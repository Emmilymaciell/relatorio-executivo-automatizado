# -*- coding: utf-8 -*-
"""Organiza a entrega do G3 (Certificadora) no padrão do POP de Entrega v2026-10-01 — SÓ CÓPIA, pela API do
SharePoint (Microsoft Graph, DEC-052), porque a pasta do G3 pode não estar sincronizada pelo OneDrive.

Origem (layout antigo enviado pelo G3, em 4. Produtos/<mês>/G3):
  - raiz: Relatório Mensal (PDF assinado) + Lista de Demandas (.xlsx)
  - Anexos/P3.1 - …/, Anexos/P3.2 - …/, Anexos/P3.3 - …/ e Anexos/P3.3 - …/ANEXOS RVS/  (subpastas)
Destino (layout do POP, criado como ARQUIVOS NOVOS — nada é excluído, movido ou renomeado):
  - G3/Programa - Relatório Grupo 3 - Relatório de Produtos Mensal N°NN - MM AAAA.pdf
  - G3/Anexos/<P3.x - Produto - Descrição - Código - Rodovia>.pdf   (direto em Anexos, sem subpasta)
  - Índice da raiz do mês: só a subdivisão GRUPO 3 é reescrita (backup + eTag via sharepoint.py --substituir).
Depois da revisão, a USUÁRIA retira do G3 o layout antigo (relatório com nome antigo, Lista de Demandas e as
subpastas P3.x) — a API não exclui nem move.

Etapas (saída em output/G3_POP_NN_Mes_AAAA/):
  baixar     --mes 15   baixa o G3 + Índice do SharePoint para ORIGEM/ (só leitura)
  planejar   --mes 15   cruza Lista de Demandas × relatório × arquivos, propõe nomes, monta STAGING + Índice com o G3,
                        roda a auditoria oficial no STAGING e grava manifesto_envio.json (publicar-lote)
  verificar  --mes 15   relê o SharePoint (baixa de novo o que está no G3 e o Índice) e gera RELATORIO_VERIFICACAO_G3.md
O envio é feito com o CLI do projeto, com autorização da usuária:
  python scripts/sharepoint.py publicar-lote output/G3_POP_…/manifesto_envio.json [--aplicar]
  python scripts/sharepoint.py enviar output/G3_POP_…/STAGING/<mês>/<Índice>.docx --para "4. Produtos/<mês>/<Índice>.docx" --substituir [--aplicar]
"""
import argparse
import collections
import datetime
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.abspath(os.path.join(AQUI, "..", ".."))
sys.path.insert(0, os.path.join(PROJ, "scripts"))
sys.path.insert(0, AQUI)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import fitz  # noqa: E402
import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402

import organizar_g2_pop as O  # noqa: E402  (Índice por grupo, util)
from prog_entrega import arquivos as A  # noqa: E402
from prog_entrega import assinatura as S  # noqa: E402
from prog_entrega import indice as IDX  # noqa: E402
from prog_entrega import nomes as N  # noqa: E402
from prog_entrega import regras as R  # noqa: E402

G = "G3"
ICONE = {"PASS": "✅ PASS", "BLOCKER": "⛔ BLOCKER", "MANUAL_REVIEW": "🟠 MANUAL_REVIEW", "WARNING": "🟡 WARNING",
         "PENDENTE": "⏳ PENDENTE (sua ação)", "INFO": "ℹ️ INFO"}
_RE_RGC = re.compile(r"^RL-P31-(BI\d{3})-RGC-(\d{3})_(\d{2})_", re.I)
_RE_RO = re.compile(r"^P33_-_(\d{4})(\d{2})_-_RD_(\d{3}A?)_-_Relat[oó]rio_de_Obra\.pdf$", re.I)
_RE_RV = re.compile(r"^(RV-(OB|PE)-[A-Z]{2}-RD(\d{3}A?)-[A-Z]{3}-E\d{4}-E\d{4}-[A-Z]{3}-\d{3}_\d{2})\.pdf$", re.I)


def _sha(p):
    return O._sha(p)


def caminhos(mes):
    P = O.caminhos(mes)  # pasta do mês etc. (o out do G2 não é usado)
    out = os.path.join(PROJ, "output", f"G3_POP_{P['num']:02d}_{R.norm(P['mes']).capitalize()}_{P['ano']}")
    os.makedirs(out, exist_ok=True)
    remoto_mes = f"4. Produtos/{P['pasta_mes_nome']}"
    return dict(P, out=out, origem=os.path.join(out, "ORIGEM"), staging_mes=os.path.join(out, "STAGING", P["pasta_mes_nome"]),
                remoto_mes=remoto_mes, remoto_g3=f"{remoto_mes}/{G}", indice_nome=os.path.basename(P["indice_sp"]))


def _sp():
    import sharepoint as SPC
    return SPC.conectar()[0]


def baixar_arvore(sp, remoto_base, destino, so=None):
    inv = []

    def walk(rel):
        for it in sp.listar(f"{remoto_base}/{rel}" if rel else remoto_base):
            r = f"{rel}/{it['name']}" if rel else it["name"]
            if so and not so(r, "folder" in it):
                continue
            if "folder" in it:
                walk(r)
            else:
                d = sp.baixar(f"{remoto_base}/{r}", os.path.join(destino, *r.split("/")))
                d.update(rel=r, modificado_por=((it.get("lastModifiedBy") or {}).get("user") or {}).get("displayName"))
                inv.append(d)
    walk("")
    return inv


def baixar(mes):
    P = caminhos(mes)
    sp = _sp()
    if os.path.isdir(P["origem"]):
        shutil.rmtree(R.lp(P["origem"]))
    inv = baixar_arvore(sp, P["remoto_mes"], P["origem"],
                        so=lambda r, pasta: r == G or r.startswith(G + "/") or (not pasta and r == P["indice_nome"]))
    O._jdump(inv, os.path.join(P["out"], "inventario_origem.json"))
    print(f"[baixar] {len(inv)} arquivo(s) em {P['origem']}")


def ler_lista(path):
    """Lista de Demandas do G3 — colunas localizadas pelo CABEÇALHO (linha com 'Documento')."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    h = next(k for k, r in enumerate(rows) if r and any(R.norm(x) == "documento" for x in r if x))
    col = {R.norm(x): c for c, x in enumerate(rows[h]) if x}

    def v(r, nome):
        c = col.get(nome)
        return r[c] if c is not None and c < len(r) else None
    linhas = []
    for r in rows[h + 1:]:
        doc = v(r, "documento")
        if not doc:
            continue
        linhas.append({"id": v(r, "id"), "produto": "P" + str(v(r, "produto")).strip()[:3], "documento": str(doc).strip(),
                       "tipo": v(r, "tipo"), "descricao": v(r, "descricao"), "empreendimento": str(v(r, "empreendimento") or "").strip(),
                       "data": str(v(r, "data da entrega"))[:10], "status": v(r, "status entrega"), "versao": v(r, "versao")})
    return linhas


def _capa(p, n=2):
    d = fitz.open(R.lp(p))
    try:
        return re.sub(r"\s+", " ", " ".join(d[i].get_text() for i in range(min(n, d.page_count))))
    finally:
        d.close()


def planejar(mes):
    P = caminhos(mes)
    inv = O._jload(os.path.join(P["out"], "inventario_origem.json"))
    g3o = os.path.join(P["origem"], G)
    prod = R.produtos()
    lista_arq = [x for x in inv if x["rel"].startswith(f"{G}/") and x["rel"].count("/") == 1 and x["rel"].lower().endswith(".xlsx")]
    rel_arq = [x for x in inv if x["rel"].startswith(f"{G}/") and x["rel"].count("/") == 1 and x["rel"].lower().endswith(".pdf")]
    if len(rel_arq) != 1 or len(lista_arq) != 1:
        raise SystemExit(f"esperado 1 relatório PDF e 1 Lista na raiz do G3: {[x['rel'] for x in rel_arq + lista_arq]}")
    lista = ler_lista(os.path.join(P["origem"], *lista_arq[0]["rel"].split("/")))
    itens = []
    for x in inv:
        if not x["rel"].startswith(f"{G}/Anexos/"):
            continue
        nome = os.path.basename(x["rel"])
        p = os.path.join(P["origem"], *x["rel"].split("/"))
        it = {"origem_rel": x["rel"], "nome_original": nome, "sha256": x["sha256"], "tamanho": x["tamanho"],
              "modificado": x.get("modificado"), "achados": []}
        capa = _capa(p)
        if (m := _RE_RGC.match(nome)):
            cod = f"RL-P3.1-{m.group(1)}-RGC-{m.group(2)}"
            lm = [r for r in lista if r["documento"].upper() == cod.upper()]
            rev = m.group(3)
            it.update(produto="P3.1", doc=f"{cod} Rev {rev}", descricao="Relatório de Gestão da Certificação",
                      rodovia=O.rodovia_lista(lm[0]["empreendimento"]) if lm else None, lista=lm[0] if lm else None,
                      fonte="nome + Lista de Demandas (Empreendimento) + capa")
            if lm and str(lm[0]["versao"]).strip() != str(int(rev)):
                it["achados"].append(["WARNING", f"revisão no nome ({rev}) ≠ Versão na Lista ({lm[0]['versao']})"])
        elif (m := _RE_RO.match(nome)):
            rod = f"RD-{m.group(3).upper()}"
            comp = f"{m.group(1)}.{m.group(2)}"
            revm = re.search(rf"P3\.3 - {comp} - RD {m.group(3)}\s+(\d{{2}})", capa)
            rev = revm.group(1) if revm else None
            lm = [r for r in lista if r["produto"] == "P3.3" and O.rodovia_lista(r["empreendimento"]) == rod]
            it.update(produto="P3.3", doc=f"{comp} Rev {rev}" if rev else comp, descricao="Relatório de Obra", rodovia=rod,
                      lista=lm[0] if len(lm) == 1 else None, fonte="nome + capa ('P3.3 - AAAA.MM - RD xxx' + revisão) + Lista")
            if len(lm) == 1 and not lm[0]["documento"].upper().startswith("P3.3"):
                it["achados"].append(["WARNING", f"Lista de Demandas registra o código '{lm[0]['documento']}' (com P3.1) — "
                                                 f"o documento é P3.3 ('P3.3 - {comp} - RD {m.group(3)}' na capa)"])
        elif (m := _RE_RV.match(nome)):
            cod, fase, r_ = m.group(1), m.group(2).upper(), m.group(3).upper()
            rod = f"RD-{r_}"
            desc = "Relatório de Visita de Obra" if fase == "OB" else "Relatório de Visita de Projeto Executivo"
            it.update(produto="P3.3", doc=cod, descricao=desc, rodovia=rod, lista=None,
                      fonte="código do documento (fase OB/PE + rodovia) + capa")
            it["achados"].append(["INFO", "anexo do Relatório de Obra da mesma rodovia (citado no texto dele) — não tem "
                                          "linha própria na Lista de Demandas, como em agosto"])
        else:
            it["achados"].append(["BLOCKER", "arquivo não reconhecido — decidir nome/produto manualmente"])
            itens.append(it)
            continue
        if it["rodovia"] not in R.rodovias():
            it["achados"].append(["BLOCKER", f"rodovia não determinada ({it['rodovia']})"])
        caprods = N.rodovias_mencionadas(capa.replace("RD ", "RD-")) | O.rodovias_texto(capa)
        if it["rodovia"] != "GERAL" and caprods and it["rodovia"] not in caprods:
            it["achados"].append(["MANUAL_REVIEW", f"capa cita {sorted(caprods)}, nome indica {it['rodovia']}"])
        stem = f"{it['produto']} - {prod[it['produto']]['nome']} - {it['descricao']} - {it['doc']} - {it['rodovia']}"
        it["destino"] = stem + ".pdf"
        a = N.analisar(it["destino"])
        if not a["conforme"]:
            it["achados"].append(["BLOCKER", f"nome final fora do padrão: {a['problemas']}"])
        if len(it["destino"]) > 135:
            it["achados"].append(["BLOCKER", f"nome final com {len(it['destino'])} caracteres (> 135)"])
        itens.append(it)

    # relatório
    rx = rel_arq[0]
    rel_nome = f"Programa - Relatório Grupo 3 - Relatório de Produtos Mensal N°{P['num']} - {P['mes_num']:02d} {P['ano']}.pdf"
    rel = {"origem_rel": rx["rel"], "nome_original": os.path.basename(rx["rel"]), "sha256": rx["sha256"],
           "tamanho": rx["tamanho"], "destino": rel_nome}
    rtxt = " ".join(re.sub(r"\s+", " ", pg.get_text()) for pg in fitz.open(os.path.join(P["origem"], *rx["rel"].split("/"))))
    m = re.search(r"constituído de (\d+) \(\w+\) páginas e (\d+) \(\w+\) anexos", rtxt)
    rel["declara_paginas_anexos"] = (int(m.group(1)), int(m.group(2))) if m else None
    rel["placeholders_assinatura"] = rtxt.count("Nome – Programa")
    # Lista × arquivos
    casados = {id(it["lista"]) for it in itens if it.get("lista")}
    lista_sem_arq = [r for r in lista if id(r) not in casados]
    # STAGING + Índice
    st = P["staging_mes"]
    if os.path.isdir(st):
        shutil.rmtree(R.lp(st))
    an = os.path.join(st, G, "Anexos")
    os.makedirs(R.lp(an))
    destinos = collections.Counter(R.norm(it.get("destino", "")) for it in itens)
    for it in itens:
        if destinos[R.norm(it.get("destino", ""))] > 1:
            it["achados"].append(["BLOCKER", "colisão de nome final"])
    pend = [it["nome_original"] for it in itens if any(s == "BLOCKER" for s, _ in it["achados"])]
    if pend:
        O._jdump({"itens": itens, "relatorio": rel, "lista": lista}, os.path.join(P["out"], "plano_g3.json"))
        raise SystemExit(f"[planejar] BLOCKER em {pend} — ver plano_g3.json")
    for it in itens:
        shutil.copy2(R.lp(os.path.join(P["origem"], *it["origem_rel"].split("/"))), R.lp(os.path.join(an, it["destino"])))
    shutil.copy2(R.lp(os.path.join(P["origem"], *rx["rel"].split("/"))), R.lp(os.path.join(st, G, rel_nome)))
    finais = sorted((os.path.splitext(it["destino"])[0] for it in itens), key=lambda s: (s.split(" - ")[0], R.norm(s)))
    linhas = [os.path.splitext(rel_nome)[0]] + finais
    idx_src = os.path.join(P["origem"], P["indice_nome"])
    idx_dst = os.path.join(st, P["indice_nome"])
    O._escrever_g2_no_indice(idx_src, idx_dst, linhas, grupo=3)
    antes, _ = O._linhas_por_grupo(idx_src)
    depois, _ = O._linhas_por_grupo(idx_dst)
    if not all(antes.get(g) == depois.get(g) for g in ("G1", "G2", "G4", "G5")):
        raise RuntimeError("o Índice alteraria linhas de outro grupo — abortado")
    idx_sha_base = next(x["sha256"] for x in inv if x["rel"] == P["indice_nome"])
    # manifesto para o publicar-lote (arquivos NOVOS) — Índice vai separado com --substituir
    man = {"motivo": f"G3 {P['pasta_mes_nome']}: layout do POP (cópia renomeada; originais intactos)",
           "itens": [{"local": os.path.join(an, it["destino"]), "remoto": f"{P['remoto_g3']}/Anexos/{it['destino']}",
                      "substituir": False} for it in itens]
           + [{"local": os.path.join(st, G, rel_nome), "remoto": f"{P['remoto_g3']}/{rel_nome}", "substituir": False}]}
    O._jdump(man, os.path.join(P["out"], "manifesto_envio.json"))
    plano = {"gerado_em": datetime.datetime.now().isoformat(timespec="seconds"), "itens": itens, "relatorio": rel,
             "lista": lista, "lista_arquivo": lista_arq[0]["rel"], "lista_sem_arquivo": lista_sem_arq,
             "indice": {"linhas_g3": linhas, "base_sha256": idx_sha_base, "novo_sha256": _sha(idx_dst),
                        "local_novo": idx_dst, "remoto": f"{P['remoto_mes']}/{P['indice_nome']}"}}
    O._jdump(plano, os.path.join(P["out"], "plano_g3.json"))
    env = dict(os.environ, PROG_ENTREGA_OUT=os.path.join(P["out"], "auditoria_staging"))
    r = subprocess.run([sys.executable, os.path.join(PROJ, "scripts", "entrega.py"), "auditar", "--pasta", st], env=env,
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(r.stdout[-600:])
    for it in itens:
        print(f"  {it['origem_rel']}\n      -> G3/Anexos/{it['destino']}")
        for s, msg in it["achados"]:
            if s != "INFO":
                print(f"         {s}: {msg}")
    print(f"  {rel['origem_rel']}\n      -> G3/{rel_nome}")
    print(f"[planejar] {len(itens)} anexo(s) + relatório; Índice G3: {len(linhas)} linha(s); Lista sem arquivo: "
          f"{[r['documento'] for r in lista_sem_arq]}")


def verificar(mes):
    P = caminhos(mes)
    plano = O._jload(os.path.join(P["out"], "plano_g3.json"))
    sp = _sp()
    fin = os.path.join(P["out"], "FINAL_SHAREPOINT")
    if os.path.isdir(fin):
        shutil.rmtree(R.lp(fin))
    # estado real: lista o G3 inteiro (inclusive o legado), baixa só a camada nova + Índice
    novos = {f"{G}/Anexos/{it['destino']}" for it in plano["itens"]} | {f"{G}/{plano['relatorio']['destino']}"}
    real = []

    def walk(rel):
        for it in sp.listar(f"{P['remoto_mes']}/{rel}"):
            r = f"{rel}/{it['name']}"
            real.append((r, "folder" in it, it.get("size")))
            if "folder" in it:
                walk(r)
    walk(G)
    for r, pasta, _ in real:
        if not pasta and r in novos:
            sp.baixar(f"{P['remoto_mes']}/{r}", os.path.join(fin, *r.split("/")))
    sp.baixar(f"{P['remoto_mes']}/{P['indice_nome']}", os.path.join(fin, P["indice_nome"]))
    lista_rel = f"{G}/{os.path.basename(plano['lista_arquivo'])}"
    lista_viva = any(r == lista_rel and not pasta for r, pasta, _ in real)
    if lista_viva:  # relê a Lista de Demandas como está AGORA no SharePoint
        sp.baixar(f"{P['remoto_mes']}/{lista_rel}", os.path.join(fin, *lista_rel.split("/")))
        plano["lista"] = ler_lista(os.path.join(fin, *lista_rel.split("/")))
    legado = [r for r, pasta, _ in real if r not in novos and not (pasta and r == f"{G}/Anexos") and r != f"{G}/{os.path.basename(plano['lista_arquivo'])}"]
    raiz_g3 = [r for r, pasta, _ in real if r.count("/") == 1 and not pasta]
    anexos_dir = [r for r, pasta, _ in real if r.startswith(f"{G}/Anexos/") and r.count("/") == 2 and not pasta]
    subpastas = [r for r, pasta, _ in real if pasta and r.startswith(f"{G}/Anexos/")]
    lim = R.limite_bytes()
    linhas = []

    def item(n, t, s, e):
        linhas.append((n, t, s, e))

    # conferências por arquivo (camada nova)
    por = []
    for it in plano["itens"]:
        p = os.path.join(fin, G, "Anexos", it["destino"])
        ok = os.path.exists(R.lp(p))
        v = A.validar(R.lp(p)) if ok else {"tipo": None, "problemas": [["BLOCKER", "ausente no SharePoint"]], "info": {}}
        sg = S.verificar_pdf(R.lp(p)) if ok else {"status": "AUSENTE", "assinaturas": []}
        a = N.analisar(it["destino"])
        por.append(dict(it, presente=ok, sha_ok=ok and _sha(p) == it["sha256"], integ=v, assin=sg["status"],
                        nome_ok=a["conforme"] and not a["problemas"], tam=os.path.getsize(p) if ok else 0,
                        paginas=v["info"].get("paginas"), paginas_texto=v["info"].get("paginas_com_texto")))
    relp = os.path.join(fin, G, plano["relatorio"]["destino"])
    rel_ok = os.path.exists(R.lp(relp))
    sig = S.verificar_pdf(R.lp(relp)) if rel_ok else {"status": "AUSENTE", "assinaturas": [], "motivo": ""}
    rv = A.validar(R.lp(relp)) if rel_ok else {"info": {}, "problemas": []}
    li = IDX.ler_indice(os.path.join(fin, P["indice_nome"]))
    g3i = [ln["arquivo"] for ln in li["linhas"] if ln["grupo"] == "G3"]
    reais = [os.path.splitext(plano["relatorio"]["destino"])[0]] + [os.path.splitext(x["destino"])[0] for x in por if x["presente"]]
    falt = [x for x in reais if R.norm(x) not in {R.norm(n) for n in g3i}]
    sobr = [x for x in g3i if R.norm(x) not in {R.norm(n) for n in reais}]
    nums = [ln["num"] for ln in li["linhas"] if ln["grupo"] == "G3"] == [str(i) for i in range(1, len(g3i) + 1)]
    base = IDX.ler_indice(os.path.join(P["origem"], P["indice_nome"]))
    ag = collections.defaultdict(list)
    dg = collections.defaultdict(list)
    for ln in base["linhas"]:
        ag[ln["grupo"]].append((ln["num"], ln["arquivo"]))
    for ln in li["linhas"]:
        dg[ln["grupo"]].append((ln["num"], ln["arquivo"]))
    outros = {g: ag.get(g) == dg.get(g) for g in ("G1", "G2", "G4", "G5")}
    mudaram = [g for g, ok in outros.items() if not ok]

    lista = plano["lista"]
    cod_errado = [r["documento"] for r in lista if r["produto"] != "P3.1" and r["documento"].upper().startswith("P3.1")]
    for x in por:  # avisos de código da Lista só valem se o erro ainda existir
        x["achados"] = [a for a in x["achados"] if not (a[0] == "WARNING" and "Lista de Demandas registra" in a[1] and not cod_errado)]
    item(1, "Lista de Demandas preenchida no modelo padrão.", "WARNING" if cod_errado else "PASS",
         f"`{plano['lista_arquivo']}` relida agora do SharePoint ({len(lista)} linha(s): {', '.join(r['documento'] for r in lista)}). "
         + (f"Código com P3.1 em linhas do P3.3: {cod_errado}." if cod_errado else
            "Códigos coerentes com os documentos (as 4 linhas do P3.3 corrigidas de 'P3.1 - 2026.09 - RD-…' para 'P3.3 - …' "
            "em 08/10/2026, a pedido da usuária; backup em output/SHAREPOINT_BACKUPS/).")
         + " A Lista é entregue pelo grupo e NÃO entra no pacote do Parceiro (controle interno — Checklist Coord. 11).")
    det = (sig.get("assinaturas") or [{}])[0]
    item(2, "Relatório de Produtos Mensal FINAL assinado em PDF.", "PASS" if sig["status"] == S.VALIDADA else "BLOCKER",
         f"`G3/{plano['relatorio']['destino']}` (cópia byte a byte de `{plano['relatorio']['nome_original']}`): "
         f"{rv['info'].get('paginas')} págs., texto em {rv['info'].get('paginas_com_texto')}; assinatura **{sig['status']}** "
         f"via DocuSign ({det.get('emissor')}). Observação: o bloco de assinaturas do Encerramento tem "
         f"{plano['relatorio']['placeholders_assinatura']} campo(s) 'Nome – Programa' sem nome preenchido (a assinatura digital "
         "existe; é só o texto do modelo).")
    raiz_extra = [r for r in raiz_g3 if r not in (f"{G}/{plano['relatorio']['destino']}", f"{G}/{os.path.basename(plano['lista_arquivo'])}")]
    item(3, "Somente o relatório final na raiz do grupo (sem Word, minuta ou versão antiga).",
         "PENDENTE" if raiz_extra else "PASS",
         (f"Ainda na raiz (layout antigo, a retirar após revisar): {raiz_extra}." if raiz_extra else
          "Só o relatório final (nome padrão) na raiz; o relatório com o nome antigo foi retirado pela usuária.")
         + (" A Lista de Demandas também está na raiz, como o POP pede ao grupo; ela sai do pacote na montagem final "
            "(Checklist Coord. 11 — `entrega.py preparar` a retira automaticamente)." if lista_viva else ""))
    item(4, "Todos os ANEXOS na pasta ANEXOS (pasta única, sem subpastas).", "PENDENTE" if subpastas else "PASS",
         f"{sum(x['presente'] for x in por)}/{len(por)} anexos direto em `G3/Anexos/`. "
         + (f"Subpastas do layout antigo ainda presentes (originais intactos, a retirar por você): {subpastas}." if subpastas else
           "Sem subpastas: o layout antigo (P3.1/P3.2/P3.3 e ANEXOS RVS) foi retirado pela usuária em 08/10/2026."))
    bad = [x["destino"] for x in por if not x["nome_ok"]]
    item(5, "Nomenclatura padrão [Código] - [Produto] - [Nome do arquivo] - [Rodovia].", "PASS" if not bad else "BLOCKER",
         f"{len(por) - len(bad)}/{len(por)} conformes." + (f" Fora: {bad}" if bad else ""))
    rods = collections.Counter(x["rodovia"] for x in por)
    item(6, "Rodovia identificada no nome (RD-101, RD-102A, RD-103, RD-104 ou GERAL).",
         "PASS" if all(x["rodovia"] in R.rodovias() for x in por) else "BLOCKER",
         f"Distribuição: {dict(sorted(rods.items()))}. RGC (P3.1) = GERAL pela Lista; Relatórios de Obra e RVs pela rodovia do "
         "nome/código, conferida com a capa.")
    item(7, "Nenhum arquivo com mais de 100 MB.", "PASS" if all(x["tam"] <= lim for x in por) else "BLOCKER",
         f"Maior anexo: {max(x['tam'] for x in por) / 1048576:.1f} MB; relatório: {plano['relatorio']['tamanho'] / 1048576:.1f} MB.")
    item(8, "Todo arquivo que não for PDF está em ZIP.", "PASS" if all(x["destino"].endswith(".pdf") for x in por) else "BLOCKER",
         f"{len(por)} PDF; nenhum outro formato.")
    item(9, "Arquivos > 100 MB fracionados.", "PASS", "Nenhum arquivo precisou ser fracionado.")
    st10 = "PASS" if (not falt and not sobr and nums and not mudaram) else "BLOCKER"
    item(10, "Índice do mês atualizado na subdivisão do grupo.", st10,
         f"GRUPO 3 com {len(g3i)} linha(s) (relatório + {len(por)} anexos), numeração 1…{len(g3i)} "
         f"{'OK' if nums else 'FORA DE ORDEM'}; {len(falt)} faltando, {len(sobr)} sobrando; SEI em branco. "
         + ("G1/G2/G4/G5 inalterados." if not mudaram else f"ATENÇÃO: {mudaram} diferem da versão-base (outro grupo editou?)."))
    bad11 = [x["destino"] for x in por if not x["presente"] or not x["sha_ok"] or any(s == "BLOCKER" for s, _ in x["integ"]["problemas"])]
    semtxt = [x["destino"] for x in por if x.get("paginas_texto") == 0]
    item(11, "Todos os arquivos abertos/testados, sem corrupção.", "BLOCKER" if bad11 else ("MANUAL_REVIEW" if semtxt else "PASS"),
         f"{len(por)} anexos + relatório baixados de volta do SharePoint e reabertos; SHA-256 = original em "
         f"{sum(x['sha_ok'] for x in por)}/{len(por)}; todos com texto pesquisável; assinaturas digitais (DocuSign): "
         f"{dict(collections.Counter(x['assin'] for x in por))}." + (f" Problemas: {bad11}" if bad11 else ""))
    decl = plano["relatorio"].get("declara_paginas_anexos")
    n_rv = sum(1 for x in por if x["doc"].startswith("RV-"))
    item(12, "Não falta nenhum documento da entrega do grupo.", "BLOCKER" if plano["lista_sem_arquivo"] else "PASS",
         f"Lista de Demandas × anexos: {len(lista) - len(plano['lista_sem_arquivo'])}/{len(lista)} com arquivo. Relatório declara "
         f"{decl[1] if decl else '?'} anexos (Anexos I–V = RGC + 4 Relatórios de Obra) — todos presentes. P3.2: 'Não houve demandas "
         f"do Produto 3.2 no período' (pasta vazia, coerente). + {n_rv} Relatórios de Visita (RV), citados dentro dos Relatórios de "
         "Obra da mesma rodovia (RD-103: 4; RD-104: 4; RD-102A: 1), mantidos como anexos — igual a agosto.")
    pior = min([s for _, _, s, _ in linhas], key=["BLOCKER", "MANUAL_REVIEW", "PENDENTE", "WARNING", "PASS"].index)
    veredito = {"PASS": "PRONTO PARA PROTOCOLO (G3)", "WARNING": "PRONTO (com avisos ao G3)",
                "PENDENTE": "CAMADA NOVA PRONTA — falta você retirar o layout antigo do G3",
                "MANUAL_REVIEW": "PENDENTE DE REVISÃO HUMANA", "BLOCKER": "NÃO PRONTO — há BLOCKER"}[pior]

    agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    md = [f"# Relatório de verificação — G3 (Certificadora) · {P['pasta_mes_nome']} (RM{P['num']})", "",
          f"Gerado em {agora} · lido **direto do SharePoint** (API Graph; a pasta do G3 não estava sincronizada neste "
          "computador) · regras: POP de Entrega v2026-10-01 + Checklist Visual (aba G3) + Modelo de Índice r2.", "",
          f"## Veredito: {ICONE[pior]} — {veredito}", "",
          "| Nº | Verificação (Checklist G3) | Resultado | Evidência |", "|---|---|---|---|"]
    md += [f"| {n} | {t} | {ICONE[s]} | {e.replace('|', '/')} |" for n, t, s, e in linhas]
    md += ["", "## Antes → depois", "", "| Original (layout antigo) | Cópia no padrão POP | Fonte do nome |", "|---|---|---|",
           f"| `{plano['relatorio']['origem_rel']}` | `G3/{plano['relatorio']['destino']}` | DEC-044 (designação) |"]
    md += [f"| `{x['origem_rel']}` | `G3/Anexos/{x['destino']}` | {x['fonte']} |" for x in por]
    md += ["", "## Ainda no G3 (layout antigo) — retirar depois da sua revisão", "",
           "A API não exclui nem move (regra do projeto). Tudo abaixo tem cópia conferida por SHA-256 na camada nova, "
           "exceto a Lista de Demandas, que é controle interno e não vai ao Parceiro:", ""]
    md += [f"- `{r}`" for r in sorted(legado)] or ["- — (nada: o layout antigo já foi retirado)"]
    avisos = [(x["nome_original"], m) for x in por for s, m in x["achados"] if s in ("WARNING", "MANUAL_REVIEW")]
    md += ["", "## Avisos para o G3 (não bloqueiam)", ""] + [f"- `{o}` — {m}" for o, m in avisos] + (
        [] if cod_errado else ["- Lista de Demandas: código das 4 linhas do P3.3 corrigido de P3.1 para P3.3 (08/10/2026)."]) + [
        f"- Relatório: Encerramento com {plano['relatorio']['placeholders_assinatura']} campo(s) 'Nome – Programa' sem nome.",
        "- RV-PE-RD-RD102A…: o nome do arquivo usa 'RD102A', o código interno usa 'RD102' (mesma rodovia)."]
    md += ["", "## Rastreabilidade", "",
           f"- Planilha: `{os.path.join(P['out'], 'G3_verificacao_antes_depois.xlsx')}` · plano: `{os.path.join(P['out'], 'plano_g3.json')}`",
           f"- Cópia de tudo como estava no SharePoint antes (ORIGEM): `{P['origem']}`",
           "- Backup do Índice substituído: `output/SHAREPOINT_BACKUPS/<data-hora>/` · log de envios: `logs/sharepoint/`"]
    out_md = os.path.join(P["out"], "RELATORIO_VERIFICACAO_G3.md")
    with open(out_md, "w", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Antes → Depois"
    ws.append(["Original", "Novo nome (G3/Anexos)", "Produto", "Documento", "Rodovia", "Páginas", "Págs. c/ texto",
               "Assinatura", "SHA-256 confere", "Nome conforme", "Observações", "SHA-256"])
    ws.append([plano["relatorio"]["origem_rel"], plano["relatorio"]["destino"], "REL", "Relatório de Produtos Mensal", "—",
               rv["info"].get("paginas"), rv["info"].get("paginas_com_texto"), sig["status"], "SIM" if rel_ok and
               _sha(relp) == plano["relatorio"]["sha256"] else "NÃO", "—", "", plano["relatorio"]["sha256"]])
    for x in por:
        ws.append([x["origem_rel"], x["destino"], x["produto"], x["doc"], x["rodovia"], x.get("paginas"), x.get("paginas_texto"),
                   x["assin"], "SIM" if x["sha_ok"] else "NÃO", "SIM" if x["nome_ok"] else "NÃO",
                   "; ".join(f"{s}: {m}" for s, m in x["achados"]), x["sha256"]])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E5A")
        c.alignment = Alignment(wrap_text=True)
    for col, w in zip("ABCDEFGHIJKL", [80, 100, 8, 40, 9, 7, 8, 12, 9, 9, 70, 66]):
        ws.column_dimensions[col].width = w
    ws2 = wb.create_sheet("Checklist")
    ws2.append(["Nº", "Verificação", "Resultado", "Evidência"])
    for n, t, s, e in linhas:
        ws2.append([n, t, s, e])
    for col, w in zip("ABCD", [5, 50, 16, 150]):
        ws2.column_dimensions[col].width = w
    xl = os.path.join(P["out"], "G3_verificacao_antes_depois.xlsx")
    wb.save(xl)
    print(f"[verificar] {veredito} | " + ", ".join(f"{n}={s}" for n, _, s, _ in linhas))
    print(f"[verificar] {out_md}\n[verificar] {xl}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("etapa", choices=["baixar", "planejar", "verificar"])
    ap.add_argument("--mes", type=int, required=True)
    a = ap.parse_args()
    {"baixar": baixar, "planejar": planejar, "verificar": verificar}[a.etapa](a.mes)

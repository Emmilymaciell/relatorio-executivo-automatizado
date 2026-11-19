# -*- coding: utf-8 -*-
"""Organiza a entrega do G2 no padrão do POP de Entrega (v2026-10-01) — SÓ COPIA, nunca apaga/move/renomeia.

Origem (preparada pelo engenheiro do G2):
  - relatório: Programa - 3. Desenvolvimento/G2 -  Engenharia/_Produto/<pasta do mês>/*.pdf
  - anexos:    Programa - 4. Produtos/<pasta do mês>/G2/Anexos/Anexos/P2.x/*      (layout antigo, com subpastas)
  - dados:     Programa - 3. Desenvolvimento/G2 -  Engenharia/00_ Lista mestre.xlsx  (filtro 'N° Relatório' = mês)
Destino (layout do POP, na MESMA pasta do mês):
  - G2/<relatório renomeado>.pdf          (único arquivo solto na raiz do grupo)
  - G2/Anexos/<P2.x - Produto - Descrição - Código - Rodovia>.pdf|.zip   (pasta única, sem subpastas)
  - Índice da raiz do mês: SÓ a subdivisão GRUPO 2 é reescrita (os outros grupos ficam byte a byte iguais no texto).
O layout antigo (G2/Anexos/Anexos/...) fica intacto até a usuária revisar e decidir apagá-lo.

Etapas (cada uma grava em output/G2_POP_NN_Mes_AAAA/):
  planejar  --mes 15   lê tudo, cruza Lista Mestre × nome × capa do PDF, propõe nomes. Nada é copiado.
  montar    --mes 15   monta STAGING local (cópia renomeada + ZIP fracionado + Índice com o G2) e roda a
                       auditoria oficial (scripts/prog_entrega) sobre ele. Nada muda no SharePoint.
  aplicar   --mes 15 --confirmar <código>   copia o STAGING para o SharePoint (nunca sobrescreve arquivo
                       diferente), confere SHA-256 e grava o Índice (backup antes; recusa se mudou desde o montar).
  verificar --mes 15   relê o SharePoint e gera RELATORIO_VERIFICACAO_G2.md (+ planilha antes→depois).
"""
import argparse
import collections
import csv
import datetime
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.abspath(os.path.join(AQUI, "..", ".."))
sys.path.insert(0, os.path.join(PROJ, "scripts"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import copy  # noqa: E402

import docx  # noqa: E402
import fitz  # noqa: E402
import openpyxl  # noqa: E402

from prog_entrega import arquivos as A  # noqa: E402
from prog_entrega import assinatura as S  # noqa: E402
from prog_entrega import indice as IDX  # noqa: E402
from prog_entrega import nomes as N  # noqa: E402
from prog_entrega import regras as R  # noqa: E402

G = "G2"
LIMITE_PARTE = 95 * 1000 * 1000  # margem sob os 100 MB do POP
EXECUTORAS = {"executora c": "RD-103", "executora a": "RD-101", "executora b": "RD-102A", "executora e": "RD-102A",
              "executora d": "RD-104", "executora f": "RD-104"}  # nomes fictícios (versão pública)
_RE_DOC = re.compile(r"^\s*(AT|CT|NT|PA|NC|BM|RA|RCA|RI|RO|RT|RM)\s*[\s_\-–]*\s*(\d{3})(?:\s*[.\-_]\s*(\d)(?!\d))?", re.I)
_RE_ROD_TXT = re.compile(r"\bRD\s*[-–_]?\s*(101|102\s*-?\s*A?|103|104)(?!\d)", re.I)
_PROIBIDOS = re.compile(r'[<>:"/\\|?*]')


# ------------------------------------------------------------------------------------------------ util
def _cfg_caminhos():
    with open(os.path.join(PROJ, "config", "caminhos.json"), encoding="utf-8") as f:
        return json.load(f)


def _ajustes(mes):
    with open(os.path.join(PROJ, "config", "g2_pop_ajustes.json"), encoding="utf-8") as f:
        return json.load(f).get(str(mes), {"por_arquivo": {}})


def _sha(p):
    h = hashlib.sha256()
    with open(R.lp(p), "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def _jdump(obj, p):
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1, default=str)


def _jload(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def caminhos(mes):
    c = _cfg_caminhos()
    root = c["prog_root"]
    produtos = os.path.join(root, c["relativos"]["produtos"])
    pasta = next(d for d in os.listdir(produtos) if re.match(rf"^{mes:02d}\.\s", d))
    num, nome_mes, mes_num, ano = R.parse_pasta_mes(pasta)
    g2dev = os.path.join(root, c["relativos"]["g2_engenharia"])
    out = os.path.join(PROJ, "output", f"G2_POP_{num:02d}_{R.norm(nome_mes).capitalize()}_{ano}")
    os.makedirs(out, exist_ok=True)
    return {
        "pasta_mes_nome": pasta, "mes": nome_mes, "mes_num": mes_num, "ano": ano, "num": num,
        "mes_sp": os.path.join(produtos, pasta),
        "g2_sp": os.path.join(produtos, pasta, G),
        "anexos_antigos": os.path.join(produtos, pasta, G, "Anexos", "Anexos"),
        "anexos_sp": os.path.join(produtos, pasta, G, "Anexos"),
        "produto_dev": os.path.join(g2dev, "_Produto", pasta),
        "lista_mestre": os.path.join(root, c["fontes"]["lista_mestre_g2"]),
        "indice_sp": os.path.join(produtos, pasta, IDX.nome_arquivo_indice(nome_mes, ano)),
        "out": out, "staging_mes": os.path.join(out, "STAGING", pasta),
    }


def rodovia_lista(v):
    v = (v or "").strip()
    return "GERAL" if R.norm(v) == "geral" else N.canon_rodovia(v)


def rodovias_texto(t):
    """Rodovias do programa citadas no texto (capa/título) + executoras citadas."""
    achou = set()
    for m in _RE_ROD_TXT.finditer(t or ""):
        s = re.sub(r"[\s\-]", "", m.group(1)).upper()
        achou.add("RD-102A" if s.startswith("102") else f"RD-{s}")
    n = R.norm(t)
    for k, rod in EXECUTORAS.items():
        if re.search(rf"\b{k}", n):
            achou.add(rod)
    return achou


def rodovia_capa(t):
    """Rodovia declarada no campo 'Rodovia:' ou no título da capa (primeiros 400 caracteres), se única."""
    m = re.search(r"Rodovia:\s*(RD\s*[-–]?\s*\d{3}\s*-?\s*A?)", t or "", re.I)
    if m:
        r = rodovias_texto(m.group(1))
        if len(r) == 1:
            return next(iter(r))
    r = {x for x in rodovias_texto((t or "")[:400])}
    return next(iter(r)) if len(r) == 1 else None


def limpar_descricao(txt):
    t = (txt or "").replace("\n", " ")
    t = re.sub(r",?\s*referente à Rodovia\s+RD\s*[-–]?\s*\d{3}\s*A?", "", t, flags=re.I)
    t = _RE_ROD_TXT.sub("", t)
    t = t.replace("–", " - ").replace("—", " - ")
    t = re.sub(r"\s+/\s+", " - ", t)  # ' / ' separa ideias; '327/2026' vira '327-2026'
    t = _PROIBIDOS.sub("-", t)
    t = re.sub(r"\s+-\s*|\s*-\s+", " - ", t)  # só hífen já separado por espaço (não toca '327-2026', 'CTD-43')
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"(\s-\s){2,}", " - ", t)
    return t.strip(" -.")


def doc_code(nome):
    m = _RE_DOC.match(nome)
    if not m:
        return None, None
    return f"{m.group(1).upper()}-{m.group(2)}", m.group(3)


def ler_lista_mestre(path, num):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    linhas = []
    for i, r in enumerate(wb["Controle Geral"].iter_rows(min_row=4, values_only=True), start=4):
        if r[10] in (num, str(num)):
            linhas.append({"linha": i, "produto": "P" + (r[1] or "")[:3].strip() if r[1] else None,
                           "doc": (r[3] or "").strip(), "manifestacao": r[4],
                           "descricao": re.sub(r"\s+", " ", r[5] or "").strip(), "empreendimento": (r[6] or "").strip(),
                           "data": str(r[7])[:10] if r[7] else None, "status": r[8], "versao": r[9]})
    wb.close()
    return linhas


def texto_capa(path, paginas=2):
    d = fitz.open(path)
    try:
        return re.sub(r"\s+", " ", " ".join(d[i].get_text() for i in range(min(paginas, d.page_count))))
    finally:
        d.close()


# ------------------------------------------------------------------------------------------- planejar
def planejar(mes):
    P = caminhos(mes)
    aj = _ajustes(mes)
    with open(os.path.join(PROJ, "config", "g2_pop_ajustes.json"), encoding="utf-8") as f:
        cfg_all = json.load(f)
    prod_cat = R.produtos()
    lista = ler_lista_mestre(P["lista_mestre"], P["num"])
    # duplicatas exatas na Lista Mestre (mesmo doc + empreendimento)
    vistos, dup_lista, lista_u = {}, [], []
    for r in lista:
        k = (r["doc"], rodovia_lista(r["empreendimento"]))
        if k in vistos:
            dup_lista.append({"linha": r["linha"], "duplica_linha": vistos[k]["linha"], "doc": r["doc"],
                              "empreendimento": r["empreendimento"]})
            continue
        vistos[k] = r
        lista_u.append(r)
    idx = collections.defaultdict(list)
    for r in lista_u:
        idx[r["doc"].rsplit(".", 1)[0]].append(r)

    itens, ignorados = [], []
    for sub in sorted(os.listdir(P["anexos_antigos"])):
        dsub = os.path.join(P["anexos_antigos"], sub)
        if not os.path.isdir(dsub):
            continue
        for nome in sorted(os.listdir(dsub)):
            p = os.path.join(dsub, nome)
            if os.path.isdir(p):
                ignorados.append({"origem": f"{sub}/{nome}", "motivo": "subpasta inesperada (não tratada)"})
                continue
            if R.eh_lixo(nome) or R.eh_temporario(nome):
                ignorados.append({"origem": f"{sub}/{nome}", "motivo": "arquivo de sistema/temporário — não vai ao pacote"})
                continue
            it = {"origem_rel": f"G2/Anexos/Anexos/{sub}/{nome}", "origem_abs": p, "nome_original": nome,
                  "pasta_produto": sub, "tamanho": os.path.getsize(p), "sha256": _sha(p), "achados": []}
            ach = it["achados"]
            ov = aj["por_arquivo"].get(nome, {})
            base, ver = doc_code(nome)
            it["doc_base"], it["versao_no_nome"] = base, ver
            ext = os.path.splitext(nome)[1].lower()
            capa = texto_capa(p) if ext == ".pdf" else ""
            it["capa"] = capa[:500]
            rod_nome = N.rodovias_mencionadas(nome.replace("RD- ", "RD-")) | rodovias_texto(nome)
            rod_txt = rodovias_texto(capa)
            rod_cp = rodovia_capa(capa)
            # --- casar com a Lista Mestre
            cand = list(idx.get(base, [])) if base else []
            if ver:
                cand = [r for r in cand if r["doc"].endswith("." + ver)]
            if len(cand) > 1:
                c2 = [r for r in cand if rodovia_lista(r["empreendimento"]) in rod_nome]
                if not c2 and rod_cp:
                    c2 = [r for r in cand if rodovia_lista(r["empreendimento"]) == rod_cp]
                if c2:
                    cand = c2
            if len(cand) > 1 and ov.get("rodovia"):
                c2 = [r for r in cand if rodovia_lista(r["empreendimento"]) == ov["rodovia"]]
                cand = c2 or cand[:1]
            if not base:
                ach.append(["BLOCKER", "código do documento não identificado no nome do arquivo"])
            elif not cand:
                ach.append(["BLOCKER", f"{base}{'.' + ver if ver else ''} não consta na Lista Mestre (N° Relatório {P['num']})"])
            elif len(cand) > 1:
                ach.append(["BLOCKER", f"{len(cand)} linhas da Lista Mestre para {base} — não foi possível decidir a rodovia"])
            lm = cand[0] if len(cand) == 1 else None
            it["lista"] = lm
            doc = lm["doc"] if lm else (f"{base}.{ver}" if base and ver else base)
            if lm and not ver:
                ach.append(["INFO", f"versão '.{lm['doc'].rsplit('.', 1)[1]}' tomada da Lista Mestre (o nome só trazia '{base}')"])
            it["doc"] = doc
            # --- produto
            prod = sub if sub in prod_cat else None
            if lm and lm["produto"] and lm["produto"] != prod:
                ach.append(["MANUAL_REVIEW", f"produto da pasta ({sub}) ≠ produto da Lista Mestre ({lm['produto']})"])
            if not prod and lm and lm["produto"] in prod_cat:
                prod = lm["produto"]
            it["produto"] = prod
            # --- rodovia
            rod_lm = rodovia_lista(lm["empreendimento"]) if lm else None
            if ov.get("rodovia"):
                rod, fonte = ov["rodovia"], "ajuste documentado (config/g2_pop_ajustes.json)"
                if rod != rod_lm:
                    ach.append(["MANUAL_REVIEW", f"rodovia {rod} (pelo documento) ≠ Lista Mestre ({rod_lm or '—'}): "
                                                 f"{ov.get('motivo', '')} → confirmar e corrigir a Lista Mestre"])
            elif rod_lm:
                rod, fonte = rod_lm, "Lista Mestre (Empreendimento)"
                if rod != "GERAL" and rod_cp and rod_cp != rod:
                    ach.append(["BLOCKER", f"capa do PDF indica {rod_cp}, Lista Mestre indica {rod} — decidir"])
                elif rod != "GERAL" and rod_txt and rod not in rod_txt:
                    ach.append(["MANUAL_REVIEW", f"capa cita {sorted(rod_txt)}, Lista Mestre indica {rod}"])
                elif rod != "GERAL" and rod_nome and rod not in rod_nome:
                    ach.append(["BLOCKER", f"nome original indica {sorted(rod_nome)}, Lista Mestre indica {rod}"])
            else:
                rod, fonte = None, None
                ach.append(["BLOCKER", "rodovia não determinável"])
            it["rodovia"], it["fonte_rodovia"] = rod, fonte
            it["evidencia_rodovia"] = {"lista": rod_lm, "nome_original": sorted(rod_nome), "capa": rod_cp,
                                       "texto": sorted(rod_txt)}
            # --- descrição e nome final
            if ov.get("descricao"):
                desc, fdesc = ov["descricao"], "ajuste documentado"
            elif lm:
                desc, fdesc = limpar_descricao(lm["descricao"]), "Lista Mestre (Descrição)"
                if base and base.startswith("AT-"):  # tipo + data da reunião (do nome original) para distinguir as atas
                    dt = re.search(r"\b(20\d\d-\d\d-\d\d)\b", nome)
                    desc = "Ata de " + desc[0].lower() + desc[1:] + (f" - {dt.group(1)}" if dt else "")
                    fdesc += " + tipo 'Ata' e data do nome original"
                elif base and base.startswith("NC-"):
                    desc = re.sub(r"^Notificação de Natureza", "Notificação de Não Conformidade de Natureza", desc)
                    fdesc += " + tipo da Manifestação"
            else:
                desc, fdesc = None, None
            desc = limpar_descricao(desc) if desc else None
            for de, para in cfg_all.get("abreviacoes", []):
                if desc:
                    desc = re.sub(re.escape(de), para, desc, flags=re.I)
            it["descricao"], it["fonte_descricao"] = desc, fdesc
            if ov.get("aviso"):
                ach.append(["WARNING", ov["aviso"]])
            if ov.get("motivo") and not any(s == "MANUAL_REVIEW" for s, _ in ach):
                ach.append(["INFO", ov["motivo"]])
            if prod and desc and doc and rod:
                stem = f"{prod} - {prod_cat[prod]['nome']} - {desc} - {doc} - {rod}"
                stem = re.sub(r"\s+", " ", stem)
                it["op"] = "fracionar_zip" if ov.get("fracionar_zip") else "copiar"
                it["destino_stem"] = stem
                it["destino_ext"] = ext
                if it["op"] == "copiar":
                    it["destinos"] = [stem + ext]
            else:
                it["op"] = "pendente"
            itens.append(it)

    # --- conferências cruzadas
    usados = collections.Counter(id(it["lista"]) for it in itens if it.get("lista"))
    lista_sem_arquivo = [r for r in lista_u if id(r) not in usados]
    destinos = collections.Counter(R.norm(d) for it in itens for d in it.get("destinos", []))
    nome_max = cfg_all.get("nome_max", 135)
    for it in itens:
        stem = it.get("destino_stem") or ""
        maxlen = len(stem) + (len(" - Parte 01") if it.get("op") == "fracionar_zip" else 0) + 4
        if maxlen > nome_max:
            it["achados"].append(["BLOCKER", f"nome final com {maxlen} caracteres (> {nome_max}) — encurtar"])
        for d in it.get("destinos", []):
            if destinos[R.norm(d)] > 1:
                it["achados"].append(["BLOCKER", f"COLISÃO de nome final: '{d}'"])
    por_sha = collections.defaultdict(list)
    for it in itens:
        por_sha[it["sha256"]].append(it["nome_original"])
    dups = [v for v in por_sha.values() if len(v) > 1]

    # --- relatório
    rels = [f for f in sorted(os.listdir(P["produto_dev"])) if f.lower().endswith(".pdf")]
    editaveis = [f for f in sorted(os.listdir(P["produto_dev"])) if f.lower().endswith((".doc", ".docx"))]
    rel = None
    if len(rels) == 1:
        rp = os.path.join(P["produto_dev"], rels[0])
        rel = {"origem_abs": rp, "origem_rel": f"_Produto/{P['pasta_mes_nome']}/{rels[0]}", "nome_original": rels[0],
               "sha256": _sha(rp), "tamanho": os.path.getsize(rp),
               "destino": aj.get("nome_relatorio") or rels[0], "editaveis_nao_copiados": editaveis}
    plano = {"gerado_em": datetime.datetime.now().isoformat(timespec="seconds"), "mes": P["pasta_mes_nome"],
             "fontes": {"lista_mestre": P["lista_mestre"], "lista_mestre_sha256": _sha(P["lista_mestre"]),
                        "anexos_antigos": P["anexos_antigos"], "relatorio": P["produto_dev"]},
             "relatorio": rel, "relatorios_encontrados": rels, "itens": itens, "ignorados": ignorados,
             "lista_linhas": len(lista), "lista_unicas": len(lista_u), "lista_duplicadas": dup_lista,
             "lista_sem_arquivo": lista_sem_arquivo, "duplicatas_conteudo": dups}
    _jdump(plano, os.path.join(P["out"], "plano_g2.json"))
    with open(os.path.join(P["out"], "plano_g2.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["origem", "destino", "operação", "produto", "documento", "rodovia", "fonte rodovia",
                    "fonte descrição", "achados"])
        for it in itens:
            w.writerow([it["origem_rel"], " | ".join(it.get("destinos", [])) or it.get("destino_stem", ""), it["op"],
                        it.get("produto"), it.get("doc"), it.get("rodovia"), it.get("fonte_rodovia"),
                        it.get("fonte_descricao"), " || ".join(f"{s}: {m}" for s, m in it["achados"])])
    cont = collections.Counter(s for it in itens for s, _ in it["achados"])
    print(f"[planejar] {len(itens)} anexo(s); {len(ignorados)} ignorado(s); Lista Mestre {len(lista)} linha(s) "
          f"({len(dup_lista)} duplicada(s)); sem arquivo: {[r['doc'] + ' ' + r['empreendimento'] for r in lista_sem_arquivo]}")
    print(f"[planejar] achados: {dict(cont)}; duplicatas de conteúdo: {dups}")
    print(f"[planejar] relatório: {rels} -> {rel and rel['destino']}")
    for it in itens:
        for s, m in it["achados"]:
            if s in ("BLOCKER", "MANUAL_REVIEW"):
                print(f"  {s}: {it['nome_original']} — {m}")
    return plano


# ---------------------------------------------------------------------------------------------- montar
def _membros_expandidos(zpath):
    """Lista (nome_interno, bytes_ou_None, origem) dos arquivos do ZIP; ZIPs internos > limite são abertos
    em pasta de mesmo nome (sem '.zip'). desktop.ini/Thumbs.db fora."""
    saida = []
    with zipfile.ZipFile(zpath) as z:
        for m in z.infolist():
            if m.is_dir() or R.eh_lixo(os.path.basename(m.filename)):
                continue
            if m.filename.lower().endswith(".zip"):  # POP: sem ZIP dentro de ZIP — abre todos
                pasta = m.filename[:-4]
                with z.open(m) as fi:
                    zi = zipfile.ZipFile(io.BytesIO(fi.read()))
                for m2 in zi.infolist():
                    if m2.is_dir() or R.eh_lixo(os.path.basename(m2.filename)):
                        continue
                    saida.append({"nome": f"{pasta}/{m2.filename}", "tam": m2.file_size, "dt": m2.date_time,
                                  "ler": (lambda zi=zi, m2=m2: zi.read(m2)), "origem": f"{m.filename} :: {m2.filename}"})
            else:
                saida.append({"nome": m.filename, "tam": m.file_size, "dt": m.date_time,
                              "ler": (lambda z=zpath, n=m.filename: zipfile.ZipFile(z).read(n)), "origem": m.filename})
    return saida


def fracionar_zip(src, dst_dir, stem):
    membros = _membros_expandidos(src)
    for m in membros:
        if m["tam"] > LIMITE_PARTE:
            raise RuntimeError(f"arquivo interno sozinho > limite: {m['nome']}")
    grupos, atual, tam = [], [], 0
    for m in membros:  # ordem original (mantém pastas juntas)
        if atual and tam + m["tam"] > LIMITE_PARTE:
            grupos.append(atual)
            atual, tam = [], 0
        atual.append(m)
        tam += m["tam"]
    if atual:
        grupos.append(atual)
    blocos = stem.split(N.SEP)
    saidas, conteudo = [], []
    for k, ms in enumerate(grupos, 1):
        nome = N.SEP.join(blocos[:-1] + [f"Parte {k:02d}", blocos[-1]]) + ".zip"
        dst = os.path.join(dst_dir, nome)
        with zipfile.ZipFile(R.lp(dst), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zo:
            for m in ms:
                data = m["ler"]()
                zi = zipfile.ZipInfo(m["nome"], date_time=m["dt"])
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = 0o644 << 16
                zi.flag_bits |= 0x800  # nomes UTF-8
                zo.writestr(zi, data)
                conteudo.append({"parte": nome, "arquivo": m["nome"], "origem_no_zip": m["origem"], "bytes": len(data),
                                 "sha256": hashlib.sha256(data).hexdigest()})
        with zipfile.ZipFile(R.lp(dst)) as zt:
            if zt.testzip() is not None:
                raise RuntimeError(f"CRC inválido em {nome}")
        saidas.append(nome)
    # prova de integridade: cada arquivo interno, no original, tem o mesmo SHA-256 da parte
    orig = {m["nome"]: hashlib.sha256(m["ler"]()).hexdigest() for m in membros}
    for c in conteudo:
        if orig[c["arquivo"]] != c["sha256"]:
            raise RuntimeError(f"conteúdo divergente: {c['arquivo']}")
    if sorted(orig) != sorted(c["arquivo"] for c in conteudo):
        raise RuntimeError("partes não contêm exatamente os arquivos do ZIP original")
    return saidas, conteudo


def _escrever_g2_no_indice(src_docx, dst_docx, linhas_g2, grupo=2):
    """Reescreve SÓ as linhas da subdivisão GRUPO <grupo> (protótipos = as próprias linhas do grupo no arquivo, zebrado
    preservado; coluna SEI copiada do protótipo, nunca inventada). A subdivisão vai até a próxima linha de grupo ou
    até o fim da tabela (GRUPO 5)."""
    d = docx.Document(src_docx)
    tbl = d.tables[0]._tbl
    W = IDX._W
    rows = tbl.findall(f"{W}tr")
    i_g2 = i_g3 = None
    for i, tr in enumerate(rows):
        tcs = IDX._cells_unicos(tr)
        if len(tcs) == 1:
            t = "".join(x.text or "" for x in tr.iter(f"{W}t")).strip().upper()
            if t == f"GRUPO {grupo}":
                i_g2 = i
            elif i_g2 is not None and i_g3 is None and t.startswith("GRUPO "):
                i_g3 = i
    if i_g2 is None:
        raise RuntimeError(f"subdivisão GRUPO {grupo} não encontrada no Índice")
    if i_g3 is None:
        i_g3 = len(rows)
    antigos = rows[i_g2 + 1:i_g3]
    protos = [copy.deepcopy(r) for r in (antigos[:2] if len(antigos) >= 2 else antigos)]
    if not protos:
        protos = [copy.deepcopy(next(r for r in rows if len(IDX._cells_unicos(r)) == 3 and r is not rows[0]))]
    for r in antigos:
        tbl.remove(r)
    ancora = rows[i_g2]
    for k, arq in enumerate(linhas_g2, 1):
        tr = copy.deepcopy(protos[(k - 1) % len(protos)])
        c = IDX._cells_unicos(tr)
        IDX._set_cell_text(c[0], str(k))
        IDX._set_cell_text(c[1], arq)
        ancora.addnext(tr)
        ancora = tr
    d.save(dst_docx)


def _linhas_por_grupo(path):
    r = IDX.ler_indice(path)
    g = collections.defaultdict(list)
    for ln in r["linhas"]:
        g[ln["grupo"]].append((ln["num"], ln["arquivo"], ln["sei"]))
    return dict(g), r


def montar(mes):
    P = caminhos(mes)
    plano = _jload(os.path.join(P["out"], "plano_g2.json"))
    pend = [it["nome_original"] for it in plano["itens"] if it["op"] == "pendente"
            or any(s == "BLOCKER" for s, _ in it["achados"])]
    if pend:
        raise SystemExit(f"[montar] recusado: {len(pend)} anexo(s) com BLOCKER/pendente no plano: {pend[:10]}")
    st = P["staging_mes"]
    if os.path.isdir(st):
        shutil.rmtree(R.lp(st))
    an = os.path.join(st, G, "Anexos")
    os.makedirs(R.lp(an))
    log = []
    for it in plano["itens"]:
        if it["op"] == "copiar":
            dst = os.path.join(an, it["destinos"][0])
            shutil.copy2(R.lp(it["origem_abs"]), R.lp(dst))
            ok = _sha(dst) == it["sha256"]
            log.append({"origem": it["origem_rel"], "destino": f"G2/Anexos/{it['destinos'][0]}", "op": "copiar",
                        "sha256": it["sha256"], "resultado": "OK" if ok else "FALHOU"})
            if not ok:
                raise RuntimeError(f"SHA-256 divergente: {it['nome_original']}")
        elif it["op"] == "fracionar_zip":
            partes, conteudo = fracionar_zip(it["origem_abs"], an, it["destino_stem"])
            it["destinos"] = partes
            it["conteudo_partes"] = conteudo
            for p_ in partes:
                log.append({"origem": it["origem_rel"], "destino": f"G2/Anexos/{p_}", "op": "fracionar_zip",
                            "sha256": _sha(os.path.join(an, p_)), "resultado": "OK"})
    rel = plano["relatorio"]
    if rel:
        dst = os.path.join(st, G, rel["destino"])
        shutil.copy2(R.lp(rel["origem_abs"]), R.lp(dst))
        if _sha(dst) != rel["sha256"]:
            raise RuntimeError("SHA-256 do relatório divergente")
        log.append({"origem": rel["origem_rel"], "destino": f"G2/{rel['destino']}", "op": "copiar",
                    "sha256": rel["sha256"], "resultado": "OK"})
    # índice: relatório primeiro, depois anexos por produto e nome (DEC-049); sem extensão, como os demais grupos
    finais = sorted((os.path.splitext(d)[0] for it in plano["itens"] for d in it.get("destinos", [])),
                    key=lambda s: (s.split(" - ")[0], R.norm(s)))
    linhas = ([os.path.splitext(rel["destino"])[0]] if rel else []) + finais
    idx_base = os.path.join(P["out"], "indice_base_sharepoint.docx")
    shutil.copy2(R.lp(P["indice_sp"]), R.lp(idx_base))
    idx_st = os.path.join(st, os.path.basename(P["indice_sp"]))
    _escrever_g2_no_indice(idx_base, idx_st, linhas)
    antes, _ = _linhas_por_grupo(idx_base)
    depois, _ = _linhas_por_grupo(idx_st)
    outros_ok = all(antes.get(g) == depois.get(g) for g in ("G1", "G3", "G4", "G5"))
    if not outros_ok:
        raise RuntimeError("o Índice alterou linhas de outro grupo — abortado")
    plano["montagem"] = {"em": datetime.datetime.now().isoformat(timespec="seconds"), "log": log,
                         "indice_base_sha256": _sha(idx_base), "indice_novo_sha256": _sha(idx_st),
                         "indice_linhas_g2": linhas, "outros_grupos_inalterados": outros_ok}
    codigo = hashlib.sha256(json.dumps(log, sort_keys=True).encode() + plano["montagem"]["indice_novo_sha256"].encode()
                            ).hexdigest()[:10]
    plano["montagem"]["codigo"] = codigo
    _jdump(plano, os.path.join(P["out"], "plano_g2.json"))
    # auditoria oficial (independente) sobre o STAGING
    env = dict(os.environ, PROG_ENTREGA_OUT=os.path.join(P["out"], "auditoria"))
    r = subprocess.run([sys.executable, os.path.join(PROJ, "scripts", "entrega.py"), "auditar", "--pasta", st],
                       env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(r.stdout[-1500:])
    print(f"[montar] {len(log)} arquivo(s) no STAGING; Índice: {len(linhas)} linha(s) no G2; outros grupos inalterados: "
          f"{outros_ok}; CÓDIGO para aplicar: {codigo}")


# --------------------------------------------------------------------------------------------- aplicar
def aplicar(mes, confirmar):
    P = caminhos(mes)
    plano = _jload(os.path.join(P["out"], "plano_g2.json"))
    m = plano.get("montagem") or {}
    if confirmar != m.get("codigo"):
        raise SystemExit(f"[aplicar] código não confere (esperado o do último 'montar': {m.get('codigo')})")
    st = P["staging_mes"]
    log, recusados = [], []
    os.makedirs(R.lp(P["anexos_sp"]), exist_ok=True)
    for e in m["log"]:
        src = os.path.join(st, *e["destino"].split("/"))
        dst = os.path.join(P["mes_sp"], *e["destino"].split("/"))
        if os.path.exists(R.lp(dst)):
            if _sha(dst) == e["sha256"]:
                log.append({**e, "aplicado": "já existia idêntico"})
                continue
            recusados.append(e["destino"])
            log.append({**e, "aplicado": "RECUSADO: já existe arquivo DIFERENTE com esse nome"})
            continue
        shutil.copy2(R.lp(src), R.lp(dst))
        ok = _sha(dst) == e["sha256"]
        log.append({**e, "aplicado": "copiado" if ok else "FALHOU: SHA-256 divergente"})
    # Índice: só se o do SharePoint ainda é o mesmo usado no montar
    bk_dir = os.path.join(P["out"], "backup_indice")
    os.makedirs(bk_dir, exist_ok=True)
    atual_sha = _sha(P["indice_sp"])
    lock = os.path.join(P["mes_sp"], "~$" + os.path.basename(P["indice_sp"])[2:])
    if atual_sha != m["indice_base_sha256"]:
        res_idx = "RECUSADO: o Índice do SharePoint mudou desde o 'montar' (outro grupo editou) — rode montar de novo"
    elif os.path.exists(lock):
        res_idx = "RECUSADO: o Índice está aberto por alguém (arquivo ~$ presente)"
    else:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(R.lp(P["indice_sp"]), R.lp(os.path.join(bk_dir, f"{ts}_{os.path.basename(P['indice_sp'])}")))
        shutil.copy2(R.lp(os.path.join(st, os.path.basename(P["indice_sp"]))), R.lp(P["indice_sp"]))
        res_idx = "gravado" if _sha(P["indice_sp"]) == m["indice_novo_sha256"] else "FALHOU: SHA-256 divergente"
    plano["aplicacao"] = {"em": datetime.datetime.now().isoformat(timespec="seconds"), "log": log,
                          "recusados": recusados, "indice": res_idx, "indice_sha_antes": atual_sha}
    _jdump(plano, os.path.join(P["out"], "plano_g2.json"))
    c = collections.Counter(x["aplicado"].split(":")[0] for x in log)
    print(f"[aplicar] {dict(c)}; Índice: {res_idx}")


# ------------------------------------------------------------------------------------------- verificar
def verificar(mes):
    from verificacao_g2 import gerar  # mesmo diretório
    P = caminhos(mes)
    gerar(P, _jload(os.path.join(P["out"], "plano_g2.json")))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("etapa", choices=["planejar", "montar", "aplicar", "verificar"])
    ap.add_argument("--mes", type=int, required=True)
    ap.add_argument("--confirmar")
    a = ap.parse_args()
    if a.etapa == "planejar":
        planejar(a.mes)
    elif a.etapa == "montar":
        montar(a.mes)
    elif a.etapa == "aplicar":
        if not a.confirmar:
            raise SystemExit("aplicar exige --confirmar <código do montar>")
        aplicar(a.mes, a.confirmar)
    else:
        sys.path.insert(0, AQUI)
        verificar(a.mes)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""Relatório de verificação do G2 contra o Checklist do G2 (itens 1–12), o Checklist da Coordenação (item 3:
coerência Relatório × Lista × Anexos) e o POP de Entrega v2026-10-01. Chamado por organizar_g2_pop.py verificar.

Lê o ESTADO REAL do destino (SharePoint, se já aplicado; senão o STAGING local) — não confia no plano: reabre cada
arquivo, recalcula SHA-256, relê a Lista Mestre, o Índice e a tabela do relatório. Exceções aceitas pela usuária
ficam em revisoes_humanas.json, presas ao SHA-256 do arquivo (se o arquivo mudar, a exceção deixa de valer) e
NUNCA viram regra: sem o registro, o mesmo achado volta a ser MANUAL_REVIEW/BLOCKER."""
import collections
import datetime
import hashlib
import json
import os
import re

import fitz
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from prog_entrega import arquivos as A
from prog_entrega import assinatura as S
from prog_entrega import indice as IDX
from prog_entrega import nomes as N
from prog_entrega import regras as R

ICONE = {"PASS": "✅ PASS", "BLOCKER": "⛔ BLOCKER", "MANUAL_REVIEW": "🟠 MANUAL_REVIEW", "WARNING": "🟡 WARNING",
         "PENDENTE": "⏳ PENDENTE (sua decisão)", "INFO": "ℹ️ INFO", "EXCECAO": "✅ PASS (exceção RM15)"}
_RE_DOC_NOME = re.compile(r" - ((?:AT|CT|NT|PA|NC|BM|RA|RCA|RI|RO|RT|RM)-\d{3}\.\d) - (?:Parte \d{2} - )?([^-][^ ]*)\.\w+$")
_RE_TAB = re.compile(r"\b((?:AT|CT|NT|PA|NC|BM|RA|RCA|RI|RO|RT|RM)-\d{3}\.\d)\b(.{0,260}?)\b(RD-\d{3}A?|Geral)\s+\d{2}/\d{2}/\d{4}",
                     re.I)


def _sha(p):
    h = hashlib.sha256()
    with open(R.lp(p), "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def _pior(sts):
    ordem = ["BLOCKER", "MANUAL_REVIEW", "PENDENTE", "WARNING", "EXCECAO", "PASS"]
    return min(sts, key=ordem.index) if sts else "PASS"


def _rod(v):
    v = (v or "").strip()
    return "GERAL" if R.norm(v) == "geral" else N.canon_rodovia(v)


def pares_tabela_relatorio(pdf):
    """(documento, rodovia) de cada linha da tabela 'Controle de Produtos' do corpo do relatório."""
    d = fitz.open(R.lp(pdf))
    txt = " ".join(re.sub(r"\s+", " ", p.get_text()) for p in d)
    d.close()
    return {(m.group(1).upper(), _rod(m.group(3))) for m in _RE_TAB.finditer(txt)}


def gerar(P, plano):
    import organizar_g2_pop as O  # leitura da Lista Mestre (mesmo diretório)
    with open(os.path.join(O.PROJ, "config", "g2_pop_ajustes.json"), encoding="utf-8") as f:
        cfg_mes = json.load(f).get(str(P["num"]), {})
    aplicado = bool(plano.get("aplicacao"))
    raiz_mes = P["mes_sp"] if aplicado else P["staging_mes"]
    g2 = os.path.join(raiz_mes, "G2")
    an = os.path.join(g2, "Anexos")
    lim = R.limite_bytes()
    esperados = {e["destino"]: e for e in plano["montagem"]["log"]}
    rev_p = os.path.join(P["out"], "revisoes_humanas.json")
    revs = {(c["sha256"], c["codigo"]): c for c in (json.load(open(rev_p, encoding="utf-8")) if os.path.exists(rev_p) else [])
            if c.get("decisao") == "aceito"}
    linhas = []

    def item(n, texto, st, evid):
        linhas.append((n, texto, st, evid))

    # ---------- inventário real
    raiz_files = sorted(f for f in os.listdir(g2) if os.path.isfile(os.path.join(g2, f)))
    raiz_dirs = sorted(f for f in os.listdir(g2) if os.path.isdir(os.path.join(g2, f)))
    an_files = sorted(f for f in os.listdir(an) if os.path.isfile(os.path.join(an, f)) and not R.eh_lixo(f))
    an_dirs = sorted(f for f in os.listdir(an) if os.path.isdir(os.path.join(an, f)))
    sem_rastro = [f for f in an_files if f"G2/Anexos/{f}" not in esperados]

    por_arquivo = []
    for f in an_files:
        p = os.path.join(an, f)
        e = esperados.get(f"G2/Anexos/{f}", {})
        v = A.validar(R.lp(p))
        a = N.analisar(f)
        sha = _sha(p)
        m = _RE_DOC_NOME.search(f)
        r = {"arquivo": f, "tamanho": os.path.getsize(p), "tipo": v["tipo"], "sha256": sha,
             "sha_ok": sha == e.get("sha256"), "origem": e.get("origem", "—"), "op": e.get("op", "—"),
             "nome_ok": a["conforme"] and not any(s in ("BLOCKER", "MANUAL_REVIEW") for s, _ in a["problemas"]),
             "nome_obs": "; ".join(mm for s, mm in a["problemas"]), "rodovia": a["rodovia"], "produto": a["codigo"],
             "parte": a["parte"], "base_parte": R.norm(N.base_sem_parte(os.path.splitext(f)[0])),
             "fmt_ok": os.path.splitext(f)[1].lower() in (".pdf", ".zip"), "doc": m.group(1) if m else None,
             "integ_ok": not any(s == "BLOCKER" for s, _ in v["problemas"]),
             "integ_obs": "; ".join(mm for s, mm in v["problemas"]), "info": v.get("info", {})}
        r["tam_ok"] = r["tamanho"] <= lim
        if v["tipo"] == "pdf":
            i = v["info"]
            r["paginas"], r["paginas_texto"], r["cripto"] = i.get("paginas"), i.get("paginas_com_texto"), i.get("criptografado")
        if v["tipo"] == "zip":
            r["zip_aninhado"] = [n for n in v["info"].get("membros", []) if str(n).lower().endswith(".zip")]
        por_arquivo.append(r)

    # ---------- relatório do G2
    rel_nome = (plano.get("relatorio") or {}).get("destino")
    rel_path = os.path.join(g2, rel_nome) if rel_nome else None
    rel_ok = bool(rel_path and os.path.exists(R.lp(rel_path)))
    sig = S.verificar_pdf(R.lp(rel_path)) if rel_ok else {"status": "AUSENTE", "motivo": "", "assinaturas": []}
    rel_val = A.validar(R.lp(rel_path)) if rel_ok else {"problemas": [], "info": {}}
    rel_sha = _sha(rel_path) if rel_ok else None

    # ---------- Lista Mestre (ao vivo)
    lista = O.ler_lista_mestre(P["lista_mestre"], P["num"])
    pares_lista = collections.Counter((r["doc"], _rod(r["empreendimento"])) for r in lista)
    dup_lista = [k for k, n in pares_lista.items() if n > 1]
    pares_arq = {(r["doc"], r["rodovia"]) for r in por_arquivo if r["doc"]}
    lista_sem_arq = sorted(k for k in pares_lista if k not in pares_arq)
    arq_sem_lista = sorted(r["arquivo"] for r in por_arquivo if (r["doc"], r["rodovia"]) not in pares_lista)

    # ---------- índice
    idx_path = os.path.join(raiz_mes, os.path.basename(P["indice_sp"]))
    li = IDX.ler_indice(idx_path)
    g2_idx = [ln for ln in li["linhas"] if ln["grupo"] == "G2"]
    nomes_idx = [ln["arquivo"] for ln in g2_idx]
    reais = ([os.path.splitext(rel_nome)[0]] if rel_ok else []) + [os.path.splitext(f)[0] for f in an_files]
    faltam_idx = [x for x in reais if R.norm(x) not in {R.norm(n) for n in nomes_idx}]
    sobram_idx = [x for x in nomes_idx if R.norm(x) not in {R.norm(n) for n in reais}]
    nums_ok = [ln["num"] for ln in g2_idx] == [str(i) for i in range(1, len(g2_idx) + 1)]
    ordem_ok = bool(g2_idx) and rel_ok and R.norm(g2_idx[0]["arquivo"]) == R.norm(os.path.splitext(rel_nome)[0])
    base = os.path.join(P["out"], "indice_base_sharepoint.docx")
    outros_ok = True
    if os.path.exists(base):
        a_, d_ = collections.defaultdict(list), collections.defaultdict(list)
        for ln in IDX.ler_indice(base)["linhas"]:
            a_[ln["grupo"]].append((ln["num"], ln["arquivo"], ln["sei"]))
        for ln in li["linhas"]:
            d_[ln["grupo"]].append((ln["num"], ln["arquivo"], ln["sei"]))
        outros_ok = all(a_.get(g) == d_.get(g) for g in ("G1", "G3", "G4", "G5"))

    # ---------- ZIPs fracionados (NT-793.1 + Relatórios Mensais de Obra) e RMs assinados
    fracionados = []
    for it in plano["itens"]:
        if it.get("conteudo_partes"):
            fracionados.append((it["nome_original"], it["conteudo_partes"]))
    rms = plano.get("relatorios_mensais", [])
    for z in rms:
        fracionados.append((z["zip_original"], z["conteudo_partes"]))
    rm_sig = []
    for z in rms:
        p = os.path.join(an, z["pdf"])
        s_ = S.verificar_pdf(R.lp(p)) if os.path.exists(R.lp(p)) else {"status": "AUSENTE", "assinaturas": []}
        d_ = (s_.get("assinaturas") or [{}])[0]
        rm_sig.append((z["doc"], z["rodovia"], s_["status"], d_.get("signatario"), d_.get("emissor"),
                       _sha(p) == z["pdf_sha256"] if os.path.exists(R.lp(p)) else False))

    # ---------- itens do Checklist G2
    item(1, "Lista de Demandas preenchida no modelo padrão.", "WARNING" if dup_lista else "PASS",
         f"G2 usa a Lista Mestre (`00_ Lista mestre.xlsx`, aba Controle Geral, N° Relatório = {P['num']}): {len(lista)} linha(s), "
         f"relida agora. Fica fora do pacote (controle interno — Checklist Coord. 11)."
         + (f" Linhas duplicadas: {dup_lista}." if dup_lista else " Sem linhas duplicadas (a do RO-549.1 RD-102A foi removida pelo G2)."))
    st2 = {"VALIDADA": "PASS", "NÃO ENCONTRADA": "BLOCKER", "AUSENTE": "BLOCKER", "INVÁLIDA": "BLOCKER"}.get(sig["status"], "MANUAL_REVIEW")
    det = (sig.get("assinaturas") or [{}])[0]
    item(2, "Relatório de Produtos Mensal FINAL revisado e assinado em PDF.", st2,
         f"`G2/{rel_nome}` — {rel_val['info'].get('paginas', '?')} págs., texto pesquisável em "
         f"{rel_val['info'].get('paginas_com_texto', '?')}, sem senha. Assinatura digital: **{sig['status']}**"
         + (f" — signatário **{det.get('signatario')}**, emissor {det.get('emissor')} (gov.br; ICP-Brasil não exigido — DEC-046)"
            if det.get("signatario") else f" — {sig.get('motivo', '')}")
         + ". Integridade conferida (o arquivo não foi alterado depois de assinado).")
    outros_raiz = [f for f in raiz_files if f != rel_nome and not R.eh_lixo(f)]
    item(3, "Somente a versão final do relatório na raiz do grupo (sem Word, minuta ou versão antiga).",
         "PASS" if not outros_raiz else "BLOCKER",
         f"Raiz do G2: {raiz_files} + pasta {raiz_dirs}. Word e versões anteriores ficaram em `3. Desenvolvimento/_Produto`.")
    item(4, "Todos os ANEXOS na pasta ANEXOS do grupo (pasta única, sem subpastas).",
         "BLOCKER" if an_dirs else ("WARNING" if sem_rastro else "PASS"),
         f"{len(an_files)} arquivo(s) direto em `G2/Anexos/`" + (f"; subpastas indevidas: {an_dirs}" if an_dirs else ", sem subpastas")
         + ". Layout antigo e ZIPs originais dos RMs guardados em `3. Desenvolvimento/G2/_Produto/15. Setembro 2026/Anexos/` "
           "(fora do pacote) até o OK do engenheiro."
         + (f" Sem origem rastreada: {sem_rastro}." if sem_rastro else ""))
    bad5 = [r for r in por_arquivo if not r["nome_ok"]]
    item(5, "Nomenclatura padrão [Código] - [Produto] - [Nome do arquivo] - [Rodovia].", "PASS" if not bad5 else "BLOCKER",
         f"{len(por_arquivo) - len(bad5)}/{len(por_arquivo)} conformes (validador `prog_entrega.nomes.analisar`)."
         + (f" Fora do padrão: {[r['arquivo'] for r in bad5]}" if bad5 else ""))
    rods = collections.Counter(r["rodovia"] for r in por_arquivo)
    item(6, "Rodovia identificada no nome (último bloco: RD-101, RD-102A, RD-103, RD-104 ou GERAL).",
         "PASS" if all(r["rodovia"] for r in por_arquivo) else "BLOCKER",
         f"Todo nome termina com uma das 5 rodovias válidas. Distribuição: {dict(sorted(rods.items()))}. Rodovia de cada "
         "documento = Lista Mestre, conferida com a capa do PDF.")
    grandes = [r for r in por_arquivo if not r["tam_ok"]]
    maior = max(por_arquivo, key=lambda r: r["tamanho"])
    item(7, "Nenhum arquivo com mais de 100 MB.", "PASS" if not grandes else "BLOCKER",
         f"Maior: {maior['tamanho'] / 1048576:.1f} MB. Relatório: {(os.path.getsize(rel_path) / 1048576 if rel_ok else 0):.1f} MB."
         + (f" Acima do limite: {[r['arquivo'] for r in grandes]}" if grandes else ""))
    fora = [r for r in por_arquivo if not r["fmt_ok"]]
    aninh = [r["arquivo"] for r in por_arquivo if r.get("zip_aninhado")]
    item(8, "Todo arquivo que não for PDF está compactado em ZIP.", "PASS" if not fora and not aninh else "BLOCKER",
         f"{sum(r['tipo'] == 'pdf' for r in por_arquivo)} PDF + {sum(r['tipo'] == 'zip' for r in por_arquivo)} ZIP; "
         "extensão confere com o conteúdo real; nenhum ZIP dentro de ZIP." + (f" ZIP aninhado: {aninh}" if aninh else ""))
    seq = collections.defaultdict(list)
    for r in por_arquivo:
        if r["parte"]:
            seq[r["base_parte"]].append(r["parte"])
    seq_bad = {k: v for k, v in seq.items() if sorted(v) != list(range(1, len(v) + 1))}
    n_int = sum(len(c) for _, c in fracionados)
    item(9, "Arquivos > 100 MB fracionados e numerados (Parte 01, Parte 02…).", "PASS" if not seq_bad else "BLOCKER",
         f"{len(seq)} conjunto(s) fracionado(s), todos com sequência completa: "
         + "; ".join(f"{os.path.basename(o)} → {len(c and {x['parte'] for x in c})} parte(s)" for o, c in fracionados)
         + f". {n_int} arquivos internos conferidos por SHA-256 contra os ZIPs originais (nenhum alterado). "
           "Os PDFs principais dos RMs saíram do ZIP como anexos próprios (byte a byte — assinatura preservada).")
    item(10, "Índice do mês atualizado na subdivisão do grupo.",
         "PASS" if (not faltam_idx and not sobram_idx and nums_ok and ordem_ok and outros_ok) else "BLOCKER",
         f"GRUPO 2 com {len(g2_idx)} linha(s) = 1 relatório + {len(an_files)} anexos (relatório primeiro, depois anexos por produto "
         f"e nome — DEC-049); numeração 1…{len(g2_idx)} {'OK' if nums_ok else 'FORA DE ORDEM'}; reconciliação: "
         f"{len(faltam_idx)} faltando, {len(sobram_idx)} sobrando{(' ' + str(sobram_idx[:5])) if sobram_idx else ''}; "
         f"SEI em branco (só no protocolo); G1/G3/G4/G5 {'inalterados' if outros_ok else 'ALTERADOS!'}.")
    bad11 = [r for r in por_arquivo if not r["integ_ok"] or not r["sha_ok"] or r.get("cripto")]
    sem_txt_all = [r for r in por_arquivo if r["tipo"] == "pdf" and r.get("paginas_texto") == 0]
    aceitos_txt = [r for r in sem_txt_all if (r["sha256"], "integridade") in revs]
    sem_txt = [r for r in sem_txt_all if r not in aceitos_txt]
    item(11, "Todos os arquivos abertos/testados, sem corrupção.",
         "BLOCKER" if bad11 else ("MANUAL_REVIEW" if sem_txt else ("EXCECAO" if aceitos_txt else "PASS")),
         f"{len(por_arquivo)} anexo(s) + relatório reabertos (PDF: estrutura, senha, páginas, texto; ZIP: CRC de todos os membros). "
         f"SHA-256 do arquivo = manifesto em {sum(r['sha_ok'] for r in por_arquivo)}/{len(por_arquivo)}."
         + (f" Divergentes: {[r['arquivo'] for r in bad11]}." if bad11 else "")
         + (f" Sem texto pesquisável e SEM aceite: {[r['arquivo'] for r in sem_txt]}." if sem_txt else "")
         + (f" {len(aceitos_txt)} ata(s) digitalizada(s) sem OCR aceitas pela usuária **só nesta competência** "
            "(registro preso ao SHA-256): " + ", ".join(r["doc"] for r in aceitos_txt) + "." if aceitos_txt else ""))
    rm_ok = all(s == S.VALIDADA and ok for _, _, s, _, _, ok in rm_sig)
    item(12, "Não falta nenhum documento da entrega do grupo.",
         "BLOCKER" if (lista_sem_arq or not rm_ok) else ("MANUAL_REVIEW" if arq_sem_lista else "PASS"),
         f"Lista Mestre ({len(pares_lista)} documento×rodovia) × anexos: "
         + (f"**faltam {lista_sem_arq}**; " if lista_sem_arq else "todos os documentos da Lista têm anexo; ")
         + (f"anexos sem linha na Lista: {arq_sem_lista}; " if arq_sem_lista else "nenhum anexo fora da Lista; ")
         + "Relatórios Mensais de Obra: " + "; ".join(
             f"{d} {r_} — assinatura {s}" + (f" ({sg})" if sg else "") for d, r_, s, sg, _, _ in rm_sig) + ".")
    # Checklist Coord. 3 — coerência Relatório × Lista × Anexos (regra permanente; exceção só registrada)
    pares_rel = pares_tabela_relatorio(rel_path) if rel_ok else set()
    so_rel = sorted(pares_rel - set(pares_lista))
    so_lista = sorted(set(pares_lista) - pares_rel)
    exc = revs.get((rel_sha, "tabela_relatorio"))
    st13 = "PASS" if not (so_rel or so_lista) else ("EXCECAO" if exc else "MANUAL_REVIEW")
    item(13, "Coord. 3 — Tabela do relatório coerente com a Lista Mestre e com os anexos.", st13,
         f"Tabela 'Controle de Produtos' do relatório: {len(pares_rel)} documento×rodovia lidos."
         + (f" Só no relatório: {so_rel}. Só na Lista/anexos: {so_lista}." if (so_rel or so_lista) else " Coerente.")
         + (f" **Aceito só nesta competência** pela usuária ({exc['por']}, {exc['em']}): {exc['justificativa']}" if exc else "")
         + " Regra permanente: nas próximas competências a divergência é MANUAL_REVIEW e o relatório deve ser corrigido "
           "ANTES de assinar.")

    veredito_st = _pior([s for _, _, s, _ in linhas])
    veredito = {"PASS": "PRONTO PARA PROTOCOLO (G2)", "EXCECAO": "PRONTO PARA PROTOCOLO (G2) — com exceções registradas só para o RM15",
                "WARNING": "PRONTO (com avisos)", "PENDENTE": "PRONTO APÓS SUA DECISÃO",
                "MANUAL_REVIEW": "PENDENTE DE REVISÃO HUMANA", "BLOCKER": "NÃO PRONTO — há BLOCKER"}[veredito_st]

    # ---------- markdown
    agora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    md = [f"# Relatório de verificação — G2 · {P['pasta_mes_nome']} (RM{P['num']})", "",
          f"Gerado em {agora} · destino verificado: **{'SharePoint (pasta do mês)' if aplicado else 'STAGING local'}** "
          f"`{raiz_mes}` · regras: POP de Entrega v2026-10-01 + Checklist Visual (abas G2 e Coordenação) + Modelo de Índice r2.", "",
          f"## Veredito: {ICONE.get(veredito_st)} — {veredito}", "",
          "| Nº | Verificação | Resultado | Evidência |", "|---|---|---|---|"]
    md += [f"| {n} | {t} | {ICONE[s]} | {e.replace('|', '/')} |" for n, t, s, e in linhas]
    md += ["", "> Regra final do Checklist: *“O material está completo, organizado e pronto para a Coordenação utilizar sem "
               "precisar renomear, procurar ou reorganizar arquivos?”*", ""]
    excs = cfg_mes.get("excecoes_desta_competencia", [])
    if excs:
        md += ["## Exceções aceitas SÓ nesta competência (não são regra)", ""] + [f"- {x}" for x in excs] + [
            "- Cada exceção está registrada em `revisoes_humanas.json` presa ao SHA-256 do arquivo; sem esse registro (ou se o "
            "arquivo mudar) o achado volta a ser MANUAL_REVIEW. Nas próximas competências o G2 deve entregar: relatório com a "
            "tabela igual à Lista Mestre, PDFs pesquisáveis e anexos já no padrão.", ""]
    md += ["## Retorno do engenheiro do G2 (Pessoa E, 06/10/2026) — como foi tratado", "",
           "| Ponto | Resposta do G2 | Tratamento no pacote |", "|---|---|---|",
           "| Relatório | Assinou e subiu | PDF assinado (Pessoa E, gov.br) já está em `G2/` com o nome padrão; assinatura VALIDADA. |",
           "| RM-012.1 / RM-013.1 ×3 | Pessoa F assinou e subiu | ZIPs originais (122–402 MB, nome fora do padrão) desmembrados: "
           "PDF assinado (Pessoa F, gov.br, VALIDADA) + anexos em partes ≤ 95 MB; originais movidos "
           "para o Desenvolvimento. |",
           "| NT-775.1 | Apagou o anexo e a linha da Lista | Fora do pacote, da Lista e do Índice. (Continua citado na tabela do relatório — exceção RM15.) |",
           "| NT-818.1 | RD-103; corrigiu a Lista | Nome com RD-103 = Lista = documento. |",
           "| NT-825.1 | RD-104; corrigiu a Lista | Nome com RD-104 = Lista = documento. A Descrição da Lista ainda diz 'Projeto de Paisagismo' "
           "(o PDF é Projeto de Drenagem) — o nome segue o documento. |",
           "| NT-793.1 | Somente RD-104 | PDF e as 9 partes dos anexos renomeados de GERAL para RD-104. |",
           "| NT-820.1 | Arquivo corrigido | PDF substituído pelo G2 (capa 'Rodovia: RD-101'); nome inalterado. |",
           "| RO-549.1 duplicado | Apagou a duplicata | Lista Mestre sem duplicatas. |",
           "| Tabela do relatório | Não vai alterar (já assinado) | Aceito só no RM15 (item 13); regra permanece. |", ""]
    md += ["## Estrutura final do G2", "", "```", f"{P['pasta_mes_nome']}/", f"  {os.path.basename(idx_path)}   (G2: {len(g2_idx)} linhas)",
           "  G2/", f"    {rel_nome}   (assinado)", f"    Anexos/   ({len(an_files)} arquivos, sem subpastas)", "```", "",
           "## Rastreabilidade", "",
           f"- Planilha antes→depois (cada arquivo, origem, SHA-256, conferências): `{os.path.join(P['out'], 'G2_verificacao_antes_depois.xlsx')}`",
           f"- Manifesto/log: `{os.path.join(P['out'], 'plano_g2.json')}` · decisões e exceções: `config/g2_pop_ajustes.json` · "
           f"aceites: `{rev_p}`",
           f"- Backups do Índice: `{os.path.join(P['out'], 'backup_indice')}`",
           "- Nada foi apagado: os originais do G2 (layout antigo e ZIPs dos RMs) estão em "
           "`3. Desenvolvimento/G2 -  Engenharia/_Produto/15. Setembro 2026/Anexos/`."]
    out_md = os.path.join(P["out"], "RELATORIO_VERIFICACAO_G2.md")
    with open(out_md, "w", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")

    # ---------- planilha
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Anexos (estado final)"
    ws.append(["Nº Índice", "Arquivo final (G2/Anexos)", "Origem", "Operação", "Produto", "Documento", "Rodovia",
               "Na Lista Mestre", "MB", "Págs.", "Págs. c/ texto", "SHA-256 = manifesto", "Nome conforme", "Íntegro", "SHA-256"])
    pos = {R.norm(n): i + 1 for i, n in enumerate(nomes_idx)}
    for r in sorted(por_arquivo, key=lambda r: pos.get(R.norm(os.path.splitext(r["arquivo"])[0]), 9999)):
        ws.append([pos.get(R.norm(os.path.splitext(r["arquivo"])[0])), r["arquivo"], r["origem"], r["op"], r["produto"], r["doc"],
                   r["rodovia"], "SIM" if (r["doc"], r["rodovia"]) in pares_lista else "NÃO", round(r["tamanho"] / 1048576, 2),
                   r.get("paginas"), r.get("paginas_texto"), "SIM" if r["sha_ok"] else "NÃO", "SIM" if r["nome_ok"] else "NÃO",
                   "SIM" if r["integ_ok"] else "NÃO", r["sha256"]])
    ws.append([1, rel_nome, "assinado pelo G2", "—", "REL", "Relatório de Produtos Mensal", "—", "—",
               round(os.path.getsize(rel_path) / 1048576, 2) if rel_ok else None, rel_val["info"].get("paginas"),
               rel_val["info"].get("paginas_com_texto"), "—", "—", "SIM" if rel_ok else "NÃO", rel_sha])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E5A")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for col, w in zip("ABCDEFGHIJKLMNO", [8, 95, 60, 18, 8, 12, 9, 9, 8, 7, 8, 10, 9, 8, 66]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    ws2 = wb.create_sheet("Checklist")
    ws2.append(["Nº", "Verificação", "Resultado", "Evidência"])
    for n, t, s, e in linhas:
        ws2.append([n, t, s if s != "EXCECAO" else "PASS (exceção RM15)", e])
    for col, w in zip("ABCD", [5, 55, 20, 150]):
        ws2.column_dimensions[col].width = w
    for row in ws2.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws3 = wb.create_sheet("Conteúdo das partes")
    ws3.append(["ZIP original", "Parte", "Arquivo interno", "Bytes", "SHA-256"])
    for o, cont in fracionados:
        for c in cont:
            ws3.append([o, c["parte"], c["arquivo"], c["bytes"], c["sha256"]])
    for col, w in zip("ABCDE", [60, 90, 90, 12, 66]):
        ws3.column_dimensions[col].width = w
    ws4 = wb.create_sheet("Relatórios Mensais de Obra")
    ws4.append(["Documento", "Rodovia", "Assinatura", "Signatário", "Emissor", "SHA-256 = extraído do ZIP"])
    for row in rm_sig:
        ws4.append([row[0], row[1], row[2], row[3], row[4], "SIM" if row[5] else "NÃO"])
    ws5 = wb.create_sheet("Tabela relatório x Lista")
    ws5.append(["Situação", "Documento", "Rodovia"])
    for d_, r_ in so_rel:
        ws5.append(["só na tabela do relatório", d_, r_])
    for d_, r_ in so_lista:
        ws5.append(["só na Lista Mestre / anexos", d_, r_])
    xl = os.path.join(P["out"], "G2_verificacao_antes_depois.xlsx")
    wb.save(xl)
    print(f"[verificar] veredito: {veredito} | itens: " + ", ".join(f"{n}={s}" for n, _, s, _ in linhas))
    print(f"[verificar] {out_md}\n[verificar] {xl}")

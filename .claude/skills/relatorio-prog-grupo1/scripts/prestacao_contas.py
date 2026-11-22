# -*- coding: utf-8 -*-
"""
prestacao_contas.py — eventos de PRESTAÇÃO DE CONTAS para a Timeline do P1.1 (DEC-060, regra permanente, 16/11/2026).

A prestação de contas é um FLUXO ITERATIVO, não um evento único:
  envio (Programa) → análise da Órgão Contratante → retorno/resposta do Programa → nova análise → … → aprovação.
Uma mesma prestação (ex.: 3º trimestre da RD-101) gera eventos em VÁRIOS meses. A Timeline da competência
representa os EVENTOS OCORRIDOS NO MÊS — não só as prestações novas.

Fonte: planilha "CONTROLE DE PRAZOS - PRESTAÇÃO DE CONTAS" (importada antes pela regra DEC-059). Layout:
  linha de cabeçalho com 'REFERÊNCIA', 'PERÍODO', 'DATA DO ENVIO Programa', 'Nª ANALISE Órgão Contratante', 'Nº RETORNO Programa',
  …, 'APROVAÇÃO'. Cada linha = uma prestação: REFERÊNCIA = rodovia/obra (RD-xxx) ou 'Programa' (Estruturadora/
  Contrato). Células podem ter data, '-', ou texto multi-linha 'dd/mm/aaaa - Executora\\ndd/mm/aaaa - Programa'.

Nada é inferido além da fonte: o tipo do evento vem do CABEÇALHO da coluna; o ator ('Executora', 'Programa',
'Certificação') vem do texto da célula; o rótulo da referência é mantido como está na fonte (RD-102 ≠ RD-102A no
cadastro do Controle Timeline). A redação só usa o que a coluna/célula comprova.
"""
import datetime, re, unicodedata

RX_DATA_TXT = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})(?:\s*-\s*([^\n\r]+))?")
RX_COL = re.compile(r"(\d+)\s*[ªºoa°]?\s*(analise|retorno)", re.I)


def _norm(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s or "")) if unicodedata.category(c) != "Mn").lower().strip()


def _tipo_coluna(cab):
    """('ENVIO', None) | ('ANALISE', n) | ('RETORNO', n) | ('APROVACAO', None) | (None, None)."""
    c = _norm(cab)
    if "envio" in c:
        return "ENVIO", None
    if "aprova" in c:
        return "APROVACAO", None
    m = RX_COL.search(c.replace("ª", "").replace("º", ""))
    if m:
        return ("ANALISE" if m.group(2).startswith("anal") else "RETORNO"), int(m.group(1))
    return None, None


def _valores_celula(v):
    """Lista de (date, ator|None, texto_original) de uma célula."""
    if v is None:
        return []
    if isinstance(v, datetime.datetime):
        return [(v.date(), None, v.strftime("%d/%m/%Y"))]
    if isinstance(v, datetime.date):
        return [(v, None, v.strftime("%d/%m/%Y"))]
    s = str(v).strip()
    if not s or s == "-":
        return []
    out = []
    for m in RX_DATA_TXT.finditer(s):
        try:
            d = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
        ator = (m.group(4) or "").strip() or None
        out.append((d, ator, m.group(0).strip()))
    return out


def classificar_referencia(ref):
    """'OBRA' (rodovia RD-xxx) | 'ESTRUTURADORA' (Programa/contrato) | 'INDEFINIDA' (não se classifica pelo nome)."""
    r = _norm(ref).upper().replace(" ", "")
    if re.fullmatch(r"RD-?\d{2,3}[A-Z]?", r):
        return "OBRA"
    if r in ("PROGRAMA", "ESTRUTURADORA", "CONTRATO"):
        return "ESTRUTURADORA"
    return "INDEFINIDA"


def ler_planilha(path):
    """Lê a planilha e devolve a lista de TODOS os eventos (todas as competências)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    eventos = []
    for ws in wb.worksheets:
        cab_row, cols = None, {}
        for r in range(1, min(ws.max_row, 15) + 1):
            vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if any("refer" in _norm(v) for v in vals) and any("envio" in _norm(v) for v in vals):
                cab_row = r
                for c, v in enumerate(vals, start=1):
                    t, n = _tipo_coluna(v)
                    if t:
                        cols[c] = (t, n, str(v).strip())
                    elif "refer" in _norm(v):
                        cols[c] = ("REF", None, "")
                    elif "period" in _norm(v):
                        cols[c] = ("PERIODO", None, "")
                break
        if not cab_row:
            continue
        c_ref = next(c for c, t in cols.items() if t[0] == "REF")
        c_per = next((c for c, t in cols.items() if t[0] == "PERIODO"), None)
        for r in range(cab_row + 1, ws.max_row + 1):
            ref = ws.cell(r, c_ref).value
            if not ref:
                continue
            ref = str(ref).strip()
            periodo = str(ws.cell(r, c_per).value or "").strip() if c_per else ""
            for c, (tipo, n, cab) in sorted(cols.items()):
                if tipo in ("REF", "PERIODO"):
                    continue
                v = ws.cell(r, c).value
                if tipo == "APROVACAO" and isinstance(v, str) and _norm(v) in ("sim", "aprovado", "aprovada"):
                    # aprovação sem data: não gera evento datado (registrado como situação)
                    eventos.append({"referencia": ref, "periodo": periodo, "tipo": "APROVACAO_SEM_DATA",
                                    "rodada": None, "data": None, "ator": None, "classe": classificar_referencia(ref),
                                    "fonte": f"{ws.title}!{ws.cell(r, c).coordinate}", "valor": str(v)})
                    continue
                for d, ator, txt in _valores_celula(v):
                    eventos.append({"referencia": ref, "periodo": periodo, "tipo": tipo, "rodada": n, "data": d,
                                    "ator": ator, "classe": classificar_referencia(ref),
                                    "fonte": f"{ws.title}!{ws.cell(r, c).coordinate}", "valor": txt,
                                    "coluna": cab})
    return eventos


def eventos_da_competencia(eventos, ano, mes):
    return sorted((e for e in eventos if e["data"] and e["data"].year == int(ano) and e["data"].month == int(mes)),
                  key=lambda e: (e["data"], e["referencia"], e["tipo"]))


def identificacao(e):
    """'Prestação de Contas de Obra da RD-101 (3º Trimestre – março a maio de 2026)' |
    'Prestação de Contas da Estruturadora (Programa) referente a setembro de 2026'. Rótulo/período vêm da fonte ou das
    evidências documentais (aplicar_evidencias)."""
    per = (e.get("periodo") or "").strip()
    ref = e.get("referencia_doc") or e["referencia"]
    if e["classe"] == "OBRA":
        det = per.replace("TRIMESTRE", "Trimestre") if per else ""
        if e.get("periodo_doc"):
            det = (det + " – " if det else "") + e["periodo_doc"]
        return f"Prestação de Contas de Obra da {ref}" + (f" ({det})" if det else "")
    if e["classe"] == "ESTRUTURADORA":
        per = e.get("periodo_doc") or per.lower()
        return (f"Prestação de Contas da Estruturadora (Programa) referente a {per}" if per else
                "Prestação de Contas da Estruturadora (Programa)")
    return f"Prestação de Contas {ref}" + (f" ({per})" if per else "")


def redigir(e, proximo=None):
    """Texto institucional do evento, só com o que a fonte comprova. `proximo` = evento seguinte da mesma prestação
    (qualquer mês): se a análise da Órgão Contratante foi seguida de retorno do Programa, a análise teve apontamentos."""
    ident = identificacao(e)
    n = e.get("rodada")
    ator = (e.get("ator") or "").strip()
    if e["tipo"] == "ENVIO":
        if ator:
            return f"Envio ao Parceiro de {ator.lower()} da {ident}."
        return f"Envio ao Parceiro da {ident}."
    if e["tipo"] == "ANALISE":
        if ator and _norm(ator) not in ("orgao contratante",):
            return f"{n}ª análise da {ident} registrada pela {ator}."
        cont = ""
        if proximo and proximo["tipo"] == "RETORNO":
            cont = ", com apontamentos que demandaram nova manifestação do Programa"
        elif proximo and proximo["tipo"] == "APROVACAO":
            cont = ", seguida de aprovação"
        return f"Retorno da Órgão Contratante com a {n}ª análise da {ident}{cont}."
    if e["tipo"] == "RETORNO":
        if _norm(ator) == "executora":
            return f"Resposta da Executora aos apontamentos da {n}ª análise da Órgão Contratante na {ident}."
        return (f"Encaminhamento pelo Programa da resposta aos apontamentos da {n}ª análise da Órgão Contratante na {ident}, "
                f"para continuidade da análise.")
    if e["tipo"] == "APROVACAO":
        return f"Aprovação da {ident}."
    return f"{ident}: {e.get('valor')}"


def _com_ref(texto, e):
    r = e.get("documento")
    return texto[:-1] + f" ({r})." if r and texto.endswith(".") else texto


def aplicar_evidencias(eventos, ev):
    """Aplica as evidências documentais (SharePoint) — DEC-060 rev.: rótulo corrigido (o original fica em
    'referencia_fonte'), período documentado, referência do documento de cada evento e eventos complementares."""
    if not ev:
        return eventos
    for e in eventos:
        k = f"{e['referencia']}|{e.get('periodo','')}"
        rot = (ev.get("rotulos") or {}).get(e["referencia"])
        if rot:
            e["referencia_fonte"] = e["referencia"]; e["referencia_doc"] = rot["para"]; e["evidencia_rotulo"] = rot["evidencia"]
        if (ev.get("periodos") or {}).get(k):
            e["periodo_doc"] = ev["periodos"][k]
        rk = f"{k}|{e['tipo']}|{e.get('rodada') or ''}"
        refs = ev.get("referencias") or {}
        # com ator ('Executora'/'Programa'): só a chave com o ator — documento do Programa não se atribui à executora
        doc = refs.get(f"{rk}|{e['ator']}") if e.get("ator") else refs.get(rk)
        if doc:
            e["documento"] = doc
        fp = (ev.get("fontes_pasta") or {})
        fpk = f"{rk}|{e['ator']}" if e.get("ator") else rk
        if fp.get(fpk):
            e["fonte"] = e["fonte"] + " | " + fp[fpk]
    for x in ev.get("eventos_extra") or []:
        d = datetime.date.fromisoformat(x["data"])
        e = {"referencia": x["referencia"], "periodo": x.get("periodo", ""), "tipo": x["tipo"], "rodada": x.get("rodada"),
             "data": d, "ator": x.get("ator"), "classe": classificar_referencia(x["referencia"]),
             "fonte": x.get("fonte", "evidência documental"), "valor": "", "texto_fixo": x.get("texto")}
        k = f"{e['referencia']}|{e['periodo']}"
        if (ev.get("periodos") or {}).get(k):
            e["periodo_doc"] = ev["periodos"][k]
        eventos.append(e)
    return eventos


def _proximo(e, todos):
    """Evento seguinte da MESMA prestação (ordem natural das colunas: envio, análise 1, retorno 1, análise 2…)."""
    def ordem(x):
        t = {"ENVIO": 0, "ANALISE": 1, "RETORNO": 2, "APROVACAO": 9}.get(x["tipo"], 8)
        return ((x.get("rodada") or 0) * 3 + t) if t < 9 else 999
    mesma = [x for x in todos if x["referencia"] == e["referencia"] and x["periodo"] == e["periodo"] and x["data"]]
    seguintes = [x for x in mesma if ordem(x) > ordem(e)]
    return min(seguintes, key=ordem) if seguintes else None


def marcos_da_competencia(path, ano, mes, evidencias=None):
    """Eventos do mês já redigidos: [{data, referencia, classe, tipo, rodada, ator, texto, fonte}].
    `evidencias` = dict ou caminho de JSON (aplicar_evidencias)."""
    if isinstance(evidencias, str):
        import json
        with open(evidencias, encoding="utf-8") as f:
            evidencias = json.load(f)
    todos = aplicar_evidencias(ler_planilha(path), evidencias)
    out = []
    for e in eventos_da_competencia(todos, ano, mes):
        txt = e.get("texto_fixo") or _com_ref(redigir(e, _proximo(e, todos)), e)
        out.append(dict(e, texto=txt))
    return out


def agrupar_por_data(marcos):
    """Um marco de Timeline por DATA (o slide comporta ~1 marcador por data — DEC-025). Devolve
    [{data, empreendimentos, eventos:[…]}] em ordem cronológica."""
    por = {}
    for m in marcos:
        por.setdefault(m["data"], []).append(m)
    return [{"data": d, "empreendimentos": sorted({(m.get("referencia_doc") or m["referencia"]) if m["classe"] == "OBRA"
                                                   else "Geral" for m in ms}),
             "eventos": ms} for d, ms in sorted(por.items())]


if __name__ == "__main__":
    import json, sys
    sys.stdout.reconfigure(encoding="utf-8")
    p, a, m = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
    evid = sys.argv[4] if len(sys.argv) > 4 else None
    for x in marcos_da_competencia(p, a, m, evid):
        print(x["data"].strftime("%d/%m/%Y"), "|", x["referencia"], "|", x["tipo"], x.get("rodada") or "", "|",
              x.get("ator") or "", "|", x["texto"], "|", x["fonte"])

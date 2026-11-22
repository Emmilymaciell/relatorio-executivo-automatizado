# -*- coding: utf-8 -*-
"""
medicao_pdf.py — lê o QUADRO DA MEDIÇÃO (planilha "Nª MEDIÇÃO" do Parceiro) impresso como PDF vetorial — p.ex. na última
página do "Relatório de Gestão do Contrato" do Parceiro — e devolve os valores por produto, grupo e total (DEC-062).

Uso como FONTE PROVISÓRIA quando o Excel editável oficial da medição ainda não chegou: os números vêm do PDF, a
planilha editável é RECONSTRUÍDA a partir da do mês anterior (reconstruir_planilha_medicao.py) e conferida campo a
campo com este extrato. Só lê o que está impresso; célula vazia/'-' => None (nunca estimado).
"""
import re, sys

COLS = ["quant", "valor_produto", "valor_aditivo", "medidos_ate_aditivo", "valor_total_contrato",
        "pct_mes", "pct_acum", "quant_medida", "nota", "valor_final", "valor_mes", "valor_acumulado"]
RX_ID = re.compile(r"^[1-5]\.[1-6]$")


def _num(s):
    s = s.strip().replace("R$", "").strip()
    if s in ("", "-"):
        return None
    pct = s.endswith("%")
    s = s.rstrip("%").replace(".", "").replace(",", ".")
    try:
        v = float(s)
    except ValueError:
        return None
    return v / 100 if pct else v


def localizar_pagina(pdf_path):
    """(índice da página, nº da medição) da página que tem o título 'Nª MEDIÇÃO' + 'VALOR ACUMULADO'."""
    import fitz
    d = fitz.open(pdf_path)
    for i in range(d.page_count - 1, -1, -1):
        t = d[i].get_text()
        m = re.search(r"(\d{1,2})\s*ª\s*MEDI[ÇC][ÃA]O", t, re.I)
        if m and "VALOR" in t.upper() and "ACUMULADO" in t.upper():
            return i, int(m.group(1))
    return None, None


def extrair(pdf_path, pagina=None):
    """{'medicao': N, 'pagina': p (1-based), 'produtos': {id: {...}}, 'grupos': {g: {...}}, 'total': {...},
    'pct_execucao': {...}, 'observacao': str}."""
    import fitz
    d = fitz.open(pdf_path)
    if pagina is None:
        pagina, nmed = localizar_pagina(pdf_path)
    else:
        nmed = None
    pg = d[pagina]
    words = pg.get_text("words")  # x0,y0,x1,y1,word,block,line,wno
    RXN = re.compile(r"^[\d\.\,\-]+%?$")

    def linhas_por_y(ws, tol=3.0):
        out = []
        for w in sorted(ws, key=lambda w: (w[1] + w[3]) / 2):
            yc = (w[1] + w[3]) / 2
            if out and abs(out[-1][0] - yc) <= tol:
                out[-1][1].append(w)
            else:
                out.append([yc, [w]])
        return out

    linhas = linhas_por_y(words)
    # linhas de produto: começam com o ID 'g.p'; o limite esquerdo dos números = após a palavra da unidade
    def numeros(ws, xmin):
        return [w for w in sorted(ws, key=lambda w: w[0]) if w[0] > xmin and RXN.match(w[4])]

    brutas = {}
    for yc, ws in linhas:
        ids = [w for w in ws if RX_ID.match(w[4])]
        if not ids:
            continue
        uni = [w for w in ws if w[4].upper() in ("MENSAL", "OBRA", "MÊS", "MÊS")]
        xmin = max(w[2] for w in uni) if uni else ids[0][2] + 250
        brutas[ids[0][4]] = (yc, numeros(ws, xmin))
    # colunas = bordas direitas (x1) da linha mais completa (12 valores)
    ref = max(brutas.values(), key=lambda t: len(t[1]))[1]
    if len(ref) != len(COLS):
        raise ValueError(f"layout inesperado: linha de referência com {len(ref)} valores (esperado {len(COLS)})")
    xs = [w[2] for w in ref]

    def atribuir(nums):
        cel = {c: None for c in COLS}
        for w in nums:
            j = min(range(len(xs)), key=lambda i: abs(xs[i] - w[2]))
            if abs(xs[j] - w[2]) < 18:
                cel[COLS[j]] = _num(w[4])
        return cel

    produtos = {pid: atribuir(nums) for pid, (yc, nums) in brutas.items()}
    # linhas de GRUPO / TOTAL / % execução: rótulo + valores nas colunas
    grupos, total, pct_exec = {}, None, None
    for yc, ws in linhas:
        txt = " ".join(w[4] for w in sorted(ws, key=lambda w: w[0])).upper()
        cel = atribuir([w for w in ws if RXN.match(w[4])])
        m = re.search(r"GRUPO\s+(\d)", txt)
        if m and not any(RX_ID.match(w[4]) for w in ws):
            grupos[int(m.group(1))] = cel
        elif "TOTAL DA MEDI" in txt:
            total = cel
        elif "EXECU" in txt and "CONTRATO" in txt:
            pct_exec = cel
    t = pg.get_text()
    obs = t[t.upper().find("OBSERVA"):].strip() if "OBSERVA" in t.upper() else ""
    return {"medicao": nmed, "pagina": pagina + 1, "produtos": produtos, "grupos": grupos, "total": total,
            "pct_execucao": pct_exec, "observacao": obs}


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    import json
    r = extrair(sys.argv[1])
    print("medição", r["medicao"], "página", r["pagina"])
    for k, v in r["produtos"].items():
        print(k, v)
    print("GRUPOS", json.dumps(r["grupos"]))
    print("TOTAL", r["total"]); print("%EXEC", r["pct_execucao"]); print(r["observacao"])

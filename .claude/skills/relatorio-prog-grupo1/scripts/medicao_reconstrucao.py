# -*- coding: utf-8 -*-
"""
medicao_reconstrucao.py — Planilha de Medição RECONSTRUÍDA a partir do quadro impresso da medição (DEC-062).

Quando o Excel editável oficial da medição N (Parceiro/Pessoa H) ainda não chegou, mas o quadro da medição N está impresso
num documento oficial (ex.: última página do "Relatório de Gestão do Contrato" do Parceiro):

  reconstruir  -> copia a Planilha de Medição do mês anterior (modelo; preserva abas, estilos, fórmulas, mesclagens),
                  replica o bloco da medição N-1 como bloco da medição N (Excel COM, fórmulas relativas), preenche SÓ
                  os insumos da medição (QUANT MEDIDA e NOTA AVALIATIVA por produto) com o que está no PDF, recalcula e
                  CONFERE campo a campo (por produto, grupo, empresa, total e % de execução) contra o PDF.
                  Qualquer divergência > R$ 0,01 / 0,01 p.p. => FALHA (a planilha não é promovida).
  comparar     -> quando o Excel OFICIAL chegar: compara o bloco da medição N do oficial com o da reconstruída
                  (valores, insumos, fórmulas-chave, totais). Equivalentes => o oficial passa a ser a fonte canônica.

A proveniência fica gravada (a) nas propriedades do arquivo (Comentários/Assunto) e (b) em
01_FONTES/EXTERNAS/_proveniencia.json + relatório de conferência em 05_LOGS. No relatório do G1 não se menciona a
reconstrução (decisão da usuária, 16/11/2026): o quadro financeiro é apresentado normalmente.

Uso:
  python medicao_reconstrucao.py reconstruir --pdf <pdf> --modelo <planilha N-1.xlsx> --saida <planilha N.xlsx>
  python medicao_reconstrucao.py comparar --oficial <xlsx> --reconstruida <xlsx> --medicao N
"""
import argparse, json, os, re, shutil, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
ABA = "Planilha de Medição - SUPRIMIDA"
# deslocamentos dentro do bloco de 7 colunas da aba SUPRIMIDA (DEC-034)
OFF = {"pct_mes": 0, "pct_acum": 1, "quant_medida": 2, "nota": 3, "valor_final": 4, "valor_mes": 5, "valor_acumulado": 6}
TOL_RS, TOL_PCT = 0.01, 0.0001


def _col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def blocos(ws_values):
    """{N: coluna inicial} pelos cabeçalhos 'Nª MEDIÇÃO' da linha 2 (openpyxl)."""
    out = {}
    for c in ws_values[2]:
        m = re.match(r"\s*(\d+)\s*ª\s*MEDI", str(c.value or ""), re.I)
        if m:
            out[int(m.group(1))] = c.column
    return out


def linhas_produto(ws):
    """{'1.1': linha, …} pela coluna D; grupos/total por rótulo na coluna H."""
    prods, grupos, total = {}, {}, None
    for r in range(5, 32):
        d = str(ws.cell(r, 4).value or "").strip()
        if re.fullmatch(r"[1-5]\.[1-6]", d):
            prods[d] = r
        h = str(ws.cell(r, 8).value or "").strip().upper()
        m = re.match(r"GRUPO\s*(\d)", h)
        if m:
            grupos[int(m.group(1))] = r
        if h == "TOTAL":
            total = r
    return prods, grupos, total


def ler_bloco(path, n):
    """Valores (data_only) do bloco da medição n: produtos, grupos, total, %execução, empresas, colunas fixas."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[ABA]
    b = blocos(ws)
    if n not in b:
        raise KeyError(f"{os.path.basename(path)}: não há bloco '{n}ª MEDIÇÃO' na aba {ABA}")
    c0 = b[n]
    prods, grupos, total = linhas_produto(ws)
    g = lambda r, k: ws.cell(r, c0 + OFF[k]).value
    out = {"bloco": _col(c0), "produtos": {}, "grupos": {}, "empresas": {}}
    for pid, r in prods.items():
        out["produtos"][pid] = {k: g(r, k) for k in OFF}
        out["produtos"][pid].update(quant=ws.cell(r, 7).value, valor_produto=ws.cell(r, 8).value,
                                    valor_aditivo=ws.cell(r, 9).value, medidos_ate_aditivo=ws.cell(r, 10).value,
                                    valor_total_contrato=ws.cell(r, 11).value)
    for gn, r in grupos.items():
        out["grupos"][gn] = {k: g(r, k) for k in ("valor_final", "valor_mes", "valor_acumulado")}
        out["grupos"][gn].update(valor_aditivo=ws.cell(r, 9).value, medidos_ate_aditivo=ws.cell(r, 10).value,
                                 valor_total_contrato=ws.cell(r, 11).value)
    out["total"] = {k: g(total, k) for k in ("valor_final", "valor_mes", "valor_acumulado")}
    out["total"].update(valor_aditivo=ws.cell(total, 9).value, medidos_ate_aditivo=ws.cell(total, 10).value,
                        valor_total_contrato=ws.cell(total, 11).value)
    out["pct_execucao"] = {k: g(total - 1, k) for k in ("valor_final", "valor_mes", "valor_acumulado")}
    for r in range(total + 1, total + 9):
        nome = str(ws.cell(r, c0 + 1).value or "").strip()
        if nome:
            out["empresas"][nome] = g(r, "valor_mes")
    return out


def conferir(planilha, pdf):
    """Lista de divergências (campo, planilha, pdf) entre o bloco lido da planilha e o extrato do PDF."""
    dif, ok = [], 0

    def cmp(rotulo, a, b, pct=False):
        nonlocal ok
        a0 = 0.0 if a in (None, "") else float(a)
        b0 = 0.0 if b in (None, "") else float(b)
        tol = 0.00005 if pct else TOL_RS
        # o PDF imprime % com 2 casas: compara arredondado
        if pct:
            a0, b0 = round(a0, 4), round(b0, 4)
        if abs(a0 - b0) > tol + 1e-9:
            dif.append({"campo": rotulo, "planilha": a, "pdf": b})
        else:
            ok += 1

    for pid, p in pdf["produtos"].items():
        q = planilha["produtos"].get(pid)
        if q is None:
            dif.append({"campo": f"{pid}", "planilha": "ausente", "pdf": "presente"}); continue
        for k in ("valor_final", "valor_mes", "valor_acumulado", "valor_aditivo", "medidos_ate_aditivo",
                  "valor_total_contrato", "valor_produto", "quant", "quant_medida"):
            cmp(f"{pid}.{k}", q.get(k), p.get(k))
        for k in ("pct_mes", "pct_acum", "nota"):
            if p.get(k) is not None or k != "nota":
                cmp(f"{pid}.{k}", q.get(k), p.get(k), pct=True)
    for gn, gp in pdf["grupos"].items():
        for k in ("valor_final", "valor_mes", "valor_acumulado", "valor_aditivo", "medidos_ate_aditivo",
                  "valor_total_contrato"):
            cmp(f"GRUPO {gn}.{k}", planilha["grupos"].get(gn, {}).get(k), gp.get(k))
    for k in ("valor_final", "valor_mes", "valor_acumulado", "valor_aditivo", "medidos_ate_aditivo",
              "valor_total_contrato"):
        cmp(f"TOTAL.{k}", planilha["total"].get(k), pdf["total"].get(k))
    for k in ("valor_mes", "valor_acumulado"):
        cmp(f"%EXECUÇÃO.{k}", planilha["pct_execucao"].get(k), pdf["pct_execucao"].get(k), pct=True)
    return dif, ok


def reconstruir(pdf_path, modelo, saida, n=None, propriedade=""):
    import medicao_pdf
    import openpyxl
    ext = medicao_pdf.extrair(pdf_path)
    n = n or ext["medicao"]
    if os.path.exists(saida):
        raise FileExistsError(f"já existe (não sobrescrevo): {saida}")
    wbv = openpyxl.load_workbook(modelo, data_only=True)
    b = blocos(wbv[ABA])
    if n - 1 not in b:
        raise KeyError(f"modelo sem bloco da {n-1}ª medição")
    if n in b:
        raise ValueError(f"modelo já tem bloco da {n}ª medição")
    prods, _g, total = linhas_produto(wbv[ABA])
    c_prev = b[n - 1]
    c_new = c_prev + 7
    shutil.copy2(modelo, saida)
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(saida))
        ws = wb.Worksheets(ABA)
        last_row = max(total + 9, 39)
        src = ws.Range(f"{_col(c_prev)}1:{_col(c_prev + 6)}{last_row}")
        dst = ws.Range(f"{_col(c_new)}1")
        src.Copy(dst)  # fórmulas relativas: a medição N referencia a N-1 (acumulados) — igual ao padrão do modelo
        for i in range(7):
            ws.Columns(c_new + i).ColumnWidth = ws.Columns(c_prev + i).ColumnWidth
        ws.Cells(2, c_new).Value = f"{n}ª MEDIÇÃO"
        insumos = {}
        for pid, r in prods.items():
            p = ext["produtos"].get(pid, {})
            for k in ("quant_medida", "nota"):
                cel = ws.Cells(r, c_new + OFF[k])
                ant = cel.Value  # valor herdado do bloco N-1 (o insumo da medição anterior)
                novo = p.get(k)
                if novo is None:
                    # '-' ou vazio no PDF: mantém o padrão do modelo (0 onde o modelo tinha 0; vazio onde era vazio)
                    novo = 0 if (ant not in (None, "") and k == "quant_medida") else None
                cel.Value = novo if novo is not None else ""
                insumos[f"{pid}.{k}"] = {"medicao_anterior": ant, "medicao_nova": novo}
        if propriedade:
            try:
                wb.BuiltinDocumentProperties("Comments").Value = propriedade
                wb.BuiltinDocumentProperties("Subject").Value = f"{n}ª medição — planilha reconstruída (DEC-062)"
            except Exception:
                pass
        xl.CalculateFull()
        wb.Save(); wb.Close(False)
    finally:
        xl.Quit()
    # releitura dos valores calculados (o Excel gravou o cache) e conferência
    lido = ler_bloco(saida, n)
    dif, ok = conferir(lido, ext)
    return {"medicao": n, "pagina_pdf": ext["pagina"], "modelo": modelo, "saida": saida, "bloco": lido["bloco"],
            "insumos": insumos, "conferidos": ok, "divergencias": dif, "empresas": lido["empresas"],
            "observacao_pdf": ext["observacao"].split("\n")[:4], "extrato_pdf": ext}


def comparar(oficial, reconstruida, n):
    a, b = ler_bloco(oficial, n), ler_bloco(reconstruida, n)
    dif = []
    for sec in ("produtos", "grupos"):
        for k, va in a[sec].items():
            vb = b[sec].get(k, {})
            for c, x in va.items():
                y = vb.get(c)
                if isinstance(x, (int, float)) or isinstance(y, (int, float)):
                    if abs(float(x or 0) - float(y or 0)) > (TOL_PCT if "pct" in c or c == "nota" else TOL_RS):
                        dif.append({"campo": f"{sec}.{k}.{c}", "oficial": x, "reconstruida": y})
    for c in a["total"]:
        if abs(float(a["total"][c] or 0) - float(b["total"].get(c) or 0)) > TOL_RS:
            dif.append({"campo": f"total.{c}", "oficial": a["total"][c], "reconstruida": b["total"].get(c)})
    return {"medicao": n, "equivalentes": not dif, "divergencias": dif}


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("reconstruir"); s.add_argument("--pdf", required=True); s.add_argument("--modelo", required=True)
    s.add_argument("--saida", required=True); s.add_argument("--medicao", type=int); s.add_argument("--relatorio")
    s.add_argument("--propriedade", default="")
    s = sub.add_parser("comparar"); s.add_argument("--oficial", required=True)
    s.add_argument("--reconstruida", required=True); s.add_argument("--medicao", type=int, required=True)
    a = ap.parse_args(argv)
    if a.cmd == "reconstruir":
        r = reconstruir(a.pdf, a.modelo, a.saida, a.medicao, a.propriedade)
        if a.relatorio:
            with open(a.relatorio, "w", encoding="utf-8") as f:
                json.dump(r, f, ensure_ascii=False, indent=1, default=str)
        print(f"{r['medicao']}ª medição — bloco {r['bloco']} — conferidos {r['conferidos']} campos — "
              f"divergências {len(r['divergencias'])}")
        for d in r["divergencias"]:
            print("  DIVERGE", d)
        print("empresas (valor mês):", r["empresas"])
        return 1 if r["divergencias"] else 0
    r = comparar(a.oficial, a.reconstruida, a.medicao)
    print(json.dumps(r, ensure_ascii=False, indent=1, default=str))
    return 0 if r["equivalentes"] else 1


if __name__ == "__main__":
    sys.exit(main())

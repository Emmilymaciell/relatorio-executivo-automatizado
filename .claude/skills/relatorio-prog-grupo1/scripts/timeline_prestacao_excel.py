# -*- coding: utf-8 -*-
"""
timeline_prestacao_excel.py — leva os eventos de PRESTAÇÃO DE CONTAS da competência para o Controle Timeline
(aba 'Base Timeline') — DEC-060.

- Lê a planilha de prestação de contas IMPORTADA (fontes_externas, DEC-059) via prestacao_contas.py.
- Agrupa os eventos por DATA (1 marcador por data no slide — DEC-025) e grava UMA linha por data:
  Competência = 1º dia do mês do relatório; Data do Marco; Categoria 'Prestação de Contas'; Empreendimento(s);
  Título curto (para o slide; vem de TITULOS quando informado, senão de um título padrão); Detalhamento = a redação
  institucional COMPLETA de cada evento; Referência = planilha + células de origem; Incluir = 'Sim'.
- Idempotente: linhas desta competência com Categoria 'Prestação de Contas' e Referência iniciada por
  'Controle de Prazos – Prestação de Contas' são substituídas; nada mais é alterado (outras categorias e meses ficam).
- Grava numa CÓPIA (o original é publicado depois com `sharepoint.py enviar --substituir`, que faz backup).
Se o mês não tiver evento, NÃO grava nada (Timeline sem macroatividade segue o comportamento histórico).
"""
import datetime, os, shutil, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prestacao_contas as PC

REF_PREFIXO = "Controle de Prazos – Prestação de Contas"


def titulo_padrao(grupo):
    evs = grupo["eventos"]
    if len(evs) == 1:
        e = evs[0]
        ident = PC.identificacao(e).replace("Prestação de Contas do ", "Prestação de Contas do ").replace(" da RD-", " RD-")
        if e["tipo"] == "ENVIO":
            return f"Envio da {ident} ao Parceiro"
        if e["tipo"] == "ANALISE":
            return f"{e['rodada']}ª análise da Órgão Contratante na {ident}"
        if e["tipo"] == "RETORNO":
            return f"Resposta do Programa à {e['rodada']}ª análise da {ident}"
        if e["tipo"] == "APROVACAO":
            return f"Aprovação da {ident}"
    return "; ".join(PC.redigir(e).rstrip(".") for e in evs)


def gravar(xlsx_entrada, xlsx_saida, planilha_pc, ano, mes, titulos=None, evidencias=None):
    import openpyxl
    marcos = PC.marcos_da_competencia(planilha_pc, ano, mes, evidencias)
    grupos = PC.agrupar_por_data(marcos)
    if not grupos:
        return {"linhas": 0, "grupos": [], "motivo": "sem eventos de prestação de contas na competência"}
    wb = openpyxl.load_workbook(xlsx_entrada)   # só LEITURA (para achar linhas); a gravação é por COM
    ws = wb["Base Timeline"]
    comp = datetime.datetime(int(ano), int(mes), 1)
    # linhas que esta rotina gravou antes para a competência (serão limpas) + linhas livres
    limpar, livres = [], []
    for r in range(4, ws.max_row + 1):
        c = ws.cell(r, 3).value
        ref = str(ws.cell(r, 11).value or "")
        if isinstance(c, datetime.datetime) and (c.year, c.month) == (comp.year, comp.month) \
                and ref.startswith(REF_PREFIXO):
            limpar.append(r)
            livres.append(r)
        elif ws.cell(r, 3).value is None and ws.cell(r, 4).value is None:
            livres.append(r)
    livres.sort()
    feitas = []
    for g, r in zip(grupos, livres):
        d = g["data"]
        tit = (titulos or {}).get(d.strftime("%d/%m")) or titulo_padrao(g)
        det = " ".join(e["texto"] for e in g["eventos"])
        fontes = "; ".join(sorted({e["fonte"] for e in g["eventos"]}))
        docs = "; ".join(dict.fromkeys(e["documento"] for e in g["eventos"] if e.get("documento")))
        if docs:
            fontes = fontes + " | " + docs
        feitas.append({"linha": r, "data": d.strftime("%d/%m/%Y"), "titulo": tit, "detalhe": det,
                       "empreendimentos": "; ".join(g["empreendimentos"]),
                       "referencia": f"{REF_PREFIXO} ({fontes})"})
    wb.close()
    # GRAVAÇÃO via Excel COM (openpyxl descartaria partes do arquivo: listas, formatação condicional, desenhos).
    # O openpyxl só calculou QUAIS linhas e QUAIS valores; a cópia de saída é a original intacta + estes valores.
    shutil.copy2(xlsx_entrada, xlsx_saida)
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    try:
        wbc = xl.Workbooks.Open(os.path.abspath(xlsx_saida))
        ws = wbc.Worksheets("Base Timeline")
        for r in limpar:
            for col in range(3, 13):
                if col not in (5, 6):
                    ws.Cells(r, col).ClearContents()
        for f in feitas:
            r = f["linha"]; d = datetime.datetime.strptime(f["data"], "%d/%m/%Y")
            ws.Cells(r, 3).Value = (datetime.date(comp.year, comp.month, 1) - datetime.date(1899, 12, 30)).days
            ws.Cells(r, 4).Value = (d.date() - datetime.date(1899, 12, 30)).days   # serial puro (sem hora — DEC-039)
            ws.Cells(r, 4).NumberFormat = "dd/mm/yyyy"
            ws.Cells(r, 7).Value = "Prestação de Contas"
            ws.Cells(r, 8).Value = f["empreendimentos"]
            ws.Cells(r, 9).Value = f["titulo"]
            ws.Cells(r, 10).Value = f["detalhe"]
            ws.Cells(r, 11).Value = f["referencia"]
            ws.Cells(r, 12).Value = "Sim"
        try:
            wbc.Worksheets("Timeline Mensal").Range("C3").Value = (datetime.date(comp.year, comp.month, 1) - datetime.date(1899, 12, 30)).days
        except Exception:
            pass
        wbc.Save(); wbc.Close(False)
    finally:
        xl.Quit()
    return {"linhas": len(feitas), "grupos": feitas}


if __name__ == "__main__":
    import json
    sys.stdout.reconfigure(encoding="utf-8")
    ent, sai, pc, a, m = sys.argv[1:6]
    tit = json.load(open(sys.argv[6], encoding="utf-8")) if len(sys.argv) > 6 and sys.argv[6] != "-" else None
    evid = sys.argv[7] if len(sys.argv) > 7 else None
    print(json.dumps(gravar(ent, sai, pc, int(a), int(m), tit, evid), ensure_ascii=False, indent=1, default=str))

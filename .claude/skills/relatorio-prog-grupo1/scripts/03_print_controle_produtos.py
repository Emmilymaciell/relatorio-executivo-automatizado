# -*- coding: utf-8 -*-
"""
03_print_controle_produtos.py — (a) CRIA o anexo 'P1.1 - Controle de Entrega de Produtos de {mês}.xlsx'
(cópia do acumulado filtrada à Versão N, ordenada por data, IDs 1..n) e (b) gera o PRINT do corpo
(landscape, SEM coluna ID, sem cortar texto, paginação balanceada, ~220 dpi). Parametrizado (prog_config).
- Não altera o acumulado (só lê/copia). Anexo salvo no work e copiado ao SharePoint (versionado) fora do DRY_RUN.
- DRY_RUN: gera anexo/print SÓ no work; não copia para o SharePoint.
Preserva DEC-009 (sem ID, sem corte, balanceada) e DEC-016 (imagens recortadas depois em 08).
"""
import os, sys, shutil, atexit, math, collections
import win32com.client as win32
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

def uniq(p):
    if not os.path.exists(p): return p
    b,e=os.path.splitext(p); i=2
    while os.path.exists(f"{b}_v{i}{e}"): i+=1
    return f"{b}_v{i}{e}"

def main():
    N=cfg.numero(); mes=cfg.mes_min(); WK=cfg.work_dir()
    ACC=cfg.controle_demandas_path()
    anexo_work=os.path.join(WK,"02_TRABALHO",f"P1.1 - Controle de Entrega de Produtos de {mes}.xlsx")
    shutil.copy2(ACC, anexo_work)
    xl=win32.gencache.EnsureDispatch("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
    try: xl.AutomationSecurity=1
    except: pass
    atexit.register(lambda: (xl.Quit() if xl else None))

    # ===== (a) ANEXO: manter só a Versão N, congelar IDs, ordenar por data, renumerar 1..n =====
    wb=xl.Workbooks.Open(anexo_work,False,False); ws=wb.Worksheets("Controle Geral")
    last=max(ws.Cells(ws.Rows.Count,5).End(-4162).Row, ws.Cells(ws.Rows.Count,3).End(-4162).Row)
    idrng=ws.Range(ws.Cells(4,2),ws.Cells(last,2)); idrng.Value=idrng.Value  # congela fórmulas de ID
    if ws.AutoFilterMode: ws.AutoFilterMode=False
    ws.Range(ws.Cells(3,2),ws.Cells(last,11)).AutoFilter(Field=10, Criteria1="<>"+str(N))
    try: ws.Range(ws.Cells(4,2),ws.Cells(last,11)).SpecialCells(12).EntireRow.Delete()
    except Exception: pass
    if ws.AutoFilterMode: ws.AutoFilterMode=False
    last=ws.Cells(ws.Rows.Count,5).End(-4162).Row
    n=last-3
    # ordenar por Data Entrega (col I=9)
    rng=ws.Range(ws.Cells(4,2),ws.Cells(last,11)); ws.Sort.SortFields.Clear()
    ws.Sort.SortFields.Add(Key=ws.Range(ws.Cells(4,9),ws.Cells(last,9)),Order=1)
    ws.Sort.SetRange(rng); ws.Sort.Header=0; ws.Sort.Apply()
    ws.Range(ws.Cells(4,2),ws.Cells(last,2)).Value=tuple((i+1,) for i in range(n))  # renumerar 1..n
    for s in list(wb.Worksheets):
        if s.Name!="Controle Geral":
            try: s.Visible=0
            except: pass
    wb.Save(); wb.Close(True)
    cfg.log(f"anexo Controle de Entregas: {n} linhas (Versão {N}) -> {anexo_work}")
    # copiar anexo para o SharePoint (P1.1), versionado, fora do DRY_RUN
    if cfg.dry_run():
        cfg.log("DRY_RUN: anexo NÃO copiado ao SharePoint.")
    else:
        dst=uniq(os.path.join(cfg.anexo_p11_dir(), cfg.nome_anexo_g1("P1.1", f"Controle de Entrega de Produtos de {mes}", ".xlsx")))
        shutil.copy2(anexo_work,dst); cfg.log(f"anexo -> {dst}")

    # ===== (b) PRINT do corpo (landscape, sem ID) =====
    pw=os.path.join(WK,"02_TRABALHO","controle_print.xlsx"); shutil.copy2(anexo_work,pw)
    wb=xl.Workbooks.Open(pw,False,False); ws=wb.Worksheets("Controle Geral")
    last=ws.Cells(ws.Rows.Count,5).End(-4162).Row
    ws.Columns("A").Hidden=True; ws.Columns("B").Hidden=True; ws.Columns("K").Hidden=True  # spacer, ID, Versão
    for col,w in [("C",30),("D",13),("E",26),("F",17),("G",50),("H",13),("I",12),("J",13)]:
        ws.Columns(col).ColumnWidth=w; ws.Columns(col).Hidden=False
    data=ws.Range(ws.Cells(4,3),ws.Cells(last,10))
    data.Font.Size=9; data.WrapText=True; data.VerticalAlignment=-4160; data.HorizontalAlignment=-4131
    ws.Range("C3:J3").WrapText=True
    cpl={3:31,5:27,6:18,7:53}  # Produto/Documento/Tipo/Descrição
    for r in range(4,last+1):
        lines=1
        for c,cw in cpl.items():
            v=ws.Cells(r,c).Value
            if v: lines=max(lines, math.ceil(len(str(v))/cw))
        ws.Rows(r).RowHeight=lines*11.0+3
    ws.Rows("3").RowHeight=22
    # paginação balanceada (última página cheia)
    heights=[ws.Rows(r).RowHeight for r in range(4,last+1)]; total=sum(heights); target=total/12
    try: ws.ResetAllPageBreaks()
    except: pass
    cum=0.0
    for i,r in enumerate(range(4,last)):
        cum+=heights[i]
        if cum>=target*0.985:
            try: ws.HPageBreaks.Add(ws.Rows(r+1))
            except Exception: pass
            cum=0.0
    ps=ws.PageSetup; ps.Orientation=2; ps.Zoom=False; ps.FitToPagesWide=1; ps.FitToPagesTall=False
    ps.PrintArea="$C$1:$J$"+str(last); ps.PrintTitleRows="$1:$3"
    ps.LeftMargin=xl.InchesToPoints(0.25); ps.RightMargin=xl.InchesToPoints(0.25)
    ps.TopMargin=xl.InchesToPoints(0.3); ps.BottomMargin=xl.InchesToPoints(0.3); ps.CenterHorizontally=True
    wb.Save()
    pdf=os.path.join(WK,"04_PDFS_TEMP","controle_LS.pdf")
    if os.path.exists(pdf):
        try: os.remove(pdf)   # Excel não sobrescreve PDF pré-existente -> "documento não foi salvo" (DEC-038)
        except OSError: pass
    ws.ExportAsFixedFormat(0,pdf)
    wb.Close(True); xl.Quit()
    import fitz
    doc=fitz.open(pdf); outdir=os.path.join(WK,"03_IMAGENS","CONTROLE_PRODUTOS")
    for f in os.listdir(outdir):
        if f.endswith(".png"): os.remove(os.path.join(outdir,f))
    for i,pg in enumerate(doc): pg.get_pixmap(dpi=220).save(os.path.join(outdir,f"controle_p{i+1:02d}.png"))
    cfg.log(f"print Controle: {doc.page_count} páginas landscape"); doc.close()

if __name__=="__main__":
    main()

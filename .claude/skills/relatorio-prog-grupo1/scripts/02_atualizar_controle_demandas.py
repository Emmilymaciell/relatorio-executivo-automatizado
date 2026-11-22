# -*- coding: utf-8 -*-
"""
02_atualizar_controle_demandas.py — Acrescenta ao acumulado 'P.1.1 - Controle de demandas Programa.xlsx'
as linhas da competência (Versão = N), preservando as linhas antigas (DEC-002).
- Faz BACKUP com timestamp antes de qualquer alteração (obrigatório).
- IDEMPOTENTE: se já existirem linhas da Versão N (re-execução), remove-as antes de reanexar.
- Datas gravadas como SERIAL PURO (inteiro), sem hora e imune ao offset de fuso do win32com
  (senão saíam 03:00 e o Power BI errava); formato herdado da coluna (DEC-001/DEC-039). 'Sem data' como texto.
- IDs contínuos como VALOR (base_id + i + 1), nunca fórmula (o Power BI/usuário querem só o número) — DEC-039.
- DRY_RUN (PROG_DRY_RUN=1): NÃO abre/modifica o acumulado; só reporta o que faria.
Fonte das linhas: output/RM.../02_TRABALHO/entregas_rmN.json (rode 01 antes).
"""
import os, sys, json, shutil, datetime, atexit, collections
import win32com.client as win32
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

def main():
    N=cfg.numero(); WK=cfg.work_dir(); TR=os.path.join(WK,"02_TRABALHO")
    ACC=cfg.controle_demandas_path()
    src=os.path.join(TR,f"entregas_rm{N}.json")
    if not os.path.exists(src):
        cfg.log(f"ERRO: {src} não encontrado. Rode 01_consolidar_entregas.py antes."); sys.exit(1)
    rows=json.load(open(src,encoding="utf-8"))
    cfg.log(f"Atualizar acumulado: {len(rows)} linhas da Versão {N}. Arquivo: {ACC}")

    if cfg.dry_run():
        cfg.log("DRY_RUN: não vou abrir nem modificar o acumulado. (Faria backup + remover Versão N existente + acrescentar.)")
        return

    # BACKUP obrigatório (timestamp)
    ts=datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    bkp=os.path.join(WK,"06_BACKUP",f"P.1.1 - Controle de demandas PROG__BACKUP_{ts}.xlsx")
    shutil.copy2(ACC,bkp); cfg.log(f"backup: {bkp}")

    xl=win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible=False; xl.DisplayAlerts=False; xl.ScreenUpdating=False
    try: xl.AutomationSecurity=1
    except: pass
    atexit.register(lambda: (xl.Quit() if xl else None))
    wb=xl.Workbooks.Open(ACC, False, False)
    ws=wb.Worksheets("Controle Geral")

    # IDEMPOTÊNCIA: remover linhas já existentes da Versão N (re-execução)
    last=max(ws.Cells(ws.Rows.Count,5).End(-4162).Row, ws.Cells(ws.Rows.Count,3).End(-4162).Row)
    if ws.AutoFilterMode: ws.AutoFilterMode=False
    ws.Range(ws.Cells(3,2),ws.Cells(last,11)).AutoFilter(Field=10, Criteria1=str(N))  # col K = 10º a partir de B
    try:
        vis=ws.Range(ws.Cells(4,2),ws.Cells(last,11)).SpecialCells(12)  # xlCellTypeVisible
        vis.EntireRow.Delete(); cfg.log("removidas linhas pré-existentes da Versão N (idempotência)")
    except Exception:
        cfg.log("nenhuma linha pré-existente da Versão N (ok)")
    if ws.AutoFilterMode: ws.AutoFilterMode=False

    last=max(ws.Cells(ws.Rows.Count,5).End(-4162).Row, ws.Cells(ws.Rows.Count,3).End(-4162).Row)
    start=last+1; end=last+len(rows)
    cfg.log(f"última linha existente: {last}; ID atual: {ws.Cells(last,2).Value}")

    # herdar formato da última linha
    ws.Range(ws.Cells(last,2),ws.Cells(last,11)).Copy()
    ws.Range(ws.Cells(start,2),ws.Cells(end,11)).PasteSpecial(Paste=-4122)  # só formatos
    xl.CutCopyMode=False

    base_id=int(ws.Cells(last,2).Value)   # último ID existente (valor, mesmo que fosse fórmula)
    for i,r in enumerate(rows):
        rr=start+i
        ws.Cells(rr,2).Value=base_id+i+1           # ID contínuo como VALOR, nunca fórmula (DEC-039)
        ws.Cells(rr,3).Value=r["produto"]; ws.Cells(rr,4).Value=r["empresa"]
        ws.Cells(rr,5).Value=r["doc"];     ws.Cells(rr,6).Value=r["tipo"]
        ws.Cells(rr,7).Value=r["desc"];    ws.Cells(rr,8).Value=r["emp"]
        if r["data"]=="Sem data":
            ws.Cells(rr,9).Value="Sem data"
        else:
            # Serial de data PURO (inteiro): sem hora e imune ao offset de fuso que o win32com
            # aplica ao converter datetime -> COM (senão a data sai com 03:00). DEC-039.
            y,m,d=map(int,r["data"].split("-"))
            ws.Cells(rr,9).Value=(datetime.date(y,m,d)-datetime.date(1899,12,30)).days
        ws.Cells(rr,10).Value=r["status"]; ws.Cells(rr,11).Value=N
    # NÃO sobrescrever NumberFormat (herda o formato de data local — via PasteSpecial acima) — DEC-001
    wb.Save()

    lastnew=ws.Cells(ws.Rows.Count,5).End(-4162).Row
    kv=ws.Range(ws.Cells(4,11),ws.Cells(lastnew,11)).Value
    cnt=collections.Counter(str(v[0]).replace(".0","") if v[0] is not None else "" for v in kv)
    cfg.log(f"versões após: {dict(sorted(cnt.items(),key=lambda x:int(x[0]) if x[0].isdigit() else 99))}")
    cfg.log(f"data exemplo: {ws.Cells(start,9).Text} (numfmt {ws.Cells(start,9).NumberFormat}); "
            f"IDs {ws.Cells(start,2).Value}..{ws.Cells(end,2).Value}")
    wb.Close(SaveChanges=True); xl.Quit()
    cfg.log(f"acumulado atualizado: +{len(rows)} linhas Versão {N}")

if __name__=="__main__":
    main()

# -*- coding: utf-8 -*-
"""
06_linha_tempo_medicoes.py — Monta a 'Linha do Tempo de Medições — Status Financeiro'
e renderiza uma imagem por bloco de 5 medições (5+5+resto...). Só escreve arquivos locais.
Dados (valores por medição + status) vêm de cfg.dados() (config/dados_mes.py); NADA fixo aqui.
Preserva o layout validado (cores, tamanhos, larguras, DEC de negrito/estreito para o bloco final).
"""
import os, sys, atexit, re, unicodedata
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

def _norm(s):
    return "".join(c for c in unicodedata.normalize("NFD",str(s)) if unicodedata.category(c)!="Mn").upper().strip()

# Status por "mais avançado" na progressão (uma medição pode ter várias colunas de status no Excel:
# ex.: med1 = Faturado→Processado→Pago; med5 = Pago/Processado/Em análise). O status exibido é o de
# maior avanço (Pago > Processado > Faturado > Em análise). Reproduz exatamente o baseline validado.
_RANK={"NAO FATURADO":0,"NÃO FATURADO":0,"EM ANALISE":1,"EM ANÁLISE":1,"FATURADO":2,"PROCESSADO":3,"PAGO":4}
_LABEL={0:"Não faturado",1:"Em análise",2:"Processado",3:"Processado",4:"Pago"}

def medicoes_do_excel():
    """Lê valores e status de cada medição direto da aba STATUS_MEDICOES do Quadro Financeiro
    (DEC-030). Retorna (empresas, {med:[valores]}, {med:status}) ou None se a aba não existir.
    Colunas de cada 'MEDIÇÃO n' são descobertas pelo cabeçalho (respeitando células mescladas);
    o valor é o da 1ª coluna do grupo (iguais) e o status é o de maior avanço no grupo."""
    src=cfg.find_in_dir(cfg.anexo_p11_dir(), "p1.1", "quadro financeiro", ext=".xlsx")
    if not src or not os.path.exists(src): return None
    wb=openpyxl.load_workbook(src, data_only=True)
    ws=None
    for name in wb.sheetnames:
        if _norm(name).startswith("STATUS_MEDICOES"): ws=wb[name]; break
    if ws is None: return None
    def span(row,col):
        for mr in ws.merged_cells.ranges:
            if mr.min_row<=row<=mr.max_row and mr.min_col<=col<=mr.max_col:
                return list(range(mr.min_col, mr.max_col+1))
        return [col]
    # linha de cabeçalho = a que contém "MEDIÇÃO 1"
    hr=None
    for r in range(1, ws.max_row+1):
        for c in range(1, ws.max_column+1):
            v=ws.cell(r,c).value
            if v and re.match(r'MEDICAO\s*1\b', _norm(v)): hr=r; break
        if hr: break
    if hr is None: return None
    med_cols={}
    for c in range(1, ws.max_column+1):
        v=ws.cell(hr,c).value
        if not v: continue
        m=re.match(r'MEDICAO\s*(\d+)', _norm(v))
        if m and int(m.group(1)) not in med_cols:
            med_cols[int(m.group(1))]=span(hr,c)
    if not med_cols: return None
    srow=hr+1  # linha de status
    emp=[]; emp_rows=[]; r=hr+2
    while r<=ws.max_row:
        lab=ws.cell(r,1).value
        if lab in (None,""): break
        emp.append(str(lab).strip()); emp_rows.append(r)
        if _norm(lab)=="TOTAL": break
        r+=1
    vals={}; status={}
    for n,cols in med_cols.items():
        vcol=cols[0]
        vals[n]=[float(ws.cell(er,vcol).value or 0) for er in emp_rows]
        best=-1
        for c in cols:
            sv=ws.cell(srow,c).value
            if sv is None: continue
            rk=_RANK.get(_norm(sv))
            if rk is not None and rk>best: best=rk
        status[n]=_LABEL.get(best,"Pago") if best>=0 else "Pago"
    return emp, vals, status

DGREEN=PatternFill("solid",fgColor="1F7A54")
HDR=PatternFill("solid",fgColor="2E8B62")
PAGO=PatternFill("solid",fgColor="A9D08E")
PROC=PatternFill("solid",fgColor="FFE699")
ANALISE=PatternFill("solid",fgColor="F4B183")
NAOFAT=PatternFill("solid",fgColor="A6A6A6")   # Não faturado = CINZA (DEC-041)
TOTALF=PatternFill("solid",fgColor="1F7A54")
LEGP=PatternFill("solid",fgColor="A9D08E"); LEGPR=PatternFill("solid",fgColor="FFE699"); LEGA=PatternFill("solid",fgColor="F4B183"); LEGNF=PatternFill("solid",fgColor="A6A6A6")
thin=Side(style="thin",color="FFFFFF")
border=Border(left=thin,right=thin,top=thin,bottom=thin)
def money(v): return "R$ "+format(v,",.2f").replace(",","X").replace(".",",").replace("X",".")
def status_fill(s):
    return NAOFAT if s=="Não faturado" else (ANALISE if s=="Em análise" else (PROC if s=="Processado" else PAGO))

def main():
    ex=medicoes_do_excel()
    if ex:
        emp,vals,status=ex
        cfg.log(f"Linha do Tempo: dados lidos do Excel STATUS_MEDICOES do Quadro Financeiro "
                f"({len(vals)} medições, empresas={emp}) — DEC-030")
    else:
        D=cfg.dados()
        emp=list(getattr(D,"MEDICOES_EMPRESAS",["DYNA","CG","JURIDICO","CERT","TOTAL"]))
        vals=D.MEDICOES; status=D.MEDICOES_STATUS
        cfg.log("Linha do Tempo: aba STATUS_MEDICOES não encontrada no Excel — usando cfg.dados() (fallback)")
    meds_all=sorted(vals.keys())
    WK=cfg.work_dir()

    wb=openpyxl.Workbook()
    def build(ws, meds):
        a,b=meds[0],meds[-1]; ncol=1+len(meds); lastcol=get_column_letter(ncol)
        ws.merge_cells(f"A1:{lastcol}1")
        ws["A1"]="LINHA DO TEMPO DE MEDIÇÕES — STATUS FINANCEIRO"
        tsize=13 if len(meds)>=5 else 10
        ws["A1"].fill=DGREEN; ws["A1"].font=Font(b=True,color="FFFFFF",size=tsize); ws["A1"].alignment=Alignment("center","center")
        ws.merge_cells(f"A2:{lastcol}2")
        ws["A2"]=f"Contrato n.º 000/0000  |  Programa  |  Medições {a} a {b}"
        ws["A2"].fill=DGREEN; ws["A2"].font=Font(color="FFFFFF",size=10); ws["A2"].alignment=Alignment("center","center")
        ws["A4"]="  ● Pago"; ws["A4"].fill=LEGP; ws["A4"].font=Font(b=True,size=9)
        ws["B4"]="  ● Processado"; ws["B4"].fill=LEGPR; ws["B4"].font=Font(b=True,size=9)
        ws["C4"]="  ● Em análise"; ws["C4"].fill=LEGA; ws["C4"].font=Font(b=True,size=9)
        if ncol>=4:  # 4ª legenda (Não faturado) cabe a partir de 3 medições no bloco
            ws["D4"]="  ● Não faturado"; ws["D4"].fill=LEGNF; ws["D4"].font=Font(b=True,size=9)
        ws.cell(6,1,"Empresa").fill=HDR
        for j,m in enumerate(meds): ws.cell(6,2+j,f"MEDIÇÃO {m}").fill=HDR
        ws.cell(7,1,"Empresa").fill=HDR
        for j,m in enumerate(meds):
            c=ws.cell(7,2+j,status[m]); c.fill=status_fill(status[m]); c.font=Font(b=True,size=9)
        for cc in range(1,ncol+1):
            ws.cell(6,cc).font=Font(b=True,color="FFFFFF",size=10); ws.cell(6,cc).alignment=Alignment("center","center")
            ws.cell(7,cc).alignment=Alignment("center","center")
        for i,e in enumerate(emp):
            row=8+i; ce=ws.cell(row,1,e); ce.alignment=Alignment("center","center")
            istotal=(e=="TOTAL")
            ce.fill=TOTALF if istotal else PatternFill("solid",fgColor="F2F2F2")
            ce.font=Font(b=True,color="FFFFFF" if istotal else "000000",size=9)
            for j,m in enumerate(meds):
                v=vals[m][i]; c=ws.cell(row,2+j,money(v))
                c.alignment=Alignment("right","center"); c.font=Font(b=istotal, color="FFFFFF" if istotal else "000000", size=9)
                c.fill= TOTALF if istotal else status_fill(status[m])
        lastrow=7+len(emp)
        for r in range(6,lastrow+1):
            for cc in range(1,ncol+1): ws.cell(r,cc).border=border
        ws.column_dimensions["A"].width=13
        medw=17 if len(meds)>=5 else 22
        for j in range(len(meds)): ws.column_dimensions[get_column_letter(2+j)].width=medw
        for r in range(1,lastrow+1): ws.row_dimensions[r].height=20
        ws.print_area=f"A1:{lastcol}{lastrow}"
        ws.page_setup.orientation="portrait"; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1
        ws.sheet_properties.pageSetUpPr=PageSetupProperties(fitToPage=True)
        ws.page_margins.left=ws.page_margins.right=0.3; ws.page_margins.top=ws.page_margins.bottom=0.4

    blocks=[meds_all[i:i+5] for i in range(0,len(meds_all),5)]
    sheets=[]
    for bi,blk in enumerate(blocks):
        name=f"M{blk[0]}_{blk[-1]}"
        ws=wb.active if bi==0 else wb.create_sheet(name)
        if bi==0: ws.title=name
        build(ws,blk); sheets.append((name,f"linha_med_{bi+1}"))
    tmp=os.path.join(WK,"02_TRABALHO","linha_medicoes.xlsx"); wb.save(tmp)

    # valida TOTAL = soma das 4 empresas
    for m in meds_all:
        s=sum(vals[m][:4]); assert abs(s-vals[m][4])<0.05, (m,s,vals[m][4])
    cfg.log("totais das medições conferem")

    xl=win32.gencache.EnsureDispatch("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
    atexit.register(lambda: xl.Quit())
    wbx=xl.Workbooks.Open(tmp)
    import fitz
    outdir=os.path.join(WK,"03_IMAGENS","FINANCEIRO")
    for sh,tag in sheets:
        ws=wbx.Worksheets(sh); pdf=os.path.join(WK,"04_PDFS_TEMP",f"{tag}.pdf")
        if os.path.exists(pdf):
            try: os.remove(pdf)   # Excel não sobrescreve PDF pré-existente (DEC-038)
            except OSError: pass
        ws.ExportAsFixedFormat(0,pdf)
        doc=fitz.open(pdf); doc[0].get_pixmap(dpi=220).save(os.path.join(outdir,f"{tag}.png")); doc.close()
        cfg.log(f"render {tag} ({sh})")
    wbx.Close(False); xl.Quit()
    cfg.log(f"linha do tempo de medições concluída ({len(sheets)} blocos)")

if __name__=="__main__":
    main()

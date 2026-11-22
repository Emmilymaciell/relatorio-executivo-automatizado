# -*- coding: utf-8 -*-
"""
15_gerar_painel_financeiro.py — PRODUZ o painel 'P1.1 - Quadro Financeiro do mês de {mês}.xlsx'
a partir de (a) TEMPLATE = painel do MÊS ANTERIOR (preserva formatação + gráfico) e (b) DADOS =
'Medicao Parceiro\\Planilha de Medição - Estruturadora - {Mês}.xlsx' (fonte oficial). Ver DEC-034.
Edita via Excel COM (preserva o gráfico). NÃO usa o espelho Parceiro.

Regras validadas contra o painel-oráculo de julho/2026 (12ª medição):
- por produto (BASE_PARCEIRO): %acum, valor final, valor acumulado direto do bloco da medição;
- por empresa: valor mês direto; ACUM = soma do valor acum dos produtos da empresa; %=acum/contrato;
  saldo = contrato − acum (contrato estático do template); TOTAL por soma;
- gráfico (PAINEL C:G, l.84-87) = contrato(estático)/acum/saldo/% por empresa;
- STATUS_MEDICOES: coluna da medição N = valor mês por empresa + status (status_medicao);
- "Qntd. Acum." (BASE col H) = recorrência template+incremento; no mês de transição use --seed-qntd.

Modos:
  python 15_gerar_painel_financeiro.py --build            # copia template + preenche via COM -> WORK/07_SAIDA
  python 15_gerar_painel_financeiro.py --build --validate ORACULO.xlsx --seed-qntd ORACULO.xlsx
  python 15_gerar_painel_financeiro.py --validate ORACULO.xlsx          # só valida a LÓGICA (openpyxl)
Env: PROG_MEDICAO_N, PROG_PAINEL_TEMPLATE, PROG_PLANILHA_MEDICAO.
"""
import os, sys, re, shutil, json
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
import openpyxl
from openpyxl.utils import get_column_letter

EMP_ALIAS={"DYNA":"Dynatest","CG":"Consultoria","CERT":"Certificadora","JURIDICO":"Escritório Jurídico"}
EMP_ORDEM=["Dynatest","Consultoria","Escritório Jurídico","Certificadora"]     # ordem canônica DYNA,CG,JURIDICO,CERT
EMP_SIGLA={"Dynatest":"DYNA","Consultoria":"CG","Escritório Jurídico":"JURIDICO","Certificadora":"CERT"}
# Cor de preenchimento por status (RGB hex) — usada nas células de Status do RESUMO + legenda.
STATUS_COR={"Pago":"A9D18E","Processado":"F2CF00","Em análise":"F4B183","Não faturado":"A6A6A6"}  # Não faturado = CINZA (DEC-041)
def _com_color(rgbhex):  # RGB hex -> inteiro BGR do Excel COM
    r=int(rgbhex[0:2],16); g=int(rgbhex[2:4],16); b=int(rgbhex[4:6],16); return b*65536+g*256+r

def medicao_n():
    v=os.environ.get("PROG_MEDICAO_N")
    if v: return int(v)
    try:
        mn=getattr(cfg.dados(),"MEDICAO_NUM",None)
        if mn: return int(mn)
    except Exception: pass
    return cfg.numero()-1

def planilha_medicao_path():
    env=os.environ.get("PROG_PLANILHA_MEDICAO")
    if env and os.path.exists(env): return env
    dp=cfg.dir_desenv_produtos()
    mi=cfg._find_sub(dp,"medicao parceiro") or os.path.join(dp,"Medicao Parceiro")
    mes=cfg.mes_min()
    return cfg.find_in_dir(mi,"planilha de medi", mes, ext=".xlsx") or cfg.find_in_dir(mi,"planilha de medi", ext=".xlsx")

def painel_template_path():
    env=os.environ.get("PROG_PAINEL_TEMPLATE")
    if env and os.path.exists(env): return env
    prod=cfg.dir_produtos(); num_ant=cfg.numero()-1; alvo=None
    for s in sorted(os.listdir(prod)):
        if s.strip().startswith(f"{num_ant}.") and os.path.isdir(os.path.join(prod,s)): alvo=os.path.join(prod,s); break
    if not alvo: return None
    anex=os.path.join(alvo,"G1","Anexos")
    # layout legado (Anexos/P1.1 …) ou POP (ANEXOS única, produto no nome — DEC-043)
    p11=cfg._find_sub(anex,"p1.1") if os.path.isdir(anex) else None
    if p11: return cfg.find_in_dir(p11,"quadro financeiro", ext=".xlsx")
    if not os.path.isdir(anex): return None
    achado=cfg.find_in_dir(anex,"p1.1","quadro financeiro", ext=".xlsx")
    if achado: return achado
    # mês já publicado no padrão POP: o painel está num ZIP individual (DEC-043) -> extrai para 01_FONTES (DEC-063)
    z=cfg.find_in_dir(anex,"p1.1","quadro financeiro", ext=".zip")
    if z:
        import zipfile
        dst=os.path.join(cfg.work_dir(),"01_FONTES","painel_mes_anterior"); os.makedirs(dst,exist_ok=True)
        with zipfile.ZipFile(z) as zf:
            xs=[n for n in zf.namelist() if n.lower().endswith(".xlsx")]
            if xs:
                zf.extract(xs[0],dst); return os.path.join(dst,xs[0])
    return None

def extrair_planilha(planilha, N):
    wb=openpyxl.load_workbook(planilha, data_only=True)
    if N>=7:
        ws=[w for w in wb.worksheets if "suprimida" in w.title.lower()][0]
        off=dict(pmes=0,pacum=1,quant=2,nota=3,vfinal=4,vmes=5,vacum=6)
    else:
        ws=[w for w in wb.worksheets if "global" in w.title.lower()][0]
        off=dict(pmes=0,pacum=1,vmes=2,vacum=3)
    blk=None
    for j in range(1,ws.max_column+1):
        v=ws.cell(2,j).value
        if v and "MEDI" in str(v).upper():
            m=re.search(r'(\d+)',str(v))
            if m and int(m.group(1))==N: blk=j
    if not blk: raise RuntimeError(f"bloco da {N}ª medição não achado na Planilha")
    def PV(row,key):
        if key not in off: return None
        v=ws.cell(row,blk+off[key]).value
        return float(v) if isinstance(v,(int,float)) else None
    prod={}
    for r in range(4,40):
        idv=ws.cell(r,4).value
        if idv and re.match(r'^\d+\.\d+$',str(idv)):
            fx=lambda c: (float(ws.cell(r,c).value) if isinstance(ws.cell(r,c).value,(int,float)) else None)
            prod[str(idv)]=dict(pacum=PV(r,'pacum'),pmes=PV(r,'pmes'),quant=PV(r,'quant'),
                                 nota=PV(r,'nota'),vfinal=PV(r,'vfinal'),vmes=PV(r,'vmes'),vacum=PV(r,'vacum'),
                                 # colunas fixas do produto (G=QUANT, H=VALOR DO PRODUTO, K=VALOR TOTAL DO CONTRATO)
                                 quantc=fx(7), vprod=fx(8), contrato=fx(11))
    comp={}
    for r in range(32,40):
        for cc in range(max(1,blk-6),blk+2):
            nv=ws.cell(r,cc).value
            if nv and any(k in str(nv).lower() for k in ("dynatest","juridico","certificadora","consultoria")):
                key={"dynatest":"Dynatest","juridico":"Escritório Jurídico","certificadora":"Certificadora","consultoria":"Consultoria"}[
                    next(k for k in ("dynatest","juridico","certificadora","consultoria") if k in str(nv).lower())]
                comp[key]=PV(r,'vmes'); break
    tot_mes=tot_acum=None
    for r in range(28,40):
        rowtxt=" ".join(str(ws.cell(r,c).value) for c in range(1,blk+2) if ws.cell(r,c).value)
        if "TOTAL" in rowtxt.upper() and PV(r,'vmes'):
            tot_mes=PV(r,'vmes'); tot_acum=PV(r,'vacum'); break
    wb.close()
    return dict(prod=prod, comp=comp, tot_mes=tot_mes, tot_acum=tot_acum, N=N)

def ler_template(template):
    wb=openpyxl.load_workbook(template, data_only=True)
    base=wb["BASE_PARCEIRO"]; prod_emp={}; base_rows={}; prod_contrato={}; base_H={}
    for r in range(4,40):
        pid=base.cell(r,2).value
        if pid and re.match(r'^\d+\.\d+$',str(pid)):
            prod_emp[str(pid)]=str(base.cell(r,11).value).strip(); base_rows[str(pid)]=r
            prod_contrato[str(pid)]=base.cell(r,6).value
            base_H[str(pid)]=base.cell(r,8).value   # Qntd. Acum. do mês anterior (recorrência)
    comp_contrato={}
    for pid,emp in prod_emp.items():
        a=EMP_ALIAS.get(emp)
        if a and isinstance(prod_contrato.get(pid),(int,float)): comp_contrato[a]=comp_contrato.get(a,0.0)+prod_contrato[pid]
    rez=wb["RESUMO_EMPRESAS_M8"]; rez_rows={}
    for r in range(6,11):
        nm=str(rez.cell(r,2).value or "").upper()
        for key,a in [("DYNA","Dynatest"),("CG","Consultoria"),("CONSULTORIA","Consultoria"),("JURIDICO","Escritório Jurídico"),("CERT","Certificadora"),("CERTIFICADORA","Certificadora"),("TOTAL","TOTAL")]:
            if key in nm: rez_rows[r]=a; break
    # STATUS_MEDICOES: mapear medição -> coluna (pelo cabeçalho "MEDIÇÃO n"); achar linha do cabeçalho
    st=wb["STATUS_MEDICOES_1A8"]; status_hr=None; status_med_cols={}
    def _n(s):
        import unicodedata
        return "".join(c for c in unicodedata.normalize("NFD",str(s or "")) if unicodedata.category(c)!="Mn").upper()
    for r in range(1,st.max_row+1):
        for c in range(1,st.max_column+1):
            if re.match(r'MEDICAO\s*1\b',_n(st.cell(r,c).value)): status_hr=r; break
        if status_hr: break
    if status_hr:
        for c in range(1,st.max_column+1):
            m=re.match(r'MEDICAO\s*(\d+)',_n(st.cell(status_hr,c).value))
            if m and int(m.group(1)) not in status_med_cols: status_med_cols[int(m.group(1))]=c
    # MEDICAO_*: tabela detalhada "PLANILHA DE MEDIÇÃO" (produto na col D; SUBTOTAL GRUPO n; TOTAL GERAL) — DEC-063
    med_sheet=next((s for s in wb.sheetnames if s.upper().startswith("MEDICAO_")),None)
    med_rows={}; med_sub={}; med_tot=None; med_lbl=None
    if med_sheet:
        ms=wb[med_sheet]
        for r in range(1,ms.max_row+1):
            d=ms.cell(r,4).value; b=str(ms.cell(r,2).value or "").strip().upper()
            if d is not None and re.match(r'^\d+\.\d+$',str(d).strip()) and r>20: med_rows[str(d).strip()]=r
            m=re.match(r"SUBTOTAL GRUPO\s*(\d)",b)
            if m: med_sub[int(m.group(1))]=r
            if b.startswith("TOTAL GERAL"): med_tot=r
            if re.match(r"VALOR DA \d+.? MEDI",_n(b)): med_lbl=r
    wb.close()
    return dict(prod_emp=prod_emp, base_rows=base_rows, prod_contrato=prod_contrato, comp_contrato=comp_contrato,
                rez_rows=rez_rows, base_H=base_H, status_hr=status_hr, status_med_cols=status_med_cols,
                med_sheet=med_sheet, med_rows=med_rows, med_sub=med_sub, med_tot=med_tot, med_lbl=med_lbl)

def computar(dados, tmpl):
    prod=dados["prod"]; comp_acum={a:0.0 for a in EMP_ORDEM}
    for pid,emp in tmpl["prod_emp"].items():
        a=EMP_ALIAS.get(emp)
        if a and pid in prod and isinstance(prod[pid]["vacum"],(int,float)): comp_acum[a]+=prod[pid]["vacum"]
    cc=tmpl["comp_contrato"]; comp={}
    for a in EMP_ORDEM:
        mes=dados["comp"].get(a); acum=comp_acum.get(a,0.0); contr=cc.get(a)
        comp[a]=dict(mes=mes,acum=acum,contrato=contr,pct=(acum/contr if contr else None),saldo=(contr-acum if contr else None))
    tot_contr=sum(v for v in cc.values() if isinstance(v,(int,float)))
    tot=dict(mes=dados["tot_mes"],acum=dados["tot_acum"],contrato=tot_contr,
             pct=(dados["tot_acum"]/tot_contr if tot_contr else None),saldo=(tot_contr-dados["tot_acum"] if tot_contr else None))
    return comp, tot

def build_writes(dados, tmpl, comp, tot, N, seed_wb=None):
    """Retorna {sheet: {(row,col): value}} com todas as células a escrever."""
    P=cfg.params(); prod=dados["prod"]
    mesc,ano=cfg.competencia_medicao()   # DEC-063: competência de execução da medição
    try: status_med=P.get("status_medicao") or getattr(cfg.dados(),"STATUS_MEDICAO","Em análise")
    except Exception: status_med=P.get("status_medicao") or "Em análise"
    W={}
    def put(sheet,cell,val):
        from openpyxl.utils.cell import coordinate_to_tuple
        r,c=coordinate_to_tuple(cell); W.setdefault(sheet,{})[(r,c)]=val
    # ---- BASE_PARCEIRO ----
    put("BASE_PARCEIRO","B2",f"BASE TÉCNICA - PLANILHA DE MEDIÇÃO Parceiro ({N}ª MEDIÇÃO)")
    put("BASE_PARCEIRO","H3","Qntd. Acum.")
    for pid,r in tmpl["base_rows"].items():
        if pid not in prod: continue
        p=prod[pid]
        if p["pacum"] is not None: put("BASE_PARCEIRO",f"G{r}",p["pacum"])
        if p["vfinal"] is not None: put("BASE_PARCEIRO",f"I{r}",p["vfinal"])
        if p["vacum"]  is not None: put("BASE_PARCEIRO",f"J{r}",p["vacum"])
        # Qntd. Acum. (H): seed do oráculo (transição) ou recorrência template+incremento (DEC-034)
        if seed_wb is not None:
            hv=seed_wb["BASE_PARCEIRO"].cell(r,8).value
            if hv is not None: put("BASE_PARCEIRO",f"H{r}",hv)
        else:
            base_h=tmpl.get("base_H",{}).get(pid)
            if isinstance(base_h,(int,float)):
                q=p.get("quant"); nt=p.get("nota")
                inc=(q*(nt if nt is not None else 1.0)) if isinstance(q,(int,float)) else 0.0
                put("BASE_PARCEIRO",f"H{r}",base_h+inc)
    # ---- RESUMO_EMPRESAS_M8 ----
    for r,emp in tmpl["rez_rows"].items():
        if emp=="TOTAL":
            put("RESUMO_EMPRESAS_M8",f"C{r}",tot["mes"]); put("RESUMO_EMPRESAS_M8",f"D{r}",tot["acum"])
            put("RESUMO_EMPRESAS_M8",f"E{r}",tot["pct"]); put("RESUMO_EMPRESAS_M8",f"F{r}",tot["saldo"])
        else:
            e=comp[emp]
            put("RESUMO_EMPRESAS_M8",f"C{r}",e["mes"]); put("RESUMO_EMPRESAS_M8",f"D{r}",e["acum"])
            put("RESUMO_EMPRESAS_M8",f"E{r}",e["pct"]); put("RESUMO_EMPRESAS_M8",f"F{r}",e["saldo"])
    # ---- PAINEL_FINANCEIRO_M8 ----
    # DEC-063: a medição é identificada pela competência de EXECUÇÃO (mês anterior ao do relatório)
    mesc,ano=cfg.competencia_medicao()
    put("PAINEL_FINANCEIRO_M8","I4",f"Referência: {N}ª Medição – {mesc}/{ano}")
    put("PAINEL_FINANCEIRO_M8","F7",f"Valor da\n{N}ª Medição")
    put("PAINEL_FINANCEIRO_M8","D8",tot["acum"]); put("PAINEL_FINANCEIRO_M8","F8",tot["mes"])
    put("PAINEL_FINANCEIRO_M8","H8",tot["saldo"]); put("PAINEL_FINANCEIRO_M8","J8",tot["pct"])
    put("PAINEL_FINANCEIRO_M8","N18",tot["saldo"])
    rows_pc={0:54,1:55,2:56,3:57}   # DYNA,CG,JURIDICO,CERT
    for i,emp in enumerate(EMP_ORDEM):
        r=rows_pc[i]; e=comp[emp]
        put("PAINEL_FINANCEIRO_M8",f"D{r}",e["mes"]); put("PAINEL_FINANCEIRO_M8",f"F{r}",e["acum"])
        put("PAINEL_FINANCEIRO_M8",f"J{r}",e["pct"]); put("PAINEL_FINANCEIRO_M8",f"L{r}",e["saldo"])
        put("PAINEL_FINANCEIRO_M8",f"N{r}",status_med)
    put("PAINEL_FINANCEIRO_M8","D58",tot["mes"]); put("PAINEL_FINANCEIRO_M8","F58",tot["acum"])
    put("PAINEL_FINANCEIRO_M8","J58",tot["pct"]); put("PAINEL_FINANCEIRO_M8","L58",tot["saldo"])  # % do TOTAL era estático (39,07%) até o RM15
    # gráfico l.84-87: D=acum,E=saldo,F=%acum,G=%saldo (C=contrato estático)
    for i,emp in enumerate(EMP_ORDEM):
        r=84+i; e=comp[emp]
        put("PAINEL_FINANCEIRO_M8",f"D{r}",e["acum"]); put("PAINEL_FINANCEIRO_M8",f"E{r}",e["saldo"])
        put("PAINEL_FINANCEIRO_M8",f"F{r}",e["pct"]); put("PAINEL_FINANCEIRO_M8",f"G{r}",(1-e["pct"]) if e["pct"] is not None else None)
    # ---- MAPEAMENTO_EMPRESAS (l.5-8: E=mes,F=acum,H=%) ----
    for i,emp in enumerate(EMP_ORDEM):
        r=5+i; e=comp[emp]
        put("MAPEAMENTO_EMPRESAS",f"E{r}",e["mes"]); put("MAPEAMENTO_EMPRESAS",f"F{r}",e["acum"]); put("MAPEAMENTO_EMPRESAS",f"H{r}",e["pct"])
    # ---- MEDICAO_08_PROG (textos + totais + per empresa) ----
    put("MEDICAO_08_PROG","B3",f"{N}ª MEDIÇÃO"); put("MEDICAO_08_PROG","B27",f"PLANILHA DE MEDIÇÃO – {N}ª MEDIÇÃO")
    put("MEDICAO_08_PROG","H9",tot["acum"]); put("MEDICAO_08_PROG","H10",tot["mes"])
    put("MEDICAO_08_PROG","H11",tot["saldo"]); put("MEDICAO_08_PROG","H12",tot["pct"])
    for i,emp in enumerate(EMP_ORDEM):
        put("MEDICAO_08_PROG",f"G{15+i}",comp[emp]["mes"]); put("MEDICAO_08_PROG",f"G{21+i}",comp[emp]["acum"])
    # ---- MEDICAO_*: tabela detalhada inteira da medição N (DEC-063 — determinação (I) do Parceiro) ----
    # Até o RM15 só os totais do topo eram atualizados; a tabela por produto e o rótulo "Valor da Nª medição"
    # ficavam com os dados da 12ª/11ª medição.
    ms=tmpl.get("med_sheet") or "MEDICAO_08_PROG"
    if tmpl.get("med_lbl"): put(ms,f"B{tmpl['med_lbl']}",f"Valor da {N}ª medição:")
    TR="—"
    def v0(x): return x if isinstance(x,(int,float)) else 0
    sub={}
    for pid,r in tmpl.get("med_rows",{}).items():
        p=prod.get(pid)
        if not p: continue
        q=p.get("quant"); nt=p.get("nota")
        medido=isinstance(q,(int,float)) and q>0
        put(ms,f"G{r}",p["quantc"] if p.get("quantc") else TR)
        put(ms,f"H{r}",p["vprod"] if p.get("vprod") else TR)
        put(ms,f"I{r}",p.get("contrato"))
        put(ms,f"J{r}",v0(p.get("pacum")))
        put(ms,f"K{r}",q if medido else TR)
        put(ms,f"L{r}",nt if (medido and isinstance(nt,(int,float))) else TR)
        put(ms,f"M{r}",v0(p.get("vfinal"))); put(ms,f"N{r}",v0(p.get("vmes"))); put(ms,f"O{r}",v0(p.get("vacum")))
        g=int(pid.split(".")[0]); s=sub.setdefault(g,[0.0,0.0,0.0,0.0])
        s[0]+=v0(p.get("contrato")); s[1]+=v0(p.get("vfinal")); s[2]+=v0(p.get("vmes")); s[3]+=v0(p.get("vacum"))
    for g,r in tmpl.get("med_sub",{}).items():
        if g in sub:
            put(ms,f"I{r}",sub[g][0]); put(ms,f"M{r}",sub[g][1]); put(ms,f"N{r}",sub[g][2]); put(ms,f"O{r}",sub[g][3])
    if tmpl.get("med_tot") and sub:
        r=tmpl["med_tot"]; ti=sum(s[0] for s in sub.values()); to=sum(s[3] for s in sub.values())
        put(ms,f"I{r}",ti); put(ms,f"J{r}",(to/ti) if ti else None)
        put(ms,f"M{r}",sum(s[1] for s in sub.values())); put(ms,f"N{r}",sum(s[2] for s in sub.values())); put(ms,f"O{r}",to)
    # ---- TEXTOS de título com o nº da medição ----
    put("PAINEL_FINANCEIRO_M8","B2",f"QUADRO FINANCEIRO CONSOLIDADO — {N}ª MEDIÇÃO")
    put("PAINEL_FINANCEIRO_M8","B52",f"RESUMO FINANCEIRO POR EMPRESA — {N}ª MEDIÇÃO")
    put("RESUMO_EMPRESAS_M8","B2",f"RESUMO FINANCEIRO POR EMPRESA — {N}ª MEDIÇÃO")
    put("RESUMO_EMPRESAS_M8","B3",f"Contrato n.º 000/0000  |  Programa  |  Referência: {mesc}/{ano}")
    put("STATUS_MEDICOES_1A8","A3",f"Contrato n.º 000/0000  |  Programa  |  Medições 1 a {N}")
    put("MEDICAO_08_PROG","B7",f"QUADRO FINANCEIRO CONSOLIDADO — ATÉ A {N}ª MEDIÇÃO")
    put("MEDICAO_08_PROG","B14",f"DISTRIBUIÇÃO POR EMPRESA — {N}ª MEDIÇÃO")
    put("MEDICAO_08_PROG","B27",f"PLANILHA DE MEDIÇÃO — {N}ª MEDIÇÃO")
    # ---- STATUS_MEDICOES_1A8: coluna da medição N (valor mês por empresa + status) ----
    # linhas: hr=cabeçalho; hr+1=status; hr+2..hr+5=DYNA/CG/JURIDICO/CERT; hr+6=TOTAL.
    status_fmt=None  # (src_col, dst_col) para copiar formato no aplicar_com
    try: MSTAT=getattr(cfg.dados(),"MEDICOES_STATUS",{})
    except Exception: MSTAT={}
    if seed_wb is not None:
        # transição: coluna N já existe no oráculo; preencher lá
        sto=seed_wb["STATUS_MEDICOES_1A8"]; col_n=None
        for c in range(2,sto.max_column+1):
            h=sto.cell(7,c).value
            if h and re.search(rf'\b{N}\b',str(h)): col_n=c
        if col_n:
            L=get_column_letter(col_n)
            for i,emp in enumerate(EMP_ORDEM): put("STATUS_MEDICOES_1A8",f"{L}{9+i}",comp[emp]["mes"])
            put("STATUS_MEDICOES_1A8",f"{L}13",tot["mes"]); put("STATUS_MEDICOES_1A8",f"{L}8",status_med)
            put("STATUS_MEDICOES_1A8",f"{L}7",f"MEDIÇÃO {N}")
    else:
        # produção: template tem medições até N-1; ADICIONAR a coluna da medição N (nova coluna)
        hr=tmpl.get("status_hr") or 7; mc=tmpl.get("status_med_cols",{})
        prev=mc.get(N-1)
        if prev:
            new_col=prev+1; L=get_column_letter(new_col); status_fmt=(prev,new_col)
            put("STATUS_MEDICOES_1A8",f"{L}{hr}",f"MEDIÇÃO {N}")
            put("STATUS_MEDICOES_1A8",f"{L}{hr+1}",MSTAT.get(N,status_med))
            for i,emp in enumerate(EMP_ORDEM): put("STATUS_MEDICOES_1A8",f"{L}{hr+2+i}",comp[emp]["mes"])
            put("STATUS_MEDICOES_1A8",f"{L}{hr+6}",tot["mes"])
            # atualizar o status das medições anteriores conforme MEDICOES_STATUS (ex.: N-1 -> Pago)
            for k,col in mc.items():
                s=MSTAT.get(k)
                if s: put("STATUS_MEDICOES_1A8",f"{get_column_letter(col)}{hr+1}",s)
    # ---- SEED (mês de transição): copiar do oráculo as abas aux não-deriváveis ----
    if seed_wb is not None:
        for shname in ("STATUS_MEDICOES_1A8","MEDICAO_08_PROG"):
            so=seed_wb[shname]
            for r in range(1,so.max_row+1):
                for c in range(1,so.max_column+1):
                    v=so.cell(r,c).value
                    if v is not None: W.setdefault(shname,{})[(r,c)]=v
    return W, status_med, status_fmt

# ---------- aplicar via COM (preserva gráfico) ----------
def aplicar_com(template, writes, dest, status_fmt=None, status_med=None):
    import win32com.client as win32, atexit
    tmp=os.path.join(cfg.work_dir(),"02_TRABALHO","painel_build.xlsx"); shutil.copy2(template,tmp)
    xl=win32.gencache.EnsureDispatch("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
    atexit.register(lambda: (xl.Quit() if xl else None))
    wb=xl.Workbooks.Open(tmp,False,False)
    shname={ws.Name:ws for ws in wb.Worksheets}
    # cor das células de Status do RESUMO (N54-57) conforme o status do mês; legenda "Não faturado" na mesma cor
    cor=STATUS_COR.get(status_med)
    if cor:
        pf=shname.get("PAINEL_FINANCEIRO_M8")
        if pf:
            try:
                for r in (54,55,56,57): pf.Cells(r,14).Interior.Color=_com_color(cor)   # N54..N57
                if status_med=="Não faturado": pf.Cells(50,11).Interior.Color=_com_color(cor)  # legenda K50
            except Exception as e: cfg.log(f"  aviso: falha ao colorir status: {e}")
    # copiar formato da coluna da medição anterior -> nova coluna do STATUS (antes de escrever os valores)
    if status_fmt:
        src,dst=status_fmt; st=shname.get("STATUS_MEDICOES_1A8")
        if st:
            try:
                st.Columns(src).Copy(); st.Columns(dst).PasteSpecial(-4122)  # xlPasteFormats
                xl.CutCopyMode=False
            except Exception as e: cfg.log(f"  aviso: falha ao copiar formato STATUS col {src}->{dst}: {e}")
    for sheet,cells in writes.items():
        if sheet not in shname: cfg.log(f"  (aba ausente no template: {sheet})"); continue
        ws=shname[sheet]
        for (r,c),v in cells.items():
            try: ws.Cells(r,c).Value=v
            except Exception as e: cfg.log(f"  falha {sheet}!{get_column_letter(c)}{r}: {e}")
    wb.Save(); wb.Close(True); xl.Quit()
    shutil.copy2(tmp,dest); cfg.log(f"  painel salvo: {dest}")

def aplicar_openpyxl(template, writes, dest):
    wb=openpyxl.load_workbook(template)
    for sheet,cells in writes.items():
        if sheet not in wb.sheetnames: continue
        ws=wb[sheet]
        for (r,c),v in cells.items(): ws.cell(r,c).value=v
    wb.save(dest); cfg.log(f"  (openpyxl, sem gráfico) painel salvo: {dest}")

def validar(gerado, oraculo):
    wg=openpyxl.load_workbook(gerado, data_only=True); wo=openpyxl.load_workbook(oraculo, data_only=True)
    def close(a,b):
        if isinstance(a,(int,float)) and isinstance(b,(int,float)): return abs(a-b)<=0.5
        return (a or "")==(b or "") or str(a)==str(b)
    diffs=[]
    for sh in wo.sheetnames:
        if sh not in wg.sheetnames: continue
        A=wg[sh]; B=wo[sh]
        for r in range(1,max(A.max_row,B.max_row)+1):
            for c in range(1,max(A.max_column,B.max_column)+1):
                a=A.cell(r,c).value; b=B.cell(r,c).value
                if a is None and b is None: continue
                if not close(a,b): diffs.append((sh,f"{get_column_letter(c)}{r}",a,b))
    wg.close(); wo.close(); return diffs

def main():
    N=medicao_n(); mes=cfg.mes_min(); WK=cfg.work_dir()
    planilha=planilha_medicao_path(); template=painel_template_path()
    if not planilha or not os.path.exists(planilha): cfg.log(f"ERRO: Planilha de Medição não encontrada (N={N})."); sys.exit(1)
    if not template or not os.path.exists(template): cfg.log("ERRO: painel-template do mês anterior não encontrado."); sys.exit(1)
    cfg.log(f"painel: medição {N} | Planilha={os.path.basename(planilha)} | template={os.path.basename(template)}")
    dados=extrair_planilha(planilha,N); tmpl=ler_template(template); comp,tot=computar(dados,tmpl)
    cfg.log(f"  totais: mês={tot['mes']:.2f} acum={tot['acum']:.2f} saldo={tot['saldo']:.2f} %={tot['pct']*100:.2f}")
    seed_wb=None
    if "--seed-qntd" in sys.argv:
        seed_wb=openpyxl.load_workbook(sys.argv[sys.argv.index("--seed-qntd")+1], data_only=True)
    writes,status_med,status_fmt=build_writes(dados,tmpl,comp,tot,N,seed_wb=seed_wb)
    ncells=sum(len(v) for v in writes.values())
    cfg.log(f"  {ncells} células a preencher em {len(writes)} abas (status medição='{status_med}')")
    if status_fmt: cfg.log(f"  STATUS: nova coluna da medição {N} (col {get_column_letter(status_fmt[1])}, formato copiado da {get_column_letter(status_fmt[0])})")
    if "--build" in sys.argv:
        dest=os.path.join(WK,"07_SAIDA",f"P1.1 - Quadro Financeiro do mês de {mes}.xlsx")
        os.makedirs(os.path.dirname(dest),exist_ok=True)
        if "--openpyxl" in sys.argv: aplicar_openpyxl(template,writes,dest)
        else: aplicar_com(template,writes,dest,status_fmt=status_fmt,status_med=status_med)
        if "--validate" in sys.argv:
            diffs=validar(dest, sys.argv[sys.argv.index("--validate")+1])
            imaged=[d for d in diffs if d[0] in ("PAINEL_FINANCEIRO_M8","RESUMO_EMPRESAS_M8","BASE_PARCEIRO","STATUS_MEDICOES_1A8")]
            cfg.log(f"  VALIDAÇÃO vs oráculo: {len(diffs)} diffs totais; {len(imaged)} nas abas que viram imagem")
            for d in diffs[:40]: cfg.log(f"    DIFF {d[0]}!{d[1]}: got={d[2]} exp={d[3]}")

if __name__=="__main__":
    main()

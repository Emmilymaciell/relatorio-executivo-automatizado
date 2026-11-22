# -*- coding: utf-8 -*-
"""
16_reunioes_do_mestre.py — PRODUZ o anexo mensal 'P1.2 - Controle de Reuniões do mês de {mês}.xlsx'
FILTRANDO o mês do MESTRE '…\\Desenvolvimento de Produtos\\P1.2 - Controle de Reuniões.xlsx'
(abas "Reuniões de GOV 1-1" e "Reuniões Stakehold 1-2" contêm todos os meses). Ver DEC-036.
NÃO usa o espelho Parceiro nem um mensal já pronto. O `04_print_reunioes.py` lê o anexo produzido aqui.

Preserva a estrutura (linhas 1-3 de cabeçalho + colunas) e copia só as linhas cujo mês da 1ª coluna
de data casa com a competência. Se o mês não existir no mestre, PARA com aviso (peça ao usuário para
atualizar o mestre — DEC-037).

Saída: copiada ao anexo do mês (P1.2) fora do DRY_RUN; sempre gerada no work.
Env: PROG_REUNIOES_MESTRE (força o caminho do mestre).
"""
import os, sys, shutil, datetime, re
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
import openpyxl

def mestre_path():
    env=os.environ.get("PROG_REUNIOES_MESTRE")
    if env and os.path.exists(env): return env
    dp=cfg.dir_desenv_produtos()
    return cfg.find_in_dir(dp,"controle de reuni", ext=".xlsx", exclude=("do mês","do mes"))

def _date_of(row_cells):
    for v in row_cells[:3]:
        if isinstance(v,datetime.datetime): return v
    return None

def main():
    P=cfg.params(); mes=cfg.mes_min(); mesc=P["mes"]; ano=cfg.ano(); mnum=int(P["mes_numero"]); WK=cfg.work_dir()
    src=mestre_path()
    if not src or not os.path.exists(src):
        cfg.log("ERRO: mestre de reuniões não encontrado."); sys.exit(1)
    cfg.log(f"reuniões: filtrando {mesc}/{ano} do mestre {os.path.basename(src)}")
    wb=openpyxl.load_workbook(src)  # com formatação (não data_only, para preservar estilos ao copiar)
    wbv=openpyxl.load_workbook(src, data_only=True)  # valores p/ ler datas
    total_mes=0
    for sheet in wb.sheetnames:
        ws=wb[sheet]; wsv=wbv[sheet]
        # achar linha de cabeçalho (contém "Data")
        hdr=None
        for r in range(1,8):
            rowtxt=" ".join(str(wsv.cell(r,c).value or "") for c in range(1,5)).lower()
            if "data" in rowtxt and ("participante" in rowtxt or "assunto" in rowtxt): hdr=r; break
        if not hdr: hdr=3
        # linhas do mês
        keep=[]
        for r in range(hdr+1, wsv.max_row+1):
            vals=[wsv.cell(r,c).value for c in range(1,wsv.max_column+1)]
            d=_date_of(vals)
            if d and d.year==ano and d.month==mnum: keep.append(r)
        total_mes+=len(keep)
        cfg.log(f"  {sheet}: {len(keep)} reunião(ões) de {mesc}/{ano}")
        # apagar da cópia (ws) todas as linhas de dados que NÃO são do mês (de baixo p/ cima)
        for r in range(ws.max_row, hdr, -1):
            if r not in keep:
                ws.delete_rows(r,1)
    if total_mes==0:
        cfg.log(f"AVISO: o mestre NÃO tem reuniões de {mesc}/{ano}. Atualize o mestre "
                f"'P1.2 - Controle de Reuniões.xlsx' com as reuniões do mês (ou informe ao operador). "
                f"NÃO usar cópia do espelho Parceiro (DEC-037).")
        cfg.set_qa_warning("reunioes_mestre_sem_mes",
                           f"Mestre de reuniões sem {mesc}/{ano} — anexo mensal não pôde ser produzido do mestre.")
    work_out=os.path.join(WK,"02_TRABALHO",f"P1.2 - Controle de Reuniões do mês de {mes}.xlsx")
    wb.save(work_out); cfg.log(f"  anexo mensal (work): {work_out}")
    if not cfg.dry_run() and total_mes>0:
        dst=os.path.join(cfg.anexo_p12_dir(), cfg.nome_anexo_g1("P1.2", f"Controle de Reuniões do mês de {mes}", ".xlsx"))
        # versionar se já existir
        if os.path.exists(dst):
            b,e=os.path.splitext(dst); i=2
            while os.path.exists(f"{b}_v{i}{e}"): i+=1
            dst=f"{b}_v{i}{e}"
        shutil.copy2(work_out,dst); cfg.log(f"  anexo -> {dst}")

if __name__=="__main__":
    main()

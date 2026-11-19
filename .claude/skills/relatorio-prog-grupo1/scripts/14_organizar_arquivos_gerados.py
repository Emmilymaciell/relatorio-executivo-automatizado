# -*- coding: utf-8 -*-
"""
14_organizar_arquivos_gerados.py — Organiza cópias dos arquivos GERADOS (corpo + anexos) numa pasta
'Arquivos Gerados' para controle do usuário. NÃO modifica originais (apenas copia).
- Destino padrão: SharePoint '…/Desenvolvimento de Produtos/Obsoleto/Programa RM{N} {Mês} {ano} - Arquivos Gerados'.
- DRY_RUN (PROG_DRY_RUN=1): destino local em output/RM.../07_SAIDA/Arquivos_Gerados (não escreve no SharePoint).
Nomes/mês/nº derivam de config; anexos são localizados por padrão (find_in_dir).
"""
import os, sys, shutil, glob
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

def main():
    P=cfg.params(); N=cfg.numero(); mes=cfg.mes_min(); mes_cap=P["mes"]; ano=cfg.ano()
    WK=cfg.work_dir(); TR=os.path.join(WK,"02_TRABALHO"); OUT=os.path.join(WK,"07_SAIDA"); IMG=os.path.join(WK,"03_IMAGENS")

    folder_name=f"Programa RM{N} {mes_cap} {ano} - Arquivos Gerados"
    if cfg.dry_run():
        DEST=os.path.join(OUT,"Arquivos_Gerados"); cfg.log(f"DRY_RUN: destino LOCAL {DEST} (sem escrever no SharePoint)")
    else:
        dp=cfg.dir_desenv_produtos(); obs=cfg._find_sub(dp,"obsoleto") or os.path.join(dp,"Obsoleto")
        os.makedirs(obs,exist_ok=True); DEST=os.path.join(obs,folder_name)
    CORPO=os.path.join(DEST,"Corpo"); ANEXOS=os.path.join(DEST,"Anexos"); CORPO_IMG=os.path.join(CORPO,"Imagens usadas no corpo")
    for d in [DEST,CORPO,ANEXOS,CORPO_IMG]: os.makedirs(d, exist_ok=True)

    def cp(src,dstdir,newname=None):
        if not src or not os.path.exists(src): cfg.log(f"  (ausente) {os.path.basename(src) if src else src}"); return
        dst=os.path.join(dstdir, newname or os.path.basename(src))
        # nomes do POP + pasta Obsoleto passam de 260 caracteres: usar caminho longo do Windows (\\?\)
        lp=lambda p: p if p.startswith("\\\\?\\") or os.name!="nt" else "\\\\?\\"+os.path.abspath(p)
        shutil.copy2(lp(src),lp(dst))
        cfg.log(f"  -> {os.path.relpath(dst,DEST)}")

    cfg.log("== CORPO (fontes editáveis + relatório final) ==")
    base=cfg.doc_final_basename()
    # relatório final: pega a versão mais recente (base + _v*) em 07_SAIDA
    for ext in (".docx",".pdf"):
        cands=sorted(glob.glob(os.path.join(OUT, base+"*"+ext)))
        if cands: cp(cands[-1],CORPO, base+ext)
        else: cfg.log(f"  (ausente) relatório final {ext}")
    tl=[p for p in (os.path.join(TR,f"Timeline_ate_{mes}{ano}.pptx"), os.path.join(TR,f"Timeline_{mes_cap}{ano}.pptx"))
        if os.path.exists(p)]   # 07 (legado) ou 17 (DEC-035) grava o editável com nomes diferentes
    cp(tl[0] if tl else None,CORPO,f"P1.1 - Timeline (editável) - {mes} {ano}.pptx")
    cp(os.path.join(TR,"controle_print.xlsx"),CORPO,"Controle de Produtos (layout do print no corpo).xlsx")
    cp(os.path.join(TR,"reunioes_print.xlsx"),CORPO,"Reuniões (layout do print no corpo).xlsx")
    cp(os.path.join(TR,"linha_medicoes.xlsx"),CORPO,"Linha do Tempo de Medições (quadros).xlsx")
    cp(os.path.join(TR,"Quadro_Financeiro_WORK.xlsx"),CORPO,"Quadro Financeiro (copia de trabalho para imagens).xlsx")
    for fol in ["CONTROLE_PRODUTOS","DASHBOARD","TIMELINE","REUNIOES_P11","REUNIOES_P12","FINANCEIRO"]:
        sub=os.path.join(CORPO_IMG,fol); os.makedirs(sub,exist_ok=True)
        for f in glob.glob(os.path.join(IMG,fol,"*.png")): shutil.copy2(f,sub)
    cfg.log("  -> Imagens usadas no corpo\\ (todas as PNG)")

    cfg.log("== ANEXOS (cópias dos arquivos oficiais dos anexos) ==")
    p11=cfg.anexo_p11_dir(); p12=cfg.anexo_p12_dir()
    anexos_src=[
        cfg.find_in_dir(p11,"p1.1","controle de entrega", ext=".xlsx"),
        cfg.find_in_dir(p11,"p1.1","dashboard", ext=".pdf"),
        cfg.find_in_dir(p11,"p1.1","quadro financeiro", ext=".xlsx"),
        cfg.find_in_dir(p11,"p1.1","timeline", ext=".pdf"),
        cfg.find_in_dir(p12,"p1.2","controle de reuni", ext=".xlsx"),
    ]
    for src in anexos_src: cp(src,ANEXOS)

    open(os.path.join(DEST,"LEIA-ME.txt"),"w",encoding="utf-8").write(
f"""ARQUIVOS GERADOS PARA O RELATÓRIO MENSAL N° {N} (Grupo 1 / {mes_cap} {ano})

CORPO\\  -> arquivos editáveis usados para montar o CORPO do relatório:
  - Relatório final (DOCX editável + PDF)
  - Timeline (PPTX editável) - altere aqui a linha do tempo
  - Excels de layout dos prints (Controle de Produtos, Reuniões, Linha do Tempo de Medições, Quadro Financeiro)
  - Imagens usadas no corpo\\ -> todas as imagens (PNG) inseridas no relatório

ANEXOS\\ -> cópias dos arquivos oficiais dos anexos (originais permanecem na pasta de Anexos da competência):
  - P1.1 Controle de Entrega de Produtos, Dashboard, Quadro Financeiro, Timeline
  - P1.2 Controle de Reuniões

Obs.: para trocar uma imagem do corpo, edite o arquivo editável correspondente (PPTX/XLSX),
gere o print e substitua a imagem no DOCX.
""")
    cfg.log(f"LEIA-ME.txt criado. PASTA: {DEST}")

if __name__=="__main__":
    main()

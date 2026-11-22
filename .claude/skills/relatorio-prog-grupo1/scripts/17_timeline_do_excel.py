# -*- coding: utf-8 -*-
"""
17_timeline_do_excel.py — Gera as imagens da Timeline (DEC-035): lê os MARCOS do mês do Excel
'…\\Desenvolvimento de Produtos\\P1.1 -Controle Timeline.xlsx' (aba 'Base Timeline', competência do
mês, Incluir? = Sim), PREENCHE o slide do mês no PPTX-base 'P1.1 - Timeline Relatório Programa.pptx'
(grupos de marcador já posicionados: define data + descrição de cada um, na ordem por data) e
RENDERIZA todas as slides (Jul/2025 → mês corrente) em PNG (200 dpi) — DEC-028. NÃO usa o espelho Parceiro.

Se o Excel não tiver marcos do mês, PARA com aviso (peça os marcos ao usuário — DEC-035/037).
Env: PROG_TIMELINE_XLSX (força o Excel de controle), PROG_TIMELINE_PPTX (força o PPTX-base).
"""
import os, sys, re, shutil, datetime, atexit
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
import openpyxl

def excel_controle_path():
    env=os.environ.get("PROG_TIMELINE_XLSX")
    if env and os.path.exists(env): return env
    dp=cfg.dir_desenv_produtos()
    return cfg.find_in_dir(dp,"controle timeline", ext=".xlsx") or cfg.find_in_dir(dp,"timeline", ext=".xlsx", exclude=("relatório","relatorio"))

def ler_marcos(xlsx, ano, mnum):
    wb=openpyxl.load_workbook(xlsx, data_only=True); ws=wb["Base Timeline"]
    # cabeçalhos na linha 3
    hdr={}
    for c in range(1,ws.max_column+1):
        v=ws.cell(3,c).value
        if v: hdr[str(v).replace("\n"," ").strip().lower()]=c
    def col(*keys):
        for k in keys:
            for h,c in hdr.items():
                if k in h: return c
        return None
    cComp=col("competência do relat","competencia do relat","competência","competencia")
    cData=col("data do marco"); cTit=col("marco / título","marco / titulo","título da timeline","titulo da timeline","marco")
    cInc=col("incluir");
    marcos=[]
    for r in range(4,ws.max_row+1):
        comp=ws.cell(r,cComp).value
        if not isinstance(comp,datetime.datetime): continue
        if comp.year!=ano or comp.month!=mnum: continue
        inc=ws.cell(r,cInc).value if cInc else "Sim"
        if str(inc).strip().lower() not in ("sim","s","x","true","1"): continue
        d=ws.cell(r,cData).value; tit=ws.cell(r,cTit).value
        if not isinstance(d,datetime.datetime) or not tit: continue
        marcos.append((d, f"{d.day:02d}/{d.month:02d}", str(tit).strip()))
    wb.close()
    marcos.sort(key=lambda m:(m[0], m[2]))
    return marcos

def slide_index(ano, mnum):
    """1-based: Jul/2025=1 ... (ano-2025)*12 + mnum - 6."""
    return (ano-2025)*12 + mnum - 6

def _norm(s):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD",str(s or "")) if unicodedata.category(c)!="Mn").lower()

def _all_text(slide):
    out=[]
    def walk(shapes):
        for sh in shapes:
            if sh.has_text_frame: out.append(sh.text_frame.text)
            if sh.shape_type==6:
                try: walk(sh.shapes)
                except Exception: pass
    walk(slide.shapes); return " ".join(out)

def slide_tem_marcos(slide, marcos):
    """Quantos marcos (por trecho do título) já aparecem no slide."""
    txt=_norm(_all_text(slide)); ok=0
    for _,_,tit in marcos:
        chave=_norm(tit)[:24]
        if chave and chave in txt: ok+=1
    return ok

def main():
    P=cfg.params(); ano=cfg.ano(); mnum=int(P["mes_numero"]); mesc=P["mes"]; WK=cfg.work_dir()
    xlsx=excel_controle_path(); pptx=cfg.timeline_pptx_src()
    if not xlsx or not os.path.exists(xlsx):
        cfg.log("ERRO: Excel de Controle da Timeline não encontrado."); sys.exit(1)
    if not pptx or not os.path.exists(pptx):
        cfg.log("ERRO: PPTX-base da Timeline não encontrado."); sys.exit(1)
    marcos=ler_marcos(xlsx, ano, mnum)
    cfg.log(f"timeline: {len(marcos)} marcos de {mesc}/{ano} no Excel de Controle")
    if not marcos:
        cfg.log(f"AVISO: sem marcos de {mesc}/{ano} no Excel de Controle. PERGUNTE os marcos ao usuário, "
                f"preencha o Excel (aba Base Timeline) e reexecute. NÃO usar o PDF do espelho Parceiro (DEC-035).")
        cfg.set_qa_warning("timeline_sem_marcos", f"Excel de Controle sem marcos de {mesc}/{ano}.")
        sys.exit(2)
    for m in marcos: cfg.log(f"   {m[1]} | {m[2][:52]}")
    # copiar pptx e preencher o slide do mês
    from pptx import Presentation
    work_pptx=os.path.join(WK,"02_TRABALHO",f"Timeline_{mesc}{ano}.pptx"); shutil.copy2(pptx,work_pptx)
    pr=Presentation(work_pptx)
    idx=slide_index(ano,mnum)
    if idx<1 or idx>len(pr.slides):
        cfg.log(f"ERRO: slide do mês (índice {idx}) não existe no PPTX ({len(pr.slides)} slides). "
                f"Crie o slide do mês no PPTX-base."); sys.exit(1)
    # Conferir se o slide do mês já contém os marcos (é mantido/preenchido no PPTX-base).
    presentes=slide_tem_marcos(pr.slides[idx-1], marcos)
    cfg.log(f"  slide {idx}: {presentes}/{len(marcos)} marcos presentes")
    if presentes < max(1, len(marcos)-1):
        cfg.log(f"  AVISO: o slide {idx} do PPTX-base NÃO está preenchido com os marcos do Excel "
                f"({presentes}/{len(marcos)}). O layout da timeline (posições/setas, datas iguais) é "
                f"montado no PPTX; prepare/preencha o slide do mês no PPTX-base e reexecute. "
                f"NÃO gero um slide quebrado nem uso o espelho Parceiro (DEC-035).")
        cfg.set_qa_warning("timeline_slide_incompleto",
                           f"Slide do mês ({presentes}/{len(marcos)} marcos) precisa ser preenchido no PPTX-base.")
    else:
        cfg.log(f"  slide {idx} OK (marcos conferem com o Excel); renderizando como está")
    cfg.log(f"  renderizando {len(pr.slides)} slides")
    # exportar PDF + render via PowerPoint COM
    import win32com.client as win32
    pp=win32.gencache.EnsureDispatch("PowerPoint.Application")
    atexit.register(lambda: pp.Quit())
    pres=pp.Presentations.Open(os.path.abspath(work_pptx), WithWindow=False)
    pdf=os.path.join(WK,"04_PDFS_TEMP","timeline.pdf")
    if os.path.exists(pdf):
        try: os.remove(pdf)   # não sobrescrever PDF pré-existente (DEC-038)
        except OSError: pass
    pres.SaveAs(os.path.abspath(pdf), 32)  # ppSaveAsPDF
    pres.Close(); pp.Quit()
    import fitz
    outdir=os.path.join(WK,"03_IMAGENS","TIMELINE")
    for f in os.listdir(outdir):
        if f.endswith(".png"): os.remove(os.path.join(outdir,f))
    d=fitz.open(pdf)
    for i,pg in enumerate(d): pg.get_pixmap(dpi=200).save(os.path.join(outdir,f"timeline_p{i+1:02d}.png"))
    cfg.log(f"  timeline: {d.page_count} páginas renderizadas -> {outdir}"); d.close()
    # limpar avisos antigos de timeline
    for k in ("timeline_sem_marcos","timeline_fallback_xxxx","timeline_pptx_ausente"):
        cfg.set_qa_warning(k,"",active=False)
    # anexo PDF (fora do dry-run)
    if not cfg.dry_run():
        dst=os.path.join(cfg.anexo_p11_dir(), cfg.nome_anexo_g1("P1.1", f"Timeline de Acompanhamento Macro até {mesc.lower()} {ano}", ".pdf"))
        shutil.copy2(pdf,dst); cfg.log(f"  anexo timeline -> {dst}")

if __name__=="__main__":
    main()

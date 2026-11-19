# -*- coding: utf-8 -*-
"""
07_timeline.py — Monta a Timeline macro do mês no PPT de acompanhamento (SharePoint, cópia local),
duplicando os grupos-modelo REAIS (círculo/seta/caixas) para cada marco do mês (DEC-013: seta curta
que não cruza o texto; descrição ancorada no topo). Exporta PDF e renderiza TODAS as slides
(timeline acumulada, do 1º mês até o mês corrente) -> PNGs para o corpo do relatório (DEC-028).
- MARCOS vêm de cfg.dados() (config/dados_mes.py). Mês/ano e nº de marcos são derivados (nada fixo).
- Slide-alvo e slide-modelo são DESCOBERTOS por conteúdo (placeholder 'XXXX' / marco com data-modelo),
  sem números de slide fixos. Só escreve arquivos locais.
"""
import os, sys, shutil, atexit, re
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

LIGHT=0+191*256+99*65536   # 00BF63
DARK=28+85*256+96*65536    # 1C5560

def deep(grp):
    for it in grp.GroupItems:
        yield it
        if it.Type==6: yield from deep(it)
def find_circle(grp):
    for it in deep(grp):
        if it.Type==5: return it   # freeform
    return None
def find_arrow(grp):
    for it in deep(grp):
        if it.Type==1: return it   # autoshape
    return None
def find_date_desc(grp):
    date=desc=None
    for it in deep(grp):
        if it.HasTextFrame and it.TextFrame.HasText:
            t=it.TextFrame.TextRange.Text.strip()
            if re.match(r'^\d\d/\d\d$', t) or t=="DD/MM": date=it
            elif len(t)>=4 and t not in ("/","|"):
                if desc is None or len(t)>len(desc.TextFrame.TextRange.Text): desc=it
    return date,desc
def circle_offset(grp):
    c=find_circle(grp); return (c.Left + c.Width/2) - grp.Left

def main():
    P=cfg.params(); D=cfg.dados()
    MARCOS=D.MARCOS
    mes_cap=P["mes"]; ano=P["ano"]; WK=cfg.work_dir()
    SRC=cfg.timeline_pptx_src()   # editável acumulado compartilhado (DEC-025) -> anexo P1.1 -> 01_FONTES
    if not SRC or not os.path.exists(SRC):
        # DEC-022: sem o .pptx-fonte, interromper SOMENTE esta etapa. NÃO inventar/reconstruir a
        # Timeline nem usar o PDF anterior como fonte editável. As demais etapas seguem (processos
        # separados); a pendência é registrada no QA.
        cfg.log("TIMELINE PPTX NÃO LOCALIZADA — etapa da Timeline interrompida (nada gerado). "
                "NÃO será inventada nem reconstruída a partir do PDF (DEC-022). Forneça o .pptx-fonte "
                "em G1/Anexos/P1.1 (ou em 01_FONTES) e reexecute o 07.")
        cfg.set_qa_warning("timeline_pptx_ausente",
            "TIMELINE PPTX NÃO LOCALIZADA — a Timeline não foi gerada nesta execução; forneça o "
            ".pptx-fonte e reexecute o 07 (DEC-022).", active=True)
        return
    # .pptx localizado: limpa eventual pendência anterior (idempotência)
    cfg.set_qa_warning("timeline_pptx_ausente", "", active=False)
    WORK=os.path.join(WK,"02_TRABALHO",f"Timeline_ate_{mes_cap.lower()}{ano}.pptx"); shutil.copy2(SRC,WORK)

    ppt=win32.gencache.EnsureDispatch("PowerPoint.Application"); ppt.Visible=True
    atexit.register(lambda: ppt.Quit())
    pres=ppt.Presentations.Open(WORK, WithWindow=False)

    # --- descobrir slide-alvo (placeholder 'XXXX') e o grupo-modelo "abaixo" nele ---
    below_tmpl=None; target=None
    for sl in pres.Slides:
        for sh in sl.Shapes:
            if sh.Type==6:
                for it in deep(sh):
                    if it.HasTextFrame and it.TextFrame.HasText and "XXXX" in it.TextFrame.TextRange.Text:
                        below_tmpl=sh; target=sl; break
                if below_tmpl: break
        if below_tmpl: break
    tl_fallback = target is None
    if tl_fallback:
        target=pres.Slides(pres.Slides.Count)  # fallback: último slide (DEC-020)
        cfg.log("WARNING (DEC-020): placeholder 'XXXX' NÃO encontrado; usando o ÚLTIMO slide como "
                "fallback. A Timeline será gerada, mas esta etapa NÃO é aprovada automaticamente: "
                "exige CONFERÊNCIA VISUAL DA TIMELINE antes de considerar o QA aprovado.")
    # Registra/limpa o aviso de QA de forma idempotente (some se o placeholder for achado numa re-execução)
    cfg.set_qa_warning(
        "timeline_fallback_xxxx",
        "CONFERÊNCIA VISUAL DA TIMELINE NECESSÁRIA — placeholder 'XXXX' não encontrado; o script 07 "
        "usou o último slide como fallback (DEC-020). Confira a slide do mês antes de aprovar.",
        active=tl_fallback)
    # --- descobrir grupo-modelo "acima" (marco estilizado em qualquer slide) ---
    above_tmpl=None
    for sl in pres.Slides:
        for sh in sl.Shapes:
            if sh.Type==6:
                txt="".join(it.TextFrame.TextRange.Text for it in deep(sh) if it.HasTextFrame and it.TextFrame.HasText)
                if "Autorização de Repasse" in txt or re.search(r'\d\d/\d\d', txt):
                    above_tmpl=sh; break
        if above_tmpl: break
    if above_tmpl is None and below_tmpl is not None:
        above_tmpl=below_tmpl
    cfg.log(f"slide-alvo #{target.SlideIndex}; below_tmpl={below_tmpl.Name if below_tmpl else None}; "
            f"above_tmpl={above_tmpl.Name if above_tmpl else None}")

    # copiar o modelo "acima" para o slide-alvo
    above_tmpl.Copy(); target.Shapes.Paste()
    above_on=target.Shapes(target.Shapes.Count)
    off_below=circle_offset(below_tmpl) if below_tmpl else 0.0
    off_above=circle_offset(above_on)
    top_below=below_tmpl.Top if below_tmpl else above_on.Top
    top_above=above_on.Top

    # Posições por DATA ÚNICA: dois marcos na MESMA data compartilham o MESMO ponto (1 bolinha,
    # 2 setas — uma acima e uma abaixo), conforme a regra do usuário (DEC-013/DEC-025).
    uniq=[]
    for _dt,_,_,_ in MARCOS:
        if _dt not in uniq: uniq.append(_dt)
    mU=len(uniq)
    # eixo alargado (mas com recuo nas bordas p/ a descrição das pontas não ser cortada)
    X0,X1=90.0,760.0
    xs_u=[X0+(X1-X0)*k/(mU-1) for k in range(mU)] if mU>1 else [(X0+X1)/2]
    XP={_dt:xs_u[k] for k,_dt in enumerate(uniq)}
    # largura da descrição = no máx. o espaçamento entre pontos (evita colidir com o vizinho do mesmo lado)
    gap = (X1-X0)/(mU-1) if mU>1 else 300.0
    DW=min(150.0, gap-3.0); made=[]; seen=set()   # largura <= espaçamento: nunca colide na horizontal
    # Escalonamento vertical (imita o ajuste manual, DEC-025): alterna nível LONGE/PERTO entre
    # descrições do MESMO lado, permitindo caixas mais largas sem colidir com o vizinho.
    a_ord=[j for j,(d,ds,ab,li) in enumerate(MARCOS) if ab]
    b_ord=[j for j,(d,ds,ab,li) in enumerate(MARCOS) if not ab]
    level={}
    for k,j in enumerate(a_ord): level[j]=k%2   # 0=longe, 1=perto
    for k,j in enumerate(b_ord): level[j]=k%2
    # marcadores ACIMA sob o título "Timeline" (centro do slide) não usam o nível ALTO (colidiria
    # com o título) — força o nível PERTO ali.
    cx=(X0+X1)/2
    for j,(d,ds,ab,li) in enumerate(MARCOS):
        if ab and abs(XP[d]-cx)<90: level[j]=1
    for i,(dt,desc,above,light) in enumerate(MARCOS):
        dup = dt in seen          # 2º marco da MESMA data (ponto compartilhado): data já existe
        seen.add(dt)
        tmpl= above_on if above else below_tmpl
        tmpl.Copy(); target.Shapes.Paste()
        g=target.Shapes(target.Shapes.Count)
        off = off_above if above else off_below
        g.Left = XP[dt]-off
        g.Top = top_above if above else top_below
        col = LIGHT if light else DARK
        c=find_circle(g)
        if c: c.Fill.ForeColor.RGB=col
        a=find_arrow(g)
        if a: a.Line.ForeColor.RGB=col; a.Fill.ForeColor.RGB=col
        dbox,sbox=find_date_desc(g)
        # (1) data única no ponto compartilhado: no 2º marco da mesma data, apaga a data (deixa vazia)
        if dbox: dbox.TextFrame.TextRange.Text = ("" if dup else dt)
        if sbox:
            lvl=level.get(i,0)   # 0=longe do eixo, 1=perto
            sbox.TextFrame.TextRange.Text=desc
            sbox.Width=DW; sbox.Left=XP[dt]-DW/2
            try: sbox.TextFrame.TextRange.ParagraphFormat.Alignment=2
            except: pass
            try: sbox.TextFrame.TextRange.Font.Size=11
            except: pass
            try: sbox.TextFrame2.WordWrap=-1
            except: pass
            sbox.Height=64
            if above:
                # descrição colada à borda inferior da caixa (perto da seta); nível longe/perto
                sbox.Top = 126 if lvl==0 else 190
                try: sbox.TextFrame.VerticalAnchor=4   # bottom: texto hugs a base -> seta conecta
                except: pass
                if a:                                  # seta do ponto ATÉ logo abaixo do texto
                    top=sbox.Top+sbox.Height+4
                    bot=280 if dup else 305            # (2) dup: para ANTES do eixo (deixa a data)
                    a.Top=top; a.Height=max(12, bot-top)
            else:
                sbox.Top = 372 if lvl==1 else 430
                try: sbox.TextFrame.VerticalAnchor=1   # top
                except: pass
                if a:
                    top=330 if dup else 313            # (2) dup: começa DEPOIS da data
                    bot=sbox.Top-4
                    a.Top=top; a.Height=max(12, bot-top)
        made.append(g)
    cfg.log(f"marcos criados: {len(made)}")
    below_tmpl.Delete(); above_on.Delete()

    # subtítulo + título 'Timeline' (não sobrepor a faixa verde)
    for sh in target.Shapes:
        if sh.HasTextFrame and sh.TextFrame.HasText:
            t=sh.TextFrame.TextRange.Text
            if "Progresso das atividades" in t:
                sh.TextFrame.TextRange.Text=f"                Progresso das atividades executadas  [ {mes_cap}/{ano} ]"
            if t.strip()=="Timeline":
                sh.TextFrame.WordWrap=0; sh.Top=124
    pres.Save()
    PDF=os.path.join(WK,"04_PDFS_TEMP",f"Timeline_ate_{mes_cap.lower()}{ano}.pdf")
    pres.SaveAs(PDF, 32)
    tgt_index=target.SlideIndex
    pres.Close(); ppt.Quit()

    # renderizar TODAS as slides -> PNGs (landscape) para o corpo do relatório.
    # DEC-028: o corpo exibe a Timeline ACUMULADA COMPLETA (do 1º mês até o mês corrente),
    # uma página por mês/slide, e NÃO apenas a slide do mês. O 09 embute todas as timeline_p*.png
    # em ordem cronológica (slide 1 = 1º mês; última = mês corrente).
    import fitz
    outdir=os.path.join(WK,"03_IMAGENS","TIMELINE")
    for f in os.listdir(outdir):
        if f.endswith(".png"): os.remove(os.path.join(outdir,f))
    doc=fitz.open(PDF)
    for i,pg in enumerate(doc):
        pg.get_pixmap(dpi=200).save(os.path.join(outdir,f"timeline_p{i+1:02d}.png"))
    n=doc.page_count; doc.close()
    cfg.log(f"timeline: PDF salvo e {n} slides (1º mês -> {mes_cap}/{ano}; slide do mês #{tgt_index}) "
            f"-> timeline_p01..p{n:02d}.png")

if __name__=="__main__":
    main()

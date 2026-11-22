# -*- coding: utf-8 -*-
"""
09_montar_word.py — Monta o corpo do Word a partir do relatório-base da CG (SharePoint, cópia local),
inserindo as seções P1.1 e P1.2 com textos (cfg.dados()) e as imagens dimensionadas à área útil
(img_aspects.json de 08). P1.3/P1.4 (conteúdo CG) são PRESERVADOS. Só escreve arquivos locais.
Mês/ano/nº da medição são derivados de config; nenhum texto/dado do mês fica fixo aqui.
"""
import os, sys, shutil, glob, json
from copy import deepcopy
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def main():
    T=cfg.dados()
    P=cfg.params(); mes=cfg.mes_min(); ano=cfg.ano(); MED_NUM=getattr(T,"MEDICAO_NUM","")
    WK=cfg.work_dir(); IMG=os.path.join(WK,"03_IMAGENS")
    SRC=cfg.base_am_docx() or os.path.join(WK,"01_FONTES","BASE_RevAM.docx")
    if not SRC or not os.path.exists(SRC):
        cfg.log(f"ERRO: relatório-base da CG (.docx) não encontrado ({SRC})."); sys.exit(1)
    WORK=os.path.join(WK,"02_TRABALHO",f"RM{cfg.numero()}_montado.docx")
    shutil.copy2(SRC,WORK)
    # Capa: data = último dia da competência (DEC-058). Corrige a cópia de trabalho se a CG trouxe outra data.
    import capa
    _alvo=cfg.data_extenso(cfg.data_capa()); _antes,_res=capa.garantir_data_capa(WORK,_alvo)
    if _res is None: cfg.log(f"Capa: data '{_antes}' já é o fim da competência (DEC-058).")
    else:
        cfg.log(f"Capa: data '{_antes}' -> '{_alvo}' (DEC-058) {_res}")
        if not _res.get("preview_regenerada"):
            cfg.set_qa_warning("capa_preview","CAPA: texto da data corrigido, mas a imagem impressa da capa não foi regenerada (PowerPoint indisponível) — conferir a capa")
        else: cfg.set_qa_warning("capa_preview","",active=False)
    doc=Document(WORK)
    ASP=json.load(open(os.path.join(WK,"02_TRABALHO","img_aspects.json")))
    # Layout das tabelas de reuniões (04): largura = largura útil do corpo, escala 1:1 (DEC-027).
    _lay=os.path.join(WK,"02_TRABALHO","reunioes_layout.json")
    LAYREU=json.load(open(_lay,encoding="utf-8")) if os.path.exists(_lay) else {}
    PW,PH=6.8,8.5     # área útil retrato (pol.)
    PH_FULL=10.1      # retrato "página cheia": imagens que ocupam a página toda (ex.: Dashboard),
                      # descendo até quase o risco do rodapé sem estourar (body útil ~10,3in). Ver DEC-024.
    LW,LH=10.4,6.5    # área útil paisagem
    # Largura EXATA do corpo de texto justificado (retrato) = page_w - margens laterais.
    # Usada por imagens que devem casar 100% com o início/fim do texto (ex.: Quadro Financeiro
    # Consolidado): a imagem é redimensionada para essa largura exata. DEC-029.
    _ret=[s for s in doc.sections if s.page_width<=s.page_height]
    BODY_W = round(_ret[0].page_width.inches-_ret[0].left_margin.inches-_ret[0].right_margin.inches,3) if _ret else 6.891
    def aspect(relpath):
        a=ASP.get(relpath); return a["aspect"] if a else 1.4
    def idx_of(pred, start=0):
        for i in range(start,len(doc.paragraphs)):
            if pred(doc.paragraphs[i].text): return i
        return None
    def title(ref, text):
        p=ref.insert_paragraph_before()
        try: p.style=doc.styles['List Bullet']
        except: pass
        r=p.add_run(text); r.bold=False; r.font.size=Pt(11)
        p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(4)
        return p
    def body(ref, text, bold=False):
        p=ref.insert_paragraph_before(text); p.alignment=AL.JUSTIFY
        for r in p.runs: r.font.size=Pt(11); r.bold=bold
        p.paragraph_format.space_after=Pt(5); return p
    def body_seg(ref, segs):
        p=ref.insert_paragraph_before(); p.alignment=AL.JUSTIFY
        for txt,bold in segs:
            r=p.add_run(txt); r.font.size=Pt(11); r.bold=bold
        p.paragraph_format.space_after=Pt(5); return p
    def _pic(ref, path, w):
        p=ref.insert_paragraph_before(); p.alignment=AL.CENTER
        p.add_run().add_picture(path, width=Inches(w))
        p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(6)
        return p
    def fit_img(ref, folder, fname, orient='portrait', area=None):
        path=os.path.join(IMG,folder,fname)
        if not os.path.exists(path): return
        a=aspect(f"{folder}/{fname}")
        if area: uw,uh=area
        else: uw,uh=(LW,LH) if orient=='landscape' else (PW,PH)
        _pic(ref, path, min(uw, uh*a))
    def fit_img_bodywidth(ref, folder, fname, width_in, max_h):
        """Insere a imagem com largura EXATA = width_in (largura do corpo justificado), casando
        início/fim com o texto (DEC-029). Só reduz a largura se, nessa largura, a altura passar
        de max_h — para não estourar a página."""
        path=os.path.join(IMG,folder,fname)
        if not os.path.exists(path): return
        a=aspect(f"{folder}/{fname}")
        w=width_in
        if w/a > max_h: w=max_h*a
        _pic(ref, path, w)
    def tabela_reunioes(ref, folder, chave):
        """Tabelas de reuniões: largura EXATA = largura útil do corpo daquela página, para que as
        bordas esquerda/direita coincidam com o início/fim do texto justificado (DEC-027). A altura
        de cada parte já foi dimensionada pelo 04 para caber na página — aqui não se redimensiona
        por altura (isso estreitaria a tabela e desalinharia). Se o 04 concluiu que a tabela
        aproveita melhor a página em PAISAGEM, ela entra em seção própria (como Controle/Timeline).
        """
        L=LAYREU.get(chave)
        arqs=sorted(glob.glob(os.path.join(IMG,folder,"*.png")))
        if not arqs: return
        paisagem = bool(L) and L.get("orientacao")=="paisagem"
        if paisagem: secbreak(ref,'portrait')
        for f in arqs:
            if L: _pic(ref, f, L["largura_in"])
            else: fit_img(ref, folder, os.path.basename(f), 'portrait')
        if paisagem: secbreak(ref,'landscape')

    # ---- Seções PENDENTES (versão EM ELABORAÇÃO — DEC-064) ----
    # dados_mes.PENDENCIAS = {"reunioes": "texto", "dashboard": "texto", ...}. No lugar do conteúdo que depende de
    # insumo ainda não recebido entra UM quadro sombreado e emoldurado, com o rótulo "SEÇÃO EM ELABORAÇÃO" — fácil de
    # achar e substituir (basta fornecer o insumo e remontar). O 13 reprova o PDF enquanto houver quadro pendente.
    PEND=getattr(T,"PENDENCIAS",{}) or {}
    def pendente(ref, chave):
        txt=PEND.get(chave)
        if not txt: return False
        p=ref.insert_paragraph_before(); p.alignment=AL.LEFT
        pPr=p._p.get_or_add_pPr()
        bdr=OxmlElement('w:pBdr')
        for lado in ('top','left','bottom','right'):
            e=OxmlElement(f'w:{lado}'); e.set(qn('w:val'),'single'); e.set(qn('w:sz'),'8'); e.set(qn('w:space'),'6')
            e.set(qn('w:color'),'1C5560'); bdr.append(e)
        pPr.append(bdr)
        shd=OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'EAF1F2')
        pPr.append(shd)
        r=p.add_run("SEÇÃO EM ELABORAÇÃO — "); r.bold=True; r.font.size=Pt(10.5); r.font.color.rgb=RGBColor(0x1C,0x55,0x60)
        r=p.add_run(txt); r.font.size=Pt(10.5); r.font.color.rgb=RGBColor(0x1C,0x55,0x60)
        p.paragraph_format.space_before=Pt(8); p.paragraph_format.space_after=Pt(10)
        p.paragraph_format.left_indent=Inches(0.1); p.paragraph_format.right_indent=Inches(0.1)
        cfg.set_qa_warning(f"pendente_{chave}", f"SEÇÃO EM ELABORAÇÃO ({chave}): {txt}")
        return True
    for _k in ("reunioes","dashboard","p12","p12_consolidado","financeiro"):
        if _k not in PEND: cfg.set_qa_warning(f"pendente_{_k}","",active=False)

    def pagebreak(ref):
        p=ref.insert_paragraph_before()
        run=p.add_run(); br=OxmlElement('w:br'); br.set(qn('w:type'),'page'); run._r.append(br)
    def secbreak(ref, orient):
        sp=deepcopy(doc.sections[-1]._sectPr)
        for t in sp.findall(qn('w:titlePg')): sp.remove(t)
        pgSz=sp.find(qn('w:pgSz'))
        if orient=='landscape':
            pgSz.set(qn('w:w'),'16838'); pgSz.set(qn('w:h'),'11906'); pgSz.set(qn('w:orient'),'landscape')
            pgMar=sp.find(qn('w:pgMar'))
            if pgMar is not None:
                pgMar.set(qn('w:left'),'720'); pgMar.set(qn('w:right'),'720'); pgMar.set(qn('w:top'),'900'); pgMar.set(qn('w:bottom'),'720')
        else:
            pgSz.set(qn('w:w'),'11906'); pgSz.set(qn('w:h'),'16838')
            if pgSz.get(qn('w:orient')): pgSz.set(qn('w:orient'),'portrait')
        p=ref.insert_paragraph_before(); p._p.get_or_add_pPr().append(sp); return p

    # remover stubs "Produtos desenvolvidos no período" das seções P1.1/P1.2
    i_p11=idx_of(lambda t:t.strip()=="Produtos desenvolvidos no período")
    i_p12h=idx_of(lambda t:t.strip().startswith("Produto 1.2"))
    i_p12=idx_of(lambda t:t.strip()=="Produtos desenvolvidos no período", start=i_p12h)
    i_p13h=idx_of(lambda t:t.strip().startswith("Produto 1.3"))
    def dele(a,b):
        for k in range(b-1,a,-1):
            el=doc.paragraphs[k]._element; el.getparent().remove(el)
    dele(i_p12,i_p13h)
    i_p12h=idx_of(lambda t:t.strip().startswith("Produto 1.2")); i_p11=idx_of(lambda t:t.strip()=="Produtos desenvolvidos no período")
    dele(i_p11,i_p12h)

    # ===== P1.1 =====
    ref=doc.paragraphs[idx_of(lambda t:t.strip().startswith("Produto 1.2"))]
    title(ref,f"Planilha de Controle de Produtos referente ao mês de {mes}.")
    for t in T.CONTROLE: body(ref,t)
    secbreak(ref,'portrait')
    for f in sorted(glob.glob(os.path.join(IMG,"CONTROLE_PRODUTOS","controle_p*.png"))):
        fit_img(ref,"CONTROLE_PRODUTOS",os.path.basename(f),'landscape')
    secbreak(ref,'landscape')
    title(ref,f"Dashboard Analítico das Entregas Executadas em {mes}.")
    for t in T.DASHBOARD: body(ref,t)
    if not pendente(ref,"dashboard"):
        pagebreak(ref)
        # Dashboard: imagem "página cheia" — ocupa quase toda a área útil (perto do rodapé), DEC-024.
        fit_img(ref,"DASHBOARD","dashboard_p01.png",'portrait', area=(PW, PH_FULL))
    # A próxima seção começa em página nova via page_break_before (NÃO inserir parágrafo de quebra
    # vazio: com o Dashboard quase preenchendo a página, o parágrafo vazio geraria uma página em
    # branco). DEC-024.
    _tl=title(ref,f"Linha do Tempo do Progresso das Atividades Executadas em {mes} de {ano}.")
    _tl.paragraph_format.page_break_before=True
    for t in T.TIMELINE: body(ref,t)
    secbreak(ref,'portrait')
    for f in sorted(glob.glob(os.path.join(IMG,"TIMELINE","timeline_p*.png"))):
        fit_img(ref,"TIMELINE",os.path.basename(f),'landscape')
    secbreak(ref,'landscape')
    title(ref,f"Reuniões de Governança realizadas durante o mês de {mes}.")
    if not pendente(ref,"reunioes"):
        for t in T.REUNIOES: body(ref,t)
        tabela_reunioes(ref,"REUNIOES_P11","gov_p11")
    _mm,_am=cfg.competencia_medicao()   # DEC-063: competência de EXECUÇÃO da medição (mês anterior)
    title(ref,f"Análise do processo de medição do mês anterior - {_mm.lower()}/{_am} - {MED_NUM}ª medição.")
    body(ref,T.MEDICAO_INTRO)
    # Quadro Financeiro Consolidado: largura EXATA = largura do corpo justificado (DEC-029).
    fit_img_bodywidth(ref,"FINANCEIRO","quadro_consolidado.png", BODY_W, PH)
    body_seg(ref,T.MEDICAO_FIG_SEG)
    fit_img(ref,"FINANCEIRO","resumo_empresas.png",'portrait')
    body_seg(ref,T.MEDICAO_AVAL_SEG)
    # Base Técnica Parceiro: pode vir como 1 imagem (base_parceiro.png) ou 2 (base_parceiro_p1/p2) quando foi
    # dividida no 05 para minimizar o espaço em branco antes da quebra de página (DEC-031).
    _bases=sorted(glob.glob(os.path.join(IMG,"FINANCEIRO","base_parceiro*.png")))
    for _b in _bases:
        fit_img(ref,"FINANCEIRO",os.path.basename(_b),'portrait')
    body_seg(ref,T.MEDICAO_PLANO_SEG)
    body_seg(ref,T.MEDICAO_FIM_SEG)
    for f in sorted(glob.glob(os.path.join(IMG,"FINANCEIRO","linha_med_*.png"))):
        fit_img(ref,"FINANCEIRO",os.path.basename(f),'portrait')
    # Síntese do Plano de Gestão Integrada de Riscos (conteúdo mínimo do P1.1 — item 3.6.2.1 do contrato;
    # determinação (III) do Parceiro para a competência de outubro/2026 — DEC-065). Fonte: produtos do P4.3 do mês.
    _riscos=getattr(T,"RISCOS",None)
    if _riscos:
        title(ref,f"Síntese do Plano de Gestão Integrada de Riscos - {mes}/{ano}.")
        for t in _riscos.get("intro",[]): body(ref,t)
        _tab=_riscos.get("tabela")
        if _tab:
            tb=doc.add_table(rows=1,cols=len(_tab["colunas"])); tb.style='Table Grid'; tb.alignment=WD_TABLE_ALIGNMENT.CENTER
            for j,c in enumerate(_tab["colunas"]):
                cel=tb.rows[0].cells[j]; cel.text=""; rr=cel.paragraphs[0].add_run(c); rr.bold=True; rr.font.size=Pt(9)
                rr.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
                sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:color'),'auto'); sh.set(qn('w:fill'),'1C5560')
                cel._tc.get_or_add_tcPr().append(sh)
            for linha in _tab["linhas"]:
                cs=tb.add_row().cells
                for j,v in enumerate(linha):
                    cs[j].text=""; rr=cs[j].paragraphs[0].add_run(str(v)); rr.font.size=Pt(9)
                    cs[j].paragraphs[0].alignment=AL.LEFT   # célula justificada espaça mal as palavras
            for cel in tb.rows[0].cells: cel.paragraphs[0].alignment=AL.LEFT
            for row in tb.rows:
                for j,w in enumerate(_tab.get("larguras",[])):
                    row.cells[j].width=Inches(w)
            ref._p.addprevious(tb._tbl)
            cap=ref.insert_paragraph_before(_tab.get("fonte","")); cap.alignment=AL.LEFT
            for r in cap.runs: r.font.size=Pt(9); r.italic=True
        for t in _riscos.get("fecho",[]): body(ref,t)

    # ===== P1.2 =====
    ref3=doc.paragraphs[idx_of(lambda t:t.strip().startswith("Produto 1.3"))]
    title(ref3,f"Integração Institucional e Consolidação de Processos executados em {mes}.")
    _p12_pend=pendente(ref3,"p12")
    if not _p12_pend:
        for t in T.P12: body(ref3,t)
        tabela_reunioes(ref3,"REUNIOES_P12","inst_p12")
    # Texto consolidado do P1.2 ("Relatório de Relacionamento Institucional"), APÓS a tabela de
    # reuniões — redigido a cada mês a partir do modelo + planilha de reuniões 1.2 (DEC-033).
    _cons=getattr(T,"P12_CONSOLIDADO",None)
    if _p12_pend:
        title(ref3,f"Relatório de Relacionamento Institucional do mês de {mes} de {ano}.")
        pendente(ref3,"p12_consolidado")
    elif _cons:
        title(ref3,f"Relatório de Relacionamento Institucional do mês de {mes} de {ano}.")
        for t in _cons: body(ref3,t)
    else:
        cfg.log("AVISO (DEC-033): P12_CONSOLIDADO ausente — texto consolidado do P1.2 NÃO inserido "
                "(o QA vai acusar). Redija-o a partir do modelo + planilha de reuniões 1.2.")

    # Cada título "Produto 1.X – ..." deve iniciar em PÁGINA NOVA (DEC-032). Antes, só o 1.1 tinha
    # quebra (as demais dependiam do fluxo e caíam no meio da página — ex.: 1.2 e 1.4). Aplica
    # page_break_before em todos e remove quebra manual redundante no parágrafo anterior (evita
    # página em branco dupla no que já tinha quebra).
    for k,p in enumerate(doc.paragraphs):
        if p.style.name.startswith("Heading") and p.text.strip().startswith("Produto 1."):
            p.paragraph_format.page_break_before=True
            if k>0:
                for r in doc.paragraphs[k-1].runs:
                    for br in r._r.findall(qn('w:br')):
                        if br.get(qn('w:type'))=='page': r._r.remove(br)

    doc.save(WORK)
    cfg.log(f"Word montado: {WORK} (paras: {len(doc.paragraphs)})")

if __name__=="__main__":
    main()

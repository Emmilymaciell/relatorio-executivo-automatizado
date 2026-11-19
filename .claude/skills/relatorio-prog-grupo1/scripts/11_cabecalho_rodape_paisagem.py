# -*- coding: utf-8 -*-
"""
11_cabecalho_rodape_paisagem.py — Adapta cabeçalho/rodapé das seções em PAISAGEM (linha mais larga,
Programa à esquerda; rodapé -CONFIDENCIAL- + nº de página) SEM tocar nas seções retrato. Limpa todo o
elemento <w:hdr>/<w:ftr> antes de reconstruir (evita header duplicado). Só escreve arquivos locais.
Parametrizado por prog_config (nenhum caminho fixo).
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL, WD_TAB_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def border(p, side):
    pPr=p._p.get_or_add_pPr()
    pBdr=pPr.find(qn('w:pBdr'))
    if pBdr is None: pBdr=OxmlElement('w:pBdr'); pPr.append(pBdr)
    e=OxmlElement('w:'+side); e.set(qn('w:val'),'single'); e.set(qn('w:sz'),'6'); e.set(qn('w:space'),'6'); e.set(qn('w:color'),'000000')
    pBdr.append(e)
def page_field(p):
    r=p.add_run()
    fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),'begin'); r._r.append(fc)
    it=OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text=' PAGE '; r._r.append(it)
    fc2=OxmlElement('w:fldChar'); fc2.set(qn('w:fldCharType'),'end'); r._r.append(fc2)

def main():
    N=cfg.numero(); WK=cfg.work_dir(); WORK=os.path.join(WK,"02_TRABALHO",f"RM{N}_montado.docx")
    if not os.path.exists(WORK):
        cfg.log(f"ERRO: {WORK} não encontrado. Rode 09/10 antes."); sys.exit(1)
    d=Document(WORK)
    # Nenhuma seção pode ter "primeira página diferente": senão a 1ª página da seção usa um
    # cabeçalho/rodapé de "primeira página" (vazio) e fica SEM cabeçalho/rodapé — era o caso da
    # seção final (pág. 35). Todas as páginas devem ter cabeçalho e rodapé (DEC-032).
    n_dfp=0
    for sec in d.sections:
        if sec.different_first_page_header_footer:
            sec.different_first_page_header_footer=False; n_dfp+=1
    if n_dfp: cfg.log(f"'primeira página diferente' removido de {n_dfp} seção(ões) (DEC-032)")
    n_ls=0
    for sec in d.sections:
        is_ls = sec.orientation==WD_ORIENT.LANDSCAPE or sec.page_width>sec.page_height
        if not is_ls: continue
        n_ls+=1
        textw = sec.page_width - sec.left_margin - sec.right_margin
        sec.header.is_linked_to_previous=False
        h=sec.header; hel=h._element
        for child in list(hel): hel.remove(child)
        hp=h.add_paragraph(); hp.alignment=AL.LEFT
        r=hp.add_run("Programa | Consórcio do Programa")
        r.font.size=Pt(9); r.bold=False; r.font.color.rgb=RGBColor(0x40,0x40,0x40)
        border(hp,'bottom')
        sec.footer.is_linked_to_previous=False
        f=sec.footer; fel=f._element
        for child in list(fel): fel.remove(child)
        fp=f.add_paragraph(); border(fp,'top')
        tabs=fp.paragraph_format.tab_stops
        tabs.add_tab_stop(textw, WD_TAB_ALIGNMENT.RIGHT)
        r1=fp.add_run("-CONFIDENCIAL-"); r1.font.size=Pt(9); r1.font.color.rgb=RGBColor(0x40,0x40,0x40)
        fp.add_run("\t"); page_field(fp)
        for rr in fp.runs: rr.font.size=Pt(9)
    d.save(WORK)
    cfg.log(f"seções paisagem adaptadas: {n_ls}")

if __name__=="__main__":
    main()

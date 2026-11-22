# -*- coding: utf-8 -*-
"""
10_encerramento_assinaturas_anexos.py — Ajusta ENCERRAMENTO (negrito seletivo: só Nº/mês/período),
data única (Cidade, DEC-003/005), 2 blocos de assinatura (Pessoa B/Pessoa A — DEC-006/041), remove tabelas
genéricas 'Nome – Programa', e adiciona 'ANEXOS' (título sem numeração — DEC-007) + Quadro de Anexos.
Só escreve arquivos locais. Número/mês/ano/período/nomes de anexos derivam de config (nada fixo).
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def data_extenso(ddmmaaaa):
    d,m,a=ddmmaaaa.split("/")
    return f"{int(d):02d} de {cfg.MESES_PT[int(m)-1].lower()} de {a}"

def main():
    P=cfg.params(); N=cfg.numero(); mes=cfg.mes_min(); ano=cfg.ano()
    mm=int(P["mes_numero"]); yy=int(ano)%100
    WK=cfg.work_dir(); WORK=os.path.join(WK,"02_TRABALHO",f"RM{N}_montado.docx")
    if not os.path.exists(WORK):
        cfg.log(f"ERRO: {WORK} não encontrado. Rode 09 antes."); sys.exit(1)
    d=Document(WORK)
    def idx_of(pred, start=0):
        for i in range(start,len(d.paragraphs)):
            if pred(d.paragraphs[i].text): return i
        return None
    # remover tabelas genéricas 'Nome - Programa'
    for tb in list(d.tables):
        joined=" ".join(c.text for r in tb.rows for c in r.cells)
        if "Nome – Programa" in joined or "Nome - Programa" in joined:
            tb._tbl.getparent().remove(tb._tbl)
    # ENCERRAMENTO (negrito seletivo)
    ie=idx_of(lambda t:t.strip()=="ENCERRAMENTO")
    ib=idx_of(lambda t:t.startswith("O Relatório Mensal") or t.startswith("O Relatório de Produtos Mensal"), start=ie)
    per_ini=data_extenso(P["periodo_inicio"]); per_fim=data_extenso(P["periodo_fim"])
    ENC_SEG=[("O ",False),(f"{cfg.termo_documento()} N° {N}",True),
     (", referente à consolidação e à apresentação das atividades desenvolvidas durante o mês de ",False),
     (f"{mes} de {ano}",True),(", referente ao período de ",False),
     (f"{per_ini} até {per_fim}",True),
     (", em conformidade ao Contrato N° 000/0000-Parceiro.",False)]
    if ib is not None:
        p=d.paragraphs[ib]
        for child in list(p._p):
            if child.tag!=qn('w:pPr'): p._p.remove(child)
        for txt,bold in ENC_SEG:
            rr=p.add_run(txt); rr.bold=bold; rr.font.size=Pt(11)
        p.alignment=AL.JUSTIFY
    # data única (Cidade) = data da capa = último dia da competência (DEC-058)
    doc_date=cfg.data_documento()
    data_str=f"Cidade, {data_extenso(doc_date)}"
    for i in range(ie, len(d.paragraphs)):
        p=d.paragraphs[i]
        if p.text.strip().startswith("Cidade") or p.text.strip().startswith("São Paulo"):
            for child in list(p._p):
                if child.tag!=qn('w:pPr'): p._p.remove(child)
            p.add_run(data_str).font.size=Pt(11); p.alignment=AL.LEFT; break
    for j in range(len(d.paragraphs)-1, idx_of(lambda t:t.strip().startswith("Cidade"),start=ie), -1):
        if d.paragraphs[j].text.strip().startswith(("Cidade","São Paulo")):
            d.paragraphs[j]._element.getparent().remove(d.paragraphs[j]._element)
    date_para=d.paragraphs[idx_of(lambda t:t.strip().startswith("Cidade"),start=ie)]

    # blocos de assinatura (2) — DEC-006
    sig=d.add_table(rows=1, cols=2); sig.alignment=WD_TABLE_ALIGNMENT.CENTER
    def fill(cell,name,prod):
        cell.paragraphs[0].text=""
        for _ in range(2):
            bp=cell.add_paragraph(); bp.paragraph_format.space_after=Pt(0)
        pl=cell.add_paragraph(); pl.alignment=AL.CENTER; pl.paragraph_format.space_after=Pt(0)
        pl.add_run("_"*38).font.size=Pt(10)
        for line,bold in [(name,True),("Programa – Consórcio do Programa",False),(prod,False)]:
            pp=cell.add_paragraph(); pp.alignment=AL.CENTER; pp.paragraph_format.space_after=Pt(0); pp.paragraph_format.space_before=Pt(0)
            r=pp.add_run(line); r.bold=bold; r.font.size=Pt(10)
    fill(sig.cell(0,0),"Pessoa B","Produto P1.3 e P1.4 – Relatório mensal – Grupo 1")
    # Signatário de P1.1/P1.2 por competência (parametros.json "assinante_p11_p12"); só o NOME no bloco —
    # nunca imagem/assinatura simulada: o documento só é assinado depois de finalizado (usuária, 16/11/2026).
    fill(sig.cell(0,1),P.get("assinante_p11_p12") or "Pessoa A","Produto P1.1 e P1.2 – Relatório mensal – Grupo 1")
    date_para._p.addnext(sig._tbl)

    # ANEXOS (título sem numeração) + Quadro de Anexos
    pana=d.add_paragraph()
    try: pana.style=d.styles['Heading 1']
    except: pass
    pana.add_run("ANEXOS").bold=True
    pPr=pana._p.get_or_add_pPr(); numPr=OxmlElement('w:numPr')
    il=OxmlElement('w:ilvl'); il.set(qn('w:val'),'0'); ni=OxmlElement('w:numId'); ni.set(qn('w:val'),'0')
    numPr.append(il); numPr.append(ni); pPr.append(numPr)
    # Quadro de Anexos = VARREDURA REAL das pastas de produto do G1 (P1.1..P1.4): lista TODOS os
    # arquivos que os grupos subiram, com o nome EXATO que está na pasta, na ordem P1.1→P1.4 (DEC-042).
    # NUNCA agrega/zipa; nunca mexe nos anexos de P1.3/P1.4. Numeração romana conforme a qtd de arquivos.
    def _roman(n):
        vals=[(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
              (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]
        s=''
        for v,sym in vals:
            while n>=v: s+=sym; n-=v
        return s
    reais=cfg.listar_anexos_reais()
    # POP Quadro 4 (DEC-043/066): no pacote, todo anexo que não é PDF nem ZIP vai em ZIP individual com o mesmo nome.
    # O Quadro de Anexos usa o nome FINAL (o mesmo do Índice e da pasta do mês), não o da cópia de trabalho.
    if cfg.layout_anexos()=="plano":
        reais=[(s, n if n.lower().endswith((".pdf",".zip")) else os.path.splitext(n)[0]+".zip") for s,n in reais]
    anx=[(f"Anexo {_roman(i+1)}", nome) for i,(sub,nome) in enumerate(reais)]
    if not anx:
        cfg.log("AVISO: nenhum arquivo de anexo encontrado em G1\\Anexos (P1.1..P1.4).")
    else:
        cfg.log(f"anexos (varredura real): {len(anx)} arquivo(s) — " +
                ", ".join(f"{s.split(' - ')[0]}:{sum(1 for ss,_ in reais if ss==s)}" for s in dict.fromkeys(ss for ss,_ in reais)))
    tb=d.add_table(rows=2+len(anx), cols=2)
    try: tb.style=d.styles['Table Grid']
    except: pass
    tb.alignment=WD_TABLE_ALIGNMENT.CENTER
    def shade(cell,color):
        sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),color); cell._tc.get_or_add_tcPr().append(sh)
    h=tb.cell(0,0).merge(tb.cell(0,1)); h.paragraphs[0].alignment=AL.CENTER
    r=h.paragraphs[0].add_run("LISTA DE DOCUMENTOS ANEXOS"); r.bold=True; r.font.color.rgb=RGBColor(255,255,255); r.font.size=Pt(11); shade(h,"1F7A54")
    for j,txt in [(0,"REF."),(1,"DESCRIÇÃO")]:
        c=tb.cell(1,j); c.paragraphs[0].alignment=AL.CENTER; rr=c.paragraphs[0].add_run(txt); rr.bold=True; rr.font.color.rgb=RGBColor(255,255,255); shade(c,"2E8B62")
    for i,(rf,dc) in enumerate(anx):
        c0=tb.cell(2+i,0); c0.paragraphs[0].alignment=AL.CENTER; c0.paragraphs[0].add_run(rf).bold=True
        tb.cell(2+i,1).text=dc
    for row in tb.rows: row.cells[0].width=Inches(1.3); row.cells[1].width=Inches(5.2)
    d.save(WORK)
    cfg.log(f"encerramento+assinaturas+anexos OK (tabelas: {len(d.tables)})")

if __name__=="__main__":
    main()

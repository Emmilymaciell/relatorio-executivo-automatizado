# -*- coding: utf-8 -*-
"""md_para_docx.py — converte os relatórios em Markdown da Coordenação (conformidade, prontidão) para .docx simples:
títulos (#/##), parágrafos, listas (- e 1.), tabelas (|…|), **negrito**, *itálico* e `código`. Sem dependência de
pandoc. Uso: python scripts/manutencao/md_para_docx.py <entrada.md> [<saida.docx>]"""
import os, re, sys
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

RX = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)")


def runs(par, texto, tam=10):
    for parte in RX.split(texto):
        if not parte:
            continue
        if parte.startswith("**"):
            r = par.add_run(parte[2:-2]); r.bold = True
        elif parte.startswith("`"):
            r = par.add_run(parte[1:-1]); r.font.name = "Consolas"
        elif parte.startswith("*"):
            r = par.add_run(parte[1:-1]); r.italic = True
        else:
            r = par.add_run(parte)
        r.font.size = Pt(tam)


def sombra(cel, cor):
    sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), cor)
    cel._tc.get_or_add_tcPr().append(sh)


def converter(md, saida):
    d = docx.Document()
    for s in d.sections:
        s.left_margin = s.right_margin = Cm(1.8); s.top_margin = s.bottom_margin = Cm(1.6)
    st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10)
    linhas = open(md, encoding="utf-8").read().splitlines()
    i = 0
    while i < len(linhas):
        l = linhas[i].rstrip()
        if l.startswith("|"):
            bloco = []
            while i < len(linhas) and linhas[i].startswith("|"):
                bloco.append(linhas[i]); i += 1
            rows = [[c.strip() for c in b.strip("|").split("|")] for b in bloco if not re.match(r"^\|[\s:\-|]+\|$", b)]
            t = d.add_table(rows=len(rows), cols=len(rows[0])); t.style = "Table Grid"
            for a, row in enumerate(rows):
                for b, c in enumerate(row):
                    cel = t.cell(a, b); cel.text = ""
                    runs(cel.paragraphs[0], c, 9)
                    if a == 0:
                        sombra(cel, "1C5560")
                        for r in cel.paragraphs[0].runs:
                            r.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            d.add_paragraph()
            continue
        if l.startswith("# "):
            h = d.add_heading(l[2:], level=0)
        elif l.startswith("## "):
            d.add_heading(l[3:], level=1)
        elif l.strip() == "---":
            pass
        elif re.match(r"^\s*- ", l):
            texto = re.sub(r"^\s*- ", "", l)
            while i + 1 < len(linhas) and re.match(r"^\s{2,}\S", linhas[i + 1]) and not re.match(r"^\s*- ", linhas[i + 1]):
                i += 1; texto += " " + linhas[i].strip()
            p = d.add_paragraph(style="List Bullet"); runs(p, texto)
            if l.startswith("  "):
                p.paragraph_format.left_indent = Cm(1.2)
        elif re.match(r"^\d+\. ", l):
            # numeração literal (a do Word continuaria a contagem entre seções)
            texto = l.strip()
            while i + 1 < len(linhas) and re.match(r"^\s{2,}\S", linhas[i + 1]):
                i += 1; texto += " " + linhas[i].strip()
            p = d.add_paragraph(); runs(p, texto)
            p.paragraph_format.left_indent = Cm(0.6); p.paragraph_format.first_line_indent = Cm(-0.6)
        elif l.startswith("  ") and l.strip() and d.paragraphs:
            runs(d.paragraphs[-1], " " + l.strip())
        elif l.strip():
            # parágrafo: junta linhas de continuação
            texto = l.strip()
            while i + 1 < len(linhas) and linhas[i + 1].strip() and not re.match(r"^(#|\||- |\s+- |\d+\. |---)", linhas[i + 1]):
                i += 1; texto += " " + linhas[i].strip()
            runs(d.add_paragraph(), texto)
        i += 1
    d.save(saida)
    return saida


if __name__ == "__main__":
    md = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(md)[0] + ".docx"
    print(converter(md, out))

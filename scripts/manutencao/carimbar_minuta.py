"""Gera cópias carimbadas "MINUTA – EM REVISÃO" dos 4 documentos do POP para divulgação informal (DEC-050).

Os originais em docs/referencias-oficiais/ são só LIDOS (copiados para uma pasta temporária, processados lá e
copiados para o destino). Usado em 02/10/2026 para a divulgação via WhatsApp antes da aprovação formal.

    python scripts/manutencao/carimbar_minuta.py [--destino PASTA] [--texto "MINUTA – EM REVISÃO"]

Como cada formato recebe a marca (decisões tomadas depois de testar e conferir no PDF exportado):
- DOCX: VML de marca d'água no cabeçalho, igual ao que o Word grava ("PowerPlusWaterMarkObject"), no último
  parágrafo que já existe (parágrafo novo aumentaria o cabeçalho e repaginaria). Documento de uma página feito de
  tabela sombreada (Modelo de Índice): a marca vai no corpo, porque o Word desenha o cabeçalho ATRÁS das células.
- XLSX: WordArt semitransparente em cada aba + texto no cabeçalho de impressão.
- PDF: texto sobreposto com opacidade; wrap_contents() isola a CTM do conteúdo original (sem isso saía espelhado).
"""
import argparse
import os
import shutil
import stat
import tempfile
from pathlib import Path

import docx
import fitz
import win32com.client as win32
from docx.oxml import parse_xml
from docx.oxml.ns import qn

REPO = Path(__file__).resolve().parents[2]
REF = REPO / "docs" / "referencias-oficiais"
TXT = "MINUTA – EM REVISÃO"
SUF = " - MINUTA EM REVISAO"

ARQS = [  # (original vigente, nome da cópia carimbada)
    (REF / "v2026-10-01" / "POP - Fluxo Mensal de Entrega de Produtos Programa.docx",
     "POP - Fluxo Mensal de Entrega de Produtos Programa" + SUF + ".docx"),
    (REF / "v2026-10-01" / "Checklist_Visual_Entregas_PROG_FINAL.xlsx",
     "Checklist Visual de Entregas Programa" + SUF + ".xlsx"),
    (REF / "v2026-10-01-r2" / "Modelo de Indice - Prestacao de Contas Programa.docx",
     "Modelo de Indice - Entrega de Produtos Programa" + SUF + ".docx"),
    (REF / "v2026-10-01" / "Fluxograma - Fluxo Mensal de Entrega de Produtos Programa.pdf",
     "Fluxograma - Fluxo Mensal de Entrega de Produtos Programa" + SUF + ".pdf"),
]
MARCA_NO_CORPO = ("Modelo de Indice",)


def rgb(r, g, b):
    return r + g * 256 + b * 65536


# ---------------------------------------------------------------- DOCX (VML)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
      'xmlns:w10="urn:schemas-microsoft-com:office:word"')
SHAPETYPE = '''<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>
<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/>
<v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>
<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>
<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>
<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>
<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>'''


def _run_marca(texto, n, larg, rel):
    alt = larg / 6.5
    xml = parse_xml(f'''<w:p {NS}><w:r><w:rPr><w:noProof/></w:rPr><w:pict>
{SHAPETYPE if n == 1 else ''}
<v:shape id="PowerPlusWaterMarkObject{n}" o:spid="_x0000_s{2049 + n}" type="#_x0000_t136"
 style="position:absolute;margin-left:0;margin-top:0;width:{larg:.0f}pt;height:{alt:.0f}pt;rotation:315;z-index:251660000;mso-position-horizontal:center;mso-position-horizontal-relative:{rel};mso-position-vertical:center;mso-position-vertical-relative:{rel}"
 o:allowincell="f" fillcolor="#c00000" stroked="f"><v:fill opacity=".25"/>
<v:textpath style="font-family:&quot;Arial Black&quot;;font-size:1pt" string="{texto}"/>
<w10:wrap anchorx="{rel}" anchory="{rel}"/></v:shape></w:pict></w:r></w:p>''')
    return xml.find(qn("w:r"))


def carimbar_docx(path, texto):
    d = docx.Document(str(path))
    if path.name.startswith(MARCA_NO_CORPO):
        s = d.sections[0]
        larg = min(s.page_width.pt, s.page_height.pt) * 0.8
        d.element.body.find(qn("w:p")).append(_run_marca(texto, 1, larg, "page"))
        d.save(str(path))
        return 1
    n = 0
    for i, s in enumerate(d.sections):
        hdrs = [s.header]
        if s.different_first_page_header_footer:
            hdrs.append(s.first_page_header)
        if d.settings.odd_and_even_pages_header_footer:
            hdrs.append(s.even_page_header)
        for h in hdrs:
            if i > 0 and h.is_linked_to_previous:
                continue
            n += 1
            larg = min(s.page_width.pt, s.page_height.pt) * 0.8
            run = _run_marca(texto, n, larg, "margin")
            ps = h._element.findall(qn("w:p"))
            (ps[-1] if ps else h._element).append(run)
    d.save(str(path))
    return n


# ---------------------------------------------------------------- XLSX / PDF
def carimbar_xlsx(path, texto):
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    n = 0
    try:
        wb = app.Workbooks.Open(str(path))
        for ws in wb.Worksheets:
            if ws.ProtectContents:
                print("AVISO aba protegida (sem carimbo):", ws.Name)
                continue
            ur = ws.UsedRange
            w, h = ur.Width, ur.Height
            sh = ws.Shapes.AddTextEffect(0, texto, "Arial Black", 54, False, False, ur.Left, ur.Top)
            sh.Line.Visible = False
            sh.Fill.ForeColor.RGB = rgb(192, 0, 0)
            sh.Fill.Transparency = 0.72
            sh.LockAspectRatio = True
            sh.Width = max(min(w * 0.8, 620), 300)
            sh.Rotation = 335
            sh.Left = ur.Left + (w - sh.Width) / 2
            sh.Top = ur.Top + max((min(h, 700) - sh.Height) / 2, 20)
            sh.Placement = 3  # não move/redimensiona com as células
            sh.Name = "Carimbo_Minuta"
            ws.PageSetup.CenterHeader = '&"Arial"&B&14&KC00000' + texto
            n += 1
        wb.Save()
        wb.Close(False)
    finally:
        app.Quit()
    return n


def carimbar_pdf(path, texto):
    doc = fitz.open(str(path))
    font = fitz.Font(fontfile="C:/Windows/Fonts/ariblk.ttf")
    for page in doc:
        r = page.rect
        page.wrap_contents()
        fs = min(r.width, r.height) / 13
        tl = font.text_length(texto, fontsize=fs)
        c = fitz.Point(r.width / 2, r.height / 2)
        tw = fitz.TextWriter(r)
        tw.append(fitz.Point(c.x - tl / 2, c.y + fs / 3), texto, font=font, fontsize=fs)
        ang = 25 if r.width > r.height else 45
        tw.write_text(page, color=(0.75, 0, 0), opacity=0.28, morph=(c, fitz.Matrix(ang)), overlay=True)
    n = doc.page_count
    tmp = path.with_suffix(".tmp.pdf")
    doc.save(str(tmp), garbage=3, deflate=True)
    doc.close()
    os.replace(tmp, path)
    return n


def exportar_pdf_verificacao(paths, pasta):
    """PDF dos Office carimbados (só para conferência visual; ficam na pasta temporária)."""
    w = win32.DispatchEx("Word.Application")
    w.Visible = False
    try:
        for p in paths:
            if p.suffix == ".docx":
                d = w.Documents.Open(str(p), ReadOnly=True)
                d.ExportAsFixedFormat(str(pasta / (p.stem + ".verif.pdf")), 17)
                d.Close(False)
    finally:
        w.Quit()


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--destino", type=Path, default=REPO / "output" / "Divulgacao_POP_MINUTA")
    ap.add_argument("--texto", default=TXT)
    a = ap.parse_args()
    trab = Path(tempfile.mkdtemp(prefix="carimbo_minuta_"))
    feitos = []
    for src, nome in ARQS:
        dst = trab / nome
        shutil.copy2(src, dst)
        os.chmod(dst, stat.S_IWRITE | stat.S_IREAD)  # a cópia herda o somente-leitura do original
        fn = {".docx": carimbar_docx, ".xlsx": carimbar_xlsx, ".pdf": carimbar_pdf}[dst.suffix]
        print(f"{dst.name}: {fn(dst, a.texto)} marca(s)")
        feitos.append(dst)
    exportar_pdf_verificacao(feitos, trab)
    a.destino.mkdir(parents=True, exist_ok=True)
    for p in feitos:
        shutil.copy2(p, a.destino / p.name)
    print("Destino:", a.destino)
    print("Conferência visual (PDFs .verif):", trab)


if __name__ == "__main__":
    main()

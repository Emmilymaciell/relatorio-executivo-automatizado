# -*- coding: utf-8 -*-
"""
capa.py — data da CAPA do Relatório de Produtos Mensal do G1 (DEC-058).

A capa vem da CG como um OBJETO PowerPoint embutido no DOCX (word/embeddings/*.pptx) com uma imagem de
pré-visualização (EMF) que é o que o Word imprime. Regra permanente: a data da capa = ÚLTIMO DIA DO MÊS DA
COMPETÊNCIA (cfg.data_capa()), nunca a data real de elaboração.

- ler_data_capa(docx)              -> 'dd de mês de aaaa' encontrada no PPTX embutido (ou None)
- corrigir_data_capa(docx, nova)   -> troca a data no PPTX embutido e, se o PowerPoint estiver disponível,
                                      regenera a pré-visualização (EMF) a partir do slide corrigido.
O 09 chama garantir_data_capa() na cópia de trabalho; o 13 confere no DOCX final.
"""
import io, os, re, shutil, tempfile, zipfile

RX_DATA = re.compile(r"\b(\d{1,2}) de ([a-zç]+) de (\d{4})\b", re.I)


def _pptx_embutidos(z):
    return [n for n in z.namelist() if n.startswith("word/embeddings/") and n.lower().endswith(".pptx")]


def _textos_pptx(blob):
    from pptx import Presentation
    p = Presentation(io.BytesIO(blob))
    for s in p.slides:
        for sh in s.shapes:
            if sh.has_text_frame:
                yield sh.text_frame.text


def ler_data_capa(docx):
    with zipfile.ZipFile(docx) as z:
        for n in _pptx_embutidos(z):
            for t in _textos_pptx(z.read(n)):
                m = RX_DATA.search(t)
                if m:
                    return m.group(0)
    return None


def _trocar_no_pptx(blob, nova):
    """Troca a data (run a run, preservando a formatação) em todas as caixas do PPTX. Devolve (bytes, nº trocas)."""
    from pptx import Presentation
    p = Presentation(io.BytesIO(blob))
    n = 0
    for s in p.slides:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                txt = "".join(r.text for r in para.runs)
                if not RX_DATA.search(txt) or not para.runs:
                    continue
                novo = RX_DATA.sub(nova, txt)
                para.runs[0].text = novo
                for r in para.runs[1:]:
                    r.text = ""
                n += 1
    out = io.BytesIO()
    p.save(out)
    return out.getvalue(), n


def _preview_emf(pptx_path, emf_path):
    """Exporta o slide 1 do PPTX corrigido como EMF (PowerPoint COM). Levanta exceção se não houver Office."""
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(os.path.abspath(pptx_path), True, False, False)
        try:
            pres.Slides(1).Export(os.path.abspath(emf_path), "EMF")
        finally:
            pres.Close()
    finally:
        app.Quit()


def corrigir_data_capa(docx, nova, regenerar_preview=True):
    """Reescreve o DOCX no lugar com a data `nova` ('31 de outubro de 2026') na capa embutida.
    Devolve dict com o que foi feito. A pré-visualização (EMF referenciada pelo v:imagedata da capa) é
    regenerada pelo PowerPoint; se não for possível, 'preview_regenerada' = False (o 13 reprova a capa)."""
    res = {"trocas": 0, "preview_regenerada": False}
    with zipfile.ZipFile(docx) as z:
        itens = {n: z.read(n) for n in z.namelist()}
        infos = {i.filename: i for i in z.infolist()}
    emb = [n for n in itens if n.startswith("word/embeddings/") and n.lower().endswith(".pptx")]
    preview = None
    for n in emb:
        novo, k = _trocar_no_pptx(itens[n], nova)
        if k:
            itens[n] = novo
            res["trocas"] += k
            rels = itens.get("word/_rels/document.xml.rels", b"").decode("utf-8")
            alvo = os.path.relpath(n, "word").replace("\\", "/")
            m = re.search(r'Id="([^"]+)"[^>]*Target="' + re.escape(alvo) + '"', rels) or \
                re.search(r'Target="' + re.escape(alvo) + r'"[^>]*Id="([^"]+)"', rels)
            if m:
                doc = itens["word/document.xml"].decode("utf-8")
                o = re.search(r'<w:object[^>]*>(?:(?!</w:object>).)*?r:id="' + m.group(1) + '"', doc, re.S)
                if o:
                    im = re.search(r'<v:imagedata r:id="([^"]+)"', o.group(0))
                    if im:
                        t = re.search(r'Id="' + im.group(1) + r'"[^>]*Target="([^"]+)"', rels) or \
                            re.search(r'Target="([^"]+)"[^>]*Id="' + im.group(1) + '"', rels)
                        if t and t.group(1).lower().endswith(".emf"):
                            preview = ("word/" + t.group(1), novo)
    if not res["trocas"]:
        return res
    if regenerar_preview and preview:
        with tempfile.TemporaryDirectory() as tmp:
            pp, ee = os.path.join(tmp, "capa.pptx"), os.path.join(tmp, "capa.emf")
            open(pp, "wb").write(preview[1])
            try:
                _preview_emf(pp, ee)
                itens[preview[0]] = open(ee, "rb").read()
                res["preview_regenerada"] = True
            except Exception as e:  # sem Office: o texto foi corrigido, a imagem impressa não
                res["erro_preview"] = str(e)
    tmpdoc = docx + ".tmp"
    with zipfile.ZipFile(tmpdoc, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in itens.items():
            z.writestr(infos[n], b)
    shutil.move(tmpdoc, docx)
    return res


def garantir_data_capa(docx, data_ext):
    """Confere a capa; se a data for outra, corrige. Devolve (data_lida_antes, resultado|None)."""
    antes = ler_data_capa(docx)
    if antes is not None and antes.lower() == data_ext.lower():
        return antes, None
    return antes, corrigir_data_capa(docx, data_ext)

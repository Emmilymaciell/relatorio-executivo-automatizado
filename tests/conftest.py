# -*- coding: utf-8 -*-
"""Fixtures ARTIFICIAIS (nenhum documento real do Programa) para a pipeline da Coordenação.

Monta uma pasta mensal '15. Setembro 2026' num diretório temporário que imita o SharePoint local:
  <tmp>/Programa/Programa - 3. Desenvolvimento/...   (raiz para o arquivo interno da publicação)
  <tmp>/Programa/Programa - 4. Produtos/15. Setembro 2026/{Índice, G1..G5}
"""
import datetime
import json
import os
import shutil
import sys
import zipfile

import pytest

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(RAIZ, "scripts"))
sys.path.insert(0, os.path.join(RAIZ, ".claude", "skills", "relatorio-prog-grupo1", "scripts"))

import fitz  # noqa: E402

PRODUTOS_TESTE = {"G1": ["P1.1", "P1.2", "P1.3", "P1.4"], "G2": ["P2.1", "P2.2", "P2.4"],
                  "G3": ["P3.1", "P3.2", "P3.3"], "G4": ["P4.1", "P4.2", "P4.3"],
                  "G5": ["P5.1", "P5.2", "P5.3", "P5.4"]}


# ------------------------------------------------------------------------------------ PDFs
def make_pdf(path, texto="Documento de teste", paginas=1, imagem_ruido_kb=0):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = fitz.open()
    for i in range(paginas):
        pg = doc.new_page()
        pg.insert_text((72, 72), f"{texto} — página {i + 1}", fontsize=12)
        if imagem_ruido_kb:
            pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 400, 400), False)
            pix.set_rect(pix.irect, (0, 0, 0))
            raw = os.urandom(400 * 400 * 3)
            pix = fitz.Pixmap(fitz.csRGB, 400, 400, raw, False)
            pg.insert_image(fitz.Rect(72, 100, 472, 500), pixmap=pix)
    doc.save(path)
    doc.close()
    return path


def make_pdf_sem_texto(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = fitz.open()
    pg = doc.new_page()
    pix = fitz.Pixmap(fitz.csRGB, 50, 50, bytes([200]) * (50 * 50 * 3), False)
    pg.insert_image(fitz.Rect(72, 72, 300, 300), pixmap=pix)
    doc.save(path)
    doc.close()
    return path


def make_pdf_protegido(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = fitz.open()
    doc.new_page().insert_text((72, 72), "protegido")
    doc.save(path, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw="dono", user_pw="senha")
    doc.close()
    return path


def make_pdf_restrito(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = fitz.open()
    doc.new_page().insert_text((72, 72), "restrito")
    doc.save(path, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw="dono", user_pw="",
             permissions=fitz.PDF_PERM_ACCESSIBILITY)
    doc.close()
    return path


_CERT = {}


def _cert():
    if _CERT:
        return _CERT["cert"], _CERT["key"]
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    nome = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "SIGNATARIO DE TESTE"),
                      x509.NameAttribute(NameOID.ORGANIZATION_NAME, "AC TESTE (NAO E ICP-BRASIL)")])
    agora = datetime.datetime(2026, 1, 1)
    cert = (x509.CertificateBuilder().subject_name(nome).issuer_name(nome).public_key(key.public_key())
            .serial_number(4242).not_valid_before(agora).not_valid_after(agora + datetime.timedelta(days=3650))
            .sign(key, hashes.SHA256()))
    _CERT.update(cert=cert, key=key)
    return cert, key


def make_pdf_assinado(path, texto="Relatório de Produtos Mensal assinado"):
    """PDF com assinatura PAdES (adbe.pkcs7.detached) REAL, gerada com certificado de teste."""
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.serialization import pkcs7
    os.makedirs(os.path.dirname(path), exist_ok=True)
    conteudo = f"BT /F1 12 Tf 72 720 Td ({texto.encode('latin-1', 'replace').decode('latin-1')}) Tj ET".encode("latin-1")
    hexlen = 16384
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] /SigFlags 3 >> >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 6 0 R >> >> /Annots [5 0 R] >>",
        b"<< /Length " + str(len(conteudo)).encode() + b" >>\nstream\n" + conteudo + b"\nendstream",
        b"<< /Type /Annot /Subtype /Widget /FT /Sig /Rect [0 0 0 0] /T (Signature1) /V 7 0 R /P 3 0 R /F 132 >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached "
        b"/ByteRange [0 0000000000 0000000000 0000000000] /Contents <" + b"0" * hexlen + b"> >>",
    ]
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + o + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for o in offs:
        out += f"{o:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    raw = bytes(out)
    ini = raw.index(b"/Contents <") + len(b"/Contents ")
    fim = ini + hexlen + 2
    br = [0, ini, fim, len(raw) - fim]
    br_txt = f"[0 {br[1]:010d} {br[2]:010d} {br[3]:010d}]".encode()
    raw = raw.replace(b"[0 0000000000 0000000000 0000000000]", br_txt)
    dados = raw[:ini] + raw[fim:]
    cert, key = _cert()
    sig = (pkcs7.PKCS7SignatureBuilder().set_data(dados).add_signer(cert, key, hashes.SHA256())
           .sign(serialization.Encoding.DER, [pkcs7.PKCS7Options.DetachedSignature, pkcs7.PKCS7Options.Binary]))
    h = sig.hex().encode()
    assert len(h) <= hexlen
    raw = raw[:ini + 1] + h + b"0" * (hexlen - len(h)) + raw[fim - 1:]
    with open(path, "wb") as f:
        f.write(raw)
    return path


def adulterar(path):
    raw = bytearray(open(path, "rb").read())
    i = raw.index(b"Tj ET") - 3
    raw[i] = ord("X") if raw[i] != ord("X") else ord("Y")
    open(path, "wb").write(bytes(raw))


def make_pdf_visual(path):
    return make_pdf(path, "Documento assinado digitalmente via gov.br por FULANO")


# ------------------------------------------------------------------------------- Office/ZIP
def make_xlsx(path, formulas=True):
    import openpyxl
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["A2"] = 10, 20
    ws["A3"] = "=SUM(A1:A2)" if formulas else 30
    wb.save(path)
    return path


def make_docx(path, texto="minuta"):
    import docx
    os.makedirs(os.path.dirname(path), exist_ok=True)
    d = docx.Document()
    d.add_paragraph(texto)
    d.save(path)
    return path


def make_zip(path, membros):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in membros.items():
            z.writestr(n, data)
    return path


def make_lista(path):
    import openpyxl
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = openpyxl.Workbook()
    wb.active.append(["Produto", "Documento"])
    wb.save(path)
    return path


def nome_anexo(cod, desc, rod="GERAL", ext=".pdf"):
    import prog_entrega.regras as R
    return f"{cod} - {R.produtos()[cod]['nome']} - {desc} - {rod}{ext}"


# --------------------------------------------------------------------------- pasta mensal
@pytest.fixture
def ambiente(tmp_path, monkeypatch):
    raiz = tmp_path / "Programa"
    prod = raiz / "Programa - 4. Produtos"
    (raiz / "Programa - 3. Desenvolvimento" / "G1 - Coordenacao" / "Desenvolvimento de Produtos").mkdir(parents=True)
    prod.mkdir(parents=True)
    monkeypatch.setenv("PROG_ROOT", str(raiz))
    monkeypatch.setenv("PROG_ENTREGA_OUT", str(tmp_path / "out"))
    monkeypatch.delenv("PROG_ENTREGA_CONFIG", raising=False)
    return {"raiz": raiz, "produtos": prod, "tmp": tmp_path}


@pytest.fixture
def config_limite_pequeno(tmp_path, monkeypatch):
    """Config com limite de 1 MB (testa fracionamento sem gerar arquivos de 100 MB)."""
    import prog_entrega.regras as R
    d = json.load(open(os.path.join(RAIZ, "config", "entrega_pop.json"), encoding="utf-8"))
    d["limite_tamanho_mb"] = 1
    d["indice"]["modelo"] = os.path.join(RAIZ, d["indice"]["modelo"])
    d["feriados_arquivo"] = os.path.join(RAIZ, d["feriados_arquivo"])
    p = tmp_path / "entrega_pop_teste.json"
    p.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setenv("PROG_ENTREGA_CONFIG", str(p))
    R._CFG = None
    yield
    R._CFG = None


def montar_mes_correto(prod_dir, nome="15. Setembro 2026"):
    """Pasta mensal conforme o POP: em cada grupo, relatório assinado + Lista + ANEXOS com nomes padrão."""
    from prog_entrega import indice as IDX
    m = os.path.join(str(prod_dir), nome)
    for g, cods in PRODUTOS_TESTE.items():
        gd = os.path.join(m, g)
        make_pdf_assinado(os.path.join(gd, f"Programa - Relatório Grupo {g[1]} - Relatório de Produtos Mensal N°15 - 09 2026.pdf"))
        make_lista(os.path.join(gd, f"PROG_Lista de demandas {g}.xlsx"))
        for c in cods:
            make_pdf(os.path.join(gd, "ANEXOS", nome_anexo(c, f"Documento de {c}", "RD-101")),
                     f"anexo {c}")
    IDX.criar_modelo_mes(os.path.join(m, IDX.nome_arquivo_indice("Setembro", 2026)), "Setembro", 2026,
                         periodo=("01/09/2026", "30/09/2026"))
    return m

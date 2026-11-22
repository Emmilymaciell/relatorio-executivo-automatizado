# -*- coding: utf-8 -*-
"""Validação de integridade por tipo REAL de arquivo (assinatura de bytes, não só extensão).

Todas as funções são SOMENTE LEITURA: abrem os arquivos sem gravar (planilhas são lidas como ZIP/XML,
nunca salvas por biblioteca — fórmulas e memórias de cálculo não são tocadas).
Retorno: {"tipo", "problemas": [[status, mensagem]], "info": {...}}.
"""
import os
import re
import zipfile

from . import regras as R

VERSAO_VALIDADOR = "arquivos-1"
_OOXML = {".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm"}


def tipo_real(path):
    with open(path, "rb") as f:
        head = f.read(8)
    # assinatura ZIP primeiro: um ZIP cujo 1º membro é um PDF "stored" traz '%PDF-' em claro no 1º KB
    if head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06"):
        ext = os.path.splitext(path)[1].lower()
        return "ooxml" if ext in _OOXML else "zip"
    if head.startswith(b"%PDF-") or b"%PDF-" in _ler_inicio(path):
        return "pdf"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "ole"  # .doc/.xls/.ppt legados
    return "outro"


def _ler_inicio(path, n=1024):
    with open(path, "rb") as f:
        return f.read(n)


def validar(path):
    """Despacha pelo tipo real. Também confere coerência extensão × conteúdo."""
    out = {"tipo": None, "problemas": [], "info": {}}
    try:
        tam = os.path.getsize(path)
    except OSError as e:
        out["problemas"].append([R.BLOCKER, f"arquivo inacessível: {e}"])
        return out
    out["info"]["tamanho"] = tam
    if tam == 0:
        out["tipo"] = "vazio"
        out["problemas"].append([R.BLOCKER, "arquivo vazio (0 bytes)"])
        return out
    t = tipo_real(path)
    out["tipo"] = t
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf" and t != "pdf":
        out["problemas"].append([R.BLOCKER, f"extensão .pdf mas o conteúdo não é PDF (tipo real: {t})"])
        return out
    if ext == ".zip" and t != "zip":
        out["problemas"].append([R.BLOCKER, f"extensão .zip mas o conteúdo não é ZIP (tipo real: {t})"])
        return out
    if t == "pdf":
        _pdf(path, out)
    elif t == "zip":
        _zip(path, out)
    elif t == "ooxml":
        _ooxml(path, ext, out)
    elif t == "ole":
        out["problemas"].append([R.WARNING, f"formato Office legado ({ext}) — integridade só superficial; "
                                            "preferir formato atual (.xlsx/.docx) dentro do ZIP"])
    return out


# --------------------------------------------------------------------------------------------- PDF
def _pdf(path, out):
    import fitz
    try:
        doc = fitz.open(path)
    except Exception as e:
        out["problemas"].append([R.BLOCKER, f"PDF corrompido / não abre: {e}"])
        return
    info = out["info"]
    try:
        if doc.needs_pass:
            out["problemas"].append([R.BLOCKER, "PDF protegido por SENHA de abertura"])
            return
        info["paginas"] = doc.page_count
        info["criptografado"] = bool(doc.is_encrypted or (doc.metadata or {}).get("encryption"))
        if info["criptografado"]:
            perm = doc.permissions
            falta = [n for n, bit in (("imprimir", fitz.PDF_PERM_PRINT), ("copiar", fitz.PDF_PERM_COPY))
                     if not perm & bit]
            if falta:
                out["problemas"].append([R.BLOCKER, f"PDF com RESTRIÇÃO de uso (sem permissão para {', '.join(falta)})"])
            else:
                out["problemas"].append([R.WARNING, "PDF criptografado (sem restrição de impressão/cópia)"])
        if getattr(doc, "is_repaired", False):
            out["problemas"].append([R.WARNING, "estrutura do PDF danificada — foi reparada ao abrir"])
        if doc.page_count == 0:
            out["problemas"].append([R.BLOCKER, "PDF sem páginas"])
            return
        com_texto, falhas = 0, []
        for i in range(doc.page_count):  # uma página por vez (não carrega o documento inteiro)
            try:
                if doc.load_page(i).get_text("text").strip():
                    com_texto += 1
            except Exception as e:
                falhas.append(f"p.{i + 1}: {e}")
        info["paginas_com_texto"] = com_texto
        if falhas:
            out["problemas"].append([R.BLOCKER, f"{len(falhas)} página(s) ilegível(is): {falhas[:3]}"])
        if com_texto == 0:
            out["problemas"].append([R.MANUAL, "PDF sem texto extraível (provável digitalização) — POP exige PDF "
                                               "pesquisável; confirmar/OCR"])
        elif com_texto < doc.page_count / 2:
            out["problemas"].append([R.WARNING, f"só {com_texto}/{doc.page_count} páginas com texto pesquisável"])
    finally:
        doc.close()


# --------------------------------------------------------------------------------------------- ZIP
def _zip(path, out):
    info = out["info"]
    try:
        with zipfile.ZipFile(path) as z:
            membros = [m for m in z.infolist() if not m.is_dir()]
            info["membros"] = [m.filename for m in membros]
            info["tamanho_descompactado"] = sum(m.file_size for m in membros)
            if not membros:
                out["problemas"].append([R.BLOCKER, "ZIP vazio"])
                return
            if any(m.flag_bits & 0x1 for m in membros):
                out["problemas"].append([R.BLOCKER, "ZIP com SENHA (membros criptografados)"])
                return
            ruins = [m.filename for m in membros if m.filename.startswith(("/", "\\")) or ".." in m.filename.replace("\\", "/").split("/")]
            if ruins:
                out["problemas"].append([R.BLOCKER, f"ZIP com caminhos inseguros: {ruins[:3]}"])
            bad = z.testzip()  # lê tudo e confere CRC
            if bad:
                out["problemas"].append([R.BLOCKER, f"ZIP corrompido (CRC inválido em '{bad}')"])
            aninhados = [m.filename for m in membros if m.filename.lower().endswith(".zip")]
            if aninhados:
                out["problemas"].append([R.WARNING, f"ZIP dentro de ZIP: {aninhados[:3]}"])
    except zipfile.BadZipFile as e:
        out["problemas"].append([R.BLOCKER, f"ZIP corrompido / não abre: {e}"])
    except Exception as e:
        out["problemas"].append([R.BLOCKER, f"falha ao testar ZIP: {e}"])


# ------------------------------------------------------------------------------- Office (OOXML)
_RE_F = re.compile(rb"<(?:\w+:)?f[ >/]")


def _ooxml(path, ext, out):
    info = out["info"]
    try:
        with zipfile.ZipFile(path) as z:
            nomes = z.namelist()
            if "[Content_Types].xml" not in nomes:
                out["problemas"].append([R.BLOCKER, "arquivo Office corrompido (sem [Content_Types].xml)"])
                return
            bad = z.testzip()
            if bad:
                out["problemas"].append([R.BLOCKER, f"arquivo Office corrompido (CRC em '{bad}')"])
                return
            if ext in (".xlsx", ".xlsm"):
                abas = [n for n in nomes if n.startswith("xl/worksheets/") and n.endswith(".xml")]
                formulas = 0
                for n in abas:
                    formulas += len(_RE_F.findall(z.read(n)))
                info["abas"] = len(abas)
                info["formulas"] = formulas
                if not abas:
                    out["problemas"].append([R.BLOCKER, "planilha sem abas legíveis"])
    except zipfile.BadZipFile as e:
        out["problemas"].append([R.BLOCKER, f"arquivo Office corrompido / não abre: {e}"])

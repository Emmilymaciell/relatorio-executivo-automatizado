# -*- coding: utf-8 -*-
"""Tradução caminho LOCAL (pasta sincronizada pelo OneDrive) <-> caminho REMOTO (dentro da biblioteca) e
bloqueios de escrita. Funções puras (testáveis sem rede)."""
import os
import unicodedata
from urllib.parse import quote


class CaminhoInvalido(Exception):
    pass


def _norm(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn").lower()


def normalizar_remoto(remoto):
    """'\\a\\b/' -> 'a/b' (sem barras nas pontas). Recusa '..'."""
    partes = [p for p in str(remoto).replace("\\", "/").split("/") if p not in ("", ".")]
    if any(p == ".." for p in partes):
        raise CaminhoInvalido(f"caminho remoto com '..': {remoto}")
    return "/".join(partes)


def remoto_de_local(local, prog_root, mapa, prefixo="Programa - "):
    """Ex.: <prog_root>\\Programa - 3. Desenvolvimento\\G1\\x.xlsx -> '3. Desenvolvimento/G1/x.xlsx'."""
    loc = os.path.abspath(local)
    raiz = os.path.abspath(prog_root)
    try:
        rel = os.path.relpath(loc, raiz)
    except ValueError:
        rel = ".."
    if rel == "." or rel.startswith(".."):
        raise CaminhoInvalido(f"o caminho não está sob a raiz sincronizada do Programa ({raiz}): {local}")
    partes = rel.split(os.sep)
    topo = partes[0]
    alvo = None
    for k, v in (mapa or {}).items():
        if _norm(k) == _norm(topo):
            alvo = v
            break
    if alvo is None:
        if prefixo and _norm(topo).startswith(_norm(prefixo)):
            alvo = topo[len(prefixo):]
        else:
            raise CaminhoInvalido(f"pasta local '{topo}' sem correspondência na biblioteca (config/sharepoint.json)")
    return normalizar_remoto("/".join([alvo] + partes[1:]))


def local_de_remoto(remoto, prog_root, mapa, prefixo="Programa - "):
    partes = normalizar_remoto(remoto).split("/")
    topo = partes[0]
    pasta = None
    for k, v in (mapa or {}).items():
        if _norm(v) == _norm(topo):
            pasta = k
            break
    if pasta is None:
        pasta = f"{prefixo}{topo}"
    return os.path.join(os.path.abspath(prog_root), pasta, *partes[1:])


def bloqueio_escrita(remoto, bloqueados):
    """Trecho bloqueado encontrado no caminho (ou None)."""
    n = _norm(normalizar_remoto(remoto))
    for b in bloqueados or ():
        if _norm(b) in n:
            return b
    return None


def exigir_escrita_permitida(remoto, bloqueados):
    b = bloqueio_escrita(remoto, bloqueados)
    if b:
        raise CaminhoInvalido(f"escrita PROIBIDA em caminho que contém '{b}': {remoto}")
    if not normalizar_remoto(remoto):
        raise CaminhoInvalido("escrita na raiz da biblioteca não é permitida")


def codificar(remoto):
    """Codifica cada segmento para a sintaxe /root:/{caminho}: da Graph."""
    return "/".join(quote(p, safe="") for p in normalizar_remoto(remoto).split("/"))

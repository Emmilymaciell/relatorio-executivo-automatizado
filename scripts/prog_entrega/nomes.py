# -*- coding: utf-8 -*-
"""Nomenclatura dos anexos (POP Quadro 3):

    [CÓDIGO DO PRODUTO] - [NOME DO PRODUTO] - [NOME DO ARQUIVO] - [RODOVIA]
    ex.: P2.2 - Análise de Projeto e RPFO - Nota Técnica de Aceitação de Documentos - NT-749.1 - RD-101

Separador = " - " (hífen com espaços). O NOME DO ARQUIVO pode conter " - " internamente (como no exemplo
do POP); por isso: 1º bloco = código, 2º = produto, último = rodovia, o meio = nome do arquivo.
Fragmentos: "Parte NN" (dois dígitos) em qualquer bloco do nome do arquivo (ex.: "... - Parte 01 - RD-101").
"""
import os
import re

from . import regras as R

SEP = " - "
_RE_COD = re.compile(r"^P\d\.\d+$")
_RE_COD_SOLTO = re.compile(r"\bP\s?(\d)\s?[.,_]\s?(\d+)\b", re.I)
_RE_PARTE = re.compile(r"\bParte\s+(\d{1,3})\b", re.I)
_RE_ROD = re.compile(r"(?<![A-Za-z0-9])RD[\s_\-]?(\d{3})\s?([A-Z])?(?![A-Za-z0-9])", re.I)
_RE_NUM_INICIAL = re.compile(r"^\s*\d{1,3}\s*[.)\-]\s+")


def canon_rodovia(txt):
    t = (txt or "").strip()
    if R.norm(t) == "geral":
        return "GERAL"
    m = _RE_ROD.fullmatch(t.replace(" ", "")) or _RE_ROD.fullmatch(t)
    if not m:
        return None
    cand = f"RD-{m.group(1)}{(m.group(2) or '').upper()}"
    return cand if cand in R.rodovias() else None


def rodovias_mencionadas(txt):
    achadas = set()
    for m in _RE_ROD.finditer(txt or ""):
        cand = f"RD-{m.group(1)}{(m.group(2) or '').upper()}"
        if cand in R.rodovias():
            achadas.add(cand)
    return achadas


def parte(nome):
    m = _RE_PARTE.search(nome)
    return int(m.group(1)) if m else None


def base_sem_parte(stem):
    return re.sub(r"\s*-?\s*\bParte\s+\d{1,3}\b\s*", " ", stem, flags=re.I).strip(" -")


def analisar(nome_arquivo):
    """Analisa o nome contra o padrão. Retorna dict com partes reconhecidas e lista de problemas
    [[status, msg]]. Não decide nada sobre o conteúdo."""
    stem, ext = os.path.splitext(nome_arquivo)
    probs = []
    blocos = [b.strip() for b in stem.split(SEP)]
    out = {"stem": stem, "ext": ext.lower(), "codigo": None, "produto": None, "descricao": None,
           "rodovia": None, "parte": parte(stem), "conforme": False, "problemas": probs}
    if "_" in stem and SEP not in stem:
        probs.append([R.WARNING, "usa '_' como separador (padrão é ' - ')"])
    if len(blocos) < 4:
        probs.append([R.BLOCKER, "nome fora do padrão [Código] - [Produto] - [Nome do arquivo] - [Rodovia] "
                                 f"({len(blocos)} bloco(s) separados por ' - ')"])
    cod = blocos[0] if blocos else ""
    if _RE_COD.match(cod):
        out["codigo"] = cod
        if cod not in R.produtos():
            probs.append([R.BLOCKER, f"código de produto inexistente: {cod}"])
    else:
        probs.append([R.BLOCKER, "não começa pelo código do produto (ex.: 'P2.2')"])
    if len(blocos) >= 4:
        out["produto"] = blocos[1]
        out["descricao"] = SEP.join(blocos[2:-1])
        rod = canon_rodovia(blocos[-1])
        if rod and blocos[-1] == rod:
            out["rodovia"] = rod
        elif rod:
            out["rodovia"] = rod
            probs.append([R.WARNING, f"rodovia '{blocos[-1]}' fora da grafia padrão '{rod}'"])
        else:
            probs.append([R.BLOCKER, f"último bloco '{blocos[-1]}' não é rodovia válida "
                                     f"({' | '.join(R.rodovias())})"])
        if out["codigo"] in R.produtos():
            p = R.produtos()[out["codigo"]]
            aceitos = {R.norm(p["nome"])} | {R.norm(a) for a in p.get("aliases", [])}
            if R.norm(out["produto"]) not in aceitos:
                probs.append([R.WARNING, f"nome do produto '{out['produto']}' difere do catálogo "
                                         f"('{p['nome']}')"])
    n = R.norm(stem)
    for termo in R.config()["termos_genericos"]:
        if re.search(rf"(^|[\s\-_(]){re.escape(R.norm(termo))}($|[\s\-_)])", n):
            probs.append([R.MANUAL, f"termo genérico/provisório no nome: '{termo}' — confirmar se é versão final"])
            break
    if re.search(r"\(\d+\)$", stem) or re.search(r"\bv\d+\b", n):
        probs.append([R.WARNING, "marcador de versão/cópia no nome (ex.: '(1)', 'v2')"])
    lim = R.config().get("nome_max_caracteres")
    if lim and len(nome_arquivo) > lim:
        probs.append([R.WARNING, f"nome com {len(nome_arquivo)} caracteres (> {lim}, heurística)"])
    out["conforme"] = not any(s == R.BLOCKER for s, _ in probs)
    return out


def _limpar_descricao(txt, codigo=None):
    t = _RE_NUM_INICIAL.sub("", txt)
    t = t.replace("_", " ")
    t = re.sub(r"\s+-\s+|\s*–\s*", " - ", t)
    if codigo:
        t = re.sub(rf"^\s*{re.escape(codigo)}\s*[-.]?\s*", "", t)
        t = re.sub(rf"^\s*{re.escape(codigo.replace('.', ''))}(?![0-9])\s*-?\s*", "", t)  # 'P33' = P3.3
    t = _RE_ROD.sub("", t)
    t = re.sub(r"\(\s*\)", "", t)
    t = re.sub(r"\s+", " ", t).strip(" -.")
    t = re.sub(r"(\s*-\s*){2,}", " - ", t)
    return t.strip(" -")


def propor(nome_arquivo, produto_pasta=None, contexto=""):
    """Propõe nome no padrão a partir do nome atual + produto da subpasta de origem (se houver) +
    rodovias citadas no nome/caminho. Retorna (nome_proposto | None, motivo, completo: bool)."""
    a = analisar(nome_arquivo)
    stem, ext = a["stem"], a["ext"]
    if a["conforme"] and not any(s == R.WARNING and "rodovia" in m for s, m in a["problemas"]):
        return nome_arquivo, "já conforme", True
    cod = a["codigo"] if a["codigo"] in R.produtos() else None
    if not cod and produto_pasta in R.produtos():
        cod = produto_pasta
    if not cod:
        m = _RE_COD_SOLTO.search(stem)
        if m and f"P{m.group(1)}.{m.group(2)}" in R.produtos():
            cod = f"P{m.group(1)}.{m.group(2)}"
    if not cod:
        return None, "código do produto não determinável pelo nome nem pela pasta", False
    nome_prod = R.produtos()[cod]["nome"]
    if a["conforme"] or (a["codigo"] == cod and a["descricao"]):
        desc = a["descricao"]
    else:
        desc = _limpar_descricao(stem, cod)
        for al in [nome_prod] + R.produtos()[cod].get("aliases", []):
            desc = re.sub(rf"^\s*{re.escape(al)}\s*-\s*", "", desc, flags=re.I)
    if not desc:
        return None, "descrição do arquivo vazia após limpeza", False
    rods = rodovias_mencionadas(stem) or rodovias_mencionadas(contexto)
    if a["rodovia"]:
        rod = a["rodovia"]
    elif len(rods) == 1:
        rod = rods.pop()
    elif len(rods) > 1:
        rod = "GERAL"  # POP: GERAL quando abranger mais de uma rodovia
    else:
        return f"{cod} - {nome_prod} - {desc} - [RODOVIA?]{ext}", "rodovia não identificável pelo nome", False
    novo = f"{cod} - {nome_prod} - {desc} - {rod}{ext}"
    novo = re.sub(r"\s+", " ", novo)
    return novo, "proposta derivada da pasta do produto + nome + rodovia citada", analisar(novo)["conforme"]

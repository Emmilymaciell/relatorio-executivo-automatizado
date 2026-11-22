# -*- coding: utf-8 -*-
"""Inventário da pasta mensal: cada arquivo recebe local (raiz do mês / raiz do grupo / ANEXOS / subpasta
de ANEXOS / outra pasta), classe (índice, relatório PDF, relatório editável, Lista de Demandas, anexo,
temporário, lixo de sistema, fora da estrutura), tamanho, mtime e SHA-256. Somente leitura."""
import os
import re

from . import regras as R

_RE_GRUPO = re.compile(r"^G([1-5])$", re.I)


def _produto_de_pasta(partes):
    for p in partes:
        cod = p.split(" ")[0].strip()
        if cod in R.produtos():
            return cod
    return None


def classificar(rel_partes, nome):
    """(grupo, local, classe, subpasta_anexos, produto_pasta)."""
    cfg = R.config()
    if R.eh_lixo(nome):
        classe_forcada = R.K_LIXO
    elif R.eh_temporario(nome):
        classe_forcada = R.K_TEMP
    else:
        classe_forcada = None
    pastas = rel_partes[:-1]
    ext = os.path.splitext(nome)[1].lower()
    if not pastas:
        classe = R.K_INDICE if (R.casa_algum(nome, cfg["padroes_indice"]) and ext in (".docx", ".pdf")) else R.K_FORA
        return None, "raiz_mes", classe_forcada or classe, None, None
    m = _RE_GRUPO.match(pastas[0].strip())
    if not m:
        return None, "outra_pasta", classe_forcada or R.K_FORA, None, None
    grupo = f"G{m.group(1)}"
    if len(pastas) == 1:
        if R.casa_algum(nome, cfg["padroes_lista_demandas"]):
            classe = R.K_LISTA
        elif R.casa_algum(nome, cfg["padroes_relatorio"]) and ext == ".pdf":
            classe = R.K_RELATORIO
        elif R.casa_algum(nome, cfg["padroes_relatorio"]) and ext in (".doc", ".docx", ".docm", ".odt"):
            classe = R.K_RELATORIO_EDITAVEL
        else:
            classe = R.K_FORA
        return grupo, "raiz_grupo", classe_forcada or classe, None, None
    if R.norm(pastas[1]) == "anexos":
        sub = pastas[2] if len(pastas) > 2 else None
        classe = R.K_LISTA if R.casa_algum(nome, cfg["padroes_lista_demandas"]) else R.K_ANEXO
        return (grupo, "anexos_sub" if sub else "anexos", classe_forcada or classe, sub,
                _produto_de_pasta(pastas[2:]))
    return grupo, "outra_pasta", classe_forcada or R.K_FORA, None, None


def inventariar(pasta_mes, cache):
    itens, dirs_anexos = [], {}
    pasta_mes = R.lp(pasta_mes)
    for raiz, dirs, files in os.walk(pasta_mes):
        dirs.sort()
        rel_dir = os.path.relpath(raiz, pasta_mes)
        partes_dir = [] if rel_dir == "." else rel_dir.split(os.sep)
        if len(partes_dir) == 2 and _RE_GRUPO.match(partes_dir[0]) and R.norm(partes_dir[1]) == "anexos":
            dirs_anexos.setdefault(f"G{partes_dir[0][1]}", []).append(partes_dir[1])
        for f in sorted(files):
            partes = partes_dir + [f]
            abs_ = os.path.join(raiz, f)
            grupo, local, classe, sub, prod = classificar(partes, f)
            st = os.stat(abs_)
            it = {"rel": "/".join(partes), "abs": abs_, "nome": f, "ext": os.path.splitext(f)[1].lower(),
                  "grupo": grupo, "local": local, "classe": classe, "subpasta_anexos": sub,
                  "produto_pasta": prod, "tamanho": st.st_size, "mtime": int(st.st_mtime)}
            if classe not in (R.K_LIXO,):
                it["sha256"] = cache.sha256(abs_)
            itens.append(it)
    grupos_presentes = sorted(g for g in R.config()["grupos"] if os.path.isdir(os.path.join(pasta_mes, g)))
    return {"pasta": R.sem_lp(pasta_mes), "itens": itens, "pastas_anexos": dirs_anexos, "grupos_presentes": grupos_presentes}

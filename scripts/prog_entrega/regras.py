# -*- coding: utf-8 -*-
"""Regras e configuração da entrega mensal ao Parceiro (POP v2026-10-01).

Fonte das regras variáveis: config/entrega_pop.json (limite de tamanho, rodovias, produtos, padrões
de arquivos internos...). Regras documentais estáveis estão aqui e em docs/dominio/. A autoridade
de cada regra está em docs/referencias-oficiais/README.md.
"""
import json
import os
import re
import unicodedata

VERSAO_AUTOMACAO = "prog_entrega 1.0.0"

# --- Classificação de cada verificação (não usar booleanos) -------------------------------------
PASS = "PASS"
FIXED = "FIXED_AUTOMATICALLY"
WARNING = "WARNING"
MANUAL = "MANUAL_REVIEW"
BLOCKER = "BLOCKER"
NA = "N/A"
STATUS_ORDEM = [BLOCKER, MANUAL, WARNING, FIXED, PASS, NA]

# --- Veredito final -----------------------------------------------------------------------------
READY = "READY FOR PARCEIRO"
PENDENTE_HUMANO = "PENDENTE DE REVISÃO HUMANA"
BLOCKED = "BLOQUEADO"

# --- Classificação de arquivos do inventário ----------------------------------------------------
K_INDICE = "indice"
K_RELATORIO = "relatorio_pdf"
K_RELATORIO_EDITAVEL = "relatorio_editavel"
K_LISTA = "lista_demandas"
K_ANEXO = "anexo"
K_TEMP = "temporario"
K_LIXO = "lixo_sistema"
K_FORA = "fora_da_estrutura"


def project_dir():
    env = os.environ.get("PROG_PROJECT_DIR")
    if env and os.path.isdir(env):
        return env
    return os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))


def norm(s):
    """minúsculas, sem acento, espaços colapsados — para comparar nomes."""
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"\s+", " ", s).strip()


_CFG = None


def config(path=None):
    """Carrega config/entrega_pop.json (cache). PROG_ENTREGA_CONFIG sobrepõe o caminho (testes)."""
    global _CFG
    path = path or os.environ.get("PROG_ENTREGA_CONFIG") or os.path.join(project_dir(), "config", "entrega_pop.json")
    if _CFG is None or _CFG.get("_path") != path:
        with open(path, encoding="utf-8") as f:
            _CFG = json.load(f)
        _CFG["_path"] = path
    return _CFG


def limite_bytes():
    return int(config()["limite_tamanho_mb"]) * 1024 * 1024


def produtos():
    return {k: v for k, v in config()["produtos"].items() if not k.startswith("_")}


def produtos_do_grupo(grupo):
    return [k for k, v in produtos().items() if v["grupo"] == grupo]


def rodovias():
    return list(config()["rodovias"])


def lp(p):
    """Caminho estendido do Windows (prefixo de caminho longo) — evita o limite de 260 caracteres
    (as pastas do G2 são profundas)."""
    if os.name != "nt" or not p:
        return p
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def sem_lp(p):
    if p and p.startswith("\\\\?\\UNC\\"):
        return "\\\\" + p[8:]
    return p[4:] if p and p.startswith("\\\\?\\") else p


def resolve_path(rel):
    return rel if os.path.isabs(rel) else os.path.join(project_dir(), rel)


MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
         "Setembro", "Outubro", "Novembro", "Dezembro"]

_RE_PASTA_MES = re.compile(r"^\s*(\d{1,2})\.\s*([A-Za-zçÇãÃéÉ]+)\s+(\d{4})\s*$")


def parse_pasta_mes(nome):
    """'16. Outubro 2026' -> (16, 'Outubro', 10, 2026) ou None."""
    m = _RE_PASTA_MES.match(nome)
    if not m:
        return None
    num, mes, ano = int(m.group(1)), m.group(2), int(m.group(3))
    idx = next((i for i, x in enumerate(MESES) if norm(x) == norm(mes)), None)
    if idx is None:
        return None
    return num, MESES[idx], idx + 1, ano


def periodo_do_mes(mes_num, ano):
    """Competência = mês civil (01 ao último dia) — regra vigente em config/parametros*.json
    (periodo_inicio/periodo_fim) desde o RM13. Ver docs/dominio/processo-entrega-mensal.md §Período."""
    import calendar
    ult = calendar.monthrange(ano, mes_num)[1]
    return f"01/{mes_num:02d}/{ano}", f"{ult:02d}/{mes_num:02d}/{ano}"


def eh_temporario(nome):
    low = nome.lower()
    cfg = config()
    if any(low.startswith(p) for p in cfg["prefixos_temporarios"]):
        return True
    return any(low.endswith(e) for e in cfg["extensoes_temporarias"])


def eh_lixo(nome):
    return nome.lower() in config()["arquivos_lixo"]


def casa_algum(nome, padroes):
    n = norm(str(nome).replace("_", " "))
    return any(norm(p) in n for p in padroes)

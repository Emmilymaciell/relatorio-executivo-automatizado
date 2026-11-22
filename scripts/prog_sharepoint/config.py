# -*- coding: utf-8 -*-
"""Configuração do acesso Graph: config/sharepoint.json (versionado, sem segredo) + credenciais (fora do repo).

Precedência das credenciais: variáveis de ambiente PROG_SP_TENANT_ID / PROG_SP_CLIENT_ID / PROG_SP_CLIENT_SECRET
> arquivo 'arquivo_credenciais' (KEY=VALOR, '#' comenta) > tenant/client do sharepoint.json. O segredo nunca é
impresso nem gravado em log (ver mascarar()).
"""
import json
import os

RAIZ_PROJETO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
ARQ_CONFIG = os.path.join(RAIZ_PROJETO, "config", "sharepoint.json")
CHAVES = ("PROG_SP_TENANT_ID", "PROG_SP_CLIENT_ID", "PROG_SP_CLIENT_SECRET")


class ErroConfig(Exception):
    pass


def carregar(caminho=None):
    with open(caminho or os.environ.get("PROG_SP_CONFIG") or ARQ_CONFIG, encoding="utf-8") as f:
        c = json.load(f)
    return {k: v for k, v in c.items() if not k.startswith("_")}


def caminho_credenciais(cfg):
    return os.path.expandvars(os.environ.get("PROG_SP_CREDENCIAIS") or cfg["arquivo_credenciais"])


def _ler_arquivo(path):
    out = {}
    if not os.path.isfile(path):
        return out
    with open(path, encoding="utf-8-sig") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            k, v = linha.split("=", 1)
            out[k.strip()] = v.strip().strip('"').strip("'")
    return out


def _dentro(path, raiz):
    try:
        return os.path.commonpath([os.path.abspath(path).lower(), os.path.abspath(raiz).lower()]) == \
            os.path.abspath(raiz).lower()
    except ValueError:
        return False


def credenciais(cfg):
    """Devolve dict tenant_id/client_id/client_secret/origem. Recusa arquivo de credenciais dentro do repositório."""
    arq = caminho_credenciais(cfg)
    if os.path.isfile(arq) and _dentro(arq, RAIZ_PROJETO):
        raise ErroConfig(f"o arquivo de credenciais não pode ficar dentro do repositório: {arq}")
    doarq = _ler_arquivo(arq)
    tenant = os.environ.get("PROG_SP_TENANT_ID") or doarq.get("PROG_SP_TENANT_ID") or cfg.get("tenant_id")
    client = os.environ.get("PROG_SP_CLIENT_ID") or doarq.get("PROG_SP_CLIENT_ID") or cfg.get("client_id")
    segredo = os.environ.get("PROG_SP_CLIENT_SECRET") or doarq.get("PROG_SP_CLIENT_SECRET")
    origem = "variável de ambiente" if os.environ.get("PROG_SP_CLIENT_SECRET") else (arq if segredo else None)
    if not segredo or segredo.upper().startswith("COLE_"):
        raise ErroConfig(
            "client secret não configurado. Cole o VALOR do segredo em PROG_SP_CLIENT_SECRET= no arquivo\n"
            f"  {arq}\n(ou defina a variável de ambiente PROG_SP_CLIENT_SECRET). Nunca dentro do repositório.")
    if not tenant or not client:
        raise ErroConfig("tenant_id/client_id ausentes (config/sharepoint.json).")
    return {"tenant_id": tenant, "client_id": client, "client_secret": segredo, "origem": origem}


def mascarar(segredo):
    if not segredo:
        return "<vazio>"
    return f"{segredo[:3]}…({len(segredo)} caracteres)"

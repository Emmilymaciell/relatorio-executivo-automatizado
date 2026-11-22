# -*- coding: utf-8 -*-
"""sharepoint.py — acesso ao SharePoint do Programa pela Microsoft Graph (app-only, Entra ID - Programa) — DEC-052.

Comandos (ALVO = caminho local sob a pasta sincronizada OU caminho na biblioteca, ex.: "4. Produtos/15. Setembro 2026"):
  testar                                  token, permissões do app, site, biblioteca e pastas da raiz (leitura)
  listar   [ALVO]                         conteúdo de uma pasta (leitura)
  info     ALVO                           metadados de um item (leitura)
  baixar   ALVO [--destino ARQ]           baixa um arquivo (padrão: output/SHAREPOINT_DOWNLOADS/...)
  comparar ARQ_LOCAL                      arquivo local sincronizado × SharePoint (tamanho e SHA-256)
  comparar-pasta PASTA_LOCAL [--recursivo] pasta sincronizada × SharePoint (nomes/tamanhos; sai 3 se divergir)
  enviar   ARQ_LOCAL [--para ALVO] [--substituir] [--aplicar] [--motivo TEXTO]
           DRY-RUN por padrão. Sem --para, o destino é o equivalente remoto do próprio caminho local.
           Arquivo existente só com --substituir: baixa o atual para output/SHAREPOINT_BACKUPS/<data>/ antes,
           envia com If-Match (eTag) e confere o SHA-256 depois (Office: OK_CONTEUDO = só metadados do SharePoint
           diferem). Nunca escreve em '5. Compartilhado Parceiro'.
  publicar-lote MANIFESTO.json [--aplicar] vários envios de uma vez: {"motivo": "...", "itens": [{"local": "...",
           "remoto": "...", "substituir": false}]}. Planeja todos; se algum estiver BLOQUEADO/ERRO, não envia nenhum.

Não há excluir. Mover/renomear/copiar no servidor: só via scripts/manutencao/reorganizar_mes_pop.py (DEC-070). Credenciais: config/sharepoint.json + segredo fora do repo (docs/COMO-USAR.md § 8).
"""
import argparse
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, "..", ".claude", "skills", "relatorio-prog-grupo1", "scripts"))

from prog_sharepoint import config as K  # noqa: E402
from prog_sharepoint.graph import Graph, ErroGraph  # noqa: E402
from prog_sharepoint.operacoes import SharePoint  # noqa: E402
from prog_sharepoint.caminhos import CaminhoInvalido  # noqa: E402


def conectar():
    import prog_config as cfg_g1  # mesma raiz sincronizada que o resto da automação usa
    cfg = K.carregar()
    cred = K.credenciais(cfg)
    g = Graph(cred["tenant_id"], cred["client_id"], cred["client_secret"])
    return SharePoint(cfg, g, cfg_g1.prog_root(), K.RAIZ_PROJETO), cred


def _p(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2, default=str))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("testar")
    s = sub.add_parser("listar"); s.add_argument("alvo", nargs="?", default="")
    s = sub.add_parser("info"); s.add_argument("alvo")
    s = sub.add_parser("baixar"); s.add_argument("alvo"); s.add_argument("--destino")
    s = sub.add_parser("espelhar"); s.add_argument("alvo"); s.add_argument("--destino", required=True)
    s.add_argument("--so-pasta", action="store_true", help="não desce nas subpastas")
    s = sub.add_parser("comparar"); s.add_argument("local")
    s = sub.add_parser("comparar-pasta"); s.add_argument("local"); s.add_argument("--recursivo", action="store_true")
    s = sub.add_parser("enviar"); s.add_argument("local"); s.add_argument("--para")
    s.add_argument("--substituir", action="store_true"); s.add_argument("--aplicar", action="store_true")
    s.add_argument("--motivo", default="")
    s = sub.add_parser("publicar-lote"); s.add_argument("manifesto"); s.add_argument("--aplicar", action="store_true")
    a = ap.parse_args(argv)

    try:
        sp, cred = conectar()
        if a.cmd == "testar":
            print(f"Credencial: client {cred['client_id']} · segredo {K.mascarar(cred['client_secret'])} "
                  f"· origem {cred['origem']}")
            _p(sp.testar())
        elif a.cmd == "listar":
            for i in sp.listar(a.alvo):
                tipo = "PASTA" if "folder" in i else f"{i.get('size', 0):>12,}"
                print(f"{tipo:>12}  {i.get('lastModifiedDateTime', '')[:19]}  {i['name']}")
        elif a.cmd == "info":
            it = sp.item(a.alvo)
            _p({k: it.get(k) for k in ("name", "size", "eTag", "lastModifiedDateTime", "webUrl", "file", "folder")}
               if it else {"existe": False, "remoto": sp.remoto(a.alvo)})
        elif a.cmd == "baixar":
            r = sp.remoto(a.alvo)
            destino = a.destino or os.path.join(K.RAIZ_PROJETO, "output", "SHAREPOINT_DOWNLOADS", *r.split("/"))
            _p(sp.baixar(a.alvo, destino))
        elif a.cmd == "espelhar":
            _p(sp.espelhar(a.alvo, os.path.abspath(a.destino), recursivo=not a.so_pasta))
        elif a.cmd == "comparar":
            _p(sp.comparar(os.path.abspath(a.local)))
        elif a.cmd == "comparar-pasta":
            r = sp.comparar_pasta(a.local, recursivo=a.recursivo)
            _p(r)
            print("\nEM DIA com o SharePoint." if r["em_dia"] else
                  "\nATENÇÃO: pasta local DIFERENTE do SharePoint (aguarde a sincronização do OneDrive ou confira).")
            return 0 if r["em_dia"] else 3
        elif a.cmd == "publicar-lote":
            with open(a.manifesto, encoding="utf-8") as f:
                man = json.load(f)
            res = sp.publicar_lote(man, aplicar=a.aplicar)
            for p in res["planos"]:
                v = (p.get("verificacao") or {}).get("status", "")
                print(f"{p['acao']:<10} {v:<12} {p['remoto']}" + (f"  ({p['motivo']})" if p.get("motivo") else ""))
            for e in res["erros"]:
                print(f"{'ERRO':<10} {'':<12} {e['local']}: {e['erro']}")
            if not res["pode_aplicar"]:
                print("\nLOTE NÃO ENVIADO: corrija os itens BLOQUEADO/ERRO (nada foi gravado).")
                return 2
            if not a.aplicar:
                print(f"\nDRY-RUN: {res['total']} item(ns) prontos. Repita com --aplicar para gravar.")
        elif a.cmd == "enviar":
            res = sp.enviar(a.local, a.para, substituir=a.substituir, aplicar=a.aplicar, motivo=a.motivo)
            _p(res)
            if res["acao"] == "BLOQUEADO":
                return 2
            if not a.aplicar:
                print("\nDRY-RUN: nada foi enviado. Repita com --aplicar para gravar.")
    except (K.ErroConfig, ErroGraph, CaminhoInvalido, FileNotFoundError) as e:
        print(f"ERRO: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

# -*- coding: utf-8 -*-
"""entrega.py — CLI da etapa da COORDENAÇÃO EXECUTIVA (POP de Entrega de Produtos Programa).

Comandos (todos DRY-RUN por padrão; nada é enviado ao Parceiro em nenhum modo):
  estrutura   --mes 16 [--aplicar]          padroniza a pasta mensal (ANEXOS única por grupo + Índice na raiz)
  indice      --mes 15 [--aplicar]          só cria o Índice "em branco" do mês na raiz (nunca sobrescreve)
  auditar     --mes 15                      inventário + validação + plano de correções (somente leitura)
  preparar    --mes 15                      STAGING + FINAL locais (output/ENTREGA_...), revalida, manifesto
  publicar    --mes 15 --confirmar <código> aplica o FINAL na própria pasta do mês no SharePoint
                                            (arquiva antes tudo que sai/muda; revalida depois)

--mes aceita o nº da pasta (ex.: 16 -> '16. Outubro 2026') ou --pasta "<caminho completo>".
Ver docs/COMO-USAR.md.
"""
import argparse
import datetime
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
sys.path.insert(0, os.path.join(AQUI, "..", ".claude", "skills", "relatorio-prog-grupo1", "scripts"))

from prog_entrega import regras as R  # noqa: E402
from prog_entrega import estrutura as E  # noqa: E402


def _dir_produtos():
    import prog_config as cfg  # reutiliza a localização do SharePoint da automação do G1
    return cfg.dir_produtos()


def resolver_pasta(args):
    if args.pasta:
        if not os.path.isdir(args.pasta):
            sys.exit(f"ERRO: pasta não existe: {args.pasta}")
        return os.path.abspath(args.pasta)
    prod = _dir_produtos()
    for s in sorted(os.listdir(prod)):
        info = R.parse_pasta_mes(s)
        if info and info[0] == int(args.mes):
            return os.path.join(prod, s)
    sys.exit(f"ERRO: pasta mensal nº {args.mes} não encontrada em {prod}")


def _log_dir():
    d = os.path.join(R.project_dir(), "logs", "entrega")
    os.makedirs(d, exist_ok=True)
    return d


def _ts():
    return datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")


def _imprimir_plano(plano):
    print(f"\nPASTA: {plano['pasta']}\nCOMPETÊNCIA: {plano['competencia']}")
    if not plano["acoes"]:
        print("  (nenhuma ação necessária)")
    for a in plano["acoes"]:
        print(f"  - {a['acao']}: {a.get('antes') or ''} -> {a.get('depois') or ''}  [{a['regra']}]")
    for p in plano["pendencias"]:
        print(f"  ! PENDÊNCIA: {p}")


def cmd_estrutura(args, so_indice=False):
    pasta = resolver_pasta(args)
    plano = E.planejar_estrutura(pasta)
    if so_indice:
        plano["acoes"] = [a for a in plano["acoes"] if a["acao"] in ("criar_indice", "substituir_indice")]
        plano["pendencias"] = [p for p in plano["pendencias"] if "índice" in p]
    _imprimir_plano(plano)
    if not args.aplicar:
        print("\nDRY-RUN: nada foi alterado. Use --aplicar para executar este plano.")
        return 0
    nome = os.path.basename(pasta).replace(" ", "_")
    res = E.aplicar_plano(plano, os.path.join(_log_dir(), f"{'indice' if so_indice else 'estrutura'}_{nome}_{_ts()}.json"))
    falhas = [a for a in res["acoes"] if a["resultado"] != "OK"]
    for a in res["acoes"]:
        print(f"  {a['resultado']:>6}  {a['acao']}: {a.get('antes') or a.get('depois')}")
    print(f"\nAPLICADO: {len(res['acoes']) - len(falhas)} ok, {len(falhas)} falha(s). Log em logs/entrega/.")
    return 1 if falhas else 0


def cmd_auditar(args, modo):
    from prog_entrega import pipeline as PL
    pasta = resolver_pasta(args)
    if modo == "publicar":
        return PL.publicar(pasta, args.confirmar)
    return PL.executar(pasta, preparar=(modo == "preparar"))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for nome in ("estrutura", "indice", "auditar", "preparar", "publicar"):
        sp = sub.add_parser(nome)
        g = sp.add_mutually_exclusive_group(required=True)
        g.add_argument("--mes", type=int, help="nº da pasta mensal (ex.: 15 -> '15. Setembro 2026')")
        g.add_argument("--pasta", help="caminho completo da pasta mensal")
        if nome in ("estrutura", "indice"):
            sp.add_argument("--aplicar", action="store_true", help="executa (sem isto: dry-run)")
        if nome == "publicar":
            sp.add_argument("--confirmar", required=True,
                            help="código do plano exibido pelo 'preparar' (garante que a usuária aprovou ESTE plano)")
    args = ap.parse_args(argv)
    if args.cmd == "estrutura":
        return cmd_estrutura(args)
    if args.cmd == "indice":
        return cmd_estrutura(args, so_indice=True)
    return cmd_auditar(args, args.cmd)


if __name__ == "__main__":
    sys.exit(main())

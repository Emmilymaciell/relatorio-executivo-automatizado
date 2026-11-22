# -*- coding: utf-8 -*-
"""Reorganiza a pasta do mês (Programa - 4. Produtos/NN. Mês AAAA) no formato do POP — TODOS os grupos — DEC-070.

Só com autorização explícita da usuária. Ordem obrigatória:
  1. snapshot   cópia NO SERVIDOR da pasta do mês, como os grupos enviaram, para
                3. Desenvolvimento/G1 - Coordenacao/Entregas Recebidas/<pasta do mês> (+ manifesto com SHA-256,
                autor e data de cada arquivo). Sem snapshot conferido (OK), nada é reorganizado.
  2. planejar   lê config/organizacao_pop.json[<mês>] (renomeações, retiradas e arquivos novos, cada um com motivo),
                confere contra o espelho local (origem existe, SHA-256, nome final no padrão, sem colisão, caminho
                <= 250) e gera o Índice completo (G1–G5) a partir do estado FINAL. Grava plano + código.
  3. aplicar    --confirmar <código>: envia os arquivos novos, move/renomeia pela API (If-Match, conflito = falha,
                quickXorHash conferido), tira do pacote o que sai (move para a pasta do snapshot, nunca apaga) e
                substitui o Índice (backup + eTag).
  4. verificar  relista o SharePoint e compara com o estado final planejado (nomes, tamanhos, Índice 100%).

  python scripts/manutencao/reorganizar_mes_pop.py snapshot --mes 16 [--aplicar]
  python scripts/manutencao/reorganizar_mes_pop.py planejar --mes 16
  python scripts/manutencao/reorganizar_mes_pop.py aplicar  --mes 16 --confirmar <código>
  python scripts/manutencao/reorganizar_mes_pop.py verificar --mes 16
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.abspath(os.path.join(AQUI, "..", ".."))
sys.path.insert(0, os.path.join(PROJ, "scripts"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import sharepoint as SPCLI  # noqa: E402
from prog_entrega import regras as R  # noqa: E402

DEST_SNAPSHOT = "3. Desenvolvimento/G1 - Coordenacao/Entregas Recebidas"
PRODUTOS = "4. Produtos"


def _jdump(obj, p):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1, default=str)


def _jload(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def _sha(p):
    h = hashlib.sha256()
    with open(R.lp(p), "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def contexto(mes):
    sp, _ = SPCLI.conectar()
    pastas = [i["name"] for i in sp.listar(PRODUTOS) if "folder" in i and re.match(rf"^{mes:02d}\.\s", i["name"])]
    if len(pastas) != 1:
        raise SystemExit(f"pasta do mês {mes} não encontrada (ou ambígua) em {PRODUTOS}: {pastas}")
    pasta = pastas[0]
    num, nome_mes, _, ano = R.parse_pasta_mes(pasta)
    out = os.path.join(PROJ, "output", f"ENTREGA_{num:02d}_{nome_mes.capitalize()}_{ano}", "05_REORGANIZACAO")
    os.makedirs(out, exist_ok=True)
    return {"sp": sp, "pasta": pasta, "mes_rem": f"{PRODUTOS}/{pasta}", "snap_rem": f"{DEST_SNAPSHOT}/{pasta}",
            "out": out, "num": num, "mes": nome_mes, "ano": ano}


# ---------------------------------------------------------------------------------------------- snapshot
def snapshot(mes, aplicar):
    c = contexto(mes)
    sp = c["sp"]
    r = sp.copiar_pasta(c["mes_rem"], c["snap_rem"], aplicar=aplicar,
                        motivo=f"Cópia como enviado pelos grupos (antes da reorganização POP) — {c['pasta']}")
    print(json.dumps({k: v for k, v in r.items() if k != "divergentes"}, ensure_ascii=False, indent=1))
    if not aplicar or r.get("verificacao") != "OK":
        if r.get("divergentes"):
            print("DIVERGENTES:", r["divergentes"])
        return r
    arqs, _ = sp.arvore(c["mes_rem"])
    espelho = {}
    pe = os.path.join(os.environ.get("PROG_ROOT", ""), "_espelho.json")
    if os.path.isfile(pe):
        espelho = _jload(pe)
    man = {"gerado_em": datetime.datetime.now().isoformat(timespec="seconds"), "origem": c["mes_rem"],
           "copia": c["snap_rem"], "verificacao_copia": r["verificacao"], "arquivos": []}
    for rel, v in sorted(arqs.items()):
        e = espelho.get(f"{c['mes_rem']}/{rel}") or {}
        man["arquivos"].append({"arquivo": rel, "tamanho": v["size"], "quickXorHash": v["qxh"],
                                "sha256": e.get("sha256") if e.get("tamanho") == v["size"] else None,
                                "enviado_por": v["criado_por"], "enviado_em": v["criado_em"],
                                "modificado_por": v["modificado_por"], "modificado_em": v["modificado_em"]})
    pj = os.path.join(c["out"], f"Manifesto da cópia como enviado - {c['pasta']}.json")
    _jdump(man, pj)
    linhas = [f"# Cópia como enviado — {c['pasta']}", "",
              f"Cópia feita no servidor em {man['gerado_em'][:16].replace('T', ' ')}, ANTES da reorganização no "
              f"formato do POP, a partir de `{c['mes_rem']}`. Conferência arquivo a arquivo (tamanho e "
              f"quickXorHash): **{r['verificacao']}** ({len(arqs)} arquivos).", "",
              "Observação: a subdivisão GRUPO 1 do Índice e os anexos P1.1 da Coordenação já tinham sido gravados "
              "pela Coordenação em 16/11/2026; o Índice exatamente como os grupos deixaram está ao lado, em "
              "`… - Índice como recebido (antes da Coordenação).docx`.", "",
              "| Arquivo | Tamanho | Enviado por | Enviado em (UTC) | SHA-256 |", "|---|---|---|---|---|"]
    for a in man["arquivos"]:
        linhas.append(f"| {a['arquivo']} | {a['tamanho']} | {a['enviado_por'] or '—'} | {a['enviado_em'] or '—'} | "
                      f"{(a['sha256'] or '—')[:16]}… |")
    pm = pj[:-5] + ".md"
    with open(pm, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")
    print(f"[snapshot] manifesto: {pj}")
    return r


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("etapa", choices=["snapshot", "planejar", "aplicar", "verificar"])
    ap.add_argument("--mes", type=int, required=True)
    ap.add_argument("--aplicar", action="store_true")
    ap.add_argument("--confirmar")
    a = ap.parse_args()
    if a.etapa == "snapshot":
        snapshot(a.mes, a.aplicar)
    else:
        import reorganizar_mes_pop_plano as P  # noqa: E402  (mesmo diretório)
        getattr(P, a.etapa)(a.mes, contexto, confirmar=a.confirmar)


if __name__ == "__main__":
    sys.path.insert(0, AQUI)
    main()

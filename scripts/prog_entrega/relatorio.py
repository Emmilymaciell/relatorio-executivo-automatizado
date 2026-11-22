# -*- coding: utf-8 -*-
"""Saídas: relatório de prontidão (Markdown), achados (JSON/CSV), manifesto do pacote (JSON/CSV)."""
import csv
import json
import os
from collections import Counter, defaultdict

from . import regras as R

_ORD = {s: i for i, s in enumerate(R.STATUS_ORDEM)}


def pior(statuses):
    ss = [s for s in statuses if s]
    return min(ss, key=lambda s: _ORD.get(s, 99)) if ss else R.NA


def veredito(achados_finais):
    c = Counter(a["status"] for a in achados_finais)
    if c[R.BLOCKER]:
        return R.BLOCKED
    if c[R.MANUAL]:
        return R.PENDENTE_HUMANO
    return R.READY


# Checklist oficial da COORDENAÇÃO EXECUTIVA (aba do Checklist_Visual_Entregas_PROG_FINAL.xlsx), item a item,
# com as etapas da auditoria que o sustentam. Itens sem verificação automática possível = MANUAL.
CHECKLIST_COORD = [
    (1, "Confirmar o recebimento de G2, G3, G4 e G5 até o 3º dia útil.", ["relatorio", "estrutura"], ["G2", "G3", "G4", "G5"]),
    (2, "Confirmar o recebimento de P1.3 e P1.4 do G1 até o 3º dia útil.", ["anexos"], ["G1"]),
    (3, "Conferir Lista de Demandas + Relatório de Produtos Mensal + ANEXOS + Índice de cada grupo.", ["lista_demandas", "relatorio", "anexos", "indice"], None),
    (4, "Identificar pendências e solicitar correção imediata ao grupo responsável.", None, None),
    (5, "Finalizar P1.1 e P1.2.", ["relatorio"], ["G1"]),
    (6, "Consolidar o Relatório de Produtos Mensal do G1.", ["relatorio"], ["G1"]),
    (7, "Providenciar e conferir todas as assinaturas necessárias.", ["assinatura"], None),
    (8, "Conferir a nomenclatura de todos os anexos.", ["nomenclatura"], None),
    (9, "Conferir o limite de 100 MB por arquivo e a compactação em ZIP dos arquivos que não forem PDF.", ["tamanho", "formato"], None),
    (10, "Conferir o Índice cumulativo completo — G1 a G5.", ["indice"], None),
    (11, "Remover do pacote final as Listas de Demandas.", ["lista_demandas"], None),
    (12, "Fora da pasta ANEXOS: remover Word, minutas, duplicatas e versões anteriores do Relatório.", ["relatorio", "duplicidade", "versoes"], None),
    (13, "Em cada grupo, deixar na raiz somente o Relatório final assinado em PDF + a pasta ANEXOS.", ["estrutura", "relatorio"], None),
    (14, "Na raiz da pasta do mês, manter o Índice completo e atualizado.", ["indice"], None),
    (15, "Abrir/testar os arquivos e verificar a integridade do conjunto documental.", ["integridade", "planilha", "fragmentos"], None),
    (16, "Fazer a conferência final para protocolo e confirmar que o pacote está pronto para envio ao Parceiro.", "*", None),
]


def checklist(achados):
    linhas = []
    for n, txt, etapas, grupos in CHECKLIST_COORD:
        if etapas is None:
            pend = [a for a in achados if a["status"] in (R.BLOCKER, R.MANUAL)]
            st = R.PASS if not pend else R.MANUAL
            linhas.append((n, txt, st, f"{len(pend)} pendência(s) a solicitar aos grupos" if pend else "sem pendências"))
            continue
        sel = [a for a in achados if (etapas == "*" or a["etapa"] in etapas) and (not grupos or a["grupo"] in grupos)]
        st = pior([a["status"] for a in sel]) if sel else R.NA
        if st == R.FIXED:
            st = R.PASS
        linhas.append((n, txt, st, Counter(a["status"] for a in sel).most_common()))
    return linhas


def _md_tab(cab, linhas):
    out = ["| " + " | ".join(cab) + " |", "|" + "---|" * len(cab)]
    for l in linhas:
        out.append("| " + " | ".join(str(x).replace("|", "/").replace("\n", " ") for x in l) + " |")
    return "\n".join(out)


def escrever(out_dir, ctx, ach_origem, ach_final, transformacoes, manifesto):
    os.makedirs(out_dir, exist_ok=True)
    base = ach_final if ach_final is not None else ach_origem
    ver = veredito(base)
    if ach_final is None:
        ver = f"{ver} (PREVISTO — só auditoria; o pacote ainda não foi montado/revalidado)"
    todos = ach_origem + (ach_final or [])
    with open(os.path.join(out_dir, "achados.json"), "w", encoding="utf-8") as f:
        json.dump({"contexto": ctx, "veredito": ver, "achados": todos}, f, ensure_ascii=False, indent=1)
    with open(os.path.join(out_dir, "achados.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["fase", "status", "etapa", "grupo", "arquivo", "mensagem", "regra", "acao"])
        for a in todos:
            w.writerow([a["fase"], a["status"], a["etapa"], a["grupo"] or "", a["arquivo"] or "", a["mensagem"],
                        a["regra"], a["acao"] or ""])
    if manifesto is not None:
        with open(os.path.join(out_dir, "manifesto.json"), "w", encoding="utf-8") as f:
            json.dump(manifesto, f, ensure_ascii=False, indent=1)
        with open(os.path.join(out_dir, "manifesto.csv"), "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["caminho_relativo", "nome", "extensao", "tamanho_bytes", "grupo", "classificacao", "sha256",
                        "resultado_validacao"])
            for m in manifesto["arquivos"]:
                w.writerow([m["caminho_relativo"], m["nome"], m["extensao"], m["tamanho"], m["grupo"] or "",
                            m["classificacao"], m["sha256"], m["resultado_validacao"]])
    with open(os.path.join(out_dir, "transformacoes.json"), "w", encoding="utf-8") as f:
        json.dump(transformacoes or [], f, ensure_ascii=False, indent=1)

    L = []
    L.append(f"# Relatório de Prontidão — Entrega ao Parceiro — {ctx['competencia']}\n")
    L.append(f"**VEREDITO: {ver}**\n")
    L.append(_md_tab(["Campo", "Valor"], [
        ("Competência", ctx["competencia"]), ("Período", f"{ctx['periodo'][0]} a {ctx['periodo'][1]}"),
        ("Pasta analisada", ctx["pasta"]), ("Modo", ctx["modo"]), ("Executado em", ctx["executado_em"]),
        ("Versão da automação", ctx["versao"]), ("Regras", ctx["regras"]),
        ("Prazo grupos (3º dia útil)", f"{ctx['prazos']['grupos']} — {ctx['prazos']['base']}"),
        ("Prazo Coordenação/envio (5º dia útil)", ctx["prazos"]["coordenacao"]),
        ("Arquivos inventariados (origem)", ctx["n_origem"]),
        ("Arquivos no pacote final", ctx.get("n_final", "—")),
        ("Código de publicação", ctx.get("codigo_publicacao") or "—"),
    ]))
    L.append("\n> O pacote **nunca** é enviado ao Parceiro por esta automação. Ela só prepara e prova a prontidão.\n")
    L.append("## Status por grupo\n")
    linhas = []
    for g in R.config()["grupos"]:
        sel = [a for a in base if a["grupo"] == g]
        c = Counter(a["status"] for a in sel)
        rel = next((a["mensagem"] for a in base if a["grupo"] == g and a["etapa"] == "assinatura"), "—")
        linhas.append((g, pior([a["status"] for a in sel]), c[R.BLOCKER], c[R.MANUAL], c[R.WARNING], c[R.FIXED],
                       rel[:120]))
    L.append(_md_tab(["Grupo", "Status", "BLOCKER", "MANUAL", "WARNING", "FIXED", "Assinatura do relatório"], linhas))
    L.append("\n## Checklist da Coordenação Executiva (Checklist Visual oficial)\n")
    L.append(_md_tab(["Nº", "Verificação", "Status", "Base"], checklist(base)))
    for titulo, st in (("BLOCKERS (impedem o protocolo)", R.BLOCKER), ("Revisão humana necessária", R.MANUAL),
                       ("Avisos", R.WARNING)):
        sel = [a for a in base if a["status"] == st]
        L.append(f"\n## {titulo} — {len(sel)}\n")
        if sel:
            L.append(_md_tab(["Grupo", "Etapa", "Arquivo", "Problema", "Regra", "Ação possível"],
                             [(a["grupo"] or "—", a["etapa"], a["arquivo"] or "—", a["mensagem"], a["regra"],
                               a["acao"] or "—") for a in sel]))
    exc = [a for a in base if "[EXCEÇÃO ACEITA pela Coordenação" in (a.get("mensagem") or "")]
    L.append(f"\n## Exceções aceitas pela Coordenação nesta competência (revisão humana) — {len(exc)}\n")
    if exc:
        L.append("> Não contam como pendência no veredito; valem só para este conteúdo (SHA-256) e esta competência.\n")
        L.append(_md_tab(["Grupo", "Etapa", "Arquivo", "Status original", "Achado + decisão"],
                         [(a["grupo"] or "—", a["etapa"], a["arquivo"] or "—", a.get("status_original", "—"),
                           a["mensagem"]) for a in exc]))
    fixed = [a for a in ach_origem if a["status"] == R.FIXED]
    L.append(f"\n## Correções automáticas seguras — {len(fixed)}\n")
    if fixed:
        L.append(_md_tab(["Grupo", "Etapa", "Arquivo (origem)", "O que foi feito", "Regra"],
                         [(a["grupo"] or "—", a["etapa"], a["arquivo"] or "—", a["mensagem"], a["regra"]) for a in fixed]))
    if transformacoes:
        L.append("\n## Transformações executadas (antes → depois → motivo → regra → resultado)\n")
        L.append(_md_tab(["Antes", "Operação", "Depois", "Motivo", "Regra", "Resultado"],
                         [(t["antes"], t["operacao"], ", ".join(t["depois"] or ["—"]), t["motivo"], t["regra"],
                           t["resultado"]) for t in transformacoes]))
    if manifesto:
        L.append(f"\n## Manifesto do pacote final — {len(manifesto['arquivos'])} arquivo(s), "
                 f"{manifesto['tamanho_total'] / 1048576:.1f} MB\n")
        L.append("Lista completa com SHA-256 em `manifesto.csv` / `manifesto.json`.\n")
        por_g = defaultdict(list)
        for m in manifesto["arquivos"]:
            por_g[m["grupo"] or "raiz"].append(m)
        L.append(_md_tab(["Grupo", "Arquivos", "MB", "Maior (MB)", "Formatos"],
                         [(g, len(ms), f"{sum(x['tamanho'] for x in ms) / 1048576:.1f}",
                           f"{max(x['tamanho'] for x in ms) / 1048576:.1f}",
                           ", ".join(f"{k}:{v}" for k, v in Counter(x['extensao'] for x in ms).items()))
                          for g, ms in sorted(por_g.items())]))
    L.append("\n## Pergunta-teste (regra final do Checklist)\n")
    L.append("> *Se o Parceiro recebesse esta pasta agora, conseguiria protocolar tudo imediatamente, sem renomear, "
             "reorganizar ou solicitar correção?*\n")
    L.append(f"**Resposta da automação: {'SIM' if ver == R.READY else 'NÃO'}** — {ver}.\n")
    with open(os.path.join(out_dir, "RELATORIO_PRONTIDAO.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return ver

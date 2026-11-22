# -*- coding: utf-8 -*-
"""Auditoria de uma pasta mensal contra o POP + Checklist.

`auditar(inv, cache, fase)`:
  fase="ORIGEM" -> achados + PLANO do pacote final (o que copiar/zipar/renomear/fracionar/excluir).
  fase="FINAL"  -> revalidação do pacote montado: o mesmo código roda sobre o FINAL; qualquer coisa que
                   ainda "precisaria de correção" (FIXED) vira BLOCKER, porque o FINAL tem de estar pronto.
Cada achado: status (PASS/FIXED_AUTOMATICALLY/WARNING/MANUAL_REVIEW/BLOCKER/N/A), etapa, regra (fonte),
grupo, arquivo, mensagem, ação possível.
"""
import os
import re
from collections import defaultdict

from . import arquivos as A
from . import assinatura as S
from . import indice as IDX
from . import nomes as N
from . import regras as R

_RE_VERSAO = re.compile(r"\b(v\d+|rev\.?\s?[a-z0-9&]+|final|revisad[oa]|assinad[oa]|c[oó]pia|copy|"
                        r"ultima versao|versao final)\b|\(\d+\)|[_\-]+$", re.I)


class Achados:
    def __init__(self, fase):
        self.fase = fase
        self.itens = []

    def add(self, status, etapa, regra, msg, grupo=None, arquivo=None, acao=None, sha=None, codigo=None,
            persistente=False):
        """persistente=True: pendência da ORIGEM que o pacote NÃO resolve e que a revalidação do FINAL não
        consegue re-detectar (ex.: arquivo fora da estrutura deixado de fora, colisão, Lista ausente).
        Entra no veredito final mesmo assim — nada some silenciosamente."""
        if self.fase == "FINAL" and status == R.FIXED:
            status, msg = R.BLOCKER, "pacote FINAL ainda requer correção: " + msg
        self.itens.append({"fase": self.fase, "status": status, "etapa": etapa, "regra": regra, "grupo": grupo,
                           "arquivo": arquivo, "mensagem": msg, "acao": acao, "sha256": sha, "codigo": codigo,
                           "persistente": persistente})


def _val(cache, it):
    return cache.validacao(it["sha256"], A.VERSAO_VALIDADOR, lambda: A.validar(it["abs"]))


def _sig(cache, it):
    return cache.validacao(it["sha256"], S.VERSAO, lambda: S.verificar_pdf(it["abs"]))


def _chave_versao(stem):
    s = R.norm(stem.replace("_", " "))
    s = _RE_VERSAO.sub(" ", s)
    return re.sub(r"[\s\-.]+", " ", s).strip()


def _status_assinatura(st):
    return {S.VALIDADA: R.PASS, S.NAO_ENCONTRADA: R.BLOCKER, S.VISUAL: R.MANUAL,
            S.INVALIDA: R.BLOCKER, S.NAO_VERIFICAVEL: R.MANUAL}.get(st, R.MANUAL)


def _problemas_arquivo(ach, it, v, etapa, regra):
    for st, msg in v["problemas"]:
        ach.add(st, etapa, regra, msg, it["grupo"], it["rel"], sha=it["sha256"], codigo="integridade")


def auditar(inv, cache, fase="ORIGEM", hoje=None):
    cfg = R.config()
    ach = Achados(fase)
    plano = {"final": [], "excluidos": [], "renomeacoes": []}
    itens = inv["itens"]
    lim = R.limite_bytes()
    pasta_nome = os.path.basename(inv["pasta"].rstrip("\\/"))
    info_mes = R.parse_pasta_mes(pasta_nome)

    # ---------------------------------------------------------------- 1. estrutura geral
    for g in cfg["grupos"]:
        if g not in inv["grupos_presentes"]:
            ach.add(R.BLOCKER, "estrutura", "POP Quadro 2", f"pasta do grupo {g} ausente", g,
                    acao="criar a pasta e receber a entrega do grupo")
        anx = inv["pastas_anexos"].get(g, [])
        if len(anx) > 1:
            ach.add(R.BLOCKER, "estrutura", "POP Definições (ANEXOS = subpasta única)",
                    f"mais de uma pasta de anexos: {anx}", g, acao="unificar manualmente", persistente=True)
        elif len(anx) == 1 and anx[0] != cfg["nome_pasta_anexos"]:
            ach.add(R.FIXED, "estrutura", "POP Definições (ANEXOS)",
                    f"pasta '{anx[0]}' será padronizada como '{cfg['nome_pasta_anexos']}' no pacote", g)
    for it in itens:
        if it["classe"] == R.K_LIXO:
            continue
        if it["classe"] == R.K_TEMP:
            ach.add(R.FIXED, "estrutura", "Checklist Coord. item 12 (arquivos provisórios)",
                    "arquivo temporário/provisório excluído do pacote", it["grupo"], it["rel"], sha=it.get("sha256"))
            plano["excluidos"].append({"origem": it["rel"], "motivo": "temporário"})
        elif it["classe"] == R.K_FORA:
            onde = {"raiz_mes": "solto na raiz do mês", "raiz_grupo": "solto na raiz do grupo (fora de ANEXOS)",
                    "outra_pasta": "em pasta fora da estrutura"}.get(it["local"], it["local"])
            ach.add(R.BLOCKER, "estrutura", "POP Quadro 2 / Checklist Coord. 13",
                    f"arquivo {onde}: na raiz do grupo só pode ficar o Relatório final assinado + ANEXOS",
                    it["grupo"], it["rel"], acao="mover para ANEXOS com nome padrão OU retirar do pacote (decisão humana)",
                    sha=it.get("sha256"), persistente=True)

    # ---------------------------------------------------------------- 2. por grupo
    finais_por_grupo = {}
    for g in cfg["grupos"]:
        gi = [it for it in itens if it["grupo"] == g]
        finais_por_grupo[g] = _auditar_grupo(g, gi, cache, ach, plano, lim, fase, pasta_nome)

    # ---------------------------------------------------------------- 2b. comprimento do caminho publicado
    lim_cam = cfg.get("caminho_max_windows")
    if lim_cam:
        for g, es in finais_por_grupo.items():
            for e in es:
                if e["op"] == "excluir":
                    continue
                cam = os.path.join(R.sem_lp(inv["pasta"]), *e["destino"].split("/"))
                if len(cam) > lim_cam:
                    ach.add(R.WARNING, "nomenclatura", "Windows MAX_PATH (cópia da pasta p/ o Parceiro)",
                            f"caminho com {len(cam)} caracteres na pasta do mês (> {lim_cam}) — pode falhar ao "
                            "copiar pelo Windows Explorer; encurtar o nome", g, e["origem"], sha=e["sha256"])

    # ---------------------------------------------------------------- 3. índice
    _auditar_indice(inv, itens, finais_por_grupo, ach, plano, fase, info_mes)
    return ach, plano, finais_por_grupo


def _lista_arquivada(pasta_nome, g):
    """Lista de Demandas do grupo já ARQUIVADA no arquivo interno (DEC-045) — caso de uma pasta do mês já publicada:
    o POP manda retirá-la do pacote, então a ausência na pasta não é pendência. Devolve o caminho ou None."""
    if not pasta_nome:
        return None
    try:
        import prog_config as cfg_g1
        base = os.path.join(cfg_g1.prog_root(), R.config()["arquivo_interno"]["relativo_prog_root"], pasta_nome)
    except Exception:
        return None
    base = R.lp(base)
    if not os.path.isdir(base):
        return None
    for data in sorted(os.listdir(base), reverse=True):        # publicação mais recente primeiro
        gd = os.path.join(base, data, g)
        if os.path.isdir(gd):
            for f in os.listdir(gd):
                if R.casa_algum(f, R.config()["padroes_lista_demandas"]):
                    return os.path.join(gd, f)
    return None


def _auditar_grupo(g, gi, cache, ach, plano, lim, fase, pasta_nome=None):
    cfg = R.config()
    finais = []  # (destino_rel, entrada_do_plano)

    # --- Lista de Demandas (controle interno: usa-se, mas NÃO vai ao pacote)
    listas = [it for it in gi if it["classe"] == R.K_LISTA]
    if listas:
        for it in listas:
            ach.add(R.FIXED, "lista_demandas", "POP Quadro 2 / Etapa 6 / Checklist Coord. 11",
                    "Lista de Demandas conferida e RETIRADA do pacote final (permanece no arquivo interno)",
                    g, it["rel"], sha=it["sha256"])
            plano["excluidos"].append({"origem": it["rel"], "motivo": "Lista de Demandas (controle interno)"})
    elif fase == "ORIGEM" and _lista_arquivada(pasta_nome, g):
        ach.add(R.PASS, "lista_demandas", "POP Quadro 2 / Etapa 6 / Checklist Coord. 11",
                "Lista de Demandas já retirada do pacote e guardada no arquivo interno (pasta do mês publicada): "
                + os.path.basename(_lista_arquivada(pasta_nome, g)), g)
    elif fase == "ORIGEM":
        ach.add(R.MANUAL, "lista_demandas", "POP Quadro 1 / Checklist Grupos item 1",
                "Lista de Demandas não encontrada na pasta do grupo", g,
                acao="solicitar ao grupo (é insumo da conferência e do P1.1)", persistente=True)

    # --- Relatório de Produtos Mensal
    rels = [it for it in gi if it["classe"] == R.K_RELATORIO]
    for it in [x for x in gi if x["classe"] == R.K_RELATORIO_EDITAVEL]:
        ach.add(R.FIXED, "relatorio", "POP Etapa 6 / Checklist Coord. 12",
                "versão editável/minuta do relatório RETIRADA do pacote (fica no arquivo interno)", g, it["rel"],
                sha=it["sha256"])
        plano["excluidos"].append({"origem": it["rel"], "motivo": "Word/minuta do relatório"})
    escolhido = None
    infos = []
    for it in rels:
        v = _val(cache, it)
        sig = _sig(cache, it) if v["tipo"] == "pdf" and not any(s == R.BLOCKER for s, _ in v["problemas"]) else None
        infos.append((it, v, sig))
    if not rels:
        ach.add(R.BLOCKER, "relatorio", "POP Quadro 1/2 / Checklist Grupos item 2",
                "Relatório de Produtos Mensal FINAL (PDF assinado) ausente na raiz do grupo", g,
                acao="G1: concluir P1.1/P1.2 e assinar; G2–G5: solicitar ao grupo")
    elif len(rels) == 1:
        escolhido = infos[0]
    else:
        por_hash = defaultdict(list)
        for x in infos:
            por_hash[x[0]["sha256"]].append(x)
        unicos = [v[0] for v in por_hash.values()]
        for v in por_hash.values():
            for dup in v[1:]:
                ach.add(R.FIXED, "relatorio", "Checklist Grupos item 3", "cópia idêntica do relatório retirada do pacote",
                        g, dup[0]["rel"], sha=dup[0]["sha256"])
                plano["excluidos"].append({"origem": dup[0]["rel"], "motivo": "duplicata exata do relatório"})
        if len(unicos) == 1:
            escolhido = unicos[0]
        else:
            validados = [x for x in unicos if x[2] and x[2]["status"] == S.VALIDADA]
            com_estrutura = [x for x in unicos if x[2] and x[2].get("assinaturas")]
            alvo = validados if len(validados) == 1 else (com_estrutura if len(com_estrutura) == 1 else None)
            if alvo:
                escolhido = alvo[0]
                for x in unicos:
                    if x is not escolhido:
                        ach.add(R.FIXED, "relatorio", "POP Etapa 6 / Checklist Grupos item 3",
                                f"versão não assinada/anterior do relatório retirada do pacote (mantida a assinada: "
                                f"'{escolhido[0]['nome']}')", g, x[0]["rel"], sha=x[0]["sha256"])
                        plano["excluidos"].append({"origem": x[0]["rel"], "motivo": "versão anterior/não assinada do relatório"})
            else:
                ach.add(R.BLOCKER, "relatorio", "Checklist Grupos item 3 / Checklist Coord. 12",
                        f"{len(unicos)} versões diferentes do relatório e não é possível determinar a final: "
                        f"{[x[0]['nome'] for x in unicos]}", g, acao="grupo/Coordenação indicar a versão final",
                        persistente=True)
    if escolhido:
        it, v, sig = escolhido
        _problemas_arquivo(ach, it, v, "relatorio", "POP Quadro 4 (Formato/Integridade)")
        if sig:
            st = _status_assinatura(sig["status"])
            det = sig["assinaturas"][0] if sig.get("assinaturas") else {}
            quem = f" — signatário: {det.get('signatario')}; emissor: {det.get('emissor')} ({det.get('emissor_org')})" if det.get("signatario") else ""
            ach.add(st, "assinatura", "POP Quadro 4 (Assinaturas) / Checklist Grupos item 2",
                    f"assinatura do relatório: {sig['status']} — {sig['motivo']}{quem}", g, it["rel"],
                    acao=None if st == R.PASS else "conferir/reassinar o relatório", sha=it["sha256"], codigo="assinatura_relatorio")
            if sig["status"] == S.VALIDADA and cfg.get("exigir_icp_brasil") and not det.get("icp_brasil"):
                ach.add(R.BLOCKER, "assinatura", "config exigir_icp_brasil", "assinatura válida mas NÃO é ICP-Brasil",
                        g, it["rel"], sha=it["sha256"])
        if it["tamanho"] > lim:
            ach.add(R.BLOCKER, "tamanho", "POP Quadro 4 (Tamanho)", f"relatório com {it['tamanho']/1048576:.1f} MB "
                    f"(> {cfg['limite_tamanho_mb']} MB) — relatório assinado não pode ser fracionado automaticamente",
                    g, it["rel"], sha=it["sha256"])
        destino = f"{g}/{it['nome']}"
        e = {"origem": it["rel"], "origem_abs": it["abs"], "sha256": it["sha256"], "destino": destino,
             "op": "copiar", "grupo": g, "papel": "relatorio", "produto": "REL"}
        plano["final"].append(e)
        finais.append(e)

    # --- ANEXOS
    anexos = [it for it in gi if it["classe"] == R.K_ANEXO]
    if not anexos:
        ach.add(R.MANUAL, "anexos", "POP Quadro 1 / Checklist Grupos item 4",
                "nenhum anexo na pasta ANEXOS do grupo", g, acao="confirmar com o grupo se não há anexos no mês")
    propostas = []
    for it in anexos:
        v = _val(cache, it)
        _problemas_arquivo(ach, it, v, "integridade", "POP Quadro 4 (Integridade) / Checklist Grupos item 11")
        if it["local"] == "anexos_sub":
            ach.add(R.FIXED, "estrutura", "POP Quadro 2 (ANEXOS sem subpastas)",
                    f"arquivo em subpasta '{os.path.dirname(it['rel'].split('/', 2)[2])}' será colocado direto em ANEXOS",
                    g, it["rel"], sha=it["sha256"])
        an = N.analisar(it["nome"])
        final_nome, rodovia_ok = it["nome"], an["conforme"]
        if not an["conforme"] or any("grafia padrão" in m for _, m in an["problemas"]):
            prop, motivo, completo = N.propor(it["nome"], it["produto_pasta"], contexto=it["rel"])
            if completo and prop and cfg.get("renomear_propostas_completas"):
                final_nome = prop
                ach.add(R.FIXED, "nomenclatura", "POP Quadro 3", f"renomeado no pacote: '{it['nome']}' -> '{prop}' "
                        f"({motivo})", g, it["rel"], sha=it["sha256"])
                plano["renomeacoes"].append({"origem": it["rel"], "antes": it["nome"], "depois": prop, "motivo": motivo,
                                             "regra": "POP Quadro 3"})
            else:
                for st, msg in an["problemas"]:
                    ach.add(st, "nomenclatura", "POP Quadro 3 / Checklist Grupos itens 5–6", msg, g, it["rel"],
                            acao=(f"sugestão: '{prop}'" if prop else motivo), sha=it["sha256"], codigo="nome")
        else:
            for st, msg in an["problemas"]:
                ach.add(st, "nomenclatura", "POP Quadro 3", msg, g, it["rel"], sha=it["sha256"], codigo="nome")
        e = {"origem": it["rel"], "origem_abs": it["abs"], "sha256": it["sha256"], "grupo": g, "papel": "anexo",
             "produto": N.analisar(final_nome)["codigo"] or it["produto_pasta"], "tipo": v["tipo"],
             "tamanho": it["tamanho"], "info": v.get("info", {})}
        stem, ext = os.path.splitext(final_nome)
        if v["tipo"] == "pdf":
            e.update(op="copiar", destino=f"{g}/{cfg['nome_pasta_anexos']}/{stem}.pdf")
            if it["tamanho"] > lim:
                sig = _sig(cache, it)
                if sig.get("assinaturas"):
                    ach.add(R.BLOCKER, "tamanho", "POP Quadro 4 (Tamanho)", f"PDF ASSINADO com {it['tamanho']/1048576:.1f} MB "
                            "— fracionar quebraria a assinatura; o grupo deve fracionar e assinar cada parte", g, it["rel"],
                            sha=it["sha256"])
                else:
                    e["op"] = "fracionar_pdf"
                    ach.add(R.FIXED, "tamanho", "POP Quadro 4 / Checklist Grupos item 9",
                            f"PDF com {it['tamanho']/1048576:.1f} MB será fracionado em partes ≤ {cfg['limite_tamanho_mb']} MB "
                            "(Parte 01, Parte 02...) sem alterar conteúdo", g, it["rel"], sha=it["sha256"])
        elif v["tipo"] == "zip":
            e.update(op="copiar", destino=f"{g}/{cfg['nome_pasta_anexos']}/{stem}.zip")
            if it["tamanho"] > lim:
                if len(v["info"].get("membros", [])) > 1:
                    e["op"] = "reempacotar_zip"
                    ach.add(R.FIXED, "tamanho", "POP Quadro 4", "ZIP > limite será reempacotado em partes por arquivo",
                            g, it["rel"], sha=it["sha256"])
                else:
                    ach.add(R.BLOCKER, "tamanho", "POP Quadro 4", "ZIP com um único arquivo > limite — fracionamento "
                            "exige decisão do grupo", g, it["rel"], sha=it["sha256"])
        else:
            e.update(op="zipar", destino=f"{g}/{cfg['nome_pasta_anexos']}/{stem}.zip", nome_interno=stem + ext.lower())
            ach.add(R.FIXED, "formato", "POP Quadro 3 / Etapa 2 / Checklist Grupos item 8",
                    f"arquivo não-PDF ({ext or 'sem extensão'}) será compactado em ZIP (original preservado dentro)",
                    g, it["rel"], sha=it["sha256"])
            if it["tamanho"] > lim:
                ach.add(R.WARNING, "tamanho", "POP Quadro 4", "arquivo > limite antes de compactar — o tamanho do ZIP "
                        "será conferido na revalidação final", g, it["rel"], sha=it["sha256"])
        if ext.lower() in (".xlsx", ".xlsm"):
            f = v["info"].get("formulas")
            if f == 0:
                ach.add(R.WARNING, "planilha", "POP Quadro 4 (editável correspondente) / PROMPT Parceiro",
                        "planilha sem nenhuma fórmula — se for memória de cálculo, conferir se é a versão editável",
                        g, it["rel"], sha=it["sha256"])
            elif f:
                ach.add(R.PASS, "planilha", "PROMPT Parceiro (fórmulas preservadas)", f"planilha íntegra com {f} fórmula(s)",
                        g, it["rel"], sha=it["sha256"])
        plano["final"].append(e)
        finais.append(e)
        propostas.append(e)

    # zip já no FINAL: conteúdo é planilha? (revalidação) -> fórmulas dentro do zip não são reabertas aqui

    # --- colisões, duplicatas exatas, versões, fragmentos (sobre os NOMES FINAIS)
    por_dest = defaultdict(list)
    for e in finais:
        por_dest[R.norm(e["destino"])].append(e)
    for k, es in por_dest.items():
        if len(es) > 1:
            if len({x["sha256"] for x in es}) == 1:
                for x in es[1:]:
                    x["op"] = "excluir"
                    ach.add(R.FIXED, "duplicidade", "Checklist Coord. 12 (duplicatas)",
                            f"duplicata exata de '{es[0]['origem']}' com o mesmo nome final — mantida uma cópia", g, x["origem"],
                            sha=x["sha256"])
                    plano["excluidos"].append({"origem": x["origem"], "motivo": "duplicata exata"})
            else:
                ach.add(R.BLOCKER, "duplicidade", "POP Quadro 4 (versão vigente) / regra 'nunca sobrescrever'",
                        f"COLISÃO: {len(es)} arquivos DIFERENTES resultariam no mesmo nome '{es[0]['destino']}': "
                        f"{[x['origem'] for x in es]}", g, acao="renomear/decidir manualmente", persistente=True)
                for x in es[1:]:
                    x["op"] = "excluir"
    ativos = [e for e in finais if e["op"] != "excluir" and e["papel"] == "anexo"]
    por_sha = defaultdict(list)
    for e in ativos:
        por_sha[e["sha256"]].append(e)
    _copia = re.compile(r"\s*\(\d+\)(?=\s*(-|$))|\b(c[oó]pia|copy)\b", re.I)
    for sha, es in por_sha.items():
        if len(es) > 1:
            limpos = [x for x in es if not _copia.search(os.path.splitext(os.path.basename(x["destino"]))[0])]
            if len(limpos) == 1:
                for x in es:
                    if x is not limpos[0]:
                        x["op"] = "excluir"
                        ach.add(R.FIXED, "duplicidade", "POP Quadro 4 (somente a versão vigente)",
                                f"cópia idêntica ('{os.path.basename(x['destino'])}') retirada; mantido "
                                f"'{os.path.basename(limpos[0]['destino'])}'", g, x["origem"], sha=sha)
                        plano["excluidos"].append({"origem": x["origem"], "motivo": "duplicata exata (cópia '(n)')"})
                continue
            ach.add(R.MANUAL, "duplicidade", "POP Quadro 4 (somente a versão vigente)",
                    f"conteúdo IDÊNTICO com nomes diferentes: {[os.path.basename(x['destino']) for x in es]}", g,
                    acao="manter só um (decisão humana)", sha=sha, codigo="dup_nomes")
    por_ver = defaultdict(list)
    for e in ativos:
        stem = os.path.splitext(os.path.basename(e["destino"]))[0]
        ext = os.path.splitext(e["destino"])[1].lower()
        por_ver[_chave_versao(N.base_sem_parte(stem)) + "|" + str(N.parte(stem)) + "|" + ext].append(e)
    for k, es in por_ver.items():
        if len({x["sha256"] for x in es}) > 1:
            ach.add(R.BLOCKER, "versoes", "POP Quadro 4 (versão vigente) / Checklist Coord. 12",
                    f"VERSÕES CONFLITANTES do mesmo documento: {[os.path.basename(x['destino']) for x in es]}", g,
                    acao="grupo indicar a vigente; as demais saem do pacote")
    frag = defaultdict(list)
    for e in ativos:
        stem = os.path.splitext(os.path.basename(e["destino"]))[0]
        p = N.parte(stem)
        if p is not None:
            frag[R.norm(N.base_sem_parte(stem))].append(p)
    for base, ps in frag.items():
        esperado = list(range(1, max(ps) + 1))
        if sorted(ps) != esperado:
            falt = sorted(set(esperado) - set(ps))
            dup = sorted({p for p in ps if ps.count(p) > 1})
            ach.add(R.BLOCKER, "fragmentos", "POP Quadro 3 / Checklist Grupos item 9",
                    f"sequência de partes incompleta/inconsistente em '{base}': presentes {sorted(ps)}"
                    + (f"; faltando {['Parte %02d' % x for x in falt]}" if falt else "")
                    + (f"; repetidas {dup}" if dup else ""), g, acao="solicitar as partes faltantes ao grupo")
        else:
            ach.add(R.PASS, "fragmentos", "POP Quadro 3", f"partes 01..{max(ps):02d} completas em '{base}'", g)
    # planilha sem PDF correspondente
    stems_pdf = {R.norm(os.path.splitext(os.path.basename(e["destino"]))[0]) for e in ativos if e["destino"].endswith(".pdf")}
    for e in ativos:
        nome_orig = e.get("nome_interno") or os.path.basename(e["origem"])
        if os.path.splitext(nome_orig)[1].lower() in (".xlsx", ".xlsm", ".xls"):
            if R.norm(os.path.splitext(os.path.basename(e["destino"]))[0]) not in stems_pdf:
                ach.add(R.WARNING, "planilha", "POP Quadro 3 ('editável deverá corresponder à versão em PDF')",
                        "planilha sem PDF correspondente de mesmo nome — conferir se o PDF é exigido", g, e["origem"],
                        sha=e["sha256"])
    # produtos ativos sem anexo / recebidos
    cods = {e["produto"] for e in ativos}
    for c in sorted(x for x in cods if x):
        n = sum(1 for e in ativos if e["produto"] == c)
        ach.add(R.PASS, "anexos", "POP Etapa 1", f"{c}: {n} anexo(s) recebido(s)", g)
    for c in R.produtos_do_grupo(g):
        if c in R.config()["produtos_ativos_padrao"] and c not in cods and ativos:
            ach.add(R.WARNING, "anexos", "POP Etapa 1 (todos os anexos)",
                    f"nenhum anexo identificado para {c} - {R.produtos()[c]['nome']} — confirmar se houve demanda no mês", g)
    return finais


def _auditar_indice(inv, itens, finais_por_grupo, ach, plano, fase, info_mes):
    indices = [it for it in itens if it["classe"] == R.K_INDICE]
    lidos = []
    if not indices:
        if fase == "ORIGEM":
            ach.add(R.FIXED, "indice", "POP Quadro 2 / Checklist Coord. 14",
                    "Índice ausente na raiz do mês — será gerado no pacote a partir dos arquivos finais (Modelo aprovado)")
        else:
            ach.add(R.BLOCKER, "indice", "POP Quadro 2", "Índice ausente na raiz do pacote final")
    for it in indices:
        r = IDX.ler_indice(it["abs"])
        if r["erros"]:
            ach.add(R.WARNING if fase == "ORIGEM" else R.BLOCKER, "indice", "Modelo de Índice",
                    f"índice ilegível: {r['erros']}", None, it["rel"])
        lidos.append((it, r))
    if len(indices) > 1:
        ach.add(R.FIXED, "indice", "POP Quadro 2 (índice único)", f"{len(indices)} índices na raiz — o pacote terá UM "
                "índice regenerado (SEI de todos preservado)", None)
    sei = {}
    for it, r in lidos:
        for ln in r["linhas"]:
            if ln["sei"]:
                for k in (R.norm(ln["arquivo"]), R.norm(os.path.splitext(ln["arquivo"])[0])):
                    sei[k] = ln["sei"]
    plano["sei_por_arquivo"] = sei
    reais = {}
    for g, es in finais_por_grupo.items():
        for e in es:
            if e["op"] != "excluir":
                reais[R.norm(os.path.basename(e["destino"]))] = g
    reais_sem_ext = {os.path.splitext(k)[0]: g for k, g in reais.items()}
    divergencias = 0
    for it, r in lidos:
        vistos = defaultdict(int)
        for ln in r["linhas"]:
            k = R.norm(ln["arquivo"])
            vistos[k] += 1
            g_real = reais.get(k) or reais_sem_ext.get(os.path.splitext(k)[0]) or reais_sem_ext.get(k)
            if not g_real:
                divergencias += 1
                ach.add(R.WARNING if fase == "ORIGEM" else R.BLOCKER, "indice", "Modelo de Índice / Checklist Coord. 10",
                        f"linha do índice SEM arquivo correspondente: '{ln['arquivo']}' ({ln['grupo']})", ln["grupo"], it["rel"])
            elif ln["grupo"] and g_real != ln["grupo"]:
                divergencias += 1
                ach.add(R.WARNING if fase == "ORIGEM" else R.BLOCKER, "indice", "Modelo de Índice",
                        f"'{ln['arquivo']}' listado no {ln['grupo']} mas está no {g_real}", ln["grupo"], it["rel"])
        for k, n in vistos.items():
            if n > 1:
                divergencias += 1
                ach.add(R.WARNING if fase == "ORIGEM" else R.BLOCKER, "indice", "Modelo de Índice",
                        f"arquivo listado {n}x no índice: '{k}'", None, it["rel"])
        listados = set(vistos) | {os.path.splitext(k)[0] for k in vistos}
        faltam = defaultdict(list)
        for k, g in reais.items():
            if k not in listados and os.path.splitext(k)[0] not in listados:
                faltam[g].append(k)
        for g, ks in sorted(faltam.items()):
            divergencias += len(ks)
            ach.add(R.WARNING if fase == "ORIGEM" else R.BLOCKER, "indice", "POP Etapa 3 / Checklist Grupos item 10",
                    f"{len(ks)} arquivo(s) do {g} NÃO listado(s) no índice (o grupo não atualizou sua subdivisão)"
                    + (f": {ks[:5]}{' …' if len(ks) > 5 else ''}"), g, it["rel"])
        por_g = defaultdict(list)
        for ln in r["linhas"]:
            por_g[ln["grupo"]].append(ln["num"])
        for g_, nums in por_g.items():  # Modelo r2: 1, 2, 3… reiniciando em cada grupo (DEC-049)
            if [str(i) for i in range(1, len(nums) + 1)] != nums:
                ach.add(R.WARNING if fase == "ORIGEM" else R.BLOCKER, "indice", "Modelo de Índice r2 (numeração por grupo)",
                        f"numeração do {g_} não é 1, 2, 3…: {nums[:15]}", g_, it["rel"])
        n_sei = sum(1 for ln in r["linhas"] if ln["sei"])
        ach.add(R.NA, "indice", "Modelo de Índice (PROTOCOLO/SEI)",
                f"SEI preenchido em {n_sei}/{len(r['linhas'])} linha(s) — o nº SEI é atribuído no protocolo; nunca inventado",
                None, it["rel"])
    if fase == "ORIGEM" and lidos:
        ach.add(R.FIXED, "indice", "POP Quadro 2 / Checklist Coord. 10 e 14",
                f"índice do pacote será REGENERADO a partir dos arquivos finais ({len(reais)} arquivo(s); "
                f"{divergencias} divergência(s) corrigida(s); SEI preservado em {len(sei) // 2 if sei else 0})")
    if fase == "FINAL" and lidos and divergencias == 0:
        ach.add(R.PASS, "indice", "POP Quadro 2 / Checklist Coord. 10", "índice reconcilia 100% com os arquivos do pacote")

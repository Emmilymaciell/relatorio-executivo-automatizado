# -*- coding: utf-8 -*-
"""Estrutura padrão da pasta mensal (POP Quadro 2):

  {NN. Mês AAAA}/
    Índice de Arquivos - Entrega de Produtos Programa - {Mês} {AAAA}.docx   (raiz do mês; Modelo r2)
    G1/ … G5/
      <Relatório de Produtos Mensal final assinado>.pdf                  (raiz do grupo)
      ANEXOS/                                                           (pasta ÚNICA, sem subpastas)

`planejar_estrutura` NUNCA altera nada; `aplicar_plano` executa só ações seguras:
  - criar pasta ANEXOS / pasta do grupo ausente;
  - renomear 'Anexos' -> 'ANEXOS' (só caixa);
  - remover subpasta de ANEXOS **comprovadamente vazia** (0 arquivos; só lixo de sistema como
    desktop.ini/Thumbs.db) — usa os.rmdir, que falha se houver qualquer coisa;
  - criar o Índice do mês a partir do Modelo (nunca sobrescreve).
Subpasta com qualquer arquivo real NÃO é tocada: vira pendência (o achatamento de anexos com
conteúdo é feito pela pipeline da Coordenação, com registro antes->depois).
"""
import datetime
import json
import os
import stat

from . import indice as IDX
from . import regras as R


def _arquivos_reais(d):
    reais, lixo = [], []
    for raiz, _dirs, files in os.walk(d):
        for f in files:
            (lixo if R.eh_lixo(f) else reais).append(os.path.join(raiz, f))
    return reais, lixo


def _layout_vigente(mes, ano, periodo):
    """Assinatura da tabela de um índice em branco gerado AGORA com o Modelo vigente (comparação)."""
    import tempfile
    tmp = os.path.join(tempfile.mkdtemp(prefix="idx"), "i.docx")
    IDX.criar_modelo_mes(tmp, mes, ano, periodo=periodo)
    return IDX.assinatura_tabela(tmp)


def planejar_estrutura(pasta_mes, criar_indice=True):
    nome = os.path.basename(pasta_mes.rstrip("\\/"))
    info = R.parse_pasta_mes(nome)
    if not info:
        raise ValueError(f"Pasta mensal fora do padrão 'NN. Mês AAAA': {nome}")
    _num, mes, mes_num, ano = info
    anexos_nome = R.config()["nome_pasta_anexos"]
    acoes, pend = [], []
    for g in R.config()["grupos"]:
        gdir = os.path.join(pasta_mes, g)
        if not os.path.isdir(gdir):
            acoes.append({"acao": "criar_pasta", "antes": None, "depois": gdir,
                          "motivo": f"pasta do grupo {g} ausente", "regra": "POP Quadro 2"})
        cands = [a for a in (os.listdir(gdir) if os.path.isdir(gdir) else [])
                 if os.path.isdir(os.path.join(gdir, a)) and R.norm(a) == "anexos"]
        if len(cands) > 1:
            pend.append(f"{g}: mais de uma pasta 'anexos' ({cands}) — decisão humana")
            continue
        adir = os.path.join(gdir, cands[0]) if cands else os.path.join(gdir, anexos_nome)
        if not cands:
            acoes.append({"acao": "criar_pasta", "antes": None, "depois": adir,
                          "motivo": "pasta ANEXOS ausente", "regra": "POP Definições/Quadro 2"})
        for s in sorted(os.listdir(adir)) if cands else []:
            sdir = os.path.join(adir, s)
            if not os.path.isdir(sdir):
                continue
            reais, lixo = _arquivos_reais(sdir)
            if reais:
                pend.append(f"{g}: subpasta '{s}' tem {len(reais)} arquivo(s) — não removida "
                            "(achatamento é etapa da pipeline da Coordenação)")
            else:
                acoes.append({"acao": "remover_pasta_vazia", "antes": sdir, "depois": None,
                              "lixo_sistema": [os.path.basename(x) for x in lixo],
                              "motivo": "subpasta por produto dentro de ANEXOS (vazia)",
                              "regra": "POP Quadro 2 — ANEXOS é subpasta única; produto vai no NOME do arquivo"})
        if cands and cands[0] != anexos_nome:
            acoes.append({"acao": "renomear_pasta", "antes": adir,
                          "depois": os.path.join(gdir, anexos_nome),
                          "motivo": f"padronizar '{cands[0]}' -> '{anexos_nome}'", "regra": "POP Definições (ANEXOS)"})
    if criar_indice:
        destino = os.path.join(pasta_mes, IDX.nome_arquivo_indice(mes, ano))
        existentes = [f for f in os.listdir(pasta_mes) if os.path.isfile(os.path.join(pasta_mes, f))
                      and R.casa_algum(f, R.config()["padroes_indice"])]
        base = {"periodo": R.periodo_do_mes(mes_num, ano), "mes": mes, "ano": ano}
        antigos = [f for f in existentes if os.path.normcase(os.path.join(pasta_mes, f)) != os.path.normcase(destino)]
        if os.path.exists(destino):
            if IDX.eh_indice_em_branco(destino) and IDX.assinatura_tabela(destino) != _layout_vigente(mes, ano, base["periodo"]):
                acoes.append({"acao": "substituir_indice", "antes": destino, "depois": destino, **base,
                              "motivo": "índice em branco num layout anterior -> layout vigente (cópia guardada antes)",
                              "regra": "Modelo de Índice r2, só por grupo (DEC-049)"})
            else:
                pend.append(f"índice vigente já existe na raiz: {os.path.basename(destino)} — mantido")
        elif not existentes:
            acoes.append({"acao": "criar_indice", "antes": None, "depois": destino, **base,
                          "motivo": "índice cumulativo na raiz do mês", "regra": "POP Quadro 2 + Modelo de Índice"})
        for f in antigos:
            p = os.path.join(pasta_mes, f)
            if IDX.eh_indice_em_branco(p):
                acoes.append({"acao": "substituir_indice", "antes": p, "depois": destino, **base,
                              "motivo": "índice em branco no modelo anterior -> modelo vigente (cópia guardada antes)",
                              "regra": "Modelo de Índice r2 (DEC-049)"})
            else:
                pend.append(f"índice '{f}' já tem arquivos/SEI preenchidos — NÃO substituído (migrar com a usuária)")
    return {"pasta": pasta_mes, "competencia": f"{mes}/{ano}", "acoes": acoes, "pendencias": pend}


def aplicar_plano(plano, log_path=None):
    resultados = []
    # ordem segura: remover subpastas vazias -> renomear -> criar -> índice
    ordem = {"remover_pasta_vazia": 0, "renomear_pasta": 1, "criar_pasta": 2, "criar_indice": 3, "substituir_indice": 4}
    for a in sorted(plano["acoes"], key=lambda x: ordem[x["acao"]]):
        r = dict(a)
        try:
            if a["acao"] == "remover_pasta_vazia":
                reais, lixo = _arquivos_reais(a["antes"])
                if reais:  # re-checagem imediatamente antes de agir
                    raise RuntimeError(f"deixou de estar vazia: {len(reais)} arquivo(s)")
                for x in lixo:
                    os.chmod(x, 0o666)
                    os.remove(x)
                for raiz, _dirs, _f in os.walk(a["antes"], topdown=False):
                    # o OneDrive marca pastas como somente-leitura (atributo R) — impede rmdir mesmo
                    # vazia. Limpa SÓ o atributo; os.rmdir continua falhando se houver conteúdo.
                    os.chmod(raiz, stat.S_IWRITE | stat.S_IREAD | stat.S_IEXEC)
                    os.rmdir(raiz)
            elif a["acao"] == "renomear_pasta":
                if os.path.normcase(a["antes"]) == os.path.normcase(a["depois"]):
                    tmp = a["antes"] + "__tmp_case"
                    os.rename(a["antes"], tmp)
                    os.rename(tmp, a["depois"])
                else:
                    if os.path.exists(a["depois"]):
                        raise RuntimeError("destino já existe")
                    os.rename(a["antes"], a["depois"])
            elif a["acao"] == "criar_pasta":
                os.makedirs(a["depois"], exist_ok=True)
            elif a["acao"] == "criar_indice":
                ok = IDX.criar_modelo_mes(a["depois"], a["mes"], a["ano"], periodo=a["periodo"])
                if not ok:
                    raise RuntimeError("índice já existia — não sobrescrito")
            elif a["acao"] == "substituir_indice":
                if not IDX.eh_indice_em_branco(a["antes"]):  # re-checagem imediatamente antes de agir
                    raise RuntimeError("o índice antigo passou a ter conteúdo — não substituído")
                import shutil
                from . import hashcache as H
                bkp = os.path.join(os.environ.get("PROG_ENTREGA_LOGS") or os.path.join(R.project_dir(), "logs", "entrega"),
                                   "indices_substituidos",
                                   os.path.basename(plano["pasta"].rstrip("\\/")), os.path.basename(a["antes"]))
                os.makedirs(os.path.dirname(bkp), exist_ok=True)
                shutil.copy2(a["antes"], bkp)
                if H.sha256_arquivo(bkp) != H.sha256_arquivo(a["antes"]):
                    raise RuntimeError("cópia de segurança divergente — nada alterado")
                os.chmod(a["antes"], stat.S_IWRITE)
                os.remove(a["antes"])  # só depois da cópia conferida (antes e depois podem ser o mesmo caminho)
                if not os.path.exists(a["depois"]):
                    IDX.criar_modelo_mes(a["depois"], a["mes"], a["ano"], periodo=a["periodo"])
                r["backup"] = bkp
            r["resultado"] = "OK"
        except Exception as e:
            r["resultado"] = f"FALHOU: {e}"
        resultados.append(r)
    saida = {"pasta": plano["pasta"], "competencia": plano["competencia"],
             "executado_em": datetime.datetime.now().isoformat(timespec="seconds"),
             "versao": R.VERSAO_AUTOMACAO, "acoes": resultados, "pendencias": plano["pendencias"]}
    if log_path:
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(saida, f, ensure_ascii=False, indent=2)
    return saida

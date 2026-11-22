# -*- coding: utf-8 -*-
"""Orquestração da etapa da Coordenação.

  executar(pasta, preparar=False)   AUDITORIA (dry-run): inventário + validação + plano. Nada é escrito
                                    fora de output/ENTREGA_<NN>_<Mês>_<AAAA>/.
  executar(pasta, preparar=True)    + STAGING (cópia verificada) + FINAL (layout POP) + REVALIDAÇÃO FINAL
                                    + MANIFESTO + plano de publicação com CÓDIGO.
  publicar(pasta, codigo)           aplica o FINAL na própria pasta do mês (decisão da usuária 01/10/2026:
                                    "o pacote final é a pasta de produtos do SharePoint"). Antes, ARQUIVA
                                    (backup local + arquivo interno no SharePoint) tudo que sai ou muda;
                                    só então remove/grava. Depois reinventaria e exige igualdade com o
                                    manifesto. Nunca envia nada ao Parceiro.
Estado retomável em estado.json (lido pelo /session-start).
"""
import datetime
import hashlib
import json
import os
import shutil
import stat
import sys

from . import calendario as CAL
from . import hashcache as H
from . import inventario as INV
from . import pacote as PK
from . import regras as R
from . import relatorio as REL
from . import validacao as V


def dir_saida(pasta):
    info = R.parse_pasta_mes(os.path.basename(pasta.rstrip("\\/")))
    if not info:
        raise ValueError(f"pasta mensal fora do padrão 'NN. Mês AAAA': {pasta}")
    num, mes, _m, ano = info
    base = os.environ.get("PROG_ENTREGA_OUT") or os.path.join(R.project_dir(), "output")
    return R.lp(os.path.join(base, f"ENTREGA_{num:02d}_{R.norm(mes).capitalize()}_{ano}")), info


def _agora():
    return datetime.datetime.now().isoformat(timespec="seconds")


MARCA_EXCECAO = "[EXCEÇÃO ACEITA pela Coordenação"


def _revisoes(out):
    p = os.path.join(out, "revisoes_humanas.json")
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else []


def _confirmacoes(out):
    """revisoes_humanas.json da competência (decisões da usuária; nunca por conta própria). Tipos de entrada:
    - {"sha256","codigo","decisao":"aceito"}           MANUAL_REVIEW -> PASS (só o MESMO conteúdo: sha256)
    - {"sha256","codigo","decisao":"aceito_excecao"}   BLOCKER ou MANUAL -> PASS como EXCEÇÃO DESTA COMPETÊNCIA,
                                                         listada no relatório (DEC-057)
    - {"grupo","etapa","decisao":"aceito"|"aceito_excecao"}  achado SEM arquivo (ex.: Lista de Demandas ausente)
    - {"sha256","decisao":"relatorio_do_grupo"}        PDF da raiz do grupo designado como o relatório (nome fora
                                                         do padrão "Relatório…") — tratado por _designacoes
    Toda entrada vale só para o conteúdo/competência registrados; arquivo alterado = confirmação deixa de valer."""
    conf = {}
    for c in _revisoes(out):
        if c.get("decisao") not in ("aceito", "aceito_excecao"):
            continue
        if c.get("sha256"):
            conf[(c["sha256"], c.get("codigo"))] = c
        elif c.get("grupo") and c.get("etapa"):
            conf[("grupo", c["grupo"], c["etapa"])] = c
    return conf


def _designacoes(out):
    return {c["sha256"]: c for c in _revisoes(out) if c.get("decisao") == "relatorio_do_grupo" and c.get("sha256")}


def _aplicar_designacoes(inv, desig):
    for it in inv["itens"]:
        c = desig.get(it.get("sha256"))
        if c and it.get("local") == "raiz_grupo" and it.get("ext") == ".pdf":
            it["classe"] = R.K_RELATORIO
            it["designado_por"] = f"{c.get('por')} em {c.get('em')}"


def _aplicar_confirmacoes(achados, conf):
    for a in achados:
        c = conf.get((a.get("sha256"), a.get("codigo")))
        if c is None and not a.get("arquivo"):
            c = conf.get(("grupo", a.get("grupo"), a.get("etapa")))
        if not c:
            continue
        if c["decisao"] == "aceito" and a["status"] == R.MANUAL:
            a["status"] = R.PASS
            a["mensagem"] += f" [CONFIRMADO por revisão humana: {c.get('por')} em {c.get('em')} — {c.get('justificativa', '')}]"
        elif c["decisao"] == "aceito_excecao" and a["status"] in (R.BLOCKER, R.MANUAL):
            a["status_original"] = a["status"]
            a["status"] = R.PASS
            a["mensagem"] += f" {MARCA_EXCECAO}: {c.get('por')} em {c.get('em')} — {c.get('justificativa', '')}]"


def _salvar_estado(out, **kw):
    p = os.path.join(out, "estado.json")
    est = json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}
    est.update(kw, atualizado_em=_agora(), versao=R.VERSAO_AUTOMACAO)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(est, f, ensure_ascii=False, indent=2)
    return est


def _manifesto(inv, achados_final):
    por_arq = {}
    for a in achados_final:
        if a["arquivo"]:
            por_arq.setdefault(a["arquivo"], []).append(a["status"])
    arqs = []
    for it in inv["itens"]:
        if it["classe"] == R.K_LIXO:
            continue
        arqs.append({"caminho_relativo": it["rel"], "nome": it["nome"], "extensao": it["ext"], "tamanho": it["tamanho"],
                     "grupo": it["grupo"], "classificacao": it["classe"], "sha256": it["sha256"],
                     "resultado_validacao": REL.pior(por_arq.get(it["rel"], [R.PASS]))})
    return {"gerado_em": _agora(), "arquivos": arqs, "tamanho_total": sum(a["tamanho"] for a in arqs)}


def _plano_publicacao(inv_origem, inv_final):
    src = {R.norm(it["rel"]): it for it in inv_origem["itens"] if it["classe"] != R.K_LIXO}
    fin = {R.norm(it["rel"]): it for it in inv_final["itens"] if it["classe"] != R.K_LIXO}
    remover = sorted(it["rel"] for k, it in src.items() if k not in fin or fin[k]["sha256"] != it["sha256"])
    gravar = sorted(it["rel"] for k, it in fin.items() if k not in src or src[k]["sha256"] != it["sha256"])
    manter = sorted(it["rel"] for k, it in fin.items() if k in src and src[k]["sha256"] == it["sha256"])
    plano = {"remover_da_pasta_do_mes_apos_arquivar": remover, "gravar_do_final": gravar, "manter_intactos": manter,
             "hash_origem": {it["rel"]: it["sha256"] for it in src.values()},
             "hash_final": {it["rel"]: it["sha256"] for it in fin.values()}}
    plano["codigo"] = hashlib.sha256(json.dumps(plano, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:10]
    return plano


def executar(pasta, preparar=False, hoje=None):
    out, info = dir_saida(pasta)
    os.makedirs(out, exist_ok=True)
    num, mes, mes_num, ano = info
    cache = H.Cache(os.path.join(out, "cache.json"))
    print(f"[1/6] INVENTÁRIO da origem: {pasta}")
    inv0 = INV.inventariar(pasta, cache)
    desig = _designacoes(out)
    _aplicar_designacoes(inv0, desig)
    os.makedirs(os.path.join(out, "01_INVENTARIO"), exist_ok=True)
    json.dump(inv0, open(os.path.join(out, "01_INVENTARIO", "inventario_origem.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print(f"      {len(inv0['itens'])} arquivo(s)")
    print("[2/6] VALIDAÇÃO + RECONCILIAÇÃO (origem)")
    ach0, plano, _ = V.auditar(inv0, cache, "ORIGEM", hoje)
    conf = _confirmacoes(out)
    _aplicar_confirmacoes(ach0.itens, conf)
    ctx = {"competencia": f"{mes}/{ano}", "periodo": R.periodo_do_mes(mes_num, ano), "pasta": pasta,
           "modo": "PREPARAÇÃO (staging + final locais)" if preparar else "AUDITORIA (dry-run, somente leitura)",
           "executado_em": _agora(), "versao": R.VERSAO_AUTOMACAO, "regras": R.config()["versao_regras"],
           "prazos": CAL.prazos_da_competencia(mes_num, ano), "n_origem": len(inv0["itens"])}
    json.dump(plano, open(os.path.join(out, "01_INVENTARIO", "plano_pacote.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    ach_final = manifesto = trans = None
    if preparar:
        print("[3/6] STAGING (cópia fiel, conferida por SHA-256)")
        PK.montar_staging(inv0, os.path.join(out, "02_STAGING"))
        print("[4/6] CORREÇÕES SEGURAS -> FINAL")
        trans, _ent = PK.montar_final(plano, os.path.join(out, "02_STAGING"), os.path.join(out, "03_FINAL", os.path.basename(pasta)), info)
        print("[5/6] REVALIDAÇÃO do FINAL (todas as verificações de novo)")
        final_root = os.path.join(out, "03_FINAL", os.path.basename(pasta))
        inv1 = INV.inventariar(final_root, cache)
        _aplicar_designacoes(inv1, desig)
        a1, _p, _f = V.auditar(inv1, cache, "FINAL", hoje)
        _aplicar_confirmacoes(a1.itens, conf)
        for t in trans:
            if str(t["resultado"]).startswith("FALHOU"):
                a1.add(R.BLOCKER, "montagem", t["regra"] or "pacote", f"transformação falhou: {t['resultado']}", None, t["antes"])
        herdadas = [dict(a, fase="ORIGEM (pendência herdada)") for a in ach0.itens
                    if a.get("persistente") and a["status"] in (R.BLOCKER, R.MANUAL)]
        ach_final = a1.itens + herdadas
        manifesto = _manifesto(inv1, ach_final)
        pub = _plano_publicacao(inv0, inv1)
        json.dump(pub, open(os.path.join(out, "plano_publicacao.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        ctx.update(n_final=len(manifesto["arquivos"]), codigo_publicacao=pub["codigo"])
    print("[6/6] RELATÓRIO")
    ver = REL.escrever(os.path.join(out, "04_RELATORIOS"), ctx, ach0.itens, ach_final, trans, manifesto)
    cache.salvar()
    base = ach_final if ach_final is not None else ach0.itens
    from collections import Counter
    cont = Counter(a["status"] for a in base)
    est = _salvar_estado(out, competencia=ctx["competencia"], pasta=pasta,
                         etapa="PREPARADO" if preparar else "AUDITADO", veredito=ver,
                         inventariados=len(inv0["itens"]), por_status=dict(cont),
                         por_grupo={g: REL.pior([a["status"] for a in base if a["grupo"] == g]) for g in R.config()["grupos"]},
                         codigo_publicacao=ctx.get("codigo_publicacao"),
                         relatorio=R.sem_lp(os.path.join(out, "04_RELATORIOS", "RELATORIO_PRONTIDAO.md")))
    print(f"\nVEREDITO: {ver}")
    print("  " + ", ".join(f"{k}={v}" for k, v in sorted(cont.items())))
    print(f"  relatório: {R.sem_lp(est['relatorio'])}")
    if preparar:
        print(f"  pacote FINAL local: {R.sem_lp(os.path.join(out, '03_FINAL'))}")
        if ver.startswith(R.BLOCKED):
            print("  publicação BLOQUEADA até zerar os BLOCKERs.")
        else:
            print(f"  para aplicar na pasta do SharePoint (com arquivamento prévio): "
                  f"python scripts/entrega.py publicar --pasta \"{pasta}\" --confirmar {ctx['codigo_publicacao']}")
    return 0 if ver == R.READY else (2 if ver.startswith(R.BLOCKED) else 1)


def _rm(p):
    os.chmod(p, stat.S_IWRITE)
    os.remove(p)


def publicar(pasta, codigo):
    out, info = dir_saida(pasta)
    pp = os.path.join(out, "plano_publicacao.json")
    est_p = os.path.join(out, "estado.json")
    if not (os.path.exists(pp) and os.path.exists(est_p)):
        sys.exit("ERRO: rode 'preparar' antes de 'publicar'.")
    plano = json.load(open(pp, encoding="utf-8"))
    est = json.load(open(est_p, encoding="utf-8"))
    if codigo != plano["codigo"]:
        sys.exit(f"ERRO: código de confirmação não confere (esperado o do último 'preparar'). Nada foi alterado.")
    if str(est.get("veredito", "")).startswith(R.BLOCKED):
        sys.exit("ERRO: pacote com BLOCKER — publicação recusada. Nada foi alterado.")
    cache = H.Cache(os.path.join(out, "cache.json"))
    inv_agora = INV.inventariar(pasta, cache)
    agora = {it["rel"]: it["sha256"] for it in inv_agora["itens"] if it["classe"] != R.K_LIXO}
    if {R.norm(k): v for k, v in agora.items()} != {R.norm(k): v for k, v in plano["hash_origem"].items()}:
        sys.exit("ERRO: a pasta do mês MUDOU depois do 'preparar' (arquivo novo/alterado/removido). "
                 "Rode 'preparar' de novo. Nada foi alterado.")
    final_root = os.path.join(out, "03_FINAL", os.path.basename(pasta))
    pasta = R.lp(pasta)
    import prog_config as cfg_g1  # localização do SharePoint (mesma da automação G1)
    arq_int = R.lp(os.path.join(cfg_g1.prog_root(), R.config()["arquivo_interno"]["relativo_prog_root"],
                           os.path.basename(pasta), datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")))
    bkp = os.path.join(out, "05_BACKUP_PUBLICACAO", datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S"))
    log = []
    # 1) ARQUIVAR tudo que sai/muda (backup local + arquivo interno no SharePoint), conferindo hash
    for rel in plano["remover_da_pasta_do_mes_apos_arquivar"]:
        src = os.path.join(pasta, *rel.split("/"))
        for raiz in (bkp, arq_int):
            dst = os.path.join(raiz, *rel.split("/"))
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            if H.sha256_arquivo(dst) != plano["hash_origem"][rel]:
                sys.exit(f"ERRO: cópia de arquivamento divergente ({rel}). Interrompido ANTES de remover qualquer coisa.")
        log.append({"antes": rel, "depois": f"ARQUIVO INTERNO + backup local", "motivo": "sai/muda no pacote",
                    "regra": "nunca perder original", "resultado": "ARQUIVADO"})
    # 2) remover da pasta do mês
    for rel in plano["remover_da_pasta_do_mes_apos_arquivar"]:
        _rm(os.path.join(pasta, *rel.split("/")))
        log.append({"antes": rel, "depois": None, "motivo": "retirado do pacote (arquivado)", "regra": "POP Etapa 6",
                    "resultado": "REMOVIDO"})
    # 3) gravar o FINAL
    for rel in plano["gravar_do_final"]:
        dst = os.path.join(pasta, *rel.split("/"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(os.path.join(final_root, *rel.split("/")), dst)
        log.append({"antes": None, "depois": rel, "motivo": "pacote final", "regra": "POP Quadro 2", "resultado": "GRAVADO"})
    # 4) remover diretórios que ficaram vazios e padronizar 'ANEXOS'
    for raiz, dirs, files in os.walk(pasta, topdown=False):
        if raiz == pasta:
            continue
        rel = os.path.relpath(raiz, pasta).split(os.sep)
        reais = [f for f in os.listdir(raiz) if not R.eh_lixo(f)]
        if not reais and not (len(rel) == 2 and R.norm(rel[1]) == "anexos") and len(rel) > 1:
            for f in os.listdir(raiz):
                _rm(os.path.join(raiz, f))
            os.chmod(raiz, stat.S_IWRITE | stat.S_IREAD | stat.S_IEXEC)
            os.rmdir(raiz)
            log.append({"antes": os.path.relpath(raiz, pasta), "depois": None, "motivo": "pasta vazia após achatamento",
                        "regra": "POP Quadro 2", "resultado": "REMOVIDA"})
    for g in R.config()["grupos"]:
        gdir = os.path.join(pasta, g)
        for a in (os.listdir(gdir) if os.path.isdir(gdir) else []):
            if R.norm(a) == "anexos" and a != R.config()["nome_pasta_anexos"]:
                tmp = os.path.join(gdir, a + "__tmp_case")
                os.rename(os.path.join(gdir, a), tmp)
                os.rename(tmp, os.path.join(gdir, R.config()["nome_pasta_anexos"]))
    # 5) REVALIDAR a pasta publicada contra o manifesto do FINAL
    inv_pub = INV.inventariar(pasta, cache)
    pub = {R.norm(it["rel"]): it["sha256"] for it in inv_pub["itens"] if it["classe"] != R.K_LIXO}
    fin = {R.norm(k): v for k, v in plano["hash_final"].items()}
    ok = pub == fin
    with open(os.path.join(out, "log_publicacao.json"), "w", encoding="utf-8") as f:
        json.dump({"executado_em": _agora(), "pasta": pasta, "arquivo_interno": arq_int, "backup_local": bkp,
                   "codigo": codigo, "operacoes": log, "conferencia_final": "IDÊNTICA AO MANIFESTO" if ok else
                   {"faltando": sorted(set(fin) - set(pub)), "sobrando": sorted(set(pub) - set(fin)),
                    "divergentes": sorted(k for k in set(fin) & set(pub) if fin[k] != pub[k])}}, f, ensure_ascii=False, indent=1)
    cache.salvar()
    _salvar_estado(out, etapa="PUBLICADO" if ok else "PUBLICAÇÃO DIVERGENTE", publicado_em=_agora(),
                   arquivo_interno=arq_int)
    print(f"PUBLICAÇÃO {'OK — pasta do mês idêntica ao manifesto do pacote final' if ok else 'COM DIVERGÊNCIA (ver log)'}")
    print(f"  originais arquivados em: {arq_int}\n  backup local: {bkp}\n  log: {os.path.join(out, 'log_publicacao.json')}")
    print("  Nada foi enviado ao Parceiro. A cópia para a pasta compartilhada com o Parceiro é feita pela usuária.")
    return 0 if ok else 3

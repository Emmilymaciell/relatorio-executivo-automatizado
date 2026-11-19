# -*- coding: utf-8 -*-
"""Atualização INCREMENTAL da entrega do G2 já organizada por organizar_g2_pop.py (depois do retorno do engenheiro).

Trabalha sobre o estado REAL de <pasta do mês>/G2/Anexos (fonte da verdade agora) e:
  1. renomeia as cópias cuja rodovia mudou por decisão registrada em config/g2_pop_ajustes.json (ex.: NT-793.1 GERAL -> RD-104);
  2. transforma cada ZIP 'Relatório Mensal de Obras - RM-0xx.1 - GOxxx - MM AAAA.zip' enviado pelo G2 em
     PDF assinado (byte a byte, assinatura preservada) + ZIP(s) 'Parte NN' <= 95 MB com a pasta Anexos;
     o ZIP original é MOVIDO (nunca apagado) para a pasta do layout antigo no Desenvolvimento;
  3. reescreve só a subdivisão GRUPO 2 do Índice a partir dos arquivos reais (backup antes);
  4. regrava o manifesto (plano_g2.json -> 'montagem.log') para a verificação.

  python scripts/manutencao/atualizar_g2_pop.py --mes 15            # dry-run: mostra o que faria
  python scripts/manutencao/atualizar_g2_pop.py --mes 15 --aplicar  # executa
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import shutil
import sys
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import organizar_g2_pop as O  # noqa: E402
from organizar_g2_pop import IDX, N, R, S, _jdump, _jload, _sha  # noqa: E402

_RE_RM_ZIP = re.compile(r"Relat[oó]rio Mensal de Obras - (RM-\d{3}\.\d) - (RD\s*-?\s*\d{3}\s*A?) - .*\.zip$", re.I)


def _cfg(mes):
    with open(os.path.join(O.PROJ, "config", "g2_pop_ajustes.json"), encoding="utf-8") as f:
        return json.load(f)[str(mes)]


def _partes_zip(membros, dst_dir, stem_sem_rod, rod, ler):
    """membros: [(nome_interno, ZipInfo)] na ordem original -> partes <= LIMITE_PARTE. Retorna (nomes, conteudo)."""
    grupos, atual, tam = [], [], 0
    for nome, zi in membros:
        if zi.file_size > O.LIMITE_PARTE:
            raise RuntimeError(f"arquivo interno sozinho > limite: {nome}")
        if atual and tam + zi.file_size > O.LIMITE_PARTE:
            grupos.append(atual)
            atual, tam = [], 0
        atual.append((nome, zi))
        tam += zi.file_size
    if atual:
        grupos.append(atual)
    nomes, conteudo = [], []
    for k, ms in enumerate(grupos, 1):
        nome_parte = f"{stem_sem_rod} - Parte {k:02d} - {rod}.zip" if len(grupos) > 1 else f"{stem_sem_rod} - {rod}.zip"
        dst = os.path.join(dst_dir, nome_parte)
        with zipfile.ZipFile(R.lp(dst), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zo:
            for nome, zi in ms:
                data = ler(zi)
                z2 = zipfile.ZipInfo(nome, date_time=zi.date_time)
                z2.compress_type = zipfile.ZIP_DEFLATED
                z2.external_attr = 0o644 << 16
                z2.flag_bits |= 0x800
                zo.writestr(z2, data)
                conteudo.append({"parte": nome_parte, "arquivo": nome, "origem_no_zip": nome, "bytes": len(data),
                                 "sha256": hashlib.sha256(data).hexdigest()})
        with zipfile.ZipFile(R.lp(dst)) as zt:
            if zt.testzip() is not None:
                raise RuntimeError(f"CRC inválido em {nome_parte}")
        nomes.append(nome_parte)
    return nomes, conteudo


def executar(mes, aplicar):
    P = O.caminhos(mes)
    cfg = _cfg(mes)
    plano = _jload(os.path.join(P["out"], "plano_g2.json"))
    an = P["anexos_sp"]
    rmcfg = cfg["relatorios_mensais_obra"]
    pasta_orig = os.path.join(O._cfg_caminhos()["prog_root"], *rmcfg["pasta_originais"].split("/"))
    prod = "P2.4"
    nome_prod = R.produtos()[prod]["nome"]
    lista = O.ler_lista_mestre(P["lista_mestre"], P["num"])
    acoes, log_extra, zips_rm = [], [], []
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # 1) renomeações por decisão registrada (rodovia mudou)
    for f in sorted(os.listdir(an)):
        if " - NT-793.1 - " in f and re.search(r" - GERAL\.(pdf|zip)$", f):
            novo = re.sub(r" - GERAL\.(pdf|zip)$", r" - RD-104.\1", f)
            acoes.append(("renomear", f, novo, "NT-793.1: engenheiro confirmou RD-104 (06/10/2026)"))
    # 2) ZIPs dos Relatórios Mensais de Obra
    for f in sorted(os.listdir(an)):
        m = _RE_RM_ZIP.search(f)
        if not m:
            continue
        doc = m.group(1)
        rod = O.rodovias_texto(m.group(2))
        rod = rod.pop() if len(rod) == 1 else None
        lm = [r for r in lista if r["doc"] == doc and O.rodovia_lista(r["empreendimento"]) == rod]
        if not rod or len(lm) != 1:
            raise SystemExit(f"[atualizar] {f}: {doc}/{rod} não casa com UMA linha da Lista Mestre")
        zips_rm.append({"zip": f, "doc": doc, "rod": rod, "lista_linha": lm[0]["linha"]})
        acoes.append(("desmembrar_rm", f, f"{prod} - {nome_prod} - {rmcfg['descricao_pdf']} - {doc} - {rod}.pdf "
                                          f"+ anexos em partes", "POP: PDF fora de ZIP; ≤ 100 MB por arquivo"))
        acoes.append(("mover_original", f, os.path.join(pasta_orig, f), "original do G2 preservado fora do pacote"))
    for a in acoes:
        print(f"  {a[0]}: {a[1]}  ->  {a[2]}   [{a[3]}]")
    if not aplicar:
        print("[atualizar] dry-run — nada alterado. Rode com --aplicar.")
        return

    # --- execução
    bk = os.path.join(P["out"], "backup_indice")
    os.makedirs(bk, exist_ok=True)
    shutil.copy2(R.lp(P["indice_sp"]), R.lp(os.path.join(bk, f"{ts}_{os.path.basename(P['indice_sp'])}")))
    for tipo, a, b, motivo in acoes:
        if tipo == "renomear":
            src, dst = os.path.join(an, a), os.path.join(an, b)
            if os.path.exists(R.lp(dst)):
                raise SystemExit(f"já existe: {b}")
            sha = _sha(src)
            os.rename(R.lp(src), R.lp(dst))
            assert _sha(dst) == sha
            log_extra.append({"renomeado_de": a, "para": b, "sha256": sha, "motivo": motivo})
    stg = os.path.join(P["out"], "STAGING_RM")
    if os.path.isdir(stg):
        shutil.rmtree(R.lp(stg))
    os.makedirs(R.lp(stg))
    rm_itens = []
    for z in zips_rm:
        src = os.path.join(an, z["zip"])
        sha_zip = _sha(src)
        with zipfile.ZipFile(R.lp(src)) as zf:
            infos = [i for i in zf.infolist() if not i.is_dir() and not R.eh_lixo(os.path.basename(i.filename))]
            raiz = [i for i in infos if "/" not in i.filename and i.filename.lower().endswith(".pdf")]
            if len(raiz) != 1:
                raise SystemExit(f"{z['zip']}: esperado 1 PDF na raiz do ZIP, achei {[i.filename for i in raiz]}")
            pdf_nome = f"{prod} - {nome_prod} - {rmcfg['descricao_pdf']} - {z['doc']} - {z['rod']}.pdf"
            data = zf.read(raiz[0])
            with open(R.lp(os.path.join(stg, pdf_nome)), "wb") as fo:
                fo.write(data)
            sig = S.verificar_pdf(R.lp(os.path.join(stg, pdf_nome)))
            if sig["status"] != S.VALIDADA:
                raise SystemExit(f"{pdf_nome}: assinatura {sig['status']} após extração — abortado")
            resto = [(i.filename, i) for i in infos if i is not raiz[0]]
            stem = f"{prod} - {nome_prod} - {rmcfg['descricao_anexos']} - {z['doc']}"
            partes, conteudo = _partes_zip(resto, stg, stem, z["rod"], zf.read)
            orig = {i.filename: hashlib.sha256(zf.read(i)).hexdigest() for i in infos}
        for c in conteudo:
            if orig[c["arquivo"]] != c["sha256"]:
                raise RuntimeError(f"conteúdo divergente: {c['arquivo']}")
        if sorted(orig) != sorted([raiz[0].filename] + [c["arquivo"] for c in conteudo]):
            raise RuntimeError(f"{z['zip']}: partes + PDF não reproduzem o ZIP original")
        det = sig["assinaturas"][0]
        rm_itens.append({"zip_original": z["zip"], "zip_sha256": sha_zip, "doc": z["doc"], "rodovia": z["rod"],
                         "pdf": pdf_nome, "pdf_sha256": hashlib.sha256(data).hexdigest(),
                         "assinatura": {"status": sig["status"], "signatario": det.get("signatario"),
                                        "emissor": det.get("emissor")},
                         "partes": partes, "conteudo_partes": conteudo, "lista_linha": z["lista_linha"]})
    # copiar staging -> SharePoint (nunca sobrescreve) e conferir
    for f in sorted(os.listdir(stg)):
        dst = os.path.join(an, f)
        if os.path.exists(R.lp(dst)):
            raise SystemExit(f"já existe no SharePoint: {f}")
        shutil.copy2(R.lp(os.path.join(stg, f)), R.lp(dst))
        if _sha(dst) != _sha(os.path.join(stg, f)):
            raise RuntimeError(f"SHA-256 divergente ao copiar {f}")
    # mover os ZIPs originais do G2 para fora do pacote
    os.makedirs(R.lp(pasta_orig), exist_ok=True)
    for z in rm_itens:
        src, dst = os.path.join(an, z["zip_original"]), os.path.join(pasta_orig, z["zip_original"])
        if os.path.exists(R.lp(dst)):
            raise SystemExit(f"destino já existe: {dst}")
        shutil.copy2(R.lp(src), R.lp(dst))
        if _sha(dst) != z["zip_sha256"]:
            raise RuntimeError(f"cópia do original divergente: {dst}")
        os.remove(R.lp(src))
        z["original_movido_para"] = dst

    # 3) Índice (só GRUPO 2) a partir dos arquivos reais
    rel = plano["relatorio"]["destino"]
    finais = sorted((os.path.splitext(f)[0] for f in os.listdir(an) if os.path.isfile(os.path.join(an, f))
                     and not R.eh_lixo(f)), key=lambda s: (s.split(" - ")[0], R.norm(s)))
    linhas = [os.path.splitext(rel)[0]] + finais
    base = os.path.join(P["out"], "indice_base_sharepoint.docx")
    shutil.copy2(R.lp(P["indice_sp"]), R.lp(base))
    novo = os.path.join(P["out"], "indice_novo.docx")
    O._escrever_g2_no_indice(base, novo, linhas)
    antes, _ = O._linhas_por_grupo(base)
    depois, _ = O._linhas_por_grupo(novo)
    if not all(antes.get(g) == depois.get(g) for g in ("G1", "G3", "G4", "G5")):
        raise RuntimeError("o Índice alteraria linhas de outro grupo — abortado")
    if _sha(P["indice_sp"]) != _sha(base):
        raise SystemExit("o Índice mudou durante a atualização — rode de novo")
    shutil.copy2(R.lp(novo), R.lp(P["indice_sp"]))

    # 4) manifesto atualizado (estado real)
    antigos = {e["destino"]: e for e in plano["montagem"]["log"]}
    ren = {f"G2/Anexos/{x['para']}": f"G2/Anexos/{x['renomeado_de']}" for x in log_extra}
    rm_por_nome = {}
    for z in rm_itens:
        for n in [z["pdf"]] + z["partes"]:
            rm_por_nome[n] = z
    novo_log = []
    for f in sorted(os.listdir(an)):
        p = os.path.join(an, f)
        if not os.path.isfile(p) or R.eh_lixo(f):
            continue
        dest = f"G2/Anexos/{f}"
        sha = _sha(p)
        ant = antigos.get(dest) or antigos.get(ren.get(dest, ""))
        if f in rm_por_nome:
            z = rm_por_nome[f]
            e = {"origem": f"G2/Anexos/{z['zip_original']} (enviado pelo G2 em 06/10/2026)", "destino": dest,
                 "op": "desmembrar_rm", "sha256": sha, "resultado": "OK"}
        elif ant:
            e = dict(ant, destino=dest, sha256=sha)
            if ant["sha256"] != sha:
                e["origem"] = ant["origem"] + " → substituído pelo G2 em 06/10/2026 (versão corrigida)"
                e["op"] = "substituido_pelo_g2"
            if dest in ren:
                e["op"] = e.get("op", "copiar") + "+renomeado"
        else:
            e = {"origem": "(arquivo colocado direto pelo G2)", "destino": dest, "op": "externo", "sha256": sha,
                 "resultado": "SEM ORIGEM RASTREADA"}
        novo_log.append(e)
    rp = os.path.join(P["g2_sp"], rel)
    novo_log.append({"origem": "assinado e colocado pelo G2 (Pessoa E) em 06/10/2026 13:28", "destino": f"G2/{rel}",
                     "op": "substituido_pelo_g2", "sha256": _sha(rp), "resultado": "OK"})
    plano.setdefault("historico", []).append({"em": datetime.datetime.now().isoformat(timespec="seconds"),
                                              "etapa": "atualizar", "renomeacoes": log_extra,
                                              "relatorios_mensais": [{k: v for k, v in z.items() if k != "conteudo_partes"}
                                                                     for z in rm_itens],
                                              "indice_linhas_g2": len(linhas)})
    plano["montagem"]["log"] = novo_log
    plano["montagem"]["indice_linhas_g2"] = linhas
    plano["montagem"]["outros_grupos_inalterados"] = True
    plano["relatorios_mensais"] = rm_itens
    _jdump(plano, os.path.join(P["out"], "plano_g2.json"))
    print(f"[atualizar] {len(log_extra)} renomeado(s); {len(rm_itens)} RM desmembrado(s) "
          f"({sum(len(z['partes']) for z in rm_itens)} parte(s)); Índice G2: {len(linhas)} linha(s); anexos: {len(novo_log) - 1}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mes", type=int, required=True)
    ap.add_argument("--aplicar", action="store_true")
    a = ap.parse_args()
    executar(a.mes, a.aplicar)

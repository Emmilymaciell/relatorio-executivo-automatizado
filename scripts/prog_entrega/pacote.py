# -*- coding: utf-8 -*-
"""STAGING (cópia fiel da origem, conferida por SHA-256) e FINAL (pacote no layout do POP).

Nada aqui escreve na pasta de origem. Toda transformação é registrada como
ANTES -> DEPOIS -> MOTIVO -> REGRA -> RESULTADO.
"""
import datetime
import os
import shutil
import stat
import zipfile

from . import hashcache as H
from . import indice as IDX
from . import nomes as N
from . import regras as R


def _limpar_dir_gerado(d):
    """Remove um diretório GERADO pela automação (output/...). Nunca chamado para pastas de origem."""
    if not os.path.isdir(d):
        return

    def _onerr(func, path, _exc):
        os.chmod(path, stat.S_IWRITE)
        func(path)
    shutil.rmtree(d, onexc=_onerr)


def montar_staging(inv, staging_root):
    _limpar_dir_gerado(staging_root)
    log = []
    for it in inv["itens"]:
        if it["classe"] == R.K_LIXO:
            continue
        dst = os.path.join(staging_root, *it["rel"].split("/"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(it["abs"], dst)
        ok = H.sha256_arquivo(dst) == it["sha256"]
        log.append({"antes": it["rel"], "depois": "STAGING/" + it["rel"], "motivo": "snapshot de evidência",
                    "regra": "nunca trabalhar sobre originais", "resultado": "OK" if ok else "FALHOU: hash divergente"})
        if not ok:
            raise RuntimeError(f"cópia para staging divergiu: {it['rel']}")
    return log


def _zip_det(dst, membros):
    """ZIP determinístico (mesma entrada => mesmos bytes): data do arquivo de origem, ordem fixa."""
    with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for nome_interno, src in membros:
            mt = datetime.datetime.fromtimestamp(os.path.getmtime(src))
            dt = (max(mt.year, 1980), mt.month, mt.day, mt.hour, mt.minute, mt.second - mt.second % 2)
            zi = zipfile.ZipInfo(nome_interno, date_time=dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            with open(src, "rb") as f:
                z.writestr(zi, f.read())
    with zipfile.ZipFile(dst) as z:
        bad = z.testzip()
        if bad:
            raise RuntimeError(f"ZIP gerado com CRC inválido: {bad}")
        nomes = sorted(z.namelist())
    if nomes != sorted(n for n, _ in membros):
        raise RuntimeError("conteúdo do ZIP não corresponde aos arquivos de origem")


def _nome_parte(stem, k):
    blocos = stem.split(N.SEP)
    if len(blocos) >= 4 and N.canon_rodovia(blocos[-1]):
        return N.SEP.join(blocos[:-1] + [f"Parte {k:02d}", blocos[-1]])
    return f"{stem}{N.SEP}Parte {k:02d}"


def _fracionar_pdf(src, dst_dir, stem, lim):
    import fitz
    doc = fitz.open(src)
    partes, ini, k = [], 0, 1
    media = os.path.getsize(src) / max(doc.page_count, 1)
    alvo = lim * 0.9

    def _bytes(a, b):
        tmp = fitz.open()
        tmp.insert_pdf(doc, from_page=a, to_page=b)
        data = tmp.tobytes(garbage=3, deflate=True)
        tmp.close()
        return data

    while ini < doc.page_count:
        # estimativa pelo tamanho médio da página; encolhe até caber (sem O(n²) de páginas)
        n = max(1, min(doc.page_count - ini, int(alvo // max(media, 1))))
        while True:
            data = _bytes(ini, ini + n - 1)
            if len(data) <= lim:
                break
            if n == 1:
                raise RuntimeError(f"a página {ini + 1} sozinha excede o limite — não fracionável automaticamente")
            n = max(1, int(n * 0.8))
        ultima = ini + n - 1
        nome = _nome_parte(stem, k) + ".pdf"
        with open(os.path.join(dst_dir, nome), "wb") as f:
            f.write(data)
        partes.append((nome, ini + 1, ultima + 1))
        ini, k = ultima + 1, k + 1
    total = sum(fitz.open(os.path.join(dst_dir, n)).page_count for n, _, _ in partes)
    if total != doc.page_count:
        raise RuntimeError("soma das páginas das partes difere do original")
    return partes


def _reempacotar_zip(src, dst_dir, stem, lim, tmp_dir):
    with zipfile.ZipFile(src) as z:
        membros = [m for m in z.infolist() if not m.is_dir()]
        grupos, atual, tam = [], [], 0
        for m in sorted(membros, key=lambda m: m.filename):
            if m.compress_size > lim:
                raise RuntimeError(f"membro '{m.filename}' sozinho excede o limite")
            if atual and tam + m.compress_size > lim * 0.95:
                grupos.append(atual)
                atual, tam = [], 0
            atual.append(m)
            tam += m.compress_size
        if atual:
            grupos.append(atual)
        saidas = []
        for k, ms in enumerate(grupos, 1):
            nome = _nome_parte(stem, k) + ".zip"
            extra = []
            for m in ms:
                p = os.path.join(tmp_dir, f"{k}", m.filename)
                os.makedirs(os.path.dirname(p), exist_ok=True)
                with z.open(m) as fi, open(p, "wb") as fo:
                    shutil.copyfileobj(fi, fo)
                extra.append((m.filename, p))
            _zip_det(os.path.join(dst_dir, nome), extra)
            saidas.append(nome)
    return saidas


def montar_final(plano, staging_root, final_root, info_mes):
    """Monta o FINAL a partir do STAGING conforme o plano. Retorna (log de transformações, entradas_indice)."""
    _limpar_dir_gerado(final_root)
    os.makedirs(final_root)
    lim = R.limite_bytes()
    log, entradas = [], {g: [] for g in R.config()["grupos"]}
    tmp = os.path.join(os.path.dirname(final_root), "_tmp_reempacote")
    for e in plano["final"]:
        if e["op"] == "excluir":
            continue
        src = os.path.join(staging_root, *e["origem"].split("/"))
        dst = os.path.join(final_root, *e["destino"].split("/"))
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        stem = os.path.splitext(os.path.basename(dst))[0]
        reg = {"antes": e["origem"], "operacao": e["op"], "regra": "", "motivo": ""}
        try:
            if os.path.exists(dst):
                raise RuntimeError("destino já existe no FINAL (colisão) — não sobrescrito")
            if e["op"] == "copiar":
                shutil.copy2(src, dst)
                if H.sha256_arquivo(dst) != e["sha256"]:
                    raise RuntimeError("hash divergente após cópia")
                gerados = [os.path.basename(dst)]
                reg.update(motivo="arquivo conforme (cópia fiel)" if os.path.basename(e["origem"]) == os.path.basename(dst)
                           else "renomeado/achatado para o padrão", regra="POP Quadros 2 e 3")
            elif e["op"] == "zipar":
                _zip_det(dst, [(e["nome_interno"], src)])
                gerados = [os.path.basename(dst)]
                reg.update(motivo="não-PDF compactado em ZIP", regra="POP Quadro 3 / Etapa 2")
            elif e["op"] == "fracionar_pdf":
                partes = _fracionar_pdf(src, os.path.dirname(dst), stem, lim)
                gerados = [p[0] for p in partes]
                reg.update(motivo=f"PDF > {R.config()['limite_tamanho_mb']} MB fracionado: "
                                  + "; ".join(f"{n} (p. {a}-{b})" for n, a, b in partes), regra="POP Quadro 4")
            elif e["op"] == "reempacotar_zip":
                gerados = _reempacotar_zip(src, os.path.dirname(dst), stem, lim, tmp)
                reg.update(motivo="ZIP > limite reempacotado em partes", regra="POP Quadro 4")
            else:
                raise RuntimeError(f"operação desconhecida {e['op']}")
            reg["depois"] = [f"{os.path.dirname(e['destino'])}/{n}" for n in gerados]
            reg["resultado"] = "OK"
            for n in gerados:
                entradas[e["grupo"]].append((e.get("produto") or "?", n, e["papel"]))
        except Exception as ex:
            reg["depois"] = None
            reg["resultado"] = f"FALHOU: {ex}"
        log.append(reg)
    _limpar_dir_gerado(tmp)
    # pastas ANEXOS vazias também existem no pacote (estrutura do POP)
    for g in R.config()["grupos"]:
        os.makedirs(os.path.join(final_root, g, R.config()["nome_pasta_anexos"]), exist_ok=True)
    # índice regenerado: relatório primeiro, depois anexos por produto e nome
    ordem_prod = list(R.produtos())
    ent_idx = {}
    for g, lst in entradas.items():
        rel = [("REL", n) for p, n, papel in lst if papel == "relatorio"]
        anx = sorted([(p, n) for p, n, papel in lst if papel != "relatorio"],
                     key=lambda x: (ordem_prod.index(x[0]) if x[0] in ordem_prod else 999, R.norm(x[1])))
        ent_idx[g] = rel + anx
    _num, mes, mes_num, ano = info_mes
    nome_idx = IDX.nome_arquivo_indice(mes, ano)
    n = IDX.escrever_indice(os.path.join(final_root, nome_idx), mes, ano, ent_idx,
                            periodo=R.periodo_do_mes(mes_num, ano), sei_por_arquivo=plano.get("sei_por_arquivo"))
    log.append({"antes": "(índices da origem)", "operacao": "gerar_indice", "depois": [nome_idx],
                "motivo": f"índice regenerado com {n} arquivo(s) reais, na ordem G1→G5", "regra": "POP Quadro 2 + Modelo",
                "resultado": "OK"})
    return log, ent_idx

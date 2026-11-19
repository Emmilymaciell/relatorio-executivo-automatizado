# -*- coding: utf-8 -*-
"""Etapas planejar / aplicar / verificar de reorganizar_mes_pop.py (DEC-070). Ver o docstring de lá."""
import collections
import datetime
import hashlib
import io
import json
import os
import shutil
import tempfile
import zipfile

from reorganizar_mes_pop import DEST_SNAPSHOT, PROJ, _jdump, _jload, _sha

from prog_entrega import assinatura as S
from prog_entrega import indice as IDX
from prog_entrega import nomes as N
from prog_entrega import regras as R

LIMITE_PARTE = 95 * 1000 * 1000
LIMITE_ARQUIVO = 100 * 1000 * 1000
CAMINHO_MAX = 250


def _cfg(mes):
    with open(os.path.join(PROJ, "config", "organizacao_pop.json"), encoding="utf-8") as f:
        return json.load(f)[str(mes)]


def _espelho_mes(c):
    raiz = os.environ.get("PROG_ROOT")
    if not raiz:
        raise SystemExit("defina PROG_ROOT (espelho de leitura em dia — sharepoint.py comparar-pasta)")
    return os.path.join(raiz, "Programa - 4. Produtos", c["pasta"])


def _inventario(base):
    inv = {}
    for dp, _, fs in os.walk(R.lp(base)):
        for f in fs:
            p = os.path.join(dp, f)
            rel = os.path.relpath(p, R.lp(base)).replace("\\", "/")
            inv[rel] = {"tamanho": os.path.getsize(p), "sha256": _sha(p), "abs": p}
    return inv


# ------------------------------------------------------------------------------------- arquivos novos
def _membros(zf):
    """[(nome_no_pacote, ler(), tamanho)] na ordem original; ZIP interno aberto em pasta de mesmo nome."""
    """Recursivo: ZIP dentro de ZIP (em qualquer nível) vira pasta de mesmo nome — POP: sem ZIP dentro de ZIP."""
    out = []

    def walk(z, prefixo):
        for m in z.infolist():
            if m.is_dir() or R.eh_lixo(os.path.basename(m.filename)):
                continue
            nome = prefixo + m.filename
            if m.filename.lower().endswith(".zip"):
                walk(zipfile.ZipFile(io.BytesIO(z.read(m))), nome[:-4] + "/")
            else:
                out.append((nome, (lambda z=z, m=m: z.read(m)), m.file_size, m.date_time))
    walk(zf, "")
    return out


def _gravar_zip(dst, membros):
    with zipfile.ZipFile(R.lp(dst), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zo:
        for nome, ler, _, dt in membros:
            zi = zipfile.ZipInfo(nome, date_time=dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            zi.flag_bits |= 0x800
            zo.writestr(zi, ler())
    with zipfile.ZipFile(R.lp(dst)) as zt:
        if zt.testzip() is not None:
            raise RuntimeError(f"CRC inválido: {dst}")


def _desmembrar(src, item, dst_root):
    """PDF da raiz do ZIP -> arquivo próprio (byte a byte); resto -> ZIP(s) <= 95 MB. Prova: cada arquivo interno das
    saídas tem o SHA-256 do original e o conjunto é exatamente o do ZIP de origem."""
    saidas = []
    with zipfile.ZipFile(R.lp(src)) as zf:
        membros = _membros(zf)
        raiz = [m for m in membros if "/" not in m[0] and m[0].lower().endswith(".pdf")]
        if len(raiz) != 1:
            raise RuntimeError(f"{src}: esperado 1 PDF na raiz do ZIP, achei {[m[0] for m in raiz]}")
        pdf_dst = os.path.join(dst_root, *item["pdf"].split("/"))
        os.makedirs(R.lp(os.path.dirname(pdf_dst)), exist_ok=True)
        dados_pdf = raiz[0][1]()
        with open(R.lp(pdf_dst), "wb") as fo:
            fo.write(dados_pdf)
        assin = None
        if item.get("assinado"):
            sig = S.verificar_pdf(R.lp(pdf_dst))
            if sig["status"] != S.VALIDADA:
                raise RuntimeError(f"{item['pdf']}: assinatura {sig['status']} após extração")
            a0 = sig["assinaturas"][0]
            assin = {"status": sig["status"], "signatario": a0.get("signatario"), "emissor": a0.get("emissor")}
        saidas.append({"rel": item["pdf"], "abs": pdf_dst, "conteudo": [raiz[0][0]], "assinatura": assin})
        resto = [m for m in membros if m is not raiz[0]]
        for m in resto:
            if m[2] > LIMITE_PARTE:
                raise RuntimeError(f"arquivo interno sozinho > 95 MB: {m[0]}")
        grupos, atual, tam = [], [], 0
        for m in resto:
            if atual and tam + m[2] > LIMITE_PARTE:
                grupos.append(atual)
                atual, tam = [], 0
            atual.append(m)
            tam += m[2]
        if atual:
            grupos.append(atual)
        orig = {m[0]: hashlib.sha256(m[1]()).hexdigest() for m in membros}
        for k, ms in enumerate(grupos, 1):
            sufixo = f" - Parte {k:02d}" if len(grupos) > 1 else ""
            rel = f"{item['anexos_stem']}{sufixo} - {item['rodovia']}.zip"
            dst = os.path.join(dst_root, *rel.split("/"))
            _gravar_zip(dst, ms)
            with zipfile.ZipFile(R.lp(dst)) as zt:
                for n in zt.namelist():
                    if hashlib.sha256(zt.read(n)).hexdigest() != orig[n]:
                        raise RuntimeError(f"conteúdo divergente em {rel}: {n}")
            saidas.append({"rel": rel, "abs": dst, "conteudo": [m[0] for m in ms]})
    if sorted(orig) != sorted(n for s in saidas for n in s["conteudo"]):
        raise RuntimeError(f"{src}: as saídas não reproduzem exatamente o conteúdo do ZIP")
    return saidas


def _zip_individual(src, item, dst_root):
    dst = os.path.join(dst_root, *item["zip"].split("/"))
    os.makedirs(R.lp(os.path.dirname(dst)), exist_ok=True)
    with open(R.lp(src), "rb") as f:
        dados = f.read()
    mt = datetime.datetime.fromtimestamp(os.path.getmtime(R.lp(src))).timetuple()[:6]
    _gravar_zip(dst, [(item["interno"], lambda: dados, len(dados), mt)])
    with zipfile.ZipFile(R.lp(dst)) as zt:
        if zt.namelist() != [item["interno"]] or hashlib.sha256(zt.read(item["interno"])).hexdigest() != \
                hashlib.sha256(dados).hexdigest():
            raise RuntimeError(f"ZIP individual divergente: {item['zip']}")
    return [{"rel": item["zip"], "abs": dst, "conteudo": [item["interno"]]}]


# -------------------------------------------------------------------------------------------- planejar
def _papel(rel):
    partes = rel.split("/")
    if len(partes) == 1:
        return "indice" if partes[0].lower().startswith(("índice", "indice")) else "solto_raiz"
    g = partes[0]
    if len(partes) == 2:
        n = R.norm(partes[1])
        if partes[1].lower().endswith((".xlsx", ".xls")) and "lista de demandas" in n:
            return "lista"
        if partes[1].lower().endswith((".doc", ".docx")):
            return "word"
        if partes[1].lower().endswith(".pdf") and any(p in n for p in R.config()["padroes_relatorio"]):
            return "relatorio"
        return "solto_grupo"
    if len(partes) == 3 and partes[1] == "ANEXOS":
        return "anexo"
    return "subpasta"


def planejar(mes, contexto, confirmar=None):
    c = contexto(mes)
    cfg = _cfg(mes)
    base = _espelho_mes(c)
    inv = _inventario(base)
    print(f"[planejar] espelho: {len(inv)} arquivo(s) em {base}")
    estado = {rel: {"origem": rel, "tipo": "existente", "tamanho": v["tamanho"], "sha256": v["sha256"]}
              for rel, v in inv.items()}
    ops_mover, ops_retirar, ops_enviar, problemas = [], [], [], []
    for r in cfg["renomear"]:
        if r["de"] not in estado:
            problemas.append(f"origem inexistente: {r['de']}")
            continue
        if r["para"] in estado:
            problemas.append(f"destino já ocupado: {r['para']}")
            continue
        estado[r["para"]] = estado.pop(r["de"])
        ops_mover.append({"de": r["de"], "para": r["para"], "motivo": r["motivo"], "sha256": inv[r["de"]]["sha256"],
                          "tamanho": inv[r["de"]]["tamanho"]})
    novos_dir = os.path.join(c["out"], "NOVOS")
    if os.path.isdir(novos_dir):
        shutil.rmtree(R.lp(novos_dir))
    for it in cfg["novos"]:
        src = os.path.join(PROJ, *it["externo"].split("/")) if it.get("externo") else inv[it["de"]]["abs"]
        saidas = _desmembrar(src, it, novos_dir) if it["tipo"] == "desmembrar_zip" else _zip_individual(src, it, novos_dir)
        for s in saidas:
            if s["rel"] in estado:
                problemas.append(f"arquivo novo colide com existente: {s['rel']}")
            sha = _sha(s["abs"])
            estado[s["rel"]] = {"origem": it.get("de") or it.get("externo"), "tipo": "novo", "tamanho": os.path.getsize(
                R.lp(s["abs"])), "sha256": sha}
            ops_enviar.append({"local": s["abs"], "para": s["rel"], "sha256": sha, "tamanho": estado[s["rel"]]["tamanho"],
                               "origem": it.get("de") or it.get("externo"), "conteudo": s["conteudo"],
                               "assinatura": s.get("assinatura"), "motivo": it["motivo"]})
    pastas_retirar = []
    for r in cfg["retirar"]:
        if r.get("pasta"):
            pastas_retirar.append(r)
            continue
        if r["de"] not in estado:
            problemas.append(f"retirar: inexistente {r['de']}")
            continue
        estado.pop(r["de"])
        ops_retirar.append({"de": r["de"], "para_snapshot": f"{c['pasta']} - Retirados do pacote/{r['de']}",
                            "motivo": r["motivo"], "sha256": inv[r["de"]]["sha256"], "tamanho": inv[r["de"]]["tamanho"]})
    for r in pastas_retirar:
        sobra = [k for k in estado if k.startswith(r["de"] + "/")]
        if sobra:
            problemas.append(f"pasta a retirar ainda teria arquivos: {r['de']} -> {sobra[:3]}")
        ops_retirar.append({"de": r["de"], "pasta": True, "para_snapshot": f"{c['pasta']} - Retirados do pacote/{r['de']}",
                            "motivo": r["motivo"]})

    # ---- conferência do estado final
    final = {}
    nomes_norm = collections.Counter()
    for rel, v in sorted(estado.items()):
        papel = _papel(rel)
        nome = rel.split("/")[-1]
        a = {"papel": papel, **v, "problemas": []}
        if papel in ("subpasta", "solto_raiz", "solto_grupo"):
            a["problemas"].append(["BLOCKER", f"fora do layout do POP ({papel})"])
        if papel == "anexo":
            an = N.analisar(nome)
            a["problemas"] += [list(p) for p in an["problemas"]]
            g = rel.split("/")[0]
            if an["codigo"] and R.produtos().get(an["codigo"], {}).get("grupo") != g:
                a["problemas"].append(["BLOCKER", f"produto {an['codigo']} não é do {g}"])
            if an["codigo"] in R.produtos() and an["produto"] != R.produtos()[an["codigo"]]["nome"] and not (
                    an["codigo"] == "P1.4" and an["produto"] == "TIC"):
                a["problemas"].append(["WARNING", f"produto '{an['produto']}' ≠ catálogo '{R.produtos()[an['codigo']]['nome']}'"])
        if v["tamanho"] > LIMITE_ARQUIVO:
            a["problemas"].append(["BLOCKER", f"arquivo > 100 MB ({v['tamanho']})"])
        caminho = f"4. Produtos/{c['pasta']}/{rel}"
        caminho_parceiro = f"5. Compartilhado Parceiro/05. Produtos/{c['pasta']}/{rel}"
        a["caminho_len"] = max(len(caminho), len(caminho_parceiro))
        if a["caminho_len"] > CAMINHO_MAX:
            a["problemas"].append(["WARNING", f"caminho com {a['caminho_len']} caracteres (> {CAMINHO_MAX})"])
        nomes_norm[R.norm(rel)] += 1
        final[rel] = a
    for rel in final:
        if nomes_norm[R.norm(rel)] > 1:
            final[rel]["problemas"].append(["BLOCKER", "colisão de nome (ignorando acento/caixa)"])
    por_sha = collections.defaultdict(list)
    for rel, a in final.items():
        if a["papel"] == "anexo":
            por_sha[a["sha256"]].append(rel)
    dups = [v for v in por_sha.values() if len(v) > 1]

    # ---- Índice completo (G1–G5) a partir do estado final
    entradas = {}
    for g in R.config()["grupos"]:
        rels = [r for r in final if r.startswith(g + "/")]
        reps = sorted(r.split("/")[-1] for r in rels if final[r]["papel"] == "relatorio")
        reps += [x for x in cfg.get("indice_extra", {}).get(g, []) if x not in reps]
        anex = sorted((r.split("/")[-1] for r in rels if final[r]["papel"] == "anexo"),
                      key=lambda s: (s.split(" - ")[0], R.norm(s)))
        entradas[g] = [("REL", x) for x in reps] + [(x.split(" - ")[0], x) for x in anex]
    nome_idx = IDX.nome_arquivo_indice(c["mes"].capitalize(), c["ano"])
    idx_local = os.path.join(c["out"], nome_idx)
    import calendar
    mnum = R.parse_pasta_mes(c["pasta"])[2]
    ult = calendar.monthrange(c["ano"], mnum)[1]
    total_idx = IDX.escrever_indice(idx_local, c["mes"].capitalize(), c["ano"], entradas,
                                    periodo=(f"01/{mnum:02d}/{c['ano']}", f"{ult}/{mnum:02d}/{c['ano']}"))
    lido = IDX.ler_indice(idx_local)
    linhas_idx = [(l["grupo"], l["arquivo"]) for l in lido["linhas"]]
    esperado = [(g, x) for g in entradas for _, x in entradas[g]]
    if linhas_idx != esperado:
        problemas.append("Índice gerado não relê igual ao planejado")
    anexos_e_rel = {(r.split("/")[0], r.split("/")[-1]) for r, a in final.items() if a["papel"] in ("anexo", "relatorio")}
    sem_linha = sorted(anexos_e_rel - set(linhas_idx))
    sem_arquivo = sorted(set(linhas_idx) - anexos_e_rel)

    cont = collections.Counter(s for a in final.values() for s, _ in a["problemas"])
    plano = {"gerado_em": datetime.datetime.now().isoformat(timespec="seconds"), "pasta": c["pasta"],
             "mes_rem": c["mes_rem"], "snapshot_rem": c["snap_rem"], "espelho": base,
             "enviar": ops_enviar, "mover": ops_mover, "retirar": ops_retirar,
             "indice": {"local": idx_local, "remoto": f"{c['mes_rem']}/{nome_idx}", "sha256": _sha(idx_local),
                        "linhas": total_idx, "por_grupo": {g: len(v) for g, v in entradas.items()},
                        "linhas_sem_arquivo": sem_arquivo, "arquivos_sem_linha": sem_linha},
             "final": final, "problemas_plano": problemas, "achados_final": dict(cont), "duplicatas_conteudo": dups}
    corpo = json.dumps([ops_enviar, ops_mover, ops_retirar, plano["indice"]["sha256"]], sort_keys=True, default=str)
    plano["codigo"] = hashlib.sha256(corpo.encode()).hexdigest()[:10]
    _jdump(plano, os.path.join(c["out"], "plano_reorganizacao.json"))
    print(f"[planejar] enviar {len(ops_enviar)} | mover/renomear {len(ops_mover)} | retirar {len(ops_retirar)} | "
          f"Índice {total_idx} linha(s) {plano['indice']['por_grupo']}")
    print(f"[planejar] estado final: {len(final)} arquivo(s); achados {dict(cont)}; duplicatas {dups}")
    print(f"[planejar] Índice: linhas sem arquivo {sem_arquivo}; arquivos sem linha {sem_linha}")
    for rel, a in final.items():
        for s, m in a["problemas"]:
            print(f"   {s}: {rel} — {m}")
    if problemas:
        print("[planejar] PROBLEMAS DO PLANO:", *problemas, sep="\n   ")
    print(f"[planejar] CÓDIGO para aplicar: {plano['codigo']}")
    return plano


# --------------------------------------------------------------------------------------------- aplicar
def aplicar(mes, contexto, confirmar=None):
    c = contexto(mes)
    sp = c["sp"]
    pp = os.path.join(c["out"], "plano_reorganizacao.json")
    plano = _jload(pp)
    if not confirmar or confirmar != plano["codigo"]:
        raise SystemExit(f"código não confere (esperado o do último planejar: {plano['codigo']})")
    if plano["problemas_plano"] or plano["achados_final"].get("BLOCKER"):
        raise SystemExit("plano com problemas/BLOCKER — corrija e planeje de novo")
    snap = sp.item(c["snap_rem"])
    if not snap or "folder" not in snap:
        raise SystemExit("sem a cópia 'como enviado' no Desenvolvimento — rode snapshot primeiro")
    perm = [c["mes_rem"], DEST_SNAPSHOT]
    log = []

    def reg(**k):
        log.append(dict(k, quando=datetime.datetime.now().isoformat(timespec="seconds")))
        _jdump(dict(plano, aplicacao=log), pp)

    motivo = f"Reorganização POP de {c['pasta']} (DEC-070)"
    for e in plano["enviar"]:
        rem = f"{c['mes_rem']}/{e['para']}"
        r = sp.enviar(e["local"], rem, aplicar=True, motivo=motivo) if not sp.item(rem) else {"acao": "JA_EXISTE"}
        if r.get("acao") == "JA_EXISTE":
            arqs = sp.item(rem)
            ok = arqs.get("size") == e["tamanho"]
            reg(op="enviar", para=e["para"], resultado="já existia" + ("" if ok else " (TAMANHO DIFERENTE)"))
            if not ok:
                raise SystemExit(f"já existe arquivo diferente: {rem}")
            continue
        st = (r.get("verificacao") or {}).get("status")
        reg(op="enviar", para=e["para"], resultado=st)
        if st not in ("OK", "OK_CONTEUDO"):
            raise SystemExit(f"envio não conferiu: {rem} -> {st}")
    for e in plano["mover"]:
        o, d = f"{c['mes_rem']}/{e['de']}", f"{c['mes_rem']}/{e['para']}"
        if not sp.item(o) and sp.item(d):
            reg(op="mover", de=e["de"], para=e["para"], resultado="já feito")
            continue
        r = sp.mover(o, d, perm, aplicar=True, motivo=f"{motivo}: {e['motivo']}")
        reg(op="mover", de=e["de"], para=e["para"], resultado=r.get("verificacao") or r.get("acao"))
        if r.get("verificacao") != "OK":
            raise SystemExit(f"mover falhou: {e['de']} -> {r}")
    for e in plano["retirar"]:
        o, d = f"{c['mes_rem']}/{e['de']}", f"{DEST_SNAPSHOT}/{e['para_snapshot']}"
        if not sp.item(o) and sp.item(d):
            reg(op="retirar", de=e["de"], resultado="já feito")
            continue
        if e.get("pasta"):
            arqs, _ = sp.arvore(o)
            if arqs:
                raise SystemExit(f"pasta a retirar não está vazia: {o} ({list(arqs)[:3]})")
        r = sp.mover(o, d, perm, aplicar=True, motivo=f"{motivo}: retirado do pacote — {e['motivo']}")
        reg(op="retirar", de=e["de"], para=d, resultado=r.get("verificacao") or r.get("acao"))
        if r.get("verificacao") != "OK":
            raise SystemExit(f"retirar falhou: {e['de']} -> {r}")
    ix = plano["indice"]
    r = sp.enviar(ix["local"], ix["remoto"], substituir=True, aplicar=True, motivo=f"{motivo}: Índice completo G1–G5")
    reg(op="indice", para=ix["remoto"], resultado=(r.get("verificacao") or {}).get("status"),
        backup=(r.get("backup") or {}).get("destino"))
    c2 = collections.Counter(x["resultado"] for x in log)
    print(f"[aplicar] concluído: {dict(c2)}")


# ------------------------------------------------------------------------------------------- verificar
def verificar(mes, contexto, confirmar=None):
    c = contexto(mes)
    sp = c["sp"]
    plano = _jload(os.path.join(c["out"], "plano_reorganizacao.json"))
    arqs, pastas = sp.arvore(c["mes_rem"])
    final = plano["final"]
    falta = sorted(set(final) - set(arqs))
    sobra = sorted(set(arqs) - set(final))
    tam = sorted(k for k in set(final) & set(arqs) if final[k]["tamanho"] != arqs[k]["size"]
                 and not k.lower().endswith(".docx"))
    with tempfile.TemporaryDirectory() as t:
        b = sp.baixar(plano["indice"]["remoto"], os.path.join(t, "i.docx"))
        lido = IDX.ler_indice(b["destino"])
    linhas = {(x["grupo"], x["arquivo"]) for x in lido["linhas"]}
    reais = {(k.split("/")[0], k.split("/")[-1]) for k, a in final.items() if a["papel"] in ("anexo", "relatorio")}
    res = {"verificado_em": datetime.datetime.now().isoformat(timespec="seconds"), "arquivos_sharepoint": len(arqs),
           "arquivos_planejados": len(final), "faltando": falta, "sobrando": sobra, "tamanho_diferente": tam,
           "subpastas": sorted(pastas), "indice_linhas": len(lido["linhas"]),
           "indice_sem_arquivo": sorted(linhas - reais), "arquivos_sem_indice": sorted(reais - linhas)}
    res["resultado"] = "OK" if not (falta or sobra or tam or res["arquivos_sem_indice"]) and \
        set(res["subpastas"]) == {f"G{i}" for i in range(1, 6)} | {f"G{i}/ANEXOS" for i in range(1, 6)} else "DIVERGENTE"
    _jdump(res, os.path.join(c["out"], "verificacao_reorganizacao.json"))
    print(json.dumps(res, ensure_ascii=False, indent=1))

# -*- coding: utf-8 -*-
"""
fontes_externas.py — PRESERVAÇÃO DE INSUMOS EXTERNOS (DEC-059, regra permanente da usuária, 16/11/2026).

Todo arquivo que chega de FORA do ambiente canônico (Downloads, Desktop, anexo da conversa, pendrive, e-mail,
pasta temporária, arquivo de terceiro) e que vai alimentar QUALQUER resultado do projeto (relatório, timeline,
quadro financeiro, reuniões, dashboard, anexos, histórico, controles, auditoria) é PRIMEIRO importado:

  1. cópia byte a byte para  <work_dir da competência>/01_FONTES/EXTERNAS/<nome original>   (somente-leitura);
  2. registro de proveniência em  01_FONTES/EXTERNAS/_proveniencia.json  (nome e caminho de origem, data/hora
     da ingestão, competência, finalidade, classificação, SHA-256, tamanho, data de modificação da origem, obs.);
  3. cópia DURÁVEL e compartilhada no SharePoint:
     'Desenvolvimento de Produtos/Fontes Externas/<pasta do mês>/'  (original + _proveniencia.json),
     feita pela API (`sharepoint.py publicar-lote`, arquivo NOVO, nunca sobrescreve) — o plano é gerado aqui.

Só DEPOIS disso o arquivo é usado — sempre a partir da cópia importada, nunca da origem (que pode ser apagada).
O original importado é IMUTÁVEL: reimportar o mesmo conteúdo (mesmo SHA-256) é idempotente; um conteúdo NOVO com o
mesmo nome entra como '<nome> (2).<ext>' — nada é sobrescrito. Se o insumo também alimenta um artefato canônico
(acumulado, mestre de reuniões, Controle Timeline, Planilha de Medição), esse artefato é atualizado pelas regras
próprias dele (backup, só acrescentar etc.).

Uso:
  python fontes_externas.py inspecionar "<pasta de entrada>"          (classifica, não copia)
  python fontes_externas.py importar "<arquivo>" --finalidade "…" [--classe X] [--obs "…"]
  python fontes_externas.py listar
  python fontes_externas.py plano-sharepoint                           (gera manifesto p/ publicar-lote)
"""
import argparse, datetime, hashlib, json, os, re, shutil, stat, sys, unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# classe -> (palavras no NOME, palavras no CONTEÚDO). Ordem = prioridade. Sem casamento seguro => AMBIGUO.
CLASSES = [
    ("PRESTACAO_CONTAS", ["prestacao de contas", "prestacao contas"], ["prestacao de contas", "analise orgao contratante", "retorno programa"]),
    ("MEDICAO_PARCEIRO_XLSX", ["planilha de medicao", "medicao"], ["medicao", "valor acumulado"]),
    ("RELATORIO_GESTAO_CONTRATO", ["relatorio de gestao de contrato", "relatorio de gestao do contrato"],
     ["relatorio de gestao do contrato"]),
    ("REUNIOES", ["reunioes", "reuniao"], ["reuniao", "pauta", "participantes"]),
    ("DASHBOARD_POWERBI", ["dashboard"], ["analise de entrega"]),
    ("TIMELINE", ["timeline"], ["marco"]),
]


def _norm(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn").lower()


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def _amostra_texto(path, limite=6000):
    """Texto inicial do arquivo (xlsx: células; pdf: 1ª e última páginas; outros: vazio)."""
    low = path.lower()
    try:
        if low.endswith((".xlsx", ".xlsm")):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            out = []
            for ws in wb.worksheets[:3]:
                for row in ws.iter_rows(max_row=6, values_only=True):
                    out += [str(v) for v in row if v is not None]
            return " ".join(out)[:limite]
        if low.endswith(".pdf"):
            import fitz
            d = fitz.open(path)
            t = d[0].get_text() + " " + d[d.page_count - 1].get_text()
            return t.replace("​", "")[:limite]
    except Exception:
        return ""
    return ""


def classificar(path):
    """(classe, motivo). Nome e conteúdo precisam apontar a mesma classe; só um deles => classe com '?'
    (confirmar); nada => AMBIGUO."""
    nome = _norm(os.path.basename(path))
    txt = _norm(_amostra_texto(path))
    for classe, kn, kc in CLASSES:
        n_ok = any(k in nome for k in kn)
        c_ok = any(k in txt for k in kc)
        if n_ok and c_ok:
            return classe, "nome e conteúdo"
    for classe, kn, kc in CLASSES:
        if any(k in nome for k in kn) or any(k in txt for k in kc):
            return classe + "?", "só nome ou só conteúdo — confirmar"
    return "AMBIGUO", "sem correspondência — pendência (não importar sem decisão)"


def _dir_externas():
    import prog_config as cfg
    d = os.path.join(cfg.work_dir(), "01_FONTES", "EXTERNAS")
    os.makedirs(d, exist_ok=True)
    return d


def _registro_path():
    return os.path.join(_dir_externas(), "_proveniencia.json")


def registro():
    p = _registro_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {"_regra": "DEC-059 — preservação de insumos externos", "itens": []}


def _salvar(reg):
    with open(_registro_path(), "w", encoding="utf-8") as f:
        json.dump(reg, f, ensure_ascii=False, indent=1)


def importar(origem, finalidade, classe=None, obs="", competencia=None):
    """Copia a origem para 01_FONTES/EXTERNAS (somente-leitura) e registra a proveniência. Idempotente por SHA-256.
    Devolve o item do registro (com 'arquivo' = caminho da CÓPIA a usar)."""
    import prog_config as cfg
    origem = os.path.abspath(origem)
    if not os.path.isfile(origem):
        raise FileNotFoundError(origem)
    h = sha256(origem)
    reg = registro()
    for it in reg["itens"]:
        if it["sha256"] == h:
            it.setdefault("reimportacoes", []).append(datetime.datetime.now().isoformat(timespec="seconds"))
            _salvar(reg)
            return it
    dest_dir = _dir_externas()
    nome = os.path.basename(origem)
    dest = os.path.join(dest_dir, nome)
    k = 2
    while os.path.exists(dest):  # nunca sobrescreve um original importado
        base, ext = os.path.splitext(nome)
        dest = os.path.join(dest_dir, f"{base} ({k}){ext}")
        k += 1
    shutil.copy2(origem, dest)
    if sha256(dest) != h:
        os.remove(dest)
        raise IOError(f"cópia divergente (SHA-256) para {origem}")
    os.chmod(dest, stat.S_IREAD)  # imutável
    P = cfg.params()
    cls, motivo = (classe, "informada") if classe else classificar(origem)
    it = {
        "arquivo": dest,
        "nome_original": nome,
        "origem": origem,
        "ingestao": datetime.datetime.now().isoformat(timespec="seconds"),
        "modificado_na_origem": datetime.datetime.fromtimestamp(os.path.getmtime(origem)).isoformat(timespec="seconds"),
        "competencia": competencia or f"{P['mes']}/{P['ano']}",
        "relatorio": cfg.numero(),
        "finalidade": finalidade,
        "classe": cls,
        "classificacao": motivo,
        "sha256": h,
        "tamanho": os.path.getsize(dest),
        "obs": obs,
        "sharepoint": None,
    }
    reg["itens"].append(it)
    _salvar(reg)
    return it


def caminho_importado(classe):
    """Cópia importada mais recente de uma classe (ex.: 'PRESTACAO_CONTAS'), ou None."""
    its = [i for i in registro()["itens"] if i["classe"].rstrip("?") == classe]
    return its[-1]["arquivo"] if its else None


def plano_sharepoint():
    """Manifesto para `sharepoint.py publicar-lote` (arquivos NOVOS em Desenvolvimento de Produtos/Fontes Externas)."""
    import prog_config as cfg
    pasta = os.path.basename(cfg.month_folder())
    base = f"3. Desenvolvimento/G1 - Coordenacao/Desenvolvimento de Produtos/Fontes Externas/{pasta}"
    itens = [{"local": i["arquivo"], "remoto": f"{base}/{os.path.basename(i['arquivo'])}"}
             for i in registro()["itens"] if not i.get("sharepoint")]
    itens.append({"local": _registro_path(), "remoto": f"{base}/_proveniencia.json", "substituir": True})
    return {"motivo": f"DEC-059: preservação de insumos externos — {pasta}", "itens": itens}


def marcar_publicado(remoto_base):
    reg = registro()
    for i in reg["itens"]:
        if not i.get("sharepoint"):
            i["sharepoint"] = f"{remoto_base}/{os.path.basename(i['arquivo'])}"
    _salvar(reg)


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("inspecionar"); s.add_argument("pasta")
    s = sub.add_parser("importar"); s.add_argument("arquivo"); s.add_argument("--finalidade", required=True)
    s.add_argument("--classe"); s.add_argument("--obs", default="")
    sub.add_parser("listar")
    s = sub.add_parser("plano-sharepoint"); s.add_argument("--saida")
    a = ap.parse_args(argv)
    if a.cmd == "inspecionar":
        for f in sorted(os.listdir(a.pasta)):
            p = os.path.join(a.pasta, f)
            if os.path.isfile(p) and not f.startswith("~$"):
                c, m = classificar(p)
                print(f"{c:<28} {m:<40} {f}")
    elif a.cmd == "importar":
        print(json.dumps(importar(a.arquivo, a.finalidade, a.classe, a.obs), ensure_ascii=False, indent=1))
    elif a.cmd == "listar":
        for i in registro()["itens"]:
            print(f"{i['classe']:<26} {i['sha256'][:12]}  {i['nome_original']}  <- {i['origem']}")
    elif a.cmd == "plano-sharepoint":
        man = plano_sharepoint()
        out = a.saida or os.path.join(_dir_externas(), "_manifesto_sharepoint.json")
        with open(out, "w", encoding="utf-8") as f:
            json.dump(man, f, ensure_ascii=False, indent=1)
        print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())

# -*- coding: utf-8 -*-
"""Integração Graph/SharePoint (DEC-052) — SEM rede: Graph falsa em memória. Nada real é lido nem gravado."""
import io
import json
import os
import urllib.error
from urllib.parse import unquote

import pytest

from prog_sharepoint import caminhos as C
from prog_sharepoint import config as K
from prog_sharepoint.graph import Graph, ErroGraph
from prog_sharepoint.operacoes import SharePoint, partes_office_divergentes

CFG = {"hostname": "x.sharepoint.com", "site_path": "/sites/exemplo", "biblioteca": "Programa",
       "mapa_local_remoto": {"Programa - 3. Desenvolvimento": "3. Desenvolvimento", "Programa - 4. Produtos": "4. Produtos"},
       "prefixo_local": "Programa - ", "bloqueados_escrita": ["Compartilhado Parceiro"], "upload_simples_max_mb": 4,
       "backups_dir": "bk", "log_dir": "lg"}


class GraphFalsa:
    """Biblioteca em memória: {caminho_remoto: bytes}. Registra as escritas."""

    def __init__(self, arquivos=None):
        self.arquivos = dict(arquivos or {})
        self.escritas = []
        self.claims = {"roles": ["Sites.Selected"]}

    def token(self):
        return "t"

    def _caminho(self, url):
        p = url.split("/root", 1)[1].split("?")[0]
        if p.startswith(":/"):
            p = p[2:]
            p, _, resto = p.partition(":")
            return unquote(p), resto
        return "", p

    def get(self, url):
        if url.startswith("/sites/") and "/drives" not in url:
            return {"id": "S", "displayName": "Programa", "webUrl": "https://x/sites/exemplo"}
        p, resto = self._caminho(url)
        if p in self.arquivos:
            return {"name": p.split("/")[-1], "size": len(self.arquivos[p]), "eTag": f"e{len(self.escritas)}",
                    "file": {}, "@microsoft.graph.downloadUrl": "dl://" + p}
        if any(k.startswith(p + "/") for k in self.arquivos):
            return {"name": p.split("/")[-1], "folder": {}}
        raise ErroGraph(404, "itemNotFound", p)

    def paginar(self, url):
        if url.endswith("/drives"):
            return iter([{"id": "D", "name": "Documentos", "webUrl": "https://x/sites/exemplo/Programa"}])
        pasta, _ = self._caminho(url)
        filhos = {}
        for k, v in self.arquivos.items():
            if k.startswith(pasta + "/"):
                resto = k[len(pasta) + 1:].split("/")
                filhos[resto[0]] = {"name": resto[0], "folder": {}} if len(resto) > 1 else                     {"name": resto[0], "size": len(v), "file": {}}
        return iter(filhos.values())

    def baixar_url(self, url, fh):
        fh.write(self.arquivos[url[len("dl://"):]])

    def requisitar(self, metodo, url, corpo=None, cabecalhos=None, autenticar=True, bruto=False):
        p, _ = self._caminho(url)
        self.escritas.append((metodo, p, dict(cabecalhos or {}), url))
        self.arquivos[p] = corpo
        return {"webUrl": "https://x/" + p}


@pytest.fixture
def ambiente(tmp_path):
    raiz = tmp_path / "SP"
    (raiz / "Programa - 4. Produtos" / "15. Setembro 2026" / "G1").mkdir(parents=True)
    proj = tmp_path / "proj"
    proj.mkdir()
    return raiz, proj


def _sp(ambiente, g):
    raiz, proj = ambiente
    return SharePoint(CFG, g, str(raiz), str(proj))


# ------------------------------------------------------------------ caminhos
def test_mapeia_local_para_remoto(ambiente):
    raiz, _ = ambiente
    loc = raiz / "Programa - 3. Desenvolvimento" / "G1 - Coordenacao" / "Medição.xlsx"
    assert C.remoto_de_local(str(loc), str(raiz), CFG["mapa_local_remoto"]) == \
        "3. Desenvolvimento/G1 - Coordenacao/Medição.xlsx"
    # sem entrada no mapa: tira o prefixo 'Programa - '
    assert C.remoto_de_local(str(raiz / "Programa - 9. Outra" / "a.pdf"), str(raiz), {}) == "9. Outra/a.pdf"
    assert C.local_de_remoto("4. Produtos/x.pdf", str(raiz), CFG["mapa_local_remoto"]) == \
        str(raiz / "Programa - 4. Produtos" / "x.pdf")


def test_recusa_fora_da_raiz_e_ponto_ponto(ambiente, tmp_path):
    raiz, _ = ambiente
    with pytest.raises(C.CaminhoInvalido):
        C.remoto_de_local(str(tmp_path / "fora.xlsx"), str(raiz), CFG["mapa_local_remoto"])
    with pytest.raises(C.CaminhoInvalido):
        C.normalizar_remoto("4. Produtos/../segredo")


def test_codifica_segmentos():
    assert C.codificar("4. Produtos/a #1%/b.pdf") == "4.%20Produtos/a%20%231%25/b.pdf"


@pytest.mark.parametrize("alvo", ["5. Compartilhado Parceiro/x.pdf", "4. Produtos/5. COMPARTILHADO parceiro/x.pdf",
                                  "5. Compartilhado Parceiro/x.pdf"])
def test_bloqueia_compartilhado_parceiro_sem_rede(ambiente, tmp_path, alvo):
    f = tmp_path / "a.pdf"
    f.write_bytes(b"%PDF")
    g = GraphFalsa()
    with pytest.raises(C.CaminhoInvalido):
        _sp(ambiente, g).enviar(str(f), alvo, aplicar=True)
    assert g.escritas == []


# ------------------------------------------------------------------ credenciais
def test_credenciais_de_arquivo_e_env(tmp_path, monkeypatch):
    for k in K.CHAVES:
        monkeypatch.delenv(k, raising=False)
    arq = tmp_path / "sp.env"
    arq.write_text("# comentário\nPROG_SP_CLIENT_SECRET=abc123segredo\n", encoding="utf-8")
    monkeypatch.setenv("PROG_SP_CREDENCIAIS", str(arq))
    c = K.credenciais({"tenant_id": "T", "client_id": "C", "arquivo_credenciais": "x"})
    assert (c["tenant_id"], c["client_id"], c["client_secret"]) == ("T", "C", "abc123segredo")
    monkeypatch.setenv("PROG_SP_CLIENT_SECRET", "doenv")
    assert K.credenciais({"tenant_id": "T", "client_id": "C", "arquivo_credenciais": "x"})["client_secret"] == "doenv"
    assert "abc123segredo" not in K.mascarar("abc123segredo")


def test_credenciais_ausentes_ou_placeholder(tmp_path, monkeypatch):
    for k in K.CHAVES:
        monkeypatch.delenv(k, raising=False)
    arq = tmp_path / "sp.env"
    arq.write_text("PROG_SP_CLIENT_SECRET=COLE_AQUI_O_VALOR\n", encoding="utf-8")
    monkeypatch.setenv("PROG_SP_CREDENCIAIS", str(arq))
    with pytest.raises(K.ErroConfig):
        K.credenciais({"tenant_id": "T", "client_id": "C", "arquivo_credenciais": "x"})


def test_recusa_arquivo_de_credenciais_dentro_do_repo(monkeypatch):
    for k in K.CHAVES:
        monkeypatch.delenv(k, raising=False)
    dentro = os.path.join(K.RAIZ_PROJETO, "config", "sharepoint.json")  # qualquer arquivo existente no repo
    monkeypatch.setenv("PROG_SP_CREDENCIAIS", dentro)
    with pytest.raises(K.ErroConfig):
        K.credenciais({"tenant_id": "T", "client_id": "C", "arquivo_credenciais": "x"})


def test_config_versionada_sem_segredo():
    texto = open(K.ARQ_CONFIG, encoding="utf-8").read().lower()
    assert "client_secret\"" not in texto and "secret\":" not in texto
    c = K.carregar()
    assert c["hostname"].endswith(".sharepoint.com") and "Compartilhado Parceiro" in c["bloqueados_escrita"]


# ------------------------------------------------------------------ enviar
def test_enviar_dry_run_nao_escreve(ambiente, tmp_path):
    f = tmp_path / "Planilha.xlsx"
    f.write_bytes(b"PK novo")
    g = GraphFalsa()
    r = _sp(ambiente, g).enviar(str(f), "3. Desenvolvimento/G1/Planilha.xlsx")
    assert r["acao"] == "NOVO" and r["dry_run"] and not r["aplicado"]
    assert g.escritas == []


def test_enviar_existente_sem_substituir_bloqueia(ambiente, tmp_path):
    f = tmp_path / "T.xlsx"
    f.write_bytes(b"PK novo")
    g = GraphFalsa({"3. Desenvolvimento/G1/T.xlsx": b"PK antigo"})
    r = _sp(ambiente, g).enviar(str(f), "3. Desenvolvimento/G1/T.xlsx", aplicar=True)
    assert r["acao"] == "BLOQUEADO" and g.escritas == []
    assert g.arquivos["3. Desenvolvimento/G1/T.xlsx"] == b"PK antigo"


def test_enviar_novo_usa_conflict_fail_e_confere(ambiente, tmp_path):
    f = tmp_path / "n.pdf"
    f.write_bytes(b"%PDF novo")
    g = GraphFalsa()
    r = _sp(ambiente, g).enviar(str(f), "4. Produtos/15. Setembro 2026/G1/n.pdf", aplicar=True, motivo="teste")
    assert r["aplicado"] and r["verificacao"]["status"] == "OK"
    metodo, _, cab, url = g.escritas[0]
    assert metodo == "PUT" and "conflictBehavior=fail" in url and "If-Match" not in cab
    _, proj = ambiente
    linhas = open(next(os.path.join(proj, "lg", x) for x in os.listdir(proj / "lg")), encoding="utf-8").readlines()
    assert json.loads(linhas[-1])["motivo"] == "teste"


def test_substituir_faz_backup_e_if_match(ambiente, tmp_path):
    f = tmp_path / "T.xlsx"
    f.write_bytes(b"PK novo")
    alvo = "3. Desenvolvimento/G1/T.xlsx"
    g = GraphFalsa({alvo: b"PK antigo"})
    r = _sp(ambiente, g).enviar(str(f), alvo, substituir=True, aplicar=True)
    assert r["acao"] == "SUBSTITUIR" and r["aplicado"]
    bk = r["backup"]["destino"]
    assert open(bk, "rb").read() == b"PK antigo" and os.path.exists(bk + ".meta.json")
    metodo, _, cab, url = g.escritas[0]
    assert cab.get("If-Match") and "conflictBehavior=replace" in url
    assert g.arquivos[alvo] == b"PK novo"


# ------------------------------------------------------------------ cliente HTTP
class _Resp(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def test_graph_token_e_retry_429():
    chamadas = []

    def abrir(req, timeout=0):
        chamadas.append(req.full_url)
        if "oauth2" in req.full_url:
            return _Resp(json.dumps({"access_token": "a.eyJyb2xlcyI6WyJTaXRlcy5TZWxlY3RlZCJdfQ.s",
                                     "expires_in": 3600}).encode())
        if len([c for c in chamadas if "graph" in c]) == 1:
            raise urllib.error.HTTPError(req.full_url, 429, "x", {"Retry-After": "0"}, io.BytesIO(b"{}"))
        assert req.headers["Authorization"] == "Bearer a.eyJyb2xlcyI6WyJTaXRlcy5TZWxlY3RlZCJdfQ.s"
        return _Resp(b'{"ok": true}')

    g = Graph("T", "C", "S", abrir=abrir, dormir=lambda s: None)
    assert g.get("/sites/root") == {"ok": True}
    assert g.claims["roles"] == ["Sites.Selected"]


def test_graph_403_explica_permissao():
    def abrir(req, timeout=0):
        if "oauth2" in req.full_url:
            return _Resp(b'{"access_token": "a.e30.s", "expires_in": 3600}')
        raise urllib.error.HTTPError(req.full_url, 403, "x", {}, io.BytesIO(
            b'{"error": {"code": "accessDenied", "message": "Access denied"}}'))

    with pytest.raises(ErroGraph) as e:
        Graph("T", "C", "S", abrir=abrir).get("/sites/x")
    assert e.value.status == 403 and "Sites.Selected" in str(e.value)


# ------------------------------------------------------------------ Office e sincronia
def _xlsx(path, planilha=b"<sheet>1</sheet>", meta=b"<props/>", extra=None):
    import zipfile
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", b"<t/>" + (b"x" if extra else b""))
        z.writestr("docProps/core.xml", meta)
        z.writestr("xl/worksheets/sheet1.xml", planilha)
        z.writestr("xl/_rels/workbook.xml.rels", b"<r/>" + (b"c" if extra else b""))
        for n, v in (extra or {}).items():
            z.writestr(n, v)


def test_office_so_metadados_do_sharepoint_e_conteudo_igual(tmp_path):
    a, b, c = tmp_path / "a.xlsx", tmp_path / "b.xlsx", tmp_path / "c.xlsx"
    _xlsx(a)
    _xlsx(b, meta=b"<props editado/>", extra={"customXml/item1.xml": b"<x/>", "[trash]/0000.dat": b"z"})
    _xlsx(c, planilha=b"<sheet>2</sheet>")
    assert partes_office_divergentes(str(a), str(b)) == []
    assert partes_office_divergentes(str(a), str(c)) == ["xl/worksheets/sheet1.xml"]


def test_comparar_pasta_detecta_divergencias(ambiente):
    raiz, _ = ambiente
    loc = raiz / "Programa - 4. Produtos" / "15. Setembro 2026" / "G1"
    (loc / "igual.pdf").write_bytes(b"1234")
    (loc / "maior.pdf").write_bytes(b"123456")
    (loc / "so_local.pdf").write_bytes(b"x")
    (loc / "~$temp.xlsx").write_bytes(b"x")
    base = "4. Produtos/15. Setembro 2026/G1"
    g = GraphFalsa({f"{base}/igual.pdf": b"1234", f"{base}/maior.pdf": b"1234", f"{base}/so_sp.pdf": b"y"})
    r = _sp(ambiente, g).comparar_pasta(str(loc))
    assert r["iguais"] == 1 and not r["em_dia"]
    assert r["so_local"] == [f"{base}/so_local.pdf"] and r["so_sharepoint"] == [f"{base}/so_sp.pdf"]
    assert r["tamanho_diferente"][0]["item"] == f"{base}/maior.pdf"
    assert g.escritas == []


def test_publicar_lote_tudo_ou_nada(ambiente, tmp_path):
    a, b = tmp_path / "a.pdf", tmp_path / "b.pdf"
    a.write_bytes(b"%PDF a")
    b.write_bytes(b"%PDF b")
    base = "4. Produtos/15. Setembro 2026/G1"
    g = GraphFalsa({f"{base}/b.pdf": b"%PDF antigo"})
    sp = _sp(ambiente, g)
    man = {"motivo": "RM15", "itens": [{"local": str(a), "remoto": f"{base}/a.pdf"},
                                        {"local": str(b), "remoto": f"{base}/b.pdf"}]}
    r = sp.publicar_lote(man, aplicar=True)  # b existe e não tem substituir -> nada é enviado
    assert not r["pode_aplicar"] and not r["aplicado"] and g.escritas == []
    man["itens"][1]["substituir"] = True
    r = sp.publicar_lote(man, aplicar=True)
    assert r["aplicado"] and [p["acao"] for p in r["planos"]] == ["NOVO", "SUBSTITUIR"]
    assert g.arquivos[f"{base}/b.pdf"] == b"%PDF b"


def test_publicar_lote_recusa_parceiro_no_meio(ambiente, tmp_path):
    a = tmp_path / "a.pdf"
    a.write_bytes(b"%PDF")
    g = GraphFalsa()
    man = {"itens": [{"local": str(a), "remoto": "4. Produtos/x/a.pdf"},
                     {"local": str(a), "remoto": "5. Compartilhado Parceiro/a.pdf"}]}
    r = _sp(ambiente, g).publicar_lote(man, aplicar=True)
    assert r["erros"] and not r["aplicado"] and g.escritas == []

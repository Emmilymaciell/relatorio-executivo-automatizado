# -*- coding: utf-8 -*-
"""DEC-070 — mover/renomear, criar pasta e copiar pasta no servidor (Graph falsa, sem rede)."""
import hashlib
import os
import sys
from urllib.parse import unquote

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
from prog_sharepoint import caminhos as C  # noqa: E402
from prog_sharepoint.graph import ErroGraph  # noqa: E402
from prog_sharepoint.operacoes import SharePoint  # noqa: E402

CFG = {"hostname": "x", "site_path": "/sites/exemplo", "biblioteca": "Programa", "bloqueados_escrita": ["Compartilhado Parceiro"],
       "mapa_local_remoto": {}, "log_dir": "logs"}


class GraphItens:
    """Biblioteca em memória com ids/eTags: {caminho: bytes} para arquivos; pastas implícitas ou explícitas."""

    def __init__(self, arquivos):
        self.arq = dict(arquivos)
        self.pastas = set()
        for k in self.arq:
            partes = k.split("/")
            for i in range(1, len(partes)):
                self.pastas.add("/".join(partes[:i]))
        self.ids = {}
        self.escritas = []
        self.claims = {}

    def token(self):
        return "t"

    def _dormir(self, s):
        pass

    def _id(self, p):
        return self.ids.setdefault(p, "I" + hashlib.md5(p.encode()).hexdigest()[:8])

    def _por_id(self, i):
        return next(p for p, v in self.ids.items() if v == i and (p in self.arq or p in self.pastas))

    def _meta(self, p):
        if p in self.arq:
            return {"id": self._id(p), "name": p.split("/")[-1], "size": len(self.arq[p]), "eTag": "e",
                    "file": {"hashes": {"quickXorHash": hashlib.sha1(self.arq[p]).hexdigest()}}}
        if p in self.pastas or p == "":
            return {"id": self._id(p), "name": p.split("/")[-1], "folder": {}, "size": 0}
        raise ErroGraph(404, "itemNotFound", p)

    def _caminho(self, url):
        p = url.split("/root", 1)[1].split("?")[0]
        if p.startswith(":/"):
            p, _, resto = p[2:].partition(":")
            return unquote(p), resto
        return "", p

    def get(self, url):
        if url.startswith("/sites/") and "/drives" not in url:
            return {"id": "S"}
        return self._meta(self._caminho(url)[0])

    def paginar(self, url):
        if url.endswith("/drives"):
            return iter([{"id": "D", "name": "Programa", "webUrl": "https://x/sites/exemplo/Programa"}])
        pasta, _ = self._caminho(url)
        filhos = {p.split("/")[-1]: p for p in list(self.arq) + list(self.pastas)
                  if p.rsplit("/", 1)[0] == pasta and p != pasta and "/" in p}
        return iter(self._meta(p) for p in filhos.values())

    def requisitar(self, metodo, url, corpo=None, cabecalhos=None, autenticar=True, bruto=False, com_cabecalhos=False):
        self.escritas.append((metodo, url, corpo, dict(cabecalhos or {})))
        if metodo == "POST" and url.endswith("/children"):
            pai, _ = self._caminho(url)
            self.pastas.add(f"{pai}/{corpo['name']}" if pai else corpo["name"])
            return {}
        if metodo == "PATCH":
            o = self._por_id(url.split("/items/")[1].split("?")[0])
            pai = self._por_id(corpo["parentReference"]["id"]) if "parentReference" in corpo else o.rsplit("/", 1)[0]
            d = f"{pai}/{corpo['name']}"
            assert d not in self.arq and d not in self.pastas
            for k in [k for k in self.arq if k == o or k.startswith(o + "/")]:
                self.arq[d + k[len(o):]] = self.arq.pop(k)
            for k in [k for k in self.pastas if k == o or k.startswith(o + "/")]:
                self.pastas.discard(k)
                self.pastas.add(d + k[len(o):])
            return {}
        if metodo == "POST" and "/copy" in url:
            o = self._por_id(url.split("/items/")[1].split("/copy")[0])
            d = f"{self._por_id(corpo['parentReference']['id'])}/{corpo['name']}"
            for k in [k for k in self.arq if k.startswith(o + "/")]:
                self.arq[d + k[len(o):]] = self.arq[k]
            for k in [k for k in self.pastas if k == o or k.startswith(o + "/")]:
                self.pastas.add(d + k[len(o):])
            return ({}, {"Location": "https://monitor/1"}) if com_cabecalhos else {}
        if url.startswith("https://monitor"):
            return {"status": "completed"}
        raise AssertionError(f"inesperado: {metodo} {url}")


@pytest.fixture
def sp(tmp_path):
    g = GraphItens({"4. Produtos/16. Out/G3/ANEXOS/P3.1 - X/a.pdf": b"aaa",
                    "4. Produtos/16. Out/G3/ANEXOS/b.pdf": b"bb",
                    "4. Produtos/16. Out/G2/c.pdf": b"c"})
    return SharePoint(CFG, g, str(tmp_path / "SP"), str(tmp_path)), g


PERM = ["4. Produtos/16. Out"]


def test_mover_dry_run_nao_escreve(sp):
    s, g = sp
    r = s.mover("4. Produtos/16. Out/G3/ANEXOS/P3.1 - X/a.pdf", "4. Produtos/16. Out/G3/ANEXOS/P3.1 - novo - GERAL.pdf", PERM)
    assert r["acao"] == "MOVER" and r["dry_run"] and not g.escritas


def test_mover_renomeia_para_pasta_pai_e_confere_hash(sp):
    s, g = sp
    r = s.mover("4. Produtos/16. Out/G3/ANEXOS/P3.1 - X/a.pdf", "4. Produtos/16. Out/G3/ANEXOS/P3.1 - novo - GERAL.pdf",
                PERM, aplicar=True)
    assert r["verificacao"] == "OK" and not r["origem_existe"]
    assert g.arq["4. Produtos/16. Out/G3/ANEXOS/P3.1 - novo - GERAL.pdf"] == b"aaa"
    patch = [e for e in g.escritas if e[0] == "PATCH"][0]
    assert "conflictBehavior=fail" in patch[1] and patch[3].get("If-Match") == "e"
    assert "parentReference" in patch[2]


def test_mover_recusa_destino_existente_e_fora_do_permitido(sp):
    s, g = sp
    r = s.mover("4. Produtos/16. Out/G3/ANEXOS/b.pdf", "4. Produtos/16. Out/G2/c.pdf", PERM, aplicar=True)
    assert r["acao"] == "BLOQUEADO" and not [e for e in g.escritas if e[0] == "PATCH"]
    with pytest.raises(C.CaminhoInvalido):
        s.mover("4. Produtos/16. Out/G2/c.pdf", "3. Desenvolvimento/c.pdf", PERM, aplicar=True)
    with pytest.raises(C.CaminhoInvalido):
        s.mover("4. Produtos/16. Out/G2/c.pdf", "5. Compartilhado Parceiro/c.pdf", ["5. Compartilhado Parceiro"], aplicar=True)


def test_criar_pasta_idempotente(sp):
    s, g = sp
    r = s.criar_pasta("3. Desenvolvimento/G1/Entregas Recebidas", aplicar=True)
    assert r["criadas"] == ["3. Desenvolvimento", "3. Desenvolvimento/G1", "3. Desenvolvimento/G1/Entregas Recebidas"]
    assert s.criar_pasta("3. Desenvolvimento/G1/Entregas Recebidas", aplicar=True)["criadas"] == []


def test_copiar_pasta_no_servidor_e_confere(sp):
    s, g = sp
    seco = s.copiar_pasta("4. Produtos/16. Out", "3. Desenvolvimento/Recebidas/16. Out")
    assert seco["dry_run"] and seco["arquivos"] == 3 and not g.escritas
    r = s.copiar_pasta("4. Produtos/16. Out", "3. Desenvolvimento/Recebidas/16. Out", aplicar=True)
    assert r["verificacao"] == "OK" and r["arquivos_destino"] == 3 and r["pastas_iguais"]
    assert g.arq["3. Desenvolvimento/Recebidas/16. Out/G3/ANEXOS/P3.1 - X/a.pdf"] == b"aaa"
    assert s.copiar_pasta("4. Produtos/16. Out", "3. Desenvolvimento/Recebidas/16. Out")["acao"] == "BLOQUEADO"


# ---------------------------------------------------------------- arquivos novos do plano (DEC-070)
def _plano_mod():
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts", "manutencao"))
    import reorganizar_mes_pop_plano as P
    return P


def test_desmembrar_zip_pdf_proprio_e_zip_interno_aberto(tmp_path, monkeypatch):
    import io
    import zipfile
    P = _plano_mod()
    monkeypatch.setattr(P, "LIMITE_PARTE", 10)
    interno = io.BytesIO()
    with zipfile.ZipFile(interno, "w") as zi:
        zi.writestr("a.pdf", b"AAAAAAAA")
        zi.writestr("b.pdf", b"BBBBBBBB")
    src = tmp_path / "RAD.zip"
    with zipfile.ZipFile(src, "w") as z:
        z.writestr("RAD-003.1.pdf", b"%PDF-relatorio")
        z.writestr("Anexos.zip", interno.getvalue())
    item = {"pdf": "G2/ANEXOS/P2.4 - X - RAD-003.1 - RD-102A.pdf", "anexos_stem": "G2/ANEXOS/P2.4 - X - Anexos - RAD-003.1",
            "rodovia": "RD-102A"}
    saidas = P._desmembrar(str(src), item, str(tmp_path / "out"))
    rels = [s["rel"] for s in saidas]
    assert rels == ["G2/ANEXOS/P2.4 - X - RAD-003.1 - RD-102A.pdf",
                    "G2/ANEXOS/P2.4 - X - Anexos - RAD-003.1 - Parte 01 - RD-102A.zip",
                    "G2/ANEXOS/P2.4 - X - Anexos - RAD-003.1 - Parte 02 - RD-102A.zip"]
    assert (tmp_path / "out" / "G2" / "ANEXOS" / "P2.4 - X - RAD-003.1 - RD-102A.pdf").read_bytes() == b"%PDF-relatorio"
    with zipfile.ZipFile(tmp_path / "out" / "G2" / "ANEXOS" / "P2.4 - X - Anexos - RAD-003.1 - Parte 01 - RD-102A.zip") as z:
        assert z.namelist() == ["Anexos/a.pdf"] and z.read("Anexos/a.pdf") == b"AAAAAAAA"  # sem ZIP dentro de ZIP


def test_zip_individual_conteudo_identico(tmp_path):
    import zipfile
    P = _plano_mod()
    src = tmp_path / "p.xlsx"
    src.write_bytes(b"PK-planilha")
    item = {"zip": "G5/ANEXOS/P5.3 - Y - Planilha (ID 8) - RD-101.zip", "interno": "P5.3 - Y - Planilha (ID 8) - RD-101.xlsx"}
    P._zip_individual(str(src), item, str(tmp_path / "out"))
    with zipfile.ZipFile(tmp_path / "out" / "G5" / "ANEXOS" / "P5.3 - Y - Planilha (ID 8) - RD-101.zip") as z:
        assert z.read("P5.3 - Y - Planilha (ID 8) - RD-101.xlsx") == b"PK-planilha"


def test_papel_dos_arquivos_no_layout_do_pop():
    P = _plano_mod()
    assert P._papel("Índice de Arquivos - Entrega de Produtos Programa - Outubro 2026.docx") == "indice"
    assert P._papel("G3/PROG_Lista de demandas (G3).xlsx") == "lista"
    assert P._papel("G2/Programa - Relatório Grupo 2 - Relatório Mensal N°16 - 10 2026.pdf") == "relatorio"
    assert P._papel("G3/ANEXOS/P3.1 - Gestao/x.pdf") == "subpasta"
    assert P._papel("G4/ANEXOS/P4.1 - Relatório de Portfólio - X - GERAL.pdf") == "anexo"


def test_zip_aninhado_em_qualquer_nivel_vira_pasta(tmp_path):
    import io
    import zipfile
    P = _plano_mod()
    n2 = io.BytesIO()
    with zipfile.ZipFile(n2, "w") as z:
        z.writestr("ctd.dwg", b"D")
    n1 = io.BytesIO()
    with zipfile.ZipFile(n1, "w") as z:
        z.writestr("1. CTDs/CTD 19.zip", n2.getvalue())
        z.writestr("x.pdf", b"X")
    src = tmp_path / "a.zip"
    with zipfile.ZipFile(src, "w") as z:
        z.writestr("Anexos.zip", n1.getvalue())
    with zipfile.ZipFile(src) as z:
        assert sorted(m[0] for m in P._membros(z)) == ["Anexos/1. CTDs/CTD 19/ctd.dwg", "Anexos/x.pdf"]

# -*- coding: utf-8 -*-
"""Compatibilidade da automação do G1 com os DOIS layouts de anexos (DEC-043) e com a nova designação
'Relatório de Produtos Mensal' (DEC-044). Árvores artificiais em tmp — nada real é tocado."""
import json
import os

import pytest

import prog_config as cfg

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _arvore(tmp_path, nome_mes, layout):
    raiz = tmp_path / "SP"
    dp = raiz / "Programa - 3. Desenvolvimento" / "G1 - Coordenacao" / "Desenvolvimento de Produtos"
    dp.mkdir(parents=True)
    g1 = raiz / "Programa - 4. Produtos" / nome_mes / "G1"
    if layout == "por_produto":
        for s in ("P1.1 - Gestao Executiva", "P1.2 - Relacoes Institucionais", "P1.3 - Comunicacao Integrada",
                  "P1.4 - Tecnologia de Informacao de Comunicacoes"):
            (g1 / "Anexos" / s).mkdir(parents=True)
        (g1 / "Anexos" / "P1.3 - Comunicacao Integrada" / "01. Carta.pdf").write_bytes(b"%PDF")
        (g1 / "Anexos" / "P1.1 - Gestao Executiva" / "P1.1 - Quadro Financeiro do mês de agosto.xlsx").write_bytes(b"PK")
    else:
        (g1 / "ANEXOS").mkdir(parents=True)
        (g1 / "ANEXOS" / "P1.4 - TIC - Gestão Integrada - 2026-10-31 - GERAL.pdf").write_bytes(b"%PDF")
        (g1 / "ANEXOS" / "P1.3 - Comunicação Integrada - Carta 015-2026 - GERAL.pdf").write_bytes(b"%PDF")
        (g1 / "ANEXOS" / "P1.1 - Gestão Executiva - Quadro Financeiro do mês de outubro - GERAL.xlsx").write_bytes(b"PK")
        (g1 / "ANEXOS" / "Thumbs.db").write_bytes(b"x")
    (g1 / "PROG_Lista de demandas G1 Rev.CG.xlsx").write_bytes(b"PK")
    return raiz, dp


@pytest.fixture
def params(tmp_path, monkeypatch):
    def _mk(num, mes, mes_num, extra=None):
        p = json.load(open(os.path.join(RAIZ, "config", "parametros-exemplo.json"), encoding="utf-8"))
        p.update(numero_relatorio=num, mes=mes, mes_numero=mes_num, ano=2026,
                 periodo_inicio=f"01/{mes_num:02d}/2026", periodo_fim=f"30/{mes_num:02d}/2026")
        p.update(extra or {})
        f = tmp_path / f"params_{num}.json"
        f.write_text(json.dumps(p), encoding="utf-8")
        monkeypatch.setenv("PROG_PARAMS", str(f))
        monkeypatch.setenv("PROG_WORK_DIR", str(tmp_path / "work"))
    return _mk


def test_layout_legado_continua_igual(tmp_path, monkeypatch, params):
    raiz, _ = _arvore(tmp_path, "14. Agosto 2026", "por_produto")
    monkeypatch.setenv("PROG_ROOT", str(raiz))
    params(14, "Agosto", 8)
    assert cfg.layout_anexos() == "por_produto"
    assert cfg.anexo_p11_dir().endswith("P1.1 - Gestao Executiva")
    assert cfg.nome_anexo_g1("P1.1", "Quadro Financeiro do mês de agosto", ".xlsx") == "P1.1 - Quadro Financeiro do mês de agosto.xlsx"
    reais = cfg.listar_anexos_reais()
    assert [n for _s, n in reais] == ["P1.1 - Quadro Financeiro do mês de agosto.xlsx", "01. Carta.pdf"]


def test_layout_plano_pop(tmp_path, monkeypatch, params):
    raiz, _ = _arvore(tmp_path, "16. Outubro 2026", "plano")
    monkeypatch.setenv("PROG_ROOT", str(raiz))
    params(16, "Outubro", 10)
    assert cfg.layout_anexos() == "plano"
    assert os.path.basename(cfg.anexo_p11_dir()) == "ANEXOS" == os.path.basename(cfg.anexo_p12_dir())
    assert cfg.nome_anexo_g1("P1.2", "Controle de Reuniões do mês de outubro", ".xlsx") == \
        "P1.2 - Relações Institucionais - Controle de Reuniões do mês de outubro - GERAL.xlsx"
    reais = cfg.listar_anexos_reais()
    assert [s for s, _n in reais] == ["P1.1", "P1.3", "P1.4"]  # ordem por produto; Thumbs.db ignorado
    assert cfg.find_in_dir(cfg.anexo_p11_dir(), "p1.1", "quadro financeiro", ext=".xlsx")


def test_override_de_layout(tmp_path, monkeypatch, params):
    raiz, _ = _arvore(tmp_path, "16. Outubro 2026", "plano")
    monkeypatch.setenv("PROG_ROOT", str(raiz))
    params(16, "Outubro", 10, {"layout_anexos": "por_produto"})
    assert cfg.layout_anexos() == "por_produto"


def test_designacao_relatorio_de_produtos_mensal(tmp_path, monkeypatch, params):
    raiz, _ = _arvore(tmp_path, "16. Outubro 2026", "plano")
    monkeypatch.setenv("PROG_ROOT", str(raiz))
    params(16, "Outubro", 10)
    assert cfg.doc_final_basename().startswith("Programa - Relatório Grupo 1 - Relatório de Produtos Mensal N°16 - ")
    params(16, "Outubro", 10, {"termo_documento": "Relatório Mensal"})  # re-geração com nome histórico
    assert "Relatório Mensal N°16" in cfg.doc_final_basename()


def test_lista_de_demandas_no_arquivo_interno_apos_publicacao(tmp_path, monkeypatch, params):
    import pathlib, shutil, tempfile
    curto = pathlib.Path(tempfile.mkdtemp(prefix="g1"))  # caminho curto: a árvore real passa de 260 no tmp do pytest
    raiz, dp = _arvore(curto, "16. Outubro 2026", "plano")
    monkeypatch.setenv("PROG_ROOT", str(raiz))
    params(16, "Outubro", 10)
    lista = raiz / "Programa - 4. Produtos" / "16. Outubro 2026" / "G1" / "PROG_Lista de demandas G1 Rev.CG.xlsx"
    arq = dp / "Controle Interno Entregas" / "16. Outubro 2026" / "2026-11-09_100000" / "G1"
    arq.mkdir(parents=True)
    (arq / lista.name).write_bytes(lista.read_bytes())
    lista.unlink()  # como após a publicação do pacote
    try:
        assert cfg.lista_demandas_path("G1") == str(arq / lista.name)
    finally:
        shutil.rmtree(curto, ignore_errors=True)

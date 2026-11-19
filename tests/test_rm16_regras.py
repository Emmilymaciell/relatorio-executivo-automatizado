# -*- coding: utf-8 -*-
"""Regras permanentes introduzidas no RM16 (16/11/2026) — DEC-058…066. Tudo com fixtures ARTIFICIAIS em tmp.

- DEC-058 data da capa = último dia da competência (genérico) + leitura/correção da capa embutida (PPTX no DOCX)
- DEC-059 preservação de insumos externos (cópia imutável, SHA-256, proveniência, idempotência, sem sobrescrever)
- DEC-060 prestações de contas iterativas na Timeline (eventos do MÊS; envio/análise/retorno/aprovação; sem evento)
- DEC-061 corte temporal do Controle de Produtos
- DEC-062 conferência PDF × planilha reconstruída (sem inventar valores)
- DEC-063 medição nomeada pela competência de execução
- DEC-064 nome 'EM ELABORAÇÃO' enquanto houver pendências
- DEC-052 (ext.) espelho somente-leitura do SharePoint
"""
import datetime
import importlib.util
import io
import json
import os
import stat
import zipfile

import pytest

import prog_config as cfg

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SCRIPTS = os.path.join(RAIZ, ".claude", "skills", "relatorio-prog-grupo1", "scripts")


def _mod(nome):
    spec = importlib.util.spec_from_file_location(nome.replace(".py", ""), os.path.join(SCRIPTS, nome))
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m


@pytest.fixture
def params(tmp_path, monkeypatch):
    def _mk(num, mes, mes_num, ano=2026, extra=None, dados=None):
        p = json.load(open(os.path.join(RAIZ, "config", "parametros-exemplo.json"), encoding="utf-8"))
        p.update(numero_relatorio=num, mes=mes, mes_numero=mes_num, ano=ano,
                 periodo_inicio=f"01/{mes_num:02d}/{ano}", periodo_fim=cfg.fim_competencia(ano, mes_num))
        p.update(extra or {})
        f = tmp_path / f"params_{num}.json"
        f.write_text(json.dumps(p), encoding="utf-8")
        monkeypatch.setenv("PROG_PARAMS", str(f))
        monkeypatch.setenv("PROG_WORK_DIR", str(tmp_path / "work"))
        dm = tmp_path / f"dados_{num}.py"
        dm.write_text(dados or "MEDICAO_NUM = 1\n", encoding="utf-8")
        monkeypatch.setenv("PROG_DADOS", str(dm))
    return _mk


# ------------------------------------------------------------------------------------------------ DEC-058
@pytest.mark.parametrize("ano,mes,esperado", [(2026, 10, "31/10/2026"), (2026, 11, "30/11/2026"),
                                              (2026, 12, "31/12/2026"), (2028, 2, "29/02/2028"),
                                              (2027, 2, "28/02/2027")])
def test_data_capa_fim_da_competencia(params, ano, mes, esperado):
    params(16, cfg.MESES_PT[mes - 1], mes, ano=ano)
    assert cfg.data_capa() == esperado
    assert cfg.data_recorrente() == esperado


def test_data_capa_ignora_data_de_elaboracao_no_mes_seguinte(params, capsys):
    params(16, "Outubro", 10, extra={"data_documento": "05/11/2026"})
    assert cfg.data_documento() == "31/10/2026"
    assert "ignorada" in capsys.readouterr().out
    assert cfg.data_extenso(cfg.data_capa()) == "31 de outubro de 2026"


def _docx_com_capa(path, texto_data):
    from pptx import Presentation
    from pptx.util import Pt
    pr = Presentation()
    s = pr.slides.add_slide(pr.slide_layouts[6])
    tb = s.shapes.add_textbox(0, 0, Pt(300), Pt(40)); tb.text_frame.text = "Relatório Mensal N° 16"
    tb = s.shapes.add_textbox(0, Pt(50), Pt(300), Pt(40)); tb.text_frame.text = texto_data
    b = io.BytesIO(); pr.save(b)
    import docx
    d = docx.Document(); d.add_paragraph("corpo"); d.save(path)
    with zipfile.ZipFile(path, "a") as z:
        z.writestr("word/embeddings/Microsoft_PowerPoint_Presentation.pptx", b.getvalue())


def test_capa_le_e_corrige_data(tmp_path):
    capa = _mod("capa.py")
    p = str(tmp_path / "base.docx")
    _docx_com_capa(p, "05 de novembro de 2026")
    assert capa.ler_data_capa(p) == "05 de novembro de 2026"
    antes, res = capa.garantir_data_capa(p, "31 de outubro de 2026") if False else (None, None)
    res = capa.corrigir_data_capa(p, "31 de outubro de 2026", regenerar_preview=False)
    assert res["trocas"] == 1 and capa.ler_data_capa(p) == "31 de outubro de 2026"
    # idempotente: já correta => não mexe
    antes, res2 = capa.garantir_data_capa(p, "31 de outubro de 2026")
    assert antes == "31 de outubro de 2026" and res2 is None


def test_competencia_da_medicao_e_mes_anterior(params):
    params(16, "Outubro", 10)
    assert cfg.competencia_medicao() == ("Setembro", 2026)
    params(19, "Janeiro", 1, ano=2027)
    assert cfg.competencia_medicao() == ("Dezembro", 2026)


# ------------------------------------------------------------------------------------------------ DEC-064
def test_nome_em_elaboracao_enquanto_houver_pendencias(params):
    params(16, "Outubro", 10, dados="PENDENCIAS = {'reunioes': 'aguardando'}\n")
    assert cfg.doc_final_basename().endswith("01.10 a 31.10 - EM ELABORAÇÃO")
    params(16, "Outubro", 10, dados="PENDENCIAS = {}\n")
    assert cfg.doc_final_basename().endswith("01.10 a 31.10")


# ------------------------------------------------------------------------------------------------ DEC-059
def test_importar_fonte_externa_preserva_imutavel_e_idempotente(params, tmp_path):
    params(16, "Outubro", 10)
    fe = _mod("fontes_externas.py")
    origem = tmp_path / "Downloads" / "novembro"; origem.mkdir(parents=True)
    a = origem / "CONTROLE DE PRAZOS - PRESTAÇÃO DE CONTAS.xlsx"; a.write_bytes(b"conteudo-1")
    it = fe.importar(str(a), "Timeline", classe="PRESTACAO_CONTAS")
    copia = it["arquivo"]
    assert os.path.exists(copia) and open(copia, "rb").read() == b"conteudo-1"
    assert not (os.stat(copia).st_mode & stat.S_IWRITE)  # somente-leitura
    for k in ("nome_original", "origem", "ingestao", "competencia", "finalidade", "sha256"):
        assert it[k]
    assert it["competencia"] == "Outubro/2026"
    # reimportar o mesmo conteúdo: não duplica
    it2 = fe.importar(str(a), "Timeline", classe="PRESTACAO_CONTAS")
    assert it2["arquivo"] == copia and len(fe.registro()["itens"]) == 1
    # conteúdo novo com o mesmo nome: não sobrescreve o original importado
    a.write_bytes(b"conteudo-2")
    it3 = fe.importar(str(a), "Timeline", classe="PRESTACAO_CONTAS")
    assert it3["arquivo"] != copia and " (2)" in it3["arquivo"]
    assert open(copia, "rb").read() == b"conteudo-1"
    # a origem pode sumir: a cópia continua
    os.remove(str(a))
    assert os.path.exists(copia) and fe.caminho_importado("PRESTACAO_CONTAS") == it3["arquivo"]


def test_classificacao_ambigua_nao_inventa(tmp_path):
    fe = _mod("fontes_externas.py")
    f = tmp_path / "arquivo qualquer.bin"; f.write_bytes(b"x")
    assert fe.classificar(str(f))[0] == "AMBIGUO"


# ------------------------------------------------------------------------------------------------ DEC-060
def _planilha_pc(path):
    import openpyxl
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["CONTROLE DE PRAZOS - PRESTAÇÃO DE CONTAS"])
    ws.append(["REFERÊNCIA", "PERÍODO", "DATA DO ENVIO Programa", "1ª ANALISE Órgão Contratante", "1º RETORNO Programa",
               "2ª ANALISE Órgão Contratante", "2º RETORNO Programa", "3ª ANALISE Órgão Contratante", "3º RETORNO Programa", "APROVAÇÃO"])
    D = datetime.datetime
    # prestação antiga com movimentação em outubro (iterativa)
    ws.append(["RD-101", "3º TRIMESTRE", D(2026, 8, 31), D(2026, 9, 14), D(2026, 10, 2), D(2026, 10, 7),
               D(2026, 10, 19), D(2026, 10, 23), "28/10/2026 - Executora\n29/10/2026 - Programa", None])
    # análise em outubro seguida de retorno só em novembro
    ws.append(["RD-102", "3º TRIMESTRE", "16/09/2026\n18/09/2026 - Complemento", D(2026, 10, 28), D(2026, 11, 6)])
    # Estruturadora: envio em outubro
    ws.append(["Programa", "Setembro", D(2026, 10, 14)])
    # aprovada em setembro (não aparece em outubro)
    ws.append(["RD-104", "2º TRIMESTRE", D(2026, 7, 21), D(2026, 7, 28), D(2026, 7, 30), None, None, None, None, "SIM"])
    wb.save(path)


def test_prestacao_contas_eventos_iterativos_do_mes(tmp_path):
    pc = _mod("prestacao_contas.py")
    p = str(tmp_path / "pc.xlsx"); _planilha_pc(p)
    ev = pc.marcos_da_competencia(p, 2026, 10)
    datas = [(e["data"].strftime("%d/%m"), e["referencia"], e["tipo"], e.get("ator")) for e in ev]
    assert ("02/10", "RD-101", "RETORNO", None) in datas            # prestação antiga gera evento no mês
    assert ("07/10", "RD-101", "ANALISE", None) in datas
    assert ("28/10", "RD-101", "RETORNO", "Executora") in datas      # célula multi-linha: ator preservado
    assert ("29/10", "RD-101", "RETORNO", "Programa") in datas
    assert ("28/10", "RD-102", "ANALISE", None) in datas
    assert ("14/10", "Programa", "ENVIO", None) in datas
    assert not any(r == "RD-104" for _d, r, _t, _a in datas)         # sem movimento em outubro
    # nada de setembro/novembro
    assert all(e["data"].month == 10 for e in ev)
    txt = {(e["data"].strftime("%d/%m"), e["referencia"], e["tipo"]): e["texto"] for e in ev}
    # análise seguida de retorno (mesmo em outro mês) => apontamentos; rótulo da fonte preservado (RD-102)
    assert "com apontamentos" in txt[("28/10", "RD-102", "ANALISE")] and "RD-102" in txt[("28/10", "RD-102", "ANALISE")]
    assert txt[("14/10", "Programa", "ENVIO")] == "Envio ao Parceiro da Prestação de Contas da Estruturadora (Programa) referente a setembro."
    assert pc.classificar_referencia("RD-102A") == "OBRA" and pc.classificar_referencia("Programa") == "ESTRUTURADORA"
    assert pc.classificar_referencia("Contrato X") == "INDEFINIDA"


def test_prestacao_contas_mes_sem_evento_nao_inventa(tmp_path):
    pc = _mod("prestacao_contas.py")
    p = str(tmp_path / "pc.xlsx"); _planilha_pc(p)
    assert pc.marcos_da_competencia(p, 2026, 3) == []
    te = _mod("timeline_prestacao_excel.py")
    r = te.gravar(p, str(tmp_path / "nao_criado.xlsx"), p, 2026, 3)
    assert r["linhas"] == 0 and not os.path.exists(tmp_path / "nao_criado.xlsx")


def test_agrupar_por_data_um_marcador_por_data(tmp_path):
    pc = _mod("prestacao_contas.py")
    p = str(tmp_path / "pc.xlsx"); _planilha_pc(p)
    g = pc.agrupar_por_data(pc.marcos_da_competencia(p, 2026, 10))
    d28 = next(x for x in g if x["data"].day == 28)
    assert len(d28["eventos"]) == 2 and d28["empreendimentos"] == ["RD-101", "RD-102"]


# ------------------------------------------------------------------------------------------------ DEC-067
def test_controle_sem_corte_temporal_tudo_que_foi_declarado_entra():
    """DEC-061 revogada: data de upload/entrega não muda a competência; linha repetida na fonte não é descartada."""
    c01 = _mod("01_consolidar_entregas.py")
    assert not hasattr(c01, "aplicar_corte_temporal") and not hasattr(c01, "corte_temporal_ativo")
    fonte = open(os.path.join(SCRIPTS, "01_consolidar_entregas.py"), encoding="utf-8").read()
    assert "duplicadas_na_fonte" in fonte and "declarados_por_fonte" in fonte


def test_evidencias_corrigem_rotulo_e_acrescentam_evento(tmp_path):
    pc = _mod("prestacao_contas.py")
    p = str(tmp_path / "pc.xlsx"); _planilha_pc(p)
    ev = {"rotulos": {"RD-102": {"para": "RD-102A", "evidencia": "Manifestação"}},
          "periodos": {"RD-102|3º TRIMESTRE": "março a maio de 2026", "Programa|Setembro": "setembro de 2026"},
          "referencias": {"RD-101|3º TRIMESTRE|RETORNO|3|Programa": "Manifestação nº 043/2026"},
          "eventos_extra": [{"data": "2026-10-30", "referencia": "RD-104", "periodo": "3º TRIMESTRE", "tipo": "RETORNO",
                             "rodada": 1, "ator": "Executora", "texto": "Encaminhamento pela executora.", "fonte": "SP"}]}
    m = pc.marcos_da_competencia(p, 2026, 10, ev)
    t = {(x["data"].day, x["referencia"], x.get("ator")): x["texto"] for x in m}
    assert "Prestação de Contas de Obra da RD-102A (3º Trimestre – março a maio de 2026)" in t[(28, "RD-102", None)]
    assert "Manifestação nº 043/2026" in t[(29, "RD-101", "Programa")]
    assert "Manifestação" not in t[(28, "RD-101", "Executora")]     # documento do Programa não vai para a executora
    assert t[(14, "Programa", None)] == "Envio ao Parceiro da Prestação de Contas da Estruturadora (Programa) referente a setembro de 2026."
    assert (30, "RD-104", "Executora") in t
    g = pc.agrupar_por_data(m)
    assert next(x for x in g if x["data"].day == 28)["empreendimentos"] == ["RD-101", "RD-102A"]


def test_assinante_p11_p12_configuravel():
    fonte = open(os.path.join(SCRIPTS, "10_encerramento_assinaturas_anexos.py"), encoding="utf-8").read()
    assert 'P.get("assinante_p11_p12")' in fonte


# ------------------------------------------------------------------------------------------------ DEC-062
def _pdf_quadro(path):
    """Quadro mínimo no layout do Parceiro (texto vetorial): cabeçalho + 2 produtos + total."""
    import fitz
    d = fitz.open(); pg = d.new_page(width=1190, height=842)
    pg.insert_text((900, 60), "15ª MEDIÇÃO", fontsize=20)
    pg.insert_text((50, 80), "VALOR ACUMULADO", fontsize=6)
    xs = [430, 480, 540, 600, 660, 700, 740, 780, 820, 880, 940, 1000]
    def linha(y, rot, vals):
        pg.insert_text((40, y), rot, fontsize=7)
        pg.insert_text((330, y), "MENSAL", fontsize=7)
        for x, v in zip(xs, vals):
            if v:
                w = fitz.get_text_length(v, fontsize=7); pg.insert_text((x - w, y), v, fontsize=7)
    linha(120, "1.1", ["28,00", "50.000,00", "1.400.000,00", "400.000,00", "1.800.000,00", "3,00%", "50,00%",
                      "1,00", "100,00%", "50.000,00", "50.000,00", "900.000,00"])
    linha(140, "1.3", ["1,00", "25.000,00", "25.000,00", "150.000,00", "175.000,00", "0,00%", "100,00%",
                      "-", "", "-", "-", "175.000,00"])
    d.save(path)


def test_extrai_quadro_pdf_e_confere_sem_inventar(tmp_path):
    mp = _mod("medicao_pdf.py")
    mr = _mod("medicao_reconstrucao.py")
    p = str(tmp_path / "q.pdf"); _pdf_quadro(p)
    ext = mp.extrair(p)
    assert ext["medicao"] == 15
    assert ext["produtos"]["1.1"]["valor_acumulado"] == 900000.00
    assert ext["produtos"]["1.3"]["quant_medida"] is None and ext["produtos"]["1.3"]["nota"] is None  # '-'/vazio
    planilha = {"produtos": {"1.1": dict(ext["produtos"]["1.1"]), "1.3": dict(ext["produtos"]["1.3"])},
                "grupos": {}, "total": {}, "pct_execucao": {}}
    ext2 = dict(ext, grupos={}, total={}, pct_execucao={})
    dif, ok = mr.conferir(planilha, ext2)
    assert not dif and ok > 10
    planilha["produtos"]["1.1"]["valor_acumulado"] += 0.02   # 2 centavos de diferença => divergência
    dif, _ = mr.conferir(planilha, ext2)
    assert any(d["campo"] == "1.1.valor_acumulado" for d in dif)


# ------------------------------------------------------------------------------------------------ espelho
def test_espelhar_baixa_arvore_e_reaproveita(tmp_path):
    from test_sharepoint import GraphFalsa, CFG
    from prog_sharepoint.operacoes import SharePoint
    g = GraphFalsa({"4. Produtos/16. Outubro 2026/G1/a.pdf": b"A", "4. Produtos/16. Outubro 2026/G2/ANEXOS/b.pdf": b"BB"})
    sp = SharePoint(CFG, g, str(tmp_path / "SP"), str(tmp_path))
    dest = str(tmp_path / "ESP")
    r = sp.espelhar("4. Produtos/16. Outubro 2026", dest)
    assert r["baixados"] == 2 and r["arquivos"] == 2
    assert open(os.path.join(dest, "Programa - 4. Produtos", "16. Outubro 2026", "G2", "ANEXOS", "b.pdf"), "rb").read() == b"BB"
    man = json.load(open(os.path.join(dest, "_espelho.json"), encoding="utf-8"))
    assert man["4. Produtos/16. Outubro 2026/G1/a.pdf"]["sha256"]
    assert not g.escritas  # somente leitura

# -*- coding: utf-8 -*-
"""Testes da pipeline da Coordenação (POP de Entrega de Produtos). Fixtures artificiais — ver conftest.py."""
import json
import os
import shutil
import zipfile

import pytest

from conftest import (adulterar, make_docx, make_lista, make_pdf, make_pdf_assinado, make_pdf_protegido,
                      make_pdf_restrito, make_pdf_sem_texto, make_pdf_visual, make_xlsx, make_zip,
                      montar_mes_correto, nome_anexo)
from prog_entrega import assinatura as S
from prog_entrega import estrutura as E
from prog_entrega import hashcache as H
from prog_entrega import indice as IDX
from prog_entrega import inventario as INV
from prog_entrega import nomes as N
from prog_entrega import pipeline as PL
from prog_entrega import regras as R
from prog_entrega import validacao as V


def auditar(pasta):
    cache = H.Cache()
    inv = INV.inventariar(pasta, cache)
    ach, plano, finais = V.auditar(inv, cache, "ORIGEM")
    return ach.itens, plano, inv


def achados_de(pasta_out):
    return json.load(open(os.path.join(pasta_out, "04_RELATORIOS", "achados.json"), encoding="utf-8"))


def tem(achados, status, trecho, fase=None):
    return any(a["status"] == status and trecho.lower() in a["mensagem"].lower()
               and (fase is None or a["fase"] == fase) for a in achados)


def out_dir(pasta):
    return R.sem_lp(PL.dir_saida(pasta)[0])


# =========================================================================== assinatura
class TestAssinatura:
    def test_assinatura_valida(self, tmp_path):
        p = make_pdf_assinado(str(tmp_path / "a.pdf"))
        r = S.verificar_pdf(p)
        assert r["status"] == S.VALIDADA, r
        assert r["assinaturas"][0]["signatario"] == "SIGNATARIO DE TESTE"
        assert r["assinaturas"][0]["icp_brasil"] is False

    def test_assinatura_adulterada_e_invalida(self, tmp_path):
        p = make_pdf_assinado(str(tmp_path / "a.pdf"))
        adulterar(p)
        assert S.verificar_pdf(p)["status"] == S.INVALIDA

    def test_sem_assinatura(self, tmp_path):
        assert S.verificar_pdf(make_pdf(str(tmp_path / "n.pdf")))["status"] == S.NAO_ENCONTRADA

    def test_assinatura_so_visual_exige_revisao_humana(self, tmp_path):
        assert S.verificar_pdf(make_pdf_visual(str(tmp_path / "v.pdf")))["status"] == S.VISUAL

    def test_alteracao_apos_assinatura_nao_valida(self, tmp_path):
        p = make_pdf_assinado(str(tmp_path / "a.pdf"))
        with open(p, "ab") as f:  # atualização incremental com conteúdo de página (ex.: rubrica)
            f.write(b"8 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj\n%%EOF\n")
        r = S.verificar_pdf(p)
        assert r["status"] == S.NAO_VERIFICAVEL and "ALTERADO" in r["motivo"]


# ============================================================================== arquivos
class TestArquivos:
    def test_pdf_invalido(self, tmp_path):
        p = tmp_path / "x.pdf"
        p.write_bytes(b"isto nao e um pdf")
        from prog_entrega import arquivos as A
        v = A.validar(str(p))
        assert any(s == R.BLOCKER for s, _ in v["problemas"])

    def test_pdf_corrompido(self, tmp_path):
        from prog_entrega import arquivos as A
        p = make_pdf(str(tmp_path / "c.pdf"), paginas=3)
        raw = open(p, "rb").read()
        open(p, "wb").write(raw[:200])
        assert any(s == R.BLOCKER for s, _ in A.validar(p)["problemas"])

    def test_pdf_com_senha(self, tmp_path):
        from prog_entrega import arquivos as A
        v = A.validar(make_pdf_protegido(str(tmp_path / "s.pdf")))
        assert any(s == R.BLOCKER and "SENHA" in m for s, m in v["problemas"])

    def test_pdf_com_restricao(self, tmp_path):
        from prog_entrega import arquivos as A
        v = A.validar(make_pdf_restrito(str(tmp_path / "r.pdf")))
        assert any(s == R.BLOCKER and "RESTRIÇÃO" in m for s, m in v["problemas"])

    def test_pdf_nao_pesquisavel(self, tmp_path):
        from prog_entrega import arquivos as A
        v = A.validar(make_pdf_sem_texto(str(tmp_path / "img.pdf")))
        assert any(s == R.MANUAL and "pesquisável" in m for s, m in v["problemas"])

    def test_zip_invalido(self, tmp_path):
        from prog_entrega import arquivos as A
        p = tmp_path / "z.zip"
        p.write_bytes(b"PK\x03\x04lixo")
        assert any(s == R.BLOCKER for s, _ in A.validar(str(p))["problemas"])

    def test_zip_crc_corrompido(self, tmp_path):
        from prog_entrega import arquivos as A
        p = make_zip(str(tmp_path / "z.zip"), {"a.txt": "x" * 5000})
        raw = bytearray(open(p, "rb").read())
        raw[60] ^= 0xFF
        open(p, "wb").write(bytes(raw))
        assert any(s == R.BLOCKER for s, _ in A.validar(p)["problemas"])

    def test_extensao_zip_com_conteudo_pdf(self, tmp_path):
        from prog_entrega import arquivos as A
        p = make_pdf(str(tmp_path / "x.pdf"))
        z = tmp_path / "x.zip"
        shutil.copy(p, z)
        assert any(s == R.BLOCKER and "não é ZIP" in m for s, m in A.validar(str(z))["problemas"])

    def test_zip_com_pdf_stored_no_inicio_continua_zip(self, tmp_path):
        """Regressão RM15: ZIP cujo 1º membro é PDF gravado sem compressão traz '%PDF-' no 1º KB — continua ZIP."""
        import zipfile
        from prog_entrega import arquivos as A
        p = make_pdf(str(tmp_path / "x.pdf"))
        z = tmp_path / "x.zip"
        with zipfile.ZipFile(z, "w", compression=zipfile.ZIP_STORED) as zf:
            zf.write(p, "x.pdf")
        v = A.validar(str(z))
        assert v["tipo"] == "zip" and not any(s == R.BLOCKER for s, _ in v["problemas"])

    def test_planilha_com_formulas_e_lida_sem_alterar(self, tmp_path):
        from prog_entrega import arquivos as A
        p = make_xlsx(str(tmp_path / "f.xlsx"))
        antes = H.sha256_arquivo(p)
        v = A.validar(p)
        assert v["info"]["formulas"] >= 1
        assert H.sha256_arquivo(p) == antes  # validação NÃO regrava a planilha


# ================================================================================= nomes
class TestNomes:
    def test_nome_correto(self):
        a = N.analisar("P2.2 - Análise de Projeto e RPFO - Nota Técnica de Aceitação de Documentos - NT-749.1 - RD-101.pdf")
        assert a["conforme"] and a["rodovia"] == "RD-101" and a["codigo"] == "P2.2"

    def test_geral_valido(self):
        assert N.analisar("P1.4 - TIC - Gestão Integrada - 2026-08-31 - GERAL.pdf")["conforme"]

    def test_nome_incorreto(self):
        a = N.analisar("documento final.pdf")
        assert not a["conforme"]

    def test_sem_rodovia(self):
        a = N.analisar("P2.1 - Coordenação Serviços de Engenharia - Riscos associados - Programa.pdf")
        assert not a["conforme"] and any("rodovia" in m for _, m in a["problemas"])

    def test_termo_generico_pede_revisao(self):
        a = N.analisar("P2.1 - Coordenação Serviços de Engenharia - arquivo novo - RD-104.pdf")
        assert any(s == R.MANUAL for s, _ in a["problemas"])

    def test_proposta_completa_pela_pasta_e_rodovia(self):
        novo, _m, completo = N.propor("P33_-_202608_-_RD_101_-_Relatório_de_Obra.pdf", "P3.3")
        assert completo and novo.startswith("P3.3 - Relatório de Inspeção - ") and novo.endswith("- RD-101.pdf")

    def test_proposta_incompleta_sem_rodovia(self):
        novo, _m, completo = N.propor("Justificativa Plano de Inspeção.pdf", "P3.2")
        assert not completo and "[RODOVIA?]" in novo


# ========================================================================= estrutura SP
class TestEstrutura:
    def test_remove_so_subpasta_vazia_e_cria_indice(self, ambiente):
        m = ambiente["produtos"] / "16. Outubro 2026"
        (m / "G1" / "Anexos" / "P1.1 - Gestao Executiva").mkdir(parents=True)
        (m / "G2" / "Anexos" / "P2.2 - Analise").mkdir(parents=True)
        (m / "G2" / "Anexos" / "P2.2 - Analise" / "real.pdf").write_bytes(b"%PDF-1.4")
        plano = E.planejar_estrutura(str(m))
        res = E.aplicar_plano(plano)
        assert all(a["resultado"] == "OK" for a in res["acoes"])
        assert not (m / "G1" / "ANEXOS" / "P1.1 - Gestao Executiva").exists()
        assert (m / "G2" / "ANEXOS" / "P2.2 - Analise" / "real.pdf").exists()  # com arquivo: intocada
        assert any("não removida" in p for p in res["pendencias"])
        assert (m / IDX.nome_arquivo_indice("Outubro", 2026)).exists()
        # idempotente e nunca sobrescreve o índice
        plano2 = E.planejar_estrutura(str(m))
        assert not [a for a in plano2["acoes"] if a["acao"] in ("criar_indice", "remover_pasta_vazia", "substituir_indice")]

    def test_substitui_indice_antigo_so_se_em_branco(self, ambiente, monkeypatch, tmp_path):
        monkeypatch.setenv("PROG_ENTREGA_LOGS", str(tmp_path / "logs"))
        m = ambiente["produtos"] / "17. Novembro 2026"
        (m / "G1" / "ANEXOS").mkdir(parents=True)
        antigo = m / "Índice de Arquivos - Prestação de Contas Programa - Novembro 2026.docx"
        IDX.criar_modelo_mes(str(antigo), "Novembro", 2026)
        m2 = ambiente["produtos"] / "18. Dezembro 2026"
        (m2 / "G1" / "ANEXOS").mkdir(parents=True)
        preenchido = m2 / "Índice de Arquivos - Prestação de Contas Programa - Dezembro 2026.docx"
        IDX.escrever_indice(str(preenchido), "Dezembro", 2026, {"G1": [("P1.1", "algo.pdf")]})
        r1 = E.aplicar_plano(E.planejar_estrutura(str(m)))
        assert not antigo.exists() and (m / IDX.nome_arquivo_indice("Novembro", 2026)).exists()
        assert all(a["resultado"] == "OK" for a in r1["acoes"])
        assert os.path.exists(next(a["backup"] for a in r1["acoes"] if a["acao"] == "substituir_indice"))
        r2 = E.aplicar_plano(E.planejar_estrutura(str(m2)))
        assert preenchido.exists() and any("NÃO substituído" in p for p in r2["pendencias"])

    def test_regenera_indice_em_branco_de_layout_anterior_no_mesmo_nome(self, ambiente, monkeypatch, tmp_path):
        """Índice vigente em branco mas com subdivisão por produto (layout anterior) -> regenerado só por grupo."""
        import copy as _c
        import docx
        monkeypatch.setenv("PROG_ENTREGA_LOGS", str(tmp_path / "logs"))
        m = ambiente["produtos"] / "16. Outubro 2026"
        (m / "G1" / "ANEXOS").mkdir(parents=True)
        p = m / IDX.nome_arquivo_indice("Outubro", 2026)
        IDX.criar_modelo_mes(str(p), "Outubro", 2026, periodo=("01/10/2026", "31/10/2026"))
        d = docx.Document(str(p))
        tbl = d.tables[0]._tbl
        tr_grupo = [tr for tr in tbl if tr.tag.endswith("}tr")][1]
        extra = _c.deepcopy(tr_grupo)
        tr_grupo.addnext(extra)  # simula linha de subdivisão (layout anterior)
        d.save(str(p))
        plano = E.planejar_estrutura(str(m))
        assert [a["acao"] for a in plano["acoes"] if "indice" in a["acao"]] == ["substituir_indice"]
        E.aplicar_plano(plano)
        assert p.exists() and IDX.eh_indice_em_branco(str(p))
        assert not [a for a in E.planejar_estrutura(str(m))["acoes"] if "indice" in a["acao"]]  # idempotente


# ================================================================================= índice
class TestIndice:
    def test_modelo_e_leitura(self, tmp_path):
        p = str(tmp_path / "i.docx")
        IDX.criar_modelo_mes(p, "Setembro", 2026, periodo=("01/09/2026", "30/09/2026"))
        r = IDX.ler_indice(p)
        assert "Setembro/2026" in r["competencia"] and r["linhas"] == [] and IDX.eh_indice_em_branco(p)
        import docx
        dd = docx.Document(p)
        assert dd.paragraphs[0].text == "ÍNDICE DE ARQUIVOS : ENTREGA DE PRODUTOS DO PROGRAMA"
        assert not any("PRESTAÇÃO DE CONTAS" in x.text for x in dd.paragraphs)
        linhas = [[c.text for c in r_.cells] for r_ in dd.tables[0].rows]
        mescladas = [l[0] for l in linhas if len(set(l)) == 1 and l[0]]
        assert mescladas == ["GRUPO 1", "GRUPO 2", "GRUPO 3", "GRUPO 4", "GRUPO 5"]  # só grupos, sem produtos
        nums = [l[0] for l in linhas[1:] if len(set(l)) > 1]
        assert nums == ["1", "2", "3", "4"] * 5  # 1..4 reiniciando em cada grupo, como o Modelo r2

    def test_escrita_preserva_sei_e_numera(self, tmp_path):
        p = str(tmp_path / "i.docx")
        IDX.escrever_indice(p, "Setembro", 2026, {"G1": [("REL", "rel.pdf"), ("P1.1", "a.pdf")], "G2": [("P2.1", "b.pdf")]},
                            sei_por_arquivo={"a.pdf": "12345678"})
        r = IDX.ler_indice(p)
        assert [l["num"] for l in r["linhas"]] == ["1", "2", "1"]  # reinicia em cada grupo (DEC-049)
        assert [l["grupo"] for l in r["linhas"]] == ["G1", "G1", "G2"]
        assert r["linhas"][1]["sei"] == "12345678" and r["linhas"][0]["sei"] is None


# ======================================================================= pipeline (origem)
class TestAuditoria:
    def test_relatorio_ausente(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        for f in os.listdir(os.path.join(m, "G3")):
            if f.endswith(".pdf"):
                os.remove(os.path.join(m, "G3", f))
        ach, _p, _i = auditar(m)
        assert tem(ach, R.BLOCKER, "Relatório de Produtos Mensal FINAL")

    def test_relatorio_nao_assinado(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        rel = [f for f in os.listdir(os.path.join(m, "G4")) if f.endswith(".pdf")][0]
        make_pdf(os.path.join(m, "G4", rel), "relatório sem assinatura")
        ach, _p, _i = auditar(m)
        assert tem(ach, R.BLOCKER, "NÃO ENCONTRADA")

    def test_dois_relatorios_sem_criterio(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_pdf(os.path.join(m, "G2", "Programa - Relatório Grupo 2 - Relatório Mensal N°15 - final 2.pdf"), "v2")
        make_pdf_assinado(os.path.join(m, "G2", "Programa - Relatório Grupo 2 - Relatório Mensal N°15 - rev.pdf"), "outra")
        ach, _p, _i = auditar(m)
        assert tem(ach, R.BLOCKER, "versões diferentes do relatório")

    def test_relatorio_assinado_prevalece_sobre_nao_assinado(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_pdf(os.path.join(m, "G5", "Programa - Relatório Grupo 5 - Relatório Mensal N°15 - 09 2026 sem assinatura.pdf"))
        ach, plano, _i = auditar(m)
        assert tem(ach, R.FIXED, "versão não assinada/anterior do relatório retirada")
        assert any("sem assinatura" in x["origem"] for x in plano["excluidos"])

    def test_word_do_relatorio_fora_de_anexos_sai_do_pacote(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_docx(os.path.join(m, "G2", "Programa - Relatório Grupo 2 - Relatório Mensal N°15 - 09 2026.docx"))
        ach, plano, _i = auditar(m)
        assert tem(ach, R.FIXED, "versão editável/minuta do relatório RETIRADA")

    def test_arquivo_solto_na_raiz_do_grupo_bloqueia(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_pdf(os.path.join(m, "G1", "ART_assinada.pdf"))
        ach, _p, _i = auditar(m)
        assert tem(ach, R.BLOCKER, "solto na raiz do grupo")

    def test_colisao_ao_achatar(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        n = nome_anexo("P2.4", "Ata de Reunião 01", "RD-103")
        make_pdf(os.path.join(m, "G2", "ANEXOS", "Atas A", n), "conteúdo A")
        make_pdf(os.path.join(m, "G2", "ANEXOS", "Atas B", n), "conteúdo B")
        ach, _p, _i = auditar(m)
        assert tem(ach, R.BLOCKER, "COLISÃO")

    def test_duplicata_exata_com_marcador_de_copia(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        a = os.path.join(m, "G3", "ANEXOS", nome_anexo("P3.1", "Relatório de Gestão", "RD-104"))
        make_pdf(a, "igual")
        shutil.copy(a, a.replace(" - RD-104.pdf", " (1) - RD-104.pdf"))
        ach, _p, _i = auditar(m)
        assert tem(ach, R.FIXED, "cópia idêntica")

    def test_versoes_diferentes_bloqueiam(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_pdf(os.path.join(m, "G4", "ANEXOS", nome_anexo("P4.1", "Relatório de Portfólio", "GERAL")), "v1")
        make_pdf(os.path.join(m, "G4", "ANEXOS", nome_anexo("P4.1", "Relatório de Portfólio v2", "GERAL")), "v2")
        ach, _p, _i = auditar(m)
        assert tem(ach, R.BLOCKER, "VERSÕES CONFLITANTES")

    def test_parte_faltando(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        for k in (1, 3):
            make_pdf(os.path.join(m, "G2", "ANEXOS", nome_anexo("P2.2", f"Projeto Executivo - Parte {k:02d}", "RD-102A")))
        ach, _p, _i = auditar(m)
        assert tem(ach, R.BLOCKER, "Parte 02")

    def test_sem_rodovia_bloqueia(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_pdf(os.path.join(m, "G5", "ANEXOS", "P5.1 - Acompanhamento Jurídico - Parecer - Parceiro.pdf"))
        ach, _p, _i = auditar(m)
        assert tem(ach, R.BLOCKER, "não é rodovia válida")

    def test_indice_divergente(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        idx = os.path.join(m, IDX.nome_arquivo_indice("Setembro", 2026))
        IDX.escrever_indice(idx, "Setembro", 2026, {"G1": [("P1.1", "arquivo que nao existe.pdf")]})
        ach, _p, _i = auditar(m)
        assert tem(ach, R.WARNING, "SEM arquivo correspondente")
        assert tem(ach, R.WARNING, "NÃO listado")

    def test_lista_ausente_pede_revisao(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        for f in os.listdir(os.path.join(m, "G3")):
            if "Lista" in f:
                os.remove(os.path.join(m, "G3", f))
        ach, _p, _i = auditar(m)
        assert tem(ach, R.MANUAL, "Lista de Demandas não encontrada")


# ================================================================ pipeline completa (E2E)
class TestPipelineCompleta:
    def test_pacote_pronto_e_publicacao(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_xlsx(os.path.join(m, "G1", "ANEXOS", nome_anexo("P1.1", "Quadro Financeiro", "GERAL", ".xlsx")))
        make_pdf(os.path.join(m, "G1", "ANEXOS", nome_anexo("P1.1", "Quadro Financeiro", "GERAL")))
        make_docx(os.path.join(m, "G2", "Programa - Relatório Grupo 2 - Relatório Mensal N°15 - 09 2026.docx"))
        rc = PL.executar(m, preparar=True)
        out = out_dir(m)
        d = achados_de(out)
        fin = [a for a in d["achados"] if a["fase"] != "ORIGEM"]
        assert d["veredito"] == R.READY, [a for a in fin if a["status"] in (R.BLOCKER, R.MANUAL)]
        assert rc == 0
        final = R.lp(os.path.join(out, "03_FINAL", os.path.basename(m)))
        # Listas e Word do relatório: no STAGING sim, no FINAL não
        assert os.path.exists(os.path.join(out, "02_STAGING", "G1", "PROG_Lista de demandas G1.xlsx"))
        todos_final = [os.path.join(r, f) for r, _d, fs in os.walk(final) for f in fs]
        assert not [f for f in todos_final if "Lista de demandas" in f]
        assert not [f for f in todos_final if f.endswith(".docx") and "Índice" not in f]
        # planilha virou ZIP com o arquivo ORIGINAL (bytes idênticos = fórmulas preservadas)
        z = [f for f in todos_final if f.endswith("Quadro Financeiro - GERAL.zip")][0]
        with zipfile.ZipFile(z) as zz:
            assert zz.testzip() is None
            interno = zz.read(zz.namelist()[0])
        orig = open(os.path.join(m, "G1", "ANEXOS", nome_anexo("P1.1", "Quadro Financeiro", "GERAL", ".xlsx")), "rb").read()
        assert interno == orig
        # índice na raiz reconcilia
        assert tem(fin, R.PASS, "reconcilia 100%")
        # manifesto com SHA-256 de todos os arquivos
        man = json.load(open(os.path.join(out, "04_RELATORIOS", "manifesto.json"), encoding="utf-8"))
        assert all(len(x["sha256"]) == 64 for x in man["arquivos"])
        # PUBLICAÇÃO na própria pasta do mês (ambiente artificial): arquiva antes, depois fica idêntica ao manifesto
        plano = json.load(open(os.path.join(out, "plano_publicacao.json"), encoding="utf-8"))
        with pytest.raises(SystemExit):
            PL.publicar(m, "codigo-errado")
        assert os.path.exists(os.path.join(m, "G1", "PROG_Lista de demandas G1.xlsx"))  # nada mudou
        assert PL.publicar(m, plano["codigo"]) == 0
        assert not os.path.exists(os.path.join(m, "G1", "PROG_Lista de demandas G1.xlsx"))
        arq = ambiente["raiz"] / "Programa - 3. Desenvolvimento" / "G1 - Coordenacao" / "Desenvolvimento de Produtos" / "Controle Interno Entregas"
        # R.lp: o arquivo interno é gravado com prefixo de caminho longo; sem ele os.walk omite
        # subpastas além de MAX_PATH (o tmp_path do Windows + nomes da versão pública passam de 260)
        arquivados = [f for r, _d, fs in os.walk(R.lp(str(arq))) for f in fs]
        assert "PROG_Lista de demandas G1.xlsx" in arquivados  # original preservado
        log = json.load(open(os.path.join(out, "log_publicacao.json"), encoding="utf-8"))
        assert log["conferencia_final"] == "IDÊNTICA AO MANIFESTO"

    def test_pacote_bloqueado_nao_publica(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_pdf(os.path.join(m, "G2", "ANEXOS", "documento final.pdf"))
        rc = PL.executar(m, preparar=True)
        assert rc == 2
        plano = json.load(open(os.path.join(out_dir(m), "plano_publicacao.json"), encoding="utf-8"))
        with pytest.raises(SystemExit):
            PL.publicar(m, plano["codigo"])
        assert os.path.exists(os.path.join(m, "G2", "ANEXOS", "documento final.pdf"))

    def test_origem_alterada_apos_preparar_bloqueia_publicacao(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        PL.executar(m, preparar=True)
        plano = json.load(open(os.path.join(out_dir(m), "plano_publicacao.json"), encoding="utf-8"))
        make_pdf(os.path.join(m, "G3", "ANEXOS", nome_anexo("P3.2", "Plano novo", "RD-103")))
        with pytest.raises(SystemExit):
            PL.publicar(m, plano["codigo"])

    def test_fracionamento_acima_do_limite(self, ambiente, config_limite_pequeno):
        m = montar_mes_correto(ambiente["produtos"])
        make_pdf(os.path.join(m, "G2", "ANEXOS", nome_anexo("P2.4", "Relatório Mensal de Obras", "RD-104")),
                 "grande", paginas=5, imagem_ruido_kb=1)
        PL.executar(m, preparar=True)
        out = out_dir(m)
        final_anx = R.lp(os.path.join(out, "03_FINAL", os.path.basename(m), "G2", "ANEXOS"))
        partes = sorted(f for f in os.listdir(final_anx) if "Relatório Mensal de Obras" in f)
        assert len(partes) >= 2 and all("Parte 0" in p for p in partes), partes
        assert all(os.path.getsize(os.path.join(final_anx, p)) <= R.limite_bytes() for p in partes)
        d = achados_de(out)
        assert not tem([a for a in d["achados"] if a["fase"] == "FINAL"], R.BLOCKER, "Parte")

    def test_reexecucao_e_deterministica(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        make_xlsx(os.path.join(m, "G4", "ANEXOS", nome_anexo("P4.2", "Controle Financeiro", "GERAL", ".xlsx")))
        PL.executar(m, preparar=True)
        man1 = json.load(open(os.path.join(out_dir(m), "04_RELATORIOS", "manifesto.json"), encoding="utf-8"))
        PL.executar(m, preparar=True)
        man2 = json.load(open(os.path.join(out_dir(m), "04_RELATORIOS", "manifesto.json"), encoding="utf-8"))
        h = lambda man: sorted((x["caminho_relativo"], x["sha256"]) for x in man["arquivos"] if not x["nome"].endswith(".docx"))
        assert h(man1) == h(man2)

    def test_revisao_humana_registrada_libera_manual(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        rel = [f for f in os.listdir(os.path.join(m, "G5")) if f.endswith(".pdf")][0]
        make_pdf_visual(os.path.join(m, "G5", rel))
        PL.executar(m, preparar=True)
        out = out_dir(m)
        assert achados_de(out)["veredito"] == R.PENDENTE_HUMANO
        sha = H.sha256_arquivo(os.path.join(m, "G5", rel))
        json.dump([{"sha256": sha, "codigo": "assinatura_relatorio", "decisao": "aceito", "por": "teste",
                    "em": "2026-10-01", "justificativa": "conferido no original físico"}],
                  open(os.path.join(out, "revisoes_humanas.json"), "w", encoding="utf-8"))
        PL.executar(m, preparar=True)
        assert achados_de(out)["veredito"] == R.READY

    # ---------------------------------------------------- exceções da competência (DEC-057)
    def _revisoes(self, out, entradas):
        json.dump(entradas, open(os.path.join(out, "revisoes_humanas.json"), "w", encoding="utf-8"), ensure_ascii=False)

    def test_excecao_aceita_libera_blocker_de_nome_e_aparece_no_relatorio(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        arq = os.path.join(m, "G5", "ANEXOS", "P5.1 - Acompanhamento Jurídico - Parecer - Parceiro.pdf")
        make_pdf(arq)
        PL.executar(m, preparar=True)
        out = out_dir(m)
        assert achados_de(out)["veredito"] == R.BLOCKED
        sha = H.sha256_arquivo(arq)
        # "aceito" simples NÃO libera BLOCKER (só MANUAL)
        self._revisoes(out, [{"sha256": sha, "codigo": "nome", "decisao": "aceito", "por": "teste", "em": "2026-10-14"}])
        PL.executar(m, preparar=True)
        assert achados_de(out)["veredito"] == R.BLOCKED
        self._revisoes(out, [{"sha256": sha, "codigo": "nome", "decisao": "aceito_excecao", "por": "teste",
                              "em": "2026-10-14", "justificativa": "só nesta competência"}])
        PL.executar(m, preparar=True)
        assert achados_de(out)["veredito"] == R.READY
        md = open(os.path.join(out, "04_RELATORIOS", "RELATORIO_PRONTIDAO.md"), encoding="utf-8").read()
        assert "Exceções aceitas pela Coordenação" in md and "EXCEÇÃO ACEITA" in md

    def test_excecao_nao_vale_para_conteudo_alterado(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        arq = os.path.join(m, "G5", "ANEXOS", "P5.1 - Acompanhamento Jurídico - Parecer - Parceiro.pdf")
        make_pdf(arq, "versão 1")
        sha = H.sha256_arquivo(arq)
        PL.executar(m)
        out = out_dir(m)
        self._revisoes(out, [{"sha256": sha, "codigo": "nome", "decisao": "aceito_excecao", "por": "t", "em": "x"}])
        make_pdf(arq, "versão 2 — conteúdo diferente")
        PL.executar(m, preparar=True)
        assert achados_de(out)["veredito"] == R.BLOCKED

    def test_relatorio_do_grupo_designado_por_revisao(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        g5 = os.path.join(m, "G5")
        rel = [f for f in os.listdir(g5) if f.endswith(".pdf")][0]
        novo = os.path.join(g5, "P5.1 - Relatório de Acompanhamento Jurídico.pdf")
        os.replace(os.path.join(g5, rel), novo)
        PL.executar(m, preparar=True)
        out = out_dir(m)
        d = achados_de(out)
        assert d["veredito"] == R.BLOCKED and tem(d["achados"], R.BLOCKER, "ausente na raiz")
        self._revisoes(out, [{"sha256": H.sha256_arquivo(novo), "decisao": "relatorio_do_grupo", "por": "teste",
                              "em": "2026-10-14"}])
        PL.executar(m, preparar=True)
        d = achados_de(out)
        assert d["veredito"] == R.READY
        assert tem(d["achados"], R.PASS, "VALIDADA")

    def test_aceite_de_achado_sem_arquivo_por_grupo_e_etapa(self, ambiente):
        m = montar_mes_correto(ambiente["produtos"])
        for f in os.listdir(os.path.join(m, "G2")):
            if "Lista" in f:
                os.remove(os.path.join(m, "G2", f))
        PL.executar(m, preparar=True)
        out = out_dir(m)
        assert achados_de(out)["veredito"] == R.PENDENTE_HUMANO
        self._revisoes(out, [{"grupo": "G2", "etapa": "lista_demandas", "decisao": "aceito", "por": "teste",
                              "em": "2026-10-14", "justificativa": "G2 usa a Lista Mestre"}])
        PL.executar(m, preparar=True)
        assert achados_de(out)["veredito"] == R.READY

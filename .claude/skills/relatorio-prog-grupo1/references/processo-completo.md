# Processo Completo — Relatório Grupo 1 Programa (end-to-end)

Visão detalhada do processo mensal. Leia junto com as references específicas de cada etapa.
Ferramentas: **Python 3.12** com `win32com` (Word/Excel/PowerPoint COM), `openpyxl`, `pandas`,
`python-docx`, `python-pptx`, `PyMuPDF (fitz)`, `Pillow`. Office instalado (COM disponível).

## Contexto
Programa = Consórcio (Dynatest, Certificadora, Consultoria de Gestão – CG, Escritório Jurídico). O Relatório do
**Grupo 1** = Coordenação Geral Executiva. Produtos: **P1.1** (Coordenação Geral, Dynatest),
**P1.2** (Relações Institucionais, Dynatest), **P1.3** (Comunicação Integrada, CG),
**P1.4** (TIC, CG). O relatório **consolida entregas** de todos os grupos (G2 Dynatest/eng.,
G3 Certificadora/certificação, G4 CG, G5 Escritório Jurídico).

## Modelo de dados das entregas
- **Acumulado** `P.1.1 - Controle de demandas Programa.xlsx`, aba **"Controle Geral"** (cabeçalho na
  linha 3): colunas B=ID, C=Produto, D=Empresa, E=Documento, F=Tipo, G=Descrição, H=Empreendimento,
  I=Data Entrega, J=Status Entrega, **K=Versão Relatório**. É a **fonte do Power BI** (Dashboard).
- A competência é filtrada por **Versão Relatório = N** (N = número do relatório).
- **G2 (engenharia)**: NÃO vem da pasta mensal. Vem da **Lista Mestre** `00_ Lista mestre.xlsx`
  (aba "Controle Geral", cabeçalho linha 3), filtrando **N° Relatório = N**.
- **G1 (1.3/1.4), G3, G4, G5**: cada um tem `PROG_Lista de demandas ...xlsx` na **pasta mensal**
  (aba Grupo1/Grupo3/Grupo4/Grupo2, cabeçalho linha 3). Regra: usar **todas as linhas populadas**
  (a coluna "Versão" costuma estar vazia; as listas são do mês). Datas podem estar na data de
  compilação (início do mês seguinte) — isso é normal.
- **Recorrentes P1.1/P1.2** (produzidos pela Coordenação): clonar as linhas da versão anterior
  (N-1) no acumulado, atualizando nome do documento/mês/data(=data do documento) e Versão=N:
  `P1.1 - Relatório Grupo 1 - Relatório Mensal N°{n} - {periodo}`, `P1.1 - Timeline de Acompanhamento
  Macro até {mês} {ano}`, `P1.1 - Dashboard da Análise de Entrega de Produtos - {mês}`,
  `P1.1 - Controle de Entrega de Produtos de {mês}`, `P1.1 - Quadro Financeiro do mês de {mês}`,
  `P1.2 - Controle de Reuniões do mês de {mês}`. Empresa = Dynatest.
- **Total da competência** = G1 + G2 + G3 + G4 + G5 + 6 recorrentes. (RM13: 271 = 11+167+5+50+32+6.)

## Normalização de datas e status (na consolidação)
- Data real (datetime/serial) ou string conversível (dd/mm/aaaa, dd.mm.aaaa, aaaa-mm-dd) → data.
- Qualquer outro valor (N/A, "-", "em definição", vazio, texto) → **"Sem data"** (exatamente).
- Status default "Entregue" se vazio; empreendimento default "Geral".
- Dedup por chave composta (Produto+Empresa+Documento+Tipo+Descrição+Empreendimento); preferir data
  real a "Sem data"; registrar duplicidades no log.

## Fases de prontidão de uma competência (ver SKILL §6b)
1. **DESCOBERTA** — localizar as fontes do mês (pasta mensal; listas G1/G3/G4/G5; Lista Mestre G2;
   Controle de Demandas; Reuniões; Quadro Financeiro; relatório CG; Dashboard; Timeline; anexos) e
   registrar o que existe/falta.
2. **PREPARAÇÃO** — criar `config/parametros.json` e, se necessário, `config/dados_mes.py` **só** com o
   não derivável das fontes (DEC-019). Em MODO REAL, sem `dados_mes.py` a produção não roda.
3. **EXECUÇÃO** — primeiro `PROG_DRY_RUN=1`; depois produção real com fontes completas; QA obrigatório.

## Etapas (ver a tabela do SKILL.md para reference+script de cada uma)
0. **Preparar**: `python prog_config.py` (confere PROG_ROOT, caminhos, pasta mensal). Cria pasta de
   trabalho em `output/RM{n}_{ano}_{mm}/` (subpastas 01_FONTES..07_SAIDA + 03_IMAGENS/*). Backup do
   acumulado em `06_BACKUP` com timestamp. **Antes de ler as fontes**: `python scripts/sharepoint.py comparar-pasta
   "<…\Desenvolvimento de Produtos>" --recursivo` (e o `G1` do mês) → só seguir com "EM DIA" (DEC-052).
1. **Consolidar entregas** → `02_TRABALHO/entregas_rmN.json` + estatísticas (por empresa/status/
   produto/empreendimento/tipo).
2. **Atualizar acumulado** (backup → acrescentar linhas Versão=N; datas dd/mm/aaaa; IDs contínuos).
   Habilita o refresh do Power BI (feito pelo usuário depois).
3. **Anexo Controle de Entregas** (`P1.1 - Controle de Entrega de Produtos de {mês}.xlsx` na pasta
   P1.1) + **print** landscape (ver DEC-009).
4. **Dashboard**: o usuário faz refresh do Power BI, filtra o mês e **exporta o PDF** para a pasta de anexos do
   G1 (`cfg.anexo_p11_dir()`) com o nome de `cfg.nome_anexo_g1()`: POP = `P1.1 - Gestão Executiva - Dashboard da
   Análise de Entrega de Produtos - {mês} - GERAL.pdf`; legado = `P1.1 - Dashboard da Análise de Entrega de Produtos
   - {mês}.pdf` em `Anexos\P1.1`. Converter em imagem.
   (Cálculos dos indicadores vêm da base, não do screenshot.)
5. **Timeline**: usar o PPT corrigido; duplicar marcos reais; exportar PDF (histórico até o mês) →
   anexo `P1.1 - Timeline de Acompanhamento Macro até {mês} {ano}.pdf` + imagens. Se o placeholder
   `XXXX` não for encontrado, o `07` usa o último slide como fallback e marca a etapa para
   **conferência visual** no QA (DEC-020) — não aprovar sem revisar.
6. **Reuniões**: cópia mensal `P1.2 - Controle de Reuniões do mês de {mês}.xlsx` (pasta P1.2),
   filtrada ao mês; tabela em escala 1:1 com escolha automática de orientação e verificação de corte
   no PDF (ver `reunioes.md` e DEC-027; **exige a 2ª passagem 04→09** para medir o espaço real).
7. **Medição**: usar o Quadro Financeiro do mês (fonte primária); validar; gerar imagens (quadro
   consolidado, resumo por empresa, base técnica Parceiro) + Linha do Tempo de Medições (3 quadros).
8. **Recortar** todas as imagens (bordas brancas) e registrar aspectos (`img_aspects.json`).
9. **Redigir** P1.1 (5 seções) e P1.2 (1 seção "Integração Institucional…") fiéis ao RM anterior, a
   partir das **fontes reais** do mês (não do exemplo). Preencher `config/dados_mes.py` **apenas** com o
   não derivável das fontes (DEC-019/DEC-019b). **MODO REAL exige `config/dados_mes.py`**: sem ele,
   `cfg.dados()` interrompe com "DADOS MENSAIS NÃO CONFIGURADOS" (o exemplo de julho só entra em MODO
   TESTE, `PROG_USE_EXAMPLE_DATA=1`).
10. **Montar Word**: copiar a base CG (capa gráfica + P1.3 + P1.4 + cabeçalho/rodapé + TOC);
    remover os stubs de P1.1/P1.2 ("Não houve demandas… / Ou / Lista de Demandas") e inserir o
    conteúdo (títulos bullet não-negrito + texto + imagens dimensionadas); seções **paisagem** para
    Controle e Timeline; Dashboard sozinho na página; textos colados às imagens.
11. **Encerramento** (Cidade, data única, negrito seletivo) + **2 assinaturas** + **Quadro de Anexos**;
    remover blocos genéricos "Nome – Programa".
12. **Cabeçalho/rodapé paisagem** (full-width; não tocar retrato).
13. **Finalizar**: abrir no Word COM, atualizar campos + sumário (TOC), repaginar, exportar **PDF**,
    salvar DOCX+PDF em `07_SAIDA` e **versionar** na pasta `G1` do mês (`_v2/_v3…` se já existir). Sem gravação
    local possível: `PROG_DRY_RUN=1` + `scripts/sharepoint.py publicar-lote` (ver `sharepoint-api.md`).
14. **QA** (`qa.md`) — obrigatório. Veredito **APROVADO** só se não houver pendências, incluindo os
    avisos persistentes de `cfg.qa_warnings()` (ex.: fallback da Timeline, DEC-020).
15. **Organizar** cópias em `Obsoleto\Programa RM{n} {Mês} {ano} - Arquivos Gerados\{Corpo,Anexos}`.

## Estrutura da base CG (documento de trabalho)
O relatório da CG (`...RevCG.docx`) já traz: capa gráfica (imagem), Apresentação (com "N° {n}" e
período), headings "Produto 1.1/1.2/1.3/1.4" com "Descrição das atividades atribuídas" +
"Produtos desenvolvidos no período", e P1.3/P1.4 **preenchidos** pela CG. P1.1/P1.2 vêm como
**stub** ("Não houve demandas… / Ou / Lista de Demandas") — é onde entra nosso conteúdo. O
encerramento pode vir com "São Paulo," e sem data; e há tabelas genéricas de assinatura "Nome – Programa"
que devem ser removidas. Sumário = campo TOC (atualizar no fim). Ver `estrutura-relatorio.md`.

## Teste de reprodução (rodar mentalmente ao manter a Skill)
Dado apenas o projeto (CLAUDE.md + Skill + references + scripts + config + baseline) e o pedido
"Crie o Relatório Mensal nº N de {mês}/{ano}": é possível executar sem esta conversa? Se faltar
algo, gravar no arquivo adequado e repetir até "sim".

## Checkpoints (perguntar só quando necessário) — ver SKILL §5.

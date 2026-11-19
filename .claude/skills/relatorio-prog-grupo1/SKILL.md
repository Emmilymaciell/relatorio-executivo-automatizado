---
name: relatorio-prog-grupo1
description: >
  Produzir e atualizar mensalmente o Relatório de Produtos Mensal do Grupo 1 (antes chamado
  "Relatório Mensal") – Coordenação Geral Executiva do Programa: preparar fontes (SharePoint), atualizar o controle de demandas, gerar os
  anexos (Controle de Entregas, Dashboard, Timeline, Reuniões, Quadro Financeiro),
  redigir os textos de P1.1/P1.2 fiéis ao modelo, preservar P1.3/P1.4 da CG, gerar as
  imagens, montar o Word (capa, sumário, encerramento, assinaturas, anexos), validar (QA)
  e entregar DOCX+PDF na pasta do mês. Acione quando o usuário pedir para criar/atualizar
  o Relatório de Produtos Mensal / Relatório Mensal do Grupo 1 (ex.: "Crie o Relatório Mensal nº 16 de outubro
  de 2026"). Para auditar/empacotar a entrega de TODOS os grupos ao Parceiro use a Skill entrega-coordenacao.
---

# Skill: Relatório Grupo 1 – Programa

## 1. Quando acionar
Quando o usuário pedir para **produzir/atualizar o Relatório de Produtos Mensal do Grupo 1** (Coordenação
Geral Executiva) do Programa para um dado mês. Não use para outros grupos/relatórios. A conferência e o pacote de
G1–G5 para o Parceiro ficam na Skill `entrega-coordenacao` (POP de Entrega v2026-10-01 — `docs/dominio/`).

## 2. Parâmetros e dados de entrada (mensais)
- **Parâmetros** em `config/parametros.json` (copie de `parametros-exemplo.json`). Se o usuário der o
  mês/número na conversa, atualize o JSON. Campos: `numero_relatorio`, `mes`, `mes_numero`,
  `ano`, `periodo_inicio`, `periodo_fim`, `mes_anterior`, `data_documento` (pode ser null),
  `status_medicao` (pode ser null). Todos os scripts leem via `scripts/prog_config.py`.
- **Conteúdo do mês** em `config/dados_mes.py` (lido por `cfg.dados()`). **Não é um relatório manual**:
  guarda só o que **não** se deriva com segurança das fontes (ex.: `MARCOS_EXTRAS`, `STATUS_MEDICAO`
  quando necessário, `OBSERVACOES_EXTRAORDINARIAS`, `INSTRUCOES_ESPECIFICAS_MES`,
  `AJUSTES_MANUAIS_AUTORIZADOS`, parâmetros de QA). Os textos de P1.1/P1.2 e os `MARCOS` deriváveis
  vêm das **fontes reais** (DEC-019/DEC-019b).
- **MODO REAL × MODO TESTE (DEC-019):** em execução real, `cfg.dados()` exige `config/dados_mes.py`
  (ou `PROG_DADOS` apontando um `.py`); **sem isso a execução PARA** com `DADOS MENSAIS NÃO
  CONFIGURADOS`. O exemplo `assets/texts_exemplo_julho2026.py` **só** é usado como fixture quando o
  operador define **`PROG_USE_EXAMPLE_DATA=1`** — nunca automaticamente.

## 3. Regras críticas (resumo — detalhes nas references)
- Segurança SharePoint: só o **acumulado** `P.1.1 - Controle de demandas Programa.xlsx` pode ser
  ALTERADO (com backup, só acrescentando linhas). Nada mais é movido/excluído/sobrescrito.
- **PRODUZIR DO ZERO das FONTES OFICIAIS todo mês (DEC-034/035/036/037).** Fontes por componente:
  entregas = listas G1/G3/G4/G5 + Lista Mestre G2; **medição = `Medicao Parceiro\Planilha de Medição -
  Estruturadora - {Mês}.xlsx`** (painel = copiar o do mês anterior + preencher da Planilha); **timeline
  = marcos de `P1.1 -Controle Timeline.xlsx`** (conferir; se faltar, PERGUNTAR ao usuário; depois gerar
  o slide); **reuniões = filtrar o mês do mestre `P1.2 - Controle de Reuniões.xlsx`**; **dashboard = PDF
  exportado pelo usuário** do Power BI; base CG = `…RevCG.docx`. **NUNCA** usar como fonte o espelho
  `5. Compartilhado Parceiro` nem `Obsoleto` (são cópias de entrega). Faltou fonte oficial → **perguntar**.
- **Nunca inventar dados.** Fidelidade ao layout/textos do RM anterior (só troca de dados).
- Datas em Excel no formato **dd/mm/aaaa** (PT-BR) e **sem hora** (senão o Power BI quebra).
- **POP de Entrega (01/10/2026):** a partir de **Outubro/2026** a pasta do G1 tem `ANEXOS` **única** (sem
  subpastas P1.x) e os anexos gerados seguem `[Código] - [Produto] - [Arquivo] - [Rodovia]` — automático por
  `cfg.layout_anexos()`/`cfg.nome_anexo_g1()` (DEC-043; meses antigos seguem o layout legado). O arquivo final se
  chama **Relatório de Produtos Mensal** (`cfg.termo_documento()`, DEC-044). Depois da publicação do pacote, a Lista
  de Demandas fica no arquivo interno e `cfg.lista_demandas_path()` a encontra lá (DEC-045).
- **SharePoint: pasta sincronizada + API (DEC-052).** Os scripts leem e gravam pela pasta sincronizada. A API
  (`python scripts/sharepoint.py`) confere se essa pasta está em dia **antes de gerar** (`comparar-pasta`) e
  publica fontes recebidas fora do SharePoint (medição, Timeline preenchida) ou os produtos, quando a gravação
  local não for possível (`enviar`/`publicar-lote`, dry-run → autorização → `--aplicar`). Ver
  `references/sharepoint-api.md`.
- **A regra mais recente do usuário prevalece.** Sempre leia `references/decisoes-e-erros-conhecidos.md`
  ANTES de gerar prints/textos, para não repetir erros já corrigidos.

## 4. Workflow — ordem obrigatória
Execute na ordem. Antes de cada etapa, **leia a reference indicada** (progressive disclosure).
Os scripts estão em `scripts/` e leem `config/parametros.json` via `prog_config.py`.

| # | Etapa | Reference a ler | Script |
|---|-------|-----------------|--------|
| 0 | Preparar: localizar PROG_ROOT, criar pasta de trabalho, backup do acumulado; **conferir a sincronização** de `Desenvolvimento de Produtos` e do `G1` do mês (não gerar com pasta desatualizada) | `fontes-sharepoint.md`, `processo-completo.md`, `sharepoint-api.md` | `prog_config.py` (testar); `scripts/sharepoint.py comparar-pasta … --recursivo` |
| 0b | **Insumos de fora do SharePoint** (Downloads, chat, e-mail): importar ANTES de usar — cópia imutável + SHA-256 + proveniência + cópia em `Fontes Externas` (DEC-059). OneDrive sem hidratar ⇒ espelho de leitura + `PROG_ROOT` (DEC-066) | `fontes-sharepoint.md`, DEC-059/066 | `fontes_externas.py inspecionar/importar/plano-sharepoint`; `scripts/sharepoint.py espelhar` |
| 1 | Consolidar entregas da competência (G1 1.3/1.4, G2 Lista Mestre por Nº Relatório, G3/G4/G5, + recorrentes P1.1/P1.2) — **todas** as linhas declaradas (DEC-067; sem corte por data; repetidas mantidas e registradas); G2: conciliar Lista Mestre × `G2\ANEXOS` | `processo-completo.md`, `regras-excel.md`, DEC-067 | `01_consolidar_entregas.py` |
| 2 | Atualizar acumulado (backup → acrescentar linhas da versão N; datas dd/mm/aaaa; IDs contínuos) | `regras-excel.md`, `decisoes-e-erros-conhecidos.md` | `02_atualizar_controle_demandas.py` |
| 3 | Anexo + print Controle de Produtos (landscape, SEM ID, sem cortar texto, paginação balanceada) | `regras-excel.md`, `regras-layout.md` | `03_print_controle_produtos.py` |
| 4 | Dashboard: **usuário exporta** o PDF do Power BI (o acumulado já está em Versão N) → colocar na pasta de anexos do G1 (`cfg.anexo_p11_dir()`: layout POP = `G1\Anexos` com o nome `P1.1 - Gestão Executiva - Dashboard da Análise de Entrega de Produtos - {mês} - GERAL.pdf`; legado = `G1\Anexos\P1.1`) → **gerar a imagem** `dashboard_p01.png`. **ETAPA OBRIGATÓRIA antes do 09** (foi esquecida no RM14 — o `18` acusa se faltar) — DEC-040 | `dashboard.md` | `19_dashboard_imagem.py` |
| 5 | Timeline: **primeiro levar os eventos de Prestação de Contas do mês para o Controle Timeline** (iterativos: envio/análise/retorno/aprovação; Obra × Estruturadora confirmado no SharePoint — DEC-060/068; `prestacao_contas.py` → `timeline_prestacao_excel.py`; mês novo no PPTX: `timeline_slide_mes.py`); depois **ler os marcos de `P1.1 -Controle Timeline.xlsx`** (aba Base Timeline, competência do mês, Incluir=Sim). **Conferir se estão atualizados; se faltarem/incompletos, PERGUNTAR ao usuário e preencher o Excel ANTES de gerar** (DEC-035; preencher numa cópia e publicar com `sharepoint.py enviar --substituir`, que faz backup). Gerar o **slide mensal** no PPTX-base (preencher os grupos de marcador posicionados, na ordem por data — sem duplicar) + exportar PDF e renderizar **TODAS as slides** (acumulada: 1º mês → corrente — DEC-028). Nunca usar PDF do espelho Parceiro | `timeline.md` | `07_timeline.py` |
| 6 | Reuniões: **produzir o anexo mensal filtrando o mês do MESTRE** `P1.2 - Controle de Reuniões.xlsx` (DEC-036); tabela em **escala 1:1** (largura = corpo), fonte legível, negrito/alinhamentos, quebras entre reuniões completas, **retrato×paisagem automático**. Nunca pegar mensal pronto do Parceiro | `reunioes.md`, `regras-layout.md` | `04_print_reunioes.py` |
| 7 | Medição (nomeada pela **competência de execução** — DEC-063; Excel oficial ausente mas quadro impresso pelo Parceiro ⇒ planilha reconstruída e conferida campo a campo, `medicao_reconstrucao.py` — DEC-062): **fonte = `Medicao Parceiro\Planilha de Medição - Estruturadora - {Mês}.xlsx`** (DEC-034). **Produzir o painel** copiando o `P1.1 - Quadro Financeiro` do **mês anterior** (template) e preenchendo da Planilha (win32 p/ preservar o gráfico; validar totais). Depois gerar imagens (quadro consolidado recortado no fim do RESUMO + largura = corpo — DEC-029; **Base Técnica em 2** — DEC-031); **Linha do Tempo 5+5+resto do STATUS_MEDICOES com legendas** — DEC-030 (`06` após `05`). Nunca usar o painel do espelho Parceiro | `medicao.md` | `15_gerar_painel_financeiro.py` (novo), `05_imagens_financeiro.py`, `06_linha_tempo_medicoes.py` |
| 8 | Recortar TODAS as imagens (bordas brancas) | `regras-layout.md` | `08_recortar_imagens.py` |
| 8b | Anexos .xlsx de P1.1/P1.2 → **ZIP individual** na pasta ANEXOS (nome final usado no Quadro de Anexos e no Índice) — DEC-066 | `anexos.md`, DEC-066 | `20_zip_anexos_g1.py` |
| 9 | Redigir P1.1/P1.2 a partir das **fontes reais** (não do exemplo), seguindo o **texto-base fixo** (P1.1 Reuniões de Governança ≈7 parágrafos por tema/agente + síntese — DEC-026; **P1.2: texto introdutório + tabela de reuniões + texto consolidado "Relatório de Relacionamento Institucional" após a tabela, redigido do modelo + planilha de reuniões 1.2 — obrigatório, DEC-033**). **+ Síntese do Plano de Gestão Integrada de Riscos** a partir do P4.3 do mês (`RISCOS` — DEC-065); insumo pendente ⇒ `PENDENCIAS` (quadro "SEÇÃO EM ELABORAÇÃO", arquivo "… - EM ELABORAÇÃO" — DEC-064); data da capa/encerramento = fim da competência (DEC-058). Preencher `config/dados_mes.py` só com o **não derivável** (`MARCOS_EXTRAS`, `STATUS_MEDICAO` s/ necessário, `OBSERVACOES_EXTRAORDINARIAS`, `AJUSTES_MANUAIS_AUTORIZADOS`, params de QA; e `MARCOS`/`MEDICOES`/`MEDICAO_NUM`/`QA_ESPERADOS` quando não deriváveis). **MODO REAL exige este arquivo** (DEC-019) | `regras-redacao.md`, `estrutura-relatorio.md` | criar `config/dados_mes.py` (estrutura em `assets/texts_exemplo_julho2026.py`); lido por `cfg.dados()` |
| 10 | Montar Word: base CG (capa+P1.3+P1.4), inserir P1.1/P1.2, seções paisagem, imagens dimensionadas; **cada título "Produto 1.X" inicia em página nova** (DEC-032) | `regras-layout.md`, `estrutura-relatorio.md` | `09_montar_word.py` |
| 11 | Encerramento (Cidade, data única, negrito seletivo) + 2 assinaturas + Quadro de Anexos | `estrutura-relatorio.md`, `anexos.md` | `10_encerramento_assinaturas_anexos.py` |
| 12 | Cabeçalho/rodapé das páginas paisagem (full-width; portrait intocado); **remove "primeira página diferente" de todas as seções** p/ nenhuma página ficar sem cabeçalho/rodapé (DEC-032) | `regras-layout.md` | `11_cabecalho_rodape_paisagem.py` |
| 13 | Finalizar: atualizar sumário/campos, exportar PDF, versionar em G1 (se a gravação local não for possível: `PROG_DRY_RUN=1` + `sharepoint.py publicar-lote`; depois `comparar-pasta` no G1) | `processo-completo.md`, `sharepoint-api.md` | `12_finalizar_pdf.py` |
| 14 | QA (obrigatório) | `qa.md` | `13_qa.py` |
| 15 | Organizar arquivos gerados (Corpo/ + Anexos/) em Obsoleto | `processo-completo.md` | `14_organizar_arquivos_gerados.py` |
| 16 | **REVISÃO FINAL (obrigatória, última etapa)** — confere que nada faltou na montagem/anexos: imagens do corpo (pega **Dashboard esquecido**), Dashboard no PDF, Timeline com nº de páginas coerente, status da medição no painel, valores esperados, e **cada anexo da lista existe fisicamente**. Só entregar com REVISÃO FINAL = APROVADO (DEC-040) | `qa.md` | `18_revisao_final.py` |

### 4b. Segunda passagem obrigatória (DEC-027)
Depois do `09`, **reexecute `04` e `09`** (e siga daí com `10`→`14`). Motivo: a 1ª passagem só consegue
*estimar* o espaço que sobra abaixo do texto do P1.1 (estimativa enviesada em ~0,30"); a 2ª mede no
**DOCX montado** onde a tabela realmente cai e reagrupa as reuniões para preencher a página. O texto
acima da tabela não muda entre passagens, então a medida converge numa única repetição.

## 5. Checkpoints (pergunte só nestes casos)
- `data_documento` ausente e não encontrável na capa da CG.
- `status_medicao` não determinável pela planilha financeira.
- Export do **Power BI** (Dashboard) — exige interface/autenticação; peça o PDF ao usuário.
- Timeline: se não houver PPT atualizado/corrigido para o mês, peça-o (a versão do corpo é a base).
- Duas fontes oficiais conflitantes sem critério seguro.
Em todos: conclua tudo o que é independente do bloqueio, salve checkpoint, e só então pergunte.

## 6. Scripts reutilizáveis
Todos em `scripts/`, **totalmente parametrizados** por `config/parametros.json` (via `prog_config.py`)
e pelo conteúdo do mês (`config/dados_mes.py` via `cfg.dados()`). Nenhum valor da competência é fixo.
Ver `scripts/README.md` para o papel de cada um, a ordem de execução e os modos por variável de ambiente:
- `PROG_DRY_RUN=1` — somente-leitura: `02`/`03`/`12`/`14` não escrevem nos arquivos oficiais do
  SharePoint. Em especial, o `12` gera DOCX/PDF só em `07_SAIDA` e **não** publica no `G1` (DEC-021).
- `PROG_USE_EXAMPLE_DATA=1` — **MODO TESTE/FIXTURE**: permite `cfg.dados()` usar o exemplo de Julho/2026
  quando não há `config/dados_mes.py`. **Nunca** ligar numa execução real (DEC-019).
- `PROG_WORK_DIR=<pasta>` — redireciona **toda** a saída/temporários da execução (subpastas
  `01_FONTES..07_SAIDA`) para essa pasta, **sem** alterar a raiz permanente `PROG_PROJECT_DIR` (DEC-023).
- `PROG_REUNIOES_FONTE_PT=<pt>` — alvo **visual** da fonte do conteúdo da tabela de reuniões
  (default 9,5; escala 1:1, então é o tamanho que sai no PDF). `PROG_REUNIOES_ORIENTACAO=auto|retrato|
  paisagem` força a orientação (default `auto` = a de menor desperdício). `PROG_BASE_DOCX=<docx>` fixa
  o relatório-base (útil quando a base do mês está em subpasta, ex.: `G1\Testes`).
- `PROG_TIMELINE_PPTX=<arquivo.pptx>` — força a fonte da Timeline (senão usa o editável acumulado
  compartilhado em *Desenvolvimento de Produtos* → anexo P1.1 → `01_FONTES`) (DEC-025).
- Timeline (`07`): fonte = **editável acumulado compartilhado** (1 slide/mês); datas iguais = **1 ponto
  com 2 setas** + data única; descrições **escalonadas**; setas de comprimento variável (DEC-025). Sem
  o `.pptx`-fonte, interrompe com "TIMELINE PPTX NÃO LOCALIZADA" + pendência de QA (DEC-022).
Regras dos scripts: aceitam parâmetros; geram log; não excluem; não sobrescrevem original
sem versionar; backup antes de alterar o acumulado; falham de forma clara; re-executáveis (idempotentes).
Timeline: se o placeholder `XXXX` não for encontrado, o `07` usa o último slide como fallback, registra
WARNING e marca a etapa para **conferência visual** no QA (DEC-020) — não aprova automaticamente.

## 6b. Fases de uma nova competência (prontidão)
1. **DESCOBERTA** — conferir a sincronização (`sharepoint.py comparar-pasta`) e localizar as fontes do mês: pasta mensal em `4. Produtos`; `PROG_Lista de demandas`
   G1/G3/G4/G5; **Lista Mestre G2**; **Controle de Demandas** (acumulado); **Controle de Reuniões**;
   **Quadro Financeiro**; **relatório CG** (base); **Dashboard** (PDF do Power BI, exportado pelo
   usuário); **Timeline** (PPTX do mês); demais anexos. Registrar o que existe e o que falta.
2. **PREPARAÇÃO** — criar `config/parametros.json` (nº, mês, período, `data_documento`/`status_medicao`
   se conhecidos) e, **se necessário**, `config/dados_mes.py` **somente** com o não derivável das fontes
   (DEC-019/DEC-019b). Em MODO REAL, sem `dados_mes.py` a produção **não** roda.
3. **EXECUÇÃO** — primeiro `PROG_DRY_RUN=1` (valida sem tocar em arquivos oficiais); depois a produção
   real quando as fontes estiverem completas; QA obrigatório ao fim (inclui os avisos de DEC-020).

## 7. Critérios de conclusão
- DOCX+PDF finais na pasta `G1` do mês (versionados, sem sobrescrever).
- QA APROVADO (`references/qa.md`) — nº/mês/período coerentes; sem placeholders; imagens
  legíveis sem invadir cabeçalho/rodapé; **cabeçalho/rodapé em TODAS as páginas e cada "Produto 1.X"
  em página nova** (checagens automáticas do `13` — DEC-032); anexos existem fisicamente; nada
  destrutivo no SharePoint.
- Cópia dos arquivos gerados em `Obsoleto\...Arquivos Gerados\{Corpo,Anexos}`.
- Resumo de entrega ao usuário (caminhos, números, medição, QA, pendências).
- Depois da **assinatura** (externa), o PDF assinado volta para a raiz do `G1` e a Coordenação roda
  `python scripts/entrega.py auditar --mes N` (Skill `entrega-coordenacao`) — o DOCX e o PDF não assinado ficam
  fora do pacote automaticamente.

## 8. Manutenção da Skill (quando o usuário disser "vire padrão / atualize a Skill")
1. Identifique exatamente qual regra mudou (assunto + reference).
2. Atualize a **reference** adequada (substitua a regra antiga pela nova).
3. Atualize `SKILL.md` **só** se a mudança afeta o workflow central.
4. Atualize o **script** correspondente se houver comportamento determinístico.
5. Registre em `references/decisoes-e-erros-conhecidos.md` (novo DEC ou atualização), preservando
   o histórico (o que era antes, por que mudou).
6. **Não** acrescente regra conflitante no fim — substitua a vigente.
7. Rode o teste de reprodução (§ "Teste" em `processo-completo.md`) e valide o QA.
8. Informe ao usuário exatamente quais arquivos da Skill foram alterados.

# Scripts — Relatório Grupo 1 Programa

Python 3.12 + `win32com` (Office COM), `openpyxl`, `python-docx`, `python-pptx`, `PyMuPDF`, `Pillow`.
Parametrização via **`prog_config.py`** (lê `config/parametros.json` + `config/caminhos.json`;
localiza `PROG_ROOT`; deriva a pasta mensal; fornece todos os caminhos e a pasta de trabalho
`output/RM{n}_{ano}_{mm}/`). Rode `python prog_config.py` para conferir os caminhos antes de tudo.

## Ordem de execução (ver a tabela do SKILL.md)
| Script | Papel | Reference |
|--------|-------|-----------|
| `prog_config.py` | Config/caminhos/pasta de trabalho/log (importado por todos) | fontes-sharepoint.md |
| `01_consolidar_entregas.py` | Consolida entregas da competência → JSON + stats | processo-completo.md, regras-excel.md |
| `02_atualizar_controle_demandas.py` | Backup + acrescenta linhas da versão N no acumulado | regras-excel.md, DEC-001/002 |
| `03_print_controle_produtos.py` | Anexo + print Controle (landscape, sem ID, balanceado) | regras-excel.md, DEC-009 |
| `04_print_reunioes.py` | Cópia mensal + tabela Reuniões em escala 1:1 (calibrada), negrito/alinhamentos, corte verificado no PDF, retrato×paisagem automático | reunioes.md, DEC-027 |
| `05_imagens_financeiro.py` | Imagens do Quadro Financeiro (consolidado recortado no fim do RESUMO+largura=corpo DEC-029; **Base Técnica dividida em 2** DEC-031) | medicao.md, DEC-014/029/031 |
| `06_linha_tempo_medicoes.py` | Quadros Linha do Tempo de Medições (blocos de 5; **lidos do Excel STATUS_MEDICOES**, com legendas; fallback `cfg.dados()`) | medicao.md, DEC-014/030 |
| `07_timeline.py` | Timeline (copia marcadores reais do PPT) → PDF/imagens | timeline.md, DEC-013 |
| `08_recortar_imagens.py` | Recorta bordas brancas de todas as PNG + aspectos | regras-layout.md, DEC-016 |
| `09_montar_word.py` | Monta o Word (P1.1/P1.2, seções paisagem, fit imagens) | regras-layout.md, estrutura-relatorio.md |
| `10_encerramento_assinaturas_anexos.py` | Encerramento + 2 assinaturas + Quadro de Anexos | estrutura-relatorio.md, anexos.md, DEC-005/006/007 |
| `11_cabecalho_rodape_paisagem.py` | Cabeçalho/rodapé full-width nas seções paisagem | regras-layout.md, DEC-015 |
| `12_finalizar_pdf.py` | Atualiza TOC/campos, exporta PDF, versiona em G1 | processo-completo.md |
| `13_qa.py` | QA (varredura do PDF): textos/nº/período/medição + **layout: cabeçalho/rodapé em todas as páginas e cada "Produto 1.X" em página nova** (DEC-032) | qa.md |
| `14_organizar_arquivos_gerados.py` | Copia gerados p/ Obsoleto\...\{Corpo,Anexos} | processo-completo.md |
| `15_gerar_painel_financeiro.py` | **PRODUZ o painel Quadro Financeiro** copiando o do mês anterior (template) + preenchendo da `Planilha de Medição - Estruturadora - {Mês}` (Medicao Parceiro); win32 preserva o gráfico; `--validate ORACULO` confere célula-a-célula (DEC-034) | medicao.md |
| `16_reunioes_do_mestre.py` | **PRODUZ o anexo mensal de Reuniões** filtrando o mês do mestre `P1.2 - Controle de Reuniões.xlsx` (roda ANTES do `04`). Se o mestre não tiver o mês, avisa (DEC-036) | reunioes.md |
| `17_timeline_do_excel.py` | **Timeline dos marcos do Excel** `P1.1 -Controle Timeline.xlsx` (Base Timeline, competência do mês, Incluir=Sim): confere/preenche o slide do mês no PPTX-base e renderiza todas as slides (DEC-035/028). Se o slide não tiver os marcos, avisa (não corrompe) | timeline.md |
| `18_revisao_final.py` | **REVISÃO FINAL (última etapa, após o 13/QA)** — vai além do QA de texto: (1) cada pasta de imagens do corpo tem PNG (pega **Dashboard esquecido**); (2) no PDF: nº/período, cabeçalho+rodapé em todas as páginas, cada Produto 1.X no topo, texto consolidado P1.2, **imagem do Dashboard presente**, nº de páginas da Timeline coerente, valores QA_ESPERADOS, datas de reuniões do mês; **status da medição** conferido no painel-anexo (RESUMO); (3) **cada anexo da "LISTA DE DOCUMENTOS ANEXOS" existe fisicamente** nas pastas. Veredito REVISÃO FINAL (DEC-040) | qa.md |
| `19_dashboard_imagem.py` | **Gera `dashboard_p01.png`** do PDF do Dashboard (exportado pelo usuário) em `G1\Anexos\P1.1`. **Rodar ANTES do 09** — esquecer isso deixa o relatório sem o Dashboard (RM14); o `18` acusa. Se o PDF faltar, avisa (checkpoint) — DEC-040 | dashboard.md |

> **Ordem completa (produção mensal):** `01`→`02`→ `16` (reuniões do mestre, ou usar o arquivo do mês) → `03`→`04` → `15` (painel) → `05`→`06` → `17` (timeline) → **`19` (Dashboard)** → `08`→`09` → (2ª passagem `04`+`09`) → `10`→`11`→`12`→`13` (QA) → `14` (organizar) → **`18` (REVISÃO FINAL)**. Só entregar/publicar com QA **e** REVISÃO FINAL = APROVADO.

> **Ordem com os geradores novos (produção do zero):** …`02`→ `16` (reuniões do mestre) → `04` → `15` (painel) → `05`→`06` → `17` (timeline) → `08`→`14`. `15/16/17` substituem a etapa de "pegar pronto"; nunca usar o espelho Parceiro (DEC-037).

`../assets/texts_exemplo_julho2026.py` = **FIXTURE de desenvolvimento** (EXEMPLO Julho/2026):
textos P1.1/P1.2 + segmentos de negrito da medição **e** os dados variáveis `MARCOS` (timeline),
`MEDICOES`/`MEDICOES_STATUS`/`MEDICOES_EMPRESAS` (linha do tempo de medições), `MEDICAO_NUM` e
`QA_ESPERADOS`. Serve como **referência de estrutura** do `config/dados_mes.py` e como fixture de teste.

### `cfg.dados()` — MODO REAL × MODO TESTE (DEC-019)
- **MODO REAL** (padrão): resolve `env PROG_DADOS` (caminho `.py` explícito) → `config/dados_mes.py`.
  Se **nenhum** existir, **PARA** com `RuntimeError` **"DADOS MENSAIS NÃO CONFIGURADOS"** — não usa o
  exemplo, para impedir contaminação factual de outra competência.
- **MODO TESTE/FIXTURE**: só com **`PROG_USE_EXAMPLE_DATA=1`** (explícito, nunca automático) o exemplo de
  Julho/2026 é carregado, com WARNING no console.
- **`dados_mes.py` não é um relatório manual** (DEC-019b): textos analíticos de P1.1/P1.2 e `MARCOS`
  deriváveis vêm das **fontes reais**; o arquivo guarda só o não derivável (`MARCOS_EXTRAS`,
  `STATUS_MEDICAO` quando necessário, `OBSERVACOES_EXTRAORDINARIAS`, `INSTRUCOES_ESPECIFICAS_MES`,
  `AJUSTES_MANUAIS_AUTORIZADOS`, parâmetros de QA).

## Como rodar para um novo mês (3 fases — ver SKILL §6b)
**FASE 1 — DESCOBERTA:** localizar as fontes do mês (pasta mensal, listas G1/G3/G4/G5, Lista Mestre G2,
Controle de Demandas, Reuniões, Quadro Financeiro, relatório CG, Dashboard, Timeline, anexos).
**FASE 2 — PREPARAÇÃO:**
1. `cp config/parametros-exemplo.json config/parametros.json` e ajustar (nº, mês, ano, período,
   `data_documento`).
2. Criar `config/dados_mes.py` **somente** com o não derivável das fontes (`MARCOS_EXTRAS`,
   `STATUS_MEDICAO` s/ necessário, `OBSERVACOES_EXTRAORDINARIAS`, `AJUSTES_MANUAIS_AUTORIZADOS`,
   parâmetros de QA). Estrutura de referência em `assets/texts_exemplo_julho2026.py` (é FIXTURE, não
   copiar dados de julho). **MODO REAL exige este arquivo** (DEC-019); sem ele a produção não roda.
3. `python prog_config.py` para conferir os caminhos (deve apontar para a pasta do mês, ex.: `14. Agosto 2026`).
**FASE 3 — EXECUÇÃO:**
4. `set PROG_DRY_RUN=1` para validar sem escrever em arquivos oficiais (02/03/14).
5. Rodar `01`→`14` na ordem (produção real quando as fontes estiverem completas). QA obrigatório ao fim.

> Testes/desenvolvimento sem dados reais: `set PROG_USE_EXAMPLE_DATA=1` autoriza o fixture de Julho/2026
> (WARNING no console). **Nunca** ligar isso numa execução real (DEC-019).

## Regras de todos os scripts
Aceitam parâmetros (config, não hardcode); geram log (`cfg.log`); **não excluem** arquivos; **não
sobrescrevem** original sem versionar; **backup** antes de alterar o acumulado; falham de forma clara;
re-executáveis sem duplicar; saída determinística sempre que possível; fecham o Office (`atexit`+Quit)
e limpam processos órfãos.

## ✅ Estado de parametrização (MIGRAÇÃO CONCLUÍDA)
Todos os scripts operacionais (`01`–`11`, `13`, `14`) foram **migrados para o padrão do
`prog_config.py`**. Nenhum contém mais caminho absoluto da sessão do RM13, pasta de trabalho fixa,
`"julho"`, `"13"`, período de julho ou nomes de anexos fixos — tudo vem de `config/parametros.json`
(via `cfg`) e do conteúdo do mês (via `cfg.dados()`). A LÓGICA validada que produziu o baseline v5
foi **preservada** (mesmo algoritmo/ordem/constantes de layout); só mudou a **origem dos valores**.

Melhorias de robustez introduzidas na migração:
- `03` agora **cria o anexo** Controle de Entregas (filtra Versão N, ordena, renumera 1..n) **e** gera
  o print do corpo — antes a criação do anexo era um passo avulso da sessão.
- `05` localiza as abas do Quadro Financeiro por **prefixo** (o sufixo do nº da medição muda a cada mês).
- `06` divide as medições em **blocos de 5** automaticamente (não mais fixo em 1-5/6-10/11-12).
- `07` **descobre** o slide-alvo (placeholder `XXXX`) e o slide-modelo por conteúdo (sem nº de slide fixo)
  e deriva o nº de marcos de `MARCOS`. Se `XXXX` não for achado, usa o último slide como fallback,
  loga WARNING e marca **"CONFERÊNCIA VISUAL DA TIMELINE NECESSÁRIA"** no QA (DEC-020).
- `09`/`10`/`11`/`13` derivam títulos, encerramento, nomes de anexos e checagens de mês/ano/nº/medição.

### DRY_RUN / somente-leitura
`PROG_DRY_RUN=1` faz os scripts que tocam arquivos oficiais **não escreverem no SharePoint**:
- `02` não abre/altera o acumulado (só reporta o que faria, após simular backup).
- `03` gera anexo/print **só no work**, sem copiar para a pasta P1.1.
- `12` gera DOCX/PDF **só em `07_SAIDA`** e **não publica no `G1`** (log "DRY_RUN: publicação no
  SharePoint ignorada"). Em execução real, versiona no `G1` sem sobrescrever (`_v2/_v3…`). (DEC-021)
- `14` grava o arquivo-organizador em `output/RM.../07_SAIDA/Arquivos_Gerados` (destino local).
Os demais scripts só escrevem em arquivos locais de trabalho (seguros por natureza).

### Isolamento de execução — `PROG_WORK_DIR` (DEC-023)
`PROG_WORK_DIR=<pasta>` redireciona **toda** a saída/temporários (`01_FONTES..07_SAIDA`, imagens,
logs, backups, QA-flags) para essa pasta, **sem** alterar `PROG_PROJECT_DIR` (raiz permanente: config/,
Skill). Se não informado, usa o padrão `output/RM{n}_{ano}_{mm}`. Como todos os scripts usam
`cfg.work_dir()`, basta setar a env — nenhum script precisa de edição.

### Timeline — fonte e técnica (DEC-025)
Fonte = **editável acumulado compartilhado** (1 slide/mês) em *Desenvolvimento de Produtos*
(`cfg.timeline_pptx_src()`: env `PROG_TIMELINE_PPTX` → compartilhado → anexo P1.1 → `01_FONTES`).
O `07` preenche a slide do mês duplicando os marcos-modelo (`DD/MM`/`XXXXXX`), com:
**datas iguais = 1 ponto + 2 setas** (data única), **descrições escalonadas** (perto/longe alternado),
**setas de comprimento variável** (do ponto até logo abaixo/acima do texto), largura ≤ espaçamento.
Gera um **ponto de partida**; o editável é refinável à mão e sobrescreve o compartilhado (com backup).

### Timeline sem `.pptx`-fonte (DEC-022)
Se o `07` não localizar o `.pptx`, ele **interrompe só a etapa da Timeline** com "TIMELINE PPTX NÃO
LOCALIZADA", registra pendência de QA e **não** inventa nem usa o PDF anterior. Fornecido o `.pptx`,
reexecutar o `07` limpa a pendência. (Fallback do `XXXX` quando o `.pptx` existe continua sob DEC-020.)

### MODO REAL × MODO TESTE (dados do mês) — DEC-019
- **Execução real**: exige `config/dados_mes.py` (ou `PROG_DADOS` apontando um `.py`). Sem isso,
  `cfg.dados()` **para** com `DADOS MENSAIS NÃO CONFIGURADOS` — jamais usa o exemplo de julho.
- **Teste/desenvolvimento**: `PROG_USE_EXAMPLE_DATA=1` autoriza o fixture `assets/texts_exemplo_
  julho2026.py` (com WARNING). Nunca ativado automaticamente; nunca em execução real.

### Scripts auxiliares (RM16, 16/11/2026 — DEC-058…066)
| Script | Papel |
|---|---|
| `fontes_externas.py` | importa e preserva insumos de fora do SharePoint (cópia imutável, SHA-256, proveniência) — DEC-059 |
| `prestacao_contas.py` | eventos iterativos de Prestação de Contas do mês, já redigidos — DEC-060 |
| `timeline_prestacao_excel.py` | grava esses eventos no Controle Timeline (Excel COM, idempotente) — DEC-060 |
| `timeline_slide_mes.py` | cria o slide do mês no PPTX-base a partir dos marcadores reais — DEC-040/060 |
| `medicao_pdf.py` | extrai o quadro "Nª MEDIÇÃO" de PDF vetorial do Parceiro — DEC-062 |
| `medicao_reconstrucao.py` | reconstrói a Planilha de Medição do quadro impresso e confere; compara com o oficial — DEC-062 |
| `capa.py` | lê/corrige a data da capa embutida (fim da competência) — DEC-058 |
| `20_zip_anexos_g1.py` | anexos .xlsx de P1.1/P1.2 em ZIP individual (POP) — DEC-066 |

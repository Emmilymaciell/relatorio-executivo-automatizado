# Decisões e Erros Conhecidos — Relatório Grupo 1 Programa

> **LEIA ANTES de gerar prints, textos e de montar o Word.** Cada item registra um erro
> cometido em versões anteriores e a **regra vigente** que deve prevalecer. Regra mais
> recente do usuário sempre vence. Não repita os comportamentos "incorretos".
> (Baseline consolidado a partir da produção do RM13/Julho-2026, versões v1→v5.)

## DEC-001 — Formato de data nas planilhas
**Incorreto:** aplicar formato `dd/mm/yyyy` nas células de data; e gravar datas com hora
(datetime 00:00:00).
**Por quê:** no Excel PT-BR o código de ano é `aaaa`; `yyyy` renderiza o literal "yyyy"
(ex.: "31/07/yyyy"). E a query do Power BI faz `type text`→`type date`: com hora ou com o
literal quebrado, TODAS as linhas datadas dão erro ("258 erros" na atualização).
**Regra vigente:** usar formato **`dd/mm/aaaa`** (ou herdar o formato das linhas existentes –
formato de data curto/local) e gravar **datas sem hora**. Validar renderização ("31/07/2026").
**Evidência:** usuária apontou "31/07/yyyy" e os 258 erros do Power BI; correção confirmada.

## DEC-002 — Atualização do acumulado (Controle de demandas)
**Incorreto:** IDs iniciando em valor que **colide** com IDs existentes; risco de reordenar/
alterar linhas antigas.
**Regra vigente:** **só ACRESCENTAR** as linhas da nova versão (`Versão Relatório = N`) ao final;
**nunca** alterar linhas antigas (restaurar do backup e reanexar, se preciso); IDs **contínuos**
como **VALOR** (`base_id + i + 1`), **não fórmula** (DEC-039); **backup** antes de qualquer alteração.
**Evidência:** usuária: "era pra vc alterar só adicionando mais linhas... não ajustando os antigos".

## DEC-003 — Data do documento e cidade
**Sequência A→C:** (A) documento sem data; (B) usei "10 de agosto"; (C) a **capa gráfica** da CG
já trazia **"05 de agosto de 2026"** → alinhei tudo a 05/08.
**SUPERADA pela DEC-058 (16/11/2026): data da capa/encerramento = último dia do mês da competência.**
**Regra anterior:** `data_documento` = a data impressa na **capa da CG** (fonte preferencial);
usar a MESMA data no encerramento; cidade **"Cidade,"** (não "São Paulo,"); a data aparece
**uma única vez**. Se a capa não trouxer data, perguntar ao usuário.
**Evidência:** capa v da CG com 05/08; usuária: "pode deixar 05/08".

## DEC-004 — Capa
**Incorreto:** criar uma capa de TEXTO no início (duplicando a capa).
**Por quê:** a CG já entrega a **capa gráfica** (imagem de página inteira: logos, "N° 13",
período, data). python-docx não "vê" o texto dela (é imagem) — não concluir que não há capa.
**Regra vigente:** **usar a capa gráfica original da CG**; não inserir capa própria.
**Evidência:** página 2 revelou a capa gráfica; a capa de texto foi removida.

## DEC-005 — Encerramento (negrito)
**Incorreto:** encerramento praticamente todo em negrito.
**Regra vigente:** negrito **seletivo**, apenas em: **"Relatório Mensal N° {n}"**, **"{mês} de {ano}"**
e **"{período}"** (01 de … até 31 de …). Restante em peso normal. Corrigir a falha gramatical do
modelo ("durante o mês de referente ao período") — **não** perpetuar erros do RM anterior.
**Evidência:** usuária, ponto 8 da rodada v4.

## DEC-006 — Assinaturas
**Incorreto:** manter blocos genéricos "Nome – Programa"; espaçamento grande entre linha e nome.
**Regra vigente:** **dois blocos lado a lado** (tabela 1×2), no padrão do RM11:
- Esquerda: **Pessoa B** — Programa – Consórcio do Programa — Produto **P1.3 e P1.4** – Relatório mensal – Grupo 1
- Direita: **Pessoa C** — Programa … — Produto **P1.1 e P1.2** – Relatório mensal – Grupo 1
Nome/cargo **imediatamente abaixo** da linha de assinatura. **Remover** qualquer bloco "Nome – Programa".
**Evidência:** RM11 (Word) e correções v3/v4.

## DEC-007 — Anexos e sumário
**Incorreto:** "ANEXOS" numerado no sumário ("3 ANEXOS"); esquecer o Quadro de Anexos.
**Regra vigente:** "ENCERRAMENTO" e "ANEXOS" como títulos **sem numeração** (Heading com numeração
suprimida). Incluir o **Quadro de Anexos** = tabela "LISTA DE DOCUMENTOS ANEXOS" (REF. / DESCRIÇÃO)
com os anexos reais do mês (padrão do RM11). Ver `anexos.md`.
**Evidência:** pontos 1 e 28 da rodada v4.

## DEC-008 — Textos fiéis (não resumir) + títulos de campo
**Incorreto:** textos curtos/reinterpretados; títulos de campo em **negrito**.
**Regra vigente:** reproduzir os parágrafos do **RM11/RM12** com a MESMA estrutura, objetivo,
semântica, linguagem e **tamanho**; alterar **apenas os dados** do mês. Títulos de campo
(Controle de Produtos, Dashboard, Timeline, Reuniões, Análise da medição, Integração
Institucional) em estilo **lista/bullet, SEM negrito, com uma linha em branco acima**.
**Evidência:** pontos 2, 3, 8, 19, 31 das rodadas v3/v4.

## DEC-009 — Print do Controle de Produtos
**Incorreto:** com coluna ID; retrato/estreito; linhas de altura fixa (texto cortado); última
página quase vazia (2 linhas + branco); escalas diferentes entre páginas.
**Regra vigente:** **paisagem**; **SEM coluna ID** (nem "Versão"); colunas Produto/Empresa/
Documento/Tipo/Descrição/Empreendimento/Data Entrega/Status (Descrição a mais larga); **altura
mínima por linha** (compacta 1-linha, expande só o necessário — **nunca cortar**); **paginação
balanceada** (distribuir as linhas ~igualmente; última página cheia); largura ~ da linha do
cabeçalho; alta resolução (~220 dpi). Ver `regras-excel.md`/`regras-layout.md`.
**Evidência:** pontos 3, 5, 6 das rodadas v3/v4 (ex.: página 20 quase vazia; PA-637.1 alto demais).

## DEC-010 — Print das Reuniões
**Incorreto:** paisagem; coluna Anexos ocupando espaço; poucas reuniões/página; linhas ora
sobrepostas (texto cortado).
**Regra vigente:** **retrato**; **ocultar a coluna "Anexos"** (não há anexos nestas reuniões);
alargar **Assuntos** (e Participantes); **Participantes centralizado verticalmente**; **logo do
Programa reposicionada** ao fim da área de Assuntos; altura mínima por linha **sem sobrepor**; mais
reuniões por página (aproveitar até perto do rodapé); largura ~ da linha do cabeçalho.
**Evidência:** pontos 4, 6, 13, 21 das rodadas v3/v4.

## DEC-011 — Separação das reuniões P1.1 × P1.2
**Contexto:** as abas nomeadas "Reuniões de GOV 1-1"/"Stakehold 1-2" não estavam atualizadas p/ o mês.
**Regra vigente:** **"Controle Geral"** (reuniões de governança/coordenação) → seção **P1.1
"Reuniões de Governança"**; **"Reuniões Stakehold 1-2"** → seção **P1.2 (Integração
Institucional)**. Nunca misturar as bases. Documentar a classificação.
**Evidência:** análise das abas do Controle de Reuniões (RM13).

## DEC-012 — Dashboard
**Incorreto:** texto → imagem → resto do texto (imagem no meio); imagem pequena com muito branco.
**Regra vigente:** **todo o texto primeiro**, depois a imagem; imagem **sozinha na sua página**,
**recortada** (sem bordas/branco inferior), **grande** e centralizada, sem invadir rodapé. O PDF do
Power BI é **exportado pelo usuário** (não automatizável headless). Ver `dashboard.md`.
**Evidência:** ponto 4 (v4): "texto → espaço vazio → imagem → resto".

## DEC-013 — Timeline
**Incorreto:** recriar bolinhas/setas "por aproximação"; setas cruzando texto/datas; título
"Timeline" quebrando linha ou colado à faixa verde; imagem invadindo rodapé.
**Regra vigente:** **usar o PPT já corrigido** que está no corpo (o usuário pode enviar a versão
do mês); **copiar os elementos REAIS** do PPT (grupos de marcador = bolinha + seta + caixas de
texto) e duplicá-los; **9 marcos** de julho como exemplo (ajustar ao mês); **4 bolinhas verde-claras
(#00BF63) + demais verde-escuras (#1C5560)**, alternando; **setas reais** (bloco), curtas, que
**não cruzam o texto**; datas do lado oposto da seta; **título "Timeline" em 1 linha, abaixo da
faixa verde**; **uma página horizontal por mês** (multi-página); **não invadir o rodapé**. Ver `timeline.md`.
**FUTURO (pedido p/ próximas versões, ainda não aplicado):** quando **dois eventos têm a mesma
data**, compartilhar a **mesma bolinha** (um p/ cima, outro p/ baixo); a seta do de cima deve
"pular" a área do número da data sem sobrepô-lo.
**Evidência:** pontos 5, 11, 12 (v4) e mensagens de correção da timeline.

## DEC-014 — Medição / Quadro Financeiro
**Regra vigente:** medição = **12ª medição, referência Julho/2026, status "Em análise"** (RM13 —
recalcular a cada mês a partir do Quadro Financeiro). Texto **fiel ao RM11** (introdução idêntica),
com **negrito seletivo** (número da medição, valores R$, %). **Preservar as cores/identidade** do
Quadro Financeiro nos prints (não recriar com paleta diferente). **"Linha do Tempo de Medições"**
dividida em **3 quadros: medições 1-5, 6-10, 11-12**, cada um **ampliado à largura útil**, com as
**mesmas cores** do quadro (Pago=verde, Processado=amarelo, Em análise=laranja). Ver `medicao.md`.
**Evidência:** pontos 5, 6, 7 (medição) das rodadas v3/v4.

## DEC-015 — Cabeçalho/rodapé das páginas em paisagem
**Regra vigente:** páginas horizontais recebem cabeçalho/rodapé **adaptados à largura maior**
(linha full-width, "Programa | CONSÓRCIO…" à esquerda, "-CONFIDENCIAL-" + nº de página no rodapé),
via seções próprias — **sem duplicar** e **sem alterar** o cabeçalho/rodapé das páginas verticais.
**Evidência:** ponto 2 (v4); erro intermediário de cabeçalho duplicado (corrigido limpando o `<w:hdr>`).

## DEC-016 — Dimensionamento das imagens
**Regra vigente:** para TODA imagem/print: (1) recortar bordas brancas; (2) preservar proporção;
(3) dimensionar ao **máximo dentro da área útil** (entre a linha do cabeçalho e a do rodapé),
**sem invadir** cabeçalho/rodapé; (4) centralizar horizontalmente; (5) manter o texto próximo da
imagem (sem grandes vazios). Ver `regras-layout.md` (função de fit).
**Evidência:** ponto 1 (v4): imagens invadindo rodapé e com muito branco.

## DEC-017 — Fontes Canva (capa/timeline)
**Contexto:** a capa e a timeline usam fontes Canva **Garet** e **Open Sauce**, não instaladas na
máquina; Office as substitui por fontes próximas.
**Regra vigente:** aceitável para revisão; **conferência visual** recomendada. Não é erro de
processo. (Se necessário, instalar as fontes para fidelidade total.)

## DEC-018 — Fonte da medição quando o Parceiro não está disponível
**Contexto:** a pasta `Medicao Parceiro` estava **vazia** localmente.
**Regra vigente:** usar o **Quadro Financeiro do mês** fornecido pela equipe como **fonte primária**
da medição; conferir matematicamente (soma por empresa = total; saldo = contrato − acumulado;
% = acumulado/contrato). Se o Parceiro estiver disponível, usar como validação cruzada; em divergência,
**registrar no QA e avisar** (não escolher em silêncio).

## DEC-019 — Proibição de fallback factual de competência em execução real
**Comportamento anterior:** `prog_config.dados()` resolvia o conteúdo do mês na ordem
`PROG_DADOS` → `config/dados_mes.py` → **`assets/texts_exemplo_julho2026.py`** (EXEMPLO). Ou seja,
faltando `config/dados_mes.py`, o processo caía **automaticamente** no exemplo de Julho/2026.
**Risco:** um relatório de agosto, setembro ou qualquer mês poderia usar **silenciosamente** dados
factuais do RM13 — `MARCOS`, `MEDICOES`, textos, `QA_ESPERADOS`, valores R$/% de julho — contaminando
a competência atual. Isso viola "nunca inventar dados".
**Regra vigente:** distinção explícita entre **MODO REAL** e **MODO TESTE/FIXTURE**.
- **MODO REAL** (padrão): `dados()` usa **apenas** `PROG_DADOS` (caminho `.py` explícito) **ou**
  `config/dados_mes.py`. Se **nenhum** existir, a execução **PARA** com `RuntimeError`:
  `DADOS MENSAIS NÃO CONFIGURADOS` (config/dados_mes.py não encontrado; execução real interrompida
  para impedir reutilização de dados factuais de outra competência). **Nunca** usa o exemplo.
- **MODO TESTE/FIXTURE**: o exemplo `assets/texts_exemplo_julho2026.py` só é usado quando o operador
  **explicitamente** define `PROG_USE_EXAMPLE_DATA=1` — e o script imprime um WARNING de que são dados
  de Julho/2026. **Nunca** é ativado automaticamente.
- **`dados_mes.py` não é um relatório manual** (ver DEC-019b abaixo). Guarda só o que NÃO se deriva com
  segurança das fontes; os textos analíticos de P1.1/P1.2 e os `MARCOS` deriváveis vêm das fontes reais.
**Onde:** `scripts/prog_config.py` (`dados()`, `use_example_data()`, `DADOS_NAO_CONFIGURADOS_MSG`);
documentado em `SKILL.md`, `scripts/README.md`, `processo-completo.md`, `fontes-sharepoint.md`, `qa.md`.
**Evidência:** decisão do usuário ao promover a arquitetura à v1.0 (proteção `dados_mes`).

### DEC-019b — `dados_mes.py` não deve virar um relatório manual
**Regra vigente:** `config/dados_mes.py` **não** deve exigir digitar todo mês os textos de P1.1/P1.2.
A automação **produz os textos analíticos a partir das fontes reais** (Controle de Demandas, Dashboard,
Controle de Reuniões, Quadro Financeiro, planilhas e documentos mensais). `dados_mes.py` guarda
**apenas** o que não é derivável com segurança, por exemplo: `MARCOS_EXTRAS`, `STATUS_MEDICAO` (quando
necessário), `OBSERVACOES_EXTRAORDINARIAS`, `INSTRUCOES_ESPECIFICAS_MES`, `AJUSTES_MANUAIS_AUTORIZADOS`
e eventuais parâmetros de QA. Quando os `MARCOS` puderem ser extraídos de uma base estruturada oficial,
**prefira a base**; não duplicar manualmente dado que pode ser calculado das fontes.

## DEC-020 — Timeline: fallback do slide exige conferência visual (QA)
**Contexto:** o script `07_timeline.py` procura a slide-alvo pelo placeholder **`XXXX`**. Se não achar,
usa o **último slide** como fallback — comportamento útil, que **não** deve ser removido.
**Risco:** o fallback pode escrever os marcos na slide errada sem que ninguém perceba (aprovação
silenciosa de uma timeline potencialmente incorreta).
**Regra vigente:** manter o fallback, porém **não aprovar automaticamente** a etapa quando ele ocorrer:
1. gerar a Timeline normalmente;
2. registrar **WARNING** no log (`cfg.log`);
3. gravar um aviso de QA persistente (`cfg.set_qa_warning("timeline_fallback_xxxx", …)`);
4. o `13_qa.py` lê `cfg.qa_warnings()` e marca **"CONFERÊNCIA VISUAL DA TIMELINE NECESSÁRIA"** como
   pendência → veredito **REVISÃO NECESSÁRIA** (não APROVADO) até a conferência visual.
O aviso é **idempotente**: se numa reexecução o placeholder `XXXX` for encontrado, o aviso é limpo.
**Onde:** `scripts/07_timeline.py`, `scripts/13_qa.py`, `scripts/prog_config.py`
(`set_qa_warning`/`qa_warnings`); documentado em `timeline.md` e `qa.md`.
**Evidência:** decisão do usuário ao promover a arquitetura à v1.0 (segurança do fallback da timeline).

## DEC-021 — Proteção DRY_RUN na publicação final
**Comportamento anterior:** `12_finalizar_pdf.py` sempre copiava o DOCX+PDF finais para a pasta `G1`
do mês (SharePoint), **mesmo** com `PROG_DRY_RUN=1` — o único script de escrita sem guarda de DRY_RUN
(apontado no teste de regressão RM13).
**Risco:** uma execução de teste/regressão publicaria arquivos na pasta oficial do mês no SharePoint.
**Regra vigente:**
- **`PROG_DRY_RUN=1`**: gera DOCX/PDF **somente localmente** (`07_SAIDA`); **não** copia/sobrescreve
  nada no `G1`; registra no log **"DRY_RUN: publicação no SharePoint ignorada"**.
- **Execução real**: mantém a publicação no `G1`, **nunca sobrescrevendo** — usa versionamento
  `_v2/_v3…` (`uniq`), como já era a regra.
A publicação foi isolada no helper `_publicar_g1(g1, docx_out, pdf_out, base)` (testável sem Office COM).
**Onde:** `scripts/12_finalizar_pdf.py`; `qa.md`, `scripts/README.md`, `SKILL.md`.
**Evidência:** teste de regressão RM13 (script 12 era o único a escrever no SharePoint sem guarda).

## DEC-022 — Ausência do PPTX-fonte da Timeline
**Comportamento anterior:** sem o `.pptx` da Timeline, o `07` caía num caminho `Timeline_RASCUNHO.pptx`
e encerrava com **erro genérico** (`sys.exit(1)`), sem orientação nem registro de QA.
**Regra vigente:** se o `.pptx`-fonte não for localizado (em `G1/Anexos/P1.1` nem em `01_FONTES`):
1. **NÃO** inventar/reconstruir a Timeline; **NÃO** usar o PDF anterior como fonte editável;
2. interromper **somente** a etapa da Timeline (as demais seguem — são processos separados);
3. emitir a mensagem clara **"TIMELINE PPTX NÃO LOCALIZADA"** e orientar como fornecer o arquivo;
4. registrar pendência de QA (`cfg.set_qa_warning("timeline_pptx_ausente", …)`), que some quando o
   `.pptx` é fornecido e o `07` reexecuta (idempotente).
**Preserva a DEC-020**: quando o `.pptx` EXISTE mas o placeholder `XXXX` não é achado, mantém o
fallback do último slide + a pendência de conferência visual. São situações distintas.
**Onde:** `scripts/07_timeline.py`; `timeline.md`, `qa.md`.
**Evidência:** teste de regressão RM13 (Timeline só existia como PDF nas fontes oficiais).

## DEC-023 — Separação entre projeto permanente e diretório de execução
**Contexto:** para isolar as saídas do teste de regressão foi preciso apontar `PROG_PROJECT_DIR` para
uma pasta de teste, o que também desloca a busca de `config/`. Faltava um jeito de redirecionar
**só** a saída/execução.
**Regra vigente:** nova variável **`PROG_WORK_DIR`**:
- **`PROG_PROJECT_DIR`** = raiz **permanente** do projeto (config/, Skill, etc.).
- **`PROG_WORK_DIR`** = raiz da **execução atual** (todas as subpastas `01_FONTES..07_SAIDA`,
  imagens, logs, backups e QA-flags vão para lá).
- Se `PROG_WORK_DIR` não for informado, mantém-se o comportamento padrão
  (`project_dir()/output/RM{n}_{ano}_{mm}`). Nenhum caminho permanente da Skill/config é quebrado.
Como todos os scripts usam `cfg.work_dir()`, a mudança é central (não exige edição script a script).
**Onde:** `scripts/prog_config.py` (`work_dir`); `scripts/README.md`, `SKILL.md`.
**Evidência:** teste de regressão RM13 (isolamento de saída sem alterar a raiz do projeto).

## DEC-024 — Dashboard "página cheia" (imagem grande, sem página em branco)
**Incorreto:** a imagem do Dashboard era dimensionada pela área retrato padrão (≈6,8 × 8,5in),
ficando **pequena** — ~2,1in de espaço em branco entre a base da imagem e o risco do rodapé, além de
recuos laterais grandes. Escala/aproveitamento inadequados vs. o padrão visual desejado.
**Regra vigente:** o Dashboard usa a área **"página cheia"** `area=(PW, PH_FULL)` com `PH_FULL≈10,1in`
(body útil retrato ≈10,3in): a imagem ocupa quase toda a área útil entre os riscos — **base bem
próxima do risco do rodapé** (~0,5in), **recuos laterais mínimos**, **proporção preservada** (sem
distorção). Para não invadir o rodapé nem estourar para outra página, `PH_FULL` fica abaixo do body.
**Sem página em branco:** a imagem quase preenche a página, então **não** se insere um parágrafo de
quebra vazio depois dela (ele criaria uma página em branco). Em vez disso, a seção seguinte (título da
Timeline) começa em página nova via **`page_break_before=True`**.
**Onde:** `scripts/09_montar_word.py` (`PH_FULL`, `fit_img(..., area=...)`, `page_break_before` no
título da Timeline); `references/regras-layout.md`, `dashboard.md`, `qa.md`.
**Evidência:** revisão do usuário sobre o RM13 reproduzido (Dashboard pequeno com muito branco);
medição pós-ajuste: altura 8,5→10,1in, gap inferior 2,12→0,52in, 44 páginas (sem página em branco).

## DEC-025 — Timeline: editável acumulado compartilhado + técnica de montagem
**Fonte (editável ACUMULADO):** a Timeline é um **PPT compartilhado** com **uma slide por mês**
(Jul/2025 → mês corrente), mantido em `…\G1 - Coordenacao\Desenvolvimento de Produtos\`
(ex.: `Compartilhado Timeline Relatório Programa.pptx`). É a **REFERÊNCIA oficial**, atualizada mês a mês.
`prog_config.timeline_pptx_src()` resolve: **env `PROG_TIMELINE_PPTX`** → pasta *Desenvolvimento de
Produtos* (compartilhado) → anexo P1.1 do mês → `01_FONTES`. (Antes o `07` só olhava no anexo do mês,
por isso a Timeline não entrava no relatório.) Se nada existir, pendência de QA (DEC-022).
**Slide-modelo:** a cada mês há uma slide com os elementos-modelo e placeholders **`DD/MM`/`XXXXXX`**;
o `07` **duplica esses elementos reais** (ponto, seta, caixas), **nunca cria do zero** (DEC-013).
**Fluxo mensal:** `07` preenche a slide do mês (a partir do modelo) → **sobrescreve** o editável
compartilhado (com backup) → exporta o **anexo PDF** (`P1.1 - Timeline de Acompanhamento Macro até
{mês} {ano}.pdf`) → usa a imagem no relatório. O editável pode ser **refinado manualmente** (é a
referência); o `07` gera um **ponto de partida** fiel à técnica abaixo.
**Técnica de montagem (aprendida do ajuste manual do usuário):**
1. **Datas iguais = MESMO ponto** (1 bolinha) com **2 setas** (uma acima, uma abaixo). Ver `MARCOS`
   com a mesma data em sequência (um `above=True`, outro `False`).
2. **Data única** por ponto: no 2º marco da mesma data, a data é apagada (não duplicar).
3. A **seta do lado da data** é **curta e começa depois da data** (não sobrepõe o número); a do outro
   lado é normal, ligando o ponto ao texto.
4. **Descrições escalonadas**: alterna nível **perto/longe** da linha entre descrições do mesmo lado
   (permite ler sem colidir). Marcadores **sob o título "Timeline"** (centro) usam o nível baixo.
5. **Setas de comprimento variável**: cada seta vai do ponto até **logo abaixo/acima** do seu texto
   (sem seta sobre texto — DEC-013).
6. **Largura da descrição ≤ espaçamento** entre pontos (não sobrepõe o vizinho); cores **verde-claro
   `#00BF63` / verde-escuro `#1C5560` alternadas** por ponto.
**Onde:** `scripts/07_timeline.py` (merge por data, escalonamento, setas variáveis, data única),
`scripts/prog_config.py` (`timeline_pptx_src`, env `PROG_TIMELINE_PPTX`); `references/timeline.md`,
`SKILL.md`, `scripts/README.md`. **Evidência:** o usuário ajustou manualmente o slide de julho; a
medição do slide ajustado guiou a codificação (níveis, comprimentos de seta, data única).

## DEC-026 — Texto-base do P1.1 "Reuniões de Governança" (padrão fixo de ~7 parágrafos)
**Incorreto:** texto de Reuniões de Governança **curto** (ex.: 3 parágrafos genéricos), pouco descritivo.
**Regra vigente:** o P1.1 Reuniões de Governança segue SEMPRE um **texto-base de ≈7 parágrafos densos**:
(1) **abertura** (mês + focos + Coordenação Executiva como instância integradora entre Parceiro/Órgão Contratante/
Órgão Supervisor/executoras/certificadora/jurídico); (2–6) **parágrafos por TEMA/AGENTE** presentes no mês
(governança interna/RACI; institucional Parceiro-Órgão Contratante; FUP/técnico-operacional por empreendimento
RD-101/102A/103/104; certificação/conformidade – Certificadora; governança financeira – CG; jurídico-
institucional – Escritório Jurídico/Procuradoria/TAG/TCE); (7) **síntese** ("Em síntese, {mês} foi marcado por…").
A cada mês **mantêm-se estrutura, semântica e tamanho**, moldando apenas os **dados do mês** (assuntos
reais da aba de governança). Se um agente não teve reunião, substitui-se aquele parágrafo pelo tema que
dominou (ex.: julho = FUPs conjuntos no lugar de reuniões separadas por empresa). O **P1.2** mantém o
mesmo espírito, com **parágrafo de consolidação/síntese** ao final.
**Onde:** `references/regras-redacao.md` (o molde completo); exemplo vigente em
`assets/texts_exemplo_julho2026.py` (`REUNIOES`); modelos históricos: RM abril/maio/junho.
**Evidência:** o usuário forneceu 3 meses (abr/mai/jun) como padrão e aprovou o texto de julho moldado.

## DEC-027 — Tabela de Reuniões: escala 1:1 com o corpo, calibrada e verificada no PDF
**Incorreto (v1.4 e anteriores):** a tabela era um print "ajustado no olho": `FitToPagesWide=1` comprimia
a planilha, o Word ainda a reduzia por altura, e o resultado saía com **conteúdo em ~6,5 pt** (contra
11 pt do texto corrido) e **cabeçalho em ~12 pt** — desproporcional; **largura 6,755"** (desalinhada do
texto justificado de 6,891"); **logo do Programa ausente** no print; alturas por fórmula fixa
(`RowHeight=L*13+7`) deixando grandes vazios e, ainda assim, **cortando a última linha** de algumas
células; Participantes centralizado apenas na vertical.
**Regra vigente:** a imagem é **consequência de uma cópia do Excel corretamente formatada**, nunca de
correção posterior no PNG. Ordem obrigatória: dados → cópia → formatação → dimensões → quebras →
área/escala de impressão → renderização → validação → inserção.
1. **Escala 1:1** — margens de impressão calculadas para que a largura natural de B+C+D seja a largura
   útil do corpo (6,891" retrato / 10,4" paisagem), `Zoom=100` (**nunca** FitToPages), imagem inserida
   nessa mesma largura. Assim **pt no Excel = pt no PDF** e as bordas coincidem com o início/fim do
   texto justificado.
2. **Calibração impresso × reportado** — o Excel reporta `Width/Height` em pt que **não** são os pt
   impressos (medido nesta máquina: largura ×1,0225, altura ×0,9357). Sem isso a tabela sai 2,4% mais
   larga (e o Word a reduz, encolhendo a letra) e a paginação superestima as alturas em 6,5%, gerando
   branco. Os fatores são medidos num PDF de calibração; **formas (logo) continuam em pt reportados** —
   usar os impressos empurra a logo fora da área de impressão e ela sai cortada.
3. **Fontes pelo alvo visual** — conteúdo 9,5 pt, cabeçalho 11 pt, título da faixa 13 pt.
4. **Negrito reaplicado do zero** — órgão + `:` em Participantes; título de cada bloco em Assuntos.
5. **Alinhamentos** — Data e **Participantes centralizados (H+V)**; Assuntos topo/esquerda; wrap.
6. **Alturas mínimas verificadas no PDF** — `AutoFit` (coluna 1,5% mais estreita) + folga 1 pt e depois
   o ciclo medir-corte → aumentar o mínimo → repaginar (até 3×). Duas checagens: última linha de texto
   contra a borda **e** ausência de termos no PDF (quando falta mais de uma linha o Excel não desenha
   as seguintes, então o corte é invisível para a checagem geométrica).
7. **Partes com identidade completa** — cada parte é exportada de uma cópia da aba com só as suas
   linhas + 1-3; faixa, **logo** e cabeçalho repetem de verdade (`PrintTitleRows` não repete formas).
8. **Aproveitamento vertical** — a 1ª parte usa o espaço que sobra abaixo do texto do P1.1 **medido no
   DOCX montado** (2ª passagem 04→09; a estimativa em documento isolado erra ~0,30" e custou uma
   reunião inteira de página). Faixa 24 pt e espaçador 3 pt: cada ponto se repete em todas as partes.
9. **Retrato × paisagem automático** — as duas orientações são formatadas, paginadas e renderizadas; a
   de menor **desperdício** vence. Desperdício = espaço perdido na página do texto + sobras das partes;
   em paisagem contam **todas** as sobras e a página do texto inteira (mudar de orientação obriga página
   nova). Em julho: retrato 3,99" × paisagem 15,71" → retrato, mesmo a tabela ficando 37% mais baixa em
   paisagem (a altura útil cai para 5,9" e continuam 3 páginas).
10. **`08` não recorta pastas com `.precropped`** — o 04 já recorta exatamente na tinta; recortar de
   novo com `pad` acrescenta borda branca e desalinha a tabela do corpo.
**Instância dedicada do Office** — `DispatchEx` para o Word: não se liga ao Word do usuário nem herda
instância travada ("servidor RPC não está disponível"); fechar/encerrar nunca descarta uma medida.
**Onde:** `scripts/04_print_reunioes.py` (reescrito), `scripts/09_montar_word.py`
(`tabela_reunioes`, largura exata + seção paisagem quando escolhida), `scripts/08_recortar_imagens.py`
(`.precropped`), `scripts/prog_config.py` (`PROG_BASE_DOCX`), `assets/logo_prog.png`,
`references/reunioes.md`, `references/regras-excel.md`, `references/qa.md`, `SKILL.md` §4/§4b.
Substitui a regra de altura da **DEC-010**; mantém a resolução de aba da **DEC-011**.
**Evidência:** RM13/julho regenerado (57 páginas) — pág. 36 com 3 reuniões (89% de ocupação), 37 com 2
(69%), 38 com 1 seguida do texto da medição; tabela 6,891" alinhada ao texto; logo em todas as partes;
nenhum corte. Limite conhecido: sobra de 3,17" na pág. 37 porque a reunião de 31/07 mede 3,94" —
zerar exigiria dividir uma reunião entre páginas ou tabela nativa do Word (não implementados).

## DEC-028 — Timeline no corpo: ACUMULADA COMPLETA (do 1º mês ao mês corrente), não só a slide do mês
**Incorreto (até v1.5):** o `07_timeline.py` renderizava **apenas a slide do mês** para o corpo
(`pg=doc[tgt_index-1]` → um único `timeline_p01.png`). O relatório saía com **uma** página de Timeline
(só o mês corrente), divergindo do baseline (que traz a linha do tempo inteira) e encurtando o
documento (RM13 caiu de ~58 para 46 páginas). Note que a própria `references/timeline.md` já
especificava "uma página por mês, do início até o mês" — era o **script** que estava fora da regra.
**Regra vigente:** a Timeline entra no corpo **COMPLETA/ACUMULADA** — **uma página horizontal por
mês, do 1º mês (Jul/2025) até o mês corrente**, em ordem cronológica (a última página é o mês do
relatório). O `07` exporta o PPT inteiro para PDF e **renderiza TODAS as slides** em
`timeline_p01..pNN.png` (200 dpi); o `09` (que já faz `glob('timeline_p*.png')` ordenado) embute
todas. **Nunca** renderizar só a slide do mês.
**Onde:** `scripts/07_timeline.py` (bloco de render: laço sobre todas as páginas do PDF, em vez de
`doc[tgt_index-1]`; docstring atualizada), `references/timeline.md` (passo 7 e "No corpo do
relatório"). Não altera a técnica de montagem da slide do mês (DEC-013/DEC-020/DEC-022/DEC-025).
**Evidência:** RM13/julho (teste local, DRY_RUN) regenerado com as 13 páginas da Timeline
(Jul/2025→Jul/2026) → **58 páginas** (era 46). Aprovado pela usuária em 19/08/2026.

## DEC-029 — Quadro Financeiro Consolidado: print até o fim do RESUMO POR EMPRESA + largura = corpo
**Incorreto (até v1.6):** o `05` usava `PrintArea = ws.UsedRange` para a aba `PAINEL_FINANCEIRO_*`.
O UsedRange ia até a última linha usada (ex.: `$A$1:$R$91`), **incluindo uma tabela auxiliar/rascunho**
(linhas ~83-87, que alimenta gráficos) **abaixo** da tabela "RESUMO FINANCEIRO POR EMPRESA". Esse
rascunho vazava no print da 1ª imagem da seção de medição. Além disso, o `09` inseria a imagem com o
fit padrão (`min(PW, PH*aspect)`), que — por a imagem ser alta — resultava numa **largura menor que o
corpo** do texto, desalinhada do início/fim do parágrafo justificado. Obs.: `references/medicao.md`
já pedia "recortar até a linha do TOTAL do resumo"; era o **script** que estava fora da regra.
**Regra vigente:**
1. **Área de impressão por conteúdo** — o `05` acha o título **"RESUMO FINANCEIRO POR EMPRESA"** e anda
   para baixo enquanto a tabela é contígua (coluna "Empresa"/"TOTAL" preenchida), parando na 1ª linha
   vazia; recorta o `PrintArea` na linha do **TOTAL**. Exclui o rascunho. Sem linha fixa (muda por mês).
2. **Largura = corpo justificado** — o `09` insere o `quadro_consolidado` com **largura EXATA do corpo**
   (retrato: `page_w − margens laterais`, medido no próprio DOCX ≈ **6,891"**), casando início/fim com o
   texto. Só reduz se, nessa largura, a altura passar da área útil (não estoura a página). Vale só para
   esta imagem; as demais seguem o fit padrão.
**Onde:** `scripts/05_imagens_financeiro.py` (detecção do fim do RESUMO → PrintArea recortado),
`scripts/09_montar_word.py` (`BODY_W` medido nas seções + `fit_img_bodywidth` para o quadro
consolidado), `references/medicao.md` (bloco "Imagens (script 05)").
**Evidência:** RM13/julho (teste local, DRY_RUN) — pág. 39: print termina no TOTAL do RESUMO POR
EMPRESA (sem rascunho) e a imagem mede **6,891"** = largura do corpo (22 imagens confirmadas nessa
largura). Documento 58 → **57 páginas**. Aprovado pela usuária em 19/08/2026.

## DEC-030 — Linha do Tempo de Medições: lida do Excel (STATUS_MEDICOES), 5+5+…, com legendas
**Incorreto (até v1.7):** o `06` montava a "Linha do Tempo de Medições — Status Financeiro" a partir de
`cfg.dados()` (MEDICOES/MEDICOES_STATUS em `config/dados_mes.py`) — dado **manual**, que precisava ser
preenchido à mão e, se ausente, sumia do relatório (foi o que aconteceu após reexecutar só o `05`, que
apaga as PNGs do financeiro, sem reexecutar o `06`: nada aparecia após "Por fim, apresenta-se o
acompanhamento financeiro…").
**Regra vigente:** o `06` lê os dados **direto da aba `STATUS_MEDICOES` do Quadro Financeiro** (a mesma
fonte oficial das demais imagens): descobre a linha de cabeçalho ("MEDIÇÃO n", respeitando células
**mescladas**), as colunas de cada medição e as linhas de empresa (até TOTAL). O **valor** é o da 1ª
coluna do grupo (iguais) e o **status** é o de **maior avanço** entre as colunas daquela medição
(**Pago > Processado > Em análise**; medições com histórico de status ocupam várias colunas, ex.: med1
Faturado→Processado→**Pago**, med5 **Pago**/Processado/Em análise). Essa regra reproduz **exatamente**
os 12 status do baseline (valores idênticos aos de `cfg.dados()` na conferência). Render mantém as
**legendas** e as cores (DEC-014) e divide em **blocos de 5** (1-5/6-10/11-12). Fallback: se a aba não
existir, usa `cfg.dados()`. Ordem no `05`/`06`: como o `05` limpa a pasta `03_IMAGENS/FINANCEIRO`, o
`06` **sempre** roda depois do `05`.
**Onde:** `scripts/06_linha_tempo_medicoes.py` (`medicoes_do_excel()` + regra de status; render
inalterado), `references/medicao.md`, `SKILL.md`.
**Evidência:** RM13/julho (teste local) — 3 quadros (1-5/6-10/11-12) com legendas aparecem após "Por
fim…"; valores conferem 100% com o Excel. Aprovado pela usuária em 19/08/2026.

## DEC-031 — Base Técnica Parceiro: dividida em 2 imagens p/ minimizar branco antes da quebra
**Incorreto (até v1.7):** a Base Técnica saía como **uma** imagem (~5,5" de altura). Como o espaço
abaixo do texto de avaliação é ~5,3", ela **estourava por pouco** e era empurrada inteira para a página
seguinte, deixando um **branco grande** (~5") no fim da página do texto.
**Regra vigente:** o `05` **mede a altura cumulativa de cada produto** no próprio print (localiza as
linhas pela coluna ID `d.d`; escala = a mesma da inserção no corpo) e **divide a Base em 2 imagens**: a
**parte 1** recebe o máximo de produtos cuja altura cumulativa ≤ alvo `PROG_BASE_PART1_IN` (default
**5,25"**) — preenchendo o resto da página; a **parte 2** recebe os demais. **Título + cabeçalho de
coluna são repetidos** nas duas (via `PrintTitleRows`; título/cabeçalho localizados **por conteúdo**:
"BASE TÉCNICA" e a linha com "PRODUTO"+"UNIDADE" — robusto ao offset do UsedRange no COM). Se o alvo não
produzir `0<n1<total`, mantém **imagem única**. Tudo em **uma passagem** (o `05` não precisa do doc
montado). O `09` insere `base_parceiro_p1/p2` (glob `base_parceiro*.png`).
**Onde:** `scripts/05_imagens_financeiro.py` (`_base_cumheights()` + divisão), `scripts/09_montar_word.py`
(glob das partes), `references/medicao.md`, `SKILL.md`.
**Evidência:** RM13/julho (teste local) — parte1=19 produtos (~5,20") preenche a pág. 40; parte2=1
produto abre a pág. 41 com título+cabeçalho; sem branco grande. Aprovado pela usuária em 19/08/2026.
Limite: o alvo 5,25" é fixo (não mede o doc montado); se um mês tiver texto de avaliação muito maior/
menor, ajustar `PROG_BASE_PART1_IN`.

## DEC-032 — Layout: cada "Produto 1.X" em página nova + cabeçalho/rodapé em todas as páginas + QA
**Incorreto (até v1.8):**
1. Só o título **"Produto 1.1"** iniciava em página nova (tinha uma quebra manual antes). **1.2, 1.3 e
   1.4** dependiam do fluxo do conteúdo e caíam **no meio da página** (1.2 na pág. 42→43; 1.4 na pág.
   48, colada ao fim do 1.3).
2. A **seção final** do documento (herdada do sectPr do corpo da base CG) tinha
   **`w:titlePg` (primeira página diferente)** ligado; como o cabeçalho/rodapé de "primeira página"
   estava vazio, a **1ª página dessa seção saía sem cabeçalho nem rodapé** (era a **pág. 35**).
3. O QA não verificava nem quebras de página dos títulos nem cabeçalho/rodapé por página.
**Regra vigente:**
1. **Todo** título "Produto 1.X – …" (estilo Heading, texto começa com "Produto 1.") recebe
   **`page_break_before=True`** no `09`; quebras manuais redundantes no parágrafo anterior são removidas
   (evita página em branco dupla).
2. O `11` força **`different_first_page_header_footer=False`** em **todas** as seções → nenhuma página
   fica sem cabeçalho/rodapé. Todas as páginas devem ter cabeçalho ("CONSÓRCIO") e rodapé
   ("CONFIDENCIAL").
3. **Verificação antes de enviar (no `13`):** (a) lista páginas sem cabeçalho/rodapé; (b) confere que
   cada "Produto 1.X – Relatório …" inicia no **topo** de uma página (Y ≤ 30% da altura). Qualquer
   falha vira pendência (**REVISÃO NECESSÁRIA**), impedindo aprovação automática.
**Onde:** `scripts/09_montar_word.py` (page_break_before nos títulos), `scripts/11_cabecalho_rodape_
paisagem.py` (remove primeira-página-diferente de todas as seções), `scripts/13_qa.py` (2 checagens de
layout), `references/qa.md`, `SKILL.md`, `scripts/README.md`.
**Evidência:** RM13/julho (teste local) — 1.1→p5, 1.2→p43, 1.3→p47, 1.4→p50, todos em Y=57pt (topo);
cabeçalho+rodapé em todas as 60 páginas (pág. 35 corrigida); QA acusa OK nas duas checagens novas; sem
página em branco. Aprovado pela usuária em 19/08/2026.

## DEC-033 — P1.2: texto consolidado "Relatório de Relacionamento Institucional" (obrigatório)
**Incorreto (até v1.9):** o P1.2 terminava na **tabela de reuniões 1.2**; faltava o **texto
consolidado** *"Relatório de Relacionamento Institucional do mês de {mês} de {ano}"* que deve vir
**depois** da tabela (o relatório saía sem esse campo).
**Regra vigente:** o P1.2 tem, após a tabela de reuniões, um **texto consolidado obrigatório**
(`cfg.dados().P12_CONSOLIDADO`; título gerado pelo `09`). É **redigido a cada mês** a partir do
**modelo do mês anterior** (estrutura/coesão — molde vigente em `assets/texts_exemplo_julho2026.py`,
`P12_CONSOLIDADO` = julho/2026) **atualizando os dados com a planilha de reuniões 1.2 do mês**
(`P1.2 - Controle de Reuniões`, aba de stakeholders — ler os *assuntos* e refletir temas/
empreendimentos/eventos reais, sem inventar). Estrutura fixa em ~12 parágrafos por tema (definição →
panorama → governança → repositório/sede → engenharia → certificação → não conformidades → medições/
financeiro → suporte/CERT → jurídico → síntese → governança documental). Gravar em `config/dados_mes.py`.
O `13` (QA) **exige** o campo (procura "Relatório de Relacionamento Institucional do mês de"); ausência
= pendência. Ao colar de PDFs, **remover cabeçalho/rodapé** ("Programa | CONSÓRCIO…", "-CONFIDENCIAL- {pág}")
e rejuntar frases partidas.
**Onde:** `assets/texts_exemplo_julho2026.py` (`P12_CONSOLIDADO`), `scripts/09_montar_word.py` (insere
após a tabela 1.2), `scripts/13_qa.py` (checagem de presença), `references/regras-redacao.md`,
`references/estrutura-relatorio.md`, `SKILL.md`.
**Evidência:** RM13/julho (teste local) — texto consolidado na pág. 46, após a tabela de reuniões 1.2,
título "…do mês de julho de 2026", 12 parágrafos limpos (sem cabeçalho/rodapé); QA "presente" OK;
62 páginas. Aprovado pela usuária em 19/08/2026.

## DEC-034 — Medição: fonte oficial = Planilha de Medição (Medicao Parceiro) + painel do mês anterior como template
**Contexto:** na 1ª geração real (RM13/julho, 24/08/2026) eu havia pego o painel `P1.1 - Quadro
Financeiro` do **espelho Parceiro** — errado (a usuária: "é pra você produzir tudo… todo mês do zero").
**Regra vigente:** os **números da medição** vêm SEMPRE de
`…\Desenvolvimento de Produtos\Medicao Parceiro\Planilha de Medição - Estruturadora - {Mês}.xlsx`
(aba `Global` = med 1–6; `SUPRIMIDA` = med 7–12+; bloco da medição N pelo cabeçalho "Nª MEDIÇÃO";
produtos ID col D; empresas linhas 33–36; total 37). O **painel formatado** (`P1.1 - Quadro Financeiro
do mês de {mês}.xlsx`, com gráfico) é **produzido**: copia-se o painel do **mês anterior** (template,
preserva formatação/gráfico) e preenchem-se as ~240 células com os dados da Planilha. **Lógica validada
100%** contra o painel-oráculo de julho: por produto (%acum/valor final/valor acum direto do bloco);
por empresa (valor mês direto; **acum = soma dos produtos da empresa**; %=acum/contrato; saldo=contrato−
acum); totais e gráfico por soma/cópia; "Qntd. Acum." por **recorrência** (mês anterior + incremento).
Editar via **Excel COM** (preserva o gráfico). **Nunca** usar o espelho Parceiro nem inventar valores.
**Onde:** `references/medicao.md`, `references/fontes-sharepoint.md`, `SKILL.md`; script a construir
`scripts/15_gerar_painel_financeiro.py`. **Supera** o legado DEC-018 (pasta Medicao Parceiro "vazia").
**Evidência:** RM13/julho — total da medição, acumulado e valores por empresa (mês e acumulado por
soma-de-produtos) **batem célula-a-célula** com o painel. *(Valores omitidos na versão pública.)*

## DEC-035 — Timeline: fonte dos marcos = Excel "P1.1 -Controle Timeline.xlsx"; gerar o slide dos marcos
**Contexto:** o PPT compartilhado tinha o slide do mês **vazio** e eu rendeziei o PDF do Parceiro — errado.
A usuária criou um **Excel de controle dos marcos** e definiu o processo.
**Regra vigente:** os **marcos** vêm de `…\Desenvolvimento de Produtos\P1.1 -Controle Timeline.xlsx`
(aba `Base Timeline`; competência = 1º dia do mês, formatada "mês/ano"; usar `Incluir? = Sim`). **Todo
mês conferir se os marcos do mês estão lá; se não/incompletos, PERGUNTAR ao usuário, preencher o Excel
e só então gerar.** A partir dos marcos, **gerar o slide mensal** no PPTX-base `P1.1 - Timeline
Relatório Programa.pptx` (a slide do mês já traz os grupos de marcador posicionados — **preencher** cada um
com data+descrição na ordem por data, sem duplicar) e renderizar todas as slides (DEC-028). **Nunca**
usar o PDF do espelho Parceiro como fonte.
**Onde:** `references/timeline.md`, `references/fontes-sharepoint.md`, `SKILL.md`; ajuste do
`scripts/07_timeline.py` (ler marcos do Excel + preencher grupos).
**Evidência:** RM13/julho — Base Timeline linhas 48–56 = os 9 marcos de julho/2026 (09/07 a 30/07),
idênticos aos validados contra a timeline oficial.

## DEC-036 — Reuniões: produzir o anexo mensal do MESTRE (filtrar o mês), não pegar mensal pronto
**Regra vigente:** o anexo `P1.2 - Controle de Reuniões do mês de {mês}.xlsx` é **produzido** filtrando
o mês do **mestre** `…\Desenvolvimento de Produtos\P1.2 - Controle de Reuniões.xlsx` (abas "Reuniões de
GOV 1-1" e "Reuniões Stakehold 1-2" contêm todos os meses). O `04` lê o anexo mensal; então é preciso
gerá-lo antes (filtro por data do mês). **Nunca** copiar um mensal pronto do espelho Parceiro/Obsoleto.
**Onde:** `references/fontes-sharepoint.md`, `references/reunioes.md`, `SKILL.md`; ajuste do fluxo do `04`.

## DEC-037 — Produzir do zero das fontes oficiais; NUNCA usar o espelho Parceiro/Obsoleto; Dashboard = export do usuário
**Regra vigente (a usuária, 24/08/2026):** todo mês o relatório é **produzido do zero** a partir das
**fontes oficiais** (listas de demandas G1/G3/G4/G5 + Lista Mestre G2; Planilha de Medição; Controle
Timeline; mestre de Reuniões; base CG). O **espelho `5. Compartilhado Parceiro`** e a pasta `Obsoleto` são
**cópias de entrega**, **nunca fonte**. Se uma fonte oficial faltar/estiver desatualizada, **perguntar
ao usuário** antes de gerar — nunca substituir por cópia nem inventar. **Dashboard:** exportado pelo
**usuário** do Power BI (checkpoint) — a automação só usa o PDF que ele salvar; nunca publica Power BI.
**Onde:** `references/fontes-sharepoint.md`, `SKILL.md`, `CLAUDE.md`.
**Contexto:** correção após a 1ª geração real do RM13 ter usado o espelho Parceiro para quadro/dashboard/
timeline/reuniões. Os NÚMEROS de julho já publicados estavam corretos (batem com a Planilha oficial);
o que mudou foi a **proveniência/processo** para os próximos meses.

## DEC-038 — Export COM falha com PDF pré-existente ("documento não foi salvo"); e QA "TODO" falso-positivo
**Sintoma:** `ws.ExportAsFixedFormat(0, pdf)` (Excel) e `SaveAs2/ExportAsFixedFormat` (Word/PowerPoint)
falhavam com `com_error … 'O documento não foi salvo.'` (código -2147018887) de forma **intermitente
mas reprodutível**, quando o **PDF de destino já existia** de um run anterior (o Excel/Office não
sobrescreve um PDF pré-existente que considere em uso). Diagnóstico enganoso: parecia "degradação do
COM na sessão", mas o export ISOLADO do mesmo arquivo/abas funcionava 100% — a diferença era só o
`os.remove(pdf)` antes. Custou muito tempo no 1º run real (RM13); o `04` já tinha o remove (por isso
funcionava), os demais não.
**Correção (vigente):** **apagar o PDF de destino antes de exportar** em TODOS os scripts que exportam:
`03` (controle_LS), `05` (`_export_render` fin_*), `06` (linha_med/status), `04` (prévia Word),
`12` (SaveAs2 docx/pdf), `17` (timeline SaveAs). Padrão: `if os.path.exists(pdf): os.remove(pdf)`.
Operacional: manter `04_PDFS_TEMP` limpo entre runs também ajuda.
**QA (`13`) — falso-positivo "TODO":** a checagem de placeholders casava "TODO" com palavras
portuguesas ("todos", "toda", "metodologia") por ser case-insensitive/substring. Agora "TODO" é
verificado **case-sensitive + palavra inteira** (`\bTODO\b`), sem falso-positivo. Os demais
placeholders seguem como antes.
**Evidência:** RM13/julho _v2 (regeração 100% das fontes oficiais): após o fix, `05`/`06`/`12` rodaram
de ponta a ponta e o QA deu **APROVADO (0 pendências)**, 61 páginas.

## DEC-039 — Acumulado: ID como VALOR e Data Entrega como SERIAL PURO (sem hora / sem fuso)
**Incorreto (até v1.13):** o `02` gravava (a) o **ID como fórmula** `=B{r-1}+1` e (b) a Data Entrega
como `datetime.datetime(y,m,d)`. Ao passar um datetime Python para o Excel via **win32com**, o COM
aplicava o **offset do fuso local (−3h)**, gravando a data com **03:00** — e o Power BI, ao tipar
`text→date`, dava erro/lia com hora. O ID como fórmula também não era desejado (a usuária quer o
**número em si** na coluna ID).
**Regra vigente:** no append do `02`:
- **ID = VALOR** (`base_id + i + 1`, com `base_id` = último ID existente), **nunca fórmula**.
- **Data Entrega = SERIAL de data PURO** (inteiro `(date(y,m,d) − date(1899,12,30)).days`), que **não**
  sofre conversão de fuso e **não tem hora**; o **formato** é herdado da coluna (PasteSpecial de formatos
  da última linha — DEC-001). 'Sem data' continua como texto.
**Correção retroativa (RM13/julho):** as 267 linhas da V13 que já tinham saído com 03:00 foram
corrigidas in-place via Excel COM (`Value=int(Value2)`), com backup antes
(`output/RM13_PROD_backups/…__BACKUP_pre_datafix.xlsx`); IDs já estavam como valor no arquivo corrigido.
**Evidência:** usuária confirmou que o Power BI passou a ler/atualizar **sem erro** após a correção;
verificação: `datas com horário = 0`, `IDs com fórmula = 0`.
**Onde:** `scripts/02_atualizar_controle_demandas.py` (ID valor + data serial), `references/regras-excel.md`.

## DEC-040 — Etapa de REVISÃO FINAL (18) + status da medição colorido (inclui "Não faturado" vermelho)
**Contexto:** no RM14/agosto, ao montar, EU ESQUECI de renderizar a imagem do Dashboard (a pasta
work/03_IMAGENS/DASHBOARD ficou vazia e o QA de texto não pega isso) e o status do RESUMO saiu
"Em análise" (amarelo) quando deveria ser "Não faturado" (vermelho). A usuária pediu uma etapa final
de revisão pra flagrar esse tipo de falha.
**Regra vigente:**
- **`18_revisao_final.py`** roda como ÚLTIMA etapa (após 13/QA): confere imagens do corpo (pega
  Dashboard esquecido), Dashboard no PDF, Timeline (nº de páginas), cada Produto 1.X no topo, texto
  consolidado P1.2, valores QA_ESPERADOS, status da medição no painel-anexo, e **cada anexo da LISTA
  DE DOCUMENTOS ANEXOS existe fisicamente**. Só entregar com REVISÃO FINAL = APROVADO.
- **Status da medição colorido por valor** (`15`): mapa STATUS_COR = {Pago:A9D18E verde, Processado:
  F2CF00 amarelo, Em análise:F4B183 laranja, **Não faturado** (era FF0000 vermelho → **CINZA A6A6A6 desde
  DEC-041**)}; o `15` pinta N54-57
  do RESUMO (PAINEL) e a legenda "Não faturado" conforme `status_medicao`. O `06` (linha do tempo)
  passou a suportar **"Não faturado" (vermelho)** no _RANK/_LABEL/status_fill + legenda (antes só
  Pago/Processado/Em análise). `status_medicao` vem do `config/parametros.json` (ex.: ago/2026 = "Não
  faturado").
- **Dashboard**: render `dashboard_p01.png` do PDF do usuário é etapa obrigatória antes do `09` (o 18
  passa a acusar se faltar).
**Evidência:** RM14/agosto — após as correções, QA APROVADO (0) e REVISÃO FINAL APROVADO (0);
status "Não faturado" vermelho no quadro consolidado, no resumo por empresa e na linha do tempo (med 13).

## DEC-041 — Correções da 2ª revisão do RM14 (cor "Não faturado" = CINZA, assinatura, seta da timeline)
**Contexto:** na 2ª revisão do RM14/agosto a usuária pediu 3 ajustes (que substituem/afinam o DEC-040):
1. **"Não faturado" = CINZA `A6A6A6`** (não mais vermelho FF0000). Aplicado em `15` (STATUS_COR → pinta
   N54-57 do RESUMO + legenda K50) e em `06` (NAOFAT/LEGNF + legenda D4 com fonte preta). Vale para o
   quadro consolidado, o resumo por empresa e a linha do tempo de medições.
2. **Assinatura do bloco P1.1/P1.2** passou de "Pessoa C" para **"Pessoa A"** (`10`, `qa.md`). O bloco P1.3/P1.4 continua "Pessoa B".
3. **Seta da timeline (slide de mês com 1 marco):** a base da seta deve **encostar a bolinha** (sobre a
   linha do eixo) e a ponta terminar **logo abaixo do texto, sem sobrepor**. No RM14 a seta flutuava
   sobre o texto e não tocava a bolinha. Corrigido no **PPTX-base** (`P1.1 - Timeline Relatório Programa.pptx`,
   com backup): transladar a AutoShape da seta para baixo até a ponta ficar ~253pt (abaixo do "Parceiro") e a
   base ~301pt (encostando a bolinha em ~305pt no eixo). Como o `17` **recopia o PPTX-base** a cada run, a
   correção precisa ficar no PPTX-base (não só na cópia de trabalho). Depois `17` re-renderiza + republica
   o anexo. Regra permanente em `timeline.md`.
**Re-run das correções (ordem):** `15`(painel cinza) → publicar anexo (backup) → `05` → `06` →
(corrigir seta no PPTX-base) → `17` → `08` → `09` → `10`(nome) → `11` → `12`(publica _v2) → `13` → `18`.
**Evidência:** RM14/agosto v2 — QA APROVADO (0) e REVISÃO FINAL APROVADO (0); status cinza, seta encostando
a bolinha e abaixo do texto, assinatura "Pessoa A". Publicado como `..._v2.docx/.pdf` no G1
(sem apagar o v1).

## DEC-042 — Anexos: nunca zipar/mexer em P1.3/P1.4; Quadro de Anexos = varredura real das pastas
**Erro:** no RM14 a automação **criou zips** `Evidências 1.3/1.4 MM.YY.zip` (compactando os individuais que
os grupos subiram) porque a lista de anexos do `10` era **hardcoded** e esperava um único zip por evidência.
Resultado: os anexos ficaram **duplicados** (soltos + zipados). A usuária: *"APAGAR O ZIP, DEIXAR SÓ O QUE
ELES SUBIRAM… NUNCA MEXER NOS ANEXOS DOS PRODUTOS 1.3 E 1.4… NÃO ZIPAR NADA… NA LISTA DE ANEXOS SÓ CONFERIR
O QUE TEM EM CADA PASTA E LISTAR TODOS COM OS NOMES QUE ESTÃO LÁ."*
**Correção (permanente):**
- **NUNCA** zipar/agregar anexos; **NUNCA** renomear/mover/apagar/mexer nos anexos de **P1.3/P1.4** (só a
  usuária, manualmente). Deixar exatamente como os grupos enviam.
- **Quadro de Anexos = varredura real**: novo helper `cfg.listar_anexos_reais()` varre `G1\Anexos\{P1.1..P1.4}`
  (ordem P1.1→P1.4, arquivos por nome), ignora `~$`/`.tmp`/`Thumbs.db`/`desktop.ini` e o zip legado
  `Evidências 1.X …`, e devolve **um item por arquivo real** com o **nome exato**. O `10` monta a tabela
  "LISTA DE DOCUMENTOS ANEXOS" (REF romano | DESCRIÇÃO) a partir dessa lista (nada hardcoded, nº de linhas =
  nº de arquivos). O `18` valida que **cada arquivo real está listado no PDF**, que **não há `.zip`** na lista
  e que **não sobrou zip** `Evidências 1.X` nas pastas.
- **Corpo de P1.3/P1.4**: continua **intocado** (conteúdo CG preservado — CLAUDE.md; nunca alterar).
**RM14/agosto:** removidos os 2 zips (backup em `06_BACKUP\zips_removidos_*`); relatório regenerado com **17
anexos** reais (P1.1:4, P1.2:1, P1.3:6, P1.4:6), Anexo I–XVII com os nomes exatos. Publicado `..._v3.docx/.pdf`
no G1; QA APROVADO (0) e REVISÃO FINAL APROVADO (0).
> **Atualizada pela DEC-043 (POP 01/10/2026)** — ver abaixo o que continua e o que mudou.

---
> **A partir daqui: POP de Entrega de Produtos Programa (v2026-10-01).** Hierarquia: POP > Checklist > Modelo de
> Índice > Fluxograma > regras do repo. Fontes em `docs/referencias-oficiais/` (README = registro + SHA-256).
> Domínio completo em `docs/dominio/processo-entrega-mensal.md`; mapa antigo×novo em `docs/dominio/regras-pop-x-antigas.md`.

## DEC-043 — Anexos no padrão do POP: pasta ANEXOS única, produto no NOME (supera parcialmente a DEC-042)
**Regra nova (POP Quadros 2/3; usuária 01/10/2026):** em cada grupo, uma pasta única `ANEXOS`, **sem subpastas por
produto**; nome `[Código] - [Produto] - [Nome do arquivo] - [Rodovia]` (rodovia RD-101/RD-102A/RD-103/RD-104/GERAL);
todo não-PDF em `.ZIP`; ≤ 100 MB por arquivo (maior = Parte 01, 02…).
**Na automação do G1:** `prog_config.layout_anexos()` detecta o layout do mês — `por_produto` (legado, meses com
`Anexos/P1.x …`) ou `plano` (POP; padrão a partir de **16. Outubro 2026**, já reorganizado). No `plano`,
`anexo_p11_dir()`/`anexo_p12_dir()` = `G1\ANEXOS` e os anexos gerados usam `cfg.nome_anexo_g1()` →
`P1.1 - Gestão Executiva - Quadro Financeiro do mês de outubro - GERAL.xlsx`. `listar_anexos_reais()` lê a pasta
plana e ordena pelo prefixo P1.x. Override: `parametros.json "layout_anexos"`.
**O que da DEC-042 CONTINUA:** a automação do G1 **não** renomeia/move/apaga anexos de P1.3/P1.4 e o Quadro de
Anexos = varredura real com o nome exato.
**O que MUDOU:** "nunca zipar" valia para o ZIP *agregado* `Evidências 1.X` (continua proibido). O POP exige ZIP
**individual** de cada arquivo não-PDF — isso é feito **só no pacote** pela pipeline da Coordenação
(`scripts/entrega.py`), com o original preservado e registro antes→depois. Renomear para o padrão também só ocorre
no pacote, e só quando a proposta é completa (código/produto pela pasta + rodovia citada); o original vai para o
arquivo interno (DEC-045).

## DEC-044 — Designação oficial: "Relatório de Produtos Mensal"
**Regra (usuária, 01/10/2026):** o documento passa a se chamar **Relatório de Produtos Mensal**.
`cfg.termo_documento()` (padrão "Relatório de Produtos Mensal") alimenta o nome do arquivo final
(`Programa - Relatório Grupo 1 - Relatório de Produtos Mensal N°{n} - {período}`), a frase do encerramento (`10`) e o
**Documento** da linha recorrente no acumulado (`01`). **Não mudou:** o **Tipo** "Relatório Mensal" no acumulado
(categoria lida pelo Power BI/histórico — mudar exige decisão da usuária e ajuste do .pbix). QA (`13`/`18`) aceita as
duas grafias ao checar o nº antigo. Re-gerar competência antiga com o nome histórico: `"termo_documento":
"Relatório Mensal"` no `parametros.json`. Textos da CG (capa/apresentação) não são alterados.

## DEC-045 — Pacote final = a própria pasta do mês em "4. Produtos"; originais arquivados antes
**Decisão da usuária (01/10/2026):** "o pacote final eu vou só copiar a pasta do SharePoint nossa para a pasta
compartilhada com o Parceiro". Logo: a pipeline monta STAGING e FINAL **locais** (`output/ENTREGA_NN_Mes_AAAA/`),
revalida o FINAL e só então, com o **código de confirmação** do plano, `publicar` deixa a pasta do mês idêntica ao
FINAL. Antes de remover/substituir qualquer coisa, copia (e confere SHA-256) para (a) backup local e (b) **arquivo
interno** `…\Desenvolvimento de Produtos\Controle Interno Entregas\{pasta do mês}\{data}\` — é para lá que vão as
Listas de Demandas, Word/minutas e versões anteriores do relatório (`cfg.lista_demandas_path()` procura lá depois).
Recusa publicar se houver BLOCKER ou se a pasta mudou desde o `preparar`. **Nunca** envia nada ao Parceiro.

## DEC-046 — Assinaturas: verificação criptográfica real; gov.br ≠ ICP-Brasil; alteração pós-assinatura
`prog_entrega/assinatura.py` confere o hash do ByteRange × messageDigest e a assinatura com a chave do certificado
(suporta BER). Classes: VALIDADA / INVÁLIDA / NÃO VERIFICÁVEL / POSSÍVEL ASSINATURA VISUAL — REVISÃO HUMANA / NÃO
ENCONTRADA. Cadeia e revogação **não** são verificadas offline (limitação registrada). Acréscimo pós-assinatura só de
LTV (DSS/VRI) é aceito; qualquer outro (página/imagem/anotação) = NÃO VERIFICÁVEL → revisão humana.
**Evidência (auditoria somente-leitura de 14. Agosto 2026, 01/10):** relatório G2 e anexos com gov.br/ICP-Brasil
VALIDADOS; G3 (DocuSign + LTV) VALIDADO; **G1 `…NB014…_assinado.pdf`: assinado digitalmente (gov.br, Pessoa G) e DEPOIS alterado por "Preencher e Assinar" (imagem/rubrica na página)** → não aprovado automaticamente;
G4 sem assinatura digital; G5 só com indício visual (gov.br impresso). O POP pede "assinaturas digitais… quando
exigíveis" e não cita ICP-Brasil; a exigência ICP-Brasil vinha só do PROMPT antigo → `exigir_icp_brasil=false`
(config) até decisão da usuária; o emissor é sempre informado.

## DEC-047 — Reorganização do SharePoint (01/10/2026, autorizada por escrito nesta data)
`16. Outubro`, `17. Novembro`, `18. Dezembro 2026`: removidas as subpastas **vazias** por produto (0 arquivos,
re-checado antes de cada remoção; o OneDrive marcava as pastas como somente-leitura — só o atributo foi limpo),
`Anexos`→`ANEXOS` em G1–G5 e Índice criado na raiz. Índice "em branco" (competência + período + subdivisões por
grupo/produto, sem dados de anexos — *layout superado pela DEC-049: hoje só por grupo*) criado também na raiz de `07. Janeiro` … `15. Setembro 2026` (pedido da
usuária). Nada existente foi sobrescrito. Logs: `logs/entrega/estrutura_*.json` e `indice_*.json`.
Comando reprodutível: `python scripts/entrega.py estrutura --mes N [--aplicar]`.

## DEC-048 — Competência = mês civil (o corte "26→25" NÃO é regra vigente)
Verificado em 01/10/2026: `parametros*.json` (RM13/RM14: 01→último dia), nomes dos relatórios ("01.08 a 31.08"),
scripts e POP — não há regra 26→25 em lugar nenhum do repositório nem nos documentos oficiais. O G2 é filtrado por
**Nº Relatório** na Lista Mestre (não por data). Se a usuária confirmar outro corte, registrar aqui e ajustar
`periodo_inicio/periodo_fim` + `prog_entrega.regras.periodo_do_mes`.


## DEC-049 — Modelo de Índice r2: título "ENTREGA DE PRODUTOS DO PROGRAMA" e numeração por grupo
**Regra (usuária, 01/10/2026, com o Modelo revisado anexado):** o título passa a ser **"ÍNDICE DE ARQUIVOS : ENTREGA
DE PRODUTOS Programa"** (sai "– PRESTAÇÃO DE CONTAS") e a coluna Nº **reinicia em 1 em cada GRUPO** (1, 2, 3, 4…), porque
os grupos preenchem o índice ao mesmo tempo — numeração global gerava conflito. **Sem subdivisão por produto**
(usuária, 01/10/2026): só as linhas "GRUPO N" do Modelo; o produto está no nome de cada arquivo (padrão do POP). O
índice do pacote lista, em cada grupo, o Relatório de Produtos Mensal e depois os anexos por produto e nome. O nome do arquivo também sai
"Prestação de Contas": `Índice de Arquivos - Entrega de Produtos Programa - {Mês} {ano}.docx`.
**Onde:** Modelo em `docs/referencias-oficiais/v2026-10-01-r2/` (o anterior fica em `v2026-10-01/`, SUBSTITUÍDO);
`config/entrega_pop.json` (`indice.modelo`, `indice.nome_arquivo`); `indice.py` (modelo do mês e índice preenchido
numeram por grupo); `validacao.py` (confere 1, 2, 3… dentro de cada grupo).
**SharePoint (mesma data, pedido da usuária):** os 12 índices em branco do modelo anterior (07…18) foram trocados pelo
r2 (e, numa 2ª passagem, regenerados só por grupo, linhas 1–4 do Modelo) já preenchidos com competência/período — cada antigo foi conferido **em branco** imediatamente antes e
guardado em `logs/entrega/indices_substituidos/` (SHA-256 conferido). Comando: `python scripts/entrega.py indice
--mes N [--aplicar]` (só substitui índice EM BRANCO — de nome antigo ou de layout anterior; preenchido vira
pendência; idempotente).

## DEC-050 — POP de Entrega ainda SEM aprovação formal; divulgação informal como MINUTA (02/10/2026)
**Fato (usuária, 02/10/2026):** a Coordenadora Geral (**Pessoa A**) **ainda não revisou nem aprovou oficialmente** o POP,
o Checklist, o Modelo de Índice e o Fluxograma de `docs/referencias-oficiais/` (v2026-10-01 + r2). Eles continuam
sendo a referência de trabalho do projeto (hierarquia inalterada), mas com status **"aprovação formal pendente"** —
não chamar de "aprovados" em comunicações externas até a usuária avisar.
**Divulgação:** a pedido da usuária, foram geradas cópias carimbadas "MINUTA – EM REVISÃO" dos 4 documentos + uma
mensagem objetiva, para ela enviar **informalmente pelo WhatsApp** aos representantes de cada consorciada, para que
já organizem a entrega no novo formato. Registro, arquivos exatos e SHA-256: `docs/divulgacao/2026-10-02-minuta-whatsapp/`.
Script: `scripts/manutencao/carimbar_minuta.py` (originais só lidos). O envio é sempre feito pela usuária.
**Quando a usuária avisar a aprovação:** registrar data/aprovador aqui; se houver alteração, importar como nova
versão em `referencias-oficiais/`; preparar a comunicação oficial com os documentos sem carimbo.
**Aberto:** a mensagem pede o formato novo já em setembro, mas `15. Setembro 2026` está no layout antigo (G3–G5 sem
subpasta de anexos) — decisão da usuária pendente (criar `ANEXOS` com autorização × setembro como transição).
**Atualização 06/10/2026 (leitura da pasta, sem auditoria):** os próprios grupos criaram `Anexos` em G3, G4 e G5 de
`15. Setembro 2026`; o G2 subiu com subpastas aninhadas (`Anexos/Anexos/P2.1|P2.2|P2.4`, fora do "ANEXOS sem
subpastas" do POP). A decisão "criar ANEXOS × transição" ficou praticamente superada; a conformidade fica para o
`entrega.py auditar --mes 15`.

## DEC-051 — Organização do G2 no padrão POP e exceções SÓ do RM15 (06–07/10/2026)
**Feito (a pedido da usuária):** o G2 de `15. Setembro 2026` foi organizado no padrão do POP **só por cópia**:
`scripts/manutencao/organizar_g2_pop.py` (planejar → montar STAGING + auditoria oficial → aplicar → verificar) e
`atualizar_g2_pop.py` (incremental, após o retorno do engenheiro). Nome = `P2.x - Produto - Descrição da Lista Mestre -
Código - Rodovia`; rodovia = Lista Mestre conferida com a capa do PDF; nome ≤ 135 caracteres (caminho ≤ 250 também na
pasta do Parceiro); ZIP > 100 MB → partes ≤ 95 MB, ZIPs internos abertos em pasta, SHA-256 de cada arquivo interno
conferido; Relatório Mensal de Obra (RM) enviado em ZIP → PDF assinado como anexo próprio (byte a byte) + anexos em
partes. Índice: só a subdivisão GRUPO 2 é reescrita (outros grupos conferidos inalterados). Originais do G2 ficam em
`3. Desenvolvimento/G2/_Produto/<mês>/Anexos/` até o OK do engenheiro; quem apaga é a usuária. Ajustes e motivos por
arquivo: `config/g2_pop_ajustes.json`. Verificação: `output/G2_POP_NN_*/RELATORIO_VERIFICACAO_G2.md`.
**Exceções aceitas SOMENTE no RM15** (registradas em `revisoes_humanas.json`, presas ao SHA-256):
(a) 5 atas digitalizadas sem OCR (AT-678.1, 179.1, 182.1, 184.1, 185.1); (b) tabela "Controle de Produtos" do
relatório assinado divergente da Lista Mestre/anexos (NT-775.1, NT-818.1, NT-825.1) — o G2 já tinha assinado.
**Regra permanente (não muda):** PDF de anexo deve ser pesquisável (POP Quadro 4) e a tabela do relatório deve bater
com a Lista Mestre e com os anexos (Checklist Coord. 3 — item 13 da verificação). Sem registro de exceção, esses
achados são MANUAL_REVIEW; exceção não se estende a outra competência nem a outro arquivo.

## DEC-052 — Acesso ao SharePoint do Programa pela Microsoft Graph (app-only) (07/10/2026)
**Contexto:** o conector Microsoft 365 da sessão é só leitura e só enxerga o tenant da Dynatest (não vê
`contoso.sharepoint.com`). A escrita na pasta sincronizada pelo OneDrive depende da permissão da sessão e da
sincronização. A usuária forneceu o app **Entra ID - Programa** ("Programa - Contrato Cidade"; tenant `f29aad96-…`, client
`069340c0-…`) e pediu a integração para o projeto todo, independente da sessão.
**Feito:** `scripts/sharepoint.py` (CLI) + `scripts/prog_sharepoint/` (config, caminhos, graph, operacoes), só com a
biblioteca padrão (urllib, nenhuma dependência nova).
- Configuração **sem segredo** em `config/sharepoint.json` (versionado). O **client secret** fica em
  `%USERPROFILE%\.prog\sharepoint.env` (fora do repo e do OneDrive) ou em `PROG_SP_CLIENT_SECRET`. O código recusa
  arquivo de credenciais dentro do repositório e nunca imprime o segredo.
- Leitura: `testar`, `listar`, `info`, `baixar`, `comparar` (arquivo) e `comparar-pasta` (nomes/tamanhos, sem baixar
  nada; serve para saber se o OneDrive está em dia antes de gerar ou auditar).
- Escrita: `enviar` e `publicar-lote`.
- `testar` em 07/10/2026: permissões do app = `Sites.ReadWrite.All` + `Sites.Selected`; biblioteca `Programa`.
**Regras de escrita (aplicam a regra permanente 1 ao novo canal):**
- DRY-RUN sem `--aplicar`.
- Destino existente → BLOQUEADO, salvo `--substituir`, que **baixa o atual para `output/SHAREPOINT_BACKUPS/<data>/`**
  (com `.meta.json`) e envia com `If-Match` (eTag), para não sobrescrever uma versão que mudou no meio do caminho.
- Depois do envio, baixa de volta e confere: `OK` (SHA-256 igual), `OK_CONTEUDO` (Office em que só `docProps`,
  `customXml` e `.rels` diferem, porque o SharePoint grava metadados ao receber) ou `DIVERGENTE`.
- `publicar-lote` planeja todos os itens antes; se algum estiver BLOQUEADO ou com erro, não envia nenhum.
- Recusa qualquer caminho com "Compartilhado Parceiro". Registro em `logs/sharepoint/AAAA-MM.jsonl`.
- **Não existe** excluir, mover nem renomear. *(DEC-070: mover/renomear e copiar no servidor passaram a existir só para a reorganização autorizada da pasta do mês; excluir continua não existindo.)*
- Toda escrita exige autorização explícita da usuária na conversa.
**Não muda:** os scripts do G1 e o `entrega.py` continuam lendo e gravando pela pasta sincronizada
(`caminhos.json`). O acumulado continua só pelo `02`. Procedimento por etapa: `references/sharepoint-api.md`
(G1) e `.claude/skills/entrega-coordenacao/references/fluxo-operacional.md` (entrega). Testes:
`tests/test_sharepoint.py` (Graph falsa, sem rede).
**Primeiro uso (07/10/2026, autorizado):** `Medicao Parceiro/Planilha de Medição - Estruturadora - Setembro.xlsx` (novo,
OK_CONTEUDO, 0 células diferentes) e `P1.1 -Controle Timeline.xlsx` (substituído com backup em
`output/SHAREPOINT_BACKUPS/20261007-123130/`, SHA-256 OK). Recomendado trocar o segredo, que foi colado no chat.

## DEC-053 — G3 do RM15 organizado no padrão POP pela API (08/10/2026)
**Contexto:** o G3 (Certificadora) entregou setembro no layout antigo (relatório e Lista de Demandas na raiz;
`Anexos/P3.1|P3.2|P3.3` com subpasta `ANEXOS RVS`), e a pasta do G3 **não estava sincronizada** pelo OneDrive nesta
máquina (vazia localmente). **Feito (a pedido da usuária):** `scripts/manutencao/organizar_g3_pop.py`
(baixar → planejar → envio pelo `sharepoint.py` → verificar), só com ARQUIVOS NOVOS: 14 anexos renomeados direto em
`G3/Anexos/` + relatório renomeado na raiz + Índice (só GRUPO 3) substituído com backup/eTag. Nomes: RGC (P3.1) =
`Relatório de Gestão da Certificação - RL-P3.1-BI008-RGC-001 Rev NN - GERAL`; Relatório de Obra (P3.3) =
`Relatório de Obra - AAAA.MM Rev NN - RD-xxx`; RV = `Relatório de Visita de Obra|de Projeto Executivo - <código> -
RD-xxx` (os RVs são citados nos Relatórios de Obra da mesma rodovia e já iam como anexos em agosto).
**Concluído em 08/10/2026:** a usuária retirou o layout antigo (relatório com nome antigo e subpastas P3.x); a Lista de
Demandas fica na raiz do G3 (o grupo a entrega; sai do pacote na montagem final). A pedido da usuária, a Lista foi
corrigida: código das 4 linhas do P3.3 de 'P3.1 - 2026.09 - RD-…' para 'P3.3 - 2026.09 - RD-…' (Excel COM numa cópia,
substituída pela API com backup + eTag). Verificação: 12/12 PASS — `output/G3_POP_15_Setembro_2026/RELATORIO_VERIFICACAO_G3.md`.

## DEC-054 — Painel financeiro: % Executado do TOTAL era fixo; template do mês anterior no layout POP (13/10/2026)
**Erro (até o RM14):** o `15` atualizava, na linha TOTAL, valor mês, acumulado e saldo, mas **não o % Executado**
(`RESUMO_EMPRESAS_M8!E{total}` e `PAINEL_FINANCEIRO_M8!J58`). O valor **39,07%** vinha fixo do template e saiu igual no
RM14, onde o correto era 44,71%: no quadro consolidado, no resumo por empresa do anexo e nas imagens do relatório.
O texto do relatório e o cabeçalho do painel (`J8`) estavam certos.
**Correção:** o `15` grava `tot["pct"]` nas duas células. Validação do RM15: nenhuma célula com 39,07%; TOTAL = 47,42%,
igual a `J8`, ao texto e à Planilha de Medição.
**Também:** `painel_template_path()` só achava o painel do mês anterior em `Anexos\P1.1 …` (layout legado). Agora
procura também na pasta de anexos plana do POP (DEC-043), senão o RM16 (template = setembro) não encontraria o modelo.
**Regra permanente (usuária, 13/10/2026: "colocar como regra para não ocorrer de novo"):** o `18_revisao_final.py`
reprova se, no painel-anexo, o % Executado de qualquer linha do resumo (empresas e TOTAL, `J54..J58`) não for igual a
acumulado ÷ contrato, ou se `J58` ≠ `J8`. A checagem foi validada contra os dois meses: setembro passa e agosto é
reprovado (TOTAL 39,07% × 44,71%).
**RM14 já entregue:** correção do anexo/relatório de agosto pendente de decisão da usuária (o mês já foi enviado).

## DEC-055 — Produto fora do padrão vira "linha em branco" no Power BI: normalizar pelo código (13/10/2026)
**Erro (RM15):** a Lista de Demandas da CG trouxe 7 linhas com o produto "1.3 RELATÓRIO DE COMUNICAÇÃO INTEGRADA". O
nome padrão é "1.3 COMUNICAÇÃO SOCIAL" (versões 11–14 e aba `Listas` do acumulado, onde o Power BI liga produto →
frente). No dashboard apareceu uma frente em branco com 7 itens, e a Coordenação Executiva ficou com 11 em vez de 18.
**Regra permanente:** o `01` normaliza o nome do produto **pelo código** (`1.3`, `2.4`…): o nome canônico é o da
**versão mais recente do acumulado** (< N) que usou aquele código. Cada troca vai para o log como `AJUSTE produto`.
Código desconhecido gera um AVISO e não é trocado. **Produto vazio** herda o produto do mesmo documento em outra linha
do mesmo grupo (RM15: RO-549.1 RD-102A → 2.4); se não houver referência, gera AVISO ("corrigir na fonte"). As listas
dos grupos não são alteradas, só a consolidação. Avisar o grupo para usar o nome padrão.
**RM15:** 01 → 02 (Versão 15 refeita, com backup) → 03 → anexo do Controle republicado. Contagens inalteradas (298).

## DEC-056 — Fechamento do RM15: tabela de reuniões cortada, descrições recorrentes, anexos ZIP, caminho longo (14/10/2026)
1. **Tabela de reuniões com a última linha cortada (`04`).** A aba GOV da mestre tinha **quebras de página manuais**
   (linhas 7 e 10). Elas fatiavam a parte em 2 páginas, e parte fatiada **desliga** a checagem de corte no PDF
   (`cortes_no_pdf`). Correção: `ResetAllPageBreaks()` na cópia de impressão de cada parte (a paginação é do `04`) e
   medição complementar da altura **por parágrafo** (`medir_por_paragrafo`). Resultado: P1.1 em 1 parte de 8,15", sem
   corte (conferido visualmente).
2. **Descrições das linhas recorrentes (`01`).** Clonadas da versão N-1, traziam "…de **Junho**" desde o RM12, e a do
   P1.2 veio vazia do RM14. Agora o mês na descrição é trocado pelo mês corrente; se a descrição da N-1 estiver vazia, usa
   a última não vazia do histórico. O histórico (V12–V14) não foi alterado.
3. **ZIP na lista de anexos (`18`).** A checagem reprovava qualquer `.zip` (DEC-042). Pelo POP, o ZIP **individual** de um
   arquivo que não é PDF é o formato correto (DEC-043). Agora só reprova o ZIP **agregado** "Evidências 1.X".
4. **Caminho longo (`14`).** Os nomes do POP somados à pasta `Obsoleto\… - Arquivos Gerados\Anexos` passam de 260
   caracteres. O `14` agora copia com o prefixo `\?\`. O PPTX da Timeline é aceito com o nome do `17`
   (`Timeline_{Mês}{ano}.pptx`).
5. **Reuniões de setembro.** A lista final da Coordenação (14/10) foi acrescentada à mestre pela API, com backup
   (`output/RM15_PREPARACAO/reunioes_set26_aplicar.py`). Só houve limpeza de forma: quebras de linha, grafia de empresas
   e pontuação final. Data **24/07 → 24/09** na Governança semanal (digitação). Todas as mudanças estão registradas em
   `reunioes_set26_alteracoes.json`.
6. **Índice: GRUPO 1 gerado pelos nomes REAIS dos arquivos da pasta**, e não copiado das linhas digitadas pelo grupo (o
   travessão "–" da Carta 007 não batia com o arquivo, que usa "-").

## DEC-057 — Exceções da competência pela Coordenação: `aceito_excecao`, `relatorio_do_grupo`, aceite por grupo+etapa (14/10/2026)
**Contexto (RM15):** a usuária decidiu deixar passar, só neste mês:
- os anexos com nome sem rodovia (G1/CG, G4, G5);
- a assinatura manual do relatório do G4;
- a falta da Lista de Demandas na pasta do G2, que usa a Lista Mestre.

Confirmou também que `P5.1 - Relatório de Acompanhamento Jurídico.pdf` é o relatório do G5 (assinatura digital VALIDADA).
Até então, `revisoes_humanas.json` só transformava MANUAL em PASS.
**Regra:** o `revisoes_humanas.json` da competência aceita, sempre registrando **quem / quando / por quê**:
- `"decisao":"aceito"`: MANUAL → PASS (como antes);
- `"decisao":"aceito_excecao"` + `sha256` + `codigo`: BLOCKER ou MANUAL → PASS **como exceção**. A exceção aparece numa
  seção própria do relatório de prontidão ("Exceções aceitas pela Coordenação"), com o status original;
- `{"grupo","etapa","decisao"}`: aceite de achado **sem arquivo** (ex.: Lista de Demandas ausente);
- `"decisao":"relatorio_do_grupo"` + `sha256`: designa um PDF da raiz do grupo como o relatório, mesmo com nome fora do
  padrão "Relatório…". A assinatura dele continua sendo verificada.

Tudo fica preso ao **SHA-256**: se o arquivo mudar, a exceção deixa de valer. Também não passa para outra competência.
**Nunca** registrar por conta própria: só com decisão explícita da usuária na conversa.
**Onde:** `scripts/prog_entrega/pipeline.py` (`_confirmacoes`, `_designacoes`, `_aplicar_confirmacoes`), `relatorio.py`
(seção de exceções). Testes em `tests/test_entrega.py`, 4 novos.
**RM15:** 36 exceções de nome/assinatura + 1 aceite G2 + 1 designação G5. Pacote final revalidado: restam 1 BLOCKER
(relatório do G1 aguardando a assinatura da Pessoa A) e 1 MANUAL (PDF do P1.3 sem texto pesquisável).
**Publicação do RM15 (14/10/2026, autorizada pela usuária, código `d6c9b8fb6c`):** a pasta `15. Setembro 2026` ficou
idêntica ao manifesto. 51 itens saíram da pasta do mês, todos arquivados antes (SHA-256 conferido) em
`Controle Interno Entregas\15. Setembro 2026\2026-10-14_143045` e no backup local. Os itens retirados foram:
- as Listas de Demandas;
- o Word RevCG;
- os originais que, no pacote, viraram ZIP ou ganharam o nome corrigido.

O G1 entrou com assinatura **manual** (imagem), registrada como exceção do RM15, igual ao G4.
**Ajuste no validador:** numa pasta já publicada, a Lista de Demandas ausente não vira pendência se estiver no arquivo
interno do mês (`validacao._lista_arquivada`).

## DEC-058 — Data da CAPA = último dia do mês da competência (16/11/2026, regra permanente da usuária)
**Regra:** a data da capa — e, por coerência, a do encerramento ("Cidade, …") e a "Data Entrega" das linhas
recorrentes P1.1/P1.2 — é o **último dia calendário do mês da competência** (out → 31/10; nov → 30/11; fev bissexto →
29/02). **Nunca** a data real de elaboração no mês seguinte. **Supera a DEC-003** (antes: data impressa na capa da CG).
**Onde:** `prog_config.data_capa()` / `fim_competencia()` / `data_documento()` (o campo legado `data_documento` do
`parametros.json` é ignorado com aviso se divergir); `capa.py` lê a data no PPTX embutido da capa (OLE da CG) e, se
for outra, corrige o texto e regenera a imagem impressa (EMF) pelo PowerPoint; o `09` chama `garantir_data_capa()` na
cópia de trabalho; o `10` usa `cfg.data_documento()`; o `13` reprova se a capa do DOCX final ou o encerramento
divergirem. Testes: `tests/test_rm16_regras.py`. **RM16:** a capa da CG já trazia "31 de outubro de 2026".

## DEC-059 — Preservação de insumos externos antes do uso (16/11/2026, regra permanente da usuária)
**Regra:** todo arquivo vindo de fora do ambiente canônico (Downloads, Desktop, anexo da conversa, pendrive, e-mail,
pasta temporária, arquivo de terceiro) que alimente qualquer resultado é **primeiro importado**:
- cópia byte a byte e somente-leitura em `output/RM{n}_{ano}_{mm}/01_FONTES/EXTERNAS/`;
- registro em `_proveniencia.json`: nome e caminho de origem, data/hora da ingestão, data de modificação na origem,
  competência, finalidade, classe, SHA-256 e tamanho;
- cópia **durável** no SharePoint em `Desenvolvimento de Produtos/Fontes Externas/{NN. Mês ano}/` (arquivo novo, nunca
  substitui).

Só depois o dado é usado, sempre a partir da cópia. Reimportar o mesmo conteúdo é idempotente; conteúdo novo com o
mesmo nome vira `nome (2).ext` — o original importado nunca é sobrescrito. Arquivo AMBÍGUO não é importado sem decisão
(vira pendência). Se o insumo alimenta um artefato canônico (acumulado, Controle Timeline, mestre de reuniões, Planilha
de Medição), esse artefato é atualizado pelas regras dele.
**Onde:** `scripts/fontes_externas.py` (`inspecionar`, `importar`, `listar`, `plano-sharepoint`). Caixa de entrada usual
da usuária: `C:\caminho\exemplo\Downloads\<mês da entrega>` (transitória). Testes: `test_rm16_regras.py`.
**RM16:** importados `CONTROLE DE PRAZOS - PRESTAÇÃO DE CONTAS.xlsx` e `Relatório de Gestão de Contrato - Setembro -
2026 - RF-27-10-26.pdf`, publicados em `Fontes Externas/16. Outubro 2026` com a ficha de proveniência da 15ª medição.

## DEC-060 — Timeline: Prestações de Contas iterativas (16/11/2026, regra permanente da usuária)
**Regra:** a Timeline do P1.1 consulta também a planilha de **Prestação de Contas** (importada pela DEC-059) e registra
os **eventos ocorridos na competência**: envio, análise/retorno da Órgão Contratante, resposta/reenvio do Programa ou da executora,
atendimento parcial, nova pendência e aprovação. Não registra só as prestações novas: uma prestação antiga com
movimentação no mês gera evento naquele mês.
- Rodovia (RD-xxx) = prestação da obra; "Programa" = prestação da Estruturadora/contrato. O rótulo fica como na fonte (o
  Controle Timeline distingue RD-102 de RD-102A).
- Redação técnica e institucional, só com o que a coluna/célula comprova. "Retorno da Órgão Contratante com a Nª análise … com
  apontamentos que demandaram nova manifestação do Programa" só quando há retorno posterior registrado.
- Sem evento no mês ⇒ nada é inventado (comportamento histórico).
- Um marcador por **data** no slide: eventos da mesma data são agrupados, com o detalhamento completo na coluna
  "Detalhamento".

**Onde:**
- `prestacao_contas.py`: leitura, classificação e redação.
- `timeline_prestacao_excel.py`: grava a Base Timeline por **Excel COM** (o openpyxl descartava partes do arquivo).
- `timeline_slide_mes.py`: cria o slide do mês no PPTX-base duplicando o mês anterior e os marcadores reais
  (automatiza a DEC-040), com a descrição centrada na bolinha e hífen inseparável em "RD‑101".
- `references/timeline.md`. Testes: `test_rm16_regras.py`.

**RM16:** 11 eventos em 8 datas (02, 07, 14, 19, 22, 23, 28 e 29/10).

## DEC-061 — ~~Corte temporal do Controle de Produtos~~ — **REVOGADA em 16/11/2026 pela DEC-067** (histórico abaixo)
**Origem:** no Relatório de Gestão do Contrato – Setembro/2026, o Parceiro apontou que o RM15 incluiu 17 registros de outubro
e 4 sem data, e determinou "adotar critério de corte temporal… segregando as entregas posteriores ao período de
referência e eliminando os registros sem data", a comprovar no RM16.
**Regra (a partir do RM16; override `parametros.json "corte_temporal"`):** o `01` mantém só as entregas com data no
período, ou anteriores ainda não reportadas como "Entregue" no acumulado.
- **Posteriores** ao fim do mês: segregadas em `02_TRABALHO/corte_temporal_rmN.json` e incluídas **automaticamente** na
  competência seguinte (lidas pelo `01` do RM N+1).
- **Sem data**: eliminado; cobrar a data do grupo.
- O que já foi reportado como Entregue com a mesma chave não conta de novo.
- O texto do Controle declara o critério.

**RM16:** 216 → 201 itens. 15 posteriores segregados para o RM17 (5 do G3 de 09/11 e 10 do G2). 17 anteriores mantidos
(13 do G4 com data 30/09 e 4 do G2).

## DEC-062 — 15ª medição: planilha RECONSTRUÍDA do quadro impresso pelo Parceiro (16/11/2026)
**Contexto:** o Excel editável oficial da 15ª medição (Parceiro/Pessoa H) não havia chegado. O quadro "15ª MEDIÇÃO" está na
p.84 do Relatório de Gestão do Contrato – Setembro/2026 (PDF vetorial, assinado pelos fiscais em 30/10/2026).
**Regra:** nessa situação:
1. Extrair o quadro por coordenadas (`medicao_pdf.py`): só o que está impresso; "-" ou vazio = None.
2. Reconstruir a planilha copiando a do mês anterior e replicando o bloco N-1 como bloco N via Excel COM. Preencher
   só QUANT MEDIDA e NOTA AVALIATIVA (`medicao_reconstrucao.py reconstruir`).
3. Conferir campo a campo: produtos, grupos, total e % de execução. Divergência acima de R$ 0,01 bloqueia.
4. Registrar a proveniência: propriedades do arquivo e ficha JSON em `Fontes Externas`.

**Status:** o do documento; se não houver status formal, "Em análise". **No relatório não se menciona a reconstrução**
(usuária, 16/11/2026).
**Quando o oficial chegar:** importar (DEC-059) e rodar
`medicao_reconstrucao.py comparar --oficial … --reconstruida … --medicao N`.
- Equivalentes: promover o oficial (substituir com backup) e atualizar a ficha.
- Divergentes: refazer 15→05→06→09…18.

**RM16:** 271 campos conferidos, 0 divergências. Total, acumulado e % conferidos *(valores omitidos na versão pública)*. Status "Em
análise" (o PDF só diz "encaminha-se… para pagamento").

## DEC-063 — Medição nomeada pela competência de EXECUÇÃO; aba de medição detalhada completa (determinação (I) do Parceiro)
**Erro (até o RM15):**
- a referência dizia "Referência: 14ª Medição – Setembro/2026", ou seja, o mês de processamento;
- a tabela "PLANILHA DE MEDIÇÃO" da aba `MEDICAO_*` do painel trazia os valores da 12ª/11ª, porque o `15` só atualizava
  os totais do topo.

**Regra:** `cfg.competencia_medicao()` = mês anterior ao do relatório.
- Título da seção: "Análise do processo de medição do mês anterior - {mês da execução}/{ano} - {N}ª medição".
- Painel: "Referência: {N}ª Medição – {mês da execução}/{ano}" e "Contrato… | Referência: {mês da execução}/{ano}".
- O `15` preenche a tabela detalhada inteira (por produto, subtotais e TOTAL GERAL) e o rótulo "Valor da Nª medição".
- O painel do mês anterior também é lido de dentro do ZIP do POP.

## DEC-064 — Versão EM ELABORAÇÃO (16/11/2026)
**Regra:** com insumo pendente, o relatório é gerado assim mesmo, com um **quadro "SEÇÃO EM ELABORAÇÃO — …"**
(sombreado, emoldurado) no lugar exato do conteúdo dependente. O controle é por `PENDENCIAS` em `config/dados_mes.py`
(chaves `reunioes`, `p12`, `p12_consolidado`, `dashboard`). Enquanto houver chave:
- o nome do arquivo ganha " - EM ELABORAÇÃO" (`cfg.status_documento()`);
- o QA (`13`) reprova;
- a versão vai para `Desenvolvimento de Produtos/Relatório G1 em Elaboração/{NN. Mês ano}/` — **nunca** para a raiz do
  G1 do mês.

Para finalizar: fornecer o insumo, remover a chave e remontar (o sufixo some sozinho).

## DEC-065 — P1.1: síntese do Plano de Gestão Integrada de Riscos (determinação (III) do Parceiro; 16/11/2026)
**Regra:** todo mês o P1.1 traz, após a medição, "Síntese do Plano de Gestão Integrada de Riscos - {mês}/{ano}." (item
3.6.2.1 do contrato). A fonte são os produtos do P4.3 do mês: matriz do Programa, matriz IRIS e relatório de evolução.
O conteúdo cobre:
- totais por nível e a variação;
- unidades e riscos prioritários;
- tabela dos riscos críticos do Programa com dono do risco e medida de mitigação (sem e-mails/telefones).

Conteúdo em `dados_mes.RISCOS`; montagem no `09`.

## DEC-066 — Anexos não-PDF da Coordenação como ZIP individual; Quadro de Anexos com o nome final (16/11/2026)
**Regra:**
- O `20_zip_anexos_g1.py` roda depois de 03/05/15/04 e antes do 10. Ele compacta os anexos .xlsx de P1.1/P1.2 em ZIP
  individual na pasta ANEXOS; o .xlsx fica em `07_SAIDA`.
- O `10` lista os anexos não-PDF pelo nome `.zip` final, o mesmo do Índice e do pacote.
- `05` e `18` leem o painel de dentro do ZIP (`cfg.painel_anexo_xlsx()`).
- Nunca mexe em P1.3/P1.4.

**Também (16/11/2026):** o OneDrive desta máquina deixou de hidratar arquivos "somente online" (acesso negado).
Solução:
1. `sharepoint.py espelhar <remoto> --destino C:\caminho\exemplo\PROG_ESPELHO` cria uma cópia de leitura com
   `_espelho.json` (SHA-256, eTag, data).
2. Os scripts rodam com `PROG_ROOT` apontando para o espelho.
3. A publicação é feita com `publicar-lote`, conferindo antes que o original não mudou desde o espelho.

Raiz curta para não passar de 260 caracteres.

## DEC-067 — Controle de Demandas = TUDO o que os grupos declararam (revoga a DEC-061; usuária, 16/11/2026)
**Erro:** o corte temporal (DEC-061) tirou de outubro 15 itens por terem data/entrega em novembro (5 do Certificadora
de 09/11 e 10 do G2), levando-os ao RM17; e a dedup por chave composta descartava linha repetida da fonte.
**Regra vigente:** o Controle de Demandas da competência reproduz integralmente o que os grupos **declararam**: G1/G3/G4/G5
= todas as linhas das Listas de Demandas da entrega do mês; **G2 = Lista Mestre filtrada pelo N° Relatório da
competência** (equivale à Lista de Demandas). A data de upload/conclusão **não** muda a competência (atraso é achado da
auditoria do POP). Linha repetida na fonte é **mantida** e registrada (`entregas_stats.json → duplicadas_na_fonte`) para o
grupo corrigir. O `01` registra `declarados_por_fonte` = `considerados_por_fonte`. **G2:** conciliar cada linha da Lista
Mestre filtrada com os arquivos de `G2\ANEXOS` (código do documento + rodovia) e listar o que faltar.
**Correção do RM16 (16/11/2026):** acumulado refeito a partir do original anterior ao RM16 — Versão 16 = **217** (G1 13,
G2 144, G3 5, G4 34, G5 15, recorrentes 6); a versão de 201 linhas ficou como evidência nos backups
(`output/SHAREPOINT_BACKUPS/20261116-0811*`, `06_BACKUP/…065541.xlsx`); `corte_temporal_rm16__REVOGADO_DEC-067.json`.
G2: 139 de 144 linhas com anexo direto; faltam os RM-013.1 (RD-104) e RM-014.1 (RD-101/102A/103); divergências PA-823.1
(lista RD-102A × arquivo RD-103) e PA-847.1 (lista RD-104 duas vezes × arquivos RD-104 e RD-101).
**A regra de data permanente é só a da CAPA (DEC-058).**

## DEC-068 — Timeline: Prestação de Contas de Obra × da Estruturadora, com evidência do SharePoint (rev. da DEC-060)
**Regra:** nomenclatura institucional "Prestação de Contas de Obra da RD-xxx (Nº Trimestre – período)" e "Prestação de
Contas da Estruturadora (Programa) referente a {mês ano}". A classificação e os dados se confirmam no SharePoint:
`2. Compartilhado interno/07_Resposta Prestação de Contas - Obras/<Nº TRIMESTRE>/<RD>/REV n - <data>` (obras) e
`2. Compartilhado interno/00_Prestação de Contas/0x_Resposta Prestação de Contas - Mensal|Trimestral|Anual - Programa`
(Estruturadora). As evidências da competência (rótulo correto, período, documento de cada evento — Relatório n de Análise
da Órgão Contratante, Manifestação nº — e eventos documentados que a planilha não registra) ficam num JSON
(`02_TRABALHO/prestacao_evidencias_rmN.json`, publicado em `Fontes Externas`) lido por
`prestacao_contas.marcos_da_competencia(..., evidencias)`. O valor original fica na fonte bruta (`referencia_fonte`).
**RM16:** "RD-102" da planilha → **RD-102A** (Manifestação nº 000/2026: "Prestação de Contas Trimestral da Rodovia
RD-102A"); evento acrescentado: 30/10, Ofício Executora D nº 000/2026 (RD-104); 12 eventos em 9 marcos.

## DEC-069 — RM16: assinante, G4 e ordem dos trabalhos (usuária, 16/11/2026)
- Assinatura do P1.1/P1.2: **Pessoa D** (`parametros.json "assinante_p11_p12"`); só o nome no bloco; status AGUARDANDO
  ASSINATURA FINAL — nada de assinatura simulada; assina depois de finalizado.
- G4: a assinatura manual do relatório de outubro foi **aceita nesta entrega** (`revisoes_humanas.json`, presa ao SHA-256);
  a característica segue registrada na auditoria.
- A auditoria de conformidade completa G1–G5 fica **para depois** da produção do relatório; a entrega dos grupos não é
  alterada nesta etapa (o que o grupo pôs na pasta conta).
  *(Atualizado pela DEC-070: a usuária autorizou, na mesma data, a reorganização das entregas no formato do POP.)*

## DEC-070 — Reorganização da pasta do mês no formato do POP, com cópia "como enviado" antes (usuária, 16/11/2026)
**Pedido:** "salva uma cópia igual do que todo mundo enviou e como enviou em uma pasta no desenvolvimento. E na pasta de
produtos pode ajustar tudo pro formato do Parceiro e completar o índice sem nenhuma divergência" — incluindo colocar os RMs
do G2 que faltavam e renomear os anexos de todos os grupos que não seguiram o POP.
**Regra permanente (ordem obrigatória):**
1. **Cópia "como enviado"** NO SERVIDOR (Graph `copy`, bytes idênticos) da pasta do mês para
   `3. Desenvolvimento/G1 - Coordenacao/Entregas Recebidas/<pasta do mês>`, conferida arquivo a arquivo (tamanho +
   quickXorHash), com manifesto ao lado (autor, data de envio e SHA-256 de cada arquivo). Sem cópia OK, nada é mexido.
2. Decisões por arquivo em `config/organizacao_pop.json[<mês>]` — renomear, retirar do pacote, arquivos novos —, cada uma
   com motivo. Quando o próprio documento contradiz a lista do grupo, vale o documento (DEC-051) e o grupo corrige a lista.
   As Listas de Demandas **não** são editadas (ficam na raiz do grupo e saem do pacote na montagem final).
3. `scripts/manutencao/reorganizar_mes_pop.py planejar` confere o estado final (nome no padrão, produto do grupo,
   colisão, ≤ 100 MB, caminho ≤ 250 também na pasta do Parceiro) e gera o Índice completo G1–G5 pelos nomes reais, com
   extensão (relatório primeiro, depois anexos por produto e nome; numeração por grupo). → código.
4. `aplicar --confirmar <código>`: envia os novos, move/renomeia pela API (`If-Match`, conflito = falha, quickXorHash
   conferido), **retira do pacote movendo** para `Entregas Recebidas/<pasta> - Retirados do pacote/` (nunca apaga) e
   substitui o Índice (backup + eTag). `verificar`: SharePoint × estado planejado × Índice.
**API (`prog_sharepoint`):** novas operações `copiar_pasta`, `mover` (restrito às pastas permitidas; recusa
`Compartilhado Parceiro`), `criar_pasta`, `arvore`. **Excluir continua não existindo.** Testes:
`tests/test_sharepoint_reorganizar.py`.
**RM16 (16/11/2026):** cópia de 234 arquivos (963 MB) OK. Decisões: G1 CG 13 anexos + "GERAL"; G2 26 renomeações (forma;
PA-823.1 → RD-102A, NT-835.1 → RD-101 e PA-821.1 → RD-103 pela capa; naturezas de NC-535.1/NC-536.1 pelo documento;
PA-846.1 é Parecer), PA-847.1 "RD-101" retirado (cópia idêntica do RD-104), RAD-003.1 desmembrado (PDF + anexos, sem
ZIP em ZIP) e os 4 RMs (RM-013.1 RD-104; RM-014.1 RD-101/102A/103) no padrão DEC-051 (PDF assinado VALIDADO + 12
partes ≤ 95 MB); G3 17 anexos para `ANEXOS` sem subpastas; G4 34 (produto do catálogo, rodovia no fim, conta 70.507-1 =
RD-102A pela Lista do G4); G5 9 + 4 planilhas em ZIP individual ("ID n" dentro da descrição). Índice: 245 linhas.
**Pendências de conteúdo (dos grupos, não corrigidas pela Coordenação):** Lista Mestre do G2 (NT-835.1 e PA-821.1 com
rodovia errada; PA-847.1 repetido), Lista do G3 (P3.1 → P3.3), Lista do G4 (P4.2 "09.2026"), OCR (G1 P1.3 nº 2, atas do
G2), planilhas do G5 sem PDF correspondente, `.7z` dentro dos anexos do PA-335.x.

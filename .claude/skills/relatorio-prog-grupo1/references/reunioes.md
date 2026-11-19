# Reuniões (Controle de Reuniões)

## Fonte
`P1.2 - Controle de Reuniões.xlsx` — usar a **versão mais recente** enviada (comparar mtime/quantidade
de reuniões). Abas relevantes:
- **"Controle Geral"** — log geral de reuniões de **governança/coordenação** (Órgão Contratante semanal, CERT,
  Escritório Jurídico, Parceiro, executoras Executora A/Executora B/Executora C, Executora D…). → alimenta **P1.1 "Reuniões de Governança"**.
- **"Reuniões Stakehold 1-2"** — reuniões **institucionais/stakeholder**. → alimenta **P1.2
  "Integração Institucional"**.
- (As abas "Reuniões de GOV 1-1"/"Stakehold 1-2" podem não estar atualizadas para o mês; a base do
  mês costuma estar em "Controle Geral". Ver DEC-011.)
Cabeçalho na linha 3: B=Data, C=Participantes, D=Assuntos, E=Anexos.

## Cópia mensal (anexo)
`P1.2 - Controle de Reuniões do mês de {mês}.xlsx` na pasta P1.2 dos anexos. Filtrar cada aba ao
**mês** (apagar linhas fora do mês na cópia; **nunca** no original). Ocultar a aba não usada.

## Tabela no corpo (script 04) — DEC-027
A imagem é **consequência de uma cópia do Excel corretamente formatada**, em **escala 1:1**: nunca se
corrige o PNG depois. Ordem: dados → cópia → formatação → dimensões → quebras → área/escala de
impressão → renderização → validação → inserção.
- **Escala 1:1**: as margens de impressão da cópia são calculadas para que a largura natural de
  `Data+Participantes+Assuntos` seja a **largura útil do corpo** (6,891" no retrato; 10,4" nas páginas
  paisagem). Com `Zoom=100` e a imagem inserida nessa mesma largura, **pt no Excel = pt no PDF**.
- **Calibração obrigatória**: o Excel **reporta** `Width/Height` em pt que **não** são os pt impressos
  (medido: largura ×1,0225, altura ×0,9357). O script mede os dois fatores num PDF de calibração e
  dimensiona/pagina em **pontos impressos**. Formas (logo) continuam posicionadas em pt reportados.
- **Fontes (alvo visual, ≈ texto corrido do relatório)**: conteúdo **9,5 pt**, cabeçalho de coluna
  11 pt, título da faixa 13 pt (`PROG_REUNIOES_FONTE_PT` ajusta).
- **Larguras**: Assuntos leva toda a folga (retrato 11,5/20,5/68%; paisagem 7,6/15/77,4%); Data sempre
  larga o suficiente para `dd/mm/aaaa` em uma linha.
- **Negrito reaplicado do zero** (independe da formatação recebida): em Participantes, o órgão e os
  `:`; em Assuntos, o título de cada bloco (linhas que não começam com bullet).
- **Alinhamentos**: Data centralizada (H+V); **Participantes centralizado (H+V)**; Assuntos topo/
  esquerda; wrap sempre.
- **Alturas**: `AutoFit` medido com a coluna 1,5% mais estreita, + folga 1 pt — e depois o **ciclo de
  verificação de corte no PDF renderizado**: mede a última linha de texto de cada linha contra a
  própria borda, soma o mínimo necessário e **repagina** (até 3 passos). Conferência dupla: nenhum
  termo das células pode desaparecer do PDF (quando a sobra passa de uma linha o Excel simplesmente
  não desenha as linhas seguintes).
- **Faixa/logo validadas antes do print**: título presente, faixa contínua em B1:D1 e **logo Programa à
  direita, inteira dentro da faixa** (reinserida de `assets/logo_prog.png` se faltar). Faixa 24 pt e
  espaçador 3 pt — cada ponto ali se repete em **todas** as partes.
- **Quebras**: sempre entre **reuniões completas**; cada parte é exportada de uma cópia da aba contendo
  só as suas linhas + linhas 1-3, então **faixa, logo e cabeçalho se repetem** com as mesmas larguras.
- **Aproveitamento vertical**: a 1ª parte é dimensionada pelo espaço que sobra abaixo do texto do P1.1,
  **medido no DOCX montado** (exige 2ª passagem — ver SKILL.md §4). Sem espaço útil (< 2,5"), a tabela
  começa na página seguinte, sem comprimir. O log informa a ocupação de cada parte e marca
  `<<< EVITÁVEL` quando a próxima reunião caberia na sobra.
- **Retrato × paisagem automático**: o script formata, pagina e renderiza nas **duas** orientações e
  publica a de **menor desperdício** (`PROG_REUNIOES_ORIENTACAO=auto|retrato|paisagem`). Desperdício =
  espaço perdido na página do texto + sobras das partes; em paisagem contam-se **todas** as sobras e a
  página do texto inteira (mudar de orientação obriga página nova).
- Gerar dois conjuntos: governança (P1.1) e institucional (P1.2), em pastas de imagem separadas.
- `08` **não recorta de novo** pastas com marcador `.precropped` (o 04 já recorta exatamente na tinta;
  recortar com `pad` desalinharia a tabela do corpo do texto).

## Textos (ver `regras-redacao.md`)
- **P1.1 "Reuniões de Governança"**: síntese técnica **somente** das reuniões de governança do mês
  (aba Controle Geral) — decisões, acompanhamentos, prestação de contas, riscos, articulação.
- **P1.2 "Integração Institucional…"**: narrativa institucional (aba Stakehold + interações
  documentadas), organizada por frentes; incluir o parágrafo-nota do RM11 (apresentação em síntese
  técnica: planilha com datas/participantes/assuntos). **Não misturar** com a base P1.1.

## Aproveitamento — limite conhecido (RM13/julho)
Retrato venceu com folga: 3,99" de branco contra 15,71" da paisagem (em paisagem a tabela fica 37%
mais baixa, mas perde-se a página do texto e a altura útil cai para 5,9"). O branco que **resta** no
retrato é de granularidade: uma reunião de 3,94" não entra numa sobra de 3,17". Zerar isso exigiria
**dividir uma reunião entre páginas** ou trocar a imagem por **tabela nativa do Word** — nenhuma das
duas está implementada (decisão do usuário pendente).

# Timeline (Linha do Tempo de Acompanhamento Macro)

## Fonte dos MARCOS = Excel de Controle da Timeline (DEC-035 — atualizado)
- **FONTE OFICIAL DOS MARCOS = `…\Desenvolvimento de Produtos\P1.1 -Controle Timeline.xlsx`.**
  Aba **`Base Timeline`**: cada linha é um marco com colunas `Competência do Relatório` (data, ex.:
  01/07/2026 formatada "julho/2026"), `Data do Marco` (dd/mm/aaaa), `Categoria`, `Empreendimento(s)`,
  `Marco / Título da Timeline`, `Detalhamento`, `Referência / SEI / Ofício`, `Incluir na Timeline?`
  (usar só "Sim"). Aba `Timeline Mensal` = saída filtrada por competência (posição PPT acima/abaixo,
  data "dd/mm", texto pronto). Aba `Listas` = domínios.
- **Processo mensal (obrigatório):** ler os marcos da competência do mês (Competência = 1º dia do mês
  do relatório). **Conferir se estão atualizados**; **se não houver marcos (ou incompletos), PERGUNTAR
  ao usuário quais são os marcos ANTES de gerar** o relatório, **preencher o Excel** com o que ele
  informar (aba Base Timeline) e só então prosseguir. **Nunca inventar marco** (DEC-037).
- **A partir dos marcos, GERAR o slide mensal** da timeline no layout/elementos/alinhamentos corretos
  (ver "Construção" abaixo) e renderizar. O PPTX-base `P1.1 - Timeline Relatório Programa.pptx`
  (`cfg.timeline_pptx_src()`) tem uma slide por mês; a slide do mês corrente já traz os **grupos de
  marcador posicionados** (Group 9 × nº de marcos, alternando acima/abaixo) — **preencher** cada grupo
  com data+descrição do marco (na ordem por data), **sem duplicar**. **NUNCA** usar o PDF renderizado
  do espelho Parceiro como fonte.
- Legado (não é mais a fonte dos marcos): antes usava-se o PPT "editável acumulado" direto; agora os
  marcos vêm do Excel de Controle e o PPT é só o **canvas de layout** para render.
- **Mês NOVO (o PPTX ainda não tem o slide do mês) — DEC-040:** criar o slide via PowerPoint COM:
  `pres.Slides(slide_do_mes_anterior).Duplicate()`; na cópia (1) trocar a faixa "[ Mês /ano ]" para o
  mês novo (replace do nome do mês em todas as textboxes, inclusive dentro de grupos); (2) reduzir aos
  marcadores do mês — **apagar os grupos de marcador extras** e, no(s) mantido(s), setar a **data**
  "DD/MM" e a **descrição** do marco. A caixa de DATA às vezes vem VAZIA (era o "2º marco de data
  igual") → escrever "DD/MM" na textbox vazia abaixo do eixo + Font teal **#1C5560** bold. **Backup do
  PPTX antes.** Depois o `17` renderiza (confere que o slide tem os marcos do Excel).
- **Não recriar** a timeline do zero nem por aproximação: **duplicar os elementos-modelo reais** da
  slide do mês (placeholders `DD/MM`/`XXXXXX`) — ponto, seta, caixas (DEC-013).
- O `07` gera um **ponto de partida** fiel à técnica (DEC-025); o editável pode ser **refinado à mão**
  e sobrescreve o compartilhado (com backup). Exporta o anexo PDF `P1.1 - Timeline de Acompanhamento
  Macro até {mês} {ano}.pdf`.

## Prestações de Contas — eventos ITERATIVOS do mês (DEC-060, regra permanente)
Todo mês, ANTES de conferir/perguntar os marcos, consultar a planilha de **Prestação de Contas** (`CONTROLE DE PRAZOS -
PRESTAÇÃO DE CONTAS.xlsx`, recebida da Coordenação; importar antes pela DEC-059):
0. **Conferir no SharePoint** (DEC-068) as pastas `07_Resposta Prestação de Contas - Obras/<trimestre>/<RD>/REV n - <data>` e
   `00_Prestação de Contas/…- Programa`: rótulo correto da rodovia, período, Relatório/Manifestação de cada evento e eventos
   documentados que a planilha não tem → `02_TRABALHO/prestacao_evidencias_rmN.json` (publicar em `Fontes Externas`).
1. `python prestacao_contas.py <planilha importada> <ano> <mês> [evidencias.json]` lista os eventos DO MÊS (envio, Nª análise da Órgão Contratante,
   Nº retorno do Programa/da executora, aprovação), já redigidos. Prestação antiga com movimentação no mês GERA evento.
2. `python timeline_prestacao_excel.py <Controle Timeline atual> <saída> <planilha> <ano> <mês> [títulos.json|-] [evidencias.json]` grava
   na Base Timeline UMA linha por data (título curto para o slide; detalhamento institucional completo; referência =
   planilha + células) — por Excel COM, idempotente. Mês sem evento ⇒ não grava nada (não inventar).
3. Outros marcos do mês (repasses, licenças, decisões) seguem a regra DEC-035: conferir/perguntar.
4. Mês NOVO no PPTX-base: `python timeline_slide_mes.py <PPTX-base> <Controle Timeline> <ano> <mês> <pasta de backup>`
   (duplica o slide do mês anterior, troca a faixa, recria N marcadores alternados acima/abaixo a partir dos modelos
   reais). Renderizar e CONFERIR visualmente; ajustar títulos (≤ ~80 caracteres) se houver colisão. Depois o `17`.

## Técnica de montagem — datas múltiplas e não-sobreposição (DEC-025)
1. **Mesma data = MESMO ponto** com **2 setas** (uma acima + uma abaixo); **data única** (a 2ª é apagada).
2. A **seta do lado da data** é **curta e começa depois da data**; a do outro lado liga o ponto ao texto.
3. **Descrições escalonadas** (alterna perto/longe da linha no mesmo lado); marcadores **sob o título
   "Timeline"** usam o nível baixo (não colidem com ele).
4. **Setas de comprimento variável** — vão do ponto até logo abaixo/acima de cada texto (sem seta
   sobre texto). **Largura da descrição ≤ espaçamento** entre pontos (não invade o vizinho).
- O PPT tem **uma página por mês** (Jul/2025 → mês corrente), com uma faixa verde superior
  ("Progresso das atividades executadas [ {Mês}/{ano} ]" + logo), o título "Timeline", um eixo
  horizontal (linha teal `#1C5560`) e **marcadores**.
- **Marcador = GROUP** contendo: caixa de **data** ("DD/MM"), a **seta** (AutoShape com linha na cor
  do marcador), um sub-grupo com a **bolinha** (freeform preenchido) e a caixa de **descrição**.

## Seta do marcador — base na bolinha, ponta abaixo do texto (DEC-041)
- A **base** da seta deve **encostar a bolinha** (que fica sobre a linha do eixo, y≈305pt) e a **ponta**
  terminar **logo abaixo da descrição, sem sobrepor** o texto. Nunca deixar a seta atravessando o texto
  nem flutuando com um vão até a bolinha.
- **Mês com 1 marco (marcador central)**: a seta é vertical, base ~301pt (topo da bolinha) e ponta ~253pt
  (abaixo da última linha do texto). Se, após corrigir/criar o slide, a seta ficar sobre o texto ou solta,
  **transladar a AutoShape da seta** (mesma altura) até esses alvos. Como o `17` **recopia o PPTX-base** a
  cada execução, a correção deve ficar no **PPTX-base** `P1.1 - Timeline Relatório Programa.pptx` (backup antes),
  não só na cópia de trabalho.

## Cores dos marcadores
- **Verde-claro `#00BF63`** e **verde-escuro/teal `#1C5560`** (bolinha e seta na mesma cor).
- Para o mês: p.ex. **9 marcos → 4 verde-claros + 5 verde-escuros, alternando** (ajustar ao nº de marcos).

## PPTX-fonte ausente → interromper só a Timeline (DEC-022)
Se o `.pptx`-fonte não for localizado (em `G1/Anexos/P1.1` nem em `01_FONTES`), o `07`: (1) **não**
inventa/reconstrói a Timeline e **não** usa o PDF anterior como fonte editável; (2) interrompe apenas
esta etapa (as demais seguem); (3) loga **"TIMELINE PPTX NÃO LOCALIZADA"**; (4) registra pendência de
QA (`timeline_pptx_ausente`), que some quando o `.pptx` é fornecido e o `07` reexecuta. Note que o mês
concluído costuma ter só o **PDF** da Timeline no anexo — o PDF **não** serve de fonte; é preciso o `.pptx`.

## Fallback do slide-alvo → conferência visual obrigatória (DEC-020)
O `07` acha a slide-alvo pelo placeholder **`XXXX`**. Se **não** encontrar, usa o **último slide** como
fallback (comportamento mantido). Quando isso acontece, o script: (1) gera a Timeline; (2) loga
**WARNING**; (3) grava `cfg.set_qa_warning("timeline_fallback_xxxx", …)`. O `13_qa.py` então marca
**"CONFERÊNCIA VISUAL DA TIMELINE NECESSÁRIA"** e o veredito vira **REVISÃO NECESSÁRIA** — a etapa
**não** é aprovada automaticamente até conferir visualmente a slide do mês. O aviso é idempotente
(some se, numa reexecução, o `XXXX` for encontrado).

## Construção (script 07, PowerPoint COM)
1. Copiar o PPT para trabalho. Trabalhar na **slide do mês** (a última / a slide-template com
   placeholders "DD/MM"/"XXXXXX", se existir).
2. Obter os **templates reais**: um marcador "abaixo" (desc abaixo do eixo) e um "acima" (desc acima),
   copiando grupos existentes (ex.: da slide de fevereiro).
3. Para cada marco: `Duplicate()` do template adequado; posicionar em X (distribuir ao longo do eixo,
   com recuo nas bordas); **recolorir** bolinha (`Fill.ForeColor.RGB`) e seta (`Line.ForeColor.RGB`)
   para claro/escuro; escrever a **data** e a **descrição** (substituir placeholders); centralizar/
   estreitar a caixa de descrição.
4. **Setas não podem cruzar o texto nem a data** (DEC-013): descrição no topo ancorada ao topo
   (`TextFrame.VerticalAnchor=Top`), seta **curta** começando **abaixo** do texto e subindo até a
   bolinha (para marcadores "acima"); para "abaixo", seta curta descendo do eixo, sem tocar a
   descrição. Datas ficam do **lado oposto** da seta.
5. **Título "Timeline"**: `WordWrap=Off` (uma linha) e posicionado **abaixo da faixa verde**
   (não colado/sobreposto).
6. **Não invadir o rodapé** no relatório: a imagem é dimensionada pela função de fit (paisagem).
7. Exportar o PPT inteiro para **PDF** (histórico até o mês) = anexo `P1.1 - Timeline de
   Acompanhamento Macro até {mês} {ano}.pdf`; renderizar **TODAS as slides** em PNG (200 dpi,
   `timeline_p01..pNN.png`, ordem cronológica) para o corpo — **NÃO** apenas a slide do mês (DEC-028).

## No corpo do relatório
A Timeline entra **COMPLETA/ACUMULADA**: uma **página horizontal por mês**, do **1º mês (Jul/2025)
até o mês corrente** — grandes e legíveis (DEC-013/DEC-028). Nunca só a slide do mês. O `09` embute
todas as `timeline_p*.png` em ordem cronológica. Cabeçalho/rodapé paisagem adaptados.

## Coordenadas de referência (pt; slide ~842×595 pt) — usadas no RM13
Eixo em y≈305. Descrições "acima": caixa `Top≈166`, `VerticalAnchor=Top`, largura ≈150; seta
`Top≈262, Height≈33` (curta, abaixo do texto → bolinha). Descrições "abaixo": caixa `Top≈378`; seta
`Top≈311, Height≈24`. Título "Timeline" `Top≈124`. X dos 9 marcadores: de ~105 a ~735 (recuo nas
bordas). Ajustar para o nº de marcos do mês.

## FUTURO (pedido para próximas versões — ainda NÃO aplicado)
- Quando **dois eventos têm a mesma data** (ex.: dois em 15/07, dois em 20/07): **compartilhar a
  mesma bolinha** — um evento para cima, outro para baixo.
- Para o evento de cima, a seta deve **"pular" a área do número da data** (interromper/deslocar o
  início) para **não sobrepor o número**, mantendo o mesmo alinhamento vertical da bolinha.

# Regras de Layout (Word e imagens)

## Área útil e dimensionamento das imagens (DEC-016)
Para **toda** imagem/print:
1. **Recortar bordas brancas** (auto-crop; script `08_recortar_imagens.py`, PIL).
2. **Preservar proporção**.
3. Dimensionar ao **máximo dentro da área útil** — o espaço entre a **linha do cabeçalho** e a
   **linha do rodapé** — sem invadir nenhum dos dois.
4. **Centralizar** horizontalmente; distribuir verticalmente equilibrado; texto **próximo** da
   imagem (sem grandes vazios injustificados).

Função de "fit" (em `09_montar_word.py`): largura = `min(largura_util, altura_util × aspecto)`, onde
aspecto = w/h da imagem recortada. Áreas úteis usadas (polegadas):
- **Retrato**: largura ≈ 6,8 · altura ≈ 8,5.
- **Retrato "página cheia"** (`PH_FULL ≈ 10,1`): para imagens que devem ocupar quase toda a página
  (ex.: **Dashboard**), descendo até perto do risco do rodapé (body útil ≈ 10,3in) — ver DEC-024.
- **Paisagem**: largura ≈ 10,4 · altura ≈ 6,5.
Isso garante largura próxima do risco do cabeçalho **e** que a altura não invade o rodapé.
`fit_img(..., area=(w,h))` aceita uma área sob medida (usado no Dashboard).

## Orientação por seção
- **Paisagem** (seções próprias, com quebra de seção): **Controle de Produtos** e **Timeline**
  (tabelas/gráficos largos).
- **Retrato**: Dashboard, Reuniões, Quadro Financeiro, Base Técnica, Linha do Tempo de Medições, textos.
  A **tabela de Reuniões** é a única cuja orientação é **decidida por medição** a cada mês (retrato ou
  paisagem, a de menor desperdício — DEC-027) e cuja largura é **exatamente** a largura útil do corpo
  (6,891" retrato / 10,4" paisagem), sem redimensionamento por altura no Word.
- Um print grande por página; a última página de uma tabela não deve ficar quase vazia (ver
  paginação balanceada em `regras-excel.md`).

## Quebras de seção (paisagem no meio do documento)
Em python-docx: inserir um parágrafo com `sectPr` (cópia do sectPr do corpo, alterando `pgSz`
orient/w/h e `pgMar`) para marcar o **fim** da seção anterior; o conteúdo entre marcadores fica na
seção da orientação desejada. Manter as referências de header/footer (herdadas) e remover `titlePg`.

## Cabeçalho/rodapé das páginas em paisagem (DEC-015; script 11)
Dar às seções paisagem header/footer **próprios**: desvincular (`is_linked_to_previous=False`),
**limpar todo o conteúdo do `<w:hdr>`/`<w:ftr>`** (não só parágrafos — também tabelas/logo, senão
**duplica**), e recriar: header = "Programa | Consórcio do Programa" à esquerda
com borda inferior full-width; footer = "-CONFIDENCIAL-" à esquerda + nº de página (campo PAGE) à
direita (tab no fim da largura útil) com borda superior. **Não** alterar o header/footer das páginas
retrato. Detectar seções paisagem por `orientation==LANDSCAPE` ou `page_width>page_height`.

## Dashboard sozinho na página (DEC-012 + DEC-024)
Inserir **quebra de página** depois de todo o texto do Dashboard (a imagem inicia em nova página).
A imagem do Dashboard usa a **área "página cheia"** (`area=(PW, PH_FULL)`), ocupando quase toda a
área útil — base próxima do risco do rodapé, recuos laterais mínimos, proporção preservada (DEC-024).
Para a imagem ficar sozinha **sem** gerar página em branco, a seção seguinte (título da Timeline)
começa em página nova via **`page_break_before`** — **não** inserir um parágrafo de quebra vazio após
o Dashboard (com a imagem quase preenchendo a página, o parágrafo vazio criaria uma página em branco).

## Títulos de campo
List Bullet, sem negrito, `space_before ≈ 10pt`, `space_after ≈ 4pt` (linha em branco acima).

## Fontes
Capa e Timeline usam fontes Canva (Garet/Open Sauce) não instaladas → Office substitui. Aceitável;
conferência visual (DEC-017).

## Finalização (Word COM)
Abrir no Word, `Fields.Update()` em todas as StoryRanges + `TablesOfContents.Update()` (sumário),
`Repaginate()`, exportar PDF (`SaveAs2(..., FileFormat=17)`). Sempre fechar/`Quit()` o Word (usar
`atexit`); ao final, matar processos órfãos EXCEL/WINWORD/POWERPNT (evita locks).

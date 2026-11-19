# Medição / Quadro Financeiro

## Fonte (DEC-034 — atualizado)
**FONTE OFICIAL DOS NÚMEROS = `…\Desenvolvimento de Produtos\Medicao Parceiro\Planilha de Medição -
Estruturadora - {Mês}.xlsx`** (todo mês, com o nome do mês no arquivo). Abas:
- `Planilha de Medição - Global` = medições **1ª–6ª** (pré-aditivo; blocos de 4 sub-colunas: %mês,
  %acum, valor mês, valor acumulado).
- `Planilha de Medição - SUPRIMIDA` = medições **7ª–12ª…** (pós-aditivo; blocos de 7 sub-colunas por
  medição: %mês, %acum, quant medida, nota avaliativa, valor final, **valor mês**, **valor acumulado**).
  O bloco da medição N é achado pelo cabeçalho "Nª MEDIÇÃO" na linha 2; produtos com ID em col D
  (1.1..5.4), empresas nas linhas 33–36 (Dynatest, Escritório Jurídico, Certificadora, Consultoria), TOTAL na 37.
- `ESCOPO DO SERVIÇO` (referência).
**Todo mês: procurar nessa pasta o arquivo do respectivo mês.** Se faltar/estiver desatualizado,
**perguntar ao usuário** antes de gerar. **NUNCA** pegar do espelho `5. Compartilhado Parceiro` nem de
mês anterior como dado (só como template de formatação — ver abaixo). Nunca inventar valor (DEC-037).

## Painel Quadro Financeiro = TEMPLATE do mês anterior + dados da Planilha (DEC-034)
O relatório usa um **painel formatado** `P1.1 - Quadro Financeiro do mês de {mês}.xlsx` (abas
`PAINEL_FINANCEIRO_*`, `RESUMO_EMPRESAS_*`, `STATUS_MEDICOES_1A*`, `MEDICAO_*`, `BASE_PARCEIRO`,
`MAPEAMENTO_EMPRESAS`, `LEGENDAS_E_OBS` + **gráfico de barras** por empresa). Esse painel é **todo de
valores fixos (sem fórmulas)**; muda ~240 células/mês. **Processo (produzir do zero a cada mês):**
1. **Template = o painel do mês anterior** (ex.: julho copia o de junho, em `4. Produtos\{mês
   anterior}\G1\Anexos\P1.1 …`). Copiar o arquivo (preserva formatação + gráfico) via `shutil`;
   editar células via **Excel COM** (win32) para preservar o gráfico.
2. **Preencher com os dados da Planilha de Medição oficial** do mês (bloco da medição N). Regras
   **validadas 100% contra o painel correto** (RM13/julho):
   - **Por produto** (BASE_PARCEIRO, linhas 4–23 por ID): `% Acum`←`%acum`(bloco+1); `Valor Final`←`valor
     final`(bloco+4); `Valor Acum`←`valor acumulado`(bloco+6). *(Valor mês/unit/contrato conforme a
     Planilha; contrato só muda em aditivo.)*
   - **Por empresa** (RESUMO/PAINEL/MAPEAMENTO/MEDICAO): valor mês←linha da empresa (bloco+5);
     **acumulado da empresa = SOMA do `valor acumulado` dos produtos daquela empresa** (mapa
     produto→empresa na col K do BASE_PARCEIRO, estático); **% = acum/contrato**; **saldo = contrato −
     acum** (contrato estático do template).
   - **Totais** = linha 37 da Planilha (mês, acum) e somas; **saldo total = contrato total − acum total**.
   - **Gráfico** (PAINEL D84:E87 etc.) = cópias de acum/saldo por empresa.
   - **STATUS_MEDICOES**: acrescentar a coluna da medição N (valores por empresa + status); status do
     mês = `status_medicao` do config; medições anteriores conforme `MEDICOES_STATUS` (dados_mes).
   - **Textos**: "{N}ª Medição", "Referência: {N}ª Medição – {mês}/{ano}", "PLANILHA DE MEDIÇÃO – {N}".
   - **"Qntd. Acum." (BASE_PARCEIRO col H)** = quantidade acumulada com ajuste de nota ao longo das
     medições (NÃO é valor direto da Planilha). **Recorrência:** `Qntd.Acum.(mês) = Qntd.Acum.(mês
     anterior, do template) + incremento do mês`. Julho/2026 foi o mês de transição (o painel de junho
     ainda usava "Qntd." mensal); a partir de julho o template já traz o acumulado.
3. Salvar o painel em `…\{mês}\G1\Anexos\P1.1 …` e então o `05`/`06` geram as imagens dele.
Script: `15_gerar_painel_financeiro.py` (a construir) — valida os totais (Σ empresas = total; Σ
produtos por empresa = acum empresa; saldo = contrato − acum) e, para o mês com painel-oráculo
conhecido, confere célula-a-célula. Abas legadas mantêm nomes "M8"/"08"/"1A8" (fixos, não são o nº do mês).

## Excel oficial ausente — planilha RECONSTRUÍDA do quadro impresso (DEC-062)
Se o Excel da medição N não chegou mas o quadro "Nª MEDIÇÃO" está impresso num documento oficial do Parceiro (ex.: última
página do Relatório de Gestão do Contrato): importar o PDF (DEC-059) →
`python medicao_reconstrucao.py reconstruir --pdf <pdf> --modelo "<Planilha … - {Mês anterior}.xlsx>" --saida
"<…/Planilha de Medição - Estruturadora - {Mês}.xlsx>" --relatorio <05_LOGS/conferencia.json> --propriedade "<origem>"`
(só segue com 0 divergência) → ficha de proveniência em `Fontes Externas` → publicar como arquivo novo em `Medicao Parceiro`.
Status: o do documento; sem status formal ⇒ "Em análise". **Não mencionar a reconstrução no relatório.** Quando o
oficial chegar: `medicao_reconstrucao.py comparar --oficial … --reconstruida … --medicao N`.

## Número/competência/status da medição
- **Report N ↔ medição (N-1).** RM12 (junho) = 11ª medição; **RM13 (julho) = 12ª medição**; **RM16 (outubro) = 15ª**.
- **Nomear pela competência de EXECUÇÃO (DEC-063):** a medição N-1 analisada no RM N é a do **mês anterior**
  (`cfg.competencia_medicao()`): "Referência: 15ª Medição – Setembro/2026" no RM16. O `15` preenche também a tabela
  detalhada da aba `MEDICAO_*` (por produto, subtotais, TOTAL GERAL) — antes ficava com dados da 12ª.
- **Status** vem do Quadro (RESUMO/STATUS): ex. RM13 = **"Em análise"** (pela Órgão Contratante). Preencher
  `status_medicao` no config; se ausente e indeterminável, perguntar.
- Título da seção: *"Análise do processo de medição do mês anterior - {mês}/{ano} - {N-1}ª medição"*.

## Validação matemática (obrigatória)
- Σ(valor mês por empresa) = valor total da medição.
- Σ(acumulado por empresa) = acumulado total.
- saldo = contrato − acumulado; % = acumulado / contrato.
- **% Executado da linha TOTAL** também = acumulado total ÷ contrato total (não herdar do template — DEC-054);
  o `18` reprova se `J54..J58` ≠ F/H ou se `J58` ≠ `J8`.
- Totais do texto = totais dos quadros. Divergência real vs Parceiro → registrar no QA e avisar (não
  escolher em silêncio).

## Imagens (script 05) e Linha do Tempo (script 06)
- Trabalhar em **cópia** do Quadro. Exportar como imagem (portrait, ~200 dpi, recortar): **Quadro
  Financeiro Consolidado**, **Resumo por Empresa**, **Base Técnica Parceiro**. **Preservar cores/identidade**
  (não recriar com outra paleta — DEC-014).
- **Quadro Financeiro Consolidado — área de impressão (DEC-029):** o `05` **detecta por conteúdo** o
  fim da tabela **"RESUMO FINANCEIRO POR EMPRESA"** (acha o título e anda até a última linha contígua =
  linha do **TOTAL**) e recorta o print ali. A aba `PAINEL_FINANCEIRO_*` tem, mais abaixo, uma **tabela
  auxiliar (rascunho)** que alimenta gráficos e que **NÃO** pode aparecer no relatório. Detecção sem
  linha fixa (a posição muda a cada mês).
- **Quadro Financeiro Consolidado — largura no corpo (DEC-029):** inserido com **largura EXATA = largura
  do corpo de texto justificado** (retrato: `page_w − margens`, ≈6,891"), casando início/fim com o
  texto. O `09` mede a largura do corpo no próprio documento e só reduz se, nessa largura, a altura
  passar da área útil da página (não estoura). Vale para esta imagem; as demais seguem o fit padrão.
- **Base Técnica Parceiro — divisão em 2 (DEC-031):** a tabela costuma ser alta demais para caber inteira
  logo abaixo do texto de avaliação e, por pouco, é empurrada para a página seguinte, deixando um
  **branco grande**. O `05` **mede a altura cumulativa de cada produto** (pela coluna ID no próprio
  print) e divide a Base em **duas imagens**: a **parte 1** leva o máximo de produtos que cabem no alvo
  `PROG_BASE_PART1_IN` (default **5,25"**) para preencher o resto da página; a **parte 2** leva o resto.
  **Título + cabeçalho de coluna são repetidos** nas duas partes (via `PrintTitleRows`; detecção do
  título/cabeçalho por conteúdo). O `09` insere `base_parceiro_p1/p2` em sequência. Se o alvo não gerar uma
  divisão válida (0<n1<total), sai **imagem única**. Ajuste o alvo se um mês deixar branco/estouro.
- **Linha do Tempo de Medições (DEC-030)**: **3 quadros — medições 1-5, 6-10, 11-12** (blocos de 5). O
  `06` **lê os dados direto da aba `STATUS_MEDICOES` do Quadro Financeiro** (valores por empresa e
  status; status = o **mais avançado** entre as colunas da medição — Pago>Processado>Em análise —
  reproduzindo o baseline; fallback: `cfg.dados()`). Renderiza **mantendo as legendas** e as **mesmas
  cores**: Pago=verde `#A9D08E`, Processado=amarelo `#FFE699`, Em análise=laranja `#F4B183`, cabeçalho
  verde-escuro `#1F7A54`. Entram **após** o texto "Por fim, apresenta-se o acompanhamento financeiro…".
  No corpo, **ampliar cada quadro à largura útil**.

## Texto (fiel ao RM11; negrito seletivo)
Intro idêntica ao RM11 (lógica formal de avaliação por produto). Parágrafo de valores com **negrito
apenas** em: "{N-1}ª medição", valores R$, %. Avaliação: linguagem prudente + status "em análise pela
Órgão Contratante" quando aplicável; citar Produto 1.3 sem medição (supressão) se for o caso. Fechar com "Por
fim, apresenta-se o acompanhamento financeiro das medições do Programa até o mês anterior:" e as imagens.
Intercalar: texto → quadro consolidado → texto(valores) → resumo → texto(avaliação) → base técnica →
texto(plano de ação) → "Por fim…" → 3 quadros da linha do tempo.

## Baseline (RM13 / 12ª medição, ref. Julho/2026) — recalcular a cada mês
Valores FICTÍCIOS (versão pública): contrato **R$ 20.000.000,00**; medição do mês **R$ 500.000,00**; acumulado
**R$ 6.000.000,00**; **30,00%**; saldo **R$ 14.000.000,00**; status **Em análise**. Por empresa (mês): DYNA
200.000,00; CG 120.000,00; JURIDICO 100.000,00; CERT 80.000,00.

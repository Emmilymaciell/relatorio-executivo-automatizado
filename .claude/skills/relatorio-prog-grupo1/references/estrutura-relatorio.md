# Estrutura do Relatório (ordem e componentes)

Segue o padrão RM11/RM12. Base = documento da CG (`...RevCG.docx`).

1. **CAPA** — gráfica, da CG (objeto PowerPoint embutido: logos Programa/Dynatest/Escritório Jurídico/CG/CERT,
   "RELATÓRIO Programa", "Grupo 1 – Coordenação Geral Executiva", "Relatório Mensal N° {n}",
   "Referente ao período de execução de {dd.mm.aa} até {dd.mm.aa}", data). **Não recriar** (DEC-004).
   Sem cabeçalho/rodapé na capa. **Data da capa = último dia do mês da competência** (DEC-058; `capa.py` confere e
   corrige no `09`; o `13` reprova se divergir). O encerramento usa a mesma data.
2. **SUMÁRIO** — campo TOC (Word). Atualizar no fim. "ENCERRAMENTO" e "ANEXOS" **sem número** (DEC-007).
3. **1 APRESENTAÇÃO** — texto institucional fixo; parágrafo variável já traz "Relatório Mensal N° {n},
   referente ao período de 01 de {Mês}/{ano} até 31 de {Mês}/{ano}". Conferir número/mês.
4. **2 COORDENAÇÃO GERAL EXECUTIVA**
   - **2.1 Produto 1.1 – Relatório de Coordenação Geral**
     - 2.1.1 Descrição das atividades atribuídas (texto fixo do TR — manter).
     - 2.1.2 Produtos desenvolvidos no período — **nosso conteúdo**, nesta ordem (títulos bullet,
       sem negrito, linha em branco acima — DEC-008):
       1. *Planilha de Controle de Produtos referente ao mês de {mês}* + prints (landscape).
       2. *Dashboard Analítico das Entregas Executadas em {mês}* + imagem (sozinha na página).
       3. *Linha do Tempo do Progresso das Atividades Executadas em {mês} de {ano}* + imagens (landscape).
       4. *Reuniões de Governança realizadas durante o mês de {mês}* + prints (portrait).
       5. *Análise do processo de medição do mês anterior - {mês da EXECUÇÃO}/{ano} - {N}ª medição* + quadros (DEC-063).
       6. *Síntese do Plano de Gestão Integrada de Riscos - {mês}/{ano}* + tabela dos riscos críticos (DEC-065).
     - Insumo ainda não recebido ⇒ quadro "SEÇÃO EM ELABORAÇÃO — …" no lugar do conteúdo (`dados_mes.PENDENCIAS`) e
       arquivo "… - EM ELABORAÇÃO" (DEC-064).
   - **2.2 Produto 1.2 – Relatório de Relações Institucionais**
     - 2.2.1 Descrição das atividades atribuídas (texto fixo do TR).
     - 2.2.2 Produtos desenvolvidos no período — 1 subseção: *Integração Institucional e
       Consolidação de Processos executados em {mês}* + print (aba Stakehold) **+ texto consolidado
       "Relatório de Relacionamento Institucional do mês de {mês} de {ano}" APÓS a tabela**
       (obrigatório; redigido do modelo + planilha de reuniões 1.2 — DEC-033; ver `regras-redacao.md`).
   - **2.3 Produto 1.3 – Relatório de Comunicação Integrada** — conteúdo **da CG** (preservar).
   - **2.4 Produto 1.4 – Relatório de Tecnologia de Informações de Comunicações (TIC)** — **da CG** (preservar).
5. **ENCERRAMENTO** — "O Relatório Mensal N° {n}, referente à consolidação e à apresentação das
   atividades desenvolvidas durante o mês de {mês} de {ano}, referente ao período de 01 de {mês} de
   {ano} até 31 de {mês} de {ano}, em conformidade ao Contrato N° 000/0000-Parceiro." (negrito seletivo,
   DEC-005) → "Cidade, {data por extenso}" (uma vez) → **duas assinaturas lado a lado** (DEC-006).
6. **ANEXOS** — título sem numeração + **Quadro de Anexos** (tabela LISTA DE DOCUMENTOS ANEXOS). Ver `anexos.md`.

## Cabeçalho/rodapé
- Retrato: header "Programa | Consórcio do Programa"; footer "-CONFIDENCIAL-" +
  nº de página. (Vêm da base; **não alterar**.)
- Paisagem: versão adaptada full-width (DEC-015; script 11).
- Capa: sem cabeçalho/rodapé.

## P1.3/P1.4 — revisão cruzada
Preservar textos/figuras/tabelas. Só corrigir referências antigas evidentes (mês/N° anteriores) se
houver suporte no material do mês. Não inventar atualização da CG. (No RM13, P1.3/P1.4 já vieram
corretos para julho — nenhuma correção necessária.)

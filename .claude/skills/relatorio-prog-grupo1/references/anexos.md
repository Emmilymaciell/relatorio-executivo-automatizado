# Anexos

## Layout da pasta de anexos (DEC-043 — POP 01/10/2026)
- **Até Setembro/2026 (`por_produto`)**: `G1\Anexos\{P1.1 …, P1.2 …, P1.3 …, P1.4 …}\arquivo` (regras abaixo).
- **A partir de Outubro/2026 (`plano`)**: `G1\ANEXOS\arquivo` (sem subpastas). Nome dos anexos gerados por nós:
  `P1.1 - Gestão Executiva - Controle de Entrega de Produtos de {mês} - GERAL.xlsx`,
  `P1.1 - Gestão Executiva - Dashboard da Análise de Entrega de Produtos - {mês} - GERAL.pdf` (nome a usar ao
  exportar o Power BI), `P1.1 - Gestão Executiva - Quadro Financeiro do mês de {mês} - GERAL.xlsx`,
  `P1.1 - Gestão Executiva - Timeline de Acompanhamento Macro até {mês} {ano} - GERAL.pdf`,
  `P1.2 - Relações Institucionais - Controle de Reuniões do mês de {mês} - GERAL.xlsx`.
  P1.3/P1.4: a CG envia já no padrão do POP; a automação do G1 continua **sem** renomear/mover.
- Detecção automática por `cfg.layout_anexos()`; forçar com `"layout_anexos"` no `parametros.json`.
- No **pacote final** (Skill `entrega-coordenacao`) cada não-PDF vira ZIP individual e nomes fora do padrão
  com proposta completa são padronizados — sempre com o original arquivado (DEC-045).

## Regra (DEC-007)
- Título **"ANEXOS"** como heading **sem numeração** (suprimir a numeração automática do Heading 1 —
  `numPr` com `numId=0`) para o sumário mostrar "ANEXOS" (não "3 ANEXOS").
- **Não** criar a lista do zero: replicar o padrão do **RM11 (Word)** → **Quadro de Anexos**:
  tabela "LISTA DE DOCUMENTOS ANEXOS" (cabeçalho verde `#1F7A54` mesclado 2 colunas) com sub-cabeçalho
  **REF. | DESCRIÇÃO** (verde `#2E8B62`), depois uma linha por anexo. Col REF ≈ 1,3"; DESCRIÇÃO ≈ 5,2".
- **Listar TODOS os arquivos que existem fisicamente** em `G1\Anexos\{P1.1,P1.2,P1.3,P1.4}`, com o
  **nome EXATO** que está na pasta. Ordem: P1.1, P1.2, P1.3, P1.4 (e por nome dentro da pasta).
  Numerar Anexo I, II, III… (romanos) conforme a **quantidade real** de arquivos. **DEC-042.**
- **NUNCA agregar/zipar** os anexos, e **NUNCA mexer** nos anexos de P1.3/P1.4 (nem renomear/mover/
  apagar) — deixar exatamente como os grupos enviam; só a usuária, manualmente, altera esses anexos.
- A geração é automática: o `10` chama `cfg.listar_anexos_reais()` (varre as 4 pastas, ignora
  temporários/`~$`/`Thumbs.db` e o zip legado `Evidências 1.X …`) e monta uma linha por arquivo.
  Não listar PDFs/PNGs temporários criados só para o corpo (esses ficam no work, não no anexo).

## Anexos recorrentes do mês (P1.1/P1.2 — nomes que a automação gera)
- `P1.1 - Controle de Entrega de Produtos de {mês}.xlsx`
- `P1.1 - Dashboard da Análise de Entrega de Produtos - {mês}.pdf`
- `P1.1 - Quadro Financeiro do mês de {mês}.xlsx`
- `P1.1 - Timeline de Acompanhamento Macro até {mês} {ano}.pdf`
- `P1.2 - Controle de Reuniões do mês de {mês}.xlsx`
- **P1.3 e P1.4**: os arquivos **individuais** que os grupos subiram (quantidade e nomes variam por mês)
  — listar todos, um a um, pelo nome real. **Sem zip** (DEC-042).

## Onde ficam
`…\4. Produtos\{pasta mensal}\G1\Anexos\P1.1 - Gestao Executiva\` (Controle, Dashboard, Quadro
Financeiro, Timeline), `…\P1.2 - Relacoes Institucionais\` (Reuniões), `…\P1.3 …\` e `…\P1.4 …\`
(evidências individuais, como os grupos enviam). Descobrir as subpastas por busca (nomes variam por mês).

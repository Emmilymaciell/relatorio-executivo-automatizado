# Regras antigas do repositório × POP v2026-10-01

Levantamento de 01/10/2026. **O POP prevalece**; nada foi silenciado. Coluna "Situação" = o que o sistema faz hoje.

| # | Regra antiga (onde) | O que o POP/Checklist diz | Conflito? | Situação / decisão |
|---|---|---|---|---|
| 1 | Anexos do G1 em `G1\Anexos\P1.1…P1.4` (subpastas por produto) — `anexos.md`, `prog_config` | ANEXOS **única**, sem subpastas; produto no **nome** [Quadro 2/3] | SIM | DEC-043: layout detectado por mês; `plano` a partir de Out/2026 (pastas 16–18 já reorganizadas — DEC-047). Meses antigos inalterados. |
| 2 | "NUNCA zipar nem mexer em P1.3/P1.4" — DEC-042 | Todo não-PDF em **ZIP**; nomes padronizados; o pacote não pode exigir correção pelo Parceiro | PARCIAL | DEC-043: a automação do G1 continua sem tocar P1.3/P1.4; o ZIP individual e o nome padrão são aplicados **só no pacote** pela Coordenação, com original arquivado. O ZIP agregado "Evidências 1.X" segue proibido. |
| 3 | Nomes dos anexos gerados `P1.1 - Quadro Financeiro do mês de agosto.xlsx` | `[Código] - [Produto] - [Arquivo] - [Rodovia]` | SIM | `cfg.nome_anexo_g1()` → `P1.1 - Gestão Executiva - … - GERAL.xlsx` no layout `plano`. |
| 4 | Planilhas de anexo (.xlsx) entregues soltas | não-PDF em ZIP + "editável deve corresponder ao PDF quando aplicável" | SIM | Pacote zipa cada planilha (bytes idênticos dentro do ZIP). Falta de PDF correspondente = WARNING. **Pendente:** o G1 ainda não gera o PDF das suas planilhas (Controle, Quadro, Reuniões) — ver STATE. |
| 5 | "Relatório Mensal" (nome do arquivo/encerramento) | "Relatório de Produtos Mensal" [Checklist] + pedido da usuária | SIM | DEC-044 (`termo_documento`). Tipo "Relatório Mensal" no acumulado mantido (Power BI). |
| 6 | Saída do G1: DOCX+PDF versionados (`_v2`, `_v3`) na raiz do `G1` — `12_finalizar_pdf.py` | Na raiz do grupo **só** o relatório final assinado | SIM (no pacote) | Mantido durante a produção (é a pasta de trabalho); no pacote a Coordenação retira Word e versões anteriores (arquivo interno) — DEC-045. |
| 7 | Lista de Demandas fica na pasta do mês (fonte do `01`) | Lista é interna e **sai** do pacote | SIM (no pacote) | Publicação move para o arquivo interno; `cfg.lista_demandas_path()` procura lá depois. |
| 8 | Índice: não existia | Índice cumulativo na raiz do mês [Quadro 2 + Modelo] | NOVO | `prog_entrega/indice.py` (cria modelo do mês; regenera e reconcilia no pacote). |
| 9 | PLANO_ajuste_entrega_SEI (27/09): pasta `Grupo 1` plana **sem** ANEXOS; índice por grupo; nome `P1. Programa Relatório…` | ANEXOS existe; índice **único** na raiz do mês; não há regra de nome do relatório | SIM | PLANO superado (arquivado em `docs/historico/`). Respostas do POP às Q1–Q6: Q1 PDF+ZIP; Q2 pasta do mês; Q4 SEI em branco; Q6 um ZIP por arquivo. |
| 10 | PROMPT antigo: "ICP-Brasil na prestação de contas" | POP: "assinaturas digitais… quando exigíveis" (sem ICP) | DIVERGÊNCIA DE FONTE | DEC-046: `exigir_icp_brasil=false` até decisão da usuária; emissor sempre reportado. |
| 11 | PROMPT antigo: datas no nome `2026-08-31`; "- Assinado" | POP só exemplifica data no nome (P1.4) | NÃO | Não exigido; não validado. |
| 12 | Corte "26 do mês anterior → 25" (citado na missão) | POP não define corte | — | Não existe no repo; vigente = mês civil (DEC-048). |
| 13 | `5. Compartilhado Parceiro` = espelho de entrega, nunca fonte (DEC-037) | POP: a Coordenação "encaminha ao Parceiro" | NÃO | Mantido. A automação nunca escreve lá; a usuária copia a pasta do mês. |
| 14 | Segurança SharePoint: só o acumulado pode ser alterado (CLAUDE.md regra 1) | POP Etapa 6: Coordenação **remove** itens da pasta final | SIM | Exceção controlada (DEC-045/047): estrutura/publicação só com autorização por escrito, arquivamento prévio com hash, log antes→depois; acumulado continua só-acrescentar. |

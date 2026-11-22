# STATE — Entrega de Produtos - Coordenação (Programa)

> **Versão pública (sanitizada):** `docs/sessoes/`, `docs/competencias/`, `docs/historico/` e `docs/divulgacao/` foram removidos (registros operacionais do contrato). Referências a eles abaixo ficam sem destino.


> Snapshot operacional. Quem chega (pessoa, máquina ou Claude novo) lê isto primeiro, depois `docs/HANDOFF.md`.
> Não é diário: histórico em `docs/sessoes/`.

- **Última atualização:** 2026-11-16 (sessão 005 — RM16/outubro EM ELABORAÇÃO; regras DEC-058…070)
- **Branch:** `master` · **Remoto:** `origin` → GitHub **privado** `prog-relatorio-grupo1`. Push só a pedido. ·
  **Tag vigente:** `v2.0.0`
- **HEAD:** `a283660` (RM16 em elaboração + reorganização DEC-070) + commit de encerramento da sessão 005 (só docs).
  Working tree limpo ao fechar a sessão 005. Sem push.
- **Ambiente:** segredo do SharePoint em `%USERPROFILE%\.prog\sharepoint.env`. **OneDrive desta máquina não hidrata
  arquivos "somente online" (16/11/2026)** → trabalhar com o espelho de leitura `C:\caminho\exemplo\PROG_ESPELHO`
  (`sharepoint.py espelhar`) e `PROG_ROOT` apontando para ele; publicar pela API (DEC-066).

## Objetivo
Produzir o Relatório de Produtos Mensal do G1 e preparar/validar o pacote mensal G1–G5 pronto para protocolo no Parceiro,
conforme o POP de Entrega v2026-10-01 (minuta). Nunca enviar nada ao Parceiro.

## Fase
**Outubro/2026 (RM16) EM ELABORAÇÃO** — estado e retomada em **`docs/competencias/RM16-2026-10.md`** (fonte única).
- CONCLUÍDO: Controle de Demandas **217** (100% do declarado; DEC-067 revogou o corte temporal), base do Power BI,
  Timeline (12 eventos/9 marcos, Obra × Estruturadora, RD-102A), Quadro Financeiro da 15ª com o **Excel oficial**
  (idêntico à reconstruída), síntese de riscos, anexos P1.1 (3), relatório EM ELABORAÇÃO (assinante Pessoa D).
- CONCLUÍDO (DEC-070): pasta `16. Outubro 2026` reorganizada no formato do POP para TODOS os grupos (cópia "como
  enviado" antes, em `3. Desenvolvimento/G1 - Coordenacao/Entregas Recebidas/`), 4 RMs do G2 incluídos (G2 144/144),
  Índice completo G1–G5 (245 linhas). Auditoria: 1 BLOCKER (relatório G1 final) + 12 MANUAL (OCR, Lista do G2).
- PENDENTE: reuniões de outubro; PDF do Dashboard; assinatura da Pessoa D; auditoria completa G1–G5 (depois do relatório).
- **Setembro/2026 (RM15):** publicado e READY FOR PARCEIRO (código `d6c9b8fb6c`).

## Concluído (verificado)
- Regras permanentes desta sessão: DEC-058 (data da capa = fim da competência), DEC-059 (preservação de insumos
  externos), DEC-060 (Prestações de Contas iterativas na Timeline), DEC-061 (corte temporal — determinação II do Parceiro),
  DEC-062 (planilha de medição reconstruída do quadro impresso), DEC-063 (medição pela competência de execução + aba
  detalhada — determinação I), DEC-064 (versão EM ELABORAÇÃO), DEC-065 (síntese de riscos no P1.1 — determinação III),
  DEC-066 (ZIP dos anexos da Coordenação; espelho de leitura); DEC-061 **revogada** pela DEC-067 (Controle = tudo o
  que foi declarado); DEC-068 (Prestação de Contas de Obra × Estruturadora com evidência do SharePoint); DEC-069 (Pessoa D,
  G4 aceito, auditoria depois).
- DEC-070 (reorganização da pasta do mês no formato do POP, com cópia "como enviado" antes; API ganhou mover/copiar,
  nunca excluir).
- Última validação: `python -m pytest tests -q -p no:cacheprovider` → **103 passed** (2026-11-16).

## Parcial / pendente (fora do RM16)
1. **RM14 (agosto):** painel com % do TOTAL errado; correção só em `output/RM14_2026_08/CORRECAO_TOTAL_PCT/` (decisão).
2. **Reuniões de agosto/2026** ausentes na mestre.
3. **Determinações do Parceiro (Relatório de Gestão do Contrato – set/2026)** a comprovar em outubro: (I) ✔, (III) ✔;
   (II) corte temporal — **não adotado por decisão da usuária** (DEC-067; tratar a resposta ao Parceiro com ela);
   (IV) P1.4/CG — trilha de aprovação das publicações (cobrar a CG); (V) Controle de Reuniões — fazer na finalização.
4. Observação do Parceiro (não determinação): os 7 anexos históricos do P1.3 contam como "entregues" no Controle e elevam o
   indicador da Coordenação Executiva — decisão da usuária pendente.
5. Organização de `Desenvolvimento de Produtos`, aprovação do POP pela Pessoa A (DEC-050), rename da pasta do projeto.

## Decisões abertas (dependem da usuária)
- **Controle de Demandas × arquivos (rodovia):** NT-835.1 e PA-821.1 estão como RD-102A no Controle (V16 = 217, igual
  à Lista Mestre do G2), mas os documentos dizem RD-101 e RD-103 (arquivos já renomeados, DEC-051). Corrigir as 2 linhas
  no acumulado agora (refaz e republica a V16, com backup) **ou** esperar o G2 corrigir a Lista Mestre? — perguntar.
- ICP-Brasil × assinatura digital (DEC-046). G4 de outubro: assinatura manual **aceita** (DEC-069).
- G2 sem Lista de Demandas na pasta (usa a Lista Mestre) — exceção do RM15 não vale para outubro.
- Tipo "Relatório Mensal" no acumulado/Power BI (DEC-044). Feriados estaduais/municipais.

## Riscos
| ID | Risco | Mitigação |
|---|---|---|
| R1 | `publicar` altera a pasta oficial do mês | código do plano + autorização; arquiva antes com SHA-256 |
| R2 | Caminhos > 250 caracteres (G2/G3/G5) | nomes ≤ 135 no G2; espelho em raiz curta |
| R3 | Cadeia/revogação de certificados não verificada offline | declarado; emissor sempre informado |
| R4 | POP sem aprovação formal | DEC-050: tratar como minuta externamente |
| R5 | Medição | Excel oficial recebido e idêntico à reconstruída (16/11) |
| R6 | OneDrive sem hidratar | espelho de leitura + `comparar-pasta` antes de publicar |
| R7 | Prazos do POP descumpridos (G1/CG, G3, G4, Coordenação) | Relatório de Conformidade para cobrança |

## Entrega mensal — checkpoints
- `output/ENTREGA_16_Outubro_2026/` — auditoria PRELIMINAR (16/11; completa fica para depois do relatório);
  `04_RELATORIOS/RELATORIO_PRONTIDAO.md` e `Relatorio de Conformidade da Entrega - Outubro 2026.{md,docx}`.
- `output/RM16_2026_10/` — trabalho do G1 (01_FONTES/EXTERNAS com proveniência; 05_LOGS com QA, revisão final,
  conferência da medição, manifesto de publicação).
- `output/ENTREGA_15_Setembro_2026/` — PUBLICADO (RM15).

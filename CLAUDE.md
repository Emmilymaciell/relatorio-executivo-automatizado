# Programa — Entrega de Produtos - Coordenação

> **Versão pública (sanitizada):** `docs/sessoes/`, `docs/competencias/`, `docs/historico/` e `docs/divulgacao/` foram removidos (registros operacionais do contrato). Referências a eles abaixo ficam sem destino.


Projeto da **Coordenação Geral Executiva do Programa** (Consórcio do Programa: Dynatest,
Certificadora, Consultoria de Gestão, Escritório Jurídico). Escopo (ampliado em 01/10/2026 — evolução do antigo
"Automação do Relatório Mensal Grupo 1", mesmo repositório):
1. **produzir** todo mês o **Relatório de Produtos Mensal do Grupo 1** (P1.1/P1.2) e manter acumulados/anexos;
2. **conferir e preparar** a entrega mensal de G1–G5 ao Parceiro conforme o **POP de Entrega de Produtos** —
   pacote pronto para protocolo, com evidência verificável. **Nunca** enviar/protocolar/publicar no Parceiro.

## Sessões (a conversa é descartável; o repositório é a memória)
- Abrir: **`/session-start`** (só lê e reconstrói; para e aguarda). Encerrar: **`/session-close`** (persiste e para).
- Estado: `docs/STATE.md` · handoff curto: `docs/HANDOFF.md` · histórico: `docs/sessoes/` · uso: `docs/COMO-USAR.md`.

## Qual Skill usar
| Pedido | Skill |
|---|---|
| "Crie o Relatório (de Produtos) Mensal nº N de <mês>" | `relatorio-prog-grupo1` |
| "Audite/confira a entrega", "prepare o pacote do Parceiro", "está pronto para protocolo?" | `entrega-coordenacao` |
| Salvar versão no Git (só a pedido) | `versionar` |

## Fontes canônicas do processo de entrega
`docs/referencias-oficiais/v2026-10-01/` (POP, Checklist, Fluxograma) + `v2026-10-01-r2/` (Modelo de Índice vigente:
"ÍNDICE DE ARQUIVOS : ENTREGA DE PRODUTOS DO PROGRAMA", só por grupo, numeração 1…4 reiniciando em cada grupo — DEC-049)
— **read-only**; registro e SHA-256 no README. Hierarquia: POP > Checklist > Modelo de Índice > Fluxograma > regras do repo > docs antigas >
outputs > inferência. Resumo do domínio: `docs/dominio/processo-entrega-mensal.md`. Divergências antigas×POP:
`docs/dominio/regras-pop-x-antigas.md`. Decisões (G1 e entrega): `.claude/skills/relatorio-prog-grupo1/references/decisoes-e-erros-conhecidos.md`.

## Divisão do Grupo 1
- **P1.1 – Gestão Executiva / Coordenação Geral** (Dynatest) e **P1.2 – Relações Institucionais** (Dynatest) →
  produzidos pela Coordenação no 4º/5º dia útil (automação deste repo).
- **P1.3 – Comunicação Integrada** e **P1.4 – TIC** (CG) → entregues pela CG até o 3º dia útil; **preservar**.
Demais grupos: G2 Dynatest (engenharia), G3 Certificadora (certificação), G4 CG (portfólio/financeiro/riscos),
G5 Escritório Jurídico (jurídico) — entregam até o 3º dia útil.

## Regras permanentes (valem em toda sessão)
1. **Segurança SharePoint**: NÃO excluir, mover, renomear ou sobrescrever originais. Exceções, **sempre com
   autorização explícita da usuária na conversa**: (a) acumulado `P.1.1 - Controle de demandas Programa.xlsx` — só
   ACRESCENTAR linhas, com backup; (b) `scripts/entrega.py estrutura/indice --aplicar` (só cria, padroniza e remove
   subpasta **vazia**); (c) `scripts/entrega.py publicar --confirmar <código>` (arquiva tudo antes, com SHA-256, e
   reconfere contra o manifesto — DEC-045); (d) `scripts/manutencao/reorganizar_mes_pop.py` (cópia "como enviado" no
   Desenvolvimento ANTES; renomeia/move no padrão do POP; retira do pacote movendo, nunca apaga — DEC-070). Nunca
   publicar Power BI. Nunca enviar e-mail. Nunca assinar. Nunca escrever em `5. Compartilhado Parceiro`. **API Graph** (`scripts/sharepoint.py`, DEC-052) = outro canal, **mesmas
   regras**: leitura livre; `enviar` só com autorização, arquivo novo por padrão, substituir só com `--substituir`
   (backup + eTag + SHA-256); sem excluir; mover/renomear só pela exceção (d). Segredo **nunca** no repo/chat/log
   (`%USERPROFILE%\.prog\sharepoint.env`).
2. **Nunca inventar dados** (entrega, marco, reunião, valor, %, risco, nº SEI, assinatura, rodovia, versão vigente).
   **G1: produzir do zero das FONTES OFICIAIS** todo mês (DEC-034/035/036/037): medição = `Medicao Parceiro\Planilha
   de Medição - Estruturadora - {Mês}.xlsx`; timeline = `P1.1 -Controle Timeline.xlsx` (se faltar, PERGUNTAR);
   reuniões = filtrar o mês do mestre `P1.2 - Controle de Reuniões.xlsx`; dashboard = PDF exportado pela usuária.
   **NUNCA usar como fonte o espelho `5. Compartilhado Parceiro` nem `Obsoleto`.** Faltou fonte → perguntar.
   Insumo vindo de FORA (Downloads, chat, e-mail): **importar e preservar antes de usar** (DEC-059). Data da capa =
   último dia da competência (DEC-058). Timeline considera Prestações de Contas iterativas do mês (DEC-060).
3. **Preservar layout/identidade visual** (relatório RM11/RM12/baseline RM13 v5; Modelo de Índice aprovado).
4. **Scripts determinísticos** (Python + Office COM) para arquivos; raciocínio para análise, redação e QA.
5. **Progressive disclosure**: leia só a reference/doc da etapa em execução.
6. **QA obrigatório**: G1 = `13_qa` + `18_revisao_final`; entrega = relatório de prontidão do pacote FINAL
   revalidado. "Script terminou" ≠ "validado". READY FOR PARCEIRO só com 0 BLOCKER e 0 MANUAL_REVIEW.
7. **Checkpoint**: pergunte só o que não dá para descobrir (status da medição ausente; export do Power BI;
   decisões de versão/assinatura/destino de arquivo). Resolva o resto e documente.
8. **Parâmetros variáveis em `config/`** (`parametros.json` do mês; `entrega_pop.json` com limite 100 MB, rodovias,
   produtos; `feriados.json`) — nunca fixos no código.
9. **A regra mais recente da usuária prevalece.** Mudança permanente → nova DEC no log + reference/doc
   correspondente (substituir a regra vigente, preservar o histórico).
10. **Designação oficial do documento: "Relatório de Produtos Mensal"** (DEC-044).
11. **Git**: commit/tag só por pedido explícito (protocolo global); **nunca push**.

## Estrutura
- `config/` — parâmetros (G1 e POP), caminhos, feriados. `docs/` — domínio, referências oficiais, arquitetura
  (`docs/ARQUITETURA.md`: camadas, módulos, pipeline, mapa de migração), estado, handoff, sessões, histórico.
- `.claude/skills/` — `relatorio-prog-grupo1` (G1: SKILL + references + scripts 01–19 + `prog_config.py`),
  `entrega-coordenacao`, `session-start`, `session-close`.
- `scripts/entrega.py` + `scripts/prog_entrega/` — pipeline da Coordenação. `tests/` — pytest (fixtures artificiais).
- `scripts/sharepoint.py` + `scripts/prog_sharepoint/` — acesso ao SharePoint do Programa pela Graph (`config/sharepoint.json`).
- `examples/baseline/` — RM13 v5 = **referência VISUAL**, não fonte de dados. `output/`, `logs/` — não versionados.

**Baseline visual do G1:** Relatório nº 13 (Julho/2026), v5 — `examples/baseline-versao5.md`.

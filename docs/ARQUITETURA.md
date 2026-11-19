# Arquitetura — Entrega de Produtos - Coordenação

> **Versão pública (sanitizada):** `docs/sessoes/`, `docs/competencias/`, `docs/historico/` e `docs/divulgacao/` foram removidos (registros operacionais do contrato). Referências a eles abaixo ficam sem destino.


Repositório único (Git local, sem remoto). Evolução do antigo `relatorio-executivo-automatizado` — **não** é projeto
novo: a automação do G1 continua sendo o núcleo do P1.1/P1.2.

## Camadas
| Camada | Onde | Papel |
|---|---|---|
| Regras/conhecimento do domínio | `docs/dominio/`, `docs/referencias-oficiais/` (canônicos, read-only) | o que é o processo e de onde vem cada regra |
| Decisões | `.claude/skills/relatorio-prog-grupo1/references/decisoes-e-erros-conhecidos.md` (DEC-001…052) | log único de decisões/erros (G1 **e** entrega) |
| Configuração | `config/` — `parametros.json` (mês do G1; não versionado), `caminhos.json`, `entrega_pop.json` (regras variáveis do POP), `feriados.json`, `dados_mes.py` (não versionado) | nada de mês/limite fixo no código |
| Geração do G1 (P1.1/P1.2) | Skill `relatorio-prog-grupo1` → `scripts/01…19` + `prog_config.py` | relatório do G1 (inalterado no essencial; ver DEC-043/044) |
| Coordenação / pacote Parceiro | `scripts/entrega.py` (CLI) + `scripts/prog_entrega/` | inventário → validação → reconciliação → staging → final → revalidação → manifesto → relatório → publicação |
| Acesso SharePoint (API) | `scripts/sharepoint.py` (CLI) + `scripts/prog_sharepoint/` + `config/sharepoint.json` | Microsoft Graph app-only: testar/listar/info/baixar/comparar/enviar (dry-run, backup+eTag+SHA-256; sem excluir/mover) — DEC-052 |
| Skills | `.claude/skills/{relatorio-prog-grupo1, entrega-coordenacao, session-start, session-close}` | como o Claude conduz cada fluxo |
| Estado/continuidade | `docs/STATE.md` (snapshot), `docs/HANDOFF.md` (curto), `docs/sessoes/` (histórico), `output/ENTREGA_*/estado.json` (checkpoint da entrega) | a conversa é descartável; o repositório é a memória |
| Testes | `tests/` (pytest, fixtures artificiais) | 71 testes: pipeline, assinatura, nomes, índice, estrutura, compat. G1, SharePoint/Graph (sem rede) |
| Saídas | `output/RM{n}_…` (G1), `output/ENTREGA_NN_Mes_AAAA/` (Coordenação), `logs/entrega/` | não versionadas |

## Módulos `scripts/prog_entrega/`
`regras` (config, status, caminho longo) · `inventario` · `hashcache` (SHA-256 com cache tamanho+mtime; resultados
por conteúdo) · `arquivos` (PDF/ZIP/Office por tipo real) · `assinatura` (CMS/PAdES) · `nomes` (POP Quadro 3 +
propostas) · `indice` (Modelo aprovado) · `validacao` (auditoria ORIGEM/FINAL) · `pacote` (staging/final, ZIP
determinístico, fracionamento) · `relatorio` (MD/JSON/CSV/manifesto + checklist oficial) · `calendario` · `estrutura`
(pastas do SharePoint) · `pipeline` (orquestração + publicação).

## Pipeline da Coordenação
```
ORIGEM (pasta do mês, SharePoint)  --inventário/validação-->  achados ORIGEM + plano
   └─ STAGING  (cópia fiel, SHA-256 conferido)          output/ENTREGA_*/02_STAGING
        └─ FINAL (layout POP: listas/Word fora, ANEXOS plana, ZIP, partes, índice regenerado)  03_FINAL
             └─ REVALIDAÇÃO FINAL (mesmo código; o que "precisaria corrigir" vira BLOCKER)
                  └─ MANIFESTO (SHA-256) + RELATÓRIO DE PRONTIDÃO + plano/código de publicação   04_RELATORIOS
                       └─ PUBLICAR --confirmar <código>: arquiva → remove → grava → reconfere = manifesto
```
Status: PASS · FIXED_AUTOMATICALLY · WARNING · MANUAL_REVIEW · BLOCKER · N/A. Veredito: READY FOR PARCEIRO /
PENDENTE DE REVISÃO HUMANA / BLOQUEADO.

## Mapa de migração ATUAL → NOVO (01/10/2026)
| Antes | Depois | Como |
|---|---|---|
| Escopo "Relatório Grupo 1" | "Entrega de Produtos - Coordenação" (G1 + Coordenação) | CLAUDE.md, docs |
| `PLANO_ajuste_entrega_SEI.md`, `PROMPT_apresentacao_nova_estrutura_SEI.md` (raiz) | `docs/historico/` (marcados como superados) | movidos (histórico preservado) |
| — | `scripts/entrega.py` + `scripts/prog_entrega/` | novo, aditivo |
| — | Skills `entrega-coordenacao`, `session-start`, `session-close` | novas |
| `prog_config` com subpastas P1.x fixas | `layout_anexos()` + `nome_anexo_g1()` + `termo_documento()` | compatível com meses antigos (testes) |
| Sem testes | `tests/` (pytest) | novo |
| Sem estado persistido | `docs/STATE.md`, `docs/HANDOFF.md`, `docs/sessoes/`, `project.yaml` | novo (convenção do plugin `os`) |
| Pasta `relatorio-executivo-automatizado` | `Entrega de Produtos - Coordenação` | **manual**, script `scripts/manutencao/renomear_projeto.ps1` (sessão precisa estar fechada) |

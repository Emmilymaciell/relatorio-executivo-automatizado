---
name: session-close
description: Protocolo oficial de ENCERRAMENTO de sessão do projeto Entrega de Produtos - Coordenação. Para o desenvolvimento novo, reconcilia Git/disco/estado/testes/entrega, persiste no repositório tudo que só existe na conversa (estado, handoff, histórico, decisões), registra validações reais e para. Use quando a usuária digitar /session-close.
disable-model-invocation: true
argument-hint: "[observação final, opcional]"
---

# /session-close — encerramento com persistência

**Pare o desenvolvimento novo.** Não inicie feature, próxima tarefa, refatoração oportunista nem análise extra.
Só é permitido: ler, rodar verificações apropriadas e **escrever os arquivos de continuidade** abaixo.
Nada de commit/tag/push por conta própria (política Git do projeto: só com pedido explícito — se a usuária quiser,
ela pede "salva essa versão" / `/versionar`). Observação da usuária: $ARGUMENTS

## Snapshot
- Git: !`git status --short --branch`
- HEAD: !`git log --oneline -1`

## 1. Pergunta obrigatória
"Existe alguma informação importante que está **só nesta conversa**?" (decisão, descoberta, bug, hipótese, limite,
risco, pendência, autorização dada, arquivo criado/alterado, saída gerada, estado da entrega). Se sim → persistir.

## 2. Onde persistir (separação semântica — não jogar tudo num arquivo)
| Tipo | Arquivo | Regra |
|---|---|---|
| REGRAS permanentes | `CLAUDE.md` / `docs/dominio/*` | só se uma regra mudou (a mais recente da usuária prevalece) |
| DECISÕES | `.claude/skills/relatorio-prog-grupo1/references/decisoes-e-erros-conhecidos.md` | nova DEC-NNN (antes/depois/por quê/evidência); nunca apagar histórico |
| ESTADO (snapshot) | `docs/STATE.md` | reescrever compacto: objetivo, fase, competência, último marco, trabalho atual, próximo passo, blockers, pendências, decisões abertas, branch, HEAD, não-commitados, última validação. **Não é diário.** |
| HANDOFF (curto) | `docs/HANDOFF.md` | reescrever: onde paramos · próxima ação · arquivos a consultar · arquivos que NÃO precisam ser relidos · blockers · riscos · checkpoints |
| HISTÓRICO | `docs/sessoes/SESSAO-NNN-AAAA-MM-DD.md` | novo arquivo: concluído/parcial/investigado, descobertas, arquivos criados/alterados, validações |
| ENTREGA (checkpoint) | `output/ENTREGA_*/estado.json` | gerado pela pipeline; só **citar** no STATE (não editar à mão) |

## 3. Validação — só o apropriado ao que foi feito
- Código em `scripts/` ou `tests/` mudou → `python -m pytest tests -q -p no:cacheprovider`.
- Só documentação → conferir links/caminhos citados (existem?).
- Entrega mensal mexida → citar `estado.json` + `04_RELATORIOS/RELATORIO_PRONTIDAO.md`.
Registrar **comando exato + resultado real**. Não executado = "NÃO EXECUTADO — motivo — risco residual".
Nunca escrever "testes passaram" sem ter rodado nesta sessão.

## 4. Relatório
```
SESSÃO        Número · Nome · Objetivo
REPOSITÓRIO   Diretório · Branch · HEAD · Working tree
CONCLUÍDO / PARCIAL / INVESTIGADO-NÃO IMPLEMENTADO
DESCOBERTAS · DECISÕES · QUESTÕES ABERTAS
VALIDAÇÕES (comando → resultado)
ALTERAÇÕES NÃO COMMITADAS (lista)
DOCUMENTAÇÃO atualizada (arquivos)
ESTADO DA ENTREGA (competência · etapa · veredito)
CHECKPOINT · PRÓXIMO PASSO EXATO · ARQUIVOS DE ALTO SINAL · RISCOS
GIT FINAL (status; "sem commit — aguardando pedido explícito" se for o caso; nunca push)
```
Termine **exatamente** com:

Handoff concluído. Nenhuma nova implementação foi iniciada.

E **PARE**.

---
name: session-start
description: Protocolo oficial de ABERTURA de sessão do projeto Entrega de Produtos - Coordenação. Reconstrói o contexto só pelo repositório (regras, estado, Git, handoff, entrega mensal em andamento) sem alterar nada, mostra divergências e para aguardando autorização. Use quando a usuária digitar /session-start.
disable-model-invocation: true
disallowed-tools: Edit Write NotebookEdit
argument-hint: "[objetivo da sessão, opcional]"
---

# /session-start — reconstrução de contexto (SOMENTE LEITURA)

A conversa é descartável; **o repositório é a memória**. Nesta skill você **não** implementa, edita, refatora,
instala, corrige, commita, faz push, apaga, reseta, faz stash nem inicia a próxima tarefa. Só lê e relata.
Comandos de shell permitidos: apenas leitura (`git status/log/diff/show/rev-parse`, `ls`, `cat`, `python -m pytest
--collect-only` se precisar contar testes). Nunca rode `entrega.py ... --aplicar` nem `publicar`.

Objetivo informado pela usuária (pode estar vazio): $ARGUMENTS

## Snapshot automático
- Git: !`git status --short --branch`
- Últimos commits: !`git log --oneline -8`
- HEAD: !`git rev-parse --short HEAD`

## Sequência (progressive disclosure — não leia o repositório inteiro)
1. **Regras**: `CLAUDE.md` (projeto). Não releia os documentos oficiais; eles estão resumidos em
   `docs/dominio/processo-entrega-mensal.md` — abra só se a próxima tarefa tocar o processo de entrega.
2. **Estado**: `docs/STATE.md` (snapshot) → `docs/HANDOFF.md` (onde paramos + próxima ação + o que não reler).
3. **Histórico**: só o arquivo mais recente de `docs/sessoes/` (nº da última sessão → esta é a seguinte).
4. **Entrega mensal em andamento** (se houver): para cada `output/ENTREGA_*/estado.json`, ler competência, etapa
   (AUDITADO/PREPARADO/PUBLICADO), veredito, contagens por status/grupo, código de publicação, data. Não reprocessar.
   G1: o mês corrente está em `config/parametros.json` (pode não existir numa máquina nova — é normal).
5. **Próxima tarefa**: a registrada no HANDOFF/STATE. Abra **só** os arquivos que ela exigir.

## Reconciliar (mostrar divergências; NUNCA corrigir aqui)
- HEAD/branch registrados no STATE × `git rev-parse`/`git branch`.
- "Alterações não commitadas" do STATE × `git status` (arquivos novos/alterados não explicados = **mostrar**;
  podem ser da usuária — não tocar).
- Próximo passo do STATE já feito? (conferir `git log` / disco).
- STATE mais antigo que o último commit? → trabalho não registrado.
- `estado.json` da entrega × arquivos de `04_RELATORIOS/` (existem?) × data da pasta do mês.
- Testes: o STATE diz quantos passavam e com qual comando — não rode a suíte inteira por padrão; só se a usuária pedir.

## Relatório (adaptar; seções vazias = "—")
```
PROJETO      Diretório: …   Objetivo: Entrega de Produtos - Coordenação (G1 + pacote Parceiro)
SESSÃO       Número: NNN   Nome sugerido: …   Objetivo: …
GIT          Branch: …   HEAD: …   Status: limpo | N alterados / M novos (listar os relevantes)
ÚLTIMO ESTADO VERIFICÁVEL   (data do STATE + último commit)
CONCLUÍDO / PARCIALMENTE CONCLUÍDO / EM ANDAMENTO / PENDENTE
ESTADO DA ENTREGA MENSAL    competência · etapa · veredito · BLOCKER/MANUAL por grupo · código de publicação
PRÓXIMA TAREFA EXATA        + CRITÉRIOS DE CONCLUSÃO
ARQUIVOS DE ALTO SINAL      (só os necessários para a próxima tarefa)
DECISÕES RELEVANTES         (DEC-xxx que restringem a próxima tarefa)
VALIDAÇÃO                   (último resultado registrado: comando + resultado; não reexecutado)
INCONSISTÊNCIAS             (STATE × Git × disco × handoff × entrega)
PENDÊNCIAS / BLOQUEIOS
```
Termine **exatamente** com:

Contexto reconstruído. Nenhuma alteração foi realizada.

E **PARE**. Aguarde a autorização da usuária antes de qualquer ação.

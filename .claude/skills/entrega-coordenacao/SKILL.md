---
name: entrega-coordenacao
description: Etapa da COORDENAÇÃO EXECUTIVA do Programa antes do envio mensal ao Parceiro - auditar a pasta mensal (G1 a G5), validar Lista de Demandas, Relatório de Produtos Mensal, assinaturas, anexos, nomes, ZIP, 100 MB, versões, duplicidades e Índice; montar o pacote final pronto para protocolo, revalidar, gerar manifesto e relatório de prontidão. Use para "auditar a entrega de setembro", "conferir os grupos", "preparar o pacote do Parceiro", "está pronto para protocolo?". Sempre em dry-run primeiro; nunca envia nada ao Parceiro.
argument-hint: "[nº da pasta mensal, ex.: 15]"
---

# Entrega de Produtos — Coordenação (pré-Parceiro)

Regras do domínio: `docs/dominio/processo-entrega-mensal.md` (leia antes da 1ª execução da sessão).
Assinaturas: `docs/dominio/assinaturas.md`. Divergências antigas×POP: `docs/dominio/regras-pop-x-antigas.md`.
Comandos, saídas e protocolo de publicação: `references/fluxo-operacional.md` (leia ao executar).

## Fluxo (pare entre as etapas que alteram algo)
0. **Conferir a sincronização** (API, só leitura): `python scripts/sharepoint.py comparar-pasta "<pasta local do mês>"
   --recursivo`. Auditar só com "EM DIA". Se divergir, aguardar o OneDrive ou avisar a usuária (DEC-052).
1. **Auditar (dry-run, somente leitura)** — `python scripts/entrega.py auditar --mes $ARGUMENTS`
   Leia `output/ENTREGA_*/04_RELATORIOS/RELATORIO_PRONTIDAO.md` e apresente por grupo: BLOCKER, MANUAL, WARNING,
   e as correções automáticas que serão feitas. Pendências dos grupos → lista para a usuária cobrar (POP Etapa 4).
2. **Preparar (local)** — `python scripts/entrega.py preparar --mes N` → STAGING + FINAL em `output/`, revalidação
   completa do FINAL, manifesto SHA-256, código de publicação. Nada no SharePoint muda.
3. **Revisão humana** — itens MANUAL_REVIEW só viram PASS com registro em `revisoes_humanas.json` (preso ao SHA-256),
   feito **com a decisão da usuária** (nunca por conta própria).
   Exceção desta competência (inclusive BLOCKER), relatório de nome fora do padrão e achado sem arquivo: tipos
   `aceito_excecao`, `relatorio_do_grupo` e `{grupo, etapa}` — sempre listados no relatório de prontidão (DEC-057).
4. **Publicar (altera o SharePoint)** — **só com autorização explícita da usuária nesta conversa**, citando o código:
   `python scripts/entrega.py publicar --mes N --confirmar <código>`. Arquiva antes (backup local + arquivo interno),
   aplica o FINAL na pasta do mês, reconfere contra o manifesto. Depois, `comparar-pasta` de novo para confirmar que o
   SharePoint ficou igual.
5. **PARE** e apresente "PACOTE PRONTO PARA Parceiro" (ou o que falta). A cópia para a pasta compartilhada com o Parceiro,
   protocolo, SEI e e-mail são da usuária — **nunca** automatizar.

## Regras invioláveis
- Nunca considerar "validado" porque o script terminou: o veredito é o do relatório (READY só com 0 BLOCKER e 0 MANUAL).
- Nunca inventar nº SEI, assinatura, rodovia ou versão vigente; ambiguidade = BLOCKER/MANUAL para a usuária decidir.
- Nunca editar originais à mão nem documentos de `docs/referencias-oficiais/`.
- G1 (P1.1/P1.2) é produzido pela Skill `relatorio-prog-grupo1`; esta Skill só audita/empacota o resultado assinado.
- API do SharePoint (`scripts/sharepoint.py`): para conferir e para envios avulsos autorizados. Nunca exclui, move
  ou renomeia, nem grava em `5. Compartilhado Parceiro`. Ver `references/fluxo-operacional.md` § SharePoint pela API.
- Estrutura das pastas futuras: `python scripts/entrega.py estrutura --mes N` (dry-run) / `--aplicar` com autorização.

# Fluxo operacional da Coordenação — comandos, saídas, leitura e autorização

## Comandos (raiz do projeto; Python 3.12)
| Etapa | Comando | Escreve onde |
|---|---|---|
| Padronizar pastas de um mês futuro | `python scripts/entrega.py estrutura --mes 19` → revisar → `--aplicar` | pasta do mês (só cria / remove subpasta **vazia** / cria Índice) |
| Índice em branco de um mês | `python scripts/entrega.py indice --mes 19 [--aplicar]` | raiz do mês (nunca sobrescreve) |
| Conferir sincronização (antes de auditar e depois de publicar) | `python scripts/sharepoint.py comparar-pasta "<…\Programa - 4. Produtos
. Setembro 2026>" --recursivo` | nada (só leitura pela API) |
| Auditar | `python scripts/entrega.py auditar --mes 15` | só `output/ENTREGA_15_Setembro_2026/` |
| Preparar | `python scripts/entrega.py preparar --mes 15` | só `output/…` (02_STAGING, 03_FINAL, 04_RELATORIOS) |
| Publicar | `python scripts/entrega.py publicar --mes 15 --confirmar <código>` | pasta do mês + arquivo interno + backup local |
`--pasta "<caminho>"` substitui `--mes` (ex.: testes). Códigos de saída: 0 READY · 1 pendente revisão humana · 2 bloqueado.

## Saídas (`output/ENTREGA_NN_Mes_AAAA/`)
- `01_INVENTARIO/inventario_origem.json`, `plano_pacote.json` — o que foi visto e o plano (copiar/zipar/renomear/fracionar/excluir).
- `02_STAGING/` — cópia fiel da origem (evidência; SHA-256 conferido).
- `03_FINAL/<pasta do mês>/` — pacote no layout do POP.
- `04_RELATORIOS/RELATORIO_PRONTIDAO.md` (para a usuária), `achados.json/.csv`, `manifesto.json/.csv` (caminho, nome,
  extensão, tamanho, grupo, classificação, SHA-256, resultado), `transformacoes.json` (antes→depois→motivo→regra→resultado).
- `plano_publicacao.json` (remover após arquivar / gravar / manter + **código**), `estado.json` (checkpoint lido pelo
  /session-start), `revisoes_humanas.json` (opcional), `log_publicacao.json`, `05_BACKUP_PUBLICACAO/`.
- `cache.json` — SHA-256 por (caminho, tamanho, mtime) e resultados por conteúdo; apagar é seguro (só recalcula).

## Como ler o relatório
- **BLOCKER**: impede o protocolo (ex.: relatório sem assinatura digital, nome sem rodovia, versões conflitantes,
  arquivo solto na raiz do grupo, PDF com senha, colisão de nomes, parte faltando, .zip que não é ZIP).
- **MANUAL_REVIEW**: precisa de olho humano (assinatura só visual, assinatura alterada depois de assinar, PDF sem texto,
  Lista ausente, conteúdo idêntico com nomes diferentes, termo genérico).
- **FIXED_AUTOMATICALLY** (só na origem): o que o pacote corrige sozinho (Lista/Word/versão não assinada fora,
  achatamento, ZIP, fracionamento, nome com proposta completa, índice regenerado).
- **WARNING**: atenção sem bloquear (planilha sem PDF correspondente, produto sem anexo, caminho longo...).
- Fase `ORIGEM (pendência herdada)`: problema da origem que o pacote não resolve — entra no veredito.

## Protocolo de autorização (obrigatório)
1. `estrutura --aplicar`, `indice --aplicar` e `publicar` **alteram o SharePoint** → pedir autorização explícita
   da usuária **nesta conversa**, mostrando o plano (dry-run) e, para publicar, o **código**.
2. Autorização vale para aquela ação/aquele mês; não generalizar.
3. Depois de publicar: mostrar `log_publicacao.json` → `conferencia_final` deve ser "IDÊNTICA AO MANIFESTO".
4. Nunca copiar para `Programa - 5. Compartilhado Parceiro`, nunca e-mail/SEI/protocolo.

## SharePoint pela API (DEC-052)
- A auditoria lê a pasta **sincronizada**. Se o OneDrive ainda não baixou o que um grupo subiu, a auditoria acusa
  falta que não existe, ou deixa passar um arquivo que já foi trocado. Por isso, **antes de auditar** rode
  `comparar-pasta` no mês (`--recursivo`). Se não estiver "EM DIA", aguarde a sincronização ou avise a usuária e
  registre no relatório. Também dá para listar sem depender da sincronização: `sharepoint.py listar "4. Produtos/15. Setembro 2026/G3"`.
- `publicar` continua pelo `entrega.py`, na pasta sincronizada: arquiva, remove e grava, coisas que a API **não faz**
  de propósito (sem excluir/mover). Depois de publicar, rode `comparar-pasta` de novo para confirmar que o SharePoint
  ficou igual ao FINAL.
- Arquivo avulso que a usuária mandar por fora (ex.: relatório assinado de um grupo por e-mail): `sharepoint.py
  enviar <arq> --para "4. Produtos/<mês>/<Gx>/<nome no padrão>"` (dry-run → autorização → `--aplicar`). Arquivo novo
  por padrão; trocar um existente só com `--substituir` (backup automático).
- Nunca pela API: `5. Compartilhado Parceiro` (bloqueado no código), excluir.
- **Reorganizar a pasta do mês no formato do POP (DEC-070)** — só com autorização explícita: `python scripts/manutencao/reorganizar_mes_pop.py snapshot --mes N [--aplicar]` (cópia "como enviado" no servidor para `3. Desenvolvimento/G1 - Coordenacao/Entregas Recebidas/`) → decisões em `config/organizacao_pop.json[N]` → `planejar` (código) → `aplicar --confirmar <código>` (envia novos, move/renomeia, retira do pacote movendo para o Desenvolvimento, Índice completo) → `verificar`. Rode com `PROG_ROOT` no espelho em dia.

## Correções que a automação NÃO faz (devolver ao grupo)
Escolher entre versões diferentes; inventar rodovia/código; fracionar PDF assinado (quebra a assinatura); OCR;
assinar; decidir destino de arquivo solto na raiz; preencher SEI.

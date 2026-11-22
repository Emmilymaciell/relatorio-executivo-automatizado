# Como usar — Entrega de Produtos - Coordenação

Tudo roda na raiz do projeto (Claude Code aberto nesta pasta). Python 3.12 + Office instalado.

## 1. ABRIR SESSÃO
```
/session-start
```
Reconstrói o contexto pelo repositório, mostra o estado da entrega do mês e **para** esperando você autorizar.

## 2. GERAR / ATUALIZAR O RELATÓRIO DE PRODUTOS MENSAL DO G1 (P1.1/P1.2) — 4º/5º dia útil
Peça ao Claude (Skill `relatorio-prog-grupo1`):
> *Crie o Relatório de Produtos Mensal nº 15 referente a setembro de 2026.*

Ele prepara `config/parametros.json`, roda os scripts `01`→`19` na ordem do `SKILL.md` (sempre `PROG_DRY_RUN=1`
primeiro), faz QA (`13`) e Revisão Final (`18`). Checkpoints em que ele vai te pedir algo: export do Dashboard
(Power BI), marcos da Timeline se faltarem, status da medição se não estiver na planilha.
Dashboard: exporte com o nome `P1.1 - Gestão Executiva - Dashboard da Análise de Entrega de Produtos - <mês> -
GERAL.pdf` em `G1\ANEXOS` (meses a partir de Outubro/2026; até Setembro, na pasta `Anexos\P1.1 …`).
Depois mande assinar; o PDF assinado volta para a raiz do `G1`.

## 3. AUDITAR A ENTREGA DO MÊS (G1–G5) — somente leitura
```
python scripts/entrega.py auditar --mes 15
```
ou peça: *"Audite a entrega de setembro"* (Skill `entrega-coordenacao`). Relatório:
`output\ENTREGA_15_Setembro_2026\04_RELATORIOS\RELATORIO_PRONTIDAO.md`. Use a lista de BLOCKER/MANUAL para cobrar os
grupos (POP Etapa 4) e rode de novo depois das correções (é idempotente; usa cache).

## 4. PREPARAR O PACOTE (local, não mexe no SharePoint)
```
python scripts/entrega.py preparar --mes 15
```
Monta STAGING + FINAL em `output\…`, revalida tudo no FINAL, gera manifesto (SHA-256) e um **código de publicação**.
Item que precisa do seu olho (MANUAL_REVIEW) e que você aceitar: o Claude registra em `revisoes_humanas.json`.

## 5. APLICAR O PACOTE NA PASTA DO MÊS (SharePoint) — só com sua autorização
```
python scripts/entrega.py publicar --mes 15 --confirmar <código>
```
Antes de mexer, arquiva tudo que sai/muda (backup local + `Desenvolvimento de Produtos\Controle Interno Entregas`).
Recusa se houver BLOCKER ou se a pasta mudou desde o `preparar`. Depois, **você** copia a pasta do mês para a pasta
compartilhada com o Parceiro. Nada é enviado automaticamente.

## 6. PREPARAR AS PASTAS DE UM MÊS NOVO (estrutura do POP + Índice)
```
python scripts/entrega.py estrutura --mes 19
python scripts/entrega.py estrutura --mes 19 --aplicar
```

## 7. ENCERRAR
```
/session-close
```
Persiste estado/handoff/histórico, registra validações reais e para. Para salvar versão no Git, peça
*"salva essa versão"* (ou `/versionar`) — nunca é automático.

## 8. SHAREPOINT PELA API (Microsoft Graph — DEC-052)
Credencial: cole o **valor** do client secret em `%USERPROFILE%\.prog\sharepoint.env` (linha
`PROG_SP_CLIENT_SECRET=`). Nunca no repositório, no OneDrive, em e-mail ou no chat.
```
python scripts/sharepoint.py testar
python scripts/sharepoint.py listar "4. Produtos/15. Setembro 2026/G1"
python scripts/sharepoint.py comparar "<arquivo local sincronizado>"
python scripts/sharepoint.py comparar-pasta "<pasta local sincronizada>" --recursivo
python scripts/sharepoint.py publicar-lote manifesto.json
python scripts/sharepoint.py enviar "<arquivo>" --para "3. Desenvolvimento/.../nome.xlsx"
python scripts/sharepoint.py enviar "<arquivo>" --para "..." --aplicar
```
`enviar` é dry-run sem `--aplicar`; arquivo que já existe só com `--substituir` (backup em
`output/SHAREPOINT_BACKUPS/`, eTag e SHA-256 conferidos). Nunca grava em `5. Compartilhado Parceiro`; não exclui,
move nem renomeia. Log: `logs/sharepoint/`. `comparar-pasta` mostra se o OneDrive está em dia (use antes de gerar
o G1 e antes de auditar). `publicar-lote` envia vários arquivos de uma vez: se um item falhar no plano, nenhum é
enviado. Passo a passo do G1: `.claude/skills/relatorio-prog-grupo1/references/sharepoint-api.md`.

## Testes
```
python -m pytest tests -q -p no:cacheprovider
```

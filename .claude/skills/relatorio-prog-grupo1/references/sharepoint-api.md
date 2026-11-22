# SharePoint pela API (Microsoft Graph) no fluxo do G1 — DEC-052

Dois canais para o MESMO SharePoint (`contoso.sharepoint.com/sites/exemplo`, biblioteca `Programa`):
- **Pasta sincronizada (OneDrive)**: é por ela que os scripts 01–19 LEEM as fontes e gravam os produtos
  (`config/caminhos.json`). Continua sendo o caminho de trabalho.
- **API** (`python scripts/sharepoint.py …`, na raiz do projeto): para conferir a sincronização e publicar
  sem depender dela. Segredo em `%USERPROFILE%\.prog\sharepoint.env` (nunca no repo/chat/log). Se o
  `testar` falhar por segredo ausente/expirado, PEDIR à usuária para atualizar esse arquivo.

As regras não mudam: as fontes oficiais e as proibições (`5. Compartilhado Parceiro`, `Obsoleto` como fonte) valem
igual, e só se ALTERA o que a regra permanente 1 permite, com autorização explícita da usuária na conversa.

## Quando usar, por etapa
| Etapa | Uso da API | Comando |
|---|---|---|
| 0 Preparar | Conferir se a pasta local está **em dia** antes de ler as fontes. Se não estiver, aguardar o OneDrive ou avisar a usuária, e **não** gerar com fonte desatualizada | `comparar-pasta "<…\Desenvolvimento de Produtos>" --recursivo` e `comparar-pasta "<…\4. Produtos\NN. Mês ano\G1>" --recursivo` |
| Fonte recebida fora do SharePoint (chat, Downloads, e-mail: planilha de medição, marcos, reuniões) | Conferir o arquivo, prepará-lo em `output/RM{n}_PREPARACAO/` e **publicar** com autorização | `enviar <arq> --para "<remoto>"` (arquivo novo); `--substituir` quando trocar um existente (faz backup) |
| 5 Timeline (marcos informados pela usuária) | Preencher a `Base Timeline` numa **cópia** (Excel COM no Temp, linhas novas na competência `01/MM/AAAA`, `Incluir=Sim`) e publicar substituindo o original | `enviar "<cópia>" --para "3. Desenvolvimento/G1 - Coordenacao/Desenvolvimento de Produtos/P1.1 -Controle Timeline.xlsx" --substituir` |
| 7 Medição | A planilha do mês entra como **arquivo novo** `Medicao Parceiro/Planilha de Medição - Estruturadora - {Mês}.xlsx`. Antes, comparar com a do mês anterior: só o bloco da nova medição pode mudar | `enviar … --para "…/Medicao Parceiro/Planilha de Medição - Estruturadora - {Mês}.xlsx"` |
| 13 Finalizar / anexos | Padrão: os scripts gravam pela pasta sincronizada. **Se a gravação local não for possível** (sem permissão ou OneDrive parado): rodar com `PROG_DRY_RUN=1` (gera tudo em `output/RM…/07_SAIDA`), montar um manifesto com os destinos (DOCX/PDF na raiz do `G1` do mês; anexos P1.1/P1.2 na pasta de anexos do G1, `cfg.anexos_dir()`, com o nome de `cfg.nome_anexo_g1()`) e publicar | `publicar-lote manifesto.json`, conferir o plano e depois `--aplicar` |
| Depois de publicar | Confirmar que o SharePoint e a pasta local ficaram iguais | `comparar-pasta "<…\G1 do mês>" --recursivo` |
| OneDrive não hidrata ("somente online", acesso negado) — DEC-066 | Espelho de LEITURA numa raiz curta e scripts com `PROG_ROOT` no espelho; publicar só pela API, depois de conferir que o original não mudou (SHA-256 do `_espelho.json`) | `espelhar "<remoto>" --destino "C:\caminho\exemplo\PROG_ESPELHO" [--so-pasta]` |

**Exceção:** o acumulado `P.1.1 - Controle de demandas Programa.xlsx` (etapa 2) continua sendo alterado só pelo
`02_atualizar_controle_demandas.py` na pasta sincronizada, com backup e só acrescentando linhas. Não publicar pela
API uma versão editada fora do `02`.

## Leitura do resultado
- `enviar`/`publicar-lote` sem `--aplicar` = só o plano. `NOVO` · `SUBSTITUIR` (com `--substituir`) ·
  `BLOQUEADO` (já existe e falta `--substituir`). No lote, um item BLOQUEADO ou com ERRO faz o lote inteiro
  não ser enviado.
- Verificação depois do envio:
  - `OK`: idêntico byte a byte.
  - `OK_CONTEUDO`: arquivo Office em que só os metadados que o SharePoint grava (`docProps`/`customXml`) diferem.
  - `DIVERGENTE`: conferir antes de seguir.
- Backups dos arquivos substituídos ficam em `output/SHAREPOINT_BACKUPS/<data-hora>/` (com `.meta.json`); o log de
  envios fica em `logs/sharepoint/AAAA-MM.jsonl`.
- **Não existe** excluir, mover nem renomear pela API. Reorganizar pastas é decisão da usuária, com plano aprovado.

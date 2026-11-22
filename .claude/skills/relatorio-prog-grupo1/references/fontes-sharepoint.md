# Fontes no SharePoint (Programa)

## Caminho lógico vs local
- **Lógico** (URL): `https://contoso.sharepoint.com/sites/exemplo/Programa/` → `3. Desenvolvimento`,
  `4. Produtos`, etc.
- **Local** (descoberto nesta máquina): raiz `C:\caminho\exemplo\Dynatest Engenharia`, com
  subpastas prefixadas `Programa - ` (ex.: `Programa - 3. Desenvolvimento`, `Programa - 4. Produtos`).

## Acesso pela API (Microsoft Graph) — DEC-052
Além da pasta sincronizada, o projeto acessa a biblioteca pela API com `python scripts/sharepoint.py`
(`testar`, `listar`, `info`, `baixar`, `comparar`, `comparar-pasta`, `enviar`, `publicar-lote`). Mapeamento: a pasta
local `Programa - X` corresponde a `X` na biblioteca (`config/sharepoint.json`). Use a API para conferir a sincronização
antes de ler as fontes e para publicar o que a regra permite. Procedimento por etapa: `sharepoint-api.md`.

## Localização automática (se o caminho mudar)
`prog_config.prog_root()`: usa `PROG_ROOT` (env) → `config/caminhos.json` → **busca** em `~` por
uma pasta que contenha subpastas cujo nome (sem acento/caixa) tenha "3. desenvolvimento" e
"4. produtos". Todos os `_find_sub` ignoram acentos/maiúsculas/espaços. Portanto **não depender de
nome exato** ("Coordenacao"/"Coordenação", "Gestao"/"Gestão", espaços duplos).

## Fontes por categoria (nomes reais)
### MODIFICÁVEL (com backup, só acrescentar)
- `…\G1 - Coordenacao\Desenvolvimento de Produtos\P.1.1 - Controle de demandas Programa.xlsx`
  — acumulado; **fonte do Power BI**. Só ACRESCENTAR linhas da versão N. Backup antes.

### SOMENTE LEITURA — fonte oficial de dados
- `…\G2 -  Engenharia\00_ Lista mestre.xlsx` — entregas de engenharia (filtrar `N° Relatório`).
- `…\4. Produtos\{pasta mensal}\G1\PROG_Lista de demandas G1 Rev.CG.xlsx` — entregas 1.3 e 1.4.
- `…\{pasta mensal}\G3\PROG_Lista de demandas (G3).xlsx` — Certificadora.
- `…\{pasta mensal}\G4\PROG_Lista de demandas G4.xlsx` — CG.
- `…\{pasta mensal}\G5\...PROG_Lista de demandas ...julho.xlsx` — Escritório Jurídico.
- `…\{pasta mensal}\G1\Programa - Relatório Grupo 1 - ...RevCG.docx` — **base** do Word (copiar, não editar).
- **Medição do mês** — `…\Desenvolvimento de Produtos\Medicao Parceiro\Planilha de Medição - Estruturadora - {Mês}.xlsx`
  (DEC-034). O painel `P1.1 - Quadro Financeiro` é **produzido** a partir dela (`medicao.md`); não é fonte.
- **P1.2 - Controle de Reuniões** — usar a **versão mais recente** enviada (comparar data de
  modificação/quantidade de reuniões). Original em `…\Desenvolvimento de Produtos\P1.2 - Controle de Reuniões.xlsx`.
- **Timeline PPTX** — o usuário fornece/corrige a versão do mês; a versão do corpo é a base.

### SOMENTE MODELO VISUAL (não é fonte de dados)
- Quadro Financeiro/Timeline/Reuniões do **mês anterior** (`…\4. Produtos\{mês anterior}\G1\Anexos\...`):
  referência de **layout**, nunca de números.
- Baseline v5 do RM13 em `examples/baseline/` do projeto: referência **visual**, não factual.
- Fixture `.claude/skills/relatorio-prog-grupo1/assets/texts_exemplo_julho2026.py`: **EXEMPLO de
  estrutura** do `config/dados_mes.py` (dados de Julho/2026). **Não é fonte** e **nunca** deve ser
  usado como dados de outra competência — só como fixture de teste com `PROG_USE_EXAMPLE_DATA=1`
  (DEC-019). Em MODO REAL, o conteúdo do mês vem das fontes oficiais + `config/dados_mes.py`.

### NUNCA MODIFICAR
Qualquer original do SharePoint (exceto o acumulado autorizado). Não excluir, mover, renomear,
sobrescrever, alterar permissões/compartilhamentos.

## Pasta mensal (produtos) — padrão POP a partir de Outubro/2026 (DEC-043/047)
`…\Programa - 4. Produtos\{NN}. {Mês} {ano}\` = raiz com o **Índice** (`Índice de Arquivos - Entrega de Produtos Programa -
{Mês} {ano}.docx`, Modelo r2) + `G1..G5\` cada um com o Relatório de Produtos Mensal assinado na raiz e `ANEXOS\` única. As
pastas 16/17/18 (out–dez/2026) já foram reorganizadas em 01/10/2026; meses futuros: `python scripts/entrega.py
estrutura --mes N`. Essa pasta **é** o pacote enviado ao Parceiro (a usuária copia) — DEC-045. Internos (Listas,
Word, versões) vão para `…\Desenvolvimento de Produtos\Controle Interno Entregas\{pasta do mês}\` na publicação.

### Layout legado (até Setembro/2026)
`…\Programa - 4. Produtos\{numero}. {Mês} {ano}\` (ex.: `13. Julho 2026`) com `G1..G5`. Em `G1\Anexos\`:
`P1.1 - Gestao Executiva`, `P1.2 - Relacoes Institucionais`, `P1.3 - Comunicacao Integrada`
(Evidências 1.3 …zip), `P1.4 - Tecnologia de Informacao de Comunicacoes` (Evidências 1.4 …zip).
> Nomes das subpastas variam por mês (ex.: junho usou "P1.1 - Coordenação Executiva"). Descobrir a
> pasta P1.1/P1.2 existente por busca ("p1.1"/"p1.2") — não criar pasta duplicada.

## Arquivos somente na nuvem (placeholders OneDrive)
Se um arquivo estiver como placeholder: **não mover/excluir**; apenas **abrir/copiar** para a pasta
de trabalho para forçar a hidratação; aguardar o download concluir. Somente para leitura/cópia.

## Escolha do arquivo mais atualizado
Quando houver mais de um candidato numa pasta: (1) conferir os **cabeçalhos** (uma lista de entregas
tem Produto/Empresa/Documento/Tipo/Descrição/Empreendimento/Data/Status); (2) excluir arquivos de
Quadro Financeiro/Reuniões/Dashboard/Timeline/modelos; (3) entre listas válidas, comparar **data de
modificação** e conteúdo; escolher a mais recente com dados atualizados. Registrar a escolha (arquivo,
caminho, grupo, mtime, nº de linhas, motivo) num `manifest_fontes.json`.

## Medição Parceiro
Pasta `…\Desenvolvimento de Produtos\Medicao Parceiro`. **FONTE OFICIAL DA MEDIÇÃO (DEC-034):**
`Planilha de Medição - Estruturadora - {Mês}.xlsx` (todo mês, com o nome do mês). Global = med 1–6,
SUPRIMIDA = med 7–12+. O **painel** `P1.1 - Quadro Financeiro` é **produzido** copiando o do mês
anterior e preenchendo com essa Planilha (ver `medicao.md`). *(Legado DEC-018 — "usar o Quadro
pronto"/"pasta vazia" — foi SUPERADO: a pasta agora tem a Planilha oficial.)*

## FONTES OFICIAIS por componente (produzir do zero todo mês — DEC-037)
- **Entregas:** listas `PROG_Lista de demandas` G1/G3/G4/G5 (pasta mensal) + `00_ Lista mestre.xlsx` (G2).
- **Medição/painel:** `Medicao Parceiro\Planilha de Medição - Estruturadora - {Mês}.xlsx` (DEC-034).
- **Timeline (marcos):** `Desenvolvimento de Produtos\P1.1 -Controle Timeline.xlsx`, aba `Base Timeline`
  (DEC-035). Conferir/atualizar; se faltar, **perguntar ao usuário** antes de gerar.
- **Reuniões:** **mestre** `Desenvolvimento de Produtos\P1.2 - Controle de Reuniões.xlsx` — **filtrar o
  mês** (abas "Reuniões de GOV 1-1" e "Reuniões Stakehold 1-2" têm todos os meses) para produzir o anexo
  mensal (DEC-036). Não usar um mensal já pronto de outra pasta.
- **Dashboard:** o **usuário exporta** o PDF do Power BI (`.pbix` lê o acumulado atualizado) e coloca em
  pasta de anexos do G1 (`cfg.anexo_p11_dir()`: POP = `G1\Anexos`, nome `P1.1 - Gestão Executiva - Dashboard da Análise
  de Entrega de Produtos - {mês} - GERAL.pdf`; legado = `G1\Anexos\P1.1`); a automação nunca publica Power BI (DEC-037).
- **Base CG (capa+P1.3+P1.4):** `…\{mês}\G1\Programa - Relatório Grupo 1 - …RevCG.docx` (fornecida pela CG).

## Insumos de FORA do SharePoint (Downloads, chat, e-mail…) — importar ANTES de usar (DEC-059)
`python fontes_externas.py inspecionar "<pasta de entrada>"` (classifica; AMBÍGUO = pendência) →
`python fontes_externas.py importar "<arquivo>" --finalidade "…"` (cópia imutável + SHA-256 + proveniência em
`01_FONTES/EXTERNAS/_proveniencia.json`) → `plano-sharepoint` + `sharepoint.py publicar-lote` (cópia durável em
`Desenvolvimento de Produtos/Fontes Externas/{NN. Mês ano}/`). Usar SEMPRE a cópia importada. Prestação de Contas →
Timeline (DEC-060); Relatório de Gestão do Contrato do Parceiro → quadro da medição (DEC-062) e determinações da fiscalização.

## OneDrive sem hidratar ("somente online", acesso negado) — espelho de leitura (DEC-066)
`python scripts/sharepoint.py espelhar "4. Produtos/NN. Mês ano" --destino "C:\caminho\exemplo\PROG_ESPELHO"` (+ G1 do
mês anterior, `Desenvolvimento de Produtos` com `--so-pasta`, `Medicao Parceiro`, Lista Mestre do G2 por `baixar`) e rodar
os scripts com `PROG_ROOT=C:\caminho\exemplo\PROG_ESPELHO`. Publicar só pela API, conferindo antes que o original não
mudou (SHA-256 do `_espelho.json`).

## PROIBIDO como fonte: espelho `5. Compartilhado Parceiro` (DEC-037)
`Programa - 5. Compartilhado Parceiro\05. Produtos\…` é **cópia de entrega**, **não é fonte**. Nunca ler dados
dele nem do `Obsoleto` para produzir o relatório. Todo mês o relatório é **produzido do zero** a partir
das fontes oficiais acima. Se uma fonte oficial faltar, **perguntar ao usuário** — nunca substituir por
cópia do Parceiro/obsoleto nem inventar.

# Regras de Excel (planilhas, prints, COM)

## Ferramentas
- **Excel COM (win32com)** para preservar formatação/gráficos/validação de dados e para
  `ExportAsFixedFormat` (PDF). `openpyxl` só para leitura/consolidação/validação (openpyxl **remove**
  a validação de dados ao salvar — não usar para salvar arquivos com dropdowns).
- Sempre `xl.DisplayAlerts=False`; `xl.AutomationSecurity=1`; fechar com `atexit`+`Quit()`.
- Se o arquivo estiver aberto/instância travada: matar EXCEL órfão antes; conferir se há lock
  `~$arquivo.xlsx` (arquivo aberto pelo usuário) — se houver, pedir para fechar.

## Datas (DEC-001) — crítico
Formato **`dd/mm/aaaa`** (PT-BR); **NUNCA** `dd/mm/yyyy` (vira "yyyy"). Gravar **datas sem hora**.
Preferir **herdar** o formato das células de data já existentes (formato de data curto/local).

## Atualizar o acumulado (script 02) (DEC-002)
1. **Backup** para `06_BACKUP` com timestamp; restaurar o original a partir do backup (garante que
   linhas antigas fiquem intactas).
2. Achar a última linha; copiar o **formato** da última linha para o bloco novo (PasteSpecial formats);
   escrever colunas C..K das linhas da versão N (ID via fórmula `=B{r-1}+1`; K=Versão=N); Data Entrega
   como datetime (sem hora); **sem** sobrescrever formato de data.
3. Validar: contagem por versão antes/depois; v(N) = total consolidado; IDs contínuos.

## Consolidação (script 01)
G2 = Lista Mestre filtrada por `N° Relatório = N`; G1(1.3/1.4)/G3/G4/G5 = todas as linhas populadas
das listas mensais; + 6 recorrentes (clonar do acumulado v(N-1)). Normalizar datas/status
(ver `processo-completo.md`). Dedup por chave composta.
**Produto normalizado pelo código** (nome da versão mais recente do acumulado; produto vazio herda do mesmo
documento) — nome fora do padrão vira linha em branco no Power BI (DEC-055). Conferir `AJUSTE produto` no log.

## Print do Controle de Produtos (script 03) (DEC-009)
Trabalhar numa **cópia** do anexo. **Ocultar** colunas A (spacer), **B (ID)** e **K (Versão)**.
Colunas visíveis e larguras (unid. Excel, ajustar): Produto 30, Empresa 13, Documento 26, Tipo 17,
**Descrição 50 (mais larga)**, Empreendimento 13, Data 12, Status 13. Wrap em Produto/Documento/Tipo/
Descrição; **VerticalAlignment=Top**; fonte 9.
- **Altura por linha calculada** (não uniforme): `linhas = max(ceil(len/CPL))` para Produto/Documento/
  Tipo/Descrição (CPL ≈ largura em chars: 31/27/18/53); `RowHeight = linhas*11 + 3`. Assim 1-linha
  fica compacta e multi-linha só o necessário (**nunca cortar**).
- **Paginação balanceada** (última página cheia): `total=Σ alturas`; `alvo=total/12`; inserir
  `HPageBreaks` quando o acumulado ≥ alvo. Evita a "página com 2 linhas + branco".
- PageSetup: **paisagem**, `FitToPagesWide=1`, `FitToPagesTall=False`, `PrintArea=$C$1:$J${last}`,
  `PrintTitleRows=$1:$3` (repete cabeçalho), margens ~0,25". Exportar PDF; renderizar PNG @220 dpi.

## Tabela das Reuniões (script 04) (DEC-027; substitui a altura de DEC-010)
Cópia do anexo mensal; **ocultar coluna E (Anexos)**. **Nunca** `FitToPagesWide` (comprime a tabela e
some com o tamanho da letra): `Zoom=100` + margens calculadas para a tabela caber **em escala 1:1** na
largura útil do corpo. `PrintArea=$B$1:$D${fim}` numa cópia da aba com **só as linhas da parte** (assim
faixa/logo/cabeçalho se repetem de verdade em cada parte — `PrintTitleRows` não repete formas).
**Calibrar impresso × reportado** (o Excel reporta pt diferentes dos impressos) e dimensionar em pt
impressos. Alturas por `AutoFit` (coluna 1,5% mais estreita) + verificação de **corte medida no PDF**.
Negrito parcial reaplicado por `Characters()` — com early binding use `GetCharacters()`.
Detalhes completos em `reunioes.md`.

## Anexo Controle de Entregas (script complementar)
Cópia do acumulado atualizado; manter só a versão N (AutoFilter Versão≠N → apagar linhas visíveis;
ou congelar IDs a valores e renumerar 1..N no anexo). Ordenar por Data (Sem data ao final). É o anexo
oficial `P1.1 - Controle de Entrega de Produtos de {mês}.xlsx` (a versão do print no corpo NÃO leva ID).

## Regras gerais dos prints
- Alta resolução (≥ 200 dpi). Recortar bordas brancas depois (script 08).
- Nunca cortar/esconder conteúdo. Nunca distorcer. Larguras nas colunas do Excel, não esticando imagem.

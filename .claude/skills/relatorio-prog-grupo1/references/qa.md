# QA — Controle de Qualidade (obrigatório antes de entregar)

Rodar o script `13_qa.py` (varre o texto do PDF) **e** fazer a conferência visual das páginas-chave.
Resultado: `APROVADO` ou `REVISÃO NECESSÁRIA` em `output/.../05_LOGS/qa_RM{n}.txt`. O veredito é
**APROVADO só se não houver nenhuma pendência**, incluindo os **avisos persistentes** de
`cfg.qa_warnings()` gerados por etapas anteriores.

## 0. Avisos persistentes / MODO REAL (bloqueiam a aprovação automática)
- **Dados do mês (DEC-019)**: a produção só ocorre em MODO REAL com `config/dados_mes.py` (ou
  `PROG_DADOS`). Se o QA foi rodado sobre um relatório gerado em MODO TESTE
  (`PROG_USE_EXAMPLE_DATA=1`), isso **não** é entrega válida — os dados são de julho/2026.
- **Timeline (DEC-020)**: se o `07` caiu no fallback do último slide (placeholder `XXXX` não achado),
  há o aviso **"CONFERÊNCIA VISUAL DA TIMELINE NECESSÁRIA"** → veredito **REVISÃO NECESSÁRIA** até a
  conferência visual da slide do mês.
- **Timeline PPTX ausente (DEC-022)**: se o `.pptx`-fonte não foi localizado, o `07` registra
  **"TIMELINE PPTX NÃO LOCALIZADA"** e não gera a Timeline → pendência; obter o `.pptx` e reexecutar o `07`.
- **Dashboard**: o PDF do Power BI da **competência atual** deve existir no anexo P1.1. Ausência =
  pendência (pedir o export ao usuário; nunca reutilizar o de mês anterior). Ver `dashboard.md`.
- **Publicação final (DEC-021)**: em `PROG_DRY_RUN=1`, o `12` **não** publica no `G1` (log
  "DRY_RUN: publicação no SharePoint ignorada"); o DOCX/PDF ficam só em `07_SAIDA`. Uma entrega real
  exige execução **sem** DRY_RUN para versionar no `G1`.

## A. Automático (texto do PDF)
- **Número**: "N° {n}" presente; **não** existe "Relatório Mensal N° {n-1}" como nº corrente.
- **Competência/período**: mês e "01 … {mês} … 31 … {mês}" coerentes (capa/apresentação/encerramento).
- **Datas**: sem literal "/yyyy" nem "/aaaa"; encerramento "Cidade, {data}" **uma única vez**.
- **Cidade**: sem "São Paulo".
- **Assinaturas**: "Pessoa B" e "Pessoa A" presentes; **sem** "Nome – Programa".
- **Anexos**: "LISTA DE DOCUMENTOS ANEXOS" presente; sumário mostra "ANEXOS" (sem "3 ANEXOS").
- **Medição**: "{N-1}ª medição" presente; valores e % conferem (Σ empresas = total; saldo = contrato −
  acumulado; % = acumulado/contrato).
- **Contagens**: total/entregues/em elaboração presentes e = base consolidada.
- **Placeholders**: sem TODO/PREENCHER/XXX/`<<`/`>>`/`[INSERIR]` (cuidado: "todo/todos" em PT é falso
  positivo — checar case-sensitive `\bTODO\b`).
- **Cabeçalho/rodapé em TODAS as páginas (DEC-032)**: toda página contém o cabeçalho ("CONSÓRCIO") e o
  rodapé ("CONFIDENCIAL"). Nenhuma seção pode ter "primeira página diferente" (senão a 1ª página da
  seção sai sem cabeçalho/rodapé — era o caso da pág. 35). O `13` lista as páginas faltantes.
- **Quebra de página dos títulos (DEC-032)**: cada título "Produto 1.X – Relatório …" inicia no **topo
  de uma página nova** (o `13` checa a posição Y de cada título no PDF; se um não estiver no topo,
  acusa a página e o Y). Antes, 1.2 e 1.4 caíam no meio da página.

## B. Visual (renderizar e conferir página a página)
- **Nenhuma imagem invade cabeçalho ou rodapé**; bordas brancas recortadas; imagens centralizadas e
  ampliadas ao máximo; sem grandes vazios; texto próximo das imagens.
- **Páginas horizontais** com cabeçalho/rodapé adaptados (full-width, Programa à esquerda); **páginas
  verticais** com header/footer originais **inalterados**; sem cabeçalho/rodapé duplicado.
- **Controle de Produtos**: paisagem, sem coluna ID, sem texto cortado, última página cheia (não 2
  linhas + branco), escala consistente entre páginas.
- **Dashboard**: sozinho na sua página, **ocupando quase toda a área útil** (base próxima do risco do
  rodapé, recuos laterais mínimos, proporção preservada); **sem página em branco** após ele (DEC-024).
- **Timeline**: título "Timeline" em 1 linha, abaixo da faixa verde; setas reais que **não cruzam**
  texto/datas; sem invadir rodapé; uma página por mês.
- **Reuniões** (DEC-027): sem coluna Anexos; **bordas esquerda/direita coincidindo com o início/fim do
  texto justificado**; conteúdo legível (≈ fonte do corpo, 9,5 pt reais — não 6-7 pt); **logo Programa
  presente e inteira** na faixa de **todas** as partes; cabeçalho `Data | Participantes | Assuntos`
  repetido; órgãos de Participantes e títulos de Assuntos em negrito; Data e Participantes
  centralizados, Assuntos topo/esquerda; **nenhum texto cortado** (conferir a última linha de cada
  célula contra a borda) e **nenhum termo faltando**; alturas sem grandes vazios; quebras entre
  reuniões completas; mesma largura de coluna em todas as partes; ocupação vertical registrada no log
  (nenhuma parte com `<<< EVITÁVEL`).
- **Medição**: quadro consolidado grande, cores preservadas; Linha do Tempo em 3 quadros ampliados,
  cores fiéis; negrito seletivo (não tudo em negrito).
- **P1.3/P1.4**: preservados (CG); sem referência antiga (mês/N-1) fora de contexto histórico.
- **Sumário**: aponta páginas corretas; nada cortado; sem páginas em branco desnecessárias.
- **Assinaturas**: 2 blocos lado a lado; nome/cargo colados à linha.

## C. Segurança (não destrutivo)
- Nenhum original do SharePoint excluído/movido/renomeado/sobrescrito. Só o **acumulado** alterado,
  **com backup** e **apenas acrescentando** linhas. Anexos listados **existem** fisicamente. Nenhum
  anexo presente na pasta que deveria entrar foi esquecido. Não houve publicação de Power BI/e-mail.

## D. Coerência semântica (texto × fonte)
Controle×base; Dashboard×dashboard; Timeline×marcos; Reuniões P1.1×aba Controle Geral; P1.2×aba
Stakehold; Medição×Quadro Financeiro; P1.3/P1.4×material CG. Corrigir o texto se não sustentado.

## E. Erros específicos desta sessão que o QA deve pegar (ver decisoes-e-erros-conhecidos.md)
Datas "yyyy"; IDs sobrepostos no acumulado; alteração de linhas antigas; capa duplicada; "São Paulo";
data repetida no encerramento; encerramento todo em negrito; "Nome – Programa"; "3 ANEXOS"; imagens
invadindo rodapé/pequenas; Controle com ID/última página vazia; Reuniões com Anexos/sobreposição;
Timeline com setas cruzando texto/título colado; Linha do Tempo de Medições pequena/cortada
("11 e 12"); cabeçalho paisagem duplicado.

## Revisão final (`18_revisao_final.py`) — checagens do painel financeiro
- Status da medição (`status_medicao`) presente no RESUMO do painel-anexo (`N54..N57`).
- **% Executado = acumulado ÷ contrato em todas as linhas do resumo, inclusive o TOTAL (`J54..J58`), e `J58` = `J8`**
  (DEC-054; o TOTAL saiu fixo em 39,07% até o RM14).

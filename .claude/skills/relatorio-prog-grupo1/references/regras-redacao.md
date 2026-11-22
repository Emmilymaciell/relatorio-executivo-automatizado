# Regras de Redação

## Princípio
Reproduzir os textos do **relatório anterior (RM11 editável / RM12 assinado)** com a MESMA
estrutura, objetivo, semântica, coesão, tipo de linguagem e **tamanho aproximado dos parágrafos**;
**alterar apenas os dados** do mês (números, datas, mês, marcos, valores). **Não resumir**, não
reinterpretar, não criar versão curta (DEC-008). Modelo de referência: `assets/texts_exemplo_julho2026.py`.

## Estilo
Formal, técnico, institucional, objetivo, coeso; compatível com engenharia/gestão pública. Sem
marketing, exageros, opiniões. Mesmo nível de detalhamento do RM anterior.

## Nunca inventar
Entrega, marco, reunião, valor, %, risco, empreendimento, decisão, participante. Todo número
reproduzível a partir das fontes. Quando a análise depender de interpretação, use formulação técnica
prudente (ex.: "em regra, recebidos com avaliação satisfatória"), sem afirmar dados inexistentes.

## Fonte de cada texto (não misturar bases)
- **Controle de Produtos**: qualitativo (tipos de documentos, empresas, frentes, empreendimentos) +
  contagens da base consolidada.
- **Dashboard**: total/entregues/em elaboração/%, distribuição por frente, por empresa, temporal
  (dia de pico). Números da base (não do screenshot).
- **Timeline**: **somente** os marcos da timeline do mês (autorizados). Nunca escrever marco sem fonte.
- **Reuniões de Governança (P1.1)**: **somente** a aba "Controle Geral" (governança).
- **P1.2 Integração Institucional**: aba "Reuniões Stakehold 1-2" + narrativa institucional das
  interações documentadas. Nunca misturar com a base P1.1 (DEC-011). Incluir o parágrafo-nota do
  RM11 sobre apresentação em síntese técnica (planilha datas/participantes/assuntos).
- **Medição**: valores da planilha financeira; texto fiel ao RM11 com negrito seletivo (ver `medicao.md`).

## Texto-base — P1.1 Reuniões de Governança (padrão fixo, DEC-026)
O texto de **Reuniões de Governança (P1.1)** segue SEMPRE este molde (≈7 parágrafos densos); a cada mês
**mantém-se a estrutura, a semântica e o tamanho** e só se **moldam os dados** do mês. Não encurtar.
1. **Abertura:** "Durante o mês de {mês}, as reuniões de governança do Consórcio de Apoio à Estruturação
   de Projetos (Programa) concentraram-se em {focos do mês}… do Programa de Gestão de Obras do Fundo Estadual de Infraestrutura.
   A Coordenação Executiva manteve sua atuação como instância integradora entre Parceiro, Órgão Contratante, (Órgão Supervisor,)
   empresas estruturadoras, certificadora e assessoramento jurídico, assegurando alinhamento entre as
   equipes e rastreabilidade das decisões."
2–6. **Parágrafos por TEMA/AGENTE** (usar os que existirem no mês; ordenar pelo que dominou a pauta):
   governança interna do Programa (estrutura, matriz de responsabilidades/RACI, pontos focais); reuniões
   institucionais (Parceiro/Órgão Contratante — prestação de contas, repasses, medidas estruturantes); FUP/técnico-
   operacional (status por empreendimento RD-101/102A/103/104: geotecnia, drenagem, desapropriação,
   projetos; riscos de avanço sem projeto aprovado); certificação e conformidade (Certificadora —
   não conformidades, certificação como condição de repasse); governança financeira (Consultoria de Gestão —
   prestação de contas, medições, painéis/BI, riscos); jurídico-institucional (Escritório Jurídico/Procuradoria —
   reajuste/aditivos, TAG, segurança jurídica, tratativas TCE/MP).
7. **Síntese:** "Em síntese, {mês} foi marcado por {balanço dos eixos}… As reuniões realizadas no período
   ampliaram a rastreabilidade das decisões, reforçaram a fundamentação técnica e documental e
   contribuíram para a mitigação de riscos e a continuidade operacional dos empreendimentos vinculados
   ao Fundo Estadual de Infraestrutura."
**Como moldar:** ler os assuntos reais das reuniões do mês (aba de governança) e escrever cada parágrafo
com a mesma densidade dos meses-modelo (abr/mai/jun em `assets/`). Se um agente não teve reunião no mês,
substituir aquele parágrafo pelo tema que dominou (ex.: em julho, FUPs conjuntos no lugar de reuniões
separadas por empresa). Exemplo vigente (julho): `assets/texts_exemplo_julho2026.py` (campo `REUNIOES`).
> O mesmo espírito vale para **P1.2 Integração Institucional**: parágrafos analíticos + **parágrafo de
> consolidação/síntese** ao final ("Em síntese, …").

## Texto-base — P1.2 "Relatório de Relacionamento Institucional" (consolidado, DEC-033)
Além do texto introdutório de P1.2 (`P12`) e da **tabela de reuniões 1.2**, o P1.2 tem, **DEPOIS da
tabela**, um **texto consolidado** intitulado *"Relatório de Relacionamento Institucional do mês de
{mês} de {ano}"* (o `09` gera o título; os parágrafos vêm de `cfg.dados().P12_CONSOLIDADO`). Esse campo
é **obrigatório em todo mês** — o `13` (QA) acusa se faltar.
**Como gerar a cada mês (processo):**
1. **Modelo de referência = o texto do mês anterior** (estrutura/coesão/densidade). O molde vigente está
   em `assets/texts_exemplo_julho2026.py` (campo `P12_CONSOLIDADO`, julho/2026) — **estrutura**, não
   dados.
2. **Fonte dos dados = a planilha de reuniões 1.2 do mês** (`P1.2 - Controle de Reuniões`, aba de
   stakeholders): ler os **assuntos** de cada reunião e refletir no texto os temas, empreendimentos,
   eventos e decisões efetivamente registrados no mês. **Nunca inventar** (mesma regra dos demais textos).
3. **Reescrever mantendo a MESMA estrutura por temas** (≈12 parágrafos): (1) definição do relatório;
   (2) panorama do mês + empresas; (3) governança e gestão integrada; (4) repositório/comunicação/sede;
   (5) engenharia (Dynatest); (6) certificação; (7) não conformidades/notificações/sanções;
   (8) medições/financeiro/prestação de contas; (9) suporte às construtoras/representação CERT;
   (10) jurídico-contratual (Escritório Jurídico); (11) síntese integrada; (12) governança documental e efetividade.
   Se um tema não teve tratativa no mês, adaptar o parágrafo ao que dominou (não deixar vazio nem inventar).
4. **Gravar** o resultado em `config/dados_mes.py` como `P12_CONSOLIDADO = [ ...parágrafos... ]`. O `09`
   insere automaticamente após a tabela de reuniões 1.2.
> Ao copiar/colar textos de PDFs, **remover cabeçalho/rodapé** que entram no meio ("Programa | CONSÓRCIO…",
> "-CONFIDENCIAL- {pág}") e **rejuntar** frases partidas por eles.

## Títulos de campo
Estilo lista/bullet (List Bullet), **sem negrito**, uma linha em branco acima (DEC-008).

## Negrito seletivo (onde o modelo usa)
- Encerramento: só "Relatório Mensal N° {n}", "{mês} de {ano}", período (DEC-005).
- Medição (parágrafos de valores): número da medição, valores R$, %, frases-chave e produtos citados
  (padrão RM11 — ver `medicao.md`).

## Coerência com a fonte (QA semântico)
Texto do Controle × base; texto Dashboard × dashboard; texto Timeline × marcos; texto Reuniões ×
aba correspondente; texto Medição × planilha financeira; P1.3/P1.4 × material da CG. Se o texto
afirmar algo não sustentado pela fonte, corrigir o texto.

# Dashboard Analítico de Entregas

## Origem e limitação
- Power BI Desktop `P.1.1 - Dashboard - Análise de Entregas Programa.pbix`, cuja fonte é o **acumulado**
  `P.1.1 - Controle de demandas Programa.xlsx` (query SharePoint). Por isso a **ordem importa**:
  **primeiro** atualizar o acumulado (script 02), **depois** o Power BI.
- **Não automatizável headless** (exige interface/autenticação/refresh). **Checkpoint**: o **usuário**
  faz o refresh, filtra o mês (slicer da competência) e **exporta o PDF** para a pasta P1.1 como
  `P1.1 - Dashboard da Análise de Entrega de Produtos - {mês}.pdf`. Não publicar o BI.

## Disponibilidade do PDF por competência
- **Meses já concluídos**: o PDF do Dashboard normalmente **já está** em
  `…\{mês}\G1\Anexos\P1.1 - Gestao Executiva\` (foi exportado quando o relatório foi produzido) —
  reprodutível a partir dali (só leitura).
- **Competência NOVA**: o PDF **precisa ser exportado** pelo usuário **após atualizar o PBIX**
  (refresh + filtro do mês). Sem isso, não há Dashboard do mês.
- **Ausência do PDF da competência atual** = **pendência de QA** (não gerar o corpo dando a etapa por
  concluída): registrar e pedir o export ao usuário (checkpoint).
- **NUNCA** reutilizar o Dashboard de mês anterior como se fosse o atual — é fonte factual da
  competência errada (mesma lógica de DEC-019 para dados). O texto dos indicadores vem sempre da
  **base consolidada** do mês, não de um screenshot antigo.

## Imagem
Converter o PDF exportado em PNG (PyMuPDF, ~200 dpi) → `03_IMAGENS/DASHBOARD`. Recortar bordas
brancas (script 08). No corpo: **todo o texto primeiro, imagem sozinha na página** (DEC-012).
A imagem usa a área **"página cheia"** (`area=(PW, PH_FULL)`, `PH_FULL≈10,1in`): ocupa quase toda a
área útil entre os riscos — base bem próxima do risco do rodapé, recuos laterais mínimos, **sem
distorcer** a proporção e **sem** gerar página em branco (título seguinte com `page_break_before`).
Ver DEC-024 e `regras-layout.md`.

## Indicadores do texto (calcular da BASE, não do screenshot)
A partir de `entregas_stats.json` (consolidação): total; entregues e em elaboração (com %);
distribuição por **frente** (macro): Gestão de Engenharia/Meio Ambiente/Desapropriação (=G2),
Gestão de Projetos e Contabilidade (=G4), Assessoramento e Consultoria Jurídica (=G5), Coordenação
Executiva (=G1), Certificação de Projetos e Obras (=G3); por **empresa** (Dynatest/CG/Escritório Jurídico/CERT);
distribuição **temporal** (dia de pico). Conferir que os números do texto batem com o dashboard.

## Baseline (RM13/Julho-2026) — para conferência, recalcular a cada mês
*(Contagens omitidas na versão pública.)* Conferir: total = entregues + em elaboração; soma por frente =
soma por empresa = total; dia de pico e nº de documentos do pico.

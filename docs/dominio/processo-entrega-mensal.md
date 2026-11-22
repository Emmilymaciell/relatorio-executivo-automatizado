# Processo mensal de Entrega de Produtos Programa → Parceiro

> Fonte canônica: **POP v2026-10-01** + **Checklist Visual** + **Modelo de Índice** + **Fluxograma**
> (`docs/referencias-oficiais/v2026-10-01/`). Citações entre colchetes = onde está a regra.
> Este arquivo resume; em conflito, o documento oficial prevalece.

## 1. Atores
- **Grupos executores**: G2 (Dynatest – engenharia), G3 (Certificadora – certificação), G4 (CG – portfólio,
  financeiro, riscos), G5 (Escritório Jurídico – jurídico).
- **Grupo 1** (Coordenação Geral Executiva): P1.1 Gestão Executiva e P1.2 Relações Institucionais (Dynatest, feitos
  pela Coordenação); P1.3 Comunicação Integrada e P1.4 TIC (CG).
- **Coordenação Executiva**: consolida, confere, organiza o pacote e encaminha ao Parceiro.
- **Parceiro**: recebe o pacote pronto para protocolo (SEI).

## 2. Prazos [POP Quadro 1, Etapas 1/5/7; Checklist]
| Quem | O quê | Até |
|---|---|---|
| G2, G3, G4, G5 | Lista de Demandas; Relatório de Produtos Mensal **final, revisado e assinado em PDF**; todos os anexos em `ANEXOS`; sua subdivisão do Índice | **3º dia útil** do mês seguinte à competência |
| G1 – P1.3 e P1.4 (CG) | as mesmas exigências, para P1.3/P1.4 | **3º dia útil** |
| Coordenação – P1.1 e P1.2 | finalizar P1.1/P1.2, consolidar o G1, providenciar assinaturas, revisar a conformidade de todos, organizar e enviar | **5º dia útil** (consolidação no 4º e 5º) |

Dias úteis: a automação **estima** (seg–sex menos feriados nacionais de `config/feriados.json`); feriados estaduais /
Cidade não estão configurados — o resultado é sempre rotulado "estimado".

## 3. Particularidade do G1
P1.3/P1.4 entram no fluxo do 3º dia como qualquer grupo. P1.1/P1.2 são **produzidos pela Coordenação** no 4º/5º dia,
a partir das entregas dos demais: é a automação histórica deste repositório (Skill `relatorio-prog-grupo1`):
consolidar entregas → acumulado `P.1.1 - Controle de demandas Programa.xlsx` (só acrescenta) → anexos P1.1/P1.2
(Controle de Entregas, Dashboard, Quadro Financeiro, Timeline, Reuniões) → textos → Word → PDF → QA → Revisão final.
Depois o relatório vai para **assinatura** (externa: gov.br/ICP) e volta assinado.

## 4. Período da competência
Mês civil (01 → último dia) — DEC-048. O G2 é filtrado por "Nº Relatório" na Lista Mestre.

## 5. O que deve existir no PACOTE FINAL [POP Quadro 2, Etapa 6; Checklist Coord. 11–14]
```
NN. Mês AAAA/
  Índice de Arquivos - Entrega de Produtos Programa - Mês AAAA.docx   (cumulativo, G1…G5, na RAIZ do mês)
  G1/ … G5/
    <Relatório de Produtos Mensal FINAL ASSINADO>.pdf              (ÚNICO arquivo solto na raiz do grupo)
    ANEXOS/                                                        (pasta ÚNICA, sem subpastas)
      P1.3 - Comunicação Integrada - <nome do arquivo> - GERAL.pdf
      P1.1 - Gestão Executiva - Quadro Financeiro do mês de outubro - GERAL.zip   (planilha dentro do ZIP)
```
O pacote final **é a própria pasta do mês** em `Programa - 4. Produtos` (a usuária copia essa pasta para a pasta
compartilhada com o Parceiro) — DEC-045.

## 6. O que NÃO pode estar no pacote [POP Quadro 2/Etapa 6; Checklist Coord. 11–12]
- **Listas de Demandas** (instrumento interno — usadas na conferência e guardadas no arquivo interno).
- Word/minuta do relatório, PDF anterior/duplicado, "final 2", versões intermediárias **fora de ANEXOS**.
- Arquivos temporários/provisórios; arquivos soltos na raiz do grupo ou do mês (exceto Índice e relatório).
- Atenção: **dentro de ANEXOS** editáveis/ZIP legítimos **ficam** (o Checklist diz expressamente).

## 7. Anexos e nomenclatura [POP Quadro 3; Checklist itens 5–9]
`[CÓDIGO DO PRODUTO] - [NOME DO PRODUTO] - [NOME DO ARQUIVO] - [RODOVIA]`
- Rodovia: `RD-101`, `RD-102A`, `RD-103`, `RD-104` ou `GERAL` (mais de uma rodovia / conjunto do programa).
- Evitar nomes genéricos ("arquivo final", "novo", "versão 2", "cópia", "teste").
- Exemplos do POP: `P1.4 - TIC - Gestão Integrada - 2026-08-31 - GERAL`;
  `P2.2 - Análise de Projeto e RPFO - Nota Técnica de Aceitação de Documentos - NT-749.1 - RD-101`.
- Nomes curtos dos produtos: `config/entrega_pop.json` (de acordo com as pastas oficiais do SharePoint).

## 8. Formatos, ZIP e tamanho [POP Quadros 3/4; Etapa 2]
- PDF pesquisável, sem senha/restrição. **Todo arquivo que não for PDF vai em .ZIP**; o editável deve corresponder
  ao PDF quando aplicável (planilha: Excel com fórmulas dentro do ZIP + PDF correspondente).
- **Máx. 100 MB por arquivo**; maior → fracionar em `Parte 01`, `Parte 02`… e conferir que todas as partes existem.

## 9. Índice [Modelo de Índice r2; POP Etapa 3]
Título "ÍNDICE DE ARQUIVOS : ENTREGA DE PRODUTOS DO PROGRAMA". Tabela Nº | ARQUIVO | PROTOCOLO/SEI, subdividida por
GRUPO 1–5 (**só por grupo — sem subdivisão por produto**: o produto já está no nome de cada arquivo), cumulativa
na competência. **A numeração reinicia em 1 em cada grupo** (1, 2, 3, 4…), porque os grupos preenchem ao mesmo
tempo. Ordem dentro do grupo: Relatório de Produtos Mensal primeiro, depois os anexos por produto e nome (DEC-049). Cada grupo preenche a
sua subdivisão; a Coordenação confere. O nº SEI é atribuído **no protocolo** — nunca inventado; antes disso o campo
fica `- [SEI nº ____ ]`. A automação regenera o Índice do pacote a partir dos arquivos reais (preservando SEI já
preenchido) e exige reconciliação 100%.

## 10. Assinaturas [POP Quadro 4; Checklist item 2]
O Relatório de Produtos Mensal deve estar **assinado (digitalmente) em PDF**. Validação: ver
`docs/dominio/assinaturas.md`. Imagem de rubrica ≠ assinatura digital. Exigência específica de ICP-Brasil: não consta
no POP aprovado (DEC-046) — `exigir_icp_brasil` em config.

## 11. Quando a entrega está PRONTA [Checklist — regra final]
> "Se o Parceiro recebesse esta pasta agora, conseguiria protocolar tudo imediatamente, sem renomear, reorganizar ou
> solicitar correção?"

Na automação: **READY FOR PARCEIRO** só com **0 BLOCKER e 0 MANUAL_REVIEW** no pacote FINAL revalidado (pendências da
origem que o pacote não resolve também contam). MANUAL pode ser liberado por registro de revisão humana
(`revisoes_humanas.json`, preso ao SHA-256 do arquivo).

## 12. Responsabilidade da Coordenação (o que esta automação faz)
Confere recebimento e aderência (Etapa 4), sinaliza pendências para os grupos corrigirem, finaliza o G1, organiza o
pacote (Etapa 6) e **prepara** o envio (Etapa 7). **Não** envia, não protocola, não sobe no SEI, não publica nada no
espaço do Parceiro — isso é feito pela usuária.

## 13. Fluxo [Fluxograma]
Grupo prepara → valida (nome/formato/integridade/100 MB) → atualiza Índice → disponibiliza no SharePoint (3º dia) →
Coordenação confere → **há pendências?** sim: grupo corrige e reenvia / não: finaliza P1.1/P1.2 e consolida G1 →
assinaturas → organiza e revisa pacote → **pacote pronto?** não: ajusta / sim: envia ao Parceiro (5º dia).

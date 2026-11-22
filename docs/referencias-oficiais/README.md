# Referências oficiais (CANÔNICAS, somente leitura)

> **Versão pública:** os documentos oficiais do processo de entrega (POP, Checklist Visual, Modelo de Índice e
> Fluxograma) **não são publicados** — são documentos internos do contrato. No lugar do Modelo de Índice há um
> **modelo fictício mínimo** em `ficticio/Modelo de Indice - Entrega de Produtos (ficticio).docx`, com a mesma
> estrutura que `scripts/prog_entrega/indice.py` espera (parágrafo COMPETÊNCIA; tabela Nº | ARQUIVO | PROTOCOLO / SEI;
> linha mesclada "GRUPO N"; 4 linhas zebradas por grupo). Os scripts de `scripts/manutencao/carimbar_minuta.py`
> dependem dos originais e não rodam na versão pública.

No projeto original, estes documentos governam a entrega mensal ao Parceiro, são **imutáveis** (somente-leitura no
disco) e versionados por subpasta `vAAAA-MM-DD/` com SHA-256 registrado.

## Hierarquia de autoridade (em caso de conflito)
1. POP aprovado → 2. Checklist aprovado → 3. Modelo de Índice → 4. Fluxograma → 5. regras atuais do repositório
que não conflitem → 6. documentação antiga → 7. outputs antigos → 8. inferências. Conflito nunca é silenciado:
registrar em `docs/dominio/regras-pop-x-antigas.md` e no log de decisões (DEC-xxx).

## Onde cada regra virou código
| Regra | Implementação |
|---|---|
| limite 100 MB, rodovias, produtos, padrões de arquivos internos | `config/entrega_pop.json` |
| prazos 3º/5º dia útil | `config/entrega_pop.json` + `config/feriados.json` (estimativa) + `scripts/prog_entrega/calendario.py` |
| estrutura do mês / ANEXOS única | `scripts/prog_entrega/estrutura.py`, `validacao.py` |
| nomenclatura | `scripts/prog_entrega/nomes.py` |
| Índice | `scripts/prog_entrega/indice.py` (linhas copiadas do Modelo r2 — layout preservado; numeração por grupo) |
| integridade / PDF / ZIP / planilhas | `scripts/prog_entrega/arquivos.py` |
| assinaturas | `scripts/prog_entrega/assinatura.py` |
| checklist da Coordenação item a item | `scripts/prog_entrega/relatorio.py` (`CHECKLIST_COORD`) |
| pacote final / publicação | `scripts/prog_entrega/pacote.py`, `pipeline.py` |

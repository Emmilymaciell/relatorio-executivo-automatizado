# CHANGELOG

## [Não lançado] — 2026-10-02
### Alterado
- Índice no **Modelo r2** (`docs/referencias-oficiais/v2026-10-01-r2/`): título "ÍNDICE DE ARQUIVOS : ENTREGA DE
  PRODUTOS Programa", só subdivisões por GRUPO (sem produto), numeração reiniciando em 1 em cada grupo; nome do arquivo
  `Índice de Arquivos - Entrega de Produtos Programa - {Mês} {ano}.docx` (DEC-049). Comando `indice` substitui índice
  em branco de nome/layout anterior (com cópia) e nunca toca índice preenchido. Testes: 50.

## [2.0.0] — 2026-10-01 (tag v2.0.0) — "Entrega de Produtos - Coordenação"
### Adicionado
- Pipeline da Coordenação (`scripts/entrega.py`, `scripts/prog_entrega/`): inventário, validação (PDF/ZIP/Office pelo
  tipo real, assinatura digital com verificação criptográfica, nomes do POP, 100 MB, fragmentos, duplicidades,
  versões, Índice), staging, pacote final, revalidação, manifesto SHA-256, relatório de prontidão, publicação com
  arquivamento e código de confirmação; padronização de pastas mensais e Índice a partir do Modelo aprovado.
- Referências oficiais do POP em `docs/referencias-oficiais/v2026-10-01/` (read-only, SHA-256).
- Skills `entrega-coordenacao`, `session-start`, `session-close`; `docs/` (domínio, arquitetura, uso, estado, handoff).
- Testes pytest (48) com fixtures artificiais (inclui PDF assinado de verdade com certificado de teste).
### Alterado
- G1: layout de anexos `plano` (POP) a partir de Out/2026 com nomes padronizados; "Relatório de Produtos Mensal";
  Lista de Demandas também procurada no arquivo interno (DEC-043/044/045).
### Corrigido (descoberto na calibração com agosto/2026 real)
- Relatórios com `_` no nome não eram reconhecidos; CMS BER truncado por `rstrip`; ICP-Brasil por substring;
  PDF+ZIP de mesmo nome tratados como "versões"; caminhos > 260 caracteres.

## [1.x] — 2026-08-11 … 2026-09-27
Automação do Relatório Mensal do Grupo 1 (tags `v1.0`…`v1.4`; commits até "v1.17"). Histórico detalhado em
`git log` e DEC-001…042.

# scripts/ (nível do projeto)

- `entrega.py` — **CLI da Coordenação** (POP de Entrega de Produtos): `estrutura`, `indice`, `auditar`, `preparar`,
  `publicar`. Ver `docs/COMO-USAR.md` e a Skill `.claude/skills/entrega-coordenacao/`.
- `prog_entrega/` — módulos da pipeline (inventário, validação, assinatura, nomes, índice, pacote, relatório...).
  Arquitetura: `docs/ARQUITETURA.md`.
- `sharepoint.py` + `prog_sharepoint/` — acesso ao SharePoint do Programa pela **Microsoft Graph** (app-only):
  `testar`, `listar`, `info`, `baixar`, `comparar`, `enviar` (dry-run; backup + eTag + SHA-256). DEC-052,
  `docs/COMO-USAR.md` § 8.

Os scripts CANÔNICOS do **Relatório de Produtos Mensal do G1** ficam na Skill:
`.claude/skills/relatorio-prog-grupo1/scripts/` (parametrizados por `config/parametros.json` via `prog_config.py`).

# -*- coding: utf-8 -*-
"""Acesso ao SharePoint do Programa pela Microsoft Graph (app-only) — DEC-052.

Só biblioteca padrão (urllib); nenhum segredo no repositório. Módulos:
  config    — config/sharepoint.json + credenciais (env PROG_SP_* ou arquivo fora do repo)
  caminhos  — local sincronizado <-> caminho na biblioteca; bloqueios de escrita
  graph     — cliente HTTP (token client_credentials, retry 429/5xx, paginação, upload em sessão)
  operacoes — testar, listar, info, baixar, comparar, enviar (dry-run por padrão; backup + If-Match ao substituir)

Por desenho NÃO existe excluir, mover nem renomear.
"""

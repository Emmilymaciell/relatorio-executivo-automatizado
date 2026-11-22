# -*- coding: utf-8 -*-
"""prog_entrega — automação da etapa da COORDENAÇÃO EXECUTIVA (POP de Entrega de Produtos Programa).

Fluxo: inventário -> validação -> reconciliação -> correções seguras (staging) -> revalidação ->
pacote final -> revalidação final -> manifesto -> relatório de prontidão. Nunca envia nada ao Parceiro.
Ponto de entrada: scripts/entrega.py (ver docs/COMO-USAR.md).
"""

# Sanitização desta versão pública

Cópia sanitizada de um projeto em uso real. Nada aqui é dado de cliente ou de terceiro: o que aparece como exemplo é inventado. Detalhe completo da sanitização: relatório interno da autora.

## O que foi removido

Planilhas, documentos e índices reais; modelo oficial do Índice (trocado por modelo fictício); valores financeiros.

## O que foi trocado

Programa, parceiro, órgãos e rodovias → nomes genéricos; **códigos de documento da Lista Mestre → fictícios** (deslocamento fixo de +500 por prefixo, preservando ordem e faixas); valores → fictícios.

## O que é fictício

`docs/referencias-oficiais/ficticio/` (modelo de Índice fictício) e fixtures artificiais geradas em `tmp_path`.

## Como verificar

```bat
py -3.12 -m pytest tests -q        # 103 testes
```

Varreduras (2026-11-19): gitleaks `dir` só com falsos positivos justificados; busca por caminho local, e-mail, nomes do cliente e links privados sem ocorrências reais.

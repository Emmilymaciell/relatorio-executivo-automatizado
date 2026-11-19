# Validação de assinaturas (implementação: `scripts/prog_entrega/assinatura.py`)

| Classe | Significado | Status na auditoria (relatório do grupo) |
|---|---|---|
| **VALIDADA** | há dicionário de assinatura (ByteRange) e: hash do conteúdo coberto = messageDigest assinado **e** a assinatura confere com a chave do certificado do signatário **e** a última assinatura cobre o arquivo inteiro (ou o acréscimo posterior é só LTV: DSS/VRI/metadados) | PASS |
| **ASSINATURA DIGITAL INVÁLIDA** | estrutura presente mas hash/assinatura não conferem (alterado após assinar ou corrompido) | BLOCKER |
| **ASSINATURA DIGITAL NÃO VERIFICÁVEL NO AMBIENTE** | formato/algoritmo não suportado, CMS ilegível, ou o arquivo foi **alterado depois da assinatura** com conteúdo (página/imagem/anotação — ex.: "Preencher e Assinar") | MANUAL_REVIEW |
| **POSSÍVEL ASSINATURA VISUAL — REVISÃO HUMANA** | sem estrutura digital, mas o texto tem indícios ("assinado digitalmente", gov.br, DocuSign…) — típico de PDF impresso/achatado | MANUAL_REVIEW |
| **NÃO ENCONTRADA** | nenhum indício | BLOCKER |

## O que NÃO é verificado (limitação declarada)
- **Cadeia de confiança e revogação** (CRL/OCSP) — exigiriam as raízes ICP-Brasil/gov.br e acesso à rede. O emissor é
  informado (ex.: "AC Final do Governo Federal do Brasil v1 (Gov-Br)", "AC Certisign RFB G5 (ICP-Brasil)").
- `icp_brasil` = o certificado do signatário (ou da cadeia embutida) tem `O=ICP-Brasil`. gov.br **não** é ICP-Brasil.
- Formato `adbe.pkcs7.sha1` e CMS sem atributos assinados → NÃO VERIFICÁVEL.

## Revisão humana
Um MANUAL_REVIEW pode ser aceito registrando em `output/ENTREGA_NN_Mes_AAAA/revisoes_humanas.json`:
`[{"sha256": "<do arquivo>", "codigo": "assinatura_relatorio", "decisao": "aceito", "por": "...", "em": "AAAA-MM-DD",
"justificativa": "..."}]`. Vale só para aquele conteúdo: se o arquivo mudar, a confirmação deixa de valer.

## Evidência de calibração (01/10/2026, `14. Agosto 2026`, somente leitura)
gov.br (G2, ART, parecer Parceiro) e ICP-Brasil (Carta 012) VALIDADAS; DocuSign+LTV (G3) VALIDADA; G1 assinado e depois
modificado por "Preencher e Assinar" → NÃO VERIFICÁVEL; G4 sem assinatura → NÃO ENCONTRADA; G5 → VISUAL. Testes
com assinatura real gerada por certificado de teste + adulteração: `tests/test_entrega.py::TestAssinatura`.

# -*- coding: utf-8 -*-
"""Detecção e verificação de assinatura digital em PDF (PAdES / adbe.pkcs7.detached / ETSI.CAdES.detached).

Classificação (nunca gera falso positivo):
  VALIDADA                     – existe assinatura digital; o hash do conteúdo coberto (ByteRange) bate com o
                                 messageDigest assinado E a assinatura criptográfica confere com a chave do
                                 certificado do signatário; a última assinatura cobre o arquivo inteiro.
                                 (Cadeia de confiança/revogação ICP-Brasil NÃO é verificada offline — o
                                 emissor é apenas informado; ver docs/dominio/assinaturas.md.)
  ASSINATURA DIGITAL INVÁLIDA  – estrutura presente, mas hash/assinatura não conferem (arquivo alterado após
                                 assinar, ou assinatura corrompida).
  ASSINATURA DIGITAL NÃO VERIFICÁVEL NO AMBIENTE – estrutura presente, mas formato/algoritmo não suportado
                                 aqui (ex.: adbe.pkcs7.sha1) ou o CMS não pôde ser interpretado.
  POSSÍVEL ASSINATURA VISUAL — REVISÃO HUMANA – sem estrutura digital, mas há indício visual/textual
                                 (ex.: "assinado digitalmente", gov.br, DocuSign) — pode ser cópia impressa.
  NÃO ENCONTRADA               – nenhum indício.
"""
import hashlib
import re

VALIDADA = "VALIDADA"
INVALIDA = "ASSINATURA DIGITAL INVÁLIDA"
NAO_VERIFICAVEL = "ASSINATURA DIGITAL NÃO VERIFICÁVEL NO AMBIENTE"
VISUAL = "POSSÍVEL ASSINATURA VISUAL — REVISÃO HUMANA"
NAO_ENCONTRADA = "NÃO ENCONTRADA"
VERSAO = "assinatura-2"

_RE_SIG = re.compile(rb"/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]")
_RE_SUB = re.compile(rb"/SubFilter\s*/([A-Za-z0-9.\-_]+)")
_INDICIOS = ["assinado digitalmente", "assinado de forma digital", "assinatura digital", "documento assinado",
             "docusign", "assinado eletronicamente", "assinatura eletronica", "gov.br", "clicksign",
             "d4sign", "autentique", "icp-brasil", "codigo verificador", "código verificador"]

OID_SHA = {"1.3.14.3.2.26": "sha1", "2.16.840.1.101.3.4.2.1": "sha256",
           "2.16.840.1.101.3.4.2.2": "sha384", "2.16.840.1.101.3.4.2.3": "sha512"}
OID_MESSAGE_DIGEST = "1.2.840.113549.1.9.4"
OID_SIGNING_TIME = "1.2.840.113549.1.9.5"


# ------------------------------------------------------------------- DER mínimo (somente leitura)
def _der(b, i):
    """(tag, inicio_conteudo, fim_conteudo, fim_total) do TLV em b[i]."""
    tag = b[i]
    j = i + 1
    ln = b[j]
    j += 1
    if ln == 0x80:  # BER comprimento indefinido: conteúdo termina no EOC (00 00)
        k = j
        while not (b[k] == 0 and b[k + 1] == 0):
            k = _der(b, k)[3]
        return tag, j, k, k + 2
    if ln & 0x80:
        n = ln & 0x7F
        ln = int.from_bytes(b[j:j + n], "big")
        j += n
    return tag, j, j + ln, j + ln


def _filhos(b, ini, fim):
    out, i = [], ini
    while i < fim:
        t, ci, cf, f = _der(b, i)
        out.append((t, i, ci, cf, f))
        i = f
    return out


def _oid(b):
    vals, v = [], 0
    for x in b:
        v = (v << 7) | (x & 0x7F)
        if not x & 0x80:
            vals.append(v)
            v = 0
    primeiro = vals[0]
    return ".".join(map(str, [min(primeiro // 40, 2), primeiro - 40 * min(primeiro // 40, 2)] + vals[1:]))


def _parse_cms(cms):
    """Extrai do SignedData: certificados (DER), digestAlgorithm, signedAttrs (DER), assinatura,
    algoritmo de assinatura e o serial do signatário."""
    _t, ci, cf, _ = _der(cms, 0)  # ContentInfo
    ch = _filhos(cms, ci, cf)
    if _oid(cms[ch[0][2]:ch[0][3]]) != "1.2.840.113549.1.7.2":
        raise ValueError("ContentInfo não é signedData")
    _t, sci, scf, _ = _der(cms, ch[1][2])  # [0] EXPLICIT -> SignedData SEQUENCE
    sd = _filhos(cms, sci, scf)
    certs, signer_infos = [], None
    for t, i0, c0, c1, f in sd:
        if t == 0xA0:  # [0] IMPLICIT certificates
            certs = [cms[a:e] for (_tt, a, _c, _cc, e) in _filhos(cms, c0, c1) if _tt == 0x30]
        if t == 0x31:
            signer_infos = (c0, c1)
    if not signer_infos:
        raise ValueError("SignedData sem signerInfos")
    sis = _filhos(cms, *signer_infos)
    _t, a, c0, c1, f = sis[0]
    si = _filhos(cms, c0, c1)
    # SignerInfo: version, sid, digestAlgorithm, [0] signedAttrs?, signatureAlgorithm, signature
    k = 1
    sid = si[k]
    serial = None
    if sid[0] == 0x30:  # issuerAndSerialNumber
        partes = _filhos(cms, sid[2], sid[3])
        serial = int.from_bytes(cms[partes[1][2]:partes[1][3]], "big")
    k += 1
    dig_alg = _oid(cms[_filhos(cms, si[k][2], si[k][3])[0][2]:_filhos(cms, si[k][2], si[k][3])[0][3]])
    k += 1
    signed_attrs = None
    if si[k][0] == 0xA0:
        signed_attrs = bytearray(cms[si[k][1]:si[k][4]])
        signed_attrs[0] = 0x31  # para verificar, o [0] IMPLICIT vira SET OF
        signed_attrs = bytes(signed_attrs)
        attrs = _filhos(cms, si[k][2], si[k][3])
        k += 1
    else:
        attrs = []
    sig_alg = _oid(cms[_filhos(cms, si[k][2], si[k][3])[0][2]:_filhos(cms, si[k][2], si[k][3])[0][3]])
    k += 1
    signature = cms[si[k][2]:si[k][3]]
    md, stime = None, None
    for _t, a0, c0_, c1_, _f in attrs:
        p = _filhos(cms, c0_, c1_)
        oid = _oid(cms[p[0][2]:p[0][3]])
        vals = _filhos(cms, p[1][2], p[1][3])
        if oid == OID_MESSAGE_DIGEST:
            md = cms[vals[0][2]:vals[0][3]]
        elif oid == OID_SIGNING_TIME:
            stime = cms[vals[0][2]:vals[0][3]].decode("ascii", "replace")
    return {"certs": certs, "dig_alg": dig_alg, "signed_attrs": signed_attrs, "sig_alg": sig_alg,
            "signature": signature, "message_digest": md, "serial": serial, "signing_time": stime}


def _verificar_assinatura(cert, sig_alg, signature, dados, hash_name):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.asymmetric import ec, padding, rsa
    h = {"sha1": hashes.SHA1(), "sha256": hashes.SHA256(), "sha384": hashes.SHA384(),
         "sha512": hashes.SHA512()}[hash_name]
    pub = cert.public_key()
    if isinstance(pub, rsa.RSAPublicKey):
        if sig_alg == "1.2.840.113549.1.1.10":
            pub.verify(signature, dados, padding.PSS(mgf=padding.MGF1(h), salt_length=padding.PSS.AUTO), h)
        else:
            pub.verify(signature, dados, padding.PKCS1v15(), h)
    elif isinstance(pub, ec.EllipticCurvePublicKey):
        pub.verify(signature, dados, ec.ECDSA(h))
    else:
        raise ValueError("tipo de chave não suportado")


def _nome(n):
    from cryptography.x509.oid import NameOID
    def g(oid):
        v = n.get_attributes_for_oid(oid)
        return v[0].value if v else ""
    return g(NameOID.COMMON_NAME), g(NameOID.ORGANIZATION_NAME)


def _uma(raw, br, contents_hex):
    from cryptography import x509
    a, b, c, d = br
    dados = raw[a:a + b] + raw[c:c + d]
    cms = bytes.fromhex(re.sub(rb"\s", b"", contents_hex).decode("ascii"))
    # NÃO usar rstrip(b"\x00"): em BER (comprimento indefinido) o CMS termina com EOCs 00 00, que se
    # confundem com o preenchimento. Delimita-se pelo próprio TLV externo.
    cms = cms[:_der(cms, 0)[3]]
    p = _parse_cms(cms)
    hn = OID_SHA.get(p["dig_alg"])
    if not hn:
        return {"status": NAO_VERIFICAVEL, "motivo": f"algoritmo de digest não suportado ({p['dig_alg']})"}
    certs = [x509.load_der_x509_certificate(c_) for c_ in p["certs"]]
    signer = next((c_ for c_ in certs if p["serial"] is not None and c_.serial_number == p["serial"]),
                  certs[0] if len(certs) == 1 else None)
    det = {"signing_time": p["signing_time"]}
    if signer is not None:
        cn, org = _nome(signer.subject)
        icn, iorg = _nome(signer.issuer)
        det.update({"signatario": cn, "organizacao": org, "emissor": icn, "emissor_org": iorg,
                    # certificados da ICP-Brasil têm O=ICP-Brasil (igualdade, não substring)
                    "icp_brasil": any((x or "").strip().lower() == "icp-brasil" for x in
                                      [iorg] + [_nome(c_.issuer)[1] for c_ in certs] + [_nome(c_.subject)[1] for c_ in certs])})
    digest = hashlib.new(hn, dados).digest()
    if p["signed_attrs"] is None:
        return {"status": NAO_VERIFICAVEL, "motivo": "CMS sem atributos assinados (formato não suportado)", **det}
    if p["message_digest"] != digest:
        return {"status": INVALIDA, "motivo": "hash do conteúdo NÃO confere com o assinado (documento alterado "
                                              "após a assinatura ou assinatura corrompida)", **det}
    if signer is None:
        return {"status": NAO_VERIFICAVEL, "motivo": "certificado do signatário não localizado no CMS", **det}
    try:
        _verificar_assinatura(signer, p["sig_alg"], p["signature"], p["signed_attrs"], hn)
    except Exception as e:
        if type(e).__name__ == "InvalidSignature":
            return {"status": INVALIDA, "motivo": "assinatura criptográfica não confere com o certificado", **det}
        return {"status": NAO_VERIFICAVEL, "motivo": f"verificação não suportada: {e}", **det}
    return {"status": VALIDADA, "motivo": "integridade e assinatura criptográfica conferidas "
                                          "(cadeia/revogação não verificadas offline)", **det}


def verificar_pdf(path):
    """Retorna {"status", "motivo", "assinaturas": [...], "subfilters": [...]} para o PDF."""
    with open(path, "rb") as f:
        raw = f.read()
    tam = len(raw)
    sigs = []
    for m in _RE_SIG.finditer(raw):
        br = tuple(int(x) for x in m.groups())
        # /SubFilter e /Contents do MESMO objeto de assinatura (entre "N 0 obj" e "endobj")
        ini_obj = raw.rfind(b" obj", 0, m.start())
        fim_obj = raw.find(b"endobj", m.end())
        obj = raw[ini_obj:fim_obj if fim_obj > 0 else m.end() + 4000]
        ms = _RE_SUB.search(obj)
        sub = ms.group(1).decode() if ms else None
        mc = re.search(rb"/Contents\s*<([0-9A-Fa-f\s]+)>", obj)
        if not mc:
            sigs.append({"status": NAO_VERIFICAVEL, "motivo": "ByteRange sem /Contents legível", "byterange": br,
                         "cobre_arquivo_inteiro": False})
            continue
        contents = mc.group(1)
        if sub == "ETSI.RFC3161":  # carimbo de tempo do documento — não é assinatura de pessoa
            sigs.append({"status": "CARIMBO_DE_TEMPO", "motivo": "carimbo de tempo (RFC 3161)", "byterange": br,
                         "subfilter": sub, "cobre_arquivo_inteiro": br[0] == 0 and br[2] + br[3] == tam})
            continue
        try:
            r = _uma(raw, br, contents)
        except Exception as e:
            r = {"status": NAO_VERIFICAVEL, "motivo": f"CMS não interpretável: {e}"}
        r["byterange"] = br
        r["subfilter"] = sub
        r["cobre_arquivo_inteiro"] = br[0] == 0 and br[2] + br[3] == tam
        if not r["cobre_arquivo_inteiro"]:
            r["acrescimo_posterior"] = _analisar_acrescimo(raw[br[2] + br[3]:])
        sigs.append(r)
    subs = sorted({s.decode() for s in _RE_SUB.findall(raw)})
    if not sigs:
        return {"status": _indicio_visual(path), "assinaturas": [], "subfilters": subs,
                "motivo": "nenhum dicionário de assinatura digital (/ByteRange) no PDF"}
    carimbos = [s for s in sigs if s["status"] == "CARIMBO_DE_TEMPO"]
    sigs = [s for s in sigs if s["status"] != "CARIMBO_DE_TEMPO"]
    if not sigs:
        return {"status": NAO_VERIFICAVEL, "assinaturas": carimbos, "subfilters": subs,
                "motivo": "só há carimbo de tempo de documento, sem assinatura de pessoa"}
    validas = [s for s in sigs if s["status"] == VALIDADA]
    if any(s["status"] == INVALIDA for s in sigs):
        st, mot = INVALIDA, next(s["motivo"] for s in sigs if s["status"] == INVALIDA)
    elif validas and (any(s["cobre_arquivo_inteiro"] for s in validas)
                      or any(c["cobre_arquivo_inteiro"] for c in carimbos)):
        st, mot = VALIDADA, validas[0]["motivo"]
    elif validas and validas[-1].get("acrescimo_posterior", {}).get("somente_ltv"):
        st, mot = VALIDADA, validas[-1]["motivo"] + "; acréscimo posterior só de validação de longo prazo (DSS/LTV)"
    elif validas:
        tipos = validas[-1].get("acrescimo_posterior", {}).get("tipos", [])
        st, mot = NAO_VERIFICAVEL, ("assinatura íntegra, mas o arquivo foi ALTERADO após a assinatura "
                                    f"(acrescentados: {', '.join(tipos) or 'conteúdo'}) — ex.: rubrica/imagem por "
                                    "'Preencher e Assinar'. Revisar/reassinar")
    else:
        st, mot = NAO_VERIFICAVEL, sigs[0]["motivo"]
    return {"status": st, "motivo": mot, "assinaturas": sigs + carimbos, "subfilters": subs}


# Tipos que um acréscimo de validação de longo prazo (LTV) pode conter sem alterar o conteúdo visível.
# StructTreeRoot = árvore de marcação/acessibilidade reescrita por plataformas como a DocuSign.
_TIPOS_LTV = {"Catalog", "Metadata", "XRef", "ObjStm", "DSS", "VRI", "DocTimeStamp", "Sig", "StructTreeRoot"}


def _analisar_acrescimo(tail):
    """O que foi acrescentado APÓS a assinatura? Só DSS/VRI/metadados/catálogo = LTV (legítimo).
    Páginas, anotações, imagens, fontes = alteração de conteúdo. Object streams são descompactados."""
    import zlib
    tipos = set(t.decode() for t in re.findall(rb"/Type\s*/(\w+)", tail))
    for mm in re.finditer(rb"/Type\s*/ObjStm.*?stream\r?\n", tail, re.S):
        ini = mm.end()
        fim = tail.find(b"endstream", ini)
        try:
            tipos |= set(t.decode() for t in re.findall(rb"/Type\s*/(\w+)", zlib.decompress(tail[ini:fim])))
        except Exception:
            tipos.add("ObjStm-ilegível")
    tem_dss = b"/DSS" in tail
    return {"tipos": sorted(tipos), "tem_dss": tem_dss,
            "somente_ltv": tem_dss and tipos <= _TIPOS_LTV}


def _indicio_visual(path):
    try:
        import fitz
        doc = fitz.open(path)
        n = doc.page_count
        paginas = list(range(n)) if n <= 6 else list(range(3)) + list(range(n - 3, n))
        txt = " ".join(doc.load_page(i).get_text() for i in paginas).lower()
        doc.close()
    except Exception:
        return NAO_ENCONTRADA
    import unicodedata
    t = "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")
    return VISUAL if any(k in t for k in _INDICIOS) else NAO_ENCONTRADA

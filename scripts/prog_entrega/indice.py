# -*- coding: utf-8 -*-
"""Índice de Arquivos : Entrega de Produtos Programa (Modelo aprovado v2026-10-01-r2 — numeração POR GRUPO).

O layout NÃO é recriado: toda linha nova é cópia (deepcopy) de uma linha-protótipo do próprio Modelo
(cabeçalho 'Nº | ARQUIVO | PROTOCOLO / SEI', linha de grupo 'GRUPO N' mesclada, linha de arquivo).
O Modelo orienta: "registrar os arquivos em ordem dentro da subdivisão do respectivo grupo. O índice é
cumulativo na competência mensal; adicionar novas linhas conforme necessário."

Funções:
  criar_modelo_mes(...)  -> índice "em branco" do mês: competência + período; GRUPO 1..5 com as linhas
                            numeradas 1..4 do Modelo (sem subdivisão por produto — DEC-049).
  ler_indice(path)       -> estrutura lida (grupo, produto, nº, arquivo, SEI) para reconciliação.
  escrever_indice(...)   -> índice preenchido a partir da lista REAL de arquivos (preservando SEI).
"""
import copy
import os
import re

import docx
from docx.oxml.ns import qn

from . import regras as R

SEI_VAZIO = "- [SEI nº __________________ ]"
_RE_SEI = re.compile(r"SEI\s*n[ºo°]?\s*[:\-]?\s*([0-9][0-9.\-/]{3,})", re.I)
_RE_GRUPO = re.compile(r"^\s*GRUPO\s+([1-5])\s*$", re.I)
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _modelo_path():
    return R.resolve_path(R.config()["indice"]["modelo"])


def nome_arquivo_indice(mes, ano):
    return R.config()["indice"]["nome_arquivo"].format(mes=mes, ano=ano)


def _tabela(d):
    if not d.tables:
        raise ValueError("Índice sem tabela — não corresponde ao Modelo aprovado.")
    return d.tables[0]


def _cells_unicos(tr):
    return tr.findall(f"{_W}tc")


def _prototipos(tbl):
    """(header_tr, grupo_tr, [arquivo_tr_impar, arquivo_tr_par]) a partir do Modelo (o Modelo alterna
    o fundo das linhas de arquivo — zebrado — e isso é preservado)."""
    rows = tbl._tbl.findall(f"{_W}tr")
    header = rows[0]
    grupo = next(r for r in rows if len(_cells_unicos(r)) == 1)
    arquivos = [r for r in rows[1:] if len(_cells_unicos(r)) == 3][:2]
    return header, grupo, arquivos


def _set_cell_text(tc, texto, bold=None):
    """Troca o texto mantendo o rPr do último run com texto (formatação do Modelo)."""
    ps = tc.findall(f"{_W}p")
    p = ps[0]
    for extra in ps[1:]:
        tc.remove(extra)
    runs = p.findall(f"{_W}r")
    keep = None
    for r in runs:
        if r.find(f"{_W}t") is not None:
            keep = r
    if keep is None and runs:
        keep = runs[-1]
    for r in runs:
        if r is not keep:
            p.remove(r)
    if keep is None:
        keep = p.makeelement(f"{_W}r", {})
        p.append(keep)
    for t in keep.findall(f"{_W}t"):
        keep.remove(t)
    t = keep.makeelement(f"{_W}t", {})
    t.text = texto
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    keep.append(t)
    if bold is not None:
        rpr = keep.find(f"{_W}rPr")
        if rpr is None:
            rpr = keep.makeelement(f"{_W}rPr", {})
            keep.insert(0, rpr)
        for b in rpr.findall(f"{_W}b"):
            rpr.remove(b)
        b = rpr.makeelement(f"{_W}b", {})
        if not bold:
            b.set(qn("w:val"), "0")
        rpr.insert(0, b)


def _set_fill(tc, fill):
    tcpr = tc.find(f"{_W}tcPr")
    shd = tcpr.find(f"{_W}shd") if tcpr is not None else None
    if shd is not None:
        shd.set(qn("w:fill"), fill)


def _linha_grupo(proto, texto):
    tr = copy.deepcopy(proto)
    _set_cell_text(_cells_unicos(tr)[0], texto, bold=True)
    return tr


def _linha_arquivo(proto, num="", arquivo="", sei=SEI_VAZIO):
    tr = copy.deepcopy(proto)
    c = _cells_unicos(tr)
    _set_cell_text(c[0], str(num))
    _set_cell_text(c[1], arquivo)
    _set_cell_text(c[2], sei)
    return tr


def _preencher_competencia(d, texto):
    for p in d.paragraphs:
        if "COMPET" in p.text.upper():
            runs = p.runs
            alvo = next((r for r in runs if "_" in r.text), None)
            if alvo is not None:
                alvo.text = re.sub(r"_+", texto, alvo.text, count=1)
                for r in runs:
                    if r is not alvo and set(r.text) <= {"_"} and r.text:
                        r.text = ""
            else:
                p.add_run(" " + texto)
            return
    raise ValueError("Parágrafo 'COMPETÊNCIA' não encontrado no Modelo.")


def _linhas_por_grupo_no_modelo(tbl):
    """Quantas linhas de arquivo o Modelo traz em cada GRUPO (Modelo r2: 4 — numeradas 1 a 4)."""
    n, contando, cont = 0, False, []
    for tr in tbl._tbl.findall(f"{_W}tr"):
        tcs = _cells_unicos(tr)
        if len(tcs) == 1:
            if contando:
                cont.append(n)
            contando, n = True, 0
        elif contando and len(tcs) == 3:
            n += 1
    if contando:
        cont.append(n)
    return max(cont) if cont else 4


def _montar(d, secoes, competencia_txt):
    """secoes: lista de (grupo, [linhas (num, arquivo, sei)]). Só linhas de GRUPO e de arquivo, como no Modelo
    (sem subdivisão por produto — o produto está no nome do arquivo; usuária, 01/10/2026, DEC-049)."""
    _preencher_competencia(d, competencia_txt)
    tbl = _tabela(d)
    header, p_grupo, p_arqs = _prototipos(tbl)
    p_grupo, p_arqs = copy.deepcopy(p_grupo), [copy.deepcopy(x) for x in p_arqs]
    for tr in tbl._tbl.findall(f"{_W}tr"):
        if tr is not header:
            tbl._tbl.remove(tr)
    for grupo, linhas in secoes:
        tbl._tbl.append(_linha_grupo(p_grupo, f"GRUPO {grupo[1:]}"))
        for k, (num, arq, sei) in enumerate(linhas):  # zebrado reinicia em cada grupo, como no Modelo
            tbl._tbl.append(_linha_arquivo(p_arqs[k % len(p_arqs)], num, arq, sei or SEI_VAZIO))


def _competencia(mes, ano, periodo):
    return f"{mes}/{ano}" + (f" (período: {periodo[0]} a {periodo[1]})" if periodo else "")


def criar_modelo_mes(destino, mes, ano, periodo=None, modelo=None, sobrescrever=False):
    """Índice do mês SEM dados de anexos: só competência/período; em cada GRUPO as linhas numeradas 1…N do
    Modelo (N = linhas por grupo do próprio Modelo). Nunca sobrescreve um índice existente."""
    if os.path.exists(destino) and not sobrescrever:
        return False
    d = docx.Document(modelo or _modelo_path())
    n = _linhas_por_grupo_no_modelo(_tabela(d))
    secoes = [(g, [(i, "", SEI_VAZIO) for i in range(1, n + 1)]) for g in R.config()["grupos"]]
    _montar(d, secoes, _competencia(mes, ano, periodo))
    os.makedirs(os.path.dirname(destino) or ".", exist_ok=True)
    d.save(destino)
    return True


def escrever_indice(destino, mes, ano, entradas, periodo=None, sei_por_arquivo=None, modelo=None):
    """entradas: {grupo: [(produto_ou_'REL', nome_arquivo)]} já na ORDEM final (relatório primeiro, depois
    anexos por produto e nome). Numeração REINICIA em 1 em cada grupo (DEC-049). SEI preservado por nome do
    arquivo; ausente => campo vazio do Modelo (nunca inventado). Retorna o total de arquivos."""
    sei_por_arquivo = sei_por_arquivo or {}
    d = docx.Document(modelo or _modelo_path())
    secoes, total = [], 0
    for g in R.config()["grupos"]:
        linhas = []
        for n, (_sub, arq) in enumerate(entradas.get(g, []), 1):
            sei = sei_por_arquivo.get(R.norm(arq))
            linhas.append((n, arq, f"- [SEI nº {sei} ]" if sei else SEI_VAZIO))
        total += len(linhas)
        secoes.append((g, linhas))
    _montar(d, secoes, _competencia(mes, ano, periodo))
    d.save(destino)
    return total


def assinatura_tabela(path):
    """Textos das linhas da tabela (para saber se um índice em branco está no layout vigente)."""
    d = docx.Document(path)
    return [tuple("".join(t.text or "" for t in tc.iter(f"{_W}t")).strip() for tc in _cells_unicos(tr))
            for tr in _tabela(d)._tbl.findall(f"{_W}tr")]


def eh_indice_em_branco(path):
    """True se o índice não tem NENHUM arquivo registrado nem SEI (só o esqueleto gerado)."""
    r = ler_indice(path)
    return not r["erros"] and not r["linhas"]


def ler_indice(path):
    """Lê o índice: {'competencia': str, 'linhas': [{grupo, subdivisao, num, arquivo, sei}], 'erros': []}."""
    out = {"competencia": None, "linhas": [], "erros": []}
    try:
        d = docx.Document(path)
    except Exception as e:  # docx corrompido
        out["erros"].append(f"não abre como .docx: {e}")
        return out
    for p in d.paragraphs:
        if "COMPET" in p.text.upper():
            out["competencia"] = p.text.split(":", 1)[-1].strip()
            break
    if not d.tables:
        out["erros"].append("sem tabela")
        return out
    grupo = sub = None
    for tr in d.tables[0]._tbl.findall(f"{_W}tr"):
        tcs = _cells_unicos(tr)
        textos = ["".join(t.text or "" for t in tc.iter(f"{_W}t")).strip() for tc in tcs]
        if len(tcs) == 1:
            m = _RE_GRUPO.match(textos[0])
            if m:
                grupo, sub = f"G{m.group(1)}", None
            else:
                sub = textos[0]
            continue
        if len(tcs) >= 3 and R.norm(textos[1]) == "arquivo":
            continue  # cabeçalho
        if len(tcs) >= 3:
            num, arq, sei_txt = textos[0], textos[1], textos[2]
            if not arq:
                continue
            m = _RE_SEI.search(sei_txt)
            out["linhas"].append({"grupo": grupo, "subdivisao": sub, "num": num, "arquivo": arq,
                                  "sei": m.group(1) if m else None})
    return out

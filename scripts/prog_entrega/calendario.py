# -*- coding: utf-8 -*-
"""Nº dia útil ESTIMADO (seg–sex menos feriados de config/feriados.json).

Não é calendário oficial: estaduais/municipais não estão configurados e a lista não foi confirmada pela
usuária (campo 'confirmado_pela_usuaria'). Por isso todo resultado é rotulado como "estimado"."""
import datetime
import json
import os

from . import regras as R


def _feriados():
    p = R.resolve_path(R.config()["feriados_arquivo"])
    if not os.path.exists(p):
        return set(), False
    d = json.load(open(p, encoding="utf-8"))
    fs = set()
    for bloco in ("nacionais", "estaduais_municipais"):
        for _ano, lst in (d.get(bloco) or {}).items():
            fs |= {datetime.date.fromisoformat(x) for x in lst}
    return fs, bool(d.get("confirmado_pela_usuaria"))


def nesimo_dia_util(ano, mes, n):
    fs, _ = _feriados()
    d = datetime.date(ano, mes, 1)
    k = 0
    while True:
        if d.weekday() < 5 and d not in fs:
            k += 1
            if k == n:
                return d
        d += datetime.timedelta(days=1)


def prazos_da_competencia(mes_num, ano):
    """A entrega da competência M ocorre no mês M+1."""
    m, a = (1, ano + 1) if mes_num == 12 else (mes_num + 1, ano)
    pz = R.config()["prazos"]
    _fs, confirmado = _feriados()
    return {"mes_entrega": f"{m:02d}/{a}",
            "grupos": nesimo_dia_util(a, m, pz["grupos_dia_util"]).strftime("%d/%m/%Y"),
            "coordenacao": nesimo_dia_util(a, m, pz["coordenacao_dia_util"]).strftime("%d/%m/%Y"),
            "base": "estimado: seg–sex menos feriados NACIONAIS de config/feriados.json"
                    + ("" if confirmado else " (lista não confirmada; feriados estaduais/municipais não incluídos)")}

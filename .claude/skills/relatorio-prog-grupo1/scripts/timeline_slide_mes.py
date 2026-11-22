# -*- coding: utf-8 -*-
"""
timeline_slide_mes.py — CRIA o slide do mês no PPTX-base da Timeline a partir dos marcos do Controle Timeline
(automatiza o procedimento manual da DEC-040; DEC-060).

1. backup do PPTX-base; 2. se o slide do mês não existe: duplica o slide do mês anterior (PowerPoint COM);
3. troca o mês da faixa "[ Mês/ano ]"; 4. usa os marcadores REAIS do slide como modelos (um "acima", um "abaixo":
   'Group 9' com a descrição acima/abaixo do eixo) e recria N marcadores, alternando acima/abaixo (o 1º acima),
   posicionados pela BOLINHA em X distribuídos entre as margens; 5. escreve data 'DD/MM' e título.
Não recria a timeline do zero (DEC-013): só duplica os elementos-modelo. O 17 depois renderiza tudo.

Uso: python timeline_slide_mes.py <pptx_base> <controle_timeline.xlsx> <ano> <mes>   [--margem-esq 70 --margem-dir 772]
"""
import os, shutil, sys, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
EIXO_Y = 305.0


def slide_index(ano, mes):
    return (ano - 2025) * 12 + mes - 6


def _walk(shapes):
    for i in range(1, shapes.Count + 1):
        s = shapes.Item(i)
        yield s
        if s.Type == 6:  # msoGroup
            for j in range(1, s.GroupItems.Count + 1):
                yield s.GroupItems.Item(j)


def _textos(group):
    out = []
    for j in range(1, group.GroupItems.Count + 1):
        g = group.GroupItems.Item(j)
        if g.HasTextFrame:
            out.append((g, g.TextFrame.TextRange.Text))
    return out


def _bolinha(group):
    """Item freeform pequeno (a bolinha) do marcador: o de menor área entre os freeforms."""
    cands = [group.GroupItems.Item(j) for j in range(1, group.GroupItems.Count + 1)
             if group.GroupItems.Item(j).Type == 5]
    return min(cands, key=lambda s: s.Width * s.Height) if cands else None


def _tem_data(group):
    import re
    return any(re.fullmatch(r"\s*\d\d/\d\d\s*", t) for _g, t in _textos(group))


def _tipo(group):
    """'acima' se a caixa de descrição (maior texto) está acima do eixo; 'abaixo' caso contrário.
    Só marcadores (grupo com caixa de data DD/MM) — a faixa do topo NÃO é marcador."""
    import re
    if not _tem_data(group):
        return None
    desc = [(g, t) for g, t in _textos(group) if len(t.strip()) > 6 and not re.fullmatch(r"\d\d/\d\d", t.strip())]
    if not desc:
        return None
    g = max(desc, key=lambda x: len(x[1]))[0]
    return "acima" if g.Top + g.Height / 2 < EIXO_Y else "abaixo"


def preencher(group, data, titulo, largura=None, x=None):
    import re
    caixas = _textos(group)
    datas = [g for g, t in caixas if re.fullmatch(r"\s*\d\d/\d\d\s*", t)]
    desc = [(g, t) for g, t in caixas if len(t.strip()) > 6 and g not in datas]
    for g in datas:
        g.TextFrame.TextRange.Text = data
    if desc:
        g = max(desc, key=lambda x: len(x[1]))[0]
        import re as _r
        g.TextFrame.TextRange.Text = _r.sub(r"RD-(\d)", lambda m: "RD\u2011" + m.group(1), titulo)  # hífen inseparável (U+2011): não quebra RD-101
        if largura and x is not None:  # descrição centrada na bolinha, sem invadir o vizinho do mesmo lado
            g.TextFrame.WordWrap = True
            w = largura
            g.Width = w
            g.Left = max(5.0, min(837.0 - w, x - w / 2))


def criar(pptx_base, marcos, ano, mes, x_ini=76.0, x_fim=766.0, backup_dir=None):
    """marcos = [(data 'DD/MM', título)] em ordem. Devolve dict com o que fez."""
    import win32com.client
    if backup_dir:
        os.makedirs(backup_dir, exist_ok=True)
        shutil.copy2(pptx_base, os.path.join(backup_dir, "BACKUP_" + os.path.basename(pptx_base)))
    app = win32com.client.DispatchEx("PowerPoint.Application")
    res = {"slide": slide_index(ano, mes), "marcos": len(marcos)}
    try:
        pres = app.Presentations.Open(os.path.abspath(pptx_base), False, False, False)
        idx = slide_index(ano, mes)
        if pres.Slides.Count >= idx:
            res["acao"] = "slide já existe — não alterado"
            pres.Close()
            return res
        if pres.Slides.Count != idx - 1:
            raise RuntimeError(f"PPTX com {pres.Slides.Count} slides; esperado {idx-1} (mês anterior)")
        novo = pres.Slides(idx - 1).Duplicate().Item(1)
        # faixa do mês
        ant = f"{MESES[(mes - 2) % 12]}/{ano if mes > 1 else ano - 1}"
        for s in _walk(novo.Shapes):
            if s.HasTextFrame and ant in s.TextFrame.TextRange.Text:
                s.TextFrame.TextRange.Replace(ant, f"{MESES[mes-1]}/{ano}")
        # modelos
        grupos = [novo.Shapes.Item(i) for i in range(1, novo.Shapes.Count + 1) if novo.Shapes.Item(i).Type == 6
                  and _bolinha(novo.Shapes.Item(i)) is not None and _tipo(novo.Shapes.Item(i))]
        acima = next((g for g in grupos if _tipo(g) == "acima"), None)
        abaixo = next((g for g in grupos if _tipo(g) == "abaixo"), None)
        if not (acima and abaixo):
            raise RuntimeError("slide anterior sem os dois modelos de marcador (acima/abaixo)")
        n = len(marcos)
        xs = [x_ini + (x_fim - x_ini) * i / (n - 1) for i in range(n)] if n > 1 else [(x_ini + x_fim) / 2]
        criados = []
        for i, ((data, tit), x) in enumerate(zip(marcos, xs)):
            modelo = acima if i % 2 == 0 else abaixo
            g = modelo.Duplicate().Item(1)
            g.Top = modelo.Top
            b = _bolinha(g)
            g.Left = g.Left + (x - (b.Left + b.Width / 2))
            esp = (xs[1] - xs[0]) if n > 1 else 300.0
            preencher(g, data, tit, largura=min(200.0, 1.78 * esp), x=x)
            criados.append(g.Name)
        for g in grupos:
            g.Delete()
        pres.Save()
        pres.Close()
        res.update(acao="slide criado", xs=[round(x) for x in xs])
    finally:
        app.Quit()
    return res


def marcos_do_excel(xlsx, ano, mes):
    import importlib.util
    spec = importlib.util.spec_from_file_location("t17", os.path.join(os.path.dirname(__file__), "17_timeline_do_excel.py"))
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return [(d, t) for _dt, d, t in m.ler_marcos(xlsx, ano, mes)]


if __name__ == "__main__":
    import json
    sys.stdout.reconfigure(encoding="utf-8")
    pptx, xlsx, a, m = sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4])
    bk = sys.argv[5] if len(sys.argv) > 5 else None
    mm = marcos_do_excel(xlsx, a, m)
    print(json.dumps(criar(pptx, mm, a, m, backup_dir=bk), ensure_ascii=False))

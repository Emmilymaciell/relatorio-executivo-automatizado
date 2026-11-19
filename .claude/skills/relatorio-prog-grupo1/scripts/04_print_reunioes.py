# -*- coding: utf-8 -*-
"""
04_print_reunioes.py — Gera as imagens da tabela "Programa | Controle de Reuniões" a partir do
anexo 'P1.2 - Controle de Reuniões do mês de {mês}.xlsx' (SharePoint, LEITURA).
- Governança (P1.1)     -> 03_IMAGENS/REUNIOES_P11/gov_p11_pNN.png
- Institucional (P1.2)  -> 03_IMAGENS/REUNIOES_P12/inst_p12_pNN.png

PRINCÍPIO (DEC-027): a imagem inserida no relatório é CONSEQUÊNCIA de uma CÓPIA do Excel
corretamente formatada — nunca de correção posterior no PNG. Ordem obrigatória:
    dados -> cópia do Excel -> formatação -> dimensões -> quebras -> área/escala de impressão
    -> renderização -> validação visual/programática -> inserção no relatório

ESCALA 1:1 (chave de tudo): as margens de impressão da cópia são calculadas para que a
LARGURA NATURAL da tabela (B+C+D) seja igual à LARGURA ÚTIL DO CORPO do relatório. Com
Zoom=100 (sem FitToPages) e a imagem inserida no Word com essa mesma largura, o fator de
escala total é 1,0 — logo **o tamanho em pt no Excel é o tamanho visual no PDF final**.
Consequências: a fonte do conteúdo é definida pelo alvo visual (≈ fonte do texto corrido) e
as bordas esquerda/direita da tabela coincidem com o início/fim do texto justificado.

Só escreve arquivos locais de trabalho (nunca altera o anexo original).
Parametrizado por prog_config (nenhum caminho/mês fixo).

Variáveis de ambiente (opcionais):
    PROG_REUNIOES_FONTE_PT   — alvo visual da fonte do CONTEÚDO (default 9.5)
    PROG_BASE_DOCX           — .docx-base (CG) usado para medir a geometria/espaço da página
    PROG_REUNIOES_SEM_PREVIA — 1 para não gerar o PDF de prévia/validação (mais rápido)
    PROG_REUNIOES_ORIENTACAO — auto (default) | retrato | paisagem
"""
import os, sys, re, json, glob, math, shutil, atexit
from collections import Counter
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

# ---------------------------------------------------------------- parâmetros visuais
DPI              = 300      # resolução dos PNGs (escala 1:1 -> 300 dpi reais no PDF)
A4_W_PT, A4_H_PT = 595.28, 841.89
PT_IN            = 72.0

FONTE_CONTEUDO_PT   = float(os.environ.get("PROG_REUNIOES_FONTE_PT", "9.5"))
FONTE_CABECALHO_PT  = FONTE_CONTEUDO_PT + 1.5   # Data/Participantes/Assuntos
FONTE_TITULO_PT     = FONTE_CONTEUDO_PT + 3.5   # "Programa | Controle de Reuniões"
# Proporções Data/Participantes/Assuntos (somam 1,0). Assuntos leva toda a largura extra: é ele
# que manda na altura da tabela. Data fica sempre larga o bastante p/ dd/mm/aaaa em uma linha.
PROP_RETRATO        = (0.115, 0.205, 0.680)
PROP_PAISAGEM       = (0.076, 0.150, 0.774)
LARG_PAISAGEM_IN    = 10.4    # largura útil das páginas paisagem deste relatório (= Controle/Timeline)
ALT_PAISAGEM_IN     = 6.5     # altura útil das páginas paisagem (DEC-015/024)
ORIENTACAO          = os.environ.get("PROG_REUNIOES_ORIENTACAO", "auto").lower()

ALT_FAIXA_PT        = 24.0   # linha 1 (faixa do título + logo) — proporção do modelo
                             # manual; cada pt aqui se repete em TODAS as partes
ALT_ESPACADOR_PT    = 3.0    # linha 2 (respiro entre faixa e cabeçalho)
FOLGA_LINHA_PT      = 1.0    # folga sobre o AutoFit (Excel x impressão)
ESTREITA_AUTOFIT    = 0.985  # AutoFit medido com a coluna 1,5% mais estreita (+ medição por parágrafo — DEC-056): garante
                             # que a última linha não seja cortada na impressão
MARGEM_LR_PT        = 25.0   # margens laterais da cópia (só espaço branco: é recortado;
                             # o que importa é sobrar largura imprimível > tabela)
MARGEM_TB_PT        = 14.4   # margens superior/inferior da cópia (0,2")
RESERVA_RODAPE_PT   = 14.4   # respiro sobre o risco do rodapé do relatório (DEC-024)
RESERVA_PRIMEIRA_PT = 7.2    # respiro da 1ª parte: ela para no fim da área de texto
                             # (o rodapé fica ABAIXO da margem inferior), então 0,1" basta
MIN_PRIMEIRA_IN     = 2.5    # espaço mínimo útil p/ começar a tabela na página do texto
LIM_ROWHEIGHT_PT    = 405.0  # limite do Excel para altura de linha (409)

xlCenter, xlLeft, xlTop = -4108, -4131, -4160
BULLETS = ("-", "–", "—", "•", "*", "▪", "·")


def chars(cel, ini, qtd):
    """Range.Characters(Start, Length) — com early binding (gencache) a propriedade
    parametrizada é exposta como GetCharacters."""
    try:
        return cel.GetCharacters(ini, qtd)
    except AttributeError:
        return cel.Characters(ini, qtd)


# ---------------------------------------------------------------- geometria do relatório
def base_docx():
    env = os.environ.get("PROG_BASE_DOCX")
    if env and os.path.exists(env):
        return env
    return cfg.base_am_docx()


def geometria_corpo():
    """Geometria da seção RETRATO do relatório-base: largura útil do corpo (= largura da
    tabela) e altura útil da página. É o contrato horizontal exigido: borda esquerda/direita
    da tabela alinhadas ao início/fim do texto justificado."""
    E = 914400.0
    g = dict(larg_in=6.891, pag_h_pt=A4_H_PT, top_pt=56.7, bottom_pt=42.5, fonte="padrão(A4/RM13)")
    p = base_docx()
    try:
        from docx import Document
        s = Document(p).sections[0]
        g = dict(larg_in=(s.page_width - s.left_margin - s.right_margin) / E,
                 pag_h_pt=(s.page_height / E) * PT_IN,
                 top_pt=(s.top_margin / E) * PT_IN,
                 bottom_pt=(s.bottom_margin / E) * PT_IN,
                 fonte=os.path.basename(p))
    except Exception as e:
        cfg.log(f"  [geometria] usando padrão ({e})")
    g["alt_pagina_pt"] = g["pag_h_pt"] - g["top_pt"] - g["bottom_pt"] - RESERVA_RODAPE_PT
    return g


# ---------------------------------------------------------------- esqueleto Word (medida/prévia)
def _doc_esqueleto(WK, nome):
    """DOCX vazio que herda do relatório-base os estilos, o tema (fontes) e a geometria da
    seção retrato. Serve para medir o espaço restante e para gerar a prévia — com o mesmo
    motor de layout do Word usado no relatório final."""
    from docx import Document
    from docx.oxml.ns import qn
    src = base_docx()
    tmp = os.path.join(WK, "02_TRABALHO", nome)
    if src and os.path.exists(src):
        shutil.copy2(src, tmp)
        d = Document(tmp)
        body = d.element.body
        for el in list(body):
            if el.tag != qn('w:sectPr'):
                body.remove(el)
        sect = body.find(qn('w:sectPr'))
        if sect is not None:
            for t in sect.findall(qn('w:titlePg')):
                sect.remove(t)   # 1ª página = página normal (como no meio do relatório)
    else:
        d = Document()
    return d, tmp


def _titulo_par(d, texto):
    """Mesma formatação de title() do 09_montar_word.py."""
    from docx.shared import Pt
    p = d.add_paragraph()
    try: p.style = d.styles['List Bullet']
    except Exception: pass
    r = p.add_run(texto); r.bold = False; r.font.size = Pt(11)
    p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(4)
    return p


def _corpo_par(d, texto):
    """Mesma formatação de body() do 09_montar_word.py."""
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
    p = d.add_paragraph(texto); p.alignment = AL.JUSTIFY
    for r in p.runs: r.font.size = Pt(11)
    p.paragraph_format.space_after = Pt(5)
    return p


def _img_par(d, path, largura_in):
    """Mesma formatação de _pic() do 09_montar_word.py."""
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
    p = d.add_paragraph(); p.alignment = AL.CENTER
    p.add_run().add_picture(path, width=Inches(largura_in))
    p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(6)
    return p


def _word():
    """Instância DEDICADA do Word (DispatchEx): não se liga a um Word já aberto — assim não
    interfere nos documentos do usuário nem herda uma instância travada (RPC indisponível)."""
    try:
        wd = win32.DispatchEx("Word.Application")
    except Exception:
        wd = win32.gencache.EnsureDispatch("Word.Application")
    wd.Visible = False; wd.DisplayAlerts = 0
    return wd


def espaco_real_no_montado(WK, geom, titulo):
    """Espaço (pt) disponível abaixo do texto do P1.1 medido NO RELATÓRIO REAL montado por uma
    passagem anterior do 09 — é a única medida sem viés (mesma seção, mesmo cabeçalho, mesma
    quebra de parágrafos). O texto acima da tabela não muda entre passagens, por isso a medida
    converge: 04 -> ... -> 09 e, na 2ª passagem, 04 já pagina com o espaço verdadeiro.
    Devolve None quando ainda não existe DOCX montado (1ª passagem)."""
    montado = os.path.join(WK, "02_TRABALHO", f"RM{cfg.numero()}_montado.docx")
    if not os.path.exists(montado):
        return None
    alvo = titulo[:38].lower()
    y = pag = None
    wd = doc = None
    try:
        wd = _word()
        doc = wd.Documents.Open(os.path.abspath(montado), False, True)   # ReadOnly
        doc.Repaginate()
        achou = False
        for i in range(1, int(doc.Paragraphs.Count) + 1):
            r = doc.Paragraphs(i).Range
            if not achou:
                if alvo in (r.Text or "").strip().lower(): achou = True
                continue
            if int(r.InlineShapes.Count) > 0:
                y = float(r.Information(6)); pag = int(r.Information(3)); break
    except Exception as e:
        cfg.log(f"  [espaço] falha ao medir no montado ({type(e).__name__}: {e})")
    finally:
        # fechar/encerrar NUNCA pode descartar uma medida já obtida
        try:
            if doc is not None: doc.Close(0)
        except Exception: pass
        try:
            if wd is not None: wd.Quit()
        except Exception: pass
    if y is None:
        return None
    limite = geom["pag_h_pt"] - geom["bottom_pt"] - RESERVA_PRIMEIRA_PT
    disp = max(0.0, limite - y)
    cfg.log(f"  [espaço] medido no RELATÓRIO montado: tabela do P1.1 comeca na pag. {pag} "
            f"em y={y:.0f}pt -> disponivel {disp/PT_IN:.2f}in (medida sem vies)")
    return disp


def medir_espaco_apos_texto(WK, geom, titulo, paragrafos):
    """Altura (pt) realmente disponível na página onde o texto do P1.1 termina — medida no
    Word, com os estilos do relatório. O P1.1 'Reuniões de Governança' começa sempre no topo
    de uma página nova (o 09 fecha a seção paisagem da Timeline antes), por isso a medida
    feita num documento só com esse texto é fiel ao relatório final."""
    try:
        d, tmp = _doc_esqueleto(WK, "medida_reunioes.docx")
        _titulo_par(d, titulo)
        for t in paragrafos: _corpo_par(d, t)
        from docx.shared import Pt
        sent = d.add_paragraph(); sent.paragraph_format.space_before = Pt(3)  # onde a imagem entra
        d.save(tmp)
        wd = _word()
        try:
            doc = wd.Documents.Open(os.path.abspath(tmp)); doc.Repaginate()
            rng = doc.Paragraphs(doc.Paragraphs.Count).Range
            y   = float(rng.Information(6))   # wdVerticalPositionRelativeToPage (pt)
            pag = int(rng.Information(3))     # wdActiveEndPageNumber
            doc.Close(0)
        finally:
            wd.Quit()
        limite = geom["pag_h_pt"] - geom["bottom_pt"] - RESERVA_RODAPE_PT
        disp   = max(0.0, limite - y - 3.0)   # -3pt = space_before do parágrafo da imagem
        cfg.log(f"  [espaço] texto do P1.1 termina na pág. {pag} em y={y:.0f}pt "
                f"-> disponível abaixo: {disp/PT_IN:.2f}in")
        return disp
    except Exception as e:
        cfg.log(f"  [espaço] não foi possível medir ({e}); usando página cheia")
        return geom["alt_pagina_pt"]


# ---------------------------------------------------------------- formatação da cópia Excel
def ultima_linha(ws):
    last = 0
    for col in (2, 3, 4):
        last = max(last, ws.Cells(ws.Rows.Count, col).End(-4162).Row)
    return last


def ajustar_larguras(ws, alvo_pt, prop):
    """Larguras de B/C/D em PROP_COLUNAS do alvo. O Excel quantiza a largura em pixels, por
    isso converge por iteração (ColumnWidth em caracteres <-> Width em pontos)."""
    alvos = [alvo_pt * x for x in prop]
    for _ in range(10):
        cur = [float(ws.Columns(c).Width) for c in (2, 3, 4)]
        if max(abs(a - c) for a, c in zip(alvos, cur)) < 0.4:
            break
        for i, c in enumerate((2, 3, 4)):
            if cur[i] <= 0: continue
            cw = float(ws.Columns(c).ColumnWidth)
            ws.Columns(c).ColumnWidth = max(1.0, cw * alvos[i] / cur[i])
    real = sum(float(ws.Columns(c).Width) for c in (2, 3, 4))
    cfg.log("  [larguras] " + " + ".join(f"{float(ws.Columns(c).Width):.1f}" for c in (2, 3, 4))
            + f" = {real:.1f}pt ({real/PT_IN:.3f}in; alvo {alvo_pt/PT_IN:.3f}in)")
    return real


def negrito_participantes(ws, r):
    """Órgão/sigla que antecede os participantes SEMPRE em negrito, incluindo os ':'.
    Reaplica do zero (não confia na formatação recebida)."""
    cel = ws.Cells(r, 3); v = cel.Value
    if not isinstance(v, str) or not v.strip(): return
    cel.Font.Bold = False
    pos = 1
    for linha in v.split("\n"):
        i = linha.find(":")
        if i >= 0:
            chars(cel, pos, i + 1).Font.Bold = True
        pos += len(linha) + 1


def negrito_assuntos(ws, r):
    """Título de cada bloco/reunião em negrito; itens (linhas de bullet) em regular."""
    cel = ws.Cells(r, 4); v = cel.Value
    if not isinstance(v, str) or not v.strip(): return
    cel.Font.Bold = False
    pos = 1
    for k, linha in enumerate(v.split("\n")):
        t = linha.strip()
        if t and not t.startswith(BULLETS):
            chars(cel, pos, len(linha)).Font.Bold = True
        pos += len(linha) + 1


def normalizar_data(ws, r):
    """Data em dd/mm/aaaa (DEC-001) quando a célula é data de verdade."""
    cel = ws.Cells(r, 2)
    try:
        if hasattr(cel.Value, "year"):
            cel.NumberFormat = "dd/mm/aaaa"
    except Exception:
        pass


def validar_faixa_e_logo(ws, x_esq_pt, larg_pt, logo_png):
    """Título, faixa e logo — exigidos e validados ANTES do print (req. 1)."""
    avisos = []
    if not str(ws.Cells(1, 2).Value or "").strip():
        ws.Cells(1, 2).Value = "   Programa | Controle de Reuniões"
        avisos.append("título da faixa estava vazio (restaurado)")
    ws.Range("B1:D1").Font.Size = FONTE_TITULO_PT
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(1, 2).VerticalAlignment = xlCenter
    # faixa contínua: replica a cor de B1 nas colunas do print
    try:
        cor = ws.Cells(1, 2).Interior.Color
        if ws.Cells(1, 2).Interior.Pattern == -4142:   # xlPatternNone
            avisos.append("faixa do título sem preenchimento")
        else:
            ws.Range("B1:D1").Interior.Color = cor
    except Exception as e:
        avisos.append(f"cor da faixa não verificada ({e})")
    # logo à direita, dentro da faixa
    top, alt = float(ws.Rows(1).Top), float(ws.Rows(1).Height)
    x_dir = x_esq_pt + larg_pt
    shapes = [ws.Shapes(i + 1) for i in range(ws.Shapes.Count)]
    if not shapes:
        if logo_png and os.path.exists(logo_png):
            shapes = [ws.Shapes.AddPicture(logo_png, False, True, x_dir - 60, top + 3, -1, -1)]
            avisos.append("logo ausente na planilha: inserida de assets/logo_prog.png")
        else:
            avisos.append("ATENÇÃO: logo Programa ausente e sem asset de reposição")
    for sh in shapes:
        try:
            asp = (float(sh.Width) / float(sh.Height)) if float(sh.Height) else 1.865
            h = max(8.0, alt - 8.0)     # cabe INTEIRA dentro da faixa (nada invade o cabeçalho)
            sh.LockAspectRatio = 0
            sh.Height = h; sh.Width = h * asp
            sh.Top = top + (alt - h) / 2.0
            sh.Left = x_dir - float(sh.Width) - 3.0
            if float(sh.Left) < x_esq_pt:
                avisos.append("logo mais larga que a tabela (reduzida)")
                sh.Width = larg_pt * 0.25; sh.Height = sh.Width / asp
                sh.Left = x_dir - float(sh.Width) - 3.0
        except Exception as e:
            avisos.append(f"logo não reposicionada ({e})")
    for a in avisos:
        cfg.log(f"  [faixa/logo] {a}")
    return avisos


def medir_por_paragrafo(wb, ws, last, colunas):
    """Altura necessária por linha = Σ alturas do AutoFit de CADA parágrafo da célula, medido numa aba temporária
    com a mesma largura (já estreitada), fonte e tamanho; o 1º parágrafo de Assuntos é medido em negrito (título).
    Devolve {linha: maior soma entre as colunas}. DEC-056."""
    tmp = wb.Worksheets.Add()
    alturas = {}
    try:
        for c in colunas:
            tmp.Columns(1).ColumnWidth = float(ws.Columns(c).ColumnWidth)
            cel = tmp.Cells(1, 1)
            cel.WrapText = True
            for r in range(4, last + 1):
                src = ws.Cells(r, c)
                cel.Font.Name = src.Font.Name
                cel.Font.Size = src.Font.Size
                total = 0.0
                for k, par in enumerate(str(src.Text or "").split("\n")):
                    cel.Value = par if par.strip() else " "
                    cel.Font.Bold = (c == 4 and k == 0) or (c == 3 and ":" in par)
                    tmp.Rows(1).AutoFit()
                    total += float(tmp.Rows(1).RowHeight)
                alturas[r] = max(alturas.get(r, 0.0), total)
    finally:
        xl = wb.Application
        da = xl.DisplayAlerts; xl.DisplayAlerts = False
        tmp.Delete(); xl.DisplayAlerts = da
        ws.Activate()
    return alturas


def formatar(wb, ws, tag, WK, alvo_pt, logo_png, prop, orient):
    """Aplica NA CÓPIA DO EXCEL tudo o que precisa aparecer no print: fonte, negrito parcial,
    alinhamentos, quebra de texto, larguras, alturas mínimas suficientes, faixa e logo."""
    last = ultima_linha(ws)
    if last < 4:
        return None
    ws.Columns("A").ColumnWidth = 2
    try: ws.Columns("E").Hidden = True      # Anexos (não usado nestas reuniões)
    except Exception: pass

    # 1ª aproximação das larguras -> calibração -> larguras definitivas em pontos IMPRESSOS
    ajustar_larguras(ws, alvo_pt, prop)
    fx, fy = calibrar_impressao(wb, ws, tag, WK,
                                sum(float(ws.Columns(c).Width) for c in (2, 3, 4)), orient)
    larg_rep = ajustar_larguras(ws, alvo_pt / fx, prop)
    larg_real = larg_rep * fx
    cfg.log(f"  [larguras] impressa = {larg_real:.1f}pt ({larg_real/PT_IN:.3f}in) "
            f"= largura util do corpo")

    corpo = ws.Range(ws.Cells(4, 2), ws.Cells(last, 4))
    corpo.Font.Size = FONTE_CONTEUDO_PT
    corpo.WrapText = True
    # alinhamentos obrigatórios (req. 5)
    data = ws.Range(ws.Cells(4, 2), ws.Cells(last, 2))
    data.HorizontalAlignment = xlCenter; data.VerticalAlignment = xlCenter
    part = ws.Range(ws.Cells(4, 3), ws.Cells(last, 3))
    part.HorizontalAlignment = xlCenter; part.VerticalAlignment = xlCenter
    txt = ws.Range(ws.Cells(4, 4), ws.Cells(last, 4))
    txt.HorizontalAlignment = xlLeft;    txt.VerticalAlignment = xlTop
    ws.Range("B3:D3").Font.Size = FONTE_CABECALHO_PT
    ws.Range("B3:D3").WrapText = True

    for r in range(4, last + 1):
        normalizar_data(ws, r)
        negrito_participantes(ws, r)
        negrito_assuntos(ws, r)

    # alturas: faixa/espaçador fixos; cabeçalho e conteúdo pelo MENOR AutoFit + folga (req. 6).
    # A faixa é dimensionada ANTES de posicionar a logo (a logo é ajustada à altura da faixa).
    ws.Rows(1).RowHeight = ALT_FAIXA_PT
    ws.Rows(2).RowHeight = ALT_ESPACADOR_PT

    # posições/tamanhos de FORMAS usam os pontos REPORTADOS pelo Excel (não os impressos):
    # com a largura impressa a logo passaria da coluna D e a impressão a cortaria.
    validar_faixa_e_logo(ws, float(ws.Cells(1, 2).Left), larg_rep, logo_png)

    ws.Range("B3:D3").Rows.AutoFit()
    ws.Rows(3).RowHeight = max(float(ws.Rows(3).RowHeight) + 2.0, FONTE_CABECALHO_PT * 1.45)
    estourou = []
    largs = [float(ws.Columns(c).ColumnWidth) for c in (3, 4)]
    for i, c in enumerate((3, 4)):
        ws.Columns(c).ColumnWidth = largs[i] * ESTREITA_AUTOFIT
    for r in range(4, last + 1):
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Rows.AutoFit()
        ws.Rows(r).RowHeight = float(ws.Rows(r).RowHeight)   # congela o valor medido
    # DEC-056: o AutoFit de célula com VÁRIAS linhas (\n) às vezes ignora quebras automáticas (RM15: linha com 9
    # linhas visuais recebeu altura de ~6,7). Mede cada parágrafo isoladamente (AutoFit de 1 parágrafo é confiável),
    # com a mesma fonte/largura estreitada, e usa a SOMA quando for maior que o AutoFit da célula inteira.
    soma = medir_por_paragrafo(wb, ws, last, (3, 4))
    for r in range(4, last + 1):
        if soma.get(r, 0.0) > float(ws.Rows(r).RowHeight):
            ws.Rows(r).RowHeight = min(LIM_ROWHEIGHT_PT, soma[r])
    for i, c in enumerate((3, 4)):
        ws.Columns(c).ColumnWidth = largs[i]
    for r in range(4, last + 1):
        h = float(ws.Rows(r).RowHeight) + FOLGA_LINHA_PT
        h = max(h, FONTE_CONTEUDO_PT * 1.35 + 3.0)
        if h > LIM_ROWHEIGHT_PT:
            estourou.append(r); h = LIM_ROWHEIGHT_PT
        ws.Rows(r).RowHeight = h
    if estourou:
        cfg.log(f"  AVISO: linhas {estourou} excedem o limite de altura do Excel "
                f"({LIM_ROWHEIGHT_PT:.0f}pt) — serão divididas pela impressão")
    # alturas convertidas para PONTOS IMPRESSOS (é o que decide a paginação na página do Word)
    h_tit = sum(float(ws.Rows(r).Height) for r in (1, 2, 3)) * fy
    rep = {r: float(ws.Rows(r).Height) for r in range(4, last + 1)}
    return dict(last=last, larg_pt=larg_real, h_titulos=h_tit, fx=fx, fy=fy, rep=rep,
                alturas={r: h * fy for r, h in rep.items()})


# ---------------------------------------------------------------- paginação + exportação
def paginar(alturas, ordem, max_primeira, max_demais):
    """Quebras SEMPRE entre reuniões completas; nunca corta uma reunião ao meio. A divisão é
    determinada pelo conteúdo real e pelo espaço disponível (req. 7/8)."""
    partes, atual, acc = [], [], 0.0
    lim = max_primeira
    for r in ordem:
        h = alturas[r]
        if atual and acc + h > lim:
            partes.append(atual); atual, acc, lim = [], 0.0, max_demais
        atual.append(r); acc += h
    if atual: partes.append(atual)
    return partes


def exportar_pdf_parte(wb, ws, linhas, tag, k, WK, larg_pt, orient=1):
    """Exporta UMA parte como PDF em escala 100%: cópia da aba contendo apenas as linhas da
    parte, precedidas da faixa/título/logo e do cabeçalho reais — por isso a identidade visual e
    o cabeçalho Data|Participantes|Assuntos se repetem em todas as partes, com as MESMAS
    larguras de coluna."""
    last = ultima_linha(ws)
    pdf = os.path.join(WK, "04_PDFS_TEMP", f"reunioes_{tag}_p{k:02d}.pdf")
    ws.Copy(None, wb.Worksheets(wb.Worksheets.Count))
    w2 = wb.Worksheets(wb.Worksheets.Count)
    try:
        r0, r1 = linhas[0], linhas[-1]
        if r1 < last: w2.Rows(f"{r1+1}:{last}").Delete()
        if r0 > 4:    w2.Rows(f"4:{r0-1}").Delete()
        fim_l = 3 + len(linhas)
        # DEC-056: quebras de página MANUAIS herdadas da mestre (RM15: aba GOV com quebras nas linhas 7 e 10)
        # fatiavam a parte em 2 páginas — e parte fatiada desliga a checagem de corte. A paginação é nossa.
        w2.ResetAllPageBreaks()
        ps = w2.PageSetup
        ps.Orientation = orient; ps.PaperSize = 9
        ps.Zoom = 100                      # 1:1 — nada de FitToPages (não comprime a tabela)
        ps.PrintArea = f"$B$1:$D${fim_l}"
        ps.PrintTitleRows = ""
        ps.TopMargin = MARGEM_TB_PT; ps.BottomMargin = MARGEM_TB_PT
        ps.HeaderMargin = 0; ps.FooterMargin = 0
        ps.CenterHorizontally = False; ps.CenterVertically = False
        # a tabela tem de caber na largura imprimível SEM redução de escala, senão o Excel a
        # fatia em duas páginas (colunas B/C numa, D noutra)
        for m in (MARGEM_LR_PT, 12.0, 6.0):
            ps.LeftMargin = m; ps.RightMargin = m
            if int(w2.VPageBreaks.Count) == 0: break
        if int(w2.VPageBreaks.Count) > 0:
            cfg.log(f"  AVISO [{tag} p{k}] a tabela ({larg_pt:.0f}pt) nao cabe na largura "
                    f"imprimivel do papel: seria fatiada horizontalmente")
        if os.path.exists(pdf): os.remove(pdf)
        w2.ExportAsFixedFormat(0, pdf)
    finally:
        w2.Delete()
    return pdf


def _tokens(txt):
    return [t for t in re.findall(r"[0-9A-Za-zÀ-ÿ]{4,}", (txt or "").lower())]


def _spans(pg):
    out = []
    for b in pg.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for sp in l["spans"]:
                out.append(sp["bbox"])
    return out


def _bandas(pg, larg_pt):
    """Linhas horizontais de borda da tabela no PDF, de cima para baixo."""
    ys = []
    for dr in pg.get_drawings():
        for it in dr["items"]:
            try:
                if it[0] == "l":
                    a, b = it[1], it[2]
                    if abs(a.y - b.y) < 0.4 and abs(a.x - b.x) > 0.8 * larg_pt:
                        ys.append(a.y)
                elif it[0] == "re":
                    r = it[1]
                    if r.height < 2.0 and r.width > 0.8 * larg_pt:
                        ys.append(r.y0)
            except Exception:
                pass
    ys.sort(); out = []
    for y in ys:
        if not out or y - out[-1] > 1.0: out.append(y)
    return out


def _bordas_verticais(pg, alt_min=40.0):
    """Linhas verticais de borda da tabela no PDF, da esquerda para a direita."""
    xs = []
    for dr in pg.get_drawings():
        for it in dr["items"]:
            try:
                if it[0] == "re" and it[1].width < 2.0 and it[1].height > alt_min:
                    xs.append(it[1].x0)
                elif it[0] == "l" and abs(it[1].x - it[2].x) < 0.4 and abs(it[1].y - it[2].y) > alt_min:
                    xs.append(it[1].x)
            except Exception:
                pass
    xs.sort(); out = []
    for x in xs:
        if not out or x - out[-1] > 1.0: out.append(x)
    return out


def calibrar_impressao(wb, ws, tag, WK, larg_reportada, orient=1):
    """O Excel REPORTA Width/Height em pontos que NÃO são os pontos IMPRESSOS (efeito de DPI/driver:
    nesta máquina largura ×1,0225 e altura ×0,9375). Como todo o dimensionamento aqui é feito em
    pontos impressos — a tabela tem de sair com a largura EXATA do corpo e a paginação depende da
    altura real —, os dois fatores são medidos num PDF de calibração (1 reunião) e aplicados.
    Devolve (fx, fy); (1.0, 1.0) se a medição falhar."""
    import fitz
    fx = fy = 1.0
    try:
        pdf = exportar_pdf_parte(wb, ws, [4], f"{tag}_calib", 0, WK, larg_reportada, orient)
        doc = fitz.open(pdf)
        try:
            pg = doc[0]
            vs = _bordas_verticais(pg)
            hs = _bandas(pg, larg_reportada * 0.75)
            if len(vs) >= 2 and larg_reportada > 0:
                fx = (vs[-1] - vs[0]) / larg_reportada
            if len(hs) >= 3:
                rep = float(ws.Rows(3).Height) + float(ws.Rows(4).Height)
                if rep > 0: fy = (hs[-1] - hs[0]) / rep
        finally:
            doc.close()
        if not (0.8 <= fx <= 1.25): fx = 1.0
        if not (0.8 <= fy <= 1.25): fy = 1.0
    except Exception as e:
        cfg.log(f"  [calibração] falhou ({e}); assumindo 1:1")
    cfg.log(f"  [calibração] impresso/reportado: largura ×{fx:.4f} | altura ×{fy:.4f}")
    return fx, fy


def cortes_no_pdf(pdf, linhas, larg_pt):
    """Texto CORTADO medido no PDF RENDERIZADO (não na planilha): devolve, por linha, quantos
    pontos faltam para a última linha de texto caber dentro da própria borda (req. 6.2-6.4)."""
    import fitz
    faltas = {}
    doc = fitz.open(pdf)
    try:
        if doc.page_count != 1:
            return faltas                    # parte fatiada: mapeamento banda->linha não vale
        pg = doc[0]
        bandas = _bandas(pg, larg_pt)
        if len(bandas) < 3: return faltas
        sp = _spans(pg)
        pares = list(zip(bandas, bandas[1:]))[1:]   # descarta a banda do cabeçalho
        for i, (a, b) in enumerate(pares):
            if i >= len(linhas): break
            dentro = [x for x in sp if a - 0.5 <= x[1] < b - 0.5]
            if not dentro: continue
            fundo = max(x[3] for x in dentro)
            if fundo > b - 0.3:
                faltas[linhas[i]] = max(faltas.get(linhas[i], 0.0), fundo - b + 1.0)
    finally:
        doc.close()
    return faltas


def texto_ausente(pdf, ws, linhas):
    """Conferência complementar: nenhum termo das células pode desaparecer do PDF."""
    import fitz
    esperado = Counter()
    for r in linhas:
        for c in (2, 3, 4):
            esperado.update(_tokens(str(ws.Cells(r, c).Text)))
    obtido = Counter()
    doc = fitz.open(pdf)
    for pg in doc: obtido.update(_tokens(pg.get_text()))
    doc.close()
    return esperado - obtido


def _crop_ink(png, pad=0):
    """Recorta exatamente o retângulo com tinta (bordas da tabela). pad=0 preserva a
    equivalência 1:1 entre largura da tabela e largura da imagem."""
    from PIL import Image
    im = Image.open(png).convert("RGB")
    bbox = im.point(lambda v: 255 if v < 250 else 0).convert("L").getbbox()
    if not bbox: return None
    x0, y0, x1, y1 = bbox
    w, h = im.size
    box = (max(0, x0 - pad), max(0, y0 - pad), min(w, x1 + pad), min(h, y1 + pad))
    im.crop(box).save(png)
    return (box[2] - box[0], box[3] - box[1])


def pngs_do_pdf(pdf, tag, k, WK, pasta):
    import fitz
    doc = fitz.open(pdf); saidas = []
    for i, pg in enumerate(doc):
        sufixo = "" if i == 0 else chr(97 + i)
        png = os.path.join(WK, "03_IMAGENS", pasta, f"{tag}_p{k:02d}{sufixo}.png")
        pg.get_pixmap(dpi=DPI).save(png)
        _crop_ink(png)
        saidas.append(png)
    if doc.page_count > 1:
        cfg.log(f"  AVISO [{tag} p{k}] a parte gerou {doc.page_count} paginas no Excel")
    doc.close()
    return saidas


def avaliar(wb, sheet, tag, WK, geo, logo_png):
    """Formata, pagina e RENDERIZA a tabela numa orientação; devolve as partes, os PDFs e a
    MÉTRICA DE DESPERDÍCIO (polegadas de página em branco que a tabela deixa no corpo).

    Desperdício = espaço perdido na página do texto + sobras das partes:
      - RETRATO: a tabela começa logo abaixo do texto, então nada se perde ali; a sobra da ÚLTIMA
        parte não conta (o texto da seção seguinte flui logo abaixo dela).
      - PAISAGEM: a mudança de orientação obriga página nova, então TODO o espaço que sobrou abaixo
        do texto é perdido; e a sobra da última parte também é perdida (a seção seguinte volta ao
        retrato em página nova)."""
    ws = wb.Worksheets(sheet)
    alvo_pt = geo["larg_in"] * PT_IN
    info = formatar(wb, ws, tag, WK, alvo_pt, logo_png, geo["prop"], geo["orient"])
    if not info:
        cfg.log(f"[{sheet}] vazio"); return None
    last, larg_pt, h_tit = info["last"], info["larg_pt"], info["h_titulos"]
    alturas, rep, fy = info["alturas"], info["rep"], info["fy"]
    ordem = list(range(4, last + 1))
    max_1 = geo["disp1"] - h_tit
    max_n = geo["dispN"] - h_tit

    partes, pdfs = [], []
    for passo in range(1, 4):
        partes = paginar(alturas, ordem, max_1, max_n)
        pdfs, faltas = [], {}
        for k, linhas in enumerate(partes, start=1):
            pdf = exportar_pdf_parte(wb, ws, linhas, tag, k, WK, larg_pt, geo["orient"])
            pdfs.append((pdf, linhas))
            for r, f in cortes_no_pdf(pdf, linhas, larg_pt).items():
                faltas[r] = max(faltas.get(r, 0.0), f)
            aus = texto_ausente(pdf, ws, linhas)
            if aus:
                culpadas = []
                for r in linhas:
                    t = Counter()
                    for c in (2, 3, 4): t.update(_tokens(str(ws.Cells(r, c).Text)))
                    if t & aus: culpadas.append(r)
                for r in (culpadas or linhas):
                    faltas[r] = max(faltas.get(r, 0.0), FONTE_CONTEUDO_PT * 1.3)
        if not faltas:
            break
        for r, f in sorted(faltas.items()):
            ws.Rows(r).RowHeight = min(LIM_ROWHEIGHT_PT, rep[r] + max(2.0, f / fy))
            rep[r] = float(ws.Rows(r).Height); alturas[r] = rep[r] * fy
    else:
        cfg.log(f"  AVISO [{tag}] ainda havia texto cortado apos 3 ajustes")

    sobras, desperdicio = [], geo["perda_pagina_texto"]
    for k, linhas in enumerate(partes, start=1):
        lim = max_1 if k == 1 else max_n
        sobra = lim - sum(alturas[r] for r in linhas)
        sobras.append(sobra)
        if k < len(partes) or geo["conta_ultima"]:
            desperdicio += max(0.0, sobra)
    cfg.log(f"  [{tag}] {geo['nome'].upper()}: {sum(alturas.values())/PT_IN:.2f}in de conteudo em "
            f"{len(partes)} parte(s) de {geo['larg_in']:.2f}in | ocupacao "
            + " + ".join(f"{100*(1-so/(max_1 if i == 0 else max_n)):.0f}%"
                         for i, so in enumerate(sobras))
            + f" | DESPERDICIO {desperdicio/PT_IN:.2f}in")
    return dict(partes=partes, pdfs=pdfs, larg_pt=larg_pt, h_tit=h_tit, alturas=alturas,
                sobras=sobras, desperdicio=desperdicio, geo=geo)


def publicar(res, tag, pasta, WK, geom):
    """Gera os PNG definitivos da orientação vencedora e o bloco do layout para o 09."""
    for f in glob.glob(os.path.join(WK, "03_IMAGENS", pasta, "*.png")):
        os.remove(f)
    geo = res["geo"]; saidas = []
    for k, (pdf, linhas) in enumerate(res["pdfs"], start=1):
        for png in pngs_do_pdf(pdf, tag, k, WK, pasta):
            from PIL import Image
            w, h = Image.open(png).size
            saidas.append(dict(arquivo=os.path.basename(png), reunioes=len(linhas), px=[w, h],
                               altura_in=round(h / w * geo["larg_in"], 3)))
            cfg.log(f"  [{tag}] {os.path.basename(png)}: {len(linhas)} reuniao(oes) -> "
                    f"{geo['larg_in']:.3f} x {saidas[-1]['altura_in']:.2f}in")
    open(os.path.join(WK, "03_IMAGENS", pasta, ".precropped"), "w").close()
    return dict(orientacao=geo["nome"], largura_in=round(geo["larg_in"], 4),
                largura_tabela_pt=round(res["larg_pt"], 2),
                escala=round(geo["larg_in"] * PT_IN / res["larg_pt"], 4),
                fonte_conteudo_pt=FONTE_CONTEUDO_PT, fonte_cabecalho_pt=FONTE_CABECALHO_PT,
                desperdicio_in=round(res["desperdicio"] / PT_IN, 2),
                espaco_primeira_in=round(geo["disp1"] / PT_IN, 3), partes=saidas)


def geometrias(geom, disp_retrato):
    """As duas candidatas. Em ambas a largura da tabela = largura útil do corpo daquela página."""
    cabe = disp_retrato >= MIN_PRIMEIRA_IN * PT_IN
    return [
        dict(nome="retrato", larg_in=geom["larg_in"], prop=PROP_RETRATO, orient=1,
             disp1=disp_retrato if cabe else geom["alt_pagina_pt"], dispN=geom["alt_pagina_pt"],
             perda_pagina_texto=0.0 if cabe else disp_retrato, conta_ultima=False),
        dict(nome="paisagem", larg_in=LARG_PAISAGEM_IN, prop=PROP_PAISAGEM, orient=2,
             disp1=ALT_PAISAGEM_IN * PT_IN, dispN=ALT_PAISAGEM_IN * PT_IN,
             perda_pagina_texto=disp_retrato, conta_ultima=True),
    ]


def escolher(wb, sheet, base_tag, pasta, WK, geom, disp_retrato, logo_png):
    """Testa as orientações candidatas e publica a de MENOR desperdício (ORIENTACAO=auto);
    'retrato'/'paisagem' forçam uma delas."""
    cands = [g for g in geometrias(geom, disp_retrato)
             if ORIENTACAO in ("auto", g["nome"])]
    res = []
    for g in cands:
        r = avaliar(wb, sheet, f"{base_tag}_{g['nome'][:3]}", WK, g, logo_png)
        if r: res.append(r)
    if not res: return None
    melhor = min(res, key=lambda r: (round(r["desperdicio"], 1), len(r["partes"])))
    if len(res) > 1:
        cfg.log(f"  [{base_tag}] ESCOLHIDA: {melhor['geo']['nome'].upper()} "
                + " | ".join(f"{r['geo']['nome']} {r['desperdicio']/PT_IN:.2f}in em "
                             f"{len(r['partes'])} pag." for r in res))
    return publicar(melhor, base_tag, pasta, WK, geom)


# ---------------------------------------------------------------- prévia (validação visual)
def gerar_previa(WK, geom, titulo, paragrafos, lay):
    """Renderiza no Word a página real (texto + tabela) para validação visual do PDF —
    critério de aprovação exigido: o que vale é o PDF renderizado, não a planilha."""
    if lay.get("orientacao") == "paisagem":
        cfg.log("  [prévia] tabela em PAISAGEM (seção própria): conferir no PDF final")
        return None
    try:
        d, tmp = _doc_esqueleto(WK, "previa_reunioes.docx")
        _titulo_par(d, titulo)
        for t in paragrafos: _corpo_par(d, t)
        for p in lay["partes"]:
            _img_par(d, os.path.join(WK, "03_IMAGENS", "REUNIOES_P11", p["arquivo"]), lay["largura_in"])
        d.save(tmp)
        out = os.path.join(WK, "04_PDFS_TEMP", "previa_reunioes_p11.pdf")
        wd = _word()
        try:
            doc = wd.Documents.Open(os.path.abspath(tmp)); doc.Repaginate()
            if os.path.exists(out):
                try: os.remove(out)   # não sobrescrever PDF pré-existente (DEC-038)
                except OSError: pass
            doc.ExportAsFixedFormat(os.path.abspath(out), 17)
            npag = int(doc.ComputeStatistics(2)); doc.Close(0)
        finally:
            wd.Quit()
        cfg.log(f"  [prévia] {out} ({npag} páginas)")
        _validar_previa(out, geom, lay)
        return out
    except Exception as e:
        cfg.log(f"  [prévia] não gerada ({e})")
        return None


def _validar_previa(pdf, geom, lay):
    """Conferência programática no PDF: largura da imagem = largura do corpo, bordas alinhadas
    às margens do texto e nada invadindo a área do rodapé."""
    import fitz
    doc = fitz.open(pdf)
    lim_inf = geom["pag_h_pt"] - geom["bottom_pt"]
    ok = True
    for i, pg in enumerate(doc):
        for b in pg.get_image_info():
            x0, y0, x1, y1 = b["bbox"]
            larg = (x1 - x0) / PT_IN
            if abs(larg - lay["largura_in"]) > 0.02:
                cfg.log(f"  [QA prévia] pág.{i+1}: largura da imagem {larg:.3f}in "
                        f"≠ largura do corpo {lay['largura_in']:.3f}in"); ok = False
            if y1 > lim_inf + 1:
                cfg.log(f"  [QA prévia] pág.{i+1}: imagem invade a área do rodapé "
                        f"(y={y1/PT_IN:.2f}in)"); ok = False
    doc.close()
    cfg.log("  [QA prévia] " + ("alinhamento e limites OK" if ok else "REVISAR (ver avisos acima)"))
    return ok


# ---------------------------------------------------------------- main
def main():
    mes = cfg.mes_min(); WK = cfg.work_dir()
    ANEX = cfg.find_in_dir(cfg.anexo_p12_dir(), "p1.2", "controle de reuni", ext=".xlsx") \
        or os.path.join(cfg.anexo_p12_dir(), cfg.nome_anexo_g1("P1.2", f"Controle de Reuniões do mês de {mes}", ".xlsx"))
    if not os.path.exists(ANEX):
        cfg.log(f"ERRO: anexo de reuniões não encontrado: {ANEX}"); sys.exit(1)
    WORK = os.path.join(WK, "02_TRABALHO", "reunioes_print.xlsx")
    shutil.copy2(ANEX, WORK)
    logo_png = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                            "..", "assets", "logo_prog.png"))

    geom = geometria_corpo()
    cfg.log(f"geometria do corpo ({geom['fonte']}): largura útil {geom['larg_in']:.3f}in, "
            f"altura útil {geom['alt_pagina_pt']/PT_IN:.2f}in | fonte do conteúdo "
            f"{FONTE_CONTEUDO_PT}pt (visual, escala 1:1)")
    titulo = f"Reuniões de Governança realizadas durante o mês de {mes}."
    try:
        paragrafos = list(cfg.dados().REUNIOES)
    except Exception as e:
        paragrafos = []; cfg.log(f"  [espaço] textos do P1.1 indisponíveis ({e})")
    disp = espaco_real_no_montado(WK, geom, titulo)
    if disp is None:
        disp = medir_espaco_apos_texto(WK, geom, titulo, paragrafos) if paragrafos \
            else geom["alt_pagina_pt"]

    xl = win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    try: xl.AutomationSecurity = 1
    except Exception: pass
    atexit.register(lambda: xl.Quit())
    wb = xl.Workbooks.Open(WORK, False, False)

    def pick(cands, default):
        names = [wb.Worksheets(k + 1).Name for k in range(wb.Worksheets.Count)]
        for c in cands:
            for n in names:
                if c.lower() in n.lower(): return n
        return default
    gov  = pick(["gov 1-1", "reuniões de gov", "reunioes de gov"], "Controle Geral")
    inst = pick(["stakehold"], "Reuniões Stakehold 1-2")
    cfg.log(f"aba governança P1.1: '{gov}' | institucional P1.2: '{inst}'")

    lay = {}
    l11 = escolher(wb, gov, "gov_p11", "REUNIOES_P11", WK, geom, disp, logo_png)
    if l11: lay["gov_p11"] = l11
    try:
        # P1.2: o texto que antecede a tabela não começa no topo da página -> página cheia
        l12 = escolher(wb, inst, "inst_p12", "REUNIOES_P12", WK, geom,
                       geom["alt_pagina_pt"], logo_png)
        if l12: lay["inst_p12"] = l12
    except Exception as e:
        cfg.log(f"aba institucional: {e}")

    # Save do xlsx de trabalho é NÃO CRÍTICO (as imagens PNG já foram geradas; downstream usa
    # as imagens + o reunioes_layout.json). Não deixar um erro de Save/COM abortar a etapa (DEC-038).
    try: wb.Save()
    except Exception as e: cfg.log(f"aviso: Save do xlsx de reuniões falhou (não crítico): {e}")
    try: wb.Close(False)
    except Exception: pass
    try: xl.Quit()
    except Exception: pass
    with open(os.path.join(WK, "02_TRABALHO", "reunioes_layout.json"), "w", encoding="utf-8") as f:
        json.dump(lay, f, ensure_ascii=False, indent=1)

    if l11 and paragrafos and not os.environ.get("PROG_REUNIOES_SEM_PREVIA"):
        gerar_previa(WK, geom, titulo, paragrafos, l11)
    cfg.log("prints de reuniões concluídos")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
05_imagens_financeiro.py — Gera as imagens (PNG) do Quadro Financeiro a partir do
anexo 'P1.1 - Quadro Financeiro do mês de {mês}.xlsx' (SharePoint, leitura).
As abas trazem sufixo do nº da medição (ex.: _M8, _1A8) que MUDA a cada mês; por isso
as abas são localizadas por PREFIXO, não por nome fixo. Só escreve arquivos locais.
Parametrizado por prog_config.
"""
import os, sys, shutil, unicodedata, re
import win32com.client as win32
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

def _norm(s):
    return "".join(c for c in unicodedata.normalize("NFD",str(s)) if unicodedata.category(c)!="Mn").upper()

def _base_cumheights(pdf_path, placement_w=6.8):
    """Mede, no print da Base Técnica, a altura CUMULATIVA (em polegadas, já na escala em que a
    imagem é inserida no corpo) até o fim de cada produto. Usa a coluna ID ('1.1','2.3'...) para
    localizar cada linha. Retorna [] se não achar as linhas. Base da divisão em 2 (DEC-031)."""
    import fitz
    doc=fitz.open(pdf_path); pg=doc[0]; words=pg.get_text("words"); doc.close()
    ids=sorted([(w[1],w[3]) for w in words if re.match(r'^\d+\.\d+$', w[4])])
    if not ids: return []
    x0=min(w[0] for w in words); x1=max(w[2] for w in words)
    y0=min(w[1] for w in words); y1=max(w[3] for w in words)
    if x1<=x0 or y1<=y0: return []
    aspect=(x1-x0)/(y1-y0)
    place_w=min(placement_w, 8.5*aspect); place_h=place_w/aspect   # = mesma regra do fit_img (09)
    return [ (yb-y0)/(y1-y0)*place_h for (_yt,yb) in ids ]

def find_sheet(wb, prefix):
    p=_norm(prefix)
    for ws in wb.Worksheets:
        if _norm(ws.Name).startswith(p): return ws
    for ws in wb.Worksheets:
        if p in _norm(ws.Name): return ws
    return None

def main():
    WK=cfg.work_dir()
    SRC=cfg.find_in_dir(cfg.anexo_p11_dir(), "p1.1", "quadro financeiro", ext=".xlsx")
    if not SRC:
        # POP (DEC-066): na pasta do mês o painel fica como ZIP individual -> extrai a planilha para 02_TRABALHO
        z=cfg.find_in_dir(cfg.anexo_p11_dir(), "p1.1", "quadro financeiro", ext=".zip")
        if z:
            import zipfile
            with zipfile.ZipFile(z) as zf:
                xs=[n for n in zf.namelist() if n.lower().endswith(".xlsx")]
                if xs: zf.extract(xs[0], os.path.join(WK,"02_TRABALHO")); SRC=os.path.join(WK,"02_TRABALHO",xs[0])
    if not SRC or not os.path.exists(SRC):
        cfg.log("ERRO: anexo 'Quadro Financeiro' (.xlsx) não encontrado em P1.1."); sys.exit(1)
    WORK=os.path.join(WK,"02_TRABALHO","Quadro_Financeiro_WORK.xlsx"); shutil.copy2(SRC,WORK)

    xl=win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible=False; xl.DisplayAlerts=False; xl.ScreenUpdating=False
    wb=xl.Workbooks.Open(WORK)
    # (prefixo da aba, orientação 1=retrato/2=paisagem, tag da imagem)
    plan=[("PAINEL_FINANCEIRO",1,"quadro_consolidado"),
          ("RESUMO_EMPRESAS",2,"resumo_empresas"),
          ("BASE_PARCEIRO",1,"base_parceiro"),
          ("STATUS_MEDICOES",2,"status_medicoes")]
    import fitz
    outdir=os.path.join(WK,"03_IMAGENS","FINANCEIRO")
    for f in os.listdir(outdir):
        if f.endswith(".png"): os.remove(os.path.join(outdir,f))
    for prefix,orient,tag in plan:
        ws=find_sheet(wb,prefix)
        if ws is None:
            cfg.log(f"aviso: aba com prefixo '{prefix}' não encontrada"); continue
        ur=ws.UsedRange; ps=ws.PageSetup
        ps.Orientation=orient; ps.Zoom=False; ps.FitToPagesWide=1; ps.FitToPagesTall=False
        area=ur.Address
        # PAINEL_FINANCEIRO: o print deve terminar no FIM da tabela "RESUMO FINANCEIRO POR
        # EMPRESA" (linha do TOTAL). A aba tem, mais abaixo, uma tabela auxiliar (rascunho que
        # alimenta gráficos) que NÃO pode sair no relatório. Detecção por conteúdo, sem linha
        # fixa (a posição muda a cada mês). DEC-029.
        if _norm(prefix).startswith("PAINEL_FINANCEIRO"):
            first_row=ur.Row; first_col=ur.Column
            last_col=first_col+ur.Columns.Count-1; rmax=first_row+ur.Rows.Count-1
            rr=None
            for r in range(first_row, rmax+1):
                line=""
                for c in range(first_col,last_col+1):
                    v=ws.Cells(r,c).Value
                    if v is not None: line+=" "+_norm(v)
                if "RESUMO FINANCEIRO POR EMPRESA" in line: rr=r; break
            if rr:
                bcol=first_col+1  # coluna "Empresa"/"TOTAL" do RESUMO (2ª col da área usada)
                last_row=rr; r=rr+1
                while r<=rmax:  # anda enquanto a tabela é contígua; para na 1ª linha vazia (exclui rascunho)
                    v=ws.Cells(r,bcol).Value
                    if v is None or str(v).strip()=="": break
                    last_row=r; r+=1
                area=f"{ws.Cells(first_row,first_col).Address}:{ws.Cells(last_row,last_col).Address}"
                cfg.log(f"  [PAINEL] print recortado no fim do RESUMO POR EMPRESA "
                        f"(título linha {rr} -> TOTAL linha {last_row}); rascunho abaixo excluído")
            else:
                cfg.log("  [PAINEL] AVISO: 'RESUMO FINANCEIRO POR EMPRESA' não localizado; usando UsedRange.")
        def _export_render(area_i, titlerows, out_tag):
            ps.PrintArea=area_i; ps.PrintTitleRows=titlerows
            ps.LeftMargin=xl.InchesToPoints(0.2); ps.RightMargin=xl.InchesToPoints(0.2)
            ps.TopMargin=xl.InchesToPoints(0.3); ps.BottomMargin=xl.InchesToPoints(0.3)
            ps.CenterHorizontally=True
            pdf=os.path.join(WK,"04_PDFS_TEMP",f"fin_{out_tag}.pdf")
            if os.path.exists(pdf):
                try: os.remove(pdf)   # Excel NÃO sobrescreve PDF pré-existente -> "documento não foi salvo" (DEC-038)
                except OSError: pass
            ws.ExportAsFixedFormat(0,pdf)
            doc=fitz.open(pdf); n=doc.page_count
            for i,pg in enumerate(doc):
                suffix="" if n==1 else f"_p{i+1}"
                pg.get_pixmap(dpi=200).save(os.path.join(outdir,f"{out_tag}{suffix}.png"))
            cfg.log(f"  {ws.Name} ({area_i}) -> {out_tag}: {n} pág."); doc.close(); return pdf

        if not _norm(prefix).startswith("BASE_PARCEIRO"):
            _export_render(area, "", tag); continue

        # ===== BASE_PARCEIRO: divide em 2 p/ minimizar o branco antes da quebra de página (DEC-031) =====
        # A parte 1 preenche o resto da página (logo abaixo do texto) e a parte 2 (título+cabeçalho
        # repetidos) segue na próxima. O nº de produtos da parte 1 é MEDIDO da própria tabela:
        # entra o máximo de produtos cuja altura cumulativa <= alvo (PROG_BASE_PART1_IN, default 5,25").
        fr=ur.Row; fc=ur.Column; lc=fc+ur.Columns.Count-1; lr=fr+ur.Rows.Count-1
        title_row=header_row=None
        for r in range(fr, lr+1):
            rowtxt=""
            for c in range(fc,lc+1):
                v=ws.Cells(r,c).Value
                if v is not None: rowtxt+=" "+_norm(v)
            if title_row is None and "BASE TECNICA" in rowtxt: title_row=r
            if header_row is None and "PRODUTO" in rowtxt and "UNIDADE" in rowtxt: header_row=r
            if header_row: break
        base_pdf=_export_render(area, "", tag)      # imagem única (serve p/ medir as alturas)
        cum=_base_cumheights(base_pdf)
        n1=0
        if header_row and cum:
            target=float(os.environ.get("PROG_BASE_PART1_IN","5.25"))
            for i,ch in enumerate(cum):
                if ch<=target: n1=i+1
            if 0<n1<len(cum):
                prod0=header_row+1; split=prod0+n1-1
                tr=f"${title_row or header_row}:${header_row}"   # repete título+cabeçalho em cada parte
                a1=f"{ws.Cells(prod0,fc).Address}:{ws.Cells(split,lc).Address}"
                a2=f"{ws.Cells(split+1,fc).Address}:{ws.Cells(lr,lc).Address}"
                _export_render(a1, tr, f"{tag}_p1")
                _export_render(a2, tr, f"{tag}_p2")
                try: os.remove(os.path.join(outdir,f"{tag}.png"))   # descarta a imagem única
                except OSError: pass
                cfg.log(f"  [BASE] dividida em 2 (DEC-031): alvo p1={target:.2f}in -> parte1={n1} produto(s) "
                        f"(~{cum[n1-1]:.2f}in), parte2={len(cum)-n1} produto(s); título=lin{title_row}, cab=lin{header_row}")
            else:
                cfg.log(f"  [BASE] imagem única (n1={n1} de {len(cum)} produtos; alvo {target:.2f}in)")
        else:
            cfg.log("  [BASE] imagem única (não foi possível medir linhas/cabeçalho).")
    wb.Close(SaveChanges=False); xl.Quit()
    cfg.log("imagens do financeiro concluídas")

if __name__=="__main__":
    main()

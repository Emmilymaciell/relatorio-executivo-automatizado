# -*- coding: utf-8 -*-
"""
12_finalizar_pdf.py — Finaliza o DOCX montado: atualiza campos + sumário (TOC),
repagina, exporta PDF, e versiona DOCX+PDF na pasta G1 do mês (sem sobrescrever).
Parametrizado via prog_config (config/parametros.json). Não altera originais.

DRY_RUN (PROG_DRY_RUN=1): gera DOCX/PDF apenas LOCALMENTE (07_SAIDA); NÃO copia/sobrescreve
nada no G1 do SharePoint (registra "DRY_RUN: publicação no SharePoint ignorada"). Ver DEC-021.

Uso:  python 12_finalizar_pdf.py <caminho_docx_montado>
      (se omitido, usa output/RM.../02_TRABALHO/RM_montado.docx)

Este script serve também de PADRÃO de parametrização/robustez para os demais.
"""
import os, sys, atexit
import win32com.client as win32
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

def uniq(path):
    if not os.path.exists(path): return path
    b, e = os.path.splitext(path); i = 2
    while os.path.exists(f"{b}_v{i}{e}"): i += 1
    return f"{b}_v{i}{e}"

def _publicar_g1(g1, docx_out, pdf_out, base):
    """Versiona DOCX+PDF na pasta G1 do mês (SharePoint), sem sobrescrever (usa uniq -> _v2/_v3...).
    Respeita PROG_DRY_RUN (DEC-021): em DRY_RUN NÃO copia nada para o SharePoint. Retorna a lista de
    destinos efetivamente escritos (vazia em DRY_RUN ou se G1 não existir)."""
    import shutil
    if cfg.dry_run():
        cfg.log("DRY_RUN: publicação no SharePoint ignorada"); return []
    if not (g1 and os.path.isdir(g1)):
        cfg.log(f"AVISO: pasta G1 não encontrada ({g1}); DOCX/PDF ficaram em 07_SAIDA."); return []
    done = []
    for ext, src in [(".docx", docx_out), (".pdf", pdf_out)]:
        dst = uniq(os.path.join(g1, base + ext)); shutil.copy2(src, dst); done.append(dst)
        cfg.log(f"G1 <- {os.path.basename(dst)}")
    return done

def main():
    P = cfg.params(); WK = cfg.work_dir()
    montado = sys.argv[1] if len(sys.argv) > 1 else os.path.join(WK, "02_TRABALHO", f"RM{cfg.numero()}_montado.docx")
    if not os.path.exists(montado):
        cfg.log(f"ERRO: DOCX montado não encontrado: {montado}"); sys.exit(1)
    base = cfg.doc_final_basename()
    out = os.path.join(WK, "07_SAIDA"); os.makedirs(out, exist_ok=True)
    docx_out = os.path.join(out, base + ".docx"); pdf_out = os.path.join(out, base + ".pdf")

    wd = win32.gencache.EnsureDispatch("Word.Application"); wd.Visible = False; wd.DisplayAlerts = 0
    atexit.register(lambda: (wd.Quit() if wd else None))
    d = wd.Documents.Open(montado)
    try:
        for st in d.StoryRanges:
            try: st.Fields.Update()
            except Exception: pass
        for toc in d.TablesOfContents: toc.Update()
        d.Repaginate()
        cfg.log(f"páginas: {d.ComputeStatistics(2)}")
        for _f in (docx_out, pdf_out):     # não sobrescrever arquivo pré-existente (DEC-038)
            if os.path.exists(_f):
                try: os.remove(_f)
                except OSError: pass
        d.SaveAs2(docx_out); d.SaveAs2(pdf_out, FileFormat=17)  # 17 = wdFormatPDF
    finally:
        d.Close(False); wd.Quit()
    cfg.log(f"SAÍDA: {docx_out} | {pdf_out}")

    # versionar na pasta G1 do mês (nunca sobrescrever; respeita PROG_DRY_RUN — DEC-021)
    _publicar_g1(cfg.g1_dir(), docx_out, pdf_out, base)

if __name__ == "__main__":
    main()

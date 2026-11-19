# -*- coding: utf-8 -*-
"""
19_dashboard_imagem.py — Gera a imagem do Dashboard (dashboard_p01.png) a partir do PDF do Power BI
que o USUÁRIO exportou e colocou em G1\\Anexos\\P1.1 (P1.1 - Dashboard da Análise de Entrega de
Produtos - {mês}.pdf). Etapa OBRIGATÓRIA antes do 09 (montar) — foi esquecida no RM14/agosto e o
Dashboard não entrou no relatório; agora é um passo explícito do fluxo (DEC-040). O 18 (revisão
final) acusa se a imagem faltar.

Só lê o PDF e grava um PNG local (work/03_IMAGENS/DASHBOARD/dashboard_p01.png). dpi=200. Parametrizado.
Se o PDF do Dashboard não existir no anexo P1.1, PARA com aviso (peça o export do Power BI ao usuário).
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
import fitz

def main():
    WK=cfg.work_dir()
    src=cfg.find_in_dir(cfg.anexo_p11_dir(), "p1.1", "dashboard", ext=".pdf")
    if not src or not os.path.exists(src):
        cfg.log("ERRO: PDF do Dashboard não encontrado em G1\\Anexos\\P1.1. O usuário precisa EXPORTAR "
                "o Dashboard do Power BI (o acumulado já está atualizado) e salvar lá. (checkpoint)")
        cfg.set_qa_warning("dashboard_pdf_ausente", "Dashboard PDF ausente — usuário deve exportar do Power BI.")
        sys.exit(2)
    outdir=os.path.join(WK,"03_IMAGENS","DASHBOARD"); os.makedirs(outdir,exist_ok=True)
    for f in os.listdir(outdir):
        if f.endswith(".png"): os.remove(os.path.join(outdir,f))
    d=fitz.open(src)
    d[0].get_pixmap(dpi=200).save(os.path.join(outdir,"dashboard_p01.png"))
    cfg.log(f"dashboard_p01.png gerado de: {os.path.basename(src)} ({d.page_count} pág. no PDF)")
    d.close()
    cfg.set_qa_warning("dashboard_pdf_ausente", "", active=False)

if __name__=="__main__":
    main()

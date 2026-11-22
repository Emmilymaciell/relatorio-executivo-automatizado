# -*- coding: utf-8 -*-
"""
20_zip_anexos_g1.py — POP Quadro 4 (DEC-066): os anexos GERADOS pela Coordenação (P1.1/P1.2) que não são PDF
(Controle de Entrega, Quadro Financeiro, Controle de Reuniões — .xlsx) vão para a pasta ANEXOS do G1 como ZIP
INDIVIDUAL com o mesmo nome. O .xlsx sai da pasta de anexos e fica guardado em 07_SAIDA (fonte editável).
Só atua em arquivos P1.1/P1.2 (nunca mexe em P1.3/P1.4 da CG — DEC-042). Idempotente; layout 'plano' apenas.
Rodar DEPOIS do 03/05/15/04 (que leem os .xlsx) e ANTES do 10 (Quadro de Anexos com o nome final).
"""
import os, shutil, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg


def main():
    if cfg.layout_anexos() != "plano":
        cfg.log("20: layout legado — nada a zipar"); return
    d = cfg.anexo_p11_dir(); saida = os.path.join(cfg.work_dir(), "07_SAIDA")
    feitos = []
    for f in sorted(os.listdir(d)):
        p = os.path.join(d, f)
        if not os.path.isfile(p) or not f[:4] in ("P1.1", "P1.2") or f.lower().endswith((".pdf", ".zip")):
            continue
        z = os.path.join(d, os.path.splitext(f)[0] + ".zip")
        with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(p, arcname=f)
        with zipfile.ZipFile(z) as zf:
            if zf.testzip() is not None:
                raise RuntimeError(f"ZIP corrompido: {z}")
        shutil.copy2(p, os.path.join(saida, f))
        os.remove(p)
        feitos.append(os.path.basename(z))
    cfg.log(f"20: {len(feitos)} anexo(s) da Coordenação em ZIP individual: {feitos}")


if __name__ == "__main__":
    main()

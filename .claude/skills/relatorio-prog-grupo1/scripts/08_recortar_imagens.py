# -*- coding: utf-8 -*-
"""
08_recortar_imagens.py — Autocrop das bordas brancas de TODAS as imagens geradas e
gravação de img_aspects.json (proporções) para o 09 dimensionar cada figura à área útil.
Só escreve arquivos locais. Parametrizado por prog_config (nenhum caminho fixo).
"""
import os, sys, json
from PIL import Image
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

def autocrop(path, pad=8, thr=244):
    im=Image.open(path).convert("RGB"); w,h=im.size; px=im.load()
    minx,miny,maxx,maxy=w,h,0,0; step=2
    for y in range(0,h,step):
        for x in range(0,w,step):
            r,g,b=px[x,y]
            if r<thr or g<thr or b<thr:
                if x<minx:minx=x
                if x>maxx:maxx=x
                if y<miny:miny=y
                if y>maxy:maxy=y
    if maxx<=minx or maxy<=miny: return None
    minx=max(0,minx-pad); miny=max(0,miny-pad); maxx=min(w,maxx+pad); maxy=min(h,maxy+pad)
    im2=im.crop((minx,miny,maxx,maxy)); im2.save(path); return im2.size

def main():
    WK=cfg.work_dir(); IMG=os.path.join(WK,"03_IMAGENS")
    folders=["CONTROLE_PRODUTOS","DASHBOARD","TIMELINE","REUNIOES_P11","REUNIOES_P12","FINANCEIRO"]
    aspects={}
    for fol in folders:
        d=os.path.join(IMG,fol)
        if not os.path.isdir(d): continue
        # Pastas com marcador '.precropped' já foram recortadas EXATAMENTE na tinta pelo script
        # que as gerou (ex.: 04 nas tabelas de reuniões, escala 1:1). Recortar de novo com 'pad'
        # acrescentaria borda branca e quebraria o alinhamento da tabela com o corpo do texto.
        pre=os.path.exists(os.path.join(d,".precropped"))
        for f in sorted(os.listdir(d)):
            if not f.lower().endswith(".png"): continue
            p=os.path.join(d,f)
            sz=Image.open(p).size if pre else autocrop(p)
            if sz:
                aspects[f"{fol}/{f}"]={"w":sz[0],"h":sz[1],"aspect":round(sz[0]/sz[1],4)}
                cfg.log(f"{fol}/{f}: {sz[0]}x{sz[1]} aspect={sz[0]/sz[1]:.3f}"
                        + (" (pré-recortada)" if pre else ""))
    json.dump(aspects, open(os.path.join(WK,"02_TRABALHO","img_aspects.json"),"w"), indent=1)
    cfg.log(f"img_aspects.json salvo: {len(aspects)} imagens")

if __name__=="__main__":
    main()

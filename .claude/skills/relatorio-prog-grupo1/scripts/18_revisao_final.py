# -*- coding: utf-8 -*-
"""
18_revisao_final.py — REVISÃO FINAL (última etapa do fluxo, após o 13/QA). Vai além do QA de texto:
confere que NADA ficou faltando na MONTAGEM e nos ANEXOS antes de considerar o relatório pronto.
Não escreve nos arquivos oficiais; só lê o PDF final, as imagens do corpo (work) e as pastas de anexos,
e grava um relatório de revisão local. Parametrizado (nada fixo da competência).

Checagens:
 1. Imagens do corpo (work/03_IMAGENS): cada pasta esperada tem ≥1 PNG — pega, p.ex., o Dashboard
    esquecido (pasta DASHBOARD vazia).
 2. PDF final: existe; nº do relatório presente e o anterior ausente; período/mês; cabeçalho+rodapé em
    TODAS as páginas; cada "Produto 1.X" no topo de página nova; texto consolidado do P1.2;
    Dashboard (página com imagem grande após o título "Dashboard Analítico"); Timeline (nº de páginas
    de imagem coerente com o nº de meses); valores/contagens esperados (cfg.dados().QA_ESPERADOS);
    status da medição (params.status_medicao) presente; datas de reuniões do mês presentes.
 3. Anexos físicos: lê a "LISTA DE DOCUMENTOS ANEXOS" do PDF e confere que CADA arquivo listado existe
    de fato nas pastas G1\\Anexos (P1.1/P1.2/P1.3/P1.4). Pega anexo faltando/renomeado.
"""
import os, sys, re, glob, datetime
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
import fitz

def _norm(s):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD",str(s or "")) if unicodedata.category(c)!="Mn").lower()

def main():
    P=cfg.params(); N=cfg.numero(); mes=cfg.mes_min(); ano=cfg.ano(); WK=cfg.work_dir()
    D=cfg.dados()
    res=[]
    def chk(ok,label,detail=""): res.append(((("OK " if ok else "!! ")+label), detail))

    # ---------- 1) imagens do corpo ----------
    IMG=os.path.join(WK,"03_IMAGENS")
    for fol in ["CONTROLE_PRODUTOS","DASHBOARD","TIMELINE","REUNIOES_P11","REUNIOES_P12","FINANCEIRO"]:
        pngs=glob.glob(os.path.join(IMG,fol,"*.png"))
        chk(len(pngs)>0, f"imagens do corpo: {fol}", f"{len(pngs)} PNG" if pngs else "PASTA VAZIA — imagem não gerada!")

    # ---------- 2) PDF final ----------
    PDF=os.path.join(WK,"07_SAIDA",cfg.doc_final_basename()+".pdf")
    if not os.path.exists(PDF):
        chk(False,"PDF final existe",PDF)
        _grava(res,WK,N); _veredito(res); return
    doc=fitz.open(PDF); Np=doc.page_count
    full="".join(doc[i].get_text() for i in range(Np)); low=full.lower()
    chk(True,"PDF final existe",f"{Np} páginas")
    chk((f"N° {N}" in full or f"Nº {N}" in full or f"N°{N}" in full),f"contém 'N° {N}'")
    chk(all(f"{t} {s} {N-1}" not in full for t in ("Relatório Mensal","Relatório de Produtos Mensal") for s in ("N°","Nº")),
        f"sem nº de relatório antigo (N° {N-1})")
    # cabeçalho+rodapé em todas as páginas
    semhf=[i+1 for i in range(Np) if ("CONSÓRCIO" not in doc[i].get_text()) or ("CONFIDENCIAL" not in doc[i].get_text())]
    chk(not semhf,"cabeçalho+rodapé em TODAS as páginas", f"faltando em {semhf}" if semhf else "")
    # cada Produto 1.X no topo
    for x in ["1.1","1.2","1.3","1.4"]:
        loc=None
        for i in range(Np):
            rs=doc[i].search_for(f"Produto {x} – Relat")
            if rs: loc=(i,min(r.y0 for r in rs))
        if loc: chk(loc[1]<=0.30*doc[loc[0]].rect.height,f"'Produto {x}' no topo de página nova",f"pág {loc[0]+1}")
        else: chk(False,f"'Produto {x}' localizado")
    # texto consolidado P1.2
    chk("relatório de relacionamento institucional do mês de" in low,"P1.2: texto consolidado presente")
    # Dashboard: página com imagem grande logo após o título (busca normalizada, sem acento)
    dash_ok=False; alvo=_norm("Dashboard Analítico das Entregas")
    for i in range(Np):
        if alvo in _norm(doc[i].get_text()):
            for j in (i,i+1):
                if j<Np and doc[j].get_images(): dash_ok=True
            break
    chk(dash_ok,"Dashboard: imagem presente no corpo","imagem do painel do Power BI")
    # Timeline: nº de páginas-imagem (>= nº de meses desde jul/2025)
    meses_prev=(ano-2025)*12 + int(P["mes_numero"]) - 6
    tl=len(glob.glob(os.path.join(IMG,"TIMELINE","*.png")))
    chk(tl>=meses_prev, f"Timeline: {tl} páginas (>= {meses_prev} meses)")
    # valores/contagens esperados
    for fig in getattr(D,"QA_ESPERADOS",[]):
        chk(fig in full, f"valor/contagem '{fig}' presente")
    # status da medição: aparece na IMAGEM do quadro (não no texto) -> conferir no painel-anexo
    st=P.get("status_medicao")
    if st:
        try:
            import openpyxl
            pan=cfg.painel_anexo_xlsx()
            wq=openpyxl.load_workbook(pan,data_only=True); pf=wq["PAINEL_FINANCEIRO_M8"]
            vals={_norm(pf.cell(r,14).value) for r in (54,55,56,57)}  # N54..N57 = Status por empresa
            wq.close()
            chk(_norm(st) in vals, f"status da medição '{st}' no RESUMO do painel", f"encontrados: {vals}")
        except Exception as e:
            chk(False, f"status da medição '{st}' verificável", str(e))
    # DEC-054: % Executado de CADA linha do RESUMO (empresas + TOTAL) = acumulado ÷ contrato (o TOTAL ficou fixo em
    # 39,07% até o RM14). Confere o painel-anexo: PAINEL J54..J58 = F/H e RESUMO_EMPRESAS col E = D/contrato.
    try:
        import openpyxl
        pan=cfg.painel_anexo_xlsx()
        wq=openpyxl.load_workbook(pan,data_only=True); pf=wq["PAINEL_FINANCEIRO_M8"]
        ruins=[]
        for r in (54,55,56,57,58):
            acum,contr,pct=pf.cell(r,6).value,pf.cell(r,8).value,pf.cell(r,10).value
            if not all(isinstance(x,(int,float)) for x in (acum,contr,pct)) or not contr or abs(acum/contr-pct)>5e-5:
                ruins.append(f"J{r}={pct} (esperado {acum/contr if isinstance(acum,(int,float)) and contr else '?'})")
        jt=pf["J8"].value; j58=pf["J58"].value
        if not (isinstance(jt,(int,float)) and isinstance(j58,(int,float)) and abs(jt-j58)<5e-5):
            ruins.append(f"J58={j58} ≠ cabeçalho J8={jt}")
        wq.close()
        chk(not ruins, "painel: % Executado (empresas e TOTAL) = acumulado ÷ contrato", "; ".join(ruins))
    except Exception as e:
        chk(False, "painel: % Executado verificável", str(e))
    # reuniões do mês: alguma data dd/mm do mês nas tabelas
    mm=f"/{int(P['mes_numero']):02d}/"
    datas_mes=len(re.findall(r'\b\d{2}'+re.escape(mm)+str(ano)+r'\b', full)) + len(re.findall(r'\b\d{2}/'+f"{int(P['mes_numero']):02d}"+r'\b', full))
    chk(datas_mes>0, "reuniões: datas do mês presentes", f"{datas_mes} ocorrências")

    # ---------- 3) anexos: a LISTA do relatório deve refletir os arquivos REAIS das pastas ----------
    # Fonte da verdade = varredura das pastas P1.1..P1.4 (mesmo helper do script 10). Cada arquivo real
    # que os grupos subiram deve aparecer na "LISTA DE DOCUMENTOS ANEXOS" do PDF; e não pode haver zip
    # agregado nem na lista nem nas pastas (DEC-042 — nunca zipar, nunca mexer nos anexos de P1.3/P1.4).
    reais=cfg.listar_anexos_reais()
    fulln=_norm(full).replace(" ","")   # sem espaços: o nome quebra em várias linhas na coluna estreita
    if reais:
        for sub,nome in reais:
            chave=_norm(os.path.splitext(nome)[0]).replace(" ","")[:22]
            chk(chave in fulln, f"anexo listado no relatório: {nome[:46]}")
        # DEC-043: ZIP INDIVIDUAL de arquivo não-PDF enviado pelo grupo é o formato do POP (permitido). Proibido é só
        # o ZIP AGREGADO 'Evidências 1.X …' (DEC-042) — e qualquer .zip listado tem de ser um arquivo real da pasta.
        pos=_norm(full).find("lista de documentos anexos")
        lista=_norm(full)[pos:] if pos>=0 else ""
        zips_reais={_norm(os.path.splitext(n)[0]).replace(" ","")[:22] for _,n in reais if n.lower().endswith(".zip")}
        agregado=bool(re.search(r"evid[eê]ncias\s+1\.\d", lista))
        chk(pos>=0 and not agregado and (".zip" not in lista or zips_reais),
            "sem .zip agregado 'Evidências 1.X' na lista de anexos (DEC-042/043)",
            f"zips individuais do grupo listados: {len(zips_reais)}" if zips_reais else "")
    else:
        chk(False,"anexos reais encontrados nas pastas","nenhum arquivo — conferir manualmente")
    # nenhum zip 'Evidências 1.X' pode ter sobrado fisicamente nas pastas de P1.3/P1.4
    dirs_ev=([cfg.anexos_dir()] if cfg.layout_anexos()=="plano" else
             [cfg._find_sub(cfg.anexos_dir(),"p1.3"), cfg._find_sub(cfg.anexos_dir(),"p1.4")])
    zsobra=[f for d in dirs_ev if d and os.path.isdir(d) for f in os.listdir(d) if cfg._ZIP_AUTO.match(f)]
    chk(not zsobra, "sem zip 'Evidências 1.X' remanescente nas pastas", "; ".join(zsobra))
    doc.close()
    _grava(res,WK,N); _veredito(res)

def _veredito(res):
    fails=[l for l,_ in res if l.startswith("!!")]
    v="APROVADO" if not fails else "REVISÃO NECESSÁRIA"
    print(f"\nREVISÃO FINAL: {v}   ({len(fails)} pendência(s))\n")
    for l,d in res: print(l+(("  -> "+d) if d else ""))
    cfg.log(f"REVISÃO FINAL {v}: {len(fails)} pendência(s)")
    if fails: cfg.log("PENDÊNCIAS REVISÃO: "+"; ".join(fails)); sys.exit(1)

def _grava(res,WK,N):
    p=os.path.join(WK,"05_LOGS",f"revisao_final_RM{N}.txt")
    open(p,"w",encoding="utf-8").write("\n".join(l+((" -> "+d) if d else "") for l,d in res))

if __name__=="__main__":
    main()

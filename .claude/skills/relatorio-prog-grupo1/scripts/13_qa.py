# -*- coding: utf-8 -*-
"""
13_qa.py — QA automático do PDF final: sem placeholders/artefatos, consistência do nº do relatório
(N presente, N-1 ausente), período/mês, cidade (Cidade) + data do documento, nº da medição, e as
strings-chave esperadas do mês (cfg.dados().QA_ESPERADOS — valores financeiros e contagens). Só lê o
PDF e grava um relatório de QA local. Parametrizado (nada fixo da competência).
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg
import fitz

def data_extenso(ddmmaaaa):
    d,m,a=ddmmaaaa.split("/")
    return f"{int(d):02d} de {cfg.MESES_PT[int(m)-1].lower()} de {a}"

def main():
    P=cfg.params(); N=cfg.numero(); mes=cfg.mes_min(); ano=cfg.ano()
    D=cfg.dados(); MED_NUM=getattr(D,"MEDICAO_NUM","")
    WK=cfg.work_dir()
    PDF=os.path.join(WK,"07_SAIDA",cfg.doc_final_basename()+".pdf")
    if not os.path.exists(PDF):
        cfg.log(f"ERRO: PDF final não encontrado: {PDF}"); sys.exit(1)
    doc=fitz.open(PDF); Np=doc.page_count
    full="".join(doc[i].get_text() for i in range(Np))
    lowfull=full.lower()
    def pages_with(kw): return [i+1 for i in range(Np) if kw.lower() in doc[i].get_text().lower()]
    res=[]
    def chk(label, cond, detail=""): res.append((("OK " if cond else "!! ")+label, detail))

    import re
    def pages_with_cs_word(kw):  # case-sensitive, palavra inteira (placeholders tipo 'TODO')
        rx=re.compile(rf'\b{re.escape(kw)}\b')
        return [i+1 for i in range(Np) if rx.search(doc[i].get_text())]
    for bad in ["PREENCHER","XXX","<<",">>","[INSERIR]","/yyyy","São Paulo",
                "Não houve demandas","Lista de Demandas","imagem a ser inserida"]:
        hits=pages_with(bad); chk(f"sem '{bad}'", not hits, f"páginas {hits}" if hits else "")
    # 'TODO' (placeholder): case-sensitive + palavra inteira, para NÃO casar "todos"/"toda"/"metodologia"
    hits=pages_with_cs_word("TODO"); chk("sem 'TODO' (placeholder)", not hits, f"páginas {hits}" if hits else "")
    chk(f"contém 'N° {N}'/'Nº {N}'", (f"N° {N}" in full or f"Nº {N}" in full or f"N°{N}" in full))
    chk(f"não há nº de relatório antigo (N° {N-1})",
        all(f"{t} {s}{e}{N-1}" not in full for t in ("Relatório Mensal","Relatório de Produtos Mensal")
            for s in ("N°","Nº") for e in (" ","")), str(pages_with(f"relatório mensal n° {N-1}")))
    if MED_NUM: chk(f"contém '{MED_NUM}ª medição'", f"{MED_NUM}ª medição".lower() in lowfull)
    for fig in getattr(D,"QA_ESPERADOS",[]):
        chk(f"valor/contagem '{fig}' presente", fig in full)
    per_ini=P["periodo_inicio"][:2]; per_fim=P["periodo_fim"][:2]
    chk(f"período ({per_ini}..{per_fim} {mes}) presente",
        (per_ini in full and per_fim in full and mes in lowfull))
    doc_date=cfg.data_documento()   # = último dia da competência (DEC-058)
    chk(f"encerramento com 'Cidade, {data_extenso(doc_date)}'", f"Cidade, {data_extenso(doc_date)}" in full)
    # Capa (DEC-058): data da capa embutida no DOCX final = último dia da competência
    try:
        import capa
        docx_final=os.path.join(WK,"07_SAIDA",cfg.doc_final_basename()+".docx")
        lida=capa.ler_data_capa(docx_final) if os.path.exists(docx_final) else None
        chk(f"capa com a data '{data_extenso(doc_date)}' (fim da competência)",
            lida is not None and lida.lower()==data_extenso(doc_date), f"capa traz: {lida!r}")
    except Exception as e:
        chk("capa: data verificável", False, f"erro ao ler a capa: {e}")

    # ===== Layout: cabeçalho/rodapé e quebras de página dos títulos Produto 1.X (DEC-032) =====
    # (1) TODAS as páginas devem ter cabeçalho ("CONSÓRCIO") e rodapé ("CONFIDENCIAL").
    sem_hf=[i+1 for i in range(Np)
            if ("CONSÓRCIO" not in doc[i].get_text()) or ("CONFIDENCIAL" not in doc[i].get_text())]
    chk("cabeçalho+rodapé em TODAS as páginas", not sem_hf, f"faltando em {sem_hf}" if sem_hf else "")
    # (2) Cada título "Produto 1.X – Relatório …" deve iniciar no TOPO de uma página nova.
    for x in ["1.1","1.2","1.3","1.4"]:
        loc=None
        for i in range(Np):
            rs=doc[i].search_for(f"Produto {x} – Relat")
            if rs: loc=(i, min(r.y0 for r in rs))   # última ocorrência = título no corpo (não o sumário)
        if loc:
            pno,y0=loc; phh=doc[pno].rect.height
            chk(f"'Produto {x}' inicia em página nova (topo)", y0<=0.30*phh, f"pág {pno+1}, y={y0:.0f}/{phh:.0f}pt")
        else:
            chk(f"'Produto {x}' localizado no corpo", False, "título não encontrado no PDF")
    # (3) P1.2: texto consolidado "Relatório de Relacionamento Institucional" presente (DEC-033)
    chk("P1.2: texto consolidado 'Relatório de Relacionamento Institucional' presente",
        "relatório de relacionamento institucional do mês de" in lowfull)

    # Versão EM ELABORAÇÃO (DEC-064): nenhum quadro "SEÇÃO EM ELABORAÇÃO" pode sobrar na versão final
    hits=pages_with("SEÇÃO EM ELABORAÇÃO")
    chk("sem quadro 'SEÇÃO EM ELABORAÇÃO' (versão final)", not hits, f"páginas {hits}" if hits else "")
    # Avisos persistentes de QA gerados por etapas anteriores (ex.: fallback da Timeline — DEC-020).
    # Cada aviso conta como pendência: impede a APROVAÇÃO automática até a conferência visual.
    for _k,_msg in cfg.qa_warnings().items():
        res.append(("!! "+_msg, "conferência manual necessária (não aprovar sem revisar)"))

    fails=[l for l,_ in res if l.startswith("!!")]
    veredito = "APROVADO" if not fails else "REVISÃO NECESSÁRIA"
    print(f"PÁGINAS: {Np}   VEREDITO QA: {veredito}   ({len(fails)} pendência(s))\n")
    for label,detail in res: print(label + (("  -> "+detail) if detail else ""))
    doc.close()
    qa_path=os.path.join(WK,"05_LOGS",f"qa_RM{N}.txt")
    open(qa_path,"w",encoding="utf-8").write(
        f"QA RM{N} - {Np} páginas - VEREDITO: {veredito}\n\n"
        +"\n".join(l+((" -> "+d) if d else "") for l,d in res))
    cfg.log(f"QA {veredito}: {Np} páginas, {len(fails)} pendência(s). Relatório: {qa_path}")
    if fails: cfg.log("PENDÊNCIAS: "+"; ".join(fails))

if __name__=="__main__":
    main()

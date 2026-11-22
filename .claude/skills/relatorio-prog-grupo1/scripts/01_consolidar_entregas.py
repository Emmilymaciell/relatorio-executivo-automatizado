# -*- coding: utf-8 -*-
"""
01_consolidar_entregas.py — Consolida as entregas da competência (Versão = N) a partir das
fontes REAIS do SharePoint (leitura), gerando output/RM.../02_TRABALHO/entregas_rmN.json + stats.
Parametrizado por config/parametros.json (prog_config). SÓ LEITURA de fontes; escreve apenas
arquivos locais de trabalho. Preserva: regra "Sem data" (DEC), dedup por chave composta,
recorrentes P1.1/P1.2 clonados da versão N-1. Idempotente (regrava os JSON de trabalho).
"""
import os, sys, json, datetime, openpyxl
from collections import Counter
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import prog_config as cfg

def norm(s): return "" if s is None else str(s).strip()

def to_date(v):
    """datetime.date ou None (=> 'Sem data'). DEC: qualquer não-data vira 'Sem data'."""
    if isinstance(v,(datetime.datetime,datetime.date)):
        return v.date() if isinstance(v,datetime.datetime) else v
    s=norm(v)
    if not s: return None
    for fmt in ("%d/%m/%Y","%d/%m/%y","%Y-%m-%d","%d.%m.%Y","%Y-%m-%d %H:%M:%S"):
        try: return datetime.datetime.strptime(s,fmt).date()
        except: pass
    return None

def colmap(hdr, wanted):
    m={}
    for key,names in wanted.items():
        idx=None
        for nm in names:
            for i,h in enumerate(hdr):
                if h.strip().lower()==nm.lower(): idx=i;break
            if idx is not None: break
        if idx is None:
            for nm in names:
                for i,h in enumerate(hdr):
                    if nm.lower() in h.strip().lower(): idx=i;break
                if idx is not None: break
        m[key]=idx
    return m

WANT={"produto":["Produto"],"empresa":["Empresa Responsável","Empresa"],"doc":["Documento"],
      "tipo":["Tipo","Parecer","Manifestação","Manifestacao"],"desc":["Descrição"],"emp":["Empreendimento"],
      "data":["Data da entrega","Data Entrega"],"status":["Status Entrega","Status"],
      "nrel":["N° Relat","Nº Relat","Relatório"]}

def data_sheet(wb):
    """Aba cujo cabeçalho (linha 3) contém 'Produto' e 'Documento'."""
    for ws in wb.worksheets:
        try:
            hdr=[norm(c.value) for c in ws[3]]
        except Exception:
            continue
        low=[h.lower() for h in hdr]
        if any("produto"==h or "produto" in h for h in low) and any("documento" in h for h in low):
            return ws
    return wb.worksheets[0]

# DEC-061 (corte temporal) REVOGADA pela usuária em 16/11/2026 (DEC-067): o Controle de Demandas reproduz o que os grupos
# DECLARARAM nas Listas de Demandas da competência (G2: Lista Mestre filtrada pelo N° Relatório). A data de upload/entrega
# NÃO muda a competência — atraso é assunto da auditoria do POP.


def main():
    P=cfg.params(); N=cfg.numero(); WK=cfg.work_dir(); TR=os.path.join(WK,"02_TRABALHO")
    cfg.log(f"Consolidação RM{N} — {P['mes']}/{P['ano']} (competência = Versão {N})")
    rows=[]
    def add(src,produto,empresa,doc,tipo,desc,emp,dataval,status):
        d=to_date(dataval)
        rows.append({"_src":src,"produto":norm(produto),"empresa":norm(empresa),"doc":norm(doc),
            "tipo":norm(tipo),"desc":norm(desc),"emp":norm(emp) or "Geral",
            "data": d.isoformat() if d else "Sem data","status":norm(status) or "Entregue"})

    # ---- listas mensais dos grupos (todas as linhas populadas) ----
    for group in ["G1","G3","G4","G5"]:
        path=cfg.lista_demandas_path(group)
        if not path:
            cfg.log(f"AVISO: lista de demandas de {group} não encontrada — pulando."); continue
        cfg.log(f"{group}: {os.path.basename(path)}")
        wb=openpyxl.load_workbook(path,data_only=True); ws=data_sheet(wb)
        hdr=[norm(c.value) for c in ws[3]]; cm=colmap(hdr,WANT)
        for r in ws.iter_rows(min_row=4,values_only=True):
            prod=norm(r[cm["produto"]]) if cm["produto"] is not None else ""
            doc=norm(r[cm["doc"]]) if cm["doc"] is not None else ""
            if not prod and not doc: continue
            add(group,prod,r[cm["empresa"]],doc,r[cm["tipo"]],r[cm["desc"]],r[cm["emp"]],
                r[cm["data"]] if cm["data"] is not None else None, r[cm["status"]])
        wb.close()

    # ---- G2 Lista Mestre: N° Relatório == N ----
    lm=cfg.lista_mestre_g2_path()
    wb=openpyxl.load_workbook(lm,data_only=True); ws=data_sheet(wb)
    hdr=[norm(c.value) for c in ws[3]]; cm=colmap(hdr,WANT)
    for r in ws.iter_rows(min_row=4,values_only=True):
        if all(c in (None,"") for c in r): continue
        nrel=norm(r[cm["nrel"]]) if cm["nrel"] is not None else ""
        if nrel.replace(".0","")!=str(N): continue
        add("G2",r[cm["produto"]],r[cm["empresa"]],r[cm["doc"]],r[cm["tipo"]],r[cm["desc"]],r[cm["emp"]],
            r[cm["data"]] if cm["data"] is not None else None,r[cm["status"]])
    wb.close()
    cfg.log(f"G2 (Lista Mestre, N° Relatório={N}): {sum(1 for x in rows if x['_src']=='G2')} linhas")

    # ---- recorrentes P1.1/P1.2 clonados da versão N-1 do acumulado ----
    wb=openpyxl.load_workbook(cfg.controle_demandas_path(),data_only=True)
    ws=wb["Controle Geral"]; prev=str(N-1); vprev={}; desc_hist={}  # desc_hist: chave -> (versão, última desc não vazia)
    for r in ws.iter_rows(min_row=4,values_only=True):
        prod=norm(r[2]); ver=norm(r[10]).replace(".0","")
        if (prod.startswith("1.1") or prod.startswith("1.2")) and norm(r[6]):
            k=" ".join(norm(r[4]).split())[:22]
            try: vn=int(ver)
            except ValueError: vn=-1
            if vn>=desc_hist.get(k,(-2,""))[0]: desc_hist[k]=(vn,norm(r[6]))
        if ver==prev and (prod.startswith("1.1") or prod.startswith("1.2")):
            vprev[norm(r[4])[:25]]={"produto":prod,"empresa":norm(r[3]),"tipo":norm(r[5]),
                                    "desc":norm(r[6]),"emp":norm(r[7])}
    wb.close()
    mes=cfg.mes_min(); anoP=P["ano"]; ini,fim=cfg.periodo_ddmm()
    recur=[
     (f"P1.1 - Relatório Grupo 1 - {cfg.termo_documento()} N°{N} - {ini} a {fim}","P1.1 - Relatório Grupo 1"),
     (f"P1.1 - Timeline de Acompanhamento Macro até {mes} {anoP}","P1.1 - Timeline de Acompanhamento Macr"),
     (f"P1.1 - Dashboard da Análise de Entrega de Produtos - {mes}","P1.1 - Dashboard da Análise de Entrega"),
     (f"P1.1 - Controle de Entrega de Produtos de {mes}","P1.1 - Controle de Entrega de Produtos"),
     (f"P1.1 - Quadro Financeiro do mês de {mes}","P1.1 - Quadro Financeiro do mês de"),
     (f"P1.2 - Controle de Reuniões do mês de {mes}","P1.2 -  Controle de Reuniões do mês de"),
    ]
    DATA_REC=cfg.data_recorrente()
    for newdoc,key in recur:
        base=None
        for k,v in vprev.items():
            if k.startswith(key[:22]): base=v;break
        if base is None:
            base={"produto":"1.1 COORDENAÇÃO GERAL" if newdoc.startswith("P1.1") else "1.2 RELAÇÕES INSTITUCIONAIS",
                  "empresa":"Dynatest","tipo":"Outros","desc":"","emp":"Geral"}
        # DEC-056: descrição vazia na versão anterior (RM14: P1.2 sem descrição) -> última não vazia do histórico
        if not base["desc"]:
            base=dict(base, desc=desc_hist.get(" ".join(key.split())[:22],(0,""))[1])
        # DEC-056: a descrição clonada trazia o mês de origem ("…de Junho" desde o RM12) — trocar pelo mês corrente
        import re as _re2
        _MESES="janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro"
        base=dict(base, desc=_re2.sub(rf"\b({_MESES})\b",
                                      lambda m: P["mes"] if m.group(0)[0].isupper() else P["mes"].lower(),
                                      base["desc"], flags=_re2.I))
        add("REC",base["produto"],"Dynatest",newdoc,
            # TIPO permanece 'Relatório Mensal': é categoria usada pelo Power BI/histórico (DEC-044)
            "Relatório Mensal" if "Relatório Grupo 1" in newdoc else "Outros",
            base["desc"],base["emp"] or "Geral",DATA_REC,"Entregue")

    # ---- normalização do PRODUTO (DEC-055) ----
    # O Power BI mapeia produto -> frente pelo NOME (aba Listas do acumulado); nome fora do padrão vira linha em
    # branco no dashboard (RM15: CG enviou "1.3 RELATÓRIO DE COMUNICAÇÃO INTEGRADA" em vez de "1.3 COMUNICAÇÃO SOCIAL").
    # Regra: o nome canônico de cada código (ex.: "1.3") é o da versão MAIS RECENTE do acumulado (< N) que o usou.
    import re as _re
    canon={}  # código -> (versão, nome)
    wb=openpyxl.load_workbook(cfg.controle_demandas_path(),read_only=True,data_only=True)
    for r in wb["Controle Geral"].iter_rows(min_row=4,values_only=True):
        prod=norm(r[2]); m=_re.match(r"^(\d\.\d)\s",prod)
        try: ver=int(float(r[10]))
        except (TypeError,ValueError): continue
        if m and ver<N and ver>=canon.get(m.group(1),(-1,""))[0]: canon[m.group(1)]=(ver,prod)
    wb.close()
    # produto vazio: herda o produto do MESMO documento em outra linha do mesmo grupo (ex.: RO-549.1 por rodovia)
    por_doc={(r["_src"],r["doc"]):r["produto"] for r in rows if r["produto"] and r["doc"]}
    ajustes=Counter()
    for r in rows:
        if not r["produto"]:
            herd=por_doc.get((r["_src"],r["doc"]))
            if herd:
                ajustes[(f"(vazio) {r['_src']} {r['doc']}",herd)]+=1; r["produto"]=herd
            else:
                cfg.log(f"AVISO: produto VAZIO sem referência — {r['_src']} | {r['doc']} | {r['desc'][:60]} (corrigir na fonte)")
            continue
        m=_re.match(r"^(\d\.\d)\b",r["produto"])
        if not m or m.group(1) not in canon:
            cfg.log(f"AVISO: produto sem código conhecido no acumulado — {r['_src']} | '{r['produto']}' (conferir no Power BI)")
            continue
        alvo=canon[m.group(1)][1]
        if r["produto"]!=alvo:
            ajustes[(r["produto"],alvo)]+=1; r["produto"]=alvo
    for (de,para),n in ajustes.items():
        cfg.log(f"AJUSTE produto ({n} linha(s)): '{de}' -> '{para}'  [DEC-055]")

    # ---- TODAS as linhas declaradas entram (DEC-067). Linha repetida na fonte NÃO é descartada: só é registrada para
    # conferência (o grupo declarou; a correção é dele). Antes (até 16/11/2026) a dedup por chave composta descartava.
    from collections import defaultdict as _dd
    vistos=_dd(list)
    for i,r in enumerate(rows):
        vistos[(r["_src"],r["produto"],r["empresa"],r["doc"],r["tipo"],r["desc"],r["emp"])].append(i)
    duplicadas=[{"fonte":k[0],"doc":k[3],"desc":k[5],"emp":k[6],"ocorrencias":len(v)} for k,v in vistos.items() if len(v)>1]
    for d in duplicadas:
        cfg.log(f"AVISO: linha repetida na fonte ({d['ocorrencias']}x) — {d['fonte']} | {d['doc']} | {d['emp']} | {d['desc'][:60]} (mantida; conferir com o grupo)")
    final=rows
    declarados=dict(Counter(r["_src"] for r in rows))
    cfg.log(f"DECLARADOS por fonte = CONSIDERADOS: {declarados}")

    out_json=os.path.join(TR,f"entregas_rm{N}.json")
    json.dump(final,open(out_json,"w",encoding="utf-8"),ensure_ascii=False,indent=2)
    stats={"total":len(final),
           "by_source":dict(Counter(r["_src"] for r in final)),
           "by_empresa":dict(Counter(r["empresa"] for r in final)),
           "by_status":dict(Counter(r["status"] for r in final)),
           "by_produto":dict(Counter(r["produto"] for r in final)),
           "by_empreendimento":dict(Counter(r["emp"] for r in final)),
           "by_tipo":dict(Counter(r["tipo"] for r in final)),
           "sem_data":sum(1 for r in final if r["data"]=="Sem data"),
           "declarados_por_fonte":declarados, "considerados_por_fonte":dict(Counter(r["_src"] for r in final)),
           "duplicadas_na_fonte":duplicadas}
    json.dump(stats,open(os.path.join(TR,"entregas_stats.json"),"w",encoding="utf-8"),ensure_ascii=False,indent=2)
    cfg.log(f"TOTAL competência: {len(final)} | por empresa: {stats['by_empresa']} | por status: {stats['by_status']}")
    cfg.log(f"salvo: {out_json} + entregas_stats.json")

if __name__=="__main__":
    main()

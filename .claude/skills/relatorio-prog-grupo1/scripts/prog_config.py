# -*- coding: utf-8 -*-
"""
prog_config.py — módulo comum de configuração e caminhos para a automação do
Relatório Mensal Grupo 1 (Programa). Todos os scripts importam este módulo para
NÃO terem mês/ano/número fixos no código.

Uso:
    import prog_config as cfg
    P = cfg.params()          # dict de config/parametros.json
    C = cfg.caminhos()        # dict de config/caminhos.json
    root = cfg.prog_root()    # raiz LOCAL sincronizada do SharePoint Programa
    mf = cfg.month_folder()   # ex.: "14. Agosto 2026"
    acc = cfg.controle_demandas_path()
    ...
Variáveis de ambiente (opcionais) para sobrepor:
    PROG_PROJECT_DIR, PROG_ROOT, PROG_PARAMS
"""
import os, json, glob, datetime, unicodedata

MESES_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho",
            "Agosto","Setembro","Outubro","Novembro","Dezembro"]

def project_dir():
    """Raiz do projeto de automação (esta pasta permanente)."""
    env=os.environ.get("PROG_PROJECT_DIR")
    if env and os.path.isdir(env): return env
    # sobe a partir de scripts/ -> skill/ -> skills/ -> .claude/ -> projeto
    here=os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(here, "..","..","..",".."))

def _cfg_dir(): return os.path.join(project_dir(),"config")

def params():
    p=os.environ.get("PROG_PARAMS") or os.path.join(_cfg_dir(),"parametros.json")
    if not os.path.exists(p):
        p=os.path.join(_cfg_dir(),"parametros-exemplo.json")
    with open(p,encoding="utf-8") as f: return json.load(f)

def caminhos():
    with open(os.path.join(_cfg_dir(),"caminhos.json"),encoding="utf-8") as f:
        return json.load(f)

def _norm(s):
    return "".join(c for c in unicodedata.normalize("NFD",s) if unicodedata.category(c)!="Mn").lower()

def prog_root():
    """Localiza a raiz LOCAL sincronizada da biblioteca Programa.
    1) env PROG_ROOT; 2) caminhos.json['prog_root']; 3) busca por uma pasta que
    contenha subpastas '...3. Desenvolvimento' e '...4. Produtos'."""
    env=os.environ.get("PROG_ROOT")
    if env and os.path.isdir(env): return env
    try:
        c=caminhos().get("prog_root")
        if c and os.path.isdir(c): return c
    except Exception: pass
    home=os.path.expanduser("~")
    for base in [home]+[os.path.join(home,d) for d in os.listdir(home) if os.path.isdir(os.path.join(home,d))]:
        try: subs=os.listdir(base)
        except Exception: continue
        has_dev=any("3. desenvolvimento" in _norm(s) for s in subs)
        has_prod=any("4. produtos" in _norm(s) for s in subs)
        if has_dev and has_prod: return base
    raise RuntimeError("PROG_ROOT não localizado. Defina em config/caminhos.json ou variável PROG_ROOT.")

def _find_sub(parent, needle):
    """Encontra subpasta cujo nome (sem acento/caixa) contém 'needle'.
    Retorna None se 'parent' ainda não existir (ex.: pasta do mês não criada)."""
    if not parent or not os.path.isdir(parent): return None
    n=_norm(needle)
    for s in os.listdir(parent):
        if n in _norm(s) and os.path.isdir(os.path.join(parent,s)):
            return os.path.join(parent,s)
    return None

def dir_desenvolvimento(): return _find_sub(prog_root(),"3. desenvolvimento")
def dir_produtos():        return _find_sub(prog_root(),"4. produtos")
def dir_g1_coord():
    return _find_sub(dir_desenvolvimento(),"g1")
def dir_desenv_produtos():
    return _find_sub(dir_g1_coord(),"desenvolvimento de produtos")

def controle_demandas_path():
    return os.path.join(dir_desenv_produtos(),"P.1.1 - Controle de demandas Programa.xlsx")
def lista_mestre_g2_path():
    g2=_find_sub(dir_desenvolvimento(),"g2")
    return os.path.join(g2,"00_ Lista mestre.xlsx")

def month_folder():
    """Nome da pasta mensal em 4. Produtos, ex.: '14. Agosto 2026'.
    Convenção descoberta: pasta = '{numero_relatorio}. {Mes} {ano}'."""
    P=params(); num=P["numero_relatorio"]; mes=P["mes"]; ano=P["ano"]
    prod=dir_produtos()
    # tenta achar pasta existente que comece com o número
    for s in os.listdir(prod):
        if s.strip().startswith(f"{num}.") and os.path.isdir(os.path.join(prod,s)):
            return os.path.join(prod,s)
    return os.path.join(prod, f"{num}. {mes} {ano}")

def g1_dir():        return os.path.join(month_folder(),"G1")
def anexos_dir():
    """Pasta de anexos do G1 (ANEXOS no padrão do POP; 'Anexos' nos meses antigos — Windows não diferencia)."""
    g1=g1_dir()
    if os.path.isdir(g1):
        for s in os.listdir(g1):
            if _norm(s)=="anexos" and os.path.isdir(os.path.join(g1,s)): return os.path.join(g1,s)
    return os.path.join(g1,"ANEXOS")

def layout_anexos():
    """'por_produto' (legado: ANEXOS/P1.1 - …/arquivo) ou 'plano' (POP de Entrega v2026-10-01: ANEXOS única,
    produto no NOME do arquivo — DEC-043). Override: parametros.json 'layout_anexos'. Automático: se a pasta
    ANEXOS do mês tem subpastas 'P1.x' => por_produto; senão => plano (padrão a partir de Outubro/2026)."""
    v=(params().get("layout_anexos") or "").strip().lower()
    if v in ("plano","por_produto"): return v
    d=anexos_dir()
    if os.path.isdir(d) and any(_norm(s).startswith("p1.") and os.path.isdir(os.path.join(d,s)) for s in os.listdir(d)):
        return "por_produto"
    return "plano"

def anexo_p11_dir():
    d=anexos_dir()
    if layout_anexos()=="plano": return d
    return _find_sub(d,"p1.1") or os.path.join(d,"P1.1 - Gestao Executiva")
def anexo_p12_dir():
    d=anexos_dir()
    if layout_anexos()=="plano": return d
    return _find_sub(d,"p1.2") or os.path.join(d,"P1.2 - Relacoes Institucionais")

def _produtos_pop():
    """Catálogo de produtos do POP (config/entrega_pop.json — fonte única dos nomes curtos)."""
    try:
        with open(os.path.join(_cfg_dir(),"entrega_pop.json"),encoding="utf-8") as f:
            return {k:v for k,v in json.load(f)["produtos"].items() if not k.startswith("_")}
    except Exception:
        return {}

def nome_anexo_g1(cod, descricao, ext, rodovia="GERAL"):
    """Nome dos anexos que a automação GERA para P1.1/P1.2.
    layout 'plano' (POP Quadro 3): '[Código] - [Produto] - [Nome do arquivo] - [Rodovia]' (ex.:
    'P1.1 - Gestão Executiva - Quadro Financeiro do mês de outubro - GERAL.xlsx').
    layout 'por_produto' (legado, meses até Setembro/2026): '[Código] - [Nome do arquivo]' (nome histórico)."""
    if layout_anexos()=="plano":
        nome=_produtos_pop().get(cod,{}).get("nome",cod)
        return f"{cod} - {nome} - {descricao} - {rodovia}{ext}"
    return f"{cod} - {descricao}{ext}"

# Padrão do zip que a automação NÃO deve mais gerar (era "Evidências 1.3 08.26.zip") — DEC-042.
import re as _re
_ZIP_AUTO=_re.compile(r'^evid[eê]ncias\s+1\.\d.*\.zip$', _re.I)
def _ignora_anexo(fname):
    """True se o arquivo NÃO deve entrar na lista de anexos (temporários, lixo de SO, ou o
    zip agregado que a automação criava por engano — DEC-042)."""
    low=fname.lower()
    if fname.startswith("~$") or low.endswith(".tmp"): return True
    if low in ("thumbs.db","desktop.ini",".ds_store"): return True
    if _ZIP_AUTO.match(fname): return True
    return False
def listar_anexos_reais():
    """Varre as pastas de PRODUTO do G1 (P1.1..P1.4) sob Anexos e devolve, em ordem
    (P1.1, P1.2, P1.3, P1.4; arquivos por nome), a lista de (subpasta, nome_arquivo) dos anexos
    REAIS que os grupos subiram. NUNCA agrega/zipa; lista exatamente os arquivos existentes com o
    nome que está lá (DEC-042). Ignora temporários e o zip 'Evidências 1.X ...' legado.
    Layout 'plano' (POP, DEC-043): arquivos direto em ANEXOS; o produto vem do prefixo do nome
    ('P1.3 - …'); devolve (código, nome) — arquivo sem código entra no fim (é pendência de nome)."""
    base=anexos_dir(); out=[]
    if not base or not os.path.isdir(base): return out
    if layout_anexos()=="plano":
        def _cod(f):
            m=_re.match(r"^\s*(P1\.\d)\b",f)
            return m.group(1) if m else "P9.9"
        fs=[f for f in os.listdir(base) if os.path.isfile(os.path.join(base,f)) and not _ignora_anexo(f)]
        for f in sorted(fs, key=lambda x:(_cod(x),_norm(x))):
            out.append((_cod(f) if _cod(f)!="P9.9" else "?", f))
        return out
    subs=[s for s in os.listdir(base) if os.path.isdir(os.path.join(base,s))]
    subs.sort(key=lambda s:_norm(s))   # P1.1..P1.4
    for s in subs:
        d=os.path.join(base,s)
        for f in sorted(os.listdir(d), key=lambda x:_norm(x)):
            if os.path.isfile(os.path.join(d,f)) and not _ignora_anexo(f):
                out.append((s,f))
    return out

def painel_anexo_xlsx():
    """Painel 'Quadro Financeiro' do mês na pasta de anexos: o .xlsx, ou — no padrão POP (DEC-066) — o .xlsx de dentro
    do ZIP individual (extraído para 02_TRABALHO). None se não houver."""
    d=anexo_p11_dir()
    x=find_in_dir(d,"p1.1","quadro financeiro",ext=".xlsx")
    if x: return x
    z=find_in_dir(d,"p1.1","quadro financeiro",ext=".zip")
    if not z: return None
    import zipfile
    dst=os.path.join(work_dir(),"02_TRABALHO")
    with zipfile.ZipFile(z) as zf:
        xs=[n for n in zf.namelist() if n.lower().endswith(".xlsx")]
        if not xs: return None
        zf.extract(xs[0],dst)
    return os.path.join(dst,xs[0])

def timeline_pptx_src():
    """PPTX-fonte da Timeline = editável ACUMULADO compartilhado (uma slide por mês; atualizado
    mês a mês; é a REFERÊNCIA oficial). Ordem de busca (DEC-025):
      1) 'Desenvolvimento de Produtos' (compartilhado, ex.: 'Compartilhado Timeline Relatório Programa.pptx');
      2) anexo P1.1 do mês; 3) rascunho local em work/01_FONTES.
    Override opcional: env PROG_TIMELINE_PPTX (caminho .pptx explícito).
    Retorna None se nenhum existir (07 então registra pendência de QA — DEC-022)."""
    env=os.environ.get("PROG_TIMELINE_PPTX")
    if env and os.path.exists(env): return env
    return (find_in_dir(dir_desenv_produtos(), "timeline", ext=".pptx")
            or find_in_dir(anexo_p11_dir(), "timeline", ext=".pptx")
            or find_in_dir(os.path.join(work_dir(),"01_FONTES"), "timeline", ext=".pptx"))

# --- diretórios de trabalho do projeto (locais, não SharePoint) ---
def work_dir():
    """Diretório de TRABALHO/EXECUÇÃO (subpastas 01_FONTES..07_SAIDA + 03_IMAGENS/*).
    Ordem de resolução (DEC-023 — separa projeto permanente de execução):
      1) env PROG_WORK_DIR  -> raiz da execução atual (redireciona TODA a saída/temporários
         sem alterar PROG_PROJECT_DIR nem quebrar caminhos permanentes da Skill/config);
      2) padrão -> project_dir()/output/RM{n}_{ano}_{mm}.
    project_dir() continua sendo a RAIZ PERMANENTE (config/, etc.)."""
    env=os.environ.get("PROG_WORK_DIR")
    if env:
        d=env
    else:
        P=params(); d=os.path.join(project_dir(),"output",f"RM{P['numero_relatorio']}_{P['ano']}_{int(P['mes_numero']):02d}")
    for sub in ["01_FONTES","02_TRABALHO","03_IMAGENS","04_PDFS_TEMP","05_LOGS","06_BACKUP","07_SAIDA"]:
        os.makedirs(os.path.join(d,sub),exist_ok=True)
    for img in ["CONTROLE_PRODUTOS","DASHBOARD","TIMELINE","REUNIOES_P11","REUNIOES_P12","FINANCEIRO"]:
        os.makedirs(os.path.join(d,"03_IMAGENS",img),exist_ok=True)
    return d

def mes_min(): return params()["mes"].lower()   # 'agosto'
def ano(): return params()["ano"]
def numero(): return int(params()["numero_relatorio"])

def dry_run():
    """True se PROG_DRY_RUN estiver setado (1/true/yes). Em DRY_RUN, os scripts NÃO
    escrevem em arquivos oficiais do SharePoint (só simulam/loga)."""
    return str(os.environ.get("PROG_DRY_RUN","")).lower() in ("1","true","yes","on")

def use_example_data():
    """True se PROG_USE_EXAMPLE_DATA estiver setado (1/true/yes/on): MODO TESTE/FIXTURE.
    Habilita o uso do exemplo assets/texts_exemplo_julho2026.py como fixture de
    desenvolvimento quando NÃO há config/dados_mes.py. NUNCA é ativado automaticamente:
    é uma decisão explícita do operador. Em execução real, deixe-o desligado (DEC-019)."""
    return str(os.environ.get("PROG_USE_EXAMPLE_DATA","")).lower() in ("1","true","yes","on")

def periodo_ddmm():
    """('01.08','31.08') a partir de periodo_inicio/fim (dd/mm/aaaa)."""
    P=params()
    ini=P["periodo_inicio"][:5].replace("/",".")
    fim=P["periodo_fim"][:5].replace("/",".")
    return ini,fim

TERMO_DOCUMENTO_PADRAO = "Relatório de Produtos Mensal"   # nomenclatura oficial (usuária, 01/10/2026 — DEC-044)

def termo_documento():
    """Designação do documento. Padrão: 'Relatório de Produtos Mensal'. Para RE-GERAR competências antigas
    com o nome histórico, usar parametros.json "termo_documento": "Relatório Mensal"."""
    return (params().get("termo_documento") or TERMO_DOCUMENTO_PADRAO).strip()

STATUS_EM_ELABORACAO = "EM ELABORAÇÃO"

def status_documento():
    """'EM ELABORAÇÃO' enquanto config/dados_mes.py tiver PENDENCIAS (seções aguardando insumo) — DEC-064; '' quando
    não houver. Override: parametros.json "status_documento" ("" força o nome final)."""
    P=params()
    if "status_documento" in P: return (P.get("status_documento") or "").strip()
    try:
        return STATUS_EM_ELABORACAO if (getattr(dados(),"PENDENCIAS",None) or {}) else ""
    except Exception:
        return ""

def doc_final_basename():
    """Nome-base do relatório, ex.: 'Programa - Relatório Grupo 1 - Relatório de Produtos Mensal N°15 - 01.09 a 30.09'.
    Versão de desenvolvimento com pendências: '… - 01.10 a 31.10 - EM ELABORAÇÃO' (DEC-064; o sufixo sai sozinho
    quando PENDENCIAS fica vazio)."""
    ini,fim=periodo_ddmm()
    base=f"Programa - Relatório Grupo 1 - {termo_documento()} N°{numero()} - {ini} a {fim}"
    st=status_documento()
    return f"{base} - {st}" if st else base

def fim_competencia(ano_, mes_numero):
    """Último dia CALENDÁRIO do mês da competência, 'dd/mm/aaaa' (fevereiro bissexto incluído)."""
    import calendar
    a=int(ano_); m=int(mes_numero)
    return f"{calendar.monthrange(a,m)[1]:02d}/{m:02d}/{a}"

def data_capa():
    """DATA DA CAPA / DO DOCUMENTO = ÚLTIMO DIA DO MÊS DA COMPETÊNCIA (DEC-058, regra permanente da usuária,
    16/11/2026). Nunca a data real de elaboração (que cai no mês seguinte). Vale para a capa, o encerramento
    ('Cidade, …') e a 'Data Entrega' das linhas recorrentes P1.1/P1.2. O campo legado 'data_documento' do
    parametros.json NÃO a substitui: se divergir, é ignorado (com aviso) — supera a DEC-003."""
    P=params()
    return fim_competencia(P["ano"], P["mes_numero"])

def data_documento():
    """Alias de data_capa() (DEC-058). Avisa se o parametros.json ainda trouxer outra 'data_documento'."""
    d=data_capa(); leg=params().get("data_documento")
    if leg and leg!=d:
        print(f"AVISO: parametros.json data_documento={leg} ignorada; DEC-058 usa o fim da competência ({d}).")
    return d

def data_extenso(ddmmaaaa):
    """'31/10/2026' -> '31 de outubro de 2026'."""
    d,m,a=ddmmaaaa.split("/")
    return f"{int(d):02d} de {MESES_PT[int(m)-1].lower()} de {a}"

def competencia_medicao():
    """(Mês, ano) de EXECUÇÃO da medição analisada no relatório = mês ANTERIOR à competência do relatório.
    DEC-063 (determinação (I) do Parceiro, Relatório de Gestão do Contrato de set/2026): a medição é denominada pela
    competência de EXECUÇÃO, não pelo mês de processamento. Ex.: RM16 (outubro/2026) -> 15ª medição = setembro/2026."""
    P=params(); m=int(P["mes_numero"]); a=int(P["ano"])
    m,a=(12,a-1) if m==1 else (m-1,a)
    return MESES_PT[m-1], a

def data_recorrente():
    """Data (dd/mm/aaaa) da 'Data Entrega' das linhas recorrentes P1.1/P1.2 = data da capa (DEC-058)."""
    return data_capa()

def find_in_dir(folder, *needles, ext=None, exclude=()):
    """Primeiro arquivo em 'folder' cujo nome (sem acento/caixa) contenha TODOS os needles.
    ext opcional (ex.: '.xlsx'). Ignora temporários ~$ e nomes em 'exclude'."""
    if not folder or not os.path.isdir(folder): return None
    ns=[_norm(n) for n in needles]
    for f in sorted(os.listdir(folder)):
        if f.startswith("~$"): continue
        if ext and not f.lower().endswith(ext): continue
        nf=_norm(f)
        if any(_norm(x) in nf for x in exclude): continue
        if all(n in nf for n in ns): return os.path.join(folder,f)
    return None

def lista_demandas_path(group):
    """Caminho do 'PROG_Lista de demandas' do grupo. group in {'G1','G3','G4','G5'}.
    G1 fica em G1/ (pasta mensal); os demais em G{n}/. Nome varia por mês -> busca por padrão.
    Depois da PUBLICAÇÃO do pacote (POP Etapa 6) a Lista sai da pasta do mês e fica no ARQUIVO INTERNO
    (…Desenvolvimento de Produtos/Controle Interno Entregas/{pasta do mês}/{data}/G{n}) — usa a mais recente."""
    folder=g1_dir() if group=="G1" else os.path.join(month_folder(), group)
    achado=find_in_dir(folder, "lista de demandas", ext=".xlsx")
    if achado: return achado
    try:
        arq=os.path.join(dir_desenv_produtos(),"Controle Interno Entregas",os.path.basename(month_folder()))
        for exec_ in sorted(os.listdir(arq), reverse=True) if os.path.isdir(arq) else []:
            achado=find_in_dir(os.path.join(arq,exec_,group), "lista de demandas", ext=".xlsx")
            if achado: return achado
    except Exception:
        pass
    return None

def base_am_docx():
    """Relatório da CG (base do Word) na pasta G1 do mês (…RevCG.docx).
    Override opcional: env PROG_BASE_DOCX (caminho .docx explícito) — útil quando a base do
    mês está numa subpasta (ex.: G1/Testes) ou em testes de regressão."""
    env=os.environ.get("PROG_BASE_DOCX")
    if env and os.path.exists(env): return env
    return find_in_dir(g1_dir(), "relatório grupo 1", "revcg", ext=".docx") \
        or find_in_dir(g1_dir(), "relatorio grupo 1", "rev", ext=".docx")

# Mensagem de parada do MODO REAL quando não há dados da competência (DEC-019).
DADOS_NAO_CONFIGURADOS_MSG = (
    "DADOS MENSAIS NÃO CONFIGURADOS\n\n"
    "config/dados_mes.py não foi encontrado para a competência atual.\n"
    "A execução real foi interrompida para impedir reutilização de dados\n"
    "factuais de outra competência.\n\n"
    "Como resolver:\n"
    "  - MODO REAL: crie config/dados_mes.py com as informações desta competência\n"
    "    que NÃO são deriváveis com segurança das fontes (ex.: STATUS_MEDICAO quando\n"
    "    necessário, MARCOS_EXTRAS, OBSERVACOES_EXTRAORDINARIAS, AJUSTES_MANUAIS_\n"
    "    AUTORIZADOS, parâmetros de QA). Ver scripts/README.md e DEC-019.\n"
    "  - MODO TESTE/FIXTURE (só desenvolvimento): defina PROG_USE_EXAMPLE_DATA=1 para\n"
    "    usar o exemplo assets/texts_exemplo_julho2026.py. JAMAIS numa execução real:\n"
    "    os dados são de Julho/2026 (RM13) e contaminariam o relatório do mês."
)

def _load_pyfile(path):
    import importlib.util
    spec=importlib.util.spec_from_file_location("prog_dados_mes",path)
    m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m); return m

def example_data_path():
    """Caminho do EXEMPLO/fixture de Julho/2026 (assets/texts_exemplo_julho2026.py)."""
    here=os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(here,"..","assets","texts_exemplo_julho2026.py"))

def dados():
    """Módulo de CONTEÚDO do mês. Ver DEC-019 (Proibição de fallback factual em execução real).

    dados_mes.py NÃO é um relatório manual: guarda apenas o que NÃO se deriva com segurança
    das fontes (STATUS_MEDICAO quando necessário, MARCOS_EXTRAS, OBSERVACOES_EXTRAORDINARIAS,
    INSTRUCOES_ESPECIFICAS_MES, AJUSTES_MANUAIS_AUTORIZADOS, parâmetros de QA). Textos analíticos
    de P1.1/P1.2 e MARCOS deriváveis vêm das fontes reais (Controle de Demandas, Dashboard,
    Reuniões, Quadro Financeiro), não daqui.

    Resolução:
      MODO REAL  -> (1) env PROG_DADOS (caminho .py explícito) ; ou (2) config/dados_mes.py.
                    Se nenhum existir, PARA com 'DADOS MENSAIS NÃO CONFIGURADOS'.
      MODO TESTE -> só com PROG_USE_EXAMPLE_DATA=1 (nunca automático): usa o fixture
                    assets/texts_exemplo_julho2026.py (dados de Julho/2026).
    """
    # (1) override explícito por caminho — decisão consciente do operador
    env=os.environ.get("PROG_DADOS")
    if env and os.path.exists(env):
        return _load_pyfile(env)
    # (2) dados reais da competência atual
    real=os.path.join(_cfg_dir(),"dados_mes.py")
    if os.path.exists(real):
        return _load_pyfile(real)
    # (3) sem dados reais: fixture SOMENTE em MODO TESTE explícito
    if use_example_data():
        ex=example_data_path()
        if os.path.exists(ex):
            print("WARNING: MODO TESTE/FIXTURE ativo (PROG_USE_EXAMPLE_DATA=1) — usando dados de "
                  "EXEMPLO de Julho/2026 (RM13). NÃO use em execução real (DEC-019).")
            return _load_pyfile(ex)
        raise RuntimeError("MODO TESTE ativo, mas o fixture não foi encontrado: "+ex)
    # (4) MODO REAL sem dados da competência: interromper para não contaminar (DEC-019)
    raise RuntimeError(DADOS_NAO_CONFIGURADOS_MSG)

def _qa_flags_path():
    """Arquivo local (na pasta de trabalho) com avisos de QA persistentes da competência."""
    return os.path.join(work_dir(),"05_LOGS","qa_warnings.json")

def set_qa_warning(key, message, active=True):
    """Registra (active=True) ou limpa (active=False) um aviso de QA persistente, por chave.
    Idempotente: reexecutar um script atualiza o mesmo 'key' em vez de duplicar. Usado, por
    exemplo, quando a Timeline cai no fallback do último slide (DEC-020) — o QA então NÃO
    aprova automaticamente essa etapa até a conferência visual."""
    p=_qa_flags_path(); data={}
    if os.path.exists(p):
        try: data=json.load(open(p,encoding="utf-8"))
        except Exception: data={}
    if active: data[key]=message
    else: data.pop(key,None)
    os.makedirs(os.path.dirname(p),exist_ok=True)
    with open(p,"w",encoding="utf-8") as f: json.dump(data,f,ensure_ascii=False,indent=2)

def qa_warnings():
    """Dict {chave: mensagem} dos avisos de QA persistentes desta competência (ou {})."""
    p=_qa_flags_path()
    if not os.path.exists(p): return {}
    try: return json.load(open(p,encoding="utf-8"))
    except Exception: return {}

def stamp():
    # NOTE: em scripts de workflow, prefira passar timestamp externamente.
    return datetime.datetime.now().strftime("%Y-%m-%d_%H%M")

def log(msg):
    P=params()
    d=os.path.join(work_dir(),"05_LOGS"); os.makedirs(d,exist_ok=True)
    line=f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    with open(os.path.join(d,"log_execucao.txt"),"a",encoding="utf-8") as f: f.write(line+"\n")
    print(line)

if __name__=="__main__":
    print("project_dir:",project_dir())
    try:
        print("prog_root:",prog_root())
        print("controle_demandas:",controle_demandas_path())
        print("lista_mestre_g2:",lista_mestre_g2_path())
        print("month_folder:",month_folder())
        print("anexo_p11:",anexo_p11_dir())
    except Exception as e:
        print("aviso:",e)

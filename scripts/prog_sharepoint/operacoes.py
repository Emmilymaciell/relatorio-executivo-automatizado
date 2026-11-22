# -*- coding: utf-8 -*-
"""Operações de alto nível no SharePoint do Programa. Escrita só por enviar(): dry-run por padrão; arquivo existente só é
substituído com substituir=True, depois de BAIXAR o atual para backup (SHA-256 + metadados) e com If-Match (eTag),
para não sobrescrever uma versão que mudou no meio do caminho. Depois do envio, baixa de volta e confere o SHA-256.
Sem excluir (por desenho). Mover/renomear e cópia no servidor só pela reorganização autorizada da pasta do mês
(DEC-070), restrita às pastas passadas em `permitidos`."""
import datetime
import hashlib
import json
import os
import tempfile
import time

from . import caminhos as C
from .graph import ErroGraph

OFFICE = (".docx", ".xlsx", ".pptx", ".docm", ".xlsm", ".pptm")


_METADADOS_OFFICE = ("docprops/", "customxml/", "[trash]/")


def partes_office_divergentes(a, b):
    """Partes do pacote Office (ZIP) que diferem entre a e b, IGNORANDO o que o SharePoint acrescenta/reescreve ao
    receber o arquivo: docProps/, customXml/, [trash]/, [Content_Types].xml, *.rels e entradas de diretório.
    Lista vazia = conteúdo (planilhas, texto, slides, mídia) byte a byte igual."""
    import zipfile

    def partes(p):
        with zipfile.ZipFile(p) as z:
            return {n: z.read(n) for n in z.namelist()
                    if not n.endswith("/") and not n.lower().startswith(_METADADOS_OFFICE)
                    and n != "[Content_Types].xml" and not n.endswith(".rels")}
    try:
        pa, pb = partes(a), partes(b)
    except zipfile.BadZipFile:
        return ["<não é pacote Office válido>"]
    return sorted(n for n in set(pa) | set(pb) if pa.get(n) != pb.get(n))


def sha256_arquivo(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloco in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloco)
    return h.hexdigest()


class SharePoint:
    def __init__(self, cfg, graph, prog_root, raiz_projeto):
        self.cfg, self.g, self.prog_root, self.raiz = cfg, graph, prog_root, raiz_projeto
        self._site = self._drive = None

    # ------------------------------------------------------------ resolução
    def remoto(self, alvo):
        """Aceita caminho local (sob prog_root) ou remoto ('4. Produtos/...')."""
        if os.path.isabs(str(alvo)):
            return C.remoto_de_local(alvo, self.prog_root, self.cfg.get("mapa_local_remoto"),
                                     self.cfg.get("prefixo_local", "Programa - "))
        return C.normalizar_remoto(alvo)

    def site(self):
        if not self._site:
            self._site = self.g.get(f"/sites/{self.cfg['hostname']}:{self.cfg['site_path']}")
        return self._site

    def drive(self):
        if not self._drive:
            from urllib.parse import unquote
            alvo = self.cfg["biblioteca"].lower()
            drives = list(self.g.paginar(f"/sites/{self.site()['id']}/drives"))
            for d in drives:
                seg = unquote(d.get("webUrl", "")).rstrip("/").split("/")[-1].lower()
                if seg == alvo or d.get("name", "").lower() == alvo:
                    self._drive = d
                    break
            if not self._drive:
                nomes = ", ".join(f"{d.get('name')} ({d.get('webUrl')})" for d in drives)
                raise ErroGraph(404, "bibliotecaNaoEncontrada", f"biblioteca '{self.cfg['biblioteca']}' não "
                                f"encontrada no site. Disponíveis: {nomes}")
        return self._drive

    def _url_item(self, remoto, sufixo=""):
        r = C.normalizar_remoto(remoto)
        base = f"/drives/{self.drive()['id']}/root"
        return base + (f":/{C.codificar(r)}:" if r else "") + sufixo

    # ------------------------------------------------------------ leitura
    def item(self, alvo):
        try:
            return self.g.get(self._url_item(self.remoto(alvo)))
        except ErroGraph as e:
            if e.status == 404:
                return None
            raise

    def listar(self, alvo=""):
        r = self.remoto(alvo) if alvo else ""
        return list(self.g.paginar(self._url_item(r, "/children")))

    def baixar(self, alvo, destino):
        it = self.item(alvo)
        if not it:
            raise ErroGraph(404, "itemNotFound", f"não existe no SharePoint: {self.remoto(alvo)}")
        if "folder" in it:
            raise ErroGraph(400, "pasta", "baixar() é só para arquivo")
        url = it.get("@microsoft.graph.downloadUrl") or self.g.get(
            self._url_item(self.remoto(alvo)) + "?select=id,@microsoft.graph.downloadUrl").get(
            "@microsoft.graph.downloadUrl")
        os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
        with open(destino, "wb") as fh:
            self.g.baixar_url(url, fh)
        return {"destino": destino, "sha256": sha256_arquivo(destino), "tamanho": os.path.getsize(destino),
                "eTag": it.get("eTag"), "modificado": it.get("lastModifiedDateTime"), "webUrl": it.get("webUrl")}

    def espelhar(self, alvo, destino_raiz, recursivo=True):
        """Copia de LEITURA de uma pasta do SharePoint para `destino_raiz`, mantendo a árvore com o prefixo local
        ('4. Produtos/16. Outubro 2026' -> '<destino_raiz>/Programa - 4. Produtos/16. Outubro 2026'). Serve quando o
        OneDrive não hidrata os arquivos 'somente online' (DEC-058): os scripts rodam com PROG_ROOT=<destino_raiz>.
        Idempotente: arquivo já baixado com o mesmo tamanho e a mesma data de modificação não é baixado de novo.
        Grava `<destino_raiz>/_espelho.json` com eTag, data e SHA-256 de cada arquivo (proveniência)."""
        prefixo = self.cfg.get("prefixo_local", "Programa - ")
        raiz_r = self.remoto(alvo)
        man_path = os.path.join(destino_raiz, "_espelho.json")
        man = {}
        if os.path.exists(man_path):
            with open(man_path, encoding="utf-8") as f:
                man = json.load(f)
        res = {"remoto": raiz_r, "destino": destino_raiz, "baixados": 0, "reaproveitados": 0}

        def local_de(rp):
            partes = rp.split("/")
            p = os.path.abspath(os.path.join(destino_raiz, prefixo + partes[0], *partes[1:]))
            if os.name == "nt" and len(p) >= 250 and not p.startswith("\\\\?\\"):
                p = "\\\\?\\" + p  # caminho longo (nomes do POP) — DEC-056
            return p

        def walk(rpath):
            for i in self.listar(rpath):
                rp = f"{rpath}/{i['name']}"
                if "folder" in i:
                    os.makedirs(local_de(rp), exist_ok=True)
                    if recursivo:
                        walk(rp)
                    continue
                dest = local_de(rp)
                ant = man.get(rp)
                if ant and os.path.isfile(dest) and os.path.getsize(dest) == i.get("size") \
                        and ant.get("modificado") == i.get("lastModifiedDateTime"):
                    res["reaproveitados"] += 1
                    continue
                info = self.baixar(rp, dest)
                man[rp] = {k: info[k] for k in ("sha256", "tamanho", "eTag", "modificado")}
                res["baixados"] += 1

        os.makedirs(local_de(raiz_r), exist_ok=True)
        walk(raiz_r)
        man["_atualizado_em"] = datetime.datetime.now().isoformat(timespec="seconds")
        with open(man_path, "w", encoding="utf-8") as f:
            json.dump(man, f, ensure_ascii=False, indent=1)
        res["arquivos"] = sum(1 for k in man if not k.startswith("_") and k.startswith(raiz_r + "/"))
        return res

    def comparar(self, local):
        """Arquivo local sincronizado × versão no SharePoint (tamanho e SHA-256)."""
        it = self.item(local)
        res = {"local": local, "remoto": self.remoto(local), "existe_remoto": bool(it),
               "tamanho_local": os.path.getsize(local), "sha256_local": sha256_arquivo(local)}
        if it and "file" in it:
            with tempfile.TemporaryDirectory() as tmp:
                b = self.baixar(local, os.path.join(tmp, "x"))
            res.update(tamanho_remoto=b["tamanho"], sha256_remoto=b["sha256"], modificado_remoto=b["modificado"],
                       igual=b["sha256"] == res["sha256_local"])
        return res

    def comparar_pasta(self, local_dir, recursivo=False):
        """Pasta local sincronizada × SharePoint, sem baixar nada: nomes e tamanhos. Serve para saber se o OneDrive
        está em dia ANTES de gerar o relatório a partir da pasta local."""
        ignorar = ("thumbs.db", "desktop.ini", ".ds_store")
        res = {"pasta": self.remoto(local_dir), "so_local": [], "so_sharepoint": [], "tamanho_diferente": [],
               "iguais": 0}

        def walk(ldir, rpath):
            loc = {n: os.path.join(ldir, n) for n in os.listdir(ldir)
                   if not n.startswith("~$") and n.lower() not in ignorar and not n.lower().endswith(".tmp")}
            rem = {i["name"]: i for i in self.listar(rpath)}
            for n in sorted(set(loc) | set(rem)):
                rp = f"{rpath}/{n}"
                if n not in rem:
                    res["so_local"].append(rp)
                elif n not in loc:
                    res["so_sharepoint"].append(rp)
                elif "folder" in rem[n]:
                    if recursivo and os.path.isdir(loc[n]):
                        walk(loc[n], rp)
                elif os.path.isfile(loc[n]) and os.path.getsize(loc[n]) != rem[n].get("size"):
                    res["tamanho_diferente"].append({"item": rp, "local": os.path.getsize(loc[n]),
                                                     "sharepoint": rem[n].get("size"),
                                                     "modificado_sharepoint": rem[n].get("lastModifiedDateTime")})
                else:
                    res["iguais"] += 1
        walk(os.path.abspath(local_dir), res["pasta"])
        res["em_dia"] = not (res["so_local"] or res["so_sharepoint"] or res["tamanho_diferente"])
        return res

    # ------------------------------------------------------------ escrita
    def enviar(self, local, alvo=None, substituir=False, aplicar=False, motivo=""):
        local = os.path.abspath(local)
        if not os.path.isfile(local):
            raise FileNotFoundError(local)
        remoto = self.remoto(alvo) if alvo else self.remoto(local)
        C.exigir_escrita_permitida(remoto, self.cfg.get("bloqueados_escrita"))
        tamanho = os.path.getsize(local)
        sha_local = sha256_arquivo(local)
        existente = self.item(remoto)
        if existente and "folder" in existente:
            raise C.CaminhoInvalido(f"o destino é uma PASTA: {remoto}")
        plano = {"acao": "SUBSTITUIR" if existente else "NOVO", "local": local, "remoto": remoto,
                 "tamanho": tamanho, "sha256_local": sha_local, "aplicado": False}
        if existente:
            plano.update(eTag_atual=existente.get("eTag"), modificado_atual=existente.get("lastModifiedDateTime"),
                         tamanho_atual=existente.get("size"))
            if not substituir:
                plano["acao"] = "BLOQUEADO"
                plano["motivo"] = "já existe no SharePoint — use --substituir (faz backup antes)"
                return plano
        if not aplicar:
            plano["dry_run"] = True
            return plano

        if existente:
            plano["backup"] = self._backup(remoto)
        conflito = "replace" if existente else "fail"
        cab = {"If-Match": existente["eTag"]} if existente else {}
        limite = float(self.cfg.get("upload_simples_max_mb", 4)) * 1024 * 1024
        if tamanho <= limite:
            with open(local, "rb") as f:
                dados = f.read()
            cab.update({"Content-Type": "application/octet-stream"})
            novo = self.g.requisitar("PUT", self._url_item(remoto, "/content") +
                                     f"?@microsoft.graph.conflictBehavior={conflito}", corpo=dados, cabecalhos=cab)
        else:
            sessao = self.g.requisitar("POST", self._url_item(remoto, "/createUploadSession"), cabecalhos=cab,
                                       corpo={"item": {"@microsoft.graph.conflictBehavior": conflito}})
            novo = self.g.enviar_em_sessao(sessao["uploadUrl"], local, tamanho)
        plano["aplicado"] = True
        plano["webUrl"] = novo.get("webUrl")
        plano["verificacao"] = self._verificar(remoto, sha_local, local)
        self._log(dict(plano, motivo=motivo))
        return plano

    def publicar_lote(self, manifesto, aplicar=False):
        """manifesto = {"motivo": str, "itens": [{"local", "remoto"?, "substituir"?: bool}]}. Planeja TODOS antes;
        se algum item estiver BLOQUEADO/inválido, não envia NENHUM. Com aplicar=True envia em ordem e para no 1º erro."""
        motivo = manifesto.get("motivo", "")
        itens = manifesto.get("itens") or []
        if not itens:
            raise C.CaminhoInvalido("manifesto sem itens")
        planos, erros = [], []
        for it in itens:
            try:
                planos.append(self.enviar(it["local"], it.get("remoto"), substituir=bool(it.get("substituir")),
                                          aplicar=False))
            except (C.CaminhoInvalido, FileNotFoundError, ErroGraph) as e:
                erros.append({"local": it.get("local"), "erro": str(e)})
        bloqueados = [p for p in planos if p["acao"] == "BLOQUEADO"]
        res = {"motivo": motivo, "total": len(itens), "planos": planos, "erros": erros, "aplicado": False,
               "pode_aplicar": not (erros or bloqueados)}
        if not aplicar or not res["pode_aplicar"]:
            return res
        enviados = []
        for it in itens:
            enviados.append(self.enviar(it["local"], it.get("remoto"), substituir=bool(it.get("substituir")),
                                        aplicar=True, motivo=motivo))
        res.update(planos=enviados, aplicado=True)
        return res

    # ------------------------------------------------------------ reorganização (DEC-070)
    # Mover/renomear e copiar no servidor existem SÓ para organizar a pasta do mês no padrão do POP, com autorização
    # explícita da usuária e depois de guardar a cópia "como enviado" no Desenvolvimento. Excluir continua NÃO existindo.
    def arvore(self, alvo):
        """{caminho relativo: {size, qxh, id, eTag}} de todos os ARQUIVOS sob a pasta + conjunto de subpastas."""
        raiz = self.remoto(alvo)
        arqs, pastas = {}, set()

        def walk(rem, rel):
            for it in self.listar(rem):
                r2 = f"{rel}/{it['name']}" if rel else it["name"]
                if "folder" in it:
                    pastas.add(r2)
                    walk(f"{rem}/{it['name']}", r2)
                else:
                    arqs[r2] = {"size": it.get("size"), "qxh": (it.get("file") or {}).get("hashes", {}).get(
                        "quickXorHash"), "id": it.get("id"), "eTag": it.get("eTag"),
                        "criado_em": it.get("createdDateTime"), "modificado_em": it.get("lastModifiedDateTime"),
                        "criado_por": ((it.get("createdBy") or {}).get("user") or {}).get("displayName"),
                        "modificado_por": ((it.get("lastModifiedBy") or {}).get("user") or {}).get("displayName")}
        walk(raiz, "")
        return arqs, pastas

    def _exigir_reorganizavel(self, remoto, permitidos):
        C.exigir_escrita_permitida(remoto, self.cfg.get("bloqueados_escrita"))
        r = C._norm(C.normalizar_remoto(remoto))
        if not any(r == C._norm(C.normalizar_remoto(p)) or r.startswith(C._norm(C.normalizar_remoto(p)) + "/")
                   for p in (permitidos or [])):
            raise C.CaminhoInvalido(f"reorganização só é permitida sob {permitidos}: {remoto}")

    def criar_pasta(self, alvo, aplicar=False):
        """Cria a pasta (e as intermediárias). Idempotente; nunca substitui."""
        remoto = self.remoto(alvo)
        C.exigir_escrita_permitida(remoto, self.cfg.get("bloqueados_escrita"))
        partes = remoto.split("/")
        criadas = []
        for i in range(1, len(partes) + 1):
            sub = "/".join(partes[:i])
            it = self.item(sub)
            if it:
                if "folder" not in it:
                    raise C.CaminhoInvalido(f"existe um ARQUIVO com o nome da pasta: {sub}")
                continue
            criadas.append(sub)
            if aplicar:
                pai = "/".join(partes[:i - 1])
                self.g.requisitar("POST", self._url_item(pai, "/children"), corpo={
                    "name": partes[i - 1], "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
        if aplicar and criadas:
            self._log({"acao": "CRIAR_PASTA", "remoto": remoto, "criadas": criadas})
        return {"remoto": remoto, "criadas": criadas, "aplicado": bool(aplicar)}

    def mover(self, origem, destino, permitidos, aplicar=False, motivo=""):
        """Move e/ou renomeia um item (arquivo ou pasta) DENTRO das pastas `permitidos`. Recusa destino existente
        (conflictBehavior=fail), usa If-Match (eTag) e confere depois que tamanho e quickXorHash não mudaram."""
        o, d = self.remoto(origem), self.remoto(destino)
        self._exigir_reorganizavel(o, permitidos)
        self._exigir_reorganizavel(d, permitidos)
        it = self.item(o)
        if not it:
            raise ErroGraph(404, "itemNotFound", f"origem não existe: {o}")
        if self.item(d):
            return {"acao": "BLOQUEADO", "origem": o, "destino": d, "motivo": "destino já existe"}
        hashes = (it.get("file") or {}).get("hashes", {})
        plano = {"acao": "MOVER", "origem": o, "destino": d, "tamanho": it.get("size"), "eTag": it.get("eTag"),
                 "qxh": hashes.get("quickXorHash"), "aplicado": False}
        if not aplicar:
            plano["dry_run"] = True
            return plano
        pai, nome = d.rsplit("/", 1)
        self.criar_pasta(pai, aplicar=True)
        corpo = {"name": nome}
        if pai != o.rsplit("/", 1)[0]:
            corpo["parentReference"] = {"id": self.item(pai)["id"]}
        cab = {"If-Match": it["eTag"]} if it.get("eTag") else {}
        self.g.requisitar("PATCH", f"/drives/{self.drive()['id']}/items/{it['id']}"
                                   "?@microsoft.graph.conflictBehavior=fail", corpo=corpo, cabecalhos=cab)
        novo = self.item(d)
        ok = bool(novo) and novo.get("size") == it.get("size") and (
            "folder" in it or (novo.get("file") or {}).get("hashes", {}).get("quickXorHash") == hashes.get("quickXorHash"))
        plano.update(aplicado=True, verificacao="OK" if ok else "DIVERGENTE", origem_existe=bool(self.item(o)))
        self._log(dict(plano, motivo=motivo))
        return plano

    def copiar_pasta(self, origem, destino, aplicar=False, motivo="", timeout=3600, intervalo=10):
        """Cópia NO SERVIDOR de uma pasta inteira (bytes idênticos, sem baixar). O destino não pode existir.
        Confere arquivo a arquivo (caminho, tamanho e quickXorHash) contra a origem."""
        o, d = self.remoto(origem), self.remoto(destino)
        C.exigir_escrita_permitida(d, self.cfg.get("bloqueados_escrita"))
        it = self.item(o)
        if not it or "folder" not in it:
            raise C.CaminhoInvalido(f"origem não é pasta: {o}")
        if self.item(d):
            return {"acao": "BLOQUEADO", "origem": o, "destino": d, "motivo": "destino já existe"}
        arqs_o, pastas_o = self.arvore(o)
        plano = {"acao": "COPIAR_PASTA", "origem": o, "destino": d, "arquivos": len(arqs_o),
                 "bytes": sum(v["size"] or 0 for v in arqs_o.values()), "aplicado": False}
        if not aplicar:
            plano["dry_run"] = True
            return plano
        pai, nome = d.rsplit("/", 1)
        self.criar_pasta(pai, aplicar=True)
        _, cab = self.g.requisitar("POST", f"/drives/{self.drive()['id']}/items/{it['id']}/copy"
                                           "?@microsoft.graph.conflictBehavior=fail", corpo={
            "parentReference": {"driveId": self.drive()["id"], "id": self.item(pai)["id"]}, "name": nome},
            com_cabecalhos=True)
        monitor = {k.lower(): v for k, v in (cab or {}).items()}.get("location")
        inicio, estado = time.time(), None
        while time.time() - inicio < timeout:
            if monitor:
                try:
                    estado = self.g.requisitar("GET", monitor, autenticar=False).get("status")
                except ErroGraph:
                    estado = None
            if estado in ("failed", "cancelled"):
                break
            if estado in (None, "completed") and self.item(d):
                arqs_d, pastas_d = self.arvore(d)
                if len(arqs_d) >= len(arqs_o):
                    break
            self.g._dormir(intervalo) if hasattr(self.g, "_dormir") else time.sleep(intervalo)
        arqs_d, pastas_d = self.arvore(d) if self.item(d) else ({}, set())
        difs = [k for k in sorted(set(arqs_o) | set(arqs_d))
                if (arqs_o.get(k) or {}).get("size") != (arqs_d.get(k) or {}).get("size")
                or (arqs_o.get(k) or {}).get("qxh") != (arqs_d.get(k) or {}).get("qxh")]
        plano.update(aplicado=True, estado_monitor=estado, arquivos_destino=len(arqs_d),
                     pastas_iguais=pastas_o == pastas_d, divergentes=difs[:50],
                     verificacao="OK" if not difs and pastas_o == pastas_d else "DIVERGENTE")
        self._log(dict(plano, motivo=motivo))
        return plano

    def _verificar(self, remoto, sha_local, local=None):
        with tempfile.TemporaryDirectory() as tmp:
            b = self.baixar(remoto, os.path.join(tmp, "v"))
            if b["sha256"] == sha_local:
                return {"status": "OK", "sha256_remoto": b["sha256"]}
            if remoto.lower().endswith(OFFICE) and local:
                difs = partes_office_divergentes(local, b["destino"])
                if not difs:
                    return {"status": "OK_CONTEUDO", "sha256_remoto": b["sha256"], "tamanho_remoto": b["tamanho"],
                            "nota": "conteúdo idêntico; o SharePoint só acrescentou metadados (docProps/customXml)"}
                return {"status": "DIVERGENTE", "sha256_remoto": b["sha256"], "partes_divergentes": difs[:20]}
        return {"status": "DIVERGENTE", "sha256_remoto": b["sha256"], "tamanho_remoto": b["tamanho"],
                "nota": "conteúdo no servidor difere do enviado — conferir"}

    def _backup(self, remoto):
        carimbo = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        base = os.path.join(self.raiz, self.cfg.get("backups_dir", "output/SHAREPOINT_BACKUPS"), carimbo)
        destino = os.path.abspath(os.path.join(base, *C.normalizar_remoto(remoto).split("/")))
        if os.name == "nt" and len(destino) >= 240 and not destino.startswith("\\\\?\\"):
            destino = "\\\\?\\" + destino  # caminho longo (DEC-056)
        info = self.baixar(remoto, destino)
        with open(destino + ".meta.json", "w", encoding="utf-8") as f:
            json.dump(dict(info, remoto=remoto), f, ensure_ascii=False, indent=2)
        return info

    def _log(self, registro):
        d = os.path.join(self.raiz, self.cfg.get("log_dir", "logs/sharepoint"))
        os.makedirs(d, exist_ok=True)
        registro = dict(registro, quando=datetime.datetime.now().isoformat(timespec="seconds"))
        with open(os.path.join(d, datetime.date.today().strftime("%Y-%m") + ".jsonl"), "a", encoding="utf-8") as f:
            f.write(json.dumps(registro, ensure_ascii=False) + "\n")

    # ------------------------------------------------------------ diagnóstico
    def testar(self):
        self.g.token()
        claims = self.g.claims
        out = {"app": claims.get("app_displayname"), "tenant": claims.get("tid"),
               "permissoes_app": sorted(claims.get("roles", []))}
        s = self.site()
        out["site"] = {"nome": s.get("displayName"), "webUrl": s.get("webUrl")}
        d = self.drive()
        out["biblioteca"] = {"nome": d.get("name"), "webUrl": d.get("webUrl")}
        out["raiz"] = sorted(i["name"] for i in self.listar(""))
        return out

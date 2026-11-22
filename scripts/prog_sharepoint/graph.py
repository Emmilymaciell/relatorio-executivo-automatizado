# -*- coding: utf-8 -*-
"""Cliente mínimo da Microsoft Graph (só urllib). Token app-only (client_credentials) mantido só em memória."""
import base64
import json
import time
import urllib.error
import urllib.parse
import urllib.request

GRAPH = "https://graph.microsoft.com/v1.0"
LOGIN = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
CHUNK = 320 * 1024 * 32  # 10 MiB, múltiplo de 320 KiB (exigência do upload em sessão)
TENTATIVAS = 5


class ErroGraph(Exception):
    def __init__(self, status, codigo, mensagem, url=""):
        self.status, self.codigo, self.mensagem, self.url = status, codigo, mensagem, url
        dica = ""
        if status in (401, 403):
            dica = (" — verifique no Entra ID se o app tem permissão de APLICATIVO (Sites.Selected com o site Programa "
                    "liberado, ou Sites.Read.All/Sites.ReadWrite.All) com consentimento do administrador")
        super().__init__(f"HTTP {status} {codigo}: {mensagem}{dica}")


def _decodificar_jwt(token):
    try:
        carga = token.split(".")[1]
        carga += "=" * (-len(carga) % 4)
        return json.loads(base64.urlsafe_b64decode(carga))
    except Exception:
        return {}


class Graph:
    def __init__(self, tenant_id, client_id, client_secret, abrir=None, dormir=time.sleep):
        self._tenant, self._client, self._segredo = tenant_id, client_id, client_secret
        self._abrir = abrir or urllib.request.urlopen
        self._dormir = dormir
        self._token, self._expira = None, 0
        self.claims = {}

    # ------------------------------------------------------------------ token
    def token(self):
        if self._token and time.time() < self._expira - 60:
            return self._token
        dados = urllib.parse.urlencode({
            "client_id": self._client, "client_secret": self._segredo,
            "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}).encode()
        req = urllib.request.Request(LOGIN.format(tenant=self._tenant), data=dados, method="POST",
                                     headers={"Content-Type": "application/x-www-form-urlencoded"})
        try:
            with self._abrir(req, timeout=60) as r:
                j = json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            corpo = _ler_erro(e)
            raise ErroGraph(e.code, corpo.get("error", "token"),
                            (corpo.get("error_description") or "").split("\r\n")[0], LOGIN) from None
        self._token = j["access_token"]
        self._expira = time.time() + int(j.get("expires_in", 3600))
        self.claims = _decodificar_jwt(self._token)
        return self._token

    # ------------------------------------------------------------------ HTTP
    def requisitar(self, metodo, url, corpo=None, cabecalhos=None, autenticar=True, bruto=False,
                   com_cabecalhos=False):
        """com_cabecalhos=True devolve (json, cabeçalhos) — usado pela cópia assíncrona (Location do monitor)."""
        if not url.startswith("http"):
            url = GRAPH + url
        hdr = dict(cabecalhos or {})
        dados = corpo
        if isinstance(corpo, (dict, list)):
            dados = json.dumps(corpo).encode("utf-8")
            hdr.setdefault("Content-Type", "application/json")
        for tentativa in range(1, TENTATIVAS + 1):
            h = dict(hdr)
            if autenticar:
                h["Authorization"] = "Bearer " + self.token()
            req = urllib.request.Request(url, data=dados, method=metodo, headers=h)
            try:
                with self._abrir(req, timeout=300) as r:
                    conteudo = r.read()
                    if bruto:
                        return conteudo
                    j = json.loads(conteudo.decode("utf-8")) if conteudo else {}
                    if com_cabecalhos:
                        return j, dict(r.headers.items()) if getattr(r, "headers", None) else {}
                    return j
            except urllib.error.HTTPError as e:
                if e.code in (429, 500, 502, 503, 504) and tentativa < TENTATIVAS:
                    espera = e.headers.get("Retry-After") if e.headers else None
                    self._dormir(min(int(espera) if espera and espera.isdigit() else 2 ** tentativa, 60))
                    continue
                corpo_erro = _ler_erro(e).get("error", {})
                if isinstance(corpo_erro, str):
                    corpo_erro = {"code": corpo_erro, "message": ""}
                raise ErroGraph(e.code, corpo_erro.get("code", ""), corpo_erro.get("message", ""), url) from None

    def get(self, url):
        return self.requisitar("GET", url)

    def paginar(self, url):
        while url:
            j = self.get(url)
            yield from j.get("value", [])
            url = j.get("@odata.nextLink")

    def baixar_url(self, url_download, destino_fh):
        """URL pré-autenticada (@microsoft.graph.downloadUrl): sem cabeçalho Authorization."""
        req = urllib.request.Request(url_download, method="GET")
        with self._abrir(req, timeout=600) as r:
            while True:
                bloco = r.read(1024 * 1024)
                if not bloco:
                    break
                destino_fh.write(bloco)

    def enviar_em_sessao(self, url_sessao, caminho_local, tamanho):
        """PUT dos blocos na URL da sessão (pré-autenticada). Devolve o driveItem final."""
        resultado = {}
        with open(caminho_local, "rb") as f:
            inicio = 0
            while inicio < tamanho:
                bloco = f.read(CHUNK)
                fim = inicio + len(bloco) - 1
                resultado = self.requisitar("PUT", url_sessao, corpo=bloco, autenticar=False, cabecalhos={
                    "Content-Length": str(len(bloco)), "Content-Range": f"bytes {inicio}-{fim}/{tamanho}"})
                inicio = fim + 1
        return resultado


def _ler_erro(e):
    try:
        return json.loads(e.read().decode("utf-8"))
    except Exception:
        return {}

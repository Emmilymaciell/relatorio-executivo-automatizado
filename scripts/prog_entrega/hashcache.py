# -*- coding: utf-8 -*-
"""SHA-256 com cache (caminho + tamanho + mtime_ns) e cache de RESULTADOS de validação por conteúdo.

- O hash só é reaproveitado se tamanho E mtime_ns forem idênticos; qualquer mudança real => recalcula.
- Resultados de validação são indexados por (sha256, versão do validador): mesmos bytes => mesmo
  resultado; um arquivo novo/alterado (ex.: ZIP gerado no staging) tem outro hash e é validado do zero.
"""
import hashlib
import json
import os

_BLOCO = 1024 * 1024


def sha256_arquivo(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(_BLOCO)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


class Cache:
    def __init__(self, path=None):
        self.path = path
        self.hashes, self.validacoes = {}, {}
        if path and os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    d = json.load(f)
                self.hashes, self.validacoes = d.get("hashes", {}), d.get("validacoes", {})
            except Exception:
                pass  # cache corrompido = sem cache (nunca é fonte de verdade)

    def sha256(self, path):
        st = os.stat(path)
        k = os.path.normcase(os.path.abspath(path))
        e = self.hashes.get(k)
        if e and e["size"] == st.st_size and e["mtime_ns"] == st.st_mtime_ns:
            return e["sha256"]
        h = sha256_arquivo(path)
        self.hashes[k] = {"size": st.st_size, "mtime_ns": st.st_mtime_ns, "sha256": h}
        return h

    def validacao(self, sha, versao, func):
        k = f"{sha}:{versao}"
        if k not in self.validacoes:
            self.validacoes[k] = func()
        return self.validacoes[k]

    def salvar(self):
        if not self.path:
            return
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"hashes": self.hashes, "validacoes": self.validacoes}, f, ensure_ascii=False)
        os.replace(tmp, self.path)

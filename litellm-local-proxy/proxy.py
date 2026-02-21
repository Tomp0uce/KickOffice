#!/usr/bin/env python3
"""
LiteLLM Local Proxy - Référence
================================
Proxy HTTP local qui ajoute les headers d'authentification LiteLLM
à chaque requête avant de la transmettre au serveur LiteLLM de l'entreprise.

Usage:
    cp .auth.env.template .auth.env
    # Éditez .auth.env avec vos credentials
    python proxy.py

Le proxy écoute sur http://localhost:4000
"""

import os
import json
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.request import urlopen, Request
from urllib.error import URLError

# Chargement des credentials depuis .auth.env
def load_env(path=".auth.env"):
    env = {}
    if not os.path.exists(path):
        print(f"[WARN] {path} introuvable. Copiez .auth.env.template en .auth.env")
        return env
    with open(path) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, value = line.partition("=")
                env[key.strip()] = value.strip()
    return env

ENV = load_env()
LITELLM_BASE_URL = ENV.get("LITELLM_BASE_URL", "https://litellm.kickmaker.net/v1")
LITELLM_USER_KEY = ENV.get("LITELLM_USER_KEY", "")
LITELLM_USER_EMAIL = ENV.get("LITELLM_USER_EMAIL", "")
PROXY_PORT = int(ENV.get("PROXY_PORT", "4000"))


class ProxyHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length)

        target_url = LITELLM_BASE_URL.rstrip("/") + self.path

        headers = {
            "Content-Type": self.headers.get("Content-Type", "application/json"),
            "Authorization": self.headers.get("Authorization", ""),
            "X-User-Key": LITELLM_USER_KEY,
            "X-OpenWebUi-User-Email": LITELLM_USER_EMAIL,
        }

        req = Request(target_url, data=body, headers=headers, method="POST")
        try:
            with urlopen(req) as resp:
                self.send_response(resp.status)
                for key, value in resp.headers.items():
                    self.send_header(key, value)
                self.end_headers()
                self.wfile.write(resp.read())
        except URLError as e:
            self.send_response(502)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())

    def log_message(self, format, *args):
        print(f"[PROXY] {self.address_string()} - {format % args}")


if __name__ == "__main__":
    if not LITELLM_USER_KEY or not LITELLM_USER_EMAIL:
        print("[ERROR] LITELLM_USER_KEY et LITELLM_USER_EMAIL sont requis dans .auth.env")
        exit(1)

    server = HTTPServer(("localhost", PROXY_PORT), ProxyHandler)
    print(f"[PROXY] Démarrage sur http://localhost:{PROXY_PORT}")
    print(f"[PROXY] Cible : {LITELLM_BASE_URL}")
    print(f"[PROXY] Email : {LITELLM_USER_EMAIL}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("[PROXY] Arrêt.")
        server.server_close()

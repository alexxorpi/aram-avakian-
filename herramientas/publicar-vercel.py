#!/usr/bin/env python3
"""Publica el sitio en Vercel con su API (sin instalar nada).

Uso:  python3 herramientas/publicar-vercel.py
La autenticación la pone el entorno (proxy) o la variable VERCEL_TOKEN.
"""
import hashlib, json, os, subprocess, sys, urllib.request, urllib.error

API = "https://api.vercel.com"
NOMBRE = "avakian-inmobiliaria"
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OMITIR = ("herramientas/", "capturas/", "LEEME.md", ".gitignore")


def llamar(metodo, ruta, datos=None, cabeceras=None):
    h = {"User-Agent": "sitio-estatico/1.0"}
    if os.environ.get("VERCEL_TOKEN"):
        h["Authorization"] = "Bearer " + os.environ["VERCEL_TOKEN"]
    h.update(cabeceras or {})
    req = urllib.request.Request(API + ruta, data=datos, method=metodo, headers=h)
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            cuerpo = r.read()
    except urllib.error.HTTPError as e:
        sys.exit(f"{metodo} {ruta} -> {e.code}: {e.read().decode()[:400]}")
    return json.loads(cuerpo) if cuerpo else {}


def main():
    archivos = subprocess.check_output(["git", "ls-files"], cwd=RAIZ, text=True).split("\n")
    archivos = [a for a in archivos if a and not a.startswith(OMITIR)]
    lista = []
    for a in archivos:
        contenido = open(os.path.join(RAIZ, a), "rb").read()
        sha = hashlib.sha1(contenido).hexdigest()
        llamar("POST", "/v2/files", contenido, {
            "Content-Type": "application/octet-stream", "x-vercel-digest": sha, "Content-Length": str(len(contenido))})
        lista.append({"file": a, "sha": sha, "size": len(contenido)})
    print(f"{len(lista)} archivos subidos")
    cuerpo = json.dumps({
        "name": NOMBRE, "files": lista, "target": "production",
        "projectSettings": {"framework": None, "buildCommand": None, "installCommand": None, "outputDirectory": None},
    }).encode()
    d = llamar("POST", "/v13/deployments?skipAutoDetectionConfirmation=1", cuerpo, {"Content-Type": "application/json"})
    print("id:", d.get("id"), "| estado:", d.get("readyState"))
    print("url:", d.get("url"))
    print("alias:", d.get("alias"))


if __name__ == "__main__":
    main()

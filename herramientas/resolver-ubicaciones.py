#!/usr/bin/env python3
"""Convierte los links de Google Maps de la hoja en coordenadas.

Lee la hoja a través del conector (FUENTE en assets/app.js), abre cada link de Google Maps
que encuentre en «Ubicación» o «Dirección» y guarda sus coordenadas en datos/ubicaciones.json.
El sitio usa ese archivo para poner el pin. Lo ejecuta GitHub cada pocos minutos.
"""
import json, os, re, sys, urllib.request, urllib.error

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SALIDA = os.path.join(RAIZ, "datos", "ubicaciones.json")
PATRONES = [r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)", r"@(-?\d+\.\d+),(-?\d+\.\d+)",
            r"[?&](?:q|ll|query|center)=(-?\d+\.\d+)(?:,|%2C)(-?\d+\.\d+)"]
UA = {"User-Agent": "Mozilla/5.0 (compatible; sitio-inmobiliario/1.0)"}


class SinRedireccion(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):
        return None


def coords(texto):
    for p in PATRONES:
        m = re.search(p, texto)
        if m:
            return f"{m.group(1)}, {m.group(2)}"
    return ""


def resolver(url):
    abrir = urllib.request.build_opener(SinRedireccion)
    for _ in range(5):
        c = coords(url)
        if c:
            return c
        try:
            r = abrir.open(urllib.request.Request(url, headers=UA), timeout=30)
            return coords(r.read(400000).decode("utf-8", "ignore"))
        except urllib.error.HTTPError as e:
            sig = e.headers.get("Location")
            if not sig:
                return ""
            url = sig
    return ""


def main():
    js = open(os.path.join(RAIZ, "assets", "app.js"), encoding="utf-8").read()
    m = re.search(r"^const FUENTE = '([^']+)'", js, re.M)
    if not m:
        print("Sin FUENTE configurada; nada que hacer.")
        return
    datos = json.load(urllib.request.urlopen(urllib.request.Request(m.group(1), headers=UA), timeout=120))
    links = set()
    for nombre, filas in datos.get("hojas", {}).items():
        if not re.match(r"(?i)(propiedades|desarrollos)", nombre) or not filas:
            continue
        cols = [i for i, t in enumerate(filas[0]) if re.match(r"(?i)(ubicaci|direcci)", str(t))]
        for fila in filas[1:]:
            for i in cols:
                v = str(fila[i]).strip() if i < len(fila) else ""
                if re.match(r"(?i)https?://", v) and not coords(v):
                    links.add(v)
    try:
        guardado = json.load(open(SALIDA, encoding="utf-8"))
    except (OSError, ValueError):
        guardado = {}
    nuevo = {k: v for k, v in guardado.items() if k in links}      # se olvidan los links que ya no están
    for link in sorted(links - set(nuevo)):
        c = resolver(link)
        print(("OK  " if c else "SIN ") + link + (" -> " + c if c else ""))
        if c:
            nuevo[link] = c
    if nuevo != guardado:
        with open(SALIDA, "w", encoding="utf-8") as f:
            json.dump(nuevo, f, ensure_ascii=False, indent=1, sort_keys=True)
            f.write("\n")
        print("ubicaciones.json actualizado")
    else:
        print("Sin cambios")


if __name__ == "__main__":
    sys.exit(main())

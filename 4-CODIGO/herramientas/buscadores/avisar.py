#!/usr/bin/env python3
"""Avisa a los buscadores que el sitio cambió (IndexNow): Bing, Yandex, Seznam, Naver y Yep lo
reciben a la vez; DuckDuckGo, Ecosia y Yahoo usan lo de Bing. Google no usa IndexNow: a Google se le
da el sitemap una vez en Search Console y él vuelve solo (ver docs/DOMINIO-Y-BUSCADORES.md).

    python3 avisar.py                      # la dirección principal que diga el servidor (robots.txt)
    python3 avisar.py https://nom.qd.je    # una en particular

Lee las direcciones del sitemap.xml que sirve ese sitio, así que se corre DESPUÉS de publicar.
La clave es el archivo static/<clave>.txt (32 letras hexadecimales); el sitio tiene que servirlo.
Respuestas: 200 o 202 = recibido; 403 = la clave no se encontró en el sitio; 422 = direcciones de
otro sitio; 429 = demasiados avisos (esperar).
"""
import glob
import json
import os
import re
import sys
import urllib.error
import urllib.request

AQUI = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.normpath(os.path.join(AQUI, "..", "..", "app", "src", "main", "resources", "static"))


def bajar(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "nomcolab-indexnow/1"}), timeout=30) as r:
        return r.read().decode("utf-8")


def main():
    claves = [os.path.basename(f)[:-4] for f in glob.glob(os.path.join(STATIC, "*.txt")) if re.fullmatch(r"[0-9a-f]{32}\.txt", os.path.basename(f))]
    if len(claves) != 1:
        sys.exit(f"avisar: se esperaba una sola clave <32 hex>.txt en {STATIC}, hay {len(claves)}")
    clave = claves[0]
    sitio = (sys.argv[1] if len(sys.argv) > 1 else "https://nomcolab.dpdns.org").rstrip("/")
    if len(sys.argv) == 1:
        # La dirección principal la dice el propio servidor en robots.txt (NOM_CANONICO).
        m = re.search(r"^Sitemap:\s*(https?://[^/\s]+)/sitemap\.xml", bajar(sitio + "/robots.txt"), re.M)
        if m: sitio = m.group(1)
    if bajar(f"{sitio}/{clave}.txt").strip() != clave:
        sys.exit(f"avisar: {sitio}/{clave}.txt no tiene la clave: publica el servidor primero")
    urls = re.findall(r"<loc>([^<]+)</loc>", bajar(sitio + "/sitemap.xml"))
    host = re.sub(r"^https?://", "", sitio)
    urls = [u for u in urls if re.sub(r"^https?://", "", u).split("/")[0] == host]
    if not urls:
        sys.exit("avisar: el sitemap no tiene direcciones de " + host)
    cuerpo = json.dumps({"host": host, "key": clave, "keyLocation": f"{sitio}/{clave}.txt", "urlList": urls}).encode()
    pedido = urllib.request.Request("https://api.indexnow.org/indexnow", data=cuerpo, method="POST",
                                    headers={"Content-Type": "application/json; charset=utf-8", "User-Agent": "nomcolab-indexnow/1"})
    try:
        with urllib.request.urlopen(pedido, timeout=30) as r:
            print(f"IndexNow: {r.status} — {len(urls)} direcciones de {host} avisadas")
    except urllib.error.HTTPError as e:
        sys.exit(f"IndexNow: {e.code} {e.reason} — {e.read().decode('utf-8', 'replace')[:300]}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Toma las capturas de evidencia del plan de pruebas (../imagenes/ev-*.jpg).

    ~/.cache/nom-pruebas/venv/bin/python evidencias.py     # con el servidor local en el puerto 8099

Las pantallas de entrada, de crear cuenta y de términos se toman aquí; las del espacio y los
editores salen de las capturas de la presentación (7-PRESENTACION/fuentes/capturas.py), para
que los dos documentos muestren lo mismo.
"""
import os
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", os.path.expanduser("~/.cache/nom-pruebas/navegadores"))
from PIL import Image
from playwright.sync_api import sync_playwright

AQUI = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(AQUI, "..", "imagenes")
APP = os.path.join(AQUI, "..", "..", "7-PRESENTACION", "imagenes", "app")
BASE = os.environ.get("NOM_BASE", "http://localhost:8099")


def jpg(origen, nombre, ancho=None):
    im = Image.open(origen).convert("RGB")
    if ancho: im = im.resize((ancho, round(im.height * ancho / im.width)), Image.LANCZOS)
    im.save(os.path.join(IMG, nombre), quality=88)


def main():
    tmp = os.path.join(IMG, "_tmp.png")
    with sync_playwright() as p:
        nav = p.chromium.launch()
        pag = nav.new_page(viewport={"width": 606, "height": 800})
        pag.goto(BASE + "/", wait_until="networkidle"); pag.wait_for_timeout(800)
        pag.screenshot(path=tmp, full_page=True); jpg(tmp, "ev-entrar.jpg")
        pag.click("text=Crear cuenta"); pag.wait_for_timeout(500)
        pag.fill("#re", "WACK"); pag.focus("#ru")
        pag.screenshot(path=tmp, full_page=True); jpg(tmp, "ev-crear.jpg")
        pag.close()
        # Términos: la parte de arriba (título, buscador de la ayuda e índice del artículo).
        pag = nav.new_page(viewport={"width": 1366, "height": 768})
        pag.goto(BASE + "/legal/terminos", wait_until="networkidle"); pag.wait_for_timeout(500)
        fin = pag.evaluate("(() => { const i = document.querySelector('.aviso'); return i ? Math.floor(i.getBoundingClientRect().top) - 8 : 768; })()")
        pag.screenshot(path=tmp, clip={"x": 0, "y": 0, "width": 1366, "height": min(fin, 768)}); jpg(tmp, "ev-terminos.jpg", 1100)
        nav.close()
    os.remove(tmp)
    jpg(os.path.join(APP, "espacio-pc.png"), "ev-espacio.jpg", 1100)
    jpg(os.path.join(APP, "texto-pc.png"), "ev-texto.jpg", 1100)
    jpg(os.path.join(APP, "hoja-cel.png"), "ev-hoja-cel.jpg", 520)
    print("listo: 6 capturas en", os.path.abspath(IMG))


if __name__ == "__main__":
    main()

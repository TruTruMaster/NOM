#!/usr/bin/env python3
"""Capturas de teléfono para la ficha de Google Play (1080 × 1920, proporción 9:16; Play pide que el lado largo no
pase del doble del corto) y el gráfico de funciones de 1024 × 500. Salen a 4-CODIGO/marca/play/.

    NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python 4-CODIGO/herramientas/play/capturas_play.py

Crea una cuenta y un espacio de ejemplo en ese servidor (usar el de desarrollo, no la Pi)."""
import base64, os, random, sys
from pathlib import Path
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", os.path.expanduser("~/.cache/nom-pruebas/navegadores"))
CODIGO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(CODIGO / "pruebas" / "interfaz"))
import nomprueba as N  # noqa: E402

SALIDA = CODIGO / "marca" / "play"
DOCUMENTOS = (("texto", "acta", "TEXTO", "Acta de la reunión de equipo"), ("hoja", "calificaciones", "HOJA", "Calificaciones del parcial"),
              ("pres", "feria-ciencias", "PRESENTACION", "Feria de ciencias"), ("diag", "flujo", "DIAGRAMA", "Diagrama de flujo del registro"))


def capturas():
    suf = str(random.randint(100, 999))
    c = N.cuenta("equipo" + suf); c2 = N.cuenta("ana" + suf); c3 = N.cuenta("luis" + suf)
    esp = N.espacio(c["token"], "Proyecto de Ciencias")
    for otro, permiso in ((c2, "EDICION"), (c3, "LECTURA")):
        N.api("POST", f"/api/espacios/{esp}/accesos", {"nombre": otro["nombre"], "permiso": permiso}, c["token"])
    nav = N.Navegador(movil=True)
    nav.ctx.close()
    nav.ctx = nav.b.new_context(viewport={"width": 360, "height": 640}, is_mobile=True, has_touch=True, device_scale_factor=3)
    nav.p = nav.ctx.new_page(); p = nav.p
    nav.entrar(c); p.wait_for_timeout(1500)
    p.evaluate("() => { try { NOM.hito('rueda'); } catch (e) {} }")
    docs = {}
    for clave, pid, modo, nombre in DOCUMENTOS:
        ops = p.evaluate("id => NOM.plantillas.LISTA.find(x => x.id === id).crear()", pid)
        docs[clave] = N.documento(c["token"], esp, modo, nombre, ops)
    p.goto(N.B + f"/#/espacio/{esp}"); p.wait_for_timeout(2000); p.screenshot(path=str(SALIDA / "1-espacio.png"))
    for i, clave in enumerate(("texto", "hoja", "pres", "diag"), 2):
        p.goto(N.B + f"/#/doc/{docs[clave]}"); p.wait_for_timeout(2800); p.screenshot(path=str(SALIDA / f"{i}-{clave}.png"))
    print("errores de la página:", nav.errores[:3])
    nav.cerrar()


def grafico():
    """1024 × 500: la marca sobre el fondo de la marca; lo importante va al centro (Play recorta las orillas)."""
    from playwright.sync_api import sync_playwright
    d = lambda f: "data:image/svg+xml;base64," + base64.b64encode((CODIGO / "marca" / f).read_bytes()).decode()
    fuente = "data:font/woff2;base64," + base64.b64encode((CODIGO / "app/src/main/resources/static/fonts/nunito-latin-wght-normal.woff2").read_bytes()).decode()
    html = f"""<!doctype html><html><head><meta charset="utf-8"><style>
@font-face {{ font-family: Nunito; src: url({fuente}) format('woff2'); font-weight: 200 1000; }}
html, body {{ margin: 0; }} body {{ width: 1024px; height: 500px; background: #1D1B36; font-family: Nunito, sans-serif; color: #fff;
  display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 22px; }}
img {{ height: 96px; }} p {{ margin: 0; font-size: 34px; font-weight: 800; letter-spacing: -.3px; }}
span {{ font-size: 20px; font-weight: 600; color: #C9C8DE; }}
</style></head><body><img src="{d('logo-horizontal-oscuro.svg')}" alt=""><p>Documentos en equipo, con Internet o sin él.</p>
<span>Texto · hoja de cálculo · presentación · diagrama — abre Word, Excel y PowerPoint</span></body></html>"""
    with sync_playwright() as pw:
        b = pw.chromium.launch(); pg = b.new_page(viewport={"width": 1024, "height": 500})
        pg.set_content(html); pg.wait_for_timeout(500); pg.screenshot(path=str(SALIDA / "grafico-1024x500.png")); b.close()


if __name__ == "__main__":
    SALIDA.mkdir(parents=True, exist_ok=True)
    grafico()
    capturas()
    print("listo:", sorted(x.name for x in SALIDA.iterdir()))

#!/usr/bin/env python3
"""Vista del editor de texto (como Word) con un Word real (w09): la cuadrícula se ve, no mueve ningún renglón y no se
imprime; el modo de lectura trae todo el texto en columnas, va de pantalla en pantalla con las flechas y Esc sale;
«Nueva ventana» abre el mismo documento.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_vistas_texto.py [capturas]
"""
import base64, os, re, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "office"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402
from prueba_marcas import JS_RENGLONES  # noqa: E402

res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:300]}"), flush=True)


def main():
    arch = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real/archivos/w09-informe.docx"
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    V = "NOM.editorActual.modoApi.vistas"
    try:
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(2500)
        p.evaluate(f"() => {V}.cuadricula(false)"); p.wait_for_timeout(300)
        r0 = p.evaluate(JS_RENGLONES)
        p.evaluate(f"() => {V}.cuadricula(true)"); p.wait_for_timeout(500)
        hay = p.evaluate("() => getComputedStyle(document.querySelector('.wd-pagina'), '::after').backgroundImage")
        anota("la cuadrícula se ve en las hojas", "repeating-linear-gradient" in hay, hay[:120])
        anota("…y no mueve ningún renglón", p.evaluate(JS_RENGLONES) == r0)
        p.emulate_media(media="print"); p.wait_for_timeout(200)
        anota("…ni se imprime", p.evaluate("() => getComputedStyle(document.querySelector('.wd-pagina'), '::after').content") in ("none", "normal"))
        p.emulate_media(media="screen"); p.evaluate(f"() => {V}.cuadricula(false)")
        # modo de lectura
        # (lo que se ve de cada párrafo: con el número de sus campos, sin lo que solo pinta la pantalla)
        parrafos = [re.sub(r"\s+", "", x) for x in p.evaluate("() => [...document.querySelectorAll('.wd-cuerpo > *')].map((e) => { const c = e.cloneNode(true); c.querySelectorAll('.wd-pintado, .wd-guionv, br.wd-rompe, .wd-salto, .wd-saltocol, .wd-cursor, .wd-tab-ctl, .wd-tab-redim').forEach((n) => n.remove()); return c.textContent; })")]
        p.evaluate(f"() => {{ window.__lec = {V}.lectura(); }}"); p.wait_for_timeout(1200)
        tl = re.sub(r"\s+", "", p.evaluate("() => document.querySelector('.wd-lectura-texto').textContent"))
        k, faltan = 0, []
        for x in (q for q in parrafos if q):
            j = tl.find(x, k)
            if j < 0: faltan.append(x[:40])
            else: k = j + len(x)
        anota(f"el modo de lectura trae los {sum(1 for q in parrafos if q)} párrafos, en su orden", not faltan, faltan[:5])
        tot = p.evaluate("() => window.__lec.total()")
        p.keyboard.press("ArrowRight"); p.wait_for_timeout(400)
        anota(f"…en {tot} pantallas, y la flecha pasa a la siguiente", tot > 1 and p.evaluate("() => window.__lec.pagina()") == 1, tot)
        p.keyboard.press("Escape"); p.wait_for_timeout(300)
        anota("…y Esc sale", p.locator(".wd-lectura").count() == 0)
        # nueva ventana
        url = p.evaluate(f"() => {{ let u = null; const o = window.open; window.open = (x) => {{ u = x; return null; }}; {V}.nuevaVentana(); window.open = o; return [u, location.href]; }}")
        anota("«Nueva ventana» abre el mismo documento", url[0] == url[1], url)
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nVistas del texto: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

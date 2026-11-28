#!/usr/bin/env python3
"""Lo que se quitó por verse «de plantilla» (0.24.1, 27-nov) sigue quitado, en el sitio real:
  - el logo sin degradados; la letra del sistema en la interfaz;
  - «Entrar» de alto normal en computadora (antes 25 px) y el título «Entrar con tu cuenta»;
  - en inglés, la «o» del separador y el enlace para bajar la app traducidos;
  - «Para empezar» como lista, no como fila de tarjetas;
  - la tarjeta del espacio con franja de color y la línea en palabras (sin cuadro con la inicial);
  - dentro del espacio: sin «Todo» aparte, «Destacados» se quita tocándolo otra vez, «Propietario» sin corona.
Crea una cuenta desechable (se borra al terminar) y deja capturas.

  [NOM_NAVEGADOR=chromium] ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_aspecto.py SALIDA
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/aspecto"
os.makedirs(SALIDA, exist_ok=True)
bien, mal = [], []
def check(nombre, ok, dato=""):
    (bien if ok else mal).append(nombre)
    print(("BIEN " if ok else "MAL  ") + nombre + (f"  [{dato}]" if dato != "" else ""))

C = N.cuenta(); esp = N.espacio(C["token"], "Laboratorio de redes")
N.documento(C["token"], esp, "TEXTO", "Reporte de la práctica")
errores = []
with sync_playwright() as pw:
    motor = os.environ.get("NOM_NAVEGADOR", "firefox")
    b = getattr(pw, motor).launch()
    # 1. Entrada en computadora (español)
    ctx = b.new_context(viewport={"width": 1366, "height": 768}); p = ctx.new_page()
    p.on("pageerror", lambda e: errores.append(str(e)))
    p.goto(N.B + "/"); p.wait_for_selector("#u", timeout=30000); p.wait_for_timeout(600)
    alto = p.evaluate("() => document.querySelector('button[type=submit]').getBoundingClientRect().height")
    check("«Entrar» mide al menos 36 px de alto en computadora", alto >= 36, round(alto, 1))
    check("título «Entrar con tu cuenta»", "Entrar con tu cuenta" in p.inner_text("form"))
    check("sin «Hola de nuevo»", "Hola de nuevo" not in p.inner_text("body"))
    grad = p.evaluate("() => document.querySelectorAll('linearGradient, radialGradient').length")
    check("el logo no lleva degradados", grad == 0, grad)
    letra = p.evaluate("() => getComputedStyle(document.body).fontFamily")
    check("la interfaz usa la letra del sistema", "Nunito" not in letra, letra[:40])
    p.screenshot(path=f"{SALIDA}/1-entrar.png"); ctx.close()
    # 2. Entrada en inglés
    ctx = b.new_context(viewport={"width": 1366, "height": 768}, locale="en-US"); p = ctx.new_page()
    p.goto(N.B + "/"); p.wait_for_selector("#u", timeout=30000); p.wait_for_timeout(1500)
    seps = p.evaluate("() => [...document.querySelectorAll('.separador')].filter(e => e.offsetParent).map(e => e.textContent.trim())")
    check("en inglés el separador dice «or»", seps and all(s == "or" for s in seps), seps)
    pie = p.evaluate("() => (document.querySelector('.servidor-pie') || {}).textContent || ''")
    check("en inglés el enlace de la app va traducido", "más reciente" not in pie, pie.strip()[:70])
    p.screenshot(path=f"{SALIDA}/2-entrar-en.png"); ctx.close()
    # 3. Inicio y espacio con la cuenta desechable
    for movil in (False, True):
        opc = dict(viewport={"width": 390, "height": 844}, has_touch=True, device_scale_factor=2, **({} if motor == "firefox" else {"is_mobile": True})) if movil else dict(viewport={"width": 1366, "height": 768})
        ctx = b.new_context(**opc); p = ctx.new_page(); p.on("pageerror", lambda e: errores.append(str(e)))
        p.goto(N.B + "/"); p.wait_for_selector("#u", timeout=30000)
        p.fill("#u", C["nombre"]); p.fill("#c", N.CL); p.click("button[type=submit]")
        p.wait_for_selector(".tarjeta-espacio", timeout=30000); p.wait_for_timeout(1200)
        q = "celular" if movil else "computadora"
        if not movil:
            check("«Para empezar» es una lista", p.evaluate("() => !!document.querySelector('section.pasos ol.pasos-lista li') && !document.querySelector('.pasos-item, .pasos-anillo')"))
            check("sin «Primeros pasos» ni «Hazla tuya»", not any(t in p.inner_text("body") for t in ("Primeros pasos", "Hazla tuya", "unos minutos")))
            t = p.evaluate("() => { const a = document.querySelector('.tarjeta-espacio'); return { borde: getComputedStyle(a).borderLeftWidth, inicial: !!a.querySelector('.icono-espacio'), meta: a.querySelector('.meta').textContent.trim().replace(/\\s+/g, ' ') }; }")
            check("tarjeta del espacio con franja de color", t["borde"] == "4px", t["borde"])
            check("tarjeta del espacio sin cuadro con la inicial", not t["inicial"])
            check("la línea de la tarjeta va en palabras", t["meta"].startswith("1 documento · 1 persona · "), t["meta"])
            almacen = p.evaluate("() => { const a = document.querySelector('.almacen-linea'); return a ? !!a.querySelector('.modo-insignia') : null; }")
            check("aviso de almacenamiento sin cuadrito de color", almacen is False, almacen)
        p.screenshot(path=f"{SALIDA}/3-inicio-{q}.png", full_page=True)
        p.click(".tarjeta-espacio"); p.wait_for_selector("[data-docs] .doc-tarjeta", timeout=30000); p.wait_for_timeout(800)
        if not movil:
            check("sin el chip «Todo» repetido", p.evaluate("() => !document.querySelector('[data-carpeta=\"*\"]')"))
            chip = p.evaluate("() => { const c = document.querySelector('.encabezado .chip'); return { t: c.textContent.trim(), ico: !!c.querySelector('svg') }; }")
            check("«Propietario» sin corona", chip["t"] == "Propietario" and not chip["ico"], chip)
            p.click("[data-destacados]"); p.wait_for_timeout(300)
            a1 = p.get_attribute("[data-destacados]", "aria-pressed")
            p.click("[data-destacados]"); p.wait_for_timeout(300)
            a2 = p.get_attribute("[data-destacados]", "aria-pressed")
            n = p.evaluate("() => document.querySelectorAll('[data-docs] .doc-tarjeta').length")
            check("«Destacados» se quita tocándolo otra vez y vuelve a mostrar todo", a1 == "true" and a2 == "false" and n == 1, (a1, a2, n))
        p.screenshot(path=f"{SALIDA}/4-espacio-{q}.png", full_page=True)
        ctx.close()
    b.close()
check("sin errores de JavaScript", not errores, errores[:3])
print(f"\n{len(bien)} bien, {len(mal)} mal")

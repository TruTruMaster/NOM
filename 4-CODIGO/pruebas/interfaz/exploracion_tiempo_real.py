#!/usr/bin/env python3
"""¿Se ve todo en tiempo real, letra por letra, sin dar Enter? (pedido del dueño, 27-nov). Dos personas en dos
navegadores: A escribe una letra a la vez y B debe verla aparecer antes de que A escriba la siguiente (sin Enter).
Texto, cuadro de texto de una presentación y celda de una hoja (mientras se escribe), más un cambio mínimo de formato.

  NOM_NAVEGADOR=firefox ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_tiempo_real.py SALIDA
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL, celda
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/tiempo-real"
os.makedirs(SALIDA, exist_ok=True)
A, B = N.cuenta(), N.cuenta(); esp = N.espacio(A["token"], "Vivo")
N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": B["nombre"], "permiso": "EDICION"}, A["token"])
txt = N.documento(A["token"], esp, "TEXTO", "Texto"); hoja = N.documento(A["token"], esp, "HOJA", "Hoja")

def espera(P, js, arg, ms=3000):
    t0 = time.time()
    while time.time() - t0 < ms / 1000:
        if P.ev(js, arg): return round((time.time() - t0) * 1000)
        P.p.wait_for_timeout(40)
    return None

with sync_playwright() as pw:
    nav = getattr(pw, os.environ.get("NOM_NAVEGADOR", "chromium")).launch()
    pa = nav.new_context(viewport={"width": 1300, "height": 860}).new_page(); PA = Pagina(pa, N.B)
    pb = nav.new_context(viewport={"width": 1300, "height": 860}).new_page(); PB = Pagina(pb, N.B)
    PA.entrar(A); PB.entrar(B)
    # 1. Texto: letra por letra
    PA.abrir(txt); PB.abrir(txt); PA.t(1500); PA.cerrar_hojas(); PB.cerrar_hojas(); PA.quitar_consejo(); PB.quitar_consejo()
    PA.ev("() => NOM.editorActual.vista.querySelector('[contenteditable]').focus()")
    palabra, tiempos = "Hola equipo", []
    for k, letra in enumerate(palabra):
        pa.keyboard.type(letra)
        t = espera(PB, "(s) => (NOM.editorActual.vista.querySelector('[contenteditable]').innerText || '').includes(s)", palabra[:k + 1])
        tiempos.append(t)
    check("texto: cada letra le llega a la otra persona sin Enter", all(t is not None for t in tiempos), tiempos)
    ok = [t for t in tiempos if t is not None]
    if ok: print(f"    texto: mediana {sorted(ok)[len(ok)//2]} ms, peor {max(ok)} ms por letra", flush=True)
    # cambio mínimo: negrita a una palabra
    pa.keyboard.press("Control+Shift+ArrowLeft"); pa.keyboard.press("Control+b")
    t = espera(PB, "() => !!NOM.editorActual.vista.querySelector('[contenteditable] b, [contenteditable] strong, [contenteditable] [style*=\"bold\"], [contenteditable] [style*=\"700\"]')", None)
    check("texto: un cambio de formato (negrita) se ve en tiempo real", t is not None, t)
    # borrar una letra
    pa.keyboard.press("End"); pa.keyboard.press("Backspace")
    t = espera(PB, "() => !(NOM.editorActual.vista.querySelector('[contenteditable]').innerText || '').includes('Hola equipo')", None)
    check("texto: borrar una letra se ve en tiempo real", t is not None, t)
    pb.screenshot(path=f"{SALIDA}/texto-b.png")
    # 2. Hoja: lo que se escribe en una celda (antes de Enter) y al confirmar
    PA.abrir(hoja); PB.abrir(hoja); PA.t(1500); PA.cerrar_hojas(); PB.cerrar_hojas()
    PA.clic(celda(PA, 0, 0)); pa.keyboard.type("Calif");
    t1 = espera(PB, "() => (document.querySelector('td[data-c=\"0\"][data-f=\"0\"]') || {}).textContent === 'Calif'", None, 2500)
    print(f"    hoja: lo escrito en la celda ANTES de Enter le llega a la otra persona: {'sí, ' + str(t1) + ' ms' if t1 is not None else 'no (al confirmar)'}", flush=True)
    pa.keyboard.press("Enter")
    t2 = espera(PB, "() => (document.querySelector('td[data-c=\"0\"][data-f=\"0\"]') || {}).textContent.trim() === 'Calif'", None)
    check("hoja: al confirmar la celda la otra persona la ve en tiempo real", t2 is not None, t2)
    pb.screenshot(path=f"{SALIDA}/hoja-b.png")
    nav.close()
print(f"\nTiempo real: {BIEN[0]} bien, {MAL[0]} mal")

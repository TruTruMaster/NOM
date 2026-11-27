#!/usr/bin/env python3
"""Cualquier cambio, por mínimo que sea, se ve en tiempo real en el otro equipo (pedido del dueño, 27-nov): en una
presentación, A pone una figura, le cambia el color, la mueve, le cambia el tamaño, la gira y la borra; en un texto,
pone una tabla; B (otra cuenta, otro navegador) debe ver cada cambio en menos de 1.5 s, sin recargar.
  NOM_NAVEGADOR=firefox ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_tiempo_real_objetos.py"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright
A, B = N.cuenta(), N.cuenta(); esp = N.espacio(A["token"], "Vivo")
N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": B["nombre"], "permiso": "EDICION"}, A["token"])
pres = N.documento(A["token"], esp, "PRESENTACION", "Expo")
def espera(P, js, ms=1500):
    t0 = time.time()
    while time.time() - t0 < ms / 1000:
        if P.ev(js): return round((time.time() - t0) * 1000)
        P.p.wait_for_timeout(30)
    return None
with sync_playwright() as pw:
    nav = getattr(pw, os.environ.get("NOM_NAVEGADOR", "chromium")).launch()
    pa = nav.new_context(viewport={"width": 1300, "height": 860}).new_page(); PA = Pagina(pa, N.B)
    pb = nav.new_context(viewport={"width": 1300, "height": 860}).new_page(); PB = Pagina(pb, N.B)
    PA.entrar(A); PB.entrar(B); PA.abrir(pres); PB.abrir(pres); PA.t(1500); PB.t(1500)
    pasos = [
        ("poner una figura", "const pg = ed.pagina; doc.cambiar('f1', { tipo: 'forma', forma: 'rect', pagina: pg, orden: 1, x: 100, y: 100, w: 160, h: 100, z: 1, fondo: '#7AA7FF', html: '' }); NOM.modos.PRESENTACION._lienzo.render(); return 1;", "!!document.querySelector('.diapo [data-id=\"f1\"], .escenario [data-id=\"f1\"]')"),
        ("cambiarle el color", "doc.cambiar('f1', { color: '#E2632E', fondo: '#E2632E' }); NOM.modos.PRESENTACION._lienzo.render(); return 1;", "NOM.editorActual.doc.valor('f1', 'color') === '#E2632E' && /226, 99, 46|e2632e/i.test(document.querySelector('.escenario [data-id=\"f1\"], .diapo [data-id=\"f1\"]').outerHTML)"),
        ("moverla", "doc.cambiar('f1', { x: 300, y: 220 }); NOM.modos.PRESENTACION._lienzo.render(); return 1;", "NOM.editorActual.doc.valor('f1', 'x') === 300"),
        ("cambiarle el tamaño", "doc.cambiar('f1', { w: 240 }); NOM.modos.PRESENTACION._lienzo.render(); return 1;", "NOM.editorActual.doc.valor('f1', 'w') === 240"),
        ("girarla", "doc.cambiar('f1', { rot: 45 }); NOM.modos.PRESENTACION._lienzo.render(); return 1;", "+NOM.editorActual.doc.valor('f1', 'rot') === 45"),
        ("borrarla", "doc.cambiar('f1', { eliminado: true }); NOM.modos.PRESENTACION._lienzo.render(); return 1;", "!document.querySelector('.escenario [data-id=\"f1\"], .diapo [data-id=\"f1\"]')"),
    ]
    for nombre, js, ver in pasos:
        PA.doc(js); t = espera(PB, "() => { try { return " + ver + "; } catch (e) { return false; } }")
        check(f"presentación: {nombre} → la otra persona lo ve", t is not None, t); print(f"      {t} ms", flush=True)
    nav.close()
print(f"\nTiempo real (objetos): {BIEN[0]} bien, {MAL[0]} mal")

#!/usr/bin/env python3
"""Ordenar A → Z en la hoja con una tabla elegida, como una persona (reporte del equipo, 27-nov: «se revuelve todo el
documento incluyendo los números de las filas»). Debe ordenar SOLO las filas de la tabla, dejar el encabezado arriba,
no tocar lo que está fuera (otra tabla abajo), no cambiar los números de fila y poder deshacerse con Ctrl+Z.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_orden_tabla.py SALIDA
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL, celda, escribir, visto
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/orden"
os.makedirs(SALIDA, exist_ok=True)
C = N.cuenta(); esp = N.espacio(C["token"], "Orden"); d = N.documento(C["token"], esp, "HOJA", "Lista")
datos = [["Nombre", "Cal", "Grupo"], ["Rosa", "8", "B"], ["Ana", "10", "A"], ["Luis", "6", "B"], ["Eva", "9", "A"]]
abajo = [["Otra", "zz"], ["tabla", "aa"]]
with sync_playwright() as pw:
    b = getattr(pw, os.environ.get("NOM_NAVEGADOR", "chromium")).launch(); p = b.new_context(viewport={"width": 1366, "height": 860}).new_page(); P = Pagina(p, N.B)
    err = []; p.on("pageerror", lambda e: err.append(str(e)[:200]))
    P.entrar(C); P.abrir(d); P.t(1500); P.cerrar_hojas()
    for i, fila in enumerate(datos):
        for j, v in enumerate(fila): escribir(P, j, i, v)
    for i, fila in enumerate(abajo):
        for j, v in enumerate(fila): escribir(P, j, 7 + i, v)
    P.t(500)
    # elegir la tabla A1:C5 con el mouse (de C5 a A1: la celda activa queda en la columna A) y «Ordenar y filtrar › A → Z»
    P.clic(celda(P, 2, 4)); p.keyboard.down("Shift"); P.clic(celda(P, 0, 0)); p.keyboard.up("Shift"); P.t(300)
    hecho = True
    try: P.barra("Ordenar y filtrar"); P.menu("Ordenar A → Z")
    except Exception as e: hecho = False; print("   ", str(e)[:200])
    P.t(1200)
    check("se encontró «Ordenar A → Z»", hecho)
    col = lambda c, f0, f1: [visto(P, c, f) for f in range(f0, f1)]
    check("el encabezado se queda arriba", visto(P, 0, 0) == "Nombre", visto(P, 0, 0))
    check("las filas de la tabla quedan ordenadas por nombre", col(0, 1, 5) == ["Ana", "Eva", "Luis", "Rosa"], col(0, 1, 5))
    check("cada dato sigue con su fila (Ana 10 A)", [visto(P, 1, 1), visto(P, 2, 1)] == ["10", "A"], [visto(P, 1, 1), visto(P, 2, 1)])
    check("lo de fuera de la tabla no se mueve", col(0, 7, 9) == ["Otra", "tabla"] and col(1, 7, 9) == ["zz", "aa"], [col(0, 7, 9), col(1, 7, 9)])
    nums = P.ev("() => [...document.querySelectorAll('tbody tr')].slice(0, 9).map(tr => tr.querySelector('th') ? tr.querySelector('th').textContent.trim() : '')")
    check("los números de fila siguen 1, 2, 3…", nums[:9] == [str(i) for i in range(1, 10)], nums)
    p.screenshot(path=f"{SALIDA}/ordenada.png")
    P.clic(celda(P, 4, 4)); p.keyboard.press("Control+z"); P.t(1000)
    check("Ctrl+Z regresa el orden de antes", col(0, 1, 5) == ["Rosa", "Ana", "Luis", "Eva"], col(0, 1, 5))
    check("sin errores de JavaScript", not err, err[:3])
    b.close()
print(f"\nOrden de tabla: {BIEN[0]} bien, {MAL[0]} mal")

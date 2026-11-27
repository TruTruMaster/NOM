#!/usr/bin/env python3
"""Lo nuevo del 27-nov, usado como una persona en el sitio real (teclado y mouse de verdad):

  1. Ajustes → «Al mantener presionado o clic derecho: Lista» → clic derecho en el texto abre un menú (role=menu) con
     las mismas acciones que la ruedita; Esc lo cierra; se vuelve a «Ruedita».
  2. Ajustes → Atajos: se ve la lista y «Volver a los de siempre».
  3. Ajustes → Idioma «English»: la interfaz cambia a inglés (y el texto del documento no); se vuelve a «Español».
  4. Hoja: BUSCARV con coincidencia aproximada escrito con el teclado da lo mismo que Excel; una fórmula escrita en la
     barra de fórmulas con Enter queda solo en su celda (no se duplica abajo).
  5. Presentación: una figura arrastrada contra la orilla derecha no cambia de tamaño.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_novedades.py SALIDA [--movil]
"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else "/tmp/novedades"
MOVIL = "--movil" in sys.argv
os.makedirs(SALIDA, exist_ok=True)
C = N.cuenta(); esp = N.espacio(C["token"], "Novedades")
texto = N.documento(C["token"], esp, "TEXTO", "Carta")
hoja = N.documento(C["token"], esp, "HOJA", "Calificaciones")
pres = N.documento(C["token"], esp, "PRESENTACION", "Expo")

def ajustes(P):
    P.ev("() => NOM.vistaEspacios.preferencias()"); P.t(1000)
def boton(P, texto, dentro=".velo .hoja"):
    return P.ev("([t, d]) => { const b = [...document.querySelectorAll(d + ' button')].find(x => x.textContent.trim() === t || x.textContent.trim().startsWith(t)); if (b) { b.click(); return true; } return false; }", [texto, dentro])

with sync_playwright() as pw:
    b = pw.chromium.launch()
    opc = dict(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True, device_scale_factor=2) if MOVIL else dict(viewport={"width": 1366, "height": 860})
    ctx = b.new_context(**opc); p = ctx.new_page(); P = Pagina(p, N.B)
    errores = []; p.on("pageerror", lambda e: errores.append(str(e)[:200]))
    P.entrar(C)

    # 1-2. Ajustes: lista y atajos
    p.goto(N.B + "/#/"); P.t(1500); ajustes(P)
    txt = P.ev("() => document.querySelector('.velo .hoja').innerText")
    check("Ajustes muestra «Atajos», «Vibrar al tocar» y «Ruedita · Lista»", "Atajos" in txt and "Lista" in txt and ("Vibrar" in txt or MOVIL is False), txt[:300])
    p.screenshot(path=f"{SALIDA}/1-ajustes.png", full_page=True)
    check("elegir «Lista»", boton(P, "Lista"))
    check("Guardar", boton(P, "Guardar")); P.t(1200)
    P.abrir(texto); P.t(1500)
    P.ev("() => { const c = NOM.editorActual.vista.querySelector('[contenteditable]'); c.focus(); }")
    p.keyboard.type("Hola equipo, este es el informe."); P.t(600)
    caja = P.ev("() => { const r = NOM.editorActual.vista.querySelector('[contenteditable] p, [contenteditable]').getBoundingClientRect(); return [r.left + 40, r.top + 10]; }")
    if not MOVIL: p.mouse.click(caja[0], caja[1], button="right")
    else: P.ev("() => NOM.rueda && NOM.rueda.abrirEn ? 0 : 0"); p.touchscreen.tap(caja[0], caja[1])
    P.t(900)
    menu = P.ev("() => { const m = document.querySelector('[role=menu]'); return m ? [...m.querySelectorAll('[role=menuitem]')].map(x => x.textContent.trim()).slice(0, 12) : null; }")
    p.screenshot(path=f"{SALIDA}/2-lista.png")
    if not MOVIL:
        check("clic derecho abre la lista clásica con acciones", bool(menu) and len(menu) >= 3, menu)
        p.keyboard.press("Escape"); P.t(400)
        check("Esc cierra la lista", P.ev("() => !document.querySelector('[role=menu]')"))
    # volver a ruedita
    ajustes(P); boton(P, "Ruedita"); boton(P, "Guardar"); P.t(800)
    if not MOVIL:
        p.mouse.click(caja[0], caja[1], button="right"); P.t(900)
        check("con «Ruedita» vuelve la ruedita", P.ev("() => !!document.querySelector('.rueda')"))
        p.keyboard.press("Escape"); P.t(300)

    # 3. Idioma
    ajustes(P); check("elegir «English»", boton(P, "English")); P.t(800)
    p.screenshot(path=f"{SALIDA}/3-ajustes-en.png")
    boton(P, "Save") or boton(P, "Guardar"); P.t(1200)
    P.abrir(texto); P.t(1500)
    cuerpo = P.ev("() => document.body.innerText")
    p.screenshot(path=f"{SALIDA}/3-editor-en.png")
    check("en inglés: la barra dice «Insert» y el documento sigue en español", "Insert" in cuerpo and "Hola equipo" in cuerpo, cuerpo[:300])
    check("en inglés: <html lang=en>", P.ev("() => document.documentElement.lang") == "en")
    p.goto(N.B + "/#/"); P.t(1500)
    p.screenshot(path=f"{SALIDA}/3-inicio-en.png")
    ajustes(P); boton(P, "Español"); P.t(500); boton(P, "Save") or boton(P, "Guardar"); P.t(1200)
    check("de vuelta en español", "Inicio" in P.ev("() => document.body.innerText") and P.ev("() => document.documentElement.lang") == "es")

    # 4. Hoja: BUSCARV aproximado (tabla de calificaciones, como la usa un profe) y la barra de fórmulas
    from prueba_funciones import celda, escribir, visto
    P.abrir(hoja); P.t(2000); P.cerrar_hojas()
    datos = [["0", "Reprobado"], ["6", "Suficiente"], ["7", "Bien"], ["8", "Muy bien"], ["9", "Excelente"]]
    for i, (n, t) in enumerate(datos): escribir(P, 0, i, n); escribir(P, 1, i, t)
    casos = [("=BUSCARV(7.5;A1:B5;2)", "Bien"), ("=BUSCARV(9.9;A1:B5;2;VERDADERO)", "Excelente"), ("=BUSCARV(5;A1:B5;2)", "Reprobado"),
             ("=BUSCARV(8;A1:B5;2;FALSO)", "Muy bien"), ("=BUSCARV(8.5;A1:B5;2;FALSO)", "#N/A"), ('=COINCIDIR(6.5;A1:A5)', "2")]
    for k, (f, esp) in enumerate(casos): escribir(P, 3, k, f)
    P.t(600)
    for k, (f, esp) in enumerate(casos): check(f"hoja: {f} = {esp} (como Excel)", visto(P, 3, k) == esp, visto(P, 3, k))
    # barra de fórmulas: clic en F1, escribir en la barra, Enter → solo F1 tiene la fórmula, F2 queda vacía
    P.clic(celda(P, 5, 0)); P.t(200)
    fx = P.ev("() => { const i = document.querySelector('[data-fx]'); if (!i) return null; i.focus(); return true; }")
    if fx:
        p.keyboard.type("=A2+A3"); p.keyboard.press("Enter"); P.t(800)
        check("barra de fórmulas + Enter: la fórmula queda solo en su celda (no se duplica abajo)", visto(P, 5, 0) == "13" and (visto(P, 5, 1) or "") == "", [visto(P, 5, 0), visto(P, 5, 1)])
    else: check("la hoja tiene barra de fórmulas", False)
    p.screenshot(path=f"{SALIDA}/4-hoja.png")

    # 5. Presentación: arrastrar una figura contra la orilla derecha no la encoge
    P.abrir(pres); P.t(2000); P.cerrar_hojas()
    if not MOVIL:
        P.barra("Figuras"); P.t(500)
        P.ev("() => { const b = [...document.querySelectorAll('.velo button, .menu button, [role=menu] button, .hoja button')].find(x => /Rect/i.test(x.textContent + (x.getAttribute('aria-label') || ''))); b && b.click(); }"); P.t(800)
        ids = P.doc("return doc.vivos().filter(b => b.tipo === 'forma').map(b => [b.id || b._id, b.w, b.h]);")
        f = P.ev("() => { const e = document.querySelector('.bl.sel, .bl.seleccionado, .bl[data-id]'); if (!e) return null; const r = e.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2, r.width, r.height]; }")
        if f:
            p.mouse.move(f[0], f[1]); p.mouse.down(); p.mouse.move(f[0] + 1500, f[1], steps=25); p.mouse.up(); P.t(800)
            g = P.ev("() => { const e = document.querySelector('.bl.sel, .bl.seleccionado, .bl[data-id]'); const r = e.getBoundingClientRect(); return [r.width, r.height]; }")
            check("figura arrastrada contra la orilla derecha conserva su tamaño", abs(g[0] - f[2]) < 2 and abs(g[1] - f[3]) < 2, [f[2:], g])
        else: check("se insertó una figura para arrastrar", False, ids)
        p.screenshot(path=f"{SALIDA}/5-figura-orilla.png")
    check("sin errores de JavaScript", not errores, errores[:3])
    b.close()
print(f"\nNovedades: {BIEN[0]} bien, {MAL[0]} mal")

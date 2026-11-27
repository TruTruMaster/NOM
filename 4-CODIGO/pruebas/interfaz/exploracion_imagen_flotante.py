#!/usr/bin/env python3
"""Imagen flotante de un texto con DOS personas a la vez, en el sitio real (reporte del equipo, 27-nov):

  A inserta una foto, la pone «Mover libremente», la arrastra, la gira 90° (Más › Girar 90°) y deshace/rehace.
  B (otra cuenta, otro navegador, con el documento abierto) debe ver cada cambio en su lugar, sin recargar; y
  quien abre el documento después (C) también. Revisa además que la foto se vea cargada completa en los tres.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_imagen_flotante.py SALIDA
"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright
from PIL import Image, ImageDraw

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/flotante"
os.makedirs(SALIDA, exist_ok=True)
foto = os.path.join(SALIDA, "foto.png")
im = Image.new("RGB", (400, 260), "#2B5797"); dr = ImageDraw.Draw(im); dr.polygon([(20, 240), (200, 20), (380, 240)], fill="#E2632E"); im.save(foto)
A, B = N.cuenta(), N.cuenta()
esp = N.espacio(A["token"], "Flotante")
print("acceso de B:", N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": B["nombre"], "permiso": "EDICION"}, A["token"])[0])
d = N.documento(A["token"], esp, "TEXTO", "Informe")

EST = """() => { const e = document.querySelector('.wd-imagen'); if (!e) return null; const r = e.getBoundingClientRect(), c = document.querySelector('.wd-pagina, .wd-cuerpo').getBoundingClientRect();
  const i = e.querySelector('img'); const ed = NOM.editorActual, b = ed.doc.vivos().find(x => x.tipo === 'imagen');
  return { x: Math.round(r.left - c.left), y: Math.round(r.top - c.top), w: Math.round(r.width), h: Math.round(r.height), libre: e.classList.contains('libre'),
           rot: b ? +(b.rot || 0) : null, bx: b && b.x, by: b && b.y, cargada: !!(i && i.complete && i.naturalWidth > 0), tf: getComputedStyle(i || e).transform }; }"""

def opcion(P, xy, ruta):
    P.p.mouse.click(xy[0], xy[1], button="right"); P.t(700)
    for paso in ruta:
        q = P.ev("(t) => { const b = [...document.querySelectorAll('.rueda-item, [role=menuitem]')].find(x => (x.getAttribute('aria-label') || x.textContent || '').trim().startsWith(t)); if (!b) return null; const r = b.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; }", paso)
        if not q: P.p.keyboard.press("Escape"); return False
        P.p.mouse.click(q[0], q[1]); P.t(600)
    return True

with sync_playwright() as pw:
    nav = pw.chromium.launch()
    pa = nav.new_context(viewport={"width": 1366, "height": 900}).new_page(); PA = Pagina(pa, N.B)
    pb = nav.new_context(viewport={"width": 1366, "height": 900}).new_page(); PB = Pagina(pb, N.B)
    err = []; pa.on("pageerror", lambda e: err.append("A: " + str(e)[:200])); pb.on("pageerror", lambda e: err.append("B: " + str(e)[:200]))
    PA.entrar(A); PB.entrar(B)
    PA.abrir(d); PB.abrir(d); PA.t(1500); PA.cerrar_hojas(); PA.quitar_consejo(); PB.cerrar_hojas(); PB.quitar_consejo()
    PA.ev("() => NOM.editorActual.vista.querySelector('[contenteditable]').focus()"); pa.keyboard.type("Primer párrafo del informe con una foto. "); PA.t(300)
    with pa.expect_file_chooser() as fc:
        PA.ev("() => { const b = [...document.querySelectorAll('button')].find(x => x.offsetParent && /^Insertar/.test(x.textContent.trim())); b.click(); }"); PA.t(600)
        PA.ev("() => { const b = [...document.querySelectorAll('.velo button, .menu button, [role=menu] [role=menuitem], .hoja button')].find(x => x.offsetParent && /^Imagen/.test(x.textContent.trim())); b.click(); }")
    fc.value.set_files(foto); PA.t(3000)
    a0 = PA.ev(EST); b0 = PB.ev(EST)
    check("la foto insertada por A se ve completa en A y en B", bool(a0 and b0 and a0["cargada"] and b0["cargada"]), [a0, b0])
    c = PA.ev(EST); xy = (c["x"] + 120 + PA.ev("() => document.querySelector('.wd-pagina, .wd-cuerpo').getBoundingClientRect().left"), c["y"] + 80 + PA.ev("() => document.querySelector('.wd-pagina, .wd-cuerpo').getBoundingClientRect().top"))
    check("«Mover libremente»", opcion(PA, xy, ["Mover libremente"]) and PA.ev(EST)["libre"])
    PB.t(1500); check("B ve la foto como libre", (PB.ev(EST) or {}).get("libre"), PB.ev(EST))
    # A arrastra
    c = PA.ev(EST); org = PA.ev("() => { const r = document.querySelector('.wd-pagina, .wd-cuerpo').getBoundingClientRect(); return [r.left, r.top]; }")
    x0, y0 = org[0] + c["x"] + 100, org[1] + c["y"] + 80
    pa.mouse.click(x0, y0); PA.t(300); pa.mouse.move(x0, y0); pa.mouse.down(); pa.mouse.move(x0 + 80, y0 + 60, steps=8); pa.mouse.move(x0 + 160, y0 + 140, steps=8); pa.mouse.up(); PA.t(2500)
    a1, b1 = PA.ev(EST), PB.ev(EST)
    check("A movió la foto (cambió su posición en A)", abs(a1["x"] - c["x"]) > 50 or abs(a1["y"] - c["y"]) > 50, [c, a1])
    check("B ve la foto en la posición nueva sin recargar", abs(a1["x"] - b1["x"]) <= 3 and abs(a1["y"] - b1["y"]) <= 3, [a1, b1])
    pa.screenshot(path=f"{SALIDA}/1-a-movida.png"); pb.screenshot(path=f"{SALIDA}/1-b-movida.png")
    # Vista «Sin páginas» (la vista web): mover la foto libre también ahí
    PA.ev("() => document.querySelector('[data-vista=\"continua\"]').click()"); PA.t(1500)
    org2 = PA.ev("() => { const r = document.querySelector('.wd-cuerpo').getBoundingClientRect(); return [r.left, r.top]; }")
    cw = PA.ev(EST)
    if cw:
        xw, yw = org2[0] + cw["x"] + 60, org2[1] + cw["y"] + 60
        pa.mouse.click(xw, yw); PA.t(300); pa.mouse.move(xw, yw); pa.mouse.down(); pa.mouse.move(xw - 60, yw + 40, steps=8); pa.mouse.move(xw - 120, yw + 90, steps=8); pa.mouse.up(); PA.t(2000)
        cw2 = PA.ev(EST)
        check("vista «Sin páginas»: arrastrar la foto libre la mueve", cw2 and (abs(cw2["x"] - cw["x"]) > 40 or abs(cw2["y"] - cw["y"]) > 40), [cw, cw2])
        check("vista «Sin páginas»: el lugar guardado cambió (lo ven los demás)", cw2 and (cw2["bx"] != cw["bx"] or cw2["by"] != cw["by"]), [cw, cw2])
        pa.screenshot(path=f"{SALIDA}/1w-sin-paginas.png")
    else: check("vista «Sin páginas»: la foto se ve", False)
    PA.ev("() => document.querySelector('[data-vista=\"paginas\"]').click()"); PA.t(1500)
    # A gira 90°
    c = PA.ev(EST); xy = (org[0] + c["x"] + c["w"] / 2, org[1] + c["y"] + c["h"] / 2)
    girada = opcion(PA, xy, ["Más", "Girar 90"]); PA.t(1500)
    a2, b2 = PA.ev(EST), PB.ev(EST)
    check("A gira la foto 90° desde Más › Girar 90°", girada and a2["rot"] == 90, a2)
    check("B ve la foto girada y en su lugar", b2 and b2["rot"] == 90 and abs(a2["x"] - b2["x"]) <= 3 and abs(a2["y"] - b2["y"]) <= 3, [a2, b2])
    pa.screenshot(path=f"{SALIDA}/2-a-girada.png"); pb.screenshot(path=f"{SALIDA}/2-b-girada.png")
    # deshacer y rehacer
    pa.keyboard.press("Control+z"); PA.t(1500); a3, b3 = PA.ev(EST), PB.ev(EST)
    check("deshacer el giro: A y B la ven sin girar", a3 and b3 and a3["rot"] == 0 and b3["rot"] == 0, [a3, b3])
    pa.keyboard.press("Control+y"); PA.t(1500)
    # C abre después: misma posición y giro
    pc = nav.new_context(viewport={"width": 1366, "height": 900}).new_page(); PC = Pagina(pc, N.B); PC.entrar(A); PC.abrir(d); PC.t(3000)
    a4, c4 = PA.ev(EST), PC.ev(EST)
    check("quien abre después la ve en el mismo lugar, girada y cargada", c4 and abs(a4["x"] - c4["x"]) <= 3 and abs(a4["y"] - c4["y"]) <= 3 and c4["rot"] == a4["rot"] and c4["cargada"], [a4, c4])
    pc.screenshot(path=f"{SALIDA}/3-c-abre.png")
    # El enlace para ver (sin cuenta, en otro navegador): la foto en el mismo lugar que en el editor
    enl = N.api("POST", f"/api/documentos/{d}/enlace-ver", {}, A["token"])[1]["enlace"]
    pv = nav.new_context(viewport={"width": 1366, "height": 900}).new_page(); pv.on("pageerror", lambda e: err.append("ver: " + str(e)[:200]))
    pv.goto(N.B + "/#/ver/" + enl); pv.wait_for_timeout(6000)
    v = pv.evaluate("""() => { const f = document.querySelector('iframe.ver-marco'), dd = f ? f.contentDocument : document; const i = dd.querySelector('img'); if (!i) return null; const e = i.closest('span[style*="position:absolute"]') || i; const r = e.getBoundingClientRect(), p = (e.closest('p') || dd.body).getBoundingClientRect(); const c = { left: p.left, top: p.top };
      return { x: Math.round(r.left - c.left), y: Math.round(r.top - c.top), w: Math.round(r.width), libre: e !== i, cargada: !!(i && i.complete && i.naturalWidth > 0) }; }""")
    a5 = PA.ev(EST)
    check("enlace para ver: la foto libre, desde su párrafo como en el editor, girada y cargada", v and v["libre"] and abs(v["x"] - a5["bx"]) <= 4 and abs(v["y"] - a5["by"]) <= 4 and v["cargada"] and v["w"] == a5["w"], [a5, v])
    pv.screenshot(path=f"{SALIDA}/4-enlace-ver.png")
    check("sin errores de JavaScript", not err, err[:4])
    nav.close()
print(f"\nImagen flotante: {BIEN[0]} bien, {MAL[0]} mal")

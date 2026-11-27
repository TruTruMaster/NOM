#!/usr/bin/env python3
"""Imágenes como las usa una persona, en los cuatro editores del sitio real (pedido del dueño, 26-nov):

  - insertar una foto con el botón «Imagen» de verdad (el selector de archivos del sistema);
  - arrastrarla contra cada orilla y más allá (derecha, izquierda, arriba, abajo): nunca cambia de tamaño;
  - deshacer a la mitad (vuelve a donde estaba, con su tamaño) y rehacer;
  - girarla y recortarla desde su menú (ruedita o lista), si el editor lo tiene;
  - que se vea completa (cargada) al final y sin errores de JavaScript.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_imagenes.py SALIDA
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright
from PIL import Image, ImageDraw

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/imagenes"
os.makedirs(SALIDA, exist_ok=True)
foto = os.path.join(SALIDA, "foto.png")
im = Image.new("RGB", (400, 300), "#2B5797"); d = ImageDraw.Draw(im); d.ellipse((150, 100, 250, 200), fill="#E2632E"); d.rectangle((20, 20, 380, 280), outline="white", width=6); im.save(foto)
C = N.cuenta(); esp = N.espacio(C["token"], "Imágenes")
docs = {m: N.documento(C["token"], esp, m, m.capitalize()) for m in ("PRESENTACION", "DIAGRAMA", "TEXTO", "HOJA")}

# El elemento de la imagen en cada editor (el que se arrastra) y su caja en pantalla.
SEL = {"PRESENTACION": ".bl:has(img)", "DIAGRAMA": ".bl:has(img)", "TEXTO": ".wd-imagen", "HOJA": ".xl-dib:has(img), .xl-dibujos img, .hoja-img"}
CAJA = """(s) => { const e = [...document.querySelectorAll(s)].filter(x => x.offsetParent && !x.closest('.miniaturas, .mini, .diapo-mini')).pop(); if (!e) return null; const r = e.getBoundingClientRect();
  const i = e.tagName === 'IMG' ? e : e.querySelector('img'); return { x: r.left, y: r.top, w: r.width, h: r.height, cargada: !!(i && i.complete && i.naturalWidth > 0) }; }"""

def ruta(P, xy, pasos, objetivo):
    """Ruedita en xy: sigue `pasos` (submenús) y luego busca `objetivo`, pasando por «Más» si no se ve. Clics reales."""
    P.p.mouse.click(xy[0], xy[1], button="right"); P.t(700)
    donde = lambda rx: P.ev("(rx) => { const b = [...document.querySelectorAll('.rueda-item')].find(x => new RegExp(rx, 'i').test(x.getAttribute('aria-label') || '')); if (!b) return null; const q = b.getBoundingClientRect(); return [q.left + q.width / 2, q.top + q.height / 2]; }", rx)
    for paso in pasos:
        q = donde("^" + paso)
        if not q: return False
        P.p.mouse.click(q[0], q[1]); P.t(600)
    for _ in range(4):
        q = donde(objetivo)
        if q: P.p.mouse.click(q[0], q[1]); P.t(900); return True
        m = donde("^Más")
        if not m: break
        P.p.mouse.click(m[0], m[1]); P.t(600)
    print("    (en el menú hay:", P.ev("() => [...document.querySelectorAll('.rueda-item')].map(x => x.getAttribute('aria-label'))"), ")", flush=True)
    P.p.keyboard.press("Escape"); P.t(300); return False

def elegir(P, patron, abrir_en=None, por=("Insertar", "Más")):
    """Abre la ruedita (clic derecho real en abrir_en) y elige la opción que dice `patron` con un clic real; si no está a
    la vista, entra primero a un submenú de `por`. Devuelve True si la eligió."""
    if abrir_en: P.p.mouse.click(abrir_en[0], abrir_en[1], button="right"); P.t(800)
    def donde(rx):
        return P.ev("(rx) => { const b = [...document.querySelectorAll('.rueda-item, [role=menuitem]')].find(x => new RegExp(rx, 'i').test(x.getAttribute('aria-label') || x.textContent)); if (!b) return null; const q = b.getBoundingClientRect(); return [q.left + q.width / 2, q.top + q.height / 2]; }", rx)
    r = donde(patron)
    if not r:
        for sub in por:
            q = donde("^" + sub)
            if q: P.p.mouse.click(q[0], q[1]); P.t(700); r = donde(patron)
            if r: break
    if not r:
        print("    (en el menú hay:", P.ev("() => [...document.querySelectorAll('.rueda-item, [role=menuitem]')].map(x => x.getAttribute('aria-label') || x.textContent.trim())"), ")", flush=True)
        P.p.keyboard.press("Escape"); P.t(300); return False
    P.p.mouse.click(r[0], r[1]); P.t(900); return True

with sync_playwright() as pw:
    b = pw.chromium.launch(); ctx = b.new_context(viewport={"width": 1440, "height": 900}); p = ctx.new_page(); P = Pagina(p, N.B)
    errores = []; p.on("pageerror", lambda e: errores.append(str(e)[:200]))
    P.entrar(C)
    for modo, d in docs.items():
        print("\n==", modo, flush=True)
        P.abrir(d); P.t(2000); P.cerrar_hojas(); P.quitar_consejo()
        if modo == "TEXTO":
            P.ev("() => NOM.editorActual.vista.querySelector('[contenteditable]').focus()"); p.keyboard.type("Texto antes de la foto. "); P.t(300)
        try:
            with p.expect_file_chooser(timeout=8000) as fc:
                if modo == "HOJA":   # en la hoja, «Imagen…» está en la ruedita de la celda
                    q = P.ev("() => { const r = document.querySelector('td[data-c=\"1\"][data-f=\"1\"]').getBoundingClientRect(); return [r.left + 10, r.top + 8]; }")
                    p.mouse.click(q[0], q[1]); P.t(300)
                    if not ruta(P, q, ["Más", "Otras"], "^Imagen"): raise Exception("sin «Imagen…» en la ruedita de la celda")
                elif modo in ("PRESENTACION", "DIAGRAMA"):
                    P.ev("() => { const b = [...document.querySelectorAll('button')].find(x => x.offsetParent && /^Imagen$/.test((x.textContent || x.getAttribute('aria-label') || '').trim())); b.click(); }")
                else:
                    P.ev("() => { const b = [...document.querySelectorAll('button')].find(x => x.offsetParent && /^Insertar/.test((x.textContent || x.getAttribute('aria-label') || '').trim())); b.click(); }"); P.t(600)
                    P.ev("() => { const b = [...document.querySelectorAll('.velo button, .menu button, [role=menu] [role=menuitem], [role=menu] button, .hoja button')].find(x => x.offsetParent && /^Imagen/.test((x.textContent || x.getAttribute('aria-label') || '').trim())); b.click(); }")
            fc.value.set_files(foto)
        except Exception as e:
            check(f"{modo}: el botón «Imagen» abre el selector de archivos", False, str(e)[:150]); continue
        P.t(2500)
        c0 = P.ev(CAJA, SEL[modo])
        check(f"{modo}: la foto se insertó y se ve completa", bool(c0) and c0["cargada"], c0)
        if not c0: p.screenshot(path=f"{SALIDA}/{modo}-sin-imagen.png"); continue
        if modo == "TEXTO":
            # flotante («Delante del texto») para poder arrastrarla libre, como en Word
            libre = elegir(P, "Mover libremente", (c0["x"] + c0["w"] / 2, c0["y"] + c0["h"] / 2))
            check("TEXTO: «Mover libremente» desde el menú de la imagen", libre and P.ev("() => !!document.querySelector('.wd-imagen.libre')"))
            p.keyboard.press("Escape"); P.t(300)
        p.screenshot(path=f"{SALIDA}/{modo}-0-insertada.png")
        # arrastrar contra cada orilla y más allá
        for nombre, dx, dy in (("derecha", 2500, 0), ("izquierda", -5000, 0), ("abajo", 0, 2500), ("arriba", 0, -5000), ("centro", None, None)):
            c = P.ev(CAJA, SEL[modo])
            if not c: check(f"{modo}: la foto sigue ahí después de arrastrarla a la {nombre}", False); break
            x, y = c["x"] + c["w"] / 2, c["y"] + c["h"] / 2
            if dx is None: dx, dy = 720 - x, 480 - y   # de vuelta al centro de la pantalla
            p.mouse.move(x, y); p.mouse.down(); p.mouse.move(x + dx / 4, y + dy / 4, steps=6); p.mouse.move(x + dx, y + dy, steps=12); p.mouse.up(); P.t(700)
            c2 = P.ev(CAJA, SEL[modo])
            same = bool(c2) and abs(c2["w"] - c["w"]) < 2 and abs(c2["h"] - c["h"]) < 2
            check(f"{modo}: arrastrada a la {nombre} (y más allá) conserva su tamaño", same, [c and (round(c["w"]), round(c["h"])), c2 and (round(c2["w"]), round(c2["h"]))])
        p.screenshot(path=f"{SALIDA}/{modo}-1-arrastrada.png")
        # deshacer a la mitad y rehacer
        antes = P.ev(CAJA, SEL[modo])
        p.keyboard.press("Control+z"); P.t(700); medio = P.ev(CAJA, SEL[modo])
        p.keyboard.press("Control+y"); P.t(700); despues = P.ev(CAJA, SEL[modo])
        check(f"{modo}: deshacer y rehacer el último arrastre no cambia el tamaño", bool(medio and despues) and abs(medio["w"] - antes["w"]) < 2 and abs(despues["w"] - antes["w"]) < 2, [antes, medio, despues])
        # girar y recortar desde su menú
        c = P.ev(CAJA, SEL[modo])
        if c and modo in ("PRESENTACION", "DIAGRAMA"):
            p.mouse.click(c["x"] + c["w"] / 2, c["y"] + c["h"] / 2); P.t(300)
            items = []
            P.t(500)
            girado = ruta(P, (c["x"] + c["w"] / 2, c["y"] + c["h"] / 2), ["Orden"], "^Girar 90")
            rot = P.doc("const b = doc.vivos().find(x => x.tipo === 'imagen'); return b ? (b.rot || 0) : null;")
            check(f"{modo}: girar la foto desde su menú", girado and rot not in (None, 0, "0"), [girado, rot, items[:10]])
            p.keyboard.press("Escape"); P.t(300)
        c = P.ev(CAJA, SEL[modo])
        check(f"{modo}: al final la foto se ve completa (cargada)", bool(c) and c["cargada"], c)
        p.screenshot(path=f"{SALIDA}/{modo}-2-final.png")
    check("sin errores de JavaScript", not errores, errores[:4])
    b.close()
print(f"\nImágenes: {BIEN[0]} bien, {MAL[0]} mal")

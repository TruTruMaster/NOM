#!/usr/bin/env python3
"""¿Desaparece algún botón al elegir opciones de accesibilidad? (reporte del dueño, 26-nov)

Para cada combinación de letra (normal, grande, muy grande), tema (claro, oscuro), «reducir movimiento» y colores
forzados (alto contraste del sistema), en computadora (1366 × 768) y celular (390 × 844): abre el inicio y los cuatro
editores y junta cada botón que se ve con su nombre. Un botón está MAL si existe con la letra normal y con la opción
(a) ya no se ve, (b) queda fuera de la pantalla, (c) queda cortado por su barra (overflow) o (d) su texto/ícono tiene
el mismo color que su fondo. Deja capturas de cada pantalla con cada opción.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_botones_accesibilidad.py SALIDA
"""
import os, sys, json, itertools
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/botones"
os.makedirs(SALIDA, exist_ok=True)
C = N.cuenta(); esp = N.espacio(C["token"], "Accesible")
docs = {m: N.documento(C["token"], esp, m, m.capitalize()) for m in ("TEXTO", "HOJA", "PRESENTACION", "DIAGRAMA")}

BOTONES = r"""() => {
  const vw = innerWidth, vh = innerHeight, out = [];
  const nombre = (b) => (b.getAttribute('aria-label') || b.title || b.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 40);
  const color = (c) => { const m = c.match(/[\d.]+/g); return m ? m.slice(0, 3).map(Number) : [0, 0, 0]; };
  const fondoDe = (el) => { for (let x = el; x; x = x.parentElement) { const c = getComputedStyle(x).backgroundColor; if (c && !/rgba\(0, 0, 0, 0\)|transparent/.test(c)) return color(c); } return [255, 255, 255]; };
  for (const b of document.querySelectorAll('button, [role=button], a.btn, select')) {
    if (b.closest('.velo, .hoja, .rueda, [hidden], .paleta-velo, .menu')) continue;
    const n = nombre(b); if (!n) continue;
    const r = b.getBoundingClientRect(), cs = getComputedStyle(b);
    const visible = b.offsetParent !== null && r.width > 1 && r.height > 1 && cs.visibility !== 'hidden' && +cs.opacity > 0.05;
    let cortado = false;
    for (let x = b.parentElement; x && x !== document.body; x = x.parentElement) {
      const s = getComputedStyle(x);
      if (/(hidden|clip|auto|scroll)/.test(s.overflowX + s.overflowY)) { const q = x.getBoundingClientRect(); if (r.right > q.right + 2 || r.left < q.left - 2 || r.bottom > q.bottom + 2 || r.top < q.top - 2) { cortado = /(hidden|clip)/.test(s.overflowX + s.overflowY) ? 'oculto' : 'desplazar'; break; } }
    }
    const fuera = r.right > vw + 1 || r.left < -1 || r.bottom > vh + 1 || r.top < -1;
    const fg = color(cs.color), bg = fondoDe(b), dif = Math.max(...fg.map((v, i) => Math.abs(v - bg[i])));
    out.push({ n, visible, fuera, cortado, invisibleColor: visible && dif < 30 && !b.querySelector('svg, img'), x: Math.round(r.left), w: Math.round(r.width) });
  }
  return out;
}"""

def pantallas(P):
    yield "inicio", N.B + "/#/"
    for m, d in docs.items(): yield m.lower(), N.B + "/#/doc/" + d

with sync_playwright() as pw:
    b = pw.chromium.launch()
    for disp, vp in (("pc", {"width": 1366, "height": 768}), ("cel", {"width": 390, "height": 844})):
        base = {}
        combos = [("normal", "claro", False, "none")] + [(l, t, mv, fc) for l, t, mv, fc in itertools.product(("grande", "muy-grande"), ("claro", "oscuro"), (False, True), ("none",))] + [("normal", "claro", False, "active"), ("muy-grande", "oscuro", True, "active")]
        for letra, tema, mov, forz in combos:
            opc = dict(viewport=vp, forced_colors=forz)
            if disp == "cel": opc.update(is_mobile=True, has_touch=True, device_scale_factor=2)
            ctx = b.new_context(**opc); p = ctx.new_page(); P = Pagina(p, N.B); P.entrar(C)
            P.ev(f"() => {{ const pr = Object.assign(NOM.leer('nom.pref', {{}}), {{ tema: '{tema}', movimientoReducido: {str(mov).lower()} }}); NOM.aplicarPreferencias(pr); NOM.guardar('nom.letra', '{letra}'); }}")
            clave = f"{disp}-{letra}-{tema}-{'quieto' if mov else 'mov'}-{forz}"
            for nombre, url in pantallas(P):
                p.goto(url); P.t(300); p.reload(); P.t(2500); P.cerrar_hojas()
                bs = P.ev(BOTONES)
                p.screenshot(path=f"{SALIDA}/{clave}-{nombre}.png")
                if letra == "normal" and tema == "claro" and not mov and forz == "none":
                    base[nombre] = {x["n"]: x for x in bs if x["visible"] and not x["fuera"] and not x["cortado"]}
                    continue
                ahora = {}
                for x in bs: ahora.setdefault(x["n"], x)
                perdidos = []
                for n0 in base.get(nombre, {}):
                    x = ahora.get(n0)
                    if not x or not x["visible"]: perdidos.append(n0 + " (no se ve)")
                    elif x["fuera"]: perdidos.append(n0 + " (fuera de la pantalla)")
                    elif x["cortado"] == "oculto": perdidos.append(n0 + " (cortado por su barra)")
                    elif x["invisibleColor"]: perdidos.append(n0 + " (mismo color que el fondo)")
                check(f"{clave} · {nombre}: ningún botón desaparece", not perdidos, perdidos[:12])
            ctx.close()
    b.close()
print(f"\nBotones con accesibilidad: {BIEN[0]} bien, {MAL[0]} mal")

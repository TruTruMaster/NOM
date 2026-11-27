#!/usr/bin/env python3
"""¿Algo fuera de su lugar? (reporte del equipo, 27-nov). En computadora (1366 × 768) y celular (390 × 844), con
contenido de verdad: inicio, espacio, Ajustes, invitar, chat, los cuatro editores con contenido y sus menús. En cada
pantalla busca:
  - elementos visibles que se salen de la pantalla a lo ancho (barra horizontal sin querer);
  - botones o textos que se salen de su caja (texto cortado o encimado sobre otro botón);
  - botones que se enciman entre sí.
Deja una captura por pantalla y la lista de lo encontrado.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_layout.py SALIDA
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL, escribir
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/layout"
os.makedirs(SALIDA, exist_ok=True)
C = N.cuenta(); esp = N.espacio(C["token"], "Equipo de la clase de Sistemas Distribuidos")
docs = {m: N.documento(C["token"], esp, m, n) for m, n in (("TEXTO", "Informe final de la práctica de laboratorio número tres"), ("HOJA", "Calificaciones del segundo parcial"), ("PRESENTACION", "Exposición del proyecto"), ("DIAGRAMA", "Mapa de la arquitectura"))}

REVISAR = r"""() => {
  const vw = innerWidth, out = [];
  const vis = (e) => { const r = e.getBoundingClientRect(), s = getComputedStyle(e); return r.width > 2 && r.height > 2 && s.visibility !== 'hidden' && s.display !== 'none' && +s.opacity > 0.05 && e.offsetParent !== null; };
  const nombre = (e) => ((e.getAttribute('aria-label') || e.textContent || e.className || e.tagName) + '').trim().replace(/\s+/g, ' ').slice(0, 40);
  const modal = [...document.querySelectorAll('.velo')].filter((v) => v.offsetParent !== null || getComputedStyle(v).position === 'fixed').pop();
  const enDesplazable = (e) => { for (let x = e.parentElement; x && x !== document.body; x = x.parentElement) { const s = getComputedStyle(x); if (/(auto|scroll)/.test(s.overflowX)) return true; } return false; };
  const fuera = (e) => (modal && !modal.contains(e)) || !!e.closest('.wd-hoja, .wd-pagina, .wd-cuerpo, .rejilla-hoja, .diapo, .escenario, .mundo, .lienzo, [data-mundo], .miniaturas, .ver-marco, .xl, table.hoja, .mini-cont, [data-rejilla]');
  // 1. se sale de la pantalla a lo ancho
  for (const e of document.querySelectorAll('button, a, input, select, h1, h2, h3, p, label, .chip, .btn')) {
    if (!vis(e) || fuera(e)) continue;
    const r = e.getBoundingClientRect();
    if ((r.right > vw + 2 || r.left < -2) && !enDesplazable(e)) out.push('se sale de la pantalla: ' + nombre(e) + ` (${Math.round(r.left)}→${Math.round(r.right)} de ${vw})`);
  }
  // 2. texto que se sale de su botón (cortado o encimado)
  for (const b of document.querySelectorAll('button, .btn, .chip')) {
    if (!vis(b) || fuera(b)) continue;
    const s = getComputedStyle(b);
    if (b.scrollWidth > b.clientWidth + 3 && s.overflowX !== 'auto' && s.textOverflow !== 'ellipsis' && s.whiteSpace !== 'normal') out.push('texto cortado en el botón: ' + nombre(b) + ` (${b.scrollWidth} > ${b.clientWidth})`);
  }
  // 3. botones encimados
  const bs = [...document.querySelectorAll('button')].filter((b) => vis(b) && !fuera(b) && !b.closest('.rueda, .menu, [role=menu]'));
  const rs = bs.map((b) => b.getBoundingClientRect());
  for (let i = 0; i < bs.length; i++) for (let j = i + 1; j < bs.length; j++) {
    if (bs[i].contains(bs[j]) || bs[j].contains(bs[i])) continue;
    const a = rs[i], c = rs[j], ix = Math.min(a.right, c.right) - Math.max(a.left, c.left), iy = Math.min(a.bottom, c.bottom) - Math.max(a.top, c.top);
    if (ix > 6 && iy > 6 && ix * iy > 0.25 * Math.min(a.width * a.height, c.width * c.height)) out.push('botones encimados: ' + nombre(bs[i]) + ' / ' + nombre(bs[j]));
  }
  if (document.documentElement.scrollWidth > vw + 2) out.push('la página tiene barra horizontal: ' + document.documentElement.scrollWidth + ' > ' + vw);
  return [...new Set(out)].slice(0, 15);
}"""

with sync_playwright() as pw:
    b = pw.chromium.launch()
    for disp, opc in (("pc", dict(viewport={"width": 1366, "height": 768})), ("cel", dict(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True, device_scale_factor=2))):
        ctx = b.new_context(**opc); p = ctx.new_page(); P = Pagina(p, N.B); err = []; p.on("pageerror", lambda e: err.append(str(e)[:150]))
        P.entrar(C)
        def mirar(nombre):
            P.t(900); p.screenshot(path=f"{SALIDA}/{disp}-{nombre}.png")
            r = P.ev(REVISAR); check(f"{disp} · {nombre}: nada fuera de su lugar", not r, r)
        p.goto(N.B + "/#/"); P.t(2000); P.cerrar_hojas(); mirar("inicio")
        p.goto(N.B + "/#/espacio/" + esp); P.t(2000); P.cerrar_hojas(); mirar("espacio")
        P.ev("() => NOM.vistaEspacios.preferencias()"); mirar("ajustes"); P.cerrar_hojas()
        P.ev("() => { const b = [...document.querySelectorAll('button')].find(x => x.offsetParent && /^Invitar/.test((x.getAttribute('aria-label') || x.textContent).trim())); b && b.click(); }"); mirar("invitar"); P.cerrar_hojas()
        P.ev("() => { const b = [...document.querySelectorAll('button')].find(x => x.offsetParent && /^Chat/.test((x.getAttribute('aria-label') || x.textContent).trim())); b && b.click(); }"); mirar("chat"); P.cerrar_hojas()
        for m, d in docs.items():
            P.abrir(d); P.t(1500); P.cerrar_hojas(); P.quitar_consejo()
            if m == "TEXTO":
                P.ev("() => NOM.editorActual.vista.querySelector('[contenteditable]').focus()"); p.keyboard.type("Introducción. " * 20); P.t(500)
            if m == "HOJA":
                for i, fila in enumerate([["Nombre completo del alumno", "Calificación"], ["María Fernanda", "9.5"], ["José Luis", "8"]]):
                    for j, v in enumerate(fila): escribir(P, j, i, v)
            mirar(m.lower())
            # los menús de la barra de arriba (el «⋮» y «Más»)
            P.ev("() => { const b = [...document.querySelectorAll('button')].find(x => x.offsetParent && /Más opciones|^Más$/.test((x.getAttribute('aria-label') || x.textContent).trim())); b && b.click(); }")
            mirar(m.lower() + "-menu"); P.cerrar_hojas(); p.keyboard.press("Escape")
        check(f"{disp}: sin errores de JavaScript", not err, err[:3])
        ctx.close()
    b.close()
print(f"\nLayout: {BIEN[0]} bien, {MAL[0]} mal")

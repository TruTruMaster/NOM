#!/usr/bin/env python3
"""Revisión automática de accesibilidad con axe-core (reglas WCAG 2.0, 2.1 y 2.2, niveles A y AA) en las pantallas
principales, en computadora y en celular: entrada, avisos legales, ayuda, inicio, un espacio, los cuatro editores,
Ajustes, la paleta de comandos y el tema oscuro. Cada pantalla «bien» = ninguna regla falla.

Lo automático encuentra solo una parte de las barreras (nombres, contraste, estructura); lo demás se prueba a mano
con un lector de pantalla (ver docs/PUBLICACION-Y-CUMPLIMIENTO.md, sección 5).

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_axe.py [salida.json]

axe-core lo baja pruebas/preparar.sh (npm) a ~/.cache/nom-pruebas/axe; otra ruta con AXE=/ruta/axe.min.js.
La política de seguridad de contenido de NOM no deja inyectar guiones (bien): la prueba abre un contexto que la ignora.
"""
import json, os, sys
from pathlib import Path
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", os.path.expanduser("~/.cache/nom-pruebas/navegadores"))
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import nomprueba as N  # noqa: E402

AXE = os.environ.get("AXE", os.path.expanduser("~/.cache/nom-pruebas/axe/node_modules/axe-core/axe.min.js"))
TAGS = ["wcag2a", "wcag2aa", "wcag21a", "wcag21aa", "wcag22aa"]
res = {}
cuenta_ok = []


def revisar(p, nombre):
    p.add_script_tag(content=open(AXE, encoding="utf-8").read())
    r = p.evaluate("async t => { const r = await axe.run(document, { runOnly: { type: 'tag', values: t } });"
                   " return r.violations.map(v => ({ id: v.id, impacto: v.impact, ayuda: v.help, n: v.nodes.length,"
                   " ejemplos: v.nodes.slice(0, 3).map(x => x.target.join(' ') + ' :: ' + (x.failureSummary || '').slice(0, 200)) })) }", TAGS)
    res[nombre] = r
    cuenta_ok.append(not r)
    if r: print(f"  MAL  {nombre}: " + "; ".join(f"{v['id']} ({v['impacto']}, {v['n']}): {v['ejemplos'][0][:160]}" for v in r), flush=True)
    else: print(f"  bien {nombre}", flush=True)


def main():
    if not os.path.exists(AXE): sys.exit(f"Falta axe-core en {AXE}: correr pruebas/preparar.sh")
    if not N.esperar_servidor(): sys.exit("El servidor no contesta en " + N.B)
    c = N.cuenta(); esp = N.espacio(c["token"], "Revisión de accesibilidad")
    docs = {m: N.documento(c["token"], esp, m, "Prueba " + m.lower()) for m in ("TEXTO", "HOJA", "PRESENTACION", "DIAGRAMA")}
    for movil in (False, True):
        tam = "celular" if movil else "computadora"
        print(f"== {tam}", flush=True)
        nav = N.Navegador(movil=movil)
        opc = dict(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True, device_scale_factor=2) if movil else dict(viewport={"width": 1400, "height": 900})
        nav.ctx.close(); nav.ctx = nav.b.new_context(bypass_csp=True, **opc); nav.p = nav.ctx.new_page(); p = nav.p
        p.on("pageerror", lambda e: nav.errores.append(str(e)))
        for ruta in ("/", "/legal/privacidad", "/legal/accesibilidad", "/ayuda/"):
            p.goto(N.B + ruta); p.wait_for_timeout(1500); revisar(p, f"{tam}: {ruta}")
        nav.entrar(c); p.wait_for_timeout(1200)
        p.goto(N.B + "/#/"); p.wait_for_timeout(1500); revisar(p, f"{tam}: inicio")
        p.goto(N.B + f"/#/espacio/{esp}"); p.wait_for_timeout(1500); revisar(p, f"{tam}: espacio")
        for modo, d in docs.items():
            p.goto(N.B + f"/#/doc/{d}"); p.wait_for_timeout(2500); revisar(p, f"{tam}: editor {modo.lower()}")
        for nombre, js in (("Ajustes", "NOM.vistaEspacios.preferencias()"), ("paleta de comandos", "document.querySelector('[data-a=paleta]').click()")):
            p.goto(N.B + f"/#/espacio/{esp}"); p.wait_for_timeout(1300)
            p.evaluate("() => { " + js + " }"); p.wait_for_timeout(900); revisar(p, f"{tam}: {nombre}")
        p.emulate_media(color_scheme="dark")
        for ruta, nombre in ((f"/#/espacio/{esp}", "espacio, tema oscuro"), (f"/#/doc/{docs['TEXTO']}", "texto, tema oscuro")):
            p.goto(N.B + ruta); p.wait_for_timeout(2000); revisar(p, f"{tam}: {nombre}")
        if nav.errores: print("  MAL  errores de JavaScript: " + "; ".join(nav.errores[:3])); cuenta_ok.append(False)
        nav.cerrar()
    if len(sys.argv) > 1: json.dump(res, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"Accesibilidad (axe-core): {sum(cuenta_ok)} bien, {len(cuenta_ok) - sum(cuenta_ok)} mal")
    sys.exit(0 if all(cuenta_ok) else 1)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Revisión automática de accesibilidad con axe-core (reglas WCAG 2.0, 2.1 y 2.2, niveles A y AA) en las pantallas
principales, en computadora y en celular: entrada, avisos legales, ayuda, inicio, un espacio, los cuatro editores
CON contenido (texto con tabla e imagen, hoja con datos, presentación con figuras, diagrama con conectores),
Ajustes abiertos, paleta de comandos, tema oscuro y colores forzados. Cada pantalla «bien» = ninguna regla falla.

Lo automático encuentra solo una parte de las barreras (nombres, contraste, estructura); lo demás se prueba a mano
con un lector de pantalla (ver docs/ACCESIBILIDAD.md y docs/PUBLICACION-Y-CUMPLIMIENTO.md, sección 5).

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
PNG_DATA = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAFElEQVR4nGP4z8CAFWEXHbQSDAwAPaMI+K6hSU8AAAAASUVORK5CYII="
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


def poblar_editores(p, docs):
    """Llena los 4 documentos con contenido real (tabla e imagen en texto, datos en hoja, figuras en presentación, conectores en diagrama)."""
    # 1. Texto: párrafo, tabla con encabezado e imagen con texto alternativo
    p.goto(N.B + f"/#/doc/{docs['TEXTO']}")
    p.wait_for_selector(".wd-cuerpo", timeout=20000)
    p.wait_for_timeout(1200)
    p.evaluate(r"""(src) => {
      const ed = NOM.editorActual, doc = ed && ed.doc, api = NOM.modos.TEXTO && NOM.modos.TEXTO.textoApi;
      if (!doc || !api) return;
      api.insertarTramos([{ x: "Informe de accesibilidad\n" }, { x: "Resumen de verificación en cuatro editores.\n" }]);
      const refTab = NOM.DocCRDT.nuevoId();
      doc.cambiar(refTab, { tipo: "tabla", enTexto: true, filasTexto: 2, colsTexto: 2, encabezado: true,
        "c:A1": "Editor", "c:B1": "Estado", "c:A2": "Texto", "c:B2": "Listo" }, { origen: "comando" });
      api.insertarBloque({ tipo: "tabla", ref: refTab });
      api.insertarImagen({ src, w: 120, h: 120 });
      const img = doc.vivos().find((b) => b.tipo === "imagen");
      if (img) doc.cambiar(img.id, { alt: "Muestra gráfica de prueba" }, { origen: "comando" });
      if (doc.guardarYa) doc.guardarYa();
    }""", PNG_DATA)
    p.wait_for_timeout(1200)

    # 2. Hoja: encabezados, datos numéricos y fórmula =SUMA(B2:B3)
    p.goto(N.B + f"/#/doc/{docs['HOJA']}")
    p.wait_for_selector(".rejilla-hoja", timeout=20000)
    p.wait_for_timeout(1200)
    p.evaluate(r"""() => {
      const ed = NOM.editorActual, doc = ed && ed.doc;
      if (!doc) return;
      const hoja = (NOM.formulas && NOM.formulas.hojasDe(doc)[0]) || doc.vivos().find((b) => b.tipo === "tabla" || b.tipo === "hoja");
      if (!hoja) return;
      doc.cambiar(hoja.id, {
        nombre: "Resumen",
        "c:A1": "Concepto", "c:B1": "Cantidad", "n:A1": true, "n:B1": true,
        "c:A2": "Libros", "c:B2": "12",
        "c:A3": "Cuadernos", "c:B3": "8",
        "c:A4": "Total", "c:B4": "=SUMA(B2:B3)",
      });
      if (doc.guardarYa) doc.guardarYa();
    }""")
    p.wait_for_timeout(1200)

    # 3. Presentación: título y figuras
    p.goto(N.B + f"/#/doc/{docs['PRESENTACION']}")
    p.wait_for_selector(".diapo", timeout=20000)
    p.wait_for_timeout(1200)
    p.evaluate(r"""() => {
      const ed = NOM.editorActual, doc = ed && ed.doc;
      if (!doc) return;
      const pg = ed.pagina;
      doc.cambiar("t1", { tipo: "texto", estilo: "h1", pagina: pg, orden: 1, x: 80, y: 40, w: 800, h: 80, z: 1, texto: "Presentación accesible", html: "<div>Presentación accesible</div>" });
      doc.cambiar("f1", { tipo: "forma", forma: "rect", pagina: pg, orden: 2, x: 100, y: 170, w: 260, h: 140, z: 2, html: "<div>Bloque principal</div>", fondo: "#D9E8FF" });
      doc.cambiar("f2", { tipo: "forma", forma: "elipse", pagina: pg, orden: 3, x: 460, y: 170, w: 240, h: 140, z: 3, html: "<div>Detalle</div>", fondo: "#DDF5E6" });
      if (NOM.modos.PRESENTACION && NOM.modos.PRESENTACION._lienzo) NOM.modos.PRESENTACION._lienzo.render();
      if (doc.guardarYa) doc.guardarYa();
    }""")
    p.wait_for_timeout(1200)

    # 4. Diagrama: figuras unidas con conector
    p.goto(N.B + f"/#/doc/{docs['DIAGRAMA']}")
    p.wait_for_selector(".pizarra", timeout=20000)
    p.wait_for_timeout(1200)
    p.evaluate(r"""() => {
      const ed = NOM.editorActual, doc = ed && ed.doc;
      if (!doc) return;
      const pg = ed.pagina;
      doc.cambiar("n1", { tipo: "forma", forma: "rect", pagina: pg, orden: 1, x: 80, y: 100, w: 180, h: 90, z: 1, html: "<div>Inicio</div>" });
      doc.cambiar("n2", { tipo: "forma", forma: "rombo", pagina: pg, orden: 2, x: 360, y: 100, w: 180, h: 110, z: 2, html: "<div>Revisión</div>" });
      doc.cambiar("n3", { tipo: "forma", forma: "rect", pagina: pg, orden: 3, x: 640, y: 100, w: 180, h: 90, z: 3, html: "<div>Publicado</div>" });
      doc.cambiar("c1", { tipo: "conector", desde: "n1", hasta: "n2", pagina: pg, orden: 0 });
      doc.cambiar("c2", { tipo: "conector", desde: "n2", hasta: "n3", pagina: pg, orden: 0 });
      if (NOM.modos.DIAGRAMA && NOM.modos.DIAGRAMA._lienzo) NOM.modos.DIAGRAMA._lienzo.render();
      if (doc.guardarYa) doc.guardarYa();
    }""")
    p.wait_for_timeout(1200)


def revisar_ajustes(p, esp, etiqueta):
    p.goto(N.B + f"/#/espacio/{esp}")
    p.wait_for_timeout(1300)
    p.evaluate("() => { NOM.vistaEspacios.preferencias(); }")
    p.wait_for_selector("section.hoja[role=dialog]", timeout=10000)
    p.wait_for_timeout(900)
    revisar(p, etiqueta)
    p.evaluate("() => { NOM.cerrarHoja && NOM.cerrarHoja(); }")


def main():
    if not os.path.exists(AXE): sys.exit(f"Falta axe-core en {AXE}: correr pruebas/preparar.sh")
    if not N.esperar_servidor(): sys.exit("El servidor no contesta en " + N.B)
    c = N.cuenta(); esp = N.espacio(c["token"], "Revisión de accesibilidad")
    docs = {m: N.documento(c["token"], esp, m, "Prueba " + m.lower()) for m in ("TEXTO", "HOJA", "PRESENTACION", "DIAGRAMA")}
    poblado = False
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
        if not poblado:
            poblar_editores(p, docs)
            poblado = True
        p.goto(N.B + "/#/"); p.wait_for_timeout(1500); revisar(p, f"{tam}: inicio")
        p.goto(N.B + f"/#/espacio/{esp}"); p.wait_for_timeout(1500); revisar(p, f"{tam}: espacio")
        for modo, d in docs.items():
            p.goto(N.B + f"/#/doc/{d}"); p.wait_for_timeout(2500); revisar(p, f"{tam}: editor {modo.lower()} con contenido")
        revisar_ajustes(p, esp, f"{tam}: Ajustes")
        p.goto(N.B + f"/#/espacio/{esp}"); p.wait_for_timeout(1300)
        p.evaluate("() => { document.querySelector('[data-a=paleta]').click(); }"); p.wait_for_timeout(900)
        revisar(p, f"{tam}: paleta de comandos")

        # Tema oscuro: inicio, espacio, los 4 editores con contenido y Ajustes abiertos
        p.emulate_media(color_scheme="dark", forced_colors="none")
        for ruta, nombre in (
            ("/#/", "inicio, tema oscuro"),
            (f"/#/espacio/{esp}", "espacio, tema oscuro"),
            (f"/#/doc/{docs['TEXTO']}", "texto con contenido, tema oscuro"),
            (f"/#/doc/{docs['HOJA']}", "hoja con contenido, tema oscuro"),
            (f"/#/doc/{docs['PRESENTACION']}", "presentacion con contenido, tema oscuro"),
            (f"/#/doc/{docs['DIAGRAMA']}", "diagrama con contenido, tema oscuro"),
        ):
            p.goto(N.B + ruta); p.wait_for_timeout(2000); revisar(p, f"{tam}: {nombre}")
        revisar_ajustes(p, esp, f"{tam}: Ajustes, tema oscuro")

        # Colores forzados (alto contraste del sistema): claro y oscuro, en inicio, editores y Ajustes
        for esquema, etq_esq in (("light", "colores forzados"), ("dark", "colores forzados + tema oscuro")):
            p.emulate_media(forced_colors="active", color_scheme=esquema)
            for ruta, nombre in (
                ("/#/", f"inicio, {etq_esq}"),
                (f"/#/espacio/{esp}", f"espacio, {etq_esq}"),
                (f"/#/doc/{docs['TEXTO']}", f"texto con contenido, {etq_esq}"),
                (f"/#/doc/{docs['HOJA']}", f"hoja con contenido, {etq_esq}"),
                (f"/#/doc/{docs['PRESENTACION']}", f"presentacion con contenido, {etq_esq}"),
                (f"/#/doc/{docs['DIAGRAMA']}", f"diagrama con contenido, {etq_esq}"),
            ):
                p.goto(N.B + ruta); p.wait_for_timeout(2000); revisar(p, f"{tam}: {nombre}")
            revisar_ajustes(p, esp, f"{tam}: Ajustes, {etq_esq}")

        p.emulate_media(color_scheme="light", forced_colors="none")
        if nav.errores: print("  MAL  errores de JavaScript: " + "; ".join(nav.errores[:3])); cuenta_ok.append(False)
        nav.cerrar()
    if len(sys.argv) > 1: json.dump(res, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"Accesibilidad (axe-core): {sum(cuenta_ok)} bien, {len(cuenta_ok) - sum(cuenta_ok)} mal")
    sys.exit(0 if all(cuenta_ok) else 1)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Clasificador de diapositivas (Presentar → Clasificador de diapositivas, como la vista de PowerPoint) con un
PowerPoint real (p13-presentacion.pptx, con tres secciones): todas las diapositivas en la cuadrícula con sus secciones, arrastrar una
cambia el orden, una oculta se ve atenuada, Ctrl+D duplica, Supr elimina, las flechas y Entrar abren la elegida. Y la
Vista de lectura: en la ventana, sin pantalla completa, con su barra para avanzar, y Esc la cierra.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_clasificador.py [capturas]
"""
import base64, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "office"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:400]}"), flush=True)


def main():
    arch = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real/archivos/p13-presentacion.pptx"
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    paginas = "() => NOM.editorActual.doc.paginas()"
    try:
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
        pags = p.evaluate(paginas)
        secs = p.evaluate("() => { const s = NOM.ooxml.seccionesPptx(NOM.editorActual.doc); return s ? s.length : 0; }")
        p.evaluate("() => { window.__clas = NOM.editorActual.modoApi._pres.clasificador(); }"); p.wait_for_timeout(1200)
        dlg = p.locator("section.hoja[role=dialog]").last
        anota(f"las {len(pags)} diapositivas en la cuadrícula, con sus {secs} secciones", dlg.locator(".clas-item").count() == len(pags) and dlg.locator(".clas-seccion").count() == secs, [dlg.locator(".clas-item").count(), dlg.locator(".clas-seccion").count()])
        # arrastrar la primera después de la tercera
        a, b = dlg.locator(".clas-item").nth(0), dlg.locator(".clas-item").nth(2)
        bb = b.bounding_box()
        a.drag_to(b, target_position={"x": bb["width"] * 0.8, "y": bb["height"] / 2}); p.wait_for_timeout(600)
        q = p.evaluate(paginas)
        anota("arrastrar la 1 después de la 3 cambia el orden", q[:3] == [pags[1], pags[2], pags[0]], [pags[:3], q[:3]])
        # una oculta se ve atenuada
        p.evaluate("() => { const d = NOM.editorActual.doc, x = d.paginas().find((y) => !d.valor('_doc', 'oculta:' + y)); NOM.editorActual.modoApi._pres.alternarOculta(x); }"); p.wait_for_timeout(300)
        p.evaluate("() => window.__clas.cerrar()"); p.wait_for_timeout(300)
        p.evaluate("() => { window.__clas = NOM.editorActual.modoApi._pres.clasificador(); }"); p.wait_for_timeout(1000)
        dlg = p.locator("section.hoja[role=dialog]").last
        estado = p.evaluate("() => NOM.editorActual.doc.paginas().map((x) => !!NOM.editorActual.doc.valor('_doc', 'oculta:' + x))")
        clases = p.evaluate("() => NOM.editorActual.doc.paginas().map((x) => document.querySelector(`.clas-item[data-p=\"${CSS.escape(x)}\"]`).classList.contains('oculta'))")
        anota("las ocultas se ven atenuadas (y las demás no)", estado == clases and any(estado), [estado, clases])
        # duplicar y eliminar con el teclado
        dlg.locator(".clas-item").nth(0).click(); p.wait_for_timeout(200)
        p.keyboard.press("Control+d"); p.wait_for_timeout(500)
        anota("Ctrl+D la duplica", len(p.evaluate(paginas)) == len(pags) + 1, len(p.evaluate(paginas)))
        p.keyboard.press("Delete"); p.wait_for_timeout(500)
        anota("Supr la elimina", len(p.evaluate(paginas)) == len(pags), len(p.evaluate(paginas)))
        # flechas y Entrar
        dlg.locator(".clas-item").nth(0).click(); p.keyboard.press("ArrowRight"); p.keyboard.press("ArrowRight"); p.keyboard.press("Enter"); p.wait_for_timeout(700)
        anota("las flechas y Entrar abren la elegida (la 3)", p.evaluate("() => NOM.editorActual.pagina") == p.evaluate(paginas)[2] and p.locator("section.hoja[role=dialog] .clasificador").count() == 0)
        # Vista de lectura: en la ventana (sin pantalla completa), con su barra para ir adelante y atrás
        p.evaluate("() => NOM.editorActual.modoApi._pres.presentar({ lectura: true, desde: 'principio' })"); p.wait_for_timeout(1200)
        rep = p.locator(".reproductor.lectura")
        c0 = rep.locator("[data-c]").inner_text()
        rep.locator(".pres-lectura [data-sig-btn]").click(); p.wait_for_timeout(900)
        c1 = rep.locator("[data-c]").inner_text()
        anota("Vista de lectura: en la ventana, sin pantalla completa, y su barra avanza", rep.count() == 1 and p.evaluate("() => !document.fullscreenElement") and c0 != c1, [c0, c1])
        p.keyboard.press("Escape"); p.wait_for_timeout(600)
        anota("…y Esc la cierra", p.locator(".reproductor").count() == 0)
        # Página de notas: una hoja por diapositiva, con sus notas, que se escriben ahí mismo
        p.evaluate("() => { window.__notas = NOM.editorActual.modoApi._pres.paginaNotas(); }"); p.wait_for_timeout(1000)
        dlg = p.locator("section.hoja[role=dialog]").last
        n_hojas = dlg.locator(".nota-pagina").count()
        dlg.locator(".nota-texto").nth(1).fill("Recordar los números del trimestre"); p.wait_for_timeout(400)
        guardada = p.evaluate("() => { const d = NOM.editorActual.doc; return d.valor('_doc', 'nota:' + d.paginas()[1]); }")
        anota("Página de notas: una hoja por diapositiva y las notas se escriben ahí", n_hojas == len(p.evaluate(paginas)) and guardada == "Recordar los números del trimestre", [n_hojas, guardada])
        p.evaluate("() => document.body.classList.add('imprime-notas')"); p.emulate_media(media="print"); p.wait_for_timeout(300)
        visibles = p.evaluate("() => [document.querySelector('.wd, .lienzo, [data-escena]') ? getComputedStyle(document.querySelector('[data-escena]') || document.body).display : 'x', getComputedStyle(document.querySelector('.nota-pagina')).display]")
        p.emulate_media(media="screen"); p.evaluate("() => document.body.classList.remove('imprime-notas')")
        anota("…y al imprimir solo salen las hojas de notas", visibles[1] != "none", visibles)
        p.evaluate("() => window.__notas.cerrar()")
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nClasificador de diapositivas: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

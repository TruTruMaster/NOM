#!/usr/bin/env python3
"""Celda de enfoque de la hoja (Vista → Celda de enfoque, como Excel 2024): con ella prendida, la fila y la columna de la
celda elegida se resaltan (todas sus celdas, menos la elegida), siguen a la selección al moverse con el teclado, no
cambian el libro ni se imprimen, y al apagarla no queda nada.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_enfoque.py [capturas]
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
import nomprueba as N  # noqa: E402

res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:300]}"), flush=True)


JS = """() => { const e = [...document.querySelectorAll('td.enfoque')]; const s = document.querySelector('td.sel');
  return { n: e.length, filas: [...new Set(e.map((x) => +x.dataset.f))].length, cols: [...new Set(e.map((x) => +x.dataset.c))].length,
    ok: e.every((x) => +x.dataset.f === +s.dataset.f || +x.dataset.c === +s.dataset.c), sel: [+s.dataset.c, +s.dataset.f],
    todas: document.querySelectorAll('td[data-c]').length, filasT: new Set([...document.querySelectorAll('td[data-c]')].map((x) => x.dataset.f)).size,
    colsT: new Set([...document.querySelectorAll('td[data-c]')].map((x) => x.dataset.c)).size }; }"""


def main():
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Enfoque")
    nav = N.Navegador(movil=False); nav.entrar(a); p = nav.p
    try:
        d = N.documento(a["token"], esp, "HOJA", "Enfoque")
        p.goto(N.B + "/#/doc/" + d); p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.api", timeout=30000)
        p.wait_for_timeout(600)
        antes = p.evaluate("() => JSON.stringify(NOM.editorActual.modoApi.api.tabla())")
        p.locator('td[data-c="2"][data-f="4"]').click(); p.wait_for_timeout(200)
        p.evaluate("() => NOM.editorActual.modoApi.api.enfoque(true)"); p.wait_for_timeout(300)
        x = p.evaluate(JS)
        anota("se resaltan la fila y la columna de C5", x["ok"] and x["sel"] == [2, 4] and x["n"] == x["filasT"] + x["colsT"] - 2, x)
        p.keyboard.press("ArrowDown"); p.keyboard.press("ArrowRight"); p.wait_for_timeout(300)
        y = p.evaluate(JS)
        anota("…y siguen a la selección (D6)", y["ok"] and y["sel"] == [3, 5], y)
        anota("…sin cambiar el libro", p.evaluate("() => JSON.stringify(NOM.editorActual.modoApi.api.tabla())") == antes)
        p.emulate_media(media="print"); p.wait_for_timeout(200)
        anota("no se imprime", p.evaluate("() => getComputedStyle(document.querySelector('td.enfoque')).boxShadow") == "none")
        p.emulate_media(media="screen")
        p.evaluate("() => NOM.editorActual.modoApi.api.enfoque(false)"); p.wait_for_timeout(300)
        anota("al apagarla no queda nada", p.evaluate("() => document.querySelectorAll('td.enfoque').length") == 0)
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nCelda de enfoque: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

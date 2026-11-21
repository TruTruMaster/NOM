#!/usr/bin/env python3
"""Estilos de celda (Más → Estilos de celda…) aplicados en NOM contra Excel 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/estilos_celda_real.py

x13-estilos-celda.xlsx (hecho en Excel con windows/hacer_estilos_celda.py: los 47 estilos, en cada fila un texto y un
número) abierto en NOM: a cada celda se le quita su estilo con «Normal» y se le aplica otra vez desde NOM. Descargado,
Excel lo pinta (windows/verdad.py) y cada celda debe verse igual que la de x13: texto, formato de número, fuente, tamaño, negrita, cursiva, subrayado, color de
letra y relleno. Además: la galería aplica el estilo con un clic, «Normal» quita lo de antes y Excel abre el libro sin
avisos.
"""
import base64, json, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
CAMPOS = ["texto", "fmt", "fuente", "tam", "b", "i", "u", "color", "fondo"]
res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:700]}"), flush=True)


def celdas(v):
    return {(x["f"], x["c"]): x for x in v["hojas"][0]["celdas"]}


def main():
    import remoto
    x13 = json.loads((REAL / "verdad/x13-estilos-celda.verdad.json").read_text(encoding="utf-8"))
    nombres = [x["texto"] for x in sorted(x13["hojas"][0]["celdas"], key=lambda q: q["f"]) if x["c"] == 1]
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    try:
        # (sobre el propio x13 abierto en NOM —la misma letra del libro y los mismos anchos—: a cada celda se le quita su
        # estilo con «Normal» y se le pone otra vez el suyo desde NOM)
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        x13x = REAL / "archivos/x13-estilos-celda.xlsx"
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(x13x.read_bytes()).decode(), x13x.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
        r = p.evaluate("""(ns) => { const H = NOM.editorActual.modoApi.api;
          return ns.map((n, i) => { H.estiloCelda('Normal', { c1: 1, c2: 2, f1: i, f2: i }); return H.estiloCelda(n, { c1: 1, c2: 2, f1: i, f2: i }); }); }""", nombres)
        anota(f"los {len(nombres)} estilos se aplican", all(x == 2 for x in r), r)
        p.wait_for_timeout(800)
        destino = REAL / "archivos/x13-nom.xlsx"
        with p.expect_download(timeout=120000) as dl:
            p.evaluate("() => NOM.formatos.exportarListo(NOM.editorActual.doc, 'xlsx', 'x13-nom')")
        dl.value.save_as(destino)
        remoto.subir([destino], "nomoffice/archivos")
        remoto.correr_office("verdad-x13nom", "verdad.py", "archivos\\x13-nom.xlsx", timeout=600)
        vn = REAL / "verdad/x13-nom.verdad.json"; remoto.bajar("nomoffice/archivos/verdad/x13-nom.verdad.json", vn)
        a, b = celdas(x13), celdas(json.loads(vn.read_text(encoding="utf-8")))
        # (la letra predeterminada del libro no es del estilo: un libro nuevo de Excel 2024 es Aptos Narrow y uno de NOM,
        # Calibri; se compara la letra contra la predeterminada de cada libro)
        from collections import Counter
        for cs in (a, b):
            pred = Counter(x.get("fuente") for x in cs.values()).most_common(1)[0][0]
            for x in cs.values():
                if x.get("fuente") == pred: x["fuente"] = "(la del libro)"
        malos = []
        for f in range(1, len(nombres) + 1):
            for col in (2, 3):
                x, y = a.get((f, col), {}), b.get((f, col), {})
                d = {k: (x.get(k), y.get(k)) for k in CAMPOS if x.get(k) != y.get(k)}
                if d: malos.append((nombres[f - 1], col, d))
        anota(f"Excel pinta las {2 * len(nombres)} celdas igual que las de su libro (texto, número, letra, color y relleno)", not malos, malos[:6])
        # la galería y «Normal»
        p.locator('td[data-c="4"][data-f="0"]').click(); p.wait_for_timeout(200)
        p.evaluate("() => { const t = NOM.editorActual.modoApi.api.tabla(); NOM.editorActual.doc.cambiar(t.id, { 'c:E1': 'hola', 'n:E1': true, 'r:E1': '#FF0000' }); }"); p.wait_for_timeout(300)
        p.evaluate("() => NOM.editorActual.modoApi.api.dialogos.estilosCelda()"); p.wait_for_timeout(400)
        p.locator('section.hoja[role=dialog] [data-estilo-celda="Good"]').click(); p.wait_for_timeout(400)
        e1 = p.evaluate("() => { const t = NOM.editorActual.modoApi.api.tabla(); return [t['k:E1'], t['r:E1'], t['n:E1'] || false]; }")
        anota("la galería pone «Bueno» con un clic (y reemplaza el relleno, como Excel)", e1[0] == "#006100" and e1[1] == "#C6EFCE", e1)
        p.evaluate("() => NOM.editorActual.modoApi.api.estiloCelda('Normal', { c1: 4, c2: 4, f1: 0, f2: 0 })"); p.wait_for_timeout(300)
        e2 = p.evaluate("() => { const t = NOM.editorActual.modoApi.api.tabla(); return [t['k:E1'] || null, t['r:E1'] || null, t['n:E1'] || null]; }")
        anota("«Normal» le quita todo", e2 == [None, None, None], e2)
        import office_abre
        rr = office_abre.abrir([destino], "estilos-celda")
        anota("Excel abre x13-nom.xlsx sin avisos ni reparar", office_abre.bien(rr.get(destino.name)), json.dumps(rr.get(destino.name), ensure_ascii=False)[:200])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nEstilos de celda contra Excel 2024: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

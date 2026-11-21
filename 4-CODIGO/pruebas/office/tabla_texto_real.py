#!/usr/bin/env python3
"""Convertir texto en tabla y tabla en texto (Más → Insertar) en NOM contra Word 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/tabla_texto_real.py [--solo-nom]

Los casos van a Word (windows/tabla_texto_word.py: Range.ConvertToTable y Table.ConvertToText) y se hacen igual en NOM
(un documento por caso, escrito con el teclado): tabulaciones, comas, otro carácter, por párrafos con 2 columnas, sin
decir el separador (Word adivina), celdas vacías; y de tabla a texto con tabulaciones, comas («, »), párrafos y otro
carácter. Además, una tabla hecha así se descarga y Word lee las mismas celdas, y Ctrl+Z la regresa a texto.
"""
import json, os, sys, tempfile
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import nomprueba as N  # noqa: E402

CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas"))
VERDAD = CACHE / "office-real/tabla-texto-word.json"
CASOS = [
    {"id": "tabs", "parrafos": ["Nombre\tEdad\tCiudad", "Ana\t30\tPuebla", "Luis\t25"], "sep": 1},
    {"id": "comas", "parrafos": ["a,b,c", "1,2,3", "x,,z"], "sep": 2},
    {"id": "otro", "parrafos": ["uno;dos", "tres;cuatro;cinco"], "sep": ";"},
    {"id": "parrafos", "parrafos": ["uno", "dos", "tres", "cuatro", "cinco"], "sep": 0, "cols": 2},
    {"id": "adivina-tabs", "parrafos": ["a\tb", "c\td"], "sep": None},
    {"id": "adivina-comas", "parrafos": ["a,b", "c,d"], "sep": None},
    {"id": "adivina-nada", "parrafos": ["uno", "dos"], "sep": None},
    {"id": "tabs-vacios", "parrafos": ["a\t\tc", "\tb\t"], "sep": 1},
    {"id": "a-texto-tabs", "tabla": [["Mes", "Ventas"], ["Enero", "120"], ["Febrero", ""]], "sep": 1, "a_texto": True},
    {"id": "a-texto-comas", "tabla": [["a", "b", "c"], ["1", "2", "3"]], "sep": 2, "a_texto": True},
    {"id": "a-texto-parrafos", "tabla": [["a", "b"], ["c", "d"]], "sep": 0, "a_texto": True},
    {"id": "a-texto-otro", "tabla": [["a", "b"], ["c", "d"]], "sep": "|", "a_texto": True},
]
SEP = {0: "parrafo", 1: "tab", 2: "coma"}


def verdad_word():
    import remoto
    with tempfile.TemporaryDirectory() as d:
        ent = Path(d) / "tt-casos.json"; ent.write_text(json.dumps(CASOS, ensure_ascii=False), encoding="utf-8")
        remoto.subir([ent], "nomoffice")
        remoto.correr_office("tt", "tabla_texto_word.py", "tt-casos.json", "tt-verdad.json", timeout=600)
        if not remoto.bajar("nomoffice/tt-verdad.json", VERDAD): raise SystemExit("Word no dejó el resultado")
    return json.loads(VERDAD.read_text(encoding="utf-8"))


JS_TABLA = """() => { const d = NOM.editorActual.doc, t = d.vivos().find((b) => b.tipo === 'tabla' && b.enTexto); if (!t) return null;
  const out = []; for (let i = 0; i < t.filasTexto; i++) { const f = []; for (let j = 0; j < t.colsTexto; j++) f.push(String(t['c:' + NOM.formulas.ref(j, i)] || '')); out.push(f); } return out; }"""
JS_PARRAFOS = "() => NOM.textoRico.parrafosDe(NOM.editorActual.doc).map((p) => p.tramos.map((t) => t.x || '').join(''))"


def escribir(p, parrafos):
    p.locator(".wd [contenteditable=true]").first.click(); p.wait_for_timeout(200)
    for i, t in enumerate(parrafos):
        if i: p.keyboard.press("Enter")
        for k, trozo in enumerate(t.split("\t")):
            if k: p.keyboard.press("Tab")
            if trozo: p.keyboard.insert_text(trozo)
    p.wait_for_timeout(400)


def main():
    v = json.loads(VERDAD.read_text(encoding="utf-8")) if "--solo-nom" in sys.argv and VERDAD.exists() else verdad_word()
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Tablas")
    nav = N.Navegador(movil=False); nav.entrar(a); p = nav.p
    bien = mal = 0
    def anota(n, ok, det=""):
        nonlocal bien, mal
        bien += ok; mal += not ok
        print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:500]}"), flush=True)
    T = "NOM.editorActual.modoApi.tablas"
    try:
        for c in CASOS:
            d = N.documento(a["token"], esp, "TEXTO", c["id"])
            p.goto(N.B + "/#/doc/" + d); p.wait_for_selector(".wd [contenteditable=true]", timeout=30000); p.wait_for_timeout(600)
            sep = SEP.get(c["sep"], c["sep"]) if c["sep"] is not None else None
            if c.get("a_texto"):
                escribir(p, ["\t".join(f) for f in c["tabla"]])
                p.keyboard.press("Control+a")
                p.evaluate(f"() => {T}.textoEnTabla('tab')"); p.wait_for_timeout(500)
                p.keyboard.press("Control+Home")
                r = p.evaluate(f"(s) => {T}.tablaEnTexto(s)", sep); p.wait_for_timeout(500)
                nom = [x for x in p.evaluate(JS_PARRAFOS)]
                while nom and nom[-1] == "": nom.pop()
                exp = [x for x in v[c["id"]]["texto"]]
                while exp and exp[-1] == "": exp.pop()
                anota(c["id"], nom == exp, f"Word {exp} · NOM {nom} · {r}")
            else:
                escribir(p, c["parrafos"])
                p.keyboard.press("Control+a")
                r = p.evaluate(f"(o) => {T}.textoEnTabla(o.s, o.c)", {"s": sep, "c": c.get("cols")}); p.wait_for_timeout(500)
                nom = p.evaluate(JS_TABLA)
                anota(c["id"], nom == v[c["id"]].get("celdas"), f"Word {v[c['id']].get('celdas')} · NOM {nom} · {r}")
        # descargada: Word lee la misma tabla; Ctrl+Z la regresa a texto
        import remoto
        d = N.documento(a["token"], esp, "TEXTO", "descargar")
        p.goto(N.B + "/#/doc/" + d); p.wait_for_selector(".wd [contenteditable=true]", timeout=30000); p.wait_for_timeout(600)
        escribir(p, ["Antes de la tabla.", "Producto\tPrecio", "Pera\t12", "Manzana\t15", "Después de la tabla."])
        p.keyboard.press("Control+Home"); p.keyboard.press("ArrowDown"); p.keyboard.press("Shift+ArrowDown"); p.keyboard.press("Shift+ArrowDown"); p.keyboard.press("Shift+End")
        p.evaluate(f"() => {T}.textoEnTabla(null)"); p.wait_for_timeout(600)
        destino = CACHE / "office-real/tabla-texto-nom.docx"
        with p.expect_download(timeout=120000) as dl:
            p.evaluate("() => NOM.formatos.exportarListo(NOM.editorActual.doc, 'docx', 'tabla-texto-nom')")
        dl.value.save_as(destino)
        remoto.subir([destino], "nomoffice/entrada")
        remoto.correr_office("tv", "tablas_ver.py", "entrada\\tabla-texto-nom.docx", "tablas-ver.json", timeout=300)
        tmp = CACHE / "office-real/tablas-ver.json"; remoto.bajar("nomoffice/tablas-ver.json", tmp)
        w = json.loads(tmp.read_text(encoding="utf-8"))
        anota("descargada: Word lee la misma tabla y el texto de antes y de después", w.get("tablas") == [[["Producto", "Precio"], ["Pera", "12"], ["Manzana", "15"]]] and "Antes de la tabla." in w.get("parrafos", []) and "Después de la tabla." in w.get("parrafos", []), w)
        p.keyboard.press("Control+z"); p.wait_for_timeout(700)
        en_texto = "() => NOM.textoRico.parrafosDe(NOM.editorActual.doc).some((q) => q.tramos.some((t) => t.e && t.e.tipo === 'tabla'))"
        anota("Ctrl+Z la regresa a texto", not p.evaluate(en_texto) and "Pera\t12" in p.evaluate(JS_PARRAFOS), p.evaluate(JS_PARRAFOS))
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nTexto y tablas contra Word 2024: {bien} bien, {mal} mal")
    return 1 if mal else 0


if __name__ == "__main__":
    sys.exit(main())

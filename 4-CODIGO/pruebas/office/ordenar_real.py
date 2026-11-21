#!/usr/bin/env python3
"""Ordenar párrafos (Inicio → Párrafo → Ordenar) en NOM contra Word 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/ordenar_real.py [--solo-nom]

Los casos van a Word (windows/ordenar_word.py: Range.Sort con su idioma) y se ordenan igual en NOM (un documento por caso,
escrito con el teclado, todo elegido): texto en español y en inglés, al revés, con y sin distinguir mayúsculas, signos,
guiones y apóstrofos, números y letras, por número, por fecha y con encabezado. Además: el formato de cada párrafo (una
lista, un título) se va con él, Ctrl+Z regresa el orden, y Word lee lo que descarga NOM en el mismo orden.
Fuera de la comparación: números con «+», «%» o entre paréntesis (Word los ordena de una manera que no se pudo explicar).
"""
import json, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import nomprueba as N  # noqa: E402

CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas"))
VERDAD = CACHE / "office-real/ordenar-word.json"
TIPO = {0: "texto", 1: "numero", 2: "fecha"}
CASOS = [
    {"id": "texto", "parrafos": ["pera", "Manzana", "árbol", "zanahoria", "Ñandú", "nuez", "Avión", "avena", "éxito", "Eco", "chile", "Ciruela", "llave", "luna", "10 perros", "2 gatos", "", "_guion", "(paréntesis)", "Zorro"]},
    {"id": "texto-desc", "parrafos": ["pera", "Manzana", "árbol", "zanahoria", "Ñandú", "nuez", "Avión", "avena", "", "Zorro"], "orden": 1},
    {"id": "mayusculas", "parrafos": ["casa", "Casa", "CASA", "cama", "Cama", "árbol", "Árbol"], "mayusculas": True},
    {"id": "iguales", "parrafos": ["casa", "Casa", "CASA", "cása", "casa "]},
    {"id": "caso-rev", "parrafos": ["CASA", "Casa", "casa", "CAMA", "cama"]},
    {"id": "caso-rev-sens", "parrafos": ["CASA", "Casa", "casa", "CAMA", "cama"], "mayusculas": True},
    {"id": "guiones", "parrafos": ["co-op", "coop", "co op", "cop", "o'neil", "oneil", "one", "la casa", "lacasa", "la-casa", "lab"]},
    {"id": "simbolos", "parrafos": ["#uno", "!dos", "@tres", "$cuatro", "%cinco", "&seis", "*siete", "+ocho", "-nueve", "/diez", ":once", "?doce", "[trece", "~catorce", "¿quince", "¡dieciséis", "«diecisiete", "\"dieciocho"]},
    {"id": "digitos-letras", "parrafos": ["1a", "1", "a1", "11", "1 b", "1b", "A", "½"]},
    {"id": "numeros", "parrafos": ["10 perros", "2 gatos", "-3 frío", "1.5 kilos", "sin número", "100", "1,000 pesos", "x7", "007", "", "3e2"], "tipo": 1},
    {"id": "numeros-desc", "parrafos": ["10 perros", "2 gatos", "-3 frío", "1.5 kilos", "sin número", "100"], "tipo": 1, "orden": 1},
    {"id": "encabezado", "parrafos": ["Frutas", "pera", "manzana", "uva"], "encabezado": True},
    {"id": "fechas-es", "parrafos": ["15/03/2026", "2/1/2025", "31/12/2024", "no es fecha", "1 de mayo de 2026", "2026-01-01"], "tipo": 2},
    {"id": "fechas-en", "parrafos": ["3/15/2026", "1/2/2025", "12/31/2024", "no es fecha", "2026-01-01", "May 5, 2026"], "tipo": 2, "idioma": 1033},
    {"id": "ingles", "parrafos": ["pera", "Ñandú", "nuez", "chile", "Ciruela", "llave", "luna", "árbol", "avena"], "idioma": 1033},
]


def verdad_word():
    import remoto, tempfile
    with tempfile.TemporaryDirectory() as d:
        ent = Path(d) / "ordenar-casos.json"
        ent.write_text(json.dumps([dict({"tipo": 0, "orden": 0, "idioma": 2058}, **c) for c in CASOS], ensure_ascii=False), encoding="utf-8")
        remoto.subir([ent], "nomoffice")
        remoto.correr_office("ordenar", "ordenar_word.py", "ordenar-casos.json", "ordenar-verdad.json", timeout=600)
        if not remoto.bajar("nomoffice/ordenar-verdad.json", VERDAD): raise SystemExit("Word no dejó el resultado")
    return json.loads(VERDAD.read_text(encoding="utf-8"))


JS_PARRAFOS = "() => NOM.textoRico.parrafosDe(NOM.editorActual.doc).map((p) => p.tramos.map((t) => t.x || '').join(''))"


def escribir(p, parrafos):
    p.locator(".wd [contenteditable=true]").first.click(); p.wait_for_timeout(200)
    for i, t in enumerate(parrafos):
        if i: p.keyboard.press("Enter")
        if t: p.keyboard.insert_text(t)
    p.wait_for_timeout(500)


def main():
    v = json.loads(VERDAD.read_text(encoding="utf-8")) if "--solo-nom" in sys.argv and VERDAD.exists() else verdad_word()
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Ordenar")
    nav = N.Navegador(movil=False); nav.entrar(a); p = nav.p
    bien = mal = 0
    def anota(n, ok, det=""):
        nonlocal bien, mal
        bien += ok; mal += not ok
        print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:400]}"), flush=True)
    try:
        for c in CASOS:
            d = N.documento(a["token"], esp, "TEXTO", c["id"])
            p.goto(N.B + "/#/doc/" + d); p.wait_for_selector(".wd [contenteditable=true]", timeout=30000); p.wait_for_timeout(600)
            if c.get("idioma") == 1033: p.evaluate("() => NOM.editorActual.doc.cambiar('_doc', { idioma: 'en' })")
            escribir(p, c["parrafos"])
            p.keyboard.press("Control+a")
            r = p.evaluate("(o) => NOM.editorActual.modoApi.ordenar.parrafos(o)", {"tipo": TIPO[c.get("tipo", 0)], "desc": c.get("orden") == 1, "encabezado": bool(c.get("encabezado")), "mayus": bool(c.get("mayusculas"))})
            p.wait_for_timeout(500)
            nom = p.evaluate(JS_PARRAFOS)[:len(c["parrafos"])]
            anota(c["id"], nom == v[c["id"]], f"Word {v[c['id']]} · NOM {nom} · {r}")
        # el formato se va con su párrafo, Ctrl+Z regresa y Word lee el mismo orden
        d = N.documento(a["token"], esp, "TEXTO", "formato")
        p.goto(N.B + "/#/doc/" + d); p.wait_for_selector(".wd [contenteditable=true]", timeout=30000); p.wait_for_timeout(600)
        escribir(p, ["uva", "pera", "manzana"])
        p.keyboard.press("Control+Home"); p.keyboard.press("ArrowDown")   # «pera»: título 1
        p.evaluate("() => NOM.editorActual.modoApi.ordenar.estilo('h1')"); p.wait_for_timeout(400)
        p.keyboard.press("Control+a")
        p.evaluate("() => NOM.editorActual.modoApi.ordenar.parrafos({ tipo: 'texto' })"); p.wait_for_timeout(600)
        est = p.evaluate("() => NOM.textoRico.parrafosDe(NOM.editorActual.doc).map((q) => [q.tramos.map((t) => t.x || '').join(''), q.pa.estilo || ''])")
        anota("el estilo se va con su párrafo («pera», Título 1)", est[:3] == [["manzana", ""], ["pera", "h1"], ["uva", ""]], est)
        # descargado: Word lee el mismo orden
        import remoto
        destino = CACHE / "office-real/ordenado-nom.docx"
        with p.expect_download(timeout=120000) as dl:
            p.evaluate("() => NOM.formatos.exportarListo(NOM.editorActual.doc, 'docx', 'ordenado-nom')")
        dl.value.save_as(destino)
        remoto.subir([destino], "nomoffice/entrada")
        remoto.correr_office("cc-ord", "cambios_ver.py", "entrada\\ordenado-nom.docx", "ordenado.cc.json", timeout=300)
        tmp = CACHE / "office-real/ordenado.cc.json"
        remoto.bajar("nomoffice/ordenado.cc.json", tmp)
        w = json.loads(tmp.read_text(encoding="utf-8"))
        anota("Word lee lo que descarga NOM en el mismo orden", [x for x in w.get("tal_cual", []) if x][:3] == ["manzana", "pera", "uva"], w.get("tal_cual"))
        p.keyboard.press("Control+z"); p.wait_for_timeout(700)
        anota("Ctrl+Z regresa el orden de antes", p.evaluate(JS_PARRAFOS)[:3] == ["uva", "pera", "manzana"], p.evaluate(JS_PARRAFOS))
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nOrdenar contra Word 2024: {bien} bien, {mal} mal")
    return 1 if mal else 0


if __name__ == "__main__":
    sys.exit(main())

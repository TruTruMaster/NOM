#!/usr/bin/env python3
"""Estilos de párrafo creados y modificados en NOM (Más → Formato → Estilos…) contra Word 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/estilos_real.py

1. Un documento de NOM con un estilo propio «Nota importante» (Arial, 14 pt, negrita, rojo oscuro, centrado, 6 pt antes y
   12 después) aplicado a un párrafo y «Título 1» modificado (verde, 20 pt): NOM los pinta así.
2. Descargado, Word ve los mismos estilos (windows/estilos_ver.py: nombre y definición) y cada párrafo con el suyo.
3. Word lo pinta (windows/verdad.py) y NOM pinta los mismos renglones (comparar_real.py).
4. Abierto otra vez en NOM, el estilo propio sigue en la lista y en su párrafo.
5. En un Word que ya existía (w03), modificar «Título 1» cambia su definición en el archivo; Word lo abre sin reparar.
"""
import base64, json, os, subprocess, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas"))
REAL = CACHE / "office-real"
NOMBRE = "w20-estilos-nom.docx"
NOTA = {"n": "Nota importante", "fuente": "Arial", "tam": 14, "b": True, "color": "#C00000", "alin": "centro", "antes": 6, "despues": 12}
TITULO1 = {"color": "#00B050", "tam": 20}
res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:600]}"), flush=True)


def word_ver(docx, *estilos):
    import remoto
    remoto.subir([docx], "nomoffice/entrada")
    remoto.correr_office("est-" + docx.stem, "estilos_ver.py", f"entrada\\{docx.name}", f"{docx.stem}.est.json", *[f'"{e}"' for e in estilos], timeout=300)
    tmp = REAL / f"{docx.stem}.est.json"; remoto.bajar(f"nomoffice/{docx.stem}.est.json", tmp)
    return json.loads(tmp.read_text(encoding="utf-8"))


def bajar(p, destino):
    with p.expect_download(timeout=120000) as dl:
        p.evaluate("(n) => NOM.formatos.exportarListo(NOM.editorActual.doc, 'docx', n)", destino.stem)
    dl.value.save_as(destino)
    return destino


def main():
    import remoto
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    try:
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'Estilos', modo: 'TEXTO', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function(U.JS_LISTO, timeout=60000); p.wait_for_timeout(1000)
        p.locator(".wd [contenteditable=true]").first.click(); p.wait_for_timeout(200)
        for i, t in enumerate(["Informe del trimestre", "Aviso: las inscripciones cierran el viernes.", "El resto del texto va con el estilo Normal y no cambia.", "Otro título"]):
            if i: p.keyboard.press("Enter")
            p.keyboard.insert_text(t)
        p.wait_for_timeout(500)
        E = "NOM.editorActual.modoApi.estilos"
        ir = lambda i: (p.keyboard.press("Control+Home"), [p.keyboard.press("ArrowDown") for _ in range(i)])
        ir(1); clave = p.evaluate(f"(d) => {{ const k = {E}.guardar(null, d); {E}.aplicar(k); return k; }}", NOTA); p.wait_for_timeout(400)
        p.evaluate(f"(d) => {E}.guardar('h1', d)", TITULO1); p.wait_for_timeout(300)
        ir(0); p.evaluate(f"() => {E}.aplicar('h1')"); ir(3); p.evaluate(f"() => {E}.aplicar('h1')"); p.wait_for_timeout(800)
        vis = p.evaluate("""() => [...document.querySelectorAll('.wd-cuerpo > *')].slice(0, 4).map((e) => { const s = getComputedStyle(e); return [e.tagName, s.fontSize, s.fontWeight, s.color, s.textAlign, s.fontFamily.split(',')[0]]; })""")
        anota("NOM pinta el estilo propio (14 pt, negrita, rojo oscuro, centrado, Arial)", vis[1][1] == "18.6667px" and vis[1][2] == "700" and vis[1][3] == "rgb(192, 0, 0)" and vis[1][4] == "center" and ("AR" in vis[1][5].upper() or "Arial" in vis[1][5]), vis[1])
        anota("…y Título 1 modificado (verde, 20 pt) en sus dos párrafos", vis[0][0] == "H1" and vis[0][3] == "rgb(0, 176, 80)" and vis[0][1] == "26.6667px" and vis[3][3] == "rgb(0, 176, 80)", [vis[0], vis[3]])
        anota("«Nota importante» está en la lista de estilos", any(s["t"] == "Nota importante" for s in p.evaluate(f"() => {E}.lista()")))
        destino = bajar(p, REAL / "archivos" / NOMBRE)
        w = word_ver(destino, "Nota importante", "Heading 1")
        st = w.get("estilos", {})
        n = st.get("Nota importante", {})
        anota("Word: «Nota importante» es un estilo propio con Arial, 14 pt, negrita, rojo oscuro, centrado, 6 y 12 pt", n.get("propio") and n.get("fuente") == "Arial" and n.get("tam") == 14 and n.get("b") == -1 and n.get("color") == "#C00000" and n.get("alin") == 1 and n.get("antes") == 6 and n.get("despues") == 12, n)
        h = st.get("Heading 1", {})
        anota("Word: «Título 1» con el verde y los 20 pt de NOM", h.get("color") == "#00B050" and h.get("tam") == 20, h)
        ps = w.get("parrafos", [])
        anota("Word: cada párrafo con su estilo", [x["estilo"] for x in ps[:4]] == ["Heading 1", "Nota importante", "Normal", "Heading 1"], [x["estilo"] for x in ps[:4]])
        # cómo lo pinta Word y cómo NOM
        remoto.subir([destino], "nomoffice/archivos")
        remoto.correr_office("verdad-est", "verdad.py", f"archivos\\{NOMBRE}", timeout=600)
        v = REAL / "verdad" / (NOMBRE[:-5] + ".verdad.json"); remoto.bajar(f"nomoffice/archivos/verdad/{v.name}", v)
        # abierto otra vez en NOM
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(destino.read_bytes()).decode(), destino.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
        otra = p.evaluate("() => ({ lista: NOM.editorActual.modoApi.estilos.lista().map((s) => s.t), p: NOM.textoRico.parrafosDe(NOM.editorActual.doc).map((q) => q.pa.estilo || '') })")
        anota("abierto otra vez en NOM: «Nota importante» sigue en la lista y en su párrafo", "Nota importante" in otra["lista"] and str(otra["p"][1]).startswith("u:"), otra)
        # un Word que ya existía
        w03 = REAL / "archivos/w03-estilos-listas.docx"
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(w03.read_bytes()).decode(), w03.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
        p.evaluate(f"() => {E}.guardar('h1', {{ color: '#7030A0' }})"); p.wait_for_timeout(500)
        mod = bajar(p, REAL / "w03-con-titulo1.docx")
        w2 = word_ver(mod, "Heading 1")
        anota("w03 con «Título 1» cambiado aquí: Word ve el morado en su definición", w2.get("estilos", {}).get("Heading 1", {}).get("color") == "#7030A0", w2.get("estilos"))
        import office_abre
        r = office_abre.abrir([destino, mod], "estilos")
        for a in (destino, mod): anota(f"Word abre {a.name} sin avisos ni reparar", office_abre.bien(r.get(a.name)), json.dumps(r.get(a.name), ensure_ascii=False)[:200])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    sal = Path.home() / "NOM/trabajo/estilos-real"
    r = subprocess.run([sys.executable, str(AQUI / "comparar_real.py"), str(REAL / "archivos" / NOMBRE), "--salida", str(sal)], capture_output=True, text=True)
    linea = next((x for x in r.stdout.splitlines() if x.startswith(NOMBRE)), r.stdout[-300:] + r.stderr[-300:])
    rep = json.loads((sal / "resumen.json").read_text(encoding="utf-8"))[0] if (sal / "resumen.json").exists() else {}
    anota(f"NOM pinta los renglones como Word: {linea}", rep.get("pct_iguales", 0) >= 99, linea)
    print(f"\nEstilos contra Word 2024: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

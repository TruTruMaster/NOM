#!/usr/bin/env python3
"""Tabla de contenido y tabla de ilustraciones de NOM contra Word 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/referencias_real.py

1. Un documento de NOM con una tabla de contenido, una tabla de ilustraciones, títulos 1 a 3, tres imágenes con pie y un
   índice alfabético (marcas XE con entradas, subentradas y «Véase») en varias páginas: NOM pinta cada entrada con su página.
2. Descargado, Word ve las dos tablas con las mismas entradas y números (windows/referencias_ver.py) y, después de
   actualizarlas Word mismo, quedan igual (los números de página de NOM son los de Word); los pies llevan su campo SEQ.
3. Word abre el archivo sin avisos.
"""
import base64, io, json, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import nomprueba as N  # noqa: E402

SAL = Path.home() / "NOM/trabajo/referencias-real"
res = []
LOREM = ("Este párrafo cuenta cómo se hizo el trabajo, con datos del trimestre y lo que se encontró en cada zona. "
         "Se repite para llenar la página y que los títulos caigan en páginas distintas, como en un reporte de verdad. ")


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:900]}"), flush=True)


def png(w, h, color):
    from PIL import Image
    im = Image.new("RGB", (w, h), color); b = io.BytesIO(); im.save(b, "PNG")
    return "data:image/png;base64," + base64.b64encode(b.getvalue()).decode()


def word_ver(docx):
    import remoto
    remoto.subir([docx], "nomoffice/entrada")
    remoto.correr_office("ref-" + docx.stem, "referencias_ver.py", f"entrada\\{docx.name}", f"{docx.stem}.ref.json", f"{docx.stem}.word.docx", timeout=300)
    tmp = SAL / f"{docx.stem}.ref.json"; remoto.bajar(f"nomoffice/{docx.stem}.ref.json", tmp)
    try: remoto.bajar(f"nomoffice/{docx.stem}.word.docx", SAL / f"{docx.stem}.word.docx")   # (actualizado por Word)
    except Exception: pass
    return json.loads(tmp.read_text(encoding="utf-8"))


def main():
    SAL.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Referencias")
    nav = N.Navegador(movil=False); nav.entrar(a); p = nav.p
    try:
        d = N.documento(a["token"], esp, "TEXTO", "Reporte con referencias")
        p.goto(N.B + "/#/doc/" + d); p.wait_for_selector(".wd [contenteditable=true]", timeout=30000); p.wait_for_timeout(600)
        p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.textoApi", timeout=30000)
        p.locator(".wd [contenteditable=true]").first.click(); p.wait_for_timeout(200)
        A = "NOM.editorActual.modoApi.textoApi"
        par = lambda t, e=None: {"tramos": [{"x": t}], "a": {"estilo": e} if e else {}}
        p.evaluate(f"() => {A}.insertarBloque({{ tipo: 'indice' }})"); p.wait_for_timeout(200)
        p.evaluate(f"() => {A}.insertarBloque({{ tipo: 'indice', modo: 'ilus' }})"); p.wait_for_timeout(200)
        p.evaluate(f"() => {A}.insertarBloque({{ tipo: 'salto' }})"); p.wait_for_timeout(200)
        pies = []
        def seccion(titulo, estilo, n, pie=None, color=(40, 120, 200), salto=False):
            p.evaluate(f"(ps) => {A}.insertarParrafos(ps)", [par(titulo, estilo)] + [par(LOREM * 3) for _ in range(n)]); p.wait_for_timeout(250)
            if pie:
                p.evaluate(f"(s) => {A}.insertarImagen({{ src: s, w: 300, h: 160 }})", png(300, 160, color)); p.wait_for_timeout(400)
                p.evaluate("(t) => { const d = NOM.editorActual.doc, im = d.vivos().filter((b) => b.tipo === 'imagen' && !b.pie).pop(); if (im) d.cambiar(im.id, { pie: t }); }", pie)
                pies.append(pie); p.wait_for_timeout(250)
                p.evaluate(f"(ps) => {A}.insertarParrafos(ps)", [par("Después de la figura.")]); p.wait_for_timeout(200)
            if salto: p.evaluate(f"() => {A}.insertarBloque({{ tipo: 'salto' }})"); p.wait_for_timeout(200)
        seccion("Introducción", "h1", 3, "Mapa de la zona norte")
        seccion("Antecedentes", "h2", 4, "Ventas por mes", (200, 90, 40))
        seccion("Resultados", "h1", 2, salto=True)
        seccion("Detalle por zona", "h2", 5, "Tabla resumen", (60, 160, 90))
        seccion("Notas finales", "h3", 3)
        # índice alfabético: marcas XE en varias páginas (entradas y subentradas) y el índice al final
        p.evaluate(f"() => {A}.insertarBloque({{ tipo: 'salto' }})"); p.wait_for_timeout(200)
        marcas = [("energía", 'XE "energía"'), ("agua potable", 'XE "agua:potable"'), ("árboles", 'XE "Árbol"'), ("agua residual", 'XE "agua:residual"'), ("zona norte", 'XE "zona"'), ("energía solar", 'XE "energía"'), ("ver bosque", 'XE "bosque" \\t "Véase Árbol"')]
        for k, (txt, instr) in enumerate(marcas):
            p.evaluate(f"(ps) => {A}.insertarParrafos(ps)", [par(f"Sobre {txt}: " + LOREM * 2)] + ([par(LOREM * 3) for _ in range(3)] if k % 2 else [])); p.wait_for_timeout(200)
            p.evaluate(f"(i) => {A}.marcarEntrada(i)", instr); p.wait_for_timeout(200)
        p.evaluate(f"(ps) => {A}.insertarParrafos(ps)", [par("Índice", "h1")]); p.wait_for_timeout(200)
        p.evaluate(f"() => {A}.insertarBloque({{ tipo: 'indice', modo: 'alfa', cols: 2 }})"); p.wait_for_timeout(200)
        p.wait_for_timeout(2500)
        vista = p.evaluate("""() => [...document.querySelectorAll('.wd-indice')].map((o) => [...o.querySelectorAll('.wd-ind-e')].map((a) => a.querySelector('.wd-ind-tx').textContent + '\\t' + a.querySelector('.wd-ind-pg').textContent))""")
        npag = p.evaluate("() => document.querySelectorAll('.wd-pagina').length")
        print("   NOM:", vista, npag, "páginas")
        cont, ilus = (vista + [[], []])[:2]
        alfa = [x for x in p.evaluate("() => [...document.querySelectorAll('.wd-ind-alfa p')].map((x) => x.textContent)")]
        print("   NOM índice:", alfa)
        anota("NOM: el índice alfabético con sus entradas, subentradas y páginas", len(alfa) == 7 and alfa[0].startswith("agua") and alfa[1].startswith("potable, "), alfa)
        anota("NOM: la tabla de contenido con los 6 títulos y su página", len(cont) == 6 and all(x.split("\t")[1].isdigit() for x in cont), cont)
        anota("NOM: la tabla de ilustraciones con las 3 figuras («Figura 1. …») y su página", [x.split("\t")[0] for x in ilus] == [f"Figura {i + 1}. {t}" for i, t in enumerate(pies)] and all(x.split("\t")[1].isdigit() for x in ilus), ilus)
        destino = SAL / "w21-referencias.docx"
        with p.expect_download(timeout=120000) as dl:
            p.evaluate("(n) => NOM.formatos.exportarListo(NOM.editorActual.doc, 'docx', n)", destino.stem)
        dl.value.save_as(destino)
        w = word_ver(destino)
        wc = (w.get("contenido") or [{}])[0]; wi = (w.get("ilustraciones") or [{}])[0]
        print("   Word:", w.get("paginas"), "páginas;", wc, wi)
        anota("Word ve la tabla de contenido como la escribió NOM", wc.get("antes") == cont, [wc.get("antes"), cont])
        anota("actualizada por Word queda igual (mismos títulos y páginas)", wc.get("despues") == cont, [wc.get("despues"), cont])
        anota("Word ve la tabla de ilustraciones como la escribió NOM", wi.get("antes") == ilus, [wi.get("antes"), ilus])
        anota("actualizada por Word queda igual (los pies llevan SEQ Figura)", wi.get("despues") == ilus and len(w.get("seq", [])) == 3, [wi.get("despues"), ilus, w.get("seq")])
        wx = (w.get("indices") or [{}])[0]
        print("   Word índice:", wx, w.get("secciones"))
        anota("Word ve el índice alfabético como lo escribió NOM", wx.get("antes") == alfa, [wx.get("antes"), alfa])
        anota("actualizado por Word queda igual (mismas entradas, subentradas y páginas)", wx.get("despues") == alfa, [wx.get("despues"), alfa])
        anota("el índice va en su sección continua de 2 columnas, como lo deja Word", any(x["cols"] == 2 for x in w.get("secciones", [])), w.get("secciones"))
        anota(f"las mismas páginas que Word ({npag}), también después de actualizar todo", w.get("paginas") == npag and w.get("paginas_despues") == npag, [w.get("paginas"), w.get("paginas_despues"), npag])
        import office_abre
        r = office_abre.abrir([destino], "referencias")
        anota("Word abre el archivo sin avisos ni reparar", office_abre.bien(r.get(destino.name)), json.dumps(r.get(destino.name), ensure_ascii=False)[:300])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nTablas de contenido e ilustraciones contra Word 2024: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

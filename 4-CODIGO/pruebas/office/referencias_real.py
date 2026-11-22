#!/usr/bin/env python3
"""Tabla de contenido y tabla de ilustraciones de NOM contra Word 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/referencias_real.py

1. Un documento de NOM con una tabla de contenido, una tabla de ilustraciones, títulos 1 a 3 y tres imágenes con pie en
   varias páginas: NOM pinta cada entrada con su número de página.
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
    remoto.correr_office("ref-" + docx.stem, "referencias_ver.py", f"entrada\\{docx.name}", f"{docx.stem}.ref.json", timeout=300)
    tmp = SAL / f"{docx.stem}.ref.json"; remoto.bajar(f"nomoffice/{docx.stem}.ref.json", tmp)
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
        p.wait_for_timeout(2500)
        vista = p.evaluate("""() => [...document.querySelectorAll('.wd-indice')].map((o) => [...o.querySelectorAll('.wd-ind-e')].map((a) => a.querySelector('.wd-ind-tx').textContent + '\\t' + a.querySelector('.wd-ind-pg').textContent))""")
        npag = p.evaluate("() => document.querySelectorAll('.wd-pagina').length")
        print("   NOM:", vista, npag, "páginas")
        cont, ilus = (vista + [[], []])[:2]
        anota("NOM: la tabla de contenido con los 5 títulos y su página", len(cont) == 5 and all(x.split("\t")[1].isdigit() for x in cont), cont)
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
        anota(f"las mismas páginas que Word ({npag})", w.get("paginas") == npag, [w.get("paginas"), npag])
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

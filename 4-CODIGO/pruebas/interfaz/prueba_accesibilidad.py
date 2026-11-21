#!/usr/bin/env python3
"""Comprobar accesibilidad (Revisar), con las reglas del comprobador de Office, en un Word, un Excel y un PowerPoint hechos
a propósito (python-docx, openpyxl y python-pptx) con problemas conocidos, y lo que no debe salir:

  Word: imagen sin texto alternativo (otra con él, no), tabla sin fila de encabezado (otra con «repetir como encabezado»,
        no), Título 1 → Título 3, tres renglones vacíos seguidos, espacios para acomodar, un vínculo «haz clic aquí» (otro
        con texto claro, no), letra gris claro y el Título 3 azul claro de la plantilla. La imagen con texto alternativo
        conserva el suyo al abrirse.
  Excel: «Sheet1», celdas combinadas y letra gris claro («Ventas», no).
  PowerPoint: una diapositiva sin título, dos con el mismo título, una imagen sin texto alternativo y una tabla sin
        encabezado (otra con él, no).
Además el cuadro abre, «Ir» lleva al párrafo, y un documento sin problemas lo dice.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_accesibilidad.py [capturas]
"""
import base64, io, sys, tempfile
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "office"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAFElEQVR4nGP4z8CAFWEXHbQSDAwAPaMI+K6hSU8AAAAASUVORK5CYII=")
res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:600]}"), flush=True)


def hacer_docx(d):
    import docx
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import RGBColor, Inches
    doc = docx.Document()
    doc.add_heading("Informe", level=1)
    doc.add_heading("Detalle", level=3)
    img = d / "p.png"; img.write_bytes(PNG)
    doc.add_paragraph().add_run().add_picture(str(img), width=Inches(1))
    doc.add_paragraph().add_run().add_picture(str(img), width=Inches(1))
    doc.inline_shapes[1]._inline.docPr.set("descr", "Gráfica de ventas de 2026")
    for _ in range(3): doc.add_paragraph("")
    doc.add_paragraph("Nombre     Apellido")
    def vinculo(texto):
        p = doc.add_paragraph("Para más información, ")
        rid = doc.part.relate_to("https://ejemplo.mx", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), rid)
        r = OxmlElement("w:r"); t = OxmlElement("w:t"); t.text = texto; r.append(t); h.append(r); p._p.append(h)
    vinculo("haz clic aquí")
    vinculo("la guía de inscripción")
    p = doc.add_paragraph(); r = p.add_run("Texto gris muy claro"); r.font.color.rgb = RGBColor(0xC8, 0xC8, 0xC8)
    t1 = doc.add_table(rows=2, cols=2); t1.cell(0, 0).text = "a"
    t2 = doc.add_table(rows=2, cols=2); t2.cell(0, 0).text = "Mes"
    trPr = t2.rows[0]._tr.get_or_add_trPr(); trPr.append(OxmlElement("w:tblHeader"))
    doc.add_paragraph("Fin.")
    f = d / "acc.docx"; doc.save(f); return f


def hacer_xlsx(d):
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook(); h = wb.active; h.title = "Sheet1"
    h["A1"] = "Título combinado"; h.merge_cells("A1:B1"); h["A2"] = "gris"; h["A2"].font = Font(color="FFD0D0D0")
    v = wb.create_sheet("Ventas"); v["A1"] = "bien"; v["A2"] = "negro"; v["A2"].font = Font(color="FF000000")
    f = d / "acc.xlsx"; wb.save(f); return f


def hacer_pptx(d):
    from pptx import Presentation
    from pptx.util import Inches
    pr = Presentation()
    s1 = pr.slides.add_slide(pr.slide_layouts[5]); s1.shapes.title.text = "Resultados"
    pic = s1.shapes.add_picture(io.BytesIO(PNG), Inches(1), Inches(2), Inches(1), Inches(1))
    pic._element.nvPicPr.cNvPr.set("descr", "")
    s2 = pr.slides.add_slide(pr.slide_layouts[6]); s2.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text_frame.text = "Sin título"
    s3 = pr.slides.add_slide(pr.slide_layouts[5]); s3.shapes.title.text = "Resultados"
    t = s3.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1)).table; t.first_row = False; t.cell(0, 0).text = "x"
    t2 = s3.shapes.add_table(2, 2, Inches(1), Inches(4), Inches(4), Inches(1)).table; t2.cell(0, 0).text = "Mes"
    f = d / "acc.pptx"; pr.save(f); return f


ESPERADO = {
    "acc.docx": {("error", "Falta texto alternativo"): 1, ("error", "Falta el encabezado de la tabla"): 1, ("consejo", "Se salta un nivel de título"): 1,
                 ("advertencia", "Caracteres en blanco repetidos (renglones vacíos para acomodar)"): 1, ("advertencia", "Caracteres en blanco repetidos (espacios para acomodar)"): 1,
                 ("advertencia", "Texto de vínculo poco claro"): 1,
                 # (la letra gris y el Título 3 de la plantilla de python-docx: #4F81BD da 4.0 a 1 contra el blanco)
                 ("advertencia", "Texto con poco contraste"): 2},
    "acc.xlsx": {("consejo", "Nombre de hoja predeterminado"): 1, ("advertencia", "Celdas combinadas"): 1, ("advertencia", "Texto con poco contraste"): 1},
    "acc.pptx": {("error", "Falta el título de la diapositiva"): 1, ("consejo", "Diapositivas con el mismo título"): 1, ("error", "Falta texto alternativo"): 1, ("error", "Falta el encabezado de la tabla"): 1},
}


def main():
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    try:
        with tempfile.TemporaryDirectory() as td:
            d = Path(td)
            for arch in (hacer_docx(d), hacer_xlsx(d), hacer_pptx(d)):
                p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
                p.evaluate(C.JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
                p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
                lista = p.evaluate("() => NOM.accesibilidad.revisar(NOM.editorActual.doc).map((x) => [x.nivel, x.regla.replace(/ \\(\\d[\\d.]* a 1\\)$/, ''), x.donde])")
                hallado = {}
                for nv, regla, _ in lista: hallado[(nv, regla)] = hallado.get((nv, regla), 0) + 1
                exp = ESPERADO[arch.name]
                anota(f"{arch.name}: lo que se esperaba y nada más", hallado == exp, f"NOM {lista} · esperado {exp}")
                if arch.suffix == ".docx":
                    p.evaluate("() => NOM.accesibilidad.abrir(NOM.editorActual)"); p.wait_for_timeout(500)
                    dlg = p.locator("section.hoja[role=dialog]").last
                    anota("…el cuadro lo muestra por grupos (errores, advertencias, consejos)", dlg.locator("[data-ir]").count() == len(lista) and "Errores" in dlg.inner_text(), dlg.inner_text()[:200])
                    dlg.locator("[data-ir]").first.click(); p.wait_for_timeout(700)
                    anota("…y «Ir» lleva al párrafo", p.evaluate("() => !!document.querySelector('.acc-marca')"))
                    dlg.locator("[data-no]").click(); p.wait_for_timeout(300)
        # un documento sin problemas
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'Limpio', modo: 'TEXTO', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function(U.JS_LISTO, timeout=60000); p.wait_for_timeout(1000)
        p.locator(".wd [contenteditable=true]").first.click(); p.keyboard.type("Un párrafo claro y sencillo."); p.wait_for_timeout(500)
        p.evaluate("() => NOM.accesibilidad.abrir(NOM.editorActual)"); p.wait_for_timeout(400)
        anota("un documento sin problemas lo dice", "No se encontraron problemas" in p.locator("section.hoja[role=dialog]").last.inner_text())
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nAccesibilidad: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

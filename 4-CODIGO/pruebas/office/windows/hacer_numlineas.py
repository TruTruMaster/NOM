"""w16-numlineas.docx: números de línea de Word (Disposición → Números de línea) para medir dónde y cómo los pone.

Cuatro secciones: (1) reiniciar en cada página, de 1 en 1, con un párrafo vacío, uno con «Suprimir números de línea»,
una tabla, un cuadro de texto, una nota al pie y encabezado; (2) cada 5, empezando en 3, a 0.5" del texto, continuo;
(3) reiniciar en cada sección, en dos columnas; (4) sin números. Lo mide verdad.py y el PDF que exporta Word.
"""
import os
from comun import word, ARCHIVOS, TEXTOS

w = word()
d = w.Documents.Add()
try:
    d.PageSetup.PageWidth, d.PageSetup.PageHeight = 612, 792
    for lado in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"): setattr(d.PageSetup, lado, 72)
    estilo = d.Styles(-1)
    estilo.ParagraphFormat.SpaceAfter = 6
    r = d.Content

    def parrafo(texto):
        d.Content.InsertParagraphAfter()
        p = d.Paragraphs(d.Paragraphs.Count)
        p.Range.Text = texto
        p = d.Paragraphs(d.Paragraphs.Count)
        p.NoLineNumber = False          # (el párrafo nuevo copia el formato del de arriba)
        return p

    def seccion():
        d.Content.InsertParagraphAfter()
        d.Paragraphs(d.Paragraphs.Count).Range.InsertBreak(2)   # salto de sección, página siguiente

    d.Paragraphs(1).Range.Text = "Números de línea que se reinician en cada página."
    for i in range(6): parrafo(TEXTOS[i % len(TEXTOS)])
    parrafo("")                                   # párrafo vacío
    p = parrafo("Este párrafo tiene «Suprimir números de línea»: " + TEXTOS[1])
    p.NoLineNumber = True
    parrafo("Antes de la tabla.")
    t = d.Tables.Add(d.Paragraphs(d.Paragraphs.Count).Range, 2, 2)
    for f in (1, 2):
        for c in (1, 2): t.Cell(f, c).Range.Text = f"celda {f}.{c} con texto"
    parrafo("Después de la tabla, con una nota al pie.")
    pn = d.Paragraphs(d.Paragraphs.Count).Range
    d.Footnotes.Add(Range=d.Range(pn.End - 1, pn.End - 1), Text="Texto de la nota al pie.")
    caja = d.Shapes.AddTextbox(1, 300, 400, 180, 60, d.Paragraphs(d.Paragraphs.Count).Range)
    caja.TextFrame.TextRange.Text = "Cuadro de texto con dos renglones de texto adentro."
    for i in range(14): parrafo(TEXTOS[(i + 2) % len(TEXTOS)])
    d.Sections(1).Headers(1).Range.Text = "Encabezado de la sección uno"
    ln = d.Sections(1).PageSetup.LineNumbering
    ln.Active = True; ln.RestartMode = 0; ln.CountBy = 1; ln.StartingNumber = 1   # 0 = cada página
    seccion()
    parrafo("Sección dos: cada 5, empieza en 3, a media pulgada, continuo.")
    for i in range(12): parrafo(TEXTOS[(i + 4) % len(TEXTOS)])
    ln = d.Sections(2).PageSetup.LineNumbering
    ln.Active = True; ln.RestartMode = 2; ln.CountBy = 5; ln.StartingNumber = 3; ln.DistanceFromText = 36   # 2 = continuo
    seccion()
    parrafo("Sección tres: se reinicia en la sección, en dos columnas.")
    for i in range(10): parrafo(TEXTOS[(i + 1) % len(TEXTOS)])
    d.Sections(3).PageSetup.TextColumns.SetCount(2)
    ln = d.Sections(3).PageSetup.LineNumbering
    ln.Active = True; ln.RestartMode = 1; ln.CountBy = 1; ln.StartingNumber = 1   # 1 = cada sección
    seccion()
    parrafo("Sección cuatro: sin números de línea.")
    for i in range(3): parrafo(TEXTOS[i % len(TEXTOS)])
    d.Sections(4).PageSetup.LineNumbering.Active = False
    d.Sections(4).PageSetup.TextColumns.SetCount(1)
    salida = os.path.join(ARCHIVOS, "w16-numlineas.docx")
    d.SaveAs2(salida, 16)
    d.ExportAsFixedFormat(os.path.join(ARCHIVOS, "w16-numlineas.pdf"), 17)
    print("hecho", d.ComputeStatistics(2), "páginas", d.Sections.Count, "secciones", flush=True)
finally:
    d.Close(0)
    w.Quit()

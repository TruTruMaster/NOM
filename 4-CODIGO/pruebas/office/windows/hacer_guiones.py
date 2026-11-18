"""w13-guiones.docx: la división automática de palabras de Word, en español (es-MX) y en inglés (en-US), con texto
justificado y a la izquierda, en columnas angostas (muchos cortes), con la zona de división de siempre (0.25") y una más
chica. Para medir con verdad.py y comparar_real.py."""
import os
from comun import word, ARCHIVOS

ES = ("La colaboración en tiempo real permite que varias personas editen el mismo documento sin perder ningún cambio, "
      "incluso cuando la conexión falla. Cada sección del documento incluye tablas, imágenes y referencias que deben "
      "conservarse exactamente igual al abrirlo en otro programa. Según los resultados de la encuesta aplicada a "
      "doscientos cuarenta y ocho estudiantes, la mayoría prefiere trabajar en equipo cuando las herramientas son "
      "sencillas, transparentes y extraordinariamente confiables para la administración universitaria.")
EN = ("Real-time collaboration allows several people to edit the same document without losing any change, even when the "
      "connection fails. Each section of the document includes tables, images and references that must be preserved "
      "exactly when opened in another application. According to the results of the survey, most students prefer "
      "working in teams when the tools are simple, transparent and extraordinarily reliable for university administration.")

w = word()
d = w.Documents.Add()
try:
    d.AutoHyphenation = True
    d.HyphenationZone = 18
    d.HyphenateCaps = True
    d.Content.Font.Name = "Calibri"
    d.Content.Font.Size = 11
    r = d.Content
    for k, (texto, lang) in enumerate([(ES, 2058), (EN, 1033), (ES, 2058), (EN, 1033), (ES, 2058), (EN, 1033)]):
        r = d.Content
        r.InsertParagraphAfter()
        p = d.Paragraphs(d.Paragraphs.Count)
        p.Range.Text = texto
        p = d.Paragraphs(d.Paragraphs.Count)
        p.Range.LanguageID = lang
        p.Alignment = 3 if k % 2 == 0 else 0          # justificado / izquierda
        p.RightIndent = [180, 220, 260, 120, 200, 240][k]   # renglones angostos: muchos cortes
    d.Paragraphs(1).Range.Delete()
    salida = os.path.join(ARCHIVOS, "w13-guiones.docx")
    d.SaveAs2(salida, 16)
    print("hecho", d.ComputeStatistics(2), "páginas", flush=True)
finally:
    d.Close(0)
    w.Quit()

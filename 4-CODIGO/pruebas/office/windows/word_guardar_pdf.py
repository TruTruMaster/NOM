"""Abre un .docx en Word, lo guarda otra vez como .docx de Word (con sus propiedades: autor, título, sin comentarios del
programa que lo generó) y lo exporta a PDF con las fuentes reales y las etiquetas de accesibilidad. El nombre de usuario
de Word se pone solo mientras se guarda y después se deja como estaba.
Uso: word_guardar_pdf.py DATOS.json, con {"entrada", "docx", "pdf", "salida", "autor", "titulo"} (rutas relativas a la
carpeta nomoffice; así el autor y el título, con espacios y acentos, no pasan por la línea de órdenes)."""
import json, os, sys
from comun import word

p = json.load(open(sys.argv[1], encoding="utf-8"))
entrada, docx, pdf, salida = (os.path.abspath(p[k]) for k in ("entrada", "docx", "pdf", "salida"))
autor, titulo = p["autor"], p.get("titulo")
w = word()
datos = {"abrio": False}
antes = (w.UserName, w.UserInitials)
try:
    w.UserName = autor
    w.UserInitials = "".join(x[0] for x in autor.split()[:3]).upper()
    d = w.Documents.Open(entrada, False, False, False, "", "", False, "", "", 0, 0, True, False, 0, True)
    try:
        datos["abrio"] = True
        props = d.BuiltInDocumentProperties
        props("Author").Value = autor
        props("Comments").Value = ""
        props("Subject").Value = ""
        props("Keywords").Value = ""
        if titulo: props("Title").Value = titulo
        d.Repaginate()
        datos["paginas"] = d.ComputeStatistics(2)
        d.SaveAs2(docx, 16)
        # ExportAsFixedFormat(ruta, PDF, abrir, optimizar para impresión, todo, desde, hasta, documento, propiedades,
        # IRM, marcadores de títulos, etiquetas de accesibilidad, mapa de bits si falta letra, PDF/A)
        d.ExportAsFixedFormat(pdf, 17, False, 0, 0, 1, 1, 0, True, True, 1, True, True, False)
        datos["docx"] = os.path.exists(docx)
        datos["pdf"] = os.path.exists(pdf)
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    try: w.UserName, w.UserInitials = antes
    except Exception: pass
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False)
print("hecho", datos, flush=True)

"""Abre un .docx en Word y lo guarda como PDF (como «Guardar como PDF»), con las propiedades del documento y los
títulos como marcadores. Dice si Word pidió reparar el archivo y cuántas páginas tiene.
Uso: word_a_pdf.py ENTRADA.docx SALIDA.pdf SALIDA.json"""
import json, os, sys
from comun import word

entrada, pdf, salida = (os.path.abspath(x) for x in sys.argv[1:4])
w = word()
datos = {"abrio": False}
try:
    d = w.Documents.Open(entrada, False, True, False, "", "", False, "", "", 0, 0, True, False, 0, True)
    try:
        datos["abrio"] = True
        d.Repaginate()
        datos["paginas"] = d.ComputeStatistics(2)
        # ExportAsFixedFormat(ruta, PDF, abrir, optimizar para impresión, todo, desde, hasta, documento, propiedades,
        # IRM, marcadores de títulos, etiquetas de accesibilidad, mapa de bits si falta letra, PDF/A)
        d.ExportAsFixedFormat(pdf, 17, False, 0, 0, 1, 1, 0, True, True, 1, True, True, False)
        datos["pdf"] = os.path.exists(pdf)
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False)
print("hecho", datos, flush=True)

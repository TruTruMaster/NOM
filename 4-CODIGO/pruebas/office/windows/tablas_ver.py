"""Las tablas de un .docx como las lee Word: de cada una, sus celdas (texto). Uso: tablas_ver.py ENTRADA.docx SALIDA.json"""
import json, os, sys
from comun import word

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
w = word()
datos = {"abrio": False}
try:
    d = w.Documents.Open(entrada, False, True)
    try:
        datos["abrio"] = True
        datos["tablas"] = [[[t.Cell(i, j).Range.Text[:-2] for j in range(1, t.Columns.Count + 1)] for i in range(1, t.Rows.Count + 1)] for t in d.Tables]
        datos["parrafos"] = [p.Range.Text.rstrip("\r\x07") for p in d.Paragraphs]
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

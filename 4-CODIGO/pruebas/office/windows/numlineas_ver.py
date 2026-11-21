"""Abre un .docx en Word y deja su PDF y cómo quedaron los números de línea (por sección) y «Suprimir números de
línea» (por párrafo). Uso: numlineas_ver.py ENTRADA.docx SALIDA.json SALIDA.pdf"""
import json, os, sys, time
from comun import word

entrada, salida, pdf = (os.path.abspath(x) for x in sys.argv[1:4])
w = word()
datos = {"abrio": False}
try:
    d = w.Documents.Open(entrada, False, True, False, "", "", False, "", "", 0, 0, True, False, 0, True)
    try:
        datos["abrio"] = True
        d.ActiveWindow.View.Type = 3
        d.Repaginate()
        datos["secciones"] = []
        for s in d.Sections:
            ln = s.PageSetup.LineNumbering
            datos["secciones"].append({"activo": bool(ln.Active), "cada": ln.CountBy, "inicio": ln.StartingNumber, "distancia": round(ln.DistanceFromText, 2), "reinicio": ln.RestartMode})
        datos["suprimidos"] = [i + 1 for i, p in enumerate(d.Paragraphs) if p.NoLineNumber]
        datos["paginas"] = d.ComputeStatistics(2)
        d.ExportAsFixedFormat(pdf, 17)
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False)
print("hecho", datos.get("abrio"), flush=True)

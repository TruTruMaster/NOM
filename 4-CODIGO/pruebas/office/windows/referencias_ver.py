"""Las tablas de contenido y de ilustraciones de un .docx como las ve Word: el texto de cada entrada tal como viene en
el archivo y después de que Word las actualiza (Update), con su número de página; y cuántas páginas tiene.
Uso: referencias_ver.py ENTRADA.docx SALIDA.json"""
import json, os, sys
from comun import word

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
w = word()
datos = {"abrio": False}


def entradas(rg):
    return [x for x in rg.Text.replace("\x07", "").split("\r") if x.strip()]


try:
    d = w.Documents.Open(entrada, False, False)
    try:
        datos["abrio"] = True
        datos["paginas"] = d.ComputeStatistics(2)
        tocs = []
        for i in range(1, d.TablesOfContents.Count + 1):
            t = d.TablesOfContents(i)
            antes = entradas(t.Range); t.Update(); tocs.append({"antes": antes, "despues": entradas(d.TablesOfContents(i).Range)})
        datos["contenido"] = tocs
        tofs = []
        for i in range(1, d.TablesOfFigures.Count + 1):
            t = d.TablesOfFigures(i)
            antes = entradas(t.Range); t.Update(); tofs.append({"antes": antes, "despues": entradas(d.TablesOfFigures(i).Range)})
        datos["ilustraciones"] = tofs
        datos["seq"] = [f.Code.Text.strip() for f in d.Fields if f.Type == 12]
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

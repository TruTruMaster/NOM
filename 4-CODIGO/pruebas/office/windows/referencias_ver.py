"""Las tablas de contenido y de ilustraciones y los índices de un .docx como los ve Word: el texto de cada entrada tal como viene en
el archivo y después de que Word las actualiza (Update), con su número de página; y cuántas páginas tiene.
Uso: referencias_ver.py ENTRADA.docx SALIDA.json"""
import json, os, sys
from comun import word

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
w = word()
datos = {"abrio": False}


def entradas(rg):
    return [x.replace("\x0c", "") for x in rg.Text.replace("\x07", "").split("\r") if x.replace("\x0c", "").strip()]


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
        ixs = []
        for i in range(1, d.Indexes.Count + 1):
            t = d.Indexes(i)
            antes = entradas(t.Range); t.Update(); ixs.append({"antes": antes, "despues": entradas(d.Indexes(i).Range)})
        datos["indices"] = ixs
        datos["secciones"] = [{"cols": d.Sections(k).PageSetup.TextColumns.Count, "inicio": d.Sections(k).PageSetup.SectionStart,
                               "pags": [d.Sections(k).Range.Characters(1).Information(3), d.Sections(k).Range.Information(3)],
                               "fin": d.Sections(k).Range.Text[-60:]} for k in range(1, d.Sections.Count + 1)]
        datos["paginas_despues"] = d.ComputeStatistics(2)
        datos["seq"] = [f.Code.Text.strip() for f in d.Fields if f.Type == 12]
        if len(sys.argv) > 3: d.SaveAs2(os.path.abspath(sys.argv[3]), 16)   # (una copia ya actualizada por Word, para ver cómo la guarda)
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

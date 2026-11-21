"""Los minigráficos de un .xlsx como los lee Excel (Range.SparklineGroups): de cada hoja, cada grupo con su tipo (1 línea,
2 columna, 3 pérdidas y ganancias), su rango de datos, sus celdas, su color y qué puntos marca.
Uso: minigraficos_ver.py ENTRADA.xlsx SALIDA.json"""
import json, os, sys
from comun import excel

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
x = excel()
datos = {"abrio": False, "hojas": {}}
rgb = lambda c: "#%02X%02X%02X" % (c & 255, (c >> 8) & 255, (c >> 16) & 255) if isinstance(c, (int, float)) and 0 <= c < 2 ** 24 else None
try:
    wb = x.Workbooks.Open(entrada, 0, True)
    try:
        datos["abrio"] = True
        for h in wb.Worksheets:
            gs = h.Cells.SparklineGroups
            lista = []
            for i in range(1, gs.Count + 1):
                g = gs.Item(i)
                lista.append({"tipo": g.Type, "datos": g.SourceData, "celdas": g.Location.Address.replace("$", ""), "n": g.Count, "color": rgb(g.SeriesColor.Color),
                              "marcas": bool(g.Points.Markers.Visible), "alto": bool(g.Points.Highpoint.Visible), "bajo": bool(g.Points.Lowpoint.Visible),
                              "neg": bool(g.Points.Negative.Visible), "grosor": g.LineWeight})
            datos["hojas"][h.Name] = sorted(lista, key=lambda q: q["celdas"])
    finally:
        wb.Close(False)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    x.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

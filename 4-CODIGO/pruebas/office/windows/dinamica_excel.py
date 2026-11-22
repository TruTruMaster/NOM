"""Excel hace tablas dinámicas sobre los datos de la primera hoja de ENTRADA.xlsx (como Insertar → Tabla dinámica) y dice
lo que muestra cada celda (texto, valor, negrita, relleno y color de letra que se ven) de cada una.
Uso: dinamica_excel.py ENTRADA.xlsx SALIDA.json CASOS.json [COPIA.xlsx]
CASOS.json: [{ "id", "origen": "A1:E31", "destino": "H3", "filas": ["Región"…], "columnas": [...], "valores": [["Ventas",
"sum"|"count"|"average"|"max"|"min"]…], "filtro": "Producto" | null, "estilo": "PivotStyleLight16" }]. Cada caso en una hoja
nueva copiada de la primera (para que no se encimen)."""
import json, os, sys
from comun import excel

entrada, salida, casos = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2]), json.load(open(sys.argv[3], encoding="utf-8"))
FN = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139}
rgb = lambda c: (lambda n: "#%02X%02X%02X" % (n & 255, (n >> 8) & 255, (n >> 16) & 255))(int(c)) if isinstance(c, (int, float)) and 0 <= c < 2 ** 24 else None
x = excel()
datos = {"casos": {}}


def col(coleccion, k):
    try: return coleccion(k)
    except TypeError: return coleccion.Item(k)

try:
    wb = x.Workbooks.Open(entrada, 0, False)
    try:
        base = wb.Worksheets(1)
        for k, cs in enumerate(casos):
            base.Copy(None, wb.Worksheets(wb.Worksheets.Count))
            h = wb.Worksheets(wb.Worksheets.Count)
            h.Name = "c" + str(k + 1)
            try:
                pc = wb.PivotCaches().Create(1, h.Range(cs["origen"]), 8)
                pt = pc.CreatePivotTable(h.Range(cs["destino"]), "TD" + str(k + 1), True, 8)
                if cs.get("estilo"): pt.TableStyle2 = cs["estilo"]
                for i, f in enumerate(cs.get("filas", [])): pf = col(pt.PivotFields, f); pf.Orientation = 1; pf.Position = i + 1
                for i, f in enumerate(cs.get("columnas", [])): pf = col(pt.PivotFields, f); pf.Orientation = 2; pf.Position = i + 1
                if cs.get("filtro"): pf = col(pt.PivotFields, cs["filtro"]); pf.Orientation = 3
                for f, fn in cs.get("valores", []): pt.AddDataField(col(pt.PivotFields, f), None, FN[fn])
                rg = pt.TableRange2
                celdas = {}
                for c in rg:
                    d = c.DisplayFormat
                    v = c.Value
                    celdas[c.Address.replace("$", "")] = {"t": c.Text, "v": v if isinstance(v, (int, float, str)) or v is None else str(v), "b": bool(d.Font.Bold),
                                                         "f": rgb(d.Interior.Color), "k": rgb(d.Font.Color), "ind": c.IndentLevel, "fmt": c.NumberFormat}
                datos["casos"][cs["id"]] = {"rango": rg.Address.replace("$", ""), "rango1": pt.TableRange1.Address.replace("$", ""), "celdas": celdas,
                                            "datos": [col(pt.DataFields, i).Name for i in range(1, pt.DataFields.Count + 1)]}
            except Exception as e:
                datos["casos"][cs["id"]] = {"error": str(e)[:300]}
        if len(sys.argv) > 4: wb.SaveAs(os.path.abspath(sys.argv[4]), 51)   # (el libro con las tablas dinámicas, para ver cómo las guarda)
    finally:
        wb.Close(False)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    x.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", flush=True)

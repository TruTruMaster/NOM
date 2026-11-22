"""Las tablas con formato (ListObjects) de un .xlsx como las ve Excel: de cada una, su nombre, su rango, su estilo, sus
opciones (encabezado, totales, franjas, primera y última columna, filtro), los nombres de sus columnas y, de cada celda,
el relleno, el color de la letra y la negrita que se ven (DisplayFormat); y el valor y la fórmula de las CELDAS pedidas
(de la primera hoja). Uso: tablas_formato_ver.py ENTRADA.xlsx SALIDA.json [CELDA …]"""
import json, os, sys
from comun import excel

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
x = excel()
rgb = lambda c: (lambda n: "#%02X%02X%02X" % (n & 255, (n >> 8) & 255, (n >> 16) & 255))(int(c)) if isinstance(c, (int, float)) and 0 <= c < 2 ** 24 else None
datos = {"abrio": False, "hojas": {}}
try:
    wb = x.Workbooks.Open(entrada, 0, True)
    try:
        datos["abrio"] = True
        for h in wb.Worksheets:
            lista = []
            for i in range(1, h.ListObjects.Count + 1):
                lo = h.ListObjects(i)
                try: est = lo.TableStyle.Name
                except Exception: est = None
                celdas = {}
                for c in lo.Range:
                    d = c.DisplayFormat
                    celdas[c.Address.replace("$", "")] = [rgb(d.Interior.Color), rgb(d.Font.Color), bool(d.Font.Bold)]
                lista.append({"nombre": lo.Name, "rango": lo.Range.Address.replace("$", ""), "estilo": est, "enc": bool(lo.ShowHeaders), "tot": bool(lo.ShowTotals),
                              "rayasF": bool(lo.ShowTableStyleRowStripes), "rayasC": bool(lo.ShowTableStyleColumnStripes), "primeraC": bool(lo.ShowTableStyleFirstColumn),
                              "ultimaC": bool(lo.ShowTableStyleLastColumn), "filtro": bool(lo.ShowAutoFilter), "cols": [lo.ListColumns(k).Name for k in range(1, lo.ListColumns.Count + 1)],
                              "celdas": celdas})
            datos["hojas"][h.Name] = lista
        datos["valores"] = {c: [wb.Worksheets(1).Range(c).Value, wb.Worksheets(1).Range(c).Formula] for c in sys.argv[3:]}
    finally:
        wb.Close(False)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    x.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

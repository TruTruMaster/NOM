"""Las tablas dinámicas de un .xlsx como las ve Excel: de cada hoja, cada una con su nombre, su rango (con el filtro), sus
campos de filas, columnas, filtro y valores, y el texto de cada celda; y después de actualizarla Excel (RefreshTable),
otra vez el texto de cada celda. Uso: dinamicas_ver.py ENTRADA.xlsx SALIDA.json"""
import json, os, sys
from comun import excel

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
x = excel()
datos = {"abrio": False, "hojas": {}}


def col(coleccion, k):
    try: return coleccion(k)
    except TypeError: return coleccion.Item(k)


def nombres(c):
    try: return [col(c, i).Name for i in range(1, c.Count + 1)]
    except Exception: return []


try:
    wb = x.Workbooks.Open(entrada, 0, False)
    try:
        datos["abrio"] = True
        for h in wb.Worksheets:
            lista = []
            pts = h.PivotTables()
            for i in range(1, pts.Count + 1):
                pt = col(pts, i)
                antes = {c.Address.replace("$", ""): c.Text for c in pt.TableRange2}
                e = {"nombre": pt.Name, "rango": pt.TableRange2.Address.replace("$", ""), "filas": nombres(pt.RowFields), "columnas": nombres(pt.ColumnFields),
                     "filtros": nombres(pt.PageFields), "valores": nombres(pt.DataFields), "antes": antes}
                try:
                    pt.RefreshTable()
                    e["despues"] = {c.Address.replace("$", ""): c.Text for c in pt.TableRange2}
                except Exception as ex:
                    e["error_actualizar"] = str(ex)[:200]
                lista.append(e)
            datos["hojas"][h.Name] = lista
    finally:
        wb.Close(False)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    x.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

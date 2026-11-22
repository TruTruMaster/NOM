"""Con qué colores pinta Excel una tabla (Insertar → Tabla) en un libro: abre el .xlsx, hace una tabla con el estilo
pedido sobre RANGO y dice el color que se ve en cada celda (DisplayFormat) y los colores del tema del libro.
Uso: tabla_tema.py ENTRADA.xlsx SALIDA.json RANGO ESTILO"""
import json, os, sys
from comun import excel

entrada, salida, rango, estilo = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2]), sys.argv[3], sys.argv[4]
x = excel()
rgb = lambda c: (lambda n: "#%02X%02X%02X" % (n & 255, (n >> 8) & 255, (n >> 16) & 255))(int(c)) if isinstance(c, (int, float)) and 0 <= c < 2 ** 24 else None
datos = {}
try:
    wb = x.Workbooks.Open(entrada, 0, False)
    try:
        h = wb.Worksheets(1)
        lo = h.ListObjects.Add(1, h.Range(rango), None, 1)
        lo.TableStyle = estilo
        datos["celdas"] = {c.Address.replace("$", ""): {"fondo": rgb(c.DisplayFormat.Interior.Color), "letra": rgb(c.DisplayFormat.Font.Color), "b": bool(c.DisplayFormat.Font.Bold)} for c in h.Range(rango)}
        tema = wb.Theme.ThemeColorScheme
        datos["tema"] = [rgb(tema.Colors(i).RGB) for i in range(1, 13)]
    finally:
        wb.Close(False)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    x.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", flush=True)

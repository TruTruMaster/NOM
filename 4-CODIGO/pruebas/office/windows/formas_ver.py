"""Las formas e imágenes de cada hoja de un .xlsx como las ve Excel (Worksheet.Shapes): nombre, tipo (13 imagen, 3
gráfica…), celda donde empieza, ancho y alto en puntos y texto alternativo. Uso: formas_ver.py ENTRADA.xlsx SALIDA.json"""
import json, os, sys
from comun import excel

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
x = excel()
datos = {"abrio": False, "hojas": {}}
try:
    wb = x.Workbooks.Open(entrada, 0, True)
    try:
        datos["abrio"] = True
        for h in wb.Worksheets:
            out = []
            for i in range(1, h.Shapes.Count + 1):
                s = h.Shapes(i)
                try: alt = s.AlternativeText
                except Exception: alt = None
                out.append({"nombre": s.Name, "tipo": s.Type, "celda": s.TopLeftCell.Address.replace("$", ""), "ancho": round(s.Width, 1), "alto": round(s.Height, 1), "alt": alt})
            datos["hojas"][h.Name] = out
    finally:
        wb.Close(False)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    x.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

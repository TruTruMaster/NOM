"""Dónde parte Excel cada hoja de un .xlsx en páginas para imprimir: de cada hoja, su disposición de página (como la lee
Excel), cuántas páginas, y la fila y la columna donde empieza cada página (HPageBreaks/VPageBreaks, los automáticos y los
puestos a mano, con la vista de saltos de página). Uso: paginas_ver.py ENTRADA.xlsx SALIDA.json"""
import json, os, sys
from comun import excel

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
x = excel()
datos = {"abrio": False, "hojas": []}
try:
    wb = x.Workbooks.Open(entrada, 0, True)
    try:
        datos["abrio"] = True
        for h in wb.Worksheets:
            h.Activate()
            wb.Windows(1).View = 2   # xlPageBreakPreview: Excel calcula los saltos automáticos
            ps = h.PageSetup
            r = {"nombre": h.Name}
            try:
                r["setup"] = {"orientacion": ps.Orientation, "papel": ps.PaperSize, "zoom": ps.Zoom, "ancho": ps.FitToPagesWide, "alto": ps.FitToPagesTall,
                              "margenes": [ps.LeftMargin, ps.RightMargin, ps.TopMargin, ps.BottomMargin, ps.HeaderMargin, ps.FooterMargin],
                              "area": ps.PrintArea, "titulosF": ps.PrintTitleRows, "titulosC": ps.PrintTitleColumns, "orden": ps.Order,
                              "cuadricula": bool(ps.PrintGridlines), "encabezados": bool(ps.PrintHeadings)}
                r["paginas"] = ps.Pages.Count
            except Exception as e:
                r["error_setup"] = str(e)[:200]
            def saltos(col, que):
                out = []
                for i in range(1, col.Count + 1):
                    try:
                        b = col(i)
                        out.append({"en": b.Location.Row if que == "f" else b.Location.Column, "mano": b.Type == -4135})
                    except Exception as e:
                        out.append({"error": str(e)[:100]})
                return out
            r["saltosF"] = saltos(h.HPageBreaks, "f")
            r["saltosC"] = saltos(h.VPageBreaks, "c")
            u = h.UsedRange
            r["usado"] = u.Address.replace("$", "")
            datos["hojas"].append(r)
            wb.Windows(1).View = 1
    finally:
        wb.Close(False)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    x.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), len(datos["hojas"]), flush=True)

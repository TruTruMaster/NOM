"""Lo que ve Excel 2024 en un libro: cada celda usada (fórmula, valor y texto que se muestra) y los nombres definidos.
Para comparar con NOM lo que NOM exporta (nombres, cálculo manual, esquema…).

  python libro_ver.py ENTRADA.xlsx SALIDA.json
"""
import json, os, sys
from comun import excel, BASE

# (rutas relativas a la carpeta de estos programas: Excel abre con la suya)
ENT = sys.argv[1] if os.path.isabs(sys.argv[1]) else os.path.join(BASE, sys.argv[1])
SAL = sys.argv[2] if os.path.isabs(sys.argv[2]) else os.path.join(BASE, sys.argv[2])
def col(j):
    r = ""
    while j: j, k = divmod(j - 1, 26); r = chr(65 + k) + r
    return r


x = excel()
res = {"abrio": False}
try:
    wb = x.Workbooks.Open(ENT, UpdateLinks=0, ReadOnly=True)
    res["abrio"] = True
    x.Calculate()
    res["calculo"] = {-4105: "automatico", -4135: "manual", 2: "semiautomatico"}.get(x.Calculation, x.Calculation)
    res["nombres"] = [{"nombre": n.Name, "refiere": n.RefersTo, "visible": bool(n.Visible)} for n in wb.Names]
    hojas = []
    for h in wb.Worksheets:
        u = h.UsedRange
        celdas = {}
        for i in range(u.Row, u.Row + u.Rows.Count):
            for j in range(u.Column, u.Column + u.Columns.Count):
                c = h.Cells(i, j)
                if c.Formula in (None, ""): continue
                celdas[col(j) + str(i)] = {"f": c.Formula, "v": c.Value2, "t": c.Text}
        filas = {}
        for i in range(1, u.Row + u.Rows.Count + 1):
            r = h.Rows(i)
            if r.OutlineLevel > 1 or r.Hidden: filas[str(i)] = {"nivel": r.OutlineLevel, "oculta": bool(r.Hidden)}
        cols = {}
        for j in range(1, u.Column + u.Columns.Count + 1):
            c = h.Columns(j)
            if c.OutlineLevel > 1 or c.Hidden: cols[str(j)] = {"nivel": c.OutlineLevel, "oculta": bool(c.Hidden)}
        try: h.Activate(); vf = bool(wb.Windows(1).DisplayFormulas)
        except Exception: vf = None
        hojas.append({"nombre": h.Name, "celdas": celdas, "filas": filas, "columnas": cols, "formulas_visibles": vf})
    res["hojas"] = hojas
    wb.Close(False)
except Exception as e:
    res["error"] = str(getattr(e, "excepinfo", None) or e)[:400]
finally:
    x.Quit()
json.dump(res, open(SAL, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", res.get("abrio"), flush=True)

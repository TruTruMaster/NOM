"""Los pasos de «Buscar objetivo» de Excel 2024: el resultado con 1, 2, 3… iteraciones como máximo (Application.MaxIterations)
y con distintas tolerancias (MaxChange), para sacar el método que usa Excel (analisis/objetivo_excel.py).

  python objetivo_pasos.py SALIDA.json
"""
import json, sys
from comun import excel

CASOS = [
    ("x2", "=B1^2", 1, 2), ("x2b", "=B1^2", 3, 2), ("x2c", "=B1^2", 10, 50), ("x3", "=B1^3-B1", 2, 5),
    ("exp", "=EXP(B1)", 0, 10), ("log", "=LN(B1)", 1, 2), ("seno", "=SIN(B1)", 0.1, 0.5), ("lineal", "=2*B1+3", 1, 100),
    ("pago", "=PMT(B1/12,60,-100000)", 0.05, 2000), ("cero", "=B1^2", 0, 4), ("neg", "=B1^2", -1, 9), ("inv", "=1/B1", 1, 4),
]
x = excel()
wb = x.Workbooks.Add()
h = wb.Worksheets(1)
res = {}
try:
    for nombre, form, x0, meta in CASOS:
        filas = []
        for tol in (0.001, 1e-6):
            for k in list(range(1, 11)) + [100]:
                x.MaxIterations = k
                x.MaxChange = tol
                h.Range("B1").Value = x0
                h.Range("B2").Formula = form
                ok = h.Range("B2").GoalSeek(Goal=meta, ChangingCell=h.Range("B1"))
                filas.append({"tol": tol, "k": k, "ok": bool(ok), "x": h.Range("B1").Value2, "f": h.Range("B2").Value2})
        res[nombre] = {"formula": form, "x0": x0, "meta": meta, "pasos": filas}
        print(nombre, filas[-1], flush=True)
finally:
    x.MaxIterations = 100; x.MaxChange = 0.001
    wb.Close(False)
    x.Quit()
json.dump(res, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=0)

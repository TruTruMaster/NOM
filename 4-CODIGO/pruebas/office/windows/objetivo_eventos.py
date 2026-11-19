"""Cada valor que prueba «Buscar objetivo» de Excel 2024, para sacar su método (analisis/objetivo_excel.py).

Sin macros: con el cálculo iterativo encendido (una vuelta), la celda D1 se refiere a sí misma y le agrega el valor de la
celda que cambia (B1) y el de la fórmula (B2) cada vez que Excel calcula la hoja; Buscar objetivo calcula la hoja con
cada valor que prueba, así que D1 queda con la historia (15 cifras, lo que Excel escribe de un número). Además, con el
evento SheetCalculate se cuenta cuántas veces calculó.

  python objetivo_eventos.py SALIDA.json
"""
import json, sys
import pythoncom
import win32com.client as W
from win32com.client import gencache

CASOS = [
    ("x2", "=B1^2", 1, 2), ("x2b", "=B1^2", 3, 2), ("x2c", "=B1^2", 10, 50), ("x3", "=B1^3-B1", 2, 5),
    ("exp", "=EXP(B1)", 0, 10), ("log", "=LN(B1)", 1, 2), ("seno", "=SIN(B1)", 0.1, 0.5), ("lineal", "=2*B1+3", 1, 100),
    ("pago", "=PMT(B1/12,60,-100000)", 0.05, 2000), ("cero", "=B1^2", 0, 4), ("neg", "=B1^2", -1, 9), ("inv", "=1/B1", 1, 4),
    ("imposible", "=B1^2", 3, -1), ("grande", "=B1*1000000", 1, 5), ("chico", "=B1/1000000", 1, 5),
]
cuenta = [0]


class Eventos:
    def OnSheetCalculate(self, sh):
        cuenta[0] += 1


x = W.DispatchWithEvents(gencache.EnsureDispatch(W.DispatchEx("Excel.Application")), Eventos)
x.Visible = False; x.DisplayAlerts = False
wb = x.Workbooks.Add()
h = wb.Worksheets(1)
res = {}
try:
    for nombre, form, x0, meta in CASOS:
        for tol in (0.001, 1e-6):
            x.Iteration = False
            h.Range("D1").Value = ""
            h.Range("B1").Value = x0
            x.MaxChange = tol
            x.Iteration = True
            x.MaxIterations = 1
            # (D1 es precedente de la fórmula: Buscar objetivo calcula los precedentes de su celda con cada valor que prueba)
            h.Range("D1").Formula = '=D1&";"&TEXT(B1,"0.00000000000000E+00")'
            h.Range("B2").Formula = form + "+0*LEN(D1)"
            antes = h.Range("D1").Value2 or ""
            pythoncom.PumpWaitingMessages(); cuenta[0] = 0
            ok = h.Range("B2").GoalSeek(Goal=meta, ChangingCell=h.Range("B1"))
            pythoncom.PumpWaitingMessages()
            hist = (h.Range("D1").Value2 or "")[len(antes):]
            pasos = []
            for t in hist.split(";"):
                try: pasos.append(float(t))
                except ValueError: pass
            res[f"{nombre}|{tol}"] = {"formula": form, "x0": x0, "meta": meta, "tol": tol, "ok": bool(ok), "x": h.Range("B1").Value2,
                                      "f": h.Range("B2").Value2, "pasos": pasos, "calculos": cuenta[0]}
            print(nombre, tol, len(pasos), "pasos,", cuenta[0], "cálculos", flush=True)
finally:
    x.Iteration = False; x.MaxIterations = 100; x.MaxChange = 0.001
    wb.Close(False)
    x.Quit()
json.dump(res, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=0)

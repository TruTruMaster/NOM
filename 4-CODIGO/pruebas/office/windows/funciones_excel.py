"""Calcula fórmulas en Excel de verdad: escribe los datos y cada fórmula (con Formula2, como si se tecleara en
Excel 2024: matrices dinámicas incluidas), recalcula todo y guarda el valor, el texto que muestra la celda y el error
exacto. Lo usa pruebas/office/funciones_excel.py.

  correr.ps1 funciones funciones_excel.py <entrada.json: {datos, casos}> <salida.json>
"""
import json, os, sys
from comun import excel

base = os.path.join(os.path.expanduser("~"), "nomoffice")
ent, sal = [a if os.path.isabs(a) else os.path.join(base, a) for a in sys.argv[1:3]]
e = json.load(open(ent, encoding="utf8"))
ERRORES = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!", -2146826252: "#NUM!",
           -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826238: "#GETTING_DATA", -2146826245: "#SPILL!", -2146826243: "#CALC!",
           -2146826236: "#FIELD!", -2146826235: "#BLOCKED!", -2146826234: "#CONNECT!", -2146826233: "#UNKNOWN!", -2146826229: "#BUSY!"}
x = excel()
wb = x.Workbooks.Add()
try:
    h = wb.Worksheets(1)
    for k, v in e["datos"].items():
        h.Range(k).Value2 = v
    # (cada fórmula en su renglón de la columna J, lejos de los datos: lo que se derrame cabe a la derecha y abajo no estorba
    # porque va de 20 en 20 renglones)
    paso = int(e.get("paso", 20))
    for i, f in enumerate(e["casos"]):
        c = h.Cells(1 + i * paso, 10)
        try: c.Formula2 = "=" + f
        except Exception as ex: c.Value2 = "#ERROR_AL_ESCRIBIR " + str(ex)[:80]
    x.CalculateFull()
    r = []
    for i, f in enumerate(e["casos"]):
        c = h.Cells(1 + i * paso, 10)
        v = c.Value2
        o = {"texto": c.Text}
        if isinstance(v, int) and v in ERRORES and c.HasFormula and str(c.Text).startswith("#"): o["e"] = ERRORES[v]
        elif isinstance(v, bool): o["b"] = v
        elif isinstance(v, (int, float)): o["n"] = float(v)
        elif v is None: o["t"] = ""
        else: o["t"] = str(v)
        try:
            if c.HasSpill: o["derrame"] = [c.SpillingToRange.Rows.Count, c.SpillingToRange.Columns.Count]
        except Exception: pass
        r.append(o)
    json.dump(r, open(sal, "w", encoding="utf8"), ensure_ascii=False)
    print(f"calculadas {len(r)} fórmulas", flush=True)
finally:
    wb.Close(False)
    x.Quit()

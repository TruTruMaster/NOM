"""Abre un .xlsx en Excel y dice sus escenarios (Datos → Análisis de hipótesis → Administrador de escenarios): de cada hoja,
cada escenario con sus celdas cambiantes, valores y comentario. Con HOJA y RESULTADO, además hace el informe resumen de
Excel (Scenarios.CreateSummary) y devuelve sus valores; con MOSTRAR, muestra ese escenario y devuelve las celdas.
Uso: escenarios_ver.py ENTRADA.xlsx SALIDA.json [HOJA RESULTADO [MOSTRAR]]"""
import json, os, sys
from comun import excel

entrada, salida = (os.path.abspath(x) for x in sys.argv[1:3])
hoja_res = sys.argv[3] if len(sys.argv) > 3 else None
resultado = sys.argv[4] if len(sys.argv) > 4 else None
mostrar = sys.argv[5] if len(sys.argv) > 5 else None
x = excel()
datos = {"abrio": False, "hojas": {}}
try:
    wb = x.Workbooks.Open(entrada, 0, False)
    try:
        datos["abrio"] = True
        for h in wb.Worksheets:
            lista = []
            for i in range(1, h.Scenarios().Count + 1):
                s = h.Scenarios(i)
                v = s.Values
                vals = list(v) if isinstance(v, (tuple, list)) else [v]
                try: com = s.Comment
                except Exception: com = ""   # (sin comentario, Excel da error al leerlo)
                lista.append({"nombre": s.Name, "celdas": s.ChangingCells.Address.replace("$", ""), "valores": vals,
                              "comentario": com or "", "bloqueado": bool(s.Locked), "oculto": bool(s.Hidden)})
            datos["hojas"][h.Name] = lista
        if hoja_res and resultado:
            h = wb.Worksheets(hoja_res)
            if mostrar:
                h.Scenarios(mostrar).Show(); x.Calculate()
                datos["mostrado"] = {c.Address.replace("$", ""): c.Value2 for c in h.Range("A1:H20") if c.Value2 is not None}
            antes = {s.Name for s in wb.Worksheets}
            h.Scenarios().CreateSummary(ReportType=1, ResultCells=h.Range(resultado))
            nueva = [s for s in wb.Worksheets if s.Name not in antes][0]
            u = nueva.UsedRange
            datos["resumen"] = {"hoja": nueva.Name, "celdas": {nueva.Cells(i, j).Address.replace("$", ""): nueva.Cells(i, j).Value2
                                for i in range(1, u.Row + u.Rows.Count) for j in range(1, u.Column + u.Columns.Count) if nueva.Cells(i, j).Value2 is not None}}
    finally:
        wb.Close(False)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    x.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

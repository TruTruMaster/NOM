"""x11-escenarios.xlsx: escenarios de Excel (Datos → Análisis de hipótesis → Administrador de escenarios), hechos en Excel,
para medir cómo los lee NOM, y su informe resumen (Escenarios.CreateSummary). Uso: hacer_escenarios.py SALIDA.xlsx"""
import os, sys
import pythoncom
import win32com.client
from comun import excel

salida = os.path.abspath(sys.argv[1])
x = excel()
wb = x.Workbooks.Add()
try:
    h = wb.Worksheets(1)
    h.Name = "Plan"
    datos = [("Precio", 120), ("Unidades", 500), ("Costo unitario", 70), ("Gastos fijos", 15000), ("Ingresos", "=B1*B2"), ("Costos", "=B3*B2+B4"), ("Utilidad", "=B5-B6")]
    for i, (a, b) in enumerate(datos, 1):
        h.Cells(i, 1).Value = a
        if isinstance(b, str): h.Cells(i, 2).Formula = b
        else: h.Cells(i, 2).Value = b
    sc = h.Scenarios()
    VAR = lambda xs: win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, xs)
    sc.Add(Name="Pesimista", ChangingCells=h.Range("B1:B3"), Values=VAR([100, 400, 75]), Comment="Ventas bajas")
    sc.Add(Name="Base", ChangingCells=h.Range("B1:B3"), Values=VAR([120, 500, 70]), Comment="")
    sc.Add(Name="Optimista", ChangingCells=h.Range("B1:B3"), Values=VAR([140, 650, 66]), Comment="Ventas altas")
    sc.Add(Name="Solo precio", ChangingCells=h.Range("B1"), Values=VAR([150]))
    sc.Item("Optimista").Show()
    sc.Item("Base").Show()
    sc.CreateSummary(ReportType=1, ResultCells=h.Range("B5:B7"))
    wb.SaveAs(salida, 51)   # xlOpenXMLWorkbook
finally:
    wb.Close(False)
    x.Quit()
print("hecho", flush=True)

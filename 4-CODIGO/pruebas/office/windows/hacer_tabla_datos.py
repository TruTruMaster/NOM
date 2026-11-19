"""x10-tablas-datos.xlsx: tablas de datos de Excel 2024 (Análisis de hipótesis → Tabla de datos), para que NOM las abra,
las calcule y las guarde como Excel: de una variable en columna, de una variable en fila y de dos variables, sobre un
pago mensual (PAGO), y otra con dos fórmulas.
"""
import os
from comun import excel, ARCHIVOS

x = excel()
wb = x.Workbooks.Add()
try:
    h = wb.Worksheets(1)
    h.Name = "Hipotesis"
    for ref, v in {"A1": "Tasa", "B1": 0.05, "A2": "Plazo", "B2": 60, "A3": "Monto", "B3": 100000, "A4": "Pago", "B4": "=PMT(B1/12,B2,-B3)",
                   "A5": "Total", "B5": "=B4*B2"}.items():
        if isinstance(v, str) and v.startswith("="): h.Range(ref).Formula = v
        else: h.Range(ref).Value = v
    # una variable, en columna (las tasas abajo; la fórmula arriba a la derecha), con dos fórmulas
    for i, tasa in enumerate([0.03, 0.04, 0.05, 0.06, 0.07]): h.Range(f"D{3 + i}").Value = tasa
    h.Range("E2").Formula = "=B4"
    h.Range("F2").Formula = "=B5"
    h.Range("D2:F7").Table(ColumnInput=h.Range("B1"))
    # una variable, en fila (los plazos a la derecha; la fórmula abajo a la izquierda)
    for j, plazo in enumerate([12, 24, 36, 48]): h.Cells(10, 9 + j).Value = plazo   # I10:L10
    h.Range("H11").Formula = "=B4"
    h.Range("H10:L11").Table(RowInput=h.Range("B2"))
    # dos variables (la fórmula en la esquina)
    h.Range("D14").Formula = "=B4"
    for j, plazo in enumerate([12, 24, 36, 48]): h.Cells(14, 5 + j).Value = plazo     # E14:H14
    for i, tasa in enumerate([0.03, 0.05, 0.07]): h.Cells(15 + i, 4).Value = tasa       # D15:D17
    h.Range("D14:H17").Table(RowInput=h.Range("B2"), ColumnInput=h.Range("B1"))
    h.Range("E3:F7,I11:L11,E15:H17").NumberFormat = "0.00"
    x.Calculate()
    salida = os.path.join(ARCHIVOS, "x10-tablas-datos.xlsx")
    wb.SaveAs(salida, 51)
    print("hecho", salida, flush=True)
finally:
    wb.Close(False)
    x.Quit()

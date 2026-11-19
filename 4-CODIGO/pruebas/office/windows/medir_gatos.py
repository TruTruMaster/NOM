"""Cuándo Excel 2024 muestra «####» porque un número no cabe en su columna: valores con formato de moneda y de número en
columnas de varios anchos (Calibri 11). Deja medir_gatos.json en archivos/: [{ancho_px, formato, valor, texto}].
"""
import os
from comun import excel, guardar_json, ARCHIVOS

x = excel()
wb = x.Workbooks.Add()
res = []
try:
    h = wb.Worksheets(1)
    valores = [1.5, 12.25, 123.5, 999.99, 1234.5, 1887.12, 9999.99, 12345.67, 99999.99, 123456.78, 1234567.89]
    formatos = ["$#,##0.00", "#,##0.00", "0.00", "$#,##0"]
    anchos = [6, 7, 8, 8.43, 9, 10]
    col = 1
    for w in anchos:
        for fm in formatos:
            h.Columns(col).ColumnWidth = w
            for i, v in enumerate(valores):
                c = h.Cells(i + 1, col); c.Value = v; c.NumberFormat = fm
            col += 1
    x.Calculate()
    col = 1
    for w in anchos:
        for fm in formatos:
            px = h.Columns(col).Width / 0.75
            for i, v in enumerate(valores):
                res.append({"ancho_car": w, "ancho_px": round(px, 2), "formato": fm, "valor": v, "texto": h.Cells(i + 1, col).Text})
            col += 1
    guardar_json(os.path.join(ARCHIVOS, "medir_gatos.json"), res)
    print("hecho", len(res), flush=True)
finally:
    wb.Close(False)
    x.Quit()

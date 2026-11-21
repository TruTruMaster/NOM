"""Cuándo Excel 2024 muestra «####» con formatos que dejan un hueco al final («_)» del formato de moneda y de
contabilidad): cada formato con y sin el hueco, 11 valores y la columna de píxel en píxel. Uso: medir_gatos2.py
[FUENTE] (Calibri por omisión; p. ej. "Aptos Narrow"). Deja medir_gatos2-<fuente>.json en archivos/:
[{px, formato, valor, texto}] con el ancho real de la columna en píxeles.
"""
import os, sys
from comun import excel, guardar_json, ARCHIVOS

FUENTE = sys.argv[1] if len(sys.argv) > 1 else "Calibri"

x = excel()
wb = x.Workbooks.Add()
res = []
try:
    h = wb.Worksheets(1)
    # la fuente del estilo Normal (la del libro): con ella decide Excel si un número cabe (no con la de la celda)
    wb.Styles("Normal").Font.Name = FUENTE; wb.Styles("Normal").Font.Size = 11
    h.Cells.Font.Name = FUENTE; h.Cells.Font.Size = 11
    valores = [1.5, 12.25, 123.5, 999.99, 1234.5, 1887.12, 9999.99, 12345.67, 99999.99, 123456.78, -1887.12]
    formatos = ['"$"#,##0.00_);[Red]("$"#,##0.00)', '"$"#,##0.00', '#,##0.00_)', '#,##0.00', '0.00_)', '0.00',
                '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)', '#,##0_);(#,##0)']
    # de 30 a 110 píxeles, de uno en uno (Calibri 11: 7 px por carácter y 5 de margen)
    anchos = [round((px - 5) / 7, 4) for px in range(30, 111)]
    col = 1
    for w in anchos:
        for fm in formatos:
            h.Columns(col).ColumnWidth = max(0.01, w)
            for i, v in enumerate(valores):
                c = h.Cells(i + 1, col); c.Value = v; c.NumberFormat = fm
            col += 1
    x.Calculate()
    col = 1
    for w in anchos:
        for fm in formatos:
            px = h.Columns(col).Width / 0.75
            for i, v in enumerate(valores):
                res.append({"px": round(px, 2), "formato": fm, "valor": v, "texto": h.Cells(i + 1, col).Text})
            col += 1
    guardar_json(os.path.join(ARCHIVOS, "medir_gatos2-normal-" + FUENTE.replace(" ", "_") + ".json"), res)
finally:
    wb.Close(False)
    x.Quit()
print("hecho", len(res), flush=True)

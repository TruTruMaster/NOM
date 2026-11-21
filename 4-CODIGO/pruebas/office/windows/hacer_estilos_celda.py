"""x13-estilos-celda.xlsx: los estilos de celda de Excel (Inicio → Estilos de celda), uno por fila, hechos en Excel con el
tema de Office: A el nombre del estilo, B un texto y C un número con ese estilo. Sirve para sacar cómo los ve NOM y para
medir que aplicarlos en NOM da lo mismo. Uso: hacer_estilos_celda.py SALIDA.xlsx"""
import os, sys
from comun import excel

ESTILOS = ["Normal", "Good", "Bad", "Neutral", "Calculation", "Check Cell", "Explanatory Text", "Input", "Linked Cell", "Note", "Output", "Warning Text",
           "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Title", "Total"] + \
          [f"Accent{i}" for i in range(1, 7)] + [f"{p}% - Accent{i}" for p in (20, 40, 60) for i in range(1, 7)] + \
          ["Comma", "Comma [0]", "Currency", "Currency [0]", "Percent"]
salida = os.path.abspath(sys.argv[1])
x = excel()
wb = x.Workbooks.Add()
try:
    h = wb.Worksheets(1)
    h.Name = "Estilos"
    for i, n in enumerate(ESTILOS, 1):
        h.Cells(i, 1).Value = n
        h.Cells(i, 2).Value = "Ejemplo"
        h.Cells(i, 3).Value = 1234.5
        for c in (2, 3):
            try: h.Cells(i, c).Style = n
            except Exception as e: h.Cells(i, 4).Value = "error: " + str(e)[:80]
    h.Columns(1).ColumnWidth = 18
    wb.SaveAs(salida, 51)
finally:
    wb.Close(False)
    x.Quit()
print("hecho", len(ESTILOS), flush=True)

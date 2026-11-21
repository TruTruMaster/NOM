"""x14-minigraficos.xlsx: minigráficos de Excel (Insertar → Minigráficos), hechos en Excel: de línea (con marcadores, el
punto alto y el bajo), de columna (con el negativo en otro color) y de pérdidas y ganancias, en grupos de varias celdas, y
uno con una celda vacía en los datos. Uso: hacer_minigraficos.py SALIDA.xlsx"""
import os, sys
from comun import excel

salida = os.path.abspath(sys.argv[1])
x = excel()
wb = x.Workbooks.Add()
rgb = lambda r, g, b: r + g * 256 + b * 65536
try:
    h = wb.Worksheets(1)
    h.Name = "Ventas"
    datos = [["Zona", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Tendencia"],
             ["Norte", 12, 15, 9, 18, 21, 17], ["Sur", 30, 22, 25, 28, 35, 40], ["Centro", -5, 3, 8, -2, 6, 10],
             ["Este", 7, 7, 9, 4, 12, 3], ["Oeste", 1, -1, 1, 1, -1, 1], ["Costa", 4, None, 6, 8, None, 5]]
    for i, fila in enumerate(datos, 1):
        for j, v in enumerate(fila, 1):
            if v is not None: h.Cells(i, j).Value = v
    g1 = h.Range("H2:H3").SparklineGroups.Add(Type=1, SourceData="B2:G3")          # línea
    g1.Points.Markers.Visible = True
    g1.Points.Highpoint.Visible = True
    g1.Points.Lowpoint.Visible = True
    g1.SeriesColor.Color = rgb(0, 112, 192)
    g1.LineWeight = 1.5
    g2 = h.Range("H4:H5").SparklineGroups.Add(Type=2, SourceData="B4:G5")          # columna
    g2.Points.Negative.Visible = True
    g2.Points.Negative.Color.Color = rgb(192, 0, 0)
    g2.SeriesColor.Color = rgb(112, 173, 71)
    g3 = h.Range("H6").SparklineGroups.Add(Type=3, SourceData="B6:G6")             # pérdidas y ganancias
    g4 = h.Range("H7").SparklineGroups.Add(Type=1, SourceData="B7:G7")             # con celdas vacías
    h.Columns(8).ColumnWidth = 22
    for f in range(2, 8): h.Rows(f).RowHeight = 24
    wb.SaveAs(salida, 51)
finally:
    wb.Close(False)
    x.Quit()
print("hecho", flush=True)

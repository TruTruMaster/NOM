"""x12-paginas.xlsx: la disposición de página de Excel (márgenes, orientación, tamaño, escala, ajustar, área de impresión,
títulos, saltos a mano, cuadrícula y encabezados), una hoja por caso, para medir dónde parte Excel cada hoja en páginas
(windows/paginas_ver.py) y compararlo con NOM. Uso: hacer_paginas.py SALIDA.xlsx"""
import os, sys
from comun import excel

salida = os.path.abspath(sys.argv[1])
x = excel()
wb = x.Workbooks.Add()
pulg = x.InchesToPoints


def llenar(h, filas=120, cols=26):
    # números cortos en todo el rango (lo que importa son los anchos y los altos)
    datos = [[(f * 7 + c * 3) % 1000 for c in range(cols)] for f in range(filas)]
    h.Range(h.Cells(1, 1), h.Cells(filas, cols)).Value = datos


CASOS = [
    ("base", {}),
    ("anchas", {"anchos": {1: 5, 2: 20, 3: 12, 4: 30, 6: 15, 7: 3, 9: 25, 12: 40}, "altos": {3: 30, 5: 8, 10: 45, 20: 60, 33: 4}}),
    ("horizontal-a4", {"Orientation": 2, "PaperSize": 9}),
    ("angostos", {"margenes": (0.25, 0.25, 0.75, 0.75, 0.3, 0.3)}),
    ("anchos-margenes", {"margenes": (1, 1, 1, 1, 0.5, 0.5)}),
    ("escala75", {"Zoom": 75}),
    ("escala130", {"Zoom": 130}),
    ("ajustar-ancho", {"ajustar": (1, False)}),
    ("ajustar-una", {"ajustar": (1, 1)}),
    ("area", {"PrintArea": "$B$2:$H$60"}),
    ("titulos", {"PrintTitleRows": "$1:$2", "PrintTitleColumns": "$A:$A"}),
    ("manuales", {"saltos": ("A20", "E1")}),
    ("letra14", {"letra": 14}),
    ("oficio", {"PaperSize": 5}),
    ("cuadricula", {"PrintGridlines": True, "PrintHeadings": True}),
    ("orden", {"Order": 2, "CenterHorizontally": True, "CenterVertically": True, "CenterHeader": "Página &P de &N", "CenterFooter": "&A"}),
]
try:
    while wb.Worksheets.Count < len(CASOS): wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    for i, (nombre, c) in enumerate(CASOS, 1):
        h = wb.Worksheets(i)
        h.Name = nombre
        llenar(h)
        if c.get("letra"): h.Cells.Font.Size = c["letra"]
        for col, ancho in c.get("anchos", {}).items(): h.Columns(col).ColumnWidth = ancho
        for fila, alto in c.get("altos", {}).items(): h.Rows(fila).RowHeight = alto
        ps = h.PageSetup
        for k in ("Orientation", "PaperSize", "Zoom", "PrintArea", "PrintTitleRows", "PrintTitleColumns", "PrintGridlines", "PrintHeadings", "Order",
                  "CenterHorizontally", "CenterVertically", "CenterHeader", "CenterFooter"):
            if k in c: setattr(ps, k, c[k])
        if "margenes" in c:
            iz, de, ar, ab, en, pi = c["margenes"]
            ps.LeftMargin, ps.RightMargin, ps.TopMargin, ps.BottomMargin, ps.HeaderMargin, ps.FooterMargin = (pulg(v) for v in (iz, de, ar, ab, en, pi))
        if "ajustar" in c:
            ps.Zoom = False
            ps.FitToPagesWide, ps.FitToPagesTall = c["ajustar"]
        if "saltos" in c:
            hf, vc = c["saltos"]
            h.HPageBreaks.Add(Before=h.Range(hf))
            h.VPageBreaks.Add(Before=h.Range(vc))
    wb.Worksheets(1).Activate()
    wb.SaveAs(salida, 51)   # xlOpenXMLWorkbook
finally:
    wb.Close(False)
    x.Quit()
print("hecho", len(CASOS), flush=True)

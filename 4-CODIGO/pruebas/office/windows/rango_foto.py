"""Una imagen PNG de un rango de una hoja tal como lo pinta Excel en la pantalla (Range.CopyPicture → un gráfico vacío →
Chart.Export; no necesita impresora). Uso: rango_foto.py ENTRADA.xlsx HOJA RANGO SALIDA.png"""
import os, sys, time
from comun import excel

entrada, hoja, rango, salida = os.path.abspath(sys.argv[1]), sys.argv[2], sys.argv[3], os.path.abspath(sys.argv[4])
x = excel()
try:
    wb = x.Workbooks.Open(entrada, 0, True)
    try:
        h = wb.Worksheets(hoja)
        h.Activate()
        r = h.Range(rango)
        for intento in range(5):
            try:
                r.CopyPicture(1, int(os.environ.get("NOM_FORMATO", "1")))   # xlScreen, xlPicture (o 2: xlBitmap)
                break
            except Exception:
                time.sleep(0.5)
        co = h.ChartObjects().Add(0, 0, r.Width, r.Height)
        co.Activate()
        co.Chart.ChartArea.Format.Line.Visible = 0
        time.sleep(1)
        co.Chart.Paste()
        time.sleep(1)
        co.Chart.Export(salida, "PNG")
        co.Delete()
    finally:
        wb.Close(False)
finally:
    x.Quit()
print("hecho", os.path.exists(salida), flush=True)

"""Abre un .docx en Word y lo guarda como texto con un salto en cada fin de renglón (y en PDF), para saber dónde partió
Word cada palabra con «Guiones automáticos». Uso: guiones_txt.py ENTRADA.docx SALIDA.txt [SALIDA.pdf]"""
import os, sys, time
from comun import word

entrada, salida = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
pdf = os.path.abspath(sys.argv[3]) if len(sys.argv) > 3 else None
w = word()
try:
    t0 = time.time()
    d = w.Documents.Open(entrada, False, True, False, "", "", False, "", "", 0, 0, True, False, 0, True)
    try:
        d.ActiveWindow.View.Type = 3
        d.Repaginate()
        # FileFormat 7 = texto Unicode; InsertLineBreaks: un salto donde termina cada renglón
        d.SaveAs2(FileName=salida, FileFormat=7, Encoding=1200, InsertLineBreaks=True, LineEnding=0, AddToRecentFiles=False)
        if pdf: d.ExportAsFixedFormat(pdf, 17)
        print("hecho", round(time.time() - t0, 1), "s", d.ComputeStatistics(2), "páginas", flush=True)
    finally:
        d.Close(0)
finally:
    w.Quit()

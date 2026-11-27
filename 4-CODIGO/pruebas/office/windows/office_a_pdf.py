"""Abre cada archivo de Office en su programa (Word, Excel o PowerPoint de verdad) y lo guarda como PDF: la referencia
de cómo se ve. Dice si el programa no lo pudo abrir.
Uso: office_a_pdf.py SALIDA.json ARCHIVO [ARCHIVO …]   (cada PDF queda junto a su archivo, con .pdf al final)"""
import json, os, sys
from comun import word, excel, powerpoint

salida = os.path.abspath(sys.argv[1]); res = {}
for a in (os.path.abspath(x) for x in sys.argv[2:]):
    pdf = a + ".pdf"; r = res[os.path.basename(a)] = {"abrio": False}
    ext = a.lower().rsplit(".", 1)[-1]
    try:
        if ext == "docx":
            w = word()
            try:
                try: d = w.Documents.Open(a, False, True, False, "", "", False, "", "", 0, 0, True, False, 0, True)
                except Exception as e0:
                    r["sin_reparar"] = str(e0)[:200]   # Word no lo abre tal cual: se intenta «Abrir y reparar»
                    d = w.Documents.Open(a, False, True, False, "", "", False, "", "", 0, 0, True, True, 0, True); r["reparado"] = True
                r["abrio"] = True; d.Repaginate(); r["paginas"] = d.ComputeStatistics(2)
                d.ExportAsFixedFormat(pdf, 17, False, 0, 0, 1, 1, 0, True, True, 1, True, True, False); d.Close(0)
            finally: w.Quit()
        elif ext == "xlsx":
            x = excel()
            try:
                try: wb = x.Workbooks.Open(a, 0, True)
                except Exception as e0:
                    r["sin_reparar"] = str(e0)[:200]   # Excel no lo abre tal cual: se intenta repararlo (CorruptLoad = 1)
                    wb = x.Workbooks.Open(a, 0, True, None, None, None, True, None, None, False, False, None, False, False, 1); r["reparado"] = True
                r["abrio"] = True; r["hojas"] = wb.Worksheets.Count
                try: wb.ExportAsFixedFormat(0, pdf)
                except Exception as e1: r["sin_pdf"] = str(e1)[:120]   # (sin impresora instalada Excel no exporta)
                wb.Close(False)
            finally: x.Quit()
        elif ext == "pptx":
            p = powerpoint()
            pr = p.Presentations.Open(a, True, False, False); r["abrio"] = True; r["diapositivas"] = pr.Slides.Count
            pr.SaveAs(pdf, 32); pr.Close()
        r["pdf"] = os.path.exists(pdf)
    except Exception as e:
        r["error"] = str(e)[:400]
    print(os.path.basename(a), r, flush=True)
json.dump(res, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

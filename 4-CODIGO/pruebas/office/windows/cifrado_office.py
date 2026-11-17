"""Archivos de Office con contraseña, con Office de verdad.

  python cifrado_office.py hacer CONTRASEÑA        guarda con contraseña w01, x01 y p01 (archivos\\cifrado-*.docx/xlsx/pptx)
  python cifrado_office.py abrir CONTRASEÑA RUTAS… abre en Office cada archivo con esa contraseña y dice qué trae

Deja archivos\\cifrado.json con el resultado (si abrió, cuántos párrafos, hojas o diapositivas y el primer texto).
"""
import glob, json, os, sys, time, traceback
from comun import *


def hacer(clave):
    out = []
    w = word()
    try:
        d = w.Documents.Open(os.path.join(ARCHIVOS, "w01-letra.docx"), False, True)
        ruta = os.path.join(ARCHIVOS, "cifrado-word.docx")
        d.SaveAs2(ruta, 16, False, clave)
        d.Close(0); out.append(ruta)
    finally:
        w.Quit(0)
    x = excel()
    try:
        wb = x.Workbooks.Open(os.path.join(ARCHIVOS, "x01-formato.xlsx"))
        ruta = os.path.join(ARCHIVOS, "cifrado-excel.xlsx")
        wb.SaveAs(ruta, 51, clave)
        wb.Close(False); out.append(ruta)
    finally:
        x.Quit()
    p = powerpoint()
    pr = p.Presentations.Open(os.path.join(ARCHIVOS, "p01-texto.pptx"), True, False, False)
    try:
        pr.Password = clave
        ruta = os.path.join(ARCHIVOS, "cifrado-powerpoint.pptx")
        pr.SaveAs(ruta)
        out.append(ruta)
    finally:
        pr.Close()
    print("hechos:", out)


def abrir(clave, rutas):
    res = []
    for ruta in rutas:
        ruta = os.path.abspath(ruta)
        r = {"archivo": os.path.basename(ruta), "abrio": False}
        ext = os.path.splitext(ruta)[1].lower()
        try:
            if ext == ".docx":
                w = word()
                try:
                    d = w.Documents.Open(ruta, False, True, False, clave)
                    r.update(abrio=True, parrafos=d.Paragraphs.Count, texto=d.Paragraphs(1).Range.Text[:80])
                    d.Close(0)
                finally:
                    w.Quit(0)
            elif ext == ".xlsx":
                x = excel()
                try:
                    wb = x.Workbooks.Open(ruta, 0, True, None, clave)
                    r.update(abrio=True, hojas=[s.Name for s in wb.Worksheets], a1=str(wb.Worksheets(1).Range("A1").Text)[:80])
                    wb.Close(False)
                finally:
                    x.Quit()
            elif ext == ".pptx":
                p = powerpoint()
                pr = p.Presentations.Open(ruta + "::" + clave + "::", True, False, False)
                try:
                    sl = pr.Slides
                    r.update(abrio=True, diapositivas=sl.Count)
                    try: r["texto"] = sl(1).Shapes(1).TextFrame.TextRange.Text[:80]
                    except Exception: pass
                finally:
                    pr.Close()
        except Exception as e:
            r["error"] = str(getattr(e, "excepinfo", None) or e)[:300]
        print(r, flush=True)
        res.append(r)
    guardar_json(os.path.join(ARCHIVOS, "cifrado.json"), res)


if __name__ == "__main__":
    import pythoncom
    pythoncom.CoInitialize()
    a = sys.argv[1:]
    if a[0] == "hacer": hacer(a[1])
    else:
        rutas = []
        for x in a[2:]: rutas += glob.glob(x) or glob.glob(os.path.join(BASE, x))
        abrir(a[1], rutas)

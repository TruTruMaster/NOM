"""¿Office acepta la «contraseña para modificar» de un archivo? Abre cada uno con la contraseña buena y con una mala.

  python modificar_ver.py CONTRASEÑA RUTAS…
"""
import os, sys
from comun import *


def probar(ruta, clave):
    ext = os.path.splitext(ruta)[1].lower()
    r = {}
    for nombre, c in [("buena", clave), ("mala", "contraseña-mala")]:
        try:
            if ext == ".docx":
                w = word()
                try:
                    d = w.Documents.Open(ruta, False, False, False, "", "", False, c)
                    r[nombre] = "edita" if not d.ReadOnly else "solo lectura"
                    d.Close(0)
                finally:
                    w.Quit(0)
            elif ext == ".xlsx":
                x = excel()
                try:
                    wb = x.Workbooks.Open(ruta, 0, False, None, "", c)
                    r[nombre] = "edita" if not wb.ReadOnly else "solo lectura"
                    wb.Close(False)
                finally:
                    x.Quit()
            elif ext == ".pptx":
                p = powerpoint()
                pr = p.Presentations.Open(ruta + "::::" + c, False, False, False)
                r[nombre] = "edita" if not pr.ReadOnly else "solo lectura"
                pr.Close()
        except Exception as e:
            r[nombre] = "error: " + str(getattr(e, "excepinfo", None) or e)[:140]
    return r


if __name__ == "__main__":
    import pythoncom
    pythoncom.CoInitialize()
    for ruta in sys.argv[2:]:
        ruta = os.path.abspath(ruta if os.path.isabs(ruta) else os.path.join(BASE, ruta))
        print(os.path.basename(ruta), probar(ruta, sys.argv[1]), flush=True)

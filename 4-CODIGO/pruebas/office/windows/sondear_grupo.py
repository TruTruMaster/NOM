"""Prueba en qué condiciones Word deja agrupar formas por automatización (para armar archivos con grupos).

  python sondear_grupo.py
"""
import os
from comun import *

pythoncom.CoInitialize()
w = word()
d = w.Documents.Add()
res = []
def intento(nombre, visible, tipo_ajuste, mismo_ancla):
    global d
    try:
        w.Visible = visible
        if visible:
            try: w.WindowState = 0; w.Left = -3000; w.Top = 0
            except Exception: pass
        d.Content.Text = "Uno\rDos\rTres"
        anc1 = d.Paragraphs(1).Range; anc2 = d.Paragraphs(2).Range if not mismo_ancla else anc1
        a = d.Shapes.AddShape(1, 50, 50, 80, 40, anc1); b = d.Shapes.AddShape(9, 150, 60, 60, 60, anc2)
        for s in (a, b): s.WrapFormat.Type = tipo_ajuste
        a.Name = nombre + "a"; b.Name = nombre + "b"
        nombres = W.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [a.Name, b.Name])
        g = d.Shapes.Range(nombres).Group()
        res.append((nombre, "bien", g.Name))
    except Exception as e:
        msg = e.excepinfo[2] if hasattr(e, "excepinfo") and e.excepinfo else str(e)
        res.append((nombre, "falla", msg[:120]))
    for s in list(d.Shapes):
        try: s.Delete()
        except Exception: pass
try:
    for vis in (False, True):
        for t in (3, 0, 5):
            for mismo in (True, False):
                intento(f"v{int(vis)}t{t}m{int(mismo)}", vis, t, mismo)
    try: print("compatibilidad:", d.CompatibilityMode)
    except Exception: pass
finally:
    d.Close(False); w.Quit()
for r in res: print(r)

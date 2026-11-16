"""Posición horizontal (en puntos, desde la página) de cada letra de los renglones que contienen un texto.

  python sondear_renglon.py ARCHIVO.docx "texto del renglón" [...]
Sirve para ver cómo acomoda Word un renglón (cuánto mide cada letra y cada espacio).
"""
import os, sys, json
from comun import *
pythoncom.CoInitialize()
w = word()
d = w.Documents.Open(os.path.abspath(sys.argv[1]), False, True, False, "", "", False, "", "", 0, 0, True, False, 0, True)
d.ActiveWindow.View.Type = 3
res = []
for marca in sys.argv[2:]:
    r = d.Content
    f = r.Find
    if not f.Execute(marca): res.append({"marca": marca, "no": True}); continue
    ini = r.Start
    linea0 = d.Range(ini, ini).Information(10)       # wdFirstCharacterLineNumber
    # retroceder al inicio del renglón
    k = ini
    while k > 0 and d.Range(k - 1, k - 1).Information(10) == linea0 and d.Range(k - 1, k).Text not in ("\r",): k -= 1
    pos = []
    j = k
    while True:
        c = d.Range(j, j + 1)
        if c.Text in ("\r", "") or d.Range(j, j).Information(10) != linea0: break
        pos.append([c.Text, d.Range(j, j).Information(5)])
        j += 1
        if len(pos) > 300: break
    fin = d.Range(j, j).Information(5)
    res.append({"marca": marca, "letras": pos, "fin": fin, "fuente": d.Range(k, k + 1).Font.Name, "tam": d.Range(k, k + 1).Font.Size})
print(json.dumps(res, ensure_ascii=False))
d.Close(0); w.Quit(0)

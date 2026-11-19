"""Mide en Word real cuánto avanza un texto con cada formato de letra: versalitas, escala de caracteres (w:w),
espaciado, elevado o bajado. Para sacar la regla con la que NOM corta los renglones como Word.

Cada caso es un párrafo «|texto|» en una hoja muy ancha (no se corta): con Range.Information(5) (posición horizontal
desde la orilla de la hoja, en puntos) se lee dónde empieza y dónde termina el texto. Deja medir_letras.json en
archivos/: [{fuente, tam, caso, texto, ancho}].
"""
import json, os
from comun import word, guardar_json, ARCHIVOS

FUENTES = ["Calibri", "Aptos", "Times New Roman", "Arial", "Cambria", "Georgia", "Verdana", "Segoe UI"]
TAMANOS = [11, 12]
TEXTOS = {
    "minus": "abcdefghijklmnopqrstuvwxyz áéíóúñü",
    "mayus": "ABCDEFGHIJKLMNOPQRSTUVWXYZ ÁÉÍÓÚÑÜ",
    "frase": "Referencia sutil, Énfasis intenso: 248 estudiantes (73.5 %)",
}
# (nombre, cómo se aplica al rango)
CASOS = [
    ("normal", None),
    ("versalitas", lambda r: setattr(r.Font, "SmallCaps", True)),
    ("mayusculas", lambda r: setattr(r.Font, "AllCaps", True)),
    ("escala150", lambda r: setattr(r.Font, "Scaling", 150)),
    ("escala80", lambda r: setattr(r.Font, "Scaling", 80)),
    ("escala66", lambda r: setattr(r.Font, "Scaling", 66)),
    ("espaciado2", lambda r: setattr(r.Font, "Spacing", 2)),
    ("espaciado-1", lambda r: setattr(r.Font, "Spacing", -1)),
    ("elevado6", lambda r: setattr(r.Font, "Position", 6)),
    ("superindice", lambda r: setattr(r.Font, "Superscript", True)),
    ("versalitas-escala150", lambda r: (setattr(r.Font, "SmallCaps", True), setattr(r.Font, "Scaling", 150))),
]

w = word()
d = w.Documents.Add()
res = []
try:
    d.PageSetup.PageWidth, d.PageSetup.PageHeight = 1584, 792
    d.PageSetup.LeftMargin = d.PageSetup.RightMargin = 36
    d.AutoHyphenation = False
    est = d.Styles(-1)
    est.Font.Kerning = 0
    est.ParagraphFormat.SpaceAfter = 0
    w.ActiveWindow.View.Type = 3   # diseño de impresión (Information necesita la página)
    filas = []
    primero = True
    for fuente in FUENTES:
        for tam in TAMANOS:
            for cn, fn in CASOS:
                for tn, tx in TEXTOS.items():
                    if primero: primero = False
                    else: d.Content.InsertParagraphAfter()
                    p = d.Paragraphs(d.Paragraphs.Count)
                    p.Range.Text = "|" + tx + "|"
                    p = d.Paragraphs(d.Paragraphs.Count)
                    p.Range.Font.Name = fuente
                    p.Range.Font.Size = tam
                    p.Range.Font.Kerning = 0
                    ini = p.Range.Start + 1
                    r = d.Range(ini, ini + len(tx))
                    if fn: fn(r)
                    filas.append((fuente, tam, cn, tn, tx, ini))
    for fuente, tam, cn, tn, tx, ini in filas:
        x0 = d.Range(ini, ini).Information(5)
        x1 = d.Range(ini + len(tx), ini + len(tx)).Information(5)
        res.append({"fuente": fuente, "tam": tam, "caso": cn, "texto": tn, "cadena": tx, "ancho": round(x1 - x0, 3), "x0": x0})
    guardar_json(os.path.join(ARCHIVOS, "medir_letras.json"), res)
    print("hecho", len(res), "medidas", flush=True)
finally:
    d.Close(0)
    w.Quit()

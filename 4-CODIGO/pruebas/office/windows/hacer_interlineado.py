"""p12-interlineado.pptx: dónde pone PowerPoint cada renglón de un cuadro de texto, por fuente y tamaño.

Una diapositiva por fuente; en cada una, cuadros de texto de 3 renglones (con saltos de línea) en varios tamaños,
con interlineado sencillo, 1.5 y exacto, sin autoajuste, anclados arriba y con los márgenes de siempre. verdad.py saca
la posición de cada renglón (BoundTop, BoundHeight) y analisis/interlineado_ppt.py la compara con NOM.
"""
import json, os
from comun import powerpoint, ARCHIVOS

import sys
FUENTES = ["Calibri", "Aptos", "Aptos Display", "Arial", "Times New Roman", "Segoe UI", "Georgia", "Verdana", "Cambria",
           "Calibri Light", "Century Gothic", "Tahoma"]
TAMANOS = [12, 18, 24, 32, 44]
ESPACIOS = [("sencillo", 1.0), ("uno y medio", 1.5), ("exacto", None)]
if "--mas" in sys.argv:   # más interlineados (el 90 % es el de los títulos de los temas de Office)
    FUENTES = ["Calibri", "Aptos", "Aptos Display", "Arial", "Times New Roman", "Segoe UI"]
    ESPACIOS = [("80", 0.8), ("90", 0.9), ("sencillo", 1.0), ("110", 1.1), ("120", 1.2), ("uno y medio", 1.5), ("doble", 2.0), ("exacto", None)]

pp = powerpoint()
pres = pp.Presentations.Add(False)
meta = []
try:
    pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight = 960, 540
    for fuente in FUENTES:
        for nombre_esp, mult in ESPACIOS:
            s = pres.Slides.Add(pres.Slides.Count + 1, 12)   # en blanco
            x = 10
            for tam in TAMANOS:
                caja = s.Shapes.AddTextbox(1, x, 20, 170, 500)
                tf = caja.TextFrame
                tf.AutoSize = 0
                tf.WordWrap = True
                tf.VerticalAnchor = 1
                tr = tf.TextRange
                tr.Text = "HxxH xHHx\rxHHx HxxH\rHxxH xHHx"   # (sin descendentes ni acentos: abajo de la tinta = línea base)
                tr.Font.Name = fuente
                tr.Font.Size = tam
                pf = tr.ParagraphFormat
                pf.SpaceBefore = 0
                pf.SpaceAfter = 0
                if mult is None:
                    pf.LineRuleWithin = False
                    pf.SpaceWithin = round(tam * 1.4, 1)
                else:
                    pf.LineRuleWithin = True
                    pf.SpaceWithin = mult
                meta.append({"diapo": s.SlideIndex, "forma": caja.Name, "fuente": fuente, "tam": tam, "espacio": nombre_esp,
                             "mult": mult, "exacto": None if mult is not None else round(tam * 1.4, 1)})
                x += 188
    nombre = "p13-interlineado-mas" if "--mas" in sys.argv else "p12-interlineado"
    salida = os.path.join(ARCHIVOS, nombre + ".pptx")
    pres.SaveAs(salida, 24)
    json.dump(meta, open(os.path.join(ARCHIVOS, nombre + ".json"), "w", encoding="utf-8"), ensure_ascii=False)
    print("hecho", pres.Slides.Count, "diapositivas,", len(meta), "cuadros", flush=True)
finally:
    pres.Close()
    # (PowerPoint es una sola copia en el equipo: solo se cierra si no queda ninguna presentación abierta)
    try:
        if pp.Presentations.Count == 0: pp.Quit()
    except Exception: pass

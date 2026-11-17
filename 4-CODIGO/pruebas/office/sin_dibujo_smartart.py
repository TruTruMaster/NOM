#!/usr/bin/env python3
"""Copia de un .pptx (o .docx/.xlsx) con sus SmartArt SIN el dibujo guardado (ppt/diagrams/drawingN.xml):
como los guardan otros programas. PowerPoint los vuelve a acomodar con su diseño (layout) al abrirlos; NOM
también tiene que hacerlo.

  python3 sin_dibujo_smartart.py ENTRADA.pptx SALIDA.pptx
"""
import re, sys, zipfile

ent, sal = sys.argv[1], sys.argv[2]
zi = zipfile.ZipFile(ent)
nombres = zi.namelist()
dibujos = {n for n in nombres if re.search(r"/diagrams/drawing\d*\.xml$", n)}
quitar = set(dibujos) | {n.rsplit("/", 1)[0] + "/_rels/" + n.rsplit("/", 1)[1] + ".rels" for n in dibujos}
with zipfile.ZipFile(sal, "w", zipfile.ZIP_DEFLATED) as zo:
    for n in nombres:
        if n in quitar: continue
        d = zi.read(n)
        if n == "[Content_Types].xml":
            t = d.decode("utf8")
            t = re.sub(r'<Override PartName="/[^"]*/diagrams/drawing\d*\.xml"[^>]*/>', "", t)
            d = t.encode("utf8")
        elif n.endswith(".rels"):
            t = d.decode("utf8")
            t = re.sub(r'<Relationship [^>]*Type="[^"]*/diagramDrawing"[^>]*/>', "", t)
            d = t.encode("utf8")
        elif re.search(r"/diagrams/data\d*\.xml$", n):
            # el dato que apunta al dibujo (dsp:dataModelExt) se va también
            t = d.decode("utf8")
            t = re.sub(r'<a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">.*?</a:ext>', "", t, flags=re.S)
            t = re.sub(r"<a:extLst>\s*</a:extLst>", "", t)
            d = t.encode("utf8")
        zo.writestr(zi.getinfo(n), d)
print(f"{len(dibujos)} dibujos quitados -> {sal}")

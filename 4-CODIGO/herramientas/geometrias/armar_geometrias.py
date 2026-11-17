#!/usr/bin/env python3
"""Convierte las 187 geometrías predefinidas de Office (ECMA-376, presetShapeDefinitions.xml) a un archivo
JavaScript compacto que usa NOM para dibujar las formas de PowerPoint, Word y Excel igual que Office.

  python3 armar_geometrias.py presetShapeDefinitions.xml ../../app/src/main/resources/static/js/geometrias-datos.js

El archivo de definiciones es parte de la norma ECMA-376 (Office Open XML, parte 1, anexo D); se puede bajar
de ecma-international.org o de los proyectos libres que lo incluyen (LibreOffice, Apache POI).
"""
import json, sys
import xml.etree.ElementTree as ET

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def num_o_nombre(v):
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (TypeError, ValueError):
        return v


def guias(nodo):
    return [[g.get("name"), g.get("fmla")] for g in (nodo.findall(A + "gd") if nodo is not None else [])]


def punto(p):
    pt = p.find(A + "pt")
    return [num_o_nombre(pt.get("x")), num_o_nombre(pt.get("y"))]


def main(origen, destino):
    raiz = ET.parse(origen).getroot()
    salida = {}
    for forma in raiz:
        nombre = forma.tag
        d = {}
        av = guias(forma.find(A + "avLst"))
        gd = guias(forma.find(A + "gdLst"))
        if av: d["av"] = av
        if gd: d["gd"] = gd
        r = forma.find(A + "rect")
        if r is not None: d["rect"] = [num_o_nombre(r.get(k)) for k in ("l", "t", "r", "b")]
        trazos = []
        pl = forma.find(A + "pathLst")
        for p in (pl.findall(A + "path") if pl is not None else []):
            t = {}
            if p.get("w"): t["w"] = int(p.get("w"))
            if p.get("h"): t["h"] = int(p.get("h"))
            if p.get("fill") and p.get("fill") != "norm": t["f"] = p.get("fill")
            if p.get("stroke") == "false": t["s"] = 0
            cmds = []
            for c in p:
                tag = c.tag.replace(A, "")
                if tag == "moveTo": cmds.append(["M"] + punto(c))
                elif tag == "lnTo": cmds.append(["L"] + punto(c))
                elif tag == "arcTo": cmds.append(["A"] + [num_o_nombre(c.get(k)) for k in ("wR", "hR", "stAng", "swAng")])
                elif tag == "quadBezTo": cmds.append(["Q"] + [v for pt in c.findall(A + "pt") for v in (num_o_nombre(pt.get("x")), num_o_nombre(pt.get("y")))])
                elif tag == "cubicBezTo": cmds.append(["C"] + [v for pt in c.findall(A + "pt") for v in (num_o_nombre(pt.get("x")), num_o_nombre(pt.get("y")))])
                elif tag == "close": cmds.append(["Z"])
            t["c"] = cmds
            trazos.append(t)
        d["p"] = trazos
        salida[nombre] = d
    js = ("/* NOM · Las 187 geometrías predefinidas de Office (ECMA-376, parte 1, anexo D:\n"
          "   presetShapeDefinitions.xml). Las arma herramientas/geometrias/armar_geometrias.py; las dibuja\n"
          "   js/geometrias.js. av: ajustes (manijas) por omisión; gd: guías (fórmulas); rect: dónde va el texto;\n"
          "   p: trazos (M L A Q C Z), con relleno f (none, darken, lighten…) y s: 0 sin contorno. */\n"
          "window.NOM = window.NOM || {};\nNOM.GEOMETRIAS = " + json.dumps(salida, ensure_ascii=False, separators=(",", ":")) + ";\n")
    open(destino, "w").write(js)
    print(f"{len(salida)} geometrías, {len(js) // 1024} KB")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])

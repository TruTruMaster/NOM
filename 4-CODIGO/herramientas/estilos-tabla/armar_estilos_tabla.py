#!/usr/bin/env python3
"""Convierte los estilos de tabla y de tabla dinámica predefinidos de Excel (ECMA-376, presetTableStyles.xml:
60 de tabla y 84 de tabla dinámica) a un archivo JavaScript compacto que usa NOM para pintar las
tablas de un libro de Excel igual que Excel (Excel no guarda en el archivo los estilos que trae de fábrica).

  python3 armar_estilos_tabla.py presetTableStyles.xml ../../app/src/main/resources/static/js/estilos-tabla-datos.js

El archivo de definiciones es parte de la norma ECMA-376 (parte 1, «SpreadsheetMLStyles»); también lo traen
proyectos libres como Apache POI.

Formato de salida, por estilo y por elemento (wholeTable, headerRow, firstRowStripe…):
  f: relleno · k: color de letra · b: negrita · i: cursiva · s: tamaño de la franja
  bd: { l, r, t, b, h, v: [estilo, color] }  (h y v son las líneas de adentro)
Los colores van como «t4» o «t4:0.79998» (color del tema y su tinte), «#RRGGBB» o «i22» (indexado).
"""
import json, sys
import xml.etree.ElementTree as ET


def color(n):
    if n is None: return None
    if n.get("theme") is not None:
        t = n.get("tint")
        # (el tinte con todos sus decimales: Excel trunca la luminosidad y un redondeo puede cambiar un tono)
        return f"t{n.get('theme')}" + (f":{t}" if t and float(t) != 0 else "")
    if n.get("rgb"): return "#" + n.get("rgb")[-6:].upper()
    if n.get("indexed") is not None: return "i" + n.get("indexed")
    if n.get("auto") == "1": return "t1"
    return None


def formato(dxf):
    o = {}
    fuente = dxf.find("font")
    if fuente is not None:
        b = fuente.find("b")
        if b is not None and b.get("val", "1") not in ("0", "false"): o["b"] = 1
        i = fuente.find("i")
        if i is not None and i.get("val", "1") not in ("0", "false"): o["i"] = 1
        c = color(fuente.find("color"))
        if c: o["k"] = c
    pf = dxf.find("fill/patternFill")
    if pf is not None and pf.get("patternType", "solid") != "none":
        # en un formato diferencial (dxf) el color del relleno liso va en bgColor
        c = color(pf.find("bgColor")) or color(pf.find("fgColor"))
        if c: o["f"] = c
    borde = dxf.find("border")
    if borde is not None:
        bd = {}
        for lado, k in (("left", "l"), ("right", "r"), ("top", "t"), ("bottom", "b"), ("horizontal", "h"), ("vertical", "v")):
            e = borde.find(lado)
            if e is not None and e.get("style") and e.get("style") != "none":
                bd[k] = [e.get("style"), color(e.find("color")) or "t1"]
        if bd: o["bd"] = bd
    return o


def main(origen, destino):
    raiz = ET.parse(origen).getroot()
    salida = {}
    for est in raiz:
        dxfs = [formato(d) for d in est.findall("dxfs/dxf")]
        ts = est.find("tableStyles/tableStyle")
        if ts is None: continue
        elems = ts.findall("tableStyleElement")
        ids = [int(e.get("dxfId", "-1")) for e in elems]
        # en este archivo unos estilos cuentan los formatos desde 1 y otros desde 0
        base = 1 if ids and max(ids) >= len(dxfs) else 0
        d = {}
        for e in elems:
            i = int(e.get("dxfId", "-1")) - base
            f = dict(dxfs[i]) if 0 <= i < len(dxfs) else {}
            if e.get("size") and e.get("size") != "1": f["s"] = int(e.get("size"))
            d[e.get("type")] = f
        salida[est.tag] = d
    js = ("/* NOM · Estilos de tabla y de tabla dinámica que Excel trae de fábrica (ECMA-376, parte 1:\n"
          "   presetTableStyles.xml). Los arma herramientas/estilos-tabla/armar_estilos_tabla.py; los usa js/ooxml.js\n"
          "   al importar un libro (las tablas con formato de Excel se ven con su estilo). Por elemento: f relleno,\n"
          "   k color de letra, b negrita, i cursiva, s tamaño de la franja, bd bordes {l r t b h v: [estilo, color]};\n"
          "   colores «t4:0.8» (tema y tinte), «#RRGGBB» o «i22» (indexado). */\n"
          "window.NOM = window.NOM || {};\nNOM.ESTILOS_TABLA = " + json.dumps(salida, ensure_ascii=False, separators=(",", ":")) + ";\n")
    open(destino, "w").write(js)
    print(f"{len(salida)} estilos, {len(js) // 1024} KB")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])

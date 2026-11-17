#!/usr/bin/env python3
"""Arma la biblioteca de efectos de animación de PowerPoint que usa NOM, con lo que escribe PowerPoint de verdad.

  python3 armar_presets.py a01-animaciones.pptx a01-animaciones.json a02-opciones.pptx a02-opciones.json \
          ../../app/src/main/resources/static/js/animaciones-office-datos.js

Los .pptx los hace `pruebas/office/windows/animaciones_office.py` (catalogo y opciones) con PowerPoint de Office 2024:
una diapositiva por efecto, con su dirección. De cada una se toma el árbol del efecto (el cTn con presetClass y lo
que tiene adentro) tal cual, sin sus números de nodo, su tipo de inicio ni su retraso (eso lo pone NOM), y con el
destino como «@» (el bloque al que se le aplique). Así NOM crea y escribe cada efecto igual que PowerPoint.
"""
import json, re, sys, zipfile
import xml.etree.ElementTree as ET

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
NS = {"http://schemas.openxmlformats.org/presentationml/2006/main": "p", "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
      "http://schemas.microsoft.com/office/powerpoint/2010/main": "p14", "http://schemas.microsoft.com/office/powerpoint/2012/main": "p15",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r"}

# Nombres en español (como los muestra PowerPoint en español, o lo más cercano)
NOMBRES = {
    "entr": {1: "Aparecer", 2: "Desplazar hacia dentro", 3: "Persianas", 4: "Cuadro", 5: "Cuadros bicolores", 6: "Círculo", 7: "Arrastrar hacia dentro",
             8: "Rombo", 9: "Disolver hacia dentro", 10: "Desvanecer", 11: "Destello único", 12: "Asomar", 13: "Más", 14: "Barras aleatorias",
             15: "Espiral hacia dentro", 16: "Dividir", 17: "Expandir", 18: "Barras diagonales", 19: "Rotar", 20: "Cuña", 21: "Rueda", 22: "Barrido",
             23: "Acercar", 25: "Bumerán", 26: "Rebote", 27: "Revelar con color", 28: "Créditos", 29: "Aliviar", 30: "Flotar",
             31: "Aumentar y girar", 34: "Velocidad de la luz", 35: "Molinete", 37: "Elevar", 38: "Silbido", 39: "Línea fina", 40: "Desplegar",
             41: "Látigo", 42: "Flotar hacia dentro", 43: "Giro central", 45: "Girar", 47: "Descender", 48: "Honda", 49: "Rotación",
             50: "Estirar", 51: "Cremallera", 52: "Arco hacia arriba", 53: "Zoom", 54: "Deslizar", 55: "Ampliar", 56: "Voltear", 58: "Plegar"},
    "exit": {1: "Desaparecer", 2: "Desplazar hacia fuera", 3: "Persianas", 4: "Cuadro", 5: "Cuadros bicolores", 6: "Círculo", 7: "Arrastrar hacia fuera",
             8: "Rombo", 9: "Disolver hacia fuera", 10: "Desvanecer", 11: "Destello único", 12: "Asomar hacia fuera", 13: "Más", 14: "Barras aleatorias",
             15: "Espiral hacia fuera", 16: "Dividir", 17: "Contraer", 18: "Barras diagonales", 19: "Rotar", 20: "Cuña", 21: "Rueda", 22: "Barrido",
             23: "Alejar", 25: "Bumerán", 26: "Rebote", 27: "Revelar con color", 28: "Créditos", 29: "Aliviar", 30: "Flotar",
             31: "Reducir y girar", 34: "Velocidad de la luz", 35: "Molinete", 37: "Hundir", 38: "Silbido", 39: "Línea fina", 40: "Plegar",
             41: "Látigo", 42: "Flotar hacia fuera", 43: "Giro central", 45: "Girar", 47: "Descender", 48: "Honda", 49: "Rotación",
             50: "Estirar", 51: "Cremallera", 52: "Arco hacia abajo", 53: "Zoom", 54: "Deslizar", 55: "Contraer", 56: "Voltear", 58: "Plegar"},
    "emph": {1: "Color de relleno", 2: "Cambiar fuente", 3: "Color de fuente", 4: "Tamaño de fuente", 5: "Estilo de fuente", 6: "Aumentar o encoger",
             7: "Color de línea", 8: "Girar", 9: "Transparencia", 10: "Destello en negrita", 14: "Explosión", 15: "Revelar en negrita",
             16: "Color con pincel", 18: "Subrayado", 19: "Color de objeto", 20: "Onda de color", 21: "Color complementario",
             22: "Color complementario 2", 23: "Color de contraste", 24: "Oscurecer", 25: "Desaturar", 26: "Pulso", 27: "Parpadeo",
             28: "Aumentar con color", 30: "Aclarar", 31: "Énfasis de estilo", 32: "Balancear", 33: "Resaltado vertical", 34: "Onda", 35: "Intermitente",
             36: "Brillo"},
    "path": {1: "Círculo", 2: "Triángulo rectángulo", 3: "Rombo", 4: "Hexágono", 5: "Estrella de 5 puntas", 6: "Luna creciente", 7: "Cuadrado",
             8: "Trapecio", 9: "Corazón", 10: "Octágono", 11: "Estrella de 6 puntas", 12: "Balón", 13: "Triángulo equilátero", 14: "Paralelogramo",
             15: "Pentágono", 16: "Estrella de 4 puntas", 17: "Estrella de 8 puntas", 18: "Lágrima", 19: "Estrella puntiaguda", 20: "Cuadrado curvo",
             21: "X curva", 22: "Ocho vertical", 23: "Estrella ondulada", 24: "Rizo", 25: "Sierra", 26: "Ocho horizontal", 27: "Cacahuate",
             28: "Ocho por cuatro", 29: "Neutrón", 30: "Remolino", 31: "Frijol", 32: "Más", 33: "Triángulo invertido", 34: "Cuadrado invertido",
             35: "Izquierda", 36: "Giro a la derecha", 37: "Arco hacia abajo", 38: "Zigzag", 39: "Curva en S 2", 40: "Onda sinusoidal",
             41: "Rebote a la izquierda", 42: "Abajo", 43: "Giro hacia arriba", 44: "Arco hacia arriba", 45: "Latido", 46: "Espiral a la derecha",
             47: "Onda", 48: "Curva a la izquierda", 49: "Diagonal abajo a la derecha", 50: "Giro hacia abajo", 51: "Arco a la izquierda",
             52: "Embudo", 53: "Resorte", 54: "Rebote a la derecha", 55: "Espiral a la izquierda", 56: "Diagonal arriba a la derecha",
             57: "Giro arriba a la derecha", 58: "Arco a la derecha", 59: "Curva en S 1", 60: "Onda que decae", 61: "Curva a la derecha",
             62: "Escalera hacia abajo", 63: "Derecha", 64: "Arriba"},
}
# Categorías del panel de PowerPoint
CATEGORIA = {
    "entr": {"Básico": [1, 3, 4, 5, 6, 8, 9, 13, 14, 10, 2, 22, 16, 21, 20, 18, 17, 12, 7, 11, 15, 19, 23], "Sutil": [53, 37, 42, 31, 45, 55, 58, 26],
             "Moderado": [25, 29, 30, 34, 43, 47, 48, 49, 50, 52, 54, 56], "Llamativo": [27, 28, 35, 38, 39, 40, 41, 51]},
    "emph": {"Básico": [1, 3, 4, 5, 6, 7, 8, 9], "Sutil": [10, 15, 16, 18, 19, 21, 22, 23, 24, 25, 26, 30, 31, 33], "Moderado": [20, 27, 28, 32, 34, 36], "Llamativo": [14]},
}
DIRECCIONES = {1: "Desde arriba", 2: "Desde la derecha", 3: "Desde arriba a la derecha", 4: "Desde abajo", 6: "Desde abajo a la derecha",
               8: "Desde la izquierda", 9: "Desde arriba a la izquierda", 12: "Desde abajo a la izquierda", 5: "Vertical", 10: "Horizontal",
               16: "Hacia dentro", 32: "Hacia fuera", 21: "Horizontal hacia dentro", 26: "Horizontal hacia fuera", 37: "Vertical hacia fuera",
               42: "Vertical hacia dentro", 36: "Hacia fuera", 528: "Hacia dentro ligeramente", 272: "Hacia fuera ligeramente"}


def nombre_q(tag):
    ns, _, loc = tag[1:].partition("}")
    return (NS.get(ns, "x") + ":" + loc) if ns else tag


def a_nodo(el):
    a = {}
    for k, v in el.attrib.items():
        if k.startswith("{"): k = nombre_q(k)
        a[k] = v
    hs = [a_nodo(c) for c in el]
    if el.text and el.text.strip(): hs.insert(0, el.text.strip())
    return [nombre_q(el.tag), a, hs] if hs else [nombre_q(el.tag), a]


def limpiar(n):
    """Sin números de nodo; el destino como «@»."""
    if not isinstance(n, list): return n
    nom, a = n[0], dict(n[1]) if len(n) > 1 else {}
    loc = nom.split(":")[-1]
    if loc == "cTn": a.pop("id", None)
    if loc in ("spTgt", "inkTgt"): a.pop("spid", None); a["bid"] = "@"
    hs = [limpiar(c) for c in (n[2] if len(n) > 2 else [])]
    return [nom, a, hs] if hs else [nom, a]


def dur_natural(ctn):
    """Lo más largo de los comportamientos (inicio + duración × repeticiones)."""
    D = 0
    def tiempo(c):
        b = 0
        st = c.find(P + "stCondLst")
        if st is not None:
            for cd in st.findall(P + "cond"):
                d = cd.get("delay")
                if d and d != "indefinite": b = float(d); break
        d = c.get("dur"); d = None if d is None else (float("inf") if d == "indefinite" else float(d))
        r = c.get("repeatCount"); r = 1 if r is None else (float("inf") if r == "indefinite" else float(r) / 1000)
        return b, d, r, c.get("autoRev") in ("1", "true")
    def rec(lst, base):
        nonlocal D
        if lst is None: return
        for n in lst:
            loc = n.tag.split("}")[-1]
            if loc in ("par", "seq"):
                c = n.find(P + "cTn")
                if c is not None:
                    b, _, _, _ = tiempo(c); rec(c.find(P + "childTnLst"), base + b)
                continue
            cb = n.find(P + "cBhvr")
            c = cb.find(P + "cTn") if cb is not None else None
            if c is None: continue
            b, d, r, rev = tiempo(c)
            d = d or 0
            ad = d * r * (2 if rev else 1)
            if ad != float("inf"): D = max(D, base + b + ad)
    rec(ctn.find(P + "childTnLst"), 0)
    return round(D)


def efectos_de(pptx, info):
    z = zipfile.ZipFile(pptx)
    out = []
    for x in json.load(open(info, encoding="utf8"))["diapositivas"]:
        t = ET.fromstring(z.read(f"ppt/slides/slide{x['diapo']}.xml"))
        for c in t.iter(P + "cTn"):
            if c.get("presetClass"):
                out.append((x, c))
    return out


def main(argv):
    pares, salida = argv[:-1], argv[-1]
    presets = {}
    for i in range(0, len(pares), 2):
        for x, c in efectos_de(pares[i], pares[i + 1]):
            pcl, pid, sub = c.get("presetClass"), int(c.get("presetID")), int(c.get("presetSubtype") or 0)
            efecto = x["efecto"]
            # solo efectos «puros» (no por letra, palabra o párrafo, ni con opciones de tiempo: esas se aplican encima)
            if re.search(r"por |repetir|autorreversa|rebobinar|suaves|rebote al|retraso|después|Con y|propia|Atenuar|Ocultar|al revés", efecto): continue
            if re.search(r"\d+ ?%|grados|rojo|verde", efecto): continue
            clave = f"{pcl}:{pid}:{sub}"
            if clave in presets: continue
            nodo = limpiar(a_nodo(c))
            a = nodo[1]
            for k in ("nodeType",): a.pop(k, None)
            nodo[2] = [h for h in nodo[2] if not (isinstance(h, list) and h[0] == "p:stCondLst")] if len(nodo) > 2 else []
            nombre = NOMBRES.get(pcl, {}).get(pid) or efecto
            presets[clave] = {"n": nombre, "en": efecto.split(" ")[0], "c": nodo, "d0": dur_natural(c) or 1,
                              "dir": DIRECCIONES.get(sub) if pcl in ("entr", "exit") and sub else None, "texto": bool(x.get("texto"))}
    cats = {k: {cat: [i for i in ids if any(p.startswith(f"{k}:{i}:") for p in presets)] for cat, ids in v.items()} for k, v in CATEGORIA.items()}
    js = ("/* Efectos de animación de PowerPoint (Office 2024) tal como los escribe PowerPoint. Generado por\n"
          " * herramientas/animaciones/armar_presets.py desde los archivos que hizo PowerPoint; no editar a mano. */\n"
          "(function () {\n  const NOM = (window.NOM = window.NOM || {});\n"
          f"  NOM.animOfficeDatos = {{ presets: {json.dumps(presets, ensure_ascii=False, separators=(',', ':'))},\n"
          f"    categorias: {json.dumps(cats, ensure_ascii=False, separators=(',', ':'))} }};\n}})();\n")
    open(salida, "w", encoding="utf8").write(js)
    print(len(presets), "efectos;", sum(1 for k in presets if k.startswith("entr")), "de entrada,", sum(1 for k in presets if k.startswith("emph")), "de énfasis,",
          sum(1 for k in presets if k.startswith("exit")), "de salida,", sum(1 for k in presets if k.startswith("path")), "trayectorias;", len(js) // 1024, "KB")


if __name__ == "__main__":
    main(sys.argv[1:])

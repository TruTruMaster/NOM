#!/usr/bin/env python3
"""Dónde dibuja PowerPoint cada renglón contra dónde lo dibuja NOM (p12-interlineado, windows/hacer_interlineado.py).

Mide la tinta de cada renglón en las imágenes a 1920 px (2 px por punto) de PowerPoint (verdad.py) y de NOM
(comparar_real.py): arriba de la tinta y abajo (el tercer renglón, «Éxito final», no tiene descendentes: su borde de
abajo es la línea base). Da la diferencia NOM − PowerPoint por fuente, tamaño e interlineado.
"""
import json, os, statistics as st
from collections import defaultdict
from pathlib import Path
import numpy as np
from PIL import Image

C = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"


def renglones(img, x0, x1, y0, y1):
    """Los grupos de filas con tinta (arriba, abajo) en ese rectángulo."""
    a = np.asarray(img.convert("L"))[max(0, y0):y1, max(0, x0):x1] < 140
    filas = a.any(axis=1)
    grupos, ini = [], None
    for i, f in enumerate(filas):
        if f and ini is None: ini = i
        if not f and ini is not None: grupos.append((ini + y0, i - 1 + y0)); ini = None
    if ini is not None: grupos.append((ini + y0, len(filas) - 1 + y0))
    return grupos


def main():
    v = json.load(open(C / "verdad/p12-interlineado.verdad.json"))
    meta = {(m["diapo"], m["forma"]): m for m in json.load(open(C / "archivos/p12-interlineado.json", encoding="utf-8"))}
    res = defaultdict(list)
    for s in v["diapositivas"]:
        pp = Image.open(C / "verdad/p12png" / s["png"])
        nom = Image.open(C / "comparar/p12-interlineado" / f"nom-{s['n']:02d}.png")
        for f in s["formas"]:
            m = meta.get((s["n"], f["nombre"]))
            if not m or len(f.get("lineas") or []) < 3: continue
            x0, x1 = int((f["x"] + 6) * 2), int((f["x"] + f["w"] - 6) * 2)
            y0, y1 = int(f["y"] * 2), int((f["lineas"][2]["y"] + f["lineas"][2]["h"]) * 2 + m["tam"] * 2)
            gp, gn = renglones(pp, x0, x1, y0, y1), renglones(nom, x0, x1, y0, y1)
            if len(gp) < 3 or len(gn) < 3: res[(m["fuente"], m["espacio"])].append((m["tam"], None)); continue
            # (arriba del primero, abajo del tercero = línea base del tercero, paso entre la base 1→3)
            d_arriba = (gn[0][0] - gp[0][0]) / 2
            d_base3 = (gn[2][1] - gp[2][1]) / 2
            res[(m["fuente"], m["espacio"])].append((m["tam"], (round(d_arriba, 1), round(d_base3, 1))))
    print("NOM − PowerPoint en puntos: (arriba de la tinta del 1.er renglón, línea base del 3.º) por tamaño")
    for k in sorted(res):
        print(f"{k[0]:16s} {k[1]:12s}", "  ".join(f"{t}:{d}" for t, d in sorted(res[k], key=lambda x: x[0])))


if __name__ == "__main__":
    main()

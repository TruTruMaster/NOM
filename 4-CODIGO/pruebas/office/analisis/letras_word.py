#!/usr/bin/env python3
"""Regla de Word para el ancho de versalitas, escala de caracteres y espaciado (datos de windows/medir_letras.py).

  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/analisis/letras_word.py medir_letras.json
"""
import json, sys, collections
from pathlib import Path
from fontTools.ttLib import TTFont, TTCollection

OF = Path.home() / ".cache/nom-pruebas/fuentes-office"
ARCH = {"Calibri": "windows/calibri.ttf", "Aptos": "aptos/Aptos/30153066857.ttf", "Times New Roman": "windows/times.ttf", "Arial": "windows/arial.ttf",
        "Cambria": "windows/cambria.ttc", "Georgia": "windows/georgia.ttf", "Verdana": "windows/verdana.ttf", "Segoe UI": "windows/segoeui.ttf"}
_c = {}
def fuente(n):
    if n not in _c:
        p = OF / ARCH[n]
        f = TTCollection(str(p)).fonts[0] if p.suffix.lower() == ".ttc" else TTFont(str(p))
        _c[n] = (f["hmtx"].metrics, f.getBestCmap(), f["head"].unitsPerEm)
    return _c[n]
def ancho(n, txt, tam):
    m, cmap, upem = fuente(n)
    return sum(m[cmap[ord(ch)]][0] for ch in txt) / upem * tam

def main():
    d = json.load(open(sys.argv[1]))
    por = collections.defaultdict(dict)
    for x in d: por[(x["fuente"], x["tam"], x["texto"])][x["caso"]] = x
    # 1. normal: ¿las medidas de la fuente dan lo que mide Word?
    errs = [abs(v["normal"]["ancho"] - ancho(k[0], v["normal"]["cadena"], k[1])) for k, v in por.items()]
    print(f"normal: error máx {max(errs):.3f} pt, medio {sum(errs)/len(errs):.3f}")
    # 2. versalitas: minúsculas → mayúscula a tam*k; lo demás igual. Probar k
    for k_ in [0.65, 0.7, 0.75, 0.8]:
        e = []
        for (f, t, tn), v in por.items():
            s = v["normal"]["cadena"]
            pred = sum(ancho(f, ch.upper(), t * k_) if ch.islower() else ancho(f, ch, t) for ch in s)
            e.append(v["versalitas"]["ancho"] - pred)
        print(f"versalitas k={k_}: error medio {sum(e)/len(e):+.3f}, máx {max(map(abs, e)):.3f}")
    # el k exacto por caso (solo minúsculas)
    ks = []
    for (f, t, tn), v in por.items():
        if tn != "minus": continue
        s = v["normal"]["cadena"]
        fijo = sum(ancho(f, ch, t) for ch in s if not ch.islower())
        mayus = sum(ancho(f, ch.upper(), t) for ch in s if ch.islower())
        k = (v["versalitas"]["ancho"] - fijo) / mayus
        ks.append(k); print(f"  {f:16} {t:4} k={k:.4f}  (tam·k = {t*k:.3f})")
    # 3. escala: ancho = normal × escala?
    for c, fac in [("escala150", 1.5), ("escala80", 0.8), ("escala66", 0.66)]:
        e = [v[c]["ancho"] - v["normal"]["ancho"] * fac for v in por.values()]
        print(f"{c}: error medio {sum(e)/len(e):+.3f}, máx {max(map(abs, e)):.3f}")
    # 4. espaciado: + n pt por letra?
    for c, n in [("espaciado2", 2), ("espaciado-1", -1)]:
        e = [v[c]["ancho"] - (v["normal"]["ancho"] + n * len(v["normal"]["cadena"])) for v in por.values()]
        print(f"{c}: error medio {sum(e)/len(e):+.3f}, máx {max(map(abs, e)):.3f}")
    for c in ["elevado6", "mayusculas", "superindice", "versalitas-escala150"]:
        print(c, [round(v[c]["ancho"] - v["normal"]["ancho"], 2) for v in list(por.values())[:6]])

main()

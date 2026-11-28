#!/usr/bin/env python3
"""Deja el logo de NOM en colores planos (27-nov-2026): sin degradados, sin resplandor y sin brillo.

Recorre los SVG del logo (archivos sueltos y los que van dentro de las páginas HTML) y cambia cada
degradado por su color plano, según el color con que empezaba:
  fondo  #2B2952 → #15142A   →  #2B2952 (noche)
  lengua #A9A9FF → #6767E8   →  #E4E4EE (clara, sobre el mosaico oscuro)
  lengua #7070EC → #4646BE   →  #2B2952 (sobre papel)
  fuego  #FF9F57 → #EE4F38   →  #F06A3D
El resplandor (radial) y el brillo (blanco transparente) se quitan junto con su rectángulo.
Uso: logo_plano.py ARCHIVO...  (cambia los archivos en su lugar; los que no tienen el logo no se tocan)
"""
import re
import sys

PLANOS = {"#2B2952": "#2B2952", "#A9A9FF": "#E4E4EE", "#7070EC": "#2B2952", "#FF9F57": "#F06A3D"}
GRAD = re.compile(r'<(linear|radial)Gradient id="([^"]+)"[^>]*>(.*?)</\1Gradient>', re.S)


def aplanar(s):
    planos, quitar = {}, set()
    for tipo, gid, cuerpo in GRAD.findall(s):
        m = re.search(r'stop-color="([^"]+)"', cuerpo)
        color = m.group(1).upper() if m else ""
        if tipo == "radial" or color in ("#FFF", "#FFFFFF"):
            quitar.add(gid)
        elif color in PLANOS:
            planos[gid] = PLANOS[color]
    if not planos and not quitar:
        return s
    for gid in quitar:   # el rectángulo del resplandor o del brillo
        s = re.sub(r'\s*<rect [^>]*fill="url\(#%s\)"\s*/>' % re.escape(gid), "", s)
    for gid, c in planos.items():
        s = s.replace('fill="url(#%s)"' % gid, 'fill="%s"' % c)
    usados = planos.keys() | quitar
    s = GRAD.sub(lambda m: "" if m.group(2) in usados else m.group(0), s)
    s = re.sub(r"\s*<defs>\s*</defs>", "", s)
    return s


if __name__ == "__main__":
    for ruta in sys.argv[1:]:
        with open(ruta, encoding="utf-8") as f:
            a = f.read()
        b = aplanar(a)
        if a != b:
            with open(ruta, "w", encoding="utf-8") as f:
                f.write(b)
            print("plano:", ruta)

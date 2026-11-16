#!/usr/bin/env python3
"""Imágenes de prueba para los archivos que arma Office (fotos, logotipo con transparencia, SVG).

Se hacen aquí (Linux) y se copian a Windows: así los archivos de Office siempre llevan las mismas.
  python3 imagenes.py CARPETA
"""
import math, random, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter, ImageFont

def foto(w, h, semilla):
    """Paisaje inventado: cielo degradado, sol, montañas con ruido (parece foto y tiene detalle)."""
    r = random.Random(semilla)
    im = Image.new("RGB", (w, h))
    px = im.load()
    for y in range(h):
        t = y / h
        c = (int(40 + 120 * t), int(110 + 90 * t), int(200 + 40 * t))
        for x in range(w): px[x, y] = c
    d = ImageDraw.Draw(im)
    d.ellipse([w * 0.7, h * 0.1, w * 0.85, h * 0.1 + w * 0.15], fill=(255, 220, 120))
    for capa, (col, base) in enumerate([((90, 110, 140), 0.55), ((60, 90, 70), 0.68), ((40, 70, 40), 0.8)]):
        pts = [(0, h)]
        y = h * base
        for x in range(0, w + 40, 40):
            y += r.uniform(-30, 30) * (1 - capa * 0.2)
            y = max(h * (base - 0.2), min(h * (base + 0.1), y))
            pts.append((x, y))
        pts.append((w, h)); d.polygon(pts, fill=col)
    im = im.filter(ImageFilter.GaussianBlur(0.6))
    ruido = Image.effect_noise((w, h), 18).convert("RGB")
    return Image.blend(im, ruido, 0.06)

def logo(n):
    im = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.ellipse([8, 8, n - 8, n - 8], fill=(29, 122, 76, 255), outline=(10, 60, 40, 255), width=10)
    d.polygon([(n * 0.3, n * 0.7), (n * 0.5, n * 0.25), (n * 0.7, n * 0.7)], fill=(255, 255, 255, 230))
    d.ellipse([n * 0.42, n * 0.52, n * 0.58, n * 0.68], fill=(255, 196, 0, 255))
    return im

def barras(w, h):
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    cols = [(68, 114, 196), (237, 125, 49), (165, 165, 165), (255, 192, 0), (91, 155, 213), (112, 173, 71)]
    for i, v in enumerate([0.8, 0.55, 0.9, 0.35, 0.65, 0.5]):
        x0 = 40 + i * (w - 80) / 6
        d.rectangle([x0 + 10, h - 40 - v * (h - 80), x0 + (w - 80) / 6 - 10, h - 40], fill=cols[i])
    d.line([30, h - 40, w - 30, h - 40], fill="black", width=2)
    return im

SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="400" height="400">
<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#1D7A4C"/><stop offset="1" stop-color="#7FD1A8"/></linearGradient></defs>
<rect x="5" y="5" width="90" height="90" rx="18" fill="url(#g)"/>
<path d="M25 70 L50 25 L75 70 Z" fill="#fff" opacity="0.9"/>
<circle cx="50" cy="58" r="8" fill="#FFC400"/>
</svg>"""

def main(carpeta):
    c = Path(carpeta); c.mkdir(parents=True, exist_ok=True)
    foto(1600, 1000, 1).save(c / "foto1.jpg", quality=88)
    foto(1200, 1200, 2).save(c / "foto2.jpg", quality=85)
    foto(900, 1400, 3).save(c / "foto3-vertical.jpg", quality=85)
    logo(512).save(c / "logo.png")
    barras(800, 500).save(c / "barras.png")
    (c / "icono.svg").write_text(SVG)
    # mosaico pequeño para rellenos con imagen
    t = Image.new("RGB", (64, 64), (250, 240, 220)); d = ImageDraw.Draw(t)
    d.rectangle([0, 0, 31, 31], fill=(220, 120, 60)); d.rectangle([32, 32, 63, 63], fill=(60, 120, 220))
    t.save(c / "mosaico.png")
    print("listo", sorted(p.name for p in c.iterdir()))

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "imagenes")

#!/usr/bin/env python3
"""¿Con qué anchos acomoda Word cada letra? Compara las posiciones exactas de cada glifo en el PDF de Word
(mutool trace) con varios modelos, usando la fuente real (HarfBuzz: avance + interletraje + ligaduras).

  avances_word.py ARCHIVO.pdf FUENTE.ttf [--pagina N] [--tam PT] [--familia "Aptos"]

Modelos del paso entre un glifo y el siguiente (px a 96 ppp):
  exacto   avance + interletraje, sin redondear
  M1       redondeo (mitad hacia arriba) del avance, sin interletraje
  M2       redondeo de (avance + interletraje)
  M3       redondeo del avance + redondeo del interletraje
Dice qué porcentaje de pasos acierta cada modelo (±0.05 px).
"""
import argparse, math, re, subprocess, sys
from collections import Counter
import uharfbuzz as hb

def rmu(x): return math.floor(x + 0.5)

def glifos_pdf(pdf, pagina, familia):
    t = subprocess.run(["mutool", "trace", pdf, str(pagina)], capture_output=True, text=True).stdout
    lineas = {}
    for sp in re.finditer(r'<span font="([^"]+)"[^>]*trm="([\d.]+) [^"]*">(.*?)</span>', t, re.S):
        fuente, tam = sp.group(1), float(sp.group(2))
        if not re.search(r"\+" + re.escape(familia) + r"$", fuente): continue
        for g in re.finditer(r'<g unicode="([^"]*)" glyph="(\d+)" x="([\d.]+)" y="([\d.]+)" adv="([\d.]+)"', sp.group(3)):
            u = g.group(1).replace("&amp;", "&").replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">").replace("&apos;", "'")
            lineas.setdefault((round(float(g.group(4)), 1), tam), []).append((u, int(g.group(2)), float(g.group(3))))
    return lineas

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf"); ap.add_argument("ttf")
    ap.add_argument("--pagina", type=int, default=1); ap.add_argument("--familia", default="Aptos")
    ap.add_argument("--ppem-entero", dest="ppem_entero", action="store_true", help="el tamaño en px por em, redondeado")
    a = ap.parse_args()
    blob = hb.Blob.from_file_path(a.ttf); face = hb.Face(blob); font = hb.Font(face); upm = face.upem
    lineas = glifos_pdf(a.pdf, a.pagina, a.familia)
    aciertos = Counter(); total = 0; ejemplos = []
    for (y, tam), gs in sorted(lineas.items(), reverse=True):
        if len(gs) < 3: continue
        texto = "".join(g[0] for g in gs)
        buf = hb.Buffer(); buf.add_str(texto); buf.guess_segment_properties()
        hb.shape(font, buf, {"kern": True, "liga": True, "clig": True, "calt": True})
        infos, pos = buf.glyph_infos, buf.glyph_positions
        # avance de HarfBuzz con interletraje y sin él (para separar el interletraje)
        buf2 = hb.Buffer(); buf2.add_str(texto); buf2.guess_segment_properties()
        hb.shape(font, buf2, {"kern": False, "liga": True, "clig": True, "calt": True})
        sin = buf2.glyph_positions
        if len(infos) != len(gs) or len(sin) != len(gs):
            continue   # (ligaduras: otra cantidad de glifos; se salta el renglón)
        ppem = rmu(tam * 96 / 72) if a.ppem_entero else tam * 96 / 72
        k = ppem / upm   # px por unidad (Word fija la letra en un número entero de píxeles por em)
        for i in range(len(gs) - 1):
            obs = (gs[i + 1][2] - gs[i][2]) * 96 / 72
            av = sin[i].x_advance * k; ke = (pos[i].x_advance - sin[i].x_advance) * k
            mod = {"exacto": av + ke, "M1": rmu(av), "M2": rmu(av + ke), "M3": rmu(av) + rmu(ke)}
            total += 1
            for n, v in mod.items():
                if abs(v - obs) <= 0.05: aciertos[n] += 1
            if abs(mod["M2"] - obs) > 0.05 and len(ejemplos) < 12: ejemplos.append((gs[i][0] + gs[i + 1][0], round(obs, 3), {n: round(v, 3) for n, v in mod.items()}))
    print(f"{total} pasos entre glifos")
    for n in ("exacto", "M1", "M2", "M3"): print(f"  {n:7s} {100 * aciertos[n] / max(1, total):5.1f} %")
    for e in ejemplos: print("   falla M2:", e)

if __name__ == "__main__":
    main()

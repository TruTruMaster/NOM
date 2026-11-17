import sys, pickle, subprocess, numpy as np
V = "/home/user/.cache/nom-pruebas/office-real/archivos/a01-animaciones.mp4"
por, ts, tr = pickle.load(open(V + ".marcas.pkl", "rb"))
def cuadro(d, t):
    f0, f1 = tr[d]
    f = min(range(f0, f1 + 1), key=lambda i: abs((ts[i] - ts[f0]) * 1000 - t))
    b = subprocess.run(["ffmpeg", "-v", "error", "-i", V, "-fps_mode", "passthrough", "-vf", f"select=eq(n\\,{f})", "-vframes", "1", "-f", "rawvideo", "-pix_fmt", "rgb24", "-"], capture_output=True).stdout
    return np.frombuffer(b, np.uint8).reshape(720, 1280, 3).astype(float), round((ts[f] - ts[f0]) * 1000)
# figura: x 380..580 pt, y 200..340 pt  ->  px*4/3: x 506.7..773.3, y 266.7..453.3
def alfa(px):  # 0 = blanco, 1 = azul (2F5597)
    azul = np.array([0x2F, 0x55, 0x97], float)
    return np.clip((255 - px.mean(axis=-1)) / (255 - azul.mean()), 0, 1.2)
def ancho_borde(v):
    """largo (px) de la transición 10%-90% en un perfil 1D."""
    v = np.asarray(v); i10 = np.where(v > 0.1)[0]; i90 = np.where(v > 0.9)[0]
    return v
for d, t, eje, pos, nom in [(43, 150, "col", 640, "wipe"), (43, 300, "col", 640, "wipe"), (5, 150, "col", 640, "blinds"), (31, 150, "fila", 360, "barn"),
                            (7, 1000, "fila", 360, "box"), (11, 1000, "fila", 360, "circle"), (15, 1500, "fila", 360, "diamond"), (25, 1000, "fila", 280, "plus"), (9, 200, "fila", 280, "checker")]:
    im, tt = cuadro(d, t)
    a = alfa(im)
    v = a[:, pos] if eje == "col" else a[pos, :]
    lo, hi = (260, 460) if eje == "col" else (500, 780)
    seg = v[lo:hi]
    # transiciones: donde cruza 0.5; ancho 10-90 alrededor de cada cruce
    cr = [i for i in range(1, len(seg)) if (seg[i - 1] - 0.5) * (seg[i] - 0.5) < 0]
    anchos = []
    for c in cr:
        j0 = c; j1 = c
        sube = seg[c] > seg[c - 1]
        while 0 < j0 and ((seg[j0] > 0.1) if sube else (seg[j0] < 0.9)): j0 -= 1
        while j1 < len(seg) - 1 and ((seg[j1] < 0.9) if sube else (seg[j1] > 0.1)): j1 += 1
        anchos.append((lo + c, j1 - j0))
    print(f"{nom:8s} d{d} t={tt}: cruces (pos, ancho 10-90 px@1280): {anchos}")

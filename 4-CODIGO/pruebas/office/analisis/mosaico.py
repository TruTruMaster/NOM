"""Mosaico de cuadros de PowerPoint: mosaico.py VIDEO salida.png diapo:t1,t2,... [diapo:...]  (recorta la zona de la figura)"""
import sys, subprocess, pickle, os, json
sys.path.insert(0, "/home/user/NOM/4-CODIGO/pruebas/office")
import animaciones_real as R
from PIL import Image, ImageDraw
import numpy as np
video, salida = sys.argv[1], sys.argv[2]
cache = video + ".marcas.pkl"
if os.path.exists(cache): por, ts, tramos = pickle.load(open(cache, "rb"))
else: por, ts, tramos = R.marcas(video); pickle.dump((por, ts, tramos), open(cache, "wb"))
pedidos = []
for a in sys.argv[3:]:
    d, _, tl = a.partition(":")
    for t in tl.split(","):
        f0, f1 = tramos[int(d)]
        f = min(range(f0, f1 + 1), key=lambda i: abs((ts[i] - ts[f0]) * 1000 - float(t)))
        pedidos.append((int(d), float(t), f))
idx = sorted({f for _, _, f in pedidos})
W, H = 1280, 720
p = subprocess.Popen(["ffmpeg", "-v", "error", "-i", video, "-fps_mode", "passthrough", "-f", "rawvideo", "-pix_fmt", "rgb24", "-"], stdout=subprocess.PIPE)
fr, n = {}, 0
while n <= idx[-1]:
    b = p.stdout.read(W * H * 3)
    if len(b) < W * H * 3: break
    if n in idx: fr[n] = np.frombuffer(b, np.uint8).reshape(H, W, 3).copy()
    n += 1
p.kill()
caja = (380, 160, 900, 520)  # x0,y0,x1,y1 en px del video
cw, ch = caja[2] - caja[0], caja[3] - caja[1]
cols = 4
filas = (len(pedidos) + cols - 1) // cols
im = Image.new("RGB", (cols * (cw + 6), filas * (ch + 20)), "white")
dr = ImageDraw.Draw(im)
for i, (d, t, f) in enumerate(pedidos):
    x, y = (i % cols) * (cw + 6), (i // cols) * (ch + 20)
    im.paste(Image.fromarray(fr[f][caja[1]:caja[3], caja[0]:caja[2]]), (x, y + 18))
    dr.text((x + 2, y + 2), f"d{d} t={int(t)}", fill="black")
im.save(salida); print(salida, im.size)

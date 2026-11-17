"""Sigue cada letra de un texto que entra volando (Fly por letra) en el video de PowerPoint: por columnas de la letra final."""
import sys, pickle, subprocess, json, numpy as np
V = sys.argv[1]; d = int(sys.argv[2])
por, ts, tr = pickle.load(open(V + ".marcas.pkl", "rb"))
f0, f1 = tr[d]
p = subprocess.Popen(["ffmpeg", "-v", "error", "-i", V, "-fps_mode", "passthrough", "-vf", f"select=between(n\\,{f0}\\,{f1})", "-f", "rawvideo", "-pix_fmt", "gray", "-"], stdout=subprocess.PIPE)
cuadros = []
for f in range(f0, f1 + 1):
    b = p.stdout.read(1280 * 720)
    if len(b) < 1280 * 720: break
    cuadros.append(np.frombuffer(b, np.uint8).reshape(720, 1280)[:690].astype(int))
final = cuadros[-1]
cols = np.nonzero((final < 128).any(axis=0))[0]
# columnas agrupadas en letras (huecos de 2 px o más)
grupos, ini = [], cols[0]
for a, b in zip(cols, cols[1:]):
    if b - a > 1: grupos.append((ini, a)); ini = b
grupos.append((ini, cols[-1]))
y_final = []
for g in grupos:
    ys = np.nonzero((final[:, g[0]:g[1] + 1] < 128).any(axis=1))[0]; y_final.append(ys.min())
res = {"grupos": [list(map(int, g)) for g in grupos], "y_final": [int(v) for v in y_final], "t": [], "top": []}
for i, c in enumerate(cuadros):
    t = round((ts[f0 + i] - ts[f0]) * 1000)
    tops = []
    for g in grupos:
        ys = np.nonzero((c[:, g[0]:g[1] + 1] < 128).any(axis=1))[0]
        tops.append(int(ys.min()) if len(ys) else None)
    res["t"].append(t); res["top"].append(tops)
json.dump(res, open(sys.argv[3], "w"))
print(len(grupos), "letras", res["grupos"], "y_final", res["y_final"])
for i in range(0, 40, 3): print(res["t"][i], res["top"][i])

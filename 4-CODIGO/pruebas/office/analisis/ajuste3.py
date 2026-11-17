import json, re, math, sys
import numpy as np
exec(open("ajuste_camino.py").read().split("modelos = {}")[0].replace("datos = json.load(open(sys.argv[1])); path = sys.argv[2]; D = float(sys.argv[3]); acc = dec = float(sys.argv[4]) if len(sys.argv) > 4 else 0.5", "datos = json.load(open(sys.argv[1])); path = sys.argv[2]; D = float(sys.argv[3]); acc = dec = 0.5"))
import os
ESC = np.array([1, 1]) if os.environ.get("FRAC") else np.array([960, 540])
pts, L = densa(ESC)
obs = []
for t, v in datos:
    if v is None or t > D: continue
    o = np.array([(v[0] - base[0]) / 960, (v[1] - base[1]) / 540])
    # el punto del camino más cercano, buscando cerca de donde iba (el camino pasa dos veces por el origen)
    dist = np.hypot(*(pts - o).T * np.array([[960], [540]]))
    k = int(np.argmin(dist)) if not obs else int(np.argmin(dist + 1e6 * (np.arange(len(pts)) < obs[-1][2] - 5)))
    obs.append((t / D, L[k] / L[-1], k))
def smil(s, a=0.5, d=0.5):
    R = 1 / (1 - a / 2 - d / 2)
    if s < a: return R * s * s / (2 * a)
    if s <= 1 - d: return R * (s - a / 2)
    t = 1 - s; return 1 - R * t * t / (2 * d)
cand = {"smil": smil, "smootherstep": lambda s: 6 * s ** 5 - 15 * s ** 4 + 10 * s ** 3, "smoothstep": lambda s: 3 * s * s - 2 * s ** 3,
        "seno": lambda s: (1 - math.cos(math.pi * s)) / 2, "lineal": lambda s: s}
for n, f in cand.items():
    e = [f(s) - o for s, o, _ in obs if 0.02 < s < 0.98]
    print(f"{n:13s} rms {np.sqrt(np.mean(np.square(e))):.4f}  max {np.max(np.abs(e)):.4f}")
print([ (round(s,3), round(o,3)) for s, o, _ in obs[::3]])
from math import comb
def ibeta(s, a):
    # I_s(a, a) para a entero o real (por integración numérica)
    xs = np.linspace(0, 1, 2001); y = (xs * (1 - xs)) ** (a - 1); c = np.cumsum(y); c /= c[-1]
    return np.interp(s, xs, c)
res = []
for a in np.arange(2.0, 4.01, 0.1):
    for off in range(-40, 41, 5):
        e = [ibeta(min(1, max(0, s + off / D)), a) - o for s, o, _ in obs if 0.02 < s < 0.98]
        res.append((np.sqrt(np.mean(np.square(e))), round(a, 1), off))
res.sort(); print(res[:6])
res2 = []
for a in np.arange(0.3, 0.71, 0.02):
    for off in range(-40, 41, 5):
        e = [smil(min(1, max(0, s + off / D)), a, a) - o for s, o, _ in obs if 0.02 < s < 0.98]
        res2.append((np.sqrt(np.mean(np.square(e))), round(a, 2), off))
res2.sort(); print(res2[:4])

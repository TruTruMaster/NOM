import json, re, math, sys
import numpy as np
exec(open("ajuste_camino.py").read().split("modelos = {}")[0].replace("datos = json.load(open(sys.argv[1])); path = sys.argv[2]; D = float(sys.argv[3]); acc = dec = float(sys.argv[4]) if len(sys.argv) > 4 else 0.5", "datos = json.load(open(sys.argv[1])); path = sys.argv[2]; D = float(sys.argv[3]); acc = dec = 0.5"))
ptsF, LF = densa(np.array([1, 1])); ptsP, LP = densa(np.array([960, 540]))
def por_largo(pts, L, s):
    l = s * L[-1]; k = np.searchsorted(L, l); k = min(max(k, 1), len(L) - 1); r = (l - L[k - 1]) / max(1e-12, L[k] - L[k - 1]); return pts[k - 1] + (pts[k] - pts[k - 1]) * r
def por_param(s):
    n = len(segs); q = min(n - 1e-9, s * n); k = int(q); return np.array(punto(segs[k], q - k))
def warp2(u, a, d):
    if not (a or d): return u
    R = 1 / (1 - a / 2 - d / 2)
    if u < a * D: return R * u * u / (2 * a * D)
    if u <= (1 - d) * D: return R * (u - a * D / 2)
    t = D - u; return D - R * t * t / (2 * d * D)
mejor = []
for modo in ("largo_px", "largo_frac", "param"):
    for a in np.arange(0.30, 0.71, 0.02):
        for off in range(-60, 61, 10):
            err = []
            for t, v in datos:
                if v is None or t > D + 200: continue
                tt = min(D, max(0, t + off)); s = warp2(tt, a, a) / D
                x, y = por_largo(ptsP, LP, s) if modo == "largo_px" else por_largo(ptsF, LF, s) if modo == "largo_frac" else por_param(s)
                err.append(math.hypot(base[0] + x * 960 - v[0], base[1] + y * 540 - v[1]))
            mejor.append((np.mean(err), np.max(err), modo, round(a, 2), off))
mejor.sort()
for m in mejor[:8]: print(m)

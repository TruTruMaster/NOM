import json, re, math, sys
import numpy as np
datos = json.load(open(sys.argv[1])); path = sys.argv[2]; D = float(sys.argv[3]); acc = dec = float(sys.argv[4]) if len(sys.argv) > 4 else 0.5
base = datos[0][1][:2]
def warp(u):
    a, d = acc, dec
    if not (a or d): return u
    R = 1 / (1 - a / 2 - d / 2)
    if u < a * D: return R * u * u / (2 * a * D)
    if u <= (1 - d) * D: return R * (u - a * D / 2)
    t = D - u; return D - R * t * t / (2 * d * D)
toks = re.findall(r"[MLCZEmlcze]|-?\d*\.?\d+", path)
segs = []; i = 0; x = y = 0; cmd = "M"; x0 = y0 = 0
while i < len(toks):
    if re.match(r"[A-Za-z]", toks[i]): cmd = toks[i]; i += 1
    C = cmd.upper(); rel = cmd.islower()
    if C in "EZ":
        if C == "Z" and (x, y) != (x0, y0): segs.append(("L", [(x, y), (x0, y0)]))
        continue
    if C == "M": x, y = float(toks[i]), float(toks[i + 1]); i += 2; x0, y0 = x, y; cmd = "L"; continue
    if C == "L": nx, ny = float(toks[i]), float(toks[i + 1]); i += 2; nx, ny = (x + nx, y + ny) if rel else (nx, ny); segs.append(("L", [(x, y), (nx, ny)])); x, y = nx, ny; continue
    if C == "C":
        v = [float(t) for t in toks[i:i + 6]]; i += 6
        p = [(v[0], v[1]), (v[2], v[3]), (v[4], v[5])]
        if rel: p = [(x + a, y + b) for a, b in p]
        segs.append(("C", [(x, y)] + p)); x, y = p[2]; continue
def punto(seg, t):
    if seg[0] == "L": (a, b), (c, d) = seg[1]; return a + (c - a) * t, b + (d - b) * t
    p0, p1, p2, p3 = seg[1]; u = 1 - t
    return tuple(u ** 3 * p0[k] + 3 * u * u * t * p1[k] + 3 * u * t * t * p2[k] + t ** 3 * p3[k] for k in range(2))
def densa(escala):
    pts = []
    for s in segs:
        for k in range(200): pts.append(punto(s, k / 200))
    pts.append(punto(segs[-1], 1))
    P = np.array(pts) * escala
    L = np.concatenate([[0], np.cumsum(np.hypot(*np.diff(P, axis=0).T))])
    return np.array(pts), L
modelos = {}
for nombre, esc in (("arco_fraccion", np.array([1, 1])), ("arco_px", np.array([960, 540]))):
    pts, L = densa(esc)
    def f(s, pts=pts, L=L):
        l = s * L[-1]; k = np.searchsorted(L, l); k = min(max(k, 1), len(L) - 1); r = (l - L[k - 1]) / max(1e-12, L[k] - L[k - 1])
        return pts[k - 1] + (pts[k] - pts[k - 1]) * r
    modelos[nombre] = f
def param(s):
    n = len(segs); q = min(n - 1e-9, s * n); k = int(q); return np.array(punto(segs[k], q - k))
modelos["parametro"] = param
for nombre, f in modelos.items():
    err = []
    for t, v in datos:
        if v is None or t > D: continue
        s = warp(min(D, t)) / D
        x, y = f(s)
        px, py = base[0] + x * 960, base[1] + y * 540
        err.append(math.hypot(px - v[0], py - v[1]))
    print(f"{nombre:14s} error medio {np.mean(err):6.2f} px, máximo {np.max(err):6.2f} px")

# fracción del recorrido (por largo en px y por parámetro) que mejor explica cada posición observada
pts, L = densa(np.array([960, 540]))
Npar = len(segs)
print("t(ms)  s_lineal  smil(acc/dec)  frac_largo_px  frac_param  ")
for t, v in datos:
    if v is None or t > D + 100: continue
    obs = np.array([(v[0] - base[0]) / 960, (v[1] - base[1]) / 540])
    k = np.argmin(np.hypot(*(pts - obs).T * np.array([[960], [540]])))
    s_lin = min(1, t / D)
    print(f"{t:5d}  {s_lin:6.3f}  {warp(min(D, t)) / D:6.3f}  {L[k] / L[-1]:6.3f}  {k / (len(pts) - 1):6.3f}")

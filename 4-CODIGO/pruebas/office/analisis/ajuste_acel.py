import json, numpy as np
info = json.load(open('/home/user/.cache/nom-pruebas/office-real/archivos/a04-aceleracion.json'))['diapositivas']
def obs_de(n):
    d = json.load(open(f"acel{n}.json"))
    x0 = d[0][1][0]
    return [(t / 2000, (v[0] - x0) / 480) for t, v in d if v and t <= 2100]
def smil(s, a, b):
    s = min(1, max(0, s))
    if a + b <= 0: return s
    if a + b > 1: k = 1 / (a + b); a *= k; b *= k
    R = 1 / (1 - a / 2 - b / 2)
    if s < a: return R * s * s / (2 * a)
    if s <= 1 - b: return R * (s - a / 2)
    t = 1 - s; return 1 - R * t * t / (2 * b)
for x in info[:16]:
    n = x['diapo']; a, b = x['acel'], x['desac']
    o = obs_de(n)
    e = [smil(s, a, b) - p for s, p in o if s <= 1.0]
    rms = float(np.sqrt(np.mean(np.square(e))))
    muestra = [(round(s, 2), round(p, 3), round(smil(s, a, b), 3)) for s, p in o[::6] if s <= 1.0]
    print(f"a={a} d={b}: rms SMIL {rms:.4f}  ", muestra)

def ppt(s, a, b):
    s = min(1, max(0, s))
    if a + b <= 0: return s
    if a + b > 1: k = 1 / (a + b); a *= k; b *= k
    R = 1 / (1 - a / 2 - b / 2)
    if s < a: u = s / a; return R * a * (u ** 3 - u ** 4 / 2)
    if s <= 1 - b: return R * (s - a / 2)
    u = (1 - s) / b; return 1 - R * b * (u ** 3 - u ** 4 / 2)
print("---- modelo con rampa suave (smoothstep en la velocidad)")
tot = {}
for off in (0, 8, 16, 24):
    errs = []
    for x in info[:16]:
        n = x['diapo']; a, b = x['acel'], x['desac']
        o = obs_de(n)
        e = [ppt(s - off / 2000, a, b) - p for s, p in o if s <= 1.0]
        errs.append(float(np.sqrt(np.mean(np.square(e)))))
    tot[off] = errs
    print(f"desfase {off} ms: rms medio {np.mean(errs):.4f}, peor {max(errs):.4f}", [round(v, 4) for v in errs])

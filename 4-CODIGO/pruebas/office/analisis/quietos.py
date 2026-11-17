import pickle, json, subprocess, numpy as np
V = "/home/user/.cache/nom-pruebas/office-real/archivos/a03-transiciones.mp4"
por, ts, tr = pickle.load(open(V + ".marcas.pkl", "rb"))
W, H = 160, 90
p = subprocess.Popen(["ffmpeg", "-v", "error", "-i", V, "-fps_mode", "passthrough", "-vf", f"scale={W}:{H}:flags=area", "-f", "rawvideo", "-pix_fmt", "gray", "-"], stdout=subprocess.PIPE)
prev, difs = None, []
while True:
    b = p.stdout.read(W * H)
    if len(b) < W * H: break
    a = np.frombuffer(b, np.uint8).astype(np.int16).reshape(H, W)
    difs.append(0.0 if prev is None else float(np.abs(a - prev).mean())); prev = a
n = len(difs)
# cuadros quietos: igual al anterior
quieto = [d < 0.15 for d in difs]
runs = []; i = 0
while i < n:
    if quieto[i]:
        j = i
        while j + 1 < n and quieto[j + 1]: j += 1
        runs.append((i - 1, j)); i = j + 1
    else: i += 1
# el rato quieto de cada diapositiva: el tramo quieto (desde el cuadro anterior) de >= 0.5 s con su marca
quietas = {}
for a0, a1 in runs:
    if a0 < 0: a0 = 0
    if ts[min(a1 + 1, n - 1)] - ts[a0] < 0.5: continue
    m = [por[k] for k in range(a0, a1 + 1) if por[k] is not None]
    if not m: continue
    v = max(set(m), key=m.count)
    quietas.setdefault(v, (a0, a1))
print(len(quietas), "diapositivas con rato quieto")
info = json.load(open(V.replace(".mp4", ".json")))
res = {}
for x in sorted(info["diapositivas"], key=lambda x: x["diapo"]):
    k = x["diapo"]
    if k - 1 in quietas and k in quietas:
        ini = ts[quietas[k - 1][1] + 1] if quietas[k - 1][1] + 1 < n else None   # primer cuadro que cambia
        fin = ts[quietas[k][0]]  # primer cuadro quieto de la siguiente (ya terminó)
        res[k] = (round(ini, 3), round(fin, 3), round(fin - ini, 3), x["transicion"])
dur = [v[2] for v in res.values()]
print(len(res), "transiciones; duración medida: media", round(np.mean(dur), 3), "mín", min(dur), "máx", max(dur))
print([v for k, v in sorted(res.items()) if abs(v[2] - 2.0) > 0.15][:30])
json.dump({str(k): v for k, v in res.items()}, open("/home/user/NOM/trabajo/animaciones-17nov/a03-tiempos.json", "w"))

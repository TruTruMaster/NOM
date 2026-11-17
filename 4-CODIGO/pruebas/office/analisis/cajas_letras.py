import pickle, subprocess, json, numpy as np
V = "/home/user/.cache/nom-pruebas/office-real/archivos/a05-letras.mp4"
por, ts, tr = pickle.load(open(V + ".marcas.pkl", "rb"))
info = json.load(open(V.replace(".mp4", ".json")))
def cuadros(d):
    f0, f1 = tr[d]
    p = subprocess.Popen(["ffmpeg", "-v", "error", "-i", V, "-fps_mode", "passthrough", "-vf", f"select=between(n\\,{f0}\\,{f1})", "-f", "rawvideo", "-pix_fmt", "gray", "-"], stdout=subprocess.PIPE)
    out = []
    for f in range(f0, f1 + 1):
        b = p.stdout.read(1280 * 720)
        if len(b) < 1280 * 720: break
        out.append(((ts[f] - ts[f0]) * 1000, np.frombuffer(b, np.uint8).reshape(720, 1280)[:690].astype(int)))
    return out
K = 0.75  # px del video -> pt
res = {}
for x in info["diapositivas"]:
    d = x["diapo"]; cs = cuadros(d); fin = cs[-1][1]
    cols = np.nonzero((fin < 128).any(axis=0))[0]
    # la H: el primer grupo de columnas
    g0 = cols[0]; g1 = g0
    for c in cols[1:]:
        if c - g1 > 2: break
        g1 = c
    ysH = np.nonzero((fin[:, g0:g1 + 1] < 128).any(axis=1))[0]
    topH, basH = ysH.min() * K, (ysH.max() + 1) * K
    clave = x["efecto"]
    if clave.startswith("vuela"):
        pts = []
        for t, im in cs:
            ys = np.nonzero((im[:, g0:g1 + 1] < 128).any(axis=1))[0]
            if len(ys) and t < 520 and ys.min() * K > topH + 2: pts.append((t, ys.min() * K))
        if len(pts) >= 2:
            (t1, y1), (t2, y2) = pts[0], pts[-1]
            v = (y1 - y2) / (t2 - t1)
            D = v * 500
        else: D = None
        res[clave] = {"topH": round(topH, 2), "baseH": round(basH, 2), "D": round(D, 2) if D else None}
    else:
        c = None
        for t, im in cs[1:8]:
            ys, xs = np.nonzero(im[:, g0 - 20:g1 + 20] < 200)
            if 5 < len(ys) < 400: c = round(ys.mean() * K, 2); break
        res[clave] = {"topH": round(topH, 2), "baseH": round(basH, 2), "centroZoom": c}
for k, v in res.items(): print(k, v)
json.dump(res, open("cajas_letras.json", "w"))

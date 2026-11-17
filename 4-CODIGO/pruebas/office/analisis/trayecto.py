"""Centro de la figura azul en cada cuadro de una diapositiva del video (en px de la diapositiva de 960×540)."""
import sys, pickle, subprocess, numpy as np, json
V = sys.argv[1]; d = int(sys.argv[2])
por, ts, tr = pickle.load(open(V + ".marcas.pkl", "rb"))
f0, f1 = tr[d]
p = subprocess.Popen(["ffmpeg", "-v", "error", "-i", V, "-fps_mode", "passthrough", "-vf", f"select=between(n\\,{f0}\\,{f1})", "-f", "rawvideo", "-pix_fmt", "rgb24", "-"], stdout=subprocess.PIPE)
res = []
for f in range(f0, f1 + 1):
    b = p.stdout.read(1280 * 720 * 3)
    if len(b) < 1280 * 720 * 3: break
    a = np.frombuffer(b, np.uint8).reshape(720, 1280, 3).astype(int)
    a = a[:690]
    azul = (abs(a[..., 0] - 0x2F) < 40) & (abs(a[..., 1] - 0x55) < 40) & (abs(a[..., 2] - 0x97) < 40)
    ys, xs = np.nonzero(azul)
    if len(xs) < 20: res.append((round((ts[f] - ts[f0]) * 1000), None)); continue
    res.append((round((ts[f] - ts[f0]) * 1000), [round(xs.mean() * 0.75, 1), round(ys.mean() * 0.75, 1), round((xs.max() - xs.min()) * 0.75, 1), round((ys.max() - ys.min()) * 0.75, 1)]))
json.dump(res, open(sys.argv[3], "w"))
for r in res[::4]: print(r)

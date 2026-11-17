"""Saca cuadros a resolución completa de una diapositiva del video de PowerPoint: cuadros_diapo.py VIDEO DIAPO salida_prefijo t1,t2,... (ms desde que empieza)"""
import sys, json, subprocess
sys.path.insert(0, "/home/user/NOM/4-CODIGO/pruebas/office")
import animaciones_real as R
video, diapo, pref, tiempos = sys.argv[1], int(sys.argv[2]), sys.argv[3], [float(x) for x in sys.argv[4].split(",")]
import pickle, os
cache = video + ".marcas.pkl"
if os.path.exists(cache): por, ts, tramos = pickle.load(open(cache, "rb"))
else:
    por, ts, tramos = R.marcas(video); pickle.dump((por, ts, tramos), open(cache, "wb"))
f0, f1 = tramos[diapo]
for t in tiempos:
    # el cuadro más cercano a t
    f = min(range(f0, f1 + 1), key=lambda i: abs((ts[i] - ts[f0]) * 1000 - t))
    out = f"{pref}-{int(t):05d}.png"
    subprocess.run(["ffmpeg", "-v", "error", "-y", "-i", video, "-vf", f"select=eq(n\\,{f})", "-vframes", "1", out])
    print(out, f, round((ts[f] - ts[f0]) * 1000))

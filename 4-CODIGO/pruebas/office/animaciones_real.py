#!/usr/bin/env python3
"""Mide las animaciones y transiciones de NOM contra el video que hace PowerPoint de verdad (CreateVideo).

  animaciones_real.py ARCHIVO.pptx [--salida DIR] [--diapos 1-40] [--cuadros 8] [--desfase 0] [--transiciones]

ARCHIVO lo hace `windows/animaciones_office.py` (catalogo, opciones o transiciones) y su video `ARCHIVO.mp4` con
`animaciones_office.py video`: una diapositiva por efecto, que empieza sola, con una marca binaria abajo a la izquierda
(su número y una comprobación). Se lee la marca de cada cuadro del video para saber en qué cuadro empieza cada
diapositiva; se toman cuadros a lo largo de la animación y NOM pinta la misma diapositiva en el mismo instante con su
motor de animaciones (con el servidor local en el puerto 8099). Se compara píxel por píxel (sin la marca) y se deja
el parecido de cada diapositiva, el reporte y, de lo que sale peor, las imágenes lado a lado.
"""
import argparse, base64, json, subprocess, sys
from pathlib import Path
import numpy as np
from PIL import Image, ImageDraw

AQUI = Path(__file__).parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import nomprueba as N, prueba_ui as U, comparar_real as C  # noqa: E402

REAL = C.REAL
ANCHO, ALTO = 640, 360   # tamaño de la comparación (el video es de 1280×720)


def marcas(video):
    """Número de diapositiva de cada cuadro (o None) y el tiempo de cada cuadro."""
    info = json.loads(subprocess.run(["ffprobe", "-v", "error", "-show_entries", "stream=width,height", "-of", "json", str(video)], capture_output=True, text=True).stdout)["streams"][0]
    W = info["width"]; k = W / 960.0
    par = lambda v: int(v) // 2 * 2
    x0, y0, ww, hh = 0, par(524 * k), par(160 * k), par(16 * k)
    p = subprocess.Popen(["ffmpeg", "-v", "error", "-i", str(video), "-fps_mode", "passthrough", "-vf", f"crop={ww}:{hh}:{x0}:{y0}", "-f", "rawvideo", "-pix_fmt", "gray", "-"], stdout=subprocess.PIPE)
    ts = subprocess.run(["ffprobe", "-v", "error", "-select_streams", "v", "-show_entries", "frame=pts_time", "-of", "csv=p=0", str(video)], capture_output=True, text=True).stdout.split()
    ts = [float(t.strip(",")) for t in ts if t.strip(",").strip()]
    sl = []
    while True:
        b = p.stdout.read(ww * hh)
        if len(b) < ww * hh: break
        a = np.frombuffer(b, np.uint8).reshape(hh, ww)
        v = 0
        for i in range(12):
            cx, cy = int((6 + i * 12 + 5) * k), int((528 + 5) * k) - y0
            v = v * 2 + (1 if a[cy - 2:cy + 3, cx - 2:cx + 3].mean() < 128 else 0)
        num, chk = v >> 3, v & 7
        sl.append(num if num > 0 and num % 7 == chk else None)
    tramos = {}
    for i, v in enumerate(sl):
        if v is None: continue
        tramos.setdefault(v, [i, i])[1] = i
    return sl, ts[:len(sl)], tramos


def cuadros_video(video, indices):
    """Los cuadros pedidos del video, a ANCHO×ALTO (RGB)."""
    W, H = ANCHO, ALTO
    p = subprocess.Popen(["ffmpeg", "-v", "error", "-i", str(video), "-fps_mode", "passthrough", "-vf", f"scale={W}:{H}:flags=area", "-f", "rawvideo", "-pix_fmt", "rgb24", "-"], stdout=subprocess.PIPE)
    quiero, out, n = set(indices), {}, 0
    ultimo = max(indices) if indices else -1
    while n <= ultimo:
        b = p.stdout.read(W * H * 3)
        if len(b) < W * H * 3: break
        if n in quiero: out[n] = np.frombuffer(b, np.uint8).reshape(H, W, 3).copy()
        n += 1
    p.kill()
    return out


def sin_marca(a):
    a = a.copy()
    a[int(ALTO * 520 / 540):, : int(ANCHO * 165 / 960)] = 255
    return a


def suavizar(a, r=3):
    """Promedio en cuadros de (2r+1)² px: para los efectos al azar, que PowerPoint dibuja distinto cada vez."""
    a = a.astype(np.float32)
    k = 2 * r + 1
    c = np.cumsum(np.cumsum(np.pad(a, ((r + 1, r), (r + 1, r), (0, 0)), mode="edge"), axis=0), axis=1)
    return (c[k:, k:] - c[:-k, k:] - c[k:, :-k] + c[:-k, :-k]) / (k * k)


def parecido_np(a, b, azar=False):
    """100 = iguales: diferencia media (RGB) en la zona donde alguna de las dos tiene algo pintado (sin la marca)."""
    a, b = sin_marca(a).astype(np.int16), sin_marca(b).astype(np.int16)
    if azar: a, b = suavizar(a), suavizar(b)
    ga, gb = a.mean(axis=2), b.mean(axis=2)
    zona = (ga < 245) | (gb < 245)
    if not zona.any(): return 100.0
    z = zona.copy()
    z[1:, :] |= zona[:-1, :]; z[:-1, :] |= zona[1:, :]; z[:, 1:] |= zona[:, :-1]; z[:, :-1] |= zona[:, 1:]
    d = np.abs(a - b).mean(axis=2)[z]
    return round(float(100 - d.mean() / 2.55), 1)


def lado(a, b, ruta, titulo):
    im = Image.new("RGB", (ANCHO * 2 + 10, ALTO + 24), "white")
    im.paste(Image.fromarray(a), (0, 24)); im.paste(Image.fromarray(b), (ANCHO + 10, 24))
    d = ImageDraw.Draw(im); d.text((4, 4), "PowerPoint", fill="black"); d.text((ANCHO + 14, 4), "NOM  " + titulo, fill="black")
    im.save(ruta)


JS_PREPARAR = r"""
async ([i, k]) => {
  const doc = NOM.editorActual.doc, p = doc.paginas()[i];
  let cont = document.getElementById('__anim_prueba');
  if (!cont) { cont = document.createElement('div'); cont.id = '__anim_prueba'; document.body.append(cont); }
  const T = NOM.diapo.tam(doc);
  cont.style.cssText = `position:fixed;left:0;top:0;z-index:99999;background:#fff;overflow:hidden;width:${T.w * k}px;height:${T.h * k}px`;
  cont.innerHTML = `<div class="diapo diapo-foto" style="${NOM.diapo.estiloDiapo(doc, p)}position:absolute;left:0;top:0;transform-origin:0 0;transform:scale(${k});border-radius:0;box-shadow:none;margin:0"></div>`;
  const d = cont.firstElementChild;
  if (window.__motor) window.__motor.detener();
  if (NOM.diapo.aplicar) NOM.diapo.aplicar(d, doc, p);
  d.style.position = 'absolute'; d.style.left = '0px'; d.style.top = '0px'; d.style.transformOrigin = '0 0'; d.style.transform = `scale(${k})`;
  d.innerHTML = NOM.lienzoVista.htmlDiapositiva(doc, p);
  await document.fonts.ready;
  const imgs = [...d.querySelectorAll('img')]; await Promise.all(imgs.map((im) => im.complete ? 0 : new Promise((r) => { im.onload = im.onerror = r; })));
  window.__motor = NOM.lienzoVista.motor(d, doc, p, () => k);
  const lista = doc.valor('_doc', 'anim:' + p) || [];
  return { pasos: window.__motor ? window.__motor.pasos : -1, dur: window.__motor ? window.__motor.duraciones() : [], efectos: lista.length, trans: doc.valor('_doc', 'trans:' + p) || null };
}
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("archivo")
    ap.add_argument("--salida", default=str(Path.home() / ".cache/nom-pruebas/animaciones"))
    ap.add_argument("--diapos", default="")
    ap.add_argument("--cuadros", type=int, default=8)
    ap.add_argument("--desfase", type=float, default=0.0, help="ms que se suman al tiempo de cada cuadro")
    a = ap.parse_args()
    arch = Path(a.archivo)
    if not arch.exists(): arch = REAL / "archivos" / arch.name
    video = arch.with_suffix(".mp4")
    info = json.loads(arch.with_suffix(".json").read_text(encoding="utf8"))
    lista = info.get("diapositivas", [])
    if a.diapos:
        r0, _, r1 = a.diapos.partition("-"); r0 = int(r0); r1 = int(r1 or r0)
        lista = [x for x in lista if r0 <= x["diapo"] <= r1]
    sal = Path(a.salida) / arch.stem; sal.mkdir(parents=True, exist_ok=True)
    for v in sal.glob("*.png"): v.unlink()
    print("leyendo las marcas del video…", flush=True)
    por_cuadro, ts, tramos = marcas(video)
    # qué cuadros se comparan de cada diapositiva
    plan = []
    for x in lista:
        tr = tramos.get(x["diapo"])
        if not tr: continue
        f0, f1 = tr
        dur = float(x.get("dur_total") or 0)
        if not (dur < 1e6): dur = 5.0
        fin = min(f1, f0 + int((dur + 0.15) / 0.033) + 1)
        n = max(2, a.cuadros)
        idx = sorted(set([f0 + round(j * (fin - f0) / (n - 1)) for j in range(n)] + [min(f1, fin + 3)]))
        plan.append((x, f0, idx))
    todos = sorted({i for _, _, l in plan for i in l})
    print(f"{len(plan)} diapositivas, {len(todos)} cuadros del video…", flush=True)
    fr = cuadros_video(video, todos)
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False)
    nav.ctx.close(); nav.ctx = nav.b.new_context(viewport={"width": 1300, "height": 760}); nav.p = nav.ctx.new_page(); nav.p.on("pageerror", lambda e: nav.errores.append(str(e)))
    nav.entrar(c)
    p = nav.p
    rep = {"archivo": arch.name, "diapositivas": []}
    try:
        C.importar(nav, esp, arch)
        k = 1280 / 960
        for x, f0, idx in plan:
            info_d = p.evaluate(JS_PREPARAR, [x["diapo"] - 1, k])
            vals = []
            for j, f in enumerate(idx):
                t = (ts[f] - ts[f0]) * 1000 + a.desfase
                p.evaluate("(t) => { if (window.__motor) window.__motor.cuadro(0, t); }", t)
                png = p.screenshot(clip={"x": 0, "y": 0, "width": 1280, "height": 720})
                im = Image.open(__import__("io").BytesIO(png)).convert("RGB").resize((ANCHO, ALTO), Image.BOX)
                b = np.asarray(im)
                ref = fr.get(f)
                if ref is None: continue
                v = parecido_np(ref, b, azar=bool(__import__("re").search(r"Dissolve|RandomBars|Disolver|Barras", str(x.get("efecto") or ""))))
                vals.append(v)
                if v < 92: lado(ref, b, sal / f"d{x['diapo']:03d}-c{j:02d}-{v:.0f}.png", f"{x.get('efecto') or x.get('transicion')} t={t:.0f} ms")
            d = {"diapo": x["diapo"], "efecto": x.get("efecto") or x.get("transicion"), "salida": x.get("salida"), "parecido": round(sum(vals) / max(1, len(vals)), 1),
                 "minimo": min(vals) if vals else None, "cuadros": vals, "nom": info_d}
            rep["diapositivas"].append(d)
            print(f"  {x['diapo']:3d} {str(d['efecto'])[:28]:28s}{' (salida)' if x.get('salida') else '         '} {d['parecido']:5.1f} % (mín {d['minimo']})", flush=True)
    finally:
        nav.cerrar()
    vs = [d["parecido"] for d in rep["diapositivas"]]
    rep["parecido_medio"] = round(sum(vs) / max(1, len(vs)), 1)
    rep["errores_js"] = nav.errores[-10:]
    (sal / "reporte.json").write_text(json.dumps(rep, ensure_ascii=False, indent=1))
    print(f"{arch.name}: {len(vs)} diapositivas, parecido medio {rep['parecido_medio']} %; ≥ 98 %: {sum(1 for v in vs if v >= 98)}, < 90 %: {sum(1 for v in vs if v < 90)}", flush=True)
    if rep["errores_js"]: print("errores:", rep["errores_js"])


if __name__ == "__main__":
    main()

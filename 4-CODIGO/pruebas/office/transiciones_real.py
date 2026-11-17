#!/usr/bin/env python3
"""Mide las transiciones de NOM contra el video que hace PowerPoint de verdad.

  transiciones_real.py [a03-transiciones.pptx] [--salida DIR] [--diapos 5-40] [--cuadros 8]

El archivo lo hace `windows/animaciones_office.py transiciones`: cada diapositiva con un fondo de color, un texto y una
transición distinta de 2 s, y avanza sola. En el video se busca dónde empieza y termina cada transición (los cuadros
que cambian después de un rato quieto), se toman cuadros a lo largo de ella y NOM pinta la misma transición, de la
diapositiva anterior a esa, en el mismo avance. Se compara píxel por píxel y deja el reporte y lo peor lado a lado.
"""
import argparse, io, json, subprocess, sys
from pathlib import Path
import numpy as np
from PIL import Image

AQUI = Path(__file__).parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import nomprueba as N, comparar_real as C, animaciones_real as R  # noqa: E402

JS_PREPARAR = r"""
async ([i, k, manual]) => {
  const doc = NOM.editorActual.doc, pags = doc.paginas(), p = pags[i], ant = pags[i - 1];
  let cont = document.getElementById('__anim_prueba');
  if (!cont) { cont = document.createElement('div'); cont.id = '__anim_prueba'; document.body.append(cont); }
  const T = NOM.diapo.tam(doc);
  cont.style.cssText = `position:fixed;left:0;top:0;z-index:99999;background:#fff;overflow:hidden;width:${T.w * k}px;height:${T.h * k}px`;
  cont.innerHTML = '<div class="diapo diapo-foto"></div>';
  const d = cont.firstElementChild;
  // la diapositiva anterior, como se veía (su fondo y su HTML)
  NOM.diapo.aplicar(d, doc, ant);
  d.innerHTML = NOM.lienzoVista.htmlDiapositiva(doc, ant);
  const viejo = { html: d.innerHTML, fondo: d.style.background };
  NOM.diapo.aplicar(d, doc, p);
  d.style.position = 'absolute'; d.style.left = '0px'; d.style.top = '0px'; d.style.transformOrigin = '0 0'; d.style.transform = `scale(${k})`;
  d.style.borderRadius = '0'; d.style.boxShadow = 'none'; d.style.margin = '0';
  d.innerHTML = NOM.lienzoVista.htmlDiapositiva(doc, p);
  await document.fonts.ready;
  const tr = doc.valor('_doc', 'transOffice:' + p);
  if (window.__trans) { try { window.__trans.fin(); } catch (e) {} }
  window.__trans = tr ? NOM.transOffice.jugar(d, viejo, tr, T, { manual: true }) : null;
  return { tr };
}
"""


def inicios(video):
    """Cuándo empieza cada transición: el primer cuadro que cambia después del rato quieto de la diapositiva anterior
    (cada diapositiva queda quieta 1 s y su marca se lee). Devuelve {diapositiva: segundo de inicio} y los tiempos."""
    import pickle
    cache = Path(str(video) + ".marcas.pkl")
    if cache.exists(): por, ts, _ = pickle.load(open(cache, "rb"))
    else: por, ts, tr = R.marcas(video); pickle.dump((por, ts, tr), open(cache, "wb"))
    W, H = 160, 90
    p = subprocess.Popen(["ffmpeg", "-v", "error", "-i", str(video), "-fps_mode", "passthrough", "-vf", f"scale={W}:{H}:flags=area", "-f", "rawvideo", "-pix_fmt", "gray", "-"], stdout=subprocess.PIPE)
    prev, difs = None, []
    while True:
        b = p.stdout.read(W * H)
        if len(b) < W * H: break
        a = np.frombuffer(b, np.uint8).astype(np.int16).reshape(H, W)
        difs.append(0.0 if prev is None else float(np.abs(a - prev).mean())); prev = a
    n = min(len(difs), len(ts))
    quieto = [d < 0.15 for d in difs[:n]]
    quietas, i = {}, 0
    while i < n:
        if quieto[i]:
            j = i
            while j + 1 < n and quieto[j + 1]: j += 1
            a0 = max(0, i - 1)
            if ts[min(j + 1, n - 1)] - ts[a0] >= 0.5:
                m = [por[k] for k in range(a0, j + 1) if por[k] is not None]
                if m: quietas.setdefault(max(set(m), key=m.count), (a0, j))
            i = j + 1
        else: i += 1
    ini = {k: (quietas[k - 1][1] + 1) for k in quietas if k - 1 in quietas and quietas[k - 1][1] + 1 < n}
    return ini, ts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("archivo", nargs="?", default="a03-transiciones.pptx")
    ap.add_argument("--salida", default=str(Path.home() / ".cache/nom-pruebas/animaciones"))
    ap.add_argument("--diapos", default="")
    ap.add_argument("--cuadros", type=int, default=8)
    a = ap.parse_args()
    arch = Path(a.archivo)
    if not arch.exists(): arch = R.REAL / "archivos" / arch.name
    video = arch.with_suffix(".mp4")
    info = json.loads(arch.with_suffix(".json").read_text(encoding="utf8"))
    lista = sorted(info.get("diapositivas", []), key=lambda x: x["diapo"])
    sal = Path(a.salida) / arch.stem; sal.mkdir(parents=True, exist_ok=True)
    for v in sal.glob("*.png"): v.unlink()
    print("buscando las transiciones en el video…", flush=True)
    ini, ts = inicios(video)
    if a.diapos:
        r0, _, r1 = a.diapos.partition("-"); r0 = int(r0); r1 = int(r1 or r0)
    plan = []
    for x in lista:
        k = x["diapo"]
        if k not in ini or str(x.get("transicion", "")) in ("ppEffectNone", "ppEffectCut"): continue
        if a.diapos and not (r0 <= k <= r1): continue
        f0 = ini[k] - 1   # el último cuadro quieto (avance 0)
        t0 = ts[ini[k]] - 0.0165   # el cambio empezó entre ese cuadro y el anterior
        fin = t0 + 2.0
        idx = [f for f in range(f0, len(ts)) if ts[f] <= fin + 0.05]
        n = max(2, a.cuadros)
        sel = sorted(set(idx[round(j * (len(idx) - 1) / (n - 1))] for j in range(n)))
        plan.append((x, t0, 2.0, sel))
    print(f"{len(plan)} transiciones a medir", flush=True)
    fr = R.cuadros_video(video, sorted({f for _, _, _, l in plan for f in l}))
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False)
    nav.ctx.close(); nav.ctx = nav.b.new_context(viewport={"width": 1300, "height": 760}); nav.p = nav.ctx.new_page(); nav.p.on("pageerror", lambda e: nav.errores.append(str(e)))
    nav.entrar(c)
    p = nav.p
    rep = {"archivo": arch.name, "transiciones": []}
    try:
        C.importar(nav, esp, arch)
        k = 1280 / 960
        for x, t0, durv, sel in plan:
            info_d = p.evaluate(JS_PREPARAR, [x["diapo"] - 1, k, True])
            dur = (info_d.get("tr") or {}).get("dur") or 2000
            vals = []
            for j, f in enumerate(sel):
                prog = max(0.0, (ts[f] - t0) * 1000 / dur)
                p.evaluate("(q) => { if (window.__trans) window.__trans.cuadro(q); }", prog)
                png = p.screenshot(clip={"x": 0, "y": 0, "width": 1280, "height": 720})
                b = np.asarray(Image.open(io.BytesIO(png)).convert("RGB").resize((R.ANCHO, R.ALTO), Image.BOX))
                ref = fr.get(f)
                if ref is None: continue
                v = R.parecido_np(ref, b, azar=bool(__import__("re").search(r"Dissolve|RandomBars|Glitter|Honeycomb|Shred|Vortex|Ripple", x.get("transicion", ""))))
                vals.append(v)
                if v < 90: R.lado(ref, b, sal / f"t{x['diapo']:03d}-c{j:02d}-{v:.0f}.png", f"{x['transicion'][8:]} p={prog:.2f}")
            d = {"diapo": x["diapo"], "transicion": x["transicion"], "dur_video": round(durv, 3), "parecido": round(sum(vals) / max(1, len(vals)), 1), "minimo": min(vals) if vals else None, "cuadros": vals}
            rep["transiciones"].append(d)
            print(f"  {x['diapo']:3d} {x['transicion'][8:]:28s} {d['parecido']:5.1f} % (mín {d['minimo']})  [video {durv:.2f} s]", flush=True)
    finally:
        nav.cerrar()
    vs = [d["parecido"] for d in rep["transiciones"]]
    rep["parecido_medio"] = round(sum(vs) / max(1, len(vs)), 1)
    rep["errores_js"] = nav.errores[-10:]
    (sal / "reporte.json").write_text(json.dumps(rep, ensure_ascii=False, indent=1))
    print(f"{arch.name}: {len(vs)} transiciones, parecido medio {rep['parecido_medio']} %; ≥ 95 %: {sum(1 for v in vs if v >= 95)}, < 85 %: {sum(1 for v in vs if v < 85)}", flush=True)
    if rep["errores_js"]: print("errores:", rep["errores_js"])


if __name__ == "__main__":
    main()

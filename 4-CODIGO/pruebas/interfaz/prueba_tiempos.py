#!/usr/bin/env python3
"""Mide cuánto tarda NOM en lo que una persona nota, contra el sitio real (por omisión www.nom.qd.je).

Con conexión normal y con la «3G lenta» del navegador (Chromium, por CDP):
  · primera carga (sin caché) y carga con caché, hasta ver el formulario de entrar
  · entrar → inicio
  · abrir cada tipo de documento (texto, hoja, presentación, diagrama) y un Word, Excel y PowerPoint grandes importados
  · tecla → letra en pantalla (Event Timing del navegador: de la tecla al siguiente cuadro pintado)
  · cambio de otra persona → visible en mi pantalla (dos páginas, el mismo reloj)
  · cambiar de diapositiva y de hoja, presentar
  · volver de sin red: cuánto tarda en llegar a la otra persona lo escrito sin conexión
  · tareas largas (Long Tasks > 50 ms) en todo el recorrido; las de > 100 ms se listan

Metas (DISENO-UI §5): tecla → pantalla < 50 ms; abrir un documento normal < 1 s; cambio en vivo < 300 ms;
nada que congele más de 100 ms.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_tiempos.py SALIDA [--solo normal|3g] [--office CARPETA]

Crea dos cuentas desechables y las borra al final. No hace carga: una sola persona a la vez.
"""
import argparse, base64, json, os, statistics, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from playwright.sync_api import sync_playwright

B = N.B
# Perfiles de red de las herramientas de desarrollo de Chromium («Slow 3G» y «Fast 3G»).
REDES = {
    "normal": None,
    "3g": dict(offline=False, latency=2000, downloadThroughput=400 * 1024 / 8 * 0.8, uploadThroughput=400 * 1024 / 8 * 0.8),
}
OBS = r"""() => {
  if (window.__tiempos) return;
  const t = window.__tiempos = { largas: [], eventos: [] };
  try { new PerformanceObserver(l => l.getEntries().forEach(e => t.largas.push([Math.round(e.startTime), Math.round(e.duration), location.hash]))).observe({ type: 'longtask', buffered: true }); } catch (e) {}
  try { new PerformanceObserver(l => l.getEntries().forEach(e => t.eventos.push([e.name, Math.round(e.duration), Math.round(e.processingEnd - e.startTime)]))).observe({ type: 'event', durationThreshold: 16, buffered: true }); } catch (e) {}
}"""
FILAS = []


def fila(red, que, ms, meta=None, nota=""):
    ok = "" if meta is None else ("cumple" if ms is not None and ms < meta else "NO cumple")
    FILAS.append((red, que, ms, meta, ok, nota))
    print(f"  [{red}] {que}: {'—' if ms is None else f'{ms:.0f} ms'}" + (f"  (meta < {meta} ms: {ok})" if meta else "") + (f"  {nota}" if nota else ""), flush=True)


def contexto(b, red, movil=False):
    opc = dict(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True, device_scale_factor=2) if movil else dict(viewport={"width": 1400, "height": 900})
    ctx = b.new_context(**opc)
    p = ctx.new_page()
    p.errores = []
    p.on("pageerror", lambda e: p.errores.append(str(e)[:300]))
    p.on("dialog", lambda d: d.accept())
    cdp = ctx.new_cdp_session(p)
    if REDES[red]:
        cdp.send("Network.enable")
        cdp.send("Network.emulateNetworkConditions", REDES[red])
    p.add_init_script("(" + OBS + ")()")
    return ctx, p


def entrar(p, c):
    p.goto(B + "/", wait_until="commit", timeout=180000); p.wait_for_function("() => window.NOM && NOM.api", timeout=120000)
    p.evaluate("() => { try { NOM.api.limpiarSesion(); } catch (e) {} }")
    p.goto(B + "/", wait_until="commit", timeout=180000); p.wait_for_selector("#u", timeout=120000)
    p.fill("#u", c["nombre"]); p.fill("#c", N.CL)
    t0 = time.perf_counter()
    p.click("button[type=submit]")
    p.wait_for_selector(".riel, .barra-movil", state="attached", timeout=120000)
    ms = (time.perf_counter() - t0) * 1000
    p.evaluate("() => document.querySelectorAll('.velo').forEach(v => v.remove())")
    return ms


def abrir(p, doc):
    """Abre un documento como lo hace la página (cambio de dirección dentro de la app) y mide hasta ver el editor armado."""
    p.evaluate("() => { location.hash = '#/'; }"); p.wait_for_timeout(600)
    ms = p.evaluate("""(id) => new Promise((ok) => {
      const t0 = performance.now(); location.hash = '#/doc/' + id;
      const listo = () => { const ed = NOM.editorActual; return ed && ed.doc && ed.vista && ed.vista.isConnected && ed.vista.querySelector('.papel, .wd-pagina, table, .lienzo, .mini, svg, [contenteditable]'); };
      const ver = () => listo() ? requestAnimationFrame(() => ok(performance.now() - t0)) : (performance.now() - t0 > 60000 ? ok(null) : setTimeout(ver, 5));
      ver(); })""", doc)
    p.wait_for_timeout(500)
    p.evaluate("() => { document.querySelectorAll('.velo').forEach(v => v.remove()); const b = [...document.querySelectorAll('.consejo button')].find(x => /Entendido/.test(x.textContent)); if (b) b.click(); }")
    return ms


def teclas(p, red):
    """Escribe 40 letras en el texto y lee del Event Timing cuánto tardó cada tecla en verse."""
    ed = p.evaluate("() => { const e = [...document.querySelectorAll('[contenteditable=true]')].find(x => x.offsetParent); if (!e) return null; const r = e.getBoundingClientRect(); return [r.x + 40, r.y + 20]; }")
    if not ed: fila(red, "tecla → letra en pantalla", None, 50, "no se encontró dónde escribir"); return
    p.mouse.click(*ed); p.wait_for_timeout(300)
    p.evaluate("() => { window.__tiempos.eventos = []; }")
    p.keyboard.type("La comida de hoy fue sopa de fideo y agua.", delay=60)
    p.wait_for_timeout(800)
    ev = p.evaluate("() => window.__tiempos.eventos")
    dur = sorted(d for (n, d, _) in ev if n in ("keydown", "keypress", "input", "beforeinput", "keyup"))
    # Event Timing solo reporta los de 16 ms o más (redondeados a 8 ms); los demás tardaron menos de 16 ms.
    tot = 42 * 2
    if not dur: fila(red, "tecla → letra en pantalla (p75 y peor)", 16, 50, "todas < 16 ms"); return
    p75 = dur[int(len(dur) * 0.75)] if len(dur) >= tot * 0.25 else 16
    fila(red, "tecla → letra en pantalla (p75)", p75, 50, f"{len(dur)} eventos ≥ 16 ms de ~{tot}; peor {dur[-1]} ms")


def en_vivo(pa, pb, red, n=5):
    """A escribe una marca; B la ve. Mismo equipo, mismo reloj (Date.now)."""
    res = []
    for i in range(n):
        marca = f"VIVO{i}X{int(time.time()*1000) % 100000}"
        pb.evaluate("""(m) => { window.__visto = null; const ob = new MutationObserver(() => { if (!window.__visto && document.body.innerText.includes(m)) { window.__visto = Date.now(); ob.disconnect(); } });
          ob.observe(document.body, { subtree: true, childList: true, characterData: true }); }""", marca)
        t0 = pa.evaluate("() => Date.now()")
        pa.keyboard.type(marca, delay=0)
        t_fin = pa.evaluate("() => Date.now()")
        for _ in range(200):
            v = pb.evaluate("() => window.__visto")
            if v: break
            pb.wait_for_timeout(50)
        if v: res.append(v - t_fin)
        pa.keyboard.press("Enter"); pa.wait_for_timeout(700)
    if res: fila(red, f"cambio de otra persona → visible (mediana de {len(res)})", statistics.median(res), 300, f"todas: {res}")
    else: fila(red, "cambio de otra persona → visible", None, 300, "B no vio la marca en 10 s")


def principal():
    ap = argparse.ArgumentParser()
    ap.add_argument("salida"); ap.add_argument("--solo", nargs="*", default=["normal", "3g"])
    ap.add_argument("--office", default=os.path.expanduser("~/NOM/trabajo/expo27/office"))
    ap.add_argument("--sin-office", action="store_true", help="no importar los Office grandes (más rápido)")
    a = ap.parse_args(); os.makedirs(a.salida, exist_ok=True)
    N.esperar_servidor()
    c1, c2 = N.cuenta(), N.cuenta()
    e = N.espacio(c1["token"], "Tiempos")
    N.api("PUT", f"/api/espacios/{e}/accesos", {"usuario": c2["nombre"], "permiso": "EDICION"}, c1["token"])
    docs = {m: N.documento(c1["token"], e, m, "Normal " + m.lower()) for m in ("TEXTO", "HOJA", "PRESENTACION", "DIAGRAMA")}
    pw = sync_playwright().start(); b = pw.chromium.launch()
    try:
        # Importar los archivos grandes una vez (con conexión normal) para poder medir abrirlos.
        ctx, p = contexto(b, "normal"); entrar(p, c1); abrir(p, docs["TEXTO"])
        grandes = {}
        for nom in (sorted(os.listdir(a.office)) if os.path.isdir(a.office) and not a.sin_office else []):
            ruta = os.path.join(a.office, nom); kb = os.path.getsize(ruta) // 1024
            b64 = base64.b64encode(open(ruta, "rb").read()).decode()
            t0 = time.perf_counter()
            imp = p.evaluate("""async ([b64, nom, esp]) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
              const t0 = performance.now(); const r = await NOM.formatos.importar(new File([u], nom)); const t1 = performance.now();
              // Como el botón «Importar» del espacio: las imágenes van al almacén y el documento se crea con referencias.
              const d = await NOM.binarios.crearEnEspacio({ id: esp }, r.nombre, r.modo, r.ops);
              return { ms: t1 - t0, subir: performance.now() - t1, n: r.ops.length, id: d && d.id };}""", [b64, nom, e])
            fila("normal", f"importar {nom} ({kb} KB) en la página", imp["ms"], None, f"{imp['n']} cambios")
            fila("normal", f"subir {nom} a la nube (imágenes y documento)", imp["subir"], None)
            if imp.get("id"): grandes[nom] = imp["id"]
        ctx.close()

        for red in a.solo:
            print(f"== red: {red}", flush=True)
            ctx, p = contexto(b, red)
            t0 = time.perf_counter(); p.goto(B + "/", wait_until="commit", timeout=180000); p.wait_for_selector("#u", timeout=180000)
            fila(red, "primera carga sin caché (hasta ver «Entrar»)", (time.perf_counter() - t0) * 1000, None)
            nav = p.evaluate("() => { const n = performance.getEntriesByType('navigation')[0]; const f = performance.getEntriesByName('first-contentful-paint')[0]; return { fcp: f ? f.startTime : null, bytes: performance.getEntriesByType('resource').reduce((s, r) => s + (r.transferSize || 0), n.transferSize || 0) }; }")
            fila(red, "primera carga: primer dibujo (FCP)", nav["fcp"], None, f"{nav['bytes']//1024} KB bajados")
            t0 = time.perf_counter(); p.reload(wait_until="commit", timeout=180000); p.wait_for_selector("#u", timeout=180000)
            fila(red, "carga con caché (hasta ver «Entrar»)", (time.perf_counter() - t0) * 1000, None)
            fila(red, "entrar → inicio", entrar(p, c1), None)
            for m, d in docs.items():
                fila(red, f"abrir {m.lower()} normal", abrir(p, d), 1000)
            for nom, d in grandes.items():
                fila(red, f"abrir {nom} importado", abrir(p, d), None)
            # Cambiar de diapositiva y de hoja en los grandes
            for nom, d in grandes.items():
                if nom.endswith(".pptx"):
                    abrir(p, d)
                    ms = p.evaluate("""() => new Promise(ok => { const m = [...document.querySelectorAll('button.mini')]; if (m.length < 2) return ok(null);
                      const t0 = performance.now(); m[1].click(); requestAnimationFrame(() => requestAnimationFrame(() => ok(performance.now() - t0))); })""")
                    fila(red, "cambiar de diapositiva (PowerPoint grande)", ms, 100)
                    t0 = time.perf_counter(); p.keyboard.press("F5")
                    try: p.wait_for_selector(".reproductor", timeout=30000); fila(red, "presentar (F5 → primera diapositiva)", (time.perf_counter() - t0) * 1000, 1000)
                    except Exception: fila(red, "presentar (F5)", None, 1000, "no apareció el reproductor")
                    p.keyboard.press("ArrowRight"); p.wait_for_timeout(400); p.keyboard.press("Escape"); p.wait_for_timeout(400); p.keyboard.press("Escape")
                if nom.endswith(".xlsx"):
                    abrir(p, d)
                    ms = p.evaluate("""() => new Promise(ok => { const h = [...document.querySelectorAll('[data-hoja]')]; if (h.length < 2) return ok(null);
                      const t0 = performance.now(); h[1].click(); requestAnimationFrame(() => requestAnimationFrame(() => ok(performance.now() - t0))); })""")
                    fila(red, "cambiar de hoja (Excel grande)", ms, 100, "" if ms is not None else "una sola hoja")
            # Escribir y en vivo
            abrir(p, docs["TEXTO"]); teclas(p, red)
            ctx2, q = contexto(b, red); entrar(q, c2); abrir(q, docs["TEXTO"]); q.wait_for_timeout(1500)
            ed = p.evaluate("() => { const e = [...document.querySelectorAll('[contenteditable=true]')].find(x => x.offsetParent); const r = e.getBoundingClientRect(); return [r.x + 40, r.y + 20]; }")
            p.mouse.click(*ed); p.keyboard.press("Control+End"); p.keyboard.press("Enter")
            en_vivo(p, q, red)
            # Sin red y de regreso
            ctx.set_offline(True); p.wait_for_timeout(800)
            marca = f"SINRED{int(time.time()) % 100000}"
            p.keyboard.type(marca); p.wait_for_timeout(800)
            q.evaluate("""(m) => { window.__visto = null; const ob = new MutationObserver(() => { if (!window.__visto && document.body.innerText.includes(m)) { window.__visto = Date.now(); ob.disconnect(); } });
              ob.observe(document.body, { subtree: true, childList: true, characterData: true }); }""", marca)
            t0 = q.evaluate("() => Date.now()"); ctx.set_offline(False)
            if REDES[red]: p.context.new_cdp_session(p).send("Network.emulateNetworkConditions", REDES[red])
            v = None
            for _ in range(600):
                v = q.evaluate("() => window.__visto")
                if v: break
                q.wait_for_timeout(100)
            fila(red, "volver de sin red → lo escrito llega a la otra persona", (v - t0) if v else None, None, "" if v else "no llegó en 60 s")
            largas = p.evaluate("() => window.__tiempos.largas") + q.evaluate("() => window.__tiempos.largas")
            graves = [l for l in largas if l[1] > 100]
            fila(red, "tarea más larga (congela la página)", max([l[1] for l in largas] or [0]), 100, f"{len(largas)} tareas > 50 ms, {len(graves)} > 100 ms: {graves[:8]}")
            if p.errores or q.errores: print("  errores de JavaScript:", (p.errores + q.errores)[:6])
            ctx2.close(); ctx.close()
    finally:
        b.close(); pw.stop()
    with open(os.path.join(a.salida, "tiempos-" + "-".join(a.solo) + ".md"), "w") as f:
        f.write("| Red | Qué | ms | Meta | |\n|---|---|---|---|---|\n")
        for red, que, ms, meta, ok, nota in FILAS:
            f.write(f"| {red} | {que}{(' — ' + nota) if nota else ''} | {'—' if ms is None else round(ms)} | {('< ' + str(meta)) if meta else ''} | {ok} |\n")
    print("tabla en", os.path.join(a.salida, "tiempos-" + "-".join(a.solo) + ".md"))


if __name__ == "__main__":
    principal()

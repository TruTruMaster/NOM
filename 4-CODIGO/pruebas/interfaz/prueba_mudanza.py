"""Cambio de dirección en el navegador (js/mudanza.js): de la dirección de antes a la nueva, sin perder lo guardado.

    ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_mudanza.py [ruta/al/nom.jar]   (NOM_NAVEGADOR=firefox)

Arranca el servidor en un puerto libre y usa dos orígenes del mismo equipo como «dirección de antes» (localhost) y
«nueva» (127.0.0.1), con window.NOM_MUDANZA_PRUEBA. Revisa:
  1. sin nada guardado, la página de antes cambia de dirección al momento (conserva la ruta);
  2. con algo guardado, pasa por mudanza.html, y con un clic llega todo a la nueva (localStorage e IndexedDB);
  3. lo que la nueva ya tenía no se pisa;
  4. solo mudanza.html relaja Cross-Origin-Opener-Policy; lo demás sigue en same-origin.
"""
import os, shutil, socket, subprocess, sys, tempfile, time, urllib.request
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", os.path.expanduser("~/.cache/nom-pruebas/navegadores"))
from playwright.sync_api import sync_playwright

AQUI = os.path.dirname(os.path.abspath(__file__))
JAR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(AQUI, "..", "..", "app", "target", "nom.jar")
with socket.socket() as s:
    s.bind(("127.0.0.1", 0)); P = s.getsockname()[1]
VIEJA, NUEVA = f"http://localhost:{P}", f"http://127.0.0.1:{P}"
D = tempfile.mkdtemp(prefix="nom-mudanza-")
srv = subprocess.Popen([shutil.which("java") or "java", "-jar", JAR],
                       env=dict(os.environ, NOM_PUERTO=str(P), NOM_DATOS=os.path.join(D, "datos"), NOM_DESCUBRIMIENTO="false"),
                       stdout=open(os.path.join(D, "servidor.log"), "w"), stderr=subprocess.STDOUT)
fallas, total = [], 0
def revisar(nombre, ok, detalle=""):
    global total
    total += 1
    print(("OK  " if ok else "MAL ") + nombre + (f"  ({detalle})" if detalle and not ok else ""))
    if not ok: fallas.append(nombre)

try:
    for _ in range(120):
        try: urllib.request.urlopen(NUEVA + "/api/red/info", timeout=2); break
        except Exception: time.sleep(1)
    def coop(ruta):
        return urllib.request.urlopen(NUEVA + ruta, timeout=10).headers.get("Cross-Origin-Opener-Policy")
    revisar("4. mudanza.html relaja COOP", coop("/mudanza.html") == "unsafe-none", coop("/mudanza.html"))
    revisar("4. la portada sigue en same-origin", coop("/") == "same-origin", coop("/"))
    revisar("4. mudanza.html no va a buscadores",
            urllib.request.urlopen(NUEVA + "/mudanza.html", timeout=10).headers.get("X-Robots-Tag") == "noindex")

    with sync_playwright() as p:
        nav = getattr(p, os.environ.get("NOM_NAVEGADOR", "chromium")).launch()
        prueba = f"window.NOM_MUDANZA_PRUEBA = {{ vieja: '{VIEJA}', nueva: '{NUEVA}' }};"

        # 1. Sin nada guardado: cambia de dirección al momento, con la misma ruta.
        ctx = nav.new_context(); ctx.add_init_script(prueba)
        pag = ctx.new_page()
        pag.goto(VIEJA + "/?x=1#hola")
        pag.wait_for_url(NUEVA + "/**", timeout=15000)
        revisar("1. sin nada guardado va a la nueva", pag.url.startswith(NUEVA + "/"), pag.url)
        revisar("1. conserva la búsqueda y el ancla", pag.url.endswith("/?x=1#hola"), pag.url)
        ctx.close()

        # 2 y 3. Con algo guardado en la de antes y algo ya en la nueva.
        ctx = nav.new_context(); ctx.add_init_script(prueba)
        pag = ctx.new_page()
        pag.goto(NUEVA + "/mudanza.html")   # sin ventana madre: solo avisa
        pag.evaluate("localStorage.setItem('nom.compartida', 'de-la-nueva')")
        pag.goto(VIEJA + "/mudanza.html")   # en la de antes, sin que corra la app: preparar lo guardado
        pag.evaluate("""() => new Promise((ok, mal) => {
            localStorage.setItem('nom.compartida', 'de-la-vieja');
            localStorage.setItem('nom.solo-vieja', 'si');
            const r = indexedDB.open('nom-boveda', 1);
            r.onupgradeneeded = () => r.result.createObjectStore('cosas');
            r.onsuccess = () => {
              const tx = r.result.transaction('cosas', 'readwrite'), s = tx.objectStore('cosas');
              s.put({ texto: 'documento sin cuenta', bytes: new Uint8Array([1, 2, 3]), cuando: new Date(5) }, 'doc1');
              s.put(new Blob(['archivo'], { type: 'text/plain' }), 'blob1');
              tx.oncomplete = () => { r.result.close(); ok(); }; tx.onerror = () => mal(tx.error);
            };
            r.onerror = () => mal(r.error);
        })""")
        pag.goto(VIEJA + "/?abrir=1")
        pag.wait_for_url(VIEJA + "/mudanza.html**", timeout=15000)
        revisar("2. con algo guardado pasa por mudanza.html", "/mudanza.html" in pag.url, pag.url)
        with ctx.expect_page(timeout=15000) as emergente:
            pag.click("#pasar")
        emergente.value.wait_for_load_state()
        pag.wait_for_url(NUEVA + "/**", timeout=30000)
        revisar("2. al terminar sigue en la nueva, en la misma ruta", pag.url == NUEVA + "/?abrir=1", pag.url)
        pag.goto(NUEVA + "/mudanza.html")
        ls = pag.evaluate("[localStorage.getItem('nom.solo-vieja'), localStorage.getItem('nom.compartida')]")
        revisar("2. pasó lo de localStorage", ls[0] == "si", ls)
        revisar("3. no pisó lo que la nueva ya tenía", ls[1] == "de-la-nueva", ls)
        bd = pag.evaluate("""() => new Promise((ok) => {
            const r = indexedDB.open('nom-boveda');
            r.onsuccess = () => {
              const db = r.result;
              if (!db.objectStoreNames.contains('cosas')) return ok(null);
              const s = db.transaction('cosas').objectStore('cosas'), d = s.get('doc1'), b = s.get('blob1');
              b.onsuccess = async () => ok({ texto: d.result && d.result.texto, bytes: d.result && Array.from(d.result.bytes),
                cuando: d.result && d.result.cuando.getTime(), blob: b.result ? await b.result.text() : null });
            };
        })""")
        revisar("2. pasó el documento de IndexedDB", bool(bd) and bd["texto"] == "documento sin cuenta" and bd["bytes"] == [1, 2, 3] and bd["cuando"] == 5, bd)
        revisar("2. pasó el archivo (Blob)", bool(bd) and bd["blob"] == "archivo", bd)
        revisar("2. la ventana emergente se cerró", emergente.value.is_closed())
        ctx.close()
        nav.close()
finally:
    srv.terminate()
    try: srv.wait(20)
    except Exception: srv.kill()
    shutil.rmtree(D, ignore_errors=True)

print(f"\n{total - len(fallas)} de {total}")
sys.exit(1 if fallas else 0)

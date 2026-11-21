"""La app con su propio origen (0.19) y los equipos que comparten: el suyo, uno nuevo y uno de antes.

  python pruebas/reales/origen_equipos.py [SERIE_ADB] [SERVIDOR_NUEVO] [SERVIDOR_DE_ANTES]
  (por omisión emulator-5554, http://10.0.2.2:8096 y http://10.0.2.2:8097, vistos desde el teléfono)

- Con un servidor nuevo, la app pide y lee la respuesta (CORS con el origen nuevo).
- Con uno de antes de la 0.19 (no conoce el origen nuevo: el navegador no deja leer), la app no dice «sin conexión»:
  dice que ese equipo tiene una versión anterior.
- El teléfono como servidor: la app enciende el suyo (http://127.0.0.1:8080) y le habla desde su origen.
La app nueva ya instalada, con la depuración del WebView (APK de depuración) y node (cdp.mjs).
"""
import os, json, subprocess, sys, time
from pathlib import Path

SER = sys.argv[1] if len(sys.argv) > 1 else "emulator-5554"
NUEVO = sys.argv[2] if len(sys.argv) > 2 else "http://10.0.2.2:8096"
VIEJO = sys.argv[3] if len(sys.argv) > 3 else "http://10.0.2.2:8097"
AQUI = Path(__file__).resolve().parent
CDP = "http://localhost:9223"
# la app que se prueba (otra con distinto nombre convive con la de siempre: NOM_PAQUETE=mx.wack.nom.prueba)
PAQUETE = os.environ.get("NOM_PAQUETE", "mx.wack.nom")
res = []


def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det else ""), flush=True)


def adb(*a, t=60):
    return subprocess.run(["adb", "-s", SER, *a], capture_output=True, text=True, timeout=t).stdout.strip()


def enlazar():
    for _ in range(40):
        pid = adb("shell", "pidof", PAQUETE)
        if pid:
            adb("forward", "tcp:9223", f"localabstract:webview_devtools_remote_{pid.split()[0]}"); return True
        time.sleep(0.5)
    return False


def app(expr, ms=60000):
    r = subprocess.run(["node", str(AQUI / "cdp.mjs"), CDP, expr, str(ms)], capture_output=True, text=True, timeout=ms / 1000 + 15)
    try: return json.loads(r.stdout.strip() or "null")
    except Exception: return {"error": (r.stdout + r.stderr)[:300]}


PEDIR = """(async (base) => { const antes = NOM.api.servidor(); NOM.api.fijarServidor(base);
  try { const r = await NOM.api.get('/api/red/info'); return { ok: true, version: r && r.version, servidor: NOM.api.servidor() }; }
  catch (e) { return { ok: false, estado: e.estado, mensaje: e.message }; }
  finally { NOM.api.fijarServidor(antes); } })"""

adb("shell", "cmd", "connectivity", "airplane-mode", "disable")
adb("shell", "am", "force-stop", PAQUETE); time.sleep(1)
adb("shell", "monkey", "-p", PAQUETE, "-c", "android.intent.category.LAUNCHER", "1"); time.sleep(4)
enlazar()
t0 = time.time()
while time.time() - t0 < 30 and app("!!(window.NOM && NOM.api)") is not True: time.sleep(0.5)
anota("la app abre desde su propio origen", app("location.origin") == "https://appassets.androidplatform.net", str(app("location.origin")))
r = app(PEDIR + f"({json.dumps(NUEVO)})")
anota("con un equipo nuevo, la app lee sus respuestas", isinstance(r, dict) and r.get("ok"), str(r)[:160])
r = app(PEDIR + f"({json.dumps(VIEJO)})")
anota("con un equipo de antes de la 0.19, la app dice que es una versión anterior (no «sin conexión»)",
      isinstance(r, dict) and not r.get("ok") and "versión anterior" in str(r.get("mensaje")), str(r)[:200])
r = app(PEDIR + "('http://10.0.2.2:1')")
anota("con un equipo que no contesta, sigue diciendo «sin conexión»", isinstance(r, dict) and not r.get("ok") and "Sin conexión" in str(r.get("mensaje")), str(r)[:200])
# El teléfono como servidor: la app enciende el suyo y le habla (sin cambiarse a él).
e = app("(() => { try { return NOM.anfitrion.encender(false); } catch (x) { return { error: String(x) }; } })()")
anota("la app enciende el servidor del teléfono", isinstance(e, dict) and e.get("activo"), str(e)[:200])
local = (e or {}).get("local") if isinstance(e, dict) else None
r = app(PEDIR + f"({json.dumps(local)})") if local else None
anota("…y le habla desde su propio origen (CORS del teléfono)", isinstance(r, dict) and r.get("ok"), f"{local} · {str(r)[:160]}")
v = app(f"""(async () => {{ const r = await fetch({json.dumps((local or '') + '/api/red/info')}, {{ headers: {{ Origin: 'x' }} }}); return r.status; }})()""")
anota("…con fetch directo también", v in (200, 401), str(v))
app("(NOM.anfitrion.detener(), 1)")
# La ayuda: lo que encuentra «Buscar» lleva el nombre del archivo (una dirección limpia rompería sus enlaces relativos)
app("(location.href = '/ayuda/buscar/index.html?q=espacio', 1)"); time.sleep(2)
t0 = time.time()
while time.time() - t0 < 20 and not app("document.querySelectorAll('.lista-articulos a').length > 0"): time.sleep(0.5)
ligas = app("[...document.querySelectorAll('.lista-articulos a')].map((a) => a.getAttribute('href'))") or []
anota(f"la ayuda busca dentro de la app ({len(ligas)} resultados) y cada uno lleva su index.html", bool(ligas) and all(str(x).endswith("index.html") for x in ligas), str(ligas[:3]))
app("(document.querySelector('.lista-articulos a').click(), 1)"); time.sleep(2)
v = app("({ ruta: location.pathname, estilos: [...document.styleSheets].length > 0 && [...document.styleSheets].every((h) => h.cssRules.length > 0) })")
anota("…y el artículo abre con sus estilos", isinstance(v, dict) and str(v.get("ruta")).endswith("/index.html") and v.get("estilos"), str(v))
app("(location.href = '/index.html', 1)"); time.sleep(2)
print(f"\nOrigen propio y equipos: {sum(1 for r in res if r[1])} de {len(res)} bien")
sys.exit(0 if all(r[1] for r in res) else 1)

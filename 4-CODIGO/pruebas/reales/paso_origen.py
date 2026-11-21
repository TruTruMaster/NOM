"""La app pasa a su propio origen (0.19) sin perder nada de lo guardado en el teléfono.

  python pruebas/reales/paso_origen.py APK_DE_ANTES APK_NUEVO [SERIE_ADB] [SERVIDOR_VISTO_DESDE_EL_TELEFONO]
  (por omisión emulator-5554 y http://10.0.2.2:8096; los dos APK de depuración, para manejar el WebView por CDP)

1. Instala desde cero la app de antes (interfaz en file://), entra por la pantalla de entrar, escribe en un documento
   con red y luego sin red (modo avión: el cambio queda pendiente en el teléfono) y toma una foto de lo guardado
   (localStorage y cada registro de IndexedDB, resumido con SHA-256).
2. Sin red todavía, instala encima la app nueva y la abre: pasa los datos a https://appassets.androidplatform.net.
3. Compara la foto: todo lo de antes está igual, la sesión sigue abierta y, al volver la red, lo escrito sin red llega
   al servidor. Al volver a abrir ya no pasa nada (se hace una sola vez) y la ayuda abre dentro de la app.
Necesita el servidor de pruebas con el FiltroCors que conoce el origen nuevo y node (cdp.mjs).
"""
import os, hashlib, json, subprocess, sys, time, urllib.request
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "interfaz"))
import nomprueba as N  # noqa: E402

VIEJA, NUEVA = sys.argv[1], sys.argv[2]
SER = sys.argv[3] if len(sys.argv) > 3 else "emulator-5554"
DESDE_TEL = sys.argv[4] if len(sys.argv) > 4 else "http://10.0.2.2:8096"
AQUI = Path(__file__).resolve().parent
CDP = "http://localhost:9223"
# la app que se prueba (otra con distinto nombre convive con la de siempre: NOM_PAQUETE=mx.wack.nom.prueba)
PAQUETE = os.environ.get("NOM_PAQUETE", "mx.wack.nom")
if PAQUETE == "mx.wack.nom" and not SER.startswith("emulator"):
    sys.exit("Esta prueba desinstala la app: en un teléfono de verdad, con otra copia (NOM_PAQUETE=mx.wack.nom.prueba).")
res = []


def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det else ""), flush=True)


def adb(*a, t=120):
    return subprocess.run(["adb", "-s", SER, *a], capture_output=True, text=True, timeout=t).stdout.strip()


def enlazar():
    for _ in range(40):
        pid = adb("shell", "pidof", PAQUETE)
        if pid:
            adb("forward", "tcp:9223", f"localabstract:webview_devtools_remote_{pid.split()[0]}")
            return True
        time.sleep(0.5)
    return False


def app(expr, ms=60000):
    r = subprocess.run(["node", str(AQUI / "cdp.mjs"), CDP, expr, str(ms)], capture_output=True, text=True, timeout=ms / 1000 + 15)
    try: return json.loads(r.stdout.strip() or "null")
    except Exception: return {"error": (r.stdout + r.stderr)[:300]}


def esperar_app(expr, s=30):
    t0 = time.time()
    while time.time() - t0 < s:
        v = app(expr, 15000)
        if v and not (isinstance(v, dict) and v.get("error")): return v, time.time() - t0
        time.sleep(0.5)
    return None, time.time() - t0


# Sin red de verdad: algunos teléfonos (el A55) dejan el Wi-Fi prendido en modo avión. Se apaga también, y al final
# queda como estaba.
WIFI_ANTES = adb("shell", "settings", "get", "global", "wifi_on").strip()


def sin_red():
    adb("shell", "cmd", "connectivity", "airplane-mode", "enable"); adb("shell", "svc", "wifi", "disable")


def con_red():
    adb("shell", "cmd", "connectivity", "airplane-mode", "disable")
    if WIFI_ANTES not in ("0", ""): adb("shell", "svc", "wifi", "enable")


import atexit  # noqa: E402
atexit.register(con_red)   # aunque la prueba se corte, el teléfono no se queda sin red


def abrir():
    adb("shell", "am", "force-stop", PAQUETE); time.sleep(1)
    adb("shell", "monkey", "-p", PAQUETE, "-c", "android.intent.category.LAUNCHER", "1")
    time.sleep(3)
    return enlazar()


# Foto de lo guardado: cada clave de localStorage con su valor y cada registro de IndexedDB con un resumen de su
# contenido (los bytes cuentan como bytes; el orden de las propiedades no importa).
FOTO = r"""(async () => {
  const enc = (v) => {
    if (v === null || typeof v !== 'object') return v;
    if (v instanceof ArrayBuffer) return { bytes: [...new Uint8Array(v)].join(',') };
    if (ArrayBuffer.isView(v)) return { [v.constructor.name]: [...new Uint8Array(v.buffer, v.byteOffset, v.byteLength)].join(',') };
    if (v instanceof Date) return { fecha: v.getTime() };
    if (Array.isArray(v)) return v.map(enc);
    const o = {}; for (const k of Object.keys(v).sort()) o[k] = enc(v[k]); return o;
  };
  const sha = async (t) => [...new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(t)))].map((b) => b.toString(16).padStart(2, '0')).join('');
  const pedido = (r) => new Promise((ok, mal) => { r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error); });
  const ls = {}; for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); ls[k] = localStorage.getItem(k); }
  const idb = {};
  for (const n of ['nom-boveda', 'nom-binarios']) {
    const db = await new Promise((ok) => { let nueva = false; const r = indexedDB.open(n); r.onupgradeneeded = () => { nueva = true; };
      r.onsuccess = () => { if (nueva) { r.result.close(); indexedDB.deleteDatabase(n); ok(null); } else ok(r.result); }; r.onerror = () => ok(null); });
    if (!db) continue;
    idb[n] = { version: db.version };
    for (const a of db.objectStoreNames) {
      const s = db.transaction(a, 'readonly').objectStore(a);
      const ks = await pedido(s.getAllKeys()), vs = await pedido(db.transaction(a, 'readonly').objectStore(a).getAll());
      const m = {}; for (let i = 0; i < ks.length; i++) m[JSON.stringify(enc(ks[i]))] = await sha(JSON.stringify(enc(vs[i])));
      idb[n][a] = m;
    }
    db.close();
  }
  return { origen: location.origin, protocolo: location.protocol, ls, idb };
})()"""

a = N.cuenta()
esp = N.espacio(a["token"], "Paso de origen")
dt = N.documento(a["token"], esp, "TEXTO", "Escrito antes de actualizar")

# --- 1. La app de antes, desde cero
adb("shell", "cmd", "connectivity", "airplane-mode", "disable")
adb("uninstall", PAQUETE)
r = subprocess.run(["adb", "-s", SER, "install", "-r", VIEJA], capture_output=True, text=True, timeout=300)
anota("se instala la app de antes", "Success" in r.stdout, (r.stdout + r.stderr)[-120:])
anota("la app de antes abre (y se puede manejar por CDP)", abrir())
v, _ = esperar_app("location.protocol", 30)
anota("la app de antes abre su interfaz desde sus archivos (file:)", v == "file:", str(v))
app(f"(NOM.api.fijarServidor({json.dumps(DESDE_TEL)}), location.hash = '#/', setTimeout(() => location.reload(), 200), 1)")
time.sleep(4); enlazar()
v, _ = esperar_app("!!document.querySelector('#u')", 30)
app(f"""(() => {{ const u = document.querySelector('#u'), c = document.querySelector('#c');
  const fija = (el, v) => {{ Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, v); el.dispatchEvent(new Event('input', {{ bubbles: true }})); }};
  fija(u, {json.dumps(a['nombre'])}); fija(c, {json.dumps(N.CL)}); document.querySelector('button[type=submit]').click(); return 1; }})()""")
v, t = esperar_app("!!(NOM.api.token() && NOM.api.usuario())", 30)
anota("la app de antes entra con usuario y contraseña", bool(v), f"{t:.1f} s")
token = app("NOM.api.token()")
app(f"(location.hash = '#/doc/{dt}', 1)")
v, _ = esperar_app("!!(NOM.editorActual && NOM.editorActual.vista && NOM.editorActual.vista.isConnected && document.querySelector('.wd-cuerpo'))", 30)
anota("abre el documento", bool(v))
ESCRIBIR = """(async (t) => { const c = document.querySelector('.wd-cuerpo'); c.focus(); const s = getSelection(); const r = document.createRange();
  r.selectNodeContents(c); r.collapse(false); s.removeAllRanges(); s.addRange(r); document.execCommand('insertText', false, t);
  await new Promise((ok) => setTimeout(ok, 1500)); return c.innerText; })"""


def texto_servidor():
    st, d = N.api("GET", f"/api/documentos/{dt}", None, a["token"])
    return json.dumps(d, ensure_ascii=False) if st == 200 else ""


app(ESCRIBIR + "('Con red antes de actualizar. ')")
t0 = time.time()
while time.time() - t0 < 20 and "Con red antes de actualizar" not in texto_servidor(): time.sleep(0.5)
anota("lo escrito con red llega al servidor", "Con red antes de actualizar" in texto_servidor())
sin_red(); time.sleep(4)
app(ESCRIBIR + "('Escrito sin red antes de actualizar. ')")
# un archivo de 1.5 MB guardado sin red: queda en nom-binarios, pendiente de subir (y cruza el puente en partes)
ref = app(f"""(async () => {{ const u = new Uint8Array(1500000); for (let i = 0; i < u.length; i += 65536) crypto.getRandomValues(u.subarray(i, Math.min(u.length, i + 65536)));
  const d = [...new Uint8Array(await crypto.subtle.digest('SHA-256', u))].map((b) => b.toString(16).padStart(2, '0')).join('');
  return [await NOM.binarios.guardar(u, 'image/png', {{ espacio: {json.dumps(esp)} }}), d]; }})()""", 60000)
anota("sin red: un archivo de 1.5 MB queda guardado en el teléfono", isinstance(ref, list) and len(ref) == 2, str(ref)[:120])
huella = ref[0].split("/")[-1].split(":")[-1].split(".")[0] if isinstance(ref, list) else ""
time.sleep(3)   # la copia del documento y la cola de cambios se guardan en la bóveda (en segundo plano)
antes = app(FOTO, 120000)
anota("foto de lo guardado en la app de antes", isinstance(antes, dict) and antes.get("protocolo") == "file:",
      f"{len((antes or {}).get('ls', {}))} claves, " + ", ".join(f"{n}: {sum(len(x) for k, x in d.items() if k != 'version')} registros" for n, d in (antes or {}).get("idb", {}).items()))
anota("…y lo escrito sin red todavía no está en el servidor", "Escrito sin red antes de actualizar" not in texto_servidor())

# --- 2. La app nueva encima, sin red
r = subprocess.run(["adb", "-s", SER, "install", "-r", NUEVA], capture_output=True, text=True, timeout=300)
anota("se instala la app nueva encima de la de antes", "Success" in r.stdout, (r.stdout + r.stderr)[-120:])
t0 = time.time(); abrir()
v, t = esperar_app("location.origin === 'https://appassets.androidplatform.net' && location.pathname === '/index.html' && !!window.NOM && !!NOM.api", 90)
anota("la app nueva abre su interfaz desde su propio origen", bool(v), f"{time.time() - t0:.1f} s desde que abre")
time.sleep(2)
despues = app(FOTO, 120000)
faltan = [k for k, x in antes["ls"].items() if despues["ls"].get(k) != x]
anota(f"las {len(antes['ls'])} claves de localStorage están igual", not faltan, str(faltan[:5]))
for n, d in antes["idb"].items():
    for al, regs in d.items():
        if al == "version": continue
        nuevos = (despues["idb"].get(n) or {}).get(al) or {}
        mal = [k for k, h in regs.items() if nuevos.get(k) != h]
        anota(f"IndexedDB {n}/{al}: los {len(regs)} registros están igual", not mal, str(mal[:3]))
anota("la sesión sigue abierta (el mismo token)", app("NOM.api.token()") == token)
anota("no pide entrar otra vez", not app("!!document.querySelector('#u')"))
anota("Atrás no regresa a la página de paso (la historia se limpió)", app("history.length") == 1, str(app("history.length")))

# --- 3. Vuelve la red: lo escrito sin red llega
con_red()
t0 = time.time()
while time.time() - t0 < 90 and "Escrito sin red antes de actualizar" not in texto_servidor(): time.sleep(1)
anota("al volver la red, lo escrito sin red con la app de antes llega al servidor", "Escrito sin red antes de actualizar" in texto_servidor(), f"{time.time() - t0:.1f} s")


def binario():
    q = urllib.request.Request(N.B + f"/api/espacios/{esp}/binarios/{huella}", headers={"Authorization": "Bearer " + a["token"]})
    try:
        with urllib.request.urlopen(q, timeout=10) as r: return hashlib.sha256(r.read()).hexdigest()
    except Exception: return None


t0 = time.time()
while time.time() - t0 < 60 and binario() is None: time.sleep(1)
anota("…y el archivo guardado sin red sube con los mismos bytes", isinstance(ref, list) and binario() == ref[1], f"{huella[:12]}… en {time.time() - t0:.1f} s")
app(f"(location.hash = '#/doc/{dt}', 1)")
v, _ = esperar_app("!!(NOM.editorActual && document.querySelector('.wd-cuerpo') && document.querySelector('.wd-cuerpo').innerText.includes('Escrito sin red antes de actualizar'))", 30)
anota("la app nueva abre el documento con todo lo escrito", bool(v))
app(ESCRIBIR + "('Escrito con la app nueva. ')")
t0 = time.time()
while time.time() - t0 < 20 and "Escrito con la app nueva" not in texto_servidor(): time.sleep(0.5)
anota("lo que escribe la app nueva llega al servidor (CORS con el origen nuevo)", "Escrito con la app nueva" in texto_servidor())

# --- 4. Una sola vez; la ayuda dentro de la app
abrir()
v, t = esperar_app("location.origin === 'https://appassets.androidplatform.net' && !!window.NOM && !!NOM.api", 30)
anota("al abrirla otra vez va directo a la interfaz (sin paso)", bool(v) and app("location.pathname") == "/index.html", f"{t:.1f} s")
anota("…y con la sesión abierta", app("NOM.api.token()") == token)
app("(NOM.ayuda('empezar'), 1)")
v, _ = esperar_app("location.pathname === '/ayuda/empezar/index.html' && document.readyState === 'complete' && document.querySelectorAll('a').length > 3", 20)
anota("la ayuda abre dentro de la app", bool(v), str(app("location.href")))
v, _ = esperar_app("(async () => { const r = await fetch('/fonts/nunito-latin-wght-normal.woff2'); return r.ok && r.headers.get('content-type') === 'font/woff2'; })()", 10)
anota("la app sirve sus archivos con su tipo (fetch en el mismo origen)", bool(v))
v = app("(async () => (await fetch('/no-existe.js')).status)()")
anota("lo que no existe da 404 (no sale a la red)", v == 404, str(v))
app("(location.href = '/index.html', 1)"); time.sleep(2); enlazar()

# Se borra la cuenta de prueba (con su espacio y su documento).
st = N.api("POST", "/api/cuentas/eliminar", {"contrasena": N.CL, "confirmacion": "BORRAR"}, a["token"])[0]
anota("se borra la cuenta de prueba", st == 200, str(st))
print(f"\nPaso al origen propio: {sum(1 for r in res if r[1])} de {len(res)} bien")
sys.exit(0 if all(r[1] for r in res) else 1)

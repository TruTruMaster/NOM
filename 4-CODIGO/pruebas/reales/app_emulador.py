"""La app Android (en el emulador o en un teléfono por adb) y una computadora en el mismo documento.

  python pruebas/reales/app_emulador.py [SERIE_ADB] [SERVIDOR_VISTO_DESDE_EL_TELEFONO]
  (por omisión emulator-5554 y http://10.0.2.2:8099, que es la laptop vista desde el emulador)

La app entra con su cuenta por la pantalla de entrar, abre un documento de texto y una hoja; la
computadora (Chromium) entra con otra cuenta del mismo espacio. Las dos escriben y se ven; luego la app se
queda sin red (modo avión), escribe, vuelve la red y todo se junta sin rama (menos de 2 minutos).
Necesita el servidor de pruebas en el 8099 y node (para manejar el WebView por CDP: cdp.mjs).
"""
import json, os, subprocess, sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "interfaz"))
import nomprueba as N  # noqa: E402
SER = sys.argv[1] if len(sys.argv) > 1 else "emulator-5554"
DESDE_TEL = sys.argv[2] if len(sys.argv) > 2 else "http://10.0.2.2:8099"
AQUI = Path(__file__).resolve().parent
CDP = "http://localhost:9223"
res = []
# Por depuración INALÁMBRICA (la serie es una dirección: «ip:puerto» o «[ipv6]:puerto») el modo avión corta la propia
# conexión de adb y el teléfono se queda sin red y sin poder quitarle el modo avión (pasó el 18-nov con el A55): ahí no
# se corta la red; lo de «sin red» se prueba en el emulador o con cable.
INALAMBRICO = SER.startswith("[") or (":" in SER and not SER.startswith("emulator"))


def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det else ""), flush=True)


def adb(*a, t=60):
    return subprocess.run(["adb", "-s", SER, *a], capture_output=True, text=True, timeout=t).stdout.strip()


def enlazar():
    pid = adb("shell", "pidof", "mx.wack.nom")
    adb("forward", "tcp:9223", f"localabstract:webview_devtools_remote_{pid}")


def app(expr, ms=60000):
    r = subprocess.run(["node", str(AQUI / "cdp.mjs"), CDP, expr, str(ms)], capture_output=True, text=True, timeout=ms / 1000 + 15)
    try: return json.loads(r.stdout.strip() or "null")
    except Exception: return {"error": (r.stdout + r.stderr)[:300]}


def esperar_app(expr, s=30):
    t0 = time.time()
    while time.time() - t0 < s:
        v = app(expr)
        if v and not (isinstance(v, dict) and v.get("error")): return True, time.time() - t0
        time.sleep(0.5)
    return False, time.time() - t0


a, b = N.cuenta(), N.cuenta()
esp = N.espacio(a["token"], "App y computadora")
N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": b["nombre"], "permiso": "EDICION"}, a["token"])
dt = N.documento(a["token"], esp, "TEXTO", "Acta en la app")
dh = N.documento(a["token"], esp, "HOJA", "Hoja en la app")

# --- La app: servidor de la laptop y entrar con usuario y contraseña por la pantalla de entrar
if not INALAMBRICO: adb("shell", "cmd", "connectivity", "airplane-mode", "disable")
adb("shell", "am", "force-stop", "mx.wack.nom"); time.sleep(1)
adb("shell", "monkey", "-p", "mx.wack.nom", "-c", "android.intent.category.LAUNCHER", "1"); time.sleep(6)
enlazar()
app(f"(NOM.anfitrion && NOM.anfitrion.detener(), NOM.api.limpiarSesion(), NOM.api.fijarServidor({json.dumps(DESDE_TEL)}), NOM.guardar('nom.infoServidor', null), location.hash = '#/', setTimeout(() => location.reload(), 200), 1)")
time.sleep(6); enlazar()
ok, _ = esperar_app("!!document.querySelector('#u')", 30)
anota("la app enseña la pantalla de entrar del servidor de la computadora", ok, str(app("NOM.api.servidor()")))
app(f"""(() => {{ const u = document.querySelector('#u'), c = document.querySelector('#c');
  const fija = (el, v) => {{ Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, v); el.dispatchEvent(new Event('input', {{ bubbles: true }})); }};
  fija(u, {json.dumps(a['nombre'])}); fija(c, {json.dumps(N.CL)}); document.querySelector('button[type=submit]').click(); return 1; }})()""")
ok, t = esperar_app("!!(NOM.api.token() && NOM.api.usuario())", 30)
anota("la app entra con usuario y contraseña", ok, f"{t:.1f} s")

# --- Texto: los dos escriben
app(f"(location.hash = '#/doc/{dt}', 1)")
ok, _ = esperar_app("!!(NOM.editorActual && NOM.editorActual.vista && NOM.editorActual.vista.isConnected && document.querySelector('.wd-cuerpo'))", 30)
anota("la app abre el documento de texto", ok)
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    nav = pw.chromium.launch(); ctx = nav.new_context(viewport={"width": 1300, "height": 860}); p = ctx.new_page(); err = []
    p.on("pageerror", lambda e: err.append(str(e)[:200]))
    p.goto(N.B + "/"); p.wait_for_selector("#u"); p.fill("#u", b["nombre"]); p.fill("#c", N.CL); p.click("button[type=submit]")
    p.wait_for_function("() => !!localStorage.getItem('nom.token')"); p.wait_for_timeout(800)
    p.goto(N.B + "/#/doc/" + dt); p.wait_for_selector(".wd-cuerpo"); p.wait_for_timeout(1500)
    ESCRIBIR_APP = """(async (t) => { const c = document.querySelector('.wd-cuerpo'); c.focus(); const s = getSelection(); const r = document.createRange();
      r.selectNodeContents(c); r.collapse(false); s.removeAllRanges(); s.addRange(r); document.execCommand('insertText', false, t);
      await new Promise((ok) => setTimeout(ok, 1200)); return c.innerText; })"""
    app(ESCRIBIR_APP + "('Desde la app. ')")
    t0 = time.time()
    try: p.wait_for_function("() => document.querySelector('.wd-cuerpo').innerText.includes('Desde la app')", timeout=20000); anota("lo de la app llega a la computadora", True, f"{time.time() - t0:.1f} s")
    except Exception: anota("lo de la app llega a la computadora", False)
    p.click(".wd-cuerpo"); p.keyboard.press("Control+End"); p.keyboard.type("Desde la computadora. ")
    ok, t = esperar_app("document.querySelector('.wd-cuerpo').innerText.includes('Desde la computadora')", 20)
    anota("lo de la computadora llega a la app", ok, f"{t:.1f} s")
    # --- Sin red en la app (modo avión) un momento (no por depuración inalámbrica: ver INALAMBRICO)
    if INALAMBRICO: print("(omitido: sin red en un teléfono conectado por Wi-Fi; se prueba en el emulador)", flush=True)
    else: adb("shell", "cmd", "connectivity", "airplane-mode", "enable"); time.sleep(4)
    app(ESCRIBIR_APP + "('Escrito en la app sin red. ')")
    p.click(".wd-cuerpo"); p.keyboard.press("Control+End"); p.keyboard.type("La computadora sigue escribiendo. ")
    time.sleep(3)
    estado = app("(document.querySelector('.estado-sync, [data-sync], .ed-estado') || {}).innerText || ''")
    if not INALAMBRICO: adb("shell", "cmd", "connectivity", "airplane-mode", "disable")
    ok, t = esperar_app("document.querySelector('.wd-cuerpo').innerText.includes('La computadora sigue escribiendo')", 60)
    anota("al volver la red, a la app le llega lo que escribió la computadora", ok, f"{t:.1f} s · estado sin red: {str(estado)[:60]!r}")
    try: p.wait_for_function("() => document.querySelector('.wd-cuerpo').innerText.includes('Escrito en la app sin red')", timeout=60000); anota("y a la computadora le llega lo que la app escribió sin red", True)
    except Exception: anota("y a la computadora le llega lo que la app escribió sin red", False)
    time.sleep(2)
    ramas = N.api("GET", f"/api/documentos/{dt}/ramas", None, a["token"])
    anota("un corte corto no abre rama (se junta solo)", ramas[0] == 200 and not (ramas[1] or {}).get("ramas", ramas[1] if isinstance(ramas[1], list) else []), str(ramas)[:150])
    t_app = app("document.querySelector('.wd-cuerpo').innerText"); t_pc = p.evaluate("() => document.querySelector('.wd-cuerpo').innerText")
    anota("la app y la computadora quedan con el mismo texto", str(t_app).strip() == t_pc.strip(), repr(t_pc[:140]))
    # --- Hoja
    app(f"(location.hash = '#/doc/{dh}', 1)")
    ok, _ = esperar_app("!!document.querySelector('td[data-c=\"0\"][data-f=\"0\"]')", 30)
    anota("la app abre la hoja", ok)
    app("""(async () => { const ed = NOM.editorActual; const b = ed.doc.deLaPagina(ed.pagina)[0];
      ed.doc.cambiar(b.id, { 'c:A1': '12', 'c:A2': '30', 'c:A3': '=SUMA(A1:A2)' }); await new Promise((ok) => setTimeout(ok, 1500)); return 1; })()""")
    p.goto(N.B + "/#/doc/" + dh); p.wait_for_selector('td[data-c="0"][data-f="2"]'); p.wait_for_timeout(1500)
    v = p.evaluate("() => document.querySelector('td[data-c=\"0\"][data-f=\"2\"]').textContent")
    anota("la hoja de la app llega a la computadora y la fórmula calcula", v.strip() == "42", v)
    celda = p.evaluate("() => { const c = document.querySelector('td[data-c=\"1\"][data-f=\"0\"]'); const r = c.getBoundingClientRect(); return [r.x + r.width / 2, r.y + r.height / 2]; }")
    p.mouse.click(*celda); p.keyboard.type("Hola desde la PC"); p.keyboard.press("Enter")
    ok, t = esperar_app("(document.querySelector('td[data-c=\"1\"][data-f=\"0\"]') || {}).textContent === 'Hola desde la PC'", 20)
    anota("lo que escribe la computadora en la hoja se ve en la app", ok, f"{t:.1f} s")
    anota("sin errores de JavaScript en la computadora", not err, str(err[:2]))
    nav.close()
if not INALAMBRICO: adb("shell", "cmd", "connectivity", "airplane-mode", "disable")
# Las dos cuentas de prueba se borran (con sus espacios y documentos): así no queda basura en la nube.
app("(NOM.api.limpiarSesion(), location.hash = '#/', 1)")
for c in (a, b):
    st = N.api("POST", "/api/cuentas/eliminar", {"contrasena": N.CL, "confirmacion": "BORRAR"}, c["token"])[0]
    anota(f"se borra la cuenta de prueba {c['nombre']}", st == 200, str(st))
print(f"\nApp y computadora: {sum(1 for r in res if r[1])} de {len(res)} bien")
sys.exit(0 if all(r[1] for r in res) else 1)

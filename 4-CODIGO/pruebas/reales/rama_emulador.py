"""B-6 con la app Android: más de 2 minutos sin red mientras otra persona cambia el mismo documento.

  python pruebas/reales/rama_emulador.py [SERIE_ADB] [SERVIDOR_VISTO_DESDE_EL_TELEFONO] [SEGUNDOS]

Esperado (decisión del equipo): lo escrito sin red NO se junta solo con el original; se abre una rama
(documento nuevo ligado al original) con lo de la app, y el original queda solo con lo de la computadora.
"""
import json, subprocess, sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "interfaz"))
import nomprueba as N  # noqa: E402
SER = sys.argv[1] if len(sys.argv) > 1 else "emulator-5554"
DESDE_TEL = sys.argv[2] if len(sys.argv) > 2 else "http://10.0.2.2:8099"
SEG = int(sys.argv[3]) if len(sys.argv) > 3 else 135
AQUI = Path(__file__).resolve().parent
# Esta prueba corta la red del teléfono (modo avión) más de 2 minutos: por depuración INALÁMBRICA eso corta también la
# conexión de adb y el teléfono se queda en modo avión (pasó el 18-nov con el A55). Solo emulador o cable.
if SER.startswith("[") or (":" in SER and not SER.startswith("emulator")):
    sys.exit("rama_emulador.py corta la red del teléfono: córrela en el emulador o con el teléfono por cable (no por Wi-Fi)")
res = []
def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det else ""), flush=True)
def adb(*a, t=60): return subprocess.run(["adb", "-s", SER, *a], capture_output=True, text=True, timeout=t).stdout.strip()
def enlazar(): adb("forward", "tcp:9223", f"localabstract:webview_devtools_remote_{adb('shell', 'pidof', 'mx.wack.nom')}")
def app(expr, ms=60000):
    r = subprocess.run(["node", str(AQUI / "cdp.mjs"), "http://localhost:9223", expr, str(ms)], capture_output=True, text=True, timeout=ms / 1000 + 15)
    try: return json.loads(r.stdout.strip() or "null")
    except Exception: return {"error": (r.stdout + r.stderr)[:300]}
def esperar_app(expr, s=30):
    t0 = time.time()
    while time.time() - t0 < s:
        v = app(expr)
        if v and not (isinstance(v, dict) and v.get("error")): return True
        time.sleep(0.5)
    return False
def texto_doc(doc, tok):
    st, d = N.api("GET", f"/api/documentos/{doc}", None, tok)
    return json.dumps(d, ensure_ascii=False) if st == 200 else ""

a, b = N.cuenta(), N.cuenta()
esp = N.espacio(a["token"], "Rama en la app")
N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": b["nombre"], "permiso": "EDICION"}, a["token"])
dt = N.documento(a["token"], esp, "TEXTO", "Acta con rama")
adb("shell", "cmd", "connectivity", "airplane-mode", "disable")
adb("shell", "am", "force-stop", "mx.wack.nom"); time.sleep(1)
adb("shell", "monkey", "-p", "mx.wack.nom", "-c", "android.intent.category.LAUNCHER", "1"); time.sleep(6); enlazar()
app(f"(NOM.api.limpiarSesion(), NOM.api.fijarServidor({json.dumps(DESDE_TEL)}), location.hash = '#/', setTimeout(() => location.reload(), 200), 1)")
time.sleep(6); enlazar(); esperar_app("!!document.querySelector('#u')", 30)
app(f"""(() => {{ const u = document.querySelector('#u'), c = document.querySelector('#c');
  const fija = (el, v) => {{ Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, v); el.dispatchEvent(new Event('input', {{ bubbles: true }})); }};
  fija(u, {json.dumps(a['nombre'])}); fija(c, {json.dumps(N.CL)}); document.querySelector('button[type=submit]').click(); return 1; }})()""")
anota("la app entra", esperar_app("!!NOM.api.token()", 30))
app(f"(location.hash = '#/doc/{dt}', 1)")
anota("la app abre el documento", esperar_app("!!document.querySelector('.wd-cuerpo')", 30))
ESCRIBIR = """(async (t) => { const c = document.querySelector('.wd-cuerpo'); c.focus(); const s = getSelection(); const r = document.createRange();
  r.selectNodeContents(c); r.collapse(false); s.removeAllRanges(); s.addRange(r); document.execCommand('insertText', false, t);
  await new Promise((ok) => setTimeout(ok, 1000)); return c.innerText; })"""
app(ESCRIBIR + "('Antes del corte. ')"); time.sleep(2)
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    nav = pw.chromium.launch(); p = nav.new_page(viewport={"width": 1300, "height": 860})
    p.goto(N.B + "/"); p.wait_for_selector("#u"); p.fill("#u", b["nombre"]); p.fill("#c", N.CL); p.click("button[type=submit]")
    p.wait_for_function("() => !!localStorage.getItem('nom.token')"); p.wait_for_timeout(1500)
    p.goto(N.B + "/#/doc/" + dt); p.wait_for_selector(".wd-cuerpo"); p.wait_for_timeout(1500)
    adb("shell", "cmd", "connectivity", "airplane-mode", "enable")
    t0 = time.time(); k = 0
    while time.time() - t0 < SEG:
        k += 1
        app(ESCRIBIR + f"('Sin red {k}. ')")
        try:
            p.click(".wd-cuerpo", timeout=10000); p.keyboard.press("Control+End"); p.keyboard.type(f"Nube {k}. "); p.wait_for_timeout(200)
        except Exception as e:
            foto = str(Path.home() / ".cache/nom-pruebas/registros/reales/rama-computadora.png"); p.screenshot(path=foto)
            print("la computadora ya no tiene el editor:", p.url, p.evaluate("() => document.body.innerText.slice(0, 300)"), foto, flush=True)
            break
        time.sleep(15)
    adb("shell", "cmd", "connectivity", "airplane-mode", "disable")
    print(f"(la app estuvo {time.time() - t0:.0f} s sin red)", flush=True)
    # El aviso de la rama dura 12 s: se busca mientras está a la vista.
    aviso = ""
    t1 = time.time()
    while time.time() - t1 < 40 and "rama" not in str(aviso).lower():
        aviso = app("[...document.querySelectorAll('.aviso, .alerta')].map((x) => x.innerText).join(' | ')"); time.sleep(1)
    time.sleep(3)
    st, ramas = N.api("GET", f"/api/documentos/{dt}/ramas", None, a["token"])
    lista = ramas if isinstance(ramas, list) else (ramas or {}).get("ramas", [])
    anota("se abrió una rama (no se juntó solo)", st == 200 and len(lista) >= 1, str(ramas)[:200])
    orig = texto_doc(dt, a["token"])
    anota("el original NO tiene lo escrito sin red", "Sin red 1" not in orig and "Nube 1" in orig)
    if lista:
        rid = lista[0].get("id") if isinstance(lista[0], dict) else lista[0]
        rama = texto_doc(rid, a["token"])
        anota("la rama tiene lo escrito sin red", "Sin red 1" in rama, rid)
    anota("la app avisa de la rama", "rama" in str(aviso).lower(), str(aviso)[:200])
    nav.close()
print(f"\nRama en la app: {sum(1 for r in res if r[1])} de {len(res)} bien")
sys.exit(0 if all(r[1] for r in res) else 1)

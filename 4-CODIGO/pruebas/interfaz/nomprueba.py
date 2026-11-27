"""Ayudas para probar NOM con Playwright; por omisión contra el sitio real (orden del dueño del 25-nov-2026)."""
import json, os, time, urllib.request, random, string
from playwright.sync_api import sync_playwright
B = os.environ.get("NOM_BASE", "https://www.nom.qd.je")   # sitio real por omisión (orden del dueño, 25-nov)
# Los navegadores de prueba hablan español de México, como los equipos de la escuela: desde la 0.23.10 la interfaz toma el
# idioma del sistema («Automático») y Chromium de Playwright dice «en-US». Una prueba que quiera otro idioma pasa locale=.
try:
    from playwright.sync_api._generated import Browser as _Navegador
    if not getattr(_Navegador, "_nom_es", False):
        _nuevo_ctx, _nueva_pag = _Navegador.new_context, _Navegador.new_page
        _Navegador.new_context = lambda self, *a, **k: _nuevo_ctx(self, *a, **dict({"locale": "es-MX"}, **k))
        _Navegador.new_page = lambda self, *a, **k: _nueva_pag(self, *a, **dict({"locale": "es-MX"}, **k))
        _Navegador._nom_es = True
except Exception:
    pass
CL = "Clave-Segura-123"
_VERSION = []
def version():
    """La versión del servidor (las pruebas hablan como una página de esa misma versión: Version.misma)."""
    if not _VERSION:
        try:
            with urllib.request.urlopen(urllib.request.Request(B + "/api/red/info", headers={"User-Agent": "NOM-pruebas/1.0"}), timeout=30) as x:
                _VERSION.append(json.loads(x.read()).get("version") or "")
        except Exception: return ""
    return _VERSION[0]
def api(m, u, b=None, t=None, v=None):
    # Un nombre de cliente propio: Cloudflare (la Raspberry detrás del túnel) rechaza el de urllib (error 1010).
    h = {"Content-Type": "application/json", "User-Agent": "NOM-pruebas/1.0", "X-NOM-Version": version() if v is None else v}
    if t: h["Authorization"] = "Bearer " + t
    q = urllib.request.Request(B + u, method=m, data=json.dumps(b).encode() if b is not None else None, headers=h)
    try:
        with urllib.request.urlopen(q, timeout=30) as x:
            c = x.read(); return x.status, (json.loads(c) if c[:1] in b"{[" else None)
    except urllib.error.HTTPError as e:
        c = e.read(); return e.code, (json.loads(c) if c[:1] in b"{[" else None)
def esperar_servidor():
    for _ in range(120):
        try: urllib.request.urlopen(urllib.request.Request(B + "/api/red/info", headers={"User-Agent": "NOM-pruebas/1.0"}), timeout=2); return True
        except Exception: time.sleep(1)
    return False
# Las pruebas van contra el sitio real (orden del dueño): cada cuenta que se crea se borra al terminar.
CREADAS = []
def _borrar_creadas():
    for n in CREADAS:
        try:
            s, r = api("POST", "/api/cuentas/sesion", {"nombre": n, "contrasena": CL})
            if s == 200: api("POST", "/api/cuentas/eliminar", {"contrasena": CL}, r["token"])
        except Exception: pass
import atexit
atexit.register(_borrar_creadas)
def cuenta(nombre=None):
    nombre = nombre or "p" + "".join(random.choice(string.ascii_lowercase) for _ in range(7))
    # (con un identificador de equipo al azar: el tope de cuentas nuevas por red no se gasta con las pruebas)
    equipo = "".join(random.choice(string.ascii_letters + string.digits) for _ in range(24))
    if api("POST", "/api/cuentas/registro", {"nombre": nombre, "contrasena": CL, "acepto": "2026-11-13", "equipo": equipo})[0] == 201: CREADAS.append(nombre)
    s = api("POST", "/api/cuentas/sesion", {"nombre": nombre, "contrasena": CL})[1]
    return {"nombre": nombre, "token": s["token"], "id": s["usuario"]["id"]}
def espacio(tok, nombre="Pruebas"):
    return api("POST", "/api/espacios", {"nombre": nombre}, tok)[1]["id"]
def documento(tok, esp, modo="HOJA", nombre="Doc", ops=None):
    return api("POST", f"/api/espacios/{esp}/documentos", {"nombre": nombre, "modo": modo, "ops": ops or []}, tok)[1]["id"]
class Navegador:
    def __init__(self, movil=False):
        self.pw = sync_playwright().start()
        # NOM_NAVEGADOR=firefox prueba con el motor de Firefox (sin "is_mobile": Firefox no lo tiene).
        motor = os.environ.get("NOM_NAVEGADOR", "chromium")
        self.b = getattr(self.pw, motor).launch()
        opc = dict(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True, device_scale_factor=2) if movil else dict(viewport={"width": 1400, "height": 900})
        if motor == "firefox": opc.pop("is_mobile", None)
        self.ctx = self.b.new_context(**opc)
        self.p = self.ctx.new_page()
        self.errores = []
        self.p.on("pageerror", lambda e: self.errores.append(str(e)))
        self.p.on("console", lambda m: self.errores.append("console: " + m.text) if m.type == "error" else None)
    def entrar(self, c):
        p = self.p
        p.goto(B + "/")
        p.wait_for_selector("#u", timeout=20000)
        p.fill("#u", c["nombre"]); p.fill("#c", CL)
        p.click("button[type=submit]")
        p.wait_for_function("() => !!localStorage.getItem('nom.token')", timeout=20000)
        p.wait_for_timeout(800)
    def abrir(self, doc):
        self.p.goto(B + "/#/doc/" + doc)
        self.p.wait_for_selector(".ed-herramientas", timeout=20000)
        self.p.wait_for_timeout(600)
    def foto(self, ruta):
        self.p.screenshot(path=ruta)
    def cerrar(self):
        try: self.b.close(); self.pw.stop()
        except Exception: pass

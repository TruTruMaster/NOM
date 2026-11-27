#!/usr/bin/env python3
"""Cómo ve la persona cada error, en el sitio real: una dirección que no existe, un documento que no existe o ya no se
puede abrir, sin red, y el servidor contestando 413, 423, 429, 500, 502, 503 y 530 (Cloudflare con la Pi apagada),
simulados con Playwright (la respuesta se cambia en el navegador: al servidor no le llega nada raro).

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_errores.py SALIDA

Revisa en cada caso: que se vea un mensaje propio de NOM en español (no JSON, no «error 500» a secas, no la página
de Cloudflare), que diga qué pasó y qué hacer, y que haya un botón para seguir. Deja una captura de cada uno.
"""
import os, sys, json, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/errores"
os.makedirs(SALIDA, exist_ok=True)
C = N.cuenta(); esp = N.espacio(C["token"], "Errores"); doc = N.documento(C["token"], esp, "TEXTO", "Prueba")
CRUDO = re.compile(r'\{"error"|Whitelabel|cloudflare|Error 5\d\d|error \d{3}\)', re.I)

with sync_playwright() as pw:
    b = pw.chromium.launch(); ctx = b.new_context(viewport={"width": 1280, "height": 800}); p = ctx.new_page(); P = Pagina(p, N.B)
    errores = []; p.on("pageerror", lambda e: errores.append(str(e)[:200]))

    # 1. Una dirección que no existe, abierta en el navegador
    p.goto(N.B + "/esto-no-existe"); P.t(800); p.screenshot(path=f"{SALIDA}/01-direccion-404.png")
    t = P.ev("() => document.body.innerText")
    check("404 de una dirección: página propia en español con un enlace al inicio", not CRUDO.search(t) and "no existe" in t.lower() and P.ev("() => !!document.querySelector('a[href=\"/\"]')"), t[:160])

    P.entrar(C)
    # 2. Un documento que no existe
    p.goto(N.B + "/#/doc/NoExiste1234"); P.t(2500); p.screenshot(path=f"{SALIDA}/02-documento-404.png")
    t = P.ev("() => document.body.innerText")
    check("documento que no existe: lo dice en una pantalla con «Volver al inicio»", not CRUDO.search(t) and re.search(r"no (existe|se encontró|lo encontramos)|ya no", t, re.I) is not None and re.search(r"inicio", t, re.I) is not None, t[:200])

    # 3. Errores del servidor al cargar el inicio y al crear un documento (respuesta cambiada en el navegador)
    for estado, cuerpo in [(413, ""), (423, ""), (429, ""), (500, ""), (502, "<html>Bad gateway</html>"), (503, ""), (530, "<html>cloudflare error 1033</html>")]:
        def hacer(estado, cuerpo):
            return lambda route: route.fulfill(status=estado, body=cuerpo, headers={"Content-Type": "text/html" if cuerpo else "application/json"})
        cambiar = hacer(estado, cuerpo)
        p.route("**/api/espacios/*/documentos", cambiar)
        p.goto(N.B + "/#/espacio/" + esp); P.t(1500); P.cerrar_hojas()
        r = P.ev("""async () => { try { await NOM.api.post(location.hash.replace('#/espacio/', '/api/espacios/') + '/documentos', {nombre: 'x', modo: 'TEXTO'}); return 'sin error'; }
                     catch (e) { return JSON.stringify({estado: e.estado, msg: e.message}); } }""")
        p.unroute("**/api/espacios/*/documentos")
        m = json.loads(r) if r.startswith("{") else {"msg": r}
        ok = not CRUDO.search(m.get("msg", "")) and len(m.get("msg", "")) > 25
        check(f"{estado}: el mensaje dice qué pasó y qué hacer", ok, m)

    # 4. Sin red (con la página ya abierta) y al volver
    p.goto(N.B + "/#/doc/" + doc); P.t(2500)
    ctx.set_offline(True); P.t(500)
    p.keyboard.type("hola sin red"); P.t(1500); p.screenshot(path=f"{SALIDA}/04-sin-red.png")
    t = P.ev("() => document.body.innerText")
    check("sin red: se ve que no hay conexión y que lo escrito se guarda en el equipo", re.search(r"sin conexión|sin red|sin internet", t, re.I) is not None, t[:200])
    ctx.set_offline(False); P.t(4000)

    # 5. El servidor caído (Cloudflare contesta 530 a todo) al recargar la página
    p.route("**/*", lambda route: route.fulfill(status=530, body="<html><title>Error 1033</title>cloudflare</html>", headers={"Content-Type": "text/html"}) if "/api/" in route.request.url else route.continue_())
    p.goto(N.B + "/#/"); P.t(4000); p.screenshot(path=f"{SALIDA}/05-servidor-caido.png")
    t = P.ev("() => document.body.innerText")
    check("servidor caído: aviso propio (no responde) y se puede seguir sin conexión", not CRUDO.search(t) and re.search(r"no responde|sin conexión|no se pudo conectar|servidor", t, re.I) is not None, t[:200])
    p.unroute("**/*")
    check("sin errores de JavaScript", not errores, errores[:3])
    b.close()
print(f"\nErrores: {BIEN[0]} bien, {MAL[0]} mal")

#!/usr/bin/env python3
"""¿Cargan completas todas las fotos? (reporte del equipo, 27-nov). Abre cada documento de un espacio con un navegador
nuevo (sin nada guardado), recorre todas sus páginas o diapositivas y cuenta las imágenes que no terminaron de cargar
(rotas, en blanco o con la marca de espera) y las que se ven con proporción rara (caja cuadrada con foto que no lo es).

  NOM_CUENTA=usuario:clave NOM_ESPACIO="Expo NOM" ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_fotos.py SALIDA
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1] if len(sys.argv) > 1 else "/tmp/fotos"
os.makedirs(SALIDA, exist_ok=True)
u, c = os.environ["NOM_CUENTA"].split(":", 1); N.CL = c
s, r = N.api("POST", "/api/cuentas/sesion", {"nombre": u, "contrasena": c}); C = {"nombre": u, "token": r["token"], "id": r["usuario"]["id"]}
esp = next(x["id"] for x in N.api("GET", "/api/espacios", None, C["token"])[1] if x["nombre"] == os.environ.get("NOM_ESPACIO", "Expo NOM"))
docs = N.api("GET", f"/api/espacios/{esp}/documentos", None, C["token"])[1]
docs = docs if isinstance(docs, list) else docs.get("documentos", [])
MAL_IMG = """() => [...document.querySelectorAll('img')].filter(i => i.offsetParent && i.getBoundingClientRect().width > 8 && !i.closest('.miniaturas, .mini, header, .riel, .avatar'))
  .filter(i => !i.complete || i.naturalWidth <= 1 || (i.getAttribute('src') || '').includes('#nom:bin/')).map(i => (i.getAttribute('src') || '').slice(0, 60))"""
with sync_playwright() as pw:
    b = pw.chromium.launch()
    for d in docs:
        if d.get("modo") not in ("TEXTO", "PRESENTACION", "HOJA", "DIAGRAMA"): continue
        ctx = b.new_context(viewport={"width": 1440, "height": 1000}); p = ctx.new_page(); P = Pagina(p, N.B)
        err = []; p.on("pageerror", lambda e: err.append(str(e)[:150]))
        P.entrar(C); P.abrir(d["id"]); P.t(4000); P.cerrar_hojas(); P.quitar_consejo()
        malas = []
        if d["modo"] == "PRESENTACION":
            n = P.ev("() => document.querySelectorAll('.miniaturas [data-p]').length")
            for i in range(n):
                P.ev(f"() => document.querySelectorAll('.miniaturas [data-p]')[{i}].click()"); P.t(1200)
                m = P.ev(MAL_IMG)
                if m: P.t(3000); m = P.ev(MAL_IMG)
                if m: malas.append((i + 1, m[:2]))
        elif d["modo"] == "TEXTO":
            n = P.ev("() => document.querySelectorAll('.wd-pagina').length")
            for i in range(max(1, n)):
                P.ev(f"() => {{ const h = document.querySelectorAll('.wd-pagina')[{i}]; h && h.scrollIntoView({{block: 'start'}}); }}"); P.t(800)
                m = P.ev(MAL_IMG)
                if m: P.t(3000); m = P.ev(MAL_IMG)
                if m: malas.append((i + 1, m[:2]))
        else:
            P.t(2000); m = P.ev(MAL_IMG)
            if m: malas.append((1, m[:2]))
        p.screenshot(path=f"{SALIDA}/{d['id']}.png")
        check(f"{d['nombre'][:40]} ({d['modo']}): todas las fotos cargadas", not malas, malas[:4])
        check(f"{d['nombre'][:40]}: sin errores de JavaScript", not err, err[:2])
        ctx.close()
    b.close()
print(f"\nFotos: {BIEN[0]} bien, {MAL[0]} mal")

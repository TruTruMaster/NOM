#!/usr/bin/env python3
"""Importa archivos de Office reales con el botón «Importar archivo», como una persona, y deja una captura de cada
página (Word), cada hoja (Excel) y cada diapositiva (PowerPoint), con lo que salió mal en cada una:
errores de JavaScript, imágenes que no cargaron completas y avisos que vio la persona.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/exploracion_archivos_dueno.py SALIDA ARCHIVO [ARCHIVO …]
  (NOM_CUENTA=usuario:clave y NOM_ESPACIO=nombre importan en esa cuenta y ese espacio y no borran nada;
   sin NOM_CUENTA se usa una cuenta desechable que se borra al final)
"""
import os, sys, time, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1]; ARCHIVOS = sys.argv[2:]
os.makedirs(SALIDA, exist_ok=True)
cuenta_fija = os.environ.get("NOM_CUENTA")
if cuenta_fija:
    u, c = cuenta_fija.split(":", 1)
    N.CL = c   # (Pagina.entrar escribe N.CL como contraseña)
    s, r = N.api("POST", "/api/cuentas/sesion", {"nombre": u, "contrasena": c})
    C = {"nombre": u, "contrasena": c, "token": r["token"], "id": r["usuario"]["id"]}
    esp = next(x["id"] for x in N.api("GET", "/api/espacios", None, C["token"])[1] if x.get("nombre") == os.environ.get("NOM_ESPACIO", "Expo NOM"))
else:
    C = N.cuenta(); esp = N.espacio(C["token"], "Archivos " + str(int(time.time()) % 100000))

IMAGENES_MAL = """() => [...document.querySelectorAll('.editor img, .vista img, main img')].filter(i => i.offsetParent && i.getBoundingClientRect().width > 4)
  .filter(i => !i.complete || i.naturalWidth === 0).map(i => (i.getAttribute('src') || '').slice(0, 80))"""
informe = {}
with sync_playwright() as pw:
    b = pw.chromium.launch()
    ctx = b.new_context(viewport={"width": 1440, "height": 1200}, accept_downloads=True); p = ctx.new_page(); P = Pagina(p, N.B)
    errores = []
    p.on("pageerror", lambda e: errores.append("pageerror: " + str(e)[:300]))
    p.on("console", lambda m: m.type == "error" and errores.append("console: " + m.text[:300]))
    P.entrar(C)
    for ruta in ARCHIVOS:
        nombre = os.path.basename(ruta); base = os.path.splitext(nombre)[0].replace(" ", "_").replace("(", "").replace(")", "")
        errores.clear(); r = informe[nombre] = {"errores": [], "imagenes_mal": [], "avisos": [], "capturas": []}
        print("==", nombre, flush=True)
        p.goto(N.B + "/#/espacio/" + esp); P.t(1500); P.cerrar_hojas()
        P.ev("() => { const b = [...document.querySelectorAll('button, a')].find(x => /Nuevo documento|Crear documento/.test((x.getAttribute('aria-label') || '') + ' ' + x.textContent) && x.offsetParent); if (b) b.click(); else { const f = document.querySelector('[aria-label=\"Nuevo documento\"], .fab'); f && f.click(); } }")
        P.t(800)
        t0 = time.time()
        with p.expect_file_chooser(timeout=15000) as fc:
            P.ev("() => document.querySelector('[data-importar]').click()")
        fc.value.set_files(ruta)
        P.t(800)
        # «¿Dónde se guarda?» (0.23.10): en el espacio
        P.ev("() => { const b = [...document.querySelectorAll('.velo button, .hoja button')].find(x => /^En «/.test(x.textContent.trim())); b && b.click(); }")
        try:
            p.wait_for_function("() => location.hash.startsWith('#/doc/') && window.NOM && NOM.editorActual && NOM.editorActual.vista && NOM.editorActual.vista.isConnected", timeout=180000)
        except Exception as e:
            r["errores"].append("no abrió el documento: " + str(e)[:200]); p.screenshot(path=f"{SALIDA}/{base}-no-abre.png"); r["errores"] += errores[:]; continue
        r["segundos_hasta_abrir"] = round(time.time() - t0, 1)
        r["doc"] = p.url.split("#/doc/")[-1]
        P.t(4000); P.quitar_consejo(); P.cerrar_hojas()
        modo = P.ev("() => NOM.editorActual.modo || (document.querySelector('.miniaturas') ? 'PRESENTACION' : document.querySelector('.pestana-hoja') ? 'HOJA' : 'TEXTO')")
        r["modo"] = modo
        if modo == "PRESENTACION" or P.ev("() => !!document.querySelector('.miniaturas [data-p]')"):
            n = P.ev("() => document.querySelectorAll('.miniaturas [data-p]').length"); r["diapositivas"] = n
            for i in range(n):
                P.ev(f"() => document.querySelectorAll('.miniaturas [data-p]')[{i}].click()"); P.t(1800)
                f = f"{SALIDA}/{base}-d{i+1:02d}.png"; p.screenshot(path=f); r["capturas"].append(f)
                mal = P.ev(IMAGENES_MAL)
                if mal: r["imagenes_mal"].append([i + 1, mal])
        elif P.ev("() => !!document.querySelector('.pestana-hoja')"):
            n = P.ev("() => document.querySelectorAll('.pestana-hoja').length"); r["hojas"] = n
            for i in range(n):
                P.ev(f"() => document.querySelectorAll('.pestana-hoja')[{i}].click()"); P.t(1800)
                f = f"{SALIDA}/{base}-h{i+1:02d}.png"; p.screenshot(path=f); r["capturas"].append(f)
                mal = P.ev(IMAGENES_MAL)
                if mal: r["imagenes_mal"].append([i + 1, mal])
        else:
            n = P.ev("() => document.querySelectorAll('.wd-pagina').length"); r["paginas"] = n
            for i in range(max(n, 1)):
                P.ev(f"() => {{ const h = document.querySelectorAll('.wd-pagina')[{i}]; h && h.scrollIntoView({{block: 'start'}}); }}"); P.t(900)
                caja = P.ev(f"() => {{ const h = document.querySelectorAll('.wd-pagina')[{i}]; if (!h) return null; const r = h.getBoundingClientRect(); return {{x: r.left, y: r.top, width: r.width, height: Math.min(r.height, innerHeight - r.top)}}; }}")
                f = f"{SALIDA}/{base}-p{i+1:02d}.png"; p.screenshot(path=f, clip=caja) if caja else p.screenshot(path=f); r["capturas"].append(f)
                mal = P.ev(IMAGENES_MAL)
                if mal: r["imagenes_mal"].append([i + 1, mal])
        r["avisos"] = P.ev("() => [...document.querySelectorAll('.aviso, .toast, [role=alert], [role=status]')].map(x => x.textContent.trim()).filter(Boolean).slice(0, 10)")
        r["errores"] += errores[:]
        print(json.dumps({k: v for k, v in r.items() if k != "capturas"}, ensure_ascii=False)[:1500], flush=True)
    b.close()
json.dump(informe, open(f"{SALIDA}/informe.json", "w"), ensure_ascii=False, indent=1)
# (la cuenta desechable la borra nomprueba al salir)
print("listo:", SALIDA)

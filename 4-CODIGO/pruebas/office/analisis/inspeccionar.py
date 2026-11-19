#!/usr/bin/env python3
"""Abre un .docx/.xlsx/.pptx en NOM (como `comparar_real.py`) y corre un JavaScript en la página: para depurar.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/analisis/inspeccionar.py \
      ARCHIVO "() => …" [--antes "window.__x = 1"] [--captura salida.png]
"""
import argparse, base64, json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent)); sys.path.insert(0, str(Path(__file__).parent.parent.parent / "interfaz"))
import comparar_real as C
import nomprueba as N
import prueba_ui as U

ap = argparse.ArgumentParser()
ap.add_argument("archivo"); ap.add_argument("js", nargs="+"); ap.add_argument("--antes"); ap.add_argument("--captura")
a = ap.parse_args()
N.esperar_servidor()
c, esp = C.entrar()
nav = N.Navegador()
nav.ctx.close(); nav.ctx = nav.b.new_context(viewport={"width": 1300, "height": 1150}); nav.p = nav.ctx.new_page()
nav.p.on("console", lambda m: print("consola:", m.text[:300]) if m.type in ("error", "warning") else None)
if a.antes: nav.p.add_init_script(a.antes)
nav.entrar(c)
nav.p.evaluate("() => { const t = NOM.leer('nom.texto', {}); t.vista = 'paginas'; t.zoom = 1; localStorage.setItem('nom.texto', JSON.stringify(t)); }")
f = Path(a.archivo)
if not f.exists(): f = C.REAL / "archivos" / f.name
p = nav.p
try:
    p.evaluate(C.JS_IMPORTAR, [base64.b64encode(f.read_bytes()).decode(), f.name, esp])
    p.wait_for_function(U.JS_LISTO, timeout=120000)
    p.evaluate("() => document.fonts.ready"); p.wait_for_timeout(1500)
    if f.suffix == ".docx": p.evaluate("() => NOM.editorActual.repaginar && NOM.editorActual.repaginar()"); p.wait_for_timeout(800)
    for js in a.js:
        r = p.evaluate(js)
        print(json.dumps(r, ensure_ascii=False, indent=1) if not isinstance(r, str) else r)
    if a.captura: p.screenshot(path=a.captura, full_page=False)
finally:
    nav.cerrar()

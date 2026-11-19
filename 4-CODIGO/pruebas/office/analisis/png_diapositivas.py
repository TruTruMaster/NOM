#!/usr/bin/env python3
"""El PNG de cada diapositiva que exporta NOM (NOM.diapo.diapoPng) contra la imagen que exporta PowerPoint real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/analisis/png_diapositivas.py p01-texto.pptx …

Deja los PNG en ~/NOM/trabajo/png-diapositivas/ y dice el parecido de cada uno (la misma medida de comparar_real.py).
"""
import base64, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent)); sys.path.insert(0, str(Path(__file__).parent.parent.parent / "interfaz"))
import comparar_real as C
import nomprueba as N
import prueba_ui as U

SALE = Path.home() / "NOM/trabajo/png-diapositivas"
SALE.mkdir(parents=True, exist_ok=True)
N.esperar_servidor()
c, esp = C.entrar()
nav = N.Navegador()
nav.entrar(c)
try:
    for a in sys.argv[1:]:
        f = Path(a)
        if not f.exists(): f = C.REAL / "archivos" / f.name
        p = nav.p
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(f.read_bytes()).decode(), f.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000)
        p.evaluate("() => document.fonts.ready"); p.wait_for_timeout(1200)
        pngs = p.evaluate("""async () => { const d = NOM.editorActual.doc; const out = [];
          for (const pg of d.paginas()) out.push(await NOM.diapo.diapoPng(d, pg, 1)); return out; }""")
        vals = []
        for i, u in enumerate(pngs, 1):
            ruta = SALE / f"{f.stem}-{i:02d}.png"
            ruta.write_bytes(base64.b64decode(u.split(",", 1)[1]))
            ref = C.REAL / "verdad" / f"{f.stem}-{i:02d}.png"
            if ref.exists(): vals.append(round(float(C.parecido_img(ref, ruta)), 1))
        print(f"{f.name}: PNG de NOM contra PowerPoint, parecido {vals} · medio {round(sum(vals) / len(vals), 1) if vals else None}", flush=True)
finally:
    nav.cerrar()

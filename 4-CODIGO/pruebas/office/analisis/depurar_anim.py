"""Abre un .pptx en NOM, prepara una diapositiva con el motor y corre un JS de depuración en el tiempo t.
   depurar_anim.py ARCHIVO DIAPO T 'js expr (recibe d = elemento de la diapositiva, m = motor)'"""
import sys, json
sys.path.insert(0, "/home/user/NOM/4-CODIGO/pruebas/office"); sys.path.insert(0, "/home/user/NOM/4-CODIGO/pruebas/interfaz")
import nomprueba as N, comparar_real as C, animaciones_real as R
from pathlib import Path
arch = Path(sys.argv[1]); diapo = int(sys.argv[2]); t = float(sys.argv[3]); js = sys.argv[4]
if not arch.exists(): arch = C.REAL / "archivos" / arch.name
N.esperar_servidor(); c, esp = C.entrar()
nav = N.Navegador(movil=False)
nav.ctx.close(); nav.ctx = nav.b.new_context(viewport={"width": 1300, "height": 760}); nav.p = nav.ctx.new_page(); nav.p.on("pageerror", lambda e: nav.errores.append(str(e)))
nav.p.on("console", lambda m: print("consola:", m.text))
nav.entrar(c); p = nav.p
try:
    C.importar(nav, esp, arch)
    print(p.evaluate(R.JS_PREPARAR, [diapo - 1, 1280 / 960]))
    p.evaluate("(t) => window.__motor && window.__motor.cuadro(0, t)", t)
    print(json.dumps(p.evaluate("() => { const d = document.querySelector('#__anim_prueba > .diapo'), m = window.__motor; return (" + js + "); }"), ensure_ascii=False, indent=1))
    if len(sys.argv) > 5: p.screenshot(path=sys.argv[5], clip={"x": 0, "y": 0, "width": 1280, "height": 720})
finally:
    print("errores:", nav.errores[:5]); nav.cerrar()

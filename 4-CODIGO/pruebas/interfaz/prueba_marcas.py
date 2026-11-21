#!/usr/bin/env python3
"""Mostrar marcas de formato (¶, como Word) en el editor de texto.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_marcas.py [ARCHIVO.docx]

Con un Word real (w09: tabulaciones con relleno, índices) y uno escrito aquí: al prender las marcas se ven ¶ al final de
cada párrafo, un punto en cada espacio y una flecha en cada tabulación, ningún renglón se mueve (los mismos renglones en
el mismo lugar), el texto del documento no cambia, se puede escribir con las marcas prendidas y al apagarlas todo queda
como estaba. No se imprimen.
"""
import base64, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "office"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:400]}"), flush=True)


# los renglones de cada párrafo: [párrafo, renglón, arriba, izquierda, ancho] (redondeado a medio píxel)
JS_RENGLONES = r"""() => { const out = []; document.querySelectorAll('.wd-cuerpo > *').forEach((el, i) => {
  const r = document.createRange(); r.selectNodeContents(el);
  const vistos = new Map();
  for (const q of r.getClientRects()) { if (q.width < 0.5) continue; const k = Math.round(q.top * 2) / 2; const v = vistos.get(k); vistos.set(k, v ? [Math.min(v[0], q.left), Math.max(v[1], q.right)] : [q.left, q.right]); }
  [...vistos.entries()].sort((a, b) => a[0] - b[0]).forEach(([t, [l, rr]], j) => out.push([i, j, t, Math.round(l * 2) / 2, Math.round((rr - l) * 2) / 2])); });
  return out; }"""
JS_TEXTO = "() => NOM.textoRico.parrafosDe(NOM.editorActual.doc).map((p) => p.tramos.map((t) => t.x || '').join('')).join('\\n')"
JS_CUENTA = """() => ({ me: document.querySelectorAll('.wd-me').length, mt: document.querySelectorAll('.wd-tabw').length,
  pilcrow: [...document.querySelectorAll('.wd-cuerpo > p')].some((e) => getComputedStyle(e, '::after').content.includes('¶')) ? '¶' : 'no' })"""


def main():
    # (la regresión pasa la carpeta de capturas: entonces, el w09 de siempre)
    arch = Path(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].endswith(".docx") else Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real/archivos/w09-tabulaciones.docx"
    if not arch.exists():
        cand = sorted((arch.parent).glob("w09*.docx")); arch = cand[0] if cand else arch
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    try:
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(2500)
        api = "NOM.editorActual.modoApi.marcasFormato"
        # (la base: repintado sin marcas, para comparar marcas contra repintado y no contra el primer pintado)
        p.evaluate(f"() => {api}.alternar(false)"); p.wait_for_timeout(2500)
        print("   ocultos:", p.evaluate("() => document.querySelectorAll('.wd-oculto').length"), "acomodados:", p.evaluate("() => document.querySelectorAll('.wd-acomodo').length"))
        r0, t0 = p.evaluate(JS_RENGLONES), p.evaluate(JS_TEXTO)
        p.evaluate(f"() => {api}.alternar(true)"); p.wait_for_timeout(2500)
        r1, t1, cu = p.evaluate(JS_RENGLONES), p.evaluate(JS_TEXTO), p.evaluate(JS_CUENTA)
        print("   con marcas, acomodados:", p.evaluate("() => document.querySelectorAll('.wd-acomodo').length"))
        anota(f"{arch.name}: se ven ¶, los espacios y las tabulaciones", cu["me"] > 50 and cu["mt"] > 0 and "¶" in cu["pilcrow"], cu)
        cerca = lambda a, b: a[:2] == b[:2] and all(abs(x - y) <= 1 for x, y in zip(a[2:], b[2:]))   # (1 px: el redondeo)
        dif = [(a, b) for a, b in zip(r0, r1) if not cerca(a, b)]
        anota(f"…y ningún renglón se mueve ({len(r0)} renglones)", len(r0) == len(r1) and not dif, f"{len(r0)} / {len(r1)} · {dif[:4]}")
        anota("…ni cambia el texto del documento", t0 == t1)
        p.evaluate(f"() => {api}.alternar(false)"); p.wait_for_timeout(2500)
        r2 = p.evaluate(JS_RENGLONES)
        anota("al apagarlas todo queda como estaba", len(r2) == len(r0) and all(cerca(a, b) for a, b in zip(r0, r2)) and not p.evaluate(JS_CUENTA)["me"], [(a, b) for a, b in zip(r0, r2) if not cerca(a, b)][:3])
        # escribir con las marcas prendidas
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'Marcas', modo: 'TEXTO', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function(U.JS_LISTO, timeout=60000); p.wait_for_timeout(1000)
        p.evaluate(f"() => {api}.alternar(true)"); p.wait_for_timeout(500)
        p.locator(".wd [contenteditable=true]").first.click(); p.wait_for_timeout(200)
        p.keyboard.type("Hola mundo"); p.keyboard.press("Tab"); p.keyboard.type("fin."); p.keyboard.press("Enter"); p.keyboard.type("Otro dos")
        for _ in range(3): p.keyboard.press("Backspace")
        p.keyboard.type("párrafo")
        p.wait_for_timeout(800)
        t = p.evaluate(JS_TEXTO)
        anota("se escribe y se borra con las marcas prendidas", t.startswith("Hola mundo\tfin.\nOtro párrafo"), repr(t))
        anota("…y se ven en lo escrito", p.evaluate(JS_CUENTA)["me"] >= 2, p.evaluate(JS_CUENTA))
        p.emulate_media(media="print"); p.wait_for_timeout(300)
        anota("no se imprimen", p.evaluate("() => getComputedStyle(document.querySelector('.wd-me'), '::after').content") in ("none", "normal"))
        p.emulate_media(media="screen")
        p.evaluate(f"() => {api}.alternar(false)")
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nMarcas de formato: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

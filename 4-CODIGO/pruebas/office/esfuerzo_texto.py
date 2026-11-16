#!/usr/bin/env python3
"""Lo pesado que una persona sí llega a hacer en un documento largo (una tesis), medido por la interfaz.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/esfuerzo_texto.py ARCHIVO.docx

1. Buscar una palabra común (Ctrl+F) y reemplazar todas.
2. Escribir dentro de una celda de una tabla del documento.
3. Seleccionar todo (Ctrl+A), copiar y pegar al final: el documento queda del doble.
4. Deshacer el pegado (Ctrl+Z) y volver a como estaba.
5. Seleccionar todo y borrar; deshacer.
Cada paso: que termine, cuánto tarda y que la otra persona (otra pestaña con otra cuenta) quede igual.
"""
import base64, sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "interfaz"))
import nomprueba as N  # noqa: E402

LISTO = "() => !!(window.NOM && NOM.editorActual && NOM.editorActual.doc && NOM.editorActual.vista && NOM.editorActual.vista.isConnected && !NOM.editorActual.doc.pendientes.length)"
LARGO = "() => document.querySelector('.wd-cuerpo').innerText.length"
res = []


def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det else ""), flush=True)


def main():
    archivo = Path(sys.argv[1])
    a, b = N.cuenta(), N.cuenta()
    esp = N.espacio(a["token"], "Esfuerzo")
    N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": b["nombre"], "permiso": "EDICION"}, a["token"])
    A = N.Navegador(); A.p.set_default_timeout(300000); A.entrar(a); p = A.p
    A.ctx.grant_permissions(["clipboard-read", "clipboard-write"], origin=N.B)
    p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
    doc = p.evaluate("""async ([b64, nombre, espacio]) => {
      const bin = atob(b64), bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const r = await NOM.formatos.importar(new File([bytes], nombre));
      const d = await NOM.api.post(`/api/espacios/${espacio}/documentos`, { nombre: r.nombre, modo: r.modo, ops: NOM.DocCRDT.partir(r.ops) }, { ms: 120000 });
      location.hash = "#/doc/" + d.id; return d.id; }""", [base64.b64encode(archivo.read_bytes()).decode(), archivo.name, esp])
    p.wait_for_function(LISTO); p.wait_for_timeout(2500)
    ctxB = A.b.new_context(viewport={"width": 1300, "height": 860}); pb = ctxB.new_page()
    pb.goto(N.B + "/"); pb.wait_for_selector("#u"); pb.fill("#u", b["nombre"]); pb.fill("#c", N.CL); pb.click("button[type=submit]")
    pb.wait_for_function("() => !!localStorage.getItem('nom.token')"); pb.wait_for_timeout(1500)
    pb.goto(N.B + "/#/doc/" + doc); pb.wait_for_function(LISTO); pb.wait_for_timeout(1500)
    n0 = p.evaluate(LARGO)
    print(f"documento: {n0} letras")

    # Se compara el CONTENIDO (el texto del documento y las celdas de sus tablas), no lo que se pinta: quien
    # escribe en una tabla ve además la barrita de la tabla («A la celda, A la fila…»), y la otra persona no.
    CONTENIDO = """() => { const d = NOM.editorActual.doc; const v = d.secuencia('cuerpo').visibles();
      const t = d.vivos().filter((b) => b.tipo === 'tabla').map((b) => b.id + ':' + Object.keys(b).filter((k) => k.startsWith('c:')).sort().map((k) => k + '=' + b[k]).join('|')).sort().join('\\n');
      return v.map((x) => (x.e ? '\\u0001' : x.c)).join('') + '\\u0002' + t; }"""

    def igual_en_b(etq, s=120):
        t0 = time.time()
        while time.time() - t0 < s:
            if p.evaluate(CONTENIDO) == pb.evaluate(CONTENIDO): return anota(f"{etq}: la otra persona queda igual", True, f"{time.time() - t0:.1f} s")
            time.sleep(1)
        anota(f"{etq}: la otra persona queda igual", False, f"pantalla {p.evaluate(LARGO)} contra {pb.evaluate(LARGO)}")

    # 1. Buscar y reemplazar (Ctrl+H abre el panel con el campo de reemplazo)
    t0 = time.time()
    p.click(".wd-cuerpo"); p.keyboard.press("Control+h"); p.wait_for_timeout(800)
    if p.evaluate("() => !!document.querySelector('.wd-buscar [data-q]') && !!document.querySelector('.wd-buscar [data-r]')"):
        p.fill(".wd-buscar [data-q]", "frecuencia"); p.wait_for_timeout(1200)
        cuantos = p.evaluate("() => document.querySelector('.wd-buscar').innerText.replace(/\\s+/g, ' ').slice(0, 60)")
        p.fill(".wd-buscar [data-r]", "FRECUENCIA"); t0 = time.time()
        p.click(".wd-buscar [data-b=todos]"); p.wait_for_function(LISTO); p.wait_for_timeout(1500)
        hay = p.evaluate("() => (document.querySelector('.wd-cuerpo').innerText.match(/FRECUENCIA/g) || []).length")
        anota("buscar y reemplazar todo", hay > 0, f"panel: {cuantos!r} · {hay} reemplazos en {time.time() - t0:.1f} s")
    else: anota("buscar y reemplazar todo", False, "Ctrl+H no abrió el panel con reemplazo")
    p.keyboard.press("Escape"); p.wait_for_timeout(500)
    igual_en_b("reemplazar")

    # 2. Escribir en una celda de una tabla
    celda = p.evaluate("() => { const c = [...document.querySelectorAll('.wd-cuerpo .wd-celda')].find((x) => x.offsetParent && x.innerText.trim().length > 2); if (!c) return null; c.scrollIntoView({block: 'center'}); const r = c.getBoundingClientRect(); return [r.x + r.width - 6, r.y + r.height / 2]; }")
    if celda:
        p.mouse.click(*celda); p.keyboard.press("End"); p.keyboard.type(" CELDA-A"); p.wait_for_timeout(800)
        p.wait_for_function(LISTO)
        ok = p.evaluate("() => JSON.stringify(NOM.editorActual.doc.estado).includes('CELDA-A')")
        anota("escribir en una celda de una tabla", ok)
        try: pb.wait_for_function("() => JSON.stringify(NOM.editorActual.doc.estado).includes('CELDA-A')", timeout=20000); anota("la celda le llega a la otra persona", True)
        except Exception: anota("la celda le llega a la otra persona", False)
    else: anota("escribir en una celda de una tabla", False, "no hay tablas a la vista")

    # 3. Seleccionar todo, copiar y pegar al final
    p.click(".wd-cuerpo"); p.keyboard.press("Control+a"); p.wait_for_timeout(500)
    t0 = time.time(); p.keyboard.press("Control+c"); p.wait_for_timeout(1500)
    p.keyboard.press("Control+End"); p.wait_for_timeout(300)
    p.keyboard.press("Control+v")
    try:
        p.wait_for_function(f"() => document.querySelector('.wd-cuerpo').innerText.length > {int(n0 * 1.8)}", timeout=300000)
        p.wait_for_function(LISTO, timeout=300000)
        anota("copiar todo y pegarlo al final (el documento queda del doble)", True, f"{p.evaluate(LARGO)} letras en {time.time() - t0:.1f} s")
    except Exception as e:
        anota("copiar todo y pegarlo al final (el documento queda del doble)", False, f"{p.evaluate(LARGO)} letras · {str(e)[:120]}")
    igual_en_b("pegar")

    # 4. Deshacer el pegado
    t0 = time.time(); p.keyboard.press("Control+z")
    try:
        p.wait_for_function(f"() => document.querySelector('.wd-cuerpo').innerText.length < {int(n0 * 1.3)}", timeout=300000)
        p.wait_for_function(LISTO, timeout=300000)
        anota("deshacer el pegado", True, f"{p.evaluate(LARGO)} letras en {time.time() - t0:.1f} s")
    except Exception as e:
        anota("deshacer el pegado", False, f"{p.evaluate(LARGO)} letras")
    igual_en_b("deshacer")

    # 5. Seleccionar todo y borrar; deshacer
    p.click(".wd-cuerpo"); p.keyboard.press("Control+a"); t0 = time.time(); p.keyboard.press("Delete")
    try:
        p.wait_for_function("() => document.querySelector('.wd-cuerpo').innerText.length < 200", timeout=300000); p.wait_for_function(LISTO, timeout=300000)
        anota("seleccionar todo y borrar", True, f"{time.time() - t0:.1f} s")
    except Exception: anota("seleccionar todo y borrar", False, f"{p.evaluate(LARGO)} letras")
    igual_en_b("borrar todo")
    t0 = time.time(); p.keyboard.press("Control+z")
    try:
        p.wait_for_function(f"() => document.querySelector('.wd-cuerpo').innerText.length > {int(n0 * 0.9)}", timeout=300000); p.wait_for_function(LISTO, timeout=300000)
        anota("deshacer el borrado: vuelve todo", True, f"{p.evaluate(LARGO)} letras en {time.time() - t0:.1f} s")
    except Exception: anota("deshacer el borrado: vuelve todo", False, f"{p.evaluate(LARGO)} letras")
    igual_en_b("deshacer el borrado")
    errores = [e for e in A.errores if "Content Security" not in e]
    anota("sin errores de JavaScript", not errores, " | ".join(errores[:3]))
    A.cerrar()
    malos = [r for r in res if not r[1]]
    print(f"\nEsfuerzo en un documento largo: {len(res) - len(malos)} de {len(res)} bien")
    sys.exit(1 if malos else 0)


if __name__ == "__main__":
    main()

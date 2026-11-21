#!/usr/bin/env python3
"""Archivos reales de Office (Word, Excel, PowerPoint) usados como los usaría un equipo, de punta a punta y
por la interfaz.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/editar_real.py SALIDA ARCHIVO...

Por cada archivo:
  1. Ana lo importa con «Nuevo documento → Importar archivo» (el selector de archivos de verdad).
  2. Ana edita con teclado y mouse: en Excel escribe en una celda vacía de la primera hoja; en PowerPoint
     agrega texto a un cuadro de texto (doble clic); en Word, al final del primer párrafo con texto.
  3. Beto (otra cuenta con permiso de edición, en Firefox) abre el documento, ve lo de Ana y escribe lo
     suyo en otro lugar; a Ana le llega sin recargar.
  4. Ana recarga la página: sigue todo.
  5. Ana descarga el archivo y se compara con el original: lo escrito por los dos está; las demás partes
     salen idénticas; el validador de Microsoft (si está) no encuentra errores nuevos; Word, Excel o PowerPoint
     de verdad lo abren sin avisos ni «reparar».
Sale con código 1 si algo falla. Necesita el servidor de pruebas en NOM_BASE (por omisión el 8099).
"""
import json, os, re, signal, subprocess, sys, time, zipfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "interfaz"))
import nomprueba as N  # noqa: E402

LISTO = "() => !!(window.NOM && NOM.editorActual && NOM.editorActual.doc && NOM.editorActual.vista && NOM.editorActual.vista.isConnected && !NOM.editorActual.doc.pendientes.length)"
# ¿El documento tiene este texto? En el estado o, si se escribió letra por letra (cada letra es un elemento de la
# secuencia y en el estado quedan separadas), en el texto visible de cada secuencia.
TIENE = """(t) => { const d = NOM.editorActual.doc; if (JSON.stringify(d.estado).includes(t)) return true;
  try { for (const n of Object.keys(d.estado.seq || {})) { const s = d.secuencia(n); if (s && s.texto().includes(t)) return true; } } catch (e) { }
  return false; }"""
EXT = {"HOJA": "xlsx", "PRESENTACION": "pptx", "TEXTO": "docx"}
para_office = []   # (lo descargado, el original, nombre)
resultados = []


def anota(archivo, nombre, ok, det=""):
    resultados.append((archivo, nombre, ok, det))
    print(("  BIEN " if ok else "  MAL  ") + nombre + (f" · {det}" if det else ""), flush=True)


class Persona(N.Navegador):
    """Una persona más con el Playwright ya abierto (no se puede arrancar dos veces en el mismo hilo)."""
    def __init__(self, pw, motor="firefox"):
        self.pw = pw
        self.b = getattr(pw, motor).launch()
        self.ctx = self.b.new_context(viewport={"width": 1280, "height": 860}, accept_downloads=True)
        self.p = self.ctx.new_page()
        self.errores = []
        self.p.on("pageerror", lambda e: self.errores.append(str(e)))
        self.p.on("console", lambda m: self.errores.append("console: " + m.text) if m.type == "error" else None)

    def cerrar(self):
        try: self.b.close()
        except Exception: pass


# ------------------------------------------------------------ ediciones por tipo de documento
CELDA_VACIA = """(salto) => { const tds = [...document.querySelectorAll('td[data-c][data-f]')].filter((t) => t.offsetParent && !t.textContent.trim() && !t.hasAttribute('colspan') && !t.hasAttribute('rowspan'));
  const vista = document.querySelector('.hoja-rejilla, .rejilla, table') ; const lim = innerHeight - 80;
  const visibles = tds.filter((t) => { const r = t.getBoundingClientRect(); return r.top > 120 && r.bottom < lim && r.left > 40 && r.right < innerWidth - 20; });
  const t = visibles[Math.min(salto, visibles.length - 1)]; if (!t) return null; const r = t.getBoundingClientRect();
  return { xy: [r.x + r.width / 2, r.y + r.height / 2], ref: t.dataset.c + ',' + t.dataset.f }; }"""


def editar_hoja(p, texto, salto):
    c = p.evaluate(CELDA_VACIA, salto)
    if not c: return "no hay una celda vacía a la vista"
    p.mouse.click(*c["xy"]); p.wait_for_timeout(200)
    p.keyboard.type(texto); p.keyboard.press("Enter"); p.wait_for_timeout(400)
    return "celda " + c["ref"]


CUADRO = """(n) => { const els = [...document.querySelectorAll('.bl')].filter((x) => !x.closest('.mini') && x.offsetParent && x.textContent.trim().length > 0);
  const el = els[Math.min(n, els.length - 1)]; if (!el) return null; el.scrollIntoView({block: 'nearest'}); const r = el.getBoundingClientRect();
  return { xy: [r.x + Math.min(r.width / 2, 120), r.y + Math.min(r.height / 2, 30)], id: el.dataset.id, cuantos: els.length }; }"""


def nuevo_cuadro(p, texto, n):
    """Una presentación sin texto (solo figuras): se agrega un cuadro de texto con la barra, como lo haría
    una persona (también es una edición real: una figura nueva dentro de una diapositiva importada)."""
    b = p.evaluate("() => { const b = document.querySelector('.ed-herramientas [aria-label=\"Cuadro de texto\"]'); if (!b || !b.offsetParent) return null; const r = b.getBoundingClientRect(); return [r.x + r.width / 2, r.y + r.height / 2]; }")
    area = p.evaluate("(n) => { const a = document.querySelector('.escenario .diapo, .diapo[data-mundo]'); if (!a) return null; const r = a.getBoundingClientRect(); return [r.x + r.width * (0.25 + 0.3 * n), r.y + r.height * 0.82]; }", n)
    if not b or not area: return "no hay botón «Cuadro de texto» o diapositiva a la vista"
    # «Cuadro de texto» inserta el cuadro y ya queda escribiendo en él (no hay que hacer clic en la diapositiva).
    p.mouse.click(*b); p.wait_for_timeout(500)
    p.keyboard.type(texto.strip()); p.wait_for_timeout(300)
    p.keyboard.press("Escape"); p.wait_for_timeout(300)
    p.evaluate("() => document.activeElement && document.activeElement.blur && document.activeElement.blur()")
    p.wait_for_timeout(500)
    return "cuadro de texto nuevo (la diapositiva no tenía texto)"


def editar_diapo(p, texto, n):
    # Si la diapositiva a la vista no tiene cuadros con texto, se prueba con las siguientes (hasta 6).
    for k in range(6):
        c = p.evaluate(CUADRO, n)
        if c: break
        mini = p.evaluate("(k) => { const m = [...document.querySelectorAll('.mini')][k + 1]; if (!m) return null; const r = m.getBoundingClientRect(); return [r.x + r.width / 2, r.y + r.height / 2]; }", k)
        if not mini: break
        p.mouse.click(*mini); p.wait_for_timeout(500)
    if not c:
        first = p.evaluate("() => { const m = document.querySelector('.mini'); if (!m) return null; const r = m.getBoundingClientRect(); return [r.x + r.width / 2, r.y + r.height / 2]; }")
        if first: p.mouse.click(*first); p.wait_for_timeout(500)
        return nuevo_cuadro(p, texto, n)
    # Un doble clic, como lo hace una persona (antes: clic y doble clic casi seguidos, como un triple clic,
    # que en Firefox abre y vuelve a cerrar el texto).
    # (con muchas pruebas a la vez, Firefox tarda: se espera hasta 5 s y, si no abrió, otro doble clic)
    abierto = "() => !!document.querySelector('.bl [contenteditable=true], .bl.editando')"
    for _ in range(2):
        p.mouse.dblclick(*c["xy"])
        try: p.wait_for_function(abierto, timeout=5000); break
        except Exception: p.wait_for_timeout(300)
    if not p.evaluate(abierto):
        return "el doble clic no abrió el texto del cuadro"
    p.keyboard.press("Control+End"); p.keyboard.type(texto); p.wait_for_timeout(300)
    p.keyboard.press("Escape"); p.wait_for_timeout(300)
    p.evaluate("() => document.activeElement && document.activeElement.blur && document.activeElement.blur()")
    p.wait_for_timeout(500)
    return f"cuadro {c['id']} ({c['cuantos']} con texto a la vista)"


def editar_texto(p, texto, al_inicio):
    ok = p.evaluate("""(inicio) => { const ed = document.querySelector('.wd-cuerpo'); if (!ed) return false;
      const ps = [...ed.children].filter((x) => x.textContent.trim().length > 3 && !x.querySelector('[data-obj]'));
      const q = inicio ? (ps[1] || ps[0]) : ps[0]; if (!q) return false; q.scrollIntoView({block: 'center'});
      const r = document.createRange(); r.selectNodeContents(q); r.collapse(inicio); const s = getSelection(); s.removeAllRanges(); s.addRange(r); ed.focus(); return true; }""", al_inicio)
    if not ok: return "no hay un párrafo con texto"
    p.keyboard.type(texto); p.wait_for_timeout(400)
    return "párrafo"


def editar(p, modo, quien, n):
    marca = f"PRUEBA-{quien}"
    if modo == "HOJA": return editar_hoja(p, marca, n)
    if modo == "PRESENTACION": return editar_diapo(p, " " + marca, n)
    return editar_texto(p, (marca + " ") if n else (" " + marca), bool(n))


# ------------------------------------------------------------ comparar lo descargado con el original
def partes(ruta):
    """Partes del .zip por nombre, como las lee la norma de Office: sin distinguir mayúsculas y con las
    rutas al derecho (algunos programas las guardan con diagonal invertida; la app las corrige)."""
    with zipfile.ZipFile(ruta) as z:
        return {i.filename.replace("\\", "/").lower(): z.read(i.filename) for i in z.infolist() if not i.filename.endswith("/")}


def probar(pw, archivo, salida):
    nombre = archivo.name
    print(f"\n== {nombre} ({archivo.stat().st_size // 1024} KB)", flush=True)
    a, b = N.cuenta(), N.cuenta()
    esp = N.espacio(a["token"], "Archivos reales")
    N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": b["nombre"], "permiso": "EDICION"}, a["token"])
    A = Persona(pw, "chromium"); B = None
    try:
        A.p.set_default_timeout(120000)
        A.entrar(a)
        A.p.goto(N.B + "/#/espacio/" + esp); A.p.wait_for_timeout(1200)
        A.p.evaluate("(e) => NOM.vistaEspacios.nuevoDocumento({id: e, nombre: 'Archivos reales'})", esp)
        A.p.wait_for_selector(".velo [data-importar]")
        t0 = time.perf_counter()
        with A.p.expect_file_chooser() as fc:
            A.p.click(".velo [data-importar]")
        fc.value.set_files(str(archivo))
        try:
            A.p.wait_for_function("() => location.hash.startsWith('#/doc/')", timeout=120000)
        except Exception:
            avisos = A.p.evaluate("() => [...document.querySelectorAll('.aviso')].map((x) => x.textContent.trim()).join(' | ')")
            anota(nombre, "importar por la interfaz", False, avisos[:200]); return
        A.p.wait_for_function(LISTO); A.p.wait_for_timeout(2000)
        doc = A.p.evaluate("() => location.hash.slice(6)")
        modo = A.p.evaluate("() => NOM.editorActual.modo")
        aviso = A.p.evaluate("() => [...document.querySelectorAll('.avisos .aviso')].map((x) => x.textContent.trim()).join(' | ')")
        anota(nombre, "importar por la interfaz", True, f"{modo}, listo en {time.perf_counter() - t0:.1f} s" + (f" · aviso: {aviso[:160]}" if aviso else ""))
        A.foto(str(salida / f"{archivo.stem}-1-abierto.png"))
        # Ana edita
        donde = editar(A.p, modo, "A", 0)
        A.p.wait_for_function(LISTO)
        anota(nombre, "Ana edita con teclado y mouse", A.p.evaluate(TIENE, "PRUEBA-A"), donde)
        # Beto lo abre en Firefox, ve lo de Ana y escribe lo suyo
        B = Persona(pw, "firefox"); B.p.set_default_timeout(120000)
        B.entrar(b)
        B.p.goto(N.B + "/#/doc/" + doc); B.p.wait_for_function(LISTO); B.p.wait_for_timeout(1500)
        anota(nombre, "Beto (Firefox) abre el documento y tiene lo de Ana", B.p.evaluate(TIENE, "PRUEBA-A"))
        donde = editar(B.p, modo, "B", 1)
        B.p.wait_for_function(LISTO)
        t0 = time.perf_counter()
        try:
            A.p.wait_for_function(TIENE, arg="PRUEBA-B", timeout=20000)
            anota(nombre, "lo de Beto le llega a Ana sin recargar", True, f"{(time.perf_counter() - t0) * 1000:.0f} ms · {donde}")
        except Exception:
            anota(nombre, "lo de Beto le llega a Ana sin recargar", False, donde)
        A.foto(str(salida / f"{archivo.stem}-2-editado.png"))
        # Recargar
        A.p.reload(); A.p.wait_for_function(LISTO); A.p.wait_for_timeout(1000)
        anota(nombre, "al recargar sigue lo de los dos", A.p.evaluate(TIENE, "PRUEBA-A") and A.p.evaluate(TIENE, "PRUEBA-B"))
        # Descargar desde la app
        ext = EXT.get(modo, archivo.suffix[1:])
        destino = salida / f"exportado-{archivo.stem}.{ext}"
        with A.p.expect_download(timeout=180000) as dl:
            A.p.evaluate("(x) => NOM.formatos.exportarListo(NOM.editorActual.doc, x, NOM.editorActual.doc.valor('_doc', 'nombre') || 'documento')", ext)
        dl.value.save_as(str(destino))
        po, ps = partes(archivo), partes(destino)
        juntos = b"".join(ps.values())
        anota(nombre, "lo descargado trae lo que escribieron los dos", b"PRUEBA-A" in juntos and b"PRUEBA-B" in juntos)
        iguales = [k for k in po if k in ps and po[k] == ps[k]]
        cambiaron = sorted(k for k in po if k in ps and po[k] != ps[k])
        # Al cambiar celdas se quita la cadena de cálculo (xl/calcChain.xml) y se pide recalcular al abrir:
        # es lo que conviene (Excel la rehace; una cadena vieja lo hace pedir «reparar»). Eso toca también
        # [Content_Types].xml, las relaciones del libro y workbook.xml.
        faltan = sorted(k for k in po if k not in ps and k.lower() not in {x.lower() for x in ps} and k != "xl/calcchain.xml")
        anota(nombre, "las demás partes del archivo salen idénticas", not faltan and len(cambiaron) <= (6 if modo == "HOJA" else 3),
              f"{len(iguales)} de {len(po)} idénticas; cambiaron {cambiaron[:6]}" + (f"; FALTAN {faltan[:6]}" if faltan else ""))
        if (Path.home() / ".dotnet" / ("dotnet.exe" if os.name == "nt" else "dotnet")).exists():
            v = subprocess.run([sys.executable, str(Path(__file__).resolve().parent / "validar_openxml.py"), str(destino), "--corpus", str(archivo.parent)], capture_output=True, text=True)
            anota(nombre, "validador de Microsoft: sin errores nuevos", v.returncode == 0, (v.stdout.strip().splitlines() or [""])[-1][:200] if v.returncode == 0 else v.stdout.strip()[:300])
        # se abre en Office de verdad al final, junto con los demás (un solo viaje a la computadora con Windows)
        para_office.append((destino, archivo, nombre))
        errores = [e for e in A.errores + (B.errores if B else []) if "Content Security" not in e]
        anota(nombre, "sin errores de JavaScript", not errores, " | ".join(errores[:3])[:300])
    except Exception as e:
        anota(nombre, "la prueba terminó sin trabarse", False, f"{type(e).__name__}: {str(e)[:300]}")
        try: A.foto(str(salida / f"{archivo.stem}-error.png"))
        except Exception: pass
    finally:
        if B: B.cerrar()
        A.cerrar()


def main():
    if len(sys.argv) < 3: sys.exit(__doc__)
    salida = Path(sys.argv[1]).resolve(); salida.mkdir(parents=True, exist_ok=True)
    archivos = [Path(x).resolve() for x in sys.argv[2:]]
    if not N.esperar_servidor(): sys.exit("No hay servidor en " + N.B)
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        for f in archivos: probar(pw, f, salida)
    # Office de verdad (Word, Excel o PowerPoint en la computadora con Windows) abre cada descarga sin avisos ni «reparar»
    import office_abre
    res = office_abre.abrir([d for d, _, _ in para_office] + [o for _, o, _ in para_office], "editar-real")
    for d, o, nombre in para_office:
        if not res: anota(nombre, "Office de verdad lo abre sin avisos", False, "no se pudo llegar a la computadora con Windows"); continue
        r_o = res.get(d.name) if d.name != o.name else None
        if office_abre.bien(r_o): anota(nombre, "Office de verdad lo abre sin avisos ni reparar", True)
        elif not office_abre.bien(res.get(o.name)): print(f"  ({nombre}: Office también se queja del original)")
        else: anota(nombre, "Office de verdad lo abre sin avisos ni reparar", False, json.dumps(r_o, ensure_ascii=False)[:300])
    malos = [r for r in resultados if not r[2]]
    print(f"\nArchivos reales: {len(resultados) - len(malos)} de {len(resultados)} bien, en {len(archivos)} archivos")
    for r in malos: print("  MAL", r[0], "·", r[1], "·", r[3][:200])
    json.dump(resultados, open(salida / "resultados.json", "w"), ensure_ascii=False, indent=1)
    sys.exit(1 if malos else 0)


if __name__ == "__main__":
    main()

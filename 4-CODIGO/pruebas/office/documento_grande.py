#!/usr/bin/env python3
"""Un documento de Word grande y real, de punta a punta, por la interfaz (como lo haría una persona).

  python pruebas/office/documento_grande.py ARCHIVO.docx SALIDA [--movil]

Pasos y lo que mide:
  1. Importa el archivo con «Nuevo documento → Importar archivo» (selector de archivos de verdad):
     tiempo de leerlo, de subirlo y de abrirlo; aviso que muestra; páginas, tablas e imágenes pintadas.
  2. Escribe al principio, a la mitad y al final con el teclado: cuánto tarda cada letra en verse.
  3. Otra persona (otro navegador, otra cuenta con permiso de edición) abre el documento: tiempo de
     abrir, que vea lo escrito, y que lo que ella escribe le llegue a la primera.
  4. Recarga la página: lo escrito sigue ahí.
  5. Descarga el Word (menú → Descargar Word) y lo compara con el original: texto completo, párrafos,
     tablas e imágenes; lo pasa por el validador de Microsoft si está (validar_openxml.py).
Sale con código 1 si algo falla. Necesita el servidor de pruebas en NOM_BASE (por omisión el 8099).
"""
import json, os, re, subprocess, sys, time, zipfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "interfaz"))
import nomprueba as N  # noqa: E402

LISTO = "() => !!(window.NOM && NOM.editorActual && NOM.editorActual.doc && NOM.editorActual.vista && NOM.editorActual.vista.isConnected && !NOM.editorActual.doc.pendientes.length)"
resultados = []


def anota(nombre, ok, det=""):
    resultados.append((nombre, ok, det))
    print(("BIEN " if ok else "MAL  ") + nombre + (f" · {det}" if det else ""), flush=True)


def texto_docx(ruta):
    """Texto de cada párrafo (w:t, también los de tablas y cuadros), párrafos de primer nivel, tablas e imágenes."""
    with zipfile.ZipFile(ruta) as z:
        x = z.read("word/document.xml").decode("utf8")
    cuerpo = x.split("<w:body>", 1)[-1]
    parrafos = []
    for p in re.findall(r"<w:p[ >].*?</w:p>", cuerpo, flags=re.S):
        parrafos.append("".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p)))
    # Párrafos de primer nivel: los que no están dentro de una tabla ni de un cuadro de texto.
    sin_anidados = re.sub(r"<w:tbl>.*?</w:tbl>", "", re.sub(r"<mc:AlternateContent>.*?</mc:AlternateContent>|<w:pict>.*?</w:pict>|<w:object.*?</w:object>", "", cuerpo, flags=re.S), flags=re.S)
    primer_nivel = len(re.findall(r"<w:p[ >/]", sin_anidados))
    imagenes = len(re.findall(r"<a:blip ", cuerpo)) + len(re.findall(r"<v:imagedata ", re.sub(r"<mc:Fallback>.*?</mc:Fallback>", "", cuerpo, flags=re.S)))
    return {"parrafos": parrafos, "primer_nivel": primer_nivel, "tablas": cuerpo.count("<w:tbl>"), "imagenes": imagenes,
            "letras": sum(len(p) for p in parrafos)}


class otra_persona(N.Navegador):
    """Segunda persona en Firefox, con el Playwright ya abierto por la primera."""
    def __init__(self, A):
        self.pw = A.pw
        self.b = A.pw.firefox.launch()
        self.ctx = self.b.new_context(viewport={"width": 1280, "height": 860})
        self.p = self.ctx.new_page()
        self.errores = []
        self.p.on("pageerror", lambda e: self.errores.append(str(e)))
        self.p.on("console", lambda m: self.errores.append("console: " + m.text) if m.type == "error" else None)
    def cerrar(self):
        try: self.b.close()
        except Exception: pass


def medir_tecleo(p, texto):
    """Teclea letra por letra y mide cuánto tarda cada una en aparecer en el editor."""
    tiempos = []
    for ch in texto:
        antes = p.evaluate("() => document.querySelector('.wd-cuerpo').textContent.length")
        t0 = time.perf_counter()
        p.keyboard.type(ch)
        p.wait_for_function("(n) => document.querySelector('.wd-cuerpo').textContent.length !== n", arg=antes, timeout=15000)
        tiempos.append((time.perf_counter() - t0) * 1000)
    lentas = [(i, texto[i], round(t)) for i, t in enumerate(tiempos) if t > 400]
    tiempos.sort()
    return {"p50": round(tiempos[len(tiempos) // 2]), "max": round(tiempos[-1]), "n": len(tiempos), "lentas": lentas}


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    archivo, salida = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
    movil = "--movil" in sys.argv
    salida.mkdir(parents=True, exist_ok=True)
    orig = texto_docx(archivo)
    print(f"Original: {len(orig['parrafos'])} párrafos, {orig['letras']} letras, {orig['tablas']} tablas, {orig['imagenes']} imágenes, {archivo.stat().st_size} bytes")
    if not N.esperar_servidor():
        sys.exit("No hay servidor en " + N.B)
    a, b = N.cuenta(), N.cuenta()
    esp = N.espacio(a["token"], "Documento grande")
    N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": b["nombre"], "permiso": "EDICION"}, a["token"])

    A = N.Navegador(movil=movil)
    consola = []   # avisos de la página (p. ej., si el Word no pudo partir del original y se armó desde cero)
    A.p.on("console", lambda m: consola.append(f"{m.type}: {m.text[:400]}") if m.type in ("warning", "error") else None)
    A.p.set_default_timeout(180000)
    A.entrar(a)
    A.p.goto(N.B + "/#/espacio/" + esp)
    A.p.wait_for_timeout(1500)
    # 1. Importar por la interfaz
    A.p.evaluate("(e) => NOM.vistaEspacios.nuevoDocumento({id: e, nombre: 'Documento grande'})", esp)
    A.p.wait_for_selector(".velo [data-importar]")
    t0 = time.perf_counter()
    with A.p.expect_file_chooser() as fc:
        A.p.click(".velo [data-importar]")
    fc.value.set_files(str(archivo))
    try:
        A.p.wait_for_function("() => location.hash.startsWith('#/doc/')", timeout=300000)
    except Exception:
        avisos = A.p.evaluate("() => [...document.querySelectorAll('.aviso')].map((x) => x.textContent.trim())")
        anota("importar por la interfaz abre el documento", False, f"avisos: {avisos} · errores: {A.errores[:3]}")
        A.foto(str(salida / "importar-fallo.png")); A.cerrar(); return terminar()
    t_subido = time.perf_counter() - t0
    A.p.wait_for_function(LISTO, timeout=300000)
    t_abierto = time.perf_counter() - t0
    A.p.wait_for_timeout(3000)
    doc = A.p.evaluate("() => location.hash.slice(6)")
    aviso = A.p.evaluate("() => [...document.querySelectorAll('.avisos .aviso')].map((x) => x.textContent.trim()).join(' | ')")
    info = A.p.evaluate("""() => { const c = document.querySelector('.wd-cuerpo');
      return { paginas: document.querySelectorAll('.wd-pagina').length, parrafos: c.children.length, letras: c.textContent.length,
               tablas: c.querySelectorAll('table').length, imagenes: c.querySelectorAll('img').length, objetos: c.querySelectorAll('[data-obj]').length,
               alto: c.scrollHeight }; }""")
    anota("importar por la interfaz abre el documento", True, f"subido en {t_subido:.1f} s, listo en {t_abierto:.1f} s · aviso: {aviso or '—'}")
    anota("se pintan los párrafos", info["parrafos"] >= orig["primer_nivel"] * 0.97, f"{info['parrafos']} de {orig['primer_nivel']} de primer nivel · " + json.dumps(info))
    anota("se pintan todas las tablas", info["tablas"] >= orig["tablas"], f"{info['tablas']} de {orig['tablas']}")
    anota("el texto completo está en el editor", info["letras"] >= orig["letras"] * 0.97, f"{info['letras']} de {orig['letras']} letras")
    A.foto(str(salida / "1-abierto.png"))
    st, meta = N.api("GET", f"/api/documentos/{doc}", None, a["token"])
    print("Servidor:", st, {k: meta.get(k) for k in ("tamano", "bytes", "version", "nombre") if isinstance(meta, dict) and k in meta})

    # 2. Escribir al principio, a la mitad y al final
    cuerpo = A.p.locator(".wd-cuerpo")
    A.p.evaluate("""() => { const c = document.querySelector('.wd-cuerpo'); const p = c.firstElementChild; const r = document.createRange();
      r.setStart(p, 0); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); c.focus(); }""")
    t_ini = medir_tecleo(A.p, "INICIO-PRUEBA ")
    anota("escribir al principio", t_ini["p50"] < 250 and t_ini["max"] < 1500, f"p50 {t_ini['p50']} ms, máx. {t_ini['max']} ms · lentas {t_ini['lentas']}")
    A.p.evaluate("""() => { const c = document.querySelector('.wd-cuerpo'); const ps = [...c.children].filter((x) => x.textContent.trim().length > 20);
      const p = ps[Math.floor(ps.length / 2)]; p.scrollIntoView({block: 'center'}); const r = document.createRange();
      const t = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode(); r.setStart(t, 0); r.collapse(true);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r); c.focus(); }""")
    t_mit = medir_tecleo(A.p, "MITAD-PRUEBA ")
    anota("escribir a la mitad", t_mit["p50"] < 250 and t_mit["max"] < 1500, f"p50 {t_mit['p50']} ms, máx. {t_mit['max']} ms · lentas {t_mit['lentas']}")
    A.p.keyboard.press("Control+End")
    A.p.wait_for_timeout(500)
    t_fin = medir_tecleo(A.p, " FINAL-PRUEBA")
    anota("escribir al final", t_fin["p50"] < 250 and t_fin["max"] < 1500, f"p50 {t_fin['p50']} ms, máx. {t_fin['max']} ms · lentas {t_fin['lentas']}")
    A.p.wait_for_function(LISTO, timeout=60000)
    A.foto(str(salida / "2-escrito.png"))

    # 3. Otra persona
    B = otra_persona(A)
    B.p.set_default_timeout(180000)
    B.entrar(b)
    t0 = time.perf_counter()
    B.p.goto(N.B + "/#/doc/" + doc)
    B.p.wait_for_function(LISTO, timeout=300000)
    t_b = time.perf_counter() - t0
    vistos = B.p.evaluate("() => ['INICIO-PRUEBA', 'MITAD-PRUEBA', 'FINAL-PRUEBA'].map((x) => document.querySelector('.wd-cuerpo').textContent.includes(x))")
    anota("la otra persona abre el documento y ve lo escrito", all(vistos), f"abrió en {t_b:.1f} s · {vistos}")
    B.p.evaluate("""() => { const c = document.querySelector('.wd-cuerpo'); const p = c.children[3] || c.firstElementChild; const r = document.createRange();
      r.setStart(p, 0); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); c.focus(); }""")
    B.p.keyboard.type("DESDE-B ")
    t0 = time.perf_counter()
    try:
        A.p.wait_for_function("() => document.querySelector('.wd-cuerpo').textContent.includes('DESDE-B')", timeout=30000)
        anota("lo que escribe la otra persona llega a la primera", True, f"{(time.perf_counter() - t0) * 1000:.0f} ms")
    except Exception:
        anota("lo que escribe la otra persona llega a la primera", False, "no llegó en 30 s")
    B.p.wait_for_function(LISTO, timeout=60000)
    A.p.wait_for_function(LISTO, timeout=60000)

    # 4. Recargar
    A.p.reload()
    A.p.wait_for_function(LISTO, timeout=300000)
    A.p.wait_for_timeout(1500)
    sigue = A.p.evaluate("() => ['INICIO-PRUEBA', 'MITAD-PRUEBA', 'FINAL-PRUEBA', 'DESDE-B'].map((x) => document.querySelector('.wd-cuerpo').textContent.includes(x))")
    anota("al recargar sigue todo lo escrito", all(sigue), str(sigue))

    # 5. Descargar Word desde el menú
    destino = salida / ("exportado-" + archivo.stem + ".docx")
    try:
        with A.p.expect_download(timeout=300000) as dl:
            A.p.evaluate("() => NOM.formatos.exportarListo(NOM.editorActual.doc, 'docx', NOM.editorActual.doc.valor('_doc', 'nombre') || 'documento')")
        dl.value.save_as(str(destino))
        anota("descargar Word", destino.stat().st_size > 0, f"{destino.stat().st_size} bytes")
    except Exception as e:
        anota("descargar Word", False, str(e)[:200])
    if destino.exists():
        sal = texto_docx(destino)
        anota("el Word descargado conserva las tablas", sal["tablas"] >= orig["tablas"], f"{sal['tablas']} de {orig['tablas']}")
        anota("el Word descargado conserva las imágenes", sal["imagenes"] >= orig["imagenes"], f"{sal['imagenes']} de {orig['imagenes']}")
        juntos_o = "".join(orig["parrafos"]); juntos_s = "".join(sal["parrafos"])
        for marca in ("INICIO-PRUEBA ", "MITAD-PRUEBA ", " FINAL-PRUEBA", "DESDE-B "):
            juntos_s = juntos_s.replace(marca, "", 1)
        anota("el Word descargado trae todo el texto (sin lo agregado, igual al original)", juntos_o == juntos_s,
              f"{len(juntos_s)} de {len(juntos_o)} letras" + ("" if juntos_o == juntos_s else " · primera diferencia en " + str(next((i for i, (x, y) in enumerate(zip(juntos_o, juntos_s)) if x != y), min(len(juntos_o), len(juntos_s))))
                                                              + " · " + json.dumps(A.p.evaluate("() => NOM.ooxml.ultimoDocx || null")) + " · consola: " + " | ".join(consola[-4:])))
        # Word de verdad (en la computadora con Windows) abre lo descargado sin avisos ni «reparar»
        sys.path.insert(0, str(Path(__file__).resolve().parent)); import office_abre
        r_w = office_abre.abrir([destino], "documento-grande").get(destino.name)
        anota("Word de verdad abre el Word descargado sin avisos ni reparar", office_abre.bien(r_w), json.dumps(r_w, ensure_ascii=False)[:200] if r_w else "sin la computadora con Windows")
        val = Path(__file__).resolve().parent / "validar_openxml.py"
        if (Path.home() / ".dotnet" / ("dotnet.exe" if os.name == "nt" else "dotnet")).exists():
            v = subprocess.run([sys.executable, str(val), str(destino), "--corpus", str(archivo.parent)], capture_output=True, text=True)
            anota("validador de Microsoft: sin errores nuevos", v.returncode == 0, v.stdout.strip().splitlines()[-1] if v.stdout.strip() else v.stderr[-200:])
    errores = [e for e in A.errores + B.errores if "Content Security Policy" not in e]
    anota("sin errores de JavaScript", not errores, " | ".join(errores[:4]))
    B.cerrar(); A.cerrar()
    return terminar()


def terminar():
    malos = [r for r in resultados if not r[1]]
    print(f"\n{len(resultados) - len(malos)} de {len(resultados)} bien")
    return 1 if malos else 0


if __name__ == "__main__":
    sys.exit(main())

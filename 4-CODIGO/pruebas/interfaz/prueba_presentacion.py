"""La pestaña Presentación de PowerPoint en NOM, por la interfaz (como lo haría una persona).

  python pruebas/interfaz/prueba_presentacion.py SALIDA

Ocultar una diapositiva desde su menú (se ve atenuada); F5 empieza desde el principio y se salta la oculta; Mayús+F5
empieza en la actual (aunque esté oculta); agregar una sección y contraerla; una presentación personalizada hecha con su
ventana se presenta en su orden; «repetir hasta Esc» vuelve a la primera; ensayar intervalos guarda cuánto se quedó cada
diapositiva.
"""
import os, sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
import nomprueba as N  # noqa: E402

SALIDA = Path(sys.argv[1] if len(sys.argv) > 1 else Path.home() / ".cache" / "nom-pruebas" / "salida")
res = []


def check(nombre, ok, info=""):
    res.append(ok); print(f"  {'ok ' if ok else 'MAL'}  {nombre}" + (f"   [{str(info)[:300]}]" if not ok and info != "" else ""), flush=True)


def main():
    SALIDA.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    c = N.cuenta(); esp = N.espacio(c["token"], "Presentar")
    doc = N.documento(c["token"], esp, "PRESENTACION", "Presentación de prueba")
    P = N.Navegador(movil=False); P.entrar(c)
    p = P.p
    LISTO = "() => !!(window.NOM && NOM.editorActual && NOM.editorActual.doc && document.querySelector('.mini'))"
    contador = lambda: p.evaluate("() => (document.querySelector('.reproductor [data-c]') || {}).textContent || ''")
    def rueda(*textos):
        """Elige en la ruedita (clic derecho sobre una miniatura la abre), entrando a sus submenús."""
        for t in textos:
            p.wait_for_selector(f'.rueda-item[aria-label="{t}"]', timeout=5000); p.wait_for_timeout(250)
            x, y = p.evaluate("(t) => { const r = document.querySelector(`.rueda-item[aria-label=\"${t}\"]`).getBoundingClientRect(); return [r.x + r.width / 2, r.y + r.height / 2]; }", t)
            p.mouse.click(x, y); p.wait_for_timeout(350)
    def menu(texto):
        p.evaluate("""(t) => { const b = [...document.querySelectorAll('.menu:not(.saliendo) [role=menuitem]')].find((x) => x.textContent.trim().startsWith(t)); b.scrollIntoView({ block: 'nearest' }); b.click(); }""", texto)
        p.wait_for_timeout(350)
    try:
        p.goto(N.B + "/#/doc/" + doc); p.wait_for_function(LISTO, timeout=30000)
        p.evaluate("""() => { const d = NOM.editorActual.doc, ps = ['a1', 'a2', 'a3', 'a4', 'a5'];
          d.lote(() => { d.cambiar('_doc', { paginas: ps }); ps.forEach((pg, k) => d.cambiar('t' + pg, { tipo: 'texto', pagina: pg, x: 80, y: 60, w: 600, h: 80, html: 'Diapositiva ' + (k + 1) })); });
          NOM.editorActual.pagina = 'a1'; return 1; }""")
        p.wait_for_function("() => document.querySelectorAll('.mini').length === 5", timeout=10000)
        p.wait_for_timeout(1500)   # (las miniaturas se repintan un momento al llegar las diapositivas)
        # 1. ocultar la 3 desde su miniatura (clic derecho: la ruedita, «Más» → «Ocultar»)
        p.click('.mini[data-p="a3"]', button="right"); p.wait_for_timeout(250); rueda("Más", "Ocultar")
        p.wait_for_timeout(400)
        check("la diapositiva 3 queda oculta y se ve atenuada", p.evaluate("() => !!document.querySelector('.mini.oculta[data-p=\"a3\"]') && NOM.editorActual.doc.valor('_doc', 'oculta:a3') === true"))
        # 2. F5 desde el principio, saltándose la oculta
        p.click('.mini[data-p="a2"]'); p.wait_for_timeout(200)
        p.keyboard.press("F5"); p.wait_for_selector(".reproductor", timeout=5000); p.wait_for_timeout(400)
        vistos = [contador()]
        for _ in range(3): p.keyboard.press("ArrowRight"); p.wait_for_timeout(350); vistos.append(contador())
        textos = p.evaluate("() => document.querySelector('.reproductor [data-d]').textContent")
        check("F5 empieza en la primera aunque estés en otra (1 / 4) y salta la oculta", vistos == ["1 / 4", "2 / 4", "3 / 4", "4 / 4"] and "Diapositiva 5" in textos, vistos)
        p.keyboard.press("Escape"); p.wait_for_timeout(400)
        # 3. Mayús+F5 desde la actual, aunque esté oculta
        p.click('.mini[data-p="a3"]'); p.wait_for_timeout(200)
        p.keyboard.press("Shift+F5"); p.wait_for_selector(".reproductor", timeout=5000); p.wait_for_timeout(400)
        t3 = p.evaluate("() => document.querySelector('.reproductor [data-d]').textContent")
        check("Mayús+F5 empieza en la diapositiva actual, aunque esté oculta", "Diapositiva 3" in t3, t3)
        p.keyboard.press("Escape"); p.wait_for_timeout(400)
        # 4. agregar una sección en la 3 y contraerla
        p.click('.mini[data-p="a3"]', button="right"); p.wait_for_timeout(250); rueda("Más", "Agregar sección aquí")
        p.wait_for_selector(".velo .hoja input", timeout=5000)
        p.fill(".velo .hoja input", "Resultados"); p.click(".velo .hoja [data-si]"); p.wait_for_timeout(500)
        cabezas = p.evaluate("() => [...document.querySelectorAll('.mini-seccion')].map((x) => x.textContent.trim())")
        check("aparecen las dos secciones (la de antes y «Resultados»)", len(cabezas) == 2 and cabezas[0].startswith("Sección predeterminada") and cabezas[1].startswith("Resultados"), cabezas)
        p.click(".mini-seccion:nth-of-type(2), .mini-seccion >> nth=1"); p.wait_for_timeout(400)
        check("al contraer «Resultados» sus 3 diapositivas no se ven", p.evaluate("() => document.querySelectorAll('.mini').length") == 2, p.evaluate("() => document.querySelectorAll('.mini').length"))
        p.click(".mini-seccion >> nth=1"); p.wait_for_timeout(300)
        # 5. una presentación personalizada con su ventana: la 5 y la 1, en ese orden
        p.click('.ed-herramientas [aria-label="Presentar"]'); p.wait_for_timeout(300); menu("Presentaciones personalizadas")
        p.click(".velo .hoja [data-nueva]"); p.wait_for_selector(".velo .hoja [data-filas] input[type=checkbox]", timeout=5000)
        for pg in ("a5", "a1"):
            p.check(f'.velo .hoja input[type=checkbox][data-p="{pg}"]'); p.wait_for_timeout(150)
        p.fill(".velo .hoja [data-nombre]", "Corta"); p.click(".velo .hoja [data-si]"); p.wait_for_timeout(400)
        pers = p.evaluate("() => NOM.editorActual.doc.valor('_doc', 'personalizadas')")
        check("se guarda la personalizada «Corta» con la 5 y la 1", isinstance(pers, list) and pers and pers[0]["nombre"] == "Corta" and pers[0]["d"] == ["a5", "a1"], pers)
        p.click('.velo .hoja [data-pres="0"]'); p.wait_for_selector(".reproductor", timeout=5000); p.wait_for_timeout(400)
        orden = [p.evaluate("() => document.querySelector('.reproductor [data-d]').textContent")]
        p.keyboard.press("ArrowRight"); p.wait_for_timeout(350); orden.append(p.evaluate("() => document.querySelector('.reproductor [data-d]').textContent"))
        check("la personalizada se presenta en su orden (5 y luego 1)", "Diapositiva 5" in orden[0] and "Diapositiva 1" in orden[1] and contador() == "2 / 2", [o[-14:] for o in orden])
        p.keyboard.press("Escape"); p.wait_for_timeout(400)
        # 6. Configurar presentación: repetir hasta Esc
        p.click('.ed-herramientas [aria-label="Presentar"]'); p.wait_for_timeout(300); menu("Configurar presentación")
        p.check(".velo .hoja [data-rep]"); p.click(".velo .hoja [data-si]"); p.wait_for_timeout(300)
        p.keyboard.press("F5"); p.wait_for_selector(".reproductor", timeout=5000); p.wait_for_timeout(300)
        for _ in range(4): p.keyboard.press("ArrowRight"); p.wait_for_timeout(300)
        check("con «repetir hasta Esc», después de la última vuelve a la primera", contador() == "1 / 4" and p.evaluate("() => !!document.querySelector('.reproductor')"), contador())
        p.keyboard.press("Escape"); p.wait_for_timeout(400)
        p.evaluate("() => NOM.editorActual.doc.cambiar('_doc', { presentacion: null })")
        # 7. ensayar intervalos: ~1.2 s en cada una y se conservan
        p.click('.ed-herramientas [aria-label="Presentar"]'); p.wait_for_timeout(300); menu("Ensayar intervalos")
        p.wait_for_selector(".reproductor .pv-ensayo", timeout=5000)
        for _ in range(4): p.wait_for_timeout(1200); p.keyboard.press("ArrowRight")
        p.wait_for_selector(".velo .hoja [data-si]", timeout=5000); p.click(".velo .hoja [data-si]"); p.wait_for_timeout(500)
        tiempos = p.evaluate("() => ['a1', 'a2', 'a4', 'a5'].map((pg) => (NOM.editorActual.doc.valor('_doc', 'transOffice:' + pg) || {}).advTm)")
        check("ensayar intervalos guarda cuánto se quedó cada una (≈1.2 s) y la oculta no", all(isinstance(t, (int, float)) and 900 <= t <= 2500 for t in tiempos) and p.evaluate("() => (NOM.editorActual.doc.valor('_doc', 'transOffice:a3') || {}).advTm == null"), tiempos)
        check("sin errores de JavaScript", not P.errores, P.errores[:3])
    finally:
        try: p.screenshot(path=str(SALIDA / "presentacion.png"))
        except Exception: pass
        P.cerrar()
    print(f"\nPresentación: {sum(res)} bien, {len(res) - sum(res)} mal")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

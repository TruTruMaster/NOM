#!/usr/bin/env python3
"""Renglones cortados como Word (modo-texto.js, acomodarComoWord): los cortes de pantalla no son parte del documento.

Arma un Word con párrafos justificados y a la izquierda, lo importa y comprueba:
  1. hay cortes de pantalla (<br class="wd-rompe">) y el texto del documento NO tiene saltos de más;
  2. copiar un pedazo que cruza un corte no lleva saltos al portapapeles (ni guiones de pantalla);
  3. escribir a la mitad de un renglón pone la letra donde va y el párrafo se vuelve a acomodar;
  4. lo descargado como .docx no trae saltos de renglón (w:br) en esos párrafos;
  5. en la vista sin páginas no hay cortes de pantalla;
  6. con «Guiones automáticos» (un Word en español con un párrafo en inglés): el guion del final del renglón es de
     pantalla (ni el texto, ni copiar, ni escribir, ni lo descargado lo llevan) y cada párrafo se parte con su idioma.
"""
import base64, io, os, sys, zipfile
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Pt

TEXTO = ("La colaboración en tiempo real permite que varias personas editen el mismo documento sin perder ningún "
         "cambio, incluso cuando la conexión falla. Cada sección del documento incluye tablas, imágenes y referencias "
         "que deben conservarse exactamente igual al abrirlo en otro programa. Según los resultados de la encuesta "
         "aplicada a doscientos cuarenta y ocho estudiantes, la mayoría prefiere trabajar en equipo cuando las "
         "herramientas son sencillas, transparentes y confiables.")

fallas = []
def check(nombre, ok, detalle=""):
    print(("  ok   " if ok else "  MAL  ") + nombre + ("" if ok else f"   [{detalle}]"), flush=True)
    if not ok: fallas.append(nombre)


def docx_prueba():
    d = Document()
    for alin in (WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = d.add_paragraph()
        r = p.add_run(TEXTO); r.font.name = "Calibri"; r.font.size = Pt(11)
        p.alignment = alin
    # como un índice: tabulación derecha con puntos en 6 pulgadas (432 pt desde el margen)
    for t in ("Introducción\t4", "Marco teórico\t12"):
        p = d.add_paragraph()
        r = p.add_run(t); r.font.name = "Calibri"; r.font.size = Pt(11)
        p.paragraph_format.tab_stops.add_tab_stop(Pt(432), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    b = io.BytesIO(); d.save(b); return b.getvalue()


EN = ("Real-time collaboration allows several people to edit the same document without losing any change, even when the "
      "connection fails. Each section of the document includes tables, images and references that must be preserved "
      "exactly when opened in another application.")


def docx_guiones():
    """Español justificado y a la izquierda, y un párrafo con el idioma en inglés; angostos (muchos cortes)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    d = Document()
    d.settings.element.append(OxmlElement("w:autoHyphenation"))
    for texto, alin, lang in ((TEXTO, WD_ALIGN_PARAGRAPH.JUSTIFY, None), (TEXTO, WD_ALIGN_PARAGRAPH.LEFT, None), (EN, WD_ALIGN_PARAGRAPH.LEFT, "en-US")):
        p = d.add_paragraph()
        r = p.add_run(texto); r.font.name = "Calibri"; r.font.size = Pt(11)
        if lang:
            l = OxmlElement("w:lang"); l.set(qn("w:val"), lang); r._element.get_or_add_rPr().append(l)
        p.alignment = alin
        p.paragraph_format.right_indent = Pt(220)
    # el idioma del documento: español
    st = d.styles.element.find(qn("w:docDefaults"))
    rpr = st.find(qn("w:rPrDefault")).find(qn("w:rPr"))
    l = rpr.find(qn("w:lang"))
    if l is None: l = OxmlElement("w:lang"); rpr.append(l)
    l.set(qn("w:val"), "es-MX")
    b = io.BytesIO(); d.save(b); return b.getvalue()


TEXTO_DOM = """(par) => { let t = ''; (function rec(n) { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType === 3) t += c.nodeValue; else if (c.nodeType === 1 && !c.classList.contains('wd-pintado') && !c.classList.contains('wd-guion')) rec(c); } })(par); return t; }"""


def guiones(P, esp):
    p = P.p
    r = p.evaluate("""async ([b64, esp]) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
      const r = await NOM.formatos.importar(new File([u], 'guiones.docx')); const espacio = (await NOM.api.get('/api/espacios')).find((e) => e.id === esp);
      const d = await NOM.binarios.crearEnEspacio(espacio, r.nombre, r.modo, r.ops); return { id: d.id }; }""", [base64.b64encode(docx_guiones()).decode(), esp])
    p.goto(N.B + "/#/doc/" + r["id"])
    p.wait_for_function("() => document.querySelectorAll('.wd-cuerpo .wd-guionv').length >= 3", timeout=30000)
    p.wait_for_timeout(800)
    g = p.evaluate("""() => [...document.querySelectorAll('.wd-cuerpo > p')].slice(0, 3).map((el) => ({ gi: el.dataset.gi || null, n: el.querySelectorAll('.wd-guionv').length,
        bien: [...el.querySelectorAll('.wd-guionv')].every((x) => x.textContent === '-' && x.nextSibling && x.nextSibling.nodeName === 'BR') }))""")
    check(f"guiones automáticos: guiones de pantalla al final del renglón ({[x['n'] for x in g]})", sum(x["n"] for x in g) >= 3 and all(x["bien"] for x in g), g)
    check("cada párrafo con el idioma de su texto (es-MX, es-MX, en-US)", [x["gi"] for x in g] == ["es-MX", "es-MX", "en-US"], [x["gi"] for x in g])
    textos = p.evaluate(f"() => [...document.querySelectorAll('.wd-cuerpo > p')].slice(0, 3).map({TEXTO_DOM})")
    check("…y el texto del documento no los lleva", textos == [TEXTO, TEXTO, EN], [t[:40] for t in textos])
    # copiar a través de un corte con guion
    copiado = p.evaluate("""() => { const g = document.querySelector('.wd-cuerpo .wd-guionv'), par = g.closest('p');
        let antes = g.previousSibling; while (antes && antes.nodeType !== 3) antes = antes.lastChild || antes.previousSibling;
        let despues = g.nextSibling.nextSibling; while (despues && despues.nodeType !== 3) despues = despues.firstChild || despues.nextSibling;
        const r = document.createRange(); r.setStart(antes, Math.max(0, antes.nodeValue.length - 3)); r.setEnd(despues, Math.min(3, despues.nodeValue.length));
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        const dt = new DataTransfer(); const ev = new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true });
        par.dispatchEvent(ev); return { texto: dt.getData('text/plain'), html: dt.getData('text/html'), prevenido: ev.defaultPrevented, esperado: antes.nodeValue.slice(-3) + despues.nodeValue.slice(0, 3) }; }""")
    check("copiar a través de un guion: la palabra entera, sin guion ni salto", copiado["prevenido"] and copiado["texto"] == copiado["esperado"].replace("\u00ad", ""), copiado)
    # escribir junto al guion (al principio del renglón de abajo)
    xy = p.evaluate("""() => { const g = document.querySelector('.wd-cuerpo .wd-guionv'); let d = g.nextSibling.nextSibling; while (d && d.nodeType !== 3) d = d.firstChild || d.nextSibling;
        const r = document.createRange(); r.setStart(d, 0); r.setEnd(d, 1); const q = r.getClientRects()[0]; return [q.left + 0.5, q.top + q.height / 2]; }""")
    p.mouse.click(xy[0], xy[1]); p.wait_for_timeout(200)
    p.keyboard.type("Q"); p.wait_for_timeout(1200)
    t1 = p.evaluate(f"() => ({TEXTO_DOM})(document.querySelector('.wd-cuerpo > p'))")
    todos = p.evaluate(f"() => [...document.querySelectorAll('.wd-cuerpo > p')].slice(0, 3).map({TEXTO_DOM})")
    k = t1.find("Q")
    check("escribir junto al guion: la letra queda en la palabra", k >= 0 and t1.replace("Q", "") == TEXTO, [t1[max(0, k - 20):k + 20] if k >= 0 else None, [x.find("Q") for x in todos], len(t1), len(TEXTO)])
    n = p.evaluate("() => document.querySelector('.wd-cuerpo > p').querySelectorAll('.wd-guionv').length")
    check("…y el párrafo se vuelve a partir con guiones", n >= 1, n)
    b64 = p.evaluate("""async () => { let u = NOM.ooxml.docx(NOM.editorActual.doc); if (u && u.then) u = await u;
        let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }""")
    z = zipfile.ZipFile(io.BytesIO(base64.b64decode(b64)))
    xml = z.read("word/document.xml").decode("utf-8")
    import re
    txt = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml))
    check("lo descargado: sin guiones de pantalla en el texto", TEXTO in txt.replace("Q", "") and EN in txt and txt.replace("Q", "").count("-") == (TEXTO + TEXTO + EN).count("-"), txt[:60])
    check("…y sigue con «Guiones automáticos»", "<w:autoHyphenation" in z.read("word/settings.xml").decode("utf-8"))
    # en la vista sin páginas los parte el navegador (guiones suaves): el cursor dentro de uno no pierde lo escrito
    p.evaluate("() => { const b = document.querySelector('[data-vista=\"continua\"]'); if (b) b.click(); }")
    p.wait_for_function("() => !document.querySelector('.wd-cuerpo br.wd-rompe') && document.querySelector('.wd-cuerpo .wd-guion')", timeout=15000)
    p.evaluate("""() => { const g = document.querySelectorAll('.wd-cuerpo > p')[1].querySelector('.wd-guion'); const t = g.firstChild;
        g.closest('.wd-cuerpo').focus(); const r = document.createRange(); r.setStart(t, 1); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }""")
    p.wait_for_timeout(150)
    p.keyboard.type("W"); p.wait_for_timeout(1200)
    t2 = p.evaluate(f"() => ({TEXTO_DOM})(document.querySelectorAll('.wd-cuerpo > p')[1])")
    check("sin páginas: el cursor dentro de un guion suave no pierde lo escrito", "W" in t2 and t2.replace("W", "") == TEXTO, t2[:80])
    p.evaluate("() => { const b = document.querySelector('[data-vista=\"paginas\"]'); if (b) b.click(); }")


def elegir(p, texto):
    """Toca la opción del menú abierto que empieza con ese texto (desplazándolo si hace falta, como una persona)."""
    xy = p.evaluate("""(t) => { const b = [...document.querySelectorAll('.menu:not(.saliendo) [role=menuitem]')].find((x) => x.textContent.trim().startsWith(t));
        if (!b) return null; b.scrollIntoView({ block: 'nearest' }); const r = b.getBoundingClientRect(); return [r.x + r.width / 2, r.y + r.height / 2]; }""", texto)
    if not xy: raise AssertionError("no está en el menú: " + texto)
    p.mouse.click(xy[0], xy[1]); p.wait_for_timeout(350)


def numeros_linea(P):
    """Números de línea desde «Más» → «Números de línea» (un menú aparte, como el de Word)."""
    p = P.p
    p.click('.ed-herramientas [aria-label="Más"]'); p.wait_for_timeout(300)
    caja = p.evaluate("() => { const m = document.querySelector('.menu:not(.saliendo)'); const r = m.getBoundingClientRect(); return { arriba: r.top, abajo: r.bottom, alto: innerHeight, desplaza: m.scrollHeight > m.clientHeight }; }")
    check("el menú «Más» cabe en la pantalla (si es largo, se desplaza)", caja["arriba"] >= 0 and caja["abajo"] <= caja["alto"], caja)
    elegir(p, "Números de línea")
    titulo = p.evaluate("() => (document.querySelector('.menu:not(.saliendo) .menu-titulo') || {}).textContent || ''")
    check("«Números de línea» abre su propio menú", titulo == "Números de línea", titulo)
    elegir(p, "Continuos")
    p.wait_for_function("() => document.querySelectorAll('.wd-numlin').length >= 5", timeout=15000)
    nums = p.evaluate("() => [...document.querySelectorAll('.wd-numlin')].map((x) => +x.textContent)")
    check(f"continuos: se ven los números ({len(nums)}) y siguen 1, 2, 3…", nums[:5] == [1, 2, 3, 4, 5] and all(b == a + 1 for a, b in zip(nums, nums[1:])), nums[:12])
    p.click('.ed-herramientas [aria-label="Más"]'); p.wait_for_timeout(300)
    elegir(p, "Números de línea")
    marcado = p.evaluate("""() => { const b = [...document.querySelectorAll('.menu:not(.saliendo) [role=menuitem]')].find((x) => x.textContent.trim().startsWith('Continuos')); return !!(b && b.querySelector('svg.ico-s')); }""")
    check("…y en el menú «Continuos» queda marcado", marcado)
    elegir(p, "Ninguno")
    p.wait_for_function("() => !document.querySelector('.wd-numlin')", timeout=15000)
    check("«Ninguno» los quita", p.evaluate("() => document.querySelectorAll('.wd-numlin').length") == 0)


def dos_lados(P, esp):
    """Texto a los dos lados de un dibujo (w17, hecho en Word): cada renglón de la franja en dos tramos con un hueco."""
    p = P.p
    arch = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real/archivos/w17-dos-lados.docx"
    if not arch.exists():
        print("  (sin w17-dos-lados.docx: se salta lo de los dos lados)"); return
    r = p.evaluate("""async ([b64, esp]) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
      const r = await NOM.formatos.importar(new File([u], 'w17.docx')); const espacio = (await NOM.api.get('/api/espacios')).find((e) => e.id === esp);
      const d = await NOM.binarios.crearEnEspacio(espacio, r.nombre, r.modo, r.ops); return { id: d.id }; }""", [base64.b64encode(arch.read_bytes()).decode(), esp])
    p.goto(N.B + "/#/doc/" + r["id"])
    p.wait_for_function("() => document.querySelectorAll('.wd-cuerpo .wd-hueco').length >= 8", timeout=40000)
    p.wait_for_timeout(800)
    i = p.evaluate("() => [...document.querySelectorAll('.wd-cuerpo > p')].findIndex((x) => x.querySelector('.wd-hueco') && x.textContent.startsWith('Para calcular'))")
    check("a los dos lados: el párrafo junto al dibujo lleva sus tramos (con hueco)", i >= 0, i)
    if i < 0: return
    antes = p.evaluate(f"() => ({TEXTO_DOM})(document.querySelectorAll('.wd-cuerpo > p')[{i}])")
    check("…y el texto del documento no cambia (el hueco es de pantalla)", "Para calcular el promedio" in antes and antes == antes.replace("  ", " "), antes[:60])
    # escribir justo antes del primer hueco (al final del tramo de la izquierda)
    p.evaluate(f"() => document.querySelectorAll('.wd-cuerpo > p')[{i}].scrollIntoView({{ block: 'center' }})"); p.wait_for_timeout(400)
    xy = p.evaluate(f"""() => {{ const par = document.querySelectorAll('.wd-cuerpo > p')[{i}], h = par.querySelector('.wd-hueco'); let t = h.previousSibling;
        while (t && t.nodeType !== 3) t = t.lastChild || t.previousSibling;
        const k = t.nodeValue.replace(/\s+$/, '').length; const r = document.createRange(); r.setStart(t, k - 1); r.setEnd(t, k); const q = r.getClientRects()[0]; return [q.right - 0.5, q.top + q.height / 2]; }}""")
    p.mouse.click(xy[0], xy[1]); p.wait_for_timeout(200)
    p.keyboard.type("Q"); p.wait_for_timeout(1500)
    despues = p.evaluate(f"() => ({TEXTO_DOM})(document.querySelectorAll('.wd-cuerpo > p')[{i}])")
    check("escribir al final del tramo de la izquierda: la letra queda en su lugar", "ParaQ" in despues.replace(" ", "")[:12] and despues.replace("Q", "") == antes, despues[:40])
    check("…y el párrafo se vuelve a partir en tramos", p.evaluate(f"() => document.querySelectorAll('.wd-cuerpo > p')[{i}].querySelectorAll('.wd-hueco').length") >= 2)
    # copiar a través de un hueco: sin el hueco ni saltos
    cp = p.evaluate(f"""() => {{ const par = document.querySelectorAll('.wd-cuerpo > p')[{i}], h = par.querySelectorAll('.wd-hueco')[1];
        let a = h.previousSibling; while (a && a.nodeType !== 3) a = a.lastChild || a.previousSibling;
        let b = h.nextSibling; while (b && b.nodeType !== 3) b = b.firstChild || b.nextSibling;
        const r = document.createRange(); r.setStart(a, Math.max(0, a.nodeValue.length - 3)); r.setEnd(b, Math.min(4, b.nodeValue.length));
        const s = getSelection(); s.removeAllRanges(); s.addRange(r);
        const dt = new DataTransfer(); const ev = new ClipboardEvent('copy', {{ clipboardData: dt, bubbles: true, cancelable: true }});
        par.dispatchEvent(ev); return {{ texto: dt.getData('text/plain'), esperado: a.nodeValue.slice(-3) + b.nodeValue.slice(0, 4) }}; }}""")
    check("copiar a través de un hueco: el texto seguido, sin el hueco", cp["texto"] == cp["esperado"], cp)
    b64 = p.evaluate("""async () => { let u = NOM.ooxml.docx(NOM.editorActual.doc); if (u && u.then) u = await u;
        let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }""")
    xml = zipfile.ZipFile(io.BytesIO(base64.b64decode(b64))).read("word/document.xml").decode("utf-8")
    import re
    txt = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml))
    check("lo descargado trae lo escrito y nada del hueco", "ParaQ" in txt.replace(" ", "") and "Para" in txt, txt[:40])


def main():
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Acomodo")
    P = N.Navegador(movil=False); P.entrar(a)
    p = P.p
    try:
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(1200)
        r = p.evaluate("""async ([b64, esp]) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
          const r = await NOM.formatos.importar(new File([u], 'acomodo.docx')); const espacio = (await NOM.api.get('/api/espacios')).find((e) => e.id === esp);
          const d = await NOM.binarios.crearEnEspacio(espacio, r.nombre, r.modo, r.ops); return { id: d.id }; }""", [base64.b64encode(docx_prueba()).decode(), esp])
        p.evaluate("() => { const t = NOM.leer('nom.texto', {}); t.vista = 'paginas'; localStorage.setItem('nom.texto', JSON.stringify(t)); }")
        p.goto(N.B + "/#/doc/" + r["id"])
        p.wait_for_function("() => document.querySelectorAll('.wd-cuerpo br.wd-rompe').length > 0", timeout=30000)
        p.wait_for_timeout(800)
        # 1. cortes de pantalla y texto del documento sin saltos de más
        n = p.evaluate("() => document.querySelectorAll('.wd-cuerpo br.wd-rompe').length")
        check(f"hay cortes de renglón de pantalla ({n})", n >= 6, n)
        modelo = p.evaluate("""() => { const el = document.querySelector('.wd-cuerpo'); const out = [];
            for (const par of el.children) { let t = ''; (function rec(n) { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType === 3) t += c.nodeValue; else if (c.nodeType === 1 && !c.classList.contains('wd-pintado') && !c.classList.contains('wd-guion')) rec(c); } })(par); out.push(t); }
            return out; }""")
        check("el texto de cada párrafo es el del documento (los cortes no agregan nada)", modelo[:3] == [TEXTO] * 3, [m[:40] for m in modelo[:3]])
        # tabulación derecha con puntos: el número termina en la parada y los puntos no son texto del documento
        tab = p.evaluate("""() => { const ps = [...document.querySelectorAll('.wd-cuerpo > p')].filter((x) => /Introducción|Marco teórico/.test(x.textContent));
            return ps.map((el) => { const sp = el.querySelector('.wd-tabw'); const r = el.getBoundingClientRect(), z = r.width / el.offsetWidth;
              const tn = [...el.childNodes].concat([...el.querySelectorAll('span')].flatMap((s) => [...s.childNodes])).filter((n) => n.nodeType === 3 && /\\d/.test(n.nodeValue)).pop();
              const rg = document.createRange(); rg.selectNodeContents(tn); const q = rg.getBoundingClientRect();
              return { spacer: !!sp, puntos: sp ? sp.textContent.length : 0, finNumero: (q.right - r.left) / z * 0.75 }; }); }""")
        check("tabulación derecha con puntos: hay relleno de puntos", all(t["spacer"] and t["puntos"] > 20 for t in tab), tab)
        check("…y el número termina en la parada (432 pt)", all(abs(t["finNumero"] - 432) < 1.5 for t in tab), [round(t["finNumero"], 1) for t in tab])
        # 2. copiar cruzando un corte
        copiado = p.evaluate("""() => { const par = document.querySelector('.wd-cuerpo > p'); const br = par.querySelector('br.wd-rompe');
            const antes = br.previousSibling, despues = br.nextSibling;
            const r = document.createRange(); r.setStart(antes, Math.max(0, antes.nodeValue.length - 8)); r.setEnd(despues, Math.min(8, despues.nodeValue.length));
            const s = getSelection(); s.removeAllRanges(); s.addRange(r);
            const dt = new DataTransfer(); const ev = new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true });
            par.dispatchEvent(ev); return { texto: dt.getData('text/plain'), html: dt.getData('text/html'), prevenido: ev.defaultPrevented }; }""")
        check("copiar a través de un corte: sin salto de renglón en el texto", copiado["prevenido"] and "\n" not in copiado["texto"] and len(copiado["texto"]) >= 14, copiado)
        check("…ni en el HTML (sin el corte de pantalla)", "wd-rompe" not in copiado["html"] and "<br" not in copiado["html"], copiado["html"][:120])
        # 3. escribir a la mitad de un renglón
        xy = p.evaluate("""() => { const par = document.querySelector('.wd-cuerpo > p'); const tn = par.querySelector('br.wd-rompe').nextSibling;
            const r = document.createRange(); r.setStart(tn, 10); r.setEnd(tn, 11); const q = r.getClientRects()[0]; return [q.left + 1, q.top + q.height / 2]; }""")
        p.mouse.click(xy[0], xy[1]); p.wait_for_timeout(200)
        p.keyboard.type("ZZ"); p.wait_for_timeout(900)
        t0 = p.evaluate("""() => { const par = document.querySelector('.wd-cuerpo > p'); let t = ''; (function rec(n) { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType === 3) t += c.nodeValue; else if (c.nodeType === 1 && !c.classList.contains('wd-pintado')) rec(c); } })(par); return t; }""")
        check("escribir a la mitad de un renglón: la letra queda dentro del texto", "ZZ" in t0 and t0.replace("ZZ", "") == TEXTO, t0[:80])
        n2 = p.evaluate("() => document.querySelector('.wd-cuerpo > p').querySelectorAll('br.wd-rompe').length")
        check("…y el párrafo se vuelve a acomodar", n2 >= 2, n2)
        # 4. descargar como Word
        b64 = p.evaluate("""async () => { let u = NOM.ooxml.docx(NOM.editorActual.doc); if (u && u.then) u = await u;
            let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }""")
        xml = zipfile.ZipFile(io.BytesIO(base64.b64decode(b64))).read("word/document.xml").decode("utf-8")
        check("lo descargado no trae saltos de renglón (w:br) en esos párrafos", "<w:br/>" not in xml and "<w:br " not in xml, xml.count("<w:br"))
        # 5. vista sin páginas
        p.evaluate("() => { const b = document.querySelector('[data-vista=\"continua\"]'); if (b) b.click(); }"); p.wait_for_timeout(1500)
        n3 = p.evaluate("() => document.querySelectorAll('.wd-cuerpo br.wd-rompe').length")
        check("en la vista sin páginas no hay cortes de pantalla", n3 == 0, n3)
        # 6. guiones automáticos
        p.evaluate("() => { const t = NOM.leer('nom.texto', {}); t.vista = 'paginas'; localStorage.setItem('nom.texto', JSON.stringify(t)); }")
        guiones(P, esp)
        numeros_linea(P)
        dos_lados(P, esp)
        errores = P.errores if hasattr(P, "errores") else []
        check("sin errores de JavaScript", not errores, errores[:3])
    finally:
        P.cerrar()
    print(f"\nAcomodo de renglones como Word: {'todo bien' if not fallas else str(len(fallas)) + ' mal'}")
    return 1 if fallas else 0


if __name__ == "__main__":
    sys.exit(main())

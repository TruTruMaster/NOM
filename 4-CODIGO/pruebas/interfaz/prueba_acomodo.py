#!/usr/bin/env python3
"""Renglones cortados como Word (modo-texto.js, acomodarComoWord): los cortes de pantalla no son parte del documento.

Arma un Word con párrafos justificados y a la izquierda, lo importa y comprueba:
  1. hay cortes de pantalla (<br class="wd-rompe">) y el texto del documento NO tiene saltos de más;
  2. copiar un pedazo que cruza un corte no lleva saltos al portapapeles (ni guiones de pantalla);
  3. escribir a la mitad de un renglón pone la letra donde va y el párrafo se vuelve a acomodar;
  4. lo descargado como .docx no trae saltos de renglón (w:br) en esos párrafos;
  5. en la vista sin páginas no hay cortes de pantalla.
"""
import base64, io, os, sys, zipfile
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
        errores = P.errores if hasattr(P, "errores") else []
        check("sin errores de JavaScript", not errores, errores[:3])
    finally:
        P.cerrar()
    print(f"\nAcomodo de renglones como Word: {'todo bien' if not fallas else str(len(fallas)) + ' mal'}")
    return 1 if fallas else 0


if __name__ == "__main__":
    sys.exit(main())

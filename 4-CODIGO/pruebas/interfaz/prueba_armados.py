#!/usr/bin/env python3
"""Documentos armados a propósito al descargarlos (Markdown, HTML, SVG) y al importar un CSV enorme, en la página real.

Un .docx con una tabla se importa y a su tabla se le ponen valores que no salen de Word (combinada gigante, 100 millones
de filas, «colores» con CSS y comillas); también una imagen de otro sitio y medidas con comillas en un diagrama.
Revisa: que nada se pida a otro sitio, que no se trabe, que lo descargado no lleve estilos ni atributos de más.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_armados.py
"""
import base64, io, os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright
from docx import Document

d = Document(); d.add_paragraph("Antes de la tabla"); t = d.add_table(rows=2, cols=2); t.cell(0, 0).text = "A1"; t.cell(1, 1).text = "B2"
o = io.BytesIO(); d.save(o); B64 = base64.b64encode(o.getvalue()).decode()
c = N.cuenta()
with sync_playwright() as pw:
    b = pw.chromium.launch(); p = b.new_context(viewport={"width": 1280, "height": 800}).new_page(); P = Pagina(p, N.B); P.entrar(c)
    fuera = []; p.on("request", lambda r: fuera.append(r.url) if "ejemplo.invalid" in r.url else None)
    # 1) texto plano de HTML con una imagen de otro sitio
    p.evaluate("() => NOM.formatos.textoPlano('<p>hola<img src=\"https://ejemplo.invalid/pixel.png\"></p>')"); p.wait_for_timeout(1500)
    check("sacar el texto de un HTML no pide sus imágenes a otro sitio", not fuera, fuera)
    # 2) tabla armada dentro de un texto → HTML y Markdown
    t0 = time.time()
    r = p.evaluate("""async (b64) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
      const imp = await NOM.formatos.importar(new File([u], 'a.docx'));
      const ops = imp.ops.map((o) => o.campos && o.campos.tipo === 'tabla' ? Object.assign({}, o, { campos: Object.assign({}, o.campos, {
        'm:A1': '99999x99999', filasTexto: 100000000, colsTexto: 100000,
        'r:A1': 'red;position:fixed;inset:0;z-index:9999', 'k:B2': 'red" onmouseover="alert(1)' }) }) : o);
      const doc = NOM.formatos.docDeOps('a', 'TEXTO', ops);
      const t0 = performance.now(); let h, md;
      try { h = NOM.formatos.aHtml(doc, 'a'); md = NOM.formatos.aMarkdown(doc); } catch (e) { return { ms: performance.now() - t0, error: String(e), tabla: false, fixed: false, onmouse: false }; }
      return { ms: performance.now() - t0, largo: h.length, fixed: h.includes('position:fixed'), onmouse: /onmouseover="/.test(h), md: md.length, tabla: h.includes('<table') };
    }""", B64)
    check("descargar en HTML un texto con una tabla armada no se traba ni truena (< 5 s)", not r.get("error") and r["ms"] < 5000 and time.time() - t0 < 60, r)
    check("…la tabla sigue saliendo", r["tabla"], r)
    check("…sin estilos colados («position:fixed») ni atributos de más («onmouseover=»)", not r["fixed"] and not r["onmouse"], r)
    # 3) diagrama con una imagen de otro sitio y medidas con comillas → SVG y HTML
    r = p.evaluate("""() => { const ops = [{ id: 'x:1', bloque: 'im1', t: 2, n: 'x', campos: { tipo: 'imagen', pagina: 'p1', src: 'https://ejemplo.invalid/a.png', x: '0" onload="alert(1)', y: 0, w: 100, h: 100 } },
        { id: 'x:2', bloque: 'im2', t: 2, n: 'x', campos: { tipo: 'imagen', pagina: 'p1', src: 'data:image/png;base64,iVBORw0KGgo=', x: 10, y: 10, w: 50, h: 50 } }];
      const doc = NOM.formatos.docDeOps('d', 'DIAGRAMA', ops);
      let s = '', h = ''; try { s = NOM.formatos.aSvg(doc, 'p1') || ''; h = NOM.formatos.aHtml(doc, 'd'); } catch (e) { return { error: String(e), externa: true, onload: true, propia: false }; }
      return { externa: s.includes('ejemplo.invalid') || h.includes('ejemplo.invalid'), onload: /onload=/.test(s), propia: s.includes('data:image/png') }; }""")
    check("el .svg y el .html descargados no llevan imágenes de otro sitio", not r["externa"], r)
    check("…ni medidas que se vuelvan atributos («onload=»)", not r["onload"], r)
    check("…y la imagen propia sí sale", r["propia"], r)
    # 4) CSV de 600 mil filas: se importa o da un mensaje, nunca «Maximum call stack»
    r = p.evaluate("""async () => { const t = 'a\\n'.repeat(600000); try { const x = await NOM.formatos.importar(new File([t], 'g.csv')); return { ok: true, modo: x.modo }; } catch (e) { return { ok: false, msg: String(e && e.message) }; } }""")
    check("importar un CSV de 600 mil filas no truena por la pila", r.get("ok") or "call stack" not in r.get("msg", ""), r)
    check("sin errores de JavaScript", not P.errores, P.errores)
    b.close()
print(f"\nArmados: {BIEN[0]} bien, {MAL[0]} mal")
sys.exit(1 if MAL[0] else 0)

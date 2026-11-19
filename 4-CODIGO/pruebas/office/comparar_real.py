#!/usr/bin/env python3
"""Compara cómo se ve un archivo en NOM contra cómo lo pinta Microsoft Office de verdad.

La referencia sale de Office en Windows (`office_real.py` + `windows/verdad.py`): por cada renglón de
cada página, su texto exacto y su posición. Aquí se importa el mismo archivo en NOM (Chromium
con Playwright), se mide cada renglón pintado y se compara:

  - renglones idénticos (mismo texto de principio a fin = Word cortó en el mismo lugar)
  - renglones en la misma página y a qué distancia vertical quedaron
  - número de páginas
  - imágenes lado a lado por página: Office | NOM | diferencias (en rojo lo que solo está en
    Office, en azul lo que solo está en NOM)

  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/comparar_real.py ARCHIVO.docx … [--movil] [--salida DIR]

Los archivos y la referencia se buscan en ~/.cache/nom-pruebas/office-real/{archivos,verdad}.
Necesita el servidor local (puerto 8099) y PLAYWRIGHT_BROWSERS_PATH=~/.cache/nom-pruebas/navegadores.
"""
import argparse, base64, difflib, json, os, re, subprocess, sys, unicodedata
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent)); sys.path.insert(0, str(Path(__file__).parent.parent / "interfaz"))
import nomprueba as N
import prueba_ui as U
from PIL import Image, ImageChops, ImageDraw

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"

JS_IMPORTAR = r"""
async ([b64, nombre, espacio]) => {
  const bin = atob(b64), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const r = await NOM.formatos.importar(new File([bytes], nombre));
  const d = await NOM.api.post(`/api/espacios/${espacio}/documentos`, { nombre: r.nombre, modo: r.modo, ops: NOM.DocCRDT.partir(r.ops) });
  location.hash = "#/doc/" + d.id;
  return { id: d.id, modo: r.modo, aviso: r.aviso || null };
}
"""

# Cada renglón pintado en el editor de texto: se mide letra por letra (rectángulo de cada carácter) y
# se parte donde la letra baja de renglón. Devuelve las páginas con sus renglones (en puntos, desde la
# esquina de la hoja) y el texto del encabezado y del pie.
JS_RENGLONES = r"""
() => {
  const cuerpo = document.querySelector('.wd-cuerpo');
  const pags = [...document.querySelectorAll('.wd-pagina')].map((p) => p.getBoundingClientRect());
  const pagDe = (y) => { for (let i = 0; i < pags.length; i++) if (y >= pags[i].top - 2 && y < pags[i].bottom + 2) return i; return -1; };
  const PT = 0.75 / (window.__zoomNom || 1);
  const letras = [];
  const r = document.createRange();
  // el texto del cuerpo y, con secciones en columnas, lo que sigue de un párrafo partido (otra capa: va justo
  // después de su párrafo); lo recortado de un párrafo partido no se ve y no cuenta
  const kidsC = [...cuerpo.children];
  for (const raiz of [cuerpo, ...document.querySelectorAll('.wd-continuaciones')]) {
    const tw = document.createTreeWalker(raiz, NodeFilter.SHOW_TEXT, { acceptNode: (n) => {
      const e = n.parentElement;
      // (las marcas del editor de un salto de página o de columna no son texto del documento)
      // (el guion de pantalla de la división automática: en la verdad de Word no es una letra del renglón)
      if (!e || e.closest('.wd-corte, [data-obj] .wd-ctl, .wd-salto, .wd-saltocol, .wd-guion, .wd-pintado, script, style')) return NodeFilter.FILTER_REJECT;
      const cs = getComputedStyle(e); if (cs.display === 'none' || cs.visibility === 'hidden') return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT; } });
    let n, parrafo = null, numPar = -1, corte = null, pr = null, zz = 1;
    while ((n = tw.nextNode())) {
      const p = raiz === cuerpo ? n.parentElement.closest('.wd-cuerpo > *') : n.parentElement.closest('.wd-continuaciones > *');
      if (p !== parrafo) {
        parrafo = p; numPar = raiz === cuerpo ? kidsC.indexOf(p) : +(p && p.dataset.par) + 0.5;
        corte = p && p.dataset.clip != null ? +p.dataset.clip : null; pr = p ? p.getBoundingClientRect() : null; zz = p && pr ? pr.height / (p.offsetHeight || 1) || 1 : 1;
      }
      // la llamada de una nota (su número volado): en Word es el carácter de control 2, que la verdad lee como un espacio;
      // aquí cuenta igual (un solo carácter de control en el lugar de su primera letra)
      const nota = n.parentElement.closest('.wd-nota[data-obj]');
      const s = nota ? (n === (nota.__primerTexto = nota.__primerTexto || n) ? '\x02' + ''.padEnd(n.data.length - 1, '\x01') : ''.padEnd(n.data.length, '\x01')) : n.data;
      // lo que se VE: «Mayúsculas» de Word es text-transform (el texto del documento no cambia), como lo lee Word
      const tt = getComputedStyle(n.parentElement).textTransform;
      const ver = (c) => (tt === 'uppercase' ? c.toUpperCase() : tt === 'lowercase' ? c.toLowerCase() : c);
      for (let i = 0; i < s.length; i++) {
        let j = i + 1; const c0 = s.charCodeAt(i); if (c0 >= 0xD800 && c0 < 0xDC00 && j < s.length) j++;
        r.setStart(n, i); r.setEnd(n, j);
        const q = r.getClientRects(); if (!q.length) { i = j - 1; continue; }
        const b = q[q.length - 1];
        if (corte != null && pr && (b.top + b.bottom) / 2 - pr.top > corte * zz) { i = j - 1; continue; }
        if (s[i] === '\x01') { i = j - 1; continue; }   // (el resto del número de una nota: no cuenta)
        letras.push({ c: ver(s.slice(i, j)), x: b.left, y: b.top, h: b.height, w: b.width, p: numPar });
        i = j - 1;
      }
    }
  }
  letras.sort((a, b) => a.p - b.p);
  // renglones: una letra empieza renglón si queda toda debajo de lo que lleva el renglón, o si baja (más de media altura)
  // y regresa a la izquierda (un subíndice después de un superíndice baja, pero sigue hacia la derecha: mismo renglón)
  const renglones = [];
  let cur = null;
  for (const L of letras) {
    const baja = cur && (L.y > cur.y + Math.max(4, cur.h * 0.5) || L.y + L.h < cur.y);
    const nuevo = !cur || L.p !== cur.p || L.y >= cur.abajo - 1 || L.y + L.h <= cur.arriba + 1 || (baja && L.x < cur.ultX - 1);
    if (nuevo) { cur = { texto: '', x: L.x, y: L.y, h: L.h, p: L.p, fin: L.x + L.w, arriba: L.y, abajo: L.y + L.h, ultX: L.x }; renglones.push(cur); }
    cur.texto += L.c; cur.fin = Math.max(cur.fin, L.x + L.w); cur.x = Math.min(cur.x, L.x);
    cur.arriba = Math.min(cur.arriba, L.y); cur.abajo = Math.max(cur.abajo, L.y + L.h); cur.ultX = L.x;
  }
  const salida = pags.map((p, i) => ({ n: i + 1, ancho: p.width * PT, alto: p.height * PT, renglones: [], enc: '', pie: '' }));
  const fuera = [];
  for (const R of renglones) {
    const i = pagDe(R.y + R.h / 2);
    const o = { texto: R.texto, x: (R.x - (i >= 0 ? pags[i].left : 0)) * PT, y: (R.y - (i >= 0 ? pags[i].top : 0)) * PT, w: (R.fin - R.x) * PT, h: R.h * PT, p: R.p };
    if (i >= 0) salida[i].renglones.push(o); else fuera.push(o);
  }
  document.querySelectorAll('.wd-pagina').forEach((p, i) => {
    const e = p.querySelector('.wd-enc'), f = p.querySelector('.wd-pie');
    salida[i].enc = e ? e.innerText : ''; salida[i].pie = f ? f.innerText : '';
  });
  return { paginas: salida, fuera, letras: letras.length };
}
"""


def limpio(t):
    """Texto de un renglón para comparar: sin marcas de control de Word, espacios normalizados."""
    if t is None: return ""
    t = unicodedata.normalize("NFC", t)
    t = t.replace(" ", " ").replace(" ", " ").replace(" ", " ").replace(" ", " ").replace("​", "")
    t = re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def renglones_word(verdad):
    """Renglones del cuerpo de Word por página (sin encabezado, pie ni notas)."""
    out = []
    for pg in verdad.get("paginas", []):
        rs = []
        for z in pg["zonas"]:
            if z.get("historia") not in ("cuerpo", None) or z.get("tipo") not in (0, None): continue
            for li in z["lineas"]:
                if li.get("t") == 0 and li.get("texto") is not None:
                    rs.append({"texto": li["texto"], "x": li["x"], "y": li["y"], "h": li["h"]})
        rs.sort(key=lambda r: (round(r["y"]), r["x"]))
        out.append(rs)
    return out


def columnas_de(rs):
    """Dónde empiezan las columnas de texto de una página (x en pt), si la página está en columnas: grupos de al menos
    4 renglones que empiezan a más de 60 pt uno del otro y que comparten altura en la página."""
    if len(rs) < 8: return None
    xs = sorted(r["x"] for r in rs if limpio(r["texto"]))
    grupos = []
    for x in xs:
        if grupos and x - grupos[-1][-1] <= 60: grupos[-1].append(x)
        else: grupos.append([x])
    grupos = [g for g in grupos if len(g) >= 4]
    if len(grupos) < 2: return None
    inicios = [min(g) for g in grupos]
    col = lambda r: max([k for k, x0 in enumerate(inicios) if r["x"] >= x0 - 8] or [0])
    rangos = [[r["y"] for r in rs if col(r) == k] for k in range(len(inicios))]
    if not all(rangos) or min(max(y) for y in rangos) < max(min(y) for y in rangos): return None   # no comparten altura: no son columnas
    return inicios


def ordenar_columnas(word, nom):
    """En las páginas en columnas, los renglones de los dos en orden de lectura (columna por columna, cada una de arriba
    a abajo): Word los da de arriba a abajo mezclando las columnas."""
    for k, rs in enumerate(word):
        inicios = columnas_de(rs)
        if not inicios: continue
        col = lambda r: max([j for j, x0 in enumerate(inicios) if r["x"] >= x0 - 8] or [0])
        rs.sort(key=lambda r: (col(r), round(r["y"]), r["x"]))
        if k < len(nom): nom[k].sort(key=lambda r: (col(r), round(r["y"]), r["x"]))


def comparar_renglones(word, nom):
    """Alinea los renglones de los dos (en orden) y cuenta cuántos quedaron idénticos y en su página."""
    W = [(p, limpio(r["texto"]), r) for p, rs in enumerate(word) for r in rs if limpio(r["texto"])]
    Nn = [(p, limpio(r["texto"]), r) for p, rs in enumerate(nom) for r in rs if limpio(r["texto"])]
    sm = difflib.SequenceMatcher(None, [t for _, t, _ in W], [t for _, t, _ in Nn], autojunk=False)
    iguales = misma_pag = 0; dy = []
    diferencias = []
    for op, a0, a1, b0, b1 in sm.get_opcodes():
        if op == "equal":
            for k in range(a1 - a0):
                iguales += 1
                pw, _, rw = W[a0 + k]; pn, _, rn = Nn[b0 + k]
                if pw == pn: misma_pag += 1; dy.append(rn["y"] - rw["y"])
        else:
            diferencias.append({"word": [f"p{W[i][0] + 1}: {W[i][1]}" for i in range(a0, a1)][:12],
                                "nom": [f"p{Nn[i][0] + 1}: {Nn[i][1]}" for i in range(b0, b1)][:12]})
    dy.sort()
    return {"renglones_word": len(W), "renglones_nom": len(Nn), "iguales": iguales, "misma_pagina": misma_pag,
            "pct_iguales": round(100 * iguales / max(1, len(W)), 1), "pct_misma_pagina": round(100 * misma_pag / max(1, len(W)), 1),
            "dy_mediana_pt": round(dy[len(dy) // 2], 1) if dy else None,
            "dy_max_pt": round(max(abs(x) for x in dy), 1) if dy else None, "diferencias": diferencias}


def pdf_a_png(pdf, carpeta, dpi=96):
    carpeta.mkdir(parents=True, exist_ok=True)
    for v in carpeta.glob("office-*.png"): v.unlink()
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(carpeta / "office")], check=True)
    return sorted(carpeta.glob("office-*.png"))


def lado_a_lado(a, b, ruta, titulo=""):
    """Office | NOM | diferencias."""
    A = sobre_blanco(a) if a else None
    B = sobre_blanco(b) if b else None
    w = max(x.width for x in (A, B) if x); h = max(x.height for x in (A, B) if x)
    lienzo = Image.new("RGB", (w * 3 + 40, h + 30), "white")
    d = ImageDraw.Draw(lienzo)
    d.text((10, 8), "Office  " + titulo, fill="black"); d.text((w + 30, 8), "NOM", fill="black"); d.text((2 * w + 40, 8), "diferencias", fill="black")
    if A: lienzo.paste(A, (0, 30))
    if B: lienzo.paste(B, (w + 20, 30))
    if A and B:
        a2 = Image.new("RGB", (w, h), "white"); a2.paste(A, (0, 0))
        b2 = Image.new("RGB", (w, h), "white"); b2.paste(B, (0, 0))
        ga, gb = a2.convert("L"), b2.convert("L")
        sa = ga.point(lambda v: 255 if v < 160 else 0); sb = gb.point(lambda v: 255 if v < 160 else 0)
        rojo = ImageChops.subtract(sa, sb); azul = ImageChops.subtract(sb, sa)
        dif = Image.new("RGB", (w, h), "white")
        base = Image.merge("RGB", [ga.point(lambda v: 200 + v // 5)] * 3)
        dif.paste(base, (0, 0))
        dif.paste((220, 0, 0), (0, 0), rojo); dif.paste((0, 60, 230), (0, 0), azul)
        lienzo.paste(dif, (2 * w + 40, 30))
    lienzo.save(ruta)


def entrar():
    c = N.cuenta(); esp = N.espacio(c["token"], "Office real")
    return c, esp


def medir_docx(nav, esp, archivo, carpeta, verdad):
    p = nav.p
    r = p.evaluate(JS_IMPORTAR, [base64.b64encode(archivo.read_bytes()).decode(), archivo.name, esp])
    p.wait_for_function(U.JS_LISTO, timeout=120000)
    p.evaluate("() => document.fonts.ready"); p.wait_for_timeout(1500)
    p.evaluate("() => NOM.editorActual.repaginar && NOM.editorActual.repaginar()"); p.wait_for_timeout(800)
    m = p.evaluate(JS_RENGLONES)
    m["aviso"] = r.get("aviso")
    m["paginacion"] = p.evaluate("() => NOM.editorActual.paginacion || null")
    # capturas de cada página
    n = p.evaluate("() => document.querySelectorAll('.wd-pagina').length")
    p.evaluate("() => document.querySelectorAll('.consejo').forEach((e) => e.remove())")
    caps = []
    for i in range(n):
        p.evaluate("(i) => document.querySelectorAll('.wd-pagina')[i].scrollIntoView({block: 'start'})", i)
        p.wait_for_timeout(250)
        caja = p.evaluate("(i) => { const r = document.querySelectorAll('.wd-pagina')[i].getBoundingClientRect(); return {x: r.left, y: r.top, width: r.width, height: r.height}; }", i)
        ruta = carpeta / f"nom-{i + 1:02d}.png"
        p.screenshot(path=str(ruta), clip=caja)
        caps.append(ruta)
    return m, caps


JS_DIAPO = r"""
async ([i, k]) => {
  // (k: escala para que la diapositiva quepa a lo ancho de la pantalla, como en el celular)
  const doc = NOM.editorActual.doc, p = doc.paginas()[i];
  let cont = document.getElementById('__diapo_prueba');
  if (!cont) { cont = document.createElement('div'); cont.id = '__diapo_prueba'; cont.style.cssText = 'position:fixed;left:0;top:0;z-index:99999;background:#fff'; document.body.append(cont); }
  const T = NOM.diapo.tam(doc);
  cont.innerHTML = `<div class="diapo diapo-foto" style="${NOM.diapo.estiloDiapo(doc, p)}position:relative;left:0;top:0;transform:none;border-radius:0;box-shadow:none;margin:0">${NOM.lienzoVista.htmlDiapositiva(doc, p)}</div>`;
  cont.style.width = T.w + 'px'; cont.style.height = T.h + 'px';
  cont.style.transformOrigin = '0 0'; cont.style.transform = k && k !== 1 ? `scale(${k})` : '';
  await document.fonts.ready; await new Promise((r) => setTimeout(r, 120));
  // (como en el editor, al presentar y en el proyector: los renglones de PowerPoint en su altura)
  if (NOM.diapo.ajustarRenglones) NOM.diapo.ajustarRenglones(cont);
  const imgs = [...cont.querySelectorAll('img')]; await Promise.all(imgs.map((im) => im.complete ? 0 : new Promise((r) => { im.onload = im.onerror = r; })));
  return { w: T.w, h: T.h, n: doc.paginas().length };
}
"""
JS_CELDAS = r"""
() => {
  const out = {}, rej = document.querySelector('td[data-c]') ? document.querySelector('td[data-c]').closest('table') : null;
  if (!rej) return { celdas: {}, cols: [], filas: [] };
  rej.querySelectorAll('td[data-c]').forEach((td) => {
    const cs = getComputedStyle(td);
    const x = out[(+td.dataset.f + 1) + ',' + (+td.dataset.c + 1)] = { t: td.innerText, fondo: cs.backgroundColor, img: cs.backgroundImage, color: cs.color, b: +cs.fontWeight >= 600, i: cs.fontStyle === 'italic',
      al: cs.textAlign, fuente: cs.fontFamily, tam: parseFloat(cs.fontSize) * 0.75, w: td.getBoundingClientRect().width, h: td.getBoundingClientRect().height };
    // las celdas que tapa una combinada tienen su relleno (en Excel cada una lo reporta) y ningún texto
    for (let dy = 0; dy < (td.rowSpan || 1); dy++) for (let dx = 0; dx < (td.colSpan || 1); dx++)
      if (dx || dy) out[(+td.dataset.f + 1 + dy) + ',' + (+td.dataset.c + 1 + dx)] = Object.assign({}, x, { t: '' });
  });
  // ancho PINTADO de cada columna (el de su encabezado), no el que dice el <col>
  const cols = [...rej.querySelectorAll('thead th[data-col]')].map((th) => th.getBoundingClientRect().width);
  const filas = [...rej.querySelectorAll('tr[data-fila]')].map((tr) => ({ f: +tr.dataset.fila + 1, h: tr.getBoundingClientRect().height }));
  const a1 = rej.querySelector('td[data-c="0"][data-f="0"]'), r = a1 ? a1.getBoundingClientRect() : null;
  return { celdas: out, cols, filas, a1: r ? { x: r.left, y: r.top } : null };
}
"""


EN_ES = {}
for en, es in zip("January February March April May June July August September October November December".split(), "enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre".split()):
    EN_ES[en.lower()] = es; EN_ES[en[:3].lower()] = es[:3]
for en, es in zip("Sunday Monday Tuesday Wednesday Thursday Friday Saturday".split(), "domingo lunes martes miércoles jueves viernes sábado".split()):
    EN_ES[en.lower()] = es; EN_ES[en[:3].lower()] = es[:3]
EN_ES.update({"am": "a. m.", "pm": "p. m.", "true": "verdadero", "false": "falso", "#value!": "#¡valor!", "#div/0!": "#¡div/0!", "#n/a": "#n/d", "#name?": "#¿nombre?",
              "#ref!": "#¡ref!", "#num!": "#¡num!", "#null!": "#¡nulo!", "#spill!": "#¡desbordamiento!", "#calc!": "#¡calc!"})


def igual_excel(a, b):
    """Mismo texto de celda, sin contar el idioma (Excel en inglés frente a NOM en español) ni los
    espacios que Excel agrega para llenar el ancho (formato contable)."""
    norm = lambda t: re.sub(r"\s+", " ", (t or "").replace("\u00a0", " ")).strip().lower()
    a, b = norm(a), norm(b)
    if a == b: return True
    def es(t):
        t = re.sub(r"#[a-z/0!?]+[!?]?", lambda m: EN_ES.get(m.group(0), m.group(0)), t)
        return re.sub(r"[a-z]+", lambda m: EN_ES.get(m.group(0), m.group(0)), t)
    quita = lambda t: t.replace("¡", "").replace("¿", "").replace(" ", "")
    # la «fecha corta del sistema» (formato 14 de Excel) sale con el orden de la región de Windows: en esa
    # computadora (en-US) mes/día/año; NOM y un Excel de México la muestran día/mes/año
    m1, m2 = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2,4})(.*)", a), re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2,4})(.*)", b)
    if m1 and m2 and int(m1[1]) == int(m2[2]) and int(m1[2]) == int(m2[1]) and m1[3] == m2[3] and m1[4].strip() == m2[4].strip(): return True
    return quita(es(a)) == quita(es(b)) or quita(es(a)) == quita(b)


def rgb_hex(c):
    m = re.match(r"rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)", c or "")
    if not m: return None
    if m.group(4) is not None and float(m.group(4)) == 0: return None
    return "#%02X%02X%02X" % tuple(int(m.group(k)) for k in (1, 2, 3))


def importar(nav, esp, archivo):
    p = nav.p
    r = p.evaluate(JS_IMPORTAR, [base64.b64encode(archivo.read_bytes()).decode(), archivo.name, esp])
    p.wait_for_function(U.JS_LISTO, timeout=180000)
    p.evaluate("() => document.fonts.ready"); p.wait_for_timeout(1500)
    return r


def sobre_blanco(ruta):
    """La imagen sobre fondo blanco (Excel exporta las gráficas con fondo transparente)."""
    im = Image.open(ruta).convert("RGBA"); fondo = Image.new("RGBA", im.size, "white")
    return Image.alpha_composite(fondo, im).convert("RGB")


def parecido_img(a, b, tam=(480, 270)):
    """Qué tanto se parecen dos imágenes (0 a 100): 100 - diferencia media por píxel en escala de grises reducida."""
    import numpy as np
    A = np.asarray(sobre_blanco(a).convert("L").resize(tam), dtype=float); B = np.asarray(sobre_blanco(b).convert("L").resize(tam), dtype=float)
    return round(100 - np.abs(A - B).mean() / 2.55, 1)


def medir_pptx(nav, esp, archivo, carpeta, verdad, nombre):
    p = nav.p
    r = importar(nav, esp, archivo)
    n = p.evaluate("() => NOM.editorActual.doc.paginas().length")
    diapos = []
    ancho = (p.viewport_size or {}).get("width") or 1300
    for i in range(n):
        k = min(1.0, ancho / 960.0)
        info = p.evaluate(JS_DIAPO, [i, k])
        k = min(1.0, ancho / float(info["w"]))
        if k < 1: info = p.evaluate(JS_DIAPO, [i, k])
        ruta = carpeta / f"nom-{i + 1:02d}.png"
        p.screenshot(path=str(ruta), clip={"x": 0, "y": 0, "width": info["w"] * k, "height": info["h"] * k})
        Image.open(ruta).resize((1920, int(round(1920 * info["h"] / info["w"])))).save(ruta)
        of = REAL / "verdad" / f"{nombre}-{i + 1:02d}.png"
        d = {"n": i + 1}
        if of.exists():
            d["parecido"] = parecido_img(of, ruta)
            lado_a_lado(of, ruta, carpeta / f"lado-{i + 1:02d}.png", f"diapositiva {i + 1}")
        diapos.append(d)
    p.evaluate("() => { const c = document.getElementById('__diapo_prueba'); if (c) c.remove(); }")
    vals = [d["parecido"] for d in diapos if "parecido" in d]
    return {"aviso": r.get("aviso"), "diapositivas_office": len(verdad.get("diapositivas", [])), "diapositivas_nom": n,
            "parecido_medio": round(sum(vals) / max(1, len(vals)), 1), "parecido_min": min(vals) if vals else None, "por_diapositiva": diapos}


def medir_xlsx(nav, esp, archivo, carpeta, verdad, nombre):
    p = nav.p
    r = importar(nav, esp, archivo)
    pags = p.evaluate("() => NOM.editorActual.doc.paginas()")
    hojas_vis = [h for h in verdad.get("hojas", []) if h.get("visible") == -1]
    nombres_nom = p.evaluate("() => { const d = NOM.editorActual.doc; return d.paginas().map((pg) => { const h = NOM.formulas.hojasDe(d).find((x) => x.pagina === pg); return h ? NOM.formulas.nombreDeHoja(d, h) : ''; }); }")
    res = {"aviso": r.get("aviso"), "hojas_office": len(verdad.get("hojas", [])), "hojas_visibles_office": len(hojas_vis), "hojas_nom": len(pags), "hojas": []}
    tot_txt = ok_txt = tot_fondo = ok_fondo = tot_col = ok_col = 0
    tot_fil, ok_fil = [0], [0]
    for hi, h in enumerate(hojas_vis):
        if h["nombre"] not in nombres_nom: continue
        pid = pags[nombres_nom.index(h["nombre"])]
        p.evaluate("(pid) => { const ed = NOM.editorActual; const b = document.querySelector(`.pestana-hoja[data-pag=\"${pid}\"]`); if (b) b.click(); }", pid)
        p.wait_for_timeout(900)
        m = p.evaluate(JS_CELDAS)
        malas, fondos_malos = [], []
        f_ocultas = {fi["f"] for fi in h.get("filas", []) if fi.get("oculta")}; c_ocultas = {co["c"] for co in h.get("columnas", []) if co.get("oculta")}
        for c in h.get("celdas", []):
            if c["f"] in f_ocultas or c["c"] in c_ocultas: continue   # Excel no las muestra
            k = f"{c['f']},{c['c']}"; n = m["celdas"].get(k)
            tot_txt += 1
            tn = (n or {}).get("t", "").replace("\n", " ").strip(); to = (c.get("texto") or "").replace("\n", " ").strip()
            if tn == to or igual_excel(to, tn): ok_txt += 1
            elif len(malas) < 60: malas.append({"celda": k, "excel": to, "nom": tn, "formula": c.get("formula")})
            if c.get("fondo"):
                tot_fondo += 1
                # degradado (4000 lineal, 4001 desde el centro): Excel no da un color de relleno; basta que se pinte un degradado
                if c.get("trama") in (4000, 4001): bien = bool(n) and ("gradient" in (n.get("img") or "") or "svg" in (n.get("img") or ""))
                else: bien = bool(n) and rgb_hex(n["fondo"]) == c["fondo"]
                if bien: ok_fondo += 1
                elif len(fondos_malos) < 40: fondos_malos.append({"celda": k, "excel": c["fondo"], "trama": c.get("trama"), "nom": (n or {}).get("fondo")})
        for col in h.get("columnas", [])[:30]:
            j = col["c"] - 1
            if j < len(m["cols"]) and not col.get("oculta"):
                tot_col += 1
                if abs(m["cols"][j] - col["w"] * 96 / 72) <= 2: ok_col += 1
        filas_ok = sum(1 for fi in h.get("filas", [])[:60] if not fi.get("oculta") and any(abs(x["h"] - fi["h"] * 96 / 72) <= 1.5 for x in m["filas"] if x["f"] == fi["f"]))
        filas_tot = sum(1 for fi in h.get("filas", [])[:60] if not fi.get("oculta") and any(x["f"] == fi["f"] for x in m["filas"]))
        tot_fil[0] += filas_tot; ok_fil[0] += filas_ok
        altos_malos = [{"fila": fi["f"], "excel": round(fi["h"] * 96 / 72, 1), "nom": [round(x["h"], 1) for x in m["filas"] if x["f"] == fi["f"]]} for fi in h.get("filas", [])[:60]
                       if not fi.get("oculta") and any(x["f"] == fi["f"] for x in m["filas"]) and not any(abs(x["h"] - fi["h"] * 96 / 72) <= 1.5 for x in m["filas"] if x["f"] == fi["f"])]
        hoja = {"nombre": h["nombre"], "texto_distinto": malas, "fondo_distinto": fondos_malos, "alto_distinto": altos_malos[:20]}
        # imagen: cuadrícula desde A1 en los dos
        of = REAL / "verdad" / f"{nombre}-h{verdad['hojas'].index(h) + 1}.png"
        if of.exists() and m.get("a1"):
            cap = h.get("captura", {})
            ruta = carpeta / f"nom-h{hi + 1}.png"
            p.screenshot(path=str(ruta), clip={"x": max(0, m["a1"]["x"] - 40), "y": max(0, m["a1"]["y"] - 24), "width": 1250, "height": 800})
            off = Image.open(of); x0, y0 = cap.get("a1_x", 40) - 40, cap.get("a1_y", 24) - 24
            off.crop((max(0, x0), max(0, y0), max(0, x0) + 1250, max(0, y0) + 800)).save(carpeta / f"office-h{hi + 1}.png")
            lado_a_lado(carpeta / f"office-h{hi + 1}.png", ruta, carpeta / f"lado-h{hi + 1}.png", h["nombre"])
            hoja["parecido"] = parecido_img(carpeta / f"office-h{hi + 1}.png", ruta, (625, 400))
        # cada gráfica contra la imagen que Excel sacó de ella (Chart.Export): se busca la de NOM por su lugar
        if h.get("graficas"):
            nom_g = p.evaluate("""() => {
              const a1 = document.querySelector('td[data-c="0"][data-f="0"]'); if (!a1) return [];
              const r0 = a1.getBoundingClientRect();
              return [...document.querySelectorAll('.xl-dibujos > .xl-dib')].map((el, i) => { const r = el.getBoundingClientRect(); return { i, x: r.left - r0.left, y: r.top - r0.top, w: r.width, h: r.height, vacio: el.classList.contains('xl-dib-vacio'), svg: !!el.querySelector('svg') }; });
            }""")
            hoja["graficas"] = []
            for gi, g in enumerate(h["graficas"]):
                png = REAL / "verdad" / (g.get("png") or "")
                if not g.get("png") or not png.exists(): continue
                try: Image.open(png).load()
                except Exception: hoja["graficas"].append({"titulo": g.get("titulo"), "tipo": g.get("tipo"), "sin_referencia": True}); continue
                ex, ey = g["x"] * 96 / 72, g["y"] * 96 / 72
                cands = sorted(nom_g, key=lambda e: abs(e["x"] - ex) + abs(e["y"] - ey) + abs(e["w"] - g["w"] * 96 / 72) + abs(e["h"] - g["h"] * 96 / 72))
                d = {"titulo": g.get("titulo"), "tipo": g.get("tipo")}
                if cands and abs(cands[0]["x"] - ex) + abs(cands[0]["y"] - ey) < 40:
                    e = cands[0]; ruta_g = carpeta / f"nom-h{hi + 1}-g{gi + 1}.png"
                    p.locator(".xl-dibujos > .xl-dib").nth(e["i"]).screenshot(path=str(ruta_g))
                    d.update(vacio=e["vacio"], parecido=parecido_img(png, ruta_g, (360, 250)))
                    lado_a_lado(png, ruta_g, carpeta / f"lado-h{hi + 1}-g{gi + 1}.png", str(g.get("titulo") or ""))
                else: d["falta"] = True
                hoja["graficas"].append(d)
        res["hojas"].append(hoja)
    res.update(celdas=tot_txt, texto_igual=ok_txt, pct_texto=round(100 * ok_txt / max(1, tot_txt), 1), fondos=tot_fondo, fondo_igual=ok_fondo,
               pct_fondo=round(100 * ok_fondo / max(1, tot_fondo), 1), columnas=tot_col, ancho_igual=ok_col, pct_ancho=round(100 * ok_col / max(1, tot_col), 1),
               pct_alto=round(100 * ok_fil[0] / max(1, tot_fil[0]), 1))
    return res


def comparar(archivo, movil, salida, nav, esp):
    archivo = Path(archivo)
    if not archivo.exists(): archivo = REAL / "archivos" / archivo.name
    nombre = archivo.stem
    verdad = json.loads((REAL / "verdad" / (nombre + ".verdad.json")).read_text())
    carpeta = Path(salida) / (nombre + ("-movil" if movil else ""))
    carpeta.mkdir(parents=True, exist_ok=True)
    for v in carpeta.glob("*.png"): v.unlink()
    rep = {"archivo": archivo.name, "movil": movil}
    if archivo.suffix == ".docx":
        m, caps = medir_docx(nav, esp, archivo, carpeta, verdad)
        word = renglones_word(verdad)
        nom = [pg["renglones"] for pg in m["paginas"]]
        ordenar_columnas(word, nom)
        rep.update(paginas_word=len(word), paginas_nom=len(nom), aviso=m.get("aviso"), paginacion=m.get("paginacion"),
                   fuera_de_pagina=len(m["fuera"]), **comparar_renglones(word, nom))
        pdf = REAL / "verdad" / (nombre + ".pdf")
        if pdf.exists():
            offs = pdf_a_png(pdf, carpeta)
            for i in range(max(len(offs), len(caps))):
                lado_a_lado(offs[i] if i < len(offs) else None, caps[i] if i < len(caps) else None, carpeta / f"lado-{i + 1:02d}.png", f"página {i + 1}")
            # parecido de cada página como imagen (lo que no es texto: formas, grupos, rellenos, líneas)
            par = [round(parecido_img(offs[i], caps[i], (306, 396)), 1) for i in range(min(len(offs), len(caps)))]
            if par: rep.update(parecido_paginas=par, parecido_medio=round(sum(par) / len(par), 1), parecido_min=min(par))
    elif archivo.suffix == ".pptx":
        rep.update(medir_pptx(nav, esp, archivo, carpeta, verdad, nombre))
    elif archivo.suffix == ".xlsx":
        rep.update(medir_xlsx(nav, esp, archivo, carpeta, verdad, nombre))
    rep["errores_js"] = nav.errores[-10:]
    (carpeta / "reporte.json").write_text(json.dumps(rep, ensure_ascii=False, indent=1))
    if archivo.suffix == ".docx":
        print(f"{archivo.name}: páginas Word {rep.get('paginas_word')} / NOM {rep.get('paginas_nom')} · renglones idénticos "
              f"{rep.get('iguales')}/{rep.get('renglones_word')} ({rep.get('pct_iguales')} %) · en su página {rep.get('pct_misma_pagina')} % · "
              f"desfase vertical mediana {rep.get('dy_mediana_pt')} pt · imagen {rep.get('parecido_medio')} % (mínimo {rep.get('parecido_min')} %)", flush=True)
    elif archivo.suffix == ".pptx":
        print(f"{archivo.name}: diapositivas {rep.get('diapositivas_office')} / {rep.get('diapositivas_nom')} · parecido medio {rep.get('parecido_medio')} % (mínimo {rep.get('parecido_min')} %)", flush=True)
    else:
        print(f"{archivo.name}: hojas {rep.get('hojas_visibles_office')} / {rep.get('hojas_nom')} · texto de celdas igual {rep.get('pct_texto')} % de {rep.get('celdas')} · "
              f"fondos {rep.get('pct_fondo')} % de {rep.get('fondos')} · anchos de columna {rep.get('pct_ancho')} % · altos de fila {rep.get('pct_alto')} %", flush=True)
        for h in rep.get("hojas", []):
            for g in h.get("graficas", []):
                print(f"   gráfica {g.get('titulo')!s:32} {'sin imagen de Excel' if g.get('sin_referencia') else 'falta' if g.get('falta') else ('vacía ' if g.get('vacio') else '') + str(g.get('parecido')) + ' %'}", flush=True)
    return rep


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("archivos", nargs="+")
    ap.add_argument("--movil", action="store_true")
    ap.add_argument("--salida", default=str(REAL / "comparar"))
    a = ap.parse_args()
    N.esperar_servidor()
    c, esp = entrar()
    nav = N.Navegador(movil=a.movil)
    if not a.movil: nav.ctx.close(); nav.ctx = nav.b.new_context(viewport={"width": 1300, "height": 1150}); nav.p = nav.ctx.new_page(); nav.p.on("pageerror", lambda e: nav.errores.append(str(e)))
    # (experimentos: NOM_EXPERIMENTO="window.__x = 1" se corre antes de cada página)
    if os.environ.get("NOM_EXPERIMENTO"): nav.p.add_init_script(os.environ["NOM_EXPERIMENTO"])
    nav.entrar(c)
    nav.p.evaluate("() => { const t = NOM.leer('nom.texto', {}); t.vista = 'paginas'; t.zoom = 1; NOM.guardar && NOM.guardar('nom.texto', t); localStorage.setItem('nom.texto', JSON.stringify(t)); }")
    # los consejos de la primera vez tapan la página en las capturas: se dan por vistos
    nav.p.evaluate("() => { ['consejo.TEXTO', 'consejo.HOJA', 'consejo.PRESENTACION', 'consejo.DIAGRAMA', 'rueda'].forEach((k) => NOM.hito && NOM.hito(k)); }")
    reps = []
    try:
        for f in a.archivos: reps.append(comparar(f, a.movil, a.salida, nav, esp))
    finally:
        nav.cerrar()
    Path(a.salida).mkdir(parents=True, exist_ok=True)
    (Path(a.salida) / ("resumen-movil.json" if a.movil else "resumen.json")).write_text(json.dumps(reps, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

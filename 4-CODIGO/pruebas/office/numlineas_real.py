#!/usr/bin/env python3
"""Números de línea de Word contra los de NOM: cada número que Word pone (en el PDF que exporta, con su hoja, su valor
y su lugar) tiene que estar en NOM en la misma hoja, con el mismo valor, a la misma altura (±1.5 pt) y con el borde
derecho en el mismo lugar (±1.5 pt). También cuenta los números de más.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/numlineas_real.py [ARCHIVO.docx]

El archivo (por omisión w16-numlineas.docx, de windows/hacer_numlineas.py) necesita junto su PDF de Word
(ARCHIVO.pdf, ExportAsFixedFormat). Lee el PDF con pdftotext -bbox.
"""
import argparse, base64, json, os, re, subprocess, sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"


def numeros_word(pdf):
    """[(hoja, valor, centro vertical pt, borde derecho pt, primera palabra de su renglón, centro del renglón)] de los
    números del margen izquierdo de cada columna."""
    x = subprocess.run(["pdftotext", "-bbox", str(pdf), "-"], capture_output=True, text=True).stdout
    out = []
    for n, pg in enumerate(x.split("<page ")[1:], 1):
        ws = [(float(a), float(b), float(c), float(d), t) for a, b, c, d, t in re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)</word>', pg)]
        for w in ws:
            if not re.fullmatch(r"\d+", w[4]): continue
            # un número de línea: lo que sigue en su renglón empieza a la derecha, con hueco (la distancia al texto), y
            # nada antes de él en su renglón (no es un número dentro del texto)
            mismos = [v for v in ws if abs(v[3] - w[3]) < 1.5 and v is not w]
            if any(w[0] - 8 < v[2] <= w[0] + 0.5 for v in mismos): continue   # (pegado a texto de su izquierda: es del texto)
            sig = min((v for v in mismos if v[0] > w[2]), key=lambda v: v[0], default=None)
            if sig and sig[0] - w[2] < 8: continue
            if w[3] - w[1] < 9: continue   # (la llamada de una nota al pie: más chica)
            # (un renglón vacío: lo que sigue es el número de la otra columna, o nada)
            vacio = not sig or (re.fullmatch(r"\d+", sig[4]) and sig[0] - w[2] > 100)
            cy = round((w[1] + w[3]) / 2, 2)
            out.append((n, int(w[4]), cy, round(w[2], 2), None if vacio else sig[4], cy if vacio else round((sig[1] + sig[3]) / 2, 2)))
    return out


JS_NOM = r"""
() => {
  const z = window.__zoomNom || 1, PT = 0.75 / z;
  // la primera palabra del renglón a la derecha del número (y el centro de su letra)
  const palabra = (x, y) => {
    for (let dx = 3; dx < 120; dx += 2) {
      const rg = document.caretRangeFromPoint(x + dx, y);
      if (!rg || rg.startContainer.nodeType !== 3 || !rg.startContainer.parentElement.closest('.wd-cuerpo, .wd-continuaciones')) continue;
      const t = rg.startContainer, v = t.nodeValue;
      let a = rg.startOffset, b = rg.startOffset;
      while (a > 0 && /\S/.test(v[a - 1])) a--;
      while (b < v.length && /\S/.test(v[b])) b++;
      if (b <= a) continue;
      const r2 = document.createRange(); r2.setStart(t, a); r2.setEnd(t, b); const q = r2.getClientRects()[0];
      if (!q || q.left < x - 1 || Math.abs((q.top + q.bottom) / 2 - y) > q.height) continue;   // (la de la otra columna o de otro renglón: no)
      return [v.slice(a, b), (q.top + q.bottom) / 2];
    }
    return [null, y];
  };
  return [...document.querySelectorAll('.wd-pagina')].flatMap((pg, i) => {
    const r = pg.getBoundingClientRect();
    return [...pg.querySelectorAll('.wd-numlin')].map((s) => {
      s.scrollIntoView({ block: 'center' }); const q2 = s.getBoundingClientRect(), r3 = pg.getBoundingClientRect();
      const [w, ly] = palabra(q2.right, (q2.top + q2.bottom) / 2);
      const cy = ((q2.top + q2.bottom) / 2 - r3.top) * PT;
      return [i + 1, +s.textContent, cy, (q2.right - r3.left) * PT, w, w == null ? cy : (ly - r3.top) * PT, cy]; });
  });
}
"""


def abrir(p, esp, arch):
    p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
    p.evaluate(C.JS_IMPORTAR, [base64.b64encode(Path(arch).read_bytes()).decode(), Path(arch).name, esp])
    p.wait_for_function(U.JS_LISTO, timeout=120000)
    p.evaluate("() => document.fonts.ready"); p.wait_for_timeout(1500)
    p.evaluate("() => NOM.editorActual.repaginar && NOM.editorActual.repaginar()"); p.wait_for_timeout(1200)


def comparar(word, nom):
    usados, bien, mal = set(), 0, []
    norm = lambda t: re.sub(r"[^\w]", "", t or "")
    for w in word:
        cand = [i for i, x in enumerate(nom) if i not in usados and x[0] == w[0] and x[1] == w[1]]
        if not cand: mal.append(("falta en NOM", w)); continue
        i = min(cand, key=lambda i: abs(nom[i][2] - w[2]))
        usados.add(i)
        x = nom[i]
        # mismo renglón (su primera palabra), el borde derecho en su lugar y a la misma altura que su renglón
        dy = (x[6] - x[5]) - (w[2] - w[5])
        if norm(x[4]) == norm(w[4]) and abs(x[3] - w[3]) <= 1.5 and abs(dy) <= 1.0: bien += 1
        else: mal.append(("distinto", w, [x[4], round(x[3], 2), round(dy, 2)]))
    de_mas = [nom[i] for i in range(len(nom)) if i not in usados]
    return bien, mal, de_mas


def reportar(nombre, word, bien, mal, de_mas):
    print(f"{nombre}: {bien} de {len(word)} números de Word iguales en NOM (hoja, valor, el mismo renglón, borde ±1.5 pt, altura en su renglón ±1 pt) · de más en NOM: {len(de_mas)}")
    for m in mal[:12]: print("  ", m)
    for x in de_mas[:8]: print("   de más:", [x[0], x[1], round(x[2], 1), round(x[3], 1), x[4]])
    return bien == len(word) and not de_mas


JS_BAJAR = r"""async () => { let u = NOM.ooxml.docx(NOM.editorActual.doc); if (u && u.then) u = await u;
  let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }"""


def word_ver(docx, salida):
    """Word abre el .docx: su PDF (junto a la salida) y cómo dice que quedaron los números de línea."""
    import remoto
    remoto.subir([docx], "nomoffice/entrada")
    r = remoto.correr_office("numlin-" + docx.stem, "numlineas_ver.py", f"entrada\\{docx.name}", f"{docx.stem}.json", f"{docx.stem}.pdf", timeout=900)
    if not remoto.bajar(f"nomoffice/{docx.stem}.json", salida / f"{docx.stem}.word.json") or not remoto.bajar(f"nomoffice/{docx.stem}.pdf", salida / f"{docx.stem}.pdf"):
        raise SystemExit("Word no dejó resultado: " + r.stdout[-300:])
    return json.loads((salida / f"{docx.stem}.word.json").read_text(encoding="utf-8")), salida / f"{docx.stem}.pdf"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("archivo", nargs="?", default=str(REAL / "archivos/w16-numlineas.docx"))
    ap.add_argument("--sin-vuelta", action="store_true", help="solo el archivo de Word (sin cambiar en NOM y abrir en Word)")
    ap.add_argument("--salida", default=str(Path.home() / ".cache/nom-pruebas/registros/numlineas"))
    a = ap.parse_args()
    arch = Path(a.archivo)
    sal = Path(a.salida); sal.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False)
    nav.ctx.close(); nav.ctx = nav.b.new_context(viewport={"width": 1300, "height": 1150}); nav.p = nav.ctx.new_page(); nav.p.on("pageerror", lambda e: nav.errores.append(str(e)))
    nav.entrar(c)
    nav.p.evaluate("() => { const t = NOM.leer('nom.texto', {}); t.vista = 'paginas'; t.zoom = 1; localStorage.setItem('nom.texto', JSON.stringify(t)); }")
    todo_bien = True
    try:
        p = nav.p
        # 1. el archivo de Word, como lo pone Word
        abrir(p, esp, arch)
        word = numeros_word(arch.with_suffix(".pdf"))
        todo_bien &= reportar(arch.name, word, *comparar(word, p.evaluate(JS_NOM)))
        if not a.sin_vuelta:
            # 2. la vuelta: se cambian en NOM (cada 2, empezar en 3, a 0.8 cm, reiniciar en cada sección; un párrafo
            # suprimido y otro que ya no), se descarga y Word lo abre: tiene que decir lo mismo y ponerlos igual
            p.evaluate("() => NOM.editorActual.modoApi.textoApi.numerosDeLinea([2, 3, 30.2, 's'])"); p.wait_for_timeout(800)
            p.evaluate("""() => { const ps = [...document.querySelectorAll('.wd-cuerpo > p')].filter((x) => x.textContent.trim().length > 30);
                const pon = (par) => { let t = par; while (t && t.nodeType !== 3) t = t.firstChild; const r = document.createRange(); r.setStart(t, 1); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); };
                pon(ps[2]); NOM.editorActual.modoApi.textoApi.suprimirNumerosLinea();
                const sup = [...document.querySelectorAll('.wd-cuerpo > p[data-sinnl]')][0]; if (sup) { pon(sup); NOM.editorActual.modoApi.textoApi.suprimirNumerosLinea(); } }""")
            p.wait_for_timeout(1500)
            vuelta = sal / "numlineas-vuelta.docx"
            vuelta.write_bytes(base64.b64decode(p.evaluate(JS_BAJAR)))
            wj, pdf = word_ver(vuelta, sal)
            esperado = {"activo": True, "cada": 2, "inicio": 3, "distancia": 0.8 * 72 / 2.54, "reinicio": 1}
            props = all(s_["activo"] and s_["cada"] == 2 and s_["inicio"] == 3 and abs(s_["distancia"] - esperado["distancia"]) < 0.6 and s_["reinicio"] == 1 for s_ in wj.get("secciones", []))
            print(f"vuelta: Word abre lo descargado ({wj.get('paginas')} páginas) con los números de NOM en sus {len(wj.get('secciones', []))} secciones: {'sí' if props else 'NO: ' + str(wj.get('secciones'))}; suprimidos {wj.get('suprimidos')}")
            todo_bien &= props
            abrir(p, esp, vuelta)
            word2 = numeros_word(pdf)
            todo_bien &= reportar("vuelta (lo que descarga NOM, en Word)", word2, *comparar(word2, p.evaluate(JS_NOM)))
        if nav.errores: print("errores de JS:", nav.errores[:3]); todo_bien = False
    finally:
        nav.cerrar()
    print("números de línea:", "todo bien" if todo_bien else "con diferencias")
    return 0 if todo_bien else 1


if __name__ == "__main__":
    sys.exit(main())

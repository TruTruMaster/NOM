#!/usr/bin/env python3
"""Abre un .docx en NOM y muestra, del párrafo que contiene un texto, sus estilos calculados y sus renglones
(texto, x, ancho de la tinta) y el ancho de la caja: para ver por qué un renglón corta distinto que en Word.

  parrafo_word.py ARCHIVO.docx "texto del párrafo"     (con el servidor en NOM_BASE)
"""
import json, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI.parent)); sys.path.insert(0, str(AQUI.parent.parent / "interfaz"))
import nomprueba as N, comparar_real as C  # noqa: E402

JS = r"""
(marca) => {
  const ps = [...document.querySelectorAll('.wd-cuerpo [data-p], .wd-cuerpo p, .wd-cuerpo div')].filter((e) => (e.textContent || '').includes(marca));
  const el = ps.sort((a, b) => (a.textContent.length - b.textContent.length))[0];
  if (!el) return { no: true };
  const cs = getComputedStyle(el);
  const r = el.getBoundingClientRect();
  // renglones: agrupar los rectángulos de cada letra por su "top"
  const rg = document.createRange(), txt = [], w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let n; const lineas = new Map();
  while ((n = w.nextNode())) for (let i = 0; i < n.length; i++) {
    rg.setStart(n, i); rg.setEnd(n, i + 1);
    const q = rg.getClientRects()[0]; if (!q) continue;
    const k = Math.round(q.top);
    if (!lineas.has(k)) lineas.set(k, { t: '', x0: q.left, x1: q.right });
    const L = lineas.get(k); L.t += n.data[i]; L.x0 = Math.min(L.x0, q.left); L.x1 = Math.max(L.x1, q.right);
  }
  const k = 72 / 96;
  return { tag: el.tagName, clase: el.className, estilo: el.getAttribute('style'), fuente: cs.fontFamily, tam: cs.fontSize, ws: cs.wordSpacing, ls: cs.letterSpacing,
    alin: cs.textAlign, kern: cs.fontKerning, lig: cs.fontVariantLigatures, ancho_pt: r.width * k, pad: [cs.paddingLeft, cs.paddingRight, cs.marginLeft, cs.marginRight, cs.textIndent],
    renglones: [...lineas.values()].map((L) => ({ t: L.t, x0_pt: Math.round((L.x0 - r.left) * k * 100) / 100, ancho_pt: Math.round((L.x1 - L.x0) * k * 100) / 100 })) };
}
"""

def main():
    archivo, marca = Path(sys.argv[1]), sys.argv[2]
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c)
    try:
        C.importar(nav, esp, archivo)
        nav.p.evaluate("() => NOM.editorActual.repaginar && NOM.editorActual.repaginar()"); nav.p.wait_for_timeout(800)
        print(json.dumps(nav.p.evaluate(JS, marca), ensure_ascii=False, indent=1))
    finally:
        nav.cerrar()

if __name__ == "__main__":
    main()

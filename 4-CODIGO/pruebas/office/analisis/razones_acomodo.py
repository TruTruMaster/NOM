#!/usr/bin/env python3
"""Por qué un renglón de NOM no sale igual al de Word: qué párrafos acomoda NOM como Word y cuáles deja al navegador.

Abre cada .docx en NOM (como `comparar_real.py`), y por cada párrafo dice si se acomodó con la regla de Word
(`acomodarComoWord`) y, si no, la razón (`window.__razAcomodo`). Con `--reporte DIR` toma las diferencias de una corrida
de `comparar_real.py --salida DIR` y cuenta cuántos renglones distintos cae en cada razón.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/analisis/razones_acomodo.py \
      ARCHIVO.docx … [--reporte ~/NOM/trabajo/19nov/base]
"""
import argparse, base64, collections, json, re, sys, unicodedata
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent)); sys.path.insert(0, str(Path(__file__).parent.parent.parent / "interfaz"))
import comparar_real as C
import nomprueba as N
import prueba_ui as U

JS = r"""
() => {
  const raz = window.__razAcomodo || new Map();
  const cuerpo = document.querySelector('.wd-cuerpo');
  const out = [];
  const txt = (el) => { const c = el.cloneNode(true); c.querySelectorAll('.wd-pintado, .wd-guion, script, style').forEach((x) => x.remove()); return c.textContent; };
  for (const el of cuerpo.children) {
    const t = txt(el);
    let r = el.dataset.acw != null ? 'acomodado' : raz.get(el) || (el.dataset.acv != null ? 'cabe en un renglón' : 'sin revisar');
    if (el.querySelector('table')) r = 'tabla';
    else if (el.querySelector('[data-obj] .wd-cuadro, [data-obj] .forma-texto, [data-obj] svg foreignObject')) r = r + ' (con cuadro de texto)';
    out.push({ tag: el.tagName, cls: el.className, alin: el.dataset.alin || '', t: t.slice(0, 400), r });
  }
  return out;
}
"""

def norm(s):
    s = unicodedata.normalize("NFC", s)
    return re.sub(r"\s+", " ", s.replace("­", "")).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("archivos", nargs="+")
    ap.add_argument("--reporte")
    a = ap.parse_args()
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador()
    nav.ctx.close(); nav.ctx = nav.b.new_context(viewport={"width": 1300, "height": 1150}); nav.p = nav.ctx.new_page()
    nav.p.add_init_script("window.__razAcomodo = new Map();")
    nav.entrar(c)
    nav.p.evaluate("() => { const t = NOM.leer('nom.texto', {}); t.vista = 'paginas'; t.zoom = 1; localStorage.setItem('nom.texto', JSON.stringify(t)); }")
    total = collections.Counter()
    try:
        for f in a.archivos:
            f = Path(f)
            if not f.exists(): f = C.REAL / "archivos" / f.name
            p = nav.p
            p.evaluate(C.JS_IMPORTAR, [base64.b64encode(f.read_bytes()).decode(), f.name, esp])
            p.wait_for_function(U.JS_LISTO, timeout=120000)
            p.evaluate("() => document.fonts.ready"); p.wait_for_timeout(1500)
            p.evaluate("() => NOM.editorActual.repaginar && NOM.editorActual.repaginar()"); p.wait_for_timeout(800)
            pars = p.evaluate(JS)
            cnt = collections.Counter(x["r"] for x in pars)
            print(f"== {f.name}: {len(pars)} párrafos · " + ", ".join(f"{k} {v}" for k, v in cnt.most_common()))
            if a.reporte:
                rep = Path(a.reporte) / f.stem / "reporte.json"
                if rep.exists():
                    difs = json.loads(rep.read_text())["diferencias"]
                    pn = [(norm(x["t"]), x) for x in pars]
                    por = collections.Counter()
                    for d in difs:
                        for ln in d["nom"] or d["word"]:
                            t = norm(ln.split(": ", 1)[1] if ": " in ln else ln)
                            if len(t) < 3: continue
                            dueño = next((x for n, x in pn if t[:60] in n), None)
                            r = dueño["r"] if dueño else "fuera del cuerpo (celda, cuadro, encabezado, nota)"
                            por[r] += 1
                            if dueño and r == "acomodado": print(f"   [acomodado pero distinto] {t[:110]}")
                    total.update(por)
                    print("   renglones distintos por razón: " + ", ".join(f"{k} {v}" for k, v in por.most_common()))
    finally:
        nav.cerrar()
    if total: print("TOTAL renglones distintos por razón: " + ", ".join(f"{k} {v}" for k, v in total.most_common()))


if __name__ == "__main__":
    main()

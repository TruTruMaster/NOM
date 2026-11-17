#!/usr/bin/env python3
"""La vuelta de las animaciones y transiciones: lo que se hace o se cambia en NOM tiene que salir a PowerPoint igual.

  animaciones_vuelta.py [archivo.pptx …] [--salida DIR]

Para cada archivo (por defecto el catálogo de PowerPoint a01-animaciones y las transiciones a03):
  1. lo importa en la página de NOM y lo exporta sin tocar (las diapositivas deben salir intactas);
  2. cambia en NOM la duración y el retraso de cada efecto y la duración de cada transición (así se escribe de nuevo el
     p:timing y la p:transition) y lo exporta;
  3. vuelve a importar lo exportado y comprueba que cada diapositiva tenga los mismos efectos (clase, número, dirección,
     inicio, retraso y duración) y la misma transición, y que el XML sea válido (lo revisa después validar_openxml.py).
Deja los .pptx en la carpeta de salida para abrirlos en PowerPoint de verdad (vuelta_office.py / abrir_office.py).
"""
import argparse, base64, json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import fidelidad as F

JS = r"""
async ([b64, nombre]) => {
  const bin = atob(b64), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const aB64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const resumen = (doc) => doc.paginas().map((p) => ({
    anim: (doc.valor('_doc', 'anim:' + p) || []).map((x) => [x.efecto, x.inicio, Math.round(+x.retraso || 0), Math.round(+x.dur || 0), x.disp ? 'disparador' : '']),
    trans: (() => { const t = doc.valor('_doc', 'transOffice:' + p); return t ? [t.t, JSON.stringify(t.a || {}), Math.round(+t.dur || 0), t.advTm == null ? null : Math.round(t.advTm)] : null; })() }));
  const r = {};
  try {
    const imp = await NOM.formatos.importar(new File([bytes], nombre));
    const doc = NOM.formatos.docDeOps(imp.nombre, imp.modo, imp.ops);
    r.antes = resumen(doc);
    r.intacta = aB64(await NOM.ooxml.pptx(doc, false));
    // cambiar: cada efecto dura 1.5 veces y empieza 100 ms después; cada transición dura 1.25 veces
    doc.paginas().forEach((p) => {
      const l = doc.valor('_doc', 'anim:' + p);
      if (Array.isArray(l) && l.length) doc.cambiar('_doc', { ['anim:' + p]: l.map((x) => Object.assign({}, x, { dur: Math.round((+x.dur || 500) * 1.5), retraso: Math.round(+x.retraso || 0) + 100 })) });
      const t = doc.valor('_doc', 'transOffice:' + p);
      if (t && t.dur) doc.cambiar('_doc', { ['transOffice:' + p]: Object.assign({}, t, { dur: Math.round(t.dur * 1.25) }) });
    });
    r.esperado = resumen(doc);
    r.cambiada = aB64(await NOM.ooxml.pptx(doc, false));
    // y de vuelta
    const sal = Uint8Array.from(atob(r.cambiada), (c) => c.charCodeAt(0));
    const imp2 = await NOM.formatos.importar(new File([sal], 'vuelta.pptx'));
    const doc2 = NOM.formatos.docDeOps(imp2.nombre, imp2.modo, imp2.ops);
    r.despues = resumen(doc2);
  } catch (e) { r.error = String(e && e.stack || e).slice(0, 900); }
  return r;
}
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("archivos", nargs="*")
    ap.add_argument("--salida", default=str(F.CACHE / "animaciones-vuelta"))
    a = ap.parse_args()
    real = F.CACHE / "office-real" / "archivos"
    archivos = [Path(x) if Path(x).exists() else real / x for x in (a.archivos or ["a01-animaciones.pptx", "a02-opciones.pptx", "a03-transiciones.pptx", "p08-animaciones.pptx"])]
    sal = Path(a.salida); sal.mkdir(parents=True, exist_ok=True)
    from playwright.sync_api import sync_playwright
    malos = 0
    with sync_playwright() as pw:
        nav = pw.chromium.launch(); pag = nav.new_page()
        pag.goto(F.BASE + "/"); pag.wait_for_function("() => window.NOM && NOM.ooxml && NOM.formatos && NOM.formatos.docDeOps && NOM.animOffice", timeout=30000)
        for f in archivos:
            r = pag.evaluate(JS, [base64.b64encode(f.read_bytes()).decode(), f.name])
            if r.get("error"): print(f"{f.name}: ERROR {r['error']}"); malos += 1; continue
            (sal / (f.stem + "-intacta.pptx")).write_bytes(base64.b64decode(r["intacta"]))
            (sal / (f.stem + "-cambiada.pptx")).write_bytes(base64.b64decode(r["cambiada"]))
            difs = []
            for i, (e, d) in enumerate(zip(r["esperado"], r["despues"])):
                # (los efectos instantáneos —Aparecer de 1 ms— no cambian de duración)
                ea = [x[:3] + ([round(x[3] / 2)] if x[3] > 2 else []) + x[4:] for x in e["anim"]]; da = [x[:3] + ([round(x[3] / 2)] if x[3] > 2 else []) + x[4:] for x in d["anim"]]
                if ea != da: difs.append((i + 1, "animaciones", [x for x, y in zip(ea, da) if x != y][:3] or ea[len(da):], [y for x, y in zip(ea, da) if x != y][:3] or da[len(ea):]))
                # (sin efecto, la duración no cuenta: solo el avance automático)
                te, td = e["trans"], d["trans"]
                if te and te[0] == "corte0": te = [te[0], te[3]]
                if td and td[0] == "corte0": td = [td[0], td[3]]
                if (te or [None])[0:3] != (td or [None])[0:3]: difs.append((i + 1, "transición", te, td))
            n_ef = sum(len(p["anim"]) for p in r["esperado"]); n_tr = sum(1 for p in r["esperado"] if p["trans"])
            print(f"{f.name}: {len(r['esperado'])} diapositivas, {n_ef} efectos, {n_tr} transiciones; diferencias tras la vuelta: {len(difs)}")
            for x in difs[:12]: print("   ", x)
            malos += 1 if difs else 0
        nav.close()
    print("listo:", "todo igual" if not malos else f"{malos} archivos con diferencias", "·", sal)
    sys.exit(1 if malos else 0)


if __name__ == "__main__":
    main()

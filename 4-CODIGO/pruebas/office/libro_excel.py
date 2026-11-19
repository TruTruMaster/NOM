#!/usr/bin/env python3
"""Lo que se hace en NOM con un libro (nombres definidos, cálculo manual, esquema…) contra Excel 2024 real.

Cada caso arma una hoja en NOM con la API de la hoja, la exporta a .xlsx, la vuelve a abrir en NOM (la vuelta) y la abre
Excel en la PC de Windows (windows/libro_ver.py): Excel tiene que abrirla y calcular lo mismo que NOM, celda por celda,
con los mismos nombres (y lo propio de cada caso).

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/libro_excel.py [--solo caso,…]
"""
import argparse, base64, json, os, sys, tempfile
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import remoto
import nomprueba as N

SALE = Path.home() / "NOM/trabajo/libro-excel"

# Cada caso: celdas (NOM en español), lo que se hace con la API de la hoja (JS) y lo que se revisa además.
CASOS = {
    "nombres": {
        "celdas": {"A1": "10", "A2": "20", "A3": "30", "A4": "40", "B1": "=SUMA(Ventas)", "B2": "=A1*IVA", "B3": "=Doble(5)",
                   "B4": "=PROMEDIO(Ventas)*IVA", "B5": "=CONTAR(Ventas)+Doble(IVA)", "B6": "=SI(SUMA(Ventas)>Meta;\"sí\";\"no\")"},
        "js": """(H, hoja) => { const g = H.nombres.guardar; return [g("IVA", "=0.16"), g("Ventas", "=" + hoja + "!$A$1:$A$4"), g("Doble", "=LAMBDA(x;x*2)"), g("Meta", "=90")]; }""",
        "nombres": ["Doble", "IVA", "Meta", "Ventas"],
    },
    # Cálculo manual: la fórmula se queda con su valor hasta F9 (como Excel) y el libro sale con calcMode="manual"
    "calculo-manual": {
        "celdas": {"A1": "3", "B1": "=A1*2", "C1": "=SUMA(A1:B1)"},
        "js": """(H, hoja) => { const mal = [], t = H.tabla(), d = NOM.editorActual.doc, v = (r) => { const p = NOM.formulas.parseRef(r); return H.evaluar().valor(p.c, p.f); };
          H.calculo.fijar("manual"); H.calculo.ahora(false);
          if (v("B1") !== 6) mal.push("al empezar B1 = " + v("B1"));
          d.cambiar(t.id, { "c:A1": "5" });
          if (v("B1") !== 6) mal.push("manual: B1 cambió sin F9 (" + v("B1") + ")");
          if (document.querySelector("[data-calcular]").hidden) mal.push("no aparece «Calcular»");
          H.calculo.ahora(false);
          if (v("B1") !== 10 || v("C1") !== 15) mal.push("después de F9 B1 = " + v("B1") + ", C1 = " + v("C1"));
          return mal; }""",
        "revisar": lambda x, r: [] if x.get("calculo") == "manual" else [f"Excel abrió en cálculo {x.get('calculo')}"],
    },
}

JS_CASO = r"""
async ([c, js]) => {
  const ed = NOM.editorActual, H = ed.modoApi.api, t = H.tabla(), F = NOM.formulas, d = ed.doc;
  const campos = {};
  for (const [r, v] of Object.entries(c.celdas)) campos["c:" + r] = v;
  d.cambiar(t.id, campos);
  const hoja = F.prefijoHoja(F.nombreDeHoja(d, t)).replace(/!$/, "");
  const res = (0, eval)(js)(H, hoja);
  await new Promise((r) => setTimeout(r, 200));
  const ev = H.evaluar(), vals = {};
  for (const r of Object.keys(c.celdas)) { const p = F.parseRef(r); const v = ev.valor(p.c, p.f); vals[r] = v && v.codigo ? v.codigo : typeof v === "object" && v !== null && v.valueOf ? v.valueOf() : v; }
  const bin = await NOM.ooxml.xlsx(d);
  let s = ""; for (let i = 0; i < bin.length; i += 0x8000) s += String.fromCharCode.apply(null, bin.subarray(i, i + 0x8000));
  return { res, vals, xlsx: btoa(s), nombres: Object.values(d.valor("_doc", "nombres") || {}).map((o) => o.n).sort() };
}
"""

JS_VUELTA = r"""
async ([b64, nombre, espacio]) => {
  const bin = atob(b64), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const r = await NOM.formatos.importar(new File([bytes], nombre));
  const doc = r.ops.find((o) => o.bloque === "_doc" && o.campos && o.campos.nombres);
  const t = r.ops.filter((o) => o.campos && Object.keys(o.campos).some((k) => k.startsWith("c:")));
  const formulas = {}; t.forEach((o) => Object.entries(o.campos).forEach(([k, v]) => { if (k.startsWith("c:")) formulas[k.slice(2)] = v; }));
  return { nombres: doc ? Object.values(doc.campos.nombres).map((o) => o.n).sort() : [], formulas };
}
"""


def igual(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)): return abs(a - b) <= 1e-9 * max(1, abs(a))
    if isinstance(a, bool) or isinstance(b, bool): return a == b
    return str(a) == str(b)


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--solo", default="")
    a = ap.parse_args()
    solo = [x for x in a.solo.split(",") if x]
    SALE.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    cu = N.cuenta(); esp = N.espacio(cu["token"], "Libro")
    nav = N.Navegador(movil=False); nav.entrar(cu)
    bien = mal = 0
    try:
        for nombre, c in CASOS.items():
            if solo and nombre not in solo: continue
            d = N.documento(cu["token"], esp, "HOJA", nombre)
            nav.p.goto(N.B + "/#/doc/" + d); nav.p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.api", timeout=30000)
            nav.p.wait_for_timeout(500)
            r = nav.p.evaluate(JS_CASO, [{k: v for k, v in c.items() if not callable(v)}, c["js"]])
            arch = SALE / f"{nombre}.xlsx"; arch.write_bytes(base64.b64decode(r["xlsx"]))
            fallas = []
            if any(x for x in (r["res"] or []) if x): fallas.append(f"NOM no aceptó: {r['res']}")
            # la vuelta: NOM abre lo que exportó
            v = nav.p.evaluate(JS_VUELTA, [r["xlsx"], arch.name, esp])
            if c.get("nombres") and v["nombres"] != c["nombres"]: fallas.append(f"vuelta: nombres {v['nombres']}")
            for ref, fx in c["celdas"].items():
                if str(fx).startswith("=") and v["formulas"].get(ref) != fx.replace(";", ";"):
                    if str(v["formulas"].get(ref)).replace(" ", "") != fx.replace(" ", ""): fallas.append(f"vuelta: {ref} {v['formulas'].get(ref)!r} (era {fx!r})")
            # Excel
            remoto.subir([arch], "nomoffice/entrada")
            rr = remoto.correr_office("libro-" + nombre, "libro_ver.py", f"entrada\\{arch.name}", f"libro-{nombre}.json", timeout=600)
            sal = SALE / f"{nombre}.excel.json"
            if not remoto.bajar(f"nomoffice/libro-{nombre}.json", sal): fallas.append("Excel no dejó resultado: " + rr.stdout[-300:])
            else:
                x = json.loads(sal.read_text(encoding="utf-8"))
                if not x.get("abrio"): fallas.append("Excel no lo abrió: " + str(x.get("error")))
                else:
                    if c.get("nombres"):
                        nx = sorted(n["nombre"] for n in x.get("nombres", []) if n.get("visible"))
                        if nx != c["nombres"]: fallas.append(f"nombres en Excel {nx}")
                    cel = x["hojas"][0]["celdas"]
                    for ref, nv in r["vals"].items():
                        ex = cel.get(ref, {}).get("v")
                        if not igual(nv, ex): fallas.append(f"{ref}: NOM {nv!r} · Excel {ex!r} ({cel.get(ref, {}).get('f')})")
                    if c.get("revisar"): fallas += c["revisar"](x, r)
            ok = not fallas
            bien += ok; mal += not ok
            print(("  bien " if ok else "  MAL  ") + nombre + ("" if ok else "\n     " + "\n     ".join(fallas)), flush=True)
    finally:
        nav.cerrar()
    print(f"\nLibro de NOM contra Excel 2024: {bien} bien, {mal} mal")
    return 1 if mal else 0


if __name__ == "__main__":
    sys.exit(main())

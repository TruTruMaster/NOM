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
    # Auditoría: los precedentes y dependientes directos de cada celda, como DirectPrecedents/DirectDependents de Excel
    "auditoria": {
        "celdas": {"A1": "2", "A2": "3", "D1": "0.1", "B1": "=A1*2", "B2": "=SUMA(A1:A2)+B1", "C1": "=B2*Tasa", "C2": "=B1+$A$2"},
        "js": """async (H, hoja) => { H.nombres.guardar("Tasa", "=" + hoja + "!$D$1"); await new Promise((r) => setTimeout(r, 200));
          const A = H.auditoria, F = NOM.formulas, datos = {};
          const celdas = (l) => { const out = []; l.forEach((x) => { if (x.c1 != null) { for (let f = x.f1; f <= x.f2; f++) for (let c = x.c1; c <= x.c2; c++) out.push(F.ref(c, f)); } else out.push(F.ref(x.c, x.f)); }); return [...new Set(out)].sort(); };
          for (const r of ["A1", "A2", "D1", "B1", "B2", "C1", "C2"]) { const p = F.parseRef(r); datos[r] = { prec: celdas(A.precedentes(p.c, p.f).aqui), dep: celdas(A.dependientes(p.c, p.f).aqui) }; }
          A.rastrear("p", 2, 0); A.rastrear("p", 2, 0); const n = A.flechas().length;
          return { mal: n >= 3 ? [] : ["flechas de C1 en dos niveles: " + n], datos }; }""",
        "revisar": lambda x, r: revisar_auditoria(x, r),
    },
    # Evaluar fórmula: cada paso (la fórmula con lo ya calculado escrito como valor) lo calcula Excel y tiene que dar lo
    # mismo que la celda
    "evaluar": {
        "celdas": {"A1": "2", "A2": "3", "B1": "=A1*2", "C1": "=SUMA(A1:A2)+B1*2", "C2": "=SI(A1>2;\"sí\";\"no\")&\" x\"", "C3": "=(A1+A2)*(B1-1)",
                   "C4": "=PROMEDIO(A1:A2;B1)/FILAS(A1:A2)", "C5": "=-A1^2+B1%", "C6": "=Y(A1<A2;B1=4)", "C7": "=INDICE(A1:B2;2;1)*10"},
        "js": """async (H, hoja) => { const F = NOM.formulas, t = H.tabla(), ev = H.evaluar(), datos = {}, expr = [];
          for (const r of ["C1", "C2", "C3", "C4", "C5", "C6", "C7"]) { const p = F.parseRef(r); const pasos = F.pasosEvaluar(t["c:" + r], ev, p.c, p.f);
            datos[r] = pasos.map((x) => x.texto); pasos.slice(0, -1).forEach((x) => expr.push([r, NOM.ooxml.formulaExcel(x.texto, 9)])); }
          return { mal: [], datos, expr: expr.map((x) => x[1].replace(/^=/, "")), donde: expr.map((x) => x[0]) }; }""",
        "expr": True,
        "revisar": lambda x, r: revisar_evaluar(x, r),
    },
    # Mostrar fórmulas (Ctrl+`): la hoja muestra la fórmula escrita y sale con showFormulas="1"
    "mostrar-formulas": {
        "celdas": {"A1": "2", "A2": "=A1*3", "A3": "=SUMA(A1:A2)"},
        "js": """async (H, hoja) => { const mal = []; H.formulasVisibles(); await new Promise((r) => setTimeout(r, 400));
          const td = document.querySelector('td[data-c="0"][data-f="1"]');
          if (!td || td.textContent.trim() !== "=A1*3") mal.push("A2 muestra " + (td && td.textContent));
          return mal; }""",
        "revisar": lambda x, r: [] if x["hojas"][0].get("formulas_visibles") is True else [f"Excel: fórmulas visibles = {x['hojas'][0].get('formulas_visibles')}"],
    },
    # Tabla de datos de dos variables hecha en NOM (Datos → Análisis de hipótesis → Tabla de datos): Excel la abre como suya
    # (cada celda del interior con =TABLE(B1,B2)) y calcula lo mismo
    "tabla-datos": {
        "celdas": {"A1": "x", "B1": "2", "A2": "y", "B2": "3", "A4": "=B1*B2+1", "B4": "1", "C4": "2", "D4": "3", "A5": "10", "A6": "20", "A7": "30"},
        "js": """async (H, hoja) => { const r = H.herramientas.tablaDatos({ c1: 0, c2: 3, f1: 3, f2: 6 }, "B1", "B2"); await new Promise((q) => setTimeout(q, 300));
          return typeof r === "string" ? [r] : []; }""",
        "revisar": lambda x, r: [f"{k}: Excel dice {c.get('f')!r}" for k, c in celdas_de_hoja(x, "B5:D7").items() if c.get("f") != "=TABLE(B1,B2)"],
    },
    # Esquema: filas 2-6 en un grupo y 3-4 adentro (nivel 2), contraído; columnas B-C agrupadas y contraídas
    "esquema": {
        "celdas": {"A1": "Mes", "A2": "1", "A3": "2", "A4": "3", "A5": "4", "A6": "5", "A7": "=SUMA(A2:A6)", "B1": "x", "C1": "y", "D1": "=B1&C1"},
        "js": """async (H, hoja) => { const E = H.esquema, mal = [];
          E.agrupar("f", 1, { c1: 0, c2: 0, f1: 1, f2: 5 }); E.agrupar("f", 1, { c1: 0, c2: 0, f1: 2, f2: 3 });
          E.agrupar("c", 1, { c1: 1, c2: 2, f1: 0, f2: 0 });
          E.alternar("f", 6); E.alternar("c", 3);
          await new Promise((r) => setTimeout(r, 400));
          const t = H.tabla();
          if (!t["oh:2"] || !t["oh:6"] || t["oh:7"]) mal.push("NOM no contrajo las filas");
          if (!t["ow:B"] || !t["ow:C"] || t["ow:D"]) mal.push("NOM no contrajo las columnas");
          if (document.querySelector('[data-filah="1"]')) mal.push("NOM sigue mostrando la fila 2");
          if (!document.querySelector('[data-esq-f="6"]')) mal.push("no está el botón + de la fila 7");
          return mal; }""",
        "esquema": {"ol:2": 1, "ol:3": 2, "ol:4": 2, "ol:5": 1, "ol:6": 1, "olc:B": 1, "olc:C": 1, "oh:2": True, "oh:6": True, "ow:B": True, "ow:C": True},
        "revisar": lambda x, r: (
            ([] if x["hojas"][0]["filas"] == {"2": {"nivel": 2, "oculta": True}, "3": {"nivel": 3, "oculta": True}, "4": {"nivel": 3, "oculta": True},
                                               "5": {"nivel": 2, "oculta": True}, "6": {"nivel": 2, "oculta": True}} else [f"filas en Excel {x['hojas'][0]['filas']}"])
            + ([] if x["hojas"][0]["columnas"] == {"2": {"nivel": 2, "oculta": True}, "3": {"nivel": 2, "oculta": True}} else [f"columnas en Excel {x['hojas'][0]['columnas']}"])),
    },
}

JS_CASO = r"""
async ([c, js]) => {
  const ed = NOM.editorActual, H = ed.modoApi.api, t = H.tabla(), F = NOM.formulas, d = ed.doc;
  const campos = {};
  for (const [r, v] of Object.entries(c.celdas)) campos["c:" + r] = v;
  d.cambiar(t.id, campos);
  const hoja = F.prefijoHoja(F.nombreDeHoja(d, t)).replace(/!$/, "");
  const res = await (0, eval)(js)(H, hoja);
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
  const esquema = {}; t.forEach((o) => Object.entries(o.campos).forEach(([k, v]) => { if (/^(ol|olc|oh|ow):/.test(k)) esquema[k] = v; }));
  r.ops.forEach((o) => o.campos && Object.entries(o.campos).forEach(([k, v]) => { if (/^(ol|olc|oh|ow):/.test(k)) esquema[k] = v; }));
  return { nombres: doc ? Object.values(doc.campos.nombres).map((o) => o.n).sort() : [], formulas, esquema };
}
"""


def celdas_de_hoja(x, rango):
    """Las celdas de un rango («B5:D7») de la primera hoja de lo que leyó Excel: {«B5»: {f, v, t}}."""
    import re
    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", rango)
    col = lambda t: sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(t)))
    letra = lambda n: (letra((n - 1) // 26) if n > 26 else "") + chr(65 + (n - 1) % 26)
    cs = x["hojas"][0].get("celdas", {})
    return {f"{letra(c)}{f}": (cs.get(f"{letra(c)}{f}") or {}) for f in range(int(m.group(2)), int(m.group(4)) + 1) for c in range(col(m.group(1)), col(m.group(3)) + 1)}


def celdas_de(addr):
    """«A1:A2,B1» (dirección de Excel) → ["A1", "A2", "B1"]."""
    import re
    def col(x): n = 0; [n := n * 26 + ord(ch) - 64 for ch in x]; return n
    def letra(n): r = ""; exec("while n:\n n, k = divmod(n - 1, 26)\n r = chr(65 + k) + r", {}, loc := {"n": n, "r": r}); return loc["r"]
    out = set()
    for parte in [p for p in addr.split(",") if p]:
        a, _, b = parte.partition(":"); b = b or a
        ma, mb = re.match(r"([A-Z]+)(\d+)", a), re.match(r"([A-Z]+)(\d+)", b)
        for f in range(int(ma.group(2)), int(mb.group(2)) + 1):
            for c in range(col(ma.group(1)), col(mb.group(1)) + 1): out.add(letra(c) + str(f))
    return sorted(out)


def revisar_evaluar(x, r):
    mal, cel, res = [], x["hojas"][0]["celdas"], r["res"]
    ev = x.get("evaluados") or []
    if len(ev) != len(res["expr"]): return [f"Excel calculó {len(ev)} de {len(res['expr'])} pasos"]
    for e, donde, v in zip(res["expr"], res["donde"], ev):
        final = cel.get(donde, {}).get("v")
        if not igual(v, final): mal.append(f"{donde}: el paso «{e}» da {v!r} en Excel (la celda da {final!r})")
    return mal


def revisar_auditoria(x, r):
    mal = []
    cel = x["hojas"][0]["celdas"]
    for ref, d in r["res"]["datos"].items():
        ex = cel.get(ref, {})
        for k in ("prec", "dep"):
            e = celdas_de(ex.get(k, "")) if ex else []
            if e != d[k]: mal.append(f"{ref} {k}: NOM {d[k]} · Excel {e}")
    return mal


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
            malos = r["res"].get("mal", []) if isinstance(r["res"], dict) else [x for x in (r["res"] or []) if x]
            if malos: fallas.append(f"NOM no aceptó: {malos}")
            # la vuelta: NOM abre lo que exportó
            v = nav.p.evaluate(JS_VUELTA, [r["xlsx"], arch.name, esp])
            if c.get("nombres") and v["nombres"] != c["nombres"]: fallas.append(f"vuelta: nombres {v['nombres']}")
            if c.get("esquema") and {k: v["esquema"].get(k) for k in c["esquema"]} != c["esquema"]: fallas.append(f"vuelta: esquema {v['esquema']}")
            for ref, fx in c["celdas"].items():
                if str(fx).startswith("=") and v["formulas"].get(ref) != fx.replace(";", ";"):
                    if str(v["formulas"].get(ref)).replace(" ", "") != fx.replace(" ", ""): fallas.append(f"vuelta: {ref} {v['formulas'].get(ref)!r} (era {fx!r})")
            # Excel
            remoto.subir([arch], "nomoffice/entrada")
            extra = []
            if c.get("expr"):
                ej = SALE / f"{nombre}.expr.json"; ej.write_text(json.dumps(r["res"]["expr"], ensure_ascii=False), encoding="utf-8")
                remoto.subir([ej], "nomoffice/entrada"); extra = [f"entrada\\{ej.name}"]
            rr = remoto.correr_office("libro-" + nombre, "libro_ver.py", f"entrada\\{arch.name}", f"libro-{nombre}.json", *extra, timeout=600)
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

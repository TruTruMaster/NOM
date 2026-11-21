#!/usr/bin/env python3
"""Escenarios de Excel (Datos → Análisis de hipótesis → Administrador de escenarios) en NOM contra Excel 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/escenarios_real.py [x11-escenarios.xlsx]

1. x11 (hecho en Excel con windows/hacer_escenarios.py): NOM lee sus 4 escenarios (nombres, celdas, valores y comentarios).
2. Mostrar «Optimista» en NOM da las mismas celdas que en Excel (windows/escenarios_ver.py).
3. El informe resumen de NOM tiene los mismos números que el de Excel (Scenarios.CreateSummary), con «Optimista» mostrado,
   y el mismo formato que el resumen que hizo Excel dentro de x11 (como lo importa NOM: rellenos, letra, rayas, grupos…).
4. Cambiados en NOM (uno nuevo y uno quitado) y descargados: Excel lee los mismos escenarios y abre el libro sin avisos.
5. Un libro hecho en NOM desde cero con dos escenarios y un nombre definido: Excel los lee y su resumen da lo mismo que el
   de NOM (también el nombre en vez de $B$3).
6. El Administrador de escenarios: agregar uno con el formulario y mostrarlo.
Se corre desde Linux (Excel por SSH) con el servidor de pruebas en NOM_BASE.
"""
import base64, json, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
SAL = Path.home() / "NOM/trabajo/escenarios-real"
res = []


def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {str(det)[:300]}" if det and not ok else ""), flush=True)


def excel_ver(xlsx, hoja=None, resultado=None, mostrar=None):
    import remoto
    remoto.subir([xlsx], "nomoffice/entrada")
    args = [f"entrada\\{xlsx.name}", f"{xlsx.stem}.esc.json"] + ([hoja, resultado] if hoja else []) + ([f'"{mostrar}"'] if mostrar else [])
    r = remoto.correr_office("esc-" + xlsx.stem, "escenarios_ver.py", *[f'"{a}"' if " " in a and not a.startswith('"') else a for a in args], timeout=600)
    if not remoto.bajar(f"nomoffice/{xlsx.stem}.esc.json", SAL / f"{xlsx.stem}.esc.json"):
        raise SystemExit("Excel no dejó resultado: " + (r.stdout or "")[-300:])
    return json.loads((SAL / f"{xlsx.stem}.esc.json").read_text(encoding="utf-8"))


def num(v):
    try: return round(float(v), 6)
    except Exception: return v


def lista_nom(p):
    return p.evaluate("() => NOM.editorActual.modoApi.api.escenarios.lista().lista.map((s) => ({ nombre: s.n, celdas: s.c.map((x) => x[0]), valores: s.c.map((x) => x[1]), comentario: s.com || '' }))")


def compara_lista(nombre, nom, excel):
    exp = [{"nombre": e["nombre"], "celdas": e["celdas"], "valores": [num(v) for v in e["valores"]], "comentario": e["comentario"]} for e in excel]
    got = []
    for e in nom:
        cs = e["celdas"]
        rango = cs[0] + ":" + cs[-1] if len(cs) > 1 else cs[0]
        got.append({"nombre": e["nombre"], "celdas": rango, "valores": [num(v) for v in e["valores"]], "comentario": e["comentario"]})
    anota(nombre, got == exp, f"NOM {got} · Excel {exp}")


JS_BAJAR = r"""async () => { const doc = NOM.editorActual.doc; if (NOM.binarios && NOM.binarios.refsDoc(doc).size) await NOM.binarios.precargar(doc);
  let u = NOM.ooxml.xlsx(doc); if (u && u.then) u = await u;
  let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }"""
JS_RESUMEN = r"""(r) => { const R = NOM.editorActual.modoApi.api.escenarios.resumen(r); if (typeof R === 'string') return { error: R };
  const t = NOM.editorActual.modoApi.api.tabla(), ev = NOM.editorActual.modoApi.api.evaluar(), out = {};
  for (const k in t) if (k.startsWith('c:') && t[k] != null && t[k] !== '') { const p = NOM.formulas.parseRef(k.slice(2)); const v = ev.valor(p.c, p.f); out[k.slice(2)] = v && typeof v === 'object' && v.valueOf ? +v : v; }
  return { hoja: R.hoja, celdas: out }; }"""


def compara_resumen(nombre, nom, excel):
    # los números de la tabla (valores actuales y de cada escenario): desde D6 hasta la última columna y fila con números
    malos = []
    for k, v in excel["celdas"].items():
        if isinstance(v, (int, float)):
            w = nom["celdas"].get(k)
            if num(w) != num(v): malos.append(f"{k}: Excel {v} · NOM {w}")
    nums = sum(1 for v in excel["celdas"].values() if isinstance(v, (int, float)))
    anota(f"{nombre}: los {nums} números iguales a los del resumen de Excel", nums > 0 and not malos, malos[:6])


JS_FORMATO = r"""() => { const d = NOM.editorActual.doc, hs = NOM.formulas.hojasDe(d), de = (n) => hs.find((h) => NOM.formulas.nombreDeHoja(d, h) === n);
  const P = /^(n|i|u|k|r|al|aj|bd|bx|tl|tp|ax|ol|oh|olc|aw|ap|ahx):/;
  const fmt = (h) => { const o = {}; for (const k in h) if (P.test(k) && h[k] != null && h[k] !== false && h[k] !== '') o[k] = h[k]; o.rejilla = h.xl ? h.xl.rejilla : null; return o; };
  const a = de('Scenario Summary'), b = de('Resumen del escenario'); return a && b ? [fmt(a), fmt(b)] : null; }"""


def compara_formato(nombre, par):
    """Todo igual, menos el ancho de las columnas de valores (D…): Excel en inglés lo ajusta a «Current Values:» y NOM, a
    «Valores actuales:», que es más largo (como el Excel en español); de esas se pide que todas midan lo mismo y no menos."""
    if not par: return anota(nombre, False, "falta una de las dos hojas")
    a, b = par
    aparte = lambda k: k.startswith(("aw:", "ap:")) and k[3:] not in ("A", "B", "C")
    malos = [f"{k}: Excel {a.get(k)!r} · NOM {b.get(k)!r}" for k in sorted(set(a) | set(b)) if not aparte(k) and json.dumps(a.get(k), sort_keys=True) != json.dumps(b.get(k), sort_keys=True)]
    anota(f"{nombre} ({len(a)} propiedades)", not malos, malos[:8])
    anchos = {k: v for k, v in b.items() if aparte(k) and k.startswith("ap:")}
    anota("…y las columnas de valores, todas del mismo ancho y no más angostas que en Excel", len(set(anchos.values())) == 1 and min(anchos.values()) >= a.get("ap:D", 0) and len(anchos) == len([k for k in a if aparte(k) and k.startswith("ap:")]), anchos)


def rotulos(celdas):
    """Los rótulos de la columna C (las celdas cambiantes y de resultado: $B$1 o su nombre)."""
    return {k: v for k, v in celdas.items() if k.startswith("C") and k[1:].isdigit()}


def ir_a_hoja(p, nombre):
    p.evaluate("(n) => { const b = [...document.querySelectorAll('.pestana-hoja')].find((x) => x.textContent.trim() === n); b && b.click(); }", nombre)
    p.wait_for_timeout(800)


def main():
    arch = Path(sys.argv[1]) if len(sys.argv) > 1 else REAL / "archivos/x11-escenarios.xlsx"
    SAL.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    try:
        verdad = excel_ver(arch, "Plan", "B5:B7", "Optimista")
        # 1. lo que lee NOM
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1200)
        ir_a_hoja(p, "Plan")
        compara_lista("NOM lee los escenarios del libro de Excel", lista_nom(p), verdad["hojas"]["Plan"])
        # 2. mostrar «Optimista»
        p.evaluate("() => NOM.editorActual.modoApi.api.escenarios.mostrar(2)"); p.wait_for_timeout(600)
        celdas = p.evaluate("""() => { const ev = NOM.editorActual.modoApi.api.evaluar(), out = {}; for (let f = 0; f < 7; f++) { const v = ev.valor(1, f); out['B' + (f + 1)] = v && typeof v === 'object' && v.valueOf ? +v : v; } return out; }""")
        exp = {k: v for k, v in verdad["mostrado"].items() if k.startswith("B")}
        anota("mostrar «Optimista»: las mismas celdas que en Excel", {k: num(v) for k, v in celdas.items()} == {k: num(v) for k, v in exp.items()}, f"NOM {celdas} · Excel {exp}")
        # 3. el resumen
        rnom = p.evaluate(JS_RESUMEN, "B5:B7"); p.wait_for_timeout(600)
        compara_resumen("el resumen de NOM", rnom, verdad["resumen"])
        anota("…con los mismos rótulos ($B$1…)", rotulos(rnom["celdas"]) == rotulos(verdad["resumen"]["celdas"]), f"NOM {rotulos(rnom['celdas'])} · Excel {rotulos(verdad['resumen']['celdas'])}")
        compara_formato("…y el formato del resumen que hizo Excel en x11", p.evaluate(JS_FORMATO))
        anota("…y va en una hoja nueva antes de la de los escenarios", p.evaluate("() => { const ps = [...document.querySelectorAll('.pestana-hoja')].map((x) => x.textContent.trim()); return ps.indexOf('Resumen del escenario') === ps.indexOf('Plan') - 1; }"))
        # 4. cambiados en NOM y descargados
        ir_a_hoja(p, "Plan")
        p.evaluate("() => { const E = NOM.editorActual.modoApi.api.escenarios; E.guardar('Nuevo', 'B2', ['800'], 'de NOM'); E.quitar(3); return 1; }"); p.wait_for_timeout(800)
        cambiado = SAL / "x11-cambiado.xlsx"; cambiado.write_bytes(base64.b64decode(p.evaluate(JS_BAJAR)))
        v2 = excel_ver(cambiado)
        esperado = [e for e in verdad["hojas"]["Plan"] if e["nombre"] != "Solo precio"] + [{"nombre": "Nuevo", "celdas": "B2", "valores": ["800"], "comentario": "de NOM"}]
        compara_lista("cambiados en NOM: Excel lee los mismos escenarios", lista_nom(p), esperado)
        anota("…y Excel también", [e["nombre"] for e in v2["hojas"].get("Plan", [])] == [e["nombre"] for e in esperado] and v2["hojas"]["Plan"][-1]["comentario"] == "de NOM", v2["hojas"].get("Plan"))
        # 5. un libro hecho en NOM desde cero
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'Escenarios de NOM', modo: 'HOJA', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function(U.JS_LISTO, timeout=60000); p.wait_for_timeout(1000)
        p.evaluate("""() => { const H = NOM.editorActual.modoApi.api, t = H.tabla(); NOM.editorActual.doc.cambiar(t.id, { 'c:A1': 'a', 'c:B1': '10', 'c:A2': 'b', 'c:B2': '20', 'c:A3': 'a×b', 'c:B3': '=B1*B2+1' }); return 1; }""")
        p.wait_for_timeout(500)
        p.evaluate("() => { const E = NOM.editorActual.modoApi.api.escenarios; E.guardar('Uno', 'B1:B2', ['1', '2'], 'chico'); E.guardar('Dos', 'B1', ['30'], ''); return 1; }"); p.wait_for_timeout(500)
        hoja0 = p.evaluate("() => NOM.formulas.nombreDeHoja(NOM.editorActual.doc, NOM.editorActual.modoApi.api.tabla())")
        malo = p.evaluate("(h) => NOM.editorActual.modoApi.api.nombres.guardar('Producto', `='${h}'!$B$3`)", hoja0)
        if malo: anota("el nombre definido", False, malo)
        p.wait_for_timeout(300)
        nuevo = SAL / "desde-nom.xlsx"; nuevo.write_bytes(base64.b64decode(p.evaluate(JS_BAJAR)))
        nombre_hoja = p.evaluate("() => NOM.formulas.nombreDeHoja(NOM.editorActual.doc, NOM.editorActual.modoApi.api.tabla())")
        v3 = excel_ver(nuevo, nombre_hoja, "B3")
        compara_lista("hecho en NOM desde cero: Excel lee los escenarios", lista_nom(p), v3["hojas"].get(nombre_hoja, []))
        rnom2 = p.evaluate(JS_RESUMEN, "B3"); p.wait_for_timeout(500)
        compara_resumen("…y el resumen de NOM", rnom2, v3["resumen"])
        anota("…con el nombre definido en vez de $B$3, como Excel", rotulos(rnom2["celdas"]) == rotulos(v3["resumen"]["celdas"]) and "Producto" in rotulos(rnom2["celdas"]).values(), f"NOM {rotulos(rnom2['celdas'])} · Excel {rotulos(v3['resumen']['celdas'])}")
        # 6. el Administrador de escenarios de verdad: agregar con el formulario y mostrar
        ir_a_hoja(p, hoja0)
        p.evaluate("() => NOM.editorActual.modoApi.api.dialogos.escenarios()"); p.wait_for_timeout(600)
        n_lista = p.evaluate("() => document.querySelectorAll('.hoja-panel input[name=esc], [role=dialog] input[name=esc], input[name=esc]').length")
        anota("el administrador muestra los 2 escenarios", n_lista == 2, n_lista)
        p.locator("button[data-agr]").click(); p.wait_for_timeout(500)
        dlg = p.locator("section.hoja[role=dialog]").last
        dlg.locator("input[data-nombre]").fill("Grande")
        dlg.locator("input[data-celdas]").fill("B1:B2"); dlg.locator("input[data-celdas]").dispatch_event("input"); p.wait_for_timeout(200)
        dlg.locator('input[data-v="B1"]').fill("7"); dlg.locator('input[data-v="B2"]').fill("9")
        com = dlg.locator("input[data-com]").input_value()
        anota("el comentario de uno nuevo empieza como en Excel («Creado por … el …»)", com.startswith("Creado por ") and " el " in com, com)
        dlg.locator("button[data-si]").click(); p.wait_for_timeout(700)
        lista = lista_nom(p)
        anota("…y con el formulario queda el tercero", [e["nombre"] for e in lista] == ["Uno", "Dos", "Grande"] and lista[2]["valores"] == ["7", "9"], lista)
        dlg = p.locator("section.hoja[role=dialog]").last
        dlg.locator("input[name=esc][value='2']").check(); p.wait_for_timeout(200)
        dlg.locator("button[data-ver]").click(); p.wait_for_timeout(700)
        b3 = p.evaluate("() => { const v = NOM.editorActual.modoApi.api.evaluar().valor(1, 2); return v && typeof v === 'object' ? +v : v; }")
        anota("…y «Mostrar» lo pone en la hoja (B3 = 7×9+1)", b3 == 64, b3)
        import office_abre
        r = office_abre.abrir([cambiado, nuevo], "escenarios")
        for a in (cambiado, nuevo):
            anota(f"Excel abre {a.name} sin avisos ni reparar", office_abre.bien(r.get(a.name)), json.dumps(r.get(a.name), ensure_ascii=False)[:200])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nEscenarios contra Excel 2024: {sum(1 for x in res if x[1])} de {len(res)} bien")
    return 0 if all(x[1] for x in res) else 1


if __name__ == "__main__":
    sys.exit(main())

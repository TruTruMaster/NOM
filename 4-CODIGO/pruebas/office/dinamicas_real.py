#!/usr/bin/env python3
"""Tablas dinámicas hechas en NOM (Insertar → Tabla dinámica) contra Excel 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/dinamicas_real.py

1. Unos datos de ventas (30 filas) en un libro de NOM; Excel hace con ellos seis tablas dinámicas (windows/dinamica_excel.py):
   un campo en filas, dos valores, un campo en columnas, dos campos en filas (subtotales), un filtro y otro estilo.
2. NOM hace las mismas (cada una en una hoja nueva, en A3, como Excel) y cada celda dice lo mismo que en Excel (el texto
   en español: «Etiquetas de fila», «Suma de…», «Total general»), con la misma negrita, relleno, letra y sangría.
3. Descargado, Excel ve las seis tablas dinámicas con sus campos; al actualizarlas él mismo no cambian (windows/dinamicas_ver.py).
4. En un libro de Excel que ya trae tablas dinámicas (x05), una hecha en NOM se suma a las de antes.
5. Excel abre todo sin avisos.
"""
import base64, json, os, re, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
SAL = Path.home() / "NOM/trabajo/dinamicas"
CAMPOS = ["Región", "Vendedor", "Producto", "Unidades", "Ventas"]
CASOS = [
    {"id": "r1v1", "filas": ["Región"], "valores": [["Ventas", "sum"]]},
    {"id": "r1v2", "filas": ["Vendedor"], "valores": [["Ventas", "sum"], ["Unidades", "count"]]},
    {"id": "r1c1", "filas": ["Región"], "columnas": ["Producto"], "valores": [["Unidades", "sum"]]},
    {"id": "r2v1", "filas": ["Región", "Producto"], "valores": [["Ventas", "average"]]},
    {"id": "f1", "filas": ["Producto"], "filtro": "Región", "valores": [["Unidades", "max"], ["Unidades", "min"]]},
    {"id": "medio9", "filas": ["Vendedor"], "columnas": ["Región"], "valores": [["Ventas", "sum"]], "estilo": "PivotStyleMedium9"},
]
res = []
EN = [("Etiquetas de fila", "Row Labels"), ("Etiquetas de columna", "Column Labels"), ("Total general", "Grand Total"), ("(Todas)", "(All)"),
      ("Suma de ", "Sum of "), ("Cuenta de ", "Count of "), ("Promedio de ", "Average of "), ("Máx. de ", "Max of "), ("Mín. de ", "Min of ")]


def en(s):
    for a, b in EN: s = s.replace(a, b) if isinstance(s, str) else s
    return s


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:900]}"), flush=True)


def rc(a):
    m = re.match(r"([A-Z]+)(\d+)", a)
    return sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(m.group(1)))) - 1, int(m.group(2)) - 1


def ref(c, f):
    s, c = "", c + 1
    while c: c, m = divmod(c - 1, 26); s = chr(65 + m) + s
    return f"{s}{f + 1}"


def remoto_excel(script, entrada, salida, *extra, timeout=900):
    import remoto
    remoto.subir([entrada] + [e for e in extra if isinstance(e, Path)], "nomoffice/entrada")
    args = [f"entrada\\{entrada.name}", salida] + [f"entrada\\{e.name}" if isinstance(e, Path) else e for e in extra]
    remoto.correr_office(Path(salida).stem, script, *args, timeout=timeout)
    remoto.bajar(f"nomoffice/{salida}", SAL / salida)
    return json.loads((SAL / salida).read_text(encoding="utf-8"))


def bajar(p, nombre):
    destino = SAL / nombre
    with p.expect_download(timeout=120000) as dl:
        p.evaluate("(n) => NOM.formatos.exportarListo(NOM.editorActual.doc, 'xlsx', n)", destino.stem)
    dl.value.save_as(destino)
    return destino


JS_VER = """(r) => { const hex = (s) => { const m = /rgba?\\((\\d+), (\\d+), (\\d+)(?:, ([\\d.]+))?/.exec(s || ''); if (!m || m[4] === '0') return null; return '#' + [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase(); };
  const out = {}; for (let f = r.f1; f <= r.f2; f++) for (let c = r.c1; c <= r.c2; c++) { const td = document.querySelector(`td[data-c="${c}"][data-f="${f}"]`); if (!td) continue; const s = getComputedStyle(td);
    out[c + ',' + f] = { t: td.textContent, f: hex(s.backgroundColor) || '#FFFFFF', k: hex(s.color), b: td.hasAttribute('data-negrita') || +s.fontWeight >= 600, ind: /padding-left/.test(td.getAttribute('style') || '') ? Math.round((parseFloat(s.paddingLeft) - 2) / 9) : 0 }; } return out; }"""


def main():
    SAL.mkdir(parents=True, exist_ok=True)
    filas = json.loads((SAL / "d01-ventas.json").read_text(encoding="utf-8"))
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    p.set_viewport_size({"width": 1800, "height": 1200})
    try:
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'dinamicas', modo: 'HOJA', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.api", timeout=60000); p.wait_for_timeout(800)
        p.evaluate("""([campos, filas]) => { const t = NOM.editorActual.modoApi.api.tabla(), c = {}, L = (k) => NOM.formulas.colLetra(k);
          campos.forEach((h, k) => (c[`c:${L(k)}1`] = h)); filas.forEach((r, i) => r.forEach((v, k) => (c[`c:${L(k)}${i + 2}`] = String(v)))); c.filas = 40; c.cols = 10; c.nombre = 'Datos'; NOM.editorActual.doc.cambiar(t.id, c); }""", [CAMPOS, filas])
        p.wait_for_timeout(500)
        datos = bajar(p, "d04-datos-nom.xlsx")
        casos_x = [{"id": cs["id"], "origen": "A1:E31", "destino": "H5" if cs.get("filtro") else "H3", "filas": cs["filas"], "columnas": cs.get("columnas", []), "valores": cs["valores"], "filtro": cs.get("filtro"), "estilo": cs.get("estilo")} for cs in CASOS]
        (SAL / "casos-prueba.json").write_text(json.dumps(casos_x, ensure_ascii=False), encoding="utf-8")
        verdad = remoto_excel("dinamica_excel.py", datos, "verdad-prueba.json", SAL / "casos-prueba.json")["casos"]
        # NOM: cada una en una hoja nueva
        ci = CAMPOS.index
        hechas = {}
        for cs in CASOS:
            p.evaluate("() => { const ed = NOM.editorActual; const t0 = NOM.formulas.hojasDe(ed.doc).find((h) => NOM.formulas.nombreDeHoja(ed.doc, h) === 'Datos'); ed.pagina = t0.pagina; ed.modoApi.render && ed.modoApi.render(); }")
            p.wait_for_timeout(300)
            r = p.evaluate("""(d) => { const x = NOM.editorActual.modoApi.api.dinamicas.crear(d); return typeof x === 'string' ? x : { n: x.n, area: x.area }; }""",
                           {"origen": {"hoja": "Datos", "c1": 0, "f1": 0, "c2": 4, "f2": 30}, "destino": "nueva", "filas": [ci(n) for n in cs["filas"]], "columnas": [ci(n) for n in cs.get("columnas", [])],
                            "valores": [{"c": ci(n), "fn": fn} for n, fn in cs["valores"]], "filtro": {"c": ci(cs["filtro"]), "sel": None} if cs.get("filtro") else None, "estilo": cs.get("estilo")})
            p.wait_for_timeout(600)
            if isinstance(r, str): hechas[cs["id"]] = r; continue
            hechas[cs["id"]] = {"area": r["area"], "vista": p.evaluate(JS_VER, r["area"])}
        anota("NOM hizo las seis tablas dinámicas", all(isinstance(v, dict) for v in hechas.values()), {k: v for k, v in hechas.items() if not isinstance(v, dict)})
        mal_t, mal_f, n_c = [], [], 0
        for cs in CASOS:
            h, V = hechas.get(cs["id"]), verdad.get(cs["id"], {})
            if not isinstance(h, dict) or "celdas" not in V: mal_t.append((cs["id"], "falta", V.get("error"))); continue
            dc = rc(V["rango"].split(":")[0])[0] - h["area"]["c1"]   # (Excel en H, NOM en A de una hoja nueva)
            for a, e in V["celdas"].items():
                n_c += 1
                cc, ff = rc(a); m = h["vista"].get(f"{cc - dc},{ff}")
                if not m: mal_t.append((cs["id"], a, "sin celda")); continue
                if en(m["t"]) != e["t"]: mal_t.append((cs["id"], a, m["t"], e["t"]))
                tinta = lambda k: "#000000" if k in (None, "#1B1C26") else k
                if m["f"] != e["f"] or tinta(m["k"]) != tinta(e["k"]) or m["b"] != e["b"] or m["ind"] != e["ind"]: mal_f.append((cs["id"], a, m, e))
        anota(f"cada celda dice lo mismo que en Excel ({n_c} celdas)", not mal_t and n_c > 100, f"{len(mal_t)}: {mal_t[:8]}")
        anota("con la misma negrita, relleno, letra y sangría", not mal_f, f"{len(mal_f)}: {mal_f[:6]}")
        v = remoto_excel("dinamicas_ver.py", bajar(p, "dinamicas-nom.xlsx"), "dinamicas-nom.dv.json")
        pts = [x for hs in v.get("hojas", {}).values() for x in hs]
        anota("Excel ve las seis tablas dinámicas con sus campos", len(pts) == 6 and all(x["valores"] for x in pts), [v.get("error"), [(x["nombre"], x["rango"], x["filas"], x["valores"]) for x in pts]])
        anota("al actualizarlas Excel, no cambian", pts and all(x.get("despues") == {k: en(t) for k, t in x["antes"].items()} or x.get("despues") == x["antes"] for x in pts),
              [(x["nombre"], x.get("error_actualizar"), [(k, x["antes"].get(k), t) for k, t in (x.get("despues") or {}).items() if en(x["antes"].get(k)) != t][:4]) for x in pts])
        # en un libro de Excel con tablas dinámicas
        x05 = REAL / "archivos/x05-tablas-dinamicas.xlsx"
        antes = remoto_excel("dinamicas_ver.py", x05, "x05-antes.dv.json")
        n_antes = sum(len(x) for x in antes["hojas"].values())
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(x05.read_bytes()).decode(), x05.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
        hoja = p.evaluate("() => NOM.formulas.nombreDeHoja(NOM.editorActual.doc, NOM.editorActual.modoApi.api.tabla())")
        r3 = p.evaluate("""() => { const x = NOM.editorActual.modoApi.api.dinamicas.crear({ origen: { c1: 0, f1: 0, c2: 6, f2: 60 }, destino: { c: 12, f: 40 }, filas: [1], valores: [{ c: 4, fn: 'sum' }] }); return typeof x === 'string' ? x : x.area; }""")
        v2 = remoto_excel("dinamicas_ver.py", bajar(p, "x05-con-dinamica.xlsx"), "x05-despues.dv.json")
        n_despues = sum(len(x) for x in v2["hojas"].values())
        anota(f"en x05 (con {n_antes} tablas dinámicas): Excel ve las de antes y la nueva", n_despues == n_antes + 1, [r3, hoja, n_antes, n_despues, v2.get("error")])
        import office_abre
        r = office_abre.abrir([SAL / "dinamicas-nom.xlsx", SAL / "x05-con-dinamica.xlsx"], "dinamicas")
        for a in ("dinamicas-nom.xlsx", "x05-con-dinamica.xlsx"): anota(f"Excel abre {a} sin avisos ni reparar", office_abre.bien(r.get(a)), json.dumps(r.get(a), ensure_ascii=False)[:300])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nTablas dinámicas contra Excel 2024: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

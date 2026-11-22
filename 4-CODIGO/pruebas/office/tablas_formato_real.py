#!/usr/bin/env python3
"""«Dar formato como tabla» hecho en NOM contra Excel 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/tablas_formato_real.py

1. En un libro de NOM se hace una tabla con cada uno de los 60 estilos de Excel (Claro 1-21, Medio 1-28, Oscuro 1-11) y
   otras con las opciones (fila de totales, primera y última columna, columnas con bandas, sin filas con bandas, sin
   encabezados). Descargado, Excel ve cada una con su nombre, rango, estilo, opciones y columnas (windows/tablas_formato_ver.py).
2. Lo que se ve: el relleno, el color de la letra y la negrita de cada celda en NOM, igual que en Excel (DisplayFormat).
3. Convertir en rango deja el formato en las celdas; insertar una fila dentro agranda la tabla.
4. En un libro de Excel que ya trae tablas (t02-todo-excel.xlsx), una tabla hecha en NOM se suma a las de antes.
5. Excel abre todo sin avisos.
"""
import base64, json, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
SAL = Path.home() / "NOM/trabajo/tablas-formato-real"
res = []
ESTILOS = [f"TableStyleLight{i}" for i in range(1, 22)] + [f"TableStyleMedium{i}" for i in range(1, 29)] + [f"TableStyleDark{i}" for i in range(1, 12)]
# (estilo, opciones) de las tablas con opciones
VARIANTES = [("TableStyleMedium2", {"tot": 1}), ("TableStyleMedium9", {"primeraC": True, "ultimaC": True}), ("TableStyleLight9", {"rayasC": True, "rayasF": False}),
             ("TableStyleDark2", {"tot": 1, "ultimaC": True}), ("TableStyleLight15", {"primeraC": True, "rayasC": True}), ("TableStyleMedium15", {"filtro": False})]


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:900]}"), flush=True)


def excel_ver(xlsx):
    import remoto
    remoto.subir([xlsx], "nomoffice/entrada")
    remoto.correr_office("tf-" + xlsx.stem, "tablas_formato_ver.py", f"entrada\\{xlsx.name}", f"{xlsx.stem}.tf.json", timeout=900)
    tmp = SAL / f"{xlsx.stem}.tf.json"; remoto.bajar(f"nomoffice/{xlsx.stem}.tf.json", tmp)
    return json.loads(tmp.read_text(encoding="utf-8"))


def bajar(p, nombre):
    destino = SAL / nombre
    with p.expect_download(timeout=120000) as dl:
        p.evaluate("(n) => NOM.formatos.exportarListo(NOM.editorActual.doc, 'xlsx', n)", destino.stem)
    dl.value.save_as(destino)
    return destino


JS_DATOS = """(n) => { const t = NOM.editorActual.modoApi.api.tabla(), c = {}, L = (k) => NOM.formulas.colLetra(k);
  const filas = Math.ceil(n / 6) * 6 + 2;
  for (let i = 0; i < n; i++) { const c0 = (i % 6) * 4, f0 = Math.floor(i / 6) * 6;
    ['Nombre', 'Zona', 'Ventas'].forEach((h, k) => (c[`c:${L(c0 + k)}${f0 + 1}`] = h));
    for (let r = 0; r < 3; r++) { c[`c:${L(c0)}${f0 + 2 + r}`] = 'P' + (i * 3 + r); c[`c:${L(c0 + 1)}${f0 + 2 + r}`] = ['Norte', 'Sur', 'Este'][r]; c[`c:${L(c0 + 2)}${f0 + 2 + r}`] = String(10 * (r + 1) + i); } }
  c.filas = filas; c.cols = 26; NOM.editorActual.doc.cambiar(t.id, c); return filas; }"""
# lo que se ve en NOM de cada celda de un rango: [relleno, letra, negrita]
JS_VER = """(r) => { const hex = (s) => { const m = /rgba?\\((\\d+), (\\d+), (\\d+)(?:, ([\\d.]+))?/.exec(s || ''); if (!m || m[4] === '0') return null; return '#' + [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase(); };
  const out = {}; for (let f = r.f1; f <= r.f2; f++) for (let c = r.c1; c <= r.c2; c++) { const td = document.querySelector(`td[data-c="${c}"][data-f="${f}"]`); if (!td) continue; const s = getComputedStyle(td);
    out[NOM.formulas.ref(c, f)] = [hex(s.backgroundColor) || '#FFFFFF', hex(s.color), td.hasAttribute('data-negrita') || +s.fontWeight >= 600]; } return out; }"""


def main():
    SAL.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    p.set_viewport_size({"width": 2400, "height": 1600})
    try:
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'tablas', modo: 'HOJA', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.api", timeout=60000); p.wait_for_timeout(800)
        n = len(ESTILOS) + len(VARIANTES)
        p.evaluate(JS_DATOS, n); p.wait_for_timeout(500)
        hechas = []
        for i, (est, op) in enumerate([(e, {}) for e in ESTILOS] + VARIANTES):
            c0, f0 = (i % 6) * 4, (i // 6) * 6
            r = p.evaluate("([est, r]) => { const x = NOM.editorActual.modoApi.api.tablasFormato.crear(est, { rango: r }); return typeof x === 'string' ? x : x.id; }", [est, {"c1": c0, "f1": f0, "c2": c0 + 2, "f2": f0 + 3}])
            if op:
                r2 = p.evaluate("([id, op]) => { const x = NOM.editorActual.modoApi.api.tablasFormato.opciones(id, op); return typeof x === 'string' ? x : 'ok'; }", [r, op])
                if r2 != "ok": print("   opciones:", est, op, r2)
            hechas.append((est, op, r))
        p.wait_for_timeout(1200)
        los = p.evaluate("() => NOM.editorActual.modoApi.api.tablasFormato.lista()")
        anota(f"NOM hizo las {n} tablas", len(los) == n, [x for x in hechas if not isinstance(x[2], str) or len(x[2]) < 4][:5])
        p.locator('td[data-c="25"][data-f="0"]').scroll_into_view_if_needed(); p.locator('td[data-c="25"][data-f="0"]').click(); p.wait_for_timeout(400)
        p.evaluate("() => { const r = document.querySelector('.rejilla-hoja'); if (r) { r.scrollLeft = 0; r.scrollTop = 0; } }"); p.wait_for_timeout(500)
        vista_nom = {}
        for lo in los:
            p.locator(f'td[data-c="{lo["c1"]}"][data-f="{lo["f1"]}"]').scroll_into_view_if_needed(); p.wait_for_timeout(60)
            vista_nom[lo["n"]] = p.evaluate(JS_VER, {"c1": lo["c1"], "f1": lo["f1"], "c2": lo["c2"], "f2": lo["f2"]})
        v = excel_ver(bajar(p, "tablas-nom.xlsx"))
        xl = {lo["nombre"]: lo for lo in list(v.get("hojas", {}).values())[0]} if v.get("hojas") else {}
        anota(f"Excel ve las {n} tablas", len(xl) == n, [v.get("error"), len(xl)])
        mal_meta, mal_celdas, celdas = [], [], 0
        for lo in los:
            e = xl.get(lo["n"])
            if not e: mal_meta.append((lo["n"], "falta")); continue
            esperado = {"estilo": lo["estilo"], "enc": bool(lo["enc"]), "tot": bool(lo["tot"]), "rayasF": lo.get("rayasF") is not False, "rayasC": bool(lo.get("rayasC")),
                        "primeraC": bool(lo.get("primeraC")), "ultimaC": bool(lo.get("ultimaC")), "filtro": lo.get("filtro") is not False}
            dif = {k: (val, e[k]) for k, val in esperado.items() if e[k] != val}
            rng = f"{N_ref(lo['c1'], lo['f1'])}:{N_ref(lo['c2'], lo['f2'])}"
            if e["rango"] != rng: dif["rango"] = (rng, e["rango"])
            if e["cols"] != ["Nombre", "Zona", "Ventas"]: dif["cols"] = e["cols"]
            if dif: mal_meta.append((lo["n"], dif))
            for ref, (fo, le, b) in e["celdas"].items():
                celdas += 1
                mio = vista_nom[lo["n"]].get(ref)
                # (la letra «automática» de un libro de NOM es la tinta de NOM, #1B1C26, en todas sus celdas; Excel la da negra)
                tinta = lambda k: "#000000" if k in (None, "#1B1C26") else k
                if not mio or mio[0] != fo or tinta(mio[1]) != tinta(le) or bool(mio[2]) != b:
                    mal_celdas.append((lo["n"], lo["estilo"], ref, "NOM", mio, "Excel", [fo, le, b]))
        anota("cada tabla con su rango, estilo, opciones y columnas", not mal_meta, mal_meta[:6])
        anota(f"lo que se ve: relleno, letra y negrita de las {celdas} celdas, como Excel", not mal_celdas and celdas > 0, f"{len(mal_celdas)} distintas: {mal_celdas[:8]}")
        # convertir en rango e insertar una fila dentro
        uno = los[0]
        p.evaluate("(id) => NOM.editorActual.modoApi.api.tablasFormato.convertir(id)", uno["id"]); p.wait_for_timeout(500)
        tv = p.evaluate("(r) => NOM.editorActual.modoApi.api.tabla()['r:' + r]", N_ref(uno["c1"], uno["f1"]))
        anota("convertir en rango: ya no es tabla y el encabezado conserva su relleno", len(p.evaluate("() => NOM.editorActual.modoApi.api.tablasFormato.lista()")) == n - 1 and (tv or "").upper() == xl[uno["n"]]["celdas"][N_ref(uno["c1"], uno["f1"])][0], tv)
        dos = los[1]
        p.locator(f'td[data-c="{dos["c1"]}"][data-f="{dos["f1"] + 1}"]').click(); p.wait_for_timeout(200)
        p.evaluate("() => NOM.editorActual.modoApi.api.insertarFilas ? NOM.editorActual.modoApi.api.insertarFilas(false) : null")
        p.wait_for_timeout(500)
        d2 = [x for x in p.evaluate("() => NOM.editorActual.modoApi.api.tablasFormato.lista()") if x["id"] == dos["id"]]
        anota("insertar una fila dentro agranda la tabla", d2 and d2[0]["f2"] == dos["f2"] + 1, d2)
        # en un libro de Excel con tablas
        t02 = REAL / "archivos/t02-todo-excel.xlsx"
        antes = excel_ver(t02)
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(t02.read_bytes()).decode(), t02.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
        hoja = list(antes["hojas"].keys())[0]
        p.evaluate("""() => { const t = NOM.editorActual.modoApi.api.tabla(), c = {}; ['Clave', 'Monto'].forEach((h, k) => (c[`c:${NOM.formulas.colLetra(30 + k)}101`] = h));
          for (let r = 0; r < 3; r++) { c[`c:AE${102 + r}`] = 'K' + r; c[`c:AF${102 + r}`] = String(5 * r); } if ((t.filas || 0) < 110) c.filas = 110; if ((t.cols || 0) < 34) c.cols = 34; NOM.editorActual.doc.cambiar(t.id, c); }""")
        p.wait_for_timeout(500)
        r3 = p.evaluate("() => { const x = NOM.editorActual.modoApi.api.tablasFormato.crear('TableStyleMedium4', { rango: { c1: 30, f1: 100, c2: 31, f2: 103 } }); return typeof x === 'string' ? x : x.n; }")
        v2 = excel_ver(bajar(p, "t02-con-tabla.xlsx"))
        nuevas = v2["hojas"].get(hoja, [])
        n_antes = sum(len(x) for x in antes["hojas"].values()); n_despues = sum(len(x) for x in v2["hojas"].values())
        anota(f"en t02 (con {n_antes} tablas): Excel ve las de antes y la nueva en AE101:AF104", n_despues == n_antes + 1 and any(x["rango"] == "AE101:AF104" and x["estilo"] == "TableStyleMedium4" for x in nuevas), [r3, n_antes, n_despues, [x["rango"] for x in nuevas]])
        import office_abre
        r = office_abre.abrir([SAL / "tablas-nom.xlsx", SAL / "t02-con-tabla.xlsx"], "tablas-formato")
        for a in ("tablas-nom.xlsx", "t02-con-tabla.xlsx"): anota(f"Excel abre {a} sin avisos ni reparar", office_abre.bien(r.get(a)), json.dumps(r.get(a), ensure_ascii=False)[:300])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nDar formato como tabla contra Excel 2024: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


def N_ref(c, f):
    s, c = "", c + 1
    while c: c, m = divmod(c - 1, 26); s = chr(65 + m) + s
    return f"{s}{f + 1}"


if __name__ == "__main__":
    sys.exit(main())

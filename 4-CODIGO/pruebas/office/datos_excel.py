#!/usr/bin/env python3
"""Herramientas de datos de la hoja (quitar duplicados, texto en columnas, buscar objetivo) contra Excel 2024 real.

  datos_excel.py [--solo-excel]

1. Arma los casos (CASOS) y los manda a la computadora con Windows: windows/datos_excel.py los hace en Excel y
   devuelve cada celda (valor y texto que se ve). La verdad queda en ~/.cache/nom-pruebas/office-real/datos-excel.json.
2. Hace lo mismo en NOM (la API de la hoja: NOM.editorActual.hojaApi.herramientas) y compara celda por celda.
"""
import json, os, sys, tempfile
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import remoto

CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas"))
VERDAD = CACHE / "office-real/datos-excel.json"

NOMBRES = [["Nombre", "Grupo", "Cal"], ["Ana", "A", 9], ["Luis", "B", 7.5], ["ana", "A", 9], ["Eva", "A", 10], ["ANA", "B", 9],
           ["Luis", "B", 7.5], ["Eva ", "A", 10], ["Juan", "", 6], ["Juan", "", 6], ["1", "C", 1], [1, "C", 1]]
CASOS = [
    {"id": "dup-todas", "op": "quitar_duplicados", "datos": NOMBRES, "rango": "A1:C12", "columnas": [1, 2, 3], "encabezado": True},
    {"id": "dup-col1", "op": "quitar_duplicados", "datos": NOMBRES, "rango": "A1:C12", "columnas": [1], "encabezado": True},
    {"id": "dup-col12", "op": "quitar_duplicados", "datos": NOMBRES, "rango": "A1:C12", "columnas": [1, 2], "encabezado": True},
    {"id": "dup-sin-encabezado", "op": "quitar_duplicados", "datos": NOMBRES[1:], "rango": "A1:C11", "columnas": [1, 3], "encabezado": False},
    {"id": "txt-coma", "op": "texto_columnas", "rango": "A1:A6", "sep": {"coma": True},
     "datos": [["a,b,c"], ["1,2,3"], ["x,,y"], ["solo"], ["3.5,007,-5"], ["\"Pérez, Ana\",20,MX"]]},
    {"id": "txt-punto-coma", "op": "texto_columnas", "rango": "A1:A4", "sep": {"punto_coma": True},
     "datos": [["uno;dos;tres"], ["'a;b';c"], ["10%;$25;1e3"], ["TRUE;FALSE;verdadero"]]},
    {"id": "txt-espacio-seguidos", "op": "texto_columnas", "rango": "A1:A3", "sep": {"espacio": True}, "seguidos": True,
     "datos": [["Ana   Pérez  López"], ["  inicio con espacios"], ["5-  (7)  2026-03-12"]]},
    {"id": "txt-espacio", "op": "texto_columnas", "rango": "A1:A2", "sep": {"espacio": True}, "seguidos": False,
     "datos": [["a  b"], [" c d"]]},
    {"id": "txt-otro", "op": "texto_columnas", "rango": "A1:A3", "sep": {"otro": "|"},
     "datos": [["2026-11-18|2026-01-05|1e-2"], ["a|b"], ["|vacío|"]]},
    {"id": "txt-anchos", "op": "texto_columnas", "rango": "A1:A3", "anchos": [0, 4, 9],
     "datos": [["ABCD12345resto"], ["XY  9    final"], ["corto"]]},
    # Consolidar: dos tablas con rótulos (filas y columnas en otro orden, una con algo que la otra no tiene) y por posición
    {"id": "cons-rotulos", "op": "consolidar", "destino": "I1", "fuentes": ["A1:C4", "E1:G5"], "funcion": "suma", "fila": True, "columna": True,
     "datos": [["", "Ene", "Feb", None, "", "Feb", "Mar"], ["Ana", 1, 2, None, "Luis", 5, 6], ["Luis", 3, 4, None, "ana", 7, 8], ["Eva", 9, "x", None, "Juan", 1, 1], [None, None, None, None, "Ana", 2, 2]]},
    {"id": "cons-posicion", "op": "consolidar", "destino": "A8", "fuentes": ["A1:B3", "D1:F2"], "funcion": "promedio",
     "datos": [[1, 2, None, 10, 20, 30], [3, 4, None, 40, 50, 60], [5, "t", None, None, None, None]]},
    {"id": "cons-contar", "op": "consolidar", "destino": "H1", "fuentes": ["A1:B4", "D1:E3"], "funcion": "contar", "columna": True,
     "datos": [["a", 1, None, "b", 2], ["b", "x", None, "a", None], ["a", 3, None, "c", 4], ["c", None, None, None, None]]},
    {"id": "cons-max", "op": "consolidar", "destino": "A6", "fuentes": ["A1:C3", "E1:G3"], "funcion": "max", "fila": True,
     "datos": [["x", "y", "z", None, "z", "x", "w"], [1, 5, 2, None, 9, 0, 4], [7, -1, 3, None, -2, 8, 6]]},
    # (el orden de los rótulos: casos para sacar la regla de Excel)
    {"id": "cons-orden-1", "op": "consolidar", "destino": "A6", "fuentes": ["A1:B3", "D1:E3"], "funcion": "suma", "fila": True, "columna": True,
     "datos": [["", "x", None, "", "x"], ["a", 1, None, "c", 1], ["b", 1, None, "d", 1]]},
    {"id": "cons-orden-2", "op": "consolidar", "destino": "A6", "fuentes": ["A1:B3", "D1:E3"], "funcion": "suma", "fila": True, "columna": True,
     "datos": [["", "x", None, "", "x"], ["a", 1, None, "b", 1], ["b", 1, None, "c", 1]]},
    {"id": "cons-orden-3", "op": "consolidar", "destino": "A6", "fuentes": ["A1:B3", "D1:E3"], "funcion": "suma", "fila": True, "columna": True,
     "datos": [["", "x", None, "", "x"], ["a", 1, None, "c", 1], ["b", 1, None, "a", 1]]},
    {"id": "cons-orden-4", "op": "consolidar", "destino": "A6", "fuentes": ["A1:B3", "D1:E3"], "funcion": "suma", "columna": True,
     "datos": [["a", 1, None, "c", 1], ["b", 1, None, "a", 1], [None, None, None, None, None]]},
    {"id": "cons-orden-5", "op": "consolidar", "destino": "A6", "fuentes": ["A1:C2", "E1:G2"], "funcion": "suma", "fila": True, "columna": True,
     "datos": [["", "x", "y", None, "", "z", "x"], ["a", 1, 1, None, "a", 1, 1]]},
    {"id": "cons-orden-6", "op": "consolidar", "destino": "A8", "fuentes": ["A1:B4", "D1:E4"], "funcion": "suma", "fila": True, "columna": True,
     "datos": [["", "x", None, "", "x"], ["a", 1, None, "d", 1], ["b", 1, None, "a", 1], ["c", 1, None, "e", 1]]},
    {"id": "obj-lineal", "op": "buscar_objetivo", "celda": "B2", "valor": 100, "cambiando": "B1", "datos": [["x", 1], ["2x+3", "=B1*2+3"]]},
    {"id": "obj-raiz", "op": "buscar_objetivo", "celda": "B2", "valor": 2, "cambiando": "B1", "datos": [["x", 1], ["x^2", "=B1^2"]]},
    {"id": "obj-pago", "op": "buscar_objetivo", "celda": "B2", "valor": 2000, "cambiando": "B1", "datos": [["tasa", 0.05], ["pago", "=PMT(B1/12,60,-100000)"]]},
    {"id": "obj-imposible", "op": "buscar_objetivo", "celda": "B2", "valor": -1, "cambiando": "B1", "datos": [["x", 3], ["x^2", "=B1^2"]]},
    # Tabla de datos (Datos → Análisis de hipótesis → Tabla de datos), hecha en NOM y en Excel (Range.Table)
    {"id": "td-columna", "op": "tabla_datos", "rango": "A4:C8", "col": "B1",
     "datos": [["x", 2], [], [], ["", "=B1*10", "=B1^2+1"], [1], [2.5], [-3], [10]]},
    {"id": "td-fila", "op": "tabla_datos", "rango": "A4:E6", "fila": "B1",
     "datos": [["x", 2], [], [], ["", 1, 2, 3, 4], ["=B1*10"], ["=B1/4+0.5"]]},
    {"id": "td-dos", "op": "tabla_datos", "rango": "A4:D7", "fila": "B1", "col": "B2",
     "datos": [["x", 2], ["y", 3], [], ["=B1*B2+1", 1, 2, 3], [10], [20], [30]]},
]


def verdad_excel():
    with tempfile.TemporaryDirectory() as d:
        ent = Path(d) / "datos-casos.json"
        ent.write_text(json.dumps(CASOS, ensure_ascii=False), encoding="utf-8")
        remoto.subir([ent], "nomoffice")
        r = remoto.correr_office("datos", "datos_excel.py", "datos-casos.json", "datos-verdad.json", timeout=900)
        print(r.stdout[-800:])
        if not remoto.bajar("nomoffice/datos-verdad.json", VERDAD): raise SystemExit("Excel no dejó el resultado")
    return json.load(open(VERDAD, encoding="utf-8"))


JS_CASO = r"""
async (c) => {
  const ed = NOM.editorActual, H = ed.modoApi.api, t = H.tabla(), F = NOM.formulas;
  const campos = { filas: Math.max(30, c.datos.length + 5), cols: 10 };
  for (let i = 0; i < 40; i++) for (let j = 0; j < 10; j++) campos["c:" + F.ref(j, i)] = null;
  c.datos.forEach((fila, i) => fila.forEach((v, j) => { if (v !== null && v !== "") campos["c:" + F.ref(j, i)] = String(v); }));
  ed.doc.cambiar(t.id, campos);
  await new Promise((r) => setTimeout(r, 150));
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(c.rango || "A1:A1");
  const col = (x) => [...x].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const r = { c1: col(m[1]), f1: +m[2] - 1, c2: col(m[3]), f2: +m[4] - 1 };
  let res = null;
  if (c.op === "quitar_duplicados") res = H.herramientas.quitarDuplicados(r, c.columnas.map((k) => r.c1 + k - 1), !!c.encabezado);
  else if (c.op === "texto_columnas") res = H.herramientas.textoEnColumnas(r, c.anchos ? { anchos: c.anchos } : { sep: c.sep, seguidos: !!c.seguidos, calificador: c.calificador == null ? '"' : c.calificador });
  else if (c.op === "consolidar") res = H.herramientas.consolidar(c.destino, c.fuentes, c.funcion, !!c.fila, !!c.columna);
  else if (c.op === "tabla_datos") res = H.herramientas.tablaDatos(r, c.fila || "", c.col || "");
  else if (c.op === "buscar_objetivo") {
    const o = H.herramientas.buscarObjetivo(c.celda, c.valor, c.cambiando);
    await new Promise((r) => setTimeout(r, 150));
    return { objetivo: typeof o === "string" ? { error: o } : { logrado: o.logrado, x: o.x, valor: o.valor } };
  }
  else return { no: true };
  await new Promise((r) => setTimeout(r, 150));
  const ev = H.evaluar(), out = [], vals = [];
  for (let f = 0; f < c.filasVer; f++) {
    const fila = [], fv = [];
    for (let k = 0; k < c.colsVer; k++) {
      const td = document.querySelector(`td[data-c="${k}"][data-f="${f}"]`);
      fila.push(td ? td.textContent.trim() : F.formatear(ev.valor(k, f)));
      const v = ev.valor(k, f); fv.push(v == null || v === "" ? null : typeof v === "object" && v.valueOf ? +v : v);
    }
    out.push(fila); vals.push(fv);
  }
  return { res, celdas: out, valores: vals };
}
"""


def con_nom(verdad):
    sys.path.insert(0, str(AQUI.parent / "interfaz"))
    import nomprueba as N
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Datos")
    nav = N.Navegador(movil=False); nav.entrar(a)
    bien = mal = 0
    try:
        for c in CASOS:
            v = verdad.get(c["id"], {})
            if "celdas" not in v: continue
            d = N.documento(a["token"], esp, "HOJA", c["id"])
            nav.p.goto(N.B + "/#/doc/" + d); nav.p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.api", timeout=30000)
            nav.p.wait_for_timeout(500)
            filas, cols = len(v["celdas"]), max(len(x) for x in v["celdas"])
            r = nav.p.evaluate(JS_CASO, dict(c, filasVer=filas, colsVer=cols))
            if r.get("no"): continue
            if "objetivo" in r:
                # Buscar objetivo: Excel no publica su método y su valor final difiere en la sexta cifra; se compara lo que
                # Excel garantiza: si encontró solución y que la fórmula quede a menos de 0.001 (su «cambio máximo»)
                o = r["objetivo"]
                igual = "error" not in o and bool(o.get("logrado")) == bool(v.get("logrado")) and (
                    not v.get("logrado") or (abs(o["valor"] - c["valor"]) < 0.001 and abs(v["resultado"] - c["valor"]) < 0.001))
                bien += igual; mal += not igual
                print(("  bien " if igual else "  MAL  ") + c["id"] + f"  (Excel: {'logrado' if v.get('logrado') else 'sin solución'}, x={v.get('cambio')}; "
                      f"NOM: {'logrado' if o.get('logrado') else 'sin solución'}, x={o.get('x')}, valor={o.get('valor')})", flush=True)
                continue
            xl = [[k["t"] for k in fila] for fila in v["celdas"]]
            # (donde Excel muestra #### la columna no le cabe: se compara el valor)
            def igual_celda(i, j):
                x, n = v["celdas"][i][j], r["celdas"][i][j]
                # (y las fechas: la PC de prueba muestra mes/día/año —Excel en inglés de EE. UU.— y NOM día/mes/año, como
                # en México; se compara la fecha misma)
                fecha = isinstance(x["v"], (int, float)) and "/" in x["t"]
                if fecha or (x["t"].startswith("#") and set(x["t"]) == {"#"}):
                    a, b = x["v"], r["valores"][i][j]
                    return a == b or (isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) < 1e-9)
                return x["t"].strip() == n.strip()
            igual = all(igual_celda(i, j) for i in range(len(xl)) for j in range(len(xl[i])))
            bien += igual; mal += not igual
            print(("  bien " if igual else "  MAL  ") + c["id"] + ("" if igual else f"\n     Excel {xl}\n     NOM   {r['celdas']}"), flush=True)
    finally:
        nav.cerrar()
    print(f"\nHerramientas de datos contra Excel 2024: {bien} bien, {mal} mal")
    return mal


def main():
    if "--solo-nom" in sys.argv and VERDAD.exists():
        return 1 if con_nom(json.load(open(VERDAD, encoding="utf-8"))) else 0
    v = verdad_excel()
    for c in CASOS:
        x = v.get(c["id"], {})
        print(f"== {c['id']}" + (f"  ERROR {x['error']}" if "error" in x else ""))
        for fila in x.get("celdas", []): print("   ", [f"{k['t']}" for k in fila])
        if "cambio" in x: print("    cambio:", x["cambio"], "resultado:", x.get("resultado"), "logrado:", x.get("logrado"))
    if "--solo-excel" in sys.argv: return 0
    return 1 if con_nom(v) else 0


if __name__ == "__main__":
    sys.exit(main())

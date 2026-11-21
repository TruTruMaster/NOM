#!/usr/bin/env python3
"""Herramientas de datos de la hoja (quitar duplicados, texto en columnas, buscar objetivo, consolidar, tablas de datos,
relleno rápido) contra Excel 2024 real.

  datos_excel.py [--solo-excel | --solo-nom] [--casos PREFIJO[,PREFIJO…]]

1. Arma los casos (CASOS) y los manda a la computadora con Windows: windows/datos_excel.py los hace en Excel y
   devuelve cada celda (valor y texto que se ve). La verdad queda en ~/.cache/nom-pruebas/office-real/datos-excel.json.
2. Hace lo mismo en NOM (la API de la hoja: NOM.editorActual.hojaApi.herramientas) y compara celda por celda.
"""
import json, os, re, sys, tempfile
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


def _ff(id_, entradas, ejemplos, **kw):
    """Un caso de relleno rápido: las columnas de entrada (una lista por fila) y, en la columna siguiente, los ejemplos de
    las primeras filas; Excel (Range.FlashFill) y NOM llenan el resto de esa columna."""
    n = len(entradas[0]) if entradas and isinstance(entradas[0], list) else 1
    filas = [(list(e) if isinstance(e, list) else [e]) + ([ejemplos[i]] if i < len(ejemplos) else []) for i, e in enumerate(entradas)]
    col = chr(65 + n)
    return dict({"id": id_, "op": "relleno_rapido", "datos": filas, "rango": f"{col}1:{col}{len(entradas)}"}, **kw)


NOMBRES_FF = ["Ana Pérez López", "Luis Gómez", "María José Ruiz Díaz", "Pedro Sánchez", "Sofía Hernández Mora", "Jorge Luis Borges"]
CORREOS_FF = ["ana.perez@correo.mx", "luis_gomez@uni.edu.mx", "maria@x.com", "j.ruiz@empresa.com.mx"]
CASOS += [
    _ff("ff-nombre", NOMBRES_FF, ["Ana"]),
    _ff("ff-apellido-ultimo", NOMBRES_FF, ["López", "Gómez"]),
    _ff("ff-apellido-segundo", NOMBRES_FF, ["Pérez"]),
    _ff("ff-iniciales", ["Ana Pérez", "Luis Gómez", "María Ruiz", "Pedro Sánchez"], ["A. Pérez"]),
    _ff("ff-iniciales-tres", ["Ana María Pérez", "Luis Felipe Gómez", "Rosa Elena Ruiz"], ["AMP"]),
    _ff("ff-mayusculas", ["ana pérez", "luis gómez", "maría ruiz", "josé de la o"], ["ANA PÉREZ"]),
    _ff("ff-minusculas", ["ANA PÉREZ", "Luis Gómez", "MaRía RuIz"], ["ana pérez"]),
    # (a propósito distinto: Excel 2024 no cambia las mayúsculas a «Nombre Propio» y arma «Aluis Pgómez» y «Aaría P ruiz»
    # con letras fijas; NOM da lo que se pidió)
    _ff("ff-propio", ["aNA péREZ", "LUIS GÓMEZ", "maría ruiz"], ["Ana Pérez"],
        distinto=[["aNA péREZ", "Ana Pérez"], ["LUIS GÓMEZ", "Luis Gómez"], ["maría ruiz", "María Ruiz"]]),
    _ff("ff-unir-coma", [["Ana", "Pérez"], ["Luis", "Gómez"], ["María", "Ruiz"]], ["Pérez, Ana"]),
    _ff("ff-correo", [["Ana", "Lopez"], ["Luis", "Gomez"], ["Maria", "Ruiz"]], ["ana.lopez@empresa.mx"]),
    _ff("ff-usuario", CORREOS_FF, ["ana.perez"]),
    _ff("ff-dominio", CORREOS_FF, ["correo.mx"]),
    _ff("ff-numero", ["Pedido 1234 enviado", "Pedido 56 pendiente", "Pedido 789012 entregado", "Pedido 3 cancelado"], ["1234"]),
    _ff("ff-telefono", [5512345678, 3398765432, 8110002000], ["(55) 1234-5678"]),
    _ff("ff-codigo-medio", ["ABC-123-XY", "DE-4567-Z", "FGHI-89-UVW"], ["123"]),
    _ff("ff-parentesis", ["Monterrey (NL)", "Guadalajara (JAL)", "Mérida (YUC)", "Toluca (MEX)"], ["NL"]),
    _ff("ff-antes-coma", ["Guadalajara, Jalisco", "Puebla, Puebla", "San Luis Potosí, SLP", "León, Guanajuato"], ["Guadalajara"]),
    _ff("ff-primeras3", ["Enero", "Febrero", "Marzo", "Abril", "Mayo"], ["Ene"]),
    _ff("ff-ultimas4", ["CURP1234", "RFC98765", "ID4321", "XYZ0001"], ["1234"]),
    _ff("ff-intercambiar", ["Pérez, Ana", "Gómez, Luis", "Ruiz, María"], ["Ana Pérez"]),
    _ff("ff-dos-columnas", [["MX-001", "Ana"], ["US-002", "Luis"], ["CA-003", "Eva"]], ["Ana (MX)"]),
    _ff("ff-encabezado", ["Nombre completo", "Ana Pérez", "Luis Gómez", "Eva Ruiz"], ["Nombre", "Ana"]),
    # (más difíciles: dos ejemplos que cambian lo que daría uno solo, decimales, filas que no siguen el patrón, un
    # encabezado que no es ejemplo, usuarios con la inicial, los últimos cuatro de un número, insertar un guion)
    _ff("ff-dos-ejemplos", NOMBRES_FF, ["Pérez", "Gómez"]),
    _ff("ff-decimal", ["Precio: 12.50 MXN", "Precio: 7.05 MXN", "Precio: 1250.75 MXN", "Precio: 3 MXN"], ["12.50"]),
    _ff("ff-falta", ["Pedido 1234 enviado", "Pedido pendiente", "Pedido 77 listo", "Otro 5"], ["1234"]),
    _ff("ff-encabezado-distinto", ["Nombre completo", "Ana Pérez", "Luis Gómez", "Eva Ruiz"], ["Primer nombre", "Ana"]),
    _ff("ff-usuario-inicial", ["Ana Lopez", "Luis Gomez", "Maria Jose Ruiz", "Eva Mora"], ["alopez"]),
    _ff("ff-ultimos4", ["Cuenta 0012345678", "Cuenta 9988776655", "Tarjeta 4111222233334444"], ["5678"]),
    _ff("ff-guion", ["ABC123", "XY45", "QRST6789", "A1"], ["ABC-123"]),
    _ff("ff-puntos", ["21.11.2026", "05.01.2025", "30.06.2024"], ["2026"]),
    _ff("ff-palabras", ["uno-dos-tres", "rojo-verde-azul", "a-b-c"], ["uno dos tres"]),
]


def verdad_excel(casos):
    """Lo que hace Excel con esos casos (se agrega a lo que ya estaba en VERDAD: con --casos se pueden pedir solo unos)."""
    with tempfile.TemporaryDirectory() as d:
        ent = Path(d) / "datos-casos.json"
        ent.write_text(json.dumps(casos, ensure_ascii=False), encoding="utf-8")
        remoto.subir([ent], "nomoffice")
        r = remoto.correr_office("datos", "datos_excel.py", "datos-casos.json", "datos-verdad.json", timeout=900)
        print(r.stdout[-800:])
        nueva = Path(d) / "verdad.json"
        if not remoto.bajar("nomoffice/datos-verdad.json", nueva): raise SystemExit("Excel no dejó el resultado")
        todo = json.load(open(VERDAD, encoding="utf-8")) if VERDAD.exists() else {}
        todo.update(json.load(open(nueva, encoding="utf-8")))
        VERDAD.parent.mkdir(parents=True, exist_ok=True)
        json.dump(todo, open(VERDAD, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    return todo


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
  else if (c.op === "relleno_rapido") res = H.herramientas.rellenoRapido ? H.herramientas.rellenoRapido(r) : "no hay relleno rápido";
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


def elegidos():
    """Los casos a correr: todos, o con --casos PREFIJO[,PREFIJO…] los que empiezan así."""
    if "--casos" not in sys.argv: return CASOS
    pref = sys.argv[sys.argv.index("--casos") + 1].split(",")
    return [c for c in CASOS if any(c["id"].startswith(x) for x in pref)]


def teclado_relleno(nav, N, a, esp):
    """El relleno rápido con el teclado, como en Excel: un ejemplo, elegir la celda de abajo, Ctrl+E; y Ctrl+Z lo quita."""
    d = N.documento(a["token"], esp, "HOJA", "ff-teclado")
    p = nav.p
    p.goto(N.B + "/#/doc/" + d); p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.api", timeout=30000)
    p.wait_for_timeout(500)
    p.evaluate("""() => { const ed = NOM.editorActual, t = ed.modoApi.api.tabla();
      ed.doc.cambiar(t.id, { 'c:A1': 'Ana Pérez', 'c:A2': 'Luis Gómez', 'c:A3': 'Eva Ruiz', 'c:B1': 'Ana' }); return 1; }""")
    p.wait_for_timeout(400)
    p.locator('td[data-c="1"][data-f="1"]').click(); p.wait_for_timeout(200)
    p.keyboard.press("Control+e"); p.wait_for_timeout(600)
    val = lambda: p.evaluate("() => { const ev = NOM.editorActual.modoApi.api.evaluar(); return [1, 2].map((f) => ev.valor(1, f)); }")
    llenas = val()
    p.keyboard.press("Control+z"); p.wait_for_timeout(600)
    despues = val()
    ok = llenas == ["Luis", "Eva"] and all(x in (None, "") for x in despues)
    print(("  bien " if ok else "  MAL  ") + "ff-teclado (Ctrl+E llena B2:B3; Ctrl+Z lo quita)" + ("" if ok else f"  llenas {llenas} · después de Ctrl+Z {despues}"), flush=True)
    return ok


def con_nom(verdad):
    sys.path.insert(0, str(AQUI.parent / "interfaz"))
    import nomprueba as N
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Datos")
    nav = N.Navegador(movil=False); nav.entrar(a)
    bien = mal = 0
    try:
        for c in elegidos():
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
                # (un número que no cabe en la columna de Excel —64 px— se ve «5.51E+09»; las de NOM son más anchas)
                cientifico = isinstance(x["v"], (int, float)) and re.match(r"^-?\d+(\.\d+)?E[+-]\d+$", x["t"])
                if fecha or cientifico or (x["t"].startswith("#") and set(x["t"]) == {"#"}):
                    a, b = x["v"], r["valores"][i][j]
                    return a == b or (isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) < 1e-9)
                return x["t"].strip() == n.strip()
            if c.get("distinto"):
                # a propósito distinto de Excel: lo que debe dar NOM
                igual = [[x.strip() for x in fila] for fila in r["celdas"]] == c["distinto"]
                bien += igual; mal += not igual
                print(("  bien " if igual else "  MAL  ") + c["id"] + f" (a propósito distinto: Excel {xl})" + ("" if igual else f"\n     NOM   {r['celdas']}"), flush=True)
                continue
            igual = all(igual_celda(i, j) for i in range(len(xl)) for j in range(len(xl[i])))
            bien += igual; mal += not igual
            print(("  bien " if igual else "  MAL  ") + c["id"] + ("" if igual else f"\n     Excel {xl}\n     NOM   {r['celdas']}"), flush=True)
        if any(c["op"] == "relleno_rapido" for c in elegidos()):
            ok = teclado_relleno(nav, N, a, esp)
            bien += ok; mal += not ok
    finally:
        nav.cerrar()
    print(f"\nHerramientas de datos contra Excel 2024: {bien} bien, {mal} mal")
    return mal


def main():
    if "--solo-nom" in sys.argv and VERDAD.exists():
        return 1 if con_nom(json.load(open(VERDAD, encoding="utf-8"))) else 0
    v = verdad_excel(elegidos())
    for c in elegidos():
        x = v.get(c["id"], {})
        print(f"== {c['id']}" + (f"  ERROR {x['error']}" if "error" in x else ""))
        for fila in x.get("celdas", []): print("   ", [f"{k['t']}" for k in fila])
        if "cambio" in x: print("    cambio:", x["cambio"], "resultado:", x.get("resultado"), "logrado:", x.get("logrado"))
    if "--solo-excel" in sys.argv: return 0
    return 1 if con_nom(v) else 0


if __name__ == "__main__":
    sys.exit(main())

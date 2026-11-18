#!/usr/bin/env python3
"""Compara las funciones de la hoja de NOM con Excel 2024 de verdad, fórmula por fórmula.

Excel (en la computadora con Windows, por automatización) escribe los datos y cada fórmula como si se tecleara
(Formula2: con matrices dinámicas), recalcula y devuelve el valor, el error exacto y el texto de la celda; después se
evalúan las mismas fórmulas con el motor de NOM (en la página real) y se comparan: el mismo número (hasta 1e-9
relativo), el mismo lógico, el mismo texto o el mismo error.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/funciones_excel.py [--detalle]

Código de salida 1 si alguna función da otro resultado.
"""
import os, subprocess, sys, tempfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import fidelidad as F

DATOS = {  # hoja de datos: columnas A a C
    "A1": 10, "A2": 20, "A3": 30, "A4": 40, "A5": 55, "A6": 7.5, "A7": -3, "A8": 12,
    "B1": 1, "B2": 2, "B3": 3, "B4": 5, "B5": 8, "B6": 13, "B7": 21, "B8": 34,
    "C1": "hola-mundo-fin", "C2": "Texto con  espacios", "C3": "abc", "C4": 46340.6076388889, "C5": "ñandú", "C6": True, "C7": "", "C8": "12",
    "E1": -1000, "E2": 300, "E3": 420, "E4": 680,
    "F1": 1, "F2": 3, "F3": 3, "F4": 5,
}
# (fórmula con nombres en inglés y comas). FECHA de referencia: 46340 = 14/11/2026 (sábado).
CASOS = [
    # fecha y hora
    "HOUR(C4)", "MINUTE(C4)", "SECOND(C4)", "TIME(14,35,10)", 'TIMEVALUE("14:35")', 'TIMEVALUE("2:35 PM")', "HOUR(0.75)",
    "WEEKNUM(46340)", "WEEKNUM(46340,2)", "WEEKNUM(DATE(2027,1,1))", "WEEKNUM(DATE(2026,1,4),2)", "ISOWEEKNUM(46340)", "ISOWEEKNUM(DATE(2027,1,1))", "ISOWEEKNUM(DATE(2024,12,30))",
    "DAYS360(DATE(2026,1,31),DATE(2026,3,31))", "DAYS360(DATE(2026,2,28),DATE(2026,3,31))", "DAYS360(DATE(2026,1,30),DATE(2026,7,31),TRUE)", "DAYS360(DATE(2024,2,29),DATE(2026,11,14))",
    "WORKDAY(46340,5)", "WORKDAY(46340,-3)", "WORKDAY(DATE(2026,11,13),3,DATE(2026,11,16))", "NETWORKDAYS(DATE(2026,11,1),DATE(2026,11,30))",
    "NETWORKDAYS(DATE(2026,11,1),DATE(2026,11,30),DATE(2026,11,16))", "NETWORKDAYS(DATE(2026,11,30),DATE(2026,11,1))",
    # información y lógicas
    "ISNA(NA())", "ISNA(1/0)", "ISERR(1/0)", "ISERR(NA())", "IFNA(VLOOKUP(99,A1:B8,2,FALSE),\"no\")", "IFNA(5,\"no\")", "ISEVEN(4)", "ISEVEN(7)", "ISODD(7)", "ISODD(-4)",
    "ISLOGICAL(C6)", "ISLOGICAL(A1)", "ISNONTEXT(A1)", "ISNONTEXT(C3)", "SWITCH(B2,1,\"uno\",2,\"dos\",\"otro\")", "SWITCH(B4,1,\"uno\",2,\"dos\",\"otro\")", "SWITCH(\"b\",\"a\",1,\"b\",2)",
    "XOR(TRUE,FALSE)", "XOR(TRUE,TRUE)", "XOR(A1>5,B1>5,B8>5)", "N(C3)", "N(A6)", "N(TRUE)", "T(C3)&\"|\"&T(A1)", "TYPE(A1)", "TYPE(C3)", "TRUE()", "FALSE()",
    # texto
    "TEXTBEFORE(C1,\"-\")", "TEXTBEFORE(C1,\"-\",2)", "TEXTBEFORE(C1,\"-\",-1)", "TEXTAFTER(C1,\"-\")", "TEXTAFTER(C1,\"-\",-1)", "CLEAN(\"a\"&CHAR(7)&\"b\")", "CODE(C3)", "CODE(\"Z\")",
    "UNICODE(C5)", "UNICHAR(241)", "UNICHAR(8364)", "FIXED(1234.567)", "FIXED(1234.567,1)", "FIXED(-1234.567,0,TRUE)", "FIXED(0.5,3)", "ROMAN(1994)", "ROMAN(4)", "ROMAN(2026)",
    # matemáticas
    "CEILING(7.3,2)", "CEILING(-7.3,-2)", "CEILING(2.5,1)", "CEILING(-2.5,2)", "FLOOR(7.3,2)", "FLOOR(-7.3,-2)", "FLOOR(-2.5,2)", "CEILING.MATH(7.3)", "CEILING.MATH(-7.3)", "CEILING.MATH(-7.3,2,1)",
    "CEILING.MATH(6.1,5)", "FLOOR.MATH(7.8)", "FLOOR.MATH(-7.3)", "FLOOR.MATH(-7.3,2,1)", "MROUND(17,5)", "MROUND(1.3,0.2)", "MROUND(-10,-3)", "EVEN(3)", "EVEN(-1.2)", "EVEN(4)", "ODD(2)", "ODD(-2.2)", "ODD(0)",
    "GCD(12,18)", "GCD(A1:A4)", "LCM(4,6)", "LCM(B1:B5)", "FACT(5)", "FACT(0)", "COMBIN(8,3)", "COMBIN(50,6)", "PERMUT(8,3)", "QUOTIENT(17,5)", "QUOTIENT(-17,5)", "SUMSQ(A1:A4)", "SUMSQ(3,4)",
    "ASIN(0.5)", "ACOS(0.5)", "ATAN(1)", "ATAN2(1,1)", "ATAN2(-1,2)", "SINH(1)", "COSH(1)", "TANH(0.5)", "AVERAGEA(A1:A4,C3,C6)", "AVERAGEA(C1:C3)", "MAXA(A7,C6)", "MINA(A1:A4,C3)",
    # estadística
    "PERCENTILE(A1:A8,0.3)", "PERCENTILE(A1:A8,0)", "PERCENTILE(A1:A8,1)", "PERCENTILE(B1:B8,0.9)", "QUARTILE(A1:A8,1)", "QUARTILE(A1:A8,2)", "QUARTILE(A1:A8,3)", "QUARTILE(B1:B8,4)",
    "GEOMEAN(B1:B8)", "HARMEAN(B1:B8)", "AVEDEV(A1:A8)", "SLOPE(A1:A8,B1:B8)", "INTERCEPT(A1:A8,B1:B8)", "FORECAST(10,A1:A8,B1:B8)", "TREND(A1:A8,B1:B8,10)", "CORREL(A1:A8,B1:B8)",
    "PEARSON(A1:A8,B1:B8)", "RSQ(A1:A8,B1:B8)", "COVARIANCE.P(A1:A8,B1:B8)", "COVARIANCE.S(A1:A8,B1:B8)", "COVAR(A1:A8,B1:B8)", "RANK.AVG(20,A1:A8)", "RANK.AVG(3,F1:F4)", "RANK.AVG(3,F1:F4,1)",
    # financieras
    "PMT(0.05/12,60,100000)", "PMT(0.05/12,60,100000,0,1)", "PMT(0,10,1000)", "PMT(0.08,5,-20000,5000)", "FV(0.06/12,120,-500)", "FV(0.06/12,120,-500,-1000,1)", "FV(0,12,-100)",
    "PV(0.08/12,240,-800)", "PV(0.05,10,0,10000)", "NPER(0.01,-250,5000)", "NPER(0,-250,5000)", "NPER(0.01,-100,0,3000)", "RATE(60,-1887.12,100000)", "RATE(12,-90,1000)", "RATE(10,0,-1000,2000)",
    "NPV(0.1,E2:E4)", "NPV(0.08,E1:E4)", "IRR(E1:E4)", "IPMT(0.05/12,1,60,100000)", "IPMT(0.05/12,30,60,100000)", "IPMT(0.05/12,2,60,100000,0,1)", "PPMT(0.05/12,1,60,100000)",
    "PPMT(0.05/12,60,60,100000)", "SLN(30000,7500,10)", "EFFECT(0.12,12)", "NOMINAL(0.126825,12)",
    # referencias
    "ROW(A7)", "COLUMN(C5)", "ROWS(A1:B8)", "COLUMNS(A1:C8)", "ROW(B3:C9)", "ADDRESS(3,2)", "ADDRESS(3,2,4)", "ADDRESS(10,28,2)", "INDIRECT(\"A\"&2)", "SUM(INDIRECT(\"A1:A4\"))",
    "SUM(OFFSET(A1,1,0,3,1))", "OFFSET(A1,2,1)", "SUM(OFFSET(A1:A2,2,1))", "ROW()", "COLUMN()",
]
WIN = os.environ.get("NOM_WINDOWS", "user@100.65.29.5")
AQUI = Path(__file__).resolve().parent


def con_excel(dir_tmp):
    """Excel 2024 en la computadora con Windows calcula cada fórmula (por SSH, en la sesión de escritorio)."""
    import json
    ent = Path(dir_tmp) / "funciones-entrada.json"
    ent.write_text(json.dumps({"datos": DATOS, "casos": CASOS}, ensure_ascii=False), encoding="utf-8")
    import remoto
    remoto.subir([ent], "nomoffice/funciones")
    r = remoto.correr_office("funciones", "funciones_excel.py", "funciones\\funciones-entrada.json", "funciones\\salida.json", timeout=1800)
    if "[fin 0]" not in r.stdout: sys.exit("Excel no terminó: " + (r.stdout + r.stderr)[-800:])
    remoto.bajar("nomoffice/funciones/salida.json", Path(dir_tmp) / "salida.json")
    return json.loads((Path(dir_tmp) / "salida.json").read_text(encoding="utf-8"))


JS = r"""
([datos, casos]) => {
  const F = NOM.formulas, celdas = Object.assign({}, datos);
  // (en el mismo lugar que en Excel: columna J, de 20 en 20 renglones; así ROW() y COLUMN() dan lo mismo)
  casos.forEach((f, i) => { celdas[F.ref(9, i * 20)] = "=" + f; });
  const crudo = (c, f) => { const v = celdas[F.ref(c, f)]; return v == null ? "" : v === true ? "=VERDADERO" : v === false ? "=FALSO" : String(v); };
  const ev = F.evaluador(crudo, { filas: casos.length * 20 + 5, cols: 14 });
  return casos.map((_, i) => { const v = ev.valor(9, i * 20); return v instanceof F.ErrorF ? { e: v.codigo } : v instanceof F.FechaF ? { n: v.serie } : typeof v === "number" ? { n: v } : typeof v === "boolean" ? { b: v } : { t: String(v) }; });
}
"""


def igual(xl, nuestro):
    """Excel (valor, error o texto de la celda) contra el valor de NOM."""
    if "e" in xl: return "e" in nuestro and str(nuestro["e"]).upper() == xl["e"].upper()
    if "e" in nuestro: return False
    if "b" in xl: return "b" in nuestro and nuestro["b"] == xl["b"]
    if "n" in xl:
        if "n" not in nuestro: return False
        x = xl["n"]
        return abs(x - nuestro["n"]) <= 1e-9 * max(1.0, abs(x))
    return "t" in nuestro and nuestro["t"] == xl.get("t", "")


def main():
    detalle = "--detalle" in sys.argv
    from playwright.sync_api import sync_playwright
    with tempfile.TemporaryDirectory() as d:
        de_xl = con_excel(d)
    with sync_playwright() as pw:
        nav = pw.chromium.launch(); pag = nav.new_page()
        pag.goto(F.BASE + "/"); pag.wait_for_function("() => window.NOM && NOM.formulas && NOM.formulas.evaluador", timeout=30000)
        nuestros = pag.evaluate(JS, [DATOS, CASOS])
        total = pag.evaluate("() => NOM.formulas.NOMBRES.length")
        nav.close()
    bien = mal = 0
    for f, xl, n in zip(CASOS, de_xl, nuestros):
        if igual(xl, n):
            bien += 1
            if detalle: print(f"  ok   {f} = {xl.get('texto')}")
        else: mal += 1; print(f"  MAL  {f}: Excel = {xl}, NOM = {n}")
    print(f"\nFunciones contra Excel 2024: {bien} bien, {mal} mal, de {len(CASOS)} fórmulas ({total} funciones en la hoja)")
    sys.exit(1 if mal else 0)


if __name__ == "__main__":
    main()

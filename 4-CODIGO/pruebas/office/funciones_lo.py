#!/usr/bin/env python3
"""Compara las funciones de la hoja de NOM con LibreOffice, fórmula por fórmula.

Arma un .xlsx con datos y una fórmula por renglón (con los nombres en inglés, como los guarda
Excel), lo abre LibreOffice sin pantalla para que calcule y lo pasa a CSV; después evalúa las
mismas fórmulas con el motor de NOM (en la página real) y compara los resultados.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/funciones_lo.py [--detalle]

Código de salida 1 si alguna función da otro resultado. Una función que LibreOffice no conoce se
anota como «sin comparar» (no cuenta como falla).
"""
import csv, os, subprocess, signal, sys, tempfile
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
# (TYPE(TRUE) no se compara: LibreOffice dice 1 —para él un lógico es un número— y Excel dice 4, que es lo que da NOM.)
# Lo que LibreOffice escribe con la configuración regional en inglés y lo que escribe NOM
LOGICOS = {"TRUE": True, "FALSE": False, "VERDADERO": True, "FALSO": False}
PREFIJO = {"IFNA": "_xlfn.", "SWITCH": "_xlfn.", "XOR": "_xlfn.", "TEXTBEFORE": "_xlfn.", "TEXTAFTER": "_xlfn.", "UNICHAR": "_xlfn.", "UNICODE": "_xlfn.", "ISOWEEKNUM": "_xlfn.",
           "CEILING.MATH": "_xlfn.", "FLOOR.MATH": "_xlfn.", "COVARIANCE.P": "_xlfn.", "COVARIANCE.S": "_xlfn.", "RANK.AVG": "_xlfn."}


def para_excel(f):
    import re
    return re.sub(r"\b([A-Z][A-Z0-9.]*)\(", lambda m: PREFIJO.get(m.group(1), "") + m.group(1) + "(", f)


def con_libreoffice(dir_tmp):
    import openpyxl
    wb = openpyxl.Workbook(); h = wb.active
    for k, v in DATOS.items(): h[k] = v
    for i, f in enumerate(CASOS):
        h[f"G{i + 1}"] = "=" + para_excel(f)
        # con formato de número fijo: si no, LibreOffice escribe "-$1,887.12" o "0.42%" y se pierden decimales
        h[f"G{i + 1}"].number_format = "0.000000000000"
    x = Path(dir_tmp) / "funciones.xlsx"; wb.save(x)
    # 44 = coma, 34 = comillas, 76 = UTF-8; el 1 del final: los números tal como están guardados (sin formato)
    pr = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{F.CACHE / 'lo-perfil'}", "--headless", "--convert-to", 'csv:Text - txt - csv (StarCalc):44,34,76,1,,1033,false,true',
                           "--outdir", str(dir_tmp), str(x)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
    try: pr.wait(timeout=180)
    except subprocess.TimeoutExpired:
        os.killpg(pr.pid, signal.SIGKILL); pr.wait(); sys.exit("LibreOffice no terminó")
    filas = list(csv.reader(open(Path(dir_tmp) / "funciones.csv", encoding="utf-8")))
    return [(filas[i][6] if i < len(filas) and len(filas[i]) > 6 else "") for i in range(len(CASOS))]


JS = r"""
([datos, casos]) => {
  const F = NOM.formulas, celdas = Object.assign({}, datos);
  casos.forEach((f, i) => { celdas["G" + (i + 1)] = "=" + f; });
  const crudo = (c, f) => { const v = celdas[F.ref(c, f)]; return v == null ? "" : v === true ? "=VERDADERO" : v === false ? "=FALSO" : String(v); };
  const ev = F.evaluador(crudo, { filas: casos.length + 5, cols: 8 });
  return casos.map((_, i) => { const v = ev.valor(6, i); return v instanceof F.ErrorF ? { e: v.codigo } : v instanceof F.FechaF ? { n: v.serie } : typeof v === "number" ? { n: v } : typeof v === "boolean" ? { b: v } : { t: String(v) }; });
}
"""


def igual(lo, nuestro):
    """LibreOffice (texto del CSV) contra el valor de NOM."""
    if "e" in nuestro: return lo.startswith(("#", "Err:")) and not lo.startswith(("#NAME", "Err:508", "Err:509"))
    if "b" in nuestro:
        if lo.upper() in LOGICOS: return LOGICOS[lo.upper()] == nuestro["b"]
        try: return float(lo) in (0.0, 1.0) and bool(float(lo)) == nuestro["b"]   # con formato de número, un lógico sale 1 o 0
        except ValueError: return False
    if "n" in nuestro:
        if lo.upper() in LOGICOS: return float(LOGICOS[lo.upper()]) == nuestro["n"]
        try: x = float(lo)
        except ValueError: return False
        return abs(x - nuestro["n"]) <= 1e-7 * max(1.0, abs(x))
    return lo == nuestro["t"]


def main():
    detalle = "--detalle" in sys.argv
    from playwright.sync_api import sync_playwright
    with tempfile.TemporaryDirectory() as d:
        de_lo = con_libreoffice(d)
    with sync_playwright() as pw:
        nav = pw.chromium.launch(); pag = nav.new_page()
        pag.goto(F.BASE + "/"); pag.wait_for_function("() => window.NOM && NOM.formulas && NOM.formulas.evaluador", timeout=30000)
        nuestros = pag.evaluate(JS, [DATOS, CASOS])
        total = pag.evaluate("() => NOM.formulas.NOMBRES.length")
        nav.close()
    bien = mal = sin = 0
    for f, lo, n in zip(CASOS, de_lo, nuestros):
        if lo.startswith(("#NAME", "Err:508", "Err:509")) or (lo == "" and "t" not in n):
            sin += 1; print(f"  --   {f}: LibreOffice no la calcula ({lo or 'vacío'}); NOM = {n}"); continue
        if igual(lo, n):
            bien += 1
            if detalle: print(f"  ok   {f} = {lo}")
        else: mal += 1; print(f"  MAL  {f}: LibreOffice = {lo!r}, NOM = {n}")
    print(f"\nFunciones contra LibreOffice: {bien} bien, {mal} mal, {sin} sin comparar, de {len(CASOS)} fórmulas ({total} funciones en la hoja)")
    sys.exit(1 if mal else 0)


if __name__ == "__main__":
    main()

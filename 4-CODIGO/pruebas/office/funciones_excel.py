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
    # base de datos (A20:C24) y criterios (E20:F21); flujos con fechas (G20:H24)
    "A20": "Nombre", "B20": "Edad", "C20": "Pago", "A21": "Ana", "B21": 20, "C21": 100, "A22": "Beto", "B22": 35, "C22": 250, "A23": "Ceci", "B23": 28, "C23": 175, "A24": "Dani", "B24": 41, "C24": 90,
    "E20": "Edad", "F20": "Pago", "E21": ">25", "F21": ">100",
    "G20": -10000, "G21": 2750, "G22": 4250, "G23": 3250, "G24": 2750, "H20": 39448, "H21": 39508, "H22": 39751, "H23": 39859, "H24": 39904,
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
    # 18-nov-2026: las que faltaban (matemáticas, fechas, estadística, bases de datos, financieras, ingeniería)
    "ACOSH(2)", "ASINH(1.5)", "ATANH(0.5)", "ACOT(2)", "ACOTH(3)", "COT(0.7)", "COTH(0.7)", "CSC(0.7)", "CSCH(0.7)", "SEC(0.7)", "SECH(0.7)",
    "CEILING.PRECISE(-4.3,2)", "FLOOR.PRECISE(-4.3,2)", "ISO.CEILING(4.3,2)", "COMBINA(4,3)", "PERMUTATIONA(3,2)", "FACTDOUBLE(7)", "FACTDOUBLE(8)",
    "MULTINOMIAL(2,3,4)", "SERIESSUM(2,0,1,B1:B4)", "SQRTPI(2)", "SUMX2MY2(A1:A4,B1:B4)", "SUMX2PY2(A1:A4,B1:B4)", "SUMXMY2(A1:A4,B1:B4)",
    "NETWORKDAYS.INTL(DATE(2026,11,1),DATE(2026,11,30),11)", "NETWORKDAYS.INTL(DATE(2026,11,1),DATE(2026,11,30),\"0000011\")", "NETWORKDAYS.INTL(DATE(2026,11,1),DATE(2026,11,30),7)",
    "WORKDAY.INTL(DATE(2026,11,13),5,11)", "WORKDAY.INTL(DATE(2026,11,13),-4,\"1000001\")",
    "YEARFRAC(DATE(2024,1,15),DATE(2026,11,14))", "YEARFRAC(DATE(2024,1,15),DATE(2026,11,14),1)", "YEARFRAC(DATE(2024,1,15),DATE(2026,11,14),2)",
    "YEARFRAC(DATE(2024,1,15),DATE(2026,11,14),3)", "YEARFRAC(DATE(2024,1,31),DATE(2024,3,31),4)", "YEARFRAC(DATE(2026,2,28),DATE(2026,8,31),0)",
    "DEVSQ(A1:A8)", "KURT(A1:A8)", "SKEW(A1:A8)", "SKEW.P(A1:A8)", "TRIMMEAN(A1:A8,0.25)", "STANDARDIZE(42,40,1.5)", "FISHER(0.75)", "FISHERINV(0.97)",
    "PHI(0.75)", "GAUSS(1.5)", "GAMMA(2.5)", "GAMMA(-1.5)", "GAMMALN(10)", "GAMMALN.PRECISE(4.5)", "ERF(0.745)", "ERF(0.5,1.2)", "ERFC(1)", "ERF.PRECISE(-0.3)",
    "NORM.DIST(42,40,1.5,TRUE)", "NORM.DIST(42,40,1.5,FALSE)", "NORM.S.DIST(1.333,TRUE)", "NORM.S.DIST(1.333,FALSE)", "NORM.INV(0.908789,40,1.5)", "NORM.S.INV(0.025)",
    "EXPON.DIST(0.2,10,TRUE)", "EXPON.DIST(0.2,10,FALSE)", "BINOM.DIST(6,10,0.5,FALSE)", "BINOM.DIST(6,10,0.5,TRUE)", "POISSON.DIST(2,5,FALSE)", "POISSON.DIST(2,5,TRUE)",
    "GAMMA.DIST(10.00001131,9,2,FALSE)", "GAMMA.DIST(10.00001131,9,2,TRUE)", "GAMMA.INV(0.068094,9,2)", "CHISQ.DIST(0.5,1,TRUE)", "CHISQ.DIST(2,3,FALSE)",
    "CHISQ.DIST.RT(18.307,10)", "CHISQ.INV(0.93,1)", "CHISQ.INV.RT(0.050001,10)", "T.DIST(60,1,TRUE)", "T.DIST(8,3,FALSE)", "T.DIST.2T(1.959999998,60)",
    "T.DIST.RT(1.959999998,60)", "T.INV(0.75,2)", "T.INV.2T(0.546449,60)", "F.DIST(15.2069,6,4,TRUE)", "F.DIST(15.2069,6,4,FALSE)", "F.DIST.RT(15.2069,6,4)",
    "F.INV(0.01,6,4)", "F.INV.RT(0.01,6,4)", "BETA.DIST(2,8,10,TRUE,1,3)", "BETA.DIST(2,8,10,FALSE,1,3)", "BETA.INV(0.685470581,8,10,1,3)",
    "CONFIDENCE.NORM(0.05,2.5,50)", "CONFIDENCE.T(0.05,1,50)", "PERCENTILE.EXC(A1:A8,0.25)", "QUARTILE.EXC(A1:A8,3)",
    "NORMDIST(42,40,1.5,TRUE)", "NORMSDIST(1.333)", "NORMINV(0.908789,40,1.5)", "NORMSINV(0.908789)", "EXPONDIST(0.2,10,TRUE)", "BINOMDIST(6,10,0.5,FALSE)",
    "CHIDIST(18.307,10)", "CHIINV(0.050001,10)", "TDIST(1.96,60,2)", "TDIST(1.96,60,1)", "TINV(0.05,60)", "FDIST(15.2069,6,4)", "FINV(0.01,6,4)", "CONFIDENCE(0.05,2.5,50)",
    "GAMMADIST(10.00001131,9,2,TRUE)", "GAMMAINV(0.068094,9,2)", "BETADIST(2,8,10,1,3)", "BETAINV(0.685470581,8,10,1,3)",
    "DSUM(A20:C24,\"Pago\",E20:E21)", "DAVERAGE(A20:C24,3,E20:E21)", "DCOUNT(A20:C24,\"Pago\",E20:F21)", "DCOUNTA(A20:C24,\"Nombre\",E20:E21)", "DMAX(A20:C24,\"Pago\",E20:E21)",
    "DMIN(A20:C24,\"Pago\",E20:E21)", "DPRODUCT(A20:C24,\"Edad\",E20:E21)", "DGET(A20:C24,\"Nombre\",E20:F21)", "DSTDEV(A20:C24,\"Pago\",E20:E21)", "DSTDEVP(A20:C24,\"Pago\",E20:E21)",
    "DVAR(A20:C24,\"Pago\",E20:E21)", "DVARP(A20:C24,\"Pago\",E20:E21)",
    "SYD(30000,7500,10,1)", "SYD(30000,7500,10,10)", "DDB(2400,300,10,1)", "DDB(2400,300,10,8)", "DDB(2400,300,120,2,1.5)", "DB(1000000,100000,6,1,7)", "DB(1000000,100000,6,7,7)",
    "DB(1000000,100000,6,3)", "VDB(2400,300,10,0,1)", "VDB(2400,300,10,6,10)", "VDB(2400,300,120,6,18,1.5)", "VDB(2400,300,10,0,0.875,1.5)",
    "XNPV(0.09,G20:G24,H20:H24)", "XIRR(G20:G24,H20:H24)", "MIRR(E1:E4,0.1,0.12)", "PDURATION(0.025,2000,2200)", "RRI(96,10000,11000)", "FVSCHEDULE(1,{0.09,0.11,0.1})",
    "ISPMT(0.1/12,1,36,8000000)", "DOLLARDE(1.02,16)", "DOLLARFR(1.125,16)",
    "BIN2HEX(11111011,4)", "BIN2HEX(1110)", "BIN2OCT(1001,3)", "HEX2BIN(\"F\",8)", "HEX2OCT(\"3B4E\")", "OCT2BIN(3,3)", "OCT2HEX(100,4)", "HEX2BIN(\"FFFFFFFFFF\")",
    "BITAND(13,25)", "BITOR(23,10)", "BITXOR(5,3)", "BITLSHIFT(4,2)", "BITRSHIFT(13,2)", "DELTA(5,4)", "DELTA(5,5)", "GESTEP(5,4)", "GESTEP(-4,-5)",
    # lote 2: búsqueda y texto
    "LOOKUP(25,A1:A4,B1:B4)", "LOOKUP(55,A1:A5,B1:B5)", "LOOKUP(5,A1:A4,B1:B4)", "LOOKUP(\"c\",{\"a\",\"b\",\"c\",\"d\";1,2,3,4})", "LOOKUP(3,{1,2,3;\"x\",\"y\",\"z\"})",
    "NUMBERVALUE(\"2.500,27\",\",\",\".\")", "NUMBERVALUE(\"3.5%\")", "NUMBERVALUE(\" 1 234 \")", "VALUETOTEXT(A6)", "VALUETOTEXT(C3,1)", "VALUETOTEXT(1/3)", "VALUETOTEXT(TRUE)",
    "ARRAYTOTEXT(A1:B2)", "ARRAYTOTEXT(A1:C3,1)", "ENCODEURL(\"http://contoso.sharepoint.com/Finance/Profit and Loss Statement.xlsx\")", "ENCODEURL(\"ñandú & café?\")",
    "ASC(\"abc\")", "DBCS(\"abc\")", "FINDB(\"m\",C1)", "SEARCHB(\"MUNDO\",C1)", "REPLACEB(C3,2,1,\"X\")",
    # complejos y Bessel
    "COMPLEX(3,4)", "COMPLEX(3,-1,\"j\")", "COMPLEX(0,1)", "COMPLEX(-2.5,0)", "IMAGINARY(\"3+4i\")", "IMREAL(\"6-9i\")", "IMABS(\"5+12i\")", "IMARGUMENT(\"3+4i\")",
    "IMCONJUGATE(\"3+4i\")", "IMCOS(\"1+i\")", "IMSIN(\"4+3i\")", "IMCOSH(\"4+3i\")", "IMSINH(\"4+3i\")", "IMTAN(\"4+3i\")", "IMCOT(\"4+3i\")", "IMSEC(\"4+3i\")",
    "IMCSC(\"4+3i\")", "IMSECH(\"4+3i\")", "IMCSCH(\"4+3i\")", "IMEXP(\"1+i\")", "IMLN(\"3+4i\")", "IMLOG10(\"3+4i\")", "IMLOG2(\"3+4i\")", "IMSQRT(\"1+i\")",
    "IMPOWER(\"2+3i\",3)", "IMDIV(\"-238+240i\",\"10+24i\")", "IMSUB(\"13+4j\",\"5+3j\")", "IMSUM(\"3+4i\",\"5-3i\")", "IMPRODUCT(\"3+4i\",\"5-3i\")", "IMPRODUCT(\"1+2i\",30)",
    "BESSELJ(1.9,2)", "BESSELJ(0.5,0)", "BESSELJ(10,1)", "BESSELJ(3,5)", "BESSELY(2.5,1)", "BESSELY(0.5,0)", "BESSELY(10,3)", "BESSELI(1.5,1)", "BESSELI(4,0)", "BESSELI(2,3)",
    "BESSELK(1.5,1)", "BESSELK(0.5,0)", "BESSELK(3,4)",
    # estadística
    "BINOM.DIST.RANGE(60,0.75,48)", "BINOM.DIST.RANGE(60,0.75,45,50)", "BINOM.INV(6,0.5,0.75)", "CRITBINOM(6,0.5,0.75)", "HYPGEOM.DIST(1,4,8,20,TRUE)", "HYPGEOM.DIST(1,4,8,20,FALSE)",
    "HYPGEOMDIST(1,4,8,20)", "NEGBINOM.DIST(10,5,0.25,TRUE)", "NEGBINOM.DIST(10,5,0.25,FALSE)", "NEGBINOMDIST(10,5,0.25)", "LOGNORM.DIST(4,3.5,1.2,TRUE)", "LOGNORM.DIST(4,3.5,1.2,FALSE)",
    "LOGNORM.INV(0.039084,3.5,1.2)", "LOGNORMDIST(4,3.5,1.2)", "LOGINV(0.039084,3.5,1.2)", "WEIBULL.DIST(105,20,100,TRUE)", "WEIBULL.DIST(105,20,100,FALSE)", "WEIBULL(105,20,100,TRUE)",
    "T.TEST(A1:A8,B1:B8,2,1)", "T.TEST(A1:A8,B1:B8,2,2)", "T.TEST(A1:A8,B1:B8,1,3)", "TTEST(A1:A8,B1:B8,2,3)", "F.TEST(A1:A8,B1:B8)", "FTEST(A1:A8,B1:B8)",
    "CHISQ.TEST(A1:B4,E1:F4)", "CHITEST(A1:A4,B1:B4)", "Z.TEST(A1:A8,10)", "Z.TEST(A1:A8,10,5)", "ZTEST(A1:A8,20)", "INDEX(MODE.MULT(1,2,3,4,3,2,1,2,3),1)",
    "INDEX(MODE.MULT(1,2,3,4,3,2,1,2,3),2)", "PERCENTRANK.INC(A1:A8,20)", "PERCENTRANK.INC(A1:A8,15,2)", "PERCENTRANK.EXC(A1:A8,20)", "PERCENTRANK.EXC(A1:A8,33)", "PERCENTRANK(A1:A8,12)",
    "PROB({0,1,2,3},{0.2,0.3,0.1,0.4},2)", "PROB({0,1,2,3},{0.2,0.3,0.1,0.4},1,3)", "STEYX(A1:A8,B1:B8)",
    "INDEX(LINEST(A1:A8,B1:B8),1)", "INDEX(LINEST(A1:A8,B1:B8),2)", "INDEX(LINEST(A1:A8,B1:B8,TRUE,TRUE),2,1)", "INDEX(LINEST(A1:A8,B1:B8,TRUE,TRUE),3,1)",
    "INDEX(LINEST(A1:A8,B1:B8,TRUE,TRUE),3,2)", "INDEX(LINEST(A1:A8,B1:B8,TRUE,TRUE),4,1)", "INDEX(LINEST(A1:A8,B1:B8,TRUE,TRUE),4,2)", "INDEX(LINEST(A1:A8,B1:B8,TRUE,TRUE),5,1)",
    "INDEX(LINEST(A1:A8,B1:B8,FALSE),1)", "INDEX(LOGEST(B1:B8),1)", "INDEX(LOGEST(B1:B8),2)", "GROWTH(B1:B8,,9)", "GROWTH(B1:B4,A1:A4,50)",
    # financieras
    "CUMIPMT(0.09/12,30*12,125000,13,24,0)", "CUMIPMT(0.09/12,30*12,125000,1,1,0)", "CUMPRINC(0.09/12,30*12,125000,13,24,0)", "CUMPRINC(0.09/12,30*12,125000,1,1,0)",
    "COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,1)", "COUPDAYS(DATE(2011,1,25),DATE(2011,11,15),2,1)", "COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,15),2,1)",
    "COUPNCD(DATE(2011,1,25),DATE(2011,11,15),2,1)", "COUPPCD(DATE(2011,1,25),DATE(2011,11,15),2,1)", "COUPNUM(DATE(2007,1,25),DATE(2008,11,15),2,1)",
    "COUPDAYBS(DATE(2011,1,25),DATE(2011,11,15),2,0)", "COUPDAYS(DATE(2011,1,25),DATE(2011,11,15),2,0)", "COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,15),2,0)",
    "COUPDAYBS(DATE(2011,1,25),DATE(2011,11,30),4,4)", "COUPDAYSNC(DATE(2011,1,25),DATE(2011,11,30),4,4)", "COUPPCD(DATE(2011,3,15),DATE(2015,2,28),2,0)",
    "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)", "PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,4,1)", "PRICE(DATE(2008,2,15),DATE(2008,6,15),0.05,0.06,100,2,0)",
    "YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)", "YIELD(DATE(2008,2,15),DATE(2008,6,15),0.05,99.5,100,2,0)",
    "DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)", "MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
    "DISC(DATE(2018,1,25),DATE(2018,6,15),97.975,100,1)", "INTRATE(DATE(2008,2,15),DATE(2008,5,15),1000000,1014420,2)", "PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)",
    "RECEIVED(DATE(2008,2,15),DATE(2008,5,15),1000000,0.0575,2)", "YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)", "PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)",
    "YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)", "ACCRINT(DATE(2008,3,1),DATE(2008,8,31),DATE(2008,5,1),0.1,1000,2,0)",
    "ACCRINTM(DATE(2008,4,1),DATE(2008,6,15),0.1,1000,3)", "TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)", "TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)",
    "TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)", "TBILLEQ(DATE(2008,3,31),DATE(2009,1,30),0.0914)", "AMORLINC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)",
    "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,1,0.15,1)", "AMORDEGRC(2400,DATE(2008,8,19),DATE(2008,12,31),300,5,0.15,1)",
    "ISFORMULA(A1)", "ISFORMULA(J1)", "ISFORMULA(C1)",
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


# Las que Excel resuelve iterando hasta su precisión documentada («0.000001 por ciento»): la tasa vale si difiere menos
# de 1e-8 (Excel no publica su método y su resultado tampoco es la raíz exacta).
ITERATIVAS = ("IRR(", "XIRR(", "RATE(")


def igual(xl, nuestro, formula=""):
    """Excel (valor, error o texto de la celda) contra el valor de NOM."""
    ES = {"#VALOR!": "#VALUE!", "#¡VALOR!": "#VALUE!", "#NOMBRE?": "#NAME?", "#¿NOMBRE?": "#NAME?", "#NULO!": "#NULL!", "#¡NULO!": "#NULL!", "#¡DIV/0!": "#DIV/0!", "#¡REF!": "#REF!", "#¡NUM!": "#NUM!"}
    if "e" in xl: return "e" in nuestro and ES.get(str(nuestro["e"]).upper(), str(nuestro["e"]).upper()) == xl["e"].upper()
    if "e" in nuestro: return False
    if "b" in xl: return "b" in nuestro and nuestro["b"] == xl["b"]
    if "n" in xl:
        if "n" not in nuestro: return False
        x = xl["n"]
        if formula.startswith(ITERATIVAS): return abs(x - nuestro["n"]) <= 1e-8
        return abs(x - nuestro["n"]) <= 1e-9 * max(1.0, abs(x))
    return "t" in nuestro and nuestro["t"] == xl.get("t", "")


# Diferencias que se conocen y están anotadas en docs/PARIDAD-OFFICE.md (se muestran, pero no hacen fallar la prueba:
# así la regresión sigue avisando de lo nuevo). Si un cambio las arregla, la prueba las cuenta como «bien».
CONOCIDAS = {
    "BESSELJ(10,1)": "con x >= 8 Excel se aparta ~5e-7 del valor exacto; no se encontró su método",
}


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
    bien = mal = conocidas = 0
    for f, xl, n in zip(CASOS, de_xl, nuestros):
        if igual(xl, n, f):
            bien += 1
            if detalle: print(f"  ok   {f} = {xl.get('texto')}")
        elif f in CONOCIDAS: conocidas += 1; print(f"  CONOCIDA  {f}: Excel = {xl}, NOM = {n} ({CONOCIDAS[f]})")
        else: mal += 1; print(f"  MAL  {f}: Excel = {xl}, NOM = {n}")
    print(f"\nFunciones contra Excel 2024: {bien} bien, {mal} mal" + (f", {conocidas} diferencia(s) conocida(s)" if conocidas else "") + f", de {len(CASOS)} fórmulas ({total} funciones en la hoja)")
    sys.exit(1 if mal else 0)


if __name__ == "__main__":
    main()

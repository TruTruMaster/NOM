"""Arma con Excel de verdad libros que usan casi todo lo que Excel sabe hacer.

  python hacer_excel.py [x01 x02 ...]      (sin nombres: todos)

Deja `archivos\\xNN-*.xlsx` y `archivos\\xNN-*.funciones.json` (qué aceptó Excel y qué no).
"""
import os, sys, time
from comun import *


class Libro:
    def __init__(self, x, nombre):
        self.x, self.nombre = x, nombre
        self.wb = x.Workbooks.Add()
        self.b = Bitacora(nombre)
        self.primera = True

    def hoja(self, nombre):
        if self.primera: ws = self.wb.Worksheets(1); self.primera = False
        else: ws = self.wb.Worksheets.Add(After=self.wb.Worksheets(self.wb.Worksheets.Count))
        ws.Name = nombre
        return ws

    def guardar(self):
        wb = self.wb
        with self.b.paso("calcular todo"): self.x.CalculateFull()
        wb.Worksheets(1).Activate()
        ruta = os.path.join(ARCHIVOS, self.nombre + ".xlsx")
        wb.SaveAs(ruta, 51)
        guardar_json(os.path.join(ARCHIVOS, self.nombre + ".funciones.json"), self.b.resumen())
        wb.Close(False)
        return ruta


def poner(ws, celda, filas):
    """Escribe una tabla (lista de filas) desde la celda."""
    r0 = ws.Range(celda)
    for i, f in enumerate(filas):
        for j, v in enumerate(f):
            if v is not None: r0.Offset(i + 1, j + 1).Value = v


VENTAS = [["Fecha", "Región", "Vendedor", "Producto", "Unidades", "Precio", "Total"]]
for k in range(60):
    VENTAS.append([f"2026-{1 + k % 12:02d}-{1 + (k * 7) % 28:02d}", ["Norte", "Sur", "Centro", "Occidente"][k % 4], ["Ana", "Luis", "Marta", "Pedro", "Sofía"][k % 5],
                   ["Cuaderno", "Lápiz", "Mochila", "Calculadora", "Regla", "Tableta"][(k * 3) % 6], 5 + (k * 37) % 60, [25.5, 8.0, 450.0, 320.75, 15.25, 3999.0][(k * 3) % 6], None])


def ventas(ws, celda="A1"):
    poner(ws, celda, VENTAS)
    n = len(VENTAS)
    for i in range(2, n + 1):
        ws.Cells(i, 1).Value = VENTAS[i - 1][0]
        ws.Cells(i, 1).NumberFormat = "dd/mm/yyyy"
        ws.Cells(i, 7).Formula = f"=E{i}*F{i}"
    ws.Range(f"F2:G{n}").NumberFormat = "$#,##0.00"
    ws.Range("A1:G1").Font.Bold = True
    ws.Columns("A:G").AutoFit()
    return n


# ---------------------------------------------------------------- x01: formato de celdas
def x01(x):
    L = Libro(x, "x01-formato"); b = L.b
    ws = L.hoja("Números")
    ws.Range("A1").Value = "Formato"; ws.Range("B1").Value = "Positivo"; ws.Range("C1").Value = "Negativo"; ws.Range("D1").Value = "Cero"; ws.Range("E1").Value = "Texto"
    ws.Range("A1:E1").Font.Bold = True
    formatos = ["General", "0", "0.00", "#,##0", "#,##0.00", "$#,##0.00", '_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "-"??_-;_-@_-', "0%", "0.00%",
                "0.00E+00", "# ?/?", "# ??/??", "d/m/yyyy", "dd-mmm-yy", "dddd", 'd "de" mmmm "de" yyyy', "h:mm AM/PM", "h:mm:ss", "[h]:mm:ss",
                "mm:ss.0", "@", "00000", "[<=9999999]###-####;(###) ###-####", "[Red]-#,##0.00", "#,##0.00;[Red]-#,##0.00",
                "[Blue]#,##0;[Red]-#,##0;[Green]0", '0.0,"K"', '0.0,,"M"', '"Total: "0.00', ";;;", "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
                "[$€-x-euro2] #,##0.00", "[$-es-MX]dddd, d mmmm", "#,##0.00 [$MXN]", "0.000", "#,##0.0;(#,##0.0)", "[>=1000]#,##0,\"k\";0", "yyyy-mm-dd hh:mm",
                "0\"°C\"", "?.??", "#\" piezas\"", "[DBNum1]0"]
    for k, f in enumerate(formatos):
        with b.paso("formato " + f):
            fila = k + 2
            ws.Cells(fila, 1).NumberFormat = "@"; ws.Cells(fila, 1).Value = f
            for col, v in [(2, 46123.4567), (3, -1234.5), (4, 0), (5, "texto")]:
                c = ws.Cells(fila, col); c.NumberFormat = f; c.Value = v
    ws.Columns("A").ColumnWidth = 42; ws.Columns("B:E").ColumnWidth = 22

    ws = L.hoja("Letra")
    filas = [("Aptos Narrow 11 (por omisión)", {}), ("Calibri 11", {"Name": "Calibri"}), ("Arial 10", {"Name": "Arial", "Size": 10}),
             ("Times New Roman 12", {"Name": "Times New Roman", "Size": 12}), ("Courier New 10", {"Name": "Courier New", "Size": 10}),
             ("Aptos 14 negrita", {"Name": "Aptos", "Size": 14, "Bold": True}), ("Cambria 16 cursiva", {"Name": "Cambria", "Size": 16, "Italic": True}),
             ("Georgia 20", {"Name": "Georgia", "Size": 20}), ("Segoe UI 9", {"Name": "Segoe UI", "Size": 9}), ("Consolas 11", {"Name": "Consolas"}),
             ("Subrayado sencillo", {"Underline": 2}), ("Subrayado doble", {"Underline": -4119}), ("Subrayado contable", {"Underline": 4}),
             ("Subrayado doble contable", {"Underline": 5}), ("Tachado", {"Strikethrough": True}), ("Rojo", {"Color": rgb(192, 0, 0)}),
             ("Color de tema Énfasis 1", {"ThemeColor": 5}), ("Énfasis 2 aclarado 40 %", {"ThemeColor": 6, "TintAndShade": 0.4}),
             ("Arial Black 12", {"Name": "Arial Black", "Size": 12}), ("Verdana 8", {"Name": "Verdana", "Size": 8})]
    for k, (t, f) in enumerate(filas):
        with b.paso("letra " + t):
            c = ws.Cells(k + 1, 1); c.Value = t
            for a, v in f.items(): setattr(c.Font, a, v)
    with b.paso("superíndice y subíndice dentro de la celda"):
        c = ws.Cells(22, 1); c.Value = "m2 y H2O"; c.GetCharacters(2, 1).Font.Superscript = True; c.GetCharacters(7, 1).Font.Subscript = True
    with b.paso("texto enriquecido en una celda"):
        c = ws.Cells(23, 1); c.Value = "Rojo, negrita azul, cursiva verde y normal"
        c.GetCharacters(1, 4).Font.Color = rgb(255, 0, 0)
        ch = c.GetCharacters(7, 12).Font; ch.Bold = True; ch.Color = rgb(0, 0, 255)
        ch = c.GetCharacters(21, 14).Font; ch.Italic = True; ch.Color = rgb(0, 128, 0)
    with b.paso("salto de línea dentro de la celda"):
        c = ws.Cells(24, 1); c.Value = "Primera línea\nSegunda línea\nTercera"; c.WrapText = True
    ws.Columns("A").ColumnWidth = 45

    ws = L.hoja("Relleno y bordes")
    colores = [rgb(255, 0, 0), rgb(255, 192, 0), rgb(255, 255, 0), rgb(146, 208, 80), rgb(0, 176, 80), rgb(0, 176, 240), rgb(0, 112, 192), rgb(112, 48, 160)]
    for k, c in enumerate(colores):
        with b.paso(f"relleno sólido {k}"): ws.Cells(1, k + 1).Interior.Color = c; ws.Cells(1, k + 1).Value = k + 1
    for k in range(6):
        for j, t in enumerate([0.8, 0.6, 0.4, -0.25, -0.5]):
            with b.paso(f"relleno de tema {k + 5} matiz {t}"):
                c = ws.Cells(2 + j, k + 1); c.Interior.ThemeColor = k + 5; c.Interior.TintAndShade = t; c.Value = t
    tramas = [(-4125, "gris 50"), (-4126, "gris 75"), (-4124, "gris 25"), (17, "gris 12.5"), (18, "gris 6.25"), (-4128, "horizontal"), (-4166, "vertical"),
              (-4121, "diagonal abajo"), (-4162, "diagonal arriba"), (9, "cuadros"), (10, "semigris"), (11, "horizontal fina"), (15, "rejilla"), (16, "rombos")]
    for k, (pt, nombre) in enumerate(tramas):
        with b.paso("trama " + nombre):
            c = ws.Cells(8 + k // 7, 1 + k % 7); c.Interior.Pattern = pt; c.Interior.PatternColor = rgb(0, 112, 192); c.Interior.Color = rgb(255, 242, 204); c.Value = nombre
    with b.paso("degradado lineal"):
        c = ws.Range("A11:C12"); c.Merge(); c.Value = "Degradado lineal"
        c.Interior.Pattern = 4000; c.Interior.Gradient.Degree = 90
        cs = c.Interior.Gradient.ColorStops; cs.Clear(); cs.Add(0).Color = rgb(255, 255, 255); cs.Add(1).Color = rgb(0, 112, 192)
    with b.paso("degradado desde el centro"):
        c = ws.Range("E11:G12"); c.Merge(); c.Value = "Degradado radial"
        c.Interior.Pattern = 4001; g = c.Interior.Gradient; g.RectangleLeft = 0.5; g.RectangleRight = 0.5; g.RectangleTop = 0.5; g.RectangleBottom = 0.5
        cs = g.ColorStops; cs.Clear(); cs.Add(0).Color = rgb(255, 255, 255); cs.Add(1).Color = rgb(237, 125, 49)
    estilos = [(1, 2, "continuo delgado"), (1, -4138, "continuo medio"), (1, 4, "continuo grueso"), (-4119, 4, "doble"), (-4118, 2, "punteado"),
               (-4115, 2, "guiones"), (4, 2, "guion punto"), (5, 2, "guion punto punto"), (13, -4138, "guion punto inclinado"), (1, 1, "capilar"),
               (-4115, -4138, "guiones medio"), (4, -4138, "guion punto medio")]
    for k, (ls, wt, nombre) in enumerate(estilos):
        with b.paso("borde " + nombre):
            c = ws.Cells(14 + 2 * (k // 4), 1 + 2 * (k % 4)); c.Value = nombre
            for lado in (7, 8, 9, 10):
                bd = c.Borders(lado); bd.LineStyle = ls; bd.Weight = wt; bd.Color = colores[k % 8]
    with b.paso("bordes diagonales"):
        c = ws.Range("A22"); c.Value = "diagonales"
        for lado in (5, 6): c.Borders(lado).LineStyle = 1; c.Borders(lado).Color = rgb(192, 0, 0)
    with b.paso("contorno grueso alrededor de un rango con rejilla interior"):
        r = ws.Range("C22:F26"); r.Value = 7
        for lado in (11, 12): r.Borders(lado).LineStyle = 1; r.Borders(lado).Weight = 1
        r.BorderAround(1, 4, 0, rgb(31, 56, 100))
    for k in range(8): ws.Columns(k + 1).ColumnWidth = 14

    ws = L.hoja("Alineación")
    largo = "Texto largo que no cabe en la celda y se desborda hacia la derecha"
    casos = [("general", 1), ("izquierda", -4131), ("centro", -4108), ("derecha", -4152), ("rellenar", 5), ("justificar", -4130),
             ("centrar en la selección", 7), ("distribuido", -4117)]
    for k, (nombre, h) in enumerate(casos):
        with b.paso("alineación horizontal " + nombre):
            c = ws.Cells(k + 1, 1); c.Value = nombre if h != 5 else "*-"; c.HorizontalAlignment = h
            if h == 7: ws.Range(ws.Cells(k + 1, 1), ws.Cells(k + 1, 3)).HorizontalAlignment = 7
    ws.Columns("A").ColumnWidth = 20
    for k, (nombre, v) in enumerate([("arriba", -4160), ("centro", -4108), ("abajo", -4107), ("justificado", -4130), ("distribuido", -4117)]):
        with b.paso("alineación vertical " + nombre):
            c = ws.Cells(10, k + 1); c.Value = nombre + " " + ("con más texto para ver el reparto vertical" if v in (-4130, -4117) else ""); c.VerticalAlignment = v; c.WrapText = v in (-4130, -4117)
    ws.Rows(10).RowHeight = 60
    with b.paso("ajustar texto"): c = ws.Range("A12"); c.Value = largo; c.WrapText = True
    with b.paso("reducir hasta ajustar"): c = ws.Range("B12"); c.Value = largo; c.ShrinkToFit = True
    with b.paso("sangría 1, 2 y 3"):
        for k in range(3): c = ws.Cells(13 + k, 1); c.Value = f"sangría {k + 1}"; c.IndentLevel = k + 1
    for k, (nombre, o) in enumerate([("45 grados", 45), ("90 grados", 90), ("-90 grados", -90), ("-45 grados", -45), ("vertical apilado", -4166)]):
        with b.paso("orientación " + nombre):
            c = ws.Cells(17, k + 1); c.Value = nombre; c.Orientation = o
    ws.Rows(17).RowHeight = 80
    with b.paso("desbordamiento a la derecha, a la izquierda y a ambos lados"):
        ws.Range("A19").Value = largo
        ws.Range("D20").Value = "Alineado a la derecha se desborda a la izquierda"; ws.Range("D20").HorizontalAlignment = -4152
        ws.Range("C21").Value = "Centrado se desborda a ambos lados"; ws.Range("C21").HorizontalAlignment = -4108
        ws.Range("A22").Value = largo; ws.Range("B22").Value = "vecina con algo (corta el desborde)"
    with b.paso("número que no cabe (####)"):
        ws.Range("F19").Value = 123456789.123; ws.Range("F19").NumberFormat = "#,##0.00"; ws.Columns("F").ColumnWidth = 6
        ws.Range("F20").Value = 46123; ws.Range("F20").NumberFormat = "dddd, d mmmm yyyy"
    with b.paso("celdas combinadas"):
        ws.Range("A24:C24").Merge(); ws.Range("A24").Value = "Combinada horizontal"; ws.Range("A24").HorizontalAlignment = -4108
        ws.Range("D24:D27").Merge(); ws.Range("D24").Value = "Vertical"; ws.Range("D24").VerticalAlignment = -4108
        ws.Range("A25:C27").Merge(); ws.Range("A25").Value = "Bloque combinado 3×3 con texto ajustado " * 2; ws.Range("A25").WrapText = True
    with b.paso("derecha a izquierda en una celda"):
        c = ws.Range("A29"); c.Value = "نص عربي في خلية"; c.ReadingOrder = -5004
    with b.paso("filas y columnas ocultas y de alto/ancho propio"):
        ws.Rows(30).RowHeight = 4; ws.Rows(31).Hidden = True; ws.Columns("H").Hidden = True; ws.Columns("I").ColumnWidth = 2; ws.Columns("J").ColumnWidth = 60
        ws.Range("J30").Value = "Columna J muy ancha"; ws.Range("I30").Value = "estrecha"

    ws = L.hoja("Estilos de celda")
    nombres = ["Good", "Bad", "Neutral", "Title", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Total", "Input", "Output", "Calculation",
               "Check Cell", "Explanatory Text", "Linked Cell", "Note", "Warning Text", "Accent1", "20% - Accent1", "40% - Accent2", "60% - Accent3",
               "Accent4", "Accent5", "Accent6", "Currency", "Percent", "Comma", "Comma [0]", "Currency [0]"]
    for k, n in enumerate(nombres):
        with b.paso("estilo de celda " + n):
            c = ws.Cells(k + 1, 1); c.Value = n if n not in ("Currency", "Percent", "Comma", "Comma [0]", "Currency [0]") else 1234.567
            c.Style = n; ws.Cells(k + 1, 2).Value = n
    with b.paso("estilo de celda propio"):
        st = L.wb.Styles.Add("Destacado NOM"); st.Font.Bold = True; st.Font.Color = rgb(29, 122, 76); st.Interior.Color = rgb(232, 245, 238); st.NumberFormat = "0.0"
        c = ws.Cells(32, 1); c.Value = 9.876; c.Style = "Destacado NOM"
    ws.Columns("A:B").ColumnWidth = 22
    return L.guardar()


# ---------------------------------------------------------------- x02: fórmulas
FORMULAS = [
    ("Suma", "=SUM(Datos!E2:E61)"), ("Promedio", "=AVERAGE(Datos!E2:E61)"), ("Contar", "=COUNT(Datos!E2:E61)"), ("Contar no vacías", "=COUNTA(Datos!B2:B61)"),
    ("Contar si", '=COUNTIF(Datos!B2:B61,"Norte")'), ("Contar si varios", '=COUNTIFS(Datos!B2:B61,"Sur",Datos!E2:E61,">30")'),
    ("Sumar si", '=SUMIF(Datos!B2:B61,"Centro",Datos!G2:G61)'), ("Sumar si varios", '=SUMIFS(Datos!G2:G61,Datos!B2:B61,"Norte",Datos!C2:C61,"Ana")'),
    ("Promedio si", '=AVERAGEIF(Datos!D2:D61,"Lápiz",Datos!E2:E61)'), ("Mínimo", "=MIN(Datos!G2:G61)"), ("Máximo", "=MAX(Datos!G2:G61)"),
    ("Mediana", "=MEDIAN(Datos!E2:E61)"), ("Moda", "=MODE.SNGL(Datos!E2:E61)"), ("Desviación", "=STDEV.S(Datos!E2:E61)"), ("Varianza", "=VAR.S(Datos!E2:E61)"),
    ("Redondear", "=ROUND(PI(),4)"), ("Redondear arriba", "=ROUNDUP(2.341,1)"), ("Redondear abajo", "=ROUNDDOWN(-2.349,2)"), ("Entero", "=INT(-7.5)"),
    ("Residuo", "=MOD(17,5)"), ("Potencia", "=POWER(2,10)"), ("Raíz", "=SQRT(2)"), ("Absoluto", "=ABS(-3.5)"), ("Producto", "=PRODUCT(2,3,4)"),
    ("Suma producto", "=SUMPRODUCT(Datos!E2:E11,Datos!F2:F11)"), ("Exponencial", "=EXP(1)"), ("Logaritmo natural", "=LN(10)"), ("Log base 10", "=LOG10(1000)"),
    ("Seno", "=SIN(PI()/6)"), ("Coseno", "=COS(0)"), ("Tangente", "=TAN(PI()/4)"), ("Arco tangente 2", "=ATAN2(1,1)"), ("Grados", "=DEGREES(PI())"),
    ("Factorial", "=FACT(10)"), ("Combinaciones", "=COMBIN(10,3)"), ("Máximo común divisor", "=GCD(48,36)"), ("Mínimo común múltiplo", "=LCM(4,6)"),
    ("Truncar", "=TRUNC(8.987,1)"), ("Signo", "=SIGN(-4)"), ("Techo", "=CEILING.MATH(4.2)"), ("Piso", "=FLOOR.MATH(4.8)"), ("Múltiplo", "=MROUND(17,5)"),
    ("Si", '=IF(B2>1000,"alto","bajo")'), ("Si varios", '=IFS(B3<10,"poco",B3<100,"medio",TRUE,"mucho")'), ("Y", "=AND(TRUE,1=1)"), ("O", "=OR(FALSE,2>3)"),
    ("No", "=NOT(TRUE)"), ("Xo", "=XOR(TRUE,FALSE)"), ("Si error", '=IFERROR(1/0,"no se puede")'), ("Si no disponible", '=IFNA(NA(),"falta")'),
    ("Cambiar", '=SWITCH(3,1,"uno",2,"dos",3,"tres","otro")'), ("Elegir", '=CHOOSE(2,"rojo","verde","azul")'),
    ("Buscar V", '=VLOOKUP("Mochila",Datos!D2:G61,3,FALSE)'), ("Buscar H", '=HLOOKUP("Precio",Datos!A1:G61,3,FALSE)'),
    ("Buscar X", '=XLOOKUP("Calculadora",Datos!D2:D61,Datos!F2:F61,"no hay")'), ("Índice y coincidir", '=INDEX(Datos!C2:C61,MATCH(MAX(Datos!G2:G61),Datos!G2:G61,0))'),
    ("Coincidir X", '=XMATCH("Regla",Datos!D2:D61)'), ("Desref", "=SUM(OFFSET(Datos!E2,0,0,5,1))"), ("Indirecto", '=INDIRECT("Datos!B"&5)'),
    ("Fila y columna", "=ROW(C7)*100+COLUMN(C7)"), ("Filas", "=ROWS(Datos!A1:G61)"), ("Izquierda", '=LEFT("Colaboración",5)'),
    ("Derecha", '=RIGHT("Colaboración",4)'), ("Extraer", '=MID("Colaboración",3,4)'), ("Largo", '=LEN("Colaboración")'), ("Encontrar", '=FIND("b","Colaboración")'),
    ("Hallar", '=SEARCH("C","colaboración")'), ("Sustituir", '=SUBSTITUTE("a-b-c","-","/")'), ("Reemplazar", '=REPLACE("2026",3,2,"99")'),
    ("Mayúsculas", '=UPPER("hola mundo")'), ("Minúsculas", '=LOWER("HOLA")'), ("Nombre propio", '=PROPER("ana maría lópez")'),
    ("Recortar", '=TRIM("  mucho   espacio  ")'), ("Concatenar", '=CONCAT("NOM","-","COLAB")'), ("Unir texto", '=TEXTJOIN(", ",TRUE,"a","","b","c")'),
    ("Texto con formato", '=TEXT(1234.5,"$#,##0.00")'), ("Texto de fecha", '=TEXT(DATE(2026,11,15),"dddd d \\de mmmm")'), ("Valor", '=VALUE("42.5")'),
    ("Repetir", '=REPT("★",5)'), ("Igual", '=EXACT("a","A")'), ("Carácter", "=CHAR(65)"), ("Código", '=CODE("A")'), ("Unicar", "=UNICHAR(9731)"),
    ("Texto antes", '=TEXTBEFORE("nombre@correo.mx","@")'), ("Texto después", '=TEXTAFTER("nombre@correo.mx","@")'),
    ("Fecha", "=DATE(2026,11,15)"), ("Año", "=YEAR(DATE(2026,11,15))"), ("Mes", "=MONTH(DATE(2026,11,15))"), ("Día de la semana", "=WEEKDAY(DATE(2026,11,15),2)"),
    ("Número de semana", "=WEEKNUM(DATE(2026,11,15))"), ("Fin de mes", "=EOMONTH(DATE(2026,2,10),0)"), ("Fecha más meses", "=EDATE(DATE(2026,1,31),1)"),
    ("Días laborables", "=NETWORKDAYS(DATE(2026,11,1),DATE(2026,11,30))"), ("Día laborable", "=WORKDAY(DATE(2026,11,13),5)"),
    ("Diferencia de fechas", '=DATEDIF(DATE(2000,5,20),DATE(2026,11,15),"y")'), ("Días", "=DAYS(DATE(2026,12,25),DATE(2026,11,15))"),
    ("Hora", "=TIME(14,30,15)"), ("Horas de una hora", "=HOUR(TIME(14,30,15))"), ("Valor de fecha", '=DATEVALUE("2026-11-15")'),
    ("Pago", "=PMT(0.12/12,24,-50000)"), ("Valor futuro", "=FV(0.05/12,120,-1000)"), ("Valor presente", "=PV(0.08,10,-5000)"),
    ("Valor neto", "=NPV(0.1,-10000,3000,4200,6800)"), ("Tasa interna", "=IRR({-10000,3000,4200,6800})"), ("Tasa", "=RATE(36,-1500,40000)"),
    ("Periodos", "=NPER(0.01,-500,10000)"), ("Depreciación", "=SLN(30000,7500,10)"), ("Está vacía", "=ISBLANK(Z99)"), ("Es número", "=ISNUMBER(B2)"),
    ("Es texto", '=ISTEXT("a")'), ("Es error", "=ISERROR(1/0)"), ("Tipo", '=TYPE("a")'), ("Percentil", "=PERCENTILE.INC(Datos!E2:E61,0.9)"),
    ("Cuartil", "=QUARTILE.INC(Datos!E2:E61,1)"), ("Jerarquía", "=RANK.EQ(Datos!E5,Datos!E2:E61)"), ("K-ésimo mayor", "=LARGE(Datos!G2:G61,3)"),
    ("K-ésimo menor", "=SMALL(Datos!G2:G61,2)"), ("Correlación", "=CORREL(Datos!E2:E61,Datos!G2:G61)"), ("Pronóstico", "=FORECAST.LINEAR(70,Datos!G2:G61,Datos!E2:E61)"),
    ("Convertir", '=CONVERT(100,"km","mi")'), ("Decimal a binario", "=DEC2BIN(42)"), ("Hexadecimal", '=HEX2DEC("FF")'), ("Romano", "=ROMAN(2026)"),
    ("Arábigo", '=ARABIC("MMXXVI")'), ("Base", "=BASE(255,16)"), ("Sea (LET)", "=LET(x,5,y,7,x*y+1)"), ("Lambda", "=LAMBDA(n,n*n+1)(6)"),
    ("Error dividir entre cero", "=1/0"), ("Error de nombre", "=NOEXISTE(1)"), ("Error de valor", '="a"+1'), ("Error de número", "=SQRT(-1)"),
    ("Error de referencia", "=INDEX(A1:A2,5)"), ("No disponible", "=NA()"),
    ("Nombre definido", "=IVA*100"), ("Referencia 3D", "=SUM(Enero:Marzo!B2)"), ("Referencia estructurada", "=SUM(TablaVentas[Total])"),
    ("Contar en tabla", '=COUNTIFS(TablaVentas[Región],"Norte")'), ("Matriz antigua (CSE)", None), ("Agregar", "=AGGREGATE(9,6,Datos!E2:E61)"),
    ("Subtotal", "=SUBTOTAL(109,Datos!G2:G61)"),
]
DINAMICAS = [("Ordenar", "=SORT(Datos!C2:C11)"), ("Ordenar por", "=SORTBY(Datos!C2:C6,Datos!E2:E6,-1)"), ("Filtrar", '=FILTER(Datos!C2:C61,Datos!E2:E61>60,"ninguno")'),
             ("Únicos", "=UNIQUE(Datos!B2:B61)"), ("Secuencia", "=SEQUENCE(3,4,10,5)"), ("Transponer", "=TRANSPOSE(Datos!A1:C1)"),
             ("Apilar vertical", "=VSTACK({1,2},{3,4})"), ("Tomar", "=TAKE(Datos!C2:D20,3)"), ("Dividir texto", '=TEXTSPLIT("a,b;c,d",",",";")'),
             ("Por fila (BYROW)", "=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))"), ("Mapear", "=MAP({1,2,3},LAMBDA(v,v*10))"), ("Desbordado bloqueado", "=SEQUENCE(3)")]


def x02(x):
    L = Libro(x, "x02-formulas"); b = L.b
    ws = L.hoja("Fórmulas")
    dat = L.hoja("Datos"); n = ventas(dat)
    with b.paso("tabla de Excel para referencias estructuradas"):
        lo = dat.ListObjects.Add(1, dat.Range(f"A1:G{n}"), None, 1); lo.Name = "TablaVentas"
    for mes in ["Enero", "Febrero", "Marzo"]:
        h = L.wb.Worksheets.Add(After=L.wb.Worksheets(L.wb.Worksheets.Count)); h.Name = mes; h.Range("B2").Value = {"Enero": 100, "Febrero": 250, "Marzo": 75}[mes]
    with b.paso("nombre definido IVA"):
        L.wb.Names.Add("IVA", "=0.16")
    ws.Range("A1").Value = "Función"; ws.Range("B1").Value = "Resultado"; ws.Range("C1").Value = "Fórmula"; ws.Range("A1:C1").Font.Bold = True
    ws.Range("B2").Value = 1500; ws.Range("B3").Value = 55
    fila = 4
    for nombre, f in FORMULAS:
        with b.paso("fórmula " + nombre):
            ws.Cells(fila, 1).Value = nombre
            if f is None:
                ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 2)).FormulaArray = "=SUM(Datos!E2:E11*Datos!F2:F11)"; ws.Cells(fila, 3).Value = "'{=SUM(Datos!E2:E11*Datos!F2:F11)}"
            else:
                ws.Cells(fila, 2).Formula2 = f; ws.Cells(fila, 3).Value = "'" + f
        fila += 1
    fila += 1
    ws.Cells(fila, 1).Value = "Matrices dinámicas (se desbordan hacia abajo y a la derecha)"; ws.Cells(fila, 1).Font.Bold = True; fila += 1
    for nombre, f in DINAMICAS:
        with b.paso("matriz dinámica " + nombre):
            ws.Cells(fila, 1).Value = nombre; ws.Cells(fila, 3).Value = "'" + f; ws.Cells(fila, 4).Formula2 = f
            if nombre == "Desbordado bloqueado": ws.Cells(fila + 1, 4).Value = "estorbo"
        fila += 12 if nombre in ("Únicos", "Filtrar", "Ordenar") else 5
    ws.Columns("A").ColumnWidth = 26; ws.Columns("B").ColumnWidth = 18; ws.Columns("C").ColumnWidth = 60
    with b.paso("tabla de datos (análisis Y si)"):
        t = L.hoja("Y si")
        t.Range("A1").Value = "Tasa"; t.Range("B1").Value = 0.1; t.Range("A2").Value = "Plazo"; t.Range("B2").Value = 12; t.Range("A3").Value = "Monto"; t.Range("B3").Value = 10000
        t.Range("D5").Formula = "=PMT(B1/12,B2,-B3)"
        for k, v in enumerate([0.05, 0.08, 0.1, 0.12, 0.15]): t.Cells(6 + k, 4).Value = v
        for k, v in enumerate([6, 12, 24, 36]): t.Cells(5, 5 + k).Value = v
        t.Range("D5:H10").Table(t.Range("B2"), t.Range("B1"))
        t.Range("E6:H10").NumberFormat = "$#,##0.00"
    return L.guardar()


# ---------------------------------------------------------------- x03: formato condicional y validación
def x03(x):
    L = Libro(x, "x03-condicional"); b = L.b
    ws = L.hoja("Condicional")
    import random
    rnd = random.Random(7)
    ws.Range("A1").Value = "Regla"
    for j in range(10): ws.Cells(1, j + 2).Value = f"C{j + 1}"
    reglas = ["mayor que 50", "entre 20 y 40", "texto contiene", "duplicados", "10 superiores", "10 % inferiores", "arriba del promedio",
              "barras degradadas", "barras sólidas con negativos", "escala 2 colores", "escala 3 colores", "flechas", "semáforos", "banderas",
              "valoración 4", "cuartos 5", "solo íconos", "fórmula filas pares", "varias reglas", "fechas"]
    for i, r in enumerate(reglas):
        ws.Cells(i + 2, 1).Value = r
        for j in range(10):
            v = rnd.randint(-30, 100) if "negativos" in r else rnd.randint(0, 100)
            ws.Cells(i + 2, j + 2).Value = v if r != "texto contiene" else ["error", "bien", "error grave", "ok", "revisar"][j % 5]
    fila = lambda r: ws.Range(ws.Cells(reglas.index(r) + 2, 2), ws.Cells(reglas.index(r) + 2, 11))
    with b.paso("valor mayor que"):
        fc = fila("mayor que 50").FormatConditions.Add(1, 5, "=50"); fc.Interior.Color = rgb(255, 199, 206); fc.Font.Color = rgb(156, 0, 6)
    with b.paso("valor entre"):
        fc = fila("entre 20 y 40").FormatConditions.Add(1, 1, "=20", "=40"); fc.Interior.Color = rgb(255, 235, 156); fc.Font.Color = rgb(156, 87, 0)
    with b.paso("texto que contiene"):
        fc = fila("texto contiene").FormatConditions.Add(9, None, None, None, "error", 0); fc.Interior.Color = rgb(255, 199, 206); fc.Font.Bold = True
    with b.paso("duplicados"):
        fc = fila("duplicados").FormatConditions.AddUniqueValues(); fc.DupeUnique = 1; fc.Interior.Color = rgb(198, 239, 206)
    with b.paso("10 superiores"):
        fc = fila("10 superiores").FormatConditions.AddTop10(); fc.TopBottom = 1; fc.Rank = 3; fc.Interior.Color = rgb(198, 239, 206)
    with b.paso("10 % inferiores"):
        fc = fila("10 % inferiores").FormatConditions.AddTop10(); fc.TopBottom = 0; fc.Rank = 20; fc.Percent = True; fc.Font.Color = rgb(192, 0, 0); fc.Font.Bold = True
    with b.paso("arriba del promedio"):
        fc = fila("arriba del promedio").FormatConditions.AddAboveAverage(); fc.AboveBelow = 0; fc.Interior.Color = rgb(221, 235, 247)
    with b.paso("barras de datos degradadas"):
        db = fila("barras degradadas").FormatConditions.AddDatabar(); db.BarColor.Color = rgb(99, 142, 198); db.BarFillType = 1
    with b.paso("barras sólidas con negativos"):
        db = fila("barras sólidas con negativos").FormatConditions.AddDatabar(); db.BarFillType = 0; db.BarColor.Color = rgb(0, 176, 80)
        db.NegativeBarFormat.ColorType = 0; db.NegativeBarFormat.Color.Color = rgb(255, 0, 0); db.AxisPosition = 0
    with b.paso("escala de 2 colores"):
        cs = fila("escala 2 colores").FormatConditions.AddColorScale(2); cs.ColorScaleCriteria(1).FormatColor.Color = rgb(255, 255, 255); cs.ColorScaleCriteria(2).FormatColor.Color = rgb(99, 190, 123)
    with b.paso("escala de 3 colores"):
        cs = fila("escala 3 colores").FormatConditions.AddColorScale(3)
    for r, conjunto in [("flechas", 1), ("semáforos", 4), ("banderas", 3), ("valoración 4", 13), ("cuartos 5", 17), ("solo íconos", 10)]:
        with b.paso("íconos " + r):
            ic = fila(r).FormatConditions.AddIconSetCondition(); ic.IconSet = L.wb.IconSets(conjunto)
            if r == "solo íconos": ic.ShowIconOnly = True
            if r == "banderas": ic.ReverseOrder = True
    with b.paso("fórmula: filas pares"):
        rg = ws.Range("A24:K40")
        for i in range(24, 41):
            for j in range(1, 12): ws.Cells(i, j).Value = (i * j) % 17
        fc = rg.FormatConditions.Add(2, None, "=MOD(ROW(),2)=0"); fc.Interior.Color = rgb(242, 242, 242)
    with b.paso("varias reglas con prioridad y detener"):
        rg = fila("varias reglas")
        f1 = rg.FormatConditions.Add(1, 5, "=80"); f1.Interior.Color = rgb(0, 176, 80); f1.StopIfTrue = True
        f2 = rg.FormatConditions.Add(1, 5, "=50"); f2.Interior.Color = rgb(255, 192, 0)
        f3 = rg.FormatConditions.AddDatabar()
    with b.paso("fechas: el mes pasado y la semana próxima"):
        rg = fila("fechas")
        for j in range(10): ws.Cells(reglas.index("fechas") + 2, j + 2).Formula = f"=DATE(2026,10,1)+{j * 7}"
        rg.NumberFormat = "dd/mm"
        fc = rg.FormatConditions.Add(1, 7, "=DATE(2026,10,20)"); fc.Font.Color = rgb(112, 48, 160); fc.Font.Italic = True
    ws.Columns("A").ColumnWidth = 26
    v = L.hoja("Validación")
    casos = [("Lista desplegable", 3, 1, "Sí,No,Tal vez", None, "Sí"), ("Lista desde un rango", 3, 1, "=$F$2:$F$6", None, "Rojo"),
             ("Número entero entre 1 y 10", 1, 1, "1", "10", 5), ("Decimal mayor que 0", 2, 5, "0", None, 2.5),
             ("Fecha de 2026", 4, 1, "=DATE(2026,1,1)", "=DATE(2026,12,31)", None), ("Hora de clase", 5, 1, "=TIME(7,0,0)", "=TIME(15,0,0)", None),
             ("Texto de 3 a 10 letras", 6, 1, "3", "10", "Hola"), ("Personalizada (par)", 7, None, "=MOD(B9,2)=0", None, 4)]
    for c, k in enumerate(["Rojo", "Verde", "Azul", "Amarillo", "Negro"]): v.Cells(c + 2, 6).Value = k
    for i, (t, tipo, op, f1, f2, valor) in enumerate(casos):
        with b.paso("validación " + t):
            c = v.Cells(i + 2, 2); v.Cells(i + 2, 1).Value = t
            dv = c.Validation; dv.Delete()
            if f2 is not None: dv.Add(tipo, 1, op, f1, f2)
            elif op is not None: dv.Add(tipo, 1, op, f1)
            else: dv.Add(tipo, 1, 1, f1)
            dv.InputTitle = "Ayuda"; dv.InputMessage = "Escribe: " + t.lower(); dv.ErrorTitle = "Valor no válido"; dv.ErrorMessage = "Revisa la regla: " + t.lower()
            if valor is not None: c.Value = valor
            if tipo == 4: c.Formula = "=DATE(2026,6,15)"; c.NumberFormat = "dd/mm/yyyy"
            if tipo == 5: c.Formula = "=TIME(9,30,0)"; c.NumberFormat = "h:mm"
    with b.paso("validación con aviso de advertencia"):
        c = v.Range("B12"); v.Range("A12").Value = "Advertencia (deja pasar)"; c.Validation.Add(1, 2, 1, "0", "100"); c.Value = 50
    v.Columns("A").ColumnWidth = 30
    return L.guardar()


# ---------------------------------------------------------------- x04: gráficas
TIPOS = [("Columnas agrupadas", 51), ("Columnas apiladas", 52), ("Columnas 100 %", 53), ("Columnas 3D", -4100), ("Barras agrupadas", 57),
         ("Barras apiladas", 58), ("Líneas", 4), ("Líneas con marcadores", 65), ("Líneas apiladas", 63), ("Circular", 5), ("Circular 3D", -4102),
         ("Circular separada", 69), ("Anillo", -4120), ("Áreas", 1), ("Áreas apiladas", 76), ("Dispersión", -4169), ("Dispersión suavizada", 72),
         ("Burbujas", 15), ("Radial", -4151), ("Radial relleno", 82), ("Cotizaciones", 88), ("Superficie", 83), ("Histograma", 118), ("Pareto", 122),
         ("Caja y bigotes", 121), ("Cascada", 119), ("Rectángulos", 117), ("Proyección solar", 120), ("Embudo", 123)]


def x04(x):
    L = Libro(x, "x04-graficas"); b = L.b
    ws = L.hoja("Gráficas")
    datos = [["Mes", "Ventas", "Costos", "Utilidad"], ["Ene", 120, 80, 40], ["Feb", 150, 95, 55], ["Mar", 90, 70, 20], ["Abr", 180, 110, 70], ["May", 160, 100, 60], ["Jun", 210, 130, 80]]
    for i, f in enumerate(datos):
        for j, v in enumerate(f): ws.Cells(i + 1, j + 1).Value = v
    jer = [["Región", "Estado", "Ventas"], ["Norte", "Sonora", 40], ["Norte", "Chihuahua", 35], ["Centro", "CDMX", 90], ["Centro", "Puebla", 30], ["Sur", "Oaxaca", 20], ["Sur", "Chiapas", 15]]
    for i, f in enumerate(jer):
        for j, v in enumerate(f): ws.Cells(i + 1, j + 7).Value = v
    for i, v in enumerate([3, 7, 7, 8, 12, 15, 15, 16, 18, 22, 25, 25, 30, 31, 40, 52, 60]): ws.Cells(i + 1, 11).Value = v
    bolsa = [["Día", "Alto", "Bajo", "Cierre"], ["L", 105, 98, 102], ["M", 108, 100, 107], ["X", 110, 104, 105], ["J", 107, 99, 100], ["V", 112, 101, 111]]
    for i, f in enumerate(bolsa):
        for j, v in enumerate(f): ws.Cells(i + 1, j + 13).Value = v
    cascada = [["Concepto", "Monto"], ["Inicio", 100], ["Ingresos", 60], ["Gastos", -40], ["Impuestos", -15], ["Final", 105]]
    for i, f in enumerate(cascada):
        for j, v in enumerate(f): ws.Cells(i + 1, j + 18).Value = v
    for k, (nombre, tipo) in enumerate(TIPOS):
        with b.paso("gráfica " + nombre):
            fuente = ws.Range("A1:D7")
            if tipo in (5, -4102, 69, -4120): fuente = ws.Range("A1:B7")
            if tipo in (117, 120): fuente = ws.Range("G1:I7")
            if tipo in (118, 122, 121): fuente = ws.Range("K1:K17")
            if tipo in (88,): fuente = ws.Range("M1:P6")
            if tipo in (119,): fuente = ws.Range("R1:S6")
            if tipo in (123,): fuente = ws.Range("A1:B7")
            if tipo == 15: fuente = ws.Range("B1:D7")
            sh = ws.Shapes.AddChart2(-1, tipo, 10 + (k % 3) * 370, 160 + (k // 3) * 260, 360, 250)
            ch = sh.Chart; ch.SetSourceData(fuente)
            if tipo != 88: ch.ChartType = tipo
            ch.HasTitle = True; ch.ChartTitle.Text = nombre
    with b.paso("gráfica combinada con eje secundario"):
        sh = ws.Shapes.AddChart2(-1, 51, 10, 160 + 10 * 260, 500, 300); ch = sh.Chart; ch.SetSourceData(ws.Range("A1:D7"))
        s3 = ch.SeriesCollection(3); s3.ChartType = 65; s3.AxisGroup = 2; ch.HasTitle = True; ch.ChartTitle.Text = "Combinada con eje secundario"
    with b.paso("gráfica con formato detallado"):
        sh = ws.Shapes.AddChart2(-1, 51, 520, 160 + 10 * 260, 560, 320); ch = sh.Chart; ch.SetSourceData(ws.Range("A1:C7"))
        ch.HasTitle = True; ch.ChartTitle.Text = "Formato detallado"; ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 16
        ch.HasLegend = True; ch.Legend.Position = -4160
        ax = ch.Axes(2); ax.HasTitle = True; ax.AxisTitle.Text = "Pesos"; ax.MinimumScale = 0; ax.MaximumScale = 250; ax.MajorUnit = 50; ax.TickLabels.NumberFormat = "$#,##0"
        ax.HasMinorGridlines = True
        cx = ch.Axes(1); cx.HasTitle = True; cx.AxisTitle.Text = "Mes"
        s1 = ch.SeriesCollection(1); s1.HasDataLabels = True; s1.DataLabels().Position = 2; s1.Format.Fill.ForeColor.RGB = rgb(29, 122, 76)
        s2 = ch.SeriesCollection(2); s2.Format.Fill.TwoColorGradient(1, 1); s2.Format.Fill.ForeColor.RGB = rgb(255, 192, 0); s2.Format.Fill.BackColor.RGB = rgb(192, 0, 0)
        tl = s1.Trendlines().Add(-4132); tl.DisplayEquation = True; tl.DisplayRSquared = True
        s2.HasErrorBars = True; s2.ErrorBar(1, 1, 2, 10)
        ch.ChartArea.Format.Line.Visible = True; ch.ChartArea.RoundedCorners = True
        ch.PlotArea.Format.Fill.ForeColor.RGB = rgb(248, 248, 248)
        ch.HasDataTable = True
    for k, estilo in enumerate([201, 205, 209, 214, 220, 230]):
        with b.paso(f"estilo de gráfica {estilo}"):
            sh = ws.Shapes.AddChart2(estilo, 51, 10 + (k % 3) * 370, 160 + (11 + k // 3) * 260 + 80, 360, 240); sh.Chart.SetSourceData(ws.Range("A1:D7"))
    with b.paso("minigráficos de línea, columna y ganancia/pérdida"):
        s = L.hoja("Minigráficos")
        for i in range(6):
            for j in range(12): s.Cells(i + 2, j + 2).Value = ((i + 1) * (j + 3) * 7) % 23 - (8 if i == 5 else 0)
            s.Cells(i + 2, 1).Value = f"Serie {i + 1}"
        s.Range("N2:N3").SparklineGroups.Add(1, "B2:M3")
        g = s.Range("N4:N5").SparklineGroups.Add(2, "B4:M5"); g.Points.Highpoint.Visible = True; g.Points.Highpoint.Color.Color = rgb(192, 0, 0)
        g = s.Range("N6:N7").SparklineGroups.Add(3, "B6:M7")
        g1 = s.Range("N2:N3").SparklineGroups(1); g1.Points.Markers.Visible = True; g1.SeriesColor.Color = rgb(0, 112, 192)
        s.Columns("N").ColumnWidth = 25; s.Rows("2:7").RowHeight = 28
    return L.guardar()


# ---------------------------------------------------------------- x05: tablas, filtros y tablas dinámicas
def x05(x):
    L = Libro(x, "x05-tablas-dinamicas"); b = L.b
    ws = L.hoja("Ventas"); n = ventas(ws)
    lo = None
    with b.paso("tabla con estilo y fila de totales"):
        lo = ws.ListObjects.Add(1, ws.Range(f"A1:G{n}"), None, 1); lo.Name = "Ventas"; lo.TableStyle = "TableStyleMedium2"
        lo.ShowTotals = True; lo.ListColumns(5).TotalsCalculation = 1; lo.ListColumns(7).TotalsCalculation = 1; lo.ListColumns(6).TotalsCalculation = 2
        lo.ShowTableStyleColumnStripes = True
    for k, est in enumerate(["TableStyleLight9", "TableStyleDark11", "TableStyleMedium15", "TableStyleLight1"]):
        with b.paso("tabla " + est):
            r0 = 2 + k * 9
            for i in range(6):
                for j in range(4): ws.Cells(r0 + i, 10 + j).Value = f"T{k}" if i == 0 and j == 0 else (["Col A", "Col B", "Col C", "Col D"][j] if i == 0 else (i * 3 + j))
            t = ws.ListObjects.Add(1, ws.Range(ws.Cells(r0, 10), ws.Cells(r0 + 5, 13)), None, 1); t.TableStyle = est
            if k == 1: t.ShowAutoFilterDropDown = False; t.ShowHeaders = True
    with b.paso("columna calculada en la tabla"):
        c = lo.ListColumns.Add(); c.Name = "Con IVA"; c.DataBodyRange.Formula = "=[@Total]*1.16"; c.DataBodyRange.NumberFormat = "$#,##0.00"
    with b.paso("filtro aplicado y orden"):
        f = L.hoja("Filtrada"); n2 = ventas(f)
        so = f.Sort; so.SortFields.Clear(); so.SortFields.Add(Key=f.Range(f"G2:G{n2}"), SortOn=0, Order=2); so.SetRange(f.Range(f"A1:G{n2}")); so.Header = 1; so.Apply()
        f.Range(f"A1:G{n2}").AutoFilter(Field=2, Criteria1="Norte", Operator=7)
    with b.paso("tabla dinámica"):
        td = L.hoja("Dinámica")
        pc = L.wb.PivotCaches().Create(1, "Ventas!R1C1:R%dC7" % n, 8)
        pt = pc.CreatePivotTable(td.Range("A3"), "TDVentas", True, 8)
        pt.PivotFields("Región").Orientation = 1
        pt.PivotFields("Producto").Orientation = 2
        pt.AddDataField(pt.PivotFields("Total"), "Suma de Total", -4157).NumberFormat = "$#,##0"
        pt.PivotFields("Vendedor").Orientation = 3
        pt.TableStyle2 = "PivotStyleMedium9"
    with b.paso("tabla dinámica con promedio y % del total"):
        pt2 = pc.CreatePivotTable(td.Range("A25"), "TDVendedor", True, 8)
        pt2.PivotFields("Vendedor").Orientation = 1
        df = pt2.AddDataField(pt2.PivotFields("Unidades"), "Promedio de unidades", -4106); df.NumberFormat = "0.0"
        df2 = pt2.AddDataField(pt2.PivotFields("Total"), "% del total", -4157); df2.Calculation = 8; df2.NumberFormat = "0.0%"
    with b.paso("gráfica dinámica"):
        sh = td.Shapes.AddChart2(-1, 51, 500, 20, 420, 260); sh.Chart.SetSourceData(pt.TableRange1)
    with b.paso("segmentación de datos"):
        sc = L.wb.SlicerCaches.Add2(pt, "Región"); sc.Slicers.Add(SlicerDestination=td, Name="SegRegion", Caption="Región", Top=300, Left=500, Width=150, Height=180)
    with b.paso("segmentación de la tabla"):
        sc2 = L.wb.SlicerCaches.Add2(lo, "Producto"); sc2.Slicers.Add(SlicerDestination=ws, Name="SegProducto", Caption="Producto", Top=20, Left=700, Width=150, Height=200)
    return L.guardar()


# ---------------------------------------------------------------- x06: objetos, comentarios y vínculos
def x06(x):
    L = Libro(x, "x06-objetos"); b = L.b
    ws = L.hoja("Objetos")
    for i in range(1, 40):
        for j in range(1, 8): ws.Cells(i, j).Value = i * j
    with b.paso("imagen"):
        p = ws.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 400, 10, 240, 150)
    with b.paso("imagen recortada y girada con borde"):
        p = ws.Shapes.AddPicture(os.path.join(IMG, "foto2.jpg"), False, True, 660, 10, 150, 150)
        p.PictureFormat.CropLeft = 30; p.PictureFormat.CropBottom = 20; p.Rotation = 10; p.Line.Visible = True; p.Line.Weight = 3; p.Line.ForeColor.RGB = rgb(29, 122, 76)
    with b.paso("logotipo con transparencia"):
        ws.Shapes.AddPicture(os.path.join(IMG, "logo.png"), False, True, 830, 10, 90, 90)
    with b.paso("ícono SVG"):
        ws.Shapes.AddPicture(os.path.join(IMG, "icono.svg"), False, True, 930, 10, 60, 60)
    for k, (tipo, nombre) in enumerate([(1, "Rectángulo"), (5, "Redondeado"), (9, "Elipse"), (33, "Flecha"), (92, "Estrella"), (106, "Llamada")]):
        with b.paso("forma " + nombre):
            s = ws.Shapes.AddShape(tipo, 400 + k * 95, 190, 85, 60); s.TextFrame2.TextRange.Text = nombre
            s.Fill.ForeColor.RGB = [rgb(68, 114, 196), rgb(237, 125, 49), rgb(112, 173, 71), rgb(255, 192, 0), rgb(91, 155, 213), rgb(165, 165, 165)][k]
    with b.paso("cuadro de texto"):
        t = ws.Shapes.AddTextbox(1, 400, 270, 300, 80); t.TextFrame2.TextRange.Text = "Cuadro de texto en la hoja con varias líneas de texto que se ajustan solas al ancho."
        t.Fill.ForeColor.RGB = rgb(255, 242, 204); t.Line.ForeColor.RGB = rgb(191, 143, 0)
    with b.paso("WordArt"):
        ws.Shapes.AddTextEffect(10, "Hoja con estilo", "Arial Black", 28, True, False, 720, 270)
    with b.paso("SmartArt"):
        lay = None
        for k in range(1, x.SmartArtLayouts.Count + 1):
            if x.SmartArtLayouts(k).Name == "Basic Process": lay = x.SmartArtLayouts(k); break
        sa = ws.Shapes.AddSmartArt(lay, 400, 370, 420, 120)
        for k, tx in enumerate(["Captura", "Revisión", "Entrega"]):
            if k < sa.SmartArt.AllNodes.Count: sa.SmartArt.AllNodes(k + 1).TextFrame2.TextRange.Text = tx
    with b.paso("notas (comentarios antiguos)"):
        ws.Range("A2").AddComment("Nota normal: aparece al pasar el mouse.")
        c = ws.Range("B3").AddComment("Nota siempre visible."); c.Visible = True
    with b.paso("comentarios con hilo y respuesta"):
        ct = ws.Range("C5").AddCommentThreaded("¿Este total está bien?"); ct.AddReply("Sí, lo revisé ayer.")
    with b.paso("vínculos"):
        ws.Hyperlinks.Add(ws.Range("A42"), "https://www.ipn.mx", "", "Sitio del IPN", "Ir al sitio del IPN")
        ws.Hyperlinks.Add(ws.Range("A43"), "", "Objetos!G30", "", "Ir a la celda G30")
        ws.Hyperlinks.Add(ws.Range("A44"), "mailto:contacto@ejemplo.mx", "", "", "Escribir un correo")
    with b.paso("controles de formulario"):
        cb = ws.Shapes.AddFormControl(1, 400, 500, 120, 20); cb.TextFrame.Characters().Text = "Acepto"; cb.ControlFormat.Value = 1
        dd = ws.Shapes.AddFormControl(2, 530, 500, 120, 20); ws.Range("J1:J3").Value = ws.Range("J1:J3").Value; ws.Range("J1").Value = "Uno"; ws.Range("J2").Value = "Dos"; ws.Range("J3").Value = "Tres"
        dd.ControlFormat.ListFillRange = "J1:J3"; dd.ControlFormat.ListIndex = 2
        bt = ws.Shapes.AddFormControl(0, 660, 500, 100, 24); bt.TextFrame.Characters().Text = "Botón"
        ob = ws.Shapes.AddFormControl(7, 770, 500, 110, 20); ob.TextFrame.Characters().Text = "Opción"
    with b.paso("grupo de formas"):
        s1 = ws.Shapes.AddShape(9, 400, 540, 60, 60); s2 = ws.Shapes.AddShape(1, 440, 560, 80, 30)
        g = ws.Shapes.Range([s1.Name, s2.Name]).Group(); g.Rotation = 15
    with b.paso("conector entre formas"):
        a = ws.Shapes.AddShape(1, 600, 540, 80, 40); c = ws.Shapes.AddShape(9, 760, 600, 80, 40)
        cn = ws.Shapes.AddConnector(2, 0, 0, 10, 10); cn.ConnectorFormat.BeginConnect(a, 4); cn.ConnectorFormat.EndConnect(c, 2); cn.Line.EndArrowheadStyle = 2
    return L.guardar()


# ---------------------------------------------------------------- x07: hojas, vista e impresión
def x07(x):
    L = Libro(x, "x07-hojas"); b = L.b
    ws = L.hoja("Inmovilizada")
    for i in range(1, 101):
        for j in range(1, 27): ws.Cells(i, j).Value = f"{chr(64 + j)}{i}" if i == 1 or j == 1 else i * j
    ws.Rows(1).Font.Bold = True; ws.Columns(1).Font.Bold = True
    with b.paso("paneles inmovilizados (fila 1 y columna A)"):
        ws.Activate(); win = L.wb.Windows(1); win.FreezePanes = False; win.SplitColumn = 1; win.SplitRow = 1; win.FreezePanes = True
    with b.paso("color de la pestaña"): ws.Tab.Color = rgb(29, 122, 76)
    h2 = L.hoja("Sin cuadrícula")
    for i in range(1, 20): h2.Cells(i, 2).Value = texto(i)
    with b.paso("sin cuadrícula ni encabezados, zoom 150 %"):
        h2.Activate(); win = L.wb.Windows(1); win.DisplayGridlines = False; win.DisplayHeadings = False; win.Zoom = 150
        h2.Tab.Color = rgb(255, 192, 0)
    h3 = L.hoja("Derecha a izquierda")
    with b.paso("hoja de derecha a izquierda"):
        h3.DisplayRightToLeft = True
        for i in range(1, 10): h3.Cells(i, 1).Value = f"سطر {i}"; h3.Cells(i, 2).Value = i * 10
    h4 = L.hoja("Esquema")
    for i in range(1, 31): h4.Cells(i, 1).Value = f"Fila {i}"; h4.Cells(i, 2).Value = i
    with b.paso("filas agrupadas y contraídas"):
        h4.Rows("3:8").Group(); h4.Rows("12:20").Group(); h4.Rows("14:16").Group(); h4.Outline.ShowLevels(1)
    with b.paso("columnas agrupadas"):
        for j in range(3, 9): h4.Cells(1, j).Value = f"Col {j}"
        h4.Columns("D:F").Group()
    with b.paso("ancho por omisión de las columnas"):
        h4.StandardWidth = 12
    h5 = L.hoja("Oculta"); h5.Range("A1").Value = "Hoja oculta"
    with b.paso("hoja oculta"): h5.Visible = 0
    h6 = L.hoja("Muy oculta"); h6.Range("A1").Value = "Hoja muy oculta"
    with b.paso("hoja muy oculta"): h6.Visible = 2
    h7 = L.hoja("Impresión")
    for i in range(1, 120):
        for j in range(1, 9): h7.Cells(i, j).Value = f"Encabezado {j}" if i == 1 else round(i * j * 1.37, 2)
    ps = h7.PageSetup
    for nombre, f in [("área de impresión", lambda: setattr(ps, "PrintArea", "$A$1:$H$119")), ("títulos que se repiten", lambda: setattr(ps, "PrintTitleRows", "$1:$1")),
                      ("hoja horizontal", lambda: setattr(ps, "Orientation", 2)), ("ajustar a 1 página de ancho", lambda: (setattr(ps, "Zoom", False), setattr(ps, "FitToPagesWide", 1), setattr(ps, "FitToPagesTall", False))),
                      ("encabezado y pie", lambda: (setattr(ps, "CenterHeader", '&"Arial,Bold"Reporte de prueba'), setattr(ps, "RightHeader", "&D"), setattr(ps, "CenterFooter", "Página &P de &N"), setattr(ps, "LeftFooter", "&F"))),
                      ("imprimir cuadrícula y centrar", lambda: (setattr(ps, "PrintGridlines", True), setattr(ps, "CenterHorizontally", True))),
                      ("márgenes", lambda: (setattr(ps, "TopMargin", 54), setattr(ps, "BottomMargin", 54), setattr(ps, "LeftMargin", 36), setattr(ps, "RightMargin", 36)))]:
        with b.paso("impresión: " + nombre): f()
    with b.paso("saltos de página manuales"):
        h7.HPageBreaks.Add(h7.Range("A40")); h7.HPageBreaks.Add(h7.Range("A80"))
    h8 = L.hoja("Protegida"); h8.Range("A1").Value = "Celda bloqueada"; h8.Range("B1").Value = "Celda libre"; h8.Range("B1").Locked = False
    with b.paso("hoja protegida"): h8.Protect("nom", True, True, True)
    h9 = L.hoja("Dividida")
    for i in range(1, 60): h9.Cells(i, 1).Value = i
    with b.paso("ventana dividida sin inmovilizar"):
        h9.Activate(); win = L.wb.Windows(1); win.SplitRow = 10; win.SplitColumn = 0
    with b.paso("vista de diseño de página"):
        h10 = L.hoja("Diseño de página"); h10.Range("A1").Value = "Esta hoja se guardó en vista de diseño de página"
        h10.Activate(); L.wb.Windows(1).View = 3
    with b.paso("libro protegido (estructura)"):
        L.wb.Worksheets(1).Activate(); L.wb.Protect("nom", True)
    return L.guardar()


# ---------------------------------------------------------------- x08: registro escolar y presupuesto
def x08(x):
    L = Libro(x, "x08-registro"); b = L.b
    ws = L.hoja("Calificaciones")
    enc = ["No.", "Nombre", "Grupo", "Unidad 1", "Unidad 2", "Unidad 3", "Proyecto", "Promedio", "Estado"]
    for j, t in enumerate(enc): ws.Cells(1, j + 1).Value = t
    nombres = ["Ana", "Luis", "Marta", "Pedro", "Sofía", "Diego", "Valeria", "Jorge", "Camila", "Andrés", "Lucía", "Mateo", "Paula", "Iván", "Renata"]
    apellidos = ["García", "López", "Hernández", "Martínez", "Pérez", "Sánchez", "Ramírez", "Torres", "Flores", "Díaz"]
    for i in range(60):
        f = i + 2
        ws.Cells(f, 1).Value = i + 1; ws.Cells(f, 2).Value = f"{nombres[i % 15]} {apellidos[(i * 7) % 10]} {apellidos[(i * 3) % 10]}"
        ws.Cells(f, 3).Value = f"5IM{1 + i % 4}"
        for j in range(4): ws.Cells(f, 4 + j).Value = round(5 + ((i * 13 + j * 17) % 50) / 10, 1)
        ws.Cells(f, 8).Formula = f"=ROUND(D{f}*0.25+E{f}*0.25+F{f}*0.2+G{f}*0.3,1)"
        ws.Cells(f, 9).Formula = f'=IF(H{f}>=6,"Aprobado","Reprobado")'
    with b.paso("formato de encabezado y anchos"):
        h = ws.Range("A1:I1"); h.Font.Bold = True; h.Font.Color = rgb(255, 255, 255); h.Interior.Color = rgb(31, 78, 121); h.HorizontalAlignment = -4108; h.WrapText = True
        ws.Columns("A").ColumnWidth = 5; ws.Columns("B").ColumnWidth = 30; ws.Columns("C").ColumnWidth = 8; ws.Columns("D:H").ColumnWidth = 10; ws.Columns("I").ColumnWidth = 12
        ws.Range("D2:H61").NumberFormat = "0.0"
    with b.paso("reprobados en rojo y barras en el promedio"):
        fc = ws.Range("D2:G61").FormatConditions.Add(1, 6, "=6"); fc.Font.Color = rgb(192, 0, 0); fc.Font.Bold = True
        db = ws.Range("H2:H61").FormatConditions.AddDatabar(); db.BarColor.Color = rgb(91, 155, 213)
        fc2 = ws.Range("I2:I61").FormatConditions.Add(9, None, None, None, "Reprobado", 0); fc2.Interior.Color = rgb(255, 199, 206)
    with b.paso("validación de calificaciones"):
        ws.Range("D2:G61").Validation.Add(2, 1, 1, "0", "10")
    with b.paso("inmovilizar encabezado y filtro"):
        ws.Activate(); win = L.wb.Windows(1); win.FreezePanes = False; win.SplitRow = 1; win.SplitColumn = 2; win.FreezePanes = True
        ws.Range("A1:I61").AutoFilter(Field=1)
    with b.paso("resumen por grupo con fórmulas"):
        r = L.hoja("Resumen")
        for j, t in enumerate(["Grupo", "Alumnos", "Promedio", "Aprobados", "% aprobación"]): r.Cells(1, j + 1).Value = t
        for k in range(4):
            g = f"5IM{k + 1}"; f = k + 2
            r.Cells(f, 1).Value = g
            r.Cells(f, 2).Formula = f'=COUNTIF(Calificaciones!C:C,A{f})'
            r.Cells(f, 3).Formula = f'=AVERAGEIF(Calificaciones!C:C,A{f},Calificaciones!H:H)'
            r.Cells(f, 4).Formula = f'=COUNTIFS(Calificaciones!C:C,A{f},Calificaciones!I:I,"Aprobado")'
            r.Cells(f, 5).Formula = f"=D{f}/B{f}"
        r.Range("C2:C5").NumberFormat = "0.00"; r.Range("E2:E5").NumberFormat = "0.0%"
        r.Range("A1:E1").Font.Bold = True; r.Range("A1:E5").Borders.LineStyle = 1; r.Columns("A:E").ColumnWidth = 14
        sh = r.Shapes.AddChart2(-1, 51, 20, 120, 420, 240); sh.Chart.SetSourceData(r.Range("A1:A5,C1:C5")); sh.Chart.HasTitle = True; sh.Chart.ChartTitle.Text = "Promedio por grupo"
        sh2 = r.Shapes.AddChart2(-1, 118, 460, 120, 360, 240); sh2.Chart.SetSourceData(L.wb.Worksheets("Calificaciones").Range("H1:H61"))
    with b.paso("presupuesto con moneda, totales y minigráficos"):
        p = L.hoja("Presupuesto")
        meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun"]
        conceptos = [("Ingresos", None), ("Colegiaturas", [52000, 52000, 54000, 54000, 55000, 55000]), ("Donativos", [3000, 0, 1500, 0, 4500, 2000]),
                     ("Egresos", None), ("Sueldos", [-38000] * 6), ("Servicios", [-4200, -4100, -4500, -4800, -5200, -5600]), ("Material", [-2500, -1200, -900, -3100, -800, -1500])]
        p.Cells(1, 1).Value = "Concepto"
        for j, m in enumerate(meses): p.Cells(1, j + 2).Value = m
        p.Cells(1, 8).Value = "Total"; p.Cells(1, 9).Value = "Tendencia"
        for i, (c, v) in enumerate(conceptos):
            f = i + 2; p.Cells(f, 1).Value = c
            if v is None: p.Cells(f, 1).Font.Bold = True; p.Range(p.Cells(f, 1), p.Cells(f, 9)).Interior.Color = rgb(221, 235, 247); continue
            for j, x2 in enumerate(v): p.Cells(f, j + 2).Value = x2
            p.Cells(f, 8).Formula = f"=SUM(B{f}:G{f})"; p.Cells(f, 1).IndentLevel = 1
        f = len(conceptos) + 2; p.Cells(f, 1).Value = "Saldo"; p.Cells(f, 1).Font.Bold = True
        for j in range(7): col = chr(66 + j); p.Cells(f, j + 2).Formula = f"=SUM({col}2:{col}{f - 1})"
        p.Range(f"B2:H{f}").NumberFormat = "$#,##0;[Red]-$#,##0"
        p.Range(f"A{f}:H{f}").Borders(8).LineStyle = -4119
        p.Range(f"I3:I{f}").SparklineGroups.Add(1, f"B3:G{f}")
        p.Columns("A").ColumnWidth = 16; p.Columns("B:H").ColumnWidth = 11; p.Columns("I").ColumnWidth = 16
        sh = p.Shapes.AddChart2(-1, 4, 20, 200, 520, 240); sh.Chart.SetSourceData(p.Range(f"A1:G1,A{f}:G{f}"))
    return L.guardar()


# ---------------------------------------------------------------- x09: matrices dinámicas, LET y LAMBDA a fondo
MATRICES = [
    ("Rango que derrama", "=Datos!C2:C6", 6), ("Rango por 2", "=Datos!E2:E6*2", 6), ("Comparación", "=Datos!E2:E6>40", 6),
    ("Concatenar rangos", '=Datos!B2:B5&"-"&Datos!C2:C5', 5), ("Largo de cada uno", "=LEN(Datos!D2:D6)", 6), ("Si con matriz", '=IF(Datos!E2:E6>40,"mucho","poco")', 6),
    ("Redondear matriz", "=ROUND(Datos!F2:F6/3,1)", 6), ("Mayúsculas", "=UPPER(Datos!C2:C4)", 4), ("Buscar V con varias", '=VLOOKUP({"Lápiz";"Regla"},Datos!D2:F61,3,FALSE)', 3),
    ("Buscar X fila entera", '=XLOOKUP("Mochila",Datos!D2:D61,Datos!D2:G61)', 2), ("Buscar X varias", '=XLOOKUP({"Regla","Lápiz"},Datos!D2:D61,Datos!E2:E61)', 2),
    ("Índice columna entera", "=INDEX(Datos!A2:G6,0,3)", 6), ("Índice fila entera", "=INDEX(Datos!A2:G6,2,0)", 2), ("Contar si con varias", '=COUNTIF(Datos!B2:B61,{"Norte","Sur","Centro"})', 2),
    ("Ordenar desc por col 2", "=SORT(Datos!B2:C9,2,-1)", 9), ("Ordenar dos niveles", "=SORT(Datos!B2:E13,{1,4},{1,-1})", 13), ("Ordenar por columnas", "=SORT({3,1,2;30,10,20},1,1,TRUE)", 3),
    ("Ordenar por dos claves", "=SORTBY(Datos!C2:C13,Datos!B2:B13,1,Datos!E2:E13,-1)", 13), ("Filtrar dos condiciones", '=FILTER(Datos!C2:E61,(Datos!B2:B61="Norte")*(Datos!E2:E61>30))', 10),
    ("Filtrar o", '=FILTER(Datos!D2:D61,(Datos!E2:E61<8)+(Datos!E2:E61>62),"")', 6), ("Filtrar columnas", "=FILTER(Datos!A1:G3,{1,0,1,0,0,0,1})", 4),
    ("Filtrar vacío (#CALC!)", "=FILTER(Datos!C2:C61,Datos!E2:E61>1000)", 2), ("Únicos ordenados", "=SORT(UNIQUE(Datos!D2:D61))", 8), ("Únicos una vez", "=UNIQUE({1;2;2;3;4;4;5},FALSE,TRUE)", 5),
    ("Únicos por columnas", "=UNIQUE({1,2,1,3,2},TRUE)", 2), ("Únicos de dos columnas", "=UNIQUE(Datos!B2:C21)", 16), ("Contar únicos", "=COUNTA(UNIQUE(Datos!C2:C61))", 2),
    ("Secuencia descendente", "=SEQUENCE(4,1,10,-2.5)", 5), ("Secuencia de fechas", "=TEXT(SEQUENCE(3,1,DATE(2026,11,15),7),\"dd/mm/yyyy\")", 4), ("Suma de secuencia", "=SUM(SEQUENCE(100))", 2),
    ("Transponer secuencia", "=TRANSPOSE(SEQUENCE(2,3))", 4), ("Apilar horizontal", "=HSTACK({1;2;3},{\"a\";\"b\"})", 4), ("Apilar rangos", "=VSTACK(Datos!C2:D3,Datos!C60:D61)", 5),
    ("Tomar las últimas", "=TAKE(Datos!C2:E61,-3)", 4), ("Tomar columnas", "=TAKE(Datos!A2:G4,,2)", 4), ("Excluir", "=DROP(Datos!C2:E8,4,1)", 4),
    ("Excluir del final", "=DROP({1,2,3;4,5,6;7,8,9},-1,-1)", 3), ("Elegir columnas", "=CHOOSECOLS(Datos!A2:G4,3,-1)", 4), ("Elegir filas", "=CHOOSEROWS(Datos!C2:D20,1,3,-1)", 4),
    ("En columna sin vacíos", "=TOCOL({1,\"\",3;4,5,\"\"},1)", 5), ("En fila por columnas", "=TOROW({1,2;3,4},0,TRUE)", 2), ("Ajustar en filas", "=WRAPROWS(SEQUENCE(7),3,0)", 4),
    ("Ajustar en columnas", "=WRAPCOLS(SEQUENCE(5),2)", 3), ("Expandir", "=EXPAND({1,2;3,4},3,3,\"-\")", 4), ("Dividir con varios separadores", '=TEXTSPLIT("uno, dos;tres",{",",";"})', 2),
    ("Dividir sin vacíos", '=TEXTSPLIT("a,,b,,,c",",",,TRUE)', 2), ("Dividir en filas con relleno", '=TEXTSPLIT("a=1;b;c=3","=",";",,,"?")', 4),
    ("LET con LAMBDA", "=LET(doble,LAMBDA(x,x*2),doble(21))", 2), ("LET con matriz", "=LET(v,SEQUENCE(4),SUM(v*v))", 2), ("LET anidado", "=LET(a,2,b,LET(c,3,a*c),a+b)", 2),
    ("Por columna (BYCOL)", "=BYCOL({1,2;3,4},LAMBDA(c,MAX(c)))", 2), ("Mapear dos matrices", "=MAP({1,2;3,4},{10,20;30,40},LAMBDA(a,b,a+b))", 3),
    ("Reducir", "=REDUCE(0,SEQUENCE(10),LAMBDA(a,v,a+v))", 2), ("Reducir a matriz", '=REDUCE("inicio",{"x","y"},LAMBDA(a,v,VSTACK(a,v)))', 4),
    ("Acumulado (SCAN)", "=SCAN(0,{1,2,3,4},LAMBDA(a,v,a+v))", 2), ("Armar matriz", "=MAKEARRAY(3,3,LAMBDA(r,c,r*c))", 4),
    ("Parámetro omitido", "=LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(5)", 2), ("Función con nombre", "=Hipotenusa(3,4)", 2),
    ("Frecuencia", "=FREQUENCY(Datos!E2:E61,{20,40,60})", 5), ("Producto de matrices", "=MMULT({1,2;3,4},{5;6})", 3), ("Inversa", "=MINVERSE({4,7;2,6})", 3),
    ("Determinante", "=MDETERM({4,7;2,6})", 2), ("Identidad", "=MUNIT(3)", 4), ("Suma de un derrame", "=SUM([[Rango por 2]]#)", 2), ("Filas de un derrame", "=ROWS([[Filtrar dos condiciones]]#)", 2),
    ("Derrame de otro derrame", "=SORT([[Largo de cada uno]]#,1,-1)", 6),
    ("Intersección implícita", "=@Datos!E2:E61", 2), ("Bloqueado por texto", "=SEQUENCE(3)", 4), ("Bloqueado por combinada", "=SEQUENCE(2,2)", 3),
]


def x09(x):
    L = Libro(x, "x09-matrices"); b = L.b
    ws = L.hoja("Matrices")
    dat = L.hoja("Datos"); ventas(dat)
    with b.paso("nombre con LAMBDA (Hipotenusa)"):
        L.wb.Names.Add("Hipotenusa", "=LAMBDA(a,b,SQRT(a^2+b^2))")
    ws.Range("A1").Value = "Qué"; ws.Range("B1").Value = "Fórmula"; ws.Range("D1").Value = "Resultado"; ws.Range("A1:D1").Font.Bold = True
    fila, filas = 2, {}
    for nombre, f, alto in MATRICES: filas[nombre] = fila; fila += alto + 1
    fila = 2
    for nombre, f, alto in MATRICES:
        for k, v in filas.items(): f = f.replace("[[" + k + "]]", f"D{v}")
        with b.paso("matriz " + nombre):
            ws.Cells(fila, 1).Value = nombre; ws.Cells(fila, 2).Value = "'" + f
            if nombre == "Bloqueado por texto": ws.Cells(fila + 2, 4).Value = "estorbo"
            if nombre == "Bloqueado por combinada": ws.Range(ws.Cells(fila + 1, 5), ws.Cells(fila + 1, 6)).Merge()
            ws.Cells(fila, 4).Formula2 = f
        fila += alto + 1
    ws.Columns("A").ColumnWidth = 26; ws.Columns("B").ColumnWidth = 44
    return L.guardar()


if __name__ == "__main__":
    pythoncom.CoInitialize()
    hacer = {k: v for k, v in globals().items() if k[:1] == "x" and k[1:].isdigit()}
    pedidos = sys.argv[1:] or sorted(hacer)
    x = excel()
    try:
        for k in pedidos:
            print("==", k, flush=True)
            try: print(" ->", hacer[k](x), flush=True)
            except Exception: traceback.print_exc()
    finally:
        x.Quit()

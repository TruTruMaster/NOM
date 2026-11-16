"""Arma con Word de verdad documentos que usan casi todo lo que Word sabe hacer.

  python hacer_word.py [w01 w02 ...]      (sin nombres: todos)

Cada archivo junta un grupo de funciones (letra, párrafo, listas, secciones, tablas, objetos,
referencias, revisión, un informe largo y combinaciones raras). Deja `archivos\\wNN-*.docx` y
`archivos\\wNN-*.funciones.json` (qué aceptó Word y qué no).
"""
import os, sys, time
from comun import *

U16 = lambda s: len(s.encode("utf-16-le")) // 2   # Word cuenta posiciones en UTF-16


class Doc:
    def __init__(self, w, nombre):
        self.w, self.nombre = w, nombre
        self.d = w.Documents.Add()
        self.b = Bitacora(nombre)
        self.primero = True

    def cm(self, x): return x * 72 / 2.54   # (CentimetersToPoints falla por COM fuera de un documento visible)

    def p(self, txt="", estilo=None, **fmt):
        """Párrafo nuevo al final (con su estilo); devuelve el párrafo."""
        d = self.d
        if self.primero: self.primero = False
        else: d.Content.InsertParagraphAfter()
        par = d.Paragraphs.Last
        par.Range.ParagraphFormat.Reset(); par.Range.Font.Reset()
        par.Style = estilo if estilo is not None else -1
        if txt: par.Range.InsertBefore(txt)
        par = d.Paragraphs.Last
        if fmt: self.fmt(par.Range, **fmt)
        return par

    def r(self, par, txt, **fmt):
        """Agrega texto con formato al final de un párrafo; devuelve su rango."""
        fin = par.Range.End - 1
        rg = self.d.Range(fin, fin)
        rg.InsertAfter(txt)
        rg = self.d.Range(fin, fin + U16(txt))
        if fmt: self.fmt(rg, **fmt)
        return rg

    def fin(self, par):
        """Rango vacío al final de un párrafo (antes de su marca)."""
        e = par.Range.End - 1
        return self.d.Range(e, e)

    def fmt(self, rg, **f):
        fo = rg.Font
        m = {"fuente": "Name", "tam": "Size", "b": "Bold", "i": "Italic", "u": "Underline", "color": "Color",
             "tachado": "StrikeThrough", "doble": "DoubleStrikeThrough", "sup": "Superscript", "sub": "Subscript",
             "versalitas": "SmallCaps", "mayus": "AllCaps", "oculto": "Hidden", "relieve": "Emboss", "grabado": "Engrave",
             "contorno": "Outline", "sombra": "Shadow", "esp": "Spacing", "escala": "Scaling", "pos": "Position",
             "kern": "Kerning", "colorsub": "UnderlineColor"}
        for k, v in f.items():
            if k in m: setattr(fo, m[k], v)
            elif k == "resalte": rg.HighlightColorIndex = v
            elif k == "fondo": rg.Shading.BackgroundPatternColor = v
            elif k == "alin": rg.ParagraphFormat.Alignment = v
            elif k == "tema":
                fo.TextColor.ObjectThemeColor = v[0]
                if len(v) > 1: fo.TextColor.TintAndShade = v[1]
            else: raise KeyError(k)

    def guardar(self):
        d = self.d
        with self.b.paso("actualizar campos e índices"):
            for t in d.TablesOfContents: t.Update()
            for t in d.TablesOfFigures: t.Update()
            for t in d.Indexes: t.Update()
            d.Fields.Update()
        ruta = os.path.join(ARCHIVOS, self.nombre + ".docx")
        d.SaveAs2(ruta, 16)
        guardar_json(os.path.join(ARCHIVOS, self.nombre + ".funciones.json"), self.b.resumen())
        d.Close(0)
        return ruta


FUENTES = ["Aptos", "Aptos Display", "Aptos Narrow", "Calibri", "Calibri Light", "Cambria", "Times New Roman", "Arial",
           "Arial Narrow", "Arial Black", "Courier New", "Georgia", "Verdana", "Tahoma", "Trebuchet MS", "Segoe UI",
           "Consolas", "Garamond", "Book Antiqua", "Century Gothic", "Comic Sans MS", "Impact", "Candara", "Corbel",
           "Constantia", "Franklin Gothic Medium", "Gill Sans MT", "Palatino Linotype", "Lucida Console", "Segoe Script",
           "Segoe Print", "Bahnschrift", "Sitka Text", "Ink Free", "Cascadia Code", "Lucida Sans Unicode", "MS Gothic"]
FRASE = "El veloz murciélago hindú comía feliz cardillo y kiwi. 0123456789 ¿Qué? ¡Sí! ÁÉÍÓÚ ñÑ"


# ---------------------------------------------------------------- w01: letra
def w01(w):
    D = Doc(w, "w01-letra"); b = D.b
    D.p("Formato de letra en Word", -63)
    D.p("Cada renglón usa una fuente distinta, en 11 puntos, con la misma frase para comparar anchos.", None)
    for f in FUENTES:
        with b.paso("fuente " + f):
            par = D.p(); D.r(par, f + ": ", b=True); D.r(par, FRASE, fuente=f)
    D.p("Tamaños", -2)
    for t in [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72]:
        with b.paso(f"tamaño {t}"):
            D.p(f"{t} pt Ágil", None, tam=t)
    D.p("Estilos de letra", -2)
    par = D.p("Normal, ")
    for nombre, f in [("negrita", dict(b=True)), ("cursiva", dict(i=True)), ("negrita cursiva", dict(b=True, i=True)),
                      ("tachado", dict(tachado=True)), ("doble tachado", dict(doble=True)), ("superíndice", dict(sup=True)),
                      ("subíndice", dict(sub=True)), ("Versalitas", dict(versalitas=True)), ("mayúsculas", dict(mayus=True)),
                      ("relieve", dict(relieve=True)), ("grabado", dict(grabado=True)), ("contorno", dict(contorno=True)),
                      ("sombra", dict(sombra=True))]:
        with b.paso("letra " + nombre):
            D.r(par, nombre, **f); D.r(par, ", ")
    D.r(par, "y texto oculto: ");
    with b.paso("texto oculto"): D.r(par, "ESTO NO SE VE", oculto=True)
    D.r(par, " (fin).")
    par = D.p("Subrayados: ")
    for nombre, u in [("sencillo", 1), ("solo palabras", 2), ("doble", 3), ("punteado", 4), ("grueso", 6), ("guiones", 7),
                      ("punto y guion", 9), ("punto punto guion", 10), ("ondulado", 11), ("ondulado doble", 43),
                      ("guion largo", 39), ("punteado grueso", 20), ("ondulado grueso", 27)]:
        with b.paso("subrayado " + nombre): D.r(par, nombre, u=u); D.r(par, "  ")
    par = D.p("Subrayado de color: ")
    with b.paso("subrayado de color"): D.r(par, "rojo debajo de texto azul", u=1, color=rgb(0, 70, 200), colorsub=rgb(220, 0, 0))
    par = D.p("Colores: ")
    for nombre, c in [("rojo", rgb(192, 0, 0)), ("naranja", rgb(237, 125, 49)), ("verde", rgb(0, 176, 80)), ("azul", rgb(0, 112, 192)),
                      ("morado", rgb(112, 48, 160)), ("gris", rgb(128, 128, 128))]:
        with b.paso("color " + nombre): D.r(par, nombre + " ", color=c)
    for nombre, t in [("Énfasis 1", (5,)), ("Énfasis 2 aclarado", (6, 0.4)), ("Énfasis 6 oscurecido", (10, -0.5)), ("Texto 2", (3,))]:
        with b.paso("color de tema " + nombre): D.r(par, nombre + " ", tema=t)
    par = D.p("Resaltado: ")
    for nombre, c in [("amarillo", 7), ("verde brillante", 4), ("turquesa", 3), ("rosa", 5), ("azul", 2), ("rojo", 6), ("azul oscuro", 9),
                      ("verde azulado", 10), ("verde", 11), ("violeta", 12), ("rojo oscuro", 13), ("amarillo oscuro", 14), ("gris 50", 16), ("gris 25", 15), ("negro", 1)]:
        with b.paso("resaltado " + nombre): D.r(par, nombre, resalte=c, color=rgb(255, 255, 255) if c in (1, 2, 9, 12, 13, 16) else 0); D.r(par, " ")
    par = D.p("Sombreado de carácter: ")
    with b.paso("sombreado de carácter"): D.r(par, "fondo verde claro", fondo=rgb(226, 239, 218))
    with b.paso("borde de carácter"):
        rg = D.r(par, " con borde"); rg.Font.Borders(1).LineStyle = 1; rg.Font.Borders(1).Color = rgb(0, 0, 0)
    D.p("Espaciado, escala y posición", -2)
    par = D.p()
    for nombre, f in [("expandido 3 pt", dict(esp=3)), ("comprimido 1 pt", dict(esp=-1)), ("escala 50 %", dict(escala=50)),
                      ("escala 150 %", dict(escala=150)), ("escala 200 %", dict(escala=200)), ("elevado 6 pt", dict(pos=6)),
                      ("bajado 4 pt", dict(pos=-4))]:
        with b.paso(nombre): D.r(par, nombre, **f); D.r(par, " · ")
    with b.paso("interletraje desde 8 pt"): D.p("AVATAR Wave To Ty LT con interletraje (kerning)", None, tam=28, kern=8)
    D.p("AVATAR Wave To Ty LT sin interletraje", None, tam=28)
    D.p("Efectos de texto (Word 2010 y después)", -2)
    with b.paso("efecto resplandor"):
        par = D.p("Texto con resplandor dorado", None, tam=20, b=True); g = par.Range.Font.Glow; g.Radius = 8; g.Color.RGB = rgb(255, 192, 0)
    with b.paso("efecto reflejo"):
        par = D.p("Texto con reflejo", None, tam=24, b=True); par.Range.Font.Reflection.Type = 2
    with b.paso("efecto sombra de texto"):
        par = D.p("Texto con sombra exterior", None, tam=24); s = par.Range.Font.TextShadow; s.Visible = True; s.OffsetX = 3; s.OffsetY = 3; s.ForeColor.RGB = rgb(120, 120, 120)
    with b.paso("efecto contorno y relleno degradado"):
        par = D.p("CONTORNO Y DEGRADADO", None, tam=28, b=True)
        fo = par.Range.Font; fo.Line.Visible = True; fo.Line.ForeColor.RGB = rgb(31, 56, 100); fo.Line.Weight = 1
        fo.Fill.TwoColorGradient(1, 1); fo.Fill.ForeColor.RGB = rgb(0, 112, 192); fo.Fill.BackColor.RGB = rgb(0, 176, 80)
    with b.paso("ligaduras y conjuntos estilísticos"):
        par = D.p("Ligaduras: fi fl ffi ffl — office affluent", None, fuente="Calibri", tam=16); par.Range.Font.Ligatures = 15
        par = D.p("Números de estilo antiguo 0123456789", None, fuente="Cambria", tam=16); par.Range.Font.NumberForm = 2
    D.p("Otros alfabetos y símbolos", -2)
    for nombre, t in [("griego", "Ελληνικά: Καλημέρα κόσμε, αβγδεζηθ"), ("cirílico", "Кириллица: Съешь же ещё этих мягких французских булок"),
                      ("árabe (derecha a izquierda)", "العربية: مرحبا بالعالم ١٢٣"), ("hebreo", "עברית: שלום עולם"),
                      ("chino", "中文：协作编辑文档，没有网络也可以。"), ("japonés", "日本語：こんにちは世界、カタカナ"), ("coreano", "한국어: 안녕하세요 세계"),
                      ("devanagari", "हिन्दी: नमस्ते दुनिया"), ("tailandés", "ภาษาไทย: สวัสดีชาวโลก"),
                      ("emojis", "Emojis: 😀 👍🏽 👨‍👩‍👧 🇲🇽 ❤️ ✅"), ("matemáticos", "Símbolos: ∑ ∫ √ ∞ ≠ ≤ ≥ ± × ÷ → ⇒ ∀ ∃ ∈ ∉ ⊂ ∪ ∩ ℝ ℕ α β γ π Ω"),
                      ("combinados", "Diacríticos combinados: á ë ñ ô ç — y separadores no rompibles")]:
        with b.paso("alfabeto " + nombre): D.p(t, None)
    D.p("Cambios de formato a media palabra", -2)
    with b.paso("formato a media palabra"):
        par = D.p()
        for k, ch in enumerate("Colaboración"):
            D.r(par, ch, color=[rgb(192, 0, 0), rgb(0, 112, 192), rgb(0, 176, 80)][k % 3], tam=12 + k, b=k % 2 == 0, i=k % 3 == 0)
        D.r(par, " y " + texto(0, 2))
    return D.guardar()


# ---------------------------------------------------------------- w02: párrafo
def w02(w):
    D = Doc(w, "w02-parrafo"); b = D.b; cm = D.cm
    D.p("Formato de párrafo", -63)
    for nombre, a in [("izquierda", 0), ("centrado", 1), ("derecha", 2), ("justificado", 3), ("distribuido", 4)]:
        with b.paso("alineación " + nombre):
            D.p("Alineación " + nombre + ". " + texto(1, 2), None, alin=a)
    D.p("Sangrías", -2)
    for nombre, (iz, de, pr) in [("izquierda 2 cm", (2, 0, 0)), ("derecha 3 cm", (0, 3, 0)), ("primera línea 1.25 cm", (0, 0, 1.25)),
                                 ("francesa 1.5 cm", (1.5, 0, -1.5)), ("negativa (sale al margen) -1 cm", (-1, -1, 0)), ("ambas 4 cm", (4, 4, 0))]:
        with b.paso("sangría " + nombre):
            par = D.p("Sangría " + nombre + ": " + texto(2, 2)); f = par.Format
            f.LeftIndent = cm(iz); f.RightIndent = cm(de); f.FirstLineIndent = cm(pr)
    D.p("Espacio antes y después", -2)
    for a, dsp in [(0, 0), (6, 6), (12, 0), (24, 12), (0, 30)]:
        with b.paso(f"espacio {a}/{dsp}"):
            par = D.p(f"Antes {a} pt y después {dsp} pt. " + texto(3)); par.Format.SpaceBefore = a; par.Format.SpaceAfter = dsp
    with b.paso("espacio automático"):
        par = D.p("Espacio automático antes y después (como en HTML). " + texto(4)); par.Format.SpaceBeforeAuto = True; par.Format.SpaceAfterAuto = True
    with b.paso("sin espacio entre párrafos del mismo estilo"):
        for k in range(3):
            par = D.p(f"Párrafo {k + 1} con espacio contextual: no deja espacio con su vecino del mismo estilo. ", -180)
            par.Format.SpaceAfter = 12; par.Format.SpaceBefore = 12
    D.p("Interlineado", -2)
    for nombre, (regla, val) in [("sencillo", (0, None)), ("1.15", (5, 13.8)), ("1.5 líneas", (1, None)), ("doble", (2, None)),
                                 ("mínimo 18 pt", (3, 18)), ("exacto 12 pt", (4, 12)), ("exacto 30 pt", (4, 30)), ("múltiple 3", (5, 36)),
                                 ("exacto 8 pt (la letra se corta)", (4, 8))]:
        with b.paso("interlineado " + nombre):
            par = D.p("Interlineado " + nombre + ". " + texto(5, 2)); par.Format.LineSpacingRule = regla
            if val: par.Format.LineSpacing = val
    D.p("Bordes y sombreado", -2)
    for nombre, (est, ancho, color) in [("sencillo negro", (1, 4, rgb(0, 0, 0))), ("doble azul", (7, 6, rgb(0, 112, 192))),
                                        ("grueso y delgado", (10, 18, rgb(192, 0, 0))), ("punteado", (2, 8, rgb(0, 128, 0))),
                                        ("guiones", (4, 12, rgb(112, 48, 160))), ("ondulado", (18, 6, rgb(237, 125, 49))),
                                        ("relieve 3D", (21, 24, rgb(128, 128, 128))), ("triple", (8, 6, rgb(0, 0, 0)))]:
        with b.paso("borde " + nombre):
            par = D.p("Recuadro " + nombre + ". " + texto(6))
            for lado in (-1, -2, -3, -4):
                bd = par.Borders(lado); bd.LineStyle = est; bd.LineWidth = ancho; bd.Color = color
    with b.paso("borde con sombra"):
        par = D.p("Recuadro con sombra. " + texto(7)); par.Borders.Enable = True; par.Borders.Shadow = True
    with b.paso("solo línea inferior"):
        par = D.p("Solo línea abajo (como raya horizontal)."); bd = par.Borders(-3); bd.LineStyle = 1; bd.LineWidth = 12; bd.Color = rgb(31, 56, 100)
    with b.paso("bordes agrupados"):
        for k in range(3):
            par = D.p(f"Bloque con un solo recuadro, párrafo {k + 1}. " + texto(k))
            for lado in (-1, -2, -3, -4): par.Borders(lado).LineStyle = 1
    for nombre, (fondo, textura) in [("verde claro", (rgb(226, 239, 218), 0)), ("amarillo", (rgb(255, 242, 204), 0)),
                                     ("trama 20 %", (rgb(255, 255, 255), 200)), ("trama 37.5 % (las rayas diagonales hacen fallar el PDF de Word)", (rgb(255, 255, 255), 375))]:
        with b.paso("sombreado " + nombre):
            par = D.p("Sombreado " + nombre + ". " + texto(8)); par.Shading.BackgroundPatternColor = fondo
            if textura: par.Shading.Texture = textura; par.Shading.ForegroundPatternColor = rgb(0, 112, 192)
    D.p("Tabulaciones", -2)
    with b.paso("tabulaciones con relleno"):
        for nombre, num in [("Introducción", "1"), ("Antecedentes del proyecto", "4"), ("Metodología aplicada", "12"), ("Resultados", "27")]:
            par = D.p(nombre + "\t" + num); par.TabStops.ClearAll(); par.TabStops.Add(cm(16), 2, 1)
    with b.paso("tabulaciones de todos los tipos"):
        par = D.p("Izq\tCentro\tDerecha\t1234.56\tBarra"); ts = par.TabStops; ts.ClearAll()
        ts.Add(cm(2), 0, 2); ts.Add(cm(6), 1, 0); ts.Add(cm(10), 2, 3); ts.Add(cm(13), 3, 0); ts.Add(cm(14.5), 4, 0)
        for v in ["7.5", "1,250.00", "-3.75", "100"]:
            par = D.p("Partida\t\t\t" + v); ts = par.TabStops; ts.ClearAll(); ts.Add(cm(13), 3, 1)
    with b.paso("tabulación predeterminada"):
        D.p("Uno\tdos\ttres\tcuatro\tcinco (tabuladores cada 1.25 cm)")
    D.p("Control de páginas", -2)
    for k in range(4):
        D.p(texto(k, 3))
    with b.paso("conservar con el siguiente"):
        par = D.p("Este párrafo se conserva con el siguiente.", None, b=True); par.Format.KeepWithNext = True
    with b.paso("conservar líneas juntas"):
        par = D.p("Conservar líneas juntas: " + texto(9, 5)); par.Format.KeepTogether = True
    with b.paso("salto de página antes"):
        par = D.p("Este párrafo empieza en página nueva (salto de página antes).", None, b=True); par.Format.PageBreakBefore = True
    with b.paso("sin control de viudas"):
        par = D.p("Sin control de líneas viudas y huérfanas: " + texto(10, 6)); par.Format.WidowControl = False
    D.p("Letra capital, saltos y otros", -2)
    with b.paso("letra capital"):
        par = D.p("Había una vez un documento que quería verse igual en todas partes. " + texto(11, 3)); dc = par.DropCap; dc.Position = 1; dc.LinesToDrop = 3; dc.DistanceFromText = 4
    with b.paso("letra capital en el margen"):
        par = D.p("En el margen también puede ir la capital. " + texto(0, 3)); dc = par.DropCap; dc.Position = 2; dc.LinesToDrop = 2
    with b.paso("salto de línea manual"):
        D.p("Primera línea\vsegunda línea con salto manual\vtercera línea")
    with b.paso("salto de página"):
        par = D.p("Antes del salto de página."); D.fin(par).InsertBreak(7)
    D.p("Después del salto de página.")
    with b.paso("derecha a izquierda"):
        par = D.p("فقرة من اليمين إلى اليسار مع نص عربي وأرقام 2026."); par.ReadingOrder = 0
    with b.paso("nivel de esquema en texto normal"):
        par = D.p("Texto normal con nivel de esquema 2 (aparece en el panel de navegación)."); par.OutlineLevel = 2
    with b.paso("párrafo largo justificado"):
        D.p(texto(0, 12), None, alin=3)
    return D.guardar()



# ---------------------------------------------------------------- w03: estilos y listas
def w03(w):
    D = Doc(w, "w03-estilos-listas"); b = D.b; d = D.d; cm = D.cm
    D.p("Estilos y listas", -63)
    D.p("Un subtítulo con el estilo Subtítulo", -75)
    for nivel in range(1, 7):
        with b.paso(f"título {nivel}"):
            D.p(f"Título de nivel {nivel}", -1 - nivel); D.p(texto(nivel, 1))
    for nombre, est in [("Cita", -181), ("Cita destacada", -182), ("Sin espaciado", "No Spacing"), ("Párrafo de lista", -180), ("Descripción", -35)]:
        with b.paso("estilo " + nombre): D.p(nombre + ": " + texto(3, 1), est)
    with b.paso("estilos de carácter"):
        par = D.p("Estilos de carácter: ")
        for nombre, est in [("Énfasis sutil", -261), ("Énfasis", -89), ("Énfasis intenso", -262), ("Texto en negrita", -88),
                            ("Referencia sutil", -263), ("Referencia intensa", -264), ("Título del libro", -265)]:
            rg = D.r(par, nombre); rg.Style = est; D.r(par, ", ")
    with b.paso("estilo propio de párrafo"):
        st = d.Styles.Add("Aviso importante", 1); st.BaseStyle = -1
        st.Font.Name = "Georgia"; st.Font.Size = 13; st.Font.Color = rgb(150, 30, 30); st.Font.Italic = True
        st.ParagraphFormat.LeftIndent = cm(1); st.ParagraphFormat.SpaceBefore = 10; st.ParagraphFormat.SpaceAfter = 10
        st.ParagraphFormat.Borders(-2).LineStyle = 1; st.ParagraphFormat.Borders(-2).LineWidth = 24; st.ParagraphFormat.Borders(-2).Color = rgb(150, 30, 30)
        st.ParagraphFormat.Shading.BackgroundPatternColor = rgb(253, 236, 234)
        D.p("Estilo propio «Aviso importante» basado en Normal. " + texto(7), "Aviso importante")
    with b.paso("estilo propio heredado de otro propio"):
        st2 = d.Styles.Add("Aviso menor", 1); st2.BaseStyle = "Aviso importante"; st2.Font.Size = 10; st2.Font.Color = rgb(80, 80, 80)
        D.p("Estilo «Aviso menor» heredado de «Aviso importante». " + texto(8), "Aviso menor")
    with b.paso("estilo propio de carácter"):
        st3 = d.Styles.Add("Clave", 2); st3.Font.Name = "Consolas"; st3.Font.Color = rgb(0, 97, 0); st3.Font.Shading.BackgroundPatternColor = rgb(235, 245, 235)
        par = D.p("Escribe "); rg = D.r(par, "=PROMEDIO(A1:A10)"); rg.Style = "Clave"; D.r(par, " en la celda para obtener el promedio.")
    D.p("Listas con viñetas", -2)
    gal = w.ListGalleries
    def lista(items, galeria, plantilla, nivel_de=None, continuar=False):
        pars = []
        for k, t in enumerate(items):
            pars.append(D.p(t))
        rg = d.Range(pars[0].Range.Start, pars[-1].Range.End)
        rg.ListFormat.ApplyListTemplateWithLevel(gal(galeria).ListTemplates(plantilla), continuar, 0, 2)
        if nivel_de:
            for k, n in enumerate(nivel_de):
                if n > 1: pars[k].Range.ListFormat.ListLevelNumber = n
        return pars
    for k in range(1, 8):
        with b.paso(f"viñetas de la galería {k}"):
            lista([f"Viñeta estilo {k}: " + texto(k), "Segundo elemento"], 1, k)
    D.p("Listas numeradas", -2)
    for k in range(1, 8):
        with b.paso(f"numeración de la galería {k}"):
            lista([f"Numeración estilo {k}: " + texto(k), "Segundo", "Tercero"], 2, k)
    D.p("Listas de varios niveles", -2)
    for k in range(1, 8):
        with b.paso(f"lista multinivel {k}"):
            lista([f"Nivel uno (plantilla {k})", "Nivel dos", "Nivel tres", "Nivel tres, otro", "Nivel dos otra vez", "Nivel uno otra vez", "Nivel dos bajo el segundo"], 3, k, [1, 2, 3, 3, 2, 1, 2])
    with b.paso("lista que empieza en 5"):
        pars = lista(["Empieza en cinco", "Seis", "Siete"], 2, 1)
        lt = pars[0].Range.ListFormat.ListTemplate; lv = lt.ListLevels(1); lv.StartAt = 5
        pars[0].Range.ListFormat.ApplyListTemplateWithLevel(lt, False, 0, 2)
    with b.paso("lista que continúa tras un párrafo"):
        pars = lista(["Primero de la lista que sigue", "Segundo"], 2, 1)
        D.p("Un párrafo normal en medio de la lista.")
        p2 = D.p("Tercero (continúa la numeración)")
        p2.Range.ListFormat.ApplyListTemplateWithLevel(gal(2).ListTemplates(1), True, 0, 2)
    with b.paso("viñeta con símbolo propio y color"):
        pars = lista(["Viñeta propia de color", "Otra con la misma viñeta"], 1, 1)
        lt = pars[0].Range.ListFormat.ListTemplate; lv = lt.ListLevels(1)
        lv.NumberFormat = chr(0xF0D8); lv.Font.Name = "Wingdings"; lv.Font.Color = rgb(192, 0, 0); lv.Font.Size = 14
        pars[0].Range.ListFormat.ApplyListTemplateWithLevel(lt, True, 0, 2)
    with b.paso("numeración con formato propio"):
        pars = lista(["Paso con formato «Paso 1 —»", "Siguiente paso"], 2, 1)
        lt = pars[0].Range.ListFormat.ListTemplate; lv = lt.ListLevels(1)
        lv.NumberFormat = "Paso %1 —"; lv.NumberStyle = 0; lv.Font.Bold = True; lv.Font.Color = rgb(0, 112, 192); lv.TrailingCharacter = 1
        lv.NumberPosition = cm(0.5); lv.TextPosition = cm(2.5); lv.TabPosition = cm(2.5)
        pars[0].Range.ListFormat.ApplyListTemplateWithLevel(lt, True, 0, 2)
    with b.paso("elemento de lista con varios párrafos"):
        pars = lista(["Elemento con explicación larga"], 2, 1)
        par = D.p(texto(4, 2)); par.Format.LeftIndent = pars[0].Format.LeftIndent
        lista(["Siguiente elemento"], 2, 1, continuar=True)
    with b.paso("títulos numerados (1, 1.1, 1.1.1)"):
        D.p("Títulos numerados", -2)
        lt = gal(3).ListTemplates(7) if gal(3).ListTemplates.Count >= 7 else gal(3).ListTemplates(1)
        for k, (t, n) in enumerate([("Introducción", 1), ("Objetivo", 2), ("Alcance", 2), ("Detalle del alcance", 3), ("Desarrollo", 1), ("Etapa uno", 2)]):
            par = D.p(t, -1 - n)
            par.Range.ListFormat.ApplyListTemplateWithLevel(lt, k > 0, 0, 2)
            D.p(texto(k))
    return D.guardar()


# ---------------------------------------------------------------- w04: secciones, encabezados, columnas
def w04(w):
    D = Doc(w, "w04-secciones"); b = D.b; d = D.d; cm = D.cm
    def nueva_seccion(tipo=2):
        r = d.Content; r.Collapse(0); r.InsertBreak(tipo); D.primero = True
        return d.Sections(d.Sections.Count)
    def campo(story, codigo):
        r = story.Paragraphs.Last.Range; r.MoveEnd(1, -1); r.Collapse(0)
        return d.Fields.Add(r, -1, codigo, False)
    def texto_al_final(story, t):
        r = story.Paragraphs.Last.Range; r.MoveEnd(1, -1); r.Collapse(0); r.InsertAfter(t)
    s1 = d.Sections(1)
    with b.paso("primera sección: carta, márgenes de 2.5 cm"):
        ps = s1.PageSetup; ps.PageWidth = 612; ps.PageHeight = 792
        for k in ("TopMargin", "BottomMargin", "LeftMargin", "RightMargin"): setattr(ps, k, cm(2.5))
    with b.paso("primera página distinta (portada sin encabezado)"):
        s1.PageSetup.DifferentFirstPageHeaderFooter = True
        s1.Headers(2).Range.Text = ""
    with b.paso("encabezado con texto, tabuladores y fecha"):
        h = s1.Headers(1).Range; h.Text = "NOM-COLAB · Informe de secciones\t\t"
        campo(s1.Headers(1).Range, 'DATE \\@ "d \'de\' MMMM \'de\' yyyy"')
    with b.paso("pie con «Página X de Y»"):
        f = s1.Footers(1).Range; f.Text = "Página "; campo(s1.Footers(1).Range, "PAGE"); texto_al_final(s1.Footers(1).Range, " de "); campo(s1.Footers(1).Range, "NUMPAGES")
        s1.Footers(1).Range.ParagraphFormat.Alignment = 1
    D.p("Portada del documento de secciones", -63); D.p("Esta primera página no lleva encabezado ni pie (primera página distinta).", -75)
    for k in range(3): D.p(texto(k, 2))
    with b.paso("salto de página"):
        r = d.Content; r.Collapse(0); r.InsertBreak(7); D.primero = True
    D.p("Página dos de la primera sección", -2)
    for k in range(4): D.p(texto(k + 3, 3))
    with b.paso("sección 2: A4 horizontal, dos columnas con línea, números romanos desde 1"):
        s2 = nueva_seccion(2); ps = s2.PageSetup
        ps.Orientation = 1; ps.PageWidth = 841.9; ps.PageHeight = 595.3
        ps.TextColumns.SetCount(2); ps.TextColumns.LineBetween = True; ps.TextColumns.Spacing = cm(1)
        s2.Headers(1).LinkToPrevious = False; s2.Headers(1).Range.Text = "Sección 2 — hoja horizontal a dos columnas"
        s2.Footers(1).LinkToPrevious = False; s2.Footers(1).Range.Text = ""; campo(s2.Footers(1).Range, "PAGE")
        s2.Footers(1).Range.ParagraphFormat.Alignment = 2
        pn = s2.Footers(1).PageNumbers; pn.NumberStyle = 2; pn.RestartNumberingAtSection = True; pn.StartingNumber = 1
        s2.PageSetup.DifferentFirstPageHeaderFooter = False
    D.p("Dos columnas en hoja horizontal", -2)
    for k in range(6): D.p(texto(k, 3))
    with b.paso("salto de columna"):
        par = D.p("Antes del salto de columna."); D.fin(par).InsertBreak(8); D.primero = False
    D.p("Este texto empieza en la columna siguiente por un salto de columna. " + texto(5, 2))
    with b.paso("sección 3 continua: tres columnas de ancho distinto"):
        s3 = nueva_seccion(3); tc = s3.PageSetup.TextColumns; tc.SetCount(3); tc.EvenlySpaced = False
        tc(1).Width = cm(4); tc(2).Width = cm(7); tc(3).Width = cm(7)
    D.p("Tres columnas desiguales (sección continua)", -3)
    for k in range(5): D.p(texto(k + 2, 2))
    with b.paso("sección 4: salto de sección"):
        s4 = nueva_seccion(2); ps = s4.PageSetup
    with b.paso("sección 4: una columna"):
        # primero una columna de ancho automático: la columna fija heredada (700 pt) no cabe en vertical
        ps.TextColumns.EvenlySpaced = True; ps.TextColumns.SetCount(1)
    with b.paso("sección 4: oficio vertical"):
        ps.Orientation = 0; ps.PageWidth = 612; ps.PageHeight = 1008
    with b.paso("sección 4: centrada vertical, borde de página, números de línea"):
        ps.VerticalAlignment = 1
        ps.LineNumbering.Active = True; ps.LineNumbering.CountBy = 5; ps.LineNumbering.RestartMode = 0
        s4.Borders.Enable = True
        for lado in (-1, -2, -3, -4):
            bd = s4.Borders(lado); bd.LineStyle = 7; bd.LineWidth = 6; bd.Color = rgb(31, 78, 121)
        s4.Borders.DistanceFrom = 1; s4.Borders.DistanceFromTop = 24; s4.Borders.DistanceFromLeft = 24; s4.Borders.DistanceFromBottom = 24; s4.Borders.DistanceFromRight = 24
        s4.Headers(1).LinkToPrevious = False; s4.Headers(1).Range.Text = "Sección 4 — tamaño oficio, centrado vertical y números de línea"
    D.p("Página centrada verticalmente", -2)
    for k in range(3): D.p(texto(k + 6, 2))
    with b.paso("sección 5: salto a página impar, carta"):
        s5 = nueva_seccion(5); ps = s5.PageSetup
        ps.PageWidth = 612; ps.PageHeight = 792; ps.VerticalAlignment = 0; ps.LineNumbering.Active = False
        s5.Borders.Enable = False
    with b.paso("páginas pares e impares distintas"):
        d.PageSetup.OddAndEvenPagesHeaderFooter = True
    with b.paso("encabezado con tabla de 3 celdas"):
        h = s5.Headers(1); h.LinkToPrevious = False; h.Range.Text = ""
        t = h.Range.Tables.Add(h.Range, 1, 3); t.Borders.Enable = True
        t.Cell(1, 2).Range.Text = "Encabezado de páginas impares (tabla de 3 celdas)"
    with b.paso("imagen dentro del encabezado"):
        rc = t.Cell(1, 1).Range; rc.Collapse(1)
        # (InlineShapes.AddPicture falla dentro de un encabezado si la ventana no está a la vista: se pone
        # flotante en el encabezado y luego se vuelve en línea)
        im = s5.Headers(1).Shapes.AddPicture(os.path.join(IMG, "logo.png"), False, True, 0, 0, 28, 28, rc)
        try: im.ConvertToInlineShape()
        except Exception: pass
    with b.paso("número de página dentro de la tabla del encabezado"):
        r3 = t.Cell(1, 3).Range; r3.Collapse(1); r3.Fields.Add(r3, 33)
    with b.paso("encabezado y pies de páginas pares"):
        he = s5.Headers(3); he.LinkToPrevious = False; he.Range.Text = "Encabezado de páginas PARES"; he.Range.ParagraphFormat.Alignment = 2
        fo = s5.Footers(1); fo.LinkToPrevious = False; fo.Range.Text = "Pie impar"; fo.Range.ParagraphFormat.Alignment = 2
        fe = s5.Footers(3); fe.LinkToPrevious = False; fe.Range.Text = "Pie par"; fe.Range.ParagraphFormat.Alignment = 0
    with b.paso("marca de agua BORRADOR"):
        wm = s5.Headers(1).Shapes.AddTextEffect(0, "BORRADOR", "Arial Black", 1, False, False, 0, 0)
        wm.Rotation = 315; wm.LockAspectRatio = True; wm.Height = cm(5); wm.Width = cm(17)
        wm.Fill.ForeColor.RGB = rgb(200, 200, 200); wm.Fill.Transparency = 0.5; wm.Line.Visible = False
        wm.WrapFormat.Type = 5; wm.RelativeHorizontalPosition = 0; wm.RelativeVerticalPosition = 0; wm.Left = -999995; wm.Top = -999995
    D.p("Páginas pares e impares", -2)
    for k in range(9): D.p(texto(k, 3))
    with b.paso("color de página"):
        d.Background.Fill.ForeColor.RGB = rgb(253, 250, 240); d.Background.Fill.Visible = True
        d.ActiveWindow.View.DisplayBackgrounds = True; w.Options.PrintBackgrounds = True
    with b.paso("sección 6: A5 con márgenes en espejo y encuadernación"):
        s6 = nueva_seccion(2); ps = s6.PageSetup; ps.PageWidth = 419.55; ps.PageHeight = 595.3
        ps.TopMargin = cm(1.5); ps.BottomMargin = cm(1.5); ps.LeftMargin = cm(2); ps.RightMargin = cm(1.2); ps.Gutter = cm(0.8)
    D.p("Hoja A5", -2)
    for k in range(4): D.p(texto(k + 1, 2))
    return D.guardar()


# ---------------------------------------------------------------- w05: tablas
def w05(w):
    D = Doc(w, "w05-tablas"); b = D.b; d = D.d; cm = D.cm
    D.p("Tablas", -63)
    def tabla(nr, nc, datos=None, estilo=None):
        par = D.p("")
        t = d.Tables.Add(par.Range, nr, nc, 1, 0)
        if estilo: t.Style = estilo
        if datos:
            for i, fila in enumerate(datos):
                for j, v in enumerate(fila):
                    if v is not None and i < nr and j < nc: t.Cell(i + 1, j + 1).Range.Text = str(v)
        D.primero = False
        return t
    DAT = [["Materia", "Parcial 1", "Parcial 2", "Final", "Promedio"], ["Cálculo", "8.5", "9.0", "8.0", "8.5"], ["Física", "7.0", "8.5", "9.5", "8.33"],
           ["Química", "9.5", "9.0", "10", "9.5"], ["Programación", "10", "9.5", "9.0", "9.5"], ["Inglés", "8.0", "8.0", "8.5", "8.17"]]
    with b.paso("tabla con cuadrícula"):
        D.p("Tabla con cuadrícula (Table Grid)", -2); tabla(6, 5, DAT, "Table Grid")
    for est in ["Grid Table 4 - Accent 1", "Grid Table 5 Dark - Accent 6", "List Table 3 - Accent 2", "List Table 6 Colorful - Accent 4",
                "Grid Table 1 Light", "Plain Table 3", "Grid Table 7 Colorful - Accent 5", "List Table 1 Light - Accent 3"]:
        with b.paso("estilo de tabla " + est):
            D.p("Estilo: " + est, -3); t = tabla(6, 5, DAT, est)
            t.ApplyStyleHeadingRows = True; t.ApplyStyleFirstColumn = True; t.ApplyStyleRowBands = True; t.ApplyStyleLastRow = est.startswith("Grid Table 4")
    with b.paso("celdas combinadas"):
        D.p("Celdas combinadas", -2)
        t = tabla(5, 5, [["Encabezado combinado en tres", None, None, "D", "E"], ["Vertical en tres", "b2", "c2", "d2", "e2"], [None, "b3", "c3 y d3 juntas", None, "e3"], [None, "b4", "c4", "d4", "e4"], ["a5", "b5 a e5 combinadas", None, None, None]], "Table Grid")
        t.Cell(5, 2).Merge(t.Cell(5, 5)); t.Cell(3, 3).Merge(t.Cell(3, 4)); t.Cell(2, 1).Merge(t.Cell(4, 1)); t.Cell(1, 1).Merge(t.Cell(1, 3))
    with b.paso("tabla dentro de tabla"):
        D.p("Tabla dentro de una celda", -2)
        t = tabla(2, 2, [["Celda normal", "Aquí va otra tabla:"], ["Abajo a la izquierda", None]], "Table Grid")
        t2 = t.Cell(2, 2).Tables.Add(t.Cell(2, 2).Range, 2, 3); t2.Style = "Grid Table 4 - Accent 2"
        for i in range(2):
            for j in range(3): t2.Cell(i + 1, j + 1).Range.Text = f"anidada {i + 1}.{j + 1}"
    with b.paso("sombreado, bordes y diagonales por celda"):
        D.p("Bordes y sombreado por celda", -2)
        t = tabla(4, 4, [["sin bordes", "rojo grueso abajo", "diagonal", "fondo"], ["1", "2", "3", "4"], ["5", "6", "7", "8"], ["9", "10", "11", "12"]])
        t.Borders.Enable = False
        c = t.Cell(1, 2).Borders(-3); c.LineStyle = 1; c.LineWidth = 24; c.Color = rgb(192, 0, 0)
        c = t.Cell(1, 3).Borders(-7); c.LineStyle = 1; c.Color = rgb(0, 0, 0)
        t.Cell(1, 4).Shading.BackgroundPatternColor = rgb(255, 230, 153)
        for i in range(2, 5):
            for j in range(1, 5):
                cel = t.Cell(i, j)
                cel.Shading.BackgroundPatternColor = [rgb(222, 235, 247), rgb(255, 255, 255)][(i + j) % 2]
                for lado in (-1, -2, -3, -4): bd = cel.Borders(lado); bd.LineStyle = [1, 7, 2, 4][j - 1]; bd.Color = rgb(68, 114, 196)
    with b.paso("alineación vertical, dirección del texto y márgenes de celda"):
        D.p("Alineación vertical y texto girado", -2)
        t = tabla(2, 5, [["arriba", "centro", "abajo", "texto hacia arriba", "texto hacia abajo"], ["Celda alta para ver la alineación vertical. " * 2, "", "", "", ""]], "Table Grid")
        t.Rows(1).Height = cm(3); t.Rows(1).HeightRule = 2
        for j, v in enumerate([0, 1, 3, 1, 1]): t.Cell(1, j + 1).VerticalAlignment = v
        t.Cell(1, 4).Range.Orientation = 2; t.Cell(1, 5).Range.Orientation = 3
        t.TopPadding = 6; t.BottomPadding = 6; t.LeftPadding = 10; t.RightPadding = 10
    with b.paso("anchos de columna fijos y tabla centrada"):
        D.p("Columnas de 2, 7 y 3.5 cm, tabla centrada", -2)
        t = tabla(3, 3, [["2 cm", "7 cm de ancho", "3.5 cm"], ["a", "b", "c"], ["d", "e", "f"]], "Table Grid")
        t.AllowAutoFit = False; t.Columns(1).Width = cm(2); t.Columns(2).Width = cm(7); t.Columns(3).Width = cm(3.5); t.Rows.Alignment = 1
    with b.paso("tabla al 100 % del ancho y sangría"):
        D.p("Tabla al 100 % del ancho", -2)
        t = tabla(2, 4, [["ancho", "completo", "de la", "página"], ["1", "2", "3", "4"]], "Table Grid")
        t.PreferredWidthType = 2; t.PreferredWidth = 100
        D.p("Tabla con sangría de 2 cm", -3)
        t = tabla(2, 2, [["sangría", "2 cm"], ["x", "y"]], "Table Grid"); t.Rows.LeftIndent = cm(2)
    with b.paso("fórmulas en celdas"):
        D.p("Fórmulas en la tabla (=SUM(ABOVE))", -2)
        t = tabla(5, 3, [["Concepto", "Cantidad", "Importe"], ["Papel", "3", "450.50"], ["Tinta", "2", "1,280.00"], ["Engargolado", "5", "175.25"], ["Total", "", ""]], "Grid Table 4 - Accent 6")
        t.Cell(5, 2).Formula("=SUM(ABOVE)", "0"); t.Cell(5, 3).Formula("=SUM(ABOVE)", "$#,##0.00")
        for i in range(2, 6): t.Cell(i, 3).Range.ParagraphFormat.Alignment = 2
    with b.paso("imagen y lista dentro de celdas"):
        D.p("Imagen y lista dentro de celdas", -2)
        t = tabla(1, 2, None, "Table Grid")
        t.Cell(1, 1).Range.InlineShapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True).Width = cm(6)
        rg = t.Cell(1, 2).Range; rg.Text = "Uno\rDos\rTres"; rg.ListFormat.ApplyListTemplateWithLevel(w.ListGalleries(1).ListTemplates(1), False, 0, 2)
    with b.paso("altura de fila exacta y mínima"):
        D.p("Filas de altura exacta (1 cm, el texto se corta) y mínima (2 cm)", -2)
        t = tabla(2, 2, [["exacta 1 cm: " + texto(0), "x"], ["mínima 2 cm", "y"]], "Table Grid")
        t.Rows(1).Height = cm(1); t.Rows(1).HeightRule = 2; t.Rows(2).Height = cm(2); t.Rows(2).HeightRule = 1
    with b.paso("tabla flotante con texto alrededor"):
        D.p("Tabla flotante", -2)
        t = tabla(3, 2, [["Flotante", "con texto"], ["alrededor", "a la derecha"], ["1", "2"]], "Grid Table 4 - Accent 3")
        t.PreferredWidthType = 3; t.PreferredWidth = cm(7); t.Rows.WrapAroundText = True; t.Rows.HorizontalPosition = cm(0); t.Rows.DistanceRight = cm(0.4)
        D.p(texto(1, 6))
    with b.paso("tabla larga que se parte en páginas con encabezado repetido"):
        D.p("Tabla larga con encabezado que se repite en cada página", -2)
        filas = [["#", "Nombre", "Grupo", "Calificación"]] + [[str(k), f"Estudiante {k:03d}", f"5IM{k % 9 + 1}", f"{6 + (k * 37 % 40) / 10:.1f}"] for k in range(1, 81)]
        t = tabla(len(filas), 4, filas, "Grid Table 4 - Accent 1")
        t.Rows(1).HeadingFormat = True
        t.Rows(10).AllowBreakAcrossPages = False
    D.p("Fin del documento de tablas.")
    return D.guardar()



def smartart(w, nombre):
    for k in range(1, w.SmartArtLayouts.Count + 1):
        if w.SmartArtLayouts(k).Name == nombre: return w.SmartArtLayouts(k)
    raise KeyError(nombre)


def llenar_smartart(forma, textos):
    sa = forma.SmartArt
    while sa.AllNodes.Count < len(textos): sa.AllNodes.Add()
    for k, t in enumerate(textos):
        if k < sa.AllNodes.Count: sa.AllNodes(k + 1).TextFrame2.TextRange.Text = t


def reintentar(f, veces=20):
    """Office ocupado («Call was rejected by callee»): se espera y se vuelve a intentar."""
    for k in range(veces):
        try: return f()
        except Exception as e:
            if getattr(e, "hresult", None) in (-2147418111, -2147417846) and k < veces - 1: time.sleep(1.5); continue
            raise


def datos_grafica(ch, tipo):
    """Cierra la hoja de datos que deja abierta la gráfica nueva (si queda abierta, la siguiente falla) y pone
    datos propios directo en las series (abrir la hoja para editarla hace fallar a Word por automatización)."""
    for k in range(20):
        try: ch.ChartData.Workbook.Close(); break
        except Exception as e:
            if getattr(e, "hresult", None) in (-2147418111, -2147417846): time.sleep(1); continue
            break
    meses = ["Enero", "Febrero", "Marzo", "Abril"]
    series = [("Ventas", [120, 150, 90, 180]), ("Costos", [80, 95, 70, 110]), ("Utilidad", [40, 55, 20, 70])]
    if tipo in (5, -4120): series = [("Participación", [45, 30, 15, 10])]; meses = ["Web", "Teléfono", "Sin red", "Otros"]
    sc = ch.SeriesCollection()
    for i in range(1, sc.Count + 1):
        if i > len(series): continue
        sr = ch.SeriesCollection(i); sr.Name = series[i - 1][0]; sr.Values = series[i - 1][1]
        if tipo != -4169: sr.XValues = meses


ECUACIONES = ["x=(-b±√(b^2-4ac))/(2a)", "∑_(i=1)^n▒i^2 =(n(n+1)(2n+1))/6", "∫_0^∞▒e^(-x^2)  dx=√π/2", "A=[■(a&b@c&d)]",
              "lim_(x→0)⁡〖sin⁡x/x〗=1", "(a+b)^n=∑_(k=0)^n▒(n¦k) a^k b^(n-k)", "f(x)={■(x^2&x≥0@-x&x<0)┤"]


def ecuacion(D, par, lineal, en_linea=False):
    rg = D.fin(par); rg.Text = lineal
    om = D.d.OMaths.Add(rg); om.OMaths(1).BuildUp()
    if en_linea: om.OMaths(1).Type = 1
    return om


# ---------------------------------------------------------------- w06: imágenes, formas y objetos
def w06(w):
    D = Doc(w, "w06-objetos"); b = D.b; d = D.d; cm = D.cm
    D.p("Imágenes, formas y objetos", -63)
    with b.paso("imagen en línea"):
        par = D.p("Imagen en línea con el texto: "); im = d.InlineShapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, D.fin(par)); im.Width = cm(5); im.Height = cm(3.125)
        D.r(par, " y el texto sigue después de la imagen.")
    with b.paso("imagen en línea con borde, sombra y reflejo"):
        par = D.p(""); im = d.InlineShapes.AddPicture(os.path.join(IMG, "foto2.jpg"), False, True, D.fin(par)); im.Width = cm(6); im.Height = cm(6)
        im.Line.Visible = True; im.Line.Weight = 4; im.Line.ForeColor.RGB = rgb(255, 255, 255)
        im.Shadow.Visible = True; im.Shadow.OffsetX = 4; im.Shadow.OffsetY = 4; im.Shadow.Blur = 6; im.Shadow.Transparency = 0.5
        try: im.Reflection.Type = 1
        except Exception: pass
    ajustes = [("cuadrado a la izquierda", 0, 0), ("estrecho a la derecha", 1, 2), ("arriba y abajo", 4, 0), ("detrás del texto", 5, 1), ("delante del texto", 3, 2), ("a través", 2, 1)]
    for k, (nombre, tipo, lado) in enumerate(ajustes):
        with b.paso("imagen flotante, ajuste " + nombre):
            D.p("Ajuste de texto: " + nombre, -3)
            par = D.p(texto(k, 4))
            sh = d.Shapes.AddPicture(os.path.join(IMG, ["foto1.jpg", "logo.png", "foto2.jpg", "foto1.jpg", "logo.png", "foto3-vertical.jpg"][k]), False, True, 0, 0, cm(4.5), cm(3.2), par.Range)
            if k in (1, 4): sh.Height = cm(3.2); sh.Width = cm(3.2)
            sh.WrapFormat.Type = tipo
            sh.RelativeHorizontalPosition = 0; sh.RelativeVerticalPosition = 2
            sh.Left = [-999998, -999996, -999995, cm(3), cm(8), -999998][k]; sh.Top = cm(0.2)
            if tipo in (0, 1, 2): sh.WrapFormat.DistanceLeft = 8; sh.WrapFormat.DistanceRight = 8
            if tipo == 5: sh.Fill.Transparency = 0; sh.PictureFormat.TransparentBackground = False
            D.p(texto(k + 3, 3))
    with b.paso("imagen fija a la página (posición absoluta)"):
        par = D.p("Imagen anclada a la página, a 2 cm del borde derecho y 3 cm de arriba.")
        sh = d.Shapes.AddPicture(os.path.join(IMG, "logo.png"), False, True, 0, 0, cm(2.5), cm(2.5), par.Range)
        sh.RelativeHorizontalPosition = 1; sh.RelativeVerticalPosition = 1; sh.Left = cm(21.59 - 2 - 2.5); sh.Top = cm(3); sh.WrapFormat.Type = 3
    with b.paso("imagen recortada, girada, en escala de grises"):
        par = D.p("Recortada, girada 15° y en escala de grises: ")
        sh = d.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 0, 0, cm(6), cm(3.75), par.Range)
        sh.PictureFormat.CropLeft = 60; sh.PictureFormat.CropRight = 40; sh.PictureFormat.CropTop = 20
        sh.Rotation = 15; sh.PictureFormat.ColorType = 2; sh.WrapFormat.Type = 4; sh.Left = -999995
        D.p(texto(2, 2))
    with b.paso("imagen con brillo y contraste"):
        par = D.p(""); im = d.InlineShapes.AddPicture(os.path.join(IMG, "foto2.jpg"), False, True, D.fin(par)); im.Width = cm(4); im.Height = cm(4)
        im.PictureFormat.Brightness = 0.7; im.PictureFormat.Contrast = 0.3
    with b.paso("imagen SVG"):
        par = D.p("Ícono SVG: "); im = d.InlineShapes.AddPicture(os.path.join(IMG, "icono.svg"), False, True, D.fin(par)); im.Width = cm(2); im.Height = cm(2)
    D.p("Formas", -2)
    formas = [(1, "Rectángulo"), (5, "Rectángulo redondeado"), (9, "Elipse"), (7, "Triángulo"), (33, "Flecha derecha"), (92, "Estrella de 5 puntas"),
              (21, "Corazón"), (22, "Rayo"), (63, "Decisión"), (106, "Llamada"), (103, "Onda"), (52, "Hexágono"), (17, "Cubo"), (16, "Cilindro"), (108, "Nube")]
    par = D.p("Fila de formas (con texto dentro y relleno de colores distintos):")
    ancla = D.p("")
    for k, (tipo, nombre) in enumerate(formas):
        with b.paso("forma " + nombre):
            sh = d.Shapes.AddShape(tipo, cm(0.2 + (k % 5) * 3.3), cm(0.3 + (k // 5) * 3.2), cm(3), cm(2.6), ancla.Range)
            sh.RelativeHorizontalPosition = 0; sh.RelativeVerticalPosition = 2
            sh.Fill.ForeColor.RGB = [rgb(68, 114, 196), rgb(237, 125, 49), rgb(112, 173, 71), rgb(255, 192, 0), rgb(91, 155, 213)][k % 5]
            sh.Line.ForeColor.RGB = rgb(40, 40, 40); sh.Line.Weight = 1
            sh.TextFrame.TextRange.Text = nombre; sh.TextFrame.TextRange.Font.Size = 9; sh.TextFrame.TextRange.Font.Color = rgb(255, 255, 255)
            sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1; sh.TextFrame.VerticalAnchor = 3; sh.WrapFormat.Type = 4
    for k in range(10): D.p("")
    with b.paso("relleno degradado, trama e imagen"):
        par = D.p("Rellenos: degradado, trama, imagen y transparencia")
        a1 = D.p("")
        g = d.Shapes.AddShape(1, cm(0), cm(0.3), cm(4), cm(2.5), a1.Range); g.Fill.TwoColorGradient(1, 1); g.Fill.ForeColor.RGB = rgb(0, 112, 192); g.Fill.BackColor.RGB = rgb(255, 255, 255)
        t = d.Shapes.AddShape(1, cm(4.5), cm(0.3), cm(4), cm(2.5), a1.Range); t.Fill.Patterned(5); t.Fill.ForeColor.RGB = rgb(192, 0, 0); t.Fill.BackColor.RGB = rgb(255, 242, 204)
        i2 = d.Shapes.AddShape(9, cm(9), cm(0.3), cm(4), cm(2.5), a1.Range); i2.Fill.UserPicture(os.path.join(IMG, "foto1.jpg"))
        tr = d.Shapes.AddShape(1, cm(13.5), cm(0.3), cm(3), cm(2.5), a1.Range); tr.Fill.ForeColor.RGB = rgb(0, 176, 80); tr.Fill.Transparency = 0.6
        for x in (g, t, i2, tr): x.RelativeHorizontalPosition = 0; x.RelativeVerticalPosition = 2; x.WrapFormat.Type = 4
        for k in range(5): D.p("")
    with b.paso("líneas con puntas de flecha y guiones"):
        a2 = D.p("Líneas y flechas:")
        for k, (dash, ini, fin) in enumerate([(1, 1, 2), (4, 3, 2), (3, 2, 5), (7, 6, 4)]):
            ln = d.Shapes.AddLine(cm(0.5), cm(1 + k * 0.7), cm(12), cm(1 + k * 0.7), a2.Range)
            ln.Line.DashStyle = dash; ln.Line.BeginArrowheadStyle = ini; ln.Line.EndArrowheadStyle = fin; ln.Line.Weight = 2 + k; ln.Line.ForeColor.RGB = rgb(60, 60, 60)
            ln.RelativeHorizontalPosition = 0; ln.RelativeVerticalPosition = 2
        for k in range(4): D.p("")
    with b.paso("efectos: sombra, resplandor, bordes suaves, bisel y giro 3D"):
        a3 = D.p("Efectos de forma:")
        e1 = d.Shapes.AddShape(5, cm(0), cm(0.5), cm(3.5), cm(2.5), a3.Range); e1.Shadow.Visible = True; e1.Shadow.OffsetX = 6; e1.Shadow.OffsetY = 6; e1.Shadow.Blur = 8
        e2 = d.Shapes.AddShape(9, cm(4.2), cm(0.5), cm(3.5), cm(2.5), a3.Range); e2.Glow.Radius = 10; e2.Glow.Color.RGB = rgb(255, 192, 0)
        e3 = d.Shapes.AddShape(1, cm(8.4), cm(0.5), cm(3.5), cm(2.5), a3.Range); e3.SoftEdge.Radius = 8
        e4 = d.Shapes.AddShape(1, cm(12.6), cm(0.5), cm(3.5), cm(2.5), a3.Range); e4.ThreeD.BevelTopType = 3; e4.ThreeD.BevelTopDepth = 6; e4.ThreeD.BevelTopInset = 6
        e5 = d.Shapes.AddShape(17, cm(0), cm(3.5), cm(3.5), cm(2.5), a3.Range); e5.ThreeD.RotationX = 30; e5.ThreeD.RotationY = 20; e5.ThreeD.Depth = 20
        for x in (e1, e2, e3, e4, e5): x.RelativeHorizontalPosition = 0; x.RelativeVerticalPosition = 2; x.WrapFormat.Type = 4; x.Fill.ForeColor.RGB = rgb(91, 155, 213)
        for k in range(8): D.p("")
    with b.paso("cuadro de texto con formato"):
        a4 = D.p("Cuadro de texto:")
        tb = d.Shapes.AddTextbox(1, cm(1), cm(0.5), cm(8), cm(3.5), a4.Range)
        tb.TextFrame.TextRange.Text = "Cuadro de texto con letra Georgia, márgenes internos amplios, fondo claro y borde punteado. " + texto(0)
        tb.TextFrame.TextRange.Font.Name = "Georgia"; tb.TextFrame.TextRange.Font.Size = 10
        tb.TextFrame.MarginLeft = 12; tb.TextFrame.MarginTop = 10; tb.Fill.ForeColor.RGB = rgb(242, 248, 238); tb.Line.DashStyle = 3; tb.Line.ForeColor.RGB = rgb(0, 128, 0)
        tb.RelativeHorizontalPosition = 0; tb.RelativeVerticalPosition = 2; tb.WrapFormat.Type = 0
        D.p(texto(4, 4))
    with b.paso("cuadros de texto enlazados"):
        a5 = D.p("Cuadros enlazados (el texto pasa del primero al segundo):")
        t1 = d.Shapes.AddTextbox(1, cm(0), cm(0.5), cm(6), cm(2.2), a5.Range); t2 = d.Shapes.AddTextbox(1, cm(8), cm(0.5), cm(6), cm(2.2), a5.Range)
        t1.TextFrame.Next = t2.TextFrame
        t1.TextFrame.TextRange.Text = texto(1, 4)
        for x in (t1, t2): x.RelativeHorizontalPosition = 0; x.RelativeVerticalPosition = 2; x.WrapFormat.Type = 4
        for k in range(5): D.p("")
    with b.paso("cuadro de texto girado y texto vertical"):
        a6 = D.p("Cuadro girado 20° y otro con texto vertical:")
        g1 = d.Shapes.AddTextbox(1, cm(1), cm(1), cm(5), cm(2), a6.Range); g1.TextFrame.TextRange.Text = "Girado veinte grados"; g1.Rotation = 20
        g2 = d.Shapes.AddTextbox(1, cm(9), cm(0.5), cm(2), cm(5), a6.Range); g2.TextFrame.TextRange.Text = "Texto vertical"; g2.TextFrame.Orientation = 2
        for x in (g1, g2): x.RelativeHorizontalPosition = 0; x.RelativeVerticalPosition = 2; x.WrapFormat.Type = 4
        for k in range(7): D.p("")
    with b.paso("WordArt"):
        a7 = D.p("WordArt:")
        for k, efecto in enumerate([0, 8, 15, 22]):
            wa = d.Shapes.AddTextEffect(efecto, "NOM-COLAB", "Arial Black", 28, True, False, cm(0.2 + k * 4.2), cm(0.5), a7.Range)
            wa.RelativeHorizontalPosition = 0; wa.RelativeVerticalPosition = 2; wa.WrapFormat.Type = 4
        for k in range(4): D.p("")
    with b.paso("grupo de formas"):
        a8 = D.p("Grupo de formas:")
        x1 = d.Shapes.AddShape(9, cm(1), cm(0.5), cm(3), cm(3), a8.Range); x2 = d.Shapes.AddShape(1, cm(3), cm(1.5), cm(4), cm(1.5), a8.Range)
        x3 = d.Shapes.AddShape(92, cm(6), cm(0.3), cm(2.5), cm(2.5), a8.Range)
        for k, x in enumerate((x1, x2, x3)): x.WrapFormat.Type = 3; x.Name = f"grupo{k + 1}"
        nombres = W.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["grupo1", "grupo2", "grupo3"])
        try: gr = d.Shapes.Range(nombres).Group()
        except Exception:
            for x in (x1, x2, x3): x.RelativeHorizontalPosition = 0; x.RelativeVerticalPosition = 2
            d.Shapes.Range(nombres).Select(); gr = w.Selection.ShapeRange.Group()
        gr.WrapFormat.Type = 4; gr.RelativeHorizontalPosition = 0; gr.RelativeVerticalPosition = 2; gr.Rotation = 10
        for k in range(5): D.p("")
    with b.paso("lienzo de dibujo"):
        a9 = D.p("Lienzo de dibujo:")
        cv = d.Shapes.AddCanvas(cm(0.5), cm(0.5), cm(12), cm(4), a9.Range)
        c1 = cv.CanvasItems.AddShape(1, 10, 20, 80, 40); c2 = cv.CanvasItems.AddShape(9, 200, 50, 80, 50)
        ln = cv.CanvasItems.AddLine(90, 40, 200, 75); ln.Line.EndArrowheadStyle = 2
        cv.RelativeHorizontalPosition = 0; cv.RelativeVerticalPosition = 2; cv.WrapFormat.Type = 4
        for k in range(6): D.p("")
    with b.paso("conector en codo entre dos formas"):
        a11 = D.p("Conector en codo:")
        f1 = d.Shapes.AddShape(1, cm(1), cm(0.5), cm(3), cm(1.5), a11.Range); f2 = d.Shapes.AddShape(9, cm(9), cm(2.5), cm(3), cm(1.5), a11.Range)
        for x in (f1, f2): x.RelativeHorizontalPosition = 0; x.RelativeVerticalPosition = 2; x.WrapFormat.Type = 4
        cx = d.Shapes.AddConnector(2, 10.0, 10.0, 50.0, 50.0); cx.ConnectorFormat.BeginConnect(f1, 4); cx.ConnectorFormat.EndConnect(f2, 2); cx.Line.EndArrowheadStyle = 2
        cx.RerouteConnections()
        for k in range(5): D.p("")
    with b.paso("forma libre"):
        a10 = D.p("Forma libre (polígono):")
        ff = d.Shapes.BuildFreeform(0, cm(1), cm(1)); ff.AddNodes(0, 0, cm(5), cm(0.5)); ff.AddNodes(0, 0, cm(6), cm(3)); ff.AddNodes(1, 0, cm(3), cm(4)); ff.AddNodes(0, 0, cm(1), cm(1))
        fs = ff.ConvertToShape(a10.Range); fs.Fill.ForeColor.RGB = rgb(255, 192, 0); fs.WrapFormat.Type = 4
        for k in range(5): D.p("")
    D.p("SmartArt", -2)
    for nombre, textos in [("Basic Process", ["Planear", "Construir", "Probar", "Entregar"]), ("Organization Chart", ["Dirección", "Asistente", "Área A", "Área B", "Área C"]),
                           ("Basic Cycle", ["Leer", "Escribir", "Revisar", "Publicar", "Comentar"]), ("Basic Venn", ["Web", "Teléfono", "Sin red"]),
                           ("Vertical Bullet List", ["Primera idea", "Segunda idea", "Tercera idea"]), ("Basic Pyramid", ["Base", "Medio", "Cima"])]:
        with b.paso("SmartArt " + nombre):
            D.p(nombre, -3); a = D.p("")
            sa = d.Shapes.AddSmartArt(smartart(w, nombre), cm(0.5), cm(0.2), cm(15), cm(6), a.Range)
            sa.WrapFormat.Type = 4; llenar_smartart(sa, textos)
            for k in range(10): D.p("")
    D.p("Gráficas", -2)
    for nombre, tipo in [("columnas agrupadas", 51), ("circular", 5), ("líneas con marcadores", 65), ("barras apiladas", 58), ("área", 1), ("anillo", -4120), ("dispersión", -4169), ("radial", -4151)]:
        with b.paso("gráfica " + nombre):
            D.p("Gráfica de " + nombre, -3); a = D.p("")
            g = d.InlineShapes.AddChart2(-1, tipo, D.fin(a)); g.Width = cm(12); g.Height = cm(7)
            ch = g.Chart
            datos_grafica(ch, tipo)
            ch.HasTitle = True; ch.ChartTitle.Text = "Gráfica de " + nombre
    D.p("Ecuaciones", -2)
    for k, e in enumerate(ECUACIONES):
        with b.paso("ecuación " + str(k + 1)):
            ecuacion(D, D.p(""), e)
    with b.paso("ecuación dentro del texto"):
        par = D.p("La energía es "); ecuacion(D, par, "E=mc^2", True); D.r(par, " y el área del círculo es "); ecuacion(D, par, "A=πr^2", True); D.r(par, ".")
    return D.guardar()


# ---------------------------------------------------------------- w07: referencias, campos y controles
FUENTES_BIB = [
    ("Gar20", "Book", "García", "Ana", "Trabajo colaborativo en el aula", "2020", "Ciudad de México", "Editorial Académica"),
    ("Lop18", "Book", "López", "Carlos", "Redes y comunicación sin Internet", "2018", "Guadalajara", "Universidad de Guadalajara"),
    ("Mar22", "Book", "Martínez", "Lucía", "Diseño de interfaces para todos", "2022", "Monterrey", "Prensa Técnica"),
]


def w07(w):
    D = Doc(w, "w07-referencias"); b = D.b; d = D.d; cm = D.cm
    D.p("Referencias, campos y controles", -63)
    with b.paso("tabla de contenido"):
        D.p("Contenido", "TOC Heading")
        par = D.p(""); d.TablesOfContents.Add(D.fin(par), True, 1, 3)
    with b.paso("salto de página"):
        par = D.p(""); D.fin(par).InsertBreak(7)
    for k, (t, n) in enumerate([("Introducción", 1), ("Antecedentes", 2), ("Planteamiento", 2), ("Marco teórico", 1), ("Trabajo en equipo", 2),
                                ("Herramientas", 3), ("Metodología", 1), ("Resultados", 1), ("Conclusiones", 1)]):
        D.p(t, -1 - n)
        par = D.p(texto(k, 3))
        if k == 0:
            with b.paso("notas al pie"):
                d.Footnotes.Add(Range=D.fin(par), Text="Esta es una nota al pie con texto suficiente para ocupar más de un renglón en la parte de abajo de la página.")
                par2 = D.p("Otra oración con su propia nota."); d.Footnotes.Add(Range=D.fin(par2), Text="Segunda nota al pie.")
        if k == 1:
            with b.paso("nota al final"):
                d.Endnotes.Add(Range=D.fin(par), Text="Nota al final del documento.")
            with b.paso("nota al pie con marca propia (símbolo)"):
                d.Footnotes.Add(Range=D.fin(par), Reference="*", Text="Nota con asterisco en vez de número.")
        if k == 2:
            with b.paso("marcador y vínculo interno"):
                d.Bookmarks.Add("marca_metodologia", par.Range)
            with b.paso("vínculos a páginas web y correo"):
                p3 = D.p("Visita ")
                rg = D.r(p3, "el sitio del Instituto"); d.Hyperlinks.Add(rg, "https://www.ipn.mx", "", "Abre el sitio del IPN", "el sitio del Instituto")
                D.r(p3, " o escribe a "); rg = D.r(p3, "contacto@ejemplo.mx"); d.Hyperlinks.Add(rg, "mailto:contacto@ejemplo.mx", "", "", "contacto@ejemplo.mx")
                D.r(p3, ". También puedes ir a "); rg = D.r(p3, "la sección marcada"); d.Hyperlinks.Add(rg, "", "marca_metodologia", "", "la sección marcada"); D.r(p3, ".")
        if k == 3:
            with b.paso("imagen con título (Figura 1) y referencia cruzada"):
                w.CaptionLabels.Add("Figura")
                pi = D.p(""); im = d.InlineShapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, D.fin(pi)); im.Width = cm(9); im.Height = cm(5.6)
                pi.Format.Alignment = 1
                im.Range.InsertCaption("Figura", ": Paisaje de prueba", "", 1)
                D.primero = False
                pr = D.p("Como se ve en la "); rg = D.fin(pr); rg.InsertCrossReference("Figura", 3, 1, True); D.r(pr, ", el paisaje tiene montañas.")
            with b.paso("tabla con título (Tabla 1)"):
                w.CaptionLabels.Add("Tabla")
                pt = D.p(""); t = d.Tables.Add(pt.Range, 3, 3, 1, 0); t.Style = "Grid Table 4 - Accent 1"
                for i in range(3):
                    for j in range(3): t.Cell(i + 1, j + 1).Range.Text = f"{i + 1}.{j + 1}"
                t.Range.InsertCaption("Tabla", ": Datos de ejemplo", "", 0); D.primero = False
            with b.paso("referencia cruzada a un título y a su página"):
                pr = D.p("Ver el apartado «"); rg = D.fin(pr); rg.InsertCrossReference("Heading", -1, 1, True)
                D.r(pr, "» en la página "); rg = D.fin(pr); rg.InsertCrossReference("Heading", 7, 1, True); D.r(pr, ".")
        if k == 4:
            with b.paso("citas y bibliografía"):
                for tag, tipo, ap, nom, tit, an, ci, ed in FUENTES_BIB:
                    xml = (f'<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"><b:Tag>{tag}</b:Tag><b:SourceType>{tipo}</b:SourceType>'
                           f'<b:Author><b:Author><b:NameList><b:Person><b:Last>{ap}</b:Last><b:First>{nom}</b:First></b:Person></b:NameList></b:Author></b:Author>'
                           f'<b:Title>{tit}</b:Title><b:Year>{an}</b:Year><b:City>{ci}</b:City><b:Publisher>{ed}</b:Publisher></b:Source>')
                    d.Bibliography.Sources.Add(xml)
                try: d.Bibliography.BibliographyStyle = "APA"
                except Exception: pass
                pc = D.p("El trabajo colaborativo mejora el aprendizaje ")
                for tag, *_ in FUENTES_BIB:
                    rg = D.fin(pc); d.Fields.Add(rg, -1, f"CITATION {tag} \\l 2058", False); D.r(pc, " ")
                D.r(pc, "según varios autores.")
        if k == 5:
            with b.paso("entradas de índice"):
                for palabra in ["colaboración", "documento", "equipo", "red local"]:
                    pe = D.p(f"Párrafo que menciona {palabra} para el índice.")
                    d.Indexes.MarkEntry(pe.Range, palabra)
        if k == 6:
            with b.paso("campos de fecha, autor, número de palabras, secuencia, fórmula y condición"):
                for etiqueta, codigo in [("Fecha", 'DATE \\@ "dd/MM/yyyy"'), ("Autor", "AUTHOR"), ("Archivo", "FILENAME"), ("Palabras", "NUMWORDS"),
                                         ("Páginas", "NUMPAGES"), ("Secuencia", "SEQ Paso"), ("Secuencia", "SEQ Paso"), ("Fórmula", "= 21 * 2 + 0.5"),
                                         ("Condición", 'IF 5 > 3 "cinco es mayor" "no"'), ("Símbolo", "SYMBOL 167 \\f Wingdings")]:
                    pf = D.p(etiqueta + ": "); d.Fields.Add(D.fin(pf), -1, codigo, False)
            with b.paso("controles de contenido"):
                for tipo, nombre in [(1, "Texto sin formato"), (0, "Texto enriquecido"), (6, "Fecha"), (4, "Lista desplegable"), (3, "Cuadro combinado"), (8, "Casilla"), (2, "Imagen")]:
                    pc = D.p(nombre + ": ")
                    cc = d.ContentControls.Add(tipo, D.fin(pc)); cc.Title = nombre
                    if tipo in (0, 1): cc.Range.Text = "valor de " + nombre.lower()
                    if tipo in (3, 4):
                        for o in ["Opción A", "Opción B", "Opción C"]: cc.DropdownListEntries.Add(o, o[-1])
                        cc.DropdownListEntries(2).Select()
                    if tipo == 6: cc.DateDisplayFormat = "dd/MM/yyyy"; cc.Range.Text = "15/11/2026"
                    if tipo == 8: cc.Checked = True
            with b.paso("campos de formulario antiguos"):
                pf = D.p("Formulario antiguo: "); d.FormFields.Add(D.fin(pf), 70); D.r(pf, "  casilla: "); d.FormFields.Add(D.fin(pf), 71)
    with b.paso("tabla de ilustraciones"):
        D.p("Índice de figuras", -2); par = D.p(""); d.TablesOfFigures.Add(D.fin(par), "Figura")
    with b.paso("bibliografía"):
        D.p("Bibliografía", -2); par = D.p(""); d.Fields.Add(D.fin(par), -1, "BIBLIOGRAPHY \\l 2058", False)
    with b.paso("índice alfabético"):
        D.p("Índice alfabético", -2); par = D.p(""); d.Indexes.Add(D.fin(par), 0, True, 0, 2)
    return D.guardar()


# ---------------------------------------------------------------- w08: comentarios y control de cambios
def w08(w):
    D = Doc(w, "w08-revision"); b = D.b; d = D.d
    antes = w.UserName, w.UserInitials
    try:
        D.p("Comentarios y control de cambios", -63)
        pars = [D.p(texto(k, 3)) for k in range(6)]
        with b.paso("comentarios con respuestas"):
            w.UserName = "Ana López"; w.UserInitials = "AL"
            c1 = d.Comments.Add(d.Range(pars[0].Range.Start, pars[0].Range.Start + 20), "¿Podemos citar la fuente de este dato?")
            w.UserName = "Carlos Ruiz"; w.UserInitials = "CR"
            c1.Replies.Add(c1.Scope, "Sí, viene de la encuesta de 2026.")
            c2 = d.Comments.Add(d.Range(pars[1].Range.Start + 10, pars[2].Range.Start + 30), "Este comentario abarca dos párrafos.")
            c3 = d.Comments.Add(d.Range(pars[3].Range.Start, pars[3].Range.Start + 15), "Comentario ya resuelto."); c3.Done = True
        with b.paso("cambios controlados de dos personas"):
            d.TrackRevisions = True
            w.UserName = "Ana López"; w.UserInitials = "AL"
            r = d.Range(pars[0].Range.End - 1, pars[0].Range.End - 1); r.InsertAfter(" Frase agregada con control de cambios.")
            r = d.Range(pars[1].Range.Start, pars[1].Range.Start + 12); r.Delete()
            w.UserName = "Carlos Ruiz"; w.UserInitials = "CR"
            r = d.Range(pars[2].Range.Start, pars[2].Range.Start + 25); r.Font.Bold = True
            pars[3].Format.Alignment = 1
            r = d.Range(pars[4].Range.Start + 5, pars[4].Range.Start + 5); r.InsertAfter("[insertado a media frase] ")
            r = d.Range(pars[5].Range.Start + 20, pars[5].Range.Start + 60); r.Delete()
        with b.paso("cambios en tabla"):
            pt = D.p(""); t = d.Tables.Add(pt.Range, 2, 2, 1, 0); t.Cell(1, 1).Range.Text = "fila"; t.Rows.Add(); D.primero = False
        d.TrackRevisions = False
        D.p("Fin del documento con revisiones.")
    finally:
        w.UserName, w.UserInitials = antes
    return D.guardar()


# ---------------------------------------------------------------- w09: informe largo con todo
def w09(w):
    D = Doc(w, "w09-informe"); b = D.b; d = D.d; cm = D.cm
    with b.paso("tema de documento (Retrospect) y conjunto de estilos"):
        temas = os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), r"Microsoft Office\root\Document Themes 16")
        for nombre in ["Retrospect.thmx", "Facet.thmx", "Ion.thmx"]:
            if os.path.exists(os.path.join(temas, nombre)): d.ApplyDocumentTheme(os.path.join(temas, nombre)); break
    with b.paso("portada de la galería de Word"):
        w.Templates.LoadBuildingBlocks()
        bb = None
        for t in w.Templates:
            if "Built-In Building Blocks" in t.Name:
                cat = t.BuildingBlockTypes(2).Categories(1); bb = cat.BuildingBlocks(1); break
        bb.Insert(d.Paragraphs(1).Range, True)
        D.primero = False
    with b.paso("tabla de contenido"):
        par = D.p(""); D.fin(par).InsertBreak(7); D.primero = False
        D.p("Contenido", "TOC Heading"); par = D.p(""); d.TablesOfContents.Add(D.fin(par), True, 1, 3)
        par = D.p(""); D.fin(par).InsertBreak(7); D.primero = False
    lt = w.ListGalleries(3).ListTemplates(7) if w.ListGalleries(3).ListTemplates.Count >= 7 else w.ListGalleries(3).ListTemplates(1)
    capitulos = ["Resumen ejecutivo", "Introducción", "Objetivos", "Metodología", "Resultados", "Discusión", "Conclusiones"]
    for k, cap in enumerate(capitulos):
        with b.paso("capítulo " + cap):
            h = D.p(cap, -2); h.Range.ListFormat.ApplyListTemplateWithLevel(lt, k > 0, 0, 2)
            for j in range(2):
                sub = D.p(f"{cap}: apartado {j + 1}", -3); sub.Range.ListFormat.ApplyListTemplateWithLevel(lt, True, 0, 2)
                D.p(texto(k + j, 4), None, alin=3)
                if (k + j) % 3 == 0:
                    par = D.p(texto(k + 5, 2), None, alin=3); d.Footnotes.Add(Range=D.fin(par), Text="Nota del capítulo " + cap.lower() + ".")
            if k == 3:
                pi = D.p(""); im = d.InlineShapes.AddPicture(os.path.join(IMG, "barras.png"), False, True, D.fin(pi)); im.Width = cm(12); im.Height = cm(7.5)
                pi.Format.Alignment = 1; w.CaptionLabels.Add("Figura"); im.Range.InsertCaption("Figura", ": Avance por etapa", "", 1); D.primero = False
            if k == 4:
                pt = D.p(""); t = d.Tables.Add(pt.Range, 6, 4, 1, 0); t.Style = "Grid Table 4 - Accent 1"
                for i, fila in enumerate([["Etapa", "Planeado", "Real", "Diferencia"], ["Análisis", "10", "12", "+2"], ["Diseño", "15", "14", "-1"], ["Construcción", "40", "45", "+5"], ["Pruebas", "20", "18", "-2"], ["Total", "85", "89", "+4"]]):
                    for j, v in enumerate(fila): t.Cell(i + 1, j + 1).Range.Text = v
                t.ApplyStyleLastRow = True; D.primero = False
                pg = D.p(""); g = d.InlineShapes.AddChart2(-1, 51, D.fin(pg)); g.Width = cm(14); g.Height = cm(7); pg.Format.Alignment = 1
                datos_grafica(g.Chart, 51)
                ecuacion(D, D.p(""), "σ=√(1/N ∑_(i=1)^N▒(x_i-μ)^2 )")
            if k == 5:
                lista_pars = [D.p(t) for t in ["Las personas prefieren herramientas simples.", "Sin red, la colaboración sigue siendo posible.", "Los formatos de Office deben respetarse al detalle."]]
                rg = d.Range(lista_pars[0].Range.Start, lista_pars[-1].Range.End); rg.ListFormat.ApplyListTemplateWithLevel(w.ListGalleries(1).ListTemplates(1), False, 0, 2)
    with b.paso("anexo en hoja horizontal"):
        r = d.Content; r.Collapse(0); r.InsertBreak(2); D.primero = True
        s = d.Sections(d.Sections.Count); s.PageSetup.Orientation = 1
        D.p("Anexo A. Datos completos", -2)
        pt = D.p(""); t = d.Tables.Add(pt.Range, 12, 8, 1, 0); t.Style = "List Table 3 - Accent 5"
        for i in range(12):
            for j in range(8): t.Cell(i + 1, j + 1).Range.Text = ("Col " + str(j + 1)) if i == 0 else f"{(i * 13 + j * 7) % 97}.{j}"
        D.primero = False
    with b.paso("encabezado y pie con número de página"):
        for s in d.Sections:
            s.Headers(1).Range.Text = "Informe final del proyecto"; s.Headers(1).Range.ParagraphFormat.Alignment = 2
            f = s.Footers(1).Range; f.Text = ""; f.Fields.Add(f, 33); s.Footers(1).Range.ParagraphFormat.Alignment = 1
    return D.guardar()


# ---------------------------------------------------------------- w10: combinaciones raras
def w10(w):
    D = Doc(w, "w10-raros"); b = D.b; d = D.d; cm = D.cm
    D.p("Combinaciones raras", -63)
    with b.paso("palabra larguísima y dirección web sin espacios"):
        D.p("Supercalifragilisticoespialidosoextraordinariamentelarguísimaquenocabeenunrenglónparaningunaraz " + "https://www.ejemplo.mx/una/ruta/muy/larga/sin/espacios/para/ver/como/se/parte/la/linea/" * 2)
    with b.paso("espacios y guiones especiales"):
        D.p("No\u00a0se\u00a0parte\u00a0aquí\u00a0por\u00a0los\u00a0espacios\u00a0duros " + texto(0) + " guion\u2011duro y guion\u00adsuave en palabras\u00adlar\u00adguí\u00adsimas y espacio\u200bde\u200bancho\u200bcero.")
    with b.paso("guiones automáticos en texto justificado"):
        d.AutoHyphenation = True; d.HyphenationZone = 18
        par = D.p("Separación automática de palabras: " + texto(1, 5), None, alin=3); par.Format.RightIndent = cm(8)
    with b.paso("derecha a izquierda mezclada con izquierda a derecha"):
        D.p("Texto mezclado: hola مرحبا بكم في NOM-COLAB 2026 y שלום עולם 123 fin.")
    with b.paso("letra de 1 pt y de 120 pt"):
        par = D.p("Letra diminuta de 1 punto: "); D.r(par, "casi invisible " * 5, tam=1)
        D.p("Á", None, tam=120)
    with b.paso("texto blanco sobre fondo negro y resaltado sobre sombreado"):
        par = D.p(""); D.r(par, "Blanco sobre negro ", color=rgb(255, 255, 255), fondo=rgb(0, 0, 0)); D.r(par, " resaltado amarillo sobre sombreado azul ", resalte=7, fondo=rgb(189, 215, 238))
    with b.paso("interlineado exacto menor que la letra"):
        par = D.p("Letra de 20 pt con interlineado exacto de 10 pt: las líneas se enciman. " + texto(2), None, tam=20); par.Format.LineSpacingRule = 4; par.Format.LineSpacing = 10
    with b.paso("sangría francesa mayor que la izquierda"):
        par = D.p("Francesa de 3 cm con izquierda de 1 cm (el primer renglón sale al margen). " + texto(3)); par.Format.LeftIndent = cm(1); par.Format.FirstLineIndent = cm(-3)
    with b.paso("tabuladores más allá del margen derecho"):
        par = D.p("Tab\tmás allá\tdel margen"); par.TabStops.ClearAll(); par.TabStops.Add(cm(17), 0, 1); par.TabStops.Add(cm(19), 0, 2)
    with b.paso("letra capital seguida de lista"):
        par = D.p("Capital antes de una lista. " + texto(4, 2)); par.DropCap.Position = 1; par.DropCap.LinesToDrop = 2
        ps = [D.p(t) for t in ["uno", "dos", "tres"]]; rg = d.Range(ps[0].Range.Start, ps[-1].Range.End); rg.ListFormat.ApplyListTemplateWithLevel(w.ListGalleries(2).ListTemplates(1), False, 0, 2)
    with b.paso("lista dentro de tabla dentro de cuadro de texto"):
        a = D.p("Cuadro con una tabla y una lista adentro:")
        tb = d.Shapes.AddTextbox(1, cm(1), cm(0.5), cm(12), cm(5), a.Range); tb.RelativeHorizontalPosition = 0; tb.RelativeVerticalPosition = 2; tb.WrapFormat.Type = 4
        tr = tb.TextFrame.TextRange; tr.Text = "Antes de la tabla"
        tr2 = tb.TextFrame.TextRange; tr2.InsertParagraphAfter()
        rt = tb.TextFrame.TextRange.Paragraphs(2).Range
        t = rt.Tables.Add(rt, 2, 2); t.Borders.Enable = True; t.Cell(1, 1).Range.Text = "a\rb\rc"
        t.Cell(1, 1).Range.ListFormat.ApplyListTemplateWithLevel(w.ListGalleries(1).ListTemplates(1), False, 0, 2)
        t.Cell(1, 2).Range.Text = "celda"; t.Cell(2, 1).Range.Text = "abajo"; t.Cell(2, 2).Range.Text = "fin"
        for k in range(7): D.p("")
    with b.paso("forma flotante anclada dentro de una celda"):
        pt = D.p(""); t = d.Tables.Add(pt.Range, 2, 2, 1, 0); t.Borders.Enable = True; t.Rows.Height = cm(2.5); D.primero = False
        sh = d.Shapes.AddShape(92, cm(0.2), cm(0.2), cm(1.5), cm(1.5), t.Cell(1, 2).Range); sh.WrapFormat.Type = 3; sh.LayoutInCell = True
    with b.paso("nota al pie con tabla adentro"):
        par = D.p("Oración con una nota al pie que trae una tabla."); fn = d.Footnotes.Add(Range=D.fin(par), Text="Nota con tabla: ")
        r = fn.Range; r.InsertParagraphAfter(); r2 = fn.Range.Paragraphs.Last.Range; tt = fn.Range.Tables.Add(r2, 2, 2); tt.Borders.Enable = True
        tt.Cell(1, 1).Range.Text = "n1"; tt.Cell(2, 2).Range.Text = "n4"
    with b.paso("sección continua con columnas a media página y otra vez una columna"):
        r = d.Content; r.Collapse(0); r.InsertBreak(3); D.primero = True
        s = d.Sections(d.Sections.Count); s.PageSetup.TextColumns.SetCount(2)
        for k in range(3): D.p(texto(k + 5, 2))
        r = d.Content; r.Collapse(0); r.InsertBreak(3); D.primero = True
        s = d.Sections(d.Sections.Count); s.PageSetup.TextColumns.SetCount(1)
        D.p("De vuelta a una columna en la misma página.")
    with b.paso("sección A6 diminuta y luego A3 horizontal"):
        for (wd, hg) in [(297.6, 419.5), (1190.5, 841.9)]:
            r = d.Content; r.Collapse(0); r.InsertBreak(2); D.primero = True
            s = d.Sections(d.Sections.Count); s.PageSetup.TextColumns.EvenlySpaced = True
            s.PageSetup.LeftMargin = 36; s.PageSetup.RightMargin = 36; s.PageSetup.TopMargin = 36; s.PageSetup.BottomMargin = 36
            s.PageSetup.PageWidth = wd; s.PageSetup.PageHeight = hg
            D.p(f"Hoja de {wd:.0f} × {hg:.0f} pt. " + texto(k, 3))
    with b.paso("párrafos vacíos con letra enorme"):
        for k in range(3): D.p("", None, tam=48)
        D.p("Después de tres párrafos vacíos de 48 pt.")
    with b.paso("caracteres de control y símbolos de Wingdings"):
        par = D.p("Wingdings: "); D.r(par, "JKLMNOP", fuente="Wingdings", tam=16); D.r(par, "  Symbol: "); D.r(par, "abgdpqW", fuente="Symbol", tam=14)
    with b.paso("texto con todo junto"):
        par = D.p("")
        D.r(par, "Negrita cursiva subrayada doble roja tachada en versalitas con resaltado ", b=True, i=True, u=3, color=rgb(192, 0, 0), tachado=True, versalitas=True, resalte=7)
        D.r(par, "y superíndice", sup=True); D.r(par, " junto a subíndice", sub=True, fuente="Cambria", tam=16)
    return D.guardar()


if __name__ == "__main__":
    pythoncom.CoInitialize()
    hacer = {k: v for k, v in globals().items() if k[:1] == "w" and k[1:].isdigit()}
    pedidos = sys.argv[1:] or sorted(hacer)
    w = word()
    try:
        for k in pedidos:
            print("==", k, flush=True)
            try: print(" ->", hacer[k](w), flush=True)
            except Exception: traceback.print_exc()
    finally:
        w.Quit(0)

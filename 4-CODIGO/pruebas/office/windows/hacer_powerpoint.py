"""Arma con PowerPoint de verdad presentaciones que usan casi todo lo que PowerPoint sabe hacer.

  python hacer_powerpoint.py [p01 p02 ...]      (sin nombres: todas)

PowerPoint es una sola copia en el equipo: nunca se cierra con Quit (puede ser la de la persona que lo
usa); solo se cierran las presentaciones que se abren aquí, y sin ventana.
Deja `archivos\\pNN-*.pptx` y `archivos\\pNN-*.funciones.json`.
"""
import os, sys, time
from comun import *
from win32com.client import constants as K

TEMAS_DIR = os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), r"Microsoft Office\root\Document Themes 16")


class Pres:
    def __init__(self, p, nombre, ancho=960, alto=540):
        self.p, self.nombre = p, nombre
        # con ventana (fuera de la pantalla): sin ventana, PowerPoint no deja crear gráficas ni SmartArt
        self.pr = p.Presentations.Add(WithWindow=True)
        try: p.WindowState = 1; p.Left = -16000; p.Top = 0
        except Exception: pass
        self.pr.PageSetup.SlideWidth = ancho; self.pr.PageSetup.SlideHeight = alto
        self.b = Bitacora(nombre)

    def diseno(self, nombre):
        m = self.pr.SlideMaster
        for k in range(1, m.CustomLayouts.Count + 1):
            if m.CustomLayouts(k).Name == nombre: return m.CustomLayouts(k)
        return m.CustomLayouts(min(7, m.CustomLayouts.Count))

    def diapo(self, diseno="Blank", titulo=None):
        sl = self.pr.Slides.AddSlide(self.pr.Slides.Count + 1, self.diseno(diseno))
        if titulo is not None:
            try: sl.Shapes.Title.TextFrame.TextRange.Text = titulo
            except Exception: pass
        return sl

    def guardar(self):
        ruta = os.path.join(ARCHIVOS, self.nombre + ".pptx")
        self.pr.SaveAs(ruta, 24)
        guardar_json(os.path.join(ARCHIVOS, self.nombre + ".funciones.json"), self.b.resumen())
        self.pr.Close()
        return ruta


def caja(sl, x, y, w, h, t, tam=14, **f):
    s = sl.Shapes.AddTextbox(1, x, y, w, h)
    tr = s.TextFrame2.TextRange; tr.Text = t; tr.Font.Size = tam
    for k, v in f.items(): setattr(tr.Font, k, v)
    return s


# ---------------------------------------------------------------- p01: texto
def p01(p):
    P = Pres(p, "p01-texto"); b = P.b
    with b.paso("portada con título y subtítulo"):
        sl = P.diapo("Title Slide", "Texto en PowerPoint"); sl.Shapes(2).TextFrame.TextRange.Text = "Formato de letra, párrafos, viñetas y cuadros"
    with b.paso("viñetas de cinco niveles"):
        sl = P.diapo("Title and Content", "Viñetas por nivel")
        tr = sl.Shapes(2).TextFrame2.TextRange
        tr.Text = "Primer nivel\rSegundo nivel\rTercer nivel\rCuarto nivel\rQuinto nivel\rDe vuelta al primero"
        for k, n in enumerate([1, 2, 3, 4, 5, 1]): tr.GetParagraphs(k + 1, 1).ParagraphFormat.IndentLevel = n
    with b.paso("numeraciones de todos los estilos"):
        sl = P.diapo("Title and Content", "Numeraciones")
        tr = sl.Shapes(2).TextFrame2.TextRange
        estilos = [(1, "1."), (2, "1)"), (3, "(1)"), (4, "a."), (5, "a)"), (6, "A."), (7, "i."), (8, "I.")]
        tr.Text = "\r".join(f"Estilo {e}" for _, e in estilos)
        for k, (st, _) in enumerate(estilos):
            pf = tr.GetParagraphs(k + 1, 1).ParagraphFormat; pf.Bullet.Type = 2; pf.Bullet.Style = [1, 2, 3, 5, 4, 6, 8, 7][k]
    with b.paso("viñetas de símbolo, color y tamaño"):
        sl = P.diapo("Title and Content", "Viñetas propias")
        tr = sl.Shapes(2).TextFrame2.TextRange; tr.Text = "Flecha de Wingdings roja\rPalomita verde al 150 %\rEstrella azul\rSin viñeta"
        for k, (fu, ch, c, tm) in enumerate([("Wingdings", 0xF0E8, rgb(192, 0, 0), 1.0), ("Wingdings", 0xF0FC, rgb(0, 176, 80), 1.5), ("Segoe UI Symbol", 0x2605, rgb(0, 112, 192), 1.0)]):
            bu = tr.GetParagraphs(k + 1, 1).ParagraphFormat.Bullet; bu.Type = 1; bu.Font.Name = fu; bu.Character = ch; bu.Font.Fill.ForeColor.RGB = c; bu.RelativeSize = tm
        tr.GetParagraphs(4, 1).ParagraphFormat.Bullet.Visible = False
    with b.paso("fuentes"):
        sl = P.diapo("Title Only", "Fuentes")
        fuentes = ["Aptos", "Aptos Display", "Calibri", "Cambria", "Times New Roman", "Arial", "Georgia", "Verdana", "Segoe UI", "Consolas", "Century Gothic", "Garamond", "Comic Sans MS", "Impact", "Trebuchet MS", "Franklin Gothic Medium"]
        for k, f in enumerate(fuentes):
            caja(sl, 30 + (k // 8) * 460, 110 + (k % 8) * 50, 450, 40, f + ": Ágil murciélago 0123", 18, Name=f)
    with b.paso("tamaños"):
        sl = P.diapo("Blank")
        y = 10
        for t in [8, 10, 12, 14, 18, 24, 32, 44, 60, 80]:
            caja(sl, 20, y, 920, t * 1.5, f"{t} pt Ágil", t); y += t * 1.25 + 4
    with b.paso("estilos de letra"):
        sl = P.diapo("Title Only", "Estilos de letra")
        s = caja(sl, 40, 120, 880, 380, "", 22); tr = s.TextFrame2.TextRange
        partes = [("Negrita ", {"Bold": True}), ("cursiva ", {"Italic": True}), ("subrayado ", {"UnderlineStyle": 1}), ("doble ", {"UnderlineStyle": 3}),
                  ("ondulado ", {"UnderlineStyle": 11}), ("tachado ", {"Strike": 1}), ("doble tachado ", {"Strike": 2}), ("superíndice ", {"Superscript": True}),
                  ("subíndice ", {"Subscript": True}), ("Versalitas ", {"Smallcaps": True}), ("mayúsculas ", {"Allcaps": True}), ("espaciado ", {"Spacing": 6}),
                  ("comprimido ", {"Spacing": -1.5}), ("rojo ", {"Fill.ForeColor.RGB": rgb(192, 0, 0)}), ("resaltado", {"Highlight.RGB": rgb(255, 255, 0)})]
        for t, f in partes:
            r = tr.InsertAfter(t)
            for k, v in f.items():
                obj = r.Font
                *camino, ult = k.split(".")
                for c in camino: obj = getattr(obj, c)
                setattr(obj, ult, v)
    with b.paso("alineación, interlineado y sangrías"):
        sl = P.diapo("Blank")
        for k, (al, nombre) in enumerate([(1, "izquierda"), (2, "centro"), (3, "derecha"), (4, "justificado"), (5, "distribuido")]):
            s = caja(sl, 20 + (k % 3) * 310, 20 + (k // 3) * 250, 300, 230, f"Alineación {nombre}. " + texto(k, 2), 12)
            s.TextFrame2.TextRange.ParagraphFormat.Alignment = al; s.Line.Visible = True
        s = caja(sl, 640, 280, 300, 240, "Interlineado 1.5, espacio antes 12 pt y sangría francesa. " + texto(3, 2), 12)
        pf = s.TextFrame2.TextRange.ParagraphFormat; pf.SpaceWithin = 1.5; pf.SpaceBefore = 12; pf.FirstLineIndent = -20; pf.LeftIndent = 20; s.Line.Visible = True
    with b.paso("ajuste automático: reducir el texto y ajustar la forma"):
        sl = P.diapo("Blank")
        s1 = caja(sl, 20, 20, 440, 160, texto(0, 6), 20); s1.TextFrame2.AutoSize = 2; s1.Line.Visible = True
        s2 = caja(sl, 500, 20, 440, 100, texto(1, 2), 18); s2.TextFrame2.AutoSize = 1; s2.Line.Visible = True
        s3 = caja(sl, 20, 260, 440, 120, texto(2, 6), 18); s3.TextFrame2.AutoSize = 0; s3.Line.Visible = True; s3.Height = 120
        s4 = caja(sl, 500, 260, 440, 60, "Sin ajuste de línea: este texto sigue derecho y se sale de la caja", 18); s4.TextFrame2.WordWrap = False; s4.Line.Visible = True
    with b.paso("columnas, márgenes, anclaje y texto vertical"):
        sl = P.diapo("Blank")
        s = caja(sl, 20, 20, 600, 240, texto(0, 6), 12); s.TextFrame2.Column.Number = 2; s.TextFrame2.Column.Spacing = 20; s.Line.Visible = True
        for k, (anc, n) in enumerate([(1, "arriba"), (3, "en medio"), (4, "abajo")]):
            r = sl.Shapes.AddShape(1, 640, 20 + k * 85, 300, 75); r.TextFrame2.TextRange.Text = "Anclado " + n; r.TextFrame2.VerticalAnchor = anc
            r.TextFrame2.MarginLeft = 20; r.TextFrame2.MarginTop = 4
        for k, (o, n) in enumerate([(2, "hacia arriba"), (3, "hacia abajo"), (5, "apilado"), (6, "vertical asiático")]):
            v = sl.Shapes.AddShape(1, 20 + k * 120, 300, 100, 220); v.TextFrame2.TextRange.Text = "Texto " + n; v.TextFrame2.Orientation = o
    with b.paso("WordArt y efectos de texto"):
        sl = P.diapo("Blank")
        for k in range(6):
            w = sl.Shapes.AddTextEffect(k * 5, "Efecto " + str(k + 1), "Arial Black", 36, True, False, 20 + (k % 2) * 470, 20 + (k // 2) * 170)
        s = caja(sl, 20, 470, 920, 60, "Resplandor, sombra y reflejo", 36, Bold=True)
        fo = s.TextFrame2.TextRange.Font; fo.Glow.Radius = 8; fo.Glow.Color.RGB = rgb(255, 192, 0); fo.Shadow.Visible = True; fo.Reflection.Type = 2
    with b.paso("texto deformado (arco, onda)"):
        sl = P.diapo("Blank")
        for k, wf in enumerate([5, 6, 15, 26]):
            s = caja(sl, 40 + (k % 2) * 460, 40 + (k // 2) * 250, 420, 200, "Texto deformado", 40, Bold=True); s.TextFrame2.WarpFormat = wf
    with b.paso("vínculos y texto de derecha a izquierda"):
        sl = P.diapo("Blank")
        s = caja(sl, 40, 40, 880, 50, "Sitio del IPN", 24); s.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.Address = "https://www.ipn.mx"
        s = caja(sl, 40, 120, 880, 50, "مرحبا بكم في العرض التقديمي", 28); s.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2
        caja(sl, 40, 200, 880, 50, "中文 日本語 한국어 Ελληνικά Кириллица 😀👍🏽", 28)
    return P.guardar()


# ---------------------------------------------------------------- p02: todas las formas predefinidas
def p02(p):
    P = Pres(p, "p02-formas"); b = P.b
    tipos = list(range(1, 184))
    por = 24
    for i0 in range(0, len(tipos), por):
        sl = P.diapo("Blank")
        for k, t in enumerate(tipos[i0:i0 + por]):
            with b.paso(f"forma {t}"):
                x = 15 + (k % 6) * 158; y = 12 + (k // 6) * 132
                s = sl.Shapes.AddShape(t, x + 20, y, 110, 95)
                s.Fill.ForeColor.RGB = [rgb(68, 114, 196), rgb(237, 125, 49), rgb(112, 173, 71), rgb(255, 192, 0), rgb(91, 155, 213), rgb(165, 165, 165)][k % 6]
                lb = caja(sl, x, y + 96, 150, 22, f"{t}: {s.AutoShapeType}", 9); lb.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    with b.paso("ajustes (manijas amarillas) de varias formas"):
        sl = P.diapo("Blank")
        for k, (t, aj) in enumerate([(5, [0.4]), (33, [0.7, 0.2]), (92, [0.2]), (106, [-0.6, 0.9]), (24, [0.3]), (7, [0.1])]):
            s = sl.Shapes.AddShape(t, 30 + (k % 3) * 310, 40 + (k // 3) * 250, 260, 200)
            for j, v in enumerate(aj): s.Adjustments.SetItem(j + 1, v)
    with b.paso("girar, voltear y agrupar"):
        sl = P.diapo("Blank")
        a = sl.Shapes.AddShape(33, 50, 50, 200, 100); a.Rotation = 30
        c = sl.Shapes.AddShape(33, 300, 50, 200, 100); c.Flip(0)
        d = sl.Shapes.AddShape(33, 550, 50, 200, 100); d.Flip(1)
        g1 = sl.Shapes.AddShape(9, 100, 250, 120, 120); g2 = sl.Shapes.AddShape(1, 180, 300, 160, 80); g3 = sl.Shapes.AddShape(92, 300, 240, 100, 100)
        g = sl.Shapes.Range([g1.Name, g2.Name, g3.Name]).Group(); g.Rotation = 20
        h1 = sl.Shapes.AddShape(5, 550, 250, 120, 80); h2 = sl.Shapes.AddShape(9, 680, 280, 100, 100)
        g2_ = sl.Shapes.Range([h1.Name, h2.Name]).Group()
        h3 = sl.Shapes.AddShape(7, 800, 250, 100, 100)
        gg = sl.Shapes.Range([g2_.Name, h3.Name]).Group(); gg.Flip(0); gg.Rotation = -15
    with b.paso("conectores rectos, en codo y curvos"):
        sl = P.diapo("Blank")
        a = sl.Shapes.AddShape(1, 60, 60, 160, 80); a.TextFrame2.TextRange.Text = "Inicio"
        c = sl.Shapes.AddShape(9, 700, 380, 160, 100); c.TextFrame2.TextRange.Text = "Fin"
        m = sl.Shapes.AddShape(4, 380, 220, 180, 100); m.TextFrame2.TextRange.Text = "Decisión"
        for tipo, (s1, p1, s2, p2) in [(1, (a, 4, m, 2)), (2, (m, 3, c, 1)), (3, (a, 3, c, 2))]:
            cn = sl.Shapes.AddConnector(tipo, 0, 0, 10, 10); cn.ConnectorFormat.BeginConnect(s1, p1); cn.ConnectorFormat.EndConnect(s2, p2)
            cn.Line.EndArrowheadStyle = 2; cn.Line.Weight = 2
    with b.paso("forma libre y curva"):
        sl = P.diapo("Blank")
        ff = sl.Shapes.BuildFreeform(0, 100, 100); ff.AddNodes(0, 0, 300, 80); ff.AddNodes(1, 0, 350, 250, 250, 300, 150, 260); ff.AddNodes(0, 0, 100, 100)
        s = ff.ConvertToShape(); s.Fill.ForeColor.RGB = rgb(112, 173, 71)
        pts = W.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R4, [[500, 400], [550, 100], [700, 100], [800, 400]])
        cv = sl.Shapes.AddCurve(pts); cv.Line.Weight = 4; cv.Line.ForeColor.RGB = rgb(192, 0, 0)
    with b.paso("combinar formas (unión, resta, intersección)"):
        sl = P.diapo("Blank")
        for k, op in enumerate([1, 2, 3, 4, 5]):
            a = sl.Shapes.AddShape(9, 30 + k * 185, 150, 120, 120); c = sl.Shapes.AddShape(1, 90 + k * 185, 190, 90, 90)
            sl.Shapes.Range([a.Name, c.Name]).MergeShapes(op, a)
    return P.guardar()


# ---------------------------------------------------------------- p03: rellenos, líneas y efectos
def p03(p):
    P = Pres(p, "p03-rellenos-efectos"); b = P.b
    with b.paso("degradados lineal, radial, rectangular y de ruta"):
        sl = P.diapo("Blank")
        for k, (estilo, var) in enumerate([(1, 1), (2, 1), (3, 1), (4, 2), (5, 1), (7, 1), (7, 2)]):
            s = sl.Shapes.AddShape(1, 20 + (k % 4) * 235, 20 + (k // 4) * 260, 215, 230)
            s.Fill.ForeColor.RGB = rgb(0, 112, 192); s.Fill.BackColor.RGB = rgb(255, 255, 255)
            try: s.Fill.TwoColorGradient(estilo, var)
            except Exception: s.Fill.TwoColorGradient(1, 1)
        s = sl.Shapes.AddShape(9, 725, 280, 215, 230); f = s.Fill; f.PresetGradient(1, 1, 10)
    with b.paso("degradado de varias paradas con transparencia"):
        sl = P.diapo("Blank")
        s = sl.Shapes.AddShape(1, 40, 40, 880, 200); f = s.Fill; f.TwoColorGradient(1, 1); f.ForeColor.RGB = rgb(255, 0, 0); f.BackColor.RGB = rgb(0, 0, 255)
        gs = f.GradientStops; gs.Insert(rgb(255, 255, 0), 0.33); gs.Insert(rgb(0, 255, 0), 0.66, 0.5)
        f.GradientAngle = 30
    with b.paso("tramas (48 patrones)"):
        sl = P.diapo("Blank")
        for k in range(1, 49):
            s = sl.Shapes.AddShape(1, 10 + ((k - 1) % 8) * 118, 10 + ((k - 1) // 8) * 88, 110, 80)
            try: s.Fill.Patterned(k); s.Fill.ForeColor.RGB = rgb(31, 78, 121); s.Fill.BackColor.RGB = rgb(222, 235, 247)
            except Exception: s.Fill.ForeColor.RGB = rgb(200, 0, 0)
    with b.paso("relleno con imagen (estirada y en mosaico) y textura"):
        sl = P.diapo("Blank")
        s = sl.Shapes.AddShape(9, 30, 40, 280, 280); s.Fill.UserPicture(os.path.join(IMG, "foto1.jpg"))
        s = sl.Shapes.AddShape(1, 340, 40, 280, 280); s.Fill.UserTextured(os.path.join(IMG, "mosaico.png"))
        s = sl.Shapes.AddShape(5, 650, 40, 280, 280); s.Fill.PresetTextured(5)
        s = sl.Shapes.AddShape(1, 30, 350, 900, 150); s.Fill.ForeColor.RGB = rgb(0, 176, 80); s.Fill.Transparency = 0.5
        t = caja(sl, 40, 400, 880, 60, "Encima de una forma verde al 50 % de transparencia", 28)
    with b.paso("líneas: guiones, grosores, compuestas, puntas y extremos"):
        sl = P.diapo("Blank")
        for k, (dash, w, st, ini, fin) in enumerate([(1, 1, 1, 1, 1), (2, 2, 1, 2, 2), (3, 3, 1, 3, 3), (4, 4, 1, 4, 4), (5, 5, 1, 5, 5), (6, 6, 1, 6, 6),
                                                      (7, 7, 2, 2, 3), (8, 8, 3, 3, 4), (1, 10, 4, 4, 5), (10, 3, 5, 6, 2), (11, 4, 1, 2, 6), (12, 6, 1, 1, 2)]):
            ln = sl.Shapes.AddLine(60, 30 + k * 42, 900, 30 + k * 42); l = ln.Line
            l.DashStyle = dash; l.Weight = w; l.Style = st; l.BeginArrowheadStyle = ini; l.EndArrowheadStyle = fin; l.ForeColor.RGB = rgb(40, 40, 40)
    with b.paso("contornos de forma"):
        sl = P.diapo("Blank")
        for k in range(8):
            s = sl.Shapes.AddShape(5, 20 + (k % 4) * 235, 30 + (k // 4) * 250, 200, 200); s.Fill.ForeColor.RGB = rgb(255, 255, 255)
            s.Line.Weight = 1 + k * 2; s.Line.DashStyle = 1 + k % 8; s.Line.ForeColor.RGB = [rgb(192, 0, 0), rgb(0, 112, 192), rgb(0, 176, 80), rgb(112, 48, 160)][k % 4]
    with b.paso("sombras exterior, interior y en perspectiva"):
        sl = P.diapo("Blank")
        for k, (tipo) in enumerate([1, 5, 9, 13, 17, 21, 25, 30]):
            s = sl.Shapes.AddShape(1, 30 + (k % 4) * 235, 40 + (k // 4) * 250, 180, 160); s.Fill.ForeColor.RGB = rgb(91, 155, 213)
            try: s.Shadow.Type = tipo
            except Exception: s.Shadow.Visible = True
    with b.paso("resplandor, bordes suaves y reflejo"):
        sl = P.diapo("Blank")
        for k in range(4):
            s = sl.Shapes.AddShape(9, 40 + k * 230, 40, 180, 180); s.Glow.Radius = 5 + k * 6; s.Glow.Color.RGB = [rgb(255, 192, 0), rgb(0, 176, 240), rgb(146, 208, 80), rgb(255, 0, 255)][k]
        for k in range(4):
            s = sl.Shapes.AddShape(1, 40 + k * 230, 280, 180, 120); s.SoftEdge.Type = k + 1
            s.Fill.ForeColor.RGB = rgb(237, 125, 49)
        for k in range(4):
            s = sl.Shapes.AddShape(5, 40 + k * 230, 420, 180, 60); s.Reflection.Type = k + 1
    with b.paso("biseles y giro 3D"):
        sl = P.diapo("Blank")
        for k, bt in enumerate([1, 3, 5, 7, 9, 11]):
            s = sl.Shapes.AddShape(1, 30 + (k % 3) * 310, 30 + (k // 3) * 250, 220, 180); s.ThreeD.BevelTopType = bt; s.ThreeD.BevelTopDepth = 8; s.ThreeD.BevelTopInset = 10
            s.Fill.ForeColor.RGB = rgb(68, 114, 196)
            if k >= 3: s.ThreeD.SetPresetCamera([1, 13, 25][k - 3]); s.ThreeD.Depth = 30
    with b.paso("estilos rápidos de forma"):
        sl = P.diapo("Blank")
        for k in range(42):
            s = sl.Shapes.AddShape(5, 10 + (k % 7) * 136, 10 + (k // 7) * 88, 125, 78)
            try: s.ShapeStyle = 10001 + k
            except Exception: pass
    return P.guardar()


# ---------------------------------------------------------------- p04: imágenes
def p04(p):
    P = Pres(p, "p04-imagenes"); b = P.b
    with b.paso("imagen normal, recortada y recortada a forma"):
        sl = P.diapo("Blank")
        sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 20, 20, 300, 188)
        c = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 340, 20, 300, 188)
        c.PictureFormat.CropLeft = 100; c.PictureFormat.CropRight = 60; c.PictureFormat.CropTop = 30
        f = sl.Shapes.AddPicture(os.path.join(IMG, "foto2.jpg"), False, True, 660, 20, 280, 280); f.AutoShapeType = 9
        h = sl.Shapes.AddPicture(os.path.join(IMG, "foto2.jpg"), False, True, 20, 260, 250, 250); h.AutoShapeType = 92
        r = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 300, 280, 330, 206); r.AutoShapeType = 5
    with b.paso("girada y volteada"):
        sl = P.diapo("Blank")
        a = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 40, 60, 360, 225); a.Rotation = -12
        c = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 520, 60, 360, 225); c.Flip(0)
        d = sl.Shapes.AddPicture(os.path.join(IMG, "foto3-vertical.jpg"), False, True, 380, 300, 140, 218); d.Flip(1)
    with b.paso("color: escala de grises, blanco y negro, desvaído; brillo y contraste"):
        sl = P.diapo("Blank")
        for k, (tipo, br, co) in enumerate([(2, 0.5, 0.5), (3, 0.5, 0.5), (4, 0.5, 0.5), (1, 0.75, 0.5), (1, 0.3, 0.8), (1, 0.5, 0.2)]):
            i = sl.Shapes.AddPicture(os.path.join(IMG, "foto2.jpg"), False, True, 20 + (k % 3) * 310, 20 + (k // 3) * 260, 240, 240)
            i.PictureFormat.ColorType = tipo; i.PictureFormat.Brightness = br; i.PictureFormat.Contrast = co
    with b.paso("efectos artísticos"):
        sl = P.diapo("Blank")
        for k, ef in enumerate([1, 4, 8, 13, 19, 22]):
            i = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 20 + (k % 3) * 310, 20 + (k // 3) * 260, 300, 188)
            try: i.Fill.PictureEffects.Insert(ef)
            except Exception: pass
    with b.paso("estilos de imagen: marco, sombra, reflejo, bordes suaves, bisel"):
        sl = P.diapo("Blank")
        a = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 30, 30, 280, 175); a.Line.Visible = True; a.Line.Weight = 10; a.Line.ForeColor.RGB = rgb(255, 255, 255); a.Shadow.Type = 21
        c = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 340, 30, 280, 175); c.Reflection.Type = 5
        d = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 650, 30, 280, 175); d.SoftEdge.Type = 4
        e = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 30, 300, 280, 175); e.ThreeD.BevelTopType = 3
        g = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 340, 300, 280, 175); g.ThreeD.SetPresetCamera(25); g.Shadow.Visible = True
        h = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 650, 300, 280, 175); h.Glow.Radius = 10; h.Glow.Color.RGB = rgb(0, 176, 240)
    with b.paso("transparencia de imagen y PNG con fondo transparente sobre color"):
        sl = P.diapo("Blank")
        fondo = sl.Shapes.AddShape(1, 0, 0, 960, 540); fondo.Fill.ForeColor.RGB = rgb(255, 230, 153); fondo.Line.Visible = False
        sl.Shapes.AddPicture(os.path.join(IMG, "logo.png"), False, True, 60, 100, 300, 300)
        t = sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 420, 100, 480, 300); t.Fill.Transparency = 0.5
        try: t.PictureFormat.TransparencyColor = rgb(255, 255, 255); t.PictureFormat.TransparentBackground = True
        except Exception: pass
    with b.paso("ícono SVG con color cambiado"):
        sl = P.diapo("Blank")
        for k in range(4):
            i = sl.Shapes.AddPicture(os.path.join(IMG, "icono.svg"), False, True, 60 + k * 220, 170, 180, 180)
            if k: i.Fill.ForeColor.RGB = [rgb(192, 0, 0), rgb(0, 112, 192), rgb(112, 48, 160)][k - 1]
    with b.paso("fondo de diapositiva con imagen"):
        sl = P.diapo("Title Only", "Fondo con imagen"); sl.FollowMasterBackground = False; sl.Background.Fill.UserPicture(os.path.join(IMG, "foto1.jpg"))
    return P.guardar()


# ---------------------------------------------------------------- p05: tablas y gráficas
ESTILOS_TABLA = ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}", "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
                 "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}", "{5940675A-B579-460E-94D1-54222C63F5DA}", "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}"]


def p05(p):
    P = Pres(p, "p05-tablas-graficas"); b = P.b
    for k, est in enumerate(ESTILOS_TABLA):
        with b.paso(f"tabla con estilo {k + 1}"):
            if k % 2 == 0: sl = P.diapo("Blank")
            t = sl.Shapes.AddTable(5, 4, 30 + (k % 2) * 460, 40, 420, 220).Table
            t.ApplyStyle(est, False)
            for i in range(5):
                for j in range(4): t.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Text = ["Materia", "P1", "P2", "Final"][j] if i == 0 else f"{(i * 7 + j * 3) % 10}.{j}"
            t.FirstRow = True; t.HorizBanding = True; t.FirstCol = k % 3 == 0; t.LastRow = k == 4
    with b.paso("tabla con celdas combinadas, rellenos, bordes y diagonal"):
        sl = P.diapo("Blank")
        t = sl.Shapes.AddTable(5, 5, 40, 40, 880, 400).Table
        for i in range(5):
            for j in range(5): t.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Text = f"{i + 1},{j + 1}"
        t.Cell(1, 1).Merge(t.Cell(1, 3)); t.Cell(2, 1).Merge(t.Cell(4, 1)); t.Cell(3, 3).Merge(t.Cell(4, 4))
        t.Cell(2, 2).Shape.Fill.ForeColor.RGB = rgb(255, 192, 0); t.Cell(5, 5).Shape.Fill.ForeColor.RGB = rgb(0, 176, 80)
        bd = t.Cell(5, 2).Borders(1); bd.Weight = 6; bd.ForeColor.RGB = rgb(192, 0, 0)
        dg = t.Cell(2, 5).Borders(7); dg.Weight = 2; dg.ForeColor.RGB = rgb(0, 0, 0); dg.Visible = True
        t.Cell(2, 3).Shape.TextFrame.TextRange.Text = "Texto largo que se ajusta dentro de la celda de la tabla"
        t.Cell(2, 4).Shape.TextFrame.VerticalAnchor = 4; t.Cell(2, 4).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 3
        t.Columns(1).Width = 120; t.Rows(5).Height = 80
    tipos = [("Columnas", 51), ("Barras", 57), ("Líneas", 65), ("Circular", 5), ("Anillo", -4120), ("Áreas", 76), ("Dispersión", -4169), ("Radial", 82),
             ("Cascada", 119), ("Rectángulos", 117), ("Proyección solar", 120), ("Histograma", 118), ("Caja y bigotes", 121), ("Embudo", 123), ("Columnas 3D", -4100)]
    for k, (nombre, tipo) in enumerate(tipos):
        with b.paso("gráfica " + nombre):
            if k % 2 == 0: sl = P.diapo("Blank")
            sh = sl.Shapes.AddChart2(-1, tipo, 20 + (k % 2) * 470, 30, 450, 480)
            for i in range(20):
                try: sh.Chart.ChartData.Workbook.Close(); break
                except Exception as e:
                    if getattr(e, "hresult", None) in (-2147418111, -2147417846): time.sleep(1); continue
                    break
            ch = sh.Chart; ch.HasTitle = True; ch.ChartTitle.Text = nombre
    with b.paso("gráfica con etiquetas, leyenda abajo y colores propios"):
        sl = P.diapo("Blank")
        sh = sl.Shapes.AddChart2(-1, 51, 40, 30, 880, 480)
        try: sh.Chart.ChartData.Workbook.Close()
        except Exception: pass
        ch = sh.Chart; ch.HasLegend = True; ch.Legend.Position = -4107
        for k in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(k); s.HasDataLabels = True; s.Format.Fill.ForeColor.RGB = [rgb(29, 122, 76), rgb(255, 192, 0), rgb(192, 0, 0)][(k - 1) % 3]
    return P.guardar()


# ---------------------------------------------------------------- p06: SmartArt
def p06(p):
    P = Pres(p, "p06-smartart"); b = P.b
    lays = p.SmartArtLayouts
    n = lays.Count
    elegidos = list(range(1, n + 1, max(1, n // 36)))[:36]
    for k, i in enumerate(elegidos):
        with b.paso(f"SmartArt {lays(i).Name}"):
            if k % 2 == 0: sl = P.diapo("Blank")
            sa = sl.Shapes.AddSmartArt(lays(i), 20 + (k % 2) * 470, 40, 450, 460)
            nodos = sa.SmartArt.AllNodes
            for j in range(1, min(nodos.Count, 6) + 1):
                try: nodos(j).TextFrame2.TextRange.Text = ["Planear", "Diseñar", "Construir", "Probar", "Entregar", "Mejorar"][j - 1]
                except Exception: pass
            cap = caja(sl, 20 + (k % 2) * 470, 505, 450, 30, lays(i).Name, 11)
            try:
                sa.SmartArt.Color = p.SmartArtColors((k % p.SmartArtColors.Count) + 1)
                sa.SmartArt.QuickStyle = p.SmartArtQuickStyles((k % p.SmartArtQuickStyles.Count) + 1)
            except Exception: pass
    return P.guardar()


# ---------------------------------------------------------------- p07: temas, diseños y fondos
def p07(p):
    P = Pres(p, "p07-temas"); b = P.b
    temas = sorted(f for f in os.listdir(TEMAS_DIR) if f.lower().endswith(".thmx"))[:14] if os.path.isdir(TEMAS_DIR) else []
    with b.paso("todos los diseños del patrón Office"):
        m = P.pr.SlideMaster
        for k in range(1, m.CustomLayouts.Count + 1):
            sl = P.pr.Slides.AddSlide(P.pr.Slides.Count + 1, m.CustomLayouts(k))
            for sh in sl.Shapes:
                if sh.HasTextFrame:
                    try: sh.TextFrame.TextRange.Text = f"{m.CustomLayouts(k).Name}: {sh.Name}"
                    except Exception: pass
    for t in temas:
        with b.paso("tema " + t):
            P.pr.Designs.Load(os.path.join(TEMAS_DIR, t))
            dsn = P.pr.Designs(P.pr.Designs.Count)
            for nombre in [1, 2]:
                lay = dsn.SlideMaster.CustomLayouts(nombre)
                sl = P.pr.Slides.AddSlide(P.pr.Slides.Count + 1, lay)
                try: sl.Shapes(1).TextFrame.TextRange.Text = "Tema " + t[:-5]
                except Exception: pass
                try: sl.Shapes(2).TextFrame.TextRange.Text = "Primer punto\rSegundo punto\rTercer punto"
                except Exception: pass
    with b.paso("fondos: sólido, degradado, trama, imagen en mosaico"):
        for k in range(4):
            sl = P.diapo("Title Only", ["Fondo sólido", "Fondo degradado", "Fondo con trama", "Fondo en mosaico"][k]); sl.FollowMasterBackground = False
            f = sl.Background.Fill
            if k == 0: f.Solid(); f.ForeColor.RGB = rgb(31, 56, 100)
            if k == 1: f.TwoColorGradient(1, 1); f.ForeColor.RGB = rgb(255, 255, 255); f.BackColor.RGB = rgb(91, 155, 213)
            if k == 2: f.Patterned(30); f.ForeColor.RGB = rgb(200, 200, 200); f.BackColor.RGB = rgb(255, 255, 255)
            if k == 3: f.UserTextured(os.path.join(IMG, "mosaico.png"))
    with b.paso("número de diapositiva, pie y fecha"):
        for sl in P.pr.Slides:
            hf = sl.HeadersFooters
            try: hf.SlideNumber.Visible = True; hf.Footer.Visible = True; hf.Footer.Text = "NOM-COLAB · prueba de temas"; hf.DateAndTime.Visible = True; hf.DateAndTime.UseFormat = True
            except Exception: pass
    return P.guardar()


# ---------------------------------------------------------------- p08: transiciones, animaciones, notas y secciones
TRANSICIONES = ["ppEffectFadeSmoothly", "ppEffectPushLeft", "ppEffectWipeRight", "ppEffectSplitHorizontalOut", "ppEffectRevealSmoothLeft", "ppEffectRandomBarsHorizontal",
                "ppEffectCircleOut", "ppEffectUncoverLeft", "ppEffectCoverLeft", "ppEffectFallOverLeft", "ppEffectDrapeLeft", "ppEffectCurtains",
                "ppEffectWindLeft", "ppEffectPrestige", "ppEffectFracture", "ppEffectCrush", "ppEffectPeelOffLeft", "ppEffectPageCurlDoubleLeft", "ppEffectAirplaneLeft",
                "ppEffectOrigamiLeft", "ppEffectDissolve", "ppEffectCheckerboardAcross", "ppEffectBlindsHorizontal", "ppEffectWheel1Spoke", "ppEffectRippleCenter",
                "ppEffectHoneycomb", "ppEffectGlitterDiamondLeft", "ppEffectVortexLeft", "ppEffectShredStripsIn", "ppEffectSwitchLeft", "ppEffectFlipLeft",
                "ppEffectGalleryLeft", "ppEffectCubeLeft", "ppEffectDoorsHorizontal", "ppEffectBoxLeft", "ppEffectMorphByObject"]


def p08(p):
    P = Pres(p, "p08-animaciones"); b = P.b
    for k, t in enumerate(TRANSICIONES):
        with b.paso("transición " + t):
            sl = P.diapo("Title and Content", f"Transición {k + 1}: {t[8:]}")
            sl.Shapes(2).TextFrame.TextRange.Text = "Uno\rDos\rTres"
            st = sl.SlideShowTransition; st.EntryEffect = getattr(K, t); st.Duration = 1.0 + (k % 3) * 0.5
            if k % 4 == 0: st.AdvanceOnTime = True; st.AdvanceTime = 3
            sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = f"Notas de la diapositiva {k + 1}: {texto(k)}"
    efectos = [(1, "Aparecer"), (2, "Volar"), (10, "Desvanecer"), (22, "Barrer"), (23, "Zoom"), (16, "Dividir"), (26, "Rebotar"), (61, "Girar"),
               (59, "Aumentar/encoger"), (62, "Transparencia"), (63, "Destello en negrita"), (101, "Trayectoria círculo"), (78, "Balancear"), (54, "Cambiar relleno")]
    with b.paso("animaciones de entrada, énfasis, salida y trayectoria"):
        sl = P.diapo("Blank")
        seq = sl.TimeLine.MainSequence
        for k, (ef, nombre) in enumerate(efectos):
            s = sl.Shapes.AddShape(5, 20 + (k % 5) * 188, 30 + (k // 5) * 170, 170, 120); s.TextFrame2.TextRange.Text = nombre
            e = seq.AddEffect(s, ef, 0, 1 if k % 3 == 0 else (2 if k % 3 == 1 else 3)); e.Timing.Duration = 0.5 + (k % 4) * 0.25
            if ef in (2, 10, 22, 23, 16) and k % 2 == 0:
                try: e.Exit = True
                except Exception: pass
    with b.paso("animación por párrafo y desencadenador"):
        sl = P.diapo("Title and Content", "Viñetas que aparecen una por una")
        sl.Shapes(2).TextFrame.TextRange.Text = "Primera\rSegunda\rTercera\rCuarta"
        seq = sl.TimeLine.MainSequence; seq.AddEffect(sl.Shapes(2), 2, 1, 1)
        bt = sl.Shapes.AddShape(130, 800, 440, 120, 60); bt.TextFrame2.TextRange.Text = "Mostrar"
        tg = sl.TimeLine.InteractiveSequences.Add(); e = tg.AddTriggerEffect(sl.Shapes(1), 63, 4, bt)
    with b.paso("diapositiva oculta"):
        sl = P.diapo("Title Only", "Diapositiva oculta"); sl.SlideShowTransition.Hidden = True
    with b.paso("secciones"):
        sp = P.pr.SectionProperties; sp.AddBeforeSlide(1, "Transiciones"); sp.AddBeforeSlide(len(TRANSICIONES) + 1, "Animaciones")
    with b.paso("audio y video"):
        sl = P.diapo("Title Only", "Medios")
        for f in ["prueba.mp4", "prueba.m4a", "prueba.mp3", "prueba.wav"]:
            ruta = os.path.join(IMG, f)
            if os.path.exists(ruta): sl.Shapes.AddMediaObject2(ruta, False, True, 100, 150, 400, 225)
    with b.paso("zoom de diapositiva (resumen)"):
        sl = P.diapo("Title Only", "Zoom")
        try: sl.Shapes.AddSlideZoom(2, 100, 150, 240, 135)
        except Exception: sl.Shapes.AddPicture(os.path.join(IMG, "barras.png"), False, True, 100, 150, 240, 135)
    with b.paso("vínculos a otra diapositiva y botón de acción"):
        sl = P.diapo("Blank")
        s = caja(sl, 100, 100, 600, 60, "Ir a la diapositiva 3", 28); s.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.SubAddress = "258,3,Diapositiva 3"
        a = sl.Shapes.AddShape(129, 100, 300, 100, 60)
    return P.guardar()


# ---------------------------------------------------------------- p09: presentación realista
def p09(p):
    P = Pres(p, "p09-presentacion"); b = P.b
    with b.paso("tema Retrospect o Ion"):
        for t in ["Ion.thmx", "Retrospect.thmx", "Facet.thmx"]:
            if os.path.exists(os.path.join(TEMAS_DIR, t)): P.pr.ApplyTheme(os.path.join(TEMAS_DIR, t)); break
    with b.paso("portada"):
        sl = P.diapo("Title Slide", "Proyecto NOM-COLAB"); sl.Shapes(2).TextFrame.TextRange.Text = "Colabora donde no hay Internet\rEquipo 5IM7 · Noviembre de 2026"
    with b.paso("agenda"):
        sl = P.diapo("Title and Content", "Agenda"); sl.Shapes(2).TextFrame.TextRange.Text = "Problema\rPropuesta\rArquitectura\rResultados\rSiguientes pasos"
    with b.paso("problema con imagen"):
        sl = P.diapo("Two Content", "El problema"); sl.Shapes(2).TextFrame.TextRange.Text = "Sin red no hay colaboración\rLos archivos se pierden entre versiones\rOffice y LibreOffice no se ven igual"
        try: sl.Shapes(3).Delete()
        except Exception: pass
        sl.Shapes.AddPicture(os.path.join(IMG, "foto1.jpg"), False, True, 500, 140, 420, 262)
    with b.paso("propuesta con SmartArt"):
        sl = P.diapo("Title Only", "Propuesta")
        lay = None
        for k in range(1, p.SmartArtLayouts.Count + 1):
            if p.SmartArtLayouts(k).Name == "Basic Process": lay = p.SmartArtLayouts(k); break
        sa = sl.Shapes.AddSmartArt(lay, 60, 160, 840, 280)
        for k, t in enumerate(["Crear", "Compartir", "Editar juntos", "Exportar"]):
            if k >= sa.SmartArt.AllNodes.Count: sa.SmartArt.AllNodes.Add()
            sa.SmartArt.AllNodes(k + 1).TextFrame2.TextRange.Text = t
    with b.paso("arquitectura con formas y conectores"):
        sl = P.diapo("Title Only", "Arquitectura")
        cajas = {}
        for k, (t, x, y) in enumerate([("Navegador", 80, 180), ("Aplicación", 80, 360), ("Servidor", 420, 270), ("Raspberry Pi", 760, 180), ("Teléfono", 760, 360)]):
            s = sl.Shapes.AddShape(5, x, y, 140, 70); s.TextFrame2.TextRange.Text = t; cajas[t] = s
        for a, c in [("Navegador", "Servidor"), ("Aplicación", "Servidor"), ("Servidor", "Raspberry Pi"), ("Servidor", "Teléfono")]:
            cn = sl.Shapes.AddConnector(2, 0, 0, 1, 1); cn.ConnectorFormat.BeginConnect(cajas[a], 4); cn.ConnectorFormat.EndConnect(cajas[c], 2); cn.Line.EndArrowheadStyle = 2
    with b.paso("resultados con tabla y gráfica"):
        sl = P.diapo("Title Only", "Resultados")
        t = sl.Shapes.AddTable(4, 3, 40, 150, 420, 200).Table
        for i, f in enumerate([["Prueba", "Antes", "Después"], ["Fidelidad Word", "64 %", "—"], ["Fidelidad Excel", "—", "—"], ["Fidelidad PowerPoint", "—", "—"]]):
            for j, v in enumerate(f): t.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Text = v
        sh = sl.Shapes.AddChart2(-1, 51, 500, 140, 420, 340)
        try: sh.Chart.ChartData.Workbook.Close()
        except Exception: pass
    with b.paso("cita grande"):
        sl = P.diapo("Section Header", "«Lo que no se ve igual, no se entrega»");
        try: sl.Shapes(2).TextFrame.TextRange.Text = "Lema del equipo"
        except Exception: pass
    with b.paso("cierre con transición y notas"):
        sl = P.diapo("Title Slide", "¡Gracias!"); sl.SlideShowTransition.EntryEffect = K.ppEffectFadeSmoothly
        sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "Agradecer y abrir preguntas."
    return P.guardar()


# ---------------------------------------------------------------- p10: combinaciones raras
def p10(p):
    P = Pres(p, "p10-raros", 595.3, 841.9); b = P.b
    with b.paso("diapositiva A4 vertical"):
        sl = P.diapo("Title and Content", "A4 vertical"); sl.Shapes(2).TextFrame.TextRange.Text = texto(0, 3)
    with b.paso("formas que se salen de la diapositiva"):
        sl = P.diapo("Blank")
        s = sl.Shapes.AddShape(9, -100, -100, 300, 300); s.Fill.ForeColor.RGB = rgb(255, 0, 0)
        s = sl.Shapes.AddShape(1, 450, 700, 300, 300); s.Fill.ForeColor.RGB = rgb(0, 0, 255)
    with b.paso("grupos anidados girados y volteados con texto"):
        sl = P.diapo("Blank")
        a = sl.Shapes.AddShape(1, 100, 100, 150, 80); a.TextFrame2.TextRange.Text = "A"
        c = sl.Shapes.AddShape(9, 200, 150, 100, 100); c.TextFrame2.TextRange.Text = "B"
        g = sl.Shapes.Range([a.Name, c.Name]).Group(); g.Rotation = 45
        d = sl.Shapes.AddShape(92, 300, 300, 120, 120); d.TextFrame2.TextRange.Text = "C"
        g2 = sl.Shapes.Range([g.Name, d.Name]).Group(); g2.Flip(1); g2.Rotation = 200
    with b.paso("texto enorme, diminuto y con 9 niveles"):
        sl = P.diapo("Blank")
        caja(sl, 10, 10, 575, 300, "Ñ", 200)
        caja(sl, 10, 330, 575, 20, "letra de 4 pt " * 10, 4)
        s = caja(sl, 10, 360, 575, 470, "\r".join(f"Nivel {k}" for k in range(1, 10)), 16)
        for k in range(9): s.TextFrame2.TextRange.GetParagraphs(k + 1, 1).ParagraphFormat.IndentLevel = k + 1
    with b.paso("transparencias encimadas"):
        sl = P.diapo("Blank")
        for k, c in enumerate([rgb(255, 0, 0), rgb(0, 255, 0), rgb(0, 0, 255)]):
            s = sl.Shapes.AddShape(9, 120 + k * 90, 200 + (k % 2) * 90, 260, 260); s.Fill.ForeColor.RGB = c; s.Fill.Transparency = 0.5; s.Line.Visible = False
    with b.paso("tabla con imagen de fondo en celda y texto vertical"):
        sl = P.diapo("Blank")
        t = sl.Shapes.AddTable(3, 3, 40, 100, 500, 400).Table
        t.Cell(2, 2).Shape.Fill.UserPicture(os.path.join(IMG, "foto2.jpg"))
        t.Cell(1, 1).Shape.TextFrame.Orientation = 2; t.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Vertical"
    with b.paso("línea de largo cero y forma de tamaño cero"):
        sl = P.diapo("Blank")
        sl.Shapes.AddLine(100, 100, 100, 100); s = sl.Shapes.AddShape(1, 200, 200, 0.1, 0.1)
    with b.paso("todos los alfabetos en una forma girada"):
        sl = P.diapo("Blank")
        s = sl.Shapes.AddShape(1, 60, 200, 470, 300); s.Rotation = 8
        s.TextFrame2.TextRange.Text = "Español ñ · Ελληνικά · Кириллица · العربية · עברית · 中文 · 日本語 · 한국어 · हिन्दी · ไทย · 😀👍🏽👨‍👩‍👧"
    return P.guardar()


if __name__ == "__main__":
    pythoncom.CoInitialize()
    hacer = {k: v for k, v in globals().items() if k[:1] == "p" and k[1:].isdigit()}
    pedidos = sys.argv[1:] or sorted(hacer)
    p = powerpoint()
    try: p.Visible = True
    except Exception: pass
    for k in pedidos:
        print("==", k, flush=True)
        try: print(" ->", hacer[k](p), flush=True)
        except Exception: traceback.print_exc()
    # PowerPoint es una sola copia: solo se cierra si nadie más tiene presentaciones abiertas
    try:
        if p.Presentations.Count == 0: p.Quit()
    except Exception: pass

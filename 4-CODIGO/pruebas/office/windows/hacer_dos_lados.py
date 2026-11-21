"""w17-dos-lados.docx: texto a los DOS lados de un dibujo (ajuste «Cuadrado», a ambos lados: lo de siempre en Word).
Rectángulos de 1, 2 y 3 pulgadas en varias posiciones de la columna —también con un lado muy angosto—, con 9 pt de
distancia al texto, para medir cuándo Word pone texto de cada lado y cómo parte cada renglón (verdad.py da cada tramo
como su renglón, con su x). Uso: hacer_dos_lados.py SALIDA.docx"""
import os, sys
from comun import word, texto

salida = os.path.abspath(sys.argv[1])
w = word()
d = w.Documents.Add()
try:
    ps = d.PageSetup
    ps.PageWidth, ps.PageHeight = 612, 792
    ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 72   # columna de 468 pt
    casos = []
    for ancho in (72, 144, 216):
        for izq in (0, 6, 14, 24, 36, 54, 72, 108, 144, 180, 216, 252, 300, 340, 380, 396):
            if izq + ancho <= 468: casos.append((ancho, izq))
    r = d.Range(0, 0)
    for k, (ancho, izq) in enumerate(casos):
        fin = d.Content; fin.InsertParagraphAfter()
        p = d.Paragraphs(d.Paragraphs.Count)
        p.Range.Text = f"Caso {k + 1}: dibujo de {ancho} pt a {izq} pt del borde izquierdo de la columna."
        p.Range.Font.Bold = True
        for j in range(2):
            d.Content.InsertParagraphAfter()
            q = d.Paragraphs(d.Paragraphs.Count)
            q.Range.Text = texto(k + j * 3, 5)
            q.Range.Font.Bold = False
            if j == 0:
                sh = d.Shapes.AddShape(1, izq, 6, ancho, 54, q.Range)
                sh.RelativeHorizontalPosition = 2   # columna
                sh.RelativeVerticalPosition = 2     # párrafo
                sh.Left, sh.Top = izq, 6
                sh.WrapFormat.Type = 0              # cuadrado
                sh.WrapFormat.Side = 0              # a ambos lados
                sh.WrapFormat.DistanceLeft = sh.WrapFormat.DistanceRight = 9
                sh.WrapFormat.DistanceTop = sh.WrapFormat.DistanceBottom = 0
                sh.Fill.Visible = False
                sh.Name = f"caso{k + 1}"
    d.Paragraphs(1).Range.Delete()
    d.SaveAs2(salida, 16)
finally:
    d.Close(0)
    w.Quit()
print("hecho", len(casos), flush=True)

"""Cómo muestra Word cada EMF o WMF: cada uno en su propia página, del tamaño exacto de la imagen, colocado por Word
en la esquina (0, 0) y exportado a PDF por Word. En Linux, cada página del PDF se pasa a PNG y se compara con lo que
dibuja NOM (pruebas/office/metarchivos_real.py).

  correr.ps1 meta metarchivos_word.py <carpeta con los archivos> <lista.txt: nombre<TAB>ancho<TAB>alto en px a 192 ppp> <salida.pdf>
"""
import os, sys
from comun import word

carpeta, lista, salida = [os.path.abspath(os.path.join(os.path.expanduser("~"), "nomoffice", a)) if not os.path.isabs(a) else a for a in sys.argv[1:4]]
filas = [l.split("\t") for l in open(lista, encoding="utf8").read().splitlines() if l.count("\t") >= 2]
w = word()
d = w.Documents.Add()
try:
    d.Content.Font.Size = 1
    d.Content.ParagraphFormat.SpaceAfter = 0; d.Content.ParagraphFormat.SpaceBefore = 0; d.Content.ParagraphFormat.LineSpacingRule = 0
    hechos = 0
    for i, (nombre, ancho, alto) in enumerate(filas):
        wpt, hpt = int(ancho) * 72 / 192, int(alto) * 72 / 192   # (la lista viene a 192 ppp: el doble de 96)
        if i:
            # salto de sección al final (el rango se cierra antes: si no, el salto reemplaza el párrafo con la imagen anterior)
            r = d.Content; r.Collapse(0); r.InsertBreak(2)
        s = d.Sections(d.Sections.Count)
        ps = s.PageSetup
        ps.TopMargin = 0; ps.BottomMargin = 0; ps.LeftMargin = 0; ps.RightMargin = 0; ps.HeaderDistance = 0; ps.FooterDistance = 0
        ps.PageWidth = max(36, min(1584, wpt)); ps.PageHeight = max(36, min(1584, hpt))
        anc = d.Paragraphs.Last.Range
        try:
            sh = d.Shapes.AddPicture(os.path.join(carpeta, nombre), False, True, 0, 0, wpt, hpt, anc)
            sh.WrapFormat.Type = 3
            sh.RelativeHorizontalPosition = 1; sh.RelativeVerticalPosition = 1
            sh.LockAspectRatio = 0; sh.Width = wpt; sh.Height = hpt; sh.Left = 0; sh.Top = 0
            hechos += 1
        except Exception as e:
            print("MAL", nombre, str(e)[:200], flush=True)
    d.ExportAsFixedFormat(salida, 17)
    print(f"listos {hechos} de {len(filas)} · páginas {d.ComputeStatistics(2)}", flush=True)
finally:
    d.Close(0)
    w.Quit()

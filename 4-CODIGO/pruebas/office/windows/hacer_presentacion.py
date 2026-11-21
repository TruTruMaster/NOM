"""p13-presentacion.pptx: lo de la pestaña Presentación de PowerPoint (y las secciones de Inicio), hecho en PowerPoint,
para medir cómo lo lee NOM: diapositivas ocultas, secciones, presentaciones personalizadas, intervalos ensayados y la
configuración de la presentación. Uso: hacer_presentacion.py SALIDA.pptx"""
import os, sys
import pythoncom
import win32com.client
from comun import powerpoint

salida = os.path.abspath(sys.argv[1])
pp = powerpoint()
pr = pp.Presentations.Add(0)   # sin ventana
try:
    titulos = ["Portada", "Objetivos", "Datos internos (oculta)", "Resultados", "Conclusiones", "Gracias"]
    for i, t in enumerate(titulos, 1):
        s = pr.Slides.Add(i, 2)   # ppLayoutText
        s.Shapes(1).TextFrame.TextRange.Text = t
        s.Shapes(2).TextFrame.TextRange.Text = f"Diapositiva {i} de {len(titulos)}"
    # ocultar la 3
    pr.Slides(3).SlideShowTransition.Hidden = -1
    # intervalos (como si se hubieran ensayado): 2.5 s en la 2 y 4 s en la 4; la 5 solo con clic
    for n, seg in ((2, 2.5), (4, 4.0)):
        tr = pr.Slides(n).SlideShowTransition
        tr.AdvanceOnTime = -1
        tr.AdvanceTime = seg
    # secciones: Inicio (1-2), Desarrollo (3-4), Cierre (5-6)
    sp = pr.SectionProperties
    sp.AddBeforeSlide(1, "Inicio")
    sp.AddBeforeSlide(3, "Desarrollo")
    sp.AddBeforeSlide(5, "Cierre")
    # presentaciones personalizadas
    ids = [pr.Slides(i).SlideID for i in range(1, 7)]
    arr = lambda xs: win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, xs)
    ns = pr.SlideShowSettings.NamedSlideShows
    ns.Add("Corta", arr([ids[0], ids[3], ids[5]]))
    ns.Add("Al revés", arr(list(reversed(ids))))
    # configurar la presentación: repetir hasta presionar Esc
    pr.SlideShowSettings.LoopUntilStopped = -1
    pr.SaveAs(salida, 24)   # ppSaveAsOpenXMLPresentation
finally:
    pr.Close()
print("hecho", flush=True)

"""w18-cambios.docx: cambios marcados de Word (Revisar → Control de cambios) de dos personas, para medir cómo los lee NOM
(windows/cambios_ver.py dice qué ve Word): agregar y borrar texto, reemplazar, un párrafo nuevo, un cambio en una tabla y
uno de formato. El nombre de usuario de Word se cambia para cada persona y al final se deja el que tenía.
Uso: hacer_cambios.py SALIDA.docx"""
import os, sys
from comun import word

salida = os.path.abspath(sys.argv[1])
w = word()
antes = w.UserName, w.UserInitials
d = w.Documents.Add()
try:
    def parrafo(txt):
        d.Content.InsertParagraphAfter()
        d.Paragraphs(d.Paragraphs.Count).Range.Text = txt
    d.Paragraphs(1).Range.Text = "El informe anual de ventas."
    parrafo("Las ventas crecieron en el norte y en el sur.")
    parrafo("Resumen: el año fue bueno.")
    parrafo("Fin del texto.")
    parrafo("")
    t = d.Tables.Add(d.Paragraphs(d.Paragraphs.Count).Range, 2, 2)
    for (f, c, txt) in ((1, 1, "uno"), (1, 2, "dos"), (2, 1, "tres"), (2, 2, "cuatro")): t.Cell(f, c).Range.Text = txt
    def buscar(txt, en=None):
        r = (en or d.Content).Duplicate
        f = r.Find
        f.ClearFormatting()
        if not f.Execute(txt): raise SystemExit("no está: " + txt)
        return r
    d.TrackRevisions = True
    # Ana: borra «anual » y agrega « de 2026» antes del punto; agrega una oración al final del resumen; pone «Resumen» en negrita
    w.UserName, w.UserInitials = "Ana Pérez", "AP"
    buscar("anual ").Delete()
    r = buscar("ventas.")
    d.Range(r.End - 1, r.End - 1).InsertAfter(" de 2026")
    r = buscar("el año fue bueno.")
    d.Range(r.End, r.End).InsertAfter(" Todo va bien.")
    buscar("Resumen").Font.Bold = True
    # Luis: cambia «y en el sur» por «y en el centro», agrega un párrafo, y «!» en la celda «dos» de la tabla
    w.UserName, w.UserInitials = "Luis Gómez", "LG"
    r = buscar("y en el sur")
    r.Delete()
    d.Range(r.Start, r.Start).InsertAfter("y en el centro")
    r = buscar("Fin del texto.")
    d.Range(r.Start, r.Start).InsertBefore("Párrafo nuevo de Luis.\r")
    r = t.Cell(1, 2).Range
    d.Range(r.End - 1, r.End - 1).InsertAfter("!")
    d.SaveAs2(salida, 16)
finally:
    w.UserName, w.UserInitials = antes
    d.Close(0)
    w.Quit()
print("hecho", flush=True)

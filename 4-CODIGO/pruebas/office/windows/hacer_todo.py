"""Arma con Office de verdad UN archivo por programa que junta todo lo de los archivos de prueba.

  python hacer_todo.py word|excel|powerpoint

  word        archivos\\t01-todo-word.docx: w01 a w12, uno tras otro, cada uno en su propia sección (página nueva)
  excel       archivos\\t02-todo-excel.xlsx: todas las hojas de x01 a x09 (cada libro se copia entero, para que
              sus fórmulas entre hojas, nombres, tablas, tablas dinámicas y gráficas sigan apuntando a sus hojas);
              cada hoja lleva delante el número de su libro («x02 Funciones»)
  powerpoint  archivos\\t03-todo-powerpoint.pptx: las diapositivas de p01 a p11 con su formato de origen (cada una
              con su patrón y su tema)

La idea (del dueño, 17-nov): un solo archivo con TODO de cada programa, para probar que todo sirve junto:
importar, ver, editar, descargar, .nom, espacios cifrados. Deja también NOMBRE.funciones.json con lo que hizo
Office y lo que no.
"""
import os, sys, time
from comun import *


def word_todo():
    w = word()
    b = Bitacora("t01-todo-word")
    d = w.Documents.Add()
    try:
        primero = True
        for k in range(1, 13):
            nombres = [f for f in os.listdir(ARCHIVOS) if f.startswith(f"w{k:02d}-") and f.endswith(".docx")]
            if not nombres: b.mal.append((f"w{k:02d}", "no está")); continue
            ruta = os.path.join(ARCHIVOS, nombres[0])
            try:
                fin = d.Content.End - 1
                rg = d.Range(fin, fin)
                if not primero:
                    rg.InsertBreak(2)   # wdSectionBreakNextPage
                    fin = d.Content.End - 1
                    rg = d.Range(fin, fin)
                rg.InsertFile(ruta)
                primero = False
                b.ok.append(nombres[0])
            except Exception as ex:
                b.mal.append((nombres[0], str(ex)[:200]))
        sal = os.path.join(ARCHIVOS, "t01-todo-word.docx")
        d.SaveAs2(sal, 16)
        print("páginas:", d.ComputeStatistics(2), "secciones:", d.Sections.Count, flush=True)
    finally:
        d.Close(0)
        w.Quit()
    guardar_json(os.path.join(ARCHIVOS, "t01-todo-word.funciones.json"), b.resumen())


def excel_todo():
    x = excel()
    x.DisplayAlerts = False
    # visible pero fuera de la pantalla: sin ventana Excel no copia bien las gráficas
    x.Visible = True
    try: x.WindowState = -4143; x.Left = -16000; x.Top = 0
    except Exception: pass
    b = Bitacora("t02-todo-excel")
    dest = x.Workbooks.Add()
    try:
        base = dest.Sheets.Count
        for k in range(1, 10):
            nombres = [f for f in os.listdir(ARCHIVOS) if f.startswith(f"x{k:02d}-") and f.endswith(".xlsx")]
            if not nombres: b.mal.append((f"x{k:02d}", "no está")); continue
            ruta = os.path.join(ARCHIVOS, nombres[0])
            src = x.Workbooks.Open(ruta, UpdateLinks=0, ReadOnly=True)
            try:
                try: src.Unprotect()
                except Exception: pass
                ocultas = {}
                for i in range(1, src.Sheets.Count + 1):
                    s = src.Sheets(i)
                    ocultas[s.Name] = s.Visible
                    if s.Visible != -1:
                        try: s.Visible = -1
                        except Exception as ex: b.mal.append((f"{nombres[0]}:{s.Name}", "no se pudo mostrar: " + str(ex)[:120]))
                antes = dest.Sheets.Count
                src.Sheets.Copy(After=dest.Sheets(dest.Sheets.Count))
                dest = x.ActiveWorkbook if x.ActiveWorkbook.Name == dest.Name else dest
                nuevas = dest.Sheets.Count - antes
                for i in range(antes + 1, dest.Sheets.Count + 1):
                    s = dest.Sheets(i)
                    viejo = s.Name
                    try: s.Name = (f"x{k:02d} " + viejo)[:31]
                    except Exception as ex: b.mal.append((f"{nombres[0]}:{viejo}", "renombrar: " + str(ex)[:120]))
                    v = ocultas.get(viejo, -1)
                    if v != -1:
                        try: s.Visible = v
                        except Exception: pass
                b.ok.append(f"{nombres[0]} ({nuevas} hojas)")
            except Exception as ex:
                b.mal.append((nombres[0], str(ex)[:200]))
            finally:
                src.Close(False)
        # la hoja vacía con la que nace el libro
        for i in range(base, 0, -1):
            try: dest.Sheets(i).Delete()
            except Exception: pass
        dest.Sheets(1).Activate()
        sal = os.path.join(ARCHIVOS, "t02-todo-excel.xlsx")
        dest.SaveAs(sal, 51)
        print("hojas:", dest.Sheets.Count, flush=True)
    finally:
        dest.Close(False)
        x.Quit()
    guardar_json(os.path.join(ARCHIVOS, "t02-todo-excel.funciones.json"), b.resumen())


def powerpoint_todo():
    p = powerpoint()
    b = Bitacora("t03-todo-powerpoint")
    nombres = sorted(f for f in os.listdir(ARCHIVOS) if f[:1] == "p" and f[1:3].isdigit() and f.endswith(".pptx"))
    sal = os.path.join(ARCHIVOS, "t03-todo-powerpoint.pptx")
    # se empieza con una copia del primero (su tamaño de diapositiva, su patrón y su tema)
    base = p.Presentations.Open(os.path.join(ARCHIVOS, nombres[0]), ReadOnly=False, Untitled=True, WithWindow=True)
    try:
        try: p.WindowState = 1; p.Left = -16000; p.Top = 0
        except Exception: pass
        b.ok.append(f"{nombres[0]} ({base.Slides.Count} diapositivas)")
        for n in nombres[1:]:
            ruta = os.path.join(ARCHIVOS, n)
            src = p.Presentations.Open(ruta, ReadOnly=True, Untitled=False, WithWindow=True)
            try:
                antes = base.Slides.Count
                ok = False
                # con el formato de origen: copiar las diapositivas y pegarlas «manteniendo el formato de origen»
                try:
                    src.Slides.Range().Copy()
                    time.sleep(1.0)
                    base.Windows(1).Activate()
                    base.Windows(1).View.GotoSlide(base.Slides.Count)
                    p.CommandBars.ExecuteMso("PasteSourceFormatting")
                    for _ in range(120):
                        time.sleep(0.5)
                        if base.Slides.Count >= antes + src.Slides.Count: break
                    ok = base.Slides.Count == antes + src.Slides.Count
                except Exception as ex:
                    b.mal.append((n, "pegar con formato de origen: " + str(ex)[:150]))
                if not ok:
                    # de respaldo: insertar (toma el diseño del destino)
                    base.Slides.InsertFromFile(ruta, base.Slides.Count)
                    b.mal.append((n, "se insertó con el diseño del destino"))
                b.ok.append(f"{n} ({base.Slides.Count - antes} diapositivas)")
            except Exception as ex:
                b.mal.append((n, str(ex)[:200]))
            finally:
                src.Close()
        base.SaveAs(sal, 24)
        print("diapositivas:", base.Slides.Count, flush=True)
    finally:
        base.Close()
    guardar_json(os.path.join(ARCHIVOS, "t03-todo-powerpoint.funciones.json"), b.resumen())


if __name__ == "__main__":
    import pythoncom
    pythoncom.CoInitialize()
    {"word": word_todo, "excel": excel_todo, "powerpoint": powerpoint_todo}[sys.argv[1]]()

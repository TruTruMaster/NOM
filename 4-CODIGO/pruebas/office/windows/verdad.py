"""Lo que Office de verdad hace con un archivo: cómo lo pinta y dónde queda cada cosa.

  python verdad.py ARCHIVO.docx|.xlsx|.pptx ...    (rutas relativas a nomoffice\\ o absolutas)

Por cada archivo deja junto a él (en archivos\\verdad\\):
  NOMBRE.verdad.json   medidas exactas (Word: cada renglón de cada página con su texto y posición;
                       Excel: cada celda con el texto que muestra, colores y tamaños; PowerPoint: cada
                       forma con su posición y los renglones de su texto)
  NOMBRE.pdf           cómo lo imprime Office (Word y Excel)
  NOMBRE-NN.png        cada diapositiva (PowerPoint)
Además anota si Office pidió reparar el archivo o no lo pudo abrir.
"""
import glob, os, sys, threading, time
from comun import *

SALE = os.path.join(ARCHIVOS, "verdad")
os.makedirs(SALE, exist_ok=True)


def pids(imagen):
    """Los procesos con ese nombre que corren ahora (para saber cuál abrimos nosotros)."""
    salida = os.popen(f'tasklist /FI "IMAGENAME eq {imagen}" /FO CSV /NH').read()
    return {int(c.split('","')[1]) for c in salida.splitlines() if c.startswith('"') and '","' in c and c.split('","')[1].isdigit()}


def vigilar(segundos, proceso, que, pid=None):
    """Si Office se queda trabado (un aviso de reparar invisible), cierra SOLO el proceso que abrimos nosotros
    (nunca un Office que la persona tenga abierto) y lo anota. Sin saber cuál es, solo avisa."""
    ev = threading.Event()
    def fin():
        if not ev.wait(segundos):
            if pid:
                print(f"  [trabado] {que}: más de {segundos} s; se cierra {proceso} (proceso {pid}, el nuestro)", flush=True)
                os.system(f"taskkill /F /PID {pid} >nul 2>&1")
            else: print(f"  [trabado] {que}: más de {segundos} s; no se cierra nada (no se sabe cuál {proceso} es el nuestro)", flush=True)
    threading.Thread(target=fin, daemon=True).start()
    return ev


def nuevo(imagen, antes):
    """El proceso que apareció al abrir Office, si fue exactamente uno."""
    n = pids(imagen) - antes
    return next(iter(n)) if len(n) == 1 else None


# ------------------------------------------------------------------------------------------- Word
HISTORIAS = {1: "cuerpo", 2: "notas", 3: "notasfinales", 4: "comentarios", 5: "cuadro", 6: "enc-par", 7: "enc",
             8: "pie-par", 9: "pie", 10: "enc-primera", 11: "pie-primera"}


def lineas_de(rc, prof=0):
    """Los renglones de una zona de la página. Una fila de tabla (LineType 1) trae sus celdas, cada una con sus
    renglones (y las tablas de adentro, hasta 3 niveles): así se comparan también los renglones dentro de las celdas."""
    out = []
    ls = rc.Lines
    for k in range(1, ls.Count + 1):
        ln = ls(k)
        li = {"t": ln.LineType, "x": ln.Left, "y": ln.Top, "w": ln.Width, "h": ln.Height}
        try: li["texto"] = ln.Range.Text
        except Exception: li["texto"] = None
        if li["t"] == 1 and prof < 3:
            try:
                cs = ln.Rectangles
                li["celdas"] = []
                for c in range(1, cs.Count + 1):
                    rcel = cs(c)
                    cel = {"x": rcel.Left, "y": rcel.Top, "w": rcel.Width, "h": rcel.Height}
                    try: cel["lineas"] = lineas_de(rcel, prof + 1)
                    except Exception as e: cel["error"] = str(e)[:120]
                    li["celdas"].append(cel)
            except Exception as e: li["error_celdas"] = str(e)[:120]
        out.append(li)
    return out


def verdad_word(ruta, nombre):
    antes = pids("WINWORD.EXE")
    w = word()
    datos = {"archivo": os.path.basename(ruta), "app": "word", "abrio": False}
    ev = vigilar(240, "WINWORD.EXE", nombre, nuevo("WINWORD.EXE", antes))
    try:
        t0 = time.time()
        try:
            d = w.Documents.Open(ruta, False, True, False, "", "", False, "", "", 0, 0, True, False, 0, True)   # ventana del documento visible (Word sigue oculto): sin ventana no hay páginas
        except Exception as e:
            datos["error"] = str(getattr(e, "excepinfo", None) or e)[:400]; return datos
        datos["abrio"] = True; datos["ms_abrir"] = int((time.time() - t0) * 1000)
        v = d.ActiveWindow.View
        v.Type = 3                                  # diseño de impresión
        try:
            v.RevisionsFilter.Markup = 0; v.RevisionsFilter.View = 0   # sin marcas, como queda (final)
            v.ShowRevisionsAndComments = False
        except Exception: pass
        d.Repaginate()
        est = {}
        for k, n in [("paginas", 2), ("palabras", 0), ("caracteres", 3), ("parrafos", 4), ("lineas", 1)]:
            try: est[k] = d.ComputeStatistics(n)
            except Exception: pass
        est.update(secciones=d.Sections.Count, tablas=d.Tables.Count, imagenes_en_linea=d.InlineShapes.Count,
                   formas=d.Shapes.Count, notas=d.Footnotes.Count, notas_finales=d.Endnotes.Count,
                   comentarios=d.Comments.Count, revisiones=d.Revisions.Count, campos=d.Fields.Count,
                   controles=d.ContentControls.Count, ecuaciones=d.OMaths.Count, hipervinculos=d.Hyperlinks.Count,
                   marcadores=d.Bookmarks.Count)
        datos["estadisticas"] = est
        secs = []
        for s in d.Sections:
            ps = s.PageSetup
            secs.append({"ancho": ps.PageWidth, "alto": ps.PageHeight, "orient": ps.Orientation, "arriba": ps.TopMargin,
                         "abajo": ps.BottomMargin, "izq": ps.LeftMargin, "der": ps.RightMargin, "columnas": ps.TextColumns.Count,
                         "inicio": s.Range.Information(3)})
        datos["secciones"] = secs
        paginas = []
        pgs = d.ActiveWindow.Panes(1).Pages
        for i in range(1, pgs.Count + 1):
            pg = pgs(i)
            pag = {"n": i, "ancho": pg.Width, "alto": pg.Height, "zonas": []}
            for j in range(1, pg.Rectangles.Count + 1):
                rc = pg.Rectangles(j)
                z = {"lineas": []}
                for k2, a in [("tipo", "RectangleType"), ("x", "Left"), ("y", "Top"), ("w", "Width"), ("h", "Height")]:
                    try: z[k2] = getattr(rc, a)
                    except Exception: z[k2] = None
                try: z["historia"] = HISTORIAS.get(rc.Range.StoryType, rc.Range.StoryType)
                except Exception: z["historia"] = None
                try: z["lineas"] = lineas_de(rc)
                except Exception as e: z["error"] = str(e)[:120]
                pag["zonas"].append(z)
            paginas.append(pag)
        datos["paginas"] = paginas
        pars = []
        for k, p in enumerate(d.Paragraphs):
            if k >= 3000: break
            r = p.Range
            try:
                ini = d.Range(r.Start, r.Start)
                pars.append({"texto": r.Text[:120], "pagina": ini.Information(3), "estilo": str(p.Style), "fuente": r.Font.Name,
                             "tam": r.Font.Size, "tabla": bool(r.Information(12))})
            except Exception: pass
        datos["parrafos"] = pars
        pdf = os.path.join(SALE, nombre + ".pdf")
        for intento in range(3):
            try:
                if intento == 0: d.ExportAsFixedFormat(pdf, 17, False, 0, 0, 1, 1, 0, False, True, 0, False, True, False)
                elif intento == 1: d.ExportAsFixedFormat(pdf, 17)
                else: d.SaveAs2(pdf, 17)
                datos["pdf"] = os.path.basename(pdf); break
            except Exception as e:
                datos["error_pdf"] = str(getattr(e, "excepinfo", None) or e)[:300]; time.sleep(1)
        # lo que flota: su ajuste, su página y su lugar EN LA HOJA (después del PDF: para saberlo se cambia su
        # «relativo a» por la hoja, y Word recalcula Left y Top dejándolo donde estaba; el archivo no se guarda)
        formas = []
        try:
            for k in range(1, min(d.Shapes.Count, 400) + 1):
                sh = d.Shapes(k)
                f = {"nombre": sh.Name, "tipo": sh.Type, "w": sh.Width, "h": sh.Height}
                for k2, fn in [("ajuste", lambda: sh.WrapFormat.Type), ("lado", lambda: sh.WrapFormat.Side), ("pagina", lambda: sh.Anchor.Information(3)),
                               ("relh", lambda: sh.RelativeHorizontalPosition), ("relv", lambda: sh.RelativeVerticalPosition), ("left", lambda: sh.Left), ("top", lambda: sh.Top),
                               ("dist", lambda: [sh.WrapFormat.DistanceTop, sh.WrapFormat.DistanceBottom, sh.WrapFormat.DistanceLeft, sh.WrapFormat.DistanceRight])]:
                    try: f[k2] = fn()
                    except Exception: pass
                formas.append(f)
            for k, f in enumerate(formas, 1):
                try:
                    sh = d.Shapes(k); sh.RelativeHorizontalPosition = 1; sh.RelativeVerticalPosition = 1
                    f["x_hoja"], f["y_hoja"] = sh.Left, sh.Top
                except Exception: pass
        except Exception as e: datos["error_formas"] = str(e)[:200]
        datos["formas"] = formas
        d.Close(0)
    finally:
        ev.set()
        try: w.Quit(0)
        except Exception: pass
    return datos


# ------------------------------------------------------------------------------------------- Excel
def color_de(c):
    try:
        c = int(c)
        if c < 0: return None
        return "#%02X%02X%02X" % (c & 255, (c >> 8) & 255, (c >> 16) & 255)
    except Exception: return None


def captura_ventana(hwnd, ruta):
    """Imagen de una ventana aunque esté fuera de la pantalla (PrintWindow con el contenido completo)."""
    import ctypes, win32gui, win32ui
    from PIL import Image
    l, t, r, b = win32gui.GetWindowRect(hwnd); w, h = r - l, b - t
    hdc = win32gui.GetWindowDC(hwnd); mfc = win32ui.CreateDCFromHandle(hdc); sdc = mfc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap(); bmp.CreateCompatibleBitmap(mfc, w, h); sdc.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(hwnd, sdc.GetSafeHdc(), 2)
    info = bmp.GetInfo(); bits = bmp.GetBitmapBits(True)
    im = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), bits, "raw", "BGRX", 0, 1)
    win32gui.DeleteObject(bmp.GetHandle()); sdc.DeleteDC(); mfc.DeleteDC(); win32gui.ReleaseDC(hwnd, hdc)
    return im, (l, t)


def captura_hoja(x, wb, ws, ruta):
    """La hoja como se ve en Excel a 100 %, desde A1: se recorta la cuadrícula de la ventana (con los
    encabezados de filas y columnas). Devuelve dónde quedó la esquina de A1 dentro de la imagen."""
    import win32gui
    ws.Activate(); win = wb.Windows(1)
    win.ScrollRow = 1; win.ScrollColumn = 1
    time.sleep(0.6)
    im, (l, t) = captura_ventana(x.Hwnd, ruta)
    x0 = win.PointsToScreenPixelsX(0) - l; y0 = win.PointsToScreenPixelsY(0) - t
    # se deja el renglón de los encabezados de columna y la columna de los números de fila
    izq = max(0, x0 - 40); arr = max(0, y0 - 24)
    im = im.crop((izq, arr, im.width - 30, im.height - 70))
    im.save(ruta)
    return {"a1_x": x0 - izq, "a1_y": y0 - arr, "zoom": win.Zoom}


def verdad_excel(ruta, nombre):
    antes = pids("EXCEL.EXE")
    x = excel()
    datos = {"archivo": os.path.basename(ruta), "app": "excel", "abrio": False, "hojas": []}
    ev = vigilar(600, "EXCEL.EXE", nombre, nuevo("EXCEL.EXE", antes))
    try:
        t0 = time.time()
        try: wb = x.Workbooks.Open(ruta, 0, True, None, None, None, True, None, None, False, False, None, False, None, 0)
        except Exception as e:
            datos["error"] = str(getattr(e, "excepinfo", None) or e)[:400]; return datos
        datos["abrio"] = True; datos["ms_abrir"] = int((time.time() - t0) * 1000)
        datos["fuente_estandar"] = [x.StandardFont, x.StandardFontSize]
        # Excel visible pero fuera de la pantalla: sin ventana no dibuja la cuadrícula (y la captura sale negra)
        x.ScreenUpdating = True; x.Visible = True; x.WindowState = -4143
        x.Left = -16000; x.Top = 0; x.Width = 1500; x.Height = 1300
        for hi, ws in enumerate(wb.Worksheets, 1):
            h = {"nombre": ws.Name, "visible": ws.Visible, "rtl": bool(ws.DisplayRightToLeft), "ancho_estandar": ws.StandardWidth, "alto_estandar": ws.StandardHeight}
            try: h["color_pestana"] = color_de(ws.Tab.Color) if ws.Tab.ColorIndex != -4142 else None
            except Exception: pass
            if ws.Visible == -1:
                try:
                    ws.Activate(); win = wb.Windows(1)
                    h.update(cuadricula=bool(win.DisplayGridlines), encabezados=bool(win.DisplayHeadings), zoom=win.Zoom, inmovil=bool(win.FreezePanes),
                             fila_div=win.SplitRow, col_div=win.SplitColumn, vista=win.View)
                except Exception as e: h["error_ventana"] = str(e)[:100]
            ur = ws.UsedRange
            f0, c0 = ur.Row, ur.Column
            nf, nc = ur.Rows.Count, ur.Columns.Count
            ultf, ultc = f0 + nf - 1, c0 + nc - 1
            h["usado"] = ur.Address
            h["columnas"] = [{"c": j, "w": ws.Columns(j).Width, "car": ws.Columns(j).ColumnWidth, "oculta": bool(ws.Columns(j).Hidden)} for j in range(1, min(ultc, 60) + 3)]
            h["filas"] = [{"f": i, "h": ws.Rows(i).Height, "oculta": bool(ws.Rows(i).Hidden)} for i in range(1, min(ultf, 400) + 3)]
            celdas = []
            limite = 6000
            for i in range(f0, min(ultf, 400) + 1):
                for j in range(c0, min(ultc, 60) + 1):
                    if len(celdas) >= limite: break
                    c = ws.Cells(i, j)
                    v = c.Value2
                    df = c.DisplayFormat
                    fondo = None
                    try:
                        if df.Interior.ColorIndex != -4142: fondo = color_de(df.Interior.Color)
                    except Exception: pass
                    combinada = bool(c.MergeCells)
                    if v is None and fondo is None and not combinada and c.Borders.LineStyle in (-4142, None) and not c.HasFormula: continue
                    d = {"f": i, "c": j, "texto": c.Text, "valor": v if not isinstance(v, (bytes,)) else None}
                    if c.HasFormula: d["formula"] = c.Formula
                    try:
                        fo = df.Font
                        d.update(fmt=c.NumberFormat, fuente=fo.Name, tam=fo.Size, b=bool(fo.Bold), i=bool(fo.Italic), u=fo.Underline, tachado=bool(fo.Strikethrough), color=color_de(fo.Color),
                                 fondo=fondo, trama=df.Interior.Pattern, h_al=df.HorizontalAlignment, v_al=df.VerticalAlignment, ajustar=bool(df.WrapText), giro=df.Orientation, sangria=df.IndentLevel)
                    except Exception as e: d["error_fmt"] = str(e)[:80]
                    if combinada: d["combinada"] = c.MergeArea.Address
                    try:
                        bs = {}
                        for k, lado in [("izq", 7), ("arr", 8), ("aba", 9), ("der", 10)]:
                            bd = df.Borders(lado)
                            if bd.LineStyle not in (-4142, None): bs[k] = [bd.LineStyle, bd.Weight, color_de(bd.Color)]
                        if bs: d["bordes"] = bs
                    except Exception: pass
                    celdas.append(d)
            h["celdas"] = celdas
            graficas = []
            for k, co in enumerate(ws.ChartObjects(), 1):
                g = {"nombre": co.Name, "x": co.Left, "y": co.Top, "w": co.Width, "h": co.Height}
                try: g["tipo"] = co.Chart.ChartType; g["titulo"] = co.Chart.ChartTitle.Text if co.Chart.HasTitle else None
                except Exception: pass
                png = os.path.join(SALE, f"{nombre}-h{hi}-g{k}.png")
                # Excel exporta vacía una gráfica que no está a la vista: se lleva la ventana a ella (y se reintenta)
                for intento in range(3):
                    try:
                        if intento or k > 1:
                            win = wb.Windows(1); win.ScrollRow = max(1, co.TopLeftCell.Row - 1); win.ScrollColumn = max(1, co.TopLeftCell.Column - 1)
                            if intento: co.Activate()
                            time.sleep(0.4 + intento)
                        co.Chart.Export(png, "PNG")
                        if os.path.exists(png) and os.path.getsize(png) > 0: g["png"] = os.path.basename(png); break
                    except Exception as e: g["error_png"] = str(e)[:100]
                else: g["error_png"] = g.get("error_png") or "Excel la exportó vacía"
                graficas.append(g)
            h["graficas"] = graficas
            formas = []
            for sh in ws.Shapes:
                f = {"nombre": sh.Name, "tipo": sh.Type, "x": sh.Left, "y": sh.Top, "w": sh.Width, "h": sh.Height, "giro": sh.Rotation}
                try:
                    if sh.HasTextFrame and sh.TextFrame2.HasText: f["texto"] = sh.TextFrame2.TextRange.Text
                except Exception: pass
                formas.append(f)
            h["formas"] = formas
            if ws.Visible == -1:
                try:
                    png = os.path.join(SALE, f"{nombre}-h{hi}.png")
                    h["captura"] = captura_hoja(x, wb, ws, png); h["png"] = os.path.basename(png)
                except Exception as e: h["error_png"] = str(e)[:120]
            datos["hojas"].append(h)
        try:
            pdf = os.path.join(SALE, nombre + ".pdf"); wb.ExportAsFixedFormat(0, pdf); datos["pdf"] = os.path.basename(pdf)
        except Exception as e: datos["error_pdf"] = str(getattr(e, "excepinfo", None) or e)[:200]
        wb.Close(False)
    finally:
        ev.set()
        try: x.Quit()
        except Exception: pass
    return datos


# -------------------------------------------------------------------------------------- PowerPoint
def forma_ppt(sh, prof=0):
    f = {"nombre": sh.Name, "tipo": sh.Type, "x": sh.Left, "y": sh.Top, "w": sh.Width, "h": sh.Height}
    for k, a in [("giro", "Rotation"), ("volteo_h", "HorizontalFlip"), ("volteo_v", "VerticalFlip"), ("visible", "Visible")]:
        try: f[k] = getattr(sh, a)
        except Exception: pass
    try:
        if sh.Type == 1: f["autoforma"] = sh.AutoShapeType
    except Exception: pass
    try:
        if sh.Fill.Visible: f["relleno"] = {"tipo": sh.Fill.Type, "color": color_de(sh.Fill.ForeColor.RGB), "transp": round(sh.Fill.Transparency, 3)}
    except Exception: pass
    try:
        if sh.Line.Visible: f["linea"] = {"color": color_de(sh.Line.ForeColor.RGB), "grosor": sh.Line.Weight, "guion": sh.Line.DashStyle}
    except Exception: pass
    try:
        if sh.HasTextFrame and sh.TextFrame.HasText:
            tr = sh.TextFrame.TextRange
            f["texto"] = tr.Text
            lineas = []
            for k in range(1, 300):
                ln = tr.Lines(k, 1)
                if ln.Length == 0 and k > 1: break
                lineas.append({"texto": ln.Text, "x": ln.BoundLeft, "y": ln.BoundTop, "w": ln.BoundWidth, "h": ln.BoundHeight})
                if ln.Start + ln.Length > tr.Length: break
            f["lineas"] = lineas
            r0 = tr.Characters(1, 1).Font
            f["letra"] = {"fuente": r0.Name, "tam": r0.Size, "b": bool(r0.Bold), "i": bool(r0.Italic), "color": color_de(r0.Color.RGB)}
            f["autoajuste"] = sh.TextFrame2.AutoSize
    except Exception as e: f["error_texto"] = str(e)[:100]
    try:
        if sh.HasTable:
            t = sh.Table
            f["tabla"] = [[t.Cell(i, j).Shape.TextFrame.TextRange.Text for j in range(1, t.Columns.Count + 1)] for i in range(1, t.Rows.Count + 1)]
    except Exception: pass
    try:
        if sh.HasChart: f["grafica"] = sh.Chart.ChartType
    except Exception: pass
    try:
        if sh.HasSmartArt: f["smartart"] = [sh.SmartArt.AllNodes(k).TextFrame2.TextRange.Text for k in range(1, sh.SmartArt.AllNodes.Count + 1)]
    except Exception: pass
    try:
        if sh.Type == 6 and prof < 6: f["grupo"] = [forma_ppt(sh.GroupItems(k), prof + 1) for k in range(1, sh.GroupItems.Count + 1)]
    except Exception: pass
    return f


def verdad_powerpoint(ruta, nombre):
    antes = pids("POWERPNT.EXE")
    p = powerpoint()
    datos = {"archivo": os.path.basename(ruta), "app": "powerpoint", "abrio": False, "diapositivas": []}
    ev = vigilar(600, "POWERPNT.EXE", nombre, nuevo("POWERPNT.EXE", antes))
    try:
        try: p.Visible = True
        except Exception: pass
        t0 = time.time()
        try: pr = p.Presentations.Open(ruta, True, False, True)
        except Exception as e:
            datos["error"] = str(getattr(e, "excepinfo", None) or e)[:400]; return datos
        try: p.WindowState = 1; p.Left = -16000; p.Top = 0
        except Exception: pass
        datos["abrio"] = True; datos["ms_abrir"] = int((time.time() - t0) * 1000)
        W_, H_ = pr.PageSetup.SlideWidth, pr.PageSetup.SlideHeight
        datos["ancho"], datos["alto"] = W_, H_
        esc = 1920 / W_ if W_ >= H_ else 1920 / H_
        for sl in pr.Slides:
            d = {"n": sl.SlideIndex, "diseno": sl.CustomLayout.Name, "oculta": bool(sl.SlideShowTransition.Hidden),
                 "transicion": sl.SlideShowTransition.EntryEffect, "efectos": sl.TimeLine.MainSequence.Count}
            try: d["notas"] = sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            except Exception: pass
            d["formas"] = [forma_ppt(sl.Shapes(k)) for k in range(1, sl.Shapes.Count + 1)]
            png = os.path.join(SALE, f"{nombre}-{sl.SlideIndex:02d}.png")
            try: sl.Export(png, "PNG", int(round(W_ * esc)), int(round(H_ * esc))); d["png"] = os.path.basename(png)
            except Exception as e: d["error_png"] = str(e)[:120]
            datos["diapositivas"].append(d)
        pr.Close()
    finally:
        ev.set()
        try:
            if p.Presentations.Count == 0: p.Quit()
        except Exception: pass
    return datos


# ----------------------------------------------------------------------------------- principal
def una(ruta):
    ruta = os.path.abspath(ruta)
    nombre, ext = os.path.splitext(os.path.basename(ruta))
    print("==", nombre + ext, flush=True)
    t0 = time.time()
    try:
        if ext.lower() in (".docx", ".docm", ".dotx", ".doc"): datos = verdad_word(ruta, nombre)
        elif ext.lower() in (".xlsx", ".xlsm", ".xltx", ".xls"): datos = verdad_excel(ruta, nombre)
        elif ext.lower() in (".pptx", ".pptm", ".potx", ".ppt"): datos = verdad_powerpoint(ruta, nombre)
        else: print("  no es de Office"); return
    except Exception:
        traceback.print_exc(); datos = {"archivo": nombre + ext, "error": traceback.format_exc()[-400:]}
    datos["ms"] = int((time.time() - t0) * 1000)
    guardar_json(os.path.join(SALE, nombre + ".verdad.json"), datos)
    print("  ->", "abrió" if datos.get("abrio") else "NO ABRIÓ " + str(datos.get("error", ""))[:200], datos.get("ms"), "ms", flush=True)


if __name__ == "__main__":
    pythoncom.CoInitialize()
    for a in sys.argv[1:]:
        for r in (glob.glob(a) or glob.glob(os.path.join(BASE, a))):
            una(r)

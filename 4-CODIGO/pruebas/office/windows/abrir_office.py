"""Abre con Office de verdad cada archivo exportado por NOM y anota si Office pidió repararlo.

  python abrir_office.py CARPETA        (los .docx, .xlsx y .pptx de esa carpeta, relativa a nomoffice\\)

Office se abre a la vista (fuera de la pantalla) y CON sus avisos: si al abrir aparece cualquier ventana de diálogo
(«Word encontró contenido que no se puede leer…», «Encontramos un problema con el contenido…», «PowerPoint encontró
un problema…»), se anota su texto y se cierra con «No» (sin reparar). Deja CARPETA\\resultado.json con lo de cada
archivo: si abrió, los avisos que dio, el error, y lo que Office ve dentro (páginas, hojas, diapositivas, fórmulas).
Solo se cierra lo que abrimos (PowerPoint es una sola copia para todo el equipo: no se cierra con Quit).
"""
import ctypes, glob, json, os, sys, threading, time
import win32con, win32gui, win32process
from comun import *

CARPETA = os.path.join(BASE, sys.argv[1]) if len(sys.argv) > 1 else os.path.join(BASE, "entrada", "vuelta")
CLASES_DIALOGO = ("#32770", "NUIDialog", "bosa_sdm_msword", "bosa_sdm_XL9", "bosa_sdm_Mso96")


def textos_de(hwnd):
    t = []
    def hijo(h, _):
        try:
            s = win32gui.GetWindowText(h)
            if s and s.strip(): t.append(s.strip())
        except Exception: pass
        return True
    try: win32gui.EnumChildWindows(hwnd, hijo, None)
    except Exception: pass
    return t


class Vigia:
    """Mientras Office abre un archivo, busca ventanas de diálogo de ESE proceso, anota su texto y las cierra con No."""
    def __init__(self, pid):
        self.pid, self.vistos, self.activo = pid, [], True
        self.hilo = threading.Thread(target=self.correr, daemon=True); self.hilo.start()

    def correr(self):
        ya = set()
        while self.activo:
            def cada(h, _):
                try:
                    if not win32gui.IsWindowVisible(h): return True
                    _, pid = win32process.GetWindowThreadProcessId(h)
                    if pid != self.pid: return True
                    cls = win32gui.GetClassName(h)
                    if cls in CLASES_DIALOGO and h not in ya:
                        ya.add(h)
                        self.vistos.append({"clase": cls, "titulo": win32gui.GetWindowText(h), "texto": textos_de(h)[:12]})
                        # «No» (no reparar), o cerrar
                        for boton in (win32con.IDNO, win32con.IDCANCEL):
                            try: win32gui.PostMessage(h, win32con.WM_COMMAND, boton, 0)
                            except Exception: pass
                        try: win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                        except Exception: pass
                except Exception: pass
                return True
            try: win32gui.EnumWindows(cada, None)
            except Exception: pass
            time.sleep(0.25)

    def fin(self):
        self.activo = False
        time.sleep(0.3)
        return self.vistos


def pid_de(app_hwnd):
    try: return win32process.GetWindowThreadProcessId(app_hwnd)[1]
    except Exception: return None


OCUPADO = ("-2147418111", "-2147417846", "rejected by callee", "Retry later")


def reintentar(f, veces=8):
    """Office a veces contesta «ocupado» (RPC_E_CALL_REJECTED) justo después de cerrar el archivo anterior: no es del
    archivo, se espera un momento y se vuelve a pedir."""
    for k in range(veces):
        try: return f()
        except Exception as e:
            if k == veces - 1 or not any(x in str(e) for x in OCUPADO): raise
            time.sleep(0.5 + k * 0.5)


def abrir_word(w, ruta):
    r = {"archivo": os.path.basename(ruta), "app": "word"}
    v = Vigia(WORD_PID[0])
    t0 = time.time()
    try:
        d = reintentar(lambda: w.Documents.Open(ruta, False, True, False, "", "", False, "", "", 0, 0, True, False, 0, True))
        r["abrio"] = True
        try: r["paginas"] = reintentar(lambda: d.ComputeStatistics(2)); r["parrafos"] = d.Paragraphs.Count; r["secciones"] = d.Sections.Count; r["formas"] = d.Shapes.Count; r["tablas"] = d.Tables.Count
        except Exception as e: r["error_lectura"] = str(e)[:200]
        d.Close(0)
    except Exception as e:
        r["abrio"] = False; r["error"] = str(getattr(e, "excepinfo", None) or e)[:400]
    r["ms"] = int((time.time() - t0) * 1000)
    r["avisos"] = v.fin()
    return r


def abrir_excel(x, ruta):
    r = {"archivo": os.path.basename(ruta), "app": "excel"}
    v = Vigia(EXCEL_PID[0])
    t0 = time.time()
    temp = os.environ.get("TEMP", "")
    antes = set(glob.glob(os.path.join(temp, "error*.xml")))
    try:
        wb = reintentar(lambda: x.Workbooks.Open(ruta, 0, True))
        r["abrio"] = True
        try:
            r["titulo"] = x.ActiveWindow.Caption
            r["reparado"] = "[Repaired]" in str(x.ActiveWindow.Caption) or "[Reparado]" in str(x.ActiveWindow.Caption)
            r["hojas"] = wb.Worksheets.Count
            # fórmulas que Excel no entendió y matrices que derraman (lo nuevo de NOM)
            nombre_mal = derraman = formulas = 0
            for ws in wb.Worksheets:
                try: ur = ws.UsedRange
                except Exception: continue
                try:
                    celdas = ur.SpecialCells(-4123)   # xlCellTypeFormulas
                except Exception: continue
                for c in celdas:
                    formulas += 1
                    try:
                        t = str(c.Text)
                        if t in ("#NAME?", "#¿NOMBRE?"): nombre_mal += 1
                        if c.HasSpill: derraman += 1
                    except Exception: pass
                    if formulas > 4000: break
            r["formulas"], r["nombre_mal"], r["derraman"] = formulas, nombre_mal, derraman
        except Exception as e: r["error_lectura"] = str(e)[:200]
        wb.Close(False)
    except Exception as e:
        r["abrio"] = False; r["error"] = str(getattr(e, "excepinfo", None) or e)[:400]
    r["ms"] = int((time.time() - t0) * 1000)
    r["avisos"] = v.fin()
    nuevos = set(glob.glob(os.path.join(temp, "error*.xml"))) - antes
    if nuevos:
        r["reparado"] = True
        try: r["registro_reparacion"] = open(sorted(nuevos)[-1], encoding="utf-8", errors="replace").read()[:1500]
        except Exception: pass
    return r


def abrir_powerpoint(p, ruta):
    r = {"archivo": os.path.basename(ruta), "app": "powerpoint"}
    v = Vigia(PPT_PID[0])
    t0 = time.time()
    try:
        pr = reintentar(lambda: p.Presentations.Open(ruta, True, False, True))
        r["abrio"] = True
        try: r["diapositivas"] = pr.Slides.Count; r["formas"] = sum(s.Shapes.Count for s in pr.Slides)
        except Exception as e: r["error_lectura"] = str(e)[:200]
        pr.Close()
    except Exception as e:
        r["abrio"] = False; r["error"] = str(getattr(e, "excepinfo", None) or e)[:400]
    r["ms"] = int((time.time() - t0) * 1000)
    r["avisos"] = v.fin()
    return r


WORD_PID, EXCEL_PID, PPT_PID = [None], [None], [None]


def pid_proceso(imagen, antes):
    salida = os.popen(f'tasklist /FI "IMAGENAME eq {imagen}" /FO CSV /NH').read()
    ahora = {int(c.split('","')[1]) for c in salida.splitlines() if c.startswith('"') and '","' in c and c.split('","')[1].isdigit()}
    n = ahora - antes
    return next(iter(n)) if len(n) == 1 else (next(iter(ahora)) if len(ahora) == 1 else None)


def pids(imagen):
    salida = os.popen(f'tasklist /FI "IMAGENAME eq {imagen}" /FO CSV /NH').read()
    return {int(c.split('","')[1]) for c in salida.splitlines() if c.startswith('"') and '","' in c and c.split('","')[1].isdigit()}


if __name__ == "__main__":
    pythoncom.CoInitialize()
    archivos = sorted(f for f in glob.glob(os.path.join(CARPETA, "*")) if f.lower().endswith((".docx", ".xlsx", ".pptx")))
    res = []
    porTipo = {".docx": [], ".xlsx": [], ".pptx": []}
    for f in archivos: porTipo[os.path.splitext(f)[1].lower()].append(f)
    if porTipo[".docx"]:
        a = pids("WINWORD.EXE"); w = word(); WORD_PID[0] = pid_proceso("WINWORD.EXE", a)
        w.Visible = True; w.DisplayAlerts = -1   # (con avisos: así se ve si pide reparar)
        try: w.WindowState = 0; w.Left = -3000; w.Top = 0
        except Exception: pass
        try:
            for f in porTipo[".docx"]:
                r = abrir_word(w, f); res.append(r); print(" ", r["archivo"], "abrió" if r.get("abrio") else "NO ABRIÓ", "· avisos:", len(r["avisos"]), flush=True)
        finally:
            try: w.Quit(0)
            except Exception: pass
    if porTipo[".xlsx"]:
        a = pids("EXCEL.EXE"); x = excel(); EXCEL_PID[0] = pid_proceso("EXCEL.EXE", a)
        x.Visible = True; x.DisplayAlerts = True; x.ScreenUpdating = True
        try: x.WindowState = -4143; x.Left = -16000; x.Top = 0
        except Exception: pass
        try:
            for f in porTipo[".xlsx"]:
                r = abrir_excel(x, f); res.append(r); print(" ", r["archivo"], "abrió" if r.get("abrio") else "NO ABRIÓ", "· avisos:", len(r["avisos"]), "· reparado" if r.get("reparado") else "", flush=True)
        finally:
            try: x.Quit()
            except Exception: pass
    if porTipo[".pptx"]:
        a = pids("POWERPNT.EXE"); p = powerpoint(); PPT_PID[0] = pid_proceso("POWERPNT.EXE", a)
        try: p.Visible = True; p.DisplayAlerts = 2   # ppAlertsAll
        except Exception: pass
        for f in porTipo[".pptx"]:
            r = abrir_powerpoint(p, f); res.append(r); print(" ", r["archivo"], "abrió" if r.get("abrio") else "NO ABRIÓ", "· avisos:", len(r["avisos"]), flush=True)
            try: p.WindowState = 1; p.Left = -16000
            except Exception: pass
    guardar_json(os.path.join(CARPETA, "resultado.json"), res)
    malos = [r for r in res if not r.get("abrio") or r.get("avisos") or r.get("reparado")]
    print(f"{len(res)} archivos: {len(res) - len(malos)} abrieron sin avisos ni reparar, {len(malos)} con algo", flush=True)

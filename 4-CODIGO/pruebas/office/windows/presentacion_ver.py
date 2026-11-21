"""Abre un .pptx en PowerPoint y dice lo de la pestaña Presentación: qué diapositivas están ocultas, sus intervalos,
las secciones, las presentaciones personalizadas y la configuración de la presentación.
Uso: presentacion_ver.py ENTRADA.pptx SALIDA.json"""
import json, os, sys
from comun import powerpoint

entrada, salida = (os.path.abspath(x) for x in sys.argv[1:3])
pp = powerpoint()
datos = {"abrio": False}
try:
    pr = pp.Presentations.Open(entrada, True, False, False)   # solo lectura, sin ventana
    try:
        datos["abrio"] = True
        ids = {}
        datos["diapositivas"] = []
        for s in pr.Slides:
            tr = s.SlideShowTransition
            ids[s.SlideID] = s.SlideIndex
            titulo = ""
            try: titulo = s.Shapes.Title.TextFrame.TextRange.Text
            except Exception: pass
            datos["diapositivas"].append({"n": s.SlideIndex, "titulo": titulo, "oculta": bool(tr.Hidden),
                                          "porTiempo": bool(tr.AdvanceOnTime), "segundos": round(tr.AdvanceTime, 2), "conClic": bool(tr.AdvanceOnClick)})
        sp = pr.SectionProperties
        datos["secciones"] = [{"nombre": sp.Name(i), "primera": sp.FirstSlide(i), "cuantas": sp.SlidesCount(i)} for i in range(1, sp.Count + 1)]
        ss = pr.SlideShowSettings
        datos["personalizadas"] = []
        for i in range(1, ss.NamedSlideShows.Count + 1):
            n = ss.NamedSlideShows(i)
            datos["personalizadas"].append({"nombre": n.Name, "diapositivas": [ids.get(x) for x in (n.SlideIDs or []) if x]})
        datos["configuracion"] = {"repetir": bool(ss.LoopUntilStopped), "tipo": ss.ShowType, "rango": ss.RangeType,
                                  "desde": ss.StartingSlide, "hasta": ss.EndingSlide, "avance": ss.AdvanceMode,
                                  "conNarracion": bool(ss.ShowWithNarration), "conAnimacion": bool(ss.ShowWithAnimation)}
        try: datos["configuracion"]["personalizada"] = ss.SlideShowName
        except Exception: pass
    finally:
        pr.Close()
except Exception as e:
    datos["error"] = str(e)[:300]
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

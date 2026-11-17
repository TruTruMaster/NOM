"""Animaciones y transiciones de PowerPoint hechas por PowerPoint de verdad, y su video para medir.

  python animaciones_office.py constantes            -> archivos\\a00-constantes.json (efectos, direcciones, transiciones)
  python animaciones_office.py catalogo              -> archivos\\a01-animaciones.pptx (+ .json: qué efecto va en cada diapositiva)
  python animaciones_office.py opciones              -> archivos\\a02-opciones.pptx (por letra, palabra y párrafo, repetir…)
  python animaciones_office.py transiciones          -> archivos\\a03-transiciones.pptx (cada transición, con sus variantes)
  python animaciones_office.py video NOMBRE [ALTO] [CUADROS]  -> archivos\\NOMBRE.mp4 (CreateVideo; por defecto 720 y 30)

Cada diapositiva lleva abajo a la izquierda una «marca» de 10 cuadritos negros o blancos con su número en binario:
así, en el video se sabe qué diapositiva se ve en cada cuadro y en qué momento empezó (para medir contra NOM).
En el catálogo cada efecto va «después de la anterior» (empieza solo al llegar a la diapositiva) y la diapositiva
avanza sola cuando terminan sus animaciones (más un segundo), para que el video sea igual cada vez.

PowerPoint es una sola copia en el equipo: nunca se cierra con Quit; solo se cierran las presentaciones abiertas aquí.
"""
import json, os, sys, time
from comun import *
import win32com.client as W
from win32com.client import constants as K

MARCA_BITS = 12   # 9 del número de la diapositiva y 3 de comprobación (número módulo 7)


def poner_marcas(pr):
    """La marca de cada diapositiva con su número final (después de borrar las que no sirvieron)."""
    for k in range(1, pr.Slides.Count + 1):
        sl = pr.Slides(k)
        v = (k << 3) | (k % 7)
        for i in range(MARCA_BITS):
            bit = (v >> (MARCA_BITS - 1 - i)) & 1
            q = sl.Shapes.AddShape(1, 6 + i * 12, 528, 10, 10)
            q.Line.Visible = 0
            q.Fill.ForeColor.RGB = rgb(0, 0, 0) if bit else rgb(255, 255, 255)
            q.Name = f"marca{i}"


def constantes():
    p = powerpoint()
    todas = {}
    for d in W.constants.__dicts__:
        for k, v in d.items():
            if isinstance(v, int) and (k.startswith("msoAnim") or k.startswith("ppEffect") or k.startswith("ppTransition")
                                       or k.startswith("msoAnimTrigger") or k.startswith("ppAnimate")):
                todas[k] = v
    guardar_json(os.path.join(ARCHIVOS, "a00-constantes.json"), dict(sorted(todas.items(), key=lambda kv: (kv[0].split("Effect")[0], kv[1]))))
    print(len(todas), "constantes")


class Pres:
    def __init__(self, nombre):
        self.p = powerpoint()
        self.nombre = nombre
        self.pr = self.p.Presentations.Add(WithWindow=True)
        try: self.p.WindowState = 1; self.p.Left = -16000; self.p.Top = 0
        except Exception: pass
        self.pr.PageSetup.SlideWidth = 960; self.pr.PageSetup.SlideHeight = 540
        self.b = Bitacora(nombre)
        self.lista = []

    def diapo(self):
        m = self.pr.SlideMaster
        dis = None
        for k in range(1, m.CustomLayouts.Count + 1):
            if m.CustomLayouts(k).Name == "Blank": dis = m.CustomLayouts(k)
        sl = self.pr.Slides.AddSlide(self.pr.Slides.Count + 1, dis or m.CustomLayouts(7))
        return sl

    def guardar(self, extra=None):
        ruta = os.path.join(ARCHIVOS, self.nombre + ".pptx")
        poner_marcas(self.pr)
        # el número final de cada diapositiva de la lista (las borradas corrieron a las demás)
        for x in self.lista:
            x["diapo"] = x.pop("_sl").SlideIndex if "_sl" in x else x["diapo"]
        self.pr.SaveAs(ruta, 24)
        datos = self.b.resumen(); datos["diapositivas"] = self.lista
        if extra: datos.update(extra)
        guardar_json(os.path.join(ARCHIVOS, self.nombre + ".json"), datos)
        self.pr.Close()
        return ruta


def rect(sl, x=380, y=200, w=200, h=140, color=(47, 85, 151)):
    s = sl.Shapes.AddShape(1, x, y, w, h)
    s.Line.Visible = 0
    s.Fill.ForeColor.RGB = rgb(*color)
    s.Name = "objetivo"
    return s


def caja_texto(sl, t="Animación NOM", x=230, y=200, w=500, h=120, tam=54):
    s = sl.Shapes.AddTextbox(1, x, y, w, h)
    tr = s.TextFrame2.TextRange; tr.Text = t; tr.Font.Size = tam; tr.Font.Name = "Arial"; tr.Font.Bold = True
    tr.Font.Fill.ForeColor.RGB = rgb(0, 0, 0)
    tr.ParagraphFormat.Alignment = 2
    s.TextFrame2.WordWrap = True
    s.Name = "objetivo"
    return s


def duracion_total(sl):
    """Cuánto duran las animaciones de la diapositiva (s), según PowerPoint."""
    fin = 0.0
    seq = sl.TimeLine.MainSequence
    t = 0.0
    for i in range(1, seq.Count + 1):
        e = seq.Item(i)
        rep = e.Timing.RepeatCount or 1
        if rep < 1 or rep > 100: rep = 1
        d = e.Timing.Duration * rep * (2 if e.Timing.AutoReverse else 1)
        t0 = t + e.Timing.TriggerDelayTime
        if not (d < 1e6): d = 5.0   # «hasta el final de la diapositiva»
        fin = max(fin, t0 + d)
        t = t0 + d
    return fin


TEXTO_SOLO = ("ChangeFont", "ChangeFontColor", "ChangeFontSize", "ChangeFontStyle", "BoldFlash", "BoldReveal", "BrushOnUnderline",
              "Wave", "ColorWave", "Credits", "Typewriter", "Flicker", "StyleEmphasis", "Underline", "Blast", "Shimmer", "Teeter",
              "BrushOnColor", "ColorBlend", "GrowWithColor", "VerticalGrow")


def catalogo():
    c = json.load(open(os.path.join(ARCHIVOS, "a00-constantes.json"), encoding="utf8"))
    efectos = sorted([(k, v) for k, v in c.items() if k.startswith("msoAnimEffect") and v > 0 and not k.startswith(("msoAnimEffectAfter", "msoAnimEffectRestart", "msoAnimEffect3D")) and k not in ("msoAnimEffectCustom", "msoAnimEffectRandomEffects", "msoAnimEffectMediaPlay", "msoAnimEffectMediaPause", "msoAnimEffectMediaStop", "msoAnimEffectMediaPlayFromBookmark", "msoAnimEffectChangeFont")], key=lambda kv: kv[1])
    P = Pres("a01-animaciones")
    for nombre, v in efectos:
        corto = nombre[len("msoAnimEffect"):]
        for salida in (False, True):
            sl = P.diapo()
            texto = any(corto.startswith(t) for t in TEXTO_SOLO)
            s = caja_texto(sl) if texto else rect(sl)
            try:
                e = sl.TimeLine.MainSequence.AddEffect(s, v, 0, 3)
            except Exception as ex:
                if texto: sl.Delete(); break
                s.Delete(); s = caja_texto(sl); texto = True
                try: e = sl.TimeLine.MainSequence.AddEffect(s, v, 0, 3)
                except Exception as ex2:
                    P.b.mal.append((corto, str(ex2)[:200])); sl.Delete(); break
            if salida:
                try:
                    if e.EffectType == v and e.Exit == 0:
                        e.Exit = -1
                        if e.Exit == 0: raise RuntimeError("no tiene salida")
                    else: raise RuntimeError("no tiene salida")
                except Exception:
                    sl.Delete(); break
            dur = duracion_total(sl)
            sl.SlideShowTransition.AdvanceOnTime = -1
            sl.SlideShowTransition.AdvanceTime = min(60.0, round(dur + 1.0, 2))
            P.lista.append({"_sl": sl, "diapo": sl.SlideIndex, "efecto": corto, "valor": v, "salida": salida, "texto": texto,
                            "duracion": e.Timing.Duration, "dur_total": dur, "clase": e.EffectType})
            P.b.ok.append(corto + (" (salida)" if salida else ""))
    print("diapositivas:", P.pr.Slides.Count, flush=True)
    P.guardar()


def opciones():
    """Variantes que cambian el comportamiento: por letra o palabra, por párrafo, repetir, rebobinar, autorreversa,
    inicio y fin suaves, rebote al final, atenuar después, ocultar después, retrasos y desencadenadores."""
    P = Pres("a02-opciones")
    def nueva(nombre, fn, texto=True, t="Animación NOM"):
        sl = P.diapo()
        s = caja_texto(sl, t) if texto else rect(sl)
        try:
            fn(sl, s)
            dur = duracion_total(sl)
            sl.SlideShowTransition.AdvanceOnTime = -1
            sl.SlideShowTransition.AdvanceTime = min(60.0, round(dur + 1.0, 2))
            P.lista.append({"_sl": sl, "diapo": sl.SlideIndex, "efecto": nombre, "dur_total": dur})
            P.b.ok.append(nombre)
        except Exception as ex:
            P.b.mal.append((nombre, str(ex)[:200])); sl.Delete()
    seqd = lambda sl: sl.TimeLine.MainSequence
    def por_unidad(ef, unidad, reversa=False):
        def f(sl, s):
            q = seqd(sl)
            e = q.AddEffect(s, ef, 0, 3)
            e = q.ConvertToTextUnitEffect(e, unidad)
            if reversa: q.ConvertToAnimateInReverse(e, -1)
        return f
    for ef, nom in ((K.msoAnimEffectFly, "Fly"), (K.msoAnimEffectFade, "Fade"), (K.msoAnimEffectAppear, "Appear"), (K.msoAnimEffectZoom, "Zoom"),
                    (K.msoAnimEffectWipe, "Wipe"), (K.msoAnimEffectBounce, "Bounce"), (K.msoAnimEffectFloat, "Float"), (K.msoAnimEffectSwivel, "Swivel")):
        nueva(nom + " por letra", por_unidad(ef, 1))
        nueva(nom + " por palabra", por_unidad(ef, 2))
    for ef, nom in ((K.msoAnimEffectChangeFontColor, "ChangeFontColor"), (K.msoAnimEffectGrowShrink, "GrowShrink"), (K.msoAnimEffectWave, "Wave"),
                    (K.msoAnimEffectBoldFlash, "BoldFlash"), (K.msoAnimEffectSpin, "Spin")):
        nueva(nom + " por letra", por_unidad(ef, 1))
    # por párrafo (viñetas una por una, automáticas)
    def parrafos(ef, nivel=2, invertido=False, fondo=False):
        def f(sl, s):
            s.TextFrame2.TextRange.Text = "Primera idea\rSegunda idea\rTercera idea"
            s.TextFrame2.TextRange.Font.Size = 36
            if fondo: s.Fill.Visible = -1; s.Fill.ForeColor.RGB = rgb(255, 230, 153)
            e = seqd(sl).AddEffect(s, ef, nivel, 3)
            for i in range(1, seqd(sl).Count + 1):
                seqd(sl).Item(i).Timing.TriggerType = 3
            if invertido: seqd(sl).ConvertToAnimateInReverse(seqd(sl).Item(1), -1)
            if fondo: seqd(sl).ConvertToAnimateBackground(seqd(sl).Item(1), -1)
        return f
    nueva("Fly por párrafo", parrafos(K.msoAnimEffectFly))
    nueva("Fade por párrafo con fondo", parrafos(K.msoAnimEffectFade, fondo=True))
    nueva("Wipe por párrafo", parrafos(K.msoAnimEffectWipe))
    nueva("Fly por párrafo al revés", parrafos(K.msoAnimEffectFly, invertido=True))
    # tiempos
    def tiempo(ef, **kw):
        def f(sl, s):
            e = seqd(sl).AddEffect(s, ef, 0, 3)
            for k, v in kw.items(): setattr(e.Timing, k, v)
        return f
    nueva("Spin repetir 3", tiempo(K.msoAnimEffectSpin, RepeatCount=3), texto=False)
    nueva("Spin autorreversa", tiempo(K.msoAnimEffectSpin, AutoReverse=-1), texto=False)
    nueva("GrowShrink rebobinar", tiempo(K.msoAnimEffectGrowShrink, RewindAtEnd=-1), texto=False)
    nueva("Fly inicio y fin suaves", tiempo(K.msoAnimEffectFly, Accelerate=0.5, Decelerate=0.5), texto=False)
    nueva("Fly rebote al final", tiempo(K.msoAnimEffectFly, BounceEnd=-1, BounceEndIntensity=0.5), texto=False)
    nueva("Fade velocidad 2s retraso 0.5s", tiempo(K.msoAnimEffectFade, Duration=2.0, TriggerDelayTime=0.5), texto=False)
    nueva("Teeter repetir 2", tiempo(K.msoAnimEffectTeeter, RepeatCount=2), texto=False)
    # después de la animación: atenuar a un color, ocultar, ocultar al siguiente clic
    def despues(tipo):
        def f(sl, s):
            q = seqd(sl)
            e = q.AddEffect(s, K.msoAnimEffectFade, 0, 3)
            q.ConvertToAfterEffect(e, tipo, rgb(200, 200, 200))
            q.AddEffect(rect(sl, 60, 60, 120, 80, (192, 0, 0)), K.msoAnimEffectFade, 0, 3)
        return f
    nueva("Atenuar después", despues(1))
    nueva("Ocultar después", despues(2))
    nueva("Por palabra al revés", por_unidad(K.msoAnimEffectFly, 2, True))
    # Parámetros de énfasis
    def param(ef, **kw):
        def f(sl, s):
            e = seqd(sl).AddEffect(s, ef, 0, 3)
            for k, v in kw.items():
                if k == "color": e.EffectParameters.Color2.RGB = v
                else: setattr(e.EffectParameters, k, v)
        return f
    nueva("GrowShrink 300 %", param(K.msoAnimEffectGrowShrink, Size=300), texto=False)
    nueva("GrowShrink 50 %", param(K.msoAnimEffectGrowShrink, Size=50), texto=False)
    nueva("Spin 90 grados", param(K.msoAnimEffectSpin, Amount=90), texto=False)
    nueva("Spin -720 grados", param(K.msoAnimEffectSpin, Amount=-720), texto=False)
    nueva("Transparency 25 %", param(K.msoAnimEffectTransparency, Amount=0.25), texto=False)
    nueva("ChangeFillColor rojo", param(K.msoAnimEffectChangeFillColor, color=rgb(255, 0, 0)), texto=False)
    nueva("ChangeLineColor verde", param(K.msoAnimEffectChangeLineColor, color=rgb(0, 176, 80)), texto=False)
    nueva("ChangeFontColor verde", param(K.msoAnimEffectChangeFontColor, color=rgb(0, 176, 80)))
    nueva("ChangeFontSize 200 %", param(K.msoAnimEffectChangeFontSize, Size=200))
    # direcciones
    def dire(ef, d):
        def f(sl, s):
            e = seqd(sl).AddEffect(s, ef, 0, 3); e.EffectParameters.Direction = d
        return f
    for nom, ef in (("Fly", K.msoAnimEffectFly), ("Wipe", K.msoAnimEffectWipe), ("Peek", K.msoAnimEffectPeek), ("Split", K.msoAnimEffectSplit),
                    ("Blinds", K.msoAnimEffectBlinds), ("Box", K.msoAnimEffectBox), ("Circle", K.msoAnimEffectCircle), ("Diamond", K.msoAnimEffectDiamond),
                    ("Plus", K.msoAnimEffectPlus), ("RandomBars", K.msoAnimEffectRandomBars), ("Strips", K.msoAnimEffectStrips), ("Wheel", K.msoAnimEffectWheel),
                    ("Zoom", K.msoAnimEffectZoom), ("Stretch", K.msoAnimEffectStretch), ("Checkerboard", K.msoAnimEffectCheckerboard)):
        for dn in ("Left", "Right", "Up", "Down", "UpLeft", "DownRight", "In", "Out", "Horizontal", "Vertical", "HorizontalIn", "VerticalOut", "Across"):
            d = getattr(K, "msoAnimDirection" + dn, None)
            if d is None: continue
            sl = P.diapo(); s = rect(sl)
            try:
                e = seqd(sl).AddEffect(s, ef, 0, 3); e.EffectParameters.Direction = d
                if e.EffectParameters.Direction != d: raise RuntimeError("no aceptó la dirección")
                dur = duracion_total(sl)
                sl.SlideShowTransition.AdvanceOnTime = -1; sl.SlideShowTransition.AdvanceTime = min(60.0, round(dur + 1.0, 2))
                P.lista.append({"_sl": sl, "diapo": sl.SlideIndex, "efecto": f"{nom} {dn}", "dur_total": dur})
            except Exception:
                sl.Delete()
    # varios efectos en una diapositiva: con la anterior, después de la anterior y retrasos
    def varios(sl, s):
        s.Delete()
        a = rect(sl, 80, 120, 160, 110, (192, 0, 0)); b = rect(sl, 400, 120, 160, 110, (0, 112, 192)); c = rect(sl, 720, 120, 160, 110, (0, 176, 80))
        q = seqd(sl)
        e1 = q.AddEffect(a, K.msoAnimEffectFly, 0, 3)
        e2 = q.AddEffect(b, K.msoAnimEffectFade, 0, 2); e2.Timing.TriggerDelayTime = 0.25
        e3 = q.AddEffect(c, K.msoAnimEffectZoom, 0, 3); e3.Timing.TriggerDelayTime = 0.5
        e4 = q.AddEffect(a, K.msoAnimEffectSpin, 0, 2)
        e5 = q.AddEffect(b, K.msoAnimEffectFade, 0, 3); e5.Exit = -1
    nueva("Con y después de la anterior", varios, texto=False)
    # trayectoria propia
    def ruta(sl, s):
        e = seqd(sl).AddEffect(s, K.msoAnimEffectCustom, 0, 3)
        b = e.Behaviors.Add(1)  # msoAnimTypeMotion
        b.MotionEffect.Path = "M 0 0 C 0.1 -0.3 0.3 -0.3 0.4 0 C 0.5 0.3 0.2 0.3 0 0 Z"
        e.Timing.Duration = 2
    nueva("Trayectoria propia con curvas", ruta, texto=False)
    print("diapositivas:", P.pr.Slides.Count, flush=True)
    P.guardar()


def aceleracion():
    """Una trayectoria recta (media diapositiva a la derecha, 2 s) con cada combinación de inicio y fin suaves y rebote
    al final: para medir cómo PowerPoint aplica accel/decel (no es la aceleración constante de SMIL)."""
    P = Pres("a04-aceleracion")
    pares = [(0, 0), (0.1, 0), (0.25, 0), (0.5, 0), (0.75, 0), (1, 0), (0, 0.25), (0, 0.5), (0, 1), (0.1, 0.1), (0.25, 0.25), (0.5, 0.5),
             (0.3, 0.7), (0.2, 0.3), (0.4, 0.4), (0.05, 0.05)]
    for a, d in pares:
        sl = P.diapo(); s = rect(sl, 60, 200, 200, 140)
        e = sl.TimeLine.MainSequence.AddEffect(s, K.msoAnimEffectCustom, 0, 3)
        b = e.Behaviors.Add(1); b.MotionEffect.Path = "M 0 0 L 0.5 0 E"
        e.Timing.Duration = 2.0
        e.Timing.Accelerate = a; e.Timing.Decelerate = d
        sl.SlideShowTransition.AdvanceOnTime = -1; sl.SlideShowTransition.AdvanceTime = 3.0
        P.lista.append({"_sl": sl, "diapo": sl.SlideIndex, "efecto": f"acel {a} desac {d}", "acel": a, "desac": d, "dur_total": 2.0})
    for inten in (0.25, 0.5, 0.75, 1.0):
        sl = P.diapo(); s = rect(sl, 60, 200, 200, 140)
        e = sl.TimeLine.MainSequence.AddEffect(s, K.msoAnimEffectCustom, 0, 3)
        b = e.Behaviors.Add(1); b.MotionEffect.Path = "M 0 0 L 0.5 0 E"
        e.Timing.Duration = 2.0
        try: e.Timing.BounceEnd = -1; e.Timing.BounceEndIntensity = inten
        except Exception as ex: P.b.mal.append(("rebote", str(ex)[:120]))
        sl.SlideShowTransition.AdvanceOnTime = -1; sl.SlideShowTransition.AdvanceTime = 3.0
        P.lista.append({"_sl": sl, "diapo": sl.SlideIndex, "efecto": f"rebote {inten}", "rebote": inten, "dur_total": 2.0})
    # la misma recta con un anim de ppt_x (no animMotion), para ver si es igual
    for a, d in ((0.5, 0.5), (1, 0)):
        sl = P.diapo(); s = rect(sl, 60, 200, 200, 140)
        e = sl.TimeLine.MainSequence.AddEffect(s, K.msoAnimEffectCustom, 0, 3)
        b = e.Behaviors.Add(5)  # msoAnimTypeProperty
        b.PropertyEffect.Property = 1  # msoAnimX
        for tm, v in ((0, "#ppt_x"), (1, "#ppt_x+0.5")):
            pt = b.PropertyEffect.Points.Add(); pt.Time = tm; pt.Value = v
        e.Timing.Duration = 2.0; e.Timing.Accelerate = a; e.Timing.Decelerate = d
        sl.SlideShowTransition.AdvanceOnTime = -1; sl.SlideShowTransition.AdvanceTime = 3.0
        P.lista.append({"_sl": sl, "diapo": sl.SlideIndex, "efecto": f"anim x acel {a} desac {d}", "acel": a, "desac": d, "dur_total": 2.0})
    P.guardar()


def transiciones():
    c = json.load(open(os.path.join(ARCHIVOS, "a00-constantes.json"), encoding="utf8"))
    tr = sorted([(k, v) for k, v in c.items() if k.startswith("ppEffect") and k not in ("ppEffectMixed",)], key=lambda kv: kv[1])
    P = Pres("a03-transiciones")
    colores = [(47, 85, 151), (192, 0, 0), (0, 176, 80), (112, 48, 160), (255, 192, 0)]
    for i, (nombre, v) in enumerate(tr):
        sl = P.diapo()
        # fondo de color liso distinto en cada diapositiva para ver la transición entre una y otra
        sl.FollowMasterBackground = 0
        sl.Background.Fill.Solid(); sl.Background.Fill.ForeColor.RGB = rgb(*colores[i % len(colores)])
        caja_texto(sl, nombre[len("ppEffect"):], 80, 200, 800, 120, 40)
        try:
            st = sl.SlideShowTransition; st.EntryEffect = v
            if st.EntryEffect != v: raise RuntimeError("no aceptó la transición")
            st.Duration = 2.0
            st.AdvanceOnTime = -1; st.AdvanceTime = 1.0
            P.lista.append({"_sl": sl, "diapo": sl.SlideIndex, "transicion": nombre, "valor": v})
            P.b.ok.append(nombre)
        except Exception as ex:
            P.b.mal.append((nombre, str(ex)[:200])); sl.Delete()
    print("diapositivas:", P.pr.Slides.Count, flush=True)
    P.guardar()


def video(nombre, alto="720", cuadros="30"):
    p = powerpoint()
    ruta = os.path.join(ARCHIVOS, nombre + ".pptx")
    pr = p.Presentations.Open(ruta, ReadOnly=True, Untitled=False, WithWindow=False)
    try:
        sal = os.path.join(ARCHIVOS, nombre + ".mp4")
        if os.path.exists(sal): os.remove(sal)
        t0 = time.time()
        pr.CreateVideo(sal, True, 1, int(alto), int(cuadros), 100)
        while True:
            st = pr.CreateVideoStatus
            if st in (3, 4): break  # hecho / falló
            time.sleep(2)
            if time.time() - t0 > 3000: print("TIEMPO AGOTADO", flush=True); break
        print("video", "listo" if st == 3 else "falló", st, round(time.time() - t0), "s", flush=True)
    finally:
        pr.Close()


if __name__ == "__main__":
    a = sys.argv[1:]
    {"constantes": constantes, "catalogo": catalogo, "aceleracion": aceleracion, "opciones": opciones, "transiciones": transiciones, "video": video}[a[0]](*a[1:])

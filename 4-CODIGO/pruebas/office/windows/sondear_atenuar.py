"""Sondeo: qué hace PowerPoint con «Después de la animación: atenuar a un color» según la figura.

  python sondear_atenuar.py      -> archivos\\v03-atenuar.pptx y .json (para animaciones_office.py video)

Cuatro diapositivas, todas «después de la anterior» y atenuando a RGB(200, 200, 200):
  1. rectángulo azul relleno con texto blanco (lo que arma el panel de NOM)
  2. el mismo rectángulo sin texto
  3. cuadro de texto sin relleno con letra negra (como a02-opciones)
  4. el rectángulo con texto y después un segundo efecto en otra figura (¿cuándo atenúa?)
"""
import os, sys
from comun import *
from animaciones_office import Pres, rect, caja_texto, K


def seqd(sl): return sl.TimeLine.MainSequence


def main():
    P = Pres("v03-atenuar")

    def con_texto(sl):
        s = rect(sl)
        tr = s.TextFrame2.TextRange; tr.Text = "NOM"; tr.Font.Size = 12; tr.Font.Bold = True
        tr.Font.Fill.ForeColor.RGB = rgb(255, 255, 255)
        return s

    def una(nombre, hacer_fig, segundo=None, despues=1):
        """segundo: None, o cómo empieza el segundo efecto (3 después de la anterior, 2 con la anterior, 1 al clic)"""
        sl = P.diapo()
        s = hacer_fig(sl)
        q = seqd(sl)
        e = q.AddEffect(s, K.msoAnimEffectFade, 0, 3)
        e.Timing.Duration = 1.0
        q.ConvertToAfterEffect(e, despues, rgb(200, 200, 200))
        dur = 1.0
        if segundo:
            e2 = q.AddEffect(rect(sl, 60, 60, 120, 80, (192, 0, 0)), K.msoAnimEffectFade, 0, segundo)
            e2.Timing.Duration = 0.5
            dur = 1.5 if segundo != 2 else 1.0
        P.lista.append({"_sl": sl, "efecto": nombre, "dur_total": dur + 0.6})

    una("Atenuar: rectángulo con texto", con_texto)
    una("Atenuar: rectángulo sin texto", lambda sl: rect(sl))
    una("Atenuar: cuadro de texto", lambda sl: caja_texto(sl))
    una("Atenuar: segundo después de la anterior", con_texto, segundo=3)
    una("Atenuar: segundo con la anterior", con_texto, segundo=2)
    una("Atenuar: segundo al hacer clic", con_texto, segundo=1)
    una("Ocultar al clic: segundo después", con_texto, segundo=3, despues=3)
    una("Ocultar al clic: segundo al clic", con_texto, segundo=1, despues=3)
    una("Ocultar después: segundo después", con_texto, segundo=3, despues=2)
    print("diapositivas:", P.pr.Slides.Count, flush=True)
    P.guardar()


if __name__ == "__main__":
    main()

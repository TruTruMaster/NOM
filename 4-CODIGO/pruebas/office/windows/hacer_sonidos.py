"""Arma con PowerPoint de verdad una presentación con sonidos (para la vuelta de los paneles, vuelta_paneles.py).

  python hacer_sonidos.py      -> archivos\\v04-sonidos.pptx y .funciones.json

  1. transición «Desvanecer» con un sonido propio (un WAV)
  2. un efecto de entrada con ese sonido
  3. transición con «Detener el sonido anterior»
  4. transición «Barrido» con el sonido en bucle hasta el siguiente sonido
El WAV se genera aquí (tono de 0.8 s): PowerPoint solo acepta WAV para los sonidos de efectos y transiciones.
"""
import math, os, struct, wave
from comun import *
from animaciones_office import Pres, rect, K


def wav(ruta, seg=0.8, hz=440, frec=22050):
    with wave.open(ruta, "wb") as w:
        w.setnchannels(1); w.setsampwidth(2); w.setframerate(frec)
        n = int(seg * frec)
        w.writeframes(b"".join(struct.pack("<h", int(12000 * math.sin(2 * math.pi * hz * i / frec) * min(1, (n - i) / 2000))) for i in range(n)))


def main():
    sonido = os.path.join(ARCHIVOS, "sonido-prueba.wav")
    wav(sonido)
    P = Pres("v04-sonidos")
    b = P.b
    # 1. transición con sonido propio
    sl = P.diapo(); rect(sl)
    with b.paso("transición con sonido propio"):
        t = sl.SlideShowTransition
        t.EntryEffect = K.ppEffectFade
        t.SoundEffect.ImportFromFile(sonido)
    P.lista.append({"_sl": sl, "efecto": "Transición con sonido"})
    # 2. efecto de entrada con sonido
    sl = P.diapo(); s = rect(sl)
    with b.paso("efecto con sonido"):
        e = sl.TimeLine.MainSequence.AddEffect(s, K.msoAnimEffectFade, 0, 1)
        try: e.EffectInformation.SoundEffect.ImportFromFile(sonido)
        except Exception:
            # (modelo viejo de animación: el mismo sonido por AnimationSettings)
            s.AnimationSettings.SoundEffect.ImportFromFile(sonido)
    P.lista.append({"_sl": sl, "efecto": "Efecto con sonido"})
    # 3. detener el sonido anterior
    sl = P.diapo(); rect(sl)
    with b.paso("transición: detener el sonido anterior"):
        t = sl.SlideShowTransition
        t.EntryEffect = K.ppEffectFade
        t.SoundEffect.Type = K.ppSoundStopPrevious
    P.lista.append({"_sl": sl, "efecto": "Detener el sonido anterior"})
    # 4. sonido en bucle hasta el siguiente
    sl = P.diapo(); rect(sl)
    with b.paso("transición con sonido en bucle"):
        t = sl.SlideShowTransition
        t.EntryEffect = K.ppEffectWipeLeft
        t.SoundEffect.ImportFromFile(sonido)
        t.LoopSoundUntilNext = -1
    P.lista.append({"_sl": sl, "efecto": "Sonido en bucle"})
    print("diapositivas:", P.pr.Slides.Count, flush=True)
    P.guardar()


if __name__ == "__main__":
    main()

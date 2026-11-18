"""Qué sonidos ve PowerPoint en un .pptx: por diapositiva, el de su transición y el de cada efecto.

  python leer_sonidos.py RUTA.pptx     (relativa a nomoffice\\ o absoluta)
"""
import json, os, sys
from comun import *

def main():
    p = powerpoint()
    ruta = sys.argv[1] if os.path.isabs(sys.argv[1]) else os.path.join(BASE, sys.argv[1])
    pr = p.Presentations.Open(ruta, ReadOnly=True, Untitled=False, WithWindow=False)
    try:
        out = []
        for i in range(1, pr.Slides.Count + 1):
            sl = pr.Slides(i); t = sl.SlideShowTransition; se = t.SoundEffect
            d = {"diapo": i, "transicion": t.EntryEffect, "sonido_tipo": se.Type, "sonido": se.Name if se.Type == 2 or se.Type == 3 else "", "bucle": t.LoopSoundUntilNext, "efectos": []}
            seq = sl.TimeLine.MainSequence
            for k in range(1, seq.Count + 1):
                e = seq.Item(k)
                try: s2 = e.EffectInformation.SoundEffect; d["efectos"].append({"tipo": s2.Type, "sonido": s2.Name})
                except Exception as ex: d["efectos"].append({"error": str(ex)[:80]})
            out.append(d)
        print(json.dumps(out, ensure_ascii=False))
    finally:
        pr.Close()

if __name__ == "__main__":
    main()

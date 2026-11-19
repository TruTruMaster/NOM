#!/usr/bin/env python3
"""Dónde corta Word un renglón: compara los cortes de Word (w14-renglones, hecho con windows/hacer_renglones.py y
medido con verdad.py) con varios modelos del ancho del texto, para quedarse con el que explica los cortes.

  acomodo_word.py [w14-renglones]

Los anchos salen de las fuentes de NOM con las medidas de Office (fonts/office/: el mismo avance de cada letra),
medidos en Chromium con canvas a 1000 px (sin interletraje ni ligaduras: la suma de los avances).
Para cada renglón que no es el último de su párrafo: «dentro» = lo que Word dejó en el renglón (sin el espacio del
final) y «con la siguiente» = eso + espacio + la primera palabra del renglón siguiente. Un modelo explica el corte si
dentro <= ancho disponible < con la siguiente.
"""
import json, math, os, sys
from collections import defaultdict
from pathlib import Path
from playwright.sync_api import sync_playwright

CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas"))
BASE = os.environ.get("NOM_BASE", "http://localhost:8096")
NOM_FUENTE = {"Calibri": "NOM CAL", "Aptos": "NOM A", "Times New Roman": "NOM TI", "Arial": "NOM AR", "Cambria": "NOM CA",
              "Georgia": "NOM GE", "Verdana": "NOM VE", "Segoe UI": "NOM SU"}


def lineas_word(verdad):
    out = []
    for p in verdad["paginas"]:
        for z in p["zonas"]:
            if z.get("historia") != "cuerpo": continue
            for l in z["lineas"]: out.append(l["texto"])
    return out


def avances(fuentes, letras):
    with sync_playwright() as pw:
        b = pw.chromium.launch(); pg = b.new_page()
        pg.goto(BASE + "/"); pg.wait_for_timeout(1500)
        r = pg.evaluate("""async ([fuentes, letras]) => {
            const out = {};
            for (const f of fuentes) {
              await document.fonts.load(`400 100px '${f}'`, letras.join(''));
              const c = document.createElement('canvas').getContext('2d');
              c.font = `400 1000px '${f}'`; c.fontKerning = 'none';
              const m = {}; for (const ch of letras) m[ch] = c.measureText(ch).width / 1000;
              out[f] = { m, ok: document.fonts.check(`400 100px '${f}'`) };
            }
            return out; }""", [fuentes, letras])
        b.close()
    return r


def main():
    nombre = sys.argv[1] if len(sys.argv) > 1 else "w14-renglones"
    meta = json.load(open(CACHE / "office-real/archivos" / (nombre + ".json"), encoding="utf-8"))
    verdad = json.load(open(CACHE / "office-real/verdad" / (nombre + ".verdad.json"), encoding="utf-8"))
    L = lineas_word(verdad)
    # renglones de cada párrafo
    pars, cur = [], []
    for t in L:
        cur.append(t)
        if t.endswith("\r"): pars.append(cur); cur = []
    assert len(pars) == len(meta["parrafos"]), (len(pars), len(meta["parrafos"]))
    letras = sorted(set("".join(p["texto"] for p in meta["parrafos"])) | {" "})
    cache_adv = CACHE / "office-real" / (nombre + ".avances.json")
    adv = json.load(open(cache_adv)) if cache_adv.exists() else avances(sorted(set(NOM_FUENTE.values())), letras)
    json.dump(adv, open(cache_adv, "w"))
    for f, v in adv.items():
        if not v["ok"]: print("¡la fuente no cargó!", f)
    casos = []
    for p, ls in zip(meta["parrafos"], pars):
        assert "".join(ls).rstrip("\r") == p["texto"], (p["texto"][:60], "".join(ls)[:60])
        for i in range(len(ls) - 1):
            dentro = ls[i].rstrip(" ")
            sig = ls[i + 1].rstrip("\r").split(" ")[0]
            casos.append(dict(f=p["fuente"], tam=p["tam"], alin=p["alin"], A=meta["ancho"] - p["sangria"], dentro=dentro,
                              espacios=len(ls[i]) - len(dentro), sig=sig, n_esp=dentro.count(" "), i=i, n_lineas=len(ls),
                              par=len(casos) and casos[-1].get("par", 0) + (1 if i == 0 else 0) or 0))
    print(f"{len(pars)} párrafos, {len(casos)} cortes")
    m = {f: adv[NOM_FUENTE[f]]["m"] for f in NOM_FUENTE}

    def ancho(texto, f, tam, modelo):
        a = m[f]
        if modelo == "ideal": return sum(a[c] for c in texto) * tam
        if modelo == "ppem":      # tamaño en píxeles enteros a 96 ppp
            ppem = round(tam * 96 / 72); return sum(a[c] for c in texto) * ppem * 72 / 96
        if modelo == "px96":      # cada letra a píxeles enteros (96 ppp)
            s = tam * 96 / 72; return sum(round(a[c] * s) for c in texto) * 72 / 96
        if modelo == "ppem+px96":
            ppem = round(tam * 96 / 72); return sum(round(a[c] * ppem) for c in texto) * 72 / 96
        if modelo == "twip":      # cada letra a 1/20 de punto
            return sum(round(a[c] * tam * 20) for c in texto) / 20
        if modelo == "600":       # cada letra a 600 ppp
            s = tam * 600 / 72; return sum(round(a[c] * s) for c in texto) * 72 / 600
        if modelo == "ppem600":
            ppem = round(tam * 600 / 72); return sum(a[c] for c in texto) * ppem * 72 / 600
        raise ValueError(modelo)

    for modelo in ("ideal", "ppem", "px96", "ppem+px96", "twip", "600", "ppem600"):
        por = defaultdict(lambda: [0, 0, 0, [], []])
        for c in casos:
            d = ancho(c["dentro"], c["f"], c["tam"], modelo)
            s = ancho(c["dentro"] + " " + c["sig"], c["f"], c["tam"], modelo)
            k = (c["alin"],)
            x = por[k]
            x[0] += 1
            if d > c["A"] + 1e-6: x[1] += 1; x[3].append(round(d - c["A"], 3))          # Word dejó pasar el margen
            if s <= c["A"] + 1e-6: x[2] += 1; x[4].append(round(c["A"] - s, 3))         # Word cortó aunque cabía
        print(f"\n== modelo {modelo}")
        for k, x in sorted(por.items()):
            print(f"  {k[0]:5s} cortes {x[0]:4d} · se pasa del margen {x[1]:3d} (máx {max(x[3], default=0):.3f} pt) · cortó aunque cabía {x[2]:3d} (máx holgura {max(x[4], default=0):.3f} pt)")
    # detalle del modelo ideal: lo que se pasa, en proporción del espacio
    print("\n== ideal, justificado: cuánto se pasa, contra el número de espacios del renglón")
    filas = []
    for c in casos:
        d = ancho(c["dentro"], c["f"], c["tam"], "ideal")
        if c["alin"] == "just" and d > c["A"]:
            esp = m[c["f"]][" "] * c["tam"]
            filas.append((round((d - c["A"]) / (c["n_esp"] * esp), 4), round(d - c["A"], 3), c["n_esp"], c["f"], c["tam"]))
    filas.sort(reverse=True)
    for f in filas[:15]: print("  ", f)
    print("  encogido máximo por espacio:", filas[0][0] if filas else None, "· casos", len(filas))
    # y los justificados que NO se encogieron aunque cabía encogiendo
    print("\n== justificado: cortes donde la siguiente palabra cabía encogiendo los espacios (y Word no la metió)")
    for lim in (0.1, 0.125, 0.15, 0.2, 0.25, 0.33):
        n = 0
        for c in casos:
            if c["alin"] != "just": continue
            s = ancho(c["dentro"] + " " + c["sig"], c["f"], c["tam"], "ideal")
            esp = m[c["f"]][" "] * c["tam"]
            if s - (c["n_esp"] + 1) * esp * lim <= c["A"]: n += 1
        print(f"   con encogido de {lim:.3f}: {n} cortes que debieron meter la palabra")
    json.dump(casos, open(CACHE / "office-real" / (nombre + ".cortes.json"), "w", encoding="utf-8"), ensure_ascii=False)


if __name__ == "__main__":
    main()

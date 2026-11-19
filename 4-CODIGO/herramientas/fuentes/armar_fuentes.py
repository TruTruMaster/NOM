#!/usr/bin/env python3
"""Arma todas las fuentes con las medidas de Office que lleva NOM y su CSS.

  ~/.cache/nom-pruebas/venv/bin/python armar_fuentes.py [--office DIR] [--libres DIR]

--office: carpeta con las fuentes de Windows/Office de donde se toman las medidas (no van en el
          repositorio; se copian de C:\\Windows\\Fonts y de la caché de fuentes de Office de un equipo con
          Office 2024; ver docs/PARIDAD-OFFICE-REAL.md).
--libres: carpeta con las fuentes libres que dan el dibujo (Open Sans, Gelasio, EB Garamond: se bajan de
          github.com/googlefonts y github.com/google/fonts; Lato, Liberation y Selawik del sistema). Todas OFL.

Escribe app/src/main/resources/static/fonts/office/*.woff2 y css/fuentes-office.css, y la tabla de
nombres para modo-texto.js en fonts/office/tabla.json.
"""
import argparse, json, os, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import medidas_office as M
from fontTools.ttLib import TTFont

RAIZ = Path(__file__).resolve().parents[2]
STATIC = RAIZ / "app/src/main/resources/static"
OFFICE = Path.home() / ".cache/nom-pruebas/fuentes-office"
LIBRES = Path.home() / ".cache/nom-pruebas/fuentes-libres"
SIS = Path("/usr/share/fonts/truetype")
SELAWIK = Path.home() / ".local/share/fonts/selawik"

# Dibujo de cada estilo: (archivo de la fuente libre, ejes si es variable)
def os_(w=400, wd=100): return ("OpenSans[wdth,wght].ttf", {"wght": w, "wdth": wd})
def osi(w=400, wd=100): return ("OpenSans-Italic[wdth,wght].ttf", {"wght": w, "wdth": wd})
OPEN_SANS = {"Regular": os_(), "Bold": os_(700), "Italic": osi(), "Bold Italic": osi(700)}
OPEN_SANS_ANGOSTA = {"Regular": os_(400, 75), "Bold": os_(700, 75), "Italic": osi(400, 75), "Bold Italic": osi(700, 75)}
LATO = {"Regular": (str(SIS / "lato/Lato-Regular.ttf"), None), "Bold": (str(SIS / "lato/Lato-Bold.ttf"), None), "Italic": (str(SIS / "lato/Lato-Italic.ttf"), None), "Bold Italic": (str(SIS / "lato/Lato-BoldItalic.ttf"), None)}
LATO_LIGERA = {"Regular": (str(SIS / "lato/Lato-Light.ttf"), None), "Italic": (str(SIS / "lato/Lato-LightItalic.ttf"), None)}
SELAWIK_ = {"Regular": (str(SELAWIK / "selawk.ttf"), None), "Bold": (str(SELAWIK / "selawkb.ttf"), None), "Italic": (str(SELAWIK / "selawk.ttf"), None), "Bold Italic": (str(SELAWIK / "selawkb.ttf"), None)}
GELASIO = {"Regular": ("gelasio-Gelasio[wght].ttf", {"wght": 400}), "Bold": ("gelasio-Gelasio[wght].ttf", {"wght": 700}), "Italic": ("gelasio-Gelasio-Italic[wght].ttf", {"wght": 400}), "Bold Italic": ("gelasio-Gelasio-Italic[wght].ttf", {"wght": 700})}
GARAMOND = {"Regular": ("ebgaramond-EBGaramond[wght].ttf", {"wght": 400}), "Bold": ("ebgaramond-EBGaramond[wght].ttf", {"wght": 700}), "Italic": ("ebgaramond-EBGaramond-Italic[wght].ttf", {"wght": 400}), "Bold Italic": ("ebgaramond-EBGaramond-Italic[wght].ttf", {"wght": 700})}
F = STATIC / "fonts"
CALADEA = {"Regular": (str(F / "caladea-regular.woff2"), None), "Bold": (str(F / "caladea-negrita.woff2"), None), "Italic": (str(F / "caladea-cursiva.woff2"), None), "Bold Italic": (str(F / "caladea-negrita-cursiva.woff2"), None)}
LIB = SIS / "liberation"
CARLITO = {"Regular": (str(F / "carlito-regular.woff2"), None), "Bold": (str(F / "carlito-negrita.woff2"), None), "Italic": (str(F / "carlito-cursiva.woff2"), None), "Bold Italic": (str(F / "carlito-negrita-cursiva.woff2"), None)}
LIB_SANS = {e: (str(LIB / f"LiberationSans-{n}.ttf"), None) for e, n in [("Regular", "Regular"), ("Bold", "Bold"), ("Italic", "Italic"), ("Bold Italic", "BoldItalic")]}
LIB_SERIF = {e: (str(LIB / f"LiberationSerif-{n}.ttf"), None) for e, n in [("Regular", "Regular"), ("Bold", "Bold"), ("Italic", "Italic"), ("Bold Italic", "BoldItalic")]}
LIB_MONO = {e: (str(LIB / f"LiberationMono-{n}.ttf"), None) for e, n in [("Regular", "Regular"), ("Bold", "Bold"), ("Italic", "Italic"), ("Bold Italic", "BoldItalic")]}
COMIC = {"Regular": (str(F / "comic-neue-regular.woff2"), None), "Bold": (str(F / "comic-neue-negrita.woff2"), None), "Italic": (str(F / "comic-neue-cursiva.woff2"), None), "Bold Italic": (str(F / "comic-neue-negrita-cursiva.woff2"), None)}

# (nombre en Office, nombre de la fuente nueva, archivos de Windows por estilo, dibujo, letras de la base que se ensanchan/encogen como máximo)
FAMILIAS = [
    # Calibri, Arial, Times y Courier ya tienen equivalentes de las mismas medidas (Carlito, Liberation), pero su
    # interletraje no es el de Office: Office 2024 lo activa por omisión, así que también se rehacen.
    ("Calibri", "NOM CAL", {"Regular": "windows/calibri.ttf", "Bold": "windows/calibrib.ttf", "Italic": "windows/calibrii.ttf", "Bold Italic": "windows/calibriz.ttf"}, CARLITO),
    ("Arial", "NOM AR", {"Regular": "windows/arial.ttf", "Bold": "windows/arialbd.ttf", "Italic": "windows/ariali.ttf", "Bold Italic": "windows/arialbi.ttf"}, LIB_SANS),
    ("Times New Roman", "NOM TI", {"Regular": "windows/times.ttf", "Bold": "windows/timesbd.ttf", "Italic": "windows/timesi.ttf", "Bold Italic": "windows/timesbi.ttf"}, LIB_SERIF),
    ("Courier New", "NOM CU", {"Regular": "windows/cour.ttf", "Bold": "windows/courbd.ttf", "Italic": "windows/couri.ttf", "Bold Italic": "windows/courbi.ttf"}, LIB_MONO),
    ("Aptos", "NOM A", {"Regular": "aptos/Aptos/30153066857.ttf", "Bold": "aptos/Aptos/32483553004.ttf", "Italic": "aptos/Aptos/26395700251.ttf", "Bold Italic": "aptos/Aptos/26205970649.ttf"}, OPEN_SANS),
    ("Aptos Display", "NOM AD", {"Regular": "aptos/Aptos Display/30169865670.ttf", "Bold": "aptos/Aptos Display/32677218994.ttf", "Italic": "aptos/Aptos Display/31638818185.ttf", "Bold Italic": "aptos/Aptos Display/25576969643.ttf"}, OPEN_SANS),
    ("Aptos Narrow", "NOM AN", {"Regular": "aptos/Aptos Narrow/27970306811.ttf", "Bold": "aptos/Aptos Narrow/25065980124.ttf", "Italic": "aptos/Aptos Narrow/28903934312.ttf", "Bold Italic": "aptos/Aptos Narrow/29867269125.ttf"}, OPEN_SANS_ANGOSTA),
    ("Cambria", "NOM CA", {"Regular": "windows/cambria.ttc", "Bold": "windows/cambriab.ttf", "Italic": "windows/cambriai.ttf", "Bold Italic": "windows/cambriaz.ttf"}, CALADEA),
    ("Calibri Light", "NOM CL", {"Regular": "windows/calibril.ttf", "Italic": "windows/calibrili.ttf"}, LATO_LIGERA),
    ("Segoe UI", "NOM SU", {"Regular": "windows/segoeui.ttf", "Bold": "windows/segoeuib.ttf", "Italic": "windows/segoeuii.ttf", "Bold Italic": "windows/segoeuiz.ttf"}, SELAWIK_),
    ("Georgia", "NOM GE", {"Regular": "windows/georgia.ttf", "Bold": "windows/georgiab.ttf", "Italic": "windows/georgiai.ttf", "Bold Italic": "windows/georgiaz.ttf"}, GELASIO),
    ("Verdana", "NOM VE", {"Regular": "windows/verdana.ttf", "Bold": "windows/verdanab.ttf", "Italic": "windows/verdanai.ttf", "Bold Italic": "windows/verdanaz.ttf"}, OPEN_SANS),
    ("Tahoma", "NOM TA", {"Regular": "windows/tahoma.ttf", "Bold": "windows/tahomabd.ttf"}, OPEN_SANS),
    ("Trebuchet MS", "NOM TR", {"Regular": "windows/trebuc.ttf", "Bold": "windows/trebucbd.ttf", "Italic": "windows/trebucit.ttf", "Bold Italic": "windows/trebucbi.ttf"}, OPEN_SANS),
    ("Arial Narrow", "NOM ARN", {"Regular": "windows/ARIALN.TTF", "Bold": "windows/ARIALNB.TTF", "Italic": "windows/ARIALNI.TTF", "Bold Italic": "windows/ARIALNBI.TTF"}, OPEN_SANS_ANGOSTA),
    ("Arial Black", "NOM ARB", {"Regular": "windows/ariblk.ttf"}, {"Regular": os_(800)}),
    ("Consolas", "NOM CO", {"Regular": "windows/consola.ttf", "Bold": "windows/consolab.ttf", "Italic": "windows/consolai.ttf", "Bold Italic": "windows/consolaz.ttf"}, LIB_MONO),
    ("Lucida Console", "NOM LC", {"Regular": "windows/lucon.ttf"}, LIB_MONO),
    ("Century Gothic", "NOM CG", {"Regular": "windows/GOTHIC.TTF", "Bold": "windows/GOTHICB.TTF", "Italic": "windows/GOTHICI.TTF", "Bold Italic": "windows/GOTHICBI.TTF"}, OPEN_SANS),
    ("Garamond", "NOM GA", {"Regular": "windows/GARA.TTF", "Bold": "windows/GARABD.TTF", "Italic": "windows/GARAIT.TTF"}, GARAMOND),
    ("Book Antiqua", "NOM BA", {"Regular": "windows/BKANT.TTF", "Bold": "windows/ANTQUAB.TTF", "Italic": "windows/ANTQUAI.TTF", "Bold Italic": "windows/ANTQUABI.TTF"}, GELASIO),
    ("Palatino Linotype", "NOM PL", {"Regular": "windows/pala.ttf", "Bold": "windows/palab.ttf", "Italic": "windows/palai.ttf", "Bold Italic": "windows/palabi.ttf"}, GELASIO),
    ("Comic Sans MS", "NOM CS", {"Regular": "windows/comic.ttf", "Bold": "windows/comicbd.ttf", "Italic": "windows/comici.ttf", "Bold Italic": "windows/comicz.ttf"}, COMIC),
    ("Impact", "NOM IM", {"Regular": "windows/impact.ttf"}, {"Regular": os_(800, 75)}),
    ("Franklin Gothic Medium", "NOM FG", {"Regular": "windows/framd.ttf", "Italic": "windows/framdit.ttf"}, {"Regular": os_(600), "Italic": osi(600)}),
    ("Gill Sans MT", "NOM GS", {"Regular": "windows/GIL_____.TTF", "Bold": "windows/GILB____.TTF", "Italic": "windows/GILI____.TTF", "Bold Italic": "windows/GILBI___.TTF"}, LATO),
    ("Candara", "NOM CN", {"Regular": "windows/Candara.ttf", "Bold": "windows/Candarab.ttf", "Italic": "windows/Candarai.ttf", "Bold Italic": "windows/Candaraz.ttf"}, LATO),
    ("Corbel", "NOM CB", {"Regular": "windows/corbel.ttf", "Bold": "windows/corbelb.ttf", "Italic": "windows/corbeli.ttf", "Bold Italic": "windows/corbelz.ttf"}, OPEN_SANS),
    ("Constantia", "NOM CT", {"Regular": "windows/constan.ttf", "Bold": "windows/constanb.ttf", "Italic": "windows/constani.ttf", "Bold Italic": "windows/constanz.ttf"}, CALADEA),
    ("Segoe Print", "NOM SP", {"Regular": "windows/segoepr.ttf", "Bold": "windows/segoeprb.ttf"}, COMIC),
    ("Segoe Script", "NOM SS", {"Regular": "windows/segoesc.ttf", "Bold": "windows/segoescb.ttf"}, COMIC),
    ("Bahnschrift", "NOM BH", {"Regular": "windows/bahnschrift.ttf"}, {"Regular": os_(400, 87)}),
    # (19-nov) las de w01 que faltaban: Sitka Text es la instancia «Text» de la fuente variable (opsz 11)
    ("Sitka Text", "NOM SKT", {"Regular": "windows/SitkaVF.ttf|opsz=11,wght=400", "Bold": "windows/SitkaVF.ttf|opsz=11,wght=700",
                              "Italic": "windows/SitkaVF-Italic.ttf|opsz=11,wght=400", "Bold Italic": "windows/SitkaVF-Italic.ttf|opsz=11,wght=700"}, GELASIO),
    ("Lucida Sans Unicode", "NOM LSU", {"Regular": "windows/l_10646.ttf"}, OPEN_SANS),
    ("Ink Free", "NOM IF", {"Regular": "windows/Inkfree.ttf"}, COMIC),
]
ARCHIVO_ESTILO = {"Regular": "regular", "Bold": "negrita", "Italic": "cursiva", "Bold Italic": "negrita-cursiva"}
CSS_ESTILO = {"Regular": ("normal", 400), "Bold": ("normal", 700), "Italic": ("italic", 400), "Bold Italic": ("italic", 700)}
COPY = {"Liberation": "Digitized data copyright (c) 2010 Google Corporation; (c) 2012 Red Hat, Inc. (Liberation fonts, SIL OFL 1.1).", "carlito": "Copyright (c) 2010-2013 by tyPoland Lukasz Dziedzic (Carlito, SIL OFL 1.1).", "OpenSans": "Copyright 2020 The Open Sans Project Authors (https://github.com/googlefonts/opensans).", "Lato": "Copyright (c) 2011-2015 by tyPoland Lukasz Dziedzic.",
        "selawk": "Copyright (c) Microsoft Corporation (Selawik, SIL OFL 1.1).",
        "Gelasio": "Copyright 2022 The Gelasio Project Authors (https://github.com/SorkinType/Gelasio).", "EBGaramond": "Copyright 2017 The EB Garamond Project Authors.",
        "caladea": "Copyright (c) 2012 Huerta Tipografica (Caladea).", "comic-neue": "Copyright 2014 The Comic Neue Project Authors."}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--office", default=str(OFFICE)); ap.add_argument("--libres", default=str(LIBRES))
    ap.add_argument("--solo", default="", help="armar solo estas familias (coma)")
    a = ap.parse_args()
    sale = STATIC / "fonts/office"; sale.mkdir(parents=True, exist_ok=True)
    css = ["/* NOM · Fuentes con LAS MISMAS MEDIDAS que las de Office (ancho de cada letra y alto del renglón).",
           "   Las arma herramientas/fuentes/armar_fuentes.py: el dibujo es de una fuente libre (OFL) y las medidas,",
           "   de la fuente de Office. Así un documento de Office corta sus renglones y sus páginas igual que en",
           "   Word, Excel o PowerPoint, en la computadora y en el teléfono. Licencias: fonts/LICENCIAS-FUENTES.txt.",
           "   Solo se descargan las que un documento usa de verdad. */"]
    tabla = {}
    solo = {x.strip().lower() for x in a.solo.split(",") if x.strip()}
    for oficina, nueva, archivos, dibujo in FAMILIAS:
        if solo and oficina.lower() not in solo: continue
        slug = nueva.lower().replace(" ", "-")
        for estilo, rel in archivos.items():
            # («archivo|eje=valor,…»: una instancia de una fuente variable de Office, como la usa Office)
            rel, _, ejes_obj = rel.partition("|")
            obj = Path(a.office) / rel
            base, ejes = dibujo.get(estilo) or dibujo["Regular"]
            rb = base if base.startswith("/") else str(Path(a.libres) / base)
            if not obj.exists() or not Path(rb).exists():
                print("  falta", obj if not obj.exists() else rb); continue
            destino = sale / f"{slug}-{ARCHIVO_ESTILO[estilo]}.woff2"
            cr = next((v for k, v in COPY.items() if k.lower() in rb.lower()), "")
            objetivo = TTFont(str(obj), fontNumber=0)
            if ejes_obj:
                from fontTools.varLib import instancer
                objetivo = instancer.instantiateVariableFont(objetivo, {k: float(v) for k, v in (e.split("=") for e in ejes_obj.split(","))}, updateFontNames=False)
            n, faltan = M.clonar(objetivo, M.cargar_base(rb, ejes), nueva, estilo, cr, str(destino))
            st, peso = CSS_ESTILO[estilo]
            css.append(f'@font-face {{ font-family: "{nueva}"; font-style: {st}; font-weight: {peso}; font-display: swap; src: url("../fonts/office/{destino.name}") format("woff2"); }}')
            print(f"{oficina:24s} {estilo:12s} -> {destino.name:34s} {n:4d} letras, {destino.stat().st_size // 1024} KB")
        tabla[oficina.lower()] = nueva
    if solo:
        # (con --solo se conservan las demás familias del CSS y de la tabla: solo se cambian las que se armaron)
        hechas = {nueva for oficina, nueva, _, _ in FAMILIAS if oficina.lower() in solo}
        viejo = (STATIC / "css/fuentes-office.css").read_text().splitlines()
        css = [l for l in viejo if not any(f'font-family: "{h}";' in l for h in hechas)] + [l for l in css[5:]]
        tabla = dict(json.loads((sale / "tabla.json").read_text()), **tabla)
    (STATIC / "css/fuentes-office.css").write_text("\n".join(css) + "\n")
    (sale / "tabla.json").write_text(json.dumps(tabla, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()

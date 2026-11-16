#!/usr/bin/env python3
"""Arma fuentes libres con LAS MISMAS MEDIDAS que una fuente de Office.

Un documento de Office 2024 se escribe en Aptos (Word), Aptos Narrow (Excel) o Aptos Display (títulos).
Ningún equipo con Android o Linux trae esas fuentes, y su licencia no deja repartirlas. Si se pintan con
otra fuente, cada letra mide distinto: los renglones se cortan en otro lugar y el documento tiene otro
número de páginas que en Office.

Esta herramienta hace lo mismo que se hizo con Carlito (Calibri) o Liberation (Arial, Times):
  - de la fuente de Office se toman SOLO sus medidas: el ancho de cada letra (avance), el alto del
    renglón (ascendente, descendente, interlineado) y la posición del subrayado y del tachado;
  - el dibujo de cada letra sale de una fuente libre (licencia SIL OFL 1.1); se ajusta a lo ancho para
    que cada letra ocupe exactamente el ancho que tiene en Office. Nunca se copia el dibujo de Office.
El resultado es una fuente nueva con licencia OFL, con otro nombre, que mide igual que la de Office.

  medidas_office.py --objetivo AptosRegular.ttf --base OpenSans[wdth,wght].ttf --eje wght=400,wdth=100 \\
      --familia "NOM Sans A" --estilo Regular --salida nom-sans-a-regular.woff2

Las fuentes de Office solo se leen en la computadora de quien arma (no van en el repositorio).
"""
import argparse, sys
from fontTools.ttLib import TTFont, newTable
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.varLib import instancer

# Lo que se lleva: latín (con todos los acentos de Europa), griego, cirílico, puntuación, monedas y los
# signos de uso común. El resto de los alfabetos cae a la fuente del equipo.
RANGOS = [(0x20, 0x7E), (0xA0, 0x24F), (0x250, 0x2AF), (0x2B0, 0x36F), (0x370, 0x3FF), (0x400, 0x52F), (0x1E00, 0x1EFF),
          (0x2000, 0x206F), (0x2070, 0x209F), (0x20A0, 0x20CF), (0x2100, 0x214F), (0x2150, 0x218F), (0x2190, 0x21FF),
          (0x2200, 0x22FF), (0x2300, 0x23FF), (0x25A0, 0x25FF), (0x2600, 0x26FF), (0x2700, 0x27BF), (0xFB00, 0xFB06), (0xFEFF, 0xFEFF), (0xFFFD, 0xFFFD)]


def en_rangos(cp):
    return any(a <= cp <= b for a, b in RANGOS)


def cargar_base(ruta, ejes):
    f = TTFont(ruta)
    if "fvar" in f and ejes:
        f = instancer.instantiateVariableFont(f, ejes, updateFontNames=False)
    return f


def rasgos_de(objetivo, nombre_clon, k):
    """Interletraje (kern) y ligaduras (liga) de la fuente de Office, como texto de rasgos OpenType, en
    los nombres de letra de la fuente nueva. Son medidas (cuánto se acercan dos letras; qué par de letras
    se vuelve una), no dibujo."""
    cm = objetivo.getBestCmap()
    a_clon = {}
    for cp, g in cm.items():
        n = nombre_clon.get(cp)
        if n: a_clon.setdefault(g, []).append(n)
    lin, nclase = [], [0]
    def clase(glifos):
        nombres = sorted({n for g in glifos for n in a_clon.get(g, [])})
        if not nombres: return None
        nclase[0] += 1
        lin.append(f"@c{nclase[0]} = [{' '.join(nombres)}];")
        return f"@c{nclase[0]}"
    kern_lk, liga_lk = [], []
    if "GPOS" in objetivo:
        g = objetivo["GPOS"].table
        idx = []
        for fr in g.FeatureList.FeatureRecord:
            if fr.FeatureTag == "kern":
                for i in fr.Feature.LookupListIndex:
                    if i not in idx: idx.append(i)
        for li in sorted(idx):
            lk = g.LookupList.Lookup[li]
            cuerpo = []
            for st in lk.SubTable:
                if lk.LookupType == 9: st = st.ExtSubTable
                if getattr(st, "LookupType", 2) != 2: continue
                if st.Format == 1:
                    cov = st.Coverage.glyphs
                    for i, ps in enumerate(st.PairSet):
                        for pvr in ps.PairValueRecord:
                            v = getattr(pvr.Value1, "XAdvance", 0) if pvr.Value1 else 0
                            if not v: continue
                            for a in a_clon.get(cov[i], []):
                                for b in a_clon.get(pvr.SecondGlyph, []):
                                    cuerpo.append(f"pos {a} {b} {round(v * k)};")
                elif st.Format == 2:
                    cov = set(st.Coverage.glyphs)
                    c1, c2 = st.ClassDef1.classDefs, st.ClassDef2.classDefs
                    izq = {}
                    for gl in cov: izq.setdefault(c1.get(gl, 0), []).append(gl)
                    der = {}
                    for gl, c in c2.items(): der.setdefault(c, []).append(gl)
                    # clase 0 de la derecha: todas las letras que no están en otra clase
                    der[0] = [gl for gl in a_clon if gl not in c2]
                    nombres_izq, nombres_der = {}, {}
                    for i1, rec in enumerate(st.Class1Record):
                        if i1 not in izq: continue
                        for i2, r2 in enumerate(rec.Class2Record):
                            if i2 not in der: continue
                            v = getattr(r2.Value1, "XAdvance", 0) if r2.Value1 else 0
                            if not v: continue
                            if i1 not in nombres_izq: nombres_izq[i1] = clase(izq[i1])
                            if i2 not in nombres_der: nombres_der[i2] = clase(der[i2])
                            if nombres_izq[i1] and nombres_der[i2]: cuerpo.append(f"pos {nombres_izq[i1]} {nombres_der[i2]} {round(v * k)};")
                cuerpo.append("subtable;")
            # sin separaciones al principio, al final ni dobles (feaLib no las acepta)
            limpio = []
            for x in cuerpo:
                if x == "subtable;" and (not limpio or limpio[-1] == "subtable;"): continue
                limpio.append(x)
            while limpio and limpio[-1] == "subtable;": limpio.pop()
            if limpio: kern_lk.append(limpio)
    if "GSUB" in objetivo:
        g = objetivo["GSUB"].table
        for fr in g.FeatureList.FeatureRecord:
            if fr.FeatureTag != "liga": continue
            for li in fr.Feature.LookupListIndex:
                lk = g.LookupList.Lookup[li]
                for st in lk.SubTable:
                    if lk.LookupType == 7: st = st.ExtSubTable
                    if not hasattr(st, "ligatures"): continue
                    for primero, ligs in st.ligatures.items():
                        for l in ligs:
                            comp = [primero] + list(l.Component)
                            if all(a_clon.get(x) for x in comp) and a_clon.get(l.LigGlyph):
                                regla = f"sub {' '.join(a_clon[x][0] for x in comp)} by {a_clon[l.LigGlyph][0]};"
                                if regla not in liga_lk: liga_lk.append(regla)
            break
    fea = ["languagesystem DFLT dflt;", "languagesystem latn dflt;", "languagesystem grek dflt;", "languagesystem cyrl dflt;"] + lin
    if kern_lk:
        for i, c in enumerate(kern_lk): fea += [f"lookup kern{i} {{"] + ["  " + x for x in c] + [f"}} kern{i};"]
        fea += ["feature kern {"] + [f"  lookup kern{i};" for i in range(len(kern_lk))] + ["} kern;"]
    if liga_lk:
        # las de tres letras primero (ffi antes de ff)
        liga_lk.sort(key=lambda r: -len(r.split(" by ")[0].split()))
        fea += ["feature liga {"] + ["  " + x for x in liga_lk] + ["} liga;"]
    return "\n".join(fea) if (kern_lk or liga_lk) else ""


def clonar(objetivo, base, familia, estilo, copyright_base, salida, recorte=(0.82, 1.18)):
    upm_o = objetivo["head"].unitsPerEm
    upm = base["head"].unitsPerEm
    k = upm / upm_o                         # unidades de Office → unidades de la base
    cmap_o, cmap_b = objetivo.getBestCmap(), base.getBestCmap()
    hm_o, hm_b = objetivo["hmtx"].metrics, base["hmtx"].metrics
    gs = base.getGlyphSet()
    os2o, os2b = objetivo["OS/2"], base["OS/2"]
    # Escala de dibujo: que las mayúsculas midan de alto lo mismo que en Office.
    cap_o = getattr(os2o, "sCapHeight", 0) or 0.7 * upm_o
    cap_b = getattr(os2b, "sCapHeight", 0) or 0.7 * upm
    s = (cap_o * k) / cap_b

    orden, cmap, glifos, metricas = [".notdef"], {}, {}, {}
    # .notdef: un rectángulo hueco del ancho del de Office
    adv0 = round(hm_o.get(".notdef", (upm_o // 2, 0))[0] * k)
    p = TTGlyphPen(None)
    a, d = round(cap_o * k), 0
    p.moveTo((50, d)); p.lineTo((50, a)); p.lineTo((adv0 - 50, a)); p.lineTo((adv0 - 50, d)); p.closePath()
    p.moveTo((100, d + 50)); p.lineTo((adv0 - 100, d + 50)); p.lineTo((adv0 - 100, a - 50)); p.lineTo((100, a - 50)); p.closePath()
    glifos[".notdef"] = p.glyph(); metricas[".notdef"] = (adv0, 50)

    cubica = "CFF " in base or "CFF2" in base   # las curvas cúbicas pasan a cuadráticas (TrueType)
    faltan = 0
    for cp in sorted(cmap_o):
        if not en_rangos(cp): continue
        gb = cmap_b.get(cp)
        if gb is None: faltan += 1; continue
        adv = round(hm_o[cmap_o[cp]][0] * k)       # el ancho de Office (lo que importa)
        nombre = "uni%04X" % cp if cp <= 0xFFFF else "u%05X" % cp
        # dibujo de la base, escalado parejo (s) y luego ajustado a lo ancho para caber en el avance de Office
        # las letras compuestas (con acento) se desarman en contornos sueltos
        rec = DecomposingRecordingPen(gs); gs[gb].draw(rec)
        bp = BoundsPen(gs); rec.replay(bp)
        adv_b = hm_b[gb][0] * s
        pen0 = TTGlyphPen(None)
        pen = Cu2QuPen(pen0, max_err=1.0, reverse_direction=True) if cubica else pen0
        if bp.bounds is None or adv == 0 or adv_b == 0:
            # espacio o marca sin dibujo: solo el ancho; marcas combinantes de ancho cero quedan como están
            tp = TransformPen(pen, (s, 0, 0, s, 0, 0)); rec.replay(tp)
        else:
            sx = max(recorte[0], min(recorte[1], adv / adv_b))
            x0, y0, x1, y1 = bp.bounds
            ancho_tinta = (x1 - x0) * s * sx
            izq_b = x0 * s; der_b = adv_b - x1 * s
            libre = adv - ancho_tinta
            # el espacio que queda se reparte como en la base (lado izquierdo/derecho)
            tot = izq_b + der_b
            # (si la letra se sale de su ancho —márgenes negativos— se escala su posición en proporción)
            izq = libre * (izq_b / tot) if izq_b >= 0 and der_b >= 0 and tot > 1 else x0 * s * (adv / adv_b)
            dx = izq - x0 * s * sx
            tp = TransformPen(pen, (s * sx, 0, 0, s, dx, 0)); rec.replay(tp)
        g = pen0.glyph()
        orden.append(nombre); cmap[cp] = nombre; glifos[nombre] = g
        metricas[nombre] = (adv, 0)

    fb = FontBuilder(upm, isTTF=True)
    fb.setupGlyphOrder(orden)
    fb.setupCharacterMap(cmap)
    fb.setupGlyf(glifos)
    # lsb correcto de cada glifo
    glyf = fb.font["glyf"]
    hm = {}
    for n in orden:
        g = glyf[n]; g.recalcBounds(glyf)
        hm[n] = (metricas[n][0], getattr(g, "xMin", 0) if g.numberOfContours else 0)
    fb.setupHorizontalMetrics(hm)
    hh = objetivo["hhea"]
    asc, desc, gap = round(hh.ascent * k), round(hh.descent * k), round(hh.lineGap * k)
    fb.setupHorizontalHeader(ascent=asc, descent=desc, lineGap=gap)
    nombres = {"familyName": familia, "styleName": estilo, "uniqueFontIdentifier": f"{familia} {estilo}; NOM-COLAB",
               "fullName": f"{familia} {estilo}", "psName": (familia + "-" + estilo).replace(" ", ""),
               "copyright": copyright_base + " Medidas ajustadas para NOM-COLAB (2026).",
               "licenseDescription": "This Font Software is licensed under the SIL Open Font License, Version 1.1.",
               "licenseInfoURL": "https://openfontlicense.org", "version": "Version 1.000"}
    fb.setupNameTable(nombres)
    negrita, cursiva = "Bold" in estilo, "Italic" in estilo
    fb.setupOS2(version=4, sTypoAscender=round(os2o.sTypoAscender * k), sTypoDescender=round(os2o.sTypoDescender * k), sTypoLineGap=round(os2o.sTypoLineGap * k),
                usWinAscent=round(os2o.usWinAscent * k), usWinDescent=round(os2o.usWinDescent * k),
                fsSelection=(os2o.fsSelection & 0x80) | (0x20 if negrita else 0) | (0x01 if cursiva else 0) | (0x40 if not (negrita or cursiva) else 0),
                usWeightClass=os2o.usWeightClass, usWidthClass=os2o.usWidthClass, achVendID="NOMC",
                sxHeight=round(getattr(os2o, "sxHeight", 0) * k), sCapHeight=round(cap_o * k),
                yStrikeoutSize=round(os2o.yStrikeoutSize * k), yStrikeoutPosition=round(os2o.yStrikeoutPosition * k),
                ySubscriptYOffset=round(os2o.ySubscriptYOffset * k), ySuperscriptYOffset=round(os2o.ySuperscriptYOffset * k),
                ySubscriptYSize=round(os2o.ySubscriptYSize * k), ySuperscriptYSize=round(os2o.ySuperscriptYSize * k),
                ySubscriptXSize=round(os2o.ySubscriptXSize * k), ySuperscriptXSize=round(os2o.ySuperscriptXSize * k), fsType=0)
    po = objetivo["post"]
    fb.setupPost(underlinePosition=round(po.underlinePosition * k), underlineThickness=round(po.underlineThickness * k), italicAngle=po.italicAngle)
    fb.setupHead(unitsPerEm=upm, macStyle=(1 if negrita else 0) | (2 if cursiva else 0))
    from fontTools.misc.timeTools import timestampFromString
    fb.font["head"].fontRevision = 1.0
    fb.font["head"].created = fb.font["head"].modified = timestampFromString("Sun Nov 15 00:00:00 2026")
    fb.setupDummyDSIG = None
    f = fb.font
    # interletraje y ligaduras de Office (Office 2024 los activa por omisión en Word)
    fea = rasgos_de(objetivo, cmap, k)
    if fea:
        from fontTools.feaLib.builder import addOpenTypeFeaturesFromString
        addOpenTypeFeaturesFromString(f, fea)
    if salida.endswith(".woff2"): f.flavor = "woff2"
    f.save(salida)
    return len(cmap), faltan


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--objetivo", required=True, help="fuente de Office de donde se toman las medidas")
    ap.add_argument("--base", required=True, help="fuente libre (OFL) de donde sale el dibujo")
    ap.add_argument("--eje", default="", help="ejes de la base variable, p. ej. wght=700,wdth=75")
    ap.add_argument("--familia", required=True); ap.add_argument("--estilo", required=True)
    ap.add_argument("--copyright", default="")
    ap.add_argument("--salida", required=True)
    a = ap.parse_args()
    ejes = {k: float(v) for k, v in (x.split("=") for x in a.eje.split(",") if x)}
    base = cargar_base(a.base, ejes)
    cr = a.copyright or (base["name"].getDebugName(0) or "")
    n, faltan = clonar(TTFont(a.objetivo), base, a.familia, a.estilo, cr, a.salida)
    print(f"{a.salida}: {n} letras con las medidas de Office ({faltan} que la base no tiene quedan fuera)")


if __name__ == "__main__":
    main()

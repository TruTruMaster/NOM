"""Saca de Word 2024 real dónde puede partir cada palabra con «Guiones automáticos» (su diccionario), por idioma.

Cómo: un .docx con Courier New (todas las letras del mismo ancho) y renglones de L letras; cada párrafo es un relleno
de «x» y la palabra, de modo que en el primer renglón caben exactamente k letras de la palabra más el guion. Word parte
en el punto más a la derecha que sea ≤ k (o baja la palabra entera si no hay ninguno). Se pregunta primero con
k = n − 2, luego con el punto encontrado − 1, y así: en pocas vueltas salen todos los puntos de todas las palabras.
La zona de división es la mínima (así Word siempre intenta partir) y sin tope de guiones seguidos.

  guiones_word.py PALABRAS.txt --idioma es-MX --idioma en-US --salida guiones-word.json [--vueltas 8]

PALABRAS.txt: una palabra por renglón (solo letras). El resultado se va juntando: si la salida ya existe, solo se
preguntan las palabras que faltan. Formato: {"es-MX": {"colaboración": [2, 4, 6, 8]}, ...} (posiciones entre letras).
"""
import argparse, json, re, sys, time, zipfile
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import remoto  # noqa: E402

L = 60                       # letras por renglón (--renglon: más para palabras muy largas)
ANCHO_LETRA_PT = 10 * 1229 / 2048   # Courier New a 10 pt: 1229/2048 de eme por letra


def ancho_tw():
    return int(L * ANCHO_LETRA_PT * 20) + 9   # ancho del renglón en vigésimos de punto (con menos de media letra de holgura)
MARGEN_TW = 720
LETRAS = re.compile(r"^[a-zà-öø-ÿā-ſ]{4,120}$")   # (alfabeto latino: las demás escrituras parten en cualquier letra)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W}><w:autoHyphenation/><w:hyphenationZone w:val="15"/><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"""
STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:eastAsia="Courier New"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""


def hacer_docx(consultas, ruta):
    """consultas: [(palabra, idioma, k)]. Cada una, un párrafo: relleno de L − k − 2 «x», un espacio y la palabra."""
    ps = []
    for palabra, idioma, k in consultas:
        relleno = "x" * (L - k - 2)
        ps.append(f'<w:p><w:r><w:rPr><w:lang w:val="{idioma}"/></w:rPr><w:t xml:space="preserve">{relleno} {escape(palabra)}</w:t></w:r></w:p>')
    sect = (f'<w:sectPr><w:pgSz w:w="{min(31680, ancho_tw() + 2 * MARGEN_TW)}" w:h="16838"/><w:pgMar w:top="720" w:right="{MARGEN_TW}" '
            f'w:bottom="720" w:left="{MARGEN_TW}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>')
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {W}><w:body>{"".join(ps)}{sect}</w:body></w:document>'
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/styles.xml", STYLES)


def leer(txt, consultas):
    """Para cada consulta, cuántas letras de la palabra quedaron en el primer renglón (0 = bajó entera)."""
    renglones = txt.replace("﻿", "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out, i = [], 0
    for palabra, idioma, k in consultas:
        relleno = "x" * (L - k - 2)
        while i < len(renglones) and not renglones[i].startswith(relleno):
            i += 1
        if i >= len(renglones):
            out.append(None)
            continue
        r = renglones[i][len(relleno):]
        i += 1
        r = r.lstrip(" ").rstrip(" -­‐‑¬")
        if r and not palabra.startswith(r):
            out.append(None)   # no se entiende: se descarta
            continue
        out.append(len(r))
    return out


def correr(consultas, carpeta, n):
    nombre = f"guiones-{n}"
    docx = carpeta / f"{nombre}.docx"
    hacer_docx(consultas, docx)
    remoto.subir([docx], "nomoffice/entrada")
    r = remoto.correr_office(nombre, "guiones_txt.py", f"entrada\\{docx.name}", f"{nombre}.txt", timeout=3600)
    txt = carpeta / f"{nombre}.txt"
    if not remoto.bajar(f"nomoffice/{nombre}.txt", txt):
        raise SystemExit("Word no dejó el texto: " + r.stdout[-500:])
    print("  ", r.stdout.strip().splitlines()[-2:] if r.stdout.strip() else "", flush=True)
    return leer(txt.read_text(encoding="utf-16"), consultas)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("palabras")
    ap.add_argument("--idioma", action="append", required=True)
    ap.add_argument("--salida", required=True)
    ap.add_argument("--vueltas", type=int, default=12)
    ap.add_argument("--lote", type=int, default=12000, help="párrafos por documento")
    ap.add_argument("--renglon", type=int, default=60, help="letras por renglón (más que la palabra más larga)")
    a = ap.parse_args()
    global L
    L = a.renglon
    salida = Path(a.salida)
    datos = json.loads(salida.read_text(encoding="utf-8")) if salida.exists() else {}
    palabras = []
    for w in Path(a.palabras).read_text(encoding="utf-8").split():
        w = w.strip().lower()
        if LETRAS.match(w) and w not in palabras: palabras.append(w)
    carpeta = salida.parent / (salida.stem + "-trabajo")
    carpeta.mkdir(parents=True, exist_ok=True)
    # estado: (palabra, idioma) → [puntos], próximo k
    pend = {}
    for idioma in a.idioma:
        hechos = datos.setdefault(idioma, {})
        for w in palabras:
            if w not in hechos and len(w) >= 4: pend[(w, idioma)] = ([], len(w) - 2)
    print(f"{len(pend)} palabras por preguntar", flush=True)
    vuelta = 0
    while pend and vuelta < a.vueltas:
        vuelta += 1
        cons = [(w, idioma, k) for (w, idioma), (_, k) in pend.items()]
        res = []
        for j in range(0, len(cons), a.lote):
            t0 = time.time()
            res += correr(cons[j:j + a.lote], carpeta, f"{vuelta}-{j // a.lote}")
            print(f"vuelta {vuelta}: {min(j + a.lote, len(cons))}/{len(cons)} ({time.time() - t0:.0f} s)", flush=True)
        for (w, idioma, k), p in zip(cons, res):
            pts, _ = pend[(w, idioma)]
            if p is None:
                del pend[(w, idioma)]   # (no se entendió: se pregunta otra vez la próxima)
                continue
            if p > 0 and p < len(w):
                pts.append(p)
            if p <= 1:
                datos[idioma][w] = sorted(pts)
                del pend[(w, idioma)]
            else:
                pend[(w, idioma)] = (pts, p - 1)
        salida.write_text(json.dumps(datos, ensure_ascii=False, sort_keys=True), encoding="utf-8")
    for (w, idioma), (pts, _) in pend.items():
        datos[idioma][w] = sorted(pts)
    salida.write_text(json.dumps(datos, ensure_ascii=False, sort_keys=True), encoding="utf-8")
    print("listo:", {k: len(v) for k, v in datos.items()})


if __name__ == "__main__":
    main()

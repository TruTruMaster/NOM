#!/usr/bin/env python3
"""Prueba lo más delicado de la conversión: importar un archivo real de Office, EDITAR algo en
NOM-COLAB y volver a exportar. Tiene que salir el cambio y nada más tiene que moverse.

Cada caso importa en la página real (Chromium), cambia el documento con las mismas operaciones
que usan los editores (doc.cambiar / doc.seq), exporta y revisa el resultado con openpyxl,
python-docx o python-pptx. Con --office además lo abre LibreOffice (tiene que poder).

  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/ediciones.py [--office] [--solo xlsx]
"""
import argparse, base64, os, subprocess, sys, warnings, zipfile, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import fidelidad as F

warnings.simplefilter("ignore")
SALIDA = F.CACHE / "office-ediciones"

# La edición de cada caso se escribe dentro de este programa (donde dice /*EDICION*/) antes de
# mandarlo a la página: la página no deja armar programas a partir de texto (política de
# contenido, sin 'unsafe-eval'), y no hay por qué aflojarla para una prueba.
JS = r"""
async ([b64, nombre]) => {
  const bin = atob(b64), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const aB64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const r = {};
  try {
    const imp = await NOM.formatos.importar(new File([bytes], nombre));
    let doc = NOM.formatos.docDeOps(imp.nombre, imp.modo, imp.ops);
    const ext = nombre.split(".").pop().toLowerCase();
    const exportar = (d) => (ext === "docx" ? NOM.ooxml.docx(d) : ext === "xlsx" ? NOM.ooxml.xlsx(d) : NOM.ooxml.pptx(d, false));
    r.info = await (async (doc, NOM) => {
/*EDICION*/
    })(doc, NOM);
    r.salida = aB64(await exportar(doc));
    r.docx = ext === "docx" ? NOM.ooxml.ultimoDocx : null;
    // y lo mismo pasando por el formato propio: editar, guardar .nom, abrirlo y exportar
    const nom = NOM.formatos.aNomBytes(doc);
    r.nom = nom.length;
    const imp2 = await NOM.formatos.importar(new File([nom], "x.nom"));
    const doc2 = NOM.formatos.docDeOps(imp2.nombre, imp2.modo, imp2.ops);
    r.salidaNom = aB64(await exportar(doc2));
  } catch (e) { r.error = String(e && e.stack || e).slice(0, 700); }
  return r;
}
"""

# ------------------------------------------------------------------ casos
def x_valor(r, orig, sale):
    import openpyxl
    a, b = openpyxl.load_workbook(orig), openpyxl.load_workbook(sale)
    h = b[b.sheetnames[0]]
    assert h["B2"].value == "Editado en NOM-COLAB", f"B2 = {h['B2'].value!r}"
    assert h["B2"].font.b, "B2 debería ir en negrita"
    assert h["AZ300"].value == 1234.5, f"AZ300 = {h['AZ300'].value!r}"
    assert str(h["AZ301"].value).upper().replace(" ", "") == "=SUM(AZ300:AZ300)*2", f"AZ301 = {h['AZ301'].value!r}"
    n = F.comparar_xlsx(orig, sale)
    # B2 se cambió a propósito (valor y negrita): es la única diferencia que se acepta
    def de_mas(k, det):
        ok, de = (int(float(x)) for x in det.split("/")) if "/" in det else (1, 1)
        return de - ok > (1 if k in ("valores y fórmulas", "negrita y cursiva", "formato de número") else 0)
    flojas = {k: d for k, (v, d) in n.m.items() if v < 1 and de_mas(k, d)}
    assert not flojas, f"se movió algo más: {flojas}"
    return f"{len(b.sheetnames)} hojas, resto {n.total()*100:.1f}%"

def x_hojas(r, orig, sale):
    import openpyxl
    a, b = openpyxl.load_workbook(orig), openpyxl.load_workbook(sale)
    assert b.sheetnames[0] == "Renombrada por mí", b.sheetnames
    assert "Hoja nueva" in b.sheetnames, b.sheetnames
    hn = b["Hoja nueva"]
    assert hn["A1"].value == "Total" and hn["B1"].value == 42 and hn["A1"].font.b, (hn["A1"].value, hn["B1"].value)
    assert len(b.sheetnames) == len(a.sheetnames) + 1, (a.sheetnames, b.sheetnames)
    # las demás hojas siguen con sus datos
    for nombre in a.sheetnames[1:]:
        if not hasattr(a[nombre], "iter_rows"): continue
        va = [c.value for f in a[nombre].iter_rows(max_row=40) for c in f]; vb = [c.value for f in b[nombre].iter_rows(max_row=40) for c in f]
        assert [F_norm(x) for x in va] == [F_norm(x) for x in vb], f"cambió la hoja {nombre}"
    return ", ".join(b.sheetnames[:4])

def F_norm(v):
    if hasattr(v, "text"): return "=" + str(v.text)
    return v

def x_borrar_hoja(r, orig, sale):
    import openpyxl
    a, b = openpyxl.load_workbook(orig), openpyxl.load_workbook(sale)
    assert len(b.sheetnames) == len(a.sheetnames) - 1 and a.sheetnames[1] not in b.sheetnames, (a.sheetnames, b.sheetnames)
    return f"{len(a.sheetnames)} -> {len(b.sheetnames)} hojas"

def x_formato(r, orig, sale):
    import openpyxl
    a, b = openpyxl.load_workbook(orig), openpyxl.load_workbook(sale)
    ha, hb = a[a.sheetnames[0]], b[b.sheetnames[0]]
    ref = r["info"]["ref"]
    ca, cb = ha[ref], hb[ref]
    assert F_norm(ca.value) == F_norm(cb.value), f"el valor de {ref} no debía cambiar: {ca.value!r} -> {cb.value!r}"
    assert cb.fill.fill_type == "solid" and cb.fill.fgColor.rgb[-6:].upper() == "FFE599", f"relleno de {ref}: {cb.fill.fgColor.rgb}"
    assert cb.font.name == ca.font.name and cb.font.sz == ca.font.sz, f"la fuente de {ref} cambió: {ca.font.name} {ca.font.sz} -> {cb.font.name} {cb.font.sz}"
    assert cb.number_format == ca.number_format, f"el formato de número cambió: {ca.number_format} -> {cb.number_format}"
    ancho = hb.column_dimensions["C"].width
    assert ancho and abs(ancho - (200 - 5) / 7) < 0.6, f"ancho de C = {ancho}"
    assert hb.freeze_panes == "A3", f"filas fijas: {hb.freeze_panes}"
    return f"{ref} con relleno, misma fuente ({cb.font.name} {cb.font.sz}) y formato {cb.number_format!r}"

DOCX_VARIOS = ["ms-complex2010.docx", "ms-informe-gobernanza.docx", "ms-hoja-informativa.docx", "ms-tabla.docx", "poi-listas.docx", "poi-notas-al-pie.docx", "ms-ecuaciones.docx", "ms-imagenes.docx", "ms-comentarios.docx", "ms-complex2005.docx"]
JS_TEXTO = """
            const sec = doc.secuencia("cuerpo");"""

def _d_resto(orig, sale, minimo=0.985, salvo=()):
    """Todo lo demás del documento sigue igual (las medidas de fidelidad no bajan)."""
    n = F.comparar_docx(orig, sale)
    flojas = {k: f"{v*100:.1f}% ({d})" for k, (v, d) in n.m.items() if v < minimo and k not in salvo}
    assert not flojas, f"se movió algo más: {flojas}"
    return n

def d_texto(r, orig, sale):
    a, b = F._docx_modelo(orig), F._docx_modelo(sale)
    ta = [p["texto"] for p in a["parrafos"]]; tb = [p["texto"] for p in b["parrafos"]]
    if r["info"].get("sinTexto"): return "(sin texto que editar)"
    con = [t for t in tb if "EDITADO-EN-NOM" in t]
    assert len(con) == 1, f"el texto editado aparece {len(con)} veces"
    assert con[0].replace(" EDITADO-EN-NOM ", "") in ta, "el párrafo editado no conserva el resto de su texto"
    assert len(ta) == len(tb), f"cambió el número de párrafos: {len(ta)} -> {len(tb)}"
    distintos = sum(1 for x, y in zip(ta, tb) if x != y)
    assert distintos == 1, f"cambiaron {distintos} párrafos (debía ser 1)"
    # el párrafo editado conserva su formato de párrafo
    i = tb.index(con[0])
    assert a["parrafos"][i]["pa"] == b["parrafos"][i]["pa"], f"cambió el formato del párrafo editado: {a['parrafos'][i]['pa']} -> {b['parrafos'][i]['pa']}"
    n = _d_resto(orig, sale, 0.97, salvo=("texto", "parrafos", "vínculos"))
    return f"1 de {len(tb)} párrafos cambió; resto {n.total()*100:.1f}%"

def d_agregar(r, orig, sale):
    a, b = F._docx_modelo(orig), F._docx_modelo(sale)
    ta = [p["texto"] for p in a["parrafos"]]; tb = [p["texto"] for p in b["parrafos"]]
    assert "Título agregado en NOM" in tb and "Párrafo nuevo con negrita y color." in tb and "Primer punto" in tb and "Segundo punto" in tb and tb[-1].endswith("Cierre"), tb[-6:]
    pt = b["parrafos"][tb.index("Título agregado en NOM")]
    assert pt["pa"].get("titulo") == "1", f"el título no quedó como Título 1: {pt['pa']}"
    pn = b["parrafos"][tb.index("Párrafo nuevo con negrita y color.")]
    k = pn["texto"].index("negrita")
    assert pn["letras"][k].get("b") and not pn["letras"][0].get("b") and pn["letras"][-1].get("color") == "C2372B", "formato de letra del párrafo nuevo"
    assert b["parrafos"][tb.index("Primer punto")]["pa"].get("lista"), "la lista nueva no quedó como lista"
    # todo lo que ya estaba sigue estando, en el mismo orden
    it = iter(tb)
    faltan = [t for t in ta if t.strip() and not any(t == u for u in it)]
    assert not faltan, f"se perdieron párrafos: {faltan[:3]}"
    n = _d_resto(orig, sale, 0.97, salvo=("texto", "parrafos"))
    return f"{len(ta)} -> {len(tb)} párrafos; resto {n.total()*100:.1f}%"

def d_borrar(r, orig, sale):
    a, b = F._docx_modelo(orig), F._docx_modelo(sale)
    ta = [p["texto"] for p in a["parrafos"]]; tb = [p["texto"] for p in b["parrafos"]]
    borrado = r["info"]["borrado"].replace("\ufffc", "").replace("\u2028", "\n")
    if not borrado.strip(): return "(sin texto que borrar)"
    assert ta.count(borrado) - tb.count(borrado) == 1, f"el párrafo borrado sigue: {borrado[:40]!r} ({ta.count(borrado)} -> {tb.count(borrado)})"
    assert len(tb) in (len(ta) - 1, len(ta)), f"párrafos: {len(ta)} -> {len(tb)}"
    _d_resto(orig, sale, 0.97, salvo=("texto", "parrafos", "notas al pie", "vínculos", "imágenes", "ecuaciones"))
    return f"{len(ta)} -> {len(tb)} párrafos"

# ---- MEJ-1: al editar un párrafo de Word no se pierde lo que la app no guarda
DOCX_TODOS = sorted(p.name for p in (F.CORPUS / "docx").glob("*.docx")) if (F.CORPUS / "docx").exists() else DOCX_VARIOS
_NO_PERDER = ["w:bookmarkStart", "w:bookmarkEnd", "w:commentRangeStart", "w:commentRangeEnd", "w:commentReference", "w:fldChar", "w:instrText", "w:hyperlink", "w:footnoteReference",
              "w:endnoteReference", "w:drawing", "mc:AlternateContent", "w:sdt", "w:ins", "w:del", "w:lang", "w:noProof", "w:rStyle", "w:pStyle", "w:numPr", "w:sectPr", "m:oMath", "w:object", "w:pict", "w:smartTag", "w:fldSimple"]

def d_parche(r, orig, sale):
    if not r["info"]["n"]: return "(sin texto que editar)"
    za, zb = zipfile.ZipFile(orig), zipfile.ZipFile(sale)
    otras = [n for n in za.namelist() if n != "word/document.xml" and (n not in zb.namelist() or za.read(n) != zb.read(n))]
    assert not otras, f"cambiaron partes que no se tocaron: {otras[:5]}"
    x, y = za.read("word/document.xml").decode("utf-8"), zb.read("word/document.xml").decode("utf-8")
    cuenta = lambda t, e: len(re.findall("<" + e + r"[\s>/]", t))
    perdidos = {e: f"{cuenta(x, e)} -> {cuenta(y, e)}" for e in _NO_PERDER if cuenta(y, e) < cuenta(x, e)}
    assert not perdidos, f"se perdió al editar: {perdidos} ({r.get('docx')})"
    for atr in ("w:rsidR=", "w:rsidRDefault=", "w14:paraId=", "w:eastAsia=", "w:cs=", "w:themeColor=", "w:asciiTheme="):
        assert y.count(atr) >= x.count(atr), f"se perdieron atributos {atr} {x.count(atr)} -> {y.count(atr)}"
    assert y.count("MEJ1") == r["info"]["insertados"], f"lo escrito aparece {y.count('MEJ1')} veces (debían ser {r['info']['insertados']})"
    st = r.get("docx") or {}
    assert st.get("regeneradas", 0) >= 1, f"no cambió ningún párrafo: {st}"
    assert st.get("parchados", 0) >= max(1, st["regeneradas"] - r["info"].get("tolera", 0)), f"párrafos escritos de nuevo en vez de parchados: {st}"
    F._docx_modelo(sale)
    n = _d_resto(orig, sale, 0.97, salvo=("texto", "parrafos", "letra:b"))
    return f"{st['parchados']} de {st['regeneradas']} párrafos editados salieron parchados sobre su XML original; resto {n.total()*100:.1f}%"

PPTX_VARIOS = ["ms-tipico.pptx", "ms-brownbag.pptx", "ms-of16.pptx", "ms-animacion.pptx", "ms-organigrama.pptx", "ms-tabla-grande.pptx", "poi-temas.pptx", "poi-formas.pptx", "ms-presentation.pptx"]

def _textos(pres):
    from pptx import Presentation
    r = []
    for sl in pres.slides:
        ts = []
        def rec(shapes):
            for s in shapes:
                if s.shape_type == 6 and hasattr(s, "shapes"): rec(s.shapes); continue
                ts.append((s.text_frame.text if getattr(s, "has_text_frame", False) and s.has_text_frame else "", s.shape_type, s.left, s.top, s.width, s.height))
        rec(sl.shapes); r.append(ts)
    return r

def _p_resto(orig, sale, salvo=(), minimo=0.9):
    n = F.comparar_pptx(orig, sale)
    flojas = {k: f"{v*100:.1f}% ({d})" for k, (v, d) in n.m.items() if v < minimo and k not in salvo}
    assert not flojas, f"se movió algo más: {flojas}"
    return n

def p_texto(r, orig, sale):
    from pptx import Presentation
    if r["info"].get("sinTexto"): return "(sin cuadros de texto)"
    a, b = Presentation(orig), Presentation(sale)
    ta, tb = _textos(a), _textos(b)
    pg = r["info"]["pagina"]
    assert len(ta) == len(tb), f"diapositivas {len(ta)} -> {len(tb)}"
    assert any("Texto cambiado en NOM" in t[0] for t in tb[pg]), "no aparece el texto nuevo"
    for i, (x, y) in enumerate(zip(ta, tb)):
        if i == pg:
            assert len(x) == len(y), f"en la diapositiva editada cambió el número de figuras: {len(x)} -> {len(y)}"
            dif = sum(1 for p, q in zip(sorted(t[0] for t in x), sorted(t[0] for t in y)) if p != q)
            assert dif <= 2, f"cambiaron {dif} textos en la diapositiva editada"
        else:
            assert x == y, f"cambió la diapositiva {i + 1}, que no se tocó"
    n = _p_resto(orig, sale, salvo=("texto", "negrita y cursiva", "tamaño de letra", "color de letra", "alineación del texto"))
    return f"diapositiva {pg + 1} de {len(tb)}; resto {n.total()*100:.1f}%"

def p_mover(r, orig, sale):
    from pptx import Presentation
    if r["info"].get("sinFigura"): return "(sin figuras)"
    a, b = Presentation(orig), Presentation(sale)
    ta, tb = _textos(a), _textos(b)
    pg, tam = r["info"]["pagina"], r["info"]["tam"]
    ex, ey = round(37 / tam["w"] * a.slide_width), round(41 / tam["h"] * a.slide_height)
    tol = a.slide_width * 0.004
    assert any(t[2] is not None and abs(t[2] - ex) < tol and abs(t[3] - ey) < tol for t in tb[pg]), f"ninguna figura quedó en ({ex}, {ey}): {[(t[2], t[3]) for t in tb[pg]][:6]}"
    for i, (x, y) in enumerate(zip(ta, tb)):
        if i != pg: assert x == y, f"cambió la diapositiva {i + 1}, que no se tocó"
        else: assert len(y) >= len(x), f"figuras en la diapositiva editada: {len(x)} -> {len(y)}"
    s0 = b.slides[0]
    assert s0.has_notes_slide and "Nota escrita en NOM-COLAB" in s0.notes_slide.notes_text_frame.text, "las notas no quedaron"
    n = _p_resto(orig, sale, salvo=("posición y tamaño", "notas del orador"))
    return f"figura movida en la {pg + 1}; notas puestas; resto {n.total()*100:.1f}%"

def p_agregar(r, orig, sale):
    from pptx import Presentation
    a, b = Presentation(orig), Presentation(sale)
    ta, tb = _textos(a), _textos(b)
    antes, quito = r["info"]["antes"], r["info"]["quito"]
    assert len(tb) == antes + (0 if quito else 1), f"diapositivas {len(ta)} -> {len(tb)}"
    ult = tb[-1]
    assert any("Diapositiva nueva de NOM" in t[0] for t in ult) and any("Círculo" in t[0] for t in ult), f"la diapositiva nueva no trae su contenido: {[t[0] for t in ult]}"
    for i in range(len(tb) - 1):
        assert ta[i] == tb[i], f"cambió la diapositiva {i + 1}, que no se tocó"
    return f"{len(ta)} -> {len(tb)} diapositivas"

# ---- DEF-45: lo que no debe perderse al editar una diapositiva importada (animaciones, marcadores, números)
PPTX_ANIM = ["ms-animacion.pptx", "ms-of16.pptx", "ms-tipico.pptx", "ms-brownbag.pptx"]
_RE_ID = re.compile(r'<p:cNvPr\b[^>]*?\sid="(\d+)"')

def _diapos(ruta):
    z = zipfile.ZipFile(ruta)
    return {n: z.read(n).decode("utf-8") for n in z.namelist() if re.match(r"ppt/slides/[^/]+\.xml$", n)}

def _timing(x):
    m = re.search(r"<p:timing\b.*?</p:timing>|<p:timing\b[^>]*/>", x, re.S)
    return m.group(0) if m else ""

def _sin_colgar(x, nombre):
    """Ninguna animación apunta a una figura (o párrafo) que no esté en la diapositiva."""
    ids = set(_RE_ID.findall(x)); t = _timing(x)
    colgadas = sorted(set(re.findall(r'spid="(\d+)"', t)) - ids)
    assert not colgadas, f"{nombre}: animaciones que apuntan a figuras que ya no están: {colgadas}"
    for spid, fin in re.findall(r'<p:spTgt spid="(\d+)">\s*<p:txEl>\s*<p:pRg st="\d+" end="(\d+)"', t):
        i = x.find(f'id="{spid}"'); j = x.find("</p:sp>", i)
        n = len(re.findall(r"<a:p[\s>/]", x[i:j]))
        assert int(fin) < n, f"{nombre}: animación del párrafo {fin} de la figura {spid}, que solo tiene {n}"

def _otras_partes_iguales(orig, sale, tocadas):
    a, b = zipfile.ZipFile(orig), zipfile.ZipFile(sale)
    distintas = [n for n in a.namelist() if n not in tocadas and (n not in b.namelist() or a.read(n) != b.read(n))]
    assert not distintas, f"cambiaron partes que no se tocaron: {distintas[:5]}"

def p_solo_formato(r, orig, sale):
    from pptx import Presentation
    if not r["info"]["rutas"]: return "(sin títulos)"
    a, b = _diapos(orig), _diapos(sale)
    tocadas = set(r["info"]["rutas"])
    _otras_partes_iguales(orig, sale, tocadas)
    con_b0 = 0
    for n in tocadas:
        x, y = a[n], b[n]
        assert _timing(x) == _timing(y), f"{n}: cambiaron (o se perdieron) las animaciones"
        assert _RE_ID.findall(x) == _RE_ID.findall(y), f"{n}: cambiaron los números de las figuras: {_RE_ID.findall(x)} -> {_RE_ID.findall(y)}"
        for marca in ("<p:ph", "<a:endParaRPr", "smtClean", 'lang="', "<p:spPr", "txBox=", "<a:xfrm", "<a:latin", "<a:solidFill"):
            assert x.count(marca) == y.count(marca), f"{n}: {marca} {x.count(marca)} -> {y.count(marca)}"
        # quitando el atributo de negrita, la diapositiva tiene que ser la misma letra por letra
        sin = lambda t: re.sub(r"<a:rPr/>", "", re.sub(r'\sb="[01]"', "", t))
        assert sin(x) == sin(y), f"{n}: cambió algo más que la negrita"
        con_b0 += y.count(' b="0"') - x.count(' b="0"')
    assert con_b0 > 0, "la negrita no quedó escrita"
    Presentation(sale)
    return f"{len(tocadas)} diapositiva(s): solo cambió b=\"0\"; animaciones, marcadores y números intactos"

def p_titulo_formato(r, orig, sale):
    from pptx import Presentation
    if not r["info"]["rutas"]: return "(sin títulos)"
    a, b = _diapos(orig), _diapos(sale)
    tocadas = set(r["info"]["rutas"])
    _otras_partes_iguales(orig, sale, tocadas)
    for n in tocadas:
        x, y = a[n], b[n]
        _sin_colgar(y, n)
        assert _RE_ID.findall(x) == _RE_ID.findall(y), f"{n}: cambiaron los números de las figuras"
        assert x.count("<p:ph") == y.count("<p:ph"), f"{n}: marcadores {x.count('<p:ph')} -> {y.count('<p:ph')}"
        assert x.count("txBox=") == y.count("txBox="), f"{n}: un marcador se volvió cuadro de texto suelto"
        assert ("<p:timing" in x) == ("<p:timing" in y), f"{n}: se perdieron las animaciones"
        assert "Título nuevo en NOM" in y and 'val="C2372B"' in y and 'sz="' in y, f"{n}: no quedó el texto, el color o el tamaño"
    Presentation(sale)
    return f"{len(tocadas)} diapositiva(s): título con texto, color y tamaño nuevos; siguen siendo marcadores, con sus animaciones"

def p_rehacer(r, orig, sale):
    from pptx import Presentation
    a, b = _diapos(orig), _diapos(sale)
    con = 0
    for n, y in b.items():
        _sin_colgar(y, n)
        if n in a and "<p:timing" in a[n] and "<p:timing" in y: con += 1
    for n in r["info"]["soloColor"]:
        # diapositivas donde solo se cambió el relleno de figuras sueltas: mismas figuras, mismos números
        assert sorted(_RE_ID.findall(a[n])) == sorted(_RE_ID.findall(b[n])), f"{n}: cambiaron los números de las figuras: {_RE_ID.findall(a[n])} -> {_RE_ID.findall(b[n])}"
        assert _timing(a[n]) == _timing(b[n]), f"{n}: cambiaron las animaciones"
        assert a[n].count("<p:ph") == b[n].count("<p:ph"), f"{n}: se perdieron marcadores"
    Presentation(sale)
    return f"{r['info']['color']} figura(s) con otro relleno, {r['info']['borradas']} borrada(s); {con} diapositiva(s) conservan animaciones y ninguna apunta al vacío"

CASOS = {
    "xlsx": [
        ("cambiar celdas y agregar una fórmula", ["ms-financial-sample.xlsx", "ms-complex01.xlsx", "ms-excel14.xlsx", "poi-formato-condicional.xlsx", "ms-dinamica-b1.xlsx", "poi-graficas.xlsx"], """
            const h = "hoja-p1";
            doc.cambiar(h, { "c:B2": "Editado en NOM-COLAB", "n:B2": true, "c:AZ300": "1234.5", "c:AZ301": "=SUMA(AZ300:AZ300)*2" });
            return {};""", x_valor),
        ("renombrar una hoja y agregar otra", ["ms-financial-sample.xlsx", "ms-hojas.xlsx", "ms-complex01.xlsx", "poi-estilos.xlsx"], """
            doc.cambiar("hoja-p1", { nombre: "Renombrada por mí" });
            const pags = doc.paginas().concat(["pz"]);
            doc.cambiar("_doc", { paginas: pags });
            doc.cambiar("hoja-pz", { tipo: "tabla", pagina: "pz", orden: 0, x: 40, y: 40, w: 880, h: 300, filas: 30, cols: 10, nombre: "Hoja nueva", "c:A1": "Total", "n:A1": true, "c:B1": "42" });
            return {};""", x_hojas),
        ("borrar una hoja", ["ms-hojas.xlsx", "ms-complex01.xlsx"], """
            doc.cambiar("hoja-p2", { eliminado: true });
            doc.cambiar("_doc", { paginas: doc.paginas().filter((p) => p !== "p2") });
            return {};""", x_borrar_hoja),
        ("solo formato: relleno, ancho de columna y filas fijas", ["ms-financial-sample.xlsx", "ms-rendimiento.xlsx", "poi-estilos.xlsx"], """
            const b = doc.bloque("hoja-p1");
            const ref = Object.keys(b).filter((k) => k.startsWith("c:") && /^c:[A-F]\\d$/.test(k)).sort()[0].slice(2);
            doc.cambiar("hoja-p1", { ["r:" + ref]: "#FFE599", "aw:C": 200, fijasF: 2 });
            return { ref };""", x_formato),
    ],
    "docx": [
        ("cambiar el texto de un párrafo", DOCX_VARIOS, JS_TEXTO + """
            const vis = sec.visibles();
            // el primer párrafo con al menos 12 letras seguidas
            let i = 0, n = 0;
            for (; i < vis.length; i++) { if (vis[i].e || vis[i].c === "\\n") n = 0; else if (++n >= 12) break; }
            if (i >= vis.length) { i = 0; n = 0; for (; i < vis.length; i++) { if (vis[i].e || vis[i].c === "\\n") n = 0; else if (++n >= 2) break; } }
            if (i >= vis.length) return { sinTexto: true };
            doc.seq("cuerpo", { k: "i", iz: vis[i].id, x: " EDITADO-EN-NOM " });
            return { antes: vis.length };""", d_texto),
        ("agregar al final un título, un párrafo y una lista", DOCX_VARIOS, JS_TEXTO + """
            const vis = sec.visibles();
            let iz = vis.length ? vis[vis.length - 1].id : null;
            const pon = (x, a) => { const r = doc.seq("cuerpo", Object.assign({ k: "i", iz, x }, a ? { a } : {})); const u = doc.secuencia("cuerpo").visibles(); iz = u[u.length - 1].id; };
            // Enter al final, como en el editor: la marca nueva copia el formato del párrafo
            pon("\\n", Object.assign({}, NOM.Secuencia.attrs(sec.fin))); pon("Título agregado en NOM"); pon("\\n", { estilo: "h1" });
            pon("Párrafo nuevo con "); pon("negrita", { b: true }); pon(" y color.", { color: "#C2372B" }); pon("\\n");
            pon("Primer punto"); pon("\\n", { lista: "numeros" }); pon("Segundo punto"); pon("\\n", { lista: "numeros" });
            pon("Cierre");
            return {};""", d_agregar),
        ("borrar el primer párrafo con texto", DOCX_VARIOS, JS_TEXTO + """
            const vis = sec.visibles();
            let a = 0; while (a < vis.length && (vis[a].c === "\\n" || vis[a].e)) a++;
            let b = a; while (b < vis.length && vis[b].c !== "\\n") b++;
            const borrado = vis.slice(a, b).map((x) => x.c).join("");
            doc.seq("cuerpo", { k: "v", v: false, r: NOM.Secuencia.rangos(vis.slice(a, Math.min(b + 1, vis.length))) });
            return { borrado };""", d_borrar),
        ("editar varios párrafos como al teclear, sin perder lo que la app no guarda (MEJ-1)", DOCX_TODOS, JS_TEXTO + """
            const vis = sec.visibles(), ini = [];
            // dónde termina la letra 12 de cada párrafo que tenga al menos 12 seguidas
            let n = 0;
            for (let i = 0; i < vis.length; i++) { if (vis[i].e || vis[i].c === "\\n") n = 0; else if (++n === 12) ini.push(i); }
            const paso = Math.max(1, Math.floor(ini.length / 6)), elegidos = ini.filter((_, k) => k % paso === 0).slice(0, 6);
            let insertados = 0;
            // de atrás para adelante, cada párrafo con una edición distinta: escribir, borrar y poner negrita
            elegidos.slice().reverse().forEach((i, k) => {
              if (k % 3 === 0) { doc.seq("cuerpo", { k: "i", iz: vis[i].id, x: " MEJ1 ", a: Object.assign({}, NOM.Secuencia.attrs(vis[i])) }); insertados++; }
              else if (k % 3 === 1) doc.seq("cuerpo", { k: "v", v: false, r: NOM.Secuencia.rangos(vis.slice(i - 3, i)) });
              else doc.seq("cuerpo", { k: "f", a: { b: true }, r: NOM.Secuencia.rangos(vis.slice(i - 6, i)) });
            });
            return { n: elegidos.length, insertados };""", d_parche),
    ],
    "pptx": [
        ("cambiar el texto de un cuadro", PPTX_VARIOS, """
            const b = doc.vivos().filter((x) => x.tipo === "texto" && x.html && x.pagina).sort((x, y) => (doc.paginas().indexOf(x.pagina) - doc.paginas().indexOf(y.pagina)) || (x.z || 0) - (y.z || 0))[0];
            if (!b) return { sinTexto: true };
            doc.cambiar(b.id, { html: "<div><b>Texto cambiado en NOM</b> y más</div>" });
            return { pagina: doc.paginas().indexOf(b.pagina) };""", p_texto),
        ("mover una figura y poner notas", PPTX_VARIOS, """
            const b = doc.vivos().filter((x) => x.pagina && x.x != null && (x.tipo === "forma" || x.tipo === "imagen" || x.tipo === "texto"))[0];
            if (!b) return { sinFigura: true };
            doc.cambiar(b.id, { x: 37, y: 41 });
            doc.cambiar("_doc", { ["nota:" + doc.paginas()[0]]: "Nota escrita en NOM-COLAB" });
            return { pagina: doc.paginas().indexOf(b.pagina), tam: NOM.diapo.tam(doc) };""", p_mover),
        ("agregar una diapositiva y borrar otra", PPTX_VARIOS, """
            const pags = doc.paginas(), quitada = pags.length > 1 ? pags[pags.length - 1] : null;
            if (quitada) doc.deLaPagina(quitada).forEach((b) => doc.cambiar(b.id, { eliminado: true }));
            const nuevas = pags.filter((p) => p !== quitada).concat(["pz"]);
            doc.cambiar("_doc", { paginas: nuevas });
            doc.cambiar("nz1", { tipo: "texto", pagina: "pz", orden: 1, z: 1, x: 80, y: 60, w: 600, h: 80, estilo: "h1", html: "Diapositiva nueva de NOM" });
            doc.cambiar("nz2", { tipo: "forma", figura: "elipse", pagina: "pz", orden: 2, z: 2, x: 200, y: 220, w: 220, h: 140, color: "#F06A3D", texto: "Círculo" });
            return { antes: pags.length, quito: !!quitada };""", p_agregar),
        ("solo formato: quitar la negrita de los títulos (DEF-45)", PPTX_ANIM, """
            // lo que hace «Negrita» de la ruedita sobre un título (que se ve en negrita): se la quita a todo el cuadro
            const ts = doc.vivos().filter((x) => x.tipo === "texto" && (x.estilo === "h1" || x.estilo === "h2") && x.pagina && x.ox != null);
            ts.forEach((b) => doc.cambiar(b.id, { tNegrita: false }));
            return { rutas: [...new Set(ts.map((b) => doc.valor("_doc", "os:" + b.pagina)))] };""", p_solo_formato),
        ("título con texto, color y tamaño nuevos (DEF-45)", PPTX_ANIM, """
            const vistos = new Set(), ts = doc.vivos().filter((x) => x.tipo === "texto" && (x.estilo === "h1" || x.estilo === "h2") && x.pagina && x.ox != null).filter((b) => !vistos.has(b.pagina) && vistos.add(b.pagina));
            ts.forEach((b) => doc.cambiar(b.id, { html: "<div>Título nuevo en NOM</div>", tColor: "#C2372B", tTam: 48 }));
            return { rutas: ts.map((b) => doc.valor("_doc", "os:" + b.pagina)) };""", p_titulo_formato),
        ("otro relleno en las figuras y borrar la segunda de cada diapositiva (DEF-45)", PPTX_ANIM, """
            let color = 0, borradas = 0;
            const soloColor = [];
            doc.paginas().forEach((pg, n) => {
                const bs = doc.deLaPagina(pg).filter((b) => b.ox != null).sort((x, y) => (x.z || 0) - (y.z || 0));
                const unicos = bs.filter((b) => bs.filter((c) => c.ox === b.ox).length === 1);
                const formas = unicos.filter((b) => b.tipo === "forma" && b.figura !== "linea" && b.figura !== "linea-flecha");
                formas.forEach((b) => { doc.cambiar(b.id, { color: "#F06A3D" }); color++; });
                // en las diapositivas pares, además, se borra su segunda figura (suele ser la que tiene animación)
                if (n % 2 === 1 && unicos.length >= 2) { doc.cambiar(unicos[1].id, { eliminado: true }); borradas++; }
                else if (formas.length && formas.length === bs.filter((b) => b.tipo === "forma").length) soloColor.push(doc.valor("_doc", "os:" + pg));
            });
            return { color, borradas, soloColor };""", p_rehacer),
    ],
}


# LibreOffice se queda esperando con el ORIGINAL de Microsoft de estos (comprobado el 15-nov-2026: 2 de 2
# veces, 200 s cada una, sin tocarlo). Pasarlos por LibreOffice no dice nada de la exportación y gastaba
# 6 minutos por caso (y a veces el original sí abría y daba una falla falsa). Los revisa el validador de
# Microsoft (validar_openxml.py), que sí los lee.
LO_SE_TRABA = {"ms-brownbag.pptx"}


def abre_office(ruta):
    """LibreOffice tiene que poder abrir el archivo y volverlo a guardar."""
    carpeta = SALIDA / "lo"; carpeta.mkdir(parents=True, exist_ok=True)
    ext = Path(ruta).suffix[1:]
    destino = carpeta / (Path(ruta).stem + ".pdf")
    if destino.exists(): destino.unlink()
    # LibreOffice se queda esperando con algunos archivos (también con el original de Microsoft, p. ej.
    # ms-brownbag.pptx): se corre en su propio grupo y, si no termina, se mata todo el grupo; si no,
    # soffice.bin seguía vivo con el perfil bloqueado y todas las conversiones siguientes se trababan.
    import os, signal
    pr = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{F.CACHE / 'lo-perfil'}", "--headless", "--convert-to", "pdf", "--outdir", str(carpeta), str(ruta)],
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
    try: pr.wait(timeout=180)
    except subprocess.TimeoutExpired:
        os.killpg(pr.pid, signal.SIGKILL); pr.wait()
        return None
    return destino.exists() and destino.stat().st_size > 800


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--solo", choices=["docx", "xlsx", "pptx"])
    ap.add_argument("--office", action="store_true", help="abrir también cada resultado con LibreOffice")
    a = ap.parse_args()
    SALIDA.mkdir(parents=True, exist_ok=True)
    from playwright.sync_api import sync_playwright
    bien = mal = 0
    with sync_playwright() as pw:
        nav = pw.chromium.launch(); pag = nav.new_page()
        errores = []
        pag.on("pageerror", lambda e: errores.append(str(e)))
        pag.goto(F.BASE + "/"); pag.wait_for_function("() => window.NOM && NOM.ooxml && NOM.formatos && NOM.formatos.docDeOps", timeout=30000)
        for ext, casos in CASOS.items():
            if a.solo and a.solo != ext: continue
            for titulo, archivos, edicion, revisar in casos:
                print(f"== {ext}: {titulo}")
                for nombre in archivos:
                    orig = F.CORPUS / ext / nombre
                    if not orig.exists(): print(f"   (falta {nombre})"); continue
                    r = pag.evaluate(JS.replace("/*EDICION*/", edicion), [base64.b64encode(orig.read_bytes()).decode(), nombre])
                    if r.get("error"): mal += 1; print(f"   MAL  {nombre}: {r['error'][:300]}"); continue
                    for sufijo, clave in (("", "salida"), (".desde-nom", "salidaNom")):
                        sale = SALIDA / (Path(nombre).stem + "." + re.sub(r"\W+", "-", titulo)[:24] + sufijo + "." + ext)
                        sale.write_bytes(base64.b64decode(r[clave]))
                        try:
                            assert zipfile.ZipFile(sale).testzip() is None, "zip dañado"
                            det = revisar(r, orig, sale)
                            if a.office and not sufijo and nombre not in LO_SE_TRABA:
                                r_lo = abre_office(sale)
                                # Si LibreOffice se traba, se prueba con el original de Microsoft: si con ese también, no es culpa de la exportación.
                                if r_lo is None: assert abre_office(orig) is None, "LibreOffice se trabó con lo exportado y no con el original"
                                else: assert r_lo, "LibreOffice no pudo abrirlo"
                            bien += 1
                            print(f"   ok   {nombre}{' (vía .nom)' if sufijo else ''}: {det}; {orig.stat().st_size//1024} KB -> {sale.stat().st_size//1024} KB{'' if sufijo else f', .nom {r['nom']//1024} KB'}")
                        except AssertionError as e:
                            mal += 1; print(f"   MAL  {nombre}{' (vía .nom)' if sufijo else ''}: {e}")
                        except Exception as e:
                            mal += 1; print(f"   MAL  {nombre}{' (vía .nom)' if sufijo else ''}: {type(e).__name__}: {str(e)[:260]}")
        nav.close()
    print(f"\nEdiciones: {bien} bien, {mal} mal. Errores de JavaScript: {errores[:3]}")
    sys.exit(1 if mal else 0)


if __name__ == "__main__":
    main()

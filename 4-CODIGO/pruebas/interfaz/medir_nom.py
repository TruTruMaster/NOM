#!/usr/bin/env python3
"""Cuánto pesa un .nom y por qué (pregunta del dueño, 26-nov: «cada letra un bit y cada formato un byte»).

Mide en el sitio real, con el mismo código que el botón «Descargar → NOM»: un documento vacío de cada tipo, uno de
texto con 1 000 letras escritas de corrido, el mismo con 1 000 letras y 50 cambios de formato, y los archivos de
Office que se pasen. Para cada uno: bytes del .nom, bytes del JSON sin comprimir, cuánto es imagen/Office original,
y bytes por letra.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/medir_nom.py SALIDA [ARCHIVO.docx …]
"""
import os, sys, json, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina
from playwright.sync_api import sync_playwright

SALIDA = sys.argv[1]; ARCH = sys.argv[2:]
os.makedirs(SALIDA, exist_ok=True)
C = N.cuenta(); esp = N.espacio(C["token"], "Medir")
MEDIR = r"""() => {
  const d = NOM.editorActual.doc, b = NOM.formatos.aNomBytes(d);
  const foto = NOM.formatos.fotoModelo(d), j = new TextEncoder().encode(JSON.stringify(foto)).length;
  const og = NOM.origen.leer(d);
  let letras = 0; try { letras = (NOM.editorActual.vista.querySelector('.wd-cuerpo, [contenteditable]') || {}).innerText.length || 0; } catch (e) {}
  return { nom: b.length, json: j, conOrigen: !!og, letras };
}"""
res = {}
with sync_playwright() as pw:
    br = pw.chromium.launch(); p = br.new_context(viewport={"width": 1366, "height": 900}).new_page(); P = Pagina(p, N.B); P.entrar(C)
    for modo in ("TEXTO", "HOJA", "PRESENTACION", "DIAGRAMA"):
        d = N.documento(C["token"], esp, modo, "Vacío " + modo); P.abrir(d); P.t(1500)
        res["vacío " + modo.lower()] = P.ev(MEDIR)
    d = N.documento(C["token"], esp, "TEXTO", "Mil letras"); P.abrir(d); P.t(1500)
    P.ev("() => { const c = NOM.editorActual.vista.querySelector('[contenteditable]'); c && c.focus(); }")
    # texto de verdad, sin repetirse (un texto repetido se comprime de más y daría un número engañoso)
    guia = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "docs", "GUIA-FACIL-DE-USAR.md"), encoding="utf-8").read()
    import re as _re
    texto = _re.sub(r"[^A-Za-záéíóúñÁÉÍÓÚÑüÜ0-9 ,.;:¿?¡!()]", " ", guia); texto = _re.sub(r" +", " ", texto)[2000:3000]
    p.keyboard.type(texto, delay=0); P.t(2500)
    res["texto 1000 letras"] = P.ev(MEDIR)
    # 50 cambios de formato (negrita en palabras sueltas)
    for i in range(50):
        p.keyboard.press("Control+Home"); [p.keyboard.press("Control+ArrowRight") for _ in range(i * 3)]
        p.keyboard.press("Control+Shift+ArrowRight"); p.keyboard.press("Control+b")
    P.t(2500)
    res["texto 1000 letras + 50 negritas"] = P.ev(MEDIR)
    br.close()
    for a in ARCH:
        b2 = pw.chromium.launch(); p = b2.new_context(viewport={"width": 1366, "height": 900}).new_page(); P = Pagina(p, N.B); P.entrar(C)
        p.goto(N.B + "/#/espacio/" + esp); P.t(1500); P.cerrar_hojas()
        P.ev("() => { const b = [...document.querySelectorAll('button, a')].find(x => /Nuevo documento|Crear documento/.test(x.textContent) && x.offsetParent); if (b) b.click(); }"); P.t(800)
        with p.expect_file_chooser(timeout=15000) as fc: P.ev("() => document.querySelector('[data-importar]').click()")
        fc.value.set_files(a); P.t(800)
        P.ev("() => { const b = [...document.querySelectorAll('.velo button, .hoja button')].find(x => /^En «/.test(x.textContent.trim())); b && b.click(); }")
        p.wait_for_function("() => location.hash.startsWith('#/doc/') && window.NOM && NOM.editorActual && NOM.editorActual.vista && NOM.editorActual.vista.isConnected", timeout=180000); P.t(4000)
        r = P.ev(MEDIR); r["office"] = os.path.getsize(a); res[os.path.basename(a)] = r
        b2.close()
for k, v in res.items():
    extra = f" · {v['nom'] / v['letras']:.2f} bytes por letra" if v.get("letras") and not v.get("conOrigen") else ""
    print(f"{k:38s} .nom {v['nom']:>9,} B · JSON sin comprimir {v['json']:>10,} B" + (f" · Office original {v['office']:,} B" if v.get("office") else "") + extra)
json.dump(res, open(os.path.join(SALIDA, "medidas.json"), "w"), ensure_ascii=False, indent=1)

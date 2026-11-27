#!/usr/bin/env python3
"""Lo más visible de una demostración, hecho con clics como una persona (lo que la matriz de la expo marcó sin prueba):

  1. «Importar archivo» desde el botón del espacio (un .docx de verdad) → se abre el documento.
  2. «Descargar» de la barra del editor → Word, Excel y PowerPoint; cada archivo se abre con su biblioteca.
  3. «Invitar» → «O agrega a alguien que ya tiene cuenta» con «Solo ver» → esa persona entra y no puede escribir;
     luego el propietario le cambia a «Puede editar» en la lista de miembros y ya puede.
  4. «Buscar en los documentos» del espacio encuentra un documento por lo que dice.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_demo.py SALIDA [--movil]

Crea dos cuentas desechables y las borra al final.
"""
import io, os, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import nomprueba as N
from prueba_funciones import Pagina, check, BIEN, MAL
from playwright.sync_api import sync_playwright
from docx import Document

SALIDA = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else "/tmp/demo"
MOVIL = "--movil" in sys.argv
os.makedirs(SALIDA, exist_ok=True)
ruta_docx = os.path.join(SALIDA, "informe-demo.docx")
d = Document(); d.add_heading("Informe de la práctica", 1); d.add_paragraph("Medimos la temperatura del agua cada cinco minutos.")
t = d.add_table(rows=2, cols=2); t.cell(0, 0).text = "Minuto"; t.cell(0, 1).text = "Grados"; d.save(ruta_docx)

c1, c2 = N.cuenta(), N.cuenta()
esp = N.espacio(c1["token"], "Demo " + str(int(time.time()) % 100000))
hoja = N.documento(c1["token"], esp, "HOJA", "Cuentas del grupo")
pres = N.documento(c1["token"], esp, "PRESENTACION", "Exposición")

with sync_playwright() as pw:
    b = pw.chromium.launch()
    opc = dict(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True, device_scale_factor=2) if MOVIL else dict(viewport={"width": 1366, "height": 800})
    ctx = b.new_context(accept_downloads=True, **opc); p = ctx.new_page(); P = Pagina(p, N.B); P.entrar(c1)

    # 1. Importar desde el botón del espacio
    print("== Importar archivo", flush=True)
    p.goto(N.B + "/#/espacio/" + esp); P.t(1500); P.cerrar_hojas()
    abrir_nuevo = P.ev("() => { const b = [...document.querySelectorAll('button, a')].find(x => /Nuevo documento|Crear documento/.test(x.textContent) && x.offsetParent); if (b) { b.click(); return true; } return false; }")
    if not abrir_nuevo:
        P.ev("() => { const b = document.querySelector('[aria-label=\"Nuevo documento\"], .fab'); b && b.click(); }")
    P.t(800)
    with p.expect_file_chooser(timeout=15000) as fc:
        P.ev("() => document.querySelector('[data-importar]').click()")
    fc.value.set_files(ruta_docx)
    try:
        p.wait_for_function("() => location.hash.startsWith('#/doc/') && window.NOM && NOM.editorActual && NOM.editorActual.vista && NOM.editorActual.vista.isConnected", timeout=60000)
        ok = True
    except Exception: ok = False
    P.t(1200)
    txt = P.ev("() => document.body.innerText")
    check("«Importar archivo» abre el Word importado", ok and "temperatura del agua" in txt, txt[:200])
    P.quitar_consejo()

    # 2. Descargar a Office desde la barra
    print("== Descargar", flush=True)
    def descargar(doc_id, fmt):
        if doc_id: P.abrir(doc_id)
        P.cerrar_hojas()
        P.barra("Descargar"); P.t(600)
        with p.expect_download(timeout=60000) as dl:
            P.clic(f'.velo .hoja [data-f="{fmt}"]')
        ruta = os.path.join(SALIDA, "descargado." + fmt); dl.value.save_as(ruta); P.t(400); P.cerrar_hojas()
        return ruta
    try:
        r = descargar(None, "docx")
        from docx import Document as D2
        check("«Descargar» → Word: se abre y trae el texto", "temperatura del agua" in "\n".join(x.text for x in D2(r).paragraphs), r)
    except Exception as e: check("«Descargar» → Word", False, str(e)[:200])
    try:
        r = descargar(hoja, "xlsx")
        import openpyxl
        check("«Descargar» → Excel: se abre", len(openpyxl.load_workbook(r).sheetnames) >= 1, r)
    except Exception as e: check("«Descargar» → Excel", False, str(e)[:200])
    try:
        r = descargar(pres, "pptx")
        from pptx import Presentation
        check("«Descargar» → PowerPoint: se abre", len(Presentation(r).slides) >= 1, r)
    except Exception as e: check("«Descargar» → PowerPoint", False, str(e)[:200])

    # 3. Agregar con «Solo ver» y cambiar a «Puede editar»
    print("== Permisos desde la pantalla", flush=True)
    p.goto(N.B + "/#/espacio/" + esp); P.t(1500); P.cerrar_hojas()
    P.boton("Invitar", "main"); P.t(800)
    P.ev("(n) => { const i = document.querySelector('.velo .hoja [data-agregar] input[name=usuario]'); i.value = n; i.dispatchEvent(new Event('input', {bubbles: true})); }", c2["nombre"])
    P.clic('.velo .hoja [data-agregar] [data-p="LECTURA"]')
    P.ev("() => document.querySelector('.velo .hoja [data-agregar] .btn').click()"); P.t(1500)
    acc = N.api("GET", f"/api/espacios/{esp}/accesos", None, c1["token"])[1]
    per = [a.get("permiso") for a in (acc if isinstance(acc, list) else acc.get("accesos", acc.get("miembros", []))) if a.get("nombre") == c2["nombre"]]
    check("«Agregar» con «Solo ver» deja a la persona como «Solo puede ver»", per == ["LECTURA"], [per, str(acc)[:200]])
    opc2 = dict(viewport={"width": 1280, "height": 860}) if MOVIL else dict(viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True, device_scale_factor=2)
    ctx2 = b.new_context(**opc2); B = Pagina(ctx2.new_page(), N.B); B.entrar(c2)
    docs = N.api("GET", f"/api/espacios/{esp}/documentos", None, c2["token"])[1]
    docs = docs.get("documentos", docs) if isinstance(docs, dict) else docs
    wid = [x["id"] for x in docs if x.get("modo") == "TEXTO"][0]
    B.abrir(wid); B.t(800)
    check("…y al abrir el texto la barra dice «Solo puedes ver»", "Solo puedes ver" in B.ev("() => document.body.innerText"), B.ev("() => document.body.innerText")[:200])
    P.cerrar_hojas(); p.goto(N.B + "/#/espacio/" + esp); P.t(1500); P.cerrar_hojas()
    P.boton("Invitar", "main"); P.t(800)
    P.ev("(n) => { const s = document.querySelector(`.velo .hoja select[data-cambiar=\"${n}\"]`); s.value = 'EDICION'; s.dispatchEvent(new Event('change', {bubbles: true})); }", c2["nombre"]); P.t(1500)
    acc = N.api("GET", f"/api/espacios/{esp}/accesos", None, c1["token"])[1]
    per = [a.get("permiso") for a in (acc if isinstance(acc, list) else acc.get("accesos", acc.get("miembros", []))) if a.get("nombre") == c2["nombre"]]
    check("el propietario cambia a «Puede editar» en la lista de miembros", per == ["EDICION"], per)
    B.p.reload(); B.t(800); B.abrir(wid); B.t(800)
    check("…y la otra persona ya puede escribir", not "Solo puedes ver" in B.ev("() => document.body.innerText") and B.ev("() => !!document.querySelector('[contenteditable=true]')"), B.ev("() => document.body.innerText")[:160])
    P.cerrar_hojas()

    # 4. Buscar en los documentos del espacio
    print("== Buscar", flush=True)
    p.goto(N.B + "/#/espacio/" + esp); P.t(1500); P.cerrar_hojas()
    P.ev("() => { const i = document.querySelector('[data-buscar]'); i.value = 'temperatura'; i.dispatchEvent(new Event('input', {bubbles: true})); }")
    try: p.wait_for_function("() => /Informe|informe-demo/.test(document.querySelector('main').innerText) && document.querySelectorAll('.doc-tarjeta, .resultado, [data-resultado]').length > 0", timeout=15000); hay = True
    except Exception: hay = False
    p.screenshot(path=os.path.join(SALIDA, "buscar.png"))
    check("«Buscar en los documentos» encuentra el informe por lo que dice («temperatura»)", hay, P.ev("() => document.querySelector('main').innerText")[:300])
    check("sin errores de JavaScript", not P.errores and not B.errores, (P.errores + B.errores)[:5])
    ctx2.close(); b.close()
print(f"\nDemo: {BIEN[0]} bien, {MAL[0]} mal")
sys.exit(1 if MAL[0] else 0)

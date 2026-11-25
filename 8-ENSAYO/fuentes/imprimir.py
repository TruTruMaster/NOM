"""Imprime ensayo.html a PDF tamaño carta con Chromium (Playwright), en dos versiones:

    ENTREGA/5IM7_WACK_Ensayo.pdf                 4 páginas contando la portada (portada + 3)
    ENTREGA/5IM7_WACK_Ensayo_portada_aparte.pdf  portada + 4 páginas (si la portada no cuenta)

    ~/.cache/nom-pruebas/venv/bin/python imprimir.py

La portada sale sin pie de página; las demás llevan el número. Avisa si alguna versión se pasa de páginas.
"""
import os, re, subprocess, tempfile
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", os.path.expanduser("~/.cache/nom-pruebas/navegadores"))
from playwright.sync_api import sync_playwright

AQUI = os.path.dirname(os.path.abspath(__file__))
ENTRADA = os.path.join(AQUI, "ensayo.html")
ENTREGA = os.path.join(AQUI, "..", "..", "ENTREGA")
VERSIONES = [  # (clase del body, archivo, páginas máximas contando la portada)
    ("corto", "5IM7_WACK_Ensayo.pdf", 4),
    ("", "5IM7_WACK_Ensayo_portada_aparte.pdf", 5),
]
PIE = """
<div style="width:100%;margin:0 2.4cm;display:flex;justify-content:space-between;font-family:Carlito,Arial,sans-serif;font-size:8.5px;color:#5E5A66;">
  <span>WACK · NOM-COLAB · Ensayo: ¿es un sistema distribuido?</span><span class="pageNumber"></span>
</div>"""


def imprimir(pagina, ruta, pie):
    pagina.pdf(path=ruta, format="Letter", print_background=True, prefer_css_page_size=True,
               display_header_footer=pie, header_template="<span></span>", footer_template=PIE if pie else "<span></span>")


def paginas(ruta):
    return int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", ruta], capture_output=True, text=True).stdout).group(1))


with sync_playwright() as p, tempfile.TemporaryDirectory() as tmp:
    nav = p.chromium.launch()
    pag = nav.new_page()
    pag.goto("file://" + ENTRADA)
    pag.evaluate("document.fonts.ready")
    pag.wait_for_timeout(600)
    for clase, nombre, maximo in VERSIONES:
        pag.evaluate("c => document.body.className = c", clase)
        pag.wait_for_timeout(200)
        con_pie, sin_pie = os.path.join(tmp, "a.pdf"), os.path.join(tmp, "b.pdf")
        imprimir(pag, con_pie, True)
        imprimir(pag, sin_pie, False)
        total = paginas(con_pie)
        unido = os.path.join(tmp, "u.pdf")
        subprocess.run(["mutool", "merge", "-o", unido, sin_pie, "1", con_pie, f"2-{total}"], check=True)
        salida = os.path.abspath(os.path.join(ENTREGA, nombre))
        subprocess.run(["mutool", "clean", "-gggg", "-z", unido, salida], check=True)
        print(f"{nombre}: {total} páginas (máximo {maximo})" + ("" if total <= maximo else "  ¡SE PASA!"))
    nav.close()

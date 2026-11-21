#!/usr/bin/env python3
"""Genera app/src/main/resources/static/js/estilos-celda-datos.js: los estilos de celda de Excel (Inicio → Estilos de
celda) como los ve NOM, sacados de x13-estilos-celda.xlsx (hecho en Excel 2024 con el tema de Office por
pruebas/office/windows/hacer_estilos_celda.py): de cada estilo, las propiedades que el importador le pone a una celda con
texto y a una con número. Así aplicar un estilo en NOM da lo mismo que en Excel.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python herramientas/estilos_celda.py
"""
import base64, json, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
P = AQUI.parent / "pruebas"
sys.path.insert(0, str(P / "office")); sys.path.insert(0, str(P / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

ARCH = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real/archivos/x13-estilos-celda.xlsx"
SALIDA = AQUI.parent / "app/src/main/resources/static/js/estilos-celda-datos.js"
# los nombres de Excel en español (los de la galería de Excel 2024 en español)
ES = {"Normal": "Normal", "Good": "Bueno", "Bad": "Incorrecto", "Neutral": "Neutral", "Calculation": "Cálculo", "Check Cell": "Celda de comprobación",
      "Explanatory Text": "Texto explicativo", "Input": "Entrada", "Linked Cell": "Celda vinculada", "Note": "Notas", "Output": "Salida",
      "Warning Text": "Texto de advertencia", "Heading 1": "Encabezado 1", "Heading 2": "Encabezado 2", "Heading 3": "Encabezado 3",
      "Heading 4": "Encabezado 4", "Title": "Título", "Total": "Total", "Comma": "Millares", "Comma [0]": "Millares [0]", "Currency": "Moneda",
      "Currency [0]": "Moneda [0]", "Percent": "Porcentaje"}
for i in range(1, 7):
    ES[f"Accent{i}"] = f"Énfasis{i}"
    for p in (20, 40, 60): ES[f"{p}% - Accent{i}"] = f"{p}% - Énfasis{i}"
GRUPO = lambda n: "Bueno, malo y neutral" if n in ("Normal", "Good", "Bad", "Neutral") else "Datos y modelo" if n in ("Calculation", "Check Cell", "Explanatory Text", "Input", "Linked Cell", "Note", "Output", "Warning Text") \
    else "Títulos y encabezados" if n.startswith("Heading") or n in ("Title", "Total") else "Formato de número" if n.startswith(("Comma", "Currency", "Percent")) else "Estilos de celda temáticos"

JS = r"""() => { const t = NOM.editorActual.modoApi.api.tabla(), F = NOM.formulas, out = [];
  const de = (ref) => { const o = {}; for (const k in t) { const i = k.indexOf(':'); if (i < 0) continue; const p = k.slice(0, i), r = k.slice(i + 1); if (r === ref && p !== 'c' && p !== 'h') o[p] = t[k]; } return o; };
  for (let f = 0; f < 60; f++) { const n = t['c:A' + (f + 1)]; if (!n) break; out.push([String(n), de('B' + (f + 1)), de('C' + (f + 1))]); } return out; }"""


def main():
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    try:
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(ARCH.read_bytes()).decode(), ARCH.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
        filas = p.evaluate(JS)
    finally:
        nav.cerrar()
    datos = {n: {"es": ES.get(n, n), "g": GRUPO(n), "t": t, "n": num} for n, t, num in filas}
    SALIDA.write_text("/* Estilos de celda de Excel (Inicio → Estilos de celda) como los ve NOM: sacados de x13-estilos-celda.xlsx, hecho en\n"
                      "   Excel 2024 con el tema de Office (herramientas/estilos_celda.py). De cada uno: es (su nombre en español), g (su grupo),\n"
                      "   t (lo que pone en una celda con texto) y n (en una con número). No editar a mano. */\n"
                      "window.NOM = window.NOM || {};\nNOM.ESTILOS_CELDA = " + json.dumps(datos, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
    print(len(datos), "estilos →", SALIDA)
    for n, d in list(datos.items())[:6]: print(n, d["t"], d["n"])


if __name__ == "__main__":
    main()

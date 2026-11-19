#!/usr/bin/env python3
"""Herramientas de datos de la hoja (quitar duplicados, texto en columnas, buscar objetivo) contra Excel 2024 real.

  datos_excel.py [--solo-excel]

1. Arma los casos (CASOS) y los manda a la computadora con Windows: windows/datos_excel.py los hace en Excel y
   devuelve cada celda (valor y texto que se ve). La verdad queda en ~/.cache/nom-pruebas/office-real/datos-excel.json.
2. Hace lo mismo en NOM (la API de la hoja: NOM.editorActual.hojaApi.herramientas) y compara celda por celda.
"""
import json, os, sys, tempfile
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
import remoto

CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas"))
VERDAD = CACHE / "office-real/datos-excel.json"

NOMBRES = [["Nombre", "Grupo", "Cal"], ["Ana", "A", 9], ["Luis", "B", 7.5], ["ana", "A", 9], ["Eva", "A", 10], ["ANA", "B", 9],
           ["Luis", "B", 7.5], ["Eva ", "A", 10], ["Juan", "", 6], ["Juan", "", 6], ["1", "C", 1], [1, "C", 1]]
CASOS = [
    {"id": "dup-todas", "op": "quitar_duplicados", "datos": NOMBRES, "rango": "A1:C12", "columnas": [1, 2, 3], "encabezado": True},
    {"id": "dup-col1", "op": "quitar_duplicados", "datos": NOMBRES, "rango": "A1:C12", "columnas": [1], "encabezado": True},
    {"id": "dup-col12", "op": "quitar_duplicados", "datos": NOMBRES, "rango": "A1:C12", "columnas": [1, 2], "encabezado": True},
    {"id": "dup-sin-encabezado", "op": "quitar_duplicados", "datos": NOMBRES[1:], "rango": "A1:C11", "columnas": [1, 3], "encabezado": False},
    {"id": "txt-coma", "op": "texto_columnas", "rango": "A1:A6", "sep": {"coma": True},
     "datos": [["a,b,c"], ["1,2,3"], ["x,,y"], ["solo"], ["3.5,007,-5"], ["\"Pérez, Ana\",20,MX"]]},
    {"id": "txt-punto-coma", "op": "texto_columnas", "rango": "A1:A4", "sep": {"punto_coma": True},
     "datos": [["uno;dos;tres"], ["'a;b';c"], ["10%;$25;1e3"], ["TRUE;FALSE;verdadero"]]},
    {"id": "txt-espacio-seguidos", "op": "texto_columnas", "rango": "A1:A3", "sep": {"espacio": True}, "seguidos": True,
     "datos": [["Ana   Pérez  López"], ["  inicio con espacios"], ["5-  (7)  12-03-2026"]]},
    {"id": "txt-espacio", "op": "texto_columnas", "rango": "A1:A2", "sep": {"espacio": True}, "seguidos": False,
     "datos": [["a  b"], [" c d"]]},
    {"id": "txt-otro", "op": "texto_columnas", "rango": "A1:A3", "sep": {"otro": "|"},
     "datos": [["2026-11-18|18/11/2026|11/18/2026"], ["a|b"], ["|vacío|"]]},
    {"id": "txt-anchos", "op": "texto_columnas", "rango": "A1:A3", "anchos": [0, 4, 9],
     "datos": [["ABCD12345resto"], ["XY  9    final"], ["corto"]]},
    {"id": "obj-lineal", "op": "buscar_objetivo", "celda": "B2", "valor": 100, "cambiando": "B1", "datos": [["x", 1], ["2x+3", "=B1*2+3"]]},
    {"id": "obj-raiz", "op": "buscar_objetivo", "celda": "B2", "valor": 2, "cambiando": "B1", "datos": [["x", 1], ["x^2", "=B1^2"]]},
    {"id": "obj-pago", "op": "buscar_objetivo", "celda": "B2", "valor": 2000, "cambiando": "B1", "datos": [["tasa", 0.05], ["pago", "=PMT(B1/12,60,-100000)"]]},
    {"id": "obj-imposible", "op": "buscar_objetivo", "celda": "B2", "valor": -1, "cambiando": "B1", "datos": [["x", 3], ["x^2", "=B1^2"]]},
]


def verdad_excel():
    with tempfile.TemporaryDirectory() as d:
        ent = Path(d) / "datos-casos.json"
        ent.write_text(json.dumps(CASOS, ensure_ascii=False), encoding="utf-8")
        remoto.subir([ent], "nomoffice")
        r = remoto.correr_office("datos", "datos_excel.py", "datos-casos.json", "datos-verdad.json", timeout=900)
        print(r.stdout[-800:])
        if not remoto.bajar("nomoffice/datos-verdad.json", VERDAD): raise SystemExit("Excel no dejó el resultado")
    return json.load(open(VERDAD, encoding="utf-8"))


def main():
    v = verdad_excel()
    for c in CASOS:
        x = v.get(c["id"], {})
        print(f"== {c['id']}" + (f"  ERROR {x['error']}" if "error" in x else ""))
        for fila in x.get("celdas", []): print("   ", [f"{k['t']}" for k in fila])
        if "cambio" in x: print("    cambio:", x["cambio"], "resultado:", x.get("resultado"), "logrado:", x.get("logrado"))
    if "--solo-excel" in sys.argv: return 0


if __name__ == "__main__":
    sys.exit(main())

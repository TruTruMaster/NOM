"""Lo que hace Excel 2024 con las herramientas de datos, para compararlo con NOM (pruebas/office/datos_excel.py).

  python datos_excel.py ENTRADA.json SALIDA.json

ENTRADA: una lista de casos {"id", "datos": [[...filas...]], "op": ..., parámetros}. Para cada caso se escribe la tabla
en una hoja nueva (desde A1, valores tal cual: texto o número), se aplica la herramienta con la API de Excel y se
guarda el resultado: el valor y el texto que se ve de cada celda del área usada (y, en Buscar objetivo, el valor de la
celda que cambió).

Operaciones:
  quitar_duplicados   {"rango": "A1:C9", "columnas": [1, 2], "encabezado": true}
  texto_columnas      {"rango": "A1:A9", "sep": {"tab": false, "punto_coma": false, "coma": true, "espacio": false,
                       "otro": null}, "seguidos": false, "calificador": "\""}  ó  {"anchos": [0, 5, 12]}
  buscar_objetivo     {"celda": "B3", "valor": 100, "cambiando": "B1"}
  tabla_datos         {"rango": "A4:C8", "fila": "B1", "col": "B2"} (una de las dos o las dos)
"""
import json, sys
from comun import excel

ENT, SAL = sys.argv[1], sys.argv[2]
casos = json.load(open(ENT, encoding="utf-8"))
x = excel()
wb = x.Workbooks.Add()
res = {}
try:
    for c in casos:
        h = wb.Worksheets.Add()
        op = c["op"]
        try:
            for i, fila in enumerate(c["datos"]):
                for j, v in enumerate(fila):
                    if v is None or v == "": continue
                    cel = h.Cells(i + 1, j + 1)
                    if isinstance(v, str) and v.startswith("="): cel.Formula = v
                    elif isinstance(v, str): cel.NumberFormat = "@" if c.get("texto_literal") else "General"; cel.Value = v
                    else: cel.Value = v
            if op == "quitar_duplicados":
                h.Range(c["rango"]).RemoveDuplicates(Columns=tuple(c["columnas"]) if len(c["columnas"]) > 1 else c["columnas"][0], Header=1 if c.get("encabezado") else 2)
            elif op == "texto_columnas":
                r = h.Range(c["rango"])
                if c.get("anchos") is not None:
                    info = tuple((a, 1) for a in c["anchos"])
                    r.TextToColumns(Destination=r, DataType=2, FieldInfo=info)
                else:
                    s = c["sep"]
                    r.TextToColumns(Destination=r, DataType=1, TextQualifier={"\"": 1, "'": 2, "": -4142}.get(c.get("calificador", "\""), 1),
                                    ConsecutiveDelimiter=bool(c.get("seguidos")), Tab=bool(s.get("tab")), Semicolon=bool(s.get("punto_coma")),
                                    Comma=bool(s.get("coma")), Space=bool(s.get("espacio")), Other=bool(s.get("otro")), OtherChar=s.get("otro") or "")
            elif op == "consolidar":
                # fuentes «A1:C4» de esta hoja, en R1C1 como las pide Excel; función de consolidar de Excel
                def r1c1(a1):
                    import re
                    def una(x):
                        m = re.match(r"([A-Z]+)(\d+)", x); n = 0
                        for ch in m.group(1): n = n * 26 + ord(ch) - 64
                        return f"R{m.group(2)}C{n}"
                    a, _, b = a1.partition(":")
                    return una(a) + (":" + una(b) if b else "")
                FUN = {"suma": -4157, "contar": -4112, "promedio": -4106, "max": -4136, "min": -4139, "producto": -4149, "contar_num": -4113,
                       "desvest": -4155, "desvestp": -4156, "var": -4164, "varp": -4165}
                fuentes = tuple(f"'{h.Name}'!{r1c1(f)}" for f in c["fuentes"])
                h.Range(c["destino"]).Consolidate(Sources=fuentes, Function=FUN[c["funcion"]], TopRow=bool(c.get("fila")), LeftColumn=bool(c.get("columna")), CreateLinks=False)
            elif op == "tabla_datos":
                kw = {}
                if c.get("fila"): kw["RowInput"] = h.Range(c["fila"])
                if c.get("col"): kw["ColumnInput"] = h.Range(c["col"])
                h.Range(c["rango"]).Table(**kw)
            elif op == "buscar_objetivo":
                ok = h.Range(c["celda"]).GoalSeek(Goal=c["valor"], ChangingCell=h.Range(c["cambiando"]))
                res.setdefault(c["id"], {})["logrado"] = bool(ok)
            x.Calculate()
            u = h.UsedRange
            filas = []
            for i in range(1, u.Row + u.Rows.Count):
                fila = []
                for j in range(1, u.Column + u.Columns.Count):
                    cel = h.Cells(i, j)
                    fila.append({"v": cel.Value2 if not isinstance(cel.Value2, str) else cel.Value2, "t": cel.Text})
                filas.append(fila)
            res.setdefault(c["id"], {})["celdas"] = filas
            if op == "buscar_objetivo":
                res[c["id"]]["cambio"] = h.Range(c["cambiando"]).Value2
                res[c["id"]]["resultado"] = h.Range(c["celda"]).Value2
        except Exception as e:
            res.setdefault(c["id"], {})["error"] = str(e)[:300]
        print(c["id"], "error" if "error" in res[c["id"]] else "bien", flush=True)
finally:
    wb.Close(False)
    x.Quit()
json.dump(res, open(SAL, "w", encoding="utf-8"), ensure_ascii=False, indent=0)

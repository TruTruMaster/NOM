"""Cómo convierte Word 2024 texto en tabla (Range.ConvertToTable) y tabla en texto (Table.ConvertToText), para compararlo
con NOM (pruebas/office/tabla_texto_real.py). Uso: tabla_texto_word.py CASOS.json SALIDA.json
CASOS: [{"id", "parrafos": [...], "sep": 0 párrafos | 1 tabulaciones | 2 comas | "x" (otro carácter) | null (Word
adivina), "cols": n (solo con párrafos)} | {"id", "tabla": [[...], ...], "sep": …, "a_texto": true}].
SALIDA: {id: {"celdas": [[...]]} | {"texto": [párrafos]}}."""
import json, sys
from comun import word

casos = json.load(open(sys.argv[1], encoding="utf-8"))
w = word()
res = {}
try:
    for c in casos:
        d = w.Documents.Add()
        try:
            if c.get("a_texto"):
                filas = c["tabla"]
                t = d.Tables.Add(d.Range(0, 0), len(filas), max(len(f) for f in filas))
                for i, f in enumerate(filas):
                    for j, v in enumerate(f): t.Cell(i + 1, j + 1).Range.Text = v
                kw = {} if c.get("sep") is None else {"Separator": c["sep"]}
                t.ConvertToText(**kw)
                res[c["id"]] = {"texto": d.Content.Text.split("\r")[:-1]}
            else:
                d.Content.Text = "\r".join(c["parrafos"])
                r = d.Range(0, d.Content.End - 1)
                kw = {}
                if c.get("sep") is not None: kw["Separator"] = c["sep"]
                if c.get("cols"): kw["NumColumns"] = c["cols"]
                t = r.ConvertToTable(**kw)
                celdas = [[t.Cell(i, j).Range.Text[:-2] if True else "" for j in range(1, t.Columns.Count + 1)] for i in range(1, t.Rows.Count + 1)]
                res[c["id"]] = {"celdas": celdas, "filas": t.Rows.Count, "cols": t.Columns.Count}
        except Exception as e:
            res[c["id"]] = {"error": str(e)[:300]}
        finally:
            d.Close(0)
finally:
    w.Quit()
json.dump(res, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", len(res), flush=True)

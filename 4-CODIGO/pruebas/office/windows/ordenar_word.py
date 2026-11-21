"""Cómo ordena Word 2024 los párrafos (Inicio → Párrafo → Ordenar, Range.Sort), para compararlo con NOM
(pruebas/office/ordenar_real.py). Uso: ordenar_word.py CASOS.json SALIDA.json
CASOS: [{"id", "parrafos": [...], "tipo": 0 texto | 1 número | 2 fecha, "orden": 0 ascendente | 1 descendente,
"encabezado": bool, "idioma": 2058 (español de México) | 1033…, "mayusculas": bool}]. SALIDA: {id: [párrafos ordenados]}."""
import json, os, sys
from comun import word

casos = json.load(open(sys.argv[1], encoding="utf-8"))
w = word()
res = {}
try:
    for c in casos:
        d = w.Documents.Add()
        try:
            d.Content.Text = "\r".join(c["parrafos"])
            d.Content.LanguageID = c.get("idioma", 2058)
            d.Content.Sort(ExcludeHeader=bool(c.get("encabezado")), FieldNumber="Paragraphs", SortFieldType=c.get("tipo", 0),
                           SortOrder=c.get("orden", 0), CaseSensitive=bool(c.get("mayusculas")), LanguageID=c.get("idioma", 2058))
            res[c["id"]] = d.Content.Text.split("\r")[:len(c["parrafos"])]
        except Exception as e:
            res[c["id"]] = {"error": str(e)[:300]}
        finally:
            d.Close(0)
finally:
    w.Quit()
json.dump(res, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", len(res), flush=True)

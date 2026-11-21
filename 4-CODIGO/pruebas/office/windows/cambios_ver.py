"""Lo que ve Word de los cambios marcados de un .docx: si tiene «Control de cambios», la lista de revisiones (tipo —1
agregar, 2 borrar, 3 formato…—, autor, texto) y el texto de cada párrafo con todos los cambios aceptados y con todos
rechazados (sin guardar nada). Uso: cambios_ver.py ENTRADA.docx SALIDA.json"""
import json, os, sys
from comun import word

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
w = word()
datos = {"abrio": False}


def parrafos(d):
    # cada párrafo (las celdas de una tabla terminan en \r\x07)
    return [p for p in d.Content.Text.replace("\x07", "").split("\r")]


try:
    d = w.Documents.Open(entrada, False, True)
    try:
        datos["abrio"] = True
        datos["control"] = bool(d.TrackRevisions)
        revs = []
        for i in range(1, d.Revisions.Count + 1):
            r = d.Revisions(i)
            try: txt = r.Range.Text
            except Exception: txt = None
            try: en_tabla = bool(r.Range.Information(12))   # wdWithInTable
            except Exception: en_tabla = None
            revs.append({"tipo": r.Type, "autor": r.Author, "texto": txt, "tabla": en_tabla})
        datos["revisiones"] = revs
        datos["tal_cual"] = parrafos(d)
        d.Revisions.AcceptAll()
        datos["aceptado"] = parrafos(d)
    finally:
        d.Close(0)
    d = w.Documents.Open(entrada, False, True)
    try:
        d.Revisions.RejectAll()
        datos["rechazado"] = parrafos(d)
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), len(datos.get("revisiones", [])), flush=True)

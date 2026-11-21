"""Las tabulaciones de cada párrafo de un .docx como las lee Word (Paragraph.TabStops: posición en pt, alineación —0
izquierda, 1 centro, 2 derecha, 3 decimal, 4 barra— y relleno —0 ninguno, 1 puntos, 2 guiones, 3 subrayado, 5 punto
medio—) y las predeterminadas (DefaultTabStop). Uso: tabs_ver.py ENTRADA.docx SALIDA.json"""
import json, os, sys
from comun import word

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
w = word()
datos = {"abrio": False}
try:
    d = w.Documents.Open(entrada, False, True)
    try:
        datos["abrio"] = True
        datos["predeterminada"] = d.DefaultTabStop
        ps = []
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            ts = p.TabStops
            ps.append({"texto": p.Range.Text.rstrip("\r"), "tabs": [[round(ts(k).Position, 2), ts(k).Alignment, ts(k).Leader] for k in range(1, ts.Count + 1) if not ts(k).CustomTab is False]})
        datos["parrafos"] = ps
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

"""Los estilos de un .docx como los ve Word: de cada párrafo, su estilo (nombre) y la letra con que se ve; de los estilos
pedidos, su definición (fuente, tamaño, negrita, cursiva, color —RGB—, alineación —0 izq, 1 centro, 2 der, 3 justif.—,
espacio antes y después en pt). Uso: estilos_ver.py ENTRADA.docx SALIDA.json [ESTILO …]"""
import json, os, sys
from comun import word

entrada, salida = (os.path.abspath(a) for a in sys.argv[1:3])
pedidos = sys.argv[3:]
w = word()
datos = {"abrio": False}
rgb = lambda c: "#%02X%02X%02X" % (c & 255, (c >> 8) & 255, (c >> 16) & 255) if isinstance(c, int) and 0 <= c < 2 ** 24 else None
try:
    d = w.Documents.Open(entrada, False, True)
    try:
        datos["abrio"] = True
        ps = []
        for i in range(1, d.Paragraphs.Count + 1):
            p = d.Paragraphs(i)
            f = p.Range.Font
            ps.append({"texto": p.Range.Text.rstrip("\r"), "estilo": p.Style.NameLocal, "fuente": f.Name, "tam": f.Size, "b": f.Bold, "color": rgb(f.Color), "alin": p.Alignment,
                       "antes": p.SpaceBefore, "despues": p.SpaceAfter})
        datos["parrafos"] = ps
        est = {}
        for n in pedidos:
            try:
                s = d.Styles(n)
                est[n] = {"fuente": s.Font.Name, "tam": s.Font.Size, "b": s.Font.Bold, "i": s.Font.Italic, "color": rgb(s.Font.Color), "alin": s.ParagraphFormat.Alignment,
                          "antes": s.ParagraphFormat.SpaceBefore, "despues": s.ParagraphFormat.SpaceAfter, "propio": not s.BuiltIn}
            except Exception as e:
                est[n] = {"error": str(e)[:150]}
        datos["estilos"] = est
    finally:
        d.Close(0)
except Exception as e:
    datos["error"] = str(e)[:300]
finally:
    w.Quit()
json.dump(datos, open(salida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("hecho", datos.get("abrio"), flush=True)

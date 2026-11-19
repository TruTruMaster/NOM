"""w14-renglones.docx: cientos de renglones controlados para aprender dónde corta Word un renglón.

Párrafos de palabras al azar (con acentos, signos y números) en 8 fuentes de Office, 3 tamaños, a la izquierda y
justificados, y 3 anchos distintos (sangría derecha). Sin interletraje ni ligaduras, sin guiones, interlineado
sencillo. Deja junto al documento w14-renglones.json con lo que lleva cada párrafo (fuente, tamaño, alineación,
sangría y texto) para comparar renglón por renglón con verdad.py (analisis/acomodo_word.py).
"""
import json, os, random
from comun import word, ARCHIVOS

PALABRAS = ("la el de que y en un una los las del se por con para como más pero sus le ya o este sí porque esta "
            "entre cuando muy sin sobre también me hasta hay donde quien desde todo nos durante todos uno les ni "
            "contra otros ese eso ante ellos e esto mí antes algunos qué unos yo otro otras otra él tanto esa estos "
            "mucho quienes nada muchos cual poco ella estar estas algunas algo nosotros colaboración documento "
            "información administración universidad estudiantes herramientas extraordinariamente responsabilidad "
            "conexión sección imágenes referencias exactamente programa resultados encuesta aplicada mayoría prefiere "
            "trabajar equipo sencillas transparentes confiables análisis electrónico teléfono matemáticas "
            "investigación comunicación organización año niño señal pequeño mañana después público práctica "
            "the of and to in is was for on are with as his they be at one have this from or had by word but what "
            "some we can out other were all there when up use your how said an each she which do their time if "
            "will way about many then them write would like so these her long make thing see him two has look "
            "2026 15 % 3.5 1,250 (nota) «cita» —guion— ¿qué? ¡sí! etc. p. ej. A-12 www.ejemplo.mx").split()
FUENTES = ["Calibri", "Aptos", "Times New Roman", "Arial", "Cambria", "Georgia", "Verdana", "Segoe UI"]
TAMANOS = [10, 11, 12.5]
SANGRIAS = [0, 61.3, 147.9]          # puntos: tres anchos de renglón (468, 406.7, 320.1)


def texto(rnd, n):
    ws = [rnd.choice(PALABRAS) for _ in range(n)]
    ws[0] = ws[0][:1].upper() + ws[0][1:]
    s = " ".join(ws)
    # algunas comas y puntos en medio, como en un texto de verdad
    out = []
    for i, w in enumerate(s.split(" ")):
        out.append(w + ("," if rnd.random() < 0.06 else "." if rnd.random() < 0.04 else ""))
    return " ".join(out) + "."


rnd = random.Random(14)
w = word()
d = w.Documents.Add()
meta = []
try:
    d.PageSetup.PageWidth, d.PageSetup.PageHeight = 612, 792
    for lado in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"): setattr(d.PageSetup, lado, 72)
    d.AutoHyphenation = False
    estilo = d.Styles(-1)                       # Normal
    estilo.Font.Kerning = 0
    estilo.ParagraphFormat.SpaceAfter = 0
    estilo.ParagraphFormat.SpaceBefore = 0
    estilo.ParagraphFormat.LineSpacingRule = 0  # sencillo
    primero = True
    for fuente in FUENTES:
        for tam in TAMANOS:
            for alin in (0, 3):
                for sangria in SANGRIAS:
                    t = texto(rnd, rnd.randint(70, 110))
                    if primero: r = d.Paragraphs(1).Range; primero = False
                    else:
                        d.Content.InsertParagraphAfter()
                        r = d.Paragraphs(d.Paragraphs.Count).Range
                    r.Text = t
                    p = d.Paragraphs(d.Paragraphs.Count)
                    p.Range.Font.Name = fuente
                    p.Range.Font.Size = tam
                    p.Range.Font.Kerning = 0
                    p.Alignment = alin
                    p.LeftIndent = 0
                    p.FirstLineIndent = 0
                    p.RightIndent = sangria
                    p.SpaceAfter = 0
                    p.SpaceBefore = 0
                    meta.append({"fuente": fuente, "tam": tam, "alin": "just" if alin == 3 else "izq", "sangria": sangria, "texto": t})
    salida = os.path.join(ARCHIVOS, "w14-renglones.docx")
    d.SaveAs2(salida, 16)
    json.dump({"ancho": 468, "parrafos": meta}, open(os.path.join(ARCHIVOS, "w14-renglones.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print("hecho", len(meta), "párrafos,", d.ComputeStatistics(2), "páginas", flush=True)
finally:
    d.Close(0)
    w.Quit()

"""w15-guiones-*.docx: cientos de renglones con «Guiones automáticos» para aprender CUÁNDO parte Word una palabra.

Párrafos de palabras al azar (español con es-MX, inglés con en-US; muchas palabras largas) en 4 fuentes, 3 tamaños,
a la izquierda y justificados, 3 anchos (sangría derecha). Un documento por zona de división (0.1", 0.25" —la de
siempre— y 0.5") y uno con «Limitar guiones consecutivos a» 2. Sin interletraje ni ligaduras, interlineado sencillo.
Deja junto a cada documento su .json con lo que lleva cada párrafo, para comparar renglón por renglón con verdad.py
(analisis/guiones_decision.py). Uso: hacer_guiones2.py [vocabulario.json]"""
import json, os, random, sys
from comun import word, ARCHIVOS

VOCAB = json.load(open(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "guiones_vocab.json"), encoding="utf-8"))
FUENTES = ["Calibri", "Times New Roman", "Arial", "Cambria"]
TAMANOS = [10, 11, 12]
SANGRIAS = [0, 95.5, 201.3]          # puntos: tres anchos de renglón (468, 372.5, 266.7)
DOCUMENTOS = [("w15-guiones-z18", 18, 0), ("w15-guiones-z7", 7.2, 0), ("w15-guiones-z36", 36, 0), ("w15-guiones-lim2", 18, 2)]


def texto(rnd, idioma, n):
    cortas, largas, medias = VOCAB[idioma]
    ws = []
    for _ in range(n):
        x = rnd.random()
        ws.append(rnd.choice(cortas) if x < 0.35 else rnd.choice(medias) if x < 0.55 else rnd.choice(largas))
    ws[0] = ws[0][:1].upper() + ws[0][1:]
    out = []
    for w in ws: out.append(w + ("," if rnd.random() < 0.05 else ""))
    return " ".join(out) + "."


for nombre, zona, limite in DOCUMENTOS:
    rnd = random.Random(nombre)
    w = word()
    d = w.Documents.Add()
    meta = []
    try:
        d.PageSetup.PageWidth, d.PageSetup.PageHeight = 612, 792
        for lado in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"): setattr(d.PageSetup, lado, 72)
        d.AutoHyphenation = True
        d.HyphenationZone = zona
        d.HyphenateCaps = True
        d.ConsecutiveHyphensLimit = limite
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
                        for idioma, lang in (("es", 2058), ("en", 1033)):
                            t = texto(rnd, idioma, rnd.randint(55, 80))
                            if primero: r = d.Paragraphs(1).Range; primero = False
                            else:
                                d.Content.InsertParagraphAfter()
                                r = d.Paragraphs(d.Paragraphs.Count).Range
                            r.Text = t
                            p = d.Paragraphs(d.Paragraphs.Count)
                            p.Range.Font.Name = fuente
                            p.Range.Font.Size = tam
                            p.Range.Font.Kerning = 0
                            p.Range.LanguageID = lang
                            p.Alignment = alin
                            p.LeftIndent = 0
                            p.FirstLineIndent = 0
                            p.RightIndent = sangria
                            p.SpaceAfter = 0
                            p.SpaceBefore = 0
                            meta.append({"fuente": fuente, "tam": tam, "alin": "just" if alin == 3 else "izq", "sangria": sangria,
                                         "idioma": "es-MX" if lang == 2058 else "en-US", "texto": t})
        salida = os.path.join(ARCHIVOS, nombre + ".docx")
        d.SaveAs2(salida, 16)
        json.dump({"ancho": 468, "zona": zona, "limite": limite, "parrafos": meta},
                  open(os.path.join(ARCHIVOS, nombre + ".json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
        print("hecho", nombre, len(meta), "párrafos,", d.ComputeStatistics(2), "páginas", flush=True)
    finally:
        d.Close(0)
        w.Quit()

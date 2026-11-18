"""Un documento de Word con campos de todo tipo, actualizados por Word (Fields.Update): la verdad para comparar con
«Actualizar campos» de NOM (pruebas/office/campos_real.py).

  correr.ps1 campos campos_word.py        → archivos\\campos-word.docx y archivos\\campos-word.json ({código, resultado})
"""
import json, os
from comun import word, ARCHIVOS

CAMPOS = [
    # fechas (el idioma del texto es español de México)
    'DATE \\@ "dddd, d \'de\' MMMM \'de\' yyyy"', 'DATE \\@ "d/MM/yy"', 'DATE \\@ "MMM-yyyy"', 'DATE \\@ "dd.MM.yyyy"', 'DATE',
    'CREATEDATE \\@ "d/M/yyyy"',
    # numeración de títulos
    "SEQ Figura \\* ARABIC", "SEQ Figura \\* ARABIC", "SEQ Tabla \\* ROMAN", "SEQ Figura \\c", "SEQ Tabla \\* ROMAN", "SEQ Figura \\r 10",
    "SEQ Figura", "SEQ Lista \\* ALPHABETIC", "SEQ Lista \\* alphabetic", "SEQ Lista \\* roman", "SEQ Oculto \\h",
    # referencias
    "REF marcaUno \\h", "REF marcaUno \\* Upper", "REF marcaUno \\* Lower", "REF marcaUno \\* FirstCap",
    # condiciones y fórmulas
    'IF 5 > 3 "mayor" "menor"', 'IF 2 = 3 "igual" "distinto"', 'IF "abc" = "a*" "coincide" "no coincide"', 'IF 10 <= 10 "sí" "no"',
    "= 2+3*4", '= 10/4 \\# "0.00"', '= 1234567.891 \\# "#,##0.00"', "= SUM(1,2,3)", "= AVERAGE(2,4,9)", "= MAX(3,8,1) - MIN(3,8,1)",
    "= ROUND(3.14159, 2)", "= ABS(-7)", "= 7 \\* CardText", "= 21 \\* Ordinal", "= 4 \\* ROMAN", "= 1500 \\* DollarText", "= 3 \\* OrdText",
    '= 0.256 \\# "0.0%"', "= 2^10", "= MOD(17,5)", "= 10 > 3",
    # propiedades del documento
    "AUTHOR", "TITLE", "SUBJECT", "AUTHOR \\* Upper", "TITLE \\* Caps",
    # páginas (este párrafo está en la página 1)
    "PAGE", "NUMPAGES", "PAGE \\* roman",
]

w = word()
d = w.Documents.Add()
try:
    d.Content.LanguageID = 2058   # español (México)
    d.BuiltInDocumentProperties("Title").Value = "Prueba de campos de nom"
    d.BuiltInDocumentProperties("Author").Value = "Equipo NOM"
    d.BuiltInDocumentProperties("Subject").Value = "Campos que se actualizan"
    r = d.Content
    r.InsertAfter("Campos de Word\r")
    d.Paragraphs(1).Style = -2   # Título 1
    # un marcador con texto
    r = d.Content; r.InsertAfter("El texto del marcador es: ")
    fin = d.Content.End - 1
    r = d.Range(fin, fin); r.InsertAfter("bandera de prueba")
    d.Bookmarks.Add("marcaUno", d.Range(fin, fin + len("bandera de prueba")))
    d.Content.InsertAfter("\r")
    for i, c in enumerate(CAMPOS):
        p = d.Content
        p.InsertAfter(f"C{i:02d}: ")
        e = d.Content.End - 1
        d.Fields.Add(d.Range(e, e), -1, c, False)
        d.Content.InsertAfter("\r")
    # página 2: un marcador y referencias a su página
    d.Content.InsertAfter("\r")
    e = d.Content.End - 1
    d.Range(e, e).InsertBreak(7)   # salto de página
    d.Content.InsertAfter("En la página dos: ")
    e = d.Content.End - 1
    d.Range(e, e).InsertAfter("marca dos")
    d.Bookmarks.Add("marcaDos", d.Range(e, e + len("marca dos")))
    d.Content.InsertAfter("\r")
    extra = ["PAGEREF marcaDos \\h", "PAGE", "PAGEREF marcaUno", "SEQ Figura", "REF marcaDos"]
    for j, c in enumerate(extra):
        d.Content.InsertAfter(f"P{j:02d}: ")
        e = d.Content.End - 1
        d.Fields.Add(d.Range(e, e), -1, c, False)
        d.Content.InsertAfter("\r")
    d.Fields.Update()
    d.Repaginate()
    d.Fields.Update()
    salida = os.path.join(ARCHIVOS, "campos-word.docx")
    d.SaveAs2(salida, 16)
    res = [{"codigo": f.Code.Text.strip(), "resultado": f.Result.Text} for f in d.Fields]
    json.dump(res, open(os.path.join(ARCHIVOS, "campos-word.json"), "w", encoding="utf8"), ensure_ascii=False, indent=1)
    print(f"campos: {len(res)}", flush=True)
finally:
    d.Close(0)
    w.Quit()

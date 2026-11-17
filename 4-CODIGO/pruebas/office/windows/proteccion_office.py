"""Archivos protegidos con Office de verdad (no cifrados: protección de edición).

  python proteccion_office.py CONTRASEÑA

Deja en archivos\\:
  prot-word-lectura.docx      Restringir edición: solo lectura (con contraseña)
  prot-word-comentarios.docx  solo comentarios
  prot-word-cambios.docx      solo cambios con control de cambios
  prot-word-formularios.docx  solo rellenar formularios
  prot-word-final.docx        marcado como final
  prot-word-modificar.docx    contraseña para modificar y «se recomienda abrir como de solo lectura»
  prot-excel-hoja.xlsx        hoja protegida (permite dar formato y ordenar), un rango sin bloquear, fórmulas ocultas,
                              y la estructura del libro protegida
  prot-excel-modificar.xlsx   contraseña para modificar y solo lectura recomendado
  prot-ppt-modificar.pptx     contraseña para modificar
  prot-ppt-final.pptx         marcada como final
"""
import os, sys
from comun import *


def main(clave):
    w = word()
    try:
        for nombre, tipo in [("lectura", 3), ("comentarios", 1), ("cambios", 0), ("formularios", 2)]:
            if os.path.exists(os.path.join(ARCHIVOS, f"prot-word-{nombre}.docx")): continue
            d = w.Documents.Open(os.path.join(ARCHIVOS, "w03-estilos-listas.docx"), False, True)
            d.Protect(tipo, True, clave)
            d.SaveAs2(os.path.join(ARCHIVOS, f"prot-word-{nombre}.docx"), 16)
            d.Close(0)
        # (Final = True abre un aviso en Word y lo deja esperando: se ponen las mismas propiedades que deja «Marcar como final»)
        d = w.Documents.Open(os.path.join(ARCHIVOS, "w03-estilos-listas.docx"), False, True)
        d.CustomDocumentProperties.Add("_MarkAsFinal", False, 2, True)
        try: d.BuiltInDocumentProperties("Content status").Value = "Final"
        except Exception: pass
        d.SaveAs2(os.path.join(ARCHIVOS, "prot-word-final.docx"), 16)
        d.Close(0)
        d = w.Documents.Open(os.path.join(ARCHIVOS, "w03-estilos-listas.docx"), False, True)
        d.SaveAs2(os.path.join(ARCHIVOS, "prot-word-modificar.docx"), 16, False, "", True, clave, True)
        d.Close(0)
    finally:
        w.Quit(0)
    x = excel()
    try:
        wb = x.Workbooks.Open(os.path.join(ARCHIVOS, "x08-registro.xlsx"))
        ws = wb.Worksheets(1)
        ws.Range("B3:D6").Locked = False
        ws.Range("E3:E6").FormulaHidden = True
        # Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns,
        #         AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows, AllowInsertingHyperlinks, AllowDeletingColumns,
        #         AllowDeletingRows, AllowSorting, AllowFiltering, AllowUsingPivotTables)
        ws.Protect(clave, True, True, True, False, True, False, False, False, False, False, False, False, True, False, False)
        wb.Protect(clave, True, False)
        wb.SaveAs(os.path.join(ARCHIVOS, "prot-excel-hoja.xlsx"), 51)
        wb.Close(False)
        wb = x.Workbooks.Open(os.path.join(ARCHIVOS, "x08-registro.xlsx"))
        wb.SaveAs(os.path.join(ARCHIVOS, "prot-excel-modificar.xlsx"), 51, None, clave, True)
        wb.Close(False)
    finally:
        x.Quit()
    p = powerpoint()
    pr = p.Presentations.Open(os.path.join(ARCHIVOS, "p09-presentacion.pptx"), True, False, False)
    try:
        pr.WritePassword = clave
        pr.SaveAs(os.path.join(ARCHIVOS, "prot-ppt-modificar.pptx"))
    finally:
        pr.Close()
    pr = p.Presentations.Open(os.path.join(ARCHIVOS, "p09-presentacion.pptx"), True, False, False)
    try:
        pr.CustomDocumentProperties.Add("_MarkAsFinal", False, 2, True)
        try: pr.BuiltInDocumentProperties("Content status").Value = "Final"
        except Exception: pass
        pr.SaveAs(os.path.join(ARCHIVOS, "prot-ppt-final.pptx"))
    finally:
        pr.Close()
    print("listos")


if __name__ == "__main__":
    import pythoncom
    pythoncom.CoInitialize()
    main(sys.argv[1])

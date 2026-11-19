"""Las funciones de hoja que tiene Excel de verdad (la biblioteca de tipos de WorksheetFunction), con su nombre en
inglés. Deja archivos\\funciones-excel.json. Lo usa pruebas/office/funciones_faltan.py."""
import json, os
import pythoncom
from comun import excel, ARCHIVOS

x = excel()
try:
    wf = x.WorksheetFunction
    ti = wf._oleobj_.GetTypeInfo()
    attr = ti.GetTypeAttr()
    nombres = set()
    for i in range(attr.cFuncs):
        fd = ti.GetFuncDesc(i)
        n = ti.GetNames(fd.memid)[0]
        if not n.startswith("_") and n not in ("QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetTypeInfo", "GetIDsOfNames", "Invoke", "Application", "Creator", "Parent"):
            nombres.add(n.replace("_", "."))
    json.dump(sorted(nombres), open(os.path.join(ARCHIVOS, "funciones-excel.json"), "w"), indent=0)
    print("funciones:", len(nombres), flush=True)
finally:
    x.Quit()

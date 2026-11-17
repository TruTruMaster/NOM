"""Comprueba en Office de verdad la protección de archivos que salieron de NOM.

  python proteccion_ver.py CONTRASEÑA RUTAS…

Word: tipo de protección (ProtectionType: -1 ninguna, 0 cambios, 1 comentarios, 2 formularios, 3 solo lectura), si está marcado
como final y si pide contraseña para modificar; quitar la protección con una contraseña mala (debe fallar) y con la buena.
Excel: hoja protegida, celdas bloqueadas, estructura del libro protegida, contraseña para modificar; desproteger con mala y buena.
PowerPoint: marcada como final y si pide contraseña para modificar (abrir sin ella queda de solo lectura).
Deja archivos\\proteccion_ver.json.
"""
import glob, json, os, sys
from comun import *


def ver_word(w, ruta, clave):
    r = {}
    # (con contraseña para modificar, Word la pide al abrir: se abre de solo lectura sin ella)
    d = w.Documents.Open(ruta, False, True, False, "", "", False, "")
    try:
        r["proteccion"] = d.ProtectionType
        try: r["final"] = bool(d.Final)
        except Exception: pass
        r["pide_modificar"] = bool(d.WriteReserved)
        r["solo_lectura_recomendada"] = bool(d.ReadOnlyRecommended)
        if d.ProtectionType != -1:
            try: d.Unprotect("contraseña-mala"); r["mala_quito"] = d.ProtectionType == -1
            except Exception: r["mala_quito"] = False
            try: d.Unprotect(clave); r["buena_quito"] = d.ProtectionType == -1
            except Exception as e: r["buena_quito"] = False; r["error_buena"] = str(e)[:120]
    finally:
        d.Close(0)
    if r.get("pide_modificar"):
        try:
            d = w.Documents.Open(ruta, False, False, False, "", "", False, clave)
            r["abre_con_clave_modificar"] = not d.ReadOnly
            d.Close(0)
        except Exception as e: r["abre_con_clave_modificar"] = False; r["error_modificar"] = str(e)[:120]
    return r


def paso(t): print("  ·", t, flush=True)


def ver_excel(x, ruta, clave):
    r = {}
    paso("abrir " + os.path.basename(ruta))
    wb = x.Workbooks.Open(ruta, 0, True)
    paso("abierto")
    try:
        ws = wb.Worksheets(1)
        r["hoja_protegida"] = bool(ws.ProtectContents)
        r["A1_bloqueada"] = bool(ws.Range("A1").Locked)
        r["E1_bloqueada"] = bool(ws.Range("E1").Locked)
        r["deja_formato"] = bool(ws.Protection.AllowFormattingCells)
        r["estructura_protegida"] = bool(wb.ProtectStructure)
        r["pide_modificar"] = bool(wb.WriteReserved)
        paso("leído: " + str(r))
        if ws.ProtectContents:
            try: ws.Unprotect("contraseña-mala"); r["mala_quito"] = not ws.ProtectContents
            except Exception: r["mala_quito"] = False
            try: ws.Unprotect(clave); r["buena_quito"] = not ws.ProtectContents
            except Exception as e: r["buena_quito"] = False; r["error_buena"] = str(e)[:120]
        if wb.ProtectStructure:
            try: wb.Unprotect("contraseña-mala"); r["libro_mala_quito"] = not wb.ProtectStructure
            except Exception: r["libro_mala_quito"] = False
            try: wb.Unprotect(clave); r["libro_buena_quito"] = not wb.ProtectStructure
            except Exception as e: r["libro_buena_quito"] = False
    finally:
        wb.Close(False)
    if r.get("pide_modificar"):
        try:
            wb = x.Workbooks.Open(ruta, 0, False, None, "", clave)
            r["abre_con_clave_modificar"] = not wb.ReadOnly
            wb.Close(False)
        except Exception as e: r["abre_con_clave_modificar"] = False; r["error_modificar"] = str(e)[:120]
    return r


def ver_ppt(p, ruta, clave):
    r = {}
    pr = p.Presentations.Open(ruta, True, False, False)
    try:
        try: r["final"] = bool(pr.Final)
        except Exception: pass
        r["tiene_clave_modificar"] = bool(pr.HasPasswordProtection) if hasattr(pr, "HasPasswordProtection") else None
    finally:
        pr.Close()
    try:
        pr = p.Presentations.Open(ruta + "::::" + clave, False, False, False)
        r["abre_con_clave_modificar"] = not pr.ReadOnly
        pr.Close()
    except Exception as e: r["abre_con_clave_modificar"] = False; r["error_modificar"] = str(e)[:160]
    return r


if __name__ == "__main__":
    import pythoncom
    pythoncom.CoInitialize()
    clave = sys.argv[1]
    rutas = []
    for a in sys.argv[2:]: rutas += glob.glob(a) or glob.glob(os.path.join(BASE, a))
    res = {}
    w = x = None
    try:
        for ruta in sorted(rutas):
            ruta = os.path.abspath(ruta); ext = os.path.splitext(ruta)[1].lower(); n = os.path.basename(ruta)
            try:
                if ext == ".docx":
                    w = w or word(); res[n] = ver_word(w, ruta, clave)
                elif ext == ".xlsx":
                    x = x or excel(); res[n] = ver_excel(x, ruta, clave)
                elif ext == ".pptx":
                    res[n] = ver_ppt(powerpoint(), ruta, clave)
            except Exception as e:
                res[n] = {"error": str(getattr(e, "excepinfo", None) or e)[:300]}
            print(n, res[n], flush=True)
    finally:
        if w: w.Quit(0)
        if x: x.Quit()
    guardar_json(os.path.join(ARCHIVOS, "proteccion_ver.json"), res)

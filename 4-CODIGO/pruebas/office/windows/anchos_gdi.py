"""Los anchos ENTEROS con que Windows dibuja cada carácter de un número (GetCharWidth32) en las fuentes de Excel, a
los tamaños de 8 a 20 puntos (96 ppp): con ellos decide Excel si un número cabe en su columna o sale «####».
Solo medidas (no el dibujo de la fuente). Deja anchos_gdi.json en archivos/: {fuente: {puntos: {carácter: px}}}.
"""
import ctypes, os
from comun import guardar_json, ARCHIVOS

gdi32, user32 = ctypes.windll.gdi32, ctypes.windll.user32
gdi32.CreateFontW.restype = ctypes.c_void_p
gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]; gdi32.SelectObject.restype = ctypes.c_void_p
gdi32.DeleteObject.argtypes = [ctypes.c_void_p]
gdi32.GetCharWidth32W.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint, ctypes.POINTER(ctypes.c_int)]
user32.GetDC.restype = ctypes.c_void_p
hdc = user32.GetDC(None)
CARACTERES = "0123456789.,-+$€£¥%()/:E ' "
FUENTES = ["Aptos Narrow", "Aptos", "Calibri", "Arial", "Cambria", "Times New Roman", "Segoe UI", "Verdana", "Tahoma", "Consolas", "Courier New"]
res = {}
for fam in FUENTES:
    res[fam] = {}
    for pt in range(8, 21):
        alto = -round(pt * 96 / 72)   # como MulDiv(pt, 96, 72): alto del carácter en píxeles
        for negrita in (400, 700):
            hf = gdi32.CreateFontW(alto, 0, 0, 0, negrita, 0, 0, 0, 1, 0, 0, 5, 0, fam)
            viejo = gdi32.SelectObject(hdc, hf)
            m = {}
            for ch in CARACTERES:
                w = ctypes.c_int()
                if gdi32.GetCharWidth32W(hdc, ord(ch), ord(ch), ctypes.byref(w)): m[ch] = w.value
            gdi32.SelectObject(hdc, viejo); gdi32.DeleteObject(hf)
            res[fam][f"{pt}{'b' if negrita == 700 else ''}"] = m
guardar_json(os.path.join(ARCHIVOS, "anchos_gdi.json"), res)
print("hecho", flush=True)

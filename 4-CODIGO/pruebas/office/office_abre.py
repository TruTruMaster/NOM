"""¿Lo abre Office de verdad sin avisos ni «reparar»? Para las pruebas que descargan archivos de NOM.

  import office_abre
  res = office_abre.abrir(["/ruta/a.docx", "/ruta/b.xlsx"], "ediciones")   # {nombre: {"abrio":…, "avisos":[…], …}}
  office_abre.bien(res["a.docx"])                                        # True si abrió sin avisos ni reparar

Manda todos los archivos de una vez (un solo paquete por SSH) a la computadora con Windows y corre
windows/abrir_office.py en la sesión de escritorio: Word, Excel o PowerPoint abren cada uno CON sus avisos y se anota
si pidieron reparar. Sin la computadora con Windows (o sin red), devuelve {} y quien llama decide.
"""
import json, sys, tempfile
sys.path.insert(0, str(__import__("pathlib").Path(__file__).resolve().parent))
from pathlib import Path

import remoto


def disponible():
    return remoto.disponible()


def abrir(rutas, etiqueta="prueba"):
    rutas = [Path(r) for r in rutas if Path(r).suffix.lower() in (".docx", ".xlsx", ".pptx")]
    if not rutas or not remoto.disponible(): return {}
    carpeta = f"nomoffice/entrada/abrir-{etiqueta}"
    remoto.ejecutar(f"rmdir /S /Q {carpeta} 2>nul & echo.".replace("/", "\\"))
    remoto.subir(rutas, carpeta)
    r = remoto.correr_office(f"abrir-{etiqueta}", "abrir_office.py", f"entrada\\abrir-{etiqueta}")
    if "[fin" not in r.stdout: return {}
    with tempfile.TemporaryDirectory() as d:
        sal = Path(d) / "resultado.json"
        if not remoto.bajar(f"{carpeta}/resultado.json", sal): return {}
        return {x["archivo"]: x for x in json.loads(sal.read_text(encoding="utf-8"))}


def bien(r):
    return bool(r) and bool(r.get("abrio")) and not r.get("avisos") and not r.get("reparado")

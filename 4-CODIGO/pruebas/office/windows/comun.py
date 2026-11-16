"""Lo común a los programas que corren en Windows con Microsoft Office de verdad.

Corren DENTRO de la sesión de escritorio (los lanza `correr.ps1` con una tarea programada): desde la
sesión de SSH, Office no puede guardar archivos. Word y Excel abren su propia copia (DispatchEx);
PowerPoint es una sola copia para todo el equipo, así que nunca se cierra con Quit (podría ser la
de la persona que usa el equipo): solo se cierran las presentaciones que abrimos.
"""
import contextlib, json, os, sys, time, traceback
import pythoncom
import win32com.client as W
from win32com.client import gencache

BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVOS = os.path.join(BASE, "archivos")
IMG = os.path.join(BASE, "imagenes")
os.makedirs(ARCHIVOS, exist_ok=True)


def rgb(r, g, b):
    """Color de Office (BGR en un entero)."""
    return r + g * 256 + b * 65536


def hexa(c):
    c = c.lstrip("#")
    return rgb(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))


class Bitacora:
    """Anota cada función que se intenta: si Office la aceptó o qué error dio."""
    def __init__(self, nombre):
        self.nombre, self.ok, self.mal = nombre, [], []

    @contextlib.contextmanager
    def paso(self, que):
        try:
            yield
            self.ok.append(que)
        except Exception as e:
            msg = str(e)
            if hasattr(e, "excepinfo") and e.excepinfo and e.excepinfo[2]: msg = e.excepinfo[2]
            self.mal.append((que, msg[:300]))
            print(f"  [falla] {que}: {msg[:300]}", flush=True)

    def resumen(self):
        print(f"{self.nombre}: {len(self.ok)} funciones bien, {len(self.mal)} con falla", flush=True)
        return {"archivo": self.nombre, "ok": self.ok, "mal": self.mal}


def word():
    w = gencache.EnsureDispatch(W.DispatchEx("Word.Application"))
    w.Visible = False
    w.DisplayAlerts = 0
    w.ScreenUpdating = False
    try: w.Options.CheckSpellingAsYouType = False; w.Options.CheckGrammarAsYouType = False
    except Exception: pass
    return w


def excel():
    x = gencache.EnsureDispatch(W.DispatchEx("Excel.Application"))
    x.Visible = False
    x.DisplayAlerts = False
    x.ScreenUpdating = False
    return x


def powerpoint():
    return gencache.EnsureDispatch("PowerPoint.Application")


def guardar_json(ruta, datos):
    with open(ruta, "w", encoding="utf8") as f: json.dump(datos, f, ensure_ascii=False, indent=1)


TEXTOS = [
    "El equipo revisó el avance del proyecto durante la reunión del lunes y acordó entregar el informe final antes del viernes.",
    "La colaboración en tiempo real permite que varias personas editen el mismo documento sin perder ningún cambio, incluso cuando la conexión falla.",
    "Cada sección del documento incluye tablas, imágenes y referencias que deben conservarse exactamente igual al abrirlo en otro programa.",
    "Según los resultados de la encuesta aplicada a 248 estudiantes, el 73.5 % prefiere trabajar en equipo cuando las herramientas son sencillas.",
    "Para calcular el promedio ponderado se multiplica cada calificación por su peso, se suman los productos y el total se divide entre la suma de los pesos.",
    "La biblioteca del plantel cuenta con más de doce mil volúmenes, de los cuales una tercera parte corresponde a obras de consulta y divulgación científica.",
    "Durante la práctica de laboratorio se midió la temperatura cada cinco minutos y se registraron los datos en una hoja de cálculo compartida.",
    "Los acuerdos de la sesión anterior quedaron pendientes porque faltó la firma de la coordinación académica y del área de servicios escolares.",
    "Un documento bien estructurado facilita la lectura: los títulos ordenan las ideas, las listas resumen los pasos y las tablas comparan datos.",
    "La propuesta considera tres etapas: análisis de requisitos, construcción del prototipo y pruebas con usuarios reales en distintos dispositivos.",
    "En la ciudad de México, la temporada de lluvias comienza en junio y termina a finales de septiembre, con tormentas intensas por la tarde.",
    "El presupuesto aprobado asciende a $185,420.50 MXN, distribuidos entre equipo de cómputo, licencias, capacitación y gastos de operación.",
]


def texto(i, n=1):
    """n oraciones seguidas a partir de la i (dan vueltas por la lista)."""
    return " ".join(TEXTOS[(i + k) % len(TEXTOS)] for k in range(n))

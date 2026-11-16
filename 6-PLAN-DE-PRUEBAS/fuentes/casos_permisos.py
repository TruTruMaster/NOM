#!/usr/bin/env python3
"""Tabla de permisos del plan de pruebas (sección 6.3): 6 acciones por 5 situaciones.

    python3 casos_permisos.py [http://localhost:8099]

Crea un espacio con una propietaria, una editora y una lectora, más una cuenta sin acceso, y con
cada una intenta las seis acciones. "Sí" es una respuesta 200 o 201. Solo usa la biblioteca estándar.
"""
import json, random, string, sys, urllib.request, urllib.error

B = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8099"
CL = "Clave-Segura-123"
suf = "".join(random.choice(string.ascii_lowercase) for _ in range(5))


def r(m, u, b=None, t=None):
    h = {"Content-Type": "application/json", "User-Agent": "NOM-COLAB-pruebas/1.0"}
    if t: h["Authorization"] = "Bearer " + t
    q = urllib.request.Request(B + u, method=m, data=json.dumps(b).encode() if b is not None else None, headers=h)
    try:
        with urllib.request.urlopen(q, timeout=30) as x:
            c = x.read(); return x.status, (json.loads(c) if c[:1] in b"{[" else None)
    except urllib.error.HTTPError as e:
        return e.code, None


tok = {}
for n in ("duena", "editora", "lectora", "ajena"):
    r("POST", "/api/cuentas/registro", {"nombre": n + suf, "contrasena": CL, "acepto": "2026-11-13"})
    tok[n] = r("POST", "/api/cuentas/sesion", {"nombre": n + suf, "contrasena": CL})[1]["token"]
tok["sin sesión"] = None
esp = r("POST", "/api/espacios", {"nombre": "Permisos " + suf}, tok["duena"])[1]["id"]
r("PUT", f"/api/espacios/{esp}/accesos", {"usuario": "editora" + suf, "permiso": "EDICION"}, tok["duena"])
r("PUT", f"/api/espacios/{esp}/accesos", {"usuario": "lectora" + suf, "permiso": "LECTURA"}, tok["duena"])
n = [0]


def doc_nuevo():
    n[0] += 1
    return r("POST", f"/api/espacios/{esp}/documentos", {"nombre": f"Doc {n[0]}", "modo": "TEXTO", "ops": []}, tok["duena"])[1]["id"]


def escribir(t, d):
    n[0] += 1
    return r("POST", f"/api/documentos/{d}/ops", {"ops": [{"id": f"x:{n[0]}", "t": n[0], "n": "x", "bloque": "cuerpo", "campos": {}, "s": {"k": "i", "iz": None, "x": "a"}}]}, t)[0]


doc = doc_nuevo()
ACCIONES = [
    ("Abrir el espacio y sus documentos", lambda t: r("GET", f"/api/espacios/{esp}", t=t)[0], ("Sí", "Sí", "Sí", 403, 401)),
    ("Escribir en un documento", lambda t: escribir(t, doc), ("Sí", "Sí", 403, 403, 401)),
    ("Dar o quitar accesos", lambda t: r("PUT", f"/api/espacios/{esp}/accesos", {"usuario": "ajena" + suf, "permiso": None}, t)[0], ("Sí", 403, 403, 403, 401)),
    ("Crear un código de invitación", lambda t: r("POST", f"/api/espacios/{esp}/enlaces", {"permiso": "LECTURA", "horas": 1}, t)[0], ("Sí", 403, 403, 403, 401)),
    ("Mandar un documento a la papelera", lambda t: r("DELETE", f"/api/documentos/{doc_nuevo()}", t=t)[0], ("Sí", "Sí", 403, 403, 401)),
    ("Vaciar la papelera", lambda t: r("DELETE", f"/api/espacios/{esp}/papelera", t=t)[0], ("Sí", 403, 403, 403, 401)),
]
QUIEN = ("duena", "editora", "lectora", "ajena", "sin sesión")
bien = total = 0
print(f'{"Acción":<36}' + "".join(f"{q:>12}" for q in QUIEN))
for nombre, hacer, esperado in ACCIONES:
    fila = []
    for q, esp_ in zip(reversed(QUIEN), reversed(esperado)):          # la propietaria al final: sus acciones sí cambian cosas
        c = hacer(tok[q]); obt = "Sí" if c in (200, 201, 204) else c
        total += 1; bien += str(obt) == str(esp_)
        fila.append(f'{str(obt) + ("" if str(obt) == str(esp_) else "≠" + str(esp_)):>12}')
    print(f"{nombre:<36}" + "".join(reversed(fila)))
print(f"\n{bien} de {total} combinaciones con el resultado esperado")
sys.exit(0 if bien == total else 1)

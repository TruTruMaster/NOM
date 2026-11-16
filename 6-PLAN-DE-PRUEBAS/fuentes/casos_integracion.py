#!/usr/bin/env python3
"""Casos de integración del plan de pruebas (sección 6.3): los módulos trabajando juntos.

    python3 casos_integracion.py [http://localhost:8099]

Hace, en orden, lo que haría un equipo desde la página: crear cuentas, entrar, crear un espacio,
un documento, invitar con código, escribir, cambiar permisos, guardar y restaurar una versión,
usar la papelera y cerrar sesión. Compara el código de respuesta del servidor con el esperado y
mide cuánto tardan el registro y el inicio de sesión. Solo usa la biblioteca estándar.
"""
import json, random, string, sys, time, urllib.request, urllib.error

B = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8099"
CL = "Clave-Segura-123"
LEGAL = "2026-11-13"
suf = "".join(random.choice(string.ascii_lowercase) for _ in range(5))
res = []


def r(m, u, b=None, t=None):
    h = {"Content-Type": "application/json", "User-Agent": "NOM-COLAB-pruebas/1.0"}
    if t: h["Authorization"] = "Bearer " + t
    q = urllib.request.Request(B + u, method=m, data=json.dumps(b).encode() if b is not None else None, headers=h)
    t0 = time.time()
    try:
        with urllib.request.urlopen(q, timeout=30) as x:
            c = x.read(); return x.status, (json.loads(c) if c[:1] in b"{[" else None), time.time() - t0
    except urllib.error.HTTPError as e:
        return e.code, None, time.time() - t0


def caso(clave, que, esperado, obtenido, nota=""):
    bien = str(esperado) == str(obtenido)
    res.append(dict(id=clave, caso=que, esperado=esperado, obtenido=obtenido, bien=bien, nota=nota))
    print(f'{"OK " if bien else "MAL"}  {clave}  {que:<74} esperado {esperado}  obtenido {obtenido}  {nota}')


def texto(cliente, n, x):
    return {"ops": [{"id": f"{cliente}:{n}", "t": n, "n": cliente, "bloque": "cuerpo", "campos": {}, "s": {"k": "i", "iz": None, "x": x}}]}


a, b = "ana" + suf, "beto" + suf
# INT-01 Cuentas: registro y sesión
c, _, t_reg = r("POST", "/api/cuentas/registro", {"nombre": a, "contrasena": CL, "acepto": LEGAL})
r("POST", "/api/cuentas/registro", {"nombre": b, "contrasena": CL, "acepto": LEGAL})
caso("INT-01", "Registrar una cuenta nueva", 201, c, f"{t_reg * 1000:.0f} ms")
c, s, t_ses = r("POST", "/api/cuentas/sesion", {"nombre": a, "contrasena": CL})
ta = (s or {}).get("token")
caso("INT-02", "Iniciar sesión con esa cuenta y recibir la sesión", 200, c, f"{t_ses * 1000:.0f} ms")
tb = r("POST", "/api/cuentas/sesion", {"nombre": b, "contrasena": CL})[1]["token"]
# INT-03 Cuentas + espacios
c, e, _ = r("POST", "/api/espacios", {"nombre": "Equipo " + suf}, ta)
esp = (e or {}).get("id")
yo = r("GET", f"/api/espacios/{esp}", t=ta)[1] or {}
caso("INT-03", "Con la sesión, crear un espacio: quien lo crea queda como Propietario", "201 PROPIETARIO", f'{c} {yo.get("permiso") or yo.get("miPermiso") or (yo.get("espacio") or {}).get("permiso")}')
# INT-04 Espacios + documentos
c, d, _ = r("POST", f"/api/espacios/{esp}/documentos", {"nombre": "Acta", "modo": "TEXTO", "ops": []}, ta)
doc = (d or {}).get("id")
caso("INT-04", "Crear un documento de texto dentro del espacio", 201, c)
caso("INT-05", "Otra cuenta, sin invitación, intenta abrir el espacio", 403, r("GET", f"/api/espacios/{esp}", t=tb)[0])
# INT-06 Invitación con código
c, en, _ = r("POST", f"/api/espacios/{esp}/enlaces", {"permiso": "EDICION", "horas": 2}, ta)
cod = (en or {}).get("codigo") or (en or {}).get("id")
c2 = r("POST", f"/api/enlaces/{cod}/entrar", t=tb)[0]
caso("INT-06", "El Propietario crea un código de invitación y la otra cuenta entra con él", "201 200", f"{c} {c2}")
# INT-07 Edición entre dos personas
c = r("POST", f"/api/documentos/{doc}/ops", texto(b, 1, "Hola equipo"), tb)[0]
visto = json.dumps(r("GET", f"/api/documentos/{doc}", t=ta)[1], ensure_ascii=False)
caso("INT-07", "La invitada escribe en el documento y el Propietario ve el cambio", "200 y lo ve", f'{c} y {"lo ve" if "Hola equipo" in visto else "no lo ve"}')
# INT-08 Permisos aplicados a la edición
c = r("PUT", f"/api/espacios/{esp}/accesos", {"usuario": b, "permiso": "LECTURA"}, ta)[0]
c2 = r("POST", f"/api/documentos/{doc}/ops", texto(b, 2, " de nuevo"), tb)[0]
caso("INT-08", "El Propietario la cambia a solo lectura y su siguiente cambio se rechaza", "200 403", f"{c} {c2}")
caso("INT-09", "Con solo lectura todavía puede abrir el documento", 200, r("GET", f"/api/documentos/{doc}", t=tb)[0])
# INT-10 Versiones
c = r("POST", f"/api/documentos/{doc}/versiones", {"nombre": "Antes de corregir"}, ta)[0]
vs = r("GET", f"/api/documentos/{doc}/versiones", t=ta)[1]
lista = vs if isinstance(vs, list) else (vs or {}).get("versiones", [])
caso("INT-10", "Guardar una versión con nombre y verla en el historial", "200 y 1 versión", f'{c} y {min(len(lista), 1)} versión')
num = (lista[0] if lista else {}).get("numero", 1)
caso("INT-11", "Solo lectura intenta restaurar una versión", 403, r("POST", f"/api/documentos/{doc}/versiones/{num}/restaurar", t=tb)[0])
# INT-12 Papelera
c = r("DELETE", f"/api/documentos/{doc}", t=ta)[0]
c2 = r("POST", f"/api/documentos/{doc}/restaurar", t=ta)[0]
caso("INT-12", "Mandar el documento a la papelera y restaurarlo", "200 200", f"{c} {c2}")
# INT-13 Quitar acceso
c = r("PUT", f"/api/espacios/{esp}/accesos", {"usuario": b, "permiso": None}, ta)[0]
caso("INT-13", "El Propietario le quita el acceso y ya no puede abrir el espacio", "200 403", f'{c} {r("GET", f"/api/espacios/{esp}", t=tb)[0]}')
# INT-14 Cierre de sesión
r("DELETE", "/api/cuentas/sesion", t=ta)
caso("INT-14", "Al cerrar sesión, la sesión anterior deja de servir", 401, r("GET", "/api/espacios", t=ta)[0])

bien = sum(1 for x in res if x["bien"])
print(f"\n{bien} de {len(res)}  ·  registro {t_reg * 1000:.0f} ms  ·  inicio de sesión {t_ses * 1000:.0f} ms  ·  versión", (r("GET", "/api/red/info")[1] or {}).get("version"))
json.dump(res, open(__file__.replace("fuentes/casos_integracion.py", "registros/15-nov/integracion.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
sys.exit(0 if bien == len(res) else 1)

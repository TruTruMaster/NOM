#!/usr/bin/env python3
"""Red de regresión en paralelo: corre TODAS las pruebas automáticas y deja un resumen. Igual que regresion.sh, pero
reparte las pruebas en varios carriles, cada uno con su propio servidor de NOM (su puerto y su carpeta de datos), así
que lo que tardaba más de una hora tarda lo que la prueba más larga. Corre en Linux y en Windows.

  python pruebas/regresion.py                 # todo
  python pruebas/regresion.py rapido          # sin Office ni ancho de celular
  python pruebas/regresion.py --carriles 6    # cuántas pruebas a la vez (por omisión: según los núcleos)
  python pruebas/regresion.py --solo ayuda,red   # solo esas

Necesita el jar compilado (app/target/nom-colab.jar) y la interfaz en app/src/main/resources/static (se sirve en
vivo). Los registros quedan en ~/.cache/nom-pruebas/registros/<fecha>/ y el resumen en resumen.txt de esa carpeta.
"""
import argparse, os, shutil, socket, subprocess, sys, tempfile, threading, time, urllib.request
from datetime import datetime
from pathlib import Path

AQUI = Path(__file__).resolve().parent          # 4-CODIGO/pruebas
CODIGO = AQUI.parent                             # 4-CODIGO
RAIZ = CODIGO.parent
CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache" / "nom-pruebas"))
WIN = os.name == "nt"
PY = str(CACHE / "venv" / ("Scripts/python.exe" if WIN else "bin/python"))
if not Path(PY).exists(): PY = sys.executable
JAVA = os.environ.get("JAVA") or shutil.which("java") or "java"
JAR = CODIGO / "app" / "target" / "nom-colab.jar"
STATIC = CODIGO / "app" / "src" / "main" / "resources" / "static"
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(CACHE / "navegadores"))


def libre(desde):
    p = desde
    while p < desde + 400:
        with socket.socket() as s:
            try: s.bind(("127.0.0.1", p)); return p
            except OSError: p += 1
    raise RuntimeError("sin puertos libres")


class Servidor:
    """Un servidor de NOM para un carril: su puerto, su carpeta de datos y la interfaz en vivo."""
    def __init__(self, n):
        self.puerto = libre(8110 + n * 3)
        self.datos = Path(tempfile.mkdtemp(prefix=f"nom-regresion-{n}-"))
        env = dict(os.environ, NOM_PUERTO=str(self.puerto), NOM_DATOS=str(self.datos),
                   SPRING_WEB_RESOURCES_STATIC_LOCATIONS=STATIC.as_uri() + "/")
        self.log = open(self.datos / "servidor.log", "w")
        self.p = subprocess.Popen([JAVA, "-jar", str(JAR)], env=env, stdout=self.log, stderr=subprocess.STDOUT)
        self.base = f"http://localhost:{self.puerto}"

    def esperar(self, s=120):
        t0 = time.time()
        while time.time() - t0 < s:
            try:
                urllib.request.urlopen(self.base + "/api/red/info", timeout=3).read(); return True
            except Exception: time.sleep(0.5)
        return False

    def apagar(self):
        try: self.p.terminate(); self.p.wait(20)
        except Exception:
            try: self.p.kill()
            except Exception: pass
        self.log.close()
        shutil.rmtree(self.datos, ignore_errors=True)


def pruebas(modo, out):
    """(nombre, comando, ¿necesita servidor?, segundos que suele tardar). Las que arrancan su propio servidor no usan
    el del carril. El orden de esta lista no importa: se empieza por las más largas."""
    cap = str(out / "capturas")
    i = lambda s: str(AQUI / "interfaz" / s)
    o = lambda s: str(AQUI / "office" / s)
    L = [
        ("junit", ["mvn", "-o", "-q", "test"], False, 30, CODIGO / "app"),
        ("seguridad", [sys.executable, str(AQUI / "seguridad.py"), str(JAR)], False, 190, None),
        ("limites", [sys.executable, str(RAIZ / "6-PLAN-DE-PRUEBAS" / "fuentes" / "prueba_limites.py")], False, 40, None),
        ("carga", [sys.executable, str(AQUI / "carga.py"), "--personas", "5", "--segundos", "10"], True, 20, None),
        ("cuentas_equipo", [PY, i("prueba_cuentas_equipo.py"), cap], True, 115, None),
        ("sin_internet", [PY, i("prueba_sin_internet.py"), cap], True, 100, None),
        ("e2e", [PY, i("prueba_e2e.py"), cap], True, 50, None),
        ("ayuda", [PY, i("prueba_ayuda.py"), cap], True, 15, None),
        ("red", [PY, i("prueba_red.py"), cap], True, 35, None),
        ("zip_espacio", [PY, i("prueba_zip_espacio.py"), cap], True, 40, None),
        ("presentacion", [PY, i("prueba_presentacion.py"), cap], True, 40, None),
        ("funciones", [PY, i("prueba_funciones.py"), cap], True, 225, None),
        ("rastreo", [PY, i("rastreo.py"), cap], True, 1230, None),
    ] + [(t, [PY, i(t + ".py"), cap], True, 30, None) for t in
         ("prueba_pres1", "prueba_pptx", "prueba_pptx2", "prueba_diag1", "prueba_85", "prueba_citas", "prueba_datos", "prueba_biblioteca", "prueba_fuentes_api", "prueba_acomodo", "prueba_marcas", "prueba_enfoque", "prueba_clasificador")]
    if modo != "rapido":
        R = CACHE / "corpus-reales"
        corpus = [str(p) for d in ("corpus", "corpus-web", "corpus-reales") for p in sorted((CACHE / d).glob("*/*"))]
        L += [
            ("funciones_movil", [PY, i("prueba_funciones.py"), cap, "--movil"], True, 260, None),
            ("rastreo_movil", [PY, i("rastreo.py"), cap, "--movil"], True, 1100, None),
            ("diag_movil", [PY, i("prueba_diag_movil.py"), cap], True, 40, None),
            ("funciones_firefox", [PY, i("prueba_funciones.py"), cap], True, 300, {"NOM_NAVEGADOR": "firefox"}),
            ("office_fidelidad", [PY, o("fidelidad.py")] + corpus, True, 600, None),
            ("office_modelo", [PY, o("fidelidad.py"), "--modelo"] + corpus, True, 600, None),
            ("office_ediciones", [PY, o("ediciones.py"), "--office"], True, 500, None),
            ("documento_grande", [PY, o("documento_grande.py"), str(CACHE / "corpus/docx/ms-informe-gobernanza.docx"), str(out / "documento-grande")], True, 300, None),
            ("editar_real", [PY, o("editar_real.py"), str(out / "editar-real"), str(R / "xlsx/ms-plantilla-presupuesto-universitario.xlsx"),
                             str(R / "xlsx/ms-plantilla-tablas-dinamicas.xlsx"), str(R / "pptx/ms-plantilla-presentacion-basica.pptx"),
                             str(R / "pptx/lo-group-rot.pptx"), str(CACHE / "corpus/docx/ms-complex2010.docx")], True, 300, None),
            ("office_raros", [PY, o("raros.py")], True, 120, None),
            ("funciones_excel", [PY, o("funciones_excel.py")], True, 120, None),
            ("almacen_binarios", [PY, o("almacen_binarios.py")], True, 200, None),
            ("campos_word", [PY, o("campos_real.py")], True, 120, None),
            ("metarchivos_word", [PY, o("metarchivos_real.py")], False, 300, None),
        ]
    return L


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("modo", nargs="?", default="todo")
    ap.add_argument("--carriles", type=int, default=max(2, min(8, (os.cpu_count() or 4) // 2)))
    ap.add_argument("--solo", default="")
    ap.add_argument("--sin-junit", action="store_true", help="sin las pruebas de Java (las corre quien compila el jar)")
    a = ap.parse_args()
    if not JAR.exists(): print("Falta el jar: cd 4-CODIGO/app && mvn -o -q package"); return 2
    out = CACHE / "registros" / datetime.now().strftime("%Y-%m-%d_%H%M"); (out / "capturas").mkdir(parents=True, exist_ok=True)
    res = out / "resumen.txt"
    lista = pruebas(a.modo, out)
    # cada prueba con su carpeta de capturas (corren a la vez: con una sola se pisaban los archivos del mismo nombre)
    cap = str(out / "capturas")
    lista = [(n, [(x + os.sep + n) if x == cap else x for x in cmd], srv, seg, extra) for n, cmd, srv, seg, extra in lista]
    for n, cmd, *_ in lista:
        if any(x.startswith(cap + os.sep) for x in cmd): (out / "capturas" / n).mkdir(parents=True, exist_ok=True)
    if a.solo: lista = [x for x in lista if x[0] in a.solo.split(",")]
    if a.sin_junit or not shutil.which("mvn"): lista = [x for x in lista if x[0] != "junit"]
    # las de Office necesitan el validador al final (lee lo que dejan ediciones.py y editar_real.py)
    validador = a.modo != "rapido" and not a.solo and (Path.home() / ".dotnet" / ("dotnet.exe" if WIN else "dotnet")).exists()
    lista.sort(key=lambda x: -x[3])
    con_srv = sum(1 for x in lista if x[2])
    nsrv = min(a.carriles, con_srv) if con_srv else 0
    print(f"{len(lista)} pruebas en {a.carriles} carriles ({nsrv} servidores)", flush=True)
    servidores = [Servidor(n) for n in range(nsrv)]
    try:
        for s in servidores:
            if not s.esperar(): print(f"El servidor del puerto {s.puerto} no arrancó (ver {s.datos}/servidor.log)"); return 2
        info = urllib.request.urlopen(servidores[0].base + "/api/red/info").read().decode() if servidores else ""
        cab = f"Versión: {info.split(chr(34) + 'version' + chr(34) + ':' + chr(34))[1].split(chr(34))[0] if 'version' in info else '?'}  ·  {datetime.now():%c}  ·  {a.carriles} carriles"
        res.write_text(cab + "\n", encoding="utf-8"); print(cab, flush=True)
        libres = list(servidores)
        candado = threading.Lock()
        pendientes = list(lista)
        filas = []
        t0g = time.time()

        def trabajar():
            while True:
                with candado:
                    if not pendientes: return
                    # la más larga que pueda correr ahora (si necesita servidor, que haya uno libre)
                    k = next((j for j, x in enumerate(pendientes) if not x[2] or libres), None)
                    if k is None: espera = True
                    else: espera = False; nombre, cmd, srv, _, extra = pendientes.pop(k); s = libres.pop(0) if srv else None
                if espera: time.sleep(0.5); continue
                env = dict(os.environ)
                if s: env["NOM_BASE"] = s.base
                cwd = str(CODIGO)
                if isinstance(extra, dict): env.update(extra)
                elif isinstance(extra, Path): cwd = str(extra)
                t0 = time.time()
                with open(out / f"{nombre}.txt", "w", encoding="utf-8", errors="replace") as f:
                    try: c = subprocess.run(cmd, cwd=cwd, env=env, stdout=f, stderr=subprocess.STDOUT, timeout=3600).returncode
                    except subprocess.TimeoutExpired: c = 124
                    except FileNotFoundError as e: f.write(str(e)); c = 127
                t = int(time.time() - t0)
                texto = (out / f"{nombre}.txt").read_text(encoding="utf-8", errors="replace").splitlines()
                import re
                ult = " ".join([l for l in texto if re.search(r"de [0-9]+|aprobad|correct|resultado|fall|tests run", l, re.I)][-2:])[:150]
                fila = f"{nombre:<26} {'BIEN' if c == 0 else 'FALLA':<6} {t:4d}s  {ult}"
                with candado:
                    filas.append(fila); print(fila, flush=True)
                    with open(res, "a", encoding="utf-8") as r: r.write(fila + "\n")
                    if s: libres.append(s)

        hilos = [threading.Thread(target=trabajar) for _ in range(a.carriles)]
        for h in hilos: h.start()
        for h in hilos: h.join()
        if validador:
            t0 = time.time()
            with open(out / "office_openxml.txt", "w", encoding="utf-8") as f:
                c = subprocess.run([sys.executable, str(AQUI / "office" / "validar_openxml.py"), str(CACHE / "office-ediciones")], cwd=str(CODIGO), stdout=f, stderr=subprocess.STDOUT).returncode
            fila = f"{'office_openxml':<26} {'BIEN' if c == 0 else 'FALLA':<6} {int(time.time() - t0):4d}s"
            filas.append(fila); print(fila)
            with open(res, "a", encoding="utf-8") as r: r.write(fila + "\n")
        fallas = sum(1 for f in filas if " FALLA " in f)
        fin = f"Total {int(time.time() - t0g)} s · {fallas} suite(s) con falla · {res}"
        print(fin)
        with open(res, "a", encoding="utf-8") as r: r.write(fin + "\n")
        return 1 if fallas else 0
    finally:
        for s in servidores: s.apagar()


if __name__ == "__main__":
    sys.exit(main())

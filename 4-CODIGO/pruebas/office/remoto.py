"""La computadora con Windows y Office: por SSH desde Linux, o la misma computadora si la prueba ya corre en Windows.

  import remoto
  remoto.ejecutar("dir")                    # un comando de cmd en Windows (por SSH o aquí mismo)
  remoto.subir([archivos], "nomoffice/x")   # copia archivos a esa carpeta (relativa a la carpeta del usuario)
  remoto.bajar("nomoffice/x/y.json", destino)
  remoto.correr_office(id, script, *args)   # corre un programa de Office en la sesión de escritorio (correr.ps1)
"""
import os, shutil, subprocess
from pathlib import Path

WIN = os.environ.get("NOM_WINDOWS", "user@100.65.29.5")
AQUI_WIN = os.name == "nt"
CASA = Path.home()


def ejecutar(cmd, timeout=3600):
    if AQUI_WIN:
        return subprocess.run(cmd, shell=True, cwd=str(CASA), capture_output=True, text=True, errors="replace", timeout=timeout)
    return subprocess.run(["ssh", "-o", "BatchMode=yes", "-o", "ConnectTimeout=10", WIN, cmd], capture_output=True, text=True, errors="replace", timeout=timeout)


def disponible():
    if AQUI_WIN: return True
    try: return ejecutar("echo nom", timeout=30).stdout.strip() == "nom"
    except Exception: return False


def subir(archivos, carpeta, timeout=900):
    """Copia los archivos a <carpeta del usuario>/<carpeta> (por SSH, en un solo paquete)."""
    archivos = [Path(a) for a in archivos]
    if AQUI_WIN:
        d = CASA / carpeta; d.mkdir(parents=True, exist_ok=True)
        for a in archivos: shutil.copy2(a, d / a.name)
        return
    import tarfile, tempfile
    with tempfile.TemporaryDirectory() as t:
        paq = Path(t) / "ida.tar"
        with tarfile.open(paq, "w") as tf:
            for a in archivos: tf.add(a, arcname=a.name)
        ejecutar(f'mkdir "{carpeta}" 2>nul & echo.'.replace("/", "\\"))
        subprocess.run(["scp", "-q", str(paq), f"{WIN}:{carpeta}/ida.tar"], check=True, timeout=timeout)
        ejecutar(f'cd "{carpeta}" && tar -xf ida.tar && del ida.tar'.replace("/", "\\"))


def bajar(ruta, destino, timeout=900):
    destino = Path(destino)
    if AQUI_WIN:
        o = CASA / ruta
        if o.exists(): shutil.copy2(o, destino)
        return destino.exists()
    subprocess.run(["scp", "-q", f"{WIN}:{ruta}", str(destino)], check=False, timeout=timeout)
    return destino.exists()


def correr_office(ident, script, *args, timeout=3600):
    """Corre nomoffice/<script> en la sesión de escritorio (Office solo guarda archivos ahí) y espera. Copia antes el
    programa y comun.py de pruebas/office/windows."""
    w = Path(__file__).resolve().parent / "windows"
    subir([w / script, w / "comun.py", w / "correr.ps1", w / "lanzador.py"], "nomoffice")
    return ejecutar("powershell -NoProfile -ExecutionPolicy Bypass -File nomoffice\\correr.ps1 " + " ".join([ident, script] + list(args)), timeout=timeout)

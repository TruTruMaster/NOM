#!/usr/bin/env python3
"""Maneja Microsoft Office de verdad en la computadora con Windows (por SSH) desde Linux.

  office_real.py enviar                 copia pruebas/office/windows/ e imágenes a Windows
  office_real.py correr ID script.py …  corre un programa de Office en la sesión de escritorio de Windows
  office_real.py traer PATRON [DESTINO] trae archivos de nomoffice\\archivos (por defecto a ~/.cache/nom-pruebas/office-real/archivos)
  office_real.py subir ARCHIVOS…        sube archivos a nomoffice\\entrada

La computadora se elige con NOM_WINDOWS (por defecto user@100.65.29.5, la de Tailscale).
Office solo guarda archivos si corre dentro de la sesión de escritorio: `correr.ps1` lo lanza con
una tarea programada y espera a que termine (ver windows/comun.py).
"""
import os, subprocess, sys
from pathlib import Path

WIN = os.environ.get("NOM_WINDOWS", "user@100.65.29.5")
AQUI = Path(__file__).parent / "windows"
CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"


def ssh(cmd, timeout=3600):
    return subprocess.run(["ssh", "-o", "BatchMode=yes", WIN, cmd], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout)


def enviar():
    ssh("mkdir nomoffice\\imagenes nomoffice\\entrada nomoffice\\archivos 2>nul & echo.")
    archivos = [str(p) for p in AQUI.iterdir() if p.suffix in (".py", ".ps1")]
    subprocess.run(["scp", "-q"] + archivos + [WIN + ":nomoffice/"], check=True)
    img = CACHE / "imagenes"
    if not img.exists(): subprocess.run([sys.executable, str(AQUI / "imagenes.py"), str(img)], check=True)
    subprocess.run(["scp", "-q"] + [str(p) for p in img.iterdir()] + [WIN + ":nomoffice/imagenes/"], check=True)


def correr(ident, *args):
    r = ssh("powershell -NoProfile -ExecutionPolicy Bypass -File nomoffice\\correr.ps1 " + ident + " " + " ".join(args))
    sys.stdout.write(r.stdout); sys.stderr.write(r.stderr)
    return r


def traer(patron, destino=None):
    destino = Path(destino or CACHE / "archivos"); destino.mkdir(parents=True, exist_ok=True)
    subprocess.run(["scp", "-q", WIN + ":nomoffice/archivos/" + patron, str(destino) + "/"], check=False)
    return destino


def subir(*archivos):
    ssh("mkdir nomoffice\\entrada 2>nul & echo.")
    subprocess.run(["scp", "-q"] + list(archivos) + [WIN + ":nomoffice/entrada/"], check=True)


if __name__ == "__main__":
    a = sys.argv[1:]
    if not a: print(__doc__); sys.exit(1)
    {"enviar": enviar, "correr": correr, "traer": traer, "subir": subir}[a[0]](*a[1:])

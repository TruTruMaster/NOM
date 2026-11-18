"""¿Lo abre Office de verdad sin avisos ni «reparar»? Para las pruebas que descargan archivos de NOM.

  import office_abre
  res = office_abre.abrir(["/ruta/a.docx", "/ruta/b.xlsx"], "ediciones")   # {nombre: {"abrio":…, "avisos":[…], …}}
  office_abre.bien(res["a.docx"])                                        # True si abrió sin avisos ni reparar

Manda todos los archivos de una vez (un solo paquete por SSH) a la computadora con Windows y corre
windows/abrir_office.py en la sesión de escritorio: Word, Excel o PowerPoint abren cada uno CON sus avisos y se anota
si pidieron reparar. Sin la computadora con Windows (o sin red), devuelve {} y quien llama decide.
"""
import json, os, subprocess, tarfile, tempfile
from pathlib import Path

WIN = os.environ.get("NOM_WINDOWS", "user@100.65.29.5")
AQUI = Path(__file__).resolve().parent


def _ssh(cmd, timeout=3600):
    return subprocess.run(["ssh", "-o", "BatchMode=yes", "-o", "ConnectTimeout=10", WIN, cmd], capture_output=True, text=True, errors="replace", timeout=timeout)


def disponible():
    try: return _ssh("echo nom", timeout=30).stdout.strip() == "nom"
    except Exception: return False


def abrir(rutas, etiqueta="prueba"):
    rutas = [Path(r) for r in rutas if Path(r).suffix.lower() in (".docx", ".xlsx", ".pptx")]
    if not rutas or not disponible(): return {}
    carpeta = f"entrada\\abrir-{etiqueta}"
    with tempfile.TemporaryDirectory() as d:
        paquete = Path(d) / "ida.tar"
        with tarfile.open(paquete, "w") as t:
            for r in rutas: t.add(r, arcname=r.name)
        _ssh(f"rmdir /S /Q nomoffice\\{carpeta} 2>nul & mkdir nomoffice\\{carpeta} & echo.")
        subprocess.run(["scp", "-q", str(paquete), f"{WIN}:nomoffice/{carpeta.replace(chr(92), '/')}/ida.tar"], check=True, timeout=900)
        subprocess.run(["scp", "-q", str(AQUI / "windows" / "abrir_office.py"), str(AQUI / "windows" / "comun.py"), WIN + ":nomoffice/"], check=True, timeout=120)
        _ssh(f"cd nomoffice\\{carpeta} && tar -xf ida.tar && del ida.tar")
        r = _ssh(f"powershell -NoProfile -ExecutionPolicy Bypass -File nomoffice\\correr.ps1 abrir-{etiqueta} abrir_office.py {carpeta}")
        if "[fin" not in r.stdout: return {}
        sal = Path(d) / "resultado.json"
        subprocess.run(["scp", "-q", f"{WIN}:nomoffice/{carpeta.replace(chr(92), '/')}/resultado.json", str(sal)], check=False, timeout=300)
        if not sal.exists(): return {}
        return {x["archivo"]: x for x in json.loads(sal.read_text(encoding="utf-8"))}


def bien(r):
    return bool(r) and bool(r.get("abrio")) and not r.get("avisos") and not r.get("reparado")

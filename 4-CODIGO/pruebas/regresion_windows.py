#!/usr/bin/env python3
"""Corre la regresión completa en la computadora con Windows (20 hilos, 32 GB, y Office 2024 ahí mismo) en vez de en
esta: manda el código de ahora (lo que hay en la carpeta, aunque no tenga commit), el jar y los archivos de prueba,
corre pruebas/regresion.py en paralelo allá y trae el resumen y los registros.

  python3 pruebas/regresion_windows.py [rapido] [--carriles 10] [--solo a,b]

Antes, una sola vez: pruebas/office/windows/preparar_windows.ps1 en Windows (Java 17, Python con Playwright, .NET 8).
El jar se compila aquí (mvn package); junit también corre aquí (Windows no tiene Maven).
"""
import argparse, os, subprocess, sys, tarfile, tempfile, time
from pathlib import Path

AQUI = Path(__file__).resolve().parent
CODIGO = AQUI.parent
RAIZ = CODIGO.parent
CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache" / "nom-pruebas"))
WIN = os.environ.get("NOM_WINDOWS", "user@100.65.29.5")


def ssh(cmd, timeout=600):
    return subprocess.run(["ssh", "-o", "BatchMode=yes", WIN, cmd], capture_output=True, text=True, errors="replace", timeout=timeout)


def empacar(destino):
    salta = ("__pycache__", "/target/", "node_modules", ".gradle")
    with tarfile.open(destino, "w:gz") as t:
        def agregar(p, como):
            for f in Path(p).rglob("*"):
                s = str(f)
                if f.is_file() and not any(x in s.replace("\\", "/") for x in salta): t.add(f, arcname=como + "/" + str(f.relative_to(p)).replace("\\", "/"))
        for d in ("4-CODIGO/app/src", "4-CODIGO/nucleo", "4-CODIGO/pruebas", "4-CODIGO/ayuda", "4-CODIGO/herramientas", "6-PLAN-DE-PRUEBAS/fuentes"):
            agregar(RAIZ / d, "NOM/" + d)
        t.add(CODIGO / "app/target/nom-colab.jar", arcname="NOM/4-CODIGO/app/target/nom-colab.jar")
        for d in ("corpus", "corpus-web", "corpus-reales", "openxml-bin"):
            if (CACHE / d).exists(): agregar(CACHE / d, ".cache/nom-pruebas/" + d)
        for nombre in ("t03-todo-powerpoint.pptx", "w16-numlineas.docx", "w16-numlineas.pdf"):
            f = CACHE / "office-real/archivos" / nombre
            if f.exists(): t.add(f, arcname=".cache/nom-pruebas/office-real/archivos/" + f.name)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("modo", nargs="?", default="todo")
    ap.add_argument("--carriles", type=int, default=10)
    ap.add_argument("--solo", default="")
    a = ap.parse_args()
    t0 = time.time()
    print("== jar y junit (aquí)", flush=True)
    j = subprocess.run(["mvn", "-o", "-q", "package"], cwd=CODIGO / "app", capture_output=True, text=True)
    print("junit y jar:", "BIEN" if j.returncode == 0 else "FALLA\n" + (j.stdout + j.stderr)[-1500:], flush=True)
    if j.returncode != 0: return 1
    with tempfile.TemporaryDirectory() as d:
        paq = Path(d) / "nom.tgz"
        empacar(paq)
        print(f"== a Windows ({paq.stat().st_size // 1024 // 1024} MB)", flush=True)
        ssh("rmdir /S /Q NOM 2>nul & echo.")
        subprocess.run(["scp", "-q", str(paq), WIN + ":nom.tgz"], check=True, timeout=1800)
        r = ssh("tar -xzf nom.tgz && del nom.tgz && echo ok", timeout=900)
        if "ok" not in r.stdout: print(r.stdout, r.stderr); return 1
    # las pruebas de Office (abrir lo descargado, Excel) corren ahí mismo, sin SSH
    jdk = ssh('powershell -NoProfile -Command "(Get-ChildItem $env:USERPROFILE\\nomoffice\\herramientas -Directory -Filter jdk-17* | Select-Object -First 1).FullName"').stdout.strip()
    args = " ".join([a.modo, "--carriles", str(a.carriles)] + (["--solo", a.solo] if a.solo else []) + ["--sin-junit"])
    cmd = (f'set "JAVA={jdk}\\bin\\java.exe" && set "PATH={jdk}\\bin;%USERPROFILE%\\.dotnet;%PATH%" && set "PYTHONUTF8=1" && '
           f'set "PLAYWRIGHT_BROWSERS_PATH=%USERPROFILE%\\.cache\\nom-pruebas\\navegadores" && python NOM\\4-CODIGO\\pruebas\\regresion.py {args}')
    print("== regresión en Windows", flush=True)
    p = subprocess.Popen(["ssh", "-o", "BatchMode=yes", WIN, cmd], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, errors="replace")
    for linea in p.stdout: print(linea.rstrip(), flush=True)
    p.wait()
    # el registro de esa corrida
    ult = ssh('powershell -NoProfile -Command "(Get-ChildItem $env:USERPROFILE\\.cache\\nom-pruebas\\registros -Directory | Sort-Object Name | Select-Object -Last 1).Name"').stdout.strip()
    if ult:
        dest = CACHE / "registros" / ("windows-" + ult); dest.mkdir(parents=True, exist_ok=True)
        ssh(f"cd .cache\\nom-pruebas\\registros\\{ult} && tar -cf ..\\{ult}.tar *.txt")
        subprocess.run(["scp", "-q", f"{WIN}:.cache/nom-pruebas/registros/{ult}.tar", str(dest / "r.tar")], check=False)
        if (dest / "r.tar").exists():
            with tarfile.open(dest / "r.tar") as t: t.extractall(dest, filter="data")
            (dest / "r.tar").unlink()
        print(f"Registros: {dest}")
    print(f"Total {int(time.time() - t0)} s")
    return p.returncode


if __name__ == "__main__":
    sys.exit(main())

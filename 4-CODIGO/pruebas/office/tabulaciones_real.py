#!/usr/bin/env python3
"""Tabulaciones puestas desde NOM (Más → Formato → Tabulaciones…) contra Word 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/tabulaciones_real.py

1. Un documento hecho en NOM con paradas de los cinco tipos (izquierda, centro, derecha, decimal y barra), con y sin
   relleno, y las predeterminadas cada 1.5 cm; una se pone con el cuadro (Fijar y Aceptar).
2. Descargado, Word lee las mismas paradas en cada párrafo (windows/tabs_ver.py) y las mismas predeterminadas.
3. Word lo pinta (windows/verdad.py) y comparar_real.py dice cuántos renglones pinta NOM igual que Word.
"""
import base64, json, os, subprocess, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas"))
REAL = CACHE / "office-real"
NOMBRE = "w19-tabs-nom.docx"
PT = 28.3465
PARRAFOS = [
    ("Producto\tCantidad\tPrecio", [[85.05, "i"], [226.8, "c"], [396.85, "d", "dot"]]),
    ("Manzana roja\t12\t$1,250.50", [[85.05, "i"], [226.8, "c"], [396.85, "p", "dot"]]),
    ("Pera\t3\t$7.05", [[85.05, "i"], [226.8, "c"], [396.85, "p", "dot"]]),
    ("Capítulo 1. Introducción\t15", [[453.55, "d", "dot"]]),
    ("Capítulo 2. Métodos y materiales\t27", [[453.55, "d", "hyphen"]]),
    ("Firma\tFecha", [[226.8, "i", "underscore"], [453.55, "d", "underscore"]]),
    ("uno\tdos\ttres\tcuatro\tcinco", []),
    ("Izquierda\tcentro de todo\tderecha", [[141.75, "c", "middleDot"], [453.55, "d"]]),
    ("Columna A\tColumna B", [[198.45, "b"], [212.6, "i"]]),
]
ALIN = {"i": 0, "c": 1, "d": 2, "p": 3, "b": 4}
REL = {None: 0, "dot": 1, "hyphen": 2, "underscore": 3, "middleDot": 5}
res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:600]}"), flush=True)


def main():
    import remoto
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Tabulaciones")
    nav = N.Navegador(movil=False); nav.entrar(a); p = nav.p
    try:
        d = N.documento(a["token"], esp, "TEXTO", "Tabulaciones")
        p.goto(N.B + "/#/doc/" + d); p.wait_for_selector(".wd [contenteditable=true]", timeout=30000); p.wait_for_timeout(600)
        p.locator(".wd [contenteditable=true]").first.click(); p.wait_for_timeout(200)
        for i, (t, _) in enumerate(PARRAFOS):
            if i: p.keyboard.press("Enter")
            for k, trozo in enumerate(t.split("\t")):
                if k: p.keyboard.press("Tab")
                p.keyboard.insert_text(trozo)
        p.wait_for_timeout(600)
        # las paradas de cada párrafo (la del primero, con el cuadro)
        for i, (_, tabs) in enumerate(PARRAFOS):
            p.keyboard.press("Control+Home")
            for _ in range(i): p.keyboard.press("ArrowDown")
            if i == 0:
                p.evaluate("() => NOM.editorActual.modoApi.tabulaciones.dialogo()"); p.wait_for_timeout(500)
                dlg = p.locator("section.hoja[role=dialog]").last
                for pos, al, *rel in tabs:
                    dlg.locator("[data-pos]").fill(f"{pos / PT:.2f}"); dlg.locator("[data-al]").select_option(al)
                    dlg.locator("[data-rel]").select_option(rel[0] if rel else ""); dlg.locator("[data-fijar]").click(); p.wait_for_timeout(150)
                dlg.locator("[data-def]").fill("1.5")
                dlg.locator("[data-si]").click(); p.wait_for_timeout(500)
            elif tabs:
                p.evaluate("(l) => NOM.editorActual.modoApi.tabulaciones.poner(l)", tabs); p.wait_for_timeout(300)
        nom = p.evaluate("() => NOM.textoRico.parrafosDe(NOM.editorActual.doc).map((q) => q.pa.tabs || [])")
        anota("el cuadro deja las paradas que se fijaron", [[round(x[0], 1), x[1], x[2] if len(x) > 2 else None] for x in nom[0]] == [[round(x[0], 1), x[1], x[2] if len(x) > 2 else None] for x in PARRAFOS[0][1]], nom[0])
        destino = REAL / "archivos" / NOMBRE
        with p.expect_download(timeout=120000) as dl:
            p.evaluate("(n) => NOM.formatos.exportarListo(NOM.editorActual.doc, 'docx', n)", NOMBRE[:-5])
        dl.value.save_as(destino)
        remoto.subir([destino], "nomoffice/archivos")
        remoto.correr_office("tabs-ver", "tabs_ver.py", f"archivos\\{NOMBRE}", "tabs-ver.json", timeout=300)
        tmp = REAL / "tabs-ver.json"; remoto.bajar("nomoffice/tabs-ver.json", tmp)
        w = json.loads(tmp.read_text(encoding="utf-8"))
        anota("Word: las predeterminadas cada 1.5 cm", abs(w.get("predeterminada", 0) - 1.5 * PT) < 0.3, w.get("predeterminada"))
        malos = []
        for i, (t, tabs) in enumerate(PARRAFOS):
            wp = w["parrafos"][i]
            exp = [[x[0], ALIN[x[1]], REL[x[2] if len(x) > 2 else None]] for x in tabs]
            got = [[x[0], x[1], x[2]] for x in wp["tabs"]]
            if wp["texto"] != t or len(got) != len(exp) or any(abs(g[0] - e[0]) > 0.1 or g[1:] != e[1:] for g, e in zip(got, exp)): malos.append((i, wp, exp))
        anota(f"Word lee las mismas paradas en los {len(PARRAFOS)} párrafos (posición, alineación y relleno)", not malos, malos[:3])
        # cómo lo pinta Word y cómo NOM
        remoto.correr_office("verdad-tabs", "verdad.py", f"archivos\\{NOMBRE}", timeout=600)
        v = REAL / "verdad" / (NOMBRE[:-5] + ".verdad.json")
        remoto.bajar(f"nomoffice/archivos/verdad/{v.name}", v)
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    sal = Path.home() / "NOM/trabajo/tabulaciones-real"
    r = subprocess.run([sys.executable, str(AQUI / "comparar_real.py"), str(REAL / "archivos" / NOMBRE), "--salida", str(sal)], capture_output=True, text=True)
    linea = next((x for x in r.stdout.splitlines() if x.startswith(NOMBRE)), r.stdout[-300:] + r.stderr[-300:])
    rep = json.loads((sal / "resumen.json").read_text(encoding="utf-8"))[0] if (sal / "resumen.json").exists() else {}
    anota(f"NOM pinta los renglones como Word: {linea}", rep.get("pct_iguales", 0) >= 99, linea)
    print(f"\nTabulaciones contra Word 2024: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

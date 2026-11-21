#!/usr/bin/env python3
"""Minigráficos de Excel (Insertar → Minigráficos) en NOM contra Excel 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/minigraficos_real.py

x14-minigraficos.xlsx (hecho en Excel con windows/hacer_minigraficos.py: de línea con marcas, alto y bajo, de columna con
negativos, de pérdidas y ganancias y uno con celdas vacías):
1. NOM lee los mismos grupos que Excel (windows/minigraficos_ver.py: tipo, datos, celdas, color y marcas).
2. Se ven como en Excel: la imagen del rango que da Excel (windows/rango_foto.py) y la de NOM, celda por celda, con lo
   dibujado en el mismo lugar (a 2 px) y con la misma cantidad de color (a un 15 %).
3. Descargado sin tocar, Excel ve los mismos grupos.
4. Hechos en NOM desde cero (línea y columna): Excel los ve; quitados, ya no; Excel abre todo sin avisos.
"""
import base64, json, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
SAL = Path.home() / "NOM/trabajo/minigraficos-real"
TIPO = {"line": 1, "column": 2, "stacked": 3}
res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:600]}"), flush=True)


def excel_ver(xlsx):
    import remoto
    remoto.subir([xlsx], "nomoffice/entrada")
    remoto.correr_office("mg-" + xlsx.stem, "minigraficos_ver.py", f"entrada\\{xlsx.name}", f"{xlsx.stem}.mg.json", timeout=300)
    tmp = SAL / f"{xlsx.stem}.mg.json"; remoto.bajar(f"nomoffice/{xlsx.stem}.mg.json", tmp)
    return json.loads(tmp.read_text(encoding="utf-8"))


def bajar(p, nombre):
    destino = SAL / nombre
    with p.expect_download(timeout=120000) as dl:
        p.evaluate("(n) => NOM.formatos.exportarListo(NOM.editorActual.doc, 'xlsx', n)", destino.stem)
    dl.value.save_as(destino)
    return destino


def medir(path, filas):
    from PIL import Image
    im = Image.open(path).convert("RGB"); px = im.load(); W, H = im.size
    color = lambda q: max(q) - min(q) > 60
    out = []
    for y0, y1, x0 in filas:
        pts = [(x, y) for y in range(y0, y1) for x in range(x0, W) if color(px[x, y])]
        out.append((min(x for x, _ in pts), max(x for x, _ in pts), min(y for _, y in pts) - y0, max(y for _, y in pts) - y0, len(pts)) if pts else None)
    return out


def main():
    import remoto
    SAL.mkdir(parents=True, exist_ok=True)
    arch = REAL / "archivos/x14-minigraficos.xlsx"
    verdad = excel_ver(arch)["hojas"]["Ventas"]
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    try:
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(2000)
        mg = p.evaluate("() => NOM.editorActual.modoApi.api.tabla().mg")
        def comoExcel(g):
            cs = [s[0] for s in g["s"]]; ds = [s[1].split("!")[-1] for s in g["s"]]
            return {"tipo": TIPO[g["tipo"]], "celdas": cs[0] + (":" + cs[-1] if len(cs) > 1 else ""), "n": len(cs), "color": (g.get("c") or {}).get("serie"),
                    "marcas": bool(g.get("marcas")), "alto": bool(g.get("alto")), "bajo": bool(g.get("bajo")), "neg": bool(g.get("neg")), "datos0": ds[0]}
        nom = sorted([comoExcel(g) for g in mg], key=lambda q: q["celdas"])
        exp = [{"tipo": v["tipo"], "celdas": v["celdas"], "n": v["n"], "color": v["color"], "marcas": v["marcas"], "alto": v["alto"], "bajo": v["bajo"], "neg": v["neg"],
                "datos0": v["datos"].split("!")[-1].replace("$", "").split(":")[0] + ":" + v["datos"].replace("$", "").split(":")[-1] if v["n"] == 1 else None} for v in verdad]
        for e, n in zip(exp, nom):
            if e["datos0"] is None: n["datos0"] = None
        anota(f"NOM lee los {len(verdad)} grupos como Excel (tipo, celdas, color y marcas)", nom == exp, f"NOM {nom} · Excel {exp}")
        # cómo se ven: la imagen de Excel y la de NOM del rango A1:H7
        remoto.correr_office("foto-x14", "rango_foto.py", "archivos\\x14-minigraficos.xlsx", "Ventas", "A1:H7", "x14-foto.png", timeout=300)
        remoto.bajar("nomoffice/x14-foto.png", SAL / "excel.png")
        a1, b1 = p.locator('td[data-c="0"][data-f="0"]').bounding_box(), p.locator('td[data-c="7"][data-f="6"]').bounding_box()
        p.screenshot(path=str(SAL / "nom.png"), clip={"x": a1["x"], "y": a1["y"], "width": b1["x"] + b1["width"] - a1["x"], "height": b1["y"] + b1["height"] - a1["y"]})
        from PIL import Image
        hE = Image.open(SAL / "excel.png").size[1]
        filas = [(20 + 32 * k, 51 + 32 * k, 440) for k in range(6)] if abs(hE - 212) <= 2 else []
        me, mn = medir(SAL / "excel.png", filas), medir(SAL / "nom.png", filas)
        malas = [(k, e, n) for k, (e, n) in enumerate(zip(me, mn)) if not e or not n or any(abs(x - y) > 2 for x, y in zip(e[:4], n[:4])) or abs(e[4] - n[4]) > 0.15 * e[4]]
        anota("se ven como en Excel: lo dibujado de cada celda en el mismo lugar (2 px) y con la misma cantidad de color (15 %)", filas and not malas, malas or "sin medidas")
        v2 = excel_ver(bajar(p, "x14-tal-cual.xlsx"))["hojas"].get("Ventas", [])
        anota("descargado sin tocar: Excel ve los mismos grupos", [(g["tipo"], g["celdas"], g["datos"]) for g in v2] == [(g["tipo"], g["celdas"], g["datos"]) for g in verdad], v2)
        # en el x14 abierto: quitar el de H7 (el extLst del original se reescribe sin tocar lo demás)
        p.evaluate("() => NOM.editorActual.modoApi.api.minigraficos.quitar({ c1: 7, c2: 7, f1: 6, f2: 6 })"); p.wait_for_timeout(500)
        v5 = excel_ver(bajar(p, "x14-sin-h7.xlsx"))["hojas"].get("Ventas", [])
        anota("en el x14: quitado el de H7, Excel ve los otros 3 grupos", sorted(g["celdas"] for g in v5) == sorted(g["celdas"] for g in verdad if g["celdas"] != "H7"), v5)
        # hechos en NOM
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'mg-nom', modo: 'HOJA', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.api", timeout=60000); p.wait_for_timeout(800)
        r = p.evaluate("""() => { const H = NOM.editorActual.modoApi.api, t = H.tabla(), cs = {};
          [[3, 5, 2, 8], [1, -2, 4, 6], [9, 7, 5, 3]].forEach((f, i) => f.forEach((v, j) => (cs['c:' + NOM.formulas.ref(j + 1, i + 1)] = String(v))));
          NOM.editorActual.doc.cambiar(t.id, cs);
          return [H.minigraficos.crear('line', 'B2:E3', 'F2:F3'), H.minigraficos.crear('column', 'B4:E4', 'F4')]; }""")
        p.wait_for_timeout(800)
        hechos = p.evaluate("() => document.querySelectorAll('svg.xl-mg').length")
        anota("hechos en NOM: se crean y se dibujan (3 celdas)", hechos == 3 and all(isinstance(x, dict) for x in r), [r, hechos])
        nuevo = bajar(p, "mg-nom.xlsx")
        v3 = list(excel_ver(nuevo)["hojas"].values())[0]
        anota("…y Excel los ve: una línea en F2:F3 y una columna en F4", sorted((g["tipo"], g["celdas"], g["datos"].split("!")[-1].replace("$", "")) for g in v3) == [(1, "F2:F3", "B2:E3"), (2, "F4", "B4:E4")], v3)
        p.evaluate("() => NOM.editorActual.modoApi.api.minigraficos.quitar({ c1: 5, c2: 5, f1: 1, f2: 3 })"); p.wait_for_timeout(500)
        sin = bajar(p, "mg-sin.xlsx")
        v4 = list(excel_ver(sin)["hojas"].values())[0]
        anota("quitados: Excel ya no ve ninguno", v4 == [] and p.evaluate("() => document.querySelectorAll('svg.xl-mg').length") == 0, v4)
        import office_abre
        ro = office_abre.abrir([SAL / "x14-tal-cual.xlsx", SAL / "x14-sin-h7.xlsx", nuevo, sin], "minigraficos")
        for a in ("x14-tal-cual.xlsx", "x14-sin-h7.xlsx", "mg-nom.xlsx", "mg-sin.xlsx"): anota(f"Excel abre {a} sin avisos ni reparar", office_abre.bien(ro.get(a)), json.dumps(ro.get(a), ensure_ascii=False)[:200])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nMinigráficos contra Excel 2024: {sum(1 for x in res if x)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

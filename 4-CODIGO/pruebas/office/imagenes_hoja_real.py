#!/usr/bin/env python3
"""Imágenes puestas en una hoja desde NOM (Más → Imagen…) contra Excel 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/imagenes_hoja_real.py

1. En un libro de NOM: la imagen aparece en la celda elegida, se mueve arrastrándola a otra y lleva su texto alternativo.
2. Descargado, Excel la ve (windows/formas_ver.py) como imagen en esa celda, del mismo tamaño y con su texto alternativo.
3. En un libro de Excel que ya trae dibujos (x06-objetos.xlsx): agregada aquí, Excel ve los de antes y la nueva.
4. Excel abre todo sin avisos.
"""
import base64, io, json, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
SAL = Path.home() / "NOM/trabajo/imagenes-hoja-real"
res = []


def anota(n, ok, det=""):
    res.append(ok); print(("  bien " if ok else "  MAL  ") + n + ("" if ok else f"  · {str(det)[:600]}"), flush=True)


def png(w, h):
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (w, h), (30, 120, 200)); ImageDraw.Draw(im).rectangle([4, 4, w - 5, h - 5], outline=(255, 255, 255), width=3)
    b = io.BytesIO(); im.save(b, "PNG"); return "data:image/png;base64," + base64.b64encode(b.getvalue()).decode()


def excel_ver(xlsx):
    import remoto
    remoto.subir([xlsx], "nomoffice/entrada")
    remoto.correr_office("fv-" + xlsx.stem, "formas_ver.py", f"entrada\\{xlsx.name}", f"{xlsx.stem}.fv.json", timeout=300)
    tmp = SAL / f"{xlsx.stem}.fv.json"; remoto.bajar(f"nomoffice/{xlsx.stem}.fv.json", tmp)
    return json.loads(tmp.read_text(encoding="utf-8"))


def bajar(p, nombre):
    destino = SAL / nombre
    with p.expect_download(timeout=120000) as dl:
        p.evaluate("(n) => NOM.formatos.exportarListo(NOM.editorActual.doc, 'xlsx', n)", destino.stem)
    dl.value.save_as(destino)
    return destino


def main():
    SAL.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    try:
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'img-hoja', modo: 'HOJA', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function("() => NOM.editorActual && NOM.editorActual.modoApi && NOM.editorActual.modoApi.api", timeout=60000); p.wait_for_timeout(800)
        p.evaluate("() => { const t = NOM.editorActual.modoApi.api.tabla(); NOM.editorActual.doc.cambiar(t.id, { 'c:A1': 'Logo', 'c:A2': 'Ventas', 'c:B2': '120' }); }")
        p.locator('td[data-c="2"][data-f="1"]').click(); p.wait_for_timeout(200)
        id_ = p.evaluate("(s) => NOM.editorActual.modoApi.api.imagen({ src: s, w: 120, h: 80 })", png(120, 80)); p.wait_for_timeout(800)
        img, td = p.locator(".nom-img").first, p.locator('td[data-c="2"][data-f="1"]')
        bi, bt = img.bounding_box(), td.bounding_box()
        anota("la imagen aparece en la celda elegida (C2), de 120×80", bi and abs(bi["x"] - bt["x"]) <= 2 and abs(bi["y"] - bt["y"]) <= 2 and abs(bi["width"] - 120) <= 2, [bi, bt])
        destino_td = p.locator('td[data-c="4"][data-f="4"]').bounding_box()
        p.mouse.move(bi["x"] + 10, bi["y"] + 10); p.mouse.down(); p.mouse.move(destino_td["x"] + 12, destino_td["y"] + 12, steps=8); p.mouse.up(); p.wait_for_timeout(700)
        de = p.evaluate("(i) => NOM.editorActual.doc.bloque(i).de", id_)
        anota("arrastrada, queda en E5", de == {"c": 4, "f": 4}, de)
        p.evaluate("(i) => NOM.editorActual.doc.cambiar(i, { alt: 'Logo de la empresa' })", id_); p.wait_for_timeout(300)
        v = excel_ver(bajar(p, "img-nom.xlsx"))
        fs = list(v["hojas"].values())[0]
        anota("Excel la ve: una imagen en E5, de 90×60 pt (120×80 px), con su texto alternativo", len(fs) == 1 and fs[0]["tipo"] == 13 and fs[0]["celda"] == "E5" and abs(fs[0]["ancho"] - 90) <= 1 and abs(fs[0]["alto"] - 60) <= 1 and fs[0]["alt"] == "Logo de la empresa", fs)
        # en un libro de Excel que ya trae dibujos
        x06 = REAL / "archivos/x06-objetos.xlsx"
        antes = excel_ver(x06)
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(x06.read_bytes()).decode(), x06.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)
        hoja = list(antes["hojas"].keys())[0]
        p.locator('td[data-c="0"][data-f="40"]').scroll_into_view_if_needed(); p.locator('td[data-c="0"][data-f="40"]').click(); p.wait_for_timeout(200)
        p.evaluate("(s) => NOM.editorActual.modoApi.api.imagen({ src: s, w: 64, h: 64 })", png(64, 64)); p.wait_for_timeout(800)
        v2 = excel_ver(bajar(p, "x06-con-imagen.xlsx"))
        nv = v2["hojas"].get(hoja, [])
        anota(f"en x06 (con {len(antes['hojas'][hoja])} dibujos): Excel ve los de antes y la nueva en A41", len(nv) == len(antes["hojas"][hoja]) + 1 and any(s["celda"] == "A41" and s["tipo"] == 13 for s in nv), nv[-3:])
        import office_abre
        r = office_abre.abrir([SAL / "img-nom.xlsx", SAL / "x06-con-imagen.xlsx"], "imagenes-hoja")
        for a in ("img-nom.xlsx", "x06-con-imagen.xlsx"): anota(f"Excel abre {a} sin avisos ni reparar", office_abre.bien(r.get(a)), json.dumps(r.get(a), ensure_ascii=False)[:200])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nImágenes en la hoja contra Excel 2024: {sum(res)} de {len(res)} bien")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

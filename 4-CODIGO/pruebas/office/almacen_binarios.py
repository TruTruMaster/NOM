#!/usr/bin/env python3
"""El almacén de binarios de punta a punta (imágenes y el Office original fuera del documento).

  almacen_binarios.py [ARCHIVO.pptx|.docx|.xlsx ...]      (por omisión t03-todo-powerpoint.pptx, 10 MB)

Con el servidor de pruebas en NOM_BASE:
  1. importa como la interfaz (NOM.binarios.crearEnEspacio) y el documento se crea (antes: «superaría 25 MB»);
     el estado del documento queda chico y los binarios van al almacén del espacio;
  2. lo abre: las imágenes se ven (ninguna se queda como referencia) y sin errores de JavaScript;
  3. lo descarga en su formato: sale IDÉNTICO, parte por parte, a lo que sale del mismo archivo sin almacén;
  4. .nom: guardarlo y abrirlo da el mismo documento (con las imágenes dentro);
  5. otra persona del espacio (otro navegador, sin nada en su equipo) lo abre y ve las imágenes (del servidor);
  6. el enlace para ver, sin cuenta, también las ve; y una huella que no es del documento da 404.
Sale con código 1 si algo falla.
"""
import base64, hashlib, io, json, sys, time, zipfile
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import nomprueba as N  # noqa: E402
import fidelidad as F  # noqa: E402

res = []
def anota(n, ok, det=""):
    res.append((n, bool(ok))); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det else ""), flush=True)

JS_IMPORTAR = r"""
async ([b64, nombre, esp]) => {
  const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  const r = await NOM.formatos.importar(new File([u], nombre));
  const pesoSin = JSON.stringify(r.ops).length;
  const espacio = (await NOM.api.get('/api/espacios')).find((e) => e.id === esp) || { id: esp };
  const t0 = performance.now();
  const d = await NOM.binarios.crearEnEspacio(espacio, r.nombre, r.modo, r.ops);
  // lo mismo SIN almacén (para comparar lo que se descarga): el documento en memoria con los datos dentro
  const sin = NOM.formatos.docDeOps(r.nombre, r.modo, r.ops);
  const ext = nombre.split('.').pop().toLowerCase();
  const sale = ext === 'docx' ? await NOM.ooxml.docx(sin) : ext === 'xlsx' ? await NOM.ooxml.xlsx(sin) : await NOM.ooxml.pptx(sin, false);
  let s = ''; for (let i = 0; i < sale.length; i += 0x8000) s += String.fromCharCode.apply(null, sale.subarray(i, i + 0x8000));
  return { id: d.id, ms: Math.round(performance.now() - t0), pesoSin, sinAlmacen: btoa(s), hay: NOM.binarios.hayAlmacen };
}
"""
JS_EXPORTAR = r"""
async (ext) => {
  const doc = NOM.editorActual.doc;
  await NOM.binarios.precargar(doc);
  const sale = ext === 'docx' ? await NOM.ooxml.docx(doc) : ext === 'xlsx' ? await NOM.ooxml.xlsx(doc) : await NOM.ooxml.pptx(doc, false);
  let s = ''; for (let i = 0; i < sale.length; i += 0x8000) s += String.fromCharCode.apply(null, sale.subarray(i, i + 0x8000));
  // .nom de ida y vuelta
  const nom = NOM.formatos.aNomBytes(doc);
  const imp = await NOM.formatos.importar(new File([nom], 'x.nom'));
  const d2 = NOM.formatos.docDeOps(imp.nombre, imp.modo, imp.ops);
  const sale2 = ext === 'docx' ? await NOM.ooxml.docx(d2) : ext === 'xlsx' ? await NOM.ooxml.xlsx(d2) : await NOM.ooxml.pptx(d2, false);
  let s2 = ''; for (let i = 0; i < sale2.length; i += 0x8000) s2 += String.fromCharCode.apply(null, sale2.subarray(i, i + 0x8000));
  return { conAlmacen: btoa(s), viaNom: btoa(s2), nomBytes: nom.length, refs: NOM.binarios.refsDoc(doc).size };
}
"""
JS_IMAGENES = r"""
() => {
  const imgs = [...document.querySelectorAll('img')].filter((i) => i.offsetParent || i.closest('.diapo'));
  const ref = imgs.filter((i) => (i.getAttribute('src') || '').startsWith('nom:bin/') || i.hasAttribute('data-nombin-src')).length;
  const blob = imgs.filter((i) => (i.getAttribute('src') || '').startsWith('blob:')).length;
  const rotas = imgs.filter((i) => i.complete && i.naturalWidth === 0 && !(i.getAttribute('src') || '').startsWith('data:image/gif')).length;
  const estilos = [...document.querySelectorAll('[style*="nom:bin/"]')].length;
  return { total: imgs.length, ref, blob, rotas, estilos };
}
"""

def partes(b):
    z = zipfile.ZipFile(io.BytesIO(b))
    return {n: z.read(n) for n in z.namelist()}

def iguales(a, b):
    pa, pb = partes(a), partes(b)
    dif = sorted(n for n in set(pa) | set(pb) if pa.get(n) != pb.get(n))
    return dif

def esperar_imagenes(p, s=90):
    t0 = time.time(); r = None
    while time.time() - t0 < s:
        r = p.evaluate(JS_IMAGENES)
        if r["ref"] == 0 and r["estilos"] == 0 and r["total"] > 0: return True, r
        p.wait_for_timeout(1000)
    return False, r

def otra_pagina(A, c=None):
    """Otra persona en el mismo navegador (otro contexto: nada en común con la primera); c=None, sin cuenta."""
    ctx = A.b.new_context(viewport={"width": 1300, "height": 900})
    p = ctx.new_page()
    if c:
        p.goto(N.B + "/"); p.wait_for_selector("#u", timeout=20000)
        p.fill("#u", c["nombre"]); p.fill("#c", N.CL); p.click("button[type=submit]")
        p.wait_for_function("() => !!localStorage.getItem('nom.token')", timeout=20000); p.wait_for_timeout(800)
    return ctx, p

def main():
    archivos = [Path(x) for x in sys.argv[1:]] or [F.CACHE / "office-real/archivos/t03-todo-powerpoint.pptx"]
    N.esperar_servidor()
    for arch in archivos:
        print(f"\n== {arch.name} ({arch.stat().st_size // 1024} KB)", flush=True)
        ext = arch.suffix[1:].lower()
        a, b = N.cuenta(), N.cuenta()
        esp = N.espacio(a["token"], "Almacén")
        N.api("PUT", f"/api/espacios/{esp}/accesos", {"usuario": b["nombre"], "permiso": "LECTURA"}, a["token"])
        A = N.Navegador(movil=False); A.entrar(a)
        try:
            A.p.goto(N.B + "/#/espacio/" + esp); A.p.wait_for_timeout(1500)
            r = A.p.evaluate(JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
            anota("1. se crea el documento (antes no cabía)", r.get("id"), f"{r['ms']} ms · sin almacén pesaba {r['pesoSin'] / 1e6:.1f} MB · almacén: {r.get('hay')}")
            st = N.api("GET", "/api/cuentas/almacenamiento", None, a["token"])
            st = st[1] if isinstance(st, tuple) else st
            e = [x for x in (st or {}).get("espacios", []) if x["id"] == esp]
            if e: anota("   el documento queda chico y los binarios van aparte", e[0]["bytes"] - e[0].get("binarios", 0) < 6e6 and e[0].get("binarios", 0) > 1e6,
                        f"documento y versiones {(e[0]['bytes'] - e[0].get('binarios', 0)) / 1e6:.1f} MB · binarios {e[0].get('binarios', 0) / 1e6:.1f} MB")
            A.p.goto(N.B + "/#/doc/" + r["id"]); A.p.wait_for_timeout(4000)
            ok, im = esperar_imagenes(A.p)
            anota("2. al abrirlo las imágenes se ven", ok and im["rotas"] == 0, json.dumps(im))
            x = A.p.evaluate(JS_EXPORTAR, ext)
            sal = F.CACHE / "almacen-binarios"; sal.mkdir(parents=True, exist_ok=True)
            (sal / f"{arch.stem}.con-almacen.{ext}").write_bytes(base64.b64decode(x["conAlmacen"]))
            (sal / f"{arch.stem}.sin-almacen.{ext}").write_bytes(base64.b64decode(r["sinAlmacen"]))
            dif = iguales(base64.b64decode(x["conAlmacen"]), base64.b64decode(r["sinAlmacen"]))
            anota(f"3. la descarga sale idéntica a la de sin almacén ({x['refs']} referencias)", not dif, f"partes distintas: {dif[:6]}")
            dif2 = iguales(base64.b64decode(x["viaNom"]), base64.b64decode(r["sinAlmacen"]))
            anota("4. guardado como .nom y abierto, sale igual", not dif2, f".nom de {x['nomBytes'] / 1e6:.1f} MB · distintas: {dif2[:6]}")
            # 5. otra persona (otro navegador, sin nada en su equipo)
            cb, pb = otra_pagina(A, b)
            try:
                pb.goto(N.B + "/#/doc/" + r["id"]); pb.wait_for_timeout(4000)
                ok, im = esperar_imagenes(pb)
                anota("5. otra persona del espacio las ve (del servidor)", ok and im["rotas"] == 0, json.dumps(im))
            finally: cb.close()
            # 6. enlace para ver, sin cuenta
            en = N.api("POST", f"/api/documentos/{r['id']}/enlace-ver", {}, a["token"])
            en = (en[1] if isinstance(en, tuple) else en) or {}
            tok = en.get("token") or en.get("enlace") or ""
            tok = tok.rsplit("/", 1)[-1]
            if tok:
                cv, pv = otra_pagina(A, None)
                try:
                    pv.goto(N.B + "/#/ver/" + tok); pv.wait_for_timeout(5000)
                    ok, im = esperar_imagenes(pv)
                    anota("6. el enlace para ver (sin cuenta) las ve", ok and im["rotas"] == 0, json.dumps(im))
                finally: cv.close()
                h = hashlib.sha256(b"no es de este documento").hexdigest()
                import urllib.request
                try: urllib.request.urlopen(N.B + f"/api/ver/{tok}/binarios/{h}"); cod = 200
                except urllib.error.HTTPError as ex: cod = ex.code
                anota("   …y una huella que no es del documento no se entrega", cod == 404, str(cod))
            else: anota("6. enlace para ver", False, f"no se pudo crear: {en}")
            anota("sin errores de JavaScript", not A.errores, str(A.errores[:3]))
        finally:
            A.cerrar()
    malos = [n for n, ok in res if not ok]
    print(f"\nAlmacén de binarios: {len(res) - len(malos)} de {len(res)} bien")
    return 1 if malos else 0

if __name__ == "__main__":
    sys.exit(main())

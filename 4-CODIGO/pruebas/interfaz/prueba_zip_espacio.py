"""«Descargar todo (Word, Excel y PowerPoint)» de un espacio: cada documento que vino de Office sale de su original.

  python pruebas/interfaz/prueba_zip_espacio.py SALIDA

Importa dos Word del corpus (uno con símbolos w:sym, otro con imágenes), recarga la página (lo del almacén de
binarios ya no está en la memoria), descarga el .zip del espacio desde su menú «Más» y compara cada Word con su
original: el texto de document.xml igual (con sus w:sym) y las mismas imágenes, byte por byte.
"""
import base64, hashlib, io, os, re, sys, zipfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
import nomprueba as N  # noqa: E402

CORPUS = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache" / "nom-pruebas")) / "corpus" / "docx"
ARCHIVOS = ["ms-informe-gobernanza.docx", "ms-hoja-informativa.docx"]
SALIDA = Path(sys.argv[1] if len(sys.argv) > 1 else Path.home() / ".cache" / "nom-pruebas" / "salida")
res = []


def check(nombre, ok, info=""):
    res.append(ok); print(f"  {'ok ' if ok else 'MAL'}  {nombre}" + (f"   [{str(info)[:300]}]" if info and not ok else ""), flush=True)


def partes(datos):
    z = zipfile.ZipFile(io.BytesIO(datos))
    xml = z.read("word/document.xml").decode("utf-8")
    texto = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", xml))
    simbolos = re.findall(r"<w:sym\b[^>]*/>", xml)
    medios = {n.split("/")[-1]: hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist() if n.startswith("word/media/")}
    return texto, simbolos, medios


def main():
    SALIDA.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    c = N.cuenta(); esp = N.espacio(c["token"], "Descargar todo")
    P = N.Navegador(movil=False); P.entrar(c)
    p = P.p
    try:
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(1200)
        for nombre in ARCHIVOS:
            b64 = base64.b64encode((CORPUS / nombre).read_bytes()).decode()
            r = p.evaluate("""async ([b64, esp, nombre]) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
              const r = await NOM.formatos.importar(new File([u], nombre)); const espacio = (await NOM.api.get('/api/espacios')).find((e) => e.id === esp);
              const d = await NOM.binarios.crearEnEspacio(espacio, r.nombre, r.modo, r.ops); return d.id; }""", [b64, esp, nombre])
            check(f"se importa {nombre}", bool(r), r)
        # una página nueva: lo del almacén (el original y las imágenes) ya no está en la memoria
        p.goto("about:blank"); p.goto(N.B + "/#/espacio/" + esp); p.wait_for_selector("[data-a=mas-espacio]", timeout=20000); p.wait_for_timeout(800)
        with p.expect_download(timeout=120000) as dl:
            p.click("[data-a=mas-espacio]"); p.wait_for_timeout(300)
            p.evaluate("""() => { const b = [...document.querySelectorAll('.menu:not(.saliendo) [role=menuitem]')].find((x) => x.textContent.includes('Descargar todo')); b.scrollIntoView({ block: 'nearest' }); b.click(); }""")
        destino = SALIDA / "espacio.zip"; dl.value.save_as(str(destino))
        z = zipfile.ZipFile(destino)
        nombres = z.namelist()
        check("el .zip trae los dos Word", sum(1 for n in nombres if n.endswith(".docx")) == 2, nombres)
        for nombre in ARCHIVOS:
            orig = (CORPUS / nombre).read_bytes()
            sal = next((z.read(n) for n in nombres if n.endswith(".docx") and n[:-5].lower().replace(" ", "-") in nombre.lower()), None)
            if sal is None:
                # el nombre en el .zip es el título que le puso la importación: se busca por su texto
                to = partes(orig)[0]
                sal = next((z.read(n) for n in nombres if n.endswith(".docx") and partes(z.read(n))[0][:200] == to[:200]), None)
            check(f"{nombre}: está en el .zip", sal is not None, nombres)
            if sal is None: continue
            (to, so, mo), (ts, ss, ms) = partes(orig), partes(sal)
            check(f"{nombre}: el mismo texto ({len(to)} letras)", to == ts, f"{len(ts)} de {len(to)}")
            check(f"{nombre}: los mismos símbolos de Word (w:sym: {len(so)})", so == ss, f"{len(ss)} de {len(so)}")
            check(f"{nombre}: las mismas imágenes, byte por byte ({len(mo)})", sorted(mo.values()) == sorted(ms.values()), f"{len(ms)} de {len(mo)}")
        # Un documento de este equipo con una imagen del almacén local: «Descargar copia (.nom)» desde la lista de
        # documentos, otra vez sin que la imagen esté en la memoria.
        png = next(zipfile.ZipFile(CORPUS / ARCHIVOS[1]).read(n) for n in zipfile.ZipFile(CORPUS / ARCHIVOS[1]).namelist() if n.startswith("word/media/"))
        idl = p.evaluate("""async (b64) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
          const ref = await NOM.binarios.guardar(u, 'image/jpeg', { espacio: null });
          const id = NOM.local.crear('Local con imagen', 'LIENZO', [{ id: 'local:i1', bloque: 'img1', campos: { tipo: 'imagen', src: ref, pagina: 'p1', x: 40, y: 40, w: 200, h: 150 }, t: 2, n: 'local' }]).id;
          await NOM.boveda.vaciar();   // (con cuenta, la lista de este equipo vive en la bóveda, que guarda en segundo plano)
          return id; }""",
                         base64.b64encode(png).decode())
        p.goto("about:blank"); p.goto(N.B + "/#/"); p.wait_for_selector(f'[data-menu-local="{idl}"]', timeout=20000); p.wait_for_timeout(600)
        with p.expect_download(timeout=60000) as dl:
            p.click(f'[data-menu-local="{idl}"]'); p.wait_for_timeout(300)
            p.evaluate("""() => { const b = [...document.querySelectorAll('.menu:not(.saliendo) [role=menuitem]')].find((x) => x.textContent.includes('Descargar copia')); b.click(); }""")
        nom = Path(dl.value.path()).read_bytes()
        # se abre con el mismo importador de NOM: ¿trae los datos de la imagen o solo su referencia al almacén?
        v = p.evaluate("""async ([b64, img]) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
          const r = await NOM.formatos.importar(new File([u], 'copia.nom')); const t = JSON.stringify(r.ops);
          return { datos: t.includes(img.slice(0, 400)), soloRef: t.includes('nom:bin/') && !t.includes(img.slice(0, 400)) }; }""",
                       [base64.b64encode(nom).decode(), base64.b64encode(png).decode()])
        check(f"documento de este equipo: la copia .nom lleva los datos de su imagen ({len(png)} bytes)", isinstance(v, dict) and v.get("datos"), v)
        check("sin errores de JavaScript", not P.errores, P.errores[:3])
    finally:
        P.cerrar()
    print(f"\nDescargar todo: {sum(res)} bien, {len(res) - sum(res)} mal")
    return 0 if all(res) else 1


if __name__ == "__main__":
    sys.exit(main())

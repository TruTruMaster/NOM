#!/usr/bin/env python3
"""Imágenes EMF y WMF: lo que dibuja NOM (js/emf.js) contra lo que dibuja Windows (GDI, el motor de Office).

  metarchivos_real.py [ARCHIVO.docx|.xlsx|.pptx ...]     (por omisión, todos los del corpus que traen EMF o WMF)

1. Saca cada EMF/WMF distinto de los archivos (por su huella).
2. NOM: los convierte a SVG con js/emf.js y los dibuja en Chromium al doble de su tamaño (96 ppp × 2).
3. Word: coloca cada uno en su propia página del tamaño exacto y la exporta a PDF (pruebas/office/windows/
   metarchivos_word.py, en la sesión de escritorio); cada página se pasa a PNG al mismo tamaño.
4. Compara: «parecido» (100 − diferencia media por píxel) y «tinta» (lo mismo pero solo donde alguna de las dos
   imágenes tiene algo dibujado: no se infla con el blanco). Deja el detalle y las imágenes lado a lado en
   ~/.cache/nom-pruebas/metarchivos/.
"""
import base64, hashlib, io, json, os, shutil, subprocess, sys, zipfile
from pathlib import Path
import numpy as np
from PIL import Image, ImageFilter
from playwright.sync_api import sync_playwright

AQUI = Path(__file__).resolve().parent
JS = AQUI.parent.parent / "app/src/main/resources/static/js/emf.js"
CACHE = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas"))
SAL = CACHE / "metarchivos"
WIN = os.environ.get("NOM_WINDOWS", "user@100.65.29.5")
ESCALA = 2


def juntar(archivos):
    ent = SAL / "entrada"; ent.mkdir(parents=True, exist_ok=True)
    vistos = {}
    for a in archivos:
        try: z = zipfile.ZipFile(a)
        except Exception: continue
        for n in z.namelist():
            ext = n.lower().rsplit(".", 1)[-1]
            if ext not in ("emf", "wmf"): continue
            b = z.read(n); h = hashlib.sha256(b).hexdigest()[:16]
            if h in vistos: continue
            nombre = f"{h}.{ext}"
            (ent / nombre).write_bytes(b)
            vistos[h] = {"nombre": nombre, "de": f"{Path(a).name}:{n}"}
    return list(vistos.values())


JS_DIBUJAR = r"""
async ([b64]) => {
  const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  const r = NOM.emf.aSvg(u); if (!r) return null;
  const W = Math.max(2, Math.min(2400, Math.round(r.w * %d))), H = Math.max(2, Math.min(2400, Math.round(r.h * %d)));
  const img = new Image(); img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(r.svg)));
  await img.decode();
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
  const g = cv.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, W, H); g.drawImage(img, 0, 0, W, H);
  return { W, H, png: cv.toDataURL('image/png') };
}
""" % (ESCALA, ESCALA)


def medir(a, b):
    A = np.asarray(Image.open(a).convert("L"), dtype=float)
    B = np.asarray(Image.open(b).convert("L").resize((A.shape[1], A.shape[0])), dtype=float)
    d = np.abs(A - B)
    parecido = 100 - d.mean() / 2.55
    # «tinta»: donde alguna de las dos tiene algo (un poco ensanchado, para no castigar el suavizado de bordes)
    ma = Image.fromarray(((A < 245) | (B < 245)).astype(np.uint8) * 255).filter(ImageFilter.MaxFilter(3))
    m = np.asarray(ma) > 0
    tinta = 100 - (d[m].mean() / 2.55 if m.any() else 0)
    return round(parecido, 2), round(tinta, 2)


def main():
    args = sys.argv[1:]
    if not args:
        args = []
        for sub in ("corpus", "corpus-web", "corpus-reales", "office-real/archivos"):
            for p in (CACHE / sub).rglob("*"):
                if p.suffix.lower() in (".docx", ".xlsx", ".pptx"): args.append(str(p))
    lista = juntar(args)
    print(f"{len(lista)} metarchivos distintos", flush=True)
    nom = SAL / "nom"; nom.mkdir(parents=True, exist_ok=True)
    tam = {}
    with sync_playwright() as pw:
        b = pw.chromium.launch(); pg = b.new_page()
        pg.set_content("<html><body></body></html>"); pg.add_script_tag(path=str(JS))
        for x in lista:
            r = pg.evaluate(JS_DIBUJAR, [base64.b64encode((SAL / "entrada" / x["nombre"]).read_bytes()).decode()])
            if not r: x["nom"] = None; continue
            (nom / (x["nombre"] + ".png")).write_bytes(base64.b64decode(r["png"].split(",", 1)[1]))
            tam[x["nombre"]] = (r["W"], r["H"])
        b.close()
    # Windows: Word coloca cada uno en su página y exporta a PDF (lo que de verdad se ve en Office)
    (SAL / "lista.txt").write_text("".join(f"{n}\t{w}\t{h}\n" for n, (w, h) in tam.items()), encoding="utf-8")
    import tarfile
    with tarfile.open(SAL / "meta.tar", "w") as t:
        t.add(SAL / "entrada", arcname="entrada"); t.add(SAL / "lista.txt", arcname="lista.txt")
    # (por SSH desde Linux, o aquí mismo si la prueba ya corre en la computadora con Windows: remoto.py)
    sys.path.insert(0, str(AQUI))
    import remoto
    remoto.ejecutar("if exist nomoffice\\meta rmdir /s /q nomoffice\\meta & mkdir nomoffice\\meta")
    remoto.subir([SAL / "meta.tar"], "nomoffice/meta")
    remoto.ejecutar("cd nomoffice\\meta && tar -xf meta.tar && del meta.tar")
    r = remoto.correr_office("meta", "metarchivos_word.py", "meta\\entrada", "meta\\lista.txt", "meta\\word.pdf", timeout=3600)
    print("\n".join(r.stdout.strip().splitlines()[-4:]) if r.stdout.strip() else r.stderr[-500:], flush=True)
    if not remoto.bajar("nomoffice/meta/word.pdf", SAL / "word.pdf"): print("Word no dejó el PDF"); return 1
    win = SAL / "windows"; shutil.rmtree(win, ignore_errors=True); win.mkdir(parents=True, exist_ok=True)
    # cada página del PDF de Word a 192 ppp (pypdfium2: igual en Linux y en Windows)
    import pypdfium2
    pdf = pypdfium2.PdfDocument(str(SAL / "word.pdf"))
    for (n, (w, h)), pag in zip(tam.items(), pdf):
        pag.render(scale=192 / 72).to_pil().convert("RGB").crop((0, 0, w, h)).save(win / (n + ".png"))
    pdf.close()
    res = []
    for x in lista:
        a, bw = nom / (x["nombre"] + ".png"), win / (x["nombre"] + ".png")
        if not a.exists() or not bw.exists(): res.append(dict(x, parecido=None, tinta=None)); continue
        p, t = medir(a, bw)
        res.append(dict(x, parecido=p, tinta=t))
    res.sort(key=lambda r: (r["tinta"] is not None, r["tinta"] or 0))
    (SAL / "resultado.json").write_text(json.dumps(res, indent=1, ensure_ascii=False))
    medidos = [r for r in res if r["tinta"] is not None]
    if medidos:
        print(f"medidos {len(medidos)} de {len(res)} · parecido medio {np.mean([r['parecido'] for r in medidos]):.2f} % · tinta media {np.mean([r['tinta'] for r in medidos]):.2f} %")
        print("los más distintos:")
        for r in medidos[:12]: print(f"  {r['tinta']:6.2f} %  {r['parecido']:6.2f} %  {r['nombre']}  ({r['de']})")
    sin = [r for r in res if r["tinta"] is None]
    if sin: print(f"sin medir {len(sin)}: " + ", ".join(r["nombre"] for r in sin[:10]))
    # (lo medido el 18-nov: parecido 96.2 %, tinta 88.7 %; si baja de esto, algo se descompuso en js/emf.js)
    if not medidos: return 1
    pm, tm = np.mean([r["parecido"] for r in medidos]), np.mean([r["tinta"] for r in medidos])
    bien = pm >= 95 and tm >= 87
    print(("BIEN" if bien else "MAL") + f" EMF/WMF contra Word: parecido {pm:.2f} % (mínimo 95) · tinta {tm:.2f} % (mínimo 87)")
    return 0 if bien else 1


if __name__ == "__main__":
    sys.exit(main())

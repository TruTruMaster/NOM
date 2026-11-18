#!/usr/bin/env python3
"""Campos de Word: «Actualizar campos» de NOM contra Word 2024 de verdad.

  campos_real.py [--sin-windows]

1. Word (en la computadora con Windows) arma un documento con 55 campos de todo tipo —fechas con formato, SEQ, REF,
   PAGEREF, IF, fórmulas con sus formatos, propiedades, páginas— y los actualiza (windows/campos_word.py).
2. NOM lo importa por la interfaz: cada campo tiene que llegar con su código y su último valor.
3. NOM actualiza los campos (F9) y cada valor se compara con el de Word. La hora (TIME) no entra (cambia).
4. Se descarga: Word abre el archivo sin avisos y sus campos siguen ahí.
Sale con código 1 si algo no coincide.
"""
import base64, json, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import nomprueba as N  # noqa: E402
import remoto, office_abre  # noqa: E402

CACHE = N.CACHE if hasattr(N, "CACHE") else Path.home() / ".cache/nom-pruebas"
SAL = Path.home() / ".cache/nom-pruebas/campos"; SAL.mkdir(parents=True, exist_ok=True)
res = []
def anota(n, ok, det=""):
    res.append((n, bool(ok))); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det else ""), flush=True)


def main():
    if "--sin-windows" not in sys.argv:
        r = remoto.correr_office("campos", "campos_word.py", timeout=900)
        if "[fin 0]" not in r.stdout: sys.exit("Word no terminó: " + r.stdout[-800:])
        remoto.bajar("nomoffice/archivos/campos-word.docx", SAL / "campos-word.docx")
        remoto.bajar("nomoffice/archivos/campos-word.json", SAL / "campos-word.json")
    verdad = json.loads((SAL / "campos-word.json").read_text(encoding="utf-8"))
    N.esperar_servidor()
    a = N.cuenta(); esp = N.espacio(a["token"], "Campos")
    A = N.Navegador(movil=False); A.entrar(a)
    try:
        A.p.goto(N.B + "/#/espacio/" + esp); A.p.wait_for_timeout(1500)
        r = A.p.evaluate("""async ([b64, esp]) => { const bin = atob(b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
          const r = await NOM.formatos.importar(new File([u], 'campos-word.docx')); const espacio = (await NOM.api.get('/api/espacios')).find((e) => e.id === esp);
          const d = await NOM.binarios.crearEnEspacio(espacio, r.nombre, r.modo, r.ops); return { id: d.id }; }""", [base64.b64encode((SAL / "campos-word.docx").read_bytes()).decode(), esp])
        A.p.goto(N.B + "/#/doc/" + r["id"]); A.p.wait_for_timeout(5000)
        LEER = "() => NOM.textoRico.recorrido(NOM.editorActual.doc).flatMap((p) => p.items.filter((it) => it.e && it.e.tipo === 'campo').map((it) => ({ instr: it.bl.instr, res: it.bl.res })))"
        antes = A.p.evaluate(LEER)
        anota("los campos llegan con su código y su último valor", len(antes) == len(verdad) and all(x["instr"] == v["codigo"] for x, v in zip(antes, verdad)), f"{len(antes)} de {len(verdad)}")
        A.p.click(".wd-cuerpo"); A.p.keyboard.press("End"); A.p.keyboard.press("F9"); A.p.wait_for_timeout(2500)
        despues = A.p.evaluate(LEER)
        bien = mal = 0
        for x, v in zip(despues, verdad):
            if v["codigo"].startswith("TIME"): continue
            if x["res"] == v["resultado"]: bien += 1
            else: mal += 1; print(f"   distinto  {v['codigo']!r}: Word {v['resultado']!r} · NOM {x['res']!r}")
        anota("«Actualizar campos» da lo mismo que Word", mal == 0, f"{bien} de {bien + mal}")
        # la descarga: Word la abre y los campos siguen
        sale = A.p.evaluate("""async () => { const u = await NOM.ooxml.docx(NOM.editorActual.doc); let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }""")
        (SAL / "campos-nom.docx").write_bytes(base64.b64decode(sale))
        o = office_abre.abrir([SAL / "campos-nom.docx"], "campos").get("campos-nom.docx")
        anota("Word abre lo descargado sin avisos ni reparar", office_abre.bien(o), json.dumps(o, ensure_ascii=False)[:200] if o else "")
        anota("sin errores de JavaScript", not A.errores, str(A.errores[:2]))
    finally:
        A.cerrar()
    malos = [n for n, ok in res if not ok]
    print(f"\nCampos de Word: {len(res) - len(malos)} de {len(res)} bien")
    return 1 if malos else 0


if __name__ == "__main__":
    sys.exit(main())

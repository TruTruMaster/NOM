#!/usr/bin/env python3
"""Control de cambios de Word (Revisar → Control de cambios) en NOM contra Word 2024 real.

  NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/cambios_real.py

w18-cambios.docx (hecho en Word con windows/hacer_cambios.py: cambios marcados de dos personas) y lo que ve Word de cada
archivo (windows/cambios_ver.py: la lista de revisiones y el texto con todo aceptado y con todo rechazado):
1. NOM abre w18 con cada cambio de texto como sugerencia (agregar o borrar, con su autor) y en modo sugerencia (el
   documento trae «Control de cambios»). Los de formato y los de dentro de una tabla se ven aplicados (NOM no los modela:
   las celdas de una tabla de Word son texto simple) y se conservan sin tocar al descargar; por eso «aceptar todo» y
   «rechazar todo» se comparan fuera de la tabla.
2. Descargado sin tocar: Word ve los mismos cambios.
3. Todo aceptado en NOM y descargado: Word no ve ningún cambio y el texto es el de Word con todo aceptado; todo rechazado:
   el de Word con todo rechazado.
4. Uno aceptado, uno rechazado y uno nuevo escrito en NOM: Word ve los que quedan y el nuevo, con el nombre de quien lo
   escribió en NOM, y su «aceptar todo» da lo mismo que en NOM.
5. Un documento hecho en NOM con sugerencias: Word las ve como cambios marcados.
6. Word abre todo sin reparar.
"""
import base64, json, os, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
SAL = Path.home() / "NOM/trabajo/cambios-real"
res = []


def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {str(det)[:500]}" if det and not ok else ""), flush=True)


def word_ver(docx):
    import remoto
    remoto.subir([docx], "nomoffice/entrada")
    r = remoto.correr_office("cc-" + docx.stem, "cambios_ver.py", f"entrada\\{docx.name}", f"{docx.stem}.cc.json", timeout=600)
    if not remoto.bajar(f"nomoffice/{docx.stem}.cc.json", SAL / f"{docx.stem}.cc.json"):
        raise SystemExit("Word no dejó resultado: " + (r.stdout or "")[-300:])
    return json.loads((SAL / f"{docx.stem}.cc.json").read_text(encoding="utf-8"))


TIPO = {1: "insertar", 2: "borrar"}


def de_texto(revs):
    """Los cambios de texto que NOM modela (agregar y borrar, fuera de las tablas), como [tipo, autor, texto]."""
    return [[TIPO[r["tipo"]], r["autor"], (r["texto"] or "").replace("\r", "")] for r in revs if r["tipo"] in TIPO and not r.get("tabla")]


CELDAS_W18 = {"uno", "dos", "dos!", "tres", "cuatro"}


def sin_vacios(ps, sin_tabla=False):
    return [p for p in ps if p.strip() and not (sin_tabla and p in CELDAS_W18)]


def lista_nom(p):
    return p.evaluate("() => NOM.editorActual.modoApi.sugerencias.lista().map((g) => [g.tipo, g.autor.n, g.texto])")


def bajar(p, nombre):
    destino = SAL / nombre
    with p.expect_download(timeout=180000) as dl:
        p.evaluate("(n) => NOM.formatos.exportarListo(NOM.editorActual.doc, 'docx', n)", destino.stem)
    dl.value.save_as(destino)
    return destino


def abrir(p, esp, arch):
    p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
    p.evaluate(C.JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
    p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1500)


def resolver_todas(p, aceptar):
    p.evaluate("(ok) => { const S = NOM.editorActual.modoApi.sugerencias; S.lista().slice().reverse().forEach((g) => S.resolver(g, ok)); return 1; }", aceptar)
    p.wait_for_timeout(800)


def main():
    arch = REAL / "archivos/w18-cambios.docx"
    if not arch.exists(): arch = Path.home() / "NOM/trabajo/21nov/cambios/w18-cambios.docx"
    SAL.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c); p = nav.p
    yo = p.evaluate("() => (NOM.api.usuario() || {}).nombre")
    try:
        verdad = word_ver(arch)
        esperado = de_texto(verdad["revisiones"])
        # 1. lo que ve NOM
        abrir(p, esp, arch)
        anota(f"NOM ve los {len(esperado)} cambios de texto de Word como sugerencias (tipo, autor y texto)", lista_nom(p) == esperado, f"NOM {lista_nom(p)} · Word {esperado}")
        anota("…y empieza en modo sugerencia (el documento trae «Control de cambios»)", p.evaluate("() => NOM.editorActual.modoApi.sugerencias.activo() && NOM.editorActual.doc.valor('_doc', 'controlCambios') === true"))
        # (Mostrar para revisión, como Word: un Word abre con «marcas simples»: lo borrado no se ve)
        vis = lambda: p.evaluate("() => ({ m: NOM.editorActual.modoApi.sugerencias.marcas(), b: [...document.querySelectorAll('.wd-sugb')].some((e) => e.getClientRects().length), i: [...document.querySelectorAll('.wd-sug')].some((e) => e.getClientRects().length) })")
        x0 = vis()
        p.evaluate("() => NOM.editorActual.modoApi.sugerencias.fijarMarcas('todas')"); p.wait_for_timeout(500); x1 = vis()
        p.evaluate("() => NOM.editorActual.modoApi.sugerencias.fijarMarcas('original')"); p.wait_for_timeout(500); x2 = vis()
        p.evaluate("() => NOM.editorActual.modoApi.sugerencias.fijarMarcas('simples')"); p.wait_for_timeout(500)
        anota("…con «marcas simples» como Word (lo borrado no se ve); «todas» lo muestra y «original» oculta lo agregado",
              x0 == {"m": "simples", "b": False, "i": True} and x1["b"] and x1["i"] and x2["b"] and not x2["i"], [x0, x1, x2])
        # 2. descargado sin tocar
        v = word_ver(bajar(p, "w18-tal-cual.docx"))
        anota("descargado sin tocar: Word ve los mismos cambios (también el de formato y el de la tabla)", [[r["tipo"], r["autor"], r["texto"]] for r in v["revisiones"]] == [[r["tipo"], r["autor"], r["texto"]] for r in verdad["revisiones"]] and v["control"], v["revisiones"])
        # 3. todo aceptado / todo rechazado
        resolver_todas(p, True)
        v = word_ver(bajar(p, "w18-aceptado.docx"))
        anota("todo aceptado en NOM: Word no ve cambios de texto", not de_texto(v["revisiones"]), v["revisiones"])
        anota("…y el texto es el de Word con todo aceptado", sin_vacios(v["tal_cual"]) == sin_vacios(verdad["aceptado"]), f"NOM→Word {sin_vacios(v['tal_cual'])} · Word {sin_vacios(verdad['aceptado'])}")
        abrir(p, esp, arch)
        resolver_todas(p, False)
        v = word_ver(bajar(p, "w18-rechazado.docx"))
        anota("todo rechazado en NOM: el texto es el de Word con todo rechazado (sin contar párrafos vacíos ni la tabla)", not de_texto(v["revisiones"]) and sin_vacios(v["tal_cual"], True) == sin_vacios(verdad["rechazado"], True), f"NOM→Word {sin_vacios(v['tal_cual'])} · Word {sin_vacios(verdad['rechazado'])}")
        # 4. uno aceptado, uno rechazado y uno nuevo
        abrir(p, esp, arch)
        p.evaluate("() => { const S = NOM.editorActual.modoApi.sugerencias, L = S.lista(); S.resolver(L[1], false); S.resolver(L[0], true); return 1; }"); p.wait_for_timeout(600)
        p.get_by_text("Fin del texto.", exact=True).first.click(); p.wait_for_timeout(200)
        p.keyboard.press("End"); p.keyboard.type(" Nuevo."); p.wait_for_timeout(900)
        quedan = lista_nom(p)
        anota("…en NOM: lo escrito queda como sugerencia de quien escribe", quedan and quedan[-1] == ["insertar", yo, " Nuevo."], quedan)
        v = word_ver(bajar(p, "w18-mezcla.docx"))
        anota("uno aceptado, uno rechazado y uno nuevo: Word ve los que quedan y el nuevo", de_texto(v["revisiones"]) == quedan, f"Word {de_texto(v['revisiones'])} · NOM {quedan}")
        resolver_todas(p, True)
        txt_nom = p.evaluate("() => [...document.querySelectorAll('.wd [data-p], .wd p')].map((e) => e.innerText.replace(/\\u00a0/g, ' ').trim()).filter(Boolean)")
        anota("…y su «aceptar todo» da lo mismo que en NOM", all(any(t.startswith(x) or x.startswith(t) for t in txt_nom) for x in sin_vacios(v["aceptado"]) if x not in ("uno", "dos!", "tres", "cuatro")), f"Word {sin_vacios(v['aceptado'])} · NOM {txt_nom}")
        # 5. hecho en NOM
        p.evaluate("""async (esp) => { const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'Sugerencias de NOM', modo: 'TEXTO', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function(U.JS_LISTO, timeout=60000); p.wait_for_timeout(1200)
        p.locator(".wd [contenteditable=true]").first.click(); p.wait_for_timeout(200)
        p.keyboard.type("Hola mundo."); p.wait_for_timeout(400)
        p.evaluate("() => NOM.editorActual.modoApi.sugerencias.alternar(true)"); p.wait_for_timeout(300)
        p.keyboard.press("End"); p.keyboard.type(" Adiós."); p.wait_for_timeout(300)
        p.keyboard.press("Home")
        for _ in range(5): p.keyboard.press("Shift+ArrowRight")
        p.keyboard.press("Delete"); p.wait_for_timeout(900)
        hechas = lista_nom(p)
        v = word_ver(bajar(p, "desde-nom.docx"))
        anota("hecho en NOM: Word ve sus sugerencias como cambios marcados", de_texto(v["revisiones"]) == hechas and len(hechas) == 2, f"Word {de_texto(v['revisiones'])} · NOM {hechas}")
        anota("…con todo aceptado, «mundo. Adiós.»; rechazado, «Hola mundo.»", sin_vacios(v["aceptado"]) == ["mundo. Adiós."] and sin_vacios(v["rechazado"]) == ["Hola mundo."], f"{v['aceptado']} · {v['rechazado']}")
        # 6. Word abre todo sin reparar
        import office_abre
        lista = [SAL / n for n in ("w18-tal-cual.docx", "w18-aceptado.docx", "w18-rechazado.docx", "w18-mezcla.docx", "desde-nom.docx")]
        r = office_abre.abrir(lista, "cambios")
        for a in lista:
            anota(f"Word abre {a.name} sin avisos ni reparar", office_abre.bien(r.get(a.name)), json.dumps(r.get(a.name), ensure_ascii=False)[:200])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nControl de cambios contra Word 2024: {sum(1 for x in res if x[1])} de {len(res)} bien")
    return 0 if all(x[1] for x in res) else 1


if __name__ == "__main__":
    sys.exit(main())

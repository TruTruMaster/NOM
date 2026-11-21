#!/usr/bin/env python3
"""La pestaña Presentación de PowerPoint (y las secciones de Inicio) contra PowerPoint 2024 real.

  python pruebas/office/presentacion_real.py [p13-presentacion.pptx] [--salida DIR]

1. p13 (hecho en PowerPoint con windows/hacer_presentacion.py): NOM lee la diapositiva oculta, las tres secciones, las
   dos presentaciones personalizadas, los intervalos y «repetir hasta Esc».
2. Se descarga sin tocar: PowerPoint (windows/presentacion_ver.py) dice lo mismo que del original.
3. Se cambia en NOM (mostrar la oculta y ocultar otra, borrar una diapositiva, agregar una, cambiar el nombre de una
   sección, una personalizada nueva, un intervalo y no usar intervalos): PowerPoint dice lo esperado.
4. Una presentación hecha en NOM desde cero, con lo mismo: PowerPoint dice lo esperado.
PowerPoint abre las tres sin avisos ni reparar (office_abre.py) y, si está, el validador de Microsoft no da errores.
Se corre desde Linux (usa PowerPoint por SSH) con el servidor de pruebas en NOM_BASE.
"""
import argparse, base64, json, os, subprocess, sys
from pathlib import Path
AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI)); sys.path.insert(0, str(AQUI.parent / "interfaz"))
import comparar_real as C  # noqa: E402
import nomprueba as N  # noqa: E402
import prueba_ui as U  # noqa: E402

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
res = []


def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det and not ok else ""), flush=True)


JS_BAJAR = r"""async () => { const doc = NOM.editorActual.doc; if (NOM.binarios && NOM.binarios.refsDoc(doc).size) await NOM.binarios.precargar(doc);
  let u = NOM.ooxml.pptx(doc); if (u && u.then) u = await u;
  let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }"""
JS_MODELO = r"""() => { const d = NOM.editorActual.doc, pags = d.paginas(), n = (p) => pags.indexOf(p) + 1;
  return { ocultas: pags.filter((p) => d.valor('_doc', 'oculta:' + p)).map(n),
    secciones: (NOM.ooxml.seccionesPptx(d) || []).map((s) => ({ nombre: s.nombre, diapositivas: s.d.map(n) })),
    personalizadas: (d.valor('_doc', 'personalizadas') || []).map((c) => ({ nombre: c.nombre, diapositivas: (c.d || []).filter((p) => pags.includes(p)).map(n) })),
    intervalos: pags.map((p) => { const t = d.valor('_doc', 'transOffice:' + p); return t && t.advTm != null ? t.advTm / 1000 : null; }),
    configuracion: d.valor('_doc', 'presentacion') || {} }; }"""


def powerpoint_ver(pptx, salida):
    import remoto
    remoto.subir([pptx], "nomoffice/entrada")
    r = remoto.correr_office("pres-" + pptx.stem, "presentacion_ver.py", f"entrada\\{pptx.name}", f"{pptx.stem}.json", timeout=600)
    if not remoto.bajar(f"nomoffice/{pptx.stem}.json", salida / f"{pptx.stem}.pp.json"):
        raise SystemExit("PowerPoint no dejó resultado: " + (r.stdout or "")[-300:])
    return json.loads((salida / f"{pptx.stem}.pp.json").read_text(encoding="utf-8"))


def resumen(v):
    """Lo que se compara de lo que dice PowerPoint."""
    ds = v.get("diapositivas", [])
    return {"diapositivas": len(ds), "ocultas": [d["n"] for d in ds if d["oculta"]],
            "intervalos": [d["segundos"] if d["porTiempo"] else None for d in ds],
            "secciones": [{"nombre": s["nombre"], "diapositivas": list(range(s["primera"], s["primera"] + s["cuantas"])) if s["cuantas"] else []} for s in v.get("secciones", [])],
            "personalizadas": [{"nombre": c["nombre"], "diapositivas": c["diapositivas"]} for c in v.get("personalizadas", [])],
            "repetir": v.get("configuracion", {}).get("repetir"), "usarIntervalos": v.get("configuracion", {}).get("avance") == 2}


def comparar(nombre, esperado, visto):
    for k in esperado:
        anota(f"{nombre}: {k} como en PowerPoint", esperado[k] == visto.get(k), f"esperado {esperado[k]} · PowerPoint {visto.get(k)}")


def abre_sin_avisos(archivos):
    import office_abre
    r = office_abre.abrir(archivos, "presentacion-real")
    for a in archivos:
        x = r.get(a.name)
        anota(f"PowerPoint abre {a.name} sin avisos ni reparar", office_abre.bien(x), json.dumps(x, ensure_ascii=False)[:200] if x else "sin respuesta")
    val = AQUI / "validar_openxml.py"
    if (Path.home() / ".dotnet" / ("dotnet.exe" if os.name == "nt" else "dotnet")).exists():
        v = subprocess.run([sys.executable, str(val)] + [str(a) for a in archivos], capture_output=True, text=True)
        anota("validador de Microsoft: sin errores nuevos", v.returncode == 0, (v.stdout.strip().splitlines() or [v.stderr[-200:]])[-1])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("archivo", nargs="?", default=str(REAL / "archivos/p13-presentacion.pptx"))
    ap.add_argument("--salida", default=str(Path.home() / ".cache/nom-pruebas/registros/presentacion-real"))
    a = ap.parse_args()
    arch = Path(a.archivo); sal = Path(a.salida); sal.mkdir(parents=True, exist_ok=True)
    N.esperar_servidor()
    c, esp = C.entrar()
    nav = N.Navegador(movil=False); nav.entrar(c)
    p = nav.p
    try:
        verdad = resumen(powerpoint_ver(arch, sal))
        # 1. lo que lee NOM del archivo de PowerPoint
        p.goto(N.B + "/#/espacio/" + esp); p.wait_for_timeout(800)
        p.evaluate(C.JS_IMPORTAR, [base64.b64encode(arch.read_bytes()).decode(), arch.name, esp])
        p.wait_for_function(U.JS_LISTO, timeout=120000); p.wait_for_timeout(1000)
        m = p.evaluate(JS_MODELO)
        anota("NOM lee la diapositiva oculta", m["ocultas"] == verdad["ocultas"], f"{m['ocultas']} · PowerPoint {verdad['ocultas']}")
        anota("NOM lee las secciones", m["secciones"] == verdad["secciones"], f"{m['secciones']}")
        anota("NOM lee las presentaciones personalizadas", m["personalizadas"] == verdad["personalizadas"], f"{m['personalizadas']}")
        anota("NOM lee los intervalos", m["intervalos"] == verdad["intervalos"], f"{m['intervalos']} · PowerPoint {verdad['intervalos']}")
        anota("NOM lee «repetir hasta Esc»", bool(m["configuracion"].get("repetir")) == bool(verdad["repetir"]), m["configuracion"])
        # 2. sin tocar: PowerPoint dice lo mismo
        igual = sal / "p13-sin-tocar.pptx"; igual.write_bytes(base64.b64decode(p.evaluate(JS_BAJAR)))
        comparar("sin tocar", verdad, resumen(powerpoint_ver(igual, sal)))
        # 3. cambiado en NOM (como lo hacen los menús: el mismo modelo)
        p.evaluate("""() => { const ed = NOM.editorActual, d = ed.doc, [p1, p2, p3, p4, p5, p6] = d.paginas();
          d.lote(() => {
            d.cambiar('_doc', { ['oculta:' + p3]: null, ['oculta:' + p5]: true });
            d.deLaPagina(p2).forEach((b) => d.cambiar(b.id, { eliminado: true }));
            const np = 'pn' + Date.now().toString(36);
            d.cambiar('_doc', { paginas: [p1, p3, p4, p5, p6, np] });
            const ss = NOM.ooxml.seccionesPptx(d).map((s) => (s.nombre === 'Cierre' ? Object.assign({}, s, { nombre: 'Final' }) : s));
            d.cambiar('_doc', { secciones: ss });
            const ps = (d.valor('_doc', 'personalizadas') || []).slice(); ps.push({ id: '7', nombre: 'Nueva', d: [p6, p1] });
            d.cambiar('_doc', { personalizadas: ps });
            const t5 = d.valor('_doc', 'transOffice:' + p5) || { t: 'corte0', a: {}, dur: 0, spd: 'fast' };
            d.cambiar('_doc', { ['transOffice:' + p5]: Object.assign({}, t5, { advTm: 3000 }), presentacion: { sinIntervalos: true } });
          }); return 1; }""")
        p.wait_for_timeout(1500)
        cambiada = sal / "p13-cambiada.pptx"; cambiada.write_bytes(base64.b64decode(p.evaluate(JS_BAJAR)))
        esperado = {"diapositivas": 6, "ocultas": [4], "intervalos": [None, None, 4.0, 3.0, None, None],
                    "secciones": [{"nombre": "Inicio", "diapositivas": [1]}, {"nombre": "Desarrollo", "diapositivas": [2, 3]}, {"nombre": "Final", "diapositivas": [4, 5, 6]}],
                    "personalizadas": [{"nombre": "Corta", "diapositivas": [1, 3, 5]}, {"nombre": "Al revés", "diapositivas": [5, 4, 3, 2, 1]}, {"nombre": "Nueva", "diapositivas": [5, 1]}],
                    "repetir": False, "usarIntervalos": False}
        comparar("cambiada en NOM", esperado, resumen(powerpoint_ver(cambiada, sal)))
        # 4. hecha en NOM desde cero (sin archivo de origen)
        p.evaluate("""async (esp) => { const ops = []; const d = await NOM.api.post(`/api/espacios/${esp}/documentos`, { nombre: 'Desde NOM', modo: 'PRESENTACION', ops: [] }); location.hash = '#/doc/' + d.id; return 1; }""", esp)
        p.wait_for_function(U.JS_LISTO, timeout=60000); p.wait_for_timeout(800)
        p.evaluate("""() => { const ed = NOM.editorActual, d = ed.doc, ps = ['q1', 'q2', 'q3', 'q4'];
          d.lote(() => {
            d.cambiar('_doc', { paginas: ps });
            ps.forEach((pg, k) => d.cambiar('t' + pg, { tipo: 'texto', pagina: pg, x: 80, y: 60, w: 600, h: 80, html: 'Diapositiva ' + (k + 1) }));
            d.cambiar('_doc', { ['oculta:q2']: true, secciones: [{ id: null, nombre: 'Uno', d: ['q1', 'q2'] }, { id: null, nombre: 'Dos', d: ['q3', 'q4'] }],
              personalizadas: [{ id: '0', nombre: 'Solo pares', d: ['q2', 'q4'] }], presentacion: { repetir: true },
              ['transOffice:q3']: { t: 'corte0', a: {}, dur: 0, spd: 'fast', advTm: 1500 } });
          }); return 1; }""")
        p.wait_for_timeout(1200)
        nueva = sal / "desde-nom.pptx"; nueva.write_bytes(base64.b64decode(p.evaluate(JS_BAJAR)))
        esperado2 = {"diapositivas": 4, "ocultas": [2], "intervalos": [None, None, 1.5, None],
                     "secciones": [{"nombre": "Uno", "diapositivas": [1, 2]}, {"nombre": "Dos", "diapositivas": [3, 4]}],
                     "personalizadas": [{"nombre": "Solo pares", "diapositivas": [2, 4]}], "repetir": True, "usarIntervalos": True}
        comparar("hecha en NOM", esperado2, resumen(powerpoint_ver(nueva, sal)))
        abre_sin_avisos([igual, cambiada, nueva])
        anota("sin errores de JavaScript", not nav.errores, nav.errores[:3])
    finally:
        nav.cerrar()
    print(f"\nPresentación de PowerPoint: {sum(1 for r in res if r[1])} de {len(res)} bien")
    return 0 if all(r[1] for r in res) else 1


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""La vuelta: lo que exporta NOM, abierto con Microsoft Office de verdad (en la Windows de pruebas).

Arma en ~/.cache/nom-pruebas/office-real/vuelta/ los archivos que se mandan a Office:
  - cada archivo hecho por Office (w*, x*, p* de office-real/archivos) importado, EDITADO como lo haría una persona
    (texto, celdas, fórmulas nuevas con matrices dinámicas, LET y LAMBDA, diapositivas) y exportado; también
    pasando por el formato propio (.nom)
  - documentos hechos desde cero en NOM (sin archivo de origen): texto, hoja con fórmulas y presentación
  - los resultados de ediciones.py (corpus de Microsoft y Apache POI editado), si están
  - tres archivos dañados a propósito (control: Office TIENE que quejarse de esos)
Luego los sube a Windows, corre windows/abrir_office.py y trae resultado.json.

  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/vuelta_office.py [--sin-ediciones] [--solo-armar]
"""
import argparse, base64, json, os, re, shutil, subprocess, sys, zipfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import fidelidad as F
import ediciones as E
import office_real as O

REAL = Path(os.environ.get("NOM_PRUEBAS_CACHE", Path.home() / ".cache/nom-pruebas")) / "office-real"
SALE = REAL / "vuelta"

# Ediciones sobre los archivos de Office real (además de las de ediciones.py)
X_MATRICES = """
            const h = doc.vivos().find((b) => b.tipo === "tabla" && !b.enTexto);
            if (!h) return { sinHoja: true };
            doc.cambiar(h.id, { "c:AZ300": "3", "c:AZ301": "1", "c:AZ302": "2", "c:BA300": "=ORDENAR(AZ300:AZ302)", "c:BB300": "=SECUENCIA(2;3;10;5)",
              "c:BF300": "=LET(x;AZ300;y;AZ301;x*y+1)", "c:BG300": "=LAMBDA(n;n*n+1)(6)", "c:BH300": "=FILTRAR(AZ300:AZ302;AZ300:AZ302>1)", "c:BI300": "=UNICOS({1;2;2;3})",
              "c:BJ300": "=SUMA(AZ300:AZ302*AZ300:AZ302)", "c:BK300": "=BYROW({1,2;3,4};LAMBDA(r;SUMA(r)))", "c:BL300": "=MAP({1,2,3};LAMBDA(v;v*10))",
              "c:BM300": "=DIVIDIRTEXTO(\\"a,b;c,d\\";\\",\\";\\";\\")", "c:BO300": "=APILARV({1,2};{3,4})", "c:BQ300": "=TOMAR(AZ300:BA302;2)", "c:BS300": "=SUMA(BA300#)",
              "c:BT300": "=REDUCE(0;AZ300:AZ302;LAMBDA(a;v;a+v))", "c:BU300": "=SCAN(0;{1,2,3};LAMBDA(a;v;a+v))", "c:B2": "Editado en NOM" });
            return {};"""
X09_DATOS = """
            // cambian los datos de los que dependen las matrices: lo que derraman cambia de tamaño
            const hs = doc.vivos().filter((b) => b.tipo === "tabla" && !b.enTexto);
            const d = hs.find((b) => /datos/i.test(b.nombre || ""));
            if (!d) return { sinDatos: true };
            const c = {}; for (let f = 40; f <= 61; f++) ["A", "B", "C", "D", "E", "F", "G"].forEach((l) => { c["c:" + l + f] = null; });
            c["c:E2"] = "99"; c["c:C3"] = "Zoe";
            doc.cambiar(d.id, c);
            return {};"""
CASOS_REALES = {
    "docx": [("texto", E.CASOS["docx"][0][2]), ("agregar", E.CASOS["docx"][1][2]), ("varios", E.CASOS["docx"][3][2])],
    "xlsx": [("celdas", E.CASOS["xlsx"][0][2].replace('const h = "hoja-p1";', 'const h = (doc.vivos().find((b) => b.tipo === "tabla" && !b.enTexto) || {}).id || "hoja-p1";').replace('doc.cambiar(h,', 'doc.cambiar(h,')),
             ("matrices", X_MATRICES)],
    "pptx": [("texto", E.CASOS["pptx"][0][2]), ("mover", E.CASOS["pptx"][1][2]), ("diapositivas", E.CASOS["pptx"][2][2])],
}

# Documentos desde cero (sin archivo de origen)
JS_NUEVOS = r"""
async () => {
  const aB64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const out = {};
  // hoja con fórmulas (también las nuevas), formato y una gráfica
  {
    const doc = NOM.formatos.docDeOps("Libro nuevo", "HOJA", []);
    doc.cambiar("_doc", { paginas: ["p1", "p2"] });
    const c = { tipo: "tabla", pagina: "p1", orden: 0, x: 40, y: 40, w: 880, h: 300, filas: 40, cols: 14, nombre: "Ventas" };
    const datos = [["Región", "Vendedor", "Unidades", "Precio"], ["Norte", "Ana", 12, 25.5], ["Sur", "Luis", 30, 8], ["Centro", "Marta", 7, 450], ["Norte", "Pedro", 22, 320.75], ["Sur", "Sofía", 15, 15.25], ["Centro", "Ana", 40, 3999]];
    datos.forEach((fila, f) => fila.forEach((v, k) => { c["c:" + "ABCD"[k] + (f + 1)] = String(v); }));
    c["n:A1"] = c["n:B1"] = c["n:C1"] = c["n:D1"] = true; c["r:A1"] = c["r:B1"] = c["r:C1"] = c["r:D1"] = "#DDEBF7";
    Object.assign(c, { "c:E1": "Total", "c:E2": "=C2:C7*D2:D7", "f:E2": "$#,##0.00", "c:G1": "Ordenado", "c:G2": "=ORDENAR(B2:B7)", "c:H1": "Únicos", "c:H2": "=UNICOS(A2:A7)",
      "c:I1": "Filtro", "c:I2": "=FILTRAR(B2:B7;C2:C7>14)", "c:J1": "Serie", "c:J2": "=SECUENCIA(4;2;1;1)", "c:L1": "LET", "c:L2": "=LET(t;SUMA(C2:C7);t/CONTAR(C2:C7))",
      "c:L3": "=LAMBDA(x;x^2)(9)", "c:L4": "=SUMA(E2#)", "c:L5": "=BUSCARX(\"Marta\";B2:B7;D2:D7)", "c:L6": "=SI.ERROR(1/0;\"sin dato\")", "c:L7": "=TEXTO(HOY();\"dd/mm/aaaa\")",
      "c:L8": "=PROMEDIO.SI(A2:A7;\"Norte\";C2:C7)", "c:L9": "=CONTAR.SI.CONJUNTO(A2:A7;\"Sur\";C2:C7;\">10\")", "c:L10": "=REDONDEAR(PI();4)" });
    doc.cambiar("hoja-p1", c);
    doc.cambiar("hoja-p2", { tipo: "tabla", pagina: "p2", orden: 0, x: 40, y: 40, w: 880, h: 300, filas: 20, cols: 8, nombre: "Resumen", "c:A1": "Total vendido", "c:B1": "=SUMA(Ventas!E2#)", "c:A2": "Mayor", "c:B2": "=MAX(Ventas!C2:C7)" });
    out.xlsx = aB64(await NOM.ooxml.xlsx(doc));
  }
  // texto con títulos, listas, tabla y formato
  {
    const doc = NOM.formatos.docDeOps("Documento nuevo", "TEXTO", []);
    const sec = doc.secuencia("cuerpo");
    let iz = sec.fin ? null : null;
    const pon = (x, a) => { const vis = doc.secuencia("cuerpo").visibles(); doc.seq("cuerpo", Object.assign({ k: "i", iz: vis.length ? vis[vis.length - 1].id : null, x }, a ? { a } : {})); };
    pon("Informe hecho en NOM"); pon("\n", { estilo: "h1" });
    pon("Un párrafo con "); pon("negrita", { b: true }); pon(", "); pon("cursiva", { i: true }); pon(" y "); pon("color", { color: "#C2372B" }); pon(".\n");
    pon("Primer punto"); pon("\n", { lista: "numeros" }); pon("Segundo punto"); pon("\n", { lista: "numeros" }); pon("Viñeta"); pon("\n", { lista: "vinetas" });
    pon("Cierre del documento.");
    out.docx = aB64(await NOM.ooxml.docx(doc));
  }
  // presentación con título, figuras y una tabla
  {
    const doc = NOM.formatos.docDeOps("Presentación nueva", "PRESENTACION", []);
    doc.cambiar("_doc", { paginas: ["d1", "d2"] });
    doc.cambiar("t1", { tipo: "texto", pagina: "d1", orden: 1, z: 1, x: 80, y: 60, w: 800, h: 100, estilo: "h1", html: "Presentación hecha en NOM" });
    doc.cambiar("f1", { tipo: "forma", figura: "elipse", pagina: "d1", orden: 2, z: 2, x: 120, y: 220, w: 240, h: 160, color: "#F06A3D", texto: "Círculo" });
    doc.cambiar("f2", { tipo: "forma", figura: "rect", pagina: "d1", orden: 3, z: 3, x: 420, y: 220, w: 260, h: 160, color: "#2F6FB0", texto: "Cuadro" });
    doc.cambiar("t2", { tipo: "texto", pagina: "d2", orden: 1, z: 1, x: 80, y: 60, w: 800, h: 90, estilo: "h1", html: "Segunda diapositiva" });
    doc.cambiar("t3", { tipo: "texto", pagina: "d2", orden: 2, z: 2, x: 80, y: 180, w: 800, h: 200, html: "<ul><li>Uno</li><li>Dos</li><li>Tres</li></ul>" });
    out.pptx = aB64(await NOM.ooxml.pptx(doc, false));
  }
  return out;
}
"""


def dañados(dest):
    """Tres archivos que Office NO puede abrir sin quejarse (control de que la prueba sí detecta los avisos)."""
    for ext, parte in (("docx", "word/document.xml"), ("xlsx", "xl/worksheets/sheet1.xml"), ("pptx", "ppt/slides/slide1.xml")):
        orig = next((p for p in sorted((REAL / "archivos").glob(f"*.{ext}"))), None)
        if not orig: continue
        sal = dest / f"zz-control-dañado.{ext}"
        zi = zipfile.ZipFile(orig)
        with zipfile.ZipFile(sal, "w", zipfile.ZIP_DEFLATED) as zo:
            for n in zi.namelist():
                d = zi.read(n)
                if n == parte: d = d.replace(b"</", b"<", 3)   # XML roto a propósito
                zo.writestr(zi.getinfo(n), d)


def armar(con_ediciones=True):
    if SALE.exists(): shutil.rmtree(SALE)
    SALE.mkdir(parents=True)
    from playwright.sync_api import sync_playwright
    n = 0
    with sync_playwright() as pw:
        nav = pw.chromium.launch(); pag = nav.new_page(); errores = []
        pag.on("pageerror", lambda e: errores.append(str(e)))
        pag.goto(F.BASE + "/"); pag.wait_for_function("() => window.NOM && NOM.ooxml && NOM.formatos && NOM.formatos.docDeOps", timeout=30000)
        for orig in sorted((REAL / "archivos").glob("*.*")):
            ext = orig.suffix[1:]
            if ext not in CASOS_REALES: continue
            casos = list(CASOS_REALES[ext])
            if orig.stem.startswith("x09"): casos.append(("datos", X09_DATOS))
            for nombre_caso, edicion in casos:
                r = pag.evaluate(E.JS.replace("/*EDICION*/", edicion), [base64.b64encode(orig.read_bytes()).decode(), orig.name])
                if r.get("error"): print(f"  MAL al editar {orig.name} ({nombre_caso}): {r['error'][:200]}"); continue
                for suf, clave in (("", "salida"), ("-vianom", "salidaNom")):
                    (SALE / f"{orig.stem}.{nombre_caso}{suf}.{ext}").write_bytes(base64.b64decode(r[clave])); n += 1
        r = pag.evaluate(JS_NUEVOS)
        for ext, b64 in r.items(): (SALE / f"nuevo-nom.{ext}").write_bytes(base64.b64decode(b64)); n += 1
        if errores: print("  errores de JavaScript:", errores[:3])
        nav.close()
    if con_ediciones:
        ed = F.CACHE / "office-ediciones"
        for p in sorted(ed.glob("*.*")) if ed.exists() else []:
            if p.suffix in (".docx", ".xlsx", ".pptx"): shutil.copy(p, SALE / ("corpus-" + p.name)); n += 1
    dañados(SALE)
    print(f"{n} archivos exportados por NOM + 3 de control en {SALE}")


def mandar_y_abrir():
    O.ssh("rmdir /S /Q nomoffice\\entrada\\vuelta 2>nul & mkdir nomoffice\\entrada\\vuelta & echo.")
    archivos = [str(p) for p in sorted(SALE.glob("*.*")) if p.suffix in (".docx", ".xlsx", ".pptx")]
    for k in range(0, len(archivos), 60):
        subprocess.run(["scp", "-q"] + archivos[k:k + 60] + [O.WIN + ":nomoffice/entrada/vuelta/"], check=True)
    subprocess.run(["scp", "-q", str(Path(__file__).parent / "windows" / "abrir_office.py"), O.WIN + ":nomoffice/"], check=True)
    O.correr("vuelta", "abrir_office.py", "entrada\\vuelta")
    subprocess.run(["scp", "-q", O.WIN + ":nomoffice/entrada/vuelta/resultado.json", str(SALE / "resultado.json")], check=True)
    res = json.loads((SALE / "resultado.json").read_text(encoding="utf-8"))
    control = [r for r in res if r["archivo"].startswith("zz-control")]
    nuestros = [r for r in res if not r["archivo"].startswith("zz-control")]
    mal = [r for r in nuestros if not r.get("abrio") or r.get("avisos") or r.get("reparado")]
    ctrl_ok = all((not r.get("abrio")) or r.get("avisos") or r.get("reparado") for r in control) and len(control) == 3
    print(f"\nControl (archivos dañados a propósito): {'Office se quejó de los 3 (la prueba sí detecta)' if ctrl_ok else 'NO SE DETECTÓ ALGUNO: ' + json.dumps(control, ensure_ascii=False)[:600]}")
    print(f"Exportados por NOM: {len(nuestros)} · abrieron sin avisos ni reparar: {len(nuestros) - len(mal)} · con algo: {len(mal)}")
    for r in mal: print("  ", r["archivo"], json.dumps({k: r.get(k) for k in ("abrio", "error", "avisos", "reparado", "registro_reparacion")}, ensure_ascii=False)[:700])
    xl = [r for r in nuestros if r.get("app") == "excel"]
    print(f"Excel: fórmulas leídas {sum(r.get('formulas', 0) for r in xl)}, con #NAME? {sum(r.get('nombre_mal', 0) for r in xl)}, matrices que derraman {sum(r.get('derraman', 0) for r in xl)}")
    for r in xl:
        if r.get("nombre_mal"): print("   #NAME? en", r["archivo"], r.get("nombre_mal"))
    return 0 if not mal and ctrl_ok else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--sin-ediciones", action="store_true")
    ap.add_argument("--solo-armar", action="store_true")
    a = ap.parse_args()
    armar(not a.sin_ediciones)
    if not a.solo_armar: sys.exit(mandar_y_abrir())

#!/usr/bin/env python3
"""La vuelta de lo que arman los paneles de presentación: animaciones, transiciones, vínculos, medios y sonidos.

Lo hecho en NOM tiene que volver a NOM igual, abrirse en PowerPoint de verdad sin «reparar» y pasar el validador
de Microsoft.

  vuelta_paneles.py [--salida DIR] [--sin-office]

Arma las presentaciones con las mismas funciones que usan los paneles (animaciones-panel.js → NOM.animOffice,
transiciones → NOM.transOffice, acción al hacer clic → bloque.accion):
  cero     desde cero, sin archivo de origen: entrada, énfasis, salida y trayectorias con todas las opciones del
           panel (repetir 3 veces, hasta el siguiente clic, hasta el final; rebobinar; ida y vuelta; inicio y fin
           suaves; rebote; atenuar a un color; ocultar después y con el siguiente clic; por letra y por palabra al
           revés; por párrafos al revés y con la forma; trayectoria dibujada abierta y cerrada; desencadenador por
           figura; cambiar la dirección; girar a la izquierda; aumentar 150 %; color de énfasis), una transición
           distinta por diapositiva (con opciones, duración, avance solo o sin clic) y las nueve acciones al hacer clic.
  medios   p08-animaciones.pptx de PowerPoint (con video y audio): efectos y vínculos nuevos en sus figuras,
           transiciones cambiadas, y se exporta sobre su original (los medios tienen que seguir).
  sonidos  dos .pptx con sonidos hechos con PowerPoint: v04-sonidos (windows/hacer_sonidos.py: sonido propio en una
           transición, en un efecto, «detener el sonido anterior» y en bucle) y una plantilla de Microsoft con sonido en
           un efecto; se cambian la transición y las opciones del efecto que suena: el sonido tiene que quedarse.
Para cada una: exporta, vuelve a importar lo exportado y compara efecto por efecto (clase, número, dirección,
inicio, retraso, duración, opciones, desencadenador), la transición y la acción de cada figura. Luego la sube a la
Windows, la abre en PowerPoint (windows/abrir_office.py: avisos, «reparar») y corre validar_openxml.py.

Necesita el servidor de pruebas (NOM_BASE, por omisión el 8099) y, para Office, la Windows (office_real.py).
Sale con código 1 si algo falla.
"""
import argparse, base64, json, shutil, subprocess, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import fidelidad as F

REAL = F.CACHE / "office-real" / "archivos"

JS = r"""
async ([b64, nombre, modo]) => {
  const A = NOM.animOffice, TO = NOM.transOffice;
  const aB64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const deB64 = (s) => { const bin = atob(s), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
  const r2 = (v) => (v == null ? null : typeof v === "number" ? Math.round(v * 100) / 100 : v);
  // nombre estable de un bloque (los ids cambian al importar): su texto, o su tipo y su orden en la diapositiva
  const nombreDe = (doc, pg, id) => {
    const b = doc.bloque(id) || {};
    const t = ((NOM.diapo && NOM.diapo.textoPlano(b.html || "")) || b.texto || "").trim().slice(0, 40);
    if (t) return t;
    const mismos = doc.deLaPagina(pg).filter((q) => q.tipo === b.tipo).sort((a, c) => (a.z || 0) - (c.z || 0));
    return (b.tipo || "?") + "#" + mismos.findIndex((q) => q.id === id);
  };
  const opc = (x) => { const o = A.opcionesDe(x) || {}; const s = {}; Object.keys(o).sort().forEach((k) => { const v = o[k]; s[k] = v && typeof v === "object" ? JSON.stringify(v) : r2(v); }); return s; };
  const resumen = (doc) => doc.paginas().map((p) => {
    const l = doc.valor("_doc", "anim:" + p) || [];
    const t = doc.valor("_doc", "transOffice:" + p);
    return {
      anim: l.map((x) => ({ quien: nombreDe(doc, p, x.id), clave: A.clavePreset(x) || x.efecto || "?", tipo: x.tipo, inicio: x.inicio || "clic",
        retraso: Math.round(+x.retraso || 0), dur: Math.round(+x.dur || 0), disp: x.disp ? nombreDe(doc, p, x.disp) : "", par: x.par == null ? null : x.par, opc: opc(x) })),
      trans: t ? { t: t.t, a: JSON.stringify(t.a || {}), dur: Math.round(+t.dur || 0), advTm: t.advTm == null ? null : Math.round(t.advTm), sinClic: !!t.sinClic, snd: !!t.snd, sndFin: !!t.sndFin, bucle: !!(t.snd && t.snd.bucle) } : null,
      acc: doc.deLaPagina(p).filter((b) => b.accion).map((b) => [nombreDe(doc, p, b.id), b.accion.tipo, b.accion.url || "", b.accion.pagina ? doc.paginas().indexOf(b.accion.pagina) : -1]).sort(),
    };
  });
  const r = { modo };
  try {
    let doc;
    if (modo === "cero") {
      doc = NOM.formatos.docDeOps("Paneles desde cero", "PRESENTACION", []);
      doc.cambiar("_doc", { paginas: ["d1", "d2", "d3", "d4", "d5"] });
      const pon = (id, o) => doc.cambiar(id, Object.assign({ orden: 1, z: 1 }, o));
      // diapositiva 1: texto por párrafos, figura con rebote y repetir, disparador
      pon("t1", { tipo: "texto", pagina: "d1", orden: 1, z: 1, x: 60, y: 30, w: 840, h: 80, estilo: "h1", html: "Animaciones hechas con el panel" });
      pon("t2", { tipo: "texto", pagina: "d1", orden: 2, z: 2, x: 60, y: 140, w: 520, h: 300, html: "<p>Primer párrafo de la lista</p><p>Segundo párrafo de la lista</p><p>Tercer párrafo de la lista</p>" });
      pon("f1", { tipo: "forma", figura: "elipse", pagina: "d1", orden: 3, z: 3, x: 640, y: 150, w: 240, h: 150, color: "#F06A3D", texto: "Rebote" });
      pon("f2", { tipo: "forma", figura: "rect", pagina: "d1", orden: 4, z: 4, x: 640, y: 340, w: 240, h: 110, color: "#2F6FB0", texto: "Disparador" });
      pon("f3", { tipo: "forma", figura: "rect", pagina: "d1", orden: 5, z: 5, x: 640, y: 470, w: 240, h: 50, color: "#3D9A5B", texto: "Se enciende" });
      // diapositiva 2: énfasis con todas sus opciones
      pon("t3", { tipo: "texto", pagina: "d2", orden: 1, z: 1, x: 60, y: 30, w: 840, h: 80, estilo: "h1", html: "Énfasis con opciones" });
      pon("f4", { tipo: "forma", figura: "rect", pagina: "d2", orden: 2, z: 2, x: 80, y: 180, w: 220, h: 140, color: "#8A4FBF", texto: "Girar" });
      pon("f5", { tipo: "forma", figura: "elipse", pagina: "d2", orden: 3, z: 3, x: 360, y: 180, w: 220, h: 140, color: "#D9A21B", texto: "Aumentar" });
      pon("f6", { tipo: "forma", figura: "rect", pagina: "d2", orden: 4, z: 4, x: 640, y: 180, w: 220, h: 140, color: "#2F6FB0", texto: "Color" });
      pon("t4", { tipo: "texto", pagina: "d2", orden: 5, z: 5, x: 80, y: 380, w: 780, h: 90, html: "<p>Texto que aparece letra por letra al revés</p>" });
      // diapositiva 3: salidas y después de la animación
      pon("t5", { tipo: "texto", pagina: "d3", orden: 1, z: 1, x: 60, y: 30, w: 840, h: 80, estilo: "h1", html: "Salidas y después de la animación" });
      pon("f7", { tipo: "forma", figura: "rect", pagina: "d3", orden: 2, z: 2, x: 80, y: 180, w: 220, h: 140, color: "#C2372B", texto: "Atenuar" });
      pon("f8", { tipo: "forma", figura: "rect", pagina: "d3", orden: 3, z: 3, x: 360, y: 180, w: 220, h: 140, color: "#C2372B", texto: "Ocultar después" });
      pon("f9", { tipo: "forma", figura: "rect", pagina: "d3", orden: 4, z: 4, x: 640, y: 180, w: 220, h: 140, color: "#C2372B", texto: "Ocultar al clic" });
      pon("t6", { tipo: "texto", pagina: "d3", orden: 5, z: 5, x: 80, y: 380, w: 780, h: 90, html: "<p>Palabra por palabra hacia atrás</p>" });
      // diapositiva 4: trayectorias
      pon("t7", { tipo: "texto", pagina: "d4", orden: 1, z: 1, x: 60, y: 30, w: 840, h: 80, estilo: "h1", html: "Trayectorias" });
      pon("f10", { tipo: "forma", figura: "elipse", pagina: "d4", orden: 2, z: 2, x: 80, y: 200, w: 120, h: 120, color: "#3D9A5B", texto: "Abierta" });
      pon("f11", { tipo: "forma", figura: "elipse", pagina: "d4", orden: 3, z: 3, x: 420, y: 200, w: 120, h: 120, color: "#2F6FB0", texto: "Cerrada" });
      pon("f12", { tipo: "forma", figura: "elipse", pagina: "d4", orden: 4, z: 4, x: 720, y: 200, w: 120, h: 120, color: "#F06A3D", texto: "Del catálogo" });
      // diapositiva 5: los botones de acción
      pon("t8", { tipo: "texto", pagina: "d5", orden: 1, z: 1, x: 60, y: 20, w: 840, h: 70, estilo: "h1", html: "Acciones al hacer clic" });
      const acciones = [["sig", "Siguiente"], ["ant", "Anterior"], ["primera", "Primera"], ["ultima", "Última"], ["volver", "Volver"], ["fin", "Terminar"], ["diapo", "Ir a la 3"], ["url", "Página web"], ["url", "Correo"]];
      acciones.forEach(([t, txt], k) => {
        const a = { tipo: t };
        if (t === "diapo") a.pagina = "d3";
        if (t === "url") a.url = txt === "Correo" ? "mailto:nom@example.com" : "https://www.nom.qd.je/";
        pon("b" + k, { tipo: "forma", figura: "rect", pagina: "d5", orden: 2 + k, z: 2 + k, x: 60 + (k % 3) * 290, y: 120 + Math.floor(k / 3) * 130, w: 260, h: 100, color: "#2F6FB0", texto: txt, accion: a });
      });
      const T = { w: 960, h: 540 };
      const l1 = [];
      l1.push(A.nuevoEfecto("t1", "entr:10:0", "clic"));                                           // Desvanecer
      l1.push(...A.porParrafos(A.nuevoEfecto("t2", "entr:2:4", "clic"), [0, 1, 2], true, true));     // Desplazar por párrafos, al revés, con la forma
      l1.push(A.conOpciones(Object.assign(A.nuevoEfecto("f1", "entr:53:16", "despues"), { retraso: 250, dur: 900 }), { rebote: 0.3, repetir: 3, suaveIni: 0.2, suaveFin: 0.2 }));
      l1.push(Object.assign(A.conOpciones(A.nuevoEfecto("f3", "emph:26:0", "clic"), { repetir: "clic" }), { disp: "f2" }));   // pulso por desencadenador, hasta el siguiente clic
      const l2 = [];
      l2.push(Object.assign(A.conOpciones(A.nuevoEfecto("f4", "emph:8:0", "clic"), { giro: -720, autoRev: true }), { dur: 1500 }));
      l2.push(A.conOpciones(A.nuevoEfecto("f5", "emph:6:0", "con"), { escala: 150, rebobinar: true, repetir: "fin" }));
      const col = A.nuevoEfecto("f6", "emph:19:0", "clic");
      l2.push(col ? A.conOpciones(col, { color: "#C2372B" }) : A.nuevoEfecto("f6", "emph:26:0", "clic"));
      const vari = A.variantes(A.nuevoEfecto("t4", "entr:2:4", "clic"));
      l2.push(A.conOpciones(A.cambiarEfecto(A.nuevoEfecto("t4", "entr:2:4", "clic"), (vari[vari.length - 1] || {}).clave || "entr:2:4"), { texto: "letra", pct: 15, atras: true }));
      const l3 = [];
      l3.push(A.conOpciones(A.nuevoEfecto("f7", "entr:10:0", "clic"), { despues: { color: "#C8C8C8" } }));
      l3.push(A.conOpciones(A.nuevoEfecto("f8", "emph:26:0", "clic"), { despues: "ocultar" }));
      l3.push(A.conOpciones(A.nuevoEfecto("f9", "entr:10:0", "clic"), { despues: "ocultarClic" }));
      l3.push(A.conOpciones(A.nuevoEfecto("t6", "entr:10:0", "despues"), { texto: "palabra", pct: 30, atras: true }));
      l3.push(A.nuevoEfecto("f7", "exit:10:0", "clic"));
      const l4 = [];
      const abierta = [[0, 0], [60, -40], [140, -60], [220, -20], [300, 40], [360, 120]];
      const cerrada = [[0, 0], [80, -60], [160, 0], [80, 60], [2, 3]];
      l4.push(A.trayectoriaPropia("f10", abierta, T, false));
      l4.push(A.trayectoriaPropia("f11", cerrada, T, true));
      const cat = Object.keys((NOM.animOfficeDatos || {}).presets || {}).filter((k) => k.startsWith("path:")).sort()[0];
      if (cat) l4.push(A.nuevoEfecto("f12", cat, "clic"));
      doc.cambiar("_doc", { "anim:d1": l1, "anim:d2": l2, "anim:d3": l3, "anim:d4": l4 });
      // una transición distinta por diapositiva (con opciones, duración, avance)
      const quiero = ["Desvanecer", "Empujar", "Barrido", "Cubo", "Galería"];
      doc.paginas().forEach((p, k) => {
        const e = Math.max(1, TO.CATALOGO.findIndex((x) => x[0] === quiero[k]));
        const vs = TO.CATALOGO[e][3];
        const tr = Object.assign({}, TO.deVariante(TO.CATALOGO[e], Math.min(vs.length - 1, k % Math.max(1, vs.length))));
        tr.dur = 700 + 150 * k;
        if (k === 1) tr.advTm = 4000;
        if (k === 2) { tr.advTm = 2500; tr.sinClic = true; }
        doc.cambiar("_doc", { ["transOffice:" + p]: tr });
      });
    } else {
      const imp = await NOM.formatos.importar(new File([deB64(b64)], nombre));
      doc = NOM.formatos.docDeOps(imp.nombre, imp.modo, imp.ops);
      r.importado = resumen(doc);
      const pgs = doc.paginas();
      const figuras = (p) => doc.deLaPagina(p).filter((b) => b.tipo === "forma" || b.tipo === "texto").sort((a, c) => (a.z || 0) - (c.z || 0));
      if (modo === "medios") {
        // tres diapositivas sin animaciones: efectos nuevos, vínculo, transición cambiada
        let hechas = 0;
        for (const p of pgs) {
          if (hechas >= 3) break;
          const l = doc.valor("_doc", "anim:" + p) || [];
          const fs = figuras(p);
          if (l.length || fs.length < 2) continue;
          doc.cambiar("_doc", { ["anim:" + p]: [A.nuevoEfecto(fs[0].id, "entr:10:0", "clic"), A.conOpciones(A.nuevoEfecto(fs[1].id, "emph:26:0", "despues"), { repetir: 2 })] });
          doc.cambiar(fs[1].id, { accion: hechas === 0 ? { tipo: "sig" } : hechas === 1 ? { tipo: "url", url: "https://www.nom.qd.je/" } : { tipo: "diapo", pagina: pgs[0] } });
          const e = TO.CATALOGO.findIndex((x) => x[0] === "Empujar");
          doc.cambiar("_doc", { ["transOffice:" + p]: Object.assign({}, TO.deVariante(TO.CATALOGO[e], 0), { dur: 900 }) });
          hechas++;
        }
        r.hechas = hechas;
      } else {
        // sonidos: cambiar la transición (el sonido se queda, como en el panel) y las opciones de cada efecto que suena
        for (const p of pgs) {
          const t = doc.valor("_doc", "transOffice:" + p);
          if (t && (t.snd || t.sndFin)) {
            const e = TO.CATALOGO.findIndex((x) => x[0] === "Barrido");
            const nueva = Object.assign({}, TO.deVariante(TO.CATALOGO[e], 0), { dur: 1100 });
            if (t.snd) nueva.snd = t.snd; if (t.sndFin) nueva.sndFin = true;
            doc.cambiar("_doc", { ["transOffice:" + p]: nueva });
          }
          const l = doc.valor("_doc", "anim:" + p) || [];
          if (l.length) doc.cambiar("_doc", { ["anim:" + p]: l.map((x) => (JSON.stringify(x).includes("snd") ? A.conOpciones(Object.assign({}, x, { dur: Math.round((+x.dur || 500) * 1.4) }), { suaveIni: 0.25 }) : x)) });
        }
      }
    }
    r.esperado = resumen(doc);
    const sal = await NOM.ooxml.pptx(doc, false);
    r.pptx = aB64(sal);
    const imp2 = await NOM.formatos.importar(new File([sal], "vuelta.pptx"));
    const doc2 = NOM.formatos.docDeOps(imp2.nombre, imp2.modo, imp2.ops);
    r.despues = resumen(doc2);
    // y por el formato propio (.nom): guardarlo, abrirlo y exportarlo otra vez
    const nom = NOM.formatos.aNomBytes(doc);
    r.nomBytes = nom.length;
    const imp3 = await NOM.formatos.importar(new File([nom], "vuelta.nom"));
    const doc3 = NOM.formatos.docDeOps(imp3.nombre, imp3.modo, imp3.ops);
    r.vianom = resumen(doc3);
    const sal3 = await NOM.ooxml.pptx(doc3, false);
    const imp4 = await NOM.formatos.importar(new File([sal3], "vuelta-nom.pptx"));
    r.nomDespues = resumen(NOM.formatos.docDeOps(imp4.nombre, imp4.modo, imp4.ops));
    r.pptxNom = aB64(sal3);
  } catch (e) { r.error = String(e && e.stack || e).slice(0, 1500); }
  return r;
}
"""


def comparar(esperado, despues):
    """Diferencias diapositiva por diapositiva (lista de textos)."""
    dif = []
    if len(esperado) != len(despues):
        return [f"diapositivas: {len(esperado)} → {len(despues)}"]
    for k, (a, b) in enumerate(zip(esperado, despues), 1):
        if len(a["anim"]) != len(b["anim"]):
            dif.append(f"d{k}: {len(a['anim'])} efectos → {len(b['anim'])}")
        else:
            for j, (x, y) in enumerate(zip(a["anim"], b["anim"]), 1):
                for c in ("quien", "clave", "tipo", "inicio", "retraso", "dur", "disp", "par"):
                    if x[c] != y[c]: dif.append(f"d{k} efecto {j} {c}: {x[c]!r} → {y[c]!r}")
                for o in sorted(set(x["opc"]) | set(y["opc"])):
                    u, v = x["opc"].get(o), y["opc"].get(o)
                    if isinstance(u, (int, float)) and isinstance(v, (int, float)) and abs(u - v) <= 0.02: continue
                    if u != v: dif.append(f"d{k} efecto {j} opción {o}: {u!r} → {v!r}")
        if a["trans"] != b["trans"]: dif.append(f"d{k} transición: {a['trans']} → {b['trans']}")
        if a["acc"] != b["acc"]: dif.append(f"d{k} acciones: {a['acc']} → {b['acc']}")
    return dif


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--salida", default=str(F.CACHE / "vuelta-paneles"))
    ap.add_argument("--sin-office", action="store_true")
    a = ap.parse_args()
    sal = Path(a.salida)
    if sal.exists(): shutil.rmtree(sal)
    sal.mkdir(parents=True)
    casos = [("cero", None), ("medios", REAL / "p08-animaciones.pptx"),
             ("sonidos", REAL / "v04-sonidos.pptx"),   # hecho con PowerPoint (windows/hacer_sonidos.py)
             ("sonidos", F.CACHE / "corpus/pptx/ms-rendimiento-tipico.pptx")]
    mal = 0
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        nav = pw.chromium.launch(); pag = nav.new_page(); errores = []
        pag.on("pageerror", lambda e: errores.append(str(e)))
        pag.goto(F.BASE + "/")
        pag.wait_for_function("() => window.NOM && NOM.ooxml && NOM.formatos && NOM.animOffice && NOM.transOffice && NOM.animOfficeDatos", timeout=60000)
        for modo, ruta in casos:
            b64 = base64.b64encode(ruta.read_bytes()).decode() if ruta else ""
            nombre = ruta.name if ruta else "paneles-desde-cero.pptx"
            r = pag.evaluate(JS, [b64, nombre, modo])
            if r.get("error"):
                print(f"MAL  {nombre}: {r['error'][:600]}"); mal += 1; continue
            destino = sal / f"{Path(nombre).stem}.paneles-{modo}.pptx"
            destino.write_bytes(base64.b64decode(r["pptx"]))
            (sal / f"{Path(nombre).stem}.paneles-{modo}-vianom.pptx").write_bytes(base64.b64decode(r["pptxNom"]))
            (sal / (destino.stem + ".json")).write_text(json.dumps({k: v for k, v in r.items() if k not in ("pptx", "pptxNom")}, ensure_ascii=False, indent=1), encoding="utf-8")
            efectos = sum(len(d["anim"]) for d in r["esperado"])
            trans = sum(1 for d in r["esperado"] if d["trans"])
            acc = sum(len(d["acc"]) for d in r["esperado"])
            dif = comparar(r["esperado"], r["despues"])
            difn = comparar(r["esperado"], r["vianom"]) + [("(exportado desde el .nom) " + x) for x in comparar(r["esperado"], r["nomDespues"])]
            ok = not dif and not difn
            print(f"{'BIEN' if ok else 'MAL '} {nombre} ({modo}): {efectos} efectos, {trans} transiciones, {acc} acciones; ida y vuelta {len(dif)} diferencias; por .nom ({r['nomBytes']} bytes) {len(difn)}")
            for d in (dif + [("(.nom) " + x) for x in difn])[:25]: print("     ", d)
            if dif or difn: mal += 1
            if modo == "medios":
                import zipfile
                z = zipfile.ZipFile(destino)
                medios = [n for n in z.namelist() if n.startswith("ppt/media/") and n.rsplit(".", 1)[-1] in ("mp4", "m4a", "wav", "mp3")]
                print(f"      medios dentro del .pptx exportado: {medios}")
                if len(medios) < 2: mal += 1; print("      MAL: se perdió el video o el audio")
            if modo == "sonidos":
                import zipfile
                z = zipfile.ZipFile(destino)
                snd = sum(1 for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml") and (b"<p:snd" in z.read(n) or b"<p:stSnd" in z.read(n)))
                print(f"      diapositivas con sonido en el exportado: {snd}")
                if not snd: mal += 1; print("      MAL: se perdió el sonido")
        if errores: print("errores de JavaScript:", errores[:3]); mal += 1
        nav.close()
    archivos = sorted(sal.glob("*.pptx"))
    # el validador de Microsoft (0 errores nuevos contra el original)
    r = subprocess.run([sys.executable, str(Path(__file__).parent / "validar_openxml.py"), str(sal), "--corpus", str(REAL), "--corpus", str(F.CACHE / "corpus-reales")], capture_output=True, text=True)
    print(r.stdout[-2500:], r.stderr[-800:])
    if r.returncode: mal += 1
    if not a.sin_office:
        import office_real as O
        O.ssh("rmdir /S /Q nomoffice\\entrada\\paneles 2>nul & mkdir nomoffice\\entrada\\paneles & echo.")
        # también los originales: si PowerPoint no abre el original (archivos de prueba de LibreOffice), lo exportado
        # por NOM se juzga contra eso y no cuenta como falla de NOM
        originales = [ruta for _, ruta in casos if ruta]
        for o in originales: shutil.copy(o, sal / ("original-" + o.name))
        subprocess.run(["scp", "-q"] + [str(p) for p in archivos] + [str(sal / ("original-" + o.name)) for o in originales] + [O.WIN + ":nomoffice/entrada/paneles/"], check=True)
        subprocess.run(["scp", "-q", str(Path(__file__).parent / "windows" / "abrir_office.py"), O.WIN + ":nomoffice/"], check=True)
        O.correr("paneles", "abrir_office.py", "entrada\\paneles")
        subprocess.run(["scp", "-q", O.WIN + ":nomoffice/entrada/paneles/resultado.json", str(sal / "resultado.json")], check=True)
        res = json.loads((sal / "resultado.json").read_text(encoding="utf-8"))
        bien_abre = lambda x: bool(x and x.get("abrio") and not x.get("avisos") and not x.get("reparado"))
        orig_de = {x["archivo"][len("original-"):].rsplit(".", 1)[0]: x for x in res if x["archivo"].startswith("original-")}
        for x in res:
            if x["archivo"].startswith("original-"): continue
            ok = bien_abre(x)
            o = orig_de.get(x["archivo"].split(".")[0])
            igual = not ok and o is not None and not bien_abre(o)
            print(f"{'BIEN' if ok or igual else 'MAL '} PowerPoint abre {x['archivo']}: {json.dumps({k: x.get(k) for k in ('abrio', 'avisos', 'reparado', 'diapositivas', 'error') if x.get(k) is not None}, ensure_ascii=False)[:300]}{'  (el original tampoco abre en PowerPoint)' if igual else ''}")
            if not (ok or igual): mal += 1
    print("Resultado:", "todo bien" if not mal else f"{mal} con falla")
    return 1 if mal else 0


if __name__ == "__main__":
    sys.exit(main())

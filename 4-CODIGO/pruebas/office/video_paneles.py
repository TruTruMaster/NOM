#!/usr/bin/env python3
"""¿Lo que arma el panel de animaciones se VE igual en PowerPoint? Graba el video en PowerPoint y lo compara con NOM.

  video_paneles.py [--sin-office] [--cuadros 8]

Arma en NOM (con las funciones del panel: NOM.animOffice) una presentación de una diapositiva por efecto, cada uno
con opciones del panel (rebote, inicio y fin suaves, repetir, ida y vuelta, rebobinar, atenuar a un color, ocultar
después, por letra y por palabra al revés, por párrafos, trayectorias dibujadas abierta y cerrada, color de
énfasis…), que empieza solo («después de la anterior») y con la marca binaria de animaciones_real.py abajo a la
izquierda. La exporta a .pptx, PowerPoint de verdad hace su video (windows/animaciones_office.py video) y
animaciones_real.py compara cuadro por cuadro el video con el motor de NOM.

Deja todo en ~/.cache/nom-pruebas/office-real/archivos/v02-paneles.{pptx,json,mp4} y el reporte en
~/.cache/nom-pruebas/animaciones/v02-paneles/. Necesita el servidor de pruebas (NOM_BASE) y la Windows.
"""
import argparse, base64, json, os, subprocess, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import fidelidad as F

NOMBRE = "v02-paneles"
REAL = F.CACHE / "office-real" / "archivos"

JS = r"""
async () => {
  const A = NOM.animOffice;
  const aB64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const doc = NOM.formatos.docDeOps("Paneles en video", "PRESENTACION", []);
  const T = { w: 960, h: 540 };
  const casos = [];
  // [nombre, figura o texto, armar(id) -> lista de efectos]
  const fig = { tipo: "forma", figura: "rect", prst: "rect", radio: 0, sombra: false, x: 380, y: 190, w: 200, h: 140, color: "#2F6FB0", texto: "NOM" };
  const ef = (id, k, o, extra) => { let y = A.nuevoEfecto(id, k, "despues"); if (!y) return null; if (o) y = A.conOpciones(y, o); return Object.assign(y, { dur: 1000 }, extra || {}); };
  const presets = (NOM.animOfficeDatos || {}).presets || {};
  const hay = (k) => !!presets[k];
  const colorE = Object.keys(presets).find((k) => /^emph:/.test(k) && /color/i.test((presets[k].en || "") + (presets[k].n || "")));
  const pathCat = Object.keys(presets).filter((k) => k.startsWith("path:")).sort()[0];
  const lista = [
    ["Desvanecer", fig, (id) => [ef(id, "entr:10:0")]],
    ["Desplazar con rebote", fig, (id) => [ef(id, "entr:2:4", { rebote: 0.3 })]],
    ["Zoom con inicio y fin suaves", fig, (id) => [ef(id, "entr:53:16", { suaveIni: 0.4, suaveFin: 0.4 })]],
    ["Pulso tres veces", fig, (id) => [ef(id, "emph:26:0", { repetir: 3 }, { dur: 500 })]],
    ["Girar a la izquierda, ida y vuelta", fig, (id) => [ef(id, "emph:8:0", { giro: -360, autoRev: true }, { dur: 800 })]],
    ["Aumentar 150 % y rebobinar", fig, (id) => [ef(id, "emph:6:0", { escala: 150, rebobinar: true })]],
    colorE ? ["Color de énfasis", fig, (id) => [ef(id, colorE, { color: "#C2372B" })]] : null,
    ["Desvanecer y atenuar a un color", fig, (id) => [ef(id, "entr:10:0", { despues: { color: "#C8C8C8" } })]],
    ["Pulso y ocultar después", fig, (id) => [ef(id, "emph:26:0", { despues: "ocultar" })]],
    ["Letra por letra al revés", { tipo: "texto", x: 120, y: 200, w: 720, h: 120, html: "<p>Texto letra por letra al revés</p>" }, (id) => [ef(id, "entr:10:0", { texto: "letra", pct: 10, atras: true })]],
    ["Palabra por palabra", { tipo: "texto", x: 120, y: 200, w: 720, h: 120, html: "<p>Una palabra después de otra</p>" }, (id) => [ef(id, "entr:2:4", { texto: "palabra", pct: 30 })]],
    ["Por párrafos al revés", { tipo: "texto", x: 120, y: 120, w: 720, h: 300, html: "<p>Primer párrafo</p><p>Segundo párrafo</p><p>Tercer párrafo</p>" },
      (id) => A.porParrafos(ef(id, "entr:10:0"), [0, 1, 2], true, false).map((y, i) => Object.assign(y, { inicio: i ? "despues" : "despues", dur: 600 }))],
    ["Trayectoria dibujada abierta", Object.assign({}, fig, { x: 120, w: 140, h: 100 }), (id) => [Object.assign(A.trayectoriaPropia(id, [[0, 0], [120, -80], [260, -100], [400, -40], [560, 60]], T, false), { inicio: "despues", dur: 1500 })]],
    ["Trayectoria dibujada cerrada", Object.assign({}, fig, { x: 380, w: 140, h: 100 }), (id) => [Object.assign(A.trayectoriaPropia(id, [[0, 0], [120, -90], [240, 0], [120, 90], [2, 2]], T, true), { inicio: "despues", dur: 1500 })]],
    pathCat ? ["Trayectoria del catálogo", Object.assign({}, fig, { w: 140, h: 100 }), (id) => [ef(id, pathCat, null, { dur: 1500 })]] : null,
    ["Salida: desvanecer", fig, (id) => [ef(id, "exit:10:0")]],
  ].filter(Boolean);
  const paginas = lista.map((_, i) => "v" + (i + 1));
  doc.cambiar("_doc", { paginas });
  const info = [];
  lista.forEach(([nombre, base, armar], i) => {
    const p = paginas[i], id = "o" + (i + 1);
    doc.cambiar(id, Object.assign({ pagina: p, orden: 1, z: 1 }, base));
    // la marca: 12 cuadritos (9 bits del número de la diapositiva y 3 de comprobación, número módulo 7)
    const k = i + 1, v = (k << 3) | (k % 7);
    for (let b = 0; b < 12; b++) {
      const bit = (v >> (11 - b)) & 1;
      doc.cambiar(`m${k}_${b}`, { tipo: "forma", figura: "rect", pagina: p, orden: 10 + b, z: 10 + b, x: 6 + b * 12, y: 528, w: 10, h: 10, prst: "rect", radio: 0, sombra: false, bordeAncho: 0, color: bit ? "#000000" : "#FFFFFF" });
    }
    const efs = armar(id).filter(Boolean);
    doc.cambiar("_doc", { ["anim:" + p]: efs });
    const durTotal = efs.reduce((s, x) => {
      const o = A.opcionesDe(x) || {};
      const rep = typeof o.repetir === "number" ? o.repetir : 1;
      return s + ((+x.retraso || 0) + (+x.dur || 500) * rep * (o.autoRev ? 2 : 1)) / 1000;
    }, 0);
    info.push({ diapo: k, efecto: nombre, dur_total: Math.round((durTotal + 0.6) * 100) / 100, texto: base.tipo === "texto" });
  });
  return { pptx: aB64(await NOM.ooxml.pptx(doc, false)), info };
}
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sin-office", action="store_true")
    ap.add_argument("--cuadros", default="8")
    a = ap.parse_args()
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        nav = pw.chromium.launch(); pag = nav.new_page(); errores = []
        pag.on("pageerror", lambda e: errores.append(str(e)))
        pag.goto(F.BASE + "/")
        pag.wait_for_function("() => window.NOM && NOM.ooxml && NOM.formatos && NOM.animOffice && NOM.animOfficeDatos", timeout=60000)
        r = pag.evaluate(JS)
        nav.close()
    if errores: print("errores de JavaScript:", errores[:3])
    REAL.mkdir(parents=True, exist_ok=True)
    pptx = REAL / f"{NOMBRE}.pptx"
    pptx.write_bytes(base64.b64decode(r["pptx"]))
    (REAL / f"{NOMBRE}.json").write_text(json.dumps({"diapositivas": r["info"]}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(r['info'])} diapositivas en {pptx}")
    if a.sin_office: return 0
    import office_real as O
    subprocess.run(["scp", "-q", str(pptx), O.WIN + ":nomoffice/archivos/"], check=True)
    subprocess.run(["scp", "-q", str(Path(__file__).parent / "windows" / "animaciones_office.py"), O.WIN + ":nomoffice/"], check=True)
    O.correr("videopaneles", "animaciones_office.py", "video", NOMBRE)
    O.traer(f"{NOMBRE}.mp4")
    if not (REAL / f"{NOMBRE}.mp4").exists(): print("MAL: PowerPoint no hizo el video"); return 1
    env = dict(os.environ, PLAYWRIGHT_BROWSERS_PATH=str(F.CACHE / "navegadores"))
    r2 = subprocess.run([sys.executable, str(Path(__file__).parent / "animaciones_real.py"), str(pptx), "--cuadros", a.cuadros], env=env)
    return r2.returncode


if __name__ == "__main__":
    sys.exit(main())

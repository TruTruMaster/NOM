/* NOM-COLAB · Dibuja las geometrías predefinidas de Office (rectángulos, flechas, estrellas, llamadas, cintas…)
   exactamente como Office: evalúa las fórmulas de la norma (ECMA-376, 20.1.9) para el tamaño y los ajustes de
   cada forma y devuelve los trazos en SVG. Los datos están en geometrias-datos.js.

   NOM.geomOffice.trazos(nombre, w, h, ajustes) → [{ d, relleno: "norm"|"none"|"darken"|"lighten"…, contorno }]
   NOM.geomOffice.rectTexto(nombre, w, h, ajustes) → { l, t, r, b } donde va el texto
   NOM.geomOffice.tiene(nombre)
   NOM.geomOffice.trazosDe(definición, w, h) y rectTextoDe(…): lo mismo para una geometría libre de un archivo */
(() => {
  const G = () => NOM.GEOMETRIAS || {};
  const RAD = Math.PI / 180 / 60000;   // los ángulos van en 60 000avos de grado
  /** Variables que existen en toda forma (w, h, ss, cd4, wd2…). */
  function base(w, h) {
    const ss = Math.min(w, h), ls = Math.max(w, h);
    const v = { w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ss, ls,
      cd2: 10800000, cd4: 5400000, cd8: 2700000, "3cd4": 16200000, "3cd8": 8100000, "5cd8": 13500000, "7cd8": 18900000 };
    for (const k of [2, 3, 4, 5, 6, 8, 10, 12, 32]) { v["wd" + k] = w / k; v["hd" + k] = h / k; }
    for (const k of [2, 4, 6, 8, 16, 32]) v["ssd" + k] = ss / k;
    return v;
  }
  function evaluar(f, v) {
    const p = f.trim().split(/\s+/), op = p[0];
    const a = p.slice(1).map((x) => (x in v ? v[x] : +x));
    switch (op) {
      case "val": return a[0];
      case "*/": return (a[0] * a[1]) / a[2];
      case "+-": return a[0] + a[1] - a[2];
      case "+/": return (a[0] + a[1]) / a[2];
      case "?:": return a[0] > 0 ? a[1] : a[2];
      case "abs": return Math.abs(a[0]);
      case "at2": return Math.atan2(a[1], a[0]) / RAD;
      case "cat2": return a[0] * Math.cos(Math.atan2(a[2], a[1]));
      case "sat2": return a[0] * Math.sin(Math.atan2(a[2], a[1]));
      case "cos": return a[0] * Math.cos(a[1] * RAD);
      case "sin": return a[0] * Math.sin(a[1] * RAD);
      case "tan": return a[0] * Math.tan(a[1] * RAD);
      case "max": return Math.max(a[0], a[1]);
      case "min": return Math.min(a[0], a[1]);
      case "mod": return Math.sqrt(a[0] * a[0] + a[1] * a[1] + a[2] * a[2]);
      case "pin": return a[1] < a[0] ? a[0] : a[1] > a[2] ? a[2] : a[1];
      case "sqrt": return Math.sqrt(a[0]);
      default: return 0;
    }
  }
  /** Todas las guías de la forma resueltas (los ajustes del archivo reemplazan a los de omisión). */
  function variables(g, w, h, ajustes) {
    const v = base(w, h);
    for (const [n, f] of g.av || []) v[n] = ajustes && ajustes[n] != null ? +ajustes[n] : evaluar(f, v);
    // en un archivo, «adj» y «adj1» son el mismo ajuste en las formas de uno solo
    if (ajustes) for (const k in ajustes) if (!(k in v)) v[k] = +ajustes[k];
    for (const [n, f] of g.gd || []) v[n] = evaluar(f, v);
    return v;
  }
  const n2 = (x) => Math.round(x * 100) / 100;
  /** Un punto de la elipse según el ángulo VISUAL (como Office), no el paramétrico. */
  function enElipse(wR, hR, ang) {
    const a = ang * RAD, t = Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
    return [wR * Math.cos(t), hR * Math.sin(t)];
  }
  function trazos(nombre, w, h, ajustes) {
    const g = G()[nombre];
    return g ? trazosDe(g, w, h, ajustes) : null;
  }
  /** Los trazos de una geometría dada (una de la norma o una libre de un archivo: custGeom, con el mismo formato). */
  function trazosDe(g, w, h, ajustes) {
    const v = variables(g, w, h, ajustes);
    const val = (x) => (typeof x === "number" ? x : x in v ? v[x] : +x || 0);
    return (g.p || []).map((p) => {
      const kx = p.w ? w / p.w : 1, ky = p.h ? h / p.h : 1;
      let d = "", x = 0, y = 0, x0 = 0, y0 = 0;
      for (const c of p.c) {
        switch (c[0]) {
          case "M": x = val(c[1]) * kx; y = val(c[2]) * ky; x0 = x; y0 = y; d += `M${n2(x)},${n2(y)}`; break;
          case "L": x = val(c[1]) * kx; y = val(c[2]) * ky; d += `L${n2(x)},${n2(y)}`; break;
          case "Q": { const q = c.slice(1).map((u, i) => val(u) * (i % 2 ? ky : kx)); d += `Q${n2(q[0])},${n2(q[1])} ${n2(q[2])},${n2(q[3])}`; x = q[2]; y = q[3]; break; }
          case "C": { const q = c.slice(1).map((u, i) => val(u) * (i % 2 ? ky : kx)); d += `C${n2(q[0])},${n2(q[1])} ${n2(q[2])},${n2(q[3])} ${n2(q[4])},${n2(q[5])}`; x = q[4]; y = q[5]; break; }
          case "A": {
            const wR = val(c[1]) * kx, hR = val(c[2]) * ky, st = val(c[3]), sw = val(c[4]);
            if (!wR || !hR) break;
            const [sx, sy] = enElipse(wR, hR, st), cx = x - sx, cy = y - sy;
            // un arco de 360° se parte en dos (SVG no dibuja un arco que empieza y termina en el mismo punto)
            const partes = Math.abs(sw) >= 21600000 - 1 ? 2 : 1;
            for (let k = 1; k <= partes; k++) {
              const [ex, ey] = enElipse(wR, hR, st + (sw * k) / partes);
              const nx = cx + ex, ny = cy + ey;
              d += `A${n2(wR)},${n2(hR)} 0 ${Math.abs(sw / partes) > 10800000 ? 1 : 0} ${sw > 0 ? 1 : 0} ${n2(nx)},${n2(ny)}`;
              x = nx; y = ny;
            }
            break;
          }
          case "Z": d += "Z"; x = x0; y = y0; break;
        }
      }
      return { d, relleno: p.f || "norm", contorno: p.s !== 0 };
    });
  }
  function rectTexto(nombre, w, h, ajustes) {
    return rectTextoDe(G()[nombre], w, h, ajustes);
  }
  function rectTextoDe(g, w, h, ajustes) {
    if (!g || !g.rect) return { l: 0, t: 0, r: w, b: h };
    const v = variables(g, w, h, ajustes);
    const val = (x) => (typeof x === "number" ? x : x in v ? v[x] : +x || 0);
    return { l: val(g.rect[0]), t: val(g.rect[1]), r: val(g.rect[2]), b: val(g.rect[3]) };
  }
  /** Color de un trazo según su relleno (las caras de un cubo o de una cinta van más oscuras o más claras). */
  function matiz(hex, modo) {
    const m = /^#?([0-9a-f]{6})$/i.exec(hex || "");
    if (!m || modo === "norm") return hex;
    const n = parseInt(m[1], 16), c = [n >> 16, (n >> 8) & 255, n & 255];
    const k = { darken: 0.6, darkenLess: 0.8, lighten: 0.4, lightenLess: 0.2 }[modo];
    if (k == null) return hex;
    const r = modo.startsWith("darken") ? c.map((v) => v * k) : c.map((v) => v + (255 - v) * k);
    return "#" + r.map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, "0")).join("");
  }
  // Las tramas de Office (pattFill): 8 × 8 puntos, una fila por byte (el bit alto es el punto de la izquierda).
  // Medidas en PowerPoint real: cada punto mide 1/96 de pulgada sobre la diapositiva.
  const TRAMAS = { pct5: "8000000008000000", pct10: "8000080080000800", pct20: "8800220088002200", pct25: "8822882288228822", pct30: "AA44AA11AA44AA11", pct40: "AA55AA51AA55AA15", pct50: "AA55AA55AA55AA55", pct60: "EE55BB55EE55BB55", pct70: "77DD77DD77DD77DD", pct75: "77FFDDFF77FFDDFF", pct80: "EFFFFEFFEFFFFEFF", pct90: "FFFFFFF7FFFFFF7F", dkHorz: "FFFF0000FFFF0000", dkVert: "CCCCCCCCCCCCCCCC", dkDnDiag: "CC663399CC663399", dkUpDiag: "3366CC993366CC99", smCheck: "9966669999666699", trellis: "FF66FF99FF66FF99", ltHorz: "FF000000FF000000", ltVert: "8888888888888888", ltDnDiag: "8844221188442211", ltUpDiag: "1122448811224488", smGrid: "FF888888FF888888", dotDmnd: "8000220008002200", wdDnDiag: "C1E070381C0E0783", wdUpDiag: "83070E1C3870E0C1", dashUpDiag: "0000112244880000", dashDnDiag: "0000884422110000", narVert: "5555555555555555", narHorz: "FF00FF00FF00FF00", dashVert: "8080808008080808", dashHorz: "F00000000F000000", lgConfetti: "B130031BD8C00C8D", lgGrid: "FF80808080808080", horzBrick: "FF808080FF080808", lgCheck: "F0F0F0F00F0F0F0F", smConfetti: "8008400210012004", zigZag: "8142241881422418", solidDmnd: "10387CFE7C381000", diagBrick: "0102040818244281", openDmnd: "8244281028448201", plaid: "AA55AA55F0F0F0F0", sphere: "77898F8F7798F8F8", weave: "8854224588142251", dotGrid: "AA00800080008000", divot: "0010081000800180", shingle: "038448300C020101", wave: "001825C0001825C0", horz: "FF00000000000000", vert: "8080808080808080", dnDiag: "8040201008040201", upDiag: "0102040810204080", cross: "FF80808080808080", diagCross: "8142241818244281" };
  /** Los puntos encendidos de una trama como un trazo SVG (cada punto mide «p»), o null si no se conoce. */
  function trama(prst, p) {
    const hx = TRAMAS[prst];
    if (!hx) return null;
    let d = "";
    for (let f = 0; f < 8; f++) {
      const byte = parseInt(hx.slice(f * 2, f * 2 + 2), 16);
      for (let c = 0; c < 8; c++) if (byte & (0x80 >> c)) d += `M${n2(c * p)},${n2(f * p)}h${n2(p)}v${n2(p)}h${n2(-p)}Z`;
    }
    return d;
  }
  NOM.geomOffice = { trazos, rectTexto, trazosDe, rectTextoDe, matiz, trama, tiene: (n) => !!G()[n] };
})();

/* NOM · Dibuja las geometrías predefinidas de Office (rectángulos, flechas, estrellas, llamadas, cintas…)
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
  // ------------------------------------------------------------ dibujos de Word: formas y grupos como en Office
  const escX = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const HEX6 = /^#[0-9A-Fa-f]{6}$/;
  // Guiones de Office (prstDash), en múltiplos del grosor de la línea: trazo, hueco, trazo, hueco…
  const GUIONES = { dot: [1, 3], dash: [4, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3],
    sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1] };
  let nDibujo = 0;
  /** El relleno de Office como «servidor de pintura» de SVG (degradado, trama o imagen) para una caja de w × h. */
  function defRelleno(pt, id, w, h) {
    const op = (a) => (a != null && a < 1 ? ` stop-opacity="${n2(a)}"` : "");
    const stops = () => pt.paradas.map(([o, c, a]) => `<stop offset="${n2(Math.max(0, Math.min(1, +o)))}" stop-color="${escX(c)}"${op(a)}/>`).join("");
    if (pt.tipo === "trama") {
      return `<pattern id="${id}" patternUnits="userSpaceOnUse" x="0" y="0" width="8" height="8"><rect width="8" height="8" fill="${escX(pt.bg)}"${pt.bgA != null && pt.bgA < 1 ? ` fill-opacity="${n2(pt.bgA)}"` : ""}/><path d="${trama(pt.prst, 1) || ""}" fill="${escX(pt.fg)}"${pt.fgA != null && pt.fgA < 1 ? ` fill-opacity="${n2(pt.fgA)}"` : ""}/></pattern>`;
    }
    if (pt.tipo === "img") {
      // (con recorte: la parte que se ve de la imagen llena la forma)
      const r = Array.isArray(pt.recorte) ? pt.recorte : [0, 0, 0, 0], iw = w / Math.max(0.01, 1 - r[0] - r[2]), ih = h / Math.max(0.01, 1 - r[1] - r[3]);
      return `<pattern id="${id}" patternUnits="userSpaceOnUse" x="0" y="0" width="${n2(w)}" height="${n2(h)}"><image href="${escX(pt.src)}" x="${n2(-r[0] * iw)}" y="${n2(-r[1] * ih)}" width="${n2(iw)}" height="${n2(ih)}" preserveAspectRatio="none"${pt.alfa != null && pt.alfa < 1 ? ` opacity="${n2(pt.alfa)}"` : ""}/></pattern>`;
    }
    if (pt.tipo === "lin") {
      // el ángulo va en el sentido de las agujas del reloj y el degradado cubre toda la caja (como Office)
      const a = ((+pt.ang || 0) * Math.PI) / 180, cx = w / 2, cy = h / 2, L = (Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a))) / 2;
      return `<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${n2(cx - L * Math.cos(a))}" y1="${n2(cy - L * Math.sin(a))}" x2="${n2(cx + L * Math.cos(a))}" y2="${n2(cy + L * Math.sin(a))}">${stops()}</linearGradient>`;
    }
    const cx = (pt.cx == null ? 0.5 : +pt.cx) * w, cy = (pt.cy == null ? 0.5 : +pt.cy) * h;
    if (pt.camino === "rect" || pt.camino === "shape") {
      // desde el centro hacia los lados: cuatro cuñas, cada una con su degradado hacia su orilla
      const cuña = (k, x2, y2, pts) => `<linearGradient id="${id}-${k}" gradientUnits="userSpaceOnUse" x1="${n2(cx)}" y1="${n2(cy)}" x2="${n2(x2)}" y2="${n2(y2)}">${stops()}</linearGradient><path d="M${pts}Z" fill="url(#${id}-${k})"/>`;
      const c = `${n2(cx)},${n2(cy)}`;
      return `<pattern id="${id}" patternUnits="userSpaceOnUse" x="0" y="0" width="${n2(w)}" height="${n2(h)}">${cuña("a", cx, 0, `0,0 ${n2(w)},0 ${c}`)}${cuña("b", w, cy, `${n2(w)},0 ${n2(w)},${n2(h)} ${c}`)}${cuña("c", cx, h, `${n2(w)},${n2(h)} 0,${n2(h)} ${c}`)}${cuña("d", 0, cy, `0,${n2(h)} 0,0 ${c}`)}</pattern>`;
    }
    const r = Math.max(Math.hypot(cx, cy), Math.hypot(w - cx, cy), Math.hypot(cx, h - cy), Math.hypot(w - cx, h - cy));
    return `<radialGradient id="${id}" gradientUnits="userSpaceOnUse" cx="${n2(cx)}" cy="${n2(cy)}" r="${n2(r)}">${stops()}</radialGradient>`;
  }
  /** Punta de flecha de Office (headEnd/tailEnd) en (x, y), que viene desde (x0, y0). */
  function puntaOffice(pt, x, y, x0, y0, ancho, color, op) {
    if (!pt || !pt.t || pt.t === "none") return "";
    let dx = x - x0, dy = y - y0;
    const L = Math.hypot(dx, dy) || 1; dx /= L; dy /= L;
    const TAM = { sm: 2, med: 3, lg: 5 }, base = Math.max(ancho, 1.5), W = base * (TAM[pt.w] || 3), Ln = base * (TAM[pt.l] || 3), nx = (-dy * W) / 2, ny = (dx * W) / 2;
    const bx = x - dx * Ln, by = y - dy * Ln, c = escX(color);
    if (pt.t === "oval") return `<ellipse cx="${n2(x)}" cy="${n2(y)}" rx="${n2(Ln / 2)}" ry="${n2(W / 2)}" transform="rotate(${n2((Math.atan2(dy, dx) * 180) / Math.PI)} ${n2(x)} ${n2(y)})" fill="${c}"${op}/>`;
    if (pt.t === "diamond") return `<path d="M${n2(x + dx * Ln / 2)},${n2(y + dy * Ln / 2)} L${n2(x + nx)},${n2(y + ny)} L${n2(x - dx * Ln / 2)},${n2(y - dy * Ln / 2)} L${n2(x - nx)},${n2(y - ny)} Z" fill="${c}"${op}/>`;
    if (pt.t === "arrow") return `<path d="M${n2(bx + nx)},${n2(by + ny)} L${n2(x)},${n2(y)} L${n2(bx - nx)},${n2(by - ny)}" fill="none" stroke="${c}" stroke-width="${n2(ancho)}" stroke-linejoin="miter"${op.replace("fill-opacity", "stroke-opacity")}/>`;
    if (pt.t === "stealth") return `<path d="M${n2(x)},${n2(y)} L${n2(bx + nx)},${n2(by + ny)} L${n2(x - dx * Ln * 0.6)},${n2(y - dy * Ln * 0.6)} L${n2(bx - nx)},${n2(by - ny)} Z" fill="${c}"${op}/>`;
    return `<path d="M${n2(x)},${n2(y)} L${n2(bx + nx)},${n2(by + ny)} L${n2(bx - nx)},${n2(by - ny)} Z" fill="${c}"${op}/>`;
  }
  /** Filtro CSS de los efectos de Office: sombra exterior y resplandor (los bordes suaves van con un filtro SVG). */
  function filtroEfectos(f) {
    const rgba = (c, a) => { const m = /^#?([0-9a-f]{6})$/i.exec(c || ""); if (!m) return "rgba(0,0,0,.4)"; const n = parseInt(m[1], 16); return `rgba(${n >> 16},${(n >> 8) & 255},${n & 255},${n2(a == null ? 1 : a)})`; };
    const fs = [];
    if (f.brillo && +f.brillo.r > 0) { const r = +f.brillo.r, c = rgba(f.brillo.c, f.brillo.a); fs.push(`drop-shadow(0 0 ${n2(r / 2)}px ${c})`, `drop-shadow(0 0 ${n2(r / 2)}px ${c})`); }
    if (f.sombra) fs.push(`drop-shadow(${n2(+f.sombra.dx || 0)}px ${n2(+f.sombra.dy || 0)}px ${n2(+f.sombra.b || 0)}px ${rgba(f.sombra.c, f.sombra.a)})`);
    return fs.join(" ");
  }
  /**
   * Una forma de Office en SVG, como la dibuja Word: su geometría (predefinida o libre), su relleno (color con
   * transparencia, degradado lineal o desde el centro, trama o imagen), su contorno (grosor, color, guiones, puntas de
   * flecha) y sus efectos (sombra, resplandor, bordes suaves). f: la forma que lee el importador; w × h en px.
   */
  function svgForma(f, w, h) {
    const id = "nd" + ++nDibujo;
    w = Math.max(0.01, +w || 0.01); h = Math.max(0.01, +h || 0.01);
    const tz = (f.cust ? trazosDe(f.cust, w, h, f.av || null) : trazos(f.prst || "rect", w, h, f.av || null)) || [];
    const ln = f.linea && !f.linea.no ? f.linea : null, sinRelleno = f.relleno === "none" || (!f.pintura && !HEX6.test(f.relleno || ""));
    const pintura = f.pintura && !sinRelleno ? f.pintura : null;
    const llenar = (z) => (z.relleno === "none" || sinRelleno ? "none" : z.relleno === "norm" || !z.relleno ? (pintura ? `url(#${id})` : f.relleno) : matiz(pintura && pintura.paradas ? pintura.paradas[0][1] : f.relleno, z.relleno));
    const opR = f.alfa != null && f.alfa < 1 ? ` fill-opacity="${n2(f.alfa)}"` : "";
    const anchoPx = ln ? Math.max(0.5, (+ln.w || 0.75) * (96 / 72)) : 0, colL = ln && HEX6.test(ln.c || "") ? ln.c : "#000000";
    const opL = ln && ln.alfa != null && ln.alfa < 1 ? ` stroke-opacity="${n2(ln.alfa)}"` : "";
    const g = ln && ln.guion && GUIONES[ln.guion] ? GUIONES[ln.guion] : null;
    const tapa = ln && ln.tapa ? ln.tapa : "butt";
    // (con tapa redonda los puntos de Office son de un grosor: el trazo del guion se acorta para que no crezca)
    const guion = g ? ` stroke-dasharray="${g.map((v, i) => n2(Math.max(0.01, (i % 2 === 0 && tapa !== "butt" ? v - 1 : i % 2 && tapa !== "butt" ? v + 1 : v) * anchoPx))).join(" ")}"` : "";
    const trazoAttrs = ln ? ` stroke="${escX(colL)}" stroke-width="${n2(anchoPx)}"${guion} stroke-linecap="${tapa}" stroke-linejoin="round"${opL}` : ' stroke="none"';
    let cuerpo = tz.map((z) => `<path d="${z.d}" fill="${escX(llenar(z))}"${llenar(z) === "none" ? "" : opR}${z.contorno !== false ? trazoAttrs : ' stroke="none"'}/>`).join("");
    // puntas de flecha: al principio y al final del primer trazo abierto
    if (ln && ln.puntas && tz.length) {
      const q = ((tz.find((z) => z.contorno !== false) || tz[0]).d.match(/-?\d+(\.\d+)?(e-?\d+)?/g) || []).map(Number);
      if (q.length >= 4) cuerpo += puntaOffice(ln.puntas.ini, q[0], q[1], q[2], q[3], anchoPx, colL, opL.replace("stroke", "fill")) + puntaOffice(ln.puntas.fin, q[q.length - 2], q[q.length - 1], q[q.length - 4], q[q.length - 3], anchoPx, colL, opL.replace("stroke", "fill"));
    }
    const suave = +f.suave > 0 ? `<filter id="${id}-s" x="-10%" y="-10%" width="120%" height="120%"><feMorphology in="SourceAlpha" operator="erode" radius="${n2(+f.suave / 2)}" result="e"/><feGaussianBlur in="e" stdDeviation="${n2(+f.suave / 3)}" result="b"/><feComposite in="SourceGraphic" in2="b" operator="in"/></filter>` : "";
    const defs = pintura || suave ? `<defs>${pintura ? defRelleno(pintura, id, w, h) : ""}${suave}</defs>` : "";
    const fx = filtroEfectos(f);
    return `<svg width="${n2(w)}" height="${n2(h)}" style="position:absolute;inset:0;overflow:visible${fx ? ";filter:" + fx : ""}" aria-hidden="true">${defs}<g${suave ? ` filter="url(#${id}-s)"` : ""}>${cuerpo}</g></svg>`;
  }
  /** El texto de una forma (cuadro de texto): en el rectángulo de texto de su geometría, con sus márgenes y su letra. */
  function htmlTextoForma(f, w, h, familia) {
    if (!Array.isArray(f.texto)) return "";
    const rt = f.cust ? rectTextoDe(f.cust, w, h, f.av || null) : rectTexto(f.prst || "rect", Math.max(0.01, w), h, f.av || null);
    const ins = Array.isArray(f.ins) ? f.ins : [9.6, 4.8, 9.6, 4.8], le = f.letra || {};
    const fuente = le.f && familia ? `font-family:${escX(familia(le.f))};` : "";
    const vert = f.vert ? `writing-mode:vertical-rl;${f.vert === "vert270" ? "transform:rotate(180deg);" : ""}` : "";
    return `<span class="wd-forma-txt" style="${vert}left:${n2(rt.l)}px;top:${n2(rt.t)}px;width:${n2(Math.max(1, rt.r - rt.l))}px;height:${n2(Math.max(1, rt.b - rt.t))}px;padding:${n2(+ins[1] || 0)}px ${n2(+ins[2] || 0)}px ${n2(+ins[3] || 0)}px ${n2(+ins[0] || 0)}px;justify-content:${f.ancla === "ctr" ? "center" : f.ancla === "b" ? "flex-end" : "flex-start"};font-size:${+le.tam || 11}pt;color:${HEX6.test(le.c || "") ? le.c : "#000"};${le.b ? "font-weight:700;" : ""}${fuente}">${f.texto.map((q) => `<span style="display:block;text-align:${{ center: "center", right: "right", end: "right", both: "justify" }[q.al] || "left"}">${escX(q.t) || "&nbsp;"}</span>`).join("")}</span>`;
  }
  const transformDe = (o) => ((+o.giro || 0) || o.vH || o.vV ? `transform:rotate(${n2(+o.giro || 0)}deg) scale(${o.vH ? -1 : 1},${o.vV ? -1 : 1});` : "");
  /** Una imagen de un grupo: con su recorte, su borde y sus efectos. */
  function htmlImagen(im, w, h) {
    const r = Array.isArray(im.recorte) ? im.recorte : [0, 0, 0, 0], iw = w / Math.max(0.01, 1 - r[0] - r[2]), ih = h / Math.max(0.01, 1 - r[1] - r[3]);
    const ln = im.linea && !im.linea.no ? im.linea : null, fx = filtroEfectos(im);
    return `<span style="position:absolute;inset:0;overflow:hidden${ln ? `;outline:${n2((+ln.w || 0.75) * (96 / 72))}px solid ${HEX6.test(ln.c || "") ? ln.c : "#000"}` : ""}${fx ? ";filter:" + fx : ""}"><img src="${escX(im.src)}" alt="" style="position:absolute;left:${n2(-r[0] * iw)}px;top:${n2(-r[1] * ih)}px;width:${n2(iw)}px;height:${n2(ih)}px;max-width:none"></span>`;
  }
  /**
   * HTML de un dibujo de Word del tamaño w × h: una forma (con su texto) o un grupo/lienzo con sus piezas (formas,
   * imágenes y otros grupos, cada una en su lugar, girada o volteada). familia(nombre) → la fuente CSS de un tipo de letra.
   */
  function htmlDibujo(e, w, h, familia) {
    if (e.forma) return svgForma(e.forma, w, h) + htmlTextoForma(e.forma, w, h, familia);
    if (!e.grupo) return "";
    const gr = e.grupo, fondo = HEX6.test(gr.fondo || "") ? `background:${gr.fondo};` : "", borde = gr.linea && HEX6.test(gr.linea.c || "") ? `outline:${n2((+gr.linea.w || 0.75) * (96 / 72))}px solid ${gr.linea.c};outline-offset:-1px;` : "";
    const piezas = (gr.hijos || []).map((p) => {
      const pw = Math.max(0.01, +p.w || 0), ph = Math.max(0.01, +p.h || 0);
      const dentro = p.forma ? svgForma(p.forma, pw, ph) + htmlTextoForma(p.forma, pw, ph, familia) : p.img ? htmlImagen(p.img, pw, ph) : p.grupo ? htmlDibujo({ grupo: p.grupo }, pw, ph, familia) : "";
      return `<span class="wd-pieza" style="position:absolute;left:${n2(+p.x || 0)}px;top:${n2(+p.y || 0)}px;width:${n2(pw)}px;height:${n2(ph)}px;${transformDe(p)}">${dentro}</span>`;
    }).join("");
    return `<span class="wd-grupo" style="position:absolute;inset:0;${fondo}${borde}">${piezas}</span>`;
  }
  NOM.geomOffice = { trazos, rectTexto, trazosDe, rectTextoDe, matiz, trama, tiene: (n) => !!G()[n], svgForma, htmlDibujo, transformDe };
})();

/* NOM — animaciones de PowerPoint: el motor que las reproduce como PowerPoint y lo que las lee y escribe.
 *
 * Una animación de PowerPoint es un árbol de tiempos (ECMA-376, parte 1, §19.5): contenedores en paralelo con
 * su inicio, duración, repeticiones, autorreversa, aceleración, velocidad y «filtro de tiempo», y en las hojas los
 * comportamientos que cambian algo del objeto: set (poner un valor), anim (valores en el tiempo, con fórmulas como
 * «#ppt_y-sin(pi*$)/3»), animEffect (filtros: barrido, persianas, círculo, disolver…), animScale, animRot,
 * animMotion (trayectoria), animClr (colores) y cmd (medios). Con «iterate» el efecto va por letra o por palabra.
 *
 * NOM guarda la lista de efectos de cada diapositiva en "anim:<página>" (lo del panel de animaciones: bloque,
 * inicio, retraso y duración) y, en cada efecto que vino de PowerPoint o que se hizo con sus efectos, su árbol tal
 * cual ("ppt"). Al presentar, este motor arma con esa lista el árbol de la diapositiva igual que lo arma PowerPoint
 * (un grupo por clic; dentro, un tramo por cada «después de la anterior») y lo evalúa cuadro por cuadro sobre los
 * bloques pintados, sus párrafos, sus palabras y sus letras. Cada propiedad se compone como en SMIL: lo que empezó
 * después manda; lo aditivo se suma.
 *
 * Los nodos del árbol se guardan como [nombre, atributos, hijos] (el XML de PowerPoint como datos), con el destino
 * de cada comportamiento como el bloque de NOM ("bid") en lugar del número de la figura en el archivo.
 */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});

  // ================================================================== nodos [nombre, atributos, hijos]
  const loc = (q) => { const s = String(q || ""), i = s.indexOf(":"); return i < 0 ? s : s.slice(i + 1); };
  const A = (n) => (Array.isArray(n) && n[1]) || {};
  const H = (n) => (Array.isArray(n) && n[2]) || [];
  const hijo = (n, l) => H(n).find((c) => Array.isArray(c) && loc(c[0]) === l) || null;
  const hijosDe = (n, l) => H(n).filter((c) => Array.isArray(c) && (!l || loc(c[0]) === l));
  const textoDe = (n) => H(n).filter((c) => typeof c === "string").join("");
  /** Un elemento del XML (DOM) como nodo. */
  function deDom(el) {
    const a = {};
    for (const at of el.attributes) if (!/^xmlns/.test(at.name)) a[at.name] = at.value;
    const h = [];
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1) h.push(deDom(c));
      else if ((c.nodeType === 3 || c.nodeType === 4) && c.nodeValue.trim()) h.push(c.nodeValue.trim());
    }
    return h.length ? [el.tagName, a, h] : [el.tagName, a];
  }
  const escA = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  /** Un nodo como XML. cambia(nodo) puede devolver otro nodo (para poner el número de figura del archivo). */
  function aXml(n, cambia) {
    if (typeof n === "string") return escA(n);
    if (cambia) { const r = cambia(n); if (r === null) return ""; if (r) n = r; }
    const a = A(n), at = Object.keys(a).filter((k) => k[0] !== "_" && a[k] != null).map((k) => ` ${k}="${escA(a[k])}"`).join("");
    const h = H(n);
    return h.length ? `<${n[0]}${at}>${h.map((c) => aXml(c, cambia)).join("")}</${n[0]}>` : `<${n[0]}${at}/>`;
  }
  const copia = (n) => JSON.parse(JSON.stringify(n));

  // ================================================================== fórmulas de PowerPoint
  // «#ppt_x» es el valor de la figura sin animar; «ppt_x», el de ese momento (lo que dejaron las animaciones de antes);
  // «$», el valor que se está animando (en «fmla» de un tav). ppt_x/ppt_y: centro, en fracciones de la diapositiva.
  const FUNC = { abs: Math.abs, acos: Math.acos, asin: Math.asin, atan: Math.atan, ceil: Math.ceil, cos: Math.cos, cosh: Math.cosh,
    deg: (x) => (x * 180) / Math.PI, exp: Math.exp, floor: Math.floor, ln: Math.log, max: Math.max, min: Math.min, rad: (x) => (x * Math.PI) / 180,
    rand: (x) => azar() * (x == null ? 1 : x), sin: Math.sin, sinh: Math.sinh, sqrt: Math.sqrt, tan: Math.tan, tanh: Math.tanh };
  let semilla = 1;
  const azar = () => { semilla = (semilla * 16807) % 2147483647; return (semilla - 1) / 2147483646; };
  const cacheF = new Map();
  function formula(txt) {
    if (cacheF.has(txt)) return cacheF.get(txt);
    const s = String(txt == null ? "" : txt).replace(/\s+/g, "");
    let i = 0;
    const mal = () => { throw new Error("fórmula"); };
    function prim() {
      const c = s[i];
      if (c === "(") { i++; const e = expr(); if (s[i] !== ")") mal(); i++; return e; }
      if (c === "$") { i++; return (v) => v.$; }
      if (c && /[\d.]/.test(c)) {
        const m = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(s.slice(i)); if (!m) mal(); i += m[0].length;
        const n = parseFloat(m[0]); return () => n;
      }
      let base = false;
      if (c === "#") { base = true; i++; }
      const m = /^[a-z_][a-z0-9_]*/i.exec(s.slice(i)); if (!m) mal(); i += m[0].length;
      const id = m[0].toLowerCase();
      if (s[i] === "(" && FUNC[id]) {
        i++; const args = [];
        if (s[i] !== ")") { args.push(expr()); while (s[i] === ",") { i++; args.push(expr()); } }
        if (s[i] !== ")") mal(); i++;
        const fn = FUNC[id]; return (v) => fn(...args.map((a) => a(v)));
      }
      if (id === "pi") return () => Math.PI;
      if (id === "e") return () => Math.E;
      return base ? (v) => +(v.base[id] || 0) : (v) => +(v.act[id] != null ? v.act[id] : v.base[id] || 0);
    }
    function unario() { if (s[i] === "-") { i++; const a = unario(); return (v) => -a(v); } if (s[i] === "+") { i++; return unario(); } return potencia(); }
    function potencia() { const a = prim(); if (s[i] === "^") { i++; const b = unario(); return (v) => Math.pow(a(v), b(v)); } return a; }
    function termino() {
      let a = unario();
      while (s[i] === "*" || s[i] === "/" || s[i] === "%") {
        const op = s[i++], b = unario(), x = a;
        a = op === "*" ? (v) => x(v) * b(v) : op === "/" ? (v) => { const d = b(v); return d ? x(v) / d : 0; } : (v) => x(v) % b(v);
      }
      return a;
    }
    function expr() {
      let a = termino();
      while (s[i] === "+" || s[i] === "-") { const op = s[i++], b = termino(), x = a; a = op === "+" ? (v) => x(v) + b(v) : (v) => x(v) - b(v); }
      return a;
    }
    let f = null;
    try { f = expr(); if (i < s.length) f = null; } catch (e) { f = null; }
    cacheF.set(txt, f);
    return f;
  }
  const valorNum = (txt, vars) => { const f = formula(txt); if (!f) { const n = parseFloat(txt); return isFinite(n) ? n : 0; } const r = f(vars); return isFinite(r) ? r : 0; };

  // ================================================================== colores
  const hexRgb = (h) => { const m = /^#?([0-9a-f]{6})/i.exec(h || ""); if (!m) return null; const n = parseInt(m[1], 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
  const rgbHex = (c) => "#" + c.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();
  function rgbHsl([r, g, b]) {
    r /= 255; g /= 255; b /= 255;
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
    if (mx === mn) return [0, 0, l];
    const d = mx - mn, s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    const h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [h * 60, s, l];
  }
  function hslRgb([h, s, l]) {
    h = ((h % 360) + 360) % 360 / 360;
    if (!s) return [l * 255, l * 255, l * 255];
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const f = (t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
    return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
  }
  /** Color de un nodo de DrawingML (a:srgbClr, a:schemeClr…): el que resolvió la importación (_c) o el del tema. */
  function colorNodo(n, ctx) {
    if (!n) return null;
    if (loc(n[0]) === "clrVal" || loc(n[0]) === "to" || loc(n[0]) === "from" || loc(n[0]) === "val") n = hijosDe(n)[0] || null;
    if (!n) return null;
    if (A(n)._c) return A(n)._c;
    const l = loc(n[0]), a = A(n);
    let hex = null;
    if (l === "srgbClr") hex = "#" + a.val;
    else if (l === "sysClr") hex = "#" + (a.lastClr || "000000");
    else if (l === "prstClr") hex = { black: "#000000", white: "#FFFFFF", red: "#FF0000", green: "#008000", blue: "#0000FF", yellow: "#FFFF00" }[a.val] || "#000000";
    else if (l === "schemeClr") hex = ctx && ctx.colorTema ? ctx.colorTema(a.val) : null;
    else if (l === "scrgbClr") hex = rgbHex([+a.r, +a.g, +a.b].map((v) => (v / 100000) * 255));
    else if (l === "hslClr") hex = rgbHex(hslRgb([(+a.hue || 0) / 60000, (+a.sat || 0) / 100000, (+a.lum || 0) / 100000]));
    if (!hex) return null;
    // modificadores comunes (lumMod, lumOff, tint, shade, alpha se ignora)
    let c = hexRgb(hex);
    for (const m of hijosDe(n)) {
      const v = +A(m).val / 100000, ml = loc(m[0]);
      if (ml === "lumMod" || ml === "lumOff") { const hs = rgbHsl(c); hs[2] = ml === "lumMod" ? hs[2] * v : hs[2] + v; hs[2] = Math.max(0, Math.min(1, hs[2])); c = hslRgb(hs); }
      else if (ml === "tint") c = c.map((x) => x + (255 - x) * (1 - v));
      else if (ml === "shade") c = c.map((x) => x * v);
    }
    return rgbHex(c);
  }

  // ================================================================== tiempos (SMIL)
  /** Los tiempos de un cTn: inicio (ms, relativo al padre), duración simple, repeticiones, autorreversa, aceleración… f: escala de tiempos. */
  function tiempoDe(ctn, f) {
    const a = A(ctn);
    let b = 0;
    const st = hijo(ctn, "stCondLst");
    if (st) for (const c of hijosDe(st, "cond")) { const d = A(c).delay; if (d != null && d !== "indefinite" && !A(c).evt) { b = (+d || 0) * f; break; } if (d != null && d !== "indefinite") { b = (+d || 0) * f; break; } }
    const d = a.dur == null ? null : a.dur === "indefinite" ? Infinity : Math.max(0, +a.dur || 0) * f;
    const rep = a.repeatCount == null ? 1 : a.repeatCount === "indefinite" ? Infinity : Math.max(0, +a.repeatCount || 0) / 1000;
    const repDur = a.repeatDur == null ? null : a.repeatDur === "indefinite" ? Infinity : Math.max(0, +a.repeatDur || 0) * f;
    const fin = hijo(ctn, "endCondLst");
    return {
      b, d, rep: rep || 1, repDur, rev: a.autoRev === "1" || a.autoRev === "true", acc: Math.max(0, Math.min(1, (+a.accel || 0) / 100000)), dec: Math.max(0, Math.min(1, (+a.decel || 0) / 100000)),
      spd: a.spd ? Math.max(0.01, Math.abs(+a.spd) / 100000) : 1,
      fill: a.fill || (a.dur == null && a.repeatCount == null && !fin ? "hold" : "remove"),
      tf: a.tmFilter ? leerTmFilter(a.tmFilter) : null,
      hastaClic: !!(fin && hijosDe(fin, "cond").some((c) => A(c).evt === "onNext")),
    };
  }
  function leerTmFilter(s) {
    const pts = String(s).split(";").map((p) => p.split(",").map((x) => +x.trim())).filter((p) => p.length === 2 && isFinite(p[0]) && isFinite(p[1]));
    return pts.length >= 2 ? pts : null;
  }
  const porFiltro = (tf, x) => {
    if (x <= tf[0][0]) return tf[0][1];
    for (let i = 1; i < tf.length; i++) if (x <= tf[i][0]) { const [a, b] = [tf[i - 1], tf[i]]; return b[0] === a[0] ? b[1] : a[1] + ((b[1] - a[1]) * (x - a[0])) / (b[0] - a[0]); }
    return tf[tf.length - 1][1];
  };
  /**
   * Inicio y fin suaves (accel/decel) como los hace PowerPoint, medido en sus videos (a04-aceleracion, 16 combinaciones;
   * error medio 0.12 % del recorrido): como en SMIL, la velocidad máxima es R = 1/(1 − a/2 − d/2) y en medio va
   * pareja, pero la velocidad no sube en línea recta sino suave (3u² − 2u³), así que el recorrido en la subida es
   * R·a·(u³ − u⁴/2), u = t/a (y lo mismo, al revés, en la bajada). La aceleración constante de SMIL se aleja hasta 4 %.
   */
  function acelera(u, D, a, d) {
    if (!(a || d) || !(D > 0) || !isFinite(D)) return u;
    if (a + d > 1) { const k = 1 / (a + d); a *= k; d *= k; }
    const R = 1 / (1 - a / 2 - d / 2), s = Math.max(0, Math.min(1, u / D));
    let p;
    if (a > 0 && s < a) { const q = s / a; p = R * a * (q * q * q - (q * q * q * q) / 2); }
    else if (d > 0 && s > 1 - d) { const q = (1 - s) / d; p = 1 - R * d * (q * q * q - (q * q * q * q) / 2); }
    else p = R * (s - a / 2);
    return p * D;
  }
  /**
   * «Rebote al final» (p14:bounceEnd = b): llega al final en 1 − b a velocidad pareja y rebota como un resorte
   * amortiguado; el tiempo del rebote crece con √b. Medido en los videos de PowerPoint (b = 0.25, 0.5 y 0.75):
   * error medio 0.17 % del recorrido. Devuelve el avance (puede pasar de 1).
   */
  function rebote(s, b) {
    if (!(b > 0)) return s;
    b = Math.min(0.95, b);
    const T1 = 1 - b, v = 1 / T1;
    if (s <= T1) return s * v;
    const e = Math.sqrt(b), tau = (s - T1) / e, K = 6.25, W = 15.75;
    if (s >= 1) return 1;
    return 1 + ((v * e) / W) * Math.exp(-K * tau) * Math.sin(W * tau);
  }
  /** Duración activa de un nodo en el tiempo de su padre. */
  function activa(t) {
    if (t.d == null) return 0;
    const per = t.d * (t.rev ? 2 : 1);
    if (!isFinite(per)) return Infinity;
    const ad = t.repDur != null ? t.repDur : per * t.rep;
    return ad / t.spd;
  }
  /** Tiempo simple de un nodo dado el tiempo simple de su padre: null si no ha empezado o ya se quitó. */
  function local(t, tp, inclusivo) {
    if (tp == null) return null;
    let ta = (tp - t.b) * t.spd;
    if (ta < 0) return null;
    const AD = t.AD * t.spd;
    let fin = false;
    // (inclusivo: justo al terminar todavía cuenta, aunque se quite: así lo ve lo que empieza en ese instante)
    if (ta >= AD) { if (t.fill === "remove" && !(t.hastaClic && !isFinite(AD)) && !(inclusivo && ta - AD < 0.5)) return null; ta = AD; fin = true; }
    const D = t.d;
    if (!(D > 0)) return 0;
    if (!isFinite(D)) return ta;
    const per = D * (t.rev ? 2 : 1);
    let k = Math.floor(ta / per), u = ta - k * per;
    if (fin && u < 1e-9 && ta > 0) u = per;
    if (t.rev && u > D) u = 2 * D - u;
    u = Math.max(0, Math.min(D, u));
    u = acelera(u, D, t.acc, t.dec);
    if (t.tf) u = D * porFiltro(t.tf, u / D);
    return u;
  }

  // ================================================================== filtros (animEffect)
  // Medido contra los videos de PowerPoint: cada filtro descubre la figura en un orden, un «campo» D(x, y) entre 0 y 1
  // (qué tan tarde se descubre cada punto), con un borde suave lineal de ancho G (en unidades de D): en el avance s,
  // alfa = recortar((s·(1+G) − D) / G). G = 0.145 en todos los filtros (barrido, persianas, cuadros, división, cuadro,
  // círculo, rombo, más); el borde medido en los videos coincide al píxel. La salida ("out") es la entrada al revés
  // en los filtros por franjas (persianas, cuadros bicolores, barras diagonales, aleatorios) y lo contrario de la
  // entrada en los demás (barrido, cuadro, círculo, rombo, más, dividir, cuña, rueda).
  const G_SUAVE = 0.145;
  const REVES = new Set(["blinds", "checkerboard", "strips", "randombar", "dissolve", "slide"]);
  /** El campo D de un filtro en una cuadrícula de mw×mh (cada celda, un punto de la figura de w×h). */
  function campoFiltro(tipo, arg, mw, mh, w, h, semillaEl) {
    const D = new Float32Array(mw * mh);
    const cx = w / 2, cy = h / 2, sx = w / mw, sy = h / mh;
    const n6 = 6;
    let orden = null;
    if (tipo === "dissolve") orden = ordenAzar(mw * mh, semillaEl || 11);
    if (tipo === "randombar") orden = ordenAzar(arg === "vertical" ? mw : mh, semillaEl || 7);
    // barras diagonales: celdas de 10 pt (como en PowerPoint), cada diagonal un paso
    const celda = 10 * 4 / 3 * 0.75;
    const ncx = Math.max(1, Math.ceil(w / celda)), ncy = Math.max(1, Math.ceil(h / celda));
    for (let j = 0; j < mh; j++) for (let i = 0; i < mw; i++) {
      const x = (i + 0.5) * sx, y = (j + 0.5) * sy, dx = x - cx, dy = y - cy;
      let d = 0;
      switch (tipo) {
        // «wipe(down)» descubre desde abajo; «wipe(up)», desde arriba (así lo escribe PowerPoint: «Desde abajo» = down)
        case "wipe": d = arg === "up" ? y / h : arg === "left" ? x / w : arg === "right" ? 1 - x / w : 1 - y / h; break;
        case "blinds": d = arg === "vertical" ? (x * n6 / w) % 1 : (y * n6 / h) % 1; break;
        case "checkerboard": {
          if (arg === "down") { const c = Math.min(n6 - 1, Math.floor((x * n6) / w)), P = h / n6, off = c % 2 ? P / 2 : 0; d = (((y - off) / P) % 1 + 1) % 1; }
          else { const r = Math.min(n6 - 1, Math.floor((y * n6) / h)), P = w / n6, off = r % 2 ? P / 2 : 0; d = (((x - off) / P) % 1 + 1) % 1; }
          break;
        }
        case "barn": { const vert = /vertical/i.test(arg), q = vert ? Math.abs(dx) / (w / 2) : Math.abs(dy) / (h / 2); d = /^out/i.test(arg) ? q : 1 - q; break; }
        case "box": { const q = Math.max(Math.abs(dx) / (w / 2), Math.abs(dy) / (h / 2)); d = arg === "out" ? q : 1 - q; break; }
        case "circle": { const q = Math.hypot(dx / (w / 2), dy / (h / 2)) / Math.SQRT2; d = arg === "out" ? q : 1 - q; break; }
        case "diamond": { const q = (Math.abs(dx) / (w / 2) + Math.abs(dy) / (h / 2)) / 2; d = arg === "out" ? q : 1 - q; break; }
        case "plus": { const q = Math.min(Math.abs(dx) / (w / 2), Math.abs(dy) / (h / 2)); d = arg === "out" ? q : 1 - q; break; }
        case "wedge": { const a = Math.abs((Math.atan2(dx, -dy) * 180) / Math.PI); d = a / 180; break; }
        case "wheel": { const n = Math.max(1, Math.min(8, +arg || 1)); const a = ((Math.atan2(dx, -dy) * 180) / Math.PI + 360) % 360; d = ((a * n) / 360) % 1; break; }
        case "strips": {
          const a = String(arg || "").toLowerCase();
          const ci = Math.min(ncx - 1, Math.floor(x / celda)), cj = Math.min(ncy - 1, Math.floor(y / celda));
          const ii = /left$/.test(a) ? ncx - 1 - ci : ci, jj = /^down/.test(a) ? cj : ncy - 1 - cj;
          d = (ii + jj) / Math.max(1, ncx + ncy - 2);
          break;
        }
        case "randombar": d = arg === "vertical" ? orden[i] / Math.max(1, mw - 1) : orden[j] / Math.max(1, mh - 1); break;
        case "dissolve": d = orden[j * mw + i] / Math.max(1, mw * mh - 1); break;
        default: d = 0;
      }
      D[j * mw + i] = d;
    }
    return D;
  }
  const cacheCampo = new Map();
  let lienzoMasc = null;
  /**
   * Lo que se ve de un elemento de w×h con el filtro f en el avance s del comportamiento (transicion "in" u "out").
   * Devuelve { opacidad } (desvanecer), { mascara: "url(data:…)" } o null (todo a la vista).
   */
  function mascara(filtro, s, w, h, pr, semillaEl, trans) {
    const m = /^(\w+)(?:\(([^)]*)\))?$/.exec(String(filtro || "").trim());
    if (!m) return null;
    const tipo = m[1].toLowerCase(), arg = (m[2] || "").trim();
    const sale = trans === "out";
    s = Math.max(0, Math.min(1, s));
    if (tipo === "fade" || tipo === "pixelate") return { opacidad: sale ? 1 - s : s };
    if (tipo === "image") return null;
    if (tipo === "slide") return { desliza: arg, p: sale ? 1 - s : s };
    if (!/^(wipe|blinds|checkerboard|barn|box|circle|diamond|plus|wedge|wheel|strips|randombar|dissolve)$/.test(tipo)) return { opacidad: sale ? 1 - s : s };
    const reves = REVES.has(tipo);
    const p = sale && reves ? 1 - s : s;            // avance del «descubrir»
    const contrario = sale && !reves;               // se ve lo que no se ha descubierto
    if (p >= 1) return contrario ? { nada: true } : null;
    if (p <= 0) return contrario ? null : { nada: true };
    const gs = pr && /gradientSize\s*:\s*([\d.]+)/i.exec(pr);
    const nitido = tipo === "dissolve" || tipo === "randombar";
    const G = nitido ? 1e-4 : gs ? Math.max(1e-4, +gs[1]) : G_SUAVE;
    // resolución de la máscara: 1 punto por px de la diapositiva (menos en figuras grandes con borde suave)
    const tope = nitido ? 260000 : 70000, k = Math.min(1, Math.sqrt(tope / Math.max(1, w * h)));
    const mw = Math.max(2, Math.round(w * k)), mh = Math.max(2, Math.round(h * k));
    const clave = `${tipo}(${arg})${mw}x${mh}:${Math.round(w)}x${Math.round(h)}:${semillaEl || 0}`;
    let D = cacheCampo.get(clave);
    if (!D) { D = campoFiltro(tipo, arg, mw, mh, w, h, semillaEl); if (cacheCampo.size > 200) cacheCampo.clear(); cacheCampo.set(clave, D); }
    if (typeof document === "undefined") return null;
    if (!lienzoMasc) lienzoMasc = document.createElement("canvas");
    lienzoMasc.width = mw; lienzoMasc.height = mh;
    const cx2 = lienzoMasc.getContext("2d"), img = cx2.createImageData(mw, mh), px = img.data;
    const P = p * (1 + G);
    for (let i = 0, n = mw * mh; i < n; i++) {
      let a = (P - D[i]) / G;
      a = a <= 0 ? 0 : a >= 1 ? 1 : a;
      if (contrario) a = 1 - a;
      const o = i * 4;
      px[o] = px[o + 1] = px[o + 2] = 255; px[o + 3] = Math.round(a * 255);
    }
    cx2.putImageData(img, 0, 0);
    return { mascara: `url("${lienzoMasc.toDataURL()}")` };
  }
  const cacheAzar = new Map();
  function ordenAzar(n, s) {
    const k = n + ":" + s;
    if (cacheAzar.has(k)) return cacheAzar.get(k);
    const o = Array.from({ length: n }, (_, i) => i);
    let x = s * 9301 + 49297;
    for (let i = n - 1; i > 0; i--) { x = (x * 16807) % 2147483647; const j = x % (i + 1); [o[i], o[j]] = [o[j], o[i]]; }
    cacheAzar.set(k, o);
    return o;
  }

  const n2 = (v) => Math.round(v * 100) / 100;

  // ================================================================== trayectorias (animMotion)
  /** Camino de PowerPoint («M 0 0 L 0.25 0 C … E») como lista de tramos en fracciones de la diapositiva. */
  function leerCamino(d) {
    const t = String(d || "").match(/[MLCZEmlcze]|-?\d*\.?\d+(?:e[+-]?\d+)?/g) || [];
    const tramos = [];
    let i = 0, cmd = "M", x = 0, y = 0, x0 = 0, y0 = 0;
    const num = () => +t[i++];
    while (i < t.length) {
      if (/[A-Za-z]/.test(t[i])) cmd = t[i++];
      const rel = cmd === cmd.toLowerCase(), C = cmd.toUpperCase();
      if (C === "E") break;
      if (C === "Z") { if (x !== x0 || y !== y0) tramos.push({ t: "L", p: [[x, y], [x0, y0]] }); x = x0; y = y0; if (i < t.length && !/[A-Za-z]/.test(t[i])) i++; continue; }
      if (i >= t.length || /[A-Za-z]/.test(t[i])) continue;
      if (C === "M") { const nx = num(), ny = num(); x = rel ? x + nx : nx; y = rel ? y + ny : ny; x0 = x; y0 = y; cmd = rel ? "l" : "L"; continue; }
      if (C === "L") { const nx = num(), ny = num(); const X = rel ? x + nx : nx, Y = rel ? y + ny : ny; tramos.push({ t: "L", p: [[x, y], [X, Y]] }); x = X; y = Y; continue; }
      if (C === "C") {
        const a = [num(), num()], b = [num(), num()], c = [num(), num()];
        const P1 = rel ? [x + a[0], y + a[1]] : a, P2 = rel ? [x + b[0], y + b[1]] : b, P3 = rel ? [x + c[0], y + c[1]] : c;
        tramos.push({ t: "C", p: [[x, y], P1, P2, P3] }); x = P3[0]; y = P3[1]; continue;
      }
      i++;
    }
    // tabla de largo (para recorrerlo a velocidad pareja)
    const pts = [[tramos.length ? tramos[0].p[0][0] : 0, tramos.length ? tramos[0].p[0][1] : 0]];
    for (const tr of tramos) {
      if (tr.t === "L") pts.push(tr.p[1]);
      else for (let k = 1; k <= 24; k++) { const s = k / 24, u = 1 - s, [p0, p1, p2, p3] = tr.p; pts.push([u * u * u * p0[0] + 3 * u * u * s * p1[0] + 3 * u * s * s * p2[0] + s * s * s * p3[0], u * u * u * p0[1] + 3 * u * u * s * p1[1] + 3 * u * s * s * p2[1] + s * s * s * p3[1]]); }
    }
    const lar = [0];
    for (let k = 1; k < pts.length; k++) lar.push(lar[k - 1] + Math.hypot(pts[k][0] - pts[k - 1][0], pts[k][1] - pts[k - 1][1]));
    return { pts, lar, total: lar[lar.length - 1] };
  }
  /** Punto del camino al avance s (0..1), a velocidad pareja sobre el largo. kx, ky: proporción de la diapositiva para medir el largo. */
  function puntoCamino(c, s) {
    if (!c.pts.length) return [0, 0];
    if (!(c.total > 0)) return c.pts[0];
    if (s > 1 && c.pts.length > 1) {
      // pasado el final (rebote): sigue la dirección del último tramo
      const n = c.pts.length, a = c.pts[n - 2], b = c.pts[n - 1], l = c.lar[n - 1] - c.lar[n - 2] || 1, k = ((s - 1) * c.total) / l;
      return [b[0] + (b[0] - a[0]) * k, b[1] + (b[1] - a[1]) * k];
    }
    const L = Math.max(0, Math.min(1, s)) * c.total;
    let k = 1;
    while (k < c.lar.length - 1 && c.lar[k] < L) k++;
    const a = c.lar[k - 1], b = c.lar[k], f = b > a ? (L - a) / (b - a) : 0;
    return [c.pts[k - 1][0] + (c.pts[k][0] - c.pts[k - 1][0]) * f, c.pts[k - 1][1] + (c.pts[k][1] - c.pts[k - 1][1]) * f];
  }

  // ================================================================== destinos
  const ESTILO_ID = "nom-anim-office";
  function ponerEstilo() {
    if (typeof document === "undefined" || document.getElementById(ESTILO_ID)) return;
    const st = document.createElement("style");
    st.id = ESTILO_ID;
    st.setAttribute("data-nom", "");   // (nucleo.js quita las hojas de estilo que no son de NOM)
    st.textContent = `.ao-color,.ao-color *{color:var(--ao-color)!important;-webkit-text-fill-color:var(--ao-color)!important}
.ao-fw,.ao-fw *{font-weight:var(--ao-fw)!important}.ao-fs,.ao-fs *{font-style:var(--ao-fs)!important}
.ao-td,.ao-td *{text-decoration-line:var(--ao-td)!important}
.ao-u{visibility:hidden}.ao-capa{position:absolute;left:0;top:0;width:0;height:0;overflow:visible;pointer-events:none;z-index:2}
.ao-capa>span{position:absolute;white-space:pre;display:inline-block;transform-origin:50% 50%}`;
    document.head.append(st);
  }
  /** Párrafos de un texto pintado (div y li que no tienen otros adentro), en orden. */
  function parrafosDe(cont) {
    if (!cont) return [];
    return [...cont.querySelectorAll("div, li, p")].filter((e) => !e.querySelector("div, li, p") && !e.closest(".ao-capa"));
  }
  const contTexto = (bl) => bl && (bl.querySelector(".txt.ppt, .libre-texto, .txt") || null);

  // ================================================================== el reproductor de una diapositiva
  /**
   * raiz: el elemento de la diapositiva con sus bloques (.bl[data-id]) ya pintados.
   * ctx: { lista (la de "anim:<página>"), T: {w, h} (tamaño de la diapositiva en px), bloque(id) → datos del bloque,
   *        escala() → cuántos px de pantalla mide un px de la diapositiva, colorTema(nombre) → "#RRGGBB",
   *        preset(item) → el árbol de PowerPoint de un efecto propio de NOM, alMedio(bid, cmd) }
   */
  function reproductor(raiz, ctx) {
    ponerEstilo();
    const T = ctx.T || { w: 960, h: 540 };
    const lista = (ctx.lista || []).filter((x) => x && (x.ppt || (ctx.preset && ctx.preset(x))));
    const destinos = new Map();          // clave → destino { clave, tipo, els, caja, ... }
    const porBloque = new Map();         // id → elementos .bl
    raiz.querySelectorAll(".bl[data-id]").forEach((el) => porBloque.set(el.dataset.id, el));
    const k0 = () => (ctx.escala ? ctx.escala() : 1) || 1;
    const rr = () => raiz.getBoundingClientRect();

    // ---- geometría base: en px de la diapositiva, sin animar
    function cajaBloques(els) {
      let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      els.forEach((el) => { const x = el.offsetLeft, y = el.offsetTop, w = el.offsetWidth, h = el.offsetHeight; x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x + w); y1 = Math.max(y1, y + h); });
      return isFinite(x0) ? { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } : { x: 0, y: 0, w: 0, h: 0 };
    }
    /** Caja de un elemento dentro de su bloque (px de la diapositiva), sin las transformaciones. */
    function cajaDentro(el, bl) {
      let x = 0, y = 0, e = el;
      while (e && e !== bl) { x += e.offsetLeft; y += e.offsetTop; e = e.offsetParent; if (e && !bl.contains(e) && e !== bl) break; }
      return { x: bl.offsetLeft + x, y: bl.offsetTop + y, w: el.offsetWidth, h: el.offsetHeight };
    }
    function bloquesDe(x, bid) {
      const ids = bid === x.id && Array.isArray(x.ids) && x.ids.length ? x.ids : [bid];
      return ids.map((i) => porBloque.get(i)).filter(Boolean);
    }
    /** El destino de un tgtEl: el bloque (o el grupo), su fondo, unos párrafos, o las letras/palabras (iterate). */
    function destino(x, tgt, iter) {
      const sp = tgt && hijo(tgt, "spTgt");
      if (!sp) return null;
      const bid = A(sp).bid && A(sp).bid !== "@" ? A(sp).bid : x.id;
      const els = bloquesDe(x, bid);
      if (!els.length) return null;
      const tx = hijo(sp, "txEl"), bg = hijo(sp, "bg");
      let clave = "b:" + bid, tipo = "bloque", sub = null;
      if (bg) { clave = "g:" + bid; tipo = "fondo"; }
      if (tx) {
        const pr = hijo(tx, "pRg"), cr = hijo(tx, "charRg");
        if (pr) { sub = { p: [+A(pr).st || 0, +A(pr).end || 0] }; clave = `p:${bid}:${sub.p[0]}-${sub.p[1]}`; tipo = "parrafo"; }
        else if (cr) { sub = { c: [+A(cr).st || 0, +A(cr).end || 0] }; clave = `c:${bid}:${sub.c[0]}-${sub.c[1]}`; tipo = "letras"; }
      }
      if (iter) clave += "|" + iter.type;
      if (destinos.has(clave)) return destinos.get(clave);
      const d = { clave, tipo, bid, els, sub, iter: iter ? iter.type : null };
      if (tipo === "bloque") d.caja = cajaBloques(els);
      else if (tipo === "fondo") { d.els = els.map((e) => e.querySelector(".forma-svg") || e); d.caja = cajaBloques(els); d.dentro = true; }
      else if (tipo === "parrafo") {
        const ps = parrafosDe(contTexto(els[0]));
        d.els = ps.slice(sub.p[0], sub.p[1] + 1);
        if (!d.els.length) return null;
        const cs = d.els.map((e) => cajaDentro(e, els[0]));
        const x0 = Math.min(...cs.map((c) => c.x)), y0 = Math.min(...cs.map((c) => c.y));
        d.caja = { x: x0, y: y0, w: Math.max(...cs.map((c) => c.x + c.w)) - x0, h: Math.max(...cs.map((c) => c.y + c.h)) - y0 };
        d.dentro = true;
      } else d.caja = cajaBloques(els);
      d.blq = els[0];
      if (iter) {
        d.unidades = unidades(d, iter.type);
        // por letra en una figura sin texto: PowerPoint anima la figura entera
        if (!d.unidades.length) { const sin = destino(x, tgt, null); destinos.set(clave, sin); return sin; }
      }
      destinos.set(clave, d);
      return d;
    }
    /** Letras o palabras de un destino: cada una con su caja (medida en el texto pintado) y su copia en una capa encima. */
    function unidades(d, tipo) {
      const bl = d.blq, cont = contTexto(bl);
      if (!cont) return [];
      const raices = d.tipo === "parrafo" ? d.els : [cont];
      const k = k0(), rb = bl.getBoundingClientRect();
      const giro = bl.style.transform; bl.style.transform = "none";
      const r0 = bl.getBoundingClientRect();
      const us = [];
      const rango = document.createRange();
      for (const raizT of raices) {
        const tw = document.createTreeWalker(raizT, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement && n.parentElement.closest(".ao-capa") ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
        const nodos = [];
        while (tw.nextNode()) nodos.push(tw.currentNode);
        for (const nd of nodos) {
          const s = nd.nodeValue;
          const trozos = [];
          if (tipo === "wd") { const re = /\S+\s*/g; let m; while ((m = re.exec(s))) trozos.push([m.index, m.index + m[0].length]); }
          // (PowerPoint no cuenta los espacios como letras: la letra después de un espacio no espera un turno)
          else for (let i = 0; i < s.length; i++) { if (/\s/.test(s[i])) continue; trozos.push([i, i + (/[\uD800-\uDBFF]/.test(s[i]) ? 2 : 1)]); if (/[\uD800-\uDBFF]/.test(s[i])) i++; }
          const cs = getComputedStyle(nd.parentElement);
          // la caja de cada letra según PowerPoint (medido con a05-letras: Arial, Calibri y Times de 18 a 72 pt, interlineado
          // de 1 a 2): el centro a 0.6 del alto del renglón desde arriba y el alto 2.02 × tamaño + 5.2 pt, sea cual sea la letra
          const fs = parseFloat(cs.fontSize) || 16;
          let lh = parseFloat(cs.lineHeight); if (!isFinite(lh)) lh = fs * 1.2;
          for (const [a, b] of trozos) {
            rango.setStart(nd, a); rango.setEnd(nd, b);
            const rs = [...rango.getClientRects()].filter((r) => r.width > 0 || r.height > 0);
            if (!rs.length) continue;
            const r = rs[0];
            const pos = { x: bl.offsetLeft + (r.left - r0.left) / k, y: bl.offsetTop + (r.top - r0.top) / k, w: r.width / k, h: r.height / k };
            const arribaRenglon = pos.y + pos.h / 2 - lh / 2, cy = arribaRenglon + 0.6 * lh, hp = 2.02 * fs + 5.2;
            us.push({ nodo: nd, a, b, txt: s.slice(a, b), cs, pos, caja: { x: pos.x, y: cy - hp / 2, w: pos.w, h: hp } });
          }
        }
      }
      bl.style.transform = giro;
      void rb;
      // capa con las copias; los originales se envuelven para ocultarlos sin mover el renglón
      let capa = bl.querySelector(":scope > .ao-capa");
      if (!capa) { capa = document.createElement("div"); capa.className = "ao-capa"; bl.append(capa); }
      // envolver de atrás hacia adelante para no mover los índices
      const porNodo = new Map();
      us.forEach((u) => { if (!porNodo.has(u.nodo)) porNodo.set(u.nodo, []); porNodo.get(u.nodo).push(u); });
      for (const [nd, l] of porNodo) {
        const par = nd.parentNode, s = nd.nodeValue, frag = document.createDocumentFragment();
        let pos = 0;
        l.sort((p, q) => p.a - q.a).forEach((u) => {
          if (u.a > pos) frag.append(document.createTextNode(s.slice(pos, u.a)));
          const sp = document.createElement("span"); sp.className = "ao-u"; sp.textContent = s.slice(u.a, u.b); frag.append(sp); u.orig = sp; pos = u.b;
        });
        if (pos < s.length) frag.append(document.createTextNode(s.slice(pos)));
        par.replaceChild(frag, nd);
      }
      us.forEach((u) => {
        const c = u.cs, el = document.createElement("span");
        el.textContent = u.txt.replace(/\s+$/, (m) => m);
        el.style.cssText = `left:${n2(u.pos.x - bl.offsetLeft)}px;top:${n2(u.pos.y - bl.offsetTop)}px;height:${n2(u.pos.h)}px;line-height:${n2(u.pos.h)}px;transform-origin:${n2(u.caja.x + u.caja.w / 2 - u.pos.x)}px ${n2(u.caja.y + u.caja.h / 2 - u.pos.y)}px;font:${c.fontStyle} ${c.fontWeight} ${c.fontSize} ${c.fontFamily};color:${c.color};letter-spacing:${c.letterSpacing};text-decoration:${c.textDecorationLine} ${c.textDecorationColor};text-shadow:${c.textShadow};font-kerning:${c.fontKerning};-webkit-text-stroke:${c.webkitTextStroke || "0"};`;
        capa.append(el);
        u.el = el;
        u.color = c.color;
      });
      return us;
    }

    // ---- compilar la lista en el árbol de la diapositiva
    let orden = 0;
    const efectos = [];   // { x, t (tiempo del efecto), bhs: [...], iter, dest... }
    function compilarEfecto(x) {
      const pp = x.ppt || (ctx.preset ? ctx.preset(x) : null);
      if (!pp || !pp.c) return null;
      const ctn = pp.c;
      const d0 = +pp.d0 || 0, f = d0 > 0 && x.dur > 0 ? x.dur / d0 : 1;
      const te = tiempoDe(ctn, f);
      te.b = Math.max(0, +x.retraso || 0);
      const it = hijo(ctn, "iterate");
      const iter = it ? { type: A(it).type || "el", atras: A(it).backwards === "1", pct: hijo(it, "tmPct") ? +A(hijo(it, "tmPct")).val / 100000 : null, abs: hijo(it, "tmAbs") ? +A(hijo(it, "tmAbs")).val * f : null } : null;
      const bhs = [];
      const chl = hijo(ctn, "childTnLst");
      (function rec(lst) {
        for (const n of hijosDe(lst)) {
          const l = loc(n[0]);
          if (l === "par" || l === "seq" || l === "excl") { const c = hijo(n, "cTn"); if (c) rec(hijo(c, "childTnLst")); continue; }
          const cb = hijo(n, "cBhvr") || (hijo(n, "cMediaNode") ? hijo(hijo(n, "cMediaNode"), "cTn") : null);
          const ctnB = cb ? hijo(cb, "cTn") || cb : null;
          if (!ctnB) continue;
          // los comportamientos para el reproductor de Internet Explorer (rctx="IE") PowerPoint no los usa: repiten lo
          // mismo que los suyos (en «Transparencia» la opacidad quedaría al cuadrado)
          if (cb && A(cb).rctx === "IE") continue;
          const tb = tiempoDe(ctnB, f);
          if (tb.d == null) tb.d = 0.001;
          tb.AD = activa(tb);
          const tgt = cb && hijo(cb, "tgtEl");
          const attrs = (cb && hijo(cb, "attrNameLst") ? hijosDe(hijo(cb, "attrNameLst"), "attrName").map((a) => textoDe(a)) : []);
          const rb = +(A(n)["p14:bounceEnd"] || 0) / 100000;
          bhs.push({ l, n, t: tb, tgt, attrs, aditivo: A(cb).additive || null, orden: orden++, rebote: rb > 0 ? rb : 0 });
        }
      })(chl);
      // PowerPoint aplica el inicio y fin suaves (y el filtro de tiempo) del efecto a cada comportamiento por separado,
      // no al efecto entero como SMIL (medido en «Silbido»: así coincide al medio píxel)
      if (te.acc || te.dec || te.tf) {
        bhs.forEach((b) => { if (!b.t.acc && !b.t.dec) { b.t.acc = te.acc; b.t.dec = te.dec; } if (!b.t.tf && te.tf) b.t.tf = te.tf; });
        te.acc = 0; te.dec = 0; te.tf = null;
      }
      // duración simple del efecto: lo más largo de sus comportamientos
      let D = 0;
      bhs.forEach((b) => { D = Math.max(D, b.t.b + b.t.AD); });
      if (te.d == null) te.d = D;
      // «después de la animación» (subTnLst): atenuar o esconder al terminar el efecto o al siguiente clic
      const desp = [];
      hijosDe(hijo(ctn, "subTnLst")).forEach((n) => {
        const cb = hijo(n, "cBhvr"), ctnB = cb && hijo(cb, "cTn");
        if (!ctnB) return;
        const st = hijo(ctnB, "stCondLst");
        const alFin = !!(st && hijosDe(st, "cond").some((q) => A(q).evt === "end"));
        const attrs = hijo(cb, "attrNameLst") ? hijosDe(hijo(cb, "attrNameLst"), "attrName").map((q) => textoDe(q)) : [];
        desp.push({ l: loc(n[0]), n, t: { b: 0, d: 0.001, rep: 1, rev: false, acc: 0, dec: 0, spd: 1, fill: "hold", AD: 0.001 }, tgt: hijo(cb, "tgtEl"), attrs, aditivo: null, orden: orden++, rebote: 0,
          alFin: alFin || A(ctnB).masterRel === "sameClick", rel: A(ctnB).masterRel || "nextClick" });
      });
      const ef = { x, t: te, bhs, iter, D, desp };
      return ef;
    }
    /** Los grupos de la secuencia principal y de cada desencadenador, como los arma PowerPoint. */
    function armarGrupos(items, seq) {
      const grupos = [];
      let g = null, sub = null;
      items.forEach((x, i) => {
        const ini = x.inicio || "clic";
        if (ini === "clic" || !g) { g = { auto: ini !== "clic" && !grupos.length, subs: [], seq: seq || "principal" }; grupos.push(g); sub = null; }
        if (!sub || ini === "despues") { sub = { efs: [] }; g.subs.push(sub); }
        const ef = compilarEfecto(x);
        if (ef) sub.efs.push(ef);
        void i;
      });
      return grupos;
    }
    const principal = lista.filter((x) => !x.disp);
    const disparos = new Map();
    lista.filter((x) => x.disp).forEach((x) => { if (!disparos.has(x.disp)) disparos.set(x.disp, []); disparos.get(x.disp).push(x); });
    // primero se resuelven los destinos (para saber cuántas letras tiene cada uno) y luego los tiempos
    const todosGrupos = { principal: armarGrupos(principal), disp: new Map([...disparos].map(([b, l]) => [b, armarGrupos(l.map((x, i) => (i === 0 && x.inicio !== "clic" ? Object.assign({}, x, { inicio: "clic" }) : x)), "d:" + b)])) };
    function resolver(grupos) {
      for (const g of grupos) {
        let t0 = 0;
        for (const s of g.subs) {
          s.b = t0;
          let fin = 0;
          for (const ef of s.efs) {
            ef.bhs.forEach((b) => { b.dest = destino(ef.x, b.tgt, ef.iter && ef.iter.type !== "el" ? ef.iter : null); });
            (ef.desp || []).forEach((b) => { b.dest = destino(ef.x, b.tgt, null); });
            const n = ef.iter ? Math.max(1, ...ef.bhs.map((b) => (b.dest && b.dest.unidades ? b.dest.unidades.length : 1))) : 1;
            ef.n = n;
            ef.delta = ef.iter ? (ef.iter.abs != null ? ef.iter.abs : (ef.iter.pct || 0) * ef.t.d) : 0;
            // duración activa del efecto: la de una unidad más el escalonado
            const unaVez = activa(ef.t);
            ef.t.AD = unaVez + (n - 1) * ef.delta;
            fin = Math.max(fin, ef.t.b + (isFinite(ef.t.AD) ? ef.t.AD : activa(Object.assign({}, ef.t, { rep: 1, repDur: null }))));
          }
          s.d = fin;
          t0 += fin;
        }
        g.d = t0;
      }
    }
    resolver(todosGrupos.principal);
    todosGrupos.disp.forEach(resolver);

    // ---- estado inicial: lo que entra primero empieza oculto
    const visibleIni = new Map();
    const marcarIni = (grupos) => grupos.forEach((g) => g.subs.forEach((s) => s.efs.forEach((ef) => {
      const pcl = A(ef.x.ppt ? ef.x.ppt.c : (ctx.preset(ef.x) || {}).c).presetClass || ({ e: "entr", x: "exit", n: "emph", r: "path" }[ef.x.tipo] || "emph");
      const vistos = new Set();
      ef.bhs.forEach((b) => {
        if (!b.dest || vistos.has(b.dest.clave)) return;
        vistos.add(b.dest.clave);
        if (!visibleIni.has(b.dest.clave)) visibleIni.set(b.dest.clave, pcl !== "entr");
      });
    })));
    marcarIni(todosGrupos.principal);
    todosGrupos.disp.forEach(marcarIni);

    // ---- estado de la reproducción: cuándo empezó cada grupo (ms del reloj del reproductor) y si ya terminó
    let reloj = 0;                       // tiempo actual del reproductor (ms)
    const corriendo = [];                // [{ g, inicio }] en el orden en que empezaron
    let siguiente = 0;                   // el siguiente grupo de la secuencia principal
    const sigDisp = new Map();
    let medio = new Map();               // bid → último comando de medio aplicado

    // ---- evaluar un cuadro
    function base(d) {
      const c = d.caja;
      return { vis: visibleIni.has(d.clave) ? visibleIni.get(d.clave) : true, op: 1, x: (c.x + c.w / 2) / T.w, y: (c.y + c.h / 2) / T.h, w: c.w / T.w, h: c.h / T.h, r: 0,
        sx: 1, sy: 1, mx: 0, my: 0, skx: 0, sky: 0, color: null, fill: null, stroke: null, fillOn: null, strokeOn: null, fw: null, fs: null, td: null, fsz: 1, filtros: [] };
    }
    const estados = new Map();
    function nuevoEstado(d, u) {
      const c = u == null ? d.caja : d.unidades[u].caja;
      const v = visibleIni.has(d.clave) ? visibleIni.get(d.clave) : true;
      const e = Object.assign(base(d), { x: (c.x + c.w / 2) / T.w, y: (c.y + c.h / 2) / T.h, w: c.w / T.w, h: c.h / T.h, vis: v });
      e.b0 = { ppt_x: e.x, ppt_y: e.y, ppt_w: e.w, ppt_h: e.h, ppt_r: 0, ppt_c: 0 };
      e.d = d; e.u = u;
      return e;
    }
    function estadoDe(d, u) {
      const k = u == null ? d.clave : d.clave + "#" + u;
      let e = estados.get(k);
      if (!e) { e = nuevoEstado(d, u); estados.set(k, e); }
      return e;
    }
    function aplicarBh(bh, s, u, e, ef) {
      const n = bh.n, a = A(n);
      // «ppt_x» (sin #) es el valor al EMPEZAR este comportamiento: lo que dejaron los que empezaron antes (no los que
      // empiezan con él). En «Voltear» el ancho se vuelve negativo y el corrimiento «ppt_w*0.5» debe usar el de antes.
      const vars = { $: 0, base: e.b0, act: e.foto || { ppt_x: e.x, ppt_y: e.y, ppt_w: e.w, ppt_h: e.h, ppt_r: e.r } };
      const prop = bh.attrs[0] || "";
      const pon = (p, v) => {
        switch (p) {
          case "style.visibility": e.vis = v === "visible" || v === true; break;
          case "style.opacity": e.op = Math.max(0, Math.min(1, +v)); break;
          case "ppt_x": e.x = +v; break;
          case "ppt_y": e.y = +v; break;
          case "ppt_w": e.w = +v; break;
          case "ppt_h": e.h = +v; break;
          case "r": case "style.rotation": case "ppt_r": e.r = +v; break;
          case "xshear": e.skx = +v; break;
          case "yshear": e.sky = +v; break;
          case "style.color": e.color = v; break;
          case "fillcolor": case "fillColor": case "fill.color": e.fill = v; break;
          case "stroke.color": e.stroke = v; break;
          case "fill.on": e.fillOn = v === "true" || v === true; break;
          case "fill.type": if (v === "solid") e.fillOn = e.fillOn == null ? true : e.fillOn; break;
          case "stroke.on": e.strokeOn = v === "true" || v === true; break;
          case "style.fontWeight": e.fw = v; break;
          case "style.fontStyle": e.fs = v; break;
          case "style.textDecorationUnderline": e.td = v === true || v === "true" ? "underline" : "none"; break;
          case "style.fontSize": e.fsz = +v || 1; break;
          case "imageData.grayscale": e.gris = v === "true" || v === true; break;
          // atenuar después de la animación: el color del objeto (el texto y, si tiene, el relleno)
          case "ppt_c": e.color = v; if (e.d && e.d.blq && e.d.blq.querySelector(".forma-svg path:not([fill='none'])")) e.fill = v; break;
          default: break;
        }
      };
      const actual = (p) => ({ "ppt_x": e.x, "ppt_y": e.y, "ppt_w": e.w, "ppt_h": e.h, r: e.r, "style.rotation": e.r, "style.opacity": e.op, "style.fontSize": e.fsz })[p];
      switch (bh.l) {
        case "set": {
          const to = hijo(n, "to"), v = to && hijosDe(to)[0];
          if (!v) return;
          const vl = loc(v[0]);
          let val = A(v).val;
          if (vl === "clrVal") val = colorNodo(v, ctx);
          else if (vl === "boolVal") val = val === "1" || val === "true";
          else if (/^(ppt_|r$|style\.(opacity|rotation|fontSize))/.test(prop) && typeof val === "string" && !/^(visible|hidden)$/.test(val)) val = valorNum(val, vars);
          bh.attrs.forEach((p) => pon(p, val));
          return;
        }
        case "anim": {
          const tl = hijo(n, "tavLst"), tavs = tl ? hijosDe(tl, "tav") : [];
          const tipoV = a.valueType || "num";
          const valDe = (tav) => {
            const vv = hijo(tav, "val"), c = vv && hijosDe(vv)[0];
            if (!c) return null;
            const cl = loc(c[0]);
            if (cl === "clrVal") return colorNodo(c, ctx);
            if (cl === "boolVal") return A(c).val === "1" || A(c).val === "true";
            if (tipoV === "str") return A(c).val;
            return valorNum(A(c).val, vars);
          };
          let val;
          if (tavs.length) {
            const tms = tavs.map((t, i) => (A(t).tm != null ? +A(t).tm / 100000 : tavs.length > 1 ? i / (tavs.length - 1) : 0));
            let i = 0;
            while (i < tavs.length - 1 && s >= tms[i + 1]) i++;
            const discreto = a.calcmode === "discrete" || tipoV === "str";
            // al final (o pasado el último tiempo) se usa el último tramo con f = 1: su fórmula también cuenta
            if (!discreto && i >= tavs.length - 1 && tavs.length > 1) i = tavs.length - 2;
            if (discreto || tavs.length === 1) val = valDe(tavs[Math.min(i, tavs.length - 1)]);
            else {
              // (con rebote al final el avance pasa de 1 y el valor se pasa del último, como en PowerPoint)
              const v0 = valDe(tavs[i]), v1 = valDe(tavs[i + 1]), f = tms[i + 1] > tms[i] ? Math.max(0, Math.min(s > 1 ? 2 : 1, (s - tms[i]) / (tms[i + 1] - tms[i]))) : 1;
              if (tipoV === "clr") val = mezclar(v0, v1, f, "rgb");
              else val = (+v0 || 0) + ((+v1 || 0) - (+v0 || 0)) * f;
            }
            const fm = A(tavs[Math.min(i, tavs.length - 1)]).fmla;
            if (fm && tipoV === "num") { vars.$ = +val; val = valorNum(fm, vars); }
          } else {
            const de = a.from != null ? valorNum(a.from, vars) : actual(prop) || 0;
            if (a.to != null) val = de + (valorNum(a.to, vars) - de) * s;
            else if (a.by != null) val = de + valorNum(a.by, vars) * s;
            else return;
            if (a.by != null && a.from == null) { val = valorNum(a.by, vars) * s; pon(prop, (actual(prop) || 0) + val); return; }
          }
          if (val == null) return;
          if (bh.aditivo === "sum" && tipoV === "num") val = (actual(prop) || 0) + +val;
          else if (bh.aditivo === "mult" && tipoV === "num") val = (actual(prop) || 0) * +val;
          bh.attrs.forEach((p) => pon(p, val));
          return;
        }
        case "animClr": {
          const de = hijo(n, "from") ? colorNodo(hijo(n, "from"), ctx) : null;
          const prop0 = prop === "fillcolor" || prop === "fillColor" || prop === "fill.color" ? "fill" : prop === "stroke.color" ? "stroke" : "color";
          const cur = de || e[prop0] || colorActual(e, prop0);
          let fin = null;
          const to = hijo(n, "to"), by = hijo(n, "by");
          if (to) fin = colorNodo(to, ctx);
          else if (by) {
            const h = hijo(by, "hsl"), r = hijo(by, "rgb");
            if (h && cur) { const c = rgbHsl(hexRgb(cur)); c[0] += (+A(h).h || 0) / 60000; c[1] = Math.max(0, Math.min(1, c[1] + (+A(h).s || 0) / 100000)); c[2] = Math.max(0, Math.min(1, c[2] + (+A(h).l || 0) / 100000)); fin = rgbHex(hslRgb(c)); bh.porHsl = +A(h).h || 0; }
            else if (r && cur) { const c = hexRgb(cur); fin = rgbHex([c[0] + ((+A(r).r || 0) / 100000) * 255, c[1] + ((+A(r).g || 0) / 100000) * 255, c[2] + ((+A(r).b || 0) / 100000) * 255]); }
          }
          if (!fin || !cur) return;
          const v = mezclar(cur, fin, s, a.clrSpc || "rgb", a.dir, bh.porHsl);
          bh.attrs.forEach((p) => pon(p, v));
          return;
        }
        case "animScale": {
          const by = hijo(n, "by"), to = hijo(n, "to"), fr = hijo(n, "from");
          const xy = (q) => [(+A(q).x || 100000) / 100000, (+A(q).y || 100000) / 100000];
          if (by && !to) { const [bx, byy] = xy(by); e.sx *= 1 + (bx - 1) * s; e.sy *= 1 + (byy - 1) * s; return; }
          const [fx, fy] = fr ? xy(fr) : [e.sx, e.sy], [tx, ty] = to ? xy(to) : [fx, fy];
          e.sx = fx + (tx - fx) * s; e.sy = fy + (ty - fy) * s;
          return;
        }
        case "animRot": {
          const g = (v) => (+v || 0) / 60000;
          if (a.by != null && a.to == null) { e.r += g(a.by) * s; return; }
          const de = a.from != null ? g(a.from) : e.r;
          const hasta = a.to != null ? g(a.to) : de;
          e.r = de + (hasta - de) * s;
          return;
        }
        case "animMotion": {
          if (!bh.camino && a.path) bh.camino = leerCamino(a.path);
          let px = 0, py = 0;
          if (bh.camino) [px, py] = puntoCamino(bh.camino, s);
          else {
            const pt = (q) => (q ? [(+A(q).x || 0) / 100000, (+A(q).y || 0) / 100000] : null);
            const by = pt(hijo(n, "by")), fr = pt(hijo(n, "from")) || [0, 0], to = pt(hijo(n, "to"));
            if (by) { px = by[0] * s; py = by[1] * s; } else if (to) { px = fr[0] + (to[0] - fr[0]) * s; py = fr[1] + (to[1] - fr[1]) * s; }
          }
          const ang = (+a.rAng || 0) / 60000;
          if (ang) { const c = Math.cos((ang * Math.PI) / 180), sn = Math.sin((ang * Math.PI) / 180); const qx = px * T.w, qy = py * T.h; px = (qx * c - qy * sn) / T.w; py = (qx * sn + qy * c) / T.h; }
          e.mx += px; e.my += py;
          return;
        }
        case "animEffect": {
          const pr = a.prLst || "";
          if (a.filter === "image") { const m = /opacity\s*:\s*([\d.]+)/i.exec(pr); if (m) e.op *= +m[1]; return; }
          const trans = a.transition || "in";
          if (trans === "none" && !/fade/.test(a.filter || "")) return;
          e.filtros.push({ f: a.filter, s, trans, pr, sem: ef ? ((ef.x.id || "").length * 31 + (u || 0) * 7 + bh.orden) : 7 });
          return;
        }
        case "cmd": {
          if (ctx.alMedio && bh.dest) { const clave = bh.dest.bid + "|" + a.cmd; if (s >= 0 && !medio.get(bh.ordenUnico + "")) { medio.set(bh.ordenUnico + "", true); ctx.alMedio(bh.dest.bid, a.type, a.cmd); } void clave; }
          return;
        }
        default: return;
      }
    }
    function colorActual(e, prop) {
      const bl = e.d.blq;
      if (!bl) return "#000000";
      if (prop === "color") { const t = contTexto(bl); return t ? aHex(getComputedStyle(t.querySelector("span, b, i, u") || t).color) : "#000000"; }
      const p = bl.querySelector(".forma-svg path:not([fill='none'])");
      // sin relleno, PowerPoint parte del color de relleno de la figura aunque no se vea: el blanco del tema (bg1)
      if (prop === "fill") { if (p) return aHex(p.getAttribute("fill")); const bgc = getComputedStyle(contTexto(bl) || bl).backgroundColor; return /rgba\(.*,\s*0\)$|transparent/.test(bgc) ? (ctx.colorTema && ctx.colorTema("bg1")) || "#FFFFFF" : aHex(bgc); }
      const q = bl.querySelector(".forma-svg path:not([stroke='none'])");
      return q ? aHex(q.getAttribute("stroke")) : "#000000";
    }
    function aHex(c) {
      if (!c) return "#000000";
      if (/^#[0-9a-f]{6}$/i.test(c)) return c.toUpperCase();
      const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(c);
      return m ? rgbHex([+m[1], +m[2], +m[3]]) : "#000000";
    }
    function mezclar(c0, c1, f, esp, dir, porH) {
      const a = hexRgb(c0), b = hexRgb(c1);
      if (!a || !b) return c1 || c0;
      if (esp === "hsl") {
        const p = rgbHsl(a), q = rgbHsl(b);
        let dh = porH != null ? porH / 60000 : q[0] - p[0];
        if (porH == null) { if (dir === "ccw") { if (dh > 0) dh -= 360; } else if (dh < 0) dh += 360; }
        return rgbHex(hslRgb([p[0] + dh * f, p[1] + (q[1] - p[1]) * f, p[2] + (q[2] - p[2]) * f]));
      }
      return rgbHex(a.map((v, i) => v + (b[i] - v) * f));
    }

    /** Recorre los grupos que ya empezaron y aplica, en orden de inicio, lo que esté activo en el tiempo ahora. */
    /** Lo que está activo en el instante T del reloj: cada comportamiento con su avance y su momento de inicio. */
    function activosEn(T, inclusivo, soloDest) {
      const activos = [];
      corriendo.forEach((c, ci) => {
        const tg = T - c.inicio;
        if (tg < 0) return;
        // el clic que sigue en la misma secuencia (para «hasta el siguiente clic» y lo que pasa al siguiente clic)
        const sig = corriendo.find((q, qi) => qi > ci && q.g.seq === c.g.seq && !q.g.auto);
        c.g.subs.forEach((s) => {
          const ts0 = tg - s.b;
          if (ts0 < 0) return;
          s.efs.forEach((ef) => {
            let ts = ts0;
            if (ef.t.hastaClic && sig) ts = Math.min(ts, sig.inicio - c.inicio - s.b);
            if (ef.desp && ef.desp.length) {
              const fin = ef.t.b + (isFinite(ef.t.AD) ? ef.t.AD : Infinity);
              ef.desp.forEach((bh) => {
                if (!bh.dest || (soloDest && bh.dest !== soloDest)) return;
                let k = null;
                if (bh.alFin && ts0 >= fin) k = c.inicio + s.b + fin;
                else if (!bh.alFin && sig && T >= sig.inicio) k = sig.inicio;
                if (k != null) activos.push({ bh, s: 1, u: null, ef, k: [ci, k, bh.orden] });
              });
            }
            const reps = ef.n || 1;
            for (let j = 0; j < reps; j++) {
              const jj = ef.iter && ef.iter.atras ? reps - 1 - j : j;
              const tEf = Object.assign({}, ef.t, { b: ef.t.b + j * (ef.delta || 0), AD: activa(ef.t) });
              const ue = local(tEf, ts, inclusivo);
              if (ue == null) continue;
              ef.bhs.forEach((bh) => {
                if (!bh.dest || (soloDest && bh.dest !== soloDest)) return;
                const ub = local(bh.t, ue, inclusivo);
                if (ub == null) return;
                let s2 = bh.t.d > 0 && isFinite(bh.t.d) ? ub / bh.t.d : ub > 0 || bh.t.d === 0 ? 1 : 0;
                s2 = Math.max(0, Math.min(1, s2));
                if (bh.rebote) s2 = rebote(s2, bh.rebote);
                activos.push({ bh, s: s2, u: bh.dest.unidades ? jj : null, ef, k: [ci, c.inicio + s.b + tEf.b + bh.t.b, bh.orden] });
              });
            }
          });
        });
      });
      activos.sort((p, q) => p.k[1] - q.k[1] || p.k[0] - q.k[0] || p.k[2] - q.k[2]);
      return activos;
    }
    /**
     * «ppt_x» (sin #) en una fórmula es el valor al EMPEZAR ese comportamiento: lo que dejaron, en ese instante, los
     * que empezaron antes (aunque terminen justo ahí). Se calcula una vez por destino y momento de inicio.
     */
    const fotos = new Map();
    function fotoDe(d, u, Tb, prof) {
      const clave = d.clave + "#" + (u == null ? "" : u) + "@" + Math.round(Tb * 10);
      if (fotos.has(clave)) return fotos.get(clave);
      const e = nuevoEstado(d, u);
      if ((prof || 0) < 12) {
        for (const ac of activosEn(Tb, true, d)) {
          if (ac.k[1] >= Tb - 1e-6) continue;
          if ((ac.u == null ? null : ac.u) !== (u == null ? null : u) && d.unidades) continue;
          e.foto = fotoDe(d, u, ac.k[1], (prof || 0) + 1);
          aplicarBh(ac.bh, ac.s, u, e, ac.ef);
        }
      }
      const f = { ppt_x: e.x, ppt_y: e.y, ppt_w: e.w, ppt_h: e.h, ppt_r: e.r };
      fotos.set(clave, f);
      return f;
    }
    /** Recorre los grupos que ya empezaron y aplica, en orden de inicio, lo que esté activo en el tiempo ahora. */
    function evaluar() {
      estados.clear();
      for (const ac of activosEn(reloj, false, null)) {
        const d = ac.bh.dest;
        const u = d.unidades && ac.u != null ? ac.u : null;
        const e = d.unidades && ac.u != null ? (d.unidades[ac.u] ? estadoDe(d, ac.u) : null) : estadoDe(d, null);
        if (!e) continue;
        e.foto = fotoDe(d, u, ac.k[1]);
        aplicarBh(ac.bh, ac.s, u, e, ac.ef);
      }
      pintar();
    }

    // ---- pintar: cada destino con su estado
    const ultimo = new Map();
    function pintar() {
      destinos.forEach((d) => {
        const e = estados.get(d.clave) || estadoDe(d, null);
        pintarEl(d, e, d.els, d.caja);
        if (d.unidades) {
          const eus = d.unidades.map((u, i) => estados.get(d.clave + "#" + i) || estadoDe(d, i));
          if (eus.some((eu) => eu.fsz !== 1) || d.conTamano) acomodarUnidades(d, eus);
          d.unidades.forEach((u, i) => pintarUnidad(u, eus[i]));
        }
      });
    }
    function transformaDe(e, c, giroBase) {
      const dx = (e.x - e.b0.ppt_x + e.mx) * T.w, dy = (e.y - e.b0.ppt_y + e.my) * T.h;
      const sx = (e.b0.ppt_w ? e.w / e.b0.ppt_w : 1) * e.sx, sy = (e.b0.ppt_h ? e.h / e.b0.ppt_h : 1) * e.sy;
      const r = (giroBase || 0) + e.r;
      let t = "";
      if (Math.abs(dx) > 0.01 || Math.abs(dy) > 0.01) t += `translate(${n2(dx)}px,${n2(dy)}px) `;
      if (r) t += `rotate(${Math.round(r * 100) / 100}deg) `;
      if (e.skx || e.sky) t += `skew(${n2(Math.atan(e.skx) * 180 / Math.PI)}deg,${n2(Math.atan(e.sky) * 180 / Math.PI)}deg) `;
      if (Math.abs(sx - 1) > 1e-4 || Math.abs(sy - 1) > 1e-4) t += `scale(${Math.round(sx * 10000) / 10000},${Math.round(sy * 10000) / 10000})`;
      void c;
      return t.trim();
    }
    function clipDe(e, w, h) {
      let op = 1, mascaras = [], nada = false, desliza = null;
      for (const f of e.filtros) {
        const m = mascara(f.f, f.s, w, h, f.pr, f.sem, f.trans);
        if (!m) continue;
        if (m.nada) nada = true;
        else if (m.opacidad != null) op *= m.opacidad;
        else if (m.mascara) mascaras.push(m.mascara);
        else if (m.desliza) desliza = m;
      }
      return { op: nada ? 0 : op, mascaras, desliza };
    }
    function pintarEl(d, e, els, caja) {
      els.forEach((el) => {
        const bl = d.tipo === "bloque" ? el : null;
        const datos = bl && ctx.bloque ? ctx.bloque(bl.dataset.id) || {} : {};
        const giro = d.tipo === "bloque" ? +datos.rot || 0 : 0;
        const op0 = d.tipo === "bloque" && datos.opacidad != null && datos.opacidad < 1 ? +datos.opacidad : 1;
        const w = el.offsetWidth || caja.w, h = el.offsetHeight || caja.h;
        const cl = clipDe(e, d.tipo === "fondo" ? w : caja.w, d.tipo === "fondo" ? h : caja.h);
        let st = "";
        st += `visibility:${e.vis ? "visible" : "hidden"};`;
        const op = e.op * cl.op * op0;
        if (op < 0.999 || op0 < 1) st += `opacity:${Math.round(op * 1000) / 1000};`;
        const tr = transformaDe(e, caja, giro);
        // el origen: el centro del destino (en un grupo, el centro del grupo)
        if (d.tipo === "bloque") {
          const ox = caja.x + caja.w / 2 - el.offsetLeft, oy = caja.y + caja.h / 2 - el.offsetTop;
          st += `transform-origin:${n2(ox)}px ${n2(oy)}px;transform:${tr || "none"};`;
          // la máscara está en la caja del destino (en un grupo, la del grupo): se acomoda en la del elemento
          if (cl.mascaras.length) st += cssMascara(cl.mascaras, caja.x - el.offsetLeft, caja.y - el.offsetTop, caja.w, caja.h);
        } else {
          st += `transform-origin:50% 50%;transform:${tr || "none"};`;
          if (cl.mascaras.length) st += cssMascara(cl.mascaras, 0, 0, d.tipo === "fondo" ? w : caja.w, d.tipo === "fondo" ? h : caja.h);
          if (d.tipo === "parrafo") st += "position:relative;";
        }
        if (e.gris) st += "filter:grayscale(1);";
        const clave = st + "|" + (e.color || "") + (e.fill || "") + (e.stroke || "") + (e.fw || "") + (e.fs || "") + (e.td || "") + e.fsz + (e.fillOn == null ? "" : e.fillOn);
        if (ultimo.get(el) === clave) return;
        ultimo.set(el, clave);
        if (!el.dataset.aoBase) el.dataset.aoBase = el.getAttribute("style") || "";
        el.setAttribute("style", quitarProps(el.dataset.aoBase) + st);
        // colores y letra
        const txt = d.tipo === "parrafo" ? el : contTexto(bl || el) || null;
        if (txt) {
          txt.classList.toggle("ao-color", !!e.color); if (e.color) txt.style.setProperty("--ao-color", e.color);
          txt.classList.toggle("ao-fw", !!e.fw); if (e.fw) txt.style.setProperty("--ao-fw", e.fw);
          txt.classList.toggle("ao-fs", !!e.fs); if (e.fs) txt.style.setProperty("--ao-fs", e.fs);
          txt.classList.toggle("ao-td", !!e.td); if (e.td) txt.style.setProperty("--ao-td", e.td);
          if (e.fsz !== 1 || txt.dataset.aoTam) {
            txt.querySelectorAll("*").forEach((x) => { if (x.dataset.aoTam == null) x.dataset.aoTam = parseFloat(getComputedStyle(x).fontSize) || 16; x.style.fontSize = n2(+x.dataset.aoTam * e.fsz) + "px"; });
            if (txt.dataset.aoTam == null) txt.dataset.aoTam = parseFloat(getComputedStyle(txt).fontSize) || 16;
            txt.style.fontSize = n2(+txt.dataset.aoTam * e.fsz) + "px";
          }
        }
        if (d.tipo === "bloque" || d.tipo === "fondo") {
          const svg = d.tipo === "fondo" ? el : el.querySelector(".forma-svg");
          if (svg) {
            const rell = [...svg.querySelectorAll("path")].filter((p) => p.getAttribute("fill") && p.getAttribute("fill") !== "none");
            const traz = [...svg.querySelectorAll("path")].filter((p) => p.getAttribute("stroke") && p.getAttribute("stroke") !== "none");
            const objR = rell.length ? rell : e.fill && e.fillOn ? [svg.querySelector("path")].filter(Boolean) : [];
            objR.forEach((p) => { p.style.fill = e.fill || ""; });
            traz.forEach((p) => { p.style.stroke = e.stroke || ""; });
          } else if (bl && txt) {
            if (txt.dataset.aoFondo == null) txt.dataset.aoFondo = txt.style.backgroundColor || "";
            txt.style.backgroundColor = e.fillOn === false ? "transparent" : e.fill || txt.dataset.aoFondo;
          }
        }
      });
    }
    function cssMascara(ms, x, y, w, h) {
      const l = ms.join(","), rep = ms.map(() => "no-repeat").join(","), tam = ms.map(() => `${n2(w)}px ${n2(h)}px`).join(","), pos = ms.map(() => `${n2(x)}px ${n2(y)}px`).join(",");
      return `-webkit-mask-image:${l};mask-image:${l};-webkit-mask-repeat:${rep};mask-repeat:${rep};-webkit-mask-size:${tam};mask-size:${tam};-webkit-mask-position:${pos};mask-position:${pos};` + (ms.length > 1 ? "-webkit-mask-composite:source-in;mask-composite:intersect;" : "");
    }
    const quitarProps = (s) => String(s || "").replace(/(^|;)\s*(visibility|opacity|transform|transform-origin|clip-path|-webkit-mask-[a-z]+|mask-[a-z]+|filter)\s*:[^;]*/g, "$1").replace(/;;+/g, ";").replace(/^;/, "") + (s && !/;\s*$/.test(s) ? ";" : "");
    /**
     * Letras que cambian de tamaño: en PowerPoint el texto se vuelve a acomodar (en «Aumentar con color» la palabra pasa al
     * siguiente renglón). Se cambia el tamaño de la letra original (oculta) y se vuelve a medir dónde quedó cada una.
     */
    function acomodarUnidades(d, eus) {
      const bl = d.blq, k = k0();
      d.conTamano = eus.some((eu) => eu.fsz !== 1);
      d.unidades.forEach((u, i) => {
        if (u.fs0 == null) u.fs0 = parseFloat(u.cs.fontSize) || 16;
        const fs = u.fs0 * eus[i].fsz;
        u.orig.style.fontSize = Math.abs(eus[i].fsz - 1) < 1e-4 ? "" : n2(fs) + "px";
        u.el.style.fontSize = n2(fs) + "px";
      });
      const giro = bl.style.transform; bl.style.transform = "none";
      const r0 = bl.getBoundingClientRect();
      d.unidades.forEach((u) => {
        const r = u.orig.getBoundingClientRect();
        u.pos = { x: bl.offsetLeft + (r.left - r0.left) / k, y: bl.offsetTop + (r.top - r0.top) / k, w: r.width / k, h: r.height / k };
        const cs2 = getComputedStyle(u.orig), fs2 = parseFloat(cs2.fontSize) || 16; let lh2 = parseFloat(cs2.lineHeight); if (!isFinite(lh2)) lh2 = fs2 * 1.2;
        const cy2 = u.pos.y + u.pos.h / 2 - lh2 / 2 + 0.6 * lh2, hp2 = 2.02 * fs2 + 5.2;
        u.caja = { x: u.pos.x, y: cy2 - hp2 / 2, w: u.pos.w, h: hp2 };
        u.el.style.left = n2(u.pos.x - bl.offsetLeft) + "px"; u.el.style.top = n2(u.pos.y - bl.offsetTop) + "px";
        u.el.style.height = n2(u.pos.h) + "px"; u.el.style.lineHeight = n2(u.pos.h) + "px";
        u.el.style.transformOrigin = `${n2(u.caja.x + u.caja.w / 2 - u.pos.x)}px ${n2(u.caja.y + u.caja.h / 2 - u.pos.y)}px`;
        u.el.dataset.aoBase = u.el.getAttribute("style").replace(/(^|;)\s*(visibility|opacity|transform|transform-origin|-webkit-mask-[a-z]+|mask-[a-z]+|color|font-weight|font-style|text-decoration-line)\s*:[^;]*/g, "$1");
      });
      bl.style.transform = giro;
    }
    function pintarUnidad(u, e) {
      const el = u.el;
      const c = u.caja;
      let st = `visibility:${e.vis ? "visible" : "hidden"};`;
      const cl = clipDe(e, c.w, c.h);
      const op = e.op * cl.op;
      if (op < 0.999) st += `opacity:${Math.round(op * 1000) / 1000};`;
      const tr = transformaDe(e, c, 0), ps = u.pos || c;
      // la copia está donde se ve la letra (pos); se transforma y se recorta en la caja de PowerPoint (c)
      st += `transform-origin:${n2(c.x + c.w / 2 - ps.x)}px ${n2(c.y + c.h / 2 - ps.y)}px;transform:${tr || "none"};`;
      if (cl.mascaras.length) st += cssMascara(cl.mascaras, c.x - ps.x, c.y - ps.y, c.w, c.h);
      if (e.color) st += `color:${e.color};`;
      if (e.fw) st += `font-weight:${e.fw};`;
      if (e.fs) st += `font-style:${e.fs};`;
      if (e.td) st += `text-decoration-line:${e.td};`;
      if (ultimo.get(el) === st) return;
      ultimo.set(el, st);
      if (!el.dataset.aoBase) el.dataset.aoBase = el.getAttribute("style") || "";
      el.setAttribute("style", quitarProps(el.dataset.aoBase) + st);
    }

    // ---- control
    let raf = 0, t0Real = 0, relojBase = 0, alTerminar = null;
    const grupoActivo = () => corriendo.length ? corriendo[corriendo.length - 1] : null;
    const terminoGrupo = (c) => !c || reloj - c.inicio >= c.g.d;
    function bucle() {
      raf = 0;
      reloj = relojBase + (performance.now() - t0Real);
      evaluar();
      const sigue = corriendo.some((c) => reloj - c.inicio < c.g.d || hayInfinito(c));
      if (sigue) raf = requestAnimationFrame(bucle);
      else if (alTerminar) { const f = alTerminar; alTerminar = null; f(); }
    }
    const hayInfinito = (c) => !isFinite(c.g.d) || c.g.subs.some((s) => s.efs.some((ef) => !isFinite(activa(ef.t))));
    function arrancar() { if (raf) return; relojBase = reloj; t0Real = performance.now(); raf = requestAnimationFrame(bucle); }
    function empezar(g) { corriendo.push({ g, inicio: reloj }); fotos.clear(); }
    /** Termina de golpe lo que esté corriendo (PowerPoint «busca» el final al hacer clic antes de que acabe). */
    function alFinal() {
      let fin = reloj;
      corriendo.forEach((c) => { if (isFinite(c.g.d)) fin = Math.max(fin, c.inicio + c.g.d); });
      reloj = fin;
    }
    const pasos = todosGrupos.principal.filter((g) => !g.auto).length;
    const api = {
      pasos,
      /** Al llegar a la diapositiva: lo que arranca solo. */
      inicio() {
        if (todosGrupos.principal[0] && todosGrupos.principal[0].auto) { empezar(todosGrupos.principal[0]); siguiente = 1; arrancar(); }
        else { siguiente = 0; evaluar(); }
      },
      /** El estado de antes de empezar (lo que entra después, oculto), sin arrancar nada: para la transición. */
      preparar() { if (raf) { cancelAnimationFrame(raf); raf = 0; } corriendo.length = 0; reloj = 0; siguiente = 0; fotos.clear(); evaluar(); },
      /** ¿Quedan clics de la secuencia principal? */
      quedan: () => siguiente < todosGrupos.principal.length,
      /** Un clic: lo que siga de la secuencia principal. false si ya no queda nada (pasar a la siguiente diapositiva). */
      clic() {
        if (raf) { cancelAnimationFrame(raf); raf = 0; reloj = relojBase + (performance.now() - t0Real); }
        const act = grupoActivo();
        if (act && !terminoGrupo(act) && isFinite(act.g.d)) { alFinal(); evaluar(); }
        if (siguiente >= todosGrupos.principal.length) return false;
        empezar(todosGrupos.principal[siguiente++]);
        arrancar();
        return true;
      },
      /** Sin animar: todo lo de los primeros k clics ya hecho (final: también el último a su final). */
      irAPaso(k, final) {
        if (raf) { cancelAnimationFrame(raf); raf = 0; }
        corriendo.length = 0; reloj = 0; siguiente = 0; fotos.clear();
        const gs = todosGrupos.principal;
        let hechos = 0;
        while (siguiente < gs.length && (gs[siguiente].auto || hechos < k)) {
          if (!gs[siguiente].auto) hechos++;
          empezar(gs[siguiente++]);
          alFinal();
        }
        if (final) while (siguiente < gs.length) { empezar(gs[siguiente++]); alFinal(); }
        evaluar();
      },
      esDisparador: (bid) => todosGrupos.disp.has(bid),
      /** Clic en una figura que desencadena animaciones. */
      disparar(bid) {
        const gs = todosGrupos.disp.get(bid);
        if (!gs || !gs.length) return false;
        if (raf) { cancelAnimationFrame(raf); raf = 0; reloj = relojBase + (performance.now() - t0Real); }
        const i = sigDisp.get(bid) || 0;
        empezar(gs[i % gs.length]);
        sigDisp.set(bid, i + 1);
        arrancar();
        return true;
      },
      /** Para medir: el grupo k (de la secuencia principal, contando el automático) en el instante t (ms desde que empezó). */
      cuadro(k, t) {
        if (raf) { cancelAnimationFrame(raf); raf = 0; }
        corriendo.length = 0; reloj = 0; siguiente = 0; fotos.clear();
        const gs = todosGrupos.principal;
        for (let i = 0; i < k && i < gs.length; i++) { empezar(gs[i]); alFinal(); }
        if (gs[k]) { empezar(gs[k]); reloj += t; }
        evaluar();
      },
      /** Duración (ms) de cada grupo de la secuencia principal. */
      duraciones: () => todosGrupos.principal.map((g) => g.d),
      corriendo: () => !!raf,
      alTerminar(f) { alTerminar = f; },
      detener() { if (raf) cancelAnimationFrame(raf); raf = 0; },
    };
    return api;
  }

  // ================================================================== leer p:timing (importar)
  /**
   * El p:timing de una diapositiva (elemento DOM) a la lista de NOM. bloquesDe(spid) → [ids de bloque] (una figura
   * agrupada da varios bloques). colorDe(nodoDom) → "#RRGGBB" para los colores del tema. Devuelve
   * { lista, medios, bld } o null si no hay nada.
   */
  function leerTiming(tmEl, bloquesDe, colorDe) {
    if (!tmEl) return null;
    const raiz = deDom(tmEl);
    // colores: se resuelven aquí, con el tema de la diapositiva (y se conserva el original para exportar)
    if (colorDe) (function rec(n, el) {
      if (!Array.isArray(n)) return;
      if (/^(srgbClr|schemeClr|sysClr|prstClr|hslClr|scrgbClr)$/.test(loc(n[0])) && el) { try { const c = colorDe(el); if (c) n[1] = Object.assign({}, n[1], { _c: c }); } catch (e) { /* sin color */ } }
      const hs = H(n), els = el ? [...el.children] : [];
      let j = 0;
      hs.forEach((c) => { if (Array.isArray(c)) rec(c, els[j++]); });
    })(raiz, tmEl);
    const lista = [], medios = [], ids = new Set();
    let huerfanos = 0;
    // cambiar el número de figura por el bloque; devuelve el bloque principal del efecto
    const conBloques = (n) => {
      let primero = null;
      (function rec(m) {
        if (!Array.isArray(m)) return;
        if (loc(m[0]) === "spTgt" || loc(m[0]) === "inkTgt") {
          const bs = bloquesDe(A(m).spid) || [];
          if (bs.length) { m[1] = Object.assign({}, A(m), { bid: bs[0] }); if (!primero) primero = { bid: bs[0], ids: bs, spid: A(m).spid }; }
          else { if (!primero) primero = { bid: null, ids: [], spid: A(m).spid }; }
        }
        H(m).forEach(rec);
      })(n);
      return primero;
    };
    const tnLst = hijo(raiz, "tnLst"), rootPar = tnLst && hijo(tnLst, "par"), rootCtn = rootPar && hijo(rootPar, "cTn");
    const hijosRaiz = rootCtn ? hijosDe(hijo(rootCtn, "childTnLst")) : [];
    const efectosDe = (seq, disp) => {
      const ctnS = hijo(seq, "cTn");
      const grupos = ctnS ? hijosDe(hijo(ctnS, "childTnLst"), "par") : [];
      grupos.forEach((gp) => {
        const gctn = hijo(gp, "cTn");
        hijosDe(hijo(gctn, "childTnLst"), "par").forEach((sp) => {
          const sctn = hijo(sp, "cTn");
          hijosDe(hijo(sctn, "childTnLst"), "par").forEach((ep) => {
            const ectn = hijo(ep, "cTn");
            if (!ectn) return;
            const a = A(ectn);
            const ref = conBloques(ectn);
            const nt = a.nodeType || "";
            const inicio = nt === "clickEffect" ? "clic" : nt === "afterEffect" ? "despues" : "con";
            const st = hijo(ectn, "stCondLst"), c0 = st && hijo(st, "cond");
            const retraso = c0 && A(c0).delay && A(c0).delay !== "indefinite" ? +A(c0).delay || 0 : 0;
            // lo propio del efecto, sin id, tipo de nodo ni retraso (eso lo lleva la lista)
            const c = copia(ectn);
            (function limpiar(m) { if (!Array.isArray(m)) return; if (loc(m[0]) === "cTn" && m[1]) { m[1] = Object.assign({}, m[1]); delete m[1].id; } if (loc(m[0]) === "tn" && a.id != null && A(m).val === a.id) m[1] = Object.assign({}, A(m), { val: "@ef" }); H(m).forEach(limpiar); })(c);
            delete c[1].nodeType;
            const ini = hijo(c, "stCondLst"); if (ini) c[2] = H(c).filter((h) => h !== ini);
            const pcl = a.presetClass || "";
            const tipo = { entr: "e", exit: "x", emph: "n", path: "r", verb: "m", mediacall: "m" }[pcl] || "n";
            const d0 = duracionNatural(c);
            const item = { id: ref && ref.bid, tipo, efecto: `ppt:${pcl}:${a.presetID || 0}:${a.presetSubtype || 0}`, inicio, retraso, dur: Math.round(d0) || 1, ppt: { c, d0: Math.round(d0) || 1 } };
            if (ref && ref.ids.length > 1) item.ids = ref.ids;
            if (!item.id) { item.spid = ref ? ref.spid : null; huerfanos++; }
            if (disp) item.disp = disp;
            if (A(sctn).presetClass) void 0;
            lista.push(item);
            if (item.id) ids.add(item.id);
          });
        });
      });
    };
    hijosRaiz.forEach((h) => {
      const l = loc(h[0]), ctn = hijo(h, "cTn") || (hijo(h, "cMediaNode") ? hijo(hijo(h, "cMediaNode"), "cTn") : null);
      const nt = A(ctn).nodeType;
      if (l === "seq" && nt === "mainSeq") efectosDe(h, null);
      else if (l === "seq" && nt === "interactiveSeq") {
        const st = hijo(ctn, "stCondLst"), cond = st && hijo(st, "cond"), tg = cond && hijo(cond, "tgtEl"), sp = tg && hijo(tg, "spTgt");
        const bs = sp ? bloquesDe(A(sp).spid) || [] : [];
        const bm = cond && A(cond).evt === "onMediaBookmark" ? A(cond) : null;
        const disp = bs[0] || (sp ? "spid:" + A(sp).spid : "?");
        const antes = lista.length;
        efectosDe(h, disp);
        if (bm) lista.slice(antes).forEach((x) => { x.marcador = (hijo(cond, "tn") ? A(hijo(cond, "tn")).val : "") || ""; });
      } else {
        const m = copia(h);
        conBloques(m);
        medios.push(m);
      }
    });
    // un efecto por párrafo seguido (como lo escribe PowerPoint con «por párrafos»): se marcan como un grupo
    const parDe = (x) => { let r = null; (function rec(m) { if (r != null || !Array.isArray(m)) return; if (loc(m[0]) === "pRg") r = +A(m).st; else if (loc(m[0]) === "bg") r = -1; H(m).forEach(rec); })(x.ppt.c); return r; };
    let gP = null;
    lista.forEach((x, i) => {
      const k = x.id ? parDe(x) : null;
      const prev = lista[i - 1];
      if (k == null) { gP = null; return; }
      if (gP && prev && prev.grupoP === gP && prev.id === x.id && prev.efecto === x.efecto && prev.disp === x.disp) { x.grupoP = gP; x.par = k; }
      else { gP = "i" + i; x.grupoP = gP; x.par = k; }
    });
    lista.forEach((x) => { if (x.grupoP && lista.filter((y) => y.grupoP === x.grupoP).length < 2) { delete x.grupoP; if (x.par === -1) delete x.par; } });
    const bl = hijo(raiz, "bldLst");
    const bld = bl ? hijosDe(bl).map((b) => { const m = copia(b); const bs = bloquesDe(A(b).spid) || []; if (bs.length) m[1] = Object.assign({}, A(m), { bid: bs[0] }); return m; }) : [];
    if (!lista.length && !medios.length) return null;
    return { lista, medios, bld, huerfanos };
  }
  /** Lo que dura un efecto (ms) sin repeticiones del efecto: lo más largo de sus comportamientos. */
  function duracionNatural(ctn) {
    let D = 0;
    (function rec(lst, base) {
      for (const n of hijosDe(lst)) {
        const l = loc(n[0]);
        if (l === "par" || l === "seq") { const c = hijo(n, "cTn"); if (c) { const t = tiempoDe(c, 1); rec(hijo(c, "childTnLst"), base + t.b); } continue; }
        const cb = hijo(n, "cBhvr") || (hijo(n, "cMediaNode") ? hijo(hijo(n, "cMediaNode"), "cTn") : null), c = cb ? hijo(cb, "cTn") || cb : null;
        if (!c) continue;
        const t = tiempoDe(c, 1);
        if (t.d == null) t.d = 0;
        const ad = activa(t);
        if (isFinite(ad)) D = Math.max(D, base + t.b + ad);
      }
    })(hijo(ctn, "childTnLst"), 0);
    return D;
  }

  // ================================================================== escribir p:timing (exportar)
  /**
   * El p:timing de una diapositiva con la lista de NOM, como lo escribe PowerPoint. spidDe(bid) → número de la
   * figura en el archivo (o null si no está: el efecto no se escribe). preset(x) → árbol de un efecto propio de NOM.
   * medios y bld: los de la importación (se conservan). Devuelve "" si no hay nada que escribir.
   */
  function timingXml(lista, spidDe, opc) {
    opc = opc || {};
    let id = 1, idEf = 0;
    const efs = [];
    const usados = new Set();
    // cambiar bid por spid; si el destino ya no está, el comportamiento no se escribe
    const cambia = (falta) => (n) => {
      const l = loc(n[0]);
      if (l === "spTgt" || l === "inkTgt") {
        const a = Object.assign({}, A(n));
        const sp = a.bid != null ? spidDe(a.bid) : a.spid;
        if (sp == null) { falta.v = true; return null; }
        a.spid = String(sp); delete a.bid; usados.add(a.spid);
        return [n[0], a, H(n)];
      }
      if (l === "cTn") { const a = Object.assign({ id: String(++id) }, A(n)); a.id = String(id); if (a.presetClass != null) idEf = id; return [n[0], a, H(n)]; }
      // «ocultar después de la animación» se refiere al número del efecto
      if (l === "tn" && A(n).val === "@ef") return [n[0], Object.assign({}, A(n), { val: String(idEf) }), H(n)];
      return null === n ? null : undefined;
    };
    const cambiaTodo = (falta) => { const f = cambia(falta); return (n) => { const r = f(n); return r === undefined ? n : r; }; };
    const efecto = (x, nt) => {
      const pp = x.ppt || (opc.preset ? opc.preset(x) : null);
      if (!pp || !pp.c) return "";
      const c = copia(pp.c);
      const d0 = +pp.d0 || 0, f = d0 > 0 && x.dur > 0 && Math.abs(x.dur - d0) > 0.5 ? x.dur / d0 : 1;
      if (f !== 1) escalarTiempos(c, f);
      // el destino del efecto: si se cambió el bloque (efecto propio), todos sus spTgt van al bloque del efecto
      (function rec(m) { if (!Array.isArray(m)) return; if ((loc(m[0]) === "spTgt") && (A(m).bid == null || A(m).bid === "@") && A(m).spid == null) m[1] = Object.assign({}, A(m), { bid: x.id }); H(m).forEach(rec); })(c);
      const falta = { v: false };
      const a = Object.assign({}, A(c), { nodeType: nt });
      const resto = H(c).filter((h) => !(Array.isArray(h) && loc(h[0]) === "stCondLst"));
      const nuevo = [c[0], a, [["p:stCondLst", {}, [["p:cond", { delay: String(Math.round(+x.retraso || 0)) }]]], ...resto]];
      const xml = aXml(nuevo, cambiaTodo(falta));
      if (falta.v && !xml.includes("spid=")) return "";
      return `<p:par>${xml}</p:par>`;
    };
    // la secuencia principal y las de cada desencadenador
    const armarSeq = (items, disp) => {
      const grupos = [];
      let g = null, s = null;
      items.forEach((x) => {
        const ini = x.inicio || "clic";
        if (ini === "clic" || !g) { g = { auto: ini !== "clic" && !grupos.length && !disp, subs: [] }; grupos.push(g); s = null; }
        if (!s || ini === "despues") { s = { efs: [], d: 0 }; g.subs.push(s); }
        s.efs.push(x);
      });
      const nt = (x, i) => (i === 0 && (x.inicio || "clic") === "clic" ? "clickEffect" : (x.inicio === "despues" && i === 0 ? "afterEffect" : x.inicio === "despues" ? "afterEffect" : "withEffect"));
      let out = "";
      grupos.forEach((g2) => {
        let t = 0, subs = "";
        const gid = ++id;
        g2.subs.forEach((s2, si) => {
          const sid = ++id;
          let fin = 0, cuerpo = "";
          s2.efs.forEach((x, i) => {
            const ini = x.inicio || "clic";
            const tipoN = i === 0 ? (ini === "clic" ? "clickEffect" : ini === "despues" ? "afterEffect" : (si === 0 ? "withEffect" : "afterEffect")) : "withEffect";
            const xml = efecto(x, tipoN);
            if (!xml) return;
            cuerpo += xml;
            const pp = x.ppt || (opc.preset ? opc.preset(x) : null);
            const dn = pp ? (+pp.d0 || 0) * ((+pp.d0 && x.dur) ? x.dur / +pp.d0 : 1) : 0;
            const rep = pp ? tiempoDe(pp.c, 1) : { rep: 1, rev: false };
            fin = Math.max(fin, (+x.retraso || 0) + dn * (isFinite(rep.rep) ? rep.rep : 1) * (rep.rev ? 2 : 1));
          });
          void nt;
          if (cuerpo) subs += `<p:par><p:cTn id="${sid}" fill="hold"><p:stCondLst><p:cond delay="${Math.round(t)}"/></p:stCondLst><p:childTnLst>${cuerpo}</p:childTnLst></p:cTn></p:par>`;
          t += fin;
        });
        if (!subs) return;
        const inicio = disp ? `<p:cond delay="0"/>` : g2.auto ? `<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>` : `<p:cond delay="indefinite"/>`;
        out += `<p:par><p:cTn id="${gid}" fill="hold"><p:stCondLst>${inicio}</p:stCondLst><p:childTnLst>${subs}</p:childTnLst></p:cTn></p:par>`;
      });
      return out;
    };
    const principal = lista.filter((x) => !x.disp);
    let cuerpo = "";
    id = 2;
    const sp = armarSeq(principal, null);
    if (sp) cuerpo += `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${sp}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
    const disps = new Map();
    lista.filter((x) => x.disp).forEach((x) => { if (!disps.has(x.disp)) disps.set(x.disp, []); disps.get(x.disp).push(x); });
    disps.forEach((items, disp) => {
      const spD = spidDe(disp);
      if (spD == null) return;
      const sid = ++id;
      const sq = armarSeq(items, disp);
      if (!sq) return;
      const cond = `<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${spD}"/></p:tgtEl></p:cond>`;
      cuerpo += `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${sid}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst>${cond}</p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>${sq}</p:childTnLst></p:cTn><p:nextCondLst>${cond}</p:nextCondLst></p:seq>`;
    });
    (opc.medios || []).forEach((m) => { const falta = { v: false }; const x = aXml(copia(m), cambiaTodo(falta)); if (!falta.v) cuerpo += x; });
    if (!cuerpo) return "";
    // bldLst: lo de la importación que siga vivo y una entrada por cada figura nueva con texto
    const blds = [];
    const yaB = new Set();
    (opc.bld || []).forEach((b) => {
      const a = Object.assign({}, A(b)), s2 = a.bid != null ? spidDe(a.bid) : a.spid;
      if (s2 == null || !usados.has(String(s2))) return;
      a.spid = String(s2); delete a.bid;
      yaB.add(a.spid + "|" + (a.grpId || 0));
      blds.push(aXml([b[0], a, H(b)]));
    });
    lista.forEach((x) => {
      const s2 = x.id != null ? spidDe(x.id) : null;
      if (s2 == null || !usados.has(String(s2))) return;
      const g = (x.ppt && A(x.ppt.c).grpId) || "0";
      const k = s2 + "|" + g;
      if (yaB.has(k)) return;
      yaB.add(k);
      // por párrafos: build="p" (y rev="1" si van al revés), como PowerPoint
      if (x.par != null) blds.push(`<p:bldP spid="${s2}" grpId="${g}" build="p"${x.parAtras ? ' rev="1"' : ""}${lista.some((y) => y.grupoP && y.grupoP === x.grupoP && y.par === -1) ? ' animBg="1" autoUpdateAnimBg="0"' : ""}/>`);
      else if (opc.conTexto && opc.conTexto(x.id)) blds.push(`<p:bldP spid="${s2}" grpId="${g}" animBg="1"/>`);
    });
    const tm = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${cuerpo}</p:childTnLst></p:cTn></p:par></p:tnLst>${blds.length ? `<p:bldLst>${blds.join("")}</p:bldLst>` : ""}</p:timing>`;
    // Lo de PowerPoint 2010 (p14: el rebote al final) va como lo escribe PowerPoint: en mc:AlternateContent, y para
    // los programas de antes el mismo p:timing sin esos atributos.
    if (!/\sp14:/.test(tm)) return tm;
    const sin = tm.replace(/\sp14:[A-Za-z]+="[^"]*"/g, "");
    return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">${tm}</mc:Choice><mc:Fallback>${sin}</mc:Fallback></mc:AlternateContent>`;
  }
  /** Multiplica por f los tiempos (duraciones y retrasos) dentro de un efecto. */
  function escalarTiempos(n, f) {
    (function rec(m) {
      if (!Array.isArray(m)) return;
      // (los «set» de 1 ms se quedan de 1 ms: son un instante, no una duración)
      if (loc(m[0]) === "cTn") { const a = m[1] = Object.assign({}, A(m)); if (a.dur && a.dur !== "indefinite" && +a.dur > 1) a.dur = String(Math.max(1, Math.round(+a.dur * f))); if (a.repeatDur && a.repeatDur !== "indefinite") a.repeatDur = String(Math.round(+a.repeatDur * f)); }
      if (loc(m[0]) === "cond" && A(m).delay && A(m).delay !== "indefinite") m[1] = Object.assign({}, A(m), { delay: String(Math.round(+A(m).delay * f)) });
      if (loc(m[0]) === "tmAbs") m[1] = Object.assign({}, A(m), { val: String(Math.round(+A(m).val * f)) });
      H(m).forEach(rec);
    })(n);
    return n;
  }

  // ================================================================== efectos de la biblioteca (los de PowerPoint)
  /** Los 17 efectos que tenía NOM antes, con su efecto de PowerPoint equivalente. */
  const LEGADO = { "e:aparecer": "entr:10:0", "e:deslizar": "entr:2:4", "e:deslizar-arriba": "entr:2:1", "e:deslizar-izq": "entr:2:8", "e:deslizar-der": "entr:2:2",
    "e:zoom": "entr:53:16", "e:rebotar": "entr:26:0", "e:girar": "entr:31:0", "n:pulso": "emph:26:0", "n:sacudir": "emph:32:0", "n:vuelta": "emph:8:0",
    "n:crecer": "emph:6:0", "n:parpadear": "emph:27:0", "x:desvanecer": "exit:10:0", "x:deslizar": "exit:2:4", "x:alejar": "exit:53:32", "x:girar": "exit:31:0" };
  const presets = () => (NOM.animOfficeDatos && NOM.animOfficeDatos.presets) || {};
  /** La clave de la biblioteca de un efecto de la lista ("entr:2:4"), o null. */
  const clavePreset = (x) => (x && /^ppt:/.test(x.efecto || "") ? x.efecto.slice(4) : x ? LEGADO[x.tipo + ":" + x.efecto] || null : null);
  /** El árbol de PowerPoint de un efecto: el suyo, o el de la biblioteca. */
  function presetDe(x) {
    if (!x) return null;
    if (x.ppt) return x.ppt;
    const k = clavePreset(x), p = k && presets()[k];
    return p ? { c: p.c, d0: p.d0 } : null;
  }
  /** Nombre para el panel: «Desplazar hacia dentro (desde abajo)». */
  function nombreEfecto(x) {
    const k = clavePreset(x);
    const p = k && presets()[k];
    if (p) return p.n + (p.dir ? ` (${p.dir.toLowerCase()})` : "");
    if (x && x.ppt && A(x.ppt.c).presetClass === "path" && +A(x.ppt.c).presetID === 0) return "Trayectoria personalizada";
    if (x && x.ppt) { const a = A(x.ppt.c); const q = presets()[`${a.presetClass}:${a.presetID}:0`] || Object.entries(presets()).find(([kk]) => kk.startsWith(`${a.presetClass}:${a.presetID}:`)); const pp = Array.isArray(q) ? q[1] : q; return pp ? pp.n : "Efecto de PowerPoint"; }
    return null;
  }

  // ================================================================== opciones de efecto (lo del cuadro «Opciones de efectos»)
  // Cada opción es lo mismo que escribe PowerPoint en el árbol del efecto (medido en a02-opciones):
  // por letra/palabra: <p:iterate type="lt|wd"><p:tmPct val="10000"/></p:iterate>; repetir n: repeatCount="n000"
  // (hasta el siguiente clic: "indefinite" con endCondLst onNext; hasta el final de la diapositiva: "indefinite");
  // autorreversa: autoRev="1"; rebobinar al terminar: fill="remove"; inicio y fin suaves: accel/decel (fracción de la
  // duración × 100000); rebote al final: p14:presetBounceEnd en el efecto y p14:bounceEnd en cada comportamiento que
  // mueve; después de la animación: subTnLst (atenuar: animClr de ppt_c al siguiente clic; ocultar: set de
  // visibilidad al terminar o al siguiente clic); por párrafos: un efecto por párrafo con <p:txEl><p:pRg/></p:txEl>.
  const ORDEN_CTN = ["stCondLst", "endCondLst", "endSync", "iterate", "childTnLst", "subTnLst"];
  function ponerHijo(c, nodo) {
    const l = loc(nodo[0]);
    const hs = H(c).filter((h) => !(Array.isArray(h) && loc(h[0]) === l));
    const k = ORDEN_CTN.indexOf(l);
    let i = hs.findIndex((h) => Array.isArray(h) && ORDEN_CTN.indexOf(loc(h[0])) > k);
    if (i < 0) i = hs.length;
    hs.splice(i, 0, nodo);
    c[2] = hs;
  }
  function quitarHijo(c, l) { if (Array.isArray(c[2])) c[2] = c[2].filter((h) => !(Array.isArray(h) && loc(h[0]) === l)); }
  const MUEVE = /^(anim|animMotion|animScale|animRot)$/;
  /** Recorre los comportamientos (las hojas) de un efecto. */
  function cadaComp(c, fn) {
    (function rec(lst) {
      for (const n of hijosDe(lst)) {
        const l = loc(n[0]);
        if (l === "par" || l === "seq" || l === "excl") { const k = hijo(n, "cTn"); if (k) rec(hijo(k, "childTnLst")); continue; }
        fn(n, l);
      }
    })(hijo(c, "childTnLst"));
  }
  const hexDe = (n) => { const v = n && hijosDe(n)[0]; if (!v) return null; if (A(v)._c) return A(v)._c; return loc(v[0]) === "srgbClr" ? "#" + A(v).val : null; };
  /** Las opciones de un efecto de la lista, leídas de su árbol. */
  function opcionesDe(x) {
    const pp = presetDe(x);
    if (!pp || !pp.c) return null;
    const c = pp.c, a = A(c), it = hijo(c, "iterate"), fin = hijo(c, "endCondLst");
    const op = {
      texto: it ? ({ lt: "letra", wd: "palabra" }[A(it).type] || "todo") : "todo",
      pct: it && hijo(it, "tmPct") ? Math.round(+A(hijo(it, "tmPct")).val / 1000) : 10,
      atras: !!(it && A(it).backwards === "1"),
      repetir: a.repeatCount == null ? 1 : a.repeatCount === "indefinite" ? (fin && hijosDe(fin, "cond").some((q) => A(q).evt === "onNext") ? "clic" : "fin") : Math.max(1, +a.repeatCount / 1000),
      autoRev: a.autoRev === "1" || a.autoRev === "true",
      rebobinar: a.fill === "remove",
      suaveIni: (+a.accel || 0) / 100000, suaveFin: (+a.decel || 0) / 100000,
      rebote: (+a["p14:presetBounceEnd"] || 0) / 100000,
      despues: null, giro: null, escala: null, color: null,
    };
    const sub = hijo(c, "subTnLst");
    hijosDe(sub).forEach((n) => {
      const cb = hijo(n, "cBhvr"), ct = cb && hijo(cb, "cTn");
      if (loc(n[0]) === "animClr") op.despues = { color: hexDe(hijo(n, "to")) || "#C8C8C8" };
      else if (loc(n[0]) === "set") op.despues = A(ct).masterRel === "nextClick" ? "ocultarClic" : "ocultar";
    });
    // los parámetros que PowerPoint deja cambiar: la cantidad en Girar, el tamaño en Aumentar o encoger y el color en
    // los énfasis de color (en los demás efectos esos valores son parte del efecto)
    const enfasis = a.presetClass === "emph", pid = +a.presetID;
    if (enfasis) cadaComp(c, (n, l) => {
      if (pid !== 8) return;
      if (l === "animRot" && op.giro == null && A(n).by != null) op.giro = Math.round(+A(n).by / 60000);
    });
    if (enfasis) cadaComp(c, (n, l) => {
      if (l === "animScale" && pid === 6 && op.escala == null && (hijo(n, "to") || hijo(n, "by"))) op.escala = Math.round(+A(hijo(n, "to") || hijo(n, "by")).x / 1000);
      if (l === "animClr" && op.color == null && hijo(n, "to")) op.color = hexDe(hijo(n, "to"));
    });
    return op;
  }
  /** Un efecto de la lista con otras opciones (devuelve una copia; el árbol queda en «ppt»). */
  function conOpciones(x, cambios) {
    const pp = presetDe(x);
    if (!pp || !pp.c) return x;
    cambios = cambios || {};
    // solo se toca lo que cambió: el resto del árbol queda como lo escribió PowerPoint (Balancear y Onda giran
    // distinto en cada tramo, los colores del tema siguen siendo del tema, un tmAbs sigue siendo tmAbs…)
    const orig = opcionesDe(x);
    const op = Object.assign({}, orig, cambios);
    const cambia = (...ks) => ks.some((k) => k in cambios && JSON.stringify(cambios[k]) !== JSON.stringify(orig[k]));
    const c = copia(pp.c), a = c[1] = Object.assign({}, A(c));
    // por letra o palabra
    if (cambia("texto", "pct", "atras")) {
      quitarHijo(c, "iterate");
      if (op.texto === "letra" || op.texto === "palabra") {
        const ia = { type: op.texto === "letra" ? "lt" : "wd" };
        if (op.atras) ia.backwards = "1";
        ponerHijo(c, ["p:iterate", ia, [["p:tmPct", { val: String(Math.round(Math.max(0, Math.min(100, +op.pct || 0)) * 1000)) }]]]);
      }
    }
    // repetir
    if (cambia("repetir")) {
      delete a.repeatCount; quitarHijo(c, "endCondLst");
      if (op.repetir === "clic" || op.repetir === "fin") {
        a.repeatCount = "indefinite";
        if (op.repetir === "clic") ponerHijo(c, ["p:endCondLst", {}, [["p:cond", { evt: "onNext", delay: "0" }, [["p:tgtEl", {}, [["p:sldTgt", {}]]]]]]]);
      } else if (+op.repetir > 1) a.repeatCount = String(Math.round(+op.repetir * 1000));
    }
    if (cambia("autoRev")) { if (op.autoRev) a.autoRev = "1"; else delete a.autoRev; }
    if (cambia("rebobinar")) a.fill = op.rebobinar ? "remove" : "hold";
    const pon100 = (k, v) => { if (v > 0) a[k] = String(Math.round(Math.min(1, v) * 100000)); else delete a[k]; };
    if (cambia("suaveIni")) pon100("accel", op.suaveIni);
    if (cambia("suaveFin")) pon100("decel", op.suaveFin);
    const rb = op.rebote > 0 ? String(Math.round(Math.min(0.95, op.rebote) * 100000)) : null;
    if (cambia("rebote")) { if (rb) a["p14:presetBounceEnd"] = rb; else delete a["p14:presetBounceEnd"]; }
    cadaComp(c, (n, l) => {
      if (cambia("rebote") && MUEVE.test(l)) { n[1] = Object.assign({}, A(n)); if (rb) n[1]["p14:bounceEnd"] = rb; else delete n[1]["p14:bounceEnd"]; }
      if (cambia("giro") && l === "animRot" && A(n).by != null) n[1] = Object.assign({}, A(n), { by: String(Math.round(+op.giro * 60000)) });
      // Aumentar o encoger: PowerPoint deja by="150000" y escribe el tamaño elegido en «to» (que es lo que manda)
      if (cambia("escala") && l === "animScale" && (hijo(n, "by") || hijo(n, "to"))) {
        const v = String(Math.round(op.escala * 1000));
        quitarHijo(n, "to");
        n[2] = H(n).concat([["p:to", { x: v, y: v }]]);
      }
      if (cambia("color") && op.color && l === "animClr" && hijo(n, "to")) { const to = hijo(n, "to"); to[2] = [["a:srgbClr", { val: op.color.replace("#", "").toUpperCase() }]]; }
    });
    if (!cambia("despues")) {
      const y0 = Object.assign({}, x, { ppt: { c, d0: pp.d0 } });
      if (!/^ppt:/.test(y0.efecto || "")) y0.efecto = "ppt:" + (clavePreset(x) || `${a.presetClass}:${a.presetID}:${a.presetSubtype || 0}`);
      return y0;
    }
    // después de la animación
    quitarHijo(c, "subTnLst");
    const tg = () => ["p:tgtEl", {}, [["p:spTgt", { bid: "@" }]]];
    if (op.despues && op.despues.color) {
      ponerHijo(c, ["p:subTnLst", {}, [["p:animClr", { clrSpc: "rgb", dir: "cw" }, [["p:cBhvr", { override: "childStyle" }, [["p:cTn", { dur: "1", fill: "hold", display: "0", masterRel: "nextClick", afterEffect: "1" }], tg(), ["p:attrNameLst", {}, [["p:attrName", {}, ["ppt_c"]]]]]], ["p:to", {}, [["a:srgbClr", { val: op.despues.color.replace("#", "").toUpperCase() }]]]]]]]);
    } else if (op.despues === "ocultar" || op.despues === "ocultarClic") {
      const ct = op.despues === "ocultar"
        ? ["p:cTn", { dur: "1", fill: "hold", display: "0", masterRel: "sameClick", afterEffect: "1" }, [["p:stCondLst", {}, [["p:cond", { evt: "end", delay: "0" }, [["p:tn", { val: "@ef" }]]]]]]]
        : ["p:cTn", { dur: "1", fill: "hold", display: "0", masterRel: "nextClick", afterEffect: "1" }];
      ponerHijo(c, ["p:subTnLst", {}, [["p:set", {}, [["p:cBhvr", { override: "childStyle" }, [ct, tg(), ["p:attrNameLst", {}, [["p:attrName", {}, ["style.visibility"]]]]]], ["p:to", {}, [["p:strVal", { val: "hidden" }]]]]]]]);
    }
    const y = Object.assign({}, x, { ppt: { c, d0: pp.d0 } });
    if (!/^ppt:/.test(y.efecto || "")) y.efecto = "ppt:" + (clavePreset(x) || `${a.presetClass}:${a.presetID}:${a.presetSubtype || 0}`);
    return y;
  }
  /** Las variantes (direcciones) de un efecto: [{ clave, dir }]. */
  function variantes(x) {
    const k = clavePreset(x) || (x && x.ppt ? `${A(x.ppt.c).presetClass}:${A(x.ppt.c).presetID}:${A(x.ppt.c).presetSubtype || 0}` : null);
    if (!k) return [];
    const [cl, id] = k.split(":");
    return Object.entries(presets()).filter(([kk]) => kk.startsWith(cl + ":" + id + ":")).map(([clave, p]) => ({ clave, dir: p.dir })).filter((v) => v.dir);
  }
  /** El mismo efecto con otra clave de la biblioteca (otra dirección o efecto), conservando sus opciones. */
  function cambiarEfecto(x, clave) {
    const p = presets()[clave];
    if (!p) return x;
    const op = x && presetDe(x) ? opcionesDe(x) : null;
    const tipo = { entr: "e", exit: "x", emph: "n", path: "r" }[clave.split(":")[0]] || "n";
    const y = Object.assign({}, x, { tipo, efecto: "ppt:" + clave, dur: x && x.dur && x.efecto && clavePreset(x) && clavePreset(x).split(":").slice(0, 2).join(":") === clave.split(":").slice(0, 2).join(":") ? x.dur : Math.max(1, +p.d0 || 500) });
    delete y.ppt;
    if (!op) return y;
    // (los parámetros propios de cada efecto, como el color o los grados, no pasan a otro efecto)
    const mismo = clavePreset(x) && clavePreset(x).split(":").slice(0, 2).join(":") === clave.split(":").slice(0, 2).join(":");
    const pasa = { texto: op.texto, pct: op.pct, atras: op.atras, repetir: op.repetir, autoRev: op.autoRev, rebobinar: op.rebobinar, suaveIni: op.suaveIni, suaveFin: op.suaveFin, rebote: op.rebote, despues: op.despues };
    if (mismo) Object.assign(pasa, { giro: op.giro, escala: op.escala, color: op.color });
    const base = opcionesDe(y);
    const dif = Object.keys(pasa).some((k) => JSON.stringify(pasa[k]) !== JSON.stringify(base[k]));
    return dif ? conOpciones(y, pasa) : y;
  }
  /** Un efecto nuevo de la biblioteca para un bloque. */
  function nuevoEfecto(id, clave, inicio) {
    const p = presets()[clave];
    if (!p) return null;
    const tipo = { entr: "e", exit: "x", emph: "n", path: "r" }[clave.split(":")[0]] || "n";
    // (los instantáneos, como Aparecer, no tienen duración: se quedan con la suya, 1 ms, como en PowerPoint)
    return { id, tipo, efecto: "ppt:" + clave, inicio: inicio || "clic", retraso: 0, dur: Math.max(1, +p.d0 || 500) };
  }
  /** Copia de un efecto para otro bloque (copiar animación): lo que apuntaba al bloque de antes apunta al nuevo. */
  function paraBloque(x, id) {
    const y = copia(x);
    (function rec(m) { if (!Array.isArray(m)) return; if (loc(m[0]) === "spTgt" && (A(m).bid === x.id || A(m).bid === "@")) { m[1] = Object.assign({}, A(m), { bid: "@" }); delete m[1].spid; } H(m).forEach(rec); })(y.ppt ? y.ppt.c : null);
    y.id = id; delete y.ids; delete y.spid; delete y.grupoP; delete y.par; delete y.parAtras; delete y.marcador;
    if (y.ppt) (function rec(m) { if (!Array.isArray(m)) return; if (loc(m[0]) === "spTgt") m[2] = H(m).filter((h) => !(Array.isArray(h) && loc(h[0]) === "txEl")); H(m).forEach(rec); })(y.ppt.c);
    return y;
  }
  /**
   * Por párrafos: un efecto por párrafo (como PowerPoint), cada uno con su txEl. indices: los párrafos con texto.
   * Devuelve la lista de efectos (el primero con el inicio del original; los demás, igual que él).
   */
  function porParrafos(x, indices, atras, conFondo) {
    const pp = presetDe(x);
    if (!pp || !pp.c || !indices.length) return [x];
    const g = "p" + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
    const orden = (conFondo ? [-1] : []).concat(atras ? indices.slice().reverse() : indices);
    // (con «Animar la forma adjunta» el primero es la figura sin su texto: <p:spTgt><p:bg/></p:spTgt>, como PowerPoint)
    return orden.map((k) => {
      const c = copia(pp.c);
      (function rec(m) {
        if (!Array.isArray(m)) return;
        if (loc(m[0]) === "spTgt" && (A(m).bid == null || A(m).bid === "@" || A(m).bid === x.id)) m[2] = k < 0 ? [["p:bg", {}]] : [["p:txEl", {}, [["p:pRg", { st: String(k), end: String(k) }]]]];
        H(m).forEach(rec);
      })(c);
      const y = Object.assign({}, x, { ppt: { c, d0: pp.d0 }, grupoP: g, par: k });
      if (atras) y.parAtras = true; else delete y.parAtras;
      if (!/^ppt:/.test(y.efecto || "")) y.efecto = "ppt:" + clavePreset(x);
      return y;
    });
  }
  /** Juntar un efecto por párrafos en uno solo (como un objeto). */
  function comoObjeto(x) {
    const y = Object.assign({}, x);
    delete y.grupoP; delete y.par; delete y.parAtras;
    if (y.ppt) { y.ppt = copia(y.ppt); (function rec(m) { if (!Array.isArray(m)) return; if (loc(m[0]) === "spTgt") m[2] = H(m).filter((h) => !(Array.isArray(h) && (loc(h[0]) === "txEl" || loc(h[0]) === "bg"))); H(m).forEach(rec); })(y.ppt.c); }
    return y;
  }
  /**
   * Trayectoria propia (dibujada): pts en px de la diapositiva desde el centro del bloque, T el tamaño de la
   * diapositiva. Se escribe como PowerPoint: presetClass="path" presetID="0", camino relativo en fracciones.
   */
  function trayectoriaPropia(id, pts, T, cerrada) {
    const f = (v) => String(Math.round(v * 10000) / 10000);
    const d = "M 0 0 " + pts.slice(1).map((p) => `L ${f(p[0] / T.w)} ${f(p[1] / T.h)}`).join(" ") + (cerrada ? " Z" : " E");
    const c = ["p:cTn", { presetID: "0", presetClass: "path", presetSubtype: "0", accel: "50000", decel: "50000", fill: "hold", grpId: "0" }, [["p:childTnLst", {}, [
      ["p:animMotion", { origin: "layout", path: d, pathEditMode: "relative", rAng: "0", ptsTypes: "A".repeat(pts.length) }, [["p:cBhvr", {}, [["p:cTn", { dur: "2000", fill: "hold" }], ["p:tgtEl", {}, [["p:spTgt", { bid: "@" }]]], ["p:attrNameLst", {}, [["p:attrName", {}, ["ppt_x"]], ["p:attrName", {}, ["ppt_y"]]]]]]]]]]]];
    return { id, tipo: "r", efecto: "ppt:path:0:0", inicio: "clic", retraso: 0, dur: 2000, ppt: { c, d0: 2000 } };
  }
  /** El camino (en fracciones de la diapositiva) de un efecto de trayectoria, para dibujarlo en el editor. */
  function caminoDe(x) {
    const pp = presetDe(x);
    let d = null;
    if (pp && pp.c) cadaComp(pp.c, (n, l) => { if (!d && l === "animMotion" && A(n).path) d = A(n).path; });
    return d ? leerCamino(d) : null;
  }

  NOM.animOffice = { rebote, reproductor, presetDe, clavePreset, nombreEfecto, LEGADO, opcionesDe, conOpciones, variantes, cambiarEfecto, nuevoEfecto, paraBloque, porParrafos, comoObjeto, trayectoriaPropia, caminoDe, parrafosDe, contTexto, leerTiming, timingXml, deDom, aXml, formula, valorNum, mascara, leerCamino, puntoCamino, tiempoDe, local, activa, acelera, duracionNatural, escalarTiempos, colorNodo, loc, A, H, hijo, hijosDe, copia };
})();

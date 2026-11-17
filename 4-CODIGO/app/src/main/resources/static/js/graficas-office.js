/* NOM · Gráficas de Office (DrawingML «c:chartSpace») dibujadas como Office. El mismo dibujante para
   Excel, Word y PowerPoint: se lee el XML de la gráfica con su estilo (colores, letras, ejes, leyenda, rótulos)
   y se pinta en SVG con las reglas de Office: escala automática de los ejes, ancho y separación de las barras,
   leyenda abajo, títulos, cuadrícula, rebanadas desde las 12 en punto…

   NOM.graficaOffice.leer(textoXml, ctx) → modelo (JSON; se guarda en el documento)
     ctx.color(nodoConRelleno) → "#RRGGBB" o null   (resuelve los colores del tema del archivo)
     ctx.fuente(nombre) → nombre de la letra         ("+mn-lt" → la del tema)
     ctx.acentos → ["#RRGGBB" × 6]                   (colores de las series que no dicen el suyo)
   NOM.graficaOffice.svg(modelo, w, h, opc) → texto SVG de w × h
     opc.pt: píxeles por punto (1 en una diapositiva 16:9; 4/3 en una hoja a 96 ppp)
     opc.valores(ref) → arreglo de valores vivos de un rango («Hoja1!$B$2:$B$7»), o null para usar los guardados
     opc.formato(valor, codigo) → texto (formato de número de Excel) */
(() => {
  const n2 = (x) => Math.round(x * 100) / 100;
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  const TIPOS = { barChart: "barra", bar3DChart: "barra", lineChart: "linea", line3DChart: "linea", stockChart: "linea", areaChart: "area", area3DChart: "area",
    pieChart: "pastel", pie3DChart: "pastel", ofPieChart: "pastel", doughnutChart: "anillo", scatterChart: "dispersion", bubbleChart: "dispersion", radarChart: "radar" };

  // ================================================================ lectura
  function leer(xml, ctx) {
    ctx = ctx || {};
    const d = new DOMParser().parseFromString(xml, "application/xml");
    const hs = (n, nombre) => [...(n ? n.children : [])].filter((c) => !nombre || c.localName === nombre);
    const h1 = (n, nombre) => hs(n, nombre)[0] || null;
    const todos = (n, nombre) => (n ? [...n.getElementsByTagNameNS("*", nombre)] : []);
    const at = (n, a) => (n ? n.getAttribute(a) : null);
    const val = (n, nombre, def) => { const e = h1(n, nombre); return e && at(e, "val") != null ? at(e, "val") : def; };
    const si = (n, nombre, def) => { const v = val(n, nombre, null); return v == null ? def : v === "1" || v === "true"; };
    const color = (n) => (n && ctx.color ? ctx.color(n) : null);
    const relleno = (spPr) => {
      if (!spPr) return null;
      if (h1(spPr, "noFill")) return "none";
      const sf = h1(spPr, "solidFill"); if (sf) return color(sf);
      const gf = h1(spPr, "gradFill"); if (gf) { const gs = todos(gf, "gs"); return gs.length ? color(gs[0]) : null; }
      const pf = h1(spPr, "pattFill"); if (pf) return color(h1(pf, "fgClr"));
      return null;
    };
    // transparencia de un color (a:alpha) y degradados: lineal con su ángulo o «desde el centro» con su foco
    const alfa = (n) => { const c = n && [...n.children].find((x) => /Clr$/.test(x.localName)); const a = c && [...c.children].find((x) => x.localName === "alpha"); return a ? Math.max(0, Math.min(1, +at(a, "val") / 100000)) : null; };
    const degradado = (spPr) => {
      const gf = spPr && h1(spPr, "gradFill");
      if (!gf) return null;
      const p = todos(gf, "gs").map((g) => [Math.max(0, Math.min(1, +(at(g, "pos") || 0) / 100000)), color(g), alfa(g)]).filter((q) => q[1]).sort((a, b) => a[0] - b[0]);
      if (!p.length) return null;
      const ln = h1(gf, "lin"), pa = h1(gf, "path"), fr = pa && h1(pa, "fillToRect");
      if (pa) return { radial: true, p, foco: fr ? ["l", "t", "r", "b"].map((k) => +(at(fr, k) || 0) / 100000) : [0.5, 0.5, 0.5, 0.5] };
      return { ang: ln ? +(at(ln, "ang") || 0) / 60000 : 90, p };
    };
    const linea = (spPr) => {
      const ln = spPr && h1(spPr, "ln");
      if (!ln) return null;
      if (h1(ln, "noFill")) return { no: 1 };
      const o = {};
      if (at(ln, "w")) o.w = +at(ln, "w") / 12700;
      const sf = h1(ln, "solidFill"); if (sf) o.c = color(sf);
      const gu = h1(ln, "prstDash"); if (gu && at(gu, "val") !== "solid") o.guion = at(gu, "val");
      if (at(ln, "cap") === "rnd") o.redondo = true;
      return o;
    };
    const letra = (txPr) => {
      const r = txPr && todos(txPr, "defRPr")[0];
      if (!r) return null;
      const o = {};
      if (at(r, "sz")) o.tam = +at(r, "sz") / 100;
      if (at(r, "b") != null) o.b = at(r, "b") === "1";
      if (at(r, "i") != null) o.i = at(r, "i") === "1";
      if (at(r, "cap") === "all" || at(r, "cap") === "small") o.cap = at(r, "cap");
      const sf = h1(r, "solidFill"); if (sf) o.c = color(sf);
      const la = h1(r, "latin"); if (la && at(la, "typeface")) o.f = ctx.fuente ? ctx.fuente(at(la, "typeface")) : at(la, "typeface");
      const bp = txPr && h1(txPr, "bodyPr"); if (bp && at(bp, "rot") != null && +at(bp, "rot") !== -60000000) o.rot = +at(bp, "rot") / 60000;
      return o;
    };
    const textoRico = (tx) => {
      const rich = tx && h1(tx, "rich");
      if (rich) return hs(rich, "p").map((p) => todos(p, "t").map((t) => t.textContent).join("")).join("\n");
      const sr = tx && h1(tx, "strRef");
      if (sr) return todos(sr, "v").map((v) => v.textContent).join(" ");
      const v = tx && h1(tx, "v");
      return v ? v.textContent : null;
    };
    const letraDeRico = (tx) => { const r = tx && todos(tx, "rPr")[0], dr = tx && todos(tx, "defRPr")[0]; const o = letra(tx && h1(tx, "rich")) || {}; if (r) { if (at(r, "sz")) o.tam = +at(r, "sz") / 100; if (at(r, "b") != null) o.b = at(r, "b") === "1"; const sf = h1(r, "solidFill"); if (sf) o.c = color(sf); } return Object.keys(o).length ? o : dr ? letra(tx) : null; };
    const cache = (ref) => {
      // valores guardados (numCache, strCache, numLit, strLit, multiLvlStrRef: el último nivel)
      if (!ref) return { v: [], f: null, fmt: null };
      const c = todos(ref, "numCache")[0] || todos(ref, "strCache")[0] || todos(ref, "numLit")[0] || todos(ref, "strLit")[0] || todos(ref, "lvl")[0];
      const n = c ? +(val(c, "ptCount", 0)) : 0, v = new Array(n).fill(null);
      if (c) hs(c, "pt").forEach((p) => { const i = +at(p, "idx"); const x = (h1(p, "v") || {}).textContent; if (i >= 0 && i < 100000) v[i] = x; });
      const fm = c && h1(c, "formatCode");
      return { v, f: (todos(ref, "f")[0] || {}).textContent || null, fmt: fm ? fm.textContent : null };
    };
    const disposicion = (n) => {
      const ml = n && todos(h1(n, "layout"), "manualLayout")[0];
      if (!ml) return null;
      const o = { dentro: val(ml, "layoutTarget", "outer") === "inner" };
      ["x", "y", "w", "h"].forEach((k) => { const v = val(ml, k, null); if (v != null) o[k] = +v; });
      return o.x != null || o.w != null ? o : null;
    };
    const rotulos = (n) => {
      const dl = n && h1(n, "dLbls");
      if (!dl) return null;
      const o = { val: si(dl, "showVal", false), pct: si(dl, "showPercent", false), cat: si(dl, "showCatName", false), ser: si(dl, "showSerName", false), pos: val(dl, "dLblPos", null) };
      const nf = h1(dl, "numFmt"); if (nf && at(nf, "sourceLinked") !== "1") o.fmt = at(nf, "formatCode");
      const l = letra(h1(dl, "txPr")); if (l) o.letra = l;
      return o.val || o.pct || o.cat || o.ser ? o : null;
    };

    const cs = d.documentElement, ch = h1(cs, "chart"), pa = ch && h1(ch, "plotArea");
    const m = { tipos: [], ejes: {}, letra: letra(h1(cs, "txPr")) || {}, lengua: val(cs, "lang", "en-US") };
    m.area = { relleno: relleno(h1(cs, "spPr")), linea: linea(h1(cs, "spPr")), redondo: si(cs, "roundedCorners", false), degr: degradado(h1(cs, "spPr")) };
    // vista 3D: giro en X y en Y, profundidad (% del ancho de una categoría por fila) y ejes en ángulo recto
    const v3 = ch && h1(ch, "view3D");
    if (v3) m.vista3D = { rotX: +val(v3, "rotX", 15), rotY: +val(v3, "rotY", 20), prof: +val(v3, "depthPercent", 100), alto: +val(v3, "hPercent", 100), recto: si(v3, "rAngAx", true), persp: +val(v3, "perspective", 30) };
    m.trazado = { relleno: relleno(h1(pa, "spPr")), linea: linea(h1(pa, "spPr")), disp: disposicion(pa), degr: degradado(h1(pa, "spPr")) };
    // tabla de datos debajo de la gráfica
    const dt = h1(pa, "dTable");
    if (dt) m.tablaDatos = { h: si(dt, "showHorzBorder", true), v: si(dt, "showVertBorder", true), contorno: si(dt, "showOutline", true), claves: si(dt, "showKeys", true), linea: linea(h1(dt, "spPr")), letra: letra(h1(dt, "txPr")) };
    // título: el escrito, o el de la serie (una sola serie) o «Título del gráfico»
    const ti = h1(ch, "title");
    if (ti) m.titulo = { texto: textoRico(h1(ti, "tx")), letra: Object.assign({}, letra(h1(ti, "txPr")) || {}, letraDeRico(h1(ti, "tx")) || {}), sobre: si(ti, "overlay", false), disp: disposicion(ti) };
    m.sinTitulo = si(ch, "autoTitleDeleted", false);
    // leyenda
    const le = h1(ch, "legend");
    if (le) m.leyenda = { pos: val(le, "legendPos", "r"), letra: letra(h1(le, "txPr")) || {}, sobre: si(le, "overlay", false), disp: disposicion(le), borrar: hs(le, "legendEntry").filter((e) => si(e, "delete", false)).map((e) => +val(e, "idx", -1)) };
    // ejes
    ["catAx", "valAx", "dateAx", "serAx"].forEach((tipo) => hs(pa, tipo).forEach((ax) => {
      const sc = h1(ax, "scaling"), nf = h1(ax, "numFmt");
      const e = { tipo, id: val(ax, "axId"), cruza: val(ax, "crossAx"), borrado: si(ax, "delete", false), pos: val(ax, "axPos", "b"),
        min: val(sc, "min", null), max: val(sc, "max", null), inv: val(sc, "orientation", "minMax") === "maxMin", log: val(sc, "logBase", null),
        unidad: val(ax, "majorUnit", null), fmt: nf ? at(nf, "formatCode") : null, ligado: nf ? at(nf, "sourceLinked") === "1" : true,
        rejilla: h1(ax, "majorGridlines") ? linea(h1(h1(ax, "majorGridlines"), "spPr")) || {} : null,
        rejillaMenor: h1(ax, "minorGridlines") ? linea(h1(h1(ax, "minorGridlines"), "spPr")) || {} : null,
        linea: linea(h1(ax, "spPr")), letra: letra(h1(ax, "txPr")) || {}, rotulos: val(ax, "tickLblPos", "nextTo"),
        marcas: val(ax, "majorTickMark", "out"), entre: val(ax, "crossBetween", null), cruces: val(ax, "crosses", "autoZero"), cruzaEn: val(ax, "crossesAt", null),
        salto: val(ax, "tickLblSkip", null) };
      const tt = h1(ax, "title"); if (tt) e.titulo = { texto: textoRico(h1(tt, "tx")) || "", letra: Object.assign({}, letra(h1(tt, "txPr")) || {}, letraDeRico(h1(tt, "tx")) || {}) };
      ["min", "max", "unidad", "cruzaEn"].forEach((k) => { if (e[k] != null) e[k] = +e[k]; });
      m.ejes[e.id] = e;
    }));
    // grupos de series (una gráfica combinada trae varios)
    hs(pa).forEach((g) => {
      const tipo = TIPOS[g.localName];
      if (!tipo) return;
      const gr = { tipo, dir: val(g, "barDir", "col"), grupo: val(g, "grouping", tipo === "barra" ? "clustered" : "standard"), varia: si(g, "varyColors", tipo === "pastel" || tipo === "anillo"),
        hueco: +val(g, "gapWidth", 150), solape: +val(g, "overlap", 0), agujero: +val(g, "holeSize", 50), giro: +val(g, "firstSliceAng", 0),
        ejes: hs(g, "axId").map((a) => at(a, "val")), rotulos: rotulos(g), marcadores: si(g, "marker", true), estiloDisp: val(g, "scatterStyle", "lineMarker"),
        series: [] };
      if (g.localName === "bar3DChart" || g.localName === "pie3DChart" || g.localName === "line3DChart" || g.localName === "area3DChart") gr.tresD = true;
      if (g.localName === "bar3DChart") gr.huecoProf = +val(g, "gapDepth", 150);
      hs(g, "ser").forEach((s) => {
        const sp = h1(s, "spPr"), mk = h1(s, "marker");
        const se = { idx: +val(s, "idx", 0), orden: +val(s, "order", 0), nombre: textoRico(h1(s, "tx")), relleno: relleno(sp), linea: linea(sp),
          cats: cache(h1(s, "cat") || h1(s, "xVal")), vals: cache(h1(s, "val") || h1(s, "yVal")), suave: si(s, "smooth", false), rotulos: rotulos(s),
          explota: +val(s, "explosion", 0), invierte: si(s, "invertIfNegative", false), degr: degradado(sp), opac: alfa(sp && h1(sp, "solidFill")) };
        if (mk) se.marcador = { simbolo: val(mk, "symbol", "auto"), tam: +val(mk, "size", 5), relleno: relleno(h1(mk, "spPr")), linea: linea(h1(mk, "spPr")) };
        // líneas de tendencia (con su ecuación y su R²) y barras de error
        const tls = hs(s, "trendline").map((tl) => ({ tipo: val(tl, "trendlineType", "linear"), orden: +val(tl, "order", 2), periodo: +val(tl, "period", 2), ec: si(tl, "dispEq", false),
          r2: si(tl, "dispRSqr", false), linea: linea(h1(tl, "spPr")), letra: letra(h1(h1(tl, "trendlineLbl"), "txPr")), nombre: (h1(tl, "name") || {}).textContent || null,
          adelante: +val(tl, "forward", 0), atras: +val(tl, "backward", 0) }));
        if (tls.length) se.tendencias = tls;
        const eb = hs(s, "errBars").find((e) => val(e, "errDir", "y") !== "x");
        if (eb) se.error = { tipo: val(eb, "errBarType", "both"), valor: val(eb, "errValType", "fixedVal"), v: +val(eb, "val", 1), sinTapa: si(eb, "noEndCap", false), linea: linea(h1(eb, "spPr")) };
        const pts = hs(s, "dPt").map((p) => ({ idx: +val(p, "idx", 0), relleno: relleno(h1(p, "spPr")), linea: linea(h1(p, "spPr")), explota: +val(p, "explosion", 0) }));
        if (pts.length) se.puntos = pts;
        const tx = h1(s, "tx"), sr = tx && h1(tx, "strRef"); if (sr) se.refNombre = (todos(sr, "f")[0] || {}).textContent || null;
        gr.series.push(se);
      });
      gr.series.sort((a, b) => a.orden - b.orden);
      m.tipos.push(gr);
    });
    m.acentos = ctx.acentos || ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"];
    return m;
  }

  // ================================================================ medidas
  let lienzoMedir = null;
  let idsSvg = 0;   // ids únicos de los degradados (las gráficas van en la misma página)
  function medir(texto, px, familia, negrita) {
    try {
      lienzoMedir = lienzoMedir || document.createElement("canvas").getContext("2d");
      lienzoMedir.font = `${negrita ? "bold " : ""}${px}px ${familia}`;
      return lienzoMedir.measureText(String(texto)).width;
    } catch (e) { return String(texto).length * px * 0.55; }
  }
  /** Escala automática de un eje de valores, como Excel: el cero si el mínimo es menor que 5/6 del máximo, un
      5 % de aire y la unidad de 1, 2 o 5 × 10ⁿ que deja hasta 10 divisiones. */
  function escala(dmin, dmax, e, apilado100, maxDiv) {
    if (apilado100) return { a: e.min != null ? e.min : 0, b: e.max != null ? e.max : 1, u: e.unidad || 0.1 };
    if (!isFinite(dmin) || !isFinite(dmax)) { dmin = 0; dmax = 1; }
    if (dmin === dmax) { if (dmax === 0) dmax = 1; else if (dmax > 0) dmin = 0; else dmax = 0; }
    let a = e.min != null ? e.min : dmin >= 0 ? (dmin <= (5 / 6) * dmax ? 0 : dmin) : dmin;
    let b = e.max != null ? e.max : dmax <= 0 ? (dmax >= (5 / 6) * dmin ? 0 : dmax) : dmax;
    const aire = 0.05 * (b - a || Math.abs(b) || 1);
    const b2 = e.max != null ? b : b > 0 || dmax > 0 ? b + aire : b, a2 = e.min != null ? a : a < 0 ? a - aire : a;
    let u = e.unidad;
    if (!u) {
      const r = Math.max(1e-12, b2 - a2), p = Math.pow(10, Math.floor(Math.log10(r)));
      const tope = Math.max(2, Math.min(10, maxDiv || 10));
      u = [0.1, 0.2, 0.5, 1, 2, 5, 10, 20].map((k) => k * p).find((x) => r / x <= tope) || 20 * p;
    }
    if (e.max == null) b = Math.ceil(b2 / u - 1e-9) * u;
    if (e.min == null) a = a2 < 0 || (a2 > 0 && a !== 0) ? Math.floor(a2 / u + 1e-9) * u : a;
    return { a, b, u };
  }
  /** Una cifra con n dígitos significativos, sin ceros de más (el formato General de las etiquetas de tendencia). */
  function cifras(x, n) { if (!isFinite(x)) return ""; const r = Number((+x).toPrecision(n || 5)); return String(Math.abs(r) < 1e-12 ? 0 : r); }
  const NOMBRE_TENDENCIA = { linear: () => "Lineal", exp: () => "Exponencial", log: () => "Logarítmica", power: () => "Potencial", poly: () => "Polinómica", movingAvg: (tl) => `Media móvil de ${tl.periodo || 2} per.` };
  /** Línea de tendencia de Excel: lineal, exponencial, logarítmica, potencial, polinómica o media móvil, con su
      ecuación y su R² (en las exponencial y potencial, sobre los logaritmos, como Excel). */
  function ajustar(tl, pares) {
    const P = pares.filter((q) => q[1] != null && isFinite(q[1]) && isFinite(q[0]));
    if (P.length < 2) return null;
    if (tl.tipo === "movingAvg") {
      const per = Math.max(2, tl.periodo || 2), pts = [];
      for (let i = per - 1; i < P.length; i++) pts.push([P[i][0], P.slice(i - per + 1, i + 1).reduce((a, q) => a + q[1], 0) / per]);
      return pts.length ? { puntos: pts } : null;
    }
    const recta = (pp) => {
      const n = pp.length, sx = pp.reduce((a, q) => a + q[0], 0), sy = pp.reduce((a, q) => a + q[1], 0), sxx = pp.reduce((a, q) => a + q[0] * q[0], 0), sxy = pp.reduce((a, q) => a + q[0] * q[1], 0);
      const den = n * sxx - sx * sx; if (!den) return null;
      const b = (n * sxy - sx * sy) / den, a = (sy - b * sx) / n, my = sy / n;
      const tot = pp.reduce((s2, q) => s2 + (q[1] - my) * (q[1] - my), 0), res = pp.reduce((s2, q) => s2 + (q[1] - a - b * q[0]) * (q[1] - a - b * q[0]), 0);
      return { a, b, r2: tot ? 1 - res / tot : 1 };
    };
    const mas = (x) => (x < 0 ? ` - ${cifras(-x)}` : ` + ${cifras(x)}`);
    if (tl.tipo === "exp") { const q = P.filter((p) => p[1] > 0).map((p) => [p[0], Math.log(p[1])]), r = q.length > 1 && recta(q); if (!r) return null; const c = Math.exp(r.a); return { f: (x) => c * Math.exp(r.b * x), ec: `y = ${cifras(c)}e^${cifras(r.b)}x`, r2: r.r2 }; }
    if (tl.tipo === "log") { const q = P.filter((p) => p[0] > 0).map((p) => [Math.log(p[0]), p[1]]), r = q.length > 1 && recta(q); if (!r) return null; return { f: (x) => (x > 0 ? r.a + r.b * Math.log(x) : NaN), ec: `y = ${cifras(r.b)}ln(x)${mas(r.a)}`, r2: r.r2 }; }
    if (tl.tipo === "power") { const q = P.filter((p) => p[0] > 0 && p[1] > 0).map((p) => [Math.log(p[0]), Math.log(p[1])]), r = q.length > 1 && recta(q); if (!r) return null; const c = Math.exp(r.a); return { f: (x) => c * Math.pow(x, r.b), ec: `y = ${cifras(c)}x^${cifras(r.b)}`, r2: r.r2 }; }
    if (tl.tipo === "poly") {
      const k = Math.max(2, Math.min(6, tl.orden || 2)), n = k + 1;
      // ecuaciones normales con eliminación de Gauss
      const A = Array.from({ length: n }, (_, i) => Array.from({ length: n + 1 }, (_, j) => (j < n ? P.reduce((s2, q) => s2 + Math.pow(q[0], i + j), 0) : P.reduce((s2, q) => s2 + q[1] * Math.pow(q[0], i), 0))));
      for (let i = 0; i < n; i++) {
        let piv = i; for (let r2 = i + 1; r2 < n; r2++) if (Math.abs(A[r2][i]) > Math.abs(A[piv][i])) piv = r2;
        [A[i], A[piv]] = [A[piv], A[i]];
        if (!A[i][i]) return null;
        for (let r2 = 0; r2 < n; r2++) if (r2 !== i) { const f = A[r2][i] / A[i][i]; for (let j = i; j <= n; j++) A[r2][j] -= f * A[i][j]; }
      }
      const cf = A.map((row, i) => row[n] / row[i]), f = (x) => cf.reduce((s2, c, i) => s2 + c * Math.pow(x, i), 0);
      const my = P.reduce((a, q) => a + q[1], 0) / P.length, tot = P.reduce((s2, q) => s2 + (q[1] - my) ** 2, 0), res = P.reduce((s2, q) => s2 + (q[1] - f(q[0])) ** 2, 0);
      const terminos = cf.map((c, i) => [c, i]).reverse().filter((q) => Math.abs(q[0]) > 1e-12);
      const ec = "y = " + terminos.map(([c, i], j) => `${j ? (c < 0 ? " - " : " + ") : c < 0 ? "-" : ""}${cifras(Math.abs(c))}${i ? "x" + (i > 1 ? String(i).replace(/\d/g, (d) => "⁰¹²³⁴⁵⁶⁷⁸⁹"[d]) : "") : ""}`).join("");
      return { f, ec, r2: tot ? 1 - res / tot : 1 };
    }
    const r = recta(P);
    if (!r) return null;
    return { f: (x) => r.a + r.b * x, ec: `y = ${cifras(r.b)}x${mas(r.a)}`, r2: r.r2 };
  }
  function numero(v, fmt, opc) {
    if (v == null || v === "") return "";
    if (opc && opc.formato && fmt && fmt !== "General") { try { return opc.formato(+v, fmt); } catch (e) { /* tal cual */ } }
    const x = +v;
    if (!isFinite(x)) return String(v);
    const r = Math.round(x * 1e10) / 1e10;
    return Math.abs(r) >= 1e11 ? r.toExponential(2).replace("+", "") : String(r);
  }

  // ================================================================ dibujo
  function svg(m, w, h, opc) {
    if (m && m.ex) return svgEx(m, w, h, opc);
    opc = opc || {};
    const pt = opc.pt || 1, fam = (f) => (opc.familia ? opc.familia(f || m.letra.f || "Calibri") : `'${(f || m.letra.f || "Calibri").replace(/'/g, "")}', sans-serif`);
    const L = (l, def) => Object.assign({ tam: 10, c: "#595959", b: false, f: null }, m.letra || {}, def || {}, l || {});
    const partes = [];
    // Lo que se dibuja de una serie o de un punto lleva data-s (serie) y data-c (categoría): así las animaciones de
    // PowerPoint «por serie», «por categoría» o «por elemento» (<p:graphicEl><a:chart/>) encuentran sus partes.
    let marca = "";
    const meter = partes.push.bind(partes);
    partes.push = (...xs) => meter(...xs.map((x) => (marca && typeof x === "string" ? x.replace(/^<([a-zA-Z]+)/, "<$1" + marca) : x)));
    const marcar = (si, ci) => { marca = si == null ? "" : ` data-s="${si}"` + (ci == null ? "" : ` data-c="${ci}"`); };
    // (peso y estilo siempre escritos: la gráfica no hereda la negrita del lugar donde se pinta)
    const t = (x, y, s, l, anc, extra) => partes.push(`<text x="${n2(x)}" y="${n2(y)}" font-family="${esc(fam(l.f))}" font-size="${n2(l.tam * pt)}" fill="${esc(l.c || "#595959")}" font-weight="${l.b ? 700 : 400}" font-style="${l.i ? "italic" : "normal"}" text-anchor="${anc || "start"}"${extra || ""}>${esc(l.cap === "all" ? String(s).toUpperCase() : s)}</text>`);
    // un degradado de Office como degradado de SVG (en la caja de cada figura, como «rotWithShape»)
    const degSvg = (dg) => {
      const id = "nomdg" + ++idsSvg;
      const st = dg.p.map((q) => `<stop offset="${n2(q[0])}" stop-color="${esc(q[1])}"${q[2] != null ? ` stop-opacity="${n2(q[2])}"` : ""}/>`).join("");
      if (dg.radial) {
        const [l, tt, r, b] = dg.foco, cx = (l + (1 - r)) / 2, cy = (tt + (1 - b)) / 2;
        const rad = Math.max(...[[0, 0], [1, 0], [0, 1], [1, 1]].map(([x, y]) => Math.hypot(x - cx, y - cy)));
        partes.push(`<defs><radialGradient id="${id}" cx="${n2(cx)}" cy="${n2(cy)}" r="${n2(rad)}">${st}</radialGradient></defs>`);
      } else {
        const a = ((dg.ang || 0) * Math.PI) / 180, dx = Math.cos(a) / 2, dy = Math.sin(a) / 2;
        partes.push(`<defs><linearGradient id="${id}" x1="${n2(0.5 - dx)}" y1="${n2(0.5 - dy)}" x2="${n2(0.5 + dx)}" y2="${n2(0.5 + dy)}">${st}</linearGradient></defs>`);
      }
      return `url(#${id})`;
    };
    const trazo = (ln, def) => { if (ln && ln.no) return null; const o = Object.assign({ w: 0.75, c: "#D9D9D9" }, def || {}, ln || {}); return o; };
    const GUIONES = { sysDot: [1, 1], sysDash: [3, 1], dash: [4, 3], dashDot: [4, 3, 1, 3], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1] };
    const guion = (o) => {
      if (!o || !o.guion) return "";
      const w = Math.max(0.25, (o.w || 0.75) * pt), pat = GUIONES[o.guion] || [3, 2];
      // con punta redonda la raya crece medio ancho de cada lado: se descuenta de la raya y se suma al hueco
      return ` stroke-dasharray="${pat.map((k, i) => n2(Math.max(i % 2 ? 0 : 0.001, o.redondo ? (i % 2 ? k + 1 : k - 1) * w : k * w))).join(" ")}"${o.redondo ? ' stroke-linecap="round"' : ""}`;
    };
    const attrsTrazo = (o) => (o ? ` stroke="${esc(o.c || "#000")}" stroke-width="${n2(Math.max(0.25, o.w * pt))}"${guion(o)}` : ' stroke="none"');
    // fondo de la gráfica
    const fa = m.area.degr ? degSvg(m.area.degr) : m.area.relleno === "none" ? "none" : m.area.relleno || "#FFFFFF", la = trazo(m.area.linea, opc.sinBorde ? { no: 1 } : { c: "#D9D9D9", w: 0.75 });
    partes.push(`<rect x="0.5" y="0.5" width="${n2(w - 1)}" height="${n2(h - 1)}" fill="${esc(fa)}"${la && !la.no ? attrsTrazo(la) : ""}${m.area.redondo ? ' rx="6"' : ""}/>`);
    const series = [];
    m.tipos.forEach((g, gi) => g.series.forEach((s, si) => series.push({ g, s, i: series.length })));
    const colorSerie = (k) => {
      const acc = m.acentos, base = acc[k % 6], vuelta = Math.floor(k / 6);
      if (!vuelta) return base;
      const f = vuelta === 1 ? 0.6 : vuelta === 2 ? 1.4 : 0.8, n = parseInt(base.slice(1), 16), c = [n >> 16, (n >> 8) & 255, n & 255];
      return "#" + c.map((v) => Math.round(f < 1 ? v * f : v + (255 - v) * (f - 1)).toString(16).padStart(2, "0")).join("");
    };
    // valores: los de opc.datos (lo que alguien editó en la app), los vivos de la hoja, o los guardados
    const ordenDe = (se) => series.findIndex((x) => x.s === se);
    const datos = (se) => {
      const k = ordenDe(se), od = opc.datos && opc.datos.series && opc.datos.series[k];
      const vivo = od ? od.vals : opc.valores && se.vals.f ? opc.valores(se.vals.f) : null;
      return (vivo || se.vals.v).map((v) => (v == null || v === "" || !isFinite(+v) ? null : +v));
    };
    const cats = (() => { const s0 = series[0] && series[0].s; if (!s0) return []; const vivo = opc.datos && opc.datos.cats ? opc.datos.cats : opc.valores && s0.cats.f ? opc.valores(s0.cats.f) : null; const c = vivo || s0.cats.v; const n = Math.max(c.length, ...series.map((x) => datos(x.s).length)); return Array.from({ length: n }, (_, i) => (c[i] == null ? String(i + 1) : String(c[i]))); })();
    const nombreSerie = (se, k) => { const od = opc.datos && opc.datos.series && opc.datos.series[ordenDe(se)]; if (od && od.nombre != null) return String(od.nombre); const v = opc.valores && se.refNombre ? opc.valores(se.refNombre) : null; return (v && v[0] != null ? String(v[0]) : se.nombre) || `Serie${k + 1}`; };
    const unaRebanada = m.tipos.length === 1 && (m.tipos[0].tipo === "pastel" || m.tipos[0].tipo === "anillo");
    // ---- título
    let arriba = 7 * pt, abajo = h - 7 * pt, izq = 7 * pt, der = w - 7 * pt;
    const tituloTexto = m.titulo ? (m.titulo.texto != null ? m.titulo.texto : series.length === 1 ? nombreSerie(series[0].s, 0) : /^es/i.test(m.lengua) ? "Título del gráfico" : "Chart Title") : !m.sinTitulo && series.length === 1 && unaRebanada ? nombreSerie(series[0].s, 0) : null;
    if (tituloTexto) {
      const lt = L(m.titulo && m.titulo.letra, { tam: 14 }), lineas = String(tituloTexto).split("\n");
      let y = arriba + lt.tam * pt * 0.95;
      if (m.titulo && m.titulo.disp && m.titulo.disp.y != null) y = m.titulo.disp.y * h + lt.tam * pt * 0.95;
      lineas.forEach((ln, k) => t(w / 2 + (m.titulo && m.titulo.disp && m.titulo.disp.x != null ? 0 : 0), y + k * lt.tam * pt * 1.2, ln, lt, "middle"));
      if (!(m.titulo && m.titulo.sobre)) arriba = y + (lineas.length - 1) * lt.tam * pt * 1.2 + lt.tam * pt * 0.45 + 4 * pt;
    }
    // ---- leyenda
    let entradas = [];
    if (m.leyenda) {
      const ll = L(m.leyenda.letra, { tam: 9 });
      if (unaRebanada) entradas = cats.map((c, i) => ({ texto: c, color: colorPunto(series[0].s, i), tipo: "caja" }));
      else entradas = series.map(({ g, s }, k) => ({ texto: nombreSerie(s, k), color: g.tipo === "linea" || g.tipo === "dispersion" || g.tipo === "radar" ? (s.linea && s.linea.c) || colorSerie(k) : s.relleno && s.relleno !== "none" ? s.relleno : (s.linea && s.linea.c) || colorSerie(k),
        tipo: g.tipo === "linea" || g.tipo === "dispersion" || g.tipo === "radar" ? "linea" : "caja", k,
        marca: (g.tipo === "linea" || g.tipo === "dispersion") && (s.marcador ? s.marcador.simbolo !== "none" : g.marcadores && g.tipo === "linea") ? (s.marcador && s.marcador.simbolo) || "circle" : null }));
      // barras horizontales: Excel lista las series al revés (como se ven de arriba abajo)
      if (!unaRebanada && m.tipos[0] && m.tipos[0].tipo === "barra" && m.tipos[0].dir === "bar" && (m.leyenda.pos === "b" || m.leyenda.pos === "t")) entradas.reverse();
      if (!unaRebanada) series.forEach(({ s }, k) => (s.tendencias || []).forEach((tl) => entradas.push({ texto: tl.nombre || `${NOMBRE_TENDENCIA[tl.tipo] ? NOMBRE_TENDENCIA[tl.tipo](tl) : "Tendencia"} (${nombreSerie(s, k)})`, color: (tl.linea && tl.linea.c) || colorSerie(k), tipo: "linea", guion: tl.linea })));
      entradas = entradas.filter((e, i) => !(m.leyenda.borrar || []).includes(i));
      const cajaW = 7 * pt, sep = 4 * pt, gapE = 10 * pt, lineaW = 18 * pt;
      const anchoE = (e) => (e.tipo === "linea" ? lineaW : cajaW) + sep + medir(e.texto, ll.tam * pt, fam(ll.f));
      const pos = m.leyenda.pos;
      if (pos === "b" || pos === "t") {
        // en renglones centrados
        const filas = [[]]; let ancho = 0;
        entradas.forEach((e) => { const a = anchoE(e); if (ancho + a > der - izq && filas[filas.length - 1].length) { filas.push([]); ancho = 0; } filas[filas.length - 1].push(e); ancho += a + gapE; });
        const altoF = ll.tam * pt * 1.5, total = filas.length * altoF;
        let y0 = pos === "b" ? abajo - total : arriba;
        if (m.leyenda.disp && m.leyenda.disp.y != null) y0 = m.leyenda.disp.y * h;
        filas.forEach((fila, k) => {
          const anchoF = fila.reduce((s, e) => s + anchoE(e), 0) + gapE * (fila.length - 1);
          let x = w / 2 - anchoF / 2; const yc = y0 + k * altoF + altoF / 2;
          fila.forEach((e) => { dibujarClave(e, x, yc); t(x + (e.tipo === "linea" ? lineaW : cajaW) + sep, yc + ll.tam * pt * 0.35, e.texto, ll); x += anchoE(e) + gapE; });
        });
        if (!m.leyenda.sobre) { if (pos === "b") abajo = y0 - 4 * pt; else arriba = y0 + total + 4 * pt; }
      } else {
        const anchoMax = Math.max(0, ...entradas.map(anchoE)), altoF = ll.tam * pt * 1.6, total = entradas.length * altoF;
        const x0 = pos === "l" ? izq : der - anchoMax;
        let y0 = pos === "tr" ? arriba : (arriba + abajo) / 2 - total / 2;
        if (m.leyenda.disp && m.leyenda.disp.y != null) y0 = m.leyenda.disp.y * h;
        entradas.forEach((e, k) => { const yc = y0 + k * altoF + altoF / 2; dibujarClave(e, x0, yc); t(x0 + (e.tipo === "linea" ? lineaW : cajaW) + sep, yc + ll.tam * pt * 0.35, e.texto, ll); });
        if (!m.leyenda.sobre) { if (pos === "l") izq = x0 + anchoMax + 8 * pt; else der = x0 - 8 * pt; }
      }
      function dibujarClave(e, x, yc) {
        if (e.tipo === "linea") { partes.push(`<line x1="${n2(x)}" y1="${n2(yc)}" x2="${n2(x + lineaW)}" y2="${n2(yc)}" stroke="${esc(e.color)}" stroke-width="${n2((e.guion && e.guion.w ? e.guion.w : 2.25) * pt)}" stroke-linecap="round"${e.guion ? guion(e.guion) : ""}/>`); if (e.marca) partes.push(`<circle cx="${n2(x + lineaW / 2)}" cy="${n2(yc)}" r="${n2(2.5 * pt)}" fill="${esc(e.color)}"/>`); return; }
        partes.push(`<rect x="${n2(x)}" y="${n2(yc - cajaW / 2)}" width="${n2(cajaW)}" height="${n2(cajaW)}" fill="${esc(e.color)}"/>`);
      }
    }
    function colorPunto(se, i) {
      const p = (se.puntos || []).find((q) => q.idx === i);
      if (p && p.relleno && p.relleno !== "none") return p.relleno;
      return colorSerie(i);
    }
    // ---- área de trazado
    let pl = izq, pr = der, ptop = arriba, pb = abajo;
    const disp = m.trazado.disp;
    if (disp && disp.x != null && disp.w != null) { pl = disp.x * w; pr = (disp.x + disp.w) * w; ptop = disp.y * h; pb = (disp.y + disp.h) * h; }
    if (unaRebanada) {
      dibujarPastel(m.tipos[0], series[0].s, pl, ptop, pr, pb);
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${n2(w)}" height="${n2(h)}" viewBox="0 0 ${n2(w)} ${n2(h)}">${partes.join("")}</svg>`;
    }
    // ejes: el de categorías (o de X en dispersión) y el de valores
    const ejes = Object.values(m.ejes);
    const g0 = m.tipos[0] || { ejes: [] };
    const ejeC = ejes.find((e) => g0.ejes.includes(e.id) && (e.tipo === "catAx" || e.tipo === "dateAx")) || ejes.find((e) => g0.ejes.includes(e.id) && g0.tipo === "dispersion" && (e.pos === "b" || e.pos === "t")) || { tipo: "catAx", pos: "b", letra: {}, rotulos: "nextTo" };
    const ejeV = ejes.find((e) => g0.ejes.includes(e.id) && e.tipo === "valAx" && e !== ejeC) || { tipo: "valAx", pos: "l", letra: {}, rotulos: "nextTo", rejilla: {} };
    const horizontal = g0.tipo === "barra" && g0.dir === "bar";
    if (g0.tresD && g0.tipo === "barra" && !horizontal && m.tipos.length === 1 && !(disp && disp.dentro)) {
      dibujarColumnas3D(g0, pl, ptop, pr, pb);
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${n2(w)}" height="${n2(h)}" viewBox="0 0 ${n2(w)} ${n2(h)}">${partes.join("")}</svg>`;
    }
    const apilado = (g) => g.grupo === "stacked" || g.grupo === "percentStacked";
    // eje de valores secundario (combinada): el de los grupos que no usan el eje principal; va a la derecha
    const g2 = m.tipos.find((g) => g !== g0 && !g.ejes.includes(ejeV.id) && g.tipo !== "dispersion");
    const ejeV2 = g2 ? ejes.find((e) => g2.ejes.includes(e.id) && e.tipo === "valAx") || null : null;
    const enSec = (g) => !!ejeV2 && !g.ejes.includes(ejeV.id) && g.ejes.includes(ejeV2.id);
    // extremos de los datos (con las sumas si van apiladas), de cada eje
    let dmin = Infinity, dmax = -Infinity, dmin2 = Infinity, dmax2 = -Infinity;
    const pct = m.tipos.some((g) => g.grupo === "percentStacked" && !enSec(g)), pct2 = m.tipos.some((g) => g.grupo === "percentStacked" && enSec(g));
    m.tipos.forEach((g) => {
      const ss = g.series.map(datos), sec = enSec(g);
      const toma = (lo, hi) => { if (sec) { dmax2 = Math.max(dmax2, hi); dmin2 = Math.min(dmin2, lo); } else { dmax = Math.max(dmax, hi); dmin = Math.min(dmin, lo); } };
      if (apilado(g)) {
        for (let i = 0; i < cats.length; i++) {
          let pos = 0, neg = 0;
          ss.forEach((v) => { const x = v[i] || 0; if (x >= 0) pos += x; else neg += x; });
          if (g.grupo === "percentStacked") { const tot = pos - neg || 1; pos /= tot; neg /= tot; }
          toma(neg, pos);
        }
      } else ss.forEach((v) => v.forEach((x) => { if (x != null) toma(x, x); }));
    });
    // cuántas divisiones caben: Office deja al menos unas 2.2 alturas de letra entre rótulos (en vertical)
    const lv0 = L(ejeV.letra, { tam: 9 }), lc0 = L(ejeC.letra, { tam: 9 });
    const largoEje = horizontal ? Math.max(40, pr - pl - (Math.max(0, ...cats.map((c) => medir(c, lc0.tam * pt, fam(lc0.f)))) + 6 * pt)) : Math.max(40, pb - ptop - lc0.tam * pt * 1.6);
    const maxDiv = Math.floor(largoEje / (lv0.tam * pt * (horizontal ? 3.6 : 2.2)));
    const esc_ = escala(dmin, dmax, ejeV, pct, maxDiv);
    // (Excel: el secundario lleva tantas divisiones como el principal, para que las rayas coincidan)
    const esc2 = ejeV2 ? escala(dmin2, dmax2, ejeV2, pct2, maxDiv) : null;
    // dispersión: escala también en X
    let escX = null;
    if (g0.tipo === "dispersion") {
      let xmin = Infinity, xmax = -Infinity;
      series.forEach(({ s }) => (opc.valores && s.cats.f ? opc.valores(s.cats.f) || s.cats.v : s.cats.v).forEach((x) => { if (x != null && isFinite(+x)) { xmin = Math.min(xmin, +x); xmax = Math.max(xmax, +x); } }));
      escX = escala(xmin, xmax, ejeC, false);
    }
    // rótulos de los ejes: espacio que ocupan
    const lc = L(ejeC.letra, { tam: 9 }), lv = L(ejeV.letra, { tam: 9 });
    const marcasV = [];
    for (let v = esc_.a, k = 0; v <= esc_.b + esc_.u * 1e-6 && k < 200; v = esc_.a + ++k * esc_.u) marcasV.push(Math.round(v * 1e10) / 1e10);
    // (100 % apilado: Excel pone el eje en porcentaje, «0%», aunque los datos digan otra cosa)
    const marcasDe = (e) => { const o = []; for (let v = e.a, k = 0; v <= e.b + e.u * 1e-6 && k < 200; v = e.a + ++k * e.u) o.push(Math.round(v * 1e10) / 1e10); return o; };
    const fmtV2 = (v) => (pct2 && !(ejeV2 && ejeV2.fmt && /%/.test(ejeV2.fmt)) ? Math.round(v * 100) + "%" : numero(v, ejeV2 && ejeV2.fmt, opc));
    const fmtV = (v) => (pct && !(ejeV.fmt && /%/.test(ejeV.fmt) && !ejeV.ligado) ? Math.round(v * 100) + "%" : numero(v, ejeV.ligado ? (series[0] && series[0].s.vals.fmt) || ejeV.fmt : ejeV.fmt, opc));
    const verV = !ejeV.borrado && ejeV.rotulos !== "none", verC = !ejeC.borrado && ejeC.rotulos !== "none";
    if (!disp || !disp.dentro) {
      if (horizontal) {
        if (verC) pl += Math.max(...cats.map((c) => medir(c, lc.tam * pt, fam(lc.f)))) + 6 * pt;
        if (verV) pb -= lv.tam * pt * 1.6;
      } else {
        if (verV) pl += Math.max(...marcasV.map((v) => medir(fmtV(v), lv.tam * pt, fam(lv.f)))) + 6 * pt;
        if (verC && !m.tablaDatos) pb -= lc.tam * pt * 1.6;
        if (m.tablaDatos) {
          const ld = L(m.tablaDatos.letra, { tam: 9 }), filaD = ld.tam * pt * 1.5;
          pb -= filaD * (1 + series.length);
          const colN = Math.max(0, ...series.map(({ s }, k) => medir(nombreSerie(s, k), ld.tam * pt, fam(ld.f)))) + (m.tablaDatos.claves ? 9 * pt : 3 * pt) + 3 * pt;
          pl = Math.max(pl, izq + colN);
        }
      }
      // Office deja 11 pt a la derecha del área de trazado (no 7) y, en barras horizontales, medio rótulo más
      if (!(m.leyenda && (m.leyenda.pos === "r" || m.leyenda.pos === "tr") && !m.leyenda.sobre)) pr -= 4 * pt;
      if (horizontal && verV) pr -= Math.max(0, ...marcasV.map((v) => medir(fmtV(v), lv.tam * pt, fam(lv.f)))) / 2;
      if (esc2 && !ejeV2.borrado && ejeV2.rotulos !== "none") pr -= Math.max(0, ...marcasDe(esc2).map((v) => medir(fmtV2(v), lv.tam * pt, fam(lv.f)))) + 6 * pt;
      if (ejeV.titulo) { const lt = L(ejeV.titulo.letra, { tam: 10 }); if (horizontal) pb -= lt.tam * pt * 1.6; else pl += lt.tam * pt * 1.6; }
      if (ejeC.titulo) { const lt = L(ejeC.titulo.letra, { tam: 10 }); if (horizontal) pl += lt.tam * pt * 1.6; else pb -= lt.tam * pt * 1.6; }
    }
    if (m.trazado.degr || (m.trazado.relleno && m.trazado.relleno !== "none")) partes.push(`<rect x="${n2(pl)}" y="${n2(ptop)}" width="${n2(Math.max(0, pr - pl))}" height="${n2(Math.max(0, pb - ptop))}" fill="${m.trazado.degr ? degSvg(m.trazado.degr) : esc(m.trazado.relleno)}"/>`);
    // posición de un valor y de una categoría
    const largoV = horizontal ? pr - pl : pb - ptop;
    let escAct = esc_;
    const aV = (v) => { const k = (v - escAct.a) / ((escAct.b - escAct.a) || 1); return horizontal ? (ejeV.inv ? pr - k * largoV : pl + k * largoV) : ejeV.inv ? ptop + k * largoV : pb - k * largoV; };
    const nC = Math.max(1, cats.length), entreCat = !(ejeV.entre === "midCat") || m.tipos.some((g) => g.tipo === "barra");
    const anchoCat = (horizontal ? pb - ptop : pr - pl) / (entreCat ? nC : Math.max(1, nC - 1));
    const aC = (i) => { const k = entreCat ? (i + 0.5) * anchoCat : i * anchoCat; return horizontal ? (ejeC.inv ? ptop + k : pb - k) : ejeC.inv ? pr - k : pl + k; };
    const aX = (x) => pl + ((x - escX.a) / ((escX.b - escX.a) || 1)) * (pr - pl);
    // cuadrícula y rótulos del eje de valores
    const rej = ejeV.rejilla ? trazo(ejeV.rejilla, { c: "#D9D9D9", w: 0.75 }) : null;
    marcasV.forEach((v) => {
      const p = aV(v);
      if (rej) partes.push(horizontal ? `<line x1="${n2(p)}" y1="${n2(ptop)}" x2="${n2(p)}" y2="${n2(pb)}"${attrsTrazo(rej)}/>` : `<line x1="${n2(pl)}" y1="${n2(p)}" x2="${n2(pr)}" y2="${n2(p)}"${attrsTrazo(rej)}/>`);
      if (verV) { if (horizontal) t(p, pb + lv.tam * pt * 1.25, fmtV(v), lv, "middle"); else t(pl - 4 * pt, p + lv.tam * pt * 0.35, fmtV(v), lv, "end"); }
    });
    if (escX) {
      const rx = ejeC.rejilla ? trazo(ejeC.rejilla, { c: "#D9D9D9", w: 0.75 }) : null;
      for (let x = escX.a, k = 0; x <= escX.b + escX.u * 1e-6 && k < 200; x = escX.a + ++k * escX.u) {
        const p = aX(x);
        if (rx) partes.push(`<line x1="${n2(p)}" y1="${n2(ptop)}" x2="${n2(p)}" y2="${n2(pb)}"${attrsTrazo(rx)}/>`);
        if (verC) t(p, pb + lc.tam * pt * 1.25, numero(Math.round(x * 1e10) / 1e10, ejeC.fmt, opc), lc, "middle");
      }
    }
    // rótulos del eje secundario, a la derecha
    if (esc2 && !ejeV2.borrado && ejeV2.rotulos !== "none" && !horizontal) {
      const lv2 = L(ejeV2.letra, { tam: 9 });
      escAct = esc2; marcasDe(esc2).forEach((v) => t(pr + 4 * pt, aV(v) + lv2.tam * pt * 0.35, fmtV2(v), lv2, "start")); escAct = esc_;
    }
    // ---- series
    const centros = new Map();   // dónde quedó cada punto de cada serie (para las barras de error)
    m.tipos.forEach((g) => {
      escAct = enSec(g) ? esc2 : esc_;
      const ss = g.series.map((s) => ({ s, v: datos(s) })), k0 = series.findIndex((x) => x.g === g);
      if (g.tipo === "barra") {
        const N = apilado(g) ? 1 : ss.length, sol = apilado(g) ? 100 : g.solape, hueco = g.hueco;
        const bw = anchoCat / (N - ((N - 1) * sol) / 100 + hueco / 100);
        const base = aV(Math.max(escAct.a, Math.min(escAct.b, 0)));
        const pos = new Array(nC).fill(0), neg = new Array(nC).fill(0), tot = cats.map((_, i) => ss.reduce((a, x) => a + Math.abs(x.v[i] || 0), 0) || 1);
        ss.forEach(({ s, v }, j) => {
          const col = s.relleno && s.relleno !== "none" ? s.relleno : colorSerie(k0 + j), ln = trazo(s.linea, { no: 1 });
          const pintaS = s.degr ? degSvg(s.degr) : null, opacS = s.opac != null && s.opac < 1 ? ` fill-opacity="${n2(s.opac)}"` : "";
          v.forEach((x, i) => {
            if (x == null) return;
            let a, b;
            if (apilado(g)) { const y = g.grupo === "percentStacked" ? x / tot[i] : x; if (y >= 0) { a = pos[i]; b = pos[i] += y; } else { a = neg[i]; b = neg[i] += y; } }
            else { a = 0; b = x; }
            const desde = apilado(g) ? aV(a) : base, hasta = aV(b);
            const off = (apilado(g) ? 0 : j * bw * (1 - sol / 100)) + (bw * hueco) / 200 - anchoCat / 2;
            const c = aC(i) + (horizontal ? -off - bw : off);
            const fill = (s.puntos || []).find((q) => q.idx === i && q.relleno) ? s.puntos.find((q) => q.idx === i).relleno : g.varia && ss.length === 1 ? colorSerie(i) : col;
            const r = horizontal ? { x: Math.min(desde, hasta), y: c, w: Math.abs(hasta - desde), h: bw } : { x: c, y: Math.min(desde, hasta), w: bw, h: Math.abs(hasta - desde) };
            const propio = (s.puntos || []).some((q) => q.idx === i && q.relleno);
            marcar(k0 + j, i);
            partes.push(`<rect x="${n2(r.x)}" y="${n2(r.y)}" width="${n2(r.w)}" height="${n2(r.h)}" fill="${pintaS && !propio ? pintaS : esc(fill)}"${opacS}${ln ? attrsTrazo(ln) : ""}/>`);
            if (!centros.has(s)) centros.set(s, []);
            centros.get(s)[i] = horizontal ? null : r.x + r.w / 2;
            const ro = s.rotulos || g.rotulos;
            if (ro && ro.val) {
              // dentro arriba (inEnd), dentro abajo (inBase), al centro (ctr) o afuera (outEnd, lo de siempre)
              const lr = L(ro.letra, { tam: 9 }), txt = numero(x, ro.fmt || s.vals.fmt, opc), alto = lr.tam * pt;
              if (horizontal) {
                const pos = ro.pos === "inEnd" ? [r.x + r.w - 3 * pt, "end"] : ro.pos === "inBase" ? [r.x + 3 * pt, "start"] : ro.pos === "ctr" ? [r.x + r.w / 2, "middle"] : [r.x + r.w + 3 * pt, "start"];
                t(pos[0], r.y + r.h / 2 + alto * 0.35, txt, lr, pos[1]);
              } else {
                const y = ro.pos === "inEnd" ? r.y + alto * 0.95 + 2 * pt : ro.pos === "inBase" ? r.y + r.h - 3 * pt : ro.pos === "ctr" ? r.y + r.h / 2 + alto * 0.35 : r.y - 3 * pt;
                t(r.x + r.w / 2, y, txt, lr, "middle");
              }
            }
          });
        });
        marcar(null);
      } else if (g.tipo === "linea" || g.tipo === "area" || g.tipo === "radar") {
        const acum = new Array(nC).fill(0), tot = cats.map((_, i) => ss.reduce((a, x) => a + Math.abs(x.v[i] || 0), 0) || 1);
        const capas = [];
        ss.forEach(({ s, v }, j) => {
          const col = (s.linea && s.linea.c) || (s.relleno && s.relleno !== "none" ? s.relleno : colorSerie(k0 + j));
          const ys = v.map((x, i) => { if (x == null) return null; const y = g.grupo === "percentStacked" ? x / tot[i] : x; if (apilado(g)) { acum[i] += y; return acum[i]; } return y; });
          const pts = ys.map((y, i) => (y == null ? null : [aC(i), aV(y)]));
          centros.set(s, ys.map((y, i) => aC(i)));
          capas.push({ s, col, pts, j });
        });
        if (g.tipo === "area") {
          capas.slice().reverse().forEach(({ s, col, pts, j }, jj) => {
            const ok = pts.filter(Boolean); if (!ok.length) return;
            marcar(k0 + j);
            const base = aV(Math.max(escAct.a, 0));
            const d = `M${n2(ok[0][0])},${n2(base)} ` + ok.map((p) => `L${n2(p[0])},${n2(p[1])}`).join(" ") + ` L${n2(ok[ok.length - 1][0])},${n2(base)} Z`;
            const fill = s.relleno && s.relleno !== "none" ? s.relleno : col;
            partes.push(`<path d="${d}" fill="${esc(fill)}"${trazo(s.linea, { no: 1 }) ? attrsTrazo(trazo(s.linea, { no: 1 })) : ""}/>`);
          });
        } else {
          capas.forEach(({ s, col, pts, j }) => {
            marcar(k0 + j);
            const ln = trazo(s.linea, { c: col, w: 2.25 });
            let d = "", abierto = false;
            pts.forEach((p, i) => { if (!p) { abierto = false; return; } d += `${abierto ? "L" : "M"}${n2(p[0])},${n2(p[1])} `; abierto = true; });
            if (ln && d) partes.push(`<path d="${d.trim()}" fill="none"${attrsTrazo(ln)} stroke-linejoin="round" stroke-linecap="round"/>`);
            const mk = s.marcador, sim = mk ? mk.simbolo : g.marcadores ? "auto" : "none";
            if (sim && sim !== "none") pts.forEach((p, i) => { if (!p) return; marcar(k0 + j, i); const r = ((mk && mk.tam) || 5) * pt * 0.5; const f = (mk && mk.relleno && mk.relleno !== "none") ? mk.relleno : col; partes.push(sim === "square" ? `<rect x="${n2(p[0] - r)}" y="${n2(p[1] - r)}" width="${n2(2 * r)}" height="${n2(2 * r)}" fill="${esc(f)}"/>` : sim === "diamond" ? `<path d="M${n2(p[0])},${n2(p[1] - r * 1.3)} L${n2(p[0] + r * 1.3)},${n2(p[1])} L${n2(p[0])},${n2(p[1] + r * 1.3)} L${n2(p[0] - r * 1.3)},${n2(p[1])} Z" fill="${esc(f)}"/>` : sim === "triangle" ? `<path d="M${n2(p[0])},${n2(p[1] - r * 1.2)} L${n2(p[0] + r * 1.2)},${n2(p[1] + r)} L${n2(p[0] - r * 1.2)},${n2(p[1] + r)} Z" fill="${esc(f)}"/>` : `<circle cx="${n2(p[0])}" cy="${n2(p[1])}" r="${n2(r)}" fill="${esc(f)}"/>`); });
            const ro = s.rotulos || g.rotulos;
            if (ro && ro.val) { const lr = L(ro.letra, { tam: 9 }); pts.forEach((p, i) => { marcar(k0 + j, i); if (p) t(p[0], p[1] - 6 * pt, numero(s.vals.v[i], ro.fmt || s.vals.fmt, opc), lr, "middle"); }); }
          });
        }
      } else if (g.tipo === "dispersion") {
        ss.forEach(({ s, v }, j) => {
          marcar(k0 + j);
          const xs = (opc.valores && s.cats.f ? opc.valores(s.cats.f) || s.cats.v : s.cats.v).map((x) => (x == null || !isFinite(+x) ? null : +x));
          const col = (s.marcador && s.marcador.relleno && s.marcador.relleno !== "none" ? s.marcador.relleno : null) || (s.linea && s.linea.c) || colorSerie(k0 + j);
          const pts = v.map((y, i) => (y == null || xs[i] == null ? null : [aX(xs[i]), aV(y)]));
          const conLinea = /line|smooth/i.test(g.estiloDisp) && !(s.linea && s.linea.no);
          if (conLinea) { const ok = pts.filter(Boolean); if (ok.length > 1) partes.push(`<path d="${ok.map((p, k) => `${k ? "L" : "M"}${n2(p[0])},${n2(p[1])}`).join(" ")}" fill="none"${attrsTrazo(trazo(s.linea, { c: col, w: 2.25 }))}/>`); }
          const sim = s.marcador ? s.marcador.simbolo : "auto";
          if (sim !== "none") pts.forEach((p, i) => { marcar(k0 + j, i); if (p) partes.push(`<circle cx="${n2(p[0])}" cy="${n2(p[1])}" r="${n2(((s.marcador && s.marcador.tam) || 5) * pt * 0.5)}" fill="${esc(col)}"/>`); });
        });
      }
    });
    marcar(null);
    // ---- barras de error y líneas de tendencia, encima de las series
    m.tipos.forEach((g) => {
      if (horizontal || !(g.tipo === "barra" || g.tipo === "linea" || g.tipo === "area")) return;
      escAct = enSec(g) ? esc2 : esc_;
      g.series.forEach((s) => {
        const v = datos(s), c = centros.get(s) || [], k = series.findIndex((x) => x.s === s);
        marcar(k);
        const eb = s.error;
        if (eb) {
          // valor fijo, porcentaje, desviación estándar (Excel la centra en la media) o error estándar
          const nums = v.filter((x) => x != null), n = nums.length, media = n ? nums.reduce((a, b) => a + b, 0) / n : 0;
          const sd = n > 1 ? Math.sqrt(nums.reduce((a, x) => a + (x - media) * (x - media), 0) / (n - 1)) : 0;
          const lnE = trazo(eb.linea, { c: "#595959", w: 0.75 }), tapa = 3 * pt;
          v.forEach((x, i) => {
            if (x == null || c[i] == null || !lnE) return;
            const e = eb.valor === "percentage" ? (Math.abs(x) * eb.v) / 100 : eb.valor === "stdDev" ? eb.v * sd : eb.valor === "stdErr" ? sd / Math.sqrt(Math.max(1, n)) : eb.valor === "cust" ? 0 : eb.v;
            const centro = eb.valor === "stdDev" ? media : x, ya = aV(eb.tipo === "minus" ? centro : centro + e), yb = aV(eb.tipo === "plus" ? centro : centro - e), cx = c[i];
            partes.push(`<line x1="${n2(cx)}" y1="${n2(ya)}" x2="${n2(cx)}" y2="${n2(yb)}"${attrsTrazo(lnE)}/>`);
            if (!eb.sinTapa) {
              if (eb.tipo !== "minus") partes.push(`<line x1="${n2(cx - tapa)}" y1="${n2(ya)}" x2="${n2(cx + tapa)}" y2="${n2(ya)}"${attrsTrazo(lnE)}/>`);
              if (eb.tipo !== "plus") partes.push(`<line x1="${n2(cx - tapa)}" y1="${n2(yb)}" x2="${n2(cx + tapa)}" y2="${n2(yb)}"${attrsTrazo(lnE)}/>`);
            }
          });
        }
        (s.tendencias || []).forEach((tl) => {
          const aj = ajustar(tl, v.map((y, i) => [i + 1, y]));
          if (!aj) return;
          const lnT = trazo(tl.linea, { c: colorSerie(k), w: 2 }), x0 = 1 - (tl.atras || 0), x1 = v.length + (tl.adelante || 0);
          const xPant = (x) => aC(0) + (x - 1) * anchoCat;
          let d = "";
          if (aj.puntos) d = aj.puntos.map((q, j) => `${j ? "L" : "M"}${n2(xPant(q[0]))},${n2(aV(q[1]))}`).join("");
          else { const pasos = tl.tipo === "linear" ? 1 : 40; for (let j = 0; j <= pasos; j++) { const x = x0 + ((x1 - x0) * j) / pasos, y = aj.f(x); if (isFinite(y)) d += `${d ? "L" : "M"}${n2(xPant(x))},${n2(aV(y))}`; } }
          if (d && lnT) partes.push(`<path d="${d}" fill="none"${attrsTrazo(lnT)}/>`);
          // la ecuación y el R², junto al final de la línea (5 cifras como el formato General de Excel)
          const lineas = [tl.ec && aj.ec ? aj.ec : "", tl.r2 && aj.r2 != null ? `R² = ${cifras(aj.r2, 4)}` : ""].filter(Boolean);
          if (lineas.length && !aj.puntos) {
            const lt = L(tl.letra, { tam: 9 }), xe = xPant(x1) - anchoCat * 0.25, ye = aV(aj.f(x1));
            lineas.forEach((ln, j) => t(xe, ye + (j - (lineas.length - 1) / 2) * lt.tam * pt * 1.25 + lt.tam * pt * 0.35, ln, lt, "end"));
          }
        });
      });
    });
    marcar(null);
    escAct = esc_;
    // eje de categorías: su línea y sus rótulos
    const lineaC = trazo(ejeC.linea, { c: "#D9D9D9", w: 0.75 });
    escAct = esc_;
    const yCero = aV(Math.max(esc_.a, Math.min(esc_.b, 0)));
    if (lineaC && !ejeC.borrado && !escX) partes.push(horizontal ? `<line x1="${n2(yCero)}" y1="${n2(ptop)}" x2="${n2(yCero)}" y2="${n2(pb)}"${attrsTrazo(lineaC)}/>` : `<line x1="${n2(pl)}" y1="${n2(yCero)}" x2="${n2(pr)}" y2="${n2(yCero)}"${attrsTrazo(lineaC)}/>`);
    const lineaV = trazo(ejeV.linea, { no: 1 });
    if (lineaV && !ejeV.borrado) partes.push(horizontal ? `<line x1="${n2(pl)}" y1="${n2(pb)}" x2="${n2(pr)}" y2="${n2(pb)}"${attrsTrazo(lineaV)}/>` : `<line x1="${n2(pl)}" y1="${n2(ptop)}" x2="${n2(pl)}" y2="${n2(pb)}"${attrsTrazo(lineaV)}/>`);
    // ---- tabla de datos: categorías arriba y una fila por serie (con su clave), alineadas con las columnas
    if (m.tablaDatos && !horizontal && !escX) {
      const td = m.tablaDatos, ld = L(td.letra, { tam: 9 }), filaD = ld.tam * pt * 1.5, lnD = trazo(td.linea, { c: "#D9D9D9", w: 0.75 });
      const x0 = pl - (pl - izq), y0 = pb, filas = 1 + series.length, sep = (x1, y1, x2, y2) => partes.push(`<line x1="${n2(x1)}" y1="${n2(y1)}" x2="${n2(x2)}" y2="${n2(y2)}"${attrsTrazo(lnD)}/>`);
      cats.forEach((c, i) => t(aC(i), y0 + filaD * 0.5 + ld.tam * pt * 0.35, c, ld, "middle"));
      series.forEach(({ s }, k) => {
        const yc = y0 + filaD * (k + 1.5), v = datos(s);
        v.forEach((x, i) => { if (x != null) t(aC(i), yc + ld.tam * pt * 0.35, numero(x, s.vals.fmt, opc), ld, "middle"); });
        let xn = x0 + 3 * pt;
        if (td.claves) { const cl = s.relleno && s.relleno !== "none" ? s.relleno : (s.linea && s.linea.c) || colorSerie(k); partes.push(`<rect x="${n2(xn)}" y="${n2(yc - 3 * pt)}" width="${n2(6 * pt)}" height="${n2(6 * pt)}" fill="${s.degr ? degSvg(s.degr) : esc(cl)}"/>`); xn += 9 * pt; }
        t(xn, yc + ld.tam * pt * 0.35, nombreSerie(s, k), ld, "start");
      });
      if (lnD) {
        if (td.h) for (let k = 0; k <= filas; k++) { if (k === 0 && !td.contorno) continue; sep(k === 0 ? pl : x0, y0 + k * filaD, pr, y0 + k * filaD); }
        if (td.v) for (let i = 0; i <= nC; i++) { if ((i === 0 || i === nC) && !td.contorno && i !== 0) continue; sep(pl + i * anchoCat, y0, pl + i * anchoCat, y0 + filas * filaD); }
        if (td.contorno) sep(x0, y0 + filaD, x0, y0 + filas * filaD);
      }
    }
    if (verC && !escX && !(m.tablaDatos && !horizontal)) {
      const salto = ejeC.salto || Math.max(1, Math.ceil(Math.max(...cats.map((c) => medir(c, lc.tam * pt, fam(lc.f)))) / Math.max(1, anchoCat * (horizontal ? 9 : 1.6))));
      cats.forEach((c, i) => {
        if (i % (horizontal ? 1 : 1) !== 0) return;
        const p = aC(i);
        if (horizontal) t(pl - 4 * pt, p + lc.tam * pt * 0.35, c, lc, "end");
        else {
          const ancho = medir(c, lc.tam * pt, fam(lc.f));
          if (ancho <= anchoCat - 2 * pt || salto <= 1) t(p, pb + lc.tam * pt * 1.25, c, lc, "middle");
          else partes.push(`<text x="${n2(p)}" y="${n2(pb + lc.tam * pt * 0.9)}" font-family="${esc(fam(lc.f))}" font-size="${n2(lc.tam * pt)}" fill="${esc(lc.c)}" font-weight="${lc.b ? 700 : 400}" font-style="normal" text-anchor="end" transform="rotate(-45 ${n2(p)} ${n2(pb + lc.tam * pt * 0.9)})">${esc(c)}</text>`);
        }
      });
    }
    // títulos de los ejes
    if (ejeV.titulo && ejeV.titulo.texto) { const lt = L(ejeV.titulo.letra, { tam: 10, b: false }); if (horizontal) t((pl + pr) / 2, h - 7 * pt, ejeV.titulo.texto, lt, "middle"); else partes.push(`<text x="${n2(7 * pt + lt.tam * pt)}" y="${n2((ptop + pb) / 2)}" font-family="${esc(fam(lt.f))}" font-size="${n2(lt.tam * pt)}" fill="${esc(lt.c)}" font-weight="${lt.b ? 700 : 400}" font-style="normal" text-anchor="middle" transform="rotate(-90 ${n2(7 * pt + lt.tam * pt)} ${n2((ptop + pb) / 2)})">${esc(ejeV.titulo.texto)}</text>`); }
    if (ejeC.titulo && ejeC.titulo.texto) { const lt = L(ejeC.titulo.letra, { tam: 10 }); if (horizontal) partes.push(`<text x="${n2(7 * pt + lt.tam * pt)}" y="${n2((ptop + pb) / 2)}" font-family="${esc(fam(lt.f))}" font-size="${n2(lt.tam * pt)}" fill="${esc(lt.c)}" font-weight="${lt.b ? 700 : 400}" font-style="normal" text-anchor="middle" transform="rotate(-90 ${n2(7 * pt + lt.tam * pt)} ${n2((ptop + pb) / 2)})">${esc(ejeC.titulo.texto)}</text>`); else t((pl + pr) / 2, pb + (m.tablaDatos ? L(m.tablaDatos.letra, { tam: 9 }).tam * pt * 1.5 * (1 + series.length) : lc.tam * pt * 1.6) + lt.tam * pt * 1.1, ejeC.titulo.texto, lt, "middle"); }
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${n2(w)}" height="${n2(h)}" viewBox="0 0 ${n2(w)} ${n2(h)}">${partes.join("")}</svg>`;

    // ---- columnas 3D (Excel con ejes en ángulo recto): el frente sin deformar y la profundidad hacia arriba a la
    // derecha, (sen giroY, sen giroX · cos giroY) por unidad; cada serie en su fila de profundidad (agrupación
    // «estándar») o lado a lado; cada fila mide de fondo un ancho de categoría por la profundidad del archivo.
    // Caras como Excel: el frente con el color de la serie, arriba 0.76 y el lado 0.63 de ese color.
    function dibujarColumnas3D(g, x0, y0, x1, y1) {
      const v3 = m.vista3D || {}, rx = ((v3.rotX != null ? v3.rotX : 15) * Math.PI) / 180, ry = ((v3.rotY != null ? v3.rotY : 20) * Math.PI) / 180;
      const sx = Math.abs(Math.sin(ry)), sy = Math.abs(Math.sin(rx) * Math.cos(ry));
      const ss = g.series, nC = Math.max(1, cats.length), enProf = g.grupo === "standard", apil = g.grupo === "stacked" || g.grupo === "percentStacked";
      const nF = enProf ? Math.max(1, ss.length) : 1, dp = Math.max(0.2, (v3.prof != null ? v3.prof : 100) / 100);
      const vs = ss.map(datos);
      let dmin2 = 0, dmax2 = 0;
      if (apil) for (let i = 0; i < nC; i++) { let a = 0, b = 0; vs.forEach((v) => { const x = v[i] || 0; if (x >= 0) a += x; else b += x; }); if (g.grupo === "percentStacked") { const tt = a - b || 1; a /= tt; b /= tt; } dmax2 = Math.max(dmax2, a); dmin2 = Math.min(dmin2, b); }
      else vs.forEach((v) => v.forEach((x) => { if (x != null) { dmax2 = Math.max(dmax2, x); dmin2 = Math.min(dmin2, x); } }));
      const lv = L(ejeV.letra, { tam: 9 }), lc = L(ejeC.letra, { tam: 9 });
      const esc3 = escala(dmin2, dmax2, ejeV, g.grupo === "percentStacked", Math.floor(Math.max(40, y1 - y0) / (lv.tam * pt * 2.2)));
      const marcas = []; for (let v = esc3.a, k = 0; v <= esc3.b + esc3.u * 1e-6 && k < 200; v = esc3.a + ++k * esc3.u) marcas.push(Math.round(v * 1e10) / 1e10);
      const fmt3 = (v) => (g.grupo === "percentStacked" ? Math.round(v * 100) + "%" : numero(v, ejeV.ligado ? (ss[0] && ss[0].vals.fmt) || ejeV.fmt : ejeV.fmt, opc));
      const nombres = enProf ? ss.map((s2, k) => nombreSerie(s2, k)) : [];
      const verV = !ejeV.borrado && ejeV.rotulos !== "none", verC = !ejeC.borrado && ejeC.rotulos !== "none";
      const X0 = x0 + (verV ? Math.max(0, ...marcas.map((v) => medir(fmt3(v), lv.tam * pt, fam(lv.f)))) + 6 * pt : 0);
      const X1 = x1 - (nombres.length ? Math.max(0, ...nombres.map((nm) => medir(nm, lc.tam * pt, fam(lc.f)))) + 6 * pt : 0);
      const Y1 = y1 - (verC ? lc.tam * pt * 1.6 : 0), Y0 = y0 + lv.tam * pt * 0.6;
      // el frente y la profundidad caben juntos en el área de trazado
      // (Excel no estira la caja hacia arriba: el frente mide de alto unas 0.37 veces su ancho con la altura del
      // archivo al 100 %; la caja va centrada en lo que queda)
      const kz = (nF * dp) / nC, Wf = Math.max(10, (X1 - X0) / (1 + kz * sx)), Z = Wf * kz;
      const Hf = Math.max(10, Math.min(Y1 - Y0 - Z * sy, Wf * 0.37 * ((v3.alto != null ? v3.alto : 100) / 100)));
      const bajar = Math.max(0, (Y1 - Y0 - Z * sy - Hf) / 2);
      const yv = (v) => Y1 - bajar - ((v - esc3.a) / (esc3.b - esc3.a || 1)) * Hf;
      const P = (x, y, z) => `${n2(x + z * sx)},${n2(y - z * sy)}`;
      const lin = (a, b) => partes.push(`<polyline points="${a} ${b}" fill="none" stroke="#D9D9D9" stroke-width="${n2(0.75 * pt)}"/>`);
      // paredes: la cuadrícula en la pared del fondo y en la del lado izquierdo; los valores al frente
      marcas.forEach((v) => {
        const y = yv(v);
        if (ejeV.rejilla) { lin(P(X0, y, 0), P(X0, y, Z)); lin(P(X0, y, Z), P(X0 + Wf, y, Z)); }
        if (verV) t(X0 - 6 * pt, y + lv.tam * pt * 0.35, fmt3(v), lv, "end");
      });
      const cw = Wf / nC, fz = Z / nF;
      if (verC) cats.forEach((c, i) => t(X0 + cw * (i + 0.5), Y1 - bajar + lc.tam * pt * 1.25, c, lc, "middle"));
      // series: a la derecha del piso, en medio de su fila; la que no cabe junto a la anterior no se escribe
      let ultimo = Infinity;
      nombres.forEach((nm, k) => {
        const z = fz * (k + 0.5), px = X0 + Wf + z * sx + 6 * pt, py = Y1 - bajar - z * sy;
        if (ultimo - py < lc.tam * pt * 1.3) return;
        t(px, py + lc.tam * pt * 0.35, nm, lc); ultimo = py;
      });
      // columnas: de la fila del fondo a la del frente
      const gap = (g.hueco != null ? g.hueco : 150) / 100, gapD = (g.huecoProf != null ? g.huecoProf : 150) / 100;
      const bwTot = cw / (1 + gap), bw = enProf || apil ? bwTot : bwTot / Math.max(1, ss.length), bd = fz / (1 + gapD);
      const oscuro = (c, f) => (/^#[0-9A-Fa-f]{6}$/.test(c) ? "#" + [1, 3, 5].map((i) => Math.round(parseInt(c.slice(i, i + 2), 16) * f).toString(16).padStart(2, "0")).join("") : c);
      const caja = (x, yA, yB, z0, col) => {
        const z1 = z0 + bd, yt = Math.min(yA, yB), yb = Math.max(yA, yB);
        partes.push(`<polygon points="${P(x, yb, z0)} ${P(x + bw, yb, z0)} ${P(x + bw, yt, z0)} ${P(x, yt, z0)}" fill="${esc(col)}"/>`);
        partes.push(`<polygon points="${P(x, yt, z0)} ${P(x + bw, yt, z0)} ${P(x + bw, yt, z1)} ${P(x, yt, z1)}" fill="${esc(oscuro(col, 0.76))}"/>`);
        partes.push(`<polygon points="${P(x + bw, yb, z0)} ${P(x + bw, yb, z1)} ${P(x + bw, yt, z1)} ${P(x + bw, yt, z0)}" fill="${esc(oscuro(col, 0.63))}"/>`);
      };
      const base = yv(Math.max(esc3.a, Math.min(0, esc3.b)));
      const colorDe = (s2, k, i) => { const p2 = (s2.puntos || []).find((q) => q.idx === i); return p2 && p2.relleno && p2.relleno !== "none" ? p2.relleno : s2.relleno && s2.relleno !== "none" ? s2.relleno : colorSerie(series.findIndex((x) => x.s === s2)); };
      if (enProf) {
        for (let k = ss.length - 1; k >= 0; k--) vs[k].forEach((v, i) => { if (v == null) return; caja(X0 + cw * i + (cw - bw) / 2, base, yv(v), fz * k + (fz - bd) / 2, colorDe(ss[k], k, i)); });
      } else {
        for (let i = 0; i < nC; i++) {
          let a = 0, b = 0;
          const tot = g.grupo === "percentStacked" ? vs.reduce((s2, v) => s2 + Math.abs(v[i] || 0), 0) || 1 : 1;
          ss.forEach((s2, k) => {
            const v = vs[k][i]; if (v == null) return;
            const x = X0 + cw * i + (cw - bwTot) / 2 + (apil ? 0 : bw * k);
            if (apil) { const x2 = v / tot; if (x2 >= 0) { caja(x, yv(a), yv(a + x2), (fz - bd) / 2, colorDe(s2, k, i)); a += x2; } else { caja(x, yv(b), yv(b + x2), (fz - bd) / 2, colorDe(s2, k, i)); b += x2; } }
            else caja(x, base, yv(v), (fz - bd) / 2, colorDe(s2, k, i));
          });
        }
      }
    }

    // ---- pastel y anillo: desde las 12 en punto, en el sentido del reloj; borde blanco entre rebanadas
    // ---- circular 3D como Excel: la elipse aplastada por el giro en X, la pared de enfrente (de las 3 a las 9
    // pasando por las 6) de 0.23 del radio, con la luz de izquierda a derecha; encima, las rebanadas
    function dibujarPastel3D(g, se, x1, y1, x2, y2) {
      const v = datos(se).map((x) => Math.max(0, x || 0)), tot = v.reduce((a, b) => a + b, 0) || 1;
      const rotX = ((m.vista3D && m.vista3D.rotX != null ? m.vista3D.rotX : 30) * Math.PI) / 180, k = Math.max(0.1, Math.min(1, Math.sin(rotX) * 0.94)), grosor = 0.23;
      const rx = Math.max(1, Math.min(((x2 - x1) / 2) * 0.85, (y2 - y1) / (2 * k + grosor))), ry = rx * k, gz = rx * grosor;
      const cx = (x1 + x2) / 2, cy = y1 + (y2 - y1 - (2 * ry + gz)) / 2 + ry;
      const P = (t2, dy) => `${n2(cx + rx * Math.sin(t2))},${n2(cy - ry * Math.cos(t2) + (dy || 0))}`;
      const lnS = trazo(se.linea, { c: "#FFFFFF", w: 0.75 });
      const oscuro = (c, f) => (/^#[0-9A-Fa-f]{6}$/.test(c) ? "#" + [1, 3, 5].map((i) => Math.round(parseInt(c.slice(i, i + 2), 16) * f).toString(16).padStart(2, "0")).join("") : c);
      const uid = "nomp3d" + ++idsSvg;
      let a0 = ((g.giro || 0) * Math.PI) / 180;
      const reb = v.map((x, i) => { const a = (x / tot) * Math.PI * 2, r2 = { i, x, d: a0, h: a0 + a }; a0 += a; return r2; });
      reb.forEach((r2) => {
        if (!(r2.x > 0)) return;
        const col = colorPunto(se, r2.i);
        partes.push(`<defs><linearGradient id="${uid}-${r2.i}" gradientUnits="userSpaceOnUse" x1="${n2(cx - rx)}" y1="0" x2="${n2(cx + rx)}" y2="0"><stop offset="0" stop-color="${esc(oscuro(col, 0.85))}"/><stop offset="1" stop-color="${esc(oscuro(col, 0.42))}"/></linearGradient></defs>`);
        for (const w0 of [Math.PI / 2, Math.PI / 2 + Math.PI * 2, Math.PI / 2 - Math.PI * 2]) {
          const ta = Math.max(r2.d, w0), tb = Math.min(r2.h, w0 + Math.PI);
          if (tb - ta > 1e-4) partes.push(`<path d="M${P(ta)} A${n2(rx)},${n2(ry)} 0 0 1 ${P(tb)} L${P(tb, gz)} A${n2(rx)},${n2(ry)} 0 0 0 ${P(ta, gz)} Z" fill="url(#${uid}-${r2.i})"/>`);
        }
      });
      reb.forEach((r2) => {
        if (!(r2.x > 0)) return;
        const grande = r2.h - r2.d > Math.PI ? 1 : 0;
        const d = r2.h - r2.d >= Math.PI * 2 - 1e-6 ? `M${P(0)} A${n2(rx)},${n2(ry)} 0 1 1 ${P(Math.PI)} A${n2(rx)},${n2(ry)} 0 1 1 ${P(0)} Z` : `M${n2(cx)},${n2(cy)} L${P(r2.d)} A${n2(rx)},${n2(ry)} 0 ${grande} 1 ${P(r2.h)} Z`;
        partes.push(`<path d="${d}" fill="${esc(colorPunto(se, r2.i))}"${lnS ? attrsTrazo(lnS) : ""} stroke-linejoin="round"/>`);
        const ro = se.rotulos || g.rotulos;
        if (ro && (ro.val || ro.pct || ro.cat)) {
          const lr = L(ro.letra, { tam: 9 }), rm = ro.pos === "outEnd" ? 1.12 : ro.pos === "inEnd" ? 0.8 : 0.62, mid = (r2.d + r2.h) / 2;
          const txt = [ro.cat ? cats[r2.i] : "", ro.val ? numero(r2.x, ro.fmt || se.vals.fmt, opc) : "", ro.pct ? Math.round((r2.x / tot) * 100) + "%" : ""].filter(Boolean).join("; ");
          t(cx + rx * rm * Math.sin(mid), cy - ry * rm * Math.cos(mid) + lr.tam * pt * 0.35, txt, lr, "middle");
        }
      });
    }
    function dibujarPastel(g, se, x1, y1, x2, y2) {
      if (g.tresD && g.tipo === "pastel") return dibujarPastel3D(g, se, x1, y1, x2, y2);
      const v = datos(se).map((x) => Math.max(0, x || 0)), tot = v.reduce((a, b) => a + b, 0) || 1;
      const exMax = Math.max(se.explota || 0, ...(se.puntos || []).map((q) => q.explota || 0)) / 100;
      const cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, r = Math.max(1, Math.min(x2 - x1, y2 - y1) / 2 / (1 + exMax)), ri = g.tipo === "anillo" ? (r * Math.max(10, Math.min(90, g.agujero))) / 100 : 0;
      let ang = ((g.giro || 0) - 90) * (Math.PI / 180);
      const lnS = trazo(se.linea, { c: "#FFFFFF", w: 0.75 }), kS = series.findIndex((q) => q.s === se);
      v.forEach((x, i) => {
        marcar(kS, i);
        const a = (x / tot) * Math.PI * 2, b = ang + a, grande = a > Math.PI ? 1 : 0;
        const ex = ((se.puntos || []).find((q) => q.idx === i) || {}).explota || se.explota || 0, dx = (Math.cos(ang + a / 2) * r * ex) / 100, dy = (Math.sin(ang + a / 2) * r * ex) / 100;
        const p = (rr, t2) => [cx + dx + rr * Math.cos(t2), cy + dy + rr * Math.sin(t2)];
        const [ax, ay] = p(r, ang), [bx, by] = p(r, b);
        let d;
        if (a >= Math.PI * 2 - 1e-6) d = ri ? `M${n2(cx + r)},${n2(cy)} A${n2(r)},${n2(r)} 0 1 1 ${n2(cx - r)},${n2(cy)} A${n2(r)},${n2(r)} 0 1 1 ${n2(cx + r)},${n2(cy)} M${n2(cx + ri)},${n2(cy)} A${n2(ri)},${n2(ri)} 0 1 0 ${n2(cx - ri)},${n2(cy)} A${n2(ri)},${n2(ri)} 0 1 0 ${n2(cx + ri)},${n2(cy)} Z` : `M${n2(cx - r)},${n2(cy)} A${n2(r)},${n2(r)} 0 1 1 ${n2(cx + r)},${n2(cy)} A${n2(r)},${n2(r)} 0 1 1 ${n2(cx - r)},${n2(cy)} Z`;
        else if (ri) { const [cx2, cy2] = p(ri, b), [dx2, dy2] = p(ri, ang); d = `M${n2(ax)},${n2(ay)} A${n2(r)},${n2(r)} 0 ${grande} 1 ${n2(bx)},${n2(by)} L${n2(cx2)},${n2(cy2)} A${n2(ri)},${n2(ri)} 0 ${grande} 0 ${n2(dx2)},${n2(dy2)} Z`; }
        else d = `M${n2(cx + dx)},${n2(cy + dy)} L${n2(ax)},${n2(ay)} A${n2(r)},${n2(r)} 0 ${grande} 1 ${n2(bx)},${n2(by)} Z`;
        if (x > 0) partes.push(`<path d="${d}" fill="${esc(colorPunto(se, i))}"${lnS ? attrsTrazo(lnS) : ""} stroke-linejoin="round" fill-rule="evenodd"/>`);
        const ro = se.rotulos || g.rotulos;
        if (ro && (ro.val || ro.pct || ro.cat) && x > 0) {
          const lr = L(ro.letra, { tam: 9 }), rm = ri ? (r + ri) / 2 : ro.pos === "outEnd" ? r * 1.12 : ro.pos === "inEnd" ? r * 0.8 : r * 0.62, [tx, ty] = p(rm, ang + a / 2);
          const txt = [ro.cat ? cats[i] : "", ro.val ? numero(x, ro.fmt || se.vals.fmt, opc) : "", ro.pct ? Math.round((x / tot) * 100) + "%" : ""].filter(Boolean).join("; ");
          t(tx, ty + lr.tam * pt * 0.35, txt, lr, "middle");
        }
        ang = b;
      });
      marcar(null);
    }
  }

  window.NOM = window.NOM || {};
  // ================================================================ gráficas nuevas de Excel (cx:chartSpace)
  // Histograma, Pareto, cajas y bigotes, cascada, rectángulos, proyección solar y embudo. Excel no guarda sus
  // valores: los toma de nombres ocultos («_xlchart.v1.N») que apuntan a la hoja. Aquí se leen vivos y se calculan
  // como Excel: intervalos por la regla de Scott, cuartiles exclusivos o inclusivos, rectángulos «cuadrados»…
  const TIPOS_EX = { clusteredColumn: "histograma", boxWhisker: "caja", waterfall: "cascada", treemap: "rectangulos", sunburst: "solar", funnel: "embudo" };
  function leerEx(xml, ctx) {
    ctx = ctx || {};
    const d = new DOMParser().parseFromString(xml, "application/xml");
    const todos = (n, nombre) => (n ? [...n.getElementsByTagNameNS("*", nombre)] : []);
    const h1 = (n, nombre) => [...(n ? n.children : [])].find((c) => c.localName === nombre) || null;
    const at = (n, a) => (n ? n.getAttribute(a) : null);
    const num = (x) => (x == null || x === "" || !isFinite(+x) ? null : +x);
    const ref = (f) => {
      const tx = String((f && f.textContent) || "").trim();
      if (!tx) return null;
      const r = /^_xlchart\./i.test(tx) ? (ctx.nombre ? ctx.nombre(tx) : null) : tx;
      return r ? String(r).replace(/^=/, "") : null;
    };
    const textoDe = (tx) => { if (!tx) return null; const v = todos(tx, "v")[0]; if (v) return v.textContent; const ts = todos(tx, "t"); return ts.length ? ts.map((q) => q.textContent).join("") : null; };
    const chart = todos(d, "chart")[0], series = todos(d, "series");
    const s0 = series.find((s) => at(s, "layoutId") !== "paretoLine") || series[0];
    if (!s0) throw new Error("gráfica sin series");
    const tipo = series.some((s) => at(s, "layoutId") === "paretoLine") ? "pareto" : TIPOS_EX[at(s0, "layoutId")];
    if (!tipo) throw new Error("tipo de gráfica que no se conoce");
    const id = at(h1(s0, "dataId"), "val") || "0";
    const data = todos(d, "data").find((x) => at(x, "id") === id) || todos(d, "data")[0];
    // lo guardado (si lo hay): cada nivel con sus puntos; el primero es el de las hojas del árbol
    const dim = (tp) => {
      const e = [...(data ? data.children : [])].find((x) => /Dim$/.test(x.localName) && at(x, "type") === tp);
      if (!e) return null;
      const niveles = todos(e, "lvl").map((l) => { const o = []; todos(l, "pt").forEach((p) => { o[+at(p, "idx") || 0] = p.textContent; }); return o; });
      return { f: ref(h1(e, "f")), niveles };
    };
    const tit = h1(chart, "title"), ley = h1(chart, "legend"), txS = h1(s0, "tx"), txD = txS && todos(txS, "txData")[0];
    const lp = h1(s0, "layoutPr"), bn = lp && h1(lp, "binning"), stt = lp && h1(lp, "statistics"), vis = lp && h1(lp, "visibility"), sub = lp && h1(lp, "subtotals");
    const bool = (n, a, def) => { const v = at(n, a); return v == null ? def : v === "1" || v === "true"; };
    const ejes = todos(d, "axis").map((a) => {
      const vs = h1(a, "valScaling"), cs = h1(a, "catScaling");
      return { val: !!vs, min: num(at(vs, "min")), max: num(at(vs, "max")), unidad: num(at(vs, "majorUnit")), hueco: num(at(cs, "gapWidth")), rejilla: !!h1(a, "majorGridlines"),
        rotulos: !!h1(a, "tickLabels"), oculto: at(a, "hidden") === "1", pct: todos(a, "units").some((u) => at(u, "unit") === "percentage") };
    });
    return {
      ex: true, tipo,
      titulo: tit ? { texto: textoDe(h1(tit, "tx")) } : null,
      leyenda: ley ? { pos: at(ley, "pos") || "t" } : null,
      nombre: { f: ref(txD && h1(txD, "f")), v: txD ? textoDe(txD) : null },
      cat: dim("cat"), val: dim("val") || dim("size"),
      bins: bn ? { cerrado: at(bn, "intervalClosed") || "r", ancho: num(at(h1(bn, "binSize"), "val")), cantidad: num(at(h1(bn, "binCount"), "val")), arriba: num(at(bn, "overflow")), abajo: num(at(bn, "underflow")) } : null,
      cuartil: at(stt, "quartileMethod") || "exclusive",
      vis: { media: bool(vis, "meanMarker", true), lineaMedia: bool(vis, "meanLine", false), puntos: bool(vis, "nonoutliers", true), atipicos: bool(vis, "outliers", true) },
      totales: sub ? todos(sub, "idx").map((x) => +at(x, "val")).filter((x) => isFinite(x)) : [],
      ejes,
      acentos: ctx.acentos || ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"],
      letra: { f: ctx.fuente ? ctx.fuente("+mn-lt") : null },
    };
  }

  /** Rectángulos «cuadrados» (Bruls, Huizing y van Wijk), como los acomoda Excel: renglones a lo largo del lado corto. */
  function cuadricular(valores, x, y, w, h) {
    const total = valores.reduce((s, v) => s + v, 0), out = [];
    if (!(total > 0) || !(w > 0) || !(h > 0)) return valores.map(() => ({ x, y, w: 0, h: 0 }));
    const area = valores.map((v) => (v / total) * w * h);
    let i = 0;
    const peor = (fila, lado) => { const s = fila.reduce((a, b) => a + b, 0), mx = Math.max(...fila), mn = Math.min(...fila); return Math.max((lado * lado * mx) / (s * s), (s * s) / (lado * lado * mn)); };
    while (i < area.length) {
      const lado = Math.min(w, h), fila = [area[i]];
      let j = i + 1;
      while (j < area.length && peor(fila.concat(area[j]), lado) <= peor(fila, lado)) fila.push(area[j++]);
      const s = fila.reduce((a, b) => a + b, 0);
      if (w >= h) {   // franja vertical a la izquierda
        const fw = s / h; let yy = y;
        fila.forEach((a) => { out.push({ x, y: yy, w: fw, h: a / fw }); yy += a / fw; });
        x += fw; w -= fw;
      } else {        // franja horizontal arriba
        const fh = s / w; let xx = x;
        fila.forEach((a) => { out.push({ x: xx, y, w: a / fh, h: fh }); xx += a / fh; });
        y += fh; h -= fh;
      }
      i = j;
    }
    return out;
  }
  /** Cuartil como Excel: CUARTIL.EXC (posición p·(n+1)) o CUARTIL.INC (posición p·(n−1)), con interpolación. */
  function cuartil(ord, p, exclusivo) {
    const n = ord.length;
    if (!n) return null;
    let pos = exclusivo ? p * (n + 1) - 1 : p * (n - 1);
    pos = Math.max(0, Math.min(n - 1, pos));
    const i = Math.floor(pos), f = pos - i;
    return i + 1 < n ? ord[i] + f * (ord[i + 1] - ord[i]) : ord[i];
  }

  function svgEx(m, w, h, opc) {
    opc = opc || {};
    const pt = opc.pt || 1, fam = (f) => (opc.familia ? opc.familia(f || m.letra.f || "Calibri") : `'${(f || m.letra.f || "Calibri").replace(/'/g, "")}', sans-serif`);
    const partes = [], L = (def) => Object.assign({ tam: 9, c: "#595959", b: false, f: m.letra.f }, def || {});
    const t = (x, y, s, l, anc, extra) => partes.push(`<text x="${n2(x)}" y="${n2(y)}" font-family="${esc(fam(l.f))}" font-size="${n2(l.tam * pt)}" fill="${esc(l.c)}" font-weight="${l.b ? 700 : 400}" font-style="normal" text-anchor="${anc || "start"}"${extra || ""}>${esc(s)}</text>`);
    const ok = (c) => /^#[0-9A-Fa-f]{6}$/.test(c || "");
    const acc = (k) => { const a = (m.acentos || [])[k % 6]; return ok(a) ? a : ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"][k % 6]; };
    const oscuro = (c, f) => "#" + [1, 3, 5].map((i) => Math.round(parseInt(c.slice(i, i + 2), 16) * f).toString(16).padStart(2, "0")).join("");
    partes.push(`<rect x="0.5" y="0.5" width="${n2(w - 1)}" height="${n2(h - 1)}" fill="#FFFFFF"${opc.sinBorde ? "" : ` stroke="#D9D9D9" stroke-width="${n2(Math.max(0.25, 0.75 * pt))}"`}/>`);
    // ---- los datos, vivos de la hoja (o lo guardado)
    const colsDe = (f) => { const q = /\$?([A-Z]{1,3})\$?\d+:\$?([A-Z]{1,3})\$?\d+$/.exec(String(f || "")); if (!q) return 1; const nc = (s) => s.split("").reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0); return Math.max(1, nc(q[2]) - nc(q[1]) + 1); };
    const filasDe = (dm) => {
      if (!dm) return [];
      const vivo = opc.valores && dm.f ? opc.valores(dm.f) : null;
      if (vivo) { const k = colsDe(dm.f), o = []; for (let i = 0; i < vivo.length; i += k) o.push(vivo.slice(i, i + k)); return o; }
      const nv = dm.niveles || [], n = Math.max(0, ...nv.map((l) => l.length));
      return Array.from({ length: n }, (_, i) => nv.slice().reverse().map((l) => (l[i] == null ? null : l[i])));
    };
    const numero2 = (v) => (v == null || v === "" || !isFinite(+v) ? null : +v);
    const vals = filasDe(m.val).map((r) => numero2(r[r.length - 1]));
    const catsF = filasDe(m.cat), cats = catsF.map((r) => r.filter((x) => x != null && x !== "").map(String));
    const nombreSerie = (() => { const v = opc.valores && m.nombre && m.nombre.f ? opc.valores(m.nombre.f) : null; return v && v[0] != null ? String(v[0]) : (m.nombre && m.nombre.v) || ""; })();
    const fmtN = (v) => { const r = Math.round(v * 1e6) / 1e6; return numero(r, null, opc); };
    // ---- título y leyenda
    let arriba = 7 * pt, abajo = h - 7 * pt, izq = 7 * pt, der = w - 7 * pt;
    if (m.titulo) {
      const lt = L({ tam: 14 }), tx = m.titulo.texto != null ? m.titulo.texto : nombreSerie || "Título del gráfico";
      const y = arriba + lt.tam * pt * 0.95;
      t(w / 2, y, tx, lt, "middle");
      arriba = y + lt.tam * pt * 0.45 + 4 * pt;
    }
    const leyenda = (entradas) => {
      if (!m.leyenda || !entradas.length) return;
      const ll = L({ tam: 9 }), caja = 7 * pt, sep = 4 * pt, gap = 10 * pt, altoF = ll.tam * pt * 1.5;
      const anchoE = (e) => caja + sep + medir(e.texto, ll.tam * pt, fam(ll.f));
      const total = entradas.reduce((s, e) => s + anchoE(e), 0) + gap * (entradas.length - 1);
      const y0 = m.leyenda.pos === "b" ? abajo - altoF : arriba, yc = y0 + altoF / 2;
      let x = w / 2 - total / 2;
      entradas.forEach((e) => { partes.push(`<rect x="${n2(x)}" y="${n2(yc - caja / 2)}" width="${n2(caja)}" height="${n2(caja)}" fill="${esc(e.color)}"/>`); t(x + caja + sep, yc + ll.tam * pt * 0.35, e.texto, ll); x += anchoE(e) + gap; });
      if (m.leyenda.pos === "b") abajo = y0 - 4 * pt; else arriba = y0 + altoF + 4 * pt;
    };
    const fin = () => `<svg xmlns="http://www.w3.org/2000/svg" width="${n2(w)}" height="${n2(h)}" viewBox="0 0 ${n2(w)} ${n2(h)}">${partes.join("")}</svg>`;
    const ejeV = m.ejes.find((e) => e.val) || {}, ejeC = m.ejes.find((e) => !e.val) || {};
    /** Ejes de columnas: valores a la izquierda (y a la derecha en Pareto), categorías abajo; devuelve dónde pintar. */
    const marco = (dmin, dmax, etiquetas, derecho) => {
      const lv = L(), lc = L(), alto = Math.max(40, abajo - arriba - lc.tam * pt * 1.6);
      const esc_ = escala(dmin, dmax, ejeV, false, Math.floor(alto / (lv.tam * pt * 2.2)));
      const marcas = []; for (let v = esc_.a, k = 0; v <= esc_.b + esc_.u * 1e-6 && k < 200; v = esc_.a + ++k * esc_.u) marcas.push(Math.round(v * 1e10) / 1e10);
      const verV = ejeV.rotulos !== false && !ejeV.oculto;
      let x0 = izq + (verV ? Math.max(...marcas.map((v) => medir(fmtN(v), lv.tam * pt, fam(lv.f)))) + 6 * pt : 0), x1 = der;
      const y0 = arriba + lv.tam * pt * 0.6, y1 = abajo - lc.tam * pt * 1.6;
      let marcasD = null;
      if (derecho) {
        marcasD = []; for (let k = 0; k <= 10; k++) marcasD.push(k / 10);
        x1 -= Math.max(...marcasD.map((v) => medir(Math.round(v * 100) + "%", lv.tam * pt, fam(lv.f)))) + 6 * pt;
      }
      const yv = (v) => y1 - ((v - esc_.a) / (esc_.b - esc_.a || 1)) * (y1 - y0);
      marcas.forEach((v) => {
        if (ejeV.rejilla !== false) partes.push(`<line x1="${n2(x0)}" y1="${n2(yv(v))}" x2="${n2(x1)}" y2="${n2(yv(v))}" stroke="#D9D9D9" stroke-width="${n2(0.75 * pt)}"/>`);
        if (verV) t(x0 - 6 * pt, yv(v) + lv.tam * pt * 0.35, fmtN(v), lv, "end");
      });
      if (marcasD) marcasD.forEach((v) => t(x1 + 6 * pt, y1 - v * (y1 - y0) + lv.tam * pt * 0.35, Math.round(v * 100) + "%", lv));
      const anchoC = (x1 - x0) / Math.max(1, etiquetas.length);
      if (ejeC.rotulos !== false && !ejeC.oculto) etiquetas.forEach((e, i) => t(x0 + anchoC * (i + 0.5), y1 + lc.tam * pt * 1.25, e, lc, "middle"));
      return { x0, x1, y0, y1, yv, anchoC };
    };
    const tipo = m.tipo;

    if (tipo === "histograma" || tipo === "pareto") {
      let barras;   // [{ etiqueta, valor }]
      const b = m.bins;
      if (b || !cats.length) {
        const datos = vals.filter((x) => x != null);
        if (!datos.length) return fin();
        const n = datos.length, min = Math.min(...datos), max = Math.max(...datos);
        let ancho = b && b.ancho, k = b && b.cantidad;
        if (!ancho && k) ancho = (max - min) / k;
        if (!ancho) {
          // regla de Scott (desviación de muestra), con dos cifras significativas como las pinta Excel
          const media = datos.reduce((s, x) => s + x, 0) / n, sd = Math.sqrt(datos.reduce((s, x) => s + (x - media) * (x - media), 0) / Math.max(1, n - 1));
          const crudo = (3.5 * sd) / Math.cbrt(n);
          ancho = crudo > 0 ? +crudo.toPrecision(2) : 1;
        }
        const lo0 = b && b.abajo != null ? b.abajo : min, hi0 = b && b.arriba != null ? b.arriba : max;
        const nb = Math.max(1, k || Math.ceil((hi0 - lo0) / ancho - 1e-9));
        const izqC = !b || b.cerrado !== "l";
        barras = Array.from({ length: nb }, (_, i) => {
          const a = lo0 + i * ancho, z = lo0 + (i + 1) * ancho;
          const etiqueta = izqC ? `${i === 0 && !(b && b.abajo != null) ? "[" : "("}${fmtN(a)}, ${fmtN(z)}]` : `[${fmtN(a)}, ${fmtN(z)}${i === nb - 1 && !(b && b.arriba != null) ? "]" : ")"}`;
          return { etiqueta, valor: datos.filter((x) => (izqC ? (x > a || (i === 0 && x >= a)) && x <= z + 1e-9 : x >= a - 1e-9 && (x < z || (i === nb - 1 && x <= z))) && x >= lo0 - 1e-9 && x <= hi0 + 1e-9).length };
        });
        if (b && b.abajo != null) barras.unshift({ etiqueta: `≤${fmtN(b.abajo)}`, valor: datos.filter((x) => x <= b.abajo).length });
        if (b && b.arriba != null) barras.push({ etiqueta: `>${fmtN(b.arriba)}`, valor: datos.filter((x) => x > b.arriba).length });
      } else barras = cats.map((c, i) => ({ etiqueta: c.join(" "), valor: vals[i] || 0 }));
      if (tipo === "pareto") barras = barras.map((x, i) => Object.assign(x, { i })).sort((p, q) => q.valor - p.valor || p.i - q.i);
      const mx = Math.max(0, ...barras.map((x) => x.valor));
      const f = marco(0, mx, barras.map((x) => x.etiqueta), tipo === "pareto");
      const hueco = ejeC.hueco != null ? ejeC.hueco : 0, bw = f.anchoC / (1 + hueco), col = acc(0);
      barras.forEach((x, i) => {
        const bx = f.x0 + f.anchoC * i + (f.anchoC - bw) / 2, y = f.yv(x.valor);
        partes.push(`<rect x="${n2(bx)}" y="${n2(y)}" width="${n2(bw)}" height="${n2(Math.max(0, f.y1 - y))}" fill="${col}" stroke="#FFFFFF" stroke-width="${n2(0.75 * pt)}"/>`);
      });
      partes.push(`<line x1="${n2(f.x0)}" y1="${n2(f.y1)}" x2="${n2(f.x1)}" y2="${n2(f.y1)}" stroke="#D9D9D9" stroke-width="${n2(0.75 * pt)}"/>`);
      if (tipo === "pareto") {
        const tot = barras.reduce((s, x) => s + x.valor, 0) || 1;
        let acum = 0;
        const pts = barras.map((x, i) => { acum += x.valor; return `${n2(f.x0 + f.anchoC * (i + 0.5))},${n2(f.y1 - (acum / tot) * (f.y1 - f.y0))}`; });
        partes.push(`<polyline points="${pts.join(" ")}" fill="none" stroke="${acc(1)}" stroke-width="${n2(2.25 * pt)}" stroke-linejoin="round" stroke-linecap="round"/>`);
      }
      return fin();
    }

    if (tipo === "caja") {
      // grupos por categoría (o uno solo, «1»)
      const grupos = new Map();
      vals.forEach((v, i) => { if (v == null) return; const k = cats.length ? (cats[i] || []).join(" ") : "1"; if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(v); });
      if (!grupos.size) return fin();
      const todosV = [...grupos.values()].flat();
      const f = marco(Math.min(...todosV), Math.max(...todosV), [...grupos.keys()], false);
      const hueco = ejeC.hueco != null ? ejeC.hueco : 1, bw = f.anchoC / (1 + hueco), col = acc(0), borde = oscuro(col, 0.7), sw = n2(Math.max(0.5, 0.75 * pt));
      [...grupos.values()].forEach((g, i) => {
        const ord = g.slice().sort((a, b) => a - b), exc = m.cuartil !== "inclusive";
        const q1 = cuartil(ord, 0.25, exc), q2 = cuartil(ord, 0.5, exc), q3 = cuartil(ord, 0.75, exc), iqr = q3 - q1;
        const dentro = ord.filter((x) => x >= q1 - 1.5 * iqr - 1e-9 && x <= q3 + 1.5 * iqr + 1e-9), fuera = ord.filter((x) => !dentro.includes(x));
        const lo = Math.min(...dentro), hi = Math.max(...dentro), media = ord.reduce((s, x) => s + x, 0) / ord.length;
        const cx = f.x0 + f.anchoC * (i + 0.5), bx = cx - bw / 2, tapa = bw * 0.3;
        partes.push(`<line x1="${n2(cx)}" y1="${n2(f.yv(hi))}" x2="${n2(cx)}" y2="${n2(f.yv(lo))}" stroke="${borde}" stroke-width="${sw}"/>`);
        [hi, lo].forEach((v) => partes.push(`<line x1="${n2(cx - tapa / 2)}" y1="${n2(f.yv(v))}" x2="${n2(cx + tapa / 2)}" y2="${n2(f.yv(v))}" stroke="${borde}" stroke-width="${sw}"/>`));
        partes.push(`<rect x="${n2(bx)}" y="${n2(f.yv(q3))}" width="${n2(bw)}" height="${n2(Math.max(0, f.yv(q1) - f.yv(q3)))}" fill="${col}" stroke="${borde}" stroke-width="${sw}"/>`);
        partes.push(`<line x1="${n2(bx)}" y1="${n2(f.yv(q2))}" x2="${n2(bx + bw)}" y2="${n2(f.yv(q2))}" stroke="${borde}" stroke-width="${sw}"/>`);
        const punto = (v) => partes.push(`<circle cx="${n2(cx)}" cy="${n2(f.yv(v))}" r="${n2(2 * pt)}" fill="${col}" stroke="${borde}" stroke-width="${sw}"/>`);
        if (m.vis.puntos) dentro.filter((x) => x !== lo && x !== hi).forEach(punto);   // (en las puntas del bigote no: ahí está la tapa)
        if (m.vis.atipicos) fuera.forEach(punto);
        if (m.vis.media) { const r = 3.5 * pt, y = f.yv(media); partes.push(`<path d="M${n2(cx - r)} ${n2(y - r)}L${n2(cx + r)} ${n2(y + r)}M${n2(cx + r)} ${n2(y - r)}L${n2(cx - r)} ${n2(y + r)}" stroke="${borde}" stroke-width="${n2(1.1 * pt)}" fill="none"/>`); }
      });
      partes.push(`<line x1="${n2(f.x0)}" y1="${n2(f.y1)}" x2="${n2(f.x1)}" y2="${n2(f.y1)}" stroke="#D9D9D9" stroke-width="${n2(0.75 * pt)}"/>`);
      return fin();
    }

    if (tipo === "cascada") {
      leyenda([{ texto: "Aumento", color: acc(0) }, { texto: "Disminución", color: acc(1) }, { texto: "Total", color: acc(2) }]);
      let corre = 0;
      const pasos = vals.map((v, i) => {
        const x = v || 0;
        if (m.totales.includes(i)) { corre = x; return { de: 0, a: x, tipo: 2 }; }
        const de = corre; corre += x; return { de, a: corre, tipo: x < 0 ? 1 : 0 };
      });
      const extremos = pasos.flatMap((p) => [p.de, p.a]).concat(0);
      const f = marco(Math.min(...extremos), Math.max(...extremos), vals.map((_, i) => (cats[i] || []).join(" ") || String(i + 1)), false);
      const hueco = ejeC.hueco != null ? ejeC.hueco : 0.5, bw = f.anchoC / (1 + hueco);
      pasos.forEach((p, i) => {
        const bx = f.x0 + f.anchoC * i + (f.anchoC - bw) / 2, ya = f.yv(Math.max(p.de, p.a)), yb = f.yv(Math.min(p.de, p.a));
        partes.push(`<rect x="${n2(bx)}" y="${n2(ya)}" width="${n2(bw)}" height="${n2(Math.max(0.5, yb - ya))}" fill="${acc(p.tipo)}"/>`);
        // la línea que une cada columna con la siguiente, a la altura donde quedó
        if (i + 1 < pasos.length) partes.push(`<line x1="${n2(bx + bw)}" y1="${n2(f.yv(p.a))}" x2="${n2(bx + f.anchoC)}" y2="${n2(f.yv(p.a))}" stroke="#D9D9D9" stroke-width="${n2(0.75 * pt)}"/>`);
      });
      partes.push(`<line x1="${n2(f.x0)}" y1="${n2(f.y0)}" x2="${n2(f.x0)}" y2="${n2(f.y1)}" stroke="#D9D9D9" stroke-width="${n2(0.75 * pt)}"/>`);
      partes.push(`<line x1="${n2(f.x0)}" y1="${n2(f.y1)}" x2="${n2(f.x1)}" y2="${n2(f.y1)}" stroke="#D9D9D9" stroke-width="${n2(0.75 * pt)}"/>`);
      return fin();
    }

    if (tipo === "embudo") {
      const etiquetas = vals.map((_, i) => (cats[i] || []).join(" ") || String(i + 1)), lc = L();
      const x0 = izq + Math.max(0, ...etiquetas.map((e) => medir(e, lc.tam * pt, fam(lc.f)))) + 8 * pt, x1 = der, y0 = arriba, y1 = abajo;
      const mx = Math.max(0, ...vals.map((v) => Math.abs(v || 0))) || 1, altoC = (y1 - y0) / Math.max(1, vals.length);
      const hueco = ejeC.hueco != null ? ejeC.hueco : 0.06, bh = altoC / (1 + hueco), cx = (x0 + x1) / 2;
      vals.forEach((v, i) => {
        const bw = (Math.abs(v || 0) / mx) * (x1 - x0), y = y0 + altoC * i + (altoC - bh) / 2;
        partes.push(`<rect x="${n2(cx - bw / 2)}" y="${n2(y)}" width="${n2(bw)}" height="${n2(bh)}" fill="${acc(0)}"/>`);
        t(x0 - 8 * pt, y0 + altoC * (i + 0.5) + lc.tam * pt * 0.35, etiquetas[i], lc, "end");
      });
      partes.push(`<line x1="${n2(x0)}" y1="${n2(y0)}" x2="${n2(x0)}" y2="${n2(y1)}" stroke="#D9D9D9" stroke-width="${n2(0.75 * pt)}"/>`);
      return fin();
    }

    if (tipo === "rectangulos" || tipo === "solar") {
      // el árbol: primer nivel (color) → hojas, con sus sumas; de mayor a menor como Excel
      const raiz = { hijos: [], valor: 0 };
      vals.forEach((v, i) => {
        if (v == null || !(v > 0)) return;
        let nodo = raiz;
        (cats[i] && cats[i].length ? cats[i] : [String(i + 1)]).forEach((nm) => {
          let h2 = nodo.hijos.find((x) => x.nombre === nm);
          if (!h2) { h2 = { nombre: nm, hijos: [], valor: 0, orden: nodo.hijos.length }; nodo.hijos.push(h2); }
          h2.valor += v; nodo = h2;
        });
        raiz.valor += v;
      });
      const grupos = raiz.hijos;
      grupos.forEach((g, k) => { g.color = acc(k); });
      const pintarColor = (nodo, c) => { nodo.color = nodo.color || c; nodo.hijos.forEach((x) => pintarColor(x, nodo.color)); };
      grupos.forEach((g) => pintarColor(g, g.color));
      const ordenar = (nodo) => { nodo.hijos.sort((a, b) => b.valor - a.valor || a.orden - b.orden); nodo.hijos.forEach(ordenar); };
      if (tipo === "rectangulos") leyenda(grupos.map((g) => ({ texto: g.nombre, color: g.color })));
      ordenar(raiz);
      const linea = n2(Math.max(0.75, 1.5 * pt));
      if (tipo === "rectangulos") {
        const dibujar = (nodo, r) => {
          if (!nodo.hijos.length) { partes.push(`<rect x="${n2(r.x)}" y="${n2(r.y)}" width="${n2(r.w)}" height="${n2(r.h)}" fill="${nodo.color}" stroke="#FFFFFF" stroke-width="${linea}"/>`); return; }
          cuadricular(nodo.hijos.map((x) => x.valor), r.x, r.y, r.w, r.h).forEach((q, i) => dibujar(nodo.hijos[i], q));
        };
        dibujar(raiz, { x: izq, y: arriba, w: der - izq, h: abajo - arriba });
        return fin();
      }
      // proyección solar: anillos del centro hacia afuera, desde las 12 y en el sentido del reloj
      const prof = (nodo) => (nodo.hijos.length ? 1 + Math.max(...nodo.hijos.map(prof)) : 0), niveles = Math.max(1, prof(raiz));
      const cx = (izq + der) / 2, cy = (arriba + abajo) / 2, R = Math.max(1, Math.min(der - izq, abajo - arriba) / 2), grueso = R / (niveles + 1);
      const p = (r, a) => `${n2(cx + r * Math.sin(a))} ${n2(cy - r * Math.cos(a))}`;
      const gajo = (r0, r1, a0, a1, c) => {
        if (a1 - a0 >= Math.PI * 2 - 1e-6) a1 = a0 + Math.PI * 2 - 1e-4;
        const g = a1 - a0 > Math.PI ? 1 : 0;
        partes.push(`<path d="M${p(r1, a0)}A${n2(r1)} ${n2(r1)} 0 ${g} 1 ${p(r1, a1)}L${p(r0, a1)}A${n2(r0)} ${n2(r0)} 0 ${g} 0 ${p(r0, a0)}Z" fill="${c}" stroke="#FFFFFF" stroke-width="${linea}"/>`);
      };
      const anillo = (nodo, nivel, a0, a1) => {
        let a = a0;
        nodo.hijos.forEach((x) => { const b2 = a + ((a1 - a0) * x.valor) / (nodo.valor || 1); gajo(grueso * (nivel + 1), grueso * (nivel + 2), a, b2, x.color); anillo(x, nivel + 1, a, b2); a = b2; });
      };
      anillo(raiz, 0, 0, Math.PI * 2);
      return fin();
    }
    return fin();
  }

  NOM.graficaOffice = { leer, leerEx, svg, escala };
})();

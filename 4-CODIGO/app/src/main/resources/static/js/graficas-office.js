/* NOM-COLAB · Gráficas de Office (DrawingML «c:chartSpace») dibujadas como Office. El mismo dibujante para
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
    const linea = (spPr) => {
      const ln = spPr && h1(spPr, "ln");
      if (!ln) return null;
      if (h1(ln, "noFill")) return { no: 1 };
      const o = {};
      if (at(ln, "w")) o.w = +at(ln, "w") / 12700;
      const sf = h1(ln, "solidFill"); if (sf) o.c = color(sf);
      const gu = h1(ln, "prstDash"); if (gu && at(gu, "val") !== "solid") o.guion = at(gu, "val");
      return o;
    };
    const letra = (txPr) => {
      const r = txPr && todos(txPr, "defRPr")[0];
      if (!r) return null;
      const o = {};
      if (at(r, "sz")) o.tam = +at(r, "sz") / 100;
      if (at(r, "b") != null) o.b = at(r, "b") === "1";
      if (at(r, "i") != null) o.i = at(r, "i") === "1";
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
    m.area = { relleno: relleno(h1(cs, "spPr")), linea: linea(h1(cs, "spPr")), redondo: si(cs, "roundedCorners", false) };
    m.trazado = { relleno: relleno(h1(pa, "spPr")), linea: linea(h1(pa, "spPr")), disp: disposicion(pa) };
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
      hs(g, "ser").forEach((s) => {
        const sp = h1(s, "spPr"), mk = h1(s, "marker");
        const se = { idx: +val(s, "idx", 0), orden: +val(s, "order", 0), nombre: textoRico(h1(s, "tx")), relleno: relleno(sp), linea: linea(sp),
          cats: cache(h1(s, "cat") || h1(s, "xVal")), vals: cache(h1(s, "val") || h1(s, "yVal")), suave: si(s, "smooth", false), rotulos: rotulos(s),
          explota: +val(s, "explosion", 0), invierte: si(s, "invertIfNegative", false) };
        if (mk) se.marcador = { simbolo: val(mk, "symbol", "auto"), tam: +val(mk, "size", 5), relleno: relleno(h1(mk, "spPr")), linea: linea(h1(mk, "spPr")) };
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
    // (peso y estilo siempre escritos: la gráfica no hereda la negrita del lugar donde se pinta)
    const t = (x, y, s, l, anc, extra) => partes.push(`<text x="${n2(x)}" y="${n2(y)}" font-family="${esc(fam(l.f))}" font-size="${n2(l.tam * pt)}" fill="${esc(l.c || "#595959")}" font-weight="${l.b ? 700 : 400}" font-style="${l.i ? "italic" : "normal"}" text-anchor="${anc || "start"}"${extra || ""}>${esc(s)}</text>`);
    const trazo = (ln, def) => { if (ln && ln.no) return null; const o = Object.assign({ w: 0.75, c: "#D9D9D9" }, def || {}, ln || {}); return o; };
    const attrsTrazo = (o) => (o ? ` stroke="${esc(o.c || "#000")}" stroke-width="${n2(Math.max(0.25, o.w * pt))}"${o.guion ? ` stroke-dasharray="${n2(3 * o.w * pt)} ${n2(2 * o.w * pt)}"` : ""}` : ' stroke="none"');
    // fondo de la gráfica
    const fa = m.area.relleno === "none" ? "none" : m.area.relleno || "#FFFFFF", la = trazo(m.area.linea, opc.sinBorde ? { no: 1 } : { c: "#D9D9D9", w: 0.75 });
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
        if (e.tipo === "linea") { partes.push(`<line x1="${n2(x)}" y1="${n2(yc)}" x2="${n2(x + lineaW)}" y2="${n2(yc)}" stroke="${esc(e.color)}" stroke-width="${n2(2.25 * pt)}" stroke-linecap="round"/>`); if (e.marca) partes.push(`<circle cx="${n2(x + lineaW / 2)}" cy="${n2(yc)}" r="${n2(2.5 * pt)}" fill="${esc(e.color)}"/>`); return; }
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
    const apilado = (g) => g.grupo === "stacked" || g.grupo === "percentStacked";
    // extremos de los datos (con las sumas si van apiladas)
    let dmin = Infinity, dmax = -Infinity;
    const pct = m.tipos.some((g) => g.grupo === "percentStacked");
    m.tipos.forEach((g) => {
      const ss = g.series.map(datos);
      if (apilado(g)) {
        for (let i = 0; i < cats.length; i++) {
          let pos = 0, neg = 0;
          ss.forEach((v) => { const x = v[i] || 0; if (x >= 0) pos += x; else neg += x; });
          if (g.grupo === "percentStacked") { const tot = pos - neg || 1; pos /= tot; neg /= tot; }
          dmax = Math.max(dmax, pos); dmin = Math.min(dmin, neg);
        }
      } else ss.forEach((v) => v.forEach((x) => { if (x != null) { dmax = Math.max(dmax, x); dmin = Math.min(dmin, x); } }));
    });
    // cuántas divisiones caben: Office deja al menos unas 2.2 alturas de letra entre rótulos (en vertical)
    const lv0 = L(ejeV.letra, { tam: 9 }), lc0 = L(ejeC.letra, { tam: 9 });
    const largoEje = horizontal ? Math.max(40, pr - pl - (Math.max(0, ...cats.map((c) => medir(c, lc0.tam * pt, fam(lc0.f)))) + 6 * pt)) : Math.max(40, pb - ptop - lc0.tam * pt * 1.6);
    const maxDiv = Math.floor(largoEje / (lv0.tam * pt * (horizontal ? 3.6 : 2.2)));
    const esc_ = escala(dmin, dmax, ejeV, pct, maxDiv);
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
    const fmtV = (v) => (pct && !(ejeV.fmt && /%/.test(ejeV.fmt) && !ejeV.ligado) ? Math.round(v * 100) + "%" : numero(v, ejeV.ligado ? (series[0] && series[0].s.vals.fmt) || ejeV.fmt : ejeV.fmt, opc));
    const verV = !ejeV.borrado && ejeV.rotulos !== "none", verC = !ejeC.borrado && ejeC.rotulos !== "none";
    if (!disp || !disp.dentro) {
      if (horizontal) {
        if (verC) pl += Math.max(...cats.map((c) => medir(c, lc.tam * pt, fam(lc.f)))) + 6 * pt;
        if (verV) pb -= lv.tam * pt * 1.6;
      } else {
        if (verV) pl += Math.max(...marcasV.map((v) => medir(fmtV(v), lv.tam * pt, fam(lv.f)))) + 6 * pt;
        if (verC) pb -= lc.tam * pt * 1.6;
      }
      if (ejeV.titulo) { const lt = L(ejeV.titulo.letra, { tam: 10 }); if (horizontal) pb -= lt.tam * pt * 1.6; else pl += lt.tam * pt * 1.6; }
      if (ejeC.titulo) { const lt = L(ejeC.titulo.letra, { tam: 10 }); if (horizontal) pl += lt.tam * pt * 1.6; else pb -= lt.tam * pt * 1.6; }
    }
    if (m.trazado.relleno && m.trazado.relleno !== "none") partes.push(`<rect x="${n2(pl)}" y="${n2(ptop)}" width="${n2(Math.max(0, pr - pl))}" height="${n2(Math.max(0, pb - ptop))}" fill="${esc(m.trazado.relleno)}"/>`);
    // posición de un valor y de una categoría
    const largoV = horizontal ? pr - pl : pb - ptop;
    const aV = (v) => { const k = (v - esc_.a) / ((esc_.b - esc_.a) || 1); return horizontal ? (ejeV.inv ? pr - k * largoV : pl + k * largoV) : ejeV.inv ? ptop + k * largoV : pb - k * largoV; };
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
    // ---- series
    m.tipos.forEach((g) => {
      const ss = g.series.map((s) => ({ s, v: datos(s) })), k0 = series.findIndex((x) => x.g === g);
      if (g.tipo === "barra") {
        const N = apilado(g) ? 1 : ss.length, sol = apilado(g) ? 100 : g.solape, hueco = g.hueco;
        const bw = anchoCat / (N - ((N - 1) * sol) / 100 + hueco / 100);
        const base = aV(Math.max(esc_.a, Math.min(esc_.b, 0)));
        const pos = new Array(nC).fill(0), neg = new Array(nC).fill(0), tot = cats.map((_, i) => ss.reduce((a, x) => a + Math.abs(x.v[i] || 0), 0) || 1);
        ss.forEach(({ s, v }, j) => {
          const col = s.relleno && s.relleno !== "none" ? s.relleno : colorSerie(k0 + j), ln = trazo(s.linea, { no: 1 });
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
            partes.push(`<rect x="${n2(r.x)}" y="${n2(r.y)}" width="${n2(r.w)}" height="${n2(r.h)}" fill="${esc(fill)}"${ln ? attrsTrazo(ln) : ""}/>`);
            const ro = s.rotulos || g.rotulos;
            if (ro && ro.val) { const lr = L(ro.letra, { tam: 9 }); const txt = numero(x, ro.fmt || s.vals.fmt, opc); if (horizontal) t(r.x + r.w + 3 * pt, r.y + r.h / 2 + lr.tam * pt * 0.35, txt, lr, "start"); else t(r.x + r.w / 2, (ro.pos === "ctr" ? r.y + r.h / 2 : r.y - 3 * pt) + (ro.pos === "ctr" ? lr.tam * pt * 0.35 : 0), txt, lr, "middle"); }
          });
        });
      } else if (g.tipo === "linea" || g.tipo === "area" || g.tipo === "radar") {
        const acum = new Array(nC).fill(0), tot = cats.map((_, i) => ss.reduce((a, x) => a + Math.abs(x.v[i] || 0), 0) || 1);
        const capas = [];
        ss.forEach(({ s, v }, j) => {
          const col = (s.linea && s.linea.c) || (s.relleno && s.relleno !== "none" ? s.relleno : colorSerie(k0 + j));
          const ys = v.map((x, i) => { if (x == null) return null; const y = g.grupo === "percentStacked" ? x / tot[i] : x; if (apilado(g)) { acum[i] += y; return acum[i]; } return y; });
          const pts = ys.map((y, i) => (y == null ? null : [aC(i), aV(y)]));
          capas.push({ s, col, pts, j });
        });
        if (g.tipo === "area") {
          capas.slice().reverse().forEach(({ s, col, pts }, jj) => {
            const ok = pts.filter(Boolean); if (!ok.length) return;
            const base = aV(Math.max(esc_.a, 0));
            const d = `M${n2(ok[0][0])},${n2(base)} ` + ok.map((p) => `L${n2(p[0])},${n2(p[1])}`).join(" ") + ` L${n2(ok[ok.length - 1][0])},${n2(base)} Z`;
            const fill = s.relleno && s.relleno !== "none" ? s.relleno : col;
            partes.push(`<path d="${d}" fill="${esc(fill)}"${trazo(s.linea, { no: 1 }) ? attrsTrazo(trazo(s.linea, { no: 1 })) : ""}/>`);
          });
        } else {
          capas.forEach(({ s, col, pts, j }) => {
            const ln = trazo(s.linea, { c: col, w: 2.25 });
            let d = "", abierto = false;
            pts.forEach((p, i) => { if (!p) { abierto = false; return; } d += `${abierto ? "L" : "M"}${n2(p[0])},${n2(p[1])} `; abierto = true; });
            if (ln && d) partes.push(`<path d="${d.trim()}" fill="none"${attrsTrazo(ln)} stroke-linejoin="round" stroke-linecap="round"/>`);
            const mk = s.marcador, sim = mk ? mk.simbolo : g.marcadores ? "auto" : "none";
            if (sim && sim !== "none") pts.forEach((p) => { if (!p) return; const r = ((mk && mk.tam) || 5) * pt * 0.5; const f = (mk && mk.relleno && mk.relleno !== "none") ? mk.relleno : col; partes.push(sim === "square" ? `<rect x="${n2(p[0] - r)}" y="${n2(p[1] - r)}" width="${n2(2 * r)}" height="${n2(2 * r)}" fill="${esc(f)}"/>` : sim === "diamond" ? `<path d="M${n2(p[0])},${n2(p[1] - r * 1.3)} L${n2(p[0] + r * 1.3)},${n2(p[1])} L${n2(p[0])},${n2(p[1] + r * 1.3)} L${n2(p[0] - r * 1.3)},${n2(p[1])} Z" fill="${esc(f)}"/>` : sim === "triangle" ? `<path d="M${n2(p[0])},${n2(p[1] - r * 1.2)} L${n2(p[0] + r * 1.2)},${n2(p[1] + r)} L${n2(p[0] - r * 1.2)},${n2(p[1] + r)} Z" fill="${esc(f)}"/>` : `<circle cx="${n2(p[0])}" cy="${n2(p[1])}" r="${n2(r)}" fill="${esc(f)}"/>`); });
            const ro = s.rotulos || g.rotulos;
            if (ro && ro.val) { const lr = L(ro.letra, { tam: 9 }); pts.forEach((p, i) => { if (p) t(p[0], p[1] - 6 * pt, numero(s.vals.v[i], ro.fmt || s.vals.fmt, opc), lr, "middle"); }); }
          });
        }
      } else if (g.tipo === "dispersion") {
        ss.forEach(({ s, v }, j) => {
          const xs = (opc.valores && s.cats.f ? opc.valores(s.cats.f) || s.cats.v : s.cats.v).map((x) => (x == null || !isFinite(+x) ? null : +x));
          const col = (s.marcador && s.marcador.relleno && s.marcador.relleno !== "none" ? s.marcador.relleno : null) || (s.linea && s.linea.c) || colorSerie(k0 + j);
          const pts = v.map((y, i) => (y == null || xs[i] == null ? null : [aX(xs[i]), aV(y)]));
          const conLinea = /line|smooth/i.test(g.estiloDisp) && !(s.linea && s.linea.no);
          if (conLinea) { const ok = pts.filter(Boolean); if (ok.length > 1) partes.push(`<path d="${ok.map((p, k) => `${k ? "L" : "M"}${n2(p[0])},${n2(p[1])}`).join(" ")}" fill="none"${attrsTrazo(trazo(s.linea, { c: col, w: 2.25 }))}/>`); }
          const sim = s.marcador ? s.marcador.simbolo : "auto";
          if (sim !== "none") pts.forEach((p) => { if (p) partes.push(`<circle cx="${n2(p[0])}" cy="${n2(p[1])}" r="${n2(((s.marcador && s.marcador.tam) || 5) * pt * 0.5)}" fill="${esc(col)}"/>`); });
        });
      }
    });
    // eje de categorías: su línea y sus rótulos
    const lineaC = trazo(ejeC.linea, { c: "#D9D9D9", w: 0.75 });
    const yCero = aV(Math.max(esc_.a, Math.min(esc_.b, 0)));
    if (lineaC && !ejeC.borrado && !escX) partes.push(horizontal ? `<line x1="${n2(yCero)}" y1="${n2(ptop)}" x2="${n2(yCero)}" y2="${n2(pb)}"${attrsTrazo(lineaC)}/>` : `<line x1="${n2(pl)}" y1="${n2(yCero)}" x2="${n2(pr)}" y2="${n2(yCero)}"${attrsTrazo(lineaC)}/>`);
    const lineaV = trazo(ejeV.linea, { no: 1 });
    if (lineaV && !ejeV.borrado) partes.push(horizontal ? `<line x1="${n2(pl)}" y1="${n2(pb)}" x2="${n2(pr)}" y2="${n2(pb)}"${attrsTrazo(lineaV)}/>` : `<line x1="${n2(pl)}" y1="${n2(ptop)}" x2="${n2(pl)}" y2="${n2(pb)}"${attrsTrazo(lineaV)}/>`);
    if (verC && !escX) {
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
    if (ejeC.titulo && ejeC.titulo.texto) { const lt = L(ejeC.titulo.letra, { tam: 10 }); if (horizontal) partes.push(`<text x="${n2(7 * pt + lt.tam * pt)}" y="${n2((ptop + pb) / 2)}" font-family="${esc(fam(lt.f))}" font-size="${n2(lt.tam * pt)}" fill="${esc(lt.c)}" font-weight="${lt.b ? 700 : 400}" font-style="normal" text-anchor="middle" transform="rotate(-90 ${n2(7 * pt + lt.tam * pt)} ${n2((ptop + pb) / 2)})">${esc(ejeC.titulo.texto)}</text>`); else t((pl + pr) / 2, pb + lc.tam * pt * 1.6 + lt.tam * pt * 1.1, ejeC.titulo.texto, lt, "middle"); }
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${n2(w)}" height="${n2(h)}" viewBox="0 0 ${n2(w)} ${n2(h)}">${partes.join("")}</svg>`;

    // ---- pastel y anillo: desde las 12 en punto, en el sentido del reloj; borde blanco entre rebanadas
    function dibujarPastel(g, se, x1, y1, x2, y2) {
      const v = datos(se).map((x) => Math.max(0, x || 0)), tot = v.reduce((a, b) => a + b, 0) || 1;
      const cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, r = Math.max(1, Math.min(x2 - x1, y2 - y1) / 2), ri = g.tipo === "anillo" ? (r * Math.max(10, Math.min(90, g.agujero))) / 100 : 0;
      let ang = ((g.giro || 0) - 90) * (Math.PI / 180);
      const lnS = trazo(se.linea, { c: "#FFFFFF", w: 0.75 });
      v.forEach((x, i) => {
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

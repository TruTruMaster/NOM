// Pruebas de fidelidad con PowerPoint 2024 sobre tf_ppt.pptx y casos unitarios (P1–P3).
// Uso: node 4-CODIGO/app/src/test/js/office-ppt.test.js
"use strict";
const fs = require("fs"), path = require("path"), vm = require("vm"), zlib = require("zlib");

const dirJs = process.env.NOM_JS_DIR || path.join(__dirname, "../../main/resources/static/js");
const dirRef = process.env.NOM_REF_DIR || path.join(__dirname, "../../../../../trabajo/expo27/dueno-office");
const leerJs = (nombre) => fs.readFileSync(path.join(dirJs, nombre), "utf8");

let fallas = 0;
const prueba = (nombre, ok, detalle) => {
  console.log((ok ? "OK    " : "FALLA ") + nombre + (!ok && detalle ? " — " + detalle : ""));
  if (!ok) fallas++;
};
const saltar = (archivo) => console.log(`AVISO sin el archivo ${archivo}: se salta`);

function abrirZipBuf(buf) {
  const mapa = Object.create(null);
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("ZIP sin EOCD");
  const total = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  for (let k = 0; k < total; k++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break;
    const metodo = buf.readUInt16LE(p + 10);
    const tamComp = buf.readUInt32LE(p + 20);
    const nLen = buf.readUInt16LE(p + 28), eLen = buf.readUInt16LE(p + 30), cLen = buf.readUInt16LE(p + 32);
    const offLocal = buf.readUInt32LE(p + 42);
    const nombre = buf.slice(p + 46, p + 46 + nLen).toString("utf8");
    const lnLen = buf.readUInt16LE(offLocal + 26), leLen = buf.readUInt16LE(offLocal + 28);
    const ini = offLocal + 30 + lnLen + leLen;
    const comp = buf.slice(ini, ini + tamComp);
    mapa[nombre] = () => (metodo === 8 ? zlib.inflateRawSync(comp) : comp);
    p += 46 + nLen + eLen + cLen;
  }
  return {
    texto: async (r) => (mapa[r] ? mapa[r]().toString("utf8") : null),
    textoSync: (r) => (mapa[r] ? mapa[r]().toString("utf8") : null),
    leer: async (r) => (mapa[r] ? new Uint8Array(mapa[r]()) : null),
  };
}

function crearDOMParserMinimo() {
  const desEsc = (s) => String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
  function hacerEl(tag, attrs, nsHeredado, padre) {
    const nsMapa = Object.assign(Object.create(null), nsHeredado);
    for (const [k, v] of attrs) {
      if (k === "xmlns") nsMapa[""] = v;
      else if (k.startsWith("xmlns:")) nsMapa[k.slice(6)] = v;
    }
    const i = tag.indexOf(":");
    const pref = i > 0 ? tag.slice(0, i) : "";
    const localName = i > 0 ? tag.slice(i + 1) : tag;
    const ns = nsMapa[pref] || "";
    const attrLista = attrs.map(([name, value]) => {
      const j = name.indexOf(":");
      const ap = j > 0 ? name.slice(0, j) : "";
      const al = j > 0 ? name.slice(j + 1) : name;
      return { name, value, localName: al, ns: ap ? (nsMapa[ap] || "") : "" };
    });
    const el = {
      nodeType: 1, tagName: tag, localName, namespaceURI: ns, parentNode: padre || null,
      attributes: attrLista, childNodes: [], children: [], _nsMapa: nsMapa,
      get textContent() { return this.childNodes.map((c) => (c.nodeType === 3 ? c.data : c.textContent)).join(""); },
      getAttribute(k) { const a = attrLista.find((x) => x.name === k); return a ? a.value : null; },
      getAttributeNS(u, l) { const a = attrLista.find((x) => x.localName === l && (u === "*" || x.ns === u)); return a ? a.value : null; },
      hasAttribute(k) { return attrLista.some((x) => x.name === k); },
      getElementsByTagName(nm) {
        const out = [];
        const rec = (n) => { for (const c of n.children) { if (nm === "*" || c.tagName === nm || c.localName === nm) out.push(c); rec(c); } };
        rec(this); return out;
      },
      getElementsByTagNameNS(u, l) {
        const out = [];
        const rec = (n) => { for (const c of n.children) { if ((l === "*" || c.localName === l) && (u === "*" || c.namespaceURI === u)) out.push(c); rec(c); } };
        rec(this); return out;
      },
    };
    return el;
  }
  function parsear(xml) {
    const raiz = hacerEl("#doc", [], Object.create(null), null);
    const pila = [raiz];
    const re = /<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:\s+[^\s=/>]+(?:\s*=\s*"[^"]*"|\s*=\s*'[^']*'))*)\s*(\/?)>|([^<]+)/g;
    let m;
    while ((m = re.exec(String(xml || "")))) {
      const tope = pila[pila.length - 1];
      if (m[1]) { if (pila.length > 1) pila.pop(); }
      else if (m[2]) {
        const attrs = [];
        const reA = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
        let a;
        while ((a = reA.exec(m[3] || ""))) attrs.push([a[1], desEsc(a[2] != null ? a[2] : a[3])]);
        const el = hacerEl(m[2], attrs, tope._nsMapa, tope);
        tope.childNodes.push(el); tope.children.push(el);
        if (!m[4]) pila.push(el);
      } else if (m[5] != null) {
        tope.childNodes.push({ nodeType: 3, data: desEsc(m[5]), parentNode: tope });
      }
    }
    return { documentElement: raiz.children[0] || raiz, getElementsByTagName: (n) => raiz.getElementsByTagName(n), getElementsByTagNameNS: (u, l) => raiz.getElementsByTagNameNS(u, l) };
  }
  return class DOMParser { parseFromString(s) { return parsear(s); } };
}

function crearEntorno() {
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const NOM = {
    esc, ico: () => "", $: () => null, $$: () => [], aviso: () => {}, colorCss: (v, d = "") => v || d,
    modos: {}, leer: (k, d) => d, guardar: () => {}, api: { usuario: () => ({ id: "u1" }) },
    srcBin: (v) => v,
    diapo: {
      TAMANOS: { "16:9": { w: 960, h: 540 }, "4:3": { w: 960, h: 720 }, carta: { w: 960, h: 742 }, cuadrado: { w: 720, h: 720 } },
      tam: () => ({ w: 960, h: 540 }),
      tema: () => ({ fondo: "#FFFFFF", tinta: "#1F1E26" }),
      htmlRico: (h) => String(h || ""),
      tablaHtml: () => "<table></table>",
      graficaHtml: (b) => (b.office && NOM.graficaOffice ? NOM.graficaOffice.svg(b.office, b.w || 300, b.h || 200, { pt: b.ptOffice || 1 }) : ""),
      fondoCss: () => "",
    },
  };
  const document = {
    createElement: (tag) => {
      if (tag === "canvas") return { getContext: () => ({ font: "", measureText: (t) => ({ width: String(t).length * 6 }) }) };
      return { style: {}, dataset: {}, content: { firstChild: null } };
    },
    addEventListener() {}, removeEventListener() {}, querySelector: () => null,
  };
  const window = { NOM, document, matchMedia: () => ({ matches: false }), addEventListener() {}, removeEventListener() {} };
  const sandbox = {
    window, NOM, document, matchMedia: window.matchMedia, console, TextEncoder, TextDecoder,
    atob: (s) => Buffer.from(s, "base64").toString("binary"),
    btoa: (s) => Buffer.from(s, "binary").toString("base64"),
    setTimeout, clearTimeout, DOMParser: crearDOMParserMinimo(),
    Image: class Image { set src(_) { setTimeout(() => this.onerror && this.onerror(), 0); } },
  };
  const ctx = vm.createContext(sandbox);
  const cargar = (archivo) => vm.runInContext(leerJs(archivo), ctx, { filename: archivo });
  return { NOM, cargar };
}

(async () => {
  const env = crearEntorno();
  env.cargar("crdt.js");
  env.cargar("compresion.js");
  env.cargar("origen.js");
  env.cargar("formulas.js");
  env.cargar("geometrias.js");
  env.cargar("graficas-office.js");
  env.cargar("lienzo.js");
  env.cargar("ooxml.js");

  const rutaPpt = path.join(dirRef, "tf_ppt.pptx");
  const ctxGraf = { color: (n) => { const s = n.getElementsByTagNameNS("*", "srgbClr")[0]; return s ? "#" + s.getAttribute("val") : null; }, acentos: ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"], fuente: () => "Calibri" };

  if (!fs.existsSync(rutaPpt)) {
    saltar("tf_ppt.pptx");
  } else {
    const bufPpt = fs.readFileSync(rutaPpt);
    const zipPpt = abrirZipBuf(bufPpt);
    const ppt = await env.NOM.ooxml.leerPptx({ size: bufPpt.length, arrayBuffer: async () => bufPpt.buffer.slice(bufPpt.byteOffset, bufPpt.byteOffset + bufPpt.byteLength) });

    // Armamos un Doc con las diapositivas importadas para pasar por lienzoVista.htmlDiapositiva
    const doc = new env.NOM.DocCRDT("ppt-test", null, { local: true, soloMemoria: true });
    doc.cambiar("_doc", Object.assign({ modo: "presentacion", paginas: ppt.paginas }, ppt.doc || {}));
    ppt.bloques.forEach((b, idx) => doc.cambiar(b.id || "b" + idx, Object.assign({ _vivo: true }, b)));

    // ---- P1: Diapositiva 9 líneas horizontales (h=0) sin sinPos/diagonal/bloque libre, con guiones y puntas
    const pg3 = ppt.paginas[2], pg9 = ppt.paginas[8];
    const bloquesS9 = ppt.bloques.filter((b) => b.pagina === pg9);
    const lineasS9 = bloquesS9.filter((b) => b.lineaOffice);
    const htmlS9 = env.NOM.lienzoVista.htmlDiapositiva(doc, pg9);
    prueba(
      "P1a. Diapositiva 9 importa las 3 líneas horizontales con h=0, puntas y bordeGuion (sysDot, lgDashDot)",
      lineasS9.length === 3
        && lineasS9.every((l) => l.h === 0 && l.w > 100)
        && lineasS9[0].puntas && lineasS9[0].puntas.ini && lineasS9[0].puntas.ini.t === "oval"
        && lineasS9[0].puntas.fin && lineasS9[0].puntas.fin.t === "triangle"
        && lineasS9[1].bordeGuion === "sysDot"
        && lineasS9[2].bordeGuion === "lgDashDot",
      `lineasS9=${JSON.stringify(lineasS9.map((l) => ({ x: l.x, y: l.y, w: l.w, h: l.h, g: l.bordeGuion, p: l.puntas })))}`
    );
    // Verifica que en el HTML de la diapositiva 9 ninguna línea caiga en sin-pos (bloque libre) ni con h=160 ni diagonal
    const bloquesHtmlLinea = htmlS9.match(/<div class="bl[^"]*linea-office[^"]*"[^>]*>[\s\S]*?<\/svg><\/div><\/div>/g) || [];
    const sinPosEnLineas = bloquesHtmlLinea.some((s) => s.includes("sin-pos") || s.includes("height:160px") || s.includes("width:840px"));
    const pathLineas = bloquesHtmlLinea.map((s) => (s.match(/<path d="([^"]+)"/) || [])[1] || "");
    const sonHorizontales = pathLineas.length === 3 && pathLineas.every((d) => /^M0,0L\d+(?:\.\d+)?,0$/.test(d));
    const tieneDashDot = bloquesHtmlLinea[2] && /stroke-dasharray="\d+(?:\.\d+)? \d+(?:\.\d+)? \d+(?:\.\d+)? \d+(?:\.\d+)?"/.test(bloquesHtmlLinea[2]);
    prueba(
      "P1b. Diapositiva 9 dibuja las 3 líneas horizontales en su sitio (sin 'sin-pos'/bloque libre, sin diagonal, con lgDashDot de 4 trazos)",
      bloquesHtmlLinea.length === 3 && !sinPosEnLineas && sonHorizontales && tieneDashDot,
      `n=${bloquesHtmlLinea.length} sinPos=${sinPosEnLineas} paths=${JSON.stringify(pathLineas)} dashDot=${!!tieneDashDot}`
    );

    // ---- P2: Contornos punteados (prstDash) y sombra exterior (outerShdw) en cuadros de texto (diap 3) y figuras (diap 9)
    const bloquesS3 = ppt.bloques.filter((b) => b.pagina === pg3);
    const cuadroS3 = bloquesS3.find((b) => b.tipo === "texto" && /Text box: fill, border, shadow/i.test(b.html || ""));
    const rectAzulS9 = bloquesS9.find((b) => b.tipo === "forma" && !b.lineaOffice && (b.bordeGuion === "dash" || b.borde === "#1F4E79" || /Dashed|Shadow/i.test(b.texto || "") || b.color === "#4472C4"));
    const htmlS3 = env.NOM.lienzoVista.htmlDiapositiva(doc, pg3);
    prueba(
      "P2a. Diapositiva 3 cuadro de texto lee prstDash='dash' y outerShdw y los pinta como outline dashed + drop-shadow",
      !!cuadroS3 && cuadroS3.ppt && Array.isArray(cuadroS3.ppt.borde) && cuadroS3.ppt.borde[2] === "dash"
        && !!cuadroS3.sombraOffice && cuadroS3.sombraOffice.b > 0
        && /outline:[^;]*dashed/.test(htmlS3) && /drop-shadow\(/.test(htmlS3),
      `borde=${JSON.stringify(cuadroS3 && cuadroS3.ppt && cuadroS3.ppt.borde)} sombra=${JSON.stringify(cuadroS3 && cuadroS3.sombraOffice)}`
    );
    prueba(
      "P2b. Diapositiva 9 rectángulo azul lee prstDash='dash' y outerShdw y los pinta con stroke-dasharray + drop-shadow",
      !!rectAzulS9 && rectAzulS9.bordeGuion === "dash" && !!rectAzulS9.sombraOffice && rectAzulS9.sombraOffice.b > 0
        && /stroke-dasharray=/.test(env.NOM.figuras.svgFigura(rectAzulS9))
        && /drop-shadow\(/.test(htmlS9),
      `rectAzul=${JSON.stringify(rectAzulS9 && { borde: rectAzulS9.borde, g: rectAzulS9.bordeGuion, s: rectAzulS9.sombraOffice })}`
    );

    // ---- P3: Gráficas de diapositiva 6 (pastel/dona) y diapositiva 7 (radar, burbujas, área, barra apilada)
    // P3a: Radar (chart7.xml)
    const mRadar = env.NOM.graficaOffice.leer(zipPpt.textoSync("ppt/charts/chart7.xml"), ctxGraf);
    const svgRadar = env.NOM.graficaOffice.svg(mRadar, 400, 260, { pt: 1 });
    const poligonosRadar = (svgRadar.match(/<polygon\b/g) || []).length;
    const ejesRadiales = (svgRadar.match(/<line\b/g) || []).length;
    prueba(
      "P3a. Gráfica radar (diap 7, chart7) se dibuja como radar con anillos poligonales y ejes radiales (no líneas cartesianas)",
      mRadar.tipos[0].tipo === "radar" && poligonosRadar >= 5 && ejesRadiales >= 5,
      `tipo=${mRadar.tipos[0].tipo} polygons=${poligonosRadar} lines=${ejesRadiales}`
    );

    // P3b: Burbujas (chart8.xml tiene ptCount=5 con tamaños 5, 12, 8, 20, 10)
    const mBurb = env.NOM.graficaOffice.leer(zipPpt.textoSync("ppt/charts/chart8.xml"), ctxGraf);
    const svgBurb = env.NOM.graficaOffice.svg(mBurb, 400, 260, { pt: 1 });
    const radiosBurb = [...svgBurb.matchAll(/<circle\b[^>]*\br="([\d.]+)"/g)].map((m) => +m[1]);
    const tieneLineaBurb = /<path\b/.test(svgBurb);
    prueba(
      "P3b. Gráfica de burbujas (diap 7, chart8) dibuja burbujas de radio proporcional a bubbleSize y sin línea conectora",
      mBurb.tipos[0].tipo === "burbuja" && !tieneLineaBurb && radiosBurb.length === 5
        && radiosBurb[3] > radiosBurb[1] && radiosBurb[1] > radiosBurb[4] && radiosBurb[4] > radiosBurb[2] && radiosBurb[2] > radiosBurb[0],
      `tipo=${mBurb.tipos[0].tipo} linea=${tieneLineaBurb} radios=${JSON.stringify(radiosBurb)}`
    );

    // P3c: Área (diap 7, chart6) muestra ambas series (Cost encima de Revenue) y área apilada dibuja bandas
    const mArea = env.NOM.graficaOffice.leer(zipPpt.textoSync("ppt/charts/chart6.xml"), ctxGraf);
    const svgArea = env.NOM.graficaOffice.svg(mArea, 400, 260, { pt: 1 });
    const pathsArea = [...svgArea.matchAll(/<path\b[^>]*data-s="(\d+)"[^>]*d="([^"]+)"/g)];
    const ordenArea = pathsArea.map((m) => +m[1]);
    // También probamos área apilada (grouping="stacked"): la serie 1 NO debe bajar hasta la base 0 tapando a la serie 0
    const xmlAreaApil = zipPpt.textoSync("ppt/charts/chart6.xml").replace("<c:areaChart>", '<c:areaChart><c:grouping val="stacked"/>');
    const mAreaApil = env.NOM.graficaOffice.leer(xmlAreaApil, ctxGraf);
    const svgAreaApil = env.NOM.graficaOffice.svg(mAreaApil, 400, 260, { pt: 1 });
    const pApil = [...svgAreaApil.matchAll(/<path\b[^>]*data-s="(\d+)"[^>]*d="([^"]+)"/g)];
    // En una banda apilada la serie 1 tiene 2*N puntos (curva superior + curva inferior de vuelta)
    const ptsSerie1Apil = pApil[1] ? (pApil[1][2].match(/[ML]\d/g) || []).length : 0;
    prueba(
      "P3c. Gráfica de área (diap 7, chart6) pinta serie 0 (Revenue) atrás y serie 1 (Cost) delante; área apilada pinta banda",
      ordenArea.length === 2 && ordenArea[0] === 0 && ordenArea[1] === 1 && ptsSerie1Apil >= 8,
      `ordenArea=${JSON.stringify(ordenArea)} ptsSerie1Apil=${ptsSerie1Apil}`
    );

    // P3d: Barra apilada horizontal (diap 7, chart9) conserva orden de leyenda Revenue → Cost (no al revés)
    const mBarAp = env.NOM.graficaOffice.leer(zipPpt.textoSync("ppt/charts/chart9.xml"), ctxGraf);
    const svgBarAp = env.NOM.graficaOffice.svg(mBarAp, 400, 260, { pt: 1 });
    const ixRev = svgBarAp.indexOf(">Revenue<"), ixCost = svgBarAp.indexOf(">Cost<");
    // Y en agrupada horizontal (clustered) sigue invirtiendo para coincidir de arriba abajo
    const xmlBarAg = zipPpt.textoSync("ppt/charts/chart9.xml").replace('val="stacked"', 'val="clustered"');
    const svgBarAg = env.NOM.graficaOffice.svg(env.NOM.graficaOffice.leer(xmlBarAg, ctxGraf), 400, 260, { pt: 1 });
    const ixRevAg = svgBarAg.indexOf(">Revenue<"), ixCostAg = svgBarAg.indexOf(">Cost<");
    prueba(
      "P3d. Barra apilada horizontal (diap 7, chart9) mantiene leyenda 'Revenue' antes que 'Cost'",
      ixRev > 0 && ixCost > 0 && ixRev < ixCost && ixCostAg < ixRevAg,
      `apilada: Rev=${ixRev} Cost=${ixCost}; agrupada: Rev=${ixRevAg} Cost=${ixCostAg}`
    );

    // P3e: Etiquetas del pastel (chart3) y la dona (chart4) en diap 6: solo porcentaje, letra 10 pt (dLbl por punto)
    const mPie = env.NOM.graficaOffice.leer(zipPpt.textoSync("ppt/charts/chart3.xml"), ctxGraf);
    const svgPie = env.NOM.graficaOffice.svg(mPie, 400, 260, { pt: 1 });
    const etqsPie = [...svgPie.matchAll(/<text\b[^>]*data-s="0"[^>]*font-size="([\d.]+)"[^>]*>([^<]+)<\/text>/g)].map((m) => ({ fs: +m[1], tx: m[2] }));
    const mDona = env.NOM.graficaOffice.leer(zipPpt.textoSync("ppt/charts/chart4.xml"), ctxGraf);
    const svgDona = env.NOM.graficaOffice.svg(mDona, 400, 260, { pt: 1 });
    const etqsDona = [...svgDona.matchAll(/<text\b[^>]*data-s="0"[^>]*font-size="([\d.]+)"[^>]*>([^<]+)<\/text>/g)].map((m) => ({ fs: +m[1], tx: m[2] }));
    prueba(
      "P3e. Pastel y dona (diap 6, chart3 y chart4) respetan dLbl por punto: solo porcentaje pequeño (10 pt), sin categoría",
      etqsPie.length === 4 && etqsPie.every((e) => /^\d+%$/.test(e.tx) && e.fs === 10)
        && etqsDona.length === 3 && etqsDona.every((e) => /^\d+%$/.test(e.tx) && e.fs === 10),
      `pie=${JSON.stringify(etqsPie)} dona=${JSON.stringify(etqsDona)}`
    );
  }

  // ============================================================ Casos sintéticos mínimos (sin archivos del dueño)
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
  const abrirZipReal = env.NOM.origen.abrirZip;
  async function leerPptxSintetico(slideXml) {
    const mapa = {
      "ppt/presentation.xml": `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
      "ppt/_rels/presentation.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`,
      "ppt/slides/slide1.xml": slideXml,
    };
    env.NOM.origen.abrirZip = () => ({ texto: async (r) => mapa[r] || null, leer: async () => null });
    try { return await env.NOM.ooxml.leerPptx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) }); }
    finally { env.NOM.origen.abrirZip = abrirZipReal; }
  }

  // P1 sintético: línea horizontal (cy=0) con puntas y lgDashDot + conector vertical (cx=0) con flipV y sysDot
  {
    const pptSin = await leerPptxSintetico(`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="L1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="952500" y="1905000"/><a:ext cx="1905000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>
        <a:ln w="38100"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="lgDashDot"/><a:headEnd type="oval"/><a:tailEnd type="triangle"/></a:ln></p:spPr></p:sp>
      <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="C1"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr><a:xfrm flipV="1"><a:off x="3810000" y="952500"/><a:ext cx="0" cy="1905000"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
        <a:ln w="25400"><a:solidFill><a:srgbClr val="CC0000"/></a:solidFill><a:prstDash val="sysDot"/><a:tailEnd type="arrow"/></a:ln></p:spPr></p:cxnSp>
    </p:spTree></p:cSld></p:sld>`);
    const docSin = new env.NOM.DocCRDT("ppt-p1s", null, { local: true, soloMemoria: true });
    docSin.cambiar("_doc", { modo: "presentacion", paginas: pptSin.paginas });
    pptSin.bloques.forEach((b, i) => docSin.cambiar(b.id || "b" + i, Object.assign({ _vivo: true }, b)));
    const htmlSin = env.NOM.lienzoVista.htmlDiapositiva(docSin, pptSin.paginas[0]);
    const ls = pptSin.bloques.filter((b) => b.lineaOffice);
    const divs = htmlSin.match(/<div class="bl[^"]*linea-office[^"]*"[^>]*>[\s\S]*?<\/svg><\/div><\/div>/g) || [];
    const p0 = (divs[0] && divs[0].match(/<path d="([^"]+)"/) || [])[1] || "";
    const p1 = (divs[1] && divs[1].match(/<path d="([^"]+)"/) || [])[1] || "";
    prueba(
      "P1s. Sintético: p:sp line (h=0, lgDashDot, puntas) y p:cxnSp vertical (w=0, flipV, sysDot) se dibujan rectos sin sin-pos",
      ls.length === 2 && ls[0].h === 0 && ls[0].bordeGuion === "lgDashDot" && ls[1].w === 0 && ls[1].volteoV === true && ls[1].bordeGuion === "sysDot"
        && divs.length === 2 && !divs.some((d) => d.includes("sin-pos") || d.includes("height:160px"))
        && /^M0,0L200,0$/.test(p0) && /^M0,0L0,200$/.test(p1)
        && /stroke-dasharray="\d+(?:\.\d+)? \d+(?:\.\d+)? \d+(?:\.\d+)? \d+(?:\.\d+)?"/.test(divs[0])
        && divs[0].includes("<ellipse ") && divs[1].includes("scale(1 -1)"),
      `ls=${JSON.stringify(ls.map((l) => ({ w: l.w, h: l.h, g: l.bordeGuion, fv: l.volteoV })))} p0=${p0} p1=${p1}`
    );
  }

  // P2 sintético: cuadro de texto con prstDash='sysDash' + outerShdw y figura con prstDash='dashDot' + outerShdw
  {
    const pptSin = await leerPptxSintetico(`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="T1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="2857500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="DDEBF7"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>
        <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Cuadro</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="3" name="F1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="952500" y="2381250"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:prstDash val="dashDot"/></a:ln>
        <a:effectLst><a:outerShdw blurRad="76200" dist="50800" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr></p:sp>
    </p:spTree></p:cSld></p:sld>`);
    const docSin = new env.NOM.DocCRDT("ppt-p2s", null, { local: true, soloMemoria: true });
    docSin.cambiar("_doc", { modo: "presentacion", paginas: pptSin.paginas });
    pptSin.bloques.forEach((b, i) => docSin.cambiar(b.id || "b" + i, Object.assign({ _vivo: true }, b)));
    const htmlSin = env.NOM.lienzoVista.htmlDiapositiva(docSin, pptSin.paginas[0]);
    const tb = pptSin.bloques.find((b) => b.tipo === "texto");
    const sh = pptSin.bloques.find((b) => b.tipo === "forma");
    const svgSh = sh ? env.NOM.figuras.svgFigura(sh) : "";
    prueba(
      "P2s. Sintético: prstDash (sysDash, dashDot) y outerShdw en cuadro de texto y figura generan outline dashed, stroke-dasharray y drop-shadow",
      !!tb && tb.ppt.borde[2] === "sysDash" && !!tb.sombraOffice && tb.sombraOffice.b > 0
        && !!sh && sh.bordeGuion === "dashDot" && !!sh.sombraOffice && sh.sombraOffice.b > 0
        && /outline:[^;]*dashed/.test(htmlSin) && /stroke-dasharray="\d+(?:\.\d+)? \d+(?:\.\d+)? \d+(?:\.\d+)? \d+(?:\.\d+)?"/.test(svgSh)
        && (htmlSin.match(/drop-shadow\(/g) || []).length === 2,
      `tb=${JSON.stringify(tb && { b: tb.ppt.borde, s: tb.sombraOffice })} sh=${JSON.stringify(sh && { g: sh.bordeGuion, s: sh.sombraOffice })}`
    );
  }

  // P3 sintéticos: radar, burbujas, área + área apilada, leyenda barra apilada/agrupada horizontal, dLbl pastel/dona
  {
    const xmlRad = `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:radarChart><c:radarStyle val="standard"/>
      <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>A</c:v></c:tx>
        <c:cat><c:strLit><c:ptCount val="5"/><c:pt idx="0"><c:v>E1</c:v></c:pt><c:pt idx="1"><c:v>E2</c:v></c:pt><c:pt idx="2"><c:v>E3</c:v></c:pt><c:pt idx="3"><c:v>E4</c:v></c:pt><c:pt idx="4"><c:v>E5</c:v></c:pt></c:strLit></c:cat>
        <c:val><c:numLit><c:ptCount val="5"/><c:pt idx="0"><c:v>40</c:v></c:pt><c:pt idx="1"><c:v>80</c:v></c:pt><c:pt idx="2"><c:v>60</c:v></c:pt><c:pt idx="3"><c:v>90</c:v></c:pt><c:pt idx="4"><c:v>50</c:v></c:pt></c:numLit></c:val></c:ser>
      <c:axId val="1"/><c:axId val="2"/></c:radarChart>
      <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;
    const mR = env.NOM.graficaOffice.leer(xmlRad, ctxGraf), sR = env.NOM.graficaOffice.svg(mR, 360, 240, { pt: 1 });
    prueba(
      "P3as. Sintético: c:radarChart dibuja rejilla poligonal (<polygon>) y ejes radiales (<line>)",
      mR.tipos[0].tipo === "radar" && (sR.match(/<polygon\b/g) || []).length >= 5 && (sR.match(/<line\b/g) || []).length >= 5,
      `tipo=${mR.tipos[0].tipo} poly=${(sR.match(/<polygon\b/g) || []).length} line=${(sR.match(/<line\b/g) || []).length}`
    );

    const xmlBurb = `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:bubbleChart>
      <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>B</c:v></c:tx>
        <c:xVal><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:xVal>
        <c:yVal><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:yVal>
        <c:bubbleSize><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>16</c:v></c:pt><c:pt idx="2"><c:v>9</c:v></c:pt></c:numLit></c:bubbleSize></c:ser>
      <c:axId val="1"/><c:axId val="2"/></c:bubbleChart>
      <c:valAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;
    const mB = env.NOM.graficaOffice.leer(xmlBurb, ctxGraf), sB = env.NOM.graficaOffice.svg(mB, 360, 240, { pt: 1 });
    const rs = [...sB.matchAll(/<circle\b[^>]*\br="([\d.]+)"/g)].map((m) => +m[1]);
    prueba(
      "P3bs. Sintético: c:bubbleChart dibuja burbujas proporcionales a bubbleSize sin <path>",
      mB.tipos[0].tipo === "burbuja" && !/<path\b/.test(sB) && rs.length === 3 && rs[1] > rs[2] && rs[2] > rs[0],
      `tipo=${mB.tipos[0].tipo} rs=${JSON.stringify(rs)}`
    );

    const xmlArea = (grp) => `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:areaChart>${grp ? `<c:grouping val="${grp}"/>` : ""}
      <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>S1</c:v></c:tx>
        <c:cat><c:strLit><c:ptCount val="4"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt><c:pt idx="3"><c:v>Q4</c:v></c:pt></c:strLit></c:cat>
        <c:val><c:numLit><c:ptCount val="4"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>15</c:v></c:pt><c:pt idx="3"><c:v>25</c:v></c:pt></c:numLit></c:val></c:ser>
      <c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>S2</c:v></c:tx>
        <c:val><c:numLit><c:ptCount val="4"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt><c:pt idx="3"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
      <c:axId val="1"/><c:axId val="2"/></c:areaChart>
      <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;
    const sA0 = env.NOM.graficaOffice.svg(env.NOM.graficaOffice.leer(xmlArea(""), ctxGraf), 360, 240, { pt: 1 });
    const sA1 = env.NOM.graficaOffice.svg(env.NOM.graficaOffice.leer(xmlArea("stacked"), ctxGraf), 360, 240, { pt: 1 });
    const ordA = [...sA0.matchAll(/<path\b[^>]*data-s="(\d+)"/g)].map((m) => +m[1]);
    const pA1 = [...sA1.matchAll(/<path\b[^>]*data-s="(\d+)"[^>]*d="([^"]+)"/g)];
    const nPtsA1 = pA1[1] ? (pA1[1][2].match(/[ML]\d/g) || []).length : 0;
    prueba(
      "P3cs. Sintético: área estándar pinta serie 0→1 y área apilada (stacked) dibuja banda de 2*N puntos",
      ordA.length === 2 && ordA[0] === 0 && ordA[1] === 1 && nPtsA1 === 8,
      `ordA=${JSON.stringify(ordA)} nPtsA1=${nPtsA1}`
    );

    const xmlBar = (grp) => `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="${grp}"/>
      <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Primero</c:v></c:tx><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
      <c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>Segundo</c:v></c:tx><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt></c:numLit></c:val></c:ser>
      <c:axId val="1"/><c:axId val="2"/></c:barChart>
      <c:catAx><c:axId val="1"/><c:axPos val="l"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="b"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>`;
    const sBarS = env.NOM.graficaOffice.svg(env.NOM.graficaOffice.leer(xmlBar("stacked"), ctxGraf), 360, 240, { pt: 1 });
    const sBarC = env.NOM.graficaOffice.svg(env.NOM.graficaOffice.leer(xmlBar("clustered"), ctxGraf), 360, 240, { pt: 1 });
    prueba(
      "P3ds. Sintético: barra apilada horizontal mantiene leyenda Primero→Segundo; agrupada la invierte",
      sBarS.indexOf(">Primero<") < sBarS.indexOf(">Segundo<") && sBarC.indexOf(">Segundo<") < sBarC.indexOf(">Primero<"),
      `stacked=${sBarS.indexOf(">Primero<")},${sBarS.indexOf(">Segundo<")} clustered=${sBarC.indexOf(">Primero<")},${sBarC.indexOf(">Segundo<")}`
    );

    const xmlPie = `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:pieChart><c:ser><c:idx val="0"/><c:order val="0"/>
      <c:dLbls>
        <c:dLbl><c:idx val="0"/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr><c:showVal val="0"/><c:showCatName val="0"/><c:showPercent val="1"/></c:dLbl>
        <c:dLbl><c:idx val="1"/><c:showVal val="1"/><c:showCatName val="0"/><c:showPercent val="0"/></c:dLbl>
        <c:showVal val="0"/><c:showCatName val="1"/><c:showPercent val="1"/>
      </c:dLbls>
      <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>Norte</c:v></c:pt><c:pt idx="1"><c:v>Sur</c:v></c:pt></c:strLit></c:cat>
      <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>75</c:v></c:pt><c:pt idx="1"><c:v>25</c:v></c:pt></c:numLit></c:val>
    </c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>`;
    const sP = env.NOM.graficaOffice.svg(env.NOM.graficaOffice.leer(xmlPie, ctxGraf), 360, 240, { pt: 1 });
    const eP = [...sP.matchAll(/<text\b[^>]*data-s="0"[^>]*font-size="([\d.]+)"[^>]*>([^<]+)<\/text>/g)].map((m) => ({ fs: +m[1], tx: m[2] }));
    prueba(
      "P3es. Sintético: pastel respeta c:dLbl por punto (showPercent/showVal/showCatName y tamaño txPr o 10 pt por omisión)",
      eP.length === 2 && eP[0].tx === "75%" && eP[0].fs === 10 && eP[1].tx === "25" && eP[1].fs === 10,
      `eP=${JSON.stringify(eP)}`
    );
  }

  console.log(fallas ? `${fallas} FALLA(S)` : "Todas las pruebas pasaron");
  process.exit(fallas ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });

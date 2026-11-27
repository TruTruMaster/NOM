// Pruebas de fidelidad con Excel 2024 sobre tf_excel.xlsx y excel_ooxml_stress_test_max.xlsx (X1–X2).
// Uso: node 4-CODIGO/app/src/test/js/office-excel.test.js
"use strict";
const fs = require("fs"), path = require("path"), vm = require("vm"), zlib = require("zlib");

const dirJs = process.env.NOM_JS_DIR || path.join(__dirname, "../../main/resources/static/js");
const dirRef = process.env.NOM_REF_DIR || path.join(__dirname, "../../../../../trabajo/expo27/dueno-office");
const dirRaiz = path.join(__dirname, "../../../../..");
const leerJs = (nombre) => fs.readFileSync(path.join(dirJs, nombre), "utf8");

let fallas = 0;
const prueba = (nombre, ok, detalle) => {
  console.log((ok ? "OK    " : "FALLA ") + nombre + (!ok && detalle ? " — " + detalle : ""));
  if (!ok) fallas++;
};
const saltados = new Set();
const saltar = (archivo) => { if (!saltados.has(archivo)) { saltados.add(archivo); console.log(`AVISO sin el archivo ${archivo}: se salta`); } };

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
    nombres: () => Object.keys(mapa),
    texto: (r) => (mapa[r] ? mapa[r]().toString("utf8") : null),
    leer: (r) => (mapa[r] ? new Uint8Array(mapa[r]()) : null),
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
    modos: {}, leer: (k, d) => d, guardar: () => {}, api: { usuario: () => ({ id: "u1" }) }, srcBin: (v) => v,
  };
  const document = {
    createElement: (tag) => (tag === "canvas" ? { getContext: () => ({ font: "", measureText: (t) => ({ width: String(t).length * 6 }) }) } : { style: {}, dataset: {} }),
    addEventListener() {}, removeEventListener() {}, querySelector: () => null,
  };
  const window = { NOM, document, matchMedia: () => ({ matches: false }), addEventListener() {}, removeEventListener() {} };
  const sandbox = {
    window, NOM, document, matchMedia: window.matchMedia, console, TextEncoder, TextDecoder,
    atob: (s) => Buffer.from(s, "base64").toString("binary"),
    btoa: (s) => Buffer.from(s, "binary").toString("base64"),
    setTimeout, clearTimeout, DOMParser: crearDOMParserMinimo(),
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
  env.cargar("graficas.js");
  env.cargar("graficas-office.js");
  env.cargar("ooxml.js");

  const F = env.NOM.formulas;

  // ---- X1: tf_excel.xlsx nombres definidos con ámbito de hoja y Hoja!Nombre
  const rutaTfXl = path.join(dirRef, "tf_excel.xlsx");
  if (!fs.existsSync(rutaTfXl)) {
    saltar("tf_excel.xlsx");
  } else {
    const bufXl = fs.readFileSync(rutaTfXl);
    const xl = await env.NOM.ooxml.leerXlsx({ size: bufXl.length, arrayBuffer: async () => bufXl.buffer.slice(bufXl.byteOffset, bufXl.byteOffset + bufXl.byteLength) });

    const doc = new env.NOM.DocCRDT("xl-test", null, { local: true, soloMemoria: true });
    doc.cambiar("_doc", Object.assign({ modo: "hoja" }, xl.doc || {}));
    xl.hojas.forEach((h, i) => doc.cambiar("h" + i, Object.assign({ _vivo: true, tipo: "tabla", nombre: h.nombre, orden: i }, h.campos)));

    const hojaSummary = F.hojasDe(doc).find((h) => F.nombreDeHoja(doc, h) === "Summary");
    const evSummary = F.evaluadorDeTabla(doc, hojaSummary);
    const valB26 = evSummary.valor(1, 25);   // Summary!B26 = "=Params!LocalOnly"
    prueba(
      "X1a. tf_excel.xlsx importa LocalOnly con localSheetId=3 (hoja Params) y Summary!B26 (=Params!LocalOnly) devuelve 'USD'",
      xl.doc && xl.doc.nombres && !!xl.doc.nombres["params!localonly"] && valB26 === "USD",
      `nombres=${JSON.stringify(xl.doc && xl.doc.nombres)} B26=${JSON.stringify(valB26)}`
    );
  }

  // X1b: precedencia local > global, Hoja!Nombre con espacios y error #NOMBRE? si no existe en esa hoja
  const docAmb = new env.NOM.DocCRDT("xl-amb", null, { local: true, soloMemoria: true });
  docAmb.cambiar("_doc", {
    modo: "hoja",
    nombres: {
      tasa: { n: "Tasa", d: "=0.16" },
      "norte!tasa": { n: "Tasa", d: "=0.08", h: "Norte" },
      "mi hoja!extra": { n: "Extra", d: "=100", h: "Mi Hoja" },
      "norte!doble": { n: "Doble", d: "=LAMBDA(x; x*2)", h: "Norte" },
    },
  });
  const idNorte = "hnorte", idSur = "hsur";
  docAmb.cambiar(idNorte, { _vivo: true, tipo: "tabla", nombre: "Norte", orden: 0, "c:A1": "=Tasa", "c:A2": "='Mi Hoja'!Extra", "c:A3": "=Sur!Tasa", "c:A4": "=Norte!Doble(7)" });
  docAmb.cambiar(idSur, { _vivo: true, tipo: "tabla", nombre: "Sur", orden: 1, "c:A1": "=Tasa", "c:A2": "=Norte!Tasa", "c:A3": "=Norte!Doble(9)" });
  docAmb.cambiar("hmi", { _vivo: true, tipo: "tabla", nombre: "Mi Hoja", orden: 2 });
  const evNorte = F.evaluadorDeTabla(docAmb, docAmb.bloque(idNorte));
  const evSur = F.evaluadorDeTabla(docAmb, docAmb.bloque(idSur));
  const rN1 = evNorte.valor(0, 0), rN2 = evNorte.valor(0, 1), rN3 = evNorte.valor(0, 2), rN4 = evNorte.valor(0, 3);
  const rS1 = evSur.valor(0, 0), rS2 = evSur.valor(0, 1), rS3 = evSur.valor(0, 2);
  prueba(
    "X1b. Ámbito de hoja: local gana sobre global en su hoja, 'Mi Hoja'!Nombre y LAMBDA local funcionan, hoja ajena sin local → #NOMBRE?",
    rN1 === 0.08 && rS1 === 0.16 && rS2 === 0.08 && rN2 === 100
      && (rN3 instanceof F.ErrorF && rN3.codigo === "#NOMBRE?")
      && rN4 === 14 && rS3 === 18,
    `N=[${rN1}, ${rN2}, ${rN3 && rN3.codigo || rN3}, ${rN4}] S=[${rS1}, ${rS2}, ${rS3}]`
  );

  // X1s: sintético XML completo vía leerXlsx con localSheetId="1" y fórmula =Params!LocalOnly
  {
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const xlsxSin = {
      "xl/workbook.xml": `<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/><sheet name="Params" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="LocalOnly" localSheetId="1">Params!$B$2</definedName></definedNames></workbook>`,
      "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>`,
      "xl/worksheets/sheet1.xml": `<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1"><f>Params!LocalOnly</f></c></row></sheetData></worksheet>`,
      "xl/worksheets/sheet2.xml": `<worksheet xmlns="${S}"><sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>USD</t></is></c></row></sheetData></worksheet>`,
    };
    const abrirZipReal = env.NOM.origen.abrirZip;
    env.NOM.origen.abrirZip = () => ({ texto: (r) => xlsxSin[r] || null, leer: () => null, bytes: () => null, nombres: () => Object.keys(xlsxSin) });
    let xlSin;
    try { xlSin = await env.NOM.ooxml.leerXlsx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) }); }
    finally { env.NOM.origen.abrirZip = abrirZipReal; }
    const docSin = new env.NOM.DocCRDT("xl-x1s", null, { local: true, soloMemoria: true });
    docSin.cambiar("_doc", Object.assign({ modo: "hoja" }, xlSin.doc || {}));
    xlSin.hojas.forEach((h, i) => docSin.cambiar("h" + i, Object.assign({ _vivo: true, tipo: "tabla", nombre: h.nombre, orden: i }, h.campos)));
    const evS = F.evaluadorDeTabla(docSin, docSin.bloque("h0"));
    const vSin = evS.valor(0, 0);
    prueba(
      "X1s. Sintético XML: leerXlsx importa definedName con localSheetId y evalúa =Params!LocalOnly → 'USD'",
      xlSin.doc && xlSin.doc.nombres && !!xlSin.doc.nombres["params!localonly"] && vSin === "USD",
      `nombres=${JSON.stringify(xlSin.doc && xlSin.doc.nombres)} vSin=${JSON.stringify(vSin)}`
    );
  }

  // ---- X2: ninguna gráfica produce <rect> con width o height negativos (ni en stress_test ni en tamaños chicos/cero)
  const ctxGraf = { color: (n) => { const s = n.getElementsByTagNameNS("*", "srgbClr")[0]; return s ? "#" + s.getAttribute("val") : null; }, acentos: ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"], fuente: () => "Calibri" };
  const tamanos = [[0, 0], [10, 10], [80, 60], [150, 100], [200, 120], [400, 260]];
  const negativos = [];
  let probadosRef = 0;
  for (const arch of ["excel_ooxml_stress_test_max.xlsx", "tf_excel.xlsx"]) {
    let ruta = path.join(dirRef, arch);
    if (!fs.existsSync(ruta) && !process.env.NOM_REF_DIR && arch === "excel_ooxml_stress_test_max.xlsx") ruta = path.join(dirRaiz, arch);
    if (!fs.existsSync(ruta)) { saltar(arch); continue; }
    probadosRef++;
    const z = abrirZipBuf(fs.readFileSync(ruta));
    for (const rutaC of z.nombres().filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n))) {
      const m = env.NOM.graficaOffice.leer(z.texto(rutaC), ctxGraf);
      for (const [w, h] of tamanos) {
        const svg = env.NOM.graficaOffice.svg(m, w, h, { pt: 96 / 72 });
        const mal = svg.match(/<rect\b[^>]*(?:width|height)\s*=\s*"-[^"]*"/g);
        if (mal) negativos.push(`${arch}:${rutaC}@${w}x${h} -> ${mal[0]}`);
      }
    }
  }
  if (probadosRef > 0) {
    prueba(
      "X2a. Ninguna gráfica de excel_ooxml_stress_test_max.xlsx ni tf_excel.xlsx genera <rect> con width/height negativos",
      negativos.length === 0,
      `negativos=${negativos.slice(0, 3).join("; ")}`
    );
  }

  // X2b: también svgEx (histograma, cascada, embudo, rectángulos, caja) a tamaños chicos/cero
  const negativosEx = [];
  const modelosEx = [
    { ex: true, tipo: "histograma", titulo: { texto: "H" }, cat: { niveles: [["A", "B", "C"]] }, val: { niveles: [[10, 20, 30]] }, ejes: [{ val: true }, { val: false }], acentos: ["#4472C4"], letra: {} },
    { ex: true, tipo: "cascada", titulo: { texto: "C" }, leyenda: { pos: "t" }, cat: { niveles: [["A", "B", "T"]] }, val: { niveles: [[10, -5, 5]] }, totales: [2], ejes: [{ val: true }, { val: false }], acentos: ["#4472C4", "#ED7D31", "#A5A5A5"], letra: {} },
    { ex: true, tipo: "embudo", titulo: { texto: "E" }, cat: { niveles: [["Etiqueta muy larga", "B"]] }, val: { niveles: [[100, 50]] }, ejes: [{ val: false }], acentos: ["#4472C4"], letra: {} },
    { ex: true, tipo: "rectangulos", titulo: { texto: "R" }, leyenda: { pos: "t" }, cat: { niveles: [["A", "B"]] }, val: { niveles: [[40, 60]] }, ejes: [], acentos: ["#4472C4", "#ED7D31"], letra: {} },
    { ex: true, tipo: "caja", titulo: { texto: "B" }, cat: { niveles: [["G1", "G1", "G1", "G1"]] }, val: { niveles: [[1, 2, 3, 4]] }, vis: { media: true, puntos: true, atipicos: true }, ejes: [{ val: true }, { val: false }], acentos: ["#4472C4"], letra: {} },
  ];
  for (const m of modelosEx) {
    for (const [w, h] of [[0, 0], [10, 10], [40, 30]]) {
      const svg = env.NOM.graficaOffice.svg(m, w, h, { pt: 96 / 72 });
      const mal = svg.match(/<rect\b[^>]*(?:width|height)\s*=\s*"-[^"]*"/g);
      if (mal) negativosEx.push(`${m.tipo}@${w}x${h} -> ${mal[0]}`);
    }
  }
  prueba(
    "X2b. Gráficas extendidas (svgEx) tampoco generan <rect> con width/height negativos a tamaños pequeños o cero",
    negativosEx.length === 0,
    `negativosEx=${negativosEx.slice(0, 3).join("; ")}`
  );

  // X2s: sintético XML de gráficas estándar (graficaOffice.svg) y propias (graficas.svg) a tamaños cero y chicos
  {
    const C = "http://schemas.openxmlformats.org/drawingml/2006/chart", A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const xmls = [
      `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>S1</c:v></c:tx><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>Ene</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>25</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>`,
      `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>S1</c:v></c:tx><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>Categoría muy larga</c:v></c:pt><c:pt idx="1"><c:v>Otra</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>25</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:axPos val="l"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="b"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`,
    ];
    const negSin = [];
    for (const xml of xmls) {
      const m = env.NOM.graficaOffice.leer(xml, ctxGraf);
      for (const [w, h] of [[0, 0], [10, 10], [40, 30], [80, 60]]) {
        const mal = env.NOM.graficaOffice.svg(m, w, h, { pt: 96 / 72 }).match(/<rect\b[^>]*(?:width|height)\s*=\s*"-[^"]*"/g);
        if (mal) negSin.push(`office@${w}x${h}->${mal[0]}`);
      }
    }
    const dNom = { etiquetas: ["A", "B"], series: [{ nombre: "S1", valores: [10, 20] }] };
    for (const cl of ["barras", "barras-h", "lineas", "area", "pastel"]) {
      for (const [w, h] of [[0, 0], [10, 10], [40, 30]]) {
        const mal = env.NOM.graficas.svg(dNom, cl, { ancho: w, alto: h, autonomo: true }).match(/<rect\b[^>]*(?:width|height)\s*=\s*"-[^"]*"/g);
        if (mal) negSin.push(`nom:${cl}@${w}x${h}->${mal[0]}`);
      }
    }
    prueba(
      "X2s. Sintético: gráficas estándar (graficaOffice y graficas.js) no generan <rect> con width/height negativos en tamaños 0 o chicos",
      negSin.length === 0,
      `negSin=${negSin.slice(0, 3).join("; ")}`
    );
  }

  console.log(fallas ? `${fallas} FALLA(S)` : "Todas las pruebas pasaron");
  process.exit(fallas ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });

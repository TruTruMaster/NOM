// Pruebas de fidelidad con Word 2024 sobre tf_word.docx y casos unitarios (W1–W6).
// Uso: node 4-CODIGO/app/src/test/js/office-word.test.js
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

// Lector ZIP mínimo en Node puro (solo para pruebas, usa zlib nativo).
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
      nodeType: 1,
      tagName: tag,
      localName,
      namespaceURI: ns,
      parentNode: padre || null,
      attributes: attrLista,
      childNodes: [],
      children: [],
      _nsMapa: nsMapa,
      get textContent() { return this.childNodes.map((c) => (c.nodeType === 3 ? c.data : c.textContent)).join(""); },
      getAttribute(k) { const a = attrLista.find((x) => x.name === k); return a ? a.value : null; },
      getAttributeNS(u, l) { const a = attrLista.find((x) => x.localName === l && (u === "*" || x.ns === u)); return a ? a.value : null; },
      hasAttribute(k) { return attrLista.some((x) => x.name === k); },
      getElementsByTagName(nm) {
        const out = [];
        const rec = (n) => { for (const c of n.children) { if (nm === "*" || c.tagName === nm || c.localName === nm) out.push(c); rec(c); } };
        rec(this);
        return out;
      },
      getElementsByTagNameNS(u, l) {
        const out = [];
        const rec = (n) => { for (const c of n.children) { if ((l === "*" || c.localName === l) && (u === "*" || c.namespaceURI === u)) out.push(c); rec(c); } };
        rec(this);
        return out;
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
        tope.childNodes.push(el);
        tope.children.push(el);
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
    modos: {}, leer: (k, d) => d, guardar: () => {}, api: { usuario: () => ({ id: "u1", nombre: "Prueba" }) },
    srcBin: (v) => v,
  };
  const document = {
    createElement: () => ({ style: {}, dataset: {}, classList: { add() {}, remove() {}, toggle() {}, contains: () => false }, getContext: () => ({ measureText: (t) => ({ width: String(t).length * 7 }) }) }),
    addEventListener() {}, removeEventListener() {}, querySelector: () => null,
  };
  const window = { NOM, document, matchMedia: () => ({ matches: false }), addEventListener() {}, removeEventListener() {} };
  const sandbox = {
    window, NOM, document, matchMedia: window.matchMedia, console, TextEncoder, TextDecoder,
    atob: (s) => Buffer.from(s, "base64").toString("binary"),
    btoa: (s) => Buffer.from(s, "binary").toString("base64"),
    setTimeout, clearTimeout, DOMParser: crearDOMParserMinimo(),
    CSS: { escape: (s) => String(s) },
  };
  const ctx = vm.createContext(sandbox);
  const cargar = (archivo) => vm.runInContext(leerJs(archivo), ctx, { filename: archivo });
  return { NOM, cargar };
}

(async () => {
  const env = crearEntorno();
  env.cargar("crdt.js");
  env.cargar("secuencia.js");
  env.cargar("campos.js");
  env.cargar("compresion.js");
  env.cargar("origen.js");
  env.cargar("formulas.js");
  env.cargar("ooxml.js");
  env.cargar("modo-texto.js");

  const abrirZipReal = env.NOM.origen.abrirZip;
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const textoPar = (p) => p.tramos.map((t) => t.x || "").join("");
  const cerca = (a, b) => Math.abs((+a || 0) - b) < 0.01;
  const esperadoPieSec2 = "Body odd — Page {PAGE} of {NUMPAGES} (section {SECTION} of {SECTIONPAGES})";
  const esperadoPieSec2Par = "Body even — Page {PAGE} of {NUMPAGES} (section {SECTION} of {SECTIONPAGES})";
  const esperadoPieSec5Pri = "Landscape first — Page {PAGE} of {NUMPAGES} (section {SECTION} of {SECTIONPAGES})";

  const rutaWord = path.join(dirRef, "tf_word.docx");
  let res = null;
  if (!fs.existsSync(rutaWord)) {
    saltar("tf_word.docx");
  } else {
    const bufWord = fs.readFileSync(rutaWord);
    res = await env.NOM.ooxml.leerDocx({ size: bufWord.length, arrayBuffer: async () => bufWord.buffer.slice(bufWord.byteOffset, bufWord.byteOffset + bufWord.byteLength) });

    // ---- W1: pie de página con campos PAGE / NUMPAGES / SECTION / SECTIONPAGES en su lugar
    const secs = res.parrafos.filter((p) => p.pa && p.pa.sx).map((p) => JSON.parse(p.pa.sx));
    const sec1 = secs[0] || {}, sec2 = secs[1] || {}, sec5 = secs[4] || {};

    prueba(
      "W1a. tf_word.docx Sec 2 conserva el texto literal y los 4 campos ({PAGE}/{NUMPAGES}/{SECTION}/{SECTIONPAGES}) sin agregar 'n de total'",
      sec2.ep && sec2.ep.p === esperadoPieSec2 && sec2.ep.np === "no" && sec2.sn === 2,
      `ep=${JSON.stringify(sec2.ep)} sn=${sec2.sn}`
    );
    prueba(
      "W1b. tf_word.docx lee pie par (even) de Sec 2 y primera página distinta (first) de Sec 1 y Sec 5",
      sec2.ep && sec2.ep.pp === esperadoPieSec2Par && sec2.ep.par === 1
        && sec1.ep && sec1.ep.t1 === 1 && sec1.ep.p1 === ""
        && sec5.ep && sec5.ep.t1 === 1 && sec5.ep.p1 === esperadoPieSec5Pri && sec5.sn === 5,
      `sec1.ep=${JSON.stringify(sec1.ep)} sec2.ep.pp=${JSON.stringify(sec2.ep && sec2.ep.pp)} sec5.ep=${JSON.stringify(sec5.ep)}`
    );

    // ---- W2: casillas (MS Gothic 2610/2612) y w:sym (Wingdings F0A9)
    const pCheck = res.parrafos.find((p) => textoPar(p).includes("Checkboxes:"));
    const txtCheck = pCheck ? textoPar(pCheck) : "";
    prueba(
      "W2a. tf_word.docx casillas MS Gothic 2610 y 2612 salen como ☐ y ☒ (no ˆ / U+0610)",
      txtCheck.includes("☐") && txtCheck.includes("☒") && !/[\u0610\u0612]/.test(txtCheck),
      `txtCheck=${JSON.stringify(txtCheck)}`
    );
    const pSym = res.parrafos.find((p) => textoPar(p).includes("Symbols:"));
    const txtSym = pSym ? textoPar(pSym) : "";
    prueba(
      "W2b. tf_word.docx w:sym Wingdings F0A9 mapea a Unicode (◆) y no a © (0xA9)",
      txtSym.includes("Symbols: ◆") || (txtSym.includes("◆") && !txtSym.startsWith("Symbols: ©")),
      `txtSym=${JSON.stringify(txtSym)}`
    );

    // ---- W3: PAGEREF a marcador sin valor guardado (w:dirty="true")
    const camposDoc = [];
    const marcadoresDoc = [];
    res.parrafos.forEach((p) => p.tramos.forEach((t) => {
      if (t.e && t.e.tipo === "campo") camposDoc.push(t.e);
      if (t.e && t.e.tipo === "marcador") marcadoresDoc.push(t.e);
    }));
    const cPageRef = camposDoc.find((c) => /^PAGEREF\s+bm_formatting$/i.test(c.instr));
    const mTarget = marcadoresDoc.find((m) => m.nombre === "bm_formatting" && m.lado === "ini");
    prueba(
      "W3a. tf_word.docx importa marcador bm_formatting y campo PAGEREF bm_formatting marcado como sucio",
      !!mTarget && !!cPageRef && cPageRef.sucio === true,
      `marcador=${JSON.stringify(mTarget)} campo=${JSON.stringify(cPageRef)}`
    );

    // ---- W4: estilos de fábrica de Word omitidos en styles.xml (TOC1..TOC5)
    const tocs = {};
    ["1. Text Formatting", "1.1 Character", "1.3.1 Level three", "1.3.1.1 Level four", "1.3.1.1.1 Level five"].forEach((pref, idx) => {
      const p = res.parrafos.find((q) => textoPar(q).startsWith(pref));
      tocs["TOC" + (idx + 1)] = p ? p.pa : null;
    });
    prueba(
      "W4a. tf_word.docx TOC1..TOC5 ausentes en styles.xml reciben espacio después 5 pt y sangría izquierda 0/220/440/660/880 twips",
      tocs.TOC1 && tocs.TOC1.despues === 5 && !tocs.TOC1.sangiz
        && tocs.TOC2 && tocs.TOC2.despues === 5 && cerca(tocs.TOC2.sangiz, 220 / (1440 / 2.54))
        && tocs.TOC3 && tocs.TOC3.despues === 5 && cerca(tocs.TOC3.sangiz, 440 / (1440 / 2.54))
        && tocs.TOC4 && tocs.TOC4.despues === 5 && cerca(tocs.TOC4.sangiz, 660 / (1440 / 2.54))
        && tocs.TOC5 && tocs.TOC5.despues === 5 && cerca(tocs.TOC5.sangiz, 880 / (1440 / 2.54)),
      `tocs=${JSON.stringify(tocs)}`
    );

    // ---- W5 / W6 en tf_word.docx
    const pDer = res.parrafos.find((p) => textoPar(p) === "Right aligned");
    prueba(
      "W5a. tf_word.docx importa 'Right aligned' con alin='der'",
      !!pDer && pDer.pa.alin === "der",
      `pDer=${JSON.stringify(pDer && pDer.pa)}`
    );
    const pDist = res.parrafos.find((p) => textoPar(p) === "Distributed alignment");
    prueba(
      "W6a. tf_word.docx importa w:jc='distribute' como alin='dist' (no 'just')",
      !!pDist && pDist.pa.alin === "dist",
      `pDist.pa=${JSON.stringify(pDist && pDist.pa)}`
    );
  }

  // ---- Casos sintéticos W1–W6 (no dependen de tf_word.docx)
  const fnExp = env.NOM.textoRico && env.NOM.textoRico.expandirEncPie;
  const expandido = fnExp ? fnExp(esperadoPieSec2, { pagina: "3", total: "9", seccion: "2", paginasSeccion: "2" }) : null;
  prueba(
    "W1c. expandirEncPie sustituye PAGE/NUMPAGES/SECTION/SECTIONPAGES en su lugar (igual que Word)",
    expandido === "Body odd — Page 3 of 9 (section 2 of 2)",
    `expandido=${JSON.stringify(expandido)}`
  );

  // W1s: sintético XML con 2 secciones, pie default + even + first y campos PAGE/NUMPAGES/SECTION/SECTIONPAGES
  {
    const docxW1 = {
      "word/settings.xml": `<w:settings xmlns:w="${W}"><w:evenAndOddHeaders/></w:settings>`,
      "word/document.xml": `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>
        <w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:footerReference w:type="first" r:id="rF3"/><w:titlePg/></w:sectPr></w:pPr><w:r><w:t>Sec 1</w:t></w:r></w:p>
        <w:p><w:r><w:t>Sec 2</w:t></w:r></w:p>
        <w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:footerReference w:type="default" r:id="rF1"/><w:footerReference w:type="even" r:id="rF2"/><w:footerReference w:type="first" r:id="rF3"/><w:titlePg/></w:sectPr>
      </w:body></w:document>`,
      "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rF1" Target="footer1.xml"/><Relationship Id="rF2" Target="footer2.xml"/><Relationship Id="rF3" Target="footer3.xml"/></Relationships>`,
      "word/footer1.xml": `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t xml:space="preserve">P </w:t></w:r><w:fldSimple w:instr=" PAGE "/><w:r><w:t>/</w:t></w:r><w:fldSimple w:instr=" NUMPAGES "/><w:r><w:t xml:space="preserve"> S </w:t></w:r><w:fldSimple w:instr=" SECTION "/><w:r><w:t>/</w:t></w:r><w:fldSimple w:instr=" SECTIONPAGES "/></w:p></w:ftr>`,
      "word/footer2.xml": `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t xml:space="preserve">Par </w:t></w:r><w:fldSimple w:instr=" PAGE "/></w:p></w:ftr>`,
      "word/footer3.xml": `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Primera</w:t></w:r></w:p></w:ftr>`,
    };
    env.NOM.origen.abrirZip = () => ({ texto: async (r) => docxW1[r] || null, leer: async () => null });
    const rW1 = await env.NOM.ooxml.leerDocx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) });
    env.NOM.origen.abrirZip = abrirZipReal;
    const sxs = rW1.parrafos.filter((p) => p.pa && p.pa.sx).map((p) => JSON.parse(p.pa.sx));
    const s2 = sxs[1] || {};
    prueba(
      "W1s. Sintético XML: leerDocx conserva campos PAGE/NUMPAGES/SECTION/SECTIONPAGES y pies par/primera",
      !!s2.ep && s2.ep.p === "P {PAGE}/{NUMPAGES} S {SECTION}/{SECTIONPAGES}" && s2.ep.pp === "Par {PAGE}" && s2.ep.par === 1 && s2.ep.t1 === 1 && s2.ep.p1 === "Primera" && s2.sn === 2,
      `sxs=${JSON.stringify(sxs)}`
    );
  }

  const docxSyms = {
    "word/document.xml": `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>
      <w:p>
        <w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>
        <w:r><w:sym w:font="Wingdings" w:char="F0FB"/></w:r>
        <w:r><w:sym w:font="Wingdings" w:char="F0FE"/></w:r>
        <w:r><w:sym w:font="Wingdings" w:char="F0FD"/></w:r>
        <w:r><w:sym w:font="Wingdings" w:char="F0A8"/></w:r>
        <w:r><w:sym w:font="Wingdings" w:char="F0A9"/></w:r>
        <w:r><w:sym w:font="Wingdings 2" w:char="52"/></w:r>
        <w:r><w:sym w:font="Symbol" w:char="F061"/></w:r>
        <w:r><w:sym w:font="Symbol" w:char="F0B1"/></w:r>
        <w:r><w:sym w:font="MS Gothic" w:char="2610"/></w:r>
        <w:r><w:sym w:font="MS Gothic" w:char="2612"/></w:r>
        <w:sdt><w:sdtPr><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr><w:sdtContent/></w:sdt>
      </w:p>
    </w:body></w:document>`,
  };
  env.NOM.origen.abrirZip = () => ({ texto: async (r) => docxSyms[r] || null, leer: async () => null });
  const resSyms = await env.NOM.ooxml.leerDocx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) });
  env.NOM.origen.abrirZip = abrirZipReal;
  const txtSyms = textoPar(resSyms.parrafos[0]);
  prueba(
    "W2c. w:sym Wingdings/Wingdings 2/Symbol/MS Gothic (2610/2612) y respaldo w14:checkbox mapean a Unicode correcto",
    txtSyms === "✓✗☑☒☐◆☑α±☐☒☒",
    `txtSyms=${JSON.stringify(txtSyms)}`
  );

  // W3s: sintético XML marcador + PAGEREF dirty
  {
    const docxW3 = {
      "word/document.xml": `<w:document xmlns:w="${W}"><w:body>
        <w:p><w:bookmarkStart w:id="1" w:name="bm_x"/><w:r><w:t>Destino</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText> PAGEREF bm_x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
      </w:body></w:document>`,
    };
    env.NOM.origen.abrirZip = () => ({ texto: async (r) => docxW3[r] || null, leer: async () => null });
    const rW3 = await env.NOM.ooxml.leerDocx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) });
    env.NOM.origen.abrirZip = abrirZipReal;
    const trs = rW3.parrafos.flatMap((p) => p.tramos);
    const m1 = trs.find((t) => t.e && t.e.tipo === "marcador" && t.e.nombre === "bm_x");
    const c1 = trs.find((t) => t.e && t.e.tipo === "campo" && /PAGEREF\s+bm_x/i.test(t.e.instr));
    prueba(
      "W3s. Sintético XML: leerDocx importa marcador y campo PAGEREF marcado w:dirty='true'",
      !!m1 && !!c1 && c1.e.sucio === true,
      `m1=${JSON.stringify(m1)} c1=${JSON.stringify(c1)}`
    );
  }

  // Prueba de cálculo automático en vista para campos vacíos/sucios vs valor guardado
  const fnAuto = env.NOM.textoRico && env.NOM.textoRico.valorCampoAuto;
  const docPrueba = new env.NOM.DocCRDT("w-test", null, { local: true, soloMemoria: true });
  docPrueba.cambiar("_doc", { modo: "texto" });
  const idC1 = "c1", idC2 = "c2";
  docPrueba.cambiar(idC1, { _vivo: true, tipo: "campo", instr: "PAGEREF bm_formatting", res: "", sucio: true });
  docPrueba.cambiar(idC2, { _vivo: true, tipo: "campo", instr: "PAGEREF bm_guardado", res: "7" });
  docPrueba.seq("cuerpo", { k: "i", iz: null, x: "Título" });
  const ult1 = () => { const v = docPrueba.secuencia("cuerpo").items.filter((x) => x.v); return v.length ? v[v.length - 1].id : null; };
  docPrueba.seq("cuerpo", { k: "i", iz: ult1(), e: { tipo: "marcador", nombre: "bm_formatting", lado: "ini" } });
  const idItemM = ult1();
  docPrueba.seq("cuerpo", { k: "i", iz: ult1(), e: { tipo: "campo", ref: idC1 } });
  docPrueba.seq("cuerpo", { k: "i", iz: ult1(), e: { tipo: "campo", ref: idC2 } });
  const mapaAuto = fnAuto ? fnAuto(docPrueba, { paginaDe: (id) => (id === idItemM ? 2 : 1), paginas: 9 }) : null;
  prueba(
    "W3b. valorCampoAuto calcula PAGEREF bm_formatting vacío/sucio → '2' y conserva el valor guardado ('7') si no está sucio",
    !!mapaAuto && mapaAuto.get(idC1) === "2" && !mapaAuto.has(idC2),
    `mapa=${mapaAuto ? JSON.stringify([...mapaAuto.entries()]) : null}`
  );

  // W4b: estilos integrados ausentes en styles.xml (Caption, Title, Subtitle, Quote y TOC1/TOC2)
  const docxFab = {
    "word/document.xml": `<w:document xmlns:w="${W}"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>Figura 1</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Portada</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>Sub</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>Cita</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:r><w:t>Nivel 2</w:t></w:r></w:p>
    </w:body></w:document>`,
    "word/styles.xml": `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
  };
  env.NOM.origen.abrirZip = () => ({ texto: async (r) => docxFab[r] || null, leer: async () => null });
  const resFab = await env.NOM.ooxml.leerDocx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) });
  env.NOM.origen.abrirZip = abrirZipReal;
  const [pCap, pTit, pSub, pCit, pToc2] = resFab.parrafos;
  prueba(
    "W4b. Estilos integrados ausentes en styles.xml (Caption, Title, Subtitle, Quote, TOC2) usan valores de fábrica de Word",
    pCap.pa.despues === 10 && pCap.tramos[0].a.i === true && pCap.tramos[0].a.tam === 9
      && pTit.pa.estilo === "titulo" && pTit.tramos[0].a.tam === 28
      && pSub.pa.estilo === "subtitulo" && pSub.pa.despues === 8
      && pCit.pa.estilo === "cita" && pCit.pa.alin === "centro" && pCit.tramos[0].a.i === true
      && pToc2.pa.despues === 5 && cerca(pToc2.pa.sangiz, 220 / (1440 / 2.54)),
    `cap=${JSON.stringify(pCap)} tit=${JSON.stringify(pTit)} sub=${JSON.stringify(pSub)} cit=${JSON.stringify(pCit)} toc2=${JSON.stringify(pToc2)}`
  );

  // ---- W5: Right aligned no debe encoger A1 al ancho medido del primer glifo (causa del corte 'Right / aligned')
  const fnIni = env.NOM.textoRico && env.NOM.textoRico.inicioPrimerRenglon;
  const x0Der = fnIni ? fnIni({ alin: "der", izq: 96, textIndent: 0, z: 1, rtl: false, rectPrimero: 638 }) : null;
  const x0Cen = fnIni ? fnIni({ alin: "centro", izq: 96, textIndent: 0, z: 1, rtl: false, rectPrimero: 367 }) : null;
  const x0Izq = fnIni ? fnIni({ alin: "izq", izq: 96, textIndent: 24, z: 1, rtl: false, rectPrimero: 120 }) : null;
  prueba(
    "W5. Right aligned (y centrado) usa izq+textIndent y no el borde izquierdo del primer carácter desplazado",
    x0Der === 96 && x0Cen === 96 && x0Izq === 120,
    `x0Der=${x0Der} x0Cen=${x0Cen} x0Izq=${x0Izq}`
  );

  // ---- W6: Distributed alignment (w:jc="distribute") sintético XML + espaciado + roundtrip
  {
    const docxW6 = {
      "word/document.xml": `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:t>Distribuido</w:t></w:r></w:p></w:body></w:document>`,
    };
    env.NOM.origen.abrirZip = () => ({ texto: async (r) => docxW6[r] || null, leer: async () => null });
    const rW6 = await env.NOM.ooxml.leerDocx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) });
    env.NOM.origen.abrirZip = abrirZipReal;
    prueba(
      "W6s. Sintético XML: leerDocx importa w:jc='distribute' como alin='dist'",
      rW6.parrafos[0] && rW6.parrafos[0].pa.alin === "dist",
      `pa=${JSON.stringify(rW6.parrafos[0] && rW6.parrafos[0].pa)}`
    );
  }
  const fnDist = env.NOM.textoRico && env.NOM.textoRico.espaciadoDistribuido;
  const lsDist = fnDist ? fnDist(624, 144, "Distributed alignment".length) : null;
  prueba(
    "W6b. espaciadoDistribuido reparte el sobrante entre letras para llenar todo el ancho",
    typeof lsDist === "number" && lsDist > 23 && lsDist < 24.5 && fnDist(624, 700, 10) === 0 && fnDist(624, 10, 1) === 0,
    `lsDist=${lsDist}`
  );
  // Exportación roundtrip de alin='dist' → w:jc="distribute"
  const docDist = new env.NOM.DocCRDT("w-dist", null, { local: true, soloMemoria: true });
  docDist.cambiar("_doc", { modo: "texto" });
  docDist.seq("cuerpo", { k: "i", iz: null, x: "Distributed alignment" });
  docDist.seq("cuerpo", { k: "f", r: [["fin", 0, 0]], a: { alin: "dist" } });
  const u8Exp = env.NOM.ooxml.docx(docDist);
  const zExp = abrirZipBuf(Buffer.from(u8Exp));
  const docXmlExp = await zExp.texto("word/document.xml");
  prueba(
    "W6c. Exportar a .docx conserva alin='dist' como w:jc=\"distribute\"",
    !!docXmlExp && /<w:jc\s+w:val="distribute"\/>/.test(docXmlExp),
    `docXmlExp=${docXmlExp && docXmlExp.slice(0, 300)}`
  );

  console.log(fallas ? `${fallas} FALLA(S)` : "Todas las pruebas pasaron");
  process.exit(fallas ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });

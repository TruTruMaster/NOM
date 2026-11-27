// Pruebas de defensa contra archivos de Office armados (modo-texto.js, formulas.js, geometrias.js,
// modo-hoja.js, smartart.js, ooxml.js).
// Uso: node 4-CODIGO/app/src/test/js/office-armado.test.js
"use strict";
const fs = require("fs"), path = require("path"), vm = require("vm");

const dirJs = process.env.NOM_JS_DIR || path.join(__dirname, "../../main/resources/static/js");
const leerJs = (nombre) => fs.readFileSync(path.join(dirJs, nombre), "utf8");

let fallas = 0;
const prueba = (nombre, ok, detalle) => {
  console.log((ok ? "OK    " : "FALLA ") + nombre + (!ok && detalle ? " — " + detalle : ""));
  if (!ok) fallas++;
};

// Entorno mínimo para cargar módulos en un contexto vm (permite cortar bucles infinitos con timeout).
function crearEntorno() {
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const nodo = () => {
    const el = {
      style: {},
      dataset: {},
      hidden: false,
      innerHTML: "",
      textContent: "",
      value: "",
      scrollLeft: 0,
      scrollTop: 0,
      offsetHeight: 0,
      children: [],
      classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
      append() {},
      remove() {},
      focus() {},
      blur() {},
      addEventListener() {},
      removeEventListener() {},
      querySelector: () => null,
      querySelectorAll: () => [],
      closest: () => null,
      getAttribute: () => null,
      setAttribute() {},
      removeAttribute() {},
      getBoundingClientRect: () => ({ left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 }),
    };
    return el;
  };
  const mapaDom = new Map();
  const raizHtml = (texto) => {
    const el = nodo();
    el._html = texto;
    if (texto.includes("data-rejilla")) {
      const rej = nodo(), fx = nodo(), refEl = nodo(), filtro = nodo(), graf = nodo(), esq = nodo(), band = nodo(), calc = nodo();
      mapaDom.set(el, { "[data-rejilla]": rej, "[data-fx]": fx, "[data-ref]": refEl, "[data-filtro]": filtro, "[data-graficas]": graf, "[data-esquema]": esq, "[data-bandeja]": band, "[data-calcular]": calc });
    }
    return el;
  };
  const $ = (sel, ctx) => (mapaDom.has(ctx) ? mapaDom.get(ctx)[sel] || null : null);
  const $$ = () => [];
  const NOM = {
    esc,
    ico: () => "",
    html: raizHtml,
    $,
    $$,
    aviso: () => {},
    colorCss: (v, d = "") => v || d,
    modos: {},
    rueda: { enlazar() {} },
    graficas: { TIPOS: [], rango: () => null },
    leer: (k, d) => d,
    guardar: () => {},
    api: { clave: (k) => "u1." + k, usuario: () => ({ id: "u1" }) },
    binarios: { esRef: (v) => typeof v === "string" && v.length < 90 && /^nom:bin\/([0-9a-f]{64})(?:\.([a-z0-9]{1,8}))?$/.test(v) },
    srcBin: (v) => v,
  };
  const document = {
    createElement: (tag) => (tag === "canvas" ? { getContext: () => ({ fillStyle: "#000000" }) } : nodo()),
    addEventListener() {},
    removeEventListener() {},
    querySelector: () => null,
    activeElement: null,
  };
  const window = { NOM, document, matchMedia: () => ({ matches: false }), addEventListener() {}, removeEventListener() {} };
  const sandbox = {
    window,
    NOM,
    document,
    matchMedia: window.matchMedia,
    console,
    TextEncoder,
    TextDecoder,
    atob: (s) => Buffer.from(s, "base64").toString("binary"),
    btoa: (s) => Buffer.from(s, "binary").toString("base64"),
    setTimeout,
    clearTimeout,
    DOMParser: crearDOMParserMinimo(),
  };
  const ctx = vm.createContext(sandbox);
  const cargar = (archivo) => vm.runInContext(leerJs(archivo), ctx, { filename: archivo });
  return { ctx, sandbox, NOM, cargar, nodo, mapaDom };
}

// Analizador XML mínimo para probar smartart.js y ooxml.js en Node puro sin librerías externas.
function crearDOMParserMinimo() {
  const desEsc = (s) => String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
  function hacerEl(tag, attrs, nsHeredado) {
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
      attributes: attrLista,
      childNodes: [],
      children: [],
      _nsMapa: nsMapa,
      get textContent() {
        return this.childNodes.map((c) => (c.nodeType === 3 ? c.data : c.textContent)).join("");
      },
      getAttribute(k) {
        const a = attrLista.find((x) => x.name === k);
        return a ? a.value : null;
      },
      getAttributeNS(u, l) {
        const a = attrLista.find((x) => x.localName === l && (u === "*" || x.ns === u));
        return a ? a.value : null;
      },
      getElementsByTagName(nm) {
        const out = [];
        const rec = (n) => { for (const c of n.children) { if (nm === "*" || c.tagName === nm || c.localName === nm) out.push(c); rec(c); } };
        rec(this);
        return out;
      },
      getElementsByTagNameNS(u, l) {
        const out = [];
        const rec = (n) => {
          for (const c of n.children) {
            if ((l === "*" || c.localName === l) && (u === "*" || c.namespaceURI === u)) out.push(c);
            rec(c);
          }
        };
        rec(this);
        return out;
      },
    };
    return el;
  }
  function parsear(xml) {
    const raiz = hacerEl("#doc", [], Object.create(null));
    const pila = [raiz];
    const re = /<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:\s+[^\s=/>]+(?:\s*=\s*"[^"]*"|\s*=\s*'[^']*'))*)\s*(\/?)>|([^<]+)/g;
    let m;
    while ((m = re.exec(String(xml || "")))) {
      const tope = pila[pila.length - 1];
      if (m[1]) {
        if (pila.length > 1) pila.pop();
      } else if (m[2]) {
        const attrs = [];
        const reA = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
        let a;
        while ((a = reA.exec(m[3] || ""))) attrs.push([a[1], desEsc(a[2] != null ? a[2] : a[3])]);
        const el = hacerEl(m[2], attrs, tope._nsMapa);
        tope.childNodes.push(el);
        tope.children.push(el);
        if (!m[4]) pila.push(el);
      } else if (m[5]) {
        tope.childNodes.push({ nodeType: 3, data: desEsc(m[5]) });
      }
    }
    raiz.documentElement = raiz.children[0] || null;
    return raiz;
  }
  return class DOMParser {
    parseFromString(texto) { return parsear(texto); }
  };
}

// ------------------------------------------------------------ 1. modo-texto.js: cssFuente
{
  const env = crearEntorno();
  env.cargar("modo-texto.js");
  const { cssFuente, altoFuente, envolver } = env.NOM.textoRico;

  const rCon = cssFuente("constructor");
  const rToS = cssFuente("toString");
  const rPro = cssFuente("__proto__");
  prueba(
    "1. modo-texto.js cssFuente('constructor') no devuelve una función ni texto de Object",
    typeof rCon === "string" && rCon === "'constructor','Carlito',sans-serif",
    `obtenido: ${typeof rCon} ${JSON.stringify(rCon)}`
  );
  prueba(
    "1. modo-texto.js cssFuente('toString'/'__proto__') usa fallback genérico sin heredar de Object.prototype",
    typeof rToS === "string" && rToS === "'toString','Carlito',sans-serif" && typeof rPro === "string" && rPro === "'__proto__','Carlito',sans-serif",
    `toString=${typeof rToS} (${String(rToS)}), __proto__=${typeof rPro} (${String(rPro)})`
  );
  prueba(
    "1. modo-texto.js cssFuente casos normales (Calibri, Aptos, Helvetica, desconocida)",
    cssFuente("Calibri") === "'NOM CAL','Carlito','Calibri',sans-serif"
      && cssFuente("Aptos") === "'NOM A','Aptos','Carlito',sans-serif"
      && cssFuente("Helvetica") === "'Helvetica','Liberation Sans','Arimo',sans-serif"
      && cssFuente("MiFuenteRara") === "'MiFuenteRara','Carlito',sans-serif",
    `Calibri=${cssFuente("Calibri")}`
  );

  // 1b. supFuente / altoFuente / espacioFuente (ejercitadas en envolver y altoFuente)
  const afCon = altoFuente("constructor");
  const envCon = envolver("x", { fuente: "constructor" });
  prueba(
    "1b. modo-texto.js supFuente/altoFuente/espacioFuente('constructor') devuelven respaldos numéricos",
    afCon === 1.2 && envCon.includes("--af:1.2;--ae:0.25;--supk:0.6499;"),
    `altoFuente=${typeof afCon} (${String(afCon)}), envolver=${envCon}`
  );
  const envAptos = envolver("x", { fuente: "Aptos" });
  prueba(
    "1b. modo-texto.js supFuente/altoFuente/espacioFuente caso normal (Aptos y Calibri)",
    altoFuente("Calibri") === 1.2207 && envAptos.includes("--af:1.2207;--ae:0.2031;--supk:0.6001;"),
    `altoCalibri=${altoFuente("Calibri")}, envAptos=${envAptos}`
  );
}

// ------------------------------------------------------------ 1c. modo-texto.js: combinadasTabla
{
  // Se sacan del archivo dimTabla y combinadasTabla tal cual (montar el editor de texto entero en Node no es fiable)
  // y se corren con las fórmulas reales; el HTML de la tabla se arma igual que en tablaHtml: colspan/rowspan de spans.
  function probarTablaTexto(camposTabla, timeoutMs) {
    const env = crearEntorno();
    env.cargar("formulas.js");
    const src = leerJs("modo-texto.js");
    const dim = /const dimTabla = \(b\) => \(\{[^\n]*\}\);/.exec(src);
    const ini = src.indexOf("function combinadasTabla(b) {");
    const fin = src.indexOf("return { spans, cubiertas };", ini);
    if (!dim || ini < 0 || fin < 0) throw new Error("no se encontró combinadasTabla en modo-texto.js");
    const cuerpo = src.slice(ini, src.indexOf("}", src.indexOf(";", fin)) + 1);
    const tabla = Object.assign({ id: "tab1", tipo: "tabla", enTexto: true, filasTexto: 3, colsTexto: 4 }, camposTabla);
    env.sandbox.__tabla = tabla;
    const r = vm.runInContext(`(() => { const FT = () => NOM.formulas; ${dim[0]} ${cuerpo}
      const { spans } = combinadasTabla(__tabla);
      return Object.entries(spans).map(([c, [w, h]]) => '<td data-celda="' + c + '" colspan="' + w + '" rowspan="' + h + '">').join(""); })()`, env.ctx, { timeout: timeoutMs });
    return String(r);
  }

  let hTabGig = null, eTabGig = null;
  try { hTabGig = probarTablaTexto({ "m:A1": "5000x5000", "c:A1": "T" }, 350); } catch (e) { eTabGig = e; }
  prueba(
    "1c. modo-texto.js combinadasTabla acota '5000x5000' a dimTabla (4x3) sin colgarse",
    eTabGig === null && hTabGig.includes('colspan="4" rowspan="3"') && !hTabGig.includes('colspan="5000"'),
    eTabGig ? String(eTabGig) : `html=${hTabGig.slice(0, 200)}`
  );

  let hTabFuera = null, eTabFuera = null;
  try { hTabFuera = probarTablaTexto({ "m:Z99": "5000x5000", "c:A1": "X" }, 350); } catch (e) { eTabFuera = e; }
  prueba(
    "1c. modo-texto.js combinadasTabla ignora combinada gigante fuera de dimTabla ('5000x5000' en Z99)",
    eTabFuera === null && !hTabFuera.includes("colspan=") && !hTabFuera.includes("rowspan="),
    eTabFuera ? String(eTabFuera) : "apareció colspan/rowspan inesperado"
  );

  const hTabNorm = probarTablaTexto({ "m:A1": "2x2", "c:A1": "Hola" }, 500);
  prueba(
    "1c. modo-texto.js combinadasTabla caso normal ('2x2' en A1 cubre B1, A2, B2)",
    hTabNorm === '<td data-celda="A1" colspan="2" rowspan="2">',
    "HTML de combinada normal en tabla de texto incorrecto"
  );
}

// ------------------------------------------------------------ 2. formulas.js: piezasXL (COLOR_XL)
{
  const env = crearEntorno();
  env.cargar("formulas.js");
  const { formatoXL } = env.NOM.formulas;

  const rCon = formatoXL(42, "[constructor]0");
  const rPro = formatoXL(42, "[__proto__]0");
  prueba(
    "2. formulas.js piezasXL: [constructor] no asigna Object como color",
    rCon.t === "42" && rCon.color === null,
    `color=${typeof rCon.color} (${String(rCon.color)})`
  );
  prueba(
    "2. formulas.js piezasXL: [__proto__] no asigna Object.prototype como color",
    rPro.t === "42" && rPro.color === null,
    `color=${typeof rPro.color}`
  );
  const rRojo = formatoXL(42, "[Red]0");
  const rAzul = formatoXL(1234.5, "[Azul]#,##0.00");
  prueba(
    "2. formulas.js piezasXL casos normales ([Red]0 y [Azul]#,##0.00)",
    rRojo.t === "42" && rRojo.color === "#FF0000" && rAzul.t === "1,234.50" && rAzul.color === "#0000FF",
    `rojo=${JSON.stringify(rRojo)}, azul=${JSON.stringify(rAzul)}`
  );
}

// ------------------------------------------------------------ 3. geometrias.js: htmlImagen y tablas fijas
{
  const env = crearEntorno();
  env.cargar("geometrias-datos.js");
  env.cargar("geometrias.js");
  const G = env.NOM.geomOffice;

  const hExt = G.htmlDibujo({ grupo: { hijos: [{ x: 0, y: 0, w: 50, h: 50, img: { src: "https://atacante.invalid/pixel.png" } }] } }, 100, 100);
  const hJs = G.htmlDibujo({ grupo: { hijos: [{ x: 0, y: 0, w: 50, h: 50, img: { src: "javascript:alert(1)" } }] } }, 100, 100);
  prueba(
    "3. geometrias.js htmlImagen rechaza URL externa https:// y javascript:",
    !hExt.includes("<img") && !hExt.includes("atacante.invalid") && !hJs.includes("<img") && !hJs.includes("javascript:"),
    `hExt=${hExt}`
  );

  const svgExt = G.svgForma({ prst: "rect", relleno: "#112233", pintura: { tipo: "img", src: "https://atacante.invalid/fondo.png" } }, 100, 50);
  prueba(
    "3. geometrias.js defRelleno rechaza pintura img externa https://",
    !svgExt.includes("atacante.invalid") && !svgExt.includes("<image"),
    `svgExt=${svgExt}`
  );

  const refBin = "nom:bin/" + "a".repeat(64) + ".png";
  const hData = G.htmlDibujo({ grupo: { hijos: [{ x: 0, y: 0, w: 50, h: 50, img: { src: "data:image/png;base64,AAAA" } }] } }, 100, 100);
  const hBin = G.htmlDibujo({ grupo: { hijos: [{ x: 0, y: 0, w: 50, h: 50, img: { src: refBin } }] } }, 100, 100);
  prueba(
    "3. geometrias.js htmlImagen casos normales (data:image/ y nom:bin/)",
    hData.includes('<img src="data:image/png;base64,AAAA"') && hBin.includes(`<img src="${refBin}"`),
    `hData=${hData}`
  );

  let errTrama = null, resTrama = "sin-correr";
  try { resTrama = G.trama("constructor", 1); } catch (e) { errTrama = e; }
  let errGuion = null, resGuion = "";
  try { resGuion = G.svgForma({ prst: "rect", relleno: "#112233", linea: { w: 1, c: "#000000", guion: "constructor" } }, 100, 50); } catch (e) { errGuion = e; }
  const resMatiz = G.matiz("#112233", "constructor");
  prueba(
    "3. geometrias.js tablas fijas (tiene/trazos/trama/matiz/guion) rechazan 'constructor'",
    G.tiene("constructor") === false
      && G.trazos("constructor", 100, 100) === null
      && errTrama === null && resTrama === null
      && resMatiz === "#112233"
      && errGuion === null && resGuion.startsWith("<svg") && !resGuion.includes("stroke-dasharray"),
    `tiene=${G.tiene("constructor")}, trazos=${JSON.stringify(G.trazos("constructor", 100, 100))}, errTrama=${errTrama}, matiz=${resMatiz}, errGuion=${errGuion}`
  );

  const trRect = G.trazos("rect", 100, 50);
  const tm50 = G.trama("pct50", 1);
  const svgDash = G.svgForma({ prst: "rect", relleno: "#112233", linea: { w: 1, c: "#000000", guion: "dash" } }, 100, 50);
  prueba(
    "3. geometrias.js casos normales (rect, pct50, darken, guion dash)",
    G.tiene("rect") === true
      && Array.isArray(trRect) && trRect.length === 1 && trRect[0].d.includes("M0,0")
      && typeof tm50 === "string" && tm50.startsWith("M0,0")
      && G.matiz("#112233", "darken") === "#0a141f"
      && svgDash.includes("stroke-dasharray="),
    `trRect=${JSON.stringify(trRect)}`
  );
}

// ------------------------------------------------------------ 4. modo-hoja.js: calcularCombinadas
{
  function probarHoja(camposTabla, timeoutMs) {
    const env = crearEntorno();
    env.cargar("formulas.js");
    env.cargar("modo-hoja.js");
    const tabla = Object.assign({ id: "t1", tipo: "tabla", pagina: "p1", filas: 6, cols: 5, nombre: "Hoja1" }, camposTabla);
    const bloques = { t1: tabla };
    const doc = {
      id: "d1",
      deLaPagina: () => [tabla],
      vivos: () => [tabla],
      bloques: () => [tabla],
      bloque: (id) => bloques[id] || null,
      valor: () => null,
      paginas: () => ["p1"],
      cambiar: () => {},
    };
    const lienzo = env.nodo();
    let raizCapturada = null;
    lienzo.append = (r) => { raizCapturada = r; };
    const ed = {
      doc,
      pagina: "p1",
      soloLectura: true,
      lienzo,
      seleccionar: () => {},
      pintarPresencia: () => {},
      refrescarHerramientas: () => {},
      vista: env.nodo(),
    };
    env.sandbox.__ed = ed;
    vm.runInContext("NOM.modos.HOJA.montar(__ed);", env.ctx, { timeout: timeoutMs });
    const rej = env.mapaDom.get(raizCapturada)["[data-rejilla]"];
    return rej.innerHTML;
  }

  let htmlGigante = null, errGigante = null;
  try { htmlGigante = probarHoja({ "m:A1": "5000x5000", "c:A1": "T" }, 350); } catch (e) { errGigante = e; }
  prueba(
    "4. modo-hoja.js calcularCombinadas acota '5000x5000' al tamaño de la hoja (5x6) sin colgarse",
    errGigante === null && htmlGigante.includes('colspan="5" rowspan="6"') && !htmlGigante.includes('colspan="5000"'),
    errGigante ? String(errGigante) : "no acotó colspan/rowspan a 5x6"
  );

  let htmlFuera = null, errFuera = null;
  try { htmlFuera = probarHoja({ "m:Z99": "5000x5000", "c:A1": "X" }, 350); } catch (e) { errFuera = e; }
  prueba(
    "4. modo-hoja.js calcularCombinadas ignora combinada gigante fuera de la hoja ('5000x5000' en Z99)",
    errFuera === null && !htmlFuera.includes("colspan=") && !htmlFuera.includes("rowspan="),
    errFuera ? String(errFuera) : "apareció colspan/rowspan inesperado"
  );

  const htmlNormal = probarHoja({ "m:A1": "3x2", "c:A1": "Hola" }, 500);
  prueba(
    "4. modo-hoja.js calcularCombinadas caso normal ('3x2' en A1 cubre B1, C1, A2, B2, C2)",
    htmlNormal.includes('data-id="t1!A1"')
      && htmlNormal.includes('colspan="3" rowspan="2"')
      && !htmlNormal.includes('data-id="t1!B1"')
      && !htmlNormal.includes('data-id="t1!C2"')
      && htmlNormal.includes('data-id="t1!D1"'),
    "HTML de combinada normal incorrecto"
  );
}

// ------------------------------------------------------------ 6. smartart.js: ciclos en q.padre (ancst, root, profundidad)
{
  const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const datosCicloXml = `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}">
    <dgm:ptLst>
      <dgm:pt modelId="doc0" type="doc"/>
      <dgm:pt modelId="n1" type="node"><dgm:t><a:p><a:r><a:t>Uno</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="n2" type="node"><dgm:t><a:p><a:r><a:t>Dos</a:t></a:r></a:p></dgm:t></dgm:pt>
    </dgm:ptLst>
    <dgm:cxnLst>
      <dgm:cxn type="parOf" srcId="doc0" destId="n1" srcOrd="0"/>
      <dgm:cxn type="parOf" srcId="n1" destId="n2" srcOrd="0"/>
      <dgm:cxn type="parOf" srcId="n2" destId="n1" srcOrd="0"/>
    </dgm:cxnLst>
  </dgm:dataModel>`;
  const datosNormalXml = `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}">
    <dgm:ptLst>
      <dgm:pt modelId="doc0" type="doc"/>
      <dgm:pt modelId="n1" type="node"><dgm:t><a:p><a:r><a:t>Uno</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="n2" type="node"><dgm:t><a:p><a:r><a:t>Dos</a:t></a:r></a:p></dgm:t></dgm:pt>
    </dgm:ptLst>
    <dgm:cxnLst>
      <dgm:cxn type="parOf" srcId="doc0" destId="n1" srcOrd="0"/>
      <dgm:cxn type="parOf" srcId="n1" destId="n2" srcOrd="0"/>
    </dgm:cxnLst>
  </dgm:dataModel>`;
  const disenoEje = (eje) => `<dgm:layoutDef xmlns:dgm="${DGM}">
    <dgm:layoutNode name="raiz">
      <dgm:alg type="lin"/>
      <dgm:forEach axis="ch" ptType="node">
        <dgm:forEach axis="${eje}" ptType="all">
          <dgm:layoutNode name="caja"><dgm:shape type="rect"/></dgm:layoutNode>
        </dgm:forEach>
      </dgm:forEach>
    </dgm:layoutNode>
  </dgm:layoutDef>`;
  const disenoProf = (op, val) => `<dgm:layoutDef xmlns:dgm="${DGM}">
    <dgm:layoutNode name="raiz">
      <dgm:alg type="lin"/>
      <dgm:forEach axis="des" ptType="node">
        <dgm:choose>
          <dgm:if func="depth" op="${op}" val="${val}">
            <dgm:layoutNode name="prof"><dgm:shape type="rect"/></dgm:layoutNode>
          </dgm:if>
        </dgm:choose>
      </dgm:forEach>
    </dgm:layoutNode>
  </dgm:layoutDef>`;

  function correrSmartArt(datosXml, disenoXml, timeoutMs) {
    const env = crearEntorno();
    env.cargar("smartart.js");
    const dp = new env.sandbox.DOMParser();
    env.sandbox.__p = { datos: dp.parseFromString(datosXml, "application/xml"), diseno: dp.parseFromString(disenoXml, "application/xml"), colores: null, estilo: null };
    env.sandbox.__o = { w: 1000000, h: 500000, color: () => "#336699", anchoLinea: () => 12700, medir: (t, pt) => t.length * pt * 7000 };
    return vm.runInContext("NOM.smartart.acomodar(__p, __o);", env.ctx, { timeout: timeoutMs });
  }

  let rAnc = null, eAnc = null;
  try { rAnc = correrSmartArt(datosCicloXml, disenoEje("ancst"), 300); } catch (e) { eAnc = e; }
  prueba(
    "6. smartart.js axis='ancst' con ciclo en parOf corta sin bucle infinito",
    eAnc === null && Array.isArray(rAnc),
    eAnc ? String(eAnc) : `res=${rAnc}`
  );

  let rRoot = null, eRoot = null;
  try { rRoot = correrSmartArt(datosCicloXml, disenoEje("root"), 300); } catch (e) { eRoot = e; }
  prueba(
    "6. smartart.js axis='root' con ciclo en parOf corta sin bucle infinito",
    eRoot === null && Array.isArray(rRoot),
    eRoot ? String(eRoot) : `res=${rRoot}`
  );

  let rProf = null, eProf = null;
  try { rProf = correrSmartArt(datosCicloXml, disenoProf("gt", "0"), 300); } catch (e) { eProf = e; }
  prueba(
    "6. smartart.js profundidad() con ciclo en parOf corta sin bucle infinito",
    eProf === null && Array.isArray(rProf),
    eProf ? String(eProf) : `res=${rProf}`
  );

  const rNormAnc = correrSmartArt(datosNormalXml, `<dgm:layoutDef xmlns:dgm="${DGM}">
    <dgm:layoutNode name="raiz"><dgm:alg type="lin"/>
      <dgm:forEach axis="des" ptType="node">
        <dgm:choose><dgm:if func="depth" op="equ" val="2">
          <dgm:forEach axis="ancst" ptType="all"><dgm:layoutNode name="a"><dgm:shape type="rect"/></dgm:layoutNode></dgm:forEach>
          <dgm:forEach axis="root" ptType="doc"><dgm:layoutNode name="r"><dgm:shape type="ellipse"/></dgm:layoutNode></dgm:forEach>
        </dgm:if></dgm:choose>
      </dgm:forEach>
    </dgm:layoutNode>
  </dgm:layoutDef>`, 300);
  prueba(
    "6. smartart.js caso normal sin ciclos (depth=2, ancst=2 ancestros, root=doc)",
    Array.isArray(rNormAnc)
      && rNormAnc.filter((f) => f.nombre === "a").length === 2
      && rNormAnc.filter((f) => f.nombre === "r").length === 1,
    `formas=${JSON.stringify(rNormAnc && rNormAnc.map((f) => f.nombre))}`
  );
}

// ------------------------------------------------------------ 7. ooxml.js: mapas Object.create(null)
(async () => {
  const env = crearEntorno();
  env.cargar("formulas.js");
  env.cargar("ooxml.js");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  // 7a. leerDocx: notasPie, notasFin, nombreMarcador, rels, formatoNum/textoNum con claves de Object.prototype
  const docxArmado = {
    "word/document.xml": `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>
      <w:p>
        <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="constructor"/></w:numPr></w:pPr>
        <w:r><w:footnoteReference w:id="constructor"/></w:r>
        <w:r><w:endnoteReference w:id="toString"/></w:r>
        <w:bookmarkEnd w:id="constructor"/>
        <w:hyperlink r:id="constructor"><w:r><w:t>clic</w:t></w:r></w:hyperlink>
      </w:p>
    </w:body></w:document>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  };
  env.NOM.origen = { abrirZip: () => ({ texto: async (r) => docxArmado[r] || null, leer: async () => null }) };
  const resDocxArm = await env.NOM.ooxml.leerDocx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) });
  const trArm = resDocxArm.parrafos[0].tramos;
  const notaPie = trArm.find((t) => t.e && t.e.tipo === "nota" && !t.e.final);
  const notaFin = trArm.find((t) => t.e && t.e.tipo === "nota" && t.e.final);
  const marcaFantasma = trArm.find((t) => t.e && t.e.tipo === "marcador");
  const enlaceArm = trArm.find((t) => t.x === "clic");
  prueba(
    "7. ooxml.js leerDocx: claves 'constructor'/'toString' no filtran métodos de Object.prototype",
    notaPie && notaPie.e.texto === ""
      && notaFin && notaFin.e.texto === ""
      && !marcaFantasma
      && enlaceArm && !enlaceArm.a.enlace,
    `notaPie=${typeof (notaPie && notaPie.e.texto)}, notaFin=${typeof (notaFin && notaFin.e.texto)}, marca=${!!marcaFantasma}`
  );

  // 7a normal
  const docxNormal = {
    "word/document.xml": `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>
      <w:p>
        <w:r><w:footnoteReference w:id="1"/></w:r>
        <w:hyperlink r:id="rId1"><w:r><w:t>sitio</w:t></w:r></w:hyperlink>
      </w:p>
    </w:body></w:document>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="https://ejemplo.org"/></Relationships>`,
    "word/footnotes.xml": `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:r><w:t>Nota uno</w:t></w:r></w:p></w:footnote></w:footnotes>`,
  };
  env.NOM.origen = { abrirZip: () => ({ texto: async (r) => docxNormal[r] || null, leer: async () => null }) };
  const resDocxNorm = await env.NOM.ooxml.leerDocx({ size: 100, arrayBuffer: async () => new ArrayBuffer(0) });
  const trNorm = resDocxNorm.parrafos[0].tramos;
  prueba(
    "7. ooxml.js leerDocx caso normal (nota al pie y enlace rId1)",
    trNorm.some((t) => t.e && t.e.tipo === "nota" && t.e.texto === "Nota uno")
      && trNorm.some((t) => t.x === "sitio" && t.a && t.a.enlace === "https://ejemplo.org"),
    `trNorm=${JSON.stringify(trNorm)}`
  );

  // 7b. modeloXlsx: destino y nombresXl con clave 'constructor'
  const xlsxArmado = {
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}">
      <sheets>
        <sheet name="Mala" sheetId="1" r:id="constructor"/>
        <sheet name="Buena" sheetId="2" r:id="rId1"/>
      </sheets>
      <definedNames><definedName name="IVA">0.16</definedName></definedNames>
    </workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>constructor+IVA</f></c></row></sheetData></worksheet>`,
  };
  let resXlsx = null, errXlsx = null;
  try {
    resXlsx = env.NOM.ooxml.modeloXlsx({ texto: (r) => xlsxArmado[r] || null }, 6);
  } catch (e) { errXlsx = e; }
  prueba(
    "7. ooxml.js modeloXlsx: r:id='constructor' y fórmula con 'constructor' no se rompen ni inyectan Object",
    errXlsx === null
      && resXlsx && resXlsx.hojas.length === 1
      && resXlsx.hojas[0].campos["c:A1"] === "=constructor+0.16",
    errXlsx ? String(errXlsx) : `c:A1=${resXlsx && resXlsx.hojas[0] && resXlsx.hojas[0].campos["c:A1"]}`
  );

  const xlsxNormal = {
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}">
      <sheets><sheet name="Ventas" sheetId="1" r:id="rId1"/></sheets>
      <definedNames><definedName name="IVA">0.16</definedName></definedNames>
    </workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>IVA*10</f></c></row></sheetData></worksheet>`,
  };
  const resXlsxNorm = env.NOM.ooxml.modeloXlsx({ texto: (r) => xlsxNormal[r] || null }, 6);
  prueba(
    "7. ooxml.js modeloXlsx caso normal (definedName IVA en fórmula)",
    resXlsxNorm.hojas.length === 1 && resXlsxNorm.hojas[0].campos["c:A1"] === "=0.16*10",
    `c:A1=${resXlsxNorm.hojas[0] && resXlsxNorm.hojas[0].campos["c:A1"]}`
  );

  // 7c. modeloXlsx: mergeCell acotando b.c y b.f a MAX_C-1 y MAX_F-1 (60x5000)
  const xlsxMerge = {
    "xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}">
      <sheets><sheet name="H1" sheetId="1" r:id="rId1"/></sheets>
    </workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
      <mergeCells count="2">
        <mergeCell ref="A1:XFD1048576"/>
        <mergeCell ref="B2:D4"/>
      </mergeCells>
    </worksheet>`,
  };
  const resMerge = env.NOM.ooxml.modeloXlsx({ texto: (r) => xlsxMerge[r] || null }, 13);
  const cm = resMerge.hojas[0].campos;
  prueba(
    "7c. ooxml.js mergeCell acota A1:XFD1048576 a MAX_C x MAX_F (60x5000)",
    cm["m:A1"] === "60x5000",
    `m:A1=${JSON.stringify(cm["m:A1"])}`
  );
  prueba(
    "7c. ooxml.js mergeCell caso normal (B2:D4 → 3x3)",
    cm["m:B2"] === "3x3",
    `m:B2=${JSON.stringify(cm["m:B2"])}`
  );

  console.log(fallas ? `${fallas} FALLA(S)` : "Todas las pruebas pasaron");
  process.exit(fallas ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });

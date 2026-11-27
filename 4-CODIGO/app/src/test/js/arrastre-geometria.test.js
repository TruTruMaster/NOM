// Pruebas de geometría de arrastre y redimensionamiento contra orillas en:
// - lienzo.js (diapositivas y diagramas)
// - modo-texto.js (imágenes flotantes en Word)
// - modo-hoja.js (imágenes sobre la hoja)
// Uso: node 4-CODIGO/app/src/test/js/arrastre-geometria.test.js
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const dirJs = process.env.NOM_JS_DIR || path.join(__dirname, "../../main/resources/static/js");
const leerJs = (n) => fs.readFileSync(path.join(dirJs, n), "utf8");

let fallas = 0;
const prueba = (nombre, ok, detalle) => {
  console.log((ok ? "OK    " : "FALLA ") + nombre + (!ok && detalle ? " — " + detalle : ""));
  if (!ok) fallas++;
};

function crearEntornoBasico(opciones) {
  opciones = opciones || {};
  const listenersWin = new Map();
  function crearClaseList() {
    const s = new Set();
    return {
      add(...xs) { xs.forEach((x) => x && s.add(x)); },
      remove(...xs) { xs.forEach((x) => s.delete(x)); },
      toggle(x, f) { if (f === undefined) { if (s.has(x)) s.delete(x); else s.add(x); } else if (f) s.add(x); else s.delete(x); },
      contains: (x) => s.has(x),
    };
  }
  function crearNodo(tag = "div") {
    const ls = new Map();
    const attrs = new Map();
    const el = {
      tagName: tag.toUpperCase(),
      style: {
        setProperty(k, v) { this[k] = v; },
        removeProperty(k) { delete this[k]; },
      },
      dataset: {},
      hidden: false,
      innerHTML: "",
      textContent: "",
      value: "",
      scrollLeft: 0,
      scrollTop: 0,
      offsetLeft: 0,
      offsetTop: 0,
      offsetWidth: 960,
      offsetHeight: 540,
      clientWidth: 960,
      clientHeight: 540,
      parentNode: null,
      parentElement: null,
      children: [],
      isConnected: true,
      classList: crearClaseList(),
      append(...cs) { cs.forEach((c) => { if (c && typeof c === "object") { c.parentNode = c.parentElement = this; this.children.push(c); } }); },
      remove() {
        if (this.parentNode) {
          const i = this.parentNode.children.indexOf(this);
          if (i >= 0) this.parentNode.children.splice(i, 1);
          this.parentNode = this.parentElement = null;
        }
      },
      contains(n) { let cur = n; while (cur) { if (cur === this) return true; cur = cur.parentNode; } return false; },
      focus() {},
      blur() {},
      setPointerCapture() {},
      insertAdjacentHTML() {},
      addEventListener(tipo, fn) { if (!ls.has(tipo)) ls.set(tipo, []); ls.get(tipo).push(fn); },
      removeEventListener(tipo, fn) { if (ls.has(tipo)) ls.set(tipo, ls.get(tipo).filter((x) => x !== fn)); },
      dispatchEvent(ev) {
        ev.target = ev.target || this;
        ev.currentTarget = this;
        ev.preventDefault = ev.preventDefault || (() => { ev.defaultPrevented = true; });
        ev.stopPropagation = ev.stopPropagation || (() => {});
        for (const fn of (ls.get(ev.type) || []).slice()) fn.call(this, ev);
      },
      querySelector: () => null,
      querySelectorAll: () => [],
      closest: () => null,
      getAttribute: (k) => (attrs.has(k) ? attrs.get(k) : null),
      setAttribute: (k, v) => attrs.set(k, String(v)),
      removeAttribute: (k) => attrs.delete(k),
      toggleAttribute: (k, f) => { if (f) attrs.set(k, ""); else attrs.delete(k); },
      getBoundingClientRect: () => ({ left: 0, top: 0, right: 960, bottom: 540, width: 960, height: 540 }),
    };
    return el;
  }

  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const mapaDom = new Map();
  let superficieEl = null, mundoEl = null;
  const html = (texto) => {
    const s = String(texto || "");
    const el = crearNodo();
    el._html = s;
    if (s.includes("data-minis")) {
      const minis = crearNodo("nav"), escena = crearNodo("div");
      mapaDom.set(el, { "[data-minis]": minis, "[data-escena]": escena });
    } else if (s.includes("data-escenario") || s.includes("data-plano")) {
      superficieEl = el;
      mundoEl = crearNodo("div");
      mundoEl.getBoundingClientRect = () => ({ left: 0, top: 0, right: 960, bottom: 540, width: 960, height: 540 });
      const svg = crearNodo("svg"), guias = crearNodo("svg"), marco = crearNodo("div");
      mapaDom.set(el, { "[data-mundo]": mundoEl, "[data-svg]": svg, "[data-guias]": guias, "[data-marco]": marco });
    } else if (s.includes("pres-notas")) {
      mapaDom.set(el, { textarea: crearNodo("textarea"), "[data-x]": crearNodo("button") });
    } else if (s.includes("data-zv")) {
      mapaDom.set(el, { "[data-zv]": crearNodo("span") });
    } else if (s.includes("data-rejilla")) {
      const rej = crearNodo("div"), fx = crearNodo("input"), refEl = crearNodo("span");
      const filtro = crearNodo("div"), graf = crearNodo("div"), esq = crearNodo("div"), band = crearNodo("div"), calc = crearNodo("button");
      mapaDom.set(el, { "[data-rejilla]": rej, "[data-fx]": fx, "[data-ref]": refEl, "[data-filtro]": filtro, "[data-graficas]": graf, "[data-esquema]": esq, "[data-bandeja]": band, "[data-calcular]": calc });
    } else if (s.includes('class="bl')) {
      const mId = /data-id="([^"]+)"/.exec(s);
      if (mId) el.dataset.id = mId[1];
      el._esBl = true;
      el.closest = (sel) => (sel === ".bl" ? el : null);
    } else if (s.includes('class="tirador"')) {
      const mT = /data-t="([^"]+)"/.exec(s);
      if (mT) el.dataset.t = mT[1];
      el._esTirador = true;
      el.closest = (sel) => (sel === ".tirador" ? el : sel === ".bl" ? el.parentNode : null);
    }
    return el;
  };
  const $ = (sel, ctx) => (mapaDom.has(ctx) ? mapaDom.get(ctx)[sel] || null : null);
  const $$ = (sel, ctx) => {
    if (sel === ".bl" && ctx) return ctx.children.filter((x) => x._esBl);
    return [];
  };
  const prefs = Object.assign({}, opciones.prefs || {});
  const NOM = {
    esc, ico: () => "", html, $, $$, aviso: () => {}, colorCss: (v, d = "") => v || d, colorHex: (v, d = "#000000") => v || d,
    modos: {}, rueda: { enlazar() {}, abierta: () => false }, cerrarMenu() {}, throttle: (fn) => fn, debounce: (fn) => fn,
    graficas: { TIPOS: [], PALETAS: {}, rango: () => null, svg: () => "" },
    leer: (k, d) => (k in prefs ? prefs[k] : d), guardar: (k, v) => { prefs[k] = v; },
    idCorto: () => "id1", DocCRDT: { nuevoId: () => "n1" },
    api: { clave: (k) => "u1." + k, usuario: () => ({ id: "u1" }) },
    binarios: { esRef: () => false }, srcBin: (v) => v,
  };
  const document = {
    body: crearNodo("body"),
    createElement: (tag) => (tag === "canvas" ? { getContext: () => ({ fillStyle: "#000", measureText: () => ({ width: 50 }) }) } : crearNodo(tag)),
    createElementNS: () => crearNodo("path"),
    addEventListener() {}, removeEventListener() {},
    querySelector: () => null,
    elementFromPoint: () => null,
    activeElement: null,
  };
  const window = {
    NOM, document,
    matchMedia: () => ({ matches: false }),
    addEventListener(tipo, fn) { if (!listenersWin.has(tipo)) listenersWin.set(tipo, []); listenersWin.get(tipo).push(fn); },
    removeEventListener(tipo, fn) { if (listenersWin.has(tipo)) listenersWin.set(tipo, listenersWin.get(tipo).filter((x) => x !== fn)); },
    innerWidth: 1200, innerHeight: 800,
  };
  const sandbox = {
    window, NOM, document,
    matchMedia: window.matchMedia,
    addEventListener: window.addEventListener,
    removeEventListener: window.removeEventListener,
    ResizeObserver: class { observe() {} disconnect() {} },
    MutationObserver: class { observe() {} disconnect() {} takeRecords() { return []; } },
    requestAnimationFrame: () => 0,
    CSS: { escape: (s) => String(s) },
    innerWidth: 1200, innerHeight: 800,
    console, Intl, setTimeout, clearTimeout,
  };
  const ctx = vm.createContext(sandbox);
  const cargar = (f) => vm.runInContext(leerJs(f), ctx, { filename: f });
  const getMundo = () => mundoEl;
  const getSuperficie = () => superficieEl;
  const dispatchWin = (tipo, ev) => { for (const fn of (listenersWin.get(tipo) || []).slice()) fn(ev); };
  return { ctx, sandbox, NOM, document, cargar, crearNodo, getMundo, getSuperficie, dispatchWin };
}

// ==================================================================
// 1. lienzo.js: mover y redimensionar contra orillas (diapositivas y diagramas)
// ==================================================================
{
  function montarDiapo(bloqueInicial, cuadricula) {
    const env = crearEntornoBasico({ prefs: { "nom.cuadricula": !!cuadricula } });
    env.cargar("diapositivas.js");
    env.cargar("lienzo.js");
    const bloques = { [bloqueInicial.id]: Object.assign({ pagina: "p1", z: 1 }, bloqueInicial) };
    const doc = {
      doc: () => ({}),
      paginas: () => ["p1"],
      valor: () => null,
      vivos: () => Object.values(bloques),
      deLaPagina: () => Object.values(bloques),
      bloque: (id) => bloques[id] || null,
      vivo: (id) => !!bloques[id],
      lote: (fn) => fn(),
      cambiar(id, campos) { Object.assign(bloques[id], campos); },
    };
    const ed = {
      doc, pagina: "p1", soloLectura: false, esPropietario: true,
      lienzo: env.crearNodo("div"), vista: env.crearNodo("div"),
      seleccionar() {}, pintarPresencia() {}, refrescarHerramientas() {},
      nombreDe: () => "Yo", colorDe: () => "#000",
    };
    const modo = Object.create(env.NOM.modos.PRESENTACION);
    modo.montar(ed);
    const sup = env.getSuperficie(), mundo = env.getMundo();
    // Escala 1:1 en la superficie para coordenadas de prueba exactas
    sup.getBoundingClientRect = () => ({ left: 0, top: 0, right: 1008, bottom: 588, width: 1008, height: 588 });
    mundo.getBoundingClientRect = () => ({ left: 0, top: 0, right: 960, bottom: 540, width: 960, height: 540 });
    mundo.querySelector = (sel) => {
      const m = /data-id="([^"]+)"/.exec(sel);
      return m ? mundo.children.find((x) => x.dataset.id === m[1]) || null : null;
    };
    // Recalcular escala=1
    modo._lienzo.render();
    return { env, modo, bloques, sup, mundo };
  }

  // 1a. Funciones puras de geometría en lienzo.js
  const envPuro = crearEntornoBasico();
  envPuro.cargar("diapositivas.js");
  envPuro.cargar("lienzo.js");
  const geo = envPuro.NOM.lienzoVista && envPuro.NOM.lienzoVista._geometria;
  prueba(
    "1a. lienzo.js expone geometría pura (_geometria.moverBloque y _geometria.calcularRedim)",
    !!(geo && typeof geo.moverBloque === "function" && typeof geo.calcularRedim === "function"),
    geo ? "incompleta" : "no existe NOM.lienzoVista._geometria"
  );

  if (geo && geo.moverBloque && geo.calcularRedim) {
    const b0 = { x: 700, y: 200, w: 200, h: 100, rot: 0 };
    // Al mover contra todas las orillas (incluso sacando parcialmente fuera), w y h jamás cambian
    const mDer = geo.moverBloque(b0, 220, 0);
    const mIzq = geo.moverBloque(b0, -760, 0);
    const mAba = geo.moverBloque(b0, 0, 320);
    const mArr = geo.moverBloque(b0, 0, -240);
    prueba(
      "1b. lienzo.js moverBloque contra las 4 orillas conserva w=200, h=100 y permite sacar un poco fuera",
      mDer.w === 200 && mDer.h === 100 && mDer.x === 920
        && mIzq.w === 200 && mIzq.h === 100 && mIzq.x === -60
        && mAba.w === 200 && mAba.h === 100 && mAba.y === 520
        && mArr.w === 200 && mArr.h === 100 && mArr.y === -40,
      JSON.stringify({ mDer, mIzq, mAba, mArr })
    );

    // Redimensionar jalando tirador 'w' pasando del borde derecho (x=900): borde derecho fijo en 900, h=100
    const rW = geo.calcularRedim(b0, "w", { x: 960, y: 250 }, { prop: false });
    const rWJusto = geo.calcularRedim(b0, "w", { x: 900, y: 250 }, { prop: false });
    prueba(
      "1c. lienzo.js calcularRedim tirador 'w' al/pasando el borde opuesto deja el borde derecho fijo en 900 y h=100",
      rW.x + rW.w === 900 && rW.w === 16 && rW.y === 200 && rW.h === 100
        && rWJusto.x + rWJusto.w === 900 && rWJusto.w === 16,
      JSON.stringify({ rW, rWJusto })
    );

    // Redimensionar jalando tirador 'e' pasando del borde izquierdo (x=700): borde izquierdo fijo en 700, h=100
    const rE = geo.calcularRedim(b0, "e", { x: 640, y: 250 }, { prop: false });
    prueba(
      "1d. lienzo.js calcularRedim tirador 'e' pasando el borde opuesto deja x=700 fijo y h=100",
      rE.x === 700 && rE.w === 16 && rE.y === 200 && rE.h === 100,
      JSON.stringify(rE)
    );

    // Redimensionar jalando tirador 'n' pasando del borde inferior (y=300): borde inferior fijo en 300, x=700, w=200
    const rN = geo.calcularRedim(b0, "n", { x: 800, y: 360 }, { prop: false });
    prueba(
      "1e. lienzo.js calcularRedim tirador 'n' pasando el borde opuesto deja el borde inferior fijo en 300 y w=200",
      rN.y + rN.h === 300 && rN.h === 12 && rN.x === 700 && rN.w === 200,
      JSON.stringify(rN)
    );

    // Con cuadrícula (paso=20): tirador 'w' cerca del borde derecho no corre el borde derecho de 900
    const rWGrid = geo.calcularRedim(b0, "w", { x: 895, y: 250 }, { paso: 20 });
    prueba(
      "1f. lienzo.js calcularRedim con cuadrícula en 'w' no desplaza el borde derecho fijo (900)",
      rWGrid.x + rWGrid.w === 900 && rWGrid.w === 20 && rWGrid.h === 100,
      JSON.stringify(rWGrid)
    );

    // Imagen (prop=true) con cuadrícula en esquina 'se': conserva relación 2:1 y esquina fija (700,200)
    const rPropGrid = geo.calcularRedim(b0, "se", { x: 942, y: 318 }, { prop: true, paso: 20 });
    const ratio0 = b0.w / b0.h, ratio1 = rPropGrid.w / rPropGrid.h;
    prueba(
      "1g. lienzo.js calcularRedim de imagen con cuadrícula en 'se' mantiene proporción 2:1 y ancla (700,200)",
      rPropGrid.x === 700 && rPropGrid.y === 200 && Math.abs(ratio1 - ratio0) < 0.02,
      JSON.stringify(rPropGrid)
    );
  }

  // 1h. Prueba integrada real de eventos pointer en lienzo.js (falla ANTES del arreglo en tiradores w/e y cuadrícula)
  {
    const s = montarDiapo({ id: "img1", tipo: "imagen", src: "data:image/png;base64,AA==", x: 700, y: 200, w: 200, h: 100 }, false);
    const blEl = s.mundo.children.find((x) => x.dataset.id === "img1");
    // Seleccionar para que cree tiradores
    s.sup.dispatchEvent({ type: "pointerdown", button: 0, pointerId: 1, pointerType: "mouse", clientX: 750, clientY: 250, target: blEl });
    s.sup.dispatchEvent({ type: "pointerup", button: 0, pointerId: 1, pointerType: "mouse", clientX: 750, clientY: 250, target: blEl });
    const blSel = s.mundo.children.find((x) => x.dataset.id === "img1");
    const tirW = blSel.children.find((x) => x.dataset.t === "w");
    // Arrastrar tirador 'w' justo hasta el borde derecho (x=900) y luego más allá (x=950)
    s.sup.dispatchEvent({ type: "pointerdown", button: 0, pointerId: 2, pointerType: "mouse", clientX: 700, clientY: 250, target: tirW });
    s.sup.dispatchEvent({ type: "pointermove", button: 0, pointerId: 2, pointerType: "mouse", clientX: 900, clientY: 250, target: tirW });
    const enBorde = Object.assign({}, { x: parseInt(blSel.style.left, 10), w: parseInt(blSel.style.width, 10), h: parseInt(blSel.style.height, 10) });
    s.sup.dispatchEvent({ type: "pointermove", button: 0, pointerId: 2, pointerType: "mouse", clientX: 950, clientY: 250, target: tirW });
    s.sup.dispatchEvent({ type: "pointerup", button: 0, pointerId: 2, pointerType: "mouse", clientX: 950, clientY: 250, target: tirW });
    const b = s.bloques.img1;
    prueba(
      "1h. lienzo.js interacción real: arrastrar tirador 'w' hasta/pasando borde derecho mantiene borde derecho en 900 y h=100",
      enBorde.x + enBorde.w === 900 && enBorde.h === 100 && b.x + b.w === 900 && b.w === 16 && b.h === 100,
      `enBorde=${JSON.stringify(enBorde)} final=${JSON.stringify({ x: b.x, y: b.y, w: b.w, h: b.h })}`
    );
  }
  {
    const s = montarDiapo({ id: "img2", tipo: "imagen", src: "data:image/png;base64,AA==", x: 700, y: 200, w: 200, h: 100 }, true);
    const blEl = s.mundo.children.find((x) => x.dataset.id === "img2");
    s.sup.dispatchEvent({ type: "pointerdown", button: 0, pointerId: 1, pointerType: "mouse", clientX: 750, clientY: 250, target: blEl });
    s.sup.dispatchEvent({ type: "pointerup", button: 0, pointerId: 1, pointerType: "mouse", clientX: 750, clientY: 250, target: blEl });
    const blSel = s.mundo.children.find((x) => x.dataset.id === "img2");
    const tirW = blSel.children.find((x) => x.dataset.t === "w");
    s.sup.dispatchEvent({ type: "pointerdown", button: 0, pointerId: 2, pointerType: "mouse", clientX: 700, clientY: 250, target: tirW });
    s.sup.dispatchEvent({ type: "pointermove", button: 0, pointerId: 2, pointerType: "mouse", clientX: 895, clientY: 250, target: tirW });
    s.sup.dispatchEvent({ type: "pointerup", button: 0, pointerId: 2, pointerType: "mouse", clientX: 895, clientY: 250, target: tirW });
    const b = s.bloques.img2;
    prueba(
      "1i. lienzo.js interacción real con cuadrícula: tirador 'w' cerca de 900 no corre el borde derecho",
      b.x + b.w === 900 && b.h === 100,
      `final=${JSON.stringify({ x: b.x, y: b.y, w: b.w, h: b.h })}`
    );
  }
}

// ==================================================================
// 2. modo-texto.js: imágenes flotantes en Word (mover y redimensionar)
// ==================================================================
{
  const env = crearEntornoBasico();
  env.cargar("modo-texto.js");
  const arr = env.NOM.modos.TEXTO && env.NOM.modos.TEXTO._arrastre;
  prueba(
    "2a. modo-texto.js expone geometría pura (_arrastre.moverCajaFlotante y _arrastre.redimImagenFlotante)",
    !!(arr && typeof arr.moverCajaFlotante === "function" && typeof arr.redimImagenFlotante === "function"),
    arr ? "incompleta" : "no existe NOM.modos.TEXTO._arrastre"
  );

  if (arr && arr.moverCajaFlotante && arr.redimImagenFlotante) {
    // Hoja de 816x1056, párrafo con margen izquierdo 96 y top 120
    const lim = { minX: -96, maxX: 816 - 96, minY: -120, maxY: 1056 - 120 };
    const tam = { w: 320, h: 200 };
    // Arrastrar fuerte contra borde derecho, izquierdo, inferior y superior: w y h NUNCA cambian y deja sacar un poco fuera
    const der = arr.moverCajaFlotante({ x: 100, y: 50 }, { dx: 1200, dy: 0 }, tam, lim, 32);
    const izq = arr.moverCajaFlotante({ x: 100, y: 50 }, { dx: -1200, dy: 0 }, tam, lim, 32);
    const aba = arr.moverCajaFlotante({ x: 100, y: 50 }, { dx: 0, dy: 2000 }, tam, lim, 32);
    const arrb = arr.moverCajaFlotante({ x: 100, y: 50 }, { dx: 0, dy: -2000 }, tam, lim, 32);
    prueba(
      "2b. modo-texto.js moverCajaFlotante contra las 4 orillas mantiene w=320, h=200 y acota posición con holgura visible",
      der.w === 320 && der.h === 200 && der.x === 816 - 96 - 32
        && izq.w === 320 && izq.h === 200 && izq.x === -96 - (320 - 32)
        && aba.w === 320 && aba.h === 200 && aba.y === 1056 - 120 - 32
        && arrb.w === 320 && arrb.h === 200 && arrb.y === -120 - (200 - 32),
      JSON.stringify({ der, izq, aba, arrb })
    );

    // Imagen más ancha que el cuerpo (900 > 816): no produce NaN ni inversión de límites
    const grande = arr.moverCajaFlotante({ x: 0, y: 0 }, { dx: 500, dy: 0 }, { w: 900, h: 200 }, lim, 32);
    prueba(
      "2c. modo-texto.js moverCajaFlotante con imagen más ancha que la hoja conserva w=900, h=200 sin invertirse",
      grande.w === 900 && grande.h === 200 && Number.isFinite(grande.x),
      JSON.stringify(grande)
    );

    // Redimensionar imagen flotante con alto fijo (Word .flota-w): escala h proporcionalmente sin deformar
    const rd = arr.redimImagenFlotante(300, 200, -150, 24, 816);
    prueba(
      "2d. modo-texto.js redimImagenFlotante escala alto proporcionalmente (300x200 -> 150x100)",
      rd.w === 150 && rd.h === 100,
      JSON.stringify(rd)
    );
  }

  // 2e. Comprobación real sobre el código de arrastrarImagen y llenarObjetos en modo-texto.js:
  // Simula el encogimiento CSS (max-width: 100% en position:absolute contra la orilla derecha del párrafo p de 624px)
  {
    const src = leerJs("modo-texto.js");
    // Extraer arrastrarImagen real y ejecutarla con un nodo que emula CSS max-width:100% cuando maxWidth !== 'none'
    const ini = src.indexOf("function arrastrarImagen(e, o) {");
    const fin = src.indexOf("/** Lleva una imagen en línea", ini);
    const fnArr = src.slice(ini, fin);
    const iniGeo = src.indexOf("function moverCajaFlotante(");
    const fnGeo = iniGeo >= 0 ? src.slice(iniGeo, src.indexOf("\n  const ", iniGeo) > 0 ? src.indexOf("\n  const ", iniGeo) : src.indexOf("function redimImagenFlotante", iniGeo)) : "";
    const tieneMaxWidthNoneLibre = /maxWidth\s*=[^;\n]*["']none["']/.test(src.slice(src.indexOf('if (e.tipo === "imagen")'), src.indexOf('else if (e.tipo === "tabla")')));
    const pWidth = 624, pLeft = 96, cuerpoW = 816, cuerpoH = 1056;
    const bl = { libre: true, x: 100, y: 50, w: 320, h: 200 };
    let guardado = null;
    const imgEl = { style: { width: "320px", maxWidth: tieneMaxWidthNoneLibre ? "none" : "" } };
    const cajaEl = { style: { maxWidth: tieneMaxWidthNoneLibre ? "none" : "" } };
    const o = {
      dataset: { id: "im1" },
      parentElement: { offsetLeft: pLeft, offsetTop: 120, offsetWidth: pWidth },
      style: { left: "100px", top: "50px", width: tieneMaxWidthNoneLibre ? "320px" : "", maxWidth: tieneMaxWidthNoneLibre ? "none" : "" },
      classList: { add() {}, remove() {}, contains: () => false },
      getBoundingClientRect: () => ({ left: pLeft + (parseFloat(o.style.left) || 0), top: 170, width: o.offsetWidth, height: o.offsetHeight }),
      get offsetWidth() {
        const w = parseFloat(imgEl.style.width) || 320;
        if (this.style.maxWidth === "none") return w;
        const left = parseFloat(this.style.left) || 0;
        return Math.max(0, Math.min(w, pWidth - left));
      },
      get offsetHeight() { return Math.round(this.offsetWidth * (200 / 320)); },
    };
    const listeners = new Map();
    const sbox = {
      Math, Object, parseFloat, parseInt,
      cuerpo: { clientWidth: cuerpoW, offsetHeight: cuerpoH },
      zoom: 1,
      $: (sel) => (sel === "img" ? imgEl : sel === ".wd-img-caja" ? cajaEl : null),
      sec: () => ({ item: () => ({ e: { ref: "bImg" } }) }),
      doc: () => ({ bloque: () => bl, valor: (_id, k) => bl[k], cambiar: (_id, c) => { guardado = c; } }),
      programarPaginar: () => {},
      addEventListener: (t, fn) => listeners.set(t, fn),
      removeEventListener: (t) => listeners.delete(t),
      moverCajaFlotante: arr && arr.moverCajaFlotante,
    };
    const c2 = vm.createContext(sbox);
    vm.runInContext(fnArr + "\narrastrarImagen({ pointerId: 1, clientX: 200, clientY: 200 }, __o);", Object.assign(c2, { __o: o }));
    const onMove = listeners.get("pointermove"), onUp = listeners.get("pointerup");
    // Arrastrar en 5 pasos hacia la orilla derecha (como hace el ratón del usuario)
    for (const dx of [250, 350, 450, 550, 650]) onMove({ pointerId: 1, clientX: 200 + dx, clientY: 200, preventDefault() {} });
    const anchoDuranteArrastre = o.offsetWidth;
    const altoDuranteArrastre = o.offsetHeight;
    onUp({ pointerId: 1, type: "pointerup" });
    prueba(
      "2e. modo-texto.js arrastrarImagen contra la orilla derecha NO encoge la imagen (mantiene 320x200)",
      tieneMaxWidthNoneLibre && anchoDuranteArrastre === 320 && altoDuranteArrastre === 200 && guardado && !("w" in guardado) && !("h" in guardado),
      `maxWidthNone=${tieneMaxWidthNoneLibre} tamDurante=${anchoDuranteArrastre}x${altoDuranteArrastre} guardado=${JSON.stringify(guardado)}`
    );
  }
}

// ==================================================================
// 3. modo-hoja.js: imágenes sobre la hoja (mover contra orillas)
// ==================================================================
{
  const env = crearEntornoBasico();
  env.cargar("formulas.js");
  env.cargar("modo-hoja.js");
  const arr = env.NOM.modos.HOJA && env.NOM.modos.HOJA._arrastre;
  prueba(
    "3a. modo-hoja.js expone geometría pura (_arrastre.posImagenHoja)",
    !!(arr && typeof arr.posImagenHoja === "function"),
    arr ? "incompleta" : "no existe NOM.modos.HOJA._arrastre"
  );

  if (arr && arr.posImagenHoja) {
    const cols = [0, 1, 2, 3, 4].map((c) => ({ c, x: c * 120, w: 120 })); // total 600 px
    const filas = [0, 1, 2, 3, 4].map((f) => ({ f, y: f * 32, h: 32 })); // total 160 px
    const tam = { w: 300, h: 180 };
    // Arrastrar pasado el borde izquierdo y superior -> acota en A1 (0,0) sin encoger w/h
    const pIzq = arr.posImagenHoja(-150, -80, tam, cols, filas);
    // Arrastrar pasado el borde derecho e inferior -> acota dejando parte visible, permite sacar un poco fuera, w/h intactos
    const pDer = arr.posImagenHoja(900, 400, tam, cols, filas);
    // Arrastrar dentro de una celda intermedia (C2 + offset) -> conserva offset y tamaño
    const pMed = arr.posImagenHoja(260, 45, tam, cols, filas);
    prueba(
      "3b. modo-hoja.js posImagenHoja contra las 4 orillas acota posición y mantiene w=300, h=180",
      pIzq.de.c === 0 && pIzq.de.f === 0 && pIzq.dx === 0 && pIzq.dy === 0 && pIzq.w === 300 && pIzq.h === 180
        && pDer.de.c === 4 && pDer.de.f === 4 && pDer.w === 300 && pDer.h === 180 && pDer.dx >= 0 && pDer.dy >= 0
        && pMed.de.c === 2 && pMed.de.f === 1 && pMed.dx === 20 && pMed.dy === 13 && pMed.w === 300 && pMed.h === 180,
      JSON.stringify({ pIzq, pDer, pMed })
    );
  }
}

console.log(fallas ? `${fallas} FALLA(S)` : "Todas las pruebas pasaron");
process.exit(fallas ? 1 : 0);

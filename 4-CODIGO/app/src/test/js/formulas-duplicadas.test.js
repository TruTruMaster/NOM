// Pruebas de regresión para "fórmulas que se duplican" en modo-hoja.js y relleno-rapido.js.
// Monta modo-hoja.js real con un DOM mínimo que respeta focus/blur/pointerdown/keydown/copy/cut/paste
// y un DocCRDT mínimo con pila de deshacer/rehacer y registro de operaciones.
// Uso: node 4-CODIGO/app/src/test/js/formulas-duplicadas.test.js
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const dirJs = process.env.NOM_JS_DIR || path.join(__dirname, "../../main/resources/static/js");
const leerJs = (n) => fs.readFileSync(path.join(dirJs, n), "utf8");

function crearHojaPrueba(celdasIniciales) {
  let activeElement = null;
  const listenersDoc = new Map();

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
      style: {},
      dataset: {},
      hidden: false,
      value: "",
      placeholder: "",
      textContent: "",
      title: "",
      scrollLeft: 0,
      scrollTop: 0,
      offsetLeft: 0,
      offsetTop: 0,
      offsetWidth: 80,
      offsetHeight: 24,
      selectionStart: 0,
      selectionEnd: 0,
      parentNode: null,
      children: [],
      classList: crearClaseList(),
      _html: "",
      _tds: new Map(),
      get innerHTML() { return this._html; },
      set innerHTML(v) {
        this._html = String(v || "");
        this._tds.clear();
        const re = /<td\s+data-c="(\d+)"\s+data-f="(\d+)"[^>]*>/g;
        let m;
        while ((m = re.exec(this._html))) {
          const td = crearNodo("td");
          td.dataset.c = m[1];
          td.dataset.f = m[2];
          td.parentNode = el;
          this._tds.set(m[1] + "," + m[2], td);
        }
      },
      append(c) {
        if (c && typeof c === "object") { c.parentNode = this; this.children.push(c); }
      },
      remove() {
        if (this.parentNode) {
          const i = this.parentNode.children.indexOf(this);
          if (i >= 0) this.parentNode.children.splice(i, 1);
          this.parentNode = null;
        }
        if (activeElement === this) {
          activeElement = null;
          this.dispatchEvent({ type: "blur" });
        }
      },
      contains(n) {
        let cur = n;
        while (cur) { if (cur === this) return true; cur = cur.parentNode; }
        return false;
      },
      focus() {
        if (activeElement === this) return;
        const prev = activeElement;
        activeElement = this;
        if (prev) prev.dispatchEvent({ type: "blur" });
        this.dispatchEvent({ type: "focus" });
      },
      blur() {
        if (activeElement === this) {
          activeElement = null;
          this.dispatchEvent({ type: "blur" });
        }
      },
      setSelectionRange(a, b) { this.selectionStart = a; this.selectionEnd = b; },
      scrollIntoView() {},
      setPointerCapture() {},
      addEventListener(tipo, fn, cap) {
        if (!ls.has(tipo)) ls.set(tipo, []);
        ls.get(tipo).push({ fn, cap: !!cap });
      },
      removeEventListener(tipo, fn) {
        if (!ls.has(tipo)) return;
        ls.set(tipo, ls.get(tipo).filter((x) => x.fn !== fn));
      },
      dispatchEvent(ev) {
        ev.target = ev.target || this;
        ev.currentTarget = this;
        ev.preventDefault = ev.preventDefault || (() => { ev.defaultPrevented = true; });
        ev.stopPropagation = ev.stopPropagation || (() => { ev._stop = true; });
        ev.stopImmediatePropagation = ev.stopImmediatePropagation || (() => { ev._stopImm = true; ev._stop = true; });
        const lista = (ls.get(ev.type) || []).slice();
        for (const item of lista) {
          item.fn.call(this, ev);
          if (ev._stopImm) break;
        }
        return !ev.defaultPrevented;
      },
      querySelector(sel) {
        const m = /^td\[data-c="(\d+)"\]\[data-f="(\d+)"\]$/.exec(sel);
        if (m) return this._tds.get(m[1] + "," + m[2]) || null;
        return null;
      },
      querySelectorAll() { return []; },
      closest(sel) {
        if (sel === "td" && this.tagName === "TD") return this;
        return null;
      },
      getAttribute(k) { return attrs.has(k) ? attrs.get(k) : null; },
      setAttribute(k, v) { attrs.set(k, String(v)); },
      removeAttribute(k) { attrs.delete(k); },
      toggleAttribute(k, f) { if (f) attrs.set(k, ""); else attrs.delete(k); },
      getBoundingClientRect() { return { left: 0, top: 0, right: 80, bottom: 24, width: 80, height: 24 }; },
    };
    return el;
  }

  const mapaDom = new Map();
  let rej, fx, refEl;
  const raizHtml = (texto) => {
    const tag = /^<(\w+)/.exec(String(texto || "").trim());
    const el = crearNodo(tag ? tag[1] : "div");
    el._html = texto;
    if (String(texto).includes("data-rejilla")) {
      rej = crearNodo("div");
      fx = crearNodo("input");
      refEl = crearNodo("span");
      const filtro = crearNodo("div"), graf = crearNodo("div"), esq = crearNodo("div"), band = crearNodo("div"), calc = crearNodo("button");
      mapaDom.set(el, {
        "[data-rejilla]": rej, "[data-fx]": fx, "[data-ref]": refEl,
        "[data-filtro]": filtro, "[data-graficas]": graf, "[data-esquema]": esq,
        "[data-bandeja]": band, "[data-calcular]": calc,
      });
    }
    return el;
  };
  const $ = (sel, ctx) => (mapaDom.has(ctx) ? mapaDom.get(ctx)[sel] || null : null);
  const $$ = (sel, ctx) => {
    if (sel === "td" && ctx && ctx._tds) return [...ctx._tds.values()];
    return [];
  };

  const document = {
    body: crearNodo("body"),
    createElement: (tag) => (tag === "canvas" ? { getContext: () => ({ fillStyle: "#000", font: "", measureText: () => ({ width: 50 }) }) } : crearNodo(tag)),
    addEventListener(tipo, fn) { if (!listenersDoc.has(tipo)) listenersDoc.set(tipo, []); listenersDoc.get(tipo).push(fn); },
    removeEventListener(tipo, fn) { if (listenersDoc.has(tipo)) listenersDoc.set(tipo, listenersDoc.get(tipo).filter((x) => x !== fn)); },
    querySelector: () => null,
    get activeElement() { return activeElement; },
  };
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const NOM = {
    esc, ico: () => "", html: raizHtml, $, $$, aviso: () => {}, colorCss: (v, d = "") => v || d,
    modos: {}, rueda: { enlazar() {} }, graficas: { TIPOS: [], rango: () => null },
    leer: (k, d) => d, guardar: () => {}, idCorto: () => "id1",
    api: { clave: (k) => "u1." + k, usuario: () => ({ id: "u1" }) },
    binarios: { esRef: () => false }, srcBin: (v) => v,
  };
  const window = { NOM, document, matchMedia: () => ({ matches: false }), addEventListener() {}, removeEventListener() {}, innerWidth: 1200, innerHeight: 800 };
  const timeouts = [];
  const sandbox = {
    window, NOM, document, matchMedia: window.matchMedia, console, Intl,
    innerWidth: 1200, innerHeight: 800,
    Event: function Event(type) { this.type = type; },
    setTimeout: (fn) => { timeouts.push(fn); return timeouts.length; },
    clearTimeout: () => {},
  };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(leerJs("formulas.js"), ctx, { filename: "formulas.js" });
  vm.runInContext(leerJs("relleno-rapido.js"), ctx, { filename: "relleno-rapido.js" });
  vm.runInContext(leerJs("modo-hoja.js"), ctx, { filename: "modo-hoja.js" });

  // Documento CRDT mínimo con historial de operaciones y deshacer/rehacer
  const tabla = Object.assign({ id: "hoja-p1", tipo: "tabla", pagina: "p1", filas: 10, cols: 8, nombre: "Hoja1" }, celdasIniciales || {});
  const ops = [];
  const pilaDeshacer = [];
  const pilaRehacer = [];
  let modoInst = null;
  const doc = {
    deLaPagina: () => [tabla],
    vivos: () => [tabla],
    paginas: () => ["p1"],
    bloque: (id) => (id === tabla.id ? tabla : null),
    valor: (id, k) => (id === tabla.id ? tabla[k] : undefined),
    comentarios: () => [],
    lote(fn) { fn(); },
    cambiar(id, campos, opc) {
      if (id !== tabla.id) return;
      const previos = {};
      for (const k of Object.keys(campos)) previos[k] = tabla[k] === undefined ? null : tabla[k];
      Object.assign(tabla, campos);
      ops.push({ id, campos: Object.assign({}, campos) });
      if (!opc || opc.deshacible !== false) {
        pilaDeshacer.push(previos);
        pilaRehacer.length = 0;
      }
      if (modoInst && modoInst.cambios) modoInst.cambios([{ bloque: id, campos, local: true }]);
    },
    deshacer() {
      const prev = pilaDeshacer.pop();
      if (!prev) return false;
      const actual = {};
      for (const k of Object.keys(prev)) actual[k] = tabla[k] === undefined ? null : tabla[k];
      pilaRehacer.push(actual);
      Object.assign(tabla, prev);
      if (modoInst && modoInst.cambios) modoInst.cambios([{ bloque: tabla.id, campos: prev, local: true }]);
      return true;
    },
    rehacer() {
      const sig = pilaRehacer.pop();
      if (!sig) return false;
      const actual = {};
      for (const k of Object.keys(sig)) actual[k] = tabla[k] === undefined ? null : tabla[k];
      pilaDeshacer.push(actual);
      Object.assign(tabla, sig);
      if (modoInst && modoInst.cambios) modoInst.cambios([{ bloque: tabla.id, campos: sig, local: true }]);
      return true;
    },
  };
  const ed = {
    doc, pagina: "p1", soloLectura: false, lienzo: crearNodo("div"), vista: crearNodo("div"),
    seleccionar() {}, pintarPresencia() {},
  };
  modoInst = Object.create(NOM.modos.HOJA);
  modoInst.montar(ed);

  const vaciarTimeouts = () => { while (timeouts.length) timeouts.shift()(); };
  const clicCelda = (c, f, extra) => {
    const td = rej.querySelector(`td[data-c="${c}"][data-f="${f}"]`);
    rej.dispatchEvent(Object.assign({ type: "pointerdown", target: td, pointerType: "mouse", pointerId: 1, shiftKey: false }, extra || {}));
    (listenersDoc.get("pointerup") || []).forEach((fn) => fn({ pointerId: 1 }));
    vaciarTimeouts();
  };
  const celda = (ref) => tabla["c:" + ref] || "";
  return { tabla, ops, rej, fx, doc, modoInst, clicCelda, celda, vaciarTimeouts, getActive: () => activeElement };
}

let fallas = 0;
const prueba = (nombre, ok, detalle) => {
  console.log((ok ? "OK    " : "FALLA ") + nombre + (!ok && detalle ? " — " + detalle : ""));
  if (!ok) fallas++;
};

// Caso 1: Escribir fórmula en la barra fx y pulsar Enter NO debe duplicarla en la celda de abajo ni enviar 2 ops CRDT
{
  const h = crearHojaPrueba({ "c:A1": "10", "c:B1": "20" });
  h.clicCelda(2, 0); // C1
  h.fx.focus();
  h.fx.value = "=A1+B1";
  h.fx.dispatchEvent({ type: "input" });
  h.fx.dispatchEvent({ type: "keydown", key: "Enter" });
  h.vaciarTimeouts();
  prueba(
    "1. Barra fx + Enter escribe la fórmula solo en C1 (no la duplica en C2) y envía 1 operación CRDT",
    h.celda("C1") === "=A1+B1" && h.celda("C2") === "" && h.ops.length === 1,
    `C1=${JSON.stringify(h.celda("C1"))} C2=${JSON.stringify(h.celda("C2"))} ops=${h.ops.length}`
  );
}

// Caso 2: Seleccionar celda con fórmula, enfocar barra fx sin editar y hacer clic en otra celda NO debe copiar la fórmula
{
  const h = crearHojaPrueba({ "c:A1": "10", "c:B1": "20", "c:C1": "=A1+B1" });
  h.clicCelda(2, 0); // C1
  h.fx.focus();
  h.clicCelda(3, 4); // D5
  prueba(
    "2. Enfocar barra fx en C1 (=A1+B1) y clicar D5 no duplica la fórmula en D5",
    h.celda("C1") === "=A1+B1" && h.celda("D5") === "" && h.ops.length === 0,
    `C1=${JSON.stringify(h.celda("C1"))} D5=${JSON.stringify(h.celda("D5"))} ops=${h.ops.length}`
  );
}

// Caso 3: Editar fórmula en barra fx de C1 y hacer clic en D5 guarda en C1, no en D5
{
  const h = crearHojaPrueba({ "c:A1": "10", "c:B1": "20" });
  h.clicCelda(2, 0); // C1
  h.fx.focus();
  h.fx.value = "=A1*3";
  h.fx.setSelectionRange(5, 5);
  h.fx.dispatchEvent({ type: "input" });
  h.clicCelda(3, 4); // D5
  prueba(
    "3. Editar fórmula en barra fx en C1 y clicar D5 guarda en C1 y deja D5 vacía",
    h.celda("C1") === "=A1*3" && h.celda("D5") === "" && h.ops.length === 1,
    `C1=${JSON.stringify(h.celda("C1"))} D5=${JSON.stringify(h.celda("D5"))} ops=${h.ops.length}`
  );
}

// Caso 4: Escribir '=SUMA(' en la barra fx y clicar A1 inserta la referencia en fx y no sobrescribe A1
{
  const h = crearHojaPrueba({ "c:A1": "10" });
  h.clicCelda(2, 0); // C1
  h.fx.focus();
  h.fx.value = "=SUMA(";
  h.fx.setSelectionRange(6, 6);
  h.fx.dispatchEvent({ type: "input" });
  h.clicCelda(0, 0); // A1
  prueba(
    "4. Escribir '=SUMA(' en fx y clicar A1 inserta 'A1' en fx sin sobrescribir celda A1",
    h.fx.value === "=SUMA(A1" && h.celda("A1") === "10" && h.ops.length === 0,
    `fx=${JSON.stringify(h.fx.value)} A1=${JSON.stringify(h.celda("A1"))} ops=${h.ops.length}`
  );
}

// Caso 5: Empezar a escribir fórmula en el editor en celda, clicar fx para terminarla y pulsar Enter envía 1 sola op sin duplicar en C2
{
  const h = crearHojaPrueba({ "c:A1": "5", "c:B1": "7" });
  h.clicCelda(2, 0); // C1
  h.rej.dispatchEvent({ type: "keydown", key: "=" });
  const edCelda = h.getActive();
  edCelda.value = "=A1+B1";
  edCelda.dispatchEvent({ type: "input" });
  h.fx.focus();
  h.vaciarTimeouts();
  h.fx.value = "=A1+B1+1";
  h.fx.dispatchEvent({ type: "input" });
  h.fx.dispatchEvent({ type: "keydown", key: "Enter" });
  h.vaciarTimeouts();
  prueba(
    "5. Pasar del editor en celda a barra fx y pulsar Enter guarda 1 vez en C1 y no duplica en C2",
    h.celda("C1") === "=A1+B1+1" && h.celda("C2") === "" && h.ops.length === 1,
    `C1=${JSON.stringify(h.celda("C1"))} C2=${JSON.stringify(h.celda("C2"))} ops=${h.ops.length}`
  );
}

// Caso 6: Cortar (Ctrl+X) fórmula de C1 y pegar en rango C2:C4 mueve una sola vez a C2 (no duplica en C3 y C4)
{
  const h = crearHojaPrueba({ "c:A1": "10", "c:B1": "20", "c:C1": "=A1+B1" });
  let clip = "";
  h.clicCelda(2, 0); // C1
  h.rej.dispatchEvent({ type: "cut", clipboardData: { setData: (_t, v) => { clip = v; } } });
  h.clicCelda(2, 1); // C2
  h.clicCelda(2, 3, { shiftKey: true }); // C2:C4
  h.rej.dispatchEvent({ type: "paste", clipboardData: { getData: () => clip } });
  prueba(
    "6. Cortar C1 (=A1+B1) y pegar en rango C2:C4 mueve solo a C2 sin duplicar en C3 ni C4",
    h.celda("C1") === "" && h.celda("C2") === "=A1+B1" && h.celda("C3") === "" && h.celda("C4") === "",
    `C1=${JSON.stringify(h.celda("C1"))} C2=${JSON.stringify(h.celda("C2"))} C3=${JSON.stringify(h.celda("C3"))} C4=${JSON.stringify(h.celda("C4"))}`
  );
}

// Caso 7: Copiar fórmula, pegar, deshacer y rehacer deja una sola fórmula intacta y sin repetir texto
{
  const h = crearHojaPrueba({ "c:A1": "10", "c:B1": "20", "c:A2": "3", "c:B2": "4", "c:C1": "=A1+B1" });
  let clip = "";
  h.clicCelda(2, 0);
  h.rej.dispatchEvent({ type: "copy", clipboardData: { setData: (_t, v) => { clip = v; } } });
  h.clicCelda(2, 1);
  h.rej.dispatchEvent({ type: "paste", clipboardData: { getData: () => clip } });
  const trasPegar = h.celda("C2");
  h.doc.deshacer();
  const trasDeshacer = h.celda("C2");
  h.doc.rehacer();
  const trasRehacer = h.celda("C2");
  prueba(
    "7. Copiar/pegar/deshacer/rehacer fórmula desplaza bien y no duplica texto",
    trasPegar === "=A2+B2" && trasDeshacer === "" && trasRehacer === "=A2+B2",
    `pegar=${JSON.stringify(trasPegar)} deshacer=${JSON.stringify(trasDeshacer)} rehacer=${JSON.stringify(trasRehacer)}`
  );
}

// Caso 8: Relleno rápido (Ctrl+E) junto a fórmulas no las sobrescribe ni las duplica como fórmulas
{
  const h = crearHojaPrueba({
    "c:A1": "Ana Lopez", "c:B1": "Ana",
    "c:A2": "Luis Perez", "c:B2": "=IZQUIERDA(A2;4)",
    "c:A3": "Sara Diaz",
  });
  h.clicCelda(1, 0); // B1
  h.modoInst.api.herramientas.rellenoRapido();
  prueba(
    "8. Relleno rápido conserva la fórmula de B2 y solo llena B3 con texto",
    h.celda("B1") === "Ana" && h.celda("B2") === "=IZQUIERDA(A2;4)" && h.celda("B3") === "Sara",
    `B1=${JSON.stringify(h.celda("B1"))} B2=${JSON.stringify(h.celda("B2"))} B3=${JSON.stringify(h.celda("B3"))}`
  );
}

console.log(fallas ? `${fallas} FALLA(S)` : "Todas las pruebas pasaron (8/8)");
process.exit(fallas ? 1 : 0);

// Pruebas de preferencias-extra.js, experiencia.js (atajos personalizables),
// rueda.js (háptica y lista clásica) y Ajustes en vista-espacios.js.
// Uso: node --test 4-CODIGO/app/src/test/js/preferencias-extra.test.js
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const dirStatic = path.join(__dirname, "../../main/resources/static");
const dirJs = path.join(dirStatic, "js");

// ---------------------------------------------------------------- DOM mínimo en memoria
// Permite correr los módulos reales de static/js en Node sin dependencias externas,
// con eventos en fase de captura y burbuja, querySelector/querySelectorAll y focus.
function crearDom() {
  const listeners = new WeakMap();
  const desEsc = (s) => String(s || "")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&");

  class EventTargetMin {
    addEventListener(tipo, fn, opc) {
      const cap = typeof opc === "boolean" ? opc : !!(opc && opc.capture);
      const once = !!(opc && typeof opc === "object" && opc.once);
      let m = listeners.get(this);
      if (!m) { m = new Map(); listeners.set(this, m); }
      if (!m.has(tipo)) m.set(tipo, []);
      m.get(tipo).push({ fn, cap, once });
    }
    removeEventListener(tipo, fn, opc) {
      const cap = typeof opc === "boolean" ? opc : !!(opc && opc.capture);
      const m = listeners.get(this);
      if (!m || !m.has(tipo)) return;
      m.set(tipo, m.get(tipo).filter((x) => x.fn !== fn || x.cap !== cap));
    }
    _oir(tipo, cap) {
      const m = listeners.get(this);
      return m && m.has(tipo) ? m.get(tipo).filter((x) => x.cap === cap).slice() : [];
    }
  }

  class Nodo extends EventTargetMin {
    constructor(tag, doc) {
      super();
      this.nodeType = 1;
      this.tagName = tag.toUpperCase();
      this.ownerDocument = doc;
      this.attrs = new Map();
      this.children = [];
      this.parentNode = null;
      this.style = { setProperty() {} };
      this._text = "";
      this._html = "";
      this.value = "";
      this.checked = false;
      this.disabled = false;
      this.hidden = false;
      this.tabIndex = 0;
      this.offsetWidth = 220;
      this.offsetHeight = 160;
      this.offsetParent = {};
      const self = this;
      this.dataset = new Proxy({}, {
        get(_, k) {
          const a = "data-" + String(k).replace(/[A-Z]/g, (c) => "-" + c.toLowerCase());
          return self.attrs.has(a) ? self.attrs.get(a) : undefined;
        },
        set(_, k, v) {
          const a = "data-" + String(k).replace(/[A-Z]/g, (c) => "-" + c.toLowerCase());
          self.attrs.set(a, String(v));
          return true;
        },
        deleteProperty(_, k) {
          const a = "data-" + String(k).replace(/[A-Z]/g, (c) => "-" + c.toLowerCase());
          self.attrs.delete(a);
          return true;
        },
      });
      this.classList = {
        _lista: () => String(self.attrs.get("class") || "").trim().split(/\s+/).filter(Boolean),
        _guardar: (xs) => self.attrs.set("class", xs.join(" ")),
        add(...cs) { const s = new Set(this._lista()); cs.forEach((c) => c && s.add(c)); this._guardar([...s]); },
        remove(...cs) { const s = new Set(this._lista()); cs.forEach((c) => s.delete(c)); this._guardar([...s]); },
        toggle(c, f) { const on = f === undefined ? !this.contains(c) : !!f; if (on) this.add(c); else this.remove(c); return on; },
        contains(c) { return this._lista().includes(c); },
      };
    }
    get id() { return this.attrs.get("id") || ""; }
    set id(v) { this.attrs.set("id", String(v)); }
    get className() { return this.attrs.get("class") || ""; }
    set className(v) { this.attrs.set("class", String(v)); }
    get isConnected() {
      let n = this;
      while (n) { if (n === this.ownerDocument.body || n === this.ownerDocument.documentElement) return true; n = n.parentNode; }
      return false;
    }
    get firstElementChild() { return this.children[0] || null; }
    get textContent() {
      if (!this.children.length) return this._text;
      return this._text + this.children.map((c) => c.textContent).join("");
    }
    set textContent(v) { this.children = []; this._text = String(v == null ? "" : v); }
    get innerHTML() { return this._html; }
    set innerHTML(s) {
      this._html = String(s || "");
      this.children = [];
      this._text = "";
      parsearFragmento(this._html, this, this.ownerDocument);
    }
    getAttribute(k) { return this.attrs.has(k) ? this.attrs.get(k) : null; }
    setAttribute(k, v) {
      const s = String(v);
      this.attrs.set(k, s);
      if (k === "value") this.value = s;
      if (k === "checked") this.checked = true;
      if (k === "disabled") this.disabled = true;
      if (k === "hidden") this.hidden = true;
    }
    removeAttribute(k) {
      this.attrs.delete(k);
      if (k === "checked") this.checked = false;
      if (k === "disabled") this.disabled = false;
      if (k === "hidden") this.hidden = false;
    }
    hasAttribute(k) { return this.attrs.has(k); }
    append(...ns) {
      for (const n of ns) {
        if (!n) continue;
        if (typeof n === "string") { this._text += n; continue; }
        if (n.parentNode) n.remove();
        n.parentNode = this;
        this.children.push(n);
      }
    }
    appendChild(n) { this.append(n); return n; }
    after(n) {
      if (!this.parentNode || !n) return;
      const i = this.parentNode.children.indexOf(this);
      n.parentNode = this.parentNode;
      this.parentNode.children.splice(i + 1, 0, n);
    }
    remove() {
      if (!this.parentNode) return;
      const i = this.parentNode.children.indexOf(this);
      if (i >= 0) this.parentNode.children.splice(i, 1);
      this.parentNode = null;
    }
    contains(otro) {
      let n = otro;
      while (n) { if (n === this) return true; n = n.parentNode; }
      return false;
    }
    matches(sel) { return sel.split(",").some((p) => coincideSelector(this, p.trim())); }
    closest(sel) {
      let n = this;
      while (n && n.nodeType === 1) { if (n.matches(sel)) return n; n = n.parentNode; }
      return null;
    }
    querySelectorAll(sel) {
      const out = [];
      const rec = (n) => {
        for (const c of n.children) {
          if (c.matches(sel)) out.push(c);
          rec(c);
        }
      };
      rec(this);
      return out;
    }
    querySelector(sel) { return this.querySelectorAll(sel)[0] || null; }
    focus() { this.ownerDocument.activeElement = this; }
    blur() { if (this.ownerDocument.activeElement === this) this.ownerDocument.activeElement = this.ownerDocument.body; }
    scrollIntoView() {}
    getBoundingClientRect() { return { left: 100, top: 100, right: 200, bottom: 140, width: 100, height: 40 }; }
    click() {
      if (this.disabled) return;
      const ev = crearEvento("click", { target: this });
      despachar(this, ev);
    }
  }

  function coincideSimple(el, s) {
    if (!s || s === "*") return true;
    const partes = s.match(/([a-zA-Z0-9_-]+|\.[a-zA-Z0-9_-]+|#[a-zA-Z0-9_-]+|\[[^\]]+\])/g) || [];
    for (const p of partes) {
      if (p[0] === ".") { if (!el.classList.contains(p.slice(1))) return false; }
      else if (p[0] === "#") { if (el.id !== p.slice(1)) return false; }
      else if (p[0] === "[") {
        const m = /^\[([^=\]]+)(?:=(?:"([^"]*)"|'([^']*)'|([^\]]*)))?\]$/.exec(p);
        if (!m) return false;
        const k = m[1].trim(), val = m[2] ?? m[3] ?? m[4];
        if (!el.attrs.has(k)) return false;
        if (val !== undefined && el.attrs.get(k) !== val) return false;
      } else {
        if (el.tagName !== p.toUpperCase()) return false;
      }
    }
    return partes.length > 0;
  }

  function coincideSelector(el, sel) {
    const tokens = sel.replace(/\s*>\s*/g, " > ").trim().split(/\s+/);
    let idx = tokens.length - 1, cur = el;
    if (!coincideSimple(cur, tokens[idx--])) return false;
    while (idx >= 0) {
      if (tokens[idx] === ">") {
        idx--;
        cur = cur.parentNode;
        if (!cur || cur.nodeType !== 1 || !coincideSimple(cur, tokens[idx])) return false;
        idx--;
      } else {
        const buscado = tokens[idx--];
        let p = cur.parentNode, ok = false;
        while (p && p.nodeType === 1) { if (coincideSimple(p, buscado)) { cur = p; ok = true; break; } p = p.parentNode; }
        if (!ok) return false;
      }
    }
    return true;
  }

  const VACIOS = new Set(["input", "br", "hr", "img", "meta", "link", "use", "path", "circle", "rect", "line"]);
  function parsearFragmento(cadena, padre, doc) {
    const pila = [padre];
    const re = /<!--[\s\S]*?-->|<\/([a-zA-Z0-9-]+)\s*>|<([a-zA-Z0-9-]+)([^>]*?)(\/?)>|([^<]+)/g;
    let m;
    while ((m = re.exec(cadena))) {
      const tope = pila[pila.length - 1];
      if (m[1]) {
        if (pila.length > 1) pila.pop();
      } else if (m[2]) {
        const tag = m[2].toLowerCase(), attrStr = m[3] || "", auto = !!m[4] || VACIOS.has(tag);
        const el = new Nodo(tag, doc);
        const reA = /([a-zA-Z0-9_:-]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;
        let a;
        while ((a = reA.exec(attrStr))) {
          const k = a[1], v = a[2] ?? a[3] ?? a[4] ?? "";
          el.setAttribute(k, desEsc(v));
        }
        tope.append(el);
        if (!auto) pila.push(el);
      } else if (m[5]) {
        tope._text += desEsc(m[5]);
      }
    }
  }

  function crearEvento(tipo, props) {
    const ev = Object.assign({
      type: tipo,
      key: "",
      code: "",
      ctrlKey: false,
      metaKey: false,
      altKey: false,
      shiftKey: false,
      clientX: 0,
      clientY: 0,
      button: 0,
      pointerType: "mouse",
      cancelable: true,
      defaultPrevented: false,
      _parado: false,
      _inmediato: false,
      target: null,
      currentTarget: null,
      preventDefault() { this.defaultPrevented = true; },
      stopPropagation() { this._parado = true; },
      stopImmediatePropagation() { this._parado = true; this._inmediato = true; },
      getModifierState: () => false,
    }, props);
    return ev;
  }

  const doc = new EventTargetMin();
  const win = new EventTargetMin();
  doc.documentElement = new Nodo("html", doc);
  doc.body = new Nodo("body", doc);
  doc.documentElement.append(doc.body);
  doc.activeElement = doc.body;
  doc.createElement = (tag) => new Nodo(tag, doc);
  doc.querySelector = (s) => doc.documentElement.querySelector(s);
  doc.querySelectorAll = (s) => doc.documentElement.querySelectorAll(s);
  doc.dispatchEvent = (ev) => despachar(doc, ev);

  function despachar(destino, ev) {
    ev.target = ev.target || destino;
    const camino = [];
    let n = destino;
    while (n && n.nodeType === 1) { camino.push(n); n = n.parentNode; }
    const ruta = [win, doc, ...camino.reverse()];
    // Captura: window -> document -> ... -> target
    for (const nodo of ruta) {
      if (ev._parado) break;
      ev.currentTarget = nodo;
      for (const h of nodo._oir(ev.type, true)) {
        if (h.once) nodo.removeEventListener(ev.type, h.fn, true);
        h.fn.call(nodo, ev);
        if (ev._inmediato) break;
      }
    }
    // Propiedad directa (onclick / onchange / oninput) y burbuja: target -> ... -> document -> window
    if (!ev._parado) {
      for (const nodo of ruta.slice().reverse()) {
        if (ev._parado) break;
        ev.currentTarget = nodo;
        const prop = "on" + ev.type;
        if (typeof nodo[prop] === "function") nodo[prop].call(nodo, ev);
        if (ev._inmediato) break;
        for (const h of nodo._oir(ev.type, false)) {
          if (h.once) nodo.removeEventListener(ev.type, h.fn, false);
          h.fn.call(nodo, ev);
          if (ev._inmediato) break;
        }
      }
    }
    return !ev.defaultPrevented;
  }

  return { win, doc, Nodo, crearEvento, despachar };
}

function crearEntorno(opc) {
  opc = opc || {};
  const { win, doc, crearEvento, despachar } = crearDom();
  const alm = new Map();
  const localStorage = {
    getItem: (k) => (alm.has(String(k)) ? alm.get(String(k)) : null),
    setItem: (k, v) => { alm.set(String(k), String(v)); },
    removeItem: (k) => { alm.delete(String(k)); },
    clear: () => alm.clear(),
  };

  const avisos = [];
  const puts = [];
  let hojaActual = null;
  const esc = (s) => String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  const html = (cadena) => {
    const cont = doc.createElement("div");
    cont.innerHTML = cadena.trim();
    return cont.firstElementChild;
  };
  const $ = (s, r) => (r || doc).querySelector(s);
  const $$ = (s, r) => (r || doc).querySelectorAll(s);

  const NOM = {
    VERSION: "0.23.10",
    esc,
    ico: (n) => `<svg data-ico="${n}"></svg>`,
    html,
    $,
    $$,
    colorHex: (c, d = "#5A5D6B") => c || d,
    colorCss: (c, d = "") => c || d,
    avatar: (u) => `<span class="avatar">${esc((u && u.nombre) || "?")}</span>`,
    sinContrasena: () => false,
    etiquetaEquipo: () => "Guardado en este equipo",
    haceCuanto: () => "hace un momento",
    bytes: (n) => n + " B",
    MODOS: { TEXTO: { nombre: "Texto", icono: "file-text" } },
    leer: (k, def) => { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); },
    guardar: (k, v) => { localStorage.setItem(k, JSON.stringify(v)); return true; },
    quitar: (k) => localStorage.removeItem(k),
    aviso: (t, o) => avisos.push({ texto: t, opc: o }),
    cerrarMenu: () => {},
    aplicarLetra: (v) => NOM.guardar("nom.letra", v),
    aplicarPreferencias: (p) => NOM.guardar("nom.pref", p),
    hayHoja: () => !!hojaActual,
    hoja(o) {
      if (hojaActual) hojaActual.cerrar(true);
      const velo = html(`<div class="velo"><section class="hoja"><header class="hoja-cab"><h2>${esc(o.titulo)}</h2><button data-cerrar>x</button></header><div class="hoja-cuerpo"></div>${o.pie ? '<footer class="hoja-pie"></footer>' : ""}</section></div>`);
      const cuerpo = $(".hoja-cuerpo", velo);
      if (typeof o.cuerpo === "string") cuerpo.innerHTML = o.cuerpo; else if (o.cuerpo) cuerpo.append(o.cuerpo);
      if (o.pie) {
        const pie = $(".hoja-pie", velo);
        if (typeof o.pie === "string") pie.innerHTML = o.pie; else pie.append(o.pie);
      }
      const tecla = (e) => { if (e.key === "Escape") { e.stopPropagation(); apiHoja.cerrar(); } };
      const apiHoja = {
        el: velo,
        cuerpo,
        cerrar() {
          if (!velo.isConnected) return;
          doc.removeEventListener("keydown", tecla, true);
          if (hojaActual === apiHoja) hojaActual = null;
          velo.remove();
          o.alCerrar && o.alCerrar();
        },
      };
      doc.addEventListener("keydown", tecla, true);
      doc.body.append(velo);
      hojaActual = apiHoja;
      return apiHoja;
    },
    confirmar: async () => true,
    local: { activo: () => !!opc.sinCuenta, usuario: () => ({ id: "local", nombre: "Sin cuenta" }), lista: () => [] },
    api: {
      empaquetada: false,
      token: () => (opc.sinCuenta ? null : "tok-1"),
      usuario: () => ({ id: "u1", nombre: "Ana", empresa: "Escuela" }),
      clave: (k) => "nom.u.u1." + k,
      get: async (r) => {
        if (r === "/api/cuentas/yo") return { preferencias: { tema: "auto", autoguardadoMin: 5, versionesConservar: 30, calidad: "alta", movimientoReducido: false, enfoqueAtenuar: true, idioma: "es" }, recuperacion: true };
        return {};
      },
      put: async (r, c) => { puts.push({ ruta: r, cuerpo: JSON.parse(JSON.stringify(c)) }); return c; },
      post: async () => ({}),
    },
    app: { espacios: [], irACuentas() {}, cerrarSesion() {} },
  };

  const navigatorObj = opc.navigator !== undefined ? opc.navigator : {};
  win.NOM = NOM;
  win.document = doc;
  win.localStorage = localStorage;
  win.navigator = navigatorObj;
  win.matchMedia = () => ({ matches: false, addEventListener() {} });
  win.innerWidth = 1024;
  win.innerHeight = 768;
  win.location = { hash: "#/", protocol: "https:", origin: "https://www.nom.qd.je" };
  win.requestAnimationFrame = (fn) => { fn(); return 1; };
  if (opc.NomAndroid !== undefined) win.NomAndroid = opc.NomAndroid;

  const sandbox = {
    window: win,
    document: doc,
    localStorage,
    navigator: navigatorObj,
    NOM,
    location: win.location,
    matchMedia: win.matchMedia,
    innerWidth: win.innerWidth,
    innerHeight: win.innerHeight,
    requestAnimationFrame: win.requestAnimationFrame,
    CustomEvent: class CustomEvent { constructor(t, o) { this.type = t; this.detail = o && o.detail; } },
    setTimeout: (fn, ms) => (ms === 0 ? (fn(), 0) : global.setTimeout(fn, ms)),
    clearTimeout: global.clearTimeout,
    console,
  };
  const ctx = vm.createContext(sandbox);
  const cargar = (archivo) => {
    const ruta = path.join(dirJs, archivo);
    vm.runInContext(fs.readFileSync(ruta, "utf8"), ctx, { filename: archivo });
  };
  const enviarTecla = (props, destino) => {
    const ev = crearEvento("keydown", props);
    despachar(destino || doc.activeElement || doc.body, ev);
    return ev;
  };
  const clickEn = (el) => {
    const ev = crearEvento("click", { target: el });
    despachar(el, ev);
    return ev;
  };

  return { win, doc, NOM, alm, localStorage, avisos, puts, cargar, crearEvento, despachar, enviarTecla, clickEn, hojaActual: () => hojaActual };
}

// ================================================================== 1. Atajos de teclado
test("1.1 preferencias-extra.js expone la tabla de atajos, valida prohibidos y detecta choques", () => {
  const env = crearEntorno();
  env.cargar("preferencias-extra.js");
  const P = env.NOM.prefExtra;
  assert.ok(P, "Debe existir NOM.prefExtra");
  assert.ok(Array.isArray(P.ATAJOS) && P.ATAJOS.length >= 4, "Debe exponer la tabla de atajos");
  for (const a of P.ATAJOS) {
    assert.ok(a.accion && a.nombre && typeof a.tecla === "string", "Cada fila necesita acción, nombre y tecla de siempre");
  }

  // Prohibidos del navegador: ninguno se debe permitir.
  for (const prohibido of ["Ctrl+W", "Ctrl+T", "Ctrl+N", "Ctrl+Tab", "F5", "Alt+F4"]) {
    const r = P.cambiarAtajo("paleta", prohibido);
    assert.equal(r.ok, false, `Debe rechazar ${prohibido}`);
    assert.equal(r.motivo, "prohibido", `Motivo de ${prohibido} debe ser prohibido`);
  }
  assert.equal(P.teclaDe("paleta"), "Ctrl+K", "No debe cambiar al intentar un atajo prohibido");

  // Colisión: no se permiten dos acciones con la misma tecla y debe avisar cuál choca.
  const choque = P.cambiarAtajo("paleta", "Alt+/");
  assert.equal(choque.ok, false, "Debe rechazar choque con otra acción");
  assert.equal(choque.motivo, "choca");
  assert.equal(choque.chocaCon, "paletaTexto");
  assert.ok(choque.mensaje.includes(choque.chocaNombre), "El aviso debe nombrar la acción con la que choca");

  // Cambio válido, quitar y volver a los de siempre.
  const ok = P.cambiarAtajo("paleta", "Ctrl+J");
  assert.equal(ok.ok, true);
  assert.equal(P.teclaDe("paleta"), "Ctrl+J");
  assert.equal(JSON.parse(env.localStorage.getItem("nom.pref.extra")).atajos.paleta, "Ctrl+J");

  P.quitarAtajo("paleta");
  assert.equal(P.teclaDe("paleta"), "");

  P.restablecerAtajos();
  assert.equal(P.teclaDe("paleta"), "Ctrl+K");
});

test("1.2 experiencia.js consulta los atajos personalizados (cambiar, quitar y restablecer)", () => {
  const env = crearEntorno();
  env.cargar("preferencias-extra.js");
  env.cargar("experiencia.js");
  const P = env.NOM.prefExtra;

  // Cambiar "Buscar o ir a" a Ctrl+J: Ctrl+K ya no abre la paleta, Ctrl+J sí.
  assert.equal(P.cambiarAtajo("paleta", "Ctrl+J").ok, true);
  env.enviarTecla({ key: "k", ctrlKey: true });
  assert.equal(env.doc.querySelector(".paleta"), null, "Ctrl+K ya no debe abrir la paleta");

  env.enviarTecla({ key: "j", ctrlKey: true });
  assert.ok(env.doc.querySelector(".paleta"), "Ctrl+J debe abrir la paleta");
  env.enviarTecla({ key: "Escape" });

  // Quitar el atajo de la hoja de atajos (?): ya no debe abrirla.
  P.quitarAtajo("atajos");
  env.enviarTecla({ key: "?", shiftKey: true });
  assert.equal(env.hojaActual(), null, "? no debe abrir atajos tras quitarlo");

  // Volver a los de siempre: ? vuelve a abrir la hoja de atajos.
  P.restablecerAtajos();
  env.enviarTecla({ key: "?", shiftKey: true });
  assert.ok(env.hojaActual(), "? debe volver a abrir la hoja de atajos");

  // Atajo del documento personalizado (Guardar versión -> Alt+G): dispara la acción y suprime Ctrl+S viejo.
  env.hojaActual().cerrar();
  let guardados = 0;
  const vista = env.doc.createElement("div");
  env.doc.body.append(vista);
  env.NOM.editorActual = { vista, acciones: () => [{ texto: "Guardar versión", fn: () => { guardados++; } }] };
  assert.equal(P.cambiarAtajo("guardar", "Alt+G").ok, true);
  env.enviarTecla({ key: "g", altKey: true });
  assert.equal(guardados, 1, "Alt+G debe ejecutar Guardar versión");
  const evViejo = env.enviarTecla({ key: "s", ctrlKey: true });
  assert.equal(evViejo.defaultPrevented, true, "Ctrl+S antiguo queda bloqueado al haberse cambiado");
  assert.equal(guardados, 1, "Ctrl+S antiguo no debe disparar Guardar versión");
});

test("1.3 Ajustes guarda preferencias extra en localStorage sin mandar claves extra al servidor (con cuenta y sin cuenta)", async () => {
  // Con cuenta
  const env = crearEntorno({ sinCuenta: false });
  env.cargar("preferencias-extra.js");
  env.cargar("vista-espacios.js");

  await env.NOM.vistaEspacios.preferencias();
  const h = env.hojaActual();
  assert.ok(h, "Debe abrir la hoja de Ajustes");
  assert.ok(h.cuerpo.querySelector("[data-menu-rapido]"), "Debe mostrar Ruedita · Lista");
  assert.ok(h.cuerpo.querySelector("[data-haptica]"), "Debe mostrar Vibrar al tocar");
  assert.ok(h.cuerpo.querySelector("[data-ajustes-atajos]"), "Debe mostrar la sección Atajos");

  // Cambiar a Lista, apagar vibración, cambiar atajo de paleta a Ctrl+J (y probar bloqueo de Ctrl+W + Escape no cierra Ajustes)
  env.clickEn(h.cuerpo.querySelector('[data-menu-rapido] [data-v="lista"]'));
  const chkHap = h.cuerpo.querySelector("[data-haptica]");
  chkHap.checked = false;
  env.despachar(chkHap, env.crearEvento("change", { target: chkHap }));

  env.clickEn(h.cuerpo.querySelector('[data-cambiar="paleta"]'));
  env.enviarTecla({ key: "w", ctrlKey: true });
  assert.equal(h.cuerpo.querySelector("[data-atajos-aviso]").hidden, false, "Ctrl+W debe mostrar aviso");
  env.enviarTecla({ key: "Escape" });
  assert.ok(env.hojaActual(), "Esc al cancelar captura no debe cerrar la ventana de Ajustes");

  env.clickEn(h.cuerpo.querySelector('[data-cambiar="paleta"]'));
  env.enviarTecla({ key: "j", ctrlKey: true });

  // Guardar
  await env.clickEn(h.el.querySelector("[data-guardar]"));
  await new Promise((r) => setTimeout(r, 10));

  const guardado = env.NOM.prefExtra.leer();
  assert.equal(guardado.menuRapido, "lista");
  assert.equal(guardado.haptica, false);
  assert.equal(env.NOM.prefExtra.teclaDe("paleta"), "Ctrl+J");
  assert.equal(env.puts.length, 1, "Debe llamar a PUT /api/cuentas/preferencias");
  assert.equal("atajos" in env.puts[0].cuerpo, false, "No debe mandar claves desconocidas al servidor");
  assert.equal("menuRapido" in env.puts[0].cuerpo, false);

  // Modo sin cuenta: también guarda en localStorage sin llamar a la API.
  const envLocal = crearEntorno({ sinCuenta: true });
  envLocal.cargar("preferencias-extra.js");
  envLocal.cargar("vista-espacios.js");
  await envLocal.NOM.vistaEspacios.preferencias();
  const hl = envLocal.hojaActual();
  envLocal.clickEn(hl.cuerpo.querySelector('[data-menu-rapido] [data-v="lista"]'));
  await envLocal.clickEn(hl.el.querySelector("[data-guardar]"));
  await new Promise((r) => setTimeout(r, 10));
  assert.equal(envLocal.puts.length, 0, "Sin cuenta no debe llamar a PUT");
  assert.equal(envLocal.NOM.prefExtra.leer().menuRapido, "lista");
});

// ================================================================== 2. Háptica
test("2.1 Háptica respeta encendido/apagado, intensidades, puente Android, navigator.vibrate y no truena sin nada", () => {
  // Sin navigator.vibrate ni puente Android: no truena y devuelve false.
  const envNada = crearEntorno({ navigator: {} });
  envNada.cargar("preferencias-extra.js");
  assert.equal(envNada.NOM.prefExtra.vibrar("abrir"), false);

  // Con navigator.vibrate: duraciones ordenadas suave < normal < fuerte y silencio al apagar.
  const vibs = [];
  const envNav = crearEntorno({ navigator: { vibrate: (ms) => { vibs.push(ms); return true; } } });
  envNav.cargar("preferencias-extra.js");
  const P = envNav.NOM.prefExtra;

  P.guardar({ haptica: true, hapticaNivel: "suave" });
  P.vibrar("abrir");
  P.guardar({ haptica: true, hapticaNivel: "normal" });
  P.vibrar("abrir");
  P.guardar({ haptica: true, hapticaNivel: "fuerte" });
  P.vibrar("abrir");
  assert.equal(vibs.length, 3);
  assert.ok(vibs[0] > 0 && vibs[0] < vibs[1] && vibs[1] < vibs[2] && vibs[2] <= 50, `Duraciones esperadas cortas y crecientes: ${vibs}`);

  vibs.length = 0;
  P.guardar({ haptica: false, hapticaNivel: "fuerte" });
  assert.equal(P.vibrar("abrir"), false);
  assert.equal(vibs.length, 0, "Apagado no debe vibrar");

  // Con puente Android (window.NomAndroid.vibrar): lo prefiere sobre navigator.vibrate.
  const androidVibs = [], navVibs = [];
  const envApp = crearEntorno({
    navigator: { vibrate: (ms) => { navVibs.push(ms); } },
    NomAndroid: { vibrar: (ms) => { androidVibs.push(ms); } },
  });
  envApp.cargar("preferencias-extra.js");
  envApp.NOM.prefExtra.vibrar("elegir");
  assert.equal(androidVibs.length, 1, "Debe usar el puente NomAndroid.vibrar");
  assert.equal(navVibs.length, 0, "Si el puente vibró, no duplica con navigator.vibrate");
});

test("2.2 La ruedita vibra al abrir y al elegir cuando la háptica está encendida", () => {
  const vibs = [];
  const env = crearEntorno({ navigator: { vibrate: (ms) => { vibs.push(ms); return true; } } });
  env.cargar("preferencias-extra.js");
  env.cargar("rueda.js");

  let ejecutado = false;
  env.NOM.rueda.abrir(200, 200, [
    { texto: "Negrita", icono: "bold", fn: () => { ejecutado = true; } },
    { texto: "Cursiva", icono: "italic", fn: () => {} },
  ], { tactil: true, modo: "clic" });
  assert.equal(vibs.length, 1, "Debe vibrar al abrir");

  // Elegir con teclado (Enter tras marcar primero)
  env.enviarTecla({ key: "ArrowDown" });
  vibs.length = 0;
  env.enviarTecla({ key: "Enter" });
  assert.equal(ejecutado, true, "Debe ejecutar la acción elegida");
  assert.equal(vibs.length, 1, "Debe vibrar al elegir");

  // Con háptica apagada, no vibra ni al abrir ni al elegir.
  vibs.length = 0;
  env.NOM.prefExtra.guardar({ haptica: false });
  env.NOM.rueda.abrir(200, 200, [{ texto: "Negrita", icono: "bold", fn: () => {} }], { tactil: true, modo: "clic" });
  env.enviarTecla({ key: "ArrowDown" });
  env.enviarTecla({ key: "Enter" });
  assert.equal(vibs.length, 0, "Con vibración apagada no debe vibrar");
});

// ================================================================== 3. Ruedita o lista clásica
test("3.1 Preferencia Lista abre menú contextual vertical con mismas acciones, orden, ARIA y teclado", () => {
  const env = crearEntorno();
  env.cargar("preferencias-extra.js");
  env.cargar("rueda.js");

  env.NOM.prefExtra.guardar({ menuRapido: "lista" });
  const llamadas = [];
  const items = [
    { texto: "Negrita", icono: "bold", marcado: true, fn: () => llamadas.push("negrita") },
    { texto: "Cortar", icono: "scissors", off: true, fn: () => llamadas.push("cortar") },
    { texto: "Copiar", icono: "copy", fn: () => llamadas.push("copiar") },
    { texto: "Más", icono: "more-horizontal", sub: [
      { texto: "Pegar", icono: "clipboard-paste", fn: () => llamadas.push("pegar") },
    ] },
  ];

  env.NOM.rueda.abrir(150, 150, items, { titulo: "Texto", modo: "clic" });
  assert.equal(env.NOM.rueda.abierta(), true);

  const menu = env.doc.querySelector('.rueda-lista[role="menu"]');
  assert.ok(menu, "Debe renderizar el menú contextual vertical con role=menu");
  assert.equal(env.doc.querySelector(".rueda"), null, "En modo lista no debe pintar el anillo circular");

  const opciones = menu.querySelectorAll('[role="menuitem"]');
  assert.equal(opciones.length, 4, "Debe tener las mismas 4 acciones");
  assert.deepEqual(opciones.map((o) => o.textContent.trim()), ["Negrita", "Cortar", "Copiar", "Más"]);
  assert.equal(opciones[1].getAttribute("aria-disabled"), "true", "La acción apagada conserva aria-disabled");

  // Teclado: empieza en el primero habilitado (Negrita), ArrowDown salta Cortar (off) hacia Copiar, Enter lo ejecuta.
  assert.equal(env.doc.activeElement, opciones[0]);
  env.enviarTecla({ key: "ArrowDown" });
  assert.equal(env.doc.activeElement, opciones[2], "ArrowDown debe saltar los elementos apagados");
  env.enviarTecla({ key: "Enter" });
  assert.deepEqual(llamadas, ["copiar"]);
  assert.equal(env.NOM.rueda.abierta(), false, "Tras elegir debe cerrarse");

  // Submenú y regreso con Retroceso; cierre con Esc.
  env.NOM.rueda.abrir(150, 150, items, { titulo: "Texto", modo: "clic" });
  env.enviarTecla({ key: "ArrowUp" }); // salta del primero al último ("Más")
  env.enviarTecla({ key: "Enter" });
  const subOpciones = env.doc.querySelectorAll('.rueda-lista [role="menuitem"]');
  assert.deepEqual(subOpciones.map((o) => o.textContent.trim()), ["Pegar"]);

  env.enviarTecla({ key: "Backspace" });
  assert.equal(env.doc.querySelectorAll('.rueda-lista [role="menuitem"]').length, 4, "Retroceso regresa al menú anterior");

  env.enviarTecla({ key: "Escape" });
  assert.equal(env.NOM.rueda.abierta(), false, "Esc cierra el menú lista");

  // Al abrir manteniendo presionado, soltar el dedo no cierra ni dispara una acción por accidente; el siguiente toque sí.
  env.NOM.rueda.abrir(150, 150, items, { titulo: "Texto", modo: "mantener", tactil: true });
  env.despachar(env.doc.body, env.crearEvento("pointerup", { clientX: 150, clientY: 150, pointerType: "touch" }));
  const btnCopiar = env.doc.querySelectorAll('.rueda-lista [role="menuitem"]')[2];
  env.despachar(btnCopiar, env.crearEvento("click", { target: btnCopiar, clientX: 150, clientY: 150 }));
  assert.equal(env.NOM.rueda.abierta(), true, "El clic sintético al soltar el gesto no debe disparar ni cerrar");
  assert.deepEqual(llamadas, ["copiar"]);
  env.despachar(btnCopiar, env.crearEvento("click", { target: btnCopiar, clientX: 160, clientY: 200 }));
  assert.deepEqual(llamadas, ["copiar", "copiar"], "El toque real sobre una opción sí la ejecuta");
  assert.equal(env.NOM.rueda.abierta(), false);
});

// ================================================================== 4. Carga en index.html y sw.js
test("4.1 index.html y sw.js registran js/preferencias-extra.js en el orden correcto", () => {
  const indexHtml = fs.readFileSync(path.join(dirStatic, "index.html"), "utf8");
  const swJs = fs.readFileSync(path.join(dirStatic, "sw.js"), "utf8");
  const iNucleo = indexHtml.indexOf('src="js/nucleo.js"');
  const iExtra = indexHtml.indexOf('src="js/preferencias-extra.js"');
  const iRueda = indexHtml.indexOf('src="js/rueda.js"');
  const iEspacios = indexHtml.indexOf('src="js/vista-espacios.js"');
  const iExp = indexHtml.indexOf('src="js/experiencia.js"');
  assert.ok(iExtra > iNucleo && iExtra < iRueda && iExtra < iEspacios && iExtra < iExp,
    "index.html debe cargar js/preferencias-extra.js después de nucleo.js y antes de rueda.js, vista-espacios.js y experiencia.js");
  assert.ok(swJs.includes('"js/preferencias-extra.js"'), "sw.js debe incluir js/preferencias-extra.js en BASE");
});

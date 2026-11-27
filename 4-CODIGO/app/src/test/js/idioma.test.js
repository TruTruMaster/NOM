// Pruebas de js/idioma.js, js/textos/{es,en}.json, preferencia Idioma y rendimiento del observador.
// Uso: node --test 4-CODIGO/app/src/test/js/idioma.test.js
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { performance } = require("node:perf_hooks");

const dirStatic = path.join(__dirname, "../../main/resources/static");
const dirJs = path.join(dirStatic, "js");
const dirTextos = path.join(dirJs, "textos");

// ---------------------------------------------------------------- DOM mínimo con nodos de texto y MutationObserver
function crearDomConTexto() {
  const observadores = [];
  const listeners = new WeakMap();
  const desEsc = (s) => String(s || "")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&");

  class EventTargetMin {
    addEventListener(tipo, fn, opc) {
      const cap = typeof opc === "boolean" ? opc : !!(opc && opc.capture);
      let m = listeners.get(this);
      if (!m) { m = new Map(); listeners.set(this, m); }
      if (!m.has(tipo)) m.set(tipo, []);
      m.get(tipo).push({ fn, cap });
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

  class Texto {
    constructor(data, doc) {
      this.nodeType = 3;
      this.ownerDocument = doc;
      this.parentNode = null;
      this._data = String(data == null ? "" : data);
    }
    get nodeValue() { return this._data; }
    set nodeValue(v) { this._data = String(v == null ? "" : v); }
    get data() { return this._data; }
    set data(v) { this._data = String(v == null ? "" : v); }
    get textContent() { return this._data; }
    set textContent(v) { this._data = String(v == null ? "" : v); }
    remove() {
      if (!this.parentNode) return;
      const i = this.parentNode.childNodes.indexOf(this);
      if (i >= 0) this.parentNode.childNodes.splice(i, 1);
      this.parentNode = null;
    }
  }

  class Nodo extends EventTargetMin {
    constructor(tag, doc) {
      super();
      this.nodeType = 1;
      this.tagName = tag.toUpperCase();
      this.ownerDocument = doc;
      this.attrs = new Map();
      this.childNodes = [];
      this.parentNode = null;
      this.style = { setProperty() {} };
      this.value = "";
      this.checked = false;
      this.disabled = false;
      this.hidden = false;
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
    get children() { return this.childNodes.filter((c) => c.nodeType === 1); }
    get firstElementChild() { return this.children[0] || null; }
    get isContentEditable() {
      let n = this;
      while (n && n.nodeType === 1) {
        if (n.attrs.has("contenteditable")) {
          const v = String(n.attrs.get("contenteditable")).toLowerCase();
          if (v === "false") return false;
          if (v === "" || v === "true" || v === "plaintext-only") return true;
        }
        n = n.parentNode;
      }
      return false;
    }
    get isConnected() {
      let n = this;
      while (n) { if (n === this.ownerDocument.body || n === this.ownerDocument.documentElement) return true; n = n.parentNode; }
      return false;
    }
    get textContent() { return this.childNodes.map((c) => c.textContent).join(""); }
    set textContent(v) {
      this.childNodes = [];
      const s = String(v == null ? "" : v);
      if (s) {
        const t = new Texto(s, this.ownerDocument);
        t.parentNode = this;
        this.childNodes.push(t);
        notificar(this, [t]);
      }
    }
    get innerHTML() { return ""; }
    set innerHTML(s) {
      this.childNodes = [];
      parsear(String(s || ""), this, this.ownerDocument);
      notificar(this, this.childNodes.slice());
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
      const agregados = [];
      for (const n of ns) {
        if (n == null) continue;
        const hijo = typeof n === "string" ? new Texto(n, this.ownerDocument) : n;
        if (hijo.parentNode) hijo.remove();
        hijo.parentNode = this;
        this.childNodes.push(hijo);
        agregados.push(hijo);
      }
      if (agregados.length) notificar(this, agregados);
    }
    appendChild(n) { this.append(n); return n; }
    remove() {
      if (!this.parentNode) return;
      const i = this.parentNode.childNodes.indexOf(this);
      if (i >= 0) this.parentNode.childNodes.splice(i, 1);
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
        for (const c of n.childNodes) {
          if (c.nodeType !== 1) continue;
          if (c.matches(sel)) out.push(c);
          rec(c);
        }
      };
      rec(this);
      return out;
    }
    querySelector(sel) { return this.querySelectorAll(sel)[0] || null; }
    focus() { this.ownerDocument.activeElement = this; }
    blur() {}
    scrollIntoView() {}
    getBoundingClientRect() { return { left: 0, top: 0, right: 100, bottom: 30, width: 100, height: 30 }; }
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
      } else if (el.tagName !== p.toUpperCase()) return false;
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
  function parsear(cadena, padre, doc) {
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
        el.parentNode = tope;
        tope.childNodes.push(el);
        if (!auto) pila.push(el);
      } else if (m[5]) {
        const tx = new Texto(desEsc(m[5]), doc);
        tx.parentNode = tope;
        tope.childNodes.push(tx);
      }
    }
  }

  class MutationObserverMin {
    constructor(cb) {
      this._cb = cb;
      this._objetivo = null;
      this._activo = false;
    }
    observe(target, opc) {
      this._objetivo = target;
      this._opc = opc || {};
      this._activo = true;
      if (!observadores.includes(this)) observadores.push(this);
    }
    disconnect() {
      this._activo = false;
      this._objetivo = null;
      const i = observadores.indexOf(this);
      if (i >= 0) observadores.splice(i, 1);
    }
  }

  function notificar(target, addedNodes) {
    if (!observadores.length) return;
    const rec = { type: "childList", target, addedNodes };
    for (const o of observadores.slice()) {
      if (!o._activo || !o._objetivo) continue;
      if (o._objetivo === target || (o._opc.subtree && o._objetivo.contains(target))) {
        o._cb([rec], o);
      }
    }
  }

  const doc = new EventTargetMin();
  const win = new EventTargetMin();
  doc.documentElement = new Nodo("html", doc);
  doc.documentElement.setAttribute("lang", "es");
  doc.body = new Nodo("body", doc);
  doc.documentElement.childNodes.push(doc.body);
  doc.body.parentNode = doc.documentElement;
  doc.activeElement = doc.body;
  doc.readyState = "complete";
  doc.createElement = (tag) => new Nodo(tag, doc);
  doc.createTextNode = (s) => new Texto(s, doc);
  doc.querySelector = (s) => doc.documentElement.querySelector(s);
  doc.querySelectorAll = (s) => doc.documentElement.querySelectorAll(s);

  function crearEvento(tipo, props) {
    return Object.assign({
      type: tipo, target: null, currentTarget: null, defaultPrevented: false, _parado: false,
      preventDefault() { this.defaultPrevented = true; },
      stopPropagation() { this._parado = true; },
    }, props);
  }
  function despachar(destino, ev) {
    ev.target = ev.target || destino;
    const camino = [];
    let n = destino;
    while (n && n.nodeType === 1) { camino.push(n); n = n.parentNode; }
    for (const nodo of camino) {
      if (ev._parado) break;
      ev.currentTarget = nodo;
      const prop = "on" + ev.type;
      if (typeof nodo[prop] === "function") nodo[prop].call(nodo, ev);
      for (const h of nodo._oir(ev.type, false)) h.fn.call(nodo, ev);
    }
    return !ev.defaultPrevented;
  }

  return { win, doc, Nodo, Texto, MutationObserverMin, observadores, crearEvento, despachar };
}

function crearEntorno(opc) {
  opc = opc || {};
  const dom = crearDomConTexto();
  const { win, doc, MutationObserverMin, observadores, crearEvento, despachar } = dom;
  const alm = new Map();
  const localStorage = {
    getItem: (k) => (alm.has(String(k)) ? alm.get(String(k)) : null),
    setItem: (k, v) => { alm.set(String(k), String(v)); },
    removeItem: (k) => { alm.delete(String(k)); },
  };
  const puts = [];
  let hojaActual = null;
  const esc = (s) => String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  const html = (cadena) => {
    const c = doc.createElement("div");
    c.innerHTML = cadena.trim();
    return c.firstElementChild;
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
    avatar: (u) => `<span class="avatar">${esc((u && u.nombre) || "?")}</span>`,
    sinContrasena: () => false,
    etiquetaEquipo: () => "Guardado en este equipo",
    haceCuanto: () => "hace un momento",
    bytes: (n) => n + " B",
    MODOS: { TEXTO: { nombre: "Texto", icono: "file-text" } },
    leer: (k, def) => { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); },
    guardar: (k, v) => { localStorage.setItem(k, JSON.stringify(v)); return true; },
    quitar: (k) => localStorage.removeItem(k),
    aviso: () => {},
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
      const apiHoja = {
        el: velo, cuerpo,
        cerrar() { if (!velo.isConnected) return; if (hojaActual === apiHoja) hojaActual = null; velo.remove(); o.alCerrar && o.alCerrar(); },
      };
      doc.body.append(velo);
      hojaActual = apiHoja;
      return apiHoja;
    },
    confirmar: async () => true,
    local: { activo: () => false, usuario: () => ({ id: "local", nombre: "Sin cuenta" }), lista: () => [] },
    api: {
      empaquetada: false,
      token: () => "tok-1",
      usuario: () => ({ id: "u1", nombre: "Ana", empresa: "Escuela" }),
      clave: (k) => "nom.u.u1." + k,
      get: async (r) => (r === "/api/cuentas/yo" ? { preferencias: { tema: "auto", autoguardadoMin: 5, versionesConservar: 30, calidad: "alta", movimientoReducido: false, enfoqueAtenuar: true, idioma: "es" }, recuperacion: true } : {}),
      put: async (r, c) => { puts.push({ ruta: r, cuerpo: JSON.parse(JSON.stringify(c)) }); return c; },
      post: async () => ({}),
    },
    app: { espacios: [], irACuentas() {}, cerrarSesion() {} },
  };

  const navObj = opc.navigator || { languages: ["es-MX"], language: "es-MX" };
  const fetchMock = async (url) => {
    const ruta = path.join(dirStatic, String(url).replace(/^\/+/, ""));
    if (!fs.existsSync(ruta)) return { ok: false, status: 404, json: async () => ({}) };
    const txt = fs.readFileSync(ruta, "utf8");
    return { ok: true, status: 200, json: async () => JSON.parse(txt) };
  };

  win.NOM = NOM;
  win.document = doc;
  win.localStorage = localStorage;
  win.navigator = navObj;
  win.location = { hash: opc.hash || "#/", protocol: "https:", origin: "https://www.nom.qd.je" };
  win.matchMedia = () => ({ matches: false, addEventListener() {} });
  win.MutationObserver = MutationObserverMin;
  win.fetch = fetchMock;

  const sandbox = {
    window: win,
    document: doc,
    localStorage,
    navigator: navObj,
    NOM,
    location: win.location,
    matchMedia: win.matchMedia,
    MutationObserver: MutationObserverMin,
    fetch: fetchMock,
    setTimeout: global.setTimeout,
    clearTimeout: global.clearTimeout,
    console,
  };
  const ctx = vm.createContext(sandbox);
  const cargar = (archivo) => vm.runInContext(fs.readFileSync(path.join(dirJs, archivo), "utf8"), ctx, { filename: archivo });
  return { win, doc, NOM, alm, localStorage, puts, observadores, crearEvento, despachar, cargar, hojaActual: () => hojaActual };
}

// ================================================================== 1. Diccionarios y carga
test("1.1 textos/es.json y textos/en.json existen, tienen las mismas claves y cubren la interfaz principal", () => {
  const es = JSON.parse(fs.readFileSync(path.join(dirTextos, "es.json"), "utf8"));
  const en = JSON.parse(fs.readFileSync(path.join(dirTextos, "en.json"), "utf8"));
  const kEs = Object.keys(es).sort();
  const kEn = Object.keys(en).sort();
  assert.deepEqual(kEs, kEn, "es.json y en.json deben tener exactamente las mismas claves");
  assert.ok(kEn.length >= 300, `La tabla debe cubrir las pantallas principales (tiene ${kEn.length})`);
  for (const k of kEn) {
    assert.ok(typeof en[k] === "string" && en[k].trim().length > 0, `Traducción vacía para: ${k}`);
  }
  for (const clave of ["Inicio", "Entrar", "Crear cuenta", "Ajustes y preferencias", "Negrita", "Guardar", "Cancelar", "Buscar o ir a…", "Vibrar al tocar"]) {
    assert.ok(clave in en, `Falta clave esencial: ${clave}`);
  }
});

test("1.2 index.html y sw.js registran idioma.js (justo después de nucleo.js) y los diccionarios", () => {
  const html = fs.readFileSync(path.join(dirStatic, "index.html"), "utf8");
  const sw = fs.readFileSync(path.join(dirStatic, "sw.js"), "utf8");
  assert.match(html, /<html\s+lang="es"/, "index.html debe declarar lang en <html>");
  assert.ok(
    html.includes('<script src="js/nucleo.js"></script>\n  <script src="js/idioma.js"></script>'),
    "index.html debe cargar js/idioma.js justo después de js/nucleo.js"
  );
  for (const archivo of ['"js/idioma.js"', '"js/textos/es.json"', '"js/textos/en.json"']) {
    assert.ok(sw.includes(archivo), `sw.js BASE debe incluir ${archivo}`);
  }
});

// ================================================================== 2. Preferencia «Idioma»
test("2.1 Idioma Automático resuelve es/en según navigator.languages y Ajustes lo guarda en nom.pref.extra", async () => {
  const envEs = crearEntorno({ navigator: { languages: ["es-MX", "en-US"] } });
  envEs.cargar("idioma.js");
  envEs.cargar("preferencias-extra.js");
  assert.equal(envEs.NOM.idioma.actual(), "es");
  assert.equal(envEs.doc.documentElement.getAttribute("lang"), "es");

  const envEn = crearEntorno({ navigator: { languages: ["en-US", "es-MX"] } });
  envEn.cargar("idioma.js");
  envEn.cargar("preferencias-extra.js");
  assert.equal(envEn.NOM.idioma.actual(), "en");
  assert.equal(envEn.doc.documentElement.getAttribute("lang"), "en");

  const envFr = crearEntorno({ navigator: { languages: ["fr-FR"] } });
  envFr.cargar("idioma.js");
  assert.equal(envFr.NOM.idioma.actual(), "en", "Si el sistema no empieza con es, usa inglés");

  // Ajustes: elegir English y Guardar guarda idiomaInterfaz en nom.pref.extra sin tocar Preferencias.idioma del servidor
  const envAj = crearEntorno({ navigator: { languages: ["es-MX"] } });
  envAj.cargar("idioma.js");
  envAj.cargar("preferencias-extra.js");
  envAj.cargar("vista-espacios.js");
  await envAj.NOM.vistaEspacios.preferencias();
  const h = envAj.hojaActual();
  const btnEn = h.cuerpo.querySelector('[data-idioma] [data-v="en"]');
  assert.ok(btnEn, "Ajustes debe tener el selector de Idioma (Automático · Español · English)");
  envAj.despachar(btnEn, envAj.crearEvento("click", { target: btnEn }));
  await envAj.despachar(h.el.querySelector("[data-guardar]"), envAj.crearEvento("click"));
  await new Promise((r) => setTimeout(r, 15));

  assert.equal(envAj.NOM.prefExtra.leer().idiomaInterfaz, "en");
  assert.equal(envAj.NOM.idioma.actual(), "en");
  assert.equal(envAj.doc.documentElement.getAttribute("lang"), "en");
  assert.equal(envAj.puts.length, 1);
  assert.equal(envAj.puts[0].cuerpo.idioma, "es", "No debe alterar Preferencias.idioma del servidor");
  assert.equal("idiomaInterfaz" in envAj.puts[0].cuerpo, false, "No debe mandar idiomaInterfaz al servidor");
});

// ================================================================== 3. Traducción de UI y exclusiones
test("3.1 Traduce nodos de texto, atributos (aria-label, title, placeholder, alt) y mutaciones nuevas, excluyendo documentos y nombres", async () => {
  const env = crearEntorno({ navigator: { languages: ["es-MX"] } });
  env.cargar("idioma.js");
  env.cargar("preferencias-extra.js");

  env.doc.body.innerHTML = `
    <div id="ui">
      <button aria-label="Guardar" title="Guardar versión">Guardar</button>
      <input placeholder="Busca documentos, espacios o acciones…">
      <img alt="Imagen">
      <span id="din">3 espacios</span>
    </div>
    <div class="wd-hoja"><div class="wd-cuerpo" contenteditable="true"><p>Guardar</p><p>Inicio</p></div></div>
    <div class="rejilla-hoja"><table class="xl"><tr><td>Guardar</td></tr></table></div>
    <div class="escenario"><div class="diapo" data-mundo><div class="bl">Inicio</div></div><div class="zoom-ctl"><button aria-label="Acercar">Ajustar</button></div></div>
    <div class="burbuja-chat"><p>Guardar</p></div>
    <div class="hilo-msj"><p>Inicio</p></div>
    <a class="doc-tarjeta"><div class="cuerpo"><span class="titulo">Inicio</span></div></a>
    <a class="tarjeta-espacio"><div class="titulo">Ajustes</div></a>
    <nav class="riel-lista"><a class="riel-item"><span class="crece">Inicio</span></a></nav>
  `;

  await env.NOM.idioma.aplicar("en");
  assert.equal(env.doc.documentElement.getAttribute("lang"), "en");

  // UI traducida (texto + 4 atributos + dinámico)
  const btn = env.doc.querySelector("#ui button");
  assert.equal(btn.textContent, "Save");
  assert.equal(btn.getAttribute("aria-label"), "Save");
  assert.equal(btn.getAttribute("title"), "Save version");
  assert.equal(env.doc.querySelector("#ui input").getAttribute("placeholder"), "Search documents, spaces, or actions…");
  assert.equal(env.doc.querySelector("#ui img").getAttribute("alt"), "Image");
  assert.equal(env.doc.querySelector("#din").textContent, "3 spaces");
  // Controles UI dentro de escenario (fuera de .diapo) SÍ se traducen
  assert.equal(env.doc.querySelector(".zoom-ctl button").textContent, "Fit");
  assert.equal(env.doc.querySelector(".zoom-ctl button").getAttribute("aria-label"), "Zoom in");

  // Exclusiones: contenido del documento y texto de personas intactos
  assert.equal(env.doc.querySelector(".wd-cuerpo").textContent, "GuardarInicio");
  assert.equal(env.doc.querySelector(".rejilla-hoja td").textContent, "Guardar");
  assert.equal(env.doc.querySelector(".diapo .bl").textContent, "Inicio");
  assert.equal(env.doc.querySelector(".burbuja-chat p").textContent, "Guardar");
  assert.equal(env.doc.querySelector(".hilo-msj p").textContent, "Inicio");
  assert.equal(env.doc.querySelector(".doc-tarjeta .titulo").textContent, "Inicio");
  assert.equal(env.doc.querySelector(".tarjeta-espacio .titulo").textContent, "Ajustes");
  assert.equal(env.doc.querySelector(".riel-lista .crece").textContent, "Inicio");

  // Mutación nueva (menú/aviso agregado dinámicamente) se traduce sola
  const avisoEl = env.doc.createElement("div");
  avisoEl.innerHTML = `<button aria-label="Cerrar">Cancelar</button>`;
  env.doc.body.append(avisoEl);
  assert.equal(avisoEl.querySelector("button").textContent, "Cancel");
  assert.equal(avisoEl.querySelector("button").getAttribute("aria-label"), "Close");

  // Volver a español restaura textos/atributos y apaga el observador (cero costo en español)
  await env.NOM.idioma.aplicar("es");
  assert.equal(env.doc.documentElement.getAttribute("lang"), "es");
  assert.equal(btn.textContent, "Guardar");
  assert.equal(btn.getAttribute("aria-label"), "Guardar");
  assert.equal(btn.getAttribute("title"), "Guardar versión");
  assert.equal(env.observadores.length, 0, "En español el MutationObserver debe estar desconectado");
});

// ================================================================== 4. Rendimiento del MutationObserver
test("4.1 Rendimiento: 2000 mutaciones dentro de un contenteditable cuestan menos de 1 ms por tecla", async () => {
  const env = crearEntorno({ navigator: { languages: ["en-US"] } });
  env.cargar("idioma.js");
  await env.NOM.idioma.aplicar("en");
  assert.equal(env.observadores.length, 1, "En inglés el observador debe estar activo");

  const hoja = env.doc.createElement("div");
  hoja.className = "wd-hoja";
  const cuerpo = env.doc.createElement("div");
  cuerpo.className = "wd-cuerpo";
  cuerpo.setAttribute("contenteditable", "true");
  hoja.append(cuerpo);
  env.doc.body.append(hoja);

  const p = env.doc.createElement("p");
  cuerpo.append(p);

  const t0 = performance.now();
  for (let i = 0; i < 2000; i++) {
    const sp = env.doc.createElement("span");
    sp.append("Guardar ");
    p.append(sp);
  }
  const totalMs = performance.now() - t0;
  const porMutacion = totalMs / 2000;

  assert.equal(p.childNodes[0].textContent, "Guardar ", "No debe traducir lo escrito dentro del contenteditable");
  assert.ok(porMutacion < 1, `Cada mutación debe costar < 1 ms (midió ${porMutacion.toFixed(4)} ms, total ${totalMs.toFixed(2)} ms)`);
  console.log(`  [rendimiento] 2000 mutaciones en contenteditable: ${totalMs.toFixed(2)} ms total (${porMutacion.toFixed(4)} ms/mutación)`);
});

// ================================================================== 5. Detección de idioma del documento
test("5.1 NOM.idioma.detectar(texto) distingue español e inglés y pone lang + spellcheck=true en el documento", () => {
  const env = crearEntorno();
  env.cargar("idioma.js");
  const { detectar, aplicarIdiomaDocumento } = env.NOM.idioma;

  assert.equal(detectar("Hola equipo, este es el informe final de la clase de historia."), "es");
  assert.equal(detectar("¿Dónde guardamos el documento con las gráficas y las tablas?"), "es");
  assert.equal(detectar("Hello team, this is the final report for the history class."), "en");
  assert.equal(detectar("Where did we save the document with the charts and tables?"), "en");
  assert.equal(detectar("The quick brown fox jumps over the lazy dog while working with the team."), "en");
  assert.equal(detectar("En un lugar de la Mancha, de cuyo nombre no quiero acordarme."), "es");
  assert.equal(detectar(""), "es");

  const docEl = env.doc.createElement("div");
  docEl.className = "wd-cuerpo";
  docEl.setAttribute("contenteditable", "true");
  docEl.textContent = "This document explains how our team built the project and tested every screen.";
  const detectado = aplicarIdiomaDocumento(docEl);
  assert.equal(detectado, "en");
  assert.equal(docEl.getAttribute("lang"), "en");
  assert.equal(docEl.getAttribute("spellcheck"), "true");

  docEl.textContent = "Este documento explica cómo nuestro equipo construyó el proyecto y probó cada pantalla.";
  assert.equal(aplicarIdiomaDocumento(docEl), "es");
  assert.equal(docEl.getAttribute("lang"), "es");
  assert.equal(docEl.getAttribute("spellcheck"), "true");
});

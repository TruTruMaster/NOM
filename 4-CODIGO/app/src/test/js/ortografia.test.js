// Pruebas de ortografía (js/vendor/ortografia/nspell-navegador.js, js/ortografia.js e integración).
// Uso: node --test 4-CODIGO/app/src/test/js/ortografia.test.js
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { performance } = require("node:perf_hooks");

const dirStatic = path.join(__dirname, "../../main/resources/static");
const rutaBundle = path.join(dirStatic, "js/vendor/ortografia/nspell-navegador.js");
const rutaOrtografia = path.join(dirStatic, "js/ortografia.js");
const rutaIdioma = path.join(dirStatic, "js/idioma.js");

function crearStorage(inicial) {
  const mapa = new Map(Object.entries(inicial || {}));
  return {
    getItem(k) { return mapa.has(k) ? mapa.get(k) : null; },
    setItem(k, v) { mapa.set(String(k), String(v)); },
    removeItem(k) { mapa.delete(k); },
    _mapa: mapa,
  };
}

function crearEntorno(opc) {
  opc = opc || {};
  const storage = opc.storage || crearStorage();
  const pedidos = [];
  const docMin = {
    readyState: "complete",
    documentElement: { setAttribute() {}, getAttribute() { return null; }, append() {} },
    head: { append() {} },
    body: null,
    addEventListener() {},
    querySelectorAll() { return []; },
  };
  const window = {
    NOM: {
      leer: (k, def) => { const v = storage.getItem(k); return v == null ? def : JSON.parse(v); },
      guardar: (k, v) => { storage.setItem(k, JSON.stringify(v)); return true; },
    },
    addEventListener() {},
    navigator: { language: opc.langSistema || "es-MX", languages: [opc.langSistema || "es-MX"] },
  };

  let fetchFn = opc.fetch || (async (url) => {
    pedidos.push(String(url));
    const ruta = path.join(dirStatic, String(url));
    if (!fs.existsSync(ruta)) return { ok: false, status: 404, async text() { return ""; } };
    const txt = fs.readFileSync(ruta, "utf8");
    return { ok: true, status: 200, async text() { return txt; } };
  });

  const ctx = vm.createContext({
    window,
    globalThis: window,
    document: docMin,
    navigator: window.navigator,
    localStorage: storage,
    fetch: (url) => fetchFn(url),
    setTimeout,
    clearTimeout,
    console,
    Map,
    Set,
    WeakMap,
    Promise,
  });

  vm.runInContext(fs.readFileSync(rutaIdioma, "utf8"), ctx, { filename: "idioma.js" });
  if (opc.cargarBundle !== false) {
    vm.runInContext(fs.readFileSync(rutaBundle, "utf8"), ctx, { filename: "nspell-navegador.js" });
  }
  vm.runInContext(fs.readFileSync(rutaOrtografia, "utf8"), ctx, { filename: "ortografia.js" });

  return {
    window,
    NOM: window.NOM,
    storage,
    pedidos,
    fijarFetch(fn) { fetchFn = fn; },
  };
}

test("español: carga desde disco, palabras bien/mal, sugerencias < 300 ms y caché en memoria", async (t) => {
  const { NOM, pedidos } = crearEntorno();
  assert.ok(NOM.ortografia, "NOM.ortografia existe");
  assert.strictEqual(NOM.ortografia.listo("es"), false, "no debe precargar el diccionario al inicio");

  const t0 = performance.now();
  const cargado = await NOM.ortografia.cargar("es");
  const cargaMs = performance.now() - t0;
  t.diagnostic(`Tiempo de carga diccionario español: ${cargaMs.toFixed(1)} ms`);
  console.log(`Tiempo de carga diccionario español: ${cargaMs.toFixed(1)} ms`);

  assert.ok(cargado, "cargar('es') devuelve instancia");
  assert.strictEqual(NOM.ortografia.listo("es"), true);
  assert.deepStrictEqual(pedidos, [
    "js/vendor/ortografia/es/es.aff",
    "js/vendor/ortografia/es/es.dic",
  ]);

  // Segunda llamada usa memoria sin volver a pedir los archivos.
  await NOM.ortografia.cargar("es");
  assert.strictEqual(pedidos.length, 2, "solo baja los archivos la primera vez");

  assert.strictEqual(NOM.ortografia.revisar("hola", "es"), true);
  assert.strictEqual(NOM.ortografia.revisar("computadora", "es"), true);
  assert.strictEqual(NOM.ortografia.revisar("escribió", "es"), true);

  assert.strictEqual(NOM.ortografia.revisar("ortografia", "es"), false);
  const s0 = performance.now();
  const sugOrt = NOM.ortografia.sugerencias("ortografia", "es");
  const sugOrtMs = performance.now() - s0;
  t.diagnostic(`Sugerencia 'ortografia': ${JSON.stringify(sugOrt)} (${sugOrtMs.toFixed(1)} ms)`);
  assert.ok(sugOrt.includes("ortografía"), `esperaba 'ortografía' en ${JSON.stringify(sugOrt)}`);
  assert.ok(sugOrt.length <= 5, "máximo 5 sugerencias");
  assert.ok(sugOrtMs < 300, `sugerencia tardó ${sugOrtMs.toFixed(1)} ms (>= 300 ms)`);

  assert.strictEqual(NOM.ortografia.revisar("recivir", "es"), false);
  const s1 = performance.now();
  const sugRec = NOM.ortografia.sugerencias("recivir", "es");
  const sugRecMs = performance.now() - s1;
  t.diagnostic(`Sugerencia 'recivir': ${JSON.stringify(sugRec)} (${sugRecMs.toFixed(1)} ms)`);
  assert.ok(sugRec.includes("recibir"), `esperaba 'recibir' en ${JSON.stringify(sugRec)}`);
  assert.ok(sugRec.length <= 5, "máximo 5 sugerencias");
  assert.ok(sugRecMs < 300, `sugerencia tardó ${sugRecMs.toFixed(1)} ms (>= 300 ms)`);
});

test("inglés: 'recieve' mal escrita con sugerencia 'receive' en < 300 ms", async (t) => {
  const { NOM } = crearEntorno();
  const cargado = await NOM.ortografia.cargar("en");
  assert.ok(cargado, "cargar('en') devuelve instancia");

  assert.strictEqual(NOM.ortografia.revisar("receive", "en"), true);
  assert.strictEqual(NOM.ortografia.revisar("recieve", "en"), false);

  const t0 = performance.now();
  const sugs = NOM.ortografia.sugerencias("recieve", "en");
  const ms = performance.now() - t0;
  t.diagnostic(`Sugerencia 'recieve': ${JSON.stringify(sugs)} (${ms.toFixed(1)} ms)`);
  assert.ok(sugs.includes("receive"), `esperaba 'receive' en ${JSON.stringify(sugs)}`);
  assert.ok(sugs.length <= 5, "máximo 5 sugerencias");
  assert.ok(ms < 300, `sugerencia tardó ${ms.toFixed(1)} ms (>= 300 ms)`);
});

test("diccionario personal en localStorage ('nom.ortografia.personal') persiste al recargar", async () => {
  const storage = crearStorage();
  const env1 = crearEntorno({ storage });
  await env1.NOM.ortografia.cargar("es");

  assert.strictEqual(env1.NOM.ortografia.revisar("nomcolab", "es"), false);
  env1.NOM.ortografia.agregar("nomcolab");
  assert.strictEqual(env1.NOM.ortografia.revisar("nomcolab", "es"), true);
  assert.deepStrictEqual(JSON.parse(storage.getItem("nom.ortografia.personal")), ["nomcolab"]);

  // Nueva sesión con el mismo localStorage.
  const env2 = crearEntorno({ storage });
  await env2.NOM.ortografia.cargar("es");
  assert.strictEqual(env2.NOM.ortografia.revisar("nomcolab", "es"), true);
});

test("sin red: devuelve null sin lanzar error y permite reintentar después", async () => {
  const env = crearEntorno({
    fetch: async () => { throw new Error("sin conexión"); },
  });
  const res = await env.NOM.ortografia.cargar("es");
  assert.strictEqual(res, null);
  assert.strictEqual(env.NOM.ortografia.listo("es"), false);
  assert.strictEqual(env.NOM.ortografia.revisar("ortografia", "es"), null);
  assert.deepStrictEqual(Array.from(env.NOM.ortografia.sugerencias("ortografia", "es")), []);

  // Al volver la red, vuelve a intentar y carga bien.
  env.fijarFetch(async (url) => {
    const txt = fs.readFileSync(path.join(dirStatic, String(url)), "utf8");
    return { ok: true, status: 200, async text() { return txt; } };
  });
  const ok = await env.NOM.ortografia.cargar("es");
  assert.ok(ok, "debe poder cargar al recuperarse la red");
  assert.strictEqual(env.NOM.ortografia.revisar("hola", "es"), true);
});

test("carga perezosa del motor nspell-navegador.js cuando aún no está en window", async () => {
  const scriptsPedidos = [];
  const docConScript = {
    readyState: "complete",
    documentElement: { setAttribute() {}, getAttribute() { return null; }, append() {} },
    head: {
      append(s) {
        scriptsPedidos.push(s.src);
        vm.runInContext(fs.readFileSync(path.join(dirStatic, s.src), "utf8"), ctx, { filename: s.src });
        if (s.onload) s.onload();
      },
    },
    body: null,
    createElement(tag) { return { tagName: tag.toUpperCase(), src: "", onload: null, onerror: null }; },
    addEventListener() {},
    querySelectorAll() { return []; },
  };
  const window = { NOM: {}, addEventListener() {}, navigator: { language: "es-MX", languages: ["es-MX"] } };
  const ctx = vm.createContext({
    window,
    globalThis: window,
    document: docConScript,
    navigator: window.navigator,
    localStorage: crearStorage(),
    fetch: async (url) => ({ ok: true, status: 200, async text() { return fs.readFileSync(path.join(dirStatic, String(url)), "utf8"); } }),
    setTimeout,
    clearTimeout,
    console,
    Map,
    Set,
    WeakMap,
    Promise,
  });
  vm.runInContext(fs.readFileSync(rutaIdioma, "utf8"), ctx, { filename: "idioma.js" });
  vm.runInContext(fs.readFileSync(rutaOrtografia, "utf8"), ctx, { filename: "ortografia.js" });

  assert.strictEqual(scriptsPedidos.length, 0, "no debe cargar el motor al inicio");
  const sp = await window.NOM.ortografia.cargar("en");
  assert.ok(sp, "carga motor + diccionario en la primera llamada");
  assert.deepStrictEqual(scriptsPedidos, ["js/vendor/ortografia/nspell-navegador.js"]);
  assert.strictEqual(window.NOM.ortografia.revisar("receive", "en"), true);
  assert.strictEqual(window.NOM.ortografia.revisar("recieve", "en"), false);
});

test("idioma: detecta del documento con NOM.idioma.detectar o usa el de la interfaz", () => {
  const envEs = crearEntorno({ langSistema: "es-MX" });
  assert.strictEqual(envEs.NOM.ortografia.idioma("The document is ready for the team"), "en");
  assert.strictEqual(envEs.NOM.ortografia.idioma("El documento está listo para el equipo"), "es");
  assert.strictEqual(envEs.NOM.ortografia.idioma(""), "es");

  const envEn = crearEntorno({ langSistema: "en-US" });
  assert.strictEqual(envEn.NOM.ortografia.idioma(""), "en");
});

test("integración: index.html, sw.js, licencias y modo-texto.js", () => {
  const html = fs.readFileSync(path.join(dirStatic, "index.html"), "utf8");
  const posIdioma = html.indexOf('<script src="js/idioma.js"></script>');
  const posOrt = html.indexOf('<script src="js/ortografia.js"></script>');
  assert.ok(posIdioma > 0 && posOrt > posIdioma, "index.html carga js/ortografia.js después de js/idioma.js");

  const sw = fs.readFileSync(path.join(dirStatic, "sw.js"), "utf8");
  assert.ok(sw.includes('"js/ortografia.js"'), "sw.js incluye js/ortografia.js en BASE");
  assert.ok(!sw.includes("es.aff") && !sw.includes("es.dic") && !sw.includes("en.aff") && !sw.includes("en.dic"), "sw.js no precarga diccionarios");

  const lic = fs.readFileSync(path.join(dirStatic, "legal/licencias/index.html"), "utf8");
  assert.ok(lic.includes("nspell") && lic.includes("MIT"), "licencias incluye nspell (MIT)");
  assert.ok(lic.includes("LibreOffice") && lic.includes("GPL-3.0") && lic.includes("LGPL-3.0") && lic.includes("MPL-1.1"), "licencias incluye diccionario español");
  assert.ok(lic.includes("SCOWL") && lic.includes("BSD"), "licencias incluye diccionario inglés");

  const modoTexto = fs.readFileSync(path.join(dirStatic, "js/modo-texto.js"), "utf8");
  assert.ok(modoTexto.includes("NOM.ortografia"), "modo-texto.js usa NOM.ortografia");
  assert.ok(modoTexto.includes("Ortografía…"), "modo-texto.js muestra 'Ortografía…' antes de cargar");
  assert.ok(modoTexto.includes("Agregar al diccionario"), "modo-texto.js ofrece 'Agregar al diccionario'");
  assert.ok(modoTexto.includes("Sin conexión: el diccionario no se pudo bajar"), "modo-texto.js avisa sin conexión");
});

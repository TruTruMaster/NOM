// La ruedita no pierde opciones: con más de 8, las primeras 7 y «Más» con el resto (antes se cortaba en 8).
//   node --test 4-CODIGO/app/src/test/js/rueda-paginar.test.js
const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path"), vm = require("vm");
const fuente = fs.readFileSync(path.join(__dirname, "../../main/resources/static/js/rueda.js"), "utf8");
const doc = { addEventListener() {}, removeEventListener() {}, querySelector: () => null, querySelectorAll: () => [], documentElement: { dataset: {} }, body: { append() {} } };
const window = { NOM: { leer: () => null, guardar() {} }, addEventListener() {}, matchMedia: () => ({ matches: false, addEventListener() {} }), innerWidth: 1000, innerHeight: 800 };
const ctx = vm.createContext(Object.assign({ window, document: doc, navigator: {}, setTimeout, clearTimeout, console, matchMedia: window.matchMedia, innerWidth: 1000, innerHeight: 800, addEventListener() {} }, {}));
vm.runInContext(fuente, ctx);
const paginar = window.NOM.rueda && window.NOM.rueda._paginar;

test("20 opciones: 7 a la vista y «Más» con las otras 13, que a su vez se reparten", () => {
  assert.ok(paginar, "rueda.js expone _paginar");
  const items = Array.from({ length: 20 }, (_, i) => ({ texto: "op" + (i + 1) }));
  const a = paginar(items);
  assert.strictEqual(a.length, 8);
  assert.strictEqual(a[7].texto, "Más");
  assert.strictEqual(a[7].sub.length, 13);
  const b = paginar(a[7].sub);
  assert.strictEqual(b.length, 8); assert.strictEqual(b[7].sub.length, 6);
  // todas siguen alcanzables y en orden
  const todas = [...a.slice(0, 7), ...b.slice(0, 7), ...paginar(b[7].sub)].map((x) => x.texto);
  assert.strictEqual(JSON.stringify(todas), JSON.stringify(items.map((x) => x.texto)));
});

test("8 o menos: igual que antes", () => {
  const items = Array.from({ length: 8 }, (_, i) => ({ texto: "op" + i }));
  assert.strictEqual(paginar(items).length, 8);
  assert.strictEqual(paginar([null, { texto: "a" }]).length, 1);
});

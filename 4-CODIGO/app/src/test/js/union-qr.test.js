// El QR para unirse lleva la dirección del teléfono que comparte aunque la app haya trabajado antes con la nube
// (api.lan quedaba en false: lo último que dijo NOM en Internet). Pasó en el A55 el 27-nov.
//   node --test 4-CODIGO/app/src/test/js/union-qr.test.js
const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path"), vm = require("vm");
const fuente = fs.readFileSync(path.join(__dirname, "../../main/resources/static/js/union.js"), "utf8");

function cargar(api, anfitrion) {
  const window = { NOM: { api, anfitrion, leer: () => null, guardar: () => {} }, location: { hostname: "appassets.androidplatform.net", origin: "https://appassets.androidplatform.net", search: "", hash: "" }, addEventListener() {} };
  const ctx = vm.createContext({ window, document: { addEventListener() {}, querySelector: () => null }, location: window.location, URLSearchParams, setTimeout, console });
  vm.runInContext(fuente, ctx);
  return window.NOM.union;
}

test("teléfono que comparte: el QR usa sus direcciones aunque la app viniera de la nube", () => {
  const api = { empaquetada: true, servidor: () => "http://127.0.0.1:8080", lan: false };
  const anfitrion = { esLocal: () => true, estado: () => ({ direcciones: ["http://192.168.0.58:8080"] }) };
  assert.strictEqual(JSON.stringify(cargar(api, anfitrion).bases(null)), JSON.stringify(["http://192.168.0.58:8080"]));
});

test("en la nube (lan false) y sin teléfono que comparte: sin direcciones de la red local", () => {
  const api = { empaquetada: false, servidor: () => "", lan: false };
  const u = cargar(api, { esLocal: () => false, estado: () => ({ direcciones: ["http://10.0.0.2:8080"] }) });
  assert.strictEqual(JSON.stringify(u.bases({ direcciones: ["10.0.0.9"], puerto: 8080 })), JSON.stringify(["https://appassets.androidplatform.net"]));
});

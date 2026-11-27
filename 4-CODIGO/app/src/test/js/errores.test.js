// Cada error dice qué pasó y qué hacer (js/errores.js); el servidor caído se trata como sin red.
//   node --test 4-CODIGO/app/src/test/js/errores.test.js
const test = require("node:test"), assert = require("node:assert");
global.window = {};
const E = require("../../main/resources/static/js/errores.js");

test("cada estado tiene un mensaje propio, sin «error NNN»", () => {
  for (const s of [0, 400, 401, 403, 404, 408, 409, 413, 415, 423, 429, 500, 501, 502, 503, 504, 520, 530, 418]) {
    const m = E.mensaje(s);
    assert.ok(m.length > 25, s + ": " + m);
    assert.ok(!/error \d{3}/i.test(m), s + ": " + m);
  }
  assert.match(E.mensaje(413), /grande/);
  assert.match(E.mensaje(429), /espera/i);
  assert.match(E.mensaje(530), /no responde/);
});

test("servidor caído: solo 502-504 y 520-530 sin mensaje de NOM", () => {
  assert.ok(E.caido(502, null) && E.caido(530, null) && E.caido(504, null) && E.caido(503, null));
  assert.ok(!E.caido(503, { error: "El servidor está recibiendo muchos archivos a la vez" }));   // 503 propio de NOM
  assert.ok(!E.caido(500, null) && !E.caido(404, null) && !E.caido(531, null) && !E.caido(519, null));
});

// Dónde parte NOM cada palabra (js/guiones.js con sus patrones), para comparar con lo que sacó guiones_word.py.
//   node guiones_nom.js guiones-word.json > comparacion.json
const fs = require("fs"), path = require("path"), vm = require("vm");
const JS = path.resolve(__dirname, "../../../app/src/main/resources/static/js");
const ctx = { window: {}, document: {}, console };
ctx.window.NOM = {};
ctx.window.window = ctx.window;
vm.createContext(ctx);
for (const f of ["guiones-es.js", "guiones-en-us.js", "guiones.js"]) vm.runInContext(fs.readFileSync(path.join(JS, f), "utf8").replace(/^\s*window\./gm, "window."), ctx, { filename: f });
const G = ctx.window.NOM.guiones;
const word = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const out = {};
let iguales = 0, total = 0;
for (const idioma of Object.keys(word)) {
  out[idioma] = {};
  for (const [w, pts] of Object.entries(word[idioma])) {
    const nom = G.puntos(w, idioma);
    const ok = JSON.stringify(nom) === JSON.stringify(pts);
    total++; if (ok) iguales++;
    out[idioma][w] = { word: pts, nom, ok };
  }
}
console.log(JSON.stringify({ iguales, total, palabras: out }, null, 0));

// Pruebas de accesibilidad WCAG 2.2 AA y marco mexicano (Carril B):
// verifica documentación (docs/ACCESIBILIDAD.md), reglas CSS nuevas en nom.css y editor.css,
// comportamiento en accesibilidad.js, rueda.js, experiencia.js y vista-espacios.js, y cobertura de prueba_axe.py.
// Uso: node --test 4-CODIGO/app/src/test/js/accesibilidad-wcag.test.js
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const dirCodigo = path.join(__dirname, "../../../..");
const dirStatic = path.join(__dirname, "../../main/resources/static");
const dirCss = path.join(dirStatic, "css");
const dirJs = path.join(dirStatic, "js");
const rutaDocAcc = path.join(dirCodigo, "docs/ACCESIBILIDAD.md");
const rutaAxe = path.join(dirCodigo, "pruebas/interfaz/prueba_axe.py");

const leer = (p) => fs.readFileSync(p, "utf8");

/** Extrae el bloque completo de una regla @media (maneja llaves anidadas). */
function extraerBloquesMedia(css, filtro) {
  const bloques = [];
  let idx = 0;
  while ((idx = css.indexOf("@media", idx)) !== -1) {
    const abre = css.indexOf("{", idx);
    if (abre === -1) break;
    const cabecera = css.slice(idx, abre);
    let nivel = 1, pos = abre + 1;
    while (pos < css.length && nivel > 0) {
      if (css[pos] === "{") nivel++;
      else if (css[pos] === "}") nivel--;
      pos++;
    }
    if (filtro.test(cabecera)) bloques.push(css.slice(idx, pos));
    idx = pos;
  }
  return bloques.join("\n");
}

// ---------------------------------------------------------------- 1. Documentación normativa
test("docs/ACCESIBILIDAD.md existe y cubre WCAG 2.2 A/AA nuevos y ley mexicana con fuentes primarias", () => {
  assert.equal(fs.existsSync(rutaDocAcc), true, "Falta 4-CODIGO/docs/ACCESIBILIDAD.md");
  const md = leer(rutaDocAcc);
  assert.match(md, /https:\/\/www\.w3\.org\/TR\/WCAG22/i, "Debe citar w3.org/TR/WCAG22");
  for (const sc of ["2.4.11", "2.5.7", "2.5.8", "3.2.6", "3.3.7", "3.3.8"]) {
    assert.ok(md.includes(sc), `Debe documentar el criterio WCAG 2.2 ${sc}`);
  }
  assert.match(md, /Ley General para la Inclusión de las Personas con Discapacidad/i);
  assert.match(md, /Art[ií]culo\s+32/i);
  assert.match(md, /NMX-R-050/i);
  assert.match(md, /espacios construidos|f[ií]sic/i, "Debe aclarar que la NMX-R-050 aplica a espacios físicos, no a web");
  assert.match(md, /NMX-R-099|Disposiciones/i, "Debe citar el marco mexicano de TIC/web");
  assert.match(md, /(nom\.css|editor\.css|accesibilidad\.js|rueda\.js|nucleo\.js|vista-cuentas\.js):\d+/, "Debe citar archivo:línea reales");
});

// ---------------------------------------------------------------- 2. Foco visible y no tapado (WCAG 2.4.11 / 2.4.7)
test("nom.css y editor.css definen foco visible con outline y scroll-padding para barras fijas", () => {
  const nom = leer(path.join(dirCss, "nom.css"));
  const ed = leer(path.join(dirCss, "editor.css"));
  assert.match(nom, /scroll-padding-top\s*:/, "nom.css debe definir scroll-padding-top");
  assert.match(nom, /scroll-padding-bottom\s*:/, "nom.css debe definir scroll-padding-bottom");
  assert.match(nom, /:focus-visible\s*\{[^}]*outline\s*:\s*2px\s+solid/s, "nom.css :focus-visible debe llevar outline visible");
  assert.match(ed, /scroll-padding-top\s*:/, "editor.css debe definir scroll-padding en áreas desplazables");
  assert.match(ed, /\.rueda-item:focus-visible\s*\{[^}]*outline\s*:\s*2px\s+solid/s, "editor.css debe restaurar foco visible en .rueda-item");
  assert.match(ed, /\.ao-ef:focus-visible\s*\{[^}]*outline\s*:\s*2px\s+solid\s+var\(--acento\)/s, "editor.css debe corregir outline de .ao-ef:focus-visible");
  assert.match(ed, /\.salir-enfoque:focus-visible[\s\S]*?opacity\s*:\s*1/s, "El botón salir de enfoque debe verse completo al enfocarse con teclado");
  assert.match(ed, /\.reproductor\.quieto\s+\.salir:focus-visible[\s\S]*?opacity\s*:\s*1/s, "El botón salir al presentar debe verse al enfocarse con teclado");
});

// ---------------------------------------------------------------- 3. Tamaño mínimo de objetivos (WCAG 2.5.8)
test("nom.css y editor.css garantizan objetivos mínimos de 24 px en escritorio y 44 px en táctil", () => {
  const nom = leer(path.join(dirCss, "nom.css"));
  const ed = leer(path.join(dirCss, "editor.css"));
  assert.match(nom, /min-width\s*:\s*24px;\s*min-height\s*:\s*24px/, "nom.css debe fijar mínimo 24x24 px");
  assert.match(ed, /\.casilla\s*\{[^}]*min-width\s*:\s*24px[^}]*min-height\s*:\s*24px/s);
  assert.match(ed, /\.lz-muestras\s+button\s*\{[^}]*min-width\s*:\s*24px[^}]*min-height\s*:\s*24px/s);
  assert.match(ed, /\.desplegable-celda\s*\{[^}]*min-width\s*:\s*24px[^}]*min-height\s*:\s*24px/s);
  assert.match(ed, /\.esq-btn\s*\{[^}]*min-width\s*:\s*24px[^}]*min-height\s*:\s*24px/s);
  assert.match(ed, /\.wd-vistas\s*\{[^}]*transform\s*:\s*none/s, "wd-vistas no debe encogerse por debajo de 24 px");
  assert.match(ed, /\.rueda-item\s*\{[^}]*min-height\s*:\s*44px/s, "Los elementos de la ruedita deben cumplir objetivo mínimo");

  const coarseNom = extraerBloquesMedia(nom, /pointer:\s*coarse/);
  const coarseEd = extraerBloquesMedia(ed, /pointer:\s*coarse/);
  assert.match(coarseNom, /\.segmentado\s+button[\s\S]*?min-height\s*:\s*44px/);
  assert.match(coarseNom, /\.btn\.chico[\s\S]*?min-height\s*:\s*44px/);
  assert.match(coarseEd, /\.ed-herramientas\s+\.icono-btn[\s\S]*?min-height\s*:\s*44px/);
  assert.match(coarseEd, /\.zoom-ctl\s+\.icono-btn[\s\S]*?min-height\s*:\s*44px/);
  assert.match(coarseEd, /\.wd-muestras\s+button[\s\S]*?min-height\s*:\s*44px/);
});

// ---------------------------------------------------------------- 4. Colores forzados y prefers-contrast
test("nom.css y editor.css soportan forced-colors: active y prefers-contrast: more sin ocultar controles", () => {
  const nom = leer(path.join(dirCss, "nom.css"));
  const ed = leer(path.join(dirCss, "editor.css"));

  const fcNom = extraerBloquesMedia(nom, /forced-colors:\s*active/);
  assert.ok(fcNom.length > 200, "nom.css debe tener reglas @media (forced-colors: active)");
  for (const token of ["ButtonText", "ButtonFace", "Highlight", "HighlightText", "CanvasText", "currentColor"]) {
    assert.ok(fcNom.includes(token), `forced-colors en nom.css debe usar ${token}`);
  }
  assert.match(fcNom, /\.btn\.primario/);
  assert.match(fcNom, /\.segmentado\s+button\[aria-pressed="true"\]/);
  assert.match(fcNom, /\.interruptor\s+span/);
  assert.match(fcNom, /:focus-visible/);

  const fcEd = extraerBloquesMedia(ed, /forced-colors:\s*active/);
  assert.ok(fcEd.length > 200, "editor.css debe tener reglas @media (forced-colors: active)");
  assert.match(fcEd, /\.rueda-fondo/);
  assert.match(fcEd, /\.rueda-sector\.activo/);
  assert.match(fcEd, /\.rueda-item\.activo/);
  assert.match(fcEd, /\.rueda-centro/);
  assert.match(fcEd, /\.rueda-lista-item/);
  assert.match(fcEd, /\.zoom-ctl[\s\S]*?forced-color-adjust\s*:\s*auto/, "Controles dentro de .escenario deben volver a forced-color-adjust: auto");

  const pcNom = extraerBloquesMedia(nom, /prefers-contrast:\s*more/);
  const pcEd = extraerBloquesMedia(ed, /prefers-contrast:\s*more/);
  assert.ok(pcNom.length > 100, "nom.css debe tener reglas @media (prefers-contrast: more)");
  assert.ok(pcEd.length > 80, "editor.css debe tener reglas @media (prefers-contrast: more)");
});

// ---------------------------------------------------------------- 5. Botones que no desaparecen (letra, oscuro, tarjetas, 320 px)
test("Ningún botón desaparece con letra grande/muy grande, tema oscuro, táctil, teclado o zoom 400 % (320 px)", () => {
  const nom = leer(path.join(dirCss, "nom.css"));
  const ed = leer(path.join(dirCss, "editor.css"));

  // Tarjetas: acciones visibles en táctil y con teclado
  const tactilNom = extraerBloquesMedia(nom, /hover:\s*none|pointer:\s*coarse/);
  assert.match(tactilNom, /\.doc-tarjeta\s+\.acciones\s*\{[^}]*opacity\s*:\s*1/s, "Las acciones de .doc-tarjeta deben verse siempre en pantallas táctiles");
  assert.match(nom, /:root\[data-teclado="1"\]\s+\.doc-tarjeta\s+\.acciones[\s\S]*?opacity\s*:\s*1/s, "Las acciones de .doc-tarjeta deben verse al usar el teclado");

  // Reflow 320 px (zoom 400 %)
  const angostoNom = extraerBloquesMedia(nom, /max-width:\s*(360|380|400)px/);
  assert.ok(angostoNom.length > 120, "nom.css debe incluir reglas de reflow para ~320 px");
  assert.match(angostoNom, /\.hoja-pie\s+\.btn[\s\S]*?min-width\s*:\s*0/s, "Botones del pie de Ajustes no deben desbordar a 320 px");
  assert.match(angostoNom, /\.segmentado[\s\S]*?flex-wrap\s*:\s*wrap/s);

  // Letra grande / muy grande sin recortar botones ni barras
  assert.match(nom, /:root\[data-letra\]\s+\.btn\s*\{[^}]*white-space\s*:\s*normal/s);
  assert.match(nom, /:root\[data-letra\]\s+\.hoja-pie\s*\{[^}]*flex-wrap\s*:\s*wrap/s);
  assert.match(nom, /:root\[data-letra\]\s+\.segmentado\s*\{[^}]*flex-wrap\s*:\s*wrap/s);
  assert.match(ed, /:root\[data-letra\]\s+\.ed-barra\s*\{[^}]*flex-wrap\s*:\s*wrap/s);
  assert.match(ed, /\.ed-barra\s*\{[^}]*height\s*:\s*auto/s);

  // Tema oscuro en modo de lectura, proyector y casillas marcadas
  assert.match(ed, /:root\[data-tema="oscuro"\]\s+\.wd-lectura\s*\{[^}]*background\s*:\s*var\(--fondo\)/s);
  assert.match(ed, /:root\[data-tema="claro"\]\s+\.wd-lectura\s*\{[^}]*background\s*:\s*var\(--sup\)/s);
  assert.match(ed, /\.proy-espera\s+\.entrada\s*\{[^}]*background\s*:\s*var\(--sup-2\)/s);
  assert.match(ed, /:root\[data-tema="oscuro"\]\s+\.casilla\.marcada\s*\{[^}]*color\s*:\s*var\(--sobre-acento\)/s);

  // Calidad baja: segmentado activo sin box-shadow inválido
  assert.match(nom, /:root\[data-calidad="baja"\]\s+\.segmentado\s+button\[aria-pressed="true"\]\s*\{[^}]*box-shadow\s*:\s*none/s);

  // Sin degradados en las secciones nuevas de accesibilidad (DISENO-UI.md §1)
  const marcaNom = nom.indexOf("Accesibilidad completa (WCAG 2.2");
  const marcaEd = ed.indexOf("Accesibilidad completa (WCAG 2.2");
  assert.ok(marcaNom > 0 && marcaEd > 0, "Las reglas nuevas deben estar al final de nom.css y editor.css");
  assert.doesNotMatch(nom.slice(marcaNom), /gradient\(/i, "Sin degradados en reglas nuevas de nom.css");
  assert.doesNotMatch(ed.slice(marcaEd), /gradient\(/i, "Sin degradados en reglas nuevas de editor.css");
});

// ---------------------------------------------------------------- 6. Comportamiento JS (accesibilidad.js, rueda.js, experiencia.js, vista-espacios.js)
function entornoMinimo() {
  const listeners = new Map();
  const addEv = (t, f) => { (listeners.get(t) || listeners.set(t, []).get(t)).push(f); };
  const remEv = (t, f) => { if (listeners.has(t)) listeners.set(t, listeners.get(t).filter((x) => x !== f)); };
  const emitir = (t, ev) => { for (const f of (listeners.get(t) || []).slice()) f(ev); };
  let scrollOpts = null;
  const parrafoMock = {
    classList: { s: new Set(), add(c) { this.s.add(c); }, remove(c) { this.s.delete(c); } },
    scrollIntoView(o) { scrollOpts = o; },
  };
  const docEl = { dataset: {} };
  const doc = {
    documentElement: docEl,
    body: { style: {}, append() {} },
    activeElement: null,
    addEventListener: addEv,
    removeEventListener: remEv,
    querySelectorAll(sel) { return sel === ".wd-cuerpo > *" ? [parrafoMock] : []; },
    querySelector() { return null; },
  };
  const win = {
    NOM: { html: () => ({}), esc: (s) => String(s || ""), ico: () => "", $: () => null, $$: () => [] },
    document: doc,
    addEventListener: addEv,
    removeEventListener: remEv,
    matchMedia: () => ({ matches: false, addEventListener() {} }),
    setTimeout: () => 1,
    clearTimeout() {},
  };
  win.window = win;
  return { win, doc, docEl, emitir, getScrollOpts: () => scrollOpts };
}

test("accesibilidad.js detecta navegación por teclado y respeta movimiento reducido al desplazarse", () => {
  const env = entornoMinimo();
  const ctx = vm.createContext(env.win);
  vm.runInContext(leer(path.join(dirJs, "accesibilidad.js")), ctx);
  const acc = env.win.NOM.accesibilidad;

  // Contraste WCAG: azul marca y tema oscuro
  assert.ok(acc.contraste("#2B2952", "#FFFFFF") > 13);
  assert.ok(acc.contraste("#ECEDF2", "#14151A") > 15);

  // Al presionar Tab se marca data-teclado="1" para que se vean las acciones en tarjetas
  env.emitir("keydown", { key: "Tab" });
  assert.equal(env.docEl.dataset.teclado, "1", "Tab debe activar data-teclado='1'");
  env.emitir("pointerdown", { pointerType: "mouse" });
  assert.equal(env.docEl.dataset.teclado, undefined, "El puntero de mouse limpia data-teclado");

  // Con movimiento reducido, ir() no fuerza desplazamiento suave
  env.docEl.dataset.movimiento = "reducido";
  assert.equal(typeof acc.ir, "function", "accesibilidad.js debe exponer ir() o usarlo con movimiento reducido");
  acc.ir({ doc: {} }, { parrafo: 0 });
  assert.equal(env.getScrollOpts().behavior, "auto", "Con movimiento reducido, scrollIntoView debe usar behavior: 'auto'");
});

test("rueda.js restaura el foco al cerrar, salta opciones deshabilitadas con teclado y anuncia el centro", () => {
  const src = leer(path.join(dirJs, "rueda.js"));
  assert.match(src, /aria-live="polite"/, "El centro de la ruedita debe anunciar la acción activa");
  assert.match(src, /anterior\s*=.*document\.activeElement/, "rueda.js debe guardar el foco anterior para restaurarlo al cerrar");
  assert.match(src, /anterior\.focus/, "rueda.js debe devolver el foco al cerrar");
  assert.match(src, /function siguienteHabil\([\s\S]*?abrir\(/s, "El teclado de la ruedita debe saltar elementos deshabilitados");
});

test("experiencia.js (paleta) usa aria-activedescendant y vista-espacios.js (Ajustes) etiqueta grupos segmentados", () => {
  const exp = leer(path.join(dirJs, "experiencia.js"));
  assert.match(exp, /aria-activedescendant/, "La paleta de comandos debe actualizar aria-activedescendant");

  const ve = leer(path.join(dirJs, "vista-espacios.js"));
  const ini = ve.indexOf("async function preferencias()");
  const pref = ve.slice(ini);
  for (const attr of ["data-tema", "data-idioma", "data-letra", "data-menu-rapido", "data-haptica-nivel"]) {
    const re = new RegExp(`role="group"[^>]*aria-label="[^"]+"[^>]*${attr}|${attr}[^>]*role="group"[^>]*aria-label="[^"]+"`);
    assert.match(pref, re, `El control segmentado ${attr} en Ajustes debe llevar role="group" y aria-label`);
  }
});

// ---------------------------------------------------------------- 7. prueba_axe.py ampliada
test("prueba_axe.py revisa los 4 editores con contenido, Ajustes, tema oscuro y colores forzados", () => {
  const py = leer(rutaAxe);
  assert.match(py, /forced_colors\s*=\s*"active"/, "prueba_axe.py debe probar forced_colors='active'");
  assert.match(py, /color_scheme\s*=\s*"dark"/, "prueba_axe.py debe probar color_scheme='dark'");
  assert.match(py, /tabla/i, "prueba_axe.py debe insertar una tabla en Texto");
  assert.match(py, /imagen|insertarImagen/i, "prueba_axe.py debe insertar una imagen en Texto");
  assert.match(py, /conector/i, "prueba_axe.py debe crear conectores en Diagrama");
  assert.match(py, /forma/i, "prueba_axe.py debe crear figuras en Presentación");
  assert.match(py, /SUMA|c:A1|c:B2/i, "prueba_axe.py debe llenar datos en Hoja");
  assert.match(py, /Ajustes.*oscuro|oscuro.*Ajustes|Ajustes.*forzad|forzad.*Ajustes/is, "prueba_axe.py debe revisar Ajustes en tema oscuro y colores forzados");
});

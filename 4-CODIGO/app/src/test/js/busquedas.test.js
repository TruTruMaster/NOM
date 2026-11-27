// Pruebas de BUSCARV/VLOOKUP, BUSCARH/HLOOKUP, COINCIDIR/MATCH, BUSCARX/XLOOKUP y COINCIDIRX/XMATCH
// comparadas con los resultados documentados de Excel.
// Uso: node 4-CODIGO/app/src/test/js/busquedas.test.js
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const dirJs = process.env.NOM_JS_DIR || path.join(__dirname, "../../main/resources/static/js");
const NOM = {};
const sandbox = { window: { NOM }, NOM, console, Intl };
const ctx = vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(dirJs, "formulas.js"), "utf8"), ctx, { filename: "formulas.js" });
const F = NOM.formulas;

function evalCelda(celdas, formula) {
  const mapa = new Map();
  let maxC = 1, maxF = 1;
  for (const [ref, val] of Object.entries(celdas)) {
    const p = F.parseRef(ref);
    if (!p) throw new Error("Referencia inválida en prueba: " + ref);
    mapa.set(p.c + "," + p.f, val);
    if (p.c + 1 > maxC) maxC = p.c + 1;
    if (p.f + 1 > maxF) maxF = p.f + 1;
  }
  const destC = maxC + 2, destF = 0;
  mapa.set(destC + "," + destF, formula);
  const ev = F.evaluador((c, f) => {
    const v = mapa.get(c + "," + f);
    return v === undefined ? "" : v;
  }, { filas: maxF + 5, cols: destC + 2 });
  const v = ev.valor(destC, destF);
  if (v instanceof F.ErrorF) return v.codigo;
  if (v instanceof F.FechaF) return F.textoFecha(v);
  return v;
}

// Tabla de casos con el resultado exacto de Excel:
// { id, desc, celdas, formula, excel }
const CASOS = [
  // --- BUSCARV / VLOOKUP coincidencia aproximada (VERDADERO u omitido) ---
  {
    id: "BV-01",
    desc: "BUSCARV aprox (omitido): mayor valor <= buscado en tabla ascendente",
    celdas: { A1: "0", B1: "F", A2: "60", B2: "D", A3: "70", B3: "C", A4: "80", B4: "B", A5: "90", B5: "A" },
    formula: "=BUSCARV(75; A1:B5; 2)",
    excel: "C",
  },
  {
    id: "BV-02",
    desc: "VLOOKUP aprox (TRUE): coincidencia exacta con duplicados devuelve la última fila",
    celdas: { A1: "10", B1: "primero", A2: "20", B2: "d1", A3: "20", B3: "d2", A4: "20", B4: "d3", A5: "30", B5: "fin" },
    formula: "=VLOOKUP(20, A1:B5, 2, TRUE)",
    excel: "d3",
  },
  {
    id: "BV-03",
    desc: "BUSCARV aprox (VERDADERO): buscado menor que el primero → #N/A",
    celdas: { A1: "10", B1: "a", A2: "20", B2: "b", A3: "30", B3: "c" },
    formula: "=BUSCARV(5; A1:B3; 2; VERDADERO)",
    excel: "#N/A",
  },
  {
    id: "BV-04",
    desc: "BUSCARV aprox (VERDADERO): buscado mayor que el último → última fila",
    celdas: { A1: "10", B1: "a", A2: "20", B2: "b", A3: "30", B3: "c" },
    formula: "=BUSCARV(999; A1:B3; 2; VERDADERO)",
    excel: "c",
  },
  {
    id: "BV-05",
    desc: "BUSCARV aprox (VERDADERO) sobre texto con filas vacías al final (rango sobrado) salta vacíos",
    celdas: { A1: "Ana", B1: "1", A2: "Beto", B2: "2", A3: "Carlos", B3: "3", B4: "vacio4", B5: "vacio5" },
    formula: '=BUSCARV("Diego"; A1:B5; 2; VERDADERO)',
    excel: 3,
  },
  {
    id: "BV-06",
    desc: "BUSCARV aprox (VERDADERO) sobre texto con vacío al inicio y buscado menor al primer texto → #N/A",
    celdas: { B1: "vacio1", A2: "Beto", B2: "2", A3: "Carlos", B3: "3" },
    formula: '=BUSCARV("Ana"; A1:B3; 2; VERDADERO)',
    excel: "#N/A",
  },
  {
    id: "BV-07",
    desc: "BUSCARV aprox (VERDADERO) sobre texto con vacío en medio salta el vacío",
    celdas: { A1: "Ana", B1: "1", B2: "vacio", A3: "Carlos", B3: "3", A4: "Elena", B4: "4" },
    formula: '=BUSCARV("Beto"; A1:B4; 2; VERDADERO)',
    excel: 1,
  },
  {
    id: "BV-08",
    desc: "BUSCARV aprox (VERDADERO): texto sin distinguir mayúsculas/minúsculas",
    celdas: { A1: "alfa", B1: "1", A2: "BETA", B2: "2", A3: "Gama", B3: "3" },
    formula: '=BUSCARV("beta"; A1:B3; 2; VERDADERO)',
    excel: 2,
  },
  {
    id: "BV-09",
    desc: "BUSCARV aprox: números y texto no se comparan (texto buscado en rango mixto no devuelve número)",
    celdas: { A1: "10", B1: "n1", A2: "20", B2: "n2", A3: "Beto", B3: "t1", A4: "Carlos", B4: "t2" },
    formula: '=BUSCARV("Ana"; A1:B4; 2; VERDADERO)',
    excel: "#N/A",
  },
  {
    id: "BV-10",
    desc: "BUSCARV aprox: en rango mixto ordenado (números luego texto), busca texto y salta números",
    celdas: { A1: "10", B1: "n1", A2: "20", B2: "n2", A3: "Ana", B3: "t1", A4: "Carlos", B4: "t2" },
    formula: '=BUSCARV("Beto"; A1:B4; 2; VERDADERO)',
    excel: "t1",
  },
  {
    id: "BV-11",
    desc: "BUSCARV aprox: en rango mixto ordenado (números luego texto), busca número > último número devuelve último número (no texto)",
    celdas: { A1: "10", B1: "n1", A2: "20", B2: "n2", A3: "Ana", B3: "t1", A4: "Carlos", B4: "t2" },
    formula: "=BUSCARV(99; A1:B4; 2; VERDADERO)",
    excel: "n2",
  },
  {
    id: "BV-12",
    desc: "BUSCARV aprox: búsqueda binaria real (con dato fuera de orden al inicio, pivote medio sí alcanza 20)",
    celdas: { A1: "50", B1: "mal", A2: "10", B2: "diez", A3: "20", B3: "veinte", A4: "30", B4: "treinta", A5: "40", B5: "cuarenta" },
    formula: "=BUSCARV(25; A1:B5; 2; VERDADERO)",
    excel: "veinte",
  },
  {
    id: "BV-13",
    desc: "BUSCARV aprox: tabla con encabezado de texto en fila 1 y un dato numérico en fila 2",
    celdas: { A1: "Clave", B1: "Valor", A2: "10", B2: "diez" },
    formula: "=BUSCARV(15; A1:B2; 2; VERDADERO)",
    excel: "diez",
  },
  {
    id: "BV-14",
    desc: "BUSCARV aprox: 4.º argumento vacío explícito tras separador (=BUSCARV(x;tabla;2;)) actúa como FALSO (exacta)",
    celdas: { A1: "10", B1: "a", A2: "20", B2: "b", A3: "30", B3: "c" },
    formula: "=BUSCARV(25; A1:B3; 2; )",
    excel: "#N/A",
  },
  {
    id: "BV-15",
    desc: "BUSCARV aprox: booleanos no coinciden con texto",
    celdas: { A1: "alfa", B1: "t1", A2: "beta", B2: "t2", A3: "=FALSO", B3: "b1", A4: "=VERDADERO", B4: "b2" },
    formula: "=BUSCARV(VERDADERO; A1:B4; 2; VERDADERO)",
    excel: "b2",
  },
  {
    id: "BV-16",
    desc: "BUSCARV exacta (FALSO): texto numérico '001' coincide con '001' y NO con número 1",
    celdas: { A1: "1", B1: "num", A2: "'001", B2: "txt001", A3: "'002", B3: "txt002" },
    formula: '=BUSCARV("001"; A1:B3; 2; FALSO)',
    excel: "txt001",
  },
  {
    id: "BV-17",
    desc: "BUSCARV exacta (FALSO): número 10 NO coincide con texto '10'",
    celdas: { A1: "'10", B1: "texto", A2: "'20", B2: "texto2" },
    formula: "=BUSCARV(10; A1:B2; 2; FALSO)",
    excel: "#N/A",
  },
  {
    id: "BV-18",
    desc: "BUSCARV exacta (FALSO): comodín escapado ~* busca asterisco literal",
    celdas: { A1: "axb", B1: "no", A2: "a*b", B2: "si" },
    formula: '=BUSCARV("a~*b"; A1:B2; 2; FALSO)',
    excel: "si",
  },
  {
    id: "BV-19",
    desc: "BUSCARV aprox: fechas (números de serie) en orden ascendente",
    celdas: { A1: "01/01/2026", B1: "T1", A2: "01/04/2026", B2: "T2", A3: "01/07/2026", B3: "T3" },
    formula: "=BUSCARV(FECHA(2026;5;15); A1:B3; 2; VERDADERO)",
    excel: "T2",
  },

  // --- BUSCARH / HLOOKUP ---
  {
    id: "BH-01",
    desc: "BUSCARH aprox (omitido): mayor valor <= buscado en la primera fila",
    celdas: { A1: "0", B1: "100", C1: "500", D1: "1000", A2: "0%", B2: "5%", C2: "10%", D2: "15%" },
    formula: "=BUSCARH(750; A1:D2; 2)",
    excel: 0.1,
  },
  {
    id: "BH-02",
    desc: "HLOOKUP aprox (TRUE): salta celdas vacías al final en primera fila de texto",
    celdas: { A1: "Ene", B1: "Feb", C1: "Mar", A2: "10", B2: "20", C2: "30", D2: "vacio" },
    formula: '=HLOOKUP("Maz", A1:D2, 2, TRUE)',
    excel: 30,
  },
  {
    id: "BH-03",
    desc: "BUSCARH aprox: buscado menor que el primero → #N/A",
    celdas: { A1: "10", B1: "20", C1: "30", A2: "a", B2: "b", C2: "c" },
    formula: "=BUSCARH(5; A1:C2; 2; VERDADERO)",
    excel: "#N/A",
  },

  // --- COINCIDIR / MATCH ---
  {
    id: "CO-01",
    desc: "COINCIDIR tipo 1 (ascendente) con duplicados devuelve la última posición",
    celdas: { A1: "10", A2: "20", A3: "20", A4: "30" },
    formula: "=COINCIDIR(20; A1:A4; 1)",
    excel: 3,
  },
  {
    id: "CO-02",
    desc: "MATCH tipo -1 (descendente): devuelve el menor valor >= buscado",
    celdas: { A1: "100", A2: "80", A3: "60", A4: "40", A5: "20" },
    formula: "=MATCH(65, A1:A5, -1)",
    excel: 2,
  },
  {
    id: "CO-03",
    desc: "COINCIDIR tipo -1 (descendente): buscado mayor que el primero → #N/A",
    celdas: { A1: "100", A2: "80", A3: "60" },
    formula: "=COINCIDIR(120; A1:A3; -1)",
    excel: "#N/A",
  },
  {
    id: "CO-04",
    desc: "COINCIDIR tipo -1 (descendente) sobre texto salta vacíos iniciales",
    celdas: { A2: "Zorro", A3: "Mesa", A4: "Casa" },
    formula: '=COINCIDIR("Gato"; A1:A4; -1)',
    excel: 3,
  },
  {
    id: "CO-05",
    desc: "COINCIDIR tipo 0 (exacta): texto numérico '10' no coincide con número 10",
    celdas: { A1: "10", A2: "'10", A3: "20" },
    formula: '=COINCIDIR("10"; A1:A3; 0)',
    excel: 2,
  },
  {
    id: "CO-06",
    desc: "COINCIDIR con matriz 2D → #N/A",
    celdas: { A1: "1", B1: "2", A2: "3", B2: "4" },
    formula: "=COINCIDIR(2; A1:B2; 0)",
    excel: "#N/A",
  },

  // --- BUSCARX / XLOOKUP y COINCIDIRX / XMATCH ---
  {
    id: "BX-01",
    desc: "BUSCARX modo 0 (exacta literal): '*' NO es comodín",
    celdas: { A1: "abc", B1: "mal", A2: "a*", B2: "bien" },
    formula: '=BUSCARX("a*"; A1:A2; B1:B2; "no"; 0)',
    excel: "bien",
  },
  {
    id: "BX-02",
    desc: "BUSCARX modo 2 (comodines): '*' y '?' funcionan, '~*' escapa asterisco",
    celdas: { A1: "gato", B1: "1", A2: "g*to", B2: "2", A3: "pato", B3: "3" },
    formula: '=BUSCARX("g~*t?"; A1:A3; B1:B3; "no"; 2)',
    excel: 2,
  },
  {
    id: "BX-03",
    desc: "BUSCARX modo -1 (exacta o menor siguiente) sobre texto desordenado",
    celdas: { A1: "Elena", B1: "E", A2: "Ana", B2: "A", A3: "Beto", B3: "B", A4: "Diego", B4: "D" },
    formula: '=BUSCARX("Carlos"; A1:A4; B1:B4; "no"; -1; 1)',
    excel: "B",
  },
  {
    id: "BX-04",
    desc: "BUSCARX modo 1 (exacta o mayor siguiente) sobre texto desordenado",
    celdas: { A1: "Elena", B1: "E", A2: "Ana", B2: "A", A3: "Beto", B3: "B", A4: "Diego", B4: "D" },
    formula: '=BUSCARX("Carlos"; A1:A4; B1:B4; "no"; 1; 1)',
    excel: "D",
  },
  {
    id: "BX-05",
    desc: "XLOOKUP modo 0, búsqueda -1 (último a primero) devuelve la última coincidencia",
    celdas: { A1: "x", B1: "primero", A2: "y", B2: "medio", A3: "x", B3: "ultimo" },
    formula: '=XLOOKUP("x", A1:A3, B1:B3, "no", 0, -1)',
    excel: "ultimo",
  },
  {
    id: "BX-06",
    desc: "BUSCARX modo -1, búsqueda 2 (binaria ascendente) devuelve mayor <= buscado",
    celdas: { A1: "10", B1: "a", A2: "20", B2: "b", A3: "30", B3: "c", A4: "40", B4: "d" },
    formula: '=BUSCARX(25; A1:A4; B1:B4; "no"; -1; 2)',
    excel: "b",
  },
  {
    id: "BX-07",
    desc: "BUSCARX modo 1, búsqueda 2 (binaria ascendente) devuelve menor >= buscado",
    celdas: { A1: "10", B1: "a", A2: "20", B2: "b", A3: "30", B3: "c", A4: "40", B4: "d" },
    formula: '=BUSCARX(25; A1:A4; B1:B4; "no"; 1; 2)',
    excel: "c",
  },
  {
    id: "BX-08",
    desc: "BUSCARX modo -1, búsqueda -2 (binaria descendente) devuelve mayor <= buscado",
    celdas: { A1: "40", B1: "d", A2: "30", B2: "c", A3: "20", B3: "b", A4: "10", B4: "a" },
    formula: '=BUSCARX(25; A1:A4; B1:B4; "no"; -1; -2)',
    excel: "b",
  },
  {
    id: "BX-09",
    desc: "BUSCARX modo 1, búsqueda -2 (binaria descendente) devuelve menor >= buscado",
    celdas: { A1: "40", B1: "d", A2: "30", B2: "c", A3: "20", B3: "b", A4: "10", B4: "a" },
    formula: '=BUSCARX(25; A1:A4; B1:B4; "no"; 1; -2)',
    excel: "c",
  },
  {
    id: "BX-10",
    desc: "BUSCARX modo 0, búsqueda 2 usa búsqueda binaria real (50 al inicio en vector ascendente no se encuentra)",
    celdas: { A1: "50", B1: "mal", A2: "10", B2: "a", A3: "20", B3: "b", A4: "30", B4: "c", A5: "40", B5: "d" },
    formula: '=BUSCARX(50; A1:A5; B1:B5; "no"; 0; 2)',
    excel: "no",
  },
  {
    id: "BX-11",
    desc: "BUSCARX modo 0, búsqueda -2 usa búsqueda binaria descendente real (5 al inicio no se encuentra)",
    celdas: { A1: "5", B1: "mal", A2: "40", B2: "d", A3: "30", B3: "c", A4: "20", B4: "b", A5: "10", B5: "a" },
    formula: '=BUSCARX(5; A1:A5; B1:B5; "no"; 0; -2)',
    excel: "no",
  },
  {
    id: "BX-12",
    desc: "BUSCARX modo 2 (comodines) con búsqueda binaria 2 → #VALOR! (no lo tapa si_no_está)",
    celdas: { A1: "a", B1: "1", A2: "b", B2: "2" },
    formula: '=BUSCARX("a*"; A1:A2; B1:B2; "no"; 2; 2)',
    excel: "#VALOR!",
  },
  {
    id: "BX-13",
    desc: "COINCIDIRX modo -1 con texto y número mezclados: no cruza tipos",
    celdas: { A1: "10", A2: "20", A3: "Beto", A4: "Diego" },
    formula: '=COINCIDIRX("Ana"; A1:A4; -1; 1)',
    excel: "#N/A",
  },
];

let fallas = 0;
for (const c of CASOS) {
  const real = evalCelda(c.celdas, c.formula);
  const ok = real === c.excel;
  console.log((ok ? "OK    " : "FALLA ") + c.id + " " + c.desc + (ok ? "" : ` — esperado=${JSON.stringify(c.excel)} obtenido=${JSON.stringify(real)}`));
  if (!ok) fallas++;
}
console.log(fallas ? `${fallas}/${CASOS.length} FALLA(S)` : `Todas las pruebas pasaron (${CASOS.length}/${CASOS.length})`);
process.exit(fallas ? 1 : 0);

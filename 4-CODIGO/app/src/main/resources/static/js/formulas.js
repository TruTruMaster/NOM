/* NOM-COLAB · motor de fórmulas del modo hoja.
   Acepta nombres en español y en inglés (=SUMA o =SUM) y los separadores , y ;
   Errores: #REF! #DIV/0! #NOMBRE? #VALOR! #CICLO! #N/A #NUM!
   Fechas como en Excel: un número de serie (días desde el 30/12/1899) que se muestra dd/mm/aaaa.
   Se escriben en una celda como 07/11/2026 o 2026-11-07, o con =FECHA(2026; 11; 7) y =HOY().

   Referencias (8.1): A1, $A$1, $A1, A$1, rangos A1:B5, columnas y filas enteras (A:A, 1:3) y de
   otra hoja (Hoja2!A1, 'Mi hoja'!A1:B3). Al copiar una fórmula se mueven solo las partes sin $
   (desplazar), y al insertar o borrar filas y columnas se corrigen todas (ajustarEstructura).
   Formatos de número como los de Excel ("#,##0.00", "$#,##0.00", "0%", "dd/mm/aaaa"…). */
(function () {
  "use strict";
  const NOM = window.NOM;

  const colNum = (s) => s.toUpperCase().split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const colLetra = (n) => { let s = ""; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const ref = (c, f) => colLetra(c) + (f + 1);
  const parseRef = (r) => { const m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(String(r || "").trim()); return m ? { c: colNum(m[1]) - 1, f: parseInt(m[2], 10) - 1 } : null; };

  class ErrorF { constructor(c) { this.codigo = c; } toString() { return this.codigo; } }
  const E = (c) => new ErrorF(c);

  // ------------------------------------------------------------ fechas
  const DIA = 86400000, BASE = Date.UTC(1899, 11, 30);
  /** Una fecha: su número de serie de Excel (la parte decimal es la hora). Sumarle días da otra
      fecha; restar dos fechas, los días. */
  class FechaF { constructor(serie) { this.serie = +serie; } valueOf() { return this.serie; } }
  const serieDe = (a, m, d) => Math.round((Date.UTC(a, m - 1, d) - BASE) / DIA);
  const fechaDe = (serie) => new Date(BASE + Math.floor(serie + 1e-9) * DIA);
  const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const MESES_C = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  const DIAS_S = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
  const DIAS_C = ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"];
  function textoFecha(f) {
    const d = fechaDe(f.serie);
    const t = `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
    const fr = f.serie - Math.floor(f.serie + 1e-9);
    if (fr > 1e-7) { const min = Math.round(fr * 1440); return t + ` ${String(Math.floor(min / 60) % 24).padStart(2, "0")}:${String(min % 60).padStart(2, "0")}`; }
    return t;
  }
  /** "07/11/2026", "7/11/26" (día/mes/año, como en México) o "2026-11-07". */
  function leerFecha(t) {
    let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
    if (m) return valida(+m[1], +m[2], +m[3]);
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(t);
    if (m) return valida(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
    return null;
  }
  function valida(a, m, d) {
    if (m < 1 || m > 12 || d < 1 || d > 31) return null;
    const f = new Date(Date.UTC(a, m - 1, d));
    return f.getUTCMonth() === m - 1 ? new FechaF(serieDe(a, m, d)) : null;
  }
  function hoy() { const d = new Date(); return new FechaF(serieDe(d.getFullYear(), d.getMonth() + 1, d.getDate())); }
  function ahora() { const d = new Date(); return new FechaF(serieDe(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 60 + d.getMinutes()) / 1440); }

  // ------------------------------------------------------------ criterios de CONTAR.SI, SUMAR.SI…
  /** ">=5", "<>listo", "=a", "a*" (comodines * y ?), 7…: devuelve una función que prueba un valor. */
  function criterio(c) {
    if (c instanceof ErrorF) throw c;
    if (typeof c === "number" || c instanceof FechaF) { const n = +c; return (v) => esNum(v) && +v === n; }
    if (typeof c === "boolean") return (v) => v === c;
    const t = String(c == null ? "" : c);
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(t);
    const op = m[1] || "=", resto = m[2];
    const f = leerFecha(resto.trim()), n = f ? +f : /^-?\d+([.,]\d+)?$/.test(resto.trim()) ? parseFloat(resto.trim().replace(",", ".")) : null;
    if (n != null) {
      const cmp = { "=": (a) => a === n, "<>": (a) => a !== n, "<": (a) => a < n, ">": (a) => a > n, "<=": (a) => a <= n, ">=": (a) => a >= n }[op];
      return (v) => (esNum(v) ? cmp(+v) : op === "<>");
    }
    const bajo = resto.toLowerCase();
    if (op === "=" || op === "<>") {
      const re = /[*?]/.test(bajo) ? new RegExp("^" + bajo.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$") : null;
      const igual = (v) => { const x = texto(v).toLowerCase(); return re ? re.test(x) : x === bajo; };
      return op === "=" ? igual : (v) => !igual(v);
    }
    return (v) => { const x = texto(v).toLowerCase(); return { "<": x < bajo, ">": x > bajo, "<=": x <= bajo, ">=": x >= bajo }[op]; };
  }
  const esNum = (v) => typeof v === "number" || v instanceof FechaF;

  // ------------------------------------------------------------ referencias
  const LETRA_HOJA = "A-Za-zÁÉÍÓÚÑáéíóúñÜü_";
  const RE_REF = new RegExp(
    "^(?:'((?:[^']|'')+)'!|([" + LETRA_HOJA + "][" + LETRA_HOJA + "0-9.]*)!)?" +
    "(?:(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)(?::(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+))?" +
    "|(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})" +
    "|(\\$?)(\\d+):(\\$?)(\\d+))");
  /** Lee una referencia en s desde i. Devuelve { ref, fin } o null (si es un nombre de función, no). */
  function leerRef(s, i) {
    const m = RE_REF.exec(s.slice(i));
    if (!m) return null;
    const fin = i + m[0].length, sig = s[fin] || "";
    if (/[A-Za-z0-9_(ÁÉÍÓÚÑáéíóúñ.]/.test(sig)) return null;
    let j = fin; while (s[j] === " ") j++;
    if (s[j] === "(") return null;
    const hoja = m[1] != null ? m[1].replace(/''/g, "'") : m[2] != null ? m[2] : null;
    let r;
    if (m[4]) {
      const a = { c: colNum(m[4]) - 1, f: +m[6] - 1, ac: !!m[3], af: !!m[5] };
      if (a.f < 0) return null;
      if (m[8]) { const b = { c: colNum(m[8]) - 1, f: +m[10] - 1, ac: !!m[7], af: !!m[9] }; if (b.f < 0) return null; r = { tipo: "rango", a, b }; }
      else r = { tipo: "celda", a, b: a };
    } else if (m[12]) r = { tipo: "col", a: { c: colNum(m[12]) - 1, ac: !!m[11] }, b: { c: colNum(m[14]) - 1, ac: !!m[13] } };
    else {
      const a = { f: +m[16] - 1, af: !!m[15] }, b = { f: +m[18] - 1, af: !!m[17] };
      if (a.f < 0 || b.f < 0) return null;
      r = { tipo: "fila", a, b };
    }
    r.hoja = hoja;
    return { ref: r, fin };
  }
  const necesitaComillas = (n) => !new RegExp("^[" + LETRA_HOJA + "][" + LETRA_HOJA + "0-9.]*$").test(n) || /^[A-Za-z]{1,3}\d+$/.test(n);
  const prefijoHoja = (n) => (n == null ? "" : (necesitaComillas(n) ? "'" + n.replace(/'/g, "''") + "'" : n) + "!");
  function textoRef(r) {
    const p = prefijoHoja(r.hoja);
    const cel = (x) => (x.ac ? "$" : "") + colLetra(x.c) + (x.af ? "$" : "") + (x.f + 1);
    if (r.tipo === "celda") return p + cel(r.a);
    if (r.tipo === "rango") return p + cel(r.a) + ":" + cel(r.b);
    if (r.tipo === "col") return p + (r.a.ac ? "$" : "") + colLetra(r.a.c) + ":" + (r.b.ac ? "$" : "") + colLetra(r.b.c);
    return p + (r.a.af ? "$" : "") + (r.a.f + 1) + ":" + (r.b.af ? "$" : "") + (r.b.f + 1);
  }
  /** Recorre las referencias de una fórmula (sin tocar los textos entre comillas); fn(ref) devuelve
      el texto nuevo, "#REF!" o null para dejarla igual. */
  function transformarRefs(formula, fn) {
    const s = String(formula);
    if (s[0] !== "=") return s;
    let out = "=", i = 1;
    while (i < s.length) {
      const c = s[i];
      if (c === '"') { let j = i + 1; while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; } out += s.slice(i, j + 1); i = j + 1; continue; }
      const prev = out[out.length - 1] || "";
      if (!/[A-Za-z0-9_.$ÁÉÍÓÚÑáéíóúñ]/.test(prev)) {
        const r = leerRef(s, i);
        if (r) { const n = fn(r.ref); out += n == null ? s.slice(i, r.fin) : n; i = r.fin; continue; }
      }
      // Un nombre (función o booleano) se copia entero para no leer "LOG10" como celda.
      const id = new RegExp("^[" + LETRA_HOJA + "][" + LETRA_HOJA + "0-9.]*").exec(s.slice(i));
      if (id && !/[A-Za-z0-9_.$ÁÉÍÓÚÑáéíóúñ]/.test(prev)) {
        const r = leerRef(s, i);
        if (!r) { out += id[0]; i += id[0].length; continue; }
      }
      out += c; i++;
    }
    return out;
  }
  /** Copiar una fórmula dc columnas y df filas más allá: solo se mueve lo que no lleva $. */
  function desplazar(formula, dc, df) {
    if (!dc && !df) return formula;
    return transformarRefs(formula, (r) => {
      const n = JSON.parse(JSON.stringify(r));
      const mueve = (x) => {
        if (x.c != null && !x.ac) x.c += dc;
        if (x.f != null && !x.af) x.f += df;
        return (x.c == null || x.c >= 0) && (x.f == null || x.f >= 0);
      };
      if (!mueve(n.a) || (n.b !== n.a && r.b !== r.a && !mueve(n.b))) return "#REF!";
      if (r.b === r.a) n.b = n.a;
      return textoRef(n);
    });
  }
  const normal = (n) => String(n || "").trim().toLowerCase();
  /**
   * Insertar (n > 0) o borrar (n < 0) filas o columnas de una hoja: corrige las referencias que
   * apuntan a ella. opc = { hoja: nombre de la hoja que cambia, propia: la fórmula vive en esa hoja,
   * eje: "f" (filas) o "c" (columnas), en: índice (desde 0), n }.
   */
  function ajustarEstructura(formula, opc) {
    const k = opc.eje === "c" ? "c" : "f", p = opc.en, n = opc.n;
    return transformarRefs(formula, (r) => {
      const apunta = r.hoja != null ? normal(r.hoja) === normal(opc.hoja) : !!opc.propia;
      if (!apunta) return null;
      if ((k === "f" && r.tipo === "col") || (k === "c" && r.tipo === "fila")) return null;
      const x = JSON.parse(JSON.stringify(r));
      if (x.b === undefined || r.b === r.a) x.b = x.a;
      const unico = r.tipo === "celda";
      if (n > 0) {
        if (x.a[k] >= p) x.a[k] += n;
        if (!unico && x.b[k] >= p) x.b[k] += n;
      } else {
        const q = -n, fin = p + q;
        const dentro = (v) => v >= p && v < fin;
        if (unico) {
          if (dentro(x.a[k])) return "#REF!";
          if (x.a[k] >= fin) x.a[k] -= q;
        } else {
          const ia = x.a[k], ib = x.b[k];
          if (dentro(ia) && dentro(ib)) return "#REF!";
          x.a[k] = ia >= fin ? ia - q : dentro(ia) ? p : ia;
          x.b[k] = ib >= fin ? ib - q : dentro(ib) ? p - 1 : ib;
          if (x.b[k] < x.a[k]) return "#REF!";
        }
      }
      if (unico) x.b = x.a;
      return textoRef(x);
    });
  }
  /** Al cambiar el nombre de una hoja, las fórmulas que la nombran usan el nombre nuevo. */
  function renombrarHoja(formula, viejo, nuevo) {
    return transformarRefs(formula, (r) => (r.hoja != null && normal(r.hoja) === normal(viejo) ? textoRef(Object.assign({}, r, { hoja: nuevo })) : null));
  }
  /** Las referencias de una fórmula (para resaltar celdas mientras se escribe). */
  function refsDe(formula) { const l = []; transformarRefs(formula, (r) => { l.push(r); return null; }); return l; }

  // ------------------------------------------------------------ análisis
  function tokens(s, conLugar) {
    const t = []; let i = 0;
    while (i < s.length) {
      const c = s[i];
      if (/\s/.test(c)) { i++; continue; }
      if (c === '"') {
        let j = i + 1, v = "";
        while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { v += '"'; j += 2; continue; } break; } v += s[j++]; }
        t.push({ k: "str", v }); i = j + 1; continue;
      }
      const r = leerRef(s, i);
      if (r) {
        const x = { k: "ref", v: r.ref };
        if (conLugar) x.p = i;
        i = r.fin;
        // A1#: la matriz que derrama A1
        if (s[i] === "#" && r.ref.tipo === "celda" && !/^#(REF!|DIV|NOMBRE|NAME|VALOR|VALUE|N\/|NUM|CICLO|DESBORD|SPILL|CALC|NUL)/.test(s.slice(i))) { x.derrame = true; i++; }
        t.push(x);
        continue;
      }
      const num = /^(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/.exec(s.slice(i));
      if (num) { t.push({ k: "num", v: parseFloat(num[0]) }); i += num[0].length; continue; }
      const id = /^[A-Za-zÁÉÍÓÚÑáéíóúñ_][A-Za-z0-9ÁÉÍÓÚÑáéíóúñ_.]*/.exec(s.slice(i));
      if (id) { t.push({ k: "id", v: id[0].toUpperCase().replace(/^(_XLFN\.|_XLWS\.|_XLPM\.)+/, "") }); i += id[0].length; continue; }
      const op = /^(<=|>=|<>|[-+*/^&(),;=<>%{}\\@[\]])/.exec(s.slice(i));
      if (op) { t.push({ k: "op", v: op[0] }); i += op[0].length; continue; }
      if (c === "#") { const e = /^#(REF!|DIV\/0!|NOMBRE\?|NAME\?|VALOR!|VALUE!|N\/A|N\/D|NUM!|CICLO!|DESBORDAMIENTO!|SPILL!|CALC!|NULO!|NULL!)/.exec(s.slice(i)); if (e) { t.push({ k: "err", v: e[0] }); i += e[0].length; continue; } }
      throw E("#VALOR!");
    }
    return t;
  }

  // ------------------------------------------------------------ funciones
  const plano = (args) => args.flat(Infinity);
  const nums = (args) => { const p = plano(args); const e = p.find((v) => v instanceof ErrorF); if (e) throw e; return p.filter(esNum).map((v) => +v); };
  /** Un rango como tabla (filas × columnas), para BUSCARV, INDICE… */
  const tabla = (r) => { if (!Array.isArray(r)) return [[r]]; const cols = r.cols || 1; const t = []; for (let i = 0; i < r.length; i += cols) t.push(r.slice(i, i + cols)); return t; };
  const lista = (r) => (Array.isArray(r) ? r : [r]);
  const media = (n) => n.reduce((x, y) => x + y, 0) / n.length;
  function varianza(a, muestra) {
    const n = nums(a);
    if (n.length < (muestra ? 2 : 1)) throw E("#DIV/0!");
    const m = media(n);
    return n.reduce((s, x) => s + (x - m) * (x - m), 0) / (n.length - (muestra ? 1 : 0));
  }
  const redondeo = (n, d, modo) => {
    const p = Math.pow(10, d || 0), x = Math.abs(n) * p;
    const r = modo === "mas" ? Math.ceil(x - 1e-9) : modo === "menos" ? Math.floor(x + 1e-9) : Math.round(x + 1e-9);
    return (Math.sign(n) * r) / p;
  };
  /** Pares rango–criterio de los *.SI.CONJUNTO (desde la posición 'desde'). Devuelve el filtro por índice. */
  function filtroConjunto(a, desde) {
    const pares = [];
    for (let i = desde; i + 1 < a.length; i += 2) pares.push([lista(a[i]), criterio(a[i + 1])]);
    if (!pares.length) throw E("#VALOR!");
    const largo = pares[0][0].length;
    if (pares.some(([r]) => r.length !== largo)) throw E("#VALOR!");
    return { largo, pasa: (k) => pares.every(([r, p]) => p(r[k])) };
  }
  const deFecha = (v) => { const n = num(v); return fechaDe(n); };
  const ultimoDia = (a, m) => new Date(Date.UTC(a, m + 1, 0)).getUTCDate();
  function sumarMeses(v, meses, fin) {
    const d = deFecha(v), a = d.getUTCFullYear(), m = d.getUTCMonth() + Math.trunc(meses);
    const aa = a + Math.floor(m / 12), mm = ((m % 12) + 12) % 12;
    const dia = fin ? ultimoDia(aa, mm) : Math.min(d.getUTCDate(), ultimoDia(aa, mm));
    return new FechaF(serieDe(aa, mm + 1, dia));
  }
  function buscarEn(buscado, valores, tipo) {
    if (buscado instanceof ErrorF) throw buscado;
    if (tipo === 0) {
      const p = criterio(esNum(buscado) ? +buscado : "=" + texto(buscado));
      return valores.findIndex((v) => p(v));
    }
    let mejor = -1;
    for (let i = 0; i < valores.length; i++) {
      const v = valores[i];
      const cmp = esNum(buscado) ? (esNum(v) ? +v - +buscado : null) : (!esNum(v) ? texto(v).toLowerCase().localeCompare(texto(buscado).toLowerCase()) : null);
      if (cmp == null) continue;
      if (tipo === 1 ? cmp <= 0 : cmp >= 0) mejor = i; else break;
    }
    return mejor;
  }
  const FUNCS = {
    // matemáticas
    SUMA: (a) => nums(a).reduce((x, y) => x + y, 0),
    PROMEDIO: (a) => { const n = nums(a); return n.length ? media(n) : E("#DIV/0!"); },
    MIN: (a) => { const n = nums(a); return n.length ? Math.min(...n) : 0; },
    MAX: (a) => { const n = nums(a); return n.length ? Math.max(...n) : 0; },
    CONTAR: (a) => nums(a).length,
    CONTARA: (a) => plano(a).filter((v) => v !== "" && v != null).length,
    "CONTAR.BLANCO": (a) => plano(a).filter((v) => v === "" || v == null).length,
    PRODUCTO: (a) => nums(a).reduce((x, y) => x * y, 1),
    SUMAPRODUCTO: (a) => {
      const ls = a.map(lista), n = ls[0].length;
      if (ls.some((l) => l.length !== n)) throw E("#VALOR!");
      let s = 0;
      for (let i = 0; i < n; i++) s += ls.reduce((p, l) => p * (esNum(l[i]) ? +l[i] : 0), 1);
      return s;
    },
    REDONDEAR: (a) => redondeo(num(a[0]), num(a[1] === undefined ? 0 : a[1])),
    "REDONDEAR.MAS": (a) => redondeo(num(a[0]), num(a[1] === undefined ? 0 : a[1]), "mas"),
    "REDONDEAR.MENOS": (a) => redondeo(num(a[0]), num(a[1] === undefined ? 0 : a[1]), "menos"),
    TRUNCAR: (a) => redondeo(num(a[0]), num(a[1] === undefined ? 0 : a[1]), "menos"),
    ENTERO: (a) => Math.floor(num(a[0])),
    RESIDUO: (a) => { const n = num(a[0]), d = num(a[1]); if (d === 0) throw E("#DIV/0!"); return n - d * Math.floor(n / d); },
    ABS: (a) => Math.abs(num(a[0])),
    RAIZ: (a) => { const n = num(a[0]); return n < 0 ? E("#NUM!") : Math.sqrt(n); },
    POTENCIA: (a) => Math.pow(num(a[0]), num(a[1])),
    PI: () => Math.PI,
    SIGNO: (a) => Math.sign(num(a[0])),
    EXP: (a) => Math.exp(num(a[0])),
    LN: (a) => { const n = num(a[0]); return n <= 0 ? E("#NUM!") : Math.log(n); },
    LOG: (a) => { const n = num(a[0]), b = a[1] === undefined ? 10 : num(a[1]); return n <= 0 || b <= 0 || b === 1 ? E("#NUM!") : Math.log(n) / Math.log(b); },
    LOG10: (a) => { const n = num(a[0]); return n <= 0 ? E("#NUM!") : Math.log10(n); },
    SENO: (a) => Math.sin(num(a[0])),
    COS: (a) => Math.cos(num(a[0])),
    TAN: (a) => Math.tan(num(a[0])),
    GRADOS: (a) => (num(a[0]) * 180) / Math.PI,
    RADIANES: (a) => (num(a[0]) * Math.PI) / 180,
    ALEATORIO: () => Math.random(),
    "ALEATORIO.ENTRE": (a) => { const lo = Math.ceil(num(a[0])), hi = Math.floor(num(a[1])); if (hi < lo) throw E("#NUM!"); return lo + Math.floor(Math.random() * (hi - lo + 1)); },
    // estadística
    MEDIANA: (a) => { const n = nums(a).sort((x, y) => x - y); if (!n.length) throw E("#NUM!"); const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; },
    MODA: (a) => {
      const n = nums(a), cuenta = new Map();
      let mejor = null, max = 1;
      n.forEach((x) => { const c = (cuenta.get(x) || 0) + 1; cuenta.set(x, c); if (c > max) { max = c; mejor = x; } });
      return mejor == null ? E("#N/A") : mejor;
    },
    DESVEST: (a) => Math.sqrt(varianza(a, true)),
    DESVESTP: (a) => Math.sqrt(varianza(a, false)),
    VAR: (a) => varianza(a, true),
    VARP: (a) => varianza(a, false),
    JERARQUIA: (a) => {
      const n = num(a[0]), l = nums([a[1]]), asc = a[2] !== undefined && verdad(a[2]) && a[2] !== 0;
      if (!l.includes(n)) return E("#N/A");
      return 1 + l.filter((x) => (asc ? x < n : x > n)).length;
    },
    "K.ESIMO.MAYOR": (a) => { const l = nums([a[0]]).sort((x, y) => y - x), k = Math.floor(num(a[1])); return k >= 1 && k <= l.length ? l[k - 1] : E("#NUM!"); },
    "K.ESIMO.MENOR": (a) => { const l = nums([a[0]]).sort((x, y) => x - y), k = Math.floor(num(a[1])); return k >= 1 && k <= l.length ? l[k - 1] : E("#NUM!"); },
    "CONTAR.SI": (a) => { const p = criterio(a[1]); return plano([a[0]]).filter((v) => p(v)).length; },
    "SUMAR.SI": (a) => {
      const p = criterio(a[1]), r = plano([a[0]]), sumas = a[2] === undefined ? r : plano([a[2]]);
      let total = 0;
      r.forEach((v, i) => { if (p(v) && esNum(sumas[i])) total += +sumas[i]; });
      return total;
    },
    "PROMEDIO.SI": (a) => {
      const p = criterio(a[1]), r = plano([a[0]]), vals = a[2] === undefined ? r : plano([a[2]]);
      const l = []; r.forEach((v, i) => { if (p(v) && esNum(vals[i])) l.push(+vals[i]); });
      return l.length ? media(l) : E("#DIV/0!");
    },
    "CONTAR.SI.CONJUNTO": (a) => { const f = filtroConjunto(a, 0); let n = 0; for (let k = 0; k < f.largo; k++) if (f.pasa(k)) n++; return n; },
    "SUMAR.SI.CONJUNTO": (a) => { const v = lista(a[0]), f = filtroConjunto(a, 1); let s = 0; for (let k = 0; k < f.largo; k++) if (f.pasa(k) && esNum(v[k])) s += +v[k]; return s; },
    "PROMEDIO.SI.CONJUNTO": (a) => { const v = lista(a[0]), f = filtroConjunto(a, 1), l = []; for (let k = 0; k < f.largo; k++) if (f.pasa(k) && esNum(v[k])) l.push(+v[k]); return l.length ? media(l) : E("#DIV/0!"); },
    "MAX.SI.CONJUNTO": (a) => { const v = lista(a[0]), f = filtroConjunto(a, 1), l = []; for (let k = 0; k < f.largo; k++) if (f.pasa(k) && esNum(v[k])) l.push(+v[k]); return l.length ? Math.max(...l) : 0; },
    "MIN.SI.CONJUNTO": (a) => { const v = lista(a[0]), f = filtroConjunto(a, 1), l = []; for (let k = 0; k < f.largo; k++) if (f.pasa(k) && esNum(v[k])) l.push(+v[k]); return l.length ? Math.min(...l) : 0; },
    // lógicas
    SI: (a) => { if (a[0] instanceof ErrorF) throw a[0]; return verdad(a[0]) ? (a[1] === undefined ? true : a[1]) : (a[2] === undefined ? false : a[2]); },
    "SI.CONJUNTO": (a) => { for (let i = 0; i + 1 < a.length; i += 2) { if (a[i] instanceof ErrorF) throw a[i]; if (verdad(a[i])) return a[i + 1]; } return E("#N/A"); },
    ESERROR: (a) => a[0] instanceof ErrorF,
    SIERROR: (a) => (a[0] instanceof ErrorF ? (a[1] === undefined ? "" : a[1]) : a[0]),
    ESBLANCO: (a) => a[0] === "" || a[0] == null,
    ESNUMERO: (a) => esNum(a[0]),
    ESTEXTO: (a) => typeof a[0] === "string" && a[0] !== "",
    Y: (a) => plano(a).every(verdad),
    O: (a) => plano(a).some(verdad),
    NO: (a) => !verdad(a[0]),
    // texto
    CONCATENAR: (a) => plano(a).map(texto).join(""),
    UNIRCADENAS: (a) => { const d = texto(a[0]), ign = verdad(a[1]); return plano(a.slice(2)).map(texto).filter((x) => !ign || x !== "").join(d); },
    MAYUSC: (a) => texto(a[0]).toUpperCase(),
    MINUSC: (a) => texto(a[0]).toLowerCase(),
    NOMPROPIO: (a) => texto(a[0]).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (m, p, l) => p + l.toUpperCase()),
    LARGO: (a) => Array.from(texto(a[0])).length,
    IZQUIERDA: (a) => Array.from(texto(a[0])).slice(0, a[1] === undefined ? 1 : Math.max(0, num(a[1]))).join(""),
    DERECHA: (a) => { const c = Array.from(texto(a[0])), n = a[1] === undefined ? 1 : Math.max(0, num(a[1])); return n ? c.slice(-n).join("") : ""; },
    EXTRAE: (a) => { const i = Math.floor(num(a[1])); if (i < 1) throw E("#VALOR!"); return Array.from(texto(a[0])).slice(i - 1, i - 1 + Math.max(0, num(a[2]))).join(""); },
    ESPACIOS: (a) => texto(a[0]).trim().replace(/ {2,}/g, " "),
    SUSTITUIR: (a) => {
      const t = texto(a[0]), viejo = texto(a[1]), nuevo = texto(a[2]);
      if (!viejo) return t;
      if (a[3] === undefined) return t.split(viejo).join(nuevo);
      let k = Math.floor(num(a[3])), i = -1;
      while (k-- > 0) { i = t.indexOf(viejo, i + 1); if (i < 0) return t; }
      return t.slice(0, i) + nuevo + t.slice(i + viejo.length);
    },
    REEMPLAZAR: (a) => { const c = Array.from(texto(a[0])), i = Math.floor(num(a[1])) - 1, n = Math.max(0, num(a[2])); return c.slice(0, i).join("") + texto(a[3]) + c.slice(i + n).join(""); },
    ENCONTRAR: (a) => { const i = texto(a[1]).indexOf(texto(a[0]), (a[2] === undefined ? 1 : num(a[2])) - 1); return i < 0 ? E("#VALOR!") : i + 1; },
    HALLAR: (a) => { const i = texto(a[1]).toLowerCase().indexOf(texto(a[0]).toLowerCase(), (a[2] === undefined ? 1 : num(a[2])) - 1); return i < 0 ? E("#VALOR!") : i + 1; },
    TEXTO: (a) => aplicarFormato(a[0], texto(a[1])),
    VALOR: (a) => { const v = calcularLiteral(texto(a[0]).trim()); if (!esNum(v)) throw E("#VALOR!"); return +v; },
    REPETIR: (a) => texto(a[0]).repeat(Math.max(0, Math.min(10000, Math.floor(num(a[1]))))),
    IGUAL: (a) => texto(a[0]) === texto(a[1]),
    CARACTER: (a) => String.fromCharCode(Math.floor(num(a[0]))),
    // fechas
    HOY: () => hoy(),
    AHORA: () => ahora(),
    FECHA: (a) => { const y = Math.floor(num(a[0])), m = Math.floor(num(a[1])), d = Math.floor(num(a[2])); return new FechaF(Math.round((Date.UTC(y < 100 ? 1900 + y : y, m - 1, d) - BASE) / DIA)); },
    DIAS: (a) => Math.round(num(a[0]) - num(a[1])),
    "AÑO": (a) => deFecha(a[0]).getUTCFullYear(),
    MES: (a) => deFecha(a[0]).getUTCMonth() + 1,
    DIA: (a) => deFecha(a[0]).getUTCDate(),
    DIASEM: (a) => { const d = deFecha(a[0]).getUTCDay(), t = a[1] === undefined ? 1 : num(a[1]); return t === 2 ? ((d + 6) % 7) + 1 : t === 3 ? (d + 6) % 7 : d + 1; },
    "FIN.MES": (a) => sumarMeses(a[0], num(a[1] === undefined ? 0 : a[1]), true),
    "FECHA.MES": (a) => sumarMeses(a[0], num(a[1] === undefined ? 0 : a[1]), false),
    "DIAS.LAB": (a) => {
      let i = Math.floor(num(a[0])), f = Math.floor(num(a[1])), sg = 1;
      if (f < i) { [i, f] = [f, i]; sg = -1; }
      let n = 0;
      for (let s = i; s <= f && n < 100000; s++) { const d = fechaDe(s).getUTCDay(); if (d !== 0 && d !== 6) n++; }
      return n * sg;
    },
    SIFECHA: (a) => {
      const d1 = deFecha(a[0]), d2 = deFecha(a[1]), u = texto(a[2]).toUpperCase();
      if (d2 < d1) throw E("#NUM!");
      let meses = (d2.getUTCFullYear() - d1.getUTCFullYear()) * 12 + d2.getUTCMonth() - d1.getUTCMonth();
      if (d2.getUTCDate() < d1.getUTCDate()) meses--;
      if (u === "Y" || u === "A") return Math.floor(meses / 12);
      if (u === "M") return meses;
      if (u === "D") return Math.round((d2 - d1) / DIA);
      throw E("#NUM!");
    },
    // búsqueda
    BUSCARV: (a) => {
      const t = tabla(a[1]);
      const col = Math.floor(num(a[2]));
      if (col < 1 || col > (t[0] || []).length) return E("#REF!");
      const exacta = a[3] !== undefined && !(verdad(a[3]) && a[3] !== 0);
      const i = buscarEn(a[0], t.map((f) => f[0]), exacta ? 0 : 1);
      return i < 0 ? E("#N/A") : t[i][col - 1];
    },
    BUSCARH: (a) => {
      const t = tabla(a[1]), fila = Math.floor(num(a[2]));
      if (fila < 1 || fila > t.length) return E("#REF!");
      const exacta = a[3] !== undefined && !(verdad(a[3]) && a[3] !== 0);
      const i = buscarEn(a[0], t[0], exacta ? 0 : 1);
      return i < 0 ? E("#N/A") : t[fila - 1][i];
    },
    BUSCARX: (a) => {
      const donde = lista(a[1]), que = lista(a[2]);
      const i = buscarEn(a[0], donde, 0);
      if (i < 0) return a[3] === undefined ? E("#N/A") : a[3];
      return que[i] === undefined ? E("#VALOR!") : que[i];
    },
    INDICE: (a) => {
      const t = tabla(a[0]);
      let f = a[1] === undefined ? 1 : Math.floor(num(a[1])), c = a[2] === undefined ? 1 : Math.floor(num(a[2]));
      if (t.length === 1 && a[2] === undefined) { c = f; f = 1; }
      if (f < 1 || c < 1 || f > t.length || c > (t[0] || []).length) return E("#REF!");
      return t[f - 1][c - 1];
    },
    COINCIDIR: (a) => {
      const tipo = a[2] === undefined ? 1 : Math.sign(num(a[2]));
      const i = buscarEn(a[0], lista(a[1]), tipo === 0 ? 0 : tipo);
      return i < 0 ? E("#N/A") : i + 1;
    },
    ELEGIR: (a) => { const i = Math.floor(num(a[0])); if (i < 1 || i >= a.length) return E("#VALOR!"); return a[i]; },
  };
  FUNCS.CONCAT = FUNCS.CONCATENAR;

  /** DEC.A.BIN/OCT/HEX: número decimal en la base, con negativos en complemento a 10 dígitos (como Excel). */
  function baseDe(a, b) {
    let n = Math.floor(num(a[0])); const lugares = a[1] === undefined ? 0 : Math.floor(num(a[1]));
    const lim = { 2: 512, 8: 536870912, 16: 549755813888 }[b];
    if (n < -lim || n >= lim) throw E("#NUM!");
    if (n < 0) n = Math.pow(b, 10) + n;
    const t = n.toString(b).toUpperCase();
    if (lugares && t.length > lugares) throw E("#NUM!");
    return lugares ? t.padStart(lugares, "0") : t;
  }
  /** BIN/OCT/HEX.A.DEC: de la base a decimal (10 dígitos con signo en complemento). */
  function aDec(a, b) {
    const t = texto(a[0]).trim(); if (!t) return 0;
    if (t.length > 10 || !new RegExp("^[" + "0123456789ABCDEF".slice(0, b) + "]+$", "i").test(t)) throw E("#NUM!");
    let n = parseInt(t, b);
    if (t.length === 10 && parseInt(t[0], b) >= b / 2) n -= Math.pow(b, 10);
    return n;
  }
  // ------------------------------------------------------------ más funciones (0.13.2)
  const oDef = (v, def) => (v === undefined || v === "" ? def : v);
  const fraccionDia = (v) => { const n = num(v); return n - Math.floor(n + 1e-9); };
  const segundosDia = (v) => Math.round(fraccionDia(v) * 86400) % 86400;
  const entero = (v) => Math.floor(num(v) + 1e-9);
  /** Los pares (x, y) con número en los dos lados, para regresión y correlación. */
  function pares(ry, rx) {
    const ys = lista(ry), xs = rx === undefined ? ys.map((_, i) => i + 1) : lista(rx);
    if (ys.length !== xs.length) throw E("#N/A");
    const p = [];
    for (let i = 0; i < ys.length; i++) { if (ys[i] instanceof ErrorF) throw ys[i]; if (xs[i] instanceof ErrorF) throw xs[i]; if (esNum(xs[i]) && esNum(ys[i])) p.push([+xs[i], +ys[i]]); }
    if (!p.length) throw E("#DIV/0!");
    const mx = media(p.map((q) => q[0])), my = media(p.map((q) => q[1]));
    let sxx = 0, syy = 0, sxy = 0;
    p.forEach(([x, y]) => { sxx += (x - mx) * (x - mx); syy += (y - my) * (y - my); sxy += (x - mx) * (y - my); });
    return { n: p.length, mx, my, sxx, syy, sxy };
  }
  const recta = (ry, rx) => { const q = pares(ry, rx); if (q.sxx === 0) throw E("#DIV/0!"); const m = q.sxy / q.sxx; return { m, b: q.my - m * q.mx, q }; };
  function percentil(r, k) {
    const n = nums([r]).sort((x, y) => x - y);
    if (!n.length || !(k >= 0 && k <= 1)) throw E("#NUM!");
    const pos = k * (n.length - 1), i = Math.floor(pos + 1e-12), fr = pos - i;
    return fr < 1e-12 || i + 1 >= n.length ? n[i] : n[i] + fr * (n[i + 1] - n[i]);
  }
  /** Los valores como los cuentan PROMEDIOA, MAXA y MINA: el texto vale 0, VERDADERO 1, lo vacío no cuenta. */
  const valoresA = (a) => plano(a).filter((v) => v !== "" && v != null).map((v) => { if (v instanceof ErrorF) throw v; return esNum(v) ? +v : typeof v === "boolean" ? +v : 0; });
  const festivos = (r) => new Set(r === undefined ? [] : plano([r]).filter(esNum).map((v) => Math.floor(+v)));
  const laboral = (serie, fest) => { const d = fechaDe(serie).getUTCDay(); return d !== 0 && d !== 6 && !fest.has(serie); };
  function mcd(x, y) { while (y) { [x, y] = [y, x % y]; } return x; }
  const enterosPos = (a) => { const n = nums(a).map((v) => Math.floor(v)); if (n.some((v) => v < 0)) throw E("#NUM!"); return n; };
  function factorial(n) { n = Math.floor(n); if (n < 0 || n > 170) throw E("#NUM!"); let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; }
  const conMiles = (n, dec, miles) => { const d = Math.max(0, Math.min(20, Math.floor(dec))), x = redondeo(Math.abs(n), Math.floor(dec)); return x.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: miles }); };
  /** Texto antes o después de un separador (TEXTOANTES y TEXTODESPUES). */
  function partirEn(a, antes) {
    const t = texto(a[0]), sep = texto(a[1]), k = a[2] === undefined ? 1 : Math.trunc(num(a[2])), sinMayus = a[3] !== undefined && num(a[3]) === 1;
    if (k === 0) throw E("#VALOR!");
    if (!sep) return antes ? (k > 0 ? "" : t) : (k > 0 ? t : "");
    const T = sinMayus ? t.toLowerCase() : t, S = sinMayus ? sep.toLowerCase() : sep, pos = [];
    for (let i = T.indexOf(S); i >= 0; i = T.indexOf(S, i + S.length)) pos.push(i);
    const i = k > 0 ? pos[k - 1] : pos[pos.length + k];
    if (i === undefined) { if (a[5] !== undefined) return a[5]; throw E("#N/A"); }
    return antes ? t.slice(0, i) : t.slice(i + sep.length);
  }
  // ---- financieras (los pagos que salen de tu bolsillo son negativos, como en Excel)
  const fin = {
    vf: (r, n, pago, va, tipo) => (r === 0 ? -(va + pago * n) : -(va * Math.pow(1 + r, n) + pago * (1 + r * tipo) * (Math.pow(1 + r, n) - 1) / r)),
    va: (r, n, pago, vf, tipo) => (r === 0 ? -(vf + pago * n) : -(vf + pago * (1 + r * tipo) * (Math.pow(1 + r, n) - 1) / r) / Math.pow(1 + r, n)),
    pago: (r, n, va, vf, tipo) => { if (n === 0) throw E("#NUM!"); return r === 0 ? -(va + vf) / n : -(r * (va * Math.pow(1 + r, n) + vf)) / ((1 + r * tipo) * (Math.pow(1 + r, n) - 1)); },
    interes: (r, per, n, va, vf, tipo) => {
      if (per < 1 || per > n) throw E("#NUM!");
      if (tipo === 1 && per === 1) return 0;
      const i = fin.vf(r, per - 1, fin.pago(r, n, va, vf, tipo), va, tipo) * r;
      return tipo === 1 ? i / (1 + r) : i;
    },
  };
  const args5 = (a, i) => [a[i] === undefined ? 0 : num(a[i]), a[i + 1] !== undefined && num(a[i + 1]) !== 0 ? 1 : 0];
  /** Raíz de f cerca de x0 (Newton con derivada numérica; si no llega, bisección entre -0.99 y 10). */
  function raiz(f, x0) {
    let x = x0;
    for (let k = 0; k < 60; k++) {
      const y = f(x), h = Math.max(1e-7, Math.abs(x) * 1e-7), d = (f(x + h) - y) / h;
      if (!isFinite(y) || !isFinite(d) || d === 0) break;
      const x1 = x - y / d;
      if (!isFinite(x1) || x1 <= -1) break;
      if (Math.abs(x1 - x) < 1e-11) return x1;
      x = x1;
    }
    let lo = -0.99, hi = 10, flo = f(lo), fhi = f(hi);
    if (!isFinite(flo) || !isFinite(fhi) || flo * fhi > 0) throw E("#NUM!");
    for (let k = 0; k < 200; k++) { const m = (lo + hi) / 2, fm = f(m); if (Math.abs(fm) < 1e-12 || hi - lo < 1e-12) return m; if (flo * fm <= 0) { hi = m; fhi = fm; } else { lo = m; flo = fm; } }
    return (lo + hi) / 2;
  }
  const ROMANOS = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  Object.assign(FUNCS, {
    // fecha y hora
    HORA: (a) => Math.floor(segundosDia(a[0]) / 3600),
    MINUTO: (a) => Math.floor((segundosDia(a[0]) % 3600) / 60),
    SEGUNDO: (a) => segundosDia(a[0]) % 60,
    NSHORA: (a) => { const s = num(a[0]) * 3600 + num(a[1]) * 60 + num(oDef(a[2], 0)); if (s < 0) throw E("#NUM!"); return (s % 86400) / 86400; },
    HORANUMERO: (a) => {
      const m = /^(?:.*\s)?(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap])?\.?\s*m?\.?$/i.exec(texto(a[0]).trim());
      if (!m) throw E("#VALOR!");
      let h = +m[1];
      if (m[4]) h = (h % 12) + (/p/i.test(m[4]) ? 12 : 0);
      if (h > 23 || +m[2] > 59 || +(m[3] || 0) > 59) throw E("#VALOR!");
      return (h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400;
    },
    FECHANUMERO: (a) => { const f = a[0] instanceof FechaF ? a[0] : leerFecha(texto(a[0]).trim()); if (!f) throw E("#VALOR!"); return Math.floor(+f); },
    "NUM.DE.SEMANA": (a) => {
      const s = entero(a[0]), tipo = a[1] === undefined ? 1 : Math.floor(num(a[1]));
      if (tipo === 21) return FUNCS["ISO.NUM.DE.SEMANA"]([a[0]]);
      const inicio = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 }[tipo];
      if (inicio === undefined) throw E("#NUM!");
      const d = fechaDe(s), ene1 = serieDe(d.getUTCFullYear(), 1, 1);
      return Math.floor((s - ene1 + ((fechaDe(ene1).getUTCDay() - inicio + 7) % 7)) / 7) + 1;
    },
    "ISO.NUM.DE.SEMANA": (a) => {
      const s = entero(a[0]), jueves = s + 3 - ((fechaDe(s).getUTCDay() + 6) % 7);
      return Math.floor((jueves - serieDe(fechaDe(jueves).getUTCFullYear(), 1, 1)) / 7) + 1;
    },
    DIAS360: (a) => {
      const d1 = deFecha(a[0]), d2 = deFecha(a[1]), europeo = a[2] !== undefined && verdad(a[2]) && a[2] !== 0;
      let x1 = d1.getUTCDate(), x2 = d2.getUTCDate();
      const m1 = d1.getUTCMonth(), m2 = d2.getUTCMonth(), y1 = d1.getUTCFullYear(), y2 = d2.getUTCFullYear();
      if (europeo) { if (x1 === 31) x1 = 30; if (x2 === 31) x2 = 30; }
      else {
        if (x1 === ultimoDia(y1, m1)) x1 = 30;
        if (x2 === 31 && x1 >= 30) x2 = 30;
      }
      return (y2 - y1) * 360 + (m2 - m1) * 30 + (x2 - x1);
    },
    "DIA.LAB": (a) => {
      let s = entero(a[0]), faltan = Math.trunc(num(a[1]));
      const fest = festivos(a[2]), paso = faltan < 0 ? -1 : 1;
      for (let k = 0; faltan !== 0 && k < 200000; k++) { s += paso; if (laboral(s, fest)) faltan -= paso; }
      return new FechaF(s);
    },
    "DIAS.LAB": (a) => {
      let i = entero(a[0]), f = entero(a[1]), sg = 1;
      if (f < i) { [i, f] = [f, i]; sg = -1; }
      const fest = festivos(a[2]);
      let n = 0;
      for (let s = i; s <= f && s - i < 200000; s++) if (laboral(s, fest)) n++;
      return n * sg;
    },
    // información y lógicas
    ESNOD: (a) => a[0] instanceof ErrorF && a[0].codigo === "#N/A",
    ESERR: (a) => a[0] instanceof ErrorF && a[0].codigo !== "#N/A",
    NOD: () => E("#N/A"),
    "SI.ND": (a) => { if (a[0] instanceof ErrorF) { if (a[0].codigo === "#N/A") return a[1] === undefined ? "" : a[1]; throw a[0]; } return a[0]; },
    "ES.PAR": (a) => Math.floor(Math.abs(num(a[0]))) % 2 === 0,
    "ES.IMPAR": (a) => Math.floor(Math.abs(num(a[0]))) % 2 === 1,
    ESLOGICO: (a) => typeof a[0] === "boolean",
    ESNOTEXTO: (a) => !(typeof a[0] === "string" && a[0] !== ""),
    CAMBIAR: (a) => {
      if (a[0] instanceof ErrorF) throw a[0];
      const igual = (x, y) => (esNum(x) && esNum(y) ? +x === +y : typeof x === typeof y && texto(x).toLowerCase() === texto(y).toLowerCase());
      for (let i = 1; i + 1 < a.length; i += 2) if (igual(a[0], a[i])) return a[i + 1];
      return a.length % 2 === 0 ? a[a.length - 1] : E("#N/A");
    },
    XO: (a) => plano(a).filter(verdad).length % 2 === 1,
    N: (a) => { if (a[0] instanceof ErrorF) throw a[0]; return esNum(a[0]) ? +a[0] : a[0] === true ? 1 : 0; },
    T: (a) => { if (a[0] instanceof ErrorF) throw a[0]; return typeof a[0] === "string" ? a[0] : ""; },
    TIPO: (a) => (a[0] instanceof ErrorF ? 16 : Array.isArray(a[0]) ? 64 : typeof a[0] === "boolean" ? 4 : esNum(a[0]) || a[0] === "" || a[0] == null ? 1 : 2),
    // texto
    TEXTOANTES: (a) => partirEn(a, true),
    TEXTODESPUES: (a) => partirEn(a, false),
    LIMPIAR: (a) => Array.from(texto(a[0])).filter((c) => c.charCodeAt(0) >= 32).join(""),
    CODIGO: (a) => { const t = texto(a[0]); if (!t) throw E("#VALOR!"); return t.codePointAt(0); },
    UNICODE: (a) => { const t = texto(a[0]); if (!t) throw E("#VALOR!"); return t.codePointAt(0); },
    UNICAR: (a) => { const n = Math.floor(num(a[0])); if (n < 1 || n > 0x10FFFF) throw E("#VALOR!"); return String.fromCodePoint(n); },
    MONEDA: (a) => { const n = num(a[0]), d = a[1] === undefined ? 2 : num(a[1]); return (redondeo(n, Math.floor(d)) < 0 ? "-" : "") + "$" + conMiles(n, d, true); },
    DECIMAL: (a) => { const n = num(a[0]), d = a[1] === undefined ? 2 : num(a[1]); return (redondeo(n, Math.floor(d)) < 0 ? "-" : "") + conMiles(n, d, !(a[2] !== undefined && verdad(a[2]) && a[2] !== 0)); },
    // 16-nov-2026 (paridad con Excel)
    SUBTOTALES: (a) => {
      const f = Math.floor(num(a[0])) % 100, resto = a.slice(1);
      const m = { 1: "PROMEDIO", 2: "CONTAR", 3: "CONTARA", 4: "MAX", 5: "MIN", 6: "PRODUCTO", 7: "DESVEST", 8: "DESVESTP", 9: "SUMA", 10: "VAR", 11: "VARP" }[f];
      if (!m) throw E("#VALOR!");
      return FUNCS[m](resto);
    },
    AGREGAR: (a) => {
      const f = Math.floor(num(a[0])), op = Math.floor(num(oDef(a[1], 0)));
      const m = { 1: "PROMEDIO", 2: "CONTAR", 3: "CONTARA", 4: "MAX", 5: "MIN", 6: "PRODUCTO", 7: "DESVEST", 8: "DESVESTP", 9: "SUMA", 10: "VAR", 11: "VARP", 12: "MEDIANA", 13: "MODA" }[f];
      if (!m && f !== 14 && f !== 15 && f !== 16 && f !== 19) throw E("#VALOR!");
      const sinErr = (v) => (Array.isArray(v) ? v.map(sinErr) : v instanceof ErrorF ? "" : v);
      const datos = [6, 7, 2, 3].includes(op) ? sinErr(a.slice(2, f >= 14 ? 3 : undefined)) : a.slice(2, f >= 14 ? 3 : undefined);
      if (f === 14) return FUNCS["K.ESIMO.MAYOR"]([datos[0], a[3]]);
      if (f === 15) return FUNCS["K.ESIMO.MENOR"]([datos[0], a[3]]);
      if (f === 16) return FUNCS.PERCENTIL([datos[0], a[3]]);
      if (f === 19) return FUNCS.PERCENTIL([datos[0], num(a[3]) / 4]);
      return FUNCS[m](datos);
    },
    COINCIDIRX: (a) => {
      const vals = lista(a[1]).flat(Infinity), modo = Math.floor(num(oDef(a[2], 0))), busq = Math.floor(num(oDef(a[3], 1)));
      const orden = busq < 0 ? vals.map((_, i) => vals.length - 1 - i) : vals.map((_, i) => i);
      const b = a[0];
      if (modo === 0 || modo === 2) {
        const p = criterio(esNum(b) ? +b : "=" + texto(b));
        for (const i of orden) if (p(vals[i])) return i + 1;
        return E("#N/A");
      }
      let mejor = -1;
      for (const i of orden) {
        const v = vals[i]; if (!esNum(v) || !esNum(b)) { if (texto(v).toLowerCase() === texto(b).toLowerCase()) return i + 1; continue; }
        const d = +v - +b;
        if (d === 0) return i + 1;
        if (modo === -1 && d < 0 && (mejor < 0 || +v > +vals[mejor])) mejor = i;
        if (modo === 1 && d > 0 && (mejor < 0 || +v < +vals[mejor])) mejor = i;
      }
      return mejor < 0 ? E("#N/A") : mejor + 1;
    },
    CONVERTIR: (a) => {
      const n = num(a[0]), de = texto(a[1]), a2 = texto(a[2]);
      const T = { C: 1, F: 1, K: 1, cel: 1, fah: 1, kel: 1 };
      if (T[de] && T[a2]) {
        const k = de === "C" || de === "cel" ? n + 273.15 : de === "F" || de === "fah" ? (n - 32) * 5 / 9 + 273.15 : n;
        return a2 === "C" || a2 === "cel" ? k - 273.15 : a2 === "F" || a2 === "fah" ? (k - 273.15) * 9 / 5 + 32 : k;
      }
      const U = {
        m: ["l", 1], km: ["l", 1000], cm: ["l", 0.01], mm: ["l", 0.001], mi: ["l", 1609.344], ft: ["l", 0.3048], in: ["l", 0.0254], yd: ["l", 0.9144], Nmi: ["l", 1852], ang: ["l", 1e-10],
        g: ["m", 1], kg: ["m", 1000], mg: ["m", 0.001], lbm: ["m", 453.59237], ozm: ["m", 28.349523125], ton: ["m", 907184.74],
        sec: ["t", 1], s: ["t", 1], mn: ["t", 60], min: ["t", 60], hr: ["t", 3600], day: ["t", 86400], d: ["t", 86400], yr: ["t", 31557600],
        l: ["v", 1], L: ["v", 1], lt: ["v", 1], ml: ["v", 0.001], gal: ["v", 3.785411784], qt: ["v", 0.946352946], pt: ["v", 0.473176473], cup: ["v", 0.2365882365], oz: ["v", 0.0295735295625], tsp: ["v", 0.00492892159375], tbs: ["v", 0.01478676478125], m3: ["v", 1000],
        J: ["e", 1], kJ: ["e", 1000], cal: ["e", 4.1868], kcal: ["e", 4186.8], Wh: ["e", 3600], kWh: ["e", 3600000],
        W: ["p", 1], kW: ["p", 1000], HP: ["p", 745.69987158227], h: ["p", 745.69987158227], Pa: ["pr", 1], atm: ["pr", 101325], mmHg: ["pr", 133.322], psi: ["pr", 6894.75729316836], bar: ["pr", 100000],
        "km/h": ["vel", 1000 / 3600], "m/s": ["vel", 1], mph: ["vel", 0.44704], kn: ["vel", 1852 / 3600], ha: ["a", 10000], m2: ["a", 1], km2: ["a", 1e6], ft2: ["a", 0.09290304], ac: ["a", 4046.8564224],
      };
      const x = U[de], y = U[a2];
      if (!x || !y || x[0] !== y[0]) return E("#N/A");
      return (n * x[1]) / y[1];
    },
    "DEC.A.BIN": (a) => baseDe(a, 2, 10), "DEC.A.OCT": (a) => baseDe(a, 8, 10), "DEC.A.HEX": (a) => baseDe(a, 16, 10),
    "BIN.A.DEC": (a) => aDec(a, 2), "OCT.A.DEC": (a) => aDec(a, 8), "HEX.A.DEC": (a) => aDec(a, 16),
    BASE: (a) => { const n = Math.floor(num(a[0])), b = Math.floor(num(a[1])), l = Math.floor(num(oDef(a[2], 0))); if (n < 0 || b < 2 || b > 36) throw E("#NUM!"); return n.toString(b).toUpperCase().padStart(l, "0"); },
    "NUMERO.ARABE": (a) => {
      const t = texto(a[0]).trim().toUpperCase(); if (!t) return 0;
      const V = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 }; let s = 0, neg = t[0] === "-";
      const r = neg ? t.slice(1) : t;
      for (let i = 0; i < r.length; i++) { const v = V[r[i]], w = V[r[i + 1]] || 0; if (!v) throw E("#VALOR!"); s += v < w ? -v : v; }
      return neg ? -s : s;
    },
    "NUMERO.ROMANO": (a) => { let n = Math.floor(num(a[0])), r = ""; if (n < 0 || n > 3999) throw E("#VALOR!"); for (const [v, t] of ROMANOS) while (n >= v) { r += t; n -= v; } return r; },
    // matemáticas
    "MULTIPLO.SUPERIOR": (a) => { const n = num(a[0]), m = a[1] === undefined ? 1 : num(a[1]); if (m === 0) return 0; if (n > 0 && m < 0) throw E("#NUM!"); return Math.ceil(n / m - 1e-10) * m; },
    "MULTIPLO.INFERIOR": (a) => { const n = num(a[0]), m = a[1] === undefined ? 1 : num(a[1]); if (m === 0) { if (n === 0) return 0; throw E("#DIV/0!"); } if (n > 0 && m < 0) throw E("#NUM!"); return Math.floor(n / m + 1e-10) * m; },
    "MULTIPLO.SUPERIOR.MAT": (a) => { const n = num(a[0]), m = Math.abs(a[1] === undefined ? 1 : num(a[1])), lejos = a[2] !== undefined && num(a[2]) !== 0; if (m === 0) return 0; return n < 0 && lejos ? -Math.ceil(-n / m - 1e-10) * m : Math.ceil(n / m - 1e-10) * m; },
    "MULTIPLO.INFERIOR.MAT": (a) => { const n = num(a[0]), m = Math.abs(a[1] === undefined ? 1 : num(a[1])), cerca = a[2] !== undefined && num(a[2]) !== 0; if (m === 0) return 0; return n < 0 && cerca ? -Math.floor(-n / m + 1e-10) * m : Math.floor(n / m + 1e-10) * m; },
    "REDOND.MULT": (a) => { const n = num(a[0]), m = num(a[1]); if (m === 0) return 0; if (n * m < 0) throw E("#NUM!"); return Math.round(n / m + 1e-10) * m; },
    "REDONDEA.PAR": (a) => { const n = num(a[0]); return Math.sign(n) * Math.ceil(Math.abs(n) / 2 - 1e-10) * 2; },
    "REDONDEA.IMPAR": (a) => { const n = num(a[0]), x = Math.abs(n); return (n < 0 ? -1 : 1) * (Math.ceil((x - 1) / 2 - 1e-10) * 2 + 1); },
    "M.C.D": (a) => enterosPos(a).reduce((x, y) => mcd(x, y), 0),
    "M.C.M": (a) => enterosPos(a).reduce((x, y) => (x === 0 || y === 0 ? 0 : (x / mcd(x, y)) * y), 1),
    FACT: (a) => factorial(num(a[0])),
    COMBINAT: (a) => { const n = Math.floor(num(a[0])), k = Math.floor(num(a[1])); if (n < 0 || k < 0 || k > n) throw E("#NUM!"); let r = 1; for (let i = 1; i <= Math.min(k, n - k); i++) r = (r * (n - Math.min(k, n - k) + i)) / i; return Math.round(r); },
    PERMUTACIONES: (a) => { const n = Math.floor(num(a[0])), k = Math.floor(num(a[1])); if (n < 0 || k < 0 || k > n) throw E("#NUM!"); let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r; },
    COCIENTE: (a) => { const d = num(a[1]); if (d === 0) throw E("#DIV/0!"); return Math.trunc(num(a[0]) / d); },
    "SUMA.CUADRADOS": (a) => nums(a).reduce((x, y) => x + y * y, 0),
    ASENO: (a) => { const n = num(a[0]); if (n < -1 || n > 1) throw E("#NUM!"); return Math.asin(n); },
    ACOS: (a) => { const n = num(a[0]); if (n < -1 || n > 1) throw E("#NUM!"); return Math.acos(n); },
    ATAN: (a) => Math.atan(num(a[0])),
    ATAN2: (a) => { const x = num(a[0]), y = num(a[1]); if (x === 0 && y === 0) throw E("#DIV/0!"); return Math.atan2(y, x); },
    SENOH: (a) => Math.sinh(num(a[0])),
    COSH: (a) => Math.cosh(num(a[0])),
    TANH: (a) => Math.tanh(num(a[0])),
    PROMEDIOA: (a) => { const n = valoresA(a); return n.length ? media(n) : E("#DIV/0!"); },
    MAXA: (a) => { const n = valoresA(a); return n.length ? Math.max(...n) : 0; },
    MINA: (a) => { const n = valoresA(a); return n.length ? Math.min(...n) : 0; },
    // estadística
    PERCENTIL: (a) => percentil(a[0], num(a[1])),
    CUARTIL: (a) => { const q = Math.floor(num(a[1])); if (q < 0 || q > 4) throw E("#NUM!"); return percentil(a[0], q / 4); },
    "MEDIA.GEOM": (a) => { const n = nums(a); if (!n.length || n.some((v) => v <= 0)) throw E("#NUM!"); return Math.exp(media(n.map(Math.log))); },
    "MEDIA.ARMO": (a) => { const n = nums(a); if (!n.length || n.some((v) => v <= 0)) throw E("#NUM!"); return n.length / n.reduce((x, y) => x + 1 / y, 0); },
    DESVPROM: (a) => { const n = nums(a); if (!n.length) throw E("#NUM!"); const m = media(n); return media(n.map((v) => Math.abs(v - m))); },
    PENDIENTE: (a) => recta(a[0], a[1]).m,
    "INTERSECCION.EJE": (a) => recta(a[0], a[1]).b,
    PRONOSTICO: (a) => { const r = recta(a[1], a[2]); return r.b + r.m * num(a[0]); },
    TENDENCIA: (a) => { const r = recta(a[0], a[1]), x = a[2] === undefined ? (a[1] === undefined ? 1 : lista(a[1]).find(esNum)) : (Array.isArray(a[2]) ? a[2][0] : a[2]); return r.b + r.m * num(x); },
    "COEF.DE.CORREL": (a) => { const q = pares(a[0], a[1]); if (q.sxx === 0 || q.syy === 0) throw E("#DIV/0!"); return q.sxy / Math.sqrt(q.sxx * q.syy); },
    "COEFICIENTE.R2": (a) => { const q = pares(a[0], a[1]); if (q.sxx === 0 || q.syy === 0) throw E("#DIV/0!"); return (q.sxy * q.sxy) / (q.sxx * q.syy); },
    "COVARIANZA.P": (a) => { const q = pares(a[0], a[1]); return q.sxy / q.n; },
    "COVARIANZA.M": (a) => { const q = pares(a[0], a[1]); if (q.n < 2) throw E("#DIV/0!"); return q.sxy / (q.n - 1); },
    "JERARQUIA.MEDIA": (a) => {
      const n = num(a[0]), l = nums([a[1]]), asc = a[2] !== undefined && verdad(a[2]) && a[2] !== 0, iguales = l.filter((x) => x === n).length;
      if (!iguales) return E("#N/A");
      return l.filter((x) => (asc ? x < n : x > n)).length + (iguales + 1) / 2;
    },
    // financieras
    PAGO: (a) => fin.pago(num(a[0]), num(a[1]), num(a[2]), ...args5(a, 3)),
    VF: (a) => fin.vf(num(a[0]), num(a[1]), num(oDef(a[2], 0)), ...args5(a, 3)),
    VA: (a) => fin.va(num(a[0]), num(a[1]), num(oDef(a[2], 0)), ...args5(a, 3)),
    NPER: (a) => {
      const r = num(a[0]), pago = num(a[1]), va = num(a[2]), [vf, tipo] = args5(a, 3);
      if (r === 0) { if (pago === 0) throw E("#NUM!"); return -(va + vf) / pago; }
      const x = (pago * (1 + r * tipo) - vf * r) / (pago * (1 + r * tipo) + va * r);
      if (!(x > 0)) throw E("#NUM!");
      return Math.log(x) / Math.log(1 + r);
    },
    TASA: (a) => {
      const n = num(a[0]), pago = num(a[1]), va = num(a[2]), [vf, tipo] = args5(a, 3), x0 = a[5] === undefined ? 0.1 : num(a[5]);
      return raiz((r) => (Math.abs(r) < 1e-10 ? va + pago * n + vf : va * Math.pow(1 + r, n) + pago * (1 + r * tipo) * (Math.pow(1 + r, n) - 1) / r + vf), x0);
    },
    VNA: (a) => { const r = num(a[0]); return nums(a.slice(1)).reduce((s2, v, i) => s2 + v / Math.pow(1 + r, i + 1), 0); },
    TIR: (a) => {
      const v = nums([a[0]]);
      if (!v.some((x) => x > 0) || !v.some((x) => x < 0)) throw E("#NUM!");
      return raiz((r) => v.reduce((s2, x, i) => s2 + x / Math.pow(1 + r, i), 0), a[1] === undefined ? 0.1 : num(a[1]));
    },
    PAGOINT: (a) => fin.interes(num(a[0]), num(a[1]), num(a[2]), num(a[3]), ...args5(a, 4)),
    PAGOPRIN: (a) => { const r = num(a[0]), per = num(a[1]), n = num(a[2]), va = num(a[3]), [vf, tipo] = args5(a, 4); return fin.pago(r, n, va, vf, tipo) - fin.interes(r, per, n, va, vf, tipo); },
    SLN: (a) => { const v = num(a[2]); if (v === 0) throw E("#DIV/0!"); return (num(a[0]) - num(a[1])) / v; },
    "INT.EFECTIVO": (a) => { const t = num(a[0]), n = Math.floor(num(a[1])); if (t <= 0 || n < 1) throw E("#NUM!"); return Math.pow(1 + t / n, n) - 1; },
    "TASA.NOMINAL": (a) => { const t = num(a[0]), n = Math.floor(num(a[1])); if (t <= 0 || n < 1) throw E("#NUM!"); return n * (Math.pow(1 + t, 1 / n) - 1); },
    // referencias (FILA, COLUMNA, FILAS, COLUMNAS, INDIRECTO y DESREF necesitan la referencia misma: van en el evaluador)
    DIRECCION: (a) => {
      const f = Math.floor(num(a[0])), c = Math.floor(num(a[1])), abs = a[2] === undefined ? 1 : Math.floor(num(a[2]));
      if (f < 1 || c < 1 || abs < 1 || abs > 4) throw E("#VALOR!");
      const r = (abs === 1 || abs === 3 ? "$" : "") + colLetra(c - 1) + (abs === 1 || abs === 2 ? "$" : "") + f;
      return (a[4] !== undefined && texto(a[4]) ? prefijoHoja(texto(a[4])) : "") + r;
    },
  });

  // ------------------------------------------------------------ matrices dinámicas (como Excel 365)
  // Una matriz es un arreglo plano por filas con su número de columnas en .cols (igual que un rango). Las
  // operaciones (+ - * / ^ & y las comparaciones) y las funciones de un valor se aplican elemento por elemento:
  // una fila o una columna sola se repite y lo que no alcanza da #N/A. La celda cuya fórmula da una matriz la
  // «derrama» a la derecha y abajo (el evaluador); si una de esas celdas tiene algo, da #DESBORDAMIENTO!.
  const colsDe = (v) => (Array.isArray(v) ? v.cols || 1 : 1);
  const filasDe = (v) => (Array.isArray(v) ? Math.ceil(v.length / colsDe(v)) : 1);
  function matriz(filas, cols, fn) {
    const out = new Array(filas * cols);
    out.cols = cols;
    for (let r = 0; r < filas; r++) for (let c = 0; c < cols; c++) out[r * cols + c] = fn(r, c);
    return out;
  }
  /** Una matriz de una lista de filas (las cortas se completan con #N/A, como APILARV). */
  const deFilas = (filas) => { const cols = filas.reduce((m, q) => Math.max(m, q.length), 0); return matriz(filas.length, cols, (r, c) => (c < filas[r].length ? filas[r][c] : E("#N/A"))); };
  const enM = (v, r, c) => (Array.isArray(v) ? v[r * colsDe(v) + c] : v);
  const filaDe = (v, r) => { const k = colsDe(v), o = lista(v).slice(r * k, r * k + k); o.cols = k; return o; };
  const columnaDe = (v, c) => { const o = []; for (let r = 0, n = filasDe(v); r < n; r++) o.push(enM(v, r, c)); o.cols = 1; return o; };
  function transponer(v) { if (!Array.isArray(v)) return v; const n = filasDe(v), k = colsDe(v); return matriz(k, n, (r, c) => v[c * k + r]); }
  /** El elemento (r, c) de un operando, con la regla de Excel para las formas distintas. */
  function elem(v, r, c) {
    if (!Array.isArray(v)) return v;
    const n = filasDe(v), k = colsDe(v), rr = n === 1 ? 0 : r, cc = k === 1 ? 0 : c;
    return rr < n && cc < k ? v[rr * k + cc] : E("#N/A");
  }
  const intenta = (fn) => { try { return fn(); } catch (e) { if (e instanceof ErrorF) return e; throw e; } };
  const unValor = (y) => (Array.isArray(y) ? (y.length === 1 ? y[0] : E("#CALC!")) : y);
  // (la celda que se calcula usó matrices: al exportar a Excel va como fórmula de matriz dinámica)
  let usoMatriz = false;
  /** fn(...valores) elemento por elemento si algún argumento es matriz. conError: los errores también llegan a fn. */
  function porElemento(args, fn, conError) {
    if (!args.some((v) => Array.isArray(v))) return fn(...args);
    usoMatriz = true;
    let filas = 1, cols = 1;
    args.forEach((v) => { if (Array.isArray(v)) { filas = Math.max(filas, filasDe(v)); cols = Math.max(cols, colsDe(v)); } });
    return matriz(filas, cols, (r, c) => {
      const xs = args.map((v) => elem(v, r, c));
      if (!conError) { const e = xs.find((x) => x instanceof ErrorF); if (e) return e; }
      return intenta(() => unValor(fn(...xs)));
    });
  }
  // Funciones de un solo valor que con una matriz dan una matriz (Excel 365): true = todos sus argumentos;
  // una lista = solo esos (los demás son rangos que la función usa completos).
  const LEVANTA = {};
  ("REDONDEAR REDONDEAR.MAS REDONDEAR.MENOS TRUNCAR ENTERO RESIDUO ABS RAIZ POTENCIA SIGNO EXP LN LOG LOG10 SENO COS TAN GRADOS RADIANES " +
    "ESERROR ESBLANCO ESNUMERO ESTEXTO NO SIERROR MAYUSC MINUSC NOMPROPIO LARGO IZQUIERDA DERECHA EXTRAE ESPACIOS SUSTITUIR REEMPLAZAR " +
    "ENCONTRAR HALLAR TEXTO VALOR REPETIR IGUAL CARACTER CONCATENAR FECHA AÑO MES DIA DIASEM FIN.MES FECHA.MES SIFECHA DIAS HORA MINUTO SEGUNDO NSHORA " +
    "HORANUMERO FECHANUMERO NUM.DE.SEMANA ISO.NUM.DE.SEMANA DIAS360 ESNOD ESERR SI.ND ES.PAR ES.IMPAR ESLOGICO ESNOTEXTO N T TEXTOANTES TEXTODESPUES " +
    "LIMPIAR CODIGO UNICODE UNICAR MONEDA DECIMAL CONVERTIR DEC.A.BIN DEC.A.OCT DEC.A.HEX BIN.A.DEC OCT.A.DEC HEX.A.DEC BASE NUMERO.ARABE " +
    "NUMERO.ROMANO MULTIPLO.SUPERIOR MULTIPLO.INFERIOR MULTIPLO.SUPERIOR.MAT MULTIPLO.INFERIOR.MAT REDOND.MULT REDONDEA.PAR REDONDEA.IMPAR FACT " +
    "COMBINAT PERMUTACIONES COCIENTE ASENO ACOS ATAN ATAN2 SENOH COSH TANH DIRECCION ALEATORIO.ENTRE PAGO VF VA NPER PAGOINT PAGOPRIN SLN " +
    "INT.EFECTIVO TASA.NOMINAL").split(" ").forEach((n) => { LEVANTA[n] = true; });
  Object.assign(LEVANTA, { BUSCARV: [0], BUSCARH: [0], BUSCARX: [0], COINCIDIR: [0], COINCIDIRX: [0], "CONTAR.SI": [1], "SUMAR.SI": [1],
    "PROMEDIO.SI": [1], INDICE: [1, 2], ELEGIR: [0], "K.ESIMO.MAYOR": [1], "K.ESIMO.MENOR": [1], JERARQUIA: [0], PERCENTIL: [1], CUARTIL: [1], CAMBIAR: [0] });
  /** Llama a una función con la regla de Excel 365 para las matrices en los argumentos de un valor. */
  function llamarFuncion(nombre, fn, args) {
    const L = nombre === "SI" ? (Array.isArray(args[0]) && args[0].length !== 1 ? [0, 1, 2] : null) : LEVANTA[nombre];
    if (!L) return fn(args);
    const a2 = args.slice(), pos = [];
    (L === true ? args.map((_, k) => k) : L).forEach((k) => {
      if (!Array.isArray(args[k])) return;
      if (args[k].length === 1) a2[k] = args[k][0]; else pos.push(k);   // (una matriz de un valor cuenta como ese valor)
    });
    if (!pos.length) return fn(a2);
    return porElemento(pos.map((k) => a2[k]), (...xs) => { const b = a2.slice(); pos.forEach((k, j) => { b[k] = xs[j]; }); const y = fn(b); if (y instanceof ErrorF) throw y; return y; }, true);
  }
  const vacioM = (v) => v === "" || v == null;
  const COLACION = typeof Intl !== "undefined" && Intl.Collator ? new Intl.Collator("es", { sensitivity: "accent" }) : null;
  /** Orden de Excel al ordenar: números < textos (sin distinguir mayúsculas) < FALSO < VERDADERO < errores. */
  const claseOrden = (v) => (v instanceof ErrorF ? 4 : typeof v === "boolean" ? 3 : esNum(v) ? 1 : 2);
  function comparaOrden(a, b, o) {
    // lo vacío siempre al final
    if (vacioM(a) || vacioM(b)) return vacioM(a) && vacioM(b) ? 0 : vacioM(a) ? 1 : -1;
    const x = claseOrden(a), y = claseOrden(b);
    let d = x - y;
    if (!d) d = x === 1 ? +a - +b : x === 2 ? (COLACION ? COLACION.compare(String(a), String(b)) : String(a).toLowerCase().localeCompare(String(b).toLowerCase())) : x === 3 ? (a ? 1 : 0) - (b ? 1 : 0) : 0;
    return d * (o < 0 ? -1 : 1);
  }
  /** Llave para comparar filas iguales (UNICOS): sin distinguir mayúsculas, como Excel. */
  const llaveUnica = (fila) => fila.map((v) => (v instanceof ErrorF ? "e" + v.codigo : esNum(v) ? "n" + +v : typeof v === "boolean" ? "b" + v : "t" + String(v == null ? "" : v).toLowerCase())).join("\u0001");
  const opcional = (v) => v === undefined || v === "";
  const lambdaDe = (v) => { if (!v || typeof v.llamar !== "function") throw E("#VALOR!"); return v; };
  function recortar(a, tomar) {
    const m = a[0], n = filasDe(m), k = colsDe(m);
    const rf = opcional(a[1]) ? null : Math.trunc(num(a[1])), rc = opcional(a[2]) ? null : Math.trunc(num(a[2]));
    if (tomar && (rf === 0 || rc === 0)) throw E("#CALC!");
    const tramo = (x, total) => (x == null ? [0, total] : tomar ? (x >= 0 ? [0, Math.min(x, total)] : [Math.max(0, total + x), total]) : x >= 0 ? [Math.min(x, total), total] : [0, Math.max(0, total + x)]);
    const [f1, f2] = tramo(rf, n), [c1, c2] = tramo(rc, k);
    if (f2 <= f1 || c2 <= c1) throw E("#CALC!");
    return matriz(f2 - f1, c2 - c1, (r, c) => enM(m, f1 + r, c1 + c));
  }
  function aplanar(a, comoFila) {
    const m = a[0], ign = Math.trunc(num(opcional(a[1]) ? 0 : a[1])), porCol = !opcional(a[2]) && verdad(a[2]);
    const n = filasDe(m), k = colsDe(m), out = [], deRango = Array.isArray(m) && m.f1 != null;
    const agrega = (v) => { if ((ign === 1 || ign === 3) && vacioM(v) && (deRango || v == null)) return; if ((ign === 2 || ign === 3) && v instanceof ErrorF) return; out.push(v); };
    if (porCol) { for (let c = 0; c < k; c++) for (let r = 0; r < n; r++) agrega(enM(m, r, c)); }
    else for (let r = 0; r < n; r++) for (let c = 0; c < k; c++) agrega(enM(m, r, c));
    if (!out.length) throw E("#CALC!");
    out.cols = comoFila ? out.length : 1;
    return out;
  }
  function ajustar(a, porFilas) {
    const v = a[0];
    if (filasDe(v) > 1 && colsDe(v) > 1) throw E("#VALOR!");
    const l = lista(v), w = Math.trunc(num(a[1])), pad = a[2] === undefined ? E("#N/A") : a[2];
    if (w < 1) throw E("#NUM!");
    const n = Math.ceil(l.length / w);
    return porFilas ? matriz(n, w, (r, c) => (r * w + c < l.length ? l[r * w + c] : pad)) : matriz(w, n, (r, c) => (c * w + r < l.length ? l[c * w + r] : pad));
  }
  function elegirPiezas(a, cols) {
    const m = a[0], n = filasDe(m), k = colsDe(m), total = cols ? k : n;
    const idx = plano(a.slice(1).filter((x) => x !== undefined)).map((v) => { let j = Math.trunc(num(v)); if (j < 0) j = total + j + 1; if (j < 1 || j > total) throw E("#VALOR!"); return j - 1; });
    if (!idx.length) throw E("#VALOR!");
    return cols ? matriz(n, idx.length, (r, c) => enM(m, r, idx[c])) : matriz(idx.length, k, (r, c) => enM(m, idx[r], c));
  }
  /** Determinante e inversa por eliminación de Gauss (MDETERM y MINVERSA). */
  function gauss(v, inversa) {
    const n = filasDe(v);
    if (n !== colsDe(v) || lista(v).some((x) => !esNum(x))) throw E("#VALOR!");
    const A = Array.from({ length: n }, (_, r) => Array.from({ length: n }, (_, c) => +enM(v, r, c)).concat(inversa ? Array.from({ length: n }, (_, c) => (r === c ? 1 : 0)) : []));
    let det = 1;
    for (let c = 0; c < n; c++) {
      let p = c;
      for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r;
      if (Math.abs(A[p][c]) < 1e-15) { if (inversa) throw E("#NUM!"); return 0; }
      if (p !== c) { [A[p], A[c]] = [A[c], A[p]]; det = -det; }
      det *= A[c][c];
      for (let r = 0; r < n; r++) if (r !== c) { const k = A[r][c] / A[c][c]; for (let j = c; j < A[r].length; j++) A[r][j] -= k * A[c][j]; }
    }
    if (!inversa) return det;
    return matriz(n, n, (r, c) => A[r][n + c] / A[r][r]);
  }
  Object.assign(FUNCS, {
    ORDENAR: (a) => {
      const m = a[0], n = filasDe(m), k = colsDe(m), porCol = !opcional(a[3]) && verdad(a[3]);
      const idx = opcional(a[1]) ? [1] : lista(a[1]).map((v) => Math.trunc(num(v)));
      const ord = opcional(a[2]) ? [1] : lista(a[2]).map((v) => num(v));
      if (idx.some((j) => j < 1 || j > (porCol ? n : k)) || ord.some((o) => o !== 1 && o !== -1)) throw E("#VALOR!");
      const piezas = Array.from({ length: porCol ? k : n }, (_, j) => (porCol ? columnaDe(m, j) : filaDe(m, j)));
      const orden = piezas.map((_, j) => j).sort((p, q) => { for (let j = 0; j < idx.length; j++) { const d = comparaOrden(piezas[p][idx[j] - 1], piezas[q][idx[j] - 1], ord[j] !== undefined ? ord[j] : ord[0]); if (d) return d; } return 0; });
      const out = deFilas(orden.map((j) => Array.from(piezas[j])));
      return porCol ? transponer(out) : out;
    },
    ORDENARPOR: (a) => {
      const m = a[0], n = filasDe(m), k = colsDe(m), claves = [];
      for (let j = 1; j < a.length; j += 2) {
        const o = opcional(a[j + 1]) ? 1 : num(a[j + 1]);
        if (a[j] === undefined || (o !== 1 && o !== -1)) throw E("#VALOR!");
        if (a[j] instanceof ErrorF) throw a[j];
        claves.push({ l: lista(a[j]), f: filasDe(a[j]), c: colsDe(a[j]), o });
      }
      if (!claves.length) throw E("#VALOR!");
      const vertical = claves.every((q) => q.c === 1 && q.f === n), horizontal = !vertical && claves.every((q) => q.f === 1 && q.c === k);
      if (!vertical && !horizontal) throw E("#VALOR!");
      const orden = Array.from({ length: vertical ? n : k }, (_, j) => j).sort((p, q) => { for (const x of claves) { const d = comparaOrden(x.l[p], x.l[q], x.o); if (d) return d; } return 0; });
      return vertical ? deFilas(orden.map((j) => Array.from(filaDe(m, j)))) : transponer(deFilas(orden.map((j) => Array.from(columnaDe(m, j)))));
    },
    FILTRAR: (a) => {
      const m = a[0], n = filasDe(m), k = colsDe(m), inc = a[1];
      if (inc instanceof ErrorF) throw inc;
      const l = lista(inc), fi = filasDe(inc), ci = colsDe(inc);
      const pasa = (v) => { if (v instanceof ErrorF) throw v; if (typeof v === "boolean") return v; if (esNum(v)) return +v !== 0; if (vacioM(v)) return false; throw E("#VALOR!"); };
      const sel = [];
      if (ci === 1 && fi === n) { for (let r = 0; r < n; r++) if (pasa(l[r])) sel.push(Array.from(filaDe(m, r))); if (sel.length) return deFilas(sel); }
      else if (fi === 1 && ci === k) { for (let c = 0; c < k; c++) if (pasa(l[c])) sel.push(Array.from(columnaDe(m, c))); if (sel.length) return transponer(deFilas(sel)); }
      else throw E("#VALOR!");
      if (a[2] !== undefined) return a[2];
      throw E("#CALC!");
    },
    UNICOS: (a) => {
      const m = a[0], porCol = !opcional(a[1]) && verdad(a[1]), unaVez = !opcional(a[2]) && verdad(a[2]);
      const piezas = Array.from({ length: porCol ? colsDe(m) : filasDe(m) }, (_, j) => Array.from(porCol ? columnaDe(m, j) : filaDe(m, j)));
      const cuenta = new Map();
      piezas.forEach((p) => { const q = llaveUnica(p); cuenta.set(q, (cuenta.get(q) || 0) + 1); });
      const vistos = new Set(), out = [];
      piezas.forEach((p) => { const q = llaveUnica(p); if (vistos.has(q)) return; vistos.add(q); if (!unaVez || cuenta.get(q) === 1) out.push(p); });
      if (!out.length) throw E("#CALC!");
      return porCol ? transponer(deFilas(out)) : deFilas(out);
    },
    SECUENCIA: (a) => {
      const f = Math.trunc(num(opcional(a[0]) ? 1 : a[0])), c = Math.trunc(num(opcional(a[1]) ? 1 : a[1]));
      const ini = num(opcional(a[2]) ? 1 : a[2]), paso = num(opcional(a[3]) ? 1 : a[3]);
      if (f === 0 || c === 0) throw E("#CALC!");
      if (f < 0 || c < 0) throw E("#VALOR!");
      if (f * c > 2000000) throw E("#NUM!");
      return matriz(f, c, (r, k) => ini + (r * c + k) * paso);
    },
    MATRIZALEAT: (a) => {
      const f = Math.trunc(num(opcional(a[0]) ? 1 : a[0])), c = Math.trunc(num(opcional(a[1]) ? 1 : a[1]));
      const lo = num(opcional(a[2]) ? 0 : a[2]), hi = num(opcional(a[3]) ? 1 : a[3]), ent = !opcional(a[4]) && verdad(a[4]);
      if (f < 1 || c < 1 || hi < lo || f * c > 2000000) throw E("#VALOR!");
      return matriz(f, c, () => (ent ? Math.ceil(lo) + Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + Math.random() * (hi - lo)));
    },
    TRANSPONER: (a) => transponer(a[0]),
    APILARV: (a) => {
      const ps = a.filter((v) => v !== undefined), filas = [];
      if (!ps.length) throw E("#VALOR!");
      ps.forEach((v) => { for (let r = 0; r < filasDe(v); r++) filas.push(Array.from(filaDe(v, r))); });
      return deFilas(filas);
    },
    APILARH: (a) => {
      const ps = a.filter((v) => v !== undefined);
      if (!ps.length) throw E("#VALOR!");
      const n = ps.reduce((m, v) => Math.max(m, filasDe(v)), 0), filas = [];
      for (let r = 0; r < n; r++) { const fila = []; ps.forEach((v) => { for (let c = 0; c < colsDe(v); c++) fila.push(r < filasDe(v) ? enM(v, r, c) : E("#N/A")); }); filas.push(fila); }
      return deFilas(filas);
    },
    TOMAR: (a) => recortar(a, true),
    EXCLUIR: (a) => recortar(a, false),
    ELEGIRCOLS: (a) => elegirPiezas(a, true),
    ELEGIRFILAS: (a) => elegirPiezas(a, false),
    ENCOL: (a) => aplanar(a, false),
    ENFILA: (a) => aplanar(a, true),
    AJUSTARFILAS: (a) => ajustar(a, true),
    AJUSTARCOLS: (a) => ajustar(a, false),
    EXPANDIR: (a) => {
      const m = a[0], n = filasDe(m), k = colsDe(m);
      const f = opcional(a[1]) ? n : Math.trunc(num(a[1])), c = opcional(a[2]) ? k : Math.trunc(num(a[2]));
      if (f < n || c < k) throw E("#VALOR!");
      const pad = a[3] === undefined ? E("#N/A") : a[3];
      return matriz(f, c, (r, q) => (r < n && q < k ? enM(m, r, q) : pad));
    },
    DIVIDIRTEXTO: (a) => {
      const t = texto(a[0]), sinMay = !opcional(a[4]) && num(a[4]) === 1, ign = !opcional(a[3]) && verdad(a[3]);
      const delims = (v) => (opcional(v) ? [] : plano([v]).map(texto).filter((d) => d !== "").sort((x, y) => y.length - x.length));
      const dc = delims(a[1]), df = delims(a[2]);
      if (!dc.length && !df.length) throw E("#VALOR!");
      const pad = a[5] === undefined ? E("#N/A") : a[5];
      const partir = (s, ds) => (ds.length ? s.split(new RegExp(ds.map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|"), sinMay ? "gi" : "g")) : [s]);
      let filas = partir(t, df).map((s) => partir(s, dc));
      if (ign) filas = filas.map((q) => q.filter((x) => x !== "")).filter((q) => q.length);
      if (!filas.length) throw E("#CALC!");
      return matriz(filas.length, Math.max(...filas.map((q) => q.length)), (r, c) => (c < filas[r].length ? filas[r][c] : pad));
    },
    FRECUENCIA: (a) => {
      const datos = nums([a[0]]), lim = nums([a[1]]), orden = lim.map((b, j) => [b, j]).sort((x, y) => x[0] - y[0]);
      const cuenta = new Array(lim.length + 1).fill(0);
      datos.forEach((d) => { const j = orden.findIndex(([b]) => d <= b); cuenta[j < 0 ? lim.length : orden[j][1]]++; });
      cuenta.cols = 1;
      return cuenta;
    },
    MMULT: (a) => {
      const x = a[0], y = a[1], n = filasDe(x), k = colsDe(x), m = colsDe(y);
      if (k !== filasDe(y) || lista(x).concat(lista(y)).some((v) => !esNum(v))) throw E("#VALOR!");
      return matriz(n, m, (r, c) => { let s = 0; for (let j = 0; j < k; j++) s += +enM(x, r, j) * +enM(y, j, c); return s; });
    },
    "M.UNIDAD": (a) => { const n = Math.trunc(num(a[0])); if (n < 1) throw E("#VALOR!"); return matriz(n, n, (r, c) => (r === c ? 1 : 0)); },
    MDETERM: (a) => gauss(a[0], false),
    MINVERSA: (a) => gauss(a[0], true),
    // con LAMBDA
    BYROW: (a) => { const m = a[0], lam = lambdaDe(a[1]), out = []; for (let r = 0, n = filasDe(m); r < n; r++) out.push(intenta(() => unValor(lam.llamar([filaDe(m, r)])))); out.cols = 1; return out; },
    BYCOL: (a) => { const m = a[0], lam = lambdaDe(a[1]), out = []; for (let c = 0, k = colsDe(m); c < k; c++) out.push(intenta(() => unValor(lam.llamar([columnaDe(m, c)])))); out.cols = out.length; return out; },
    MAP: (a) => {
      const lam = lambdaDe(a[a.length - 1]), ms = a.slice(0, -1);
      if (!ms.length) throw E("#VALOR!");
      return porElemento(ms.map((v) => (Array.isArray(v) || v === undefined ? v : [v])), (...xs) => unValor(lam.llamar(xs)), true);
    },
    REDUCE: (a) => {
      const lam = lambdaDe(a[2]), m = a[1];
      let acc = a[0] === undefined ? "" : a[0];
      for (let r = 0, n = filasDe(m); r < n; r++) for (let c = 0, k = colsDe(m); c < k; c++) acc = lam.llamar([acc, enM(m, r, c)]);
      return acc;
    },
    SCAN: (a) => {
      const lam = lambdaDe(a[2]), m = a[1];
      let acc = a[0] === undefined ? "" : a[0];
      return matriz(filasDe(m), colsDe(m), (r, c) => { acc = intenta(() => unValor(lam.llamar([acc, enM(m, r, c)]))); return acc; });
    },
    MAKEARRAY: (a) => {
      const f = Math.trunc(num(a[0])), c = Math.trunc(num(a[1])), lam = lambdaDe(a[2]);
      if (f < 1 || c < 1 || f * c > 2000000) throw E("#VALOR!");
      return matriz(f, c, (r, k) => intenta(() => unValor(lam.llamar([r + 1, k + 1]))));
    },
    // INDICE y BUSCARX como en Excel 365: la fila o la columna entera (con 0 o sin columna) y la fila de lo encontrado
    INDICE: (a) => {
      const m = a[0], n = filasDe(m), k = colsDe(m);
      let f = opcional(a[1]) ? null : Math.floor(num(a[1])), c = opcional(a[2]) ? null : Math.floor(num(a[2]));
      if (n === 1 && c === null && f !== null) { c = f; f = 1; }   // (una sola fila: el número es la columna)
      if (k === 1 && c === null) c = 1;
      if (f === null) f = n === 1 ? 1 : 0;
      if (c === null) c = k === 1 ? 1 : 0;
      if (f < 0 || c < 0 || f > n || c > k) return E("#REF!");
      if (f === 0 && c === 0) return m;
      if (f === 0) { usoMatriz = true; return n === 1 ? enM(m, 0, c - 1) : columnaDe(m, c - 1); }
      if (c === 0) { usoMatriz = true; return k === 1 ? enM(m, f - 1, 0) : filaDe(m, f - 1); }
      return enM(m, f - 1, c - 1);
    },
    BUSCARX: (a) => {
      const donde = a[1], que = a[2], ld = lista(donde), vertical = colsDe(donde) === 1;
      if (filasDe(donde) > 1 && colsDe(donde) > 1) throw E("#VALOR!");
      const modo = opcional(a[4]) ? 0 : Math.trunc(num(a[4])), busq = opcional(a[5]) ? 1 : Math.trunc(num(a[5]));
      const p = FUNCS.COINCIDIRX([a[0], donde, modo, busq < 0 ? -1 : 1]);
      if (p instanceof ErrorF) return a[3] === undefined ? E("#N/A") : a[3];
      const i = p - 1;
      if (vertical) {
        if (filasDe(que) !== ld.length) throw E("#VALOR!");
        return colsDe(que) === 1 ? enM(que, i, 0) : filaDe(que, i);
      }
      if (colsDe(que) !== ld.length) throw E("#VALOR!");
      return filasDe(que) === 1 ? enM(que, 0, i) : columnaDe(que, i);
    },
    // (las resuelve el evaluador; aquí solo para que se reconozcan y salgan en la ayuda)
    LET: () => E("#VALOR!"),
    LAMBDA: () => E("#CALC!"),
    ISOMITTED: () => E("#VALOR!"),
  });
  /** Funciones que solo existen con matrices dinámicas: la fórmula se guarda en Excel como fórmula de matriz. */
  const DINAMICAS = new Set(["ORDENAR", "ORDENARPOR", "FILTRAR", "UNICOS", "SECUENCIA", "MATRIZALEAT", "APILARV", "APILARH", "TOMAR", "EXCLUIR",
    "ELEGIRCOLS", "ELEGIRFILAS", "ENCOL", "ENFILA", "AJUSTARFILAS", "AJUSTARCOLS", "EXPANDIR", "DIVIDIRTEXTO", "BYROW", "BYCOL", "MAP", "REDUCE",
    "SCAN", "MAKEARRAY"]);
  /** Funciones que trabajan con la referencia y no con su valor (las resuelve el evaluador). */
  const DE_REFERENCIA = new Set(["FILA", "COLUMNA", "FILAS", "COLUMNAS", "INDIRECTO", "DESREF"]);
  DE_REFERENCIA.forEach((n) => { FUNCS[n] = () => E("#VALOR!"); });
  const ALIAS = {
    SUM: "SUMA", AVERAGE: "PROMEDIO", COUNT: "CONTAR", COUNTA: "CONTARA", COUNTBLANK: "CONTAR.BLANCO", PRODUCT: "PRODUCTO", SUMPRODUCT: "SUMAPRODUCTO",
    IF: "SI", IFS: "SI.CONJUNTO", AND: "Y", OR: "O", NOT: "NO", ISERROR: "ESERROR", IFERROR: "SIERROR", ISBLANK: "ESBLANCO", ISNUMBER: "ESNUMERO", ISTEXT: "ESTEXTO",
    ROUND: "REDONDEAR", ROUNDUP: "REDONDEAR.MAS", ROUNDDOWN: "REDONDEAR.MENOS", TRUNC: "TRUNCAR", INT: "ENTERO", MOD: "RESIDUO", SQRT: "RAIZ", POWER: "POTENCIA",
    SIGN: "SIGNO", SIN: "SENO", DEGREES: "GRADOS", RADIANS: "RADIANES", RAND: "ALEATORIO", RANDBETWEEN: "ALEATORIO.ENTRE",
    MEDIAN: "MEDIANA", MODE: "MODA", "MODE.SNGL": "MODA", "MODA.UNO": "MODA", STDEV: "DESVEST", "STDEV.S": "DESVEST", "DESVEST.M": "DESVEST", STDEVP: "DESVESTP",
    "STDEV.P": "DESVESTP", "DESVEST.P": "DESVESTP", "VAR.S": "VAR", "VAR.M": "VAR", "VAR.P": "VARP", RANK: "JERARQUIA", "RANK.EQ": "JERARQUIA", "JERARQUIA.EQV": "JERARQUIA",
    "JERARQUÍA": "JERARQUIA", LARGE: "K.ESIMO.MAYOR", SMALL: "K.ESIMO.MENOR",
    COUNTIF: "CONTAR.SI", SUMIF: "SUMAR.SI", AVERAGEIF: "PROMEDIO.SI", COUNTIFS: "CONTAR.SI.CONJUNTO", SUMIFS: "SUMAR.SI.CONJUNTO", AVERAGEIFS: "PROMEDIO.SI.CONJUNTO",
    MAXIFS: "MAX.SI.CONJUNTO", MINIFS: "MIN.SI.CONJUNTO",
    CONCATENATE: "CONCATENAR", TEXTJOIN: "UNIRCADENAS", UPPER: "MAYUSC", LOWER: "MINUSC", PROPER: "NOMPROPIO", LEN: "LARGO", LEFT: "IZQUIERDA", RIGHT: "DERECHA",
    MID: "EXTRAE", TRIM: "ESPACIOS", SUBSTITUTE: "SUSTITUIR", REPLACE: "REEMPLAZAR", FIND: "ENCONTRAR", SEARCH: "HALLAR", TEXT: "TEXTO", VALUE: "VALOR", REPT: "REPETIR",
    EXACT: "IGUAL", CHAR: "CARACTER", "CARÁCTER": "CARACTER",
    TODAY: "HOY", NOW: "AHORA", DATE: "FECHA", DAYS: "DIAS", "DÍAS": "DIAS", YEAR: "AÑO", MONTH: "MES", DAY: "DIA", "DÍA": "DIA", WEEKDAY: "DIASEM", EOMONTH: "FIN.MES",
    EDATE: "FECHA.MES", NETWORKDAYS: "DIAS.LAB", "DÍAS.LAB": "DIAS.LAB", DATEDIF: "SIFECHA",
    VLOOKUP: "BUSCARV", HLOOKUP: "BUSCARH", XLOOKUP: "BUSCARX", INDEX: "INDICE", "ÍNDICE": "INDICE", MATCH: "COINCIDIR", CHOOSE: "ELEGIR",
    // 0.13.2
    HOUR: "HORA", MINUTE: "MINUTO", SECOND: "SEGUNDO", TIME: "NSHORA", TIMEVALUE: "HORANUMERO", DATEVALUE: "FECHANUMERO", "FECHANÚMERO": "FECHANUMERO", "HORANÚMERO": "HORANUMERO",
    WEEKNUM: "NUM.DE.SEMANA", ISOWEEKNUM: "ISO.NUM.DE.SEMANA", "ISONUM.DE.SEMANA": "ISO.NUM.DE.SEMANA", DAYS360: "DIAS360", "DÍAS360": "DIAS360", WORKDAY: "DIA.LAB", "DÍA.LAB": "DIA.LAB",
    ISNA: "ESNOD", ISERR: "ESERR", NA: "NOD", IFNA: "SI.ND", ISEVEN: "ES.PAR", ISODD: "ES.IMPAR", ISLOGICAL: "ESLOGICO", "ESLÓGICO": "ESLOGICO", ISNONTEXT: "ESNOTEXTO",
    SWITCH: "CAMBIAR", XOR: "XO", TYPE: "TIPO",
    TEXTBEFORE: "TEXTOANTES", TEXTAFTER: "TEXTODESPUES", "TEXTODESPUÉS": "TEXTODESPUES", CLEAN: "LIMPIAR", CODE: "CODIGO", "CÓDIGO": "CODIGO", UNICHAR: "UNICAR",
    DOLLAR: "MONEDA", FIXED: "DECIMAL", FIJO: "DECIMAL", ROMAN: "NUMERO.ROMANO", "NÚMERO.ROMANO": "NUMERO.ROMANO",
    CEILING: "MULTIPLO.SUPERIOR", FLOOR: "MULTIPLO.INFERIOR", "MÚLTIPLO.SUPERIOR": "MULTIPLO.SUPERIOR", "MÚLTIPLO.INFERIOR": "MULTIPLO.INFERIOR",
    "CEILING.MATH": "MULTIPLO.SUPERIOR.MAT", "FLOOR.MATH": "MULTIPLO.INFERIOR.MAT", "MÚLTIPLO.SUPERIOR.MAT": "MULTIPLO.SUPERIOR.MAT", "MÚLTIPLO.INFERIOR.MAT": "MULTIPLO.INFERIOR.MAT",
    MROUND: "REDOND.MULT", EVEN: "REDONDEA.PAR", ODD: "REDONDEA.IMPAR", GCD: "M.C.D", LCM: "M.C.M", COMBIN: "COMBINAT", PERMUT: "PERMUTACIONES", QUOTIENT: "COCIENTE",
    SUMSQ: "SUMA.CUADRADOS", ASIN: "ASENO", SINH: "SENOH", AVERAGEA: "PROMEDIOA",
    PERCENTILE: "PERCENTIL", "PERCENTILE.INC": "PERCENTIL", "PERCENTIL.INC": "PERCENTIL", QUARTILE: "CUARTIL", "QUARTILE.INC": "CUARTIL", "CUARTIL.INC": "CUARTIL",
    GEOMEAN: "MEDIA.GEOM", HARMEAN: "MEDIA.ARMO", AVEDEV: "DESVPROM", SLOPE: "PENDIENTE", INTERCEPT: "INTERSECCION.EJE", "INTERSECCIÓN.EJE": "INTERSECCION.EJE",
    FORECAST: "PRONOSTICO", "FORECAST.LINEAR": "PRONOSTICO", "PRONOSTICO.LINEAL": "PRONOSTICO", "PRONÓSTICO": "PRONOSTICO", "PRONÓSTICO.LINEAL": "PRONOSTICO", TREND: "TENDENCIA",
    CORREL: "COEF.DE.CORREL", PEARSON: "COEF.DE.CORREL", RSQ: "COEFICIENTE.R2", "COVARIANCE.P": "COVARIANZA.P", COVAR: "COVARIANZA.P", "COVARIANCE.S": "COVARIANZA.M",
    "RANK.AVG": "JERARQUIA.MEDIA", "JERARQUÍA.MEDIA": "JERARQUIA.MEDIA",
    PMT: "PAGO", PV: "VA", FV: "VF", RATE: "TASA", NPV: "VNA", IRR: "TIR", IPMT: "PAGOINT", PPMT: "PAGOPRIN", EFFECT: "INT.EFECTIVO", NOMINAL: "TASA.NOMINAL",
    ROW: "FILA", COLUMN: "COLUMNA", ROWS: "FILAS", COLUMNS: "COLUMNAS", ADDRESS: "DIRECCION", "DIRECCIÓN": "DIRECCION", INDIRECT: "INDIRECTO", OFFSET: "DESREF",
    // 16-nov-2026
    SUBTOTAL: "SUBTOTALES", AGGREGATE: "AGREGAR", XMATCH: "COINCIDIRX", CONVERT: "CONVERTIR", DEC2BIN: "DEC.A.BIN", DEC2OCT: "DEC.A.OCT", DEC2HEX: "DEC.A.HEX",
    BIN2DEC: "BIN.A.DEC", OCT2DEC: "OCT.A.DEC", HEX2DEC: "HEX.A.DEC", ARABIC: "NUMERO.ARABE", "NÚMERO.ARABE": "NUMERO.ARABE",
    // matrices dinámicas (17-nov-2026)
    SORT: "ORDENAR", SORTBY: "ORDENARPOR", FILTER: "FILTRAR", UNIQUE: "UNICOS", "ÚNICOS": "UNICOS", SEQUENCE: "SECUENCIA", RANDARRAY: "MATRIZALEAT",
    TRANSPOSE: "TRANSPONER", VSTACK: "APILARV", HSTACK: "APILARH", TAKE: "TOMAR", DROP: "EXCLUIR", CHOOSECOLS: "ELEGIRCOLS", CHOOSEROWS: "ELEGIRFILAS",
    TOCOL: "ENCOL", TOROW: "ENFILA", WRAPROWS: "AJUSTARFILAS", WRAPCOLS: "AJUSTARCOLS", EXPAND: "EXPANDIR", TEXTSPLIT: "DIVIDIRTEXTO",
    FREQUENCY: "FRECUENCIA", MUNIT: "M.UNIDAD", MINVERSE: "MINVERSA",
  };
  const NOMBRES = Object.keys(FUNCS);
  /** Nombre en inglés para Excel (.xlsx); las funciones nuevas de Excel llevan el prefijo _xlfn. */
  const INGLES = {};
  for (const [en, es] of Object.entries(ALIAS)) if (!INGLES[es] && /^[A-Z.]+$/.test(en) && !/^(MODE\.SNGL|STDEV\.S|STDEV\.P|VAR\.S|VAR\.P|RANK\.EQ)$/.test(en)) INGLES[es] = en;
  Object.assign(INGLES, { MIN: "MIN", MAX: "MAX", ABS: "ABS", PI: "PI", EXP: "EXP", LN: "LN", LOG: "LOG", LOG10: "LOG10", COS: "COS", TAN: "TAN",
    "SI.CONJUNTO": "_xlfn.IFS", BUSCARX: "_xlfn.XLOOKUP", CONCAT: "_xlfn.CONCAT", UNIRCADENAS: "_xlfn.TEXTJOIN", "MAX.SI.CONJUNTO": "_xlfn.MAXIFS",
    "MIN.SI.CONJUNTO": "_xlfn.MINIFS", DIAS: "_xlfn.DAYS", CONCATENAR: "CONCATENATE", DESVEST: "STDEV", DESVESTP: "STDEVP", VAR: "VAR", VARP: "VARP", JERARQUIA: "RANK", MODA: "MODE" });
  Object.assign(INGLES, { SUBTOTALES: "SUBTOTAL", AGREGAR: "_xlfn.AGGREGATE", COINCIDIRX: "_xlfn.XMATCH", CONVERTIR: "CONVERT", "DEC.A.BIN": "DEC2BIN", "DEC.A.OCT": "DEC2OCT",
    "DEC.A.HEX": "DEC2HEX", "BIN.A.DEC": "BIN2DEC", "OCT.A.DEC": "OCT2DEC", "HEX.A.DEC": "HEX2DEC", "NUMERO.ARABE": "_xlfn.ARABIC", BASE: "_xlfn.BASE" });
  Object.assign(INGLES, { DIAS360: "DAYS360", "ISO.NUM.DE.SEMANA": "_xlfn.ISOWEEKNUM", "SI.ND": "_xlfn.IFNA", CAMBIAR: "_xlfn.SWITCH", XO: "_xlfn.XOR", N: "N", T: "T",
    TEXTOANTES: "_xlfn.TEXTBEFORE", TEXTODESPUES: "_xlfn.TEXTAFTER", UNICAR: "_xlfn.UNICHAR", UNICODE: "_xlfn.UNICODE", DECIMAL: "FIXED",
    "MULTIPLO.SUPERIOR.MAT": "_xlfn.CEILING.MATH", "MULTIPLO.INFERIOR.MAT": "_xlfn.FLOOR.MATH", FACT: "FACT", ACOS: "ACOS", ATAN: "ATAN", ATAN2: "ATAN2", COSH: "COSH", TANH: "TANH",
    MAXA: "MAXA", MINA: "MINA", PERCENTIL: "PERCENTILE", CUARTIL: "QUARTILE", PRONOSTICO: "FORECAST", "COEF.DE.CORREL": "CORREL", "COVARIANZA.P": "_xlfn.COVARIANCE.P",
    "COVARIANZA.M": "_xlfn.COVARIANCE.S", "JERARQUIA.MEDIA": "_xlfn.RANK.AVG", NPER: "NPER", SLN: "SLN" });
  Object.assign(INGLES, { ORDENAR: "_xlfn._xlws.SORT", FILTRAR: "_xlfn._xlws.FILTER", ORDENARPOR: "_xlfn.SORTBY", UNICOS: "_xlfn.UNIQUE",
    SECUENCIA: "_xlfn.SEQUENCE", MATRIZALEAT: "_xlfn.RANDARRAY", TRANSPONER: "TRANSPOSE", APILARV: "_xlfn.VSTACK", APILARH: "_xlfn.HSTACK",
    TOMAR: "_xlfn.TAKE", EXCLUIR: "_xlfn.DROP", ELEGIRCOLS: "_xlfn.CHOOSECOLS", ELEGIRFILAS: "_xlfn.CHOOSEROWS", ENCOL: "_xlfn.TOCOL",
    ENFILA: "_xlfn.TOROW", AJUSTARFILAS: "_xlfn.WRAPROWS", AJUSTARCOLS: "_xlfn.WRAPCOLS", EXPANDIR: "_xlfn.EXPAND", DIVIDIRTEXTO: "_xlfn.TEXTSPLIT",
    LET: "_xlfn.LET", LAMBDA: "_xlfn.LAMBDA", BYROW: "_xlfn.BYROW", BYCOL: "_xlfn.BYCOL", MAP: "_xlfn.MAP", REDUCE: "_xlfn.REDUCE", SCAN: "_xlfn.SCAN",
    MAKEARRAY: "_xlfn.MAKEARRAY", ISOMITTED: "_xlfn.ISOMITTED", FRECUENCIA: "FREQUENCY", MMULT: "MMULT", "M.UNIDAD": "_xlfn.MUNIT", MDETERM: "MDETERM",
    MINVERSA: "MINVERSE" });

  /** Ayuda al escribir: qué pide cada función y qué hace. */
  const FIRMAS = {
    SUMA: ["número1; [número2]; …", "Suma los números."], PROMEDIO: ["número1; [número2]; …", "El promedio (media) de los números."],
    MIN: ["número1; [número2]; …", "El número más chico."], MAX: ["número1; [número2]; …", "El número más grande."],
    CONTAR: ["valor1; [valor2]; …", "Cuántas celdas tienen números."], CONTARA: ["valor1; [valor2]; …", "Cuántas celdas no están vacías."],
    "CONTAR.BLANCO": ["rango", "Cuántas celdas están vacías."], PRODUCTO: ["número1; [número2]; …", "Multiplica los números."],
    SUMAPRODUCTO: ["matriz1; [matriz2]; …", "Multiplica las listas elemento por elemento y suma."],
    REDONDEAR: ["número; decimales", "Redondea a los decimales que digas."], "REDONDEAR.MAS": ["número; decimales", "Redondea hacia arriba."],
    "REDONDEAR.MENOS": ["número; decimales", "Redondea hacia abajo."], TRUNCAR: ["número; [decimales]", "Corta los decimales sin redondear."],
    ENTERO: ["número", "La parte entera (hacia abajo)."], RESIDUO: ["número; divisor", "Lo que sobra de una división."], ABS: ["número", "El valor absoluto."],
    RAIZ: ["número", "La raíz cuadrada."], POTENCIA: ["número; potencia", "Eleva un número a una potencia."], PI: ["", "El número pi (3.14159…)."],
    SIGNO: ["número", "1 si es positivo, -1 si es negativo, 0 si es cero."], EXP: ["número", "e elevado al número."], LN: ["número", "Logaritmo natural."],
    LOG: ["número; [base]", "Logaritmo (base 10 si no dices otra)."], LOG10: ["número", "Logaritmo base 10."], SENO: ["ángulo en radianes", "El seno."],
    COS: ["ángulo en radianes", "El coseno."], TAN: ["ángulo en radianes", "La tangente."], GRADOS: ["radianes", "Pasa radianes a grados."],
    RADIANES: ["grados", "Pasa grados a radianes."], ALEATORIO: ["", "Un número al azar entre 0 y 1."], "ALEATORIO.ENTRE": ["inferior; superior", "Un entero al azar entre los dos."],
    MEDIANA: ["número1; [número2]; …", "El valor de en medio."], MODA: ["número1; [número2]; …", "El valor que más se repite."],
    DESVEST: ["número1; [número2]; …", "Desviación estándar (muestra)."], DESVESTP: ["número1; [número2]; …", "Desviación estándar (población)."],
    VAR: ["número1; [número2]; …", "Varianza (muestra)."], VARP: ["número1; [número2]; …", "Varianza (población)."],
    JERARQUIA: ["número; referencia; [orden]", "El lugar de un número en la lista (1 = el más grande)."],
    "K.ESIMO.MAYOR": ["matriz; k", "El k-ésimo más grande."], "K.ESIMO.MENOR": ["matriz; k", "El k-ésimo más chico."],
    "CONTAR.SI": ["rango; criterio", "Cuántas celdas cumplen el criterio (\">5\", \"Sí\"…)."], "SUMAR.SI": ["rango; criterio; [rango_suma]", "Suma lo que cumple el criterio."],
    "PROMEDIO.SI": ["rango; criterio; [rango_promedio]", "Promedio de lo que cumple el criterio."],
    "CONTAR.SI.CONJUNTO": ["rango1; criterio1; [rango2; criterio2]; …", "Cuenta lo que cumple todos los criterios."],
    "SUMAR.SI.CONJUNTO": ["rango_suma; rango1; criterio1; …", "Suma lo que cumple todos los criterios."],
    "PROMEDIO.SI.CONJUNTO": ["rango_promedio; rango1; criterio1; …", "Promedio de lo que cumple todos los criterios."],
    "MAX.SI.CONJUNTO": ["rango_max; rango1; criterio1; …", "El más grande de lo que cumple los criterios."],
    "MIN.SI.CONJUNTO": ["rango_min; rango1; criterio1; …", "El más chico de lo que cumple los criterios."],
    SI: ["prueba; si_es_verdadero; [si_es_falso]", "Elige un valor según una condición."], "SI.CONJUNTO": ["prueba1; valor1; [prueba2; valor2]; …", "El valor de la primera prueba que se cumple."],
    ESERROR: ["valor", "VERDADERO si es un error."], SIERROR: ["valor; si_hay_error", "Otro valor si hay error."], ESBLANCO: ["valor", "VERDADERO si está vacía."],
    ESNUMERO: ["valor", "VERDADERO si es un número."], ESTEXTO: ["valor", "VERDADERO si es texto."], Y: ["lógico1; [lógico2]; …", "VERDADERO si todo se cumple."],
    O: ["lógico1; [lógico2]; …", "VERDADERO si algo se cumple."], NO: ["lógico", "Lo contrario."],
    CONCATENAR: ["texto1; [texto2]; …", "Une textos."], CONCAT: ["texto1; [texto2]; …", "Une textos."], UNIRCADENAS: ["separador; ignorar_vacías; texto1; …", "Une textos con un separador."],
    MAYUSC: ["texto", "En MAYÚSCULAS."], MINUSC: ["texto", "En minúsculas."], NOMPROPIO: ["texto", "Cada Palabra Con Mayúscula."], LARGO: ["texto", "Cuántos caracteres tiene."],
    IZQUIERDA: ["texto; [n]", "Los primeros caracteres."], DERECHA: ["texto; [n]", "Los últimos caracteres."], EXTRAE: ["texto; inicio; n", "Caracteres desde una posición."],
    ESPACIOS: ["texto", "Quita los espacios de más."], SUSTITUIR: ["texto; texto_viejo; texto_nuevo; [instancia]", "Cambia un texto por otro."],
    REEMPLAZAR: ["texto; inicio; n; texto_nuevo", "Reemplaza caracteres por posición."], ENCONTRAR: ["buscado; dentro_de; [inicio]", "Dónde está un texto (distingue mayúsculas)."],
    HALLAR: ["buscado; dentro_de; [inicio]", "Dónde está un texto."], TEXTO: ["valor; formato", "Un número con formato (\"0.00\", \"dd/mm/aaaa\")."],
    VALOR: ["texto", "Convierte texto en número."], REPETIR: ["texto; veces", "Repite un texto."], IGUAL: ["texto1; texto2", "VERDADERO si son idénticos."],
    CARACTER: ["número", "El carácter de ese código."],
    HOY: ["", "La fecha de hoy."], AHORA: ["", "La fecha y la hora de ahora."], FECHA: ["año; mes; día", "Arma una fecha."], DIAS: ["fecha_final; fecha_inicial", "Días entre dos fechas."],
    "AÑO": ["fecha", "El año."], MES: ["fecha", "El mes (1 a 12)."], DIA: ["fecha", "El día del mes."], DIASEM: ["fecha; [tipo]", "Día de la semana (1 = domingo; tipo 2: 1 = lunes)."],
    "FIN.MES": ["fecha; meses", "El último día del mes."], "FECHA.MES": ["fecha; meses", "La misma fecha unos meses después."],
 SIFECHA: ["inicio; fin; \"Y\" | \"M\" | \"D\"", "Años, meses o días entre dos fechas (sirve para la edad)."],
    BUSCARV: ["buscado; tabla; columna; [aproximada]", "Busca en la primera columna y devuelve otra (FALSO = exacta)."],
    BUSCARH: ["buscado; tabla; fila; [aproximada]", "Busca en la primera fila y devuelve otra."],
    BUSCARX: ["buscado; donde_buscar; que_devolver; [si_no_está]", "Busca y devuelve lo que está a la par."],
    INDICE: ["matriz; fila; [columna]", "El valor en esa fila y columna."], COINCIDIR: ["buscado; matriz; [tipo]", "En qué posición está (0 = exacta)."],
    ELEGIR: ["n; valor1; [valor2]; …", "El valor número n de la lista."],
    // 0.13.2
    HORA: ["hora", "La hora (0 a 23) de una fecha con hora."], MINUTO: ["hora", "Los minutos (0 a 59)."], SEGUNDO: ["hora", "Los segundos (0 a 59)."],
    NSHORA: ["hora; minuto; segundo", "Arma una hora (como fracción del día)."], HORANUMERO: ["texto", "Convierte un texto como \"14:30\" en hora."],
    FECHANUMERO: ["texto", "Convierte un texto como \"07/11/2026\" en el número de esa fecha."],
    "NUM.DE.SEMANA": ["fecha; [tipo]", "El número de semana del año (tipo 2 = la semana empieza en lunes)."], "ISO.NUM.DE.SEMANA": ["fecha", "El número de semana según la norma ISO."],
    DIAS360: ["inicio; fin; [europeo]", "Días entre dos fechas con años de 360 días (contabilidad)."],
    "DIA.LAB": ["inicio; días; [festivos]", "La fecha después de tantos días hábiles."],
    "DIAS.LAB": ["inicio; fin; [festivos]", "Días hábiles (lunes a viernes) entre dos fechas."],
    ESNOD: ["valor", "VERDADERO si es el error #N/A."], ESERR: ["valor", "VERDADERO si es un error que no sea #N/A."], NOD: ["", "El error #N/A (dato no disponible)."],
    "SI.ND": ["valor; si_no_está", "Otro valor si el resultado es #N/A."], "ES.PAR": ["número", "VERDADERO si es par."], "ES.IMPAR": ["número", "VERDADERO si es impar."],
    ESLOGICO: ["valor", "VERDADERO si es VERDADERO o FALSO."], ESNOTEXTO: ["valor", "VERDADERO si no es texto."],
    CAMBIAR: ["expresión; valor1; resultado1; …; [si_ninguno]", "Compara con una lista de valores y devuelve el resultado del que coincida."],
    XO: ["lógico1; [lógico2]; …", "VERDADERO si se cumple una cantidad impar de condiciones."], N: ["valor", "El valor como número (el texto vale 0)."],
    T: ["valor", "El valor si es texto; si no, vacío."], TIPO: ["valor", "1 número, 2 texto, 4 lógico, 16 error."],
    TEXTOANTES: ["texto; separador; [cuál]", "Lo que está antes del separador."], TEXTODESPUES: ["texto; separador; [cuál]", "Lo que está después del separador."],
    LIMPIAR: ["texto", "Quita los caracteres que no se imprimen."], CODIGO: ["texto", "El número del primer carácter."], UNICODE: ["texto", "El número Unicode del primer carácter."],
    UNICAR: ["número", "El carácter Unicode de ese número."], MONEDA: ["número; [decimales]", "El número como texto con $ y comas."],
    DECIMAL: ["número; [decimales]; [sin_comas]", "El número como texto con decimales fijos."], "NUMERO.ROMANO": ["número", "El número en romano (hasta 3999)."],
    "MULTIPLO.SUPERIOR": ["número; múltiplo", "Redondea hacia arriba al múltiplo más cercano."], "MULTIPLO.INFERIOR": ["número; múltiplo", "Redondea hacia abajo al múltiplo más cercano."],
    "MULTIPLO.SUPERIOR.MAT": ["número; [múltiplo]; [modo]", "Hacia arriba al múltiplo (o al entero) más cercano."], "MULTIPLO.INFERIOR.MAT": ["número; [múltiplo]; [modo]", "Hacia abajo al múltiplo (o al entero) más cercano."],
    "REDOND.MULT": ["número; múltiplo", "Redondea al múltiplo más cercano."], "REDONDEA.PAR": ["número", "Al par siguiente (alejándose de cero)."],
    "REDONDEA.IMPAR": ["número", "Al impar siguiente (alejándose de cero)."], "M.C.D": ["número1; [número2]; …", "Máximo común divisor."], "M.C.M": ["número1; [número2]; …", "Mínimo común múltiplo."],
    FACT: ["número", "El factorial (5! = 120)."], COMBINAT: ["n; k", "De cuántas formas se eligen k de n (sin orden)."], PERMUTACIONES: ["n; k", "De cuántas formas se ordenan k de n."],
    COCIENTE: ["número; divisor", "La parte entera de una división."], "SUMA.CUADRADOS": ["número1; [número2]; …", "Suma los cuadrados."],
    ASENO: ["número", "El arco seno (en radianes)."], ACOS: ["número", "El arco coseno (en radianes)."], ATAN: ["número", "El arco tangente (en radianes)."],
    ATAN2: ["x; y", "El ángulo del punto (x, y), en radianes."], SENOH: ["número", "Seno hiperbólico."], COSH: ["número", "Coseno hiperbólico."], TANH: ["número", "Tangente hiperbólica."],
    PROMEDIOA: ["valor1; [valor2]; …", "Promedio contando el texto como 0 y VERDADERO como 1."], MAXA: ["valor1; [valor2]; …", "El mayor, contando texto y lógicos."],
    MINA: ["valor1; [valor2]; …", "El menor, contando texto y lógicos."],
    PERCENTIL: ["rango; k", "El valor bajo el cual queda la fracción k (0 a 1) de los datos."], CUARTIL: ["rango; cuartil", "El cuartil 0 a 4 (2 = la mediana)."],
    "MEDIA.GEOM": ["número1; [número2]; …", "La media geométrica."], "MEDIA.ARMO": ["número1; [número2]; …", "La media armónica."], DESVPROM: ["número1; [número2]; …", "Promedio de las distancias a la media."],
    PENDIENTE: ["conocido_y; conocido_x", "La pendiente de la recta que mejor ajusta."], "INTERSECCION.EJE": ["conocido_y; conocido_x", "Dónde corta esa recta al eje y."],
    PRONOSTICO: ["x; conocido_y; conocido_x", "El valor que predice la recta para x."], TENDENCIA: ["conocido_y; [conocido_x]; [nueva_x]", "El valor de la recta de tendencia en nueva_x."],
    "COEF.DE.CORREL": ["rango1; rango2", "La correlación entre dos listas (-1 a 1)."], "COEFICIENTE.R2": ["conocido_y; conocido_x", "Qué tan bien ajusta la recta (0 a 1)."],
    "COVARIANZA.P": ["rango1; rango2", "Covarianza (población)."], "COVARIANZA.M": ["rango1; rango2", "Covarianza (muestra)."],
    "JERARQUIA.MEDIA": ["número; referencia; [orden]", "El lugar en la lista; los empates comparten el promedio de sus lugares."],
    PAGO: ["tasa; periodos; préstamo; [valor_final]; [tipo]", "El pago fijo de un préstamo."], VA: ["tasa; periodos; pago; [valor_final]; [tipo]", "Cuánto vale hoy una serie de pagos."],
    VF: ["tasa; periodos; pago; [valor_actual]; [tipo]", "Cuánto se junta al final ahorrando un pago fijo."], NPER: ["tasa; pago; valor_actual; [valor_final]; [tipo]", "Cuántos pagos hacen falta."],
    TASA: ["periodos; pago; valor_actual; [valor_final]; [tipo]; [estimar]", "La tasa de interés por periodo."], VNA: ["tasa; valor1; [valor2]; …", "Valor neto actual de una serie de flujos."],
    TIR: ["flujos; [estimar]", "La tasa interna de retorno."], PAGOINT: ["tasa; periodo; periodos; préstamo; [valor_final]; [tipo]", "Lo que se paga de interés en un periodo."],
    PAGOPRIN: ["tasa; periodo; periodos; préstamo; [valor_final]; [tipo]", "Lo que se abona a capital en un periodo."], SLN: ["costo; valor_residual; vida", "Depreciación en línea recta por periodo."],
    "INT.EFECTIVO": ["tasa_nominal; periodos_al_año", "La tasa efectiva anual."], "TASA.NOMINAL": ["tasa_efectiva; periodos_al_año", "La tasa nominal anual."],
    FILA: ["[referencia]", "El número de fila (de esta celda si no dices otra)."], COLUMNA: ["[referencia]", "El número de columna."],
    FILAS: ["rango", "Cuántas filas tiene el rango."], COLUMNAS: ["rango", "Cuántas columnas tiene el rango."],
    DIRECCION: ["fila; columna; [tipo]", "El nombre de la celda como texto ($A$1)."], INDIRECTO: ["texto", "El valor de la celda cuyo nombre está escrito en el texto."],
    DESREF: ["referencia; filas; columnas; [alto]; [ancho]", "El valor que está a tantas filas y columnas de una celda."],
    // matrices dinámicas: el resultado se derrama en las celdas de la derecha y de abajo
    ORDENAR: ["matriz; [columna]; [orden 1 o -1]; [por_columnas]", "Ordena las filas (de menor a mayor con 1, al revés con -1)."],
    ORDENARPOR: ["matriz; por_matriz1; [orden1]; …", "Ordena las filas según otra columna."],
    FILTRAR: ["matriz; incluir; [si_vacío]", "Solo las filas que cumplen (incluir: VERDADERO o FALSO por fila)."],
    UNICOS: ["matriz; [por_columnas]; [solo_una_vez]", "Los valores sin repetir."], SECUENCIA: ["filas; [columnas]; [inicio]; [paso]", "Una lista de números seguidos."],
    MATRIZALEAT: ["[filas]; [columnas]; [mín]; [máx]; [entero]", "Números al azar."], TRANSPONER: ["matriz", "Cambia filas por columnas."],
    APILARV: ["matriz1; [matriz2]; …", "Pone las matrices una debajo de otra."], APILARH: ["matriz1; [matriz2]; …", "Pone las matrices una al lado de otra."],
    TOMAR: ["matriz; filas; [columnas]", "Las primeras (o con negativo, las últimas) filas y columnas."], EXCLUIR: ["matriz; filas; [columnas]", "Quita las primeras (o las últimas) filas y columnas."],
    ELEGIRCOLS: ["matriz; columna1; [columna2]; …", "Solo esas columnas."], ELEGIRFILAS: ["matriz; fila1; [fila2]; …", "Solo esas filas."],
    ENCOL: ["matriz; [ignorar]; [por_columna]", "Todos los valores en una columna."], ENFILA: ["matriz; [ignorar]; [por_columna]", "Todos los valores en una fila."],
    AJUSTARFILAS: ["vector; cuántos; [relleno]", "Parte una lista en filas de tantos valores."], AJUSTARCOLS: ["vector; cuántos; [relleno]", "Parte una lista en columnas de tantos valores."],
    EXPANDIR: ["matriz; filas; [columnas]; [relleno]", "Agranda la matriz y rellena lo nuevo."],
    DIVIDIRTEXTO: ["texto; separador_columnas; [separador_filas]; [ignorar_vacíos]; [mayúsculas]; [relleno]", "Parte un texto en celdas."],
    FRECUENCIA: ["datos; intervalos", "Cuántos datos caen en cada intervalo."], MMULT: ["matriz1; matriz2", "Producto de matrices."],
    "M.UNIDAD": ["tamaño", "La matriz identidad."], MDETERM: ["matriz", "El determinante."], MINVERSA: ["matriz", "La matriz inversa."],
    LET: ["nombre1; valor1; [nombre2; valor2]; …; cálculo", "Pone nombre a resultados para usarlos en el cálculo."],
    LAMBDA: ["[parámetro1]; …; cálculo", "Una función propia: =LAMBDA(x; x*2)(5)."],
    BYROW: ["matriz; LAMBDA(fila)", "Un resultado por cada fila."], BYCOL: ["matriz; LAMBDA(columna)", "Un resultado por cada columna."],
    MAP: ["matriz1; [matriz2]; …; LAMBDA", "Aplica la función a cada valor."], REDUCE: ["[inicial]; matriz; LAMBDA(acumulado; valor)", "Acumula todos los valores en uno."],
    SCAN: ["[inicial]; matriz; LAMBDA(acumulado; valor)", "Cada paso del acumulado."], MAKEARRAY: ["filas; columnas; LAMBDA(fila; columna)", "Arma una matriz con la función."],
    ISOMITTED: ["parámetro", "VERDADERO si al llamar la LAMBDA no se dio ese parámetro."],
  };

  function num(v) {
    if (v instanceof ErrorF) throw v;
    if (typeof v === "number") return v;
    if (v instanceof FechaF) return v.serie;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (v === "" || v == null) return 0;
    if (Array.isArray(v)) { if (v.length === 1) return num(v[0]); throw E("#VALOR!"); }
    const x = calcularLiteral(String(v).trim());
    if (esNum(x)) return +x;
    throw E("#VALOR!");
  }
  const texto = (v) => (v == null ? "" : typeof v === "number" || v instanceof FechaF ? formatear(v) : typeof v === "boolean" ? (v ? "VERDADERO" : "FALSO") : Array.isArray(v) ? texto(v[0]) : String(v));
  const verdad = (v) => (typeof v === "string" ? v !== "" && v.toUpperCase() !== "FALSO" && v.toUpperCase() !== "FALSE" : !!v);

  /** Lo que se escribió en una celda (sin =): número, porcentaje, moneda, fecha o texto. */
  function calcularLiteral(t) {
    if (/^-?\d+([.,]\d+)?%?$/.test(t) && !/^-?\d{1,3},\d{3}$/.test(t)) { const n = parseFloat(t.replace(",", ".")); return t.endsWith("%") ? n / 100 : n; }
    let m = /^(-?)\$?\s?(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?$/.exec(t);
    if (m && (t.includes("$") || t.includes(","))) return parseFloat(m[1] + m[2].replace(/,/g, "") + (m[3] || ""));
    m = /^(-?)(\d{1,3}(?:,\d{3})+)(\.\d+)?%$/.exec(t);
    if (m) return parseFloat(m[1] + m[2].replace(/,/g, "") + (m[3] || "")) / 100;
    const f = leerFecha(t);
    if (f) return f;
    return t;
  }

  /** Una función hecha con LAMBDA: sus parámetros y el cálculo (los símbolos de la fórmula, sin calcular). */
  class LambdaF { constructor(params, tk, ini, fin, ambito) { this.params = params; this.tk = tk; this.ini = ini; this.fin = fin; this.ambito = ambito; } toString() { return "#CALC!"; } }
  const OMITIDO = Object.freeze({ omitido: true });
  const MAX_DERRAME = 2000000;
  const ERR_ES = { "#NAME?": "#NOMBRE?", "#VALUE!": "#VALOR!", "#SPILL!": "#DESBORDAMIENTO!", "#N/D": "#N/A", "#NULL!": "#NULO!" };
  // Fórmulas que pueden dar una matriz (las demás nunca derraman): con un rango, una matriz escrita, un derrame,
  // la intersección o una función que arma matrices sin recibir un rango.
  const PUEDE_MATRIZ = /[:{#@]|\b(SECUENCIA|SEQUENCE|MATRIZALEAT|RANDARRAY|DIVIDIRTEXTO|TEXTSPLIT|DESREF|OFFSET|INDIRECTO|INDIRECT|MAKEARRAY|LAMBDA|LET|TRANSPONER|TRANSPOSE|APILARV|APILARH|VSTACK|HSTACK|EXPANDIR|EXPAND|M\.UNIDAD|MUNIT|AJUSTARFILAS|AJUSTARCOLS|WRAPROWS|WRAPCOLS|ENCOL|ENFILA|TOCOL|TOROW)\b/i;

  /**
   * Hoja: obtener(c, f) -> texto crudo de la celda. Devuelve un evaluador con caché.
   * opc: { nombre, filas, cols, formulas() -> [[c, f, crudo]…], sinDerrame, hojas(nombre) -> { obtener, filas, cols,
   * nombre, formulas, sinDerrame } | null } para las referencias a otras hojas, las columnas enteras y los derrames.
   * Una fórmula que da una matriz se derrama en las celdas de la derecha y de abajo, como en Excel 365: si alguna
   * tiene algo, la fórmula da #DESBORDAMIENTO!. Sin formulas() (o con sinDerrame) se ve solo el primer valor.
   */
  function evaluador(obtener, opc) {
    opc = opc || {};
    const nueva = (o, nombre) => ({ obtener: o.obtener, filas: o.filas || 0, cols: o.cols || 0, cache: new Map(), nombre,
      formulas: o.formulas || null, sinDerrame: !!o.sinDerrame, combinada: o.combinada || null, derrames: new Map(), meta: new Set(), duenos: new Map(), cand: null });
    const actual = nueva({ obtener, filas: opc.filas, cols: opc.cols, formulas: opc.formulas, sinDerrame: opc.sinDerrame, combinada: opc.combinada }, opc.nombre || "");
    const otras = new Map();
    const pila = new Set();
    function hojaDe(nombre, desde) {
      if (nombre == null) return desde;
      const k = normal(nombre);
      if (opc.nombre && normal(opc.nombre) === k) return actual;
      if (otras.has(k)) return otras.get(k);
      const h = opc.hojas ? opc.hojas(nombre) : null;
      const x = h ? nueva(h, h.nombre || nombre) : null;
      otras.set(k, x);
      return x;
    }
    const derrama = (h) => !!h.formulas && !h.sinDerrame;
    /** Las celdas con fórmula de una hoja que podrían derramar, de arriba a abajo. */
    function candidatos(h) {
      if (h.cand) return h.cand;
      h.cand = [];
      if (!derrama(h)) return h.cand;
      for (const [c, f, crudo] of h.formulas()) if (PUEDE_MATRIZ.test(String(crudo))) h.cand.push({ c, f, muerto: false });
      h.cand.sort((a, b) => a.f - b.f || a.c - b.c);
      return h.cand;
    }
    /** Lo que le toca a una celda vacía por el derrame de una fórmula de arriba o de la izquierda. */
    function derramadoEn(h, c, f) {
      let saltada = false;
      for (const a of candidatos(h)) {
        if (a.f > f) break;
        if (a.muerto || a.c > c || (a.f === f && a.c === c)) continue;
        if (pila.has(h.nombre + "|" + a.c + "," + a.f)) { saltada = true; continue; }
        valorEn(h, a.c, a.f);
        const d = h.derrames.get(a.c + "," + a.f);
        if (!d) { a.muerto = true; continue; }
        if (d.bloqueado) continue;
        if (c < a.c + d.cols && f < a.f + d.filas) return { v: d.arr[(f - a.f) * d.cols + (c - a.c)], a };
      }
      return saltada ? { v: "", sinCache: true } : null;
    }
    function valorEn(h, c, f) {
      const k = c + "," + f;
      if (h.cache.has(k)) return h.cache.get(k);
      const crudo = h.obtener(c, f);
      if (crudo === "" || crudo == null) {
        const d = derrama(h) ? derramadoEn(h, c, f) : null;
        let v = d ? d.v : "";
        if (v === OMITIDO || v instanceof LambdaF) v = E("#CALC!");
        if (d && d.a) h.duenos.set(k, d.a);
        if (!(d && d.sinCache)) h.cache.set(k, v);
        return v;
      }
      const kp = h.nombre + "|" + k;
      if (pila.has(kp)) throw E("#CICLO!");
      pila.add(kp);
      const antes = usoMatriz;
      usoMatriz = false;
      let v;
      try { v = calcular(crudo, h, { c, f }); } catch (e) { v = e instanceof ErrorF ? e : E("#VALOR!"); }
      const uso = usoMatriz;
      usoMatriz = antes;
      pila.delete(kp);
      if (Array.isArray(v)) v = derramar(h, c, f, v);
      if (uso) h.meta.add(k);
      h.cache.set(k, v);
      return v;
    }
    /** Una celda cuya fórmula dio una matriz: se registra su derrame y la celda muestra el primer valor. */
    function derramar(h, c, f, arr) {
      const filas = filasDe(arr), cols = colsDe(arr);
      let bloqueado = false;
      if (derrama(h)) {
        if (arr.length > MAX_DERRAME || f + filas > 1048576 || c + cols > 16384) bloqueado = true;
        else for (let r = 0; r < filas && !bloqueado; r++) for (let q = r ? 0 : 1; q < cols; q++) {
          const x = h.obtener(c + q, f + r);
          if ((x !== "" && x != null) || (h.combinada && h.combinada(c + q, f + r))) { bloqueado = true; break; }
        }
      }
      h.derrames.set(c + "," + f, { arr, filas, cols, bloqueado });
      const v0 = arr[0] === OMITIDO || arr[0] instanceof LambdaF ? E("#CALC!") : arr[0];
      return bloqueado ? E("#DESBORDAMIENTO!") : v0;
    }
    function rango(r, h) {
      const hh = hojaDe(r.hoja, h);
      if (!hh) throw E("#REF!");
      let c1, c2, f1, f2;
      if (r.tipo === "col") { c1 = Math.min(r.a.c, r.b.c); c2 = Math.max(r.a.c, r.b.c); f1 = 0; f2 = Math.max(1, hh.filas) - 1; }
      else if (r.tipo === "fila") { f1 = Math.min(r.a.f, r.b.f); f2 = Math.max(r.a.f, r.b.f); c1 = 0; c2 = Math.max(1, hh.cols) - 1; }
      else { c1 = Math.min(r.a.c, r.b.c); c2 = Math.max(r.a.c, r.b.c); f1 = Math.min(r.a.f, r.b.f); f2 = Math.max(r.a.f, r.b.f); }
      if ((f2 - f1 + 1) * (c2 - c1 + 1) > 200000) throw E("#NUM!");
      const out = [];
      out.cols = c2 - c1 + 1;   // forma del rango (BUSCARV, INDICE… la usan)
      out.f1 = f1; out.c1 = c1;  // y su lugar (para la intersección implícita, «@»)
      for (let f = f1; f <= f2; f++) for (let c = c1; c <= c2; c++) {
        const v = valorEn(hh, c, f);
        if (v instanceof ErrorF) throw v;
        out.push(v);
      }
      return out;
    }

    function calcular(crudo, h, donde) {
      if (crudo == null || crudo === "") return "";
      const s = String(crudo);
      // «'» al principio: es texto aunque parezca número o fórmula (como el apóstrofo de Excel)
      if (s[0] === "'") return s.slice(1);
      if (s[0] !== "=") return calcularLiteral(s.trim()) === s.trim() ? s : calcularLiteral(s.trim());
      let tk = tokens(s.slice(1));
      let i = 0;
      // nombres de LET y parámetros de LAMBDA (cada uno con el ámbito de afuera)
      let ambito = null;
      const ver = () => tk[i];
      const es = (v) => tk[i] && tk[i].k === "op" && tk[i].v === v;
      const toma = (v) => { const x = tk[i]; if (!x || (v && (x.k !== "op" || x.v !== v))) throw E("#VALOR!"); i++; return x; };
      const sep = () => es(",") || es(";");

      function expr() {
        let a = suma();
        while (ver() && ver().k === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(ver().v)) {
          const o = toma().v, b = suma();
          a = porElemento([a, b], (p, q) => {
            const x = esNum(p) && esNum(q) ? [+p, +q] : [texto(p).toLowerCase(), texto(q).toLowerCase()];
            return { "=": x[0] === x[1], "<>": x[0] !== x[1], "<": x[0] < x[1], ">": x[0] > x[1], "<=": x[0] <= x[1], ">=": x[0] >= x[1] }[o];
          });
        }
        return a;
      }
      function suma() {
        let a = term();
        while (ver() && ver().k === "op" && ["+", "-", "&"].includes(ver().v)) {
          const o = toma().v, b = term();
          if (o === "&") { a = porElemento([a, b], (p, q) => texto(p) + texto(q)); continue; }
          a = porElemento([a, b], (p, q) => {
            const r = o === "+" ? num(p) + num(q) : num(p) - num(q);
            // Fecha ± días = fecha; fecha - fecha = días.
            const fa = p instanceof FechaF, fb = q instanceof FechaF;
            return (fa && !fb) || (fb && !fa && o === "+") ? new FechaF(r) : r;
          });
        }
        return a;
      }
      function term() {
        let a = pot();
        while (ver() && ver().k === "op" && ["*", "/"].includes(ver().v)) {
          const o = toma().v, b = pot();
          a = porElemento([a, b], (p, q) => {
            if (o === "/") { if (num(q) === 0) throw E("#DIV/0!"); return num(p) / num(q); }
            return num(p) * num(q);
          });
        }
        return a;
      }
      function pot() {
        let a = unario();
        while (es("^")) { toma(); const b = unario(); a = porElemento([a, b], (p, q) => Math.pow(num(p), num(q))); }
        return a;
      }
      function unario() {
        if (es("-")) { toma(); return porElemento([unario()], (p) => -num(p)); }
        if (es("+")) { toma(); return porElemento([unario()], (p) => num(p)); }
        // «@»: la intersección implícita de Excel (la celda del rango que está en la fila o columna de la fórmula)
        if (es("@")) { toma(); return interseccion(unario()); }
        let v = primario();
        while (es("%")) { toma(); v = porElemento([v], (p) => num(p) / 100); }
        return v;
      }
      function primario() {
        const x = ver();
        if (!x) throw E("#VALOR!");
        if (x.k === "num") { i++; return x.v; }
        if (x.k === "str") { i++; return x.v; }
        if (x.k === "err") { i++; throw E(ERR_ES[x.v] || x.v); }
        if (x.k === "op" && x.v === "(") { toma(); const v = expr(); toma(")"); return llamadas(v); }
        if (x.k === "op" && x.v === "{") return matrizEscrita();
        if (x.k === "ref") {
          i++;
          const r = x.v;
          if (x.derrame) return derrameDe(r);
          if (r.tipo === "celda") {
            const hh = hojaDe(r.hoja, h);
            if (!hh) throw E("#REF!");
            const v = valorEn(hh, r.a.c, r.a.f);
            if (v instanceof ErrorF) throw v;
            return v;
          }
          return rango(r, h);
        }
        if (x.k === "id") {
          i++;
          if (x.v === "VERDADERO" || x.v === "TRUE" || x.v === "FALSO" || x.v === "FALSE") {
            if (es("(")) { toma("("); toma(")"); }   // también como función: =VERDADERO()
            return x.v === "VERDADERO" || x.v === "TRUE";
          }
          const propio = buscar(x.v);
          if (es("(")) {
            // una función guardada con LET o un parámetro: =LET(f; LAMBDA(x; x*2); f(5))
            if (propio instanceof LambdaF) return llamadas(propio);
            toma("(");
            const nombre = ALIAS[x.v] || x.v;
            if (nombre === "LET") return llamadas(conNombres());
            if (nombre === "LAMBDA") return llamadas(lambda());
            if (nombre === "ISOMITTED") {
              const t = ver(), sig = tk[i + 1];
              if (t && t.k === "id" && sig && sig.k === "op" && sig.v === ")") { i += 2; return buscar(t.v) === OMITIDO; }
              return argumentos()[0] === undefined;
            }
            const fn = FUNCS[nombre];
            if (!fn) throw E("#NOMBRE?");
            if (DE_REFERENCIA.has(nombre)) return deReferencia(nombre);
            const args = argumentos();
            if (DINAMICAS.has(nombre)) usoMatriz = true;
            const r = llamarFuncion(nombre, fn, args);
            if (r instanceof ErrorF) throw r;
            return llamadas(r);
          }
          if (propio !== undefined) {
            if (propio instanceof ErrorF) throw propio;
            return propio === OMITIDO ? "" : propio;
          }
          throw E("#NOMBRE?");
        }
        throw E("#VALOR!");
      }
      /** Los argumentos de una función (ya se tomó el «(»), hasta el «)». Uno vacío queda undefined. */
      function argumentos() {
        const args = [];
        if (ver() && !es(")")) {
          args.push(sep() ? undefined : argumento());
          while (sep()) { toma(); args.push(es(")") || sep() ? undefined : argumento()); }
        }
        toma(")");
        return args;
      }
      /** Una LAMBDA seguida de (argumentos) se calcula: =LAMBDA(n; n*n+1)(6). */
      function llamadas(v) {
        while (v instanceof LambdaF && es("(")) { toma("("); v = invocar(v, argumentos()); }
        return v;
      }
      function invocar(lam, args) {
        if (args.length > lam.params.length) throw E("#VALOR!");
        const m = new Map();
        lam.params.forEach((p, k) => m.set(p, args[k] === undefined ? OMITIDO : args[k]));
        const g = [tk, i, ambito];
        tk = lam.tk; i = lam.ini; ambito = { m, padre: lam.ambito };
        try {
          const v = expr();
          if (i !== lam.fin) throw E("#VALOR!");
          return v;
        } finally { [tk, i, ambito] = g; }
      }
      function buscar(n) { for (let a = ambito; a; a = a.padre) if (a.m.has(n)) return a.m.get(n); return undefined; }
      /** LET(nombre1; valor1; …; cálculo), ya tomado el «(». */
      function conNombres() {
        const previo = ambito;
        ambito = { m: new Map(), padre: previo };
        try {
          for (;;) {
            const t = ver(), sig = tk[i + 1];
            if (t && t.k === "id" && sig && sig.k === "op" && (sig.v === "," || sig.v === ";")) {
              i += 2;
              const v = argumento();
              ambito.m.set(t.v, v);
              if (!sep()) throw E("#VALOR!");
              toma();
              continue;
            }
            if (!ambito.m.size) throw E("#VALOR!");
            const v = expr();
            toma(")");
            return v;
          }
        } finally { ambito = previo; }
      }
      /** LAMBDA(parámetro1; …; cálculo), ya tomado el «(»: el cálculo se guarda para hacerlo al llamarla. */
      function lambda() {
        const partes = [];
        let ini = i, prof = 0, j = i;
        for (; j < tk.length; j++) {
          const t = tk[j];
          if (t.k !== "op") continue;
          if (t.v === "(" || t.v === "{") prof++;
          else if (t.v === ")" || t.v === "}") { if (!prof) break; prof--; }
          else if ((t.v === "," || t.v === ";") && !prof) { partes.push([ini, j]); ini = j + 1; }
        }
        if (j >= tk.length) throw E("#VALOR!");
        partes.push([ini, j]);
        const [a, b] = partes.pop();
        if (b <= a) throw E("#VALOR!");
        // (un parámetro entre corchetes, [y], se puede omitir al llamarla: ISOMITTED dice si se dio)
        const params = partes.map(([p, q]) => {
          if (q - p === 3 && tk[p].k === "op" && tk[p].v === "[" && tk[p + 1].k === "id" && tk[p + 2].k === "op" && tk[p + 2].v === "]") return tk[p + 1].v;
          if (q - p !== 1 || tk[p].k !== "id") throw E("#VALOR!");
          return tk[p].v;
        });
        i = j + 1;
        usoMatriz = true;
        const lam = new LambdaF(params, tk, a, b, ambito);
        lam.llamar = (args) => invocar(lam, args);
        return lam;
      }
      /** {1,2,3;4,5,6}: «,» (o «\») separa columnas y «;» filas. */
      function matrizEscrita() {
        toma("{");
        const filas = [[]];
        if (!es("}")) for (;;) {
          filas[filas.length - 1].push(intenta(() => expr()));
          if (es(",") || es("\\")) { toma(); continue; }
          if (es(";")) { toma(); filas.push([]); continue; }
          break;
        }
        toma("}");
        const cols = filas[0].length;
        if (!cols || filas.some((q) => q.length !== cols)) throw E("#VALOR!");
        const out = [].concat(...filas);
        out.cols = cols;
        return out;
      }
      /** A1#: toda la matriz que derrama la celda A1. */
      function derrameDe(r) {
        const hh = hojaDe(r.hoja, h);
        if (!hh) throw E("#REF!");
        const v = valorEn(hh, r.a.c, r.a.f);
        const d = hh.derrames.get(r.a.c + "," + r.a.f);
        if (!d) {
          if (v instanceof ErrorF) throw v;
          if (String(hh.obtener(r.a.c, r.a.f))[0] === "=") return v;
          throw E("#REF!");
        }
        if (d.bloqueado) throw E("#DESBORDAMIENTO!");
        return d.arr;
      }
      function interseccion(v) {
        if (!Array.isArray(v)) return v;
        if (v.length === 1) return v[0];
        if (v.f1 != null && donde) {
          const n = filasDe(v), k = colsDe(v);
          const fi = n === 1 ? 0 : donde.f - v.f1, ci = k === 1 ? 0 : donde.c - v.c1;
          if (fi >= 0 && fi < n && ci >= 0 && ci < k) return v[fi * k + ci];
          throw E("#VALOR!");
        }
        return v[0];
      }
      /** El valor de una referencia ya armada (una celda o un rango). */
      function valorDeRef(r) {
        if (r.tipo === "celda") { const hh = hojaDe(r.hoja, h); if (!hh) throw E("#REF!"); const v = valorEn(hh, r.a.c, r.a.f); if (v instanceof ErrorF) throw v; return v; }
        return rango(r, h);
      }
      /** Filas y columnas que abarca una referencia: { f1, f2, c1, c2 }. */
      function cajaDeRef(r) {
        const hh = hojaDe(r.hoja, h);
        if (!hh) throw E("#REF!");
        if (r.tipo === "col") return { c1: Math.min(r.a.c, r.b.c), c2: Math.max(r.a.c, r.b.c), f1: 0, f2: Math.max(1, hh.filas) - 1 };
        if (r.tipo === "fila") return { f1: Math.min(r.a.f, r.b.f), f2: Math.max(r.a.f, r.b.f), c1: 0, c2: Math.max(1, hh.cols) - 1 };
        return { c1: Math.min(r.a.c, r.b.c), c2: Math.max(r.a.c, r.b.c), f1: Math.min(r.a.f, r.b.f), f2: Math.max(r.a.f, r.b.f) };
      }
      /** FILA, COLUMNA, FILAS, COLUMNAS, INDIRECTO y DESREF: el primer argumento es una referencia, no su valor. */
      function deReferencia(nombre) {
        // ¿el argumento que sigue es una referencia sola? (A1, B2:C5, Hoja2!A1)
        const refSola = () => { const t = tk[i], sig = tk[i + 1]; if (t && t.k === "ref" && !t.derrame && sig && sig.k === "op" && (sig.v === ")" || sig.v === "," || sig.v === ";")) { i++; return t.v; } return null; };
        const resto = () => { const l = []; while (sep()) { toma(); l.push(es(")") || sep() ? undefined : expr()); } toma(")"); return l; };
        if (nombre === "FILA" || nombre === "COLUMNA") {
          if (es(")")) { toma(")"); if (!donde) throw E("#VALOR!"); return (nombre === "FILA" ? donde.f : donde.c) + 1; }
          const r = refSola();
          if (!r) throw E("#VALOR!");
          toma(")");
          const c = cajaDeRef(r);
          return (nombre === "FILA" ? c.f1 : c.c1) + 1;
        }
        if (nombre === "FILAS" || nombre === "COLUMNAS") {
          const r = refSola();
          if (r) { toma(")"); const c = cajaDeRef(r); return nombre === "FILAS" ? c.f2 - c.f1 + 1 : c.c2 - c.c1 + 1; }
          const v = expr(); toma(")");
          if (!Array.isArray(v)) return 1;
          return nombre === "FILAS" ? filasDe(v) : colsDe(v);
        }
        if (nombre === "INDIRECTO") {
          const t = texto(expr()).trim(); resto();
          const r = leerRef(t, 0);
          if (!r || r.fin !== t.length) throw E("#REF!");
          return valorDeRef(r.ref);
        }
        // DESREF(referencia; filas; columnas; [alto]; [ancho])
        const r0 = refSola();
        if (!r0) throw E("#VALOR!");
        const a = resto(), c = cajaDeRef(r0);
        const df = Math.trunc(num(a[0] === undefined ? 0 : a[0])), dc = Math.trunc(num(a[1] === undefined ? 0 : a[1]));
        const alto = a[2] === undefined ? c.f2 - c.f1 + 1 : Math.trunc(num(a[2])), ancho = a[3] === undefined ? c.c2 - c.c1 + 1 : Math.trunc(num(a[3]));
        const f1 = c.f1 + df, c1 = c.c1 + dc;
        if (alto < 1 || ancho < 1 || f1 < 0 || c1 < 0) throw E("#REF!");
        const A = { c: c1, f: f1 }, B = { c: c1 + ancho - 1, f: f1 + alto - 1 };
        return valorDeRef(alto === 1 && ancho === 1 ? { tipo: "celda", a: A, b: A, hoja: r0.hoja } : { tipo: "rango", a: A, b: B, hoja: r0.hoja });
      }
      // Un error dentro de un argumento se guarda como valor (así SI y SIERROR pueden ignorarlo)
      // y se saltan los tokens que quedaron de ese argumento.
      function argumento() {
        const inicio = i;
        try { return expr(); } catch (e) {
          if (!(e instanceof ErrorF)) throw e;
          let prof = 0;
          if (i < inicio) i = inicio;
          while (i < tk.length) {
            const x = tk[i];
            if (x.k === "op" && (x.v === "(" || x.v === "{")) prof++;
            else if (x.k === "op" && (x.v === ")" || x.v === "}")) { if (prof === 0) break; prof--; }
            else if (x.k === "op" && (x.v === "," || x.v === ";") && prof === 0) break;
            i++;
          }
          return e;
        }
      }
      const v = expr();
      if (i < tk.length) throw E("#VALOR!");
      if (v instanceof LambdaF) throw E("#CALC!");
      if (v === OMITIDO) return "";
      if (Array.isArray(v)) { if (v.length === 1) return v[0] instanceof LambdaF ? E("#CALC!") : v[0]; if (!v.length) throw E("#CALC!"); return v; }
      return v;
    }

    const abierta = (c, f) => { valorEn(actual, c, f); return actual.derrames.get(c + "," + f) || null; };
    return {
      valor: (c, f) => valorEn(actual, c, f),
      limpiar: () => { actual.cache.clear(); actual.derrames.clear(); actual.meta.clear(); actual.duenos.clear(); actual.cand = null; otras.clear(); },
      // una fórmula que no está en ninguna celda, como si estuviera en (c, f): la de una regla de formato condicional
      calcular: (crudo, c, f) => { try { const v = calcular(crudo, actual, { c, f }); return Array.isArray(v) ? v[0] : v; } catch (e) { return e instanceof ErrorF ? e : E("#VALOR!"); } },
      /** Si la fórmula de (c, f) da una matriz: { filas, cols, bloqueado } (lo que derrama, o no cabe). */
      derrame: (c, f) => { const d = abierta(c, f); return d ? { filas: d.filas, cols: d.cols, bloqueado: d.bloqueado } : null; },
      /** ¿La fórmula de (c, f) es de matrices? (dio una matriz o calculó con ellas: Excel la guarda como fórmula de matriz) */
      esMatriz: (c, f) => !!abierta(c, f) || actual.meta.has(c + "," + f),
      /** La celda cuya fórmula derrama sobre (c, f), o null: { c, f }. */
      origen: (c, f) => { valorEn(actual, c, f); const a = actual.duenos.get(c + "," + f); return a ? { c: a.c, f: a.f } : null; },
    };
  }

  // ------------------------------------------------------------ mostrar valores
  function formatear(v) {
    if (v instanceof ErrorF) return v.codigo;
    if (v instanceof FechaF) return textoFecha(v);
    if (typeof v === "boolean") return v ? "VERDADERO" : "FALSO";
    if (typeof v === "number") {
      if (!isFinite(v)) return "#DIV/0!";
      const r = Math.round(v * 1e6) / 1e6;
      return Math.abs(r) >= 1e4 ? r.toLocaleString("es-MX", { maximumFractionDigits: 4 }) : String(r);
    }
    return String(v);
  }
  /** Separa un código de formato en secciones (positivo; negativo; cero) sin romper lo entre comillas. */
  function secciones(cod) {
    const r = []; let a = "", q = false;
    for (let i = 0; i < cod.length; i++) {
      const c = cod[i];
      if (c === '"') q = !q;
      if (c === "\\" && !q) { a += c + (cod[i + 1] || ""); i++; continue; }
      if (c === ";" && !q) { r.push(a); a = ""; continue; }
      a += c;
    }
    r.push(a);
    return r;
  }
  /** El código sin textos literales (para saber si es de fecha). */
  const sinLiterales = (s) => s.replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/\[[^\]]*\]/g, "");
  function esFormatoFecha(cod) {
    if (!cod || /^general$/i.test(cod) || cod === "@") return false;
    const s = sinLiterales(secciones(cod)[0]).toLowerCase();
    return /[dyh]|a{2,4}|s{1,2}|m{1,4}/.test(s) && !/[0#?]/.test(s);
  }
  function formatoFecha(serie, cod) {
    const d = fechaDe(serie), fr = serie - Math.floor(serie + 1e-9), seg = Math.round(fr * 86400);
    const hh = Math.floor(seg / 3600) % 24, mi = Math.floor(seg / 60) % 60, ss = seg % 60;
    const ampm = /am\/pm|a\/p/i.test(cod);
    const p2 = (n) => String(n).padStart(2, "0");
    const toks = [];
    const s = secciones(cod)[0];
    for (let i = 0; i < s.length;) {
      if (s[i] === '"') { const j = s.indexOf('"', i + 1); toks.push({ lit: s.slice(i + 1, j < 0 ? s.length : j) }); i = j < 0 ? s.length : j + 1; continue; }
      if (s[i] === "\\") { toks.push({ lit: s[i + 1] || "" }); i += 2; continue; }
      const m = /^(yyyy|yy|aaaa|aa|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p)/i.exec(s.slice(i));
      if (m) { toks.push({ t: m[0].toLowerCase() }); i += m[0].length; continue; }
      toks.push({ lit: s[i] }); i++;
    }
    return toks.map((x, k) => {
      if (x.lit != null) return x.lit;
      const prev = toks.slice(0, k).reverse().find((y) => y.t), sig = toks.slice(k + 1).find((y) => y.t);
      const minutos = (x.t === "m" || x.t === "mm") && ((prev && /^h/.test(prev.t)) || (sig && /^s/.test(sig.t)));
      switch (x.t) {
        case "yyyy": case "aaaa": return String(d.getUTCFullYear());
        case "yy": case "aa": return p2(d.getUTCFullYear() % 100);
        case "mmmm": return MESES[d.getUTCMonth()];
        case "mmm": return MESES_C[d.getUTCMonth()];
        case "mm": return minutos ? p2(mi) : p2(d.getUTCMonth() + 1);
        case "m": return minutos ? String(mi) : String(d.getUTCMonth() + 1);
        case "dddd": return DIAS_S[d.getUTCDay()];
        case "ddd": return DIAS_C[d.getUTCDay()];
        case "dd": return p2(d.getUTCDate());
        case "d": return String(d.getUTCDate());
        case "hh": return p2(ampm ? (hh % 12 || 12) : hh);
        case "h": return String(ampm ? (hh % 12 || 12) : hh);
        case "ss": return p2(ss);
        case "s": return String(ss);
        case "am/pm": return hh < 12 ? "a. m." : "p. m.";
        case "a/p": return hh < 12 ? "a" : "p";
      }
      return "";
    }).join("");
  }
  function formatoNumero(n, cod) {
    const secs = secciones(cod);
    let s = secs[0], neg = n < 0;
    if (n < 0 && secs[1] != null && secs[1] !== "") { s = secs[1]; neg = false; n = -n; }
    else if (n === 0 && secs[2] != null && secs[2] !== "") s = secs[2];
    else if (n < 0) n = -n;
    // Literales y marcadores.
    const partes = [];
    for (let i = 0; i < s.length;) {
      const c = s[i];
      if (c === '"') { const j = s.indexOf('"', i + 1); partes.push({ lit: s.slice(i + 1, j < 0 ? s.length : j) }); i = j < 0 ? s.length : j + 1; continue; }
      if (c === "\\") { partes.push({ lit: s[i + 1] || "" }); i += 2; continue; }
      if (c === "[") { const j = s.indexOf("]", i); i = j < 0 ? s.length : j + 1; continue; }   // [Rojo], [$-es-MX]…
      if (c === "_") { partes.push({ lit: " " }); i += 2; continue; }
      if (c === "*") { i += 2; continue; }
      partes.push(/[0#?.,%]/.test(c) ? { m: c } : { lit: c });
      i++;
    }
    const ini = partes.findIndex((p) => p.m && /[0#?]/.test(p.m));
    let fin = -1;
    for (let k = partes.length - 1; k >= 0; k--) if (partes[k].m && /[0#?]/.test(partes[k].m)) { fin = k; break; }
    const pct = partes.filter((p) => p.m === "%").length;
    if (ini < 0) {
      // Solo texto (por ejemplo "-"): se muestra tal cual.
      return (neg ? "-" : "") + partes.map((p) => (p.lit != null ? p.lit : p.m === "%" ? "%" : "")).join("");
    }
    const cuerpo = partes.slice(ini, fin + 1).map((p) => p.m || "").join("");
    const [ent, dec = ""] = cuerpo.split(".");
    const miles = ent.includes(",");
    const minEnt = (ent.match(/0/g) || []).length, decObl = (dec.match(/0/g) || []).length, decTot = (dec.match(/[0#?]/g) || []).length;
    let x = n * Math.pow(100, pct);
    let txt = x.toFixed(decTot);
    if (decTot > decObl) txt = txt.replace(new RegExp("0{1," + (decTot - decObl) + "}$"), "").replace(/\.$/, "");
    let [e, d] = txt.split(".");
    if (e.length < minEnt) e = e.padStart(minEnt, "0");
    if (minEnt === 0 && e === "0" && (d || decTot)) e = "";
    if (miles) e = e.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    const numero = e + (d ? "." + d : "");
    const pre = partes.slice(0, ini).map((p) => (p.lit != null ? p.lit : p.m === "%" ? "%" : "")).join("");
    const pos = partes.slice(fin + 1).map((p) => (p.lit != null ? p.lit : p.m === "%" ? "%" : "")).join("");
    const cero = /^[0.,]*$/.test(numero.replace(/,/g, ""));
    return (neg && !cero ? "-" : "") + pre + numero + pos;
  }
  // ------------------------------------------------------------ formato de número exactamente como Excel
  // Lo que muestra Excel (en español de México) para un valor y un código de formato: secciones
  // (positivo;negativo;cero;texto), condiciones [>=1000], colores [Red] y [ColorN], escala con comas,
  // porcentaje, notación científica, fracciones, literales, _x (espacio del ancho de x), *x (relleno),
  // @ (el texto), fechas del sistema 1900 de Excel (con su día 0 y su 29 de febrero de 1900) y horas
  // transcurridas [h]. Devuelve { t, color, relleno } (relleno: la letra que llena el ancho de la celda y
  // en qué lugar de t va, marcado con \u0000).
  const COLOR_XL = { black: "#000000", negro: "#000000", blue: "#0000FF", azul: "#0000FF", cyan: "#00FFFF", aguamarina: "#00FFFF",
    green: "#00FF00", verde: "#00FF00", magenta: "#FF00FF", red: "#FF0000", rojo: "#FF0000", white: "#FFFFFF", blanco: "#FFFFFF", yellow: "#FFFF00", amarillo: "#FFFF00" };
  const PALETA_XL = "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333".split(" ");
  /** Redondea a d decimales como Excel (mitad hacia afuera, sobre los 15 dígitos significativos). */
  function redondeaXL(x, d) {
    if (!isFinite(x)) return x;
    const n = Number(Math.abs(x).toPrecision(15));
    const r = Number(Math.round(Number(n + "e" + d)) + "e-" + d);
    return x < 0 ? -r : r;
  }
  function fijoXL(x, d) {
    const r = redondeaXL(Math.abs(x), d);
    let t = r.toFixed(Math.min(100, d));
    if (/e/i.test(t)) t = BigInt(Math.round(r)).toString();
    return t;
  }
  /** Separa una sección en piezas y saca su color, su condición y su moneda. */
  function piezasXL(sec) {
    const p = []; let color = null, cond = null;
    for (let i = 0; i < sec.length;) {
      const c = sec[i], resto = sec.slice(i);
      if (c === '"') { const j = sec.indexOf('"', i + 1); p.push({ t: "lit", v: sec.slice(i + 1, j < 0 ? sec.length : j) }); i = j < 0 ? sec.length : j + 1; continue; }
      if (c === "\\") { p.push({ t: "lit", v: sec[i + 1] || "" }); i += 2; continue; }
      if (c === "_") { p.push({ t: "lit", v: " " }); i += 2; continue; }
      if (c === "*") { p.push({ t: "relleno", v: sec[i + 1] || " " }); i += 2; continue; }
      if (c === "[") {
        const j = sec.indexOf("]", i), dentro = sec.slice(i + 1, j < 0 ? sec.length : j); i = j < 0 ? sec.length : j + 1;
        let m;
        if (COLOR_XL[dentro.toLowerCase()]) color = COLOR_XL[dentro.toLowerCase()];
        else if ((m = /^(?:color|color)\s*(\d+)$/i.exec(dentro))) color = "#" + (PALETA_XL[+m[1] - 1] || "000000");
        else if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:E[+-]?\d+)?)$/i.exec(dentro))) cond = { op: m[1], n: +m[2] };
        else if ((m = /^\$([^-\]]*)(?:-[0-9A-Za-z-]+)?$/.exec(dentro))) { if (m[1]) p.push({ t: "lit", v: m[1] }); }
        else if (/^(h+|m+|s+)$/i.test(dentro)) p.push({ t: "trans", v: dentro.toLowerCase() });
        continue;
      }
      let m = /^(yyyy|yyy|yy|y|aaaa|aa|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p|e)/i.exec(resto);
      if (m && !/^e[+-]/i.test(resto)) {
        const v = m[0].toLowerCase();
        p.push({ t: /^(am\/pm|a\/p)$/.test(v) ? "ampm" : "fecha", v, orig: m[0] }); i += m[0].length; continue;
      }
      if ((m = /^[eE]([+-])/.exec(resto))) { p.push({ t: "E", v: m[1] }); i += 2; continue; }
      if ("0#?".includes(c)) { p.push({ t: "dig", v: c }); i++; continue; }
      if (c === ".") { p.push({ t: "punto" }); i++; continue; }
      if (c === ",") { p.push({ t: "coma" }); i++; continue; }
      if (c === "%") { p.push({ t: "pct" }); i++; continue; }
      if (c === "/") { p.push({ t: "barra" }); i++; continue; }
      if (c === "@") { p.push({ t: "arroba" }); i++; continue; }
      p.push({ t: "lit", v: c }); i++;
    }
    return { p, color, cond };
  }
  const esSeccionFecha = (p) => p.some((x) => x.t === "fecha" || x.t === "ampm" || x.t === "trans") && !p.some((x) => x.t === "E");
  /** Partes de la fecha de un número de serie de Excel (sistema 1900, con sus rarezas). */
  function partesSerie(serie) {
    const ent = Math.floor(serie + 1e-9);
    let a, m, d, sem;
    if (ent === 0) { a = 1900; m = 1; d = 0; sem = 6; }
    else if (ent === 60) { a = 1900; m = 2; d = 29; sem = 3; }
    else if (ent < 60) { const f = new Date(Date.UTC(1900, 0, ent)); a = 1900; m = f.getUTCMonth() + 1; d = f.getUTCDate(); sem = (ent - 1) % 7; }
    else { const f = fechaDe(ent); a = f.getUTCFullYear(); m = f.getUTCMonth() + 1; d = f.getUTCDate(); sem = f.getUTCDay(); }
    return { a, m, d, sem };
  }
  function fechaXL(v, p) {
    if (v < 0 || v > 2958465.99999) return { t: "#".repeat(12), gato: true };
    const conSeg = p.some((x) => x.t === "fecha" && /^s/.test(x.v));
    // décimas de segundo: "ss.0" (el punto y los ceros tras los segundos)
    let decSeg = 0;
    for (let k = 0; k < p.length; k++) if (p[k].t === "punto" && p.slice(0, k).some((x) => x.t === "fecha" && /^s/.test(x.v))) { let j = k + 1; while (p[j] && p[j].t === "dig" && p[j].v === "0") { decSeg++; j++; } }
    // Excel redondea al segundo (o a la décima que se muestra) y luego corta: 10:57:38.9 con «h:mm» es 10:57
    const unidad = decSeg ? 86400 * Math.pow(10, decSeg) : 86400;
    const total = Math.round(v * unidad);   // redondeo al último lugar que se muestra (como Excel)
    const dias = Math.floor(total / unidad), resto = total - dias * unidad;
    const segT = unidad >= 86400 ? resto / (unidad / 86400) : resto * 60;
    const hh = Math.floor(segT / 3600), mi = Math.floor((segT % 3600) / 60), ss = Math.floor(segT % 60), fr = segT - Math.floor(segT);
    const F = partesSerie(dias);
    const ampm = p.some((x) => x.t === "ampm");
    const p2 = (n) => String(n).padStart(2, "0");
    let out = "", relleno = null;
    for (let k = 0; k < p.length; k++) {
      const x = p[k];
      if (x.t === "lit") { out += x.v; continue; }
      if (x.t === "relleno") { relleno = x.v; out += "\u0000"; continue; }
      if (x.t === "trans") {
        const tot = v * 86400;
        out += x.v[0] === "h" ? String(Math.floor((total / unidad) * 24 + 1e-9)).padStart(x.v.length, "0") : x.v[0] === "m" ? String(Math.floor((total / unidad) * 1440 + 1e-9)).padStart(x.v.length, "0") : String(Math.round(tot)).padStart(x.v.length, "0");
        continue;
      }
      if (x.t === "ampm") { out += x.v === "a/p" ? (hh < 12 ? "a" : "p") : hh < 12 ? "a. m." : "p. m."; continue; }
      if (x.t === "punto" && decSeg && p.slice(0, k).some((y) => y.t === "fecha" && /^s/.test(y.v))) { out += "." + String(Math.round(fr * Math.pow(10, decSeg))).padStart(decSeg, "0").slice(0, decSeg); k += decSeg; continue; }
      if (x.t === "fecha") {
        const ant = p.slice(0, k).reverse().find((y) => y.t === "fecha" || y.t === "trans"), sig = p.slice(k + 1).find((y) => y.t === "fecha" || y.t === "trans");
        const esMin = /^m{1,2}$/.test(x.v) && ((ant && /^h|^\[?h/.test(ant.v)) || (sig && /^s/.test(sig.v)) || (ant && ant.t === "trans" && ant.v[0] === "h"));
        switch (x.v) {
          case "yyyy": case "yyy": case "aaaa": case "e": out += String(F.a); break;
          case "yy": case "y": case "aa": out += p2(F.a % 100); break;
          case "mmmmm": out += MESES[F.m - 1][0]; break;
          case "mmmm": out += MESES[F.m - 1]; break;
          case "mmm": out += MESES_C[F.m - 1]; break;
          case "mm": out += esMin ? p2(mi) : p2(F.m); break;
          case "m": out += esMin ? String(mi) : String(F.m); break;
          case "dddd": out += DIAS_S[F.sem]; break;
          case "ddd": out += DIAS_C[F.sem]; break;
          case "dd": out += p2(F.d); break;
          case "d": out += String(F.d); break;
          case "hh": out += p2(ampm ? hh % 12 || 12 : hh); break;
          case "h": out += String(ampm ? hh % 12 || 12 : hh); break;
          case "ss": out += p2(ss); break;
          case "s": out += String(ss); break;
          default: out += x.orig;
        }
        continue;
      }
      if (x.t === "punto") { out += "."; continue; }
      if (x.t === "coma") { out += ","; continue; }
      if (x.t === "dig") { out += x.v === "0" ? "0" : ""; continue; }
      if (x.t === "pct") { out += "%"; continue; }
      if (x.t === "barra") { out += "/"; continue; }
    }
    return { t: out, relleno };
  }
  /** Llena los lugares de dígitos de la parte entera, de derecha a izquierda (con literales intercalados). */
  function enteroXL(digitos, lugares, miles) {
    // lugares: [{t:'dig',v}, {t:'lit',v}, ...] de la parte entera
    const solo = lugares.filter((x) => x.t === "dig");
    if (miles && solo.length === lugares.length) {
      const minimo = solo.filter((x) => x.v === "0").length, conEsp = solo.filter((x) => x.v === "?").length;
      let d = digitos === "0" && !solo.some((x) => x.v === "0") ? "" : digitos;
      if (d.length < minimo) d = d.padStart(minimo, "0");
      d = d.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      const faltan = Math.max(0, minimo + conEsp - digitos.length - (minimo ? 0 : 0));
      return (d.length < minimo + conEsp ? " ".repeat(Math.max(0, minimo + conEsp - d.replace(/,/g, "").length)) : "") + d;
    }
    const salida = [];
    let i = digitos.length - 1;
    if (digitos === "0") i = -1;   // el cero se escribe solo donde hay "0"
    for (let k = lugares.length - 1; k >= 0; k--) {
      const x = lugares[k];
      if (x.t !== "dig") { salida.unshift(x.v); continue; }
      const esPrimero = !lugares.slice(0, k).some((y) => y.t === "dig");
      if (i >= 0) {
        if (esPrimero) { salida.unshift(digitos.slice(0, i + 1)); i = -1; }
        else salida.unshift(digitos[i--]);
      } else salida.unshift(x.v === "0" ? "0" : x.v === "?" ? " " : "");
    }
    let r = salida.join("");
    if (miles) r = r.replace(/\d+/, (q) => q.replace(/\B(?=(\d{3})+(?!\d))/g, ","));
    return r;
  }
  function numeroXL(v, p) {
    let x = v;
    const pct = p.filter((y) => y.t === "pct").length;
    if (pct) x *= Math.pow(100, pct);
    // escala: comas pegadas después del último dígito de la parte entera (o al final del número)
    const idxDig = p.map((y, k) => (y.t === "dig" ? k : -1)).filter((k) => k >= 0);
    const iPunto = p.findIndex((y) => y.t === "punto"), iE = p.findIndex((y) => y.t === "E"), iBarra = p.findIndex((y) => y.t === "barra");
    const finEnt = (() => { const lim = [iPunto, iE, iBarra].filter((k) => k >= 0); const l = lim.length ? Math.min(...lim) : p.length; let u = -1; for (const k of idxDig) if (k < l) u = k; return u; })();
    let escala = 0;
    for (let k = finEnt + 1; k < p.length && p[k].t === "coma"; k++) escala++;
    // también las comas pegadas al último dígito del número (decimal o no): «0.0,"K"»
    const ultDig = idxDig.length ? idxDig[idxDig.length - 1] : -1;
    const comasFin = new Set();
    if (ultDig > finEnt) for (let k = ultDig + 1; k < p.length && p[k].t === "coma"; k++) { escala++; comasFin.add(k); }
    if (escala) x /= Math.pow(1000, escala);
    const miles = p.some((y, k) => y.t === "coma" && k < finEnt && p.slice(0, k).some((z) => z.t === "dig"));
    const comaEscala = new Set(); { let k = finEnt + 1; while (k < p.length && p[k].t === "coma") comaEscala.add(k++); }
    const lit = (y) => (y.t === "lit" ? y.v : y.t === "pct" ? "%" : y.t === "relleno" ? "\u0000" : "");
    let relleno = (p.find((y) => y.t === "relleno") || {}).v || null;
    // ---- fracción
    if (iBarra >= 0) {
      const den = []; let k = iBarra + 1;
      while (k < p.length && (p[k].t === "dig" || (p[k].t === "lit" && /\d/.test(p[k].v)))) den.push(p[k++]);
      const fijo = den.length && den.every((y) => y.t === "lit") ? +den.map((y) => y.v).join("") : null;
      const numL = []; let j = iBarra - 1;
      while (j >= 0 && p[j].t === "dig") numL.unshift(p[j--]);
      const antes = p.slice(0, j + 1), hayEntero = antes.some((y) => y.t === "dig");
      const a = Math.abs(x);
      let ent = hayEntero ? Math.floor(a) : 0, fr = a - ent, num, dn;
      if (fijo) { dn = fijo; num = Math.round(fr * dn); }
      else {
        // Excel usa las convergentes de la fracción continua (no la mejor aproximación): 0.4567 con «?/?» es 1/2
        const maxD = Math.pow(10, den.filter((y) => y.t === "dig").length) - 1;
        let h0 = 0, h1 = 1, k0 = 1, k1 = 0, r = fr;
        num = 0; dn = 1;
        for (let it = 0; it < 40; it++) {
          const a0 = Math.floor(r + 1e-12);
          const h2 = a0 * h1 + h0, k2 = a0 * k1 + k0;
          if (k2 > maxD) break;
          num = h2; dn = k2; h0 = h1; h1 = h2; k0 = k1; k1 = k2;
          const resto = r - a0; if (resto < 1e-10) break; r = 1 / resto;
        }
        if (dn === 0) { num = 0; dn = 1; }
      }
      if (num === dn && hayEntero) { ent += 1; num = 0; }
      const neg = "";
      const anchoN = Math.max(1, numL.length), anchoD = fijo ? String(fijo).length : den.filter((y) => y.t === "dig").length;
      let txtEnt = hayEntero ? enteroXL(String(ent), antes.filter((y) => y.t === "dig"), false) : "";
      const sepLit = hayEntero ? antes.slice(antes.map((y) => y.t).lastIndexOf("dig") + 1).map(lit).join("") : "";
      const preLit = antes.slice(0, antes.findIndex((y) => y.t === "dig")).map(lit).join("");
      let fracc;
      if (num === 0 && hayEntero) { fracc = " ".repeat(anchoN + 1 + anchoD); if (!txtEnt) txtEnt = "0"; }
      else fracc = String(num).padStart(anchoN, " ") + "/" + (fijo ? String(dn) : String(dn).padEnd(anchoD, " "));
      const despues = p.slice(k).map(lit).join("");
      return { t: neg + (hayEntero ? preLit + txtEnt + (num === 0 ? " ".repeat(sepLit.length) : sepLit) : antes.map(lit).join("")) + fracc + despues, relleno };
    }
    // ---- científica
    if (iE >= 0) {
      const mant = p.slice(0, iE), expL = p.slice(iE + 1).filter((y) => y.t === "dig");
      const entL = mant.filter((y, k) => y.t === "dig" && (iPunto < 0 || k < iPunto)), decL = iPunto >= 0 ? mant.filter((y, k) => y.t === "dig" && k > iPunto) : [];
      const a = Math.abs(x);
      let e = a === 0 ? 0 : Math.floor(Math.log10(a));
      const paso = entL.length > 1 && entL.some((y) => y.v === "#") ? entL.length : 1;
      if (paso > 1) e = Math.floor(e / paso) * paso; else e = e - (entL.length - 1);
      let m = a / Math.pow(10, e);
      let mt = fijoXL(m, decL.length);
      if (+mt >= Math.pow(10, Math.max(1, paso > 1 ? paso : entL.length))) { e += paso; m = a / Math.pow(10, e); mt = fijoXL(m, decL.length); }
      let [me, md = ""] = mt.split(".");
      const ex = String(Math.abs(e)).padStart(expL.length, "0");
      const signo = e < 0 ? "-" : p[iE].v === "+" ? "+" : "";
      const pre = mant.slice(0, mant.findIndex((y) => y.t === "dig")).map(lit).join("");
      return { t: pre + me + (decL.length ? "." + md : "") + "E" + signo + ex + p.slice(iE + 1).filter((y) => y.t !== "dig").map(lit).join(""), relleno };
    }
    // ---- número normal
    const decL = iPunto >= 0 ? p.filter((y, k) => y.t === "dig" && k > iPunto) : [];
    const nd = decL.length;
    const a = Math.abs(x);
    let txt = fijoXL(a, nd);
    let [e, d = ""] = txt.split(".");
    // decimales: # quita ceros finales, ? los vuelve espacios
    if (nd) {
      const chars = d.split("");
      for (let k = nd - 1; k >= 0; k--) { if (chars[k] !== "0") break; if (decL[k].v === "#") chars[k] = ""; else if (decL[k].v === "?") chars[k] = " "; else break; }
      d = chars.join("");
    }
    const primerDig = p.findIndex((y) => y.t === "dig");
    const ultEnt = iPunto >= 0 ? iPunto : (() => { let u = -1; p.forEach((y, k) => { if (y.t === "dig") u = k; }); return u + 1; })();
    const zonaEnt = primerDig >= 0 ? p.slice(primerDig, ultEnt).filter((y, k) => y.t !== "coma" && !comaEscala.has(primerDig + k)) : [];
    const txtEnt = primerDig < 0 || (iPunto >= 0 && primerDig > iPunto) ? "" : enteroXL(e, zonaEnt, miles);
    const pre = p.slice(0, Math.max(0, primerDig < 0 ? p.length : primerDig)).map(lit).join("");
    const post = p.slice(iPunto >= 0 ? iPunto + 1 : ultEnt).filter((y) => y.t !== "dig" && y.t !== "coma").map(lit).join("");
    void comasFin;
    const cuerpo = txtEnt + (iPunto >= 0 ? "." + d : "");
    return { t: pre + cuerpo + post, relleno, esCero: Number(txt) === 0 };
  }
  /** El formato General de Excel: hasta 11 caracteres; muy grande o muy chico, en notación científica. */
  function generalXL(n) {
    if (n === 0) return "0";
    const a = Math.abs(n), neg = n < 0 ? "-" : "";
    if (a >= 1e11 || a < 1e-9) {
      let [m, e] = a.toExponential(5).split("e");
      m = m.replace(/\.?0+$/, "");
      const ex = Math.abs(+e);
      return neg + m + "E" + (+e < 0 ? "-" : "+") + String(ex).padStart(2, "0");
    }
    for (let d = 10; d >= 0; d--) {
      const t = String(redondeaXL(a, d));
      const s = /e/i.test(t) ? a.toFixed(d) : t;
      if (s.replace("-", "").length <= 11) return neg + s;
    }
    return neg + String(Math.round(a));
  }
  /** Lo que muestra Excel para un valor con un formato: { t, color, relleno, gato (no cabe), num }. */
  function formatoXL(v, cod) {
    if (v instanceof ErrorF) return { t: v.codigo };
    if (v instanceof FechaF) v = v.serie;
    const codigo = cod == null || cod === "" ? "General" : String(cod);
    if (typeof v === "boolean") return { t: v ? "VERDADERO" : "FALSO" };
    const texto = typeof v === "string" && !esNum(v);
    const secs = secciones(codigo).map(piezasXL);
    if (texto || v === "" || v == null) {
      if (v === "" || v == null) return { t: "" };
      const s4 = secs[3] || (secs.length === 1 && secs[0].p.some((x) => x.t === "arroba") ? secs[0] : null);
      if (!s4) return { t: String(v) };
      let relleno = null;
      const t = s4.p.map((x) => (x.t === "arroba" ? String(v) : x.t === "lit" ? x.v : x.t === "relleno" ? ((relleno = x.v), "\u0000") : "")).join("");
      return { t, color: s4.color, relleno };
    }
    const n = +v;
    if (!isFinite(n)) return { t: "#¡DIV/0!" };
    if (/^general$/i.test(codigo.trim())) return { t: generalXL(n), num: true };
    // elegir sección
    let s, absoluto = false;
    const conCond = secs.some((x) => x.cond);
    if (conCond) {
      const cumple = (c) => !c || ({ "<": n < c.n, "<=": n <= c.n, ">": n > c.n, ">=": n >= c.n, "=": n === c.n, "<>": n !== c.n })[c.op];
      if (secs[0].cond && cumple(secs[0].cond)) s = secs[0];
      else if (secs[1] && secs[1].cond && cumple(secs[1].cond)) s = secs[1];
      else if (secs[1] && !secs[1].cond && (!secs[2] || secs[2].p.some((x) => x.t === "arroba"))) s = secs[1];
      else s = secs[2] || secs[1] || secs[0];
    } else if (n < 0 && secs.length >= 2 && secs[1].p.length) { s = secs[1]; absoluto = true; }
    else if (n === 0 && secs.length >= 3 && secs[2].p.length) s = secs[2];
    else s = secs[0];
    if (!s.p.length && secs.length > 1) return { t: "", color: s.color };   // ;;; (oculto)
    if (/^general$/i.test(s.p.map((x) => (x.t === "lit" ? x.v : x.orig || x.v || "")).join("").trim()) || (s.p.length && s.p.every((x) => x.t === "lit") && /^general$/i.test(s.p.map((x) => x.v).join("")))) {
      return { t: (absoluto ? generalXL(Math.abs(n)) : generalXL(n)), color: s.color, num: true };
    }
    if (esSeccionFecha(s.p)) { const r = fechaXL(n, s.p); return { t: r.t, color: s.color, relleno: r.relleno, gato: r.gato, num: true }; }
    if (s.p.some((x) => x.t === "arroba") && !s.p.some((x) => x.t === "dig")) return { t: s.p.map((x) => (x.t === "arroba" ? generalXL(n) : x.t === "lit" ? x.v : "")).join(""), color: s.color, num: true };
    // con condiciones, Excel da formato al valor sin signo y le antepone el menos (aun con literales)
    const sinSigno = conCond || absoluto;
    const r = numeroXL(Math.abs(n), s.p);
    let t = r.t;
    if (n < 0 && !absoluto && (conCond || !r.esCero)) t = "-" + t;
    void sinSigno;
    return { t, color: s.color, relleno: r.relleno, num: true };
  }
  /** Un valor con su formato de número (código como los de Excel). Sin formato: formatear(). */
  function aplicarFormato(v, cod) {
    if (NOM.formatoExcel !== false && !(v instanceof FechaF && (!cod || /^general$/i.test(cod)))) {
      try { return formatoXL(v, cod).t.replace(/\u0000/g, ""); } catch (e) { /* cae a lo de antes */ }
    }
    if (!cod || /^general$/i.test(cod)) return formatear(v);
    if (v instanceof ErrorF) return v.codigo;
    if (cod === "@") return texto(v);
    if (typeof v === "boolean" || v === "" || v == null) return formatear(v);
    if (!esNum(v)) return String(v);
    if (!isFinite(+v)) return "#DIV/0!";
    try { return esFormatoFecha(cod) ? formatoFecha(+v, cod) : formatoNumero(+v, cod); } catch (e) { return formatear(v); }
  }
  /** Formatos para elegir (el código es el de Excel, así se exporta igual). */
  const FORMATOS = [
    { id: "General", t: "General", ej: "1234.5" },
    { id: "0", t: "Número entero", ej: "1235" },
    { id: "#,##0.00", t: "Número con 2 decimales", ej: "1,234.50" },
    { id: "$#,##0.00", t: "Moneda", ej: "$1,234.50" },
    { id: "0%", t: "Porcentaje", ej: "50%" },
    { id: "0.00%", t: "Porcentaje con decimales", ej: "50.25%" },
    { id: "dd/mm/yyyy", t: "Fecha corta", ej: "07/11/2026" },
    { id: "d \"de\" mmmm \"de\" yyyy", t: "Fecha larga", ej: "7 de noviembre de 2026" },
    { id: "dddd d \"de\" mmmm", t: "Día y fecha", ej: "sábado 7 de noviembre" },
    { id: "hh:mm", t: "Hora", ej: "14:30" },
    { id: "@", t: "Texto", ej: "001234" },
  ];
  /** Un decimal más (+1) o menos (-1) en el formato de la celda. */
  function cambiarDecimales(cod, delta, valor) {
    if (!cod || /^general$/i.test(cod)) {
      const t = esNum(valor) ? String(Math.round(+valor * 1e6) / 1e6) : "";
      const actuales = t.includes(".") ? t.split(".")[1].length : 0;
      const n = Math.max(0, actuales + delta);
      return n ? "0." + "0".repeat(n) : "0";
    }
    if (esFormatoFecha(cod) || cod === "@") return cod;
    return secciones(cod).map((s) => {
      const m = /([0#?])(\.([0#?]*))?(?![0#?.,])/.exec(s);
      if (!m) return s;
      const idx = m.index;
      const decs = (m[3] || "").length + delta;
      const nuevo = m[1] + (decs > 0 ? "." + "0".repeat(decs) : "");
      return s.slice(0, idx) + nuevo + s.slice(idx + m[0].length);
    }).join(";");
  }
  /** Lo que alguien escribe en una celda: el valor a guardar y el formato que conviene ponerle. */
  function leerEntrada(t) {
    const s = String(t == null ? "" : t).trim();
    if (!s || s[0] === "=") return { crudo: t, formato: null };
    let m = /^-?\d+(?:([.,])(\d+))?%$/.exec(s);
    if (m) return { crudo: s, formato: m[2] ? "0." + "0".repeat(Math.min(4, m[2].length)) + "%" : "0%" };
    m = /^(-?)\$\s?(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d+))?$/.exec(s);
    if (m) return { crudo: m[1] + m[2].replace(/,/g, "") + (m[3] ? "." + m[3] : ""), formato: "$#,##0.00" };
    m = /^(-?)(\d{1,3}(?:,\d{3})+)(?:\.(\d+))?$/.exec(s);
    if (m) return { crudo: m[1] + m[2].replace(/,/g, "") + (m[3] ? "." + m[3] : ""), formato: m[3] ? "#,##0." + "0".repeat(Math.min(4, m[3].length)) : "#,##0" };
    return { crudo: t, formato: null };
  }

  // ------------------------------------------------------------ series (cuadrito de relleno)
  const LISTAS = [DIAS_S.slice(1).concat(DIAS_S[0]), DIAS_C.slice(1).concat(DIAS_C[0]), MESES, MESES_C,
    ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]];
  const quitarAcentos = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "");
  function enLista(t) {
    const b = quitarAcentos(t.toLowerCase());
    for (const l of LISTAS) { const i = l.findIndex((x) => quitarAcentos(x) === b); if (i >= 0) return { l, i }; }
    return null;
  }
  const comoLaOriginal = (modelo, t) => (modelo === modelo.toUpperCase() && modelo !== modelo.toLowerCase() ? t.toUpperCase() : modelo[0] === modelo[0].toUpperCase() ? t[0].toUpperCase() + t.slice(1) : t);
  /**
   * Continúa una serie como el cuadrito de relleno de Excel. crudos: los valores elegidos (texto
   * crudo, en orden); n: cuántos más. Las fórmulas las mueve quien llama (desplazar).
   * Un número solo se copia; dos o más números siguen el paso; fechas, días, meses y "Tema 1"
   * avanzan solos.
   */
  function serie(crudos, n) {
    const vals = crudos.map((c) => (c == null ? "" : String(c)));
    const out = [];
    const nums = vals.map((v) => { const x = v.trim() === "" ? null : calcularLiteral(v.trim()); return esNum(x) ? x : null; });
    if (vals.length && nums.every((x) => x != null)) {
      const fecha = nums.every((x) => x instanceof FechaF);
      if (nums.length === 1 && !fecha) { for (let k = 0; k < n; k++) out.push(vals[0]); return out; }
      const paso = nums.length === 1 ? 1 : (+nums[nums.length - 1] - +nums[0]) / (nums.length - 1);
      let ultimo = +nums[nums.length - 1];
      const pct = vals.every((v) => /%$/.test(v.trim()));
      for (let k = 0; k < n; k++) {
        ultimo += paso;
        if (fecha) { const d = fechaDe(ultimo); out.push(`${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`); }
        else if (pct) out.push(String(Math.round(ultimo * 100 * 1e6) / 1e6) + "%");
        else out.push(String(Math.round(ultimo * 1e9) / 1e9));
      }
      return out;
    }
    const listas = vals.map((v) => enLista(v.trim()));
    if (vals.length && listas.every((x) => x && x.l === listas[0].l)) {
      const l = listas[0].l, paso = listas.length > 1 ? (listas[listas.length - 1].i - listas[0].i) / (listas.length - 1) : 1;
      let i = listas[listas.length - 1].i;
      for (let k = 0; k < n; k++) { i = (((i + paso) % l.length) + l.length) % l.length; out.push(comoLaOriginal(vals[0].trim(), l[Math.round(i)])); }
      return out;
    }
    const conNumero = vals.map((v) => /^(.*?)(\d+)(\D*)$/.exec(v));
    if (vals.length && conNumero.every((m) => m && m[1] === conNumero[0][1] && m[3] === conNumero[0][3]) && conNumero[0][1].trim() !== "") {
      const ns = conNumero.map((m) => +m[2]);
      const paso = ns.length > 1 ? (ns[ns.length - 1] - ns[0]) / (ns.length - 1) : 1;
      let u = ns[ns.length - 1];
      for (let k = 0; k < n; k++) { u += paso; out.push(conNumero[0][1] + String(Math.round(u)).padStart(conNumero[0][2].length, "0") + conNumero[0][3]); }
      return out;
    }
    for (let k = 0; k < n; k++) out.push(vals[k % vals.length]);
    return out;
  }

  /**
   * Una fórmula que Excel guardó sin matrices dinámicas (sin t="array") se calcula como en un Excel viejo: un rango
   * que va a una operación, a un argumento de un solo valor o al resultado de la celda da solo la celda de su fila o
   * columna (la intersección implícita). Excel 365 la muestra con «@» ahí (=@A1:A10, =A:A*2 → =@A:A*2): se agrega
   * igual para que NOM-COLAB la calcule como Excel.
   */
  function conInterseccion(formula) {
    const s = String(formula || "");
    if (s[0] !== "=" || !/:/.test(s)) return s;
    let tk;
    try { tk = tokens(s.slice(1), true); } catch (e) { return s; }
    const OPS = new Set(["+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="]);
    // (estos argumentos ya se calculaban como matriz en el Excel viejo: SUMAPRODUCTO((A1:A9>5)*B1:B9) no lleva «@»)
    const DE_MATRIZ = { SUMAPRODUCTO: true, MMULT: true, FRECUENCIA: true, MDETERM: true, MINVERSA: true, TRANSPONER: true, INDICE: [0], AGREGAR: [2, 3] };
    const enMatriz = () => pila.some((c) => { const L = c.fn && DE_MATRIZ[c.fn]; return L === true || (Array.isArray(L) && L.includes(c.arg)); });
    const pila = [{ fn: null, arg: 0, ini: 0, arriba: true }], pos = [];
    for (let j = 0; j < tk.length; j++) {
      const t = tk[j];
      if (t.k === "op" && (t.v === "(" || t.v === "{")) { const p = tk[j - 1]; pila.push({ fn: t.v === "{" ? "{" : p && p.k === "id" ? ALIAS[p.v] || p.v : null, arg: 0, ini: j + 1 }); continue; }
      if (t.k === "op" && (t.v === ")" || t.v === "}")) { if (pila.length > 1) pila.pop(); continue; }
      if (t.k === "op" && (t.v === "," || t.v === ";")) { const c = pila[pila.length - 1]; c.arg++; c.ini = j + 1; continue; }
      if (t.k !== "ref" || t.v.tipo === "celda" || t.derrame) continue;
      const p = tk[j - 1], n = tk[j + 1];
      const antes = p && p.k === "op" && (OPS.has(p.v) || p.v === "@"), despues = n && n.k === "op" && (OPS.has(n.v) || n.v === "%");
      if ((p && p.k === "op" && p.v === "@") || enMatriz()) continue;
      let escalar = antes || despues;
      const ctx = pila[pila.length - 1], solo = ctx.ini === j && (!n || (n.k === "op" && [")", ",", ";", "}"].includes(n.v)));
      if (!escalar && solo) {
        if (ctx.arriba) escalar = !n;
        else if (ctx.fn && ctx.fn !== "{") { const L = ctx.fn === "SI" ? true : LEVANTA[ctx.fn]; escalar = L === true || (Array.isArray(L) && L.includes(ctx.arg)); }
      }
      if (escalar) pos.push(t.p + 1);
    }
    if (!pos.length) return s;
    let r = s;
    pos.sort((a, b) => b - a).forEach((q) => { r = r.slice(0, q) + "@" + r.slice(q); });
    return r;
  }
  /** Una fórmula de Excel en inglés (=AVERAGE(A1,B1)) en español (=PROMEDIO(A1;B1)). */
  function aEspanol(formula, conMatrices) {
    const s = String(formula || "");
    if (s[0] !== "=") return s;
    let r = "=", llaves = 0;
    for (let i = 1; i < s.length; i++) {
      const c = s[i];
      if (c === '"') { let j = i + 1; while (j < s.length) { if (s[j] === '"') { if (s[j + 1] === '"') { j += 2; continue; } break; } j++; } r += s.slice(i, j + 1); i = j; continue; }
      // (importador de Excel 5: dentro de una matriz escrita, {1,2;3,4}, la coma separa columnas y se queda)
      if (conMatrices && (c === "{" || c === "}")) { llaves += c === "{" ? 1 : -1; r += c; continue; }
      if (c === "," && llaves > 0) { r += c; continue; }
      if (c === ",") { r += ";"; continue; }
      const m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i));
      if (m && !/[A-Za-z0-9_.$!:]/.test(s[i - 1] || "")) {
        const n = m[0].toUpperCase(), sig = s.slice(i + m[0].length).trimStart()[0];
        if (sig === "(" && ALIAS[n] && /^[A-Z0-9.]+$/.test(n)) r += ALIAS[n];
        else if (sig !== "(" && sig !== "!" && (n === "TRUE" || n === "FALSE")) r += n === "TRUE" ? "VERDADERO" : "FALSO";
        else r += m[0];
        i += m[0].length - 1;
        continue;
      }
      r += c;
    }
    return r;
  }

  // ------------------------------------------------------------ hojas de un documento
  /** Las hojas (tablas de un documento de hoja de cálculo) en el orden de sus páginas. */
  function hojasDe(doc) {
    const pags = doc.paginas();
    return doc.vivos().filter((b) => b.tipo === "tabla" && !b.enTexto).sort((a, b) => pags.indexOf(a.pagina) - pags.indexOf(b.pagina) || (a.orden || 0) - (b.orden || 0));
  }
  /** Nombre de una hoja: el que le pusieron o "HojaN" según su lugar. */
  function nombreDeHoja(doc, b) {
    if (b && b.nombre) return String(b.nombre);
    const i = b ? hojasDe(doc).findIndex((x) => x.id === b.id) : -1;
    return "Hoja" + (i >= 0 ? i + 1 : 1);
  }
  const lectorDe = (b) => (c, f) => { const v = b["c:" + ref(c, f)]; return v == null ? "" : v; };
  /** Las celdas con fórmula de una tabla: [[c, f, crudo]…] (el evaluador busca ahí las que derraman). */
  const formulasDe = (b) => () => {
    const l = [];
    for (const k in b) if (k.charCodeAt(0) === 99 && k[1] === ":" && typeof b[k] === "string" && b[k][0] === "=") { const p = parseRef(k.slice(2)); if (p) l.push([p.c, p.f, b[k]]); }
    return l;
  };
  // Una hoja importada de Excel antes de la versión 5 del importador trae lo derramado como valores sueltos: ahí la
  // fórmula muestra su primer valor y no se derrama encima (si no, saldría #DESBORDAMIENTO!).
  const sinDerrameDe = (b) => !!(b.ox && +b.xlv > 0 && +b.xlv < 5);
  /** ¿(c, f) está dentro de una celda combinada? (tapa un derrame, como en Excel) */
  const combinadaDe = (b) => {
    let caja = null;
    return (c, f) => {
      if (!caja) { caja = []; for (const k in b) if (k.startsWith("m:")) { const p = parseRef(k.slice(2)), m = /^(\d+)x(\d+)$/.exec(String(b[k] || "")); if (p && m && +m[1] * +m[2] > 1) caja.push([p.c, p.f, p.c + +m[1] - 1, p.f + +m[2] - 1]); } }
      return caja.some((z) => c >= z[0] && c <= z[2] && f >= z[1] && f <= z[3]);
    };
  };
  /** Evaluador de una tabla que entiende las referencias a las demás hojas del documento. */
  function evaluadorDeTabla(doc, b) {
    const todas = doc ? hojasDe(doc) : [];
    return evaluador(lectorDe(b), {
      nombre: doc ? nombreDeHoja(doc, b) : "", filas: b.filas || b.filasTexto || 30, cols: b.cols || b.colsTexto || 10,
      formulas: formulasDe(b), sinDerrame: sinDerrameDe(b), combinada: combinadaDe(b),
      hojas: (n) => {
        const x = todas.find((h) => normal(nombreDeHoja(doc, h)) === normal(n));
        return x ? { obtener: lectorDe(x), filas: x.filas || 30, cols: x.cols || 10, nombre: nombreDeHoja(doc, x), formulas: formulasDe(x), sinDerrame: sinDerrameDe(x), combinada: combinadaDe(x) } : null;
      },
    });
  }

  NOM.formulas = {
    hojasDe, nombreDeHoja, evaluadorDeTabla,
    evaluador, formatear, aplicarFormato, formatoXL, ErrorF, FechaF, esNum, colLetra, colNum, ref, parseRef, NOMBRES, FIRMAS, ALIAS, INGLES,
    leerFecha, criterio, desplazar, ajustarEstructura, renombrarHoja, refsDe, textoRef, leerRef, transformarRefs, FORMATOS, esFormatoFecha,
    cambiarDecimales, leerEntrada, serie, calcularLiteral, prefijoHoja, textoFecha, aEspanol, formulasDe, conInterseccion,
  };
})();

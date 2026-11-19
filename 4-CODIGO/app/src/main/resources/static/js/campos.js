/* NOM · Campos de Word: fecha, número de página, títulos numerados (SEQ), referencias (REF, PAGEREF), condiciones (IF),
   fórmulas (=), propiedades del documento, combinar correspondencia…

   Como Word: cada campo guarda su código («SEQ Figura \* ARABIC») y su último valor. Al abrir se ve el último valor
   (Word tampoco recalcula al abrir, salvo la fecha y la hora, que se ven siempre al día); «Actualizar campos» (F9) los
   calcula todos otra vez, en el orden del documento, y guarda los valores nuevos (los ven todos los del documento).
   Lo que NOM no sabe calcular se queda con su último valor. Alt+F9 muestra los códigos.

   Formato de los resultados (los modificadores de Word): \@ fecha y hora («d 'de' MMMM 'de' yyyy»), \# número
   («#,##0.00»), \* Arabic, ROMAN, roman, ALPHABETIC, alphabetic, Ordinal, CardText, OrdText, Hex, DollarText, Upper,
   Lower, Caps, FirstCap (MERGEFORMAT y CHARFORMAT no cambian el texto). */
(function () {
  "use strict";
  const NOM = window.NOM;

  // ------------------------------------------------------------ leer el código
  /** «SEQ Figura \* ARABIC \r 3» → { nombre: "SEQ", args: ["Figura"], sw: { "\\*": ["ARABIC"], "\\r": "3" }, banderas: Set } */
  function leer(instr) {
    const s = String(instr || "").trim();
    const toks = [];
    let i = 0;
    while (i < s.length) {
      if (/\s/.test(s[i])) { i++; continue; }
      if (s[i] === '"') {
        let j = i + 1, t = "";
        while (j < s.length && s[j] !== '"') { if (s[j] === "\\" && s[j + 1] === '"') { t += '"'; j += 2; continue; } t += s[j++]; }
        toks.push({ t, q: true }); i = j + 1; continue;
      }
      let j = i, t = "";
      while (j < s.length && !/\s/.test(s[j])) t += s[j++];
      toks.push({ t, q: false }); i = j;
    }
    const r = { nombre: "", args: [], sw: {}, banderas: new Set(), crudo: s };
    if (!toks.length) return r;
    if (toks[0].t.startsWith("=")) { r.nombre = "="; const k = s.search(/\\[#*@]/); r.expr = (k >= 0 ? s.slice(1, k) : s.slice(1)).trim(); }
    else r.nombre = toks[0].t.toUpperCase();
    for (let k = 1; k < toks.length; k++) {
      const tk = toks[k];
      if (!tk.q && /^\\[a-z#*@!]$/i.test(tk.t)) {
        const sw = tk.t.toLowerCase();
        const siguiente = toks[k + 1];
        // los que llevan valor: \* \# \@ \r \s \l \b \c \p (según el campo); los demás son banderas
        if (["\\*", "\\#", "\\@"].includes(sw) || (r.nombre === "SEQ" && ["\\r", "\\s"].includes(sw)) || (r.nombre === "STYLEREF" && false)) {
          if (siguiente) { if (sw === "\\*") (r.sw[sw] = r.sw[sw] || []).push(siguiente.t); else r.sw[sw] = siguiente.t; k++; }
        } else r.banderas.add(sw);
      } else if (r.nombre !== "=") r.args.push(tk.t);
    }
    return r;
  }

  // ------------------------------------------------------------ formatos
  const MESES = { es: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
    en: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"] };
  // (abreviaturas de Windows en español de México, con su punto: «nov.», «mié.»)
  const MESES_C = { es: ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sep.", "oct.", "nov.", "dic."], en: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"] };
  const DIAS_C = { es: ["dom.", "lun.", "mar.", "mié.", "jue.", "vie.", "sáb."], en: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"] };
  const DIAS = { es: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"], en: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"] };
  const lengua = (idioma) => (/^es/i.test(idioma || "") ? "es" : "en");
  const p2 = (n) => String(n).padStart(2, "0");
  /** Fecha con el formato de Word (\@). */
  function fecha(d, fmt, idioma) {
    const L = lengua(idioma);
    if (!fmt) fmt = L === "es" ? "dd/MM/yyyy" : "M/d/yyyy";
    let out = "", i = 0;
    const doce = /am\/pm|a\/p/i.test(fmt);
    while (i < fmt.length) {
      const c = fmt[i];
      if (c === "'") { const j = fmt.indexOf("'", i + 1); out += fmt.slice(i + 1, j < 0 ? fmt.length : j); i = j < 0 ? fmt.length : j + 1; continue; }
      let n = 1; while (fmt[i + n] === c) n++;
      const run = fmt.slice(i, i + n);
      if (/^am\/pm/i.test(fmt.slice(i))) { const pm = d.getHours() >= 12; const t = fmt.slice(i, i + 5); out += t === "AM/PM" ? (pm ? "PM" : "AM") : t === "am/pm" ? (pm ? "pm" : "am") : pm ? "PM" : "AM"; i += 5; continue; }
      if (c === "d") out += n === 1 ? d.getDate() : n === 2 ? p2(d.getDate()) : n === 3 ? DIAS_C[L][d.getDay()] : DIAS[L][d.getDay()];
      else if (c === "M") out += n === 1 ? d.getMonth() + 1 : n === 2 ? p2(d.getMonth() + 1) : n === 3 ? MESES_C[L][d.getMonth()] : MESES[L][d.getMonth()];
      else if (c === "y") out += n <= 2 ? p2(d.getFullYear() % 100) : String(d.getFullYear());
      else if (c === "h") { const h = d.getHours() % 12 || 12; out += n === 1 ? h : p2(h); }
      else if (c === "H") out += n === 1 ? d.getHours() : p2(d.getHours());
      else if (c === "m") out += n === 1 ? d.getMinutes() : p2(d.getMinutes());
      else if (c === "s") out += n === 1 ? d.getSeconds() : p2(d.getSeconds());
      else out += run;
      i += n;
      void doce;
    }
    return out;
  }
  /** Número con el formato de Word (\#): 0, #, x, «,» (miles), «.» (decimales), %, y lo demás tal cual. */
  function numero(v, fmt) {
    if (!fmt) return Number.isInteger(v) ? String(v) : String(Math.round(v * 1e10) / 1e10);
    const secciones = fmt.split(";");
    let f = secciones[0];
    if (v < 0 && secciones[1] != null) { f = secciones[1]; v = -v; }
    else if (v === 0 && secciones[2] != null) f = secciones[2];
    const m = /[0#x,.]+/.exec(f);
    if (!m) return f;
    const pat = m[0], pre = f.slice(0, m.index), post = f.slice(m.index + pat.length);
    // (en Word, el «%» de un formato es solo texto: no multiplica por 100)
    const [ent, dec = ""] = pat.split(".");
    const nd = dec.replace(/[^0#x]/g, "").length;
    const minDec = (dec.match(/0/g) || []).length;
    let t = Math.abs(v).toFixed(nd);
    let [a, b = ""] = t.split(".");
    // decimales con «#»: se quitan los ceros de más
    while (b.length > minDec && b.endsWith("0")) b = b.slice(0, -1);
    const minEnt = (ent.match(/0/g) || []).length;
    a = a.replace(/^0+(?=\d)/, "");
    if (a === "0" && minEnt === 0) a = "";
    a = a.padStart(minEnt, "0");
    if (ent.includes(",")) a = a.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    const signo = v < 0 && secciones[1] == null ? "-" : "";
    return pre + signo + a + (b ? "." + b : nd && minDec ? "." + b : "") + post;
  }
  const romano = (n) => { let r = ""; [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]].forEach(([v, l]) => { while (n >= v) { r += l; n -= v; } }); return r; };
  // Word: A…Z, luego AA…ZZ (la misma letra repetida)
  const alfabetico = (n) => { if (n < 1) return String(n); const l = String.fromCharCode(65 + ((n - 1) % 26)); return l.repeat(Math.floor((n - 1) / 26) + 1); };
  const UNOS = { en: ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"],
    es: ["cero", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve"] };
  const DECENAS = { en: ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"], es: ["", "", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"] };
  const VEINTI = ["veinte", "veintiuno", "veintidós", "veintitrés", "veinticuatro", "veinticinco", "veintiséis", "veintisiete", "veintiocho", "veintinueve"];
  const CIENTOS = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"];
  function enLetras(n, L) {
    n = Math.floor(Math.abs(n));
    if (L === "es") {
      const menor100 = (x) => (x < 20 ? UNOS.es[x] : x < 30 ? VEINTI[x - 20] : DECENAS.es[Math.floor(x / 10)] + (x % 10 ? " y " + UNOS.es[x % 10] : ""));
      const menor1000 = (x) => (x === 100 ? "cien" : (x >= 100 ? CIENTOS[Math.floor(x / 100)] + (x % 100 ? " " : "") : "") + (x % 100 || x < 100 ? menor100(x % 100) : ""));
      if (n < 1000) return menor1000(n);
      if (n < 1e6) { const m = Math.floor(n / 1000), r = n % 1000; return (m === 1 ? "mil" : menor1000(m) + " mil") + (r ? " " + menor1000(r) : ""); }
      const mi = Math.floor(n / 1e6), r = n % 1e6; return (mi === 1 ? "un millón" : enLetras(mi, "es") + " millones") + (r ? " " + enLetras(r, "es") : "");
    }
    const menor100 = (x) => (x < 20 ? UNOS.en[x] : DECENAS.en[Math.floor(x / 10)] + (x % 10 ? "-" + UNOS.en[x % 10] : ""));
    const menor1000 = (x) => (x >= 100 ? UNOS.en[Math.floor(x / 100)] + " hundred" + (x % 100 ? " " : "") : "") + (x % 100 || x < 100 ? menor100(x % 100) : "");
    if (n < 1000) return menor1000(n);
    if (n < 1e6) { const m = Math.floor(n / 1000), r = n % 1000; return menor1000(m) + " thousand" + (r ? " " + menor1000(r) : ""); }
    const mi = Math.floor(n / 1e6), r = n % 1e6; return enLetras(mi, "en") + " million" + (r ? " " + enLetras(r, "en") : "");
  }
  const ORD_EN = { one: "first", two: "second", three: "third", five: "fifth", eight: "eighth", nine: "ninth", twelve: "twelfth" };
  function ordenEnLetras(n, L) {
    if (L === "es") {
      const o = ["", "primero", "segundo", "tercero", "cuarto", "quinto", "sexto", "séptimo", "octavo", "noveno", "décimo"];
      return n <= 10 ? o[n] : enLetras(n, "es") + "º";
    }
    const t = enLetras(n, "en"), partes = t.split(/([ -])/), u = partes[partes.length - 1];
    partes[partes.length - 1] = ORD_EN[u] || (u.endsWith("y") ? u.slice(0, -1) + "ieth" : u + "th");
    return partes.join("");
  }
  // (Word en español escribe los números con letra con la primera en mayúscula: «Siete», «Mil quinientos con 00/100»)
  const mayusEs = (t, L) => (L === "es" ? t.charAt(0).toUpperCase() + t.slice(1) : t);
  const ordinal = (n, L) => (L === "es" ? n + "º" : n + (n % 100 >= 11 && n % 100 <= 13 ? "th" : ["th", "st", "nd", "rd"][n % 10] || "th"));
  /** \* (formato general) sobre un valor (número o texto). */
  function general(v, lista, idioma) {
    const L = lengua(idioma);
    let t = typeof v === "number" ? String(v) : String(v);
    for (const f0 of lista || []) {
      const f = String(f0);
      const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
      const entero = Number.isFinite(n) ? Math.round(n) : null;
      if (f === "ROMAN" && entero != null) t = romano(entero);
      else if (f === "roman" && entero != null) t = romano(entero).toLowerCase();
      else if (f === "ALPHABETIC" && entero != null) t = alfabetico(entero);
      else if (f === "alphabetic" && entero != null) t = alfabetico(entero).toLowerCase();
      else if (/^arabic$/i.test(f) && entero != null) t = String(entero);
      else if (/^ordinal$/i.test(f) && entero != null) t = ordinal(entero, L);
      else if (/^cardtext$/i.test(f) && entero != null) t = mayusEs(enLetras(entero, L), L);
      else if (/^ordtext$/i.test(f) && entero != null) t = mayusEs(ordenEnLetras(entero, L), L);
      else if (/^hex$/i.test(f) && entero != null) t = entero.toString(16).toUpperCase();
      else if (/^dollartext$/i.test(f) && Number.isFinite(n)) { const c = Math.round((n - Math.floor(n)) * 100); t = mayusEs(enLetras(Math.floor(n), L) + (L === "es" ? " con " : " and ") + p2(c) + "/100", L); }
      else if (f === "Upper" || f === "UPPER") t = t.toUpperCase();
      else if (f === "Lower" || f === "LOWER") t = t.toLowerCase();
      else if (/^caps$/i.test(f)) t = t.replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase());
      else if (/^firstcap$/i.test(f)) t = t.charAt(0).toUpperCase() + t.slice(1);
    }
    return t;
  }
  /** Los modificadores del campo sobre su valor. */
  function conFormato(v, c, idioma) {
    if (v instanceof Date) return fecha(v, c.sw["\\@"], idioma);
    let t = v;
    if (typeof v === "number" && c.sw["\\#"] != null) t = numero(v, c.sw["\\#"]);
    else if (typeof v === "number") t = numero(v, null);
    const lista = (c.sw["\\*"] || []).filter((f) => !/^(mergeformat|charformat)$/i.test(f));
    return lista.length ? general(typeof v === "number" && c.sw["\\#"] == null ? v : t, lista, idioma) : String(t);
  }

  // ------------------------------------------------------------ fórmulas (=)
  function expresion(src, buscar) {
    const s = String(src).replace(/\s+/g, " ");
    let i = 0;
    const ver = () => { while (s[i] === " ") i++; return s[i]; };
    const toma = (t) => { ver(); if (s.startsWith(t, i)) { i += t.length; return true; } return false; };
    const FUN = {
      ABS: (a) => Math.abs(a[0]), INT: (a) => Math.trunc(a[0]), MOD: (a) => a[0] % a[1], SIGN: (a) => Math.sign(a[0]),
      ROUND: (a) => { const k = Math.pow(10, a[1] || 0); return Math.round(a[0] * k) / k; }, SUM: (a) => a.reduce((x, y) => x + y, 0),
      PRODUCT: (a) => a.reduce((x, y) => x * y, 1), AVERAGE: (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0), COUNT: (a) => a.length,
      MAX: (a) => Math.max(...a), MIN: (a) => Math.min(...a), AND: (a) => (a.every(Boolean) ? 1 : 0), OR: (a) => (a.some(Boolean) ? 1 : 0), NOT: (a) => (a[0] ? 0 : 1),
      TRUE: () => 1, FALSE: () => 0, DEFINED: (a, crudos) => (crudos[0] != null && buscar(crudos[0]) != null ? 1 : 0),
    };
    function primario() {
      const c = ver();
      if (c === "(") { i++; const v = comparar(); toma(")"); return v; }
      if (c === "-") { i++; return -primario(); }
      if (c === "+") { i++; return primario(); }
      // (el punto es el decimal; la coma separa los argumentos: «SUM(1,2,3)»)
      const m = /^(\d+(?:\.\d+)?|\.\d+)/.exec(s.slice(i));
      if (m) { i += m[0].length; return parseFloat(m[0]); }
      const id = /^[A-Za-z_][\w]*/.exec(s.slice(i));
      if (id) {
        i += id[0].length;
        const n = id[0].toUpperCase();
        if (FUN[n] && ver() === "(") {
          i++;
          const args = [], crudos = [];
          if (ver() !== ")") do { const k0 = i; args.push(comparar()); crudos.push(s.slice(k0, i).trim()); } while (toma(","));
          toma(")");
          return FUN[n](args, crudos);
        }
        if (FUN[n] && (n === "TRUE" || n === "FALSE")) return FUN[n]();
        const v = buscar(id[0]);
        const x = parseFloat(String(v == null ? "" : v).replace(/[^\d.-]/g, ""));
        if (!Number.isFinite(x)) throw new Error("!Error de sintaxis");
        return x;
      }
      throw new Error("!Error de sintaxis");
    }
    function potencia() { let v = primario(); while (toma("^")) v = Math.pow(v, primario()); if (toma("%")) v /= 100; return v; }
    function termino() { let v = potencia(); for (;;) { if (toma("*")) v *= potencia(); else if (toma("/")) { const d = potencia(); if (!d) throw new Error("!Cero divide"); v /= d; } else return v; } }
    function suma() { let v = termino(); for (;;) { if (toma("+")) v += termino(); else if (toma("-")) v -= termino(); else return v; } }
    function comparar() {
      let v = suma();
      for (;;) {
        if (toma("<=")) v = v <= suma() ? 1 : 0; else if (toma(">=")) v = v >= suma() ? 1 : 0; else if (toma("<>")) v = v !== suma() ? 1 : 0;
        else if (toma("<")) v = v < suma() ? 1 : 0; else if (toma(">")) v = v > suma() ? 1 : 0; else if (toma("=")) v = v === suma() ? 1 : 0; else return v;
      }
    }
    const v = comparar();
    if (ver() != null) throw new Error("!Error de sintaxis");
    return v;
  }
  /** IF: «IF expr1 op expr2 "verdad" "falso"». */
  function condicion(c, buscar) {
    const a = c.args;
    if (a.length < 3) return null;
    const val = (x) => { const b = buscar(x); return b != null ? String(b) : x; };
    const OPS = ["<=", ">=", "<>", "=", "<", ">"];
    let izq = a[0], op = a[1], der = a[2], k = 3;
    // (si el operador vino pegado: «IF x=1 …»)
    if (!OPS.includes(op)) { const m = /^(.*?)(<=|>=|<>|=|<|>)(.*)$/.exec(a[0]); if (!m) return null; izq = m[1]; op = m[2]; der = m[3] || a[1]; k = m[3] ? 1 : 2; }
    const si = a[k] != null ? a[k] : "", no = a[k + 1] != null ? a[k + 1] : "";
    const x = val(izq), y = val(der);
    const nx = parseFloat(x), ny = parseFloat(y);
    let r;
    if (Number.isFinite(nx) && Number.isFinite(ny) && /^-?[\d.,]+$/.test(x.trim()) && /^-?[\d.,]+$/.test(y.trim())) r = { "=": nx === ny, "<>": nx !== ny, "<": nx < ny, ">": nx > ny, "<=": nx <= ny, ">=": nx >= ny }[op];
    else {
      // texto: = y <> con comodines ? y * (como Word)
      const re = new RegExp("^" + y.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\?/g, ".").replace(/\*/g, ".*") + "$", "i");
      const igual = op === "=" || op === "<>" ? re.test(x) : null;
      r = op === "=" ? igual : op === "<>" ? !igual : { "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y }[op];
    }
    return r ? si : no;
  }

  // ------------------------------------------------------------ el documento: orden, marcadores, títulos
  /**
   * Todo lo que hace falta para calcular: los campos en orden (con su párrafo), el texto de cada marcador, el estilo de
   * cada párrafo y las propiedades del documento. opc.paginaDe(idDelElemento) y opc.paginas: lo que dice la paginación.
   */
  function contexto(doc, opc) {
    opc = opc || {};
    const T = NOM.textoRico;
    const ps = T && T.recorrido ? T.recorrido(doc) : [];
    const campos = [], marcas = {}, abiertos = {};
    const textoPar = [];
    let palabras = 0, letras = 0;
    ps.forEach((p, ip) => {
      let t = "";
      p.items.forEach((it) => {
        if (it.e) {
          if (it.e.tipo === "campo") campos.push({ id: it.id, ref: it.e.ref, bl: it.bl, par: ip, pos: t.length });
          else if (it.e.tipo === "marcador") {
            if (it.e.lado === "ini") abiertos[it.e.nombre] = { par: ip, pos: t.length, id: it.id };
            else if (abiertos[it.e.nombre]) { marcas[it.e.nombre] = Object.assign({}, abiertos[it.e.nombre], { parF: ip, posF: t.length }); }
          }
          if (it.e.tipo === "campo" && it.bl) t += String(it.bl.res || "");
          return;
        }
        t += it.c;
      });
      textoPar.push(t);
      palabras += (t.match(/\S+/g) || []).length;
      letras += t.replace(/\s/g, "").length;
    });
    for (const n in abiertos) if (!marcas[n]) marcas[n] = Object.assign({}, abiertos[n], { parF: abiertos[n].par, posF: abiertos[n].pos });
    const textoMarca = (n) => {
      const m = marcas[n];
      if (!m) return null;
      if (m.par === m.parF) return textoPar[m.par].slice(m.pos, m.posF);
      const partes = [textoPar[m.par].slice(m.pos)];
      for (let q = m.par + 1; q < m.parF; q++) partes.push(textoPar[q]);
      partes.push(textoPar[m.parF].slice(0, m.posF));
      return partes.join("\r");
    };
    const props = (doc.valor && doc.valor("_doc", "props")) || {};
    const idioma = (doc.valor && doc.valor("_doc", "idioma")) || opc.idioma || "es-MX";
    return { doc, ps, campos, marcas, textoMarca, textoPar, props, idioma, palabras, letras, paginaDe: opc.paginaDe || (() => 1), paginas: opc.paginas || 1, nombre: opc.nombre || "", usuario: opc.usuario || "", ahora: opc.ahora || new Date() };
  }
  const fechaDe = (iso) => { const d = iso ? new Date(iso) : null; return d && !isNaN(d) ? d : null; };
  const PROPS = { TITLE: "titulo", SUBJECT: "asunto", AUTHOR: "autor", KEYWORDS: "claves", COMMENTS: "comentarios", LASTSAVEDBY: "modificoPor" };
  const DOCPROPS = { title: "titulo", subject: "asunto", author: "autor", keywords: "claves", comments: "comentarios", lastsavedby: "modificoPor", company: "company", manager: "manager", category: "categoria", template: "template" };

  /**
   * El valor de un campo (ya con su formato) o null si NOM no lo sabe calcular (se queda el último). estado: lo que
   * llevan los SEQ (se va llenando en orden).
   */
  function calcular(campo, ctx, estado) {
    const bl = campo.bl || {};
    const c = leer(bl.instr);
    const buscar = (n) => { const t = ctx.textoMarca(n); return t == null ? null : t; };
    // el idioma del texto del campo (el de Word: meses, números con letra); si no, el del documento
    ctx = Object.assign({}, ctx, { idioma: bl.idioma || ctx.idioma });
    const f = (v) => conFormato(v, c, ctx.idioma);
    switch (c.nombre) {
      case "DATE": return f(ctx.ahora);
      case "TIME": return fecha(ctx.ahora, c.sw["\\@"] || (lengua(ctx.idioma) === "es" ? "HH:mm" : "h:mm am/pm"), ctx.idioma);
      case "CREATEDATE": { const d = fechaDe(ctx.props.creado); return d ? f(d) : null; }
      case "SAVEDATE": { const d = fechaDe(ctx.props.guardado); return d ? f(d) : null; }
      case "PRINTDATE": { const d = fechaDe(ctx.props.impreso); return d ? f(d) : null; }
      case "PAGE": return f(ctx.paginaDe(campo.id));
      case "NUMPAGES": case "SECTIONPAGES": return f(ctx.paginas);
      case "SECTION": return f(1);
      case "NUMWORDS": return f(ctx.palabras);
      case "NUMCHARS": return f(ctx.letras);
      case "FILENAME": return f(ctx.nombre ? ctx.nombre + (/\.docx?$/i.test(ctx.nombre) ? "" : ".docx") : "");
      case "USERNAME": return ctx.usuario ? f(ctx.usuario) : null;
      case "USERINITIALS": return ctx.usuario ? f(ctx.usuario.split(/\s+/).map((x) => x[0] || "").join("").toUpperCase()) : null;
      case "TITLE": case "SUBJECT": case "AUTHOR": case "KEYWORDS": case "COMMENTS": case "LASTSAVEDBY": {
        const v = ctx.props[PROPS[c.nombre]];
        return v != null ? f(v) : f("");
      }
      case "DOCPROPERTY": { const k = DOCPROPS[String(c.args[0] || "").toLowerCase()]; const v = k ? ctx.props[k] : null; return v != null ? f(v) : null; }
      case "QUOTE": return f(c.args.join(" "));
      case "SEQ": {
        const id = String(c.args[0] || "").toLowerCase();
        if (!id) return null;
        const s = estado.seq;
        // \s n: vuelve a 1 después de cada título de nivel n
        const nivel = c.sw["\\s"] ? +c.sw["\\s"] : 0;
        if (nivel) {
          const ultimo = estado.titulos[nivel] || -1;
          if ((s["_t" + id] ?? -2) !== ultimo) { s[id] = 0; s["_t" + id] = ultimo; }
        }
        if (c.sw["\\r"] != null) s[id] = +c.sw["\\r"];
        else if (!c.banderas.has("\\c")) s[id] = (s[id] || 0) + 1;
        else s[id] = s[id] || 0;
        if (c.banderas.has("\\h")) return "";
        return f(s[id]);
      }
      case "REF": case "": {
        const n = c.nombre === "REF" ? c.args[0] : null;
        const t = n ? buscar(n) : null;
        if (t == null) return null;
        if (c.banderas.has("\\p")) return null;
        return f(t.replace(/\r/g, " "));
      }
      case "PAGEREF": {
        const m = ctx.marcas[c.args[0]];
        if (!m) return null;
        const pg = ctx.paginaDe(m.id);
        if (c.banderas.has("\\p")) {
          const aqui = ctx.paginaDe(campo.id);
          const es = lengua(ctx.idioma) === "es";
          if (pg === aqui) return m.par < campo.par || (m.par === campo.par && m.pos < campo.pos) ? (es ? "arriba" : "above") : (es ? "abajo" : "below");
          return (es ? "en la página " : "on page ") + f(pg);
        }
        return f(pg);
      }
      case "IF": { const r = condicion(c, buscar); return r == null ? null : f(r); }
      case "=": {
        try { return f(expresion(c.expr, buscar)); }
        catch (e) { return String(e.message || "!Error de sintaxis"); }
      }
      case "STYLEREF": {
        const nombre = String(c.args[0] || "").toLowerCase();
        const nivel = (/(heading|título|titulo)\s*(\d)/.exec(nombre) || [])[2] || (/^\d$/.test(nombre) ? nombre : null);
        if (!nivel) return null;
        const est = "h" + Math.min(3, +nivel);
        for (let q = campo.par; q >= 0; q--) if (ctx.ps[q] && ctx.ps[q].pa && ctx.ps[q].pa.estilo === est) return f(ctx.textoPar[q]);
        return null;
      }
      default: return null;   // MERGEFIELD, NOTEREF y los demás: el último valor
    }
  }

  /** Calcula todos los campos en el orden del documento: [{ campo, valor }] (valor null = se queda el último). */
  function calcularTodos(ctx) {
    const estado = { seq: {}, titulos: {} };
    const porPar = new Map();
    ctx.campos.forEach((c) => { if (!porPar.has(c.par)) porPar.set(c.par, []); porPar.get(c.par).push(c); });
    const r = [];
    ctx.ps.forEach((p, ip) => {
      const est = p.pa && p.pa.estilo;
      const nv = est === "h1" ? 1 : est === "h2" ? 2 : est === "h3" ? 3 : 0;
      if (nv) for (let k = nv; k <= 9; k++) estado.titulos[k] = ip;
      for (const c of porPar.get(ip) || []) r.push({ campo: c, valor: calcular(c, ctx, estado) });
    });
    return r;
  }

  /** Lo que se ve de un campo ahora (sin «Actualizar»): su último valor; la fecha y la hora, al día (como Word). */
  function textoVisible(bl, ctx0) {
    if (!bl) return "";
    const c = leer(bl.instr);
    if ((c.nombre === "DATE" || c.nombre === "TIME") && !c.banderas.has("\\l")) {
      const idioma = bl.idioma || (ctx0 && ctx0.idioma) || "es-MX";
      try { return c.nombre === "DATE" ? conFormato(new Date(), c, idioma) : fecha(new Date(), c.sw["\\@"] || (lengua(idioma) === "es" ? "HH:mm" : "h:mm am/pm"), idioma); }
      catch (e) { return String(bl.res || ""); }
    }
    return String(bl.res == null ? "" : bl.res);
  }

  NOM.campos = { leer, fecha, numero, general, conFormato, expresion, condicion, contexto, calcular, calcularTodos, textoVisible, enLetras };
})();

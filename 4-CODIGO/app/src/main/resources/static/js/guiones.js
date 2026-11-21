/* NOM · División automática de palabras, como «Guiones automáticos» de Word.

   Word parte las palabras al final del renglón con las reglas del IDIOMA DEL TEXTO (un documento en inglés con texto en
   español se parte con las reglas del inglés: «con-ex-ión», «in-cluye»). Los puntos donde Word puede partir cada palabra
   salen de Word 2024 real (herramientas/guiones: 70 000 palabras en español y en inglés, cada una con los dos idiomas),
   aprendidos como patrones de Liang con patgen; js/guiones-es.js y js/guiones-en-us.js se cargan solo cuando un documento
   los pide.

   Cuándo parte Word (medido en Word 2024, pruebas/office/analisis/guiones_decision.py; el acomodo está en modo-texto.js):
   nunca deja menos de 3 letras al principio del renglón siguiente ni menos de 2 al final del de arriba, ni más guiones
   seguidos que «Limitar guiones consecutivos»; la «zona de división» no cambia nada en Word 2013 y después.

   Cómo se dibuja en el navegador (los párrafos que no acomoda modo-texto.js): en cada lugar donde una palabra se puede
   partir se pone un guion suave dentro de un elemento marcado (<span class="wd-guion">&shy;</span>); el navegador parte
   ahí si hace falta y dibuja el guion. Ese elemento NO es parte del documento: el editor lo ignora al leer lo escrito y al
   contar posiciones (modo-texto.js). */
(function () {
  "use strict";
  const NOM = window.NOM;
  const ARCHIVOS = { "en-us": "js/guiones-en-us.js", es: "js/guiones-es.js" };
  const compilados = {}, cargando = {};
  /** El idioma de Word («en-US», «es-MX») → el juego de patrones que hay. */
  function juego(idioma) {
    const l = String(idioma || "").toLowerCase();
    if (l.startsWith("es")) return "es";
    if (l.startsWith("en")) return "en-us";
    return null;
  }
  function compilar(j) {
    if (compilados[j]) return compilados[j];
    const d = (window.NOM_GUIONES || {})[j];
    if (!d) return null;
    const t = new Map();
    let largo = 0;
    for (const p of d.patrones.split(/\s+/)) {
      if (!p) continue;
      const letras = p.replace(/\d/g, "");
      const v = new Array(letras.length + 1).fill(0);
      let i = 0;
      for (const ch of p) { if (ch >= "0" && ch <= "9") v[i] = +ch; else i++; }
      t.set(letras, v);
      if (letras.length > largo) largo = letras.length;
    }
    const exc = new Map();
    for (const w of (d.excepciones || "").split(/\s+/)) if (w) exc.set(w.replace(/-/g, ""), w);
    return (compilados[j] = { t, largo, exc, min: d.min || [2, 3] });
  }
  /** Carga los patrones de un idioma (una sola vez). Devuelve una promesa con true si los hay. */
  function cargar(idioma) {
    const j = juego(idioma);
    if (!j) return Promise.resolve(false);
    if (compilados[j] || (window.NOM_GUIONES || {})[j]) return Promise.resolve(!!compilar(j));
    if (cargando[j]) return cargando[j];
    cargando[j] = new Promise((ok) => {
      const s = document.createElement("script");
      s.src = ARCHIVOS[j];
      s.onload = () => ok(!!compilar(j));
      s.onerror = () => ok(false);
      document.head.append(s);
    });
    return cargando[j];
  }
  const listo = (idioma) => { const j = juego(idioma); return !!(j && (compilados[j] || ((window.NOM_GUIONES || {})[j] && compilar(j)))); };
  /** Dónde se puede partir una palabra: posiciones (entre letras) en orden. */
  function puntos(palabra, idioma) {
    const j = juego(idioma), c = j && compilar(j);
    if (!c) return [];
    const w = palabra.toLowerCase();
    const [minI, minD] = c.min;
    if (w.length < minI + minD) return [];
    const ex = c.exc.get(w);
    if (ex) { const r = []; let k = 0; for (const ch of ex) { if (ch === "-") r.push(k); else k++; } return r; }
    // (Word solo parte dentro de las primeras 64 letras de una palabra: en una de 96, su último lugar es el 61, en los dos
    // idiomas)
    if (w.length > LARGO_WORD) return puntos(w.slice(0, LARGO_WORD), idioma).filter((p) => p <= LARGO_WORD - 3);
    const s = "." + w + ".", n = s.length, v = new Array(n + 1).fill(0);
    for (let i = 0; i < n; i++) {
      for (let k = i + 1; k <= Math.min(n, i + c.largo); k++) {
        const p = c.t.get(s.slice(i, k));
        if (p) for (let q = 0; q < p.length; q++) if (p[q] > v[i + q]) v[i + q] = p[q];
      }
    }
    const r = [];
    // v[k] es el valor entre s[k-1] y s[k]; en la palabra, el corte después de su letra m está en v[m + 1]
    for (let m = minI; m <= w.length - minD; m++) if (v[m + 1] % 2 === 1) r.push(m);
    return r;
  }
  const LETRA = /[A-Za-zÀ-ÖØ-öø-ÿĀ-ſ]/;
  const LARGO_WORD = 64;
  /** Letras que Word deja, como mínimo, al principio del renglón siguiente. */
  const MIN_SIGUE = 3;
  // Lo que va justo antes de una palabra y no deja partirla (medido en Word 2024: «/colaboración» no se parte; «(colaboración»,
  // «"colaboración», «#colaboración» o «colaboración/» sí). Lo de después nunca estorba.
  const NO_ANTES = /[\d@/\\_.:;,+&=*~^|\u2010-\u2015-]/;
  /**
   * Dónde se puede partir cada palabra de un texto (posiciones del texto, en orden), con las reglas de Word: palabras de 4
   * letras o más, al menos 3 letras al principio del renglón siguiente, nada en una palabra en MAYÚSCULAS si opc.mayus es
   * falso, y nada en la que va pegada después de NO_ANTES. Con `todos`, también esas: Word las usa cuando lo que no tiene
   * espacios (una dirección web) no cabe ni en un renglón entero. opc: { idioma, mayus }.
   */
  function lugares(txt, opc, todos) {
    const idioma = opc && opc.idioma, out = [];
    if (!listo(idioma)) return out;
    const re = /[A-Za-zÀ-ÖØ-öø-ÿĀ-ſ]{4,}/g;
    let m;
    while ((m = re.exec(txt))) {
      const w = m[0], antes = txt[m.index - 1];
      if (antes && (LETRA.test(antes) || (!todos && NO_ANTES.test(antes)))) continue;
      if (opc.mayus === false && w === w.toUpperCase()) continue;
      // (Word nunca deja menos de 3 letras al principio del renglón siguiente en el texto corrido)
      for (const p of puntos(w, idioma)) if (w.length - p >= MIN_SIGUE) out.push(m.index + p);
    }
    return out;
  }
  /**
   * Un texto ya escapado para HTML con los lugares para partir (el texto original va en `txt`, sin escapar). opc:
   * { idioma, mayus } (mayus: falso = no partir palabras en MAYÚSCULAS, como «No dividir palabras en mayúsculas»).
   */
  function marcar(txt, esc, opc) {
    const ls = lugares(txt, opc);
    if (!ls.length) return esc(txt);
    let out = "", k = 0;
    for (const p of ls) { out += esc(txt.slice(k, p)) + '<span class="wd-guion">\u00ad</span>'; k = p; }
    return out + esc(txt.slice(k));
  }
  NOM.guiones = { cargar, listo, puntos, lugares, marcar, juego, MIN_SIGUE };
})();

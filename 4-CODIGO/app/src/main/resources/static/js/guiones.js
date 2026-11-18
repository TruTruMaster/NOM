/* NOM · División automática de palabras, como «Guiones automáticos» de Word.

   Word parte las palabras al final del renglón con las reglas del IDIOMA DEL TEXTO (un documento en inglés con texto en
   español se parte con las reglas del inglés: «con-exión», «in-cluye»). Aquí se usan los patrones de TeX (algoritmo de
   Liang, los mismos que usan TeX y muchos procesadores de texto; js/vendor/guiones-*.js, con su licencia) y se cargan
   solo cuando un documento los pide.

   Cómo se dibuja: en cada lugar donde una palabra se puede partir se pone un guion suave dentro de un elemento
   marcado (<span class="wd-guion">&shy;</span>); el navegador parte ahí si hace falta y dibuja el guion. Ese elemento NO
   es parte del documento: el editor lo ignora al leer lo escrito y al contar posiciones (modo-texto.js). */
(function () {
  "use strict";
  const NOM = window.NOM;
  const ARCHIVOS = { "en-us": "js/vendor/guiones-en-us.js", es: "js/vendor/guiones-es.js" };
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
  /**
   * Un texto ya escapado para HTML con los lugares para partir (el texto original va en `txt`, sin escapar). opc:
   * { idioma, mayus } (mayus: falso = no partir palabras en MAYÚSCULAS, como «No dividir palabras en mayúsculas»).
   */
  function marcar(txt, esc, opc) {
    const idioma = opc && opc.idioma;
    if (!listo(idioma)) return esc(txt);
    let out = "";
    const re = /[A-Za-zÀ-ÖØ-öø-ÿĀ-ſ]{4,}/g;
    let ult = 0, m;
    while ((m = re.exec(txt))) {
      const w = m[0];
      out += esc(txt.slice(ult, m.index));
      ult = m.index + w.length;
      // (una palabra pegada a otra cosa —un número, un guion— se deja: Word tampoco la parte así)
      const antes = txt[m.index - 1], despues = txt[ult];
      if ((antes && (LETRA.test(antes) || /[\d@/\\_]/.test(antes))) || (despues && /[\d@/\\_]/.test(despues)) || (opc.mayus === false && w === w.toUpperCase())) { out += esc(w); continue; }
      const ps = puntos(w, idioma);
      if (!ps.length) { out += esc(w); continue; }
      let k = 0;
      for (const p of ps) { out += esc(w.slice(k, p)) + '<span class="wd-guion">­</span>'; k = p; }
      out += esc(w.slice(k));
    }
    return out + esc(txt.slice(ult));
  }
  NOM.guiones = { cargar, listo, puntos, marcar, juego };
})();

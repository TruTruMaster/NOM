/* NOM · Idioma de la interfaz (español / inglés) y detección del idioma del documento.
   - Preferencia guardada en localStorage ("nom.pref.extra" → idiomaInterfaz: "auto" | "es" | "en"),
     sin tocar Preferencias.idioma del servidor.
   - En español no observa el DOM (cero costo). En inglés traduce nodos de texto y los atributos
     aria-label, title, placeholder y alt usando js/textos/en.json (y restaura al volver a español).
   - Excluye el contenido de documentos, chat, comentarios y nombres escritos por las personas, con
     verificación en O(1) por lote para no costar ni 1 ms por tecla al escribir.
   - Expone NOM.idioma.detectar(texto) y aplica lang + spellcheck="true" al contenedor del documento. */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});
  const propio = (o, k) => !!o && Object.prototype.hasOwnProperty.call(o, k);

  const CLAVE_EXTRA = "nom.pref.extra";
  const ATRIBUTOS = ["aria-label", "title", "placeholder", "alt"];
  const TAGS_SIN_TEXTO = new Set(["SCRIPT", "STYLE", "NOSCRIPT", "SVG", "TEXTAREA", "CODE", "PRE"]);
  const TAGS_IGNORAR_TODO = new Set(["SCRIPT", "STYLE", "NOSCRIPT", "SVG"]);
  // Clases reales de contenido de documentos, chat, comentarios y elementos que nunca se traducen.
  const CLASES_EXCLUIDAS = new Set([
    "wd-hoja", "wd-cuerpo", "wd-celda", "wd-paginas", "wd-cursores", "papel",
    "rejilla-hoja", "xl", "editor-celda",
    "diapo", "diapo-mini", "mundo", "lienzo", "miniatura", "bl",
    "ver-diapos", "ver-hoja", "ver-texto", "ver-diag",
    "burbuja-chat", "hilo-cita", "menciones-caja",
    "ed-nombre", "avatar",
  ]);
  const PADRES_TITULO_USUARIO = new Set(["doc-tarjeta", "tarjeta-espacio", "continuar"]);
  const CLASE_TITULO = new Set(["titulo"]);
  const CLASE_CRECE = new Set(["crece"]);
  const CLASE_RIEL_LISTA = new Set(["riel-lista"]);
  const CLASE_HILO_O_RES = new Set(["hilo-msj", "resultado"]);
  const CLASE_FRAGMENTO = new Set(["fragmento"]);
  const NO_TRADUCIR_LITERAL = new Set(["Español", "English", "NOM"]);

  const diccionarios = { es: null, en: null };
  const cargando = {};
  const origTexto = new WeakMap();
  const origAtrib = new WeakMap();
  let tituloOriginal = null;

  let activo = "es";
  let tablaActual = null;
  let observador = null;
  let aplicando = false;

  // ================================================================== preferencia y resolución
  function leerPreferencia() {
    try {
      const g = NOM.leer ? NOM.leer(CLAVE_EXTRA, null) : JSON.parse(localStorage.getItem(CLAVE_EXTRA) || "null");
      const v = g && g.idiomaInterfaz;
      if (v === "es" || v === "en" || v === "auto") return v;
    } catch (e) { /* sin almacenamiento */ }
    return "auto";
  }

  function idiomaSistema() {
    const nav = (typeof navigator !== "undefined" && navigator) || (typeof window !== "undefined" && window.navigator) || null;
    if (!nav) return "es";
    const lista = Array.isArray(nav.languages) && nav.languages.length ? nav.languages : (nav.language ? [nav.language] : []);
    if (!lista.length) return "es";
    const primero = String(lista[0] || "").trim().toLowerCase();
    if (!primero) return "es";
    return primero.startsWith("es") ? "es" : "en";
  }

  function resolver(pref) {
    if (pref === "es" || pref === "en") return pref;
    return idiomaSistema();
  }

  function fijarLangHtml(cod) {
    if (document && document.documentElement) document.documentElement.setAttribute("lang", cod);
  }

  // ================================================================== carga de diccionarios JSON
  function fijarDiccionario(cod, tabla) {
    if (cod !== "es" && cod !== "en") return;
    diccionarios[cod] = tabla && typeof tabla === "object" ? tabla : {};
    if (activo === cod) tablaActual = diccionarios[cod];
  }

  async function cargarDiccionario(cod) {
    if (cod !== "es" && cod !== "en") return {};
    if (diccionarios[cod]) return diccionarios[cod];
    if (cargando[cod]) return cargando[cod];
    const url = "js/textos/" + cod + ".json";
    cargando[cod] = (async () => {
      try {
        if (typeof fetch === "function") {
          const r = await fetch(url);
          if (r && r.ok) {
            const d = await r.json();
            diccionarios[cod] = d && typeof d === "object" ? d : {};
            return diccionarios[cod];
          }
        }
      } catch (e) { /* sin red o entorno sin fetch */ }
      diccionarios[cod] = diccionarios[cod] || {};
      return diccionarios[cod];
    })();
    return cargando[cod];
  }

  // ================================================================== exclusión rápida (O(1) con caché por lote)
  function tieneClase(el, conjunto) {
    const cls = el.className;
    if (!cls || typeof cls !== "string") return false;
    if (!cls.includes(" ")) return conjunto.has(cls);
    const partes = cls.split(/\s+/);
    for (let i = 0; i < partes.length; i++) if (conjunto.has(partes[i])) return true;
    return false;
  }

  function esNodoExcluidoDirecto(el) {
    if (!el || el.nodeType !== 1) return false;
    if (TAGS_IGNORAR_TODO.has(el.tagName)) return true;
    const ce = el.getAttribute && el.getAttribute("contenteditable");
    if (ce !== null && String(ce).toLowerCase() !== "false") return true;
    if (el.isContentEditable === true && ce === null) return true;
    if (el.getAttribute) {
      if (el.getAttribute("translate") === "no") return true;
      if (el.hasAttribute("data-no-traducir") || el.hasAttribute("data-mundo") || el.hasAttribute("data-fx") || el.hasAttribute("data-esquema-lista") || el.hasAttribute("data-previa")) return true;
    }
    if (tieneClase(el, CLASES_EXCLUIDAS)) return true;
    // Nombres de documentos, espacios, comentarios y fragmentos de búsqueda escritos por personas.
    if (tieneClase(el, CLASE_TITULO)) {
      let p = el.parentNode;
      for (let k = 0; k < 3 && p && p.nodeType === 1; k++, p = p.parentNode) {
        if (tieneClase(p, PADRES_TITULO_USUARIO)) return true;
      }
    }
    if (tieneClase(el, CLASE_CRECE)) {
      let p = el.parentNode;
      for (let k = 0; k < 3 && p && p.nodeType === 1; k++, p = p.parentNode) {
        if (tieneClase(p, CLASE_RIEL_LISTA)) return true;
      }
    }
    if (el.tagName === "P" || el.tagName === "B") {
      const p = el.parentNode;
      if (p && p.nodeType === 1) {
        const gp = p.parentNode;
        if (tieneClase(p, CLASE_HILO_O_RES) || (gp && gp.nodeType === 1 && tieneClase(gp, CLASE_HILO_O_RES))) return true;
      }
    }
    if (tieneClase(el, CLASE_FRAGMENTO)) return true;
    return false;
  }

  /** Verifica si `el` o alguno de sus ancestros está excluido, guardando el resultado en `cache`. */
  function estaExcluido(el, cache) {
    if (!el || el.nodeType !== 1) return false;
    if (cache && cache.has(el)) return cache.get(el);
    const camino = [];
    let cur = el;
    let res = false;
    while (cur && cur.nodeType === 1) {
      if (cache && cache.has(cur)) { res = cache.get(cur); break; }
      camino.push(cur);
      if (esNodoExcluidoDirecto(cur)) { res = true; break; }
      cur = cur.parentNode;
    }
    if (cache) {
      // Si un ancestro está excluido, todo el camino debajo también lo está; si no, tampoco.
      for (let i = 0; i < camino.length; i++) cache.set(camino[i], res);
    }
    return res;
  }

  // ================================================================== traducción exacta + patrones dinámicos
  const PATRONES_EN = [
    [/^(\d+) espacio(s)?$/, (_, n) => `${n} ${+n === 1 ? "space" : "spaces"}`],
    [/^(\d+) documento(s)?$/, (_, n) => `${n} ${+n === 1 ? "document" : "documents"}`],
    [/^(\d+) palabra(s)?$/, (_, n) => `${n} ${+n === 1 ? "word" : "words"}`],
    [/^(\d+) solicitud(es)?$/, (_, n) => `${n} ${+n === 1 ? "request" : "requests"}`],
    [/^(\d+) documentos? · (\d+) personas?$/, (_, d, m) => `${d} ${+d === 1 ? "document" : "documents"} · ${m} ${+m === 1 ? "person" : "people"}`],
    [/^hace (\d+) min$/, (_, n) => `${n} min ago`],
    [/^hace (\d+) h$/, (_, n) => `${n} hr ago`],
    [/^Página (\d+) de (\d+)$/, (_, a, b) => `Page ${a} of ${b}`],
    [/^Versión (\d+) guardada$/, (_, n) => `Version ${n} saved`],
    [/^Columna ([A-Z]+)$/, (_, c) => `Column ${c}`],
    [/^Fila (\d+)$/, (_, f) => `Row ${f}`],
    [/^Faltan (\d+)$/, (_, n) => `${n} more needed`],
    [/^Sugerencias \((\d+)\)$/, (_, n) => `Suggestions (${n})`],
    [/^Resueltos \((\d+)\)$/, (_, n) => `Resolved (${n})`],
    [/^Errores \((\d+)\)$/, (_, n) => `Errors (${n})`],
    [/^Advertencias \((\d+)\)$/, (_, n) => `Warnings (${n})`],
    [/^Consejos \((\d+)\)$/, (_, n) => `Tips (${n})`],
    [/^Papelera \((\d+)\)$/, (_, n) => `Trash (${n})`],
    [/^Ver la papelera \((\d+)\)$/, (_, n) => `View trash (${n})`],
    [/^Antes: (.+)$/, (_, k, d) => "Before: " + traducirClaveTecla(k, d)],
    [/^Guardados en (.+)$/, (_, s, d) => "Saved in " + (propio(d, s) ? d[s] : s)],
    [/^Nuevo documento en (.+)$/, (_, s) => `New document in ${s}`],
    [/^Invitar a (.+) con código QR$/, (_, s) => `Invite to ${s} with QR code`],
    [/^Exportar todo (.+) \(\.zip\)$/, (_, s) => `Download all of ${s} (.zip)`],
    [/^Más opciones de (.+)$/, (_, s) => `More options for ${s}`],
    [/^Chat · (.+)$/, (_, s) => `Chat · ${s}`],
    [/^Descargar (Word|Excel|PowerPoint)(.*)$/, (_, a, b) => `Download ${a}${b}`],
    [/^(\d+) de (\d+) · unos minutos y tu equipo estará trabajando aquí\.$/, (_, a, b) => `${a} of ${b} · a few minutes and your team will be up and running.`],
    [/^Esa tecla ya la usa «(.+)»\. Quítala de ahí primero o elige otra\.$/, (_, nom, d) => `That key is already used by “${propio(d, nom) ? d[nom] : nom}”. Remove it there first or pick another.`],
    [/^Nada con “(.+)”$/, (_, q) => `Nothing matching “${q}”`],
    [/^([0-9.,]+ [KMGB]+) usados$/, (_, v) => `${v} used`],
    [/^([0-9.,]+ [KMGB]+) de ([0-9.,]+ [KMGB]+)$/, (_, a, b) => `${a} of ${b}`],
    [/^([0-9.,]+ [KMGB]+) · sin límite aquí$/, (_, a) => `${a} · unlimited here`],
    [/^(\d+) documentos? listos? para abrir sin Internet$/, (_, n) => `${n} ${+n === 1 ? "document" : "documents"} ready offline`],
    [/^(\d+) solo aquí$/, (_, n) => `${n} on this device only`],
    [/^(\d+) cambios? por enviar$/, (_, n) => `${n} ${+n === 1 ? "change" : "changes"} pending`],
    [/^([0-9.,]+ [KMGB]+) de ajustes y sesión$/, (_, v) => `${v} for settings & session`],
  ];

  function traducirClaveTecla(texto, dict) {
    return String(texto || "").split(" + ").map((p) => (propio(dict, p) ? dict[p] : p)).join(" + ");
  }

  function traducirCadena(raw, cod, dict) {
    if (!raw || !dict) return raw;
    const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(raw);
    const medio = m[2];
    if (!medio || NO_TRADUCIR_LITERAL.has(medio)) return raw;
    const clave = medio.replace(/\s+/g, " ");
    if (propio(dict, clave)) return m[1] + dict[clave] + m[3];
    if (cod === "en" && clave.length <= 220) {
      for (let i = 0; i < PATRONES_EN.length; i++) {
        const rx = PATRONES_EN[i][0], fn = PATRONES_EN[i][1];
        const mm = rx.exec(clave);
        if (mm) return m[1] + fn(...mm, dict) + m[3];
      }
    }
    return raw;
  }

  function traducirTextoNodo(nodo, cod, dict) {
    const padre = nodo.parentNode;
    if (padre && padre.nodeType === 1 && TAGS_SIN_TEXTO.has(padre.tagName)) return;
    const actual = nodo.nodeValue;
    if (!actual || !actual.trim()) return;
    if (cod === "es") {
      if (origTexto.has(nodo)) {
        const base = origTexto.get(nodo);
        const nuevo = dict ? traducirCadena(base, "es", dict) : base;
        if (nodo.nodeValue !== nuevo) nodo.nodeValue = nuevo;
      }
      return;
    }
    const base = origTexto.has(nodo) ? origTexto.get(nodo) : actual;
    const nuevo = traducirCadena(base, cod, dict);
    if (nuevo !== actual) {
      if (!origTexto.has(nodo)) origTexto.set(nodo, base);
      nodo.nodeValue = nuevo;
    }
  }

  function traducirAtributos(el, cod, dict) {
    if (!el || el.nodeType !== 1 || !el.getAttribute) return;
    for (let i = 0; i < ATRIBUTOS.length; i++) {
      const a = ATRIBUTOS[i];
      const val = el.getAttribute(a);
      if (!val || !val.trim()) continue;
      let mapa = origAtrib.get(el);
      if (cod === "es") {
        if (mapa && propio(mapa, a)) {
          const base = mapa[a];
          const nuevo = dict ? traducirCadena(base, "es", dict) : base;
          if (val !== nuevo) el.setAttribute(a, nuevo);
        }
        continue;
      }
      const base = mapa && propio(mapa, a) ? mapa[a] : val;
      const nuevo = traducirCadena(base, cod, dict);
      if (nuevo !== val) {
        if (!mapa) { mapa = {}; origAtrib.set(el, mapa); }
        if (!propio(mapa, a)) mapa[a] = base;
        el.setAttribute(a, nuevo);
      }
    }
  }

  function recorrer(raiz, cod, dict, cache) {
    if (!raiz) return;
    if (raiz.nodeType === 3) {
      if (!estaExcluido(raiz.parentNode, cache)) traducirTextoNodo(raiz, cod, dict);
      return;
    }
    if (raiz.nodeType !== 1) return;
    if (estaExcluido(raiz, cache)) return;
    traducirAtributos(raiz, cod, dict);
    const hijos = raiz.childNodes;
    if (!hijos || !hijos.length) return;
    for (let i = 0; i < hijos.length; i++) {
      const h = hijos[i];
      if (h.nodeType === 3) traducirTextoNodo(h, cod, dict);
      else if (h.nodeType === 1) recorrer(h, cod, dict, cache);
    }
  }

  // ================================================================== observador del DOM (solo en inglés)
  function desconectar() {
    if (observador) { observador.disconnect(); observador = null; }
  }

  function conectar() {
    if (observador || typeof MutationObserver !== "function" || !document || !document.body) return;
    observador = new MutationObserver((records) => {
      if (aplicando || activo !== "en" || !tablaActual) return;
      aplicando = true;
      try {
        const cache = new Map();
        for (let i = 0; i < records.length; i++) {
          const rec = records[i];
          const tgt = rec.target;
          const padre = tgt && (tgt.nodeType === 1 ? tgt : tgt.parentNode);
          // Corte inmediato en O(1) para mutaciones dentro de contenteditable o documentos.
          if (!padre || estaExcluido(padre, cache)) continue;
          const nodos = rec.addedNodes;
          if (!nodos || !nodos.length) continue;
          for (let j = 0; j < nodos.length; j++) recorrer(nodos[j], "en", tablaActual, cache);
        }
      } finally {
        aplicando = false;
      }
    });
    observador.observe(document.body, { childList: true, subtree: true });
  }

  function traducirTitulo(cod, dict) {
    if (!document || typeof document.title !== "string" || !document.title) return;
    if (tituloOriginal == null) tituloOriginal = document.title;
    if (cod === "es") { document.title = dict ? traducirCadena(tituloOriginal, "es", dict) : tituloOriginal; return; }
    document.title = traducirCadena(tituloOriginal, cod, dict);
  }

  async function aplicar(pref) {
    const cod = resolver(pref !== undefined ? pref : leerPreferencia());
    activo = cod;
    fijarLangHtml(cod);
    if (cod === "es") {
      desconectar();
      tablaActual = diccionarios.es;
      if (document && document.body) {
        aplicando = true;
        try {
          recorrer(document.body, "es", tablaActual, new Map());
          traducirTitulo("es", tablaActual);
        } finally {
          aplicando = false;
        }
      }
      return "es";
    }
    const dict = await cargarDiccionario("en");
    if (activo !== "en") return activo;
    tablaActual = dict;
    if (document && document.body) {
      aplicando = true;
      try {
        recorrer(document.body, "en", dict, new Map());
        traducirTitulo("en", dict);
      } finally {
        aplicando = false;
      }
      conectar();
    }
    return "en";
  }

  // ================================================================== detección del idioma del documento (pura)
  const PALABRAS_ES = new Set([
    "el", "la", "los", "las", "un", "una", "unos", "unas", "de", "del", "al", "en", "por", "para",
    "con", "sobre", "entre", "hacia", "desde", "hasta", "que", "qué", "como", "cómo", "cuando",
    "cuándo", "donde", "dónde", "porque", "pero", "sino", "aunque", "mientras", "pues", "entonces",
    "también", "además", "muy", "más", "menos", "todo", "toda", "todos", "todas", "este", "esta",
    "estos", "estas", "ese", "esa", "esos", "esas", "aquí", "allí", "ahí", "ahora", "después",
    "antes", "siempre", "nunca", "solo", "sólo", "cada", "otro", "otra", "otros", "otras", "mismo",
    "misma", "ser", "estar", "tener", "hacer", "poder", "decir", "haber", "es", "está", "están",
    "fue", "eran", "era", "somos", "eres", "soy", "hay", "han", "había", "tiene", "tienen", "puede",
    "pueden", "hace", "hacen", "su", "sus", "tu", "tus", "mi", "mis", "nuestro", "nuestra", "se",
    "le", "les", "lo", "nos", "te", "ya", "sí", "si", "ni", "y", "o", "u", "bien", "trabajo",
    "equipo", "documento", "página", "texto", "archivo", "espacio", "fecha", "nombre", "tiempo",
    "clase", "historia", "informe", "proyecto", "hola", "pantalla", "gráficas", "tablas", "guardamos",
  ]);
  const PALABRAS_EN = new Set([
    "the", "and", "of", "to", "in", "is", "that", "for", "it", "with", "as", "was", "on", "are",
    "be", "by", "this", "from", "or", "have", "an", "at", "not", "but", "they", "which", "you",
    "we", "were", "his", "her", "their", "its", "our", "your", "my", "can", "will", "would",
    "should", "could", "may", "might", "must", "been", "being", "had", "does", "did", "do", "am",
    "if", "when", "where", "what", "who", "why", "how", "than", "then", "there", "here", "these",
    "those", "into", "over", "under", "about", "after", "before", "between", "through", "during",
    "without", "within", "against", "among", "because", "while", "although", "however", "also",
    "just", "only", "very", "more", "most", "some", "any", "many", "much", "each", "every", "other",
    "another", "such", "all", "both", "few", "same", "own", "new", "first", "last", "good", "great",
    "little", "make", "work", "time", "year", "people", "way", "day", "document", "page", "file",
    "team", "project", "text", "hello", "report", "class", "history", "screen", "charts", "tables",
  ]);

  /** Función pura: detecta si un texto está en español ("es") o en inglés ("en"). */
  function detectar(texto, defecto) {
    const def = defecto === "en" ? "en" : "es";
    const s = String(texto || "").slice(0, 4000);
    if (!s.trim()) return def;
    const palabras = s.toLowerCase().match(/[a-záéíóúüñ]+/g);
    if (!palabras || !palabras.length) return def;
    let es = 0, en = 0;
    if (/[¿¡]/.test(s)) es += 3;
    for (let i = 0; i < palabras.length; i++) {
      const w = palabras[i];
      if (/[ñáéíóúü]/.test(w)) es += 2;
      if (PALABRAS_ES.has(w)) es += 2;
      if (PALABRAS_EN.has(w)) en += 2;
    }
    if (en > es) return "en";
    if (es > en) return "es";
    return def;
  }

  /** Detecta el idioma del contenido de un contenedor de documento y ajusta lang + spellcheck="true". */
  function aplicarIdiomaDocumento(cont, texto) {
    if (!cont || cont.nodeType !== 1) return "es";
    const muestra = texto != null ? String(texto) : String(cont.textContent || cont.value || "");
    const lang = detectar(muestra, activo);
    if (cont.getAttribute("lang") !== lang) cont.setAttribute("lang", lang);
    if (cont.getAttribute("spellcheck") !== "true") cont.setAttribute("spellcheck", "true");
    return lang;
  }

  function actualizarDocumentosAbiertos() {
    if (!document || !document.querySelectorAll) return;
    const sels = ".wd-cuerpo, .diapo[data-mundo], .mundo[data-mundo], .rejilla-hoja";
    const nodos = document.querySelectorAll(sels);
    for (let i = 0; i < nodos.length; i++) {
      const n = nodos[i];
      if ((n.textContent || "").trim().length >= 8) aplicarIdiomaDocumento(n);
    }
  }

  function programarDeteccionDoc() {
    const h = (typeof location !== "undefined" && location.hash) || "";
    if (!/^#\/(doc|local|ver)\//.test(h)) return;
    setTimeout(actualizarDocumentosAbiertos, 250);
    setTimeout(actualizarDocumentosAbiertos, 1000);
  }

  // Arranque: pone lang en <html> de inmediato; solo carga/observa inglés si hace falta.
  activo = resolver(leerPreferencia());
  fijarLangHtml(activo);
  if (activo === "en") {
    if (document && document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", () => { aplicar("en"); programarDeteccionDoc(); }, { once: true });
    } else {
      aplicar("en");
    }
  }
  if (typeof window !== "undefined" && window.addEventListener) {
    window.addEventListener("hashchange", programarDeteccionDoc);
  }

  NOM.idioma = {
    actual: () => activo,
    preferencia: leerPreferencia,
    sistema: idiomaSistema,
    resolver,
    aplicar,
    cargar: cargarDiccionario,
    fijarDiccionario,
    traducir: (texto) => (activo === "en" && tablaActual ? traducirCadena(texto, "en", tablaActual) : texto),
    detectar,
    aplicarIdiomaDocumento,
    actualizarDocumentosAbiertos,
    observando: () => !!observador,
  };
})();

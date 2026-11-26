/* NOM · Almacén de binarios: las imágenes y el Office original de un documento van aparte del documento.

   Dentro del documento iban como texto base64 (un tercio más pesadas) y repetidas (una vez dentro del original de
   Office y otra en cada bloque que las usa): un .pptx de 10 MB pesaba 30 MB dentro de NOM y no cabía. Ahora el
   documento guarda solo una referencia corta, «nom:bin/HUELLA.ext», y los bytes viven:
   - en el servidor, en el almacén del espacio (una sola vez; ServicioBinarios), con los permisos del espacio;
   - en este equipo, en IndexedDB («nom-binarios»): lo importado aquí se ve sin Internet y se sube cuando hay red;
   - en la memoria mientras se usan (con su URL para dibujarlos).
   En un espacio cifrado de extremo a extremo los bytes salen de aquí YA cifrados con la llave del documento: el
   servidor no ve ni el contenido ni su huella real (la huella es la de lo cifrado). El enlace para ver y la pantalla
   del proyector traen la llave del documento y los descifran igual.
   Si el servidor no tiene almacén (un teléfono con una versión vieja), todo sigue como antes, en base64.

   Al dibujar: cualquier <img>, <image>, <video>, <audio>, <source> o estilo con una referencia se cambia solo por su
   imagen (un observador de la página), así que el resto de la interfaz no tiene que saber nada de esto. */
(function () {
  "use strict";
  const NOM = window.NOM;
  const PRE = "nom:bin/";
  const RE = /^nom:bin\/([0-9a-f]{64})(?:\.([a-z0-9]{1,8}))?$/;
  const RE_EN = /nom:bin\/[0-9a-f]{64}(?:\.[a-z0-9]{1,8})?/g;
  const MIMES = {
    png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp", svg: "image/svg+xml", bmp: "image/bmp",
    tif: "image/tiff", tiff: "image/tiff", emf: "image/x-emf", wmf: "image/x-wmf", ico: "image/x-icon", avif: "image/avif",
    mp4: "video/mp4", webm: "video/webm", mov: "video/quicktime", m4a: "audio/mp4", mp3: "audio/mpeg", wav: "audio/wav", ogg: "audio/ogg",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation", pdf: "application/pdf", bin: "application/octet-stream",
  };
  const EXT = { "image/jpeg": "jpg", "image/svg+xml": "svg", "audio/mpeg": "mp3", "audio/x-wav": "wav", "audio/wave": "wav", "image/x-png": "png" };
  /** Lo más chico se queda dentro del documento (íconos, viñetas): no vale la pena una petición por eso. */
  const UMBRAL = 12 * 1024;
  const ES_DATOS = /^data:([a-z0-9.+\/-]+)(?:;[a-z0-9=._-]+)*;base64,/i;

  /** La llave de un documento en texto (base64 de URL, como la que llevan el enlace para ver y el QR) → bytes. */
  const RESERVADO = new Set(["__proto__", "constructor", "prototype"]);
  const propio = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
  const uidActual = () => (NOM.api && NOM.api.usuario && (NOM.api.usuario() || {}).id) || null;
  const llaveDeTexto = (t) => { if (!t) return null; try { return deB64(String(t).replace(/-/g, "+").replace(/_/g, "/") + "===".slice((String(t).length + 3) % 4)); } catch (e) { return null; } };
  const esRef = (v) => typeof v === "string" && v.length < 90 && RE.test(v);
  const huellaDe = (ref) => (RE.exec(ref) || [])[1] || null;
  const mimeDe = (ref) => { const e = (RE.exec(ref) || [])[2] || "bin"; return propio(MIMES, e) ? MIMES[e] : "application/octet-stream"; };
  const extDe = (mime) => { const m = String(mime || "").toLowerCase(); if (propio(EXT, m)) return EXT[m]; const e = Object.keys(MIMES).find((k) => MIMES[k] === m); return e || "bin"; };
  const hex = (u) => Array.from(u, (b) => b.toString(16).padStart(2, "0")).join("");
  const deB64 = (s) => { const b = atob(s), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; };
  const aB64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };

  // ------------------------------------------------------------ IndexedDB (lo importado aquí, para sin Internet)
  let dbP = null;
  function db() {
    if (!dbP) dbP = new Promise((ok) => {
      try {
        const r = indexedDB.open("nom-binarios", 1);
        r.onupgradeneeded = () => { const s = r.result.createObjectStore("b", { keyPath: "h" }); s.createIndex("pendiente", "pendiente"); };
        r.onsuccess = () => ok(r.result);
        r.onerror = () => ok(null);
      } catch (e) { ok(null); }
    });
    return dbP;
  }
  async function idb(modo, f) {
    const d = await db();
    if (!d) return null;
    return new Promise((ok) => {
      let res = null;
      try {
        const t = d.transaction("b", modo);
        res = f(t.objectStore("b"));
        t.oncomplete = () => ok(res && "result" in res ? res.result : null);
        t.onerror = t.onabort = () => ok(null);
      } catch (e) { ok(null); }
    });
  }
  const idbPut = (x) => idb("readwrite", (s) => s.put(x));
  const idbGet = (h) => idb("readonly", (s) => s.get(h));
  const idbDel = (h) => idb("readwrite", (s) => s.delete(h));
  const idbTodos = () => idb("readonly", (s) => s.getAll());

  // ------------------------------------------------------------ cifrado (espacios de extremo a extremo)
  const C = () => NOM.cripto;
  async function cifrar(llave, datos) {
    const s = await C().cifrar(llave, datos, C().algPreferido());
    const r = new Uint8Array(1 + s.iv.length + s.ct.length);
    r[0] = s.alg === "gcm" ? 1 : 2; r.set(s.iv, 1); r.set(s.ct, 1 + s.iv.length);
    return r;
  }
  async function descifrar(llave, b) {
    if (!b || !b.length || (b[0] !== 1 && b[0] !== 2)) throw new Error("Formato de archivo cifrado no válido");
    const alg = b[0] === 1 ? "gcm" : "ctr", n = alg === "gcm" ? 12 : 16;
    if (b.length < 1 + n) throw new Error("Archivo cifrado incompleto");
    return C().descifrar(llave, { alg, iv: b.subarray(1, 1 + n), ct: b.subarray(1 + n) });
  }
  async function huella(bytes) {
    try { if (window.crypto && crypto.subtle && window.isSecureContext !== false) return hex(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes))); }
    catch (e) { /* sin WebCrypto: la propia */ }
    return hex(C().sha256(bytes));
  }

  // ------------------------------------------------------------ contexto: de dónde se bajan y con qué llave
  // { espacio, llave (Uint8Array: la del documento, en un espacio cifrado), ver (código del enlace para ver),
  //   proyeccion (código del proyector) }
  let ctx = {};
  /** De qué contexto vino cada huella (si se usa en varios documentos abiertos). */
  const ctxDe = new Map();
  function contexto(c) {
    ctx = Object.assign({}, c || {});
    // lo que se quiso dibujar antes de saber de dónde bajarlo: otra vez
    setTimeout(() => { if (typeof document !== "undefined" && document.body) reintentar(document.body); }, 0);
    return ctx;
  }

  // ------------------------------------------------------------ memoria (lo que se está usando)
  const mem = new Map();        // huella → { bytes (en claro), mime, url }
  const cargando = new Map();   // huella → Promise<bytes>
  let enMemoria = 0;
  const TOPE_MEMORIA = 160 * 1024 * 1024;
  function recordar(h, bytes, mime) {
    const v = mem.get(h);
    if (v) { mem.delete(h); mem.set(h, v); return v; }   // al final: lo más reciente
    const x = { bytes, mime, url: null };
    mem.set(h, x);
    enMemoria += bytes.length;
    // lo que no se usa hace rato sale de la memoria (sigue en IndexedDB o en el servidor)
    for (const [k, y] of mem) {
      if (enMemoria <= TOPE_MEMORIA) break;
      if (k === h) continue;
      if (y.url) { try { URL.revokeObjectURL(y.url); } catch (e) { /* nada */ } }
      enMemoria -= y.bytes.length;
      mem.delete(k);
    }
    return x;
  }

  // ------------------------------------------------------------ subir
  /** ¿El servidor actual tiene almacén? (null = no se sabe todavía) */
  let hayAlmacen = null;
  async function subirUno(reg) {
    if (!reg || !reg.espacio || !NOM.api || !NOM.api.binario) return false;
    try {
      await NOM.api.binario("PUT", `/api/espacios/${encodeURIComponent(reg.espacio)}/binarios/${encodeURIComponent(reg.h)}`, reg.datos);
      hayAlmacen = true;
      await idbPut(Object.assign({}, reg, { pendiente: 0 }));
      return true;
    } catch (e) {
      const est = e && e.estado;
      // (una ruta que no existe es un servidor de antes, sin almacén; un 404 con su propio mensaje es que el espacio ya no está)
      if (est === 405 || (est === 404 && (!e.message || /^No existe/.test(e.message)))) hayAlmacen = false;
      else if (est === 413) {
        // Demasiado grande para el servidor: nunca entrará. Se queda en este equipo y no se vuelve a mandar.
        await idbPut(Object.assign({}, reg, { pendiente: 0, rechazado: est }));
      } else reintentarLuego();
      return false;
    }
  }
  /** Lo que no subió por algo pasajero (sin red, servidor ocupado): otra vez en 30-60 s, con azar para no llegar todos juntos. */
  let reintentoT = null;
  function reintentarLuego() {
    if (reintentoT) return;
    reintentoT = setTimeout(() => { reintentoT = null; subirPendientes(); }, 30000 + Math.random() * 30000);
  }
  let subiendo = null;
  /** Sube lo que quedó pendiente (sin red, o la app se cerró a la mitad). */
  function subirPendientes() {
    if (subiendo) return subiendo;
    const activo = uidActual();
    if (!activo || !NOM.api || !NOM.api.token || !NOM.api.token()) return Promise.resolve();
    subiendo = (async () => {
      const todos = (await idbTodos()) || [];
      for (const r of todos) if (r && r.pendiente && (!r.uid || r.uid === activo)) await subirUno(r);
    })().finally(() => { subiendo = null; });
    return subiendo;
  }
  async function limpiar(borrarDisco, usuarioId) {
    clearTimeout(reintentoT); reintentoT = null;
    for (const y of mem.values()) { if (y && y.url) { try { URL.revokeObjectURL(y.url); } catch (e) { /* nada */ } } }
    mem.clear();
    cargando.clear();
    ctxDe.clear();
    ctx = {};
    enMemoria = 0;
    if (borrarDisco) {
      const objetivo = usuarioId || uidActual();
      const todos = (await idbTodos()) || [];
      for (const r of todos) if (r && r.h && (!objetivo || !r.uid || r.uid === objetivo)) await idbDel(r.h);
    }
  }
  window.addEventListener("online", () => { subirPendientes(); });
  // (algunos WebView no mandan «online»: la API avisa cuando una petición vuelve a salir bien)
  if (NOM.api && NOM.api.alCambiarRed) NOM.api.alCambiarRed((ok) => { if (ok) setTimeout(() => { subirPendientes(); }, 500); });

  /**
   * Guarda unos bytes en el almacén del espacio y devuelve su referencia. c: { espacio, llave }.
   * Queda en este equipo de inmediato y se sube (si no hay red, cuando la haya). `esperar`: no regresa hasta que el
   * servidor lo tenga (para crear un documento que lo usa: así nadie lo abre antes de que exista).
   */
  async function guardar(bytes, mime, c, esperar) {
    c = c || ctx;
    const sale = c.llave ? await cifrar(c.llave, bytes) : bytes;
    const h = await huella(sale);
    const reg = { h, uid: uidActual(), espacio: c.espacio || null, mime: mime || "application/octet-stream", datos: sale, cifrado: c.llave ? 1 : 0, pendiente: 1, t: Date.now() };
    recordar(h, bytes, reg.mime);
    ctxDe.set(h, c);
    const antes = esperar ? await idbGet(h) : null;
    await idbPut(reg);
    const p = subirUno(reg);
    if (esperar) {
      const ok = await p;
      if (!ok && hayAlmacen === false) throw Object.assign(new Error("Este servidor no tiene almacén de archivos"), { sinAlmacen: true });
      if (!ok) {
        // Quien espera va a crear un documento con esto: sin subirlo, no lo crea (se guarda en este equipo con los
        // datos dentro). Que no se quede pendiente de subir, salvo que ya lo estuviera por otra cosa.
        if (!(antes && antes.pendiente && antes.espacio === reg.espacio)) idbPut(Object.assign({}, reg, { pendiente: 0 }));
        throw new NOM.api.ErrorApi(0, null);
      }
    }
    return PRE + h + "." + extDe(mime);
  }
  /** Una clave al azar para crear un documento: si la respuesta se pierde y se pide otra vez, el servidor da el mismo. */
  function claveNueva() {
    const u = new Uint8Array(16);
    (window.crypto || window.msCrypto).getRandomValues(u);
    return hex(u);
  }
  /** ¿Contesta el servidor? (para distinguir «sin red» de «la respuesta tardó o se perdió») */
  async function servidorContesta() {
    try { await NOM.api.get("/api/red/info", undefined, { ms: 6000 }); return true; } catch (e) { return false; }
  }

  // ------------------------------------------------------------ bajar
  async function bajarCrudo(h, c) {
    if (c.ver) return NOM.api.binario("GET", `/api/ver/${encodeURIComponent(c.ver)}/binarios/${encodeURIComponent(h)}`, null, { sinSesion: true });
    if (c.proyeccion) return NOM.api.binario("GET", `/api/proyeccion/${encodeURIComponent(c.proyeccion)}/binarios/${encodeURIComponent(h)}`, null, { sinSesion: true });
    if (!c.espacio) throw new Error("No se sabe de qué espacio es este archivo");
    return NOM.api.binario("GET", `/api/espacios/${encodeURIComponent(c.espacio)}/binarios/${encodeURIComponent(h)}`);
  }
  /** Los bytes (en claro) de una referencia: de la memoria, de este equipo o del servidor. */
  function bytes(ref, c) {
    const h = huellaDe(ref);
    if (!h) return Promise.reject(new Error("No es una referencia"));
    const v = mem.get(h);
    if (v) return Promise.resolve(v.bytes);
    if (cargando.has(h)) return cargando.get(h);
    c = c || ctxDe.get(h) || ctx;
    const p = (async () => {
      let reg = await idbGet(h), crudo = reg && reg.datos;
      if (!crudo) {
        crudo = await bajarCrudo(h, c);
        // guardado en este equipo: la siguiente vez (o sin Internet) no se vuelve a bajar
        idbPut({ h, uid: uidActual(), espacio: c.espacio || null, mime: mimeDe(ref), datos: crudo, cifrado: c.llave ? 1 : 0, pendiente: 0, t: Date.now() });
      }
      const claro = (reg ? reg.cifrado : c.llave) ? await descifrar(c.llave, crudo) : crudo;
      recordar(h, claro, (reg && reg.mime) || mimeDe(ref));
      ctxDe.set(h, c);
      return claro;
    })().finally(() => cargando.delete(h));
    cargando.set(h, p);
    return p;
  }
  /** Solo de la memoria (para lo que no puede esperar: los escritores de Office, ya con todo precargado). */
  function bytesYa(ref) { const v = mem.get(huellaDe(ref)); return v ? v.bytes : null; }
  /** Una URL para dibujarlo (blob:…). */
  async function url(ref, c) {
    const h = huellaDe(ref);
    const b = await bytes(ref, c);
    const v = mem.get(h) || recordar(h, b, mimeDe(ref));
    if (!v.url) {
      // un EMF o WMF (dibujos de Windows) se dibuja convertido a SVG (emf.js); el documento guarda el original
      const meta = NOM.emf && NOM.emf.tipo(v.bytes) ? NOM.emf.aSvg(v.bytes) : null;
      v.url = URL.createObjectURL(meta ? new Blob([meta.svg], { type: "image/svg+xml" }) : new Blob([v.bytes], { type: v.mime || mimeDe(ref) }));
    }
    return v.url;
  }
  const urlYa = (ref) => { const v = mem.get(huellaDe(ref)); return v && v.url; };
  /** data:…;base64 (para lo que necesita el binario dentro: un .nom, copiar a otro espacio). */
  function datosYa(ref) { const v = mem.get(huellaDe(ref)); return v ? "data:" + (v.mime || mimeDe(ref)) + ";base64," + aB64(v.bytes) : null; }

  // ------------------------------------------------------------ al escribir HTML: nunca «nom:bin/…» crudo
  const GIF = "data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==";
  /**
   * Para un src o un url() de CSS: la imagen ya resuelta (blob:…) o, mientras llega, un GIF transparente que lleva la
   * referencia en su fragmento («data:…#nom:bin/…»): el navegador lo carga sin error (con «nom:bin/…» crudo lo bloquea
   * la política de seguridad y deja un error en la consola) y el observador lo cambia por la imagen al llegar.
   */
  function paraSrc(v) {
    if (!esRef(v)) return NOM.emf ? NOM.emf.paraVer(v) : v;
    const u = urlYa(v);
    if (u) return u;
    url(v).catch(() => {});
    return GIF + "#" + v;
  }
  /** Para lo que se descarga con la imagen dentro (HTML, Markdown, SVG): sus datos (ya precargados), no una URL local. */
  const paraExportar = (v) => { const d = esRef(v) ? datosYa(v) || "" : v; return NOM.emf ? NOM.emf.paraVer(d) : d; };

  // ------------------------------------------------------------ recorrer valores
  /** Todas las referencias dentro de un valor (texto, lista u objeto). */
  function refsEn(v, s) {
    s = s || new Set();
    if (typeof v === "string") { if (v.indexOf(PRE) >= 0) (v.match(RE_EN) || []).forEach((r) => s.add(r)); }
    else if (Array.isArray(v)) v.forEach((x) => refsEn(x, s));
    else if (v && typeof v === "object") for (const k in v) refsEn(v[k], s);
    return s;
  }
  /** Las referencias que usa un documento. */
  function refsDoc(doc) {
    const s = new Set();
    const b = doc && doc.estado && doc.estado.bloques;
    if (!b) return s;
    for (const id in b) {
      if (doc.vivo && !doc.vivo(id)) continue;
      for (const k in b[id]) refsEn(b[id][k] && b[id][k].v, s);
    }
    return s;
  }
  /** Trae a la memoria todo lo que usa el documento (antes de exportarlo o de guardarlo como .nom). */
  async function precargar(doc, c) {
    const refs = [...refsDoc(doc)];
    const fallas = [];
    await Promise.all(refs.map((r) => bytes(r, c).catch((e) => { fallas.push(r); void e; })));
    if (fallas.length) throw Object.assign(new Error(`No se pudieron traer ${fallas.length} archivo(s) del documento (¿sin Internet y no estaban en este equipo?)`), { fallas });
    return refs.length;
  }
  /** Copia de un valor con cada referencia cambiada por sus datos (ya precargados). */
  function conDatos(v) {
    if (typeof v === "string") {
      if (v.indexOf(PRE) < 0) return v;
      if (esRef(v)) return datosYa(v) || v;
      return v.replace(RE_EN, (r) => datosYa(r) || r);
    }
    if (Array.isArray(v)) return v.map(conDatos);
    if (v && typeof v === "object") { const o = {}; for (const k in v) if (propio(v, k) && !RESERVADO.has(k)) o[k] = conDatos(v[k]); return o; }
    return v;
  }

  /**
   * Antes de crear un documento con lo importado: cada imagen grande (data:…;base64) va al almacén y en su lugar
   * queda su referencia; el Office original (los trozos de _origen) también, como un solo binario. c: { espacio, llave }.
   * Si el servidor no tiene almacén, regresa las operaciones como estaban.
   */
  /** Sin almacén: como antes, un original de más de 14 MB no se guarda dentro del documento (no cabría). */
  function sinOrigenGrande(ops) {
    const og = ops.find((o) => o.bloque === "_origen" && o.campos && o.campos.tipo === "origen");
    const max = (NOM.origen && NOM.origen.MAX_SIN_ALMACEN) || 14 * 1024 * 1024;
    return og && +og.campos.tam > max ? ops.filter((o) => o.bloque !== "_origen") : ops;
  }
  async function sacarDeOps(ops, c) {
    if (!c || !c.espacio || !NOM.api || !NOM.api.binario || hayAlmacen === false) return sinOrigenGrande(ops);
    const vistos = new Map();   // data URL → referencia (la misma imagen en muchos bloques: una sola vez)
    const nuevas = [];
    try {
      // 1) el original de Office: los trozos p0…pN de _origen → un binario
      const ogBase = ops.find((o) => o.bloque === "_origen" && o.campos && o.campos.tipo === "origen");
      if (ogBase && +ogBase.campos.n > 0) {
        let b64 = "";
        for (let i = 0; i < +ogBase.campos.n; i++) {
          const pz = ops.find((o) => o.bloque === "_origen" && o.campos && typeof o.campos["p" + i] === "string");
          if (!pz) { b64 = null; break; }
          b64 += pz.campos["p" + i];
        }
        if (b64) {
          const fmt = String(ogBase.campos.fmt || "bin");
          const ref = await guardar(deB64(b64), propio(MIMES, fmt) ? MIMES[fmt] : "application/octet-stream", c, true);
          // (los trozos pN pueden estar repartidos entre varias operaciones, también junto con otros campos: se quitan
          // de todas, y las que se quedan vacías, fuera)
          const sinTrozos = (campos) => { const r = {}; for (const k in campos) if (!/^p\d+$/.test(k)) r[k] = campos[k]; return r; };
          ops = ops.map((o) => {
            if (o.bloque !== "_origen" || !o.campos) return o;
            if (o === ogBase) return Object.assign({}, o, { campos: Object.assign(sinTrozos(o.campos), { bin: ref, n: 0 }) });
            const campos = sinTrozos(o.campos);
            return Object.keys(campos).length ? Object.assign({}, o, { campos }) : null;
          }).filter(Boolean);
        }
      }
      // 2) imágenes grandes en cualquier campo
      const cambiar = async (v) => {
        if (typeof v === "string") {
          if (v.length < UMBRAL || v.charCodeAt(0) !== 100) return v;
          const m = ES_DATOS.exec(v);
          if (!m) return v;
          if (vistos.has(v)) return vistos.get(v);
          let u;
          try { u = deB64(v.slice(m[0].length)); } catch (e) { return v; }
          const ref = await guardar(u, m[1].toLowerCase(), c, true);
          vistos.set(v, ref);
          return ref;
        }
        if (Array.isArray(v)) { const r = []; for (const x of v) r.push(await cambiar(x)); return r; }
        if (v && typeof v === "object") { const o = {}; for (const k in v) if (propio(v, k) && !RESERVADO.has(k)) o[k] = await cambiar(v[k]); return o; }
        return v;
      };
      for (const o of ops) nuevas.push(o.campos && o.bloque !== "_origen" ? Object.assign({}, o, { campos: await cambiar(o.campos) }) : o);
      // las huellas del original (qué no cambió al exportar) se calculan otra vez, ya con las referencias
      const og2 = nuevas.find((o) => o.bloque === "_origen" && o.campos && o.campos.tipo === "origen");
      if (og2 && vistos.size && c.modo && NOM.formatos && NOM.formatos.huellasOrigen) {
        try {
          const h = NOM.formatos.huellasOrigen(String(og2.campos.fmt || ""), c.nombre || "Documento", c.modo, nuevas);
          const i = nuevas.indexOf(og2);
          nuevas[i] = Object.assign({}, og2, { campos: Object.assign({}, og2.campos, { h: JSON.stringify(h) }) });
        } catch (e) { if (typeof console !== "undefined") console.warn("No se pudieron recalcular las huellas del original", e); }
      }
      return nuevas;
    } catch (e) {
      if (e && e.sinAlmacen) return sinOrigenGrande(ops);   // servidor sin almacén: como antes
      throw e;
    }
  }

  /**
   * Crear en un espacio un documento con lo importado: los binarios grandes al almacén (en un espacio cifrado, con la
   * llave del documento, que se genera aquí y se le pasa a e2e.js) y el documento solo con sus referencias.
   */
  async function crearEnEspacio(esp, nombre, modo, ops, opc) {
    opc = Object.assign({}, opc || {});
    const clave = opc.clave || claveNueva(), sinContexto = !!opc.sinContexto;
    delete opc.clave; delete opc.sinContexto;
    // Sin red (el sistema lo sabe): ni se toca el almacén; quien llama lo guarda en este equipo y lo sube al volver.
    if (navigator.onLine === false) throw new NOM.api.ErrorApi(0, null);
    const cifrado = !!(esp && (esp.e2e || (NOM.e2e && NOM.e2e.espacioCifrado && NOM.e2e.espacioCifrado(esp.id))));
    const kd = cifrado && NOM.cripto ? NOM.cripto.azar(32) : null;
    const listas = await sacarDeOps(ops, { espacio: esp.id, llave: kd, nombre, modo });
    const cuerpo = { nombre, modo, ops: NOM.DocCRDT.partir(listas), clave };
    if (kd) cuerpo._llaveDoc = kd;
    const ruta = `/api/espacios/${encodeURIComponent(esp.id)}/documentos`;
    let d;
    try { d = await NOM.api.post(ruta, cuerpo, Object.assign({ ms: 60000 }, opc)); }
    catch (e) {
      // ¿De verdad sin red? Si el servidor contesta, la respuesta tardó más de la cuenta o se perdió en el camino: se
      // pide otra vez con la misma clave (si ya lo había creado, el servidor da el mismo documento, no otro).
      if (!e || !e.red || !(await servidorContesta())) throw e;
      d = await NOM.api.post(ruta, cuerpo, Object.assign({}, opc, { ms: 240000 }));
    }
    if (d && d.id && !sinContexto) contexto({ espacio: esp.id, llave: kd, doc: d.id });
    return d;
  }

  // ------------------------------------------------------------ al dibujar: cambiar referencias por imágenes
  const ATRS = ["src", "href", "xlink:href", "poster"];
  const ES_IMG_SVG = (el) => { const t = (el.tagName || "").toLowerCase(); return t === "image" || t === "feimage"; };
  /** En un texto (atributo o estilo): cada referencia, cruda o en la marca «GIF#ref», por su URL si ya está. */
  function cambiarEn(t) {
    let falta = false;
    const r = t.replace(new RegExp("(?:" + GIF.replace(/[.*+?^${}()|[\]\\/]/g, "\\$&") + "#)?(" + RE_EN.source + ")", "g"), (todo, ref) => {
      const u = urlYa(ref);
      if (u) return u;
      falta = true;
      return todo;
    });
    return { r, falta, refs: t.match(RE_EN) || [] };
  }
  /** Un EMF o WMF que llegó como datos (data:image/x-emf…): se cambia por su dibujo en SVG. */
  function metaEl(el) {
    for (const a of ["src", "href"]) {
      if (a === "href" && !ES_IMG_SVG(el)) continue;
      const v = el.getAttribute(a);
      if (v && v.charCodeAt(0) === 100 && NOM.emf && NOM.emf.esDatos(v)) el.setAttribute(a, NOM.emf.paraVer(v));
    }
  }
  /** Cuántas veces se ha reintentado cada imagen que no se pudo mostrar todavía. */
  const intentos = new WeakMap();
  function resolverEl(el) {
    if (!el || el.nodeType !== 1 || !el.getAttribute) return;
    metaEl(el);
    for (const a of ATRS.concat(["style"])) {
      if ((a === "href" || a === "xlink:href") && !ES_IMG_SVG(el)) continue;
      const v = el.getAttribute(a);
      if (!v || v.indexOf(PRE) < 0) continue;
      const c = cambiarEn(v);
      if (c.r !== v) el.setAttribute(a, c.r);
      if (!c.falta) continue;
      // (mientras llega: un <img> con la referencia cruda pasa a la marca, para que no salga un ícono de imagen rota)
      if (a === "src" && esRef(v)) el.setAttribute(a, GIF + "#" + v);
      Promise.all(c.refs.map((r) => url(r).catch(() => null))).then(() => {
        const v2 = el.getAttribute(a);
        if (!v2 || v2.indexOf(PRE) < 0) return;
        const c2 = cambiarEn(v2);
        if (c2.r !== v2) el.setAttribute(a, c2.r);
        // Todavía no se pudo (en un espacio cifrado la llave del documento puede llegar después que la página, o la red
        // tardó): otra vez en un rato, mientras la imagen siga en pantalla.
        if (c2.falta && el.isConnected) {
          const n = (intentos.get(el) || 0) + 1;
          if (n <= 6) { intentos.set(el, n); setTimeout(() => { if (el.isConnected) resolverEl(el); }, 1000 * 2 ** (n - 1)); }
        }
      });
    }
  }
  /** Lo que se quedó esperando (por ejemplo, antes de saber de qué espacio bajarlo): otra vez. */
  function reintentar(raiz) { resolverTodo(raiz); }
  function resolverTodo(raiz) {
    if (!raiz || !raiz.querySelectorAll) return;
    if (raiz.nodeType === 1) resolverEl(raiz);
    raiz.querySelectorAll('[src*="nom:bin/"],[href*="nom:bin/"],[poster*="nom:bin/"],[style*="nom:bin/"],[src^="data:image/x-"],[href^="data:image/x-"]').forEach(resolverEl);
    // (xlink:href de SVG viejos)
    raiz.querySelectorAll("image").forEach(resolverEl);
  }
  let obs = null;
  function vigilar() {
    if (obs || typeof MutationObserver === "undefined" || !document.body) return;
    obs = new MutationObserver((ms) => {
      for (const m of ms) {
        if (m.type === "attributes") { const v = m.target && (m.target.getAttribute(m.attributeName) || ""); if (v && (v.indexOf(PRE) >= 0 || v.startsWith("data:image/x-"))) resolverEl(m.target); }
        else m.addedNodes.forEach((n) => { if (n.nodeType !== 1) return; const t = n.outerHTML || ""; if (t.indexOf(PRE) >= 0 || t.indexOf("data:image/x-") >= 0) resolverTodo(n); });
      }
    });
    obs.observe(document.body, { subtree: true, childList: true, attributes: true, attributeFilter: ["src", "href", "style", "poster"] });
    resolverTodo(document.body);
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", vigilar); else vigilar();
  // al arrancar: subir lo que haya quedado pendiente
  setTimeout(() => { if (navigator.onLine !== false) subirPendientes(); }, 4000);

  /** Atajos para quien escribe HTML con imágenes del documento (lienzo, texto, hoja, diapositivas, exportar). */
  NOM.srcBin = paraSrc;
  NOM.srcExport = paraExportar;
  NOM.binarios = {
    PRE, UMBRAL, esRef, huellaDe, mimeDe, extDe, contexto, guardar, bytes, bytesYa, url, urlYa, datosYa, refsEn, refsDoc,
    precargar, conDatos, sacarDeOps, crearEnEspacio, claveNueva, subirPendientes, resolverTodo, llaveDeTexto, paraSrc, paraExportar, limpiar,
    /** Para las pruebas: ¿tiene almacén el servidor? (null = no se ha probado) */
    get hayAlmacen() { return hayAlmacen; },
  };
})();

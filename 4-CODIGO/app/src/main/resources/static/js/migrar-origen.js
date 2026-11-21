/* NOM · la app de Android pasa a su propio origen (0.19).
   Hasta la 0.18 la interfaz se abría desde los archivos de la app (file://), y lo guardado en el teléfono vive en ese
   origen: la sesión y el servidor elegido (localStorage), la bóveda cifrada con las copias de los documentos y los
   cambios sin enviar (IndexedDB «nom-boveda») y los binarios («nom-binarios»). Desde la 0.19 la interfaz se sirve
   desde https://appassets.androidplatform.net (MainActivity), que es otro origen y empieza vacío.
   La primera vez, MainActivity abre esta página dos veces:
     1. en file:// (leer): manda todo por el puente, un mensaje por renglón (los grandes, en partes);
     2. en el origen nuevo (escribir): lo pone igual, con los mismos almacenes, índices y llaves.
   Lo de file:// no se borra. Fuera de la app (en el navegador) no hace nada. */
(function () {
  "use strict";
  const A = window.NomAndroid;
  const estado = (t) => { const e = document.getElementById("estado"); if (e) e.textContent = t; };
  if (!A || !A.migracionParte) { estado("Esta página solo sirve dentro de la app de NOM."); return; }
  const PARTE = 1 << 20;   // caracteres que cruzan el puente de una vez

  // ------------------------------------------------------------ valores de IndexedDB ↔ JSON
  const b64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const deB64 = (s) => { const b = atob(s), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; };
  const VISTAS = { Uint8Array, Int8Array, Uint8ClampedArray, Int16Array, Uint16Array, Int32Array, Uint32Array, Float32Array, Float64Array };
  /** Bytes, fechas, archivos, mapas y conjuntos van marcados con «$nom»; lo demás, tal cual. */
  async function aJson(v) {
    if (v === null || typeof v !== "object") return typeof v === "number" && !isFinite(v) ? { $nom: "n", v: String(v) } : v;
    if (v instanceof ArrayBuffer) return { $nom: "ab", b: b64(new Uint8Array(v)) };
    if (ArrayBuffer.isView(v)) return { $nom: "ta", t: v.constructor.name, b: b64(new Uint8Array(v.buffer, v.byteOffset, v.byteLength)) };
    if (v instanceof Date) return { $nom: "d", v: v.getTime() };
    if (typeof Blob !== "undefined" && v instanceof Blob) return { $nom: "blob", t: v.type, b: b64(new Uint8Array(await v.arrayBuffer())) };
    if (Array.isArray(v)) { const a = []; for (const x of v) a.push(await aJson(x)); return a; }
    if (v instanceof Map) { const a = []; for (const [k, x] of v) a.push([await aJson(k), await aJson(x)]); return { $nom: "map", v: a }; }
    if (v instanceof Set) { const a = []; for (const x of v) a.push(await aJson(x)); return { $nom: "set", v: a }; }
    const o = {};
    for (const k of Object.keys(v)) o[k] = await aJson(v[k]);
    return Object.prototype.hasOwnProperty.call(o, "$nom") ? { $nom: "o", v: o } : o;
  }
  function deJson(v) {
    if (v === null || typeof v !== "object") return v;
    if (Array.isArray(v)) return v.map(deJson);
    switch (v.$nom) {
      case "ab": return deB64(v.b).buffer;
      case "ta": {
        const u = deB64(v.b), T = Object.prototype.hasOwnProperty.call(VISTAS, v.t) ? VISTAS[v.t] : Uint8Array;
        return T === Uint8Array ? u : new T(u.buffer, 0, u.byteLength / T.BYTES_PER_ELEMENT);
      }
      case "d": return new Date(v.v);
      case "blob": return new Blob([deB64(v.b)], { type: v.t || "" });
      case "n": return Number(v.v);
      case "map": return new Map(v.v.map(([k, x]) => [deJson(k), deJson(x)]));
      case "set": return new Set(v.v.map(deJson));
      case "o": return objeto(v.v);
      default: return objeto(v);
    }
  }
  function objeto(v) {
    const o = {};
    for (const k of Object.keys(v)) Object.defineProperty(o, k, { value: deJson(v[k]), enumerable: true, writable: true, configurable: true });
    return o;
  }

  // ------------------------------------------------------------ IndexedDB
  const pedido = (r) => new Promise((ok, mal) => { r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error); });
  async function nombresBD() {
    const n = new Set(["nom-boveda", "nom-binarios"]);
    try { if (indexedDB.databases) for (const d of await indexedDB.databases()) if (d && d.name) n.add(d.name); } catch (e) { /* sin la lista */ }
    return [...n];
  }
  /** La base si existe (abrirla sin versión la crearía vacía: en ese caso se borra y da null). */
  function abrirSiExiste(nombre) {
    return new Promise((ok, mal) => {
      let nueva = false;
      const r = indexedDB.open(nombre);
      r.onupgradeneeded = () => { nueva = true; };
      r.onsuccess = () => {
        const db = r.result;
        if (!nueva) return ok(db);
        db.close(); try { indexedDB.deleteDatabase(nombre); } catch (e) { /* nada que borrar */ }
        ok(null);
      };
      r.onerror = () => mal(r.error);
      r.onblocked = () => mal(new Error("La base " + nombre + " está ocupada"));
    });
  }

  // ------------------------------------------------------------ 1. leer (file://)
  async function leer() {
    if (!A.migracionInicio()) return;
    let datos = 0, registros = 0;
    const mandar = (o) => {
      const t = JSON.stringify(o);
      for (let i = 0; i < t.length || i === 0; i += PARTE) if (!A.migracionParte(t.slice(i, i + PARTE), i + PARTE >= t.length)) throw new Error("El puente no aceptó un mensaje");
    };
    estado("Pasando la sesión y los ajustes…");
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k == null) continue;
      mandar({ t: "ls", k, v: localStorage.getItem(k) }); datos++;
    }
    if (window.indexedDB) {
      for (const nombre of await nombresBD()) {
        const db = await abrirSiExiste(nombre);
        if (!db) continue;
        const almacenes = [];
        if (db.objectStoreNames.length) {
          const tx = db.transaction([...db.objectStoreNames], "readonly");
          for (const a of db.objectStoreNames) {
            const s = tx.objectStore(a);
            almacenes.push({ n: a, kp: s.keyPath, ai: s.autoIncrement,
              ix: [...s.indexNames].map((x) => { const ix = s.index(x); return { n: x, kp: ix.keyPath, u: ix.unique, me: ix.multiEntry }; }) });
          }
        }
        mandar({ t: "bd", n: nombre, ver: db.version, almacenes });
        for (const al of almacenes) {
          estado("Pasando " + (nombre === "nom-binarios" ? "las imágenes y los archivos" : "los documentos guardados") + "…");
          const llaves = await pedido(db.transaction(al.n, "readonly").objectStore(al.n).getAllKeys());
          for (const k of llaves) {
            // cada registro en su propia lectura: con archivos grandes no se junta todo en memoria
            const v = await pedido(db.transaction(al.n, "readonly").objectStore(al.n).get(k));
            if (v === undefined) continue;
            mandar({ t: "r", n: nombre, a: al.n, k: al.kp == null ? await aJson(k) : undefined, v: await aJson(v) }); registros++;
          }
        }
        db.close();
      }
    }
    mandar({ t: "fin", datos, registros });
    A.migracionFin(true, datos + " datos y " + registros + " registros");
  }

  // ------------------------------------------------------------ 2. escribir (origen propio)
  function siguiente() {
    let t = "";
    for (;;) {
      const p = A.migracionSiguiente();
      if (!p) return t ? JSON.parse(t) : null;
      t += p.slice(1);
      if (p[0] === "1") return JSON.parse(t);
    }
  }
  function crearBD(m) {
    const abrir = (ver) => new Promise((ok, mal) => {
      const r = ver ? indexedDB.open(m.n, ver) : indexedDB.open(m.n);
      r.onupgradeneeded = () => {
        const db = r.result, tx = r.transaction;
        for (const al of m.almacenes) {
          const s = db.objectStoreNames.contains(al.n) ? tx.objectStore(al.n)
            : db.createObjectStore(al.n, al.kp == null ? { autoIncrement: !!al.ai } : { keyPath: al.kp, autoIncrement: !!al.ai });
          for (const ix of al.ix) if (!s.indexNames.contains(ix.n)) s.createIndex(ix.n, ix.kp, { unique: !!ix.u, multiEntry: !!ix.me });
        }
      };
      r.onsuccess = () => ok(r.result);
      r.onerror = () => mal(r.error);
      r.onblocked = () => mal(new Error("La base " + m.n + " está ocupada"));
    });
    // (si ya hay una de versión más nueva, se usa como está)
    return abrir(m.ver).catch((e) => (e && e.name === "VersionError" ? abrir(0) : Promise.reject(e)));
  }
  function poner(db, almacen, k, v) {
    return new Promise((ok, mal) => {
      const tx = db.transaction(almacen, "readwrite"), s = tx.objectStore(almacen);
      if (k === undefined) s.put(v); else s.put(v, k);
      tx.oncomplete = () => ok();
      tx.onerror = () => mal(tx.error);
      tx.onabort = () => mal(tx.error || new Error("No se pudo guardar"));
    });
  }
  async function escribir() {
    const bases = new Map();
    let datos = 0, registros = 0, fin = null, m;
    try {
      while ((m = siguiente())) {
        if (m.t === "ls") { localStorage.setItem(m.k, m.v); datos++; }
        else if (m.t === "bd") { estado("Guardando lo de este teléfono…"); bases.set(m.n, await crearBD(m)); }
        else if (m.t === "r") { await poner(bases.get(m.n), m.a, m.k === undefined ? undefined : deJson(m.k), deJson(m.v)); registros++; }
        else if (m.t === "fin") fin = m;
      }
    } finally {
      bases.forEach((db) => db.close());
    }
    if (!fin) throw new Error("Lo leído llegó incompleto");
    if (fin.datos !== datos || fin.registros !== registros) throw new Error(`Llegaron ${datos} de ${fin.datos} datos y ${registros} de ${fin.registros} registros`);
    A.migracionHecha(true, datos + " datos y " + registros + " registros");
  }

  const falla = (e) => {
    estado("No se pudo pasar todo: lo guardado sigue en el teléfono y se vuelve a intentar al abrir la app.");
    try { (location.protocol === "file:" ? A.migracionFin : A.migracionHecha).call(A, false, String((e && e.message) || e).slice(0, 300)); } catch (x) { /* app sin el puente */ }
  };
  (location.protocol === "file:" ? leer() : escribir()).catch(falla);
})();

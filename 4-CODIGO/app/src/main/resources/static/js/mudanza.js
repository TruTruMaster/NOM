/* NOM · la dirección de antes ya no se usa en el navegador: quien abre la página con ese nombre pasa a
   www.nom.qd.je. Lo que el navegador guarda (sesión, ajustes, documentos sin cuenta, copias y cambios sin enviar)
   es de cada dirección, así que antes de cambiar se pasa a la nueva:
     - index.html (aquí, en la dirección de antes): si el navegador no guarda nada de NOM, cambia de dirección al
       momento; si guarda algo, abre mudanza.html para pasarlo.
     - mudanza.html en la dirección de antes: con un clic abre mudanza.html en la nueva y le manda todo por postMessage.
     - mudanza.html en la nueva: recibe y guarda solo lo que no tenga ya (nunca pisa lo de la dirección nueva).
   La API sigue contestando con el nombre de antes (apps ya instaladas y pases): esto solo toca al navegador.
   Para las pruebas, window.NOM_MUDANZA_PRUEBA = { vieja, nueva } con dos orígenes de este mismo equipo. */
(function () {
  "use strict";
  const PRUEBA = window.NOM_MUDANZA_PRUEBA;
  const local = (o) => /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(o || "");
  const ENPRUEBA = PRUEBA && local(PRUEBA.vieja) && local(PRUEBA.nueva);
  const NUEVA = ENPRUEBA ? PRUEBA.nueva : "https://www.nom.qd.je";
  const VIEJAS = ENPRUEBA ? [PRUEBA.vieja] : ["https://nomcolab.dpdns.org", "https://www.nomcolab.dpdns.org"];
  const enVieja = VIEJAS.includes(location.origin);
  const enMudanza = /\/mudanza\.html$/.test(location.pathname);
  const BASES = ["nom-boveda", "nom-binarios"];
  const estado = (t) => { const e = document.getElementById("estado"); if (e) e.textContent = t; };

  // ------------------------------------------------------------ IndexedDB
  const pedido = (r) => new Promise((ok, mal) => { r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error); });
  async function nombresBD() {
    const n = new Set(BASES);
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
  /** ¿Guarda este navegador algún registro de NOM? (los ajustes sueltos no cuentan: se vuelven a elegir) */
  async function hayAlgo() {
    if (!window.indexedDB) return false;
    for (const nombre of await nombresBD()) {
      const db = await abrirSiExiste(nombre).catch(() => null);
      if (!db) continue;
      try {
        for (const a of db.objectStoreNames) if (await pedido(db.transaction(a, "readonly").objectStore(a).count())) return true;
      } finally { db.close(); }
    }
    return false;
  }
  const irANueva = () => location.replace(NUEVA + location.pathname.replace(/\/mudanza\.html$/, "/") + location.search + location.hash);

  // ------------------------------------------------------------ 1. index.html en la dirección de antes
  if (!enMudanza) {
    if (!enVieja) return;
    hayAlgo().then((si) => {
      if (!si) return irANueva();
      location.replace("/mudanza.html#" + encodeURIComponent(location.pathname + location.search + location.hash));
    }).catch(() => location.replace("/mudanza.html"));
    return;
  }

  // ------------------------------------------------------------ 2. mudanza.html en la dirección de antes: mandar
  const boton = document.getElementById("pasar");
  if (enVieja) {
    const destino = (() => { try { const r = decodeURIComponent(location.hash.slice(1)); return r.startsWith("/") && !r.startsWith("//") ? r : "/"; } catch (e) { return "/"; } })();
    const terminar = () => location.replace(NUEVA + destino);
    document.getElementById("saltar").addEventListener("click", terminar);
    boton.addEventListener("click", () => {
      boton.disabled = true;
      estado("Abriendo la dirección nueva…");
      const v = window.open(NUEVA + "/mudanza.html", "nom-mudanza");
      if (!v) { boton.disabled = false; estado("El navegador no dejó abrir la ventana: permite las ventanas emergentes de esta página y vuelve a intentarlo."); return; }
      let mandado = false;
      window.addEventListener("message", async (ev) => {
        if (ev.origin !== NUEVA || ev.source !== v || !ev.data || typeof ev.data !== "object") return;
        if (ev.data.t === "lista" && !mandado) {
          mandado = true;
          try {
            let datos = 0, registros = 0;
            estado("Pasando la sesión y los ajustes…");
            for (let i = 0; i < localStorage.length; i++) {
              const k = localStorage.key(i);
              if (k != null) { v.postMessage({ t: "ls", k, v: localStorage.getItem(k) }, NUEVA); datos++; }
            }
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
              v.postMessage({ t: "bd", n: nombre, ver: db.version, almacenes }, NUEVA);
              for (const al of almacenes) {
                estado(nombre === "nom-binarios" ? "Pasando las imágenes y los archivos…" : "Pasando los documentos guardados…");
                for (const k of await pedido(db.transaction(al.n, "readonly").objectStore(al.n).getAllKeys())) {
                  // uno por uno: con archivos grandes no se junta todo en memoria (postMessage copia blobs y bytes tal cual)
                  const x = await pedido(db.transaction(al.n, "readonly").objectStore(al.n).get(k));
                  if (x === undefined) continue;
                  v.postMessage({ t: "r", n: nombre, a: al.n, k: al.kp == null ? k : undefined, v: x }, NUEVA); registros++;
                }
              }
              db.close();
            }
            v.postMessage({ t: "fin", datos, registros }, NUEVA);
            estado("Guardando en la dirección nueva…");
          } catch (e) {
            boton.disabled = false; mandado = false;
            estado("No se pudo pasar todo (" + String((e && e.message) || e).slice(0, 120) + "). Lo tuyo sigue aquí: vuelve a intentarlo.");
          }
        } else if (ev.data.t === "hecho") {
          estado("Listo. Cambiando a la dirección nueva…");
          try { v.close(); } catch (e) { /* la cierra ella */ }
          terminar();
        } else if (ev.data.t === "error") {
          boton.disabled = false; mandado = false;
          estado("No se pudo guardar en la dirección nueva (" + String(ev.data.m || "").slice(0, 120) + "). Lo tuyo sigue aquí: vuelve a intentarlo.");
        }
      });
    });
    return;
  }

  // ------------------------------------------------------------ 3. mudanza.html en la dirección nueva: recibir
  document.getElementById("vieja").hidden = true;
  const madre = window.opener;
  if (!madre) { estado("Esta página solo sirve para pasar lo guardado desde la dirección anterior de NOM."); return; }
  estado("Recibiendo lo guardado…");
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
  /** Agrega el registro solo si no existe ya (lo de la dirección nueva manda). */
  function agregar(db, almacen, k, v) {
    return new Promise((ok, mal) => {
      if (!db.objectStoreNames.contains(almacen)) return ok();
      const tx = db.transaction(almacen, "readwrite"), s = tx.objectStore(almacen);
      const r = k === undefined ? s.add(v) : s.add(v, k);
      r.onerror = (e) => { if (r.error && r.error.name === "ConstraintError") { e.preventDefault(); e.stopPropagation(); } };
      tx.oncomplete = () => ok();
      tx.onerror = () => { if (tx.error && tx.error.name === "ConstraintError") ok(); else mal(tx.error); };
      tx.onabort = () => { if (tx.error && tx.error.name === "ConstraintError") ok(); else mal(tx.error || new Error("No se pudo guardar")); };
    });
  }
  const bases = new Map();
  let cola = Promise.resolve(), datos = 0, registros = 0, fallo = null;
  window.addEventListener("message", (ev) => {
    if (!VIEJAS.includes(ev.origin) || ev.source !== madre || !ev.data || typeof ev.data !== "object") return;
    const m = ev.data, origen = ev.origin;
    cola = cola.then(async () => {
      if (fallo) return;
      try {
        if (m.t === "ls" && typeof m.k === "string" && typeof m.v === "string") {
          if (localStorage.getItem(m.k) === null) localStorage.setItem(m.k, m.v);
          datos++;
        } else if (m.t === "bd" && typeof m.n === "string" && Array.isArray(m.almacenes)) {
          if (!bases.has(m.n)) bases.set(m.n, await crearBD(m));
        } else if (m.t === "r" && bases.has(m.n)) {
          await agregar(bases.get(m.n), m.a, m.k, m.v); registros++;
        } else if (m.t === "fin") {
          bases.forEach((db) => db.close());
          if (m.datos !== datos || m.registros !== registros) throw new Error(`llegaron ${datos} de ${m.datos} datos y ${registros} de ${m.registros} registros`);
          estado("Listo.");
          madre.postMessage({ t: "hecho", datos, registros }, origen);
        }
      } catch (e) {
        fallo = e;
        madre.postMessage({ t: "error", m: String((e && e.message) || e).slice(0, 200) }, origen);
        estado("No se pudo guardar todo: lo de la dirección anterior sigue allá.");
      }
    });
  });
  for (const o of VIEJAS) madre.postMessage({ t: "lista" }, o);
})();

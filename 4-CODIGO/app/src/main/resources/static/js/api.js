/* NOM · cliente de la API del nodo local.
   Desde el navegador la API está en el mismo origen. Dentro de la app Android
   (file://) la dirección del servidor se elige en la pantalla de inicio. */
(function () {
  "use strict";
  const NOM = window.NOM;

  const empaquetada = location.protocol === "file:" || !!window.NomAndroid;
  // La nube tiene varios nombres (www.nom.qd.je, nom.qd.je y el de antes, nomcolab.dpdns.org): es el
  // mismo servidor, con las mismas cuentas. Para la app son un solo lugar y se guarda el principal; así
  // un QR hecho con el nombre de antes no la cambia de servidor ni cierra la sesión.
  const NUBE_PRINCIPAL = "https://www.nom.qd.je";
  const NOMBRES_NUBE = /^https?:\/\/((www\.)?nom\.qd\.je|(www\.)?nomcolab\.dpdns\.org)(:443)?\/?$/i;
  const normal = (u) => (NOMBRES_NUBE.test(String(u || "")) ? NUBE_PRINCIPAL : u);
  if (empaquetada) {
    // (una vez: lo guardado con el nombre de antes pasa al principal)
    try {
      const s = NOM.leer("nom.servidor", "");
      if (s && normal(s) !== s) NOM.guardar("nom.servidor", normal(s));
      ["nom.sesionNube", "nom.infoServidor"].forEach((k) => { const g = NOM.leer(k, null); if (g && g.servidor && normal(g.servidor) !== g.servidor) NOM.guardar(k, Object.assign(g, { servidor: normal(g.servidor) })); });
    } catch (e) { /* sin almacenamiento */ }
  }
  const api = {
    empaquetada,
    NUBE_PRINCIPAL,
    /** ¿Son el mismo servidor? (los nombres de la nube cuentan como uno) */
    mismoServidor(a, b) { return normal(String(a || "").replace(/\/+$/, "")) === normal(String(b || "").replace(/\/+$/, "")); },
    esNube(u) { return NOMBRES_NUBE.test(String(u || "")); },
    servidor() { return empaquetada ? (normal(NOM.leer("nom.servidor", "") || "") || "") : ""; },
    fijarServidor(url) {
      url = String(url || "").trim().replace(/\/+$/, "");
      if (url && !/^https?:\/\//i.test(url)) {
        // Un equipo de la red (IP, localhost, nombre.local) va por http y el puerto 8080;
        // un dominio de Internet (p. ej. la versión publicada en Azure) va por https.
        const host = url.split(/[/:]/)[0].toLowerCase();
        const enLaRed = /^\d+(\.\d+){3}$/.test(host) || host === "localhost" || !host.includes(".") || host.endsWith(".local");
        url = (enLaRed ? "http://" : "https://") + url;
        if (enLaRed && !/:\d+$/.test(url.replace(/^https?:\/\//, ""))) url += ":8080";
      }
      url = normal(url);
      NOM.guardar("nom.servidor", url);
      return url;
    },
    token() { return NOM.leer("nom.token", null); },
    usuario() { return NOM.leer("nom.usuario", null); },
    fijarSesion(token, usuario) {
      NOM.guardar("nom.token", token); NOM.guardar("nom.usuario", usuario);
      if (NOM.boveda) NOM.boveda.recordarSesion(token, usuario);   // copia que sobrevive a un cierre de golpe
    },
    limpiarSesion() {
      NOM.quitar("nom.token"); NOM.quitar("nom.usuario");
      if (NOM.boveda) NOM.boveda.soltarSesion();
    },
    /** Clave de localStorage de la cuenta actual: las copias y los cambios pendientes de una cuenta
        nunca se mezclan con los de otra que use el mismo dispositivo. */
    clave(resto) { return "nom.u." + ((api.usuario() || {}).id || "anon") + "." + resto; },
    /** Datos guardados antes de separar por cuenta: pasan a la primera cuenta que entra. */
    migrarClaves() {
      try {
        Object.keys(localStorage).forEach((k) => {
          const m = /^nom\.(doc\.|cola\.|recientes$|cache\.)(.*)$/.exec(k);
          if (!m) return;
          // Pasan a la bóveda cifrada de la cuenta y desaparecen de localStorage.
          const nueva = api.clave(m[1] + m[2]);
          if (NOM.leer(nueva, undefined) === undefined) { try { NOM.guardar(nueva, JSON.parse(localStorage.getItem(k))); } catch (e) { /* dato dañado */ } }
          localStorage.removeItem(k);
        });
      } catch (e) { /* sin almacenamiento */ }
    },
    enLinea: navigator.onLine !== false,
    /** Cómo se llama este dispositivo en la bitácora de cambios. */
    dispositivo: empaquetada ? "App Android"
      : /Android|iPhone|iPad|Mobile/i.test(navigator.userAgent || "") || (window.matchMedia && matchMedia("(pointer: coarse) and (max-width: 860px)").matches)
        ? "Navegador del celular" : "Navegador de computadora",
    /** Identificador de ESTE dispositivo, solo para el tope de cuentas nuevas por hora (no es la MAC:
        ningún navegador ni app puede leerla). En la app es el ANDROID_ID del teléfono resumido (sobrevive
        a reinstalar); en el navegador, un número al azar que se guarda en él. */
    equipo() {
      try { if (window.NomAndroid && window.NomAndroid.equipo) { const x = String(window.NomAndroid.equipo() || ""); if (/^[A-Za-z0-9_-]{16,64}$/.test(x)) return x; } } catch (e) { /* app vieja */ }
      let q = NOM.leer("nom.equipo", "");
      if (!/^[A-Za-z0-9_-]{16,64}$/.test(q || "")) {
        const b = new Uint8Array(16);
        crypto.getRandomValues(b);
        q = btoa(String.fromCharCode.apply(null, b)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
        NOM.guardar("nom.equipo", q);
      }
      return q;
    },
    /** false si el servidor está en la nube: no hay "red local" que anunciar ni buscar. */
    lan: NOM.leer("nom.lan", true) !== false,
  };

  class ErrorApi extends Error {
    constructor(estado, cuerpo) {
      super((cuerpo && cuerpo.error) || (estado === 0 ? "Sin conexión: revisa el Wi-Fi o los datos" : "Algo falló (error " + estado + "). Intenta de nuevo."));
      this.estado = estado;
      this.campo = cuerpo && cuerpo.campo;
      this.detalle = cuerpo && cuerpo.detalle;
      this.red = estado === 0;
    }
  }
  api.ErrorApi = ErrorApi;

  /** Desde la 0.19 la app tiene su propio origen (https://appassets.androidplatform.net). Un equipo con una versión
      anterior (otro teléfono o una computadora que comparte) no lo conoce y el navegador no deja leer sus respuestas:
      se vería como «sin conexión». Se distingue pidiéndole algo sin leer la respuesta (no-cors): si contesta, está ahí
      y es de antes. Se revisa una vez por minuto por equipo. */
  const VIEJO = "El equipo que comparte tiene una versión anterior de NOM: actualícenla ahí para entrar desde la app (desde el navegador sí se puede).";
  const revisados = new Map();
  async function equipoViejo(base, e) {
    if (!empaquetada || location.protocol !== "https:" || !base || (e && e.name === "AbortError")) return false;
    const r = revisados.get(base);
    if (r && Date.now() - r.t < 60000) return r.viejo;
    let viejo = false;
    const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), 3000);
    try { await fetch(base + "/api/version", { mode: "no-cors", cache: "no-store", signal: ctl.signal }); viejo = true; }
    catch (x) { viejo = false; }
    finally { clearTimeout(t); }
    revisados.set(base, { t: Date.now(), viejo });
    return viejo;
  }
  api.equipoViejo = equipoViejo;

  /** Petición JSON con tiempo límite. Lanza ErrorApi (estado 0 = sin red). */
  api.pedir = async function (metodo, ruta, cuerpo, opc) {
    opc = opc || {};
    const base = api.servidor();
    if (empaquetada && !base) throw new ErrorApi(0, { error: "Primero elige dónde trabaja tu equipo" });
    const cab = { Accept: "application/json" };
    if (cuerpo !== undefined) cab["Content-Type"] = "application/json";
    // Fase 15: en un espacio cifrado de extremo a extremo, el contenido se cifra aquí antes de salir.
    const original = cuerpo;
    if (cuerpo !== undefined && !opc.crudo && NOM.e2e) {
      try { cuerpo = await NOM.e2e.alEnviar(metodo, ruta, cuerpo); }
      catch (e) { throw new ErrorApi(e.estado === 0 ? 0 : 400, { error: e.message }); }
    }
    const tok = api.token();
    if (tok) cab.Authorization = "Bearer " + tok;
    // (las llaves «_…» del cuerpo son de este equipo, como _llaveDoc: no salen. Las de DENTRO de los datos sí: antes se
    // quitaban todas y un documento perdía lo que la importación guarda así, como el color «_c» de una animación.)
    // Se arma antes de pedir: si algo no se puede convertir a JSON es un error de aquí, no «sin conexión».
    const enviar = cuerpo;
    const texto0 = enviar === undefined ? undefined : JSON.stringify(enviar, function (k, v) { return this === enviar && k.startsWith("_") ? undefined : v; });
    // (el tiempo límite cuenta desde que sale la petición: cifrar y armar un documento grande no se lo come)
    const ctl = new AbortController();
    const t = setTimeout(() => ctl.abort(), opc.ms || 9000);
    let r;
    try {
      r = await fetch(base + ruta, { method: metodo, headers: cab, body: texto0, signal: ctl.signal });
    } catch (e) {
      marcarRed(false);
      throw new ErrorApi(0, (await equipoViejo(base, e)) ? { error: VIEJO } : null);
    } finally {
      clearTimeout(t);
    }
    marcarRed(true);
    let datos = null;
    const texto = await r.text();
    if (texto) { try { datos = JSON.parse(texto); } catch (e) { datos = null; } }
    if (!r.ok) {
      const err = new ErrorApi(r.status, datos);
      if (r.status === 401 && !opc.sin401) api.alExpirar && api.alExpirar(err);
      throw err;
    }
    // Fase 15: lo cifrado de extremo a extremo se descifra al llegar.
    if (datos && !opc.crudo && NOM.e2e) {
      try { datos = await NOM.e2e.alRecibir(metodo, ruta, datos, cuerpo); }
      catch (e) { throw new ErrorApi(e.sinLlave ? 423 : 400, { error: e.message }); }
    }
    void original;
    return datos;
  };
  ["get", "post", "put", "delete"].forEach((m) => {
    api[m] = (ruta, cuerpo, opc) => api.pedir(m.toUpperCase(), ruta, m === "get" || m === "delete" ? undefined : (cuerpo === undefined ? {} : cuerpo), opc);
  });
  /**
   * Bytes crudos (el almacén de binarios): PUT manda `bytes` tal cual y devuelve el JSON de la respuesta; GET devuelve
   * los bytes (Uint8Array). Sin cifrado de extremo a extremo aquí: lo de un espacio cifrado ya viene cifrado (binarios.js).
   */
  api.binario = async function (metodo, ruta, bytes, opc) {
    opc = opc || {};
    const base = api.servidor();
    if (empaquetada && !base) throw new ErrorApi(0, { error: "Primero elige dónde trabaja tu equipo" });
    const ctl = new AbortController();
    const t = setTimeout(() => ctl.abort(), opc.ms || 180000);
    const cab = {};
    if (bytes) cab["Content-Type"] = "application/octet-stream";
    const tok = opc.sinSesion ? null : api.token();
    if (tok) cab.Authorization = "Bearer " + tok;
    let r;
    try { r = await fetch(base + ruta, { method: metodo, headers: cab, body: bytes || undefined, signal: ctl.signal }); }
    catch (e) { marcarRed(false); throw new ErrorApi(0, (await equipoViejo(base, e)) ? { error: VIEJO } : null); }
    finally { clearTimeout(t); }
    marcarRed(true);
    if (!r.ok) {
      let datos = null;
      try { datos = JSON.parse(await r.text()); } catch (e) { datos = null; }
      throw new ErrorApi(r.status, datos);
    }
    if (metodo === "GET") return new Uint8Array(await r.arrayBuffer());
    try { return JSON.parse(await r.text()); } catch (e) { return null; }
  };

  /**
   * Canal en vivo (EventSource). No admite cabeceras, así que se identifica en la dirección; para
   * que ahí no viaje el token de la sesión (las direcciones acaban en bitácoras), antes de cada
   * conexión se pide un boleto de un solo uso. Un servidor de una versión anterior no los da
   * (responde 404): con ese se sigue mandando el token, y se recuerda para no volver a preguntar.
   */
  const sinBoletos = new Set();
  async function direccionCanal(ruta, cliente) {
    const base = api.servidor(), fin = "&cliente=" + encodeURIComponent(cliente || "");
    // La pantalla del proyector se abre con su código, sin sesión: no lleva boleto ni token.
    if (!api.token() || ruta.startsWith("/api/proyeccion/")) return base + ruta + "?" + fin.slice(1);
    if (!sinBoletos.has(base)) {
      try {
        const r = await api.post("/api/cuentas/boleto", {}, { ms: 6000, sin401: true });
        if (r && r.boleto) return base + ruta + "?boleto=" + encodeURIComponent(r.boleto) + fin;
      } catch (e) {
        if (e.estado === 404 || e.estado === 405) sinBoletos.add(base);
        else throw e;
      }
    }
    return base + ruta + "?token=" + encodeURIComponent(api.token() || "") + fin;
  }

  api.eventos = function (ruta, cliente, manejadores) {
    let es, cerrado = false, reintento = 1000, timer, abierto = 0;
    const reintentar = () => { timer = setTimeout(abrir, reintento); reintento = Math.min(reintento * 2, 4000); };
    const abrir = async () => {
      let url;
      try { url = await direccionCanal(ruta, cliente); }
      catch (e) {
        if (cerrado) return;
        // Sin boleto no hay canal: sesión vencida (se avisa) o red caída (se reintenta).
        if (e.estado === 401) { cerrado = true; api.alExpirar && api.alExpirar(e); return; }
        manejadores.caido && manejadores.caido();
        return reintentar();
      }
      if (cerrado) return;
      es = new EventSource(url);
      es.addEventListener("hola", (e) => { reintento = 1000; abierto = Date.now(); marcarRed(true); manejadores.hola && manejadores.hola(JSON.parse(e.data)); });
      Object.keys(manejadores).forEach((ev) => {
        if (ev === "hola" || ev === "caido") return;
        es.addEventListener(ev, (e) => { try { manejadores[ev](JSON.parse(e.data)); } catch (x) { console.error(x); } });
      });
      es.onerror = () => {
        if (cerrado) return;
        es.close();
        // Un canal que llevaba un buen rato abierto y se corta suele ser el reenvío del nombre corto
        // (Vercel corta toda conexión a los 2 minutos): se vuelve a abrir de inmediato, sin avisar.
        // Si de verdad no hay red, el siguiente intento falla y entonces sí se avisa.
        if (abierto && Date.now() - abierto > 90000) { abierto = 0; reintento = 1000; return abrir(); }
        abierto = 0;
        manejadores.caido && manejadores.caido();
        // Puede ser red caída, sesión vencida o espacio lleno: se pregunta al servidor qué pasó.
        sondear().finally(() => { if (!cerrado) reintentar(); });
      };
    };
    async function sondear() {
      // Se pide el mismo canal con fetch solo para leer el código de estado y se corta. Aquí la
      // sesión sí va en la cabecera (fetch puede mandarla): el boleto ya se gastó.
      const ctl = new AbortController();
      const t = setTimeout(() => ctl.abort(), 4000);
      try {
        const r = await fetch(api.servidor() + ruta + "?cliente=" + encodeURIComponent(cliente || ""), { headers: Object.assign({ Accept: "text/event-stream" }, api.token() ? { Authorization: "Bearer " + api.token() } : {}), signal: ctl.signal });
        if (!r.ok) {
          let cuerpo = null;
          try { cuerpo = await r.json(); } catch (x) { /* sin cuerpo */ }
          const e = new ErrorApi(r.status, cuerpo);
          if (r.status === 401) { cerrado = true; api.alExpirar && api.alExpirar(e); }
          else if (r.status === 409 || r.status === 403 || r.status === 404) { cerrado = true; manejadores.rechazo && manejadores.rechazo(e); }
        }
      } catch (e) {
        marcarRed(false);
      } finally {
        clearTimeout(t);
        ctl.abort();
      }
    }
    abrir();
    return { cerrar() { cerrado = true; clearTimeout(timer); es && es.close(); } };
  };

  // ------------------------------------------------------------ estado de la red
  const oyentesRed = new Set();
  function marcarRed(ok) {
    if (api.enLinea === ok) return;
    api.enLinea = ok;
    sondear();
    oyentesRed.forEach((f) => f(ok));
  }
  // Algunos WebView no avisan cuando vuelve la red (en el A55, con WebView 157, navigator.onLine se queda en true con
  // el modo avión y no llega «online»): lo que estaba pendiente no salía hasta que la persona hiciera algo. Mientras la
  // última petición haya fallado por red, se le pregunta al servidor cada rato (4 s, y más espaciado hasta 15 s).
  let sondeo = null, espera = 4000;
  function sondear() {
    clearTimeout(sondeo);
    if (api.enLinea !== false) { espera = 4000; return; }
    sondeo = setTimeout(async () => {
      if (api.enLinea !== false) return;
      try { await api.get("/api/red/info", undefined, { ms: 4000 }); } catch (e) { /* sigue sin red */ }
      espera = Math.min(15000, Math.round(espera * 1.5));
      sondear();
    }, espera);
  }
  api.alCambiarRed = (f) => { oyentesRed.add(f); return () => oyentesRed.delete(f); };
  window.addEventListener("online", () => api.get("/api/red/info", undefined, { ms: 4000 }).catch(() => {}));
  window.addEventListener("offline", () => marcarRed(false));

  NOM.api = api;
})();

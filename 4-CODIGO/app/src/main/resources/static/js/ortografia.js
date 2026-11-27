/* NOM · Corrector ortográfico bajo demanda (nspell + diccionarios Hunspell de español e inglés).
   - Baja el motor (js/vendor/ortografia/nspell-navegador.js) y el diccionario (.aff y .dic)
     solo la primera vez que se usan, y los conserva en memoria.
   - Si no hay red o falla la descarga, devuelve null sin lanzar errores.
   - Guarda el diccionario personal en localStorage bajo "nom.ortografia.personal".
   - El idioma sale del texto del documento (NOM.idioma.detectar) o del idioma de la interfaz. */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});

  const RUTA_BASE = "js/vendor/ortografia/";
  const RUTA_MOTOR = RUTA_BASE + "nspell-navegador.js";
  const CLAVE_PERSONAL = "nom.ortografia.personal";
  const MAX_SUGERENCIAS = 5;
  const MAX_LARGO_SUG = 48;
  const SIN_CONEXION = "Sin conexión: el diccionario no se pudo bajar";

  const diccionarios = { es: null, en: null };
  const cargando = { es: null, en: null };
  let cargandoMotor = null;

  let personalRaw = undefined;
  let personalSet = new Set();

  // ================================================================== idioma
  function idiomaInterfaz() {
    const a = NOM.idioma && typeof NOM.idioma.actual === "function" ? NOM.idioma.actual() : "es";
    return String(a || "").trim().toLowerCase().startsWith("en") ? "en" : "es";
  }

  function normalizarIdioma(idioma) {
    if (idioma != null && String(idioma).trim()) {
      const s = String(idioma).trim().toLowerCase();
      if (s.startsWith("en")) return "en";
      if (s.startsWith("es")) return "es";
    }
    return idiomaInterfaz();
  }

  /** Idioma para revisar: el del documento (NOM.idioma.detectar) o el de la interfaz. */
  function idiomaPara(texto) {
    const def = idiomaInterfaz();
    const s = String(texto || "").trim();
    if (s && NOM.idioma && typeof NOM.idioma.detectar === "function") {
      return normalizarIdioma(NOM.idioma.detectar(s, def));
    }
    return def;
  }

  // ================================================================== diccionario personal
  function conjuntoPersonal() {
    let raw = null;
    try {
      raw = typeof localStorage !== "undefined" && localStorage ? localStorage.getItem(CLAVE_PERSONAL) : null;
    } catch (e) { raw = null; }
    if (raw === personalRaw) return personalSet;
    personalRaw = raw;
    personalSet = new Set();
    if (raw) {
      try {
        const arr = JSON.parse(raw);
        if (Array.isArray(arr)) {
          for (let i = 0; i < arr.length; i++) {
            const w = String(arr[i] || "").trim();
            if (w) personalSet.add(w);
          }
        }
      } catch (e) { /* lista dañada: se ignora */ }
    }
    return personalSet;
  }

  function guardarPersonal(lista) {
    try {
      if (typeof localStorage !== "undefined" && localStorage) {
        localStorage.setItem(CLAVE_PERSONAL, JSON.stringify(lista));
      } else if (NOM.guardar) {
        NOM.guardar(CLAVE_PERSONAL, lista);
      }
      personalRaw = JSON.stringify(lista);
      personalSet = new Set(lista);
      return true;
    } catch (e) {
      return false;
    }
  }

  function agregar(palabra) {
    const w = String(palabra || "").trim();
    if (!w) return false;
    const set = conjuntoPersonal();
    if (!set.has(w)) {
      const lista = Array.from(set);
      lista.push(w);
      guardarPersonal(lista);
    }
    if (diccionarios.es) diccionarios.es.add(w);
    if (diccionarios.en) diccionarios.en.add(w);
    return true;
  }

  // ================================================================== carga perezosa
  function motorDisponible() {
    if (typeof window !== "undefined" && typeof window.NSpell === "function") return window.NSpell;
    if (typeof globalThis !== "undefined" && typeof globalThis.NSpell === "function") return globalThis.NSpell;
    return null;
  }

  function cargarMotor() {
    const ya = motorDisponible();
    if (ya) return Promise.resolve(ya);
    if (cargandoMotor) return cargandoMotor;
    if (typeof document === "undefined" || !document.createElement) return Promise.resolve(null);
    cargandoMotor = new Promise((resolver) => {
      try {
        const s = document.createElement("script");
        s.onload = () => {
          const m = motorDisponible();
          if (!m) cargandoMotor = null;
          resolver(m);
        };
        s.onerror = () => {
          cargandoMotor = null;
          resolver(null);
        };
        s.src = RUTA_MOTOR;
        (document.head || document.documentElement).append(s);
      } catch (e) {
        cargandoMotor = null;
        resolver(null);
      }
    });
    return cargandoMotor;
  }

  /** Carga el diccionario de `idioma` ("es" o "en") la primera vez; luego lo devuelve de memoria. */
  function cargar(idioma) {
    const cod = normalizarIdioma(idioma);
    if (diccionarios[cod]) return Promise.resolve(diccionarios[cod]);
    if (cargando[cod]) return cargando[cod];
    cargando[cod] = (async () => {
      try {
        const Motor = await cargarMotor();
        if (!Motor || typeof fetch !== "function") {
          cargando[cod] = null;
          return null;
        }
        const prefijo = RUTA_BASE + cod + "/" + cod;
        const [rAff, rDic] = await Promise.all([fetch(prefijo + ".aff"), fetch(prefijo + ".dic")]);
        if (!rAff || !rAff.ok || !rDic || !rDic.ok) {
          cargando[cod] = null;
          return null;
        }
        const [aff, dic] = await Promise.all([rAff.text(), rDic.text()]);
        if (!aff || !dic) {
          cargando[cod] = null;
          return null;
        }
        const sp = Motor(aff, dic);
        const pers = conjuntoPersonal();
        pers.forEach((w) => sp.add(w));
        diccionarios[cod] = sp;
        cargando[cod] = null;
        return sp;
      } catch (e) {
        cargando[cod] = null;
        return null;
      }
    })();
    return cargando[cod];
  }

  function listo(idioma) {
    return !!diccionarios[normalizarIdioma(idioma)];
  }

  // ================================================================== revisión y sugerencias
  /** Devuelve true si está bien, false si está mal, o null si el diccionario aún no se carga. */
  function revisar(palabra, idioma) {
    const cod = normalizarIdioma(idioma);
    const sp = diccionarios[cod];
    if (!sp) return null;
    const w = String(palabra || "").trim();
    if (!w || !/\p{L}/u.test(w)) return true;
    const pers = conjuntoPersonal();
    if (pers.has(w) || pers.has(w.toLowerCase())) return true;
    return Boolean(sp.correct(w));
  }

  /** Devuelve hasta 5 sugerencias para una palabra mal escrita ([] si está bien o sin diccionario). */
  function sugerencias(palabra, idioma) {
    const cod = normalizarIdioma(idioma);
    const sp = diccionarios[cod];
    if (!sp) return [];
    const w = String(palabra || "").trim();
    if (!w || w.length > MAX_LARGO_SUG || !/\p{L}/u.test(w)) return [];
    if (revisar(w, cod) !== false) return [];
    const lista = sp.suggest(w) || [];
    return lista.slice(0, MAX_SUGERENCIAS);
  }

  NOM.ortografia = {
    cargar,
    listo,
    revisar,
    sugerencias,
    agregar,
    personal: () => Array.from(conjuntoPersonal()),
    idioma: idiomaPara,
    SIN_CONEXION,
  };
})();

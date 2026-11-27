/* NOM · Preferencias de este equipo: atajos de teclado personalizables, vibración al tocar (háptica)
   y elegir entre la ruedita o una lista clásica al mantener presionado o hacer clic derecho.
   Se guardan en localStorage bajo una clave propia ("nom.pref.extra") para no chocar con el modelo
   fijo de preferencias de la cuenta en el servidor. */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});
  const propio = (o, k) => !!o && Object.prototype.hasOwnProperty.call(o, k);

  const CLAVE = "nom.pref.extra";
  const NIVELES_HAPTICA = { suave: true, normal: true, fuerte: true };
  const MODOS_MENU = { rueda: true, lista: true };

  // Duraciones cortas (en milisegundos) por intensidad y tipo de toque.
  const HAPTICA_MS = {
    suave:  { abrir: 8,  elegir: 12, mover: 4,  toque: 6  },
    normal: { abrir: 16, elegir: 22, mover: 8,  toque: 12 },
    fuerte: { abrir: 28, elegir: 38, mover: 14, toque: 22 },
  };

  // Tabla de atajos personalizables: acción, nombre visible y tecla de siempre.
  const ATAJOS = Object.freeze([
    { accion: "paleta",        nombre: "Buscar o ir a",                        tecla: "Ctrl+K",      grupo: "En toda la app",  soloFueraCampo: true },
    { accion: "paletaTexto",   nombre: "Buscar o ir a (mientras escribes)",    tecla: "Alt+/",       grupo: "En toda la app",  enCampo: true },
    { accion: "atajos",        nombre: "Ver la lista de atajos",               tecla: "?",           grupo: "En toda la app",  soloFueraCampo: true, sinEditor: true },
    { accion: "ajustes",       nombre: "Abrir Ajustes y preferencias",         tecla: "Ctrl+,",      grupo: "En toda la app",  enCampo: true },
    { accion: "inicio",        nombre: "Ir a Inicio",                          tecla: "Alt+H",       grupo: "En toda la app",  soloFueraCampo: true },
    { accion: "nuevoDoc",      nombre: "Crear un documento nuevo",             tecla: "Alt+N",       grupo: "En toda la app",  soloFueraCampo: true },
    { accion: "tema",          nombre: "Cambiar entre claro y oscuro",         tecla: "Alt+T",       grupo: "En toda la app",  soloFueraCampo: true },
    { accion: "guardar",       nombre: "Guardar versión del documento",        tecla: "Ctrl+S",      grupo: "En el documento", enEditor: true, enCampo: true },
    { accion: "enfoque",       nombre: "Modo enfoque (sin distracciones)",     tecla: "Ctrl+.",      grupo: "En el documento", enEditor: true, enCampo: true },
    { accion: "accesibilidad", nombre: "Comprobar accesibilidad",              tecla: "Alt+Shift+A", grupo: "En el documento", enEditor: true, enCampo: true },
  ]);
  const POR_ACCION = new Map(ATAJOS.map((a) => [a.accion, a]));

  // Teclas que el navegador o el sistema no dejan cambiar (ni las básicas de teclado).
  const PROHIBIDOS = {
    "Ctrl+W": "Ctrl + W cierra la pestaña y el navegador no deja cambiarlo.",
    "Ctrl+Shift+W": "Ctrl + Mayús + W cierra la ventana y el navegador no deja cambiarlo.",
    "Ctrl+T": "Ctrl + T abre una pestaña nueva y el navegador no deja cambiarlo.",
    "Ctrl+Shift+T": "Ctrl + Mayús + T reabre una pestaña y el navegador no deja cambiarlo.",
    "Ctrl+N": "Ctrl + N abre una ventana nueva y el navegador no deja cambiarlo.",
    "Ctrl+Shift+N": "Ctrl + Mayús + N abre una ventana privada y el navegador no deja cambiarlo.",
    "Ctrl+Tab": "Ctrl + Tab cambia de pestaña y el navegador no deja cambiarlo.",
    "Ctrl+Shift+Tab": "Ctrl + Mayús + Tab cambia de pestaña y el navegador no deja cambiarlo.",
    "F5": "F5 recarga la página (fuera de presentar, el navegador no deja cambiarlo).",
    "Ctrl+F5": "Ctrl + F5 recarga la página y el navegador no deja cambiarlo.",
    "Alt+F4": "Alt + F4 cierra la ventana y el sistema no deja cambiarlo.",
    "Escape": "Esc siempre sirve para cerrar ventanas y menús.",
    "Tab": "Tab sirve para moverse entre botones con el teclado.",
    "Shift+Tab": "Mayús + Tab sirve para moverse entre botones con el teclado.",
    "Enter": "Enter sirve para aceptar o bajar de renglón.",
  };

  // Símbolos que ya llevan Mayús en el teclado (como ?): no se duplica "Shift+" delante.
  const SIMBOLOS_MAYUS = new Set(["?", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "{", "}", "|", ":", '"', "<", ">", "~", "¡", "¿"]);
  const MODS_SOLOS = new Set(["Control", "Shift", "Alt", "Meta", "AltGraph", "CapsLock", "NumLock", "ScrollLock", "Dead", "Process", "Unidentified", ""]);
  const NOMBRES_TECLA = {
    Shift: "Mayús",
    Space: "Espacio",
    Delete: "Supr",
    Backspace: "Retroceso",
    Escape: "Esc",
    ArrowUp: "↑",
    ArrowDown: "↓",
    ArrowLeft: "←",
    ArrowRight: "→",
  };
  const ALIAS_TECLA = {
    esc: "Escape",
    escape: "Escape",
    supr: "Delete",
    del: "Delete",
    delete: "Delete",
    retroceso: "Backspace",
    backspace: "Backspace",
    espacio: "Space",
    space: "Space",
    spacebar: "Space",
    " ": "Space",
    tab: "Tab",
    enter: "Enter",
    return: "Enter",
    home: "Home",
    end: "End",
    pageup: "PageUp",
    pagedown: "PageDown",
    insert: "Insert",
    arrowup: "ArrowUp",
    arrowdown: "ArrowDown",
    arrowleft: "ArrowLeft",
    arrowright: "ArrowRight",
  };

  // ================================================================== lectura y guardado
  function fabrica() {
    return { haptica: true, hapticaNivel: "normal", menuRapido: "rueda", atajos: {} };
  }

  function limpiar(raw) {
    const base = fabrica();
    if (!raw || typeof raw !== "object") return base;
    if (typeof raw.haptica === "boolean") base.haptica = raw.haptica;
    if (propio(NIVELES_HAPTICA, raw.hapticaNivel)) base.hapticaNivel = raw.hapticaNivel;
    if (propio(MODOS_MENU, raw.menuRapido)) base.menuRapido = raw.menuRapido;
    if (raw.atajos && typeof raw.atajos === "object") {
      const limpio = {};
      for (const item of ATAJOS) {
        if (!propio(raw.atajos, item.accion)) continue;
        const v = raw.atajos[item.accion];
        if (v === "") { limpio[item.accion] = ""; continue; }
        if (typeof v === "string") {
          const n = normalizarTecla(v);
          if (n && !propio(PROHIBIDOS, n) && n !== item.tecla) limpio[item.accion] = n;
        }
      }
      base.atajos = limpio;
    }
    return base;
  }

  function leer() {
    const guardado = NOM.leer ? NOM.leer(CLAVE, null) : null;
    return limpiar(guardado);
  }

  function guardar(cambios) {
    const actual = leer();
    const mezclado = Object.assign({}, actual, cambios || {});
    if (cambios && cambios.atajos && typeof cambios.atajos === "object") mezclado.atajos = Object.assign({}, cambios.atajos);
    const limpio = limpiar(mezclado);
    if (NOM.guardar) NOM.guardar(CLAVE, limpio);
    else { try { localStorage.setItem(CLAVE, JSON.stringify(limpio)); } catch (e) { /* sin almacenamiento */ } }
    return limpio;
  }

  function restablecer() {
    return guardar(fabrica());
  }

  function clonar(estado) {
    const e = limpiar(estado || leer());
    return { haptica: e.haptica, hapticaNivel: e.hapticaNivel, menuRapido: e.menuRapido, atajos: Object.assign({}, e.atajos) };
  }

  // ================================================================== atajos: normalización y validación
  function normalizarNombreTecla(k) {
    if (typeof k !== "string") return "";
    if (k === " ") return "Space";
    const t = k.trim();
    if (!t) return "";
    if (/^[a-z]$/i.test(t)) return t.toUpperCase();
    if (/^f([1-9]|1[0-2])$/i.test(t)) return t.toUpperCase();
    const min = t.toLowerCase();
    if (propio(ALIAS_TECLA, min)) return ALIAS_TECLA[min];
    return t;
  }

  /** Convierte un evento de teclado (keydown) en una tecla canónica ("Ctrl+K", "Alt+/", "?") o null si solo es un modificador. */
  function teclaDeEvento(e) {
    if (!e || typeof e.key !== "string") return null;
    if (MODS_SOLOS.has(e.key)) return null;
    const ctrl = !!(e.ctrlKey || e.metaKey);
    const alt = !!e.altKey;
    const k = normalizarNombreTecla(e.key);
    if (!k) return null;
    const shift = !!e.shiftKey && !SIMBOLOS_MAYUS.has(k);
    const partes = [];
    if (ctrl) partes.push("Ctrl");
    if (alt) partes.push("Alt");
    if (shift) partes.push("Shift");
    partes.push(k);
    return partes.join("+");
  }

  /** Descompone una combinación canónica respetando la tecla "+" (p. ej. "Ctrl++" -> { ctrl: true, ..., key: "+" }). */
  function descomponer(combo) {
    if (typeof combo !== "string") return null;
    const s = combo.trim();
    if (!s) return { ctrl: false, alt: false, shift: false, key: "" };
    let resto = s;
    let ctrl = false, alt = false, shift = false;
    while (resto.length > 1) {
      const m = /^(Ctrl|Control|Cmd|Meta|Alt|Shift|Mayús|Mayus)\s*\+\s*(.+)$/i.exec(resto);
      if (!m) break;
      const mod = m[1].toLowerCase();
      if (mod === "ctrl" || mod === "control" || mod === "cmd" || mod === "meta") ctrl = true;
      else if (mod === "alt") alt = true;
      else shift = true;
      resto = m[2];
    }
    const key = normalizarNombreTecla(resto);
    if (!key || MODS_SOLOS.has(key)) return null;
    if (SIMBOLOS_MAYUS.has(key)) shift = false;
    return { ctrl, alt, shift, key };
  }

  function normalizarTecla(entrada) {
    if (entrada === "") return "";
    if (entrada && typeof entrada === "object" && typeof entrada.key === "string") return teclaDeEvento(entrada);
    const d = descomponer(entrada);
    if (!d) return null;
    if (!d.key) return "";
    const partes = [];
    if (d.ctrl) partes.push("Ctrl");
    if (d.alt) partes.push("Alt");
    if (d.shift) partes.push("Shift");
    partes.push(d.key);
    return partes.join("+");
  }

  const tieneModificador = (combo) => {
    const d = descomponer(combo);
    return !!(d && (d.ctrl || d.alt));
  };
  const esTeclaFuncion = (combo) => {
    const d = descomponer(combo);
    return !!(d && /^F([1-9]|1[0-2])$/.test(d.key));
  };

  /** Partes listas para pintar como etiquetas <kbd> en español ("Ctrl+Shift+A" -> ["Ctrl", "Mayús", "A"]). */
  function partesTecla(combo) {
    const d = descomponer(combo);
    if (!d || !d.key) return [];
    const p = [];
    if (d.ctrl) p.push("Ctrl");
    if (d.alt) p.push("Alt");
    if (d.shift) p.push("Mayús");
    p.push(propio(NOMBRES_TECLA, d.key) ? NOMBRES_TECLA[d.key] : d.key);
    return p;
  }

  function formatearTecla(combo) {
    const p = partesTecla(combo);
    return p.length ? p.join(" + ") : "Sin atajo";
  }

  function etiquetaCorta(combo) {
    return partesTecla(combo).join(" ");
  }

  function teclaDe(accion, estado) {
    const item = POR_ACCION.get(accion);
    if (!item) return "";
    const mapa = (estado || leer()).atajos || {};
    return propio(mapa, accion) ? mapa[accion] : item.tecla;
  }

  function listarAtajos(estado) {
    const e = estado || leer();
    return ATAJOS.map((item) => {
      const actual = teclaDe(item.accion, e);
      return {
        accion: item.accion,
        nombre: item.nombre,
        grupo: item.grupo,
        teclaDefecto: item.tecla,
        tecla: actual,
        personalizado: actual !== item.tecla,
      };
    });
  }

  function validarAtajo(accion, entrada, estado) {
    const item = POR_ACCION.get(accion);
    if (!item) return { ok: false, motivo: "desconocida", mensaje: "Esa acción no existe." };
    if (entrada === "") return { ok: true, tecla: "" };
    const norm = normalizarTecla(entrada);
    if (!norm) return { ok: false, motivo: "invalida", mensaje: "Presiona una tecla válida (por ejemplo Ctrl + J o Alt + B)." };
    if (propio(PROHIBIDOS, norm)) return { ok: false, motivo: "prohibido", mensaje: PROHIBIDOS[norm] };
    if (item.enCampo && !tieneModificador(norm) && !esTeclaFuncion(norm)) {
      return { ok: false, motivo: "necesita-modificador", mensaje: "Esta acción se usa al escribir: agrega Ctrl, Alt o usa una tecla F (por ejemplo Ctrl + J)." };
    }
    const e = estado || leer();
    for (const otra of ATAJOS) {
      if (otra.accion === accion) continue;
      if (teclaDe(otra.accion, e) === norm) {
        return {
          ok: false,
          motivo: "choca",
          chocaCon: otra.accion,
          chocaNombre: otra.nombre,
          mensaje: `Esa tecla ya la usa «${otra.nombre}». Quítala de ahí primero o elige otra.`,
        };
      }
    }
    return { ok: true, tecla: norm };
  }

  /** Cambia un atajo. Si se pasa `borrador`, modifica ese objeto sin escribir todavía en localStorage. */
  function cambiarAtajo(accion, entrada, borrador) {
    const estado = borrador || leer();
    const res = validarAtajo(accion, entrada, estado);
    if (!res.ok) return res;
    const item = POR_ACCION.get(accion);
    const mapa = Object.assign({}, estado.atajos || {});
    if (res.tecla === item.tecla) delete mapa[accion];
    else mapa[accion] = res.tecla;
    if (borrador) { borrador.atajos = mapa; return { ok: true, tecla: res.tecla, estado: borrador }; }
    const nuevo = guardar({ atajos: mapa });
    return { ok: true, tecla: res.tecla, estado: nuevo };
  }

  function quitarAtajo(accion, borrador) {
    return cambiarAtajo(accion, "", borrador);
  }

  function restablecerAtajos(borrador) {
    if (borrador) { borrador.atajos = {}; return borrador; }
    return guardar({ atajos: {} });
  }

  /** ¿El evento coincide con la combinación guardada? Tolerante con Alt+/ en teclados donde "/" lleva Mayús. */
  function coincideEvento(combo, e) {
    if (!combo) return false;
    const ev = teclaDeEvento(e);
    if (!ev) return false;
    if (ev === combo) return true;
    if (combo === "Alt+/" && e.altKey && !(e.ctrlKey || e.metaKey) && e.key === "/") return true;
    return false;
  }

  /** Devuelve la acción cuya tecla coincide con el evento en el contexto actual, o null. */
  function buscarAccionPorEvento(e, ctx, estado) {
    if (!e || e.defaultPrevented) return null;
    if (e.getModifierState && e.getModifierState("AltGraph") && e.ctrlKey && e.altKey) return null;
    ctx = ctx || {};
    const eAct = estado || leer();
    for (const item of ATAJOS) {
      const t = teclaDe(item.accion, eAct);
      if (!t || !coincideEvento(t, e)) continue;
      if (item.enEditor && !ctx.enEditor) continue;
      const conMod = tieneModificador(t) || esTeclaFuncion(t);
      if (!conMod && (ctx.enCampo || ctx.hayHoja)) continue;
      if (!conMod && item.sinEditor && ctx.enEditor) continue;
      // En el texto, Ctrl+K de fábrica es para insertar enlace (como en Word).
      if (item.accion === "paleta" && t === "Ctrl+K" && ctx.enCampo) continue;
      if (item.soloFueraCampo && !item.enCampo && ctx.enCampo) continue;
      return item.accion;
    }
    return null;
  }

  /** ¿El evento corresponde al atajo de fábrica de alguna acción que la persona ya cambió o quitó? */
  function esDefectoCambiado(e, acciones, estado) {
    if (!e) return false;
    const eAct = estado || leer();
    for (const acc of acciones) {
      const item = POR_ACCION.get(acc);
      if (!item) continue;
      const actual = teclaDe(acc, eAct);
      if (actual !== item.tecla && coincideEvento(item.tecla, e)) return true;
    }
    return false;
  }

  // ================================================================== háptica (vibración)
  function duracionHaptica(tipo, nivel) {
    const n = propio(HAPTICA_MS, nivel) ? nivel : "normal";
    const tabla = HAPTICA_MS[n];
    if (typeof tipo === "number" && tipo > 0) return Math.min(80, Math.round(tipo));
    return propio(tabla, tipo) ? tabla[tipo] : tabla.toque;
  }

  function ejecutarVibracion(ms) {
    try {
      const puente = (typeof window !== "undefined" && (window.NomAndroid || window.NOMApp)) || null;
      if (puente && typeof puente.vibrar === "function") { puente.vibrar(ms); return true; }
      if (puente && typeof puente.haptica === "function") { puente.haptica(ms); return true; }
      const nav = (typeof navigator !== "undefined" && navigator) || (typeof window !== "undefined" && window.navigator) || null;
      if (nav && typeof nav.vibrate === "function") { nav.vibrate(ms); return true; }
    } catch (e) { /* dispositivo sin vibración o bloqueado */ }
    return false;
  }

  /** Vibra un instante si la preferencia está encendida. Nunca truena si el equipo no vibra. */
  function vibrar(tipo) {
    const p = leer();
    if (!p.haptica) return false;
    return ejecutarVibracion(duracionHaptica(tipo, p.hapticaNivel));
  }

  /** Vibración de muestra al probar una intensidad en Ajustes (aunque aún no se haya pulsado Guardar). */
  function probarVibracion(nivel) {
    return ejecutarVibracion(duracionHaptica("elegir", nivel));
  }

  // ================================================================== menú rápido: ruedita o lista
  function esLista() {
    return leer().menuRapido === "lista";
  }

  // ================================================================== sección «Atajos» dentro de Ajustes
  function montarAtajos(cont, borrador, alCambiar) {
    const { esc, ico, $, $$ } = NOM;
    let grabando = null;
    let mensajeAviso = "";

    const dejarDeGrabar = () => {
      if (!grabando) return;
      grabando = null;
      window.removeEventListener("keydown", capturar, true);
    };

    function capturar(e) {
      if (!grabando) return;
      if (MODS_SOLOS.has(e.key)) return;
      e.preventDefault();
      e.stopPropagation();
      if (e.stopImmediatePropagation) e.stopImmediatePropagation();
      if (e.key === "Escape") {
        mensajeAviso = "";
        dejarDeGrabar();
        pintar();
        return;
      }
      const res = cambiarAtajo(grabando, e, borrador);
      if (!res.ok) {
        mensajeAviso = res.mensaje;
        pintar();
        return;
      }
      mensajeAviso = "";
      dejarDeGrabar();
      pintar();
      alCambiar && alCambiar();
    }

    function pintar() {
      const filas = listarAtajos(borrador);
      const hayCambios = filas.some((f) => f.personalizado);
      const kbdHtml = (combo) => partesTecla(combo).map((x) => `<kbd>${esc(x)}</kbd>`).join("");
      cont.innerHTML = `<div class="atajos-ed">
        <div class="atajos-ed-cab">
          <div>
            <span class="etiqueta">${ico("keyboard", "ico-s")} Atajos <span class="suave" style="font-weight:600">(en este equipo)</span></span>
            <div class="ayuda">Presiona «Cambiar» y luego la tecla que prefieras.</div>
          </div>
          <button type="button" class="btn chico" data-atajos-fabrica ${hayCambios ? "" : "disabled"}>Volver a los de siempre</button>
        </div>
        <div class="alerta aviso" data-atajos-aviso role="alert" ${mensajeAviso ? "" : "hidden"}>${ico("triangle-alert", "ico-s")}<span>${esc(mensajeAviso)}</span></div>
        <div class="atajos-ed-lista">
          ${filas.map((f) => {
            const enGrabacion = grabando === f.accion;
            return `<div class="atajos-ed-fila ${enGrabacion ? "grabando" : ""}" data-atajo-fila="${esc(f.accion)}">
              <div class="atajos-ed-nom">
                <b>${esc(f.nombre)}</b>
                ${f.personalizado ? `<small>Antes: ${esc(formatearTecla(f.teclaDefecto))}</small>` : ""}
              </div>
              <div class="atajos-ed-tecla">
                ${enGrabacion
                  ? `<span class="atajo-grabando" role="status" aria-live="polite">Presiona una tecla… (Esc cancela)</span>`
                  : f.tecla ? kbdHtml(f.tecla) : `<span class="suave chico">Sin atajo</span>`}
              </div>
              <div class="atajos-ed-acc">
                <button type="button" class="btn chico" data-cambiar="${esc(f.accion)}">${enGrabacion ? "Cancelar" : "Cambiar"}</button>
                <button type="button" class="btn chico plano" data-quitar="${esc(f.accion)}" ${f.tecla ? "" : "disabled"}>Quitar</button>
              </div>
            </div>`;
          }).join("")}
        </div>
      </div>`;

      const bFab = $("[data-atajos-fabrica]", cont);
      if (bFab) bFab.onclick = () => {
        dejarDeGrabar();
        mensajeAviso = "";
        restablecerAtajos(borrador);
        pintar();
        alCambiar && alCambiar();
      };
      $$("[data-cambiar]", cont).forEach((b) => (b.onclick = () => {
        const acc = b.dataset.cambiar;
        if (grabando === acc) { dejarDeGrabar(); mensajeAviso = ""; pintar(); return; }
        dejarDeGrabar();
        mensajeAviso = "";
        grabando = acc;
        window.addEventListener("keydown", capturar, true);
        pintar();
      }));
      $$("[data-quitar]", cont).forEach((b) => (b.onclick = () => {
        const acc = b.dataset.quitar;
        dejarDeGrabar();
        mensajeAviso = "";
        quitarAtajo(acc, borrador);
        pintar();
        alCambiar && alCambiar();
      }));
    }

    pintar();
    return { desmontar: dejarDeGrabar, repintar: pintar };
  }

  NOM.prefExtra = {
    CLAVE,
    ATAJOS,
    PROHIBIDOS,
    HAPTICA_MS,
    fabrica,
    leer,
    guardar,
    restablecer,
    clonar,
    teclaDeEvento,
    normalizarTecla,
    partesTecla,
    formatearTecla,
    etiquetaCorta,
    teclaDe,
    listarAtajos,
    validarAtajo,
    cambiarAtajo,
    quitarAtajo,
    restablecerAtajos,
    coincideEvento,
    buscarAccionPorEvento,
    esDefectoCambiado,
    duracionHaptica,
    vibrar,
    probarVibracion,
    esLista,
    montarAtajos,
  };
})();

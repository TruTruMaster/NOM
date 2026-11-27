/* NOM · Errores: qué pasó y qué hacer, en palabras de todos los días, para cada respuesta que no salió bien.
   Lo usan api.js (el mensaje de cada error cuando el servidor no manda uno propio) y la página de «no existe».
   Los mensajes propios del servidor (en español, con el campo que falló) se siguen mostrando tal cual. */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});

  const MENSAJES = {
    0: "Sin conexión: revisa el Wi-Fi o los datos. Lo que hagas se guarda en este equipo.",
    400: "Algo de lo que se mandó no es válido. Revisa los datos y vuelve a intentarlo.",
    401: "Tu sesión terminó. Vuelve a entrar con tu usuario y contraseña.",
    403: "No tienes permiso para hacer esto. Pídeselo a quien administra el espacio.",
    404: "No lo encontramos: puede que lo hayan borrado o que ya no tengas acceso.",
    408: "NOM tardó demasiado en contestar. Vuelve a intentarlo en un momento.",
    409: "Alguien más lo cambió al mismo tiempo. Vuelve a abrirlo e intenta de nuevo.",
    413: "El archivo es demasiado grande para subirlo. Prueba con uno más pequeño o divídelo en partes.",
    415: "Ese tipo de archivo no se puede usar aquí.",
    423: "Este espacio está protegido con tu contraseña y en este equipo falta su llave. Vuelve a entrar con tu contraseña para abrirlo.",
    429: "Hiciste muchos intentos seguidos. Espera un minuto y vuelve a intentarlo.",
    500: "Algo falló en NOM. No es algo que hayas hecho tú: tu trabajo sigue en este equipo. Vuelve a intentarlo en un momento.",
  };
  const CAIDO = "NOM en Internet no responde (puede estar reiniciándose o sin conexión). Lo que hagas se guarda en este equipo y se manda cuando vuelva.";

  /** ¿La respuesta dice que el servidor no está (la Pi apagada o sin Internet: Cloudflare contesta 502/503/504/52x
      con su propia página)? Un 503 de NOM trae su mensaje en JSON: ese no cuenta. */
  function caido(estado, datos) {
    const sinMensajeDeNom = !(datos && typeof datos.error === "string");
    return sinMensajeDeNom && (estado === 502 || estado === 503 || estado === 504 || (estado >= 520 && estado <= 530));
  }

  /** El mensaje para un estado HTTP (0 = sin red). */
  function mensaje(estado) {
    if (MENSAJES[estado]) return MENSAJES[estado];
    if (caido(estado, null)) return CAIDO;
    if (estado >= 500) return MENSAJES[500];
    if (estado >= 400) return MENSAJES[400];
    return "Algo no salió bien. Vuelve a intentarlo.";
  }

  NOM.errores = { mensaje, caido, CAIDO, MENSAJES };
  if (typeof module !== "undefined" && module.exports) module.exports = NOM.errores;
})();

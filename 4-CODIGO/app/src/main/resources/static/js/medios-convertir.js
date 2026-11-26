/* NOM — convertir video y audio en el equipo antes de subirlos (docs/PLAN-MEDIOS.md, fase 1).
 *
 * Lo que la persona sube (MP4, MOV, M4V, WebM, MKV, MP3, M4A, AAC, WAV, OGG, FLAC…) se convierte aquí, con los
 * codificadores del navegador (WebCodecs, por medio de Mediabunny), al formato de NOM: un WebM con video AV1 y audio
 * Opus. Son los códecs más eficientes que el navegador decodifica solo y en tiempo real, así que el archivo queda lo
 * más ligero posible sin que se note la diferencia, y al servidor (la Pi de la casa) solo sube lo ya convertido.
 *
 * Si el equipo no tiene codificador AV1 se usa VP9. Si no puede convertir (sin WebCodecs) se sube el original, si
 * cabe en el límite y el navegador lo reproduce. También hace la vuelta: de lo de NOM a MP4 (H.264 y AAC) para que
 * PowerPoint lo reproduzca al descargar el .pptx (eso se hace en el equipo: el servidor solo manda lo ligero).
 */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});

  /** Límites (ya convertido). Pensados para muchas cuentas en el disco de 2 TB de la Pi: ver docs/PLAN-MEDIOS.md. */
  const LIMITES = {
    origenBytes: 4 * 1024 ** 3,          // lo que se elige en el equipo (se convierte antes de subir)
    videoBytes: 500 * 1024 ** 2,         // un video ya convertido
    audioBytes: 200 * 1024 ** 2,         // un audio ya convertido
    videoSeg: 60 * 60,                   // 60 minutos de video
    audioSeg: 3 * 60 * 60,               // 3 horas de audio
    ladoLargo: 1920, ladoCorto: 1080,    // hasta 1080p (lo más grande se reduce sin deformar)
    fpsMax: 60,
  };
  const ACEPTA = "video/*,audio/*,.mp4,.m4v,.mov,.webm,.mkv,.mp3,.m4a,.aac,.wav,.ogg,.oga,.opus,.flac";
  const MB = (b) => (b / 1024 / 1024 >= 10 ? Math.round(b / 1024 / 1024) : Math.round((b / 1024 / 1024) * 10) / 10) + " MB";
  const tiempo = (s) => { s = Math.round(s); const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), x = s % 60; return (h ? h + ":" + String(m).padStart(2, "0") : m) + ":" + String(x).padStart(2, "0"); };

  // ------------------------------------------------------------ cargar la biblioteca (solo cuando hace falta)
  let cargando = null;
  function cargar() {
    if (window.Mediabunny) return Promise.resolve(window.Mediabunny);
    if (cargando) return cargando;
    cargando = new Promise((ok, mal) => {
      const s = document.createElement("script");
      s.src = "js/vendor/mediabunny.min.js";
      s.onload = () => (window.Mediabunny ? ok(window.Mediabunny) : mal(new Error("No se pudo cargar el convertidor de video")));
      s.onerror = () => { cargando = null; mal(new Error("No se pudo cargar el convertidor de video (¿sin conexión la primera vez?)")); };
      document.head.append(s);
    });
    return cargando;
  }
  /** ¿Este navegador puede convertir? (WebCodecs) */
  const puedeConvertir = () => typeof window.VideoEncoder === "function" || typeof window.AudioEncoder === "function";

  /** Un error con un mensaje para la persona. */
  const errorNom = (m, codigo) => Object.assign(new Error(m), { nom: true, codigo });

  // ------------------------------------------------------------ analizar
  /** Qué es el archivo: { tipo: "video" | "audio", duracion (s), ancho, alto, fps, canales, contenedor }. */
  async function analizar(archivo) {
    const M = await cargar();
    const input = new M.Input({ formats: M.ALL_FORMATS, source: new M.BlobSource(archivo) });
    try {
      const vt = await input.getPrimaryVideoTrack(), at = await input.getPrimaryAudioTrack();
      if (!vt && !at) throw errorNom("Ese archivo no tiene video ni audio que NOM pueda leer.", "sin-pistas");
      const duracion = await input.computeDuration();
      const r = { tipo: vt ? "video" : "audio", duracion, contenedor: (await input.getMimeType().catch(() => "")) || archivo.type || "" };
      if (vt) {
        r.ancho = vt.displayWidth; r.alto = vt.displayHeight;
        r.codecVideo = vt.codec || null;
        try { const st = await vt.computePacketStats(300); r.fps = st && st.averagePacketRate ? Math.round(st.averagePacketRate * 100) / 100 : null; r.bpsVideo = st ? st.averageBitrate : null; } catch (e) { r.fps = null; }
      }
      if (at) {
        r.canales = at.numberOfChannels; r.codecAudio = at.codec || null;
        try { const st = await at.computePacketStats(300); r.bpsAudio = st ? st.averageBitrate : null; } catch (e) { r.bpsAudio = null; }
      }
      return { input, info: r, vt, at };
    } catch (e) { input.dispose && input.dispose(); throw e; }
  }

  /** El tamaño de salida: dentro de 1920 × 1080 (de pie o acostado), par y sin deformar. */
  function medidaSalida(w, h) {
    const largo = Math.max(w, h), corto = Math.min(w, h);
    const k = Math.min(1, LIMITES.ladoLargo / largo, LIMITES.ladoCorto / corto);
    const par = (v) => Math.max(2, Math.round(v / 2) * 2);
    return { ancho: par(w * k), alto: par(h * k), reducido: k < 1 };
  }

  /**
   * La tasa del video de salida (bits por segundo), en modo variable. Medido el 17-nov-2026 en Chromium: su AV1
   * (libaom en tiempo real) rinde a la misma tasa casi como x264 «medium», no mejor (720p: 36.8 contra 37.9 dB a
   * 1 Mb/s), y su modo de calidad constante no baja de ~6.4 Mb/s. Por eso lo ligero sale de no gastar de más: 60 % de
   * la tasa de un H.264 (90 % de HEVC o VP9), con un techo de 0.1 bits por píxel y cuadro (1080p a 30: ~6.2 Mb/s,
   * donde ya no se distingue) y un piso de 0.025 (~1.6 Mb/s). Los videos de teléfono (~17 Mb/s) quedan ~60 % más
   * ligeros; uno ya comprimido, menos (si no se ahorra, se sube el original).
   */
  function tasaVideo(info, med, codec, fps) {
    const pxs = med.ancho * med.alto * (fps || 30);
    const techo = pxs * 0.1, piso = pxs * 0.025;
    const FACTORES = { avc: 0.6, hevc: 0.9, vp9: 0.9, av1: 1, vp8: 0.65 };
    const factor = Object.prototype.hasOwnProperty.call(FACTORES, info.codecVideo) ? FACTORES[info.codecVideo] : 0.45;
    const escala = info.ancho && info.alto ? (med.ancho * med.alto) / (info.ancho * info.alto) : 1;
    let t = info.bpsVideo ? info.bpsVideo * factor * escala : techo;
    if (codec === "vp9") t *= 1.2;
    return Math.round(Math.max(piso, Math.min(techo * (codec === "vp9" ? 1.2 : 1), t)));
  }
  /** La tasa del audio: Opus a 128 kb/s estéreo (64 en mono) o ~75 % de la de un MP3/AAC si era menor (sin bajar de 48/32). */
  function tasaAudio(info, canales) {
    const tope = canales > 1 ? 128000 : 64000, piso = canales > 1 ? 48000 : 32000;
    if (!info.bpsAudio || /^(flac|pcm)/.test(info.codecAudio || "")) return tope;
    const factor = info.codecAudio === "opus" ? 1 : info.codecAudio === "vorbis" ? 0.9 : 0.75;
    return Math.round(Math.max(piso, Math.min(tope, info.bpsAudio * factor)));
  }

  // ------------------------------------------------------------ al formato de NOM (AV1 / Opus en WebM)
  /**
   * Convierte un archivo (File o Blob) al formato ligero de NOM.
   * opc: { alAvance(fraccion 0-1, segundosHechos), senal (AbortSignal) }
   * Devuelve { blob, mime, tipo, convertido, codecVideo, codecAudio, duracion, ancho, alto, bytesAntes, bytesDespues, nota }.
   */
  async function aLigero(archivo, opc) {
    opc = opc || {};
    if (archivo.size > LIMITES.origenBytes) throw errorNom(`El archivo pesa ${MB(archivo.size)}; el máximo para convertir es ${MB(LIMITES.origenBytes)}.`, "origen-grande");
    if (!puedeConvertir()) return original(archivo, null, "Este navegador no puede convertir video; se sube como está.");
    const M = await cargar();
    const { input, info, vt, at } = await analizar(archivo);
    try {
      const esVideo = info.tipo === "video";
      if (esVideo && info.duracion > LIMITES.videoSeg) throw errorNom(`El video dura ${tiempo(info.duracion)}; el máximo es ${tiempo(LIMITES.videoSeg)}.`, "largo");
      if (!esVideo && info.duracion > LIMITES.audioSeg) throw errorNom(`El audio dura ${tiempo(info.duracion)}; el máximo es ${tiempo(LIMITES.audioSeg)}.`, "largo");
      // códecs: AV1 (o VP9) y Opus, lo que este equipo pueda codificar
      let video;
      if (vt) {
        const med = medidaSalida(info.ancho, info.alto);
        const codec = await M.getFirstEncodableVideoCodec(["av1", "vp9"], { width: med.ancho, height: med.alto });
        if (!codec) return original(archivo, info, "Este equipo no puede codificar AV1 ni VP9; se sube como está.");
        const fps = info.fps ? Math.min(LIMITES.fpsMax, info.fps) : 30;
        video = { codec, width: med.ancho, height: med.alto, fit: "contain", bitrate: opc.tasaVideo || tasaVideo(info, med, codec, fps), forceTranscode: true, keyFrameInterval: 4 };
        if (info.fps && info.fps > LIMITES.fpsMax) video.frameRate = LIMITES.fpsMax;
      }
      let audio;
      if (at) {
        const can = Math.min(2, at.numberOfChannels || 2);
        // Opus 128 kb/s estéreo (64 kb/s en mono) ya no se distingue del original en casi ninguna grabación
        const tasa = tasaAudio(info, can);
        const codec = await M.getFirstEncodableAudioCodec(["opus"], { numberOfChannels: can, sampleRate: 48000, bitrate: tasa });
        if (!codec) return original(archivo, info, "Este equipo no puede codificar Opus; se sube como está.");
        audio = { codec, numberOfChannels: can, sampleRate: 48000, bitrate: tasa, forceTranscode: true };
      }
      const output = new M.Output({ format: new M.WebMOutputFormat(), target: new M.BufferTarget() });
      const conv = await M.Conversion.init({ input, output, video: video || { discard: true }, audio: audio || { discard: true }, showWarnings: false });
      if (!conv.isValid) {
        const por = (conv.discardedTracks || []).map((d) => d.reason).join(", ");
        return original(archivo, info, "No se pudo convertir (" + (por || "formato no soportado") + "); se sube como está.");
      }
      if (opc.alAvance) conv.onProgress = (p, t) => opc.alAvance(Math.max(0, Math.min(1, p)), t);
      const cancelar = () => conv.cancel();
      if (opc.senal) { if (opc.senal.aborted) throw errorNom("Conversión cancelada", "cancelado"); opc.senal.addEventListener("abort", cancelar, { once: true }); }
      try { await conv.execute(); }
      catch (e) { if (opc.senal && opc.senal.aborted) throw errorNom("Conversión cancelada", "cancelado"); throw e; }
      finally { if (opc.senal) opc.senal.removeEventListener("abort", cancelar); }
      const mime = esVideo ? "video/webm" : "audio/webm";
      const blob = new Blob([output.target.buffer], { type: mime });
      const r = { blob, mime, tipo: info.tipo, convertido: true, codecVideo: video ? video.codec : null, codecAudio: audio ? audio.codec : null, duracion: info.duracion,
        ancho: video ? video.width : null, alto: video ? video.height : null, bytesAntes: archivo.size, bytesDespues: blob.size, nota: null };
      // si el original ya era más ligero (y el navegador lo reproduce), se queda el original
      if (blob.size >= archivo.size * 0.97 && reproducible(archivo.type)) return original(archivo, info, "El archivo ya era ligero; se sube como está.");
      revisarTope(r);
      return r;
    } finally {
      input.dispose && input.dispose();
    }
  }
  /** Sin convertir: el archivo tal cual, si cabe y se puede reproducir. */
  function original(archivo, info, nota) {
    const tipo = info ? info.tipo : /^audio\//.test(archivo.type) ? "audio" : "video";
    if (!reproducible(archivo.type)) throw errorNom("Este equipo no puede convertir ese formato y el navegador no lo reproduce. Prueba desde una computadora con Chrome, Edge o Firefox.", "no-convierte");
    const r = { blob: archivo, mime: archivo.type, tipo, convertido: false, codecVideo: null, codecAudio: null, duracion: info ? info.duracion : null,
      ancho: info ? info.ancho : null, alto: info ? info.alto : null, bytesAntes: archivo.size, bytesDespues: archivo.size, nota };
    revisarTope(r);
    return r;
  }
  function revisarTope(r) {
    const tope = r.tipo === "video" ? LIMITES.videoBytes : LIMITES.audioBytes;
    if (r.bytesDespues > tope) throw errorNom(`Aun ${r.convertido ? "convertido" : "sin convertir"}, el ${r.tipo} pesa ${MB(r.bytesDespues)}; el máximo es ${MB(tope)}. Recórtalo o súbelo en partes.`, "grande");
  }
  /** ¿El navegador reproduce este tipo de archivo? */
  function reproducible(mime) {
    if (!mime) return false;
    const el = document.createElement(/^audio\//.test(mime) ? "audio" : "video");
    return !!el.canPlayType(mime);
  }

  // ------------------------------------------------------------ a MP4 para PowerPoint
  /**
   * De lo de NOM (o cualquier medio) a MP4 con H.264 y AAC, para el .pptx. Si el equipo no tiene AAC, el audio va
   * en Opus (PowerPoint quizá no lo suene) y se avisa en «nota».
   */
  async function aMp4(blob, opc) {
    opc = opc || {};
    if (blob && blob.size > LIMITES.origenBytes) throw errorNom(`El archivo pesa ${MB(blob.size)}; el máximo para convertir es ${MB(LIMITES.origenBytes)}.`, "origen-grande");
    const M = await cargar();
    const input = new M.Input({ formats: M.ALL_FORMATS, source: new M.BlobSource(blob) });
    try {
      const vt = await input.getPrimaryVideoTrack(), at = await input.getPrimaryAudioTrack();
      if (!vt && !at) throw errorNom("Ese archivo no tiene video ni audio.", "sin-pistas");
      const dur = await input.computeDuration().catch(() => 0);
      if (vt && dur > LIMITES.videoSeg) throw errorNom(`El video dura ${tiempo(dur)}; el máximo es ${tiempo(LIMITES.videoSeg)}.`, "largo");
      if (!vt && dur > LIMITES.audioSeg) throw errorNom(`El audio dura ${tiempo(dur)}; el máximo es ${tiempo(LIMITES.audioSeg)}.`, "largo");
      let video = { discard: true }, audio = { discard: true }, nota = null;
      if (vt) {
        const med = medidaSalida(vt.displayWidth || 1280, vt.displayHeight || 720);
        const codec = await M.getFirstEncodableVideoCodec(["avc"], { width: med.ancho, height: med.alto });
        if (!codec) throw errorNom("Este equipo no puede codificar H.264 para PowerPoint.", "sin-avc");
        video = { codec, width: med.ancho, height: med.alto, fit: "contain", bitrate: M.QUALITY_HIGH, hardwareAcceleration: "prefer-hardware" };
      }
      if (at) {
        const codec = await M.getFirstEncodableAudioCodec(["aac", "opus"], { numberOfChannels: Math.min(2, at.numberOfChannels || 2), sampleRate: 48000 });
        if (!codec) throw errorNom("Este equipo no puede codificar el audio para PowerPoint.", "sin-aac");
        if (codec !== "aac") nota = "Este equipo no tiene codificador AAC: el audio va en Opus y PowerPoint quizá no lo suene.";
        audio = { codec, bitrate: 160000, numberOfChannels: Math.min(2, at.numberOfChannels || 2) };
      }
      const output = new M.Output({ format: new M.Mp4OutputFormat({ fastStart: "in-memory" }), target: new M.BufferTarget() });
      const conv = await M.Conversion.init({ input, output, video, audio, showWarnings: false });
      if (!conv.isValid) throw errorNom("No se pudo preparar el video para PowerPoint.", "mp4");
      if (opc.alAvance) conv.onProgress = (p) => opc.alAvance(p);
      const cancelar = () => conv.cancel();
      if (opc.senal) { if (opc.senal.aborted) throw errorNom("Conversión cancelada", "cancelado"); opc.senal.addEventListener("abort", cancelar, { once: true }); }
      try { await conv.execute(); }
      catch (e) { if (opc.senal && opc.senal.aborted) throw errorNom("Conversión cancelada", "cancelado"); throw e; }
      finally { if (opc.senal) opc.senal.removeEventListener("abort", cancelar); }
      const out = new Blob([output.target.buffer], { type: vt ? "video/mp4" : "audio/mp4" });
      revisarTope({ tipo: vt ? "video" : "audio", convertido: true, bytesDespues: out.size });
      return { blob: out, nota };
    } finally {
      input.dispose && input.dispose();
    }
  }

  // ------------------------------------------------------------ elegir y convertir con su ventana de avance
  /**
   * Pide un archivo de video o audio, lo convierte mostrando el avance (se puede cancelar) y devuelve el resultado
   * de aLigero, o null si la persona canceló.
   */
  async function elegirYConvertir() {
    const archivo = await new Promise((ok) => {
      const i = document.createElement("input");
      i.type = "file"; i.accept = ACEPTA;
      i.onchange = () => ok(i.files && i.files[0] ? i.files[0] : null);
      i.click();
    });
    if (!archivo) return null;
    return convertirConAvance(archivo);
  }
  function convertirConAvance(archivo) {
    const { html, $, ico, esc } = NOM;
    const ctl = new AbortController();
    const cuerpo = html(`<div class="col"><p style="font-weight:400"><b>${esc(archivo.name)}</b> · ${MB(archivo.size)}</p>
      <p class="suave chico" data-paso style="font-weight:400">Revisando el archivo…</p>
      <div class="mc-barra" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><i data-i></i></div>
      <p class="suave chico" style="font-weight:400">Se convierte en este equipo a un formato más ligero (AV1 y Opus) antes de subirlo; la calidad se ve igual. Puedes seguir usando NOM mientras tanto si no cierras esta ventana.</p></div>`);
    const pie = html(`<div style="display:contents"><span class="crece"></span><button class="btn" type="button" data-no>${ico("x", "ico-s")} Cancelar</button></div>`);
    let fin = false;
    const h = NOM.hoja({ titulo: "Preparando el video o audio", icono: "film", cuerpo, pie, alCerrar: () => { if (!fin) ctl.abort(); } });
    $("[data-no]", pie).onclick = () => h.cerrar();
    const t0 = performance.now();
    return aLigero(archivo, {
      senal: ctl.signal,
      alAvance: (p) => {
        const b = $("[data-i]", cuerpo), bar = $(".mc-barra", cuerpo);
        if (b) b.style.width = Math.round(p * 100) + "%";
        if (bar) bar.setAttribute("aria-valuenow", String(Math.round(p * 100)));
        const s = (performance.now() - t0) / 1000, resta = p > 0.02 ? (s / p) * (1 - p) : null;
        const paso = $("[data-paso]", cuerpo);
        if (paso) paso.textContent = `Convirtiendo: ${Math.round(p * 100)} %` + (resta != null ? ` · faltan unos ${tiempo(resta)}` : "");
      },
    }).then((r) => {
      fin = true; h.cerrar();
      const ahorro = r.convertido && r.bytesAntes > 0 ? Math.round((1 - r.bytesDespues / r.bytesAntes) * 100) : 0;
      if (NOM.aviso) NOM.aviso(r.convertido ? `Listo: ${MB(r.bytesDespues)} (${ahorro > 0 ? ahorro + " % más ligero" : "mismo tamaño"})` : r.nota || "Se sube como está", { tipo: r.convertido ? "ok" : "info", ms: 3500 });
      return r;
    }).catch((e) => {
      fin = true; h.cerrar();
      if (e && e.codigo === "cancelado") return null;
      if (NOM.aviso) NOM.aviso(e && e.nom ? e.message : "No se pudo convertir: " + ((e && e.message) || e), { tipo: "error", ms: 6000 });
      return null;
    });
  }

  NOM.mediosConvertir = { LIMITES, ACEPTA, cargar, puedeConvertir, analizar, aLigero, aMp4, elegirYConvertir, convertirConAvance, medidaSalida, reproducible };
})();

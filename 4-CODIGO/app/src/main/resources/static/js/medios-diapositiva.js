/* NOM — videos y audios de PowerPoint al presentar.
 *
 * En PowerPoint un video o un audio es una imagen (la portada del video o el ícono del altavoz) con el medio aparte. Al
 * importar, ese bloque de imagen queda con «medio»: { tipo: "video" | "audio", ruta (la parte del archivo, p. ej.
 * ppt/media/media1.mp4), recorte { ini, fin } (ms que se quitan al principio y al final, p14:trim), fundido { ent, sal }
 * (ms, p14:fade), marcadores [{ n, t (ms) }] (p14:bmkLst) }. El medio NO se copia al documento: al presentar se saca del
 * archivo de Office original que NOM ya guarda dentro del documento (_origen, ver origen.js).
 *
 * Cómo se reproduce lo dice la diapositiva, como en PowerPoint: el nodo p:video / p:audio de sus tiempos
 * ("animMedios:<página>": volumen, silencio, repetir hasta detenerlo, ocultar mientras no se reproduce, seguir en las
 * diapositivas siguientes, pantalla completa) y los comandos de sus animaciones (playFrom, togglePause, pause, resume,
 * stop, seek), que el motor de animaciones manda aquí (ctx.alMedio). Al pasar el puntero aparece una barra con
 * reproducir/pausa, la posición (con los marcadores) y el sonido; las teclas son las de PowerPoint: Alt+P, Alt+Q,
 * Alt+↑/↓ y Alt+Inicio/Fin.
 */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});

  const MIME = { mp4: "video/mp4", m4v: "video/mp4", mov: "video/quicktime", webm: "video/webm", ogv: "video/ogg", wmv: "video/x-ms-wmv",
    avi: "video/x-msvideo", mpg: "video/mpeg", mpeg: "video/mpeg", mp3: "audio/mpeg", m4a: "audio/mp4", wav: "audio/wav", wma: "audio/x-ms-wma",
    aac: "audio/aac", ogg: "audio/ogg", oga: "audio/ogg", flac: "audio/flac", mid: "audio/midi", midi: "audio/midi" };
  const mimeDe = (ruta) => MIME[(String(ruta || "").split(".").pop() || "").toLowerCase()] || "application/octet-stream";
  const loc = (q) => String(q || "").replace(/^.*:/, "");
  const A = (n) => (Array.isArray(n) && n[1]) || {};
  const H = (n) => (Array.isArray(n) && n[2]) || [];
  const ico = (n) => (NOM.ico ? NOM.ico(n, "ico-s") : "");
  const reloj = (s) => { s = Math.max(0, Math.floor(s || 0)); return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`; };

  /**
   * Los ajustes de PowerPoint del medio de un bloque, del nodo p:video / p:audio de la diapositiva:
   * <p:video fullScrn="1"><p:cMediaNode vol="80000" mute="1" showWhenStopped="0" numSld="999"><p:cTn repeatCount="indefinite"/>…
   */
  function ajustesDe(nodos, bid) {
    for (const n of nodos || []) {
      if (!/^(video|audio)$/.test(loc(n && n[0]))) continue;
      const cm = H(n).find((c) => Array.isArray(c) && loc(c[0]) === "cMediaNode");
      if (!cm) continue;
      let es = false;
      (function rec(m) { if (!Array.isArray(m) || es) return; if (loc(m[0]) === "spTgt" && A(m).bid === bid) es = true; H(m).forEach(rec); })(cm);
      if (!es) continue;
      const a = A(cm), ctn = H(cm).find((c) => Array.isArray(c) && loc(c[0]) === "cTn");
      return {
        vol: a.vol != null ? Math.max(0, Math.min(1, (+a.vol || 0) / 100000)) : 1,
        mudo: a.mute === "1" || a.mute === "true",
        repetir: !!(ctn && A(ctn).repeatCount === "indefinite"),
        ocultar: a.showWhenStopped === "0" || a.showWhenStopped === "false",
        diapos: a.numSld != null ? Math.max(1, +a.numSld || 1) : 1,
        pantalla: A(n).fullScrn === "1" || A(n).fullScrn === "true",
      };
    }
    return { vol: 1, mudo: false, repetir: false, ocultar: false, diapos: 1, pantalla: false };
  }

  // ------------------------------------------------------------ estilo (con las variables de nom.css)
  const ESTILO_ID = "nom-medios-diapo";
  function ponerEstilo() {
    if (typeof document === "undefined" || document.getElementById(ESTILO_ID)) return;
    const s = document.createElement("style");
    s.id = ESTILO_ID;
    s.setAttribute("data-nom", "");   // (nucleo.js quita las hojas de estilo que no son de NOM)
    s.textContent = `
.nm-medio{position:absolute;inset:0;overflow:hidden;pointer-events:none;visibility:hidden}
.nm-medio.nm-visible{visibility:visible}
.nm-medio video{position:absolute;left:0;top:0;width:100%;height:100%;object-fit:fill}
.nm-medio.nm-pantalla{z-index:40;background:#000}
.nm-medio.nm-pantalla video{object-fit:contain}
.bl.nm-callado>*{visibility:hidden!important}
.reproductor .bl.con-medio{cursor:pointer}
.nm-barra{position:absolute;z-index:45;display:flex;align-items:center;gap:8px;height:34px;padding:0 10px;box-sizing:border-box;border-radius:var(--r-s);background:var(--acento);color:var(--sobre-acento);font:600 13px/1 var(--fuente);transform-origin:0 0;opacity:0;pointer-events:none;transition:opacity .15s var(--suave)}
.nm-barra.nm-ver{opacity:1;pointer-events:auto}
.nm-barra button{all:unset;display:grid;place-items:center;width:26px;height:26px;border-radius:var(--r-xs);cursor:pointer}
.nm-barra button:focus-visible,.nm-pista:focus-visible{box-shadow:var(--foco)}
.nm-pista{position:relative;flex:1;height:6px;border-radius:3px;cursor:pointer;outline:none}
.nm-pista::before{content:"";position:absolute;inset:0;border-radius:3px;background:var(--sobre-acento);opacity:.3}
.nm-pista>i{position:absolute;left:0;top:0;bottom:0;border-radius:3px;background:var(--sobre-acento)}
.nm-pista>b{position:absolute;top:-3px;width:8px;height:12px;margin-left:-4px;border-radius:2px;background:var(--sobre-acento);box-shadow:0 0 0 2px var(--acento)}
.nm-tiempo{font-variant-numeric:tabular-nums;white-space:nowrap}
.nm-medio-marca{position:absolute;left:6px;bottom:6px;display:inline-flex;align-items:center;gap:4px;padding:2px 7px;border-radius:var(--r-xs);background:var(--sup);color:var(--tinta);box-shadow:var(--hundido);font:600 11px/1.5 var(--fuente);pointer-events:none}
`;
    document.head.append(s);
  }
  ponerEstilo();

  /**
   * Los medios de una presentación (uno por presentación). opc: { raiz (el reproductor: ahí viven los audios que siguen
   * sonando en las diapositivas siguientes), escala() (px de pantalla por px de la diapositiva) }.
   *   montar(el, p, k): los medios de la diapositiva p (la número k) ya pintada en el; lo de la anterior se detiene.
   *   comando(bid, cmd): lo que manda el motor de animaciones. esMedio(bid), alternar(bid): clic en el medio.
   *   tecla(e): las teclas de PowerPoint (true si era para un medio). cerrar(): al terminar la presentación.
   */
  function sesion(doc, opc) {
    ponerEstilo();
    opc = opc || {};
    const k0 = () => (opc.escala ? opc.escala() : 1) || 1;
    const urls = new Map();
    let og;
    const origen = () => {
      if (og === undefined) { try { og = NOM.origen ? NOM.origen.leer(doc) : null; } catch (e) { og = null; } }
      return og;
    };
    /** El medio sacado del archivo original (una sola vez por presentación). */
    function urlDe(ruta) {
      if (urls.has(ruta)) return urls.get(ruta);
      let u = null;
      try { const o = origen(), b = o && o.zip ? o.zip.bytes(ruta) : null; if (b && b.length) u = URL.createObjectURL(new Blob([b], { type: mimeDe(ruta) })); } catch (e) { u = null; }
      urls.set(ruta, u);
      return u;
    }
    const todos = new Map();   // bid → { bid, b, aj, tipo, el, cont, bl, barra, k, empezado }
    let raiz = null, actual = null, raf = 0;

    // ---- tiempos del medio (con su recorte)
    const ini = (r) => ((r.b.medio.recorte && r.b.medio.recorte.ini) || 0) / 1000;
    const fin = (r) => { const d = r.el.duration; if (!isFinite(d) || !(d > 0)) return Infinity; return Math.max(ini(r), d - ((r.b.medio.recorte && r.b.medio.recorte.fin) || 0) / 1000); };
    const largo = (r) => { const f = fin(r); return isFinite(f) ? f - ini(r) : 0; };

    function volumen(r) {
      const fd = r.b.medio.fundido, t = r.el.currentTime;
      let v = r.aj.vol;
      if (fd && fd.ent > 0) v *= Math.max(0, Math.min(1, (t - ini(r)) / (fd.ent / 1000)));
      if (fd && fd.sal > 0 && isFinite(fin(r))) v *= Math.max(0, Math.min(1, (fin(r) - t) / (fd.sal / 1000)));
      try { r.el.volume = Math.max(0, Math.min(1, v)); } catch (e) { /* iOS no deja cambiar el volumen */ }
    }
    function sonando() { for (const r of todos.values()) if (!r.el.paused) return true; return false; }
    /** Mientras algo suena: el recorte del final, repetir, el fundido y la barra. */
    function vigilar() {
      if (raf) return;
      const paso = () => {
        raf = 0;
        for (const r of todos.values()) {
          if (r.el.paused) continue;
          if (r.el.currentTime >= fin(r) - 0.03) { if (r.aj.repetir) { r.el.currentTime = ini(r); } else { r.el.pause(); terminado(r); continue; } }
          volumen(r);
          pintarBarra(r);
        }
        if (sonando()) raf = requestAnimationFrame(paso);
      };
      raf = requestAnimationFrame(paso);
    }
    function terminado(r) {
      // sin «repetir»: se queda en el último cuadro (como PowerPoint); en pantalla completa, vuelve a la diapositiva
      if (r.aj.pantalla) pantalla(r, false);
      estado(r);
    }
    /** Lo que se ve según si suena: el video encima de la portada, «ocultar mientras no se reproduce». */
    function estado(r) {
      if (r.bl) r.bl.classList.toggle("nm-callado", !!r.aj.ocultar && r.el.paused);
      pintarBarra(r);
    }
    function pantalla(r, si) {
      if (!r.cont || !raiz) return;
      if (si && !r.cont.classList.contains("nm-pantalla")) {
        const T = { w: raiz.offsetWidth || 960, h: raiz.offsetHeight || 540 };
        r.cont.classList.add("nm-pantalla");
        Object.assign(r.cont.style, { left: "0px", top: "0px", width: T.w + "px", height: T.h + "px", right: "auto", bottom: "auto" });
        raiz.append(r.cont);
      } else if (!si && r.cont.classList.contains("nm-pantalla")) {
        r.cont.classList.remove("nm-pantalla");
        Object.assign(r.cont.style, { left: "", top: "", width: "", height: "", right: "", bottom: "" });
        if (r.bl) r.bl.append(r.cont);
      }
    }

    // ---- reproducir, pausar, detener
    function cuandoListo(r, f) { if (r.el.readyState >= 1) f(); else r.el.addEventListener("loadedmetadata", f, { once: true }); }
    function reproducir(r, desdeS) {
      cuandoListo(r, () => {
        try {
          if (desdeS != null) r.el.currentTime = Math.min(ini(r) + desdeS, isFinite(fin(r)) ? fin(r) : Infinity);
          else if (!r.empezado || r.el.currentTime >= fin(r) - 0.05) r.el.currentTime = ini(r);
        } catch (e) { /* sin posición todavía */ }
        r.empezado = true;
        if (r.cont) r.cont.classList.add("nm-visible");
        if (r.aj.pantalla) pantalla(r, true);
        volumen(r);
        const pr = r.el.play();
        // (si el navegador no deja sonar sin un toque, se reproduce sin sonido antes que no reproducir)
        if (pr && pr.catch) pr.catch(() => { if (!r.el.muted) { r.el.muted = true; r.el.play().catch(() => {}); estado(r); } });
        actual = r;
        estado(r);
        vigilar();
        mostrarBarra(r, 2200);
      });
    }
    function pausar(r) { r.el.pause(); actual = r; estado(r); mostrarBarra(r, 2200); }
    function detener(r) {
      r.el.pause();
      cuandoListo(r, () => { try { r.el.currentTime = ini(r); } catch (e) { /* sin posición */ } });
      r.empezado = false;
      if (r.aj.pantalla) pantalla(r, false);
      if (r.cont) r.cont.classList.remove("nm-visible");
      estado(r);
    }
    function alternar(bid) {
      const r = todos.get(bid);
      if (!r) return false;
      if (r.el.paused) reproducir(r, null); else pausar(r);
      return true;
    }
    /** Un comando de PowerPoint (p:cmd type="call"): playFrom(s), play, togglePause, pause, resume, stop, seek(s). */
    function comando(bid, cmd) {
      const r = todos.get(bid);
      if (!r) return;
      const c = String(cmd || "").trim();
      let m;
      if ((m = /^playFrom\(\s*([-\d.]+)\s*\)$/i.exec(c))) reproducir(r, Math.max(0, +m[1] || 0));
      else if (/^(play|resume)$/i.test(c)) reproducir(r, null);
      else if (/^togglePause$/i.test(c)) { if (r.el.paused) reproducir(r, null); else pausar(r); }
      else if (/^pause$/i.test(c)) pausar(r);
      else if (/^stop$/i.test(c)) detener(r);
      else if ((m = /^seek\(\s*([-\d.]+)\s*\)$/i.exec(c))) cuandoListo(r, () => { try { r.el.currentTime = ini(r) + Math.max(0, +m[1] || 0); } catch (e) { /* sin posición */ } pintarBarra(r); });
    }

    // ---- la barra (reproducir/pausa, posición con marcadores, tiempo y sonido)
    function crearBarra(r) {
      const bar = document.createElement("div");
      bar.className = "nm-barra";
      bar.setAttribute("role", "group");
      bar.setAttribute("aria-label", r.tipo === "audio" ? "Controles del audio" : "Controles del video");
      const marcas = (r.b.medio.marcadores || []).map((m, i) => `<b data-m="${i}" title="${NOM.esc ? NOM.esc(m.n) : ""}"></b>`).join("");
      bar.innerHTML = `<button type="button" data-nm="tocar" aria-label="Reproducir">${ico("play")}</button><div class="nm-pista" data-nm="pista" role="slider" tabindex="0" aria-label="Posición" aria-valuemin="0" aria-valuenow="0"><i></i>${marcas}</div><span class="nm-tiempo">0:00</span><button type="button" data-nm="sonido" aria-label="Silenciar">${ico("volume-2")}</button>`;
      // (lo que se toca en la barra no avanza la presentación)
      ["pointerdown", "pointerup", "click", "dblclick"].forEach((ev) => bar.addEventListener(ev, (e) => e.stopPropagation()));
      bar.addEventListener("click", (e) => {
        const b = e.target.closest("[data-nm]");
        if (!b) return;
        if (b.dataset.nm === "tocar") alternar(r.bid);
        else if (b.dataset.nm === "sonido") { r.el.muted = !r.el.muted; r.aj.mudo = r.el.muted; pintarBarra(r); }
      });
      const pista = bar.querySelector(".nm-pista");
      const ir = (e) => {
        const q = pista.getBoundingClientRect(), f = Math.max(0, Math.min(1, (e.clientX - q.left) / (q.width || 1)));
        cuandoListo(r, () => { try { r.el.currentTime = ini(r) + f * largo(r); } catch (x) { /* sin posición */ } r.empezado = true; if (r.cont) r.cont.classList.add("nm-visible"); pintarBarra(r); });
      };
      pista.addEventListener("pointerdown", (e) => { pista.setPointerCapture(e.pointerId); ir(e); });
      pista.addEventListener("pointermove", (e) => { if (pista.hasPointerCapture(e.pointerId)) ir(e); });
      pista.addEventListener("keydown", (e) => {
        if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
        e.preventDefault(); e.stopPropagation();
        cuandoListo(r, () => { r.el.currentTime = Math.max(ini(r), Math.min(fin(r), r.el.currentTime + (e.key === "ArrowLeft" ? -5 : 5))); pintarBarra(r); });
      });
      bar.addEventListener("pointerenter", () => mostrarBarra(r, 0));
      bar.addEventListener("pointerleave", () => mostrarBarra(r, 900));
      return bar;
    }
    function colocarBarra(r) {
      if (!r.barra || !r.bl || !raiz) return;
      const k = k0(), x = r.bl.offsetLeft, y = r.bl.offsetTop, w = r.bl.offsetWidth, h = r.bl.offsetHeight;
      // la barra mide lo mismo en la pantalla a cualquier tamaño de la diapositiva (se escala al revés)
      const ancho = r.tipo === "audio" ? Math.max(260, w * k) : Math.max(200, w * k - 16), alto = 34;
      const left = r.tipo === "audio" ? x + w / 2 - ancho / k / 2 : x + 8 / k;
      const top = r.tipo === "audio" ? y + h + 6 / k : y + h - (alto + 8) / k;
      Object.assign(r.barra.style, { left: left + "px", top: top + "px", width: ancho + "px", transform: `scale(${1 / k})` });
    }
    function pintarBarra(r) {
      const bar = r.barra;
      if (!bar) return;
      const t = Math.max(0, r.el.currentTime - ini(r)), L = largo(r);
      bar.querySelector(".nm-pista>i").style.width = (L > 0 ? Math.min(100, (t / L) * 100) : 0) + "%";
      bar.querySelector(".nm-tiempo").textContent = L > 0 ? `${reloj(t)} / ${reloj(L)}` : reloj(t);
      const pista = bar.querySelector(".nm-pista");
      pista.setAttribute("aria-valuemax", String(Math.round(L)));
      pista.setAttribute("aria-valuenow", String(Math.round(t)));
      bar.querySelectorAll(".nm-pista>b").forEach((b) => { const m = (r.b.medio.marcadores || [])[+b.dataset.m]; b.style.left = (L > 0 && m ? Math.min(100, ((m.t / 1000 - ini(r)) / L) * 100) : 0) + "%"; });
      const tocar = bar.querySelector('[data-nm="tocar"]'), son = bar.querySelector('[data-nm="sonido"]');
      const juega = !r.el.paused;
      if (tocar.dataset.v !== String(juega)) { tocar.dataset.v = String(juega); tocar.innerHTML = ico(juega ? "pause" : "play"); tocar.setAttribute("aria-label", juega ? "Pausa" : "Reproducir"); }
      if (son.dataset.v !== String(r.el.muted)) { son.dataset.v = String(r.el.muted); son.innerHTML = ico(r.el.muted ? "volume-x" : "volume-2"); son.setAttribute("aria-label", r.el.muted ? "Activar el sonido" : "Silenciar"); }
    }
    function mostrarBarra(r, ms) {
      if (!r.barra) return;
      clearTimeout(r.tBarra);
      colocarBarra(r);
      pintarBarra(r);
      r.barra.classList.add("nm-ver");
      if (ms) r.tBarra = setTimeout(() => { if (!r.barra.matches(":hover") && !(r.bl && r.bl.matches(":hover"))) r.barra.classList.remove("nm-ver"); }, ms);
    }

    // ---- poner y quitar los medios de una diapositiva
    function quitar(r) {
      clearTimeout(r.tBarra);
      try { r.el.pause(); r.el.removeAttribute("src"); r.el.load(); } catch (e) { /* ya no está */ }
      if (r.cont) r.cont.remove();
      if (r.tipo === "audio") r.el.remove();
      if (r.barra) r.barra.remove();
      if (actual === r) actual = null;
    }
    function montar(el, p, k) {
      // lo de la diapositiva de antes se detiene, salvo el audio que sigue sonando en las siguientes (numSld)
      for (const [bid, r] of [...todos]) {
        const sigue = r.tipo === "audio" && !r.el.paused && k > r.k && k < r.k + r.aj.diapos;
        if (sigue) { if (r.barra) { r.barra.remove(); r.barra = null; } r.bl = null; continue; }
        quitar(r);
        todos.delete(bid);
      }
      raiz = el;
      const nodos = doc.valor("_doc", "animMedios:" + p) || [];
      el.querySelectorAll(".bl[data-id]").forEach((bl) => {
        const b = doc.bloque(bl.dataset.id);
        if (!b || !b.medio || !b.medio.ruta || todos.has(b.id)) return;
        const url = urlDe(b.medio.ruta);
        if (!url) return;   // sin el archivo original (o muy grande para guardarlo): se queda la portada
        const r = { bid: b.id, b, aj: ajustesDe(nodos, b.id), tipo: b.medio.tipo === "audio" ? "audio" : "video", k, bl, empezado: false };
        r.el = document.createElement(r.tipo);
        r.el.preload = "auto";
        r.el.src = url;
        r.el.muted = r.aj.mudo;
        if (r.tipo === "video") {
          r.el.playsInline = true;
          r.el.setAttribute("playsinline", "");
          r.el.disablePictureInPicture = true;
          r.cont = document.createElement("div");
          r.cont.className = "nm-medio";
          // el video toma el recorte de su portada (PowerPoint recorta igual las dos)
          const rc = b.recorte;
          if (rc && (rc.iz || rc.d || rc.s || rc.i)) {
            const fw = 1 - (rc.iz || 0) - (rc.d || 0), fh = 1 - (rc.s || 0) - (rc.i || 0);
            if (fw > 0.05 && fh > 0.05) Object.assign(r.el.style, { width: 100 / fw + "%", height: 100 / fh + "%", left: (-(rc.iz || 0) / fw) * 100 + "%", top: (-(rc.s || 0) / fh) * 100 + "%" });
          }
          r.cont.append(r.el);
          bl.append(r.cont);
        } else (opc.raiz || document.body).append(r.el);
        r.el.addEventListener("ended", () => { if (r.aj.repetir) { try { r.el.currentTime = ini(r); } catch (e) { /* sin posición */ } r.el.play().catch(() => {}); } else terminado(r); });
        r.el.addEventListener("pause", () => estado(r));
        r.el.addEventListener("play", () => { estado(r); vigilar(); });
        if (!(r.tipo === "audio" && r.aj.ocultar)) {
          r.barra = crearBarra(r);
          el.append(r.barra);
          bl.addEventListener("pointerenter", () => mostrarBarra(r, 0));
          bl.addEventListener("pointerleave", () => mostrarBarra(r, 900));
        }
        todos.set(b.id, r);
        estado(r);
      });
    }
    /** Las teclas de PowerPoint para el medio que se tocó al último (o el primero de la diapositiva). */
    function tecla(e) {
      if (!e.altKey || e.ctrlKey || e.metaKey) return false;
      const r = actual && todos.has(actual.bid) ? actual : [...todos.values()][0];
      if (!r) return false;
      const k = (e.key || "").toLowerCase();
      if (k === "p") { alternar(r.bid); return true; }
      if (k === "q") { detener(r); return true; }
      if (e.key === "ArrowUp" || e.key === "ArrowDown") {
        r.aj.vol = Math.max(0, Math.min(1, r.aj.vol + (e.key === "ArrowUp" ? 0.1 : -0.1)));
        volumen(r); mostrarBarra(r, 1500);
        return true;
      }
      if (e.key === "Home" || e.key === "End") {
        const ms = (r.b.medio.marcadores || []).map((m) => m.t / 1000);
        if (!ms.length) return false;
        cuandoListo(r, () => {
          const t = r.el.currentTime;
          const dest = e.key === "End" ? ms.find((x) => x > t + 0.05) : [...ms].reverse().find((x) => x < t - 0.5);
          if (dest != null) r.el.currentTime = dest;
          else if (e.key === "Home") r.el.currentTime = ini(r);
          r.empezado = true;
          if (r.cont) r.cont.classList.add("nm-visible");
          mostrarBarra(r, 1500);
        });
        return true;
      }
      return false;
    }
    function cerrar() {
      if (raf) cancelAnimationFrame(raf);
      raf = 0;
      for (const r of todos.values()) quitar(r);
      todos.clear();
      urls.forEach((u) => { if (u) URL.revokeObjectURL(u); });
      urls.clear();
    }
    return { montar, comando, alternar, esMedio: (bid) => todos.has(bid), tecla, cerrar };
  }

  NOM.mediosDiapo = { sesion, ajustesDe, mimeDe };
})();

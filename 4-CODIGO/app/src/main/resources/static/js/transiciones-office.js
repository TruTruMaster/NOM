/* NOM — transiciones de PowerPoint entre diapositivas.
 *
 * Se guarda la transición de cada diapositiva tal como la escribe PowerPoint ("transOffice:<página>": el elemento con
 * su prefijo — p:, p14:, p15:prstTrans, p159:morph —, sus atributos, la duración y el avance automático). Al presentar,
 * la diapositiva que sale queda en una capa (copia de lo que se veía) y la que entra en otra; cada transición mueve,
 * recorta o gira esas capas cuadro por cuadro. Las de borde suave usan el mismo modelo medido de los filtros de las
 * animaciones (animaciones-office.js): cada punto se descubre en su turno con un borde lineal.
 */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});
  const n2 = (v) => Math.round(v * 100) / 100;

  // ================================================================== leer
  /** El elemento p:transition (DOM) de una diapositiva, ya elegida la rama de mc:AlternateContent que entiende PowerPoint 2016. */
  function leer(sld) {
    if (!sld) return null;
    const hijos = (n) => [...(n ? n.childNodes : [])].filter((c) => c.nodeType === 1);
    let tr = null;
    for (const c of hijos(sld.documentElement || sld)) {
      if (c.localName === "transition") { tr = c; break; }
      if (c.localName === "AlternateContent") {
        const ch = hijos(c).find((x) => x.localName === "Choice") || hijos(c).find((x) => x.localName === "Fallback");
        const t = ch && hijos(ch).find((x) => x.localName === "transition");
        if (t) { tr = t; break; }
      }
    }
    if (!tr) return null;
    const a = (n) => tr.getAttribute(n);
    const ef = hijos(tr).find((c) => c.localName !== "sndAc" && c.localName !== "extLst");
    const at = {};
    if (ef) for (const x of ef.attributes) if (!/^xmlns/.test(x.name)) at[x.localName] = x.value;
    const dur14 = tr.getAttributeNS("http://schemas.microsoft.com/office/powerpoint/2010/main", "dur") || a("p14:dur");
    const spd = a("spd") || "fast";
    const r = {
      t: ef ? (ef.localName === "prstTrans" ? "prst:" + (at.prst || "fallOver") : ef.localName) : "corte0",
      a: at,
      dur: dur14 != null ? +dur14 : { slow: 1000, med: 750, fast: 500 }[spd] || 500,
      spd,
    };
    if (a("advTm") != null) r.advTm = +a("advTm");
    if (a("advClick") === "0" || a("advClick") === "false") r.sinClic = true;
    if (ef && ef.localName === "cut" && !at.thruBlk) r.dur = 0;
    // sonido de la transición: se recuerda su relación para tocarlo si el archivo lo trae
    const snd = hijos(tr).find((c) => c.localName === "sndAc");
    if (snd) {
      const st = snd.querySelector("*|stSnd, stSnd"), s = st && [...st.getElementsByTagNameNS("*", "snd")][0];
      if (s) r.snd = { rid: s.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed") || s.getAttribute("r:embed"), nombre: s.getAttribute("name") || "", bucle: st.getAttribute("loop") === "1" };
      if ([...snd.getElementsByTagNameNS("*", "endSnd")].length) r.sndFin = true;
    }
    return r;
  }
  /** La transición de NOM de antes ("trans:<página>": desvanecer, deslizar, empujar, zoom) como una de PowerPoint. */
  function deLegado(t) {
    return { desvanecer: { t: "fade", a: {}, dur: 700 }, deslizar: { t: "push", a: { dir: "u" }, dur: 700 }, empujar: { t: "push", a: { dir: "l" }, dur: 700 }, zoom: { t: "zoom", a: {}, dur: 700 } }[t] || null;
  }
  /** Nombres de PowerPoint en español (para el panel de transiciones). */
  const NOMBRES = {
    corte0: "Ninguna", cut: "Cortar", fade: "Desvanecer", push: "Empujar", wipe: "Barrido", split: "Dividir", reveal: "Revelar", randomBar: "Barras aleatorias",
    circle: "Forma (círculo)", diamond: "Forma (rombo)", plus: "Forma (más)", zoom: "Forma (zoom)", cover: "Cubrir", pull: "Descubrir", flash: "Destello",
    dissolve: "Disolver", checker: "Cuadros bicolores", blinds: "Persianas", comb: "Peine", strips: "Barras diagonales", wedge: "Reloj (cuña)",
    wheel: "Reloj (rueda)", wheelReverse: "Reloj (rueda inversa)", newsflash: "Titular", random: "Aleatoria", vortex: "Vórtice", switch: "Cambiar", flip: "Voltear",
    ripple: "Onda", honeycomb: "Panal", prism: "Cubo", doors: "Puertas", window: "Ventana", ferris: "Noria", gallery: "Galería", conveyor: "Cinta",
    pan: "Desplazamiento", glitter: "Brillo", warp: "Deformar", flythrough: "Atravesar volando", shred: "Triturar", morph: "Transformar",
    "prst:fallOver": "Caída", "prst:drape": "Cortina lateral", "prst:curtains": "Cortinas", "prst:wind": "Viento", "prst:prestige": "Prestigio",
    "prst:fracture": "Fractura", "prst:crush": "Aplastar", "prst:peelOff": "Despegar", "prst:pageCurlDouble": "Curvatura de página",
    "prst:pageCurlSingle": "Curvatura de página (una)", "prst:airplane": "Avión", "prst:origami": "Origami",
  };
  const nombre = (tr) => (tr ? NOMBRES[tr.t] || tr.t : "Ninguna");

  // ================================================================== máscaras (mismo modelo que los filtros de las animaciones)
  const G = 0.145;
  let lienzo = null;
  /** Máscara de alfa: campo(x, y) ∈ [0,1] = cuándo se descubre ese punto; p = avance. */
  function mascaraCampo(w, h, p, campo, clave, nitido) {
    const tope = nitido ? 200000 : 50000, k = Math.min(1, Math.sqrt(tope / Math.max(1, w * h)));
    const mw = Math.max(2, Math.round(w * k)), mh = Math.max(2, Math.round(h * k));
    const ck = clave + ":" + mw + "x" + mh;
    let D = cache.get(ck);
    if (!D) {
      D = new Float32Array(mw * mh);
      for (let j = 0; j < mh; j++) for (let i = 0; i < mw; i++) D[j * mw + i] = campo(((i + 0.5) / mw) * w, ((j + 0.5) / mh) * h, i, j, mw, mh);
      if (cache.size > 60) cache.clear();
      cache.set(ck, D);
    }
    if (!lienzo) lienzo = document.createElement("canvas");
    lienzo.width = mw; lienzo.height = mh;
    const g = lienzo.getContext("2d"), im = g.createImageData(mw, mh), px = im.data, Gx = nitido ? 1e-4 : G, P = p * (1 + Gx);
    for (let i = 0, n = mw * mh; i < n; i++) { let a = (P - D[i]) / Gx; a = a <= 0 ? 0 : a >= 1 ? 1 : a; px[i * 4] = px[i * 4 + 1] = px[i * 4 + 2] = 255; px[i * 4 + 3] = Math.round(a * 255); }
    g.putImageData(im, 0, 0);
    return `url("${lienzo.toDataURL()}")`;
  }
  const cache = new Map();
  const azar = (n, s) => { const o = Array.from({ length: n }, (_, i) => i); let x = s * 9301 + 49297; for (let i = n - 1; i > 0; i--) { x = (x * 16807) % 2147483647; const j = x % (i + 1); [o[i], o[j]] = [o[j], o[i]]; } return o; };
  const ponMascara = (el, url) => { el.style.webkitMaskImage = url || ""; el.style.maskImage = url || ""; el.style.webkitMaskSize = el.style.maskSize = url ? "100% 100%" : ""; el.style.webkitMaskRepeat = el.style.maskRepeat = url ? "no-repeat" : ""; };

  // ================================================================== curvas
  const suave = (p) => p * p * (3 - 2 * p);
  const salida = (p) => 1 - (1 - p) * (1 - p);
  const entrada = (p) => p * p;

  // ================================================================== reproducir
  /**
   * Corre la transición tr en el escenario (el elemento de la diapositiva, ya con la diapositiva NUEVA pintada adentro).
   * viejoHtml: el HTML de la diapositiva que sale y su estilo (fondo). T: {w, h}. Devuelve { saltar(), promesa }.
   */
  function jugar(escenario, viejo, tr, T, opc) {
    opc = opc || {};
    if (!tr || tr.t === "corte0" || !(tr.dur > 0) && tr.t !== "cut") return { saltar() {}, promesa: Promise.resolve() };
    if (tr.t === "random") tr = Object.assign({}, tr, ALEATORIAS[Math.floor(Math.random() * ALEATORIAS.length)]);
    const W = T.w, H = T.h;
    // capas: la nueva (los hijos actuales del escenario) y la vieja (copia de lo que se veía)
    const nueva = document.createElement("div");
    nueva.className = "tr-capa tr-nueva";
    const fondoNuevo = escenario.style.background;
    nueva.style.cssText = `position:absolute;left:0;top:0;width:${W}px;height:${H}px;background:${fondoNuevo};overflow:hidden;`;
    while (escenario.firstChild) nueva.appendChild(escenario.firstChild);
    const vieja = document.createElement("div");
    vieja.className = "tr-capa tr-vieja";
    vieja.style.cssText = `position:absolute;left:0;top:0;width:${W}px;height:${H}px;background:${viejo.fondo || "#fff"};overflow:hidden;pointer-events:none;`;
    vieja.innerHTML = viejo.html || "";
    const escena = document.createElement("div");
    escena.className = "tr-escena";
    escena.style.cssText = `position:absolute;left:0;top:0;width:${W}px;height:${H}px;overflow:hidden;background:#000;`;
    escenario.appendChild(escena);
    const a = tr.a || {}, dir = a.dir || "", dur = Math.max(1, tr.dur || 1);
    let extra = [];
    const dosCapas = (arriba) => { escena.append(arriba === "vieja" ? nueva : vieja, arriba === "vieja" ? vieja : nueva); };
    const ef = efectoDe(tr, { W, H, nueva, vieja, escena, dir, a, extra, dosCapas });
    let raf = 0, t0 = performance.now(), terminado = false, resolver;
    const promesa = new Promise((r) => (resolver = r));
    function fin() {
      if (terminado) return;
      terminado = true;
      if (raf) cancelAnimationFrame(raf);
      // la nueva vuelve a ser el contenido del escenario
      escena.remove();
      ["transform", "opacity", "clipPath", "webkitMaskImage", "maskImage", "filter"].forEach((k) => (nueva.style[k] = ""));
      while (nueva.firstChild) escenario.appendChild(nueva.firstChild);
      resolver();
    }
    function cuadro(p) {
      ef(Math.max(0, Math.min(1, p)));
    }
    function bucle() {
      const p = (performance.now() - t0) / dur;
      if (p >= 1) { cuadro(1); fin(); return; }
      cuadro(p);
      raf = requestAnimationFrame(bucle);
    }
    if (opc.manual) { cuadro(0); return { cuadro, fin, promesa, saltar: fin }; }
    cuadro(0);
    raf = requestAnimationFrame(bucle);
    return { saltar: fin, promesa, cuadro, fin };
  }
  const ALEATORIAS = [{ t: "fade", a: {} }, { t: "push", a: { dir: "u" } }, { t: "wipe", a: {} }, { t: "split", a: {} }, { t: "randomBar", a: {} }, { t: "circle", a: {} }, { t: "cover", a: {} }, { t: "pull", a: {} }];

  /** El cuadro de cada transición: devuelve f(p) que acomoda las capas en el avance p. */
  function efectoDe(tr, c) {
    const { W, H, nueva, vieja, escena, dir, a, dosCapas } = c;
    const negro = () => { const n = document.createElement("div"); n.style.cssText = `position:absolute;left:0;top:0;width:${W}px;height:${H}px;background:#000;pointer-events:none;`; escena.appendChild(n); return n; };
    const blanco = () => { const n = negro(); n.style.background = "#fff"; return n; };
    const sentido = (d, def) => ({ l: [-1, 0], r: [1, 0], u: [0, -1], d: [0, 1], lu: [-1, -1], ru: [1, -1], ld: [-1, 1], rd: [1, 1] }[d || def] || [-1, 0]);
    const mascaraNueva = (campo, clave, nitido) => (p) => { ponMascara(nueva, p >= 1 ? "" : mascaraCampo(W, H, p, campo, clave, nitido)); };
    switch (tr.t) {
      case "cut": {
        dosCapas("nueva");
        if (a.thruBlk === "1" || a.thruBlk === "true") { const n = negro(); return (p) => { n.style.opacity = p < 1 ? 1 : 0; nueva.style.opacity = p < 1 ? 0 : 1; }; }
        return () => {};
      }
      case "fade": {
        dosCapas("nueva");
        if (a.thruBlk === "1" || a.thruBlk === "true") {
          // a negro y de negro: la vieja se oscurece en la primera mitad y la nueva aparece en la segunda
          const n = negro(); escena.insertBefore(n, nueva);
          return (p) => { vieja.style.opacity = p < 0.5 ? 1 : 0; n.style.opacity = p < 0.5 ? p * 2 : 1; nueva.style.opacity = p < 0.5 ? 0 : (p - 0.5) * 2; };
        }
        return (p) => { nueva.style.opacity = p; };
      }
      case "flash": {
        dosCapas("nueva"); const b = blanco();
        return (p) => { b.style.opacity = p < 0.5 ? p * 2 : 2 - p * 2; nueva.style.opacity = p < 0.5 ? 0 : 1; };
      }
      case "push": case "cover": case "pull": {
        const [sx, sy] = sentido(dir, "l");
        if (tr.t === "pull") { dosCapas("vieja"); return (p) => { vieja.style.transform = `translate(${n2(sx * W * p)}px,${n2(sy * H * p)}px)`; }; }
        dosCapas("nueva");
        return (p) => {
          nueva.style.transform = `translate(${n2(-sx * W * (1 - p))}px,${n2(-sy * H * (1 - p))}px)`;
          if (tr.t === "push") vieja.style.transform = `translate(${n2(sx * W * p)}px,${n2(sy * H * p)}px)`;
        };
      }
      case "wipe": {
        dosCapas("nueva");
        const d = dir || "l";
        return mascaraNueva((x, y) => (d === "l" ? 1 - x / W : d === "r" ? x / W : d === "u" ? 1 - y / H : y / H), "wipe" + d);
      }
      case "split": {
        dosCapas("nueva");
        const vert = a.orient === "vert", sale = a.dir !== "in";
        return mascaraNueva((x, y) => { const q = vert ? Math.abs(x - W / 2) / (W / 2) : Math.abs(y - H / 2) / (H / 2); return sale ? q : 1 - q; }, "split" + vert + sale);
      }
      case "circle": dosCapas("nueva"); return mascaraNueva((x, y) => Math.hypot((x - W / 2) / (W / 2), (y - H / 2) / (H / 2)) / Math.SQRT2, "circle");
      case "diamond": dosCapas("nueva"); return mascaraNueva((x, y) => (Math.abs(x - W / 2) / (W / 2) + Math.abs(y - H / 2) / (H / 2)) / 2, "diamond");
      case "plus": dosCapas("nueva"); return mascaraNueva((x, y) => Math.min(Math.abs(x - W / 2) / (W / 2), Math.abs(y - H / 2) / (H / 2)), "plus");
      case "wedge": dosCapas("nueva"); return mascaraNueva((x, y) => Math.abs((Math.atan2(x - W / 2, -(y - H / 2)) * 180) / Math.PI) / 180, "wedge");
      case "wheel": case "wheelReverse": {
        dosCapas("nueva");
        const n = Math.max(1, Math.min(8, +(a.spokes || 4))), inv = tr.t === "wheelReverse";
        return mascaraNueva((x, y) => { let g = ((Math.atan2(x - W / 2, -(y - H / 2)) * 180) / Math.PI + 360) % 360; if (inv) g = 360 - g; return ((g * n) / 360) % 1; }, "wheel" + n + inv);
      }
      case "blinds": { dosCapas("nueva"); const v = dir === "vert"; return mascaraNueva((x, y) => (v ? (x * 6) / W : (y * 6) / H) % 1, "blinds" + v); }
      case "checker": {
        dosCapas("nueva");
        const v = dir === "vert";
        return mascaraNueva((x, y) => {
          if (v) { const c = Math.min(7, Math.floor((x * 8) / W)), P = H / 4, off = c % 2 ? P / 2 : 0; return (((y - off) / P) % 1 + 1) % 1; }
          const r = Math.min(7, Math.floor((y * 8) / H)), P = W / 6, off = r % 2 ? P / 2 : 0; return (((x - off) / P) % 1 + 1) % 1;
        }, "checker" + v);
      }
      case "randomBar": {
        dosCapas("nueva");
        const v = dir === "vert", n = v ? Math.round(W / 2) : Math.round(H / 2), o = azar(n, 5);
        return mascaraNueva((x, y) => o[Math.min(n - 1, Math.floor(((v ? x / W : y / H)) * n))] / (n - 1), "randomBar" + v, true);
      }
      case "dissolve": {
        dosCapas("nueva");
        const lado = 8, nx = Math.ceil(W / lado), ny = Math.ceil(H / lado), o = azar(nx * ny, 11);
        return mascaraNueva((x, y) => o[Math.min(ny - 1, Math.floor(y / lado)) * nx + Math.min(nx - 1, Math.floor(x / lado))] / (nx * ny - 1), "dissolve", true);
      }
      case "strips": {
        dosCapas("nueva");
        const d = dir || "lu", celda = 10, nx = Math.ceil(W / celda), ny = Math.ceil(H / celda);
        return mascaraNueva((x, y) => { const i = Math.floor(x / celda), j = Math.floor(y / celda); const ii = /^l/.test(d) ? nx - 1 - i : i, jj = /u$/.test(d) ? ny - 1 - j : j; return (ii + jj) / (nx + ny - 2); }, "strips" + d);
      }
      case "comb": {
        dosCapas("nueva");
        const v = dir === "vert", n = 16;
        return (p) => {
          // franjas que entran alternadas desde los dos lados
          const q = suave(p);
          const cs = [];
          for (let i = 0; i < n; i++) {
            const desde = i % 2 ? 1 : -1;
            if (v) { const x0 = (i * W) / n, x1 = ((i + 1) * W) / n, dy = desde * H * (1 - q); cs.push(`M${n2(x0)} ${n2(Math.max(0, dy))}H${n2(x1)}V${n2(Math.min(H, H + dy))}H${n2(x0)}Z`); }
            else { const y0 = (i * H) / n, y1 = ((i + 1) * H) / n, dx = desde * W * (1 - q); cs.push(`M${n2(Math.max(0, dx))} ${n2(y0)}H${n2(Math.min(W, W + dx))}V${n2(y1)}H${n2(Math.max(0, dx))}Z`); }
          }
          nueva.style.clipPath = p >= 1 ? "" : `path('${cs.join("")}')`;
        };
      }
      case "zoom": {
        dosCapas("nueva");
        const dentro = a.dir === "in";
        return (p) => { const q = suave(p); nueva.style.transformOrigin = "50% 50%"; nueva.style.transform = `scale(${dentro ? 3 - 2 * q : Math.max(0.001, q)})`; nueva.style.opacity = dentro ? q : 1; };
      }
      case "newsflash": {
        dosCapas("nueva");
        return (p) => { nueva.style.transformOrigin = "50% 50%"; nueva.style.transform = `rotate(${n2(720 * (1 - p))}deg) scale(${Math.max(0.001, p)})`; };
      }
      case "reveal": {
        // la vieja se desvanece mientras la nueva se corre desde un lado (con negro de por medio si thruBlk)
        dosCapas("nueva");
        const [sx] = sentido(dir, "l"), negroMedio = a.thruBlk === "1" || a.thruBlk === "true";
        const n = negroMedio ? negro() : null; if (n) escena.insertBefore(n, nueva);
        return (p) => {
          const q = salida(p);
          vieja.style.opacity = negroMedio ? Math.max(0, 1 - p * 2) : 1 - p;
          vieja.style.transform = `translate(${n2(sx * W * 0.25 * p)}px,0)`;
          nueva.style.opacity = negroMedio ? Math.max(0, p * 2 - 1) : p;
          nueva.style.transform = `translate(${n2(-sx * W * 0.25 * (1 - q))}px,0)`;
          if (n) n.style.opacity = 1;
        };
      }
      // ---------------------------------------------------------------- en 3D
      case "prism": case "flip": case "switch": case "doors": case "window": case "gallery": case "conveyor": case "ferris": case "pan": case "flythrough":
      case "prst:fallOver": case "prst:drape": case "prst:curtains": case "prst:prestige": case "prst:peelOff": case "prst:pageCurlDouble": case "prst:pageCurlSingle": case "prst:airplane": case "prst:origami": case "prst:wind":
        return efecto3d(tr, c);
      // ---------------------------------------------------------------- en piezas
      case "vortex": case "ripple": case "honeycomb": case "glitter": case "shred": case "warp": case "prst:fracture": case "prst:crush":
        return efectoPiezas(tr, c);
      case "morph": return efectoMorph(tr, c);
      default:
        dosCapas("nueva");
        return (p) => { nueva.style.opacity = p; };
    }
  }

  // ================================================================== transiciones en 3D (CSS)
  function efecto3d(tr, c) {
    const { W, H, nueva, vieja, escena, dir, a, dosCapas } = c;
    escena.style.perspective = `${Math.round(Math.max(W, H) * 2.2)}px`;
    escena.style.perspectiveOrigin = "50% 50%";
    const inv = a.invX === "1" || a.invX === "true";
    const lado = /^(r)$/.test(dir) || (dir === "" && inv) ? 1 : dir === "u" ? 2 : dir === "d" ? 3 : 0;  // 0 izq, 1 der, 2 arriba, 3 abajo
    const caras = () => { [vieja, nueva].forEach((x) => { x.style.backfaceVisibility = "hidden"; x.style.webkitBackfaceVisibility = "hidden"; }); };
    switch (tr.t) {
      case "prism": {
        // cubo (y caja, girar, órbita: con el contenido o al revés)
        dosCapas("nueva"); caras();
        const horiz = lado < 2, sg = lado === 1 || lado === 3 ? -1 : 1;
        const caja = a.isInverted === "1", contenido = a.isContent === "1";
        const z = (horiz ? W : H) / 2;
        return (p) => {
          const q = suave(p), g = 90 * q * sg * (caja ? -1 : 1);
          const ax = horiz ? "rotateY" : "rotateX", sgx = horiz ? 1 : -1;
          const prof = contenido ? 0 : z;
          vieja.style.transformOrigin = nueva.style.transformOrigin = `50% 50% ${caja ? z : -z}px`;
          vieja.style.transform = `translateZ(${caja ? -prof : 0}px) ${ax}(${n2(sgx * g)}deg)`;
          nueva.style.transform = `translateZ(${caja ? -prof : 0}px) ${ax}(${n2(sgx * (g - 90 * sg * (caja ? -1 : 1)))}deg)`;
        };
      }
      case "flip": {
        dosCapas("nueva"); caras();
        const sg = dir === "r" ? -1 : 1;
        return (p) => {
          const q = suave(p), g = 180 * q * sg;
          const esc = 1 - 0.25 * Math.sin(Math.PI * p);
          vieja.style.transform = `scale(${n2(esc)}) rotateY(${n2(g)}deg)`;
          nueva.style.transform = `scale(${n2(esc)}) rotateY(${n2(g - 180 * sg)}deg)`;
        };
      }
      case "switch": {
        dosCapas("nueva");
        const sg = dir === "r" ? 1 : -1;
        return (p) => {
          const q = suave(p), arriba = q < 0.5;
          escena.append(arriba ? nueva : vieja, arriba ? vieja : nueva);
          const s = Math.sin(Math.PI * q);
          vieja.style.transform = `translateX(${n2(sg * -W * 0.45 * s)}px) rotateY(${n2(sg * 25 * s)}deg) scale(${n2(1 - 0.15 * s)})`;
          nueva.style.transform = `translateX(${n2(sg * W * 0.45 * s)}px) rotateY(${n2(sg * -25 * s)}deg) scale(${n2(1 - 0.15 * s)})`;
        };
      }
      case "doors": case "window": {
        // la vieja se abre en dos hojas (puertas) o la nueva cierra (ventana)
        dosCapas("vieja");
        const vert = dir === "vert";
        const hojas = [0, 1].map((k) => {
          const h = document.createElement("div");
          h.style.cssText = `position:absolute;overflow:hidden;${vert ? `left:0;width:${W}px;top:${k ? H / 2 : 0}px;height:${H / 2}px` : `top:0;height:${H}px;left:${k ? W / 2 : 0}px;width:${W / 2}px`};transform-origin:${vert ? (k ? "50% 100%" : "50% 0%") : k ? "100% 50%" : "0% 50%"};backface-visibility:hidden;`;
          const copia = vieja.cloneNode(true);
          copia.style.left = vert ? "0px" : `${k ? -W / 2 : 0}px`; copia.style.top = vert ? `${k ? -H / 2 : 0}px` : "0px";
          h.appendChild(copia); escena.appendChild(h); return h;
        });
        vieja.style.visibility = "hidden";
        return (p) => {
          const q = suave(p);
          hojas.forEach((h, k) => { const s = k ? 1 : -1; h.style.transform = vert ? `rotateX(${n2(-s * 100 * q)}deg)` : `rotateY(${n2(s * 100 * q)}deg)`; h.style.opacity = 1 - q * 0.6; });
          nueva.style.transform = `translateZ(${n2(-Math.max(W, H) * 0.6 * (1 - q))}px)`;
        };
      }
      case "gallery": case "conveyor": case "pan": {
        dosCapas("nueva");
        const sg = dir === "r" ? 1 : -1, vert = dir === "u" || dir === "d", sv = dir === "d" ? 1 : -1;
        return (p) => {
          const q = suave(p);
          if (tr.t === "pan") {
            if (vert) { vieja.style.transform = `translateY(${n2(sv * -H * q)}px)`; nueva.style.transform = `translateY(${n2(sv * H * (1 - q))}px)`; }
            else { vieja.style.transform = `translateX(${n2(sg * W * q)}px)`; nueva.style.transform = `translateX(${n2(-sg * W * (1 - q))}px)`; }
            return;
          }
          const g = tr.t === "gallery" ? 50 : 30;
          vieja.style.transform = `translateX(${n2(sg * W * 1.05 * q)}px) rotateY(${n2(sg * g * q)}deg) scale(${n2(1 - 0.2 * Math.sin(Math.PI * q))})`;
          nueva.style.transform = `translateX(${n2(-sg * W * 1.05 * (1 - q))}px) rotateY(${n2(-sg * g * (1 - q))}deg) scale(${n2(1 - 0.2 * Math.sin(Math.PI * q))})`;
        };
      }
      case "ferris": {
        dosCapas("nueva");
        const sg = dir === "r" ? 1 : -1;
        return (p) => {
          const q = suave(p);
          vieja.style.transformOrigin = nueva.style.transformOrigin = `50% ${n2(H * 2)}px`;
          vieja.style.transform = `rotate(${n2(sg * 40 * q)}deg)`;
          nueva.style.transform = `rotate(${n2(-sg * 40 * (1 - q))}deg)`;
        };
      }
      case "flythrough": {
        dosCapas("nueva");
        const sale = a.dir === "out", rebote = a.hasBounce === "1";
        return (p) => {
          const q = rebote ? 1 - Math.cos(p * Math.PI * 1.5) * Math.exp(-p * 3) * (1 - p) : suave(p);
          vieja.style.transform = `scale(${n2(sale ? 1 - 0.6 * q : 1 + 1.5 * q)})`; vieja.style.opacity = 1 - q;
          nueva.style.transform = `scale(${n2(sale ? 1.6 - 0.6 * q : 0.4 + 0.6 * q)})`; nueva.style.opacity = q;
        };
      }
      case "prst:fallOver": {
        dosCapas("vieja");
        return (p) => { const q = entrada(p); vieja.style.transformOrigin = inv ? "100% 100%" : "0% 100%"; vieja.style.transform = `rotateX(${n2(85 * q)}deg) rotateZ(${n2((inv ? -1 : 1) * 10 * q)}deg)`; vieja.style.opacity = 1 - Math.max(0, (p - 0.7) / 0.3); };
      }
      case "prst:drape": case "prst:curtains": {
        dosCapas("vieja");
        const dos = tr.t === "prst:curtains";
        return (p) => {
          const q = suave(p);
          vieja.style.transformOrigin = dos ? "50% 0%" : inv ? "100% 0%" : "0% 0%";
          vieja.style.transform = dos ? `scaleX(${n2(1 - q)})` : `skewY(${n2((inv ? 1 : -1) * 15 * q)}deg) scaleX(${n2(1 - q)})`;
          vieja.style.opacity = 1 - q * 0.3;
        };
      }
      case "prst:prestige": {
        dosCapas("vieja");
        return (p) => { const q = suave(p); vieja.style.transformOrigin = "50% 0%"; vieja.style.transform = `rotateX(${n2(-70 * q)}deg) translateZ(${n2(200 * q)}px)`; vieja.style.opacity = 1 - q; };
      }
      case "prst:peelOff": case "prst:pageCurlDouble": case "prst:pageCurlSingle": {
        dosCapas("vieja");
        return (p) => { const q = suave(p); vieja.style.transformOrigin = inv ? "0% 50%" : "100% 50%"; vieja.style.transform = `rotateY(${n2((inv ? 1 : -1) * 120 * q)}deg) translateZ(${n2(80 * Math.sin(Math.PI * q))}px)`; vieja.style.opacity = 1 - Math.max(0, (q - 0.6) / 0.4); };
      }
      case "prst:airplane": case "prst:origami": {
        dosCapas("vieja");
        return (p) => {
          const q = suave(p), s = inv ? -1 : 1;
          const doblar = Math.min(1, q * 2), volar = Math.max(0, q * 2 - 1);
          vieja.style.transformOrigin = "50% 50%";
          vieja.style.transform = `translate(${n2(s * W * 1.2 * volar)}px,${n2(-H * 0.6 * volar)}px) rotateX(${n2(60 * doblar)}deg) scale(${n2(1 - 0.6 * doblar)})`;
          vieja.style.opacity = 1 - volar;
        };
      }
      case "prst:wind": {
        dosCapas("vieja");
        return (p) => { const q = entrada(p), s = inv ? -1 : 1; vieja.style.transform = `translateX(${n2(s * W * 1.3 * q)}px) skewX(${n2(-s * 25 * q)}deg) rotateY(${n2(s * 30 * q)}deg)`; vieja.style.opacity = 1 - q; };
      }
      default: dosCapas("nueva"); return (p) => { nueva.style.opacity = p; };
    }
  }

  // ================================================================== transiciones en piezas
  /** La diapositiva que sale partida en piezas (copias recortadas) que se mueven, giran o se desvanecen. */
  function efectoPiezas(tr, c) {
    const { W, H, nueva, vieja, escena, dir, a, dosCapas } = c;
    dosCapas("vieja");
    vieja.style.visibility = "hidden";
    escena.style.perspective = `${Math.round(Math.max(W, H) * 2.2)}px`;
    const forma = tr.t === "shred" ? (a.pattern === "rectangle" ? "rect" : "tiras") : tr.t === "honeycomb" || (tr.t === "glitter" && a.pattern === "hexagon") ? "hex" : "cuad";
    const nx = forma === "tiras" ? 24 : tr.t === "prst:crush" ? 10 : 12, ny = forma === "tiras" ? 1 : Math.round((nx * H) / W);
    const cw = W / nx, ch = H / ny, piezas = [];
    const orden = azar(nx * ny, 3);
    for (let j = 0; j < ny; j++) for (let i = 0; i < nx; i++) {
      const pz = document.createElement("div");
      pz.style.cssText = `position:absolute;left:${n2(i * cw)}px;top:${n2(j * ch)}px;width:${n2(cw + 0.5)}px;height:${n2(ch + 0.5)}px;overflow:hidden;`;
      const copia = vieja.cloneNode(true);
      copia.style.visibility = "visible"; copia.style.left = n2(-i * cw) + "px"; copia.style.top = n2(-j * ch) + "px";
      pz.appendChild(copia); escena.appendChild(pz);
      const cx = (i + 0.5) / nx - 0.5, cy = (j + 0.5) / ny - 0.5;
      piezas.push({ el: pz, i, j, cx, cy, r: orden[j * nx + i] / (nx * ny), dist: Math.hypot(cx, cy) });
    }
    const sx = dir === "r" ? 1 : dir === "l" || !dir ? -1 : 0, sy = dir === "d" ? 1 : dir === "u" ? -1 : 0;
    return (p) => {
      piezas.forEach((pz) => {
        let tr2 = "", op = 1;
        switch (tr.t) {
          case "vortex": { const q = Math.max(0, Math.min(1, (p * 1.6 - pz.r * 0.6))); tr2 = `translate(${n2((sx || 1) * W * 0.8 * q + pz.cx * W * q)}px,${n2(sy * H * 0.8 * q + pz.cy * H * q)}px) rotate(${n2(540 * q * (pz.r - 0.5))}deg) scale(${n2(1 - 0.5 * q)})`; op = 1 - q; break; }
          case "ripple": { const q = Math.max(0, Math.min(1, p * 1.8 - pz.dist * 1.2)); tr2 = `translateZ(${n2(-200 * Math.sin(Math.PI * q))}px) scale(${n2(1 - 0.2 * q)})`; op = 1 - q; break; }
          case "honeycomb": case "glitter": { const q = Math.max(0, Math.min(1, p * 1.7 - pz.r * 0.7)); tr2 = `rotateY(${n2(180 * q)}deg) scale(${n2(1 - 0.3 * Math.sin(Math.PI * q))})`; op = q < 0.5 ? 1 : 0; break; }
          case "shred": { const q = Math.max(0, Math.min(1, p * 1.5 - (pz.i % 2) * 0.25)); tr2 = `translateY(${n2((pz.i % 2 ? -1 : 1) * H * 1.1 * q)}px)`; break; }
          case "warp": { const q = suave(p); const sale = a.dir !== "in"; tr2 = `translate(${n2(pz.cx * W * (sale ? 2 : -0.8) * q)}px,${n2(pz.cy * H * (sale ? 2 : -0.8) * q)}px) scale(${n2(sale ? 1 + q : 1 - q * 0.9)})`; op = 1 - q; break; }
          case "prst:fracture": { const q = Math.max(0, Math.min(1, p * 1.5 - pz.r * 0.5)); tr2 = `translate(${n2(pz.cx * W * 0.6 * q)}px,${n2(H * 0.9 * q * q)}px) rotate(${n2(90 * q * (pz.r - 0.5))}deg)`; op = 1 - q; break; }
          case "prst:crush": { const q = suave(p); tr2 = `translate(${n2(-pz.cx * W * q)}px,${n2(-pz.cy * H * q)}px) scale(${n2(1 - q)}) rotate(${n2(200 * q * (pz.r - 0.5))}deg)`; op = 1 - q * q; break; }
          default: op = 1 - p;
        }
        pz.el.style.transform = tr2; pz.el.style.opacity = op;
      });
      nueva.style.opacity = Math.min(1, p * 1.5);
    };
  }

  // ================================================================== transformar (morph)
  /**
   * Transformar: empareja los objetos de las dos diapositivas (primero por nombre — «!!» al principio fuerza la pareja —,
   * después por tipo y texto, y por forma y lugar) y lleva cada uno de su lugar, tamaño, giro y color al del otro.
   * Lo que no tiene pareja entra o sale desvaneciéndose. Por palabra o por letra, el texto se empareja así también.
   */
  function efectoMorph(tr, c) {
    const { W, H, nueva, vieja, escena, dosCapas } = c;
    dosCapas("nueva");
    const blsV = [...vieja.querySelectorAll(":scope > .bl[data-id]")], blsN = [...nueva.querySelectorAll(":scope > .bl[data-id]")];
    const caja = (el) => ({ x: el.offsetLeft, y: el.offsetTop, w: el.offsetWidth || 1, h: el.offsetHeight || 1, r: parseFloat((/rotate\((-?[\d.]+)deg\)/.exec(el.style.transform) || [])[1] || 0) });
    const firma = (el) => {
      const n = el.dataset.nombre || "", t = (el.textContent || "").trim().slice(0, 80);
      const tipo = el.querySelector(".libre-texto") ? "texto" : el.querySelector("img") ? "img:" + (el.querySelector("img").getAttribute("src") || "").slice(-40) : el.querySelector(".forma") ? "forma:" + (el.querySelector(".forma").dataset.figura || "") : "otro";
      return { n, t, tipo };
    };
    const fv = blsV.map(firma), fn = blsN.map(firma), usadoN = new Set(), pares = [];
    // 1) por nombre «!!»; 2) mismo tipo y mismo texto; 3) mismo tipo (más cercano)
    const pasos = [(i, j) => fv[i].n && fv[i].n === fn[j].n && /^!!/.test(fv[i].n), (i, j) => fv[i].tipo === fn[j].tipo && fv[i].t && fv[i].t === fn[j].t, (i, j) => fv[i].tipo === fn[j].tipo && !fv[i].t && !fn[j].t, (i, j) => fv[i].tipo === fn[j].tipo && fv[i].tipo === "texto"];
    const usadoV = new Set();
    for (const regla of pasos) for (let i = 0; i < blsV.length; i++) {
      if (usadoV.has(i)) continue;
      let mejor = -1, dmin = Infinity;
      for (let j = 0; j < blsN.length; j++) {
        if (usadoN.has(j) || !regla(i, j)) continue;
        const a = caja(blsV[i]), b = caja(blsN[j]), dd = Math.hypot(a.x - b.x, a.y - b.y);
        if (dd < dmin) { dmin = dd; mejor = j; }
      }
      if (mejor >= 0) { usadoV.add(i); usadoN.add(mejor); pares.push([i, mejor]); }
    }
    // los de la vieja con pareja se esconden: se mueve el de la nueva desde el lugar del viejo
    const mov = pares.map(([i, j]) => {
      const a = caja(blsV[i]), b = caja(blsN[j]), el = blsN[j];
      blsV[i].style.visibility = "hidden";
      const colV = getComputedStyle(blsV[i].querySelector(".forma-svg path") || blsV[i]).fill;
      return { el, a, b, colV, colN: getComputedStyle(blsN[j].querySelector(".forma-svg path") || blsN[j]).fill, pathN: blsN[j].querySelector(".forma-svg path"), base: el.style.transform };
    });
    const soloN = blsN.filter((_, j) => !usadoN.has(j)), soloV = blsV.filter((_, i) => !usadoV.has(i));
    // el fondo: la nueva capa (con su fondo) entra desvaneciéndose y los objetos se mueven encima
    nueva.style.background = "transparent";
    const fondo = document.createElement("div"); fondo.style.cssText = `position:absolute;left:0;top:0;width:${W}px;height:${H}px;background:${c.nueva.dataset.fondo || ""};`;
    return (p) => {
      const q = suave(p);
      mov.forEach((m) => {
        const x = m.a.x + (m.b.x - m.a.x) * q, y = m.a.y + (m.b.y - m.a.y) * q, w = m.a.w + (m.b.w - m.a.w) * q, h = m.a.h + (m.b.h - m.a.h) * q, r = m.a.r + (m.b.r - m.a.r) * q;
        const sx = w / m.b.w, sy = h / m.b.h, dx = x + w / 2 - (m.b.x + m.b.w / 2), dy = y + h / 2 - (m.b.y + m.b.h / 2);
        m.el.style.transform = p >= 1 ? m.base : `translate(${n2(dx)}px,${n2(dy)}px) rotate(${n2(r)}deg) scale(${n2(sx)},${n2(sy)})`;
        if (m.pathN && m.colV && m.colN && m.colV !== m.colN) m.pathN.style.fill = p >= 1 ? "" : mezclaCss(m.colV, m.colN, q);
      });
      soloN.forEach((el) => { el.style.opacity = q; });
      soloV.forEach((el) => { el.style.opacity = 1 - q; });
      vieja.style.opacity = 1;
      void fondo;
    };
  }
  function mezclaCss(a, b, q) {
    const p = (s) => (/rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(s) || [0, 0, 0, 0]).slice(1).map(Number);
    const x = p(a), y = p(b);
    return `rgb(${x.map((v, i) => Math.round(v + (y[i] - v) * q)).join(",")})`;
  }

  NOM.transOffice = { leer, jugar, deLegado, nombre, NOMBRES, mascaraCampo };
})();

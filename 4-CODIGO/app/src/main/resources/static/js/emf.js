/* NOM · Imágenes EMF y WMF (los «metarchivos» de Windows) dibujadas como SVG.

   Word, Excel y PowerPoint guardan así muchos dibujos: las ecuaciones del Editor de ecuaciones 3.0 y de MathType (su
   vista previa es un WMF), los objetos incrustados, las imágenes pegadas desde otros programas de Windows. El navegador
   no sabe mostrarlos; aquí se leen sus registros (MS-EMF y MS-WMF: el dibujo de GDI) y se escribe un SVG equivalente.
   El documento guarda el archivo original (al descargar sale idéntico); solo al dibujarlo se convierte (binarios.js).

   Lo que se dibuja: líneas, polígonos, curvas, rectángulos, elipses, arcos, trayectorias (rellenas, con contorno o como
   recorte), texto con su fuente, tamaño, giro, posiciones de cada letra y alineación (también la fuente Symbol, que se
   pasa a Unicode), mapas de bits (DIB, también como máscaras), tramas, degradados, transformaciones del mundo, modos
   de mapeo, ventana y área de visualización, guardar y restaurar el estado, y recortes por rectángulo y por región.
   Los EMF+ que también traen su dibujo GDI («dobles», lo normal en Office) se dibujan con el GDI. */
(function () {
  "use strict";
  const NOM = window.NOM || (window.NOM = {});

  // ------------------------------------------------------------ Symbol → Unicode (de 0x20 a 0xFE)
  const SYMBOL = (() => {
    const t = {};
    const s = " !∀#∃%&∋()∗+,−./0123456789:;<=>?≅ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_‾αβχδεφγηιϕκλμνοπθρστυϖωξψζ{|}∼";
    for (let i = 0; i < s.length; i++) t[0x20 + i] = s[i];
    const a = "€ϒ′≤⁄∞ƒ♣♦♥♠↔←↑→↓°±″≥×∝∂•÷≠≡≈…⏐⎯↵ℵℑℜ℘⊗⊕∅∩∪⊃⊇⊄⊂⊆∈∉∠∇®©™∏√⋅¬∧∨⇔⇐⇑⇒⇓◊〈®©™∑⎛⎜⎝⎡⎢⎣⎧⎨⎩⎪ 〉∫⌠⎮⌡⎞⎟⎠⎤⎥⎦⎫⎬⎭";
    for (let i = 0; i < a.length; i++) t[0xA0 + i] = a[i];
    return t;
  })();
  // MT Extra (MathType y el Editor de ecuaciones): los puntos suspensivos y unas cuantas letras
  const MTEXTRA = { 0x26: "˙", 0x4C: "⋯", 0x4D: "⋮", 0x4E: "⋰", 0x4F: "⋱", 0x6C: "ℓ", 0x68: "ℏ", 0x3C: "◃", 0x3E: "▹" };
  // Marlett (las piezas de los controles de Windows): el marco de una casilla con sus sombras; la cara no se pinta
  const MARLETT = { c: "□", d: "□", e: "□", f: "□", g: "" };
  const deMarlett = (s) => Array.from(s, (c) => (c in MARLETT ? MARLETT[c] : c)).join("");
  /** Ancho medio de las letras (fracción del em) de las fuentes más comunes: GDI lo usa con lfWidth. */
  const ANCHO_MEDIO = { arial: 0.441, "arial narrow": 0.362, "times new roman": 0.401, "courier new": 0.6, calibri: 0.43, cambria: 0.45, verdana: 0.508, tahoma: 0.45, "trebuchet ms": 0.43, symbol: 0.42, georgia: 0.45, "comic sans ms": 0.47, "segoe ui": 0.46, system: 0.45 };
  const deMtExtra = (s) => Array.from(s, (c) => { let k = c.codePointAt(0); if (k >= 0xF020 && k <= 0xF0FF) k -= 0xF000; return MTEXTRA[k] || c; }).join("");
  const deSymbol = (s) => Array.from(s, (c) => { let k = c.codePointAt(0); if (k >= 0xF020 && k <= 0xF0FF) k -= 0xF000; return SYMBOL[k] || c; }).join("");
  const CODIFICACION = { 0: "windows-1252", 1: "windows-1252", 161: "windows-1253", 162: "windows-1254", 177: "windows-1255", 178: "windows-1256", 186: "windows-1257", 204: "windows-1251", 238: "windows-1250", 163: "windows-1258", 222: "windows-874", 128: "shift_jis", 129: "euc-kr", 134: "gbk", 136: "big5" };
  const decodificadores = {};
  function ansi(bytes, charset) {
    const cod = CODIFICACION[charset] || "windows-1252";
    try { return (decodificadores[cod] = decodificadores[cod] || new TextDecoder(cod)).decode(bytes); }
    catch (e) { return String.fromCharCode.apply(null, bytes); }
  }

  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const n3 = (v) => (Math.round(v * 1000) / 1000).toString();
  const color = (cr) => "#" + [cr & 0xFF, (cr >>> 8) & 0xFF, (cr >>> 16) & 0xFF].map((v) => v.toString(16).padStart(2, "0")).join("");
  const B64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };

  /** Fuentes de Windows → lo que tiene el equipo (el SVG dentro de un <img> no puede usar las fuentes de la página). */
  function familia(nombre) {
    const n = String(nombre || "").replace(/\0.*$/, "").trim();
    const l = n.toLowerCase();
    const serif = /times|roman|georgia|cambria|garamond|book|century|serif|minion|palatino|symbol|mt extra|math/.test(l);
    const mono = /courier|consolas|mono|fixed|lucida console/.test(l);
    const alt = mono ? "'Courier New','Liberation Mono',monospace" : serif ? "'Times New Roman','Liberation Serif',Tinos,serif" : "Arial,'Liberation Sans',Arimo,Helvetica,sans-serif";
    return (n && !/^symbol$/i.test(n) ? `'${n.replace(/'/g, "")}',` : "") + alt;
  }

  // ------------------------------------------------------------ mapas de bits (DIB → BMP que el navegador sí lee)
  /** Tamaño de la cabecera más la tabla de colores de un DIB que empieza en `o`. */
  function tamCabeceraDib(dv, o) {
    const tam = dv.getUint32(o, true);
    if (tam === 12) { const bpp = dv.getUint16(o + 10, true); return 12 + (bpp <= 8 ? (1 << bpp) * 3 : 0); }
    const bpp = dv.getUint16(o + 14, true), comp = dv.getUint32(o + 16, true), usados = dv.getUint32(o + 32, true);
    let t = tam + (bpp <= 8 ? (usados || (1 << bpp)) * 4 : usados * 4);
    if ((comp === 3 || comp === 6) && tam === 40) t += comp === 6 ? 16 : 12;
    return t;
  }
  /** Un DIB (cabecera y bits, quizá separados) → data URL. null si no se entiende. */
  function dibAUrl(u8, oCab, tamCab, oBits, tamBits) {
    try {
      const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
      if (oCab + 16 > u8.length) return null;
      const ancho = dv.getInt32(oCab + 4, true), alto = dv.getInt32(oCab + 8, true);
      const comp = dv.getUint32(oCab, true) >= 40 ? dv.getUint32(oCab + 16, true) : 0;
      if (oBits == null) { tamCab = tamCabeceraDib(dv, oCab); oBits = oCab + tamCab; }
      if (tamBits == null || tamBits <= 0) {
        const bpp = dv.getUint16(oCab + (dv.getUint32(oCab, true) === 12 ? 10 : 14), true);
        tamBits = comp === 4 || comp === 5 || comp === 1 || comp === 2 ? (dv.getUint32(oCab + 20, true) || u8.length - oBits) : Math.ceil((Math.abs(ancho) * bpp) / 32) * 4 * Math.abs(alto);
      }
      tamBits = Math.min(tamBits, u8.length - oBits);
      if (tamBits <= 0) return null;
      const bits = u8.subarray(oBits, oBits + tamBits);
      if (comp === 4) return { url: "data:image/jpeg;base64," + B64(bits), w: Math.abs(ancho), h: Math.abs(alto) };
      if (comp === 5) return { url: "data:image/png;base64," + B64(bits), w: Math.abs(ancho), h: Math.abs(alto) };
      const cab = u8.subarray(oCab, oCab + tamCab);
      const f = new Uint8Array(14 + cab.length + bits.length);
      const fv = new DataView(f.buffer);
      f[0] = 0x42; f[1] = 0x4D; fv.setUint32(2, f.length, true); fv.setUint32(10, 14 + cab.length, true);
      f.set(cab, 14); f.set(bits, 14 + cab.length);
      return { url: "data:image/bmp;base64," + B64(f), w: Math.abs(ancho), h: Math.abs(alto), abajoArriba: alto > 0 };
    } catch (e) { return null; }
  }

  // ------------------------------------------------------------ el lienzo de GDI (común a EMF y WMF)
  const PS = { SOLID: 0, DASH: 1, DOT: 2, DASHDOT: 3, DASHDOTDOT: 4, NULL: 5, INSIDEFRAME: 6, USERSTYLE: 7 };
  const STOCK = {
    0: { t: "brocha", estilo: 0, color: 0xFFFFFF }, 1: { t: "brocha", estilo: 0, color: 0xC0C0C0 }, 2: { t: "brocha", estilo: 0, color: 0x808080 },
    3: { t: "brocha", estilo: 0, color: 0x404040 }, 4: { t: "brocha", estilo: 0, color: 0 }, 5: { t: "brocha", estilo: 1, color: 0 },
    6: { t: "pluma", estilo: 0, ancho: 0, color: 0xFFFFFF }, 7: { t: "pluma", estilo: 0, ancho: 0, color: 0 }, 8: { t: "pluma", estilo: PS.NULL, ancho: 0, color: 0 },
    10: { t: "fuente", alto: -12, cara: "Courier New" }, 11: { t: "fuente", alto: -12, cara: "Courier New" }, 12: { t: "fuente", alto: -12, cara: "Arial" },
    13: { t: "fuente", alto: -12, cara: "Arial", peso: 700 }, 14: { t: "fuente", alto: -12, cara: "Arial" }, 16: { t: "fuente", alto: -12, cara: "Courier New" },
    17: { t: "fuente", alto: -11, cara: "Arial" }, 18: { t: "brocha", estilo: 0, color: 0xFFFFFF }, 19: { t: "pluma", estilo: 0, ancho: 0, color: 0 },
  };

  class Lienzo {
    constructor(opc) {
      this.out = []; this.defs = []; this.nId = 0; this.dibujado = 0;
      this.mm = opc.mm || 1;               // píxeles del dispositivo por milímetro (modos de mapeo fijos)
      this.wmf = !!opc.wmf; this.caja = opc.caja || null;
      this.objetos = [];
      this.pila = [];
      this.s = {
        pluma: STOCK[7], brocha: STOCK[0], fuente: STOCK[13], texto: 0, fondo: 0xFFFFFF, modoFondo: 2, relleno: "evenodd", alinTexto: 0,
        pos: [0, 0], mundo: [1, 0, 0, 1, 0, 0], modo: opc.wmf ? 8 : 1, vOrg: [0, 0], vExt: null, ventOrg: [0, 0], ventExt: null, recorte: null, dirArco: 1, extraLetra: 0,
      };
      this.tray = null;   // la trayectoria que se está armando (BEGINPATH … ENDPATH)
      this.trayLista = null;
      this.unidadPx = 1;  // cuánto mide un píxel de la salida en unidades del dispositivo (lo fija quien arma el SVG)
    }
    guardar() { this.pila.push(JSON.parse(JSON.stringify(Object.assign({}, this.s, { pluma: null, brocha: null, fuente: null }))) ); const t = this.pila[this.pila.length - 1]; t.pluma = this.s.pluma; t.brocha = this.s.brocha; t.fuente = this.s.fuente; }
    restaurar(n) {
      // n < 0: relativo (−1 = el último guardado); n > 0: absoluto
      let k = n < 0 ? this.pila.length + n : n - 1;
      if (k < 0 || k >= this.pila.length) return;
      this.s = this.pila[k];
      this.pila.length = k;
    }
    /** La matriz de lógico → dispositivo: [a, b, c, d, e, f]. */
    matriz() {
      const s = this.s, w = s.mundo;
      let sx = 1, sy = 1, tx = 0, ty = 0;
      if (this.wmf) {
        // WMF: la ventana se lleva a la caja de la imagen (la del encabezado «placeable»)
        if (s.ventExt && s.ventExt[0] && s.ventExt[1] && this.caja) {
          const c = this.caja, W = Math.abs(c[2] - c[0]), H = Math.abs(c[3] - c[1]), x0 = Math.min(c[0], c[2]), y0 = Math.min(c[1], c[3]);
          sx = W / s.ventExt[0]; sy = H / s.ventExt[1];
          // (una ventana con medida negativa voltea el dibujo: su origen queda del otro lado de la caja)
          tx = (sx >= 0 ? x0 : x0 + W) - s.ventOrg[0] * sx; ty = (sy >= 0 ? y0 : y0 + H) - s.ventOrg[1] * sy;
        }
      } else {
        const m = s.modo;
        if (m === 7 || m === 8) {
          if (s.ventExt && s.vExt && s.ventExt[0] && s.ventExt[1]) {
            sx = s.vExt[0] / s.ventExt[0]; sy = s.vExt[1] / s.ventExt[1];
            if (m === 7) { const k = Math.min(Math.abs(sx), Math.abs(sy)); sx = Math.sign(sx) * k; sy = Math.sign(sy) * k; }
          }
        } else if (m >= 2 && m <= 6) {
          const mmPorU = { 2: 0.1, 3: 0.01, 4: 0.254, 5: 0.0254, 6: 25.4 / 1440 }[m];
          sx = mmPorU * this.mm; sy = -sx;
        }
        tx = s.vOrg[0] - s.ventOrg[0] * sx; ty = s.vOrg[1] - s.ventOrg[1] * sy;
      }
      return [w[0] * sx, w[1] * sy, w[2] * sx, w[3] * sy, w[4] * sx + tx, w[5] * sy + ty];
    }
    p(x, y) { const m = this.matriz(); return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]; }
    /** Cuánto mide en el dispositivo una longitud lógica (para el grosor de la pluma y el tamaño de la letra). */
    escala() { const m = this.matriz(); return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1; }
    escalaY() { const m = this.matriz(); return Math.hypot(m[2], m[3]) || 1; }
    id(p) { return (p || "e") + ++this.nId; }

    // ---------------------------------------------------------------- objetos
    crear(ix, o) { if (ix == null) { let i = 0; while (this.objetos[i]) i++; ix = i; } this.objetos[ix] = o; return ix; }
    elegir(ix) {
      const o = ix >= 0x80000000 ? STOCK[ix & 0x7FFFFFFF] : this.objetos[ix];
      if (!o) return;
      if (o.t === "pluma") this.s.pluma = o; else if (o.t === "brocha") this.s.brocha = o; else if (o.t === "fuente") this.s.fuente = o;
      else if (o.t === "region") this.recortarRegion(o.rects, 5, true);
    }
    borrar(ix) { if (ix < 0x80000000) this.objetos[ix] = null; }

    // ---------------------------------------------------------------- estilos
    relleno(brocha) {
      const b = brocha || this.s.brocha;
      if (!b || b.estilo === 1) return 'fill="none"';
      if (b.estilo === 2) return `fill="url(#${this.trama(b)})"`;
      if ((b.estilo === 3 || b.estilo === 5 || b.estilo === 6) && b.img) return `fill="url(#${this.patronImg(b)})"`;
      return `fill="${color(b.color)}"`;
    }
    trama(b) {
      if (b._id) return b._id;
      const t = 8 * this.unidadPx, id = (b._id = this.id("t")), c = color(b.color), w = this.unidadPx;
      const lineas = { 0: [[0, t / 2, t, t / 2]], 1: [[t / 2, 0, t / 2, t]], 2: [[0, 0, t, t]], 3: [[0, t, t, 0]], 4: [[0, t / 2, t, t / 2], [t / 2, 0, t / 2, t]], 5: [[0, 0, t, t], [0, t, t, 0]] }[b.trama] || [[0, t / 2, t, t / 2]];
      const fondo = this.s.modoFondo === 2 ? `<rect width="${n3(t)}" height="${n3(t)}" fill="${color(this.s.fondo)}"/>` : "";
      this.defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${n3(t)}" height="${n3(t)}">${fondo}${lineas.map((l) => `<line x1="${n3(l[0])}" y1="${n3(l[1])}" x2="${n3(l[2])}" y2="${n3(l[3])}" stroke="${c}" stroke-width="${n3(w)}"/>`).join("")}</pattern>`);
      return id;
    }
    patronImg(b) {
      if (b._id) return b._id;
      const id = (b._id = this.id("p")), w = b.img.w * this.unidadPx, h = b.img.h * this.unidadPx;
      this.defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${n3(w)}" height="${n3(h)}"><image width="${n3(w)}" height="${n3(h)}" preserveAspectRatio="none" href="${b.img.url}"/></pattern>`);
      return id;
    }
    contorno() {
      const p = this.s.pluma;
      if (!p || (p.estilo & 0xF) === PS.NULL) return 'stroke="none"';
      const geom = !!(p.estilo & 0x10000);
      let w = p.ancho ? p.ancho * this.escala() : 0;
      if (!w || (!geom && !this.wmf && p.ancho <= 1)) w = Math.max(w, this.unidadPx);
      w = Math.max(w, this.unidadPx * 0.75);
      const tipo = p.estilo & 0xF, cap = p.estilo & 0xF00, union = p.estilo & 0xF000;
      const g = { 1: [18, 6], 2: [3, 3], 3: [9, 6, 3, 6], 4: [9, 3, 3, 3, 3, 3] }[tipo];
      const unidad = geom ? w : this.unidadPx;
      let a = `stroke="${color(p.color)}" stroke-width="${n3(w)}"`;
      if (g) a += ` stroke-dasharray="${g.map((v) => n3((v * unidad) / (geom ? 3 : 1))).join(" ")}"`;
      else if (tipo === PS.USERSTYLE && p.guiones && p.guiones.length) a += ` stroke-dasharray="${p.guiones.map((v) => n3(v * this.escala())).join(" ")}"`;
      a += cap === 0x100 ? ' stroke-linecap="square"' : cap === 0x200 ? ' stroke-linecap="butt"' : ' stroke-linecap="round"';
      a += union === 0x1000 ? ' stroke-linejoin="bevel"' : union === 0x2000 ? ' stroke-linejoin="miter"' : ' stroke-linejoin="round"';
      return a;
    }
    recorte() { return this.s.recorte ? ` clip-path="url(#${this.s.recorte})"` : ""; }

    // ---------------------------------------------------------------- recortes
    recortarRect(x1, y1, x2, y2, excluir) {
      const a = this.p(x1, y1), b = this.p(x2, y2);
      const x = Math.min(a[0], b[0]), y = Math.min(a[1], b[1]), w = Math.abs(b[0] - a[0]), h = Math.abs(b[1] - a[1]);
      const id = this.id("c");
      if (excluir) {
        // (lo de fuera del rectángulo: un rectángulo enorme con el hueco, relleno par-impar)
        const G = 1e7;
        this.defs.push(`<clipPath id="${id}"${this.recorte()}><path clip-rule="evenodd" d="M${-G},${-G}H${G}V${G}H${-G}Z M${n3(x)},${n3(y)}h${n3(w)}v${n3(h)}h${n3(-w)}Z"/></clipPath>`);
      } else this.defs.push(`<clipPath id="${id}"${this.recorte()}><rect x="${n3(x)}" y="${n3(y)}" width="${n3(w)}" height="${n3(h)}"/></clipPath>`);
      this.s.recorte = id;
    }
    /** rects en unidades del dispositivo; modo: 1 Y, 2 O, 3 O exclusivo, 4 diferencia, 5 copiar. */
    recortarRegion(rects, modo, logico) {
      if (!rects || !rects.length) { if (modo === 5) this.s.recorte = null; return; }
      const id = this.id("c");
      const d = rects.map((r) => {
        let a = [r[0], r[1]], b = [r[2], r[3]];
        if (logico) { a = this.p(r[0], r[1]); b = this.p(r[2], r[3]); }
        return `M${n3(Math.min(a[0], b[0]))},${n3(Math.min(a[1], b[1]))}H${n3(Math.max(a[0], b[0]))}V${n3(Math.max(a[1], b[1]))}H${n3(Math.min(a[0], b[0]))}Z`;
      }).join("");
      const conActual = modo === 1 && this.s.recorte ? this.recorte() : "";
      this.defs.push(`<clipPath id="${id}"${conActual}><path d="${d}"/></clipPath>`);
      this.s.recorte = id;
    }

    // ---------------------------------------------------------------- trayectorias
    /** Agrega a la trayectoria en curso, o dibuja de una vez. `d`: lista de comandos en coordenadas del dispositivo. */
    figura(d, rellenar, perfilar) {
      if (!d) return;
      if (this.tray != null) { this.tray += d; return; }
      this.dibujar(d, rellenar, perfilar);
    }
    dibujar(d, rellenar, perfilar, brocha) {
      const f = rellenar ? this.relleno(brocha) : 'fill="none"';
      const c = perfilar ? this.contorno() : 'stroke="none"';
      if (f === 'fill="none"' && c === 'stroke="none"') return;
      this.out.push(`<path d="${d}" ${f}${rellenar ? ` fill-rule="${this.s.relleno}"` : ""} ${c}${this.recorte()}/>`);
      this.dibujado++;
    }
    lineaA(x, y) {
      const a = this.p(...this.s.pos), b = this.p(x, y);
      if (this.tray != null) { if (!this.tray) this.tray = `M${n3(a[0])},${n3(a[1])}`; this.tray += `L${n3(b[0])},${n3(b[1])}`; }
      else this.dibujar(`M${n3(a[0])},${n3(a[1])}L${n3(b[0])},${n3(b[1])}`, false, true);
      this.s.pos = [x, y];
    }
    moverA(x, y) { this.s.pos = [x, y]; if (this.tray != null) { const a = this.p(x, y); this.tray += `M${n3(a[0])},${n3(a[1])}`; } }
    poli(pts, cerrar, rellenar) {
      if (!pts.length) return "";
      let d = "";
      pts.forEach((q, i) => { const a = this.p(q[0], q[1]); d += (i ? "L" : "M") + n3(a[0]) + "," + n3(a[1]); });
      if (cerrar) d += "Z";
      void rellenar;
      return d;
    }
    bezier(pts, desdeActual) {
      let d = "", i = 0;
      if (!desdeActual) { if (!pts.length) return ""; const a = this.p(pts[0][0], pts[0][1]); d += `M${n3(a[0])},${n3(a[1])}`; i = 1; }
      else { const a = this.p(...this.s.pos); d += this.tray != null && this.tray ? "" : `M${n3(a[0])},${n3(a[1])}`; }
      for (; i + 2 < pts.length; i += 3) {
        const [c1, c2, e] = [this.p(...pts[i]), this.p(...pts[i + 1]), this.p(...pts[i + 2])];
        d += `C${n3(c1[0])},${n3(c1[1])} ${n3(c2[0])},${n3(c2[1])} ${n3(e[0])},${n3(e[1])}`;
      }
      if (pts.length) this.s.pos = pts[pts.length - 1].slice();
      return d;
    }
    /** Arco de elipse dentro de la caja, del punto inicial al final (radiales), como curvas. Lógico, y hacia abajo. */
    arco(l, t, r, b, xs, ys, xe, ye, tipo) {
      const cx = (l + r) / 2, cy = (t + b) / 2, rx = Math.abs(r - l) / 2, ry = Math.abs(b - t) / 2;
      if (!rx || !ry) return "";
      let a0 = Math.atan2(-(ys - cy) / ry, (xs - cx) / rx), a1 = Math.atan2(-(ye - cy) / ry, (xe - cx) / rx);
      const antihorario = this.s.dirArco !== 2;
      let barrido = antihorario ? a1 - a0 : a0 - a1;
      while (barrido <= 1e-9) barrido += 2 * Math.PI;
      if (xs === xe && ys === ye) barrido = 2 * Math.PI;
      if (!antihorario) barrido = -barrido;
      const punto = (a) => [cx + rx * Math.cos(a), cy - ry * Math.sin(a)];
      const n = Math.max(1, Math.ceil(Math.abs(barrido) / (Math.PI / 2) - 1e-9)), paso = barrido / n, k = (4 / 3) * Math.tan(paso / 4);
      const ini = punto(a0);
      let d = "";
      if (tipo === "pie") { const c = this.p(cx, cy), q = this.p(...ini); d = `M${n3(c[0])},${n3(c[1])}L${n3(q[0])},${n3(q[1])}`; }
      else if (tipo === "a") { const q = this.p(...ini); d = this.tray != null && this.tray ? `L${n3(q[0])},${n3(q[1])}` : `M${n3(q[0])},${n3(q[1])}`; }
      else { const q = this.p(...ini); d = `M${n3(q[0])},${n3(q[1])}`; }
      for (let i = 0; i < n; i++) {
        const t0 = a0 + i * paso, t1 = t0 + paso;
        const p0 = punto(t0), p1 = punto(t1);
        const c1 = [p0[0] - k * rx * Math.sin(t0), p0[1] - k * ry * Math.cos(t0)];
        const c2 = [p1[0] + k * rx * Math.sin(t1), p1[1] + k * ry * Math.cos(t1)];
        const [q1, q2, q3] = [this.p(...c1), this.p(...c2), this.p(...p1)];
        d += `C${n3(q1[0])},${n3(q1[1])} ${n3(q2[0])},${n3(q2[1])} ${n3(q3[0])},${n3(q3[1])}`;
      }
      if (tipo === "pie" || tipo === "chord") d += "Z";
      this.s.posArco = punto(a0 + barrido);
      return d;
    }
    elipse(l, t, r, b) { return this.arco(l, t, r, b, r, (t + b) / 2, r, (t + b) / 2, "e") + "Z"; }
    rect(l, t, r, b) { return this.poli([[l, t], [r, t], [r, b], [l, b]], true); }
    rectRedondo(l, t, r, b, w, h) {
      w = Math.min(Math.abs(w), Math.abs(r - l)); h = Math.min(Math.abs(h), Math.abs(b - t));
      if (!w || !h) return this.rect(l, t, r, b);
      const L = Math.min(l, r), R = Math.max(l, r), T = Math.min(t, b), Bo = Math.max(t, b), rx = w / 2, ry = h / 2, k = 0.5523;
      const pts = [["M", R - rx, T], ["C", R - rx + k * rx, T, R, T + ry - k * ry, R, T + ry], ["L", R, Bo - ry], ["C", R, Bo - ry + k * ry, R - rx + k * rx, Bo, R - rx, Bo],
        ["L", L + rx, Bo], ["C", L + rx - k * rx, Bo, L, Bo - ry + k * ry, L, Bo - ry], ["L", L, T + ry], ["C", L, T + ry - k * ry, L + rx - k * rx, T, L + rx, T]];
      return pts.map((q) => { let s = q[0]; for (let i = 1; i < q.length; i += 2) { const a = this.p(q[i], q[i + 1]); s += (i > 1 ? " " : "") + n3(a[0]) + "," + n3(a[1]); } return s; }).join("") + "Z";
    }

    // ---------------------------------------------------------------- texto
    /** x, y: punto de referencia (lógico); dx: avances (lógicos) o null. */
    texto(x, y, cad, dx, opciones, rectOpaco) {
      if (!cad) return;
      const f = this.s.fuente || STOCK[13];
      const s = this.s;
      if (/^marlett$/i.test(f.cara || "")) cad = deMarlett(cad);
      else if (/mt extra/i.test(f.cara || "")) cad = deMtExtra(cad);
      else if (f.symbol) cad = deSymbol(cad);
      if (rectOpaco && (opciones & 2)) this.dibujar(this.rect(rectOpaco[0], rectOpaco[1], rectOpaco[2], rectOpaco[3]), true, false, { estilo: 0, color: s.fondo });
      const alin = s.alinTexto;
      if (alin & 1) { x = s.pos[0]; y = s.pos[1]; }   // TA_UPDATECP
      const ey = this.escalaY();
      const altoL = Math.abs(f.alto || 12);
      const em = (f.alto < 0 ? altoL : altoL * (f.symbol ? 0.82 : 0.89)) * ey;
      if (em <= 0) return;
      const asc = f.alto < 0 ? em * 0.9 : altoL * ey * (f.symbol ? 0.82 : 0.8);
      const desc = f.alto < 0 ? em * 0.22 : altoL * ey - asc;
      const m = this.matriz();
      const escX = Math.hypot(m[0], m[1]) || 1;
      // el avance de cada letra, en el dispositivo
      const letras = Array.from(cad);
      let avances = null;
      if (dx && dx.length) { avances = []; for (let i = 0; i < letras.length; i++) avances.push(((dx[i] || 0) + 0) * escX); }
      const total = avances ? avances.reduce((a, v) => a + v, 0) : null;
      const o = this.p(x, y);
      // giro: el de la fuente (décimas de grado, contra las manecillas) y el del mundo
      const giroMundo = (Math.atan2(m[1], m[0]) * 180) / Math.PI;
      const giro = -(f.escape || 0) / 10 + (m[0] < 0 ? 0 : giroMundo) * (Math.sign(m[3]) || 1);
      let x0 = o[0], y0 = o[1];
      // vertical: arriba (0), abajo (8) o en la línea base (24)
      const v = alin & 24;
      if (v === 0) y0 += asc; else if (v === 8) y0 -= desc;
      // horizontal: izquierda (0), derecha (2) o centro (6)
      const h = alin & 6;
      let anclaje = "";
      if (h === 2) { if (total != null) x0 -= total; else anclaje = ' text-anchor="end"'; }
      else if (h === 6) { if (total != null) x0 -= total / 2; else anclaje = ' text-anchor="middle"'; }
      let xs = "";
      if (avances) { let acc = x0; const xsL = []; for (let i = 0; i < letras.length; i++) { xsL.push(n3(acc)); acc += avances[i]; } xs = xsL.join(" "); }
      else xs = n3(x0);
      if ((alin & 1) && total != null) { const inv = this.inversa(); if (inv) { const fin = [inv[0] * (o[0] + total) + inv[2] * o[1] + inv[4], inv[1] * (o[0] + total) + inv[3] * o[1] + inv[5]]; s.pos = fin; } }
      const estilo = [`font-family:${familia(f.cara)}`, `font-size:${n3(em)}px`, f.peso >= 600 ? "font-weight:bold" : "", f.cursiva ? "font-style:italic" : "",
        f.subrayado || f.tachado ? `text-decoration:${[f.subrayado ? "underline" : "", f.tachado ? "line-through" : ""].filter(Boolean).join(" ")}` : ""].filter(Boolean).join(";");
      // lfWidth: el ancho medio pedido contra el natural de la fuente; las letras se estiran o encogen (su lugar no cambia)
      let kx = 1;
      if (+f.ancho > 0) { const nat = (ANCHO_MEDIO[String(f.cara || "").toLowerCase()] || 0.45) * em; const pedido = Math.abs(f.ancho) * escX; if (nat > 0 && pedido > 0) kx = Math.max(0.2, Math.min(5, pedido / nat)); }
      if (kx !== 1) {
        const xsN = xs.split(" ").map((v) => n3(x0 + (+v - x0) / kx)).join(" ");
        xs = xsN;
      }
      const tr0 = giro && Math.abs(giro) > 0.01 ? `rotate(${n3(giro)} ${n3(o[0])} ${n3(o[1])})` : "";
      const tr1 = kx !== 1 ? `translate(${n3(x0)} 0) scale(${n3(kx)} 1) translate(${n3(-x0)} 0)` : "";
      const tr = tr0 || tr1 ? ` transform="${[tr0, tr1].filter(Boolean).join(" ")}"` : "";
      this.out.push(`<text x="${xs}" y="${n3(y0)}"${anclaje} fill="${color(s.texto)}" style="${esc(estilo)};white-space:pre" xml:space="preserve"${tr}${this.recorte()}>${esc(letras.join(""))}</text>`);
      this.dibujado++;
    }
    inversa() {
      const m = this.matriz(), det = m[0] * m[3] - m[1] * m[2];
      if (!det) return null;
      return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
    }

    // ---------------------------------------------------------------- imágenes
    /** Un mapa de bits en el rectángulo lógico (x, y, w, h), con la parte (sx, sy, sw, sh) de la imagen. */
    imagen(img, x, y, w, h, sx, sy, sw, sh, rop) {
      if (!img || !w || !h) return;
      const m = this.matriz();
      if (sw == null || !sw) { sx = 0; sy = 0; sw = img.w; sh = img.h; }
      // (en un DIB de abajo hacia arriba, el origen de la parte está abajo)
      if (img.abajoArriba && (sy || sh !== img.h)) sy = img.h - sy - Math.abs(sh);
      const mezcla = rop === 0x8800C6 ? ";mix-blend-mode:multiply" : rop === 0xEE0086 ? ";mix-blend-mode:screen" : rop === 0x660046 ? ";mix-blend-mode:difference" : "";
      if (mezcla) this.conFondo = true;
      const sxW = Math.sign(w) || 1, syH = Math.sign(h) || 1;
      this.out.push(`<g transform="matrix(${m.map(n3).join(" ")}) translate(${n3(x)} ${n3(y)}) scale(${sxW} ${syH})"${this.recorte()}><svg x="0" y="0" width="${n3(Math.abs(w))}" height="${n3(Math.abs(h))}" viewBox="${n3(sx)} ${n3(sy)} ${n3(Math.abs(sw))} ${n3(Math.abs(sh))}" preserveAspectRatio="none" overflow="hidden"><image width="${img.w}" height="${img.h}" preserveAspectRatio="none" href="${img.url}" style="image-rendering:auto${mezcla}"/></svg></g>`);
      this.dibujado++;
    }
    /** Rellenar un rectángulo con la brocha (PATBLT, BITBLT sin imagen) según la operación. */
    patron(x, y, w, h, rop) {
      let br = this.s.brocha;
      if (rop === 0x42) br = { estilo: 0, color: 0 };
      else if (rop === 0xFF0062) br = { estilo: 0, color: 0xFFFFFF };
      else if (rop === 0x550009 || rop === 0xAA0029) return;   // invertir el destino / no hacer nada
      this.dibujar(this.rect(x, y, x + w, y + h), true, false, br);
    }
    degradado(verts, rect, modo) {
      // rectángulos con degradado horizontal (0) o vertical (1); triángulos (2): el color medio
      if (modo === 2) {
        const [a, b, c] = rect.map((i) => verts[i]);
        if (!a || !b || !c) return;
        const col = (k) => Math.round(((a[k] + b[k] + c[k]) / 3) / 256);
        this.dibujar(this.poli([[a[0], a[1]], [b[0], b[1]], [c[0], c[1]]], true), true, false, { estilo: 0, color: col(2) | (col(3) << 8) | (col(4) << 16) });
        return;
      }
      const v1 = verts[rect[0]], v2 = verts[rect[1]];
      if (!v1 || !v2) return;
      const id = this.id("g"), p1 = this.p(v1[0], v1[1]), p2 = this.p(v2[0], v2[1]);
      const c = (v) => color(Math.round(v[2] / 256) | (Math.round(v[3] / 256) << 8) | (Math.round(v[4] / 256) << 16));
      const [x1, y1, x2, y2] = modo === 0 ? [p1[0], p1[1], p2[0], p1[1]] : [p1[0], p1[1], p1[0], p2[1]];
      this.defs.push(`<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${n3(x1)}" y1="${n3(y1)}" x2="${n3(x2)}" y2="${n3(y2)}"><stop offset="0" stop-color="${c(v1)}"/><stop offset="1" stop-color="${c(v2)}"/></linearGradient>`);
      this.out.push(`<path d="${this.rect(v1[0], v1[1], v2[0], v2[1])}" fill="url(#${id})"${this.recorte()}/>`);
      this.dibujado++;
    }

    svg(vx, vy, vw, vh, wPx, hPx) {
      // (las máscaras de GDI —Y, O de mapas de bits— se combinan con lo que hay debajo: la hoja, blanca)
      const fondo = this.conFondo ? `<rect x="${n3(vx)}" y="${n3(vy)}" width="${n3(vw)}" height="${n3(vh)}" fill="#fff"/>` : "";
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${n3(wPx)}" height="${n3(hPx)}" viewBox="${n3(vx)} ${n3(vy)} ${n3(vw)} ${n3(vh)}" preserveAspectRatio="none" overflow="hidden">` +
        (this.defs.length ? `<defs>${this.defs.join("")}</defs>` : "") + fondo + this.out.join("") + "</svg>";
    }
  }

  // ------------------------------------------------------------ EMF
  function emf(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (u8.length < 88 || dv.getUint32(0, true) !== 1 || dv.getUint32(40, true) !== 0x464D4520) return null;
    const i32 = (o) => dv.getInt32(o, true), u32 = (o) => dv.getUint32(o, true), i16 = (o) => dv.getInt16(o, true), f32 = (o) => dv.getFloat32(o, true);
    const tamCab = u32(4);
    const limites = [i32(8), i32(12), i32(16), i32(20)], marco = [i32(24), i32(28), i32(32), i32(36)];
    const disp = [i32(72), i32(76)], mm = [i32(80), i32(84)];
    let micras = null;
    if (tamCab >= 108 && u8.length >= 108) micras = [i32(100), i32(104)];
    const pxPorMm = micras && micras[0] > 0 ? (disp[0] * 1000) / micras[0] : mm[0] > 0 ? disp[0] / mm[0] : 96 / 25.4;
    const pxPorMmY = micras && micras[1] > 0 ? (disp[1] * 1000) / micras[1] : mm[1] > 0 ? disp[1] / mm[1] : pxPorMm;
    // el marco (centésimas de milímetro) en unidades del dispositivo: lo que ocupa la imagen
    let vx = (marco[0] / 100) * pxPorMm, vy = (marco[1] / 100) * pxPorMmY, vw = ((marco[2] - marco[0]) / 100) * pxPorMm, vh = ((marco[3] - marco[1]) / 100) * pxPorMmY;
    if (!(vw > 0 && vh > 0)) { vx = limites[0]; vy = limites[1]; vw = limites[2] - limites[0] + 1; vh = limites[3] - limites[1] + 1; }
    if (!(vw > 0 && vh > 0)) return null;
    const wPx = marco[2] > marco[0] ? ((marco[2] - marco[0]) / 100 / 25.4) * 96 : vw, hPx = marco[3] > marco[1] ? ((marco[3] - marco[1]) / 100 / 25.4) * 96 : vh;
    const L = new Lienzo({ mm: pxPorMm });
    L.unidadPx = vw / Math.max(1, wPx);
    const puntos = (o, n, corto) => { const r = []; for (let i = 0; i < n; i++) r.push(corto ? [i16(o + i * 4), i16(o + i * 4 + 2)] : [i32(o + i * 8), i32(o + i * 8 + 4)]); return r; };
    const region = (o, tam) => {
      if (tam < 32) return [];
      const n = u32(o + 8), r = [];
      for (let i = 0; i < n && o + 32 + i * 16 + 16 <= u8.length; i++) { const q = o + 32 + i * 16; r.push([i32(q), i32(q + 4), i32(q + 8), i32(q + 12)]); }
      return r;
    };
    let hayMas = false, gdi = 0;
    let p = 0, vueltas = 0;
    while (p + 8 <= u8.length && vueltas++ < 2000000) {
      const tipo = u32(p), tam = u32(p + 4);
      if (tam < 8 || p + tam > u8.length) break;
      const d = p + 8;
      try {
        switch (tipo) {
          case 14: p = u8.length; continue;   // EOF
          case 70: if (tam >= 16 && u32(d + 4) === 0x2B464D45) hayMas = true; break;   // comentario (EMF+ incluido)
          case 9: L.s.ventExt = [i32(d), i32(d + 4)]; break;
          case 10: L.s.ventOrg = [i32(d), i32(d + 4)]; break;
          case 11: L.s.vExt = [i32(d), i32(d + 4)]; break;
          case 12: L.s.vOrg = [i32(d), i32(d + 4)]; break;
          case 31: if (L.s.vExt) L.s.vExt = [(L.s.vExt[0] * i32(d)) / (i32(d + 4) || 1), (L.s.vExt[1] * i32(d + 8)) / (i32(d + 12) || 1)]; break;
          case 32: if (L.s.ventExt) L.s.ventExt = [(L.s.ventExt[0] * i32(d)) / (i32(d + 4) || 1), (L.s.ventExt[1] * i32(d + 8)) / (i32(d + 12) || 1)]; break;
          case 17: L.s.modo = u32(d); break;
          case 18: L.s.modoFondo = u32(d); break;
          case 19: L.s.relleno = u32(d) === 2 ? "nonzero" : "evenodd"; break;
          case 22: L.s.alinTexto = u32(d); break;
          case 24: L.s.texto = u32(d); break;
          case 25: L.s.fondo = u32(d); break;
          case 57: L.s.dirArco = u32(d); break;
          case 27: L.moverA(i32(d), i32(d + 4)); break;
          case 54: L.lineaA(i32(d), i32(d + 4)); gdi++; break;
          case 33: L.guardar(); break;
          case 34: L.restaurar(i32(d)); break;
          case 35: L.s.mundo = [f32(d), f32(d + 4), f32(d + 8), f32(d + 12), f32(d + 16), f32(d + 20)]; break;
          case 36: {
            const x = [f32(d), f32(d + 4), f32(d + 8), f32(d + 12), f32(d + 16), f32(d + 20)], modo = u32(d + 24), w = L.s.mundo;
            const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
            if (modo === 1) L.s.mundo = [1, 0, 0, 1, 0, 0];
            else if (modo === 2) L.s.mundo = mul(x, w);
            else if (modo === 3) L.s.mundo = mul(w, x);
            else if (modo === 4) L.s.mundo = x;
            break;
          }
          case 37: L.elegir(u32(d)); break;
          case 40: L.borrar(u32(d)); break;
          case 38: L.crear(u32(d), { t: "pluma", estilo: u32(d + 4), ancho: i32(d + 8), color: u32(d + 16) }); break;
          case 95: {
            const ix = u32(d), o = d + 20, n = u32(o + 20);
            const g = []; for (let i = 0; i < n && o + 24 + i * 4 + 4 <= p + tam; i++) g.push(u32(o + 24 + i * 4));
            L.crear(ix, { t: "pluma", estilo: u32(o), ancho: u32(o + 4), color: u32(o + 12), guiones: g, geom: !!(u32(o) & 0x10000) });
            break;
          }
          case 39: L.crear(u32(d), { t: "brocha", estilo: u32(d + 4), color: u32(d + 8), trama: u32(d + 12) }); break;
          case 93: case 94: {
            const ix = u32(d), offBmi = u32(d + 8), cbBmi = u32(d + 12), offBits = u32(d + 16), cbBits = u32(d + 20);
            const img = dibAUrl(u8, p + offBmi, cbBmi, p + offBits, cbBits);
            L.crear(ix, img ? { t: "brocha", estilo: 5, img, color: 0 } : { t: "brocha", estilo: 0, color: 0x808080 });
            break;
          }
          case 82: {
            const ix = u32(d), o = d + 4;
            let cara = ""; for (let i = 0; i < 32; i++) { const c = dv.getUint16(o + 28 + i * 2, true); if (!c) break; cara += String.fromCharCode(c); }
            L.crear(ix, { t: "fuente", alto: i32(o), ancho: i32(o + 4), escape: i32(o + 8), peso: i32(o + 16), cursiva: u8[o + 20], subrayado: u8[o + 21], tachado: u8[o + 22], charset: u8[o + 23], cara, symbol: u8[o + 23] === 2 || /^symbol$/i.test(cara) });
            break;
          }
          case 48: case 49: case 50: case 51: case 52: break;   // paletas
          case 30: L.recortarRect(i32(d), i32(d + 4), i32(d + 8), i32(d + 12)); break;
          case 29: L.recortarRect(i32(d), i32(d + 4), i32(d + 8), i32(d + 12), true); break;
          case 75: { const cb = u32(d), modo = u32(d + 4); if (!cb) { if (modo === 5) L.s.recorte = null; } else L.recortarRegion(region(d + 8, cb), modo, false); break; }
          case 67: { const modo = u32(d); if (L.trayLista) { const id = L.id("c"); L.defs.push(`<clipPath id="${id}"${modo === 1 && L.s.recorte ? L.recorte() : ""}><path d="${L.trayLista}" clip-rule="${L.s.relleno}"/></clipPath>`); L.s.recorte = id; } L.trayLista = null; break; }
          case 59: L.tray = ""; break;
          case 60: L.trayLista = L.tray; L.tray = null; break;
          case 68: L.tray = null; L.trayLista = null; break;
          case 61: if (L.tray != null) L.tray += "Z"; break;
          case 62: if (L.trayLista) { L.dibujar(L.trayLista, true, false); gdi++; } L.trayLista = null; break;
          case 63: if (L.trayLista) { L.dibujar(L.trayLista, true, true); gdi++; } L.trayLista = null; break;
          case 64: if (L.trayLista) { L.dibujar(L.trayLista, false, true); gdi++; } L.trayLista = null; break;
          case 65: case 66: break;
          // figuras (lo de 16 bits: puntos cortos)
          case 3: case 86: { const n = u32(d + 16); L.figura(L.poli(puntos(d + 20, n, tipo === 86), true), true, true); gdi++; break; }
          case 4: case 87: { const n = u32(d + 16); L.figura(L.poli(puntos(d + 20, n, tipo === 87), false), false, true); gdi++; break; }
          case 6: case 89: {
            const n = u32(d + 16), pts = puntos(d + 20, n, tipo === 89);
            if (L.tray != null) { pts.forEach((q) => L.lineaA(q[0], q[1])); }
            else { const ini = L.s.pos.slice(); L.figura(L.poli([ini].concat(pts), false), false, true); if (pts.length) L.s.pos = pts[pts.length - 1].slice(); }
            gdi++; break;
          }
          case 2: case 85: { const n = u32(d + 16); L.figura(L.bezier(puntos(d + 20, n, tipo === 85), false), false, true); gdi++; break; }
          case 5: case 88: {
            const n = u32(d + 16), pts = puntos(d + 20, n, tipo === 88);
            const dd = L.bezier(pts, true);
            if (L.tray != null) L.tray += dd; else L.dibujar(dd, false, true);
            gdi++; break;
          }
          case 7: case 90: case 8: case 91: {
            const np = u32(d + 16), corto = tipo === 90 || tipo === 91, cerrar = tipo === 8 || tipo === 91;
            const cuentas = []; for (let i = 0; i < np; i++) cuentas.push(u32(d + 24 + i * 4));
            let o = d + 24 + np * 4, dd = "";
            for (const c of cuentas) { dd += L.poli(puntos(o, c, corto), cerrar); o += c * (corto ? 4 : 8); }
            L.figura(dd, cerrar, true); gdi++; break;
          }
          case 56: case 92: {
            const n = u32(d + 16), corto = tipo === 92, pts = puntos(d + 20, n, corto), tps = d + 20 + n * (corto ? 4 : 8);
            let dd = "";
            for (let i = 0; i < n; i++) {
              const t = u8[tps + i], q = L.p(pts[i][0], pts[i][1]);
              if ((t & 6) === 6) dd += `M${n3(q[0])},${n3(q[1])}`;
              else if ((t & 6) === 4 && i + 2 < n) { const q2 = L.p(...pts[i + 1]), q3 = L.p(...pts[i + 2]); dd += `C${n3(q[0])},${n3(q[1])} ${n3(q2[0])},${n3(q2[1])} ${n3(q3[0])},${n3(q3[1])}`; i += 2; if (u8[tps + i] & 1) dd += "Z"; continue; }
              else dd += `L${n3(q[0])},${n3(q[1])}`;
              if (t & 1) dd += "Z";
            }
            if (n) L.s.pos = pts[n - 1].slice();
            L.figura(dd, false, true); gdi++; break;
          }
          case 43: L.figura(L.rect(i32(d), i32(d + 4), i32(d + 8), i32(d + 12)), true, true); gdi++; break;
          case 42: L.figura(L.elipse(i32(d), i32(d + 4), i32(d + 8), i32(d + 12)), true, true); gdi++; break;
          case 44: L.figura(L.rectRedondo(i32(d), i32(d + 4), i32(d + 8), i32(d + 12), i32(d + 16), i32(d + 20)), true, true); gdi++; break;
          case 45: case 46: case 47: case 55: {
            const t = { 45: "arc", 46: "chord", 47: "pie", 55: "a" }[tipo];
            const cur = L.p(...L.s.pos);
            let dd = L.arco(i32(d), i32(d + 4), i32(d + 8), i32(d + 12), i32(d + 16), i32(d + 20), i32(d + 24), i32(d + 28), t);
            if (t === "a") {
              // (ARCTO: una línea desde donde se estaba hasta el inicio del arco, y el arco)
              if (L.tray != null) L.tray += dd; else L.dibujar(`M${n3(cur[0])},${n3(cur[1])}` + dd.replace(/^M/, "L"), false, true);
              if (L.s.posArco) L.s.pos = L.s.posArco;
            }
            else L.figura(dd, t !== "arc", true);
            gdi++; break;
          }
          case 41: {
            const cx = i32(d), cy = i32(d + 4), r = u32(d + 8), a0 = (f32(d + 12) * Math.PI) / 180, sw = (f32(d + 16) * Math.PI) / 180;
            const xs = cx + r * Math.cos(a0), ys = cy - r * Math.sin(a0), xe = cx + r * Math.cos(a0 + sw), ye = cy - r * Math.sin(a0 + sw);
            const dir = L.s.dirArco; L.s.dirArco = sw >= 0 ? 1 : 2;
            L.lineaA(xs, ys);
            const dd = L.arco(cx - r, cy - r, cx + r, cy + r, xs, ys, xe, ye, "a");
            if (L.tray != null) L.tray += dd; else L.dibujar(dd, false, true);
            L.s.dirArco = dir; L.s.pos = [xe, ye]; gdi++; break;
          }
          case 15: { const x = i32(d), y = i32(d + 4); L.dibujar(L.rect(x, y, x + 1, y + 1), true, false, { estilo: 0, color: u32(d + 8) }); break; }
          case 71: case 74: {
            const cb = u32(d + 16), conBrocha = tipo === 71, br = conBrocha ? (u32(d + 20) >= 0x80000000 ? STOCK[u32(d + 20) & 0x7FFFFFFF] : L.objetos[u32(d + 20)]) : null;
            const rects = region(d + (conBrocha ? 24 : 20), cb);
            const dd = rects.map((r) => L.rect(r[0], r[1], r[2], r[3])).join("");
            if (dd) { L.dibujar(dd, true, false, br || undefined); gdi++; }
            break;
          }
          case 118: {
            const nv = u32(d + 16), nt = u32(d + 20), modo = u32(d + 24);
            const verts = []; let o = d + 28;
            for (let i = 0; i < nv; i++, o += 16) verts.push([i32(o), i32(o + 4), dv.getUint16(o + 8, true), dv.getUint16(o + 10, true), dv.getUint16(o + 12, true)]);
            for (let i = 0; i < nt; i++) {
              if (modo === 2) { L.degradado(verts, [u32(o), u32(o + 4), u32(o + 8)], 2); o += 12; }
              else { L.degradado(verts, [u32(o), u32(o + 4)], modo); o += 8; }
            }
            gdi++; break;
          }
          case 81: {
            const xD = i32(d + 16), yD = i32(d + 20), xS = i32(d + 24), yS = i32(d + 28), cxS = i32(d + 32), cyS = i32(d + 36);
            const offBmi = u32(d + 40), cbBmi = u32(d + 44), offBits = u32(d + 48), cbBits = u32(d + 52), rop = u32(d + 60), cxD = i32(d + 64), cyD = i32(d + 68);
            const img = cbBmi ? dibAUrl(u8, p + offBmi, cbBmi, p + offBits, cbBits) : null;
            if (img) L.imagen(img, xD, yD, cxD, cyD, xS, yS, cxS, cyS, rop); else L.patron(xD, yD, cxD, cyD, rop);
            gdi++; break;
          }
          case 76: case 77: case 114: case 116: {
            const xD = i32(d + 16), yD = i32(d + 20), cxD = i32(d + 24), cyD = i32(d + 28);
            const rop = tipo === 114 || tipo === 116 ? 0xCC0020 : u32(d + 32);
            const base = d + 36;   // xSrc, ySrc, xform (24), crBk, iUsage, offBmi, cbBmi, offBits, cbBits [cxSrc, cySrc]
            const xS = i32(base), yS = i32(base + 4), offBmi = u32(base + 40), cbBmi = u32(base + 44), offBits = u32(base + 48), cbBits = u32(base + 52);
            const cxS = tipo === 76 ? cxD : i32(base + 56), cyS = tipo === 76 ? cyD : i32(base + 60);
            const img = cbBmi ? dibAUrl(u8, p + offBmi, cbBmi, p + offBits, cbBits) : null;
            if (img) L.imagen(img, xD, yD, cxD, cyD, xS, yS, cxS, cyS, rop); else L.patron(xD, yD, cxD, cyD, rop);
            gdi++; break;
          }
          case 80: {
            const xD = i32(d + 16), yD = i32(d + 20), xS = i32(d + 24), yS = i32(d + 28), cx = i32(d + 32), cy = i32(d + 36);
            const offBmi = u32(d + 40), cbBmi = u32(d + 44), offBits = u32(d + 48), cbBits = u32(d + 52);
            const img = cbBmi ? dibAUrl(u8, p + offBmi, cbBmi, p + offBits, cbBits) : null;
            if (img) L.imagen(img, xD, yD, cx, cy, xS, yS, cx, cy, 0xCC0020);
            gdi++; break;
          }
          case 83: case 84: {
            const o = d + 28;   // EMRTEXT
            const x = i32(o), y = i32(o + 4), n = u32(o + 8), offS = u32(o + 12), opc = u32(o + 16), rc = [i32(o + 20), i32(o + 24), i32(o + 28), i32(o + 32)], offDx = u32(o + 36);
            if (opc & 0x10) break;   // índices de glifos: no se pueden pasar a letras
            let cad = "";
            if (tipo === 84) { for (let i = 0; i < n; i++) cad += String.fromCharCode(dv.getUint16(p + offS + i * 2, true)); }
            else cad = ansi(u8.subarray(p + offS, p + offS + n), (L.s.fuente || {}).charset || 0);
            let dx = null;
            if (offDx && p + offDx + n * 4 <= p + tam) { dx = []; const paso = opc & 0x2000 ? 8 : 4; for (let i = 0; i < n; i++) dx.push(i32(p + offDx + i * paso)); }
            // (un carácter Unicode fuera del plano básico ocupa dos lugares: su avance va en el primero)
            if (dx && Array.from(cad).length !== cad.length) { const r = []; let k = 0; for (const ch of cad) { r.push(dx[k] + (ch.length > 1 ? dx[k + 1] || 0 : 0)); k += ch.length; } dx = r; }
            L.texto(x, y, cad, dx, opc, rc);
            gdi++; break;
          }
          default: break;
        }
      } catch (e) { /* un registro raro no detiene el resto */ }
      p += tam;
    }
    if (!L.dibujado) return hayMas ? { soloMas: true } : null;
    void gdi;
    return { svg: L.svg(vx, vy, vw, vh, wPx, hPx), w: wPx, h: hPx };
  }

  // ------------------------------------------------------------ WMF
  function wmf(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let p = 0, caja = null, pulgada = 1440;
    if (u8.length >= 22 && dv.getUint32(0, true) === 0x9AC6CDD7) {
      caja = [dv.getInt16(6, true), dv.getInt16(8, true), dv.getInt16(10, true), dv.getInt16(12, true)];
      pulgada = dv.getUint16(14, true) || 1440;
      p = 22;
    }
    if (p + 18 > u8.length) return null;
    const tipoCab = dv.getUint16(p, true), tamCab = dv.getUint16(p + 2, true);
    if ((tipoCab !== 1 && tipoCab !== 2) || tamCab !== 9) return null;
    p += 18;
    const i16 = (o) => dv.getInt16(o, true), u16 = (o) => dv.getUint16(o, true), u32 = (o) => dv.getUint32(o, true);
    const L = new Lienzo({ wmf: true, caja });
    // lo que se dibuja sin caja: se mide para el tamaño
    const recs = [];
    let q = p, vueltas = 0;
    while (q + 6 <= u8.length && vueltas++ < 2000000) {
      const tam = u32(q) * 2, fn = u16(q + 4);
      if (tam < 6 || q + tam > u8.length) break;
      recs.push([q, tam, fn]);
      if (fn === 0) break;
      q += tam;
    }
    // sin encabezado con caja: el marco es la ventana vigente cuando se empieza a dibujar (como lo coloca Office)
    let ventOrg = [0, 0], ventExt = null;
    const DIBUJAN = new Set([0x0A32, 0x0521, 0x0324, 0x0325, 0x0538, 0x041B, 0x061C, 0x0418, 0x0817, 0x081A, 0x0830, 0x061D, 0x0213, 0x0F43, 0x0B41, 0x0B23, 0x0940, 0x0922, 0x0D33]);
    for (const [o, , fn] of recs) {
      if (DIBUJAN.has(fn) && ventExt) break;
      if (fn === 0x020B) ventOrg = [i16(o + 8), i16(o + 6)];
      if (fn === 0x020C) ventExt = [i16(o + 8), i16(o + 6)];
    }
    if (!caja) {
      if (!ventExt) return null;
      caja = [Math.min(ventOrg[0], ventOrg[0] + ventExt[0]), Math.min(ventOrg[1], ventOrg[1] + ventExt[1]), Math.max(ventOrg[0], ventOrg[0] + ventExt[0]), Math.max(ventOrg[1], ventOrg[1] + ventExt[1])];
      L.caja = caja;
    }
    const cw = Math.abs(caja[2] - caja[0]), ch = Math.abs(caja[3] - caja[1]);
    if (!cw || !ch) return null;
    const wPx = (cw / pulgada) * 96, hPx = (ch / pulgada) * 96;
    L.unidadPx = cw / Math.max(1, wPx);
    const vx = Math.min(caja[0], caja[2]), vy = Math.min(caja[1], caja[3]);
    const pts = (o, n) => { const r = []; for (let i = 0; i < n; i++) r.push([i16(o + i * 4), i16(o + i * 4 + 2)]); return r; };
    const logfont = (o, fin) => {
      let cara = ""; const ini = o + 18;
      for (let i = ini; i < Math.min(fin, ini + 32); i++) { if (!u8[i]) break; cara += String.fromCharCode(u8[i]); }
      return { t: "fuente", alto: i16(o), ancho: i16(o + 2), escape: i16(o + 4), peso: i16(o + 8), cursiva: u8[o + 10], subrayado: u8[o + 11], tachado: u8[o + 12], charset: u8[o + 13], cara, symbol: u8[o + 13] === 2 || /^symbol$/i.test(cara) };
    };
    const dib = (o, fin) => dibAUrl(u8.subarray(0, fin), o, null, null, null);
    for (const [o, tam, fn] of recs) {
      const d = o + 6, fin = o + tam;
      try {
        switch (fn) {
          case 0x0000: break;
          case 0x0201: L.s.fondo = u32(d); break;
          case 0x0102: L.s.modoFondo = u16(d); break;
          case 0x0103: L.s.modo = u16(d); break;
          case 0x0106: L.s.relleno = u16(d) === 2 ? "nonzero" : "evenodd"; break;
          case 0x0209: L.s.texto = u32(d); break;
          case 0x012E: L.s.alinTexto = u16(d); break;
          case 0x020B: L.s.ventOrg = [i16(d + 2), i16(d)]; break;
          case 0x020C: L.s.ventExt = [i16(d + 2), i16(d)]; break;
          case 0x020F: L.s.ventOrg = [L.s.ventOrg[0] + i16(d + 2), L.s.ventOrg[1] + i16(d)]; break;
          case 0x0410: if (L.s.ventExt) L.s.ventExt = [(L.s.ventExt[0] * i16(d + 6)) / (i16(d + 4) || 1), (L.s.ventExt[1] * i16(d + 2)) / (i16(d) || 1)]; break;
          case 0x020D: case 0x020E: case 0x0211: case 0x0412: break;   // el área de visualización no cambia la imagen
          case 0x0214: L.moverA(i16(d + 2), i16(d)); break;
          case 0x0213: L.lineaA(i16(d + 2), i16(d)); break;
          case 0x001E: L.guardar(); break;
          case 0x0127: L.restaurar(i16(d)); break;
          case 0x012D: L.elegir(u16(d)); break;
          case 0x01F0: L.borrar(u16(d)); break;
          case 0x02FA: L.crear(null, { t: "pluma", estilo: u16(d), ancho: i16(d + 2), color: u32(d + 6) }); break;
          case 0x02FC: L.crear(null, { t: "brocha", estilo: u16(d), color: u32(d + 2), trama: u16(d + 6) }); break;
          case 0x02FB: L.crear(null, logfont(d, fin)); break;
          case 0x0142: { const img = dib(d + 4, fin); L.crear(null, img ? { t: "brocha", estilo: 5, img, color: 0 } : { t: "brocha", estilo: 0, color: 0x808080 }); break; }
          case 0x01F9: case 0x00F7: L.crear(null, { t: fn === 0x00F7 ? "paleta" : "brocha", estilo: 0, color: 0x808080 }); break;
          case 0x06FF: {
            // región: se guarda como rectángulos (sus «scans»)
            const n = i16(d + 10), rects = []; let o2 = d + 20;
            for (let i = 0; i < n && o2 + 4 <= fin; i++) { const c = u16(o2), top = i16(o2 + 2), bot = i16(o2 + 4); for (let k = 0; k < c / 2; k++) rects.push([i16(o2 + 6 + k * 4), top, i16(o2 + 8 + k * 4), bot]); o2 += 6 + c * 2 + 2; }
            L.crear(null, { t: "region", rects });
            break;
          }
          case 0x012C: { const ix = u16(d), r = L.objetos[ix]; if (r && r.t === "region") L.recortarRegion(r.rects, 5, true); else L.s.recorte = null; break; }
          case 0x0416: L.recortarRect(i16(d + 6), i16(d + 4), i16(d + 2), i16(d)); break;
          case 0x0415: L.recortarRect(i16(d + 6), i16(d + 4), i16(d + 2), i16(d), true); break;
          case 0x0324: { const n = i16(d); L.figura(L.poli(pts(d + 2, n), true), true, true); break; }
          case 0x0325: { const n = i16(d); L.figura(L.poli(pts(d + 2, n), false), false, true); break; }
          case 0x0538: {
            const np = u16(d), cuentas = []; for (let i = 0; i < np; i++) cuentas.push(u16(d + 2 + i * 2));
            let o2 = d + 2 + np * 2, dd = "";
            for (const c of cuentas) { dd += L.poli(pts(o2, c), true); o2 += c * 4; }
            L.figura(dd, true, true); break;
          }
          case 0x041B: L.figura(L.rect(i16(d + 6), i16(d + 4), i16(d + 2), i16(d)), true, true); break;
          case 0x061C: L.figura(L.rectRedondo(i16(d + 10), i16(d + 8), i16(d + 6), i16(d + 4), i16(d + 2), i16(d)), true, true); break;
          case 0x0418: L.figura(L.elipse(i16(d + 6), i16(d + 4), i16(d + 2), i16(d)), true, true); break;
          case 0x0817: case 0x081A: case 0x0830: {
            const t = { 0x0817: "arc", 0x081A: "pie", 0x0830: "chord" }[fn];
            L.figura(L.arco(i16(d + 14), i16(d + 12), i16(d + 10), i16(d + 8), i16(d + 6), i16(d + 4), i16(d + 2), i16(d)), t !== "arc", true);
            break;
          }
          case 0x061D: L.patron(i16(d + 10), i16(d + 8), i16(d + 6), i16(d + 4), u32(d)); break;
          case 0x041F: { const x = i16(d + 6), y = i16(d + 4); L.dibujar(L.rect(x, y, x + 1, y + 1), true, false, { estilo: 0, color: u32(d) }); break; }
          case 0x0521: {
            const n = i16(d), bytes = u8.subarray(d + 2, d + 2 + n), o2 = d + 2 + n + (n & 1);
            L.texto(i16(o2 + 2), i16(o2), ansi(bytes, (L.s.fuente || {}).charset || 0), null, 0, null);
            break;
          }
          case 0x0A32: {
            const y = i16(d), x = i16(d + 2), n = i16(d + 4), opc = u16(d + 6);
            let o2 = d + 8, rect = null;
            if (opc & 6) { rect = [i16(o2), i16(o2 + 2), i16(o2 + 4), i16(o2 + 6)]; o2 += 8; }
            const bytes = u8.subarray(o2, o2 + n);
            o2 += n + (n & 1);
            let dx = null;
            if (o2 + n * 2 <= fin) { dx = []; for (let i = 0; i < n; i++) dx.push(i16(o2 + i * 2)); }
            const cs = (L.s.fuente || {}).charset || 0;
            let cad = ansi(bytes, cs);
            // (en las codificaciones de dos bytes, un carácter usa dos avances)
            if (dx && cad.length !== n) { const r = []; let k = 0; for (const ch of cad) { const b = new TextEncoder().encode(ch).length > 1 && cs >= 128 ? 2 : 1; r.push(dx.slice(k, k + b).reduce((a, v) => a + v, 0)); k += b; } dx = r; }
            L.texto(x, y, cad, dx, opc, rect ? [rect[0], rect[1], rect[2], rect[3]] : null);
            break;
          }
          case 0x0F43: {
            const rop = u32(d), srcH = i16(d + 6), srcW = i16(d + 8), yS = i16(d + 10), xS = i16(d + 12), dH = i16(d + 14), dW = i16(d + 16), yD = i16(d + 18), xD = i16(d + 20);
            const img = dib(d + 22, fin);
            if (img) L.imagen(img, xD, yD, dW, dH, xS, yS, srcW, srcH, rop);
            break;
          }
          case 0x0B41: case 0x0B23: {
            const rop = u32(d), sin = tam === 28;
            const srcH = i16(d + 4), srcW = i16(d + 6), yS = i16(d + 8), xS = i16(d + 10), k = sin ? 2 : 0;
            const dH = i16(d + 12 + k), dW = i16(d + 14 + k), yD = i16(d + 16 + k), xD = i16(d + 18 + k);
            const img = sin ? null : dib(d + 20, fin);
            if (img) L.imagen(img, xD, yD, dW, dH, xS, yS, srcW, srcH, rop); else L.patron(xD, yD, dW, dH, rop);
            break;
          }
          case 0x0940: case 0x0922: {
            const rop = u32(d), sin = tam === 24;
            const yS = i16(d + 4), xS = i16(d + 6), k = sin ? 2 : 0;
            const h = i16(d + 8 + k), w = i16(d + 10 + k), yD = i16(d + 12 + k), xD = i16(d + 14 + k);
            const img = sin ? null : dib(d + 16, fin);
            if (img) L.imagen(img, xD, yD, w, h, xS, yS, w, h, rop); else L.patron(xD, yD, w, h, rop);
            break;
          }
          case 0x0D33: {
            const h = i16(d + 10), w = i16(d + 12), yD = i16(d + 16), xD = i16(d + 18), yS = i16(d + 8), xS = i16(d + 14);
            const img = dib(d + 20, fin);
            if (img) L.imagen(img, xD, yD, w, h, xS, yS, w, h, 0xCC0020);
            break;
          }
          default: break;   // ESCAPE (0x0626) y lo demás: no dibujan
        }
      } catch (e) { /* un registro raro no detiene el resto */ }
    }
    if (!L.dibujado) return null;
    return { svg: L.svg(vx, vy, cw, ch, wPx, hPx), w: wPx, h: hPx };
  }

  // ------------------------------------------------------------ lo que usa el resto
  /** ¿Son bytes de un EMF o un WMF? "emf", "wmf" o null. */
  function tipo(u8) {
    if (!u8 || u8.length < 18) return null;
    if (u8.length >= 44 && u8[0] === 1 && u8[1] === 0 && u8[2] === 0 && u8[3] === 0 && u8[40] === 0x20 && u8[41] === 0x45 && u8[42] === 0x4D && u8[43] === 0x46) return "emf";
    if (u8[0] === 0xD7 && u8[1] === 0xCD && u8[2] === 0xC6 && u8[3] === 0x9A) return "wmf";
    if ((u8[0] === 1 || u8[0] === 2) && u8[1] === 0 && u8[2] === 9 && u8[3] === 0) return "wmf";
    return null;
  }
  /** Bytes de un EMF o WMF → { svg, w, h } (w y h en píxeles a 96 ppp) o null si no se pudo. */
  function aSvg(u8) {
    try {
      const t = tipo(u8);
      const r = t === "emf" ? emf(u8) : t === "wmf" ? wmf(u8) : null;
      return r && r.svg ? r : null;
    } catch (e) { if (typeof console !== "undefined") console.warn("No se pudo dibujar el metarchivo", e); return null; }
  }
  const enc = typeof TextEncoder !== "undefined" ? new TextEncoder() : null;
  /** data:image/svg+xml;base64,… del dibujo (o null). */
  function dataUrl(u8) {
    const r = aSvg(u8);
    if (!r || !enc) return null;
    return "data:image/svg+xml;base64," + B64(enc.encode(r.svg));
  }
  const RE_DATOS = /^data:image\/(?:x-emf|emf|x-wmf|wmf|x-ms-wmf|x-ms-emf);base64,/i;
  const cache = new Map();
  /** Un data URL de EMF o WMF → el de su SVG (para dibujarlo); lo demás, igual. Recuerda los últimos. */
  function paraVer(src) {
    if (typeof src !== "string" || !RE_DATOS.test(src)) return src;
    if (cache.has(src)) return cache.get(src);
    let r = null;
    try { const b = atob(src.slice(src.indexOf(",") + 1)), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); r = dataUrl(u); } catch (e) { r = null; }
    // (si no se pudo, un recuadro gris del tamaño, como un objeto que no se ve)
    r = r || "data:image/svg+xml;base64," + btoa('<svg xmlns="http://www.w3.org/2000/svg" width="64" height="48"><rect width="64" height="48" fill="#eee" stroke="#bbb"/></svg>');
    if (cache.size > 300) cache.delete(cache.keys().next().value);
    cache.set(src, r);
    return r;
  }
  NOM.emf = { tipo, aSvg, dataUrl, paraVer, esDatos: (s) => typeof s === "string" && RE_DATOS.test(s), deSymbol };
})();

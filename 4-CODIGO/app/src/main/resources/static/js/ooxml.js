/* NOM · Exportar a formatos de Office (Office Open XML), sin librerías:
   texto → .docx (Word), hoja → .xlsx (Excel), presentación y diagrama → .pptx (PowerPoint).
   Un archivo OOXML es un .zip de XML; aquí se arma comprimido, sin librerías (compresion.js). */
(function () {
  "use strict";
  /** Un texto que viene del archivo o del documento, para usarlo dentro de una expresión regular tal cual. */
  const reEsc = (t) => String(t).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  /**
   * Un símbolo de Word (w:sym): con la fuente Symbol, su letra en Unicode (F06B → κ, como la dibuja Word; js/emf.js);
   * con otra fuente, el carácter tal cual. (Antes quedaba la letra latina —«knbernaw» en lugar de «κνβερναω»— y al
   * descargar un párrafo cambiado salía así en Word.)
   */
  function textoSym(fuente, codigo) {
    const k = parseInt(codigo || "20", 16);
    if (/^symbol$/i.test(fuente || "") && NOM.emf && NOM.emf.deSymbol) return NOM.emf.deSymbol(String.fromCharCode(k >= 0xF000 ? k : k + 0xF000));
    return String.fromCharCode(k & 0xFFF);
  }
  const NOM = window.NOM;
  const F = () => NOM.formulas;
  const enc = new TextEncoder();
  const x = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]))
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  // (el color de cada persona que hizo cambios marcados en un Word: uno por nombre, como Word)
  const COLORES_CAMBIO = ["#B5179E", "#0077B6", "#2A9D8F", "#E76F51", "#6A4C93", "#C1121F", "#3A86FF", "#8338EC"];

  // ------------------------------------------------------------ zip (comprimido)
  const crc32 = (b) => NOM.compresion.crc32(b);
  const YA_COMPRIMIDO = /\.(png|jpe?g|gif|webp|mp4|m4a|mp3|wmv|avi|zip|docx|xlsx|pptx|odttf)$/i;
  /**
   * Arma un .zip. Cada archivo: { nombre, datos (texto o bytes) }. Los XML van comprimidos
   * (DEFLATE propio, compresion.js); las imágenes, que ya vienen comprimidas, se guardan tal cual.
   */
  /**
   * Relaciones externas de lo que se descarga: se quedan las páginas web, el correo y las rutas de una unidad local; una
   * plantilla de otro sitio, una ruta de red u otro esquema pasa a «about:blank». Un archivo sin nada de eso sale igual,
   * byte por byte.
   */
  function destinoPeligroso(tipo, destino) {
    const d = String(destino || "").trim(), t = String(tipo || "");
    if (/^(https?:\/\/|mailto:)/i.test(d)) return /\/attachedTemplate$/i.test(t);
    if (/^file:\/\/\/[A-Za-z]:[\\/]/i.test(d) || /^[A-Za-z]:[\\/]/.test(d)) return false;   // una unidad local
    return true;   // UNC, file://equipo, esquemas que abren programas, rutas relativas raras
  }
  function relsSeguras(datos) {
    const texto = typeof datos === "string" ? datos : new TextDecoder().decode(datos);
    if (!/TargetMode\s*=\s*["']External/i.test(texto)) return datos;
    let doc;
    try { doc = new DOMParser().parseFromString(texto, "application/xml"); } catch (e) { return datos; }
    if (!doc || doc.getElementsByTagName("parsererror").length) return datos;
    let cambio = false;
    for (const r of Array.from(doc.getElementsByTagNameNS("*", "Relationship"))) {
      if (r.getAttribute("TargetMode") !== "External") continue;
      if (destinoPeligroso(r.getAttribute("Type"), r.getAttribute("Target"))) { r.setAttribute("Target", "about:blank"); cambio = true; }
    }
    if (!cambio) return datos;
    const t = new XMLSerializer().serializeToString(doc);
    const xml = /^<\?xml/.test(t) ? t : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + t;
    return typeof datos === "string" ? xml : enc.encode(xml);
  }

  function zip(archivos, opc) {
    const partes = [], centrales = [];
    let off = 0;
    for (const a of archivos) {
      const nombre = enc.encode(a.nombre);
      // Parte que llegó dañada del archivo original: va tal cual venía (comprimida, con su CRC y su tamaño).
      const cr = a.crudo;
      const limpio = !cr && /\.rels$/i.test(a.nombre) && a.datos != null ? relsSeguras(a.datos) : a.datos;
      const datos = cr ? { length: cr.tam } : typeof limpio === "string" ? enc.encode(limpio) : limpio, crc = cr ? cr.crc : crc32(datos);
      let cuerpo = cr ? cr.datos : datos, metodo = cr ? 8 : 0;
      if (!cr && datos.length > 40 && !a.sinComprimir) {
        // Lo que ya viene comprimido (fotos) se prueba rápido: algunos PNG mal comprimidos todavía bajan bastante.
        const ya = YA_COMPRIMIDO.test(a.nombre);
        const c = NOM.compresion.comprimir(datos, { nivel: ya ? 2 : (opc && opc.nivel) || (datos.length > 3000000 ? 6 : 9) });
        if (c.length < datos.length * (ya ? 0.97 : 1)) { cuerpo = c; metodo = 8; }
      }
      const loc = new DataView(new ArrayBuffer(30));
      loc.setUint32(0, 0x04034b50, true); loc.setUint16(4, 20, true); loc.setUint16(6, 0x0800, true); loc.setUint16(8, metodo, true);
      loc.setUint16(12, 0x21, true);   // fecha fija (1980-01-01): el mismo documento da siempre el mismo archivo
      loc.setUint32(14, crc, true); loc.setUint32(18, cuerpo.length, true); loc.setUint32(22, datos.length, true);
      loc.setUint16(26, nombre.length, true);
      partes.push(new Uint8Array(loc.buffer), nombre, cuerpo);
      const cen = new DataView(new ArrayBuffer(46));
      cen.setUint32(0, 0x02014b50, true); cen.setUint16(4, 20, true); cen.setUint16(6, 20, true); cen.setUint16(8, 0x0800, true); cen.setUint16(10, metodo, true);
      cen.setUint16(14, 0x21, true);
      cen.setUint32(16, crc, true); cen.setUint32(20, cuerpo.length, true); cen.setUint32(24, datos.length, true);
      cen.setUint16(28, nombre.length, true); cen.setUint32(42, off, true);
      centrales.push(new Uint8Array(cen.buffer), nombre);
      off += 30 + nombre.length + cuerpo.length;
    }
    const tamCentral = centrales.reduce((n, p) => n + p.length, 0);
    const fin = new DataView(new ArrayBuffer(22));
    fin.setUint32(0, 0x06054b50, true); fin.setUint16(8, archivos.length, true); fin.setUint16(10, archivos.length, true);
    fin.setUint32(12, tamCentral, true); fin.setUint32(16, off, true);
    const todo = [...partes, ...centrales, new Uint8Array(fin.buffer)];
    const r = new Uint8Array(todo.reduce((n, p) => n + p.length, 0));
    let i = 0; for (const p of todo) { r.set(p, i); i += p.length; }
    return r;
  }

  // ------------------------------------------------------------ utilidades
  function ordenados(doc, pagina) {
    const pags = doc.paginas();
    return (pagina ? [pagina] : pags).flatMap((p) => doc.deLaPagina(p));
  }

  /** Qué imagen es, por sus primeros bytes: png, jpeg, gif, bmp, emf, wmf o null. */
  function tipoImagen(b) {
    if (!b || b.length < 24) return null;
    if (b[0] === 0x89 && b[1] === 0x50) return "png";
    if (b[0] === 0xFF && b[1] === 0xD8) return "jpeg";
    if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return "gif";
    if (b[0] === 0x42 && b[1] === 0x4D) return "bmp";
    return NOM.emf ? NOM.emf.tipo(b) : null;
  }
  const MIME_IMG = { png: "image/png", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", emf: "image/x-emf", wmf: "image/x-wmf" };
  /** Una imagen del documento (data:…;base64 o del almacén) → { bytes, ext, w, h }: PNG, JPEG, GIF, BMP, EMF o WMF
      (lo que Office guarda tal cual; un EMF o WMF importado sale idéntico). Otras se omiten. */
  function imagen(src) {
    let bytes;
    if (NOM.binarios && NOM.binarios.esRef(src)) {
      // del almacén de binarios (ya precargado antes de exportar)
      bytes = NOM.binarios.bytesYa(src);
    } else {
      const m = /^data:[a-z0-9.+\/-]+;base64,(.+)$/i.exec(src || "");
      if (!m) return null;
      const bin = atob(m[1]);
      bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    }
    const ext = tipoImagen(bytes);
    if (!ext) return null;
    const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let w = 800, h = 600;
    if (ext === "png") { w = v.getUint32(16); h = v.getUint32(20); }
    else if (ext === "gif") { w = v.getUint16(6, true); h = v.getUint16(8, true); }
    else if (ext === "bmp") { w = Math.abs(v.getInt32(18, true)); h = Math.abs(v.getInt32(22, true)); }
    else if (ext === "emf") { const fw = v.getInt32(32, true) - v.getInt32(24, true), fh = v.getInt32(36, true) - v.getInt32(28, true); if (fw > 0 && fh > 0) { w = (fw / 2540) * 96; h = (fh / 2540) * 96; } }
    else if (ext === "wmf") { if (v.getUint32(0, true) === 0x9AC6CDD7) { const k = v.getUint16(14, true) || 1440; w = (Math.abs(v.getInt16(10, true) - v.getInt16(6, true)) / k) * 96; h = (Math.abs(v.getInt16(12, true) - v.getInt16(8, true)) / k) * 96; } }
    else {
      for (let i = 2; i < bytes.length - 9;) {
        if (bytes[i] !== 0xFF) { i++; continue; }
        const mk = bytes[i + 1], len = (bytes[i + 2] << 8) | bytes[i + 3];
        if (mk >= 0xC0 && mk <= 0xC3) { h = (bytes[i + 5] << 8) | bytes[i + 6]; w = (bytes[i + 7] << 8) | bytes[i + 8]; break; }
        i += 2 + len;
      }
    }
    return { bytes, ext, w: Math.max(1, Math.round(w)), h: Math.max(1, Math.round(h)) };
  }

  /** HTML de un párrafo (b, i, u, br) → tramos { t, b, i, u, br }. */
  function tramos(html) {
    const d = document.createElement("div");
    d.innerHTML = html || "";
    const r = [];
    (function rec(n, st) {
      n.childNodes.forEach((c) => {
        if (c.nodeType === 3) { if (c.nodeValue) r.push({ t: c.nodeValue, ...st }); return; }
        if (c.nodeType !== 1) return;
        const tag = c.tagName.toLowerCase();
        if (tag === "br") { r.push({ br: true }); return; }
        rec(c, { b: st.b || tag === "b" || tag === "strong", i: st.i || tag === "i" || tag === "em", u: st.u || tag === "u" });
      });
    })(d, {});
    return r;
  }

  /** Fórmula de NOM (=SUMA(A1;A2)) → fórmula de Excel (SUM(A1,A2)), sin el "=". */
  const EN = { SUMA: "SUM", PROMEDIO: "AVERAGE", CONTAR: "COUNT", CONTARA: "COUNTA", SI: "IF", ESERROR: "ISERROR", SIERROR: "IFERROR",
    Y: "AND", O: "OR", NO: "NOT", REDONDEAR: "ROUND", RAIZ: "SQRT", POTENCIA: "POWER", CONCATENAR: "CONCATENATE", MAYUSC: "UPPER", MINUSC: "LOWER", LARGO: "LEN",
    BUSCARV: "VLOOKUP", "CONTAR.SI": "COUNTIF", "SUMAR.SI": "SUMIF", HOY: "TODAY", FECHA: "DATE", DIAS: "DAYS", "DÍAS": "DAYS" };
  // Los errores de NOM con el nombre que acepta Excel en un archivo (los demás Excel los da al recalcular)
  const ERR_XL = { "#VALOR!": "#VALUE!", "#NOMBRE?": "#NAME?", "#NULO!": "#NULL!", "#DESBORDAMIENTO!": "#VALUE!", "#CALC!": "#VALUE!", "#CICLO!": "#VALUE!" };
  // Matrices dinámicas en el archivo: la parte de metadatos de celda con el tipo XLDAPR (cm="1" en la celda)
  const T_META = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
  const CT_META = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
  const META_DA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;
  /** El cm (desde 1) de la matriz dinámica en la parte de metadatos de un libro, o 0 si no lo tiene. */
  function cmDeMetadatos(txt) {
    const tipos = [...txt.matchAll(/<(?:\w+:)?metadataType\b[^>]*?\bname="([^"]*)"/g)].map((q) => q[1]);
    const t = tipos.indexOf("XLDAPR") + 1;
    const fm = /<(?:\w+:)?futureMetadata\b[^>]*?\bname="XLDAPR"[^>]*>([\s\S]*?)<\/(?:\w+:)?futureMetadata>/.exec(txt);
    const cmx = /<(?:\w+:)?cellMetadata\b[^>]*>([\s\S]*?)<\/(?:\w+:)?cellMetadata>/.exec(txt);
    if (!t || !fm || !cmx) return 0;
    const v = fm[1].split(/<(?:\w+:)?bk\b[^>]*>/).slice(1).findIndex((b) => /fDynamic="1"/.test(b));
    if (v < 0) return 0;
    const k = cmx[1].split(/<(?:\w+:)?bk\b[^>]*>/).slice(1).findIndex((b) => { const rc = /<(?:\w+:)?rc\b([^>]*)>/.exec(b); return rc && new RegExp(`\\bt="${t}"`).test(rc[1]) && new RegExp(`\\bv="${v}"`).test(rc[1]); });
    return k < 0 ? 0 : k + 1;
  }
  /** Nombres de LET y parámetros de LAMBDA de una fórmula (Excel los guarda con «_xlpm.»). */
  function nombresLambda(s) {
    const out = new Set(), re = /(?:^|[^A-Za-z0-9_.])(LET|LAMBDA)\s*\(/gi;
    let m;
    while ((m = re.exec(s))) {
      let i = m.index + m[0].length, prof = 0, ini = i, enT = false;
      const partes = [];
      for (; i < s.length; i++) {
        const c = s[i];
        if (c === '"') { enT = !enT; continue; }
        if (enT) continue;
        if (c === "(" || c === "{") prof++;
        else if (c === ")" || c === "}") { if (!prof) break; prof--; }
        else if ((c === "," || c === ";") && !prof) { partes.push(s.slice(ini, i).trim()); ini = i + 1; }
      }
      partes.push(s.slice(ini, i).trim());
      const esLet = m[1].toUpperCase() === "LET";
      partes.forEach((p, k) => { if (k === partes.length - 1 || (esLet && k % 2)) return; const q = /^\[?([A-Za-z_][A-Za-z0-9_.]*)\]?$/.exec(p); if (q) out.add(q[1].toUpperCase()); });
    }
    return out;
  }
  const RE_REF_FIN = /((?:'[^']*(?:''[^']*)*'|[A-Za-zÁÉÍÓÚÑáéíóúñ_][A-Za-z0-9ÁÉÍÓÚÑáéíóúñ_.]*)!)?\$?[A-Za-z]{1,3}\$?\d+$/;
  const RE_REF_INI = /^((?:'[^']*(?:''[^']*)*'|[A-Za-zÁÉÍÓÚÑáéíóúñ_][A-Za-z0-9ÁÉÍÓÚÑáéíóúñ_.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/;
  /** Una fórmula de NOM (en español) como la guarda Excel. xlv: versión del importador con que llegó la hoja
      (antes de la 5 las matrices escritas se guardaron con «;» entre columnas). */
  function formulaExcel(f, xlv, vieja) {
    let r = "", enTexto = false, llaves = 0;
    const s = f.replace(/^=/, ""), viejo = +xlv > 0 && +xlv < 5;
    const params = /LET|LAMBDA/i.test(s) ? nombresLambda(s) : null;
    for (let i = 0; i < s.length; i++) {
      const c = s[i];
      if (c === '"') { enTexto = !enTexto; r += c; continue; }
      if (enTexto) { r += c; continue; }
      // {1,2;3,4}: la coma (o «\\») separa columnas y el punto y coma filas, igual que en Excel
      if (!viejo && (c === "{" || c === "}")) { llaves += c === "{" ? 1 : -1; r += c; continue; }
      if (llaves > 0) { r += c === "\\" ? "," : c; continue; }
      if (c === ";") { r += ","; continue; }
      // [y]: parámetro de LAMBDA que se puede omitir (Excel: _xlop.y)
      if (c === "[" && params) { const q = /^\[([A-Za-z_][A-Za-z0-9_.]*)\]/.exec(s.slice(i)); if (q && params.has(q[1].toUpperCase())) { r += "_xlop." + q[1]; i += q[0].length - 1; continue; } }
      // A1#: Excel lo guarda como _xlfn.ANCHORARRAY(A1)
      if (c === "#") { const q = RE_REF_FIN.exec(r); if (q && !/[A-Za-z0-9_.]/.test(r[q.index - 1] || "")) { r = r.slice(0, q.index) + "_xlfn.ANCHORARRAY(" + q[0] + ")"; continue; } }
      // @: la intersección implícita, _xlfn.SINGLE(…)
      if (c === "@") {
        const resto = s.slice(i + 1), q = RE_REF_INI.exec(resto), fn = /^[A-Za-zÁÉÍÓÚÑáéíóúñ_][A-Za-z0-9ÁÉÍÓÚÑáéíóúñ_.]*\s*\(/.exec(resto);
        // (en una fórmula del Excel viejo la intersección implícita no se escribe: es como se calcula ahí)
        if (vieja) continue;
        if (resto[0] === "(") { r += "_xlfn.SINGLE"; continue; }
        if (q && !fn) { r += "_xlfn.SINGLE(" + q[0] + ")"; i += q[0].length; continue; }
        if (fn) {
          let j = fn[0].length, prof = 1, t = false;
          for (; j < resto.length && prof; j++) { const ch = resto[j]; if (ch === '"') t = !t; else if (!t) { if (ch === "(") prof++; else if (ch === ")") prof--; } }
          r += "_xlfn.SINGLE(" + formulaExcel("=" + resto.slice(0, j), xlv, vieja) + ")"; i += j; continue;
        }
      }
      const m = /^[A-Za-zÁÉÍÓÚÑáéíóúñ_][A-Za-z0-9ÁÉÍÓÚÑáéíóúñ_.]*(?=\s*\()/.exec(s.slice(i));
      if (m && params && params.has(m[0].toUpperCase()) && !/[A-Za-z0-9_.$]/.test(s[i - 1] || "")) { r += "_xlpm." + m[0]; i += m[0].length - 1; continue; }
      // (lo que no es una función —un nombre definido que es una función propia: Doble(5)— se queda como se escribió)
      if (m) { const n = m[0].toUpperCase(), es = (F().ALIAS || {})[n] || n; r += (F().INGLES || {})[es] || EN[n] || (/^_xl/i.test(m[0]) ? n : m[0]); i += m[0].length - 1; continue; }
      if (params) {
        const q = /^[A-Za-zÁÉÍÓÚÑáéíóúñ_][A-Za-z0-9ÁÉÍÓÚÑáéíóúñ_.]*/.exec(s.slice(i));
        if (q && params.has(q[0].toUpperCase()) && !/[A-Za-z0-9_.$!:'\]]/.test(s[i - 1] || "") && s[i + q[0].length] !== "!") { r += "_xlpm." + q[0]; i += q[0].length - 1; continue; }
      }
      const b = /^(VERDADERO|FALSO)(?![A-Za-z0-9_.(])/i.exec(s.slice(i));
      if (b && !/[A-Za-z0-9_.]/.test(s[i - 1] || "")) { r += b[1].toUpperCase() === "VERDADERO" ? "TRUE" : "FALSE"; i += b[0].length - 1; continue; }
      r += c;
    }
    return r;
  }

  const CT = (extra) => XML + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="wmf" ContentType="image/x-wmf"/>${extra}</Types>`;
  const RELS = (rels) => XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((r, i) =>
    `<Relationship Id="${r.id || "rId" + (i + 1)}" Type="${r.tipo}" Target="${x(r.destino)}"${r.externo ? ' TargetMode="External"' : ""}/>`).join("")}</Relationships>`;
  const T = { doc: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    img: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    estilos: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    num: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    hoja: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    diapo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    maestro: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
    diseno: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    tema: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    conf: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    enlace: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    encabezado: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    pie: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    nucleo: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" };
  const EMU = 9525;   // 1 px a 96 ppp
  /**
   * La protección de edición de un elemento de Office (Word documentProtection/writeProtection, Excel sheetProtection,
   * workbookProtection y fileSharing, PowerPoint modifyVerifier): con qué se hashea la contraseña y su hash, para
   * comprobarla después (js/cifrado-office.js). `word`: la de «restringir edición» de Word pasa antes por su clave vieja.
   */
  function proteccionDe(el, word, prefijo) {
    if (!el) return null;
    const a = (k) => { for (const n of [prefijo ? prefijo + k[0].toUpperCase() + k.slice(1) : k, k]) { const v = el.getAttribute(n) || el.getAttribute("w:" + n) || el.getAttributeNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", n); if (v) return v; } return null; };
    const p = {};
    if (a("hashValue")) { p.alg = a("algorithmName"); p.hash = a("hashValue"); p.sal = a("saltValue"); p.vueltas = +(a("spinCount") || 0); }
    else if (a("hash") || a("hashData")) { p.sid = +(a("cryptAlgorithmSid") || 4); p.hash = a("hash") || a("hashData"); p.sal = a("salt") || a("saltData"); p.vueltas = +(a("cryptSpinCount") || a("spinCount") || 0); }
    else if (a("password")) p.viejo = a("password");
    if (word && p.hash) p.word = true;
    return Object.keys(p).length ? p : {};
  }
  // Lo que una hoja protegida de Excel puede dejar hacer (en el archivo, «0» = permitido; sin el atributo: no se puede).
  const PERMISOS_HOJA = ["formatCells", "formatColumns", "formatRows", "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows", "sort", "autoFilter", "pivotTables"];
  /**
   * Los escenarios de una hoja de Excel (<scenarios>, Administrador de escenarios): { lista: [{ n (nombre), c: [[celda,
   * valor]], com (comentario), b (bloqueado), o (oculto), u (usuario) }], actual, mostrado, sqref } —sqref: las celdas de
   * resultado del último resumen—. null si no hay.
   */
  function escenariosDe(se) {
    const si = (n, a) => /^(1|true)$/.test(n.getAttribute(a) || "");
    const lista = [...se.children].filter((q) => q.localName === "scenario").map((sc) => {
      const e = { n: sc.getAttribute("name") || "", c: [...sc.children].filter((q) => q.localName === "inputCells" && !si(q, "deleted")).map((q) => [String(q.getAttribute("r") || "").replace(/\$/g, "").toUpperCase(), q.getAttribute("val") || ""]) };
      const com = sc.getAttribute("comment"), u = sc.getAttribute("user");
      if (com) e.com = com;
      if (si(sc, "locked")) e.b = 1;
      if (si(sc, "hidden")) e.o = 1;
      if (u) e.u = u;
      return e;
    });
    if (!lista.length) return null;
    const r = { lista }, ac = se.getAttribute("current"), mo = se.getAttribute("show"), sq = se.getAttribute("sqref");
    if (ac != null && ac !== "") r.actual = +ac;
    if (mo != null && mo !== "") r.mostrado = +mo;
    if (sq) r.sqref = sq;
    return r;
  }
  /** <scenarios> de una hoja (lo contrario de escenariosDe); "" si no hay escenarios. */
  function xmlEscenarios(e) {
    if (!e || !Array.isArray(e.lista) || !e.lista.length) return "";
    const n = e.lista.length, a = [];
    if (e.actual != null && +e.actual >= 0 && +e.actual < n) a.push(`current="${+e.actual}"`);
    if (e.mostrado != null && +e.mostrado >= 0 && +e.mostrado < n) a.push(`show="${+e.mostrado}"`);
    if (e.sqref) a.push(`sqref="${x(e.sqref)}"`);
    return `<scenarios${a.length ? " " + a.join(" ") : ""}>` + e.lista.map((sc) => { const cs = Array.isArray(sc.c) ? sc.c : [];
      return `<scenario name="${x(sc.n)}"${sc.b ? ' locked="1"' : ""}${sc.o ? ' hidden="1"' : ""} count="${cs.length}"${sc.u ? ` user="${x(sc.u)}"` : ""}${sc.com ? ` comment="${x(sc.com)}"` : ""}>`
        + cs.map(([r, v]) => `<inputCells r="${x(r)}" val="${x(v == null ? "" : String(v))}"/>`).join("") + "</scenario>"; }).join("") + "</scenarios>";
  }
  // (una imagen puesta en NOM sobre una hoja: en su celda, de su tamaño en píxeles, como Insertar → Imágenes de Excel)
  const T_IMAGEN_XL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  const ctImagen = (ext) => ({ png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp", bmp: "image/bmp", svg: "image/svg+xml" }[String(ext).toLowerCase()] || "image/png");
  const anclaImagenXl = (rid, n, b) => `<xdr:oneCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + `<xdr:from><xdr:col>${Math.max(0, +(b.de && b.de.c) || 0)}</xdr:col><xdr:colOff>${Math.round((+b.dx || 0) * 9525)}</xdr:colOff><xdr:row>${Math.max(0, +(b.de && b.de.f) || 0)}</xdr:row><xdr:rowOff>${Math.round((+b.dy || 0) * 9525)}</xdr:rowOff></xdr:from>`
    + `<xdr:ext cx="${Math.round((+b.w || 100) * 9525)}" cy="${Math.round((+b.h || 100) * 9525)}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${2000 + n}" name="Imagen ${n}" descr="${x(b.alt || "")}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>`
    + `<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round((+b.w || 100) * 9525)}" cy="${Math.round((+b.h || 100) * 9525)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
  // ------------------------------------------------------------ minigráficos de Excel (x14:sparklineGroups)
  const URI_MG = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
  const COLORES_MG = [["serie", "colorSeries"], ["neg", "colorNegative"], ["eje", "colorAxis"], ["marc", "colorMarkers"], ["prim", "colorFirst"], ["ult", "colorLast"], ["alto", "colorHigh"], ["bajo", "colorLow"]];
  const BANDERAS_MG = [["marcas", "markers"], ["alto", "high"], ["bajo", "low"], ["prim", "first"], ["ult", "last"], ["neg", "negative"], ["eje", "displayXAxis"], ["ocultas", "displayHidden"], ["rtl", "rightToLeft"]];
  /** El ext con los minigráficos de una hoja (t.mg: [{ tipo, s: [[celda, datos]], c: {colores}, marcas…, grosor, vacias }]). */
  function xmlMinigraficos(mg) {
    if (!Array.isArray(mg) || !mg.length) return "";
    const col = (k, el, g) => (g.c && g.c[k] ? `<x14:${el} rgb="FF${x(String(g.c[k]).replace("#", "").toUpperCase())}"/>` : "");
    return `<ext uri="${URI_MG}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">`
      + mg.map((g) => {
        const at2 = [];
        if (g.grosor && Math.abs(+g.grosor - 0.75) > 0.001) at2.push(`lineWeight="${+g.grosor}"`);
        if (g.tipo === "column" || g.tipo === "stacked") at2.push(`type="${g.tipo}"`);
        at2.push(`displayEmptyCellsAs="${g.vacias === "zero" || g.vacias === "span" ? g.vacias : "gap"}"`);
        BANDERAS_MG.forEach(([k, a]) => { if (g[k]) at2.push(`${a}="1"`); });
        return `<x14:sparklineGroup ${at2.join(" ")}>${COLORES_MG.map(([k, el]) => col(k, el, g)).join("")}<x14:sparklines>`
          + (g.s || []).map(([celda, f]) => `<x14:sparkline><xm:f>${x(f)}</xm:f><xm:sqref>${x(celda)}</xm:sqref></x14:sparkline>`).join("") + "</x14:sparklines></x14:sparklineGroup>";
      }).join("") + "</x14:sparklineGroups></ext>";
  }
  /** El extLst de una hoja con los minigráficos de t.mg en vez de los que tenía (lo demás del extLst se queda). */
  function extConMinigraficos(interior, mg) {
    const sin = String(interior || "").replace(new RegExp(`<(\\w+:)?ext\\b[^>]*uri="${URI_MG.replace(/[{}]/g, "\\$&")}"[\\s\\S]*?</(\\w+:)?ext>`, "i"), "");
    return sin + xmlMinigraficos(mg);
  }
  /** La protección de una hoja de Excel (sheetProtection): { p (contraseña), permite: [...], sinSelBloq, sinSelLibre, objetos, escenarios }. */
  function protHojaDe(sp) {
    const si = (k) => /^(1|true)$/.test(sp.getAttribute(k) || ""), no = (k) => /^(0|false)$/.test(sp.getAttribute(k) || "");
    const r = { p: proteccionDe(sp, false) || {} }, permite = PERMISOS_HOJA.filter(no);
    if (permite.length) r.permite = permite;
    if (si("selectLockedCells")) r.sinSelBloq = true;
    if (si("selectUnlockedCells")) r.sinSelLibre = true;
    if (si("objects")) r.objetos = true;
    if (si("scenarios")) r.escenarios = true;
    return r;
  }
  function xmlProtHoja(pr) {
    const p = pr.p || {}, a = [];
    if (p.hash) a.push(`algorithmName="${x(p.alg || "SHA-512")}" hashValue="${x(p.hash)}" saltValue="${x(p.sal || "")}" spinCount="${+p.vueltas || 100000}"`);
    else if (p.viejo) a.push(`password="${x(p.viejo)}"`);
    a.push('sheet="1"');
    if (pr.objetos) a.push('objects="1"');
    if (pr.escenarios) a.push('scenarios="1"');
    (pr.permite || []).filter((k) => PERMISOS_HOJA.includes(k)).forEach((k) => a.push(`${k}="0"`));
    if (pr.sinSelBloq) a.push('selectLockedCells="1"');
    if (pr.sinSelLibre) a.push('selectUnlockedCells="1"');
    return `<sheetProtection ${a.join(" ")}/>`;
  }
  /** «Marcado como final» (Word y PowerPoint): la propiedad _MarkAsFinal del archivo. */
  async function marcadoFinal(z) {
    const c = await z.texto("docProps/custom.xml");
    return !!c && /name="_MarkAsFinal"[^>]*>\s*<vt:bool>\s*(true|1)\s*<\/vt:bool>/i.test(c);
  }
  /** Color de DrawingML con los cambios de Office (tint y shade en RGB lineal; lumMod, lumOff, satMod y hueMod en HSL)
      con los colores de un tema dado ({ dk1, lt1, accent1… } en hexadecimal). Lo usan las gráficas de Word. */
  function colorConTema(fill, tema) {
    const hijos2 = (n) => [...(n ? n.children : [])];
    const c = fill && (/Clr$/.test(fill.localName || "") ? fill : fill.getElementsByTagNameNS("*", "srgbClr")[0] || fill.getElementsByTagNameNS("*", "schemeClr")[0] || fill.getElementsByTagNameNS("*", "sysClr")[0]);
    if (!c) return null;
    const mapa = Object.assign({ tx1: tema.dk1, bg1: tema.lt1, tx2: tema.dk2, bg2: tema.lt2 }, tema);
    let hx = c.localName === "srgbClr" ? c.getAttribute("val") : c.localName === "schemeClr" ? mapa[c.getAttribute("val")] : c.localName === "sysClr" ? c.getAttribute("lastClr") : null;
    if (!/^[0-9a-f]{6}$/i.test(hx || "")) return null;
    hx = hx.toUpperCase();
    for (const t of hijos2(c)) {
      if (t.localName !== "shade" && t.localName !== "tint") continue;
      const k = +t.getAttribute("val") / 100000, lin = [0, 2, 4].map((i) => Math.pow(parseInt(hx.slice(i, i + 2), 16) / 255, 2.3));
      hx = lin.map((x) => (t.localName === "shade" ? x * k : 1 - (1 - x) * k)).map((x) => Math.round(Math.pow(Math.max(0, Math.min(1, x)), 1 / 2.3) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
    }
    const v = (k) => { const e = hijos2(c).find((x) => x.localName === k); return e ? +e.getAttribute("val") / 100000 : null; };
    if (v("lumMod") == null && v("lumOff") == null && v("satMod") == null && v("hueMod") == null) return "#" + hx;
    let [r, g, b] = [0, 2, 4].map((i) => parseInt(hx.slice(i, i + 2), 16) / 255);
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
    let hh = 0, sa = 0, l = (mx + mn) / 2;
    if (mx !== mn) { const d = mx - mn; sa = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); hh = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; hh /= 6; }
    if (v("hueMod") != null) hh = (hh * v("hueMod")) % 1;
    if (v("satMod") != null) sa = Math.max(0, Math.min(1, sa * v("satMod")));
    if (v("lumMod") != null) l *= v("lumMod");
    if (v("lumOff") != null) l += v("lumOff");
    l = Math.max(0, Math.min(1, l));
    const q = l < 0.5 ? l * (1 + sa) : l + sa - l * sa, p2 = 2 * l - q;
    const f = (t3) => { t3 = (t3 + 1) % 1; return t3 < 1 / 6 ? p2 + (q - p2) * 6 * t3 : t3 < 1 / 2 ? q : t3 < 2 / 3 ? p2 + (q - p2) * (2 / 3 - t3) * 6 : p2; };
    return "#" + [f(hh + 1 / 3), f(hh), f(hh - 1 / 3)].map((x) => Math.round(x * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  /** Los colores y las letras de un tema de Office (theme1.xml) para colorConTema y las gráficas. */
  function temaDeXml(xml) {
    const t = {}, f = {};
    if (!xml) return { colores: t, fuentes: f };
    const d = new DOMParser().parseFromString(xml, "application/xml"), cs = d.getElementsByTagNameNS("*", "clrScheme")[0];
    [...(cs ? cs.children : [])].forEach((c) => { const s2 = c.getElementsByTagNameNS("*", "srgbClr")[0], y = c.getElementsByTagNameNS("*", "sysClr")[0], v = s2 ? s2.getAttribute("val") : y ? y.getAttribute("lastClr") : null; if (v && /^[0-9a-f]{6}$/i.test(v)) t[c.localName] = v.toUpperCase(); });
    ["majorFont", "minorFont"].forEach((k, i) => { const n = d.getElementsByTagNameNS("*", k)[0], la = n && n.getElementsByTagNameNS("*", "latin")[0]; if (la && la.getAttribute("typeface")) f[i ? "menor" : "mayor"] = la.getAttribute("typeface"); });
    return { colores: t, fuentes: f };
  }
  // ---- HLS de Windows con enteros (HLSMAX = 240), como lo usa Excel para el tinte de los colores del tema
  const HLSMAX = 240;
  function rgbAHls(R, G, B) {
    const cMax = Math.max(R, G, B), cMin = Math.min(R, G, B), L = Math.floor(((cMax + cMin) * HLSMAX + 255) / 510);
    if (cMax === cMin) return [160, L, 0];
    const S = L <= HLSMAX / 2 ? Math.floor(((cMax - cMin) * HLSMAX + Math.floor((cMax + cMin) / 2)) / (cMax + cMin)) : Math.floor(((cMax - cMin) * HLSMAX + Math.floor((510 - cMax - cMin) / 2)) / (510 - cMax - cMin));
    const d = cMax - cMin, delta = (v) => Math.floor(((cMax - v) * 40 + Math.floor(d / 2)) / d);
    let H = R === cMax ? delta(B) - delta(G) : G === cMax ? 80 + delta(R) - delta(B) : 160 + delta(G) - delta(R);
    if (H < 0) H += HLSMAX;
    if (H > HLSMAX) H -= HLSMAX;
    return [H, L, S];
  }
  function tonoARgb(n1, n2, h) {
    if (h < 0) h += HLSMAX;
    if (h > HLSMAX) h -= HLSMAX;
    if (h < 40) return n1 + Math.trunc(((n2 - n1) * h + 20) / 40);
    if (h < 120) return n2;
    if (h < 160) return n1 + Math.trunc(((n2 - n1) * (160 - h) + 20) / 40);
    return n1;
  }
  function hlsARgb(H, L, S) {
    if (!S) { const v = Math.floor((L * 255) / HLSMAX); return [v, v, v]; }
    const m2 = L <= 120 ? Math.floor((L * (HLSMAX + S) + 120) / HLSMAX) : L + S - Math.floor((L * S + 120) / HLSMAX), m1 = 2 * L - m2;
    return [H + 80, H, H - 80].map((h) => Math.floor((tonoARgb(m1, m2, h) * 255 + 120) / HLSMAX));
  }
  const INDEXADOS_XL = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
    "800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF " +
    "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 " +
    "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333 000000 FFFFFF").split(" ");

  // ------------------------------------------------------------ Word (.docx)
  const TWIP = 15;   // 1 px (96 ppp) = 15 twips
  /** w:lnNumType de Word para [cada, inicio, distancia en px o null, reinicio "p" | "s" | "c"] ("" sin números de línea).
      Word guarda el inicio menos 1 y, sin w:restart, reinicia en cada página. */
  const lnNumTypeXml = (v) => (!Array.isArray(v) ? "" : `<w:lnNumType w:countBy="${Math.max(1, Math.min(100, Math.round(+v[0] || 1)))}"${+v[1] > 1 ? ` w:start="${Math.round(+v[1]) - 1}"` : ""}${v[2] != null && isFinite(+v[2]) ? ` w:distance="${Math.round(+v[2] * TWIP)}"` : ""}${v[3] === "c" ? ' w:restart="continuous"' : v[3] === "s" ? ' w:restart="newSection"' : ""}/>`);
  /** Un w:sectPr con estos números de línea (o sin ninguno), en su lugar del esquema. */
  function conNumLineas(sect, v) {
    if (/<w:sectPrChange\b/.test(sect)) return sect;   // (con un cambio de sección marcado: tal cual)
    const s = sect.replace(/<w:lnNumType\b[^>]*\/>/g, ""), x = lnNumTypeXml(v);
    if (!x) return s;
    if (/^<w:sectPr\b[^>]*\/>$/.test(s)) return s.replace(/\/>$/, ">") + x + "</w:sectPr>";
    const m = /<w:(pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings)\b|<\/w:sectPr>/.exec(s);
    return m ? s.slice(0, m.index) + x + s.slice(m.index) : s;
  }
  /** Lo que es de una celda de tabla (texto y formato): su firma dice si cambió desde que se importó. */
  const CAMPOS_CELDA = ["c", "n", "i", "k", "r", "al", "m", "z", "va", "td", "bxw", "bxd", "h"];
  function firmaCelda(o, ref) {
    const t = JSON.stringify(CAMPOS_CELDA.map((p) => (o[p + ":" + ref] == null || o[p + ":" + ref] === false ? null : o[p + ":" + ref])));
    let h = 0x811C9DC5;
    for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 0x01000193); }
    return (h >>> 0).toString(36) + t.length.toString(36);
  }
  const ESTILO_WORD = { titulo: "Title", subtitulo: "Subtitle", h1: "Heading1", h2: "Heading2", h3: "Heading3", cita: "Quote", codigo: "Codigo" };
  // ------------------------------------------------------------ estilos de párrafo de NOM como estilos de Word
  const ORDEN_RPR = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"];
  const ORDEN_PPR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];
  /** Los hijos de un rPr o pPr con los de «nuevos» en lugar de los suyos, en el orden que pide Word. */
  function conHijos(xml, nuevos, orden) {
    const hijos = [];
    const re = /<w:([A-Za-z]+)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:\1>)/g;
    let m;
    while ((m = re.exec(xml || ""))) if (!(m[1] in nuevos)) hijos.push([m[1], m[0]]);
    for (const k in nuevos) if (nuevos[k]) hijos.push([k, nuevos[k]]);
    return hijos.sort((a, b) => orden.indexOf(a[0]) - orden.indexOf(b[0])).map((h) => h[1]).join("");
  }
  const JC_W = { izq: "left", centro: "center", der: "right", just: "both" };
  /** Lo que un estilo de NOM cambia: { rpr: {tag: xml}, ppr: {tag: xml} } (spacing junta lo que ya tenía). */
  function piezasEstilo(e, spacingViejo) {
    const rpr = {}, ppr = {}, si = (v, tag) => (v === true ? `<w:${tag}/>` : `<w:${tag} w:val="0"/>`);
    if (e.fuente) rpr.rFonts = `<w:rFonts w:ascii="${x(e.fuente)}" w:hAnsi="${x(e.fuente)}" w:eastAsia="${x(e.fuente)}" w:cs="${x(e.fuente)}"/>`;
    if (e.b != null) { rpr.b = si(e.b, "b"); rpr.bCs = si(e.b, "bCs"); }
    if (e.i != null) { rpr.i = si(e.i, "i"); rpr.iCs = si(e.i, "iCs"); }
    if (e.color) rpr.color = `<w:color w:val="${x(String(e.color).replace("#", "").toUpperCase())}"/>`;
    if (+e.tam) { rpr.sz = `<w:sz w:val="${Math.round(+e.tam * 2)}"/>`; rpr.szCs = `<w:szCs w:val="${Math.round(+e.tam * 2)}"/>`; }
    if (e.u != null) rpr.u = `<w:u w:val="${e.u ? "single" : "none"}"/>`;
    if (e.antes != null || e.despues != null || +e.interlineado) {
      const at2 = {};
      String(spacingViejo || "").replace(/w:(\w+)="([^"]*)"/g, (q, k, v) => { at2[k] = v; return q; });
      if (e.antes != null) at2.before = String(Math.round(+e.antes * 20));
      if (e.despues != null) at2.after = String(Math.round(+e.despues * 20));
      if (+e.interlineado) { at2.line = String(Math.round(+e.interlineado * 240)); at2.lineRule = "auto"; }
      ppr.spacing = `<w:spacing ${Object.entries(at2).map(([k, v]) => `w:${k}="${x(v)}"`).join(" ")}/>`;
    }
    if (JC_W[e.alin]) ppr.jc = `<w:jc w:val="${JC_W[e.alin]}"/>`;
    return { rpr, ppr };
  }
  /** La definición de un estilo (<w:style … w:styleId="id">) en un styles.xml, con lo que cambia el estilo de NOM. */
  function parcharEstiloXml(xml, id, e) {
    const re = new RegExp(`<w:style\\b[^>]*w:styleId="${id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*>[\\s\\S]*?</w:style>`);
    const m = re.exec(xml);
    if (!m) return xml;
    let st = m[0];
    const pprM = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(st), rprM = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(st);
    const sp = pprM && /<w:spacing\b[^>]*\/>/.exec(pprM[1]);
    const { rpr, ppr } = piezasEstilo(e, sp && sp[0]);
    if (Object.keys(ppr).length) {
      const nuevo = `<w:pPr>${conHijos(pprM ? pprM[1] : "", ppr, ORDEN_PPR)}</w:pPr>`;
      st = pprM ? st.replace(pprM[0], nuevo) : rprM ? st.replace(rprM[0], nuevo + rprM[0]) : st.replace("</w:style>", nuevo + "</w:style>");
    }
    if (Object.keys(rpr).length) {
      const r2 = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(st), nuevo = `<w:rPr>${conHijos(r2 ? r2[1] : "", rpr, ORDEN_RPR)}</w:rPr>`;
      st = r2 ? st.replace(r2[0], nuevo) : st.replace("</w:style>", nuevo + "</w:style>");
    }
    return xml.replace(m[0], st);
  }
  /** El id de Word de un estilo propio de NOM («u:…»): el que traía de Word, o «Nom…». */
  const idEstiloPropio = (estP, k) => (estP[k] && estP[k].id) || "Nom" + String(k).slice(2).replace(/[^A-Za-z0-9]/g, "");
  /** Un styles.xml de NOM con los estilos de NOM modificados y los propios. */
  function conEstilosDoc(xml, estP) {
    let defs = "";
    for (const [k, e] of Object.entries(estP || {})) {
      if (!e || e.deWord) continue;
      if (ESTILO_WORD[k]) { xml = parcharEstiloXml(xml, ESTILO_WORD[k], e); continue; }
      if (!k.startsWith("u:") || !e.n) continue;
      const id = idEstiloPropio(estP, k);
      const base = `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${x(id)}"><w:name w:val="${x(e.n)}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/></w:style>`;
      defs += parcharEstiloXml(base, id, e);
    }
    return defs ? xml.replace("</w:styles>", defs + "</w:styles>") : xml;
  }

  const ALIN_WORD = { centro: "center", der: "right", just: "both" };
  const hexW = (c) => (/^#[0-9a-f]{6}$/i.test(c || "") ? c.slice(1).toUpperCase() : null);

  /** Documento de texto (secuencia letra por letra) → .docx con estilos de Word reales. */
  function docx(doc) {
    const og = NOM.ooxml.usarOrigen !== false && NOM.origen ? NOM.origen.leer(doc) : null;
    if (og && og.fmt === "docx" && og.h && og.h.u) {
      try { return docxArmar(doc, og); }
      catch (e) { if (typeof console !== "undefined") console.warn("Word: no se pudo partir del archivo original; se arma desde cero.", e); }
    }
    return docxArmar(doc, null);
  }
  /** Huellas de cada unidad del cuerpo (párrafo o tabla de Word) tal como quedó importada: { k: [todo, soloFormatoDePárrafo] }. */
  const sinOx = (pa) => { const r = Object.assign({}, pa); delete r.ox; return r; };
  // Un objeto dentro del texto (ecuación, imagen…) lleva los campos de su bloque: un campo vacío (null) no cuenta, porque
  // al guardar como .nom (y en cualquier copia) un campo vacío simplemente no está, y el párrafo parecía cambiado.
  const sinNulosE = (e) => { const r = {}; for (const k in e) if (e[k] != null) r[k] = e[k]; return r; };
  // (la huella de un párrafo no distingue una sugerencia de su marca visible —subrayado o tachado—: así es la misma al
  // importar, con las sugerencias pintadas, que al descargar el Word, con las sugerencias como cambios marcados)
  const pintada = (a) => { if (!a || (!a.sug && !a.sugBorrar)) return a; const b = Object.assign({}, a); if (b.sug) { b.u = true; b.color = "#1D7A4C"; } if (b.sugBorrar) { b.s = true; b.color = "#C2372B"; } delete b.sug; delete b.sugBorrar; return b; };
  const huellaGrupo = (g) => NOM.origen.huella(g.map((p) => [sinOx(p.pa), p.tramos.map((t) => (t.e ? { e: sinNulosE(t.e), a: pintada(t.a) } : [t.x, pintada(t.a)]))]));
  // (las huellas guardadas hasta la 0.16.0 sí contaban los campos vacíos: también valen)
  const conNulos = (g) => g.some((p) => p.tramos.some((t) => t.e && Object.keys(t.e).some((k) => t.e[k] === null)));
  const huellaGrupoVieja = (g) => NOM.origen.huella(g.map((p) => [sinOx(p.pa), p.tramos.map((t) => (t.e ? { e: t.e, a: pintada(t.a) } : [t.x, pintada(t.a)]))]));
  const mismoGrupo = (g, h) => huellaGrupo(g) === h || (conNulos(g) && huellaGrupoVieja(g) === h);
  function huellasDocx(doc) {
    const ps = NOM.textoRico.parrafosDe(doc), u = {};
    for (let i = 0; i < ps.length;) {
      const k = ps[i].pa.ox;
      let j = i;
      if (k != null) { while (j + 1 < ps.length && ps[j + 1].pa.ox === k) j++; if (u[k] == null) u[k] = [huellaGrupo(ps.slice(i, j + 1)), NOM.origen.huella(sinOx(ps[i].pa))]; }
      i = j + 1;
    }
    return u;
  }
  function docxArmar(doc, og) {
    // Con archivo de origen, lo nuevo se numera después de lo que ya trae (ver "a partir del original").
    let PRE = "", baseAbs = 10, baseNum = 20, baseNota = 0, baseDib = 0;
    const medios = [], rels = [{ id: "rIdEst", tipo: T.estilos, destino: "styles.xml" }, { id: "rIdNum", tipo: T.num, destino: "numbering.xml" }, { id: "rIdConf", tipo: T.conf, destino: "settings.xml" }];
    const ps = NOM.textoRico.parrafosDe(doc, { cambios: true });
    const estP = doc.valor("_doc", "estilosP") || {};
    let nImg = 0, nLink = 0, nCambio = 800000;
    const fechaCambio = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
    // 8.2 Listas: una definición (abstractNum) por formato de número o viñeta, y una instancia (num)
    // por cada lista que empieza (o que empieza en otro número). "continuar" reusa la anterior.
    const abstractos = new Map(), instancias = [];
    let numActual = null, numAnterior = null;
    const FMT_NUM = { "1.": ["decimal", "%1."], "1)": ["decimal", "%1)"], "a.": ["lowerLetter", "%1."], "a)": ["lowerLetter", "%1)"], "A.": ["upperLetter", "%1."], "I.": ["upperRoman", "%1."], "i.": ["lowerRoman", "%1."] };
    const abstracto = (clave) => { if (!abstractos.has(clave)) abstractos.set(clave, abstractos.size + baseAbs); return abstractos.get(clave); };
    const nuevaInstancia = (clave, inicio) => { const id = baseNum + instancias.length; instancias.push({ id, abs: abstracto(clave), inicio }); return id; };
    // Marcadores (los que usan los campos): un id por nombre
    const marcadoresW = {};
    // 8.2 Notas al pie y figuras.
    const notas = [];
    let nFigura = 0;
    const pagDoc = doc.valor("_doc", "pagina") || {};
    // "otro": tamaño que no es de los comunes (vino así de un archivo de Word), con sus medidas en px
    const TAM0 = (pagDoc.tam === "otro" && +pagDoc.w > 100 && +pagDoc.h > 100 ? [+pagDoc.w, +pagDoc.h] : null) || { carta: [816, 1056], a4: [794, 1123], oficio: [816, 1344], a5: [559, 794], media: [529, 816] }[pagDoc.tam] || [816, 1056];
    const mg0 = pagDoc.margen || {}, v0 = (v) => (typeof v === "number" ? v : 96);
    const anchoTexto = Math.round(((pagDoc.orient === "h" ? TAM0[1] : TAM0[0]) - v0(mg0.iz) - v0(mg0.d)) * TWIP);
    const CMT = 1440 / 2.54;   // twips por centímetro
    const rPr = (a) => {
      // en el orden que pide el esquema de Word
      let r = "";
      if (a.fuente) r += `<w:rFonts w:ascii="${x(a.fuente)}" w:hAnsi="${x(a.fuente)}" w:cs="${x(a.fuente)}"/>`;
      if (a.b) r += "<w:b/><w:bCs/>";
      if (a.i) r += "<w:i/><w:iCs/>";
      if (a.mayus) r += "<w:caps/>";
      if (a.vers) r += "<w:smallCaps/>";
      if (a.s) r += "<w:strike/>";
      if (a.oc) r += "<w:vanish/>";
      if (hexW(a.color)) r += `<w:color w:val="${hexW(a.color)}"/>`;
      if (+a.esp) r += `<w:spacing w:val="${Math.round(+a.esp * 20)}"/>`;
      if (+a.esc && +a.esc !== 100) r += `<w:w w:val="${Math.round(+a.esc)}"/>`;
      if (+a.pos) r += `<w:position w:val="${Math.round(+a.pos * 2)}"/>`;
      if (a.tam) r += `<w:sz w:val="${Math.round(a.tam * 2)}"/><w:szCs w:val="${Math.round(a.tam * 2)}"/>`;
      if (a.u) r += '<w:u w:val="single"/>';
      if (hexW(a.fondo)) r += `<w:shd w:val="clear" w:color="auto" w:fill="${hexW(a.fondo)}"/>`;
      if (a.va) r += `<w:vertAlign w:val="${a.va === "sup" ? "superscript" : "subscript"}"/>`;
      return r ? `<w:rPr>${r}</w:rPr>` : "";
    };
    const texto = (t, a) => {
      // Tabuladores y saltos de línea son elementos propios en Word.
      return t.split(/(\t|\u2028)/).filter((z) => z !== "").map((z) => `<w:r>${rPr(a)}${z === "\t" ? "<w:tab/>" : z === "\u2028" ? "<w:br/>" : `<w:t xml:space="preserve">${x(z)}</w:t>`}</w:r>`).join("");
    };
    const dibujo = (src, wpx, alt, rec, rot) => {
      const im = imagen(src);
      if (!im) return "";
      const n = ++nImg, id = "rId" + PRE + "Img" + n, arch = (PRE ? PRE.toLowerCase() + "-" : "") + "imagen" + n + "." + im.ext;
      medios.push({ nombre: "word/media/" + arch, datos: im.bytes });
      rels.push({ id, tipo: T.img, destino: "media/" + arch });
      const r = rec || {}, fw = Math.max(0.05, 1 - (+r.iz || 0) - (+r.d || 0)), fh = Math.max(0.05, 1 - (+r.s || 0) - (+r.i || 0));
      const w = Math.min(wpx || im.w, anchoTexto / TWIP), cx = Math.round(w * EMU), cy = Math.round(w * (im.h / im.w) * (fh / fw) * EMU);
      const recorte = rec ? `<a:srcRect l="${Math.round((+r.iz || 0) * 100000)}" t="${Math.round((+r.s || 0) * 100000)}" r="${Math.round((+r.d || 0) * 100000)}" b="${Math.round((+r.i || 0) * 100000)}"/>` : "";
      const giro = +rot ? ` rot="${Math.round(+rot * 60000)}"` : "";
      return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${baseDib + n}" name="Imagen ${n}" descr="${x(alt || "")}"/>
<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${n}" name="imagen${n}"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="${id}"/>${recorte}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm${giro}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    };
    const BORDES_TABLA = {
      todos: ["top", "left", "bottom", "right", "insideH", "insideV"],
      horizontales: ["top", "bottom", "insideH"],
      externos: ["top", "left", "bottom", "right"],
      ninguno: [],
    };
    const tabla = (b) => {
      const filas = Math.max(1, b.filasTexto || 3), cols = Math.max(1, b.colsTexto || 3);
      const fijos = Array.from({ length: cols }, (_, c) => +b["aw:" + F().colLetra(c)] || 0), suma = fijos.reduce((s2, v) => s2 + v, 0), libres = fijos.filter((v) => !v).length;
      // (el ancho de la tabla: el de Word si lo trae —px o % del texto—; si no, el del texto)
      const pctT = /^p:(\d+(\.\d+)?)$/.exec(String(b.tw || ""));
      const anchoTabla = pctT ? Math.round((anchoTexto * +pctT[1]) / 100) : +b.tw > 20 ? Math.round(+b.tw * 15) : anchoTexto;
      const anchos = fijos.map((v) => Math.round(((v || (libres ? Math.max(0, 100 - suma) / libres : 0)) / 100) * anchoTabla));
      const spans = {}, tapadas = {};
      for (const k in b) {
        if (!k.startsWith("m:") || !/^\d+x\d+$/.test(String(b[k] || ""))) continue;
        const p = F().parseRef(k.slice(2)), [w, h] = String(b[k]).split("x").map(Number);
        if (!p || w * h <= 1) continue;
        spans[k.slice(2)] = [w, h];
        for (let f = p.f; f < p.f + h; f++) for (let c = p.c; c < p.c + w; c++) if (c !== p.c || f !== p.f) tapadas[F().ref(c, f)] = { tl: k.slice(2), primeraCol: c === p.c };
      }
      const lados = BORDES_TABLA[b.bordes || "todos"] || BORDES_TABLA.todos;
      const bordes = `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map((l) => (lados.includes(l) ? `<w:${l} w:val="single" w:sz="4" w:space="0" w:color="8E8E8E"/>` : `<w:${l} w:val="nil"/>`)).join("")}</w:tblBorders>`;
      // márgenes de celda de Word (px: arriba, derecha, abajo, izquierda) y el espaciado de los párrafos de las celdas
      const mar = String(b.cmar || "").split(",").map(Number), conMar = mar.length === 4 && mar.every((v) => v >= 0 && v < 200);
      const tw = (px) => Math.round(px * 15);
      const margenes = conMar ? `<w:tblCellMar><w:top w:w="${tw(mar[0])}" w:type="dxa"/><w:left w:w="${tw(mar[3])}" w:type="dxa"/><w:bottom w:w="${tw(mar[2])}" w:type="dxa"/><w:right w:w="${tw(mar[1])}" w:type="dxa"/></w:tblCellMar>` : "";
      const esp = String(b.cesp || "").split(",").map(Number), conEsp = esp.length === 2 && esp.every((v) => v >= 0 && v < 500);
      const linea = +b.clx > 0 ? ` w:line="${tw(+b.clx)}" w:lineRule="exact"` : +b.cli > 0 ? ` w:line="${Math.round(+b.cli * 240)}" w:lineRule="auto"` : "";
      const espCelda = `<w:spacing${conEsp ? ` w:before="${tw(esp[0])}"` : ""} w:after="${conEsp ? tw(esp[1]) : 0}"${linea}/>`;
      const cabRep = Math.max(+b.cabRep || 0, b.encabezado ? 1 : 0);
      // ancho (de Word: px o % del texto), alineación y sangría
      const pct = /^p:(\d+(\.\d+)?)$/.exec(String(b.tw || ""));
      const tblW = pct ? `<w:tblW w:w="${Math.round(+pct[1] * 50)}" w:type="pct"/>` : +b.tw > 20 ? `<w:tblW w:w="${tw(+b.tw)}" w:type="dxa"/>` : `<w:tblW w:w="${anchoTexto}" w:type="dxa"/>`;
      const tblJc = b.tj === "c" ? '<w:jc w:val="center"/>' : b.tj === "d" ? '<w:jc w:val="right"/>' : "";
      const tblInd = +b.ti ? `<w:tblInd w:w="${tw(+b.ti)}" w:type="dxa"/>` : "";
      // (las líneas de cada celda que trajo de Word van en la celda)
      const conBxw = Object.keys(b).some((k) => k.startsWith("bxw:") || k.startsWith("bxd:"));
      let t = `<w:tbl><w:tblPr><w:tblStyle w:val="TablaNom"/>${tblW}${tblJc}${tblInd}${bordes}<w:tblLayout w:type="fixed"/>${margenes}<w:tblLook w:val="04A0" w:firstRow="${b.encabezado ? 1 : 0}" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${anchos.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
      for (let f = 0; f < filas; f++) {
        // la fila: que no se parta entre páginas, su alto (exacto o mínimo) y si se repite arriba de cada hoja
        const fh = /^([em])(\d+(\.\d+)?)$/.exec(String(b["fh:" + f] || ""));
        const trPr = `${b["fp:" + f] ? "<w:cantSplit/>" : ""}${fh ? `<w:trHeight w:val="${tw(+fh[2])}" w:hRule="${fh[1] === "e" ? "exact" : "atLeast"}"/>` : ""}${f < cabRep ? "<w:tblHeader/>" : ""}`;
        t += `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ""}`;
        for (let c = 0; c < cols; c++) {
          const r = F().ref(c, f), tp = tapadas[r];
          if (tp && !tp.primeraCol) continue;   // tapada a la derecha: va dentro del gridSpan
          const sp = tp ? spans[tp.tl] : spans[r];
          const base = tp ? F().parseRef(tp.tl) : { c, f };
          const refBase = tp ? tp.tl : r;
          const ancho = anchos.slice(c, c + (sp ? sp[0] : 1)).reduce((s2, v) => s2 + v, 0);
          const fondo = hexW(b["r:" + refBase]) || (b.encabezado && f === 0 ? "F2F2F2" : null);
          // bordes de la celda (de Word) en el orden del esquema: top, left, bottom, right, diagonales
          let tcB = "";
          if (conBxw && !tp) {
            const lados = {}; String(b["bxw:" + refBase] || "").split(";").forEach((q) => { const [l, e2, w2, c2] = q.split(":"); if (l) lados[l] = [e2, w2, c2]; });
            String(b["bxd:" + refBase] || "").split(";").forEach((q) => { const [l, e2, w2, c2] = q.split(":"); if (l) lados[l] = [e2, w2, c2]; });
            const linea = (n, v) => (!v ? "" : v[0] === "no" ? `<w:${n} w:val="nil"/>` : `<w:${n} w:val="${{ double: "double", dotted: "dotted", dashed: "dashed" }[v[0]] || "single"}" w:sz="${Math.max(2, Math.min(96, Math.round((+v[1] || 1) * 0.75 * 8)))}" w:space="0" w:color="${hexW(v[2]) || "auto"}"/>`);
            tcB = linea("top", lados.a) + linea("left", lados.i) + linea("bottom", lados.b) + linea("right", lados.d) + linea("tl2br", lados.tl) + linea("tr2bl", lados.tr);
            if (tcB) tcB = `<w:tcBorders>${tcB}</w:tcBorders>`;
          }
          const vAl = b["va:" + refBase] === "c" ? '<w:vAlign w:val="center"/>' : b["va:" + refBase] === "b" ? '<w:vAlign w:val="bottom"/>' : "";
          const tDir = b["td:" + refBase] === "arriba" ? '<w:textDirection w:val="btLr"/>' : b["td:" + refBase] === "abajo" ? '<w:textDirection w:val="tbRl"/>' : "";
          const tcPr = `<w:tcW w:w="${ancho}" w:type="dxa"/>${sp && sp[0] > 1 ? `<w:gridSpan w:val="${sp[0]}"/>` : ""}${sp && sp[1] > 1 ? (tp ? '<w:vMerge/>' : '<w:vMerge w:val="restart"/>') : ""}${tcB}${fondo ? `<w:shd w:val="clear" w:color="auto" w:fill="${fondo}"/>` : ""}${tDir}${vAl}`;
          const a2 = { b: !!b["n:" + refBase] || (b.encabezado && f === 0), i: !!b["i:" + refBase], color: b["k:" + refBase] || null, tam: +b["z:" + refBase] > 0 ? +b["z:" + refBase] : null };
          const jc = ALIN_WORD[b["al:" + refBase]];
          const v = tp ? "" : String(b["c:" + r] || "");
          void base;
          t += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${v.split("\n").map((l) => `<w:p><w:pPr>${espCelda}<w:ind w:firstLine="0"/>${jc ? `<w:jc w:val="${jc}"/>` : ""}</w:pPr>${l ? texto(l, a2) : ""}</w:p>`).join("")}</w:tc>`;
        }
        t += "</w:tr>";
      }
      return t + "</w:tbl>";
    };

    // ---------------------------------------------------------- a partir del archivo original
    /**
     * Arma el cuerpo con las unidades del documento.xml de origen: la que no cambió (misma huella que
     * al importar) sale textual; la que cambió se escribe de nuevo, con el formato de párrafo original
     * si ese no se tocó. Lo que NOM no muestra (marcadores, controles) se queda en su lugar.
     */
    function cuerpoConOrigen() {
      const O = NOM.origen, H = O.huella, pq = paqueteDe(og);
      const relDoc = pq.rels("").find((r) => /\/officeDocument$/.test(r.tipo || ""));
      const DOC = relDoc ? pq.resolver("", relDoc.destino) : "word/document.xml";
      const xd = pq.texto(DOC);
      if (!xd) throw new Error("falta el documento principal");
      const raiz = O.hijosXml(xd).find((h) => O.local(h.n) === "document");
      if (!raiz || raiz.n !== "w:document") throw new Error("el documento usa otro prefijo");
      const body = O.hijoXml(xd, "body", raiz.ci, raiz.cf);
      if (!body) throw new Error("documento sin cuerpo");
      const us = O.hijosXml(xd, body.ci, body.cf), hu = og.h.u || {};
      const relsD = pq.rels(DOC), carpeta = DOC.slice(0, DOC.lastIndexOf("/") + 1);
      const parte = (re) => { const r = relsD.find((q) => re.test(q.tipo || "")); return r ? pq.resolver(DOC, r.destino) : null; };
      // ---- numeración propia que no choque con la del original
      let g = 1;
      while (relsD.some((r) => (r.id || "").startsWith("rIdNom" + g)) || pq.tiene(`${carpeta}media/nom${g}-imagen1.png`) || pq.tiene(`${carpeta}media/nom${g}-imagen1.jpeg`)) g++;
      PRE = "Nom" + g; baseDib = 7000 + g * 500;
      const numRuta = parte(/\/numbering$/), numTxt = numRuta ? pq.texto(numRuta) : null;
      if (numTxt) {
        (numTxt.match(/w:abstractNumId="(\d+)"/g) || []).forEach((m) => { baseAbs = Math.max(baseAbs, +m.replace(/\D/g, "") + 1); });
        (numTxt.match(/<w:num [^>]*w:numId="(\d+)"/g) || []).forEach((m) => { baseNum = Math.max(baseNum, +/w:numId="(\d+)"/.exec(m)[1] + 1); });
      }
      const notaRuta = parte(/\/footnotes$/), notaTxt = notaRuta ? pq.texto(notaRuta) : null;
      if (notaTxt) (notaTxt.match(/<w:footnote [^>]*w:id="(-?\d+)"/g) || []).forEach((m) => { baseNota = Math.max(baseNota, +/w:id="(-?\d+)"/.exec(m)[1]); });
      // ---- estilos: los de NOM se cambian por los del original que se llamen igual
      const estRuta = parte(/\/styles$/), estTxt = estRuta ? pq.texto(estRuta) : null;
      const porNombre = {}, idsEst = new Set();
      let normal = "Normal";
      if (estTxt) {
        const re = /<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g;
        let m;
        while ((m = re.exec(estTxt))) {
          const id = (/w:styleId="([^"]*)"/.exec(m[1]) || [])[1], nm = (/<w:name w:val="([^"]*)"/.exec(m[2]) || [])[1];
          if (id == null) continue;
          idsEst.add(id);
          if (nm && porNombre[nm.toLowerCase()] == null) porNombre[nm.toLowerCase()] = id;
          if (/w:type="paragraph"/.test(m[1]) && /w:default="(1|true)"/.test(m[1])) normal = id;
        }
      }
      const NOMBRES = { Title: "title", Subtitle: "subtitle", Heading1: "heading 1", Heading2: "heading 2", Heading3: "heading 3", Quote: "quote", Codigo: "código", ListParagraph: "list paragraph", Caption: "caption",
        TextoNota: "footnote text", RefNota: "footnote reference", TOCHeading: "toc heading", TDC1: "toc 1", TDC2: "toc 2", TDC3: "toc 3", Hyperlink: "hyperlink", TablaNom: "table grid", Header: "header", Footer: "footer" };
      for (const [k, e] of Object.entries(estP)) if (k.startsWith("u:") && e && e.n && !e.deWord) NOMBRES[idEstiloPropio(estP, k)] = String(e.n).toLowerCase();
      const faltan = new Map();   // id nuestro → id con el que se agrega al original
      const idEstilo = (id) => {
        const nm = NOMBRES[id];
        if (!nm) return id;
        if (porNombre[nm] != null) return porNombre[nm];
        if (!faltan.has(id)) faltan.set(id, idsEst.has(id) ? "Nom" + id : id);
        return faltan.get(id);
      };
      const conEstilos = (t) => t.replace(/<w:(pStyle|rStyle|tblStyle) w:val="([^"]+)"\/>/g, (m, tag, id) => `<w:${tag} w:val="${idEstilo(id)}"/>`);
      // ---- objetos que salen tal cual (imagen, ecuación, llamada de nota) si su bloque no cambió
      const igualQueAlImportar = (ref) => { const hb = og.h.b && og.h.b[ref]; return !!hb && (Array.isArray(hb) ? hb[0] : hb) === NOM.formatos.huellaBloque(doc, ref); };
      const objetoDe = (u) => (e) => {
        if (!e.oe || (e.tipo !== "objeto" && (!e.ref || !igualQueAlImportar(e.ref)))) return null;
        const k = String(e.oe).lastIndexOf(":"), nombre = String(e.oe).slice(0, k), ord = +String(e.oe).slice(k + 1);
        const re = new RegExp("<(?:\\w+:)?" + reEsc(nombre) + "(?=[\\s/>])", "g");
        re.lastIndex = u.i;
        let m, n = -1;
        while ((m = re.exec(xd)) && m.index < u.f) { if (++n === ord) break; }
        if (!m || m.index >= u.f || n !== ord) return null;
        const el = O.hijosXml(xd, m.index, u.f)[0];
        if (!el) return null;
        if (nombre === "oMath" || nombre === "oMathPara") return xd.slice(el.i, el.f);
        let p = m.index;
        for (;;) { p = xd.lastIndexOf("<w:r", p - 1); if (p < u.i) return null; const c = xd.charCodeAt(p + 4); if (c === 62 || c === 32) break; }
        const run = O.hijosXml(xd, p, u.f)[0];
        if (!run || run.f < el.f) return null;
        // si la corrida trae además texto, ese texto ya va por su lado: no se repite
        if (/<w:t[ >]/.test(xd.slice(run.ci, el.i) + xd.slice(el.f, run.cf))) return null;
        return xd.slice(run.i, run.f);
      };
      const numDe = (k) => { const u = elDe(k); if (!u || O.local(u.n) !== "p") return null; const pp = O.hijoXml(xd, "pPr", u.ci, u.cf); const m = pp && /<w:numId w:val="(\d+)"\/>/.exec(xd.slice(pp.i, pp.f)); return m && m[1] !== "0" ? m[1] : null; };
      /** Para un renglón de lista nuevo: la numeración de su vecino (el anterior de la misma lista, o el siguiente). */
      const numVecino = (q) => {
        const tipo = ps[q].pa.lista;
        if (tipo !== "vinetas" && tipo !== "numeros") return null;
        for (let t = q - 1; t >= 0 && ps[t].pa.lista === tipo; t--) if (ps[t].pa.ox != null && hu[ps[t].pa.ox]) { const n = numDe(ps[t].pa.ox); if (n) return n; }
        for (let t = q + 1; t < ps.length && ps[t].pa.lista === tipo; t++) if (ps[t].pa.ox != null && hu[ps[t].pa.ox]) { const n = numDe(ps[t].pa.ox); if (n) return n; }
        return null;
      };
      // ---- MEJ-1: un párrafo editado se PARCHA sobre su XML original
      /**
       * Antes, el párrafo que cambiaba se escribía otra vez desde el modelo y se perdía lo que el modelo
       * no guarda (w:rFonts w:cs/w:eastAsia, w:lang, w:noProof, rsid, marcadores, comentarios, campos…).
       * Ahora se alinea el texto nuevo con el del párrafo original y cada corrida conserva su XML: solo
       * cambia el texto que la persona tocó y las propiedades de letra que el modelo sí maneja y que ya
       * no coinciden. Ante cualquier duda regresa null y el párrafo se escribe como antes.
       */
      const fuentesTema = (() => {
        let f = null;
        return () => {
          if (f) return f;
          f = {};
          const tr = parte(/\/theme$/), tt = tr ? pq.texto(tr) : null;
          if (tt) ["major", "minor"].forEach((k) => { const m = new RegExp("<a:" + k + "Font>\\s*<a:latin\\b[^>]*\\btypeface=\"([^\"]*)\"").exec(tt); if (m && m[1]) f[k] = O.desescapar(m[1]); });
          return f;
        };
      })();
      const ORDEN_RPR = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"];
      const GRUPOS_RPR = { fuente: ["rFonts"], b: ["b", "bCs"], i: ["i", "iCs"], mayus: ["caps"], vers: ["smallCaps"], s: ["strike", "dstrike"], oc: ["vanish"], color: ["color"], esp: ["spacing"], pos: ["position"], tam: ["sz", "szCs"], u: ["u"], fondo: ["highlight", "shd"], va: ["vertAlign"] };
      const cacheRpr = new Map(), baseDoc = doc.valor("_doc", "formato") || {}, clavesA = new WeakMap();
      /** Dos letras tienen el mismo formato si sus atributos dicen lo mismo (aunque sean objetos distintos). */
      const claveA = (a) => { if (!a) return ""; let c = clavesA.get(a); if (c == null) { c = O.estable(a); clavesA.set(a, c); } return c; };
      /** El w:rPr original de una corrida, con lo que pide el modelo (a). Devuelve el mismo texto si no hay nada que cambiar. */
      const rPrParchado = (rp, a) => {
        const clave = rp + "\u0001" + O.estable(a);
        if (cacheRpr.has(clave)) return cacheRpr.get(clave);
        let hijos = [];
        if (rp) { const g = O.finEtiqueta(rp, 0); if (rp.charCodeAt(g - 1) !== 47) hijos = O.hijosXml(rp, g + 1, rp.lastIndexOf("</")).map((h) => ({ n: O.local(h.n), x: rp.slice(h.i, h.f), h })); }
        const el = (n) => hijos.find((h) => h.n === n) || null;
        const at = (h, k) => (h ? O.attr(rp, h.h, "w:" + k) : null);
        const activo = (h) => !!h && !["0", "false", "none"].includes(at(h, "val") || "1");
        // lo que el importador lee de ESTE w:rPr (sin lo heredado), igual que attrsRun
        const d0 = {};
        if (el("b")) d0.b = activo(el("b"));
        if (el("i")) d0.i = activo(el("i"));
        if (el("u")) d0.u = activo(el("u"));
        if (el("strike") || el("dstrike")) d0.s = activo(el("strike") || el("dstrike"));
        if (el("caps")) d0.mayus = activo(el("caps"));
        if (el("smallCaps")) d0.vers = activo(el("smallCaps"));
        if (el("vanish")) d0.oc = activo(el("vanish"));
        if (el("position")) d0.pos = Math.round(+at(el("position"), "val") * 50) / 100 || null;
        const va = at(el("vertAlign"), "val");
        if (va) d0.va = va === "superscript" ? "sup" : va === "subscript" ? "sub" : null;
        const c = at(el("color"), "val");
        if (c) d0.color = /^[0-9a-f]{6}$/i.test(c) ? "#" + c.toUpperCase() : null;
        const hl = at(el("highlight"), "val"), fill = at(el("shd"), "fill");
        if (hl) d0.fondo = RESALTE_WORD[hl] || null;
        if (fill && /^[0-9a-f]{6}$/i.test(fill)) d0.fondo = "#" + fill.toUpperCase(); else if (el("shd") && d0.fondo === undefined) d0.fondo = null;
        const sz = at(el("sz"), "val");
        if (sz) d0.tam = +sz / 2;
        if (el("spacing")) d0.esp = Math.round((+at(el("spacing"), "val") / 20) * 100) / 100 || null;
        const rf = el("rFonts");
        if (rf) { const t = at(rf, "asciiTheme") || at(rf, "hAnsiTheme"); d0.fuente = (t && fuentesTema()[/^major/.test(t) ? "major" : "minor"]) || at(rf, "ascii") || at(rf, "hAnsi") || null; }
        const igual = (k, v, w) => (k === "tam" || k === "esp" || k === "pos" ? Math.abs(+v - +w) < 0.005 : k === "color" || k === "fondo" ? String(v).toUpperCase() === String(w).toUpperCase() : v === w);
        const nuevo = {
          fuente: (v) => {
            if (!rf) return [{ n: "rFonts", x: `<w:rFonts w:ascii="${x(v)}" w:hAnsi="${x(v)}" w:cs="${x(v)}"/>` }];
            // se conserva lo que el modelo no toca (w:eastAsia, w:hint) y se cambia la fuente latina
            let ats = rf.x.slice("<w:rFonts".length, rf.x.lastIndexOf("/")).replace(/\s+w:(ascii|hAnsi|cs|asciiTheme|hAnsiTheme|cstheme)="[^"]*"/g, "").replace(/\s+$/, "");
            return [{ n: "rFonts", x: `<w:rFonts w:ascii="${x(v)}" w:hAnsi="${x(v)}" w:cs="${x(v)}"${ats}/>` }];
          },
          b: () => [{ n: "b", x: "<w:b/>" }, { n: "bCs", x: "<w:bCs/>" }],
          i: () => [{ n: "i", x: "<w:i/>" }, { n: "iCs", x: "<w:iCs/>" }],
          mayus: () => [{ n: "caps", x: "<w:caps/>" }],
          vers: () => [{ n: "smallCaps", x: "<w:smallCaps/>" }],
          s: () => [{ n: "strike", x: "<w:strike/>" }],
          oc: () => [{ n: "vanish", x: "<w:vanish/>" }],
          pos: (v) => (+v ? [{ n: "position", x: `<w:position w:val="${Math.round(+v * 2)}"/>` }] : []),
          color: (v) => (hexW(v) ? [{ n: "color", x: `<w:color w:val="${hexW(v)}"/>` }] : []),
          esp: (v) => (+v ? [{ n: "spacing", x: `<w:spacing w:val="${Math.round(+v * 20)}"/>` }] : []),
          tam: (v) => [{ n: "sz", x: `<w:sz w:val="${Math.round(v * 2)}"/>` }, { n: "szCs", x: `<w:szCs w:val="${Math.round(v * 2)}"/>` }],
          u: () => [{ n: "u", x: '<w:u w:val="single"/>' }],
          fondo: (v) => (hexW(v) ? [{ n: "shd", x: `<w:shd w:val="clear" w:color="auto" w:fill="${hexW(v)}"/>` }] : []),
          va: (v) => [{ n: "vertAlign", x: `<w:vertAlign w:val="${v === "sup" ? "superscript" : "subscript"}"/>` }],
        };
        let toco = false;
        for (const k in GRUPOS_RPR) {
          const v = a[k] == null || a[k] === false || a[k] === "" ? undefined : a[k], hay = GRUPOS_RPR[k].some((n) => el(n));
          // oculto y elevado: los documentos importados antes de que NOM los leyera no los traen en el modelo; lo que
          // diga el original se queda (NOM todavía no tiene cómo quitarlos)
          if ((k === "oc" || k === "pos") && v === undefined) continue;
          if (!hay) { if (v !== undefined) { hijos.push(...nuevo[k](v)); toco = true; } continue; }
          if (v !== undefined && d0[k] != null && d0[k] !== false && igual(k, d0[k], v)) continue;
          // sin valor en el modelo: un «apagado» explícito no se anota, ni la fuente o el tamaño que son
          // los del formato base del documento. Lo demás (negrita, color, otro tamaño…) es que la persona
          // lo quitó («Borrar formato»).
          if (v === undefined && (d0[k] == null || d0[k] === false || ((k === "fuente" || k === "tam") && baseDoc[k] != null && igual(k, d0[k], baseDoc[k])))) continue;
          hijos = hijos.filter((h) => !GRUPOS_RPR[k].includes(h.n) || (k === "fuente" && v !== undefined));
          if (v !== undefined) { if (k === "fuente") hijos = hijos.filter((h) => h.n !== "rFonts"); hijos.push(...nuevo[k](v)); }
          toco = true;
        }
        let r = rp;
        if (toco) {
          const pos = (h, i) => { const k = ORDEN_RPR.indexOf(h.n); return (k < 0 ? 1000 : k) * 10000 + i; };
          hijos = hijos.map((h, i) => [pos(h, i), h]).sort((p1, p2) => p1[0] - p2[0]).map((p1) => p1[1]);
          r = hijos.length ? `<w:rPr>${hijos.map((h) => h.x).join("")}</w:rPr>` : "";
        }
        cacheRpr.set(clave, r);
        return r;
      };
      const ENVOLTURAS = ["hyperlink", "ins", "smartTag", "fldSimple", "customXml", "sdt", "sdtContent", "bdo", "dir"];
      // (un párrafo con cambios marcados de Word que se tocó aquí se vuelve a escribir desde NOM: así aceptar o rechazar en NOM
      // quita sus marcas y lo que quede pendiente sale como cambios marcados)
      const CAMBIOS_W = ["ins", "del", "moveFrom", "moveTo"];
      const OBJ = "\uFFFC";
      /** Lo que hay dentro de un párrafo del original, en el orden y con el mismo criterio con que lo lee el importador. */
      const desarmar = (xp, ini, fin) => {
        const R = { nodos: [], runs: [], s: "", raro: false, fuera: 0 };
        const andar = (i0, f0, lista, enlace) => {
          for (const h of O.hijosXml(xp, i0, f0)) {
            const ln = O.local(h.n);
            if (ln === "pPr") continue;
            if (ln === "r") {
              const run = { tipo: "r", h, abre: xp.slice(h.i, h.ci), rpr: "", items: [], enlace };
              for (const c of O.hijosXml(xp, h.ci, h.cf)) {
                const cn = O.local(c.n), xml = xp.slice(c.i, c.f);
                if (cn === "rPr") { run.rpr = xml; continue; }
                let t = null, obj = false;
                if (cn === "t") t = c.ci === c.cf ? "" : O.desescapar(xp.slice(c.ci, c.cf));
                else if (cn === "tab") t = "\t";
                else if (cn === "br" || cn === "cr") { if (O.attr(xp, c, "w:type") === "page") R.raro = true; t = "\u2028"; }
                else if (cn === "sym") t = textoSym(O.attr(xp, c, "w:font"), O.attr(xp, c, "w:char"));
                else if (["drawing", "pict", "AlternateContent", "object", "footnoteReference", "endnoteReference"].includes(cn)) obj = true;
                if (t != null && /[<&]/.test(xp.slice(c.ci, c.cf)) && cn === "t" && /<!\[CDATA\[|<[a-zA-Z]/.test(xp.slice(c.ci, c.cf))) R.raro = true;
                if (t == null && !obj && cn !== "lastRenderedPageBreak") R.fuera++;
                const it = { xml, t: obj ? OBJ : t, cn, de: R.s.length };
                if (it.t != null) R.s += it.t;
                run.items.push(it);
              }
              R.runs.push(run); lista.push(run);
            } else if (CAMBIOS_W.includes(ln)) {
              R.raro = true;
              lista.push({ tipo: "x", xml: xp.slice(h.i, h.f) });
            } else if (ENVOLTURAS.includes(ln)) {
              let en = enlace;
              if (ln === "hyperlink") { const rid = O.attr(xp, h, "r:id"), rel = rid ? relsD.find((q) => q.id === rid) : null; if (rel && /^(https?:|mailto:)/i.test(rel.destino || "")) en = rel.destino; }
              const env = { tipo: "env", abre: xp.slice(h.i, h.ci), cierra: xp.slice(h.cf, h.f), hijos: [], vacio: h.ci === h.cf && xp.charCodeAt(h.f - 2) === 47, xml: xp.slice(h.i, h.f) };
              if (!env.vacio) andar(h.ci, h.cf, env.hijos, en);
              lista.push(env);
            } else {
              if (ln === "oMath" || ln === "oMathPara") R.raro = true;
              if (ln !== "proofErr") R.fuera++;   // marcadores, comentarios, campos…: van una sola vez
              lista.push({ tipo: "x", xml: xp.slice(h.i, h.f) });
            }
          }
        };
        andar(ini, fin, R.nodos, null);
        return R;
      };
      const textoXml = (t) => t.split(/(\t|\u2028)/).filter((z) => z !== "").map((z) => (z === "\t" ? "<w:tab/>" : z === "\u2028" ? "<w:br/>" : `<w:t xml:space="preserve">${x(z)}</w:t>`)).join("");
      const parchar = (p, u, completo, partido) => {
        if (typeof completo !== "string" || !completo.startsWith("<w:p>")) return null;
        if (p.pa.lista === "tareas" || p.pa.hecho) return null;
        const xp = xd, D = desarmar(xp, u.ci, u.cf);
        if (D.raro || (partido && D.fuera)) return null;
        // el texto nuevo, letra por letra, con el formato de cada una
        let s1 = "";
        const A = [];
        for (const t of p.tramos) {
          if (t.e) {
            // un marcador no es texto en el original (va aparte y se queda donde estaba); un campo que vino del archivo es
            // el texto de su resultado (el que está entre «separate» y «end»): así el párrafo se parcha y no se reescribe
            if (t.e.tipo === "marcador") continue;
            if (t.e.tipo === "campo" && t.e.orig) {
              const v = NOM.campos ? NOM.campos.textoVisible(t.e, { idioma: doc.valor("_doc", "idioma") }) : String(t.e.res || "");
              for (let k = 0; k < v.length; k++) A.push(t.e.a || {});
              s1 += v; continue;
            }
            if (!["imagen", "objeto", "nota"].includes(t.e.tipo) || t.e.pie) return null;
            s1 += OBJ; A.push(null); continue;
          }
          const a = t.a || {};
          if (a.sug || a.sugBorrar) return null;
          for (let k = 0; k < t.x.length; k++) A.push(a);
          s1 += t.x;
        }
        const s0 = D.s, n0 = s0.length, n1 = s1.length;
        // ---- alinear: qué letras del original se conservan (y dónde quedaron) y qué se escribió nuevo
        let pre = 0;
        while (pre < n0 && pre < n1 && s0.charCodeAt(pre) === s1.charCodeAt(pre)) pre++;
        let suf = 0;
        while (suf < n0 - pre && suf < n1 - pre && s0.charCodeAt(n0 - 1 - suf) === s1.charCodeAt(n1 - 1 - suf)) suf++;
        const queda = new Int32Array(n0).fill(-1);   // letra del original → su lugar en el texto nuevo (-1: se borró)
        for (let i = 0; i < pre; i++) queda[i] = i;
        for (let i = 0; i < suf; i++) queda[n0 - 1 - i] = n1 - 1 - i;
        const ma = n0 - pre - suf, mb = n1 - pre - suf;
        if (ma > 0 && mb > 0 && ma * mb <= 1500000) {
          // en medio hubo más de un cambio: la subsecuencia común más larga dice qué letras son las mismas
          const W1 = mb + 1, L = new Uint16Array((ma + 1) * W1);
          for (let i = ma - 1; i >= 0; i--) {
            const c = s0.charCodeAt(pre + i);
            for (let j = mb - 1; j >= 0; j--) L[i * W1 + j] = c === s1.charCodeAt(pre + j) ? L[(i + 1) * W1 + j + 1] + 1 : Math.max(L[(i + 1) * W1 + j], L[i * W1 + j + 1]);
          }
          for (let i = 0, j = 0; i < ma && j < mb;) {
            if (s0.charCodeAt(pre + i) === s1.charCodeAt(pre + j) && L[i * W1 + j] === L[(i + 1) * W1 + j + 1] + 1) { queda[pre + i] = pre + j; i++; j++; }
            else if (L[(i + 1) * W1 + j] >= L[i * W1 + j + 1]) i++; else j++;
          }
        }
        // lo escrito nuevo: tramos seguidos del texto nuevo que no vienen de ninguna letra del original
        const usada = new Uint8Array(n1);
        for (let i = 0; i < n0; i++) { if (queda[i] >= 0) usada[queda[i]] = 1; else if (s0[i] === OBJ) return null; }
        const enlaceDe = (i) => { const r = D.runs.find((q) => q.items.some((it) => it.t != null && it.t !== OBJ && i >= it.de && i < it.de + it.t.length)); return r ? r.enlace || null : null; };
        const insAntes = new Map(), insDespues = new Map();
        let alFinal = null, sig = 0;   // sig: primera letra del original que queda después de lo que se va recorriendo
        for (let j = 0; j < n1;) {
          if (usada[j]) { while (sig < n0 && queda[sig] !== j) sig++; sig++; j++; continue; }
          let m = j;
          while (m < n1 && !usada[m]) { if (s1[m] === OBJ) return null; m++; }
          // entra pegado a la letra anterior (como al teclear); si esa no puede (es un objeto, o es de otro
          // enlace), antes de la siguiente; en un párrafo sin texto, al final
          const quiere = (A[j] || {}).enlace || null, ins = { j, n: m - j };
          const ant = sig - 1, siguiente = (() => { let k = sig; while (k < n0 && queda[k] < 0) k++; return k; })();
          // (las letras borradas entre la anterior y la siguiente también sirven de sitio: su corrida sigue ahí)
          let puesto = false;
          for (let k = Math.max(0, ant); k < siguiente && !puesto; k++) if (s0[k] !== OBJ && enlaceDe(k) === quiere && !insDespues.has(k)) { insDespues.set(k, ins); puesto = true; }
          if (!puesto && siguiente < n0 && s0[siguiente] !== OBJ && enlaceDe(siguiente) === quiere && !insAntes.has(siguiente)) { insAntes.set(siguiente, ins); puesto = true; }
          if (!puesto) { if (n0 === 0 && !alFinal && !quiere) alFinal = ins; else return null; }
          j = m;
        }
        let ok = true;
        const armarRun = (run) => {
          // piezas: { a, xml | txt } con el formato del modelo (a = null: lo que no es texto y va pegado a su vecina)
          const piezas = [];
          let igual = true;
          const pon = (txt, i1, entero, it) => {
            // parte el texto donde cambia el formato
            let k = 0;
            while (k < txt.length) {
              let m = k + 1;
              while (m < txt.length && (A[i1 + m] === A[i1 + k] || claveA(A[i1 + m]) === claveA(A[i1 + k]))) m++;
              const a = A[i1 + k] || {};
              if ((a.enlace || null) !== (run.enlace || null)) ok = false;
              piezas.push(entero && k === 0 && m === txt.length ? { a, xml: it.xml } : { a, txt: txt.slice(k, m) });
              if (!(entero && k === 0 && m === txt.length)) igual = false;
              k = m;
            }
          };
          for (const it of run.items) {
            if (it.t == null || it.t === OBJ) { piezas.push({ a: null, xml: it.xml }); continue; }
            const ini = it.de, fin = it.de + it.t.length;
            let a0 = -1;   // dónde empezó el trozo conservado que se está juntando
            const cierra = (hasta) => { if (a0 >= 0) { pon(it.t.slice(a0 - ini, hasta - ini), queda[a0], a0 === ini && hasta === fin, it); a0 = -1; } };
            const escribe = (ins) => { pon(s1.slice(ins.j, ins.j + ins.n), ins.j, false, it); };
            for (let i = ini; i < fin; i++) {
              if (insAntes.has(i)) { cierra(i); escribe(insAntes.get(i)); }
              if (queda[i] >= 0) { if (a0 >= 0 && queda[i] !== queda[i - 1] + 1) cierra(i); if (a0 < 0) a0 = i; }
              else { cierra(i); igual = false; }
              if (insDespues.has(i)) { cierra(i + 1); escribe(insDespues.get(i)); }
            }
            cierra(fin);
          }
          const conTexto = piezas.filter((q) => q.a);
          if (!conTexto.length) {
            // sin letras: queda solo lo que no es texto (campos, saltos, objetos); si tampoco hay, la corrida desaparece
            const resto = piezas.map((q) => q.xml || "").join("");
            return resto ? (igual ? xp.slice(run.h.i, run.h.f) : run.abre + run.rpr + resto + "</w:r>") : "";
          }
          // corridas nuevas: una por cada formato distinto seguido
          const grupos = [];
          let pend = [];
          // dos trozos de texto seguidos (lo de antes y lo de después de algo borrado) quedan en un solo w:t
          const junta = (g, q) => { const u2 = g.partes[g.partes.length - 1]; if (q.txt != null && u2 && u2.txt != null) u2.txt += q.txt; else g.partes.push(q.txt != null ? { txt: q.txt } : { xml: q.xml }); };
          for (const q of piezas) {
            if (!q.a) { if (grupos.length) grupos[grupos.length - 1].partes.push({ xml: q.xml }); else pend.push({ xml: q.xml }); continue; }
            let g = grupos[grupos.length - 1];
            if (!g || !(g.a === q.a || claveA(g.a) === claveA(q.a))) { g = { a: q.a, partes: pend }; grupos.push(g); pend = []; }
            junta(g, q);
          }
          const salida = grupos.map((g) => ({ rpr: rPrParchado(run.rpr, g.a), xml: g.partes.map((q) => (q.xml != null ? q.xml : textoXml(q.txt))).join("") }));
          if (igual && salida.length === 1 && salida[0].rpr === run.rpr) return xp.slice(run.h.i, run.h.f);
          return salida.map((g) => run.abre + g.rpr + g.xml + "</w:r>").join("");
        };
        const armar = (lista) => lista.map((nd) => {
          if (nd.tipo === "x") return nd.xml;
          if (nd.tipo === "r") return armarRun(nd);
          if (nd.vacio) return nd.xml;
          const dentro = armar(nd.hijos);
          // un enlace (u otra envoltura) al que se le borró todo el texto no se deja vacío
          return dentro || !nd.hijos.some((h) => h.tipo === "r") ? nd.abre + dentro + nd.cierra : "";
        }).join("");
        let cuerpoP = armar(D.nodos);
        if (!ok) return null;
        if (alFinal) {
          // el párrafo no tenía texto: lo escrito se agrega al final, como corridas nuevas
          let k = alFinal.j;
          const tope = alFinal.j + alFinal.n;
          while (k < tope) { let m = k + 1; while (m < tope && claveA(A[m]) === claveA(A[k])) m++; if ((A[k] || {}).enlace) return null; cuerpoP += texto(s1.slice(k, m), A[k] || {}); k = m; }
        }
        // comprobación: lo armado tiene que decir exactamente lo que dice el modelo
        const pr = O.hijosXml(completo, 5)[0], ppr = pr && O.local(pr.n) === "pPr" ? completo.slice(pr.i, pr.f) : "";
        const abre = partido ? "<w:p>" : xp.slice(u.i, u.ci);
        const fin = "<w:p>" + cuerpoP + "</w:p>", V = desarmar(fin, 5, fin.length - 6);
        if (V.s !== s1) return null;
        return abre + conEstilos(ppr) + cuerpoP + "</w:p>";
      };
      let parchados = 0;
      /**
       * Una tabla de Word que se editó sale de su XML original con solo lo que cambió: las celdas cuya firma no cambió salen
       * tal cual (imágenes, listas, tablas adentro, estilo); las que cambiaron conservan su formato de celda, de párrafo y
       * de letra con el texto nuevo; una fila nueva copia la forma de la de arriba. null si cambió la forma de la tabla
       * (columnas o celdas combinadas): entonces se escribe de nuevo.
       */
      const parcharTabla = (u, b) => {
        const hs = O.hijosXml(xd, u.ci, u.cf), trs = hs.filter((h) => O.local(h.n) === "tr");
        if (!trs.length || hs.some((h) => !["tblPr", "tblGrid", "tr", "bookmarkStart", "bookmarkEnd"].includes(O.local(h.n)))) return null;
        const filasO = [];
        for (const tr of trs) {
          const celdas = [];
          let k = 0;
          for (const h of O.hijosXml(xd, tr.ci, tr.cf)) {
            const ln = O.local(h.n);
            if (ln === "trPr" || ln === "tblPrEx") continue;
            if (ln !== "tc") return null;
            const tcPr = O.hijoXml(xd, "tcPr", h.ci, h.cf), gs = tcPr && O.hijoXml(xd, "gridSpan", tcPr.ci, tcPr.cf), vm = tcPr && O.hijoXml(xd, "vMerge", tcPr.ci, tcPr.cf);
            const span = Math.max(1, +(gs ? O.attr(xd, gs, "w:val") : 1) || 1);
            celdas.push({ k, span, tc: h, tcPr, vm: vm ? O.attr(xd, vm, "w:val") || "continue" : null });
            k += span;
          }
          filasO.push({ tr, celdas });
        }
        const nC = Math.min(63, Math.max(...filasO.map((f) => f.celdas.reduce((s, c) => s + c.span, 0))));
        if ((+b.colsTexto || 0) !== nC) return null;
        const nF = Math.max(1, +b.filasTexto || 0), deOrig = [];
        for (let f = 0; f < nF; f++) {
          let o = -1;
          for (let c = 0; c < nC && o < 0; c++) { const s = String(b["o:" + F().ref(c, f)] || ""); if (/^\d+\|/.test(s)) o = +s.split("|")[0]; }
          deOrig.push(o >= 0 && o < filasO.length ? o : -1);
        }
        if (deOrig.some((o, f) => o >= 0 && deOrig.indexOf(o) !== f)) return null;   // una fila del original repetida
        // las celdas combinadas tienen que ser las mismas que en el original (en las filas que vienen de él)
        const combO = {}, abiertas = {};
        filasO.forEach((fo, f) => fo.celdas.forEach((ce) => {
          if (ce.vm === "continue" && abiertas[ce.k]) { const a = abiertas[ce.k]; a.h++; combO[a.r] = `${a.w}x${a.h}`; return; }
          abiertas[ce.k] = ce.vm === "restart" ? { r: ce.k + "," + f, w: ce.span, h: 1 } : null;
          if (ce.span > 1 || ce.vm === "restart") combO[ce.k + "," + f] = `${ce.span}x1`;
        }));
        Object.keys(combO).forEach((x) => { if (combO[x] === "1x1") delete combO[x]; });
        const combA = {};
        for (const k2 in b) {
          if (!k2.startsWith("m:") || !/^\d+x\d+$/.test(String(b[k2] || "")) || b[k2] === "1x1") continue;
          const p = F().parseRef(k2.slice(2));
          if (!p || p.f >= nF) continue;
          if (deOrig[p.f] < 0) return null;
          combA[p.c + "," + deOrig[p.f]] = String(b[k2]);
        }
        if (JSON.stringify(Object.entries(combO).sort()) !== JSON.stringify(Object.entries(combA).sort())) return null;
        // una celda escrita de nuevo: su formato de celda (con el relleno de ahora) y el de su primer párrafo y su primera
        // letra, con el texto de ahora
        const ORDEN_TC = ["noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"];
        const celdaNueva = (ce, ref, nueva) => {
          let tcPr = ce.tcPr ? xd.slice(ce.tcPr.i, ce.tcPr.f) : "<w:tcPr></w:tcPr>";
          if (tcPr.endsWith("/>")) tcPr = "<w:tcPr></w:tcPr>";
          if (nueva) tcPr = tcPr.replace(/<w:vMerge\b[^>]*\/>|<w:vMerge\b[^>]*>\s*<\/w:vMerge>/, "");
          tcPr = tcPr.replace(/<w:shd\b[^>]*\/>/, "");
          const fondo = hexW(b["r:" + ref]);
          if (fondo) {
            const shd = `<w:shd w:val="clear" w:color="auto" w:fill="${fondo}"/>`, m = new RegExp("<w:(" + ORDEN_TC.join("|") + ")\\b").exec(tcPr);
            tcPr = m ? tcPr.slice(0, m.index) + shd + tcPr.slice(m.index) : tcPr.replace(/<\/w:tcPr>$/, shd + "</w:tcPr>");
          }
          if (tcPr === "<w:tcPr></w:tcPr>") tcPr = "";
          const p0 = O.hijoXml(xd, "p", ce.tc.ci, ce.tc.cf), ppr0 = p0 ? O.hijoXml(xd, "pPr", p0.ci, p0.cf) : null;
          let ppr = ppr0 && !xd.slice(ppr0.i, ppr0.f).endsWith("/>") ? xd.slice(ppr0.i, ppr0.f) : "";
          // (sin lo que es de otra cosa: marca de revisión del párrafo o el final de sección)
          ppr = ppr.replace(/<w:rPr\b[\s\S]*?<\/w:rPr>|<w:rPr\/>/, "").replace(/<w:jc\b[^>]*\/>/, "");
          const jc = ALIN_WORD[b["al:" + ref]];
          if (jc) ppr = ppr ? ppr.replace(/<\/w:pPr>$/, `<w:jc w:val="${jc}"/></w:pPr>`) : `<w:pPr><w:jc w:val="${jc}"/></w:pPr>`;
          // la letra: la de la primera letra de la celda, con la negrita, la cursiva, el color y el tamaño de ahora
          let rpr0 = "";
          if (p0) for (const r of O.hijosXml(xd, p0.ci, p0.cf)) { if (O.local(r.n) !== "r") continue; const rp = O.hijoXml(xd, "rPr", r.ci, r.cf); rpr0 = rp && rp.cf > rp.ci ? xd.slice(rp.ci, rp.cf) : ""; break; }
          rpr0 = rpr0.replace(/<w:(b|bCs|i|iCs|color|sz|szCs)\b[^>]*\/>/g, "");
          const a = { b: !!b["n:" + ref], i: !!b["i:" + ref], color: b["k:" + ref] || null, tam: +b["z:" + ref] > 0 ? +b["z:" + ref] : null };
          // (en el orden del esquema: lo de la fuente va antes que negrita y cursiva; color y tamaño, después)
          const fuente = (rpr0.match(/<w:(rStyle|rFonts)\b[^>]*\/>/g) || []).join(""), resto = rpr0.replace(/<w:(rStyle|rFonts)\b[^>]*\/>/g, "");
          const ab = (a.b ? "<w:b/><w:bCs/>" : "") + (a.i ? "<w:i/><w:iCs/>" : "");
          const ct = (hexW(a.color) ? `<w:color w:val="${hexW(a.color)}"/>` : "") + (a.tam ? `<w:sz w:val="${Math.round(a.tam * 2)}"/><w:szCs w:val="${Math.round(a.tam * 2)}"/>` : "");
          // (color y tamaño van antes que subrayado, sombreado y lo demás del final)
          const iFin = resto.search(/<w:(kern|position|sz|szCs|highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath|w14:|w15:)\b/);
          const antesCt = iFin < 0 ? resto : resto.slice(0, iFin), despCt = iFin < 0 ? "" : resto.slice(iFin);
          const rprX = fuente + ab + antesCt + ct + despCt ? `<w:rPr>${fuente}${ab}${antesCt}${ct}${despCt}</w:rPr>` : "";
          const v = String(b["c:" + ref] || "");
          const ps2 = v.split("\n").map((l) => `<w:p>${ppr}${l ? l.split(/(\t)/).filter((z) => z !== "").map((z) => `<w:r>${rprX}${z === "\t" ? "<w:tab/>" : `<w:t xml:space="preserve">${x(z)}</w:t>`}</w:r>`).join("") : ""}</w:p>`).join("");
          const tcAbre = xd.slice(ce.tc.i, O.finEtiqueta(xd, ce.tc.i) + 1);
          return (tcAbre.endsWith("/>") ? "<w:tc>" : tcAbre) + tcPr + ps2 + "</w:tc>";
        };
        const out = [xd.slice(u.i, trs[0].i)];
        let ultima = null;
        for (let f = 0; f < nF; f++) {
          const o = deOrig[f], fo = o >= 0 ? filasO[o] : null, base = fo || ultima;
          if (!base) return null;
          const partes = [];
          let pos = base.tr.i;
          for (const ce of base.celdas) {
            if (ce.k >= nC) return null;
            const ref = F().ref(ce.k, f), firma = String(b["o:" + ref] || "");
            partes.push(xd.slice(pos, ce.tc.i));
            const igual = fo && firma === o + "|" + firmaCelda(b, ref);
            if (fo && (ce.vm === "continue" || igual)) partes.push(xd.slice(ce.tc.i, ce.tc.f));
            else partes.push(celdaNueva(ce, ref, !fo));
            pos = ce.tc.f;
          }
          partes.push(xd.slice(pos, base.tr.f));
          // (una fila nueva no se repite como encabezado aunque la de arriba sí)
          let fila = partes.join("");
          if (!fo) fila = fila.replace(/<w:tblHeader\b[^>]*\/>/, "");
          out.push(fila);
          if (fo) ultima = fo;
        }
        out.push(xd.slice(trs[trs.length - 1].f, u.f));
        return out.join("");
      };
      const gen = (q, u, mismoPa, ultimo, partido) => {
        // una tabla de Word: se parcha sobre la original (solo lo que cambió)
        if (u && O.local(u.n) === "tbl" && NOM.ooxml.parcharTablas !== false) {
          const tb = ps[q].tramos.length === 1 && ps[q].tramos[0].e && ps[q].tramos[0].e.tipo === "tabla" ? ps[q].tramos[0].e : null;
          if (tb) {
            let t = null;
            try { t = parcharTabla(u, tb); } catch (e) { if (typeof console !== "undefined") console.warn("Word: una tabla no se pudo parchar; se escribe de nuevo.", e); }
            if (t != null) { parchados++; return t + (q === ps.length - 1 ? "<w:p/>" : ""); }
          }
        }
        const U = { ppr: null, sect: "", objeto: u ? objetoDe(u) : null, numId: null };

        if (u && O.local(u.n) === "p") {
          const pp = O.hijoXml(xd, "pPr", u.ci, u.cf);
          if (mismoPa) { U.ppr = pp ? xd.slice(pp.i, pp.f) : ""; if (!ultimo) U.ppr = U.ppr.replace(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[\s\S]*<\/w:sectPr>/, ""); }
          else {
            const sect = pp && ultimo ? O.hijoXml(xd, "sectPr", pp.ci, pp.cf) : null;
            if (sect) U.sect = xd.slice(sect.i, sect.f);
            U.numId = numVecino(q);
          }
        } else U.numId = numVecino(q);
        const completo = xmlParrafo(ps[q], q, U);
        if (u && O.local(u.n) === "p" && u.n === "w:p" && u.ci !== u.cf && NOM.ooxml.parcharParrafos !== false) {
          let t = null;
          try { t = parchar(ps[q], u, completo, !!partido); } catch (e) { if (typeof console !== "undefined") console.warn("Word: un párrafo no se pudo parchar; se escribe de nuevo.", e); }
          if (t != null) { parchados++; return t; }
        }
        return conEstilos(completo);
      };
      // ---- recorrer el documento actual y decidir unidad por unidad
      const piezas = [], usados = new Set(), cacheEl = new Map();
      let suelta = 0, textuales = 0, regeneradas = 0;
      const topDe = (ox) => +String(ox).split(".")[0];
      /** La unidad del original a la que apunta un ox ("7" = hijo 7 del cuerpo; "7.2" = dentro de su control de contenido). */
      const elDe = (ox) => {
        const c = String(ox);
        if (cacheEl.has(c)) return cacheEl.get(c);
        const partes = c.split(".");
        let u = us[+partes[0]] || null;
        for (let d = 1; d < partes.length && u; d++) { const cont = O.hijoXml(xd, "sdtContent", u.ci, u.cf); u = cont ? O.hijosXml(xd, cont.ci, cont.cf)[+partes[d]] || null : null; }
        cacheEl.set(c, u);
        return u;
      };
      const modeladas = new Set(), hojas = {};
      Object.keys(hu).forEach((c) => { modeladas.add(topDe(c)); if (c.includes(".")) (hojas[topDe(c)] = hojas[topDe(c)] || []).push(c); });
      const porRuta = (a, b) => { const x = a.split(".").map(Number), y = b.split(".").map(Number); for (let d = 0; d < Math.max(x.length, y.length); d++) { if ((x[d] || 0) !== (y[d] || 0)) return (x[d] || 0) - (y[d] || 0); } return 0; };
      Object.values(hojas).forEach((l) => l.sort(porRuta));
      // El espacio en blanco que el original tenía antes de una unidad que sale tal cual (sangría del XML):
      // así un documento que no cambió sale idéntico byte por byte aunque se arme pieza por pieza.
      const antesDe = (el) => { const idx = us.indexOf(el); if (idx < 0) return ""; const g = xd.slice(idx === 0 ? body.ci : us[idx - 1].f, el.i); return /^\s*$/.test(g) ? g : ""; };
      const sueltas = (hasta) => { for (; suelta < hasta; suelta++) { const u = us[suelta]; if (!modeladas.has(suelta) && O.local(u.n) !== "sectPr") piezas.push(antesDe(u) + xd.slice(u.i, u.f)); } };
      /** Escribe en `sale` la unidad que empieza en el párrafo i (uno nuevo, o todos los que salieron de la misma unidad del original). Devuelve dónde sigue. */
      const emitirUnidad = (i, sale) => {
        const k = ps[i].pa.ox, u = k != null && hu[k] != null && !usados.has(String(k)) ? elDe(k) : null;
        if (!u) { sale.push(gen(i, null, false, true)); regeneradas++; return i + 1; }
        let j = i;
        while (j + 1 < ps.length && ps[j + 1].pa.ox === k) j++;
        usados.add(String(k));
        const grupo = ps.slice(i, j + 1);
        if (mismoGrupo(grupo, hu[k][0])) { sale.push(antesDe(u) + xd.slice(u.i, u.f)); textuales++; }
        else {
          // La unidad cambió (se editó, o se partió con Enter). Párrafo por párrafo: el que sigue igual
          // que el original sale textual; los demás se parchan sobre su XML original (o se escriben de nuevo).
          let textual = false;
          for (let q = i; q <= j; q++) {
            const p = ps[q], unico = p.tramos.length === 1 && p.tramos[0].e;
            if (!textual && grupo.length > 1 && mismoGrupo([p], hu[k][0])) { sale.push(xd.slice(u.i, u.f)); textual = true; textuales++; continue; }
            const ant = sale[sale.length - 1];
            if (unico && unico.tipo === "salto" && q > i && ant && ant.endsWith("</w:p>") && !textual) { sale[sale.length - 1] = ant.slice(0, -6) + '<w:r><w:br w:type="page"/></w:r></w:p>'; continue; }
            if (unico && unico.tipo === "saltocol" && q > i && ant && ant.endsWith("</w:p>") && !textual) { sale[sale.length - 1] = ant.slice(0, -6) + '<w:r><w:br w:type="column"/></w:r></w:p>'; continue; }
            sale.push(gen(q, u, H(sinOx(p.pa)) === hu[k][1], q === j, grupo.length > 1)); regeneradas++;
          }
        }
        return j + 1;
      };
      const viva = (q) => { const ox = q < ps.length ? ps[q].pa.ox : null; return ox != null && hu[ox] != null && !usados.has(String(ox)) ? String(ox) : null; };
      /**
       * Un control de contenido (w:sdt: título de portada, campos de una plantilla…) al que se le editó
       * algo por dentro conserva su envoltura: antes sus párrafos salían sueltos y el control se perdía.
       * Se le quita el enlace a datos (w:dataBinding) y la marca de «texto de ejemplo» para que Word
       * muestre lo que la persona escribió y no lo vuelva a cambiar por el valor viejo.
       */
      const conModelo = new Set();   // cada clave del modelo y cada control que la contiene ("0.5.0" → "0", "0.5", "0.5.0")
      Object.keys(hu).forEach((c) => { const pt = c.split("."); for (let d = 1; d <= pt.length; d++) conModelo.add(pt.slice(0, d).join(".")); });
      const contenidoDe = (el) => (el && O.local(el.n) === "sdt" ? O.hijoXml(xd, "sdtContent", el.ci, el.cf) : null);
      /** ¿Todas las claves de este control llevan a algo que existe, pasando solo por controles de contenido? */
      const armable = (top) => hojas[top].every((c) => {
        const pt = c.split(".");
        let el = us[top];
        for (let d = 1; d < pt.length; d++) { const cont = contenidoDe(el); el = cont ? O.hijosXml(xd, cont.ci, cont.cf)[+pt[d]] : null; if (!el) return false; }
        return true;
      });
      const envolverNodo = (pref, el, q0) => {
        const cont = contenidoDe(el), hijosC = O.hijosXml(xd, cont.ci, cont.cf), dentro = [];
        let q = q0, sueltaC = 0, ult = -1, cambio = false;
        const reg0 = regeneradas;
        const sueltasC = (hasta) => { for (; sueltaC < hasta && sueltaC < hijosC.length; sueltaC++) { if (!conModelo.has(pref + "." + sueltaC)) dentro.push(xd.slice(hijosC[sueltaC].i, hijosC[sueltaC].f)); else cambio = true; } };
        while (q < ps.length) {
          const s = viva(q);
          if (s == null) {
            // un párrafo nuevo queda dentro solo si después sigue otro del mismo control
            let r = q + 1;
            while (r < ps.length && viva(r) == null) r++;
            const s2 = viva(r) || "";
            if (!(s2.startsWith(pref + ".") && +s2.slice(pref.length + 1).split(".")[0] > ult)) break;
            q = emitirUnidad(q, dentro);
            continue;
          }
          if (!s.startsWith(pref + ".")) break;
          const resto = s.slice(pref.length + 1), c = +resto.split(".")[0];
          if (c <= ult) break;
          sueltasC(c); sueltaC = c + 1; ult = c;
          if (!resto.includes(".")) q = emitirUnidad(q, dentro);
          else { const n = envolverNodo(pref + "." + c, hijosC[c], q); dentro.push(n.xml); q = n.sig; if (n.cambio) cambio = true; }
        }
        sueltasC(hijosC.length);
        if (regeneradas !== reg0) cambio = true;
        let cabeza = xd.slice(el.i, cont.ci);
        if (cambio) cabeza = cabeza.replace(/<w(?:15)?:dataBinding\b[^>]*\/>/g, "").replace(/<w:showingPlcHdr\b[^>]*\/>/g, "");
        return { xml: cabeza + dentro.join("") + xd.slice(cont.cf, el.f), sig: q, cambio };
      };
      const envolver = (i, top) => {
        if (!contenidoDe(us[top]) || !armable(top)) return null;
        const r = envolverNodo(String(top), us[top], i);
        hojas[top].forEach((hk) => usados.add(hk));
        return r;
      };
      for (let i = 0; i < ps.length;) {
        const k = ps[i].pa.ox, u = k != null && hu[k] != null && !usados.has(String(k)) ? elDe(k) : null;
        if (!u) { i = emitirUnidad(i, piezas); continue; }
        const top = topDe(k);
        sueltas(top);
        // Un control de contenido (portada, índice…) con todo su contenido intacto sale entero, con su envoltura.
        if (String(k).includes(".") && hojas[top] && hojas[top][0] === String(k)) {
          let q = i, entero = true;
          for (const hk of hojas[top]) {
            if (q >= ps.length || String(ps[q].pa.ox) !== hk) { entero = false; break; }
            let j2 = q;
            while (j2 + 1 < ps.length && String(ps[j2 + 1].pa.ox) === hk) j2++;
            if (!mismoGrupo(ps.slice(q, j2 + 1), hu[hk][0])) { entero = false; break; }
            q = j2 + 1;
          }
          if (entero) { piezas.push(antesDe(us[top]) + xd.slice(us[top].i, us[top].f)); hojas[top].forEach((hk) => usados.add(hk)); textuales++; i = q; continue; }
        }
        if (String(k).includes(".") && hojas[top]) {
          const env = envolver(i, top);
          if (env) { piezas.push(antesDe(us[top]) + env.xml); i = env.sig; continue; }
        }
        i = emitirUnidad(i, piezas);
      }
      sueltas(us.length);
      const cuerpo = piezas.join("").replace(/\u0001SEC[AB]\d*\u0001/g, "");

      function ensamblar(E) {
        const cambio = (k) => { const v = doc.valor("_doc", k), h0 = og.h.d ? og.h.d[k] : undefined; return v == null ? h0 != null : h0 !== H(v); };
        // ---- sección final: tamaño, márgenes y encabezados solo si la persona los cambió
        const su = us.find((h) => O.local(h.n) === "sectPr");
        let sect = su ? xd.slice(su.i, su.f) : "<w:sectPr/>";
        if (cambio("pagina")) {
          const mar = `<w:pgMar w:top="${E.mv(E.m.s)}" w:right="${E.mv(E.m.d)}" w:bottom="${E.mv(E.m.i)}" w:left="${E.mv(E.m.iz)}" w:header="720" w:footer="720" w:gutter="0"/>`;
          sect = sect.replace(/<w:pgSz\b[^>]*\/>/, "").replace(/<w:pgMar\b[^>]*\/>/, "");
          sect = /<w:sectPr\b[^>]*\/>$/.test(sect) ? sect.replace(/\/>$/, ">") + E.pgSz + mar + "</w:sectPr>" : sect.replace(/(<w:sectPr\b[^>]*>(?:\s*<w:(?:headerReference|footerReference|footnotePr|endnotePr|type)\b[^>]*\/>)*)/, "$1" + E.pgSz + mar);
        }
        // números de línea puestos desde NOM: valen para todo el documento (cada sección del archivo)
        let cuerpoN = cuerpo;
        if (cambio("numLineas")) {
          const v = (doc.valor("_doc", "numLineas") || {}).v;
          sect = conNumLineas(sect, v);
          cuerpoN = cuerpo.replace(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, (m) => conNumLineas(m, v));
        }
        const CAMPOS_EP = ["encabezado", "pie", "numeros", "numPos", "numFmt", "numInicio", "primeraDistinta", "encAlin", "pieAlin", "encabezadoImg", "marcaAgua"];
        const epCambio = CAMPOS_EP.some(cambio);
        const sufijo = PRE;
        if (epCambio) {
          const refs = E.refs.replace(/r:id="(rId\w+)"/g, `r:id="$1${sufijo}"`);
          sect = sect.replace(/<w:(headerReference|footerReference)\b[^>]*\/>/g, "").replace(/<w:titlePg\b[^>]*\/>/, "");
          sect = /<w:sectPr\b[^>]*\/>$/.test(sect) ? sect.replace(/\/>$/, ">") + refs + (E.primera ? "<w:titlePg/>" : "") + "</w:sectPr>" : sect.replace(/(<w:sectPr\b[^>]*>)/, "$1" + refs).replace(/<\/w:sectPr>$/, (E.primera ? "<w:titlePg/>" : "") + "</w:sectPr>");
          E.extras.forEach((e) => {
            const nombre = e.nombre.replace(/^word\//, carpeta).replace(/(header|footer)(\d)\.xml/, "nom" + g + "$1$2.xml");
            pq.pon(nombre, e.datos);
            if (!/_rels\//.test(nombre)) pq.tipo(nombre, "application/vnd.openxmlformats-officedocument.wordprocessingml." + (/header/.test(nombre) ? "header" : "footer") + "+xml");
          });
          rels.filter((r) => r.tipo === T.encabezado || r.tipo === T.pie).forEach((r) => pq.agregaRel(DOC, r.id + sufijo, r.tipo, r.destino.replace(/(header|footer)(\d)\.xml/, "nom" + g + "$1$2.xml")));
        }
        // ---- raíz: que estén declarados los prefijos que usa lo nuevo
        let abre = xd.slice(0, body.ci);
        const NS = { r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships", wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", m: "http://schemas.openxmlformats.org/officeDocument/2006/math" };
        const etiqueta = xd.slice(raiz.i, raiz.ci);
        let extra = "";
        Object.keys(NS).forEach((p) => { if (!new RegExp("\\sxmlns:" + p + "\\s*=").test(etiqueta)) extra += ` xmlns:${p}="${NS[p]}"`; });
        if (extra) abre = xd.slice(0, raiz.ci - 1) + extra + ">" + xd.slice(raiz.ci, body.ci);
        // Lo de alrededor de la sección final (saltos de línea del original) también se conserva.
        const antesSect = su ? antesDe(su) : "", trasSect = su && /^\s*$/.test(xd.slice(su.f, body.cf)) ? xd.slice(su.f, body.cf) : "";
        pq.pon(DOC, abre + cuerpoN + antesSect + sect + trasSect + xd.slice(body.cf));
        // ---- relaciones e imágenes nuevas
        const usadosEnCuerpo = (id) => cuerpo.includes('"' + id + '"');
        rels.filter((r) => (r.tipo === T.img || r.tipo === T.enlace) && usadosEnCuerpo(r.id)).forEach((r) => pq.agregaRel(DOC, r.id, r.tipo, r.destino, r.externo));
        medios.filter((mm) => epCambio || !/media\/(logo|marca)\./.test(mm.nombre)).forEach((mm) => {
          pq.pon(mm.nombre.replace(/^word\//, carpeta), mm.datos);
          pq.extension(mm.nombre.split(".").pop(), MIME_IMG[mm.nombre.split(".").pop()] || "image/png");
        });
        // ---- numeración nueva (listas hechas en NOM)
        if (instancias.length) {
          const abs = (E.numeracion.match(/<w:abstractNum w:abstractNumId="\d+">[\s\S]*?<\/w:abstractNum>/g) || []).filter((t) => +/abstractNumId="(\d+)"/.exec(t)[1] >= baseAbs).join("");
          const nums = (E.numeracion.match(/<w:num w:numId="\d+">[\s\S]*?<\/w:num>/g) || []).filter((t) => +/numId="(\d+)"/.exec(t)[1] >= baseNum).join("");
          if (numTxt) {
            const r0 = O.hijosXml(numTxt).find((h) => O.local(h.n) === "numbering"), hs = O.hijosXml(numTxt, r0.ci, r0.cf);
            const primerNum = hs.find((h) => O.local(h.n) === "num" || O.local(h.n) === "numIdMacAtCleanup"), limpia = hs.find((h) => O.local(h.n) === "numIdMacAtCleanup");
            const p1 = primerNum ? primerNum.i : r0.cf, p2 = limpia ? limpia.i : r0.cf;
            // una raíz vacía (<w:numbering/>) se abre para poder meterle contenido
            if (r0.ci === r0.f) pq.pon(numRuta, numTxt.slice(0, r0.af) + ">" + abs + nums + "</" + r0.n + ">" + numTxt.slice(r0.f));
            else pq.pon(numRuta, aplicarPiezas(numTxt, [[p1, p1, abs], [p2, p2, nums]]));
          } else {
            pq.pon(carpeta + "numbering.xml", XML + `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abs}${nums}</w:numbering>`);
            pq.tipo(carpeta + "numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
            pq.agregaRel(DOC, "rId" + PRE + "Num", T.num, "numbering.xml");
          }
        }
        // ---- notas al pie nuevas
        if (notas.length) {
          const nuevas = (E.notasXml.match(/<w:footnote w:id="\d+">[\s\S]*?<\/w:footnote>/g) || []).filter((t) => +/w:id="(\d+)"/.exec(t)[1] > baseNota).join("");
          if (notaTxt) { const k = notaTxt.lastIndexOf("</w:footnotes>"); pq.pon(notaRuta, notaTxt.slice(0, k) + conEstilos(nuevas) + notaTxt.slice(k)); }
          else {
            pq.pon(carpeta + "footnotes.xml", conEstilos(E.notasXml));
            pq.tipo(carpeta + "footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml");
            pq.agregaRel(DOC, "rId" + PRE + "Notas", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", "footnotes.xml");
          }
        }
        // ---- estilos que se modificaron aquí (Estilos…): su definición en el original, con lo cambiado
        let estMod = estTxt;
        if (estTxt) for (const [k, e] of Object.entries(estP)) {
          if (!e || e.deWord) continue;
          const id = k.startsWith("u:") ? (e.id && idsEst.has(e.id) ? e.id : null) : ESTILO_WORD[k] && porNombre[NOMBRES[ESTILO_WORD[k]]];
          if (id != null) estMod = parcharEstiloXml(estMod, id, e);
        }
        if (estMod !== estTxt && !faltan.size) pq.pon(estRuta, estMod);
        // ---- estilos de NOM que el original no tenía
        if (faltan.size) {
          let defs = "";
          faltan.forEach((idNuevo, id) => {
            const m = new RegExp('<w:style [^>]*w:styleId="' + reEsc(id) + '">[\\s\\S]*?</w:style>').exec(E.estilos);
            if (m) defs += m[0].replace('w:styleId="' + id + '"', 'w:styleId="' + idNuevo + '"').replace(/<w:(basedOn|next) w:val="Normal"\/>/g, `<w:$1 w:val="${normal}"/>`);
          });
          if (estMod) { const k = estMod.lastIndexOf("</w:styles>"); if (k > 0) pq.pon(estRuta, estMod.slice(0, k) + defs + estMod.slice(k)); }
          else {
            pq.pon(carpeta + "styles.xml", XML + `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${defs}</w:styles>`);
            pq.tipo(carpeta + "styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
            pq.agregaRel(DOC, "rId" + PRE + "Est", T.estilos, "styles.xml");
          }
        }
        return pq.zip();
      }
      return { cuerpo, ensamblar, textuales, regeneradas, parchados };
    }
    const titulos = ps.filter((p) => /^h[1-3]$/.test(p.pa.estilo || ""));
    // Las páginas de la vista del editor abierto con este documento (para los números de las tablas de contenido e
    // ilustraciones, como los deja Word); sin editor abierto, las entradas salen sin número (Word lo pone con F9).
    const edA = NOM.editorActual, pagPara = edA && edA.doc === doc && edA.modoApi && edA.modoApi.paginaPara ? edA.modoApi.paginaPara : null;
    const figuras = [];
    { let nf = 0; ps.forEach((p) => p.tramos.forEach((t) => { if (t.e && t.e.tipo === "imagen" && t.e.pie) figuras.push({ n: ++nf, pie: t.e.pie, id: t.id }); })); }
    /** Una entrada de tabla de contenido o de ilustraciones: el texto, un tabulador con puntos al margen derecho y la página. */
    const entradaIndice = (estilo, txt, pag) => `<w:p><w:pPr><w:pStyle w:val="${estilo}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${anchoTexto}"/></w:tabs></w:pPr>${texto(txt, {})}${pag != null ? `<w:r><w:tab/></w:r>${texto(String(pag), {})}` : ""}</w:p>`;
    const campoIndice = (instr, cuerpoCampo) => `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${x(instr)} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>${cuerpoCampo}<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const xmlParrafo = (p, i, U) => {
      const pa = p.pa, unico = p.tramos.length === 1 && p.tramos[0].e;
      if (unico && unico.tipo === "tabla") return tabla(unico) + (i === ps.length - 1 ? "<w:p/>" : "");
      if (unico && unico.tipo === "salto") return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      if (unico && unico.tipo === "saltocol") return '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
      if (unico && unico.tipo === "linea") return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="8E8E8E"/></w:pBdr></w:pPr></w:p>';
      if (unico && unico.tipo === "indice" && unico.modo === "alfa") {
        // Campo INDEX de Word con las entradas ya escritas («Entrada, 2, 5», las subentradas con «Índice 2»). Con columnas,
        // en su propia sección continua, como lo deja Word (las marcas \u0001SEC…\u0001 se cambian al armar el documento).
        const cols = Math.max(1, Math.min(4, +unico.cols || 2));
        const lineas = edA && edA.doc === doc && edA.modoApi && edA.modoApi.textoApi && edA.modoApi.textoApi.entradasIndice ? edA.modoApi.textoApi.entradasIndice() : [];
        const ent = lineas.map((l) => `<w:p><w:pPr><w:pStyle w:val="Indice${l.n}"/></w:pPr>${texto(l.t, {})}</w:p>`).join("") || "<w:p/>";
        if (cols === 1) return campoIndice('INDEX \\z "3082"', ent);
        // como Word: el párrafo que abre el campo cierra la sección de antes y el último del resultado cierra la de columnas
        return `<w:p><w:pPr>\u0001SECA\u0001</w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> INDEX \\c &quot;${cols}&quot; \\z &quot;3082&quot; </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>`
          + ent + `<w:p><w:pPr>\u0001SECB${cols}\u0001</w:pPr></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
      }
      if (unico && unico.tipo === "indice") {
        // Campo TOC de Word con las entradas ya escritas, con su número de página (Word lo puede actualizar con F9).
        if (unico.modo === "ilus")
          return campoIndice('TOC \\h \\z \\c "Figura"', figuras.map((g) => entradaIndice("Tabladeilustraciones", `Figura ${g.n}. ${g.pie}`, pagPara ? pagPara(null, g.id) : null)).join("") || "<w:p/>");
        const ent = titulos.map((t) => entradaIndice("TDC" + t.pa.estilo.slice(1), t.tramos.map((z) => z.x || "").join("").trim(), pagPara ? pagPara(t.k) : null)).join("");
        return `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contenido</w:t></w:r></w:p>` + campoIndice('TOC \\o "1-3" \\h \\z \\u', ent || "<w:p/>");
      }
      let ppr = "";
      if (ESTILO_WORD[pa.estilo]) ppr += `<w:pStyle w:val="${ESTILO_WORD[pa.estilo]}"/>`;
      else if (String(pa.estilo || "").startsWith("u:") && estP[pa.estilo]) ppr += `<w:pStyle w:val="${x(idEstiloPropio(estP, pa.estilo))}"/>`;
      else if (pa.lista) ppr += '<w:pStyle w:val="ListParagraph"/>';
      // entre páginas (orden del esquema: keepNext, keepLines, pageBreakBefore, widowControl); los títulos ya van
      // con el siguiente por su estilo
      const conSig = pa.conSig == null || pa.conSig === "" ? null : !!+pa.conSig;
      if (conSig === true && !/^h[123]$/.test(pa.estilo || "")) ppr += "<w:keepNext/>";
      else if (conSig === false && /^h[123]$/.test(pa.estilo || "")) ppr += '<w:keepNext w:val="0"/>';
      if (pa.juntas) ppr += "<w:keepLines/>";
      if (pa.saltoAntes) ppr += "<w:pageBreakBefore/>";
      // letra capital: el marco de Word (dropCap) con sus renglones y su distancia al texto
      const cap = /^(drop|margin)\|(\d+)\|([\d.]+)$/.exec(pa.capital || "");
      if (cap) ppr += `<w:framePr w:dropCap="${cap[1]}" w:lines="${cap[2]}" w:hSpace="${Math.round(+cap[3] * 20)}" w:wrap="around" w:vAnchor="text" w:hAnchor="${cap[1] === "margin" ? "page" : "text"}"/>`;
      if (pa.viudas != null && pa.viudas !== "" && !+pa.viudas) ppr += '<w:widowControl w:val="0"/>';
      if (U && U.numId && (pa.lista === "vinetas" || pa.lista === "numeros")) {
        // sigue la lista del archivo original (misma numeración que sus vecinos)
        ppr += `<w:numPr><w:ilvl w:val="${pa.nivel || 0}"/><w:numId w:val="${U.numId}"/></w:numPr>`;
      } else if (U && U.ppr != null && (pa.lista === "vinetas" || pa.lista === "numeros")) {
        // sale el w:pPr del archivo original, con su numeración: no hace falta una lista nueva (antes se
        // agregaba a numbering.xml una definición que nadie usaba)
      } else if (pa.lista === "vinetas") {
        if (!numActual || numActual.tipo !== "v" + (pa.vineta || "")) numActual = { tipo: "v" + (pa.vineta || ""), id: nuevaInstancia("v" + (pa.vineta || "•")) };
        ppr += `<w:numPr><w:ilvl w:val="${pa.nivel || 0}"/><w:numId w:val="${numActual.id}"/></w:numPr>`;
      } else if (pa.lista === "numeros") {
        const clave = "n" + (pa.numfmt || "1.");
        const inicio = pa.inicio != null && pa.inicio !== "" ? +pa.inicio : null;
        if (pa.continuar && numAnterior && !numActual) numActual = numAnterior;
        if (!numActual || numActual.tipo !== clave || inicio != null) numActual = { tipo: clave, id: nuevaInstancia(clave, inicio != null ? inicio : 1) };
        ppr += `<w:numPr><w:ilvl w:val="${pa.nivel || 0}"/><w:numId w:val="${numActual.id}"/></w:numPr>`;
      } else if (pa.lista !== "tareas") { if (numActual) numAnterior = numActual; numActual = null; }
      // «Suprimir números de línea» (orden del esquema: después de la numeración, antes de los bordes)
      if (pa.sinNumLin) ppr += "<w:suppressLineNumbers/>";
      const cm = (v) => Math.round((+v || 0) * CMT);
      if (pa.lista === "tareas") ppr += `<w:ind w:left="${360 * ((pa.nivel || 0) + 1)}" w:hanging="360"/>`;
      else if (!pa.lista && (pa.sangria || pa.sangiz || pa.sangder || pa.primera != null || pa.francesa)) {
        const izq = (pa.sangria || 0) * 720 + cm(pa.sangiz) + cm(pa.francesa);
        ppr += `<w:ind w:left="${izq}"${pa.sangder ? ` w:right="${cm(pa.sangder)}"` : ""}${+pa.francesa ? ` w:hanging="${cm(pa.francesa)}"` : pa.primera != null ? ` w:firstLine="${cm(pa.primera)}"` : ""}/>`;
      } else if (pa.estilo && pa.estilo !== "cita") ppr += '<w:ind w:firstLine="0"/>';
      const linea = +pa.lineaPt ? ` w:line="${Math.round(+pa.lineaPt * 20)}" w:lineRule="exact"` : pa.interlineado ? ` w:line="${Math.round(pa.interlineado * 240)}" w:lineRule="auto"` : "";
      // (orden del esquema: bordes, fondo, sentido, espacio, sangría, alineación)
      const ind = /<w:ind [^>]*\/>$/.exec(ppr);
      if (ind) ppr = ppr.slice(0, ind.index);
      const bm = /^(\d+(?:\.\d+)?) (solid|double|dotted|dashed) (#[0-9A-Fa-f]{6}) ([sizd]{1,4})$/.exec(pa.borde || "");
      if (bm) {
        const valW = { solid: "single", double: "double", dotted: "dotted", dashed: "dashed" }[bm[2]], sz = Math.max(2, Math.round((bm[2] === "double" ? +bm[1] / 3 : +bm[1]) * 8));
        ppr += `<w:pBdr>${[["s", "top"], ["z", "left"], ["i", "bottom"], ["d", "right"]].filter(([k]) => bm[4].includes(k)).map(([, l]) => `<w:${l} w:val="${valW}" w:sz="${sz}" w:space="4" w:color="${bm[3].slice(1).toUpperCase()}"/>`).join("")}</w:pBdr>`;
      }
      if (hexW(pa.sombra)) ppr += `<w:shd w:val="clear" w:color="auto" w:fill="${hexW(pa.sombra)}"/>`;
      // las tabulaciones del párrafo (las que tenía en Word: si no, al escribir el párrafo de nuevo se perdían) y «no
      // dividir con guiones» (orden del esquema: fondo, tabulaciones, guiones, sentido)
      if (Array.isArray(pa.tabs) && pa.tabs.length) {
        const TIPO_W = { i: "left", c: "center", d: "right", p: "decimal", b: "bar" }, REL = new Set(["dot", "hyphen", "underscore", "heavy", "middleDot"]);
        const tt = pa.tabs.filter((t) => Array.isArray(t) && isFinite(+t[0]) && +t[0] >= 0 && +t[0] <= 1584).map((t) => `<w:tab w:val="${TIPO_W[t[1]] || "left"}"${REL.has(t[2]) ? ` w:leader="${t[2]}"` : ""} w:pos="${Math.round(+t[0] * 20)}"/>`);
        if (tt.length) ppr += `<w:tabs>${tt.join("")}</w:tabs>`;
      }
      if (pa.sinGuiones) ppr += "<w:suppressAutoHyphens/>";
      if (pa.dir === "rtl") ppr += "<w:bidi/>";
      if (pa.antes != null || pa.despues != null || linea) ppr += `<w:spacing${pa.antes != null ? ` w:before="${Math.round(pa.antes * 20)}"` : ""}${pa.despues != null ? ` w:after="${Math.round(pa.despues * 20)}"` : ""}${linea}/>`;
      if (ind) ppr += ind[0];
      if (ALIN_WORD[pa.alin]) ppr += `<w:jc w:val="${ALIN_WORD[pa.alin]}"/>`;
      let runs = pa.lista === "tareas" ? texto(pa.hecho ? "☑\t" : "☐\t", {}) : "";
      let despues = "";
      // las sugerencias salen como cambios marcados de Word: lo propuesto para agregar en <w:ins> y lo propuesto para borrar en
      // <w:del> (con w:delText), con su autor y su fecha; seguidas del mismo autor, en uno solo
      let rev = null;
      const cambio = (a) => {
        const sg = a && a.sug ? ["ins", a.sug] : a && a.sugBorrar ? ["del", a.sugBorrar] : null;
        const k = sg ? sg[0] + "|" + (sg[1].n || "") + "|" + (sg[1].t || "") : "";
        if ((rev ? rev.k : "") === k) return "";
        let out = rev ? `</w:${rev.tipo}>` : "";
        rev = sg ? { k, tipo: sg[0] } : null;
        if (sg) out += `<w:${sg[0]} w:id="${++nCambio}" w:author="${x(sg[1].n || "NOM")}" w:date="${x(sg[1].t || fechaCambio)}">`;
        return out;
      };
      const borrado = (xml) => (rev && rev.tipo === "del" ? xml.replace(/<w:t(\s[^>]*)?>/g, "<w:delText$1>").replace(/<\/w:t>/g, "</w:delText>") : xml);
      p.tramos.forEach((t) => {
        if (t.e) {
          runs += cambio(null);
          // un objeto que no cambió desde que se importó sale tal cual estaba en el archivo original
          const tal = U && U.objeto && t.e.oe ? U.objeto(t.e) : null;
          if (tal) { runs += tal; return; }
          if (t.e.tipo === "imagen") {
            runs += dibujo(t.e.src, t.e.w, t.e.alt, t.e.recorte, t.e.rot);
            if (t.e.pie) { const nb = ++nFigura, rb = "<w:rPr><w:b/></w:rPr>"; despues += `<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>${texto("Figura ", { b: true })}<w:r>${rb}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rb}<w:instrText xml:space="preserve"> SEQ Figura \\* ARABIC </w:instrText></w:r><w:r>${rb}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rb}<w:t>${nb}</w:t></w:r><w:r>${rb}<w:fldChar w:fldCharType="end"/></w:r>${texto(". ", { b: true })}${texto(t.e.pie, {})}</w:p>`; }
          } else if (t.e.tipo === "nota") {
            notas.push(t.e.texto || "");
            runs += `<w:r><w:rPr><w:rStyle w:val="RefNota"/></w:rPr><w:footnoteReference w:id="${baseNota + notas.length}"/></w:r>`;
          } else if (t.e.tipo === "ecuacion" && NOM.ecuaciones) {
            runs += t.e.bloque ? `<m:oMathPara>${NOM.ecuaciones.omml(t.e.tex || "")}</m:oMathPara>` : NOM.ecuaciones.omml(t.e.tex || "");
          } else if (t.e.tipo === "campo") {
            // campo de Word: { código | valor } como lo escribe Word (la fecha y la hora, con el valor de ahora)
            const v = NOM.campos ? NOM.campos.textoVisible(t.e, { idioma: doc.valor("_doc", "idioma") }) : String(t.e.res || "");
            const rp = rPr(t.e.a || {});
            runs += `<w:r>${rp}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rp}<w:instrText xml:space="preserve"> ${x(t.e.instr || "")} </w:instrText></w:r><w:r>${rp}<w:fldChar w:fldCharType="separate"/></w:r>${v ? texto(v, t.e.a || {}) : ""}<w:r>${rp}<w:fldChar w:fldCharType="end"/></w:r>`;
          } else if (t.e.tipo === "marcador" && t.e.nombre) {
            if (marcadoresW[t.e.nombre] == null) marcadoresW[t.e.nombre] = 900000 + Object.keys(marcadoresW).length;
            runs += t.e.lado === "fin" ? `<w:bookmarkEnd w:id="${marcadoresW[t.e.nombre]}"/>` : `<w:bookmarkStart w:id="${marcadoresW[t.e.nombre]}" w:name="${x(t.e.nombre)}"/>`;
          } else if (t.e.texto != null && t.e.tipo !== "tabla") runs += texto(String(t.e.texto), t.a || {});
          return;
        }
        const a = Object.assign({}, t.a, pa.hecho ? { s: true } : {});
        if (a.enlace && /^(https?:|mailto:)/i.test(a.enlace)) {
          const id = "rId" + PRE + "Link" + ++nLink;
          rels.push({ id, tipo: T.enlace, destino: a.enlace, externo: true });
          // (un cambio marcado va dentro del vínculo, no al revés)
          runs += cambio(null);
          const dentro = cambio(a) + borrado(texto(t.x, Object.assign({}, a)).replace(/<w:rPr>/g, '<w:rPr><w:rStyle w:val="Hyperlink"/>').replace(/<w:r>(?!<w:rPr>)/g, '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>')) + cambio(null);
          runs += `<w:hyperlink r:id="${id}" w:history="1">${dentro}</w:hyperlink>`;
        } else runs += cambio(a) + borrado(texto(t.x, a));
      });
      runs += cambio(null);
      let pprXml = U && U.ppr != null ? U.ppr : ppr || (U && U.sect) ? `<w:pPr>${ppr}${(U && U.sect) || ""}</w:pPr>` : "";
      // (la marca de párrafo de un Word con cambios marcados: si al párrafo ya no le queda ninguna sugerencia, se quedó
      // aceptado o rechazado aquí y su marca va sin el cambio)
      if (U && U.ppr && /<w:(ins|del|moveFrom|moveTo)\b/.test(pprXml) && !p.tramos.some((t) => t.a && (t.a.sug || t.a.sugBorrar)))
        pprXml = pprXml.replace(/<w:(ins|del|moveFrom|moveTo)\b[^>]*\/>/g, "");
      return `<w:p>${pprXml}${runs}</w:p>${despues}`;
    };
    const OGR = og ? cuerpoConOrigen() : null;
    // para las pruebas: cuántas unidades salieron tal cual, cuántas cambiaron y cuántas de esas se parcharon
    if (OGR) NOM.ooxml.ultimoDocx = { textuales: OGR.textuales, regeneradas: OGR.regeneradas, parchados: OGR.parchados };
    const cuerpo = OGR ? OGR.cuerpo : ps.map((p, i) => xmlParrafo(p, i, null)).join("\n");

    // Página, encabezado y pie (8.2: número donde se elija, "Página 1 de 5", primera página distinta,
    // logo y marca de agua).
    const pag = pagDoc;
    const TAM = TAM0;
    const hor = pag.orient === "h", m = pag.margen || {};
    const mv = (v) => Math.round((typeof v === "number" ? v : 96) * TWIP);
    const ep2 = {
      enc: doc.valor("_doc", "encabezado") || "", pie: doc.valor("_doc", "pie") || "", nums: doc.valor("_doc", "numeros") !== false,
      pos: doc.valor("_doc", "numPos") || "pie-der", fmt: doc.valor("_doc", "numFmt") || "n", inicio: +doc.valor("_doc", "numInicio") || 1,
      primera: !!doc.valor("_doc", "primeraDistinta"), encAl: doc.valor("_doc", "encAlin") || "izq", pieAl: doc.valor("_doc", "pieAlin") || "izq",
      logo: doc.valor("_doc", "encabezadoImg") || null, marca: doc.valor("_doc", "marcaAgua") || "",
    };
    const extras = [];
    let refs = "";
    const NSW = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    const campo = (instr, ej) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${ej}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const numero = () => ({ pagina: texto("Página ", {}) + campo("PAGE", "1"), total: campo("PAGE", "1") + texto(" de ", {}) + campo("NUMPAGES", "1"), "pagina-total": texto("Página ", {}) + campo("PAGE", "1") + texto(" de ", {}) + campo("NUMPAGES", "1") }[ep2.fmt] || campo("PAGE", "1"));
    /** Un renglón con tres lugares (izquierda, centro y derecha) separados por tabuladores. */
    const renglon = (zona, txt, al, prefijo) => {
      const slots = { izq: "", centro: "", der: "" };
      if (txt) slots[al] += texto(txt, {});
      const [zn, an] = ep2.pos.split("-");
      if (ep2.nums && zn === zona) slots[an] += (slots[an] ? '<w:r><w:t xml:space="preserve">  </w:t></w:r>' : "") + numero();
      const media = Math.round(anchoTexto / 2);
      return `<w:p><w:pPr><w:pStyle w:val="${zona === "enc" ? "Header" : "Footer"}"/><w:tabs><w:tab w:val="clear" w:pos="4680"/><w:tab w:val="clear" w:pos="9360"/><w:tab w:val="center" w:pos="${media}"/><w:tab w:val="right" w:pos="${anchoTexto}"/></w:tabs></w:pPr>${prefijo || ""}${slots.izq}<w:r><w:tab/></w:r>${slots.centro}<w:r><w:tab/></w:r>${slots.der}</w:p>`;
    };
    const relsEnc = [];
    let logoXml = "", marcaXml = "";
    if (ep2.logo && ep2.logo.src) {
      const im = imagen(ep2.logo.src);
      if (im) {
        medios.push({ nombre: `word/media/logo.${im.ext}`, datos: im.bytes });
        relsEnc.push({ id: "rIdLogo", tipo: T.img, destino: `media/logo.${im.ext}` });
        const alto = Math.min(80, +ep2.logo.alto || 36), cy = Math.round(alto * EMU), cx = Math.round(alto * (im.w / im.h) * EMU);
        logoXml = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="114300"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="900" name="Logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="900" name="logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
      }
    }
    if (ep2.marca) {
      // La marca de agua es una imagen (texto gris girado) detrás del texto, centrada en la hoja.
      const png = marcaPng(ep2.marca);
      if (png) {
        medios.push({ nombre: "word/media/marca.png", datos: png.bytes });
        relsEnc.push({ id: "rIdMarca", tipo: T.img, destino: "media/marca.png" });
        const cx = Math.round(png.w * EMU * 0.75), cy = Math.round(png.h * EMU * 0.75);
        marcaXml = `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:align>center</wp:align></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="901" name="Marca de agua"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="901" name="marca"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdMarca"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
      }
    }
    const [zonaNum] = ep2.pos.split("-");
    const hayEnc = ep2.enc || logoXml || marcaXml || (ep2.nums && zonaNum === "enc");
    const hayPie = ep2.pie || (ep2.nums && zonaNum === "pie");
    if (hayEnc) {
      rels.push({ id: "rIdEnc", tipo: T.encabezado, destino: "header1.xml" });
      extras.push({ nombre: "word/header1.xml", datos: XML + `<w:hdr ${NSW}>${renglon("enc", ep2.enc, ep2.encAl, logoXml + marcaXml)}</w:hdr>` });
      if (relsEnc.length) extras.push({ nombre: "word/_rels/header1.xml.rels", datos: RELS(relsEnc) });
      refs += '<w:headerReference w:type="default" r:id="rIdEnc"/>';
    }
    if (hayPie) {
      rels.push({ id: "rIdPie", tipo: T.pie, destino: "footer1.xml" });
      extras.push({ nombre: "word/footer1.xml", datos: XML + `<w:ftr ${NSW}>${renglon("pie", ep2.pie, ep2.pieAl)}</w:ftr>` });
      refs += '<w:footerReference w:type="default" r:id="rIdPie"/>';
    }
    if (ep2.primera && (hayEnc || hayPie)) {
      // Primera página distinta: encabezado y pie vacíos.
      rels.push({ id: "rIdEnc1", tipo: T.encabezado, destino: "header2.xml" }, { id: "rIdPie1", tipo: T.pie, destino: "footer2.xml" });
      extras.push({ nombre: "word/header2.xml", datos: XML + `<w:hdr ${NSW}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p></w:hdr>` }, { nombre: "word/footer2.xml", datos: XML + `<w:ftr ${NSW}><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p></w:ftr>` });
      refs += '<w:headerReference w:type="first" r:id="rIdEnc1"/><w:footerReference w:type="first" r:id="rIdPie1"/>';
    }
    if (notas.length) rels.push({ id: "rIdNotas", tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", destino: "footnotes.xml" });
    const notasXml = XML + `<w:footnotes ${NSW}><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${notas.map((t, i) => `<w:footnote w:id="${baseNota + i + 1}"><w:p><w:pPr><w:pStyle w:val="TextoNota"/></w:pPr><w:r><w:rPr><w:rStyle w:val="RefNota"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${x(t)}</w:t></w:r></w:p></w:footnote>`).join("")}</w:footnotes>`;
    // la hoja y los márgenes (de la sección final y de las que abre un índice alfabético con columnas)
    const pgSzX = `<w:pgSz w:w="${(hor ? TAM[1] : TAM[0]) * TWIP}" w:h="${(hor ? TAM[0] : TAM[1]) * TWIP}"${hor ? ' w:orient="landscape"' : ""}/>`;
    const pgMarX = `<w:pgMar w:top="${mv(m.s)}" w:right="${mv(m.d)}" w:bottom="${mv(m.i)}" w:left="${mv(m.iz)}" w:header="720" w:footer="720" w:gutter="0"/>`;
    const conSecIndice = /\u0001SECA\u0001/.test(cuerpo);
    const cuerpoSec = !conSecIndice ? cuerpo : cuerpo
      .replace(/\u0001SECA\u0001/g, `<w:sectPr>${refs}${pgSzX}${pgMarX}${ep2.primera && (hayEnc || hayPie) ? "<w:titlePg/>" : ""}</w:sectPr>`)
      .replace(/\u0001SECB(\d+)\u0001/g, (_, n) => `<w:sectPr><w:type w:val="continuous"/>${pgSzX}${pgMarX}<w:cols w:num="${n}" w:space="720"/></w:sectPr>`);
    const documento = XML + `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
<w:body>${cuerpoSec || "<w:p/>"}<w:sectPr>${refs}${conSecIndice ? '<w:type w:val="continuous"/>' : ""}${pgSzX}${pgMarX}${lnNumTypeXml(doc.valor("_doc", "numLineas") ? (doc.valor("_doc", "numLineas") || {}).v : (doc.valor("_doc", "pagina") || {}).numLin)}${ep2.inicio > 1 ? `<w:pgNumType w:start="${ep2.inicio}"/>` : ""}${ep2.primera && (hayEnc || hayPie) ? "<w:titlePg/>" : ""}</w:sectPr></w:body></w:document>`;
    const ep = (id, nombre, rpr, ppr, extra) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${nombre}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>${extra || ""}${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}</w:style>`;
    // 1.4 El formato base del documento (el del espacio) es el estilo Normal de Word.
    const fb = doc.valor("_doc", "formato") || {};
    const ligera = fb.fuente ? `<w:rFonts w:ascii="${x(fb.fuente)}" w:hAnsi="${x(fb.fuente)}"/>` : '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:ascii="Calibri Light" w:hAnsi="Calibri Light"/>';
    let estilos = XML + `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${x(fb.fuente || "Calibri")}" w:hAnsi="${x(fb.fuente || "Calibri")}" w:eastAsia="${x(fb.fuente || "Calibri")}" w:cs="${x(fb.fuente || "Calibri")}"/><w:sz w:val="${Math.round((+fb.tam || 11) * 2)}"/><w:szCs w:val="${Math.round((+fb.tam || 11) * 2)}"/><w:lang w:val="es-MX"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="${Math.round((+fb.interlineado || 1.15) * 240)}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>${+fb.primera ? `<w:pPr><w:ind w:firstLine="${Math.round(+fb.primera * 567)}"/></w:pPr>` : ""}</w:style>
${ep("Caption", "caption", '<w:i/><w:iCs/><w:sz w:val="18"/><w:szCs w:val="18"/>', '<w:spacing w:after="200" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>')}
${ep("TextoNota", "footnote text", '<w:sz w:val="18"/><w:szCs w:val="18"/>', '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/>')}
<w:style w:type="character" w:styleId="RefNota"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
${ep("Title", "Title", ligera + '<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>', '<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>')}
${ep("Subtitle", "Subtitle", '<w:color w:val="5A5A5A"/><w:spacing w:val="15"/>', "")}
${ep("Heading1", "heading 1", ligera + '<w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/>', '<w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/>')}
${ep("Heading2", "heading 2", ligera + '<w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/>', '<w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/>')}
${ep("Heading3", "heading 3", ligera + '<w:color w:val="1F3763"/><w:sz w:val="24"/><w:szCs w:val="24"/>', '<w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="2"/>')}
${ep("Quote", "Quote", '<w:i/><w:iCs/><w:color w:val="404040"/>', '<w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>')}
${ep("Codigo", "Código", '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/><w:sz w:val="20"/>', '<w:shd w:val="clear" w:color="auto" w:fill="F3F3F3"/><w:spacing w:after="0"/>')}
${ep("ListParagraph", "List Paragraph", "", '<w:ind w:left="720"/><w:contextualSpacing/>')}
${ep("TOCHeading", "TOC Heading", ligera + '<w:color w:val="2F5496"/><w:sz w:val="32"/>', '<w:spacing w:before="240" w:after="0"/>')}
${ep("Tabladeilustraciones", "table of figures", "", '<w:spacing w:after="0"/>')}${ep("Indice1", "index 1", "", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="220" w:hanging="220"/>')}${ep("Indice2", "index 2", "", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="440" w:hanging="220"/>')}${ep("TDC1", "toc 1", "", '<w:spacing w:after="100"/>')}${ep("TDC2", "toc 2", "", '<w:spacing w:after="100"/><w:ind w:left="220"/>')}${ep("TDC3", "toc 3", "", '<w:spacing w:after="100"/><w:ind w:left="440"/>')}
${ep("Header", "header", "", '<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')}
${ep("Footer", "footer", "", '<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')}
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TablaNom"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map((l) => `<w:${l} w:val="single" w:sz="4" w:space="0" w:color="8E8E8E"/>`).join("")}</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`;
    const nivel = (i, fmt, txt, fuente) => `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${txt}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${360 * (i + 1)}" w:hanging="360"/></w:pPr>${fuente ? `<w:rPr><w:rFonts w:ascii="${fuente}" w:hAnsi="${fuente}"/></w:rPr>` : ""}</w:lvl>`;
    const vin = ["•", "◦", "▪"], fm = ["decimal", "lowerLetter", "lowerRoman"];
    const definicion = (clave, id) => {
      if (clave[0] === "v") {
        const v0 = clave.slice(1) || "•";
        return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => nivel(i, "bullet", i === 0 ? v0 : vin[i % 3], "Calibri")).join("")}</w:abstractNum>`;
      }
      const [f0, t0] = FMT_NUM[clave.slice(1)] || FMT_NUM["1."];
      return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => (i === 0 ? nivel(0, f0, t0) : nivel(i, fm[i % 3], `%${i + 1}.`))).join("")}</w:abstractNum>`;
    };
    const numeracion = XML + `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => nivel(i, "bullet", vin[i % 3], "Calibri")).join("")}</w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => nivel(i, fm[i % 3], `%${i + 1}.`)).join("")}</w:abstractNum>
${[...abstractos.entries()].map(([k, id]) => definicion(k, id)).join("")}
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
${instancias.map((n) => `<w:num w:numId="${n.id}"><w:abstractNumId w:val="${n.abs}"/>${n.inicio != null ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${n.inicio}"/></w:lvlOverride>` : ""}</w:num>`).join("")}</w:numbering>`;
    const ajustes = XML + `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="${Math.round((+doc.valor("_doc", "tabDef") || 36) * 20)}"/>${doc.valor("_doc", "guiones") ? `<w:autoHyphenation/><w:hyphenationZone w:val="${Math.round((+(doc.valor("_doc", "guiones") || {}).zona || 18) * 20)}"/>` : ""}<w:characterSpacingControl w:val="doNotCompress"/>${notas.length ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : ""}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
    estilos = conEstilosDoc(estilos, estP);   // (los modificados y los propios: Estilos…)
    if (OGR) return OGR.ensamblar({ estilos, numeracion, notasXml, extras, refs, relsEnc, primera: ep2.primera && (hayEnc || hayPie), pgSz: `<w:pgSz w:w="${(hor ? TAM[1] : TAM[0]) * TWIP}" w:h="${(hor ? TAM[0] : TAM[1]) * TWIP}"${hor ? ' w:orient="landscape"' : ""}/>`, mv, m });
    const titulo = doc.valor("_doc", "nombre") || "Documento";
    // Sin fecha: el mismo documento da siempre exactamente el mismo archivo.
    const nucleo = XML + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${x(titulo)}</dc:title><dc:creator>NOM</dc:creator></cp:coreProperties>`;
    return zip([
      { nombre: "[Content_Types].xml", datos: CT(`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
${extras.filter((e) => /header\d\.xml$/.test(e.nombre)).map((e) => `<Override PartName="/${e.nombre}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`).join("")}
${extras.filter((e) => /footer\d\.xml$/.test(e.nombre)).map((e) => `<Override PartName="/${e.nombre}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`).join("")}
${notas.length ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' : ""}`) },
      { nombre: "_rels/.rels", datos: RELS([{ tipo: T.doc, destino: "word/document.xml" }, { tipo: T.nucleo, destino: "docProps/core.xml" }]) },
      { nombre: "docProps/core.xml", datos: nucleo },
      { nombre: "word/document.xml", datos: documento },
      { nombre: "word/styles.xml", datos: estilos },
      { nombre: "word/numbering.xml", datos: numeracion },
      { nombre: "word/settings.xml", datos: ajustes },
      { nombre: "word/_rels/document.xml.rels", datos: RELS(rels) },
      ...extras,
      ...(notas.length ? [{ nombre: "word/footnotes.xml", datos: notasXml }] : []),
      ...medios,
    ]);
  }
  /** Marca de agua como imagen PNG (texto gris claro, girado) para Word. */
  function marcaPng(txt) {
    try {
      const cv = document.createElement("canvas"), g = cv.getContext("2d");
      const tam = 110;
      g.font = `900 ${tam}px Calibri, Carlito, Arial, sans-serif`;
      const w = Math.ceil(g.measureText(txt).width) + 40, h = tam + 40;
      const diag = Math.ceil(Math.sqrt(w * w + h * h));
      cv.width = diag; cv.height = Math.ceil(diag * 0.75);
      g.translate(cv.width / 2, cv.height / 2);
      g.rotate(-Math.PI / 5);
      g.font = `900 ${tam}px Calibri, Carlito, Arial, sans-serif`;
      g.fillStyle = "rgba(0, 0, 0, 0.10)";
      g.textAlign = "center"; g.textBaseline = "middle";
      g.fillText(txt, 0, 0);
      const datos = cv.toDataURL("image/png"), bin = atob(datos.split(",")[1]), bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return { bytes, w: cv.width, h: cv.height };
    } catch (e) { return null; }
  }

  // ------------------------------------------------------------ leer .docx (importar de Word)
  /** Lector de .zip (origen.js). Se deja "async" porque los lectores de Office lo llaman con await. */
  async function abrirZip(buf) { return NOM.origen.abrirZip(buf); }
  const RESALTE_WORD = { yellow: "#FFFF00", green: "#00FF00", cyan: "#00FFFF", magenta: "#FF00FF", blue: "#0000FF", red: "#FF0000", darkBlue: "#000080", darkCyan: "#008080",
    darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000", darkGray: "#808080", lightGray: "#C0C0C0", black: "#000000" };

  /** Ecuación de Word (OMML) → la notación de NOM (\\frac, ^, _, \\sqrt…). */
  function ommlATex(n) {
    const hs = (x) => [...(x ? x.childNodes : [])].filter((c) => c.nodeType === 1);
    const h1 = (x, nombre) => hs(x).find((c) => c.localName === nombre) || null;
    const todo = (x) => hs(x).map(ommlATex).join("");
    const g = (x) => { const t = x ? todo(x) : ""; return t.length === 1 ? t : "{" + t + "}"; };
    if (!n) return "";
    switch (n.localName) {
      case "r": return [...n.getElementsByTagNameNS("*", "t")].map((t) => t.textContent).join("");
      case "f": return `\\frac{${todo(h1(n, "num"))}}{${todo(h1(n, "den"))}}`;
      case "sSup": return g(h1(n, "e")) + "^" + g(h1(n, "sup"));
      case "sSub": return g(h1(n, "e")) + "_" + g(h1(n, "sub"));
      case "sSubSup": return g(h1(n, "e")) + "_" + g(h1(n, "sub")) + "^" + g(h1(n, "sup"));
      case "rad": { const d = todo(h1(n, "deg")); return d ? `\\sqrt[${d}]{${todo(h1(n, "e"))}}` : `\\sqrt{${todo(h1(n, "e"))}}`; }
      case "nary": {
        const pr = h1(n, "naryPr"), chr = pr && h1(pr, "chr"), c = chr ? chr.getAttribute("m:val") || chr.getAttributeNS("*", "val") || "∫" : "∫";
        const sub = todo(h1(n, "sub")), sup = todo(h1(n, "sup"));
        return c + (sub ? "_{" + sub + "}" : "") + (sup ? "^{" + sup + "}" : "") + " " + todo(h1(n, "e"));
      }
      case "d": {
        const pr = h1(n, "dPr"), ab = pr && h1(pr, "begChr"), ce = pr && h1(pr, "endChr");
        const a = ab ? ab.getAttribute("m:val") : "(", b = ce ? ce.getAttribute("m:val") : ")";
        return (a == null ? "(" : a) + hs(n).filter((c) => c.localName === "e").map(todo).join(",") + (b == null ? ")" : b);
      }
      case "func": return "\\" + todo(h1(n, "fName")).trim() + " " + todo(h1(n, "e"));
      case "acc": return `\\vec{${todo(h1(n, "e"))}}`;
      case "limLow": return todo(h1(n, "e")) + "_{" + todo(h1(n, "lim")) + "}";
      case "naryPr": case "dPr": case "fPr": case "radPr": case "sSupPr": case "sSubPr": case "accPr": case "ctrlPr": case "rPr": return "";
      default: return todo(n);
    }
  }

  /** .docx → { parrafos: [{ pa, tramos }], pagina, aviso }. */
  async function leerDocx(archivo) {
    if (archivo.size > 30 * 1024 * 1024) throw new Error("El archivo pesa más de 30 MB");
    const u8 = new Uint8Array(await archivo.arrayBuffer());
    const z = NOM.origen.abrirZip(u8);
    let uEl = null;   // la unidad del cuerpo (párrafo o tabla de primer nivel) que se está leyendo
    const ordinal = (c) => (uEl ? c.localName + ":" + Array.prototype.indexOf.call(uEl.getElementsByTagNameNS("*", c.localName), c) : null);
    const xmlDoc = await z.texto("word/document.xml");
    // Los marcadores que usa algún campo (REF, PAGEREF, NOTEREF): solo esos entran al texto (como marcas invisibles),
    // para que «Actualizar campos» encuentre su texto y su página. Los demás (_GoBack, los del índice…) no hace falta.
    const marcadoresUsados = new Set();
    String(xmlDoc || "").replace(/(?:REF|PAGEREF|NOTEREF)\s+([A-Za-z_][\w]*)/g, (m, n) => { marcadoresUsados.add(n); return m; });
    const nombreMarcador = {};   // id → nombre (bookmarkEnd solo trae el id)
    if (!xmlDoc) throw new Error("El archivo no parece un documento de Word (.docx)");
    const P = (t) => new DOMParser().parseFromString(t, "application/xml");
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const hijos = (n, nombre) => [...(n ? n.childNodes : [])].filter((c) => c.nodeType === 1 && (!nombre || c.localName === nombre));
    const hijo = (n, nombre) => hijos(n, nombre)[0] || null;
    const val = (n) => (n ? n.getAttributeNS(W, "val") || n.getAttribute("w:val") : null);
    const activo = (n) => !!n && !["0", "false", "none"].includes(val(n) || "1");
    // relaciones
    const relsTxt = await z.texto("word/_rels/document.xml.rels");
    const rels = {};
    if (relsTxt) [...P(relsTxt).getElementsByTagName("Relationship")].forEach((r) => (rels[r.getAttribute("Id")] = r.getAttribute("Target")));
    const at = (n, k) => (n ? n.getAttributeNS(W, k) || n.getAttribute("w:" + k) : null);
    const rutaDe = (destino) => (destino.startsWith("/") ? destino.slice(1) : "word/" + destino.replace(/^\.\//, ""));
    // estilos: id → nombre en inglés ("heading 1", "Title"…), y de cada uno su estilo base y sus
    // propiedades, para saber con qué letra y qué espaciado se ve de verdad cada párrafo (en un
    // documento real casi nada trae formato propio: la fuente y el interlineado vienen del estilo)
    const estilos = {}, defEstilo = {}, estilosW = {};   // (estilosW: los estilos propios del Word que usan sus párrafos)
    let rprDef = null, pprDef = null, idNormal = null, idTablaNormal = null;
    const stTxt = await z.texto("word/styles.xml");
    if (stTxt) {
      const ds = P(stTxt), dd = ds.getElementsByTagNameNS(W, "docDefaults")[0];
      if (dd) { rprDef = hijo(hijo(dd, "rPrDefault"), "rPr"); pprDef = hijo(hijo(dd, "pPrDefault"), "pPr"); }
      [...ds.getElementsByTagNameNS(W, "style")].forEach((s) => {
        const nm = (val(hijo(s, "name")) || "").toLowerCase(), id = at(s, "styleId"), tipo = at(s, "type");
        estilos[id] = nm;
        defEstilo[id] = { nombre: nm, tipo, base: val(hijo(s, "basedOn")), ppr: hijo(s, "pPr"), rpr: hijo(s, "rPr"), nodo: tipo === "table" ? s : undefined,
          propio: tipo === "paragraph" && /^(1|true)$/.test(at(s, "customStyle") || "") ? val(hijo(s, "name")) || id : undefined };
        if (tipo === "paragraph" && ["1", "true"].includes(at(s, "default") || "") && idNormal == null) idNormal = id;
        if (tipo === "table" && ["1", "true"].includes(at(s, "default") || "") && idTablaNormal == null) idTablaNormal = id;
      });
    }
    // fuentes del tema: los estilos suelen decir "la fuente del cuerpo" en vez de su nombre
    const tema = {};
    {
      const destino = Object.values(rels).find((t) => /theme\/[^/]*\.xml$/.test(t || ""));
      const tt = await z.texto(destino ? rutaDe(destino) : "word/theme/theme1.xml");
      if (tt) {
        const dt = P(tt);
        ["major", "minor"].forEach((k) => {
          const f = dt.getElementsByTagNameNS("*", k + "Font")[0], l = f && [...f.childNodes].find((c) => c.localName === "latin");
          if (l && l.getAttribute("typeface")) tema[k] = l.getAttribute("typeface");
        });
      }
    }
    // numeración: numId → abstract → formato por nivel (y el texto de cada nivel: «%1.», «Capítulo %1», «%1.%2»)
    const formatoNum = {}, textoNum = {};
    // 8.2 Notas al pie: id → texto.
    const notasPie = {};
    // Notas al final (endnotes.xml): id → texto; se numeran aparte (i, ii, iii… como Word, o el formato del archivo)
    const notasFin = {};
    const finTxt = await z.texto("word/endnotes.xml");
    if (finTxt) [...P(finTxt).getElementsByTagNameNS(W, "endnote")].forEach((fn) => { const id = fn.getAttributeNS(W, "id") || fn.getAttribute("w:id"); if (+id > 0) notasFin[id] = [...fn.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("").trim(); });
    const notasTxt = await z.texto("word/footnotes.xml");
    if (notasTxt) {
      const dn = P(notasTxt);
      [...dn.getElementsByTagNameNS(W, "footnote")].forEach((fn) => {
        const id = fn.getAttributeNS(W, "id") || fn.getAttribute("w:id");
        if (+id > 0) notasPie[id] = [...fn.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("").trim();
      });
    }
    const numTxt = await z.texto("word/numbering.xml");
    if (numTxt) {
      const d = P(numTxt), abs = {};
      const absTxt = {};
      [...d.getElementsByTagNameNS(W, "abstractNum")].forEach((a) => {
        const id = a.getAttributeNS(W, "abstractNumId") || a.getAttribute("w:abstractNumId");
        abs[id] = {}; absTxt[id] = {};
        hijos(a, "lvl").forEach((l) => { const k = l.getAttributeNS(W, "ilvl") || l.getAttribute("w:ilvl"); abs[id][k] = val(hijo(l, "numFmt")); absTxt[id][k] = val(hijo(l, "lvlText")); });
      });
      [...d.getElementsByTagNameNS(W, "num")].forEach((nn) => { const nid = nn.getAttributeNS(W, "numId") || nn.getAttribute("w:numId"), a = val(hijo(nn, "abstractNumId")); formatoNum[nid] = abs[a] || {}; textoNum[nid] = absTxt[a] || {}; });
    }
    const medios = {};
    async function imagenDe(rid) {
      const destino = rels[rid];
      if (!destino) return null;
      const ruta = destino.startsWith("/") ? destino.slice(1) : "word/" + destino.replace(/^\.\//, "");
      if (medios[ruta]) return medios[ruta];
      const b = await z.leer(ruta);
      if (!b) return null;
      const ext = (ruta.split(".").pop() || "").toLowerCase();
      const mime = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp", bmp: "image/bmp", emf: "image/x-emf", wmf: "image/x-wmf", tif: "image/tiff", tiff: "image/tiff" }[ext];
      if (!mime) return null;
      let bin = "";
      for (let i = 0; i < b.length; i += 0x8000) bin += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
      return (medios[ruta] = "data:" + mime + ";base64," + btoa(bin));
    }
    const doc = P(xmlDoc), cuerpo = doc.getElementsByTagNameNS(W, "body")[0];
    const parrafos = [];
    let aproximado = false;

    function attrsRun(rpr, base) {
      const a = Object.assign({}, base);
      if (!rpr) return a;
      const b = hijo(rpr, "b"), i = hijo(rpr, "i"), u = hijo(rpr, "u"), s = hijo(rpr, "strike") || hijo(rpr, "dstrike");
      if (b) a.b = activo(b);
      if (i) a.i = activo(i);
      if (u) a.u = activo(u);
      if (s) a.s = activo(s);
      const va = val(hijo(rpr, "vertAlign"));
      if (va) a.va = va === "superscript" ? "sup" : va === "subscript" ? "sub" : null;
      const c = val(hijo(rpr, "color"));
      if (c && /^[0-9a-f]{6}$/i.test(c)) a.color = "#" + c.toUpperCase(); else if (c === "auto") delete a.color;
      const h = val(hijo(rpr, "highlight"));
      if (h && RESALTE_WORD[h]) a.fondo = RESALTE_WORD[h]; else if (h === "none") delete a.fondo;
      const shd = hijo(rpr, "shd");
      const fill = shd && (shd.getAttributeNS(W, "fill") || shd.getAttribute("w:fill"));
      if (fill && /^[0-9a-f]{6}$/i.test(fill)) a.fondo = "#" + fill.toUpperCase();
      const sz = val(hijo(rpr, "sz"));
      if (sz) a.tam = +sz / 2;
      const f = hijo(rpr, "rFonts");
      if (f) {
        // la fuente del tema (cuerpo o títulos) manda sobre el nombre escrito, igual que en Word
        const t = at(f, "asciiTheme") || at(f, "hAnsiTheme");
        const fuente = (t && tema[/^major/.test(t) ? "major" : "minor"]) || at(f, "ascii") || at(f, "hAnsi");
        if (fuente) a.fuente = fuente;
      }
      const vs = hijo(rpr, "smallCaps"), my = hijo(rpr, "caps"), es = hijo(rpr, "spacing");
      if (vs) a.vers = activo(vs);
      if (my) a.mayus = activo(my);
      if (es && val(es) != null) a.esp = Math.round((+val(es) / 20) * 100) / 100 || null;
      // Escala de caracteres de Word (Fuente → Avanzado → Escala: 50 %, 150 %, 200 %…)
      const ww = hijo(rpr, "w");
      if (ww && val(ww) != null) a.esc = +val(ww) && +val(ww) !== 100 ? +val(ww) : null;
      // texto oculto (Word no lo muestra ni lo imprime) y elevado o bajado (w:position, en medios puntos)
      const vn = hijo(rpr, "vanish"), ps = val(hijo(rpr, "position"));
      if (vn) a.oc = activo(vn);
      if (ps != null && ps !== "" && !isNaN(+ps)) a.pos = Math.round(+ps * 50) / 100 || null;
      Object.keys(a).forEach((k) => { if (a[k] == null || a[k] === false) delete a[k]; });
      return a;
    }
    // ---- lo que se hereda: valores del documento → estilo base → estilo → lo propio
    const rBase = attrsRun(rprDef, {});
    const cacheR = {}, cacheP = {};
    /** La letra con la que se ve un estilo. De un estilo de carácter sale solo lo que él cambia. */
    function letraDe(id, prof) {
      const d = id != null ? defEstilo[id] : null;
      if (!d) return rBase;
      if (cacheR[id]) return cacheR[id];
      if ((prof || 0) > 20) return rBase;
      const base = d.base != null && defEstilo[d.base] ? letraDe(d.base, (prof || 0) + 1) : d.tipo === "character" ? {} : rBase;
      return (cacheR[id] = attrsRun(d.rpr, base));
    }
    /** Propiedades de párrafo en crudo (twips), cada una sobre lo que ya traía la base. */
    function parCrudo(ppr, base) {
      const r = Object.assign({}, base);
      if (!ppr) return r;
      const jc = val(hijo(ppr, "jc"));
      if (jc) r.jc = jc;
      const sg = hijo(ppr, "suppressAutoHyphens");
      if (sg) r.sinGuiones = !["0", "false", "off"].includes(val(sg) || "1");
      const snl = hijo(ppr, "suppressLineNumbers");
      if (snl) r.sinNumLin = !["0", "false", "off"].includes(val(snl) || "1");
      // Tabulaciones de Word: [posición en pt desde el margen, tipo (i, c, d, p = decimal, b = barra), relleno]
      const tbs = hijo(ppr, "tabs");
      if (tbs) {
        let lista = (r.tabs || []).slice();
        const TIPO = { left: "i", start: "i", center: "c", right: "d", end: "d", decimal: "p", bar: "b", num: "i" };
        for (const t of Array.from(tbs.childNodes).filter((n) => n.nodeType === 1 && n.localName === "tab")) {
          const pos = +at(t, "pos") / 20, tipo = at(t, "val") || "left";
          if (isNaN(pos)) continue;
          lista = lista.filter((x) => Math.abs(x[0] - pos) > 0.05);
          if (tipo === "clear") continue;
          const rel = at(t, "leader");
          lista.push([pos, TIPO[tipo] || "i", rel && rel !== "none" ? rel : undefined]);
        }
        lista.sort((a, b) => a[0] - b[0]);
        r.tabs = lista;
      }
      const ind = hijo(ppr, "ind");
      if (ind) {
        const g = (k) => { const v = at(ind, k); return v == null || v === "" || isNaN(+v) ? null : +v; };
        const izq = g("left") != null ? g("left") : g("start"), der = g("right") != null ? g("right") : g("end");
        if (izq != null) r.izq = izq;
        if (der != null) r.der = der;
        if (g("hanging") != null) { r.fr = g("hanging"); r.pri = 0; } else if (g("firstLine") != null) { r.pri = g("firstLine"); r.fr = 0; }
      }
      const sp = hijo(ppr, "spacing");
      if (sp) {
        const g = (k) => { const v = at(sp, k); return v == null || v === "" || isNaN(+v) ? null : +v; };
        const auto = (k) => /^(1|true|on)$/.test(at(sp, k) || "");
        if (g("before") != null) r.antes = g("before");
        if (g("after") != null) r.desp = g("after");
        // espacio automático (como en HTML): Word deja 14 pt
        if (auto("beforeAutospacing")) r.antes = 280;
        if (auto("afterAutospacing")) r.desp = 280;
        if (g("line") != null) { r.linea = g("line"); r.regla = at(sp, "lineRule") || "auto"; }
      }
      // cómo se reparte entre páginas: con el siguiente, líneas juntas, salto de página antes y control de viudas
      const kn = hijo(ppr, "keepNext"), kl = hijo(ppr, "keepLines"), pbb = hijo(ppr, "pageBreakBefore"), wc = hijo(ppr, "widowControl");
      if (kn) r.kn = activo(kn);
      if (kl) r.kl = activo(kl);
      if (pbb) r.pbb = activo(pbb);
      if (wc) r.wc = activo(wc);
      const bd = hijo(ppr, "pBdr");
      if (bd) r.bdr = bd;
      const shd = hijo(ppr, "shd");
      if (shd) { const fill = at(shd, "fill"); r.shd = fill && /^[0-9a-f]{6}$/i.test(fill) && fill.toUpperCase() !== "FFFFFF" ? "#" + fill.toUpperCase() : null; }
      const np = hijo(ppr, "numPr");
      if (np) { const id = val(hijo(np, "numId")), lvl = val(hijo(np, "ilvl")); r.num = { id: id != null ? id : (r.num || {}).id, lvl: lvl != null ? lvl : (r.num || {}).lvl }; }
      const bi = hijo(ppr, "bidi");
      if (bi) r.rtl = activo(bi);
      const cs = hijo(ppr, "contextualSpacing");
      if (cs) r.junto = activo(cs);
      return r;
    }
    const pBase = parCrudo(pprDef, {});
    function parDe(id, prof) {
      const d = id != null ? defEstilo[id] : null;
      if (!d) return pBase;
      if (cacheP[id]) return cacheP[id];
      if ((prof || 0) > 20) return pBase;
      return (cacheP[id] = parCrudo(d.ppr, d.base != null && defEstilo[d.base] ? parDe(d.base, (prof || 0) + 1) : pBase));
    }
    // El estilo Normal del archivo es el "formato base" del documento: lo que no trae formato propio
    // se ve con esa fuente, ese tamaño y ese interlineado.
    const rN = letraDe(idNormal), pN = parDe(idNormal);
    // (centímetros con 4 decimales y el factor exacto: con 2 decimales una sangría se movía hasta 0.14 pt y un renglón
    // de Word que cabía justo dejaba de caber, o al revés)
    const cmDe = (tw) => Math.round((tw / (1440 / 2.54)) * 10000) / 10000;
    // (con 4 decimales: el 1.1583 de Office 2024 —line="278"— no es el 1.15 de antes; en 40 renglones son 5 pt)
    const liDe = (c) => (c.linea && (!c.regla || c.regla === "auto") ? Math.round((c.linea / 240) * 10000) / 10000 : null);
    const liDeWord = (v) => Math.abs(v - 1.15) <= 0.04 || Math.abs(v - 1.08) <= 0.03;   // lo que aquí ya es el interlineado normal
    const formato = { word: 1 };   // (vino de Word: sus párrafos llevan el formato de sus estilos, no el de los de NOM)
    if (rN.fuente) formato.fuente = rN.fuente;
    if (rN.tam) formato.tam = rN.tam; else if (stTxt) formato.tam = 10;   // sin tamaño escrito, Word usa 10 pt
    // (un archivo con estilos que no dice el interlineado va a renglón sencillo: es lo que hace Word)
    const liDeO = (c) => (c.linea ? liDe(c) : stTxt ? 1 : null);
    // El interlineado del estilo Normal va siempre tal cual (antes se omitía si se parecía a 1.15 o 1.08 y el
    // documento se veía con 1.15: los renglones de Word no caían en el mismo lugar).
    if (liDeO(pN)) formato.interlineado = liDeO(pN);
    if (pN.pri > 0) formato.primera = cmDe(pN.pri);
    // Interletraje y ligaduras del documento. Office 2024 los activa por omisión (w:kern de 1 pt y
    // w14:ligatures «standardContextual»): sin ellos el texto ocupa más que en Word y corta en otro lugar.
    const kernDe = (rpr) => { const k = hijo(rpr, "kern"); return k ? (+val(k) || 0) / 2 : null; };
    const ligDe = (rpr) => { const l = hijo(rpr, "ligatures"); return l ? l.getAttribute("w14:val") || l.getAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "val") : null; };
    const dN = idNormal != null ? defEstilo[idNormal] : null;
    const kernN = (dN && kernDe(dN.rpr)) ?? kernDe(rprDef);
    if (kernN) formato.kern = kernN;
    const ligN = (dN && ligDe(dN.rpr)) || ligDe(rprDef);
    if (ligN && ligN !== "none") formato.lig = ligN;
    // Modo de compatibilidad del archivo (Word 2013 y después = 15): de él depende cómo justifica Word.
    {
      const aj = await z.texto("word/settings.xml");
      const m = aj && /compatibilityMode"[^>]*w:val="(\d+)"|w:val="(\d+)"[^>]*compatibilityMode"/.exec(aj);
      const modo = m ? +(m[1] || m[2]) : (aj ? 12 : 15);
      if (modo < 15) formato.compat = modo;
    }
    const liBase = formato.interlineado || 1.15;
    /** Lo que es igual al formato base no se repite en cada letra. */
    const relativa = (a) => { if (a.fuente && a.fuente === formato.fuente) delete a.fuente; if (a.tam && a.tam === formato.tam) delete a.tam; return a; };
    // ---- bordes (de párrafo y de página)
    const ESTILO_BORDE = { single: "solid", thick: "solid", double: "double", triple: "double", dotted: "dotted", dashed: "dashed", dashSmallGap: "dashed", dotDash: "dashed", dotDotDash: "dashed", wave: "solid" };
    /** Un lado de un borde de Word → [grosor en puntos, estilo, color, separación en puntos]. */
    function ladoBorde(n) {
      const v = n ? val(n) : null;
      if (!v || v === "nil" || v === "none") return null;
      const e = ESTILO_BORDE[v] || "solid", sz = +(at(n, "sz") || 4) / 8, col = at(n, "color");
      return [Math.max(0.5, Math.round((e === "double" ? sz * 3 : sz) * 100) / 100), e, col && /^[0-9a-f]{6}$/i.test(col) ? "#" + col.toUpperCase() : "#000000", +(at(n, "space") || 0)];
    }
    /** Borde de párrafo → "grosor estilo color lados" (lados: s, i, z, d). */
    function bordeDe(bd) {
      const L = { s: ladoBorde(hijo(bd, "top")), i: ladoBorde(hijo(bd, "bottom")), z: ladoBorde(hijo(bd, "left") || hijo(bd, "start")), d: ladoBorde(hijo(bd, "right") || hijo(bd, "end")) };
      const lados = Object.keys(L).filter((k) => L[k]);
      if (!lados.length) return null;
      const p = L[lados[0]];
      return `${p[0]} ${p[1]} ${p[2]} ${lados.join("")}`;
    }
    // ---- texto de lo que no se dibuja aquí (cuadro de texto, SmartArt, gráfico): al menos que se lea
    async function textoObjeto(c) {
      const dentroDe = (n, nombre) => { for (let q = n.parentNode; q && q !== c; q = q.parentNode) if (q.localName === nombre) return true; return false; };
      const cajas = [...c.getElementsByTagNameNS("*", "txbxContent")].filter((t) => !dentroDe(t, "Fallback"));
      const lineas = (raiz) => [...raiz.getElementsByTagNameNS(W, "p")].map((q) => [...q.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("")).filter((t) => t.trim());
      let t = (cajas.length ? cajas : [...c.getElementsByTagNameNS("*", "txbxContent")].slice(0, 1)).map((k) => lineas(k).join("\n")).filter(Boolean).join("\n\n");
      const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
      const parte = async (rid) => { const d = rid && rels[rid]; return d ? z.texto(rutaDe(d)) : null; };
      if (!t) {
        const dg = c.getElementsByTagNameNS("*", "relIds")[0];
        const xd = dg ? await parte(dg.getAttributeNS(R, "dm") || dg.getAttribute("r:dm")) : null;
        if (xd) {
          t = [...P(xd).getElementsByTagNameNS("*", "pt")].filter((q) => !/^(pres|parTrans|sibTrans)$/.test(q.getAttribute("type") || ""))
            .map((q) => [...q.getElementsByTagNameNS("*", "t")].map((k) => k.textContent).join(" ").trim()).filter(Boolean).join("\n");
        }
      }
      if (!t) {
        const ch = [...c.getElementsByTagNameNS("*", "chart")].find((q) => q.getAttributeNS(R, "id") || q.getAttribute("r:id"));
        const xc = ch ? await parte(ch.getAttributeNS(R, "id") || ch.getAttribute("r:id")) : null;
        if (xc) {
          const dc = P(xc), ti = dc.getElementsByTagNameNS("*", "title")[0];
          const titulo = ti ? [...ti.getElementsByTagNameNS("*", "t")].map((k) => k.textContent).join("").trim() : "";
          const series = [...dc.getElementsByTagNameNS("*", "ser")].map((q) => { const tx = [...q.childNodes].find((k) => k.localName === "tx"), v = tx && tx.getElementsByTagNameNS("*", "v")[0]; return v ? v.textContent.trim() : ""; }).filter(Boolean);
          t = [titulo, series.length ? "Series: " + series.join(", ") : ""].filter(Boolean).join("\n");
        }
      }
      return t ? t.slice(0, 4000) : null;
    }
    let temaWord = null;   // colores y letras del tema del documento (para dibujar sus gráficos)
    // ---- Dibujos de Word (DrawingML): formas y cuadros de texto (wps), imágenes (pic), grupos (wpg) y lienzos (wpc),
    // con los rellenos de Office (color con transparencia, degradado, trama, imagen), sus líneas (guiones y puntas) y
    // sus efectos (sombra, resplandor, bordes suaves).
    const RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const h2 = (n, nm) => [...(n ? n.children : [])].find((x) => x.localName === nm);
    const r1 = (v) => Math.round(v * 10) / 10;
    const colW = (n) => (n ? colorConTema(n, temaWord.colores) : null);
    /** La opacidad de un color de DrawingML (1 = opaco), de su <a:alpha>. */
    function alfaDe(n) {
      const cl = n && (/Clr$/.test(n.localName || "") ? n : [...n.children].find((x) => /Clr$/.test(x.localName)));
      const a = cl && [...cl.children].find((x) => x.localName === "alpha");
      return a ? Math.round(Math.max(0, Math.min(1, +a.getAttribute("val") / 100000)) * 1000) / 1000 : 1;
    }
    /** Relleno que no es un color liso: degradado (lineal o desde un punto), trama o imagen. */
    async function pinturaWord(spPr) {
      const gf = h2(spPr, "gradFill"), pf = h2(spPr, "pattFill"), bf = h2(spPr, "blipFill");
      if (gf) {
        const paradas = [...gf.getElementsByTagNameNS("*", "gs")].map((g) => [Math.round((+(g.getAttribute("pos") || 0) / 100000) * 10000) / 10000, colW(g), alfaDe(g)]).filter((q) => q[1]).sort((a, b) => a[0] - b[0]);
        if (!paradas.length) return null;
        const lin = h2(gf, "lin"), path = h2(gf, "path");
        if (path) {
          const fr = h2(path, "fillToRect"), v = (k) => (fr ? +(fr.getAttribute(k) || 0) / 100000 : 0.5);
          return { tipo: "rad", camino: path.getAttribute("path") || "circle", paradas, cx: Math.round(((v("l") + 1 - v("r")) / 2) * 1000) / 1000, cy: Math.round(((v("t") + 1 - v("b")) / 2) * 1000) / 1000 };
        }
        return { tipo: "lin", paradas, ang: lin ? Math.round((+(lin.getAttribute("ang") || 0) / 60000) * 100) / 100 : 90 };
      }
      if (pf) return { tipo: "trama", prst: pf.getAttribute("prst") || "pct50", fg: colW(h2(pf, "fgClr")) || "#000000", bg: colW(h2(pf, "bgClr")) || "#FFFFFF", fgA: alfaDe(h2(pf, "fgClr")), bgA: alfaDe(h2(pf, "bgClr")) };
      if (bf) {
        const blip = h2(bf, "blip"), rid = blip && (blip.getAttributeNS(RNS, "embed") || blip.getAttribute("r:embed"));
        const src = rid ? await imagenDe(rid) : null;
        if (!src) return null;
        const am = blip && [...blip.children].find((x) => x.localName === "alphaModFix"), tile = h2(bf, "tile"), sr = h2(bf, "srcRect");
        const recorte = sr ? ["l", "t", "r", "b"].map((k) => +(sr.getAttribute(k) || 0) / 100000) : null;
        return { tipo: "img", src, alfa: am ? +(am.getAttribute("amt") || 100000) / 100000 : undefined,
          mosaico: tile ? { sx: +(tile.getAttribute("sx") || 100000) / 100000, sy: +(tile.getAttribute("sy") || 100000) / 100000 } : undefined,
          recorte: recorte && recorte.some(Boolean) ? recorte : undefined };
      }
      return null;
    }
    // (grosor de las líneas del tema por índice de estilo: 0.5, 1 y 1.5 pt en el tema de Office)
    const ANCHO_LN = [0, 0.5, 1, 1.5];
    /** Contorno de una forma: grosor (pt), color, transparencia, guiones, terminación y puntas de flecha. */
    function lineaDibujo(spPr, st) {
      const ln = spPr && h2(spPr, "ln"), lr = st && h2(st, "lnRef");
      if (ln && h2(ln, "noFill")) return { no: 1 };
      const idx = lr ? +lr.getAttribute("idx") || 0 : 0;
      if (!ln && !idx) return null;
      const sf = ln && (h2(ln, "solidFill") || h2(ln, "gradFill"));
      const fuente = sf && sf.localName === "gradFill" ? sf.getElementsByTagNameNS("*", "gs")[0] : sf;
      const pd = ln && h2(ln, "prstDash"), cap = ln && ln.getAttribute("cap");
      const punta = (k) => { const e = ln && h2(ln, k), t = e && e.getAttribute("type"); return t && t !== "none" ? { t, w: e.getAttribute("w") || "med", l: e.getAttribute("len") || "med" } : null; };
      const o = { w: ln && ln.getAttribute("w") ? +ln.getAttribute("w") / 12700 : ANCHO_LN[Math.min(3, idx)] || 0.75, c: colW(fuente) || colW(lr) || "#000000" };
      const a = fuente ? alfaDe(fuente) : 1;
      if (a < 1) o.alfa = a;
      if (pd && pd.getAttribute("val") && pd.getAttribute("val") !== "solid") o.guion = pd.getAttribute("val");
      if (cap === "rnd" || cap === "sq") o.tapa = cap === "rnd" ? "round" : "square";
      if (ln && ln.getAttribute("cmpd") && ln.getAttribute("cmpd") !== "sng") o.cmpd = ln.getAttribute("cmpd");
      if (punta("headEnd") || punta("tailEnd")) o.puntas = { ini: punta("headEnd"), fin: punta("tailEnd") };
      return o;
    }
    /** Efectos de la forma: sombra exterior, resplandor y bordes suaves (en px). */
    function efectosWord(spPr) {
      const el = spPr && h2(spPr, "effectLst"), o = {};
      if (!el) return o;
      const sh = h2(el, "outerShdw");
      if (sh) { const d = +(sh.getAttribute("dist") || 0) / EMU, dir = ((+(sh.getAttribute("dir") || 0) / 60000) * Math.PI) / 180; o.sombra = { dx: r1(d * Math.cos(dir)), dy: r1(d * Math.sin(dir)), b: r1(+(sh.getAttribute("blurRad") || 0) / EMU), c: colW(sh) || "#000000", a: alfaDe(sh) }; }
      const gl = h2(el, "glow");
      if (gl) o.brillo = { r: r1(+(gl.getAttribute("rad") || 0) / EMU), c: colW(gl) || "#FFC000", a: alfaDe(gl) };
      const se = h2(el, "softEdge");
      if (se) o.suave = r1(+(se.getAttribute("rad") || 0) / EMU);
      return o;
    }
    /** Geometría libre (custGeom) en el formato de geometrias-datos.js. */
    function geomLibreW(spPr) {
      const cg = spPr && h2(spPr, "custGeom");
      if (!cg) return null;
      const kids = (n) => [...(n ? n.children : [])];
      const num = (v) => (v != null && v !== "" && !isNaN(+v) ? +v : v || 0);
      const guias = (n) => kids(n).filter((g) => g.localName === "gd").map((g) => [g.getAttribute("name"), g.getAttribute("fmla") || "val 0"]);
      const pts = (c) => [...c.getElementsByTagNameNS("*", "pt")].flatMap((q) => [num(q.getAttribute("x")), num(q.getAttribute("y"))]);
      const def = { p: [] }, av = guias(h2(cg, "avLst")), gd = guias(h2(cg, "gdLst"));
      if (av.length) def.av = av;
      if (gd.length) def.gd = gd;
      const rc = h2(cg, "rect");
      if (rc) def.rect = ["l", "t", "r", "b"].map((k) => num(rc.getAttribute(k)));
      kids(h2(cg, "pathLst")).filter((q) => q.localName === "path").forEach((q) => {
        const t = { c: [] };
        if (q.getAttribute("w")) t.w = +q.getAttribute("w");
        if (q.getAttribute("h")) t.h = +q.getAttribute("h");
        if (q.getAttribute("fill") && q.getAttribute("fill") !== "norm") t.f = q.getAttribute("fill");
        if (q.getAttribute("stroke") === "0" || q.getAttribute("stroke") === "false") t.s = 0;
        kids(q).forEach((c) => {
          const n = c.localName;
          if (n === "moveTo") t.c.push(["M", ...pts(c)]);
          else if (n === "lnTo") t.c.push(["L", ...pts(c)]);
          else if (n === "arcTo") t.c.push(["A", ...["wR", "hR", "stAng", "swAng"].map((k) => num(c.getAttribute(k)))]);
          else if (n === "quadBezTo") t.c.push(["Q", ...pts(c)]);
          else if (n === "cubicBezTo") t.c.push(["C", ...pts(c)]);
          else if (n === "close") t.c.push(["Z"]);
        });
        if (t.c.length) def.p.push(t);
      });
      return def.p.length ? def : null;
    }
    /** Lo que se dibuja de una forma de Word (wps:wsp): geometría, relleno, contorno, efectos, texto y su letra. */
    async function formaDeWsp(wsp) {
      const spPr = h2(wsp, "spPr"), st = h2(wsp, "style"), prstG = spPr && h2(spPr, "prstGeom");
      const cust = !prstG ? geomLibreW(spPr) : null;
      const prst = prstG ? prstG.getAttribute("prst") : "rect", av = {};
      if (prstG) [...prstG.getElementsByTagNameNS("*", "gd")].forEach((g) => { const m = /^val\s+(-?\d+)/.exec(g.getAttribute("fmla") || ""); if (m) av[g.getAttribute("name")] = +m[1]; });
      let relleno = null, alfa, pintura = null;
      if (spPr) {
        const sf = h2(spPr, "solidFill");
        if (h2(spPr, "noFill")) relleno = "none";
        else if (sf) { relleno = colW(sf); const a = alfaDe(sf); if (a < 1) alfa = a; }
        else if (h2(spPr, "gradFill") || h2(spPr, "pattFill") || h2(spPr, "blipFill")) {
          pintura = await pinturaWord(spPr);
          if (pintura) relleno = pintura.paradas ? pintura.paradas[0][1] : pintura.fg || "#808080";
        }
      }
      if (relleno == null && !pintura && st) { const fr = h2(st, "fillRef"); if (fr && +fr.getAttribute("idx")) relleno = colW(fr); }
      const tx = wsp.getElementsByTagNameNS("*", "txbxContent")[0];
      const vw = (n) => (n ? n.getAttributeNS(W, "val") || n.getAttribute("w:val") : null);
      const parrafos = tx ? [...tx.getElementsByTagNameNS(W, "p")].map((q) => ({ t: [...q.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join(""), al: vw(q.getElementsByTagNameNS(W, "jc")[0]) })) : [];
      const rr = tx && tx.getElementsByTagNameNS(W, "rPr")[0], sz = rr && rr.getElementsByTagNameNS(W, "sz")[0], cr = rr && rr.getElementsByTagNameNS(W, "color")[0], b0 = rr && rr.getElementsByTagNameNS(W, "b")[0];
      const rf = rr && rr.getElementsByTagNameNS(W, "rFonts")[0], fuente = rf ? rf.getAttributeNS(W, "ascii") || rf.getAttribute("w:ascii") : null;
      const fontRef = st && h2(st, "fontRef");
      const letra = { tam: sz ? +vw(sz) / 2 : 11, c: cr && /^[0-9a-f]{6}$/i.test(vw(cr) || "") ? "#" + vw(cr).toUpperCase() : colW(fontRef) || "#000000", b: !!b0 && vw(b0) !== "0" && vw(b0) !== "false" };
      if (fuente) letra.f = fuente;
      const bp = h2(wsp, "bodyPr");
      const ins = bp ? ["lIns", "tIns", "rIns", "bIns"].map((k, i) => (bp.getAttribute(k) != null ? +bp.getAttribute(k) / EMU : [9.6, 4.8, 9.6, 4.8][i])) : [9.6, 4.8, 9.6, 4.8];
      const xf = spPr && h2(spPr, "xfrm");
      const giro = xf && xf.getAttribute("rot") ? +xf.getAttribute("rot") / 60000 : 0, vH = xf && xf.getAttribute("flipH") === "1", vV = xf && xf.getAttribute("flipV") === "1";
      const vert = bp && /^(vert|vert270|eaVert|wordArtVert)$/.test(bp.getAttribute("vert") || "") ? bp.getAttribute("vert") : null;
      return Object.assign({ prst: cust ? undefined : prst, cust: cust || undefined, av: Object.keys(av).length ? av : undefined, relleno, alfa, pintura: pintura || undefined,
        linea: lineaDibujo(spPr, st), texto: parrafos.some((q) => q.t) ? parrafos : undefined, letra, ancla: (bp && bp.getAttribute("anchor")) || "t", ins,
        giro: giro || undefined, vH: vH || undefined, vV: vV || undefined, vert: vert || undefined }, efectosWord(spPr));
    }
    /** Una imagen dentro de un grupo o de un lienzo (pic:pic): su archivo, su recorte, su borde y sus efectos. */
    async function imagenDePic(pic) {
      const bf = h2(pic, "blipFill"), blip = bf && h2(bf, "blip"), rid = blip && (blip.getAttributeNS(RNS, "embed") || blip.getAttribute("r:embed"));
      const src = rid ? await imagenDe(rid) : null;
      if (!src) return null;
      const sr = bf && h2(bf, "srcRect"), recorte = sr ? ["l", "t", "r", "b"].map((k) => +(sr.getAttribute(k) || 0) / 100000) : null;
      const spPr = h2(pic, "spPr");
      return Object.assign({ src, recorte: recorte && recorte.some(Boolean) ? recorte : undefined, linea: lineaDibujo(spPr, null) || undefined }, efectosWord(spPr));
    }
    /** Las piezas de un grupo (wpg) o de un lienzo (wpc), en px dentro de la caja de w × h. */
    async function piezasDe(cont, w, h, esLienzo) {
      const gx = esLienzo ? null : h2(h2(cont, "grpSpPr"), "xfrm");
      const ch = gx && h2(gx, "chOff"), ce = gx && h2(gx, "chExt");
      const ox = ch ? +ch.getAttribute("x") || 0 : 0, oy = ch ? +ch.getAttribute("y") || 0 : 0;
      const sx = ce && +ce.getAttribute("cx") ? (w * EMU) / +ce.getAttribute("cx") : 1, sy = ce && +ce.getAttribute("cy") ? (h * EMU) / +ce.getAttribute("cy") : 1;
      const piezas = [];
      for (const el of cont.children) {
        const nm = el.localName;
        if (!["wsp", "pic", "grpSp", "cxnSp"].includes(nm)) continue;
        const sp = h2(el, nm === "grpSp" ? "grpSpPr" : "spPr"), xf = sp && h2(sp, "xfrm"), off = xf && h2(xf, "off"), ex = xf && h2(xf, "ext");
        if (!off || !ex) continue;
        const p = { x: r1(((+off.getAttribute("x") - ox) * sx) / EMU), y: r1(((+off.getAttribute("y") - oy) * sy) / EMU), w: r1((+ex.getAttribute("cx") * sx) / EMU), h: r1((+ex.getAttribute("cy") * sy) / EMU) };
        const giro = +(xf.getAttribute("rot") || 0) / 60000;
        if (giro) p.giro = giro;
        if (xf.getAttribute("flipH") === "1") p.vH = true;
        if (xf.getAttribute("flipV") === "1") p.vV = true;
        if (nm === "wsp" || nm === "cxnSp") { p.forma = await formaDeWsp(el); delete p.forma.giro; delete p.forma.vH; delete p.forma.vV; }
        else if (nm === "pic") { p.img = await imagenDePic(el); if (!p.img) continue; }
        else p.grupo = { hijos: await piezasDe(el, p.w, p.h, false) };
        piezas.push(p);
      }
      return piezas;
    }
    /** Dónde va un dibujo anclado (wp:anchor) de Word: en su lugar, delante o detrás del texto, o «arriba y abajo». */
    function lugarDe(c, w, h) {
      const fl = comoFlota(c, w, h), anc = c.getElementsByTagNameNS("*", "anchor")[0];
      const rh = anc ? Math.max(0, Math.min(9000, Math.round((+(anc.getAttribute("relativeHeight") || 251658240) - 251658240) / 1024))) : 0;
      return anc ? { libre: true, x: fl.libre ? fl.x : xDeFlota(c, w), y: fl.libre ? fl.y : yDeFlota(c), detras: fl.detras || anc.getAttribute("behindDoc") === "1" || undefined, franja: fl.ajuste === "bloque" || undefined, z: rh || undefined, fl: fl.fl } : {};
    }
    /**
     * Cómo pone Word un dibujo anclado, tal cual (lo usa la página para ponerlo en su hoja y apartar el texto):
     *   h, v: [de qué se mide, alineación o distancia en px] — horizontal: page, margin, column, character, leftMargin,
     *         rightMargin, insideMargin, outsideMargin; vertical: page, margin, paragraph, line, topMargin, bottomMargin…
     *         (la distancia también puede ser {pct} del ancho o alto de lo que se mide)
     *   a: ajuste del texto (cuadrado, estrecho, traves, arriba = arriba y abajo, nada = delante o detrás del texto)
     *   l: de qué lado va el texto (ambos, izq = solo a la izquierda, der = solo a la derecha, mayor = del más ancho)
     *   d: distancias al texto [arriba, abajo, izquierda, derecha]; ee: lo que sobresalen los efectos [izq, arriba, der, abajo]
     *   p: contorno para el ajuste estrecho o a través (puntos de 0 a 1 sobre el dibujo); t: detrás del texto
     */
    /** positionH o positionV de un ancla: directo o dentro de mc:AlternateContent (Word 2010 escribe la posición en
        porcentaje en el Choice «wp14» y en EMU en el Fallback; aquí se entienden las dos, así que va el Choice). */
    function posicionDe(anc, nm) {
      for (const x of anc.children) {
        if (x.localName === nm) return x;
        if (x.localName === "AlternateContent") {
          const ch = [...x.children].find((q) => q.localName === "Choice"), fb = [...x.children].find((q) => q.localName === "Fallback");
          for (const r of [ch, fb]) { const e = r && [...r.children].find((q) => q.localName === nm); if (e) return e; }
        }
      }
      return null;
    }
    function flotaDe(anc) {
      const hijo2 = (n, nm) => [...(n ? n.children : [])].find((x) => x.localName === nm);
      const r2 = (x) => Math.round(x * 100) / 100;
      const pos = (nm, pctNm) => {
        const e = posicionDe(anc, nm);
        if (!e) return null;
        const o = hijo2(e, "posOffset"), a = hijo2(e, "align"), pc = hijo2(e, pctNm);
        return [e.getAttribute("relativeFrom") || "", a ? a.textContent.trim() : pc ? { pct: +pc.textContent / 100000 } : o ? r2(+o.textContent / EMU) : 0];
      };
      const sp = anc.getAttribute("simplePos") === "1" ? hijo2(anc, "simplePos") : null;
      const wrap = [...anc.children].find((x) => /^wrap/.test(x.localName));
      const fl = {
        h: sp ? ["page", r2(+(sp.getAttribute("x") || 0) / EMU)] : pos("positionH", "pctPosHOffset") || ["column", 0],
        v: sp ? ["page", r2(+(sp.getAttribute("y") || 0) / EMU)] : pos("positionV", "pctPosVOffset") || ["paragraph", 0],
        a: { wrapSquare: "cuadrado", wrapTight: "estrecho", wrapThrough: "traves", wrapTopAndBottom: "arriba", wrapNone: "nada" }[wrap ? wrap.localName : "wrapNone"] || "nada",
      };
      const lado = wrap && wrap.getAttribute("wrapText");
      if (lado && lado !== "bothSides") fl.l = { left: "izq", right: "der", largest: "mayor" }[lado] || undefined;
      const d = ["distT", "distB", "distL", "distR"].map((k) => r2(+(anc.getAttribute(k) || 0) / EMU));
      if (d.some((x) => x)) fl.d = d;
      const ee = hijo2(anc, "effectExtent");
      if (ee) { const v = ["l", "t", "r", "b"].map((k) => r2(+(ee.getAttribute(k) || 0) / EMU)); if (v.some((x) => x)) fl.ee = v; }
      const poli = wrap && hijo2(wrap, "wrapPolygon");
      if (poli && (fl.a === "estrecho" || fl.a === "traves")) {
        const pts = [...poli.children].filter((x) => x.localName === "start" || x.localName === "lineTo").map((x) => [Math.round((+x.getAttribute("x") / 21600) * 10000) / 10000, Math.round((+x.getAttribute("y") / 21600) * 10000) / 10000]);
        if (pts.length >= 3) fl.p = pts;
      }
      if (anc.getAttribute("behindDoc") === "1") fl.t = true;
      return fl;
    }
    /** Una forma o un cuadro de texto de Word (wps:wsp), en su lugar (desde el párrafo y la columna). Si Word la rodea
        arriba y abajo, su franja se aparta del texto. */
    async function formaWord(c, w, h) {
      const wsp = c.getElementsByTagNameNS("*", "wsp")[0];
      if (!wsp || !w) return null;
      if (!temaWord) temaWord = temaDeXml(await z.texto("word/theme/theme1.xml"));
      return Object.assign({ forma: await formaDeWsp(wsp), w, h }, lugarDe(c, w, h));
    }
    /** Un grupo de formas (wpg:wgp) o un lienzo de dibujo (wpc:wpc) de Word, con todas sus piezas, en su lugar. */
    async function grupoWord(c, w, h) {
      if (!w) return null;
      if (!temaWord) temaWord = temaDeXml(await z.texto("word/theme/theme1.xml"));
      const wgp = c.getElementsByTagNameNS("*", "wgp")[0], wpc = !wgp ? c.getElementsByTagNameNS("*", "wpc")[0] : null, cont = wgp || wpc;
      if (!cont) return null;
      const grupo = { hijos: await piezasDe(cont, w, h, !!wpc) };
      if (wgp) {
        const xf = h2(h2(wgp, "grpSpPr"), "xfrm");
        if (xf && +xf.getAttribute("rot")) grupo.giro = +xf.getAttribute("rot") / 60000;
        if (xf && xf.getAttribute("flipH") === "1") grupo.vH = true;
        if (xf && xf.getAttribute("flipV") === "1") grupo.vV = true;
      } else {
        // el lienzo: su fondo y su contorno
        const bg = h2(wpc, "bg"), wh = h2(wpc, "whole");
        const sf = bg && h2(bg, "solidFill");
        if (sf) grupo.fondo = colW(sf);
        const l = wh ? lineaDibujo(wh, null) : null;
        if (l && !l.no) grupo.linea = l;
      }
      return Object.assign({ grupo, w, h }, lugarDe(c, w, h));
    }
    // (la posición de un ancla sin importar su ajuste, para las formas)
    function xDeFlota(c, w) { const t = comoFlota(c, w, 1, true); return t.x || 0; }
    function yDeFlota(c) { const t = comoFlota(c, 1, 1, true); return t.y || 0; }
    /** Cómo va una imagen anclada de Word (wp:anchor) en la página de NOM: sin ajuste (libre, en su lugar,
        delante o detrás del texto), cuadrado/estrecho/transparente (a la izquierda o a la derecha del texto) o arriba
        y abajo (en su propio renglón, con su alineación). Las medidas, desde el párrafo y la columna. */
    function comoFlota(c, w, h, soloPosicion) {
      const anc = c.getElementsByTagNameNS("*", "anchor")[0];
      if (!anc) return {};
      const hijo2 = (n, nm) => [...(n ? n.children : [])].find((x) => x.localName === nm);
      const wrap = [...anc.children].find((x) => /^wrap/.test(x.localName));
      const pos = (nm) => { const e = posicionDe(anc, nm); if (!e) return null; const o = hijo2(e, "posOffset"), a = hijo2(e, "align"); return { rel: e.getAttribute("relativeFrom") || "", off: o ? Math.round(+o.textContent / EMU) : null, al: a ? a.textContent.trim() : null }; };
      const ph = pos("positionH"), pv = pos("positionV");
      // ancho del texto (columna) y márgenes de la página (de la primera sección; en px a 96 ppp)
      // (la primera sección: la de la lista si ya se leyó; si no, la primera «sectPr» del documento)
      const sec = (secciones[0] && secciones[0].sect) || cuerpo.getElementsByTagNameNS(W, "sectPr")[0] || null, pgSz = sec && hijo(sec, "pgSz"), pgMar = sec && hijo(sec, "pgMar");
      const anchoPag = pgSz ? Math.round(+(pgSz.getAttributeNS(W, "w") || pgSz.getAttribute("w:w") || 12240) / 15) : 816;
      const mIz = pgMar ? Math.round(+(pgMar.getAttributeNS(W, "left") || pgMar.getAttribute("w:left") || 1440) / 15) : 96;
      const mDe = pgMar ? Math.round(+(pgMar.getAttributeNS(W, "right") || pgMar.getAttribute("w:right") || 1440) / 15) : 96;
      const mSu = pgMar ? Math.round(+(pgMar.getAttributeNS(W, "top") || pgMar.getAttribute("w:top") || 1440) / 15) : 96;
      const col = anchoPag - mIz - mDe;
      // posición horizontal dentro de la columna
      let x = 0;
      if (ph) {
        if (ph.off != null) x = ph.rel === "page" ? ph.off - mIz : ph.off;
        else if (ph.al === "right" || ph.al === "outside") x = ph.rel === "page" ? anchoPag - mIz - w : col - w;
        else if (ph.al === "center") x = (ph.rel === "page" ? anchoPag : col) / 2 - w / 2 - (ph.rel === "page" ? mIz : 0);
      }
      const y = pv && pv.off != null ? (pv.rel === "page" || pv.rel === "topMargin" ? pv.off - mSu : pv.off) : 0;
      const tipo = wrap ? wrap.localName : "wrapNone";
      if (soloPosicion) return { x: Math.round(x), y: Math.round(y) };
      // (qué va encima de qué: Word lo guarda en relativeHeight, de 1024 en 1024 desde 251658240)
      const z = Math.max(0, Math.min(9000, Math.round((+(anc.getAttribute("relativeHeight") || 251658240) - 251658240) / 1024))) || undefined;
      // (y cómo lo pone Word, completo: con esto la página lo coloca en su hoja y aparta el texto como Word)
      const fl = flotaDe(anc);
      if (tipo === "wrapNone") return { libre: true, x: Math.round(x), y: Math.round(y), detras: anc.getAttribute("behindDoc") === "1" || undefined, z, fl };
      if (tipo === "wrapTopAndBottom") return { ajuste: "bloque", alin: x + w / 2 < col * 0.4 ? (x > 4 ? undefined : "izq") : x + w / 2 > col * 0.6 ? "der" : "centro", dx: x > 4 && x + w / 2 < col * 0.4 ? Math.round(x) : undefined, z, fl };
      // cuadrado, estrecho o transparente: el texto la rodea por el lado libre
      return { ajuste: x + w / 2 > col / 2 ? "der" : "izq", z, fl };
    }
    /** El gráfico de un dibujo de Word (también dentro de mc:AlternateContent): su modelo de Office y su tamaño. */
    async function graficaWord(c) {
      if (!NOM.graficaOffice) return null;
      const RR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
      const ch = [...c.getElementsByTagNameNS("*", "chart")].find((q) => q.getAttributeNS(RR, "id") || q.getAttribute("r:id"));
      const ridG = ch && (ch.getAttributeNS(RR, "id") || ch.getAttribute("r:id")), dG = ridG && rels[ridG];
      const xc = dG ? await z.texto(rutaDe(dG)) : null;
      if (!xc) return null;
      // el tamaño: el «extent» del dibujo que contiene ese gráfico
      let dib = ch; while (dib && dib.localName !== "inline" && dib.localName !== "anchor") dib = dib.parentNode;
      const ext = dib && [...dib.children].find((q) => q.localName === "extent");
      const w = ext ? Math.round(+ext.getAttribute("cx") / EMU) : 0, h = ext ? Math.round(+ext.getAttribute("cy") / EMU) : 0;
      if (!w || !h) return null;
      if (!temaWord) temaWord = temaDeXml(await z.texto("word/theme/theme1.xml"));
      try {
        const g = NOM.graficaOffice.leer(xc, { color: (n) => colorConTema(n, temaWord.colores), fuente: (n) => (/^\+mn-/.test(n) ? temaWord.fuentes.menor : /^\+mj-/.test(n) ? temaWord.fuentes.mayor : n) || "Calibri", acentos: [1, 2, 3, 4, 5, 6].map((k) => "#" + (temaWord.colores["accent" + k] || "4472C4")) });
        return { g, w, h };
      } catch (e) { return null; }
    }
    /** Las partes de un estilo de tabla de Word con su herencia (basedOn): rPr, tcPr y tblPr de cada tipo de región. */
    function partesEstiloTabla(id, prof) {
      const d = id != null ? defEstilo[id] : null;
      if (!d || !d.nodo || (prof || 0) > 8) return {};
      const out = Object.assign({}, d.base ? partesEstiloTabla(d.base, (prof || 0) + 1) : {});
      const una = (n) => ({ rpr: hijo(n, "rPr"), tcpr: hijo(n, "tcPr"), tblpr: hijo(n, "tblPr"), ppr: hijo(n, "pPr") });
      const todo = una(d.nodo);
      out.wholeTable = Object.assign({}, out.wholeTable || {}, Object.fromEntries(Object.entries(todo).filter(([, v]) => v)));
      hijos(d.nodo, "tblStylePr").forEach((x) => { const t = x.getAttributeNS(W, "type") || x.getAttribute("w:type"); out[t] = Object.assign({}, out[t] || {}, Object.fromEntries(Object.entries(una(x)).filter(([, v]) => v))); });
      return out;
    }
    const lineaWord = (b) => {
      if (!b) return null;
      const v = b.getAttributeNS(W, "val") || b.getAttribute("w:val");
      if (!v || v === "nil" || v === "none") return "no";
      const sz = +(b.getAttributeNS(W, "sz") || b.getAttribute("w:sz") || 4) / 8, col = b.getAttributeNS(W, "color") || b.getAttribute("w:color");
      return [{ double: "double", dotted: "dotted", dashed: "dashed", dashSmallGap: "dashed", dotDash: "dashed" }[v] || "solid", Math.max(0.5, Math.round(sz * (96 / 72) * 100) / 100), /^[0-9a-f]{6}$/i.test(col || "") ? "#" + col.toUpperCase() : "#000000"];
    };
    /**
     * Los bordes que la tabla y cada celda traen escritos (encima de los del estilo, como Word): los de la tabla
     * (tblBorders: afuera y entre celdas) y los de cada celda (tcBorders, con sus diagonales). Si hay, la tabla va sin
     * las líneas de NOM por omisión: cada celda lleva las suyas.
     */
    function bordesDirectosWord(tbl, celdas, filas, nF, nC) {
      const tb = hijo(hijo(tbl, "tblPr"), "tblBorders");
      const lado = (b, n, n2) => (b ? lineaWord(hijo(b, n)) || (n2 ? lineaWord(hijo(b, n2)) : null) : null);
      const leer = (s) => { const o = {}; String(s || "").split(";").forEach((q) => { const [l, ...r] = q.split(":"); if (l) o[l] = r[0] === "no" ? "no" : r.length >= 3 ? [r[0], +r[1], r[2]] : null; }); return o; };
      const escribir = (bd) => Object.entries(bd).filter(([, v]) => v).map(([l, v]) => (v === "no" ? `${l}:no` : `${l}:${v[0]}:${v[1]}:${v[2]}`)).join(";");
      let hay = false;
      filas.forEach((tr, f) => {
        let k = 0;
        hijos(tr, "tc").forEach((tc) => {
          const tcPr = hijo(tc, "tcPr"), span = Math.max(1, +(val(hijo(tcPr, "gridSpan")) || 1)), R = F().ref(k, f);
          const vm = hijo(tcPr, "vMerge"), sigue = vm && (val(vm) || "continue") === "continue";
          const k0 = k;
          k += span;
          if (sigue || k0 >= nC) return;
          const m = String(celdas["m:" + R] || "1x1").split("x").map(Number), alto = Math.max(1, m[1] || 1);
          const tcb = hijo(tcPr, "tcBorders");
          if (!tb && !tcb) return;
          const bd = leer(celdas["bxw:" + R]);
          if (tb) {
            const a = f === 0 ? lado(tb, "top") : lado(tb, "insideH"), b = f + alto - 1 >= nF - 1 ? lado(tb, "bottom") : lado(tb, "insideH");
            const i = k0 === 0 ? lado(tb, "left", "start") : lado(tb, "insideV"), d = k0 + span >= nC ? lado(tb, "right", "end") : lado(tb, "insideV");
            if (a) bd.a = a; if (b) bd.b = b; if (i) bd.i = i; if (d) bd.d = d;
          }
          if (tcb) {
            const a = lado(tcb, "top"), b = lado(tcb, "bottom"), i = lado(tcb, "left", "start"), d = lado(tcb, "right", "end");
            if (a) bd.a = a; if (b) bd.b = b; if (i) bd.i = i; if (d) bd.d = d;
            // diagonales: de arriba a la izquierda hacia abajo a la derecha (tl2br) y al revés (tr2bl)
            const dg = [["tl", lado(tcb, "tl2br")], ["tr", lado(tcb, "tr2bl")]].filter(([, v]) => Array.isArray(v)).map(([l, v]) => `${l}:${v[0]}:${v[1]}:${v[2]}`).join(";");
            if (dg) celdas["bxd:" + R] = dg;
          }
          const txt = escribir(bd);
          if (txt) { celdas["bxw:" + R] = txt; hay = true; }
        });
      });
      // (cada celda lleva sus líneas: las de NOM por omisión, no)
      if (hay) celdas.bordes = "ninguno";
    }
    function estiloTablaWord(tbl, celdas, nF, nC) {
      const tblPr = hijo(tbl, "tblPr"), idEst = val(hijo(tblPr, "tblStyle"));
      const partes = partesEstiloTabla(idEst);
      if (!Object.keys(partes).length) return;
      // con estilo, las líneas son las del estilo (cada celda las lleva): sin las de NOM por omisión
      if (!hijo(tblPr, "tblBorders")) celdas.bordes = "ninguno";
      const lk = hijo(tblPr, "tblLook"), si = (a, def) => { if (!lk) return def; const v = lk.getAttributeNS(W, a) || lk.getAttribute("w:" + a); if (v != null) return v === "1" || v === "true"; const hx = parseInt(val(lk) || "", 16); if (isNaN(hx)) return def; return !!(hx & { firstRow: 0x20, lastRow: 0x40, firstColumn: 0x80, lastColumn: 0x100, noHBand: 0x200, noVBand: 0x400 }[a]); };
      const fl = { fr: si("firstRow", true), lr: si("lastRow", false), fc: si("firstColumn", true), lc: si("lastColumn", false), bh: !si("noHBand", false), bv: !si("noVBand", true) };
      const tamBanda = (k) => Math.max(1, +(val(hijo(tblPr, k)) || val(hijo((partes.wholeTable || {}).tblpr || null, k)) || 1));
      const bF = tamBanda("tblStyleRowBandSize"), bC = tamBanda("tblStyleColBandSize");
      const cubierta = new Set();
      Object.keys(celdas).filter((k) => k.startsWith("m:")).forEach((k) => { const tl = F().parseRef(k.slice(2)), [w2, h2] = String(celdas[k]).split("x").map(Number); for (let f = tl.f; f < tl.f + h2; f++) for (let c = tl.c; c < tl.c + w2; c++) if (f !== tl.f || c !== tl.c) cubierta.add(F().ref(c, f)); });
      for (let f = 0; f < nF; f++) for (let c = 0; c < nC; c++) {
        const R = F().ref(c, f);
        if (cubierta.has(R)) continue;
        const capas = [], pon = (t) => { if (partes[t]) capas.push([t, partes[t]]); };
        pon("wholeTable");
        const f0 = fl.fr ? 1 : 0, fN = fl.lr ? nF - 2 : nF - 1, c0 = fl.fc ? 1 : 0, cN = fl.lc ? nC - 2 : nC - 1;
        if (fl.bv && c >= c0 && c <= cN) pon(Math.floor((c - c0) / bC) % 2 ? "band2Vert" : "band1Vert");
        if (fl.bh && f >= f0 && f <= fN) pon(Math.floor((f - f0) / bF) % 2 ? "band2Horz" : "band1Horz");
        if (fl.fc && c === 0) pon("firstCol");
        if (fl.lc && c === nC - 1) pon("lastCol");
        if (fl.fr && f === 0) pon("firstRow");
        if (fl.lr && f === nF - 1) pon("lastRow");
        if (fl.fr && fl.fc && f === 0 && c === 0) pon("nwCell");
        if (fl.fr && fl.lc && f === 0 && c === nC - 1) pon("neCell");
        if (fl.lr && fl.fc && f === nF - 1 && c === 0) pon("swCell");
        if (fl.lr && fl.lc && f === nF - 1 && c === nC - 1) pon("seCell");
        let fondo = null, negrita = null, cursiva = null, color = null, mayus = null;
        const bd = {};
        for (const [t, x] of capas) {
          const shd = x.tcpr && hijo(x.tcpr, "shd"), fill = shd && (shd.getAttributeNS(W, "fill") || shd.getAttribute("w:fill"));
          if (fill && /^[0-9a-f]{6}$/i.test(fill)) fondo = "#" + fill.toUpperCase();
          // (mayúsculas del estilo: «Plain Table 3» pone en mayúsculas la primera fila y la primera columna)
          if (x.rpr) { const cp = hijo(x.rpr, "caps"); if (cp) mayus = val(cp) !== "0" && val(cp) !== "false"; }
          if (x.rpr) { const b = hijo(x.rpr, "b"); if (b) negrita = val(b) !== "0" && val(b) !== "false"; const i = hijo(x.rpr, "i"); if (i) cursiva = val(i) !== "0" && val(i) !== "false"; const k = val(hijo(x.rpr, "color")); if (k && /^[0-9a-f]{6}$/i.test(k)) color = "#" + k.toUpperCase(); }
          // bordes: los de la tabla (tblBorders: afuera y adentro) y los de la celda en esa región (tcBorders)
          const tb2 = t === "wholeTable" && x.tblpr ? hijo(x.tblpr, "tblBorders") : null;
          if (tb2) {
            const lado = (n) => lineaWord(hijo(tb2, n));
            const a = f === 0 ? lado("top") : lado("insideH"), ab = f === nF - 1 ? lado("bottom") : lado("insideH"), iz = c === 0 ? lado("left") || lado("start") : lado("insideV"), de = c === nC - 1 ? lado("right") || lado("end") : lado("insideV");
            if (a) bd.a = a; if (ab) bd.b = ab; if (iz) bd.i = iz; if (de) bd.d = de;
          }
          const tcb = x.tcpr && hijo(x.tcpr, "tcBorders");
          if (tcb) {
            const lado = (n) => lineaWord(hijo(tcb, n));
            const a = lado("top"), ab = lado("bottom"), iz = lado("left") || lado("start"), de = lado("right") || lado("end");
            if (a) bd.a = a; if (ab) bd.b = ab; if (iz) bd.i = iz; if (de) bd.d = de;
          }
        }
        if (fondo && !celdas["r:" + R] && fondo !== "#FFFFFF") celdas["r:" + R] = fondo;
        if (negrita && celdas["n:" + R] == null && celdas["c:" + R]) celdas["n:" + R] = true;
        if (cursiva && celdas["i:" + R] == null && celdas["c:" + R]) celdas["i:" + R] = true;
        if (color && !celdas["k:" + R] && color !== "#000000") celdas["k:" + R] = color;
        if (mayus && celdas["my:" + R] == null) celdas["my:" + R] = true;
        const txt = Object.entries(bd).filter(([, v]) => v).map(([l, v]) => (v === "no" ? `${l}:no` : `${l}:${v[0]}:${v[1]}:${v[2]}`)).join(";");
        if (txt) celdas["bxw:" + R] = txt;
      }
    }
    async function runs(n, base, tramos, parrafo) {
      for (const r of hijos(n)) {
        const ln = r.localName;
        if (ln === "r") {
          const rpr = hijo(r, "rPr"), rs = rpr ? val(hijo(rpr, "rStyle")) : null;
          const a = relativa(attrsRun(rpr, rs != null && defEstilo[rs] ? Object.assign({}, base, letraDe(rs)) : base));
          for (const c of hijos(r)) {
            const cn = c.localName;
            if (cn === "t" || cn === "delText") tramos.push({ x: c.textContent, a });   // (delText: lo borrado con cambios marcados)
            else if (cn === "tab") tramos.push({ x: "\t", a });
            else if (cn === "br" || cn === "cr") {
              const tipo = c.getAttributeNS(W, "type") || c.getAttribute("w:type");
              // (salto de p\u00e1gina o de columna: d\u00f3nde va dentro del p\u00e1rrafo, para partirlo ah\u00ed como Word)
              if (tipo === "page" || tipo === "column") { if (tipo === "page") parrafo.salto = true; else parrafo.saltoCol = true; (parrafo.cortes = parrafo.cortes || []).push({ k: tramos.length, t: tipo === "page" ? "salto" : "saltocol" }); }
              else tramos.push({ x: "\u2028", a });
            } else if (cn === "drawing" || cn === "pict") {
              // (una forma con relleno de imagen o un grupo con una foto también traen «blip»: solo es imagen si el dibujo lo es)
              const gd0 = c.getElementsByTagNameNS("*", "graphicData")[0], uri0 = (gd0 && gd0.getAttribute("uri")) || "";
              const esDibujo = /wordprocessingShape|wordprocessingGroup|wordprocessingCanvas|chart|diagram/.test(uri0);
              const blip = esDibujo ? null : c.getElementsByTagNameNS("*", "blip")[0], img = esDibujo ? null : c.getElementsByTagNameNS("*", "imagedata")[0];
              const rid = blip ? (blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed") || blip.getAttribute("r:embed")) : img && (img.getAttribute("r:id") || img.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"));
              const src = rid ? await imagenDe(rid) : null;
              const ext = c.getElementsByTagNameNS("*", "extent")[0];
              const w = ext ? Math.round(+ext.getAttribute("cx") / EMU) : 320, h = ext ? Math.round(+ext.getAttribute("cy") / EMU) : 0;
              // (con su texto alternativo: «descr» —o «title»— del docPr, como lo pone Word)
              const dp = c.getElementsByTagNameNS("*", "docPr")[0], alt = dp ? (dp.getAttribute("descr") || dp.getAttribute("title") || "").trim() : "";
              if (src) tramos.push({ e: Object.assign({ tipo: "imagen", src, w, h, oe: ordinal(c) }, alt ? { alt } : {}, comoFlota(c, w, h)), a: {} });
              else {
                // Forma, cuadro de texto, gráfico o SmartArt: aquí se ve como una marca y al exportar sale el original.
                const g = c.getElementsByTagNameNS("*", "graphicData")[0], uri = (g && g.getAttribute("uri")) || "";
                const que = /chart/.test(uri) ? "Gráfico" : /diagram/.test(uri) ? "SmartArt" : /wordprocessingGroup/.test(uri) ? "Grupo" : /wordprocessingCanvas/.test(uri) ? "Lienzo" : c.getElementsByTagNameNS("*", "txbxContent")[0] ? "Cuadro de texto" : /picture/.test(uri) || img ? "Imagen" : "Forma";
                // un gráfico se dibuja como Word (js/graficas-office.js), del tamaño que tiene en el documento;
                // una forma o un cuadro de texto, con su geometría de Office, su relleno, su contorno y su texto;
                // un grupo o un lienzo, con todas sus piezas en su lugar
                const gOf = /chart/.test(uri) ? await graficaWord(c) : null;
                if (gOf) Object.assign(gOf, lugarDe(c, gOf.w, gOf.h));   // (un gráfico anclado va en su lugar de la hoja)
                const fOf = !gOf && /wordprocessingShape/.test(uri) ? await formaWord(c, w, h) : !gOf && /wordprocessingGroup|wordprocessingCanvas/.test(uri) ? await grupoWord(c, w, h) : null;
                tramos.push({ e: Object.assign({ tipo: "objeto", que, oe: ordinal(c), txt: (await textoObjeto(c)) || undefined }, gOf || {}, fOf || {}), a: {} }); aproximado = true;
              }
            } else if (cn === "object" && c.getElementsByTagNameNS("*", "imagedata")[0]) {
              // Un objeto incrustado (ecuación del Editor de ecuaciones o de MathType, hoja de Excel, diapositiva…): Word
              // muestra su vista previa (casi siempre un WMF o EMF); al descargar sale el objeto original completo.
              const im = c.getElementsByTagNameNS("*", "imagedata")[0], forma = im.parentNode;
              const rid = im.getAttribute("r:id") || im.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
              const src = rid ? await imagenDe(rid) : null;
              const est = (forma && forma.getAttribute && forma.getAttribute("style")) || "";
              const med = (k) => { const m = new RegExp("(?:^|;)\\s*" + k + "\\s*:\\s*([\\d.]+)(pt|in|cm|mm|px)?", "i").exec(est); if (!m) return 0; const f = { pt: 96 / 72, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, px: 1 }[(m[2] || "px").toLowerCase()]; return +m[1] * f; };
              const w = med("width"), h = med("height");
              const que = /Equation|MathType/i.test((c.getElementsByTagNameNS("*", "OLEObject")[0] || {}).getAttribute ? c.getElementsByTagNameNS("*", "OLEObject")[0].getAttribute("ProgID") || "" : "") ? "Ecuación" : "Objeto incrustado";
              if (src && w > 0 && h > 0) tramos.push({ e: { tipo: "objeto", que, oe: ordinal(c), img: src, w: Math.round(w * 100) / 100, h: Math.round(h * 100) / 100 }, a: {} });
              else { tramos.push({ e: { tipo: "objeto", que, oe: ordinal(c) }, a: {} }); aproximado = true; }
            } else if (cn === "AlternateContent" || cn === "object") {
              const que = cn === "object" ? "Objeto incrustado" : c.getElementsByTagNameNS("*", "txbxContent")[0] ? "Cuadro de texto" : /drawingml\/2006\/chart/.test(c.innerHTML || "") ? "Gráfico" : c.getElementsByTagNameNS("*", "relIds")[0] ? "SmartArt" : "Forma";
              const gOf = que === "Gráfico" ? await graficaWord(c) : null;
              if (gOf) Object.assign(gOf, lugarDe(c, gOf.w, gOf.h));
              // una forma o un cuadro de texto de Word 2010 en adelante: su dibujo va en el Choice (el Fallback es VML)
              const dibAC = !gOf && cn === "AlternateContent" ? c.getElementsByTagNameNS("*", "drawing")[0] : null;
              const extAC = dibAC && dibAC.getElementsByTagNameNS("*", "extent")[0];
              const gdAC = dibAC && dibAC.getElementsByTagNameNS("*", "graphicData")[0], uriAC = (gdAC && gdAC.getAttribute("uri")) || "";
              const wAC = extAC ? Math.round(+extAC.getAttribute("cx") / EMU) : 0, hAC = extAC ? Math.round(+extAC.getAttribute("cy") / EMU) : 0;
              const fOf = !dibAC || !extAC ? null : /wordprocessingGroup|wordprocessingCanvas/.test(uriAC) ? await grupoWord(dibAC, wAC, hAC) : dibAC.getElementsByTagNameNS("*", "wsp")[0] ? await formaWord(dibAC, wAC, hAC) : null;
              tramos.push({ e: Object.assign({ tipo: "objeto", que, oe: ordinal(c), txt: (await textoObjeto(c)) || undefined }, gOf || {}, fOf || {}), a: {} }); aproximado = true;
            } else if (cn === "footnoteReference") {
              const id = c.getAttributeNS(W, "id") || c.getAttribute("w:id");
              // (con marca propia —«*», †— Word no le pone número ni la cuenta: la marca es el texto que sigue)
              const propia = ["1", "true", "on"].includes(c.getAttributeNS(W, "customMarkFollows") || c.getAttribute("w:customMarkFollows") || "");
              tramos.push({ e: { tipo: "nota", texto: notasPie[id] || "", oe: ordinal(c), propia: propia || undefined }, a: {} });
            } else if (cn === "endnoteReference") {
              const id = c.getAttributeNS(W, "id") || c.getAttribute("w:id");
              tramos.push({ e: { tipo: "nota", final: true, texto: notasFin[id] || "", oe: ordinal(c) }, a: {} });
            } else if (cn === "sym") tramos.push({ x: textoSym(c.getAttributeNS(W, "font") || c.getAttribute("w:font"), c.getAttributeNS(W, "char") || c.getAttribute("w:char")), a });
            else if (cn === "fldChar" && parrafo) {
              // Campo de Word: { código | último valor }. Se guarda como campo (se ve su último valor, como en Word, y
              // «Actualizar campos» lo calcula otra vez); los que NOM no calcula se quedan como texto.
              const tipoF = c.getAttributeNS(W, "fldCharType") || c.getAttribute("w:fldCharType");
              const pila = (parrafo._campos = parrafo._campos || []);
              // (el idioma del campo es el de su texto: con él se escriben los meses y los números con letra)
              const lg = rpr && hijo(rpr, "lang");
              if (tipoF === "begin") pila.push({ instr: "", etapa: "instr", ini: tramos.length, a, idioma: (lg && (lg.getAttributeNS(W, "val") || lg.getAttribute("w:val"))) || null });
              else if (tipoF === "separate" && pila.length) { const f = pila[pila.length - 1]; f.etapa = "res"; f.res0 = tramos.length; }
              else if (tipoF === "end" && pila.length) cerrarCampo(pila, tramos);
            } else if (cn === "instrText" && parrafo && parrafo._campos && parrafo._campos.length) {
              const f = parrafo._campos[parrafo._campos.length - 1];
              if (f.etapa === "instr") f.instr += c.textContent;
            }
          }
        } else if (ln === "hyperlink") {
          const rid = r.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || r.getAttribute("r:id");
          const destino = rid && rels[rid];
          await runs(r, Object.assign({}, base, destino && /^(https?:|mailto:)/i.test(destino) ? { enlace: destino } : {}), tramos, parrafo);
        } else if (ln === "oMath" || ln === "oMathPara") {
          const t = ommlATex(r).trim();
          if (t) tramos.push({ e: { tipo: "ecuacion", tex: t, bloque: ln === "oMathPara" || null, oe: ordinal(r) }, a: {} });
        } else if (ln === "fldSimple" && parrafo) {
          const ini = tramos.length, pila = (parrafo._campos = parrafo._campos || []);
          pila.push({ instr: r.getAttributeNS(W, "instr") || r.getAttribute("w:instr") || "", etapa: "res", ini, res0: ini, a: base });
          await runs(r, base, tramos, parrafo);
          cerrarCampo(pila, tramos);
        } else if (ln === "bookmarkStart") {
          const nm = r.getAttributeNS(W, "name") || r.getAttribute("w:name"), id = r.getAttributeNS(W, "id") || r.getAttribute("w:id");
          if (nm && marcadoresUsados.has(nm)) { nombreMarcador[id] = nm; tramos.push({ e: { tipo: "marcador", nombre: nm, lado: "ini" }, a: {} }); }
        } else if (ln === "bookmarkEnd") {
          const nm = nombreMarcador[r.getAttributeNS(W, "id") || r.getAttribute("w:id")];
          if (nm) tramos.push({ e: { tipo: "marcador", nombre: nm, lado: "fin" }, a: {} });
        } else if (ln === "ins" || ln === "moveTo") await runs(r, Object.assign({}, base, { sug: autorCambio(r) }), tramos, parrafo);
        else if (ln === "del" || ln === "moveFrom") await runs(r, Object.assign({}, base, { sugBorrar: autorCambio(r) }), tramos, parrafo);
        else if (["smartTag", "fldSimple", "customXml", "sdt", "sdtContent", "bdo", "dir"].includes(ln)) await runs(r, base, tramos, parrafo);
      }
    }
    /**
     * Quién hizo un cambio marcado de Word (w:ins, w:del…): como el autor de una sugerencia de NOM ({ u, n, c }) y su fecha
     * (t), para que salga igual al descargar. El color sale del nombre (siempre el mismo para la misma persona).
     */
    function autorCambio(el) {
      const n = el.getAttributeNS(W, "author") || el.getAttribute("w:author") || "Autor", t = el.getAttributeNS(W, "date") || el.getAttribute("w:date") || "";
      let k = 0;
      for (const ch of n) k = (k * 31 + ch.codePointAt(0)) >>> 0;
      return Object.assign({ u: "word:" + n, n, c: COLORES_CAMBIO[k % COLORES_CAMBIO.length] }, t ? { t } : {});
    }
    /** Los campos que NOM sabe calcular (los demás —índice, vínculos, XE…— se quedan como su texto). */
    const CAMPOS_VIVOS = new Set(["DATE", "TIME", "CREATEDATE", "SAVEDATE", "PRINTDATE", "PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION", "SEQ", "REF", "PAGEREF", "NOTEREF",
      "IF", "=", "MERGEFIELD", "AUTHOR", "TITLE", "SUBJECT", "KEYWORDS", "COMMENTS", "FILENAME", "NUMWORDS", "NUMCHARS", "DOCPROPERTY", "LASTSAVEDBY", "QUOTE", "USERNAME", "USERINITIALS", "STYLEREF"]);
    /** Termina el campo de arriba de la pila: su último valor (los tramos de su resultado) se vuelve un campo. */
    function cerrarCampo(pila, tramos) {
      const f = pila.pop();
      const padre = pila[pila.length - 1];
      const r0 = f.res0 != null ? f.res0 : tramos.length;
      const resTr = tramos.slice(r0);
      const texto = resTr.every((t) => !t.e) ? resTr.map((t) => t.x || "").join("") : null;
      if (padre && padre.etapa === "instr") { padre.instr += texto || ""; tramos.length = f.ini; return; }   // dentro del código del de afuera
      const codigo = f.instr.replace(/\s+/g, " ").trim();
      const nombre = (/^=/.test(codigo) ? "=" : (codigo.split(" ")[0] || "")).toUpperCase();
      if (!CAMPOS_VIVOS.has(nombre) || texto == null || tramos.slice(f.ini, r0).some((t) => t.e)) return;   // se queda como su texto
      tramos.length = f.ini;
      tramos.push({ e: { tipo: "campo", instr: codigo, res: texto, a: Object.assign({}, (resTr[0] && resTr[0].a) || f.a || {}), idioma: f.idioma || undefined, orig: true }, a: {} });
    }
    const ESTILO_POR_NOMBRE = (nm) => (/^heading 1$|^título 1$/.test(nm) ? "h1" : /^heading 2$|^título 2$/.test(nm) ? "h2" : /^heading [3-9]$|^título [3-9]$/.test(nm) ? "h3"
      : nm === "title" || nm === "título" ? "titulo" : nm === "subtitle" || nm === "subtítulo" ? "subtitulo" : /quote|cita/.test(nm) ? "cita" : null);
    // ---- el idioma del texto (w:lang): con él parte Word las palabras al final del renglón (js/guiones.js)
    const langDe = (rpr) => { const l = rpr && hijo(rpr, "lang"); return l ? (l.getAttributeNS(W, "val") || l.getAttribute("w:val") || null) : null; };
    const langEstilo = (id, prof) => { const d = id != null ? defEstilo[id] : null; if (!d || (prof || 0) > 20) return null; return langDe(d.rpr) || langEstilo(d.base, (prof || 0) + 1); };
    const idiomaDoc = langDe(rprDef) || "es-MX";
    /** El idioma de la mayoría de las letras de un párrafo (de sus corridas, de su estilo o del documento). */
    function idiomaParrafo(p, ppr) {
      const idSt = val(hijo(ppr, "pStyle")) != null && defEstilo[val(hijo(ppr, "pStyle"))] ? val(hijo(ppr, "pStyle")) : idNormal;
      const base = langEstilo(idSt) || idiomaDoc;
      const cuenta = new Map();
      for (const t of p.getElementsByTagNameNS(W, "t")) {
        let q = t.parentNode;
        while (q && q.localName !== "p") q = q.parentNode;
        if (q !== p) continue;   // (lo de un cuadro de texto dentro del párrafo no cuenta)
        const rpr = t.parentNode && hijo(t.parentNode, "rPr");
        const lg = langDe(rpr) || langEstilo(val(hijo(rpr, "rStyle"))) || base;
        cuenta.set(lg, (cuenta.get(lg) || 0) + t.textContent.length);
      }
      let mejor = base, n = 0;
      for (const [k, v] of cuenta) if (v > n) { mejor = k; n = v; }
      return mejor;
    }
    /** Formato del párrafo y la letra que heredan sus corridas: { pa, letra }. */
    function attrsParrafo(ppr) {
      const pa = {}, idSt = val(hijo(ppr, "pStyle")) != null && defEstilo[val(hijo(ppr, "pStyle"))] ? val(hijo(ppr, "pStyle")) : idNormal;
      const c = parCrudo(ppr, parDe(idSt)), letra = letraDe(idSt);
      // estilo: por el nombre del suyo, o del primero de sus estilos base que se reconozca
      let st = "";
      for (let id = idSt, n = 0; id != null && defEstilo[id] && n < 20; id = defEstilo[id].base, n++) {
        if (!n) st = defEstilo[id].nombre;
        const e = ESTILO_POR_NOMBRE(defEstilo[id].nombre);
        if (e) { pa.estilo = e; break; }
      }
      // un estilo propio del Word (hecho por alguien en Word, o en NOM): su párrafo lo conserva con su nombre
      if (!pa.estilo && defEstilo[idSt] && defEstilo[idSt].propio) { pa.estilo = "u:" + idSt; estilosW["u:" + idSt] = { n: defEstilo[idSt].propio, id: idSt, deWord: 1 }; }
      if (c.sinGuiones) pa.sinGuiones = true;
      if (c.sinNumLin) pa.sinNumLin = true;
      if (c.tabs && c.tabs.length) pa.tabs = c.tabs;
      // la marca de párrafo (¶) con su propia letra y tamaño: de ella sale el alto del renglón donde está
      const mr = hijo(ppr, "rPr");
      if (mr) { const am = attrsRun(mr, {}); if (+am.tam) pa.tm = +am.tam; if (am.fuente) pa.fm = am.fuente; }
      if (c.num && c.num.id != null && c.num.id !== "0") {
        const lvl = +(c.num.lvl || 0), f = (formatoNum[c.num.id] || {})[lvl] || "bullet";
        // (un nivel con formato «none» no lleva marca: Word solo le da la sangría)
        if (f !== "none") pa.lista = f === "bullet" ? "vinetas" : "numeros";
        if (lvl && f !== "none") pa.nivel = Math.min(8, lvl);
        if (pa.lista === "numeros") pa.numid = String(c.num.id);   // Word cuenta por lista (numId) aunque haya texto en medio
        if (pa.lista === "numeros") {
          // el número como lo escribe Word: un formato sencillo («1.», «a)», «I.») o su texto completo con los niveles
          const txt = (textoNum[c.num.id] || {})[lvl], FM = { decimal: "1", lowerLetter: "a", upperLetter: "A", lowerRoman: "i", upperRoman: "I", decimalZero: "01" };
          const m = txt != null ? new RegExp("^%" + (lvl + 1) + "([.)])$").exec(txt) : null;
          if (m && FM[f] && FM[f] !== "01") { const nf = FM[f] + m[1]; if (nf !== "1.") pa.numfmt = nf; }
          else if (txt != null && /%[1-9]/.test(txt)) {
            pa.numtxt = txt;
            pa.numfmts = Array.from({ length: lvl + 1 }, (_, k) => FM[(formatoNum[c.num.id] || {})[k]] || "1").join(",");
          }
        }
      } else if (/^list bullet|^lista con viñetas/.test(st)) pa.lista = "vinetas";
      else if (/^list number|^lista con números/.test(st)) pa.lista = "numeros";
      if (pa.lista) delete pa.estilo;
      const v = { center: "centro", right: "der", end: "der", both: "just", distribute: "just" }[c.jc];
      if (v) pa.alin = v;
      if (!pa.lista) {
        const izq = c.izq || 0, der = c.der || 0, pri = c.pri || 0, fr = c.fr || 0, priBase = pa.estilo ? 0 : Math.max(0, pN.pri || 0);
        // (también negativas: en Word el texto sale al margen)
        if (fr > 0) { pa.francesa = cmDe(fr); if (Math.abs(izq - fr) > 20) pa.sangiz = cmDe(izq - fr); if (priBase) pa.primera = 0; }
        else { if (Math.abs(izq) > 20) pa.sangiz = cmDe(izq); if (pri !== priBase && pri >= 0) pa.primera = cmDe(pri); }
        if (Math.abs(der) > 20) pa.sangder = cmDe(der);
      }
      const li = liDeO(c);
      if (li != null) { if (Math.abs(li - liBase) > 0.0005 && !(formato.interlineado == null && liDeWord(li))) pa.interlineado = li; }
      else if (c.linea && c.regla === "exact") pa.lineaPt = Math.round(c.linea / 2) / 10;
      else if (c.linea && c.regla === "atLeast" && c.linea / 20 > (letra.tam || formato.tam || 11) * 1.3) pa.lineaPt = Math.round(c.linea / 2) / 10;
      // espacio antes y después: solo se anota si no es el que aquí ya tiene ese tipo de párrafo. El «sin espacio entre
      // párrafos del mismo estilo» (contextualSpacing) depende de los vecinos: se resuelve al final (espaciosContextuales)
      const antes = Math.round((c.antes || 0) / 2) / 10, desp = Math.round((c.desp || 0) / 2) / 10;
      const om = pa.lista ? [0, 2] : pa.estilo ? { h1: [12, 0], h2: [2, 0], h3: [2, 0], titulo: [0, 4], cita: [10, 8] }[pa.estilo] : [0, 8];
      Object.defineProperty(pa, "_esp", { value: { antes, desp, om, st: idSt, ctx: !!c.junto }, enumerable: false, configurable: true });
      if (!om || antes !== om[0]) pa.antes = antes;
      if (!om || desp !== om[1]) pa.despues = desp;
      // entre páginas: salto antes, con el siguiente (los títulos ya lo llevan aquí), líneas juntas y viudas
      if (c.pbb) pa.saltoAntes = true;
      // letra capital (marco de Word con dropCap): «drop» (dentro del texto) o «margin» (en el margen), cuántos renglones
      // abarca y la distancia al texto (pt). El párrafo siguiente la rodea.
      const fp = hijo(ppr, "framePr"), dc = fp ? at(fp, "dropCap") : null;
      if (dc === "drop" || dc === "margin") pa.capital = `${dc}|${Math.max(1, Math.min(10, +(at(fp, "lines") || 3) || 3))}|${Math.round(((+(at(fp, "hSpace") || 0) || 0) / 20) * 10) / 10}`;
      const kn0 = /^h[123]$/.test(pa.estilo || "");
      if (!!c.kn !== kn0) pa.conSig = c.kn ? 1 : 0;
      if (c.kl) pa.juntas = true;
      if (c.wc === false) pa.viudas = 0;
      if (c.bdr) pa.borde = bordeDe(c.bdr);
      if (c.shd) pa.sombra = c.shd;
      if (c.rtl) pa.dir = "rtl";
      Object.keys(pa).forEach((k) => pa[k] == null && delete pa[k]);
      return { pa, letra, crudo: c };
    }
    // ---- secciones: cada una puede traer otra hoja y otro número de columnas
    const TAMS = { carta: [816, 1056], a4: [794, 1123], oficio: [816, 1344], a5: [559, 794], media: [529, 816] };
    function datosSeccion(sect) {
      const sz = hijo(sect, "pgSz"), co = hijo(sect, "cols"), mg = hijo(sect, "pgMar"), g = (n, k) => +(at(n, k) || 0) || 0;
      const w = sz ? g(sz, "w") / TWIP : 816, h = sz ? g(sz, "h") / TWIP : 1056;
      const colsXml = co ? hijos(co, "col") : [];
      const n = co ? Math.max(1, +(at(co, "num") || 0) || colsXml.length || 1) : 1;
      // columnas: iguales (con su espacio, 720 twips si no lo dice) o cada una con su ancho y el espacio que la sigue
      const igual = !co || at(co, "equalWidth") !== "0" || !colsXml.length;
      const esp = co && at(co, "space") != null ? g(co, "space") / TWIP : 48;
      const anchos = !igual ? colsXml.map((c) => [Math.round((g(c, "w") / TWIP) * 10) / 10, Math.round((g(c, "space") / TWIP) * 10) / 10]) : null;
      // márgenes (el de encuadernación se suma a la izquierda) y alineación vertical de la página
      const m = mg ? [Math.abs(g(mg, "top")), Math.abs(g(mg, "bottom")), g(mg, "left") + g(mg, "gutter"), g(mg, "right")].map((x) => Math.round((x / TWIP) * 10) / 10) : [96, 96, 96, 96];
      const pn = hijo(sect, "pgNumType");
      return { sect, w, h, hor: w > h, cols: n, tipo: val(hijo(sect, "type")) || "nextPage", m, igual, esp, anchos, sep: !!co && /^(1|true|on)$/.test(at(co, "sep") || ""),
        va: val(hijo(sect, "vAlign")) || "top", numIni: pn && at(pn, "start") != null ? +at(pn, "start") : null, numFmt: pn ? at(pn, "fmt") || null : null, ln: numLineasDe(hijo(sect, "lnNumType")) };
    }
    /**
     * Números de línea de Word (w:lnNumType): [cada cuántos se ve, con cuál empieza, distancia al texto en px (null =
     * automática), reinicio: "p" cada página, "s" cada sección, "c" continuo]. Word guarda el inicio menos 1 (medido: con
     * «empezar en 3» escribe w:start="2"). Sin countBy no se ve nada.
     */
    function numLineasDe(n) {
      if (!n) return null;
      const cada = Math.round(+(at(n, "countBy") || 0) || 0);
      if (cada < 1) return null;
      const dist = at(n, "distance");
      const r = at(n, "restart");
      return [Math.min(100, cada), Math.max(1, Math.min(32767, (+(at(n, "start") || 0) || 0) + 1)), dist != null && dist !== "" && +dist > 0 ? Math.round((+dist / TWIP) * 10) / 10 : null, r === "continuous" ? "c" : r === "newSection" ? "s" : "p"];
    }
    /** La hoja de una sección como la guarda el párrafo que la empieza (pa.sx): tamaño, márgenes, columnas, tipo de
        salto, alineación vertical y número de página con que empieza (si Word lo reinicia). */
    const sxDe = (s) => JSON.stringify({ w: Math.round(s.w * 10) / 10, h: Math.round(s.h * 10) / 10, m: s.m, c: s.cols, e: Math.round(s.esp * 10) / 10, a: s.anchos || undefined,
      l: s.sep ? 1 : undefined, t: s.tipo === "continuous" ? "c" : s.tipo === "oddPage" ? "i" : s.tipo === "evenPage" ? "p" : "n", v: s.va !== "top" ? s.va : undefined, n: s.numIni != null ? s.numIni : undefined,
      f: s.numFmt && s.numFmt !== "decimal" ? s.numFmt : undefined, ln: s.ln || undefined });
    const secciones = [];
    hijos(cuerpo).forEach((c) => {
      const sp = c.localName === "p" ? hijo(hijo(c, "pPr"), "sectPr") : c.localName === "sectPr" ? c : null;
      if (sp) secciones.push(datosSeccion(sp));
    });
    let nSec = 0, secPendiente = secciones.length && secciones[0].cols > 1 ? "c|" + secciones[0].cols + " columnas" : null;
    // Con secciones de otra hoja, con columnas o centradas: cada sección lleva su hoja en su primer párrafo (pa.sx)
    const geoSec = (s) => [Math.round(s.w), Math.round(s.h), s.m.join(","), s.cols, Math.round(s.esp), JSON.stringify(s.anchos), s.sep, s.va, JSON.stringify(s.ln)].join("|");
    const conSecciones = secciones.length > 0 && (new Set(secciones.map(geoSec)).size > 1 || secciones.some((s) => s.cols > 1 || s.va !== "top"));
    let sxPendiente = conSecciones ? sxDe(secciones[0]) : null, sxK = 0;
    const paSx = [];   // [pa, índice de la sección]
    /** Marca con la que empieza una sección: "p|…" si arranca en hoja nueva, "c|…" si sigue en la misma. */
    function marcaSeccion(nueva, ant) {
      const partes = [];
      if (nueva.cols !== ant.cols) partes.push(nueva.cols === 1 ? "1 columna" : nueva.cols + " columnas");
      if (Math.abs(nueva.w - ant.w) > 8 || Math.abs(nueva.h - ant.h) > 8) partes.push(`hoja ${nueva.hor ? "horizontal" : "vertical"} de ${Math.round((nueva.w / 37.7953) * 10) / 10} × ${Math.round((nueva.h / 37.7953) * 10) / 10} cm`);
      const nuevaHoja = nueva.tipo !== "continuous";
      return partes.length || nuevaHoja ? (nuevaHoja ? "p|" : "c|") + partes.join(" · ") : null;
    }
    /** Una tabla de Word: { filas, cols, celdas } (los campos del bloque «tabla»; ver modo-texto.js). prof: anidada dentro de otra. */
    async function leerTabla(c, prof) {
      const filas = hijos(c, "tr"), celdas = {};
      let cols = 0;
      // texto de una celda: sus párrafos, y si trae otra tabla adentro, sus renglones con " | "
      const textoCelda = (tc) => {
        const out = [];
        const rec = (n) => hijos(n).forEach((h) => {
          if (h.localName === "p") out.push([...h.getElementsByTagNameNS(W, "t")].filter((t) => { for (let q = t.parentNode; q && q !== h; q = q.parentNode) if (q.localName === "txbxContent" || q.localName === "Fallback") return false; return true; }).map((t) => t.textContent).join(""));
          else if (h.localName === "tbl") hijos(h, "tr").forEach((tr) => out.push(hijos(tr, "tc").map((x2) => textoCelda(x2).replace(/\n/g, " ")).join(" | ")));
          else if (h.localName === "sdt") rec(hijo(h, "sdtContent"));
        });
        rec(tc);
        return out.join("\n").trim();
      };
      const abiertas = {};   // columna → celda donde empezó una combinación vertical
      for (const [f, tr] of filas.entries()) {
        let k = 0;
        for (const tc of hijos(tr, "tc")) {
          const tcPr = hijo(tc, "tcPr");
          const span = Math.max(1, +(val(hijo(tcPr, "gridSpan")) || 1));
          const vm = hijo(tcPr, "vMerge"), vmv = vm ? val(vm) || "continue" : null;
          const ref = F().ref(k, f);
          if (vmv === "continue" && abiertas[k]) {
            const tl = abiertas[k], m = String(celdas["m:" + tl] || "1x1").split("x").map(Number);
            celdas["m:" + tl] = `${m[0]}x${m[1] + 1}`;
          } else {
            // todos los párrafos de la celda, también los que van dentro de un control de contenido
            const t = textoCelda(tc);
            if (t) celdas["c:" + ref] = t;
            // lo que el texto simple no guarda (imágenes, listas, tablas adentro, letra mezclada): para verla como Word
            // mientras nadie cambie su texto (con la huella del texto con que se importó)
            const rico = await ricoCelda(tc, prof);
            if (rico) celdas["h:" + ref] = JSON.stringify({ t: NOM.origen.huella(t || ""), b: rico });
            if (span > 1 || vmv === "restart") celdas["m:" + ref] = `${span}x1`;
            abiertas[k] = vmv === "restart" ? ref : null;
            const shd = hijo(tcPr, "shd"), fill = shd && (shd.getAttributeNS(W, "fill") || shd.getAttribute("w:fill"));
            if (fill && /^[0-9a-f]{6}$/i.test(fill) && fill.toUpperCase() !== "FFFFFF") celdas["r:" + ref] = "#" + fill.toUpperCase();
            const negritas = [...tc.getElementsByTagNameNS(W, "r")].filter((rr) => rr.textContent.trim());
            if (negritas.length && negritas.every((rr) => { const rp = hijo(rr, "rPr"); return rp && hijo(rp, "b") && val(hijo(rp, "b")) !== "0"; })) celdas["n:" + ref] = true;
            // cursiva y color si toda la celda los lleva; alineación, la de su primer párrafo
            if (negritas.length && negritas.every((rr) => { const rp = hijo(rr, "rPr"); return rp && activo(hijo(rp, "i")); })) celdas["i:" + ref] = true;
            const colores = new Set(negritas.map((rr) => { const k = val(hijo(hijo(rr, "rPr"), "color")); return k && /^[0-9a-f]{6}$/i.test(k) ? "#" + k.toUpperCase() : ""; }));
            if (colores.size === 1 && [...colores][0] && [...colores][0] !== "#000000") celdas["k:" + ref] = [...colores][0];
            const p0 = hijo(tc, "p"), jc = p0 ? parCrudo(hijo(p0, "pPr"), {}).jc : null, al = { center: "centro", right: "der", end: "der" }[jc];
            if (al) celdas["al:" + ref] = al;
            // alineación vertical y texto girado (hacia arriba o hacia abajo), como en Word
            const vAl = val(hijo(tcPr, "vAlign"));
            if (vAl === "center" || vAl === "both") celdas["va:" + ref] = "c"; else if (vAl === "bottom") celdas["va:" + ref] = "b";
            const tDir = val(hijo(tcPr, "textDirection"));
            if (/^(btLr|tbRl|tbRlV|tbLrV)$/.test(tDir || "")) celdas["td:" + ref] = /^bt/.test(tDir) ? "arriba" : "abajo";
          }
          k += span;
        }
        cols = Math.max(cols, k);
      }
      Object.keys(celdas).forEach((x) => { if (x.startsWith("m:") && celdas[x] === "1x1") delete celdas[x]; });
      // Filas como Word: alto exacto («e») o mínimo («m») en px, «permitir que la fila se divida entre páginas»
      // apagado (fp) y las filas de arriba que se repiten en cada hoja («repetir como fila de encabezado»)
      {
        let cab = 0, seguidas = true;
        filas.forEach((tr, f) => {
          const trPr = hijo(tr, "trPr"), th = hijo(trPr, "trHeight"), v = th ? +(val(th) || 0) : 0;
          if (v > 0 && at(th, "hRule") !== "auto") celdas["fh:" + f] = (at(th, "hRule") === "exact" ? "e" : "m") + Math.round((v / 15) * 100) / 100;
          if (activo(hijo(trPr, "cantSplit"))) celdas["fp:" + f] = 1;
          if (seguidas && activo(hijo(trPr, "tblHeader"))) cab++; else seguidas = false;
        });
        if (cab) celdas.cabRep = cab;
      }
      // Márgenes de celda y párrafos de las celdas: lo del documento, luego el estilo de tabla, luego el estilo del párrafo
      // y lo propio (así los pone Word: con «Tabla con cuadrícula», renglón sencillo y sin espacio después; sin estilo,
      // los 8 pt y el 1.08 del documento). Word deja 0 arriba y abajo y 0.19 cm a los lados.
      try {
        const marDe = (m) => {
          if (!m) return null;
          const g = (l, l2) => { const q = hijo(m, l) || (l2 ? hijo(m, l2) : null); if (!q) return null; const tp = at(q, "type"); return tp === "nil" ? 0 : !tp || tp === "dxa" ? (+(at(q, "w") || 0) || 0) / 15 : null; };
          return [g("top"), g("right", "end"), g("bottom"), g("left", "start")];
        };
        const deEstilo = (id, prof) => {
          const d = id != null ? defEstilo[id] : null;
          if (!d || d.tipo !== "table" || (prof || 0) > 20) return { ppr: pBase, rpr: rBase, mar: [0, 7.2, 0, 7.2] };
          const b = deEstilo(d.base, (prof || 0) + 1), m = marDe(hijo(hijo(d.nodo, "tblPr"), "tblCellMar"));
          return { ppr: parCrudo(d.ppr, b.ppr), rpr: attrsRun(d.rpr, b.rpr), mar: m ? m.map((x, k) => (x != null ? x : b.mar[k])) : b.mar };
        };
        const est = deEstilo(val(hijo(hijo(c, "tblPr"), "tblStyle")) || idTablaNormal);
        const mt = marDe(hijo(hijo(c, "tblPr"), "tblCellMar")), mar = mt ? mt.map((x, k) => (x != null ? x : est.mar[k])) : est.mar;
        celdas.cmar = mar.map((x) => Math.round(x * 100) / 100).join(",");
        // el estilo del párrafo (con su base) va encima de lo de la tabla
        const sobre = (id, base, prof) => { const d = id != null ? defEstilo[id] : null; return !d || (prof || 0) > 20 ? base : parCrudo(d.ppr, d.base != null ? sobre(d.base, base, (prof || 0) + 1) : base); };
        const p0 = (() => { const tc0 = filas.length ? hijo(filas[0], "tc") : null; return tc0 ? hijo(tc0, "p") : null; })();
        if (p0) {
          const ppr0 = hijo(p0, "pPr"), pc = parCrudo(ppr0, sobre(val(hijo(ppr0, "pStyle")) || idNormal, est.ppr));
          if (pc.linea && (!pc.regla || pc.regla === "auto")) celdas.cli = Math.round((pc.linea / 240) * 10000) / 10000;
          else if (pc.linea && pc.regla === "exact") celdas.clx = Math.round((pc.linea / 15) * 100) / 100;
          else celdas.cli = 1;
          celdas.cesp = `${Math.round(((pc.antes || 0) / 15) * 100) / 100},${Math.round(((pc.desp || 0) / 15) * 100) / 100}`;
        }
        // tamaño de letra de cada celda si todo su texto lo lleva y no es el del documento
        const tamDoc = rN.tam || 10;
        filas.forEach((tr, f) => {
          let k = 0;
          hijos(tr, "tc").forEach((tc) => {
            const span = Math.max(1, +(val(hijo(hijo(tc, "tcPr"), "gridSpan")) || 1)), ref = F().ref(k, f);
            k += span;
            const tams = new Set();
            hijos(tc, "p").forEach((pp) => {
              const base = letraDe(val(hijo(hijo(pp, "pPr"), "pStyle")) || idNormal);
              hijos(pp, "r").forEach((rr) => { if (rr.textContent.trim()) tams.add(attrsRun(hijo(rr, "rPr"), base).tam || tamDoc); });
            });
            if (tams.size === 1 && [...tams][0] !== tamDoc && celdas["c:" + ref] != null) celdas["z:" + ref] = [...tams][0];
          });
        });
      } catch (e) { /* la tabla sin sus medidas de Word */ }
      // anchos de columna (en % del ancho de la tabla) y si la tabla no lleva líneas
      const rejilla = hijos(hijo(c, "tblGrid"), "gridCol").map((gc) => +(at(gc, "w") || 0) || 0), totalRej = rejilla.reduce((a2, b2) => a2 + b2, 0);
      if (totalRej > 0 && rejilla.length === cols && cols <= 63 && new Set(rejilla).size > 1) rejilla.forEach((w2, k2) => { celdas["aw:" + F().colLetra(k2)] = Math.round((w2 / totalRej) * 10000) / 100; });
      // Estilo de tabla de Word (Tabla con cuadrícula 4 – Énfasis 2…): por capas como Word —toda la tabla, franjas,
      // primera y última fila y columna, esquinas— según lo que pide la tabla (tblLook). Lo propio de cada celda gana.
      try { estiloTablaWord(c, celdas, filas.length, Math.min(63, cols)); } catch (e) { /* la tabla sin su estilo */ }
      try { bordesDirectosWord(c, celdas, filas, filas.length, Math.min(63, cols)); } catch (e) { /* sin los bordes propios */ }
      // ancho de la tabla (de Word: en px o en % del texto; sin él, la suma de sus columnas), alineación y sangría
      {
        const tp = hijo(c, "tblPr"), tw = hijo(tp, "tblW"), tipo = at(tw, "type"), v = +(at(tw, "w") || 0) || 0;
        if (tipo === "pct" && v > 0) celdas.tw = "p:" + Math.round((String(at(tw, "w")).endsWith("%") ? parseFloat(at(tw, "w")) : v / 50) * 100) / 100;
        else if (tipo === "dxa" && v > 0) celdas.tw = Math.round((v / 15) * 100) / 100;
        else if (totalRej > 0) celdas.tw = Math.round((totalRej / 15) * 100) / 100;
        const tj = val(hijo(tp, "jc"));
        if (tj === "center") celdas.tj = "c"; else if (tj === "right" || tj === "end") celdas.tj = "d";
        const ti = hijo(tp, "tblInd"), vi = ti && (at(ti, "type") || "dxa") === "dxa" ? +(at(ti, "w") || 0) || 0 : 0;
        if (vi) celdas.ti = Math.round((vi / 15) * 100) / 100;
      }
      const tb = hijo(hijo(c, "tblPr"), "tblBorders");
      const sinLineas = (b2) => !!b2 && ["top", "left", "bottom", "right", "insideH", "insideV"].every((l) => { const q = hijo(b2, l); return !q || ["nil", "none"].includes(val(q)); });
      if (tb ? sinLineas(tb) : !val(hijo(hijo(c, "tblPr"), "tblStyle")) && !c.getElementsByTagNameNS(W, "tcBorders").length) celdas.bordes = "ninguno";
      // de qué fila del original es cada celda y su firma (al descargar, las que no cambiaron salen tal cual)
      for (let f = 0; f < filas.length; f++) for (let k2 = 0; k2 < Math.min(63, cols); k2++) { const ref = F().ref(k2, f); celdas["o:" + ref] = f + "|" + firmaCelda(celdas, ref); }
      return filas.length ? { filas: filas.length, cols: Math.min(63, cols), celdas } : null;
    }
    /**
     * Lo de una celda de Word que el texto simple no guarda (imágenes, listas, tablas adentro, letra mezclada o alineaciones
     * distintas), para dibujarla como Word: [{ r: [{ x, a } | { img, w, h }], al, m (la marca de la lista), iz, fr } | { tabla }].
     * null si con su texto y el formato de la celda basta.
     */
    async function ricoCelda(tc, prof) {
      const out = [], cuenta = {}, firmas = new Set(), alins = new Set();
      let complejo = false;
      const romano = (n) => { let r = ""; [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]].forEach(([v, l]) => { while (n >= v) { r += l; n -= v; } }); return r; };
      const letra = (n) => { let r = ""; for (; n > 0; n = Math.floor((n - 1) / 26)) r = String.fromCharCode(97 + ((n - 1) % 26)) + r; return r; };
      const numero = (fmt, n) => ({ a: letra(n), A: letra(n).toUpperCase(), i: romano(n), I: romano(n).toUpperCase(), "01": String(n).padStart(2, "0") }[fmt] || String(n));
      async function rec(n) {
        for (const h of hijos(n)) {
          if (h.localName === "p") {
            const ppr = hijo(h, "pPr"), ap = attrsParrafo(ppr), tramos = [], info = {};
            await runs(h, ap.letra, tramos, info);
            const r = [];
            for (const t of tramos) {
              if (t.e && t.e.tipo === "imagen" && t.e.src) { r.push({ img: t.e.src, w: Math.round(+t.e.w || 0), h: Math.round(+t.e.h || 0) }); complejo = true; }
              // (un objeto incrustado —ecuación vieja, hoja…— se ve en la celda con su vista previa, como en Word)
              else if (t.e && t.e.tipo === "objeto" && t.e.img) { r.push({ img: t.e.img, w: Math.round(+t.e.w || 0), h: Math.round(+t.e.h || 0) }); complejo = true; }
              // (un campo dentro de una celda: su último valor, como texto)
              else if (t.e && t.e.tipo === "campo") { r.push({ x: String(t.e.res || ""), a: t.e.a || {} }); }
              else if (t.e && t.e.tipo === "marcador") { /* nada que ver */ }
              else if (t.e) { if (t.e.txt) r.push({ x: String(t.e.txt), a: {} }); }
              else if (t.x) { r.push({ x: t.x, a: t.a || {} }); if (t.x.trim()) firmas.add(JSON.stringify(t.a || {})); }
            }
            const b = { r }, pa = ap.pa, cr = ap.crudo || {};
            if (pa.alin) b.al = pa.alin;
            alins.add(pa.alin || "");
            if (pa.lista) {
              complejo = true;
              const nivel = +pa.nivel || 0, clave = (pa.numid || pa.lista) + "|";
              Object.keys(cuenta).forEach((x) => { const [c0, n0] = x.split("#"); if (c0 === clave && +n0 > nivel) delete cuenta[x]; });
              const k = clave + "#" + nivel;
              cuenta[k] = (cuenta[k] || 0) + 1;
              if (pa.lista === "vinetas") b.m = ["•", "o", "▪"][nivel % 3];
              else {
                const fm = String(pa.numfmt || "1."), m = /^(1|a|A|i|I|01)([.)])$/.exec(fm);
                b.m = pa.numtxt ? String(pa.numtxt).replace(/%([1-9])/g, (_, d) => numero(String(pa.numfmts || "").split(",")[+d - 1] || "1", cuenta[clave + "#" + (+d - 1)] || 1)) : m ? numero(m[1], cuenta[k]) + m[2] : cuenta[k] + ".";
              }
              // sangría de Word: a la izquierda y la francesa (la marca va en ella)
              b.iz = Math.round(((cr.izq != null ? cr.izq : 720 * (nivel + 1)) / 15) * 10) / 10;
              b.fr = Math.round(((cr.fr != null && cr.fr > 0 ? cr.fr : 360) / 15) * 10) / 10;
            }
            out.push(b);
          } else if (h.localName === "tbl" && (prof || 0) < 3) {
            const t = await leerTabla(h, (prof || 0) + 1);
            if (t) { out.push({ tabla: t }); complejo = true; }
          } else if (h.localName === "sdt") await rec(hijo(h, "sdtContent"));
        }
      }
      await rec(tc);
      if (firmas.size > 1 || alins.size > 1) complejo = true;
      return complejo ? out : null;
    }
    async function bloque(n, unidad) {
      const lista = hijos(n);
      for (let idx = 0; idx < lista.length; idx++) {
        const c = lista[idx];
        // ox = qué hijo del cuerpo es (para sacar esa unidad tal cual del original al exportar)
        // (dentro de un control de contenido: "3.0", "3.1"… = hijos del contenido del hijo 3)
        const ox = unidad != null ? unidad + "." + idx : idx;
        uEl = c;
        if (c.localName === "p") {
          const ppr = hijo(c, "pPr"), ap = attrsParrafo(ppr), pa = ap.pa, tramos = [], info = {};
          pa.ox = ox;
          // (la hoja de la primera sección va en el primer párrafo del documento aunque esté dentro de un control de
          // contenido, como una portada: si no, la sección empezaba más abajo y en hoja nueva)
          if (sxPendiente != null && (unidad == null || sxK === 0)) { pa.sx = sxPendiente; paSx.push([pa, sxK]); sxPendiente = null; }
          if (unidad == null) {
            if (secPendiente != null) { pa.sec = secPendiente; secPendiente = null; }
            if (hijo(ppr, "sectPr") && secciones[nSec + 1]) { secPendiente = marcaSeccion(secciones[nSec + 1], secciones[nSec]); if (conSecciones) { sxPendiente = sxDe(secciones[nSec + 1]); sxK = nSec + 1; } nSec++; }
          }
          await runs(c, ap.letra, tramos, info);
          // (el idioma del párrafo si no es el del documento: no se pinta, solo dice cómo partir las palabras)
          { const lg = idiomaParrafo(c, ppr); if (lg && lg !== idiomaDoc) pa.lg = lg; }
          if (info.cortes && info.cortes.length) {
            // Saltos de página o de columna dentro del párrafo: el texto se parte donde va cada salto (antes del texto,
            // en medio o al final). El espacio de antes es del primer pedazo y el de después, del último (como Word:
            // un salto al principio deja el espacio de antes en la hoja anterior y el texto empieza arriba de la nueva).
            const conTexto = (ts) => ts.some((t) => t.e || (t.x && t.x.replace(/[\u2028\s]/g, "")));
            let desde = 0, primero = true;
            const pedazos = [];
            for (const ct of info.cortes) { pedazos.push({ ts: tramos.slice(desde, ct.k) }, { salto: ct.t }); desde = ct.k; }
            pedazos.push({ ts: tramos.slice(desde) });
            const conTx = pedazos.map((q) => q.ts && conTexto(q.ts));
            const ultimo = conTx.lastIndexOf(true);
            pedazos.forEach((q, j) => {
              // (si el párrafo solo trae saltos, el primero se queda con su formato y con el inicio de sección)
              if (q.salto) { parrafos.push({ pa: ultimo < 0 && primero ? pa : { ox }, tramos: [{ e: { tipo: q.salto } }] }); if (ultimo < 0) primero = false; return; }
              if (!conTx[j]) return;
              const p2 = Object.assign({}, pa);
              if (!primero) { if (p2.antes != null || pa._esp) p2.antes = 0; delete p2.saltoAntes; delete p2.sec; delete p2.sx; }
              if (j !== ultimo) p2.despues = 0;
              if (pa._esp) Object.defineProperty(p2, "_esp", { value: Object.assign({}, pa._esp, primero ? {} : { antes: 0 }, j !== ultimo ? { desp: 0 } : {}), enumerable: false, configurable: true });
              if (primero && pa.sx) { const k = paSx.findIndex((x) => x[0] === pa); if (k >= 0) paSx[k][0] = p2; }
              parrafos.push({ pa: p2, tramos: q.ts });
              primero = false;
            });
            // (un párrafo que solo trae el salto: el salto y nada más)
          }
          else parrafos.push({ pa, tramos });
        } else if (c.localName === "tbl") {
          const t = await leerTabla(c, 0);
          // (una tabla de Word no deja espacio antes ni después: el del párrafo de abajo cuenta solo)
          if (t) parrafos.push({ pa: { ox, antes: 0, despues: 0 }, tramos: [{ e: { tipo: "tabla", filas: t.filas, cols: t.cols, celdas: t.celdas } }] });
        } else if (c.localName === "sdt") await bloque(hijo(c, "sdtContent"), ox);
      }
    }
    await bloque(cuerpo);
    // «Sin espacio entre párrafos del mismo estilo» (contextualSpacing) de Word: el párrafo que lo pide no deja su espacio
    // de antes si el de arriba es de su mismo estilo, ni el de después si el de abajo lo es (medido en Word real).
    for (let i = 0; i < parrafos.length; i++) {
      const pa = parrafos[i].pa, e = pa && pa._esp;
      if (!e) continue;
      const ant = i > 0 && parrafos[i - 1].pa && parrafos[i - 1].pa._esp, sig = i + 1 < parrafos.length && parrafos[i + 1].pa && parrafos[i + 1].pa._esp;
      let a = e.antes, d = e.desp;
      if (e.ctx && ant && ant.st === e.st) a = 0;
      if (e.ctx && sig && sig.st === e.st) d = 0;
      if (a !== e.antes || d !== e.desp) {
        if (!e.om || a !== e.om[0]) pa.antes = a; else delete pa.antes;
        if (!e.om || d !== e.om[1]) pa.despues = d; else delete pa.despues;
      }
    }
    // Página: la de la PRIMERA sección, que es la hoja con la que empieza el documento (la sección
    // del final del cuerpo es la última, y puede ser otra hoja).
    const sec0 = secciones[0] || null, sect = sec0 ? sec0.sect : null;
    let pagina = null;
    const avisos = [];
    if (sect) {
      const mg = hijo(sect, "pgMar"), g = (n, k) => +(at(n, k) || 0) || 0;
      const lado = Math.min(sec0.w, sec0.h), largo = Math.max(sec0.w, sec0.h);
      const tam = Object.keys(TAMS).find((k) => Math.abs(lado - TAMS[k][0]) < 8 && Math.abs(largo - TAMS[k][1]) < 10);
      pagina = { tam: tam || "otro", orient: sec0.hor ? "h" : "v", // (en Word un margen de arriba o de abajo negativo quiere decir "fijo": la medida es la misma)
        margen: mg ? { s: Math.round(Math.abs(g(mg, "top")) / TWIP), i: Math.round(Math.abs(g(mg, "bottom")) / TWIP), iz: Math.round(g(mg, "left") / TWIP), d: Math.round(g(mg, "right") / TWIP) } : undefined };
      if (!tam) { pagina.w = Math.round(lado); pagina.h = Math.round(largo); }   // tamaño que no es de los comunes: sus medidas (px, hoja vertical)
      // números de línea (si todas las secciones los tienen igual; si no, cada sección los trae en su pa.sx)
      if (sec0.ln && !conSecciones) pagina.numLin = sec0.ln;
      const pb = hijo(sect, "pgBorders");
      const lb = pb ? ladoBorde(hijo(pb, "top")) || ladoBorde(hijo(pb, "left")) || ladoBorde(hijo(pb, "bottom")) || ladoBorde(hijo(pb, "right")) : null;
      // borde de página: grosor (pt), estilo, color y a cuántos px del canto de la hoja (o del texto) va
      if (lb) pagina.borde = { a: lb[0], e: lb[1], c: lb[2], d: Math.round((lb[3] * 96) / 72), hoja: at(pb, "offsetFrom") === "page" || undefined };
      if (!conSecciones && secciones.some((x) => Math.abs(x.w - sec0.w) > 8 || Math.abs(x.h - sec0.h) > 8)) avisos.push("Este documento cambia de tamaño u orientación de hoja a media obra: aquí se ve todo con la hoja del principio (al descargar en Word sale cada sección con la suya)");
      if (!conSecciones && secciones.some((x) => x.cols > 1)) avisos.push("Las partes en columnas se ven aquí a una sola columna, con una marca donde empiezan (al descargar en Word salen en columnas)");
    }
    // Encabezado y pie de la primera sección: su texto (izquierda, centro y derecha van separados por
    // tabuladores) y dónde va el número de página.
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    async function zona(etiqueta, sectZ) {
      const deSec = sectZ || sect, refs = deSec ? hijos(deSec, etiqueta) : [];
      const ref = refs.find((x) => (at(x, "type") || "default") === "default");
      const rid = ref && (ref.getAttributeNS(R_NS, "id") || ref.getAttribute("r:id")), destino = rid && rels[rid];
      const xml = destino ? await z.texto(rutaDe(destino)) : null;
      if (!xml) return null;
      const huecos = ["", "", ""];
      [...P(xml).getElementsByTagNameNS(W, "p")].forEach((pp) => {
        for (let q = pp.parentNode; q; q = q.parentNode) if (q.localName === "txbxContent") return;
        let out = "", campo = null, instr = "", res = "";
        const cierra = (i2, r2) => { out += /NUMPAGES|SECTIONPAGES/i.test(i2) ? "\u0002" : /(^|\s)PAGE(\s|$)/i.test(i2) ? "\u0001" : r2; };
        [...pp.getElementsByTagNameNS(W, "*")].forEach((n) => {
          const ln = n.localName;
          if (ln === "fldChar") {
            const t = at(n, "fldCharType");
            if (t === "begin") { campo = "instr"; instr = ""; res = ""; } else if (t === "separate") campo = "res"; else if (t === "end") { cierra(instr, res); campo = null; }
          } else if (ln === "instrText") instr += n.textContent;
          else if (ln === "fldSimple") cierra(at(n, "instr") || "", "");
          else if (ln === "t") {
            let simple = false;
            for (let q = n.parentNode; q && q !== pp; q = q.parentNode) if (q.localName === "fldSimple" && /PAGE/i.test(at(q, "instr") || "")) simple = true;
            if (simple) return;
            if (campo === "res") res += n.textContent; else if (!campo) out += n.textContent;
          } else if (ln === "tab" && n.parentNode && n.parentNode.localName === "r" && !campo) out += "\t";
          else if (ln === "ptab" && !campo) out += "\t";
          else if ((ln === "br" || ln === "cr") && !campo) out += " ";
        });
        if (!out.trim()) return;
        const jc = parCrudo(hijo(pp, "pPr"), parDe(val(hijo(hijo(pp, "pPr"), "pStyle")) || idNormal)).jc;
        const partes = out.split("\t").map((t) => t.trim());
        if (partes.length === 1) { const k = jc === "center" ? 1 : jc === "right" || jc === "end" ? 2 : 0; huecos[k] += (huecos[k] ? " · " : "") + partes[0]; }
        else partes.slice(0, 3).forEach((t, k) => { if (t) huecos[k] += (huecos[k] ? " · " : "") + t; });
      });
      // número de página: en qué hueco está y con qué forma ("Página 1", "1 de 5"…)
      let num = null;
      huecos.forEach((t, k) => {
        if (!t.includes("\u0001")) return;
        const fmt = /p[áa]gina\s*\u0001\s*(de|\/)\s*\u0002/i.test(t) ? "pagina-total" : /\u0001\s*(de|of|\/)\s*\u0002/i.test(t) ? "total" : /(p[áa]gina|page|p[áa]g\.?)\s*\u0001/i.test(t) ? "pagina" : "n";
        if (!num) num = { lado: ["izq", "centro", "der"][k], fmt };
        huecos[k] = t.replace(/(p[áa]gina|page|p[áa]g\.?)?\s*\u0001(\s*(de|of|\/)\s*\u0002)?/i, "").replace(/^[\s·|–-]+|[\s·|–-]+$/g, "");
      });
      const limpios = huecos.map((t) => t.replace(/[\u0001\u0002]/g, "").trim());
      const usados = limpios.filter(Boolean).length;
      // un solo texto: va con su alineación; varios: izquierda, centro y derecha separados por tabuladores
      const texto = usados <= 1 ? limpios.find(Boolean) || "" : limpios.join("\t").replace(/\t+$/, "");
      return { texto: texto.slice(0, 300), alin: usados === 1 ? ["izq", "centro", "der"][limpios.findIndex(Boolean)] : "izq", num };
    }
    const docCampos = {};
    if (Object.keys(estilosW).length) docCampos.estilosP = estilosW;
    // Para los campos de Word (AUTHOR, TITLE, CREATEDATE, SAVEDATE…, y el formato de fecha del idioma): las propiedades
    // del documento (docProps/core.xml y app.xml) y el idioma del texto (docDefaults).
    try {
      const props = {};
      const core = await z.texto("docProps/core.xml");
      if (core) {
        const dc = P(core), tx = (ln) => { const n = dc.getElementsByTagNameNS("*", ln)[0]; return n ? n.textContent.trim() : ""; };
        [["title", "titulo"], ["subject", "asunto"], ["creator", "autor"], ["keywords", "claves"], ["description", "comentarios"], ["lastModifiedBy", "modificoPor"],
          ["created", "creado"], ["modified", "guardado"], ["lastPrinted", "impreso"], ["category", "categoria"]].forEach(([k, n]) => { const v = tx(k); if (v) props[n] = v; });
      }
      const app = await z.texto("docProps/app.xml");
      if (app) { const da = P(app), tx = (ln) => { const n = da.getElementsByTagNameNS("*", ln)[0]; return n ? n.textContent.trim() : ""; }; ["Company", "Manager", "Template"].forEach((k) => { const v = tx(k); if (v) props[k.toLowerCase()] = v; }); }
      const lang = rprDef && hijo(rprDef, "lang");
      const idioma = lang && (lang.getAttributeNS(W, "val") || lang.getAttribute("w:val"));
      if (Object.keys(props).length) docCampos.props = props;
      if (idioma) docCampos.idioma = idioma;
    } catch (e) { /* sin propiedades */ }
    if (Object.keys(formato).length) docCampos.formato = formato;
    // Protección de Word: «restringir edición» (solo lectura, comentarios, cambios marcados o formularios), «contraseña para
    // modificar» con «abrir como de solo lectura» recomendado, y «marcar como final». La contraseña no se guarda: solo su hash.
    {
      const st = await z.texto("word/settings.xml");
      if (st) {
        const x = P(st), dp = x.getElementsByTagNameNS(W, "documentProtection")[0], wp = x.getElementsByTagNameNS(W, "writeProtection")[0];
        const enf = dp && /^(1|true|on)$/.test(dp.getAttributeNS(W, "enforcement") || dp.getAttribute("w:enforcement") || "");
        const ed = dp && (dp.getAttributeNS(W, "edit") || dp.getAttribute("w:edit"));
        const tipo = { readOnly: "lectura", comments: "comentarios", trackedChanges: "cambios", forms: "formularios" }[ed];
        if (enf && tipo) docCampos.proteccion = { tipo, p: proteccionDe(dp, true) };
        // «Control de cambios» de Word (Revisar): quien edite el documento lo hace como sugerencias
        const tr = x.getElementsByTagNameNS(W, "trackRevisions")[0];
        if (tr && !/^(0|false|off)$/.test(tr.getAttributeNS(W, "val") || tr.getAttribute("w:val") || "")) docCampos.controlCambios = true;
        if (wp) docCampos.modificar = { recomendada: /^(1|true|on)$/.test(wp.getAttributeNS(W, "recommended") || wp.getAttribute("w:recommended") || "") || undefined, p: proteccionDe(wp, false) };
        const dts = x.getElementsByTagNameNS(W, "defaultTabStop")[0];
        if (dts) { const v = +(dts.getAttributeNS(W, "val") || dts.getAttribute("w:val")) / 20; if (v > 0 && Math.abs(v - 36) > 0.01) docCampos.tabDef = v; }
        // «Guiones automáticos» de Word (Disposición → Guiones): la zona, el límite de guiones seguidos y las mayúsculas
        const ah = x.getElementsByTagNameNS(W, "autoHyphenation")[0];
        if (ah && !/^(0|false|off)$/.test(ah.getAttributeNS(W, "val") || ah.getAttribute("w:val") || "")) {
          const vz = (n) => { const e = x.getElementsByTagNameNS(W, n)[0]; return e ? e.getAttributeNS(W, "val") || e.getAttribute("w:val") : null; };
          const dc = x.getElementsByTagNameNS(W, "doNotHyphenateCaps")[0];
          docCampos.guiones = { zona: +(vz("hyphenationZone") || 360) / 20, limite: +(vz("consecutiveHyphenLimit") || 0) || undefined, mayus: dc && !/^(0|false|off)$/.test(dc.getAttributeNS(W, "val") || dc.getAttribute("w:val") || "") ? false : undefined };
        }
      }
      if (await marcadoFinal(z)) docCampos.final = true;
    }
    {
      const enc = await zona("headerReference"), pie = await zona("footerReference");
      if (enc && enc.texto) { docCampos.encabezado = enc.texto; if (enc.alin !== "izq") docCampos.encAlin = enc.alin; }
      if (pie && pie.texto) { docCampos.pie = pie.texto; if (pie.alin !== "izq") docCampos.pieAlin = pie.alin; }
      const num = (pie && pie.num) || (enc && enc.num);
      // sin número de página en el archivo, aquí tampoco se muestra
      if (!num) docCampos.numeros = false;
      else {
        const pos = (pie && pie.num ? "pie-" : "enc-") + num.lado;
        if (pos !== "pie-der") docCampos.numPos = pos;
        if (num.fmt !== "n") docCampos.numFmt = num.fmt;
        const ini = sect ? +(at(hijo(sect, "pgNumType"), "start") || 0) : 0;
        if (ini > 1) docCampos.numInicio = ini;
      }
      if (sect && hijo(sect, "titlePg") && activo(hijo(sect, "titlePg")) && (docCampos.encabezado || docCampos.pie || num)) docCampos.primeraDistinta = true;
    }
    // Con secciones: cada una con su encabezado y su pie (si no trae los suyos, los de la anterior, como Word) y dónde
    // va su número de página
    if (paSx.length) {
      const porSec = [];
      let antE = null, antP = null;
      for (let k = 0; k < secciones.length; k++) {
        const e = (await zona("headerReference", secciones[k].sect)) || antE, p2 = (await zona("footerReference", secciones[k].sect)) || antP;
        porSec.push({ e, p: p2 });
        antE = e; antP = p2;
      }
      paSx.forEach(([pa, k]) => {
        const z = porSec[k];
        if (!z) return;
        let g;
        try { g = JSON.parse(pa.sx); } catch (e) { return; }
        const num = (z.p && z.p.num) || (z.e && z.e.num);
        g.ep = { e: z.e ? z.e.texto : "", ea: z.e && z.e.alin !== "izq" ? z.e.alin : undefined, p: z.p ? z.p.texto : "", pa: z.p && z.p.alin !== "izq" ? z.p.alin : undefined,
          np: num ? (z.p && z.p.num ? "pie-" : "enc-") + num.lado : "no", nf: num && num.fmt !== "n" ? num.fmt : undefined };
        pa.sx = JSON.stringify(g);
      });
    }
    if (aproximado) avisos.unshift("Las formas, los gráficos y los cuadros de texto de Word no se editan aquí: se ve su texto en un recuadro y al descargar en Word salen completos");
    while (parrafos.length > 1 && !parrafos[parrafos.length - 1].tramos.length && !Object.keys(parrafos[parrafos.length - 1].pa).length) parrafos.pop();
    return { archivo: u8, parrafos: parrafos.length ? parrafos : [{ pa: {}, tramos: [] }], pagina, doc: docCampos, formato, aviso: avisos.length ? avisos.join(". ") + "." : null };
  }

  /** Valores calculados de una tabla, recortados a lo que tiene datos (para Word y PowerPoint). */
  function tablaVisible(b) {
    const ev = F().evaluador((c, f) => b["c:" + F().ref(c, f)]);
    let mf = 0, mc = 0;
    Object.keys(b).forEach((k) => { if (!k.startsWith("c:") || b[k] === "" || b[k] == null) return; const r = F().parseRef(k.slice(2)); if (r) { mf = Math.max(mf, r.f + 1); mc = Math.max(mc, r.c + 1); } });
    const filas = [];
    for (let f = 0; f < mf; f++) { const fila = []; for (let c = 0; c < mc; c++) fila.push(F().formatear(ev.valor(c, f))); filas.push(fila); }
    return filas;
  }

  // ------------------------------------------------------------ Excel (.xlsx)
  const LIM_F = 5000, LIM_C = 60;   // lo que se muestra de una hoja importada (lo demás se conserva en el archivo de origen)
  const px2ancho = (px) => Math.max(1, Math.round(((px - 5) / 7) * 100) / 100);
  const argb = (c) => (/^#[0-9a-f]{6}$/i.test(c || "") ? "FF" + c.slice(1).toUpperCase() : null);
  const NUMFMT_BASE = { General: 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "@": 49 };
  /** Rango de celdas de una gráfica, en partes para Excel: etiquetas, nombres y series (como graficas.datos). */
  function partesGrafica(t, ev, r) {
    const num = (v) => F().esNum(v);
    const val = (c, f) => ev.valor(c, f);
    const conEtiquetas = (() => { for (let f = r.f1; f <= r.f2; f++) if (r.c2 > r.c1 && !num(val(r.c1, f)) && val(r.c1, f) !== "") return true; return false; })();
    const c0 = conEtiquetas ? r.c1 + 1 : r.c1;
    let conNombres = false;
    if (r.f2 > r.f1) for (let c = c0; c <= r.c2; c++) { const v = val(c, r.f1); if (v !== "" && !num(v)) conNombres = true; }
    const f0 = conNombres ? r.f1 + 1 : r.f1;
    const series = [];
    for (let c = c0; c <= r.c2; c++) {
      let alguno = false;
      for (let f = f0; f <= r.f2; f++) if (num(val(c, f))) alguno = true;
      if (alguno) series.push({ c, nombre: conNombres ? F().formatear(val(c, r.f1)) : "Serie " + (series.length + 1), valores: Array.from({ length: r.f2 - f0 + 1 }, (_, k) => (num(val(c, f0 + k)) ? +val(c, f0 + k) : null)) });
    }
    const etiquetas = Array.from({ length: r.f2 - f0 + 1 }, (_, k) => (conEtiquetas ? F().formatear(val(r.c1, f0 + k)) : String(k + 1)));
    return { conEtiquetas, conNombres, f0, series, etiquetas };
  }
  /** XML de una gráfica nativa de Excel (barras, líneas, área, dispersión o pastel). */
  function graficaXlsx(g, hoja, t, ev) {
    const r = NOM.graficas.rango(g.rango);
    if (!r) return null;
    const pg = partesGrafica(t, ev, r);
    if (!pg.series.length) return null;
    const H = F().prefijoHoja(hoja).replace(/!$/, "");
    const abs = (c1, f1, c2, f2) => `${H}!$${F().colLetra(c1)}$${f1 + 1}:$${F().colLetra(c2)}$${f2 + 1}`;
    const col = (NOM.graficas.PALETAS[g.colores] || NOM.graficas.PALETAS.clasica).c;
    const n = pg.etiquetas.length;
    const cache = (vals) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (v == null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:numCache>`;
    const cat = pg.conEtiquetas
      ? `<c:cat><c:strRef><c:f>${x(abs(r.c1, pg.f0, r.c1, r.f2))}</c:f><c:strCache><c:ptCount val="${n}"/>${pg.etiquetas.map((e, i) => `<c:pt idx="${i}"><c:v>${x(e)}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>` : "";
    const tx = (s) => (pg.conNombres ? `<c:tx><c:strRef><c:f>${x(abs(s.c, r.f1, s.c, r.f1))}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${x(s.nombre)}</c:v></c:pt></c:strCache></c:strRef></c:tx>` : `<c:tx><c:v>${x(s.nombre)}</c:v></c:tx>`);
    const relleno = (i) => `<c:spPr><a:solidFill><a:srgbClr val="${col[i % col.length].slice(1)}"/></a:solidFill></c:spPr>`;
    const linea = (i) => `<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="${col[i % col.length].slice(1)}"/></a:solidFill></a:ln></c:spPr>`;
    const vals = (s) => `<c:val><c:numRef><c:f>${x(abs(s.c, pg.f0, s.c, r.f2))}</c:f>${cache(s.valores)}</c:numRef></c:val>`;
    const ejes = `<c:axId val="1001"/><c:axId val="1002"/>`;
    const ejeX = (titulo, pos) => `<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos || "b"}"/>${titulo ? tituloXml(titulo) : ""}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
    const ejeY = (titulo, pos) => `<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos || "l"}"/><c:majorGridlines/>${titulo ? tituloXml(titulo) : ""}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="1001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
    function tituloXml(t0) { return `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${x(t0)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`; }
    let plot;
    const clase = g.clase || "barras";
    if (clase === "pastel") {
      const s = pg.series[0];
      plot = `<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${tx(s)}${pg.etiquetas.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>${relleno(i)}</c:dPt>`).join("")}${cat}${vals(s)}</c:ser><c:firstSliceAng val="0"/></c:pieChart>`;
    } else if (clase === "dispersion") {
      const xs = pg.series.length > 1 ? pg.series[0] : null, ys = xs ? pg.series.slice(1) : pg.series;
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${ys.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx(s)}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/>${relleno(i)}</c:marker>${xs ? `<c:xVal><c:numRef><c:f>${x(abs(xs.c, pg.f0, xs.c, r.f2))}</c:f>${cache(xs.valores)}</c:numRef></c:xVal>` : ""}<c:yVal><c:numRef><c:f>${x(abs(s.c, pg.f0, s.c, r.f2))}</c:f>${cache(s.valores)}</c:numRef></c:yVal><c:smooth val="0"/></c:ser>`).join("")}${ejes}</c:scatterChart>`
        + `<c:valAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${g.ejeX ? tituloXml(g.ejeX) : ""}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="1002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>${ejeY(g.ejeY)}`;
    } else {
      const sers = pg.series.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx(s)}${clase === "lineas" ? linea(i) + `<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>` : relleno(i)}${cat}${vals(s)}${clase === "lineas" ? `<c:smooth val="0"/>` : ""}</c:ser>`).join("");
      if (clase === "lineas") plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${ejes}</c:lineChart>`;
      else if (clase === "area") plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${ejes}</c:areaChart>`;
      else plot = `<c:barChart><c:barDir val="${clase === "barras-h" ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="80"/>${ejes}</c:barChart>`;
      plot += ejeX(g.ejeX, clase === "barras-h" ? "l" : "b") + ejeY(g.ejeY, clase === "barras-h" ? "b" : "l");
    }
    return XML + `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${g.titulo ? tituloXml(g.titulo) + `<c:autoTitleDeleted val="0"/>` : `<c:autoTitleDeleted val="1"/>`}<c:plotArea><c:layout/>${plot}</c:plotArea>${pg.series.length > 1 || clase === "pastel" ? `<c:legend><c:legendPos val="${clase === "pastel" ? "r" : "b"}"/><c:overlay val="0"/></c:legend>` : ""}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  }
  const OPS_CF = { ">": "greaterThan", "<": "lessThan", ">=": "greaterThanOrEqual", "<=": "lessThanOrEqual", "=": "equal", "<>": "notEqual", entre: "between" };
  const valorCf = (v) => { const n = F().calcularLiteral(String(v == null ? "" : v).trim()); return F().esNum(n) ? String(+n) : `"${String(v == null ? "" : v).replace(/"/g, '""')}"`; };

  // ---- piezas que comparten "desde cero" y "a partir del original"
  const CELDA_FMT = ["n", "i", "u", "k", "tl", "r", "bd", "f", "al", "aj", "pr"];
  const quitaAttr = (attrs, nombre) => attrs.replace(new RegExp("\\s+" + nombre.replace(/[.:]/g, "\\$&") + "\\s*=\\s*(\"[^\"]*\"|'[^']*')", "g"), "");
  const mismoValor = (a, b) => { const n = (v) => (v == null || v === "" || v === false ? null : v); a = n(a); b = n(b); return a === b || (typeof a === "object" && typeof b === "object" && a && b && NOM.origen.estable(a) === NOM.origen.estable(b)); };

  /**
   * Fábrica de estilos y celdas de Excel. Sin `ini` arma un styles.xml completo; con `ini`
   * (los conteos del styles.xml de origen) solo junta lo NUEVO, numerado después de lo que ya hay.
   */
  function fabricaXlsx(doc, ini) {
    const O = ini && ini.origen;
    const fuentes = O ? [] : ['<font><sz val="11"/><name val="Calibri"/></font>'];
    const rellenos = O ? [] : ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const bordes = O ? [] : ["<border><left/><right/><top/><bottom/><diagonal/></border>"];
    const xfs = O ? [] : ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    const numFmts = [], dxfs = [];
    const base = O ? { fuente: O.fuentes.length, relleno: O.nRellenos, borde: O.nBordes, xf: O.xfs.length } : { fuente: 0, relleno: 0, borde: 0, xf: 0 };
    const idx = { fuente: O ? {} : { "": 0 }, relleno: O ? {} : { "": 0 }, borde: O ? {} : { "": 0 }, num: {}, xf: O ? {} : { "": 0 } };
    const registrar = (tipo, clave, lista, xml) => { if (idx[tipo][clave] == null) { lista.push(xml); idx[tipo][clave] = base[tipo] + lista.length - 1; } return idx[tipo][clave]; };
    let sigNum = O ? O.maxNum + 1 : 164;
    function numId(cod) {
      if (!cod || cod === "General") return 0;
      const cx = cod.replace(/aaaa/g, "yyyy").replace(/aa/g, "yy");
      if (NUMFMT_BASE[cx] != null) return NUMFMT_BASE[cx];
      if (O && O.codigos[cx] != null) return O.codigos[cx];
      if (idx.num[cx] == null) { const id = sigNum++; numFmts.push(`<numFmt numFmtId="${id}" formatCode="${x(cx)}"/>`); idx.num[cx] = id; }
      return idx.num[cx];
    }
    const xmlRelleno = (r) => `<fill><patternFill patternType="solid"><fgColor rgb="${r}"/><bgColor indexed="64"/></patternFill></fill>`;
    const xmlBorde = (bd) => { const lado = (l, si) => (si ? `<${l} style="thin"><color auto="1"/></${l}>` : `<${l}/>`); return `<border>${lado("left", bd.includes("i"))}${lado("right", bd.includes("d"))}${lado("top", bd.includes("a"))}${lado("bottom", bd.includes("b"))}<diagonal/></border>`; };
    const tamExcel = (tl) => (tl ? Math.round(tl * 1.0476 * 2) / 2 : 11);
    const ALIN = { izq: "left", centro: "center", der: "right" };
    /** Estilo de una celda armado solo con lo que dice el modelo (documento hecho en NOM). */
    function estilo(t, ref, esFecha) {
      const b = !!t["n:" + ref], i = !!t["i:" + ref], u = !!t["u:" + ref], k = argb(t["k:" + ref]), tl = +t["tl:" + ref] || 0, tf = String(t["tf:" + ref] || "").replace(/[<>"&]/g, "");
      const kf = [b, i, u, k, tl].join("|") + (tf ? "|" + tf : "");
      const fid = kf === "false|false|false|null|0" ? 0 : registrar("fuente", kf, fuentes, `<font>${b ? "<b/>" : ""}${i ? "<i/>" : ""}${u ? "<u/>" : ""}<sz val="${tamExcel(tl)}"/>${k ? `<color rgb="${k}"/>` : ""}<name val="${tf || "Calibri"}"/></font>`);
      const r = argb(t["r:" + ref]);
      const rid = r ? registrar("relleno", r, rellenos, xmlRelleno(r)) : 0;
      const bd = String(t["bd:" + ref] || "").split("").sort().join("");
      const bid = bd ? registrar("borde", bd, bordes, xmlBorde(bd)) : 0;
      const nid = numId(t["f:" + ref] || (esFecha ? "dd/mm/yyyy" : null));
      const al = t["al:" + ref], aj = !!t["aj:" + ref], m = !!t["m:" + ref];
      const alin = al || aj || m ? `<alignment${al ? ` horizontal="${ALIN[al] || "general"}"` : ""}${m ? ' vertical="center"' : ""}${aj ? ' wrapText="1"' : ""}/>` : "";
      // (protección de la celda: sin bloquear y fórmula oculta, para cuando la hoja se protege)
      const pr = String(t["pr:" + ref] || ""), prot = pr ? `<protection${pr.includes("u") ? ' locked="0"' : ""}${pr.includes("h") ? ' hidden="1"' : ""}/>` : "";
      const kx = [fid, rid, bid, nid, alin, prot].join("|");
      if (kx === "0|0|0|0||") return 0;
      const dentro = alin + prot;
      return registrar("xf", kx, xfs, `<xf numFmtId="${nid}" fontId="${fid}" fillId="${rid}" borderId="${bid}" xfId="0"${fid ? ' applyFont="1"' : ""}${rid ? ' applyFill="1"' : ""}${bid ? ' applyBorder="1"' : ""}${nid ? ' applyNumberFormat="1"' : ""}${alin ? ' applyAlignment="1"' : ""}${prot ? ' applyProtection="1"' : ""}${dentro ? ">" + dentro + "</xf>" : "/>"}`);
    }
    /**
     * Estilo a partir del que tenía la celda en el archivo original (sOrig): solo se toca lo que la
     * persona cambió (`camb` = llaves de formato distintas); lo demás (fuente, tema, protección…) se queda.
     */
    function estiloDesde(t, ref, sOrig, camb, fmtFecha) {
      const xf = O.xfs[sOrig] || O.xfs[0] || { attrs: "", numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alin: "", otros: "" };
      let { fontId, fillId, borderId, numFmtId, alin } = xf;
      const aplica = {};
      if (["n", "i", "u", "k", "tl"].some((k) => camb.has(k))) {
        let dentro = O.fuentes[fontId] != null ? O.fuentes[fontId] : O.fuentes[0] || "";
        const quita = (tag) => { dentro = dentro.replace(new RegExp("<" + tag + "(?=[\\s/>])[^>]*/>|<" + tag + "(?=[\\s>])[^>]*>[\\s\\S]*?</" + tag + ">", "g"), ""); };
        if (camb.has("n")) { quita("b"); if (t["n:" + ref]) dentro = "<b/>" + dentro; }
        if (camb.has("i")) { quita("i"); if (t["i:" + ref]) dentro = "<i/>" + dentro; }
        if (camb.has("u")) { quita("u"); if (t["u:" + ref]) dentro = "<u/>" + dentro; }
        if (camb.has("tl")) { quita("sz"); dentro += `<sz val="${tamExcel(+t["tl:" + ref] || 0)}"/>`; }
        if (camb.has("k")) { quita("color"); const k = argb(t["k:" + ref]); if (k) dentro += `<color rgb="${k}"/>`; }
        fontId = registrar("fuente", dentro, fuentes, `<font>${dentro}</font>`); aplica.applyFont = 1;
      }
      if (camb.has("r")) { const r = argb(t["r:" + ref]); fillId = r ? registrar("relleno", r, rellenos, xmlRelleno(r)) : 0; aplica.applyFill = 1; }
      if (camb.has("bd")) { const bd = String(t["bd:" + ref] || "").split("").sort().join(""); borderId = bd ? registrar("borde", bd, bordes, xmlBorde(bd)) : 0; aplica.applyBorder = 1; }
      if (camb.has("f") || fmtFecha) { numFmtId = numId(t["f:" + ref] || fmtFecha); aplica.applyNumberFormat = 1; }
      if (camb.has("al") || camb.has("aj")) {
        let a = alin || "";
        if (camb.has("al")) { a = quitaAttr(a, "horizontal"); const al = t["al:" + ref]; if (al) a += ` horizontal="${ALIN[al] || "general"}"`; }
        if (camb.has("aj")) { a = quitaAttr(a, "wrapText"); if (t["aj:" + ref]) a += ' wrapText="1"'; }
        alin = a; aplica.applyAlignment = 1;
      }
      let otros = xf.otros || "";
      if (camb.has("pr")) {
        const pr = String(t["pr:" + ref] || "");
        otros = otros.replace(/<protection\b[^>]*\/>|<protection\b[^>]*>[\s\S]*?<\/protection>/g, "");
        if (pr) otros = `<protection${pr.includes("u") ? ' locked="0"' : ""}${pr.includes("h") ? ' hidden="1"' : ""}/>` + otros;
        aplica.applyProtection = 1;
      }
      if (fontId === xf.fontId && fillId === xf.fillId && borderId === xf.borderId && numFmtId === xf.numFmtId && alin === xf.alin && otros === (xf.otros || "")) return sOrig;
      let attrs = xf.attrs;
      ["numFmtId", "fontId", "fillId", "borderId", ...Object.keys(aplica)].forEach((k) => { attrs = quitaAttr(attrs, k); });
      attrs = ` numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}"` + attrs + Object.keys(aplica).map((k) => ` ${k}="1"`).join("");
      const dentro = (alin && alin.trim() ? `<alignment${alin}/>` : "") + otros;
      return registrar("xf", attrs + "|" + dentro, xfs, dentro ? `<xf${attrs}>${dentro}</xf>` : `<xf${attrs}/>`);
    }
    function dxf(rg) {
      const k = argb(rg.color), f = argb(rg.fondo);
      dxfs.push(`<dxf>${k || rg.negrita ? `<font>${rg.negrita ? "<b/>" : ""}${k ? `<color rgb="${k}"/>` : ""}</font>` : ""}${f ? `<fill><patternFill><bgColor rgb="${f}"/></patternFill></fill>` : ""}</dxf>`);
      return (O ? O.nDxfs : 0) + dxfs.length - 1;
    }
    /** ¿El valor escrito es una fecha? (para darle formato de fecha si la celda no tiene uno). */
    const esFechaLit = (crudo) => crudo !== "" && !crudo.startsWith("=") && F().calcularLiteral(crudo.trim()) instanceof F().FechaF;
    /** El valor calculado como lo guarda Excel: { tipo, val } (los errores con su nombre en inglés). */
    const valorXml = (v) => (F().esNum(v) ? { tipo: "", val: `<v>${+v}</v>` } : typeof v === "boolean" ? { tipo: ' t="b"', val: `<v>${v ? 1 : 0}</v>` }
      : v instanceof F().ErrorF ? { tipo: ' t="e"', val: `<v>${x(ERR_XL[v.codigo] || v.codigo)}</v>` } : { tipo: ' t="str"', val: `<v>${x(v)}</v>` });
    /** El cm de las fórmulas de matriz dinámica (0: el libro no lo tiene y se guardan como fórmulas de matriz). */
    const cmDinamica = () => { fab.usaDA = true; return fab.cmDA == null ? 1 : fab.cmDA; };
    /** XML de una celda con el valor del modelo y el estilo `sx` (ya calculado). */
    function celda(t, ev, ref, c, f, sx) {
      const crudo = t["c:" + ref] == null ? "" : String(t["c:" + ref]), s = sx ? ` s="${sx}"` : "";
      const td = /^=\s*(?:TABLA|TABLE)\s*\(\s*(\$?[A-Z]{1,3}\$?\d+)?\s*[;,]\s*(\$?[A-Z]{1,3}\$?\d+)?\s*\)\s*$/i.exec(crudo);
      if (td) {
        // tabla de datos, como la guarda Excel: la fórmula en la primera celda de su lugar, con el lugar y las entradas
        const v = ev.valor(c, f), { tipo, val } = valorXml(v), esT = (cc, ff) => /^=\s*(TABLA|TABLE)\s*\(/i.test(String(t["c:" + F().ref(cc, ff)] || ""));
        if ((c > 0 && esT(c - 1, f)) || (f > 0 && esT(c, f - 1))) return `<c r="${ref}"${s}${tipo}>${val}</c>`;
        let c2 = c, f2 = f;
        while (esT(c2 + 1, f)) c2++;
        while (esT(c, f2 + 1)) f2++;
        const fi = td[1] ? td[1].replace(/\$/g, "").toUpperCase() : null, ci = td[2] ? td[2].replace(/\$/g, "").toUpperCase() : null;
        const at2 = fi && ci ? ` dt2D="1" dtr="1" r1="${fi}" r2="${ci}"` : fi ? ` dt2D="0" dtr="1" r1="${fi}"` : ` dt2D="0" dtr="0" r1="${ci}"`;
        return `<c r="${ref}"${s}${tipo}><f t="dataTable" ref="${ref}:${F().ref(c2, f2)}"${at2}/>${val}</c>`;
      }
      if (crudo.startsWith("=")) {
        const v = ev.valor(c, f), { tipo, val } = valorXml(v), matrizD = !!(ev.esMatriz && ev.esMatriz(c, f));
        const fx = x(formulaExcel(crudo, t.xlv, !matrizD));
        // una fórmula de matrices (da una matriz o calcula con ellas): Excel la guarda con t="array", el lugar que
        // derrama y cm (matriz dinámica); sin eso Excel la leería como una fórmula vieja (con intersección implícita)
        if (matrizD) {
          const d = ev.derrame(c, f), cm = cmDinamica();
          const area = d && !d.bloqueado && d.filas * d.cols > 1 ? `${ref}:${F().ref(c + d.cols - 1, f + d.filas - 1)}` : ref;
          return `<c r="${ref}"${s}${tipo}${cm ? ` cm="${cm}"` : ""}><f t="array" ref="${area}">${fx}</f>${val}</c>`;
        }
        return `<c r="${ref}"${s}${tipo}><f>${fx}</f>${val}</c>`;
      }
      if (crudo === "") {
        // lo que derrama la fórmula de otra celda: Excel guarda ahí el valor (sin fórmula)
        if (ev.origen && ev.origen(c, f)) { const v = ev.valor(c, f); if (v !== "" && v != null) { const { tipo, val } = valorXml(v); return `<c r="${ref}"${s}${tipo}>${val}</c>`; } }
        return `<c r="${ref}"${s}/>`;
      }
      if (crudo.startsWith("'")) return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${x(crudo.slice(1))}</t></is></c>`;
      if (t["f:" + ref] === "@") return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${x(crudo)}</t></is></c>`;
      const lit = F().calcularLiteral(crudo.trim());
      if (typeof lit === "boolean") return `<c r="${ref}"${s} t="b"><v>${lit ? 1 : 0}</v></c>`;
      return F().esNum(lit) ? `<c r="${ref}"${s}><v>${+lit}</v></c>` : `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${x(crudo)}</t></is></c>`;
    }
    const combinadas = (t) => Object.keys(t).filter((k) => k.startsWith("m:") && /^\d+x\d+$/.test(String(t[k] || ""))).map((k) => {
      const r = F().parseRef(k.slice(2)), [w, h] = String(t[k]).split("x").map(Number);
      return r && w * h > 1 ? `<mergeCell ref="${F().ref(r.c, r.f)}:${F().ref(r.c + w - 1, r.f + h - 1)}"/>` : "";
    }).filter(Boolean);
    let prioridad = (O ? 1000 : 0) + 1;
    function regla(rg) {
      if (!rg || !rg.rango) return "";
      const primero = rg.rango.split(":")[0], d = dxf(rg);
      let r0;
      if (OPS_CF[rg.op]) r0 = `<cfRule type="cellIs" dxfId="${d}" priority="${prioridad++}" operator="${OPS_CF[rg.op]}"><formula>${x(valorCf(rg.v1))}</formula>${rg.op === "entre" ? `<formula>${x(valorCf(rg.v2))}</formula>` : ""}</cfRule>`;
      else if (rg.op === "contiene") r0 = `<cfRule type="containsText" dxfId="${d}" priority="${prioridad++}" operator="containsText" text="${x(rg.v1 || "")}"><formula>${x(`NOT(ISERROR(SEARCH("${String(rg.v1 || "").replace(/"/g, '""')}",${primero})))`)}</formula></cfRule>`;
      else if (rg.op === "vacia") r0 = `<cfRule type="containsBlanks" dxfId="${d}" priority="${prioridad++}"><formula>${x(`LEN(TRIM(${primero}))=0`)}</formula></cfRule>`;
      else if (rg.op === "novacia") r0 = `<cfRule type="notContainsBlanks" dxfId="${d}" priority="${prioridad++}"><formula>${x(`LEN(TRIM(${primero}))>0`)}</formula></cfRule>`;
      return r0 ? `<conditionalFormatting sqref="${x(rg.rango)}">${r0}</conditionalFormatting>` : "";
    }
    function validacion(v) {
      if (!v || !v.rango) return "";
      const msg = v.mensaje ? ` showErrorMessage="1" error="${x(v.mensaje)}"` : ' showErrorMessage="1"';
      if (v.tipo === "lista") return `<dataValidation type="list" allowBlank="1"${msg} sqref="${x(v.rango)}"><formula1>${x('"' + (v.valores || []).join(",").slice(0, 250) + '"')}</formula1></dataValidation>`;
      if (v.tipo === "numero") {
        const a = v.min !== "" && v.min != null, b = v.max !== "" && v.max != null;
        if (!a && !b) return "";
        const op = a && b ? "between" : a ? "greaterThanOrEqual" : "lessThanOrEqual";
        return `<dataValidation type="decimal" operator="${op}" allowBlank="1"${msg} sqref="${x(v.rango)}"><formula1>${x(String(a ? v.min : v.max))}</formula1>${a && b ? `<formula2>${x(String(v.max))}</formula2>` : ""}</dataValidation>`;
      }
      return "";
    }
    const panel = (t) => { const ff = +t.fijasF || 0, fc = +t.fijasC || 0; return ff || fc ? `<pane${fc ? ` xSplit="${fc}"` : ""}${ff ? ` ySplit="${ff}"` : ""} topLeftCell="${F().ref(fc, ff)}" activePane="${ff && fc ? "bottomRight" : ff ? "bottomLeft" : "topRight"}" state="frozen"/>` : ""; };
    /** Ancla de una gráfica en el dibujo de la hoja (con sus espacios de nombres: sirve suelta). */
    const anclaGrafica = (rid, ng, colIni, fila0) => `<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" editAs="oneCell"><xdr:from><xdr:col>${colIni}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fila0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${colIni + 8}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fila0 + 16}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${ng + 1000}" name="Gráfica ${ng}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
    const fab = { fuentes, rellenos, bordes, xfs, numFmts, dxfs, estilo, estiloDesde, dxf, celda, esFechaLit, combinadas, regla, validacion, panel, anclaGrafica, usaDA: false, cmDA: null };
    return fab;
  }

  /** Hoja de cálculo → .xlsx. Si el documento vino de un Excel, se parte de ese archivo (xlsxConOrigen). */
  /** Esquema de Excel (Datos → Agrupar): atributos de la fila n (1, 2…) —oculta, nivel y, en la de resumen de abajo, si
      el grupo de arriba está contraído— y de la columna (su letra). */
  function esqFila(t, n) {
    const nv = +t["ol:" + n] || 0, ant = +t["ol:" + (n - 1)] || 0;
    return (t["oh:" + n] ? ' hidden="1"' : "") + (nv ? ` outlineLevel="${nv}"` : "") + (ant > nv && t["oh:" + (n - 1)] ? ' collapsed="1"' : "");
  }
  function esqCol(t, letra) {
    const p = F().parseRef(letra + "1"), ant = p && p.c > 0 ? F().colLetra(p.c - 1) : null;
    const nv = +t["olc:" + letra] || 0, na = ant ? +t["olc:" + ant] || 0 : 0;
    return (t["ow:" + letra] ? ' hidden="1"' : "") + (nv ? ` outlineLevel="${nv}"` : "") + (ant && na > nv && t["ow:" + ant] ? ' collapsed="1"' : "");
  }
  function nivelesEsq(t) {
    let f = 0, c = 0;
    for (const k in t) { if (k.startsWith("ol:")) f = Math.max(f, +t[k] || 0); else if (k.startsWith("olc:")) c = Math.max(c, +t[k] || 0); }
    return { f, c };
  }
  /** El modo de cálculo del libro como lo escribe Excel en calcPr (« calcMode="manual"»), o "" si es automático. */
  function calcModeDe(doc) {
    const m = doc && typeof doc.valor === "function" ? doc.valor("_doc", "calculo") : null;
    return m === "manual" ? ' calcMode="manual"' : m === "sinTablas" ? ' calcMode="autoNoTable"' : "";
  }
  /** Los nombres definidos de NOM como los escribe Excel en workbook.xml (<definedNames>…), o "" si no hay. */
  function definedNameXml(o, pref) {
    const d = formulaExcel(String(o.d || "").startsWith("=") ? o.d : "=" + (o.d || ""), 7).replace(/^=/, "");
    return `<${pref || ""}definedName name="${x(o.n)}">${x(d)}</${pref || ""}definedName>`;
  }
  function definedNamesDe(doc, pref) {
    const m = doc && typeof doc.valor === "function" ? doc.valor("_doc", "nombres") : null;
    const l = m && typeof m === "object" ? Object.values(m).filter((o) => o && o.n && o.d != null) : [];
    if (!l.length) return "";
    l.sort((a, b) => a.n.toLowerCase().localeCompare(b.n.toLowerCase()));
    return `<${pref || ""}definedNames>${l.map((o) => definedNameXml(o, pref)).join("")}</${pref || ""}definedNames>`;
  }
  function xlsx(doc) {
    const og = NOM.ooxml.usarOrigen !== false && NOM.origen ? NOM.origen.leer(doc) : null;
    if (og && og.fmt === "xlsx") {
      try { return xlsxConOrigen(doc, og); }
      catch (e) { if (typeof console !== "undefined") console.warn("Excel: no se pudo partir del archivo original; se arma desde cero.", e); }
    }
    return xlsxNuevo(doc);
  }

  function xlsxNuevo(doc) {
    const tablas = F().hojasDe(doc);
    if (!tablas.length) throw new Error("Este documento no tiene ninguna tabla");
    const nombres = [];
    const fab = fabricaXlsx(doc, null);
    const graficasArch = [], dibujos = [], relsHojas = [], mediosXl = [], tablasArch = [], nombresLo = new Set(), dinArch = [];
    const hojas = tablas.map((t, i) => {
      let n = F().nombreDeHoja(doc, t).replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || `Hoja${i + 1}`;
      while (nombres.includes(n)) n = n.slice(0, 28) + " " + (i + 1);
      nombres.push(n);
      const ev = F().evaluadorDeTabla(doc, t);
      const ncols = Math.max(t.cols || 10, 1);
      // Gráficas de esta hoja: un dibujo con una gráfica nativa por cada una.
      const gs = doc.vivos().filter((g) => g.tipo === "grafica" && g.hoja === t.id);
      const ims = doc.vivos().filter((b) => b.tipo === "imagen" && b.hoja === t.id && b.src);
      let dibujo = "";
      if (gs.length || ims.length) {
        const nd = dibujos.length + 1, anclas = [], relsDibujo = [];
        let fila0 = 1;
        const colIni = Math.min(ncols + 1, 20);
        gs.forEach((g) => {
          const xmlG = graficaXlsx(g, n, t, ev);
          if (!xmlG) return;
          graficasArch.push(xmlG);
          const ng = graficasArch.length, rid = "rIdG" + ng;
          relsDibujo.push({ id: rid, tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", destino: `../charts/chart${ng}.xml` });
          anclas.push(fab.anclaGrafica(rid, ng, colIni, fila0));
          fila0 += 18;
        });
        ims.forEach((b) => {
          const im = imagen(b.src);
          if (!im) return;
          mediosXl.push({ nombre: `xl/media/imagen${mediosXl.length + 1}.${im.ext}`, datos: im.bytes, ct: ctImagen(im.ext) });
          const nm = mediosXl.length, rid = "rIdI" + nm;
          relsDibujo.push({ id: rid, tipo: T_IMAGEN_XL, destino: `../media/imagen${nm}.${im.ext}` });
          anclas.push(anclaImagenXl(rid, nm, b));
        });
        if (anclas.length) {
          dibujos.push({ n: nd, xml: XML + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anclas.join("")}</xdr:wsDr>`, rels: relsDibujo });
          relsHojas[i] = [{ id: "rIdD1", tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", destino: `../drawings/drawing${nd}.xml` }];
          dibujo = `<drawing r:id="rIdD1"/>`;
        }
      }
      // tablas con formato (Dar formato como tabla)
      const los = tablasLoDe(t);
      if (los.length) {
        const ids = los.map((lo) => {
          tablasArch.push(xmlTablaXl(lo, tablasArch.length + 1, nombreTablaXl(lo.n, nombresLo), columnasTablaXl(lo, t, ev)));
          const rid = "rIdT" + tablasArch.length;
          (relsHojas[i] = relsHojas[i] || []).push({ id: rid, tipo: T_TABLA_XL, destino: `../tables/table${tablasArch.length}.xml` });
          return rid;
        });
        dibujo += `<tableParts count="${ids.length}">${ids.map((r) => `<tablePart r:id="${r}"/>`).join("")}</tableParts>`;
      }
      // tablas dinámicas: su definición (enlazada desde la hoja), su caché y sus registros
      dinamicasDe(doc, t).forEach((r) => {
        dinArch.push(r.partes);
        (relsHojas[i] = relsHojas[i] || []).push({ id: "rIdP" + dinArch.length, tipo: T_PV_TABLA, destino: `../pivotTables/pivotTable${dinArch.length}.xml` });
      });
      return hojaDesdeCero(t, ev, fab, (ref, esFecha) => fab.estilo(t, ref, esFecha), i === 0, dibujo);
    });
    const { fuentes, rellenos, bordes, xfs, numFmts, dxfs } = fab;
    const libro = XML + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${nombres.map((n, i) => `<sheet name="${x(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets>${definedNamesDe(doc)}<calcPr calcId="191029"${calcModeDe(doc)} fullCalcOnLoad="1"/>${dinArch.length ? `<pivotCaches>${dinArch.map((_, k) => `<pivotCache cacheId="${101 + k}" r:id="rIdPC${k + 1}"/>`).join("")}</pivotCaches>` : ""}</workbook>`;
    const estilos = XML + `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join("")}</numFmts>` : ""}<fonts count="${fuentes.length}">${fuentes.join("")}</fonts>
<fills count="${rellenos.length}">${rellenos.join("")}</fills><borders count="${bordes.length}">${bordes.join("")}</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>${dxfs.length ? `<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>` : ""}</styleSheet>`;
    const relsLibro = [...hojas.map((_, i) => ({ id: "rId" + (i + 1), tipo: T.hoja, destino: `worksheets/hoja${i + 1}.xml` })),
      ...dinArch.map((_, k) => ({ id: "rIdPC" + (k + 1), tipo: T_PV_CACHE, destino: `pivotCache/pivotCacheDefinition${k + 1}.xml` })),
      { id: "rIdEst", tipo: T.estilos, destino: "styles.xml" }, ...(fab.usaDA ? [{ id: "rIdMeta", tipo: T_META, destino: "metadata.xml" }] : [])];
    return zip([
      { nombre: "[Content_Types].xml", datos: CT(`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${fab.usaDA ? `<Override PartName="/xl/metadata.xml" ContentType="${CT_META}"/>` : ""}
${hojas.map((_, i) => `<Override PartName="/xl/worksheets/hoja${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}
${dibujos.map((d) => `<Override PartName="/xl/drawings/drawing${d.n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join("")}
${graficasArch.map((_, i) => `<Override PartName="/xl/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("")}
${mediosXl.map((m) => `<Override PartName="/${m.nombre}" ContentType="${m.ct}"/>`).join("")}
${tablasArch.map((_, i) => `<Override PartName="/xl/tables/table${i + 1}.xml" ContentType="${CT_TABLA_XL}"/>`).join("")}
${dinArch.map((_, k) => `<Override PartName="/xl/pivotTables/pivotTable${k + 1}.xml" ContentType="${CT_PV_TABLA}"/><Override PartName="/xl/pivotCache/pivotCacheDefinition${k + 1}.xml" ContentType="${CT_PV_CACHE}"/><Override PartName="/xl/pivotCache/pivotCacheRecords${k + 1}.xml" ContentType="${CT_PV_REG}"/>`).join("")}`) },
      { nombre: "_rels/.rels", datos: RELS([{ tipo: T.doc, destino: "xl/workbook.xml" }]) },
      { nombre: "xl/workbook.xml", datos: libro },
      { nombre: "xl/styles.xml", datos: estilos },
      ...(fab.usaDA ? [{ nombre: "xl/metadata.xml", datos: META_DA }] : []),
      { nombre: "xl/_rels/workbook.xml.rels", datos: RELS(relsLibro) },
      ...hojas.map((h, i) => ({ nombre: `xl/worksheets/hoja${i + 1}.xml`, datos: h })),
      ...relsHojas.map((r, i) => (r ? { nombre: `xl/worksheets/_rels/hoja${i + 1}.xml.rels`, datos: RELS(r) } : null)).filter(Boolean),
      ...dibujos.flatMap((d) => [{ nombre: `xl/drawings/drawing${d.n}.xml`, datos: d.xml }, { nombre: `xl/drawings/_rels/drawing${d.n}.xml.rels`, datos: RELS(d.rels) }]),
      ...graficasArch.map((g, i) => ({ nombre: `xl/charts/chart${i + 1}.xml`, datos: g })),
      ...mediosXl.map((m) => ({ nombre: m.nombre, datos: m.datos })),
      ...tablasArch.map((d, i) => ({ nombre: `xl/tables/table${i + 1}.xml`, datos: d })),
      ...dinArch.flatMap((d, k) => [
        { nombre: `xl/pivotTables/pivotTable${k + 1}.xml`, datos: d.tabla.replace("\u0001CACHE\u0001", String(101 + k)) },
        { nombre: `xl/pivotTables/_rels/pivotTable${k + 1}.xml.rels`, datos: RELS([{ id: "rId1", tipo: T_PV_CACHE, destino: `../pivotCache/pivotCacheDefinition${k + 1}.xml` }]) },
        { nombre: `xl/pivotCache/pivotCacheDefinition${k + 1}.xml`, datos: d.cache },
        { nombre: `xl/pivotCache/_rels/pivotCacheDefinition${k + 1}.xml.rels`, datos: RELS([{ id: "rId1", tipo: T_PV_REG, destino: `pivotCacheRecords${k + 1}.xml` }]) },
        { nombre: `xl/pivotCache/pivotCacheRecords${k + 1}.xml`, datos: d.registros },
      ]),
    ]);
  }

  // ------------------------------------------------------------ partir del archivo original
  /** Aplica reemplazos [ini, fin, texto] (que no se enciman) sobre un texto. */
  function aplicarPiezas(xml, piezas) {
    if (!piezas.length) return xml;
    const ps = piezas.map((p, k) => [p[0], p[1], p[2], k]).sort((a, b) => a[0] - b[0] || a[3] - b[3]);
    let r = "", i = 0;
    for (const [a, b, t] of ps) { if (a < i) continue; r += xml.slice(i, a) + t; i = b; }
    return r + xml.slice(i);
  }
  const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  /** El .zip de origen abierto para cambiarle partes: texto o bytes, relaciones y tipos de contenido. */
  function paqueteDe(og) {
    const Z = og.zip, nuevos = new Map(), quitados = new Set(), orden = Z.nombres();
    const dec = new TextDecoder();
    const CTN = "[Content_Types].xml";
    const P = {
      tiene: (n) => nuevos.has(n) || (Z.tiene(n) && !quitados.has(n)),
      texto(n) { if (nuevos.has(n)) { const v = nuevos.get(n); return typeof v === "string" ? v : dec.decode(v); } return quitados.has(n) ? null : Z.texto(n); },
      bytes(n) { if (nuevos.has(n)) { const v = nuevos.get(n); return typeof v === "string" ? enc.encode(v) : v; } return quitados.has(n) ? null : Z.bytes(n); },
      pon(n, datos) { if (!P.tiene(n) && !orden.includes(n)) orden.push(n); quitados.delete(n); nuevos.set(n, datos); },
      quita(n) { nuevos.delete(n); quitados.add(n); },
      nombres: () => orden.filter((n) => P.tiene(n)),
      relsDe: (ruta) => { const k = ruta.lastIndexOf("/"); return ruta.slice(0, k + 1) + "_rels/" + ruta.slice(k + 1) + ".rels"; },
      resolver(ruta, destino) {
        if (destino.startsWith("/")) return destino.slice(1);
        const b = ruta.split("/").slice(0, -1);
        destino.split("/").forEach((s) => { if (s === "..") b.pop(); else if (s !== "." && s !== "") b.push(s); });
        return b.join("/");
      },
      rels(ruta) {
        const t = P.texto(P.relsDe(ruta)), r = [];
        if (!t) return r;
        const re = /<(?:\w+:)?Relationship\b([^>]*?)\/?>/g;
        let m;
        while ((m = re.exec(t))) {
          const a = (n) => { const q = new RegExp("\\s" + n + '\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\')').exec(m[1]); return q ? NOM.origen.desescapar(q[1] != null ? q[1] : q[2]) : null; };
          r.push({ id: a("Id"), tipo: a("Type"), destino: a("Target"), externo: a("TargetMode") === "External" });
        }
        return r;
      },
      agregaRel(ruta, id, tipo, destino, externo) {
        const n = P.relsDe(ruta), nuevo = `<Relationship Id="${id}" Type="${tipo}" Target="${x(destino)}"${externo ? ' TargetMode="External"' : ""}/>`;
        let t = P.texto(n);
        if (!t) t = XML + `<Relationships xmlns="${REL_NS}"></Relationships>`;
        const k = t.lastIndexOf("</");
        t = k < 0 ? t.replace(/<Relationships([^>]*)\/>\s*$/, `<Relationships$1>${nuevo}</Relationships>`) : t.slice(0, k) + nuevo + t.slice(k);
        P.pon(n, t);
      },
      quitaRel(ruta, id) {
        const n = P.relsDe(ruta), t = P.texto(n);
        if (t) P.pon(n, t.replace(new RegExp('<(?:\\w+:)?Relationship\\b[^>]*\\sId\\s*=\\s*["\']' + id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + '["\'][^>]*?/>'), ""));
      },
      /** Declara el tipo de contenido de una parte nueva. */
      tipo(parte, ct) {
        let t = P.texto(CTN) || "";
        if (t.includes(`PartName="/${parte}"`)) return;
        const k = t.lastIndexOf("</");
        P.pon(CTN, t.slice(0, k) + `<Override PartName="/${x(parte)}" ContentType="${ct}"/>` + t.slice(k));
      },
      extension(ext, ct) {
        let t = P.texto(CTN) || "";
        if (new RegExp('<Default\\b[^>]*Extension\\s*=\\s*["\']' + reEsc(ext) + '["\']', "i").test(t)) return;
        const k = t.indexOf(">", t.indexOf("<Types")) + 1;
        P.pon(CTN, t.slice(0, k) + `<Default Extension="${ext}" ContentType="${ct}"/>` + t.slice(k));
      },
      quitaTipo(parte) {
        const t = P.texto(CTN);
        if (t) P.pon(CTN, t.replace(new RegExp('<Override\\b[^>]*PartName\\s*=\\s*["\']/' + parte.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + '["\'][^>]*?/>'), ""));
      },
      zip() {
        const lista = orden.filter((n) => P.tiene(n)).map((n) => {
          if (nuevos.has(n)) {
            // Si la parte original empezaba con la marca UTF-8 (BOM), se conserva: al leerla como texto se
            // pierde, y la parte reescrita salía 3 bytes distinta aunque su contenido fuera el mismo.
            let v = nuevos.get(n);
            if (typeof v === "string" && v.charCodeAt(0) !== 0xFEFF && Z.tiene(n)) { const o = Z.bytes(n); if (o && o[0] === 0xEF && o[1] === 0xBB && o[2] === 0xBF) v = "\uFEFF" + v; }
            return { nombre: n, datos: v };
          }
          const d = Z.bytes(n), cr = d == null && Z.crudo ? Z.crudo(n) : null;
          return cr ? { nombre: n, crudo: cr } : { nombre: n, datos: d };
        }).filter((x) => x.crudo || x.datos != null);
        // [Content_Types].xml primero, como lo escribe Office
        const sale = nombresAscii(lista);
        sale.sort((a, b) => (a.nombre === CTN ? -1 : 0) - (b.nombre === CTN ? -1 : 0));
        return zip(sale);
      },
    };
    return P;
  }
  /**
   * Open XML (OPC, ECMA-376 parte 2) pide los nombres de las partes en ASCII: lo demás va como %XX de UTF-8. Hay
   * programas (LibreOffice) que guardan el nombre tal cual («media/Cortázar.wav») y entonces PowerPoint ya no abre el
   * archivo ni el validador de Microsoft encuentra la parte. Al descargar se corrige: el nombre de la parte y cada
   * referencia a ella (destinos de las relaciones y PartName de [Content_Types].xml). Lo que ya es ASCII no se toca.
   */
  function nombresAscii(lista) {
    const cod = (s) => s.replace(/[^\x00-\x7f]+/g, (m) => encodeURIComponent(m));
    const dec = (s) => { try { return decodeURIComponent(s); } catch (e) { return s; } };
    const cambia = new Map();
    lista.forEach((a) => { if (/[^\x00-\x7f]/.test(a.nombre)) cambia.set(a.nombre, cod(a.nombre)); });
    if (!cambia.size) return lista;
    const des = (s) => (NOM.origen && NOM.origen.desescapar ? NOM.origen.desescapar(s) : s.replace(/&amp;/g, "&"));
    const resolver = (dir, destino) => {
      if (destino.startsWith("/")) return destino.slice(1);
      const b = dir ? dir.split("/").filter(Boolean) : [];
      destino.split("/").forEach((s) => { if (s === "..") b.pop(); else if (s !== "." && s !== "") b.push(s); });
      return b.join("/");
    };
    const td = new TextDecoder();
    return lista.map((a) => {
      const nombre = cambia.get(a.nombre) || a.nombre;
      if (a.crudo || !(/\.rels$/.test(a.nombre) || a.nombre === "[Content_Types].xml")) return nombre === a.nombre ? a : Object.assign({}, a, { nombre });
      let t = typeof a.datos === "string" ? a.datos : td.decode(a.datos);
      const t0 = t;
      if (a.nombre === "[Content_Types].xml") {
        t = t.replace(/(\sPartName\s*=\s*")\/([^"]*)(")/g, (m, i, pn, f) => { const n = cambia.get(dec(des(pn))); return n ? i + "/" + n + f : m; });
      } else {
        // el directorio de la parte dueña de estas relaciones ("_rels/.rels" es la raíz del paquete)
        const dir = a.nombre === "_rels/.rels" ? "" : a.nombre.replace(/_rels\/[^/]+\.rels$/, "");
        t = t.replace(/(\sTarget\s*=\s*")([^"]*)(")/g, (m, i, tg, f) => {
          const destino = dec(des(tg));
          return cambia.has(resolver(dir, destino)) ? i + cod(destino).replace(/&/g, "&amp;") + f : m;
        });
      }
      return t === t0 && nombre === a.nombre ? a : Object.assign({}, a, { nombre, datos: t });
    });
  }

  /** Lo que hace falta del styles.xml de origen para agregarle estilos sin tocar los que ya tiene. */
  function estilosOrigen(xml) {
    const O = NOM.origen, raiz = O.hijosXml(xml).find((h) => O.local(h.n) === "styleSheet");
    if (!raiz) throw new Error("styles.xml sin raíz");
    const pref = raiz.n.includes(":") ? raiz.n.split(":")[0] + ":" : "";
    const sinPref = (s) => (pref ? s.replace(new RegExp("<(/?)" + reEsc(pref), "g"), "<$1") : s);
    const hs = O.hijosXml(xml, raiz.ci, raiz.cf), sec = (n) => hs.find((h) => O.local(h.n) === n) || null;
    const dentro = (s) => (s ? O.hijosXml(xml, s.ci, s.cf) : []);
    const fuentes = dentro(sec("fonts")).map((h) => sinPref(xml.slice(h.ci, h.cf)));
    const codigos = {};
    let maxNum = 163;
    dentro(sec("numFmts")).forEach((h) => { const id = +O.attr(xml, h, "numFmtId"), cod = O.attr(xml, h, "formatCode"); if (id > maxNum) maxNum = id; if (cod != null && codigos[cod] == null) codigos[cod] = id; });
    const xfs = dentro(sec("cellXfs")).map((h) => {
      const ks = O.hijosXml(xml, h.ci, h.cf), al = ks.find((k) => O.local(k.n) === "alignment");
      return { attrs: xml.slice(h.a, h.af), numFmtId: +O.attr(xml, h, "numFmtId") || 0, fontId: +O.attr(xml, h, "fontId") || 0, fillId: +O.attr(xml, h, "fillId") || 0, borderId: +O.attr(xml, h, "borderId") || 0,
        alin: al ? xml.slice(al.a, al.af).replace(/\s+$/, "") : "", otros: sinPref(ks.filter((k) => k !== al).map((k) => xml.slice(k.i, k.f)).join("")) };
    });
    return { raiz, pref, hs, sec, fuentes, nRellenos: dentro(sec("fills")).length, nBordes: dentro(sec("borders")).length, nDxfs: dentro(sec("dxfs")).length, xfs, codigos, maxNum };
  }
  function estilosConNuevos(xml, EO, fab) {
    const O = NOM.origen, pf = (s) => (EO.pref ? s.replace(/<(\/?)(?=[A-Za-z])/g, "<$1" + EO.pref) : s), piezas = [];
    const ORD = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];
    const agrega = (nombre, items) => {
      if (!items.length) return;
      const s = EO.sec(nombre), nuevo = pf(items.join(""));
      if (s) {
        const n = O.hijosXml(xml, s.ci, s.cf).length + items.length;
        piezas.push([s.i, s.f, `<${s.n}${quitaAttr(xml.slice(s.a, s.af), "count")} count="${n}">${xml.slice(s.ci, s.cf)}${nuevo}</${s.n}>`]);
      } else {
        const k = ORD.indexOf(nombre), sig = EO.hs.find((h) => ORD.indexOf(O.local(h.n)) > k), pos = sig ? sig.i : EO.raiz.cf;
        piezas.push([pos, pos, `<${EO.pref}${nombre} count="${items.length}">${nuevo}</${EO.pref}${nombre}>`]);
      }
    };
    agrega("numFmts", fab.numFmts); agrega("fonts", fab.fuentes); agrega("fills", fab.rellenos); agrega("borders", fab.bordes); agrega("cellXfs", fab.xfs); agrega("dxfs", fab.dxfs);
    return aplicarPiezas(xml, piezas);
  }

  const ORDEN_HOJA = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
  /** Las secciones de una hoja (hijos de <worksheet>) y cómo reemplazar o insertar una en su lugar. */
  function seccionesHoja(xml) {
    const O = NOM.origen, raiz = O.hijosXml(xml).find((h) => O.local(h.n) === "worksheet");
    if (!raiz) throw new Error("hoja sin raíz");
    const pref = raiz.n.includes(":") ? raiz.n.split(":")[0] + ":" : "";
    const hs = O.hijosXml(xml, raiz.ci, raiz.cf), piezas = [];
    return {
      raiz, pref, hs, piezas,
      pf: (s) => (pref ? s.replace(/<(\/?)(?=[A-Za-z])/g, "<$1" + pref) : s),
      una: (n) => hs.find((h) => O.local(h.n) === n) || null,
      todas: (n) => hs.filter((h) => O.local(h.n) === n),
      /** Reemplaza la sección (todas las de ese nombre) o la inserta donde toca. */
      pon(nombre, nuevo) {
        const v = hs.filter((h) => O.local(h.n) === nombre);
        if (v.length) { piezas.push([v[0].i, v[v.length - 1].f, nuevo]); return; }
        if (!nuevo) return;
        const k = ORDEN_HOJA.indexOf(nombre), sig = hs.find((h) => ORDEN_HOJA.indexOf(O.local(h.n)) > k), pos = sig ? sig.i : raiz.cf;
        piezas.push([pos, pos, nuevo]);
      },
      aplicar: () => aplicarPiezas(xml, piezas),
    };
  }

  /**
   * Devuelve el XML de la hoja de origen con SOLO lo que cambió respecto a lo importado (`base`),
   * o null si no cambió nada. Lo demás (estilos, textos enriquecidos, gráficas, filtros, vínculos,
   * formato condicional, celdas fuera de lo que NOM muestra) queda intacto.
   */
  function parcharHoja(xml, base, t, fab, ev) {
    const O = NOM.origen;
    const celdas = new Map(), filasAlto = new Set(), colsCamb = new Set(), filasEsq = new Set(), colsAncho = new Set();
    let cM = false, cPane = false, cCf = false, cDv = false, cProt = false, cVf = false, cEsc = false, cMg = false;
    const llaves = new Set(Object.keys(base));
    Object.keys(t).forEach((k) => llaves.add(k));
    llaves.forEach((k) => {
      if (mismoValor(base[k], t[k])) return;
      const i2 = k.indexOf(":");
      if (i2 > 0) {
        const p = k.slice(0, i2), r = k.slice(i2 + 1);
        if ((p === "c" || CELDA_FMT.includes(p)) && /^[A-Z]{1,3}\d+$/.test(r)) { (celdas.get(r) || celdas.set(r, new Set()).get(r)).add(p); }
        else if (p === "m") cM = true;
        else if (p === "ah") filasAlto.add(+r);
        else if (p === "oh" || p === "ol") { filasEsq.add(+r); filasEsq.add(+r + 1); }   // (y la de resumen de abajo: su «collapsed»)
        else if (p === "aw") { colsCamb.add(r); colsAncho.add(r); }
        else if (p === "ow" || p === "olc") { colsCamb.add(r); const q = F().parseRef(r + "1"); if (q) colsCamb.add(F().colLetra(q.c + 1)); }
        else if (p === "cf") cCf = true;
        else if (p === "dv") cDv = true;
      } else if (k === "fijasF" || k === "fijasC") cPane = true;
      else if (k === "vf") cVf = true;
      else if (k === "prot") cProt = true;
      else if (k === "esc") cEsc = true;
      else if (k === "mg") cMg = true;
    });
    if (!celdas.size && !filasAlto.size && !filasEsq.size && !colsCamb.size && !cM && !cPane && !cCf && !cDv && !cProt && !cVf && !cEsc && !cMg) return null;
    // Matrices dinámicas: si la hoja cambió, lo que derrama cada fórmula se vuelve a escribir con sus valores de hoy
    // (y se borra lo que ya no alcanza); la fórmula, si cambió el tamaño de lo que derrama (su «ref»).
    if (celdas.size && ev.derrame) {
      const zona = (txt) => { const [a, b] = String(txt || "").split(":").map((q) => F().parseRef(q)); return a && b ? { c1: a.c, f1: a.f, c2: b.c, f2: b.f } : null; };
      const anclas = new Set();
      Object.keys(base).forEach((k) => { if (k.startsWith("fa:")) anclas.add(k.slice(3)); });
      Object.keys(t).forEach((k) => { if (k.startsWith("c:") && String(t[k])[0] === "=") anclas.add(k.slice(2)); });
      const marca = (r) => (celdas.get(r) || celdas.set(r, new Set()).get(r)).add("c");
      anclas.forEach((r) => {
        const p = F().parseRef(r);
        if (!p) return;
        const d = String(t["c:" + r] || "")[0] === "=" ? ev.derrame(p.c, p.f) : null;
        const antes = zona(base["fa:" + r]), ahora = d && !d.bloqueado && d.filas * d.cols > 1 ? { c1: p.c, f1: p.f, c2: p.c + d.cols - 1, f2: p.f + d.filas - 1 } : null;
        if (!antes && !ahora) return;
        [antes, ahora].forEach((z) => { if (z) for (let ff = z.f1; ff <= z.f2; ff++) for (let cc = z.c1; cc <= z.c2; cc++) { const q = F().ref(cc, ff); if ((cc !== p.c || ff !== p.f) && String(t["c:" + q] || "") === "") marca(q); } });
        if (!antes || !ahora || antes.c2 !== ahora.c2 || antes.f2 !== ahora.f2) marca(r);
      });
    }
    const S = seccionesHoja(xml), pf = S.pf;
    const tiene = (v) => v != null && v !== "" && v !== false;

    if (celdas.size || filasAlto.size || filasEsq.size) {
      const sd = S.una("sheetData");
      const rows = sd ? O.hijosXml(xml, sd.ci, sd.cf).filter((h) => O.local(h.n) === "row") : [];
      // Fórmulas compartidas: si cambia la celda "maestra", las que dependen de ella se escriben completas.
      if (sd && xml.indexOf('t="shared"', sd.ci) > 0) {
        const re = /<(?:\w+:)?c\b([^>]*)>\s*<(?:\w+:)?f\b([^>]*?)\/?>/g, maestras = {}, siguen = {};
        re.lastIndex = sd.ci;
        let m;
        while ((m = re.exec(xml)) && m.index < sd.cf) {
          if (!/t="shared"/.test(m[2])) continue;
          const r = (/\sr="([A-Z]+\d+)"/.exec(m[1]) || [])[1], si = (/\ssi="(\d+)"/.exec(m[2]) || [])[1];
          if (!r || si == null) continue;
          if (/\sref="/.test(m[2])) maestras[r] = si; else (siguen[si] = siguen[si] || []).push(r);
        }
        [...celdas.keys()].forEach((r) => { if (maestras[r] != null && celdas.get(r).has("c")) (siguen[maestras[r]] || []).forEach((q) => (celdas.get(q) || celdas.set(q, new Set()).get(q)).add("c")); });
      }
      const porFila = new Map();
      celdas.forEach((ks, ref) => { const p = F().parseRef(ref); if (!p) return; (porFila.get(p.f + 1) || porFila.set(p.f + 1, new Map()).get(p.f + 1)).set(p.c, { ref, ks }); });
      filasAlto.forEach((n) => { if (n > 0 && !porFila.has(n)) porFila.set(n, new Map()); });
      filasEsq.forEach((n) => { if (n > 0 && !porFila.has(n)) porFila.set(n, new Map()); });
      const altoAttr = (n) => { const h = +t["ah:" + n] || 0; return h ? ` ht="${Math.round(h * 0.75 * 100) / 100}" customHeight="1"` : ""; };
      const celdaParchada = (ref, c, f, ks, oc) => {
        const sOrig = oc ? +(O.attr(xml, oc, "s") || 0) : 0;
        const camb = new Set([...ks].filter((k) => k !== "c"));
        if (!oc) CELDA_FMT.forEach((k) => { if (tiene(t[k + ":" + ref])) camb.add(k); });
        const crudo = t["c:" + ref] == null ? "" : String(t["c:" + ref]);
        // una fecha recién escrita en una celda sin formato de número se muestra como fecha
        const xfO = fab.xfOrigen(sOrig);
        const fecha = ks.has("c") && !tiene(t["f:" + ref]) && (!xfO || !xfO.numFmtId) && (crudo.startsWith("=") ? ev.valor(c, f) instanceof F().FechaF : fab.esFechaLit(crudo)) ? "dd/mm/yyyy" : null;
        const s = camb.size || fecha ? fab.estiloDesde(t, ref, sOrig, camb, fecha) : sOrig;
        if (!ks.has("c") && oc) {
          // mismo valor (con su texto enriquecido o su fórmula original): solo cambia el estilo
          const attrs = quitaAttr(xml.slice(oc.a, oc.af), "s") + (s ? ` s="${s}"` : "");
          return oc.ci === oc.f ? `<${oc.n}${attrs}/>` : `<${oc.n}${attrs}>${xml.slice(oc.ci, oc.cf)}</${oc.n}>`;
        }
        if (crudo === "") { const hijo = fab.celda(t, ev, ref, c, f, s); return hijo.includes("<v>") ? pf(hijo) : s ? pf(`<c r="${ref}" s="${s}"/>`) : ""; }
        return pf(fab.celda(t, ev, ref, c, f, s));
      };
      const filaNueva = (n) => {
        const cs = [...porFila.get(n).entries()].sort((a, b) => a[0] - b[0]).map(([c, v]) => celdaParchada(v.ref, c, n - 1, v.ks, null)).join("");
        const alto = altoAttr(n) + (filasEsq.has(n) ? esqFila(t, n) : "");
        return cs || alto ? pf(`<row r="${n}"${alto}>`) + cs + pf("</row>") : "";
      };
      const filaParchada = (row, n, cambios) => {
        const orig = new Map(), otros = [];
        let col = 0;
        O.hijosXml(xml, row.ci, row.cf).forEach((c) => {
          if (O.local(c.n) !== "c") { otros.push(xml.slice(c.i, c.f)); return; }
          const r = O.attr(xml, c, "r"), p = r ? F().parseRef(r) : null, k = p ? p.c : col;
          col = k + 1; orig.set(k, c);
        });
        const columnas = [...new Set([...orig.keys(), ...cambios.keys()])].sort((a, b) => a - b);
        let dentro = "";
        columnas.forEach((c) => { const cb = cambios.get(c), oc = orig.get(c); dentro += cb ? celdaParchada(cb.ref, c, n - 1, cb.ks, oc || null) : xml.slice(oc.i, oc.f); });
        let attrs = quitaAttr(xml.slice(row.a, row.af), "spans");
        if (filasAlto.has(n)) attrs = quitaAttr(quitaAttr(attrs, "ht"), "customHeight") + altoAttr(n);
        if (filasEsq.has(n)) attrs = quitaAttr(quitaAttr(quitaAttr(attrs, "hidden"), "outlineLevel"), "collapsed") + esqFila(t, n);
        if (!/\sr\s*=/.test(attrs)) attrs = ` r="${n}"` + attrs;
        return `<${row.n}${attrs}>${dentro}${otros.join("")}</${row.n}>`;
      };
      const pend = [...porFila.keys()].sort((a, b) => a - b);
      let pi = 0, nFila = 0, sale = "";
      rows.forEach((row) => {
        const rn = +O.attr(xml, row, "r") || nFila + 1;
        nFila = rn;
        while (pi < pend.length && pend[pi] < rn) sale += filaNueva(pend[pi++]);
        if (pi < pend.length && pend[pi] === rn) { sale += filaParchada(row, rn, porFila.get(rn)); pi++; }
        else sale += xml.slice(row.i, row.f);
      });
      while (pi < pend.length) sale += filaNueva(pend[pi++]);
      if (sd) S.piezas.push([sd.i, sd.f, `<${sd.n}${xml.slice(sd.a, sd.af)}>${sale}</${sd.n}>`]);
      else S.pon("sheetData", pf("<sheetData>") + sale + pf("</sheetData>"));
      if (celdas.size) S.pon("dimension", "");
    }
    // (con el esquema cambiado, sheetFormatPr dice hasta qué nivel llega: Excel dibuja así sus botones)
    if (filasEsq.size || [...colsCamb].some((l) => !colsAncho.has(l))) {
      const fp = S.una("sheetFormatPr"), nv = nivelesEsq(t);
      const esqA = (nv.f ? ` outlineLevelRow="${nv.f}"` : "") + (nv.c ? ` outlineLevelCol="${nv.c}"` : "");
      if (fp) S.piezas.push([fp.i, fp.f, `<${fp.n}${quitaAttr(quitaAttr(xml.slice(fp.a, fp.af), "outlineLevelRow"), "outlineLevelCol").replace(/\/$/, "")}${esqA}/>`]);
      else if (esqA) S.pon("sheetFormatPr", pf(`<sheetFormatPr defaultRowHeight="15"${esqA}/>`));
    }
    if (colsCamb.size) {
      const sec = S.una("cols");
      let rangos = (sec ? O.hijosXml(xml, sec.ci, sec.cf) : []).map((h) => ({ min: +O.attr(xml, h, "min"), max: +O.attr(xml, h, "max"), attrs: quitaAttr(quitaAttr(xml.slice(h.a, h.af), "min"), "max") }));
      colsCamb.forEach((letra) => {
        const p = F().parseRef(letra + "1");
        if (!p) return;
        const n = p.c + 1, w = +t["aw:" + letra] || 0, nuevos = [], ancho = colsAncho.has(letra), esq = esqCol(t, letra);
        let puesto = false;
        rangos.forEach((r) => {
          if (n < r.min || n > r.max) { nuevos.push(r); return; }
          if (r.min < n) nuevos.push({ min: r.min, max: n - 1, attrs: r.attrs });
          // (el ancho, solo si cambió; el esquema —oculta, nivel, contraída—, siempre el de NOM)
          let a = quitaAttr(quitaAttr(quitaAttr(r.attrs, "hidden"), "outlineLevel"), "collapsed");
          if (ancho) { a = quitaAttr(quitaAttr(a, "width"), "customWidth"); if (w) a += ` width="${px2ancho(w)}" customWidth="1"`; }
          a += esq;
          if (a.trim()) nuevos.push({ min: n, max: n, attrs: /\swidth\s*=/.test(a) ? a : a + ' width="9.14"' });
          if (r.max > n) nuevos.push({ min: n + 1, max: r.max, attrs: r.attrs });
          puesto = true;
        });
        if (!puesto && (w || esq)) nuevos.push({ min: n, max: n, attrs: ` width="${w ? px2ancho(w) : "9.14"}"${w ? ' customWidth="1"' : ""}${esq}` });
        rangos = nuevos.sort((a, b) => a.min - b.min);
      });
      S.pon("cols", rangos.length ? pf(`<cols>${rangos.map((r) => `<col min="${r.min}" max="${r.max}"${r.attrs}/>`).join("")}</cols>`) : "");
    }
    if (cM) {
      const sec = S.una("mergeCells");
      // las combinadas que quedaron fuera de lo importado (más allá del límite) se conservan
      const fuera = (sec ? O.hijosXml(xml, sec.ci, sec.cf) : []).filter((h) => { const a = F().parseRef(String(O.attr(xml, h, "ref") || "").split(":")[0]); return a && (a.f >= LIM_F || a.c >= LIM_C); }).map((h) => xml.slice(h.i, h.f));
      const todas = fab.combinadas(t).map(pf).concat(fuera);
      S.pon("mergeCells", todas.length ? pf(`<mergeCells count="${todas.length}">`) + todas.join("") + pf("</mergeCells>") : "");
    }
    if (cVf && !cPane) {
      // «mostrar fórmulas»: solo el atributo de la vista de la hoja
      const sv = S.una("sheetViews"), v1 = sv ? O.hijosXml(xml, sv.ci, sv.cf).find((h) => O.local(h.n) === "sheetView") : null;
      if (v1) {
        const attrs = quitaAttr(xml.slice(v1.a, v1.af).replace(/\/$/, ""), "showFormulas") + (t.vf ? ' showFormulas="1"' : "");
        S.piezas.push([v1.i, v1.f, v1.ci === v1.f ? `<${v1.n}${attrs}/>` : `<${v1.n}${attrs}>${xml.slice(v1.ci, v1.cf)}</${v1.n}>`]);
      } else if (t.vf) S.pon("sheetViews", pf('<sheetViews><sheetView showFormulas="1" workbookViewId="0"/></sheetViews>'));
    }
    if (cPane) {
      const sv = S.una("sheetViews"), pane = fab.panel(t);
      const v1 = sv ? O.hijosXml(xml, sv.ci, sv.cf).find((h) => O.local(h.n) === "sheetView") : null;
      if (v1) {
        const resto = O.hijosXml(xml, v1.ci, v1.cf).filter((h) => !["pane", "selection"].includes(O.local(h.n))).map((h) => xml.slice(h.i, h.f)).join("");
        const attrs = cVf ? quitaAttr(xml.slice(v1.a, v1.af).replace(/\/$/, ""), "showFormulas") + (t.vf ? ' showFormulas="1"' : "") : xml.slice(v1.a, v1.af).replace(/\/$/, "");
        S.piezas.push([v1.i, v1.f, `<${v1.n}${attrs}>${pf(pane)}${resto}</${v1.n}>`]);
      } else if (pane) S.pon("sheetViews", pf(`<sheetViews><sheetView${t.vf ? ' showFormulas="1"' : ""} workbookViewId="0">${pane}</sheetView></sheetViews>`));
    }
    if (cCf) {
      const viejas = S.todas("conditionalFormatting").map((h) => xml.slice(h.i, h.f)).join("");
      const nuevas = Object.keys(t).filter((k) => k.startsWith("cf:") && !mismoValor(base[k], t[k])).map((k) => fab.regla(t[k])).join("");
      S.pon("conditionalFormatting", viejas + pf(nuevas));
    }
    if (cDv) {
      const sec = S.una("dataValidations"), lista = [];
      (sec ? O.hijosXml(xml, sec.ci, sec.cf) : []).forEach((h, k) => {
        const kb = "dv:x" + k;
        if (base[kb] == null || mismoValor(base[kb], t[kb])) lista.push(xml.slice(h.i, h.f));   // no importada, o importada y sin cambios
      });
      Object.keys(t).filter((k) => k.startsWith("dv:") && tiene(t[k]) && !mismoValor(base[k], t[k])).forEach((k) => { const v = fab.validacion(t[k]); if (v) lista.push(pf(v)); });
      S.pon("dataValidations", lista.length ? pf(`<dataValidations count="${lista.length}">`) + lista.join("") + pf("</dataValidations>") : "");
    }
    // la protección de la hoja (se puso, se cambió o se quitó aquí)
    if (cProt) S.pon("sheetProtection", t.prot && typeof t.prot === "object" ? pf(xmlProtHoja(t.prot)) : "");
    // los escenarios (se agregaron, se cambiaron o se quitaron aquí)
    if (cEsc) S.pon("scenarios", pf(xmlEscenarios(t.esc)));
    // los minigráficos (se agregaron, se cambiaron o se quitaron aquí): su ext en el extLst, sin tocar lo demás de él
    if (cMg) { const sec = S.una("extLst"), dentro = extConMinigraficos(sec ? xml.slice(sec.ci, sec.cf) : "", t.mg); S.pon("extLst", dentro ? pf("<extLst>") + dentro + pf("</extLst>") : ""); }
    return S.aplicar();
  }

  /** Hoja completa armada desde el modelo (documento nuevo, u hoja agregada a un libro importado). */
  function hojaDesdeCero(t, ev, fab, estiloDe, primera, dibujo) {
    const porFila = {}, celdasCon = new Set();
    Object.keys(t).forEach((k) => {
      const i2 = k.indexOf(":");
      if (i2 < 1) return;
      const p = k.slice(0, i2), ref = k.slice(i2 + 1);
      if (!["c", "n", "i", "u", "k", "f", "r", "al", "aj", "bd", "tl", "m"].includes(p) || !/^[A-Z]{1,3}\d+$/.test(ref)) return;
      if (t[k] === "" || t[k] == null || t[k] === false) return;
      celdasCon.add(ref);
    });
    // lo que derraman las fórmulas de matrices (Excel guarda esos valores en sus celdas)
    if (ev.derrame) [...celdasCon].forEach((ref) => {
      if (String(t["c:" + ref] || "")[0] !== "=") return;
      const p = F().parseRef(ref), d = ev.derrame(p.c, p.f);
      if (d && !d.bloqueado) for (let f = p.f; f < p.f + d.filas; f++) for (let c = p.c; c < p.c + d.cols; c++) celdasCon.add(F().ref(c, f));
    });
    celdasCon.forEach((ref) => { const r = F().parseRef(ref); (porFila[r.f] = porFila[r.f] || []).push({ c: r.c, ref }); });
    Object.keys(t).forEach((k) => { if (k.startsWith("ah:") && t[k]) { const f = +k.slice(3) - 1; porFila[f] = porFila[f] || []; } });
    // (las filas agrupadas u ocultas, y la de resumen de abajo de cada grupo, van aunque no tengan celdas)
    Object.keys(t).forEach((k) => { if ((k.startsWith("ol:") || k.startsWith("oh:")) && t[k]) { const f = +k.slice(3) - 1; porFila[f] = porFila[f] || []; porFila[f + 1] = porFila[f + 1] || []; } });
    const filas = Object.keys(porFila).map(Number).sort((a, b) => a - b).map((f) => {
      const celdas = porFila[f].sort((a, b) => a.c - b.c).map((c) => {
        const crudo = t["c:" + c.ref] == null ? "" : String(t["c:" + c.ref]);
        const esFecha = crudo.startsWith("=") ? ev.valor(c.c, f) instanceof F().FechaF : fab.esFechaLit(crudo);
        return fab.celda(t, ev, c.ref, c.c, f, estiloDe(c.ref, esFecha));
      }).join("");
      const h = +t["ah:" + (f + 1)] || 0, esq = esqFila(t, f + 1);
      if (!celdas && !h && !esq) return "";
      return `<row r="${f + 1}"${h ? ` ht="${Math.round(h * 0.75 * 100) / 100}" customHeight="1"` : ""}${esq}>${celdas}</row>`;
    }).join("");
    const ncols = Math.max(t.cols || 10, 1);
    // Solo las columnas con ancho propio (las demás usan el ancho por omisión): archivo más chico.
    let cols = "";
    for (let c = 0; c < ncols + 1; c++) {
      const L = F().colLetra(c), w = +t["aw:" + L] || 0, esq = esqCol(t, L);
      if (w || esq) cols += `<col min="${c + 1}" max="${c + 1}" width="${w ? px2ancho(w) : "9.14"}"${w ? ' customWidth="1"' : ""}${esq}/>`;
    }
    const nvE = nivelesEsq(t), esqFmt = (nvE.f ? ` outlineLevelRow="${nvE.f}"` : "") + (nvE.c ? ` outlineLevelCol="${nvE.c}"` : "");
    const pane = fab.panel(t);
    const vf = (t.vf ? ' showFormulas="1"' : "") + (t.xl && t.xl.rejilla === false ? ' showGridLines="0"' : "");   // (sin cuadrícula: el resumen de escenarios)
    const vista = pane ? `<sheetView${vf} workbookViewId="0"${primera ? ' tabSelected="1"' : ""}>${pane}</sheetView>` : `<sheetView${vf} workbookViewId="0"${primera ? ' tabSelected="1"' : ""}/>`;
    const combinadas = fab.combinadas(t);
    const reglas = Object.keys(t).filter((k) => k.startsWith("cf:")).map((k) => fab.regla(t[k])).join("");
    const validaciones = Object.keys(t).filter((k) => k.startsWith("dv:")).map((k) => fab.validacion(t[k])).filter(Boolean);
    return XML + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews>${vista}</sheetViews><sheetFormatPr defaultColWidth="16.43" defaultRowHeight="15"${esqFmt}/>${cols ? `<cols>${cols}</cols>` : ""}<sheetData>${filas}</sheetData>${t.prot && typeof t.prot === "object" ? xmlProtHoja(t.prot) : ""}${xmlEscenarios(t.esc)}${combinadas.length ? `<mergeCells count="${combinadas.length}">${combinadas.join("")}</mergeCells>` : ""}${reglas}${validaciones.length ? `<dataValidations count="${validaciones.length}">${validaciones.join("")}</dataValidations>` : ""}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${dibujo || ""}${Array.isArray(t.mg) && t.mg.length ? `<extLst>${xmlMinigraficos(t.mg)}</extLst>` : ""}</worksheet>`;
  }
  const CT_HOJA = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
  const CT_DIBUJO = "application/vnd.openxmlformats-officedocument.drawing+xml", CT_GRAFICA = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
  const T_DIBUJO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", T_GRAFICA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
  const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  /** .xlsx a partir del archivo del que vino el documento: se cambia solo lo editado. */
  function xlsxConOrigen(doc, og) {
    const O = NOM.origen, pq = paqueteDe(og), hojasBase = {};
    // Lo que se importó de cada hoja se vuelve a sacar del original: la diferencia es lo que editó la persona.
    // (con la misma versión del importador con que se importó: si no, un cambio de lectura parecería una edición)
    modeloXlsx(og.zip, VERSION_XL).hojas.forEach((h) => { hojasBase[h.campos.ox] = h.campos; });
    const porVersion = { [VERSION_XL]: hojasBase };
    const baseDe = (t, ox) => {
      const v = Math.max(1, Math.min(VERSION_XL, +t.xlv || 1));
      if (!porVersion[v]) { porVersion[v] = {}; modeloXlsx(og.zip, v).hojas.forEach((h) => { porVersion[v][h.campos.ox] = h.campos; }); }
      return porVersion[v][ox] || hojasBase[ox];
    };
    const tablas = F().hojasDe(doc);
    if (!tablas.length) throw new Error("Este documento no tiene ninguna tabla");
    const LIBRO = "xl/workbook.xml";
    const relsL = pq.rels(LIBRO), relEst = relsL.find((r) => /\/styles$/.test(r.tipo || ""));
    const estRuta = relEst ? pq.resolver(LIBRO, relEst.destino) : "xl/styles.xml";
    const estTxt = pq.texto(estRuta), libroTxt = pq.texto(LIBRO);
    if (!estTxt || !libroTxt) throw new Error("al archivo original le faltan partes");
    const EO = estilosOrigen(estTxt), fab = fabricaXlsx(doc, { origen: EO });
    fab.xfOrigen = (s) => EO.xfs[s] || null;
    // matrices dinámicas: el cm que ya trae el libro; si trae metadatos sin él, se guardan como fórmulas de matriz
    const relMeta = relsL.find((r) => /\/sheetMetadata$/.test(r.tipo || ""));
    const metaRuta = relMeta ? pq.resolver(LIBRO, relMeta.destino) : null;
    fab.cmDA = metaRuta && pq.tiene(metaRuta) ? cmDeMetadatos(pq.texto(metaRuta) || "") : null;
    const usados = new Set(), nombres = [];
    let cambio = false, nNueva = 0, nGraf = 0, nDib = 0, nImgXl = 0, infoTablas = null;
    const cachesNuevas = [];   // (tablas dinámicas hechas en NOM: [{ id, rid }] para las pivotCaches del libro)
    const tiene = (v) => v != null && v !== "" && v !== false;
    const planes = tablas.map((t, i) => {
      let n = F().nombreDeHoja(doc, t).replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || `Hoja${i + 1}`;
      while (nombres.includes(n)) n = n.slice(0, 28) + " " + (i + 1);
      nombres.push(n);
      const ox = t.ox && hojasBase[t.ox] && pq.tiene(t.ox) && !usados.has(t.ox) ? t.ox : null;
      if (ox) usados.add(ox);
      return { t, ox, nombre: n, ruta: ox };
    });
    planes.forEach((p, i) => {
      const t = p.t, ev = F().evaluadorDeTabla(doc, t);
      if (p.ox) {
        const r = parcharHoja(pq.texto(p.ox), baseDe(t, p.ox), t, fab, ev);
        if (r != null) { pq.pon(p.ox, r); cambio = true; }
      } else {
        // hoja creada en NOM dentro de un libro importado
        p.ruta = `xl/worksheets/nomhoja${++nNueva}.xml`;
        const estiloDe = (ref, esFecha) => { const camb = new Set(CELDA_FMT.filter((k) => tiene(t[k + ":" + ref]))); const fecha = esFecha && !tiene(t["f:" + ref]) ? "dd/mm/yyyy" : null; return camb.size || fecha ? fab.estiloDesde(t, ref, 0, camb, fecha) : 0; };
        pq.pon(p.ruta, hojaDesdeCero(t, ev, fab, estiloDe, i === 0, ""));
        cambio = true;
      }
      // tablas con formato hechas en NOM: una parte nueva cada una, junto a las que ya traía la hoja
      const los = tablasLoDe(t);
      if (los.length) {
        if (!infoTablas) {
          infoTablas = { id: 0, nombres: new Set() };
          pq.nombres().filter((n) => /^xl\/tables\/[^/]+\.xml$/i.test(n)).forEach((n) => {
            const tx = pq.texto(n) || "", id = /<(?:\w+:)?table\b[^>]*?\sid\s*=\s*"(\d+)"/.exec(tx), nm = /\sdisplayName\s*=\s*"([^"]*)"/.exec(tx);
            if (id) infoTablas.id = Math.max(infoTablas.id, +id[1]);
            if (nm) infoTablas.nombres.add(nm[1].toLowerCase());
          });
        }
        const rids = los.map((lo) => {
          const id = ++infoTablas.id, ruta = `xl/tables/nomtabla${id}.xml`, rid = "rIdNomT" + id;
          pq.pon(ruta, xmlTablaXl(lo, id, nombreTablaXl(lo.n, infoTablas.nombres), columnasTablaXl(lo, t, ev))); pq.tipo(ruta, CT_TABLA_XL);
          pq.agregaRel(p.ruta, rid, T_TABLA_XL, "/" + ruta);
          return rid;
        });
        const hx = pq.texto(p.ruta), S = seccionesHoja(hx), tp = S.una("tableParts");
        const antes = tp ? (hx.slice(tp.i, tp.f).match(/<(?:\w+:)?tablePart\b[^>]*\/>/g) || []) : [];
        const nuevos = rids.map((r) => S.pf(`<tablePart xmlns:r="${NS_R}" r:id="${r}"/>`));
        S.pon("tableParts", S.pf(`<tableParts count="${antes.length + nuevos.length}">`) + antes.join("") + nuevos.join("") + S.pf("</tableParts>"));
        pq.pon(p.ruta, S.aplicar());
        cambio = true;
      }
      // tablas dinámicas hechas en NOM: sus partes nuevas, enlazadas desde la hoja; su caché va al libro más abajo
      dinamicasDe(doc, t).forEach((r) => {
        const k = cachesNuevas.length + 1, rt = `xl/pivotTables/nomdinamica${k}.xml`, rc = `xl/pivotCache/nomcache${k}.xml`, rr = `xl/pivotCache/nomregistros${k}.xml`;
        cachesNuevas.push({ tabla: rt, cache: rc, partes: r.partes });
        pq.pon(rc, r.partes.cache); pq.tipo(rc, CT_PV_CACHE);
        pq.pon(rr, r.partes.registros); pq.tipo(rr, CT_PV_REG); pq.agregaRel(rc, "rId1", T_PV_REG, "/" + rr);
        pq.pon(rt, r.partes.tabla); pq.tipo(rt, CT_PV_TABLA); pq.agregaRel(rt, "rId1", T_PV_CACHE, "/" + rc);
        pq.agregaRel(p.ruta, "rIdNomP" + k, T_PV_TABLA, "/" + rt);
        cambio = true;
      });
      // gráficas hechas en NOM: se agregan al dibujo de la hoja (o se le crea uno)
      const gs = doc.vivos().filter((g) => g.tipo === "grafica" && g.hoja === t.id);
      const ims = doc.vivos().filter((b) => b.tipo === "imagen" && b.hoja === t.id && b.src);
      if (!gs.length && !ims.length) return;
      const anclas = [], relsG = [];
      let fila0 = 1;
      const colIni = Math.min(Math.max(t.cols || 10, 1) + 1, 20);
      gs.forEach((g) => {
        const xmlG = graficaXlsx(g, p.nombre, t, ev);
        if (!xmlG) return;
        const ng = ++nGraf, ruta = `xl/charts/nomgrafica${ng}.xml`, rid = "rIdNomG" + ng;
        pq.pon(ruta, xmlG); pq.tipo(ruta, CT_GRAFICA);
        relsG.push([rid, ruta]);
        anclas.push(fab.anclaGrafica(rid, ng, colIni, fila0));
        fila0 += 18;
      });
      ims.forEach((b) => {
        const im = imagen(b.src);
        if (!im) return;
        const ni = ++nImgXl, ruta = `xl/media/nomimagen${ni}.${im.ext}`, rid = "rIdNomI" + ni;
        pq.pon(ruta, im.bytes); pq.tipo(ruta, ctImagen(im.ext));
        relsG.push([rid, ruta, T_IMAGEN_XL]);
        anclas.push(anclaImagenXl(rid, ni, b));
      });
      if (!anclas.length) return;
      cambio = true;
      let hx = pq.texto(p.ruta), dRuta = null;
      const dEl = /<(?:\w+:)?drawing\b[^>]*?\b(?:\w+:)?id\s*=\s*"([^"]+)"/.exec(hx);
      if (dEl) { const rel = pq.rels(p.ruta).find((r) => r.id === dEl[1]); if (rel) dRuta = pq.resolver(p.ruta, rel.destino); }
      if (dRuta && pq.tiene(dRuta)) {
        const d = pq.texto(dRuta), k = d.lastIndexOf("</");
        if (k > 0) pq.pon(dRuta, d.slice(0, k) + anclas.join("") + d.slice(k));
        else dRuta = null;
      } else dRuta = null;
      if (!dRuta) {
        dRuta = `xl/drawings/nomdibujo${++nDib}.xml`;
        pq.pon(dRuta, XML + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}">${anclas.join("")}</xdr:wsDr>`);
        pq.tipo(dRuta, CT_DIBUJO);
        pq.agregaRel(p.ruta, "rIdNomD", T_DIBUJO, "/" + dRuta);
        const S = seccionesHoja(hx);
        S.pon("drawing", `<${S.pref}drawing xmlns:r="${NS_R}" r:id="rIdNomD"/>`);
        pq.pon(p.ruta, S.aplicar());
      }
      relsG.forEach(([rid, ruta, tp]) => pq.agregaRel(dRuta, rid, tp || T_GRAFICA, "/" + ruta));
    });

    // ---- el libro: nombres, hojas nuevas o borradas y su orden
    const raizL = O.hijosXml(libroTxt).find((h) => O.local(h.n) === "workbook");
    const hsL = O.hijosXml(libroTxt, raizL.ci, raizL.cf), prefL = raizL.n.includes(":") ? raizL.n.split(":")[0] + ":" : "";
    const secHojas = hsL.find((h) => O.local(h.n) === "sheets"), piezasL = [];
    const info = O.hijosXml(libroTxt, secHojas.ci, secHojas.cf).map((h) => {
      const txt = libroTxt.slice(h.i, h.f), rid = (/\s(?:\w+:)id\s*=\s*"([^"]*)"/.exec(txt) || [])[1], rel = relsL.find((r) => r.id === rid);
      return { txt, rid, ruta: rel ? pq.resolver(LIBRO, rel.destino) : null, nombre: O.attr(libroTxt, h, "name"), id: +O.attr(libroTxt, h, "sheetId") || 0 };
    });
    const renombra = (s, n) => (n === s.nombre ? s.txt : s.txt.replace(/\sname\s*=\s*("[^"]*"|'[^']*')/, ` name="${x(n)}"`));
    const importadas = info.filter((s) => s.ruta && hojasBase[s.ruta]);
    const ahora = planes.filter((p) => p.ox).map((p) => p.ox);
    const mismaLista = importadas.length === ahora.length && importadas.every((s, i) => s.ruta === ahora[i]) && !planes.some((p) => !p.ox);
    let lista;
    if (mismaLista) {
      lista = info.map((s) => { const p = planes.find((q) => q.ox === s.ruta); if (p && p.nombre !== s.nombre) cambio = true; return p ? renombra(s, p.nombre) : s.txt; });
    } else {
      cambio = true;
      let maxId = info.reduce((m, s) => Math.max(m, s.id), 0), k = 0;
      const prefR = (/\s(\w+):id\s*=/.exec(info[0] ? info[0].txt : "") || [0, "r"])[1];
      lista = planes.map((p) => {
        if (p.ox) return renombra(info.find((s) => s.ruta === p.ox), p.nombre);
        const rid = "rIdNomH" + ++k;
        pq.agregaRel(LIBRO, rid, T.hoja, "/" + p.ruta); pq.tipo(p.ruta, CT_HOJA);
        return `<${prefL}sheet name="${x(p.nombre)}" sheetId="${++maxId}" ${prefR}:id="${rid}"/>`;
      });
      info.filter((s) => !(s.ruta && hojasBase[s.ruta])).forEach((s) => lista.push(s.txt));   // hojas de gráfico y otras que no se importan
      importadas.filter((s) => !usados.has(s.ruta)).forEach((s) => { pq.quita(s.ruta); pq.quita(pq.relsDe(s.ruta)); pq.quitaTipo(s.ruta); pq.quitaRel(LIBRO, s.rid); });
      // Los nombres definidos "de una hoja" y la pestaña activa van por posición: ya no corresponden (ver abajo).
      const bv = hsL.find((h) => O.local(h.n) === "bookViews");
      if (bv) piezasL.push([bv.i, bv.f, libroTxt.slice(bv.i, bv.f).replace(/\s(activeTab|firstSheet)\s*=\s*"[^"]*"/g, "")]);
    }
    piezasL.push([secHojas.ci, secHojas.cf, lista.join("")]);
    // Nombres definidos: si los de NOM (Administrador de nombres) ya no son los del original, se escriben los de NOM; los de
    // una hoja (si las hojas siguen en su lugar) y los de Excel (_xlnm: área de impresión, filtros…) se quedan tal cual.
    {
      const dn = hsL.find((h) => O.local(h.n) === "definedNames");
      const hijosDn = dn ? O.hijosXml(libroTxt, dn.ci, dn.cf) : [];
      const norm = (t) => String(t || "").replace(/^=/, "").replace(/_xlfn\.|_xlpm\.|_xlws\./g, "").replace(/\s+/g, "");
      const orig = {};
      hijosDn.forEach((h) => {
        const nm = O.attr(libroTxt, h, "name") || "", txt = libroTxt.slice(h.i, h.f);
        if (/^_xlnm\./i.test(nm) || O.attr(libroTxt, h, "localSheetId") != null || /\shidden\s*=\s*"(1|true)"/.test(txt.slice(0, txt.indexOf(">")))) return;
        orig[nm.toLowerCase()] = norm(txt.replace(/^<[^>]*>|<\/[^>]*>$/g, "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&amp;/g, "&"));
      });
      const mNom = doc.valor("_doc", "nombres");
      const nom = mNom && typeof mNom === "object" ? mNom : null;
      // (un libro importado antes de los nombres —versión 7— no los tiene en «_doc»: los del original se quedan)
      let cambianNombres = false;
      if (nom) {
        const ks = Object.keys(nom), ko = Object.keys(orig);
        cambianNombres = ks.length !== ko.length || ks.some((k) => orig[k] == null || norm(formulaExcel(String(nom[k].d || "").startsWith("=") ? nom[k].d : "=" + nom[k].d, 7)) !== orig[k]);
      }
      if (cambianNombres || (!mismaLista && hijosDn.some((h) => O.attr(libroTxt, h, "localSheetId") != null))) {
        const pref = dn ? (dn.n.includes(":") ? dn.n.split(":")[0] + ":" : "") : prefL;
        const quedan = hijosDn.filter((h) => {
          const nm = O.attr(libroTxt, h, "name") || "", local = O.attr(libroTxt, h, "localSheetId") != null, txt = libroTxt.slice(h.i, h.f);
          if (local) return mismaLista;
          if (/^_xlnm\./i.test(nm) || /\shidden\s*=\s*"(1|true)"/.test(txt.slice(0, txt.indexOf(">")))) return true;
          return !cambianNombres;   // (los del libro: si cambiaron, se escriben los de NOM)
        }).map((h) => libroTxt.slice(h.i, h.f)).join("");
        const nuevos = cambianNombres ? Object.values(nom).filter((o) => o && o.n && o.d != null).map((o) => definedNameXml(o, pref)).join("") : "";
        const todo = quedan + nuevos;
        if (dn) piezasL.push([dn.i, dn.f, todo ? `<${dn.n}>${todo}</${dn.n}>` : ""]);
        else if (todo) {
          // (va después de «sheets» y antes de calcPr, en el orden del esquema)
          const ORDN = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences"];
          const sig = hsL.find((h) => ORDN.indexOf(O.local(h.n)) < 0), pos = sig ? sig.i : raizL.cf;
          piezasL.push([pos, pos, `<${pref}definedNames>${todo}</${pref}definedNames>`]);
        }
        cambio = true;
      }
    }
    // modo de cálculo: si en NOM se eligió otro que el del archivo, calcPr lleva el nuevo
    const modoNom = doc.valor("_doc", "calculo");
    let nuevoModo = null;
    {
      const cp0 = hsL.find((h) => O.local(h.n) === "calcPr");
      const orig = cp0 ? (O.attr(libroTxt, cp0, "calcMode") || "auto") : "auto";
      const quiere = modoNom === "manual" ? "manual" : modoNom === "sinTablas" ? "autoNoTable" : modoNom === "auto" ? "auto" : null;
      if (quiere && quiere !== orig) { nuevoModo = quiere; cambio = true; }
    }
    if (cambio) {
      // Los valores guardados de las fórmulas pueden ya no servir: Excel recalcula al abrir.
      const cc = relsL.find((r) => /\/calcChain$/.test(r.tipo || ""));
      if (cc) { const ruta = pq.resolver(LIBRO, cc.destino); pq.quita(ruta); pq.quitaTipo(ruta); pq.quitaRel(LIBRO, cc.id); }
      const cp = hsL.find((h) => O.local(h.n) === "calcPr");
      const modoAttr = (a) => (nuevoModo ? quitaAttr(a, "calcMode") + (nuevoModo !== "auto" ? ` calcMode="${nuevoModo}"` : "") : a);
      if (cp) piezasL.push([cp.i, cp.f, `<${cp.n}${modoAttr(quitaAttr(libroTxt.slice(cp.a, cp.af), "fullCalcOnLoad"))} fullCalcOnLoad="1"/>`]);
      else {
        const ORD = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr"];
        const sig = hsL.find((h) => ORD.indexOf(O.local(h.n)) < 0 || ORD.indexOf(O.local(h.n)) > 9), pos = sig ? sig.i : raizL.cf;
        piezasL.push([pos, pos, `<${prefL}calcPr${modoAttr("")} fullCalcOnLoad="1"/>`]);
      }
    }
    if (cachesNuevas.length) {
      // las cachés de las tablas dinámicas nuevas, con un cacheId que el libro no use
      const pcs = hsL.find((h) => O.local(h.n) === "pivotCaches");
      let maxId = 0;
      if (pcs) (libroTxt.slice(pcs.i, pcs.f).match(/cacheId\s*=\s*"(\d+)"/g) || []).forEach((m) => (maxId = Math.max(maxId, +m.replace(/\D/g, ""))));
      const els = cachesNuevas.map((cn, k) => {
        const id = maxId + 1 + k, rid = "rIdNomPC" + (k + 1);
        pq.agregaRel(LIBRO, rid, T_PV_CACHE, "/" + cn.cache);
        pq.pon(cn.tabla, cn.partes.tabla.replace("\u0001CACHE\u0001", String(id)));
        return `<${prefL}pivotCache cacheId="${id}" xmlns:r="${NS_R}" r:id="${rid}"/>`;
      }).join("");
      if (pcs) piezasL.push([pcs.cf, pcs.cf, els]);
      else {
        const DESPUES = ["smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];
        const sig = hsL.find((h) => DESPUES.includes(O.local(h.n))), pos = sig ? sig.i : raizL.cf;
        piezasL.push([pos, pos, `<${prefL}pivotCaches>${els}</${prefL}pivotCaches>`]);
      }
    }
    pq.pon(LIBRO, aplicarPiezas(libroTxt, piezasL));
    if (fab.usaDA && fab.cmDA == null) {
      // el libro no tenía metadatos de celda: se agrega la parte con la matriz dinámica (cm="1")
      pq.pon("xl/metadata.xml", META_DA); pq.tipo("xl/metadata.xml", CT_META);
      pq.agregaRel(LIBRO, "rIdNomMeta", T_META, "/xl/metadata.xml");
    }
    if (fab.fuentes.length || fab.rellenos.length || fab.bordes.length || fab.xfs.length || fab.numFmts.length || fab.dxfs.length) pq.pon(estRuta, estilosConNuevos(estTxt, EO, fab));
    return pq.zip();
  }

  // ------------------------------------------------------------ importar Excel (.xlsx)
  const FMT_EXCEL = { 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 5: "$#,##0", 6: "$#,##0", 7: "$#,##0.00", 8: "$#,##0.00", 9: "0%", 10: "0.00%",
    14: "dd/mm/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "hh:mm", 21: "hh:mm:ss", 22: "dd/mm/yyyy hh:mm", 37: "#,##0", 38: "#,##0", 39: "#,##0.00", 40: "#,##0.00", 49: "@" };
  // Versión 2 del importador de Excel (16-nov-2026): los formatos internos que faltaban y el «'» de los textos que
  // parecen número o fórmula. Lo importado antes (versión 1) se compara al exportar con la lectura de entonces.
  // (los 20, 21 y 47 de Excel son «h:mm», «h:mm:ss» y «mm:ss.0»; la norma publica «mmss.0», que Excel no usa)
  const FMT_EXCEL_V2 = Object.assign({}, FMT_EXCEL, { 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 20: "h:mm", 21: "h:mm:ss", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0" });
  /** .xlsx → { hojas: [{ nombre, campos }], aviso }: cada hoja con sus celdas, formato, combinadas, anchos y filas fijas. */
  /** Los nombres definidos de todo el libro en workbook.xml: { minúsculas: { n: nombre, d: definición en NOM («=…») } }. */
  function nombresDelLibro(wb) {
    const out = {};
    const re = /<(?:\w+:)?definedName\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?definedName>/g;
    let m;
    while ((m = re.exec(wb))) {
      const nm = (/\sname\s*=\s*"([^"]*)"/.exec(m[1]) || [])[1];
      if (!nm || /^_xlnm\./i.test(nm) || /\slocalSheetId\s*=/.test(m[1]) || /\shidden\s*=\s*"(1|true)"/.test(m[1])) continue;
      const t = m[2].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&").trim();
      let d;
      try { d = F().aEspanol("=" + t.replace(/_xlfn\.|_xlpm\.|_xlws\./g, "")); } catch (e) { d = "=" + t; }
      out[nm.toLowerCase()] = { n: nm, d };
    }
    return out;
  }
  async function leerXlsx(archivo) {
    if (archivo.size > 30 * 1024 * 1024) throw new Error("El archivo pesa más de 30 MB");
    const u8 = new Uint8Array(await archivo.arrayBuffer());
    const z = NOM.origen.abrirZip(u8), r = Object.assign(modeloXlsx(z), { archivo: u8 });
    // lo del libro: estructura protegida (no agregar, quitar, mover ni renombrar hojas), contraseña para modificar con
    // «se recomienda abrir como de solo lectura», y marcado como final
    const doc = {}, wb = z.texto("xl/workbook.xml");
    if (wb) {
      const x = new DOMParser().parseFromString(wb, "application/xml");
      const wp = x.getElementsByTagNameNS("*", "workbookProtection")[0], fs = x.getElementsByTagNameNS("*", "fileSharing")[0];
      const si = (n, k) => /^(1|true)$/.test(n.getAttribute(k) || "");
      if (wp && (si(wp, "lockStructure") || si(wp, "lockWindows"))) doc.libroProt = { p: proteccionDe(wp, false, "workbook") || {}, estructura: si(wp, "lockStructure") || undefined, ventanas: si(wp, "lockWindows") || undefined };
      if (fs && (fs.getAttribute("hashValue") || fs.getAttribute("reservationPassword") || si(fs, "readOnlyRecommended"))) {
        const p = proteccionDe(fs, false) || {};
        if (!p.hash && fs.getAttribute("reservationPassword")) p.viejo = fs.getAttribute("reservationPassword");
        doc.modificar = { recomendada: si(fs, "readOnlyRecommended") || undefined, p };
      }
    }
    // modo de cálculo del libro (calcPr): manual o automático excepto tablas de datos
    if (wb) {
      const cp = /<(?:\w+:)?calcPr\b([^>]*)>/.exec(wb), cm = cp ? (/\scalcMode\s*=\s*"([^"]*)"/.exec(cp[1]) || [])[1] : null;
      if (cm === "manual") doc.calculo = "manual"; else if (cm === "autoNoTable") doc.calculo = "sinTablas";
    }
    // nombres definidos del libro (los de todo el libro; los de una hoja y los de Excel —_xlnm— se conservan en el original)
    if (wb) {
      const nombres = nombresDelLibro(wb);
      if (Object.keys(nombres).length) doc.nombres = nombres;
    }
    const c = z.texto("docProps/custom.xml");
    if (c && /name="_MarkAsFinal"[^>]*>\s*<vt:bool>\s*(true|1)\s*<\/vt:bool>/i.test(c)) doc.final = true;
    if (Object.keys(doc).length) r.doc = doc;
    return r;
  }
  /**
   * Lo que NOM entiende de un libro de Excel ya abierto (z = el .zip). Síncrono: también lo usa
   * el exportador para saber qué cambió.
   * OJO: el exportador compara el documento contra lo que ESTA función lee hoy del archivo original.
   * Si aquí se empieza a leer algo nuevo de las celdas (otro color, otra propiedad), en un documento
   * importado con la versión anterior eso se verá como "la persona lo quitó". Al cambiar lo que se
   * lee hay que cuidar esos documentos (Word y PowerPoint no tienen el problema: guardan huellas).
   */
  // Versiones del importador de Excel (lo que se guarda y lo que se exporta se compara con la misma versión):
  // 1 la primera · 2 formato de números y medidas de Excel · 3 tinte exacto de Excel, colores indexados, bordes
  // con su estilo y color, y los estilos de tabla (los que Excel trae de fábrica y los del libro) · 4 alineación
  // completa, tramas, degradados y diagonales · 5 matrices dinámicas: lo derramado no se guarda como valores
  // sueltos (fa: guarda el lugar que derramaba), {1,2;3,4} con sus separadores, A1#, «@» y LAMBDA sin «_xlpm.» ·
  // 6 protección: la de cada hoja (prot), la de cada celda (pr: «u» sin bloquear, «h» fórmula oculta) y la del libro ·
  // 7 nombres definidos: las fórmulas los conservan (=B2*IVA, como Excel; antes se sustituían por su definición) y el
  // libro los guarda en «_doc» (nombres), para el Administrador de nombres · 8 esquema: el nivel de cada fila (ol:) y de
  // cada columna (olc:) agrupada · 9 «mostrar fórmulas» de la hoja (vf) · 10 tablas de datos (t="dataTable": cada celda
  // de su lugar lleva =TABLA(fila;columna), como la muestra Excel).
  // ---- Tablas con formato hechas en NOM (t.lo, «Dar formato como tabla»): xl/tables/tableN.xml como las escribe Excel.
  // Cada una: { id, n (nombre), c1, f1, c2, f2, estilo, enc, tot, rayasF, rayasC, primeraC, ultimaC, filtro }.
  const CT_TABLA_XL = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
  // ---- Tablas dinámicas hechas en NOM (t.pv, js/dinamicas.js): sus tres partes de Excel y cómo se enlazan
  const T_PV_TABLA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
  const T_PV_CACHE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
  const T_PV_REG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
  const CT_PV_TABLA = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
  const CT_PV_CACHE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
  const CT_PV_REG = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
  /** Las tablas dinámicas de la hoja `t`, armadas con los datos de ahora (las que no se pueden armar se saltan). */
  const dinamicasDe = (doc, t) => (Array.isArray(t && t.pv) && NOM.dinamicas ? t.pv.map((pv) => NOM.dinamicas.armar(doc, t, pv)).filter((r) => r && typeof r === "object") : []);
  const T_TABLA_XL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
  const FUNC_TOTAL_XL = { 101: "average", 102: "countNums", 103: "count", 104: "max", 105: "min", 107: "stdDev", 109: "sum", 110: "var" };
  const tablasLoDe = (t) => (Array.isArray(t && t.lo) ? t.lo.filter((lo) => lo && lo.c2 >= lo.c1 && lo.f2 >= lo.f1) : []);
  /** Las columnas de una tabla para Excel: el nombre (el texto del encabezado, sin repetir ni vacío) y lo de la fila de
      totales (un SUBTOTAL de la columna o un rótulo). */
  function columnasTablaXl(lo, t, ev) {
    const vistos = new Set(), out = [];
    for (let c = lo.c1; c <= lo.c2; c++) {
      let n = lo.enc ? String(F().formatear(ev.valor(c, lo.f1)) || "").trim() : "";
      if (!n) n = "Columna" + (c - lo.c1 + 1);
      let m = n, k = 2;
      while (vistos.has(m.toLowerCase())) m = n + k++;
      vistos.add(m.toLowerCase());
      const col = { n: m };
      if (lo.tot) {
        const cr = String(t["c:" + F().ref(c, lo.f2)] || "");
        const sm = /^=\s*SUBTOTAL\(\s*(\d+)\s*,/i.exec(cr);
        if (sm && FUNC_TOTAL_XL[+sm[1]]) col.fn = FUNC_TOTAL_XL[+sm[1]];
        else if (cr && !cr.startsWith("=")) col.rot = cr;
      }
      out.push(col);
    }
    return out;
  }
  function xmlTablaXl(lo, id, nombre, cols) {
    const ref = F().ref(lo.c1, lo.f1) + ":" + F().ref(lo.c2, lo.f2), si = (v) => (v ? "1" : "0");
    const af = lo.enc && lo.filtro !== false ? `<autoFilter ref="${F().ref(lo.c1, lo.f1)}:${F().ref(lo.c2, lo.f2 - (lo.tot ? 1 : 0))}"/>` : "";
    return XML + `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${x(nombre)}" displayName="${x(nombre)}" ref="${ref}"${lo.enc ? "" : ' headerRowCount="0"'}${lo.tot ? ' totalsRowCount="1"' : ""}>${af}`
      + `<tableColumns count="${cols.length}">${cols.map((c, i) => `<tableColumn id="${i + 1}" name="${x(c.n)}"${c.rot ? ` totalsRowLabel="${x(c.rot)}"` : ""}${c.fn ? ` totalsRowFunction="${c.fn}"` : ""}/>`).join("")}</tableColumns>`
      + (lo.estilo === "" ? `<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>`
        : `<tableStyleInfo name="${x(lo.estilo || "TableStyleMedium2")}" showFirstColumn="${si(lo.primeraC)}" showLastColumn="${si(lo.ultimaC)}" showRowStripes="${si(lo.rayasF !== false)}" showColumnStripes="${si(lo.rayasC)}"/>`)
      + `</table>`;
  }
  /** Nombre de tabla válido y único en el libro (Excel: letra o «_» al principio, sin espacios). */
  function nombreTablaXl(n, usados) {
    let b = String(n || "Tabla").replace(/[^\p{L}\p{N}_.]/gu, "_");
    if (!/^[\p{L}_]/u.test(b)) b = "_" + b;
    let m = b, k = 2;
    while (usados.has(m.toLowerCase())) m = b + "_" + k++;
    usados.add(m.toLowerCase());
    return m;
  }
  // ---- Estilos de tabla de Excel: lo que pinta un estilo (NOM.ESTILOS_TABLA o uno del libro) en cada celda de una tabla,
  // con la precedencia de Excel. Lo usan el importador (las tablas de un libro de Excel) y la hoja (las tablas hechas en
  // NOM con «Dar formato como tabla», que se pintan en vivo). `colorTabla` traduce «t4:0.8», «#RRGGBB» o «i22».
  /** El tema de Office de Excel 2024 (el que usa Excel en un libro sin tema propio, como los de NOM), en el orden de los
      índices de Excel: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4-9 énfasis, 10 vínculo, 11 vínculo visitado. */
  const TEMA_OFFICE_XL = ["FFFFFF", "000000", "E8E8E8", "0E2841", "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E", "467886", "96607D"];
  /** El tinte de Excel sobre un color del tema (HLS de Windows con enteros, como el importador en la versión 3). */
  function conTinteXl(hex, tinte) {
    if (!tinte) return hex;
    const [R, G, B] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)), [H, L, S] = rgbAHls(R, G, B);
    if (!S) { const L3 = Math.max(0, Math.min(HLSMAX, Math.round(tinte < 0 ? L * (1 + tinte) : L * (1 - tinte) + HLSMAX * tinte))), g = Math.floor((L3 * 255 + 120) / HLSMAX); return [g, g, g].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase(); }
    const L2 = tinte < 0 ? Math.trunc(L * (1 + tinte)) : Math.trunc((((Math.max(R, G, B) + Math.min(R, G, B)) * HLSMAX) / 510) * (1 - tinte) + (HLSMAX - HLSMAX * (1 - tinte)));
    return hlsARgb(H, Math.max(0, Math.min(HLSMAX, L2)), S).map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  /** Traductor de colores de estilo de tabla con un tema (12 colores sin «#»; si falta, el de Office 2024). */
  const colorTablaConTema = (tema) => (x) => {
    if (!x) return null;
    if (x[0] === "#") return x.toUpperCase();
    if (x[0] === "i") return INDEXADOS_XL[+x.slice(1)] ? "#" + INDEXADOS_XL[+x.slice(1)] : null;
    const m = /^t(\d+)(?::(-?[\d.eE-]+))?$/.exec(x), tm = tema || TEMA_OFFICE_XL;
    return m && tm[+m[1]] ? "#" + conTinteXl(tm[+m[1]], +(m[2] || 0)) : null;
  };
  /** Pinta en `campos` (r:, k:, n:, i:, bx:) las celdas de una tabla { c1, f1, c2, f2, enc, tot, rayasF, rayasC,
      primeraC, ultimaC } con el estilo `est2`. Lo que `campos` ya tenga gana. */
  function pintarTablaXl(tb, est2, colorTabla, campos) {
    const { c1, c2, f1, f2 } = tb, fEnc = tb.enc ? f1 : -1, fTot = tb.tot ? f2 : -1, d1 = f1 + (tb.enc ? 1 : 0), d2 = f2 - (tb.tot ? 1 : 0);
    // franjas: cada una mide «s» filas (o columnas); la primera y la segunda se alternan
    const franja = (i, ini, n1, n2) => { const k = (i - ini) % (n1 + n2), a = i - k; return k < n1 ? [0, a, a + n1 - 1] : [1, a + n1, a + n1 + n2 - 1]; };
    const s1F = (est2.firstRowStripe || {}).s || 1, s2F = (est2.secondRowStripe || {}).s || 1, s1C = (est2.firstColumnStripe || {}).s || 1, s2C = (est2.secondColumnStripe || {}).s || 1;
    for (let f = f1; f <= f2; f++) for (let c = c1; c <= c2; c++) {
      const capas = [], pon = (tipo, rc1, rf1, rc2, rf2) => { if (est2[tipo]) capas.push({ e: est2[tipo], rc1, rf1, rc2, rf2 }); };
      pon("wholeTable", c1, f1, c2, f2);
      const enDatos = f >= d1 && f <= d2;
      if (enDatos && tb.rayasC) { const [n, a, b] = franja(c, c1, s1C, s2C); pon(n ? "secondColumnStripe" : "firstColumnStripe", a, d1, Math.min(b, c2), d2); }
      if (enDatos && tb.rayasF) { const [n, a, b] = franja(f, d1, s1F, s2F); pon(n ? "secondRowStripe" : "firstRowStripe", c1, a, c2, Math.min(b, d2)); }
      if (tb.ultimaC && c === c2) pon("lastColumn", c2, f1, c2, f2);
      if (tb.primeraC && c === c1) pon("firstColumn", c1, f1, c1, f2);
      if (f === fEnc) pon("headerRow", c1, f, c2, f);
      if (f === fTot) pon("totalRow", c1, f, c2, f);
      if (f === fEnc && c === c1) pon("firstHeaderCell", c, f, c, f);
      if (f === fEnc && c === c2) pon("lastHeaderCell", c, f, c, f);
      if (f === fTot && c === c1) pon("firstTotalCell", c, f, c, f);
      if (f === fTot && c === c2) pon("lastTotalCell", c, f, c, f);
      pintarCapasXl(capas, c, f, campos, colorTabla);
      if (!campos["r:" + F().ref(c, f)]) campos["r:" + F().ref(c, f)] = "#FFFFFF";   // sin cuadrícula adentro, como Excel
    }
  }
  /** Pinta una celda con las capas de un estilo de tabla (de menor a mayor precedencia): relleno, letra y bordes
      (los de afuera de cada región o las líneas de adentro). Lo propio de la celda gana. */
  function pintarCapasXl(capas, c, f, campos, colorTabla) {
    let relleno = null, tinta = null, negrita = false, cursiva = false;
    const lados = {};
    for (const x of capas) {
      const e = x.e;
      if (e.f) relleno = e.f;
      if (e.k) tinta = e.k;
      if (e.b) negrita = true;
      if (e.i) cursiva = true;
      if (e.bd) {
        const bd = e.bd, ar = f === x.rf1 ? bd.t : bd.h, ab = f === x.rf2 ? bd.b : bd.h, iz = c === x.rc1 ? bd.l : bd.v, de = c === x.rc2 ? bd.r : bd.v;
        if (ar) lados.a = ar;
        if (ab) lados.b = ab;
        if (iz) lados.i = iz;
        if (de) lados.d = de;
      }
    }
    const R = F().ref(c, f);
    if (relleno && !campos["r:" + R]) { const x = colorTabla(relleno); if (x) campos["r:" + R] = x; }
    if (tinta && !campos["k:" + R]) { const x = colorTabla(tinta); if (x && x !== "#000000") campos["k:" + R] = x; }
    if (negrita && !campos["n:" + R]) campos["n:" + R] = true;
    if (cursiva && !campos["i:" + R]) campos["i:" + R] = true;
    if (Object.keys(lados).length) {
      const propios = {};
      String(campos["bx:" + R] || "").split(";").filter(Boolean).forEach((q) => (propios[q.split(":")[0]] = q));
      for (const l in lados) if (!propios[l]) propios[l] = `${l}:${lados[l][0]}:${colorTabla(lados[l][1]) || "#000000"}`;
      campos["bx:" + R] = Object.values(propios).join(";");
    }
  }
  const ORDEN_DINAMICA = ["wholeTable", "pageFieldLabels", "pageFieldValues", "firstColumnStripe", "secondColumnStripe", "firstRowStripe", "secondRowStripe", "blankRow",
    "firstColumn", "firstSubtotalColumn", "secondSubtotalColumn", "thirdSubtotalColumn", "firstColumnSubheading", "secondColumnSubheading", "thirdColumnSubheading",
    "firstRowSubheading", "secondRowSubheading", "thirdRowSubheading", "firstSubtotalRow", "secondSubtotalRow", "thirdSubtotalRow", "lastColumn", "headerRow", "firstHeaderCell", "totalRow"];
  const NIVEL = ["first", "second", "third"];
  /** Pinta en `campos` lo que el estilo `est2` pone en cada área de una tabla dinámica de Excel (pivotTableDefinition ya
      leída: `pd`), con la precedencia de Excel. La usan el importador y las tablas dinámicas hechas en NOM (js/dinamicas.js). */
  function pintarDinamicaXl(pd, est2, colorTabla, campos) {
    const hijos = (n, nombre) => [...(n ? n.childNodes : [])].filter((c) => c.nodeType === 1 && (!nombre || c.localName === nombre));
    const hijo = (n, nombre) => hijos(n, nombre)[0] || null;
    const at = (n, a) => (n ? n.getAttribute(a) : null);
    const loc = hijo(pd, "location"), si = hijo(pd, "pivotTableStyleInfo");
    const [a1, a2] = String(at(loc, "ref") || "").split(":").map((q) => F().parseRef(q));
    if (!a1) return;
    const c1 = a1.c, f1 = a1.f, c2 = (a2 || a1).c, f2 = (a2 || a1).f;
    const fData = f1 + +(at(loc, "firstDataRow") || 0), cData = c1 + +(at(loc, "firstDataCol") || 0), fEnc = f1 + +(at(loc, "firstHeaderRow") || 0);
    const si1 = (a) => at(si, a) === "1";
    const filasI = hijos(hijo(pd, "rowItems"), "i"), colsI = hijos(hijo(pd, "colItems"), "i");
    const nCamposF = hijos(hijo(pd, "rowFields"), "field").length, nCamposC = hijos(hijo(pd, "colFields"), "field").filter((x) => at(x, "x") !== "-2").length;
    const prof = (i) => +(at(i, "r") || 0) + Math.max(1, hijos(i, "x").length) - 1;
    const capas = new Map();
    const pon = (tipo, rc1, rf1, rc2, rf2) => {
      if (!est2[tipo]) return;
      for (let f = rf1; f <= rf2; f++) for (let c = rc1; c <= rc2; c++) {
        const k = c + "," + f;
        if (!capas.has(k)) capas.set(k, []);
        capas.get(k).push({ tipo, e: est2[tipo], rc1, rf1, rc2, rf2 });
      }
    };
    // botones de filtro de Excel: etiquetas de fila y de columna, y el valor de cada filtro
    if (nCamposF) campos["fb:" + F().ref(c1, fEnc)] = true;
    if (nCamposC && fEnc > f1) campos["fb:" + F().ref(cData, fEnc - 1)] = true;
    pon("wholeTable", c1, f1, c2, f2);
    // filtros (campos de página): encima de la tabla, con un renglón vacío en medio
    const nPag = +(at(loc, "rowPageCount") || 0), nPagC = Math.max(1, +(at(loc, "colPageCount") || 0));
    const nCamposP = hijos(hijo(pd, "pageFields"), "pageField").length;
    for (let k = 0; k < nCamposP; k++) {
      const fp = f1 - 1 - nPag + (k % Math.max(1, nPag)), cp = c1 + 3 * Math.floor(k / Math.max(1, nPag));
      if (fp < 0) continue;
      pon("pageFieldLabels", cp, fp, cp, fp);
      pon("pageFieldValues", cp + 1, fp, cp + 1, fp);
      campos["fb:" + F().ref(cp + 1, fp)] = true;
    }
    if (si1("showColStripes")) for (let c = cData; c <= c2; c++) { const s1 = (est2.firstColumnStripe || {}).s || 1, s2 = (est2.secondColumnStripe || {}).s || 1, k = (c - cData) % (s1 + s2); pon(k < s1 ? "firstColumnStripe" : "secondColumnStripe", c, fData, c, f2); }
    if (si1("showRowStripes")) for (let f = fData; f <= f2; f++) { const s1 = (est2.firstRowStripe || {}).s || 1, s2 = (est2.secondRowStripe || {}).s || 1, k = (f - fData) % (s1 + s2); pon(k < s1 ? "firstRowStripe" : "secondRowStripe", c1, f, c2, f); }
    if (si1("showRowHeaders") && cData > c1) pon("firstColumn", c1, fData, cData - 1, f2);
    // cada fila de datos según su elemento: total general, subtotal, renglón vacío o subtítulo de un nivel de afuera
    filasI.forEach((i, k) => {
      const f = fData + k, t = at(i, "t");
      if (f > f2) return;
      if (t === "grand") pon("totalRow", c1, f, c2, f);
      else if (t === "blank") pon("blankRow", c1, f, c2, f);
      else if (t && t !== "data") pon(NIVEL[Math.min(2, prof(i))] + "SubtotalRow", c1, f, c2, f);
      else if (nCamposF > 1 && prof(i) < nCamposF - 1) pon(NIVEL[Math.min(2, prof(i))] + "RowSubheading", c1, f, c2, f);
    });
    colsI.forEach((i, k) => {
      const c = cData + k, t = at(i, "t");
      if (c > c2) return;
      if (t === "grand") { if (si1("showLastColumn")) pon("lastColumn", c, f1, c, f2); }
      else if (t && t !== "data") pon(NIVEL[Math.min(2, prof(i))] + "SubtotalColumn", c, f1, c, f2);
      else if (nCamposC > 1 && prof(i) < nCamposC - 1) pon(NIVEL[Math.min(2, prof(i))] + "ColumnSubheading", c, f1, c, fData - 1);
    });
    if (si1("showColHeaders") && fData > f1) pon("headerRow", c1, f1, c2, fData - 1);
    pon("firstHeaderCell", c1, fEnc, c1, fEnc);
    for (const [k, lista] of capas) {
      const [c, f] = k.split(",").map(Number);
      lista.sort((a, b) => ORDEN_DINAMICA.indexOf(a.tipo) - ORDEN_DINAMICA.indexOf(b.tipo));
      pintarCapasXl(lista, c, f, campos, colorTabla);
    }
    // Excel no dibuja la cuadrícula dentro de una tabla dinámica (ni en sus filtros): fondo blanco donde no hay otro
    for (let f = Math.max(0, f1 - 1 - nPag); f <= f2; f++) for (let c = c1; c <= c2; c++) {
      if (f < f1 && !(capas.has(c + "," + f))) continue;
      const R = F().ref(c, f);
      if (!campos["r:" + R]) campos["r:" + R] = "#FFFFFF";
    }
  }
  const VERSION_XL = 13;   // (13: el tema del libro en t.xl.tema, para las tablas con formato hechas en NOM)
  function modeloXlsx(z, version) {
    const V = version || VERSION_XL;
    const P = (t) => new DOMParser().parseFromString(t, "application/xml");
    const libro = z.texto("xl/workbook.xml");
    if (!libro) throw new Error("El archivo no parece un libro de Excel (.xlsx)");
    const hijos = (n, nombre) => [...(n ? n.childNodes : [])].filter((c) => c.nodeType === 1 && (!nombre || c.localName === nombre));
    const hijo = (n, nombre) => hijos(n, nombre)[0] || null;
    const todos = (n, nombre) => (n ? [...n.getElementsByTagNameNS("*", nombre)] : []);
    const at = (n, a) => (n ? n.getAttribute(a) : null);
    const rels = P(z.texto("xl/_rels/workbook.xml.rels") || "<Relationships/>");
    const destino = {};
    todos(rels, "Relationship").forEach((r) => (destino[at(r, "Id")] = at(r, "Target")));
    // Textos compartidos.
    const sst = z.texto("xl/sharedStrings.xml");
    const textos = sst ? todos(P(sst), "si").map((si) => todos(si, "t").map((t) => t.textContent).join("")) : [];
    // Estilos.
    const est = P(z.texto("xl/styles.xml") || "<styleSheet/>");
    const fmts = Object.assign({}, V >= 2 ? FMT_EXCEL_V2 : FMT_EXCEL);
    todos(est, "numFmt").forEach((n) => (fmts[+at(n, "numFmtId")] = at(n, "formatCode")));
    // Colores del tema (en Excel van por número: 0 fondo 1, 1 texto 1, 2 fondo 2, 3 texto 2, 4 a 9 los
    // acentos) con su "tinte": positivo aclara, negativo oscurece.
    const temaXl = [];
    {
      const tt = z.texto("xl/theme/theme1.xml"), cs = tt ? todos(P(tt), "clrScheme")[0] : null, porNombre = {};
      hijos(cs).forEach((c) => { const v = hijo(c, "srgbClr") ? at(hijo(c, "srgbClr"), "val") : hijo(c, "sysClr") ? at(hijo(c, "sysClr"), "lastClr") : null; if (v && /^[0-9a-f]{6}$/i.test(v)) porNombre[c.localName] = v.toUpperCase(); });
      ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"].forEach((k, i) => (temaXl[i] = porNombre[k] || null));
    }
    // El tinte como lo calcula Excel (comprobado contra Excel real, igual en 37 de 37 colores): el color pasa a
    // HLS de Windows con números enteros (HLSMAX = 240) y se cambia solo la luminosidad. Para aclarar, Excel parte
    // de la luminosidad sin redondear; para oscurecer, de la redondeada; y trunca el resultado.
    const conTinte = (hex, tinte) => {
      if (!tinte) return hex;
      if (V < 3) return [0, 2, 4].map((i) => { const v = parseInt(hex.slice(i, i + 2), 16); return Math.max(0, Math.min(255, Math.round(tinte > 0 ? v + (255 - v) * tinte : v * (1 + tinte)))).toString(16).padStart(2, "0"); }).join("").toUpperCase();
      const [R, G, B] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)), [H, L, S] = rgbAHls(R, G, B);
      // (los grises —blanco, negro— redondean: «Negro, más claro 50 %» es #808080 y «Blanco, más oscuro 15 %» #D9D9D9)
      if (!S) { const L3 = Math.max(0, Math.min(HLSMAX, Math.round(tinte < 0 ? L * (1 + tinte) : L * (1 - tinte) + HLSMAX * tinte))), g = Math.floor((L3 * 255 + 120) / HLSMAX); return [g, g, g].map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0")).join("").toUpperCase(); }
      const L2 = tinte < 0 ? Math.trunc(L * (1 + tinte)) : Math.trunc((((Math.max(R, G, B) + Math.min(R, G, B)) * HLSMAX) / 510) * (1 - tinte) + (HLSMAX - HLSMAX * (1 - tinte)));
      return hlsARgb(H, Math.max(0, Math.min(HLSMAX, L2)), S).map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0")).join("").toUpperCase();
    };
    // Los 64 colores «indexados» de Excel 97 (archivos viejos o hechos por otros programas), o los que el archivo cambió.
    const indexados = INDEXADOS_XL.slice();
    todos(todos(est, "indexedColors")[0], "rgbColor").forEach((c, i) => { const v = at(c, "rgb"); if (v && /^[0-9A-F]{8}$/i.test(v)) indexados[i] = v.slice(2).toUpperCase(); });
    const color = (n) => {
      const rgb = at(n, "rgb");
      if (rgb && /^[0-9A-F]{8}$/i.test(rgb)) return "#" + (V >= 3 ? conTinte(rgb.slice(2).toUpperCase(), +(at(n, "tint") || 0)) : rgb.slice(2).toUpperCase());
      const th = at(n, "theme");
      if (th != null && temaXl[+th]) return "#" + conTinte(temaXl[+th], +(at(n, "tint") || 0));
      const ix = at(n, "indexed");
      return V >= 3 && ix != null && indexados[+ix] ? "#" + conTinte(indexados[+ix], +(at(n, "tint") || 0)) : null;
    };
    const fuentes = todos(hijo(est.documentElement, "fonts"), "font").map((f) => ({ nombre: at(hijo(f, "name"), "val") || null, b: !!hijo(f, "b") && at(hijo(f, "b"), "val") !== "0", i: !!hijo(f, "i") && at(hijo(f, "i"), "val") !== "0", u: !!hijo(f, "u") && at(hijo(f, "u"), "val") !== "none", k: hijo(f, "color") ? color(hijo(f, "color")) : null, sz: hijo(f, "sz") ? +at(hijo(f, "sz"), "val") : 11 }));
    const rellenos = todos(hijo(est.documentElement, "fills"), "fill").map((f) => { const p = hijo(f, "patternFill"); return p && at(p, "patternType") === "solid" && hijo(p, "fgColor") ? color(hijo(p, "fgColor")) : null; });
    // (versión 4) Tramas y degradados de celda, para pintarlos como Excel: { t: trama, fg, bg } (fg: el color de la
    // trama; bg: el de la celda) o { g: "lin", deg, p: paradas } / { g: "path", l, r, t, b, p } (desde el rectángulo
    // interior hacia las orillas). En el relleno simple (r:) queda el color de la celda o la mezcla del degradado.
    const rellenosX = V >= 4 ? todos(hijo(est.documentElement, "fills"), "fill").map((f) => {
      const p = hijo(f, "patternFill"), g = hijo(f, "gradientFill");
      if (p) {
        const tp = at(p, "patternType");
        if (!tp || tp === "none" || tp === "solid") return null;
        return { t: tp, fg: (hijo(p, "fgColor") && color(hijo(p, "fgColor"))) || "#000000", bg: (hijo(p, "bgColor") && color(hijo(p, "bgColor"))) || "#FFFFFF" };
      }
      if (!g) return null;
      const fr = (x) => Math.max(0, Math.min(1, +x || 0));
      const pr = todos(g, "stop").map((q) => [fr(at(q, "position")), (hijo(q, "color") && color(hijo(q, "color"))) || "#FFFFFF"]).sort((a, b) => a[0] - b[0]);
      if (!pr.length) return null;
      return at(g, "type") === "path" ? { g: "path", l: fr(at(g, "left")), r: fr(at(g, "right")), t: fr(at(g, "top")), b: fr(at(g, "bottom")), p: pr } : { g: "lin", deg: +at(g, "degree") || 0, p: pr };
    }) : [];
    const mezcla = (a, b) => "#" + [1, 3, 5].map((i) => Math.round((parseInt(a.slice(i, i + 2), 16) + parseInt(b.slice(i, i + 2), 16)) / 2).toString(16).padStart(2, "0")).join("").toUpperCase();
    const rellenoDe = (id) => rellenos[id] || (rellenosX[id] ? (rellenosX[id].t ? rellenosX[id].bg : mezcla(rellenosX[id].p[0][1], rellenosX[id].p[rellenosX[id].p.length - 1][1])) : null);
    const bordesX = todos(hijo(est.documentElement, "borders"), "border").map((b) => [["top", "a"], ["bottom", "b"], ["left", "i"], ["right", "d"]].map(([l, k]) => {
      const e = hijo(b, l), st = e && at(e, "style");
      return st && st !== "none" ? `${k}:${st}:${color(hijo(e, "color")) || "#000000"}` : "";
    }).concat(V >= 4 ? (() => {
      // (versión 4) las diagonales: u, de abajo a arriba; w, de arriba a abajo
      const dg = hijo(b, "diagonal"), st = dg && at(dg, "style"), k = st && st !== "none" ? color(hijo(dg, "color")) || "#000000" : null;
      return k ? [at(b, "diagonalUp") === "1" ? `u:${st}:${k}` : "", at(b, "diagonalDown") === "1" ? `w:${st}:${k}` : ""] : [];
    })() : []).filter(Boolean).join(";"));
    const bordes = todos(hijo(est.documentElement, "borders"), "border").map((b) => ["top", "bottom", "left", "right"].map((l, k) => (hijo(b, l) && at(hijo(b, l), "style") ? "abid"[k] : "")).join(""));
    const xfs = todos(hijo(est.documentElement, "cellXfs"), "xf").map((xf) => {
      const al = hijo(xf, "alignment"), pt = V >= 6 ? hijo(xf, "protection") : null;
      return { pr: pt ? (/^(0|false)$/.test(at(pt, "locked") || "") ? "u" : "") + (/^(1|true)$/.test(at(pt, "hidden") || "") ? "h" : "") : "", num: +at(xf, "numFmtId") || 0, fuente: fuentes[+at(xf, "fontId") || 0] || {}, relleno: rellenoDe(+at(xf, "fillId") || 0) || null, rx: rellenosX[+at(xf, "fillId") || 0] || null, borde: bordes[+at(xf, "borderId") || 0] || "", bx: bordesX[+at(xf, "borderId") || 0] || "", al: al ? { left: "izq", center: "centro", centerContinuous: "centro", right: "der" }[at(al, "horizontal")] || null : null, aj: al ? at(al, "wrapText") === "1" : false,
        // (versión 4: lo demás de la alineación de Excel —rellenar, justificar, distribuido, centrar en la selección,
        // vertical, giro del texto, sangría y reducir hasta ajustar— para pintarlo como Excel)
        ax: V >= 4 && al ? (() => {
          const o = {}, hz = at(al, "horizontal"), vt = at(al, "vertical"), rot = +(at(al, "textRotation") || 0), ind = +(at(al, "indent") || 0);
          if (hz && !/^(general|left|right|center)$/.test(hz)) o.h = hz;
          if (vt && vt !== "bottom") o.v = vt;
          if (rot > 0 && rot <= 180 || rot === 255) o.r = rot;
          if (ind > 0) o.i = Math.min(ind, 250);
          if (at(al, "shrinkToFit") === "1") o.s = 1;
          return Object.keys(o).length ? o : null;
        })() : null };
    });
    const hojasXml = todos(P(libro), "sheet");
    // Tablas de Excel (ListObject) de todas las hojas: su nombre, su rango y sus columnas, para traducir las
    // referencias estructuradas (Tabla[Columna], [@Columna]…) a referencias normales que el motor entiende.
    const tablasXl = {};
    if (V >= 2) for (const h of hojasXml) {
      let rh = destino[h.getAttribute("r:id") || h.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")] || "";
      rh = rh.startsWith("/") ? rh.slice(1) : "xl/" + rh.replace(/^\.\//, "");
      const relsH = z.texto(rh.replace(/([^/]+)$/, "_rels/$1.rels"));
      if (!relsH) continue;
      todos(P(relsH), "Relationship").filter((r) => /\/table$/.test(at(r, "Type") || "")).forEach((r) => {
        const tg = at(r, "Target") || "", rt = tg.startsWith("/") ? tg.slice(1) : rh.replace(/[^/]+$/, "") + tg;
        const norm = rt.split("/").reduce((a, x) => { if (x === "..") a.pop(); else if (x && x !== ".") a.push(x); return a; }, []).join("/");
        const tx = z.texto(norm); if (!tx) return;
        const td = P(tx).documentElement, ref = String(at(td, "ref") || "").split(":");
        const a1 = F().parseRef(ref[0] || ""), a2 = F().parseRef(ref[1] || ref[0] || "");
        if (!a1 || !a2) return;
        const cols = todos(td, "tableColumn").map((c) => at(c, "name") || "");
        const tsi = todos(td, "tableStyleInfo")[0], si = (a) => at(tsi, a) === "1" || at(tsi, a) === "true", conFiltro = hijo(td, "autoFilter");
        const info = { hoja: at(h, "name") || "", c1: a1.c, f1: a1.f, c2: a2.c, f2: a2.f, enc: at(td, "headerRowCount") === "0" ? 0 : 1, tot: +(at(td, "totalsRowCount") || 0), cols,
          estilo: tsi ? at(tsi, "name") : null, rayasF: si("showRowStripes"), rayasC: si("showColumnStripes"), primeraC: si("showFirstColumn"), ultimaC: si("showLastColumn"), filtro: conFiltro };
        [at(td, "displayName"), at(td, "name")].filter(Boolean).forEach((n) => (tablasXl[n.toLowerCase()] = info));
      });
    }
    // ---- Estilos de tabla (versión 3): el formato que Excel pinta en una tabla (encabezado, franjas, primera y
    // última columna, totales) con la precedencia de Excel. Lo propio de cada celda (su relleno, su color, sus
    // bordes) gana; el estilo solo llena lo que la celda no dice.
    const colorTabla = (x) => {
      if (!x) return null;
      if (x[0] === "#") return x.toUpperCase();
      if (x[0] === "i") return indexados[+x.slice(1)] ? "#" + indexados[+x.slice(1)] : null;
      const m = /^t(\d+)(?::(-?[\d.eE-]+))?$/.exec(x);
      return m && temaXl[+m[1]] ? "#" + conTinte(temaXl[+m[1]], +(m[2] || 0)) : null;
    };
    // estilos propios del libro (los que alguien hizo en Excel): sus elementos usan los formatos «dxf» del libro
    const estilosTablaLibro = {};
    const dxfsL = todos(hijo(est.documentElement, "dxfs"), "dxf");
    const formatoDxf = (d) => {
        const o = {}, fu = d && hijo(d, "font"), pf = d && todos(d, "patternFill")[0], bo = d && hijo(d, "border");
        if (fu) {
          const b = hijo(fu, "b"), i = hijo(fu, "i"), u = hijo(fu, "u"), st = hijo(fu, "strike");
          if (b) o.b = at(b, "val") === "0" ? 0 : 1;
          if (i) o.i = at(i, "val") === "0" ? 0 : 1;
          if (u && at(u, "val") !== "none") o.u = 1;
          if (st && at(st, "val") !== "0") o.st = 1;
          const k = color(hijo(fu, "color")); if (k) o.k = k;
        }
        if (pf && at(pf, "patternType") !== "none") { const f = color(hijo(pf, "bgColor")) || color(hijo(pf, "fgColor")); if (f) o.f = f; }
        if (bo) { const bd = {}; [["left", "l"], ["right", "r"], ["top", "t"], ["bottom", "b"], ["horizontal", "h"], ["vertical", "v"]].forEach(([l, k]) => { const e = hijo(bo, l), st = e && at(e, "style"); if (st && st !== "none") bd[k] = [st, color(hijo(e, "color")) || "#000000"]; }); if (Object.keys(bd).length) o.bd = bd; }
        return o;
    };
    if (V >= 3) {
      const ts = hijo(est.documentElement, "tableStyles");
      hijos(ts, "tableStyle").forEach((t2) => {
        const e = {};
        hijos(t2, "tableStyleElement").forEach((x) => { const f = formatoDxf(dxfsL[+at(x, "dxfId")]); if (+at(x, "size") > 1) f.s = +at(x, "size"); e[at(x, "type")] = f; });
        if (at(t2, "name")) estilosTablaLibro[at(t2, "name")] = e;
      });
    }
    // ---- Dibujos de la hoja (versión 3): imágenes, formas, líneas y gráficas con su ancla de Excel (celda y desfase).
    // Las gráficas se leen con su estilo para dibujarlas como Excel (js/graficas-office.js) con los valores vivos.
    const temaDml = {}, fuentesDml = {};
    {
      const tt = z.texto("xl/theme/theme1.xml"), dt = tt ? P(tt) : null, cs = dt ? todos(dt, "clrScheme")[0] : null;
      hijos(cs).forEach((c) => { const v = hijo(c, "srgbClr") ? at(hijo(c, "srgbClr"), "val") : hijo(c, "sysClr") ? at(hijo(c, "sysClr"), "lastClr") : null; if (v && /^[0-9a-f]{6}$/i.test(v)) temaDml[c.localName] = v.toUpperCase(); });
      Object.assign(temaDml, { tx1: temaDml.dk1, bg1: temaDml.lt1, tx2: temaDml.dk2, bg2: temaDml.lt2 });
      ["majorFont", "minorFont"].forEach((k, i) => { const f = dt && todos(dt, k)[0], l = f && hijo(f, "latin"); if (l && at(l, "typeface")) fuentesDml[i ? "menor" : "mayor"] = at(l, "typeface"); });
    }
    const PRST_DML = { black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00", gray: "808080", darkGray: "A9A9A9", lightGray: "D3D3D3", orange: "FFA500" };
    /** Color de DrawingML (gráficas y formas de la hoja) con sus cambios: tint y shade en RGB lineal, lumMod,
        lumOff, satMod y hueMod en HSL, como lo hace Office. */
    function colorDml(fill) {
      const c = fill && (fill.localName && /Clr$/.test(fill.localName) ? fill : todos(fill, "srgbClr")[0] || todos(fill, "schemeClr")[0] || todos(fill, "sysClr")[0] || todos(fill, "prstClr")[0]);
      if (!c) return null;
      let hx = c.localName === "srgbClr" ? at(c, "val") : c.localName === "schemeClr" ? temaDml[at(c, "val")] : c.localName === "sysClr" ? at(c, "lastClr") : PRST_DML[at(c, "val")];
      if (!/^[0-9a-f]{6}$/i.test(hx || "")) return null;
      hx = hx.toUpperCase();
      for (const t2 of hijos(c)) {
        if (t2.localName !== "shade" && t2.localName !== "tint") continue;
        const k = +at(t2, "val") / 100000, lin = [0, 2, 4].map((i) => Math.pow(parseInt(hx.slice(i, i + 2), 16) / 255, 2.3));
        hx = lin.map((x) => (t2.localName === "shade" ? x * k : 1 - (1 - x) * k)).map((x) => Math.round(Math.pow(Math.max(0, Math.min(1, x)), 1 / 2.3) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
      }
      const v = (k) => { const e = hijo(c, k); return e ? +at(e, "val") / 100000 : null; };
      if (v("lumMod") == null && v("lumOff") == null && v("satMod") == null && v("hueMod") == null) return "#" + hx;
      let [r, g, b] = [0, 2, 4].map((i) => parseInt(hx.slice(i, i + 2), 16) / 255);
      const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
      let hh = 0, sa = 0, l = (mx + mn) / 2;
      if (mx !== mn) { const d = mx - mn; sa = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); hh = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; hh /= 6; }
      if (v("hueMod") != null) hh = (hh * v("hueMod")) % 1;
      if (v("satMod") != null) sa = Math.max(0, Math.min(1, sa * v("satMod")));
      if (v("lumMod") != null) l *= v("lumMod");
      if (v("lumOff") != null) l += v("lumOff");
      l = Math.max(0, Math.min(1, l));
      const q = l < 0.5 ? l * (1 + sa) : l + sa - l * sa, p2 = 2 * l - q;
      const f = (t3) => { t3 = (t3 + 1) % 1; return t3 < 1 / 6 ? p2 + (q - p2) * 6 * t3 : t3 < 1 / 2 ? q : t3 < 2 / 3 ? p2 + (q - p2) * (2 / 3 - t3) * 6 : p2; };
      return "#" + [f(hh + 1 / 3), f(hh), f(hh - 1 / 3)].map((x) => Math.round(x * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
    }
    const fuenteDml = (n) => (n === "+mn-lt" || n === "+mn-ea" ? fuentesDml.menor || "Calibri" : n === "+mj-lt" || n === "+mj-ea" ? fuentesDml.mayor || "Calibri" : n);
    const rutaRel = (desde, tg) => { const rt = String(tg || "").startsWith("/") ? tg.slice(1) : desde.replace(/[^/]+$/, "") + tg; return rt.split("/").reduce((a, x) => { if (x === "..") a.pop(); else if (x && x !== ".") a.push(x); return a; }, []).join("/"); };
    const relsDeParte = (ruta) => { const t2 = z.texto(ruta.replace(/([^/]+)$/, "_rels/$1.rels")), o = {}; if (t2) todos(P(t2), "Relationship").forEach((r) => (o[at(r, "Id")] = { tipo: at(r, "Type") || "", ruta: at(r, "TargetMode") === "External" ? null : rutaRel(ruta, at(r, "Target")) })); return o; };
    const MIME = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", webp: "image/webp", emf: "image/x-emf", wmf: "image/x-wmf" };
    const dataUrlXl = (ruta) => {
      const b = ruta && z.bytes(ruta), ext = String(ruta || "").split(".").pop().toLowerCase();
      if (!b || !MIME[ext] || b.length > 8 * 1024 * 1024) return null;
      let bin = ""; for (let i = 0; i < b.length; i += 0x8000) bin += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
      return `data:${MIME[ext]};base64,${btoa(bin)}`;
    };
    const rellenoDml = (spPr) => { if (!spPr) return null; if (hijo(spPr, "noFill")) return "none"; const sf = hijo(spPr, "solidFill") || hijo(spPr, "gradFill"); return sf ? colorDml(todos(sf, "gs")[0] || sf) : null; };
    const lineaDml = (spPr, st) => {
      const ln = spPr && hijo(spPr, "ln"), lnRef = st && hijo(st, "lnRef");
      if (!ln) return lnRef && +at(lnRef, "idx") ? { w: 0.75, c: colorDml(lnRef) || "#000000" } : null;
      if (hijo(ln, "noFill")) return { no: 1 };
      const g = at(hijo(ln, "prstDash"), "val");
      return { w: at(ln, "w") ? +at(ln, "w") / 12700 : 0.75, c: colorDml(hijo(ln, "solidFill")) || (lnRef ? colorDml(lnRef) : null) || "#000000", guion: g && g !== "solid" ? g : null, fin: at(hijo(ln, "tailEnd"), "type") || null, ini: at(hijo(ln, "headEnd"), "type") || null };
    };
    function dibujosDeHoja(ruta) {
      const relsH = relsDeParte(ruta), out = [];
      // controles de formulario de la hoja: shapeId → su tipo y su estado (ctrlProps)
      const controles = {};
      {
        const xh = z.texto(ruta), dh = xh && xh.includes("<control") ? P(xh) : null;
        todos(dh, "control").forEach((c) => {
          const rid = c.getAttribute("r:id") || c.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"), rel = rid && relsH[rid];
          const xp = rel && rel.ruta ? z.texto(rel.ruta) : null, fp = xp ? P(xp).documentElement : null;
          if (!fp || !at(c, "shapeId")) return;
          controles[at(c, "shapeId")] = { tipo: at(fp, "objectType") || "Button", marcado: at(fp, "checked") === "Checked" ? 1 : at(fp, "checked") === "Mixed" ? 2 : 0,
            sel: at(fp, "sel") != null ? +at(fp, "sel") : null, rango: at(fp, "fmlaRange") || null, enlace: at(fp, "fmlaLink") || null, valor: at(fp, "val") != null ? +at(fp, "val") : null, texto: at(c, "name") || "" };
        });
      }
      Object.values(relsH).filter((r) => /\/drawing$/.test(r.tipo) && r.ruta).forEach((r) => {
        const xd = z.texto(r.ruta);
        if (!xd) return;
        const relsD = relsDeParte(r.ruta);
        const pos = (n) => (n ? { c: +(hijo(n, "col") || {}).textContent || 0, dc: (+(hijo(n, "colOff") || {}).textContent || 0) / 9525, f: +(hijo(n, "row") || {}).textContent || 0, df: (+(hijo(n, "rowOff") || {}).textContent || 0) / 9525 } : null);
        const leerEl = (el, ancla, grupo) => {
          if (!el) return;
          const nm = el.localName;
          if (nm === "AlternateContent") { const ch = hijo(el, "Choice") || hijo(el, "Fallback"); hijos(ch).forEach((x) => leerEl(x, ancla, grupo)); return; }
          // posición dentro de un grupo: su xfrm en el espacio del grupo
          const xf = todos(hijo(el, "spPr") || hijo(el, "grpSpPr") || hijo(el, "xfrm") || el, "xfrm")[0];
          const o2 = xf && hijo(xf, "off"), e2 = xf && hijo(xf, "ext");
          const a = grupo ? { grupo: grupo.idx, caja: o2 && e2 ? { x: +at(o2, "x"), y: +at(o2, "y"), w: +at(e2, "cx"), h: +at(e2, "cy") } : null } : ancla;
          const giro = xf && at(xf, "rot") ? +at(xf, "rot") / 60000 : 0, vH = xf && at(xf, "flipH") === "1", vV = xf && at(xf, "flipV") === "1";
          if (nm === "grpSp") {
            const gx = todos(hijo(el, "grpSpPr"), "xfrm")[0], go = gx && hijo(gx, "off"), ge = gx && hijo(gx, "ext"), co = gx && hijo(gx, "chOff"), ce = gx && hijo(gx, "chExt");
            if (!go || !ge || !co || !ce) return;
            // el grupo ocupa su ancla; sus hijos van en su espacio (chOff, chExt)
            const sub = { ancla: a, grupoEn: grupo, chx: +at(co, "x"), chy: +at(co, "y"), cx: +at(ce, "cx") || 1, cy: +at(ce, "cy") || 1, ex: +at(ge, "cx") || 1, ey: +at(ge, "cy") || 1 };
            out.push({ tipo: "grupo-ini", a, sub: { chx: sub.chx, chy: sub.chy, cx: sub.cx, cy: sub.cy, giro: gx && at(gx, "rot") ? +at(gx, "rot") / 60000 : 0 } });
            const gIdx = out.length - 1;
            hijos(el).forEach((x) => { if (x.localName !== "nvGrpSpPr" && x.localName !== "grpSpPr") leerEl(x, null, { idx: gIdx }); });
            return;
          }
          const base = { a: grupo ? { grupo: grupo.idx, caja: xf && o2 && e2 ? { x: +at(o2, "x"), y: +at(o2, "y"), w: +at(e2, "cx"), h: +at(e2, "cy") } : null } : a, giro: giro || undefined, vH: vH || undefined, vV: vV || undefined };
          if (nm === "pic") {
            const bl = todos(el, "blip")[0], rid = bl && (bl.getAttribute("r:embed") || bl.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed"));
            const src = rid && relsD[rid] ? dataUrlXl(relsD[rid].ruta) : null;
            if (!src) return;
            const sr = todos(hijo(el, "blipFill"), "srcRect")[0];
            const lnP = lineaDml(hijo(el, "spPr"));
            out.push(Object.assign({ tipo: "imagen", src, alt: (at(todos(el, "cNvPr")[0], "descr") || "").trim() || undefined, recorte: sr ? ["l", "t", "r", "b"].map((k) => +(at(sr, k) || 0) / 100000) : undefined, linea: lnP && !lnP.no ? lnP : undefined }, base));
            return;
          }
          if (nm === "graphicFrame") {
            const ch = todos(el, "chart")[0], rid = ch && (ch.getAttribute("r:id") || ch.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"));
            const rel = rid && relsD[rid];
            if (rel && /\/chart$/.test(rel.tipo) && NOM.graficaOffice) {
              const xc = z.texto(rel.ruta);
              if (xc) { try { out.push(Object.assign({ tipo: "grafica", g: NOM.graficaOffice.leer(xc, { color: colorDml, fuente: fuenteDml, acentos: ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((k) => "#" + (temaDml[k] || "4472C4")) }) }, base)); } catch (e) { /* una gráfica rara no detiene la hoja */ } }
              return;
            }
            // (versión 4) gráficas nuevas de Excel (histograma, Pareto, cajas, cascada, rectángulos, solar, embudo):
            // sus datos vienen de nombres ocultos del libro que apuntan a la hoja; se dibujan con los valores vivos
            if (V >= 4 && rel && /chartEx$/i.test(rel.tipo || "") && NOM.graficaOffice && NOM.graficaOffice.leerEx) {
              const xc = z.texto(rel.ruta);
              if (xc) {
                try {
                  out.push(Object.assign({ tipo: "grafica", g: NOM.graficaOffice.leerEx(xc, { color: colorDml, fuente: fuenteDml, nombre: (n) => nombresXl[String(n).toLowerCase()] || null,
                    acentos: ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((k) => "#" + (temaDml[k] || "4472C4")) }) }, base));
                  return;
                } catch (e) { /* una que no se entiende queda como recuadro */ }
              }
            }
            // SmartArt: Excel guarda al lado el dibujo ya acomodado; sus formas van como un grupo del tamaño del marco
            const relIds = todos(el, "relIds")[0];
            if (relIds) {
              const xfG = hijo(el, "xfrm"), eG = xfG && hijo(xfG, "ext");
              const dsp = Object.values(relsD).find((q) => /diagramDrawing$/.test(q.tipo) && q.ruta);
              const xdg = dsp && z.texto(dsp.ruta);
              if (xdg && eG) {
                // (Excel escribe 0 × 0 en el marco: el tamaño está en el ancla; las formas van en EMU desde la esquina del marco)
                out.push({ tipo: "grupo-ini", a, sub: { chx: 0, chy: 0, cx: +at(eG, "cx") || 0, cy: +at(eG, "cy") || 0, giro: 0, emu: true } });
                const gI = out.length - 1, arbolD = todos(P(xdg), "spTree")[0];
                hijos(arbolD).forEach((x) => { if (x.localName === "sp" || x.localName === "grpSp") leerEl(x, null, { idx: gI }); });
                return;
              }
            }
            // gráficas nuevas de Excel (cascada, rectángulos, embudo…): un recuadro con su título por ahora
            if (todos(el, "chart").length || /chartex/i.test(String(rel && rel.tipo))) out.push(Object.assign({ tipo: "grafica-ex" }, base));
            return;
          }
          if (nm === "sp" || nm === "cxnSp") {
            const spPr = hijo(el, "spPr"), prst = at(todos(spPr, "prstGeom")[0], "prst") || "rect", st = hijo(el, "style");
            const tx = hijo(el, "txBody"), parrafos = tx ? hijos(tx, "p").map((p) => ({ t: todos(p, "t").map((x) => x.textContent).join(""), al: at(todos(p, "pPr")[0], "algn") || null })) : [];
            const r0 = tx && todos(tx, "rPr")[0], dr = tx && todos(tx, "defRPr")[0], rr = r0 || dr;
            const letra = rr ? { tam: at(rr, "sz") ? +at(rr, "sz") / 100 : 11, b: at(rr, "b") === "1", i: at(rr, "i") === "1", c: colorDml(hijo(rr, "solidFill")) || (st ? colorDml(hijo(st, "fontRef")) : null) || "#000000" } : { tam: 11, c: st ? colorDml(hijo(st, "fontRef")) || "#000000" : "#000000" };
            if (rr) {
              // WordArt: el contorno de la letra (su color, aunque sea un degradado) y su sombra; y su tipo de letra
              const lnT = hijo(rr, "ln"); if (lnT && !hijo(lnT, "noFill")) { const cT = colorDml(hijo(lnT, "solidFill") || todos(lnT, "gs")[0]); if (cT) letra.contorno = { c: cT, w: at(lnT, "w") ? +at(lnT, "w") / 12700 : 0.75 }; }
              const sh = todos(hijo(rr, "effectLst"), "outerShdw")[0]; if (sh) letra.sombra = { c: colorDml(sh) || "#000000", d: (+(at(sh, "dist") || 0)) / 12700, ang: (+(at(sh, "dir") || 0)) / 60000, b: (+(at(sh, "blurRad") || 0)) / 12700 };
              const la = hijo(rr, "latin"); if (la && at(la, "typeface")) letra.f = fuenteDml(at(la, "typeface"));
            }
            const bp = tx && hijo(tx, "bodyPr");
            const ins = bp ? ["lIns", "tIns", "rIns", "bIns"].map((k, i) => (at(bp, k) != null ? +at(bp, k) / 9525 : [9.6, 4.8, 9.6, 4.8][i])) : [9.6, 4.8, 9.6, 4.8];
            // controles de formulario (casilla, lista, botón, opción…): Excel los dibuja él; aquí con su estado
            const idSp = at(todos(el, "cNvPr")[0], "id"), ctl = idSp && controles[idSp];
            if (ctl) { out.push(Object.assign({ tipo: "control", ctl, texto: todos(el, "t").map((x) => x.textContent).join("") || ctl.texto || "" }, base)); return; }
            if (at(todos(el, "cNvPr")[0], "hidden") === "1") return;
            out.push(Object.assign({ tipo: nm === "cxnSp" ? "linea" : "forma", prst, relleno: rellenoDml(spPr) || (st && +at(hijo(st, "fillRef"), "idx") ? colorDml(hijo(st, "fillRef")) : null), linea: lineaDml(spPr, st),
              texto: parrafos.some((p) => p.t) ? parrafos : undefined, letra, ancla: bp ? at(bp, "anchor") || "t" : "t", ins, ajusta: !(bp && at(bp, "wrap") === "none"), av: (() => { const o = {}; todos(todos(spPr, "prstGeom")[0], "gd").forEach((g) => { const m = /^val\s+(-?\d+)/.exec(at(g, "fmla") || ""); if (m) o[at(g, "name")] = +m[1]; }); return Object.keys(o).length ? o : undefined; })() }, base));
          }
        };
        // (los controles de formulario y otras cosas nuevas vienen dentro de «mc:AlternateContent»: se usa su Choice)
        const anclas = [];
        hijos(P(xd).documentElement).forEach((x) => { if (x.localName === "AlternateContent") { const ch = hijo(x, "Choice") || hijo(x, "Fallback"); hijos(ch).forEach((y) => anclas.push(y)); } else anclas.push(x); });
        anclas.forEach((anc) => {
          const tipoA = anc.localName;
          if (!/Anchor$/.test(tipoA)) return;
          const a = tipoA === "twoCellAnchor" ? { de: pos(hijo(anc, "from")), a: pos(hijo(anc, "to")) }
            : tipoA === "oneCellAnchor" ? { de: pos(hijo(anc, "from")), w: +at(hijo(anc, "ext"), "cx") / 9525, h: +at(hijo(anc, "ext"), "cy") / 9525 }
            : { abs: { x: +at(hijo(anc, "pos"), "x") / 9525, y: +at(hijo(anc, "pos"), "y") / 9525, w: +at(hijo(anc, "ext"), "cx") / 9525, h: +at(hijo(anc, "ext"), "cy") / 9525 } };
          hijos(anc).forEach((el) => { if (["pic", "sp", "graphicFrame", "grpSp", "cxnSp", "AlternateContent"].includes(el.localName)) leerEl(el, a, null); });
        });
      });
      return out.length ? out : null;
    }

    // ---- Formato condicional de Excel (versión 3): las reglas tal cual (con sus fórmulas en español y relativas a
    // la primera celda de su rango); las pinta el modo Excel de la hoja. Las de Excel 2010 (barras con borde o lisas,
    // más juegos de íconos) van aparte en «extLst» y se juntan por su id.
    function formatoCondicional(doc, hoja) {
      const raiz = doc.documentElement, reglas = [], usados = new Set(), ext14 = {};
      const X14 = (n) => n && /2009\/9\/main/.test(n.namespaceURI || "");
      todos(raiz, "cfRule").forEach((r) => { if (X14(r) && at(r, "id")) ext14[at(r, "id")] = r; });
      const cfvo = (n) => { const fx = todos(n, "f")[0]; return { t: at(n, "type") || "min", v: at(n, "val") != null ? at(n, "val") : fx ? fx.textContent : null, gte: at(n, "gte") !== "0" }; };
      const formulaCf = (fx, r0) => {
        let x = String(fx || "").trim().replace(/_xlfn\.|_xlws\./g, "");
        const p0 = F().parseRef(String(r0 || "A1").split(":")[0].replace(/\$/g, "")) || { c: 0, f: 0 };
        try { x = sinNombres(sinEstructuradas(x, hoja, p0.c, p0.f)); } catch (e) { /* tal cual */ }
        return F().aEspanol("=" + x);
      };
      const leer = (x, r, d14) => {
        const t = at(x, "type"), o = { r, t, p: +(at(x, "priority") || 0), s: at(x, "stopIfTrue") === "1" };
        if (at(x, "dxfId") != null && dxfsL[+at(x, "dxfId")]) o.d = formatoDxf(dxfsL[+at(x, "dxfId")]);
        if (at(x, "operator")) o.o = at(x, "operator");
        if (at(x, "text") != null) o.tx = at(x, "text");
        const fs = hijos(x, "formula").concat(hijos(x, "f")).map((n) => formulaCf(n.textContent, r[0]));
        if (fs.length) o.fx = fs;
        if (t === "top10") { o.rank = +(at(x, "rank") || 10); o.pct = at(x, "percent") === "1"; o.bottom = at(x, "bottom") === "1"; }
        if (t === "aboveAverage") { o.above = at(x, "aboveAverage") !== "0"; o.eq = at(x, "equalAverage") === "1"; o.std = +(at(x, "stdDev") || 0); }
        if (t === "timePeriod") o.tp = at(x, "timePeriod");
        const cs = hijo(x, "colorScale");
        if (cs) o.cs = { v: hijos(cs, "cfvo").map(cfvo), c: hijos(cs, "color").map((n) => color(n) || "#FFFFFF") };
        const db = hijo(x, "dataBar") || (d14 && hijo(d14, "dataBar"));
        if (db) {
          o.db = { v: hijos(db, "cfvo").map(cfvo), c: color(hijo(db, "color") || hijo(db, "fillColor")) || "#638EC6", min: +(at(db, "minLength") || 10), max: +(at(db, "maxLength") || 90), grad: true, show: at(db, "showValue") !== "0" };
          const e = d14 && hijo(d14, "dataBar");
          if (e) {
            if (at(e, "minLength") != null) o.db.min = +at(e, "minLength");
            if (at(e, "maxLength") != null) o.db.max = +at(e, "maxLength");
            o.db.grad = at(e, "gradient") !== "0";
            if (at(e, "border") === "1") o.db.borde = color(hijo(e, "borderColor")) || o.db.c;
            o.db.neg = at(e, "negativeBarColorSameAsPositive") === "1" ? o.db.c : color(hijo(e, "negativeFillColor")) || "#FF0000";
            o.db.eje = at(e, "axisPosition") || "automatic";
            o.db.ejeColor = color(hijo(e, "axisColor")) || "#000000";
            if (at(e, "direction")) o.db.dir = at(e, "direction");
            const v14 = hijos(e, "cfvo").map(cfvo); if (v14.length === 2) o.db.v = v14;
          }
        }
        const is = hijo(x, "iconSet");
        if (is) o.is = { set: at(is, "iconSet") || "3TrafficLights1", rev: at(is, "reverse") === "1", show: at(is, "showValue") !== "0", v: hijos(is, "cfvo").map(cfvo) };
        return o;
      };
      hijos(raiz, "conditionalFormatting").forEach((cf) => {
        const r = String(at(cf, "sqref") || "").split(/\s+/).filter(Boolean);
        if (!r.length) return;
        hijos(cf, "cfRule").forEach((x) => {
          const id = todos(x, "id")[0], d14 = id ? ext14[id.textContent.trim()] : null;
          if (d14) usados.add(d14);
          reglas.push(leer(x, r, d14));
        });
      });
      // reglas que solo están en las extensiones de Excel 2010 (íconos como estrellas, triángulos o cajas)
      todos(raiz, "conditionalFormatting").filter(X14).forEach((cf) => {
        const sq = todos(cf, "sqref")[0], r = String(sq ? sq.textContent : "").split(/\s+/).filter(Boolean);
        if (!r.length) return;
        hijos(cf, "cfRule").filter((x) => !usados.has(x)).forEach((x) => reglas.push(leer(x, r, x)));
      });
      return reglas.length ? reglas : null;
    }
    /** Botones de filtro de Excel en el encabezado de un filtro (de tabla o de la hoja): sin los que el archivo esconde
        (hiddenButton) y con el embudo donde hay un filtro puesto. */
    function botonesFiltro(af, c1, c2, f, campos) {
      const cols = {};
      hijos(af, "filterColumn").forEach((fc) => { cols[+at(fc, "colId")] = at(fc, "hiddenButton") === "1" ? "oculto" : hijos(fc).length ? "activo" : "normal"; });
      for (let c = c1; c <= c2 && c < MAX_C; c++) { const e = cols[c - c1]; if (e !== "oculto") campos["fb:" + F().ref(c, f)] = e === "activo" ? 2 : true; }
    }
    function estilosDeTabla(hoja, campos) {
      const vistas = new Set();
      for (const tb of Object.values(tablasXl)) {
        if (tb.hoja !== hoja || vistas.has(tb)) continue;
        vistas.add(tb);
        if (tb.filtro && tb.enc) botonesFiltro(tb.filtro, tb.c1, tb.c2, tb.f1, campos);   // botones de filtro
        if (!tb.estilo) continue;
        const est2 = estilosTablaLibro[tb.estilo] || (NOM.ESTILOS_TABLA || {})[tb.estilo];
        if (!est2) continue;
        pintarTablaXl(tb, est2, colorTabla, campos);
      }
    }
    // ---- Tablas dinámicas (versión 3): su estilo pintado sobre sus áreas (filtros, encabezado, etiquetas de fila,
    // subtítulos y subtotales, total general de filas y de columnas), con la precedencia de Excel.
    function estilosDinamicas(ruta, campos) {
      const relsH = z.texto(ruta.replace(/([^/]+)$/, "_rels/$1.rels"));
      if (!relsH) return;
      todos(P(relsH), "Relationship").filter((r) => /\/pivotTable$/.test(at(r, "Type") || "")).forEach((r) => {
        const tg = at(r, "Target") || "", rt = tg.startsWith("/") ? tg.slice(1) : ruta.replace(/[^/]+$/, "") + tg;
        const norm = rt.split("/").reduce((a, x) => { if (x === "..") a.pop(); else if (x && x !== ".") a.push(x); return a; }, []).join("/");
        const tx = z.texto(norm);
        if (!tx) return;
        const pd = P(tx).documentElement, si = hijo(pd, "pivotTableStyleInfo");
        if (!hijo(pd, "location") || !si) return;
        const est2 = estilosTablaLibro[at(si, "name")] || (NOM.ESTILOS_TABLA || {})[at(si, "name")];
        if (!est2) return;
        pintarDinamicaXl(pd, est2, colorTabla, campos);
      });
    }
    const hojaRef = (n) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
    // Nombres definidos del libro (IVA = 0.16, Ventas = Datos!$A$1:$G$61…) y el orden de las hojas (referencias 3D)
    const nombresXl = {};
    if (V >= 2) todos(P(libro), "definedName").forEach((d) => {
      const n = at(d, "name") || "";
      if (!n || /^_xlnm\./i.test(n) || at(d, "localSheetId") != null) return;
      nombresXl[n.toLowerCase()] = d.textContent.trim();
    });
    const ordenHojas = hojasXml.map((h) => at(h, "name") || "");
    /** Lo que el motor no entiende escrito como Excel lo guarda → lo mismo con referencias normales: referencias 3D
        (Enero:Marzo!B2 → Enero!B2,Febrero!B2,Marzo!B2) y nombres definidos (IVA → 0.16). */
    function sinNombres(fx) {
      if (V < 2) return fx;
      const limpia = (h) => h.replace(/^'|'$/g, "").replace(/''/g, "'");
      const tresD = (todo, h1, h2, ref) => {
        const i1 = ordenHojas.indexOf(limpia(h1)), i2 = ordenHojas.indexOf(limpia(h2));
        if (i1 < 0 || i2 < 0) return todo;
        return ordenHojas.slice(Math.min(i1, i2), Math.max(i1, i2) + 1).map((h) => hojaRef(h) + "!" + ref).join(",");
      };
      // Con espacios u otros signos, Excel escribe las dos hojas dentro de UNAS comillas: 'x02 Enero:x02 Marzo'!B2 (un nombre
      // de hoja no puede llevar «:», así que dentro de las comillas los dos puntos siempre separan la primera de la última)
      let r = fx.replace(/'((?:[^':]|'')+):((?:[^':]|'')+)'!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/g, tresD);
      r = r.replace(/('[^']+'|[A-Za-z_][\w.]*):('[^']+'|[A-Za-z_][\w.]*)!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/g, tresD);
      if (Object.keys(nombresXl).length && V < 7) {
        // fuera de los textos entre comillas: un identificador que es un nombre definido (y no una función)
        r = r.split(/("(?:[^"]|"")*")/).map((parte, i) => (i % 2 ? parte : parte.replace(/(^|[^\w.!'$\]])([A-Za-z_\\][\w.]*)(?![\w.(!\[])/g, (todo, antes, n) => {
          const d = nombresXl[n.toLowerCase()];
          return d != null ? antes + (/^[-\d.]+$/.test(d) ? d : "(" + d + ")") : todo;
        }))).join("");
        // (versión 5) una función propia guardada como nombre (MiFuncion = LAMBDA(x, x*2)): MiFuncion(5) → (LAMBDA(…))(5)
        if (V >= 5) r = r.split(/("(?:[^"]|"")*")/).map((parte, i) => (i % 2 ? parte : parte.replace(/(^|[^\w.!'$\]])([A-Za-z_\\][\w.]*)(?=\s*\()/g, (todo, antes, n) => {
          const d = nombresXl[n.toLowerCase()];
          return d != null && /^(_xlfn\.)?LAMBDA\s*\(/i.test(d) ? antes + "(" + d + ")" : todo;
        }))).join("");
      }
      return r;
    }
    /** Fórmula con referencias estructuradas → la misma con referencias normales (para la celda c,f de la hoja). */
    function sinEstructuradas(fx, hojaAct, c, f) {
      fx = sinNombres(fx);
      return F().sinEstructuradas(fx, tablasXl, hojaAct, c, f, hojaRef);
    }
    // La fuente del libro (la primera de styles.xml: Aptos Narrow 11 en Office 2024) define el ancho de las
    // columnas (el ancho del dígito más ancho) y el alto de las filas, como en Excel.
    const fuenteLibro = fuentes[0] || { nombre: "Calibri", sz: 11 };
    const MDW = Math.max(5, Math.round((fuenteLibro.sz || 11) * 96 / 72 * 0.4885));   // ≈7 px con 11 pt
    const pxCol = (w) => Math.floor(((256 * w + Math.floor(128 / MDW)) / 256) * MDW);   // fórmula de Excel
    let aviso = null;
    const MAX_F = LIM_F, MAX_C = LIM_C;
    const res = [];
    for (const h of hojasXml.slice(0, 30)) {
      const nombre = at(h, "name") || "Hoja" + (res.length + 1);
      let ruta = destino[h.getAttribute("r:id") || h.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")] || "";
      ruta = ruta.startsWith("/") ? ruta.slice(1) : "xl/" + ruta.replace(/^\.\//, "");
      const xmlH = z.texto(ruta);
      if (!xmlH) continue;
      const doc = P(xmlH);
      const campos = { nombre: nombre.slice(0, 31), ox: ruta, xlv: V };
      // Cómo se ve la hoja en Excel: fuente y tamaño del libro, ancho y alto por omisión, cuadrícula, oculta…
      {
        const fp = todos(doc, "sheetFormatPr")[0], sv = todos(doc, "sheetView")[0];
        const dcw = +at(fp, "defaultColWidth"), bcw = at(fp, "baseColWidth"), drh = +at(fp, "defaultRowHeight");
        const w0 = dcw ? pxCol(dcw) : bcw != null ? Math.round((+bcw + 0.43) * MDW + 5) : Math.round(8.43 * MDW + 5);
        const estado = at(h, "state");
        campos.xl = { fuente: fuenteLibro.nombre || "Calibri", tam: fuenteLibro.sz || 11, w0, h0: Math.round((drh || 15) * 96 / 72),
          rejilla: !(sv && at(sv, "showGridLines") === "0"), rtl: !!(sv && at(sv, "rightToLeft") === "1"), zoom: sv && +at(sv, "zoomScale") ? +at(sv, "zoomScale") : 100,
          oculta: estado === "hidden" || estado === "veryHidden" };
        if (V >= 13 && temaXl.some(Boolean)) campos.xl.tema = temaXl.map((c) => c || null);
        if (V >= 9 && sv && /^(1|true)$/.test(at(sv, "showFormulas") || "")) campos.vf = true;
      }
      let maxF = 0, maxC = 0;
      const compartidas = {}, tablasDatos = [];
      const derramadas = new Set();   // (versión 5) celdas a las que llega lo que derrama una fórmula de matriz
      todos(doc, "row").forEach((row) => {
        const ht = +at(row, "ht");
        const rnum = +at(row, "r");
        // (versión 4: Excel dibuja la fila con los píxeles enteros que caben en su alto: 80.1 pt → 106 px, no 107)
        const pxFila = (x) => (V >= 4 ? Math.floor(x / 0.75 + 1e-6) : Math.round(x / 0.75));
        if (at(row, "customHeight") === "1" && ht && rnum && rnum <= MAX_F) campos["ah:" + rnum] = pxFila(ht);
        else if (ht && rnum && rnum <= MAX_F) campos["ahx:" + rnum] = pxFila(ht);   // alto que Excel calculó (letra grande, texto ajustado)
        if (at(row, "hidden") === "1" && rnum && rnum <= MAX_F) campos["oh:" + rnum] = true;
        if (V >= 8 && +at(row, "outlineLevel") > 0 && rnum && rnum <= MAX_F) campos["ol:" + rnum] = Math.min(7, +at(row, "outlineLevel"));
        hijos(row, "c").forEach((c) => {
          const ref = at(c, "r"), p = F().parseRef(ref || "");
          if (!p) return;
          if (p.f >= MAX_F || p.c >= MAX_C) { aviso = `Solo se importaron las primeras ${MAX_F} filas y ${MAX_C} columnas`; return; }
          const tipo = at(c, "t"), v = hijo(c, "v"), f = hijo(c, "f"), s = xfs[+at(c, "s") || 0];
          let crudo = "";
          // (versión 10) tabla de datos: todo su lugar lleva la fórmula TABLA (las demás celdas solo traen su valor)
          if (V >= 10 && f && at(f, "t") === "dataTable") {
            const [a1, b1] = String(at(f, "ref") || "").split(":").map((q) => F().parseRef(q));
            const r1 = at(f, "r1") || "", r2 = at(f, "r2") || "", dos = at(f, "dt2D") === "1", fila = at(f, "dtr") === "1";
            const fx = dos ? `=TABLA(${r1};${r2})` : fila ? `=TABLA(${r1};)` : `=TABLA(;${r1})`;
            if (a1) tablasDatos.push({ a: a1, b: b1 || a1, fx });
          }
          if (f && (f.textContent || at(f, "t") === "shared")) {
            let fx = f.textContent;
            if (at(f, "t") === "shared") {
              const si = at(f, "si");
              if (fx) compartidas[si] = { f: "=" + fx, c: p.c, r: p.f };
              else if (compartidas[si]) fx = F().desplazar(compartidas[si].f, p.c - compartidas[si].c, p.f - compartidas[si].r).slice(1);
            }
            if (fx && V >= 2) fx = sinEstructuradas(fx, nombre, p.c, p.f);
            // (versión 5) matrices dinámicas: A1# (ANCHORARRAY), «@» (SINGLE) y los parámetros de LAMBDA sin «_xlpm.»
            if (fx && V >= 5) fx = fx.replace(/_xlfn\.ANCHORARRAY\(([^()]+)\)/gi, "$1#").replace(/_xlfn\.SINGLE\(/gi, "@(").replace(/_xlop\.([A-Za-z_][\w.]*)/gi, "[$1]").replace(/_xlpm\./gi, "");
            crudo = fx ? F().aEspanol("=" + fx.replace(/_xlfn\.|_xlws\./g, ""), V >= 5) : "";
            // (versión 5) sin t="array" es una fórmula del Excel viejo: «@» donde Excel 365 hace la intersección implícita
            if (crudo && V >= 5 && at(f, "t") !== "array") crudo = F().conInterseccion(crudo);
            // (versión 5) una fórmula de matriz: las demás celdas de su lugar son lo que derrama (no se guardan sueltas)
            if (V >= 5 && at(f, "t") === "array") {
              const [a1, b1] = String(at(f, "ref") || "").split(":").map((q) => F().parseRef(q));
              if (a1 && b1 && (b1.c > a1.c || b1.f > a1.f)) {
                campos["fa:" + F().ref(p.c, p.f)] = `${F().ref(a1.c, a1.f)}:${F().ref(b1.c, b1.f)}`;
                for (let ff = a1.f; ff <= Math.min(b1.f, MAX_F - 1); ff++) for (let cc = a1.c; cc <= Math.min(b1.c, MAX_C - 1); cc++) if (ff !== p.f || cc !== p.c) derramadas.add(F().ref(cc, ff));
                maxF = Math.max(maxF, Math.min(b1.f + 1, MAX_F)); maxC = Math.max(maxC, Math.min(b1.c + 1, MAX_C));
              }
            }
          } else if (V >= 5 && derramadas.has(F().ref(p.c, p.f))) crudo = "";
          else if (tipo === "s") crudo = textos[+(v ? v.textContent : 0)] || "";
          else if (tipo === "inlineStr") crudo = todos(c, "t").map((t) => t.textContent).join("");
          else if (tipo === "b") crudo = v && v.textContent === "1" ? "VERDADERO" : "FALSO";
          else if (tipo === "str" || tipo === "e") crudo = v ? v.textContent : "";
          else if (v) crudo = v.textContent;
          const R = F().ref(p.c, p.f);
          // Un texto que parece número o fecha se queda como texto (formato @).
          const esTexto = (tipo === "s" || tipo === "inlineStr") && crudo !== "" && F().esNum(F().calcularLiteral(crudo.trim()));
          // y un texto que parece número, fecha o fórmula lleva «'» (el apóstrofo de Excel): se ve tal cual
          const ambiguo = (tipo === "s" || tipo === "inlineStr" || tipo === "str") && !f && crudo !== "" && (esTexto || /^[=']/.test(crudo) || F().calcularLiteral(crudo.trim()) !== crudo.trim());
          if (crudo !== "") campos["c:" + R] = ambiguo && V >= 2 ? "'" + crudo : crudo;
          if (s) {
            const cod = esTexto ? "@" : fmts[s.num];
            if (cod && cod !== "General") campos["f:" + R] = cod.replace(/\[\$-[^\]]*\]/g, "").replace(/;@$/, "");
            if (s.fuente.b) campos["n:" + R] = true;
            if (s.fuente.i) campos["i:" + R] = true;
            if (s.fuente.u) campos["u:" + R] = true;
            if (s.fuente.k && s.fuente.k !== "#000000") campos["k:" + R] = s.fuente.k;
            if (s.fuente.sz && Math.abs(s.fuente.sz - 11) > 0.1) campos["tl:" + R] = Math.round((s.fuente.sz / 1.0476) * 2) / 2;
            if (s.fuente.sz && Math.abs(s.fuente.sz - (fuenteLibro.sz || 11)) > 0.01) campos["tp:" + R] = s.fuente.sz;   // tamaño exacto (modo Excel)
            // tipo de letra de la celda, si no es el del libro (el de su primera fuente)
            if (s.fuente.nombre && s.fuente.nombre !== ((fuentes[0] || {}).nombre || "Calibri")) campos["tf:" + R] = s.fuente.nombre;
            if (s.relleno) campos["r:" + R] = s.relleno;
            if (s.borde) campos["bd:" + R] = s.borde.split("").sort().join("");
            if (V >= 3 && s.bx) campos["bx:" + R] = s.bx;
            if (V >= 4 && s.ax) campos["ax:" + R] = s.ax;
            if (V >= 4 && s.rx) campos["rx:" + R] = s.rx;
            if (V >= 6 && s.pr) campos["pr:" + R] = s.pr;
            if (s.al) campos["al:" + R] = s.al;
            if (s.aj) campos["aj:" + R] = true;
          }
          if (crudo !== "" || s) { maxF = Math.max(maxF, p.f + 1); maxC = Math.max(maxC, p.c + 1); }
        });
      });
      todos(doc, "mergeCell").forEach((m) => {
        const [a, b] = String(at(m, "ref") || "").split(":").map((r) => F().parseRef(r));
        if (a && b && a.f < MAX_F && a.c < MAX_C) campos["m:" + F().ref(a.c, a.f)] = `${b.c - a.c + 1}x${b.f - a.f + 1}`;
      });
      todos(doc, "col").forEach((c) => {
        const w = +at(c, "width"), min = +at(c, "min"), max = Math.min(+at(c, "max"), MAX_C);
        if (!w || at(c, "customWidth") !== "1") return;
        for (let k = min; k <= max && k - min < MAX_C; k++) { campos["aw:" + F().colLetra(k - 1)] = Math.round(w * 7 + 5); campos["ap:" + F().colLetra(k - 1)] = pxCol(w); }
      });
      todos(doc, "col").forEach((c) => {
        const min = +at(c, "min"), max = Math.min(+at(c, "max"), MAX_C), nv = V >= 8 ? Math.min(7, +at(c, "outlineLevel") || 0) : 0;
        if (at(c, "hidden") === "1") for (let k = min; k <= max && k - min < MAX_C; k++) campos["ow:" + F().colLetra(k - 1)] = true;
        if (nv > 0) for (let k = min; k <= max && k - min < MAX_C; k++) campos["olc:" + F().colLetra(k - 1)] = nv;
      });
      // (versión 10) las celdas de cada tabla de datos: su fórmula TABLA (Excel solo guarda la primera)
      tablasDatos.forEach((td) => { for (let f2 = td.a.f; f2 <= td.b.f; f2++) for (let c2 = td.a.c; c2 <= td.b.c; c2++) campos["c:" + F().ref(c2, f2)] = td.fx; });
      // (versión 6) la protección de la hoja: el hash de su contraseña y lo que deja hacer
      if (V >= 6) { const sp = todos(doc, "sheetProtection")[0]; if (sp && /^(1|true)$/.test(at(sp, "sheet") || "")) campos.prot = protHojaDe(sp); }
      // (versión 11) los escenarios (Datos → Análisis de hipótesis → Administrador de escenarios)
      if (V >= 11) { const se = todos(doc, "scenarios")[0], e = se ? escenariosDe(se) : null; if (e) campos.esc = e; }
      // (versión 12) los minigráficos (Insertar → Minigráficos): sus grupos, colores, marcas y de dónde salen sus datos
      if (V >= 12) {
        const gs = [...doc.getElementsByTagNameNS("*", "sparklineGroup")].map((g) => {
          const r = { tipo: at(g, "type") === "column" ? "column" : at(g, "type") === "stacked" ? "stacked" : "line", c: {}, s: [] };
          COLORES_MG.forEach(([k, el]) => { const n = [...g.children].find((q) => q.localName === el); const cc = n ? color(n) : null; if (cc) r.c[k] = cc; });
          BANDERAS_MG.forEach(([k, a]) => { if (/^(1|true)$/.test(at(g, a) || "")) r[k] = 1; });
          if (+at(g, "lineWeight")) r.grosor = +at(g, "lineWeight");
          if (at(g, "displayEmptyCellsAs")) r.vacias = at(g, "displayEmptyCellsAs");
          [...g.getElementsByTagNameNS("*", "sparkline")].forEach((sp) => {
            const f = [...sp.children].find((q) => q.localName === "f"), sq = [...sp.children].find((q) => q.localName === "sqref");
            if (sq && sq.textContent.trim()) r.s.push([sq.textContent.trim(), f ? f.textContent.trim() : ""]);
          });
          return r.s.length ? r : null;
        }).filter(Boolean);
        if (gs.length) campos.mg = gs;
      }
      const pane = todos(doc, "pane")[0];
      if (pane && at(pane, "state") === "frozen") { if (+at(pane, "ySplit")) campos.fijasF = +at(pane, "ySplit"); if (+at(pane, "xSplit")) campos.fijasC = +at(pane, "xSplit"); }
      todos(doc, "dataValidation").forEach((dv, k) => {
        const sq = String(at(dv, "sqref") || "").split(" ")[0], f1 = hijo(dv, "formula1");
        if (!sq) return;
        const rango = sq.includes(":") ? sq : sq + ":" + sq;
        if (at(dv, "type") === "list" && f1 && /^".*"$/.test(f1.textContent)) campos["dv:x" + k] = { rango, tipo: "lista", valores: f1.textContent.slice(1, -1).split(",").map((s) => s.trim()).filter(Boolean) };
        else if (["decimal", "whole"].includes(at(dv, "type")) && f1) {
          const f2 = hijo(dv, "formula2"), op = at(dv, "operator") || "between";
          campos["dv:x" + k] = { rango, tipo: "numero", min: op === "lessThanOrEqual" ? "" : f1.textContent, max: op === "between" && f2 ? f2.textContent : op === "lessThanOrEqual" ? f1.textContent : "" };
        }
      });
      if (V >= 3) {
        const af = hijo(doc.documentElement, "autoFilter"), [q1, q2] = String(at(af, "ref") || "").split(":").map((q) => F().parseRef(q));
        if (q1) botonesFiltro(af, q1.c, (q2 || q1).c, q1.f, campos);
      }
      if (V >= 3) { const dib = dibujosDeHoja(ruta); if (dib) campos.dib = dib; }
      if (V >= 3) Object.values(relsDeParte(ruta)).filter((q) => /\/comments$/.test(q.tipo) && q.ruta).forEach((q) => {
        const xc = z.texto(q.ruta);
        if (xc) todos(P(xc), "comment").forEach((cm) => { const ref = at(cm, "ref"); if (ref && F().parseRef(ref)) campos["nt:" + ref] = true; });
      });
      if (V >= 3) { estilosDeTabla(nombre, campos); estilosDinamicas(ruta, campos); const cfx = formatoCondicional(doc, nombre); if (cfx) campos.cfx = cfx; }
      if (todos(doc, "drawing").length) aviso = aviso || "Las gráficas e imágenes del archivo no se ven aquí, pero se conservan: al descargarlo en Excel salen igual";
      campos.filas = Math.max(30, maxF + 5);
      campos.cols = Math.max(10, Math.min(MAX_C, maxC + 2));
      res.push({ nombre, campos });
    }
    if (!res.length) throw new Error("El libro no tiene hojas que se puedan leer");
    return { hojas: res, aviso };
  }

  // ------------------------------------------------------------ PowerPoint (.pptx)
  const ANCHO = 960, ALTO = 540;
  const color = (c, def) => (/^#[0-9a-f]{6}$/i.test(c || "") ? c.slice(1).toUpperCase() : def);
  const luzHex = (c) => { const n = parseInt(String(c).replace("#", ""), 16) || 0; return (0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255; };

  // ------------------------------------------------------------ 8.4 trazos SVG como formas libres de PowerPoint
  /** Arco SVG (de x1,y1 a x2,y2) → curvas de Bézier cúbicas (cada una de 90° como máximo). */
  function arcoABezier(x1, y1, rx, ry, phi, fa, fs, x2, y2) {
    if (!rx || !ry || (x1 === x2 && y1 === y2)) return [[x1, y1, x2, y2, x2, y2]];
    rx = Math.abs(rx); ry = Math.abs(ry);
    const f = (phi * Math.PI) / 180, cf = Math.cos(f), sf = Math.sin(f);
    const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2, x1p = cf * dx + sf * dy, y1p = -sf * dx + cf * dy;
    const lam = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
    if (lam > 1) { const m = Math.sqrt(lam); rx *= m; ry *= m; }
    const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p, den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
    let co = Math.sqrt(Math.max(0, num / (den || 1)));
    if (+fa === +fs) co = -co;
    const cxp = (co * rx * y1p) / ry, cyp = (-co * ry * x1p) / rx;
    const cx = cf * cxp - sf * cyp + (x1 + x2) / 2, cy = sf * cxp + cf * cyp + (y1 + y2) / 2;
    const ang = (ux, uy, vx, vy) => Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    const t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
    let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
    if (!+fs && dt > 0) dt -= 2 * Math.PI; else if (+fs && dt < 0) dt += 2 * Math.PI;
    const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9)), paso = dt / n, kk = (4 / 3) * Math.tan(paso / 4), out = [];
    const pt = (ux, uy) => [cx + rx * ux * cf - ry * uy * sf, cy + rx * ux * sf + ry * uy * cf];
    for (let j = 0; j < n; j++) {
      const a1 = t1 + j * paso, a2 = a1 + paso, c1 = Math.cos(a1), s1 = Math.sin(a1), c2 = Math.cos(a2), s2 = Math.sin(a2);
      const q1 = pt(c1 - kk * s1, s1 + kk * c1), q2 = pt(c2 + kk * s2, s2 - kk * c2), e = pt(c2, s2);
      out.push([q1[0], q1[1], q2[0], q2[1], e[0], e[1]]);
    }
    return out;
  }
  /** Trazo SVG (M, L, H, V, C, Q, A, Z absolutos, como los que hace el lienzo) → camino de DrawingML. */
  function caminoDrawingML(d, k) {
    const toks = String(d).match(/[MLHVCQAZ]|-?\d*\.?\d+(?:e-?\d+)?/gi) || [];
    let i = 0, cmd = null, x0 = 0, y0 = 0, px = 0, py = 0, r = "";
    const num = () => +toks[i++];
    const P = (a, b) => `<a:pt x="${Math.round(a * k)}" y="${Math.round(b * k)}"/>`;
    while (i < toks.length) {
      if (/^[MLHVCQAZ]$/i.test(toks[i])) cmd = toks[i++].toUpperCase();
      if (cmd === "M") { px = num(); py = num(); x0 = px; y0 = py; r += `<a:moveTo>${P(px, py)}</a:moveTo>`; cmd = "L"; }
      else if (cmd === "L") { px = num(); py = num(); r += `<a:lnTo>${P(px, py)}</a:lnTo>`; }
      else if (cmd === "H") { px = num(); r += `<a:lnTo>${P(px, py)}</a:lnTo>`; }
      else if (cmd === "V") { py = num(); r += `<a:lnTo>${P(px, py)}</a:lnTo>`; }
      else if (cmd === "C") { const a = [num(), num(), num(), num(), num(), num()]; r += `<a:cubicBezTo>${P(a[0], a[1])}${P(a[2], a[3])}${P(a[4], a[5])}</a:cubicBezTo>`; px = a[4]; py = a[5]; }
      else if (cmd === "Q") { const a = [num(), num(), num(), num()]; r += `<a:quadBezTo>${P(a[0], a[1])}${P(a[2], a[3])}</a:quadBezTo>`; px = a[2]; py = a[3]; }
      else if (cmd === "A") { const a = [num(), num(), num(), num(), num(), num(), num()]; arcoABezier(px, py, a[0], a[1], a[2], a[3], a[4], a[5], a[6]).forEach((c) => (r += `<a:cubicBezTo>${P(c[0], c[1])}${P(c[2], c[3])}${P(c[4], c[5])}</a:cubicBezTo>`)); px = a[5]; py = a[6]; }
      else if (cmd === "Z") { r += "<a:close/>"; px = x0; py = y0; cmd = null; }
      else i++;
    }
    return r;
  }
  /** Forma libre de PowerPoint con las piezas de una figura (relleno y contorno por pieza). */
  function custGeom(partes, w, h) {
    const W = Math.max(1, Math.round(w * EMU)), H = Math.max(1, Math.round(h * EMU));
    return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${partes.map((pz) => `<a:path w="${W}" h="${H}"${pz.relleno ? "" : ' fill="none"'}${pz.trazo ? "" : ' stroke="0"'}>${caminoDrawingML(pz.d, EMU)}</a:path>`).join("")}</a:pathLst></a:custGeom>`;
  }
  const LIBRES = new Set(["clase", "entidad", "entidad-debil", "relacion-debil", "atributo-multi", "actor", "paquete", "componente", "nodo", "final"]);

  // ------------------------------------------------------------ Presentación de PowerPoint: secciones, personalizadas y configuración
  /**
   * Las secciones (Inicio → Sección) en el orden de las diapositivas de ahora, como las guarda PowerPoint: cada diapositiva
   * en una sola y las de una sección juntas. Una diapositiva que quedó sin sección, o separada de las demás de la suya,
   * va con la sección de la diapositiva de antes; una sección que se quedó sin diapositivas se conserva vacía en su lugar.
   * null si no hay secciones.
   */
  function seccionesPptx(doc) {
    const ss = doc.valor("_doc", "secciones");
    if (!Array.isArray(ss) || !ss.length) return null;
    const pags = doc.paginas(), de = new Map();
    ss.forEach((sc, k) => (Array.isArray(sc.d) ? sc.d : []).forEach((pg) => { if (!de.has(pg)) de.set(pg, k); }));
    const sal = ss.map((sc) => ({ id: sc.id, nombre: String(sc.nombre == null ? "" : sc.nombre), d: [] }));
    const orden = [], cerradas = new Set();
    let k = null;
    for (const pg of pags) {
      let kp = de.has(pg) ? de.get(pg) : null;
      if (kp == null || (cerradas.has(kp) && kp !== k)) kp = k == null ? (kp == null ? 0 : kp) : k;
      if (kp !== k) { if (k != null) cerradas.add(k); if (!orden.includes(kp)) orden.push(kp); k = kp; }
      sal[kp].d.push(pg);
    }
    ss.forEach((_, i) => { if (!orden.includes(i)) { const j = orden.indexOf(i - 1); orden.splice(j < 0 ? orden.length : j + 1, 0, i); } });
    return orden.map((i) => sal[i]);
  }
  /** Las presentaciones personalizadas con las diapositivas que siguen en el documento (y un id numérico para cada una). */
  function personalizadasPptx(doc) {
    const l = doc.valor("_doc", "personalizadas");
    if (!Array.isArray(l) || !l.length) return [];
    const hay = new Set(doc.paginas()), usados = new Set();
    let sig = 0;
    l.forEach((c) => { if (/^\d+$/.test(String(c.id))) sig = Math.max(sig, +c.id + 1); });
    return l.map((c) => {
      let id = /^\d+$/.test(String(c.id)) && !usados.has(String(c.id)) ? String(c.id) : String(sig++);
      usados.add(id);
      return { id, idNom: c.id, nombre: String(c.nombre == null ? "" : c.nombre), d: (Array.isArray(c.d) ? c.d : []).filter((pg) => hay.has(pg)) };
    });
  }
  const guid = () => {
    const u = new Uint8Array(16); crypto.getRandomValues(u);
    const h = [...u].map((b) => b.toString(16).padStart(2, "0")).join("").toUpperCase();
    return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
  };
  const URI_SECCIONES = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
  /** <p:custShowLst> (ridDe(página) = r:id de la diapositiva) o "" si no hay. */
  function custShowXml(pers, ridDe) {
    if (!pers.length) return "";
    return `<p:custShowLst>${pers.map((c) => `<p:custShow name="${x(c.nombre)}" id="${c.id}"><p:sldLst>${c.d.map((pg) => ridDe(pg)).filter(Boolean).map((r) => `<p:sld r:id="${r}"/>`).join("")}</p:sldLst></p:custShow>`).join("")}</p:custShowLst>`;
  }
  /** El <p:ext> de las secciones (numDe(página) = id de la diapositiva en sldIdLst) o "" si no hay. */
  function seccionesExtXml(secs, numDe) {
    if (!secs || !secs.length) return "";
    return `<p:ext uri="${URI_SECCIONES}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${secs.map((sc) => `<p14:section name="${x(sc.nombre)}" id="${/^\{[0-9A-Fa-f-]{36}\}$/.test(String(sc.id)) ? sc.id : guid()}"><p14:sldIdLst>${sc.d.map((pg) => numDe(pg)).filter(Boolean).map((n) => `<p14:sldId id="${n}"/>`).join("")}</p14:sldIdLst></p14:section>`).join("")}</p14:sectionLst></p:ext>`;
  }
  /** Configurar presentación (presProps.xml → <p:showPr>): repetir hasta Esc, usar intervalos, animaciones, tipo y qué
      diapositivas. null si es la de siempre. */
  function showPrXml(doc, pers) {
    const c = doc.valor("_doc", "presentacion");
    if (!c || typeof c !== "object" || !Object.keys(c).length) return null;
    const p2 = c.personalizada != null ? pers.find((q) => String(q.idNom) === String(c.personalizada)) : null;
    const rg = Array.isArray(c.rango) && c.rango.length === 2 ? c.rango : null;
    return `<p:showPr${c.repetir ? ' loop="1"' : ""} showNarration="1"${c.sinAnimacion ? ' showAnimation="0"' : ""}${c.sinIntervalos ? ' useTimings="0"' : ""}>`
      + (c.tipo === "kiosco" ? '<p:kiosk/>' : c.tipo === "ventana" ? '<p:browse/>' : '<p:present/>')
      + (p2 ? `<p:custShow id="${p2.id}"/>` : rg ? `<p:sldRg st="${Math.max(1, +rg[0] || 1)}" end="${Math.max(1, +rg[1] || 1)}"/>` : '<p:sldAll/>')
      + '</p:showPr>';
  }

  function pptx(doc, esDiagrama) {
    const og = !esDiagrama && NOM.ooxml.usarOrigen !== false && NOM.origen ? NOM.origen.leer(doc) : null;
    if (og && og.fmt === "pptx" && og.h && og.h.pp) {
      try { return pptxArmar(doc, false, og); }
      catch (e) { if (typeof console !== "undefined") console.warn("PowerPoint: no se pudo partir del archivo original; se arma desde cero.", e); }
    }
    return pptxArmar(doc, esDiagrama, null);
  }
  /** De cada diapositiva importada: sus bloques en orden y de qué elemento del original salió cada uno. */
  function huellasPptx(doc) {
    const pp = {};
    doc.paginas().forEach((pg) => {
      const orden = doc.deLaPagina(pg).slice().sort((a, b) => (a.z || 0) - (b.z || 0)), e = {};
      orden.forEach((b) => { if (b.ox != null) (e[b.ox] = e[b.ox] || []).push(b.id); });
      pp[pg] = { ids: orden.map((b) => b.id), e };
    });
    return pp;
  }
  const P_GEO = ["x", "y", "w", "h", "rot", "volteoH", "volteoV", "z", "orden"], P_TXT = ["html", "texto"];
  /** Tres huellas de un bloque de diapositiva: sin su posición, sin posición ni texto, y solo la posición. */
  function parcialesPptx(b) {
    const a = {}, c = {}, g = {};
    for (const k in b) {
      if (k === "id" || b[k] == null) continue;
      if (P_GEO.includes(k)) { if (k !== "z" && k !== "orden") g[k] = b[k]; continue; }
      a[k] = b[k];
      if (!P_TXT.includes(k)) c[k] = b[k];
    }
    const r = [NOM.origen.huella(a), NOM.origen.huella(c), NOM.origen.huella(g)];
    // El formato de texto de todo el cuadro que ya traía al importar (solo si traía alguno): con eso se
    // sabe después cuál cambió la persona.
    const f = {};
    P_FMT.forEach((k) => { if (b[k] != null) f[k] = b[k]; });
    if (Object.keys(f).length) r.push(f);
    return r;
  }
  /** Formato de texto que vale para todo el cuadro (lienzo.js). */
  const P_FMT = ["tNegrita", "tCursiva", "tSubrayado", "tColor", "tFuente", "tTam", "tAlin"];
  /**
   * DEF-45: ¿lo único que cambió en un cuadro importado (además, tal vez, de su texto o su lugar) fue el
   * formato de texto de todo el cuadro? Devuelve { fmt: lo que la persona puso, texto, geo } o null.
   * q son las huellas de cuando se importó; q[3], si está, el formato que ya traía. En lo importado con
   * una versión anterior no viene: se prueba con «no traía ninguno» y con «traía ese mismo color».
   */
  function formatoCambiadoPptx(b, q) {
    const intentos = q[3] && typeof q[3] === "object" ? [q[3]] : [{}].concat(b.tColor != null ? [{ tColor: b.tColor }] : []);
    for (const f0 of intentos) {
      const b0 = Object.assign({}, b);
      P_FMT.forEach((k) => { if (f0[k] != null) b0[k] = f0[k]; else delete b0[k]; });
      const pc = parcialesPptx(b0);
      if (pc[1] !== q[1]) continue;
      const fmt = {};
      let quitado = false;
      P_FMT.forEach((k) => { const v = b[k] == null ? null : b[k], v0 = f0[k] == null ? null : f0[k]; if (v !== v0) { if (v == null) quitado = true; else fmt[k] = v; } });
      // quitar un formato que el archivo traía no se puede escribir como parche: ese cuadro se rehace
      if (quitado || !Object.keys(fmt).length) return null;
      return { fmt, texto: pc[0] !== q[0], geo: pc[2] !== q[2] };
    }
    return null;
  }
  /** Pone (o cambia) un atributo en el texto de atributos de una etiqueta. */
  const ponAttr = (ats, n, v) => { const re = new RegExp("\\s" + n + '="[^"]*"'); return re.test(ats) ? ats.replace(re, ` ${n}="${v}"`) : ats + ` ${n}="${v}"`; };
  const A_RELLENOS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"], A_TRAS_LATIN = ["ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"];
  /**
   * DEF-45: el formato de todo el cuadro escrito sobre el p:txBody ORIGINAL de una figura: solo se tocan
   * los atributos de a:rPr (y algn de a:pPr) que la persona cambió; idioma, revisión ortográfica,
   * a:endParaRPr, viñetas y todo lo demás quedan como en el archivo. sz ya viene en centésimas de punto.
   */
  function conFormatoPptx(t, fmt, sz) {
    const O = NOM.origen, a = t.indexOf("<p:txBody>"), z = t.lastIndexOf("</p:txBody>");
    if (a < 0 || z < a) return null;
    const rPr = (rp) => {
      let ats = "", hijos = [];
      if (rp) {
        const g = O.finEtiqueta(rp, 0), vacio = rp.charCodeAt(g - 1) === 47;
        ats = rp.slice("<a:rPr".length, vacio ? g - 1 : g).replace(/\s+$/, "");
        if (!vacio) { const fin = rp.lastIndexOf("</a:rPr>"); hijos = O.hijosXml(rp, g + 1, fin).map((h) => ({ n: O.local(h.n), x: rp.slice(h.i, h.f) })); }
      }
      if (fmt.tNegrita != null) ats = ponAttr(ats, "b", fmt.tNegrita ? 1 : 0);
      if (fmt.tCursiva != null) ats = ponAttr(ats, "i", fmt.tCursiva ? 1 : 0);
      if (fmt.tSubrayado != null) ats = ponAttr(ats, "u", fmt.tSubrayado ? "sng" : "none");
      if (sz) ats = ponAttr(ats, "sz", sz);
      if (fmt.tColor) {
        hijos = hijos.filter((h) => !A_RELLENOS.includes(h.n));
        const k = hijos.findIndex((h) => h.n === "ln");
        hijos.splice(k + 1, 0, { n: "solidFill", x: `<a:solidFill><a:srgbClr val="${color(fmt.tColor, "000000")}"/></a:solidFill>` });
      }
      if (fmt.tFuente) {
        hijos = hijos.filter((h) => h.n !== "latin");
        const k = hijos.findIndex((h) => A_TRAS_LATIN.includes(h.n));
        hijos.splice(k < 0 ? hijos.length : k, 0, { n: "latin", x: `<a:latin typeface="${x(fmt.tFuente)}"/>` });
      }
      return hijos.length ? `<a:rPr${ats}>${hijos.map((h) => h.x).join("")}</a:rPr>` : `<a:rPr${ats}/>`;
    };
    let cuerpo = t.slice(a, z);
    if (fmt.tNegrita != null || fmt.tCursiva != null || fmt.tSubrayado != null || sz || fmt.tColor || fmt.tFuente)
      cuerpo = cuerpo.replace(/(<a:(?:r|fld)(?:\s[^>]*)?>\s*)(<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>)?/g, (m, abre, rp) => abre + rPr(rp || ""));
    const al = fmt.tAlin && { izq: "l", centro: "ctr", der: "r", just: "just" }[fmt.tAlin];
    if (al) cuerpo = cuerpo.replace(/(<a:p>\s*)(<a:pPr\b[^>]*>)?/g, (m, abre, pp) => {
      if (!pp) return abre + `<a:pPr algn="${al}"/>`;
      const vacio = pp.charCodeAt(pp.length - 2) === 47, ats = pp.slice("<a:pPr".length, pp.length - (vacio ? 2 : 1));
      return abre + "<a:pPr" + ponAttr(ats.replace(/\s+$/, ""), "algn", al) + (vacio ? "/>" : ">");
    });
    return t.slice(0, a) + cuerpo + t.slice(z);
  }
  const TM_LISTAS = new Set(["childTnLst", "tnLst", "subTnLst", "bldLst", "stCondLst", "endCondLst", "nextCondLst", "prevCondLst", "tmplLst"]);
  /**
   * DEF-45: las animaciones (p:timing) de una diapositiva editada. Se conservan; solo se quitan los
   * efectos que apuntan a una figura que ya no está (si se quedaran, PowerPoint pide reparar el
   * archivo). vivo(spid, ultimoParrafo|null) dice si el destino sigue existiendo. Al quitar un efecto
   * se quita también lo que quede vacío hacia arriba (un grupo «al hacer clic» sin efectos, etc.).
   * Devuelve el p:timing nuevo ("" si no queda nada).
   */
  function podarTiming(xt, vivo) {
    const O = NOM.origen;
    const podar = (h) => {
      const loc = O.local(h.n);
      if (loc === "spTgt" || loc === "inkTgt" || /^bld(P|Dgm|OleChart|Graphic)$/.test(loc)) {
        const spid = O.attr(xt, h, "spid");
        if (spid == null) return xt.slice(h.i, h.f);
        const pr = loc === "spTgt" ? /<p:pRg\b[^>]*\bend="(\d+)"/.exec(xt.slice(h.i, h.f)) : null;
        return vivo(spid, pr ? +pr[1] : null) ? xt.slice(h.i, h.f) : null;
      }
      if (h.ci === h.cf) return xt.slice(h.i, h.f);
      const hs = O.hijosXml(xt, h.ci, h.cf);
      if (!hs.length) return xt.slice(h.i, h.f);
      let dentro = "", pos = h.ci, quitados = 0, quedan = 0, queda = {};
      for (const c of hs) {
        const r = podar(c);
        dentro += xt.slice(pos, c.i); pos = c.f;
        if (r == null) quitados++; else { dentro += r; quedan++; queda[O.local(c.n)] = true; }
      }
      dentro += xt.slice(pos, h.cf);
      if (!quitados) return xt.slice(h.i, h.f);
      // de una lista se puede quitar un elemento; de lo demás no (le faltaría una parte obligatoria)
      if (loc === "timing" ? !queda.tnLst : TM_LISTAS.has(loc) ? !quedan : true) return null;
      return xt.slice(h.i, h.ci) + dentro + xt.slice(h.cf, h.f);
    };
    const raiz = O.hijosXml(xt)[0];
    if (!raiz) return xt;
    const r = podar(raiz);
    return r == null ? "" : r;
  }
  /**
   * El p:timing de una página con la lista de animaciones de NOM (las de PowerPoint con su árbol; las propias de NOM con su
   * efecto de PowerPoint). spidDe(bloque) → número de la figura en el archivo; conTexto(bloque) → si lleva texto.
   */
  function timingDe(doc, p, spidDe, conTexto, sinSonidos) {
    const A = NOM.animOffice;
    if (!A) return "";
    const lista = doc.valor("_doc", "anim:" + p);
    if (!Array.isArray(lista) || !lista.length) return "";
    try {
      return A.timingXml(lista.filter((x) => x && (x.id || x.spid)), (bid) => (bid ? spidDe(bid) : null),
        { preset: (x) => A.presetDe(x), medios: doc.valor("_doc", "animMedios:" + p) || [], bld: doc.valor("_doc", "animBld:" + p) || [], conTexto, sinSonidos });
    } catch (e) { return ""; }
  }
  function pptxArmar(doc, esDiagrama, og) {
    let PRE = "", heredar = false;
    const medios = [], graficas = [], notas = [];
    const TD = esDiagrama || !NOM.diapo ? { w: ANCHO, h: ALTO } : NOM.diapo.tam(doc);
    const te = !esDiagrama && NOM.diapo ? NOM.diapo.tema(doc) : null;
    const paginas = doc.paginas().filter((p) => doc.deLaPagina(p).length || !esDiagrama);
    /** Un color CSS (#hex o rgb()) en hex para Office. */
    const hexCss = (c, def) => {
      if (!c) return def;
      const m = /^#([0-9a-f]{6})$/i.exec(c.trim());
      if (m) return m[1].toUpperCase();
      const r = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(c.trim());
      return r ? [r[1], r[2], r[3]].map((v) => (+v).toString(16).padStart(2, "0")).join("").toUpperCase() : def;
    };
    const armar = (p, n, O) => {
      // Lo oculto (capas) no se exporta, ni lo que quedó fuera de la diapositiva, en la mesa de trabajo.
      let bs = doc.deLaPagina(p).filter((b) => (b.x != null || b.tipo === "conector") && !b.oculto
        && (esDiagrama || b.tipo === "conector" || !(b.x + b.w <= 0 || b.y + b.h <= 0 || b.x >= TD.w || b.y >= TD.h)));
      let k = O ? O.k : 1, dx = 0, dy = 0;
      const FX = O ? O.fx : 1, FY = O ? O.fy : 1;
      const conPos = bs.filter((b) => b.x != null);
      if (esDiagrama && conPos.length) {
        // El pizarrón no tiene bordes: se encuadra todo dentro de la diapositiva con 30 px de margen.
        const minX = Math.min(...conPos.map((b) => b.x)), minY = Math.min(...conPos.map((b) => b.y));
        const maxX = Math.max(...conPos.map((b) => b.x + b.w)), maxY = Math.max(...conPos.map((b) => b.y + b.h));
        k = Math.min(1, (TD.w - 60) / (maxX - minX || 1), (TD.h - 60) / (maxY - minY || 1));
        dx = (TD.w - (maxX - minX) * k) / 2 - minX * k; dy = (TD.h - (maxY - minY) * k) / 2 - minY * k;
      }
      bs = bs.slice().sort((a, b) => (a.z || 0) - (b.z || 0));
      const E = (v) => Math.round(v * EMU);
      const pos = (b) => `<a:xfrm${b.rot ? ` rot="${Math.round(b.rot * 60000)}"` : ""}${b.volteoH ? ' flipH="1"' : ""}${b.volteoV ? ' flipV="1"' : ""}><a:off x="${E((b.x * k + dx) * FX)}" y="${E((b.y * k + dy) * FY)}"/><a:ext cx="${E(Math.max(1, b.w) * k * FX)}" cy="${E(Math.max(1, b.h) * k * FY)}"/></a:xfrm>`;
      const alfa = (b) => (b.opacidad != null && b.opacidad < 1 ? `<a:alpha val="${Math.round(b.opacidad * 100000)}"/>` : "");
      const linea = (b, e) => !e.trazo ? "<a:ln><a:noFill/></a:ln>"
        : `<a:ln w="${Math.round(e.ancho * EMU * k)}"><a:solidFill><a:srgbClr val="${color(e.trazo, "2B2A33")}">${alfa(b)}</a:srgbClr></a:solidFill>${e.estilo === "guiones" ? '<a:prstDash val="dash"/>' : e.estilo === "puntos" ? '<a:prstDash val="sysDot"/>' : ""}${b.figura === "linea-flecha" ? '<a:tailEnd type="triangle"/>' : ""}</a:ln>`;
      const sombra = (b) => (b.sombra === true || (b.sombra == null && b.tipo === "forma" && !NOM.figuras.esLinea(b.figura))
        ? '<a:effectLst><a:outerShdw blurRad="63500" dist="38100" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="22000"/></a:srgbClr></a:outerShdw></a:effectLst>' : "");
      const GEO = { rect: "roundRect", elipse: "ellipse", rombo: "diamond", triangulo: "triangle", "triangulo-r": "rtTriangle", pentagono: "pentagon", hexagono: "hexagon", octagono: "octagon",
        estrella: "star5", corazon: "heart", nube: "cloud", burbuja: "wedgeRoundRectCallout", flecha: "rightArrow", "flecha-doble": "leftRightArrow", chevron: "chevron", cruz: "plus",
        paralelogramo: "flowChartInputOutput", trapecio: "trapezoid", cilindro: "can", documento: "flowChartDocument", nota: "foldedCorner", linea: "line", "linea-flecha": "line" };
      const rels = [{ id: "rId1", tipo: T.diseno, destino: "../slideLayouts/slideLayout1.xml" }];
      let id = O ? O.idBase : 2;
      const pre = PRE ? PRE.toLowerCase() + "-" : "";
      const imagenRel = (src) => {
        const im = imagen(src);
        if (!im) return null;
        const m = medios.length + 1, rid = "rId" + PRE + "Img" + m;
        medios.push({ nombre: `ppt/media/${pre}imagen${m}.${im.ext}`, datos: im.bytes });
        rels.push({ id: rid, tipo: T.img, destino: `../media/${pre}imagen${m}.${im.ext}` });
        return rid;
      };
      /** Texto simple (un solo formato para todo el cuadro). */
      const parrafos = (html, tam, negrita, centro, opcT) => {
        opcT = opcT || {};
        const ps = String(html || "").split(/<br\s*\/?>|<\/p>|<\/div>/i);
        const algn = opcT.alin ? { izq: "l", centro: "ctr", der: "r" }[opcT.alin] : centro ? "ctr" : null;
        return ps.map((h) => `<a:p>${algn ? `<a:pPr algn="${algn}"/>` : ""}${tramos(h).filter((t) => !t.br).map((t) =>
          `<a:r><a:rPr lang="es-MX" sz="${Math.round(tam * k * 75)}" b="${t.b || negrita ? 1 : 0}" i="${t.i || opcT.i ? 1 : 0}"${t.u || opcT.u ? ' u="sng"' : ""} dirty="0"><a:solidFill><a:srgbClr val="${opcT.color || "1F1E26"}"/></a:solidFill><a:latin typeface="${x(opcT.fuente || "Calibri")}"/></a:rPr><a:t>${x(t.t)}</a:t></a:r>`).join("")}</a:p>`).join("");
      };
      /** 8.3 Texto con formato por partes, viñetas y numeración. */
      const parrafosRicos = (html, tam, negrita, opcT) => {
        if (!NOM.diapo) return parrafos(html, tam, negrita, false, opcT);
        const ps = NOM.diapo.parrafos(html);
        if (!ps.length) return "<a:p/>";
        const algnBase = opcT.alin ? { izq: "l", centro: "ctr", der: "r" }[opcT.alin] : null;
        return ps.map((pp) => {
          // FH: lo que la persona le puso a todo el cuadro de una figura del archivo original (DEF-45)
          const FH = heredar || {};
          const algn = pp.al ? { izq: "l", centro: "ctr", der: "r", just: "just" }[pp.al] : heredar ? (FH.tAlin ? { izq: "l", centro: "ctr", der: "r", just: "just" }[FH.tAlin] : null) : algnBase;
          const sangria = pp.lista ? ` marL="${342900 + pp.nivel * 342900}" indent="-342900"` : "";
          const vineta = pp.lista === "v" ? '<a:buFont typeface="Arial"/><a:buChar char="•"/>' : pp.lista === "n" ? '<a:buAutoNum type="arabicPeriod"/>' : "<a:buNone/>";
          // "heredar": el texto va dentro de una figura del archivo original; solo se escribe el formato
          // que la persona puso a mano y lo demás (tamaño, color, fuente, viñeta) lo sigue dando el diseño.
          const ppr = heredar && pp.lista === "v" ? `<a:pPr${pp.nivel ? ` lvl="${Math.min(8, pp.nivel)}"` : ""}${algn ? ` algn="${algn}"` : ""}/>` : `<a:pPr${algn ? ` algn="${algn}"` : ""}${sangria}>${vineta}</a:pPr>`;
          const runs = pp.runs.filter((r) => r.t).map((r) => {
            const t2 = r.tam ? (/px$/.test(r.tam) ? parseFloat(r.tam) / 1.333 : /pt$/.test(r.tam) ? parseFloat(r.tam) : tam) : tam;
            if (heredar) {
              const szH = r.tam ? Math.round(t2 * k * 100) : FH.tTam ? Math.round((FH.tTam / 1.333) * k * 100) : 0;
              const bH = r.b ? ' b="1"' : FH.tNegrita != null ? ` b="${FH.tNegrita ? 1 : 0}"` : "", colH = r.color ? hexCss(r.color, "000000") : FH.tColor ? color(FH.tColor, "000000") : null;
              return `<a:r><a:rPr lang="es-MX"${szH ? ` sz="${szH}"` : ""}${bH}${r.i || FH.tCursiva ? ' i="1"' : ""}${r.u || FH.tSubrayado ? ' u="sng"' : ""}${r.s ? ' strike="sngStrike"' : ""} dirty="0">${colH ? `<a:solidFill><a:srgbClr val="${colH}"/></a:solidFill>` : ""}${FH.tFuente ? `<a:latin typeface="${x(FH.tFuente)}"/>` : ""}</a:rPr><a:t>${x(r.t)}</a:t></a:r>`;
            }
            return `<a:r><a:rPr lang="es-MX" sz="${Math.round(t2 * k * 100)}" b="${r.b || negrita ? 1 : 0}" i="${r.i || opcT.i ? 1 : 0}"${r.u || opcT.u ? ' u="sng"' : ""}${r.s ? ' strike="sngStrike"' : ""} dirty="0"><a:solidFill><a:srgbClr val="${hexCss(r.color, opcT.color || "1F1E26")}"/></a:solidFill><a:latin typeface="${x(opcT.fuente || "Calibri")}"/></a:rPr><a:t>${x(r.t)}</a:t></a:r>`;
          }).join("");
          return `<a:p>${ppr}${runs}</a:p>`;
        }).join("");
      };
      const colorTema = (b) => (te ? (b.estilo === "h1" || b.estilo === "h2" ? te.titulo : te.texto).slice(1) : null);
      const textoDe = (b, fondo) => ({ color: color(b.tColor, colorTema(b) || (fondo && luzHex(fondo) < 0.5 ? "FFFFFF" : "1F1E26")), fuente: b.tFuente || (te && (b.estilo === "h1" || b.estilo === "h2") ? te.fuenteT : te && te.fuente) || null, alin: b.tAlin, i: b.tCursiva, u: b.tSubrayado });
      const porId = Object.fromEntries(conPos.map((b) => [b.id, b]));
      const caja = (nid, nombre, b, cuerpo, descr) => `<p:sp><p:nvSpPr><p:cNvPr id="${nid}" name="${nombre} ${nid}"${descr ? ` descr="${x(descr)}"` : ""}/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${pos(b)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>${cuerpo}</p:txBody></p:sp>`;
      /** 8.4 Clase y entidad: nombre arriba y renglones fijos que caen entre las líneas divisorias. */
      const compartimentosPptx = (b, e, opcT) => {
        const ls = String(b.texto || "").split("\n").map((l) => l.replace(/\s+$/, ""));
        const nombre = ls.shift() || "", resto = ls.filter((l) => l.trim() && !/^-{2,}$/.test(l.trim()));
        const secs = e.f === "clase" ? [resto.filter((l) => !/\(.*\)/.test(l)), resto.filter((l) => /\(.*\)/.test(l))] : [resto];
        const fs = +b.tTam || 14, L = fs * 1.35, pts = (px) => Math.max(0, Math.round(px * 0.75 * k * 100));
        const run = (t, neg, sub) => `<a:r><a:rPr lang="es-MX" sz="${Math.round(fs * 0.75 * k * 100)}" b="${neg ? 1 : 0}"${sub ? ' u="sng"' : ""} dirty="0"><a:solidFill><a:srgbClr val="${(opcT && opcT.color) || "1F1E26"}"/></a:solidFill><a:latin typeface="${x((opcT && opcT.fuente) || "Calibri")}"/></a:rPr><a:t>${x(t)}</a:t></a:r>`;
        const par = (contenido, antes, alin) => `<a:p><a:pPr algn="${alin}" marL="0" indent="0"><a:lnSpc><a:spcPts val="${pts(L)}"/></a:lnSpc><a:spcBef><a:spcPts val="${pts(antes)}"/></a:spcBef><a:buNone/></a:pPr>${contenido}</a:p>`;
        let ps = par(run(nombre, true), 0, "ctr"), cursor = 6 + L, y = 12 + L;
        const lineas = [];
        secs.forEach((sec) => {
          lineas.push(y);
          sec.forEach((l, i) => {
            const arriba = y + 5 + i * L, clave = e.f === "entidad" && /^(PK\b|#)/i.test(l.trim());
            if (arriba + L > b.h) return;
            ps += par(run(clave ? l.trim().replace(/^(PK\s*:?\s*|#\s*)/i, "") : l, false, clave), arriba - cursor, "l");
            cursor = arriba + L;
          });
          y += 5 + sec.length * L + (sec.length ? 5 : 14);
        });
        return { parrafos: ps, lineas: lineas.filter((v) => v < b.h - 2) };
      };
      /** El vínculo o la acción al hacer clic de un bloque (a:hlinkClick, con su relación), como lo escribe PowerPoint. */
      const accionXml = (b) => {
        const a = b.accion;
        if (!a || esDiagrama) return "";
        const J = { sig: "nextslide", ant: "previousslide", primera: "firstslide", ultima: "lastslide", fin: "endshow", volver: "lastslideviewed" };
        if (J[a.tipo]) return `<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=${J[a.tipo]}"/>`;
        const rid = "rIdAcc" + (rels.length + 1);
        if (a.tipo === "url" && /^(https?:|mailto:)/i.test(a.url || "")) { rels.push({ id: rid, tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", destino: a.url, externo: true }); return `<a:hlinkClick r:id="${rid}"/>`; }
        if (a.tipo === "diapo") {
          let dest = null;
          if (!O) { const k = paginas.indexOf(a.pagina); if (k >= 0) dest = `slide${k + 1}.xml`; }
          else { const os = doc.valor("_doc", "os:" + a.pagina); if (os) dest = os.slice(os.lastIndexOf("/") + 1); }
          if (!dest) return "";
          rels.push({ id: rid, tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", destino: dest });
          return `<a:hlinkClick r:id="${rid}" action="ppaction://hlinksldjump"/>`;
        }
        return "";
      };
      const figurasL = bs.map((b) => {
        if (O && O.genera && !O.genera(b)) return "";
        heredar = (O && O.hereda && O.hereda(b)) || false;
        let f;
        try { f = figura1(b); } finally { heredar = false; }
        const hx = f ? accionXml(b) : "";
        return hx ? f.replace(/<p:cNvPr\b([^>]*?)\/>/, `<p:cNvPr$1>${hx}</p:cNvPr>`) : f;
      });
      function figura1(b) {
        const nid = id++;
        if (b.tipo === "texto") {
          const tam = b.tTam ? b.tTam / 1.333 : b.estilo === "h1" ? 40 : b.estilo === "h2" ? 28 : 20;
          const neg = b.tNegrita != null ? b.tNegrita : b.estilo === "h1" || b.estilo === "h2";
          // El nombre del cuadro guarda si era título (al volver a abrirlo en NOM se recupera).
          return caja(nid, b.estilo === "h1" ? "Título" : b.estilo === "h2" ? "Encabezado" : "Texto", b, parrafosRicos(b.html, tam, neg, textoDe(b)));
        }
        if (b.tipo === "ecuacion") return caja(nid, "Ecuación", b, `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="es-MX" sz="${Math.round(((b.tTam || 32) / 1.333) * k * 100)}" dirty="0"/><a:t>${x(NOM.ecuaciones ? NOM.ecuaciones.texto(b.tex || "") : b.tex || "")}</a:t></a:r></a:p>`, "LaTeX: " + (b.tex || ""));
        if (b.tipo === "forma") {
          const e = NOM.figuras.estiloFigura(b);
          const geo = b.prst && /^[a-zA-Z0-9]+$/.test(b.prst) ? b.prst : GEO[e.f] || "roundRect";
          const av = b.prst && b.av ? Object.entries(b.av).filter(([n, v]) => /^[a-z0-9]+$/i.test(n) && isFinite(+v)).map(([n, v]) => `<a:gd name="${n}" fmla="val ${Math.round(+v)}"/>`).join("")
            : geo === "roundRect" ? `<a:gd name="adj" fmla="val ${Math.round(Math.min(50000, (e.radio / Math.max(1, Math.min(b.w, b.h))) * 100000))}"/>` : "";
          if (geo === "line") {
            return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${nid}" name="Línea ${nid}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${pos(Object.assign({}, b, { y: b.y + b.h / 2, h: 0 }))}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${linea(b, e)}</p:spPr></p:cxnSp>`;
          }
          // 8.4 Figuras de UML y entidad-relación: forma libre con el mismo trazo que en el lienzo.
          const cm = NOM.figuras.COMPARTIMENTOS && NOM.figuras.COMPARTIMENTOS.has(e.f);
          let geometria = `<a:prstGeom prst="${geo}"><a:avLst>${av}</a:avLst></a:prstGeom>`, rellenoCol = e.relleno;
          const opcT = textoDe(b, e.relleno), tam = b.tTam ? b.tTam / 1.333 : 16;
          let cuerpo = b.html && NOM.diapo ? parrafosRicos(b.html, tam, b.tNegrita !== false, Object.assign({ alin: e.f === "nota" ? null : b.tAlin || "centro" }, opcT)) : parrafos(x(b.texto || ""), tam, b.tNegrita !== false, e.f !== "nota", opcT);
          let cuerpoPr = `<a:bodyPr anchor="${e.f === "nota" ? "t" : "ctr"}" wrap="square"><a:normAutofit/></a:bodyPr>`;
          if (LIBRES.has(e.f) && NOM.figuras.partesFigura) {
            let partes = NOM.figuras.partesFigura(e, b.w, b.h);
            if (e.f === "final") { rellenoCol = e.trazo || "#2B2A33"; partes = [Object.assign({}, partes[0], { relleno: null }), partes[1]]; }
            if (cm) {
              const cp = compartimentosPptx(b, e, opcT);
              partes = partes.concat(cp.lineas.map((y) => ({ d: `M0,${y} H${b.w}`, relleno: null, trazo: e.trazo })));
              cuerpo = cp.parrafos;
              cuerpoPr = `<a:bodyPr anchor="t" wrap="square" lIns="${E(10 * k)}" tIns="${E(6 * k)}" rIns="${E(10 * k)}" bIns="0"><a:noAutofit/></a:bodyPr>`;
            } else if (e.f === "actor") cuerpoPr = '<a:bodyPr anchor="b" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr>';
            else if (e.f === "paquete") cuerpoPr = `<a:bodyPr anchor="ctr" wrap="square" tIns="${E(Math.min(24, b.h * 0.22) * k)}"><a:normAutofit/></a:bodyPr>`;
            geometria = custGeom(partes, b.w, b.h);
          }
          const relleno = rellenoCol ? `<a:solidFill><a:srgbClr val="${color(rellenoCol, "CFE0FF")}">${alfa(b)}</a:srgbClr></a:solidFill>` : "<a:noFill/>";
          return `<p:sp><p:nvSpPr><p:cNvPr id="${nid}" name="Figura ${nid}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${pos(b)}${geometria}${relleno}${linea(b, e)}${sombra(b)}</p:spPr>
<p:txBody>${cuerpoPr}<a:lstStyle/>${cuerpo}</p:txBody></p:sp>`;
        }
        if (b.tipo === "imagen" || (b.tipo === "icono" && b.png)) {
          const rid = imagenRel(b.tipo === "icono" ? b.png : b.src);
          if (!rid) return "";
          const e = { trazo: b.bordeAncho ? b.borde || "#2B2A33" : null, ancho: +b.bordeAncho || 0, estilo: b.bordeEstilo };
          const r = b.recorte, f = b.filtros || {};
          const recorte = r ? `<a:srcRect l="${Math.round((+r.iz || 0) * 100000)}" t="${Math.round((+r.s || 0) * 100000)}" r="${Math.round((+r.d || 0) * 100000)}" b="${Math.round((+r.i || 0) * 100000)}"/>` : "";
          const efectos = (b.opacidad != null && b.opacidad < 1 ? `<a:alphaModFix amt="${Math.round(b.opacidad * 100000)}"/>` : "") + (+f.gris >= 50 ? "<a:grayscl/>" : "")
            + ((f.brillo != null && +f.brillo !== 100) || (f.contraste != null && +f.contraste !== 100) ? `<a:lum bright="${Math.round(((+f.brillo || 100) - 100) * 1000)}" contrast="${Math.round(((+f.contraste || 100) - 100) * 1000)}"/>` : "");
          const geo = b.mascara ? GEO[b.mascara] || "rect" : b.radio ? "roundRect" : "rect";
          return `<p:pic><p:nvPicPr><p:cNvPr id="${nid}" name="${b.tipo === "icono" ? "Ícono" : "Imagen"} ${nid}" descr="${x(b.alt || b.nombre || "")}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}">${efectos}</a:blip>${recorte}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${pos(b)}<a:prstGeom prst="${geo}"><a:avLst/></a:prstGeom>${linea(b, e)}</p:spPr></p:pic>`;
        }
        if (b.tipo === "grafica" && !b.hoja) {
          const xmlG = graficaPptx(b);
          if (!xmlG) return "";
          graficas.push(xmlG);
          const ng = graficas.length, rid = "rId" + PRE + "Graf" + ng;
          rels.push({ id: rid, tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", destino: `../charts/${pre}chart${ng}.xml` });
          return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${nid}" name="Gráfica ${nid}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${E(b.x * k + dx)}" y="${E(b.y * k + dy)}"/><a:ext cx="${E(b.w * k)}" cy="${E(b.h * k)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
        }
        if (b.tipo === "tabla") {
          const filas = tablaVisible(b);
          if (!filas.length) return "";
          const ancho = E(b.w * k), alto = E(Math.max(1, b.h) * k), cols = Math.max(...filas.map((f) => f.length));
          const cab = b.enDiapo && b.encabezado, cCab = color(b.colorCab, "3D72F2");
          const tam = b.tTam ? b.tTam / 1.333 : b.enDiapo ? 14 : 14;
          return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${nid}" name="Tabla ${nid}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="${E(b.x * k + dx)}" y="${E(b.y * k + dy)}"/><a:ext cx="${ancho}" cy="${alto}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="${cab ? 1 : 0}" bandRow="1"/>
<a:tblGrid>${`<a:gridCol w="${Math.floor(ancho / cols)}"/>`.repeat(cols)}</a:tblGrid>${filas.map((f, fi) => `<a:tr h="${Math.floor(alto / filas.length)}">${Array.from({ length: cols }, (_, ci) => f[ci] || "").map((v) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="es-MX" sz="${Math.round(tam * k * 100)}" b="${cab && fi === 0 ? 1 : 0}" dirty="0">${cab && fi === 0 ? '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' : ""}</a:rPr><a:t>${x(v)}</a:t></a:r></a:p></a:txBody><a:tcPr>${["lnL", "lnR", "lnT", "lnB"].map((l) => `<a:${l} w="9525"><a:solidFill><a:srgbClr val="B7B4C2"/></a:solidFill></a:${l}>`).join("")}${cab && fi === 0 ? `<a:solidFill><a:srgbClr val="${cCab}"/></a:solidFill>` : fi % 2 ? '<a:solidFill><a:srgbClr val="F6F5FA"/></a:solidFill>' : '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'}</a:tcPr></a:tc>`).join("")}</a:tr>`).join("")}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
        }
        if (b.tipo === "conector" && porId[b.desde] && porId[b.hasta]) {
          // 8.4 La flecha con su recorrido (recta, codo o curva), sus puntas y sus textos.
          const r = NOM.geometria.rutaConector(porId[b.desde], porId[b.hasta], b);
          const T2 = (q) => ({ x: q.x * k + dx, y: q.y * k + dy }), pts = r.pts.map(T2);
          const col = color(b.color, "4A4854"), estilo = b.estilo === "guiones" ? '<a:prstDash val="dash"/>' : b.estilo === "puntos" ? '<a:prstDash val="sysDot"/>' : "";
          const tipo = { abierta: "arrow", hueco: "triangle", rombo: "diamond", "rombo-hueco": "diamond", circulo: "oval" }[b.tipoPunta] || "triangle";
          const pIni = b.puntas === "ambas" || b.puntas === "inicio", pFin = !b.puntas || b.puntas === "fin" || b.puntas === "ambas";
          const ln = `<a:ln w="${Math.round((+b.grosor || 2) * k * 9525)}"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill>${estilo}${pIni ? `<a:headEnd type="${tipo}"/>` : ""}${pFin ? `<a:tailEnd type="${tipo}"/>` : ""}</a:ln>`;
          let s2;
          if ((b.ruta || "recta") === "recta") {
            const p1 = pts[0], p2 = pts[pts.length - 1];
            s2 = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${nid}" name="Conector ${nid}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm${p2.x < p1.x ? ' flipH="1"' : ""}${p2.y < p1.y ? ' flipV="1"' : ""}><a:off x="${E(Math.min(p1.x, p2.x))}" y="${E(Math.min(p1.y, p2.y))}"/><a:ext cx="${E(Math.abs(p2.x - p1.x))}" cy="${E(Math.abs(p2.y - p1.y))}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>${ln}</p:spPr></p:cxnSp>`;
          } else {
            // En codo o curva: una forma libre abierta con el mismo recorrido que en el lienzo.
            const x0 = Math.min(...pts.map((q) => q.x)), y0 = Math.min(...pts.map((q) => q.y));
            const w = Math.max(1, Math.max(...pts.map((q) => q.x)) - x0), h = Math.max(1, Math.max(...pts.map((q) => q.y)) - y0);
            const L = (q) => `<a:pt x="${E(q.x - x0)}" y="${E(q.y - y0)}"/>`;
            const camino = b.ruta === "curva" ? `<a:moveTo>${L(pts[0])}</a:moveTo><a:cubicBezTo>${L(pts[1])}${L(pts[2])}${L(pts[3])}</a:cubicBezTo>` : `<a:moveTo>${L(pts[0])}</a:moveTo>${pts.slice(1).map((q) => `<a:lnTo>${L(q)}</a:lnTo>`).join("")}`;
            s2 = `<p:sp><p:nvSpPr><p:cNvPr id="${nid}" name="Conector ${nid}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E(x0)}" y="${E(y0)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${E(w)}" h="${E(h)}" fill="none">${camino}</a:path></a:pathLst></a:custGeom><a:noFill/>${ln}</p:spPr></p:sp>`;
          }
          const etq = (t, q, ancho, tamL, neg) => caja(id++, "Etiqueta", { x: (q.x - dx) / k - ancho / 2, y: (q.y - dy) / k - 14, w: ancho, h: 28 }, `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="es-MX" sz="${Math.round(tamL * k * 100)}" b="${neg ? 1 : 0}" dirty="0"/><a:t>${x(t)}</a:t></a:r></a:p>`).replace("<a:noFill/>", '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>');
          if (b.etiqueta) s2 += etq(b.etiqueta, T2(r.medio), Math.max(60, String(b.etiqueta).length * 8 + 24), 12, true);
          const cerca = (pt, u) => T2({ x: pt.x - u.x * 24 - u.y * 14, y: pt.y - u.y * 24 + u.x * 14 });
          if (b.etqIni) s2 += etq(b.etqIni, cerca(r.ini, r.dirIni), 50, 11, false);
          if (b.etqFin) s2 += etq(b.etqFin, cerca(r.fin, r.dirFin), 50, 11, false);
          return s2;
        }
        return "";
      }
      const figuras = figurasL.join("\n");
      // 8.3 Fondo de la diapositiva (color, degradado o imagen), del tema o propio.
      let fondo = "";
      // (al parchar una diapositiva del original, su fondo solo se escribe si alguien lo cambió: si no, ni se arma,
      // para no sumar una imagen ni una relación que nadie usa)
      if (!esDiagrama && NOM.diapo && !(O && O.sinFondo)) {
        const f = NOM.diapo.fondoDe(doc, p);
        let relleno = "";
        if (f && f.tipo === "color") relleno = `<a:solidFill><a:srgbClr val="${color(f.color, "FFFFFF")}"/></a:solidFill>`;
        else if (f && f.tipo === "rayas") relleno = `<a:solidFill><a:srgbClr val="${color(f.color, "FFFFFF")}"/></a:solidFill>`;
        else if (f && f.tipo === "degradado") relleno = `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="${color(f.c1, "FFFFFF")}"/></a:gs><a:gs pos="100000"><a:srgbClr val="${color(f.c2, "CFE0FF")}"/></a:gs></a:gsLst><a:lin ang="${Math.round((((+f.angulo || 135) - 90 + 360) % 360) * 60000)}" scaled="0"/></a:gradFill>`;
        else if (f && f.tipo === "imagen") { const rid = imagenRel(f.src); if (rid) relleno = `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`; }
        if (relleno) fondo = `<p:bg><p:bgPr>${relleno}<a:effectLst/></p:bgPr></p:bg>`;
      }
      // Transición y notas del orador.
      const tr = doc.valor("_doc", "trans:" + p), to = !esDiagrama && NOM.transOffice ? doc.valor("_doc", "transOffice:" + p) : null;
      const transicion = to ? NOM.transOffice.xml(to) : tr ? `<p:transition spd="med">${{ desvanecer: "<p:fade/>", deslizar: '<p:push dir="l"/>', empujar: '<p:push dir="u"/>', zoom: "<p:zoom/>" }[tr] || "<p:fade/>"}</p:transition>` : "";
      // Animaciones: el p:timing de PowerPoint con la lista de la diapositiva (cada bloque, con el número de su figura)
      const timing = !esDiagrama ? timingDe(doc, p, (bid) => { const i = bs.findIndex((b) => b.id === bid); const m = i >= 0 && /<p:cNvPr id="(\d+)"/.exec(figurasL[i] || ""); return m ? m[1] : null; }, (bid) => { const b = bs.find((q) => q.id === bid); return !!(b && (b.tipo === "texto" || b.html || b.texto)); }, true) : "";
      const nota = !esDiagrama ? doc.valor("_doc", "nota:" + p) || "" : "";
      if (nota && !O) { notas.push({ diapo: n + 1, texto: nota }); rels.push({ id: "rIdNota", tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide", destino: `../notesSlides/notesSlide${notas.length}.xml` }); }
      return { rels, bs, figurasL, fondo, transicion, nota: nota && !O ? notas.length : 0, xml: XML + `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${!esDiagrama && doc.valor("_doc", "oculta:" + p) ? ' show="0"' : ""}>
<p:cSld>${fondo}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${figuras}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transicion}${timing}</p:sld>`, timing };
    };
    const diapos = og ? [{}] : paginas.map((p, n) => armar(p, n, null));
    if (!diapos.length) throw new Error("No hay nada que exportar");
    const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
    const arbolVacio = '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>';
    const conNotas = notas.length > 0;
    const persP = !esDiagrama && !og ? personalizadasPptx(doc) : null, secsP = !esDiagrama && !og ? seccionesPptx(doc) : null;
    const presentacion = XML + `<p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst>${conNotas ? '<p:notesMasterIdLst><p:notesMasterId r:id="rIdNM"/></p:notesMasterIdLst>' : ""}
<p:sldIdLst>${diapos.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`).join("")}</p:sldIdLst><p:sldSz cx="${TD.w * EMU}" cy="${TD.h * EMU}"/><p:notesSz cx="6858000" cy="9144000"/>${persP ? custShowXml(persP, (pg) => { const i = paginas.indexOf(pg); return i < 0 ? null : "rId" + (i + 1); }) : ""}${secsP ? `<p:extLst>${seccionesExtXml(secsP, (pg) => { const i = paginas.indexOf(pg); return i < 0 ? null : String(256 + i); })}</p:extLst>` : ""}</p:presentation>`;
    const showPr = !esDiagrama ? showPrXml(doc, persP || []) : null;
    const presProps = showPr ? XML + `<p:presentationPr ${NS}>${showPr}</p:presentationPr>` : null;
    const maestro = XML + `<p:sldMaster ${NS}>${arbolVacio}<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`;
    const diseno = XML + `<p:sldLayout ${NS} type="blank" preserve="1">${arbolVacio}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
    // Notas del orador: un patrón de notas y una página de notas por diapositiva que las tenga.
    const phNota = (idx, tipo, nombre, x0, y0, w0, h0, cuerpo) => `<p:sp><p:nvSpPr><p:cNvPr id="${idx}" name="${nombre}"/><p:cNvSpPr><a:spLocks noGrp="1"${tipo === "sldImg" ? ' noRot="1" noChangeAspect="1"' : ""}/></p:cNvSpPr><p:nvPr><p:ph type="${tipo}"${tipo === "body" ? ' idx="1"' : ""}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x0}" y="${y0}"/><a:ext cx="${w0}" cy="${h0}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>${cuerpo || ""}</p:sp>`;
    const maestroNotas = XML + `<p:notesMaster ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${phNota(2, "sldImg", "Imagen de diapositiva", 685800, 685800, 5486400, 3086100)}${phNota(3, "body", "Notas", 685800, 4114800, 5486400, 4114800, '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="es-MX"/></a:p></p:txBody>')}</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
    const paginaNotas = (t) => XML + `<p:notes ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${phNota(2, "sldImg", "Imagen de diapositiva", 685800, 685800, 5486400, 3086100)}${phNota(3, "body", "Notas", 685800, 4114800, 5486400, 4114800, `<p:txBody><a:bodyPr/><a:lstStyle/>${String(t).split("\n").map((l) => `<a:p><a:r><a:rPr lang="es-MX" dirty="0"/><a:t>${x(l)}</a:t></a:r></a:p>`).join("")}</p:txBody>`)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;

    // ---------------------------------------------------------- a partir del archivo original
    /**
     * Cada diapositiva que no cambió sale tal cual. En la que cambió, cada figura sin tocar sale
     * textual del original y solo las editadas (o nuevas) se escriben de nuevo; lo que NOM no
     * muestra (SmartArt, video, marcadores vacíos) se queda en su lugar. Diseños, patrones y tema: intactos.
     */
    function conOrigen(N) {
      const O = NOM.origen, H = O.huella, pq = paqueteDe(og);
      const relP = pq.rels("").find((r) => /\/officeDocument$/.test(r.tipo || ""));
      const PRES = relP ? pq.resolver("", relP.destino) : "ppt/presentation.xml";
      const xp = pq.texto(PRES);
      if (!xp) throw new Error("falta la presentación");
      const raizP = O.hijosXml(xp).find((h) => O.local(h.n) === "presentation");
      if (!raizP || raizP.n !== "p:presentation") throw new Error("la presentación usa otro prefijo");
      const hsP = O.hijosXml(xp, raizP.ci, raizP.cf), una = (nm) => hsP.find((h) => O.local(h.n) === nm) || null;
      const szEl = una("sldSz"), lst = una("sldIdLst");
      if (!lst) throw new Error("presentación sin lista de diapositivas");
      const cx = szEl ? +O.attr(xp, szEl, "cx") : 12192000, cy = szEl ? +O.attr(xp, szEl, "cy") : 6858000;
      const kx = TD.w / (cx / EMU), ky = TD.h / (cy / EMU), km = Math.min(kx, ky);
      const esc = { k: 1 / km, fx: km / kx, fy: km / ky };
      const relsP = pq.rels(PRES), nombres = pq.nombres();
      let g = 1;
      while (nombres.some((nm) => nm.includes("/nom" + g + "-")) || relsP.some((r) => (r.id || "").startsWith("rIdNom" + g))) g++;
      PRE = "Nom" + g;
      const pre = PRE.toLowerCase() + "-", hpp = og.h.pp || {};
      const cambio = (k) => { const v = doc.valor("_doc", k), h0 = og.h.d ? og.h.d[k] : undefined; return v == null || v === "" ? h0 != null && h0 !== H("") : h0 !== H(v); };
      const igualB = (id) => { const hb = og.h.b && og.h.b[id]; return !!hb && (Array.isArray(hb) ? hb[0] : hb) === NOM.formatos.huellaBloque(doc, id); };
      const T_NOTA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide", T_NM = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
      const CT_NOTA = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
      let nNota = 0, maestroNotas = (() => { const r = relsP.find((q) => q.tipo === T_NM); return r ? pq.resolver(PRES, r.destino) : null; })();
      const piezasP = [];
      /** Notas del orador de una diapositiva: cambia el texto de su página de notas, o le crea una. */
      const ponNota = (ruta, texto) => {
        const rel = pq.rels(ruta).find((q) => q.tipo === T_NOTA);
        const parrafos = String(texto).split("\n").map((l) => `<a:p><a:r><a:rPr lang="es-MX" dirty="0"/><a:t>${x(l)}</a:t></a:r></a:p>`).join("");
        if (rel) {
          const np = pq.resolver(ruta, rel.destino), xn = pq.texto(np);
          if (!xn) return;
          // el cuadro de texto de las notas (marcador "body")
          const re = /<p:sp>[\s\S]*?<\/p:sp>/g;
          let m;
          while ((m = re.exec(xn))) {
            if (!/<p:ph\b[^>]*type="body"/.test(m[0])) continue;
            const nuevo = /<p:txBody>/.test(m[0]) ? m[0].replace(/<p:txBody>[\s\S]*<\/p:txBody>/, `<p:txBody><a:bodyPr/><a:lstStyle/>${parrafos || "<a:p/>"}</p:txBody>`) : m[0].replace("</p:sp>", `<p:txBody><a:bodyPr/><a:lstStyle/>${parrafos || "<a:p/>"}</p:txBody></p:sp>`);
            pq.pon(np, xn.slice(0, m.index) + nuevo + xn.slice(m.index + m[0].length));
            return;
          }
          return;
        }
        if (!texto) return;
        if (!maestroNotas) {
          // el original no tenía patrón de notas: se agrega uno sencillo con su tema
          maestroNotas = `ppt/notesMasters/${pre}notesMaster1.xml`;
          const tm = `ppt/theme/${pre}theme1.xml`;
          pq.pon(maestroNotas, N.maestroNotas); pq.tipo(maestroNotas, "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml");
          pq.pon(tm, tema()); pq.tipo(tm, "application/vnd.openxmlformats-officedocument.theme+xml");
          pq.agregaRel(maestroNotas, "rId1", T.tema, "/" + tm);
          pq.agregaRel(PRES, "rId" + PRE + "NM", T_NM, "/" + maestroNotas);
          const mm = una("sldMasterIdLst");
          if (mm) piezasP.push([mm.f, mm.f, `<p:notesMasterIdLst><p:notesMasterId r:id="rId${PRE}NM"/></p:notesMasterIdLst>`]);
        }
        const np = `ppt/notesSlides/${pre}notesSlide${++nNota}.xml`;
        pq.pon(np, N.paginaNotas(texto)); pq.tipo(np, CT_NOTA);
        pq.pon(pq.relsDe(np), RELS([{ id: "rId1", tipo: T_NM, destino: "/" + maestroNotas }, { id: "rId2", tipo: T.diapo, destino: "/" + ruta }]));
        pq.agregaRel(ruta, "rId" + PRE + "Nota", T_NOTA, "/" + np);
      };
      /** El diseño para las diapositivas nuevas: uno en blanco si lo hay. */
      const disenoNuevo = (() => {
        const ls = nombres.filter((nm) => /^ppt\/slideLayouts\/[^/]+\.xml$/.test(nm));
        return ls.find((nm) => /<p:sldLayout\b[^>]*\btype="blank"/.test(pq.texto(nm) || "")) || ls[ls.length - 1] || null;
      })();
      const usadas = new Set(), planes = [];
      let nNueva = 0;
      paginas.forEach((p, n) => {
        const ruta0 = doc.valor("_doc", "os:" + p), hp = hpp[p];
        if (!ruta0 || !hp || !pq.tiene(ruta0) || usadas.has(ruta0)) {
          // diapositiva hecha en NOM (o copia de otra): se arma completa sobre un diseño del original
          if (!disenoNuevo) throw new Error("el original no tiene diseños");
          const ruta = `ppt/slides/${pre}slide${++nNueva}.xml`;
          const r = armar(p, n, Object.assign({ idBase: 2 }, esc));
          pq.pon(ruta, r.xml); pq.tipo(ruta, "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
          pq.pon(pq.relsDe(ruta), RELS(r.rels.map((q) => (q.tipo === T.diseno ? { id: q.id, tipo: q.tipo, destino: "/" + disenoNuevo } : q))));
          const nota = doc.valor("_doc", "nota:" + p);
          if (nota) ponNota(ruta, nota);
          planes.push({ ruta, nueva: true });
          return;
        }
        usadas.add(ruta0);
        planes.push({ ruta: ruta0 });
        const orden = doc.deLaPagina(p).slice().sort((a, b) => (a.z || 0) - (b.z || 0)), ids = orden.map((b) => b.id);
        const cF = cambio("fondo:" + p), cT = cambio("trans:" + p) || cambio("transOffice:" + p), cN = cambio("nota:" + p), cA = cambio("anim:" + p);
        const cO = cambio("oculta:" + p);
        const mismas = ids.length === hp.ids.length && ids.every((id, i) => id === hp.ids[i]) && ids.every(igualB);
        if (mismas && !cF && !cT && !cN && !cA && !cO) return;   // intacta: ni se toca
        const xs = pq.texto(ruta0), raiz = O.hijosXml(xs).find((h) => O.local(h.n) === "sld");
        if (!raiz || raiz.n !== "p:sld") throw new Error("diapositiva con otro prefijo");
        const cSld = O.hijoXml(xs, "cSld", raiz.ci, raiz.cf), arbol = cSld && O.hijoXml(xs, "spTree", cSld.ci, cSld.cf);
        if (!arbol) throw new Error("diapositiva sin figuras");
        const us = O.hijosXml(xs, arbol.ci, arbol.cf), ps2 = [];
        // Ocultar diapositiva: solo el atributo show de <p:sld>
        if (cO) { const tag = xs.slice(raiz.i, raiz.ci).replace(/\sshow="[^"]*"/, ""); ps2.push([raiz.i, raiz.ci, doc.valor("_doc", "oculta:" + p) ? tag.replace(/\s*>$/, ' show="0">') : tag]); }
        const vivos = {};
        orden.forEach((b) => { if (b.ox != null) (vivos[b.ox] = vivos[b.ox] || []).push(b.id); });
        const intacto = (k) => { const a = hp.e[k] || [], v = vivos[k] || []; return a.length > 0 && a.length === v.length && a.every((id) => v.includes(id)) && a.every(igualB); };
        let maxId = 1;
        (xs.match(/<p:cNvPr\b[^>]*?\bid="\d+"/g) || []).forEach((m) => { maxId = Math.max(maxId, +/\bid="(\d+)"/.exec(m)[1]); });
        // Qué le pasó a cada elemento importado: "igual", "geo" (solo se movió o cambió de tamaño),
        // "texto" (cambió lo escrito) o "nuevo" (hay que escribirlo otra vez desde el modelo).
        // DEF-45: "fmt" = solo cambió el formato de texto de todo el cuadro (negrita, color, tamaño…): se
        // parcha el p:txBody original. Con "texto" puede venir además ese formato (fmts[k]).
        const modos = {}, fmts = {};
        const unoAUno = (k) => { const a = hp.e[k] || [], v = vivos[k] || []; return a.length === 1 && v.length === 1 && v[0] === a[0]; };
        const modo = (k) => {
          if (modos[k]) return modos[k];
          let m = "nuevo";
          const a = hp.e[k] || [], v = vivos[k] || [], el = us[k], q = a.length === 1 && og.h.q ? og.h.q[a[0]] : null;
          if (intacto(k)) m = "igual";
          else if (q && v.length === 1 && v[0] === a[0] && el && ["sp", "pic", "cxnSp", "graphicFrame"].includes(O.local(el.n))) {
            const pc = parcialesPptx(doc.bloque(a[0]));
            if (pc[0] === q[0]) m = "geo";
            else if (pc[1] === q[1] && O.local(el.n) === "sp" && O.hijoXml(xs, "txBody", el.ci, el.cf)) m = pc[2] === q[2] ? "texto" : "texto+geo";
            else if (O.local(el.n) === "sp" && O.hijoXml(xs, "txBody", el.ci, el.cf)) {
              const fc = formatoCambiadoPptx(doc.bloque(a[0]), q);
              if (fc) { fmts[k] = fc.fmt; m = (fc.texto ? "texto" : "fmt") + (fc.geo ? "+geo" : ""); }
            }
          }
          return (modos[k] = m);
        };
        const X = (v) => Math.round(v * esc.k * esc.fx * EMU), Y = (v) => Math.round(v * esc.k * esc.fy * EMU);
        /** El elemento original con la posición y el tamaño que tiene ahora el bloque. */
        const conGeo = (t, b) => {
          const at2 = `${b.rot ? ` rot="${Math.round(b.rot * 60000)}"` : ""}${b.volteoH ? ' flipH="1"' : ""}${b.volteoV ? ' flipV="1"' : ""}`;
          const dentro = `<a:off x="${X(b.x)}" y="${Y(b.y)}"/><a:ext cx="${X(Math.max(1, b.w))}" cy="${Y(Math.max(1, b.h))}"/>`;
          if (/^<p:graphicFrame\b/.test(t)) return t.replace(/<p:xfrm\b[^>]*>[\s\S]*?<\/p:xfrm>|<p:xfrm\b[^>]*\/>/, `<p:xfrm${at2}>${dentro}</p:xfrm>`);
          const i0 = t.indexOf("<p:spPr");
          if (i0 < 0) return t;
          const g0 = O.finEtiqueta(t, i0), vacio = t.charCodeAt(g0 - 1) === 47;
          const xf = `<a:xfrm${at2}>${dentro}</a:xfrm>`;
          if (vacio) return t.slice(0, g0 - 1) + ">" + xf + "</p:spPr>" + t.slice(g0 + 1);
          const fin = t.indexOf("</p:spPr>", g0), cuerpo = t.slice(g0 + 1, fin);
          const nuevo = /^\s*<a:xfrm\b/.test(cuerpo) ? cuerpo.replace(/<a:xfrm\b[^>]*\/>|<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/, xf) : xf + cuerpo;
          return t.slice(0, g0 + 1) + nuevo + t.slice(fin);
        };
        /** El elemento original con los párrafos nuevos (conserva su caja de texto, su forma y su marcador). */
        const conTexto = (el, fig) => {
          const m = /<p:txBody>([\s\S]*)<\/p:txBody>/.exec(fig || ""), tx = O.hijoXml(xs, "txBody", el.ci, el.cf);
          if (!m || !tx) return null;
          const ps3 = m[1].replace(/^\s*(<a:bodyPr\b[^>]*\/>|<a:bodyPr\b[\s\S]*?<\/a:bodyPr>)/, "").replace(/^\s*(<a:lstStyle\b[^>]*\/>|<a:lstStyle\b[\s\S]*?<\/a:lstStyle>)/, "");
          const cabeza = O.hijosXml(xs, tx.ci, tx.cf).filter((h) => ["bodyPr", "lstStyle"].includes(O.local(h.n))).map((h) => xs.slice(h.i, h.f)).join("");
          return xs.slice(el.i, tx.i) + "<p:txBody>" + (cabeza || "<a:bodyPr/>") + ps3 + "</p:txBody>" + xs.slice(tx.f, el.f);
        };
        /**
         * DEF-45: una figura que hubo que escribir de nuevo sigue siendo LA MISMA figura del archivo:
         * conserva su número, su nombre y su marcador de posición (p:nv…Pr del original), para que las
         * animaciones, el esquema y el diseño la sigan reconociendo. Un cuadro de texto conserva además
         * su p:spPr y su a:bodyPr (el modelo no guarda nada de eso: solo dónde está).
         */
        const NV = { sp: "nvSpPr", pic: "nvPicPr", cxnSp: "nvCxnSpPr", graphicFrame: "nvGraphicFramePr" };
        /**
         * El vínculo (a:hlinkClick) de una figura que se reescribe con su identidad original: el de NOM. El del original
         * se quita si es de los que NOM entiende (la persona lo cambió o lo quitó); los demás (abrir un archivo, una
         * macro) se quedan como venían.
         */
        const conVinculo = (res, fig, b) => {
          const cn = res && /<p:cNvPr\b[^>]*?(?:\/>|>[\s\S]*?<\/p:cNvPr>)/.exec(res);
          if (!cn) return res;
          const nuevo = (/<p:cNvPr\b[^>]*?>([\s\S]*?)<\/p:cNvPr>/.exec(fig || "") || [])[1] || "";
          const hl = b.accion ? (/<a:hlinkClick\b[^>]*?\/>/.exec(nuevo) || [])[0] || "" : "";
          const viejo = /<a:hlinkClick\b[^>]*?(?:\/>|>[\s\S]*?<\/a:hlinkClick>)/.exec(cn[0]);
          const entiende = viejo && (/hlinkshowjump|hlinksldjump/.test(viejo[0]) || !/\saction="/.test(viejo[0]) || /\saction=""/.test(viejo[0]));
          let c = cn[0];
          if (viejo && (hl || entiende)) c = c.replace(viejo[0], "");
          if (hl) c = /^<p:cNvPr\b[^>]*\/>$/.test(c) ? c.replace(/\/>$/, `>${hl}</p:cNvPr>`) : c.replace(/^(<p:cNvPr\b[^>]*?)>/, `$1>${hl}`);
          return c === cn[0] ? res : res.slice(0, cn.index) + c + res.slice(cn.index + cn[0].length);
        };
        const conIdentidad = (el, fig, b) => {
          const nl = O.local(el.n), nv = NV[nl];
          if (!nv || !fig || !fig.startsWith("<p:" + nl + ">")) return fig;
          if (nl === "sp" && b.tipo === "texto") {
            const t = conTexto(el, fig);
            if (t) { const q = og.h.q && og.h.q[b.id]; return q && parcialesPptx(b)[2] === q[2] ? t : conGeo(t, b); }
          }
          const o = O.hijoXml(xs, nv, el.ci, el.cf), i0 = fig.indexOf("<p:" + nv + ">"), i1 = fig.indexOf("</p:" + nv + ">");
          if (!o || i0 < 0 || i1 < i0) return fig;
          let orig = xs.slice(o.i, o.f);
          // el texto alternativo sí está en el modelo: si la persona lo cambió, sale el nuevo
          const d = /<p:cNvPr\b[^>]*?\sdescr="([^"]*)"/.exec(fig.slice(i0, i1));
          if (d) orig = orig.replace(/(<p:cNvPr\b[^>]*?)\sdescr="[^"]*"/, "$1").replace(/<p:cNvPr\b/, `<p:cNvPr descr="${d[1]}"`);
          return fig.slice(0, i0) + orig + fig.slice(i1 + nv.length + 5);
        };
        const r = armar(p, n, Object.assign({ idBase: maxId + 1, sinFondo: !cF,
          genera: (b) => !mismas && (b.ox == null || !hp.e[b.ox] || !us[b.ox] || /^(nuevo|texto)/.test(modo(b.ox))),
          hereda: (b) => (b.ox != null && hp.e[b.ox] && us[b.ox] && /^texto/.test(modo(b.ox)) ? fmts[b.ox] || {} : false) }, esc));
        let arbolSalida = xs.slice(arbol.ci, arbol.cf);
        if (!mismas) {
          const figDe = new Map(r.bs.map((b, i) => [b.id, r.figurasL[i]])), enBs = new Set(r.bs.map((b) => b.ox).filter((k) => k != null));
          const sale = [], puestos = new Set();
          let suelta = 0;
          while (suelta < us.length && ["nvGrpSpPr", "grpSpPr"].includes(O.local(us[suelta].n))) sale.push(xs.slice(us[suelta].i, us[suelta++].f));
          // lo que no se muestra en NOM, y lo importado que quedó fuera de la vista pero sigue igual
          const sueltas = (hasta) => { for (; suelta < hasta; suelta++) if (!hp.e[suelta] || (!enBs.has(suelta) && intacto(suelta))) sale.push(xs.slice(us[suelta].i, us[suelta].f)); };
          r.bs.forEach((b) => {
            const k = b.ox;
            if (k == null || !hp.e[k] || !us[k]) { sale.push(figDe.get(b.id) || ""); return; }
            const md = modo(k);
            if (puestos.has(k)) { if (md === "nuevo") sale.push(figDe.get(b.id) || ""); return; }
            puestos.add(k); sueltas(k);
            let t = null;
            if (md === "igual") t = xs.slice(us[k].i, us[k].f);
            else if (md === "geo") t = conGeo(xs.slice(us[k].i, us[k].f), b);
            else if (md === "texto" || md === "texto+geo") { t = conTexto(us[k], figDe.get(b.id)); if (t && md === "texto+geo") t = conGeo(t, b); if (!t) modos[k] = "nuevo"; }
            else if (md === "fmt" || md === "fmt+geo") {
              const f = fmts[k];
              t = conFormatoPptx(xs.slice(us[k].i, us[k].f), f, f.tTam ? Math.round((f.tTam / 1.333) * esc.k * 100) : 0);
              if (t && md === "fmt+geo") t = conGeo(t, b);
            }
            sale.push(t != null ? t : unoAUno(k) ? conVinculo(conIdentidad(us[k], figDe.get(b.id) || "", b), figDe.get(b.id) || "", b) : figDe.get(b.id) || "");
          });
          sueltas(us.length);
          const arbolNuevo = sale.join("");
          arbolSalida = arbolNuevo;
          ps2.push([arbol.ci, arbol.cf, arbolNuevo]);
          // Las animaciones apuntan a figuras por número. Como cada figura conserva el suyo, se quedan;
          // solo se quitan los efectos de las figuras (o párrafos) que ya no están.
          // (PowerPoint pone el p:timing dentro de mc:AlternateContent cuando usa lo de 2010, como el rebote: se poda cada uno)
          const tms = [];
          if (!cA) {
            const t0 = O.hijoXml(xs, "timing", raiz.ci, raiz.cf);
            if (t0) tms.push(t0);
            else O.hijosXml(xs, raiz.ci, raiz.cf).filter((h) => O.local(h.n) === "AlternateContent").forEach((h) => {
              let q = h.i;
              for (;;) { const i = xs.indexOf("<p:timing", q); if (i < 0 || i >= h.f) break; const f = xs.indexOf("</p:timing>", i); if (f < 0 || f > h.f) break; tms.push({ i, f: f + 11, ac: h }); q = f + 11; }
            });
          }
          for (const tm of tms) {
            const donde = new Map(), re = /<p:cNvPr\b[^>]*?\sid="(\d+)"/g;
            let mm;
            while ((mm = re.exec(arbolNuevo))) if (!donde.has(mm[1])) donde.set(mm[1], mm.index);
            const vivo = (spid, ultimo) => {
              if (!donde.has(spid)) return false;
              if (ultimo == null) return true;
              const i0 = donde.get(spid), i1 = arbolNuevo.indexOf("</p:sp>", i0);
              return ((i1 < 0 ? "" : arbolNuevo.slice(i0, i1)).match(/<a:p(?=[\s>\/])/g) || []).length > ultimo;
            };
            const antes = xs.slice(tm.i, tm.f), despues = podarTiming(antes, vivo);
            tm.queda = despues;
            if (despues !== antes && !(tm.ac && despues === "")) ps2.push([tm.i, tm.f, despues]);
          }
          // si no quedó ninguna animación, se quita el mc:AlternateContent entero (un Choice vacío no sirve)
          const acV = tms.find((tm) => tm.ac && tm.queda === "");
          if (acV) { for (let k = ps2.length - 1; k >= 0; k--) if (ps2[k][0] >= acV.ac.i && ps2[k][1] <= acV.ac.f) ps2.splice(k, 1); ps2.push([acV.ac.i, acV.ac.f, ""]); }
        }
        r.rels.filter((q) => q.tipo !== T.diseno).forEach((q) => { if (!mismas || cF) pq.agregaRel(ruta0, q.id, q.tipo, q.destino, q.externo); });
        if (cF) { const bg = O.hijoXml(xs, "bg", cSld.ci, cSld.cf); ps2.push(bg ? [bg.i, bg.f, r.fondo] : [cSld.ci, cSld.ci, r.fondo]); }
        const hsS = O.hijosXml(xs, raiz.ci, raiz.cf);
        const trEl = hsS.find((h) => O.local(h.n) === "transition") || hsS.find((h) => O.local(h.n) === "AlternateContent" && /<p:transition\b/.test(xs.slice(h.i, h.f)));
        // animaciones cambiadas en NOM: el p:timing se escribe otra vez con la lista (cada figura con el número que tiene ahora)
        let timingNuevo = null;
        if (cA) {
          const enSalida = new Set(); (arbolSalida.match(/<p:cNvPr\b[^>]*?\sid="(\d+)"/g) || []).forEach((m) => enSalida.add(/\sid="(\d+)"/.exec(m)[1]));
          const original = new Map();
          (doc.valor("_doc", "anim:" + p) || []).forEach((x) => { (function rec(m) { if (!Array.isArray(m)) return; if (/spTgt$/.test(m[0]) && m[1] && m[1].bid && m[1].spid) original.set(m[1].bid, m[1].spid); (m[2] || []).forEach(rec); })(x && x.ppt ? x.ppt.c : null); });
          const porId = new Map(orden.map((b) => [b.id, b])), figDe2 = new Map(r.bs.map((b, i) => [b.id, r.figurasL[i]]));
          const spidDe = (bid) => {
            const o = original.get(bid); if (o && enSalida.has(String(o))) return o;
            const b = porId.get(bid);
            if (b && b.ox != null && us[b.ox]) { const m = /<p:cNvPr\b[^>]*?\sid="(\d+)"/.exec(xs.slice(us[b.ox].i, us[b.ox].f)); if (m && enSalida.has(m[1])) return m[1]; }
            const m2 = /<p:cNvPr id="(\d+)"/.exec(figDe2.get(bid) || ""); return m2 && enSalida.has(m2[1]) ? m2[1] : null;
          };
          timingNuevo = timingDe(doc, p, spidDe, (bid) => { const b = porId.get(bid); return !!(b && (b.tipo === "texto" || b.html || b.texto)); });
          const tmV = hsS.find((h) => O.local(h.n) === "timing") || hsS.find((h) => O.local(h.n) === "AlternateContent" && /<p:timing\b/.test(xs.slice(h.i, h.f)));
          if (tmV) ps2.push([tmV.i, tmV.f, timingNuevo]);
          else if (timingNuevo) { const ref = trEl || hsS.find((h) => O.local(h.n) === "clrMapOvr"), pos = ref ? ref.f : cSld.f; if (!cT || !trEl) ps2.push([pos, pos, timingNuevo]); else timingNuevo = "+" + timingNuevo; }
        }
        if (cT) {
          const tras = typeof timingNuevo === "string" && timingNuevo[0] === "+" ? timingNuevo.slice(1) : "";
          // el sonido de la transición se queda si su relación sigue en esta diapositiva y es un audio (con «Aplicar a
          // todas», el rId de otra diapositiva aquí puede ser una imagen: entonces no se escribe)
          let trXml = r.transicion;
          const to2 = NOM.transOffice ? doc.valor("_doc", "transOffice:" + p) : null;
          if (to2 && to2.snd && to2.snd.rid && pq.rels(ruta0).some((q) => q.id === to2.snd.rid && /\/(audio|media)$/.test(q.tipo || ""))) trXml = NOM.transOffice.xml(to2, { snd: true });
          if (trEl) ps2.push([trEl.i, trEl.f, trXml + tras]);
          else { const cm = hsS.find((h) => O.local(h.n) === "clrMapOvr"), pos = cm ? cm.f : cSld.f; ps2.push([pos, pos, trXml + tras]); }
        }
        pq.pon(ruta0, aplicarPiezas(xs, ps2));
        if (cN) ponNota(ruta0, doc.valor("_doc", "nota:" + p) || "");
      });
      // ---- la lista de diapositivas: orden, nuevas y borradas
      const ents = O.hijosXml(xp, lst.ci, lst.cf).map((h) => {
        const txt = xp.slice(h.i, h.f), rid = (/\sr:id="([^"]*)"/.exec(txt) || [])[1], rel = relsP.find((q) => q.id === rid);
        return { txt, rid, id: +O.attr(xp, h, "id") || 0, ruta: rel ? pq.resolver(PRES, rel.destino) : null };
      });
      const mismaLista = ents.length === planes.length && ents.every((e, i) => e.ruta === planes[i].ruta);
      // el id y el r:id de cada página en la lista (las de antes conservan los suyos)
      const sldDe = planes.map((pl, i) => (mismaLista ? ents[i] : ents.find((q) => q.ruta === pl.ruta)) || null).map((e) => (e ? { id: e.id, rid: e.rid } : null));
      if (!mismaLista) {
        let maxId = ents.reduce((m, e) => Math.max(m, e.id), 255), k = 0;
        const lista = planes.map((pl, i) => {
          const e = ents.find((q) => q.ruta === pl.ruta);
          if (e) return e.txt;
          const rid = "rId" + PRE + "S" + ++k;
          pq.agregaRel(PRES, rid, T.diapo, "/" + pl.ruta);
          sldDe[i] = { id: ++maxId, rid };
          return `<p:sldId id="${maxId}" r:id="${rid}"/>`;
        });
        piezasP.push([lst.ci, lst.cf, lista.join("")]);
        const borradas = ents.filter((e) => e.ruta && !usadas.has(e.ruta));
        borradas.forEach((e) => {
          const nt = pq.rels(e.ruta).find((q) => q.tipo === T_NOTA);
          if (nt) { const np = pq.resolver(e.ruta, nt.destino); pq.quita(np); pq.quita(pq.relsDe(np)); pq.quitaTipo(np); }
          pq.quita(e.ruta); pq.quita(pq.relsDe(e.ruta)); pq.quitaTipo(e.ruta); pq.quitaRel(PRES, e.rid);
        });
      }
      // Presentaciones personalizadas y secciones: nombran diapositivas por relación y por número, así que con otra lista
      // (o si la persona las cambió) se escriben otra vez desde el documento, con las diapositivas de ahora.
      const pers = personalizadasPptx(doc);
      if (!mismaLista || cambio("secciones") || cambio("personalizadas")) {
        const porPag = new Map(paginas.map((pg, i) => [pg, sldDe[i]]));
        const ridDe = (pg) => (porPag.get(pg) || {}).rid || null, numDe = (pg) => { const q = porPag.get(pg); return q && q.id ? String(q.id) : null; };
        const cs = una("custShowLst"), nuevoCs = custShowXml(pers, ridDe);
        if (cs) piezasP.push([cs.i, cs.f, nuevoCs]);
        else if (nuevoCs) {
          const tras = ["photoAlbum", "custDataLst", "kinsoku", "defaultTextStyle", "modifyVerifier", "extLst"].map(una).find(Boolean);
          piezasP.push(tras ? [tras.i, tras.i, nuevoCs] : [raizP.cf, raizP.cf, nuevoCs]);
        }
        const ext = una("extLst"), nuevaSec = seccionesExtXml(seccionesPptx(doc), numDe);
        if (ext) {
          const xe = xp.slice(ext.i, ext.f), re = /<p:ext\b[^>]*uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}"[^>]*>[\s\S]*?<\/p:ext>/;
          const nx = re.test(xe) ? xe.replace(re, () => nuevaSec) : nuevaSec ? xe.replace(/<\/p:extLst>$/, () => nuevaSec + "</p:extLst>") : xe;
          if (nx !== xe) piezasP.push([ext.i, ext.f, /^<p:extLst>\s*<\/p:extLst>$/.test(nx) ? "" : nx]);
        } else if (nuevaSec) piezasP.push([raizP.cf, raizP.cf, `<p:extLst>${nuevaSec}</p:extLst>`]);
      }
      // Configurar presentación (presProps.xml): solo lo que se cambió; el color de la pluma y lo demás se quedan
      if (cambio("presentacion")) {
        const nuevo = showPrXml(doc, pers) || '<p:showPr showNarration="1"><p:present/><p:sldAll/></p:showPr>';
        const relPP = relsP.find((q) => /\/presProps$/.test(q.tipo || "")), rutaPP = relPP ? pq.resolver(PRES, relPP.destino) : null, xpp = rutaPP ? pq.texto(rutaPP) : null;
        if (xpp) {
          const raizPP = O.hijosXml(xpp).find((h) => O.local(h.n) === "presentationPr");
          const sp = raizPP && O.hijosXml(xpp, raizPP.ci, raizPP.cf).find((h) => O.local(h.n) === "showPr");
          const pre0 = (raizPP && raizPP.n.includes(":")) ? raizPP.n.slice(0, raizPP.n.indexOf(":") + 1) : "p:";
          const nuevoP = pre0 === "p:" ? nuevo : nuevo.replace(/<(\/?)p:/g, "<$1" + pre0);
          if (sp) {
            // los hijos que no son de qué se muestra ni de cómo (penClr, extLst…) se quedan
            const resto = sp.ci === sp.cf ? "" : O.hijosXml(xpp, sp.ci, sp.cf).filter((h) => !/^(present|browse|kiosk|sldAll|sldRg|custShow)$/.test(O.local(h.n))).map((h) => xpp.slice(h.i, h.f)).join("");
            const abre = nuevoP.slice(0, nuevoP.indexOf(">") + 1), dentro = nuevoP.slice(abre.length, nuevoP.lastIndexOf("</"));
            pq.pon(rutaPP, aplicarPiezas(xpp, [[sp.i, sp.f, abre + dentro + resto + `</${pre0}showPr>`]]));
          } else if (raizPP) pq.pon(rutaPP, aplicarPiezas(xpp, [[raizPP.ci, raizPP.ci, nuevoP]]));
        } else {
          const nr = "ppt/" + pre + "presProps.xml";
          pq.pon(nr, XML + `<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${nuevo}</p:presentationPr>`);
          pq.tipo(nr, "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml");
          pq.agregaRel(PRES, "rId" + PRE + "PP", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps", "/" + nr);
        }
      }
      if (piezasP.length) pq.pon(PRES, aplicarPiezas(xp, piezasP));
      medios.forEach((mm) => { pq.pon(mm.nombre, mm.datos); pq.extension(mm.nombre.split(".").pop(), MIME_IMG[mm.nombre.split(".").pop()] || "image/png"); });
      graficas.forEach((gx, i) => { const nm = `ppt/charts/${pre}chart${i + 1}.xml`; pq.pon(nm, gx); pq.tipo(nm, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"); });
      return pq.zip();
    }
    if (og) return conOrigen({ maestroNotas, paginaNotas });
    return zip([
      { nombre: "[Content_Types].xml", datos: CT(`<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${presProps ? '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' : ""}
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
${conNotas ? `<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/><Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` : ""}
${notas.map((_, i) => `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`).join("")}
${graficas.map((_, i) => `<Override PartName="/ppt/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("")}
${diapos.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("")}`) },
      { nombre: "_rels/.rels", datos: RELS([{ tipo: T.doc, destino: "ppt/presentation.xml" }]) },
      { nombre: "ppt/presentation.xml", datos: presentacion },
      { nombre: "ppt/_rels/presentation.xml.rels", datos: RELS([...diapos.map((_, i) => ({ id: "rId" + (i + 1), tipo: T.diapo, destino: `slides/slide${i + 1}.xml` })),
        { id: "rIdM", tipo: T.maestro, destino: "slideMasters/slideMaster1.xml" }, { id: "rIdT", tipo: T.tema, destino: "theme/theme1.xml" },
        ...(conNotas ? [{ id: "rIdNM", tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster", destino: "notesMasters/notesMaster1.xml" }] : []),
        ...(presProps ? [{ id: "rIdPP", tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps", destino: "presProps.xml" }] : [])]) },
      ...(presProps ? [{ nombre: "ppt/presProps.xml", datos: presProps }] : []),
      { nombre: "ppt/slideMasters/slideMaster1.xml", datos: maestro },
      { nombre: "ppt/slideMasters/_rels/slideMaster1.xml.rels", datos: RELS([{ tipo: T.diseno, destino: "../slideLayouts/slideLayout1.xml" }, { tipo: T.tema, destino: "../theme/theme1.xml" }]) },
      { nombre: "ppt/slideLayouts/slideLayout1.xml", datos: diseno },
      { nombre: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", datos: RELS([{ tipo: T.maestro, destino: "../slideMasters/slideMaster1.xml" }]) },
      { nombre: "ppt/theme/theme1.xml", datos: tema() },
      ...(conNotas ? [
        { nombre: "ppt/notesMasters/notesMaster1.xml", datos: maestroNotas },
        { nombre: "ppt/notesMasters/_rels/notesMaster1.xml.rels", datos: RELS([{ tipo: T.tema, destino: "../theme/theme2.xml" }]) },
        { nombre: "ppt/theme/theme2.xml", datos: tema() },
        ...notas.flatMap((nn, i) => [
          { nombre: `ppt/notesSlides/notesSlide${i + 1}.xml`, datos: paginaNotas(nn.texto) },
          { nombre: `ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`, datos: RELS([{ tipo: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster", destino: "../notesMasters/notesMaster1.xml" }, { tipo: T.diapo, destino: `../slides/slide${nn.diapo}.xml` }]) },
        ])] : []),
      ...diapos.flatMap((d, i) => [{ nombre: `ppt/slides/slide${i + 1}.xml`, datos: d.xml }, { nombre: `ppt/slides/_rels/slide${i + 1}.xml.rels`, datos: RELS(d.rels) }]),
      ...graficas.map((g, i) => ({ nombre: `ppt/charts/chart${i + 1}.xml`, datos: g })),
      ...medios,
    ]);
  }
  /** 8.3 Gráfica de una diapositiva como gráfica nativa de PowerPoint (con sus datos escritos). */
  function graficaPptx(b) {
    if (!NOM.diapo || !NOM.graficas) return null;
    const r = NOM.diapo.rangoUsado(b);
    if (!r) return null;
    const ev = F().evaluador((c, f) => b["c:" + F().ref(c, f)] || "");
    const d = NOM.graficas.datos(ev, r);
    if (!d) return null;
    const col = (NOM.graficas.PALETAS[b.colores] || NOM.graficas.PALETAS.clasica).c;
    const lit = (vals) => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (v == null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:numLit>`;
    const cat = `<c:cat><c:strLit><c:ptCount val="${d.etiquetas.length}"/>${d.etiquetas.map((e, i) => `<c:pt idx="${i}"><c:v>${x(e)}</c:v></c:pt>`).join("")}</c:strLit></c:cat>`;
    const relleno = (i) => `<c:spPr><a:solidFill><a:srgbClr val="${col[i % col.length].slice(1)}"/></a:solidFill></c:spPr>`;
    const tx = (s) => `<c:tx><c:v>${x(s.nombre)}</c:v></c:tx>`;
    const titulo = (t0) => `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${x(t0)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
    const ejes = `<c:axId val="1001"/><c:axId val="1002"/>`;
    const ejeX = (pos0) => `<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos0}"/>${b.ejeX ? titulo(b.ejeX) : ""}<c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
    const ejeY = (pos0) => `<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos0}"/><c:majorGridlines/>${b.ejeY ? titulo(b.ejeY) : ""}<c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="1001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
    const clase = b.clase || "barras";
    let plot;
    if (clase === "pastel") {
      const s = d.series[0];
      plot = `<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${tx(s)}${d.etiquetas.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>${relleno(i)}</c:dPt>`).join("")}${cat}<c:val>${lit(s.valores)}</c:val></c:ser><c:firstSliceAng val="0"/></c:pieChart>`;
    } else if (clase === "dispersion") {
      const xs = d.series.length > 1 ? d.series[0] : null, ys = xs ? d.series.slice(1) : d.series;
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${ys.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx(s)}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/>${relleno(i)}</c:marker>${xs ? `<c:xVal>${lit(xs.valores)}</c:xVal>` : ""}<c:yVal>${lit(s.valores)}</c:yVal><c:smooth val="0"/></c:ser>`).join("")}${ejes}</c:scatterChart>`
        + `<c:valAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${b.ejeX ? titulo(b.ejeX) : ""}<c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="1002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>${ejeY("l")}`;
    } else {
      const sers = d.series.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx(s)}${relleno(i)}${cat}<c:val>${lit(s.valores)}</c:val>${clase === "lineas" ? '<c:smooth val="0"/>' : ""}</c:ser>`).join("");
      if (clase === "lineas") plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${ejes}</c:lineChart>`;
      else if (clase === "area") plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${ejes}</c:areaChart>`;
      else plot = `<c:barChart><c:barDir val="${clase === "barras-h" ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="80"/>${ejes}</c:barChart>`;
      plot += ejeX(clase === "barras-h" ? "l" : "b") + ejeY(clase === "barras-h" ? "b" : "l");
    }
    return XML + `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${b.titulo ? titulo(b.titulo) + `<c:autoTitleDeleted val="0"/>` : `<c:autoTitleDeleted val="1"/>`}<c:plotArea><c:layout/>${plot}</c:plotArea>${d.series.length > 1 || clase === "pastel" ? `<c:legend><c:legendPos val="${clase === "pastel" ? "r" : "b"}"/><c:overlay val="0"/></c:legend>` : ""}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  }

  function tema() {
    const c = (n, v) => `<a:${n}><a:srgbClr val="${v}"/></a:${n}>`;
    const f = '<a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>';
    const relleno = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    return XML + `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="NOM"><a:themeElements>
<a:clrScheme name="NOM">${c("dk1", "1F1E26")}${c("lt1", "FFFFFF")}${c("dk2", "3A3752")}${c("lt2", "F5F3EF")}${c("accent1", "5B5BD6")}${c("accent2", "3D72F2")}${c("accent3", "2FA868")}${c("accent4", "F08C2A")}${c("accent5", "8B5CF6")}${c("accent6", "D6336C")}${c("hlink", "3D72F2")}${c("folHlink", "8B5CF6")}</a:clrScheme>
<a:fontScheme name="NOM"><a:majorFont>${f}</a:majorFont><a:minorFont>${f}</a:minorFont></a:fontScheme>
<a:fmtScheme name="NOM"><a:fillStyleLst>${relleno.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${'<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'.repeat(3)}</a:lnStyleLst>
<a:effectStyleLst>${"<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${relleno.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
  }

  // ------------------------------------------------------------ importar PowerPoint (.pptx)
  /**
   * .pptx → { paginas, bloques, doc, aviso }: cada diapositiva con sus cuadros de texto (con formato y
   * viñetas), figuras, imágenes (con recorte), tablas, gráficas, fondo, notas y transición.
   */
  async function leerPptx(archivo) {
    if (archivo.size > 60 * 1024 * 1024) throw new Error("El archivo pesa más de 60 MB");
    const u8 = new Uint8Array(await archivo.arrayBuffer());
    const z = NOM.origen.abrirZip(u8);
    let oxActual = null;   // qué hijo del árbol de figuras de la diapositiva se está leyendo
    const P = (t) => new DOMParser().parseFromString(t, "application/xml");
    const pres = await z.texto("ppt/presentation.xml");
    if (!pres) throw new Error("El archivo no parece una presentación de PowerPoint (.pptx)");
    const hs = (n, nombre) => [...(n ? n.childNodes : [])].filter((c) => c.nodeType === 1 && (!nombre || c.localName === nombre));
    const h1 = (n, nombre) => hs(n, nombre)[0] || null;
    const todos = (n, nombre) => (n ? [...n.getElementsByTagNameNS("*", nombre)] : []);
    const at = (n, a) => (n ? n.getAttribute(a) : null);
    const RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const rid = (n, a) => (n ? n.getAttributeNS(RNS, a) || n.getAttribute("r:" + a) : null);
    const carpeta = (ruta) => ruta.slice(0, ruta.lastIndexOf("/") + 1);
    const resolver = (base, destino) => {
      if (destino.startsWith("/")) return destino.slice(1);
      const partes = (carpeta(base) + destino).split("/"), r = [];
      partes.forEach((p) => { if (p === "..") r.pop(); else if (p && p !== ".") r.push(p); });
      return r.join("/");
    };
    /**
     * La acción al hacer clic de una figura (a:hlinkClick en su cNvPr), como la guarda NOM en el bloque:
     * { tipo: "sig" | "ant" | "primera" | "ultima" | "fin" | "volver" } (saltos de la presentación), { tipo: "diapo", ruta }
     * (otra diapositiva; la ruta se cambia por la página al terminar) o { tipo: "url", url } (página web o correo).
     */
    function accionDe(el, rels) {
      const nv = hs(el).find((c) => /^nv/.test(c.localName)), cNv = nv && h1(nv, "cNvPr"), hl = cNv && h1(cNv, "hlinkClick");
      if (!hl) return null;
      const act = at(hl, "action") || "", rid = hl.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || at(hl, "r:id") || "";
      const salto = /jump=(\w+)/.exec(act);
      if (/hlinkshowjump/.test(act) && salto) return { tipo: { nextslide: "sig", previousslide: "ant", firstslide: "primera", lastslide: "ultima", endshow: "fin", lastslideviewed: "volver" }[salto[1].toLowerCase()] || "sig" };
      const rel = rid && rels[rid];
      if (/hlinksldjump/.test(act)) return rel ? { tipo: "diapo", ruta: rel.destino } : null;
      if (!act && rel && rel.externo && /^(https?:|mailto:)/i.test(rel.bruto || "")) return { tipo: "url", url: rel.bruto };
      return null;
    }
    /**
     * El video o audio de una imagen de PowerPoint (a:videoFile / a:audioFile y p14:media en su nvPr), o undefined:
     * { tipo, ruta (la parte del archivo con el medio), recorte { ini, fin } (ms que se quitan al principio y al final,
     * p14:trim), fundido { ent, sal } (ms, p14:fade), marcadores [{ n, t }] (ms, p14:bmkLst) }. El medio no se copia: al
     * presentar se saca del archivo original que guarda el documento (js/medios-diapositiva.js). Uno vinculado (fuera
     * del archivo) no se puede reproducir y se queda como imagen.
     */
    function medioDePic(el, rels) {
      const nv = h1(el, "nvPicPr"), nvPr = nv && h1(nv, "nvPr");
      if (!nvPr) return undefined;
      const vf = h1(nvPr, "videoFile"), af = h1(nvPr, "audioFile"), m14 = todos(nvPr, "media")[0];
      if (!vf && !af && !m14) return undefined;
      const r14 = m14 && rels[rid(m14, "embed")], rL = (vf || af) && rels[rid(vf || af, "link")];
      const rel = r14 && !r14.externo ? r14 : rL && !rL.externo ? rL : null;
      if (!rel || !rel.destino) return undefined;
      const m = { tipo: af && !vf ? "audio" : "video", ruta: rel.destino };
      if (!vf && !af && /\.(mp3|m4a|wav|wma|aac|ogg|oga|flac|midi?)$/i.test(rel.destino)) m.tipo = "audio";
      const tr = m14 && todos(m14, "trim")[0], fd = m14 && todos(m14, "fade")[0], bm = m14 ? todos(m14, "bmk") : [];
      const ms = (v) => Math.max(0, Math.round(+v || 0));
      if (tr && (ms(at(tr, "st")) || ms(at(tr, "end")))) m.recorte = { ini: ms(at(tr, "st")), fin: ms(at(tr, "end")) };
      if (fd && (ms(at(fd, "in")) || ms(at(fd, "out")))) m.fundido = { ent: ms(at(fd, "in")), sal: ms(at(fd, "out")) };
      if (bm.length) m.marcadores = bm.map((b) => ({ n: at(b, "name") || "", t: ms(at(b, "time")) })).sort((p, q) => p.t - q.t);
      return m;
    }
    async function relsDe(ruta) {
      const t = await z.texto(carpeta(ruta) + "_rels/" + ruta.slice(ruta.lastIndexOf("/") + 1) + ".rels");
      const m = {};
      if (t) todos(P(t), "Relationship").forEach((r) => (m[at(r, "Id")] = { destino: resolver(ruta, at(r, "Target")), tipo: at(r, "Type") || "", externo: at(r, "TargetMode") === "External", bruto: at(r, "Target") || "" }));
      return m;
    }
    const docP = P(pres);
    const sz = todos(docP, "sldSz")[0];
    const cx = sz ? +at(sz, "cx") : 12192000, cy = sz ? +at(sz, "cy") : 6858000;
    // El tamaño de NOM con la proporción más parecida (16:9, 4:3, carta, cuadrado…).
    const proporcion = cx / cy, TAMS = NOM.diapo ? NOM.diapo.TAMANOS : { "16:9": { w: 960, h: 540 } };
    const tamDiapo = Object.keys(TAMS).reduce((m, k) => (Math.abs(TAMS[k].w / TAMS[k].h - proporcion) < Math.abs(TAMS[m].w / TAMS[m].h - proporcion) ? k : m), "16:9");
    const T = TAMS[tamDiapo];
    const kx = T.w / (cx / EMU), ky = T.h / (cy / EMU);
    const relsPres = await relsDe("ppt/presentation.xml");
    const ids = todos(docP, "sldId").map((s) => relsPres[rid(s, "id")]).filter(Boolean);
    const GEO_INV = { rect: "rect", roundRect: "rect", ellipse: "elipse", diamond: "rombo", triangle: "triangulo", rtTriangle: "triangulo-r", pentagon: "pentagono", hexagon: "hexagono", octagon: "octagono",
      star5: "estrella", heart: "corazon", cloud: "nube", wedgeRoundRectCallout: "burbuja", wedgeRectCallout: "burbuja", rightArrow: "flecha", leftRightArrow: "flecha-doble", chevron: "chevron", plus: "cruz",
      flowChartInputOutput: "paralelogramo", parallelogram: "paralelogramo", trapezoid: "trapecio", can: "cilindro", flowChartDocument: "documento", foldedCorner: "nota", flowChartProcess: "rect",
      flowChartDecision: "rombo", flowChartTerminator: "elipse", flowChartMagneticDisk: "cilindro" };
    // Colores del tema: casi todas las figuras de una presentación real dicen "acento 1" o "texto 1"
    // en vez de un color escrito.
    const esquema = {}, lineasTema = {}, fuentesTema = { mayor: null, menor: null }, fondosTema = {};
    let relsTema = {};
    const capas = new Map();   // diseño → formas heredadas del patrón y del diseño (se pintan debajo)
    // Cada patrón trae su propio tema (una presentación puede mezclar varios): colores, fuentes, grosores de línea y
    // rellenos de fondo. Se leen una vez por tema y se cambian al pasar a una diapositiva de otro patrón.
    const temas = new Map();
    async function leerTema(rutaTema) {
      if (temas.has(rutaTema)) return temas.get(rutaTema);
      const t = { esquema: {}, lineas: {}, fondos: {}, fuentes: { mayor: null, menor: null }, rels: {} };
      const tt = await z.texto(rutaTema), dt = tt ? P(tt) : null;
      const cs = dt ? todos(dt, "clrScheme")[0] : null;
      hs(cs).forEach((c) => { const v = h1(c, "srgbClr") ? at(h1(c, "srgbClr"), "val") : h1(c, "sysClr") ? at(h1(c, "sysClr"), "lastClr") : null; if (v && /^[0-9a-f]{6}$/i.test(v)) t.esquema[c.localName] = v.toUpperCase(); });
      Object.assign(t.esquema, { tx1: t.esquema.dk1, bg1: t.esquema.lt1, tx2: t.esquema.dk2, bg2: t.esquema.lt2 });
      // grosor (puntos) de las líneas de los estilos del tema: <p:style><a:lnRef idx="2"> usa el segundo
      const ls = dt ? todos(todos(dt, "lnStyleLst")[0], "ln") : [];
      ls.forEach((l, i) => (t.lineas[i + 1] = +(at(l, "w") || 9525) / 12700));
      // rellenos de fondo del tema (bgRef idx="1001"… usa el primero)
      const bgl = dt ? todos(dt, "bgFillStyleLst")[0] : null;
      hs(bgl).forEach((f, i) => { f.__tema = true; t.fondos[i + 1] = f; });
      t.rels = tt ? await relsDe(rutaTema) : {};
      // fuentes del tema: «+mj-lt» (títulos) y «+mn-lt» (cuerpo)
      ["majorFont", "minorFont"].forEach((k, i) => { const f = dt && todos(dt, k)[0], l = f && h1(f, "latin"); if (l && at(l, "typeface")) t.fuentes[i ? "menor" : "mayor"] = at(l, "typeface"); });
      temas.set(rutaTema, t);
      return t;
    }
    function usarTema(t) {
      [[esquema, t.esquema], [lineasTema, t.lineas], [fondosTema, t.fondos]].forEach(([a, b]) => { for (const k in a) delete a[k]; Object.assign(a, b); });
      fuentesTema.mayor = t.fuentes.mayor; fuentesTema.menor = t.fuentes.menor;
      relsTema = t.rels;
    }
    const relEstTablas = Object.values(relsPres).find((q) => /\/tableStyles$/.test(q.tipo));
    const xmlEstTablas = relEstTablas ? (await z.texto(relEstTablas.destino) || "").length ? P(await z.texto(relEstTablas.destino)) : null : null;
    const temaPres = await leerTema((Object.values(relsPres).find((q) => /\/theme$/.test(q.tipo)) || {}).destino || "ppt/theme/theme1.xml");
    usarTema(temaPres);
    /** Más claro o más oscuro, como lo pide el archivo (lumMod, lumOff, tint, shade: en milésimas de %).
        Como Office: shade y tint se aplican en RGB lineal (gamma 2.3); lumMod y lumOff en la luminosidad. */
    function matiz(hex, c) {
      // shade y tint primero, en el orden del archivo
      for (const t of hs(c)) {
        const k = +at(t, "val") / 100000;
        if (t.localName !== "shade" && t.localName !== "tint") continue;
        const lin = [0, 2, 4].map((i) => Math.pow(parseInt(hex.slice(i, i + 2), 16) / 255, 2.3));
        const nuevo = lin.map((x) => (t.localName === "shade" ? x * k : 1 - (1 - x) * k));
        hex = nuevo.map((x) => Math.round(Math.pow(Math.max(0, Math.min(1, x)), 1 / 2.3) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
      }
      let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
      const v = (k) => { const e = h1(c, k); return e && k !== "shade" && k !== "tint" ? +at(e, "val") / 100000 : null; };
      const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
      let h = 0, sa = 0, l = (mx + mn) / 2;
      if (mx !== mn) { const d = mx - mn; sa = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
      if (v("hueMod") != null) h = (h * v("hueMod")) % 1;
      if (v("hueOff") != null) h = (((h + (v("hueOff") * 100000) / 21600000) % 1) + 1) % 1;   // (hueOff va en 60 000avos de grado)
      if (v("satMod") != null) sa = Math.max(0, Math.min(1, sa * v("satMod")));
      if (v("satOff") != null) sa = Math.max(0, Math.min(1, sa + v("satOff")));
      if (v("lumMod") != null) l *= v("lumMod");
      if (v("lumOff") != null) l += v("lumOff");
      if (v("shade") != null) l *= v("shade");
      if (v("tint") != null) l = l * v("tint") + (1 - v("tint"));
      l = Math.max(0, Math.min(1, l));
      const q = l < 0.5 ? l * (1 + sa) : l + sa - l * sa, p2 = 2 * l - q;
      const f = (t) => { t = (t + 1) % 1; return t < 1 / 6 ? p2 + (q - p2) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p2 + (q - p2) * (2 / 3 - t) * 6 : p2; };
      [r, g, b] = sa === 0 ? [l, l, l] : [f(h + 1 / 3), f(h), f(h - 1 / 3)];
      return [r, g, b].map((x2) => Math.round(x2 * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
    }
    const hexDe = (fill, ph) => {
      const c = fill && todos(fill, "srgbClr")[0];
      if (c && /^[0-9a-f]{6}$/i.test(at(c, "val") || "")) return "#" + (hs(c).length ? matiz(at(c, "val"), c) : at(c, "val").toUpperCase());
      const e = fill && todos(fill, "schemeClr")[0], base = e && (at(e, "val") === "phClr" && ph ? String(ph).replace("#", "").toUpperCase() : esquema[at(e, "val")]);
      if (base) return "#" + (hs(e).length ? matiz(base, e) : base);
      const y = fill && todos(fill, "sysClr")[0];
      if (y && /^[0-9a-f]{6}$/i.test(at(y, "lastClr") || "")) return "#" + at(y, "lastClr").toUpperCase();
      const pr = fill && todos(fill, "prstClr")[0], nom = pr && { black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00", gray: "808080" }[at(pr, "val")];
      return nom ? "#" + nom : null;
    };
    /** Color de un elemento de color suelto (srgbClr, schemeClr…), con phClr si viene de un estilo del tema. */
    const colorDeElemento = (ce, ph) => {
      if (!ce) return null;
      let base = null;
      if (ce.localName === "srgbClr") base = (at(ce, "val") || "").toUpperCase();
      else if (ce.localName === "schemeClr") base = at(ce, "val") === "phClr" && ph ? String(ph).replace("#", "").toUpperCase() : esquema[at(ce, "val")];
      else if (ce.localName === "sysClr") base = (at(ce, "lastClr") || "").toUpperCase();
      if (!/^[0-9A-F]{6}$/.test(base || "")) return null;
      return "#" + (hs(ce).length ? matiz(base, ce) : base);
    };
    /** Dos tonos (duotone de Office): cada píxel va del primer color (oscuros) al segundo (claros). */
    function duotono(src, c1, c2) {
      return new Promise((ok) => {
        const img = new Image();
        img.onload = () => {
          try {
            const k = Math.min(1, 1600 / Math.max(img.width, img.height));
            const cv = document.createElement("canvas"); cv.width = Math.max(1, Math.round(img.width * k)); cv.height = Math.max(1, Math.round(img.height * k));
            const g = cv.getContext("2d"); g.drawImage(img, 0, 0, cv.width, cv.height);
            const d = g.getImageData(0, 0, cv.width, cv.height), px = d.data;
            const a = [1, 3, 5].map((i) => parseInt(c1.slice(i, i + 2), 16)), b = [1, 3, 5].map((i) => parseInt(c2.slice(i, i + 2), 16));
            for (let i = 0; i < px.length; i += 4) {
              const l = (0.2126 * px[i] + 0.7152 * px[i + 1] + 0.0722 * px[i + 2]) / 255;
              px[i] = a[0] + (b[0] - a[0]) * l; px[i + 1] = a[1] + (b[1] - a[1]) * l; px[i + 2] = a[2] + (b[2] - a[2]) * l;
            }
            g.putImageData(d, 0, 0);
            ok(cv.toDataURL("image/jpeg", 0.9));
          } catch (e) { ok(src); }
        };
        img.onerror = () => ok(src);
        img.src = src;
      });
    }
    /** Color de un relleno como CSS, con su transparencia (alpha) si la tiene. */
    const colorCss = (fill, ph) => {
      const hx = hexDe(fill, ph);
      if (!hx) return null;
      const cl = fill && (todos(fill, "srgbClr")[0] || todos(fill, "schemeClr")[0] || todos(fill, "sysClr")[0] || todos(fill, "prstClr")[0]);
      const al = cl && h1(cl, "alpha");
      if (!al) return hx;
      const n = parseInt(hx.slice(1), 16), a = Math.max(0, Math.min(1, +at(al, "val") / 100000));
      return `rgba(${n >> 16},${(n >> 8) & 255},${n & 255},${Math.round(a * 1000) / 1000})`;
    };
    // Cuánto cubre cada trama de Office (sus 8 × 8 puntos): a tamaño normal se ve como un tono parejo.
    const COBERTURA = { pct5: 0.05, pct10: 0.1, pct20: 0.2, pct25: 0.25, pct30: 0.3, pct40: 0.4, pct50: 0.5, pct60: 0.6, pct70: 0.7, pct75: 0.75, pct80: 0.8, pct90: 0.9,
      horz: 0.125, vert: 0.125, ltHorz: 0.25, ltVert: 0.25, dkHorz: 0.5, dkVert: 0.5, narHorz: 0.5, narVert: 0.5, dashHorz: 0.19, dashVert: 0.19, cross: 0.23,
      dnDiag: 0.125, upDiag: 0.125, ltDnDiag: 0.25, ltUpDiag: 0.25, dkDnDiag: 0.5, dkUpDiag: 0.5, wdDnDiag: 0.375, wdUpDiag: 0.375, dashDnDiag: 0.19, dashUpDiag: 0.19,
      diagCross: 0.25, smCheck: 0.5, lgCheck: 0.5, smGrid: 0.44, lgGrid: 0.23, dotGrid: 0.125, smConfetti: 0.125, lgConfetti: 0.3, horzBrick: 0.3, diagBrick: 0.25,
      solidDmnd: 0.5, openDmnd: 0.25, dotDmnd: 0.125, plaid: 0.4, sphere: 0.5, weave: 0.4, divot: 0.2, shingle: 0.25, wave: 0.25, trellis: 0.55, zigZag: 0.3 };
    function colorTrama(prst, cf, cb) {
      const a = COBERTURA[prst] != null ? COBERTURA[prst] : 0.5;
      const c = (h, i) => parseInt(String(h).slice(1 + i * 2, 3 + i * 2), 16);
      return "#" + [0, 1, 2].map((i) => Math.round(c(cf, i) * a + c(cb, i) * (1 - a)).toString(16).padStart(2, "0")).join("").toUpperCase();
    }
    /** Ancho y alto en píxeles de una imagen (data:). */
    const medidasImagen = (src) => new Promise((ok) => { const im = new Image(); im.onload = () => ok({ w: im.naturalWidth || 1, h: im.naturalHeight || 1 }); im.onerror = () => ok({ w: 64, h: 64 }); im.src = src; });
    /** Resolución (puntos por pulgada) que guarda una imagen PNG (pHYs) o JPEG (JFIF); null si no la dice. */
    function dpiImagen(src) {
      try {
        const i0 = src.indexOf(",") + 1, b64 = src.slice(i0, i0 + 87380);
        const bin = atob(b64.slice(0, b64.length - (b64.length % 4))), u = (i) => bin.charCodeAt(i) & 255, u32 = (i) => ((u(i) << 24) | (u(i + 1) << 16) | (u(i + 2) << 8) | u(i + 3)) >>> 0;
        if (bin.startsWith("\x89PNG")) {
          for (let i = 8; i + 12 <= bin.length; ) {
            const len = u32(i), tipo = bin.slice(i + 4, i + 8);
            if (tipo === "pHYs") { const ppm = u32(i + 8); return u(i + 16) === 1 && ppm ? ppm * 0.0254 : null; }
            if (tipo === "IDAT" || tipo === "IEND") break;
            i += 12 + len;
          }
        } else if (u(0) === 0xff && u(1) === 0xd8 && bin.slice(6, 10) === "JFIF") {
          const unidad = u(13), x = (u(14) << 8) | u(15);
          return unidad === 1 && x ? x : unidad === 2 && x ? x * 2.54 : null;
        }
      } catch (e) { /* sin resolución: 96 ppp */ }
      return null;
    }
    /** Geometría libre (custGeom) en el formato de geometrias-datos.js: guías, rectángulo del texto y trazos. */
    function geomLibre(spPr) {
      const cg = spPr && h1(spPr, "custGeom");
      if (!cg) return null;
      const num = (v) => (v != null && v !== "" && !isNaN(+v) ? +v : v || 0);
      const guias = (n) => (n ? hs(n).filter((g) => g.localName === "gd").map((g) => [at(g, "name"), at(g, "fmla") || "val 0"]) : []);
      const pts = (c) => todos(c, "pt").flatMap((q) => [num(at(q, "x")), num(at(q, "y"))]);
      const def = { p: [] }, av = guias(h1(cg, "avLst")), gd = guias(h1(cg, "gdLst"));
      if (av.length) def.av = av;
      if (gd.length) def.gd = gd;
      const rc = h1(cg, "rect");
      if (rc) def.rect = ["l", "t", "r", "b"].map((k) => num(at(rc, k)));
      hs(h1(cg, "pathLst")).filter((q) => q.localName === "path").forEach((q) => {
        const t = { c: [] };
        if (at(q, "w")) t.w = +at(q, "w");
        if (at(q, "h")) t.h = +at(q, "h");
        if (at(q, "fill") && at(q, "fill") !== "norm") t.f = at(q, "fill");
        if (at(q, "stroke") === "0" || at(q, "stroke") === "false") t.s = 0;
        hs(q).forEach((c) => {
          const n = c.localName;
          if (n === "moveTo") t.c.push(["M", ...pts(c)]);
          else if (n === "lnTo") t.c.push(["L", ...pts(c)]);
          else if (n === "arcTo") t.c.push(["A", ...["wR", "hR", "stAng", "swAng"].map((k) => num(at(c, k)))]);
          else if (n === "quadBezTo") t.c.push(["Q", ...pts(c)]);
          else if (n === "cubicBezTo") t.c.push(["C", ...pts(c)]);
          else if (n === "close") t.c.push(["Z"]);
        });
        if (t.c.length) def.p.push(t);
      });
      return def.p.length ? def : null;
    }
    /** Relleno de una forma que no es un color liso: degradado (lineal o radial) o imagen (con su duotono). */
    async function pinturaDe(spPr, relsX) {
      const bf = spPr && h1(spPr, "blipFill"), gf = spPr && h1(spPr, "gradFill"), pf = spPr && h1(spPr, "pattFill");
      if (pf) { const f = await fondoDeRelleno(pf, null, relsX); return f ? { tipo: "trama", prst: f.trama, fg: f.linea, bg: f.color, pp: f.pp, mezcla: f.mezcla } : null; }
      if (bf) { const f = await fondoDeRelleno(bf, null, relsX); return f && f.src ? { tipo: "img", src: f.src } : null; }
      if (!gf) return null;
      const paradas = todos(gf, "gs").map((g) => [Math.round((+(at(g, "pos") || 0) / 100000) * 10000) / 10000, colorCss(g)]).filter((q) => q[1]).sort((a, b) => a[0] - b[0]);
      if (paradas.length < 2) return null;
      const lin = h1(gf, "lin"), path = h1(gf, "path");
      if (path) {
        const fr = h1(path, "fillToRect"), v = (k) => (fr ? +(at(fr, k) || 0) / 100000 : 0.5);
        return { tipo: "rad", paradas, cx: Math.round(((v("l") + 1 - v("r")) / 2) * 1000) / 1000, cy: Math.round(((v("t") + 1 - v("b")) / 2) * 1000) / 1000 };
      }
      return { tipo: "lin", paradas, ang: lin ? Math.round((+(at(lin, "ang") || 0) / 60000) * 100) / 100 : 90 };
    }
    /** Un relleno de DrawingML (sólido, degradado, imagen, trama) → fondo de NOM (con su CSS exacto). */
    async function fondoDeRelleno(f, ph, relsX) {
      if (!f) return null;
      const nm = f.localName;
      if (nm === "solidFill") { const c = colorCss(f, ph); return c ? { tipo: "color", color: c } : null; }
      if (nm === "gradFill") {
        const gs = todos(f, "gs").map((g) => ({ p: +(at(g, "pos") || 0) / 1000, c: colorCss(g, ph) })).filter((g) => g.c).sort((a, b) => a.p - b.p);
        if (!gs.length) return null;
        const paradas = gs.map((g) => `${g.c} ${Math.round(g.p * 100) / 100}%`).join(", ");
        const lin = h1(f, "lin"), path = h1(f, "path");
        let css;
        if (path) {
          const fr = h1(path, "fillToRect"), cx = fr ? 50 + ((+(at(fr, "l") || 0)) - (+(at(fr, "r") || 0))) / 2000 : 50, cy = fr ? 50 + ((+(at(fr, "t") || 0)) - (+(at(fr, "b") || 0))) / 2000 : 50;
          css = `radial-gradient(${at(path, "path") === "circle" ? "circle" : "ellipse"} farthest-corner at ${Math.round(cx)}% ${Math.round(cy)}%, ${paradas})`;
        } else css = `linear-gradient(${Math.round((90 + (lin ? +(at(lin, "ang") || 0) / 60000 : 90)) * 100) / 100}deg, ${paradas})`;
        return { tipo: "degradado", c1: gs[0].c, c2: gs[gs.length - 1].c, angulo: 135, css };
      }
      if (nm === "blipFill") {
        // (una imagen de un estilo del tema se busca en las relaciones del TEMA: el mismo rId del patrón es otra cosa)
        const b2 = todos(f, "blip")[0], rel = b2 && ((f.__tema && relsTema[rid(b2, "embed")]) || relsX[rid(b2, "embed")]);
        let src = rel ? await dataUrl(rel.destino) : null;
        if (!src) return null;
        const duo = b2 && h1(b2, "duotone");
        if (duo) { const cs = hs(duo).map((ce) => colorDeElemento(ce, ph)).filter(Boolean); if (cs.length === 2) src = await duotono(src, cs[0], cs[1]); }
        const tile = h1(f, "tile");
        if (!tile) return { tipo: "imagen", src, ajuste: "estirar" };
        // mosaico: cada pieza mide lo que la imagen a su resolución (96 ppp si no la dice), por la escala del archivo
        const md = await medidasImagen(src), dpi = dpiImagen(src) || 96, k = (96 / dpi) * kx;
        const sx = +(at(tile, "sx") || 100000) / 100000, sy = +(at(tile, "sy") || 100000) / 100000;
        const r2 = (v) => Math.round(v * 100) / 100;
        return { tipo: "imagen", src, ajuste: "mosaico", tw: r2(md.w * k * sx), th: r2(md.h * k * sy), tx: r2((+(at(tile, "tx") || 0) / EMU) * kx), ty: r2((+(at(tile, "ty") || 0) / EMU) * ky) };
      }
      if (nm === "pattFill") {
        const fg = h1(f, "fgClr"), bg = h1(f, "bgClr"), cf = fg ? hexDe(fg, ph) : "#000000", cb = bg ? hexDe(bg, ph) : "#FFFFFF";
        // la trama exacta (y el color parejo que se ve de lejos, por si no se puede dibujar)
        return { tipo: "trama", trama: at(f, "prst") || "pct50", linea: cf, color: cb, mezcla: colorTrama(at(f, "prst") || "pct50", cf, cb), pp: Math.round(kx * 10000) / 10000 };
      }
      return null;
    }
    const medios = new Map();
    async function dataUrl(ruta) {
      if (medios.has(ruta)) return medios.get(ruta);
      const b = await z.leer(ruta);
      const ext = (ruta.split(".").pop() || "").toLowerCase();
      const mime = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp", svg: "image/svg+xml", bmp: "image/bmp", emf: "image/x-emf", wmf: "image/x-wmf" }[ext];
      let u = null;
      if (b && mime) { let bin = ""; for (let i = 0; i < b.length; i += 0x8000) bin += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); u = `data:${mime};base64,${btoa(bin)}`; }
      // Las fotos grandes se achican (máx. 1600 px) para que la presentación no pese de más.
      if (u && b.length > 350000 && /png|jpeg|webp/.test(mime)) u = await achicar(u, mime);
      medios.set(ruta, u);
      return u;
    }
    function achicar(u, mime) {
      return new Promise((ok) => {
        const img = new Image();
        img.onload = () => {
          const k = Math.min(1, 1600 / Math.max(img.width, img.height));
          if (k >= 1 && mime !== "image/png") { ok(u); return; }
          const cv = document.createElement("canvas");
          cv.width = Math.max(1, Math.round(img.width * k)); cv.height = Math.max(1, Math.round(img.height * k));
          cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
          const r = cv.toDataURL(mime === "image/png" ? "image/png" : "image/jpeg", 0.85);
          ok(r.length < u.length ? r : u);
        };
        img.onerror = () => ok(u);
        img.src = u;
      });
    }
    /** Posición de un elemento (xfrm propio o el de su marcador en el diseño o el patrón). */
    /** Ajustes de una geometría de Office (<a:avLst><a:gd name="adj" fmla="val 25000"/>): { adj: 25000 }. */
    function ajustesDe(spPr) {
      const o = {};
      todos(todos(spPr, "prstGeom")[0], "gd").forEach((g) => { const m = /^val\s+(-?[\d.]+)/.exec(g.getAttribute("fmla") || ""); if (m && g.getAttribute("name")) o[g.getAttribute("name")] = +m[1]; });
      return Object.keys(o).length ? o : undefined;
    }
    function xfrmDe(spPr, ph, heredados) {
      const xf = spPr && h1(spPr, "xfrm");
      const leer = (x) => { const o = h1(x, "off"), e = h1(x, "ext"); return o && e ? { x: +at(o, "x"), y: +at(o, "y"), w: +at(e, "cx"), h: +at(e, "cy"), rot: +(at(x, "rot") || 0), fh: at(x, "flipH") === "1", fv: at(x, "flipV") === "1" } : null; };
      let r = xf ? leer(xf) : null;
      if (!r && ph) r = [heredados["#" + (ph.idx || "")], heredados[ph.clave], heredados[ph.tipo]].find((v) => v && !v.sinPos && v.x != null) || null;
      return r;
    }
    async function marcadores(ruta) {
      const m = {};
      const t = await z.texto(ruta);
      if (!t) return m;
      const dm = P(t);
      todos(dm, "sp").forEach((sp) => {
        const ph = todos(sp, "ph")[0], xf = todos(sp, "xfrm")[0];
        if (!ph) return;
        const o = xf && h1(xf, "off"), e = xf && h1(xf, "ext");
        const v = o && e ? { x: +at(o, "x"), y: +at(o, "y"), w: +at(e, "cx"), h: +at(e, "cy"), rot: 0 } : { sinPos: true };
        const tb = todos(sp, "txBody")[0];
        v.lst = tb ? h1(tb, "lstStyle") : null; v.cuerpo = tb ? h1(tb, "bodyPr") : null;
        const tipo = at(ph, "type") || "body", idx = at(ph, "idx") || "";
        m[tipo + ":" + idx] = v;
        if (!m[tipo]) m[tipo] = v;
        if (idx) m["#" + idx] = v;   // la diapositiva se liga al marcador del diseño por su índice
      });
      // estilos de texto del patrón (títulos, cuerpo y lo demás)
      const ts = todos(dm, "txStyles")[0];
      if (ts) m.__estilos = { title: h1(ts, "titleStyle"), body: h1(ts, "bodyStyle"), other: h1(ts, "otherStyle") };
      return m;
    }
    // ---------------------------------------------------------------- texto como PowerPoint (16-nov-2026)
    // Cada propiedad se hereda como en PowerPoint: estilo de la presentación → estilo del patrón (títulos, cuerpo
    // u otros) → marcador del patrón → marcador del diseño → lista del cuadro → párrafo → tramo. El renglón mide
    // 1.2 × el tamaño × el interlineado; los márgenes internos, el anclaje y el autoajuste vienen del cuerpo.
    const kPt = Math.min(kx, ky) * (96 / 72);   // px de la diapositiva por punto (1 en 16:9)
    const defPres = todos(docP, "defaultTextStyle")[0] || null;
    const nivelDe = (lst, l) => (lst ? h1(lst, "lvl" + l + "pPr") : null);
    const WINGDINGS3 = { 0xF084: "▸", 0xF07D: "▴", 0xF07E: "▾", 0xF07C: "◂", 0xF070: "▲", 0xF071: "▼", 0xF075: "◄", 0xF076: "►", 0xF0C6: "⇨", 0xF0C7: "⇦", 0xF022: "➢" };
    const WINGDINGS = { 0xF0A7: "▪", 0xF0D8: "➢", 0xF0FC: "✓", 0xF0E8: "➔", 0xF06C: "●", 0xF06E: "■", 0xF071: "❑", 0xF075: "◆", 0xF076: "❖", 0xF0A8: "◆", 0xF0B7: "•", 0xF0AB: "★", 0xF0D7: "⮚", 0xF0F0: "⇨", 0xF0FB: "✗", 0xF0FD: "☒", 0xF0FE: "☑", 0xF0A1: "○", 0xF09F: "•", 0xF0B2: "◇" };
    function leerRPr(el, r) {
      if (!el) return r;
      if (at(el, "sz")) r.sz = +at(el, "sz") / 100;
      if (at(el, "b") != null) r.b = at(el, "b") === "1" || at(el, "b") === "true";
      if (at(el, "i") != null) r.i = at(el, "i") === "1" || at(el, "i") === "true";
      if (at(el, "u") != null) r.u = at(el, "u");
      if (at(el, "strike") != null) r.strike = at(el, "strike");
      if (at(el, "cap") != null) r.cap = at(el, "cap");
      if (at(el, "spc") != null) r.spc = +at(el, "spc") / 100;
      if (at(el, "baseline") != null) r.baseline = +at(el, "baseline") / 1000;
      const lat = h1(el, "latin"); if (lat && at(lat, "typeface")) r.fuente = at(lat, "typeface");
      const sf = h1(el, "solidFill") || h1(el, "gradFill"); if (sf) { const c = hexDe(sf); if (c) r.color = c; }
      if (h1(el, "noFill")) r.color = "transparent";
      const hl = h1(el, "highlight"); if (hl) { const c = hexDe(hl); if (c) r.resalte = c; }
      return r;
    }
    function leerPPr(el, o) {
      if (!el) return o;
      ["algn", "marL", "indent", "rtl"].forEach((k) => { if (at(el, k) != null) o[k] = at(el, k); });
      const espacio = (n) => { const c = n && hs(n)[0]; if (!c) return null; return c.localName === "spcPts" ? { pts: +at(c, "val") / 100 } : { pct: +at(c, "val") / 100000 }; };
      const ls = espacio(h1(el, "lnSpc")); if (ls) o.lnSpc = ls;
      const sb = espacio(h1(el, "spcBef")); if (sb) o.antes = sb;
      const sa = espacio(h1(el, "spcAft")); if (sa) o.despues = sa;
      if (h1(el, "buNone")) o.bu = { nada: true };
      if (h1(el, "buChar")) o.bu = { car: at(h1(el, "buChar"), "char") || "•" };
      if (h1(el, "buAutoNum")) o.bu = { auto: at(h1(el, "buAutoNum"), "type") || "arabicPeriod", inicio: +(at(h1(el, "buAutoNum"), "startAt") || 1) };
      if (h1(el, "buFont")) o.buFuente = at(h1(el, "buFont"), "typeface");
      if (h1(el, "buFontTx")) delete o.buFuente;
      if (h1(el, "buClr")) o.buColor = hexDe(h1(el, "buClr"));
      if (h1(el, "buClrTx")) delete o.buColor;
      if (h1(el, "buSzPct")) o.buTam = { pct: +at(h1(el, "buSzPct"), "val") / 100000 };
      if (h1(el, "buSzPts")) o.buTam = { pts: +at(h1(el, "buSzPts"), "val") / 100 };
      if (h1(el, "buSzTx")) delete o.buTam;
      o.r = leerRPr(h1(el, "defRPr"), Object.assign({}, o.r));
      return o;
    }
    const resolverFuente = (f) => (f === "+mj-lt" ? fuentesTema.mayor : f === "+mn-lt" ? fuentesTema.menor : f && f[0] === "+" ? fuentesTema.menor : f) || fuentesTema.menor || "Calibri";
    function numeroAuto(tipo, n) {
      const romano = (x) => { let r = ""; for (const [v, t] of [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]]) while (x >= v) { r += t; x -= v; } return r; };
      const letra = (x) => { let r = ""; while (x > 0) { r = String.fromCharCode(97 + ((x - 1) % 26)) + r; x = Math.floor((x - 1) / 26); } return r; };
      const t = String(tipo);
      let v = /^arabic/.test(t) ? String(n) : /^alphaLc/.test(t) ? letra(n) : /^alphaUc/.test(t) ? letra(n).toUpperCase() : /^romanLc/.test(t) ? romano(n) : /^romanUc/.test(t) ? romano(n).toUpperCase() : String(n);
      return /ParenBoth$/.test(t) ? `(${v})` : /ParenR$/.test(t) ? `${v})` : /Period$/.test(t) ? `${v}.` : /Minus$/.test(t) ? `${v} -` : v;
    }
    /**
     * Texto de un cuadro de PowerPoint → { html, cuerpo }. ctx: { tipoPh (title, body…, null si no es marcador),
     * claves (del marcador en el diseño y el patrón), diseno, maestro (los marcadores), enForma }.
     */
    function textoPpt(txBody, ctx) {
      const estilos = (ctx.maestro && ctx.maestro.__estilos) || {};
      const tipo = ctx.tipoPh;
      const est = !tipo ? null : /^(title|ctrTitle)$/.test(tipo) ? estilos.title : /^(body|subTitle|obj)$/.test(tipo) || tipo === "_idx" ? estilos.body : estilos.other;
      const phM = ctx.phMaestro || null, phD = ctx.phDiseno || null;
      // cuerpo: márgenes internos, anclaje, ajuste, autoajuste, columnas, texto vertical
      const cuerpo = { l: 91440, t: 45720, r: 91440, b: 45720, anc: "t", wrap: true, escala: 1, reduce: 0, cols: 1, vert: null };
      [phM && phM.cuerpo, phD && phD.cuerpo, h1(txBody, "bodyPr")].forEach((bp) => {
        if (!bp) return;
        [["lIns", "l"], ["tIns", "t"], ["rIns", "r"], ["bIns", "b"]].forEach(([a, k]) => { if (at(bp, a) != null) cuerpo[k] = +at(bp, a); });
        if (at(bp, "anchor")) cuerpo.anc = at(bp, "anchor");
        if (at(bp, "wrap")) cuerpo.wrap = at(bp, "wrap") !== "none";
        if (at(bp, "vert")) cuerpo.vert = at(bp, "vert") === "horz" ? null : at(bp, "vert");
        if (at(bp, "numCol")) cuerpo.cols = +at(bp, "numCol");
        if (at(bp, "spcCol")) cuerpo.espCol = +at(bp, "spcCol");
        const na = h1(bp, "normAutofit");
        if (na) { cuerpo.escala = at(na, "fontScale") ? +at(na, "fontScale") / 100000 : 1; cuerpo.reduce = at(na, "lnSpcReduction") ? +at(na, "lnSpcReduction") / 100000 : 0; }
        if (h1(bp, "noAutofit")) { cuerpo.escala = 1; cuerpo.reduce = 0; }
      });
      const px = (pt) => Math.round(pt * kPt * 100) / 100, emu = (v) => Math.round((v / 12700) * kPt * 100) / 100;
      const contadores = {};
      let h = "", primero = true;
      hs(txBody, "p").forEach((p) => {
        const pPr = h1(p, "pPr"), nivel = Math.min(9, (+(at(pPr, "lvl") || 0)) + 1);
        const o = { r: {} };
        [defPres && nivelDe(defPres, nivel), est && nivelDe(est, nivel), phM && phM.lst && nivelDe(phM.lst, nivel), phD && phD.lst && nivelDe(phD.lst, nivel)].forEach((el) => leerPPr(el, o));
        if (ctx.r0) Object.assign(o.r, ctx.r0);   // (texto de una celda: el color y la negrita del estilo de la tabla)
        [nivelDe(h1(txBody, "lstStyle"), nivel), pPr].forEach((el) => leerPPr(el, o));
        // tramos
        const tramos = [];
        hs(p).forEach((r) => {
          if (r.localName === "br") { tramos.push({ br: true, r: leerRPr(h1(r, "rPr"), Object.assign({}, o.r)) }); return; }
          if (r.localName !== "r" && r.localName !== "fld") return;
          const t = (h1(r, "t") || {}).textContent || "";
          if (!t) return;
          tramos.push({ t, r: leerRPr(h1(r, "rPr"), Object.assign({}, o.r)) });
        });
        const finR = leerRPr(h1(p, "endParaRPr"), Object.assign({}, o.r));
        const r0 = (tramos.find((x) => !x.br) || {}).r || finR;
        const tam0 = (r0.sz || 18) * cuerpo.escala;
        const ls = o.lnSpc || { pct: 1 };
        // (el interlineado exacto PowerPoint lo dibuja redondeado a puntos enteros: 33.6 → 34, medido)
        const lh = ls.pts ? `${px(Math.round(ls.pts))}px` : String(Math.round(1.2 * Math.max(0.1, ls.pct - cuerpo.reduce) * 10000) / 10000);
        const espacio = (e) => (!e ? 0 : e.pts ? px(e.pts) : px(e.pct * tam0 * 1.2));
        // (con la letra de su primer texto: si el párrafo heredara otra, el navegador alinea las dos por la línea base y el
        // renglón crece más que el interlineado —medido contra PowerPoint: hasta 3 pt por renglón en Segoe UI 44—)
        const st = [`font-size:${px(tam0)}px`, `line-height:${lh}`, `font-family:${NOM.textoRico ? NOM.textoRico.cssFuente(resolverFuente(r0.fuente)) : resolverFuente(r0.fuente)}`];
        if (o.algn) st.push(`text-align:${{ ctr: "center", r: "right", just: "justify", dist: "justify", l: "left" }[o.algn] || "left"}`);
        const marL = o.marL != null ? emu(+o.marL) : 0, ind = o.indent != null ? emu(+o.indent) : 0;
        if (marL) st.push(`margin-left:${marL}px`);
        if (ind) st.push(`text-indent:${ind}px`);
        // (PowerPoint no pone el espacio de antes en el primer párrafo del cuadro)
        const mt = primero ? 0 : espacio(o.antes), mb = espacio(o.despues);
        primero = false;
        if (mt) st.push(`margin-top:${mt}px`);
        if (mb) st.push(`margin-bottom:${mb}px`);
        // viñeta
        let vin = "";
        const conTexto = tramos.some((x) => !x.br && x.t.trim());
        if (o.bu && !o.bu.nada && conTexto) {
          let car;
          if (o.bu.auto) {
            const k = nivel + ":" + o.bu.auto;
            Object.keys(contadores).forEach((q) => { if (+q.split(":")[0] > nivel) delete contadores[q]; });
            contadores[k] = (contadores[k] || o.bu.inicio - 1) + 1;
            car = numeroAuto(o.bu.auto, contadores[k]);
          } else {
            car = o.bu.car || "•";
            const cp = car.codePointAt(0);
            if (/wingdings|symbol|webdings/i.test(o.buFuente || "")) {
              const mapa = /wingdings 3/i.test(o.buFuente) ? WINGDINGS3 : /wingdings 2/i.test(o.buFuente) ? {} : WINGDINGS;
              car = mapa[cp] || mapa[cp + 0xF000] || (cp >= 0xF000 || /wingdings [23]|webdings/i.test(o.buFuente) ? "•" : car);
            }
          }
          const tamV = o.buTam ? (o.buTam.pts ? o.buTam.pts : tam0 * o.buTam.pct) : tam0;
          const fuenteV = o.bu.auto ? r0.fuente : /wingdings|symbol/i.test(o.buFuente || "") ? r0.fuente : o.buFuente || r0.fuente;
          const ancho = ind < 0 ? -ind : 0;
          vin = `<span style="display:inline-block;width:${ancho}px;text-indent:0px;font-size:${px(tamV)}px;color:${esc2(o.buColor || r0.color || "")};font-family:${esc2(NOM.textoRico ? NOM.textoRico.cssFuente(resolverFuente(fuenteV)) : resolverFuente(fuenteV))};font-weight:400;font-style:normal">${NOM.esc(car)}</span>`;
          if (!ancho) vin += " ";
        } else if (!o.bu || o.bu.nada || !conTexto) Object.keys(contadores).forEach((q) => { if (+q.split(":")[0] >= nivel) delete contadores[q]; });
        let cuerpoP = "";
        tramos.forEach((x) => {
          if (x.br) { cuerpoP += "<br>"; return; }
          const r = x.r, s2 = [];
          const tam = (r.sz || 18) * cuerpo.escala;
          s2.push(`font-size:${px(r.baseline ? tam * 0.667 : tam)}px`);
          s2.push(`font-family:${NOM.textoRico ? NOM.textoRico.cssFuente(resolverFuente(r.fuente)) : resolverFuente(r.fuente)}`);
          if (r.color) s2.push(`color:${r.color}`);
          if (r.resalte) s2.push(`background-color:${r.resalte}`);
          s2.push(`font-weight:${r.b ? 700 : 400}`, `font-style:${r.i ? "italic" : "normal"}`);
          const deco = [r.u && r.u !== "none" ? "underline" : "", r.strike && r.strike !== "noStrike" ? "line-through" : ""].filter(Boolean).join(" ");
          if (deco) s2.push(`text-decoration:${deco}`);
          if (r.spc) s2.push(`letter-spacing:${px(r.spc)}px`);
          if (r.cap === "all") s2.push("text-transform:uppercase"); else if (r.cap === "small") s2.push("font-variant-caps:small-caps");
          if (r.baseline) s2.push(`vertical-align:${r.baseline > 0 ? "super" : "sub"}`);
          s2.push(`line-height:${lh}`);
          cuerpoP += `<span style="${s2.join(";")}">${NOM.esc(x.t)}</span>`;
        });
        if (!cuerpoP) cuerpoP = `<span style="font-size:${px((finR.sz || 18) * cuerpo.escala)}px;line-height:${lh}"><br></span>`;
        h += `<div style="${st.join(";")}">${vin}${cuerpoP}</div>`;
      });
      return { html: h, cuerpo: { ins: [emu(cuerpo.l), emu(cuerpo.t), emu(cuerpo.r), emu(cuerpo.b)], anc: cuerpo.anc, ajusta: cuerpo.wrap, cols: cuerpo.cols > 1 ? cuerpo.cols : undefined, espCol: cuerpo.espCol ? emu(cuerpo.espCol) : undefined, vert: cuerpo.vert || undefined } };
    }
    const esc2 = (v) => String(v || "").replace(/[;"<>]/g, "");
    /** Párrafos de PowerPoint → HTML con formato (negrita, color, tamaño, viñetas). */
    function htmlTexto(txBody, base) {
      let h = "", lista = null;
      const cerrar = () => { if (lista) { h += `</${lista}>`; lista = null; } };
      hs(txBody, "p").forEach((p) => {
        const pPr = h1(p, "pPr");
        const algn = at(pPr, "algn");
        const vin = pPr && (h1(pPr, "buChar") ? "ul" : h1(pPr, "buAutoNum") ? "ol" : null);
        const lst = vin || (base.lista && !(pPr && h1(pPr, "buNone")) ? "ul" : null);
        let runs = "";
        hs(p).forEach((r) => {
          if (r.localName === "br") { runs += "<br>"; return; }
          if (r.localName !== "r" && r.localName !== "fld") return;
          const t = (h1(r, "t") || {}).textContent || "";
          if (!t) return;
          const rp = h1(r, "rPr"), st = [];
          if (at(rp, "sz")) st.push(`font-size:${Math.round((+at(rp, "sz") / 100) * 1.333 * Math.min(kx, ky))}px`);
          const c = rp && hexDe(h1(rp, "solidFill"));
          if (c) st.push(`color:${c}`);
          let x2 = NOM.esc(t);
          if (st.length) x2 = `<span style="${st.join(";")}">${x2}</span>`;
          if (at(rp, "b") === "1") x2 = `<b>${x2}</b>`;
          if (at(rp, "i") === "1") x2 = `<i>${x2}</i>`;
          if (at(rp, "u") && at(rp, "u") !== "none") x2 = `<u>${x2}</u>`;
          if (at(rp, "strike") && at(rp, "strike") !== "noStrike") x2 = `<s>${x2}</s>`;
          runs += x2;
        });
        const al = algn ? ` style="text-align:${{ ctr: "center", r: "right", just: "justify" }[algn] || "left"}"` : "";
        if (lst) { if (lista !== lst) { cerrar(); h += `<${lst}>`; lista = lst; } h += `<li${al}>${runs || "<br>"}</li>`; }
        else { cerrar(); h += `<div${al}>${runs || "<br>"}</div>`; }
      });
      cerrar();
      return h;
    }
    const bloques = [], paginas = [], docCampos = { tamDiapo };
    // contraseña para modificar (p:modifyVerifier) y marcada como final: se abre para ver hasta que se escriba o se decida editar
    {
      const mv = todos(P(pres), "modifyVerifier")[0];
      if (mv) docCampos.modificar = { p: proteccionDe(mv, false) || {} };
      if (await marcadoFinal(z)) docCampos.final = true;
    }
    let aproximado = false, n = 0;
    for (const s of ids) {
      const ruta = s.destino, xml = await z.texto(ruta);
      if (!xml) continue;
      const pg = "p" + ++n;
      paginas.push(pg);
      let rels = await relsDe(ruta);
      const sld = P(xml);
      // Presentación → Ocultar diapositiva (<p:sld show="0">): no sale al presentar
      if (/^(0|false)$/.test(sld.documentElement.getAttribute("show") || "")) docCampos["oculta:" + pg] = true;
      // Marcadores del diseño y del patrón (para lo que no trae posición propia).
      const relDiseno = Object.values(rels).find((r) => /slideLayout$/.test(r.tipo));
      let heredados = {}, mDis = {}, mMae = {};
      if (relDiseno) {
        const relsDis = await relsDe(relDiseno.destino);
        const relMaestro = Object.values(relsDis).find((r) => /slideMaster$/.test(r.tipo));
        // el tema de ESTE patrón (antes de leer colores y fuentes)
        const relTemaM = relMaestro ? Object.values(await relsDe(relMaestro.destino)).find((r) => /\/theme$/.test(r.tipo)) : null;
        usarTema(relTemaM ? await leerTema(relTemaM.destino) : temaPres);
        mMae = relMaestro ? await marcadores(relMaestro.destino) : {};
        mDis = await marcadores(relDiseno.destino);
        // mapa de colores: el patrón dice qué es «texto 1» y «fondo 1» (en un tema oscuro, el texto es el claro)
        const leeMapa = (t) => { const m2 = t && /<(?:\w+:)?(?:clrMap|overrideClrMapping)\s([^>]*)>/.exec(t); const o2 = {}; if (m2) m2[1].replace(/(\w+)="(\w+)"/g, (_, k, v) => { o2[k] = v; }); return o2; };
        let mapa = relMaestro ? leeMapa(await z.texto(relMaestro.destino)) : {};
        const ovD = leeMapa(/<p:clrMapOvr>\s*<a:overrideClrMapping[^>]*>/.exec(await z.texto(relDiseno.destino) || "") ? (/<p:clrMapOvr>\s*(<a:overrideClrMapping[^>]*>)/.exec(await z.texto(relDiseno.destino)) || [])[1] : "");
        const ovS = leeMapa((/<p:clrMapOvr>\s*(<a:overrideClrMapping[^>]*>)/.exec(xml) || [])[1] || "");
        mapa = Object.assign({}, mapa, ovD, ovS);
        ["bg1", "tx1", "bg2", "tx2"].forEach((k) => { const v = mapa[k] || { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" }[k]; if (esquema[v]) esquema[k] = esquema[v]; });
        heredados = Object.assign({}, mMae);
        for (const k in mDis) if (!mDis[k].sinPos || !heredados[k]) heredados[k] = mDis[k];
      }
      let orden = 0;
      docCampos["os:" + pg] = ruta;   // de qué diapositiva del archivo viene esta página
      let agregar = (b) => { const id = "imp" + n + "_" + orden; bloques.push(Object.assign({ id, pagina: pg, orden: ++orden, z: orden, ox: oxActual }, b)); return id; };
      // estilos de tabla del archivo, con los colores del tema de esta diapositiva
      const estilosTablaPpt = {};
      if (xmlEstTablas) {
        const lnDe = (ln) => { if (!ln) return null; if (h1(ln, "noFill")) return "no"; const c = colorCss(h1(ln, "solidFill")); return c ? [Math.max(0.5, Math.round(((+(at(ln, "w") || 12700)) / 12700) * kPt * 100) / 100), c, null] : null; };
        todos(xmlEstTablas, "tblStyle").forEach((ts) => {
          const e = {};
          hs(ts).forEach((parte) => {
            const o = {}, tx = h1(parte, "tcTxStyle"), tcs = h1(parte, "tcStyle");
            if (tx) { if (at(tx, "b") != null) o.b = at(tx, "b") === "on"; if (at(tx, "i") != null) o.i = at(tx, "i") === "on"; const c = hs(tx).find((x) => /Clr$/.test(x.localName)); if (c) { const w = document.createElementNS(c.namespaceURI, "a:solidFill"); w.appendChild(c.cloneNode(true)); o.color = colorCss(w); } }
            const fi = tcs && h1(tcs, "fill"); if (fi) { const sf = h1(fi, "solidFill"); if (sf) o.fondo = colorCss(sf); else if (h1(fi, "noFill")) o.fondo = null; }
            const tb = tcs && h1(tcs, "tcBdr");
            if (tb) { const bd = {}; ["left", "right", "top", "bottom", "insideH", "insideV"].forEach((n) => { const x = h1(tb, n), v = x && lnDe(h1(x, "ln")); if (v) bd[n] = v; }); if (Object.keys(bd).length) o.bd = bd; }
            e[parte.localName] = o;
          });
          if (at(ts, "styleId")) estilosTablaPpt[at(ts, "styleId")] = e;
        });
      }
      /**
       * SmartArt. Si el archivo trae al lado el dibujo ya acomodado (PowerPoint 2010 en adelante lo
       * guarda), se usan esas figuras con su lugar, su color y su texto. Si no, se arma con su lista
       * de cuadros: en árbol (organigrama) si unos cuelgan de otros, o en fila si es una lista.
       * Todas llevan el mismo "ox": mientras no se toquen, al exportar sale el SmartArt original.
       */
      async function smartArt(dg, r, anchoOrig, altoOrig) {
        const relDatos = rels[rid(dg, "dm")], datos = relDatos ? await z.texto(relDatos.destino) : null;
        if (!datos) return false;
        const m = /dataModelExt[^>]*relId="([^"]+)"/.exec(datos), relDib = m ? rels[m[1]] : null;
        const dib = relDib ? await z.texto(relDib.destino) : null;
        let hechas = 0;
        if (dib && typeof document !== "undefined") {
          // Cada figura del dibujo guardado pasa por el mismo camino que las formas de la diapositiva (geometría de Office
          // con sus ajustes, formas libres, degradados, contornos, texto con su herencia), en su lugar dentro del marco.
          const fx = anchoOrig ? r.w / anchoOrig : 1, fy = altoOrig ? r.h / altoOrig : 1;
          const grupoSA = { tx: (v) => r.x + v * fx, ty: (v) => r.y + v * fy, s: (v) => v * fx, sy: (v) => v * fy };
          // El color del texto como lo pinta PowerPoint: el de la etiqueta de estilo de cada figura (su punto de presentación
          // en los datos, por modelId) en los colores del SmartArt (txFillClrLst), no el fontRef del dibujo guardado
          // (que a veces dice blanco sobre un nodo blanco).
          const etiquetas = {};
          todos(P(datos), "pt").forEach((q) => { const ps = h1(q, "prSet"); if (ps && at(ps, "presStyleLbl")) etiquetas[at(q, "modelId")] = { lbl: at(ps, "presStyleLbl"), i: +(at(ps, "presStyleIdx") || 0), n: +(at(ps, "presStyleCnt") || 1) }; });
          const relCol = rels[rid(dg, "cs")], tCol = relCol ? await z.texto(relCol.destino) : null, xCol = tCol ? P(tCol) : null;
          const tintaDe = (mid) => {
            const e = etiquetas[mid];
            if (!e || !xCol) return null;
            const sl = todos(xCol, "styleLbl").find((s) => at(s, "name") === e.lbl), lst = sl && h1(sl, "txFillClrLst"), cs = lst ? hs(lst) : [];
            if (!cs.length) return null;
            const k = at(lst, "meth") === "span" && cs.length > 1 ? Math.round((e.i / Math.max(1, e.n - 1)) * (cs.length - 1)) : e.i % cs.length;
            return cs[Math.min(k, cs.length - 1)];
          };
          const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
          const conColor = (txBody, tc) => {
            // cada corrida sin color propio lleva el de su etiqueta (y la marca de fin de párrafo)
            for (const rp of [...todos(txBody, "rPr"), ...todos(txBody, "endParaRPr"), ...todos(txBody, "defRPr")]) {
              if (hs(rp).some((x) => /^(solidFill|gradFill|noFill|pattFill|blipFill)$/.test(x.localName))) continue;
              const sf = rp.ownerDocument.createElementNS(A, "a:solidFill");
              sf.appendChild(tc.cloneNode(true));
              const ln0 = hs(rp).find((x) => x.localName === "ln");
              rp.insertBefore(sf, ln0 ? ln0.nextSibling : rp.firstChild);
            }
            for (const run of todos(txBody, "r")) if (!h1(run, "rPr")) { const rp = run.ownerDocument.createElementNS(A, "a:rPr"); run.insertBefore(rp, run.firstChild); const sf = rp.ownerDocument.createElementNS(A, "a:solidFill"); sf.appendChild(tc.cloneNode(true)); rp.appendChild(sf); }
          };
          for (const sp of todos(P(dib), "sp")) {
            const spPr = h1(sp, "spPr"), xf = spPr && h1(spPr, "xfrm"), o = xf && h1(xf, "off"), e = xf && h1(xf, "ext");
            if (!o || !e || (!+at(e, "cx") && !+at(e, "cy"))) continue;
            // (si los colores del SmartArt no dicen el del texto, PowerPoint usa el del estilo de la figura: su fontRef)
            const fRef = todos(h1(sp, "style"), "fontRef")[0];
            const tc = tintaDe(at(sp, "modelId")) || (fRef ? hs(fRef)[0] || null : null);
            const txb = h1(sp, "txBody"), conTexto = !!(txb && txb.textContent.trim());
            const txX = h1(sp, "txXfrm");
            // (cada figura recuerda su punto del SmartArt: las animaciones «uno por uno» o «por nivel» lo señalan con
            // <p:graphicEl><a:dgm id="modelId"/>)
            const k0 = bloques.length, mid = at(sp, "modelId");
            // la figura (sin su texto si el texto va en su propio rectángulo)
            const fig = sp.cloneNode(true);
            if (conTexto && txX) { const t = h1(fig, "txBody"); if (t) t.parentNode.removeChild(t); }
            else if (conTexto && tc) conColor(h1(fig, "txBody"), tc);
            await elemento(fig, grupoSA);
            hechas++;
            if (conTexto && txX) {
              // el texto, derecho, en el rectángulo que PowerPoint le guardó (txXfrm), sin relleno ni contorno
              const t = sp.cloneNode(true), sp2 = h1(t, "spPr");
              for (const x of [...hs(sp2)]) sp2.removeChild(x);
              const xo = h1(txX, "off"), xe = h1(txX, "ext"), D = t.ownerDocument;
              const nx = D.createElementNS(A, "a:xfrm"), no = D.createElementNS(A, "a:off"), ne = D.createElementNS(A, "a:ext");
              no.setAttribute("x", at(xo, "x") || "0"); no.setAttribute("y", at(xo, "y") || "0");
              ne.setAttribute("cx", at(xe, "cx") || "0"); ne.setAttribute("cy", at(xe, "cy") || "0");
              // (sin giro propio, el texto gira con su figura, como en PowerPoint)
              const giroTx = at(txX, "rot") || at(xf, "rot");
              if (giroTx && +giroTx) nx.setAttribute("rot", giroTx);
              nx.appendChild(no); nx.appendChild(ne); sp2.appendChild(nx);
              const pg = D.createElementNS(A, "a:prstGeom"); pg.setAttribute("prst", "rect"); pg.appendChild(D.createElementNS(A, "a:avLst")); sp2.appendChild(pg);
              sp2.appendChild(D.createElementNS(A, "a:noFill"));
              const ln = D.createElementNS(A, "a:ln"); ln.appendChild(D.createElementNS(A, "a:noFill")); sp2.appendChild(ln);
              const st = h1(t, "style"); if (st) st.parentNode.removeChild(st);
              const tb = h1(t, "txBody");
              if (tc && tb) conColor(tb, tc);
              await elemento(t, grupoSA);
            }
            if (mid) for (let k = k0; k < bloques.length; k++) bloques[k].dgm = mid;
          }
          if (hechas) return true;
        }
        if (dib && typeof document === "undefined") {
          const fx = anchoOrig ? r.w / anchoOrig : 1, fy = altoOrig ? r.h / altoOrig : 1;
          todos(P(dib), "sp").forEach((sp) => {
            const spPr = h1(sp, "spPr"), xf = spPr && h1(spPr, "xfrm"), o = xf && h1(xf, "off"), e = xf && h1(xf, "ext");
            if (!o || !e || !+at(e, "cx") || !+at(e, "cy")) return;
            const txBody = h1(sp, "txBody"), texto = txBody ? txBody.textContent.trim() : "";
            const est = h1(sp, "style"), sinRelleno = !!h1(spPr, "noFill");
            const relleno = h1(spPr, "solidFill") ? hexDe(h1(spPr, "solidFill")) : h1(spPr, "gradFill") ? hexDe(h1(spPr, "gradFill")) : sinRelleno ? null : est && +at(h1(est, "fillRef"), "idx") ? hexDe(h1(est, "fillRef")) : null;
            const ln = h1(spPr, "ln"), borde = ln && !h1(ln, "noFill") ? hexDe(h1(ln, "solidFill")) : null;
            const prst = at(h1(spPr, "prstGeom"), "prst");
            if (!prst && !texto) return;   // trazo libre (las líneas que unen los cuadros): no se dibuja
            if (!relleno && !borde && !texto) return;
            const g = geo({ x: r.x + +at(o, "x") * fx, y: r.y + +at(o, "y") * fy, w: +at(e, "cx") * fx, h: +at(e, "cy") * fy, rot: +(at(xf, "rot") || 0) });
            const tinta = est ? hexDe(h1(est, "fontRef")) : null;
            agregar(Object.assign({ tipo: "forma", figura: GEO_INV[prst] || "rect", radio: prst === "roundRect" ? undefined : 0, color: relleno || "#FFFFFF", relleno: relleno ? undefined : false,
              borde: borde || undefined, bordeAncho: borde ? Math.max(1, Math.round(+(at(ln, "w") || 12700) / 12700)) : 0, sombra: false, tColor: tinta && tinta !== "#000000" ? tinta : undefined,
              texto, html: texto ? htmlTexto(txBody, {}) : undefined }, g));
            hechas++;
          });
          if (hechas) return true;
        }
        // ---- sin dibujo guardado: se acomoda con su diseño, sus colores y su estilo, como PowerPoint (js/smartart.js)
        if (NOM.smartart && typeof document !== "undefined") {
          try {
            const parte = async (k) => { const rl = rels[rid(dg, k)]; const t = rl ? await z.texto(rl.destino) : null; return t ? P(t) : null; };
            const partes = { datos: P(datos), diseno: await parte("lo"), colores: await parte("cs"), estilo: await parte("qs") };
            const W = anchoOrig || r.w, H = altoOrig || r.h, fx = anchoOrig ? r.w / anchoOrig : 1, fy = altoOrig ? r.h / altoOrig : 1;
            const lienzoM = document.createElement("canvas").getContext("2d"), fam = fuentesTema.menor || "Calibri";
            const css = NOM.textoRico ? NOM.textoRico.cssFuente(fam) : fam;
            const medir = (t, pt) => { lienzoM.font = `${pt}pt ${css}`; return lienzoM.measureText(t).width * EMU; };
            const formas = partes.diseno ? NOM.smartart.acomodar(partes, { w: W, h: H, color: (el) => colorDeElemento(el), anchoLinea: (i) => (lineasTema[i] || [0, 0.75, 1, 1.5][i] || 1) * 12700, medir }) : null;
            if (formas && formas.length) {
              const nombresAv = (prst) => (((NOM.GEOMETRIAS || {})[prst] || {}).av || []).map((q) => q[0]);
              const kPx = (pt) => Math.round(pt * (96 / 72) * kx * 100) / 100;
              formas.forEach((f) => {
                const prst = NOM.geomOffice && NOM.geomOffice.tiene(f.prst) ? f.prst : "rect";
                // ajustes de la forma: los del diseño son fracciones (0.1 = 10000); los ángulos del arco, en grados
                const av = {}, ns = nombresAv(prst);
                Object.keys(f.adj || {}).forEach((k) => { const n = ns.length === 1 ? ns[0] : ns[+k - 1]; if (n) av[n] = f.arcoGrados ? Math.round(f.adj[k] * 60000) : Math.round(f.adj[k] * 100000); });
                const g = geo({ x: r.x + f.x * fx, y: r.y + f.y * fy, w: f.w * fx, h: f.h * fy, rot: (f.rot || 0) * 60000 });
                const html = f.texto ? f.texto.map((ln) => `<div style="text-align:${ln.nivel ? "left" : { l: "left", r: "right" }[f.alin] || "center"};font-size:${kPx(f.tam * (ln.nivel ? 0.8 : 1))}px;line-height:0.9${ln.nivel ? `;padding-left:${kPx(f.tam * 0.6)}px` : ""}">${ln.nivel ? "•&nbsp;" : ""}${NOM.esc(ln.t) || "&nbsp;"}</div>`).join("") : undefined;
                agregar(Object.assign({ tipo: "forma", figura: GEO_INV[prst] || "rect", prst, av: Object.keys(av).length ? av : undefined, radio: prst === "rect" ? 0 : undefined,
                  color: f.relleno || "#FFFFFF", relleno: f.relleno ? undefined : false, borde: f.linea && f.linea.c ? f.linea.c : undefined,
                  bordeAncho: f.linea && f.linea.c ? Math.max(0.25, Math.round((f.linea.w / 12700) * 100) / 100) : undefined, sombra: false,
                  tColor: f.colorTexto || undefined, texto: f.texto ? f.texto.map((x) => x.t).join("\n") : "", html,
                  ppt: f.texto ? { ins: (f.margen || [0, 0, 0, 0]).map((v) => Math.round((v / EMU) * kx * 100) / 100), anc: { t: "t", mid: "ctr", b: "b" }[f.ancla] || "ctr", ajusta: true } : undefined }, g));
              });
              return true;
            }
          } catch (e) { if (typeof console !== "undefined") console.warn("SmartArt sin dibujo: no se pudo acomodar con su diseño", e); }
        }
        // ---- (último recurso) los cuadros de su lista
        const dm = P(datos), pts = {}, hijosDe = {};
        todos(dm, "pt").forEach((q) => { const tipo = at(q, "type") || "node"; if (tipo === "doc" || tipo === "node" || tipo === "asst") pts[at(q, "modelId")] = { tipo, texto: todos(h1(q, "t"), "p").map((x2) => x2.textContent.trim()).filter(Boolean).join("\n") }; });
        todos(dm, "cxn").filter((c) => (at(c, "type") || "parOf") === "parOf" && pts[at(c, "srcId")] && pts[at(c, "destId")]).sort((a, b) => +at(a, "srcOrd") - +at(b, "srcOrd"))
          .forEach((c) => (hijosDe[at(c, "srcId")] = hijosDe[at(c, "srcId")] || []).push(at(c, "destId")));
        const raizId = Object.keys(pts).find((k) => pts[k].tipo === "doc"), arriba = (raizId && hijosDe[raizId]) || [];
        if (!arriba.length) return false;
        const color = "#" + (esquema.accent1 || "4472C4"), W0 = r.w / EMU * kx, H0 = r.h / EMU * ky, X0 = r.x / EMU * kx, Y0 = r.y / EMU * ky;
        const caja = (id, x, y, w, h) => agregar({ tipo: "forma", figura: "rect", color, borde: "#FFFFFF", bordeAncho: 1, sombra: false, tColor: "#FFFFFF", texto: pts[id].texto, html: `<div style="text-align:center">${NOM.esc(pts[id].texto).replace(/\n/g, "<br>")}</div>`,
          x: Math.round(X0 + x), y: Math.round(Y0 + y), w: Math.max(8, Math.round(w)), h: Math.max(8, Math.round(h)) });
        const enArbol = arriba.some((k) => (hijosDe[k] || []).length);
        if (!enArbol) {
          // lista: en fila si caben, en varias filas si son muchos
          const nC = arriba.length, porFila = nC <= 6 ? nC : Math.ceil(Math.sqrt(nC * (W0 / Math.max(1, H0)))), filas = Math.ceil(nC / porFila);
          const cw = W0 / porFila, ch = H0 / filas;
          arriba.forEach((id, i) => caja(id, (i % porFila) * cw + cw * 0.06, Math.floor(i / porFila) * ch + ch * 0.15, cw * 0.88, Math.min(ch * 0.7, cw * 0.7)));
          return true;
        }
        // árbol: las hojas se reparten a lo ancho y cada cuadro queda centrado sobre los suyos
        let hojas = 0, niveles = 0;
        const pos = {};
        const medir = (id, nivel, visto) => {
          if (visto.has(id) || nivel > 12) return null;
          visto.add(id); niveles = Math.max(niveles, nivel + 1);
          const hj = (hijosDe[id] || []).map((k) => medir(k, nivel + 1, visto)).filter((v) => v != null);
          const x = hj.length ? (hj[0] + hj[hj.length - 1]) / 2 : hojas++ + 0.5;
          pos[id] = { x, nivel };
          return x;
        };
        const visto = new Set();
        arriba.forEach((id) => medir(id, 0, visto));
        const cw = W0 / Math.max(1, hojas), ch = H0 / Math.max(1, niveles), bw = Math.min(cw * 0.86, 220), bh = Math.min(ch * 0.55, 80);
        const idDe = {};
        Object.keys(pos).forEach((id) => { idDe[id] = caja(id, pos[id].x * cw - bw / 2, pos[id].nivel * ch + (ch - bh) / 2, bw, bh); });
        Object.keys(pos).forEach((id) => (hijosDe[id] || []).forEach((k) => { if (idDe[k]) agregar({ tipo: "conector", desde: idDe[id], hasta: idDe[k], orden: 0 }); }));
        return true;
      }
      const geo = (r) => ({ x: Math.round((r.x / EMU) * kx), y: Math.round((r.y / EMU) * ky), w: Math.max(4, Math.round((r.w / EMU) * kx)), h: Math.max(4, Math.round((r.h / EMU) * ky)), rot: r.rot ? Math.round(r.rot / 60000) : undefined, volteoH: r.fh || undefined, volteoV: r.fv || undefined });
      /** Línea o conector de Office tal cual: de esquina a esquina de su caja (con sus volteos), su grosor en puntos,
          su punteado y sus puntas (headEnd al principio, tailEnd al final). */
      function lineaOffice(el, ajustar) {
        const spPr = h1(el, "spPr"), r = ajustar(xfrmDe(spPr, null, heredados));
        if (!r) return;
        const prst = at(todos(spPr, "prstGeom")[0], "prst") || "line";
        const est = h1(el, "style"), lnRef = est && h1(est, "lnRef"), ln = spPr && h1(spPr, "ln");
        if (ln && h1(ln, "noFill")) return;   // sin línea: no se ve
        const color = ln && h1(ln, "solidFill") ? colorCss(h1(ln, "solidFill")) : lnRef ? hexDe(lnRef) : null;
        const anchoPt = ln && at(ln, "w") ? +at(ln, "w") / 12700 : lnRef && +at(lnRef, "idx") ? lineasTema[+at(lnRef, "idx")] || 0.75 : 0.75;
        const guion = ln ? at(h1(ln, "prstDash"), "val") : null;
        const punta = (n) => { const e = ln && h1(ln, n), t = e && at(e, "type"); return t && t !== "none" ? { t, w: at(e, "w") || "med", l: at(e, "len") || "med" } : undefined; };
        const r2 = (v) => Math.round(v * 100) / 100;
        agregar({ tipo: "forma", figura: "linea", lineaOffice: true, prst: NOM.geomOffice && NOM.geomOffice.tiene(prst) ? prst : "line", av: ajustesDe(spPr), sombra: false,
          borde: color || "#000000", bordeAncho: r2(Math.max(0.1, anchoPt * kPt)), bordeEstilo: guion && guion !== "solid" ? (/dot/i.test(guion) && !/dash/i.test(guion) ? "puntos" : "guiones") : undefined,
          puntas: punta("headEnd") || punta("tailEnd") ? { ini: punta("headEnd"), fin: punta("tailEnd") } : undefined,
          x: r2((r.x / EMU) * kx), y: r2((r.y / EMU) * ky), w: r2(Math.max(0, (r.w / EMU) * kx)), h: r2(Math.max(0, (r.h / EMU) * ky)),
          rot: r.rot ? Math.round(r.rot / 60000) : undefined, volteoH: r.fh || undefined, volteoV: r.fv || undefined });
      }
      /** Una tabla de PowerPoint como la dibuja PowerPoint: anchos y altos, celdas combinadas, el estilo de la tabla
          (ppt/tableStyles.xml) por capas —toda la tabla, franjas, primera y última fila y columna, esquinas—, y lo
          propio de cada celda (relleno, bordes, márgenes, anclaje) encima. */
      function tablaComoPpt(tbl, textoCelda) {
        const tblPr = h1(tbl, "tblPr"), si = (a) => at(tblPr, a) === "1" || at(tblPr, a) === "true";
        const fl = { fr: si("firstRow"), fc: si("firstCol"), lr: si("lastRow"), lc: si("lastCol"), br: si("bandRow"), bc: si("bandCol") };
        const idEst = ((tblPr && h1(tblPr, "tableStyleId")) || {}).textContent || "";
        const est = estilosTablaPpt[idEst.trim()] || null;
        const anchos = hs(h1(tbl, "tblGrid"), "gridCol").map((g) => Math.round(((+at(g, "w") || 0) / EMU) * kx * 100) / 100);
        const trs = hs(tbl, "tr"), nF = trs.length, nC = anchos.length;
        const altos = trs.map((tr) => Math.round(((+at(tr, "h") || 0) / EMU) * ky * 100) / 100);
        const borde = (ln) => { if (!ln) return null; if (h1(ln, "noFill")) return "no"; const c = colorCss(h1(ln, "solidFill")); return c ? [Math.max(0.5, Math.round(((+(at(ln, "w") || 12700)) / 12700) * kPt * 100) / 100), c, at(h1(ln, "prstDash"), "val") || null] : null; };
        const celdas = {};
        trs.forEach((tr, f) => hs(tr, "tc").forEach((tc, c) => {
          const k = f + "," + c;
          if (at(tc, "hMerge") === "1" || at(tc, "vMerge") === "1") { celdas[k] = { o: 1 }; return; }
          const cs = Math.max(1, +(at(tc, "gridSpan") || 1)), rs = Math.max(1, +(at(tc, "rowSpan") || 1)), f2 = Math.min(nF - 1, f + rs - 1), c2 = Math.min(nC - 1, c + cs - 1);
          // capas del estilo que tocan esta celda (de menor a mayor precedencia), con la región de cada una
          const capas = [], pon = (n, r) => { if (est && est[n]) capas.push([est[n], r]); };
          pon("wholeTbl", [0, 0, nF - 1, nC - 1]);
          const f0 = fl.fr ? 1 : 0, fN = fl.lr ? nF - 2 : nF - 1, c0 = fl.fc ? 1 : 0, cN = fl.lc ? nC - 2 : nC - 1;
          if (fl.br && f >= f0 && f <= fN) pon((f - f0) % 2 ? "band2H" : "band1H", [f, 0, f2, nC - 1]);
          if (fl.bc && c >= c0 && c <= cN) pon((c - c0) % 2 ? "band2V" : "band1V", [0, c, nF - 1, c2]);
          if (fl.lc && c2 === nC - 1) pon("lastCol", [0, nC - 1, nF - 1, nC - 1]);
          if (fl.fc && c === 0) pon("firstCol", [0, 0, nF - 1, 0]);
          if (fl.lr && f2 === nF - 1) pon("lastRow", [nF - 1, 0, nF - 1, nC - 1]);
          if (fl.fr && f === 0) pon("firstRow", [0, 0, 0, nC - 1]);
          if (fl.lr && fl.lc && f2 === nF - 1 && c2 === nC - 1) pon("seCell", [f, c, f2, c2]);
          if (fl.lr && fl.fc && f2 === nF - 1 && c === 0) pon("swCell", [f, c, f2, c2]);
          if (fl.fr && fl.lc && f === 0 && c2 === nC - 1) pon("neCell", [f, c, f2, c2]);
          if (fl.fr && fl.fc && f === 0 && c === 0) pon("nwCell", [f, c, f2, c2]);
          let fondo = null; const r0 = {}, bd = {};
          for (const [e, [rf1, rc1, rf2, rc2]] of capas) {
            if (e.fondo) fondo = e.fondo;
            if (e.color) r0.color = e.color;
            if (e.b != null) r0.b = e.b;
            if (e.i != null) r0.i = e.i;
            const b = e.bd || {};
            const t = f === rf1 ? b.top : b.insideH, bo = f2 === rf2 ? b.bottom : b.insideH, l = c === rc1 ? b.left : b.insideV, ri = c2 === rc2 ? b.right : b.insideV;
            if (t) bd.t = t; if (bo) bd.b = bo; if (l) bd.l = l; if (ri) bd.r = ri;
          }
          const tcPr = h1(tc, "tcPr");
          if (tcPr) {
            if (h1(tcPr, "solidFill")) fondo = colorCss(h1(tcPr, "solidFill")) || fondo;
            else if (h1(tcPr, "noFill")) fondo = null;
            [["lnT", "t"], ["lnB", "b"], ["lnL", "l"], ["lnR", "r"]].forEach(([n, k2]) => { const x = borde(h1(tcPr, n)); if (x) bd[k2] = x; });
          }
          const mar = [["marL", 91440], ["marT", 45720], ["marR", 91440], ["marB", 45720]].map(([n, d]) => Math.round((((tcPr && at(tcPr, n) != null ? +at(tcPr, n) : d) / EMU) * kx) * 100) / 100);
          const tp = textoPpt(h1(tc, "txBody"), { tipoPh: null, maestro: mMae, r0 });
          celdas[k] = { cs: cs > 1 ? cs : undefined, rs: rs > 1 ? rs : undefined, fondo: fondo || undefined, bd: Object.keys(bd).length ? bd : undefined, html: tp.html, t: textoCelda(tc), anc: (tcPr && at(tcPr, "anchor")) || undefined, mar };
        }));
        return { anchos, altos, celdas };
      }
      const LINEAS_PRST = /^(line|straightConnector1|bentConnector[2-5]|curvedConnector[2-5])$/;
      async function elemento(el, grupo) {
        const nm = el.localName;
        if (nm === "grpSp") {
          // Grupo: sus hijos se ubican con la transformación del grupo.
          const gx = todos(h1(el, "grpSpPr"), "xfrm")[0];
          let g = grupo;
          if (gx) {
            const o = h1(gx, "off"), e = h1(gx, "ext"), co = h1(gx, "chOff"), ce = h1(gx, "chExt");
            if (o && e && co && ce) {
              const sx = (+at(ce, "cx") ? +at(e, "cx") / +at(ce, "cx") : 1), sy = (+at(ce, "cy") ? +at(e, "cy") / +at(ce, "cy") : 1);
              const prev = grupo || { tx: (v) => v, ty: (v) => v, s: (v) => v, sy: (v) => v };
              g = { tx: (v) => prev.tx(+at(o, "x") + (v - +at(co, "x")) * sx), ty: (v) => prev.ty(+at(o, "y") + (v - +at(co, "y")) * sy), s: (v) => prev.s(v * sx), sy: (v) => prev.sy(v * sy) };
            }
          }
          for (const c of hs(el)) await elemento(c, g);
          return;
        }
        const ajustar = (r) => (grupo && r ? Object.assign({}, r, { x: grupo.tx(r.x), y: grupo.ty(r.y), w: grupo.s(r.w), h: grupo.sy(r.h) }) : r);
        if (nm === "sp") {
          const nv = todos(el, "nvPr")[0], ph = todos(nv, "ph")[0];
          const phInfo = ph ? { tipo: at(ph, "type") || "body", idx: at(ph, "idx") || "", clave: (at(ph, "type") || "body") + ":" + (at(ph, "idx") || "") } : null;
          // de dónde hereda el texto: el marcador del diseño (por índice o tipo) y el del patrón (por tipo)
          const ctxTexto = { tipoPh: phInfo ? (at(ph, "type") ? phInfo.tipo : "_idx") : null };
          if (phInfo) {
            ctxTexto.phDiseno = (phInfo.idx && mDis["#" + phInfo.idx]) || mDis[phInfo.clave] || mDis[phInfo.tipo] || null;
            const tm = /^(title|ctrTitle)$/.test(phInfo.tipo) ? "title" : /^(body|subTitle|obj)$/.test(phInfo.tipo) || !at(ph, "type") ? "body" : phInfo.tipo;
            ctxTexto.phMaestro = mMae[tm] || null;
          }
          ctxTexto.maestro = mMae;
          const spPr = h1(el, "spPr"), r = ajustar(xfrmDe(spPr, phInfo, heredados));
          if (!r) return;
          const txBody = h1(el, "txBody");
          const texto = txBody ? txBody.textContent.trim() : "";
          const prst = at(todos(spPr, "prstGeom")[0], "prst");
          if (prst && LINEAS_PRST.test(prst) && !texto && !phInfo) { lineaOffice(el, ajustar); return; }
          const estSp = h1(el, "style"), conRellenoPropio = spPr && (h1(spPr, "solidFill") || h1(spPr, "noFill") || h1(spPr, "gradFill") || h1(spPr, "blipFill") || h1(spPr, "pattFill"));
          // (con su transparencia: un relleno con alpha se ve a través, como en PowerPoint)
          const relleno = spPr && h1(spPr, "solidFill") ? colorCss(h1(spPr, "solidFill")) : spPr && h1(spPr, "gradFill") ? colorCss(todos(h1(spPr, "gradFill"), "gs")[0] || h1(spPr, "gradFill")) : !conRellenoPropio && estSp && +at(h1(estSp, "fillRef"), "idx") ? hexDe(h1(estSp, "fillRef")) : null;
          // degradado o imagen de relleno, y geometría libre: es una forma aunque no tenga color liso
          const pintura = spPr && (h1(spPr, "gradFill") || h1(spPr, "blipFill") || h1(spPr, "pattFill")) ? await pinturaDe(spPr, rels) : null;
          const libre = geomLibre(spPr);
          const esTexto = !!(phInfo || todos(el, "cNvSpPr").some((c) => at(c, "txBox") === "1") || (!relleno && !pintura && !libre && (!prst || prst === "rect")));
          const cNv = todos(el, "cNvPr")[0], nombreSp = at(cNv, "name") || "", descr = at(cNv, "descr") || "";
          // Ecuación exportada por NOM: la fórmula original viaja en la descripción del cuadro.
          if (/^LaTeX: /.test(descr)) { agregar(Object.assign({ tipo: "ecuacion", tex: descr.slice(7) }, geo(r))); return; }
          const titulo = phInfo && /title|ctrTitle/.test(phInfo.tipo);
          if (esTexto) {
            if (!texto) return;
            const estilo = titulo ? (phInfo.tipo === "ctrTitle" ? "h1" : "h2") : /^Título \d/.test(nombreSp) ? "h1" : /^Encabezado \d/.test(nombreSp) ? "h2" : "p";
            const tp = textoPpt(txBody, ctxTexto);
            // (un cuadro de texto puede tener relleno y contorno propios: se ven como en PowerPoint)
            const lnT = spPr && h1(spPr, "ln");
            if (relleno) tp.cuerpo.fondo = relleno;
            if (lnT && !h1(lnT, "noFill")) { const cT = colorCss(h1(lnT, "solidFill")); if (cT) tp.cuerpo.borde = [Math.max(0.5, Math.round(((+(at(lnT, "w") || 9525)) / 12700) * kPt * 100) / 100), cT]; }
            agregar(Object.assign({ tipo: "texto", estilo, html: tp.html, ppt: tp.cuerpo }, geo(r)));
          } else {
            const ln = spPr && h1(spPr, "ln"), tinta = estSp ? hexDe(h1(estSp, "fontRef")) : null;
            agregar(Object.assign({ tipo: "forma", figura: GEO_INV[prst] || "rect", radio: prst === "rect" || prst === "flowChartProcess" || libre ? 0 : undefined, color: relleno || (pintura && pintura.paradas ? pintura.paradas[0][1] : pintura && pintura.mezcla ? pintura.mezcla : "#CFE0FF"), relleno: spPr && h1(spPr, "noFill") ? false : undefined,
              cust: libre || undefined, pintura: pintura || undefined,
              // la geometría de Office tal cual (se dibuja con js/geometrias.js) y sus ajustes (las manijas amarillas)
              prst: prst && NOM.geomOffice && NOM.geomOffice.tiene(prst) ? prst : undefined, av: ajustesDe(spPr),
              borde: ln && h1(ln, "solidFill") ? hexDe(h1(ln, "solidFill")) : estSp && +at(h1(estSp, "lnRef"), "idx") && !(ln && h1(ln, "noFill")) ? hexDe(h1(estSp, "lnRef")) || undefined : undefined,
              bordeAncho: ln && h1(ln, "noFill") ? undefined : ln && at(ln, "w") ? Math.max(0.25, Math.round((+at(ln, "w") / 12700) * 100) / 100) : estSp && +at(h1(estSp, "lnRef"), "idx") ? lineasTema[+at(h1(estSp, "lnRef"), "idx")] || 0.75 : undefined, sombra: false, tColor: tinta && tinta !== "#000000" && texto ? tinta : undefined,
              texto: texto || "", html: texto ? textoPpt(txBody, ctxTexto).html : undefined, ppt: texto ? textoPpt(txBody, ctxTexto).cuerpo : undefined }, geo(r)));
          }
          return;
        }
        if (nm === "pic") {
          const r = ajustar(xfrmDe(h1(el, "spPr"), null, heredados)), blip = todos(el, "blip")[0];
          const rel = blip && rels[rid(blip, "embed")];
          const src = rel ? await dataUrl(rel.destino) : null;
          if (!r || !src) { aproximado = true; return; }
          const sr = todos(h1(el, "blipFill"), "srcRect")[0];
          const fr = (k) => Math.max(0, Math.min(0.9, (+(at(sr, k) || 0)) / 100000));
          const recorte = sr && (at(sr, "l") || at(sr, "r") || at(sr, "t") || at(sr, "b")) ? { iz: fr("l"), d: fr("r"), s: fr("t"), i: fr("b") } : undefined;
          const prst = at(todos(h1(el, "spPr"), "prstGeom")[0], "prst");
          // Efectos de la imagen: blanco y negro, brillo, contraste y transparencia.
          const filtros = {}, lum = h1(blip, "lum"), alfaI = h1(blip, "alphaModFix");
          if (h1(blip, "grayscl")) filtros.gris = 100;
          if (lum && at(lum, "bright")) filtros.brillo = Math.round(100 + +at(lum, "bright") / 1000);
          if (lum && at(lum, "contrast")) filtros.contraste = Math.round(100 + +at(lum, "contrast") / 1000);
          agregar(Object.assign({ tipo: "imagen", src, recorte, mascara: prst && prst !== "rect" && prst !== "roundRect" ? GEO_INV[prst] : undefined, alt: at(todos(el, "cNvPr")[0], "descr") || undefined,
            filtros: Object.keys(filtros).length ? filtros : undefined, opacidad: alfaI && at(alfaI, "amt") ? Math.max(0.05, +at(alfaI, "amt") / 100000) : undefined,
            // video o audio de PowerPoint: la imagen es su portada (o el ícono del altavoz)
            medio: medioDePic(el, rels) }, geo(r)));
          return;
        }
        if (nm === "graphicFrame") {
          const xf = h1(el, "xfrm"), o = xf && h1(xf, "off"), e = xf && h1(xf, "ext");
          if (!o || !e) return;
          const r = ajustar({ x: +at(o, "x"), y: +at(o, "y"), w: +at(e, "cx"), h: +at(e, "cy") });
          const tbl = todos(el, "tbl")[0];
          if (tbl) {
            const filas = hs(tbl, "tr"), celdas = {};
            let cols = 0;
            // (cada párrafo de una celda en su renglón; un texto que parece número se queda como está: «7.0», no 7)
            const textoCelda = (tc) => hs(h1(tc, "txBody"), "p").map((q) => todos(q, "t").map((x) => x.textContent).join("")).join("\n").trim();
            filas.forEach((tr, f) => { const tcs = hs(tr, "tc"); cols = Math.max(cols, tcs.length); tcs.forEach((tc, c) => { const t = textoCelda(tc); if (t) celdas["c:" + F().ref(c, f)] = F().calcularLiteral(t) !== t || /^[=']/.test(t) ? "'" + t : t; }); });
            let office;
            try { office = tablaComoPpt(tbl, textoCelda); } catch (e) { office = undefined; }
            agregar(Object.assign({ tipo: "tabla", enDiapo: true, filas: filas.length, cols, encabezado: at(todos(tbl, "tblPr")[0], "firstRow") === "1" || undefined, office }, celdas, geo(r)));
            return;
          }
          const ch = todos(el, "chart")[0], rel = ch && rels[rid(ch, "id")];
          if (rel) {
            const t = await z.texto(rel.destino);
            if (t) {
              const cd = P(t), celdas = {};
              const tipoG = todos(cd, "barChart")[0] ? (at(todos(cd, "barDir")[0], "val") === "bar" ? "barras-h" : "barras") : todos(cd, "lineChart")[0] ? "lineas" : todos(cd, "pieChart")[0] ? "pastel" : todos(cd, "areaChart")[0] ? "area" : todos(cd, "scatterChart")[0] ? "dispersion" : "barras";
              const sers = todos(cd, "ser");
              const cats = sers[0] ? todos(h1(sers[0], "cat") || h1(sers[0], "xVal"), "pt").map((p) => (todos(p, "v")[0] || {}).textContent || "") : [];
              cats.forEach((c, i) => (celdas["c:" + F().ref(0, i + 1)] = c));
              sers.forEach((sr, j) => {
                const nombre = (todos(h1(sr, "tx"), "v")[0] || {}).textContent || "Serie " + (j + 1);
                celdas["c:" + F().ref(j + 1, 0)] = nombre;
                todos(h1(sr, "val") || h1(sr, "yVal"), "pt").forEach((p) => { celdas["c:" + F().ref(j + 1, +at(p, "idx") + 1)] = (todos(p, "v")[0] || {}).textContent || ""; });
              });
              const titulo = todos(todos(cd, "title")[0], "t").map((x2) => x2.textContent).join("");
              // y su estilo de Office (colores, letras, ejes, leyenda) para dibujarla como PowerPoint; los datos siguen
              // en las celdas del bloque (se pueden editar en la app)
              let office;
              try {
                if (NOM.graficaOffice) office = NOM.graficaOffice.leer(t, { color: (n) => hexDe(n), acentos: ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((k) => "#" + (esquema[k] || "4472C4")),
                  fuente: (n) => (/^\+mn-/.test(n) ? fuentesTema.menor : /^\+mj-/.test(n) ? fuentesTema.mayor : n) || "Calibri" });
              } catch (e) { office = undefined; }
              agregar(Object.assign({ tipo: "grafica", clase: tipoG, titulo: titulo || "", filas: Math.max(2, cats.length + 1), cols: sers.length + 1, office, ptOffice: office ? Math.round(kPt * 1000) / 1000 : undefined }, celdas, geo(r)));
              return;
            }
          }
          const dg = todos(el, "relIds")[0];
          if (dg && (await smartArt(dg, r, +at(e, "cx"), +at(e, "cy")))) return;
          aproximado = true;
          return;
        }
        if (nm === "cxnSp") { lineaOffice(el, ajustar); return; }
        if (nm === "AlternateContent") { const f = h1(el, "Fallback") || h1(el, "Choice"); if (f) for (const c of hs(f)) await elemento(c, grupo); }
      }
      const arbol = todos(sld, "spTree")[0];
      const hijosArbol = hs(arbol);
      for (let i2 = 0; i2 < hijosArbol.length; i2++) {
        oxActual = i2;
        const k0 = bloques.length;
        await elemento(hijosArbol[i2], null);
        // acción al hacer clic (vínculo o botón de acción de PowerPoint): para todo lo que salió de esta figura
        const acc = accionDe(hijosArbol[i2], rels);
        if (acc) for (let k = k0; k < bloques.length; k++) bloques[k].accion = acc;
      }
      oxActual = null;
      // Animaciones de PowerPoint (p:timing): cada efecto con su árbol tal cual, apuntando a los bloques de NOM
      // (una figura agrupada son varios bloques con el mismo «ox»). Los colores del tema se resuelven aquí.
      // (con lo de PowerPoint 2010, como el rebote al final, el p:timing va en mc:AlternateContent: se lee el Choice)
      let tmEl = hs(sld.documentElement, "timing")[0];
      if (!tmEl) for (const ac of hs(sld.documentElement, "AlternateContent")) { const ch = h1(ac, "Choice") || h1(ac, "Fallback"); const t = ch && h1(ch, "timing"); if (t) { tmEl = t; break; } }
      if (tmEl && NOM.animOffice) {
        const oxDe = new Map();
        hijosArbol.forEach((h, k) => { const c = todos(h, "cNvPr")[0]; if (c && at(c, "id") != null && !oxDe.has(at(c, "id"))) oxDe.set(at(c, "id"), k); });
        const idsDe = (spid) => { const k = oxDe.get(String(spid)); return k == null ? [] : bloques.filter((b) => b.pagina === pg && b.ox === k).map((b) => b.id); };
        try {
          const r = NOM.animOffice.leerTiming(tmEl, idsDe, (el) => colorDeElemento(el));
          if (r) {
            if (r.lista.length) docCampos["anim:" + pg] = r.lista;
            if (r.medios.length) docCampos["animMedios:" + pg] = r.medios;
            if (r.bld.length) docCampos["animBld:" + pg] = r.bld;
          }
        } catch (e) { /* una animación rara no detiene la importación: al exportar se conserva la original */ }
      }
      if (!docCampos.esquemaOffice && Object.keys(esquema).length) docCampos.esquemaOffice = Object.assign({}, esquema);
      // Fondo: el de la diapositiva, o el de su diseño, o el de su patrón (con los rellenos de fondo del tema)
      const relsDis0 = relDiseno ? await relsDe(relDiseno.destino) : {};
      const relMae0 = relDiseno ? Object.values(relsDis0).find((r) => /slideMaster$/.test(r.tipo)) : null;
      const xmlDis0 = relDiseno ? P(await z.texto(relDiseno.destino) || "<x/>") : null, xmlMae0 = relMae0 ? P(await z.texto(relMae0.destino) || "<x/>") : null;
      async function fondoDeHoja(dx, relsX) {
        const bg = dx && todos(dx, "bg")[0];
        if (!bg) return null;
        const bp = h1(bg, "bgPr");
        if (bp) { for (const f of hs(bp)) { const r2 = await fondoDeRelleno(f, null, relsX); if (r2) return r2; } return null; }
        const br = h1(bg, "bgRef");
        if (br) { const idx = +(at(br, "idx") || 0), f = fondosTema[idx - 1000] || null; return f ? await fondoDeRelleno(f, hexDe(br), relsX) : (hexDe(br) ? { tipo: "color", color: hexDe(br) } : null); }
        return null;
      }
      const fondo = (await fondoDeHoja(sld, rels)) || (await fondoDeHoja(xmlDis0, relsDis0)) || (relMae0 ? await fondoDeHoja(xmlMae0, await relsDe(relMae0.destino)) : null);
      if (fondo) docCampos["fondo:" + pg] = fondo;
      // Capa heredada: las formas del patrón y del diseño (no sus marcadores) van debajo de lo de la diapositiva,
      // como en PowerPoint. Se guarda una vez por diseño y no se escribe en la diapositiva al exportar.
      if (relDiseno && at(sld.documentElement, "showMasterSp") !== "0") {
        const clave = relDiseno.destino;
        if (!capas.has(clave)) {
          const capa = [], fuentesCapa = [];
          if (relMae0 && at(xmlDis0.documentElement, "showMasterSp") !== "0") fuentesCapa.push([xmlMae0, await relsDe(relMae0.destino)]);
          fuentesCapa.push([xmlDis0, relsDis0]);
          const relsSlide = rels, agregarSlide = agregar;
          let ord = 0;
          agregar = (b) => { const id = "cap" + capas.size + "_" + ++ord; capa.push(Object.assign({ id, orden: ord, z: 0 }, b)); return id; };
          try {
            for (const [dx, relsX] of fuentesCapa) {
              rels = relsX;
              const arbolC = todos(dx, "spTree")[0];
              for (const h of hs(arbolC)) { if (todos(h, "ph").length) continue; oxActual = null; await elemento(h, null); }
            }
          } catch (e) { /* una forma rara del patrón no detiene la importación */ }
          rels = relsSlide; agregar = agregarSlide; oxActual = null;
          capas.set(clave, capa);
          if (capa.length) docCampos["capaBl:" + clave] = capa;
        }
        if ((capas.get(clave) || []).length) docCampos["capa:" + pg] = clave;
      }
      // Transición y notas del orador
      const tr = todos(sld, "transition")[0];
      if (tr) { const c = hs(tr)[0]; if (c) docCampos["trans:" + pg] = { fade: "desvanecer", push: "deslizar", zoom: "zoom", wipe: "deslizar", cover: "deslizar", dissolve: "desvanecer" }[c.localName] || "desvanecer"; }
      // la transición de PowerPoint tal cual (las 149 de Office 2024, con su duración, dirección y avance automático)
      if (NOM.transOffice) { try { const to = NOM.transOffice.leer(sld); if (to) docCampos["transOffice:" + pg] = to; } catch (e) { /* se queda la de antes */ } }
      const relNotas = Object.values(rels).find((r) => /notesSlide$/.test(r.tipo));
      if (relNotas) {
        const t = await z.texto(relNotas.destino);
        if (t) {
          const cuerpo = todos(P(t), "sp").find((sp) => { const ph = todos(sp, "ph")[0]; return ph && at(ph, "type") === "body"; });
          const txt = cuerpo ? hs(h1(cuerpo, "txBody"), "p").map((p) => p.textContent).join("\n").trim() : "";
          if (txt) docCampos["nota:" + pg] = txt;
        }
      }
    }
    // los vínculos a otra diapositiva apuntan a su página de NOM
    const paginaDeRuta = new Map(paginas.map((pg) => [docCampos["os:" + pg], pg]));
    bloques.forEach((b) => { if (b.accion && b.accion.ruta) { const pg = paginaDeRuta.get(b.accion.ruta); if (pg) b.accion = { tipo: "diapo", pagina: pg }; else delete b.accion; } });
    // Las secciones (Inicio → Sección: p14:sectionLst, por número de diapositiva), las presentaciones personalizadas
    // (p:custShowLst, por relación) y Configurar presentación (presProps.xml: repetir hasta Esc, usar intervalos…).
    {
      const porNum = new Map(), porRid = new Map();
      todos(todos(docP, "sldIdLst")[0], "sldId").forEach((x) => {
        const r = relsPres[rid(x, "id")], pg = r && paginaDeRuta.get(r.destino);
        if (pg) { porNum.set(at(x, "id"), pg); porRid.set(rid(x, "id"), pg); }
      });
      const secs = todos(docP, "section").filter((x) => /powerpoint\/20\d\d\/main/.test(x.namespaceURI || ""))
        .map((x) => ({ id: at(x, "id") || "", nombre: at(x, "name") || "", d: todos(x, "sldId").map((y) => porNum.get(at(y, "id"))).filter(Boolean) }));
      if (secs.length) docCampos.secciones = secs;
      const pers = todos(todos(docP, "custShowLst")[0], "custShow")
        .map((c) => ({ id: at(c, "id") || "", nombre: at(c, "name") || "", d: todos(c, "sld").map((y) => porRid.get(rid(y, "id"))).filter(Boolean) }));
      if (pers.length) docCampos.personalizadas = pers;
      const relPP = Object.values(relsPres).find((r) => /\/presProps$/.test(r.tipo));
      const tpp = relPP ? await z.texto(relPP.destino) : null, sp = tpp ? todos(P(tpp), "showPr")[0] : null;
      if (sp) {
        const si = (a) => /^(1|true)$/.test(at(sp, a) || ""), no = (a) => /^(0|false)$/.test(at(sp, a) || "");
        const cfg = {};
        if (si("loop")) cfg.repetir = true;
        if (no("useTimings")) cfg.sinIntervalos = true;
        if (no("showAnimation")) cfg.sinAnimacion = true;
        if (todos(sp, "kiosk")[0]) cfg.tipo = "kiosco"; else if (todos(sp, "browse")[0]) cfg.tipo = "ventana";
        const rg = todos(sp, "sldRg")[0];
        if (rg) cfg.rango = [+at(rg, "st") || 1, +at(rg, "end") || 1];
        const cs = todos(sp, "custShow")[0], p2 = cs && pers.find((x) => x.id === at(cs, "id"));
        if (p2) cfg.personalizada = p2.id;
        if (Object.keys(cfg).length) docCampos.presentacion = cfg;
      }
    }
    if (!paginas.length) throw new Error("La presentación no tiene diapositivas que se puedan leer");
    // (lo que no se ve se conserva porque va dentro el archivo original; uno muy grande no se guarda: formatos.js lo avisa)
    const conserva = !(NOM.origen && NOM.origen.MAX_ORIGEN) || u8.length <= NOM.origen.MAX_ORIGEN;
    return { archivo: u8, paginas, bloques, doc: docCampos, aviso: aproximado ? (conserva ? "Algunos objetos de PowerPoint (videos o imágenes de otro formato) no se ven aquí, pero se conservan: al descargar en PowerPoint salen completos" : "Algunos objetos de PowerPoint (videos o imágenes de otro formato) no se ven aquí") : null };
  }

  /**
   * La protección de Office en lo que se descarga, como está AHORA en el documento: «restringir edición» y «contraseña
   * para modificar» de Word (settings.xml), estructura protegida y «contraseña para modificar» de Excel (workbook.xml),
   * «contraseña para modificar» de PowerPoint (presentation.xml) y «marcado como final» (docProps/custom.xml). Lo que se
   * quitó aquí se quita del archivo; lo que se puso, se escribe como lo escribe Office 365. Si no cambia nada, el archivo
   * sale igual. (La protección de cada hoja de Excel va con la hoja, en parcharHoja.)
   */
  function conProteccion(u8, doc, tipo) {
    let z;
    try { z = NOM.origen.abrirZip(u8); } catch (e) { return u8; }
    const v = (k) => doc.valor("_doc", k);
    const cambios = new Map();
    const cripto = (p, w) => (p && p.hash ? ` ${w}cryptProviderType="rsaAES" ${w}cryptAlgorithmClass="hash" ${w}cryptAlgorithmType="typeAny" ${w}cryptAlgorithmSid="${p.sid || 14}" ${w}cryptSpinCount="${+p.vueltas || 100000}" ${w}hash="${x(p.hash)}" ${w}salt="${x(p.sal || "")}"` : "");
    const conHash = (p, pre) => (p && p.hash ? ` ${pre ? pre + "AlgorithmName" : "algorithmName"}="${x(p.alg || "SHA-512")}" ${pre ? pre + "HashValue" : "hashValue"}="${x(p.hash)}" ${pre ? pre + "SaltValue" : "saltValue"}="${x(p.sal || "")}" ${pre ? pre + "SpinCount" : "spinCount"}="${+p.vueltas || 100000}"` : "");
    const quita = (s, tag) => s.replace(new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag.replace(":", "\\:")}>`, "g"), "");
    if (tipo === "docx") {
      const st = z.texto("word/settings.xml");
      if (st) {
        let s = quita(quita(st, "w:documentProtection"), "w:writeProtection");
        const pr = v("proteccion"), mod = v("modificar");
        const ED = { lectura: "readOnly", comentarios: "comments", cambios: "trackedChanges", formularios: "forms" };
        if (mod && (mod.recomendada || (mod.p && mod.p.hash))) s = s.replace(/(<w:settings\b[^>]*>)/, `$1<w:writeProtection${mod.recomendada ? ' w:recommended="1"' : ""}${cripto(mod.p, "w:")}/>`);
        if (pr && ED[pr.tipo]) {
          const el = `<w:documentProtection w:edit="${ED[pr.tipo]}" w:enforcement="1"${cripto(pr.p, "w:")}/>`;
          s = /<w:defaultTabStop\b/.test(s) ? s.replace(/<w:defaultTabStop\b/, el + "<w:defaultTabStop") : s.replace(/<\/w:settings>/, el + "</w:settings>");
        }
        // «Guiones automáticos»: solo se toca si cambió (si no, el archivo sale idéntico)
        const tenia = /<w:autoHyphenation\b(?![^>]*w:val="(?:0|false|off)")[^>]*\/?>/.test(s), quiere = !!v("guiones");
        if (tenia !== quiere) {
          s = quita(s, "w:autoHyphenation");
          if (quiere) s = /<w:defaultTabStop\b[^>]*\/>/.test(s) ? s.replace(/(<w:defaultTabStop\b[^>]*\/>)/, "$1<w:autoHyphenation/>") : s.replace(/(<w:settings\b[^>]*>)/, "$1<w:autoHyphenation/>");
        }
        // las tabulaciones predeterminadas (Tabulaciones…): solo si cambiaron
        { const quiere = Math.round((+v("tabDef") || 36) * 20), m = /<w:defaultTabStop\b[^>]*w:val="(\d+)"[^>]*\/>/.exec(s);
          if (m && +m[1] !== quiere) s = s.replace(m[0], `<w:defaultTabStop w:val="${quiere}"/>`); }
        // «Control de cambios» (va antes de la protección y de las tabulaciones, como pide Word)
        const teniaCC = /<w:trackRevisions\b(?![^>]*w:val="(?:0|false|off)")[^>]*\/?>/.test(s), quiereCC = !!v("controlCambios");
        if (teniaCC !== quiereCC) {
          s = quita(s, "w:trackRevisions");
          if (quiereCC) s = /<w:documentProtection\b/.test(s) ? s.replace(/<w:documentProtection\b/, "<w:trackRevisions/><w:documentProtection") : /<w:defaultTabStop\b/.test(s) ? s.replace(/<w:defaultTabStop\b/, "<w:trackRevisions/><w:defaultTabStop") : s.replace(/(<w:settings\b[^>]*>)/, "$1<w:trackRevisions/>");
        }
        if (s !== st) cambios.set("word/settings.xml", s);
      }
    } else if (tipo === "xlsx") {
      const wb = z.texto("xl/workbook.xml");
      if (wb) {
        let s = quita(quita(wb, "workbookProtection"), "fileSharing");
        const lp = v("libroProt"), mod = v("modificar");
        if (mod && (mod.recomendada || (mod.p && (mod.p.hash || mod.p.viejo)))) {
          const el = `<fileSharing${mod.recomendada ? ' readOnlyRecommended="1"' : ""}${mod.p && mod.p.viejo ? ` reservationPassword="${x(mod.p.viejo)}"` : conHash(mod.p)}/>`;
          s = /<fileVersion\b[^>]*\/>/.test(s) ? s.replace(/(<fileVersion\b[^>]*\/>)/, "$1" + el) : s.replace(/(<workbook\b[^>]*>)/, "$1" + el);
        }
        if (lp && (lp.estructura || lp.ventanas)) {
          const el = `<workbookProtection${conHash(lp.p, "workbook")}${lp.p && lp.p.viejo ? ` workbookPassword="${x(lp.p.viejo)}"` : ""}${lp.estructura ? ' lockStructure="1"' : ""}${lp.ventanas ? ' lockWindows="1"' : ""}/>`;
          s = /<bookViews\b/.test(s) ? s.replace(/<bookViews\b/, el + "<bookViews") : s.replace(/<sheets\b/, el + "<sheets");
        }
        if (s !== wb) cambios.set("xl/workbook.xml", s);
      }
    } else if (tipo === "pptx") {
      const pr0 = z.texto("ppt/presentation.xml");
      if (pr0) {
        let s = quita(pr0, "p:modifyVerifier");
        const mod = v("modificar");
        if (mod && mod.p && mod.p.hash) {
          const p = mod.p, el = `<p:modifyVerifier cryptProviderType="rsaAES" cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" cryptAlgorithmSid="${p.sid || 14}" spinCount="${+p.vueltas || 100000}" saltData="${x(p.sal || "")}" hashData="${x(p.hash)}"/>`;
          s = /<p:extLst>/.test(s) ? s.replace(/<p:extLst>(?![\s\S]*<p:extLst>)/, el + "<p:extLst>") : s.replace(/<\/p:presentation>/, el + "</p:presentation>");
        }
        if (s !== pr0) cambios.set("ppt/presentation.xml", s);
      }
    }
    // marcado como final (Word, Excel y PowerPoint): la propiedad _MarkAsFinal de docProps/custom.xml
    {
      const fin = !!v("final"), cu = z.texto("docProps/custom.xml"), tiene = !!cu && /name="_MarkAsFinal"/.test(cu);
      const marcado = tiene && /name="_MarkAsFinal"[^>]*>\s*<vt:bool>\s*(true|1)\s*<\/vt:bool>/i.test(cu);
      if (fin !== marcado) {
        if (!fin && cu) cambios.set("docProps/custom.xml", cu.replace(/<property\b[^>]*name="_MarkAsFinal"[^>]*>[\s\S]*?<\/property>/, ""));
        else if (fin && cu) {
          const ids = [...cu.matchAll(/pid="(\d+)"/g)].map((m) => +m[1]), pid = Math.max(1, ...ids) + 1;
          cambios.set("docProps/custom.xml", cu.replace(/<\/Properties>/, `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="_MarkAsFinal"><vt:bool>true</vt:bool></property></Properties>`));
        } else if (fin) {
          cambios.set("docProps/custom.xml", XML + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_MarkAsFinal"><vt:bool>true</vt:bool></property></Properties>`);
          const ct = z.texto("[Content_Types].xml"), rl = z.texto("_rels/.rels");
          if (ct && !/PartName="\/docProps\/custom\.xml"/.test(ct)) cambios.set("[Content_Types].xml", ct.replace(/<\/Types>/, '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>'));
          if (rl && !/docProps\/custom\.xml/.test(rl)) cambios.set("_rels/.rels", rl.replace(/<\/Relationships>/, '<Relationship Id="rIdNomFinal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>'));
        }
      }
    }
    if (!cambios.size) return u8;
    const sale = [];
    for (const nm of z.nombres()) {
      if (cambios.has(nm)) { sale.push({ nombre: nm, datos: cambios.get(nm) }); cambios.delete(nm); continue; }
      const e = z.entradas.get(nm);
      if (e && e.metodo === 8) sale.push({ nombre: nm, crudo: { datos: e.datos, crc: e.crc, tam: e.tam } });
      else sale.push({ nombre: nm, datos: z.bytes(nm) || new Uint8Array(0) });
    }
    for (const [nm, d] of cambios) sale.push({ nombre: nm, datos: d });
    return zip(sale);
  }

  /** Para la hoja: pinta en `campos` lo que el estilo de la tabla `lo` pone en cada celda (con el tema del libro o el de Office). */
  const tablaXl = {
    pintar(lo, campos, tema) {
      const est2 = lo.estilo === "" ? null : (NOM.ESTILOS_TABLA || {})[lo.estilo || "TableStyleMedium2"];
      if (!est2) return campos;
      pintarTablaXl({ c1: lo.c1, f1: lo.f1, c2: lo.c2, f2: lo.f2, enc: lo.enc ? 1 : 0, tot: lo.tot ? 1 : 0, rayasF: lo.rayasF !== false, rayasC: !!lo.rayasC, primeraC: !!lo.primeraC, ultimaC: !!lo.ultimaC }, est2, colorTablaConTema(tema), campos);
      return campos;
    },
    color: (x, tema) => colorTablaConTema(tema)(x),
    /** Lo mismo para una tabla dinámica, desde su pivotTableDefinition (la que arma js/dinamicas.js). */
    pintarDinamica(xml, campos, tema) {
      // (la definición guardada lleva la marca de la caché, que no es XML válido hasta que se exporta)
      const pd = new DOMParser().parseFromString(String(xml).replace("\u0001CACHE\u0001", "1"), "application/xml").documentElement, si = pd && [...pd.childNodes].find((n) => n.localName === "pivotTableStyleInfo");
      const est2 = si && (NOM.ESTILOS_TABLA || {})[si.getAttribute("name")];
      if (est2) pintarDinamicaXl(pd, est2, colorTablaConTema(tema), campos);
      return campos;
    },
  };
  NOM.ooxml = { tablaXl, zip, docx, xlsx, pptx, formulaExcel, crc32, leerDocx, leerXlsx, leerPptx, modeloXlsx, huellasDocx, huellasPptx, parcialesPptx, abrirZip, conProteccion, seccionesPptx, usarOrigen: true };
})();

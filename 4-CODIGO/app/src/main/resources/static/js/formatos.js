/* NOM · ConversorFormato del lado del cliente (CU-28 importar, CU-29 exportar).
   El editor nunca conoce un formato de archivo: todo entra y sale como bloques. */
(function () {
  "use strict";
  const NOM = window.NOM;
  const F = () => NOM.formulas;

  const IMPORTABLES = [".docx", ".xlsx", ".pptx", ".nom", ".json", ".md", ".txt", ".csv", ".tsv", ".html", ".htm"];
  // Formatos que la gente sí tiene pero que todavía no se abren: se dejan elegir en el selector para
  // decirle exactamente qué hacer (si no, el teléfono ni los mostraba y no se sabía por qué).
  const VIEJOS = {
    ".doc": ["Word 97-2003", "Word", ".docx"], ".xls": ["Excel 97-2003", "Excel", ".xlsx"], ".ppt": ["PowerPoint 97-2003", "PowerPoint", ".pptx"],
    ".odt": ["LibreOffice Writer", "LibreOffice", ".docx"], ".ods": ["LibreOffice Calc", "LibreOffice", ".xlsx"], ".odp": ["LibreOffice Impress", "LibreOffice", ".pptx"],
    ".rtf": ["texto enriquecido (RTF)", "Word", ".docx"],
  };
  const ACEPTAR = IMPORTABLES.concat(Object.keys(VIEJOS));
  const MAX_BYTES = 10 * 1024 * 1024;

  // ------------------------------------------------------------ utilidades
  const textoPlano = (h) => { const d = document.createElement("div"); d.innerHTML = h || ""; return d.textContent || ""; };
  function descargar(nombre, mime, contenido) {
    if (window.NomAndroid && window.NomAndroid.guardarArchivo) {
      const b64 = btoa(unescape(encodeURIComponent(contenido)));
      window.NomAndroid.guardarArchivo(nombre, mime, b64);
      return;
    }
    const url = URL.createObjectURL(new Blob([contenido], { type: mime + ";charset=utf-8" }));
    const a = Object.assign(document.createElement("a"), { href: url, download: nombre });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
  /** Para archivos binarios (.docx, .xlsx, .pptx, .png). */
  function descargarBytes(nombre, mime, bytes) {
    if (window.NomAndroid && window.NomAndroid.guardarArchivo) {
      let bin = "";
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      window.NomAndroid.guardarArchivo(nombre, mime, btoa(bin));
      return;
    }
    const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    const a = Object.assign(document.createElement("a"), { href: url, download: nombre });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
  const MIME = {
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  };
  const limpioNombre = (s) => String(s || "documento").replace(/[\\/:*?"<>|]+/g, "-").slice(0, 80);

  function ordenados(doc) {
    const pags = doc.paginas();
    return pags.flatMap((p) => doc.deLaPagina(p).map((b) => Object.assign(b, { _pag: pags.indexOf(p) })));
  }

  function celdasDe(b) {
    const ev = F().evaluador((c, f) => b["c:" + F().ref(c, f)]);
    let maxF = 0, maxC = 0;
    Object.keys(b).forEach((k) => {
      if (!k.startsWith("c:") || b[k] === "" || b[k] == null) return;
      const r = F().parseRef(k.slice(2)); if (!r) return;
      maxF = Math.max(maxF, r.f + 1); maxC = Math.max(maxC, r.c + 1);
    });
    const filas = [];
    for (let f = 0; f < maxF; f++) {
      const fila = [];
      for (let c = 0; c < maxC; c++) fila.push(F().formatear(ev.valor(c, f)));
      filas.push(fila);
    }
    return filas;
  }

  // ------------------------------------------------------------ exportar
  const esTexto = (doc) => (doc.valor("_doc", "modo") || "TEXTO") === "TEXTO" && NOM.textoRico;
  const textoDe = (p) => p.tramos.map((t) => t.x || "").join("");

  /** HTML de un objeto del texto (imagen, tabla, salto, línea, índice). */
  /** Notas al pie de la exportación a HTML (se listan al final). */
  let notasHtml = [];
  function objetoHtml(e, ps) {
    if (e.tipo === "imagen") return e.src ? `<figure style="margin:0;display:inline-block"><img src="${NOM.esc(NOM.srcExport ? NOM.srcExport(e.src) : e.src)}" alt="${NOM.esc(e.alt || "")}" style="max-width:100%;width:${+e.w || 320}px">${e.pie ? `<figcaption style="font-style:italic;text-align:center">${NOM.esc(e.pie)}</figcaption>` : ""}</figure>` : "";
    if (e.tipo === "tabla") {
      let h = `<table style="border-collapse:collapse">`;
      const tapadas = new Set();
      for (const k in e) if (k.startsWith("m:") && /^\d+x\d+$/.test(String(e[k] || ""))) { const p = F().parseRef(k.slice(2)), [w, hh] = String(e[k]).split("x").map(Number); for (let f = p.f; f < p.f + hh; f++) for (let c = p.c; c < p.c + w; c++) if (c !== p.c || f !== p.f) tapadas.add(F().ref(c, f)); }
      for (let f = 0; f < (e.filasTexto || 3); f++) {
        h += "<tr>";
        for (let c = 0; c < (e.colsTexto || 3); c++) {
          const r = F().ref(c, f);
          if (tapadas.has(r)) continue;
          const m = String(e["m:" + r] || "").split("x").map(Number);
          const st = [e["r:" + r] ? `background:${e["r:" + r]}` : "", e["n:" + r] || (e.encabezado && f === 0) ? "font-weight:bold" : "", e["i:" + r] ? "font-style:italic" : "", e["k:" + r] ? `color:${e["k:" + r]}` : "",
            e["al:" + r] ? `text-align:${{ centro: "center", der: "right" }[e["al:" + r]] || "left"}` : "", e.bordes === "ninguno" ? "border:0" : "border:1px solid #8E8E8E", "padding:2pt 5pt"].filter(Boolean).join(";");
          h += `<td style="${st}"${m[0] > 1 ? ` colspan="${m[0]}"` : ""}${m[1] > 1 ? ` rowspan="${m[1]}"` : ""}>${NOM.esc(e["c:" + r] || "")}</td>`;
        }
        h += "</tr>";
      }
      return h + "</table>";
    }
    if (e.tipo === "ecuacion") return NOM.ecuaciones ? NOM.ecuaciones.html(e.tex || "", !!e.bloque) : NOM.esc(e.tex || "");
    if (e.tipo === "nota") { notasHtml.push(e.texto || ""); return `<sup><a href="#nota-${notasHtml.length}" id="ref-${notasHtml.length}">${notasHtml.length}</a></sup>`; }
    if (e.texto != null) return NOM.esc(String(e.texto));
    if (e.tipo === "salto") return `<div style="break-after:page"></div>`;
    if (e.tipo === "linea") return "<hr>";
    if (e.tipo === "indice") return `<nav><b>Contenido</b>${ps.filter((p) => /^h[1-3]$/.test(p.pa.estilo || "")).map((p) => `<div style="padding-left:${(+p.pa.estilo.slice(1) - 1) * 16}px">${NOM.esc(textoDe(p))}</div>`).join("")}</nav>`;
    return "";
  }
  function aHtmlTexto(doc, nombre) {
    const ps = NOM.textoRico.parrafosDe(doc);
    notasHtml = [];
    let h = "", lista = null;
    const cerrar = () => { if (lista) { h += `</${lista}>`; lista = null; } };
    ps.forEach((p) => {
      const pa = p.pa;
      const cont = p.tramos.map((t) => (t.e ? objetoHtml(t.e, ps) : NOM.textoRico.envolver(t.x, t.a || {}))).join("") || "<br>";
      const st = [pa.alin ? `text-align:${{ centro: "center", der: "right", just: "justify" }[pa.alin]}` : "", pa.sangria && !pa.lista ? `margin-left:${pa.sangria * 36}pt` : "", pa.interlineado ? `line-height:${pa.interlineado}` : "",
        !pa.lista && (+pa.sangiz || +pa.francesa) ? `padding-left:${(+pa.sangiz || 0) + (+pa.francesa || 0)}cm` : "", !pa.lista && +pa.sangder ? `padding-right:${+pa.sangder}cm` : "",
        !pa.lista && (pa.primera != null || +pa.francesa) ? `text-indent:${(+pa.primera || 0) - (+pa.francesa || 0)}cm` : "",
        NOM.textoRico && NOM.textoRico.cssBorde ? NOM.textoRico.cssBorde(pa.borde) : "", /^#[0-9a-f]{6}$/i.test(pa.sombra || "") ? `background-color:${pa.sombra}` : "", +pa.lineaPt ? `line-height:${+pa.lineaPt}pt` : "", pa.dir === "rtl" ? "direction:rtl" : ""].filter(Boolean).join(";");
      const sa = st ? ` style="${st}"` : "";
      if (pa.lista) {
        const tag = pa.lista === "numeros" ? "ol" : "ul";
        if (lista !== tag) { cerrar(); h += `<${tag}>`; lista = tag; }
        h += `<li${pa.nivel ? ` style="margin-left:${pa.nivel * 18}pt"` : ""}>${pa.lista === "tareas" ? (pa.hecho ? "☑ " : "☐ ") : ""}${cont}</li>\n`;
        return;
      }
      cerrar();
      const tag = { h1: "h1", h2: "h2", h3: "h3", titulo: "h1", cita: "blockquote", codigo: "pre" }[pa.estilo] || "p";
      h += `<${tag}${pa.estilo === "titulo" ? ' class="titulo"' : pa.estilo === "subtitulo" ? ' class="subtitulo"' : ""}${sa}>${cont}</${tag}>\n`;
    });
    cerrar();
    if (notasHtml.length) h += `<hr><ol class="notas">${notasHtml.map((t, i) => `<li id="nota-${i + 1}">${NOM.esc(t)} <a href="#ref-${i + 1}">↩</a></li>`).join("")}</ol>\n`;
    return `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>${NOM.esc(nombre)}</title>
<style>body{font:11pt/1.15 Calibri,Carlito,"Segoe UI",Arial,sans-serif;max-width:6.5in;margin:1in auto;padding:0 1rem;color:#000}p{margin:0 0 8pt}h1,h2,h3{font-family:"Calibri Light",Carlito,Calibri,sans-serif;font-weight:normal;color:#2F5496;margin:12pt 0 0}h1{font-size:16pt}h2{font-size:13pt}h3{font-size:12pt;color:#1F3763}h1.titulo{font-size:28pt;color:#000}.subtitulo{color:#5A5A5A}blockquote{font-style:italic;color:#404040;text-align:center}pre{font:10pt "Courier New",monospace;background:#F3F3F3;padding:2pt 6pt;white-space:pre-wrap}table{border-collapse:collapse;width:100%}td{border:1px solid #8E8E8E;padding:2pt 5pt;vertical-align:top}a{color:#0563C1}</style>
</head><body>
${h}</body></html>`;
  }
  function aMarkdownTexto(doc) {
    const ps = NOM.textoRico.parrafosDe(doc);
    const md = (t) => {
      let s = (t.x || "").replace(/([*_`\\[\]#])/g, "\\$1").replace(/\u2028/g, "  \n");
      const a = t.a || {};
      if (!s.trim()) return s;
      if (a.b && a.i) s = `***${s}***`; else if (a.b) s = `**${s}**`; else if (a.i) s = `*${s}*`;
      if (a.s) s = `~~${s}~~`;
      if (a.enlace) s = `[${s}](${a.enlace})`;
      return s;
    };
    const out = [];
    let prevLista = false;
    ps.forEach((p) => {
      const pa = p.pa;
      if (p.tramos.length === 1 && p.tramos[0].e) {
        const e = p.tramos[0].e;
        if (e.tipo === "tabla") {
          const filas = [];
          for (let f = 0; f < (e.filasTexto || 3); f++) { const r = []; for (let c = 0; c < (e.colsTexto || 3); c++) r.push(String(e["c:" + F().ref(c, f)] || "").replace(/\|/g, "\\|").replace(/\n/g, " ")); filas.push("| " + r.join(" | ") + " |"); }
          if (filas.length) filas.splice(1, 0, "|" + " --- |".repeat(e.colsTexto || 3));
          out.push(filas.join("\n"));
        } else if (e.tipo === "linea" || e.tipo === "salto") out.push("---");
        else if (e.tipo === "imagen") out.push("![imagen](imagen-incrustada)");
        prevLista = false;
        return;
      }
      const cuerpo = p.tramos.map((t) => (t.e ? (t.e.tipo === "imagen" ? "![imagen](imagen-incrustada)" : "") : md(t))).join("");
      const sang = "  ".repeat(pa.nivel || 0);
      let linea;
      if (pa.lista === "vinetas") linea = sang + "- " + cuerpo;
      else if (pa.lista === "numeros") linea = sang + "1. " + cuerpo;
      else if (pa.lista === "tareas") linea = sang + (pa.hecho ? "- [x] " : "- [ ] ") + cuerpo;
      else if (pa.estilo === "h1" || pa.estilo === "titulo") linea = "# " + cuerpo;
      else if (pa.estilo === "h2") linea = "## " + cuerpo;
      else if (pa.estilo === "h3") linea = "### " + cuerpo;
      else if (pa.estilo === "cita") linea = "> " + cuerpo;
      else if (pa.estilo === "codigo") linea = "    " + textoDe(p);
      else linea = cuerpo;
      if (pa.lista && prevLista) out[out.length - 1] += "\n" + linea; else out.push(linea);
      prevLista = !!pa.lista;
    });
    return out.join("\n\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
  }

  function aHtml(doc, nombre) {
    if (esTexto(doc)) return aHtmlTexto(doc, nombre);
    const partes = ordenados(doc).map((b) => {
      if (b.tipo === "texto") {
        const tag = { h1: "h1", h2: "h2", cita: "blockquote" }[b.estilo] || "p";
        if (b.estilo === "lista") return `<ul><li>${b.html || ""}</li></ul>`;
        if (b.estilo === "tarea") return `<p>${b.hecho ? "☑" : "☐"} ${b.html || ""}</p>`;
        return `<${tag}>${b.html || ""}</${tag}>`;
      }
      if (b.tipo === "tabla") {
        // (las imágenes que flotan sobre una hoja de Excel van después de su tabla)
        const imgs = Array.isArray(b.dib) ? b.dib.filter((d) => d && d.tipo === "imagen" && d.src).map((d) => `<p><img src="${NOM.esc(NOM.srcExport ? NOM.srcExport(d.src) : d.src)}" alt=""></p>`).join("") : "";
        return "<table>" + celdasDe(b).map((r) => "<tr>" + r.map((v) => `<td>${NOM.esc(v)}</td>`).join("") + "</tr>").join("") + "</table>" + imgs;
      }
      if (b.tipo === "imagen") return `<p><img src="${NOM.esc(NOM.srcExport ? NOM.srcExport(b.src) : b.src)}" alt="${NOM.esc(b.alt || "")}"></p>`;
      if (b.tipo === "forma") return `<p><em>[${NOM.esc(b.figura)}]</em> ${NOM.esc(b.texto || "")}</p>`;
      return "";
    }).join("\n").replace(/<\/ul>\n<ul>/g, "\n");
    return `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>${NOM.esc(nombre)}</title>
<style>body{font-family:system-ui,sans-serif;max-width:46rem;margin:3rem auto;padding:0 1rem;line-height:1.7;color:#222}table{border-collapse:collapse}td{border:1px solid #ccc;padding:.3rem .6rem}blockquote{border-left:4px solid #3D72F2;margin:0;padding-left:1rem;color:#555}img{max-width:100%}</style>
</head><body>
${partes}
</body></html>`;
  }

  function aMarkdown(doc) {
    if (esTexto(doc)) return aMarkdownTexto(doc);
    return ordenados(doc).map((b) => {
      const t = textoPlano((b.html || "").replace(/<b>|<\/b>|<strong>|<\/strong>/g, "**").replace(/<i>|<\/i>|<em>|<\/em>/g, "_"));
      if (b.tipo === "texto") {
        return { h1: "# " + t, h2: "## " + t, cita: "> " + t, lista: "- " + t, tarea: (b.hecho ? "- [x] " : "- [ ] ") + t }[b.estilo] || t;
      }
      if (b.tipo === "tabla") {
        const filas = celdasDe(b);
        if (!filas.length) return "";
        const sep = "|" + filas[0].map(() => " --- ").join("|") + "|";
        return [filas[0], ...filas.slice(1)].map((r, i) => "| " + r.join(" | ") + " |" + (i === 0 ? "\n" + sep : "")).join("\n");
      }
      if (b.tipo === "forma") return `- ${b.texto || ""}`;
      if (b.tipo === "imagen") return `![${b.alt || "imagen"}](imagen-incrustada)`;
      return "";
    }).filter((x, i, a) => x !== "" || (a[i - 1] !== "")).join("\n\n").replace(/\n\n(- )/g, "\n$1") + "\n";
  }

  function aCsv(doc, pagina) {
    const t = doc.deLaPagina(pagina).find((b) => b.tipo === "tabla") || doc.vivos().find((b) => b.tipo === "tabla");
    if (!t) return null;
    return celdasDe(t).map((r) => r.map((v) => (/[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v)).join(",")).join("\r\n") + "\r\n";
  }

  const n1 = (v) => Math.round(v * 10) / 10;
  function aSvg(doc, pagina) {
    const bs = doc.deLaPagina(pagina).filter((b) => b.x != null && !b.oculto && b.tipo !== "conector");
    if (!bs.length) return null;
    const E = NOM.esc, F = NOM.figuras;
    const minX = Math.min(...bs.map((b) => b.x)) - 40, minY = Math.min(...bs.map((b) => b.y)) - 40;
    const maxX = Math.max(...bs.map((b) => b.x + b.w)) + 40, maxY = Math.max(...bs.map((b) => b.y + b.h)) + 40;
    const porId = Object.fromEntries(bs.map((b) => [b.id, b]));
    // 8.4 Flechas con su forma, puntas, estilo y textos: el mismo dibujo que en el lienzo.
    const lineas = doc.deLaPagina(pagina).filter((b) => b.tipo === "conector" && porId[b.desde] && porId[b.hasta]).map((c) => NOM.geometria.svgConector(c, porId[c.desde], porId[c.hasta]));
    const luz = (c) => { const n = parseInt(String(c || "#ffffff").slice(1), 16) || 0; return (0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255; };
    const fuente = 'font-family="Nunito, Arial, sans-serif"';
    const formas = bs.sort((a, b) => (a.z || 0) - (b.z || 0)).map((b) => {
      const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
      const giro = b.rot ? ` transform="rotate(${+b.rot} ${cx} ${cy})"` : "";
      const op = b.opacidad != null && b.opacidad < 1 ? ` opacity="${+b.opacidad}"` : "";
      const voltear = b.volteoH || b.volteoV ? ` translate(${b.volteoH ? b.w : 0} ${b.volteoV ? b.h : 0}) scale(${b.volteoH ? -1 : 1} ${b.volteoV ? -1 : 1})` : "";
      let fig = "", fondo = null, e = null;
      if (b.tipo === "forma" && F) {
        e = F.estiloFigura(b);
        fondo = e.relleno;
        const dash = e.estilo === "guiones" ? ` stroke-dasharray="${e.ancho * 3} ${e.ancho * 2}"` : e.estilo === "puntos" ? ` stroke-dasharray="0 ${e.ancho * 2}"` : "";
        const tapa = e.estilo === "puntos" || e.f === "actor" ? "round" : "butt";
        fig = `<g transform="translate(${b.x} ${b.y})${voltear}">${F.partesFigura(e, b.w, b.h).map((pz) => `<path d="${pz.d}" fill="${pz.relleno ? E(pz.relleno) : "none"}" stroke="${pz.trazo ? E(pz.trazo) : "none"}" stroke-width="${e.ancho}"${pz.trazo ? dash : ""} stroke-linejoin="round" stroke-linecap="${tapa}"/>`).join("")}</g>`;
      } else if (b.tipo === "imagen" && b.src) fig = `<image href="${E(b.src)}" x="${b.x}" y="${b.y}" width="${b.w}" height="${b.h}" preserveAspectRatio="none"${voltear ? ` transform="translate(${b.x} ${b.y})${voltear} translate(${-b.x} ${-b.y})"` : ""}/>`;
      else if (b.tipo === "icono" && F) fig = `<svg x="${b.x}" y="${b.y}" width="${b.w}" height="${b.h}" viewBox="0 0 24 24" fill="none" stroke="${E(b.color || "#2B2A33")}" stroke-width="${+b.grosor || 2}" stroke-linecap="round" stroke-linejoin="round">${F.iconoSeguro(b)}</svg>`;
      const colorT = b.tColor || (fondo && luz(fondo) < 0.5 ? "#FFFFFF" : "#1F1E26");
      const fs = +b.tTam || (b.tipo === "texto" ? 22 : 16), lh = fs * 1.3;
      const peso = F && F.enNegrita(b) ? 700 : 500;
      let txt = "";
      if (e && F.COMPARTIMENTOS.has(e.f)) {
        // Clase y entidad: nombre arriba y cada compartimento con su línea divisoria.
        const ls = String(b.texto || "").split("\n").map((l) => l.replace(/\s+$/, ""));
        const nombre = ls.shift() || "", resto = ls.filter((l) => l.trim() && !/^-{2,}$/.test(l.trim()));
        const secs = e.f === "clase" ? [resto.filter((l) => !/\(.*\)/.test(l)), resto.filter((l) => /\(.*\)/.test(l))] : [resto];
        const f2 = +b.tTam || 14, l2 = f2 * 1.35;
        let y = b.y + 6 + l2;
        txt += `<text x="${cx}" y="${y - 5}" text-anchor="middle" ${fuente} font-size="${f2}" font-weight="800" fill="${E(colorT)}">${E(nombre)}</text>`;
        y += 6;
        secs.forEach((sec) => {
          if (y < b.y + b.h - 2) txt += `<line x1="${b.x}" y1="${y}" x2="${b.x + b.w}" y2="${y}" stroke="${E(e.trazo || "#2B2A33")}" stroke-width="2"/>`;
          y += 5;
          sec.forEach((l) => {
            y += l2;
            const clave = e.f === "entidad" && /^(PK\b|#)/i.test(l.trim());
            if (y < b.y + b.h) txt += `<text x="${b.x + 10}" y="${y - 4}" ${fuente} font-size="${f2}" font-weight="600"${clave ? ' text-decoration="underline"' : ""} fill="${E(colorT)}">${E(clave ? l.trim().replace(/^(PK\s*:?\s*|#\s*)/i, "") : l)}</text>`;
          });
          y += sec.length ? 5 : 14;
        });
      } else {
        const t = b.tipo === "forma" ? (b.html && NOM.diapo ? NOM.diapo.textoPlano(b.html) : b.texto) : b.tipo === "texto" ? (NOM.diapo ? NOM.diapo.textoPlano(b.html) : textoPlano(b.html)) : "";
        if (t) {
          // Renglones como en el lienzo: se parte por palabras según el ancho de la figura.
          const caben = Math.max(4, Math.floor(((e && e.f === "actor" ? b.w + 60 : b.w) - 20) / (fs * 0.56)));
          const ls = [];
          String(t).split("\n").forEach((renglon) => {
            let l = "";
            renglon.split(/\s+/).forEach((p) => {
              if (!l) l = p;
              else if ((l + " " + p).length <= caben) l += " " + p;
              else { ls.push(l); l = p; }
            });
            ls.push(l);
          });
          ls.splice(40);
          const anchor = b.tAlin === "izq" ? ["start", b.x + 10] : b.tAlin === "der" ? ["end", b.x + b.w - 10] : ["middle", cx];
          // El actor lleva su nombre abajo; lo demás, centrado a lo alto.
          const y0 = e && e.f === "actor" ? b.y + b.h - (ls.length - 1) * lh - fs * 0.3 : cy - ((ls.length - 1) * lh) / 2 + fs * 0.35;
          txt = `<text text-anchor="${anchor[0]}" ${fuente} font-size="${fs}" font-weight="${peso}"${b.tCursiva ? ' font-style="italic"' : ""}${b.tSubrayado ? ' text-decoration="underline"' : ""} fill="${E(colorT)}">${ls.map((l, i) => `<tspan x="${anchor[1]}" y="${n1(y0 + i * lh)}">${E(l)}</tspan>`).join("")}</text>`;
        }
      }
      return `<g${giro}${op}>${fig}${txt}</g>`;
    });
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${minX} ${minY} ${maxX - minX} ${maxY - minY}" width="${maxX - minX}" height="${maxY - minY}">
<rect x="${minX}" y="${minY}" width="100%" height="100%" fill="#fff"/>
${lineas.join("\n")}
${formas.join("\n")}
</svg>`;
  }

  // ------------------------------------------------------------ el formato propio (.nom)
  /** Documento en memoria armado con operaciones (lo que devuelve importar). */
  function docDeOps(nombre, modo, ops) {
    const doc = new NOM.DocCRDT("mem-" + Math.random().toString(36).slice(2, 8), { bloques: {}, reloj: 0 }, { local: true, soloMemoria: true });
    doc._fusionar({ id: "local:c1", bloque: "_doc", campos: { nombre, modo, paginas: ["p1"] }, t: 1, n: "local" });
    (ops || []).forEach((op) => doc._fusionar(op));
    return doc;
  }
  /** Texto letra por letra → párrafos compactos: [[formatoDePárrafo | 0, [texto, formato | 0, …]], …]. */
  function parrafosDeFoto(f) {
    const ps = [];
    let tr = [], ult = null;
    const cierra = (pa) => { ps.push([pa && Object.keys(pa).length ? pa : 0, tr]); tr = []; ult = null; };
    for (const it of f.items) {
      if (it.c === "\n" && !it.e) { cierra(it.a); continue; }
      const a = it.a && Object.keys(it.a).length ? it.a : 0;
      if (it.e) { tr.push({ e: it.e }, a); ult = null; continue; }
      const k = a ? NOM.origen.estable(a) : "";
      if (ult === k) tr[tr.length - 2] += it.c; else { tr.push(it.c, a); ult = k; }
    }
    cierra(f.fin);
    return ps;
  }
  const igual = (a, b) => a === b || (a != null && b != null && typeof a === "object" && typeof b === "object" && NOM.origen.estable(a) === NOM.origen.estable(b));
  /** Huella de un bloque tal como está ahora (se compara con la que se guardó al importar). */
  function huellaBloque(doc, id) { const o = doc.bloque(id) || {}, r = {}; for (const k in o) if (k !== "id" && o[k] != null) r[k] = o[k]; return NOM.origen.huella(r); }
  /** El documento como datos simples: { d: datos generales, b: bloques, t: textos }. Sin historia. */
  function fotoModelo(doc) {
    const sinNulos = (o) => { const r = {}; for (const k in o) if (k !== "id" && o[k] != null) r[k] = o[k]; return r; };
    const d = sinNulos(doc.bloque("_doc") || {}), b = {}, t = {};
    for (const id in doc.estado.bloques) {
      if (id === "_doc" || id === NOM.origen.BLOQUE || !doc.vivo(id)) continue;
      b[id] = sinNulos(doc.bloque(id));
    }
    for (const k in doc.estado.seq || {}) if (doc.tieneSecuencia(k)) t[k] = parrafosDeFoto(doc.secuencia(k).foto());
    return { d, b, t };
  }
  const GRANDE = 256;   // un campo de texto más largo que esto (una imagen) lleva su propia huella
  /** Huellas de un documento recién importado: con ellas se sabe después qué cambió, sin guardar otra copia. */
  function huellasDe(foto) {
    const H = NOM.origen.huella, h = { d: {}, b: {}, t: {} };
    for (const k in foto.d) h.d[k] = H(foto.d[k]);
    for (const id in foto.b) {
      const x = foto.b[id];
      let g = null;
      for (const k in x) if (typeof x[k] === "string" && x[k].length > GRANDE) (g = g || {})[k] = H(x[k]);
      h.b[id] = g ? [H(x), g] : H(x);
    }
    for (const k in foto.t) h.t[k] = H(foto.t[k]);
    return h;
  }
  /**
   * El documento como archivo .nom (versión 3): binario y comprimido. Si vino de un archivo de
   * Office lleva ese archivo dentro (recomprimido todo junto, pesa menos que el original) y, del
   * documento, solo lo que cambió desde que se importó.
   */
  function aNomBytes(doc) {
    const O = NOM.origen, H = O.huella, foto = fotoModelo(doc), og = O.leer(doc);
    // "nom-colab" es la marca del formato .nom desde antes del cambio de nombre: se queda para que abran los archivos viejos y las apps viejas.
    const adj = [], json = { f: "nom-colab", v: 3 };
    let cuerpo, enOrigen = null;
    if (og && og.h && og.h.b) {
      const h = og.h, d = {}, b = {}, bq = {}, t = {}, mezcla = [];
      const nombres = og.zip.nombres();
      enOrigen = new Map();
      const danadas = {};
      nombres.forEach((n) => {
        const u = og.zip.bytes(n), cr = u == null && og.zip.crudo ? og.zip.crudo(n) : null;
        if (cr) { danadas[n] = { crc: cr.crc, tam: cr.tam }; adj.push({ datos: cr.datos, crudo: true }); return; }
        const crudo = O.YA_COMPRIMIDO.test(n); adj.push({ datos: u, crudo }); if (crudo || /\.(emf|wmf|bmp|tiff?|svg)$/i.test(n)) enOrigen.set(O.huellaBytes(u), n);
      });
      json.og = { fmt: og.fmt, nombres };
      if (Object.keys(danadas).length) json.og.danadas = danadas;
      // De una hoja de cálculo importada se guardan solo las celdas que cambiaron.
      const baseHojas = {};
      if (og.fmt === "xlsx" && NOM.ooxml.modeloXlsx) { try { NOM.ooxml.modeloXlsx(og.zip).hojas.forEach((x, i) => { baseHojas["hoja-p" + (i + 1)] = x.campos; }); } catch (e) { /* sin base: se guarda la hoja completa */ } }
      for (const k in foto.d) if (h.d[k] !== H(foto.d[k])) d[k] = foto.d[k];
      for (const id in foto.b) {
        const x = foto.b[id], hb = h.b[id];
        if (!hb) { b[id] = x; continue; }
        if ((Array.isArray(hb) ? hb[0] : hb) === H(x)) continue;
        const y = baseHojas[id];
        if (y && y.ox === x.ox) {
          const c = {}, q = [];
          for (const k in x) if (!(k in y) || !igual(x[k], y[k])) c[k] = x[k];
          for (const k in y) if (!(k in x)) q.push(k);
          b[id] = c; mezcla.push(id);
          if (q.length) bq[id] = q;
          continue;
        }
        const c = Object.assign({}, x), g = Array.isArray(hb) ? hb[1] : null;
        if (g) for (const k in g) if (typeof c[k] === "string" && H(c[k]) === g[k]) c[k] = "\u0003";   // igual que en el original
        b[id] = c;
      }
      for (const k in foto.t) if (h.t[k] !== H(foto.t[k])) t[k] = foto.t[k];
      cuerpo = { d, dx: Object.keys(h.d).filter((k) => !(k in foto.d)), b, bq, bm: mezcla, bx: Object.keys(h.b).filter((id) => !foto.b[id]), t, tx: Object.keys(h.t).filter((k) => !(k in foto.t)) };
      ["dx", "bx", "tx", "bm"].forEach((k) => { if (!cuerpo[k].length) delete cuerpo[k]; });
      if (!Object.keys(bq).length) delete cuerpo.bq;
    } else cuerpo = foto;
    // lo del almacén de binarios va dentro del .nom (con sus bytes: el archivo se abre en cualquier lado, sin servidor);
    // ya precargado por exportarListo
    if (NOM.binarios) cuerpo = NOM.binarios.conDatos(cuerpo);
    Object.assign(json, O.sacarDatos(cuerpo, adj, enOrigen, new Map()));
    return NOM.paquete.empacar(json, adj);
  }
  /** Abre un .nom versión 3: { d, b, t } y, si vino de Office, las operaciones del archivo de origen. */
  async function deNomBytes(u) {
    const O = NOM.origen, { json, adjuntos } = NOM.paquete.desempacar(u);
    if (json.f !== "nom-colab") throw new Error("Ese archivo no es un documento de NOM");
    if (!json.og) { const c = O.meterDatos({ d: json.d || {}, b: json.b || {}, t: json.t || {} }, adjuntos, null); return { d: c.d, b: c.b, t: c.t, opsOrigen: [] }; }
    // Vino de Office: se vuelve a armar ese archivo, se importa igual que la primera vez y encima van los cambios.
    const partes = new Map();
    json.og.nombres.forEach((n, i) => partes.set(n, adjuntos[i]));
    const dan = json.og.danadas || {};
    const archivo = NOM.ooxml.zip(json.og.nombres.map((n) => (dan[n] ? { nombre: n, crudo: { datos: partes.get(n), crc: dan[n].crc, tam: dan[n].tam } } : { nombre: n, datos: partes.get(n) })), { nivel: 4 });
    const imp = await importar(new File([archivo], "origen." + json.og.fmt));
    const B = fotoModelo(docDeOps(imp.nombre, imp.modo, imp.ops));
    const c = O.meterDatos({ d: json.d || {}, b: json.b || {}, t: json.t || {} }, adjuntos, partes);
    const d = Object.assign({}, B.d, c.d), b = {}, t = Object.assign({}, B.t, c.t);
    (json.dx || []).forEach((k) => delete d[k]);
    (json.tx || []).forEach((k) => delete t[k]);
    const fuera = new Set(json.bx || []), mezcla = new Set(json.bm || []);
    for (const id in B.b) if (!fuera.has(id)) b[id] = B.b[id];
    for (const id in c.b) {
      const x = c.b[id], y = B.b[id] || {};
      if (mezcla.has(id)) b[id] = Object.assign({}, y, x);
      else { b[id] = x; for (const k in x) if (x[k] === "\u0003") { if (y[k] != null) x[k] = y[k]; else delete x[k]; } }
    }
    for (const id in json.bq || {}) if (b[id]) { b[id] = Object.assign({}, b[id]); json.bq[id].forEach((k) => delete b[id][k]); }
    return { d, b, t, opsOrigen: imp.ops.filter((op) => op.bloque === O.BLOQUE) };
  }
  /** .nom versión 3 → operaciones para crear el documento. */
  function opsDeNom(r) {
    const ops = [];
    let i = 0;
    const { nombre: _n, modo: _m, ...resto } = r.d;
    if (Object.keys(resto).length) ops.push({ id: "imp:d", bloque: "_doc", campos: resto, t: 2, n: "importar" });
    for (const id in r.b) ops.push({ id: "imp:" + i++, bloque: id, campos: r.b[id], t: 2, n: "importar" });
    const partidas = NOM.DocCRDT.partir(ops);
    let k = 0;
    for (const bl in r.t) {
      const ps = r.t[bl].map(([pa, tr]) => {
        const tramos = [];
        for (let j = 0; j < tr.length; j += 2) { const a = tr[j + 1] || undefined; tramos.push(typeof tr[j] === "string" ? { x: tr[j], a } : { e: tr[j].e, a }); }
        return { a: pa || undefined, tramos };
      });
      partidas.push(...NOM.Secuencia.crearOps(bl, ps, "importar", 2, "imp" + (k++ ? k : "")));
    }
    return partidas.concat(r.opsOrigen || []);
  }
  /** El formato de antes (versión 2, JSON): se sigue pudiendo abrir. */
  function aNom(doc) {
    const bloques = {};
    for (const id in doc.estado.bloques) if (id === "_doc" || (doc.vivo(id) && id !== "_origen")) bloques[id] = doc.bloque(id);
    return JSON.stringify({ formato: "nom-colab", version: 2, exportado: new Date().toISOString(), bloques, seq: doc.estado.seq || {} });
  }
  /** Si el archivo es demasiado grande para guardarlo dentro del documento, se avisa: al descargarlo se arma de nuevo. */
  function avisoSinOrigen(aviso, archivo) {
    const max = NOM.origen && NOM.origen.MAX_ORIGEN;
    if (!archivo || !max || archivo.length <= max) return aviso;
    const t = `Este archivo pesa más de ${Math.round(max / 1048576)} MB, así que NOM no guarda el original dentro del documento: al descargarlo se arma de nuevo con lo que NOM sabe hacer (lo que no conoce, como sus videos o sonidos, no sale)`;
    return aviso ? aviso.replace(/[.\s]*$/, ". ") + t + "." : t + ".";
  }
  /** Guarda dentro del documento recién importado el archivo de Office del que vino (y sus huellas). */
  /**
   * Las huellas de lo recién importado: al exportar dicen qué cambió la persona y qué sale idéntico del original. Se
   * vuelven a calcular cuando las imágenes pasan al almacén de binarios (binarios.js), porque el documento ya no lleva
   * sus datos sino su referencia.
   */
  function huellasOrigen(fmt, nombre, modo, ops) {
    const tmp = docDeOps(nombre, modo, ops.filter((o) => o.bloque !== (NOM.origen ? NOM.origen.BLOQUE : "_origen"))), h = huellasDe(fotoModelo(tmp));
    if (fmt === "docx" && NOM.ooxml.huellasDocx && NOM.textoRico) h.u = NOM.ooxml.huellasDocx(tmp);
    if (fmt === "pptx" && NOM.ooxml.huellasPptx) {
      h.pp = NOM.ooxml.huellasPptx(tmp); h.q = {};
      tmp.vivos().forEach((b) => { if (b.ox != null) h.q[b.id] = NOM.ooxml.parcialesPptx(b); });
    }
    return h;
  }
  function conOrigen(fmt, nombre, modo, ops, archivo) {
    if (!archivo || !NOM.origen) return ops;
    try {
      const h = huellasOrigen(fmt, nombre, modo, ops);
      const og = NOM.origen.crearOps(fmt, archivo, h);
      return og.length ? ops.concat(og) : ops;
    } catch (e) { if (typeof console !== "undefined") console.warn("No se pudo guardar el archivo de origen", e); return ops; }
  }

  const EXPORTABLES = [
    { id: "docx", nombre: "Word (.docx)", desc: "Se abre en Word, Google Docs o LibreOffice Writer", icono: "file-text" },
    { id: "xlsx", nombre: "Excel (.xlsx)", desc: "Con fórmulas: Excel, Google Sheets o LibreOffice Calc", icono: "sheet" },
    { id: "pptx", nombre: "PowerPoint (.pptx)", desc: "Se abre en PowerPoint, Google Slides o LibreOffice Impress", icono: "presentation" },
    { id: "png", nombre: "Imagen (.png)", desc: "El diagrama de esta página como imagen", icono: "image" },
    { id: "nom", nombre: "NOM (.nom)", desc: "Sin pérdida: se vuelve a importar tal cual", icono: "layers" },
    { id: "html", nombre: "Página web (.html)", desc: "Para abrir en cualquier navegador", icono: "file-text" },
    { id: "md", nombre: "Markdown (.md)", desc: "Texto plano con formato ligero", icono: "type" },
    { id: "csv", nombre: "CSV (.csv)", desc: "La tabla, para Excel o LibreOffice Calc", icono: "sheet" },
    { id: "svg", nombre: "Imagen vectorial (.svg)", desc: "El diagrama de esta página", icono: "workflow" },
    { id: "pdf", nombre: "PDF (imprimir)", desc: "Usa el diálogo de impresión del sistema", icono: "download" },
  ];

  /** Formatos que tienen sentido para cada tipo de documento, el de Office primero. */
  const POR_TIPO = {
    TEXTO: ["docx", "pdf", "html", "md", "nom"],
    HOJA: ["xlsx", "csv", "pdf", "html", "nom"],
    PRESENTACION: ["pptx", "pdf", "nom"],
    DIAGRAMA: ["pptx", "png", "svg", "pdf", "nom"],
  };
  const paraTipo = (tipo) => (POR_TIPO[tipo] || POR_TIPO.TEXTO).map((id) => EXPORTABLES.find((f) => f.id === id));

  function aPng(svg) {
    return new Promise((ok, mal) => {
      const img = new Image();
      img.onload = () => {
        const k = Math.min(2, 4000 / Math.max(img.width, img.height));
        const cv = Object.assign(document.createElement("canvas"), { width: Math.round(img.width * k), height: Math.round(img.height * k) });
        cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
        cv.toBlob((b) => b ? b.arrayBuffer().then((a) => ok(new Uint8Array(a))) : mal(new Error("No se pudo crear la imagen")), "image/png");
      };
      img.onerror = () => mal(new Error("No se pudo crear la imagen"));
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
    });
  }

  /** Devuelve un aviso si algo se aproximó (CU-29 3a), o lanza si no hay qué exportar. */
  /**
   * Igual que exportar, pero antes trae a la memoria lo que el documento tiene en el almacén de binarios (imágenes y
   * el Office original): los escritores de Office y del .nom los necesitan ya. Sin red y sin tenerlos en este equipo,
   * avisa en vez de descargar un archivo incompleto.
   */
  async function exportarListo(doc, formato, nombre, pagina, opcTipo, opc) {
    if (NOM.binarios && NOM.binarios.refsDoc(doc).size) await NOM.binarios.precargar(doc);
    return exportar(doc, formato, nombre, pagina, opcTipo, opc);
  }
  function exportar(doc, formato, nombre, pagina, opcTipo, opc) {
    const base = limpioNombre(nombre);
    switch (formato) {
      case "nom": descargarBytes(base + ".nom", "application/octet-stream", aNomBytes(doc)); return null;
      // (con opc.contrasena, el archivo sale protegido con contraseña como lo hace Office: hay que escribirla para abrirlo)
      case "docx": descargarBytes(base + ".docx", MIME.docx, paraOffice(NOM.ooxml.docx(doc), doc, "docx", opc)); return hayLibres(doc, ["conector"]) ? "Las flechas no se incluyen en Word" : null;
      case "xlsx": descargarBytes(base + ".xlsx", MIME.xlsx, paraOffice(NOM.ooxml.xlsx(doc), doc, "xlsx", opc)); return null;
      case "pptx": descargarBytes(base + ".pptx", MIME.pptx, paraOffice(NOM.ooxml.pptx(doc, (doc.valor("_doc", "modo") || "") === "DIAGRAMA" || opcTipo === "DIAGRAMA"), doc, "pptx", opc)); return null;
      case "png": {
        const s = aSvg(doc, pagina);
        if (!s) throw new Error("Esta página no tiene figuras para exportar");
        aPng(s).then((b) => descargarBytes(base + ".png", "image/png", b)).catch((e) => NOM.aviso(e.message, { tipo: "error" }));
        return null;
      }
      case "html": descargar(base + ".html", "text/html", aHtml(doc, nombre)); return hayLibres(doc, ["forma", "conector"]) ? "Las figuras se exportaron como texto" : null;
      case "md": descargar(base + ".md", "text/markdown", aMarkdown(doc)); return hayLibres(doc, ["imagen", "forma"]) ? "Las imágenes y figuras se aproximaron en Markdown" : null;
      case "csv": {
        const c = aCsv(doc, pagina);
        if (!c) throw new Error("Este documento no tiene ninguna tabla");
        descargar(base + ".csv", "text/csv", c);
        return "Se exportaron los valores calculados (sin fórmulas)";
      }
      case "svg": {
        const s = aSvg(doc, pagina);
        if (!s) throw new Error("Esta página no tiene figuras para exportar");
        descargar(base + ".svg", "image/svg+xml", s);
        return null;
      }
      case "pdf":
        // En la app no existe "imprimir" del navegador: se usa la impresión de Android (también guarda PDF).
        if (window.NomAndroid && window.NomAndroid.imprimir) window.NomAndroid.imprimir(base);
        else setTimeout(() => window.print(), 80);
        return null;
      default: throw new Error("Formato no soportado");
    }
  }
  const hayLibres = (doc, tipos) => doc.vivos().some((b) => tipos.includes(b.tipo));

  // ------------------------------------------------------------ importar
  function parseCsv(texto, sep) {
    const filas = []; let fila = [], campo = "", comillas = false;
    for (let i = 0; i < texto.length; i++) {
      const c = texto[i];
      if (comillas) {
        if (c === '"' && texto[i + 1] === '"') { campo += '"'; i++; }
        else if (c === '"') comillas = false;
        else campo += c;
      } else if (c === '"') comillas = true;
      else if (c === sep) { fila.push(campo); campo = ""; }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && texto[i + 1] === "\n") i++;
        fila.push(campo); filas.push(fila); fila = []; campo = "";
      } else campo += c;
    }
    if (campo !== "" || fila.length) { fila.push(campo); filas.push(fila); }
    return filas;
  }
  const inline = (s) => NOM.esc(s).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/(^|[^\w])_(.+?)_(?=$|[^\w])/g, "$1<i>$2</i>");

  function bloquesTexto(lineas) {
    const bs = []; let orden = 1, aproximado = false;
    lineas.forEach((l) => {
      let m;
      const b = { id: NOM.DocCRDT.nuevoId(), tipo: "texto", pagina: "p1", orden: orden++, estilo: "p", html: "" };
      if ((m = /^#\s+(.*)/.exec(l))) { b.estilo = "h1"; b.html = inline(m[1]); }
      else if ((m = /^#{2,6}\s+(.*)/.exec(l))) { b.estilo = "h2"; b.html = inline(m[1]); }
      else if ((m = /^>\s?(.*)/.exec(l))) { b.estilo = "cita"; b.html = inline(m[1]); }
      else if ((m = /^\s*[-*]\s+\[([ xX])\]\s+(.*)/.exec(l))) { b.estilo = "tarea"; b.hecho = m[1] !== " "; b.html = inline(m[2]); }
      else if ((m = /^\s*([-*+]|\d+[.)])\s+(.*)/.exec(l))) { b.estilo = "lista"; b.html = inline(m[2]); }
      else if (/^\s*\|/.test(l) || /^!\[/.test(l)) { aproximado = true; b.html = NOM.esc(l); }
      else b.html = inline(l);
      bs.push(b);
    });
    return { bs, aproximado };
  }

  function comoOps(bloques, paginas) {
    const r = [];
    if (paginas) r.push({ id: "imp:p", bloque: "_doc", campos: { paginas }, t: 2, n: "importar" });
    bloques.forEach((b, i) => { const { id, ...c } = b; r.push({ id: "imp:" + i, bloque: id, campos: c, t: 2, n: "importar" }); });
    return NOM.DocCRDT.partir(r);
  }

  // ------------------------------------------------------------ archivos de Office con contraseña (js/cifrado-office.js)
  /** Pide la contraseña de un archivo de Office cifrado (se puede reintentar) y devuelve el .zip de adentro. La
      contraseña no se guarda en ningún lado: solo sirve para abrirlo en este equipo. */
  function pedirContrasena(nombreArchivo, probar) {
    return new Promise((ok, mal) => {
      const cuerpo = NOM.html(`<form class="col" autocomplete="off"><p class="suave">«${NOM.esc(nombreArchivo)}» está protegido con contraseña en Office. Escríbela para abrirlo; no se guarda en ningún lado.</p>
        <div class="campo"><label for="contr-office">Contraseña</label><div class="entrada"><input id="contr-office" type="password" data-c autocomplete="current-password" maxlength="255"></div></div>
        <p class="chico" data-err role="alert" style="color:var(--error)" hidden></p></form>`);
      const pie = NOM.html(`<div style="display:contents"><span class="crece"></span><button class="btn" type="button" data-no>Cancelar</button><button class="btn primario" type="button" data-si>Abrir</button></div>`);
      let hecho = false;
      const h = NOM.hoja({ titulo: "Archivo con contraseña", icono: "lock", cuerpo, pie, alCerrar: () => { if (!hecho) { hecho = true; mal(new Error("Se necesita la contraseña para abrir este archivo")); } } });
      const input = cuerpo.querySelector("[data-c]"), err = cuerpo.querySelector("[data-err]"), si = pie.querySelector("[data-si]");
      const intentar = () => {
        const c = input.value;
        if (!c) { input.focus(); return; }
        si.disabled = true; si.textContent = "Abriendo…"; err.hidden = true;
        // (las 100 000 vueltas de la clave tardan un momento: primero se pinta «Abriendo…»)
        setTimeout(() => {
          let zip = null;
          try { zip = probar(c); } catch (e) { err.textContent = e.message; err.hidden = false; si.disabled = false; si.textContent = "Abrir"; return; }
          if (!zip) { err.textContent = "Esa contraseña no es la del archivo."; err.hidden = false; si.disabled = false; si.textContent = "Abrir"; input.select(); return; }
          hecho = true; h.cerrar(); ok(zip);
        }, 40);
      };
      si.onclick = intentar;
      pie.querySelector("[data-no]").onclick = () => h.cerrar();
      cuerpo.onsubmit = (e) => { e.preventDefault(); intentar(); };   // (el cuerpo es el formulario)
      setTimeout(() => input.focus(), 60);
    });
  }
  /** Si el archivo de Office viene con contraseña, lo abre (con la de `opc.contrasena` o pidiéndola). */
  async function abrirSiCifrado(archivo, opc) {
    const C = NOM.cifradoOffice;
    if (!C) return archivo;
    const u = new Uint8Array(await archivo.arrayBuffer());
    if (!C.cifrado(u)) return archivo;
    let zip;
    if (opc && opc.contrasena != null) { zip = C.descifrar(u, opc.contrasena); if (!zip) throw new Error("Esa contraseña no es la del archivo"); }
    else zip = await pedirContrasena(archivo.name, (c) => C.descifrar(u, c));
    const f = new File([zip], archivo.name, { type: archivo.type });
    f.conContrasena = true;
    return f;
  }
  /** Lo que se descarga para Office, con contraseña si se pidió (Agile, como Office 365). */
  const conContrasena = (bytes, opc) => (opc && opc.contrasena && NOM.cifradoOffice ? NOM.cifradoOffice.cifrar(bytes, opc.contrasena) : bytes);
  /** Lo que se descarga para Office: con su protección de edición como está ahora (ooxml.conProteccion) y, si se pidió,
      con contraseña para abrir. */
  const paraOffice = (bytes, doc, tipo, opc) => conContrasena(NOM.ooxml.conProteccion ? NOM.ooxml.conProteccion(bytes, doc, tipo) : bytes, opc);

  async function importar(archivo, opc) {
    const nombre = archivo.name.replace(/\.[^.]+$/, "");
    const ext = (archivo.name.match(/\.[^.]+$/) || [""])[0].toLowerCase();
    if (VIEJOS[ext]) {
      const [que, prog, nuevo] = VIEJOS[ext];
      throw new Error(`Los archivos ${ext} (${que}) todavía no se abren aquí. Ábrelo en ${prog}, LibreOffice o Google Drive y guárdalo como ${nuevo}: ese sí se importa completo.`);
    }
    if (!IMPORTABLES.includes(ext)) throw new Error(`Formato no soportado. Puedes importar: ${IMPORTABLES.join(", ")}`);
    if (ext === ".docx" || ext === ".xlsx" || ext === ".pptx") {
      // (el tamaño se revisa antes de leerlo: un archivo enorme no debe cargarse completo en la memoria del equipo)
      if (archivo.size > 60 * 1024 * 1024) throw new Error("El archivo pesa más de 60 MB");
      // (un archivo de Office con contraseña: se abre con ella y después se importa como cualquier otro)
      archivo = await abrirSiCifrado(archivo, opc);
      const r = ext === ".docx" ? await importarDocx(archivo, nombre) : ext === ".xlsx" ? await importarXlsx(archivo, nombre) : await importarPptx(archivo, nombre);
      if (archivo.conContrasena) r.aviso = (r.aviso ? r.aviso + " " : "") + "El archivo tenía contraseña de Office: al descargarlo puedes ponerle una otra vez.";
      return r;
    }
    if (ext === ".nom") {
      // Versión 3: binario comprimido. (La 2 era JSON y se lee más abajo.)
      if (archivo.size > 60 * 1024 * 1024) throw new Error("El archivo pesa más de 60 MB");
      const u = new Uint8Array(await archivo.arrayBuffer());
      if (NOM.paquete.esPaquete(u)) {
        const r = await deNomBytes(u);
        return { nombre: r.d.nombre || nombre, modo: r.d.modo || "TEXTO", ops: opsDeNom(r), aviso: null };
      }
    }
    if (archivo.size > MAX_BYTES) throw new Error("El archivo pesa más de 10 MB");
    const texto = await archivo.text();
    if (/�/.test(texto.slice(0, 2000)) && ext !== ".txt") throw new Error("El archivo parece dañado o no es texto");
    let aviso = null;

    if (ext === ".nom" || ext === ".json") {
      let datos;
      try { datos = JSON.parse(texto); } catch (e) { throw new Error("El archivo .nom está dañado"); }
      if (datos.formato !== "nom-colab" || !datos.bloques) throw new Error("Ese JSON no es un documento de NOM");
      const d = datos.bloques._doc || {};
      const ops = Object.keys(datos.bloques).filter((k) => k !== "_doc").map((k, i) => {
        const { id, ...c } = datos.bloques[k];
        return { id: "imp:" + i, bloque: k, campos: c, t: 2, n: "importar" };
      });
      if (d.paginas) ops.unshift({ id: "imp:p", bloque: "_doc", campos: { paginas: d.paginas }, t: 2, n: "importar" });
      // Todos los datos generales (antes solo cuatro: se perdían fondos, notas del orador, tema…).
      Object.keys(d).forEach((k) => { if (!["id", "nombre", "modo", "paginas"].includes(k) && d[k] != null) ops.push({ id: "imp:d" + k, bloque: "_doc", campos: { [k]: d[k] }, t: 2, n: "importar" }); });
      // Texto letra por letra: se recrean sus operaciones con los mismos ids.
      let i = 0;
      for (const b in datos.seq || {}) {
        const r = datos.seq[b];
        ["ins", "vis", "fmt"].forEach((k) => (r[k] || []).forEach((e) => {
          const { t, n, ...s } = e;
          if (!s.k) s.k = k === "ins" ? "i" : k === "vis" ? "v" : "f";
          ops.push({ id: "imp:s" + i++, bloque: b, campos: {}, t, n, s });
        }));
      }
      return { nombre: d.nombre || nombre, modo: d.modo || "TEXTO", ops, aviso };
    }
    if (ext === ".csv" || ext === ".tsv") {
      const sep = ext === ".tsv" ? "\t" : (texto.split("\n")[0].split(";").length > texto.split("\n")[0].split(",").length ? ";" : ",");
      const filas = parseCsv(texto, sep).filter((r) => r.some((c) => c !== ""));
      const b = { id: NOM.DocCRDT.nuevoId(), tipo: "tabla", pagina: "p1", orden: 1, filas: Math.max(30, filas.length + 5), cols: Math.max(10, ...filas.map((r) => r.length)), x: 40, y: 40, w: 880, h: 300 };
      filas.forEach((r, f) => r.forEach((v, c) => { if (v !== "") b["c:" + F().ref(c, f)] = v; }));
      return { nombre, modo: "HOJA", ops: comoOps([b]), aviso };
    }
    if (ext === ".html" || ext === ".htm") {
      const d = new DOMParser().parseFromString(texto, "text/html");
      const lineas = [];
      d.body.querySelectorAll("h1,h2,h3,h4,p,li,blockquote").forEach((n) => {
        const t = n.textContent.trim(); if (!t) return;
        const pref = { H1: "# ", H2: "## ", H3: "## ", H4: "## ", LI: "- ", BLOCKQUOTE: "> " }[n.tagName] || "";
        lineas.push(pref + t);
      });
      if (d.body.querySelector("img,table,svg")) aviso = "Las tablas e imágenes del HTML no se importaron";
      const r = bloquesTexto(lineas.length ? lineas : [d.body.textContent.trim()]);
      return { nombre: d.title || nombre, modo: "TEXTO", ops: comoOps(r.bs), aviso };
    }
    // .md / .txt
    const lineas = texto.replace(/\r\n?/g, "\n").split("\n").filter((l) => l.trim() !== "");
    const r = bloquesTexto(lineas.length ? lineas : [""]);
    if (r.aproximado) aviso = "Algunas tablas o imágenes de Markdown quedaron como texto";
    return { nombre, modo: "TEXTO", ops: comoOps(r.bs), aviso };
  }

  /** Word (.docx) → documento de texto: estilos, listas, formato de letra, tablas, imágenes y página. */
  async function importarDocx(archivo, nombre) {
    const r = await NOM.ooxml.leerDocx(archivo);
    const ops = [], ps = [];
    let i = 0;
    const limpio = (a) => { const o = {}; for (const k in a || {}) if (a[k] != null && a[k] !== false && a[k] !== "") o[k] = a[k]; return Object.keys(o).length ? o : undefined; };
    r.parrafos.forEach((p) => {
      const tramos = p.tramos.map((t) => {
        if (!t.e) return { x: t.x, a: limpio(t.a) };
        // ids fijos (no al azar): al abrir un .nom se vuelve a importar el original y tienen que salir los mismos
        const ref = "impw" + (i++).toString(36);
        if (t.e.tipo === "imagen") {
          // (con cómo flota en Word: libre en su lugar, a un lado del texto o en su propio renglón)
          const fl = {}; ["ajuste", "libre", "x", "y", "detras", "alin", "dx", "z", "fl"].forEach((k) => { if (t.e[k] != null) fl[k] = t.e[k]; });
          ops.push({ id: "imp:b" + ref, bloque: ref, campos: Object.assign({ tipo: "imagen", enTexto: true, src: t.e.src, w: Math.min(t.e.w || 320, 624), h: t.e.h || 0, oe: t.e.oe || undefined, alt: t.e.alt || undefined }, fl), t: 1, n: "importar" });
          return { e: { tipo: "imagen", ref } };
        }
        if (t.e.tipo === "tabla") { ops.push({ id: "imp:b" + ref, bloque: ref, campos: Object.assign({ tipo: "tabla", enTexto: true, filasTexto: t.e.filas, colsTexto: t.e.cols }, t.e.celdas), t: 1, n: "importar" }); return { e: { tipo: "tabla", ref } }; }
        if (t.e.tipo === "nota" || t.e.tipo === "ecuacion") { ops.push({ id: "imp:b" + ref, bloque: ref, campos: Object.assign({ enTexto: true }, t.e), t: 1, n: "importar" }); return { e: { tipo: t.e.tipo, ref } }; }
        // campo de Word (su código y su último valor: «Actualizar campos» lo cambia) y marcador (donde empieza o
        // termina; solo los que usa algún campo)
        if (t.e.tipo === "campo") { ops.push({ id: "imp:b" + ref, bloque: ref, campos: limpio({ tipo: "campo", enTexto: true, instr: t.e.instr, res: t.e.res, a: limpio(t.e.a), idioma: t.e.idioma, orig: t.e.orig || undefined }), t: 1, n: "importar" }); return { e: { tipo: "campo", ref } }; }
        if (t.e.tipo === "marcador") return { e: { tipo: "marcador", nombre: t.e.nombre, lado: t.e.lado } };
        // (un gráfico de Word lleva además su modelo de Office y su tamaño, para dibujarlo como Word)
        return { e: limpio({ tipo: t.e.tipo, oe: t.e.oe, que: t.e.que, txt: t.e.txt, img: t.e.img, g: t.e.g, w: t.e.w, h: t.e.h, forma: t.e.forma, grupo: t.e.grupo, libre: t.e.libre, x: t.e.x, y: t.e.y, detras: t.e.detras, franja: t.e.franja, z: t.e.z, fl: t.e.fl }) };
      });
      ps.push({ tramos, a: limpio(p.pa) });
    });
    ops.push(...NOM.Secuencia.crearOps("cuerpo", ps, "importar", 2, "imp"));
    if (r.pagina) ops.push({ id: "imp:pag", bloque: "_doc", campos: { pagina: r.pagina }, t: 1, n: "importar" });
    // Formato base (el estilo Normal del archivo), encabezado, pie y número de página. Con reloj 2: va
    // después del formato que el servidor le pone a los documentos nuevos del espacio.
    if (r.doc && Object.keys(r.doc).length) ops.push({ id: "imp:doc", bloque: "_doc", campos: r.doc, t: 2, n: "importar" });
    return { nombre, modo: "TEXTO", ops: conOrigen("docx", nombre, "TEXTO", ops, r.archivo), aviso: avisoSinOrigen(r.aviso, r.archivo) };
  }

  /** Excel (.xlsx) → hoja de cálculo: una hoja por pestaña, con fórmulas, formato, combinadas y anchos. */
  async function importarXlsx(archivo, nombre) {
    const r = await NOM.ooxml.leerXlsx(archivo);
    const paginas = r.hojas.map((_, i) => "p" + (i + 1));
    const bloques = r.hojas.map((h, i) => Object.assign({ id: "hoja-" + paginas[i], tipo: "tabla", pagina: paginas[i], orden: 0, x: 40, y: 40, w: 880, h: 300 }, h.campos));
    const ops = comoOps(bloques, paginas);
    // (lo del libro: estructura protegida, contraseña para modificar, marcado como final)
    if (r.doc) ops.push(...NOM.DocCRDT.partir([{ id: "imp:doc", bloque: "_doc", campos: r.doc, t: 2, n: "importar" }]));
    return { nombre, modo: "HOJA", ops: conOrigen("xlsx", nombre, "HOJA", ops, r.archivo), aviso: avisoSinOrigen(r.aviso, r.archivo) };
  }
  /** PowerPoint (.pptx) → presentación (lo arma ooxml.leerPptx). */
  async function importarPptx(archivo, nombre) {
    if (!NOM.ooxml.leerPptx) throw new Error("Este formato todavía no se puede abrir");
    const r = await NOM.ooxml.leerPptx(archivo);
    const ops = comoOps(r.bloques, r.paginas);
    // Tamaño, fondos, transiciones y notas del orador van en los datos del documento.
    if (r.doc && Object.keys(r.doc).length) ops.push(...NOM.DocCRDT.partir([{ id: "imp:d", bloque: "_doc", campos: r.doc, t: 2, n: "importar" }]));
    return { nombre, modo: "PRESENTACION", ops: conOrigen("pptx", nombre, "PRESENTACION", ops, r.archivo), aviso: avisoSinOrigen(r.aviso, r.archivo) };
  }

  /** Lo más grande que se inserta como imagen: por peso y por tamaño en píxeles (lo que ocupa abierta en memoria). */
  const MAX_IMAGEN_BYTES = 50 * 1024 * 1024, MAX_IMAGEN_PIXELES = 120e6;

  /** Imagen local -> data URL comprimida (máx. 1600 px, JPEG 0.85) para no inflar el documento. */
  function imagenADataUrl(archivo) {
    return new Promise((ok, mal) => {
      if (!/^image\//.test(archivo.type)) return mal(new Error("Elige una imagen (PNG, JPG, GIF, WebP)"));
      // (antes de leerla: una foto enorme llenaría la memoria del celular y cerraría la página)
      if (archivo.size > MAX_IMAGEN_BYTES) return mal(new Error(`La imagen pesa ${Math.round(archivo.size / 1048576)} MB y el máximo es 50 MB: redúcela antes de insertarla`));
      // 8.3 Los GIF se guardan tal cual (si se pasan por el canvas, dejan de moverse).
      const gif = archivo.type === "image/gif" && archivo.size <= 4 * 1024 * 1024;
      // Se abre desde el archivo (sin copiarla a texto antes) y se reduce en un lienzo.
      const url = URL.createObjectURL(archivo);
      const img = new Image();
      img.onload = () => {
        if (img.width * img.height > MAX_IMAGEN_PIXELES) {
          URL.revokeObjectURL(url);
          return mal(new Error(`La imagen mide ${img.width} × ${img.height} píxeles; el máximo es 120 megapíxeles: redúcela antes de insertarla`));
        }
        if (gif) {
          URL.revokeObjectURL(url);
          const lector = new FileReader();
          lector.onerror = () => mal(new Error("No se pudo leer la imagen"));
          lector.onload = () => ok({ src: lector.result, w: img.width, h: img.height });
          lector.readAsDataURL(archivo);
          return;
        }
        try {
          const k = Math.min(1, 1600 / Math.max(img.width, img.height));
          const cv = document.createElement("canvas");
          cv.width = Math.max(1, Math.round(img.width * k)); cv.height = Math.max(1, Math.round(img.height * k));
          cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
          const png = archivo.type === "image/png" && cv.width * cv.height < 400000;
          ok({ src: cv.toDataURL(png ? "image/png" : "image/jpeg", 0.85), w: cv.width, h: cv.height });
        } catch (e) { mal(new Error("No se pudo preparar la imagen")); }
        finally { URL.revokeObjectURL(url); }
      };
      img.onerror = () => { URL.revokeObjectURL(url); mal(new Error("La imagen está dañada o su formato no se puede abrir aquí")); };
      img.src = url;
    });
  }

  function elegirArchivo(aceptar) {
    return new Promise((ok) => {
      const i = Object.assign(document.createElement("input"), { type: "file", accept: aceptar });
      i.onchange = () => ok(i.files[0] || null);
      i.click();
    });
  }

  NOM.formatos = { IMPORTABLES, ACEPTAR, EXPORTABLES, paraTipo, exportar, exportarListo, importar, huellasOrigen, aNom, aNomBytes, deNomBytes, fotoModelo, docDeOps, huellaBloque, imagenADataUrl, elegirArchivo, aHtml, aMarkdown, aSvg, textoPlano, descargar, descargarBytes, MIME, limpioNombre };
})();

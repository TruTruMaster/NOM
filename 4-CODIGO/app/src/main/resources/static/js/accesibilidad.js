/* NOM · Revisar → Comprobar accesibilidad, como el de Office: errores (falta texto alternativo, una tabla sin fila de
   encabezado, una diapositiva sin título), advertencias (contraste bajo, vínculos que no dicen a dónde llevan, celdas
   combinadas, objetos que no van en línea con el texto, espacios o renglones vacíos para acomodar) y consejos (títulos que
   se saltan un nivel, diapositivas con el mismo título, hojas con el nombre de siempre). Las reglas son las del
   comprobador de Word, Excel y PowerPoint (support.microsoft.com, «Reglas del Comprobador de accesibilidad»). */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});

  // ------------------------------------------------------------ contraste (WCAG 2, el mismo que usa Office)
  function rgb(c) {
    const m = /^#?([0-9a-f]{6})$/i.exec(String(c || "").trim());
    if (!m) return null;
    const n = parseInt(m[1], 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  function luz([r, g, b]) {
    const f = (v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b);
  }
  /** La razón de contraste entre dos colores (1 a 21), o null si alguno no se entiende. */
  function contraste(a, b) {
    const x = rgb(a), y = rgb(b);
    if (!x || !y) return null;
    const l1 = luz(x), l2 = luz(y);
    return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
  }
  /** ¿Se lee mal? Como Office: menos de 4.5 a 1 (3 a 1 en letra grande: 18 pt, o 14 pt en negrita). */
  const pocoContraste = (color, fondo, pt, negrita) => {
    const r = contraste(color, fondo || "#FFFFFF");
    return r != null && r < ((+pt || 11) >= 18 || ((+pt || 11) >= 14 && negrita) ? 3 : 4.5) ? Math.round(r * 100) / 100 : null;
  };
  // vínculos que no dicen a dónde llevan
  const VINCULO_VAGO = /^(aqu[ií]|clic aqu[ií]|haz clic aqu[ií]|da clic aqu[ií]|click here|here|m[aá]s|leer m[aá]s|ver m[aá]s|read more|more|link|enlace|v[ií]nculo|este enlace|este v[ií]nculo|https?:\/\/\S+|www\.\S+)\.?$/i;
  const recorta = (s, n = 50) => { s = String(s || "").replace(/\s+/g, " ").trim(); return s.length > n ? s.slice(0, n - 1) + "…" : s; };

  // ------------------------------------------------------------ texto (Word)
  function revisarTexto(doc) {
    const out = [], ps = NOM.textoRico ? NOM.textoRico.parrafosDe(doc) : [];
    let nivelAnt = 0, vacios = 0;
    const NIVEL = { h1: 1, h2: 2, h3: 3 };
    ps.forEach((p, i) => {
      const texto = p.tramos.map((t) => t.x || "").join(""), donde = recorta(texto) || `Párrafo ${i + 1}`;
      // títulos: no saltarse niveles
      const nv = Object.prototype.hasOwnProperty.call(NIVEL, p.pa.estilo) ? NIVEL[p.pa.estilo] : 0;
      if (nv) { if (nivelAnt && nv > nivelAnt + 1) out.push({ nivel: "consejo", regla: "Se salta un nivel de título", donde, ir: { parrafo: i } }); nivelAnt = nv; }
      // renglones vacíos seguidos para acomodar
      if (!texto.trim() && !p.tramos.some((t) => t.e)) { vacios++; if (vacios === 3) out.push({ nivel: "advertencia", regla: "Caracteres en blanco repetidos (renglones vacíos para acomodar)", donde: `Párrafo ${i + 1}`, ir: { parrafo: i } }); }
      else vacios = 0;
      if (/\S {3,}\S/.test(texto)) out.push({ nivel: "advertencia", regla: "Caracteres en blanco repetidos (espacios para acomodar)", donde, ir: { parrafo: i } });
      p.tramos.forEach((t) => {
        if (t.e) {
          const e = t.e, tipo = e.tipo;
          if (["imagen", "forma", "grafica", "smartart", "dibujo"].includes(tipo) && !String(e.alt || "").trim() && !e.decorativo)
            out.push({ nivel: "error", regla: "Falta texto alternativo", donde: `${tipo === "imagen" ? "Imagen" : tipo === "grafica" ? "Gráfico" : "Objeto"} en «${recorta(texto, 30) || "párrafo " + (i + 1)}»`, ir: { parrafo: i } });
          if (tipo === "imagen" && (e.fl || e.libre || e.ajuste)) out.push({ nivel: "advertencia", regla: "Objeto que no está en línea con el texto", donde: `Imagen en «${recorta(texto, 30) || "párrafo " + (i + 1)}»`, ir: { parrafo: i } });
          // (como Word: una fila marcada «Repetir como fila de encabezado», o el encabezado de NOM)
          if (tipo === "tabla" && !e.encabezado && !(+e.cabRep > 0)) out.push({ nivel: "error", regla: "Falta el encabezado de la tabla", donde: `Tabla del párrafo ${i + 1}`, ir: { parrafo: i } });
          return;
        }
        const a = t.a || {};
        if (a.enlace && VINCULO_VAGO.test(String(t.x || "").trim())) out.push({ nivel: "advertencia", regla: "Texto de vínculo poco claro", donde: `«${recorta(t.x, 40)}»`, ir: { parrafo: i } });
        const r = a.color && (t.x || "").trim() ? pocoContraste(a.color, a.fondo, a.tam, a.b) : null;
        if (r) out.push({ nivel: "advertencia", regla: `Texto con poco contraste (${r} a 1)`, donde: `«${recorta(t.x, 40)}»`, ir: { parrafo: i } });
      });
    });
    return out;
  }

  // ------------------------------------------------------------ hoja (Excel)
  function revisarHoja(doc) {
    const F = NOM.formulas, out = [];
    F.hojasDe(doc).forEach((h, k) => {
      const nombre = F.nombreDeHoja(doc, h);
      if (/^(Hoja|Sheet)\d+$/i.test(nombre)) out.push({ nivel: "consejo", regla: "Nombre de hoja predeterminado", donde: `«${nombre}»`, ir: { pagina: h.pagina } });
      const comb = Object.keys(h).filter((x) => x.startsWith("m:") && h[x]).map((x) => x.slice(2));
      if (comb.length) out.push({ nivel: "advertencia", regla: "Celdas combinadas", donde: `${nombre}: ${comb.slice(0, 6).join(", ")}${comb.length > 6 ? ` y ${comb.length - 6} más` : ""}`, ir: { pagina: h.pagina, ref: comb[0] } });
      const bajos = [];
      for (const x of Object.keys(h)) {
        if (!x.startsWith("k:")) continue;
        const ref = x.slice(2);
        if (h["c:" + ref] == null || h["c:" + ref] === "") continue;
        const r = pocoContraste(h[x], h["r:" + ref], h["tp:" + ref] || h["tl:" + ref], h["n:" + ref]);
        if (r) bajos.push(ref);
      }
      if (bajos.length) out.push({ nivel: "advertencia", regla: "Texto con poco contraste", donde: `${nombre}: ${bajos.slice(0, 6).join(", ")}${bajos.length > 6 ? ` y ${bajos.length - 6} más` : ""}`, ir: { pagina: h.pagina, ref: bajos[0] } });
      (doc.deLaPagina ? doc.deLaPagina(h.pagina) : []).forEach((b) => {
        if (["imagen", "grafica", "forma"].includes(b.tipo) && !String(b.alt || "").trim() && !(b.tipo === "forma" && String(b.texto || "").trim()) && !b.decorativo)
          out.push({ nivel: "error", regla: "Falta texto alternativo", donde: `${b.tipo === "grafica" ? "Gráfico" : b.tipo === "imagen" ? "Imagen" : "Forma"} en «${nombre}»`, ir: { pagina: h.pagina } });
      });
      void k;
    });
    return out;
  }

  // ------------------------------------------------------------ diapositivas (PowerPoint)
  function revisarDiapositivas(doc) {
    const out = [], pags = doc.paginas(), titulos = new Map();
    const plano = (b) => String(b.texto || String(b.html || "").replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim();
    pags.forEach((pg, i) => {
      const bs = doc.deLaPagina(pg), n = `Diapositiva ${i + 1}`;
      const tit = bs.filter((b) => b.tipo === "texto" && /^h[12]$/.test(b.estilo || "") && plano(b)).sort((a, c) => (a.y || 0) - (c.y || 0))[0];
      if (!tit) out.push({ nivel: "error", regla: "Falta el título de la diapositiva", donde: n, ir: { pagina: pg } });
      else { const t = plano(tit).toLowerCase(); (titulos.get(t) || titulos.set(t, []).get(t)).push(i + 1); }
      bs.forEach((b) => {
        const sinAlt = !String(b.alt || "").trim() && !b.decorativo;
        if (b.tipo === "imagen" && sinAlt) out.push({ nivel: "error", regla: "Falta texto alternativo", donde: `Imagen de la ${n.toLowerCase()}`, ir: { pagina: pg, id: b.id } });
        if (["grafica", "smartart", "video", "audio"].includes(b.tipo) && sinAlt) out.push({ nivel: "error", regla: "Falta texto alternativo", donde: `${b.tipo === "grafica" ? "Gráfico" : b.tipo === "smartart" ? "SmartArt" : b.tipo === "video" ? "Video" : "Audio"} de la ${n.toLowerCase()}`, ir: { pagina: pg, id: b.id } });
        if (b.tipo === "forma" && sinAlt && !plano(b)) out.push({ nivel: "error", regla: "Falta texto alternativo", donde: `Forma de la ${n.toLowerCase()}`, ir: { pagina: pg, id: b.id } });
        if (b.tipo === "tabla" && !b.encabezado) out.push({ nivel: "error", regla: "Falta el encabezado de la tabla", donde: `Tabla de la ${n.toLowerCase()}`, ir: { pagina: pg, id: b.id } });
        if (b.tipo === "texto" && b.color && plano(b)) { const r = pocoContraste(b.color, b.fondo || (doc.valor("_doc", "fondo") || {}).color, b.tam, b.negrita); if (r) out.push({ nivel: "advertencia", regla: `Texto con poco contraste (${r} a 1)`, donde: `«${recorta(plano(b), 40)}» (${n.toLowerCase()})`, ir: { pagina: pg, id: b.id } }); }
      });
    });
    titulos.forEach((nums, t) => { if (nums.length > 1) out.push({ nivel: "consejo", regla: "Diapositivas con el mismo título", donde: `«${recorta(t, 40)}»: ${nums.join(", ")}`, ir: { pagina: pags[nums[1] - 1] } }); });
    return out;
  }

  /** Lo que encuentra en el documento: [{ nivel: "error"|"advertencia"|"consejo", regla, donde, ir }], en ese orden. */
  function revisar(doc) {
    const modo = String(doc.valor("_doc", "modo") || "TEXTO").toUpperCase();
    const r = modo === "HOJA" ? revisarHoja(doc) : modo === "TEXTO" ? revisarTexto(doc) : revisarDiapositivas(doc);
    const ORDEN = { error: 0, advertencia: 1, consejo: 2 };
    return r.map((x, i) => [x, i]).sort((a, b) => ORDEN[a[0].nivel] - ORDEN[b[0].nivel] || a[1] - b[1]).map((x) => x[0]);
  }

  /** Lleva a donde está lo encontrado (el párrafo, la celda o la diapositiva). */
  function ir(ed, lugar) {
    if (!lugar) return;
    if (lugar.parrafo != null) {
      const el = document.querySelectorAll(".wd-cuerpo > *")[lugar.parrafo];
      if (el) { el.scrollIntoView({ block: "center", behavior: "smooth" }); el.classList.add("acc-marca"); setTimeout(() => el.classList.remove("acc-marca"), 1600); }
      return;
    }
    const api = ed.modoApi && ed.modoApi.api;
    if (api && api.irA && lugar.pagina) { api.irA({ pagina: lugar.pagina, ref: lugar.ref || "A1" }); return; }
    if (lugar.pagina) {
      const m = document.querySelector(`.mini[data-p="${CSS.escape(lugar.pagina)}"]`);
      if (m) m.click();
    }
  }

  const ICONO = { error: "circle-x", advertencia: "triangle-alert", consejo: "lightbulb" }, TITULO = { error: "Errores", advertencia: "Advertencias", consejo: "Consejos" };
  /** El panel de Comprobar accesibilidad. */
  function abrir(ed) {
    const { html, esc, ico } = NOM;
    const cuerpo = html(`<div class="col" style="gap:10px" data-lista></div>`);
    const pie = html(`<div style="display:contents"><button class="btn" data-otra type="button">Volver a revisar</button><span class="crece"></span><button class="btn primario" data-no type="button">Cerrar</button></div>`);
    const h = NOM.hoja({ titulo: "Comprobar accesibilidad", icono: "accessibility", cuerpo, pie, ancha: true });
    let lista = [];
    const pintar = () => {
      lista = revisar(ed.doc);
      if (!lista.length) { cuerpo.innerHTML = `<p>${ico("check", "ico-s")} No se encontraron problemas de accesibilidad: las personas con discapacidad no deberían tener dificultad para leer este documento.</p>`; return; }
      cuerpo.innerHTML = ["error", "advertencia", "consejo"].map((nv) => {
        const xs = lista.map((x, i) => [x, i]).filter(([x]) => x.nivel === nv);
        if (!xs.length) return "";
        return `<section class="col" style="gap:4px"><b>${ico(ICONO[nv], "ico-s")} ${TITULO[nv]} (${xs.length})</b>${xs.map(([x, i]) => `<div class="lista-item fila" style="gap:8px;align-items:center"><span class="crece">${esc(x.regla)}<br><span class="suave chico">${esc(x.donde)}</span></span>${x.ir ? `<button class="btn chico" type="button" data-ir="${i}">Ir</button>` : ""}</div>`).join("")}</section>`;
      }).join("") + `<p class="suave chico">Las reglas son las del Comprobador de accesibilidad de Office. El texto alternativo se pone con el menú de cada imagen u objeto.</p>`;
      cuerpo.querySelectorAll("[data-ir]").forEach((b) => (b.onclick = () => ir(ed, lista[+b.dataset.ir].ir)));
    };
    pie.querySelector("[data-otra]").onclick = pintar;
    pie.querySelector("[data-no]").onclick = () => h.cerrar();
    pintar();
    return lista;
  }

  NOM.accesibilidad = { revisar, abrir, contraste };
})();

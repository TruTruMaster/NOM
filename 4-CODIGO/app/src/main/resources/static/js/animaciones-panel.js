/* NOM — crear y editar animaciones como en PowerPoint: la galería de efectos (los 262 de PowerPoint 2024, por
 * Entrada, Énfasis, Salida y Trayectorias), las opciones de cada efecto (dirección, por letra o palabra, por
 * párrafos, inicio y fin suaves, rebote, después de la animación), los intervalos (inicio, retraso, duración,
 * repetir, rebobinar, desencadenador), el panel de animación con su línea de tiempo, copiar la animación de un bloque
 * a otro y las trayectorias dibujadas a mano.
 *
 * Todo se guarda en la lista "anim:<página>" con el árbol de PowerPoint de cada efecto (animaciones-office.js), así
 * que lo que se arma aquí se reproduce con el mismo motor y sale igual en el .pptx.
 *
 * El modo Presentación (lienzo.js) le pasa un contexto:
 *   { doc(), pagina(), seleccion(), soloLectura(), lista(), guardar(l), nombreBloque(id), bloques() → [{ id, nombre }],
 *     T(), estiloDiapo(), htmlDiapo(), motor(el, lista, escala), bloqueEnEditor(id) → elemento .bl del editor }
 */
(function () {
  "use strict";
  const NOM = window.NOM;
  const { esc, ico, html, $, $$, aviso } = NOM;
  const AO = () => NOM.animOffice;
  const datos = () => NOM.animOfficeDatos || { presets: {}, categorias: {} };

  const CLASES = [["entr", "Entrada", "e"], ["emph", "Énfasis", "n"], ["exit", "Salida", "x"], ["path", "Trayectorias", "r"]];
  const TIPO = { e: "Entrada", n: "Énfasis", x: "Salida", r: "Trayectoria", m: "Medio" };
  const CLASE_DE_TIPO = { e: "entr", n: "emph", x: "exit", r: "path" };
  const INICIOS = [["clic", "Al hacer clic"], ["con", "Con la anterior"], ["despues", "Después de la anterior"]];
  const REPETIR = [["1", "(ninguna)"], ["2", "2"], ["3", "3"], ["4", "4"], ["5", "5"], ["10", "10"], ["clic", "Hasta el siguiente clic"], ["fin", "Hasta el final de la diapositiva"]];

  /** Marca de color por tipo (como las estrellas de PowerPoint: verde, amarillo, rojo; azul para trayectorias). */
  const marca = (t) => `<span class="ao-marca ao-${esc(t || "n")}" aria-hidden="true">${ico(t === "r" ? "route" : "star", "ico-s")}</span>`;
  const nombreDe = (x) => (AO() && AO().nombreEfecto(x)) || "Efecto";
  const seg = (ms) => Math.round((+ms || 0) / 100) / 10;
  const seg2 = (ms) => Math.round((+ms || 0) / 10) / 100;

  // ------------------------------------------------------------ catálogo
  /** Los efectos de una clase (uno por efecto, con su dirección de siempre: la primera que escribe PowerPoint). */
  function catalogo(cl) {
    const out = new Map();
    Object.entries(datos().presets || {}).forEach(([k, p]) => {
      const [c, id] = k.split(":");
      if (c !== cl) return;
      if (!out.has(id)) out.set(id, { clave: k, id: +id, n: p.n, en: p.en });
    });
    return [...out.values()];
  }
  /** Las categorías del panel de PowerPoint (las de salida son las de entrada; las trayectorias van por forma). */
  function categorias(cl) {
    if (cl === "path") {
      const r = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
      return { "Básico": r(1, 18), "Especial": r(19, 34), "Líneas y curvas": r(35, 64) };
    }
    return (datos().categorias || {})[cl === "exit" ? "entr" : cl] || {};
  }
  const sinAcentos = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();

  /**
   * La galería: elegir un efecto. opc: { titulo, clase (la pestaña abierta), actual (clave), alElegir(clave), alCerrar }
   */
  function galeria(opc) {
    let cl = opc.clase || "entr", q = "";
    const cuerpo = html(`<div class="col ao-galeria">
      <div class="fila" style="flex-wrap:wrap;gap:8px"><div class="segmentado" role="tablist">${CLASES.map(([c, t, tp]) => `<button type="button" data-cl="${c}" aria-pressed="${c === cl}">${marca(tp)} ${t}</button>`).join("")}</div>
      <input type="search" class="crece" data-q placeholder="Buscar efecto" aria-label="Buscar efecto" style="min-width:160px"></div>
      <div class="ao-cats" data-cats></div></div>`);
    let elegido = false;
    const h = NOM.hoja({ titulo: opc.titulo || "Agregar animación", icono: "star", cuerpo, ancha: true, alCerrar: () => { if (!elegido && opc.alCerrar) opc.alCerrar(); } });
    const pintar = () => {
      $$("[data-cl]", cuerpo).forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.cl === cl)));
      const efs = catalogo(cl), cats = categorias(cl), tp = CLASES.find((c) => c[0] === cl)[2];
      const qq = sinAcentos(q);
      const pasa = (e) => !qq || sinAcentos(e.n).includes(qq) || sinAcentos(e.en).includes(qq);
      const usados = new Set();
      const secciones = Object.entries(cats).map(([cat, ids]) => {
        const l = ids.map((id) => efs.find((e) => e.id === id)).filter(Boolean);
        l.forEach((e) => usados.add(e.id));
        return [cat, l.filter(pasa)];
      });
      const otros = efs.filter((e) => !usados.has(e.id));
      if (otros.length) secciones.push(["Otros", otros.filter(pasa)]);
      const actualId = opc.actual ? opc.actual.split(":").slice(0, 2).join(":") : "";
      $("[data-cats]", cuerpo).innerHTML = secciones.filter(([, l]) => l.length).map(([cat, l]) => `<section><h3 class="ao-cat">${esc(cat)}</h3><div class="ao-efs">${l.map((e) =>
        `<button type="button" class="ao-ef" data-clave="${esc(e.clave)}" aria-pressed="${cl + ":" + e.id === actualId}">${marca(tp)}<span>${esc(e.n)}</span></button>`).join("")}</div></section>`).join("")
        || `<p class="suave">Ningún efecto se llama así. Prueba con otra palabra (también sirven los nombres de PowerPoint en inglés, como «Fly» o «Wipe»).</p>`;
    };
    pintar();
    cuerpo.addEventListener("click", (e) => {
      const c = e.target.closest("[data-cl]");
      if (c) { cl = c.dataset.cl; pintar(); return; }
      const b = e.target.closest("[data-clave]");
      if (b) { elegido = true; h.cerrar(); opc.alElegir(b.dataset.clave); }
    });
    $("[data-q]", cuerpo).addEventListener("input", (e) => { q = e.target.value; pintar(); });
    return h;
  }

  // ------------------------------------------------------------ vista previa (la diapositiva en chico con el motor)
  /** Pinta la diapositiva en cont y reproduce la lista dada (todos sus clics seguidos). Devuelve { jugar, detener }. */
  function vistaPrevia(cont, ctx) {
    let motor = null, vivo = true, espera = 0;
    const T = ctx.T();
    function pintar() {
      const w = cont.clientWidth || 300, k = w / T.w;
      cont.style.height = Math.round(T.h * k) + "px";
      cont.innerHTML = `<div class="ao-prev-diapo diapo-mini" style="${ctx.estiloDiapo()};width:${T.w}px;height:${T.h}px;transform:scale(${k})"></div>`;
      const d = cont.firstChild;
      d.innerHTML = ctx.htmlDiapo();
      return { d, k };
    }
    return {
      jugar(lista) {
        this.detener();
        vivo = true;
        const { d, k } = pintar();
        motor = ctx.motor(d, lista, () => k);
        if (!motor) return;
        motor.preparar();
        const siguiente = () => {
          if (!vivo || !motor) return;
          if (!motor.quedan()) return;
          motor.clic();
          motor.alTerminar(() => { espera = setTimeout(siguiente, 300); });
        };
        motor.inicio();
        if (motor.corriendo()) motor.alTerminar(() => { espera = setTimeout(siguiente, 300); });
        else siguiente();
      },
      quieta() { this.detener(); pintar(); },
      detener() { vivo = false; clearTimeout(espera); if (motor) motor.detener(); motor = null; },
    };
  }

  // ------------------------------------------------------------ unidades de la lista (un efecto por párrafos es uno)
  /** La lista en unidades: [{ desde, hasta (índices en la lista), x (el primero) }]. */
  function unidades(l) {
    const us = [];
    l.forEach((x, i) => {
      const u = us[us.length - 1];
      if (u && x.grupoP && l[u.hasta].grupoP === x.grupoP) u.hasta = i;
      else us.push({ desde: i, hasta: i, x });
    });
    return us;
  }
  /** Los párrafos con texto de un bloque (sus índices), con el bloque pintado en la vista previa. */
  function parrafosDe(el) {
    const A = AO();
    if (!el || !A) return [];
    return A.parrafosDe(A.contTexto(el)).map((p, i) => [i, (p.textContent || "").trim()]).filter(([, t]) => t).map(([i]) => i);
  }
  /** Tiempos de cada efecto dentro de su clic (como el panel de PowerPoint): [{ t, d (con repeticiones) }]. */
  function tiempos(l) {
    const out = [];
    let ini = 0, fin = 0;
    const A = AO();
    l.forEach((x) => {
      const dur = Math.max(1, +x.dur || 500), ret = Math.max(0, +x.retraso || 0);
      const op = A ? A.opcionesDe(x) : null;
      const rep = op ? (typeof op.repetir === "number" ? op.repetir : 1) * (op.autoRev ? 2 : 1) : 1;
      let t;
      if (x.disp && !out.some((o, j) => l[j].disp === x.disp)) { ini = 0; fin = 0; }
      if ((x.inicio || "clic") === "clic") { t = ret; fin = 0; }
      else if (x.inicio === "con") t = ini + ret;
      else t = fin + ret;
      const d = dur * rep, inf = op && (op.repetir === "clic" || op.repetir === "fin");
      out.push({ t, d, inf });
      ini = t; fin = Math.max(fin, t + d);
    });
    return out;
  }

  // ------------------------------------------------------------ opciones de un efecto
  /**
   * El cuadro de opciones del efecto en la posición idx de la lista (si es por párrafos, el grupo entero).
   * alCerrar(): lo que se abre después (el panel, si de ahí vino).
   */
  function opciones(ctx, idx, alCerrar) {
    const A = AO();
    let l = ctx.lista();
    const u = unidades(l).find((q) => idx >= q.desde && idx <= q.hasta);
    if (!u || !A) return;
    // (un efecto por párrafos se edita como uno; uno que anima un solo párrafo se queda con su párrafo)
    let x = l[u.desde].grupoP ? A.comoObjeto(l[u.desde]) : l[u.desde];
    const parAntes = l[u.desde].grupoP ? { atras: !!l[u.desde].parAtras, fondo: l.slice(u.desde, u.hasta + 1).some((q) => q.par === -1) } : null;
    let op = A.opcionesDe(x) || {};
    // los instantáneos (Aparecer, Desaparecer…) no tienen duración, como en PowerPoint
    const instantaneo = (+(A.presetDe(x) || {}).d0 || 0) <= 1;
    const durS = () => (instantaneo ? Math.max(0.001, (+x.dur || 1) / 1000) : Math.max(0.01, parseFloat($("[data-f=dur]", cuerpo).value) || seg(x.dur) || 0.5));
    const vars = A.variantes(x);
    const k0 = A.clavePreset(x) || "";
    const bloques = ctx.bloques();
    const conTexto = ctx.tieneTexto(x.id), esGraf = !!(ctx.esGrafica && ctx.esGrafica(x.id));
    const grafAntes = l[u.desde].graf ? { modo: l[u.desde].graf, fondo: !!l[u.desde].grafFondo } : null;
    const colorHex = (c) => (/^#[0-9a-f]{6}$/i.test(c || "") ? c : "#C8C8C8");
    const desp = !op.despues ? "" : op.despues.color ? "color" : op.despues;
    const cuerpo = html(`<div class="col ao-opc">
      <div class="ao-prev" data-prev aria-label="Vista previa"></div>
      <div class="fila" style="align-items:center;gap:8px">${marca(x.tipo)}<b class="crece recorta" data-nombre>${esc(nombreDe(x))}</b>
        <button type="button" class="btn chico" data-cambiar>${ico("replace", "ico-s")} Cambiar efecto</button>
        <button type="button" class="btn chico" data-ver>${ico("play", "ico-s")} Vista previa</button></div>
      <div class="segmentado lleno" role="tablist"><button type="button" data-tab="efecto" aria-pressed="true">Efecto</button><button type="button" data-tab="tiempo" aria-pressed="false">Intervalos</button>${conTexto ? `<button type="button" data-tab="texto" aria-pressed="false">Texto</button>` : ""}${esGraf ? `<button type="button" data-tab="grafica" aria-pressed="false">Gráfica</button>` : ""}</div>
      <div class="ao-tab" data-pan="efecto">
        ${vars.length > 1 ? `<label class="ao-campo"><span>Dirección</span><select data-f="dir">${vars.map((v) => `<option value="${esc(v.clave)}" ${v.clave === k0 ? "selected" : ""}>${esc(v.dir)}</option>`).join("")}</select></label>` : ""}
        ${op.giro != null ? `<label class="ao-campo"><span>Cantidad (grados)</span><input data-f="giro" type="number" step="15" value="${Math.abs(op.giro)}"><select data-f="giroSentido"><option value="1" ${op.giro >= 0 ? "selected" : ""}>A la derecha</option><option value="-1" ${op.giro < 0 ? "selected" : ""}>A la izquierda</option></select></label>` : ""}
        ${op.escala != null ? `<label class="ao-campo"><span>Tamaño</span><input data-f="escala" type="number" min="1" max="1000" step="10" value="${op.escala}"> %</label>` : ""}
        ${op.color != null ? `<label class="ao-campo"><span>Color</span><input data-f="color" type="color" value="${colorHex(op.color)}"></label>` : ""}
        <label class="ao-campo"><span>Inicio suave</span><input data-f="suaveIni" type="number" min="0" step="0.05" value="${seg2(op.suaveIni * (x.dur || 500))}"> s</label>
        <label class="ao-campo"><span>Fin suave</span><input data-f="suaveFin" type="number" min="0" step="0.05" value="${seg2(op.suaveFin * (x.dur || 500))}"> s</label>
        <label class="ao-campo"><span>Rebote al final</span><input data-f="rebote" type="number" min="0" step="0.05" value="${seg2(op.rebote * (x.dur || 500))}"> s</label>
        <label class="ao-check"><input data-f="autoRev" type="checkbox" ${op.autoRev ? "checked" : ""}> Revertir automáticamente (ida y vuelta)</label>
        <label class="ao-campo"><span>Después de la animación</span><select data-f="despues"><option value="">No atenuar</option><option value="color" ${desp === "color" ? "selected" : ""}>Atenuar a un color</option><option value="ocultar" ${desp === "ocultar" ? "selected" : ""}>Ocultar después de la animación</option><option value="ocultarClic" ${desp === "ocultarClic" ? "selected" : ""}>Ocultar con el siguiente clic</option></select>
          <input data-f="despColor" type="color" value="${colorHex(op.despues && op.despues.color)}" ${desp === "color" ? "" : "hidden"} aria-label="Color para atenuar"></label>
      </div>
      <div class="ao-tab" data-pan="tiempo" hidden>
        <label class="ao-campo"><span>Inicio</span><select data-f="inicio">${INICIOS.map(([v, t]) => `<option value="${v}" ${(x.inicio || "clic") === v ? "selected" : ""}>${t}</option>`).join("")}</select></label>
        <label class="ao-campo"><span>Retraso</span><input data-f="retraso" type="number" min="0" max="60" step="0.25" value="${seg(x.retraso)}"> s</label>
        <label class="ao-campo"><span>Duración</span><input data-f="dur" type="number" min="0.01" max="60" step="0.25" value="${instantaneo ? "" : seg2(x.dur || 500)}" ${instantaneo ? 'disabled placeholder="—"' : ""}> s</label>
        <label class="ao-campo"><span>Repetir</span><select data-f="repetir">${REPETIR.map(([v, t]) => `<option value="${v}" ${String(op.repetir) === v ? "selected" : ""}>${t}</option>`).join("")}</select></label>
        <label class="ao-check"><input data-f="rebobinar" type="checkbox" ${op.rebobinar ? "checked" : ""}> Rebobinar al terminar (vuelve a como estaba)</label>
        <fieldset class="ao-campo ao-disp"><legend>Desencadenadores</legend>
          <label class="ao-check"><input type="radio" name="ao-disp" value="" ${x.disp ? "" : "checked"}> Como parte de la secuencia de clics</label>
          <label class="ao-check"><input type="radio" name="ao-disp" value="b" ${x.disp ? "checked" : ""}> Iniciar el efecto al hacer clic en:</label>
          <select data-f="disp" aria-label="Figura que inicia el efecto">${bloques.map((b) => `<option value="${esc(b.id)}" ${x.disp === b.id ? "selected" : ""}>${esc(b.nombre)}</option>`).join("")}</select></fieldset>
      </div>
      ${conTexto ? `<div class="ao-tab" data-pan="texto" hidden>
        <label class="ao-campo"><span>Animar texto</span><select data-f="texto"><option value="todo">Todo a la vez</option><option value="palabra" ${op.texto === "palabra" ? "selected" : ""}>Por palabra</option><option value="letra" ${op.texto === "letra" ? "selected" : ""}>Por letra</option></select></label>
        <label class="ao-campo"><span>Retraso entre ${op.texto === "palabra" ? "palabras" : "letras"}</span><input data-f="pct" type="number" min="0" max="100" step="1" value="${op.pct}"> % de la duración</label>
        <label class="ao-check"><input data-f="atras" type="checkbox" ${op.atras ? "checked" : ""}> Letras o palabras en orden inverso</label>
        <label class="ao-campo"><span>Agrupar texto</span><select data-f="agrupar"><option value="obj">Como un objeto</option><option value="p" ${parAntes ? "selected" : ""}>Por párrafos</option></select></label>
        <label class="ao-check"><input data-f="parAtras" type="checkbox" ${parAntes && parAntes.atras ? "checked" : ""}> Párrafos en orden inverso</label>
        <label class="ao-check"><input data-f="parFondo" type="checkbox" ${parAntes && parAntes.fondo ? "checked" : ""}> Animar la forma adjunta (la figura antes que su texto)</label>
      </div>` : ""}
      ${esGraf ? `<div class="ao-tab" data-pan="grafica" hidden>
        <label class="ao-campo"><span>Agrupar gráfica</span><select data-f="graf"><option value="">Como un objeto</option>${[["series", "Por serie"], ["category", "Por categoría"], ["ptInSeries", "Por elemento de la serie"], ["ptInCategory", "Por elemento de la categoría"]].map(([v, t]) => `<option value="${v}" ${grafAntes && grafAntes.modo === v ? "selected" : ""}>${t}</option>`).join("")}</select></label>
        <label class="ao-check"><input data-f="grafFondo" type="checkbox" ${!grafAntes || grafAntes.fondo ? "checked" : ""}> Empezar la animación dibujando el fondo de la gráfica</label>
      </div>` : ""}
    </div>`);
    const pie = html(`<div style="display:contents"><span class="crece"></span><button class="btn" type="button" data-no>Cancelar</button><button class="btn primario" type="button" data-si>Guardar</button></div>`);
    let guardado = false;
    const prev = vistaPrevia($("[data-prev]", cuerpo), ctx);
    const h = NOM.hoja({ titulo: "Opciones del efecto · " + ctx.nombreBloque(x.id), icono: "sliders-horizontal", cuerpo, pie, ancha: true,
      alCerrar: () => { prev.detener(); if (!guardado && alCerrar) alCerrar(); } });
    const f = (n) => $(`[data-f="${n}"]`, cuerpo);
    const val = (n, def) => { const e = f(n); return e ? (e.type === "checkbox" ? e.checked : e.value) : def; };
    /** ¿La persona cambió este campo? (lo que no se toca no se reescribe en el árbol de PowerPoint) */
    const iniciales = Object.fromEntries($$("[data-f]", cuerpo).map((e) => [e.dataset.f, val(e.dataset.f)]));
    const toco = (n) => !!f(n) && val(n) !== iniciales[n];
    /** El efecto (o los efectos, por párrafos) con lo que dice el cuadro. */
    function armar() {
      let y = x;
      const dir = val("dir", null);
      if (dir && dir !== k0) y = A.cambiarEfecto(y, dir);
      const dur = Math.round(durS() * 1000);
      const fr = (n) => Math.max(0, Math.min(1, (parseFloat(val(n, "0")) || 0) / durS()));
      const cambios = {};
      if (toco("suaveIni") || toco("suaveFin")) {
        let si = fr("suaveIni"), sf = fr("suaveFin");
        if (si + sf > 1) { const kk = 1 / (si + sf); si *= kk; sf *= kk; }
        cambios.suaveIni = si; cambios.suaveFin = sf;
      }
      if (toco("rebote")) cambios.rebote = fr("rebote");
      if (toco("autoRev")) cambios.autoRev = !!val("autoRev", false);
      if (toco("rebobinar")) cambios.rebobinar = !!val("rebobinar", false);
      if (toco("repetir")) { const rp = val("repetir", "1"); cambios.repetir = rp === "clic" || rp === "fin" ? rp : Math.max(1, +rp || 1); }
      if (toco("despues") || toco("despColor")) { const dsp = val("despues", ""); cambios.despues = dsp === "color" ? { color: val("despColor", "#C8C8C8").toUpperCase() } : dsp || null; }
      if (toco("giro") || toco("giroSentido")) cambios.giro = (Math.abs(+val("giro", 0)) || 0) * (+val("giroSentido", 1) || 1);
      if (toco("escala")) cambios.escala = Math.max(1, +val("escala", 100) || 100);
      if (toco("color")) cambios.color = val("color", "#000000").toUpperCase();
      if (toco("texto") || toco("pct") || toco("atras")) Object.assign(cambios, { texto: val("texto", "todo"), pct: Math.max(0, Math.min(100, +val("pct", 10) || 0)), atras: !!val("atras", false) });
      y = A.conOpciones(y, cambios);
      y = Object.assign({}, y, { inicio: val("inicio", "clic"), retraso: Math.round((parseFloat(val("retraso", "0")) || 0) * 1000), dur });
      const radio = $('input[name="ao-disp"]:checked', cuerpo);
      if (radio && radio.value === "b" && val("disp", "")) y.disp = val("disp", ""); else delete y.disp;
      if (val("graf", "")) {
        // cuántas series y categorías tiene, contadas en la gráfica pintada (data-s y data-c de graficas-office.js)
        const sv = $(`.bl[data-id="${CSS.escape(x.id)}"] svg`, $("[data-prev]", cuerpo));
        const nS = sv ? new Set([...sv.querySelectorAll("[data-s]")].map((e) => e.getAttribute("data-s"))).size : 0;
        const nC = sv ? new Set([...sv.querySelectorAll("[data-c]")].map((e) => e.getAttribute("data-c"))).size : 0;
        if (nS) return A.porGrafica(y, val("graf", ""), nS, Math.max(1, nC), !!val("grafFondo", true));
      }
      if (val("agrupar", "obj") === "p") {
        const ps = parrafosDe($(`.bl[data-id="${CSS.escape(x.id)}"]`, $("[data-prev]", cuerpo)));
        if (ps.length > 1) return A.porParrafos(y, ps, !!val("parAtras", false), !!val("parFondo", false));
      }
      return [y];
    }
    const actualizar = () => {
      const d = f("despColor"); if (d) d.hidden = val("despues", "") !== "color";
      const disp = f("disp"); if (disp) disp.disabled = !($('input[name="ao-disp"]:checked', cuerpo) || {}).value;
    };
    actualizar();
    cuerpo.addEventListener("change", actualizar);
    cuerpo.addEventListener("click", (e) => {
      const t = e.target.closest("[data-tab]");
      if (t) {
        $$("[data-tab]", cuerpo).forEach((b) => b.setAttribute("aria-pressed", String(b === t)));
        $$("[data-pan]", cuerpo).forEach((p) => { p.hidden = p.dataset.pan !== t.dataset.tab; });
      }
      if (e.target.closest("[data-ver]")) prev.jugar(armar().map((y, i) => Object.assign({}, y, { inicio: i === 0 ? "clic" : y.inicio, disp: undefined })));
      if (e.target.closest("[data-cambiar]")) {
        // cambiar el efecto: se elige en la galería y se vuelve aquí con el nuevo (con las mismas opciones e intervalos)
        const y = armar();
        guardado = true;
        h.cerrar(true);
        galeria({ titulo: "Cambiar efecto", clase: CLASE_DE_TIPO[x.tipo] || "entr", actual: k0,
          alElegir: (clave) => {
            const l2 = ctx.lista();
            const nuevos = y.map((q) => A.cambiarEfecto(q, clave));
            l2.splice(u.desde, u.hasta - u.desde + 1, ...nuevos);
            ctx.guardar(l2);
            opciones(ctx, u.desde, alCerrar);
          },
          alCerrar: () => opciones(ctx, idx, alCerrar) });
      }
    });
    $("[data-no]", pie).onclick = () => h.cerrar();
    $("[data-si]", pie).onclick = () => {
      if (ctx.soloLectura()) { h.cerrar(); return; }
      const ys = armar();
      l = ctx.lista();
      l.splice(u.desde, u.hasta - u.desde + 1, ...ys);
      ctx.guardar(l);
      guardado = true;
      h.cerrar();
      if (alCerrar) alCerrar();
    };
    requestAnimationFrame(() => prev.quieta());
  }

  // ------------------------------------------------------------ el panel de animación
  function panel(ctx) {
    const A = AO();
    let elegida = -1;
    const cuerpo = html(`<div class="col ao-panel">
      <div class="ao-prev" data-prev aria-label="Vista previa"></div>
      <div class="fila" style="flex-wrap:wrap;gap:8px"><button type="button" class="btn chico" data-todo>${ico("play", "ico-s")} Reproducir todo</button><button type="button" class="btn chico" data-desde disabled>${ico("play", "ico-s")} Reproducir desde la elegida</button>
        <span class="crece"></span><span class="suave chico" data-escala></span></div>
      <div class="ao-lista" data-l role="list"></div></div>`);
    const pie = html(`<div style="display:contents"><button class="btn" type="button" data-agregar>${ico("plus", "ico-s")} Agregar animación</button><span class="crece"></span><button class="btn primario" type="button" data-listo>Listo</button></div>`);
    const prev = vistaPrevia($("[data-prev]", cuerpo), ctx);
    const h = NOM.hoja({ titulo: "Panel de animación", icono: "list-ordered", cuerpo, pie, ancha: true, alCerrar: () => prev.detener() });
    let pxMs = 0.06;
    function pintar() {
      const l = ctx.lista(), us = unidades(l), ts = tiempos(l);
      const largo = Math.max(2000, ...ts.map((t) => (t.inf ? t.t + 2000 : t.t + t.d)));
      const ancho = Math.max(120, Math.min(260, ($("[data-l]", cuerpo).clientWidth || 600) * 0.32));
      pxMs = ancho / largo;
      $("[data-escala]", cuerpo).textContent = `La barra de cada efecto: cuándo empieza y cuánto dura (hasta ${seg(largo)} s). Arrástrala para cambiar el retraso; su borde derecho, para la duración.`;
      let paso = 0, dispVisto = new Set();
      const filas = us.map((un, k) => {
        const x = un.x, t = ts[un.desde];
        let cab = "";
        if (x.disp && !dispVisto.has(x.disp)) { dispVisto.add(x.disp); cab = `<div class="ao-disp-cab">${ico("mouse-pointer-click", "ico-s")} Al hacer clic en ${esc(ctx.nombreBloque(x.disp))}</div>`; }
        if (!x.disp && (x.inicio || "clic") === "clic") paso++;
        const num = !x.disp && (x.inicio || "clic") === "clic" ? `<span class="ao-num">${paso}</span>` : `<span class="ao-num ao-sin" aria-hidden="true">${(x.inicio || "clic") === "con" ? "" : ico("clock", "ico-s")}</span>`;
        const n = un.hasta - un.desde + 1;
        const fin = l.slice(un.desde, un.hasta + 1).reduce((m, _, j) => Math.max(m, ts[un.desde + j].t + ts[un.desde + j].d), 0);
        return `${cab}<div class="ao-fila" role="listitem" data-u="${k}" aria-selected="${un.desde === elegida}">
          ${num}${marca(x.tipo)}<div class="ao-txt crece"><b class="recorta">${esc(ctx.nombreBloque(x.id))}</b><span class="suave chico recorta">${esc(nombreDe(x))}${n > 1 ? ` · ${x.graf ? { series: "por serie", category: "por categoría", ptInSeries: "por elemento de la serie", ptInCategory: "por elemento de la categoría" }[x.graf] || "por partes" : "por párrafos"} (${n})` : ""} · ${esc((INICIOS.find((q) => q[0] === (x.inicio || "clic")) || INICIOS[0])[1].toLowerCase())}</span></div>
          <div class="ao-tiempo" style="width:${ancho}px"><div class="ao-barra ao-${esc(x.tipo)}${t.inf ? " inf" : ""}" data-barra style="left:${Math.round(t.t * pxMs)}px;width:${Math.max(6, Math.round((t.inf ? 2000 : fin - t.t) * pxMs))}px" title="${seg(t.t)} s a ${t.inf ? "…" : seg(fin) + " s"}"><i data-borde aria-hidden="true"></i></div></div>
          <div class="ao-acc"><button type="button" class="icono-btn chico" data-mover="-1" aria-label="Subir" ${k === 0 ? "disabled" : ""}>${ico("arrow-up", "ico-s")}</button><button type="button" class="icono-btn chico" data-mover="1" aria-label="Bajar" ${k === us.length - 1 ? "disabled" : ""}>${ico("arrow-down", "ico-s")}</button><button type="button" class="icono-btn chico" data-opc aria-label="Opciones del efecto">${ico("sliders-horizontal", "ico-s")}</button><button type="button" class="icono-btn chico" data-quitar aria-label="Quitar">${ico("trash-2", "ico-s")}</button></div></div>`;
      });
      $("[data-l]", cuerpo).innerHTML = filas.join("") || `<p class="suave">Esta diapositiva no tiene animaciones. Elige un bloque en la diapositiva y toca «Agregar animación».</p>`;
      $("[data-desde]", cuerpo).disabled = elegida < 0;
      $("[data-agregar]", pie).disabled = ctx.soloLectura() || !ctx.seleccion();
      $("[data-agregar]", pie).title = ctx.seleccion() ? "" : "Primero elige un bloque en la diapositiva";
    }
    pintar();
    requestAnimationFrame(() => { prev.quieta(); pintar(); });
    const unidad = (e) => { const f = e.target.closest("[data-u]"); if (!f) return null; const l = ctx.lista(); return { l, us: unidades(l), k: +f.dataset.u }; };
    cuerpo.addEventListener("click", (e) => {
      if (e.target.closest("[data-todo]")) { prev.jugar(ctx.lista().filter((x) => !x.disp)); return; }
      if (e.target.closest("[data-desde]")) { const l = ctx.lista().slice(elegida).filter((x) => !x.disp); if (l[0]) l[0] = Object.assign({}, l[0], { inicio: "clic" }); prev.jugar(l); return; }
      const r = unidad(e);
      if (!r) return;
      const { l, us, k } = r, un = us[k];
      if (ctx.soloLectura() && !e.target.closest("[data-opc]") && e.target.closest("button")) return;
      const m = e.target.closest("[data-mover]");
      if (m) {
        const j = k + +m.dataset.mover, o = us[j];
        if (!o) return;
        const a = l.slice(un.desde, un.hasta + 1), b = l.slice(o.desde, o.hasta + 1);
        const ini = Math.min(un.desde, o.desde);
        l.splice(ini, a.length + b.length, ...(j < k ? a.concat(b) : b.concat(a)));
        elegida = j < k ? o.desde : un.desde + b.length;
        ctx.guardar(l); pintar(); return;
      }
      if (e.target.closest("[data-quitar]")) { l.splice(un.desde, un.hasta - un.desde + 1); elegida = -1; ctx.guardar(l); pintar(); return; }
      if (e.target.closest("[data-opc]")) { h.cerrar(true); opciones(ctx, un.desde, () => panel(ctx)); return; }
      elegida = un.desde;
      pintar();
      prev.jugar([Object.assign({}, un.x, { inicio: "clic", disp: undefined })].concat(l.slice(un.desde + 1, un.hasta + 1)));
    });
    cuerpo.addEventListener("dblclick", (e) => { const r = unidad(e); if (r && !e.target.closest("button")) { h.cerrar(true); opciones(ctx, r.us[r.k].desde, () => panel(ctx)); } });
    // arrastrar la barra (retraso) o su borde derecho (duración), con el mouse o el dedo
    cuerpo.addEventListener("pointerdown", (e) => {
      const b = e.target.closest("[data-barra]");
      if (!b || ctx.soloLectura()) return;
      const r = unidad(e);
      if (!r) return;
      e.preventDefault();
      const borde = !!e.target.closest("[data-borde]"), x0 = e.clientX, un = r.us[r.k];
      const it = r.l.slice(un.desde, un.hasta + 1).map((x) => ({ ret: +x.retraso || 0, dur: +x.dur || 500 }));
      b.setPointerCapture(e.pointerId);
      const paso = (ms) => Math.round(ms / 50) * 50;
      const mover = (ev) => {
        const dms = (ev.clientX - x0) / pxMs;
        r.l.slice(un.desde, un.hasta + 1).forEach((x, j) => {
          if (borde) x.dur = Math.max(50, paso(it[j].dur + dms / (un.hasta - un.desde + 1)));
          else if (j === 0) x.retraso = Math.max(0, paso(it[j].ret + dms));
        });
        const ts = tiempos(r.l), t = ts[un.desde];
        const fin = r.l.slice(un.desde, un.hasta + 1).reduce((m, _, j) => Math.max(m, ts[un.desde + j].t + ts[un.desde + j].d), 0);
        b.style.left = Math.round(t.t * pxMs) + "px";
        b.style.width = Math.max(6, Math.round((t.inf ? 2000 : fin - t.t) * pxMs)) + "px";
        b.title = borde ? `Duración ${seg(r.l[un.desde].dur)} s` : `Retraso ${seg(r.l[un.desde].retraso)} s`;
      };
      const soltar = () => { b.removeEventListener("pointermove", mover); b.removeEventListener("pointerup", soltar); b.removeEventListener("pointercancel", soltar); ctx.guardar(r.l); pintar(); };
      b.addEventListener("pointermove", mover);
      b.addEventListener("pointerup", soltar);
      b.addEventListener("pointercancel", soltar);
    });
    $("[data-listo]", pie).onclick = () => h.cerrar();
    $("[data-agregar]", pie).onclick = () => { h.cerrar(true); agregar(ctx, () => panel(ctx)); };
  }

  // ------------------------------------------------------------ agregar, copiar, pegar, quitar
  /** Agregar un efecto de la galería al bloque elegido (después de los que ya hay, al hacer clic). */
  function agregar(ctx, alCerrar, clase) {
    const id = ctx.seleccion();
    if (!id) { aviso("Elige primero un bloque de la diapositiva", { tipo: "aviso" }); return; }
    galeria({ titulo: "Agregar animación · " + ctx.nombreBloque(id), clase,
      alElegir: (clave) => {
        const y = AO().nuevoEfecto(id, clave, "clic");
        if (!y) return;
        const l = ctx.lista();
        l.push(y);
        ctx.guardar(l);
        const pasos = l.filter((x) => !x.disp && (x.inicio || "clic") === "clic").length;
        aviso(`${TIPO[y.tipo]}: ${nombreDe(y)} (clic ${pasos})`, { tipo: "ok", ms: 2200, accion: { texto: "Opciones", fn: () => opciones(ctx, l.length - 1, alCerrar) } });
        if (alCerrar) alCerrar();
      },
      alCerrar });
  }
  function copiar(ctx) {
    const id = ctx.seleccion();
    const l = ctx.lista().filter((x) => x.id === id);
    if (!l.length) { aviso("Este bloque no tiene animaciones que copiar", { tipo: "aviso" }); return; }
    // (un efecto por párrafos se copia una vez, recordando que iba por párrafos: en el otro bloque se vuelve a partir)
    const unos = unidades(l).map((u) => ({ x: l[u.desde], par: l[u.desde].grupoP ? { atras: !!l[u.desde].parAtras } : null }));
    NOM._animCopiada = unos;
    aviso(`Animación copiada (${unos.length} efecto${unos.length > 1 ? "s" : ""}). Elige otro bloque y toca «Pegar animación».`, { tipo: "ok", ms: 3000 });
  }
  function pegar(ctx) {
    const id = ctx.seleccion(), c = NOM._animCopiada;
    if (!id || !c || !c.length) return;
    const l = ctx.lista().filter((x) => x.id !== id);
    const ps = parrafosDe(ctx.bloqueEnEditor(id));
    c.forEach(({ x, par }) => {
      const y = AO().paraBloque(x, id);
      if (par && ps.length > 1) l.push(...AO().porParrafos(y, ps, par.atras)); else l.push(y);
    });
    ctx.guardar(l);
    aviso(`Animación pegada en ${ctx.nombreBloque(id)}`, { tipo: "ok", ms: 1800 });
  }
  function quitar(ctx) {
    const id = ctx.seleccion();
    ctx.guardar(ctx.lista().filter((x) => x.id !== id));
    aviso("Animaciones quitadas", { tipo: "ok", ms: 1600 });
  }

  // ------------------------------------------------------------ trayectoria dibujada
  /** Dibujar con el dedo o el mouse el camino que seguirá el bloque elegido (Esc cancela). */
  function dibujarTrayectoria(ctx) {
    const id = ctx.seleccion();
    const el = id && ctx.bloqueEnEditor(id);
    if (!el) { aviso("Elige primero un bloque de la diapositiva", { tipo: "aviso" }); return; }
    const b = ctx.doc().bloque(id) || {};
    const r0 = el.getBoundingClientRect(), k = r0.width / (b.w || r0.width) || 1;
    const capa = html(`<div class="ao-dibujo" role="application" aria-label="Dibuja la trayectoria"><svg width="100%" height="100%"><polyline data-linea points=""/></svg>
      <div class="ao-dibujo-aviso">${ico("route", "ico-s")} Dibuja el camino desde el bloque, sin soltar. Al soltar queda la trayectoria. <button type="button" class="btn chico" data-cancelar>Cancelar</button></div></div>`);
    document.body.append(capa);
    const linea = $("[data-linea]", capa);
    let pts = null;
    const cerrar = () => { capa.remove(); document.removeEventListener("keydown", tecla, true); };
    const tecla = (e) => { if (e.key === "Escape") { e.stopPropagation(); cerrar(); } };
    document.addEventListener("keydown", tecla, true);
    $("[data-cancelar]", capa).onclick = cerrar;
    capa.addEventListener("pointerdown", (e) => {
      if (e.target.closest("button")) return;
      e.preventDefault();
      capa.setPointerCapture(e.pointerId);
      pts = [[e.clientX, e.clientY]];
    });
    capa.addEventListener("pointermove", (e) => {
      if (!pts) return;
      const u = pts[pts.length - 1];
      if (Math.hypot(e.clientX - u[0], e.clientY - u[1]) < 4) return;
      pts.push([e.clientX, e.clientY]);
      linea.setAttribute("points", pts.map((p) => p.join(",")).join(" "));
    });
    const fin = () => {
      if (!pts) return;
      const p = pts;
      pts = null;
      cerrar();
      const largo = p.reduce((s, q, i) => (i ? s + Math.hypot(q[0] - p[i - 1][0], q[1] - p[i - 1][1]) : 0), 0);
      if (p.length < 2 || largo < 12) { aviso("La trayectoria quedó muy corta; vuelve a dibujarla", { tipo: "aviso" }); return; }
      // en px de la diapositiva, desde donde empezó el trazo (PowerPoint: el camino es relativo a la figura)
      const rel = p.map((q) => [(q[0] - p[0][0]) / k, (q[1] - p[0][1]) / k]);
      const cerrada = Math.hypot(p[p.length - 1][0] - p[0][0], p[p.length - 1][1] - p[0][1]) < 14 && largo > 60;
      const y = AO().trayectoriaPropia(id, rel, ctx.T(), cerrada);
      const l = ctx.lista();
      l.push(y);
      ctx.guardar(l);
      aviso("Trayectoria propia agregada (clic " + l.filter((x) => !x.disp && (x.inicio || "clic") === "clic").length + ")", { tipo: "ok", ms: 2400, accion: { texto: "Opciones", fn: () => opciones(ctx, l.length - 1) } });
    };
    capa.addEventListener("pointerup", fin);
    capa.addEventListener("pointercancel", () => { pts = null; cerrar(); });
  }

  // ------------------------------------------------------------ acción al hacer clic (vínculos y botones de acción)
  const ACCIONES = [["", "Nada"], ["sig", "Ir a la siguiente diapositiva"], ["ant", "Ir a la anterior"], ["primera", "Ir a la primera"], ["ultima", "Ir a la última"],
    ["volver", "Volver a la última vista"], ["fin", "Terminar la presentación"], ["diapo", "Ir a la diapositiva…"], ["url", "Abrir una página web o un correo…"]];
  /** Qué pasa al tocar el bloque elegido al presentar (como «Acción» e «Hipervínculo» de PowerPoint). */
  function accion(ctx) {
    const id = ctx.seleccion();
    if (!id) { aviso("Elige primero un bloque de la diapositiva", { tipo: "aviso" }); return; }
    const a = (ctx.doc().bloque(id) || {}).accion || {};
    const pags = ctx.paginas();
    const cuerpo = html(`<div class="col ao-opc"><p class="suave chico" style="font-weight:400">Al presentar, tocar este bloque hace lo que elijas aquí en vez de avanzar. Se guarda en el .pptx como un vínculo de PowerPoint.</p>
      <label class="ao-campo"><span>Al hacer clic</span><select data-f="tipo">${ACCIONES.map(([v, t]) => `<option value="${v}" ${(a.tipo || "") === v ? "selected" : ""}>${t}</option>`).join("")}</select></label>
      <label class="ao-campo" data-si-diapo><span>Diapositiva</span><select data-f="pagina">${pags.map((pg, k) => `<option value="${esc(pg)}" ${a.pagina === pg ? "selected" : ""}>${k + 1}. ${esc(ctx.tituloDiapo(pg))}</option>`).join("")}</select></label>
      <label class="ao-campo" data-si-url><span>Dirección</span><input data-f="url" type="url" inputmode="url" placeholder="https://… o mailto:…" value="${esc(a.url || "")}" style="flex:1 1 220px"></label></div>`);
    const pie = html(`<div style="display:contents"><span class="crece"></span><button class="btn" type="button" data-no>Cancelar</button><button class="btn primario" type="button" data-si>Guardar</button></div>`);
    const h = NOM.hoja({ titulo: "Acción al hacer clic · " + ctx.nombreBloque(id), icono: "mouse-pointer-click", cuerpo, pie });
    const f = (n) => $(`[data-f="${n}"]`, cuerpo);
    const ver = () => { $("[data-si-diapo]", cuerpo).hidden = f("tipo").value !== "diapo"; $("[data-si-url]", cuerpo).hidden = f("tipo").value !== "url"; };
    ver();
    f("tipo").onchange = ver;
    $("[data-no]", pie).onclick = () => h.cerrar();
    $("[data-si]", pie).onclick = () => {
      const t = f("tipo").value;
      let nueva = t ? { tipo: t } : null;
      if (t === "diapo") nueva.pagina = f("pagina").value;
      if (t === "url") {
        let u = f("url").value.trim();
        if (u && !/^[a-z]+:/i.test(u)) u = (/@/.test(u) && !/\//.test(u) ? "mailto:" : "https://") + u;
        if (!/^(https?:\/\/|mailto:)./i.test(u)) { aviso("Escribe una dirección que empiece con https:// o mailto:", { tipo: "aviso" }); return; }
        nueva.url = u;
      }
      if (!ctx.soloLectura()) ctx.ponerAccion(id, nueva);
      h.cerrar();
      aviso(nueva ? "Acción guardada: se usa al presentar" : "Acción quitada", { tipo: "ok", ms: 1800 });
    };
  }

  // ------------------------------------------------------------ menú y ruedita
  /** Lo que se puede hacer con las animaciones del bloque elegido (para el menú y la ruedita). */
  function acciones(ctx, plano) {
    const id = ctx.seleccion(), l = ctx.lista(), ro = ctx.soloLectura();
    const items = [];
    if (id && !ro) {
      items.push({ texto: "Agregar animación…", corto: "Agregar", icono: "plus", fn: () => agregar(ctx) });
      const efs = unidades(l).filter((u) => u.x.id === id).map((u) => ({ texto: `${nombreDe(u.x)} (${TIPO[u.x.tipo] || "Efecto"})…`, corto: nombreDe(u.x), icono: u.x.tipo === "r" ? "route" : "star", fn: () => opciones(ctx, u.desde) }));
      // (en la ruedita caben 8 casillas: si hay varios efectos van en una casilla con los suyos)
      if (plano || efs.length < 2) items.push(...efs);
      else items.push({ texto: "Opciones de sus efectos", corto: "Efectos", icono: "sliders-horizontal", sub: efs.slice(0, 8) });
      items.push({ texto: "Dibujar una trayectoria", corto: "Trayectoria", icono: "route", fn: () => dibujarTrayectoria(ctx) });
      items.push({ texto: "Copiar animación", corto: "Copiar", icono: "paintbrush", off: !l.some((x) => x.id === id), fn: () => copiar(ctx) });
      items.push({ texto: "Pegar animación", corto: "Pegar", icono: "clipboard-paste", off: !(NOM._animCopiada && NOM._animCopiada.length), fn: () => pegar(ctx) });
      items.push({ texto: "Quitar las animaciones de esto", corto: "Quitar", icono: "x", off: !l.some((x) => x.id === id), fn: () => quitar(ctx) });
      items.push({ texto: "Acción al hacer clic (vínculo)…", corto: "Vínculo", icono: "mouse-pointer-click", fn: () => accion(ctx) });
    }
    items.push({ texto: "Panel de animación…", corto: "Panel", icono: "list-ordered", fn: () => panel(ctx) });
    return items;
  }
  function menu(ctx, ancla) {
    const id = ctx.seleccion(), items = acciones(ctx, true);
    const cab = id ? [{ titulo: "Animación · " + ctx.nombreBloque(id) }] : [{ titulo: "Elige un bloque para animarlo" }];
    // las primeras de cada tipo, a un toque (como la galería corta de PowerPoint)
    const rapidas = id && !ctx.soloLectura() ? [["entr:10:0", "Desvanecer"], ["entr:2:4", "Desplazar hacia dentro"], ["entr:53:16", "Zoom"], ["emph:26:0", "Pulso"], ["emph:8:0", "Girar"], ["exit:10:0", "Desvanecer (salida)"]].map(([k, t]) => ({ texto: t, icono: k.startsWith("exit") ? "log-out" : k.startsWith("emph") ? "sparkles" : "log-in", fn: () => {
      const y = AO().nuevoEfecto(id, k, "clic"), l = ctx.lista(); l.push(y); ctx.guardar(l);
      aviso(`${TIPO[y.tipo]}: ${nombreDe(y)} (clic ${l.filter((x) => !x.disp && (x.inicio || "clic") === "clic").length})`, { tipo: "ok", ms: 2200, accion: { texto: "Opciones", fn: () => opciones(ctx, l.length - 1) } });
    } })) : [];
    NOM.menu(ancla, cab.concat(items.slice(0, 1), rapidas.length ? ["-"].concat(rapidas, ["-"]) : [], items.slice(1)));
  }
  function rueda(ctx, x, y) {
    const items = acciones(ctx);
    NOM.rueda.abrir(x == null ? innerWidth / 2 : x, y == null ? innerHeight / 2 : y, items, { titulo: "Animación", modo: "clic" });
  }

  // ------------------------------------------------------------ transiciones (como la pestaña Transiciones de PowerPoint)
  /**
   * La transición de la diapositiva p: galería, opciones del efecto, duración, cómo se avanza y aplicar a todas.
   * ctx: { doc(), paginas(), soloLectura(), T(), estiloDiapo(p), htmlDiapo(p), guardar(p, tr | null, todas) }
   */
  function transiciones(ctx, p) {
    const TO = NOM.transOffice;
    if (!TO || !TO.CATALOGO) return;
    const pags = ctx.paginas(), i = pags.indexOf(p);
    const actual = ctx.doc().valor("_doc", "transOffice:" + p) || TO.deLegado(ctx.doc().valor("_doc", "trans:" + p)) || null;
    const pos = TO.buscar(actual) || { e: 0, v: 0 };
    let e = pos.e, v = pos.v;
    let dur = actual && actual.t !== "corte0" ? +actual.dur || TO.CATALOGO[e][4] : TO.CATALOGO[e][4];
    const cats = [...new Set(TO.CATALOGO.map((x) => x[1]))];
    const cuerpo = html(`<div class="col ao-trans">
      <div class="ao-prev" data-prev aria-label="Vista previa"></div>
      <div class="ao-cats" data-cats style="max-height:34vh"></div>
      <div class="ao-tab">
        <label class="ao-campo" data-opc-fila><span>Opciones del efecto</span><select data-f="var"></select></label>
        <label class="ao-campo"><span>Duración</span><input data-f="dur" type="number" min="0.01" max="60" step="0.25"> s</label>
        <div class="ao-campo"><span>Avanzar a la siguiente</span>
          <label class="ao-check"><input data-f="clic" type="checkbox" ${actual && actual.sinClic ? "" : "checked"}> Al hacer clic</label>
          <label class="ao-check"><input data-f="auto" type="checkbox" ${actual && actual.advTm != null ? "checked" : ""}> Después de</label>
          <input data-f="advTm" type="number" min="0" max="3600" step="0.5" value="${actual && actual.advTm != null ? seg2(actual.advTm) : "5"}" aria-label="Segundos para avanzar solo"> s</div>
      </div></div>`);
    const pie = html(`<div style="display:contents"><button class="btn" type="button" data-ver>${ico("play", "ico-s")} Vista previa</button><button class="btn" type="button" data-todas>${ico("copy", "ico-s")} Aplicar a todas</button><span class="crece"></span><button class="btn" type="button" data-no>Cancelar</button><button class="btn primario" type="button" data-si>Guardar</button></div>`);
    const h = NOM.hoja({ titulo: `Transición de la diapositiva ${i + 1}`, icono: "blend", cuerpo, pie, ancha: true, alCerrar: () => { if (juego) juego.saltar(); } });
    const f = (n) => $(`[data-f="${n}"]`, cuerpo);
    let juego = null;
    function pintar() {
      $("[data-cats]", cuerpo).innerHTML = cats.map((c) => `<section><h3 class="ao-cat">${esc(c)}</h3><div class="ao-efs">${TO.CATALOGO.map((x, k) => [x, k]).filter(([x]) => x[1] === c).map(([x, k]) =>
        `<button type="button" class="ao-ef" data-e="${k}" aria-pressed="${k === e}">${ico(k === 0 ? "x" : "blend", "ico-s")}<span>${esc(x[0])}</span></button>`).join("")}</div></section>`).join("");
      const vs = TO.CATALOGO[e][3];
      $("[data-opc-fila]", cuerpo).hidden = vs.length < 2;
      f("var").innerHTML = vs.map((x, k) => `<option value="${k}" ${k === v ? "selected" : ""}>${esc(x[0])}</option>`).join("");
      const sinDur = e === 0 || TO.deVariante(TO.CATALOGO[e], v).t === "cut";
      f("dur").disabled = sinDur;
      f("dur").value = sinDur ? "" : seg2(dur);
      f("advTm").disabled = !f("auto").checked;
    }
    pintar();
    /** La transición con lo que dice el cuadro (null si no hay nada). */
    function armar() {
      const base = TO.deVariante(TO.CATALOGO[e], v);
      const tr = Object.assign({}, base);
      if (e === 0) tr.t = "corte0";
      if (!f("dur").disabled) tr.dur = Math.max(10, Math.round((parseFloat(f("dur").value) || seg2(base.dur)) * 1000));
      else tr.dur = tr.t === "cut" && tr.a.thruBlk ? base.dur : 0;
      if (f("auto").checked) tr.advTm = Math.max(0, Math.round((parseFloat(f("advTm").value) || 0) * 1000));
      if (!f("clic").checked) tr.sinClic = true;
      // el sonido que traía del archivo se queda
      if (actual && actual.snd) { tr.snd = actual.snd; if (actual.sndFin) tr.sndFin = true; }
      if (tr.t === "corte0" && tr.advTm == null && !tr.sinClic) return null;
      return tr;
    }
    function vista() {
      const cont = $("[data-prev]", cuerpo), T = ctx.T(), w = cont.clientWidth || 300, k = w / T.w;
      if (juego) juego.saltar();
      cont.style.height = Math.round(T.h * k) + "px";
      cont.innerHTML = `<div class="ao-prev-diapo diapo-mini" style="${ctx.estiloDiapo(p)};width:${T.w}px;height:${T.h}px;transform:scale(${k})"></div>`;
      const d = cont.firstChild;
      d.innerHTML = ctx.htmlDiapo(p);
      return { d, T };
    }
    function jugar() {
      const { d, T } = vista(), tr = armar();
      if (!tr) return;
      const ant = pags[i - 1];
      const fondo = ant ? (/background:([^;]+);/.exec(ctx.estiloDiapo(ant)) || [])[1] : "#000";
      juego = TO.jugar(d, { html: ant ? ctx.htmlDiapo(ant) : "", fondo: fondo || "#fff" }, tr, T);
    }
    requestAnimationFrame(vista);
    cuerpo.addEventListener("click", (ev) => {
      const b = ev.target.closest("[data-e]");
      if (!b) return;
      e = +b.dataset.e; v = 0; dur = TO.CATALOGO[e][4] || dur;
      pintar(); jugar();
    });
    cuerpo.addEventListener("change", (ev) => {
      if (ev.target === f("var")) { v = +f("var").value; pintar(); jugar(); }
      else if (ev.target === f("dur")) dur = Math.round((parseFloat(f("dur").value) || 0) * 1000) || dur;
      else pintar();
    });
    $("[data-ver]", pie).onclick = jugar;
    $("[data-no]", pie).onclick = () => h.cerrar();
    const guardar = (todas) => { if (!ctx.soloLectura()) ctx.guardar(p, armar(), todas); h.cerrar(); if (todas) aviso("La misma transición en todas las diapositivas", { tipo: "ok", ms: 1800 }); };
    $("[data-si]", pie).onclick = () => guardar(false);
    $("[data-todas]", pie).onclick = () => guardar(true);
    if (ctx.soloLectura()) { $("[data-si]", pie).disabled = true; $("[data-todas]", pie).disabled = true; }
  }

  NOM.animPanel = { transiciones, accion, galeria, opciones, panel, agregar, copiar, pegar, quitar, dibujarTrayectoria, acciones, menu, rueda, catalogo, categorias, unidades, tiempos };
})();

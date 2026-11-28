/* NOM · Experiencia: encontrar las cosas rápido y saber por dónde empezar.
   - Paleta de comandos (Ctrl+K, o Alt+/ dentro del texto): ir a documentos, espacios y acciones
     escribiendo unas letras. Muestra los atajos de cada acción.
   - Barra de carga arriba: solo aparece si algo tarda más de 160 ms y, si aparece, se queda al
     menos 350 ms (nunca parpadea).
   - Primeros pasos en el inicio de una cuenta nueva y un consejo de la ruedita en cada editor.
   - Hoja de atajos (tecla ?). */
(function () {
  "use strict";
  const NOM = window.NOM;
  const { esc, ico, html, $, $$ } = NOM;
  const api = () => NOM.api;
  // El movimiento lo decide la preferencia de la app ("Reducir movimiento"), igual en todos los equipos.
  const reducido = () => document.documentElement.dataset.movimiento === "reducido";
  const conCuenta = () => !!(NOM.api && NOM.api.token());
  const clave = (k) => (conCuenta() ? api().clave(k) : "nom.local." + k);

  // ================================================================== hitos (primeros pasos)
  const hitos = () => NOM.leer(clave("hitos"), {});
  /** ¿Se usa con mouse? Así los consejos nombran el gesto que la persona sí tiene. */
  const conMouse = () => matchMedia("(hover: hover) and (pointer: fine)").matches;
  /** Marca algo que la persona ya hizo (probar la ruedita, ajustar la app…). */
  const propio = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
  const iconoModo = (m) => (propio(MODO_ICONO, m) ? MODO_ICONO[m] : "file-text");
  NOM.hito = (nombre) => {
    const h = hitos();
    if (propio(h, nombre) && h[nombre]) return;
    h[nombre] = Date.now();
    NOM.guardar(clave("hitos"), h);
    document.dispatchEvent(new CustomEvent("nom:hito", { detail: nombre }));
  };

  // ================================================================== barra de carga
  let enCurso = 0, tMostrar = null, desde = 0, barra = null;
  const barraEl = () => barra || (barra = document.body.appendChild(html(`<div class="carga" role="progressbar" aria-label="Cargando" hidden><i></i></div>`)));
  function empezar() {
    enCurso++;
    if (enCurso === 1 && !desde) {
      clearTimeout(tMostrar);
      tMostrar = setTimeout(() => {
        const b = barraEl();
        b.hidden = false; b.classList.remove("fin"); void b.offsetWidth; b.classList.add("activa");
        desde = Date.now();
      }, 160);
    }
    let hecho = false;
    return () => {
      if (hecho) return;
      hecho = true;
      enCurso = Math.max(0, enCurso - 1);
      if (enCurso) return;
      clearTimeout(tMostrar);
      if (!desde) return;
      setTimeout(() => {
        if (enCurso) return;
        const b = barraEl(), i = b.firstElementChild;
        // Se congela donde iba y de ahí se completa (si no, regresaría a cero antes de llenarse).
        i.style.transform = getComputedStyle(i).transform;
        b.classList.remove("activa");
        void i.offsetWidth;
        i.style.transition = "transform .28s cubic-bezier(.16, 1, .3, 1)";
        i.style.transform = "scaleX(1)";
        b.classList.add("fin");
        desde = 0;
        setTimeout(() => { if (!enCurso && !desde) { b.hidden = true; b.classList.remove("fin"); i.style.transform = ""; i.style.transition = ""; } }, 480);
      }, Math.max(0, 350 - (Date.now() - desde)));
    };
  }
  NOM.cargando = { empezar, en: (promesa) => { const fin = empezar(); return Promise.resolve(promesa).finally(fin); } };

  // ================================================================== paleta de comandos
  const sinAcentos = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  /** Qué tan bien coincide lo escrito (0: no coincide). Inicio de palabra > en medio > todas las
      palabras > letras en orden. */
  function puntaje(texto, q) {
    const t = sinAcentos(texto), p = sinAcentos(q).trim();
    if (!p) return 1;
    const i = t.indexOf(p);
    if (i === 0) return 100;
    if (i > 0) return (/[\s·(/-]/.test(t[i - 1]) ? 85 : 60) - Math.min(i, 30) / 3;
    const ps = p.split(/\s+/);
    if (ps.length > 1 && ps.every((w) => t.includes(w))) return 45;
    let j = 0;
    for (const c of t) if (c === p[j]) j++;
    return j === p.length ? 20 - Math.min(t.length, 60) / 6 : 0;
  }
  const MODO_ICONO = { TEXTO: "file-text", HOJA: "sheet", PRESENTACION: "presentation", DIAGRAMA: "workflow" };
  let docsCache = { t: 0, lista: [] };
  /** Documentos de todos los espacios (para buscarlos por nombre); se refresca cada minuto. */
  async function documentosDeEspacios() {
    if (!conCuenta()) return [];
    if (Date.now() - docsCache.t < 60000) return docsCache.lista;
    const espacios = (NOM.app && NOM.app.espacios) || [];
    const listas = await Promise.all(espacios.slice(0, 12).map((e) => api().get(`/api/espacios/${encodeURIComponent(e.id)}/documentos`, undefined, { ms: 5000 })
      .then((l) => (l || []).map((d) => ({ ...d, espacio: e.nombre }))).catch(() => [])));
    docsCache = { t: Date.now(), lista: listas.flat() };
    return docsCache.lista;
  }
  function cambiarPreferencia(campos, texto) {
    const p = Object.assign({}, NOM.leer("nom.pref", {}), campos);
    NOM.aplicarPreferencias(p);
    if (conCuenta()) api().put("/api/cuentas/preferencias", p).catch(() => {});
    NOM.hito("preferencias");
    NOM.aviso(texto, { tipo: "ok", ms: 1800 });
  }
  function espacioActual() {
    const m = location.hash.match(/^#\/espacio\/([^/?]+)/);
    const lista = (NOM.app && NOM.app.espacios) || [];
    if (m) return lista.find((e) => e.id === m[1]) || null;
    const ed = editorVisible();
    return ed && ed.espacio && ed.espacio.id ? lista.find((e) => e.id === ed.espacio.id) || ed.espacio : null;
  }
  const editorVisible = () => { const ed = NOM.editorActual; return ed && ed.vista && ed.vista.isConnected ? ed : null; };

  const prefExtra = () => NOM.prefExtra || null;
  const etiquetaAtajo = (accion, defecto) => {
    const px = prefExtra();
    if (!px) return defecto;
    return px.etiquetaCorta(px.teclaDe(accion));
  };
  const kbdAtajo = (accion, defecto) => {
    const px = prefExtra();
    const partes = px ? px.partesTecla(px.teclaDe(accion)) : (defecto ? defecto.split(" ") : []);
    return partes.length ? partes.map((x) => `<kbd>${esc(x)}</kbd>`).join("") : `<span class="suave chico">Sin atajo</span>`;
  };
  const mapaAtajoEditor = { "Guardar versión": "guardar", "Modo enfoque": "enfoque", "Comprobar accesibilidad": "accesibilidad" };

  /** Todo lo que se puede buscar ahora, según dónde está la persona (docs: documentos de los espacios). */
  function elementos(docs) {
    const l = [];
    const ed = editorVisible();
    const pref = NOM.leer("nom.pref", {});
    const espacios = (NOM.app && NOM.app.espacios) || [];
    const u = conCuenta() ? api().usuario() || {} : null;
    if (ed && ed.acciones) ed.acciones().forEach((a) => {
      const copia = Object.assign({ grupo: "En este documento" }, a);
      if (propio(mapaAtajoEditor, a.texto)) copia.atajo = etiquetaAtajo(mapaAtajoEditor[a.texto], a.atajo || "");
      l.push(copia);
    });

    const nuevoDoc = (esp) => () => NOM.vistaEspacios.nuevoDocumento(esp);
    if (conCuenta()) {
      const esp = espacioActual();
      if (esp) l.push({ grupo: "Crear", texto: `Nuevo documento en ${esp.nombre}`, icono: "file-plus", atajo: "", fn: nuevoDoc(esp), claves: "crear documento nuevo" });
      else if (espacios.length === 1) l.push({ grupo: "Crear", texto: `Nuevo documento en ${espacios[0].nombre}`, icono: "file-plus", fn: nuevoDoc(espacios[0]), claves: "crear documento nuevo" });
      else if (espacios.length) l.push({ grupo: "Crear", texto: "Nuevo documento…", icono: "file-plus", claves: "crear documento nuevo", sub: () => espacios.map((e) => ({ grupo: "¿En qué espacio?", texto: e.nombre, icono: "layout-grid", fn: nuevoDoc(e) })) });
      l.push({ grupo: "Crear", texto: "Nuevo espacio", icono: "plus", fn: () => NOM.vistaEspacios.nuevoEspacio(), claves: "crear espacio equipo proyecto" });
      l.push({ grupo: "Crear", texto: "Unirse a un espacio", icono: "qr-code", fn: () => NOM.vistaEspacios.unirse(), claves: "codigo qr invitacion entrar wifi buscar" });
      l.push({ grupo: "Crear", texto: "Unirse con código QR", icono: "qr-code", fn: () => NOM.union.escanear(), claves: "qr camara codigo salon escanear" });
      if (esp && esp.permiso === "PROPIETARIO") l.push({ grupo: "Crear", texto: `Invitar a ${esp.nombre} con código QR`, icono: "qr-code", fn: () => NOM.union.invitar(esp), claves: "qr invitar salon alumnos alias apodo sin cuenta" });
      if (esp && esp.id && !ed) l.push({ grupo: "Crear", texto: `Exportar todo ${esp.nombre} (.zip)`, icono: "archive", fn: async () => NOM.vistaEspacios.exportarTodo(await api().get(`/api/espacios/${encodeURIComponent(esp.id)}`)), claves: "descargar office zip respaldo" });
    } else {
      l.push({ grupo: "Crear", texto: "Nuevo documento (en este equipo)", icono: "file-plus", fn: nuevoDoc(null), claves: "crear documento nuevo" });
      l.push({ grupo: "Crear", texto: "Crear una cuenta", icono: "user-plus", fn: () => NOM.app.irACuentas("registro"), claves: "registrarse cuenta" });
    }
    // Fase 13, parte C: instalar la página como programa (Chrome y Edge).
    if (NOM.instalar && NOM.instalar.disponible()) l.push({ grupo: "Este equipo", texto: "Instalar en esta computadora", icono: "monitor-down", fn: () => NOM.instalar.pedir(), claves: "instalar programa escritorio pwa sin internet aplicacion" });

    // Recientes primero; luego todos los documentos.
    const vistos = new Set();
    const recientes = conCuenta() ? NOM.leer(api().clave("recientes"), []) : [];
    recientes.forEach((r) => { vistos.add(r.id); l.push({ grupo: "Recientes", texto: r.nombre || "Sin título", sub2: r.espacio, icono: iconoModo(r.modo), fn: () => (location.hash = "#/doc/" + r.id), modo: r.modo }); });
    (docs || []).forEach((d) => { if (!vistos.has(d.id)) l.push({ grupo: "Documentos", texto: d.nombre || "Sin título", sub2: d.espacio, icono: iconoModo(d.modo), fn: () => (location.hash = "#/doc/" + d.id), modo: d.modo, oculto: true }); });
    (NOM.local && NOM.local.lista ? NOM.local.lista() : []).forEach((d) => l.push({ grupo: "En este equipo", texto: d.nombre || "Sin título", sub2: "Solo en este equipo", icono: iconoModo(d.modo), fn: () => (location.hash = "#/local/" + d.id), modo: d.modo, oculto: true }));
    espacios.forEach((e) => l.push({ grupo: "Espacios", texto: e.nombre, sub2: `${e.documentos} documento${e.documentos === 1 ? "" : "s"} · ${e.miembros} persona${e.miembros === 1 ? "" : "s"}`, icono: "layout-grid", color: e.color, fn: () => (location.hash = "#/espacio/" + e.id), oculto: espacios.length > 5 }));

    l.push({ grupo: "Ir a", texto: "Inicio", icono: "house", fn: () => (location.hash = "#/"), claves: "principal home" });
    l.push({ grupo: "Ir a", texto: NOM.etiquetaEquipo(), icono: "hard-drive", fn: () => (location.hash = "#/dispositivo"), claves: "local almacenamiento dispositivo archivos sin internet telefono" });
    if (NOM.anfitrion && NOM.anfitrion.disponible() && conCuenta()) l.push({ grupo: "Ir a", texto: "Trabajar sin Internet con este teléfono", icono: "smartphone", fn: () => (location.hash = "#/dispositivo"), claves: "anfitrion compartir red wifi punto de acceso" });

    const tema = pref.tema || "auto";
    l.push({ grupo: "Apariencia", texto: "Tema oscuro", icono: "moon", marcado: tema === "oscuro", fn: () => cambiarPreferencia({ tema: "oscuro" }, "Tema oscuro"), claves: "noche negro dark" });
    l.push({ grupo: "Apariencia", texto: "Tema claro", icono: "sun", marcado: tema === "claro", fn: () => cambiarPreferencia({ tema: "claro" }, "Tema claro"), claves: "dia blanco light" });
    l.push({ grupo: "Apariencia", texto: "Tema automático (como el sistema)", icono: "monitor", marcado: tema === "auto", fn: () => cambiarPreferencia({ tema: "auto" }, "Tema automático"), claves: "sistema" });
    l.push({ grupo: "Apariencia", texto: pref.movimientoReducido ? "Activar animaciones" : "Reducir movimiento", icono: "move", fn: () => cambiarPreferencia({ movimientoReducido: !pref.movimientoReducido }, pref.movimientoReducido ? "Animaciones activadas" : "Movimiento reducido"), claves: "animaciones accesibilidad" });
    l.push({ grupo: "Ayuda", texto: "Atajos de teclado", icono: "keyboard", atajo: etiquetaAtajo("atajos", "?"), fn: atajos, claves: "teclas shortcuts ayuda" });
    l.push({ grupo: "Ayuda", texto: "Verificar un recibo de escritura", icono: "circle-check", fn: () => NOM.recibo.verificar(), claves: "recibo autoria hash comprobar" });
    l.push({ grupo: "Ayuda", texto: "Ajustes y preferencias", icono: "settings", atajo: etiquetaAtajo("ajustes", ""), fn: () => NOM.vistaEspacios.preferencias(), claves: "configuracion" });
    if (conCuenta()) l.push({ grupo: "Ayuda", texto: "Cerrar sesión", icono: "log-out", fn: () => NOM.app.cerrarSesion && NOM.app.cerrarSesion(), claves: "salir logout" });
    return l;
  }

  let paleta = null;
  async function abrirPaleta(inicial) {
    if (paleta) { $("input", paleta.el).focus(); return; }
    NOM.cerrarMenu(); NOM.rueda && NOM.rueda.cerrar();
    const anterior = document.activeElement;
    const teclaPaleta = prefExtra() ? prefExtra().partesTecla(prefExtra().teclaDe("paleta")) : ["Ctrl", "K"];
    const el = html(`<div class="paleta-velo" role="presentation">
      <div class="paleta" role="dialog" aria-modal="true" aria-label="Buscar o ir a">
        <div class="paleta-buscar">${ico("search")}<input type="text" autocomplete="off" spellcheck="false" placeholder="Busca documentos, espacios o acciones…" aria-label="Buscar" role="combobox" aria-expanded="true" aria-controls="paleta-lista"><button class="paleta-esc" type="button" aria-label="Cerrar"><kbd>Esc</kbd></button></div>
        <div class="paleta-lista" id="paleta-lista" role="listbox"><div class="paleta-cargando">${'<div class="esqueleto" style="height:40px"></div>'.repeat(4)}</div></div>
        <div class="paleta-pie"><span><kbd>↑</kbd><kbd>↓</kbd> moverse</span><span><kbd>Enter</kbd> abrir</span>${teclaPaleta.length ? `<span class="solo-escritorio">${teclaPaleta.map((x) => `<kbd>${esc(x)}</kbd>`).join("")} en cualquier momento</span>` : ""}</div>
      </div></div>`);
    document.body.append(el);
    const input = $("input", el), lista = $(".paleta-lista", el);
    let todos = [], vistos = [], activo = 0, pila = [];
    const cerrar = () => {
      if (!paleta) return;
      paleta = null;
      document.removeEventListener("keydown", tecla, true);
      el.classList.add("saliendo");
      setTimeout(() => el.remove(), 160);
      if (anterior && anterior.focus && anterior.isConnected) anterior.focus({ preventScroll: true });
    };
    paleta = { el, cerrar };
    function pintar() {
      const q = input.value;
      const base = pila.length ? pila[pila.length - 1] : todos;
      vistos = base.map((it) => ({ it, p: Math.max(puntaje(it.texto, q), it.claves ? puntaje(it.claves, q) * 0.8 : 0, it.sub2 ? puntaje(it.sub2, q) * 0.5 : 0) }))
        .filter((x) => x.p > 0 && (q.trim() || !x.it.oculto || pila.length))
        .sort((a, b) => (q.trim() ? b.p - a.p : 0)).map((x) => x.it).slice(0, 40);
      activo = Math.min(activo, Math.max(0, vistos.length - 1));
      if (!vistos.length) {
        lista.innerHTML = `<div class="paleta-vacio">${ico("search")}<b>Nada con “${esc(q)}”</b><span>Prueba con otra palabra, o crea algo nuevo.</span></div>`;
        input.removeAttribute("aria-activedescendant");
        return;
      }
      let grupo = null, h = "";
      vistos.forEach((it, i) => {
        if (!q.trim() && it.grupo !== grupo) { grupo = it.grupo; h += `<div class="paleta-grupo">${esc(grupo)}</div>`; }
        h += `<button type="button" id="paleta-op-${i}" class="paleta-item ${i === activo ? "activo" : ""}" role="option" aria-selected="${i === activo}" data-i="${i}" style="--i:${Math.min(i, 12)}">
          <span class="paleta-ico ${it.modo ? "modo-" + esc(it.modo) : ""}" ${it.color ? `style="background:${NOM.colorHex(it.color)};color:#fff"` : ""}>${ico(it.icono || "chevron-right")}</span>
          <span class="crece recorta"><b>${esc(it.texto)}</b>${it.sub2 ? `<small>${esc(it.sub2)}</small>` : ""}</span>
          ${it.marcado ? ico("check", "ico-s") : ""}${it.atajo ? `<kbd>${esc(it.atajo)}</kbd>` : ""}${it.sub ? ico("chevron-right", "ico-s") : ""}</button>`;
      });
      lista.innerHTML = h;
      input.setAttribute("aria-activedescendant", "paleta-op-" + activo);
      const a = $(".paleta-item.activo", lista);
      if (a) a.scrollIntoView({ block: "nearest" });
    }
    function marcar(i) {
      activo = (i + vistos.length) % vistos.length;
      $$(".paleta-item", lista).forEach((b) => { const on = +b.dataset.i === activo; b.classList.toggle("activo", on); b.setAttribute("aria-selected", on); });
      input.setAttribute("aria-activedescendant", "paleta-op-" + activo);
      const a = $(".paleta-item.activo", lista);
      if (a) a.scrollIntoView({ block: "nearest" });
    }
    function elegir(i) {
      const it = vistos[i];
      if (!it) return;
      if (it.sub) { pila.push(it.sub()); input.value = ""; activo = 0; input.placeholder = it.texto.replace(/…$/, "") + ": elige…"; pintar(); return; }
      cerrar();
      setTimeout(() => it.fn && it.fn(), 0);
    }
    function tecla(e) {
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); if (pila.length) { pila.pop(); input.placeholder = "Busca documentos, espacios o acciones…"; pintar(); } else cerrar(); return; }
      if (e.key === "ArrowDown" || (e.key === "Tab" && !e.shiftKey)) { e.preventDefault(); marcar(activo + 1); }
      else if (e.key === "ArrowUp" || (e.key === "Tab" && e.shiftKey)) { e.preventDefault(); marcar(activo - 1); }
      else if (e.key === "Enter") { e.preventDefault(); elegir(activo); }
      else if (e.key === "Backspace" && !input.value && pila.length) { pila.pop(); pintar(); }
    }
    document.addEventListener("keydown", tecla, true);
    input.addEventListener("input", () => { activo = 0; pintar(); });
    lista.addEventListener("pointermove", (e) => { const b = e.target.closest(".paleta-item"); if (b && +b.dataset.i !== activo && e.pointerType === "mouse") marcar(+b.dataset.i); });
    lista.addEventListener("click", (e) => { const b = e.target.closest(".paleta-item"); if (b) elegir(+b.dataset.i); });
    el.addEventListener("pointerdown", (e) => { if (e.target === el) cerrar(); });
    $(".paleta-esc", el).onclick = cerrar;
    input.value = inicial || "";
    requestAnimationFrame(() => input.focus({ preventScroll: true }));
    // Se muestra al instante con lo que ya se sabe; los documentos de los espacios llegan después.
    todos = elementos(docsCache.lista);
    pintar();
    documentosDeEspacios().then((docs) => { if (paleta && paleta.el === el) { todos = elementos(docs); pintar(); } }).catch(() => {});
  }
  NOM.paleta = abrirPaleta;

  // ================================================================== hoja de atajos
  function atajos() {
    const fila = ([t, k, acc]) => `<div class="atajo"><span>${esc(t)}</span><span>${acc ? kbdAtajo(acc, k) : k.split(" ").map((x) => `<kbd>${esc(x)}</kbd>`).join("")}</span></div>`;
    const g = (titulo, filas) => `<div class="atajos-grupo"><h3>${esc(titulo)}</h3>${filas.map(fila).join("")}</div>`;
    const pie = NOM.vistaEspacios && NOM.vistaEspacios.preferencias
      ? html(`<div style="display:contents"><button type="button" class="btn" data-cambiar-atajos>${ico("settings", "ico-s")} Cambiar atajos en Ajustes</button><span class="crece"></span><button type="button" class="btn primario" data-listo>Cerrar</button></div>`)
      : null;
    const h = NOM.hoja({ titulo: "Atajos de teclado", icono: "keyboard", ancha: true, pie, cuerpo: `<div class="atajos">
      ${g("En toda la app", [["Buscar o ir a…", "Ctrl K", "paleta"], ["Buscar dentro del texto", "Alt /", "paletaTexto"], ["Estos atajos", "?", "atajos"], ["Abrir Ajustes", "Ctrl ,", "ajustes"], ["Ir a Inicio", "Alt H", "inicio"], ["Crear un documento nuevo", "Alt N", "nuevoDoc"], ["Cambiar entre claro y oscuro", "Alt T", "tema"], ["Cerrar ventanas y menús", "Esc"]])}
      ${g("Documentos", [["Guardar versión", "Ctrl S", "guardar"], ["Deshacer", "Ctrl Z"], ["Rehacer", "Ctrl Y"], ["Modo enfoque", "Ctrl .", "enfoque"], ["Comprobar accesibilidad", "Alt Mayús A", "accesibilidad"], ["Buscar y reemplazar", "Ctrl F"]])}
      ${g("Formato", [["Negrita", "Ctrl B"], ["Cursiva", "Ctrl I"], ["Subrayado", "Ctrl U"], ["Enlace (texto)", "Ctrl K"], ["Título 1, 2, 3", "Ctrl Alt 1"]])}
      ${g("Menú rápido (ruedita o lista)", [["Mantener presionado (dedo o clic)", "½ s"], ["Al instante", "Clic derecho"], ["Abrirlo donde está el cursor", "Menú"], ["Elegir directo", "1-8"], ["Recorrer", "← →"], ["Regresar de un submenú", "Retroceso"]])}
      ${g("Presentaciones y diagramas", [["Duplicar", "Ctrl D"], ["Agrupar", "Ctrl G"], ["Elegir todo", "Ctrl A"], ["Eliminar", "Supr"], ["Mover fino", "Flechas"], ["Elegir varias", "Mayús clic"]])}
      ${g("Hoja de cálculo", [["Editar celda", "F2"], ["Elegir un rango", "Mayús Flechas"], ["Siguiente celda", "Tab"], ["Bajar", "Enter"], ["Borrar", "Supr"]])}
    </div>` });
    if (pie) {
      $("[data-cambiar-atajos]", pie).onclick = () => { h.cerrar(true); NOM.vistaEspacios.preferencias(); };
      $("[data-listo]", pie).onclick = () => h.cerrar();
    }
  }
  NOM.atajos = atajos;

  function ejecutarAccionAtajo(accion, e) {
    const ed = editorVisible();
    if (accion === "paleta" || accion === "paletaTexto") { e.preventDefault(); abrirPaleta(); return true; }
    if (accion === "atajos") { e.preventDefault(); atajos(); return true; }
    if (accion === "ajustes") { if (NOM.vistaEspacios && NOM.vistaEspacios.preferencias) { e.preventDefault(); NOM.vistaEspacios.preferencias(); return true; } return false; }
    if (accion === "inicio") { e.preventDefault(); location.hash = "#/"; return true; }
    if (accion === "nuevoDoc") {
      if (!NOM.vistaEspacios || !NOM.vistaEspacios.nuevoDocumento) return false;
      const esp = conCuenta() ? (espacioActual() || ((NOM.app && NOM.app.espacios) || [])[0] || null) : null;
      if (conCuenta() && !esp) return false;
      e.preventDefault(); NOM.vistaEspacios.nuevoDocumento(esp); return true;
    }
    if (accion === "tema") {
      const actual = document.documentElement.dataset.tema === "oscuro" ? "claro" : "oscuro";
      e.preventDefault();
      cambiarPreferencia({ tema: actual }, actual === "oscuro" ? "Tema oscuro" : "Tema claro");
      return true;
    }
    if (accion === "accesibilidad") {
      if (ed && NOM.accesibilidad && NOM.accesibilidad.abrir) { e.preventDefault(); NOM.accesibilidad.abrir(ed); return true; }
      return false;
    }
    if (accion === "guardar" && ed) {
      const px = prefExtra();
      if (px && px.teclaDe("guardar") === "Ctrl+S") return false; // editor.js conserva su manejo de fábrica
      const act = ed.acciones && ed.acciones().find((x) => x.texto === "Guardar versión");
      if (act && act.fn) { e.preventDefault(); if (e.stopImmediatePropagation) e.stopImmediatePropagation(); act.fn(); return true; }
      return false;
    }
    if (accion === "enfoque" && ed) {
      const px = prefExtra();
      if (px && px.teclaDe("enfoque") === "Ctrl+.") return false; // editor.js conserva su manejo de fábrica
      e.preventDefault();
      if (e.stopImmediatePropagation) e.stopImmediatePropagation();
      if (ed.vista && ed.vista.classList.contains("enfoque")) {
        ed.vista.classList.remove("enfoque");
        if (document.fullscreenElement && document.exitFullscreen) document.exitFullscreen().catch(() => {});
      } else {
        const act = ed.acciones && ed.acciones().find((x) => x.texto === "Modo enfoque");
        if (act && act.fn) act.fn();
      }
      return true;
    }
    return false;
  }

  // ================================================================== teclado global
  document.addEventListener("keydown", (e) => {
    if (e.defaultPrevented) return;
    const t = e.target;
    const enCampo = !!(t && (t.isContentEditable || /INPUT|TEXTAREA|SELECT/.test(t.tagName)));
    const enEditor = !!editorVisible();
    const px = prefExtra();
    if (px) {
      const accion = px.buscarAccionPorEvento(e, { enCampo, enEditor, hayHoja: !!(NOM.hayHoja && NOM.hayHoja()) });
      if (accion) {
        if (ejecutarAccionAtajo(accion, e)) return;
      }
      // Si la persona cambió o quitó Ctrl+S o Ctrl+., la tecla anterior deja de disparar la acción fija de editor.js.
      if (enEditor && px.esDefectoCambiado(e, ["guardar", "enfoque"])) {
        e.preventDefault();
        if (e.stopImmediatePropagation) e.stopImmediatePropagation();
        return;
      }
      return;
    }
    const mod = e.ctrlKey || e.metaKey;
    // En el texto, Ctrl+K es "enlace" (como en Word): ahí la paleta se abre con Alt+/.
    if ((mod && !e.shiftKey && !e.altKey && e.key.toLowerCase() === "k" && !enCampo) || (e.altKey && !mod && e.key === "/")) {
      e.preventDefault(); abrirPaleta(); return;
    }
    if (e.key === "?" && !mod && !enCampo && !enEditor && !NOM.hayHoja()) { e.preventDefault(); atajos(); }
  });
  document.addEventListener("click", (e) => {
    const b = e.target.closest && e.target.closest("[data-a=paleta]");
    if (b) { e.preventDefault(); abrirPaleta(); }
  });

  // ================================================================== primeros pasos
  /** Tarjeta de primeros pasos en el inicio de una cuenta; desaparece al completarla u ocultarla. */
  NOM.primerosPasos = (cont, espacios) => {
    if (!cont || !conCuenta()) return;
    const ya = NOM.leer(clave("pasos.oculto"), false);
    if (ya) return;
    const h = hitos();
    const recientes = NOM.leer(api().clave("recientes"), []);
    const propio = espacios.find((e) => e.permiso === "PROPIETARIO") || espacios[0];
    const pasos = [
      { id: "espacio", t: "Crear un espacio", d: "Es una carpeta que compartes con tu equipo.", ok: espacios.length > 0, fn: () => NOM.vistaEspacios.nuevoEspacio() },
      { id: "documento", t: "Crear un documento", d: "De texto, hoja de cálculo, presentación o diagrama.", ok: espacios.some((e) => e.documentos > 0) || recientes.length > 0, fn: () => (propio ? NOM.vistaEspacios.nuevoDocumento(propio) : NOM.vistaEspacios.nuevoEspacio()) },
      { id: "rueda", t: "Abrir la ruedita", d: (conMouse() ? "Clic derecho" : "Mantén presionado") + " sobre lo que quieras cambiar.", ok: !!h.rueda, fn: () => (recientes[0] ? (location.hash = "#/doc/" + recientes[0].id) : propio ? NOM.vistaEspacios.nuevoDocumento(propio) : NOM.vistaEspacios.nuevoEspacio()) },
      { id: "equipo", t: "Invitar a alguien", d: "Con un código o un QR.", ok: espacios.some((e) => e.miembros > 1 || e.conexion), fn: () => (propio ? NOM.vistaEspacios.accesos(propio) : NOM.vistaEspacios.unirse()) },
      { id: "ajustes", t: "Revisar las preferencias", d: "Tema, guardado automático y movimiento.", ok: !!h.preferencias, fn: () => NOM.vistaEspacios.preferencias() },
    ];
    const hechos = pasos.filter((p) => p.ok).length;
    // Cuando ya se hizo todo, la lista deja de salir.
    if (hechos === pasos.length) return;
    const el = html(`<section class="pasos" aria-label="Para empezar">
      <div class="pasos-cab">
        <h2 class="crece">Para empezar <span class="suave">${hechos} de ${pasos.length}</span></h2>
        <button class="btn plano chico" type="button" data-ocultar>Ocultar</button>
      </div>
      <ol class="pasos-lista">${pasos.map((p, i) => `<li class="${p.ok ? "hecho" : ""}">${p.ok
        ? `<span>${esc(p.t)}</span>`
        : `<button type="button" class="enlace" data-p="${i}">${esc(p.t)}</button>`} <span class="suave">${esc(p.d)}</span></li>`).join("")}</ol>
    </section>`);
    cont.after(el);
    $("[data-ocultar]", el).onclick = () => { NOM.guardar(clave("pasos.oculto"), true); el.hidden = true; };
    $$("[data-p]", el).forEach((b) => (b.onclick = () => pasos[+b.dataset.p].fn()));
  };

  // ================================================================== consejo de la ruedita en el editor
  // Con mouse se dice "clic derecho"; con el dedo, "mantén presionado" (el gesto que esa persona sí tiene).
  const nombreMenuRapido = () => (prefExtra() && prefExtra().esLista() ? "el menú de acciones" : "la ruedita");
  const CONSEJOS = {
    TEXTO: () => (conMouse() ? "Haz clic derecho sobre el texto" : "Mantén presionado sobre el texto") + `: sale ${nombreMenuRapido()} con negrita, color, estilos y más.`,
    HOJA: () => (conMouse() ? "Haz clic derecho en una celda" : "Mantén presionada una celda") + `: sale ${nombreMenuRapido()} con formato, fórmulas y más.`,
    PRESENTACION: () => (conMouse() ? `Haz clic derecho en una figura: sale ${nombreMenuRapido()}. Para elegir varias, usa Mayús + clic.` : `Toca un cuadro ya elegido para escribir en él. Mantén presionada una figura: sale ${nombreMenuRapido()}.`),
    DIAGRAMA: () => (conMouse() ? `Haz clic derecho en una figura: sale ${nombreMenuRapido()}. Para elegir varias, usa Mayús + clic.` : `Mantén presionada una figura: sale ${nombreMenuRapido()}. Para elegir varias, sostén una y toca las demás.`),
  };
  NOM.alAbrirEditor = (ed, tipo) => {
    if (!ed || ed.soloLectura || !propio(CONSEJOS, tipo)) return;
    const k = "consejo." + tipo;
    if (hitos()[k] || hitos().rueda) return;
    setTimeout(() => {
      if (!ed.vista || !ed.vista.isConnected) return;
      const el = html(`<div class="consejo" role="status">${ico("circle-dot")}<span>${esc(CONSEJOS[tipo]())}</span><button class="btn chico" type="button">Entendido</button></div>`);
      ed.vista.append(el);
      const quitar = () => { if (!el.isConnected) return; NOM.hito(k); el.classList.add("saliendo"); setTimeout(() => el.remove(), 260); document.removeEventListener("nom:hito", alHito); };
      const alHito = (e) => { if (e.detail === "rueda") quitar(); };
      document.addEventListener("nom:hito", alHito);
      $("button", el).onclick = quitar;
      setTimeout(quitar, 16000);
    }, 1400);
  };
})();

/* NOM · Ruedita de acciones (menú radial o "pie menu").
   Como la rueda de "ping" de los juegos: un anillo de acciones alrededor del dedo o del cursor.
   Se elige por DIRECCIÓN, no por posición exacta; con práctica se hace de memoria sin mirar
   (marking menu, Kurtenbach y Buxton). Igual en computadora, celular y la app.

   Cómo se abre (siempre centrada donde está el dedo o el cursor, y solo si se mantiene):
   - Mantener presionado (dedo o clic izquierdo) ~0.55 s sin moverse. Con el clic, también al
     hacer una pausa mientras se arrastra para seleccionar texto o celdas. Con el dedo sobre texto
     manda la pulsación larga del sistema: la rueda sale en ese momento, junto con la selección.
   - Clic derecho (o la tecla Menú): al instante.
   Cómo se elige: deslizar hacia la acción y soltar; o soltar sin moverse y luego tocar la acción
   (la rueda se queda abierta hasta elegir). Al terminar la acción, la rueda se va.
   Teclado: flechas, Enter, 1-8 directo, Retroceso regresa de un submenú, Esc cierra.
   Mayús + clic derecho deja el menú normal del navegador (corrector ortográfico, etc.).

   Las 8 posiciones tienen el mismo orden en todos los editores (en el sentido del reloj, desde
   arriba), para que la mano las aprenda:
     0 Negrita · 1 Cursiva · 2 Subrayado · 3 Color · 4 Eliminar o cortar (abajo) ·
     5 Copiar o duplicar · 6 Estilo, orden o formato · 7 Más
   Lo que no aplica a la selección se muestra apagado en su lugar; las demás no se recorren. */
(function () {
  "use strict";
  const NOM = window.NOM;
  const { esc, ico, html } = NOM;

  const R = 112, RI = 38, MUERTA = 26;   // radio externo, interno y zona muerta del centro
  const MANTENER_MS = 550;                // cuánto hay que mantener presionado para que salga
  let actual = null, saliendo = null;

  function vibrar(tipo) {
    if (NOM.prefExtra && NOM.prefExtra.vibrar) return NOM.prefExtra.vibrar(tipo);
    try {
      const ms = tipo === "mover" ? 5 : tipo === "elegir" ? 16 : 12;
      if (typeof navigator !== "undefined" && navigator && typeof navigator.vibrate === "function") { navigator.vibrate(ms); return true; }
    } catch (e) { /* sin vibración */ }
    return false;
  }

  function usarLista() {
    return !!(NOM.prefExtra && NOM.prefExtra.esLista && NOM.prefExtra.esLista());
  }

  function siguienteHabil(items, desde, paso) {
    const n = items.length;
    for (let k = 0; k < n; k++) {
      const j = (desde + paso * k + n * 2) % n;
      if (items[j] && !items[j].off) return j;
    }
    return -1;
  }

  function restaurarFoco(anterior) {
    if (anterior && anterior.focus && anterior.isConnected) {
      try { anterior.focus({ preventScroll: true }); } catch (e) { /* elemento no enfocable */ }
    }
  }

  /** Versión en lista vertical (menú contextual clásico) con las mismas acciones y en el mismo orden. */
  function abrirLista(x, y, items, opc, interno) {
    const n = items.length;
    const v = visible();
    const anterior = !interno ? document.activeElement : (opc._anterior || null);
    const velo = html(`<div class="rueda-velo rueda-velo-lista" role="presentation">
      <div class="menu rueda-lista${opc.padre ? " sub" : ""}" role="menu" aria-label="${esc(opc.titulo || "Acciones rápidas")}">
        ${opc.padre
          ? `<button type="button" class="rueda-lista-atras" data-atras aria-label="Volver">${ico("chevron-left", "ico-s")}<span>${esc(opc.titulo || "Volver")}</span></button>`
          : (opc.titulo ? `<div class="menu-titulo">${esc(opc.titulo)}</div>` : "")}
        ${items.map((it, i) => `<button type="button" class="rueda-lista-item ${it.off ? "off" : ""} ${it.marcado ? "marcado" : ""}" role="menuitem" data-i="${i}" tabindex="-1"${it.off ? ' aria-disabled="true"' : ""}${it.sub ? ' aria-haspopup="menu"' : ""}>
          ${it.color ? `<span class="rueda-lista-muestra" style="background:${NOM.colorCss(it.color, "#5A5D6B")}"></span>` : ico(it.icono || "circle")}
          <span class="crece">${esc(it.texto)}</span>
          ${it.marcado ? ico("check", "ico-s") : ""}${it.sub ? ico("chevron-right", "ico-s") : ""}</button>`).join("")}
      </div></div>`);
    document.body.append(velo);
    const menu = velo.firstElementChild;
    const botones = Array.from(menu.querySelectorAll(".rueda-lista-item"));
    const w = menu.offsetWidth || 232;
    const h = menu.offsetHeight || (n * 40 + 44);
    let px = Math.min(v.x + v.w - w - 8, Math.max(v.x + 8, x));
    let py = y + 4;
    if (py + h > v.y + v.h - 8) py = y - h - 4;
    py = Math.min(v.y + v.h - h - 8, Math.max(v.y + 8, py));
    menu.style.left = px + "px";
    menu.style.top = py + "px";

    let elegido = -1, modo = opc.modo || "clic", soltado = null;
    function marcar(i, enfocar) {
      elegido = i;
      botones.forEach((b) => b.classList.toggle("activo", +b.dataset.i === i));
      if (enfocar && i >= 0 && botones[i]) botones[i].focus({ preventScroll: true });
    }
    function elegir(i) {
      const it = items[i];
      if (!it || it.off) return;
      if (opc.tactil !== false) vibrar("elegir");
      if (it.sub) { abrir(x, y, it.sub, Object.assign({}, opc, { modo: "clic", titulo: it.texto, padre: { items, opc }, _anterior: anterior }), true); return; }
      cerrar();
      it.fn && it.fn();
      if (NOM.hito) NOM.hito("rueda");
    }
    function atras() { if (opc.padre) abrir(x, y, opc.padre.items, Object.assign({}, opc.padre.opc, { _anterior: anterior }), true); else cerrar(); }
    function mover(e) {
      if (e.pointerType !== "mouse") return;
      const b = e.target && e.target.closest && e.target.closest(".rueda-lista-item");
      if (b && menu.contains(b) && !items[+b.dataset.i].off) marcar(+b.dataset.i, false);
    }
    function soltar(e) {
      if (modo !== "mantener") return;
      // Soltar el gesto que abrió la lista no la cierra ni dispara una opción por accidente.
      modo = "clic";
      soltado = { t: Date.now(), x: e.clientX, y: e.clientY };
    }
    function clic(e) {
      if (modo !== "clic") return;
      if (soltado && Date.now() - soltado.t < 500 && Math.hypot(e.clientX - soltado.x, e.clientY - soltado.y) < 12) { soltado = null; return; }
      soltado = null;
      const t = e.target;
      if (t && t.closest && t.closest("[data-atras]") && menu.contains(t)) { atras(); return; }
      const b = t && t.closest && t.closest(".rueda-lista-item");
      if (b && menu.contains(b)) { elegir(+b.dataset.i); return; }
      if (!menu.contains(t)) cerrar();
    }
    function tecla(e) {
      e.stopPropagation();
      if (e.key === "Escape") { e.preventDefault(); cerrar(); return; }
      if ((e.key === "Backspace" || e.key === "ArrowLeft") && opc.padre) { e.preventDefault(); atras(); return; }
      if (/^[1-8]$/.test(e.key) && +e.key <= n) { e.preventDefault(); elegir(+e.key - 1); return; }
      if (e.key === "ArrowDown" || (e.key === "Tab" && !e.shiftKey)) {
        e.preventDefault();
        marcar(siguienteHabil(items, elegido < 0 ? 0 : elegido + 1, 1), true);
      } else if (e.key === "ArrowUp" || (e.key === "Tab" && e.shiftKey)) {
        e.preventDefault();
        marcar(siguienteHabil(items, elegido < 0 ? n - 1 : elegido - 1, -1), true);
      } else if (e.key === "Home") {
        e.preventDefault(); marcar(siguienteHabil(items, 0, 1), true);
      } else if (e.key === "End") {
        e.preventDefault(); marcar(siguienteHabil(items, n - 1, -1), true);
      } else if (e.key === "ArrowRight" && elegido >= 0 && items[elegido] && items[elegido].sub) {
        e.preventDefault(); elegir(elegido);
      } else if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        if (elegido >= 0) elegir(elegido);
      }
    }

    const tragar = (e) => { if (e.cancelable) e.preventDefault(); e.stopPropagation(); };
    const oyentes = [
      ["pointermove", mover],
      ["pointerup", (e) => { if (modo === "mantener") { tragar(e); soltar(e); } }],
      ["click", (e) => { tragar(e); clic(e); }],
      ["pointercancel", () => { if (modo === "mantener") modo = "clic"; }],
      ["pointerdown", (e) => { if (modo === "clic") { tragar(e); if (!menu.contains(e.target)) cerrar(); } }],
      ["contextmenu", tragar],
      ["keydown", tecla],
    ];
    oyentes.forEach(([t, f]) => document.addEventListener(t, f, { capture: true, passive: false }));
    const quitarSel = document.body.style.userSelect;
    document.body.style.userSelect = "none";
    actual = {
      velo, opc,
      cerrar(sinAviso) {
        setTimeout(() => oyentes.forEach(([t, f]) => document.removeEventListener(t, f, { capture: true })), 0);
        document.body.style.userSelect = quitarSel;
        menu.removeAttribute("role");
        menu.setAttribute("aria-hidden", "true");
        if (sinAviso) { velo.remove(); return; }
        velo.classList.add("saliendo");
        saliendo = velo;
        setTimeout(() => { if (saliendo === velo) saliendo = null; velo.remove(); }, 170);
        restaurarFoco(anterior);
        opc.alCerrar && opc.alCerrar();
      },
    };
    const primero = siguienteHabil(items, 0, 1);
    if (primero >= 0) marcar(primero, true);
    if (!interno && opc.tactil !== false) vibrar("abrir");
    requestAnimationFrame(() => menu.classList.add("abierta"));
    return actual;
  }

  function sector(i, n) {
    const paso = (Math.PI * 2) / n, a0 = -Math.PI / 2 - paso / 2 + i * paso, a1 = a0 + paso;
    const p = (r, a) => `${(r * Math.cos(a)).toFixed(2)},${(r * Math.sin(a)).toFixed(2)}`;
    const largo = paso > Math.PI ? 1 : 0;
    return `M${p(RI, a0)} L${p(R, a0)} A${R},${R} 0 ${largo} 1 ${p(R, a1)} L${p(RI, a1)} A${RI},${RI} 0 ${largo} 0 ${p(RI, a0)} Z`;
  }

  /** Zona visible de la pantalla (con el teclado del celular abierto es más chica). */
  function visible() {
    const vv = window.visualViewport;
    return vv ? { x: vv.offsetLeft, y: vv.offsetTop, w: vv.width, h: vv.height } : { x: 0, y: 0, w: innerWidth, h: innerHeight };
  }

  /** Caben 8: con más, las primeras 7 y «Más» con el resto (que a su vez se reparte igual). Antes se cortaba en 8 y lo
      demás no se podía alcanzar (en la hoja, «Más › Otras» perdía Imagen, Fijar, Formato condicional, Tabla dinámica…). */
  function paginar(items) {
    items = (items || []).filter(Boolean);
    if (items.length <= 8) return items;
    return items.slice(0, 7).concat([{ texto: "Más", icono: "more-horizontal", sub: items.slice(7) }]);
  }

  /**
   * Abre la rueda centrada en (x, y). items: [{ texto, icono | color, fn, off, marcado, mantener, sub: [...] }]
   * (hasta 8). opc.modo: "mantener" (se elige al soltar) o "clic". opc.alCerrar.
   */
  function abrir(x, y, items, opc, interno) {
    if (saliendo) { saliendo.remove(); saliendo = null; }
    cerrar(!!interno);
    opc = opc || {};
    items = paginar(items);
    if (!items.length) return null;
    if (usarLista()) return abrirLista(x, y, items, opc, interno);
    const n = items.length;
    const v = visible();
    const anterior = !interno ? document.activeElement : (opc._anterior || null);
    // Cerca de la orilla la rueda se recorre hacia adentro. Al mantener presionado, la dirección se
    // mide desde donde se puso el dedo (x0, y0), no desde el centro recorrido: si no, soltar sin
    // moverse elegía la acción que quedaba de ese lado (en el número de fila: "Fijar" o "Eliminar fila").
    const x0 = x, y0 = y;
    x = Math.min(v.x + v.w - R - 8, Math.max(v.x + R + 8, x));
    y = Math.min(v.y + v.h - R - 8, Math.max(v.y + R + 8, y));
    const velo = html(`<div class="rueda-velo" role="presentation">
      <div class="rueda${opc.padre ? " sub" : ""}" role="menu" aria-label="${esc(opc.titulo || "Acciones rápidas")}" style="left:${x}px;top:${y}px">
        <div class="rueda-fondo"></div>
        <svg class="rueda-svg" viewBox="${-R} ${-R} ${R * 2} ${R * 2}" width="${R * 2}" height="${R * 2}" aria-hidden="true">
          ${items.map((it, i) => `<path class="rueda-sector ${it.off ? "off" : ""}" data-i="${i}" d="${sector(i, n)}"/>`).join("")}
        </svg>
        ${items.map((it, i) => {
          const a = -Math.PI / 2 + (i * Math.PI * 2) / n, rm = (R + RI) / 2;
          return `<button class="rueda-item ${it.off ? "off" : ""} ${it.marcado ? "marcado" : ""}" role="menuitem" data-i="${i}" tabindex="-1" ${it.off ? 'aria-disabled="true"' : ""}
            style="--i:${i};left:${(rm * Math.cos(a)).toFixed(1)}px;top:${(rm * Math.sin(a)).toFixed(1)}px" aria-label="${esc(it.texto)}">
            ${it.color ? `<span class="rueda-muestra" style="background:${NOM.colorCss(it.color, "#5A5D6B")}"></span>` : ico(it.icono || "circle")}
            <small>${esc(it.corto || it.texto)}</small>${it.sub ? `<i class="rueda-mas"></i>` : ""}</button>`;
        }).join("")}
        <div class="rueda-centro">${opc.padre ? ico("chevron-left", "ico-s") : ""}<span data-etq aria-live="polite">${esc(opc.titulo || "Elige")}</span></div>
      </div></div>`);
    document.body.append(velo);
    const rueda = velo.firstElementChild;
    const etq = rueda.querySelector("[data-etq]");
    let elegido = -1, modo = opc.modo || "clic", seMovio = false, soltado = null;

    function marcar(i) {
      if (i >= 0) seMovio = true;
      if (i === elegido) return;
      elegido = i;
      velo.querySelectorAll("[data-i]").forEach((el) => el.classList.toggle("activo", +el.dataset.i === i));
      etq.textContent = i < 0 ? (opc.titulo || "Elige") : items[i].texto;
      if (i >= 0 && opc.tactil) vibrar("mover");
    }
    function porPunto(cx, cy, desdeInicio) {
      const dx = cx - (desdeInicio ? x0 : x), dy = cy - (desdeInicio ? y0 : y);
      if (Math.hypot(dx, dy) < MUERTA) return -1;
      let a = Math.atan2(dy, dx) + Math.PI / 2 + Math.PI / n;
      a = ((a % (Math.PI * 2)) + Math.PI * 2) % (Math.PI * 2);
      return Math.floor(a / ((Math.PI * 2) / n)) % n;
    }
    function elegir(i) {
      const it = items[i];
      if (!it || it.off) return;
      if (opc.tactil !== false) vibrar("elegir");
      // El submenú se abre en el mismo lugar; su centro regresa a la rueda anterior.
      if (it.sub) { abrir(x, y, it.sub, Object.assign({}, opc, { modo: "clic", titulo: it.texto, padre: { items, opc }, _anterior: anterior }), true); return; }
      // Al hacer la acción la rueda se va.
      cerrar();
      it.fn && it.fn();
      if (NOM.hito) NOM.hito("rueda");
    }
    function atras() { if (opc.padre) abrir(x, y, opc.padre.items, Object.assign({}, opc.padre.opc, { _anterior: anterior }), true); else cerrar(); }
    function mover(e) { if (modo === "mantener" || e.pointerType === "mouse") marcar(porPunto(e.clientX, e.clientY, modo === "mantener")); }
    function soltar(e) {
      if (modo !== "mantener") return;
      const i = porPunto(e.clientX, e.clientY, true);
      // Soltar sin haber ido hacia ninguna acción: se queda abierta para elegir con un toque.
      if (i < 0) { if (!seMovio) { modo = "clic"; soltado = { t: Date.now(), x: e.clientX, y: e.clientY }; marcar(-1); } else cerrar(); return; }
      elegir(i);
    }
    function clic(e) {
      if (modo !== "clic") return;
      // El clic que el navegador manda al soltar el mismo gesto que la abrió no cuenta.
      if (soltado && Date.now() - soltado.t < 500 && Math.hypot(e.clientX - soltado.x, e.clientY - soltado.y) < 12) { soltado = null; return; }
      soltado = null;
      const d = Math.hypot(e.clientX - x, e.clientY - y);
      if (d < RI && opc.padre) { atras(); return; }
      const i = porPunto(e.clientX, e.clientY);
      if (i < 0 || d > R + 40) { cerrar(); return; }
      elegir(i);
    }
    function tecla(e) {
      e.stopPropagation();
      if (e.key === "Escape") { e.preventDefault(); cerrar(); return; }
      if (e.key === "Backspace" && opc.padre) { e.preventDefault(); atras(); return; }
      if (/^[1-8]$/.test(e.key) && +e.key <= n) { e.preventDefault(); elegir(+e.key - 1); return; }
      if (e.key === "ArrowRight" || e.key === "ArrowDown" || (e.key === "Tab" && !e.shiftKey)) {
        e.preventDefault();
        marcar(siguienteHabil(items, elegido < 0 ? 0 : elegido + 1, 1));
      } else if (e.key === "ArrowLeft" || e.key === "ArrowUp" || (e.key === "Tab" && e.shiftKey)) {
        e.preventDefault();
        marcar(siguienteHabil(items, elegido < 0 ? n - 1 : elegido - 1, -1));
      } else if (e.key === "Home") {
        e.preventDefault();
        marcar(siguienteHabil(items, 0, 1));
      } else if (e.key === "End") {
        e.preventDefault();
        marcar(siguienteHabil(items, n - 1, -1));
      } else if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        if (elegido >= 0) elegir(elegido);
      }
      const b = velo.querySelector(`.rueda-item[data-i="${elegido}"]`);
      if (b) b.focus({ preventScroll: true });
    }

    // En el documento (fase de captura): con el dedo, el puntero queda capturado por el elemento
    // donde empezó el toque, así que los eventos no llegarían al velo.
    const tragar = (e) => { if (e.cancelable) e.preventDefault(); e.stopPropagation(); };
    const oyentes = [["pointermove", mover], ["pointerup", (e) => { tragar(e); soltar(e); }], ["click", (e) => { tragar(e); clic(e); }],
      // Si el sistema se queda con el toque (pulsación larga sobre texto), se elige con otro toque.
      ["pointercancel", () => { if (modo === "mantener") { modo = "clic"; marcar(-1); } }],
      ["pointerdown", (e) => { if (modo === "clic") tragar(e); }],
      ["contextmenu", tragar], ["keydown", tecla], ["touchmove", (e) => e.cancelable && e.preventDefault()]];
    oyentes.forEach(([t, f]) => document.addEventListener(t, f, { capture: true, passive: false }));
    const quitarSel = document.body.style.userSelect;
    document.body.style.userSelect = "none";
    actual = {
      velo, opc,
      cerrar(sinAviso) {
        // El clic que sigue a soltar el botón todavía debe tragarse; se quitan un instante después.
        setTimeout(() => oyentes.forEach(([t, f]) => document.removeEventListener(t, f, { capture: true })), 0);
        document.body.style.userSelect = quitarSel;
        velo.classList.add(sinAviso ? "cambiando" : "saliendo");
        setTimeout(() => velo.remove(), sinAviso ? 90 : 170);
        if (!sinAviso) {
          restaurarFoco(anterior);
          opc.alCerrar && opc.alCerrar();
        }
      },
    };
    if (opc.teclado) {
      const primero = siguienteHabil(items, 0, 1);
      if (primero >= 0) {
        marcar(primero);
        const b = velo.querySelector(`.rueda-item[data-i="${primero}"]`);
        if (b) b.focus({ preventScroll: true });
      }
    }
    if (!interno && opc.tactil !== false) vibrar("abrir");
    requestAnimationFrame(() => rueda.classList.add("abierta"));
    return actual;
  }

  function cerrar(sinAviso) { if (actual) { const a = actual; actual = null; a.cerrar(sinAviso); } }

  /**
   * Conecta la rueda a un elemento. items(e) devuelve las acciones para ese punto (o null).
   * opc.izquierdo: también al mantener el clic izquierdo; opc.tactil: también con el dedo.
   * opc.rearmar(e): true si moverse reinicia la espera en vez de cancelarla (pausa al seleccionar).
   * opc.tactilPorSistema: con el dedo manda la pulsación larga del sistema (texto seleccionable);
   *   opc.alPulsacionLarga(e) decide: "propia" (no dejar que el sistema seleccione) o "sistema".
   * opc.alPreparar(e): se llama justo antes de abrir (para guardar la selección).
   * opc.alAbrir(e, tactil) y opc.alCerrar(e, tactil): al abrir y al cerrar (también al elegir una acción).
   */
  function enlazar(el, items, opc) {
    opc = Object.assign({ izquierdo: true, tactil: true, ms: MANTENER_MS }, opc);
    let espera = null;
    const cancelar = () => { if (espera) { clearTimeout(espera.t); espera = null; } };
    const abrirEn = (ev, x, y, modo, tactil, teclado) => {
      opc.alPreparar && opc.alPreparar(ev);
      const l = items(ev);
      if (!l) return null;
      const r = abrir(x, y, l, { modo, titulo: l.titulo || opc.titulo, tactil, teclado: !!teclado, alCerrar: opc.alCerrar ? () => opc.alCerrar(ev, tactil) : undefined });
      if (r && opc.alAbrir) opc.alAbrir(ev, tactil);
      return r;
    };
    el.addEventListener("contextmenu", (e) => {
      if (e.shiftKey) return;   // menú nativo: corrector ortográfico, copiar enlace, etc.
      const tactil = e.pointerType === "touch" || e.pointerType === "pen" || !!(e.sourceCapabilities && e.sourceCapabilities.firesTouchEvents);
      if (tactil) {
        if (!opc.tactil || (opc.ignorar && opc.ignorar(e))) return;
        cancelar();
        if (actual) { e.preventDefault(); return; }   // ya la abrió la espera propia
        if (!opc.tactilPorSistema) { e.preventDefault(); abrirEn(e, e.clientX, e.clientY, "clic", true); return; }
        // Texto: la pulsación larga del sistema elige la palabra; la rueda sale con ella, en el dedo.
        const decide = opc.alPulsacionLarga ? opc.alPulsacionLarga(e) : "sistema";
        if (decide === "propia") { e.preventDefault(); abrirEn(e, e.clientX, e.clientY, "clic", true); }
        else { const x = e.clientX, y = e.clientY; setTimeout(() => { if (!actual) abrirEn(e, x, y, "clic", true); }, 140); }
        return;
      }
      e.preventDefault();
      // Teclado (tecla Menú o Mayús+F10): se abre donde está el cursor.
      if (e.button === 0 && e.clientX === 0 && e.clientY === 0) {
        const r = opc.puntoTeclado ? opc.puntoTeclado() : el.getBoundingClientRect();
        abrirEn(e, r.left + (r.width || 0) / 2, r.top + (r.height || 0) / 2, "clic", false, true);
      }
    });
    el.addEventListener("pointerdown", (e) => {
      if (actual) return;
      if (e.pointerType === "mouse" && e.button === 2 && !e.shiftKey) {
        const r = abrirEn(e, e.clientX, e.clientY, "mantener", false);
        if (r) e.preventDefault();
        return;
      }
      const tactil = e.pointerType !== "mouse";
      if ((e.button !== 0) || (tactil && (!opc.tactil || opc.tactilPorSistema)) || (!tactil && !opc.izquierdo)) return;
      if (opc.ignorar && opc.ignorar(e)) return;
      cancelar();
      const armar = (x, y) => setTimeout(() => {
        const ev = espera && espera.ev;
        espera = null;
        if (!ev || actual) return;
        if (!abrirEn(ev, x, y, "mantener", tactil)) return;
        // El clic que termina este gesto no debe llegar al elemento de abajo.
        const tragar = (c) => { c.stopPropagation(); c.preventDefault(); };
        document.addEventListener("click", tragar, { capture: true, once: true });
        setTimeout(() => document.removeEventListener("click", tragar, { capture: true }), 600);
      }, opc.ms);
      espera = { x: e.clientX, y: e.clientY, id: e.pointerId, ev: e, rearmar: !!(opc.rearmar && opc.rearmar(e)), t: armar(e.clientX, e.clientY) };
      espera.armar = armar;
    });
    const mover = (e) => {
      if (!espera || e.pointerId !== espera.id || Math.hypot(e.clientX - espera.x, e.clientY - espera.y) <= 7) return;
      if (!espera.rearmar) { cancelar(); return; }
      // Se sigue seleccionando: la cuenta empieza otra vez desde el punto nuevo.
      clearTimeout(espera.t);
      espera.x = e.clientX; espera.y = e.clientY; espera.ev = e;
      espera.t = espera.armar(e.clientX, e.clientY);
    };
    document.addEventListener("pointermove", mover, true);
    ["pointerup", "pointercancel", "dragstart"].forEach((t) => document.addEventListener(t, (e) => { if (espera && (t === "dragstart" || e.pointerId === espera.id)) cancelar(); }, true));
    el.addEventListener("wheel", cancelar, { passive: true });
    el.addEventListener("scroll", cancelar, true);
  }

  NOM.rueda = { abrir, cerrar: () => cerrar(), enlazar, abierta: () => !!actual, _paginar: paginar };
})();

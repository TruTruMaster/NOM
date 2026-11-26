/* NOM · Módulo 1: Gestión de cuentas (diagramas A1 Crear usuario y A2 Iniciar sesión). */
(function () {
  "use strict";
  const NOM = window.NOM;
  const { esc, ico, html, $, $$, api } = { ...NOM, api: NOM.api };

  function fuerza(c) {
    let n = 0;
    if (c.length >= 8) n++;
    if (c.length >= 12) n++;
    if (/[A-Z]/.test(c) && /[a-z]/.test(c)) n++;
    if (/\d/.test(c) && /[^A-Za-z0-9]/.test(c)) n++;
    return c ? Math.max(1, Math.min(4, n)) : 0;
  }
  const textoFuerza = ["", "Débil", "Aceptable", "Buena", "Excelente"];
  /** Tiempo límite de registro e inicio de sesión: cubre el arranque en frío de la base de datos. */
  const ESPERA_BD = 90000;

  /** El código para recuperar la contraseña: se muestra una sola vez (no hay correo ni administrador). */
  function mostrarCodigo(codigo, nombre) {
    return new Promise((listo) => {
      const cuerpo = html(`<div class="col" style="align-items:center;text-align:center">
        <span class="modo-insignia" style="width:56px;height:56px;border-radius:var(--r-l);background:var(--acento-tenue);color:var(--acento-tinta)">${ico("key-round", "ico-l")}</span>
        <p style="font-weight:400">Si algún día olvidas tu contraseña, con este código pones una nueva. <b>Guárdalo</b> (foto, papel o archivo): solo se muestra ahora y nadie más lo tiene, ni quien administra NOM.</p>
        <b class="qr-codigo" style="font-size:24px;letter-spacing:.08em">${esc(codigo)}</b>
        <div class="fila"><button class="btn chico" data-copiar>${ico("copy", "ico-s")} Copiar</button><button class="btn chico" data-bajar>${ico("download", "ico-s")} Guardar archivo</button></div></div>`);
      const pie = html(`<div style="display:contents"><span class="crece"></span><button class="btn primario" data-ok>Ya lo guardé</button></div>`);
      const h = NOM.hoja({ titulo: "Tu código de recuperación", icono: "key-round", cuerpo, pie, alCerrar: () => listo() });
      $("[data-copiar]", cuerpo).onclick = () => NOM.copiarTexto(codigo).then((ok) => NOM.aviso(ok ? "Código copiado" : "No se pudo copiar: anótalo o usa «Descargar».", { tipo: ok ? "ok" : "aviso", ms: ok ? undefined : 6000 }));
      $("[data-bajar]", cuerpo).onclick = () => NOM.formatos.descargar(`Codigo de recuperacion NOM - ${nombre || "cuenta"}.txt`, "text/plain", `Cuenta: ${nombre || ""}\r\nCódigo para recuperar la contraseña de NOM: ${codigo}\r\nGuárdalo en un lugar seguro.\r\n`);
      $("[data-ok]", pie).onclick = () => h.cerrar();
    });
  }

  /** ¿Olvidaste tu contraseña?: usuario + código de recuperación + contraseña nueva. */
  function recuperar(nombre) {
    const cuerpo = html(`<form class="col" novalidate>
      <p class="suave" style="font-weight:400">Usa el código que te dio NOM al crear la cuenta (o el último que generaste en Ajustes).</p>
      <p class="suave chico" style="font-weight:400">${ico("hard-drive", "ico-s")} Lo que está en tus espacios no cambia. Lo que solo está guardado en un equipo (documentos de “${esc(NOM.etiquetaEquipo())}” y cambios sin enviar) va cifrado con tu contraseña de ahora: al entrar con la nueva te la pediremos una vez para recuperarlo, así que si la recuerdas, mejor cámbiala desde tu cuenta.</p>
      <div class="alerta" data-alerta hidden role="alert"></div>
      <div class="campo"><label for="rn">Usuario</label><div class="entrada">${ico("user")}<input id="rn" name="nombre" autocomplete="username" autocapitalize="none" value="${esc(nombre || "")}"></div></div>
      <div class="campo"><label for="rk">Código de recuperación</label><div class="entrada">${ico("key-round")}<input id="rk" name="codigo" autocapitalize="characters" autocomplete="off" spellcheck="false" placeholder="ABCD-EFGH-JKMN-PQRS"></div></div>
      <div class="campo"><label for="rp">Contraseña nueva</label><div class="entrada">${ico("lock")}<input id="rp" name="nueva" type="password" autocomplete="new-password" minlength="8"></div></div>
    </form>`);
    const pie = html(`<div style="display:contents"><button class="btn" data-no type="button">Cancelar</button><button class="btn primario" data-si>Cambiar contraseña</button></div>`);
    const h = NOM.hoja({ titulo: "Recuperar la contraseña", icono: "key-round", cuerpo, pie });
    const al = $("[data-alerta]", cuerpo);
    const enviar = async (e) => {
      e && e.preventDefault();
      const b = $("[data-si]", pie);
      b.disabled = true;
      try {
        const r = await api.post("/api/cuentas/recuperar", { nombre: cuerpo.nombre.value.trim(), codigo: cuerpo.codigo.value, nueva: cuerpo.nueva.value }, { sin401: true, ms: ESPERA_BD });
        // Fase 15: al entrar, la llave privada se abre con el código de antes y se envuelve con lo nuevo.
        try { sessionStorage.setItem("nom.e2e.recupera", JSON.stringify({ nombre: cuerpo.nombre.value.trim(), viejo: cuerpo.codigo.value, nuevo: r.recuperacion })); } catch (x) { /* sin almacenamiento */ }
        h.cerrar(true);
        NOM.aviso("Listo: ya puedes entrar con tu contraseña nueva", { tipo: "ok", ms: 5000 });
        await mostrarCodigo(r.recuperacion, cuerpo.nombre.value.trim());
      } catch (x) {
        al.className = "alerta " + (x.estado === 423 ? "aviso" : "error"); al.hidden = false;
        al.innerHTML = `${ico("triangle-alert")}<span>${esc(x.message)}</span>`;
        b.disabled = false;
      }
    };
    cuerpo.onsubmit = enviar;
    $("[data-si]", pie).onclick = enviar;
    $("[data-no]", pie).onclick = () => h.cerrar();
  }

  /** Alerta dentro de una hoja (error del servidor o cuenta bloqueada). */
  function alertaEn(al, x) {
    al.className = "alerta " + (x.estado === 423 ? "aviso" : "error"); al.hidden = false;
    al.innerHTML = `${ico(x.estado === 423 ? "lock" : "triangle-alert")}<span>${esc(x.red ? "Sin conexión: revisa el Wi-Fi o los datos e intenta de nuevo." : x.message)}</span>`;
  }

  /**
   * Cambiar la contraseña sabiendo la actual. El servidor cierra las demás sesiones y da una nueva
   * para este equipo; la llave de lo guardado aquí se envuelve con la contraseña nueva, así que
   * los documentos de este equipo y los cambios sin enviar siguen abriendo.
   */
  function cambiarContrasena() {
    const u = api.usuario() || {};
    const cuerpo = html(`<form class="col" novalidate>
      <p class="suave" style="font-weight:400">Tus otros equipos tendrán que volver a entrar con la contraseña nueva. Lo guardado en este no se pierde.</p>
      <div class="alerta" data-alerta hidden role="alert"></div>
      <input type="text" name="usuario" autocomplete="username" value="${esc(u.nombre || "")}" hidden>
      <div class="campo"><label for="ca">Contraseña actual</label><div class="entrada">${ico("lock")}<input id="ca" name="actual" type="password" autocomplete="current-password" autofocus></div></div>
      <div class="campo"><label for="cn">Contraseña nueva</label><div class="entrada">${ico("key-round")}<input id="cn" name="nueva" type="password" autocomplete="new-password" minlength="8"></div>
        <div class="ayuda">Al menos 8 letras, números o signos.</div></div>
      <div class="campo"><label for="cr">Repite la contraseña nueva</label><div class="entrada">${ico("key-round")}<input id="cr" name="repite" type="password" autocomplete="new-password" minlength="8"></div></div>
    </form>`);
    const pie = html(`<div style="display:contents"><button class="btn" data-no type="button">Cancelar</button><button class="btn primario" data-si>Cambiar contraseña</button></div>`);
    const h = NOM.hoja({ titulo: "Cambiar contraseña", icono: "key-round", cuerpo, pie });
    const al = $("[data-alerta]", cuerpo);
    const enviar = async (e) => {
      e && e.preventDefault();
      const actual = cuerpo.actual.value, nueva = cuerpo.nueva.value;
      if (!actual) return alertaEn(al, { message: "Escribe tu contraseña actual" });
      if (nueva.length < 8) return alertaEn(al, { message: "La contraseña nueva necesita al menos 8 letras, números o signos" });
      if (nueva !== cuerpo.repite.value) return alertaEn(al, { message: "Las dos contraseñas nuevas no son iguales" });
      if (nueva === actual) return alertaEn(al, { message: "La contraseña nueva es igual a la actual" });
      const b = $("[data-si]", pie);
      b.disabled = true;
      try {
        const r = await api.put("/api/cuentas/contrasena", { actual, nueva }, { sin401: true, ms: ESPERA_BD });
        api.fijarSesion(r.token, r.usuario);
        if (NOM.e2e) NOM.e2e.reenvolver(nueva).catch((x) => console.warn(x));
        try { await NOM.boveda.cambiarContrasena(nueva); }
        catch (x) { console.error(x); NOM.aviso("La contraseña cambió, pero no se pudo actualizar lo guardado en este equipo: la próxima vez que entres te pediremos la anterior para abrirlo.", { tipo: "aviso", ms: 9000 }); }
        h.cerrar(true);
        NOM.aviso("Contraseña cambiada. En tus otros equipos vuelve a entrar con la nueva.", { tipo: "ok", ms: 6000 });
      } catch (x) {
        alertaEn(al, x);
        b.disabled = false;
      }
    };
    cuerpo.onsubmit = enviar;
    $("[data-si]", pie).onclick = enviar;
    $("[data-no]", pie).onclick = () => h.cerrar();
  }

  /**
   * Borrar la cuenta: dice qué se borra, pide la contraseña y escribir el nombre de usuario. Al
   * terminar se borra también lo suyo de este equipo y se vuelve a la pantalla de entrada.
   */
  async function borrarCuenta() {
    const u = api.usuario() || {};
    let alm = null;
    try { alm = await api.get("/api/cuentas/almacenamiento"); } catch (e) { /* se explica sin las cifras */ }
    const nEsp = alm ? alm.espacios.length : null, nDoc = alm ? alm.documentos : null;
    const cuerpo = html(`<form class="col" novalidate>
      <div class="alerta error">${ico("triangle-alert")}<span><b>No se puede deshacer.</b> Nadie, ni quien administra NOM, puede recuperar la cuenta después.</span></div>
      <div class="lista">
        <div class="lista-item">${ico("trash-2")}<span class="crece"><b>Se borran ${nEsp === null ? "tus espacios" : nEsp === 1 ? "tu espacio" : "tus " + nEsp + " espacios"}${nDoc ? " y " + (nDoc === 1 ? "su documento" : "sus " + nDoc + " documentos") : ""}</b><small>Con todas sus versiones, también para las personas con quienes los compartes. Si quieres conservar algo, descárgalo antes.</small></span></div>
        <div class="lista-item">${ico("log-out")}<span class="crece"><b>Sales de los espacios de otras personas</b><small>Lo que escribiste ahí se queda: es de ese espacio.</small></span></div>
        <div class="lista-item">${ico(api.empaquetada ? "smartphone" : "hard-drive")}<span class="crece"><b>Se borra lo tuyo de este equipo</b><small>Los documentos de “${esc(NOM.etiquetaEquipo())}” y los cambios sin enviar de esta cuenta.</small></span></div>
        <div class="lista-item">${ico("archive")}<span class="crece"><b>Los respaldos se borran solos</b><small>El servidor guarda copias de los últimos días: tus datos desaparecen de ellas a más tardar en 14 días.</small></span></div>
      </div>
      <div class="alerta" data-alerta hidden role="alert"></div>
      <input type="text" name="usuario" autocomplete="username" value="${esc(u.nombre || "")}" hidden>
      <div class="campo"><label for="bc">Tu contraseña</label><div class="entrada">${ico("lock")}<input id="bc" name="contrasena" type="password" autocomplete="current-password"></div></div>
      <div class="campo"><label for="bn">Para confirmar, escribe tu nombre de usuario: <b>${esc(u.nombre || "")}</b></label><div class="entrada">${ico("user")}<input id="bn" name="nombre" autocomplete="off" autocapitalize="none" spellcheck="false"></div></div>
    </form>`);
    const pie = html(`<div style="display:contents"><button class="btn" data-no type="button">Cancelar</button><button class="btn peligro lleno" data-si>Borrar mi cuenta para siempre</button></div>`);
    const h = NOM.hoja({ titulo: "Borrar mi cuenta", icono: "trash-2", cuerpo, pie });
    const al = $("[data-alerta]", cuerpo);
    const enviar = async (e) => {
      e && e.preventDefault();
      if (!cuerpo.contrasena.value) return alertaEn(al, { message: "Escribe tu contraseña" });
      if (cuerpo.nombre.value.trim().toLowerCase() !== String(u.nombre || "").toLowerCase()) return alertaEn(al, { message: "Escribe tu nombre de usuario tal cual para confirmar" });
      const b = $("[data-si]", pie);
      b.disabled = true;
      try {
        await api.post("/api/cuentas/eliminar", { contrasena: cuerpo.contrasena.value }, { sin401: true, ms: ESPERA_BD });
        h.cerrar(true);
        await NOM.app.cuentaBorrada(u.id);
        NOM.aviso("Tu cuenta y todo lo que tenía se borraron.", { tipo: "ok", ms: 7000 });
      } catch (x) {
        alertaEn(al, x);
        b.disabled = false;
      }
    };
    cuerpo.onsubmit = enviar;
    $("[data-si]", pie).onclick = enviar;
    $("[data-no]", pie).onclick = () => h.cerrar();
  }

  /**
   * Si la contraseña cambió sin este equipo (se restableció con el código, o se cambió en otro),
   * lo que solo vivía aquí quedó apartado y cifrado con la anterior: se ofrece recuperarlo.
   * `alEntrar`: recién se inició sesión (se ofrece siempre); si no, como mucho una vez al día.
   * `pedido`: la persona lo abrió desde el menú. Cuando se ofrece solo, espera a que no haya otra
   * ventana abierta (por ejemplo, la de unirse a un espacio con un código): abrir esta la cerraría.
   */
  async function ofrecerRecuperar(alEntrar, pedido, intento) {
    if (!NOM.boveda || !NOM.boveda.activa()) return;
    if (!pedido && document.querySelector(".velo")) {
      if ((intento || 0) < 30) setTimeout(() => ofrecerRecuperar(alEntrar, false, (intento || 0) + 1).catch((e) => console.error(e)), 4000);
      return;
    }
    const r = await NOM.boveda.porRecuperar();
    NOM.app.viejaPendiente = r.hay;
    if (!r.hay) return;
    const uid = NOM.boveda.activa(), clave = "nom.viejaPospuesta." + uid;
    if (alEntrar === false && Date.now() - NOM.leer(clave, 0) < 86400000) return;
    const partes = [];
    if (r.docs) partes.push(r.docs === 1 ? "1 documento" : r.docs + " documentos");
    if (r.colas) partes.push("cambios sin enviar de " + (r.colas === 1 ? "1 documento" : r.colas + " documentos"));
    const cuerpo = html(`<form class="col" novalidate>
      <p style="font-weight:400">En este equipo hay <b>${esc(partes.join(" y "))}</b> que se guardaron con tu contraseña anterior. Siguen aquí, cifrados: escribe esa contraseña para recuperarlos.</p>
      <div class="alerta" data-alerta hidden role="alert"></div>
      <div class="campo"><label for="va">Tu contraseña anterior</label><div class="entrada">${ico("lock")}<input id="va" name="anterior" type="password" autocomplete="off" autofocus></div>
        <div class="ayuda">No se envía a ningún lado: solo abre lo guardado en este equipo.</div></div>
    </form>`);
    const pie = html(`<div style="display:contents"><button class="btn plano peligro" data-borrar type="button">Borrarlos</button><span class="crece"></span><button class="btn" data-no type="button">Ahora no</button><button class="btn primario" data-si>Recuperar</button></div>`);
    let hecho = false;
    const h = NOM.hoja({ titulo: "Guardado con tu contraseña anterior", icono: "key-round", cuerpo, pie, alCerrar: () => { if (!hecho) NOM.guardar(clave, Date.now()); } });
    const al = $("[data-alerta]", cuerpo);
    const enviar = async (e) => {
      e && e.preventDefault();
      if (!cuerpo.anterior.value) return alertaEn(al, { message: "Escribe la contraseña que tenías antes" });
      const b = $("[data-si]", pie);
      b.disabled = true;
      try {
        const x = await NOM.boveda.recuperarCon(cuerpo.anterior.value);
        hecho = true; NOM.quitar(clave);
        h.cerrar(true);
        const que = [];
        if (x.docs) que.push(x.docs === 1 ? "1 documento" : x.docs + " documentos");
        if (x.cambios) que.push(x.cambios === 1 ? "1 cambio sin enviar" : x.cambios + " cambios sin enviar");
        NOM.aviso(que.length ? "Se recuperó: " + que.join(" y ") + "." : "No había nada nuevo que recuperar.", { tipo: "ok", ms: 7000 });
        if (x.cambios && NOM.cola) NOM.cola.arrancar();
        NOM.app.viejaPendiente = x.quedan > 0;
        NOM.app.refrescar();
        // Quedó algo guardado con otra contraseña todavía más vieja: se vuelve a ofrecer.
        if (x.quedan) setTimeout(() => ofrecerRecuperar(true), 600);
      } catch (x) {
        alertaEn(al, x);
        b.disabled = false; cuerpo.anterior.select();
      }
    };
    cuerpo.onsubmit = enviar;
    $("[data-si]", pie).onclick = enviar;
    $("[data-no]", pie).onclick = () => h.cerrar();
    $("[data-borrar]", pie).onclick = async () => {
      // La confirmación abre otra hoja (cierra esta): si dice que no, se vuelve a ofrecer.
      hecho = true;
      const si = await NOM.confirmar("Borrar lo guardado con tu contraseña anterior", `Se borrarán de este equipo ${partes.join(" y ")}. Sin esa contraseña nadie puede abrirlos, y borrados no se pueden recuperar.`, { si: "Borrarlos", peligro: true });
      if (!si) return ofrecerRecuperar(true, true);
      await NOM.boveda.descartarVieja();
      NOM.quitar(clave);
      NOM.app.viejaPendiente = false;
      NOM.aviso("Se borró lo guardado con tu contraseña anterior.", { tipo: "ok" });
    };
  }

  /** "¿Tienes un código del salón?": QR con la cámara o código escrito. */
  async function entrarConCodigo() {
    if (!NOM.union) return;
    if (NOM.qr && NOM.qr.hayCamara()) return NOM.union.escanear();
    const c = await NOM.pedirTexto("Código para unirse", { etiqueta: "El código que te dio el propietario del espacio", ph: "ABCD-2345", si: "Entrar", max: 60 });
    if (c) NOM.union.usar(c);
  }

  function campoContrasena(id, auto) {
    return `<div class="campo">
      <label for="${id}">Contraseña</label>
      <div class="entrada">${ico("lock")}
        <input id="${id}" type="password" autocomplete="${auto}" required minlength="${auto === "new-password" ? 8 : 1}">
        <button type="button" class="icono-btn" data-ver aria-label="Mostrar contraseña" data-micro="Mostrar u ocultar|Revisa lo que escribiste antes de enviar.">${ico("eye")}</button>
      </div>
      <div class="ayuda" data-mayus hidden>${ico("triangle-alert", "ico-s")} Bloq Mayús está activado</div>
    </div>`;
  }

  /** El código de recuperación recién creado (para envolver también con él la llave privada). */
  let codigoNuevo = null;

  function pintar(raiz, opc) {
    opc = opc || {};
    let modo = opc.modo || "entrar";
    raiz.innerHTML = "";
    const vista = html(`
      <main class="bienvenida">
        <div class="bienvenida-col">
          <div class="marca bienvenida-marca"><svg class="logo"><use href="#logo"/></svg>NOM</div>
          <div class="tarjeta-cuenta caja" data-tarjeta></div>
          <p class="bienvenida-que">Documentos de texto, hojas de cálculo, presentaciones y diagramas para trabajar en equipo, con Internet o en el mismo Wi-Fi.</p>
          <footer class="pie-legal">
            <nav aria-label="Ayuda y legal">
              ${[["", "Ayuda"], ["preguntas-frecuentes", "Preguntas frecuentes"], ["legal/privacidad", "Privacidad"], ["legal/terminos", "Términos"], ["legal", "Todo lo legal"]]
                .map(([r, t]) => `<a href="${NOM.urlAyuda(r)}" data-ayuda="${r}">${t}</a>`).join("")}
            </nav>
          </footer>
        </div>
      </main>`);
    raiz.append(vista);
    const tarjeta = $("[data-tarjeta]", vista);

    // H-8: si la persona pidió «Cambiar dónde trabaja tu equipo», decide ella (no se conecta sola).
    if (api.empaquetada && !api.servidor()) { const cambiar = !!NOM.leer("nom.eligeServidor", false); NOM.quitar("nom.eligeServidor"); return pintarServidor(tarjeta, () => pintar(raiz, opc), { cambiar }); }

    function render() {
      tarjeta.innerHTML = "";
      tarjeta.append(modo === "entrar" ? formEntrar() : formRegistro());
    }

    // ------------------------------------------------------------ fase 13: el teléfono que comparte
    /** En un teléfono que comparte no hay cuentas con contraseña: alias (con el código o el QR de un
        espacio) o la cuenta de la nube si este equipo tiene Internet. El dueño entra solo. */
    function pintarTelefono() {
      tarjeta.innerHTML = "";
      const propio = NOM.anfitrion && NOM.anfitrion.esLocal();
      if (propio) {
        const f = html(`<div class="col" style="align-items:center;text-align:center">
          <div class="burbuja">${ico("smartphone")}</div>
          <h2>Entrando a tu teléfono…</h2>
          <p class="suave" style="font-weight:400">Este teléfono comparte sin Internet. Aquí eres su dueño, con un alias que se elige al azar: así lo verán los demás.</p>
          <div class="alerta error" data-alerta hidden role="alert"></div>
          <button class="btn primario ancho" data-dueno hidden>${ico("log-in")} Entrar a mi teléfono</button></div>`);
        tarjeta.append(f);
        enlazarComunes(f);
        const b = $("[data-dueno]", f), al = $("[data-alerta]", f);
        const intentar = async () => {
          b.disabled = true;
          try {
            const r = await NOM.acceso.entrarComoDueno();
            NOM.app.alIniciarSesion(r);
          } catch (e) {
            al.hidden = false; al.className = "alerta error";
            al.innerHTML = `${ico("triangle-alert")}<span>${esc(e.message)}</span>`;
            b.hidden = false; b.disabled = false;
          }
        };
        b.onclick = intentar;
        intentar();
        return;
      }
      const f = html(`<div class="col">
        <div><h2>Un teléfono que comparte</h2><p class="suave" style="font-weight:400;margin-top:4px">Aquí no se crean cuentas ni se escribe contraseña: entra a un espacio con su <b>código</b> o su <b>QR</b> y te toca un alias al azar. Si este equipo tiene Internet, también puedes entrar con tu cuenta.</p></div>
        <div class="alerta" data-alerta hidden role="alert"></div>
        <div class="col" data-guardados></div>
        ${NOM.qr && NOM.qr.hayCamara() ? `<button type="button" class="btn primario ancho" data-escanear>${ico("qr-code")} Escanear el código QR</button>` : ""}
        <button type="button" class="btn ${NOM.qr && NOM.qr.hayCamara() ? "" : "primario"} ancho" data-codigo>${ico("key-round")} Tengo un código de invitación</button>
        <button type="button" class="btn ancho" data-mi-cuenta hidden>${ico("cloud")} <span class="btn-sub">Entrar con mi cuenta<small>Este equipo tiene Internet: lo tuyo queda a tu nombre</small></span></button>
        <p class="suave chico" data-sin-internet style="text-align:center">${ico("wifi-off", "ico-s")} Revisando si este equipo tiene Internet…</p>
        <div class="separador" aria-hidden="true">o</div>
        <button type="button" class="btn ancho" data-sin-cuenta data-micro="Usar sin cuenta|Trabaja sin registrarte. Tus documentos se guardan solo ${NOM.aqui().en}: si ${NOM.aqui().borrar} se pierden, salvo que los descargues.">${ico("hard-drive")} <span class="btn-sub">Usar sin cuenta<small>Tus documentos se quedan solo ${NOM.aqui().en}</small></span></button>
      </div>`);
      tarjeta.append(f);
      enlazarComunes(f);
      // Los alias que este equipo ya tiene en este teléfono: se vuelve a entrar sin el código.
      NOM.acceso.infoServidor().then((i) => {
        const ls = NOM.acceso.aliasDeEsteServidor(i.servidor), cont = $("[data-guardados]", f);
        if (!ls.length || !cont) return;
        cont.innerHTML = `<span class="etiqueta">Ya entraste desde este equipo</span>` + ls.slice(0, 4).map((g, k) => `<button type="button" class="btn ancho" data-g="${k}">${ico("user-check")} <span class="btn-sub">Seguir como ${esc(g.nombre)}<small>${esc(g.espacio || "Espacio")}</small></span></button>`).join("");
        $$("[data-g]", cont).forEach((b) => (b.onclick = async () => {
          b.disabled = true;
          try {
            const r = await NOM.acceso.volverConAlias(ls[+b.dataset.g]);
            NOM.app.alIniciarSesion(r);
          } catch (e) {
            b.disabled = false;
            const al = $("[data-alerta]", f);
            al.hidden = false; al.className = "alerta error";
            al.innerHTML = `${ico("triangle-alert")}<span>${esc(e.message)}</span>`;
          }
        }));
      }).catch(() => {});
      const esc2 = $("[data-escanear]", f);
      if (esc2) esc2.onclick = () => NOM.union && NOM.union.escanear();
      $("[data-codigo]", f).onclick = async () => {
        const c = await NOM.pedirTexto("Código para unirse", { etiqueta: "El código que te dio quien comparte el espacio", ph: "ABCD-2345", si: "Entrar", max: 60 });
        if (c && NOM.union) NOM.union.usar(c);
      };
      const mi = $("[data-mi-cuenta]", f), sinInt = $("[data-sin-internet]", f);
      NOM.acceso.nubeAlcanzable().then((ok) => {
        if (!mi.isConnected) return;
        mi.hidden = !ok;
        sinInt.innerHTML = ok ? "" : `${ico("wifi-off", "ico-s")} Este equipo no tiene Internet: entra con el código o el QR de un espacio.`;
      });
      mi.onclick = async () => {
        const r = await NOM.acceso.dialogoCuenta();
        if (!r) return;
        NOM.app.alIniciarSesion(r);
      };
    }

    function cabecera(titulo, sub) {
      return `<div><h2>${titulo}</h2><p class="suave" style="margin-top:4px">${sub}</p></div>`;
    }
    const SEPARADOR = `<div class="separador" aria-hidden="true">o</div>`;

    function enlazarComunes(f) {
      // F.5 La ayuda, las preguntas frecuentes y lo legal están en el pie de la pantalla (ver pintar).
      f.querySelectorAll("[data-sin-cuenta]").forEach((b) => (b.onclick = () => NOM.app.sinCuenta()));
      f.querySelectorAll("[data-ir]").forEach((b) => (b.onclick = () => { if (modo === b.dataset.ir) return; modo = b.dataset.ir; render(); }));
      f.querySelectorAll("[data-ver]").forEach((b) => (b.onclick = () => {
        const i = b.parentElement.querySelector("input");
        const ver = i.type === "password";
        i.type = ver ? "text" : "password";
        b.innerHTML = ico(ver ? "eye-off" : "eye");
        b.setAttribute("aria-label", ver ? "Ocultar contraseña" : "Mostrar contraseña");
        i.focus();
      }));
      f.querySelectorAll('input[type="password"]').forEach((i) => {
        const avisoMayus = i.closest(".campo").querySelector("[data-mayus]");
        const rev = (e) => { if (e.getModifierState) avisoMayus.hidden = !e.getModifierState("CapsLock"); };
        i.addEventListener("keyup", rev); i.addEventListener("keydown", rev);
      });
      // En la red local se viaja por http:// (ahí no hay forma de dar un certificado de confianza):
      // se avisa para que se use solo en un Wi-Fi conocido. En el propio equipo no hace falta.
      const dir = api.empaquetada ? api.servidor() : location.origin;
      const equipo = (dir.replace(/^https?:\/\//i, "").match(/^\[[^\]]*\]|^[^/:]+/) || [""])[0].toLowerCase();
      if (/^http:/i.test(dir) && !/^(localhost|127(\.\d+){3}|\[::1\])$/.test(equipo))
        f.append(html(`<p class="servidor-pie" data-sin-cifrar>${ico("triangle-alert", "ico-s")}<span><b>Conexión sin cifrar (red local).</b> Úsala solo en un Wi-Fi de confianza: quien esté en la misma red podría ver lo que se envía.</span></p>`));
      // En la app: a qué servidor se conecta, dicho en palabras, y cómo cambiarlo (para equipos que
      // usan NOM en su propia red, sin Internet).
      if (NOM.anfitrion && NOM.anfitrion.esLocal()) {
        const p = NOM.anfitrion.pie();
        if (p) f.append(p);
      } else if (api.empaquetada) {
        const srv = api.servidor().replace(/^https?:\/\//, "");
        const internet = /(^|\.)nom\.dpdns\.org$|(^|\.)nom\.qd\.je$|\.ts\.net$/.test(srv.replace(/\/.*$/, ""));
        // H-8: a un equipo del Wi-Fi se entra con una cuenta de ESE equipo. Se dice arriba y en grande,
        // para que nadie escriba ahí la contraseña de su cuenta de Internet.
        if (!internet && !f.querySelector("[data-aviso-wifi]")) {
          (f.querySelector("[data-alerta]") || f.firstElementChild).before(html(`<div class="alerta aviso" data-aviso-wifi>${ico("wifi")}<span>Vas a entrar a <b>un equipo de este Wi-Fi</b> (${esc(srv)}), no a NOM en Internet. Aquí las cuentas son de ese equipo: <b>no escribas la contraseña de tu cuenta de Internet</b> si no conoces a quien lo comparte.</span></div>`));
        }
        const s = html(`<p class="servidor-pie">${ico(internet ? "globe" : "wifi", "ico-s")}<span>Trabajas en <b>${internet ? "NOM en Internet" : "un equipo del Wi-Fi · " + esc(srv)}</b></span>
          <button type="button" data-micro="Cambiar dónde trabajas|Si tu equipo usa NOM en otro lugar (en Internet, o en una computadora o teléfono del mismo Wi-Fi), elígelo aquí.">Cambiar</button></p>`);
        // DEF-47: al terminar de elegir se vuelve a decidir TODA la pantalla (antes regresaba al formulario
        // de usuario y contraseña aunque lo elegido fuera este mismo teléfono, donde se entra como dueño).
        $("button", s).onclick = () => pintarServidor(tarjeta, () => pintar(raiz, opc), { cambiar: true });
        f.append(s);
      } else {
        // En un navegador normal (no la app): un enlace para descargar la app de Android. El
        // archivo lo entrega la red de GitHub (Fastly), no la conexión del servidor, y
        // ".../releases/latest/..." trae siempre la versión más reciente que se haya publicado.
        f.append(html(`<p class="servidor-pie">${ico("smartphone", "ico-s")}<span>¿Usas Android? <a href="https://github.com/TruTruMaster/NOM/releases/latest/download/NOM.apk" rel="noopener" data-micro="Descargar la app para Android|Se descarga desde la red de GitHub, así que no gasta los datos del servidor. Ábrela en el teléfono para instalarla; trae siempre la versión más reciente."><b>Descarga la app</b></a> · siempre la más reciente</span></p>`));
      }
    }

    function mostrarError(f, err) {
      const caja = $("[data-alerta]", f);
      f.querySelectorAll(".entrada").forEach((x) => x.classList.remove("invalida"));
      f.querySelectorAll(".error-campo").forEach((x) => x.remove());
      clearInterval(f._cuenta);
      if (err.campo) {
        const i = f.querySelector(`[name="${err.campo}"]`);
        if (i) {
          i.closest(".entrada").classList.add("invalida");
          i.closest(".campo").append(html(`<div class="error-campo">${ico("triangle-alert", "ico-s")} ${esc(err.message)}</div>`));
          i.focus();
          caja.hidden = true;
          return;
        }
      }
      const bloqueo = err.estado === 423 && err.detalle && err.detalle.bloqueadoHasta;
      caja.className = "alerta " + (bloqueo ? "aviso" : "error");
      caja.hidden = false;
      const pintarMsg = () => {
        let m = err.message;
        if (bloqueo) {
          const s = Math.max(0, Math.round((err.detalle.bloqueadoHasta - Date.now()) / 1000));
          m = s > 0 ? `Por seguridad, la cuenta está bloqueada. Podrás intentar de nuevo en ${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}.` : "Ya puedes intentar de nuevo.";
          if (s <= 0) { clearInterval(f._cuenta); $("button[type=submit]", f).disabled = false; }
        }
        caja.innerHTML = `${ico(bloqueo ? "lock" : "triangle-alert")}<span>${esc(m)}</span>`;
      };
      pintarMsg();
      if (bloqueo) { $("button[type=submit]", f).disabled = true; f._cuenta = setInterval(pintarMsg, 1000); }
      const t = $("[data-tarjeta]", vista);
      t.classList.remove("sacudir"); void t.offsetWidth; t.classList.add("sacudir");
    }

    /** Abre la sesión (también justo después de registrarse). */
    async function iniciarSesion(nombre, contrasena, nuevo) {
      const r = await api.post("/api/cuentas/sesion", { nombre, contrasena }, { sin401: true, ms: ESPERA_BD });
      NOM.guardar("nom.ultimoNombre", nombre);
      api.fijarSesion(r.token, r.usuario);
      // La contraseña abre la bóveda cifrada de esta cuenta en este equipo (y no se guarda).
      try { await NOM.boveda.abrir(r.usuario.id, contrasena); }
      catch (e) { console.error(e); NOM.aviso("Este navegador no permite guardar documentos cifrados; se trabajará solo en línea.", { tipo: "aviso", ms: 6000 }); }
      NOM.aplicarPreferencias(r.preferencias);
      // Fase 15: el par de llaves de la cuenta (se crea la primera vez) y la privada a la bóveda.
      if (NOM.e2e) {
        let rec = null;
        try { rec = JSON.parse(sessionStorage.getItem("nom.e2e.recupera") || "null"); } catch (x) { rec = null; }
        if (rec && String(rec.nombre).toLowerCase() === String(nombre).toLowerCase()) {
          try { sessionStorage.removeItem("nom.e2e.recupera"); } catch (x) { /* nada */ }
          NOM.e2e.recuperarConCodigo(rec.viejo, contrasena, rec.nuevo).then((ok) => { if (!ok) NOM.e2e.prepararCuenta(contrasena); }).catch((x) => console.warn(x));
        } else NOM.e2e.prepararCuenta(contrasena, codigoNuevo).catch((x) => console.warn(x));
        codigoNuevo = null;
      }
      NOM.app.alIniciarSesion(r);
      if (!nuevo) setTimeout(() => revisarLegal(r), 900);
      // Ya dentro: ¿quedó algo en este equipo guardado con la contraseña anterior?
      setTimeout(() => ofrecerRecuperar(true).catch((e) => console.error(e)), nuevo ? 2600 : 1700);
    }

    async function enviando(f, fn) {
      const b = $("button[type=submit]", f);
      const original = b.innerHTML;
      b.disabled = true;
      b.innerHTML = `<span class="cargando"></span>`;
      // Si el servidor tarda en contestar, se avisa que sigue conectando.
      const lento = setTimeout(() => { if (b.isConnected) b.innerHTML = `<span class="cargando"></span> Conectando con la base de datos…`; }, 4000);
      try { await fn(); } catch (e) { mostrarError(f, e); }
      finally { clearTimeout(lento); if (b.isConnected && !f._cuenta) { b.disabled = false; b.innerHTML = original; } else if (b.isConnected) b.innerHTML = original; }
    }

    function formEntrar() {
      const f = html(`<form class="col" novalidate>
        ${cabecera("Hola de nuevo", "Entra para ver tus espacios y documentos.")}
        <div class="alerta" data-alerta hidden role="alert"></div>
        <div class="campo"><label for="u">Usuario</label>
          <div class="entrada">${ico("user")}<input id="u" name="nombre" autocomplete="username" autocapitalize="none" spellcheck="false" required value="${esc(opc.nombre || NOM.leer("nom.ultimoNombre", ""))}"></div></div>
        ${campoContrasena("c", "current-password").replace('id="c"', 'id="c" name="contrasena"')}
        <p class="campo-pie"><button type="button" class="enlace" data-olvide>¿Olvidaste tu contraseña?</button></p>
        <button class="btn primario ancho" type="submit">Entrar</button>
        <p class="cambio-modo">¿No tienes cuenta? <button type="button" class="enlace" data-ir="registro">Crear cuenta</button></p>
        ${SEPARADOR}
        <button type="button" class="btn ancho" data-codigo>${ico("qr-code")} Tengo un código de invitación</button>
        <button type="button" class="btn ancho" data-sin-cuenta data-micro="Usar sin cuenta|Trabaja sin registrarte. Tus documentos se guardan solo ${NOM.aqui().en}: si ${NOM.aqui().borrar} se pierden, salvo que los descargues.">${ico("hard-drive")} <span class="btn-sub">Usar sin cuenta<small>Tus documentos se quedan solo ${NOM.aqui().en}</small></span></button>
      </form>`);
      enlazarComunes(f);
      $("[data-olvide]", f).onclick = (e) => { e.preventDefault(); recuperar(f.nombre.value.trim()); };
      $("[data-codigo]", f).onclick = entrarConCodigo;
      f.onsubmit = (e) => {
        e.preventDefault();
        const nombre = f.nombre.value.trim(), contrasena = f.contrasena.value;
        if (!nombre) return mostrarError(f, { campo: "nombre", message: "Escribe tu usuario" });
        if (!contrasena) return mostrarError(f, { campo: "contrasena", message: "Escribe tu contraseña" });
        enviando(f, () => iniciarSesion(nombre, contrasena, false));
      };
      setTimeout(() => (f.nombre.value ? f.contrasena : f.nombre).focus(), 30);
      return f;
    }

    function formRegistro() {
      const f = html(`<form class="col" novalidate>
        ${cabecera("Crea tu cuenta", "Solo necesitas un nombre de usuario y una contraseña.")}
        <div class="alerta" data-alerta hidden role="alert"></div>
        <div class="campo"><label for="ru">Nombre de usuario</label>
          <div class="entrada">${ico("user")}<input id="ru" name="nombre" autocomplete="username" autocapitalize="none" spellcheck="false" maxlength="24" required></div>
          <div class="ayuda">De 3 a 24 letras o números, sin espacios. Así te verán los demás.</div></div>
        ${campoContrasena("rc", "new-password").replace('id="rc"', 'id="rc" name="contrasena"')}
        <div class="fila"><div class="medidor crece" data-medidor data-n="0"><i></i><i></i><i></i><i></i></div><span class="chico suave" data-fuerza style="min-width:70px;text-align:right;font-weight:600"></span></div>
        <div class="campo"><label for="re">Escuela, empresa o grupo <span class="suave" style="font-weight:600">(opcional)</span></label>
          <div class="entrada">${ico("door-open")}<input id="re" name="empresa" autocomplete="organization" maxlength="80" placeholder="Ej. tu escuela, empresa o equipo"></div></div>
        <label class="acepto"><input type="checkbox" name="acepto" required> <span>Leí y acepto los <a href="${NOM.urlAyuda("legal/terminos")}" data-ayuda="legal/terminos">Términos</a> y el <a href="${NOM.urlAyuda("legal/privacidad")}" data-ayuda="legal/privacidad">Aviso de privacidad</a> (<a href="${NOM.urlAyuda("legal/privacidad-corto")}" data-ayuda="legal/privacidad-corto">resumen</a>).</span></label>
        <button class="btn primario ancho" type="submit">Crear cuenta</button>
        <p class="suave chico" style="text-align:center">Tu contraseña se guarda protegida: nadie puede verla, ni quien administra NOM.</p>
        <p class="cambio-modo">¿Ya tienes cuenta? <button type="button" class="enlace" data-ir="entrar">Iniciar sesión</button></p>
        ${SEPARADOR}
        <button type="button" class="btn ancho" data-sin-cuenta data-micro="Usar sin cuenta|Trabaja sin registrarte. Tus documentos se guardan solo ${NOM.aqui().en}: si ${NOM.aqui().borrar} se pierden, salvo que los descargues.">${ico("hard-drive")} <span class="btn-sub">Usar sin cuenta<small>Tus documentos se quedan solo ${NOM.aqui().en}</small></span></button>
      </form>`);
      enlazarComunes(f);
      f.contrasena.addEventListener("input", () => {
        const n = fuerza(f.contrasena.value);
        $("[data-medidor]", f).dataset.n = n;
        $("[data-fuerza]", f).textContent = f.contrasena.value.length && f.contrasena.value.length < 8 ? `Faltan ${8 - f.contrasena.value.length}` : textoFuerza[n];
      });
      f.onsubmit = (e) => {
        e.preventDefault();
        const nombre = f.nombre.value.trim(), contrasena = f.contrasena.value, empresa = f.empresa.value.trim();
        if (!nombre) return mostrarError(f, { campo: "nombre", message: "Escribe un nombre de usuario" });
        if (!/^[\p{Script=Latin}0-9._-]{3,24}$/u.test(nombre) || nombre.normalize("NFKC") !== nombre) return mostrarError(f, { campo: "nombre", message: "Usa de 3 a 24 letras, números, punto, guion o guion bajo, sin espacios" });
        if (contrasena.length < 8) return mostrarError(f, { campo: "contrasena", message: "La contraseña necesita al menos 8 letras, números o signos" });
        if (!f.acepto.checked) return mostrarError(f, { message: "Para crear la cuenta, lee y acepta los Términos y el Aviso de privacidad" });
        enviando(f, async () => {
          const reg = await api.post("/api/cuentas/registro", { nombre, contrasena, empresa, acepto: NOM.LEGAL, equipo: api.equipo() }, { ms: ESPERA_BD });
          // El código de recuperación solo se ve ahora: primero se guarda, después se entra.
          if (reg && reg.recuperacion) { codigoNuevo = reg.recuperacion; await mostrarCodigo(reg.recuperacion, nombre); }
          tarjeta.innerHTML = "";
          tarjeta.append(html(`<div class="col" style="align-items:center;text-align:center;padding:24px 0">
            <div class="exito-check">${ico("check")}</div>
            <h2>¡Registro exitoso!</h2>
            <p class="suave" style="font-weight:400">Tu cuenta <b>${esc(nombre)}</b> está lista. Entrando…</p></div>`));
          // Sin volver a escribir la contraseña: se entra directo (si algo falla, al formulario de entrar).
          setTimeout(async () => {
            try { await iniciarSesion(nombre, contrasena, true); }
            catch (e) { opc.nombre = nombre; modo = "entrar"; render(); }
          }, 1200);
        });
      };
      setTimeout(() => f.nombre.focus(), 30);
      return f;
    }

    // ¿Es un teléfono que comparte? Lo último que se supo, y luego se confirma con el servidor.
    const ig = NOM.acceso && NOM.acceso.infoGuardada();
    if (ig && NOM.acceso.esTelefono(ig)) pintarTelefono();
    else render();
    if (NOM.acceso) NOM.acceso.infoServidor().then((i) => {
      if (!tarjeta.isConnected || api.token()) return;
      if (NOM.acceso.esTelefono(i)) { if (!(ig && NOM.acceso.esTelefono(ig))) pintarTelefono(); }
      else if (ig && NOM.acceso.esTelefono(ig)) render();
    }).catch(() => {});
    // En la app, al volver a la nube: si quedó guardada su sesión (al irse a un teléfono), se retoma.
    if (NOM.acceso && api.empaquetada && NOM.acceso.sesionNube()) NOM.acceso.retomarNube().then((ok) => { if (ok) NOM.app.alIniciarSesion({}); }).catch(() => {});
  }

  /** Solo en la app Android: elegir el equipo de la red que comparte los espacios. */
  function pintarServidor(tarjeta, listo, opc) {
    tarjeta.innerHTML = "";
    const f = html(`<form class="col">
      <div><h2>¿Dónde trabaja tu equipo?</h2><p class="suave" style="font-weight:400;margin-top:4px">Si te dieron un código QR, escanéalo y listo. Si no, elige una opción.</p></div>
      <div class="alerta" data-alerta hidden></div>
      <button type="button" class="btn primario ancho" data-escanear>${ico("qr-code")} Escanear el código QR</button>
      <button type="button" class="btn ancho" data-internet>${ico("cloud")} <span class="btn-sub">Usar NOM en Internet<small>Desde cualquier lugar, con datos o Wi-Fi</small></span></button>
      ${window.NomAndroid && window.NomAndroid.buscarServidores ? `<button type="button" class="btn ancho" data-buscar>${ico("radio")} <span class="btn-sub">Buscar en este Wi-Fi<small>Una computadora o un teléfono que ya comparte</small></span></button>` : ""}
      <div class="lista" data-encontrados></div>
      <details><summary class="etiqueta" style="cursor:pointer">Escribir la dirección a mano</summary>
        <div class="col" style="margin-top:10px">
          <div class="campo"><label for="srv">Dirección que te dieron</label>
            <div class="entrada">${ico("server")}<input id="srv" inputmode="url" autocapitalize="none" placeholder="192.168.1.20:8080" value="${esc(api.servidor().replace(/^https?:\/\//, ""))}"></div></div>
          <button class="btn crece" type="submit">Conectar</button>
        </div>
      </details>
      <div class="separador" aria-hidden="true">o</div>
        <button type="button" class="btn ancho" data-sin-cuenta data-micro="Usar sin cuenta|Trabaja sin registrarte. Tus documentos se guardan solo ${NOM.aqui().en}: si ${NOM.aqui().borrar} se pierden, salvo que los descargues.">${ico("hard-drive")} <span class="btn-sub">Usar sin cuenta<small>Tus documentos se quedan solo ${NOM.aqui().en}</small></span></button>
      <p class="suave chico" style="text-align:center">Sin cuenta puedes crear y editar documentos en este teléfono; nadie más los ve hasta que los subas a un espacio.</p>
    </form>`);
    f.querySelector("[data-sin-cuenta]").onclick = () => NOM.app.sinCuenta();
    // 2.1 En la app: el QR trae la dirección del servidor y el código; el espacio se abre aquí, en la app.
    f.querySelector("[data-escanear]").onclick = () => NOM.union && NOM.union.escanear();
    f.querySelector("[data-internet]").onclick = () => probar(NOM.NUBE);
    // El teléfono también puede ser el servidor de la red (Android 8+).
    const propio = NOM.anfitrion && NOM.anfitrion.opcion(listo);
    if (propio) f.querySelector("[data-alerta]").before(propio);
    tarjeta.append(f);
    const alerta = $("[data-alerta]", f);
    async function probar(dir) {
      const anterior = api.servidor();
      api.fijarServidor(dir);
      try {
        // (a un servidor que apenas se prueba no se le manda la sesión de ningún otro)
        const info = await api.get("/api/red/info", undefined, { ms: 4000, sinSesion: true });
        if (info.nom !== 1) throw new Error();
        NOM.aviso("Conectado con " + (info.equipo || dir), { tipo: "ok" });
        if (NOM.revisarVersion) NOM.revisarVersion(info);
        listo();
      } catch (e) {
        api.fijarServidor(anterior);
        alerta.className = "alerta error"; alerta.hidden = false;
        alerta.innerHTML = `${ico("wifi-off")}<span>No encontramos NOM en esa dirección. Revisa que estés en el mismo Wi-Fi que el equipo que comparte (o que tengas Internet, si es la versión en línea).</span>`;
      }
    }
    f.onsubmit = (e) => { e.preventDefault(); if (f.srv.value.trim()) probar(f.srv.value.trim()); };
    const b = $("[data-buscar]", f);
    if (b) b.onclick = () => {
      b.disabled = true; b.innerHTML = `<span class="cargando"></span> Buscando…`;
      window.__nomServidores = (json) => {
        b.disabled = false; b.innerHTML = `${ico("radio")} Buscar de nuevo`;
        const lista = JSON.parse(json || "[]");
        const cont = $("[data-encontrados]", f);
        cont.innerHTML = lista.length ? "" : `<p class="suave chico">No apareció ningún equipo. Pide al propietario que active la "Conexión del proyecto".</p>`;
        lista.forEach((s) => {
          const it = html(`<button type="button" class="lista-item" style="text-align:left">${ico("server")}<span class="crece"><b>${esc(s.equipo)}</b><small>${esc(s.host)}:${s.puerto} · ${s.espacios} espacio(s) en red</small></span>${ico("chevron-right")}</button>`);
          it.onclick = () => probar(s.host + ":" + s.puerto);
          cont.append(it);
        });
        // Primera vez: si en el Wi-Fi hay un solo servidor, se conecta solo.
        if (automatico && lista.length === 1) probar(lista[0].host + ":" + lista[0].puerto);
        automatico = false;
      };
      window.NomAndroid.buscarServidores();
    };
    // Al abrir la app se busca sin que haya que tocar nada. Solo la primera vez se conecta sola:
    // si la persona tocó "Cambiar", o ya empezó a tocar algo de esta pantalla, decide ella.
    let automatico = !!b && !(opc && opc.cambiar);
    f.addEventListener("pointerdown", () => { automatico = false; }, { capture: true });
    if (b) b.click();
  }

  /**
   * F.5 Si cambiaron los Términos o el Aviso de privacidad (o la cuenta es de antes de que se
   * pidiera aceptarlos), se piden leer y aceptar. r: la respuesta de /api/cuentas/yo o de la sesión.
   */
  let preguntandoLegal = false;
  function revisarLegal(r) {
    if (!r || !r.legal || r.aceptoVersion === r.legal || preguntandoLegal) return;
    preguntandoLegal = true;
    const nuevo = !r.aceptoVersion;
    const cuerpo = html(`<div class="col">
      <p class="suave" style="font-weight:400">${nuevo ? "Ahora NOM tiene Términos y un Aviso de privacidad." : "Actualizamos los Términos y el Aviso de privacidad."} Léelos y acéptalos para seguir usando tu cuenta.</p>
      <div class="lista">
        <a class="lista-item" href="${NOM.urlAyuda("legal/privacidad-corto")}" data-ayuda="legal/privacidad-corto">${ico("shield")}<span class="crece"><b>Aviso de privacidad (resumen)</b><small>Qué datos tratamos y para qué</small></span>${ico("chevron-right")}</a>
        <a class="lista-item" href="${NOM.urlAyuda("legal/privacidad")}" data-ayuda="legal/privacidad">${ico("shield")}<span class="crece"><b>Aviso de privacidad completo</b><small>Tus derechos y cómo ejercerlos</small></span>${ico("chevron-right")}</a>
        <a class="lista-item" href="${NOM.urlAyuda("legal/terminos")}" data-ayuda="legal/terminos">${ico("scale")}<span class="crece"><b>Términos y condiciones</b><small>Las reglas para usar NOM</small></span>${ico("chevron-right")}</a>
      </div></div>`);
    const pie = html(`<div style="display:contents"><button class="btn" data-salir>Cerrar sesión</button><span class="crece"></span><button class="btn primario" data-acepto>Leí y acepto</button></div>`);
    const h = NOM.hoja({ titulo: "Términos y privacidad", icono: "scale", cuerpo, pie, alCerrar: () => { preguntandoLegal = false; } });
    $("[data-acepto]", pie).onclick = async () => {
      try { await api.post("/api/cuentas/aceptar", { version: r.legal }); h.cerrar(true); NOM.aviso("Gracias", { tipo: "ok", ms: 1500 }); }
      catch (e) { NOM.aviso(e.message, { tipo: "error" }); }
    };
    $("[data-salir]", pie).onclick = () => { h.cerrar(true); NOM.app.cerrarSesion(); };
  }

  NOM.vistaCuentas = { pintar, mostrarCodigo, recuperar, cambiarContrasena, borrarCuenta, ofrecerRecuperar, revisarLegal };
})();

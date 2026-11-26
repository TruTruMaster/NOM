// Pruebas de llaves E2E, reloj del CRDT e importación .nom v3 (e2e.js, crdt.js, formatos.js).
// Uso: node 4-CODIGO/app/src/test/js/e2e-llaves.test.js
const fs = require("fs"), path = require("path"), nodeCrypto = require("crypto");

// Entorno mínimo de navegador para cargar los módulos del cliente sin dependencias.
global.window = { crypto: nodeCrypto.webcrypto, isSecureContext: true };
global.crypto = nodeCrypto.webcrypto;
const sesion = new Map();
global.sessionStorage = {
  get length() { return sesion.size; },
  key: (i) => Array.from(sesion.keys())[i] ?? null,
  getItem: (k) => (sesion.has(k) ? sesion.get(k) : null),
  setItem: (k, v) => { sesion.set(String(k), String(v)); },
  removeItem: (k) => { sesion.delete(String(k)); },
  clear: () => { sesion.clear(); },
};

const alm = new Map();
const puts = [];
let getApi = async (ruta) => { throw new Error("GET sin respuesta configurada: " + ruta); };

window.NOM = {
  leer: (k, def) => (alm.has(k) ? JSON.parse(alm.get(k)) : def),
  guardar: (k, v) => { alm.set(k, JSON.stringify(v)); },
  quitar: (k) => { alm.delete(k); },
  throttle: (fn) => fn,
  sinContrasena: () => false,
  aviso: () => {},
  esc: (s) => String(s || ""),
  api: {
    clave: (k) => "u1." + k,
    token: () => "tok-prueba",
    usuario: () => ({ id: "u1", nombre: "victima" }),
    dispositivo: "disp-prueba",
    get: async (ruta, params, opc) => getApi(ruta, params, opc),
    put: async (ruta, cuerpo, opc) => { puts.push({ ruta, cuerpo }); return {}; },
    post: async () => ({ aplicadas: 0 }),
    ErrorApi: class extends Error { constructor(estado) { super("API " + estado); this.estado = estado; } },
  },
};

const dirJs = process.env.NOM_JS_DIR || path.join(__dirname, "../../main/resources/static/js");
for (const archivo of ["cripto.js", "secuencia.js", "crdt.js", "e2e.js", "compresion.js", "origen.js", "formatos.js"]) {
  eval(fs.readFileSync(path.join(dirJs, archivo), "utf8"));
}

const NOM = window.NOM, C = NOM.cripto, E = NOM.e2e;
let fallas = 0;
const prueba = (nombre, ok) => { console.log((ok ? "OK    " : "FALLA ") + nombre); if (!ok) fallas++; };
const iguales = (a, b) => a instanceof Uint8Array && b instanceof Uint8Array && a.length === b.length && a.every((x, i) => x === b[i]);
const unir = (...ps) => { const r = new Uint8Array(ps.reduce((n, p) => n + p.length, 0)); let o = 0; ps.forEach((p) => { r.set(p, o); o += p.length; }); return r; };
const esperar = () => new Promise((r) => setTimeout(r, 15));

async function cifrarBytes(llave, datos) {
  const s = await C.cifrar(llave, datos, C.algPreferido());
  return "e2e:" + E._b64(unir(new Uint8Array([s.alg === "gcm" ? 1 : 2]), s.iv, s.ct));
}
async function descifrarBytes(llave, texto) {
  const b = E._deB64(texto.slice(4));
  const alg = b[0] === 1 ? "gcm" : "ctr", n = alg === "gcm" ? 12 : 16;
  return C.descifrar(llave, { alg, iv: b.subarray(1, 1 + n), ct: b.subarray(1 + n) });
}
const cifrarValor = (llave, v) => cifrarBytes(llave, C.bytes(JSON.stringify(v)));

async function derivarCon(privKey, pubJwk) {
  const pub = await crypto.subtle.importKey("jwk", pubJwk, { name: "ECDH", namedCurve: "P-256" }, false, []);
  const bits = new Uint8Array(await crypto.subtle.deriveBits({ name: "ECDH", public: pub }, privKey, 256));
  return C.sha256(unir(bits, C.bytes("nom-colab/e2e/envoltura")));
}
async function envolverPara(pubTexto, llave) {
  const eph = await crypto.subtle.generateKey({ name: "ECDH", namedCurve: "P-256" }, false, ["deriveBits"]);
  const epk = await crypto.subtle.exportKey("jwk", eph.publicKey);
  const kw = await derivarCon(eph.privateKey, JSON.parse(pubTexto));
  return JSON.stringify({ epk, s: await cifrarBytes(kw, llave) });
}
async function abrirEnvueltaCon(privTexto, texto) {
  const o = JSON.parse(texto);
  const priv = await crypto.subtle.importKey("jwk", JSON.parse(privTexto), { name: "ECDH", namedCurve: "P-256" }, false, ["deriveBits"]);
  const kw = await derivarCon(priv, o.epk);
  return descifrarBytes(kw, o.s);
}

(async () => {
  const par = await crypto.subtle.generateKey({ name: "ECDH", namedCurve: "P-256" }, true, ["deriveBits"]);
  const pub = JSON.stringify(await crypto.subtle.exportKey("jwk", par.publicKey));
  const priv = JSON.stringify(await crypto.subtle.exportKey("jwk", par.privateKey));

  const reiniciar = () => {
    E.limpiar();
    alm.clear();
    sesion.clear();
    puts.length = 0;
    NOM.guardar(NOM.api.clave("e2e.publica"), pub);
    NOM.guardar(NOM.api.clave("e2e.privada"), priv);
  };

  const kReal = C.azar(32), kFalsa = C.azar(32);
  const nombreCifrado = await cifrarValor(kReal, "Espacio real");
  const envReal = await envolverPara(pub, kReal);
  const envFalsa = await envolverPara(pub, kFalsa);

  // 1. Sustitución de llave (4_SUSTITUCION): K_falsa no abre el nombre cifrado con K_real.
  reiniciar();
  let miaServidor = envFalsa;
  getApi = async (ruta) => {
    if (ruta === "/api/espacios/esp1/llaves") return { e2e: true, mia: miaServidor, nombre: nombreCifrado, faltan: [] };
    throw new Error("Ruta inesperada: " + ruta);
  };
  const r1 = await E.llaveEspacio("esp1");
  await esperar();
  const borroPropia = puts.some((p) => p.ruta === "/api/espacios/esp1/llaves" && !p.cuerpo.usuario && p.cuerpo.llave === "");
  miaServidor = null;
  const r1Despues = await E.llaveEspacio("esp1");
  prueba("1. Sustitución: llaveEspacio devuelve null si la llave envuelta no abre el nombre", r1 === null);
  prueba("1. Sustitución: la llave falsa no queda guardada en bóveda ni memoria", NOM.leer(NOM.api.clave("e2e.esp.esp1"), null) === null && r1Despues === null);
  prueba("1. Sustitución: se manda PUT /llaves con llave vacía para quitar la propia", borroPropia);

  // 2. Llave buena: abre el nombre, se usa y se guarda.
  reiniciar();
  getApi = async (ruta) => {
    if (ruta === "/api/espacios/esp2/llaves") return { e2e: true, mia: envReal, nombre: nombreCifrado, faltan: [] };
    throw new Error("Ruta inesperada: " + ruta);
  };
  const r2 = await E.llaveEspacio("esp2");
  await esperar();
  const nombreAbierto = await E.nombreEspacio({ id: "esp2", e2e: true, nombre: nombreCifrado });
  prueba("2. Llave buena: llaveEspacio devuelve K_real", iguales(r2, kReal));
  prueba("2. Llave buena: queda guardada y descifra el nombre sin borrar en el servidor",
    NOM.leer(NOM.api.clave("e2e.esp.esp2"), null) === E._b64(kReal) && nombreAbierto === "Espacio real" && puts.length === 0);

  // 3. alEntrar: rechaza #k= falso y guarda #k= bueno.
  reiniciar();
  getApi = async (ruta) => {
    if (ruta === "/api/espacios/esp3/llaves") return { e2e: true, mia: null, nombre: nombreCifrado, faltan: [] };
    throw new Error("Ruta inesperada: " + ruta);
  };
  E.recordarInvitacion("INV-FALSA", E._b64(kFalsa));
  await E.alEntrar("INV-FALSA", "esp3");
  await esperar();
  prueba("3. alEntrar con #k= falso: no se guarda ni se sube al servidor",
    NOM.leer(NOM.api.clave("e2e.esp.esp3"), null) === null && (await E.llaveEspacio("esp3")) === null && puts.length === 0);

  reiniciar();
  getApi = async (ruta) => {
    if (ruta === "/api/espacios/esp3b/llaves") return { e2e: true, mia: null, nombre: nombreCifrado, faltan: [] };
    throw new Error("Ruta inesperada: " + ruta);
  };
  E.recordarInvitacion("INV-BUENA", E._b64(kReal));
  await E.alEntrar("INV-BUENA", "esp3b");
  await esperar();
  prueba("3. alEntrar con #k= bueno: se guarda y queda disponible en llaveEspacio",
    NOM.leer(NOM.api.clave("e2e.esp.esp3b"), null) === E._b64(kReal) && iguales(await E.llaveEspacio("esp3b"), kReal));

  // 4. repartir: si l.mia abre otra llave distinta de la que tengo, manda PUT con la mía.
  reiniciar();
  NOM.guardar(NOM.api.clave("e2e.esp.esp4"), E._b64(kReal));
  getApi = async (ruta) => {
    if (ruta === "/api/espacios/esp4/llaves") return { e2e: true, mia: envFalsa, nombre: nombreCifrado, faltan: [] };
    throw new Error("Ruta inesperada: " + ruta);
  };
  await E.repartir("esp4");
  const putMia = puts.find((p) => p.ruta === "/api/espacios/esp4/llaves" && !p.cuerpo.usuario && p.cuerpo.llave);
  const llaveSubida = putMia ? await abrirEnvueltaCon(priv, putMia.cuerpo.llave) : null;
  prueba("4. repartir: si l.mia abre otra llave distinta de la mía, manda PUT con la mía", !!putMia && iguales(llaveSubida, kReal));

  // 5. Reloj reiniciado / reuso (B_RELOJ_REINICIADO / 1_REUSO): reemplazarEstado no retrocede el reloj.
  reiniciar();
  const doc = new NOM.DocCRDT("doc-reloj", { bloques: {}, seq: {}, reloj: 0 }, { soloMemoria: true });
  let e1, e2;
  try {
    e1 = doc.seq("cuerpo", { k: "i", iz: null, x: "SECRETO" });
    doc.descartarCola();
    doc.reemplazarEstado({ bloques: {}, seq: {}, reloj: 0 });
    e2 = doc.seq("cuerpo", { k: "i", iz: null, x: "OTROTXT" });
  } finally {
    doc.cerrar();
  }
  const maxUsado = e1.t + NOM.Secuencia.caracteres(e1.x).length - 1;
  const kDoc = C.azar(32);
  const ct1 = E._letras(kDoc, e1.n, e1.t, "AAAAAAA", 1);
  const ct2 = E._letras(kDoc, e2.n, e2.t, "AAAAAAA", 1);
  prueba("5. reemplazarEstado con reloj menor: la siguiente letra usa t mayor que todos los ya usados",
    e2.t > maxUsado && ct1 !== ct2);

  // 6. Importar .nom v3 con dos bloques de texto: ningún par (n, t) de letras se repite entre bloques.
  const bytesNom = NOM.paquete.empacar({
    f: "nom-colab",
    v: 3,
    d: { nombre: "Dos textos", modo: "TEXTO" },
    b: {},
    t: {
      cuerpo: [[0, ["SECRETO BLOQUE UNO", 0]], [0, ["Segunda linea", 0]]],
      notas: [[0, ["SECRETO BLOQUE DOS", 0]]],
    },
  }, []);
  const archivoNom = {
    name: "dos-bloques.nom",
    size: bytesNom.length,
    arrayBuffer: async () => bytesNom.buffer.slice(bytesNom.byteOffset, bytesNom.byteOffset + bytesNom.byteLength),
  };
  const imp = await NOM.formatos.importar(archivoNom);
  const usados = new Set(), bloques = new Set();
  let repetidos = 0;
  for (const op of imp.ops) {
    if (!op.s || op.s.k !== "i") continue;
    bloques.add(op.bloque);
    const largo = NOM.Secuencia.caracteres(op.s.x).length;
    for (let i = 0; i < largo; i++) {
      const id = op.n + "@" + (op.t + i);
      if (usados.has(id)) repetidos++;
      usados.add(id);
    }
  }
  prueba("6. Importar .nom v3 con dos bloques de texto: ningún par (n, t) se repite entre bloques",
    bloques.size === 2 && usados.size > 0 && repetidos === 0);

  console.log(fallas ? `${fallas} FALLA(S)` : "Todas las pruebas pasaron");
  process.exit(fallas ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });

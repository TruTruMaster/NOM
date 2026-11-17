/* NOM · Archivos de Office con contraseña (MS-OFFCRYPTO).
   Office guarda un .docx, .xlsx o .pptx con contraseña dentro de un archivo compuesto (CFB, el formato OLE de
   siempre) con dos partes: «EncryptionInfo» (cómo se cifró) y «EncryptedPackage» (el .zip cifrado). Aquí se abre
   con la contraseña (cifrado «Agile» de Office 2010 en adelante y «Standard» de Office 2007) y se guarda con
   contraseña como lo hace Office 365 (Agile: AES-256 en modo CBC, SHA-512 con 100 000 vueltas, HMAC de integridad).
   Todo en JavaScript puro: en la red local por http:// el navegador no ofrece WebCrypto. Verificado con los
   vectores de FIPS-180 (SHA-1, SHA-512), FIPS-197 (AES), RFC 4231 (HMAC) y con Office real (pruebas/office). */
(function (raiz) {
  "use strict";
  const NOM = (raiz.NOM = raiz.NOM || {});

  // ------------------------------------------------------------ SHA-1
  function sha1(m) {
    const ml = m.length, nb = ((ml + 9 + 63) >> 6) << 6, b = new Uint8Array(nb);
    b.set(m); b[ml] = 0x80;
    const bits = ml * 8, dv = new DataView(b.buffer);
    dv.setUint32(nb - 4, bits >>> 0); dv.setUint32(nb - 8, Math.floor(bits / 4294967296));
    let h0 = 0x67452301, h1 = 0xefcdab89, h2 = 0x98badcfe, h3 = 0x10325476, h4 = 0xc3d2e1f0;
    const w = new Uint32Array(80);
    for (let o = 0; o < nb; o += 64) {
      for (let i = 0; i < 16; i++) w[i] = dv.getUint32(o + i * 4);
      for (let i = 16; i < 80; i++) { const x = w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16]; w[i] = (x << 1) | (x >>> 31); }
      let a = h0, bb = h1, c = h2, d = h3, e = h4;
      for (let i = 0; i < 80; i++) {
        const f = i < 20 ? (bb & c) | (~bb & d) : i < 40 ? bb ^ c ^ d : i < 60 ? (bb & c) | (bb & d) | (c & d) : bb ^ c ^ d;
        const k = i < 20 ? 0x5a827999 : i < 40 ? 0x6ed9eba1 : i < 60 ? 0x8f1bbcdc : 0xca62c1d6;
        const t = (((a << 5) | (a >>> 27)) + f + e + k + w[i]) | 0;
        e = d; d = c; c = (bb << 30) | (bb >>> 2); bb = a; a = t;
      }
      h0 = (h0 + a) | 0; h1 = (h1 + bb) | 0; h2 = (h2 + c) | 0; h3 = (h3 + d) | 0; h4 = (h4 + e) | 0;
    }
    const r = new Uint8Array(20), rv = new DataView(r.buffer);
    [h0, h1, h2, h3, h4].forEach((x, i) => rv.setUint32(i * 4, x >>> 0));
    return r;
  }

  // ------------------------------------------------------------ SHA-512 (con pares de enteros de 32 bits)
  const K512 = [
    0x428a2f98, 0xd728ae22, 0x71374491, 0x23ef65cd, 0xb5c0fbcf, 0xec4d3b2f, 0xe9b5dba5, 0x8189dbbc, 0x3956c25b, 0xf348b538, 0x59f111f1, 0xb605d019,
    0x923f82a4, 0xaf194f9b, 0xab1c5ed5, 0xda6d8118, 0xd807aa98, 0xa3030242, 0x12835b01, 0x45706fbe, 0x243185be, 0x4ee4b28c, 0x550c7dc3, 0xd5ffb4e2,
    0x72be5d74, 0xf27b896f, 0x80deb1fe, 0x3b1696b1, 0x9bdc06a7, 0x25c71235, 0xc19bf174, 0xcf692694, 0xe49b69c1, 0x9ef14ad2, 0xefbe4786, 0x384f25e3,
    0x0fc19dc6, 0x8b8cd5b5, 0x240ca1cc, 0x77ac9c65, 0x2de92c6f, 0x592b0275, 0x4a7484aa, 0x6ea6e483, 0x5cb0a9dc, 0xbd41fbd4, 0x76f988da, 0x831153b5,
    0x983e5152, 0xee66dfab, 0xa831c66d, 0x2db43210, 0xb00327c8, 0x98fb213f, 0xbf597fc7, 0xbeef0ee4, 0xc6e00bf3, 0x3da88fc2, 0xd5a79147, 0x930aa725,
    0x06ca6351, 0xe003826f, 0x14292967, 0x0a0e6e70, 0x27b70a85, 0x46d22ffc, 0x2e1b2138, 0x5c26c926, 0x4d2c6dfc, 0x5ac42aed, 0x53380d13, 0x9d95b3df,
    0x650a7354, 0x8baf63de, 0x766a0abb, 0x3c77b2a8, 0x81c2c92e, 0x47edaee6, 0x92722c85, 0x1482353b, 0xa2bfe8a1, 0x4cf10364, 0xa81a664b, 0xbc423001,
    0xc24b8b70, 0xd0f89791, 0xc76c51a3, 0x0654be30, 0xd192e819, 0xd6ef5218, 0xd6990624, 0x5565a910, 0xf40e3585, 0x5771202a, 0x106aa070, 0x32bbd1b8,
    0x19a4c116, 0xb8d2d0c8, 0x1e376c08, 0x5141ab53, 0x2748774c, 0xdf8eeb99, 0x34b0bcb5, 0xe19b48a8, 0x391c0cb3, 0xc5c95a63, 0x4ed8aa4a, 0xe3418acb,
    0x5b9cca4f, 0x7763e373, 0x682e6ff3, 0xd6b2b8a3, 0x748f82ee, 0x5defb2fc, 0x78a5636f, 0x43172f60, 0x84c87814, 0xa1f0ab72, 0x8cc70208, 0x1a6439ec,
    0x90befffa, 0x23631e28, 0xa4506ceb, 0xde82bde9, 0xbef9a3f7, 0xb2c67915, 0xc67178f2, 0xe372532b, 0xca273ece, 0xea26619c, 0xd186b8c7, 0x21c0c207,
    0xeada7dd6, 0xcde0eb1e, 0xf57d4f7f, 0xee6ed178, 0x06f067aa, 0x72176fba, 0x0a637dc5, 0xa2c898a6, 0x113f9804, 0xbef90dae, 0x1b710b35, 0x131c471b,
    0x28db77f5, 0x23047d84, 0x32caab7b, 0x40c72493, 0x3c9ebe0a, 0x15c9bebc, 0x431d67c4, 0x9c100d4c, 0x4cc5d4be, 0xcb3e42b6, 0x597f299c, 0xfc657e2a,
    0x5fcb6fab, 0x3ad6faec, 0x6c44198c, 0x4a475817];
  const W512 = new Int32Array(160);
  function sha512(m) {
    const ml = m.length, nb = ((ml + 17 + 127) >> 7) << 7, b = new Uint8Array(nb);
    b.set(m); b[ml] = 0x80;
    const dv = new DataView(b.buffer), bits = ml * 8;
    dv.setUint32(nb - 4, bits >>> 0); dv.setUint32(nb - 8, Math.floor(bits / 4294967296));
    const H = new Int32Array([0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1,
      0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179]);
    const w = W512;
    for (let o = 0; o < nb; o += 128) {
      for (let i = 0; i < 32; i++) w[i] = dv.getInt32(o + i * 4);
      for (let i = 16; i < 80; i++) {
        // s0 = rotr(x,1) ^ rotr(x,8) ^ shr(x,7) ; s1 = rotr(x,19) ^ rotr(x,61) ^ shr(x,6)
        const xh = w[(i - 15) * 2], xl = w[(i - 15) * 2 + 1];
        const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
        const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
        const yh = w[(i - 2) * 2], yl = w[(i - 2) * 2 + 1];
        const s1h = ((yh >>> 19) | (yl << 13)) ^ ((yl >>> 29) | (yh << 3)) ^ (yh >>> 6);
        const s1l = ((yl >>> 19) | (yh << 13)) ^ ((yh >>> 29) | (yl << 3)) ^ ((yl >>> 6) | (yh << 26));
        let l = (s0l >>> 0) + (s1l >>> 0) + (w[(i - 16) * 2 + 1] >>> 0) + (w[(i - 7) * 2 + 1] >>> 0);
        const h = (s0h + s1h + w[(i - 16) * 2] + w[(i - 7) * 2] + Math.floor(l / 4294967296)) | 0;
        w[i * 2] = h; w[i * 2 + 1] = l | 0;
      }
      let ah = H[0], al = H[1], bh = H[2], bl = H[3], ch = H[4], cl = H[5], dh = H[6], dl = H[7];
      let eh = H[8], el = H[9], fh = H[10], fl = H[11], gh = H[12], gl = H[13], hh = H[14], hl = H[15];
      for (let i = 0; i < 80; i++) {
        // S1 = rotr(e,14) ^ rotr(e,18) ^ rotr(e,41); ch = (e&f) ^ (~e&g)
        const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
        const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
        const chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
        let l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K512[i * 2 + 1] >>> 0) + (w[i * 2 + 1] >>> 0);
        const t1h = (hh + S1h + chh + K512[i * 2] + w[i * 2] + Math.floor(l / 4294967296)) | 0, t1l = l | 0;
        // S0 = rotr(a,28) ^ rotr(a,34) ^ rotr(a,39); maj = (a&b) ^ (a&c) ^ (b&c)
        const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
        const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
        const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch), mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
        l = (S0l >>> 0) + (mjl >>> 0);
        const t2h = (S0h + mjh + Math.floor(l / 4294967296)) | 0, t2l = l | 0;
        hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
        l = (dl >>> 0) + (t1l >>> 0); eh = (dh + t1h + Math.floor(l / 4294967296)) | 0; el = l | 0;
        dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
        l = (t1l >>> 0) + (t2l >>> 0); ah = (t1h + t2h + Math.floor(l / 4294967296)) | 0; al = l | 0;
      }
      const suma = (k, xh, xl) => { const l = (H[k + 1] >>> 0) + (xl >>> 0); H[k] = (H[k] + xh + Math.floor(l / 4294967296)) | 0; H[k + 1] = l | 0; };
      suma(0, ah, al); suma(2, bh, bl); suma(4, ch, cl); suma(6, dh, dl); suma(8, eh, el); suma(10, fh, fl); suma(12, gh, gl); suma(14, hh, hl);
    }
    const r = new Uint8Array(64), rv = new DataView(r.buffer);
    for (let i = 0; i < 16; i++) rv.setInt32(i * 4, H[i]);
    return r;
  }
  const HASHES = { SHA1: { f: sha1, n: 20, bloque: 64 }, SHA512: { f: sha512, n: 64, bloque: 128 } };
  function hmac(alg, clave, datos) {
    const h = HASHES[alg];
    let k = clave.length > h.bloque ? h.f(clave) : clave;
    const kp = new Uint8Array(h.bloque); kp.set(k);
    const ip = new Uint8Array(h.bloque + datos.length), op = new Uint8Array(h.bloque + h.n);
    for (let i = 0; i < h.bloque; i++) { ip[i] = kp[i] ^ 0x36; op[i] = kp[i] ^ 0x5c; }
    ip.set(datos, h.bloque);
    op.set(h.f(ip), h.bloque);
    return h.f(op);
  }

  // ------------------------------------------------------------ AES (128, 192 y 256), cifrar y descifrar un bloque
  const SB = new Uint8Array(256), ISB = new Uint8Array(256);
  (function () {
    const p = new Uint8Array(256), l = new Uint8Array(256);
    let x = 1;
    for (let i = 0; i < 255; i++) { p[i] = x; l[x] = i; x ^= (x << 1) ^ (x & 0x80 ? 0x1b : 0); x &= 0xff; }
    SB[0] = 0x63;
    for (let i = 1; i < 256; i++) {
      let s = p[(255 - l[i]) % 255];   // (el inverso de 1 es 1: p[0])
      s = s ^ ((s << 1) | (s >> 7)) ^ ((s << 2) | (s >> 6)) ^ ((s << 3) | (s >> 5)) ^ ((s << 4) | (s >> 4));
      SB[i] = (s ^ 0x63) & 0xff;
    }
    for (let i = 0; i < 256; i++) ISB[SB[i]] = i;
  })();
  const xt = (a) => ((a << 1) ^ (a & 0x80 ? 0x1b : 0)) & 0xff;
  const mul = (a, b) => { let r = 0; while (b) { if (b & 1) r ^= a; a = xt(a); b >>= 1; } return r; };
  function expandir(clave) {
    const nk = clave.length / 4, nr = nk + 6, w = new Uint8Array(16 * (nr + 1));
    w.set(clave);
    let rcon = 1;
    for (let i = nk; i < 4 * (nr + 1); i++) {
      let t = w.slice((i - 1) * 4, i * 4);
      if (i % nk === 0) { t = new Uint8Array([SB[t[1]] ^ rcon, SB[t[2]], SB[t[3]], SB[t[0]]]); rcon = xt(rcon); }
      else if (nk > 6 && i % nk === 4) t = t.map((v) => SB[v]);
      for (let j = 0; j < 4; j++) w[i * 4 + j] = w[(i - nk) * 4 + j] ^ t[j];
    }
    return { w, nr };
  }
  function cifrarBloque(k, b) {
    const s = b.slice(0, 16), { w, nr } = k;
    for (let i = 0; i < 16; i++) s[i] ^= w[i];
    for (let r = 1; r <= nr; r++) {
      for (let i = 0; i < 16; i++) s[i] = SB[s[i]];
      // filas
      let t = s[1]; s[1] = s[5]; s[5] = s[9]; s[9] = s[13]; s[13] = t;
      t = s[2]; s[2] = s[10]; s[10] = t; t = s[6]; s[6] = s[14]; s[14] = t;
      t = s[15]; s[15] = s[11]; s[11] = s[7]; s[7] = s[3]; s[3] = t;
      if (r < nr) for (let c = 0; c < 4; c++) {
        const a0 = s[c * 4], a1 = s[c * 4 + 1], a2 = s[c * 4 + 2], a3 = s[c * 4 + 3];
        s[c * 4] = xt(a0) ^ xt(a1) ^ a1 ^ a2 ^ a3;
        s[c * 4 + 1] = a0 ^ xt(a1) ^ xt(a2) ^ a2 ^ a3;
        s[c * 4 + 2] = a0 ^ a1 ^ xt(a2) ^ xt(a3) ^ a3;
        s[c * 4 + 3] = xt(a0) ^ a0 ^ a1 ^ a2 ^ xt(a3);
      }
      for (let i = 0; i < 16; i++) s[i] ^= w[r * 16 + i];
    }
    return s;
  }
  function descifrarBloque(k, b) {
    const s = b.slice(0, 16), { w, nr } = k;
    for (let i = 0; i < 16; i++) s[i] ^= w[nr * 16 + i];
    for (let r = nr - 1; r >= 0; r--) {
      let t = s[13]; s[13] = s[9]; s[9] = s[5]; s[5] = s[1]; s[1] = t;
      t = s[2]; s[2] = s[10]; s[10] = t; t = s[6]; s[6] = s[14]; s[14] = t;
      t = s[3]; s[3] = s[7]; s[7] = s[11]; s[11] = s[15]; s[15] = t;
      for (let i = 0; i < 16; i++) s[i] = ISB[s[i]];
      for (let i = 0; i < 16; i++) s[i] ^= w[r * 16 + i];
      if (r > 0) for (let c = 0; c < 4; c++) {
        const a0 = s[c * 4], a1 = s[c * 4 + 1], a2 = s[c * 4 + 2], a3 = s[c * 4 + 3];
        s[c * 4] = mul(a0, 14) ^ mul(a1, 11) ^ mul(a2, 13) ^ mul(a3, 9);
        s[c * 4 + 1] = mul(a0, 9) ^ mul(a1, 14) ^ mul(a2, 11) ^ mul(a3, 13);
        s[c * 4 + 2] = mul(a0, 13) ^ mul(a1, 9) ^ mul(a2, 14) ^ mul(a3, 11);
        s[c * 4 + 3] = mul(a0, 11) ^ mul(a1, 13) ^ mul(a2, 9) ^ mul(a3, 14);
      }
    }
    return s;
  }
  /** AES en modo CBC sin relleno (los datos ya vienen en múltiplos de 16, como en MS-OFFCRYPTO). */
  function cbc(clave, iv, datos, cifrar) {
    const k = expandir(clave), n = Math.ceil(datos.length / 16) * 16, out = new Uint8Array(n);
    const entrada = new Uint8Array(n); entrada.set(datos);
    let prev = iv.slice(0, 16);
    for (let o = 0; o < n; o += 16) {
      const b = entrada.subarray(o, o + 16);
      if (cifrar) { const x = new Uint8Array(16); for (let i = 0; i < 16; i++) x[i] = b[i] ^ prev[i]; prev = cifrarBloque(k, x); out.set(prev, o); }
      else { const d = descifrarBloque(k, b); for (let i = 0; i < 16; i++) out[o + i] = d[i] ^ prev[i]; prev = b.slice(); }
    }
    return out;
  }
  function ecbDescifrar(clave, datos) {
    const k = expandir(clave), out = new Uint8Array(datos.length);
    for (let o = 0; o + 16 <= datos.length; o += 16) out.set(descifrarBloque(k, datos.subarray(o, o + 16)), o);
    return out;
  }

  // ------------------------------------------------------------ archivo compuesto (CFB): leer y escribir
  const FIRMA = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  const esCFB = (u) => u && u.length >= 512 && FIRMA.every((b, i) => u[i] === b);
  /** Lee un archivo compuesto: { "Nombre": bytes, "Carpeta/Nombre": bytes, … } con todas sus corrientes. */
  function leerCFB(u) {
    if (!esCFB(u)) throw new Error("No es un archivo compuesto");
    const dv = new DataView(u.buffer, u.byteOffset, u.byteLength);
    const tamS = 1 << dv.getUint16(30, true), tamMini = 1 << dv.getUint16(32, true);
    const nFat = dv.getUint32(44, true), dir0 = dv.getUint32(48, true), corte = dv.getUint32(56, true);
    const miniFat0 = dv.getUint32(60, true), difat0 = dv.getUint32(68, true);
    const sector = (i) => 512 + i * tamS;   // (versión 3: 512; versión 4: el encabezado ocupa un sector de 4096)
    const off = (i) => (tamS === 512 ? sector(i) : tamS * (i + 1));
    const FIN = 0xfffffffe;
    // tabla de sectores (FAT): de la DIFAT del encabezado y de sus sectores extra
    const sectoresFat = [];
    for (let i = 0; i < 109 && sectoresFat.length < nFat; i++) sectoresFat.push(dv.getUint32(76 + i * 4, true));
    for (let s = difat0, vueltas = 0; s < 0xfffffffa && sectoresFat.length < nFat && vueltas < 10000; vueltas++) {
      const o = off(s), n = tamS / 4;
      for (let i = 0; i < n - 1 && sectoresFat.length < nFat; i++) sectoresFat.push(dv.getUint32(o + i * 4, true));
      s = dv.getUint32(o + (n - 1) * 4, true);
    }
    const fat = [];
    for (const s of sectoresFat) { const o = off(s); for (let i = 0; i < tamS / 4; i++) fat.push(dv.getUint32(o + i * 4, true)); }
    const cadena = (s0, tabla) => { const r = []; for (let s = s0, v = 0; s < 0xfffffffa && v < 1e7; v++) { r.push(s); s = tabla[s]; if (s === FIN) break; } return r; };
    const leer = (s0, tam) => { const sec = cadena(s0, fat), r = new Uint8Array(sec.length * tamS); sec.forEach((s, i) => r.set(u.subarray(off(s), off(s) + tamS), i * tamS)); return tam != null ? r.subarray(0, tam) : r; };
    // directorio
    const dir = leer(dir0), dd = new DataView(dir.buffer, dir.byteOffset, dir.byteLength), ent = [];
    for (let o = 0; o + 128 <= dir.length; o += 128) {
      const ln = dd.getUint16(o + 64, true), tipo = dir[o + 66];
      const nombre = ln >= 2 ? String.fromCharCode(...new Uint16Array(dir.slice(o, o + ln - 2).buffer)) : "";
      ent.push({ nombre, tipo, izq: dd.getUint32(o + 68, true), der: dd.getUint32(o + 72, true), hijo: dd.getUint32(o + 76, true), ini: dd.getUint32(o + 116, true), tam: dd.getUint32(o + 120, true) });
    }
    const raizE = ent[0], miniStream = raizE && raizE.ini < 0xfffffffa ? leer(raizE.ini, raizE.tam) : new Uint8Array(0);
    const miniFat = [];
    if (miniFat0 < 0xfffffffa) { const m = leer(miniFat0), mv = new DataView(m.buffer, m.byteOffset, m.byteLength); for (let i = 0; i < m.length / 4; i++) miniFat.push(mv.getUint32(i * 4, true)); }
    const leerMini = (s0, tam) => { const sec = cadena(s0, miniFat), r = new Uint8Array(sec.length * tamMini); sec.forEach((s, i) => r.set(miniStream.subarray(s * tamMini, (s + 1) * tamMini), i * tamMini)); return r.subarray(0, tam); };
    const out = {};
    const recorrer = (id, ruta, visto) => {
      if (id >= ent.length || id === 0xffffffff || visto.has(id)) return;
      visto.add(id);
      const e = ent[id];
      recorrer(e.izq, ruta, visto); recorrer(e.der, ruta, visto);
      const r2 = ruta ? ruta + "/" + e.nombre : e.nombre;
      if (e.tipo === 2) out[r2] = e.tam < corte ? leerMini(e.ini, e.tam) : leer(e.ini, e.tam);
      else if (e.tipo === 1) recorrer(e.hijo, r2, visto);
    };
    if (raizE) recorrer(raizE.hijo, "", new Set());
    return out;
  }
  /** Escribe un archivo compuesto (versión 3, sectores de 512) con esas corrientes ({ "Carpeta/Nombre": bytes }). */
  function escribirCFB(corrientes) {
    const S = 512, MS = 64, CORTE = 4096, LIBRE = 0xffffffff, FIN = 0xfffffffe, FATSEC = 0xfffffffd;
    // árbol de nombres
    const raizN = { nombre: "Root Entry", tipo: 5, hijos: {} };
    for (const ruta of Object.keys(corrientes)) {
      const partes = ruta.split("/");
      let n = raizN;
      partes.forEach((p, i) => {
        if (i === partes.length - 1) n.hijos[p] = { nombre: p, tipo: 2, datos: corrientes[ruta] };
        else n = n.hijos[p] = n.hijos[p] || { nombre: p, tipo: 1, hijos: {} };
      });
    }
    const nodos = [raizN];
    const cmp = (a, b) => (a.nombre.length !== b.nombre.length ? a.nombre.length - b.nombre.length : a.nombre.toUpperCase() < b.nombre.toUpperCase() ? -1 : a.nombre.toUpperCase() > b.nombre.toUpperCase() ? 1 : 0);
    // los hijos de cada carpeta en un árbol binario equilibrado (todos negros), ordenados como pide el formato
    const arbol = (lista) => {
      if (!lista.length) return LIBRE;
      const m = lista.length >> 1, n = lista[m];
      n.id = nodos.length; nodos.push(n);
      n.izq = arbol(lista.slice(0, m)); n.der = arbol(lista.slice(m + 1));
      if (n.tipo !== 2) n.hijo = arbol(Object.values(n.hijos).sort(cmp));
      return n.id;
    };
    raizN.id = 0; raizN.izq = raizN.der = LIBRE;
    raizN.hijo = arbol(Object.values(raizN.hijos).sort(cmp));
    // mini corriente (lo que pesa menos que el corte) y sectores normales
    const mini = [], grandes = [];
    nodos.forEach((n) => { if (n.tipo === 2) (n.datos.length < CORTE ? mini : grandes).push(n); });
    const miniFat = [], partesMini = [];
    let mOff = 0;
    for (const n of mini) {
      const k = Math.ceil(n.datos.length / MS);
      n.ini = k ? mOff : FIN;
      for (let i = 0; i < k; i++) miniFat.push(i === k - 1 ? FIN : mOff + i + 1);
      const b = new Uint8Array(k * MS); b.set(n.datos); partesMini.push(b); mOff += k;
    }
    const miniStream = new Uint8Array(mOff * MS);
    { let o = 0; for (const b of partesMini) { miniStream.set(b, o); o += b.length; } }
    // sectores: [corrientes grandes] [mini corriente] [mini FAT] [directorio] [FAT]
    const sectores = [], fat = [];
    const poner = (datos) => {
      const k = Math.max(1, Math.ceil(datos.length / S)), ini = sectores.length;
      for (let i = 0; i < k; i++) { const b = new Uint8Array(S); b.set(datos.subarray(i * S, (i + 1) * S)); sectores.push(b); fat.push(i === k - 1 ? FIN : ini + i + 1); }
      return ini;
    };
    for (const n of grandes) n.ini = poner(n.datos);
    raizN.ini = mOff ? poner(miniStream) : FIN; raizN.datos = { length: miniStream.length };
    let miniFat0 = FIN, nMiniFat = 0;
    if (miniFat.length) {
      const b = new Uint8Array(Math.ceil(miniFat.length / 128) * S).fill(0xff), v = new DataView(b.buffer);
      miniFat.forEach((x, i) => v.setUint32(i * 4, x, true));
      miniFat0 = poner(b); nMiniFat = b.length / S;
    }
    const dirB = new Uint8Array(Math.ceil(nodos.length / 4) * 4 * 128), dv = new DataView(dirB.buffer);
    for (let i = nodos.length; i < dirB.length / 128; i++) { dv.setUint32(i * 128 + 68, LIBRE, true); dv.setUint32(i * 128 + 72, LIBRE, true); dv.setUint32(i * 128 + 76, LIBRE, true); }
    nodos.forEach((n, i) => {
      const o = i * 128;
      for (let k = 0; k < n.nombre.length && k < 31; k++) dv.setUint16(o + k * 2, n.nombre.charCodeAt(k), true);
      dv.setUint16(o + 64, (Math.min(31, n.nombre.length) + 1) * 2, true);
      dirB[o + 66] = n.tipo; dirB[o + 67] = 1;   // negro
      dv.setUint32(o + 68, n.izq, true); dv.setUint32(o + 72, n.der, true); dv.setUint32(o + 76, n.tipo === 2 ? LIBRE : n.hijo, true);
      dv.setUint32(o + 116, n.tipo === 1 ? 0 : n.ini, true); dv.setUint32(o + 120, n.tipo === 1 ? 0 : n.datos.length, true);
    });
    const dir0 = poner(dirB);
    // la FAT se describe a sí misma: cuántos sectores ocupa (cada uno lleva 128 entradas)
    let nFat = 1;
    while (Math.ceil((sectores.length + nFat) / 128) > nFat) nFat++;
    const fat0 = sectores.length;
    for (let i = 0; i < nFat; i++) fat.push(FATSEC);
    if (nFat > 109) throw new Error("Archivo demasiado grande para cifrar aquí");
    const fatB = new Uint8Array(nFat * S).fill(0xff), fv = new DataView(fatB.buffer);
    fat.forEach((x, i) => fv.setUint32(i * 4, x, true));
    for (let i = 0; i < nFat; i++) sectores.push(fatB.subarray(i * S, (i + 1) * S));
    const enc = new Uint8Array(512), ev = new DataView(enc.buffer);
    enc.set(FIRMA);
    ev.setUint16(24, 0x003e, true); ev.setUint16(26, 3, true); ev.setUint16(28, 0xfffe, true); ev.setUint16(30, 9, true); ev.setUint16(32, 6, true);
    ev.setUint32(44, nFat, true); ev.setUint32(48, dir0, true); ev.setUint32(56, CORTE, true);
    ev.setUint32(60, miniFat0, true); ev.setUint32(64, nMiniFat, true); ev.setUint32(68, FIN, true); ev.setUint32(72, 0, true);
    for (let i = 0; i < 109; i++) ev.setUint32(76 + i * 4, i < nFat ? fat0 + i : LIBRE, true);
    const out = new Uint8Array(512 + sectores.length * S);
    out.set(enc);
    sectores.forEach((b, i) => out.set(b, 512 + i * S));
    return out;
  }

  // ------------------------------------------------------------ MS-OFFCRYPTO
  const b64 = (s) => { const t = atob(s), r = new Uint8Array(t.length); for (let i = 0; i < t.length; i++) r[i] = t.charCodeAt(i); return r; };
  const a64 = (u) => { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const unir = (...ps) => { const n = ps.reduce((s, p) => s + p.length, 0), r = new Uint8Array(n); let o = 0; for (const p of ps) { r.set(p, o); o += p.length; } return r; };
  const le32 = (n) => { const r = new Uint8Array(4); new DataView(r.buffer).setUint32(0, n >>> 0, true); return r; };
  const utf16 = (s) => { const r = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { r[i * 2] = s.charCodeAt(i) & 0xff; r[i * 2 + 1] = s.charCodeAt(i) >> 8; } return r; };
  const hex = (s) => new Uint8Array(s.match(/../g).map((x) => parseInt(x, 16)));
  const ajustar = (b, n, relleno) => { if (b.length >= n) return b.slice(0, n); const r = new Uint8Array(n).fill(relleno); r.set(b); return r; };
  const BLOQUE = { verifInput: hex("fea7d2763b4b9e79"), verifValue: hex("d7aa0f6d3061344e"), clave: hex("146e0be7abacd0d6"), hmacClave: hex("5fb2ad010cb9e1f6"), hmacValor: hex("a0677f02b22c8433") };
  const iguales = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);

  /** La clave de una contraseña (Agile): el hash con la sal, las vueltas y el bloque. */
  function hashContrasena(alg, sal, contrasena, vueltas) {
    const f = HASHES[alg].f;
    let h = f(unir(sal, utf16(contrasena)));
    const buf = new Uint8Array(4 + h.length);
    for (let i = 0; i < vueltas; i++) { new DataView(buf.buffer).setUint32(0, i, true); buf.set(h, 4); h = f(buf); }
    return h;
  }
  /** Lee la información de cifrado: { tipo: "agile" | "standard", … }. */
  function leerInfo(info) {
    const dv = new DataView(info.buffer, info.byteOffset, info.byteLength);
    const vMay = dv.getUint16(0, true), vMen = dv.getUint16(2, true);
    if (vMay === 4 && vMen === 4) {
      const x = new DOMParser().parseFromString(new TextDecoder().decode(info.subarray(8)), "application/xml");
      const kd = x.getElementsByTagNameNS("*", "keyData")[0], ek = x.getElementsByTagNameNS("*", "encryptedKey")[0], di = x.getElementsByTagNameNS("*", "dataIntegrity")[0];
      if (!kd || !ek) throw new Error("Este archivo usa un cifrado que aquí no se reconoce (por certificado)");
      const a = (n, k) => n.getAttribute(k);
      return { tipo: "agile", datos: { sal: b64(a(kd, "saltValue")), bits: +a(kd, "keyBits"), bloque: +a(kd, "blockSize"), hash: a(kd, "hashAlgorithm"), cifra: a(kd, "cipherAlgorithm"), cadena: a(kd, "cipherChaining") },
        clave: { sal: b64(a(ek, "saltValue")), bits: +a(ek, "keyBits"), bloque: +a(ek, "blockSize"), hash: a(ek, "hashAlgorithm"), vueltas: +a(ek, "spinCount"),
          vIn: b64(a(ek, "encryptedVerifierHashInput")), vVal: b64(a(ek, "encryptedVerifierHashValue")), valor: b64(a(ek, "encryptedKeyValue")) },
        integridad: di ? { clave: b64(a(di, "encryptedHmacKey")), valor: b64(a(di, "encryptedHmacValue")) } : null };
    }
    if ((vMay === 3 || vMay === 4) && vMen === 2) {
      // Standard (Office 2007): encabezado binario y verificador; AES-128/192/256 en modo ECB con SHA-1
      const tamEnc = dv.getUint32(8, true), o = 12;
      const algId = dv.getUint32(o + 8, true), bits = dv.getUint32(o + 16, true);
      const v = 12 + tamEnc, salT = dv.getUint32(v, true);
      return { tipo: "standard", alg: algId, bits: bits || 128, sal: info.slice(v + 4, v + 4 + salT), verif: info.slice(v + 4 + salT, v + 20 + salT), tamHash: dv.getUint32(v + 20 + salT, true), verifHash: info.slice(v + 24 + salT, v + 24 + salT + 32) };
    }
    throw new Error(`Cifrado de Office versión ${vMay}.${vMen}: aquí no se reconoce`);
  }
  /** La clave secreta del paquete si la contraseña es buena; null si no. */
  function claveAgile(inf, contrasena) {
    const k = inf.clave, h = hashContrasena(k.hash, k.sal, contrasena, k.vueltas), f = HASHES[k.hash].f, nb = k.bits / 8;
    const derivar = (bloque) => ajustar(f(unir(h, bloque)), nb, 0x36);
    const entrada = cbc(derivar(BLOQUE.verifInput), k.sal, k.vIn, false).subarray(0, k.sal.length);
    const valor = cbc(derivar(BLOQUE.verifValue), k.sal, k.vVal, false).subarray(0, HASHES[k.hash].n);
    if (!iguales(f(entrada), valor)) return null;
    return cbc(derivar(BLOQUE.clave), k.sal, k.valor, false).subarray(0, nb);
  }
  function descifrarAgile(inf, paquete, contrasena) {
    const clave = claveAgile(inf, contrasena);
    if (!clave) return null;
    const d = inf.datos, f = HASHES[d.hash].f, tam = Number(new DataView(paquete.buffer, paquete.byteOffset, 8).getBigUint64(0, true));
    const out = new Uint8Array(Math.ceil(tam / 16) * 16 + 4096);
    let o = 0;
    for (let i = 0, p = 8; p < paquete.length && o < tam; i++, p += 4096) {
      const seg = paquete.subarray(p, Math.min(paquete.length, p + 4096));
      const iv = ajustar(f(unir(d.sal, le32(i))), d.bloque, 0x36);
      const dec = cbc(clave, iv, seg.length % 16 ? ajustar(seg, Math.ceil(seg.length / 16) * 16, 0) : seg, false);
      out.set(dec, o); o += dec.length;
    }
    return out.subarray(0, tam);
  }
  function descifrarStandard(inf, paquete, contrasena) {
    // MS-OFFCRYPTO 2.3.4.7: H0 = SHA1(sal + contraseña), 50 000 vueltas, H final con el bloque 0, y la clave con X1/X2
    let h = sha1(unir(inf.sal, utf16(contrasena)));
    for (let i = 0; i < 50000; i++) h = sha1(unir(le32(i), h));
    h = sha1(unir(h, le32(0)));
    const n = inf.bits / 8, b1 = new Uint8Array(64).fill(0x36), b2 = new Uint8Array(64).fill(0x5c);
    for (let i = 0; i < h.length; i++) { b1[i] ^= h[i]; b2[i] ^= h[i]; }
    const clave = unir(sha1(b1), sha1(b2)).subarray(0, n);
    const v = ecbDescifrar(clave, inf.verif), vh = ecbDescifrar(clave, inf.verifHash).subarray(0, 20);
    if (!iguales(sha1(v), vh)) return null;
    const tam = Number(new DataView(paquete.buffer, paquete.byteOffset, 8).getBigUint64(0, true));
    const datos = paquete.subarray(8), cuerpo = datos.subarray(0, Math.floor(datos.length / 16) * 16);
    return ecbDescifrar(clave, cuerpo).subarray(0, tam);
  }
  /** ¿Es un archivo de Office con contraseña? (archivo compuesto con EncryptionInfo y EncryptedPackage) */
  function cifrado(u) {
    if (!esCFB(u)) return false;
    try { const c = leerCFB(u); return !!(c.EncryptionInfo && c.EncryptedPackage); } catch (e) { return false; }
  }
  /** Abre un archivo de Office con contraseña: el .zip de adentro, o null si la contraseña no es. */
  function descifrar(u, contrasena) {
    const c = leerCFB(u);
    if (!c.EncryptionInfo || !c.EncryptedPackage) throw new Error("El archivo no trae contraseña de Office");
    const inf = leerInfo(c.EncryptionInfo);
    const zip = inf.tipo === "agile" ? descifrarAgile(inf, c.EncryptedPackage, contrasena) : descifrarStandard(inf, c.EncryptedPackage, contrasena);
    if (zip && !(zip[0] === 0x50 && zip[1] === 0x4b)) return null;
    return zip;
  }
  const aleatorio = (n) => { const r = new Uint8Array(n); if (raiz.crypto && raiz.crypto.getRandomValues) raiz.crypto.getRandomValues(r); else for (let i = 0; i < n; i++) r[i] = Math.floor(Math.random() * 256); return r; };
  // las partes de «\x06DataSpaces» que Office escribe junto al paquete cifrado (MS-OFFCRYPTO 2.1)
  function espaciosDeDatos() {
    const cadena = (s) => unir(le32(s.length * 2), utf16(s), new Uint8Array((4 - ((s.length * 2) % 4)) % 4));
    const version = unir(cadena("Microsoft.Container.DataSpaces"), le32(1), le32(1), le32(1));   // lector, actualizador y escritor 1.0
    const entrada = unir(le32(1), le32(0), cadena("EncryptedPackage"), cadena("StrongEncryptionDataSpace"));
    const mapa = unir(le32(8), le32(1), le32(entrada.length + 4), entrada);
    const info = unir(le32(8), le32(1), cadena("StrongEncryptionTransform"));
    const id = cadena("{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"), nombre = cadena("Microsoft.Container.EncryptionTransform");
    const cab = unir(le32(1), id, nombre, le32(1), le32(1), le32(1));
    const primario = unir(le32(cab.length + 4), cab, le32(0), le32(0), le32(0), le32(4));
    return { "\u0006DataSpaces/Version": version, "\u0006DataSpaces/DataSpaceMap": mapa, "\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace": info,
      "\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary": primario };
  }
  /** Guarda un .docx/.xlsx/.pptx (bytes del .zip) con contraseña, como Office 365: Agile con AES-256 y SHA-512. */
  function cifrar(zip, contrasena, vueltas) {
    vueltas = vueltas || 100000;
    const salDatos = aleatorio(16), salClave = aleatorio(16), secreta = aleatorio(32), verif = aleatorio(16);
    const h = hashContrasena("SHA512", salClave, contrasena, vueltas);
    const derivar = (bloque) => ajustar(sha512(unir(h, bloque)), 32, 0x36);
    const encVIn = cbc(derivar(BLOQUE.verifInput), salClave, verif, true);
    const encVVal = cbc(derivar(BLOQUE.verifValue), salClave, sha512(verif), true);
    const encClave = cbc(derivar(BLOQUE.clave), salClave, secreta, true);
    // el paquete: tamaño (8 bytes) y segmentos de 4096 cifrados, cada uno con su vector
    const n = zip.length, segs = [];
    const tam = new Uint8Array(8); new DataView(tam.buffer).setBigUint64(0, BigInt(n), true);
    for (let i = 0, o = 0; o < n; i++, o += 4096) {
      const seg = zip.subarray(o, Math.min(n, o + 4096));
      const iv = ajustar(sha512(unir(salDatos, le32(i))), 16, 0x36);
      segs.push(cbc(secreta, iv, ajustar(seg, Math.ceil(seg.length / 16) * 16, 0), true));
    }
    const paquete = unir(tam, ...segs);
    // integridad: HMAC-SHA512 del paquete, con su clave y su valor cifrados con la clave secreta
    const claveHmac = aleatorio(64), valorHmac = hmac("SHA512", claveHmac, paquete);
    const encHmacK = cbc(secreta, ajustar(sha512(unir(salDatos, BLOQUE.hmacClave)), 16, 0x36), claveHmac, true);
    const encHmacV = cbc(secreta, ajustar(sha512(unir(salDatos, BLOQUE.hmacValor)), 16, 0x36), valorHmac, true);
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${a64(salDatos)}"/><dataIntegrity encryptedHmacKey="${a64(encHmacK)}" encryptedHmacValue="${a64(encHmacV)}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="${vueltas}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${a64(salClave)}" encryptedVerifierHashInput="${a64(encVIn)}" encryptedVerifierHashValue="${a64(encVVal)}" encryptedKeyValue="${a64(encClave)}"/></keyEncryptor></keyEncryptors></encryption>`;
    const info = unir(new Uint8Array([4, 0, 4, 0, 0x40, 0, 0, 0]), new TextEncoder().encode(xml));
    return escribirCFB(Object.assign({ EncryptionInfo: info, EncryptedPackage: paquete }, espaciosDeDatos()));
  }

  // ------------------------------------------------------------ protección de edición (no cifra: pide contraseña para cambiar)
  // MS-OFFCRYPTO 2.4.2.4: H0 = H(sal + contraseña en UTF-16), Hn = H(Hn-1 + vuelta) — la vuelta va DESPUÉS (al revés que al
  // cifrar). Excel (hoja, libro, «contraseña para modificar») y PowerPoint usan la contraseña tal cual; Word primero la pasa
  // por su clave vieja de 32 bits (ECMA-376 parte 4, 2.15.1.28) escrita en hexadecimal con los bytes al revés.
  const ALG_SID = { 4: "SHA1", 12: "SHA256", 13: "SHA384", 14: "SHA512" };
  function hashProteccion(contrasena, alg, sal, vueltas) {
    const h0 = HASHES[String(alg || "SHA512").replace(/-/g, "").toUpperCase()];
    if (!h0) return null;
    let h = h0.f(unir(sal, utf16(contrasena)));
    const buf = new Uint8Array(h.length + 4), dv = new DataView(buf.buffer);
    for (let i = 0; i < vueltas; i++) { buf.set(h, 0); dv.setUint32(h.length, i, true); h = h0.f(buf); }
    return h;
  }
  const CODIGO_INICIAL = [0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3];
  const MATRIZ = [[0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09], [0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF], [0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0],
    [0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40], [0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5], [0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A],
    [0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9], [0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0], [0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC],
    [0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10], [0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168], [0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C],
    [0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD], [0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC], [0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4]];
  /** La clave vieja de 16 bits de Excel y Word (MS-OFFCRYPTO 2.3.7.1): también es el «password» de hoja de Excel 97. */
  function claveVieja16(contrasena) {
    const b = [...contrasena.slice(0, 15)].map((ch) => { const c = ch.charCodeAt(0); return (c & 0xff) || (c >> 8); });
    let v = 0;
    if (b.length) {
      for (let i = b.length - 1; i >= 0; i--) { v = ((v >> 14) & 1) | ((v << 1) & 0x7fff); v ^= b[i]; }
      v = ((v >> 14) & 1) | ((v << 1) & 0x7fff); v ^= b.length; v ^= 0xCE4B;
    }
    return v & 0xffff;
  }
  /** La clave vieja de 32 bits de Word, en hexadecimal con los bytes al revés (como la hashea Word). */
  function claveWord(contrasena) {
    if (!contrasena) return "";
    const b = [...contrasena.slice(0, 15)].map((ch) => { const c = ch.charCodeAt(0); return (c & 0xff) || (c >> 8); });
    let alta = CODIGO_INICIAL[b.length - 1];
    b.forEach((x, i) => { const fila = 15 - b.length + i; for (let bit = 0; bit < 7; bit++) if (x & (1 << bit)) alta ^= MATRIZ[fila][bit]; });
    const baja = claveVieja16(contrasena), k = ((alta & 0xffff) << 16) | baja;
    return [k & 0xff, (k >>> 8) & 0xff, (k >>> 16) & 0xff, (k >>> 24) & 0xff].map((x) => x.toString(16).toUpperCase().padStart(2, "0")).join("");
  }
  /**
   * ¿La contraseña es la de esta protección? p: { alg | sid, hash (base64), sal (base64), vueltas, word: true si es de Word,
   * viejo: "CC1A" (la de 16 bits de Excel) }. Sin hash (protección sin contraseña): cualquiera sirve.
   */
  function contrasenaValida(p, contrasena) {
    if (!p) return true;
    if (p.viejo) return claveVieja16(contrasena).toString(16).toUpperCase().padStart(4, "0") === String(p.viejo).toUpperCase().padStart(4, "0");
    if (!p.hash) return true;
    const alg = p.alg || ALG_SID[+p.sid] || "SHA512";
    const h = hashProteccion(p.word ? claveWord(contrasena) : contrasena, alg, b64(p.sal || ""), +p.vueltas || 0);
    return !!h && a64(h) === p.hash;
  }
  /** Una protección nueva (como la escribe Office 365): SHA-512 con sal y 100 000 vueltas. */
  function nuevaProteccion(contrasena, word) {
    if (!contrasena) return {};
    const sal = aleatorio(16), h = hashProteccion(word ? claveWord(contrasena) : contrasena, "SHA512", sal, 100000);
    return { alg: "SHA-512", hash: a64(h), sal: a64(sal), vueltas: 100000, word: word || undefined };
  }

  NOM.cifradoOffice = { cifrado, descifrar, cifrar, leerCFB, escribirCFB, esCFB, contrasenaValida, nuevaProteccion,
    prueba: { sha1, sha512, hmac, cbc, ecbDescifrar, expandir, cifrarBloque, descifrarBloque, hashContrasena, hashProteccion, claveWord, claveVieja16 } };
})(typeof window !== "undefined" ? window : globalThis);

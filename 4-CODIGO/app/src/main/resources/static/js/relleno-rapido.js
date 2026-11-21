/* NOM · Relleno rápido (Datos → Relleno rápido, Ctrl+E), como el de Excel: de unos ejemplos (lo que se escribió en una
   columna junto a los datos) aprende cómo se arma cada uno con los datos de su fila y llena las demás filas.

   Un programa es una lista de piezas: texto fijo, un pedazo de una de las columnas (entre dos posiciones) o el n-ésimo
   tramo de una clase (números, mayúsculas, minúsculas, letras, letras y números, espacios o un signo), cada pedazo con su
   cambio de mayúsculas (como está, MAYÚSCULAS, minúsculas o «Nombre Propio»). Una posición es fija (contando desde el
   principio o desde el final) o el n-ésimo lugar donde termina un tramo de una clase y empieza uno de otra (desde el
   principio o desde el final). Por cada ejemplo se arma el grafo de todos los programas que lo dan (como en «Automating
   String Processing in Spreadsheets using Input-Output Examples», Gulwani, 2011), se cruzan los grafos de los ejemplos y se
   elige el programa más sencillo, con las preferencias medidas en Excel 2024 (pruebas/office/datos_excel.py, casos ff-):
   un tramo entero antes que un pedazo, contar desde el final cuando empata («Pérez» de «Ana Pérez López» → la penúltima
   palabra), un signo antes que una palabra («Guadalajara» de «Guadalajara, Jalisco» → todo lo de antes de la coma) y texto
   fijo solo para lo que no sale de los datos. Si en una fila el programa no se puede aplicar, se usa el siguiente más
   sencillo que sí (también como Excel). */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});

  // ------------------------------------------------------------------ clases de caracteres y sus tramos
  const CLASES = [["num", /\p{Nd}+/gu], ["may", /\p{Lu}+/gu], ["min", /\p{Ll}+/gu], ["let", /\p{L}+/gu], ["aln", /[\p{L}\p{Nd}]+/gu], ["esp", /\s+/gu]];
  const esSigno = (ch) => !/[\p{L}\p{Nd}\s]/u.test(ch);
  // (un signo es su propia clase: «=» y el número del carácter, para que quepa en las claves con «:»)
  const claseSigno = (ch) => "=" + ch.codePointAt(0);

  const memoTramos = new Map();
  /** Los tramos [inicio, fin) de cada clase en s (y de cada signo que tenga). */
  function tramos(s) {
    let r = memoTramos.get(s);
    if (r) return r;
    r = {};
    for (const [n, re] of CLASES) {
      const lista = [];
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(s))) lista.push([m.index, m.index + m[0].length]);
      if (lista.length) r[n] = lista;
    }
    for (const ch of new Set([...s].filter(esSigno))) {
      const lista = [];
      for (let i = s.indexOf(ch); i >= 0; i = s.indexOf(ch, lista[lista.length - 1][1])) {
        let j = i + ch.length;
        while (s.startsWith(ch, j)) j += ch.length;
        lista.push([i, j]);
      }
      r[claseSigno(ch)] = lista;
    }
    if (memoTramos.size > 20000) memoTramos.clear();
    memoTramos.set(s, r);
    return r;
  }
  /** Los lugares de s donde termina un tramo de la clase a y empieza uno de b, en orden («» en a o b: cualquiera). */
  function lugares(s, a, b) {
    const T = tramos(s);
    const fines = a ? (T[a] || []).map((x) => x[1]) : null, inicios = b ? (T[b] || []).map((x) => x[0]) : null;
    if (fines && inicios) { const hay = new Set(inicios); return fines.filter((p) => hay.has(p)); }
    return fines || inicios || [];
  }

  // ------------------------------------------------------------------ posiciones
  // «A:k»: fija (k ≥ 0 desde el principio; k < 0 desde el final: -1 es el final). «P:a:b:c»: el c-ésimo lugar donde
  // termina un tramo de a y empieza uno de b (c > 0 desde el principio, c < 0 desde el final).
  /** Todas las maneras de nombrar el lugar p de s. */
  function posicionesDe(s, p) {
    const out = new Set(["A:" + p, "A:" + (p - s.length - 1)]);
    const cls = [""].concat(Object.keys(tramos(s)));
    for (const a of cls) for (const b of cls) {
      if (!a && !b) continue;
      const L = lugares(s, a, b), i = L.indexOf(p);
      if (i >= 0) { out.add(`P:${a}:${b}:${i + 1}`); out.add(`P:${a}:${b}:${i - L.length}`); }
    }
    return out;
  }
  /** El lugar de s que nombra q, o -1 si en s no existe. */
  function lugar(s, q) {
    const x = q.split(":");
    if (x[0] === "A") { const k = +x[1], p = k >= 0 ? k : s.length + 1 + k; return p >= 0 && p <= s.length ? p : -1; }
    const L = lugares(s, x[1], x[2]), c = +x[3], i = c > 0 ? c - 1 : L.length + c;
    return i >= 0 && i < L.length ? L[i] : -1;
  }
  /** Qué tan rebuscada es una posición (menos es mejor). */
  function costoPos(q) {
    const x = q.split(":");
    // (las fijas, contando desde el principio antes que desde el final: Excel saca «5678» de «Cuenta 0012345678» desde el
    // lugar 13, y de «Tarjeta 4111222233334444», «22233334444»)
    // (y una fija en medio de una palabra cuesta más que una letra fija: de «alopez», Excel no saca la «a» de «Ana»)
    if (x[0] === "A") { const k = +x[1]; return k === 0 || k === -1 ? 0 : k > 0 ? 2 + 0.001 * k : 2.1 + 0.001 * -k; }
    const a = x[1], b = x[2], c = +x[3];
    const base = a.startsWith("=") || b.startsWith("=") ? 0.15 : a === "esp" || b === "esp" ? 0.3 : a && b ? 0.35 : 0.4;
    return base + 0.1 * (Math.abs(c) - 1) + (c > 0 ? 0.001 : 0);
  }

  // ------------------------------------------------------------------ mayúsculas
  const propio = (s) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (m, a, b) => a + b.toUpperCase());   // como NOMPROPIO
  const CAMBIO = { "": (s) => s, may: (s) => s.toUpperCase(), min: (s) => s.toLowerCase(), prop: propio };
  // (cambiar las mayúsculas de una sola letra cuesta más que una letra fija: con un ejemplo, Excel saca «alopez» de «Ana
  // Lopez» como «a» y el apellido en minúsculas —«agomez»—; de algo más largo, sí: «ana.lopez@…»)
  const costoCaso = (caso, largo) => (!caso ? 0 : largo <= 1 ? 3 : 0.1);
  const ORDEN_CASO = { "": 0, may: 1, min: 1, prop: 2 };
  /** Con qué cambio de mayúsculas el pedazo p da s. */
  function casosDe(p, s) {
    const r = new Set();
    if (p === s) r.add("");
    if (/\p{L}/u.test(s)) for (const k of ["may", "min", "prop"]) if (CAMBIO[k](p) === s) r.add(k);
    return r;
  }
  const soloSignos = (s) => !/[\p{L}\p{Nd}]/u.test(s);

  // ------------------------------------------------------------------ el grafo de un ejemplo
  // Nodos: los lugares de la salida (0…largo). Arista i→j: todas las piezas que dan salida[i:j] en ese ejemplo:
  // { c: Set de textos fijos, s: [{ n (columna), p1: Set, p2: Set (posiciones), k: Set (cambios) }], t: Set de «n:clase:c:caso» }.
  function grafo(entradas, salida) {
    const L = salida.length, sal = new Map();
    const bajas = entradas.map((s) => { const b = s.toLowerCase(); return b.length === s.length ? b : null; });
    for (let i = 0; i < L; i++) {
      const m = new Map();
      for (let j = i + 1; j <= L; j++) {
        const sub = salida.slice(i, j), baja = sub.toLowerCase(), e = { c: new Set([sub]), s: [], t: new Set(), l: sub.length };
        entradas.forEach((s, n) => {
          if (sub.length > s.length) return;
          const donde = bajas[n] != null ? bajas[n] : s, que = bajas[n] != null ? baja : sub;
          for (let k = donde.indexOf(que); k >= 0; k = donde.indexOf(que, k + 1)) {
            const casos = casosDe(s.slice(k, k + sub.length), sub);
            if (!casos.size) continue;
            e.s.push({ n, p1: posicionesDe(s, k), p2: posicionesDe(s, k + sub.length), k: casos });
            const T = tramos(s);
            for (const cls in T) {
              const lista = T[cls], idx = lista.findIndex((x) => x[0] === k && x[1] === k + sub.length);
              if (idx >= 0) casos.forEach((cs) => { e.t.add(`${n}:${cls}:${idx + 1}:${cs}`); e.t.add(`${n}:${cls}:${idx - lista.length}:${cs}`); });
            }
          }
        });
        m.set(j, e);
      }
      sal.set(i, m);
    }
    return { ini: 0, fin: L, sal };
  }

  const cruce = (a, b) => { const r = new Set(); for (const x of a) if (b.has(x)) r.add(x); return r; };
  /** Las piezas que están en las dos aristas (null si ninguna). */
  function cruzarArista(a, b) {
    const c = cruce(a.c, b.c), t = cruce(a.t, b.t), s = [];
    for (const x of a.s) for (const y of b.s) {
      if (x.n !== y.n) continue;
      const k = cruce(x.k, y.k); if (!k.size) continue;
      const p1 = cruce(x.p1, y.p1); if (!p1.size) continue;
      const p2 = cruce(x.p2, y.p2); if (!p2.size) continue;
      s.push({ n: x.n, p1, p2, k });
    }
    return c.size || t.size || s.length ? { c, s, t, l: Math.min(a.l, b.l) } : null;
  }
  /** Los programas que dan los dos grafos (el grafo producto), o null si no hay ninguno. */
  function cruzar(g1, g2) {
    const sal = new Map(), clave = (a, b) => a + "," + b, ini = clave(g1.ini, g2.ini), fin = clave(g1.fin, g2.fin);
    const pila = [[g1.ini, g2.ini]], vistos = new Set([ini]);
    while (pila.length) {
      const [a, b] = pila.pop(), sa = g1.sal.get(a), sb = g2.sal.get(b);
      if (!sa || !sb) continue;
      const m = new Map();
      for (const [a2, e1] of sa) for (const [b2, e2] of sb) {
        const e = cruzarArista(e1, e2);
        if (!e) continue;
        const kb = clave(a2, b2);
        m.set(kb, e);
        if (!vistos.has(kb)) { vistos.add(kb); pila.push([a2, b2]); }
      }
      if (m.size) sal.set(clave(a, b), m);
    }
    return podar({ ini, fin, sal });
  }
  /** Quita lo que no llega al final; null si del principio no se llega. */
  function podar(g) {
    const memo = new Map();
    const llega = (a) => {
      if (a === g.fin) return true;
      if (memo.has(a)) return memo.get(a);
      memo.set(a, false);
      const m = g.sal.get(a);
      let ok = false;
      if (m) for (const b of [...m.keys()]) { if (llega(b)) ok = true; else m.delete(b); }
      if (!ok) g.sal.delete(a);
      memo.set(a, ok);
      return ok;
    };
    return llega(g.ini) ? g : null;
  }

  // ------------------------------------------------------------------ elegir y aplicar
  const textoTramo = (s, cls, c) => { const L = tramos(s)[cls]; if (!L) return null; const i = c > 0 ? c - 1 : L.length + c; return i >= 0 && i < L.length ? s.slice(L[i][0], L[i][1]) : null; };
  const ordenar = (set) => [...set].map((q) => [q, costoPos(q)]).sort((a, b) => a[1] - b[1]);
  // (texto fijo: barato si son signos o espacios —los separadores—; una letra fija sí puede ser —«a» en «agomez»—, pero un
  // dígito casi siempre sale de los datos —«(33) 9876-5432», no «(55) …» fijo—)
  const COSTO_FIJO = (s) => { let c = 1.5; for (const ch of s) c += /\p{Nd}/u.test(ch) ? 3 : /\p{L}/u.test(ch) ? 1.5 : 0.5; return c; };
  /**
   * La pieza más sencilla de una arista: { costo, pieza, texto }. Con fila (las entradas de una fila a llenar), solo una que
   * se pueda aplicar ahí (y su texto); sin fila, la más sencilla de todas.
   */
  function mejorPieza(e, fila) {
    let mejor = null;
    const prueba = (costo, pieza, texto) => { if (texto != null && (!mejor || costo < mejor.costo - 1e-9)) mejor = { costo, pieza, texto }; };
    for (const s of e.c) prueba(COSTO_FIJO(s), { tipo: "fijo", s }, s);
    for (const q of e.t) {
      const [n, cls, c, caso] = q.split(":"), cn = +c;
      const extra = cls === "esp" || cls.startsWith("=") ? 3 : cls === "aln" ? 0.02 : 0;
      const costo = 1.3 + 0.1 * (Math.abs(cn) - 1) + (cn > 0 ? 0.001 : 0) + costoCaso(caso, e.l) + extra;
      if (mejor && costo >= mejor.costo) continue;
      const tx = fila ? textoTramo(fila[+n] || "", cls, cn) : "";
      if (fila && !tx) continue;
      if (!fila || !soloSignos(tx) || cls.startsWith("=") || cls === "esp") prueba(costo, { tipo: "tramo", n: +n, cls, c: cn, caso }, fila ? CAMBIO[caso](tx) : "");
    }
    for (const x of e.s) {
      const caso = [...x.k].sort((a, b) => ORDEN_CASO[a] - ORDEN_CASO[b])[0], cc = costoCaso(caso, e.l);
      const P1 = x.o1 || (x.o1 = ordenar(x.p1)), P2 = x.o2 || (x.o2 = ordenar(x.p2));
      const s = fila ? fila[x.n] || "" : null;
      for (const [q1, c1] of P1) {
        if (mejor && 1 + c1 + P2[0][1] + cc >= mejor.costo) break;
        const l1 = s != null ? lugar(s, q1) : 0;
        if (l1 < 0) continue;
        for (const [q2, c2] of P2) {
          const costo = 1 + c1 + c2 + cc;
          if (mejor && costo >= mejor.costo) break;
          if (s == null) { prueba(costo + (x.sep ? 3 : 0), { tipo: "pedazo", n: x.n, p1: q1, p2: q2, caso }, ""); break; }
          const l2 = lugar(s, q2);
          if (l2 <= l1) continue;
          const tx = s.slice(l1, l2);
          prueba(costo + (soloSignos(tx) ? 3 : 0), { tipo: "pedazo", n: x.n, p1: q1, p2: q2, caso }, CAMBIO[caso](tx));
          break;
        }
      }
    }
    return mejor;
  }
  /** El programa más sencillo del grafo; con fila, el más sencillo que se puede aplicar a esa fila (y lo que da). */
  function mejorPrograma(g, fila) {
    const memo = new Map();
    const ir = (a) => {
      if (a === g.fin) return { costo: 0, piezas: [], texto: "" };
      if (memo.has(a)) return memo.get(a);
      let mejor = null;
      const m = g.sal.get(a);
      if (m) for (const [b, e] of m) {
        const resto = ir(b);
        if (!resto) continue;
        const p = mejorPieza(e, fila);
        if (!p) continue;
        const costo = p.costo + resto.costo;
        if (!mejor || costo < mejor.costo - 1e-9) mejor = { costo, piezas: [p.pieza].concat(resto.piezas), texto: p.texto + resto.texto };
      }
      memo.set(a, mejor);
      return mejor;
    };
    return ir(g.ini);
  }
  /** Lo que da el programa con las entradas de una fila, o null si ahí no se puede. */
  function aplicar(piezas, fila) {
    let out = "";
    for (const p of piezas) {
      if (p.tipo === "fijo") { out += p.s; continue; }
      const s = fila[p.n] || "";
      if (p.tipo === "tramo") { const tx = textoTramo(s, p.cls, p.c); if (!tx) return null; out += CAMBIO[p.caso](tx); continue; }
      const l1 = lugar(s, p.p1), l2 = lugar(s, p.p2);
      if (l1 < 0 || l2 <= l1) return null;
      out += CAMBIO[p.caso](s.slice(l1, l2));
    }
    return out;
  }

  /**
   * Aprende de los ejemplos y llena: ejemplos [{ entradas: [texto…], salida }], filas [[texto…]] (las entradas de cada
   * fila a llenar). Devuelve { salidas: [texto|null…], programa } o null si ningún programa da todos los ejemplos.
   */
  function aprender(ejemplos) {
    let g = null;
    for (const e of ejemplos) {
      const g2 = grafo(e.entradas.map((x) => String(x == null ? "" : x)), String(e.salida));
      g = g ? cruzar(g, g2) : podar(g2);
      if (!g) return null;
    }
    return g;
  }
  function rellenar(ejemplos, filas) {
    let buenos = ejemplos.filter((e) => e.salida != null && String(e.salida) !== ""), encabezado = false;
    if (!buenos.length) return null;
    let g = aprender(buenos);
    // (el primero puede ser el encabezado —«Nombre completo» / «Primer nombre»—: Excel no lo toma como ejemplo)
    if (!g && buenos.length > 1) { g = aprender(buenos.slice(1)); encabezado = !!g; }
    if (!g) return null;
    const prog = mejorPrograma(g, null);
    if (!prog) return null;
    const salidas = filas.map((f) => {
      const fila = f.map((x) => String(x == null ? "" : x));
      if (!fila.some((x) => x !== "")) return null;
      const r = aplicar(prog.piezas, fila);
      if (r != null) return r;
      const otro = mejorPrograma(g, fila);   // el siguiente más sencillo que sí se puede aplicar aquí
      return otro ? otro.texto : null;
    });
    return { salidas, programa: prog.piezas, encabezado };
  }

  NOM.rellenoRapido = { rellenar, _grafo: grafo, _aplicar: aplicar };
})();

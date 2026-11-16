/* NOM-COLAB · SmartArt sin dibujo guardado: se acomoda con su diseño (layoutDef), sus datos, sus colores (colorsDef) y
   su estilo (styleDef), como lo hace PowerPoint al abrir un archivo que no trae el dibujo ya hecho (los que guardan
   otros programas). Interpreta el lenguaje de los diseños de la norma (ECMA-376, parte 1, 21.4): nodos de diseño,
   recorridos (forEach) por los datos, condiciones (choose/if), restricciones (constrLst) y reglas, con los algoritmos
   compuesto, línea, serpiente, ciclo, jerarquía, pirámide, texto, espacio y conector.

   NOM.smartart.acomodar({ datos, diseno, colores, estilo }, { w, h, color(el) → "#RRGGBB", anchoLinea(idx) → EMU,
     medir(texto, pt) → ancho en EMU }) → [{ x, y, w, h, rot, prst, adj, relleno, linea, texto, tam, colorTexto, alin,
     ancla, vineta }]   (medidas en EMU, desde la esquina del marco del SmartArt) */
(() => {
  const NOM = window.NOM;
  const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
  const hijos = (n) => [...(n ? n.children : [])];
  const hijo = (n, nm) => hijos(n).find((x) => x.localName === nm) || null;
  const at = (n, k) => (n ? n.getAttribute(k) : null);
  const num = (v, d) => (v == null || v === "" || v === "NaN" ? d : v === "INF" ? Infinity : +v);

  // ------------------------------------------------------------ datos: los puntos y cómo cuelgan
  function leerDatos(doc) {
    const pts = new Map();
    for (const p of doc.getElementsByTagNameNS(DGM, "pt")) {
      const id = at(p, "modelId"), tipo = at(p, "type") || "node";
      const t = hijo(p, "t"), pars = t ? [...t.getElementsByTagNameNS("*", "p")].map((q) => [...q.getElementsByTagNameNS("*", "t")].map((x) => x.textContent).join("")) : [];
      const ps = hijo(p, "prSet");
      pts.set(id, { id, tipo, texto: pars.filter((x, i) => x !== "" || i === 0), hijos: [], padre: null, sib: null, par: null,
        pres: ps ? { nombre: at(ps, "presName"), lbl: at(ps, "presStyleLbl"), idx: num(at(ps, "presStyleIdx"), 0), cnt: num(at(ps, "presStyleCnt"), 1), presAssocID: at(ps, "presAssocID") } : null });
    }
    const cxns = [...doc.getElementsByTagNameNS(DGM, "cxn")];
    cxns.filter((c) => (at(c, "type") || "parOf") === "parOf").sort((a, b) => num(at(a, "srcOrd"), 0) - num(at(b, "srcOrd"), 0)).forEach((c) => {
      const s = pts.get(at(c, "srcId")), d = pts.get(at(c, "destId"));
      if (!s || !d) return;
      s.hijos.push(d); d.padre = s;
      d.sib = pts.get(at(c, "sibTransId")) || null; d.par = pts.get(at(c, "parTransId")) || null;
      if (d.sib) d.sib.de = d;
      if (d.par) d.par.de = d;
    });
    // los puntos de presentación (pres) de cada dato: su etiqueta de estilo y su lugar entre los de su clase
    const presDe = new Map();   // id del dato → [pres]
    cxns.filter((c) => at(c, "type") === "presOf").forEach((c) => { const s = at(c, "srcId"), d = pts.get(at(c, "destId")); if (d && d.pres) (presDe.get(s) || presDe.set(s, []).get(s)).push(d.pres); });
    const raiz = [...pts.values()].find((p) => p.tipo === "doc");
    return { pts, raiz, presDe };
  }
  const esNodo = (p) => p && (p.tipo === "node" || p.tipo === "asst");

  // ------------------------------------------------------------ recorridos por los datos (forEach, presOf, cnt)
  function paso(ps, eje, tipoPt, opc) {
    const out = [];
    for (const p of ps) {
      if (!p) continue;
      switch (eje) {
        case "self": out.push(p); break;
        case "ch": out.push(...p.hijos); break;
        case "des": { const r = (q) => q.hijos.forEach((h) => { out.push(h); r(h); }); r(p); break; }
        case "desOrSelf": { out.push(p); const r = (q) => q.hijos.forEach((h) => { out.push(h); r(h); }); r(p); break; }
        case "par": if (p.padre) out.push(p.padre); break;
        case "ancst": for (let q = p.padre; q; q = q.padre) out.push(q); break;
        case "root": { let q = p; while (q.padre) q = q.padre; out.push(q); break; }
        case "followSib":
          if (tipoPt === "sibTrans") {
            const her = p.padre ? p.padre.hijos : [], k = her.indexOf(p);
            if (p.sib && (k < her.length - 1 || opc.ultimo)) out.push(p.sib);
          } else if (p.padre) out.push(...p.padre.hijos.slice(p.padre.hijos.indexOf(p) + 1));
          break;
        case "precedSib":
          if (tipoPt === "sibTrans") { const her = p.padre ? p.padre.hijos : [], k = her.indexOf(p); if (k > 0 && her[k - 1].sib) out.push(her[k - 1].sib); }
          else if (p.padre) out.push(...p.padre.hijos.slice(0, p.padre.hijos.indexOf(p)).reverse());
          break;
        default: break;
      }
    }
    const ok = (q) => {
      switch (tipoPt || "all") {
        case "all": return true;
        case "node": return q.tipo === "node" || q.tipo === "asst";
        case "nonAsst": return q.tipo === "node";
        case "asst": return q.tipo === "asst";
        case "doc": return q.tipo === "doc";
        case "parTrans": return q.tipo === "parTrans";
        case "sibTrans": return q.tipo === "sibTrans";
        case "nonNorm": return q.tipo !== "node";
        default: return true;
      }
    };
    return out.filter(ok);
  }
  function puntos(ctx, el) {
    const ejes = String(at(el, "axis") || "none").split(/\s+/), tipos = String(at(el, "ptType") || "all").split(/\s+/);
    if (ejes[0] === "none") return [];
    let ps = [ctx];
    ejes.forEach((e, k) => { ps = paso(ps, e, tipos[Math.min(k, tipos.length - 1)], { ultimo: at(el, "hideLastTrans") === "0" }); });
    const st = num(at(el, "st"), 1), cnt = num(at(el, "cnt"), 0), step = num(at(el, "step"), 1);
    if (st !== 1 || step !== 1) ps = ps.filter((_, i) => i >= st - 1 && (i - (st - 1)) % Math.max(1, step) === 0);
    return cnt > 0 ? ps.slice(0, cnt) : ps;
  }

  // ------------------------------------------------------------ el árbol de lo que se dibuja (instancias)
  function profundidad(p) { let d = 0; for (let q = p; q && q.padre; q = q.padre) d++; return d; }
  function maxProf(p) { return p.hijos.length ? 1 + Math.max(...p.hijos.map(maxProf)) : 0; }
  function condicion(c, ctx, inst) {
    const fn = at(c, "func"), op = at(c, "op") || "equ", val = at(c, "val");
    let x;
    switch (fn) {
      case "cnt": x = puntos(ctx, c).length; break;
      case "pos": { const her = ctx.padre ? ctx.padre.hijos.filter(esNodo) : [ctx]; x = her.indexOf(ctx) + 1; break; }
      case "revPos": { const her = ctx.padre ? ctx.padre.hijos.filter(esNodo) : [ctx]; x = her.length - her.indexOf(ctx); break; }
      case "posEven": { const her = ctx.padre ? ctx.padre.hijos.filter(esNodo) : [ctx]; x = (her.indexOf(ctx) + 1) % 2 === 0 ? 1 : 0; break; }
      case "posOdd": { const her = ctx.padre ? ctx.padre.hijos.filter(esNodo) : [ctx]; x = (her.indexOf(ctx) + 1) % 2 === 1 ? 1 : 0; break; }
      case "depth": x = profundidad(ctx); break;
      case "maxDepth": x = maxProf(ctx); break;
      case "var": x = (inst.vars || {})[at(c, "arg")]; if (x == null) x = { dir: "norm", hierBranch: "std", animOne: "one", animLvl: "none", resizeHandles: "rel", bulletEnabled: "0" }[at(c, "arg")]; break;
      default: return false;
    }
    const a = isNaN(+x) || isNaN(+val) ? String(x) : +x, b = isNaN(+x) || isNaN(+val) ? String(val) : +val;
    return { equ: a === b, neq: a !== b, gt: a > b, lt: a < b, gte: a >= b, lte: a <= b }[op] || false;
  }
  function instanciar(el, ctx, varsPadre, cuenta) {
    const inst = { nombre: at(el, "name") || "", lbl: at(el, "styleLbl") || null, ctx, alg: null, forma: null, presOf: [], constr: [], reglas: [], hijos: [], vars: Object.assign({}, varsPadre), v: {} };
    procesar(hijos(el), inst, ctx, cuenta);
    return inst;
  }
  function procesar(ns, inst, ctx, cuenta) {
    for (const c of ns) {
      if (++cuenta.n > 4000) return;
      switch (c.localName) {
        case "alg": { const p = {}; hijos(c).filter((x) => x.localName === "param").forEach((x) => { p[at(x, "type")] = at(x, "val"); }); inst.alg = { tipo: at(c, "type"), p }; break; }
        case "shape": {
          const adj = {};
          hijos(hijo(c, "adjLst")).forEach((a) => { adj[at(a, "idx")] = num(at(a, "val"), 0); });
          inst.forma = { tipo: at(c, "type") || "none", rot: num(at(c, "rot"), 0), oculta: at(c, "hideGeom") === "1" || at(c, "hideGeom") === "true", adj, zOff: num(at(c, "zOrderOff"), 0) };
          break;
        }
        case "presOf": inst.presOf = puntos(ctx, c); break;
        case "constrLst": hijos(c).forEach((x) => inst.constr.push(Object.fromEntries([...x.attributes].map((a) => [a.name, a.value])))); break;
        case "ruleLst": hijos(c).forEach((x) => inst.reglas.push(Object.fromEntries([...x.attributes].map((a) => [a.name, a.value])))); break;
        case "varLst": hijos(c).forEach((x) => { inst.vars[x.localName] = at(x, "val") != null ? at(x, "val") : { dir: "norm", hierBranch: "std" }[x.localName] || "1"; }); break;
        case "layoutNode": inst.hijos.push(instanciar(c, ctx, inst.vars, cuenta)); break;
        case "forEach": for (const p of puntos(ctx, c)) procesar(hijos(c), inst, p, cuenta); break;
        case "choose": {
          for (const r of hijos(c)) {
            if (r.localName === "if" && condicion(r, ctx, inst)) { procesar(hijos(r), inst, ctx, cuenta); break; }
            if (r.localName === "else") { procesar(hijos(r), inst, ctx, cuenta); break; }
          }
          break;
        }
        default: break;
      }
    }
  }

  // ------------------------------------------------------------ restricciones: tamaños, márgenes, letra y espacios
  // (los filtros: forName y ptType para «a quién», refForName y refPtType para «de quién»)
  const coincide = (inst, c, pre) => {
    const nom = c[pre === "for" ? "forName" : "refForName"], pt = c[pre === "for" ? "ptType" : "refPtType"];
    if (nom && inst.nombre !== nom) return false;
    if (!pt || pt === "all") return true;
    const t = inst.ctx ? inst.ctx.tipo : null;
    return pt === "node" ? esNodo(inst.ctx) : pt === "nonAsst" ? t === "node" : pt === "asst" ? t === "asst" : pt === "sibTrans" ? t === "sibTrans" : pt === "parTrans" ? t === "parTrans" : pt === "doc" ? t === "doc" : true;
  };
  const desc = (inst) => { const out = []; const r = (q) => q.hijos.forEach((h) => { out.push(h); r(h); }); r(inst); return out; };
  function objetivos(inst, c, pre) {
    const para = c[pre] || "self";
    if (para === "self") return [inst];
    if (para === "ch") return inst.hijos.filter((h) => coincide(h, c, pre));
    if (para === "des") return desc(inst).filter((h) => coincide(h, c, pre));
    return [];
  }
  /** Valor de una medida de una instancia (lo que ya le pusieron, o su caja). */
  const valor = (q, tipo) => (q.v[tipo] != null ? q.v[tipo] : tipo === "w" && q.caja ? q.caja.w : tipo === "h" && q.caja ? q.caja.h : undefined);
  /** Las restricciones de una instancia: las de sus hijos y descendientes (propias = false) o las suyas (propias = true). */
  function aplicarRestricciones(inst, propias) {
    for (const c of inst.constr) {
      const deSi = !c.for || c.for === "self";
      if (deSi !== !!propias) continue;
      let x;
      if (c.refType && c.refType !== "none") {
        const r = objetivos(inst, c, "refFor")[0];
        const rv = r ? valor(r, c.refType) : undefined;
        if (rv == null) continue;
        x = rv * num(c.fact, 1);
      } else if (c.val == null || c.val === "") continue;   // (op="equ" sin valor: «iguales entre sí», no un valor)
      else x = num(c.val, 0);
      for (const t of objetivos(inst, c, "for")) {
        const ant = t.v[c.type];
        if (c.op === "gte") t.v[c.type] = ant == null ? x : Math.max(ant, x);
        else if (c.op === "lte") t.v[c.type] = ant == null ? x : Math.min(ant, x);
        else t.v[c.type] = x;
      }
    }
  }

  // ------------------------------------------------------------ algoritmos: dónde va cada hijo
  const visibles = (inst) => inst.hijos;
  const tam = (h, padre) => ({ w: h.v.w != null ? h.v.w : padre.caja.w, h: h.v.h != null ? h.v.h : padre.caja.h });
  function escalarHijo(h, k) { if (k === 1) return; ["w", "h"].forEach((t) => { if (h.v[t] != null) h.v[t] *= k; }); }
  function colocar(inst) {
    // primero lo que el padre dice de sus hijos, luego lo que cada hijo dice de sí (su tamaño preferido o su proporción);
    // el algoritmo los acomoda (y los agranda o achica para caber); cada hijo se acomoda luego con su tamaño final
    aplicarRestricciones(inst, false);
    const tipo = inst.alg ? inst.alg.tipo : "composite", p = (inst.alg && inst.alg.p) || {}, C = inst.caja;
    const hs = visibles(inst);
    hs.forEach((h) => aplicarRestricciones(h, true));
    if (tipo === "lin") lineal(inst, hs, p);
    else if (tipo === "snake") serpiente(inst, hs, p);
    else if (tipo === "cycle") ciclo(inst, hs, p);
    else if (tipo === "hierRoot" || tipo === "hierChild") jerarquia(inst, hs, p, tipo);
    else if (tipo === "pyra") piramide(inst, hs, p);
    else {
      // compuesto (y los demás): cada hijo por sus restricciones l, t, w, h, r, b, ctrX y ctrY dentro de la caja
      hs.forEach((h) => {
        const v = h.v;
        let w = v.w != null ? v.w : v.l != null && v.r != null ? v.r - v.l : C.w, hh = v.h != null ? v.h : v.t != null && v.b != null ? v.b - v.t : C.h;
        const x = v.l != null ? v.l : v.r != null ? v.r - w : v.ctrX != null ? v.ctrX - w / 2 : 0;
        const y = v.t != null ? v.t : v.b != null ? v.b - hh : v.ctrY != null ? v.ctrY - hh / 2 : 0;
        h.caja = { x: C.x + x, y: C.y + y, w: Math.max(0, w), h: Math.max(0, hh) };
      });
    }
    hs.forEach((h) => { if (!h.caja) h.caja = { x: C.x, y: C.y, w: C.w, h: C.h }; h.v.w = h.caja.w; h.v.h = h.caja.h; colocar(h); });
  }
  function lineal(inst, hs, p) {
    const C = inst.caja, dir = p.linDir || "fromL", horiz = dir === "fromL" || dir === "fromR";
    const sp = inst.v.sp || 0;
    const ts = hs.map((h) => tam(h, inst));
    const largo = ts.reduce((s, t) => s + (horiz ? t.w : t.h), 0) + sp * Math.max(0, hs.length - 1);
    const ancho = Math.max(0, ...ts.map((t) => (horiz ? t.h : t.w)));
    const k = Math.min(largo ? (horiz ? C.w : C.h) / largo : 1, ancho ? (horiz ? C.h : C.w) / ancho : 1);
    let pos = 0;
    const total = largo * k, ini = ((horiz ? C.w : C.h) - total) / 2;
    hs.forEach((h, i) => {
      const t = { w: ts[i].w * k, h: ts[i].h * k };
      escalarHijo(h, k);
      const a = ini + pos, cruz = ((horiz ? C.h : C.w) - (horiz ? t.h : t.w)) / 2;
      const x = horiz ? (dir === "fromR" ? C.w - a - t.w : a) : cruz, y = horiz ? cruz : dir === "fromB" ? C.h - a - t.h : a;
      h.caja = { x: C.x + x, y: C.y + y, w: t.w, h: t.h };
      h.dir = horiz ? (dir === "fromR" ? 180 : 0) : dir === "fromB" ? 270 : 90;
      pos += (horiz ? t.w : t.h) + sp * k;
    });
  }
  function serpiente(inst, hs, p) {
    const C = inst.caja, celdas = hs.filter((h) => !(h.alg && h.alg.tipo === "sp")), sp = inst.v.sp || 0;
    const n = celdas.length;
    if (!n) return;
    const t0 = tam(celdas[0], inst);
    let mejor = null;
    for (let c = 1; c <= n; c++) {
      const r = Math.ceil(n / c), W = c * t0.w + (c - 1) * sp, H = r * t0.h + (r - 1) * sp, k = Math.min(C.w / W, C.h / H);
      if (!mejor || k > mejor.k + 1e-9) mejor = { c, r, k };
    }
    const { c, r } = mejor, k = mejor.k;
    const w = t0.w * k, h = t0.h * k, s = sp * k, W = c * w + (c - 1) * s, H = r * h + (r - 1) * s;
    const x0 = C.x + (C.w - W) / 2, y0 = C.y + (C.h - H) / 2, derecha = (p.grDir || "tL") === "tR";
    celdas.forEach((q, i) => {
      const fila = Math.floor(i / c), col = i % c, enFila = fila === r - 1 ? n - fila * c : c;
      const desp = fila === r - 1 && (p.off || "ctr") === "ctr" ? ((c - enFila) * (w + s)) / 2 : 0;
      const cc = derecha ? c - 1 - col : col;
      escalarHijo(q, k);
      q.caja = { x: x0 + desp + cc * (w + s), y: y0 + fila * (h + s), w, h };
    });
    hs.filter((h) => h.alg && h.alg.tipo === "sp").forEach((h) => { h.caja = { x: C.x, y: C.y, w: 0, h: 0 }; });
  }
  function ciclo(inst, hs, p) {
    const C = inst.caja, nodos = hs.filter((h) => esNodo(h.ctx) && !(h.alg && (h.alg.tipo === "sp" || h.alg.tipo === "conn"))), n = nodos.length;
    if (!n) return;
    const st = num(p.stAng, 0), span = num(p.spanAng, 360), centro = p.ctrShpMap === "fNode";
    const enAro = centro ? nodos.slice(1) : nodos, m = enAro.length;
    const t0 = tam(nodos[0], inst), sib = inst.v.sibSp || 0;
    let w = t0.w, h = t0.h;
    // el radio con que los nodos no se tocan (con su espacio), y todo se achica para caber
    let R = m > 1 ? (Math.max(w, h) + sib) / (2 * Math.sin(Math.PI / m)) : 0;
    if (centro) R = Math.max(R, Math.max(w, h) + sib);
    const k = Math.min(C.w / (2 * R + w), C.h / (2 * R + h));
    w *= k; h *= k; R *= k;
    const cx = C.x + C.w / 2, cy = C.y + C.h / 2;
    nodos.forEach((q) => escalarHijo(q, k));
    if (centro) nodos[0].caja = { x: cx - w / 2, y: cy - h / 2, w, h };
    enAro.forEach((q, i) => {
      const a = ((st + (i * span) / Math.max(1, m)) * Math.PI) / 180;
      q.caja = { x: cx + R * Math.sin(a) - w / 2, y: cy - R * Math.cos(a) - h / 2, w, h };
      q.angulo = st + (i * span) / Math.max(1, m);
    });
    // conectores: entre un nodo y el siguiente, sobre el aro
    const conns = hs.filter((q) => q.alg && q.alg.tipo === "conn");
    conns.forEach((q, i) => {
      const a1 = st + (i * span) / Math.max(1, m), a2 = st + ((i + 1) * span) / Math.max(1, m), am = ((a1 + a2) / 2) * Math.PI / 180;
      const tw = (q.v.w != null ? q.v.w : w * 0.3) * k, th = q.v.h != null ? q.v.h * k : tw * 0.6;
      q.caja = { x: cx + R * Math.sin(am) - tw / 2, y: cy - R * Math.cos(am) - th / 2, w: tw, h: th };
      q.arco = { cx, cy, R, a1, a2 };
      q.dir = (a1 + a2) / 2 + 90;
    });
    hs.filter((q) => q.alg && q.alg.tipo === "sp").forEach((q) => { q.caja = { x: cx, y: cy, w: 0, h: 0 }; });
  }
  /** Jerarquía (organigramas): cada nodo arriba y centrado sobre lo que cuelga de él, los hijos en fila debajo. */
  function jerarquia(inst, hs, p, tipo) {
    const C = inst.caja;
    // medir: el ancho que necesita cada hijo (recursivo por sus propios hijos de jerarquía)
    const medida = (q) => {
      if (q._m) return q._m;
      const t = tam(q, inst);
      if (q.alg && (q.alg.tipo === "hierRoot" || q.alg.tipo === "hierChild")) {
        const hj = visibles(q).map((x) => ({ x, m: medida(x) }));
        const raices = hj.filter((o) => !(o.x.alg && (o.x.alg.tipo === "hierChild" || o.x.alg.tipo === "hierRoot"))), ramas = hj.filter((o) => o.x.alg && (o.x.alg.tipo === "hierChild" || o.x.alg.tipo === "hierRoot"));
        const espH = (q.v.sibSp || t.w * 0.1), wr = Math.max(0, ...raices.map((o) => o.m.w)), hr = Math.max(0, ...raices.map((o) => o.m.h));
        const enFilaQ = q.alg.tipo === "hierChild" && !/fromT|fromB/.test((q.alg.p || {}).linDir || "");
        const wRamas = enFilaQ ? ramas.reduce((s, o) => s + o.m.w, 0) + espH * Math.max(0, ramas.length - 1) : Math.max(0, ...ramas.map((o) => o.m.w));
        const hRamas = enFilaQ ? Math.max(0, ...ramas.map((o) => o.m.h)) : ramas.reduce((s, o) => s + o.m.h, 0);
        q._m = { w: Math.max(wr, wRamas), h: hr + (ramas.length ? hr * 0.4 + hRamas : 0) };
      } else q._m = { w: t.w, h: t.h };
      return q._m;
    };
    const m = hs.map(medida);
    // (hierChild: los hijos en fila, o uno debajo de otro si su dirección es de arriba abajo)
    const enFila = tipo === "hierChild" && !/fromT|fromB/.test(p.linDir || "");
    const W = enFila ? m.reduce((s, x) => s + x.w, 0) + (inst.v.sibSp || 0) * Math.max(0, m.length - 1) : Math.max(0, ...m.map((x) => x.w));
    const H = enFila ? Math.max(0, ...m.map((x) => x.h)) : m.reduce((s, x) => s + x.h, 0) + (tipo === "hierChild" ? (inst.v.sibSp || 0) * Math.max(0, m.length - 1) : 0);
    const k = Math.min(W ? C.w / W : 1, H ? C.h / H : 1);
    if (tipo === "hierChild" && !enFila) {
      let y = C.y + (C.h - H * k) / 2;
      hs.forEach((q, i) => { q.caja = { x: C.x, y, w: m[i].w * k, h: m[i].h * k }; escalarArbol(q, k); y += (m[i].h + (inst.v.sibSp || 0)) * k; });
    } else if (tipo === "hierChild") {
      let x = C.x + (C.w - W * k) / 2;
      hs.forEach((q, i) => { q.caja = { x, y: C.y, w: m[i].w * k, h: m[i].h * k }; escalarArbol(q, k); x += (m[i].w + (inst.v.sibSp || 0)) * k; });
    } else {
      let y = C.y;
      hs.forEach((q, i) => {
        const esRama = q.alg && (q.alg.tipo === "hierChild" || q.alg.tipo === "hierRoot");
        const w = m[i].w * k, h = m[i].h * k;
        q.caja = { x: C.x + (C.w - w) / 2, y, w, h };
        escalarArbol(q, k);
        y += h + (esRama ? 0 : h * 0.4);
      });
    }
  }
  function escalarArbol(q, k) { if (k === 1) return; escalarHijo(q, k); q._m = null; q.hijos.forEach((h) => escalarArbol(h, k)); }
  function piramide(inst, hs, p) {
    const C = inst.caja, niveles = hs.filter((h) => !(h.alg && h.alg.tipo === "sp")), n = niveles.length;
    if (!n) return;
    const alto = C.h / n, abajo = (p.linDir || "fromB") === "fromT";
    niveles.forEach((q, i) => {
      const j = abajo ? i : i;
      const w = C.w * ((j + 1) / n);
      q.caja = { x: C.x + (C.w - w) / 2, y: C.y + i * alto, w, h: alto };
      q.piso = { i, n };
    });
  }

  // ------------------------------------------------------------ letra: el tamaño más grande con que el texto cabe
  function textoDe(inst) {
    const out = [];
    inst.presOf.forEach((p, k) => {
      const nivel = k === 0 ? 0 : 1 + Math.max(0, profundidad(p) - profundidad(inst.presOf[0]) - 1);
      (p.texto.length ? p.texto : [""]).forEach((t) => out.push({ t, nivel }));
    });
    return out.filter((x, i) => x.t !== "" || i === 0);
  }
  function cabe(lineas, w, h, pt, medir) {
    const EMU_PT = 12700, alto = pt * 1.2 * 0.9 * EMU_PT;
    let n = 0;
    for (const ln of lineas) {
      const ancho = w - (ln.nivel ? pt * 1.5 * EMU_PT : 0);
      if (!ln.t) { n++; continue; }
      let renglon = "";
      for (const pal of ln.t.split(/\s+/)) {
        if (medir(pal, pt * (ln.nivel ? 0.8 : 1)) > ancho) return false;   // una palabra no cabe ni sola
        const prueba = renglon ? renglon + " " + pal : pal;
        if (medir(prueba, pt * (ln.nivel ? 0.8 : 1)) > ancho && renglon) { n++; renglon = pal; } else renglon = prueba;
      }
      n++;
    }
    return n * alto <= h + 1;
  }

  // ------------------------------------------------------------ el resultado: las formas en orden de dibujo
  function acomodar(partes, opc) {
    const datos = leerDatos(partes.datos);
    const raizDis = partes.diseno && [...partes.diseno.getElementsByTagNameNS(DGM, "layoutNode")][0];
    if (!datos.raiz || !raizDis) return null;
    const raiz = instanciar(raizDis, datos.raiz, {}, { n: 0 });
    raiz.caja = { x: 0, y: 0, w: opc.w, h: opc.h };
    colocar(raiz);
    // colores y estilo de cada etiqueta (colorsDef y styleDef)
    const etiquetas = (doc) => { const m = new Map(); if (doc) for (const s of doc.getElementsByTagNameNS(DGM, "styleLbl")) m.set(at(s, "name"), s); return m; };
    const colores = etiquetas(partes.colores), estilos = etiquetas(partes.estilo);
    const colorDe = (lista, i, cnt) => {
      if (!lista) return null;
      const cs = hijos(lista);
      if (!cs.length) return null;
      const met = at(lista, "meth") || "repeat";
      const el = met === "cycle" ? cs[i % cs.length] : met === "span" && cs.length > 1 ? cs[Math.round((Math.min(i, Math.max(1, cnt) - 1) / Math.max(1, cnt - 1)) * (cs.length - 1))] : cs[Math.min(i, cs.length - 1)];
      return opc.color(el);
    };
    const formas = [];
    const recorrer = (q) => {
      const f = q.forma, alg = q.alg ? q.alg.tipo : "composite";
      // (se dibuja su forma si tiene y no está escondida; el texto de un nodo de texto se ve aunque su forma no)
      const conForma = f && f.tipo !== "none" && !f.oculta, conTexto = alg === "tx" && textoDe(q).some((x) => x.t);
      if (f && (conForma || conTexto) && q.caja && q.caja.w > 0 && q.caja.h > 0) {
        // su punto de presentación (dice la etiqueta de estilo y su número entre los de su clase)
        const ps = (q.ctx && datos.presDe.get(q.ctx.id)) || [];
        const pr = ps.find((x) => x.nombre === q.nombre) || null;
        const lbl = (pr && pr.lbl) || q.lbl || (esNodo(q.ctx) ? "node1" : q.ctx && q.ctx.tipo === "sibTrans" ? "sibTrans2D1" : "node1");
        const idx = pr ? pr.idx : 0, cnt = pr ? pr.cnt : 1;
        const cl = colores.get(lbl), es = estilos.get(lbl), st = es && hijo(es, "style");
        const fillRef = st && hijo(st, "fillRef"), lnRef = st && hijo(st, "lnRef");
        const conRelleno = !fillRef || +at(fillRef, "idx") !== 0, conLinea = lnRef && +at(lnRef, "idx") !== 0;
        let prst = f.tipo;
        const forma = { x: q.caja.x, y: q.caja.y, w: q.caja.w, h: q.caja.h, rot: f.rot || 0, prst, adj: f.adj, nombre: q.nombre,
          relleno: conForma && conRelleno ? colorDe(cl && hijo(cl, "fillClrLst"), idx, cnt) : null,
          linea: conForma && conLinea ? { w: opc.anchoLinea ? opc.anchoLinea(+at(lnRef, "idx")) : 12700, c: colorDe(cl && hijo(cl, "linClrLst"), idx, cnt) } : null,
          colorTexto: colorDe(cl && hijo(cl, "txFillClrLst"), idx, cnt) };
        // conectores: flecha (2D) en la dirección de la línea, o arco sobre el aro del ciclo (1D)
        if (prst === "conn") {
          const dim = (q.alg && q.alg.p && q.alg.p.dim) || "2D";
          if (q.arco && dim === "1D") { const a = q.arco; Object.assign(forma, { prst: "arc", x: a.cx - a.R, y: a.cy - a.R, w: 2 * a.R, h: 2 * a.R, rot: 0, adj: { 1: ((a.a1 + 6 - 90) % 360 + 360) % 360, 2: ((a.a2 - 6 - 90) % 360 + 360) % 360 }, arcoGrados: true, relleno: null, linea: forma.linea || { w: 12700, c: forma.relleno } }); }
          else { forma.prst = "rightArrow"; forma.rot = q.dir || 0; forma.adj = { 1: 0.6, 2: 0.5 }; }
        }
        if (q.piso && prst === "trapezoid") forma.adj = { 1: q.caja.w ? (q.caja.w / (2 * (q.piso.i + 1))) / Math.min(q.caja.w, q.caja.h) : 0.5 };
        // texto (nodos de texto): con el tamaño de letra que luego se iguala entre los de su nombre
        if (alg === "tx") {
          const lineas = textoDe(q);
          if (lineas.some((x) => x.t)) {
            forma.texto = lineas;
            forma.margen = ["lMarg", "tMarg", "rMarg", "bMarg"].map((k) => (q.v[k] != null ? q.v[k] : 0.3 * (q.v.primFontSz || 18)) * 12700 / 1);
            forma.tamMax = q.v.primFontSz || 65;
            forma.tamMin = Math.max(5, ...q.reglas.filter((r) => r.type === "primFontSz").map((r) => num(r.val, 5)).filter((x) => isFinite(x)));
            forma.alin = (q.alg.p.parTxLTRAlign || "ctr");
            forma.ancla = q.alg.p.txAnchorVert || "mid";
          }
        }
        formas.push(Object.assign(forma, { z: f.zOff || 0 }));
      }
      q.hijos.forEach(recorrer);
    };
    recorrer(raiz);
    // letra: el mismo tamaño para todas las de un mismo nombre, el más grande con que todas caben (como PowerPoint)
    const grupos = new Map();
    formas.filter((f) => f.texto).forEach((f) => (grupos.get(f.nombre) || grupos.set(f.nombre, []).get(f.nombre)).push(f));
    grupos.forEach((fs) => {
      let pt = Math.min(...fs.map((f) => f.tamMax)), min = Math.max(...fs.map((f) => f.tamMin));
      const ok = (x) => fs.every((f) => { const m = f.margen.map((v) => (v / (f.tamMax || 1)) * x); return cabe(f.texto, f.w - m[0] - m[2], f.h - m[1] - m[3], x, opc.medir); });
      while (pt > min && !ok(pt)) pt -= pt > 20 ? 1 : 0.5;
      fs.forEach((f) => { f.tam = pt; f.margen = f.margen.map((v) => (v / (f.tamMax || 1)) * pt); });
    });
    return formas.sort((a, b) => a.z - b.z);
  }

  NOM.smartart = { acomodar };
})();

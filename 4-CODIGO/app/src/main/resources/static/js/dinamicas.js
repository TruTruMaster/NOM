/* NOM · Tablas dinámicas (Insertar → Tabla dinámica), como Excel.

   Una tabla dinámica de una hoja (t.pv, una lista) dice de dónde salen los datos (un rango con encabezados, de esta hoja
   o de otra), dónde va, qué campos van en filas, en columnas y en el filtro, y qué valores resume (suma, cuenta,
   promedio, máximo o mínimo). calcular() la arma como Excel en el diseño compacto (el de siempre): los elementos de cada
   campo ordenados de la A a la Z, los subtotales arriba (en el renglón del elemento de afuera), el total general abajo y a
   la derecha, el filtro dos renglones arriba. Devuelve lo que se escribe en cada celda y las partes de Excel
   (pivotTableDefinition, pivotCacheDefinition y pivotCacheRecords) para guardarla como tabla dinámica de verdad; el estilo
   se pinta con la misma lógica que las tablas dinámicas que vienen en un archivo de Excel (ooxml.js).

   Como en Excel, lo escrito en las celdas se queda hasta «Actualizar» (si cambian los datos de origen). */
(function () {
  "use strict";
  const NOM = (window.NOM = window.NOM || {});
  const F = () => NOM.formulas;

  const FN = { sum: "Suma", count: "Cuenta", average: "Promedio", max: "Máx.", min: "Mín." };
  const TXT = { filas: "Etiquetas de fila", columnas: "Etiquetas de columna", total: "Total general", todas: "(Todas)", valores: "Valores" };
  const x = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

  /** El nombre de un valor resumido, como Excel en español: «Suma de Ventas», «Cuenta de Unidades»… */
  const nombreValor = (fn, campo) => `${FN[fn] || FN.sum} de ${campo}`;
  const esNum = (v) => typeof v === "number" && isFinite(v);
  /** Clave de un elemento (Excel junta «Norte» y «norte» en el mismo). */
  const clave = (v) => (v == null || v === "" ? "\u0000(en blanco)" : esNum(v) ? "n:" + v : "s:" + String(v).toLocaleLowerCase("es"));
  /** Orden de Excel para los elementos: números de menor a mayor, luego el texto de la A a la Z (sin distinguir acentos ni mayúsculas). */
  function comparar(a, b) {
    const na = esNum(a), nb = esNum(b);
    if (na && nb) return a - b;
    if (na !== nb) return na ? -1 : 1;
    return String(a).localeCompare(String(b), "es", { sensitivity: "base" }) || (String(a) < String(b) ? -1 : String(a) > String(b) ? 1 : 0);
  }
  function resumir(fn, vs) {
    const ns = vs.filter(esNum);
    if (fn === "count") return vs.filter((v) => v != null && v !== "").length;
    if (!ns.length) return fn === "sum" ? 0 : null;
    if (fn === "average") return ns.reduce((a, b) => a + b, 0) / ns.length;
    if (fn === "max") return Math.max(...ns);
    if (fn === "min") return Math.min(...ns);
    return Math.round(ns.reduce((a, b) => a + b, 0) * 1e10) / 1e10;
  }

  /**
   * Arma la tabla dinámica. def: { n, campos: [nombres], datos: [[valor…]…] (sin el encabezado), filas: [i…], columnas:
   * [i] (a lo más uno), valores: [{ c: i, fn }], filtro: { c: i, sel: valor | null } | null, destino: { c, f }, estilo }.
   * Devuelve { celdas: [[c, f, valor]…], rango (sin el filtro), rango2 (con el filtro), capas: {…}, partes: {…} } o un motivo.
   */
  function calcular(def) {
    const campos = def.campos || [], filasC = (def.filas || []).slice(0, 3), colsC = (def.columnas || []).slice(0, 1), vals = def.valores || [];
    if (!filasC.length) return "Pon al menos un campo en Filas";
    if (!vals.length) return "Pon al menos un campo en Valores";
    if (colsC.length && vals.length > 1) return "Con un campo en Columnas, un solo valor (por ahora)";
    const fil = def.filtro && def.filtro.c != null ? def.filtro : null;
    const todos = def.datos || [];
    const datos = fil && fil.sel != null ? todos.filter((r) => clave(r[fil.c]) === clave(fil.sel)) : todos;
    // los elementos de cada campo que se usa: en el orden en que aparecen (sharedItems) y ordenados (items)
    const usados = [...new Set([...filasC, ...colsC, ...(fil ? [fil.c] : [])])];
    const elem = {};
    usados.forEach((ci) => {
      const vistos = new Map();
      todos.forEach((r) => { const k = clave(r[ci]); if (!vistos.has(k)) vistos.set(k, { v: r[ci] == null ? "" : r[ci], x: vistos.size }); });
      const lista = [...vistos.values()];
      elem[ci] = { lista, orden: lista.slice().sort((a, b) => comparar(a.v, b.v)), pos: new Map(lista.map((e) => [clave(e.v), e.x])) };
    });
    const ordIdx = (ci) => new Map(elem[ci].orden.map((e, i) => [clave(e.v), i]));
    const ord = Object.fromEntries(usados.map((ci) => [ci, ordIdx(ci)]));
    // el árbol de filas: [{ ruta: [clave…], nivel, etiqueta, filas: [registros] }…] en el orden de Excel
    const filasL = [];
    const bajar = (regs, nivel, ruta) => {
      const ci = filasC[nivel], grupos = new Map();
      regs.forEach((r) => { const k = clave(r[ci]); if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(r); });
      [...grupos.keys()].sort((a, b) => ord[ci].get(a) - ord[ci].get(b)).forEach((k) => {
        const rs = grupos.get(k), v = elem[ci].orden[ord[ci].get(k)].v;
        filasL.push({ nivel, ruta: ruta.concat(k), etiqueta: v === "" ? "(en blanco)" : v, regs: rs, ultimo: nivel === filasC.length - 1 });
        if (nivel < filasC.length - 1) bajar(rs, nivel + 1, ruta.concat(k));
      });
    };
    bajar(datos, 0, []);
    // columnas: con un campo, sus elementos (con datos) y el total; si no, una por valor
    const colsL = colsC.length ? [...new Set(datos.map((r) => clave(r[colsC[0]])))].sort((a, b) => ord[colsC[0]].get(a) - ord[colsC[0]].get(b)) : null;
    const d0 = def.destino || { c: 0, f: 0 }, c0 = d0.c, f0 = d0.f + (fil ? 2 : 0);
    const celdas = [];
    const pon = (c, f, v) => celdas.push([c, f, v]);
    let fila = f0;
    const nV = vals.length, nombres = vals.map((v) => v.n || nombreValor(v.fn, campos[v.c]));
    if (fil) { pon(c0, d0.f, campos[fil.c]); pon(c0 + 1, d0.f, fil.sel == null ? TXT.todas : fil.sel); }
    if (colsL) {
      pon(c0, fila, nombres[0]); pon(c0 + 1, fila, TXT.columnas); fila++;
      pon(c0, fila, TXT.filas);
      colsL.forEach((k, i) => { const v = elem[colsC[0]].orden[ord[colsC[0]].get(k)].v; pon(c0 + 1 + i, fila, v === "" ? "(en blanco)" : v); });
      pon(c0 + 1 + colsL.length, fila, TXT.total); fila++;
    } else {
      pon(c0, fila, TXT.filas);
      nombres.forEach((n, i) => pon(c0 + 1 + i, fila, n)); fila++;
    }
    const fData = fila, sangrias = [];
    const valorDe = (regs, vi) => resumir(vals[vi].fn, regs.map((r) => r[vals[vi].c]));
    filasL.forEach((L) => {
      pon(c0, fila, L.etiqueta);
      if (L.nivel) sangrias.push([c0, fila, L.nivel]);
      if (colsL) {
        colsL.forEach((k, i) => { const rs = L.regs.filter((r) => clave(r[colsC[0]]) === k); if (rs.length) pon(c0 + 1 + i, fila, valorDe(rs, 0)); });
        pon(c0 + 1 + colsL.length, fila, valorDe(L.regs, 0));
      } else vals.forEach((_, vi) => pon(c0 + 1 + vi, fila, valorDe(L.regs, vi)));
      fila++;
    });
    pon(c0, fila, TXT.total);
    if (colsL) { colsL.forEach((k, i) => pon(c0 + 1 + i, fila, valorDe(datos.filter((r) => clave(r[colsC[0]]) === k), 0))); pon(c0 + 1 + colsL.length, fila, valorDe(datos, 0)); }
    else vals.forEach((_, vi) => pon(c0 + 1 + vi, fila, valorDe(datos, vi)));
    const c2 = c0 + (colsL ? colsL.length + 1 : nV), f2 = fila;
    const rango = { c1: c0, f1: f0, c2, f2 }, rango2 = { c1: c0, f1: fil ? d0.f : f0, c2: fil ? Math.max(c2, c0 + 1) : c2, f2 };
    const partes = xmlPartes(def, { campos, filasC, colsC, vals, nombres, fil, elem, ord, filasL, colsL, rango, fData, todos });
    return { celdas, rango, rango2, sangrias, partes, nombres };
  }

  // ------------------------------------------------------------ las partes de Excel
  function xmlPartes(def, P) {
    const { campos, filasC, colsC, vals, nombres, fil, elem, ord, filasL, colsL, rango, todos } = P;
    const ref = (c, f) => F().ref(c, f);
    const usados = new Set(Object.keys(elem).map(Number));
    // la caché: cada campo con sus elementos (los de los ejes) o el tipo de sus números; y los registros
    const cacheFields = campos.map((n, ci) => {
      if (usados.has(ci)) {
        const l = elem[ci].lista, nums = l.every((e) => esNum(e.v));
        return `<cacheField name="${x(n)}" numFmtId="0"><sharedItems${nums ? ' containsSemiMixedTypes="0" containsString="0" containsNumber="1"' + (l.every((e) => Number.isInteger(e.v)) ? ' containsInteger="1"' : "") + ` minValue="${Math.min(...l.map((e) => e.v))}" maxValue="${Math.max(...l.map((e) => e.v))}"` : ""} count="${l.length}">${l.map((e) => (esNum(e.v) ? `<n v="${e.v}"/>` : e.v === "" ? "<m/>" : `<s v="${x(e.v)}"/>`)).join("")}</sharedItems></cacheField>`;
      }
      const vs = todos.map((r) => r[ci]), ns = vs.filter(esNum);
      if (vs.length && ns.length === vs.length) return `<cacheField name="${x(n)}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"${ns.every(Number.isInteger) ? ' containsInteger="1"' : ""} minValue="${Math.min(...ns)}" maxValue="${Math.max(...ns)}"/></cacheField>`;
      return `<cacheField name="${x(n)}" numFmtId="0"><sharedItems${vs.some((v) => v == null || v === "") ? ' containsBlank="1"' : ""}/></cacheField>`;
    }).join("");
    const o = def.origen || {};
    const cache = XML + `<pivotCacheDefinition ${NS} r:id="rId1" refreshOnLoad="1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="${todos.length}">`
      + `<cacheSource type="worksheet"><worksheetSource ref="${ref(o.c1, o.f1)}:${ref(o.c2, o.f2)}" sheet="${x(o.hoja || "")}"/></cacheSource><cacheFields count="${campos.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`;
    const registros = XML + `<pivotCacheRecords ${NS} count="${todos.length}">${todos.map((r) => `<r>${campos.map((_, ci) => {
      const v = r[ci];
      if (usados.has(ci)) return `<x v="${elem[ci].pos.get(clave(v))}"/>`;
      return esNum(v) ? `<n v="${v}"/>` : v == null || v === "" ? "<m/>" : `<s v="${x(v)}"/>`;
    }).join("")}</r>`).join("")}</pivotCacheRecords>`;
    // la definición: campos (con su eje y sus elementos ordenados), elementos de filas y columnas, valores y estilo
    const eje = (ci) => (filasC.includes(ci) ? "axisRow" : colsC.includes(ci) ? "axisCol" : fil && fil.c === ci ? "axisPage" : null);
    const pivotFields = campos.map((_, ci) => {
      const a = eje(ci), dato = vals.some((v) => v.c === ci) ? ' dataField="1"' : "";
      if (!a) return `<pivotField${dato} showAll="0"/>`;
      const it = elem[ci].orden.map((e) => `<item x="${e.x}"/>`).join("") + '<item t="default"/>';
      return `<pivotField axis="${a}"${dato} showAll="0"><items count="${elem[ci].orden.length + 1}">${it}</items></pivotField>`;
    }).join("");
    // rowItems: cada renglón con el índice (en el orden) de su elemento; r = cuántos niveles de afuera repite
    const rowItems = filasL.map((L) => { const ci = filasC[L.nivel], v = ord[ci].get(L.ruta[L.nivel]); return `<i${L.nivel ? ` r="${L.nivel}"` : ""}>${v ? `<x v="${v}"/>` : "<x/>"}</i>`; }).join("") + '<i t="grand"><x/></i>';
    const nFil = filasL.length + 1;
    let colXml, colItems, fhr;
    if (colsL) {
      const ci = colsC[0];
      colXml = `<colFields count="1"><field x="${ci}"/></colFields>`;
      colItems = `<colItems count="${colsL.length + 1}">${colsL.map((k) => { const v = ord[ci].get(k); return `<i>${v ? `<x v="${v}"/>` : "<x/>"}</i>`; }).join("")}<i t="grand"><x/></i></colItems>`;
      fhr = 1;
    } else if (vals.length > 1) {
      colXml = `<colFields count="1"><field x="-2"/></colFields>`;
      colItems = `<colItems count="${vals.length}">${vals.map((_, i) => (i ? `<i i="${i}"><x v="${i}"/></i>` : "<i><x/></i>")).join("")}</colItems>`;
      fhr = 0;
    } else { colXml = ""; colItems = '<colItems count="1"><i/></colItems>'; fhr = 1; }
    const fdr = colsL ? 2 : 1;
    const loc = `<location ref="${ref(rango.c1, rango.f1)}:${ref(rango.c2, rango.f2)}" firstHeaderRow="${fhr}" firstDataRow="${fdr}" firstDataCol="1"${fil ? ' rowPageCount="1" colPageCount="1"' : ""}/>`;
    const pageXml = fil ? `<pageFields count="1"><pageField fld="${fil.c}"${fil.sel != null ? ` item="${ord[fil.c].get(clave(fil.sel))}"` : ""} hier="-1"/></pageFields>` : "";
    const dataXml = `<dataFields count="${vals.length}">${vals.map((v, i) => `<dataField name="${x(nombres[i])}" fld="${v.c}"${v.fn && v.fn !== "sum" ? ` subtotal="${v.fn}"` : ""} baseField="0" baseItem="0"/>`).join("")}</dataFields>`;
    const tabla = XML + `<pivotTableDefinition ${NS} name="${x(def.n || "TablaDinámica1")}" cacheId="\u0001CACHE\u0001" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="${TXT.valores}" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">`
      + loc + `<pivotFields count="${campos.length}">${pivotFields}</pivotFields><rowFields count="${filasC.length}">${filasC.map((ci) => `<field x="${ci}"/>`).join("")}</rowFields><rowItems count="${nFil}">${rowItems}</rowItems>`
      + colXml + colItems + pageXml + dataXml
      + `<pivotTableStyleInfo name="${x(def.estilo || "PivotStyleLight16")}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
    return { tabla, cache, registros };
  }

  /** Los datos de origen de una tabla dinámica de la hoja `t`: los nombres de los campos (la primera fila del rango) y los
      registros con sus valores ya calculados; o el motivo si no se puede. */
  function datos(doc, t, o) {
    const hoja = o.hoja ? F().hojasDe(doc).find((h) => F().nombreDeHoja(doc, h).toLowerCase() === String(o.hoja).toLowerCase()) : t;
    if (!hoja) return "No está la hoja de los datos";
    const ev = F().evaluadorDeTabla(doc, hoja);
    const val = (c, f) => { const v = ev.valor(c, f); if (v == null || v === "") return ""; if (typeof v === "number" || typeof v === "boolean") return v; if (v instanceof F().ErrorF) return String(v); const n = +v; return typeof v === "object" && !isNaN(n) ? n : String(v); };
    const campos = [];
    for (let c = o.c1; c <= o.c2; c++) { const n = String(val(c, o.f1)).trim(); if (!n) return "Cada columna de los datos necesita un nombre en la primera fila"; campos.push(n); }
    const regs = [];
    for (let f = o.f1 + 1; f <= o.f2; f++) { const r = []; for (let c = o.c1; c <= o.c2; c++) r.push(val(c, f)); if (r.some((v) => v !== "")) regs.push(r); }
    if (!regs.length) return "Los datos necesitan al menos una fila debajo de los nombres";
    return { campos, datos: regs, nombreHoja: F().nombreDeHoja(doc, hoja) };
  }
  /** La tabla dinámica `pv` de la hoja `t` armada con sus datos de ahora (o el motivo). */
  function armar(doc, t, pv) {
    const o = datos(doc, t, pv.origen);
    if (typeof o === "string") return o;
    return calcular(Object.assign({}, pv, { campos: o.campos, datos: o.datos, origen: Object.assign({}, pv.origen, { hoja: o.nombreHoja }) }));
  }

  NOM.dinamicas = { calcular, datos, armar, nombreValor, FN, TXT, comparar };
})();

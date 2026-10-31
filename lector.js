/* Lector de facturas de Belleza Lua.
   Extrae datos de facturas en PDF con texto, sin IA y sin coste:
   - lineasPdf(): reconstruye las líneas de texto del PDF con pdf.js
   - leerFactura(): proveedor, NIF, nº, fecha, base, IVA, retención y total
   - lineasGdC(): líneas de producto de las facturas de Germaine de Capuccini
   Funciona igual en el navegador y en Node (pruebas). */
(function (root) {
  "use strict";

  /* Proveedores conocidos (empresas; datos públicos). Los demás se aprenden al guardar. */
  const PROVEEDORES = {
    A03050572: { nombre: "Germaine de Capuccini", cat: "CREMAS" },
    B45441953: { nombre: "Yodeyma", cat: "COLONIAS" },
    A82018474: { nombre: "Telefónica (O2)", cat: "TELÉFONO" },
    A65067332: { nombre: "Comercializadora Regulada (luz)", cat: "LUZ" },
    IE9813461A: { nombre: "SumUp", cat: "SUM UP" },
    B19937309: { nombre: "Depilshop", cat: "MATERIAL TRABAJO" },
    A28581882: { nombre: "Alcampo", cat: "MATERIAL TRABAJO" }
  };
  const PISTAS = [
    [/GERMAINE|8412971\d{6}\s+\d{8}/i, "A03050572"],
    [/YODEYMA/i, "B45441953"],
    [/TELEF[OÓ]NICA|\bO2\b/i, "A82018474"],
    [/COMERCIALIZADORA REGULADA|NATURGY/i, "A65067332"],
    [/SUMUP|monthly-fee|comisiones por/i, "IE9813461A"],
    [/DEPILSHOP/i, "B19937309"],
    [/ALCAMPO/i, "A28581882"]
  ];

  const MESES = { enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8, septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, diciembre: 12 };

  /* "1.111,93" -> 1111.93 ; "€2.78" -> 2.78 ; "63,6€" -> 63.6 ; "***782,52" -> 782.52 */
  function numero(s) {
    if (s == null) return null;
    let t = String(s).replace(/[€*\s]/g, "").replace(/EUR/i, "");
    const neg = /^-/.test(t) || /^\(.*\)$/.test(t);
    t = t.replace(/[^\d.,]/g, "");
    if (!t) return null;
    if (t.includes(",")) t = t.replace(/\./g, "").replace(",", ".");
    else if (/\.\d{3}$/.test(t) && !/\.\d{1,2}$/.test(t)) t = t.replace(/\./g, "");
    const n = parseFloat(t);
    if (!isFinite(n)) return null;
    return Math.round((neg ? -n : n) * 100) / 100;
  }

  const RE_IMPORTE = /-?(?:€\s?)?\d{1,3}(?:\.\d{3})*,\d{1,2}(?!\d)|-?€\s?\d+\.\d{2}(?!\d)|-?\d+,\d{1,2}(?!\d)/g;
  function importes(linea) {
    const m = linea.match(RE_IMPORTE) || [];
    return m.map(numero).filter((n) => n != null);
  }

  function isoFecha(d, m, y) {
    d = +d; m = +m; y = +y;
    if (y < 100) y += 2000;
    if (!(d >= 1 && d <= 31 && m >= 1 && m <= 12 && y >= 2000 && y <= 2100)) return null;
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  function fechas(linea) {
    const out = [];
    let m;
    const r1 = /(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/g;
    while ((m = r1.exec(linea))) { const f = isoFecha(m[1], m[2], m[3]); if (f) out.push(f); }
    const r2 = /(\d{1,2})\s+de\s+([a-záéíóú]+)\s+(?:de|del)\s+(\d{4})/gi;
    while ((m = r2.exec(linea))) { const mm = MESES[m[2].toLowerCase()]; if (mm) { const f = isoFecha(m[1], mm, m[3]); if (f) out.push(f); } }
    return out;
  }

  const RE_NIF = /\b(?:ES)?([ABCDEFGHJNPQRSUVW]-?\d{7}[0-9A-J]|\d{8}-?[A-Z]|[XYZ]\d{7}[A-Z]|IE\d{7}[A-Z]{1,2})\b/g;
  function nifs(texto) {
    const out = [];
    let m;
    while ((m = RE_NIF.exec(texto))) { const n = m[1].replace(/-/g, ""); if (!out.includes(n)) out.push(n); }
    return out;
  }

  /* Reconstruye líneas a partir de los fragmentos de texto de pdf.js */
  async function lineasPdf(pdfjsLib, datos) {
    const pdf = await pdfjsLib.getDocument({ data: datos, isEvalSupported: false }).promise;
    const lineas = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
      const filas = [];
      tc.items.forEach((it) => {
        if (!it.str || !it.str.trim()) return;
        const y = it.transform[5], x = it.transform[4];
        if (Math.abs(it.transform[1]) > 0.5 && Math.abs(it.transform[0]) < 0.5) return; // texto vertical (márgenes legales)
        let f = filas.find((r) => Math.abs(r.y - y) < 2.5);
        if (!f) { f = { y, items: [] }; filas.push(f); }
        f.items.push({ x, w: it.width || 0, s: it.str });
      });
      filas.sort((a, b) => b.y - a.y);
      filas.forEach((f) => {
        f.items.sort((a, b) => a.x - b.x);
        let s = "", fin = null;
        f.items.forEach((it) => {
          if (fin != null) { const gap = it.x - fin; s += gap > 12 ? "   " : gap > 1.5 ? " " : ""; }
          s += it.s; fin = it.x + it.w;
        });
        lineas.push(s.replace(/\s+$/, ""));
      });
    }
    return lineas;
  }

  /* Busca base y cuota de IVA que cuadren con el total */
  function cuadrar(nums, total) {
    const c = [...new Set(nums.filter((n) => n > 0 && n < total + 0.01))];
    const tol = (a, b) => Math.abs(a - b) <= 0.021;
    for (const rate of [0.21, 0.10, 0.04]) {
      for (const b of c) {
        const cuota = Math.round(b * rate * 100) / 100;
        if (c.some((x) => tol(x, cuota)) && tol(b + cuota, total)) return { base: b, tipo: rate * 100, cuota, ret: 0 };
      }
    }
    for (const rate of [0.21, 0.10]) { // con retención IRPF (alquiler)
      for (const b of c) {
        const cuota = Math.round(b * rate * 100) / 100;
        for (const rr of [0.19, 0.15, 0.07]) {
          const ret = Math.round(b * rr * 100) / 100;
          if (c.some((x) => tol(x, cuota)) && c.some((x) => tol(x, ret)) && tol(b + cuota - ret, total)) return { base: b, tipo: rate * 100, cuota, ret };
        }
      }
    }
    for (const b of c) { // tipos mixtos (p. ej. 10 % + 21 %)
      for (const q of c) {
        if (q !== b && q >= b * 0.035 && q <= b * 0.215 && tol(b + q, total)) return { base: b, tipo: Math.round((q / b) * 1000) / 10, cuota: q, ret: 0, mixto: true };
      }
    }
    return null;
  }

  function leerFactura(lineas, opciones) {
    const propio = ((opciones && opciones.nifPropio) || "").toUpperCase().replace(/[^0-9A-Z]/g, "");
    const conocidos = Object.assign({}, PROVEEDORES, (opciones && opciones.aprendidos) || {});
    const texto = lineas.join("\n");
    const r = { proveedor: "", nif: "", numero: "", fecha: "", base: null, tipoIva: null, iva: null, retencion: 0, total: null, categoria: "", avisos: [] };

    // Proveedor
    const lista = nifs(texto).filter((n) => n !== propio && n !== "ES" + propio);
    let nif = lista.find((n) => conocidos[n]);
    if (!nif) { const p = PISTAS.find(([re]) => re.test(texto)); if (p) nif = p[1]; }
    if (!nif) nif = lista[0] || "";
    r.nif = nif;
    if (conocidos[nif]) { r.proveedor = conocidos[nif].nombre; r.categoria = conocidos[nif].cat; }

    // Total
    const etiquetas = [/total\s+a\s+pagar/i, /importe\s+factura/i, /total\s+factura/i, /total\s+importe\s+factura/i, /^\s*total\b(?!\s*(bruto|base|impuesto|deducc|de los))/i, /\btotal\b(?!\s*(bruto|base|impuesto|deducc|de los|ejercicio|con iva))/i];
    for (const re of etiquetas) {
      for (let i = 0; i < lineas.length && r.total == null; i++) {
        if (!re.test(lineas[i]) || /subtotal|sub-total/i.test(lineas[i])) continue;
        const after = lineas[i].slice(lineas[i].search(re));
        let v = importes(after);
        if (!v.length && i + 1 < lineas.length) v = importes(lineas[i + 1]);
        if (!v.length && i + 2 < lineas.length) v = importes(lineas[i + 2]);
        if (v.length) { const x = v[v.length - 1]; if (x > 0) r.total = x; }
      }
      if (r.total != null) break;
    }

    // Base, IVA y retención
    const todos = [];
    lineas.forEach((l) => importes(l).forEach((n) => todos.push(n)));
    if (r.total != null) {
      const exento = /exent[ao]s?\s+de\s+IVA|IVA\s*0\s*%/i.test(texto);
      const q = exento ? null : cuadrar(todos, r.total);
      if (exento) { r.base = r.total; r.tipoIva = 0; r.iva = 0; r.avisos.push("Exento de IVA"); }
      else if (q) { r.base = q.base; r.tipoIva = q.tipo; r.iva = q.cuota; r.retencion = q.ret; if (q.mixto) r.avisos.push("Varios tipos de IVA"); }
      else r.avisos.push("No se encontró la base imponible");
    } else r.avisos.push("No se encontró el total");

    // Fecha: la que sigue a una etiqueta de fecha; si no, la primera
    for (let i = 0; i < lineas.length && !r.fecha; i++) {
      if (/fecha|emitida|expedici/i.test(lineas[i]) && !/vencim|cargo|l[ií]mite|cobro|nacim/i.test(lineas[i].replace(/fecha de expedici[oó]n.*vencimiento/i, "fecha de expedicion"))) {
        const f = fechas(lineas[i]).concat(fechas(lineas[i + 1] || ""));
        if (f.length) r.fecha = f[0];
      }
    }
    if (!r.fecha) {
      const hoy = (opciones && opciones.hoy) ? new Date(opciones.hoy) : new Date();
      const ok = (f) => { const d = (new Date(f) - hoy) / 864e5; return d > -400 && d < 45; };
      const cuenta = {};
      fechas(texto).filter(ok).forEach((f) => { cuenta[f] = (cuenta[f] || 0) + 1; });
      const mejor = Object.keys(cuenta).sort((a, b) => cuenta[b] - cuenta[a])[0];
      if (mejor) r.fecha = mejor;
    }

    // Número de factura
    const reNum = [
      /factura\s*n[º°o.]*\s*[:.]?\s*([A-Z0-9][\w\-\/]*\d[\w\-\/]*)/i,
      /n[º°o]\.?\s*(?:de\s+)?factura\s*(?:n[º°o.]*)?\s*[:.]?\s*([A-Z0-9][\w\-\/ ]*\d[\w\-\/]*)/i,
      /factura\s+n[uú]m\.?\s*:?\s*([A-Z0-9][\w\-\/]*\d[\w\-\/]*)/i,
      /^\s*factura\s*:?\s+([A-Z0-9][\w\-\/]*\d[\w\-\/]*)/im
    ];
    for (const re of reNum) {
      const m = texto.match(re);
      if (m) { r.numero = m[1].trim().replace(/\s+-\s+/, "-").replace(/\s{2,}.*$/, ""); break; }
    }
    if (!r.numero) {
      for (let i = 0; i < lineas.length; i++) {
        if (/\bfactura\b/i.test(lineas[i]) && !/total|importe|electr|datos|resumen/i.test(lineas[i])) {
          const cab = lineas[i].trim().split(/\s{2,}/), sig = (lineas[i + 1] || "").trim().split(/\s{2,}/);
          const k = cab.findIndex((c) => /\bfactura\b/i.test(c));
          const okNum = (t) => /^(\d{6,}|[A-Z]{1,4}\d{5,}|\d{2,4}[-\/]\d{2,6}(?:[-\/]\d+)?)$/.test(t || "");
          if (cab.length > 1 && cab.length === sig.length && okNum(sig[k])) { r.numero = sig[k]; break; }
          const cands = (lineas[i + 1] || "").match(/\b(\d{6,}|[A-Z]{1,4}\d{5,}|\d{2,4}[-\/]\d{2,6}(?:[-\/]\d+)?)\b/g);
          if (cands) { r.numero = cands[cands.length - 1]; break; }
        }
      }
    }
    return r;
  }

  /* Líneas de producto de una factura de Germaine de Capuccini */
  function lineasGdC(lineas) {
    const out = [];
    const re = /^\s*(\d{13})\s+(\d{8})\s+(.+?)\s+(-?\d+)\s+(?:(-?\d+)\s+)?(-?\d{1,3}(?:\.\d{3})*,\d{2})\s+(?:(\d{1,2}(?:,\d+)?)\s+)?(-?\d{1,3}(?:\.\d{3})*,\d{2})\s*$/;
    lineas.forEach((l) => {
      const m = l.match(re);
      if (!m) return;
      const precio = numero(m[6]), dto = m[7] ? numero(m[7]) : 0, imp = numero(m[8]);
      let uvta, bnf;
      if (m[5] != null) { uvta = +m[4]; bnf = +m[5]; }
      else if (imp !== 0) { uvta = +m[4]; bnf = 0; }
      else { uvta = 0; bnf = +m[4]; }
      const ok = Math.abs(uvta * precio * (1 - dto / 100) - imp) < 0.06;
      out.push({ ean: m[1], cod: m[2], desc: m[3].trim(), uvta, bnf, precio, dto, importe: imp, ok });
    });
    return out;
  }

  const api = { PROVEEDORES, numero, importes, fechas, nifs, lineasPdf, leerFactura, lineasGdC, cuadrar };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Lector = api;
})(typeof window !== "undefined" ? window : globalThis);

/* Lector de facturas de Belleza Lua.
   Extrae datos de facturas en PDF con texto, sin IA y sin coste:
   - lineasPdf(): reconstruye las líneas de texto del PDF con pdf.js
   - leerFactura(): proveedor, NIF, nº, fecha, base, IVA, retención y total
   - lineasGdC(): líneas de producto de las facturas de Germaine de Capuccini
   Funciona igual en el navegador y en Node (pruebas). */
(function (root) {
  "use strict";

  /* Proveedores conocidos: solo empresas y organismos (datos públicos de sus facturas).
     Los particulares (p. ej. un casero persona física) NO van aquí: se aprenden al guardar la primera factura. */
  const PROVEEDORES = {
    // Productos para vender
    A03050572: { nombre: "Germaine de Capuccini", cat: "CREMAS" },
    B08207615: { nombre: "Proveedor ARUDE / S.Expert", cat: "CREMAS" },
    B45441953: { nombre: "Yodeyma", cat: "COLONIAS" },
    // Material de trabajo y cosmética de cabina
    B47683198: { nombre: "The Beauty Corner (Postquam)", cat: "MATERIAL TRABAJO" },
    B22405666: { nombre: "Factoría Cosmética", cat: "MATERIAL TRABAJO" },
    B36924629: { nombre: "Estética Directo", cat: "MATERIAL TRABAJO" },
    B92511427: { nombre: "Pro-Duo Spain", cat: "MATERIAL TRABAJO" },
    A08749608: { nombre: "Investigación y Desarrollo de Cosméticos", cat: "MATERIAL TRABAJO" },
    B66679085: { nombre: "Cosmética natural (FV)", cat: "MATERIAL TRABAJO" },
    B66291097: { nombre: "Amcore Balance", cat: "MATERIAL TRABAJO" },
    B19937309: { nombre: "Depilshop", cat: "MATERIAL TRABAJO" },
    A28581882: { nombre: "Alcampo", cat: "MATERIAL TRABAJO" },
    A46103834: { nombre: "Mercadona", cat: "MATERIAL TRABAJO" },
    E27766591: { nombre: "Farmacia Progreso", cat: "MATERIAL TRABAJO" },
    // Láser subcontratado
    B47758008: { nombre: "Yosemite Estética y Belleza", cat: "LÁSER" },
    // Local
    A60917978: { nombre: "AXA Seguros", cat: "SEGURO ALQUILER" },
    A65067332: { nombre: "Comercializadora Regulada (luz)", cat: "LUZ" },
    A82018474: { nombre: "Telefónica (O2)", cat: "TELÉFONO" },
    // Cobros con tarjeta y gestoría
    IE9813461A: { nombre: "SumUp", cat: "SUM UP" },
    N0074325B: { nombre: "SumUp", cat: "SUM UP" },
    B12936720: { nombre: "Declarando Asesores", cat: "FACTURA DECLARANDO" },
    B27878503: { nombre: "Ubeira Consultores (gestoría)", cat: "OTROS" },
    // Administraciones
    Q2827003A: { nombre: "Tesorería General de la Seguridad Social", cat: "SEGURIDAD SOCIAL" },
    Q2826000H: { nombre: "Agencia Tributaria", cat: "OTROS" },
    P3605700H: { nombre: "Concello de Vigo", cat: "IMPUESTO BASURA" },
    // Otros servicios y compras para el local
    B70779327: { nombre: "Digital Ladpack", cat: "OTROS" },
    B27747351: { nombre: "Vifordat (protección de datos)", cat: "OTROS" },
    B70124607: { nombre: "LOPDAT Protección de Datos", cat: "OTROS" },
    B36821718: { nombre: "Profesionales Teletécnicos", cat: "OTROS" },
    A62588421: { nombre: "Bruneau", cat: "OTROS" },
    B84818442: { nombre: "Leroy Merlin", cat: "OTROS" },
    A28812618: { nombre: "IKEA", cat: "OTROS" },
    B28262822: { nombre: "Rexel", cat: "OTROS" },
    A15340516: { nombre: "Zara Home", cat: "OTROS" },
    A63297097: { nombre: "Media Markt", cat: "OTROS" },
    B64935604: { nombre: "JYSK", cat: "OTROS" },
    B83645069: { nombre: "Maisons du Monde", cat: "OTROS" },
    B82356981: { nombre: "H&M", cat: "OTROS" },
    A85934784: { nombre: "Interflora", cat: "OTROS" },
    B74310202: { nombre: "El Bombillón", cat: "OTROS" },
    B36809036: { nombre: "Ponte las Pilas", cat: "OTROS" },
    B63367411: { nombre: "Gifiess", cat: "OTROS" }
  };
  const PISTAS = [
    [/GERMAINE|8412971\d{6}\s+\d{8}/i, "A03050572"],
    [/YODEYMA/i, "B45441953"],
    [/\bFVC\d{2}-\d{5,}\b|POSTQUAM/i, "B47683198"],
    [/TELEF[OÓ]NICA|\bO2\b/i, "A82018474"],
    [/COMERCIALIZADORA REGULADA|NATURGY/i, "A65067332"],
    [/SUMUP|monthly-fee|comisiones por/i, "IE9813461A"],
    [/DEPILSHOP/i, "B19937309"],
    [/ALCAMPO/i, "A28581882"],
    [/UBEIRA/i, "B27878503"],
    [/DECLARANDO ASESORES/i, "B12936720"],
    [/YOSEMITE/i, "B47758008"]
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

  // 2 decimales; 3-4 decimales solo si van seguidos de € (O2: "24,7934 €"), para no confundir "1 € = 166,386 PTA"
  const RE_IMPORTE = /-?(?:€\s?)?\d{1,3}(?:\.\d{3})*,\d{1,2}(?!\d)|-?(?:€\s?)?\d{1,3}(?:\.\d{3})*,\d{3,4}(?=\s?€)|-?€\s?\d+\.\d{2}(?!\d)|-?\d+,\d{1,2}(?!\d)|-?\d+,\d{3,4}(?=\s?€)/g;
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
    const r2 = /(\d{1,2})\.?\s+(?:de\s+)?([a-záéíóú]+)\s+(?:de\s+|del\s+)?(\d{4})/gi;
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
  function cuadrar(nums, total, opciones) {
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
    if (opciones && opciones.suplidos) for (const rate of [0.21, 0.10, 0.04]) { // con importes sin IVA (suplidos, tasas)
      for (const b of c) {
        const cuota = Math.round(b * rate * 100) / 100;
        if (!c.some((x) => tol(x, cuota))) continue;
        const resto = Math.round((total - b - cuota) * 100) / 100;
        if (resto > 0.5 && c.some((x) => tol(x, resto)) && !tol(resto, b)) return { base: b, tipo: rate * 100, cuota, ret: 0, sinIva: resto };
      }
    }
    if (opciones && opciones.estricto) return null;
    for (const b of c) { // tipos mixtos (p. ej. 10 % + 21 %)
      for (const q of c) {
        if (q !== b && q >= b * 0.035 && q <= b * 0.215 && tol(b + q, total)) return { base: b, tipo: Math.round((q / b) * 1000) / 10, cuota: q, ret: 0, mixto: true };
      }
    }
    return null;
  }

  /* Justificantes de ING: recibos de la Seguridad Social, pagos de impuestos, adeudos y transferencias.
     No son facturas. SS e impuestos son gastos sin factura (este justificante es el comprobante);
     los adeudos y transferencias pagan una factura que se registra aparte. */
  const MODELOS = { "115": "MODELO 115", "130": "MODELO 130", "303": "MODELO 303" };
  function numeroPunto(s) { const n = parseFloat(String(s).replace(/\./g, "").replace(",", ".")); return isFinite(n) ? Math.round(n * 100) / 100 : null; }
  function leerJustificante(lineas, r) {
    const L = lineas.map((l) => l.trim()).filter(Boolean);
    const cab = L[0] || "";
    const tras = (re) => { const i = L.findIndex((l) => re.test(l)); return i >= 0 ? (L[i + 1] || "") : ""; };
    const sinIva = (total) => { r.base = total; r.tipoIva = 0; r.iva = 0; r.retencion = 0; r.total = total; };
    if (/^Recibos Seguridad Social/i.test(cab)) {
      const fila = tras(/Titular de la domiciliaci[oó]n.*Importe/i);
      const m = fila.match(/(\d{2}\/\d{2}\/\d{4})\s+(-?[\d.]+(?:,\d{1,2})?|-?\d+\.\d{2})\s*$/);
      r.tipoDoc = "seguridad_social";
      r.nif = "Q2827003A"; r.proveedor = "Tesorería General de la Seguridad Social"; r.categoria = "SEGURIDAD SOCIAL";
      if (m) { const f = fechas(m[1]); r.fecha = f[0] || ""; sinIva(/,/.test(m[2]) ? numeroPunto(m[2]) : Math.round(parseFloat(m[2]) * 100) / 100); }
      const per = L.find((l) => /PERIODO LIQUIDACI/i.test(l));
      r.numero = per ? per.replace(/.*PERIODO LIQUIDACI[OÓ]N:\s*/i, "SS ").slice(0, 30) : "";
      r.avisos.push("Recibo de autónomos de la Seguridad Social (gasto sin factura)");
      return true;
    }
    if (/^Presentaci[oó]n \/ Pago Impuestos/i.test(cab)) {
      const fila = tras(/Fecha de cargo\s+Importe\s+Modelo/i);
      const m = fila.match(/^(\d{2}\/\d{2}\/\d{4})\s+(-?[\d.]*\d,\d{2})\s+(\d{3})/);
      r.tipoDoc = "impuesto";
      r.nif = "Q2826000H"; r.proveedor = "Agencia Tributaria";
      if (m) {
        r.fecha = fechas(m[1])[0] || ""; sinIva(numeroPunto(m[2]));
        r.categoria = MODELOS[m[3]] || "OTROS";
        const per = tras(/Ejercicio\s+Per[ií]odo\s+Concepto/i).match(/^(\d{4})\s+\S+\s+(\d[TP])/i);
        r.numero = "Modelo " + m[3] + (per ? " " + per[2].toUpperCase() + " " + per[1] : "");
        r.avisos.push("Pago del modelo " + m[3] + " a Hacienda (gasto sin factura)");
      } else { r.categoria = "OTROS"; r.avisos.push("Pago de impuestos: revisa el modelo y el importe"); }
      return true;
    }
    if (/^Adeudo recibido/i.test(cab) || /^Transferencia (emitida|ordenada)/i.test(cab) || /^TRANSFERENCIA ORDENADA/i.test(cab)) {
      r.tipoDoc = "justificante_pago";
      const id = (L.join("\n").match(/Id\.\s*emisor:\s*ES\d{2}\d{3}([A-Z]\d{7}[0-9A-J]|\d{8}[A-Z])/i) || [])[1];
      if (id) r.nif = id.toUpperCase();
      const fila = tras(/Importe( en)? euros/i);
      const imp = (fila.match(/-?\d{1,3}(?:\.\d{3})*,\d{2}/) || [])[0];
      if (imp) r.total = numeroPunto(imp);
      const f = fechas(L.slice(0, 3).join(" ")); if (f.length) r.fecha = f[0];
      r.avisos.push("Justificante de pago del banco, no una factura: si ya registraste la factura, no lo guardes (se contaría dos veces)");
      return true;
    }
    if (/^INFORMACI[OÓ]N DE LA PRESENTACI[OÓ]N DE LA/i.test(cab)) {
      r.tipoDoc = "declaracion";
      r.avisos.push("Declaración de impuestos presentada: no es un gasto; el pago se registra con el justificante del banco");
      return true;
    }
    if (/^Certificado de Movimientos/i.test(cab)) {
      r.tipoDoc = "extracto";
      r.avisos.push("Extracto de movimientos del banco: no es una factura");
      return true;
    }
    return false;
  }

  function leerFactura(lineas, opciones) {
    const propio = ((opciones && opciones.nifPropio) || "").toUpperCase().replace(/[^0-9A-Z]/g, "");
    const conocidos = Object.assign({}, PROVEEDORES, (opciones && opciones.aprendidos) || {});
    const texto = lineas.join("\n");
    const r = { proveedor: "", nif: "", numero: "", fecha: "", base: null, tipoIva: null, iva: null, retencion: 0, total: null, categoria: "", avisos: [], tipoDoc: "factura" };
    if (leerJustificante(lineas, r)) {
      if (r.nif && conocidos[r.nif]) { r.proveedor = r.proveedor || conocidos[r.nif].nombre; r.categoria = r.categoria || conocidos[r.nif].cat; }
      return r;
    }

    // Proveedor
    const lista = nifs(texto).filter((n) => n !== propio && n !== "ES" + propio);
    let nif = lista.find((n) => conocidos[n]);
    if (!nif) { const p = PISTAS.find(([re]) => re.test(texto)); if (p) nif = p[1]; }
    if (!nif) nif = lista[0] || "";
    r.nif = nif;
    if (conocidos[nif]) { r.proveedor = conocidos[nif].nombre; r.categoria = conocidos[nif].cat; }

    // Total
    const etiquetas = [/total\s+de\s+la\s+factura/i, /total\s+a\s+pagar/i, /totales?\s+a\s+pagar/i, /importe\s+a\s+pagar/i, /total\s+a\s+abonar/i, /importe\s+factura/i, /total\s+factura/i, /total\s+importe\s+factura/i, /\bT\s+O\s+T\s+A\s+L\b/, /^\s*total\b(?!\s*(bruto|base|impuesto|deducc|de los|es\b))/i, /\btotal\b(?!\s*(bruto|base|impuesto|deducc|de los|ejercicio|con iva|iva|es\b|rec\b))/i];
    const rotulados = [];
    lineas.forEach((l, i) => {
      if (!/total|importe\s+a\s+pagar/i.test(l) || /subtotal|sub-total/i.test(l)) return;
      importes(l).concat(importes(lineas[i + 1] || "")).forEach((n) => rotulados.push(Math.abs(n)));
    });
    const aPagar = (() => { for (let i = 0; i < lineas.length; i++) if (/total\s+a\s+pagar|importe\s+a\s+pagar/i.test(lineas[i])) { const v = importes(lineas[i].slice(lineas[i].search(/a\s+pagar/i))); if (v.length) return v[v.length - 1]; } return null; })();
    for (const re of etiquetas) {
      for (let i = 0; i < lineas.length && r.total == null; i++) {
        if (!re.test(lineas[i]) || /subtotal|sub-total/i.test(lineas[i])) continue;
        const after = lineas[i].slice(lineas[i].search(re));
        let v = importes(after);
        if (!v.length && i + 1 < lineas.length) v = importes(lineas[i + 1]);
        if (!v.length && i + 2 < lineas.length) v = importes(lineas[i + 2]);
        if (v.length) { const x = v[v.length - 1]; if (x !== 0) r.total = x; }
      }
      if (r.total != null) break;
    }

    // Base, IVA y retención
    const todos = [];
    lineas.forEach((l) => importes(l).forEach((n) => todos.push(n)));
    const exento = /exent[ao]s?\s+de\s+IVA|IVA\s*0\s*%/i.test(texto);
    const abs = todos.map(Math.abs);
    const aplicar = (q, signo) => {
      r.base = signo * q.base; r.tipoIva = q.tipo; r.iva = signo * q.cuota; r.retencion = signo * q.ret;
      if (q.mixto) r.avisos.push("Varios tipos de IVA");
      if (q.sinIva) r.avisos.push("Incluye " + String(q.sinIva).replace(".", ",") + " € sin IVA (suplidos o tasas)");
    };
    const opc = { suplidos: /suplido|no\s+sujet/i.test(texto) };
    let q = null;
    if (r.total != null && !exento) q = cuadrar(abs, Math.abs(r.total), opc);
    if (!q && !exento) {
      // El total rotulado no cuadra (o no hay): se busca el mayor importe que sea base + IVA
      const cands = [...new Set(abs.filter((n) => n > 0))].sort((a, b) => b - a);
      let z = null, t2 = null;
      for (const t of cands) { z = cuadrar(abs, t, { estricto: true, suplidos: opc.suplidos }); if (z) { t2 = t; break; } }
      const casi = (a, b) => Math.abs(a - b) <= 0.02;
      // Se acepta si no había total, si el "total" leído era en realidad la base, el IVA o la retención,
      // o si el nuevo total también aparece junto a un rótulo de total (p. ej. "TOTAL FACTURA" frente a "TOTAL A PAGAR")
      const T = r.total == null ? null : Math.abs(r.total);
      if (z && (T == null || [z.base, z.cuota, z.ret].some((x) => casi(x, T)) || rotulados.some((x) => casi(x, t2)))) {
        q = z;
        const signo = r.total != null && r.total < 0 ? -1 : 1;
        if (T == null || !casi(T, t2)) r.avisos.push("Total deducido de base + IVA: revísalo");
        r.total = signo * t2;
      }
    }
    if (aPagar != null && r.total != null && Math.abs(Math.abs(aPagar) - Math.abs(r.total)) > 0.02) r.avisos.push("Total a pagar distinto del total de la factura (" + String(aPagar).replace(".", ",") + " €): descuentos, bonificaciones o pagos a cuenta");
    if (r.total != null) {
      const signo = r.total < 0 ? -1 : 1;
      if (exento) { r.base = r.total; r.tipoIva = 0; r.iva = 0; r.avisos.push("Exento de IVA"); }
      else if (q) aplicar(q, signo);
      else r.avisos.push("No se encontró la base imponible");
      if (signo < 0) r.avisos.push("Abono (factura rectificativa): importes en negativo");
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

  /* Líneas de colonias de una factura de Yodeyma: "4,00 EAU TOILET. POUR FEMME 100 ML   13,99   55,96".
     No traen código de barras ni referencia: se identifican por la descripción. Las de 0 € (probadores,
     expositores, kits promocionales) no son producto para vender y se ignoran. */
  function lineasYodeyma(lineas) {
    const out = [];
    const re = /^\s*(\d{1,4},\d{2})\s+(.+?)\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})\s*$/;
    lineas.forEach((l) => {
      const m = l.match(re);
      if (!m) return;
      const uds = numero(m[1]), precio = numero(m[3]), imp = numero(m[4]);
      const desc = m[2].replace(/(\s+-?\d{1,3}(?:\.\d{3})*,\d{2})+\s*$/, "").replace(/\s+/g, " ").trim().toUpperCase();
      if (!desc || !(precio > 0) || !(uds > 0) || uds % 1) return;
      const cod = "YOD-" + desc.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60);
      const gratis = imp === 0; // regalada (precio con descuento total): entra como bonificada
      out.push({ ean: "", cod, desc, uvta: gratis ? 0 : uds, bnf: gratis ? uds : 0, precio, dto: 0, importe: imp, ok: gratis || Math.abs(uds * precio - imp) < 0.06, prov: "Yodeyma", cat: "Colonia" });
    });
    return out;
  }

  const api = { PROVEEDORES, lineasYodeyma, numero, importes, fechas, nifs, lineasPdf, leerFactura, leerJustificante, lineasGdC, cuadrar };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Lector = api;
})(typeof window !== "undefined" ? window : globalThis);

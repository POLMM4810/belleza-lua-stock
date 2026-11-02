/* Excel anual de Belleza Lua: cuentas del año explicadas, con gráficas.
   Genera el .xlsx en el propio móvil, sin librerías (zip sin comprimir + SpreadsheetML + DrawingML),
   con el valor ya calculado en cada fórmula y en cada gráfica para que la vista previa del iPhone
   muestre números y gráficas; Excel recalcula al abrir (fullCalcOnLoad).
   Basado en el generador de las apps de gastos de Paulo.
   Funciona igual en el navegador (window.ExcelAnual) y en Node (pruebas). */
(function (root) {
  "use strict";
  var MESES_C = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  var MESES_L = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

  /* Categorías del Excel anual (las de siempre de Miriam), agrupadas para entender en qué se va el dinero */
  var GRUPOS = [
    { id: "local", nombre: "El local", color: "5B7DB1" },
    { id: "producto", nombre: "Producto para vender", color: "C2185B" },
    { id: "trabajo", nombre: "Trabajo y equipo", color: "8E6BBF" },
    { id: "impuestos", nombre: "Impuestos y autónomo", color: "E08A2C" },
    { id: "servicios", nombre: "Servicios y otros", color: "6B9E78" }
  ];
  var CATS = [
    ["ALQUILER", "Alquiler del local", "local"],
    ["LUZ-TLF", "Luz y teléfono", "local"],
    ["SEGURO LOCAL", "Seguro del local", "local"],
    ["SEGURO ALQUILER", "Seguro del alquiler", "local"],
    ["IMPUESTO BASURA", "Impuesto de basura", "local"],
    ["CREMAS", "Cremas (Germaine y otras)", "producto"],
    ["COLONIAS", "Colonias (Yodeyma)", "producto"],
    ["MATERIAL TRABAJO", "Material de trabajo", "trabajo"],
    ["SERVICIO LASER", "Servicio de láser", "trabajo"],
    ["APARATOLOGIA CENTRO", "Aparatología (máquinas)", "trabajo"],
    ["MOBILIARIO", "Mobiliario", "trabajo"],
    ["EXTINTORES", "Extintores", "trabajo"],
    ["CURSOS", "Cursos y formación", "trabajo"],
    ["TASA AUTONOMO", "Cuota de autónomo (Seguridad Social)", "impuestos"],
    ["IRPF MODELO 130", "IRPF de Miriam (modelo 130)", "impuestos"],
    ["IVA MODELO 303", "IVA (modelo 303)", "impuestos"],
    ["IRPF MODELO 115", "IRPF del alquiler (modelo 115)", "impuestos"],
    ["RETENCIONES MODELO 111", "Retenciones (modelo 111)", "impuestos"],
    ["GESTORIA", "Gestoría", "servicios"],
    ["PROTECCIÓN DATOS", "Protección de datos", "servicios"],
    ["COMISIONES SUM UP", "Comisiones de SumUp (tarjeta)", "servicios"],
    ["OTROS NO FACTURA", "Otros gastos", "servicios"]
  ];
  var SIN = "SIN CLASIFICAR";
  var CAT_IDX = {}; CATS.forEach(function (c, i) { CAT_IDX[c[0]] = i; });
  /* Categoría de la app -> categoría del Excel anual */
  var DESDE_APP = {
    "ALQUILER": "ALQUILER", "SEGURO ALQUILER": "SEGURO ALQUILER", "EXTINTORES": "EXTINTORES", "LUZ": "LUZ-TLF", "TELÉFONO": "LUZ-TLF",
    "SUM UP": "COMISIONES SUM UP", "FACTURA DECLARANDO": "GESTORIA", "SEGURIDAD SOCIAL": "TASA AUTONOMO", "LÁSER": "SERVICIO LASER",
    "LETRA MÁQUINA": "APARATOLOGIA CENTRO", "IMPUESTO BASURA": "IMPUESTO BASURA", "MODELO 115": "IRPF MODELO 115", "MODELO 130": "IRPF MODELO 130",
    "MODELO 303": "IVA MODELO 303", "CREMAS": "CREMAS", "COLONIAS": "COLONIAS", "MATERIAL TRABAJO": "MATERIAL TRABAJO", "OTROS": "OTROS NO FACTURA"
  };
  var POR_NIF = { B27878503: "GESTORIA", B12936720: "GESTORIA", B27747351: "PROTECCIÓN DATOS", B70124607: "PROTECCIÓN DATOS" };
  function catDeGasto(g) { return POR_NIF[g.nif] || DESDE_APP[g.categoria] || "OTROS NO FACTURA"; }
  function catDeHist(label) { var k = String(label || "").trim().toUpperCase(); return CAT_IDX[k] != null ? k : "OTROS NO FACTURA"; }

  // ── Utilidades ──────────────────────────────────────────────────────────
  function pad(n) { return String(n).padStart(2, "0"); }
  function r2(n) { var x = Number(n || 0), s = x < 0 ? -1 : 1; return s * Number(Math.round(Number(Math.abs(x).toFixed(6) + "e2")) + "e-2"); }
  function col(i) { var s = ""; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function esc(s) { return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function numStr(v) { var s = String(Number(Number(v).toFixed(10))); return s.indexOf("e") >= 0 ? Number(v).toFixed(10) : s; }
  function sum(a) { var t = 0; for (var i = 0; i < a.length; i++) t += (typeof a[i] === "number" ? a[i] : 0); return t; }
  function q(name) { return "'" + name.replace(/'/g, "''") + "'"; }
  function at(arr, i) { if (!arr) return null; var v = arr[i]; if (v === undefined) v = arr[String(i)]; return typeof v === "number" && isFinite(v) ? v : null; }
  function eurTxt(v) { return (Math.round(v) || 0).toLocaleString("es-ES") + " €"; }

  // ── ZIP (sin comprimir) ─────────────────────────────────────────────────
  var CRC = (function () { var t = []; for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u) { var c = 0xFFFFFFFF; for (var i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function utf8(s) { return new TextEncoder().encode(s); }
  function zip(files) {
    var chunks = [], central = [], off = 0, total = 0;
    files.forEach(function (f) {
      var nb = utf8(f.name), db = utf8(f.data), crc = crc32(db);
      var h = new Uint8Array(30 + nb.length), v = new DataView(h.buffer);
      v.setUint32(0, 0x04034b50, true); v.setUint16(4, 20, true); v.setUint16(6, 0x0800, true); v.setUint16(8, 0, true);
      v.setUint16(10, 0, true); v.setUint16(12, 33, true); v.setUint32(14, crc, true); v.setUint32(18, db.length, true); v.setUint32(22, db.length, true);
      v.setUint16(26, nb.length, true); v.setUint16(28, 0, true); h.set(nb, 30);
      var c = new Uint8Array(46 + nb.length), w = new DataView(c.buffer);
      w.setUint32(0, 0x02014b50, true); w.setUint16(4, 20, true); w.setUint16(6, 20, true); w.setUint16(8, 0x0800, true); w.setUint16(10, 0, true);
      w.setUint16(12, 0, true); w.setUint16(14, 33, true); w.setUint32(16, crc, true); w.setUint32(20, db.length, true); w.setUint32(24, db.length, true);
      w.setUint16(28, nb.length, true); w.setUint32(42, off, true); c.set(nb, 46);
      chunks.push(h, db); central.push(c); off += h.length + db.length;
    });
    var csz = 0; central.forEach(function (c) { csz += c.length; });
    var e = new Uint8Array(22), ev = new DataView(e.buffer);
    ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.length, true); ev.setUint16(10, files.length, true); ev.setUint32(12, csz, true); ev.setUint32(16, off, true);
    var all = chunks.concat(central, [e]); all.forEach(function (a) { total += a.length; });
    var out = new Uint8Array(total), p = 0; all.forEach(function (a) { out.set(a, p); p += a.length; });
    return out;
  }

  // ── Estilos ─────────────────────────────────────────────────────────────
  var EUR = '#,##0.00 "€";[Red]-#,##0.00 "€";"-"', EUR0 = '#,##0 "€";[Red]-#,##0 "€";"-"', PCT = '0.0%;[Red]-0.0%;"-"', PCTS = '+0.0%;[Red]-0.0%;0.0%';
  function Styles() {
    this.fonts = ['<font><sz val="10"/><name val="Arial"/><family val="2"/></font>'];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.nf = []; this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']; this.map = {}; this.dxfs = [];
  }
  Styles.prototype.idx = function (arr, x) { var i = arr.indexOf(x); if (i < 0) { arr.push(x); i = arr.length - 1; } return i; };
  Styles.prototype.s = function (o) {
    o = o || {}; var k = JSON.stringify(o); if (this.map[k] != null) return this.map[k];
    var font = '<font>' + (o.b ? '<b/>' : '') + (o.i ? '<i/>' : '') + '<sz val="' + (o.sz || 10) + '"/><color rgb="FF' + (o.c || "1F2937") + '"/><name val="Arial"/><family val="2"/></font>';
    var fid = this.idx(this.fonts, font), flid = 0, bid = 0, nid = 0;
    if (o.fill) flid = this.idx(this.fills, '<fill><patternFill patternType="solid"><fgColor rgb="FF' + o.fill + '"/><bgColor indexed="64"/></patternFill></fill>');
    if (o.bd) {
      var ln = function (side) { var n = ({ l: "left", r: "right", t: "top", b: "bottom" })[side]; return o.bd.indexOf(side) >= 0 ? '<' + n + ' style="' + (o.bw || "thin") + '"><color rgb="FF' + (o.bc || "CBD5E1") + '"/></' + n + '>' : '<' + n + '/>'; };
      bid = this.idx(this.borders, '<border>' + ln("l") + ln("r") + ln("t") + ln("b") + '<diagonal/></border>');
    }
    if (o.nf) { var j = this.nf.indexOf(o.nf); if (j < 0) { this.nf.push(o.nf); j = this.nf.length - 1; } nid = 164 + j; }
    var al = (o.h || o.v || o.wrap || o.ind) ? '<alignment' + (o.h ? ' horizontal="' + o.h + '"' : '') + ' vertical="' + (o.v || "center") + '"' + (o.wrap ? ' wrapText="1"' : '') + (o.ind ? ' indent="' + o.ind + '"' : '') + '/>' : '';
    var xf = '<xf numFmtId="' + nid + '" fontId="' + fid + '" fillId="' + flid + '" borderId="' + bid + '" xfId="0"' + (nid ? ' applyNumberFormat="1"' : '') + ' applyFont="1"' + (flid ? ' applyFill="1"' : '') + (bid ? ' applyBorder="1"' : '') + (al ? ' applyAlignment="1">' + al + '</xf>' : '/>');
    this.xfs.push(xf); this.map[k] = this.xfs.length - 1; return this.map[k];
  };
  Styles.prototype.dxf = function (fc, bg, b) { this.dxfs.push('<dxf><font>' + (b ? '<b/>' : '') + '<color rgb="FF' + fc + '"/></font><fill><patternFill patternType="solid"><bgColor rgb="FF' + bg + '"/></patternFill></fill></dxf>'); return this.dxfs.length - 1; };
  Styles.prototype.xml = function () {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + (this.nf.length ? '<numFmts count="' + this.nf.length + '">' + this.nf.map(function (f, i) { return '<numFmt numFmtId="' + (164 + i) + '" formatCode="' + esc(f) + '"/>'; }).join("") + '</numFmts>' : '')
      + '<fonts count="' + this.fonts.length + '">' + this.fonts.join("") + '</fonts><fills count="' + this.fills.length + '">' + this.fills.join("") + '</fills>'
      + '<borders count="' + this.borders.length + '">' + this.borders.join("") + '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + '<cellXfs count="' + this.xfs.length + '">' + this.xfs.join("") + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
      + '<dxfs count="' + this.dxfs.length + '">' + this.dxfs.join("") + '</dxfs></styleSheet>';
  };

  // ── Hoja ────────────────────────────────────────────────────────────────
  function Sheet(name) { this.name = name; this.cells = {}; this.maxR = 1; this.maxC = 1; this.cols = {}; this.merges = []; this.cf = []; this.freeze = null; this.grid = true; this.charts = []; this.heights = {}; this.tab = null; }
  Sheet.prototype.put = function (r, c, v, s, f) { this.cells[r + "," + c] = { v: v, s: s || 0, f: f }; if (r > this.maxR) this.maxR = r; if (c > this.maxC) this.maxC = c; return this; };
  Sheet.prototype.get = function (r, c) { var x = this.cells[r + "," + c]; return x ? x.v : null; };
  Sheet.prototype.width = function (c, w) { this.cols[c] = w; return this; };
  Sheet.prototype.xml = function (drawRid) {
    var rows = [], self = this;
    for (var r = 1; r <= this.maxR; r++) {
      var cs = [];
      for (var c = 1; c <= this.maxC; c++) {
        var x = this.cells[r + "," + c]; if (!x) continue;
        var ref = col(c) + r, s = x.s ? ' s="' + x.s + '"' : '';
        if (x.f != null) {
          var isS = typeof x.v === "string";
          cs.push('<c r="' + ref + '"' + s + (isS ? ' t="str"' : '') + '><f>' + esc(x.f) + '</f>' + (x.v == null ? '' : '<v>' + (isS ? esc(x.v) : numStr(x.v)) + '</v>') + '</c>');
        } else if (typeof x.v === "number" && isFinite(x.v)) {
          cs.push('<c r="' + ref + '"' + s + '><v>' + numStr(x.v) + '</v></c>');
        } else if (x.v != null && x.v !== "") {
          cs.push('<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(String(x.v).slice(0, 32000)) + '</t></is></c>');
        } else cs.push('<c r="' + ref + '"' + s + '/>');
      }
      var h = this.heights[r];
      if (cs.length || h) rows.push('<row r="' + r + '"' + (h ? ' ht="' + h + '" customHeight="1"' : '') + '>' + cs.join("") + '</row>');
    }
    var colsX = Object.keys(this.cols).map(Number).sort(function (a, b) { return a - b; }).map(function (c) { return '<col min="' + c + '" max="' + c + '" width="' + self.cols[c] + '" customWidth="1"/>'; }).join("");
    var pane = this.freeze ? '<pane' + (this.freeze[1] > 1 ? ' xSplit="' + (this.freeze[1] - 1) + '"' : '') + (this.freeze[0] > 1 ? ' ySplit="' + (this.freeze[0] - 1) + '"' : '') + ' topLeftCell="' + col(this.freeze[1]) + this.freeze[0] + '" activePane="' + (this.freeze[1] > 1 && this.freeze[0] > 1 ? "bottomRight" : this.freeze[0] > 1 ? "bottomLeft" : "topRight") + '" state="frozen"/>' : '';
    var cf = this.cf.map(function (b, i) { return '<conditionalFormatting sqref="' + b.ref + '">' + b.rules.map(function (ru, j) { return '<cfRule type="expression" dxfId="' + ru.dxf + '" priority="' + (i * 10 + j + 1) + '"><formula>' + esc(ru.f) + '</formula></cfRule>'; }).join("") + '</conditionalFormatting>'; }).join("");
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + (this.tab ? '<sheetPr><tabColor rgb="FF' + this.tab + '"/></sheetPr>' : '')
      + '<dimension ref="A1:' + col(this.maxC) + this.maxR + '"/><sheetViews><sheetView' + (this.grid ? '' : ' showGridLines="0"') + ' workbookViewId="0">' + pane + '</sheetView></sheetViews>'
      + '<sheetFormatPr defaultRowHeight="15"/>' + (colsX ? '<cols>' + colsX + '</cols>' : '') + '<sheetData>' + rows.join("") + '</sheetData>'
      + (this.merges.length ? '<mergeCells count="' + this.merges.length + '">' + this.merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join("") + '</mergeCells>' : '')
      + cf + '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>'
      + '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
      + (drawRid ? '<drawing r:id="' + drawRid + '"/>' : '') + '</worksheet>';
  };

  // ── Gráficos (con los valores guardados para la vista previa del iPhone) ─
  function txPr(sz, c, b) { return '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="' + sz + '"' + (b ? ' b="1"' : '') + '><a:solidFill><a:srgbClr val="' + c + '"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="es-ES"/></a:p></c:txPr>'; }
  function strCache(vals) { return '<c:strCache><c:ptCount val="' + vals.length + '"/>' + vals.map(function (v, i) { return '<c:pt idx="' + i + '"><c:v>' + esc(v) + '</c:v></c:pt>'; }).join("") + '</c:strCache>'; }
  function numCache(vals) { return '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="' + vals.length + '"/>' + vals.map(function (v, i) { return (typeof v === "number" && isFinite(v)) ? '<c:pt idx="' + i + '"><c:v>' + numStr(v) + '</c:v></c:pt>' : ''; }).join("") + '</c:numCache>'; }
  function serXml(s, i, cats, kind) {
    var sp = kind === "line" ? '<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="' + s.color + '"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="' + s.color + '"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>'
      : '<c:spPr><a:solidFill><a:srgbClr val="' + s.color + '"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>';
    var dl = s.labels ? '<c:dLbls><c:numFmt formatCode="' + esc(EUR0) + '" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' + txPr(800, "374151") + '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : '';
    return '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/><c:tx><c:strRef><c:f>' + esc(s.nameRef) + '</c:f>' + strCache([s.name]) + '</c:strRef></c:tx>' + sp + dl
      + '<c:cat><c:strRef><c:f>' + esc(cats.ref) + '</c:f>' + strCache(cats.vals) + '</c:strRef></c:cat>'
      + '<c:val><c:numRef><c:f>' + esc(s.ref) + '</c:f>' + numCache(s.vals) + '</c:numRef></c:val>' + (kind === "line" ? '<c:smooth val="0"/>' : '') + '</c:ser>';
  }
  function chartXml(o) {
    var horiz = o.dir === "bar", ax1 = 50010, ax2 = 50020;
    var bars = o.series.filter(function (s) { return s.kind !== "line"; }), lines = o.series.filter(function (s) { return s.kind === "line"; });
    var plot = '';
    if (bars.length) plot += '<c:barChart><c:barDir val="' + (horiz ? "bar" : "col") + '"/><c:grouping val="' + (o.stacked ? "stacked" : "clustered") + '"/><c:varyColors val="0"/>'
      + bars.map(function (s, i) { return serXml(s, i, o.cats, "bar"); }).join("") + '<c:gapWidth val="' + (o.gap || 60) + '"/>' + (o.stacked ? '<c:overlap val="100"/>' : '') + '<c:axId val="' + ax1 + '"/><c:axId val="' + ax2 + '"/></c:barChart>';
    if (lines.length) plot += '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + lines.map(function (s, i) { return serXml(s, bars.length + i, o.cats, "line"); }).join("") + '<c:marker val="1"/><c:axId val="' + ax1 + '"/><c:axId val="' + ax2 + '"/></c:lineChart>';
    var axLine = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:ln></c:spPr>';
    var cat = '<c:catAx><c:axId val="' + ax1 + '"/><c:scaling><c:orientation val="' + (horiz ? "maxMin" : "minMax") + '"/></c:scaling><c:delete val="0"/><c:axPos val="' + (horiz ? "l" : "b") + '"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>' + axLine + txPr(800, "4B5563") + '<c:crossAx val="' + ax2 + '"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>';
    var val = '<c:valAx><c:axId val="' + ax2 + '"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + (horiz ? "b" : "l") + '"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E5E7EB"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="' + esc(EUR0) + '" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>' + txPr(800, "6B7280") + '<c:crossAx val="' + ax1 + '"/><c:crosses val="' + (horiz ? "max" : "autoZero") + '"/><c:crossBetween val="between"/></c:valAx>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/>'
      + '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr><a:r><a:rPr lang="es-ES" sz="1100" b="1"><a:solidFill><a:srgbClr val="111827"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>' + esc(o.title) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
      + '<c:plotArea><c:layout/>' + plot + cat + val + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>'
      + (o.series.length > 1 ? '<c:legend><c:legendPos val="t"/><c:overlay val="0"/>' + txPr(900, "374151") + '</c:legend>' : '')
      + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="E5E7EB"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="es-ES"/></a:p></c:txPr></c:chartSpace>';
  }
  function drawingXml(charts) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
      + charts.map(function (ch, i) { var a = ch.at; return '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>' + a[1] + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + a[0] + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + a[3] + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + a[2] + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        + '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (i + 2) + '" name="Gráfico ' + (i + 1) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId' + (i + 1) + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'; }).join("")
      + '</xdr:wsDr>';
  }

  // ── Ensamblado del libro ────────────────────────────────────────────────
  function pack(sheets, ST, title) {
    var files = [], ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>', '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>', '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'];
    var wbRels = [], chartN = 0, drawN = 0;
    sheets.forEach(function (sh, i) {
      var n = i + 1, rid = null;
      if (sh.charts.length) {
        drawN++; rid = "rId1";
        files.push({ name: "xl/worksheets/_rels/sheet" + n + ".xml.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + drawN + '.xml"/></Relationships>' });
        files.push({ name: "xl/drawings/drawing" + drawN + ".xml", data: drawingXml(sh.charts) });
        ct.push('<Override PartName="/xl/drawings/drawing' + drawN + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
        var dr = [];
        sh.charts.forEach(function (ch, j) {
          chartN++; files.push({ name: "xl/charts/chart" + chartN + ".xml", data: ch.xml });
          ct.push('<Override PartName="/xl/charts/chart' + chartN + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
          dr.push('<Relationship Id="rId' + (j + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + chartN + '.xml"/>');
        });
        files.push({ name: "xl/drawings/_rels/drawing" + drawN + ".xml.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + dr.join("") + '</Relationships>' });
      }
      files.push({ name: "xl/worksheets/sheet" + n + ".xml", data: sh.xml(rid) });
      ct.push('<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      wbRels.push('<Relationship Id="rId' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>');
    });
    wbRels.push('<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>');
    files.push({ name: "xl/workbook.xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr/><bookViews><workbookView activeTab="0"/></bookViews><sheets>'
      + sheets.map(function (sh, i) { return '<sheet name="' + esc(sh.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join("")
      + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>' });
    files.push({ name: "xl/_rels/workbook.xml.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wbRels.join("") + '</Relationships>' });
    files.push({ name: "xl/styles.xml", data: ST.xml() });
    files.push({ name: "_rels/.rels", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>' });
    files.push({ name: "docProps/core.xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + esc(title) + '</dc:title><dc:creator>Belleza Lua</dc:creator></cp:coreProperties>' });
    files.push({ name: "docProps/app.xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Belleza Lua</Application></Properties>' });
    files.unshift({ name: "[Content_Types].xml", data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + ct.join("") + '</Types>' });
    return zip(files);
  }

  // ── Modelo: de dónde sale cada número ───────────────────────────────────
  /* o = { hist, ingresos:[{fecha,total}], gastos:[{fecha,categoria,nif,proveedor,concepto,base,iva,total}], year, nowKey:"AAAA-MM", generado }
     Meses anteriores a hist.desde: del histórico (Excel antiguo). Desde hist.desde: de la app. */
  function modelo(o, year) {
    var H = o.hist || {}, desde = H.desde || "0000-00", now = o.nowKey;
    var hy = (H.anios || {})[String(year)] || null;
    var M = { year: year, meses: [], ing: [], gas: [], cat: {}, catAnual: null, sinCat: [], gastosApp: [] };
    CATS.forEach(function (c) { M.cat[c[0]] = []; });
    for (var m = 0; m < 12; m++) {
      var key = year + "-" + pad(m + 1);
      var src = key < desde ? "hist" : (key <= now ? "app" : "futuro");
      var mes = { key: key, src: src, cerrado: key < now, actual: key === now, datos: false };
      if (src === "hist") {
        var ing = hy ? at(hy.ingresos, m) : null, gas = hy ? at(hy.gastos, m) : null;
        M.ing[m] = ing; M.gas[m] = gas; mes.datos = ing != null || gas != null;
        var sc = 0;
        CATS.forEach(function (c) { M.cat[c[0]][m] = null; });
        if (hy && hy.cat) Object.keys(hy.cat).forEach(function (lab) { var v = at(hy.cat[lab], m); if (v != null) { var k = catDeHist(lab); M.cat[k][m] = r2((M.cat[k][m] || 0) + v); sc += v; } });
        M.sinCat[m] = gas != null && hy && hy.cat ? r2(gas - sc) : null;
        if (hy && !hy.cat && gas != null) { M.sinCat[m] = gas; M.sinDesglose = true; }
      } else if (src === "app") {
        var ingM = (o.ingresos || []).filter(function (x) { return String(x.fecha || "").slice(0, 7) === key; });
        var gasM = (o.gastos || []).filter(function (x) { return String(x.fecha || "").slice(0, 7) === key; });
        M.ing[m] = r2(sum(ingM.map(function (x) { return Number(x.total) || 0; })));
        M.gas[m] = r2(sum(gasM.map(function (x) { return Number(x.total) || 0; })));
        mes.datos = ingM.length > 0 || gasM.length > 0;
        CATS.forEach(function (c) { M.cat[c[0]][m] = null; });
        gasM.forEach(function (g) { var k = catDeGasto(g); M.cat[k][m] = r2((M.cat[k][m] || 0) + (Number(g.total) || 0)); });
        M.sinCat[m] = 0;
        M.gastosApp = M.gastosApp.concat(gasM);
        if (!mes.datos) { M.ing[m] = null; M.gas[m] = null; M.sinCat[m] = null; }
      } else {
        M.ing[m] = null; M.gas[m] = null; M.sinCat[m] = null; CATS.forEach(function (c) { M.cat[c[0]][m] = null; });
      }
      M.meses.push(mes);
    }
    if (hy && hy.cat_anual) { // 2023: solo totales anuales por categoría
      M.catAnual = {}; Object.keys(hy.cat_anual).forEach(function (lab) { var k = catDeHist(lab); M.catAnual[k] = r2((M.catAnual[k] || 0) + (hy.cat_anual[lab] || 0)); });
    }
    return M;
  }
  function aniosDisponibles(o) {
    var ys = Object.keys((o.hist || {}).anios || {}).map(Number);
    (o.ingresos || []).concat(o.gastos || []).forEach(function (x) { var y = +String(x.fecha || "").slice(0, 4); if (y) ys.push(y); });
    ys.push(o.year);
    ys = ys.filter(function (y, i) { return ys.indexOf(y) === i && y <= o.year; }).sort();
    return ys;
  }

  // ── Libro ───────────────────────────────────────────────────────────────
  function build(o) {
    var year = o.year, M = modelo(o, year), P = modelo(o, year - 1), ST = new Styles();
    var C = { navy: "4A2735", ink: "1F2937", muted: "6B7280", blue: "1D4ED8", sect: "F3E8EC", zebra: "FAF7F8", input: "FFF7CC", line: "D9CDD2", rosa: "7A4A5A" };
    var S = {
      title: ST.s({ b: 1, sz: 18, c: C.rosa }), sub: ST.s({ sz: 9, c: C.muted }), small: ST.s({ sz: 8, c: C.muted }),
      h2: ST.s({ b: 1, sz: 12, c: C.rosa }), txt: ST.s({ c: C.ink, wrap: 1, v: "top" }), txtB: ST.s({ b: 1, c: C.ink, wrap: 1, v: "top" }),
      head: ST.s({ b: 1, c: "FFFFFF", fill: C.rosa, h: "center", wrap: 1 }), headL: ST.s({ b: 1, c: "FFFFFF", fill: C.rosa, h: "left", ind: 1, wrap: 1 }),
      sect: ST.s({ b: 1, c: C.rosa, fill: C.sect, ind: 1 }), sectN: ST.s({ b: 1, c: C.rosa, fill: C.sect, nf: EUR }), sectP: ST.s({ b: 1, c: C.rosa, fill: C.sect, nf: PCT }),
      label: ST.s({ c: C.ink, ind: 1, bd: "b", bc: "EEE7EA" }), labelB: ST.s({ b: 1, c: C.ink, ind: 1, bd: "t", bc: C.line, fill: "F5EEF1" }),
      in: ST.s({ c: C.blue, nf: EUR, bd: "b", bc: "EEE7EA" }), f: ST.s({ c: "000000", nf: EUR, bd: "b", bc: "EEE7EA" }), fB: ST.s({ b: 1, c: "000000", nf: EUR, bd: "t", bc: C.line, fill: "F5EEF1" }),
      pct: ST.s({ c: "000000", nf: PCT, h: "center", bd: "b", bc: "EEE7EA" }), pctS: ST.s({ c: "000000", nf: PCTS, h: "center", bd: "b", bc: "EEE7EA" }), pctB: ST.s({ b: 1, c: "000000", nf: PCT, h: "center", bd: "t", bc: C.line, fill: "F5EEF1" }),
      fuente: ST.s({ sz: 8, c: C.muted, h: "center", bd: "b", bc: "EEE7EA" }),
      kpiL: ST.s({ c: C.ink, ind: 1, bd: "b", bc: "E5E7EB", v: "center" }), kpiV: ST.s({ b: 1, sz: 13, c: "000000", nf: EUR, bd: "b", bc: "E5E7EB", h: "right" }),
      kpiP: ST.s({ b: 1, sz: 13, c: "000000", nf: PCT, bd: "b", bc: "E5E7EB", h: "right" }), kpiS: ST.s({ b: 1, sz: 13, c: "000000", nf: PCTS, bd: "b", bc: "E5E7EB", h: "right" }),
      kpiT: ST.s({ b: 1, sz: 11, c: "000000", bd: "b", bc: "E5E7EB", h: "right" }), kpiE: ST.s({ sz: 9, c: C.muted, wrap: 1, bd: "b", bc: "E5E7EB", v: "center" }),
      input: ST.s({ b: 1, c: C.blue, fill: C.input, h: "center", bd: "lrtb", bc: "EAB308" }),
      date: ST.s({ c: C.ink, nf: "dd/mm/yyyy", h: "center", bd: "b", bc: "EEE7EA" }), t: ST.s({ c: C.ink, bd: "b", bc: "EEE7EA" }), tC: ST.s({ c: C.ink, h: "center", bd: "b", bc: "EEE7EA" }),
      year: ST.s({ b: 1, c: C.ink, h: "center", bd: "b", bc: "EEE7EA" }), legend: ST.s({ sz: 9, c: C.ink, wrap: 1 }),
      sw: function (fill) { return ST.s({ fill: fill }); }
    };
    var DX = { red: ST.dxf("9B1C1C", "FDE2E2", 1), green: ST.dxf("166534", "DCFCE7", 0) };
    var sheets = [];
    var conDatos = M.meses.map(function (x) { return x.datos; });
    var ultimo = -1; conDatos.forEach(function (d, i) { if (d) ultimo = i; });
    var nDatos = conDatos.filter(Boolean).length;
    var periodo = ultimo < 0 ? "sin datos todavía" : (nDatos === 12 ? "año completo" : "de enero a " + MESES_L[ultimo].toLowerCase());

    // ═ MES A MES ═ (la base de casi todo)
    var MM = new Sheet("MES A MES"); MM.tab = "7A4A5A"; MM.grid = false;
    [14, 14, 14, 14, 10, 15, 15, 13, 16].forEach(function (w, i) { MM.width(i + 1, w); });
    MM.put(1, 1, "Mes a mes " + year, S.title);
    MM.put(2, 1, "Lo que entra (ingresos), lo que sale (gastos) y lo que queda (beneficio) cada mes. Azul = dato; negro = cálculo.", S.sub);
    var mh = 4;
    ["Mes", "Ingresos", "Gastos", "Beneficio", "Margen", "Beneficio acumulado", "Ingresos " + (year - 1), "Ingresos vs " + (year - 1), "Datos de"].forEach(function (h, i) { MM.put(mh, i + 1, h, S.head); });
    MM.heights[mh] = 32; MM.freeze = [mh + 1, 2];
    var r0 = mh + 1, acum = 0, benef = [], acumV = [];
    for (var m = 0; m < 12; m++) {
      var r = r0 + m, ing = M.ing[m], gas = M.gas[m], d = conDatos[m];
      MM.put(r, 1, MESES_L[m], S.t);
      MM.put(r, 2, ing, S.in); MM.put(r, 3, gas, S.in);
      var b = d ? r2((ing || 0) - (gas || 0)) : null; benef[m] = b;
      MM.put(r, 4, b, S.f, d ? "B" + r + "-C" + r : null);
      MM.put(r, 5, d && ing > 0 ? b / ing : null, S.pct, d ? 'IF(B' + r + '>0,D' + r + '/B' + r + ',"")' : null);
      if (d) acum = r2(acum + b); acumV[m] = d ? acum : null;
      MM.put(r, 6, acumV[m], S.f, d ? "SUM($D$" + r0 + ":D" + r + ")" : null);
      var pi = P.ing[m]; MM.put(r, 7, pi, S.in);
      MM.put(r, 8, d && pi > 0 && ing != null ? ing / pi - 1 : null, S.pctS, d && pi > 0 ? 'IF(AND(B' + r + '<>"",G' + r + '>0),B' + r + '/G' + r + '-1,"")' : null);
      var mes = M.meses[m];
      MM.put(r, 9, mes.src === "hist" ? (d ? "Tu Excel antiguo" : "Sin datos") : mes.src === "app" ? (mes.actual ? "App (mes en curso)" : "App") : "Aún no", S.fuente);
    }
    var rT = r0 + 12, rMed = rT + 1;
    var tIng = r2(sum(M.ing)), tGas = r2(sum(M.gas)), tBen = r2(tIng - tGas), tPrev = r2(sum(P.ing.filter(function (v, i) { return conDatos[i]; })));
    MM.put(rT, 1, "Total año", S.labelB);
    MM.put(rT, 2, tIng, S.fB, "SUM(B" + r0 + ":B" + (rT - 1) + ")"); MM.put(rT, 3, tGas, S.fB, "SUM(C" + r0 + ":C" + (rT - 1) + ")");
    MM.put(rT, 4, tBen, S.fB, "B" + rT + "-C" + rT); MM.put(rT, 5, tIng > 0 ? tBen / tIng : null, S.pctB, 'IF(B' + rT + '>0,D' + rT + '/B' + rT + ',"")');
    MM.put(rT, 6, null, S.fB); MM.put(rT, 7, r2(sum(P.ing)), S.fB, "SUM(G" + r0 + ":G" + (rT - 1) + ")"); MM.put(rT, 8, null, S.fB); MM.put(rT, 9, null, S.fB);
    MM.put(rMed, 1, "Media al mes", S.label);
    MM.put(rMed, 2, nDatos ? tIng / nDatos : null, S.f, nDatos ? 'SUM(B' + r0 + ':B' + (rT - 1) + ')/' + nDatos : null);
    MM.put(rMed, 3, nDatos ? tGas / nDatos : null, S.f, nDatos ? 'SUM(C' + r0 + ':C' + (rT - 1) + ')/' + nDatos : null);
    MM.put(rMed, 4, nDatos ? tBen / nDatos : null, S.f, nDatos ? 'IFERROR(B' + rMed + '-C' + rMed + ',"")' : null);
    MM.put(rMed + 2, 1, "Cómo se lee: Beneficio = Ingresos − Gastos. Margen = qué parte de cada euro cobrado queda (20 % = de cada 100 € quedan 20 €). Beneficio acumulado = lo que llevas ganado en el año hasta ese mes.", S.small);
    MM.put(rMed + 3, 1, "«Ingresos vs " + (year - 1) + "»: +10 % = ese mes entró un 10 % más que el mismo mes del año pasado; en rojo, menos.", S.small);
    MM.cf.push({ ref: "D" + r0 + ":D" + rT + " F" + r0 + ":F" + (rT - 1), rules: [{ f: 'AND(ISNUMBER(D' + r0 + '),D' + r0 + '<0)', dxf: DX.red }] });
    sheets.push(MM);
    var refMM = function (c) { return q("MES A MES") + "!$" + c + "$" + r0 + ":$" + c + "$" + (rT - 1); };
    var catsMes = { ref: refMM("A"), vals: MESES_L.map(function (x) { return x.slice(0, 3); }) };

    // ═ GASTOS POR CATEGORÍA ═
    var G = new Sheet("GASTOS POR CATEGORÍA"); G.tab = "E08A2C"; G.grid = false;
    G.width(1, 36); for (var i = 2; i <= 13; i++) G.width(i, 10.5); G.width(14, 13); G.width(15, 12); G.width(16, 10); G.width(17, 13);
    G.put(1, 1, "En qué se va el dinero · " + year, S.title);
    G.put(2, 1, "Cada gasto del año en su categoría, agrupadas en 5 grupos. Total año, media al mes, qué parte del gasto total supone y cuánto fue el año pasado.", S.sub);
    var gh = 4; G.put(gh, 1, "Categoría", S.headL); for (i = 0; i < 12; i++) G.put(gh, 2 + i, MESES_C[i], S.head);
    G.put(gh, 14, "Total " + year, S.head); G.put(gh, 15, "Media al mes", S.head); G.put(gh, 16, "% del gasto", S.head); G.put(gh, 17, "Total " + (year - 1), S.head);
    G.heights[gh] = 30; G.freeze = [gh + 1, 2];
    var gr = gh + 1, grupoFila = {}, grupoVals = {}, filasCat = [], totalGas = tGas;
    function catPrev(k) { if (P.catAnual) return P.catAnual[k] || 0; return r2(sum(P.cat[k] || [])); }
    GRUPOS.forEach(function (g) {
      var cs = CATS.filter(function (c) { return c[2] === g.id; });
      var hr = gr; grupoFila[g.id] = hr; gr++;
      var first = gr;
      cs.forEach(function (c) {
        var vals = M.cat[c[0]];
        G.put(gr, 1, c[1], S.label);
        for (var m = 0; m < 12; m++) G.put(gr, 2 + m, vals[m], S.in);
        var t = r2(sum(vals));
        G.put(gr, 14, t, S.fB, "SUM(B" + gr + ":M" + gr + ")");
        G.put(gr, 15, nDatos ? t / nDatos : null, S.f, nDatos ? "N" + gr + "/" + nDatos : null);
        G.put(gr, 16, totalGas ? t / totalGas : null, S.pct, 'IFERROR(N' + gr + '/{TOT},"")');
        G.put(gr, 17, catPrev(c[0]) || null, S.in);
        filasCat.push(gr); gr++;
      });
      var last = gr - 1, gv = [];
      G.put(hr, 1, g.nombre, S.sect);
      for (var mm = 0; mm < 12; mm++) { var v = r2(sum(cs.map(function (c) { return M.cat[c[0]][mm]; }))); gv[mm] = conDatos[mm] ? v : null; G.put(hr, 2 + mm, gv[mm], S.sectN, conDatos[mm] ? "SUM(" + col(2 + mm) + first + ":" + col(2 + mm) + last + ")" : null); }
      var gt = r2(sum(gv)); grupoVals[g.id] = { meses: gv, total: gt, prev: r2(sum(cs.map(function (c) { return catPrev(c[0]); }))) };
      G.put(hr, 14, gt, S.sectN, "SUM(B" + hr + ":M" + hr + ")"); G.put(hr, 15, nDatos ? gt / nDatos : null, S.sectN, nDatos ? "N" + hr + "/" + nDatos : null);
      G.put(hr, 16, totalGas ? gt / totalGas : null, S.sectP, 'IFERROR(N' + hr + '/{TOT},"")'); G.put(hr, 17, grupoVals[g.id].prev, S.sectN, "SUM(Q" + first + ":Q" + last + ")");
    });
    // Diferencia con el total del mes (gastos sin categoría en el Excel antiguo)
    var hayDif = M.sinCat.some(function (v) { return v && Math.abs(v) >= 0.005; });
    var rSin = null;
    if (hayDif) {
      rSin = gr; G.put(gr, 1, M.sinDesglose ? "Sin desglose por mes (tu Excel antiguo solo tiene el total del año)" : "Sin clasificar (diferencia en tu Excel antiguo)", S.label);
      for (m = 0; m < 12; m++) G.put(gr, 2 + m, M.sinCat[m] && Math.abs(M.sinCat[m]) >= 0.005 ? M.sinCat[m] : null, S.in);
      var ts = r2(sum(M.sinCat.filter(function (v) { return v && Math.abs(v) >= 0.005; })));
      G.put(gr, 14, ts, S.fB, "SUM(B" + gr + ":M" + gr + ")"); G.put(gr, 15, nDatos ? ts / nDatos : null, S.f, nDatos ? "N" + gr + "/" + nDatos : null);
      G.put(gr, 16, totalGas ? ts / totalGas : null, S.pct, 'IFERROR(N' + gr + '/{TOT},"")'); gr++;
    }
    var rGT = gr; G.put(rGT, 1, "TOTAL GASTOS", S.labelB);
    var sumRef = function (c) { return GRUPOS.map(function (g) { return c + grupoFila[g.id]; }).concat(rSin ? [c + rSin] : []).join("+"); };
    for (m = 0; m < 12; m++) G.put(rGT, 2 + m, conDatos[m] ? M.gas[m] : null, S.fB, conDatos[m] ? sumRef(col(2 + m)) : null);
    G.put(rGT, 14, totalGas, S.fB, sumRef("N")); G.put(rGT, 15, nDatos ? totalGas / nDatos : null, S.fB, nDatos ? "N" + rGT + "/" + nDatos : null);
    G.put(rGT, 16, totalGas ? 1 : null, S.pctB, 'IFERROR(N' + rGT + '/N' + rGT + ',"")');
    G.put(rGT, 17, r2(sum(GRUPOS.map(function (g) { return grupoVals[g.id].prev; }))), S.fB, GRUPOS.map(function (g) { return "Q" + grupoFila[g.id]; }).join("+"));
    Object.keys(G.cells).forEach(function (k) { var x = G.cells[k]; if (x.f && x.f.indexOf("{TOT}") >= 0) x.f = x.f.replace("{TOT}", "$N$" + rGT); });
    G.put(rGT + 2, 1, "Cómo se lee: la fila de color de cada grupo es la suma de sus categorías. «% del gasto» = qué parte de todo lo gastado en el año se fue en esa línea.", S.small);
    if (P.catAnual) G.put(rGT + 3, 1, "El total " + (year - 1) + " por categoría viene de los totales anuales de tu Excel antiguo.", S.small);
    // Gráfica: gastos por grupo y mes (columnas apiladas)
    var gRows = GRUPOS.map(function (g) { return grupoFila[g.id]; });
    G.charts.push({ at: [rGT + 4, 0, rGT + 24, 13], xml: chartXml({ title: "Gastos de cada mes, por grupo", cats: catsMes, stacked: 1, gap: 50,
      series: GRUPOS.map(function (g, k) { return { name: g.nombre, nameRef: q(G.name) + "!$A$" + gRows[k], ref: q(G.name) + "!$B$" + gRows[k] + ":$M$" + gRows[k], vals: grupoVals[g.id].meses, color: g.color }; }) }) });
    sheets.push(G);

    // ═ HISTÓRICO ═
    var Hs = new Sheet("HISTÓRICO"); Hs.tab = "94A3B8"; Hs.grid = false;
    var ys = aniosDisponibles(o);
    Hs.width(1, 16); for (i = 0; i < ys.length; i++) Hs.width(2 + i, 13);
    Hs.width(2 + ys.length, 3);
    Hs.put(1, 1, "Los años comparados", S.title);
    Hs.put(2, 1, "Cada año con los meses que tienen datos. Sirve para ver si el negocio crece y qué meses suelen ser flojos.", S.sub);
    var yh = 4; Hs.put(yh, 1, "", S.headL); ys.forEach(function (y, k) { Hs.put(yh, 2 + k, String(y), S.head); });
    var filasY = [["Ingresos", "ing"], ["Gastos", "gas"], ["Beneficio", "ben"], ["Margen", "mar"], ["Meses con datos", "n"]];
    var Ymod = {}; ys.forEach(function (y) { Ymod[y] = y === year ? M : (y === year - 1 ? P : modelo(o, y)); });
    var resY = {};
    ys.forEach(function (y) { var X = Ymod[y], n = X.meses.filter(function (x) { return x.datos; }).length, a = r2(sum(X.ing)), g = r2(sum(X.gas)); resY[y] = { ing: a, gas: g, ben: r2(a - g), n: n }; });
    filasY.forEach(function (fy, k) {
      var rr = yh + 1 + k; Hs.put(rr, 1, fy[0], k === 2 ? S.labelB : S.label);
      ys.forEach(function (y, j) {
        var c = col(2 + j), R = resY[y];
        if (fy[1] === "ing" || fy[1] === "gas") Hs.put(rr, 2 + j, R[fy[1]], S.f, "SUM(" + c + (yh + 9) + ":" + c + (yh + 20) + ")".replace("", "") );
        else if (fy[1] === "ben") Hs.put(rr, 2 + j, R.ben, S.fB, c + (yh + 1) + "-" + c + (yh + 2));
        else if (fy[1] === "mar") Hs.put(rr, 2 + j, R.ing ? R.ben / R.ing : null, S.pct, 'IFERROR(' + c + (yh + 3) + '/' + c + (yh + 1) + ',"")');
        else Hs.put(rr, 2 + j, R.n, S.year);
      });
    });
    // Tablas mes a mes por año (ingresos y gastos)
    var ti = yh + 8; Hs.put(ti - 1, 1, "Ingresos de cada mes", S.h2);
    Hs.put(ti, 1, "Mes", S.headL); ys.forEach(function (y, k) { Hs.put(ti, 2 + k, String(y), S.head); });
    for (m = 0; m < 12; m++) { Hs.put(ti + 1 + m, 1, MESES_L[m], S.t); ys.forEach(function (y, k) { Hs.put(ti + 1 + m, 2 + k, Ymod[y].ing[m], S.in); }); }
    var tg = ti + 15; Hs.put(tg - 1, 1, "Gastos de cada mes", S.h2);
    Hs.put(tg, 1, "Mes", S.headL); ys.forEach(function (y, k) { Hs.put(tg, 2 + k, String(y), S.head); });
    for (m = 0; m < 12; m++) { Hs.put(tg + 1 + m, 1, MESES_L[m], S.t); ys.forEach(function (y, k) { Hs.put(tg + 1 + m, 2 + k, Ymod[y].gas[m], S.in); }); }
    // fórmulas de los totales anuales apuntan a esas tablas
    ys.forEach(function (y, j) { var c = col(2 + j); Hs.cells[(yh + 1) + "," + (2 + j)].f = "SUM(" + c + (ti + 1) + ":" + c + (ti + 12) + ")"; Hs.cells[(yh + 2) + "," + (2 + j)].f = "SUM(" + c + (tg + 1) + ":" + c + (tg + 12) + ")"; });
    var chH = ys.length + 3;
    Hs.charts.push({ at: [yh - 1, chH, yh + 17, chH + 9], xml: chartXml({ title: "Cada año: lo que entra, lo que sale y lo que queda", cats: { ref: q(Hs.name) + "!$B$" + yh + ":$" + col(1 + ys.length) + "$" + yh, vals: ys.map(String) }, gap: 80,
      series: [
        { name: "Ingresos", nameRef: q(Hs.name) + "!$A$" + (yh + 1), ref: q(Hs.name) + "!$B$" + (yh + 1) + ":$" + col(1 + ys.length) + "$" + (yh + 1), vals: ys.map(function (y) { return resY[y].ing; }), color: "7A4A5A" },
        { name: "Gastos", nameRef: q(Hs.name) + "!$A$" + (yh + 2), ref: q(Hs.name) + "!$B$" + (yh + 2) + ":$" + col(1 + ys.length) + "$" + (yh + 2), vals: ys.map(function (y) { return resY[y].gas; }), color: "D4A5B5" },
        { name: "Beneficio", nameRef: q(Hs.name) + "!$A$" + (yh + 3), ref: q(Hs.name) + "!$B$" + (yh + 3) + ":$" + col(1 + ys.length) + "$" + (yh + 3), vals: ys.map(function (y) { return resY[y].ben; }), color: "2E7D5B", labels: 1 }] }) });
    var palY = ["CBD5E1", "94A3B8", "D4A5B5", "7A4A5A", "C2185B", "4A2735"];
    Hs.charts.push({ at: [yh + 18, chH, yh + 36, chH + 9], xml: chartXml({ title: "Ingresos de cada mes, un color por año", cats: { ref: q(Hs.name) + "!$A$" + (ti + 1) + ":$A$" + (ti + 12), vals: MESES_L.map(function (x) { return x.slice(0, 3); }) },
      series: ys.map(function (y, k) { var c = col(2 + k); return { kind: "line", name: String(y), nameRef: q(Hs.name) + "!$" + c + "$" + ti, ref: q(Hs.name) + "!$" + c + "$" + (ti + 1) + ":$" + c + "$" + (ti + 12), vals: Ymod[y].ing, color: palY[(((palY.length - ys.length + k) % palY.length) + palY.length) % palY.length] }; }) }) });
    sheets.push(Hs);

    // ═ GASTOS DETALLE (lo apuntado en la app) ═
    var D = new Sheet("GASTOS DETALLE"); D.tab = "6B9E78";
    [12, 30, 20, 30, 34, 12, 12, 12].forEach(function (w, i) { D.width(i + 1, w); });
    D.put(1, 1, "Gastos apuntados en la app · " + year, S.title);
    D.put(2, 1, "Cada factura o recibo guardado en la app, con la categoría que usa este Excel. Los meses anteriores a la app solo tienen totales (hoja GASTOS POR CATEGORÍA).", S.sub);
    ["Fecha", "Proveedor", "Categoría en la app", "Categoría en este Excel", "Concepto / nº", "Base", "IVA", "Total"].forEach(function (h, i) { D.put(4, i + 1, h, S.head); });
    D.freeze = [5, 1]; D.heights[4] = 28;
    var lista = M.gastosApp.slice().sort(function (a, b) { return String(a.fecha).localeCompare(String(b.fecha)); });
    lista.forEach(function (g, k) {
      var rr = 5 + k, f = String(g.fecha || "");
      var serial = f ? (Date.UTC(+f.slice(0, 4), +f.slice(5, 7) - 1, +f.slice(8, 10) || 1) - Date.UTC(1899, 11, 30)) / 864e5 : null;
      D.put(rr, 1, serial, S.date); D.put(rr, 2, g.proveedor || "", S.t); D.put(rr, 3, g.categoria || "", S.t);
      var k2 = catDeGasto(g); D.put(rr, 4, CATS[CAT_IDX[k2]][1], S.t); D.put(rr, 5, g.numero || g.concepto || "", S.t);
      D.put(rr, 6, typeof g.base === "number" ? g.base : null, S.in); D.put(rr, 7, typeof g.iva === "number" ? g.iva : null, S.in); D.put(rr, 8, Number(g.total) || 0, S.in);
    });
    if (!lista.length) D.put(5, 1, "Todavía no hay gastos de la app en " + year + ".", S.sub);
    else { var rr2 = 5 + lista.length; D.put(rr2, 7, "Total", S.labelB); D.put(rr2, 8, r2(sum(lista.map(function (g) { return Number(g.total) || 0; }))), S.fB, "SUM(H5:H" + (rr2 - 1) + ")"); }
    sheets.push(D);

    // ═ RESUMEN ═ (lo primero que se mira)
    var R = new Sheet("RESUMEN " + year); R.tab = "C2185B"; R.grid = false;
    R.width(1, 38); R.width(2, 16); R.width(3, 52);
    R.put(1, 1, "Belleza Lua · Resumen " + year, S.title);
    R.put(2, 1, "Actualizado el " + (o.generado || "") + " · " + (ultimo < 0 ? "todavía sin datos" : "datos " + periodo) + ". Cambia solo el número amarillo; lo demás se calcula.", S.sub);
    R.put(4, 1, "Lo importante del año", S.h2);
    var mmB = q("MES A MES") + "!", kr = 5;
    var mejor = -1, peor = -1;
    for (m = 0; m < 12; m++) if (conDatos[m] && M.ing[m] != null) { if (mejor < 0 || M.ing[m] > M.ing[mejor]) mejor = m; if (peor < 0 || M.ing[m] < M.ing[peor]) peor = m; }
    var mediaG = nDatos ? tGas / nDatos : 0, dias = 22;
    var prevG = r2(sum(P.gas.filter(function (v, i) { return conDatos[i]; })));
    var kpis = [
      ["Ingresos del año", tIng, S.kpiV, mmB + "B" + rT, "Todo lo cobrado (" + periodo + "), IVA incluido."],
      ["Gastos del año", tGas, S.kpiV, mmB + "C" + rT, "Todo lo pagado: facturas, alquiler, autónomo e impuestos."],
      ["Beneficio (lo que queda)", tBen, S.kpiV, mmB + "D" + rT, "Ingresos − gastos. Es lo que queda antes de la declaración de la renta."],
      ["Margen", tIng > 0 ? tBen / tIng : null, S.kpiP, mmB + "E" + rT, "De cada 100 € cobrados, cuántos quedan."],
      ["Ingresos medios al mes", nDatos ? tIng / nDatos : null, S.kpiV, mmB + "B" + rMed, "Media de los meses con datos."],
      ["Gastos medios al mes", nDatos ? mediaG : null, S.kpiV, mmB + "C" + rMed, "Media de los meses con datos."],
      ["Para cubrir gastos hay que facturar al mes", nDatos ? mediaG : null, S.kpiV, "B10", "Lo mínimo para no perder dinero: igual a los gastos medios."],
      ["Días de trabajo al mes", dias, S.input, null, "← Cámbialo si trabajáis más o menos días."],
      ["Para cubrir gastos hay que facturar al día", nDatos ? mediaG / dias : null, S.kpiV, 'IFERROR(B11/B12,"")', "Gastos medios al mes ÷ días de trabajo."],
      ["Ingresos frente a " + (year - 1) + " (mismos meses)", tPrev ? tIng / tPrev - 1 : null, S.kpiS, tPrev ? "IFERROR(" + tIng + "/" + tPrev + "-1,\"\")" : null, tPrev ? "Mismos meses del año pasado: " + eurTxt(tPrev) + ". En rojo si se ha facturado menos." : "No hay datos del año pasado para comparar."],
      ["Gastos frente a " + (year - 1) + " (mismos meses)", prevG ? tGas / prevG - 1 : null, S.kpiS, prevG ? "IFERROR(" + tGas + "/" + prevG + "-1,\"\")" : null, prevG ? "Mismos meses del año pasado: " + eurTxt(prevG) + "." : "No hay datos del año pasado para comparar."],
      ["Mejor mes", mejor >= 0 ? MESES_L[mejor] + " · " + eurTxt(M.ing[mejor]) : "—", S.kpiT, null, "El mes con más ingresos."],
      ["Mes más flojo", peor >= 0 ? MESES_L[peor] + " · " + eurTxt(M.ing[peor]) : "—", S.kpiT, null, "El mes con menos ingresos (vacaciones, cierres…)."]
    ];
    kpis.forEach(function (k, j) {
      var rr = kr + j; R.put(rr, 1, k[0], S.kpiL); R.put(rr, 2, k[1], k[2], k[3] || null); R.put(rr, 3, k[4], S.kpiE); R.heights[rr] = 22;
    });
    // Fila 10 = "Para cubrir gastos al mes" (kr+6), fila 11 = días (kr+7) y 12 = al día: ajustar referencias exactas
    var rCub = kr + 6, rDias = kr + 7, rDia = kr + 8;
    R.cells[rCub + ",2"].f = nDatos ? mmB + "C" + rMed : null;
    if (!nDatos) { [kr + 4, kr + 5].forEach(function (rr) { if (R.cells[rr + ",2"]) R.cells[rr + ",2"].f = null; }); }
    R.cells[rDia + ",2"].f = nDatos ? 'IFERROR(B' + rCub + '/B' + rDias + ',"")' : null;
    R.cf.push({ ref: "B" + (kr + 2), rules: [{ f: "B" + (kr + 2) + "<0", dxf: DX.red }, { f: "B" + (kr + 2) + ">=0", dxf: DX.green }] });
    var rG = kr + kpis.length + 1;
    R.put(rG, 1, "Gráficas", S.h2);
    R.put(rG + 1, 1, "1) Cada mes: barras = ingresos y gastos; línea verde = beneficio. 2) En qué se va el dinero del año, por grupos. 3) Los años comparados (en la hoja HISTÓRICO).", S.small);
    R.charts.push({ at: [rG + 2, 0, rG + 22, 3], xml: chartXml({ title: "Cada mes: ingresos, gastos y beneficio", cats: catsMes, gap: 40,
      series: [
        { name: "Ingresos", nameRef: mmB + "$B$" + mh, ref: refMM("B"), vals: M.ing, color: "7A4A5A" },
        { name: "Gastos", nameRef: mmB + "$C$" + mh, ref: refMM("C"), vals: M.gas, color: "D4A5B5" },
        { kind: "line", name: "Beneficio", nameRef: mmB + "$D$" + mh, ref: refMM("D"), vals: benef, color: "2E7D5B" }] }) });
    var gq = q("GASTOS POR CATEGORÍA") + "!";
    R.charts.push({ at: [rG + 23, 0, rG + 40, 3], xml: chartXml({ title: "En qué se va el dinero en " + year + " (total del año)", dir: "bar", gap: 40,
      cats: { ref: gq + "$A$" + gRows[0] + ":$A$" + gRows[gRows.length - 1], vals: GRUPOS.map(function (g) { return g.nombre; }) },
      series: [{ name: "Total " + year, nameRef: gq + "$N$" + gh, ref: gq + "$N$" + gRows[0] + ":$N$" + gRows[gRows.length - 1], vals: GRUPOS.map(function (g) { return grupoVals[g.id].total; }), color: "C2185B", labels: 1 }] }) });
    // Ojo: la referencia A{primer grupo}:A{último grupo} incluye categorías entre medias; usamos una tabla propia de grupos:
    var rGr = rG + 42; R.put(rGr, 1, "Gastos por grupo (para la gráfica)", S.h2);
    R.put(rGr + 1, 1, "Grupo", S.headL); R.put(rGr + 1, 2, "Total " + year, S.head); R.put(rGr + 1, 3, "Qué incluye", S.headL);
    GRUPOS.forEach(function (g, k) {
      var rr = rGr + 2 + k; R.put(rr, 1, g.nombre, S.label); R.put(rr, 2, grupoVals[g.id].total, S.f, gq + "N" + grupoFila[g.id]);
      R.put(rr, 3, CATS.filter(function (c) { return c[2] === g.id; }).map(function (c) { return c[1]; }).join(", "), S.kpiE); R.heights[rr] = 26;
    });
    R.charts[1].xml = chartXml({ title: "En qué se va el dinero en " + year + " (total del año)", dir: "bar", gap: 40,
      cats: { ref: q(R.name) + "!$A$" + (rGr + 2) + ":$A$" + (rGr + 1 + GRUPOS.length), vals: GRUPOS.map(function (g) { return g.nombre; }) },
      series: [{ name: "Total " + year, nameRef: q(R.name) + "!$B$" + (rGr + 1), ref: q(R.name) + "!$B$" + (rGr + 2) + ":$B$" + (rGr + 1 + GRUPOS.length), vals: GRUPOS.map(function (g) { return grupoVals[g.id].total; }), color: "C2185B", labels: 1 }] });

    // ═ EMPIEZA AQUÍ ═
    var E = new Sheet("EMPIEZA AQUÍ"); E.tab = "4A2735"; E.grid = false;
    E.width(1, 28); E.width(2, 62); E.width(3, 34);
    var er = 1;
    function line(a, b, c, sa, sb, sc) { E.put(er, 1, a, sa || S.txtB); if (b != null) E.put(er, 2, b, sb || S.txt); if (c != null) E.put(er, 3, c, sc || S.txt); er++; }
    E.put(er, 1, "Belleza Lua · Cuentas del año " + year, S.title); er++;
    E.put(er, 1, "Generado por la app Belleza Lua el " + (o.generado || "") + ". Para tener la versión al día, vuelve a generarlo desde la app.", S.sub); er += 2;
    E.put(er, 1, "Qué es este archivo", S.h2); er++;
    E.put(er, 1, "Las cuentas del negocio del año en un solo sitio: cuánto entra, cuánto sale, cuánto queda y en qué se va el dinero, comparado con los años anteriores. No sustituye a la gestoría: es para entender cómo va el negocio.", S.txt); E.merges.push("A" + er + ":C" + er); E.heights[er] = 32; er += 2;
    E.put(er, 1, "Las hojas, en orden", S.h2); er++;
    line("Hoja", "Para qué sirve", "Mírala cuando…", S.headL, S.headL, S.headL);
    [["RESUMEN " + year, "Las cifras clave del año y 2 gráficas. Lo primero que hay que mirar.", "Quieras saber en 1 minuto cómo va el año."],
     ["MES A MES", "Ingresos, gastos y beneficio de cada mes, y la comparación con el año pasado.", "Quieras ver qué mes fue bueno o malo."],
     ["GASTOS POR CATEGORÍA", "Cada tipo de gasto mes a mes, en 5 grupos, con su % del total.", "Quieras saber en qué se va el dinero o qué ha subido."],
     ["HISTÓRICO", "Los años comparados (desde 2023) y los meses de cada año.", "Quieras ver si el negocio crece y qué meses suelen ser flojos."],
     ["GASTOS DETALLE", "Cada factura o recibo guardado en la app este año.", "Quieras buscar un gasto concreto."]].forEach(function (x) { line(x[0], x[1], x[2], S.txtB, S.txt, S.txt); E.heights[er - 1] = 30; });
    er++;
    E.put(er, 1, "Palabras que salen", S.h2); er++;
    [["Ingresos", "Lo cobrado a las clientas según el talonario, IVA incluido (como en tu Excel de siempre)."],
     ["Gastos", "Todo lo pagado: facturas de proveedores, alquiler, luz, gestoría, cuota de autónomo e impuestos (modelos 115, 130, 303)."],
     ["Beneficio", "Ingresos − gastos: lo que queda en el mes o en el año. No es el sueldo neto: falta la declaración de la renta."],
     ["Margen", "Beneficio ÷ ingresos. Un 30 % quiere decir que de cada 100 € cobrados quedan 30 €."],
     ["Para cubrir gastos", "Lo que hay que facturar como mínimo para no perder dinero (los gastos medios). Al día = dividido entre los días de trabajo."],
     ["Grupos de gasto", "El local · Producto para vender · Trabajo y equipo · Impuestos y autónomo · Servicios y otros."]].forEach(function (x) { line(x[0], x[1], null); E.heights[er - 1] = 30; E.merges.push("B" + (er - 1) + ":C" + (er - 1)); });
    er++;
    E.put(er, 1, "Colores", S.h2); er++;
    E.put(er, 1, "Número azul", ST.s({ b: 1, c: C.blue })); E.put(er, 2, "Dato que viene de la app o de tu Excel antiguo.", S.txt); er++;
    E.put(er, 1, "Número negro", ST.s({ b: 1, c: "000000" })); E.put(er, 2, "Cálculo hecho por el Excel (no hace falta tocarlo).", S.txt); er++;
    E.put(er, 1, "Casilla amarilla", S.input); E.put(er, 2, "Lo único que se puede cambiar a mano (los días de trabajo al mes).", S.txt); er++;
    E.put(er, 1, "Rojo", ST.s({ b: 1, c: "9B1C1C", fill: "FDE2E2" })); E.put(er, 2, "Beneficio negativo o menos que el año anterior.", S.txt); er += 2;
    E.put(er, 1, "Cómo se actualiza", S.h2); er++;
    ["1. En la app: pestaña Resumen → elegir el año con las flechas → «Actualizar Excel anual».", "2. Guardar el archivo en Archivos (iPhone) o abrirlo con Excel. Sustituye al del mes anterior.", "3. Hazlo cada mes después de sacar el Excel de Declarando."].forEach(function (t) { E.put(er, 1, t, S.txt); E.merges.push("A" + er + ":C" + er); er++; });
    er++;
    E.put(er, 1, "De dónde salen los datos", S.h2); er++;
    var H = o.hist || {};
    var desdeTxt = H.desde ? MESES_L[+H.desde.slice(5, 7) - 1].toLowerCase() + " de " + H.desde.slice(0, 4) : null;
    E.put(er, 1, H.anios ? "Hasta " + (desdeTxt ? "antes de " + desdeTxt : "que empieza la app") + ": tu Excel antiguo (" + (H.fuente || "histórico") + "). Desde " + (desdeTxt || "entonces") + ": lo apuntado en la app." : "Todo sale de lo apuntado en la app (no se ha cargado el histórico de años anteriores).", S.txt); E.merges.push("A" + er + ":C" + er); E.heights[er] = 30; er++;
    var huecos = []; ys.forEach(function (y) { Ymod[y].meses.forEach(function (x, i) { if (x.src === "hist" && !x.datos) huecos.push(MESES_C[i] + " " + y); }); });
    if (huecos.length) { E.put(er, 1, "Meses sin datos en el Excel antiguo: " + huecos.join(", ") + ". Salen vacíos.", S.txt); E.merges.push("A" + er + ":C" + er); E.heights[er] = 30; er++; }
    var difs = []; ys.forEach(function (y) { Ymod[y].sinCat.forEach(function (v, i) { if (v && Math.abs(v) >= 0.05 && Ymod[y].meses[i].src === "hist") difs.push(MESES_C[i] + " " + y + " (" + (v > 0 ? "+" : "") + String(v).replace(".", ",") + " €)"); }); });
    if (difs.length) { E.put(er, 1, "En tu Excel antiguo, el gasto total de algunos meses no coincide con la suma de sus categorías: " + difs.join(", ") + ". En GASTOS POR CATEGORÍA esa diferencia sale como «Sin clasificar» para que los totales sean los tuyos.", S.txt); E.merges.push("A" + er + ":C" + er); E.heights[er] = 44; er++; }

    var orden = [E, R, MM, G, Hs, D];
    return { bytes: pack(orden, ST, "Belleza Lua · Cuentas " + year), resumen: { ingresos: tIng, gastos: tGas, beneficio: tBen, meses: nDatos, ultimo: ultimo } };
  }

  var api = { build: build, modelo: modelo, CATS: CATS, GRUPOS: GRUPOS, catDeGasto: catDeGasto };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ExcelAnual = api;
})(typeof window !== "undefined" ? window : globalThis);

/* xl-lib.js — lector/escritor .xlsx sin dependencias + importador de presupuestos
   Funciona en navegador (DecompressionStream) y en Node >= 18. */
(function (root) {
'use strict';

var OLS = ['RDV', 'TL', 'TOL', 'ST'];
var ITEMS = ['Mantenimiento', 'Neumaticos', 'Combustible', 'Multas', 'Otros'];
var ITEM_LABEL = { Mantenimiento: 'Mantenimiento', Neumaticos: 'Neumáticos', Combustible: 'Combustible', Multas: 'Multas', Otros: 'Otros' };
var MONTHS_FULL = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
var MONTHS_SHORT = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
/* Proporción usada solo cuando un OL trae únicamente totales (caso ST) */
var ST_SPLIT = { Mantenimiento: 0.0991838290169045, Neumaticos: 0.010759828929806734, Combustible: 0.5409404056511602, Multas: 0.006302534406874315, Otros: 0.3428134019952543 };

function norm(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}
function unesc(s) {
  return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, function (_, n) { return String.fromCodePoint(+n); })
    .replace(/&#x([0-9a-fA-F]+);/g, function (_, n) { return String.fromCodePoint(parseInt(n, 16)); })
    .replace(/&amp;/g, '&');
}
function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function attr(tag, name) {
  var m = new RegExp('(?:^|\\s)' + name.replace(':', '\\:') + '="([^"]*)"').exec(tag);
  return m ? m[1] : null;
}

/* ---------------------------------------------------------------- ZIP (lectura) */
function u16(d, o) { return d[o] | (d[o + 1] << 8); }
function u32(d, o) { return (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) >>> 0; }
async function inflateRaw(u8) {
  if (typeof DecompressionStream === 'undefined') throw new Error('Este navegador no puede leer .xlsx. Usá una versión reciente de Chrome, Edge o Firefox.');
  var ds = new DecompressionStream('deflate-raw');
  var stream = new Blob([u8]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function unzip(buf) {
  var d = new Uint8Array(buf), eocd = -1, i;
  for (i = d.length - 22; i >= Math.max(0, d.length - 66000); i--) { if (u32(d, i) === 0x06054b50) { eocd = i; break; } }
  if (eocd < 0) throw new Error('El archivo no es un .xlsx válido. Si es un .xls antiguo, guardalo como .xlsx desde Excel.');
  var n = u16(d, eocd + 10), p = u32(d, eocd + 16), files = {};
  for (var k = 0; k < n; k++) {
    if (u32(d, p) !== 0x02014b50) break;
    var method = u16(d, p + 10), csize = u32(d, p + 20), nlen = u16(d, p + 28), elen = u16(d, p + 30), clen = u16(d, p + 32), lho = u32(d, p + 42);
    var name = new TextDecoder().decode(d.subarray(p + 46, p + 46 + nlen));
    files[name] = { method: method, csize: csize, lho: lho };
    p += 46 + nlen + elen + clen;
  }
  return {
    names: Object.keys(files),
    get: async function (name) {
      var f = files[name]; if (!f) return null;
      var nl = u16(d, f.lho + 26), el = u16(d, f.lho + 28), start = f.lho + 30 + nl + el;
      var raw = d.subarray(start, start + f.csize);
      if (f.method === 0) return raw;
      if (f.method === 8) return inflateRaw(raw);
      throw new Error('Compresión ZIP no soportada.');
    }
  };
}

/* ---------------------------------------------------------------- XLSX (lectura) */
function colIdx(ref) {
  var m = /^([A-Z]+)/.exec(ref || ''); if (!m) return 0;
  var n = 0; for (var i = 0; i < m[1].length; i++) n = n * 26 + m[1].charCodeAt(i) - 64;
  return n - 1;
}
function parseSheet(xml, ss) {
  var rows = [], auto = 0, rm, rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  while ((rm = rowRe.exec(xml))) {
    var rattr = rm[1], rnS = attr('<x ' + rattr, 'r');
    var rn = rnS ? parseInt(rnS, 10) - 1 : auto; auto = rn + 1;
    var body = rm[2] || '', arr = [], cm, cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    while ((cm = cellRe.exec(body))) {
      var at = '<c ' + cm[1], ref = attr(at, 'r'), t = attr(at, 't'), inner = cm[2] || '', v = null;
      var vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
      if (t === 'inlineStr') {
        var s = '', x, tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g; while ((x = tre.exec(inner))) s += x[1]; v = unesc(s);
      } else if (vm) {
        var raw = vm[1];
        if (t === 's') v = ss[parseInt(raw, 10)];
        else if (t === 'str') v = unesc(raw);
        else if (t === 'b') v = raw === '1';
        else if (t === 'e') v = null;
        else { var num = parseFloat(raw); v = isNaN(num) ? unesc(raw) : num; }
      }
      if (v !== null && v !== undefined) arr[colIdx(ref)] = v;
    }
    rows[rn] = arr;
  }
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i]) rows[i] = [];
    for (var j = 0; j < rows[i].length; j++) if (rows[i][j] === undefined) rows[i][j] = null;
  }
  return rows;
}
async function readXlsx(buf) {
  var z = await unzip(buf), td = new TextDecoder('utf-8');
  async function txt(n) { var u = await z.get(n); return u ? td.decode(u) : null; }
  var wbx = await txt('xl/workbook.xml');
  if (!wbx) throw new Error('El archivo no parece un libro de Excel (.xlsx).');
  var rels = (await txt('xl/_rels/workbook.xml.rels')) || '', relMap = {}, m;
  var relRe = /<Relationship\b[^>]*>/g;
  while ((m = relRe.exec(rels))) { var id = attr(m[0], 'Id'); if (id) relMap[id] = attr(m[0], 'Target'); }
  var ss = [], sst = await txt('xl/sharedStrings.xml');
  if (sst) {
    var siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    while ((m = siRe.exec(sst))) {
      var inner = m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, ''), s = '', t, tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      while ((t = tre.exec(inner))) s += t[1];
      ss.push(unesc(s));
    }
  }
  var names = [], sheets = {}, shRe = /<sheet\b[^>]*>/g;
  while ((m = shRe.exec(wbx))) {
    var name = unesc(attr(m[0], 'name') || ''), tg = relMap[attr(m[0], 'r:id')];
    if (!tg) continue;
    tg = tg.replace(/^\//, ''); var path = tg.indexOf('xl/') === 0 ? tg : 'xl/' + tg;
    var xml = await txt(path); if (xml == null) continue;
    names.push(name); sheets[name] = parseSheet(xml, ss);
  }
  return { names: names, sheets: sheets };
}

/* ---------------------------------------------------------------- Helpers de datos */
function parseNum(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v).replace(/[$\s ]/g, '');
  if (!s || s === '-' || s === '–') return null;
  var neg = /^\(.*\)$/.test(s); s = s.replace(/[()]/g, '');
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) { s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, ''); }
  else if (s.indexOf(',') >= 0) { s = s.replace(',', '.'); }
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) { s = s.replace(/\./g, ''); }
  var n = parseFloat(s); if (!isFinite(n)) return null;
  return neg ? -n : n;
}
function monthIndex(v) {
  if (typeof v !== 'string') return -1;
  var n = norm(v), i = MONTHS_FULL.indexOf(n); if (i >= 0) return i;
  if (n.length <= 4) { var s = MONTHS_SHORT.indexOf(n.slice(0, 3)); if (s >= 0) return s; }
  return -1;
}
function monthCols(row) {
  var cols = Array(12).fill(null), count = 0;
  for (var c = 0; c < row.length; c++) { var mi = monthIndex(row[c]); if (mi >= 0 && cols[mi] == null) { cols[mi] = c; count++; } }
  return { cols: cols, count: count };
}
function itemOf(label) {
  var n = norm(label); if (!n) return null;
  if (/total|tatal|conceptos|desvio|diferencia|^flota|^importe/.test(n)) return null;
  if (/^(mantenimiento|mantto\.?|mant\.?)$/.test(n)) return { item: 'Mantenimiento', kind: 'total' };
  if (/correctiv|preventiv|proactiv|predictiv|repuesto|respuesto|mantenimiento|mantto/.test(n)) return { item: 'Mantenimiento', kind: 'sub' };
  if (/neumatic|cubierta/.test(n)) return { item: 'Neumaticos', kind: 'total' };
  if (/multa/.test(n)) return { item: 'Multas', kind: 'total' };
  if (/combustible/.test(n)) return { item: 'Combustible', kind: 'total' };
  if (/documentac|administrativ|varios|seguridad|insumo|gps|extintor|patente|vtv|peaje|habilitac|sistema de gestion|^otros/.test(n)) return { item: 'Otros', kind: 'sub' };
  return null;
}
function newAcc() { var a = {}; ITEMS.forEach(function (i) { a[i] = { tot: null, sub: null }; }); return a; }
function addTo(acc, item, kind, vals) {
  var key = kind === 'total' ? 'tot' : 'sub', slot = acc[item];
  if (!slot[key]) slot[key] = Array(12).fill(null);
  for (var m = 0; m < 12; m++) { if (vals[m] == null) continue; slot[key][m] = (slot[key][m] || 0) + vals[m]; }
}
function finalize(acc) {
  var out = {};
  ITEMS.forEach(function (it) {
    var s = acc[it], arr = Array(12).fill(null), any = false;
    for (var m = 0; m < 12; m++) {
      var v = (s.tot && s.tot[m] != null) ? s.tot[m] : (s.sub ? s.sub[m] : null);
      if (v != null) { arr[m] = v; any = true; }
    }
    if (any) out[it] = arr;
  });
  return out;
}
function rowVals(row, cols) { return cols.map(function (c) { return c == null ? null : parseNum(row[c]); }); }
function score(o) {
  var s = 0; ['presu', 'real'].forEach(function (t) { var x = o[t] || {}; Object.keys(x).forEach(function (k) { x[k].forEach(function (v) { s += Math.abs(v || 0); }); }); });
  return s;
}
function nonEmpty(o) { return o && (Object.keys(o.presu || {}).length + Object.keys(o.real || {}).length) > 0; }

/* ---------------------------------------------------------------- Parsers */
/* 1) Plantilla estándar: columnas OL | Ítem | Ene … Dic */
function parseTemplateSheet(rows, kind, warnings, sheetName) {
  var h = -1, mc = null;
  for (var r = 0; r < Math.min(rows.length, 30); r++) {
    var row = rows[r] || [], a = norm(row[0]);
    if (/^(ol|operador|operador logistico|ol \/ operador)$/.test(a)) { var c = monthCols(row); if (c.count >= 6) { h = r; mc = c.cols; break; } }
  }
  if (h < 0) return null;
  var out = {}, accs = {};
  for (r = h + 1; r < rows.length; r++) {
    row = rows[r] || [];
    var olRaw = row[0], itRaw = row[1];
    if (olRaw == null && itRaw == null) continue;
    var ol = String(olRaw == null ? '' : olRaw).trim().toUpperCase();
    if (!ol) continue;
    if (OLS.indexOf(ol) < 0) {
      var msg = (ol === 'EB' || /EXPRESS/.test(ol)) ? 'El operador "' + olRaw + '" ya no forma parte del tablero: se omitió.' : 'Operador no reconocido "' + olRaw + '" (hoja ' + sheetName + '): se omitió. Válidos: ' + OLS.join(', ') + '.';
      if (warnings.indexOf(msg) < 0) warnings.push(msg);
      continue;
    }
    var it = itemOf(itRaw);
    if (!it) { var m2 = 'Ítem no reconocido "' + itRaw + '" (hoja ' + sheetName + ', fila ' + (r + 1) + '): se omitió.'; warnings.push(m2); continue; }
    if (!accs[ol]) accs[ol] = newAcc();
    addTo(accs[ol], it.item, it.kind, rowVals(row, mc));
  }
  Object.keys(accs).forEach(function (ol) { out[ol] = finalize(accs[ol]); });
  return out;
}
/* 2) Tabla dinámica: filas = conceptos, columnas = (.PTTO | .Real | .Desvio) por mes (formato TOL) */
function parsePivot(rows) {
  for (var r = 0; r < rows.length - 1; r++) {
    var mcR = monthCols(rows[r] || []);
    if (mcR.count < 6) continue;
    var sub = rows[r + 1] || [], pt = Array(12).fill(null), rl = Array(12).fill(null), found = false;
    for (var m = 0; m < 12; m++) {
      var c0 = mcR.cols[m]; if (c0 == null) continue;
      for (var d = 0; d < 3; d++) {
        var lab = norm(sub[c0 + d]);
        if (/ptto|presup/.test(lab)) { pt[m] = c0 + d; found = true; }
        else if (/real/.test(lab)) rl[m] = c0 + d;
      }
    }
    if (!found) continue;
    var accP = newAcc(), accR = newAcc();
    for (var q = r + 2; q < rows.length; q++) {
      var row = rows[q] || [], it = itemOf(row[0]); if (!it) continue;
      addTo(accP, it.item, it.kind, rowVals(row, pt));
      addTo(accR, it.item, it.kind, rowVals(row, rl));
    }
    var res = { presu: finalize(accP), real: finalize(accR) };
    if (nonEmpty(res)) return res;
  }
  return null;
}
/* 3) Bloques "PRESUPUESTO 2026" / "GASTOS (REAL) 2026" (formato RDV y TL) */
function parseBlocks(rows, year) {
  var accP = newAcc(), accR = newAcc(), cur = null, cols = null;
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || [], hit = null;
    for (var c = 0; c < Math.min(3, row.length); c++) {
      var m = /^\s*(presupuesto|gastos?(?:\s+real)?)\s*(20\d{2})/i.exec(typeof row[c] === 'string' ? row[c] : '');
      if (m) { hit = { type: /presupuesto/i.test(m[1]) ? 'presu' : 'real', year: +m[2] }; break; }
    }
    if (hit) { cur = hit.year === year ? hit.type : null; cols = null; continue; }
    var mcR = monthCols(row);
    if (mcR.count >= 6) { cols = mcR.cols; continue; }
    if (!cur || !cols) continue;
    var label = null;
    for (c = 0; c < Math.min(3, row.length); c++) { if (typeof row[c] === 'string' && monthIndex(row[c]) < 0 && row[c].trim()) { label = row[c]; break; } }
    if (label == null) continue;
    var it = itemOf(label); if (!it) continue;
    addTo(cur === 'presu' ? accP : accR, it.item, it.kind, rowVals(row, cols));
  }
  var res = { presu: finalize(accP), real: finalize(accR) };
  return nonEmpty(res) ? res : null;
}
/* 4) Solo totales: filas "PRESUPUESTO" y "REAL" con meses en columnas (formato ST) */
function parseTotals(rows) {
  var cols = null, pres = null, real = null;
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || [], mcR = monthCols(row);
    if (mcR.count >= 6) { cols = mcR.cols; continue; }
    if (!cols) continue;
    var lab = norm(row[0]);
    if (lab === 'presupuesto' && !pres) pres = rowVals(row, cols);
    else if (lab === 'real' && !real) real = rowVals(row, cols);
  }
  if (!pres && !real) return null;
  var res = { presu: {}, real: {}, totalsOnly: true };
  [['presu', pres], ['real', real]].forEach(function (p) {
    if (!p[1]) return;
    ITEMS.forEach(function (it) { res[p[0]][it] = p[1].map(function (v) { return v == null ? null : v * ST_SPLIT[it]; }); });
  });
  return nonEmpty(res) ? res : null;
}

function guessOL(filename) {
  var toks = String(filename || '').toUpperCase().split(/[^A-Z0-9]+/), i;
  for (i = 0; i < toks.length; i++) if (OLS.indexOf(toks[i]) >= 0) return toks[i];
  for (i = 0; i < toks.length; i++) if (toks[i] === 'EB' || toks[i] === 'EXPRESS') return 'EB';
  return null;
}

/* Punto de entrada: devuelve { kind, ols:{OL:{presu,real}}, single, needsOL, warnings, notes } */
function parseWorkbook(wbk, opts) {
  opts = opts || {};
  var year = opts.year || 2026, warnings = [], notes = [], names = wbk.names;
  var res = { kind: null, ols: {}, single: null, needsOL: false, warnings: warnings, notes: notes, sheet: null };
  var found = false, i, rows;

  /* plantilla */
  var tplOls = {};
  names.forEach(function (n) {
    var nn = norm(n), kind = /target|objetiv|meta/.test(nn) ? 'target' : /presup|ptto/.test(nn) ? 'presu' : (/gasto|real|ejecu/.test(nn) ? 'real' : null);
    if (!kind) return;
    var r = parseTemplateSheet(wbk.sheets[n], kind, warnings, n);
    if (!r) return;
    found = true;
    Object.keys(r).forEach(function (ol) {
      tplOls[ol] = tplOls[ol] || { presu: {}, real: {}, target: {} };
      tplOls[ol][kind] = r[ol];
    });
  });
  if (found) {
    res.kind = 'plantilla'; res.ols = tplOls;
    if (!Object.keys(tplOls).length) warnings.push('La plantilla no tiene filas con datos para ningún operador.');
    return res;
  }

  /* formatos originales de cada OL */
  var cands = [];
  names.forEach(function (n) {
    rows = wbk.sheets[n];
    var p = parsePivot(rows); if (p && score(p) > 0) cands.push({ k: 'pivot', sheet: n, v: p, s: score(p) + 1e18 });
    var b = parseBlocks(rows, year); if (b && score(b) > 0) cands.push({ k: 'bloques', sheet: n, v: b, s: score(b) + 1e12 });
    var t = parseTotals(rows); if (t && score(t) > 0) cands.push({ k: 'totales', sheet: n, v: t, s: score(t) });
  });
  if (!cands.length) {
    warnings.push('No se reconoció el formato. Usá la plantilla (botón "Plantilla") o un archivo con bloques "PRESUPUESTO ' + year + '" / "GASTOS ' + year + '".');
    return res;
  }
  cands.sort(function (a, b) { return b.s - a.s; });
  var best = cands[0];
  res.kind = best.k; res.sheet = best.sheet; res.single = { presu: best.v.presu, real: best.v.real }; res.needsOL = true;
  if (best.k === 'pivot') notes.push('Formato tabla dinámica (.PTTO / .Real) leído de la hoja "' + best.sheet + '".');
  if (best.k === 'bloques') notes.push('Formato por bloques "PRESUPUESTO ' + year + ' / GASTOS ' + year + '" leído de la hoja "' + best.sheet + '". Los tipos de mantenimiento (correctivo, preventivo, proactivo) se sumaron en Mantenimiento.');
  if (best.k === 'totales') notes.push('Este archivo solo trae totales mensuales (sin apertura por ítem). Se repartieron entre los 5 ítems con la proporción histórica de los otros OLs: ' + ITEMS.map(function (it) { return ITEM_LABEL[it] + ' ' + (ST_SPLIT[it] * 100).toFixed(1) + '%'; }).join(' · ') + '.');
  return res;
}

/* ---------------------------------------------------------------- Modelo */
function emptyData() {
  var d = {};
  OLS.forEach(function (ol) { d[ol] = { presu: {}, real: {}, target: {} }; ITEMS.forEach(function (it) { d[ol].presu[it] = Array(12).fill(0); d[ol].real[it] = Array(12).fill(0); d[ol].target[it] = Array(12).fill(0); }); });
  return d;
}
function emptyModel() { return { app: 'presupuesto-mercado-central', schema: 1, year: 2026, meta: { savedAt: null, savedBy: null, version: 0 }, data: emptyData() }; }
function applyImport(model, olMap) {
  var stats = { cells: 0, byOL: {} };
  Object.keys(olMap).forEach(function (ol) {
    if (!model.data[ol]) return;
    stats.byOL[ol] = { presu: 0, real: 0, target: 0 };
    ['presu', 'real', 'target'].forEach(function (t) {
      if (!model.data[ol][t]) return;
      var src = olMap[ol][t] || {};
      Object.keys(src).forEach(function (it) {
        if (!model.data[ol][t][it]) return;
        src[it].forEach(function (v, m) { if (v != null) { model.data[ol][t][it][m] = Math.round(v * 100) / 100; stats.cells++; stats.byOL[ol][t]++; } });
      });
    });
  });
  return stats;
}
function sumObj(o) { var s = 0; Object.keys(o || {}).forEach(function (k) { o[k].forEach(function (v) { s += v || 0; }); }); return s; }

/* ---------------------------------------------------------------- XLSX (escritura mínima, ZIP sin compresión) */
var CRC_T = null;
function crc32(u8) {
  if (!CRC_T) { CRC_T = new Uint32Array(256); for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; CRC_T[n] = c >>> 0; } }
  var crc = 0xFFFFFFFF; for (var i = 0; i < u8.length; i++) crc = CRC_T[(crc ^ u8[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
function zipStore(files) {
  var enc = new TextEncoder(), parts = [], central = [], offset = 0;
  function w16(a, v) { a.push(v & 255, (v >> 8) & 255); }
  function w32(a, v) { a.push(v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >>> 24) & 255); }
  files.forEach(function (f) {
    var name = enc.encode(f.name), data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data);
    var lh = []; w32(lh, 0x04034b50); w16(lh, 20); w16(lh, 0x0800); w16(lh, 0); w16(lh, 0); w16(lh, 0x21); w32(lh, crc); w32(lh, data.length); w32(lh, data.length); w16(lh, name.length); w16(lh, 0);
    parts.push(new Uint8Array(lh), name, data);
    var ch = []; w32(ch, 0x02014b50); w16(ch, 20); w16(ch, 20); w16(ch, 0x0800); w16(ch, 0); w16(ch, 0); w16(ch, 0x21); w32(ch, crc); w32(ch, data.length); w32(ch, data.length); w16(ch, name.length); w16(ch, 0); w16(ch, 0); w16(ch, 0); w16(ch, 0); w32(ch, 0); w32(ch, offset);
    central.push(new Uint8Array(ch), name);
    offset += lh.length + name.length + data.length;
  });
  var cdSize = 0; central.forEach(function (c) { cdSize += c.length; });
  var end = []; w32(end, 0x06054b50); w16(end, 0); w16(end, 0); w16(end, files.length); w16(end, files.length); w32(end, cdSize); w32(end, offset); w16(end, 0);
  var all = parts.concat(central, [new Uint8Array(end)]), total = 0; all.forEach(function (p) { total += p.length; });
  var out = new Uint8Array(total), pos = 0; all.forEach(function (p) { out.set(p, pos); pos += p.length; });
  return out;
}
function colName(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
/* sheets: [{name, rows:[[cell,...]], widths:[..]}]  cell = valor | {v, s}  s: 1=encabezado 2=número */
function writeXlsx(sheets) {
  var files = [];
  files.push({ name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map(function (s, i) { return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'; }).join('') + '</Types>' });
  files.push({ name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' });
  files.push({ name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets.map(function (s, i) { return '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join('') + '</sheets></workbook>' });
  files.push({ name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map(function (s, i) { return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; }).join('') + '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' });
  files.push({ name: 'xl/styles.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF6B800"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' });
  sheets.forEach(function (s, i) {
    var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>';
    if (s.widths) xml += '<cols>' + s.widths.map(function (w, k) { return '<col min="' + (k + 1) + '" max="' + (k + 1) + '" width="' + w + '" customWidth="1"/>'; }).join('') + '</cols>';
    xml += '<sheetData>';
    s.rows.forEach(function (row, r) {
      xml += '<row r="' + (r + 1) + '">';
      row.forEach(function (cell, c) {
        if (cell == null || cell === '') return;
        var v = cell, st = 0; if (typeof cell === 'object') { v = cell.v; st = cell.s || 0; }
        var ref = colName(c) + (r + 1);
        if (typeof v === 'number') xml += '<c r="' + ref + '"' + (st ? ' s="' + st + '"' : ' s="2"') + '><v>' + v + '</v></c>';
        else xml += '<c r="' + ref + '" t="inlineStr"' + (st ? ' s="' + st + '"' : '') + '><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
      });
      xml += '</row>';
    });
    xml += '</sheetData></worksheet>';
    files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: xml });
  });
  return zipStore(files);
}
function modelToSheets(model) {
  var hdr = ['OL', 'Ítem'].concat(MONTHS_SHORT.map(function (m) { return m.charAt(0).toUpperCase() + m.slice(1); }), ['Total']);
  function build(t) {
    var rows = [hdr.map(function (h) { return { v: h, s: 1 }; })];
    OLS.forEach(function (ol) { ITEMS.forEach(function (it) {
      var arr = model.data[ol][t][it], tot = arr.reduce(function (a, b) { return a + (b || 0); }, 0);
      rows.push([ol, ITEM_LABEL[it]].concat(arr, [tot]));
    }); });
    return { name: t === 'presu' ? 'Presupuesto' : t === 'target' ? 'Targets' : 'Gastos', rows: rows, widths: [8, 18].concat(Array(12).fill(14), [16]) };
  }
  return [build('presu'), build('real'), build('target')];
}

var api = { OLS: OLS, ITEMS: ITEMS, ITEM_LABEL: ITEM_LABEL, MONTHS_FULL: MONTHS_FULL, MONTHS_SHORT: MONTHS_SHORT, ST_SPLIT: ST_SPLIT,
  readXlsx: readXlsx, parseWorkbook: parseWorkbook, guessOL: guessOL, emptyModel: emptyModel, applyImport: applyImport, sumObj: sumObj,
  writeXlsx: writeXlsx, modelToSheets: modelToSheets, parseNum: parseNum, norm: norm };
if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.XL = api;
})(typeof window !== 'undefined' ? window : globalThis);

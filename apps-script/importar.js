/**
 * Plan financiero · Importador de movimientos
 * Bancos reconocidos: Openbank (cuenta y tarjeta), CaixaBank, ING.
 *
 * Cómo se usa: deja los archivos que descargas del banco en la carpeta «Movimientos»
 * (junto a este sheet) y pulsa  Plan financiero → Importar movimientos.
 * Requisito: servicio avanzado «Drive API» activado (Servicios → Drive API).
 */

var HOJA_MOV = 'MOVIMIENTOS', HOJA_REGLAS = 'REGLAS', HOJA_DATOS = 'MIS DATOS';
var MOV_FILA1 = 4;            // primera fila de datos en MOVIMIENTOS
var REGLAS_FILA1 = 4;         // primera fila de datos en REGLAS
var DATOS_FILA1 = 6;          // primera fila de datos en MIS DATOS
var CARPETA = 'Movimientos', PROCESADOS = 'Procesados';
var POR_REVISAR = '❓ Por revisar';
// Las tres cuentas con saldo, en el orden de las columnas B, C y D de MIS DATOS.
// Si cambias un nombre aquí, cámbialo también en BANCOS y en generador/datos-demo/config.json.
var CUENTAS_SALDO = ['CaixaBank', 'Openbank', 'ING'];

// =====================================================================
// Funciones puras (sin Google): leen tablas, normalizan, aplican reglas
// =====================================================================

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function decodeEntities_(s) {
  var map = { nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", euro: '€', ordf: 'ª', ordm: 'º',
    aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó', uacute: 'ú', ntilde: 'ñ', uuml: 'ü', ccedil: 'ç',
    Aacute: 'Á', Eacute: 'É', Iacute: 'Í', Oacute: 'Ó', Uacute: 'Ú', Ntilde: 'Ñ', Uuml: 'Ü', Ccedil: 'Ç' };
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, function (m, e) {
    if (e.charAt(0) === '#') {
      var hex = e.charAt(1).toLowerCase() === 'x';
      return String.fromCharCode(parseInt(e.slice(hex ? 2 : 1), hex ? 16 : 10));
    }
    return Object.prototype.hasOwnProperty.call(map, e) ? map[e] : m;
  });
}

/** Tabla HTML (los .xls de Openbank son HTML por dentro) → filas de texto. */
function htmlToRows_(html) {
  var rows = [];
  html.split(/<tr[\s>]/i).slice(1).forEach(function (tr) {
    tr = tr.split(/<\/tr>/i)[0];
    var cells = [], re = /<t[dh]\b[^>]*?(\/>|>([\s\S]*?)<\/t[dh]>)/gi, m;
    while ((m = re.exec(tr)) !== null) {
      var txt = m[1] === '/>' ? '' : (m[2] || '');
      cells.push(decodeEntities_(txt.replace(/<[^>]*>/g, ' ')).replace(/\s+/g, ' ').trim());
    }
    if (cells.length) rows.push(cells);
  });
  return rows;
}

/** Importe: número, "1.600,00", "-122,84", "-5.99", "1,234.56"… */
function num_(v) {
  if (typeof v === 'number') return v;
  if (v == null) return NaN;
  var s = String(v).replace(/EUR|€|\s/gi, '').replace(/−/g, '-');
  if (s === '') return NaN;
  var c = s.lastIndexOf(','), d = s.lastIndexOf('.');
  if (c > -1 && d > -1) s = c > d ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (c > -1) s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  else if (d > -1 && /^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  var n = Number(s);
  return isNaN(n) ? NaN : n;
}

/** Fecha: Date, número de serie o texto dd/mm/aaaa, dd-mm-aaaa, aaaa-mm-dd. */
function fecha_(v) {
  if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    var u = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5);
    return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
  }
  var s = String(v == null ? '' : v).trim(), m;
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/))) return new Date(+m[3], +m[2] - 1, +m[1]);
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function iso_(d) {
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}

/**
 * Cómo es la tabla de cada banco (cabeceras sin acentos y en mayúsculas).
 * Para añadir tu banco: una entrada más con las cabeceras que lo distinguen (reconoce) y qué columna
 * es la fecha, el concepto, el importe y el saldo. Basta con el principio de cada cabecera.
 */
var BANCOS = [
  { cuenta: 'ING', reconoce: ['F. VALOR', 'SUBCATEGORIA', 'DESCRIPCION', 'IMPORTE'],
    fecha: 'F. VALOR', concepto: ['DESCRIPCION', 'COMENTARIO'], importe: 'IMPORTE', saldo: 'SALDO' },
  { cuenta: 'CaixaBank', reconoce: ['FECHA', 'MOVIMIENTO', 'MAS DATOS', 'IMPORTE'],
    fecha: 'FECHA', concepto: ['MOVIMIENTO', 'MAS DATOS'], importe: 'IMPORTE', saldo: 'SALDO' },
  { cuenta: 'Openbank tarjeta', reconoce: ['FECHA OPERACION', 'HORA', 'CONCEPTO', 'SITUACION', 'IMPORTE'],
    fecha: 'FECHA OPERACION', hora: 'HORA', concepto: ['CONCEPTO'], importe: 'IMPORTE', saldo: null },
  { cuenta: 'Openbank', reconoce: ['FECHA OPERACION', 'FECHA VALOR', 'CONCEPTO', 'IMPORTE', 'SALDO'],
    fecha: 'FECHA OPERACION', concepto: ['CONCEPTO'], importe: 'IMPORTE', saldo: 'SALDO' }
];

/** Índice de la columna: primero coincidencia exacta, después «empieza por». */
function col_(cab, clave) {
  for (var i = 0; i < cab.length; i++) if (cab[i] === clave) return i;
  for (var j = 0; j < cab.length; j++) if (cab[j] && cab[j].indexOf(clave) === 0) return j;
  return -1;
}

/** Filas de un archivo → {cuenta, movs[]} o null si no se reconoce el formato. */
function tableToMovs_(filas) {
  for (var r = 0; r < Math.min(filas.length, 40); r++) {
    var cab = filas[r].map(norm_);
    for (var b = 0; b < BANCOS.length; b++) {
      var B = BANCOS[b];
      if (!B.reconoce.every(function (k) { return col_(cab, k) > -1; })) continue;
      var iF = col_(cab, B.fecha), iI = col_(cab, B.importe);
      var iS = B.saldo ? col_(cab, B.saldo) : -1, iH = B.hora ? col_(cab, B.hora) : -1;
      var iC = B.concepto.map(function (k) { return col_(cab, k); }).filter(function (i) { return i > -1; });
      var movs = [], vistos = {};
      for (var k = r + 1; k < filas.length; k++) {
        var f = filas[k], fe = fecha_(f[iF]), imp = num_(f[iI]);
        if (!fe || isNaN(imp)) continue;                        // filas vacías, «Total»…
        var concepto = iC.map(function (i) { return String(f[i] == null ? '' : f[i]).trim(); })
          .filter(function (t) { return t !== ''; }).join(' · ');
        var saldo = iS > -1 ? num_(f[iS]) : NaN;
        var hv = iH > -1 ? f[iH] : '';
        var hora = hv instanceof Date ? ('0' + hv.getHours()).slice(-2) + ':' + ('0' + hv.getMinutes()).slice(-2)
          : String(hv == null ? '' : hv).trim();
        var base = [B.cuenta, iso_(fe), hora, imp.toFixed(2), isNaN(saldo) ? '' : saldo.toFixed(2)].join('|');
        vistos[base] = (vistos[base] || 0) + 1;
        movs.push({ fecha: fe, cuenta: B.cuenta, hora: hora, concepto: concepto, importe: Math.round(imp * 100) / 100,
          saldo: isNaN(saldo) ? '' : Math.round(saldo * 100) / 100, clave: base + '#' + vistos[base] });
      }
      return { cuenta: B.cuenta, movs: movs };
    }
  }
  return null;
}

/** Primera regla que encaja: el concepto contiene el texto (sin mirar mayúsculas ni acentos). */
function aplicaReglas_(mov, reglas) {
  var c = norm_(mov.concepto), cu = norm_(mov.cuenta);
  for (var i = 0; i < reglas.length; i++) {
    var r = reglas[i];
    if (!r.patron || !r.categoria) continue;
    if (r.cuenta && norm_(r.cuenta) !== cu) continue;
    if (r.signo === '+' && !(mov.importe > 0)) continue;
    if (r.signo === '-' && !(mov.importe < 0)) continue;
    if (c.indexOf(norm_(r.patron)) > -1) return r.categoria;
  }
  return POR_REVISAR;
}

/**
 * Saldo a fin de mes de cada cuenta. Dentro de un día, el saldo final es el del movimiento
 * cuyo saldo no es el «saldo anterior» de otro movimiento del mismo día (no depende del orden).
 */
function saldosMensuales_(movs, cuentas) {
  var finDia = {}, primero = {}, ultimoMes = null;
  movs.forEach(function (m) {
    var mes = new Date(m.fecha.getFullYear(), m.fecha.getMonth(), 1);
    if (!ultimoMes || mes > ultimoMes) ultimoMes = mes;
  });
  cuentas.forEach(function (cu) {
    var porDia = {};
    movs.forEach(function (m) {
      if (m.cuenta !== cu || m.saldo === '' || m.saldo == null) return;
      (porDia[iso_(m.fecha)] = porDia[iso_(m.fecha)] || []).push(m);
    });
    var dias = Object.keys(porDia).sort();
    if (!dias.length) return;
    finDia[cu] = {};
    dias.forEach(function (d) {
      var ms = porDia[d];
      var previos = ms.map(function (m) { return (m.saldo - m.importe).toFixed(2); });
      var fin = ms.filter(function (m, i) {
        return !previos.some(function (p, j) { return j !== i && p === m.saldo.toFixed(2); });
      })[0] || ms[0];
      finDia[cu][d] = fin.saldo;
    });
    var p = fecha_(dias[0]);
    primero[cu] = new Date(p.getFullYear(), p.getMonth(), 1);
  });
  var conDatos = Object.keys(primero);
  if (!conDatos.length) return { meses: [], saldos: {} };
  var inicio = conDatos.map(function (c) { return primero[c]; }).reduce(function (a, b) { return a > b ? a : b; });
  var meses = [], saldos = {};
  cuentas.forEach(function (c) { saldos[c] = []; });
  for (var m = new Date(inicio); m <= ultimoMes; m = new Date(m.getFullYear(), m.getMonth() + 1, 1)) {
    var finMes = iso_(new Date(m.getFullYear(), m.getMonth() + 1, 0));
    meses.push(m);
    cuentas.forEach(function (c) {
      if (!finDia[c]) { saldos[c].push(''); return; }
      var dias = Object.keys(finDia[c]).filter(function (d) { return d <= finMes; }).sort();
      saldos[c].push(dias.length ? finDia[c][dias[dias.length - 1]] : '');
    });
  }
  return { meses: meses, saldos: saldos };
}

// =====================================================================
// Parte de Google Sheets / Drive
// =====================================================================

// Fechas: el sheet y el proyecto de Apps Script pueden tener zonas horarias distintas.
// Todo lo que se lee se pasa a 'aaaa-mm-dd' en la zona del sheet de origen, y lo que se
// escribe se convierte a medianoche en la zona de este sheet.
var TZ_ = null;
function tz_() { return TZ_ || (TZ_ = SpreadsheetApp.getActive().getSpreadsheetTimeZone()); }
function deSheet_(d, tz) { return fecha_(Utilities.formatDate(d, tz || tz_(), 'yyyy-MM-dd')); }
function aSheet_(d) { return Utilities.parseDate(iso_(d), tz_(), 'yyyy-MM-dd'); }

function valoresDe_(ss) {
  var tz = ss.getSpreadsheetTimeZone();
  return ss.getSheets()[0].getDataRange().getValues().map(function (fila) {
    return fila.map(function (c) {
      if (!(c instanceof Date)) return c;
      return c.getFullYear() < 1901 ? Utilities.formatDate(c, tz, 'HH:mm') : Utilities.formatDate(c, tz, 'yyyy-MM-dd');
    });
  });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Plan financiero')
    .addItem('Importar movimientos', 'importarMovimientos')
    .addItem('Reaplicar reglas a lo pendiente', 'reaplicarReglas')
    .addItem('Recalcular saldos de MIS DATOS', 'recalcularSaldos')
    .addToUi();
}

function subcarpeta_(padre, nombre) {
  var it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

function carpetas_() {
  var ss = SpreadsheetApp.getActive();
  var padre = DriveApp.getFileById(ss.getId()).getParents().next();
  var mov = subcarpeta_(padre, CARPETA);
  return { mov: mov, proc: subcarpeta_(mov, PROCESADOS) };
}

function leerFilas_(file) {
  var mime = file.getMimeType();
  if (mime === MimeType.GOOGLE_SHEETS) {
    return valoresDe_(SpreadsheetApp.openById(file.getId()));
  }
  var blob = file.getBlob(), b = blob.getBytes();
  var ole = b.length > 2 && (b[0] & 0xFF) === 0xD0 && (b[1] & 0xFF) === 0xCF;   // .xls binario
  var zip = b.length > 2 && b[0] === 0x50 && b[1] === 0x4B;                     // .xlsx
  if (ole || zip) {
    var tmp = Drive.Files.copy({ name: '_tmp_importacion', mimeType: MimeType.GOOGLE_SHEETS }, file.getId());
    try {
      return valoresDe_(SpreadsheetApp.openById(tmp.id));
    } finally {
      DriveApp.getFileById(tmp.id).setTrashed(true);
    }
  }
  var txt = blob.getDataAsString('UTF-8');
  if (txt.indexOf('�') > -1) txt = blob.getDataAsString('ISO-8859-1');
  if (/<tr[\s>]/i.test(txt)) return htmlToRows_(txt);
  var l1 = txt.split(/\r?\n/)[0];
  return Utilities.parseCsv(txt, l1.split(';').length > l1.split(',').length ? ';' : ',');
}

function leerReglas_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(HOJA_REGLAS);
  var n = sh.getLastRow() - REGLAS_FILA1 + 1;
  if (n < 1) return [];
  return sh.getRange(REGLAS_FILA1, 1, n, 4).getValues().map(function (r) {
    return { patron: String(r[0]).trim(), cuenta: String(r[1]).trim(), signo: String(r[2]).trim(), categoria: String(r[3]).trim() };
  });
}

function filasMov_(sh) {
  var n = sh.getLastRow() - MOV_FILA1 + 1;
  return n < 1 ? [] : sh.getRange(MOV_FILA1, 1, n, 11).getValues();
}

function escribirMovs_(sh, movs) {
  var fila = Math.max(sh.getLastRow() + 1, MOV_FILA1);
  var valores = movs.map(function (m) {
    return [aSheet_(m.fecha), '', m.cuenta, m.concepto, m.importe, m.saldo, m.categoria, '', '', '', m.clave];
  });
  valores.forEach(function (v, i) {
    var r = fila + i;
    v[1] = '=DATE(YEAR(A' + r + '),MONTH(A' + r + '),1)';
    v[7] = '=IFERROR(VLOOKUP(G' + r + ',\'CATEGORÍAS\'!$A$4:$C$100,2,FALSE),"Por revisar")';
    v[8] = '=IFERROR(VLOOKUP(G' + r + ',\'CATEGORÍAS\'!$A$4:$C$100,3,FALSE),"")';
  });
  enIngles_(function () { sh.getRange(fila, 1, valores.length, 11).setValues(valores); });
  var total = sh.getLastRow() - MOV_FILA1 + 1;
  sh.getRange(MOV_FILA1, 1, total, 11).sort({ column: 1, ascending: false });
}

/** Escribe fórmulas con la configuración en inglés un momento (una hoja en español separa con «;»). */
function enIngles_(fn) {
  var ss = SpreadsheetApp.getActive(), loc = ss.getSpreadsheetLocale();
  try {
    if (loc !== 'en_US') { ss.setSpreadsheetLocale('en_US'); SpreadsheetApp.flush(); }
    fn();
    SpreadsheetApp.flush();
  } finally {
    if (loc !== 'en_US') { ss.setSpreadsheetLocale(loc); SpreadsheetApp.flush(); }
  }
}

function importarMovimientos() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(HOJA_MOV), ui = SpreadsheetApp.getUi();
  var c = carpetas_(), archivos = [], it = c.mov.getFiles();
  while (it.hasNext()) archivos.push(it.next());
  archivos = archivos.filter(function (f) { return f.getName().indexOf('_tmp') !== 0; });
  if (!archivos.length) {
    ui.alert('No hay archivos nuevos en la carpeta «' + CARPETA + '».');
    return;
  }
  var existentes = {};
  filasMov_(sh).forEach(function (r) { if (r[10]) existentes[r[10]] = true; });
  var reglas = leerReglas_(), nuevas = [], resumen = [];
  archivos.forEach(function (f) {
    var filas;
    try { filas = leerFilas_(f); } catch (e) {
      resumen.push('⚠️ ' + f.getName() + ': no se pudo leer (' + e.message + ')');
      return;
    }
    var res = tableToMovs_(filas);
    if (!res) { resumen.push('⚠️ ' + f.getName() + ': formato no reconocido, se queda en la carpeta'); return; }
    var n = 0, d = 0;
    res.movs.forEach(function (m) {
      if (existentes[m.clave]) { d++; return; }
      existentes[m.clave] = true;
      m.categoria = aplicaReglas_(m, reglas);
      nuevas.push(m); n++;
    });
    resumen.push('✅ ' + f.getName() + ' (' + res.cuenta + '): ' + n + ' nuevos, ' + d + ' ya estaban');
    f.moveTo(c.proc);
  });
  if (nuevas.length) escribirMovs_(sh, nuevas);
  SpreadsheetApp.flush();
  recalcularSaldos();
  var desde = ss.getSheetByName(HOJA_DATOS).getRange(DATOS_FILA1, 1).getValue();
  var pend = filasMov_(sh).filter(function (r) { return r[7] === 'Por revisar'; });
  var recientes = desde instanceof Date ? pend.filter(function (r) { return r[0] instanceof Date && r[0] >= desde; }).length : pend.length;
  ui.alert('Importación terminada\n\n' + resumen.join('\n') + '\n\nPor revisar: ' + recientes +
    ' desde el primer mes de MIS DATOS (' + pend.length + ' en todo el histórico; los antiguos no afectan a las cuentas).');
}

function reaplicarReglas() {
  var sh = SpreadsheetApp.getActive().getSheetByName(HOJA_MOV), reglas = leerReglas_();
  var filas = filasMov_(sh), cambios = 0;
  var cats = filas.map(function (r) {
    if (r[6] !== POR_REVISAR && r[6] !== '') return [r[6]];
    var nueva = aplicaReglas_({ concepto: r[3], cuenta: r[2], importe: Number(r[4]) }, reglas);
    if (nueva !== POR_REVISAR) cambios++;
    return [nueva];
  });
  if (cats.length) sh.getRange(MOV_FILA1, 7, cats.length, 1).setValues(cats);
  SpreadsheetApp.getUi().alert(cambios + ' movimientos pendientes han encontrado categoría.');
}

/** Reescribe meses y saldos (columnas A–D) de MIS DATOS; respeta lo que tengas a mano en E y F. */
function recalcularSaldos() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(HOJA_DATOS);
  var movs = filasMov_(ss.getSheetByName(HOJA_MOV)).filter(function (r) { return r[0] instanceof Date; })
    .map(function (r) { return { fecha: deSheet_(r[0]), cuenta: r[2], importe: Number(r[4]), saldo: r[5] === '' ? '' : Number(r[5]) }; });
  var res = saldosMensuales_(movs, CUENTAS_SALDO);
  var n = Math.max(sh.getLastRow() - DATOS_FILA1 + 1, 1);
  var manual = {};
  sh.getRange(DATOS_FILA1, 1, n, 6).getValues().forEach(function (r) {
    if (r[0] instanceof Date) manual[iso_(deSheet_(r[0]))] = [r[4], r[5]];
  });
  sh.getRange(DATOS_FILA1, 1, n, 6).clearContent();
  if (!res.meses.length) return;
  var out = res.meses.map(function (m, i) {
    var man = manual[iso_(m)] || [0, 0];
    return [aSheet_(m)].concat(CUENTAS_SALDO.map(function (c) { return res.saldos[c][i]; }), [man[0] === '' ? 0 : man[0], man[1] === '' ? 0 : man[1]]);
  });
  sh.getRange(DATOS_FILA1, 1, out.length, 6).setValues(out);
}

// Para las pruebas fuera de Google
if (typeof module !== 'undefined') {
  module.exports = { htmlToRows_: htmlToRows_, num_: num_, fecha_: fecha_, iso_: iso_, tableToMovs_: tableToMovs_,
    aplicaReglas_: aplicaReglas_, saldosMensuales_: saldosMensuales_, norm_: norm_, POR_REVISAR: POR_REVISAR,
    CUENTAS_SALDO: CUENTAS_SALDO };
}

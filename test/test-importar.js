// Pruebas del importador fuera de Google:  node test/test-importar.js
// Usan los extractos de extractos-ejemplo/ y los datos de generador/datos-demo/, todo ficticio.
process.env.TZ = 'Europe/Madrid';
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const RAIZ = path.join(__dirname, '..');
const M = require(path.join(RAIZ, 'apps-script', 'importar.js'));
const ext = (f) => fs.readFileSync(path.join(RAIZ, 'extractos-ejemplo', f), 'latin1');
const demo = JSON.parse(fs.readFileSync(path.join(RAIZ, 'generador', 'datos-demo', 'movimientos.json'), 'utf8'));
const saldosDemo = JSON.parse(fs.readFileSync(path.join(RAIZ, 'generador', 'datos-demo', 'saldos.json'), 'utf8'));
const claves = (movs) => movs.map((m) => m.clave).sort();
const deCuenta = (c) => demo.filter((m) => m.cuenta === c);
let n = 0;
const ok = (txt) => { n++; console.log('  ✓ ' + txt); };

// 1) Openbank cuenta (.xls que por dentro es HTML): reconoce el banco y saca los mismos movimientos
const html = M.htmlToRows_(ext('openbank-cuenta.xls'));
const a = M.tableToMovs_(html);
assert.strictEqual(a.cuenta, 'Openbank');
assert.deepStrictEqual(claves(a.movs), claves(deCuenta('Openbank')));
ok('Openbank cuenta: ' + a.movs.length + ' movimientos, las mismas claves que los datos de ejemplo');

// 2) El mismo archivo dos veces da las mismas claves (así no duplica al reimportar)
assert.deepStrictEqual(M.tableToMovs_(html).movs.map((m) => m.clave), a.movs.map((m) => m.clave));
ok('reimportar el mismo archivo da las mismas claves');

// 3) Como si Drive lo hubiera convertido a Sheets: fechas Date y números
const conv = html.map((r) => r.map((c) => (/^\d{2}\/\d{2}\/\d{4}$/.test(c) ? M.fecha_(c) : /^-?[\d.]+,\d{2}$/.test(c) ? M.num_(c) : c)));
assert.deepStrictEqual(M.tableToMovs_(conv).movs.map((m) => m.clave), a.movs.map((m) => m.clave));
ok('texto o Date/número dan las mismas claves');

// 4) Tarjeta: sin saldo y con hora; la hora puede llegar como texto o como Date
const t = M.htmlToRows_(ext('openbank-tarjeta.xls'));
const tm = M.tableToMovs_(t);
assert.strictEqual(tm.cuenta, 'Openbank tarjeta');
assert.deepStrictEqual(claves(tm.movs), claves(deCuenta('Openbank tarjeta')));
const tconv = t.map((r) => r.map((c) => (/^\d{2}:\d{2}$/.test(c) ? new Date(1899, 11, 30, +c.slice(0, 2), +c.slice(3)) : c)));
assert.deepStrictEqual(M.tableToMovs_(tconv).movs.map((m) => m.clave), tm.movs.map((m) => m.clave));
ok('Openbank tarjeta: ' + tm.movs.length + ' movimientos, hora como texto o como Date');

// 5) CaixaBank e ING llegan como .xls binario: Apps Script los convierte con la Drive API y lee
//    los valores. Aquí se simula esa lectura: título arriba, una fila vacía, cabeceras y filas.
const comoDrive = (titulo, cab, filas) => [[titulo], [''], cab, ...filas];
const fecha = (iso) => M.fecha_(iso);
const caixa = comoDrive('Movimientos de la cuenta', ['Fecha', 'Fecha valor', 'Movimiento', 'Más datos', 'Importe', 'Saldo'],
  deCuenta('CaixaBank').map((m) => [fecha(m.fecha), fecha(m.fecha), m.concepto, '', m.importe, m.saldo]));
const cm = M.tableToMovs_(caixa);
assert.strictEqual(cm.cuenta, 'CaixaBank');
assert.deepStrictEqual(claves(cm.movs), claves(deCuenta('CaixaBank')));
ok('CaixaBank: ' + cm.movs.length + ' movimientos');
const ing = comoDrive('Cuenta Ahorro', ['F. VALOR', 'CATEGORÍA', 'SUBCATEGORÍA', 'DESCRIPCIÓN', 'COMENTARIO', 'IMAGEN', 'IMPORTE (€)', 'SALDO (€)'],
  deCuenta('ING').map((m) => [fecha(m.fecha), 'Otros', 'Otros', m.concepto, '', '', m.importe, m.saldo]));
const im = M.tableToMovs_(ing);
assert.strictEqual(im.cuenta, 'ING');
assert.deepStrictEqual(claves(im.movs), claves(deCuenta('ING')));
ok('ING: ' + im.movs.length + ' movimientos (cabeceras con acentos y «(€)»)');

// 6) Saldos a fin de mes: coinciden con los de ejemplo y no dependen del orden
const todos = demo.map((m) => ({ ...m, fecha: M.fecha_(m.fecha) }));
const s1 = M.saldosMensuales_(todos, M.CUENTAS_SALDO);
assert.deepStrictEqual(s1.meses.map(M.iso_), saldosDemo.meses);
assert.deepStrictEqual(s1.saldos, saldosDemo.saldos);
const barajados = todos.slice().sort(() => Math.random() - 0.5);
assert.deepStrictEqual(M.saldosMensuales_(barajados, M.CUENTAS_SALDO).saldos, s1.saldos);
ok('saldos de fin de mes iguales a los de ejemplo, en cualquier orden');

// 7) Números en los formatos que mandan los bancos
const casos = { '-122,84': -122.84, '1.600,00': 1600, '69.522,52': 69522.52, '-5.99': -5.99, '1,234.56': 1234.56,
  '-1.047,03': -1047.03, '800,00 EUR': 800, '1.600': 1600 };
for (const [k, v] of Object.entries(casos)) assert.strictEqual(M.num_(k), v, k);
ok('importes con coma, punto, «EUR» y separador de miles');

// 8) Un archivo que no es de ningún banco conocido
assert.strictEqual(M.tableToMovs_([['hola', 'adios'], ['1', '2']]), null);
ok('formato desconocido → null (el archivo se queda en la carpeta)');

// 9) Reglas: signo y cuenta
const regla = [{ patron: 'NOMINA', signo: '+', categoria: 'Nómina' }];
assert.strictEqual(M.aplicaReglas_({ concepto: 'Nómina trf', cuenta: 'CaixaBank', importe: 10 }, regla), 'Nómina');
assert.strictEqual(M.aplicaReglas_({ concepto: 'Nómina trf', cuenta: 'CaixaBank', importe: -10 }, regla), M.POR_REVISAR);
assert.strictEqual(M.aplicaReglas_({ concepto: 'x', cuenta: 'ING', importe: 1 }, [{ patron: 'X', cuenta: 'Openbank', categoria: 'A' }]), M.POR_REVISAR);
ok('las reglas respetan el signo y la cuenta');

console.log('\nTodas las pruebas OK (' + n + ')');

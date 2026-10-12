dash.loader("kpi", {
  description: "Lee la tabla KPI de la pestaña PANEL_DATOS del .xlsx que exporta Google Drive. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("KPI", []);
t.rows.forEach((r) => { if (r.clave === 'fecha_datos') r.valor = fecha(r.valor); });
return t.rows;
  },
});

dash.loader("meses", {
  description: "Lee la tabla MESES (saldos, ingresos, gastos y gasto por grupo de cada mes) de la pestaña PANEL_DATOS. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("MESES", ["mes"]);

return t.rows;
  },
});

dash.loader("casa", {
  description: "Lee la tabla CASA (recibos de la vivienda por concepto) de la pestaña PANEL_DATOS. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("CASA", []);

return t.rows;
  },
});

dash.loader("casa_meses", {
  description: "Lee la tabla CASA_MESES (recibos de la vivienda por mes) de la pestaña PANEL_DATOS. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("CASA_MESES", ["mes"]);

return t.rows;
  },
});

dash.loader("proyeccion", {
  description: "Lee la tabla PROYECCION (patrimonio año a año hasta la jubilación) de la pestaña PANEL_DATOS. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("PROYECCION", []);
t.rows = t.rows.filter((r) => r.edad !== '');
return t.rows;
  },
});

dash.loader("jubilacion", {
  description: "Lee la tabla JUBILACION (capital año a año en la jubilación) de la pestaña PANEL_DATOS. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("JUBILACION", []);

return t.rows;
  },
});

dash.loader("hitos", {
  description: "Lee la tabla HITOS de la pestaña PANEL_DATOS. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("HITOS", []);

return t.rows;
  },
});

dash.loader("viajes", {
  description: "Lee la tabla VIAJES (tu parte de cada viaje y el saldo de la cuenta de viajes) de la pestaña PANEL_DATOS. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("VIAJES", ["fecha"]);

return t.rows;
  },
});

dash.loader("viajes_escenarios", {
  description: "Lee la tabla VIAJES_ESCENARIOS (qué pasa con otras aportaciones) de la pestaña PANEL_DATOS. Solo convierte el archivo en filas.",
  fn: async ({ answer }) => {
const obj = typeof answer === 'string' ? JSON.parse(answer) : answer;
if (!obj || !obj.content) throw new Error('Google Drive no devolvió el archivo' + (obj && obj.error ? ': ' + JSON.stringify(obj.error) : ''));
const bin = atob(obj.content);
const buf = new Uint8Array(bin.length);
for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
const dv = new DataView(buf.buffer);
let eocd = -1;
for (let i = buf.length - 22; i >= Math.max(0, buf.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
if (eocd < 0) throw new Error('El archivo no es un .xlsx');
const nEnt = dv.getUint16(eocd + 10, true);
let p = dv.getUint32(eocd + 16, true);
const files = {};
for (let k = 0; k < nEnt; k++) {
  const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
  const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
  const off = dv.getUint32(p + 42, true);
  const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nlen));
  files[name] = { method, csize, off };
  p += 46 + nlen + xlen + clen;
}
const leer = async (name) => {
  const f = files[name];
  if (!f) return null;
  const start = f.off + 30 + dv.getUint16(f.off + 26, true) + dv.getUint16(f.off + 28, true);
  const raw = buf.subarray(start, start + f.csize);
  if (f.method === 0) return new TextDecoder().decode(raw);
  const ds = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return await new Response(ds).text();
};
const ent = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e.toLowerCase()]);
const textos = (xml) => (xml.match(/<t\b[^>]*>[\s\S]*?<\/t>|<t\b[^>]*\/>/g) || []).map((t) => ent(t.replace(/<[^>]+>/g, ''))).join('');
const wb = await leer('xl/workbook.xml');
const rels = await leer('xl/_rels/workbook.xml.rels');
const hoja = (wb.match(/<sheet\b[^>]*>/g) || []).find((s) => /name="PANEL_DATOS"/.test(s));
if (!hoja) throw new Error('El sheet no tiene la pestaña PANEL_DATOS');
const rid = (hoja.match(/r:id="([^"]+)"/) || hoja.match(/\bid="([^"]+)"/))[1];
const rel = (rels.match(/<Relationship\b[^>]*>/g) || []).find((r) => r.includes('Id="' + rid + '"'));
let target = rel.match(/Target="([^"]+)"/)[1];
target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
const ssXml = await leer('xl/sharedStrings.xml');
const ss = ssXml ? (ssXml.match(/<si\b[\s\S]*?<\/si>/g) || []).map(textos) : [];
const xml = await leer(target);
const col = (ref) => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const grid = [];
for (const fila of xml.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || []) {
  const r = +fila.match(/\br="(\d+)"/)[1] - 1;
  const celdas = [];
  for (const c of fila.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || []) {
    const ref = c.match(/\br="([A-Z]+\d+)"/)[1], t = (c.match(/\bt="(\w+)"/) || [])[1];
    const v = (c.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
    let val = '';
    if (t === 's') val = ss[+v] ?? '';
    else if (t === 'inlineStr') val = textos(c);
    else if (t === 'str' || t === 'e') val = v == null ? '' : ent(v);
    else if (t === 'b') val = v === '1';
    else if (v != null && v !== '') val = Number(v);
    celdas[col(ref)] = val;
  }
  grid[r] = celdas;
}
const fecha = (n) => typeof n === 'number' ? new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10) : n;
const tabla = (nombre, fechas) => {
  const i = grid.findIndex((f) => f && f[0] === '## ' + nombre);
  if (i < 0) throw new Error('Falta la tabla ' + nombre + ' en PANEL_DATOS');
  const columns = (grid[i + 1] || []).filter((x) => x !== '' && x != null);
  const rows = [];
  for (let k = i + 2; k < grid.length; k++) {
    const f = grid[k];
    if (!f || f[0] === '' || f[0] == null) break;
    const o = {};
    columns.forEach((c, j) => { const v = f[j]; o[c] = v == null ? '' : (fechas.includes(c) ? fecha(v) : v); });
    rows.push(o);
  }
  return { columns, rows };
};
const t = tabla("VIAJES_ESCENARIOS", []);

return t.rows;
  },
});

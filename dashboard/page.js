// ---------- cálculos a la vista en Fuentes
dash.calc('dura', {
  title: 'Hasta cuándo te dura el dinero',
  description: 'Edad a la que se acaba el capital en la jubilación, sin decisiones; «más allá» si no se acaba en el horizonte.',
  inputs: ['kpi'],
  fn: (kpi) => {
    const v = (k) => (kpi.find((r) => r.clave === k) || {}).valor;
    const sin = v('dura_sin'), max = v('dura_max');
    return [{ texto: sin >= max ? 'Más allá de los ' + max : 'Hasta los ' + sin }];
  },
});

dash.calc('gasto_grupos', {
  title: 'Gasto medio por grupo',
  description: 'Media mensual de cada grupo de gasto en los meses completos.',
  inputs: ['meses'],
  fn: (meses) => {
    const completos = meses.filter((m) => m.completo === 1);
    const grupos = Object.keys(meses[0] || {}).filter((k) => k.startsWith('g_'));
    return grupos
      .map((g) => ({ grupo: g.slice(2), media: completos.length ? completos.reduce((s, m) => s + (+m[g] || 0), 0) / completos.length : 0 }))
      .filter((r) => r.media > 0.5)
      .sort((a, b) => b.media - a.media);
  },
});

dash.calc('trayectoria', {
  title: 'Patrimonio por edad',
  description: 'Une la proyección hasta la jubilación con el capital de después, sin y con decisiones.',
  inputs: ['proyeccion', 'jubilacion'],
  fn: (proyeccion, jubilacion) => {
    const antes = proyeccion.filter((r) => r.edad !== '').map((r) => ({ edad: +r.edad, sin: +r.liquido, con: +r.liquido_con }));
    const ultima = antes.length ? antes[antes.length - 1].edad : 0;
    const despues = jubilacion.filter((r) => +r.edad > ultima).map((r) => ({ edad: +r.edad, sin: +r.capital, con: +r.capital_con }));
    return antes.concat(despues);
  },
});

// ---------- formatos
const nf = (o) => { try { return new Intl.NumberFormat('es-ES', { ...o, useGrouping: 'always' }); } catch (e) { return new Intl.NumberFormat('es-ES', o); } };
const fEur = nf({ style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
const fNum = nf({ maximumFractionDigits: 0 });
const eur = (v) => (v === '' || v == null || isNaN(+v) ? '—' : fEur.format(+v));
const pct = (v) => (v === '' || v == null || isNaN(+v) ? '—' : Math.round(+v * 100) + ' %');
const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sept', 'oct', 'nov', 'dic'];
const MESES_L = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const mesCorto = (iso) => { const [y, m] = String(iso).split('-'); return MESES[+m - 1] + ' ' + y.slice(2); };
const mesLargo = (iso) => { if (!iso) return '—'; const [y, m] = String(iso).split('-'); return MESES_L[+m - 1] + ' de ' + y; };
const fechaCorta = (iso) => { if (!iso) return '—'; const [y, m, d] = String(iso).split('-'); return +d + ' ' + MESES[+m - 1] + ' ' + y; };
const $ = (s) => document.querySelector(s);
const set = (sel, txt) => document.querySelectorAll(sel).forEach((el) => { el.textContent = txt == null || /undefined|NaN/.test(String(txt)) ? '—' : txt; el.classList.remove('dash-skeleton'); });
const colors = () => dash.colors;
const fail = (el) => { el.replaceChildren(); const d = document.createElement('div'); d.className = 'fail'; d.textContent = 'Failed to load'; el.appendChild(d); };

// ---------- piezas de gráfico
function frame(el, h, left) {
  el.replaceChildren();
  const w = Math.max(260, el.clientWidth || 320);
  const svg = d3.select(el).append('svg').attr('width', '100%').attr('viewBox', `0 0 ${w} ${h}`);
  const tt = d3.select(el).append('div').attr('class', 'tt');
  return { svg, tt, w, h, m: { t: 8, r: 8, b: 26, l: left || 56 } };
}
function yAxis(g, y, w, m, fmt) {
  const ax = g.append('g').attr('transform', `translate(${m.l},0)`).call(d3.axisLeft(y).ticks(4).tickSize(-(w - m.l - m.r)).tickFormat(fmt));
  ax.attr('font-size', null).attr('font-family', null);
  ax.select('.domain').remove();
  ax.selectAll('line').attr('stroke', 'var(--cds-chart-grid)');
  ax.selectAll('text').attr('fill', 'var(--color-fg-muted)').style('font-size', 'var(--cds-font-size-caption)');
}
function xAxis(g, x, h, m, fmt, every) {
  const ax = g.append('g').attr('transform', `translate(0,${h - m.b})`).call(d3.axisBottom(x).tickSize(0).tickPadding(8).tickFormat((d, i) => (every && i % every ? '' : fmt(d))));
  ax.attr('font-size', null).attr('font-family', null);
  ax.select('.domain').attr('stroke', 'var(--cds-chart-axis)');
  ax.selectAll('text').attr('fill', 'var(--color-fg-muted)').style('font-size', 'var(--cds-font-size-caption)');
}
function showTip(f, el, event, head, lines) {
  const box = el.getBoundingClientRect();
  f.tt.html('');
  f.tt.append('div').attr('class', 'h').text(head);
  lines.forEach(([col, name, val]) => {
    const r = f.tt.append('div');
    if (col) r.append('span').attr('class', 'sw').style('background', col);
    r.append('span').text(name + ': ');
    r.append('b').text(val);
  });
  f.tt.style('display', 'block');
  const tw = f.tt.node().offsetWidth;
  let x = event.clientX - box.left + 12;
  if (x + tw > box.width) x = event.clientX - box.left - tw - 12;
  f.tt.style('left', Math.max(0, x) + 'px').style('top', Math.max(0, event.clientY - box.top - 10) + 'px');
}
function legend(el, items) {
  el.replaceChildren();
  items.forEach(([col, name, dashed]) => {
    const s = document.createElement('span');
    const sw = document.createElement('span');
    sw.className = dashed ? 'dash' : 'sw';
    if (dashed) sw.style.color = col; else sw.style.background = col;
    s.append(sw, document.createTextNode(name));
    el.appendChild(s);
  });
}
const yMaxFmt = (v) => (Math.abs(v) >= 10000 ? fNum.format(v / 1000) + ' mil' : fNum.format(v));

// Barras apiladas por mes
function stacked(el, lg, rows, keys, names) {
  const f = frame(el, 240);
  const { svg, w, h, m } = f;
  const C = colors();
  const x = d3.scaleBand(rows.map((r) => r.mes), [m.l, w - m.r]).padding(0.3);
  const st = d3.stack().keys(keys).value((r, k) => Math.max(0, +r[k] || 0))(rows);
  const y = d3.scaleLinear([0, d3.max(st, (s) => d3.max(s, (d) => d[1])) || 1], [h - m.b, m.t]).nice();
  yAxis(svg, y, w, m, yMaxFmt);
  st.forEach((s, i) => {
    svg.append('g').selectAll('rect').data(s).join('rect')
      .attr('x', (d) => x(d.data.mes)).attr('width', x.bandwidth())
      .attr('y', (d) => y(d[1])).attr('height', (d) => Math.max(0, y(d[0]) - y(d[1]) - (i < st.length - 1 ? 0 : 0)))
      .attr('fill', C[i]).attr('stroke', 'var(--color-panel)').attr('stroke-width', 2)
      .attr('rx', i === st.length - 1 ? 3 : 0);
  });
  svg.append('g').selectAll('rect').data(rows).join('rect')
    .attr('x', (r) => x(r.mes) - x.step() * x.padding() / 2).attr('width', x.step()).attr('y', m.t).attr('height', h - m.b - m.t)
    .attr('fill', 'transparent')
    .on('mousemove', (ev, r) => showTip(f, el, ev, mesLargo(r.mes) + (r.completo === 0 ? ' (en curso)' : ''),
      keys.map((k, i) => [C[i], names[i], eur(r[k])]).concat([[null, 'Total', eur(keys.reduce((s, k) => s + (+r[k] || 0), 0))]])))
    .on('mouseleave', () => f.tt.style('display', 'none'));
  xAxis(svg, x, h, m, mesCorto, rows.length > 8 ? 2 : 0);
  legend(lg, keys.map((k, i) => [C[i], names[i]]));
}

// Barras agrupadas ingresos/gastos
function grouped(el, lg, rows) {
  const f = frame(el, 240);
  const { svg, w, h, m } = f;
  const C = colors();
  const keys = ['ingresos', 'gastos'], names = ['Ingresos', 'Gastos'];
  const x = d3.scaleBand(rows.map((r) => r.mes), [m.l, w - m.r]).padding(0.25);
  const xi = d3.scaleBand(keys, [0, x.bandwidth()]).padding(0.12);
  const y = d3.scaleLinear([0, d3.max(rows, (r) => Math.max(+r.ingresos || 0, +r.gastos || 0)) || 1], [h - m.b, m.t]).nice();
  yAxis(svg, y, w, m, yMaxFmt);
  keys.forEach((k, i) => {
    svg.append('g').selectAll('rect').data(rows).join('rect')
      .attr('x', (r) => x(r.mes) + xi(k)).attr('width', xi.bandwidth())
      .attr('y', (r) => y(Math.max(0, +r[k] || 0))).attr('height', (r) => y(0) - y(Math.max(0, +r[k] || 0)))
      .attr('rx', 3).attr('fill', C[i]).attr('opacity', (r) => (r.completo === 0 ? 0.45 : 1));
  });
  svg.append('g').selectAll('rect').data(rows).join('rect')
    .attr('x', (r) => x(r.mes) - x.step() * x.padding() / 2).attr('width', x.step()).attr('y', m.t).attr('height', h - m.b - m.t)
    .attr('fill', 'transparent')
    .on('mousemove', (ev, r) => showTip(f, el, ev, mesLargo(r.mes) + (r.completo === 0 ? ' (en curso)' : ''),
      [[C[0], 'Ingresos', eur(r.ingresos)], [C[1], 'Gastos', eur(r.gastos)], [null, 'Ahorro', eur(r.ahorro)]]))
    .on('mouseleave', () => f.tt.style('display', 'none'));
  xAxis(svg, x, h, m, mesCorto, rows.length > 8 ? 2 : 0);
  legend(lg, keys.map((k, i) => [C[i], names[i]]));
}

// Barras horizontales
function hbars(el, rows) {
  el.replaceChildren();
  const w = Math.max(260, el.clientWidth || 320);
  const rowH = 30, h = Math.max(120, rows.length * rowH + 10);
  const svg = d3.select(el).append('svg').attr('width', '100%').attr('viewBox', `0 0 ${w} ${h}`);
  const tmp = svg.append('text').style('font-size', 'var(--cds-font-size-caption)');
  const labW = d3.max(rows, (r) => { tmp.text(r.grupo); return tmp.node().getComputedTextLength(); }) || 60;
  const valW = d3.max(rows, (r) => { tmp.text(eur(r.media)); return tmp.node().getComputedTextLength(); }) || 50;
  tmp.remove();
  const x = d3.scaleLinear([0, d3.max(rows, (r) => r.media) || 1], [labW + 12, w - valW - 10]);
  const C = colors();
  const g = svg.selectAll('g').data(rows).join('g').attr('transform', (r, i) => `translate(0,${i * rowH + 6})`);
  g.append('text').attr('x', labW).attr('y', 13).attr('text-anchor', 'end').attr('fill', 'var(--color-fg)').style('font-size', 'var(--cds-font-size-caption)').text((r) => r.grupo);
  g.append('rect').attr('x', x(0)).attr('y', 4).attr('height', 14).attr('rx', 4).attr('width', (r) => Math.max(2, x(r.media) - x(0))).attr('fill', C[0]);
  g.append('text').attr('x', (r) => x(r.media) + 6).attr('y', 15).attr('fill', 'var(--color-fg-muted)').style('font-size', 'var(--cds-font-size-caption)').style('font-variant-numeric', 'tabular-nums').text((r) => eur(r.media));
}

// Líneas por edad
function lines(el, lg, rows, edadJub) {
  const f = frame(el, 280, 64);
  const { svg, w, h, m } = f;
  const C = colors();
  const x = d3.scaleLinear(d3.extent(rows, (r) => r.edad), [m.l, w - m.r]);
  const y = d3.scaleLinear([0, d3.max(rows, (r) => Math.max(r.sin, r.con)) || 1], [h - m.b, m.t]).nice();
  yAxis(svg, y, w, m, yMaxFmt);
  const ax = svg.append('g').attr('transform', `translate(0,${h - m.b})`).call(d3.axisBottom(x).ticks(Math.min(8, Math.floor(w / 60))).tickFormat(d3.format('d')).tickSize(0).tickPadding(8));
  ax.attr('font-size', null).attr('font-family', null);
  ax.select('.domain').attr('stroke', 'var(--cds-chart-axis)');
  ax.selectAll('text').attr('fill', 'var(--color-fg-muted)').style('font-size', 'var(--cds-font-size-caption)');
  if (edadJub) {
    svg.append('line').attr('x1', x(edadJub)).attr('x2', x(edadJub)).attr('y1', m.t).attr('y2', h - m.b).attr('stroke', 'var(--cds-chart-reference)').attr('stroke-dasharray', '4 4');
    svg.append('text').attr('x', x(edadJub) + 6).attr('y', m.t + 12).attr('fill', 'var(--color-fg-muted)').style('font-size', 'var(--cds-font-size-caption)').text('Jubilación');
  }
  [['sin', C[0], null], ['con', C[1], '6 5']].forEach(([k, col, dsh]) => {
    svg.append('path').datum(rows).attr('fill', 'none').attr('stroke', col).attr('stroke-width', 2)
      .attr('stroke-linecap', 'round').attr('stroke-linejoin', 'round').attr('stroke-dasharray', dsh)
      .attr('d', d3.line().x((r) => x(r.edad)).y((r) => y(Math.max(0, r[k]))));
  });
  const dot = svg.append('g').style('display', 'none');
  const rule = dot.append('line').attr('y1', m.t).attr('y2', h - m.b).attr('stroke', 'var(--cds-chart-axis)');
  const d1 = dot.append('circle').attr('r', 4).attr('fill', C[0]).attr('stroke', 'var(--color-panel)').attr('stroke-width', 2);
  const d2 = dot.append('circle').attr('r', 4).attr('fill', C[1]).attr('stroke', 'var(--color-panel)').attr('stroke-width', 2);
  svg.append('rect').attr('x', m.l).attr('y', m.t).attr('width', w - m.l - m.r).attr('height', h - m.t - m.b).attr('fill', 'transparent')
    .on('mousemove', (ev) => {
      const [px] = d3.pointer(ev);
      const e = Math.round(x.invert(px));
      const r = rows.find((q) => q.edad === e) || rows[d3.bisectCenter(rows.map((q) => q.edad), e)];
      if (!r) return;
      dot.style('display', null);
      rule.attr('x1', x(r.edad)).attr('x2', x(r.edad));
      d1.attr('cx', x(r.edad)).attr('cy', y(Math.max(0, r.sin)));
      d2.attr('cx', x(r.edad)).attr('cy', y(Math.max(0, r.con)));
      showTip(f, el, ev, r.edad + ' años', [[C[0], 'Sin decisiones', eur(r.sin)], [C[1], 'Con tus decisiones', eur(r.con)]]);
    })
    .on('mouseleave', () => { dot.style('display', 'none'); f.tt.style('display', 'none'); });
  legend(lg, [[C[0], 'Sin decisiones'], [C[1], 'Con tus decisiones (pestaña DECISIONES)', true]]);
}

// Tablas
function fillTable(sel, rows, cells) {
  const tb = $(sel + ' tbody');
  tb.replaceChildren();
  rows.forEach((r) => {
    const tr = document.createElement('tr');
    cells(r).forEach(([txt, cls, node]) => {
      const td = document.createElement('td');
      if (cls) td.className = cls;
      if (node) td.appendChild(node); else td.textContent = txt;
      tr.appendChild(td);
    });
    tb.appendChild(tr);
  });
}

// ---------- dibujo
dash.onData(() => {
  const k = dash.data('kpi');
  const kv = {};
  if (k.status === 'ok') k.data.forEach((r) => { kv[r.clave] = r.valor; });
  const okK = k.status === 'ok';
  const euroKeys = { '#k-liq': 'patrimonio_liquido', '#k-tot': 'patrimonio_total', '#k-ahorro': 'ahorro_medio', '#k-ing': 'ingresos_medio', '#k-casa': 'casa_media', '#k-vsaldo': 'viajes_saldo_fin_anio', '#k-p67': 'patrimonio_jub', '#k-p67h': 'patrimonio_jub_hoy', '.t-liq': 'patrimonio_liquido', '.t-aho': 'ahorro_medio' };
  Object.entries(euroKeys).forEach(([s, key]) => set(s, okK ? eur(kv[key]) : '—'));
  set('#k-lib', okK ? pct(kv.libertad) : '—');
  set('.t-lib', okK ? pct(kv.libertad) : '—');
  set('#k-nivel', okK ? String(kv.nivel || '') : '—');
  set('#k-vanio', okK ? String(kv.viajes_anio) : '—');
  set('#k-vaviso', okK ? String(kv.viajes_aviso || '') : '—');
  set('#k-fecha', okK ? mesLargo(kv.fecha_datos) : '—');
  set('#k-ej', okK ? String(kv.edad_jubilacion) : '—');
  set('#dec-texto', okK ? String(kv.decisiones_texto || '') : '—');
  set('#k-revisar', okK ? fNum.format(kv.por_revisar || 0) : '—');
  $('#chip-revisar').hidden = !(okK && +kv.por_revisar > 0);
  const du = dash.data('dura');
  set('#k-dura', du.status === 'ok' && du.data[0] ? du.data[0].texto : '—');

  const ms = dash.data('meses');
  const elS = $('#ch-saldos'), elF = $('#ch-flujo');
  if (ms.status === 'ok' && ms.data.length) {
    stacked(elS, $('#lg-saldos'), ms.data, ['cuenta_1', 'cuenta_2', 'cuenta_3'], [1, 2, 3].map((n) => (okK && kv['nombre_cuenta_' + n]) || 'Cuenta ' + n));
    grouped(elF, $('#lg-flujo'), ms.data);
  } else if (ms.status === 'error') { fail(elS); fail(elF); } else { elS.replaceChildren(); elF.replaceChildren(); }

  const gg = dash.data('gasto_grupos'), elG = $('#ch-grupos');
  if (gg.status === 'ok') hbars(elG, gg.data); else if (gg.status === 'error') fail(elG); else elG.replaceChildren();

  const cm = dash.data('casa_meses'), elC = $('#ch-casa');
  const ca = dash.data('casa');
  if (cm.status === 'ok' && ca.status === 'ok') {
    const conceptos = ca.data.filter((r) => +r.total12 > 0).map((r) => r.concepto);
    const desde = ms.status === 'ok' && ms.data[0] ? ms.data[0].mes : '';
    const filas = cm.data.filter((r) => !desde || r.mes >= desde).map((r) => ({ ...r, completo: ms.status === 'ok' ? ((ms.data.find((q) => q.mes === r.mes) || {}).completo ?? 1) : 1 }));
    stacked(elC, $('#lg-casa'), filas, conceptos, conceptos);
  } else if (cm.status === 'error') fail(elC); else elC.replaceChildren();
  if (ca.status === 'ok') fillTable('#tb-casa', ca.data.filter((r) => +r.total12 > 0), (r) => [[r.concepto], [eur(r.media), 'num'], [eur(r.ultimo), 'num']]);

  const tr = dash.data('trayectoria'), elT = $('#ch-futuro');
  if (tr.status === 'ok' && tr.data.length) lines(elT, $('#lg-futuro'), tr.data, okK ? +kv.edad_jubilacion : null);
  else if (tr.status === 'error') fail(elT); else elT.replaceChildren();

  const hi = dash.data('hitos');
  if (hi.status === 'ok') {
    const C = colors();
    fillTable('#tb-hitos', hi.data, (r) => {
      const bar = document.createElement('div'); bar.className = 'bar';
      const i = document.createElement('i'); i.style.width = Math.round(Math.min(1, +r.pct || 0) * 100) + '%'; i.style.background = C[0];
      bar.appendChild(i);
      return [[r.hito], ['', '', bar], [r.anio === '' ? '' : String(r.anio)], [r.estado]];
    });
  }

  const vi = dash.data('viajes');
  if (vi.status === 'ok') fillTable('#tb-viajes', vi.data, (r) => [[r.viaje], [fechaCorta(r.fecha)], [eur(r.tu_parte), 'num'], [r.estado], [r.saldo_tras === '' ? '' : eur(r.saldo_tras), 'num'], [r.llega]]);

  const es = dash.data('viajes_escenarios');
  if (es.status === 'ok') {
    const fmt = (r, v) => (typeof v !== 'number' ? String(v) : /^%|libertad/i.test(r.concepto) || /%/.test(r.concepto) ? pct(v) : eur(v));
    fillTable('#tb-esc', es.data, (r) => [[r.concepto], [fmt(r, r.ahora), 'num'], [fmt(r, r.a), 'num'], [fmt(r, r.b), 'num'], [fmt(r, r.c), 'num']]);
  }
});

const named = document.getElementById(location.hash.slice(1));
if (named) named.scrollIntoView();

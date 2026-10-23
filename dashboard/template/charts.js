/* ===========================================================================
   CHARTS (hand-built SVG - no libraries, works offline)
   Specs: thin marks (<=24px), 4px rounded data-end, hairline recessive grid,
   one axis per chart, tooltips on every mark, legend whenever >=2 series.
   =========================================================================== */
function niceTicks(lo, hi, n = 5) {
  if (!isFinite(lo) || !isFinite(hi)) return [0, 1];
  if (lo === hi) { const p = Math.abs(lo) || 1; lo -= p * .5; hi += p * .5; }
  const span = hi - lo, raw = span / n, mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag, step = (norm >= 5 ? 10 : norm >= 2 ? 5 : norm >= 1 ? 2 : 1) * mag;   // 1,2,5,10
  const t0 = Math.floor(lo / step + 1e-9) * step, out = [];
  let v = t0; for (; v <= hi + step * 1e-9; v += step) out.push(+v.toPrecision(12));
  if (out[out.length - 1] < hi - step * 1e-9) out.push(+v.toPrecision(12));
  return out;
}
function lin(d0, d1, r0, r1) { const k = (r1 - r0) / ((d1 - d0) || 1); return v => r0 + (v - d0) * k; }
function txt(x, y, s, o = {}) { return sv('text', Object.assign({ x, y, 'font-size': 11, class: 'axis-t' }, o), s); }
function chartHost(el) { el.classList.add('chart'); return Math.max(260, el.clientWidth || 600); }
function mixColor(a, b, t) {        // a,b = '#rrggbb' or 'rgb(...)'
  const pa = parseRGB(a), pb = parseRGB(b); const c = pa.map((v, i) => Math.round(v + (pb[i] - v) * t));
  return `rgb(${c[0]},${c[1]},${c[2]})`;
}
function parseRGB(s) {
  s = s.trim();
  if (s[0] === '#') { if (s.length === 4) s = '#' + [...s.slice(1)].map(c => c + c).join(''); return [1, 3, 5].map(i => parseInt(s.slice(i, i + 2), 16)); }
  const m = s.match(/[\d.]+/g); return m ? m.slice(0, 3).map(Number) : [128, 128, 128];
}
function lum(rgb) { const f = c => { c /= 255; return c <= .03928 ? c / 12.92 : Math.pow((c + .055) / 1.055, 2.4); }; return .2126 * f(rgb[0]) + .7152 * f(rgb[1]) + .0722 * f(rgb[2]); }
function inkOn(bg) { return lum(parseRGB(bg)) > .42 ? '#0b0b0b' : '#ffffff'; }
function divColor(v, maxAbs) {          // blue (positive) <-> neutral <-> red (negative)
  if (!isFinite(v)) return cssVar('--wash') || '#eee';
  const t = Math.min(1, Math.abs(v) / (maxAbs || 1)), mid = cssVar('--mid'), pole = cssVar(v >= 0 ? '--fav' : '--unf');
  return mixColor(mid, pole, Math.pow(t, .8));
}
function legendEl(items) {
  return h('div', { class: 'legend' }, items.map(it => h('span', null, it.type === 'bar' ? h('b', { style: { background: it.color } }) : h('i', { style: { borderTopColor: it.color } }), it.label)));
}
function roundedBarPath(a, b, c, w, r, horizontal) {      // data end rounded (4px), baseline end square
  if (horizontal) {      // a = baseline x, b = top y of the bar, c = tip x, w = thickness
    const dir = c >= a ? 1 : -1, len = Math.abs(c - a), rr = Math.min(r, len, w / 2);
    if (len < .5) return '';
    const yt = b, yb = b + w;
    return `M${a},${yt}H${c - dir * rr}Q${c},${yt} ${c},${yt + rr}V${yb - rr}Q${c},${yb} ${c - dir * rr},${yb}H${a}Z`;
  }
  const len = Math.abs(c - b), rr = Math.min(r, len, w / 2);   // vertical: a = left x, b = baseline y, c = tip y
  if (len < .5) return '';
  const e = c < b ? rr : -rr, xl = a, xr = a + w;
  return `M${xl},${b}V${c + e}Q${xl},${c} ${xl + rr},${c}H${xr - rr}Q${xr},${c} ${xr},${c + e}V${b}Z`;
}

/* ---------- 1. Overview trend: two aligned panels, one axis each ---------- */
function trendChart(host, o) {
  // o: {tc[],ac[],vc[],pct[],att[],inSel[], onMonth(m), selected}
  const W = chartHost(host), ml = 58, mr = 14, hA = 190, hB = 110, gap = 34, top = 6, xAx = 26, H = top + hA + gap + hB + xAx;
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Monthly target versus actual sales and variance' });
  const n = o.ac.length, pw = W - ml - mr, band = pw / n, X = i => ml + band * (i + .5);
  const allV = o.ac.concat(o.tc).map(v => v / 1e8);   // $M scale: cents/1e8 = millions
  let lo = Math.min(...allV), hi = Math.max(...allV); const pad = (hi - lo) * .08 || 1;
  const tA = niceTicks(lo - pad, hi + pad, 4), yA = lin(tA[0], tA[tA.length - 1], top + hA, top);
  tA.forEach(t => { svg.append(sv('line', { x1: ml, x2: W - mr, y1: yA(t), y2: yA(t), stroke: 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, yA(t) + 4, '$' + t.toFixed(t % 1 ? 1 : 0) + 'M', { 'text-anchor': 'end' })); });
  // gap wash between lines
  for (let i = 0; i < n - 1; i++) {
    const s = (o.vc[i] + o.vc[i + 1]) / 2, c = s >= 0 ? 'var(--fav)' : 'var(--unf)';
    const pts = [[X(i), yA(o.ac[i] / 1e8)], [X(i + 1), yA(o.ac[i + 1] / 1e8)], [X(i + 1), yA(o.tc[i + 1] / 1e8)], [X(i), yA(o.tc[i] / 1e8)]];
    svg.append(sv('polygon', { points: pts.map(p => p.join(',')).join(' '), fill: c, 'fill-opacity': .16 }));
  }
  const path = a => a.map((v, i) => (i ? 'L' : 'M') + X(i) + ',' + yA(v / 1e8)).join('');
  svg.append(sv('path', { d: path(o.tc), fill: 'none', stroke: 'var(--target)', 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
  svg.append(sv('path', { d: path(o.ac), fill: 'none', stroke: 'var(--s1)', 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
  [['tc', 'var(--target)'], ['ac', 'var(--s1)']].forEach(([k, c]) => svg.append(sv('circle', { cx: X(n - 1), cy: yA(o[k][n - 1] / 1e8), r: 4, fill: c, stroke: 'var(--surface)', 'stroke-width': 2 })));
  // panel B: variance columns from a zero baseline
  const yb0 = top + hA + gap; const vv = o.vc.map(v => v / 1e5);   // $K
  const vlo = Math.min(0, ...vv), vhi = Math.max(0, ...vv), tB = niceTicks(vlo * 1.12, vhi * 1.12 || 1, 3);
  const yB = lin(tB[0], tB[tB.length - 1], yb0 + hB, yb0);
  svg.append(txt(ml, yb0 - 10, 'Variance (actual – target), $K', { 'font-weight': 600 }));
  tB.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: yB(t), y2: yB(t), stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, yB(t) + 4, numF(t), { 'text-anchor': 'end' })));
  const bw = Math.min(24, band * .62), y0 = yB(0);
  const bars = [];
  for (let i = 0; i < n; i++) {
    const y1 = yB(vv[i]), c = vv[i] >= 0 ? 'var(--fav)' : 'var(--unf)';
    const p = sv('path', { d: roundedBarPath(X(i) - bw / 2, y0, y1, bw, 4, false), fill: c, opacity: o.inSel[i] ? 1 : .28 }); bars.push(p); svg.append(p);
  }
  // selective labels: extremes
  const iMin = vv.indexOf(Math.min(...vv)), iMax = vv.indexOf(Math.max(...vv));
  [iMin, iMax].filter((v, k, a) => a.indexOf(v) === k).forEach(i => svg.append(txt(X(i), vv[i] < 0 ? yB(vv[i]) + 13 : yB(vv[i]) - 5, numF(vv[i]), { 'text-anchor': 'middle', 'font-size': 10.5, style: 'fill:var(--ink2)' })));
  // x axis
  const every = band < 52 ? 2 : 1;
  for (let i = 0; i < n; i++) if (i % every === 0) svg.append(txt(X(i), yb0 + hB + 17, MLAB[i].replace(' ', ' '), { 'text-anchor': 'middle', 'font-size': 10.5 }));
  // hover
  const cross = sv('line', { y1: top, y2: yb0 + hB, stroke: 'var(--ink2)', 'stroke-width': 1, opacity: 0 });
  const rect = sv('rect', { x: ml, y: 0, width: pw, height: H, fill: 'transparent', style: 'cursor:pointer' });
  svg.append(cross, rect);
  const idxAt = ev => { const r = svg.getBoundingClientRect(); const x = (ev.clientX - r.left) * (W / r.width); return Math.max(0, Math.min(n - 1, Math.floor((x - ml) / band))); };
  rect.addEventListener('pointermove', ev => {
    const i = idxAt(ev); cross.setAttribute('x1', X(i)); cross.setAttribute('x2', X(i)); cross.setAttribute('opacity', .5);
    showTip(ev, MLONG[o.m[i]], [
      { color: 'var(--target)', label: 'Target', value: usd(o.tc[i] / 100) }, { color: 'var(--s1)', label: 'Actual', value: usd(o.ac[i] / 100) },
      { label: 'Variance', value: usd(o.vc[i] / 100) + '  (' + pctF(o.pct[i], 2, true) + ')' }, { label: 'Qty attainment', value: pctF(o.att[i]) },
      { label: 'Status (month total)', value: ST_GLYPH[o.st[i]] + ' ' + ST_NAME[o.st[i]] + (o.mat[i] ? '  ◆ material' : '') }], 'Click to filter to this month');
  });
  rect.addEventListener('pointerleave', () => { cross.setAttribute('opacity', 0); hideTip(); });
  rect.addEventListener('click', ev => o.onMonth && o.onMonth(o.m[idxAt(ev)]));
  host.replaceChildren(legendEl([{ label: 'Target', color: 'var(--target)' }, { label: 'Actual', color: 'var(--s1)' }, { label: 'Gap: actual above target', color: 'var(--fav)', type: 'bar' }, { label: 'Gap: actual below target', color: 'var(--unf)', type: 'bar' }]), svg);
}

/* ---------- 2. Horizontal diverging bars (clickable) ---------- */
function hbars(host, items, o = {}) {
  const W = chartHost(host), rowH = o.rowH || 30, bh = Math.min(o.barH || 18, rowH - 8), mt = 4, H = items.length * rowH + mt + (o.axis === false ? 4 : 22);
  const labW = Math.min(o.labW || 150, Math.max(70, Math.max(0, ...items.map(i => i.label.length)) * 6.4 + 6));
  const vals = items.map(i => i.value), hasNeg = vals.some(v => v < 0), hasPos = vals.some(v => v > 0);
  const lo = Math.min(0, ...vals), hi = Math.max(0, ...vals);
  const txtW = sel => Math.max(0, ...items.filter(sel).map(i => String(i.txt).length)) * 6.6 + 12;   // room for value labels
  const xL = labW + 10 + (hasNeg ? Math.max(60, txtW(i => i.value < 0)) : 0), xR = W - (hasPos ? Math.max(60, txtW(i => i.value >= 0)) : 16), x = lin(lo, hi || 1, xL, xR), x0 = x(0);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Bar chart' });
  niceTicks(lo, hi, 3).forEach(t => { if (t < lo - 1e-9 || t > hi + 1e-9) return; svg.append(sv('line', { x1: x(t), x2: x(t), y1: mt, y2: H - 18, stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 })); if (o.axis !== false) svg.append(txt(x(t), H - 4, (o.tickFmt || (v => usdC(v, 1)))(t), { 'text-anchor': 'middle', 'font-size': 10.5 })); });
  const anySel = items.some(i => i.selected);
  items.forEach((it, k) => {
    const y = mt + k * rowH + (rowH - bh) / 2, c = it.color || (it.value >= 0 ? 'var(--fav)' : 'var(--unf)');
    const g = sv('g', { class: o.onClick ? 'clk' : '', opacity: anySel && !it.selected ? .38 : 1 });
    g.append(sv('rect', { x: 0, y: mt + k * rowH, width: W, height: rowH, fill: 'transparent' }));
    g.append(txt(labW, y + bh / 2 + 4, it.label, { 'text-anchor': 'end', 'font-size': 12, style: 'fill:var(--ink)' }));
    const bar = sv('path', { d: roundedBarPath(x0, y, x(it.value), bh, 4, true), fill: c }); g.append(bar);
    const tipX = x(it.value), pos = it.value >= 0;
    g.append(txt(pos ? tipX + 6 : tipX - 6, y + bh / 2 + 4, it.txt, { 'text-anchor': pos ? 'start' : 'end', 'font-size': 11.5, style: 'fill:var(--ink)', 'font-weight': 600 }));
    g.addEventListener('pointermove', ev => { bar.style.filter = 'brightness(1.12)'; const t = o.tip ? o.tip(it) : { title: it.label, rows: [{ label: 'Value', value: it.txt }] }; showTip(ev, t.title, t.rows, o.onClick ? (it.selected ? 'Click to remove from filter' : 'Click to filter') : t.foot); });
    g.addEventListener('pointerleave', () => { bar.style.filter = ''; hideTip(); });
    if (o.onClick) g.addEventListener('click', ev => o.onClick(it, ev));
    focusable(g, it.label + ': ' + it.txt, o.onClick ? ev => o.onClick(it, ev) : null, () => { const t = o.tip ? o.tip(it) : { title: it.label, rows: [{ label: 'Value', value: it.txt }] }; return { title: t.title, rows: t.rows, foot: o.onClick ? 'Press Enter to filter' : '' }; });
    svg.append(g);
  });
  host.replaceChildren(svg);
}

/* ---------- 3. Waterfall (floating bars; truncated axis is labelled) ---------- */
function waterfall(host, steps, o = {}) {
  // steps: [{label, value, kind:'total'|'delta', color}]
  const W = chartHost(host), H = o.height || 300, ml = 60, mr = 12, mt = 26, mb = 44;
  let run = 0; const bars = steps.map(s => { if (s.kind === 'total') { run = s.value; return { s, a: 0, b: s.value, end: run }; } const a = run; run += s.value; return { s, a, b: run, end: run }; });
  const all = bars.flatMap(b => b.s.kind === 'total' ? [b.b] : [b.a, b.b, b.end]); let lo = Math.min(...all), hi = Math.max(...all);
  const totals = steps.filter(s => s.kind === 'total').map(s => s.value);
  let trunc = false, base = 0;
  if (totals.length && Math.min(...totals) > 0 && (hi - lo) < 0.5 * hi) { const dmin = Math.min(...bars.filter(b => b.s.kind === 'delta').flatMap(b => [b.a, b.b])); base = Math.min(dmin, ...totals); const sp = hi - base; base = base - sp * .6; trunc = base > 0; lo = trunc ? base : 0; }
  else lo = Math.min(0, lo);
  const ticks = niceTicks(lo, hi * 1.0, 4), y = lin(ticks[0], ticks[ticks.length - 1], H - mb, mt);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Waterfall chart' });
  ticks.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, y(t) + 4, usdC(t, ticks[1] - ticks[0] >= 1e6 ? 0 : 1), { 'text-anchor': 'end' })));
  const n = steps.length, band = (W - ml - mr) / n, bw = Math.min(46, band * .6), X = i => ml + band * (i + .5);
  bars.forEach((b, i) => {
    const top = Math.max(b.a, b.b), bot = Math.min(b.a, b.b), yT = y(Math.max(top, ticks[0])), yB = y(Math.max(bot, ticks[0]));
    const c = b.s.color || (b.s.kind === 'total' ? 'var(--s1)' : (b.s.value >= 0 ? 'var(--fav)' : 'var(--unf)'));
    const r = sv('rect', { x: X(i) - bw / 2, y: yT, width: bw, height: Math.max(1.5, yB - yT), rx: 3, fill: c, class: 'clk' });
    r.addEventListener('pointermove', ev => showTip(ev, b.s.label, [{ label: b.s.kind === 'total' ? 'Total' : 'Change', value: usd(b.s.value) }, { label: 'Running total', value: usd(b.end) }].concat(b.s.extra || []), b.s.foot));
    r.addEventListener('pointerleave', hideTip); svg.append(r);
    focusable(r, b.s.label + ': ' + usd(b.s.value), null, () => ({ title: b.s.label, rows: [{ label: b.s.kind === 'total' ? 'Total' : 'Change', value: usd(b.s.value) }, { label: 'Running total', value: usd(b.end) }].concat(b.s.extra || []), foot: b.s.foot }));
    if (i < n - 1) svg.append(sv('line', { x1: X(i) + bw / 2, x2: X(i + 1) - bw / 2, y1: y(b.end), y2: y(b.end), stroke: 'var(--axis)', 'stroke-width': 1 }));
    const lab = b.s.kind === 'total' ? usdC(b.s.value) : usdS(b.s.value);
    const up = b.s.kind === 'total' || b.s.value >= 0;
    svg.append(txt(X(i), up ? yT - 6 : yB + 14, lab, { 'text-anchor': 'middle', 'font-size': band < 70 ? 10 : 11.5, 'font-weight': 600, style: 'fill:var(--ink)' }));
    const words = b.s.label.split(' '); const l1 = words.slice(0, Math.ceil(words.length / 2)).join(' '), l2 = words.slice(Math.ceil(words.length / 2)).join(' ');
    svg.append(txt(X(i), H - mb + 16, l1, { 'text-anchor': 'middle', 'font-size': band < 70 ? 9.5 : 11 }), txt(X(i), H - mb + 29, l2, { 'text-anchor': 'middle', 'font-size': band < 70 ? 9.5 : 11 }));
  });
  host.replaceChildren(svg);
  if (trunc) host.append(h('div', { class: 'cap' }, 'Axis starts at ' + usdC(ticks[0], 0) + ' so the changes are visible – bar heights are not proportional to totals.'));
}

/* ---------- 4. Heatmap (diverging, symmetric around zero) ---------- */
function heatmap(host, o) {
  // o: {rows[], cols[], colSub[], vals[r][c], fmt(v), cellTip(r,c)->{title,rows}, onCell(r,c), selR:Set, selC:Set, rowMarg[], colMarg[], diverging, legendFmt}
  const W = chartHost(host), nr = o.rows.length, nc = o.cols.length, labW = o.labW || 118, mt = o.colSub ? 36 : 22, cw = (W - labW - 8) / nc;
  const ch = o.cellH || (nc > 12 ? 30 : 44), H = mt + nr * ch + 30;
  const all = o.vals.flat().filter(isFinite); const mx = o.maxAbs || Math.max(1e-9, ...all.map(Math.abs)), mn = Math.min(...all), mxv = Math.max(...all);
  const seqLo = cssVar('--mid'), seqHi = cssVar('--s1');
  const colorOf = v => o.diverging === false ? mixColor(seqLo, seqHi, Math.pow(Math.max(0, (v - mn) / ((mxv - mn) || 1)), .9)) : divColor(v, mx);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Heatmap' });
  o.cols.forEach((c, j) => { svg.append(txt(labW + cw * (j + .5), 14, c, { 'text-anchor': 'middle', 'font-size': cw < 40 ? 10 : 11.5, style: 'fill:var(--ink2)' })); if (o.colSub && o.colSub[j]) svg.append(txt(labW + cw * (j + .5), 27, o.colSub[j], { 'text-anchor': 'middle', 'font-size': 10 })); });
  const anyR = o.selR && o.selR.size, anyC = o.selC && o.selC.size;
  o.rows.forEach((r, i) => {
    svg.append(txt(labW - 8, mt + i * ch + ch / 2 + 4, r, { 'text-anchor': 'end', 'font-size': 12, style: 'fill:var(--ink)' }));
    o.cols.forEach((c, j) => {
      const v = o.vals[i][j], col = colorOf(v), dim = (anyR && !o.selR.has(i)) || (anyC && !o.selC.has(j));
      const g = sv('g', { class: o.onCell ? 'clk' : '', opacity: dim ? .35 : 1 });
      const rc = sv('rect', { x: labW + cw * j + 1, y: mt + i * ch + 1, width: cw - 2, height: ch - 2, rx: 3, fill: col });
      g.append(rc);
      if (cw >= 44 || (cw >= 30 && ch >= 28 && o.fmtShort)) { const s = (cw >= 44 ? o.fmt : o.fmtShort)(v); if (s) g.append(txt(labW + cw * (j + .5), mt + i * ch + ch / 2 + 4, s, { 'text-anchor': 'middle', 'font-size': cw >= 44 ? 11.5 : 9.5, style: `fill:${inkOn(col)}`, 'font-weight': 600 })); }
      g.addEventListener('pointermove', ev => { rc.style.stroke = 'var(--ink)'; rc.style.strokeWidth = 2; const t = o.cellTip(i, j); showTip(ev, t.title, t.rows, t.foot || (o.onCell ? 'Click to filter' : '')); });
      g.addEventListener('pointerleave', () => { rc.style.stroke = ''; rc.style.strokeWidth = ''; hideTip(); });
      if (o.onCell) g.addEventListener('click', () => o.onCell(i, j));
      focusable(g, r + ' · ' + c + ': ' + (isFinite(v) ? o.fmt(v) : 'no data'), o.onCell ? () => o.onCell(i, j) : null, () => { const t = o.cellTip(i, j); return { title: t.title, rows: t.rows, foot: o.onCell ? 'Press Enter to filter' : '' }; });
      svg.append(g);
    });
  });
  // scale legend
  const ly = mt + nr * ch + 12, lw = Math.min(260, W - labW), lx = labW; const steps = 40;
  for (let k = 0; k < steps; k++) { const v = o.diverging === false ? mn + (mxv - mn) * k / (steps - 1) : -mx + 2 * mx * k / (steps - 1); svg.append(sv('rect', { x: lx + lw * k / steps, y: ly, width: lw / steps + .5, height: 8, fill: colorOf(v) })); }
  const lf = o.legendFmt || (v => v.toFixed(1));
  svg.append(txt(lx, ly + 20, o.diverging === false ? lf(mn) : lf(-mx), { 'font-size': 10.5 }), txt(lx + lw, ly + 20, o.diverging === false ? lf(mxv) : lf(mx), { 'text-anchor': 'end', 'font-size': 10.5 }));
  if (o.diverging !== false) svg.append(txt(lx + lw / 2, ly + 20, lf(0), { 'text-anchor': 'middle', 'font-size': 10.5 }));
  svg.append(txt(lx + lw + 12, ly + 8, o.legendNote || '', { 'font-size': 10.5 }));
  host.replaceChildren(svg);
}

/* ---------- 5. Multi-series line / single-series column chart with crosshair ---------- */
function lineChart(host, o) {
  // o: {x[], xFull[], series:[{id,name,color,values[]}], fmt(v), tickFmt, zero:boolean, mode:'line'|'bar', onX(i), selX:Uint8Array, height, refs:[{y,label}]}
  const W = chartHost(host), H = o.height || 260, ml = o.ml || 58, mr = 16, mt = 10, mb = 28, n = o.x.length, pw = W - ml - mr, band = pw / n, X = i => ml + band * (i + .5);
  const vals = o.series.flatMap(s => s.values).filter(isFinite); let lo = Math.min(...vals, ...(o.refs || []).map(r => r.y)), hi = Math.max(...vals, ...(o.refs || []).map(r => r.y));
  if (o.zero || o.mode === 'bar') { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
  const pad = (hi - lo) * .08 || 1; const ticks = niceTicks(lo - (o.mode === 'bar' || o.zero ? 0 : pad), hi + pad, 5), y = lin(ticks[0], ticks[ticks.length - 1], H - mb, mt);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Line chart' });
  ticks.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, y(t) + 4, (o.tickFmt || o.fmt)(t), { 'text-anchor': 'end' })));
  (o.refs || []).forEach(r => { svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(r.y), y2: y(r.y), stroke: 'var(--ink2)', 'stroke-width': 1, opacity: .55 }), txt(W - mr - 4, y(r.y) - 4, r.label, { 'text-anchor': 'end', 'font-size': 10.5 })); });
  const every = band < 52 ? 2 : 1; for (let i = 0; i < n; i++) if (i % every === 0) svg.append(txt(X(i), H - 8, o.x[i].replace(' ', ' '), { 'text-anchor': 'middle', 'font-size': 10.5 }));
  if (o.mode === 'bar') {
    const s = o.series[0], bw = Math.min(24, band * .62), y0 = y(0);
    s.values.forEach((v, i) => { if (!isFinite(v)) return; svg.append(sv('path', { d: roundedBarPath(X(i) - bw / 2, y0, y(v), bw, 4, false), fill: v >= 0 ? 'var(--fav)' : 'var(--unf)', opacity: !o.selX || o.selX[i] ? 1 : .28 })); });
  } else {
    o.series.forEach(s => {
      let d = '', pen = false; s.values.forEach((v, i) => { if (!isFinite(v)) { pen = false; return; } d += (pen ? 'L' : 'M') + X(i) + ',' + y(v); pen = true; });
      svg.append(sv('path', { d, fill: 'none', stroke: s.color, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round', opacity: s.dim ? .35 : 1 }));
      let li = -1; s.values.forEach((v, i) => { if (isFinite(v)) li = i; });
      if (li >= 0) svg.append(sv('circle', { cx: X(li), cy: y(s.values[li]), r: 4, fill: s.color, stroke: 'var(--surface)', 'stroke-width': 2 }));
    });
  }
  const cross = sv('line', { y1: mt, y2: H - mb, stroke: 'var(--ink2)', 'stroke-width': 1, opacity: 0 }), dots = o.series.map(s => sv('circle', { r: 4, fill: s.color, stroke: 'var(--surface)', 'stroke-width': 2, opacity: 0 }));
  const rect = sv('rect', { x: ml, y: 0, width: pw, height: H, fill: 'transparent', style: o.onX ? 'cursor:pointer' : '' });
  svg.append(cross, ...dots, rect);
  const idxAt = ev => { const r = svg.getBoundingClientRect(); const px = (ev.clientX - r.left) * (W / r.width); return Math.max(0, Math.min(n - 1, Math.floor((px - ml) / band))); };
  rect.addEventListener('pointermove', ev => {
    const i = idxAt(ev); cross.setAttribute('x1', X(i)); cross.setAttribute('x2', X(i)); cross.setAttribute('opacity', .5);
    dots.forEach((d, k) => { const v = o.series[k].values[i]; if (isFinite(v) && o.mode !== 'bar') { d.setAttribute('cx', X(i)); d.setAttribute('cy', y(v)); d.setAttribute('opacity', 1); } else d.setAttribute('opacity', 0); });
    const rows = o.series.map(s => ({ color: s.color, label: s.name, value: o.fmt(s.values[i]) })).sort((a, b) => 0);
    showTip(ev, (o.xFull || o.x)[i], rows.concat(o.extraTip ? o.extraTip(i) : []), o.onX ? 'Click to filter to this month' : '');
  });
  rect.addEventListener('pointerleave', () => { cross.setAttribute('opacity', 0); dots.forEach(d => d.setAttribute('opacity', 0)); hideTip(); });
  if (o.onX) rect.addEventListener('click', ev => o.onX(idxAt(ev)));
  const kids = []; if (o.series.length >= 2) kids.push(legendEl(o.series.map(s => ({ label: s.name, color: s.color })))); kids.push(svg);
  host.replaceChildren(...kids);
}

/* ---------- 6. Bubble / scatter with nearest-point hit layer ---------- */
function bubble(host, o) {
  // o: {pts:[{id,label,x,y,r,color,tipRows,selected}], xFmt,yFmt, xTitle,yTitle, bands:[{y0,y1,label}], labelIds:Set, onClick(p), height}
  const W = chartHost(host), H = o.height || 340, ml = 58, mr = 16, mt = 12, mb = 40;
  const xs = o.pts.map(p => p.x), ys = o.pts.map(p => p.y);
  const tx = niceTicks(Math.min(...xs) - (Math.max(...xs) - Math.min(...xs)) * .06, Math.max(...xs) + (Math.max(...xs) - Math.min(...xs)) * .06, 5);
  const yl = Math.min(...ys, 0), yh = Math.max(...ys, 0), ty = niceTicks(yl - (yh - yl) * .08, yh + (yh - yl) * .08, 5);
  const x = lin(tx[0], tx[tx.length - 1], ml, W - mr), y = lin(ty[0], ty[ty.length - 1], H - mb, mt);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Scatter chart' });
  (o.bands || []).forEach(b => { const a = y(Math.min(ty[ty.length - 1], b.y1)), c = y(Math.max(ty[0], b.y0)); svg.append(sv('rect', { x: ml, y: a, width: W - ml - mr, height: Math.max(0, c - a), fill: 'var(--wash)' }), txt(W - mr - 4, a + 12, b.label, { 'text-anchor': 'end', 'font-size': 10.5 })); });
  ty.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, y(t) + 4, o.yFmt(t), { 'text-anchor': 'end' })));
  tx.forEach(t => svg.append(txt(x(t), H - mb + 16, o.xFmt(t), { 'text-anchor': 'middle' })));
  svg.append(txt((ml + W - mr) / 2, H - 4, o.xTitle || '', { 'text-anchor': 'middle', 'font-weight': 600 }));
  svg.append(txt(ml, mt - 1, o.yTitle || '', { 'font-weight': 600 }));
  const anySel = o.pts.some(p => p.selected);
  const dots = o.pts.map(p => { const c = sv('circle', { cx: x(p.x), cy: y(p.y), r: p.r, fill: p.color, 'fill-opacity': .82, stroke: 'var(--surface)', 'stroke-width': 2, opacity: anySel && !p.selected ? .35 : 1 }); svg.append(c); return c; });
  o.pts.forEach(p => { if (o.labelIds && o.labelIds.has(p.id)) svg.append(txt(x(p.x) + p.r + 4, y(p.y) + 4, p.label, { 'font-size': 10.5, style: 'fill:var(--ink)' })); });
  const rect = sv('rect', { x: ml, y: 0, width: W - ml - mr, height: H, fill: 'transparent', style: o.onClick ? 'cursor:pointer' : '' }); svg.append(rect);
  const near = ev => { const r = svg.getBoundingClientRect(), k = W / r.width, px = (ev.clientX - r.left) * k, py = (ev.clientY - r.top) * k; let best = -1, bd = 1e9; o.pts.forEach((p, i) => { const d = Math.hypot(x(p.x) - px, y(p.y) - py); if (d < bd) { bd = d; best = i; } }); return bd <= 36 ? best : -1; };
  let last = -1;
  rect.addEventListener('pointermove', ev => { const i = near(ev); if (last >= 0 && dots[last]) dots[last].setAttribute('r', o.pts[last].r); last = i; if (i < 0) { hideTip(); return; } dots[i].setAttribute('r', o.pts[i].r + 3); showTip(ev, o.pts[i].label, o.pts[i].tipRows, o.onClick ? 'Click to filter' : ''); });
  rect.addEventListener('pointerleave', () => { if (last >= 0) dots[last].setAttribute('r', o.pts[last].r); last = -1; hideTip(); });
  if (o.onClick) rect.addEventListener('click', ev => { const i = near(ev); if (i >= 0) o.onClick(o.pts[i]); });
  host.replaceChildren(...(o.legend ? [legendEl(o.legend)] : []), svg);
}

/* ---------- 7. Histogram ---------- */
function histogram(host, o) {
  // o: {bins:[{x0,x1,count}], vlines:[{x,label}], xFmt, height}
  const W = chartHost(host), H = o.height || 230, ml = 52, mr = 14, mt = 18, mb = 34, n = o.bins.length, pw = W - ml - mr, bw = pw / n;
  const mx = Math.max(1, ...o.bins.map(b => b.count)), ty = niceTicks(0, mx, 4), y = lin(0, ty[ty.length - 1], H - mb, mt);
  const lo = o.bins[0].x0, hi = o.bins[n - 1].x1, X = v => ml + (v - lo) / (hi - lo) * pw;
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Histogram' });
  ty.forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: 'var(--grid)', 'stroke-width': 1 }), txt(ml - 8, y(t) + 4, numF(t), { 'text-anchor': 'end' })));
  o.bins.forEach((b, i) => {
    const mid = (b.x0 + b.x1) / 2, c = b.band ? 'var(--axis)' : (mid >= 0 ? 'var(--fav)' : 'var(--unf)');
    const r = sv('path', { d: roundedBarPath(ml + i * bw + 1, y(0), y(b.count), Math.max(2, bw - 2), 3, false), fill: c, opacity: b.band ? .9 : .85 });
    const hit = sv('rect', { x: ml + i * bw, y: mt, width: bw, height: H - mb - mt, fill: 'transparent' });
    hit.addEventListener('pointermove', ev => showTip(ev, o.xFmt(b.x0) + ' to ' + o.xFmt(b.x1), [{ label: 'Lines', value: numF(b.count) }, { label: 'Share of lines', value: pctF(b.count / o.total * 100, 1) }]));
    hit.addEventListener('pointerleave', hideTip); svg.append(r, hit);
  });
  for (let t = Math.ceil(lo / 10) * 10; t <= hi; t += 10) svg.append(txt(X(t), H - mb + 15, o.xFmt(t), { 'text-anchor': 'middle' }));
  (o.vlines || []).forEach(v => svg.append(sv('line', { x1: X(v.x), x2: X(v.x), y1: mt - 4, y2: H - mb, stroke: 'var(--ink2)', 'stroke-width': 1, opacity: .6 }), txt(X(v.x), mt - 6, v.label, { 'text-anchor': 'middle', 'font-size': 10 })));
  host.replaceChildren(svg);
}

/* ---------- 8. Dumbbell (two periods) ---------- */
function dumbbell(host, o) {
  // o: {items:[{id,label,v0,v1,selected}], l0, l1, fmt, onClick}
  const W = chartHost(host), rowH = 26, mt = 22, H = mt + o.items.length * rowH + 22, labW = Math.min(190, Math.max(100, Math.max(...o.items.map(i => i.label.length)) * 6.3));
  const all = o.items.flatMap(i => [i.v0, i.v1]).filter(isFinite), lo = Math.min(...all), hi = Math.max(...all), t = niceTicks(lo - (hi - lo) * .05, hi + (hi - lo) * .05, 5);
  const x = lin(t[0], t[t.length - 1], labW + 12, W - 90);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Change between two periods' });
  t.forEach(v => svg.append(sv('line', { x1: x(v), x2: x(v), y1: mt - 6, y2: H - 18, stroke: v === 100 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }), txt(x(v), H - 4, o.fmt(v), { 'text-anchor': 'middle', 'font-size': 10.5 })));
  svg.append(txt(labW + 12, 10, 'o ' + o.l0 + '   ● ' + o.l1, { 'font-size': 11 }));
  o.items.forEach((it, k) => {
    const cy = mt + k * rowH + rowH / 2, d = it.v1 - it.v0, c = d >= 0 ? 'var(--fav)' : 'var(--unf)';
    const g = sv('g', { class: o.onClick ? 'clk' : '', opacity: o.anySel && !it.selected ? .4 : 1 });
    g.append(sv('rect', { x: 0, y: cy - rowH / 2, width: W, height: rowH, fill: 'transparent' }), txt(labW, cy + 4, it.label, { 'text-anchor': 'end', 'font-size': 12, style: 'fill:var(--ink)' }),
      sv('line', { x1: x(it.v0), x2: x(it.v1), y1: cy, y2: cy, stroke: c, 'stroke-width': 2.5, 'stroke-linecap': 'round' }),
      sv('circle', { cx: x(it.v0), cy, r: 4.5, fill: 'var(--surface)', stroke: 'var(--muted)', 'stroke-width': 2 }), sv('circle', { cx: x(it.v1), cy, r: 4.5, fill: c, stroke: 'var(--surface)', 'stroke-width': 2 }),
      txt(Math.max(x(it.v0), x(it.v1)) + 10, cy + 4, ppF(d), { 'font-size': 11, 'font-weight': 600, style: 'fill:var(--ink)' }));
    g.addEventListener('pointermove', ev => showTip(ev, it.label, [{ label: o.l0, value: o.fmt(it.v0) }, { label: o.l1, value: o.fmt(it.v1) }, { label: 'Change', value: ppF(d) }], o.onClick ? 'Click to filter' : ''));
    g.addEventListener('pointerleave', hideTip); if (o.onClick) g.addEventListener('click', () => o.onClick(it));
    focusable(g, it.label + ': ' + o.fmt(it.v0) + ' to ' + o.fmt(it.v1), o.onClick ? () => o.onClick(it) : null, () => ({ title: it.label, rows: [{ label: o.l0, value: o.fmt(it.v0) }, { label: o.l1, value: o.fmt(it.v1) }, { label: 'Change', value: ppF(d) }] }));
    svg.append(g);
  });
  host.replaceChildren(svg);
}

/* ---------- 9. Single stacked bar (parts of 100 %) ---------- */
function stack100(host, segs, o = {}) {
  // segs: [{label,value,color,txt}]
  const W = chartHost(host), H = 44, total = segs.reduce((a, s) => a + s.value, 0) || 1, svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': o.aria || 'Share breakdown' });
  let x = 0;
  segs.forEach(s => {
    const w = s.value / total * W; if (w <= 0) return;
    const r = sv('rect', { x: x + 1, y: 6, width: Math.max(0, w - 2), height: 30, rx: 4, fill: s.color });
    r.addEventListener('pointermove', ev => showTip(ev, s.label, [{ label: 'Lines', value: numF(s.value) }, { label: 'Share', value: pctF(s.value / total * 100, 1) }].concat(s.extra || [])));
    r.addEventListener('pointerleave', hideTip); svg.append(r);
    focusable(r, s.label + ': ' + numF(s.value) + ' lines', null, () => ({ title: s.label, rows: [{ label: 'Lines', value: numF(s.value) }, { label: 'Share', value: pctF(s.value / total * 100, 1) }].concat(s.extra || []) }));
    const t = s.txt || pctF(s.value / total * 100, 0); if (w > t.length * 7 + 16) svg.append(txt(x + w / 2, 26, t, { 'text-anchor': 'middle', 'font-size': 12, 'font-weight': 650, style: `fill:${/mid|target|wash/.test(s.color) ? 'var(--ink)' : '#fff'}` }));
    x += w;
  });
  host.replaceChildren(svg, legendEl(segs.map(s => ({ label: `${s.label} · ${numF(s.value)}`, color: s.color, type: 'bar' }))));
}

/* ---------- 10. Pareto (bars + cumulative line, one common % axis) ---------- */
function pareto(host, o) {
  // o: {items:[{label,share,cum,selected}], onClick}
  const W = chartHost(host), H = 300, ml = 50, mr = 14, mt = 14, mb = 96, n = o.items.length, pw = W - ml - mr, band = pw / n, X = i => ml + band * (i + .5), y = lin(0, 100, H - mb, mt);
  const svg = sv('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': 'Pareto of shortfall by product' });
  [0, 20, 40, 60, 80, 100].forEach(t => svg.append(sv('line', { x1: ml, x2: W - mr, y1: y(t), y2: y(t), stroke: t === 80 ? 'var(--ink2)' : 'var(--grid)', 'stroke-width': 1, opacity: t === 80 ? .6 : 1 }), txt(ml - 8, y(t) + 4, t + '%', { 'text-anchor': 'end' })));
  svg.append(txt(ml + 4, y(80) - 4, '80% of shortfall', { 'font-size': 10.5 }));
  const bw = Math.min(24, band * .62);
  o.items.forEach((it, i) => {
    const g = sv('g', { class: o.onClick && it.id != null ? 'clk' : '', opacity: o.anySel && !it.selected ? .4 : 1 });
    g.append(sv('rect', { x: X(i) - band / 2, y: mt, width: band, height: H - mb - mt, fill: 'transparent' }));
    g.append(sv('path', { d: roundedBarPath(X(i) - bw / 2, y(0), y(it.share), bw, 4, false), fill: 'var(--unf)' }));
    const lx = X(i), ly = H - mb + 10; g.append(sv('text', { x: lx, y: ly, 'font-size': 10.5, class: 'axis-t', transform: `rotate(-45 ${lx} ${ly})`, 'text-anchor': 'end', text: it.label }));
    g.addEventListener('pointermove', ev => showTip(ev, it.label, [{ label: 'Shortfall', value: usd(it.usd) }, { label: 'Share of total shortfall', value: pctF(it.share, 1) }, { label: 'Cumulative', value: pctF(it.cum, 1) }], o.onClick && it.id != null ? 'Click to filter' : ''));
    g.addEventListener('pointerleave', hideTip); if (o.onClick && it.id != null) g.addEventListener('click', () => o.onClick(it));
    focusable(g, it.label + ': ' + pctF(it.share, 1) + ' of shortfall', o.onClick && it.id != null ? () => o.onClick(it) : null, () => ({ title: it.label, rows: [{ label: 'Shortfall', value: usd(it.usd) }, { label: 'Share of total shortfall', value: pctF(it.share, 1) }, { label: 'Cumulative', value: pctF(it.cum, 1) }] }));
    svg.append(g);
  });
  const d = o.items.map((it, i) => (i ? 'L' : 'M') + X(i) + ',' + y(it.cum)).join('');
  svg.append(sv('path', { d, fill: 'none', stroke: 'var(--ink)', 'stroke-width': 2, 'stroke-linejoin': 'round' }));
  o.items.forEach((it, i) => svg.append(sv('circle', { cx: X(i), cy: y(it.cum), r: 3.5, fill: 'var(--ink)', stroke: 'var(--surface)', 'stroke-width': 2 })));
  host.replaceChildren(legendEl([{ label: 'Share of total shortfall (bars)', color: 'var(--unf)', type: 'bar' }, { label: 'Cumulative share (line)', color: 'var(--ink)' }]), svg);
}

/* ---------- 11. Sparkline ---------- */
function spark(vals, w = 120, hgt = 32, signed = true) {
  const v = vals.filter(isFinite); if (v.length < 2) return h('span');
  const lo = Math.min(...v), hi = Math.max(...v), x = lin(0, vals.length - 1, 2, w - 6), y = lin(lo, hi === lo ? hi + 1 : hi, hgt - 4, 4);
  let d = ''; vals.forEach((q, i) => { if (isFinite(q)) d += (d ? 'L' : 'M') + x(i) + ',' + y(q); });
  const last = vals.length - 1; const svg = sv('svg', { width: w, height: hgt, viewBox: `0 0 ${w} ${hgt}`, 'aria-hidden': 'true' });
  if (signed && lo < 0 && hi > 0) svg.append(sv('line', { x1: 0, x2: w, y1: y(0), y2: y(0), stroke: 'var(--axis)', 'stroke-width': 1 }));
  svg.append(sv('path', { d, fill: 'none', stroke: 'var(--target)', 'stroke-width': 1.5, 'stroke-linejoin': 'round' }), sv('circle', { cx: x(last), cy: y(vals[last]), r: 3.5, fill: 'var(--s1)', stroke: 'var(--surface)', 'stroke-width': 1.5 }));
  return svg;
}

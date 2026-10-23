/* ===========================================================================
   TAB 4 - REGION x CATEGORY
   =========================================================================== */
const RC = { metric: 'pct', persSort: null };
function persistenceRows(mask) {
  const g = agg(KEY_RCM, NR * NC * NM, mask), rows = [];
  for (let r = 0; r < NR; r++) for (let c = 0; c < NC; c++) {
    let months = 0, below = 0, unf = 0, sumV = 0, worst = Infinity, worstM = -1, tc = 0, ac = 0, nlines = 0;
    for (let m = 0; m < NM; m++) { const k = (r * NC + c) * NM + m; if (!g.n[k]) continue; const vc = g.as[k] - g.ts[k]; months++; nlines += g.n[k]; if (vc < 0) below++; if (statusOf(vc, g.ts[k]) === ST_UNF) unf++; sumV += vc; tc += g.ts[k]; ac += g.as[k]; if (vc < worst) { worst = vc; worstM = m; } }
    if (!months) continue;
    rows.push({ r, c, region: REG[r], category: CAT[c], months, below, share: below / months * 100, chronic: below >= PERSIST * months, unf, avg: sumV / 100 / months, worstV: worst / 100, worstM: MLONG[worstM], total: sumV / 100, pct: pct2(sumV, tc), n: nlines });
  }
  return rows;
}
const PERSIST = CFG.persist;
function tabRegCat(root) {
  const mRC = getMask('reg,cat');
  if (!maskCount(mRC)) { root.append(emptyMsg()); return; }
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const g = agg(KEY_RC, NR * NC, mRC), pers = persistenceRows(mRC), pmap = {}; pers.forEach(p => pmap[p.r * NC + p.c] = p);
  const tot = totals(mRC);
  const M = {
    pct: { l: 'Sales variance %', f: d => d.pct, fmt: v => pctF(v, 1, true), lf: v => pctF(v, 1), s: 'diverging' },
    v: { l: 'Sales variance $', f: d => d.v, fmt: v => usdC(v, 2), fmtShort: v => usdC(v, 1), lf: v => usdC(v, 1), s: 'diverging' },
    gpct: { l: 'Gross-profit variance %', f: d => d.gpct, fmt: v => pctF(v, 1, true), lf: v => pctF(v, 1), s: 'diverging' },
    real: { l: 'Selling price vs target %', f: d => d.real, fmt: v => pctF(v, 1, true), lf: v => pctF(v, 1), s: 'diverging' },
    gmc: { l: 'Gross margin change (pp)', f: d => d.gmc, fmt: v => ppF(v, 1), lf: v => v.toFixed(1), s: 'diverging' },
    att: { l: 'Qty attainment %', f: d => d.att - 100, fmt: v => pctF(v + 100, 1), lf: v => pctF(v + 100, 0), s: 'diverging', show: v => v + 100 },
    below: { l: '% of months below target', f: (d, k) => (pmap[k] ? pmap[k].share : NaN), fmt: v => v.toFixed(0) + '%', lf: v => v.toFixed(0) + '%', s: 'seq' },
    contrib: { l: 'Contribution to total sales variance (pp)', f: d => (tot.tc > 0 ? d.vc / tot.tc * 100 : NaN), fmt: v => (v > 0 ? '+' : '') + v.toFixed(2), lf: v => v.toFixed(2), s: 'diverging' }
  };
  const mm = M[RC.metric];
  const vals = REG.map((_, r) => CAT.map((__, c) => { const k = r * NC + c; return g.n[k] ? mm.f(dv(g, k), k) : NaN; }));
  const selR = new Set(S.regs), selC = new Set(S.cats);
  grid.append(card({
    id: 'rc-heat', title: 'Region × category heat map', sub: 'Blue above target, red below, grey near zero · click a cell (or Tab + Enter) to filter region and category', span: 's7',
    tools: [h('select', { class: 'sel', 'aria-label': 'Heat map measure', onchange: e => { RC.metric = e.target.value; renderView(); } }, Object.entries(M).map(([k, v]) => h('option', { value: k, selected: k === RC.metric }, v.l)))],
    build: b => heatmap(b, { rows: REG, cols: CAT, vals, fmt: mm.fmt, labW: 120, cellH: 52, diverging: mm.s === 'diverging', legendFmt: mm.lf, selR, selC, aria: 'Region by category heat map of ' + mm.l,
      onCell: (r, c) => change(() => { if (S.regs.size === 1 && S.regs.has(r) && S.cats.size === 1 && S.cats.has(c)) { S.regs.clear(); S.cats.clear(); } else { S.regs.clear(); S.cats.clear(); S.regs.add(r); S.cats.add(c); } }),
      cellTip: (r, c) => { const k = r * NC + c, p = pmap[k]; if (!g.n[k]) return { title: REG[r] + ' · ' + CAT[c], rows: [{ label: 'Lines', value: '0' }] }; return { title: REG[r] + ' · ' + CAT[c], rows: tipFor(dv(g, k), '').rows.concat(p ? [{ label: 'Months below target', value: p.below + ' of ' + p.months }] : []) }; } }),
    table: () => ({ name: 'region_category', cols: [{ k: 'region', label: 'Region' }, { k: 'category', label: 'Category' }, { k: 'v', label: 'Sales var $', align: 'r', fmt: v => usd(v) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => pctF(v, 2, true) }, { k: 'real', label: 'Price vs target', align: 'r', fmt: v => pctF(v, 2, true) }, { k: 'gv', label: 'GP var $', align: 'r', fmt: v => usd(v) }, { k: 'gpct', label: 'GP var %', align: 'r', fmt: v => pctF(v, 2, true) }, { k: 'att', label: 'Qty att. %', align: 'r', fmt: v => pctF(v) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }], rows: REG.flatMap((rn, r) => CAT.map((cn, c) => { const k = r * NC + c; return g.n[k] ? Object.assign({ region: rn, category: cn }, dv(g, k)) : null; }).filter(Boolean)) })
  }));
  grid.append(card({ id: 'rc-chronic', title: 'Persistent shortfalls', sub: `A pair is “persistent” if its sales are below target in at least ${PERSIST * 100}% of the months in view`, span: 's5', build: b => {
    const ch = pers.filter(p => p.chronic).length;
    b.append(h('div', { style: { display: 'flex', alignItems: 'baseline', gap: '10px', margin: '2px 0 10px' } }, h('span', { style: { fontSize: '40px', fontWeight: 650 } }, `${ch}`), h('span', { class: 'sub' }, `of ${pers.length} region × category pairs are persistent shortfalls`)));
    const top = pers.slice().sort((a, c) => c.share - a.share || a.total - c.total).slice(0, 7);
    b.append(h('ul', { class: 'ins' }, top.map(p => h('li', { class: p.chronic ? 'bad' : '' }, `${p.region} · ${p.category}: below target in ${p.below} of ${p.months} months (${p.share.toFixed(0)}%), total ${usdC(p.total)}`))));
    b.append(h('div', { class: 'note' }, 'A gap that repeats month after month points to something structural (the target, the market, pricing or distribution) rather than a one-off event. Which of these it is cannot be told from this data alone. Full ranking below.'));
  } }));
  // region x month heat strip
  const mRM = getMask('time,reg'), gRM = agg(KEY_RM, NR * NM, mRM), inSel = monthOk();
  const valsM = REG.map((_, r) => MON.map((__, m) => { const k = r * NM + m; return gRM.n[k] ? pct2(gRM.as[k] - gRM.ts[k], gRM.ts[k]) : NaN; }));
  const selCm = new Set(); inSel.forEach((v, i) => { if (v && (S.m0 !== 0 || S.m1 !== NM - 1 || S.years.size || S.qtrs.size)) selCm.add(i); });
  grid.append(card({ id: 'rc-rm', title: 'Region × month: sales variance %', sub: 'Every cell is a region’s monthly total vs its target · click a cell to pick the region and month', span: 's12',
    build: b => heatmap(b, { rows: REG, cols: MON.map((_, i) => MNAME[MNUM[i] - 1]), colSub: MON.map((_, i) => (MNUM[i] === 1 ? String(MYEAR[i]) : '')), vals: valsM, fmt: v => v.toFixed(1), fmtShort: v => v.toFixed(0), labW: 120, cellH: 34, legendFmt: v => pctF(v, 1), selR: new Set(S.regs), selC: selCm, aria: 'Region by month variance heat map',
      onCell: (r, m) => change(() => { S.regs.clear(); S.regs.add(r); S.years.clear(); S.qtrs.clear(); S.m0 = S.m1 = m; }),
      cellTip: (r, m) => { const k = r * NM + m; return { title: REG[r] + ' · ' + MLONG[m], rows: gRM.n[k] ? tipFor(dv(gRM, k), '').rows : [{ label: 'Lines', value: '0' }] }; } }),
    table: () => ({ name: 'region_month', cols: [{ k: 'region', label: 'Region' }].concat(MON.map((m, i) => ({ k: 'm' + i, label: MLAB[i], align: 'r', fmt: v => pctF(v, 1, true), csv: r => r['m' + i] }))), rows: REG.map((rn, r) => Object.assign({ region: rn }, ...MON.map((m, i) => ({ ['m' + i]: valsM[r][i] })))), search: false }) }));
  // dumbbell YoY attainment
  if (YEARS.length >= 2) {
    const yLo = YEARS[YEARS.length - 2], yHi = YEARS[YEARS.length - 1];
    const mT = getMask('time,reg,cat'); const KY = Float64Array.from({ length: N }, (_, i) => KEY_RC[i] * 2 + (MYEAR[A.M[i]] === yHi ? 1 : 0));
    const mY = new Uint8Array(N); for (let i = 0; i < N; i++) mY[i] = mT[i] && (MYEAR[A.M[i]] === yLo || MYEAR[A.M[i]] === yHi) ? 1 : 0;
    const gy = agg(KY, NR * NC * 2, mY);
    const items = []; for (let r = 0; r < NR; r++) for (let c = 0; c < NC; c++) { const k0 = (r * NC + c) * 2, k1 = k0 + 1; if (!gy.ts[k0] || !gy.ts[k1]) continue; items.push({ id: r * NC + c, r, c, label: REG[r] + ' · ' + CAT[c], v0: pct2(gy.as[k0], gy.ts[k0]), v1: pct2(gy.as[k1], gy.ts[k1]), selected: (S.regs.size === 0 || S.regs.has(r)) && (S.cats.size === 0 || S.cats.has(c)) && (S.regs.size || S.cats.size) ? true : false }); }
    items.sort((a, b) => (a.v1 - a.v0) - (b.v1 - b.v0));
    grid.append(card({ id: 'rc-db', title: `Is it getting better? Sales attainment ${yLo} → ${yHi}`, sub: 'Actual sales ÷ target sales per region × category, sorted by change · red = deteriorated, blue = improved · click to filter', span: 's6',
      build: b => dumbbell(b, { items, l0: String(yLo), l1: String(yHi), fmt: v => v.toFixed(1) + '%', anySel: items.some(i => i.selected), onClick: it => change(() => { S.regs.clear(); S.cats.clear(); S.regs.add(it.r); S.cats.add(it.c); }), aria: 'Attainment change by region and category' }),
      table: () => ({ name: 'attainment_change', cols: [{ k: 'label', label: 'Region · Category' }, { k: 'v0', label: String(yLo), align: 'r', fmt: v => pctF(v) }, { k: 'v1', label: String(yHi), align: 'r', fmt: v => pctF(v) }, { k: 'd', label: 'Change (pp)', align: 'r', fmt: v => ppF(v), sortVal: r => r.v1 - r.v0, csv: r => (r.v1 - r.v0).toFixed(2) }], rows: items, search: false }) }));
  }
  grid.append(card({ id: 'rc-pers', title: 'Persistence ranking: all region × category pairs', sub: 'Months below target, worst month and total variance · click a row to filter', span: YEARS.length >= 2 ? 's6' : 's12', build: b => b.append(tableView({
    name: 'persistence', rows: pers, sortKey: 'share', sortDir: -1, search: false, onRow: r => change(() => { S.regs.clear(); S.cats.clear(); S.regs.add(r.r); S.cats.add(r.c); }),
    cols: [{ k: 'region', label: 'Region' }, { k: 'category', label: 'Category' }, { k: 'below', label: 'Months below', align: 'r', fmt: (v, r) => v + ' / ' + r.months, sortVal: r => r.share }, { k: 'share', label: 'Share', align: 'r', fmt: v => v.toFixed(0) + '%' }, { k: 'chronic', label: 'Persistent', fmt: v => (v ? h('span', { class: 'chr' }, '● Persistent') : h('span', { style: { color: 'var(--ink2)' } }, 'No')), csv: r => (r.chronic ? 'Yes' : 'No') },
      { k: 'avg', label: 'Avg monthly var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'worstV', label: 'Worst month $', align: 'r', fmt: (v, r) => usd(v) + ' (' + r.worstM + ')' }, { k: 'total', label: 'Total var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }] })) }));
}

/* ===========================================================================
   TAB 5 - PRODUCTS
   =========================================================================== */
const PR = { view: 'worst', n: 10, by: 'v' };
function tabProducts(root) {
  const mP = getMask('prod');
  if (!maskCount(mP)) { root.append(emptyMsg()); return; }
  const rows = groupRows(A.P, NP, mP, PRD).filter(d => d.n).map(d => Object.assign(d, { cat: CAT[DIM.prodCat[d.id]], ci: DIM.prodCat[d.id] }));
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const key = PR.by === 'v' ? 'v' : 'pct';
  let shown = rows.slice().sort((a, b) => a[key] - b[key]);
  if (PR.view === 'worst') shown = shown.slice(0, PR.n); else if (PR.view === 'best') shown = shown.slice(-PR.n).reverse(); 
  if (PR.view === 'best' || PR.view === 'all') shown.sort((a, b) => b[key] - a[key]); else shown.sort((a, b) => a[key] - b[key]);
  const above = rows.filter(d => d.v > 0).length;
  const items = shown.map(d => ({ id: d.id, label: d.name, value: d.v, txt: usdS(d.v) + ' · ' + pctF(d.pct, 2, true), selected: S.prods.has(d.id), d, color: undefined }));
  const ctl = h('div', { class: 'tbar', style: { marginBottom: '4px' } }, seg([['worst', 'Biggest shortfalls'], ['best', 'Best performers'], ['all', 'All']], PR.view, v => { PR.view = v; renderView(); }, 'Product ranking view'),
    h('select', { class: 'sel', 'aria-label': 'How many products', onchange: e => { PR.n = +e.target.value; renderView(); } }, [5, 10, 15, 27].map(n => h('option', { value: n, selected: n === PR.n }, 'Top ' + n))),
    seg([['v', 'by $'], ['pct', 'by %']], PR.by, v => { PR.by = v; renderView(); }, 'Rank by'));
  grid.append(card({ id: 'pr-rank', title: 'Products ranked by sales variance', sub: `${above} of ${rows.length} products are above target. A ranking is not a verdict: a “best” product can still miss target.`, span: 's7', build: b => {
    b.append(ctl); const host = h('div'); b.append(host);
    hbars(host, items, { onClick: (it, ev) => pickProd(it.id, ev), rowH: 28, tip: it => { const t = tipFor(it.d, it.label + ' (' + it.d.cat + ')'); t.rows.push({ label: 'Above target', value: it.d.v > 0 ? 'Yes' : 'No' }); return t; }, aria: 'Products ranked by sales variance' });
  }, table: () => ({ name: 'products', cols: prodCols(), rows, sortKey: 'v', sortDir: 1 }) }));
  // pareto of shortfall
  const neg = rows.filter(d => d.v < 0).sort((a, b) => a.v - b.v), gross = neg.reduce((s, d) => s + d.v, 0);
  const keep = neg.slice(0, 15); let cum = 0; const pi = keep.map(d => { cum += d.v; return { id: d.id, label: d.name, usd: d.v, share: d.v / gross * 100, cum: cum / gross * 100, selected: S.prods.has(d.id) }; });
  if (neg.length > 15) { const rest = neg.slice(15), rv = rest.reduce((s, d) => s + d.v, 0); cum += rv; pi.push({ id: null, label: `Other (${rest.length})`, usd: rv, share: rv / gross * 100, cum: cum / gross * 100 }); }
  grid.append(card({ id: 'pr-pareto', title: 'Where the shortfall sits (Pareto)', sub: neg.length ? `${neg.length} products below target add up to ${usdC(gross)}; ${pi.filter(p => p.cum - p.share < 80 && p.id != null).length} reach 80%` : 'No product is below target in this selection', span: 's5',
    build: b => (neg.length ? pareto(b, { items: pi, anySel: pi.some(p => p.selected), onClick: it => pickProd(it.id) }) : b.append(emptyMsg('Nothing below target.'))), table: () => ({ name: 'pareto', cols: [{ k: 'label', label: 'Product' }, { k: 'usd', label: 'Shortfall $', align: 'r', fmt: v => usd(v) }, { k: 'share', label: 'Share %', align: 'r', fmt: v => pctF(v, 1) }, { k: 'cum', label: 'Cumulative %', align: 'r', fmt: v => pctF(v, 1) }], rows: pi, search: false }) }));
  // bubble
  const maxAbs = Math.max(1, ...rows.map(d => Math.abs(d.v))), pts = rows.map(d => ({ id: d.id, label: d.name, x: d.target, y: d.pct, r: 5 + 11 * Math.sqrt(Math.abs(d.v) / maxAbs), color: catColor(d.ci), selected: S.prods.has(d.id), tipRows: tipFor(d, '').rows.concat([{ label: 'Category', value: d.cat }]) }));
  const lab = new Set(rows.slice().sort((a, b) => a.pct - b.pct).slice(0, 3).concat(rows.slice().sort((a, b) => b.pct - a.pct).slice(0, 3)).map(d => d.id));
  grid.append(card({ id: 'pr-bubble', title: 'Size vs performance', sub: 'Each bubble is a product: planned sales (x), sales variance % (y), bubble size = dollar variance · colour = category · click to filter (the table view lists every product)', span: 's7',
    build: b => bubble(b, { pts, xFmt: v => usdC(v, 0), yFmt: v => pctF(v, 1), xTitle: 'Target sales', yTitle: 'Sales variance %', labelIds: lab, onClick: p => pickProd(p.id), height: 340, bands: [{ y0: -CFG.unfPct, y1: CFG.favPct, label: WITHIN_LABEL + ' band' }], legend: CAT.map((c, i) => ({ label: c, color: catColor(i), type: 'bar' })), aria: 'Products by size and variance' }),
    table: () => ({ name: 'products_bubble', cols: prodCols(), rows, sortKey: 'pct', sortDir: 1 }) }));
  // by category within product tab: share of lines? product table
  grid.append(card({ id: 'pr-tbl', title: 'Product scorecard', sub: 'Click a row to filter · scroll sideways for more columns', span: 's5', build: b => b.append(tableView({ name: 'product_scorecard', rows, sortKey: 'v', sortDir: 1, onRow: r => pickProd(r.id), cols: prodCols() })) }));
}
function prodCols() {
  return [{ k: 'name', label: 'Product' }, { k: 'cat', label: 'Category' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Sales var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'real', label: 'Price vs target', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) },
    { k: 'gv', label: 'GP var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'gpct', label: 'GP var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'att', label: 'Qty att. %', align: 'r', fmt: v => pctF(v) },
    { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }, { k: 'mat', label: 'Material', fmt: (v, r) => matCell(v, r.trig), csv: r => (r.mat ? 'Yes' : 'No') }, { k: 'pb', label: '% lines below target', align: 'r', fmt: v => pctF(v, 1) }];
}

/* ===========================================================================
   TAB 6 - EXCEPTIONS & LINES
   =========================================================================== */
const EX = { lineQ: '' };
function tabExceptions(root) {
  const m = getMask(), T = totals(m);
  if (!T.n) { root.append(emptyMsg()); return; }
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const g = agg(KEY_RCM, NR * NC * NM, m), cells = [];
  for (let r = 0; r < NR; r++) for (let c = 0; c < NC; c++) for (let mo = 0; mo < NM; mo++) { const k = (r * NC + c) * NM + mo; if (!g.n[k]) continue; const d = dv(g, k); d.r = r; d.c = c; d.mo = mo; d.region = REG[r]; d.category = CAT[c]; d.month = MLONG[mo]; cells.push(d); }
  const matCells = cells.filter(d => d.mat).sort((a, b) => Math.abs(b.v) - Math.abs(a.v));
  const fav = matCells.filter(d => d.v > 0).length;
  grid.append(h('section', { class: 'card s4 kpi' }, h('div', { class: 'lab' }, 'Material region × category × month cells'), h('div', { class: 'val' }, `${matCells.length} of ${cells.length}`), h('div', { class: 'dl' }, `${matCells.length - fav} below target, ${fav} above · rule: |sales var %| ≥ ${CFG.matPctV}% or |sales var $| ≥ ${usd(CFG.matUsd)} · largest first`)));
  grid.append(h('section', { class: 'card s4 kpi' }, h('div', { class: 'lab' }, 'Material individual lines'), h('div', { class: 'val' }, numF(T.nm)), h('div', { class: 'dl' }, `${pctF(T.nm / T.n * 100, 1)} of ${numF(T.n)} lines – too many to review one by one`)));
  grid.append(h('section', { class: 'card s4 kpi' }, h('div', { class: 'lab' }, 'Typical swing of a single line'), h('div', { class: 'val' }, '±' + lineSd(m).toFixed(1) + '%'), h('div', { class: 'dl' }, 'standard deviation of line sales variance %: why line-level flags are noisy')));
  grid.append(card({ id: 'ex-cells', title: 'Exceptions: material region × category × month cells', sub: 'Ranked by dollar size: start at the top · click a row to filter to it', span: 's12', build: b => b.append(tableView({ name: 'material_cells', rows: matCells, sortKey: null, pageSize: 15, onRow: r => change(() => { S.regs.clear(); S.cats.clear(); S.regs.add(r.r); S.cats.add(r.c); S.years.clear(); S.qtrs.clear(); S.m0 = S.m1 = r.mo; }),
    cols: [{ k: 'month', label: 'Month', sortVal: r => r.mo }, { k: 'region', label: 'Region' }, { k: 'category', label: 'Category' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Sales var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'prc', label: 'Price $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'gv', label: 'GP var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'att', label: 'Qty att. %', align: 'r', fmt: v => pctF(v) },
      { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }, { k: 'trig', label: 'Trigger' }, { k: 'n', label: 'Lines', align: 'r', fmt: v => numF(v) }] })) }));
  // histogram of line variance %
  const bins = []; const lo = -60, hi = 60, bw = 2.5; for (let x = lo; x < hi; x += bw) bins.push({ x0: x, x1: x + bw, count: 0, band: x >= -CFG.unfPct - 1e-9 && x + bw <= CFG.favPct + 1e-9 });
  let tot = 0; for (let i = 0; i < N; i++) { if (!m[i] || !(A.TS[i] > 0)) continue; const p = A.VAR[i] / A.TS[i] * 100; let k = Math.floor((p - lo) / bw); k = Math.max(0, Math.min(bins.length - 1, k)); bins[k].count++; tot++; }
  grid.append(card({ id: 'ex-hist', title: 'How individual lines are spread', sub: `Sales variance % of each line (lines with a target) · grey bars = ${WITHIN_LABEL}`, span: 's12', build: b => { histogram(b, { bins, total: tot, xFmt: v => (v > 0 ? '+' : v < 0 ? MINUS : '') + Math.abs(v) + '%', vlines: [{ x: 0, label: '0' }, { x: -CFG.unfPct, label: MINUS + CFG.unfPct + '%' }, { x: CFG.favPct, label: '+' + CFG.favPct + '%' }, { x: -CFG.matPctV, label: MINUS + CFG.matPctV + '%' }, { x: CFG.matPctV, label: '+' + CFG.matPctV + '%' }], aria: 'Distribution of line variance percent' }); b.append(h('div', { class: 'note' }, 'Single lines swing widely around a small average shift, so conclusions should come from totals and groups, not from individual lines.')); } }));
  // line explorer
  const lineRows = []; for (let i = 0; i < N; i++) { if (!m[i]) continue; lineRows.push({ i, id: IDS[i], month: MLONG[A.M[i]], mo: A.M[i], region: REG[A.R[i]], category: CAT[A.C[i]], product: PRD[A.P[i]], tq: A.TQ[i], aq: A.AQ[i], tp: A.TP[i] / 100, ap: A.AP[i] / 100, tcu: A.TCU[i] / 100, acu: A.ACU[i] / 100, target: A.TS[i] / 100, actual: A.AS[i] / 100, v: A.VAR[i] / 100, pct: pct2(A.VAR[i], A.TS[i]), st: A.ST[i], mat: A.MT[i] === 1, trig: TRIG_NAME[A.TRG[i]], vol: A.VOL[i] / 100, prc: A.PRC[i] / 100, gv: A.GVAR[i] / 100, cost: A.COST[i] / 100, note: NOTES[A.NT[i]], rep: A.RP[i] === 1, dq: A.DQ[i] === 1 }); }
  grid.append(card({ id: 'ex-lines', title: 'Line explorer', sub: `${numF(lineRows.length)} sales lines match your filters · sort, search and export`, span: 's12', build: b => b.append(tableView({ name: 'sales_lines_filtered', rows: lineRows, pageSize: 25, sortKey: 'id', sortDir: 1,
    cols: [{ k: 'id', label: 'Line ID' }, { k: 'month', label: 'Month', sortVal: r => r.mo }, { k: 'region', label: 'Region' }, { k: 'category', label: 'Category' }, { k: 'product', label: 'Product' }, { k: 'tq', label: 'Target qty', align: 'r', fmt: v => numF(v) }, { k: 'aq', label: 'Actual qty', align: 'r', fmt: v => numF(v) }, { k: 'tp', label: 'Target price', align: 'r', fmt: v => usd(v, 2) }, { k: 'ap', label: 'Actual price', align: 'r', fmt: v => usd(v, 2) },
      { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v, 2) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v, 2) }, { k: 'v', label: 'Sales var $', align: 'r', fmt: v => vSpan(v, usd(v, 2)) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] },
      { k: 'mat', label: 'Material', fmt: (v, r) => matCell(v, r.trig), csv: r => (r.mat ? 'Yes' : 'No') }, { k: 'vol', label: 'Volume effect $', align: 'r', fmt: v => usd(v, 2) }, { k: 'prc', label: 'Price effect $', align: 'r', fmt: v => usd(v, 2) }, { k: 'tcu', label: 'Target cost', align: 'r', fmt: v => usd(v, 2) }, { k: 'acu', label: 'Actual cost', align: 'r', fmt: v => usd(v, 2) }, { k: 'cost', label: 'Cost effect $', align: 'r', fmt: v => usd(v, 2) }, { k: 'gv', label: 'GP var $', align: 'r', fmt: v => vSpan(v, usd(v, 2)) }, { k: 'note', label: 'Note' },
      { k: 'rep', label: 'Repaired value', fmt: v => (v ? 'Yes' : ''), csv: r => (r.rep ? 'Yes' : 'No') }, { k: 'dq', label: 'Large order flag', fmt: v => (v ? 'Yes' : ''), csv: r => (r.dq ? 'Yes' : 'No') }] })) }));
}

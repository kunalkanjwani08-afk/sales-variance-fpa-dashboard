/* ===========================================================================
   TAB 2 - DRIVERS (revenue and gross-profit bridges, pricing, margin, what-if)
   =========================================================================== */
const WI = { region: null, pct: null };
const BANDS = [                       // price realisation bands (exact integer tests on cents)
  ['Discount of 5% or more', (ap, tp) => (ap - tp) * 100 <= -5 * tp],
  ['Discount of 1–5%', (ap, tp) => (ap - tp) * 100 < -1 * tp],
  ['At target price (within 1%)', (ap, tp) => (ap - tp) * 100 <= 1 * tp],
  ['Above target price (over 1%)', () => true]];
function bandOf(i) { for (let b = 0; b < BANDS.length; b++) if (BANDS[b][1](A.AP[i], A.TP[i])) return b; return BANDS.length - 1; }
const KEY_BAND = Float64Array.from({ length: N }, (_, i) => bandOf(i));

function tabDrivers(root) {
  const m = getMask(), T = totals(m);
  if (!T.n) { root.append(emptyMsg()); return; }
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const rev = [
    { label: 'Target sales', value: T.target, kind: 'total', color: 'var(--target)' },
    { label: 'Volume (pure)', value: T.pure, kind: 'delta', foot: '(Σ actual qty − Σ target qty) × average target price' },
    { label: 'Mix', value: T.mix, kind: 'delta', foot: 'Selling a different mix of products, regions and accounts than planned, valued at target prices' },
    { label: 'Price', value: T.prc, kind: 'delta', foot: 'Actual qty × (actual unit price − target unit price)' },
    { label: 'Actual sales', value: T.actual, kind: 'total', color: 'var(--s1)' }];
  const gp = [
    { label: 'Target gross profit', value: T.tgp, kind: 'total', color: 'var(--target)' },
    { label: 'Volume (pure)', value: T.ppure, kind: 'delta', foot: '(Σ actual qty − Σ target qty) × average target unit margin' },
    { label: 'Mix', value: T.pmix, kind: 'delta', foot: 'Selling a different margin mix than planned, valued at target unit margins' },
    { label: 'Price', value: T.prc, kind: 'delta', foot: 'Actual qty × (actual unit price − target unit price): price changes fall straight to gross profit' },
    { label: 'Cost', value: T.cost, kind: 'delta', foot: '−Actual qty × (actual unit cost − target unit cost)' },
    { label: 'Actual gross profit', value: T.agp, kind: 'total', color: 'var(--s1)' }];
  const chkR = Math.round((T.target + T.pure + T.mix + T.prc - T.actual) * 100) / 100, chkG = Math.round((T.tgp + T.ppure + T.pmix + T.prc + T.cost - T.agp) * 100) / 100;
  grid.append(card({ id: 'dr-pvm', title: 'Revenue bridge: price, volume, mix', sub: 'Why actual sales differ from target (current filters)', span: 's6', build: b => {
    waterfall(b, rev, { aria: 'Revenue price volume mix bridge' });
    b.append(h('div', { class: 'note' }, `Pure volume ${usdS(T.pure)}, mix ${usdS(T.mix)}, price ${usdS(T.prc)}. Bridge check: target + effects − actual = ${usd(chkR, 2)}.`));
  }, table: () => ({ name: 'revenue_bridge', cols: [{ k: 'label', label: 'Step' }, { k: 'value', label: 'USD', align: 'r', fmt: v => usd(v, 2) }], rows: rev, search: false }) }));
  grid.append(card({ id: 'dr-gp', title: 'Gross-profit bridge: volume, mix, price, cost', sub: 'Why actual gross profit differs from target (current filters)', span: 's6', build: b => {
    waterfall(b, gp, { aria: 'Gross profit bridge' });
    b.append(h('div', { class: 'note' }, `Gross margin ${pctF(T.agm)} vs ${pctF(T.tgm)} planned (${ppF(T.gmc)}). Bridge check: target + effects − actual = ${usd(chkG, 2)}.`));
  }, table: () => ({ name: 'gross_profit_bridge', cols: [{ k: 'label', label: 'Step' }, { k: 'value', label: 'USD', align: 'r', fmt: v => usd(v, 2) }], rows: gp, search: false }) }));
  // by region table (exclude region filter so all regions stay visible)
  const rr = groupRows(A.R, NR, getMask('reg'), REG).filter(d => d.n);
  grid.append(card({ id: 'dr-reg', title: 'Drivers by region', sub: 'Click a row to filter', span: 's12', build: b => b.append(tableView({
    name: 'drivers_by_region', search: false, sortKey: 'v', sortDir: 1, onRow: r => pickRegion(r.id), rows: rr,
    cols: [{ k: 'name', label: 'Region' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'pure', label: 'Volume (pure) $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'mix', label: 'Mix $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'prc', label: 'Price $', align: 'r', fmt: v => vSpan(v, usd(v)) },
      { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Sales var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'real', label: 'Price vs target', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) },
      { k: 'cost', label: 'Cost effect $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'gv', label: 'GP var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'gpct', label: 'GP var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'agm', label: 'Gross margin', align: 'r', fmt: v => pctF(v) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }] })) }));
  // pricing: did discounts pay?
  const gb = agg(KEY_BAND, BANDS.length, m), bands = BANDS.map((bd, i) => Object.assign(dv(gb, i), { name: bd[0], id: i })).filter(d => d.n);
  grid.append(card({ id: 'dr-band', title: 'Did discounts pay? Gross-profit variance by price vs target', sub: 'Each line grouped by its actual price vs its target price · bars = gross-profit variance', span: 's7', build: b => {
    hbars(b, bands.map(d => ({ id: d.id, label: d.name, value: d.gv, txt: usdS(d.gv) + ' (' + pctF(d.gpct, 1, true) + ')', d })), { labW: 190, tip: it => ({ title: it.label, rows: [{ label: 'Lines', value: numF(it.d.n) }, { label: 'Qty attainment', value: pctF(it.d.att) }, { label: 'Sales variance', value: usd(it.d.v) + ' (' + pctF(it.d.pct, 2, true) + ')' }, { label: 'Price effect', value: usd(it.d.prc) }, { label: 'Gross-profit variance', value: usd(it.d.gv) + ' (' + pctF(it.d.gpct, 2, true) + ')' }] }), aria: 'Gross profit variance by price band' });
    b.append(tableView({ name: 'price_bands', search: false, sortKey: null, rows: bands, cols: [{ k: 'name', label: 'Price vs target' }, { k: 'n', label: 'Lines', align: 'r', fmt: v => numF(v) }, { k: 'att', label: 'Qty attainment', align: 'r', fmt: v => pctF(v) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'prc', label: 'Price effect $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'gpct', label: 'GP var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }] }));
  } }));
  const cr = groupRows(A.C, NC, getMask('cat'), CAT).filter(d => d.n);
  const cItems = cr.map(d => ({ id: d.id, label: d.name, value: d.cost, txt: usdS(d.cost), selected: S.cats.has(d.id), d })).sort((a, b) => a.value - b.value);
  grid.append(card({ id: 'dr-cost', title: 'Cost effect by category', sub: 'Gross profit lost (red) or gained (blue) because unit costs differed from target · click to filter', span: 's5', build: b => hbars(b, cItems, { onClick: (it, ev) => pickCat(it.id, ev), tip: it => ({ title: it.label, rows: [{ label: 'Cost effect', value: usd(it.d.cost) }, { label: 'Gross margin', value: pctF(it.d.agm) + ' vs ' + pctF(it.d.tgm) + ' target' }, { label: 'GP variance', value: usd(it.d.gv) + ' (' + pctF(it.d.gpct, 2, true) + ')' }] }), aria: 'Cost effect by category' }),
    table: () => ({ name: 'cost_by_category', cols: [{ k: 'name', label: 'Category' }, { k: 'cost', label: 'Cost effect $', align: 'r', fmt: v => usd(v) }, { k: 'tgm', label: 'Target margin', align: 'r', fmt: v => pctF(v) }, { k: 'agm', label: 'Actual margin', align: 'r', fmt: v => pctF(v) }, { k: 'gv', label: 'GP var $', align: 'r', fmt: v => usd(v) }], rows: cr, search: false }) }));
  grid.append(card({ id: 'dr-whatif', title: 'What-if: close the sales gap in one region', sub: 'Move a region’s sales variance % and see the effect on the company (uses all current filters except region)', span: 's12', build: b => whatIf(b) }));
}
function whatIf(b) {
  const rows = groupRows(A.R, NR, getMask('reg'), REG).filter(d => d.n && d.tc > 0); if (!rows.length) { b.append(emptyMsg()); return; }
  if (WI.region == null || !rows.find(r => r.id === WI.region)) { WI.region = rows.reduce((a, c) => (c.pct < a.pct ? c : a)).id; WI.pct = null; }
  const r = rows.find(x => x.id === WI.region), rest = rows.filter(x => x.id !== r.id);
  const restT = rest.reduce((s, x) => s + x.tc, 0), restV = rest.reduce((s, x) => s + x.vc, 0), restPct = restT > 0 ? restV / restT * 100 : 0;
  const tot = rows.reduce((o, x) => ({ tc: o.tc + x.tc, vc: o.vc + x.vc }), { tc: 0, vc: 0 });
  if (WI.pct == null) WI.pct = Math.round(Math.max(r.pct, restPct) * 10) / 10;
  const lo = Math.floor(Math.min(r.pct, -12)), hi = Math.ceil(Math.max(r.pct, 6));
  const out = h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(190px,1fr))', gap: '14px', marginTop: '10px' } });
  const sel = h('select', { class: 'sel', 'aria-label': 'Region to simulate', onchange: e => { WI.region = +e.target.value; WI.pct = null; b.replaceChildren(); whatIf(b); } }, rows.map(x => h('option', { value: x.id }, x.name)));
  sel.value = r.id;
  const val = h('strong', null, pctF(WI.pct, 1, true));
  const sl = h('input', { type: 'range', min: lo, max: hi, step: 0.1, value: WI.pct, 'aria-label': 'Sales variance percent for the region', oninput: e => { WI.pct = +e.target.value; val.textContent = pctF(WI.pct, 1, true); paint(); } });
  function paint() {
    const revGain = (WI.pct - r.pct) / 100 * r.target, gmGain = revGain * (isFinite(r.agm) ? r.agm / 100 : 0), newV = tot.vc / 100 + revGain, newPct = newV / (tot.tc / 100) * 100;
    const box = (l, v, d, c) => h('div', { class: 'card kpi' }, h('div', { class: 'lab' }, l), h('div', { class: 'val ' + (c || '') }, v), h('div', { class: 'dl' }, d));
    out.replaceChildren(box('Revenue recovered', usdC(revGain), `${r.name}: ${pctF(r.pct, 2, true)} → ${pctF(WI.pct, 1, true)}`, sgnCls(revGain)), box('Gross profit recovered', usdC(gmGain), `at the region’s ${pctF(r.agm, 1)} actual gross margin`, sgnCls(gmGain)),
      box('Company sales variance', usdC(newV), `was ${usdC(tot.vc / 100)} · now ${pctF(newPct, 2, true)} of target`, sgnCls(newV)), box('Share of total gap closed', tot.vc < 0 ? pctF(revGain / -(tot.vc / 100) * 100, 0) : '—', `rest of company runs at ${pctF(restPct, 2, true)}`));
  }
  b.append(h('div', { class: 'slider' }, h('label', null, 'Region '), sel, h('span', null, 'Sales variance %'), sl, val, h('button', { class: 'tool', onclick: () => { WI.pct = Math.round(restPct * 10) / 10; sl.value = WI.pct; val.textContent = pctF(WI.pct, 1, true); paint(); } }, 'Match rest of company (' + pctF(restPct, 1, true) + ')'), h('button', { class: 'tool', onclick: () => { WI.pct = 0; sl.value = 0; val.textContent = pctF(0, 1, true); paint(); } }, 'Hit target (0%)')), out,
    h('div', { class: 'note' }, 'An illustration, not a forecast: it assumes the region’s target stays the same and its actual gross-margin rate holds on the extra sales.'));
  paint();
}

/* ===========================================================================
   TAB 3 - TRENDS
   =========================================================================== */
const TRM = {
  v: { label: 'Sales variance $', get: d => d.v, fmt: v => usd(v), tick: v => usdC(v, 1), bar: true, zero: true },
  pct: { label: 'Sales variance %', get: d => d.pct, fmt: v => pctF(v, 2, true), tick: v => pctF(v, 1), bar: true, zero: true, refs: [{ y: CFG.favPct, label: '+' + CFG.favPct + '% Favorable above' }, { y: -CFG.unfPct, label: MINUS + CFG.unfPct + '% Unfavorable below' }] },
  gv: { label: 'Gross-profit variance $', get: d => d.gv, fmt: v => usd(v), tick: v => usdC(v, 1), bar: true, zero: true },
  gpct: { label: 'Gross-profit variance %', get: d => d.gpct, fmt: v => pctF(v, 2, true), tick: v => pctF(v, 1), bar: true, zero: true },
  prc: { label: 'Price effect $', get: d => d.prc, fmt: v => usd(v), tick: v => usdC(v, 1), bar: true, zero: true },
  real: { label: 'Selling price vs target %', get: d => d.real, fmt: v => pctF(v, 2, true), tick: v => pctF(v, 1), zero: true },
  agm: { label: 'Gross margin %', get: d => d.agm, fmt: v => pctF(v), tick: v => v.toFixed(0) + '%' },
  att: { label: 'Qty attainment %', get: d => d.att, fmt: v => pctF(v), tick: v => v.toFixed(0) + '%', refs: [{ y: 100, label: '100% = on plan' }] },
  actual: { label: 'Actual sales $', get: d => d.actual, fmt: v => usd(v), tick: v => usdC(v, 0) },
  target: { label: 'Target sales $', get: d => d.target, fmt: v => usd(v), tick: v => usdC(v, 0) },
  mom: { label: 'MoM sales growth %', get: d => d.mom, fmt: v => pctF(v, 2, true), tick: v => v.toFixed(0) + '%', zero: true },
  yoy: { label: 'YoY sales growth %', get: d => d.yoy, fmt: v => pctF(v, 2, true), tick: v => v.toFixed(0) + '%', zero: true },
  ytd: { label: 'YTD sales variance $', get: d => d.ytdV, fmt: v => usd(v), tick: v => usdC(v, 1), zero: true },
  r3: { label: 'Rolling 3M sales variance $', get: d => d.r3, fmt: v => usd(v), tick: v => usdC(v, 1), zero: true },
  cum: { label: 'Cumulative sales variance $', get: d => d.cum, fmt: v => usd(v), tick: v => usdC(v, 1), zero: true },
  ts: { label: 'Tracking signal', get: d => d.ts, fmt: v => numF(v, 1), tick: v => numF(v, 0), zero: true, refs: [{ y: TS_LIMIT, label: '+' + TS_LIMIT + ' limit' }, { y: -TS_LIMIT, label: MINUS + TS_LIMIT + ' limit' }] }
};
const TR = { metric: 'v', split: 'total', seas: 'actual', phy: null };
function tabTrends(root) {
  const mAll = getMask('time'), inSel = monthOk();
  if (!maskCount(mAll)) { root.append(emptyMsg()); return; }
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const mt = TRM[TR.metric];
  const controls = h('div', { class: 'tbar', style: { marginBottom: '4px' } }, h('span', { class: 'sub' }, 'Measure'), h('select', { class: 'sel', 'aria-label': 'Trend measure', onchange: e => { TR.metric = e.target.value; renderView(); } }, Object.entries(TRM).map(([k, v]) => h('option', { value: k, selected: k === TR.metric }, v.label))),
    h('span', { class: 'sub' }, 'Split by'), seg([['total', 'Total'], ['region', 'Region'], ['category', 'Category']], TR.split, v => { TR.split = v; renderView(); }, 'Split'));
  const splitDefs = TR.split === 'region' ? { key: A.R, K: NR, names: REG, col: regColor } : TR.split === 'category' ? { key: A.C, K: NC, names: CAT, col: catColor } : null;
  let series = [];
  const allSer = monthlySeries(mAll);
  if (!splitDefs) series = [{ id: 'all', name: mt.label, color: 'var(--s1)', values: allSer.map(d => (d.n ? mt.get(d) : NaN)) }];
  else { const g = agg(Float64Array.from({ length: N }, (_, i) => splitDefs.key[i] * NM + A.M[i]), splitDefs.K * NM, mAll); for (let i = 0; i < splitDefs.K; i++) { if (!Array.from({ length: NM }, (_, k) => g.n[i * NM + k]).some(Boolean)) continue; const ser = monthlyFrom(g, i * NM); series.push({ id: i, name: splitDefs.names[i], color: splitDefs.col(i), values: ser.map(d => (d.n ? mt.get(d) : NaN)) }); } }
  const lineMode = splitDefs || !mt.bar;
  grid.append(card({
    id: 'tr-main', title: mt.label + (splitDefs ? ' by ' + TR.split : ' by month'), sub: 'Click a month to filter · hover for all series' + (TR.metric === 'ts' ? ' · limit ±' + TS_LIMIT : ''), span: 's12', build: b => {
      b.append(controls);
      const host = h('div'); b.append(host);
      lineChart(host, { x: MLAB, xFull: MLONG, series, fmt: mt.fmt, tickFmt: mt.tick, zero: !!mt.zero, mode: lineMode ? 'line' : 'bar', refs: mt.refs, height: 320, onX: selectMonth, selX: inSel, aria: mt.label + ' trend',
        extraTip: i => (splitDefs ? [] : [{ label: 'Status', value: ST_GLYPH[allSer[i].st] + ' ' + ST_NAME[allSer[i].st] }]) });
      if (TR.metric === 'ts') b.append(h('div', { class: 'note' }, 'Tracking signal = cumulative variance ÷ average absolute monthly variance. Beyond ±4 suggests persistent bias. If every month so far is below target it simply equals –(number of months), so read it together with the monthly pattern, not as separate proof.'));
    },
    table: () => ({ name: 'trend_' + TR.metric, cols: [{ k: 'month', label: 'Month' }].concat(series.map(s => ({ k: 'v' + s.id, label: s.name, align: 'r', fmt: v => mt.fmt(v), csv: r => r['v' + s.id] }))), rows: MLONG.map((l, i) => Object.assign({ month: l }, ...series.map(s => ({ ['v' + s.id]: s.values[i] })))), search: false })
  }));
  // seasonality
  const sm = { actual: ['Actual sales', d => d.actual, v => usd(v), v => usdC(v, 0)], target: ['Target sales', d => d.target, v => usd(v), v => usdC(v, 0)], pct: ['Sales variance %', d => d.pct, v => pctF(v, 2, true), v => v.toFixed(1) + '%'] }[TR.seas];
  const yS = YEARS.map((y, yi) => ({ id: y, name: String(y), color: `var(--s${yi + 1})`, values: Array.from({ length: 12 }, (_, k) => { const mi = MON.findIndex((q, i) => MYEAR[i] === y && MNUM[i] === k + 1); return mi >= 0 && allSer[mi].n ? sm[1](allSer[mi]) : NaN; }) }));
  grid.append(card({ id: 'tr-seas', title: 'Same month, different year', sub: 'Seasonality check: is the gap seasonal or constant?', span: 's6', tools: [seg([['actual', 'Actual'], ['target', 'Target'], ['pct', 'Var %']], TR.seas, v => { TR.seas = v; renderView(); }, 'Seasonality measure')],
    build: b => lineChart(b, { x: MNAME, series: yS, fmt: sm[2], tickFmt: sm[3], zero: TR.seas === 'pct', height: 280, aria: 'Month of year by year' }),
    table: () => ({ name: 'seasonality', cols: [{ k: 'month', label: 'Month' }].concat(yS.map(s => ({ k: 'y' + s.id, label: s.name, align: 'r', fmt: v => sm[2](v), csv: r => r['y' + s.id] }))), rows: MNAME.map((n, i) => Object.assign({ month: n }, ...yS.map(s => ({ ['y' + s.id]: s.values[i] })))), search: false }) }));
  // phasing: monthly share of the year
  if (TR.phy == null || !YEARS.includes(TR.phy)) TR.phy = YEARS[YEARS.length - 1];
  const yi = MON.map((_, i) => i).filter(i => MYEAR[i] === TR.phy), yt = yi.reduce((s, i) => s + allSer[i].tc, 0), ya = yi.reduce((s, i) => s + allSer[i].ac, 0);
  const ph = [{ id: 't', name: 'Target share', color: 'var(--target)', values: Array.from({ length: 12 }, (_, k) => { const i = yi.find(j => MNUM[j] === k + 1); return i != null && yt > 0 ? pct2(allSer[i].tc, yt) : NaN; }) },
    { id: 'a', name: 'Actual share', color: 'var(--s1)', values: Array.from({ length: 12 }, (_, k) => { const i = yi.find(j => MNUM[j] === k + 1); return i != null && ya > 0 ? pct2(allSer[i].ac, ya) : NaN; }) }];
  grid.append(card({ id: 'tr-phase', title: 'Phasing: share of the year’s sales by month', sub: 'If the lines differ, the target spread the year differently from how sales really came in', span: 's6', tools: [seg(YEARS.map(y => [y, String(y)]), TR.phy, v => { TR.phy = v; renderView(); }, 'Phasing year')],
    build: b => lineChart(b, { x: MNAME, series: ph, fmt: v => pctF(v, 2), tickFmt: v => v.toFixed(0) + '%', height: 280, aria: 'Monthly phasing of target and actual', extraTip: k => [{ label: 'Gap', value: ppF(ph[1].values[k] - ph[0].values[k]) }] }),
    table: () => ({ name: 'phasing_' + TR.phy, cols: [{ k: 'month', label: 'Month' }, { k: 't', label: 'Target share %', align: 'r', fmt: v => pctF(v) }, { k: 'a', label: 'Actual share %', align: 'r', fmt: v => pctF(v) }, { k: 'g', label: 'Gap (pp)', align: 'r', fmt: v => ppF(v) }], rows: MNAME.map((n, k) => ({ month: n, t: ph[0].values[k], a: ph[1].values[k], g: ph[1].values[k] - ph[0].values[k] })), search: false }) }));
  // quarters
  const qIdx = Float64Array.from({ length: N }, (_, i) => { const mo = A.M[i]; return (MYEAR[mo] - YEARS[0]) * 4 + MQ[mo] - 1; });
  const gq = agg(qIdx, YEARS.length * 4, mAll), qRows = [];
  for (let q = 0; q < YEARS.length * 4; q++) { if (!gq.n[q]) continue; const d = dv(gq, q); d.q = `${YEARS[Math.floor(q / 4)]} Q${(q % 4) + 1}`; d.qi = q; qRows.push(d); }
  qRows.forEach(d => { const p1 = qRows.find(x => x.qi === d.qi - 1), p4 = qRows.find(x => x.qi === d.qi - 4); d.qoq = p1 ? pct2(d.ac - p1.ac, p1.ac) : NaN; d.yoyq = p4 ? pct2(d.ac - p4.ac, p4.ac) : NaN; });
  grid.append(card({ id: 'tr-q', title: 'Quarterly view', sub: 'All months (time filter not applied)', span: 's12', build: b => b.append(tableView({ name: 'quarters', search: false, sortKey: null, rows: qRows,
    cols: [{ k: 'q', label: 'Quarter' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usdC(v) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usdC(v) }, { k: 'v', label: 'Sales var $', align: 'r', fmt: v => vSpan(v, usdC(v)) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'prc', label: 'Price $', align: 'r', fmt: v => vSpan(v, usdC(v)) }, { k: 'gv', label: 'GP var $', align: 'r', fmt: v => vSpan(v, usdC(v)) }, { k: 'gpct', label: 'GP var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'qoq', label: 'QoQ sales', align: 'r', fmt: v => pctF(v, 1, true) }, { k: 'yoyq', label: 'YoY sales', align: 'r', fmt: v => pctF(v, 1, true) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }] })) }));
  grid.append(card({ id: 'tr-tbl', title: 'Monthly scorecard', sub: 'Growth, year-to-date, rolling and tracking measures for the whole company · click a row to filter', span: 's12', build: b => b.append(tableView({ name: 'monthly_scorecard', search: false, rows: allSer.map((d, i) => Object.assign({ month: MLONG[i] }, d)), cols: monthCols().concat([{ k: 'mat', label: 'Material', fmt: (v, r) => matCell(v, r.trig), csv: r => (r.mat ? 'Yes' : 'No') }]), onRow: r => selectMonth(MLONG.indexOf(r.month)) })) }));
}

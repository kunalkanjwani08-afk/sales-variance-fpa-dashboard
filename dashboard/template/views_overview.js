/* ===========================================================================
   SHARED VIEW HELPERS
   =========================================================================== */
const sgnCls = v => (v > 0 ? 'pos' : v < 0 ? 'neg' : '');
function vSpan(v, txt) { return h('span', { class: sgnCls(v) }, txt); }
function monthSeries(mask) { const g = agg(A.M, NM, mask); return Array.from({ length: NM }, (_, i) => dv(g, i)); }
function selTracking(ser, inSel) {          // tracking signal over the months currently in view
  let cum = 0, abs = 0, n = 0; for (let m = 0; m < NM; m++) if (inSel[m] && ser[m].n) { cum += ser[m].vc; abs += Math.abs(ser[m].vc); n++; }
  return abs > 0 ? cum / (abs / n) : NaN;
}
function groupRows(key, K, mask, names) { const g = agg(key, K, mask); return names.map((nm, i) => Object.assign(dv(g, i), { name: nm, id: i })); }
function matCell(m, trig) { return m ? h('span', { class: 'mat' }, '◆ Yes' + (trig ? ' (' + trig + ')' : '')) : 'No'; }
function tipFor(d, name) {
  return { title: name, rows: [
    { label: 'Target sales', value: usd(d.target) }, { label: 'Actual sales', value: usd(d.actual) },
    { label: 'Sales variance', value: usd(d.v) + '  (' + pctF(d.pct, 2, true) + ')' },
    { label: 'Gross-profit variance', value: usd(d.gv) + '  (' + pctF(d.gpct, 2, true) + ')' },
    { label: 'Price effect', value: usd(d.prc) }, { label: 'Qty attainment', value: pctF(d.att) },
    { label: 'Status (group total)', value: ST_GLYPH[d.st] + ' ' + ST_NAME[d.st] + (d.mat ? '  ◆ material' : '') },
    { label: 'Lines below target', value: pctF(d.pb, 1) + ' of ' + numF(d.n) }] };
}
function itemOf(d, name, id, selected, extra) {
  return Object.assign({ id, label: name, value: d.v, txt: usdS(d.v) + '  ·  ' + pctF(d.pct, 2, true), selected, d }, extra || {});
}
function lineSum(m, s) { let t = 0; for (let i = 0; i < N; i++) if (m[i] && A.ST[i] === s) t += A.VAR[i]; return t / 100; }
function lineSd(m) { let n = 0, s = 0, q = 0; for (let i = 0; i < N; i++) if (m[i] && A.TS[i] > 0) { const p = A.VAR[i] / A.TS[i] * 100; n++; s += p; q += p * p; } return n > 1 ? Math.sqrt((q - s * s / n) / (n - 1)) : 0; }
const DISC = 5;                               // lines priced >= 5 % below target count as "discounted"
function discountSplit(m) {                   // discounted lines vs the rest
  const md = new Uint8Array(N), mo = new Uint8Array(N);
  for (let i = 0; i < N; i++) { if (!m[i]) continue; if ((A.AP[i] - A.TP[i]) * 100 <= -DISC * A.TP[i]) md[i] = 1; else mo[i] = 1; }
  return { disc: totals(md), rest: totals(mo) };
}

/* ---------- live insight sentences for the current selection ---------- */
function liveInsights() {
  const m = getMask(), T = totals(m), out = [];
  if (!T.n) return [['warn', 'No sales lines match the current filters.']];
  const inSel = monthOk(), ser = monthSeries(m);
  out.push([T.v < 0 ? 'bad' : '', `Sales are ${usdC(Math.abs(T.v))} ${T.v < 0 ? 'below' : 'above'} target (${pctF(T.pct, 2, true)}) and gross profit is ${usdC(Math.abs(T.gv))} ${T.gv < 0 ? 'below' : 'above'} target (${pctF(T.gpct, 2, true)}). Gross margin ${pctF(T.agm)} vs ${pctF(T.tgm)} planned.`]);
  out.push([T.prc < 0 ? 'warn' : '', `Revenue bridge: pure volume ${usdS(T.pure)}, mix ${usdS(T.mix)}, price ${usdS(T.prc)} (average selling price ${pctF(T.real, 2, true)} vs target). Costs vs target moved gross profit by ${usdS(T.cost)}.`]);
  const dims = [['region', A.R, NR, REG, 'reg'], ['category', A.C, NC, CAT, 'cat']];
  dims.forEach(([lab, key, K, names, ex]) => {
    const rows = groupRows(key, K, getMask(ex), names).filter(r => r.n); if (rows.length < 2) return;
    const worst = rows.reduce((a, b) => (b.v < a.v ? b : a)), best = rows.reduce((a, b) => (b.v > a.v ? b : a)), net = rows.reduce((s, r) => s + r.v, 0);
    let t = `Weakest ${lab}: ${worst.name} (${usdS(worst.v)}, ${pctF(worst.pct, 2, true)})`;
    if (net < 0 && worst.v < 0) t += `, equal to ${Math.round(worst.v / net * 100)}% of the net shortfall`;
    t += `. Strongest: ${best.name} (${usdS(best.v)}, ${pctF(best.pct, 2, true)}).`;
    out.push([worst.pct < -5 ? 'bad' : worst.v < 0 ? 'warn' : '', t]);
  });
  const below = ser.filter((d, i) => inSel[i] && d.n && d.vc < 0).length, present = ser.filter((d, i) => inSel[i] && d.n).length;
  if (present) { const worstM = ser.reduce((a, b, i) => (inSel[i] && b.n && (a === null || b.pct < ser[a].pct) ? i : a), null); out.push([below === present && present > 3 ? 'bad' : '', `${below} of ${present} month${present > 1 ? 's' : ''} in view finished below target` + (worstM != null ? `; the weakest was ${MLONG[worstM]} (${pctF(ser[worstM].pct, 2, true)}).` : '.')]); }
  const ds = discountSplit(m);
  if (ds.disc.n && ds.rest.n) out.push([ds.disc.gpct < ds.rest.gpct ? 'warn' : '', `Lines sold ${DISC}% or more below target price (${numF(ds.disc.n)} lines) reached ${pctF(ds.disc.att)} of planned units but ${pctF(ds.disc.gpct, 2, true)} gross profit vs target; all other lines: ${pctF(ds.rest.att)} of units, ${pctF(ds.rest.gpct, 2, true)} gross profit.`]);
  const pr = groupRows(A.P, NP, getMask('prod'), PRD).filter(r => r.n), neg = pr.filter(r => r.v < 0), gross = neg.reduce((s, r) => s + r.v, 0);
  if (neg.length) { const top = neg.slice().sort((a, b) => a.v - b.v).slice(0, 3); out.push(['', `${neg.length} of ${pr.length} products are below target. The three largest shortfalls (${top.map(r => r.name).join(', ')}) explain ${Math.round(top.reduce((s, r) => s + r.v, 0) / gross * 100)}% of the product-level shortfall.`]); }
  out.push(['', `${numF(T.nm)} of ${numF(T.n)} lines (${pctF(T.nm / T.n * 100, 0)}) are flagged material at line level: single lines swing a lot, so use the Exceptions tab (region × category × month) for triage.`]);
  return out;
}
function insightsList(items) { return h('ul', { class: 'ins' }, items.map(([c, t]) => h('li', { class: c }, t))); }

/* ---------- key takeaway banner (live: follows the filters) ---------- */
function takeawayBanner(T, m, mSer, inSel) {
  const sp = (c, t) => h('strong', { class: c }, t), b = t => h('strong', null, t);
  const miss = T.v < 0;
  const present = mSer.filter((d, i) => inSel[i] && d.n), below = present.filter(d => d.vc < 0).length;
  const word = v => (v < 0 ? ' below' : ' above');
  const out = [h('p', null, 'Sales finished ', sp(sgnCls(T.v), usdC(Math.abs(T.v)) + ' (' + pctF(Math.abs(T.pct), 2) + ')'), word(T.v) + ' target and gross profit ',
    sp(sgnCls(T.gv), usdC(Math.abs(T.gv)) + ' (' + pctF(Math.abs(T.gpct), 2) + ')'), word(T.gv) + ' target' + (present.length > 2 ? `, with ${below} of ${present.length} months below plan.` : '.'))];
  // drivers, largest first (revenue: price, volume and mix; gross profit: also unit costs)
  const DRV = [
    [T.prc, v => (v < 0 ? 'Lower selling prices cost ' : 'Higher selling prices added '), 'revenue'],
    [T.vol, v => (v < 0 ? 'Lower volume and mix cost ' : 'Volume and mix added '), 'revenue'],
    [T.cost, v => (v < 0 ? 'Higher unit costs took ' : 'Lower unit costs added '), 'gross profit']].filter(d => Math.abs(d[0]) >= 0.5).sort((a, c) => Math.abs(c[0]) - Math.abs(a[0]));
  if (DRV.length) {
    const p = [];
    DRV.forEach((d, i) => {
      let lead = d[1](d[0]);
      if (i > 0) { p.push(i === DRV.length - 1 ? ' and ' : ', '); lead = lead.charAt(0).toLowerCase() + lead.slice(1); }
      p.push(lead, sp(sgnCls(d[0]), usdC(Math.abs(d[0]))), ' of ' + d[2]);
    });
    p.push(T.prc < 0 && T.vol >= 0 && T.v < 0 ? ': the shortfall comes from price, not from units.' : '.');
    out.push(h('p', null, ...p));
  }
  const reg = groupRows(A.R, NR, m, REG).filter(r => r.n);
  if (reg.length >= 2) {
    const worst = reg.reduce((a, x) => (x.v < a.v ? x : a)), best = reg.reduce((a, x) => (x.v > a.v ? x : a)), net = reg.reduce((s, r) => s + r.v, 0);
    const p = [];
    if (worst.v < 0) p.push(sp('neg', worst.name), ' is the weakest region (', sp('neg', pctF(worst.pct, 2, true)), net < 0 ? `, ${Math.round(worst.v / net * 100)}% of the net shortfall)` : ')');
    if (best.v > 0) p.push(worst.v < 0 ? '; ' : '', sp('pos', best.name), ' (', sp('pos', pctF(best.pct, 2, true)), ') is the strongest.');
    else p.push('.');
    out.push(h('p', null, ...p));
  }
  const right = h('div', { class: 'tk-r' },
    h('div', { class: 'stat' }, h('b', { class: sgnCls(T.v) }, pctF(T.pct, 2, true)), h('span', null, 'sales vs target')),
    h('div', { class: 'stat' }, h('b', { class: sgnCls(T.gv) }, pctF(T.gpct, 2, true)), h('span', null, 'gross profit vs target')),
    present.length ? h('div', { class: 'stat' }, h('b', { class: below > present.length / 2 ? 'neg' : '' }, below + ' / ' + present.length), h('span', null, 'months below target')) : null);
  return h('section', { class: 'takeaway ' + (miss ? '' : 'pos'), 'aria-label': 'Key takeaway' }, h('div', { class: 'tk-l' }, h('div', { class: 'tk-k' }, (isFiltered() ? 'Key takeaway · current filters' : 'Key takeaway') + ' · synthetic data'), ...out), right);
}

/* ===========================================================================
   TAB 1 - OVERVIEW
   =========================================================================== */
const TS_LIMIT = 4;
function tabOverview(root) {
  const m = getMask(), T = totals(m);
  if (!T.n) { root.append(emptyMsg()); return; }
  const mT = getMask('time'), ser = monthSeries(mT), inSel = monthOk(), mSer = monthSeries(m), ts = selTracking(mSer, inSel);
  const cells = (() => { const g = agg(KEY_RCM, NR * NC * NM, m); let c = 0, mc = 0; for (let k = 0; k < g.K; k++) if (g.n[k]) { c++; if (isMat(g.as[k] - g.ts[k], g.ts[k])) mc++; } return [mc, c]; })();
  const monthsBelow = mSer.filter((d, i) => inSel[i] && d.n && d.vc < 0).length, monthsPresent = mSer.filter((d, i) => inSel[i] && d.n).length;
  root.append(takeawayBanner(T, m, mSer, inSel));
  const hero = h('section', { class: 'card s4 hero ' + (T.v < 0 ? '' : 'pos-t') },
    h('div', null, h('div', { class: 'sub', style: { fontWeight: 600 } }, 'SALES VARIANCE vs TARGET'),
      h('div', { class: 'big ' + sgnCls(T.v) }, usdC(T.v)),
      h('div', { class: 'row' }, h('span', { class: sgnCls(T.v), style: { fontSize: '20px', fontWeight: 650 } }, pctF(T.pct, 2, true)), stBadge(T.st), T.mat ? h('span', { class: 'mat', title: 'Material: |variance %| ≥ ' + CFG.matPctV + '% or |variance $| ≥ ' + usd(CFG.matUsd) }, '◆ Material') : null)),
    h('div', { class: 'sub' }, `${usd(T.v)} on a target of ${usd(T.target)}. Status: Favorable above +${CFG.favPct}%, Unfavorable below ${MINUS}${CFG.unfPct}%. Material (separately) at ±${CFG.matPctV}% or ${usd(CFG.matUsd)}.`),
    h('div', null, spark(ser.map(d => d.v), 220, 42), h('div', { class: 'sub' }, 'Monthly sales variance, all months (other filters applied)')));
  const tile = (lab, val, dl, cls, pri, title) => h('div', { class: 'card kpi' + (pri ? ' pri ' + (cls === 'neg' ? 'neg-t' : cls === 'pos' ? 'pos-t' : '') : ''), title: title || null }, h('div', { class: 'lab' }, lab), h('div', { class: 'val ' + (cls || '') }, val), h('div', { class: 'dl' }, dl));
  const tiles = h('div', { class: 's8 kpi-wrap' },
    h('div', { class: 'kpi-pri' },
      tile('Gross-profit variance', usdC(T.gv), `${pctF(T.gpct, 2, true)} vs target gross profit`, sgnCls(T.gv), true, 'Actual gross profit − target gross profit'),
      tile('Price effect', usdC(T.prc), `selling price ${pctF(T.real, 2, true)} vs target`, sgnCls(T.prc), true, 'Actual qty × (actual unit price − target unit price)'),
      tile('Volume effect', usdC(T.vol), `pure ${usdS(T.pure)} · mix ${usdS(T.mix)}`, sgnCls(T.vol), true, '(Actual qty − target qty) × target unit price'),
      tile('Cost effect', usdC(T.cost), 'gross profit, unit costs vs target', sgnCls(T.cost), true, '−Actual qty × (actual unit cost − target unit cost)')),
    h('div', { class: 'kpi-sec' },
      tile('Actual sales', usdC(T.actual), `${pctF(T.actual / T.target * 100)} of target`),
      tile('Target sales', usdC(T.target), `${numF(T.n)} lines · ${numF(T.tq)} units`),
      tile('Gross margin', pctF(T.agm), `target ${pctF(T.tgm)} · ${ppF(T.gmc)}`, sgnCls(T.gmc)),
      tile('Quantity attainment', pctF(T.att), `${numF(T.qv)} units vs plan`, sgnCls(T.qv)),
      tile('Lines below target', pctF(T.pb, 1), `actual < target · ${pctF(T.pu, 1)} worse than ${MINUS}${CFG.unfPct}%`, '', false, 'Share of lines whose actual sales are below target'),
      tile('Months below target', `${monthsBelow} of ${monthsPresent}`, 'months in view with sales', monthsBelow > monthsPresent / 2 ? 'neg' : ''),
      tile('Material lines', numF(T.nm), `${pctF(T.nm / T.n * 100, 1)} of lines · ${cells[0]} of ${cells[1]} cells`),
      tile('Tracking signal', isFinite(ts) ? numF(ts, 1) : '—', `limit ±${TS_LIMIT} (persistent bias)`, isFinite(ts) && Math.abs(ts) > TS_LIMIT ? 'neg' : '')));
  root.append(h('div', { class: 'grid' }, hero, tiles));
  const grid = h('div', { class: 'grid', style: { marginTop: '14px' } }); root.append(grid);

  grid.append(card({ id: 'ov-ins', title: 'What the data says right now', sub: 'Updates with every filter', span: 's4', build: b => b.append(insightsList(liveInsights())) }));
  grid.append(card({
    id: 'ov-trend', title: 'Target vs actual sales, month by month', sub: 'Click a month to filter the whole dashboard · dimmed months are outside the current time filter', span: 's8',
    build: b => trendChart(b, { m: [...Array(NM).keys()], tc: ser.map(d => d.tc), ac: ser.map(d => d.ac), vc: ser.map(d => d.vc), pct: ser.map(d => d.pct), att: ser.map(d => d.att), st: ser.map(d => d.st), mat: ser.map(d => d.mat), inSel: Array.from(inSel), onMonth: selectMonth }),
    table: () => ({ name: 'monthly_target_vs_actual', cols: monthCols(), rows: ser.map((d, i) => Object.assign({ month: MLONG[i] }, d)), sortKey: null })
  }));
  const regItems = () => groupRows(A.R, NR, getMask('reg'), REG).filter(x => x.n).sort((a, b) => a.v - b.v).map(d => itemOf(d, d.name, d.id, S.regs.has(d.id)));
  const catItems = () => groupRows(A.C, NC, getMask('cat'), CAT).filter(x => x.n).sort((a, b) => a.v - b.v).map(d => itemOf(d, d.name, d.id, S.cats.has(d.id)));
  grid.append(card({ id: 'ov-reg', title: 'Sales variance by region', sub: 'Blue = above target, red = below · click to filter (shift-click for several)', span: 's4',
    build: b => hbars(b, regItems(), { onClick: (it, ev) => pickRegion(it.id, ev), tip: it => tipFor(it.d, it.label), aria: 'Sales variance by region' }), table: () => dimTable(regItems(), 'Region') }));
  grid.append(card({ id: 'ov-cat', title: 'Sales variance by category', sub: 'Click to filter (shift-click for several)', span: 's4',
    build: b => hbars(b, catItems(), { onClick: (it, ev) => pickCat(it.id, ev), tip: it => tipFor(it.d, it.label), aria: 'Sales variance by category' }), table: () => dimTable(catItems(), 'Category') }));
  grid.append(card({ id: 'ov-status', title: 'Status of individual sales lines', sub: `Each line vs its own target: Favorable > +${CFG.favPct}%, Unfavorable < ${MINUS}${CFG.unfPct}%, otherwise ${WITHIN_LABEL}`, span: 's4', build: b => {
    const sh = h('div'); b.append(sh);
    const segs = [
      { label: 'Favorable', value: T.nf, color: 'var(--fav)', extra: [{ label: 'Variance $', value: usd(lineSum(m, ST_FAV)) }] },
      { label: WITHIN_LABEL, value: T.nw, color: 'var(--axis)', extra: [{ label: 'Variance $', value: usd(lineSum(m, ST_WITHIN)) }] },
      { label: 'Unfavorable', value: T.nu, color: 'var(--unf)', extra: [{ label: 'Variance $', value: usd(lineSum(m, ST_UNF)) }] }];
    if (T.nn) segs.push({ label: 'No target', value: T.nn, color: 'var(--muted)', extra: [{ label: 'Variance $', value: usd(lineSum(m, ST_NOTGT)) }] });
    stack100(sh, segs);
    b.append(h('div', { class: 'note' }, `Single lines swing widely (typical ±${lineSd(m).toFixed(0)}%), so line counts are noisy; the totals are what matter. ${numF(T.nu)} lines are Unfavorable and ${numF(T.nf)} Favorable.` + (T.nn ? ` ${numF(T.nn)} lines have no target (new accounts).` : '')));
  } }));
  const bd = bridgeDim(); const bm = getMask(bd.ex);
  const gb = groupRows(bd.key, bd.K, bm, bd.names).filter(x => x.n).sort((a, b) => b.v - a.v);
  const steps = [{ label: 'Target', value: T.target, kind: 'total', color: 'var(--target)' }].concat(gb.map(d => ({ label: d.name, value: d.v, kind: 'delta', extra: [{ label: 'Variance %', value: pctF(d.pct, 2, true) }] })), [{ label: 'Actual', value: T.actual, kind: 'total', color: 'var(--s1)' }]);
  grid.append(card({ id: 'ov-bridge', title: `Target-to-actual sales bridge by ${bd.label}`, sub: 'How each ' + bd.label + ' moves the company from target to actual', span: 's8', build: b => waterfall(b, steps, { aria: 'Bridge from target to actual by ' + bd.label }),
    table: () => ({ name: 'bridge', cols: [{ k: 'label', label: bd.label }, { k: 'value', label: 'USD', align: 'r', fmt: v => usd(v) }], rows: steps, search: false }) }));
  grid.append(card({ id: 'ov-hint', title: 'How to read this dashboard', sub: '', span: 's4', build: b => b.append(h('ul', { class: 'ins' },
    h('li', null, `Every figure is computed in your browser from the ${numF(N)} sales lines, in exact cents. The data is synthetic.`),
    h('li', null, 'Click any bar, month, heat-map cell or table row to filter everything (Tab + Enter works too). Shift-click adds to a selection.'),
    h('li', null, `Status and materiality are separate: a group can be ${WITHIN_LABEL} and still material in dollars.`),
    h('li', null, 'Use ⊞ Table on any chart to see the exact numbers; ⬇ Export saves the filtered lines.'))) }));
}
function bridgeDim() {
  if (S.regs.size === 1 && S.cats.size === 1) return { key: A.P, K: NP, names: PRD, ex: 'prod', label: 'product' };
  if (S.regs.size === 1) return { key: A.C, K: NC, names: CAT, ex: 'cat', label: 'category' };
  return { key: A.R, K: NR, names: REG, ex: 'reg', label: 'region' };
}
function dimTable(items, label) {
  return { name: label.toLowerCase() + '_variance', cols: [{ k: 'label', label }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Sales var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) },
    { k: 'prc', label: 'Price effect $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'gv', label: 'GP var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'gpct', label: 'GP var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) }, { k: 'att', label: 'Qty attainment', align: 'r', fmt: v => pctF(v) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }], rows: items.map(i => Object.assign({ label: i.label }, i.d)), search: false, sortKey: 'v', sortDir: 1 };
}
function monthCols() {
  return [{ k: 'month', label: 'Month' }, { k: 'target', label: 'Target $', align: 'r', fmt: v => usd(v) }, { k: 'actual', label: 'Actual $', align: 'r', fmt: v => usd(v) }, { k: 'v', label: 'Sales var $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'pct', label: 'Sales var %', align: 'r', fmt: v => vSpan(v, pctF(v, 2, true)) },
    { k: 'prc', label: 'Price effect $', align: 'r', fmt: v => vSpan(v, usd(v)) }, { k: 'gv', label: 'GP var $', align: 'r', fmt: v => vSpan(v, usd(v)) },
    { k: 'att', label: 'Qty att. %', align: 'r', fmt: v => pctF(v) }, { k: 'mom', label: 'MoM sales', align: 'r', fmt: v => pctF(v, 2, true) }, { k: 'yoy', label: 'YoY sales', align: 'r', fmt: v => pctF(v, 2, true) }, { k: 'ytdV', label: 'YTD var $', align: 'r', fmt: v => usd(v) },
    { k: 'r3', label: 'Rolling 3M var $', align: 'r', fmt: v => usd(v) }, { k: 'cum', label: 'Cumulative var $', align: 'r', fmt: v => usd(v) }, { k: 'ts', label: 'Tracking signal', align: 'r', fmt: v => numF(v, 1) }, { k: 'st', label: 'Status', fmt: v => stBadge(v), csv: r => ST_NAME[r.st] }];
}

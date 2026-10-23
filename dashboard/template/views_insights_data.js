/* ===========================================================================
   TAB 7 - INSIGHTS & RECOMMENDATIONS     TAB 8 - DATA & VALIDATION
   =========================================================================== */
let ALL_MASK = null;
function maskAll() { if (!ALL_MASK) { ALL_MASK = new Uint8Array(N).fill(1); } return ALL_MASK; }

/* Each headline statement in the written report is recomputed here from the line data.
   The EXPECTED values come from docs/key_metrics.json (written by scripts/build_insights.py),
   never from numbers typed into this file. */
function claimValue(key) {
  const all = maskAll(), T = totals(all);
  const reg = groupRows(A.R, NR, all, REG), prd = groupRows(A.P, NP, all, PRD);
  const worst = reg.reduce((a, x) => (x.v < a.v ? x : a)), best = reg.reduce((a, x) => (x.v > a.v ? x : a));
  switch (key) {
    case 'sales_variance_usd': return T.v;
    case 'sales_variance_pct': return T.pct;
    case 'gp_variance_usd': return T.gv;
    case 'gp_variance_pct': return T.gpct;
    case 'pure_volume_usd': return T.pure;
    case 'mix_usd': return T.mix;
    case 'price_effect_usd': return T.prc;
    case 'cost_effect_usd': return T.cost;
    case 'price_realisation_pct': return T.real;
    case 'qty_attainment_pct': return T.att;
    case 'gross_margin_target_pct': return T.tgm;
    case 'gross_margin_actual_pct': return T.agm;
    case 'months_below': return monthSeries(all).filter(d => d.n && d.vc < 0).length;
    case 'worst_region': return worst.name;
    case 'worst_region_pct': return worst.pct;
    case 'worst_region_share_of_net_pct': return T.vc ? pct2(Math.sign(T.vc) * worst.vc, Math.abs(T.vc)) : NaN;
    case 'best_region': return best.name;
    case 'best_region_pct': return best.pct;
    case 'products_above_target': return prd.filter(p => p.n && p.v > 0).length;
    case 'persistent_pairs': return persistenceRows(all).filter(p => p.chronic).length;
    case 'discounted_lines_qty_att_pct': return discountSplit(all).disc.att;
    case 'discounted_lines_gp_var_pct': return discountSplit(all).disc.gpct;
    case 'other_lines_gp_var_pct': return discountSplit(all).rest.gpct;
    default: return undefined;
  }
}
function headlineClaims() {
  return (D.metrics && D.metrics.claims || []).map(c => {
    const got = claimValue(c.key);
    const ok = typeof c.expected === 'string' ? got === c.expected : (isFinite(got) && Math.abs(got - c.expected) <= (c.tol || 0) + 1e-9);
    const show = typeof got === 'string' ? got : c.unit === 'usd' ? usd(got, 2) : c.unit === 'pct' ? pctF(got, 2, true) : numF(got);
    return [c.label, c.text, show, ok];
  });
}
function tabInsights(root) {
  const grid = h('div', { class: 'grid' }); root.append(grid);
  grid.append(card({ id: 'in-md', title: 'Insights & recommendations', sub: 'Written for the full period (not affected by the filters) · synthetic data', span: 's8', build: b => {
    if (D.md) {
      const box = h('div', { class: 'md' }); box.innerHTML = D.md;       // HTML built by the Python builder from the .md file, with all text escaped
      const heads = [...box.querySelectorAll('h3')];
      if (heads.length > 1) { const jb = h('div', { class: 'md-jump' }); heads.forEach(x => { const t = x.textContent.trim(); x.style.scrollMarginTop = '70px'; jb.append(h('button', { class: 'chip', onclick: () => x.scrollIntoView({ behavior: 'smooth', block: 'start' }) }, t.length > 34 ? t.slice(0, 32) + '…' : t)); }); b.append(jb); }
      b.append(box);
    } else b.append(h('div', { class: 'note' }, 'docs/Insights_and_Recommendations.md was not found, so only the live insights are shown. Run scripts/build_insights.py.'));
  } }));
  const side = h('div', { class: 's4', style: { display: 'flex', flexDirection: 'column', gap: '14px' } });
  const rows = headlineClaims(), okN = rows.filter(r => r[3]).length;
  const c1 = card({ id: 'in-claims', title: 'Does the data still say this?', sub: rows.length ? `${okN} of ${rows.length} headline statements re-computed from the line data` : 'docs/key_metrics.json not found', span: '', build: b => b.append(h('div', { class: 'tbl-wrap', style: { maxHeight: '520px' } }, h('table', { class: 't' }, h('thead', null, h('tr', null, ['', 'Statement', 'In the report', 'Recomputed now'].map(x => h('th', null, x)))),
    h('tbody', null, rows.map(r => h('tr', null, h('td', { class: r[3] ? 'pass' : 'fail' }, r[3] ? '✔' : '✖'), h('td', { style: { whiteSpace: 'normal' } }, r[0]), h('td', { style: { whiteSpace: 'normal' } }, r[1]), h('td', { style: { whiteSpace: 'normal' } }, r[2])))))))
  });
  const c2 = card({ id: 'in-live', title: 'Live insights for your current filters', sub: 'The same analysis applied to whatever you have selected', span: '', build: b => b.append(insightsList(liveInsights())) });
  side.append(c1, c2); grid.append(side);
}

/* ---------- reconciliation of the in-browser numbers with the Excel workbook ---------- */
function reconcile() {
  const xl = D.xl, out = []; if (!xl) return out;
  const add = (name, items, maxDiff, tol, detail) => out.push({ name, items, maxDiff, tol, pass: maxDiff <= tol + 1e-9, detail: detail || '' });
  const all = maskAll();
  const cmpGroups = (key, K, names, ref, label) => {
    const rows = groupRows(key, K, all, names); let dvv = 0, dg = 0, cnt = 0;
    rows.forEach(r => { const x = ref[r.name]; if (!x) { dvv = Infinity; return; } cnt++; dvv = Math.max(dvv, Math.abs(r.v - x.v)); dg = Math.max(dg, Math.abs(r.gv - x.gv), Math.abs(r.prc - x.prc), Math.abs(r.mix - x.mix)); });
    add(`${label}: sales variance (${cnt})`, cnt, dvv, 0.005); add(`${label}: gross-profit variance, price and mix (${cnt})`, cnt, dg, 0.005);
  };
  cmpGroups(A.R, NR, REG, xl.region, 'By region'); cmpGroups(A.C, NC, CAT, xl.category, 'By category'); cmpGroups(A.P, NP, PRD, xl.product, 'By product');
  const ser = monthlySeries(all), mx = (f, tol) => { let d = 0, n = 0; MON.forEach((m, i) => { const r = xl.month[m]; if (!r || r[f] == null) return; const v = { var: ser[i].v, pct: ser[i].pct, ts: ser[i].ts, ytd: ser[i].ytdV, r3: ser[i].r3, mom: ser[i].mom, yoy: ser[i].yoy, gv: ser[i].gv }[f]; if (!isFinite(v)) { d = Infinity; return; } n++; d = Math.max(d, Math.abs(v - r[f])); }); return [n, d]; };
  [['var', 'Monthly sales variance $', 0.005], ['pct', 'Monthly sales variance %', 0.0051], ['gv', 'Monthly gross-profit variance $', 0.005], ['ts', 'Tracking signal by month', 0.006], ['ytd', 'YTD sales variance', 0.005], ['r3', 'Rolling 3-month sales variance', 0.005], ['mom', 'Month-on-month sales growth %', 0.0051], ['yoy', 'Year-on-year sales growth %', 0.0051]].forEach(([f, l, t]) => { const [n, d] = mx(f); add(`${l} (${n})`, n, d, t); });
  const yrs = Object.keys(xl.year || {}); let dy = 0; yrs.forEach(y => { const m = new Uint8Array(N); for (let i = 0; i < N; i++) m[i] = MYEAR[A.M[i]] === +y ? 1 : 0; const t = totals(m); dy = Math.max(dy, Math.abs(t.v - xl.year[y].v), Math.abs(t.gv - xl.year[y].gv)); }); add('Sales and gross-profit variance by year', yrs.length, dy, 0.005);
  const T = totals(all), c = xl.company;
  const db = Math.max(...['pure', 'mix', 'prc', 'cost', 'ppure', 'pmix'].map(k => Math.abs(T[k] - c[k])));
  add('Company revenue and gross-profit bridges (6 effects)', 6, db, 0.005);
  const pers = persistenceRows(all); let dp = 0; pers.forEach(p => { const r = xl.pers[p.region + '|' + p.category]; dp = Math.max(dp, r == null ? Infinity : Math.abs(p.below - r)); });
  add(`Months below target per region × category (${pers.length})`, pers.length, dp, 0); add('Persistent shortfall pairs', 1, Math.abs(pers.filter(p => p.chronic).length - xl.persistent), 0);
  const g = agg(KEY_RCM, NR * NC * NM, all); let mc = 0; for (let k = 0; k < g.K; k++) if (g.n[k] && isMat(g.as[k] - g.ts[k], g.ts[k])) mc++;
  add('Material region × category × month cells', 1, Math.abs(mc - xl.exc_cells), 0, mc + ' vs ' + xl.exc_cells);
  let ml = 0; for (let i = 0; i < N; i++) ml += A.MT[i]; add('Material lines', 1, Math.abs(ml - xl.mat_lines), 0, numF(ml) + ' vs ' + numF(xl.mat_lines));
  const gb = agg(KEY_BAND, BANDS.length, all); let dbd = 0; BANDS.forEach((bd, i) => { const r = xl.bands[i]; if (!r) { dbd = Infinity; return; } const d = dv(gb, i); dbd = Math.max(dbd, Math.abs(d.gv - r.gv), Math.abs(d.n - r.n)); }); add('Price bands: lines and gross-profit variance (4)', 4, dbd, 0.005);
  // every line re-classified from cents vs the status and materiality written by the Python analysis
  let bad = 0; for (let i = 0; i < N; i++) { if (statusOf(A.VAR[i], A.TS[i]) !== A.ST[i] || matOf(A.VAR[i], A.TS[i]) !== A.TRG[i]) bad++; }
  add(`Line status and materiality re-derived from cents (${numF(N)} lines)`, N, bad, 0, bad + ' differences');
  return out;
}
let RECON = [];
function tabData(root) {
  const grid = h('div', { class: 'grid' }); root.append(grid);
  const okN = RECON.filter(r => r.pass).length;
  grid.append(card({ id: 'dt-rec', title: 'Reconciliation: dashboard vs Excel workbook', sub: D.xl ? `${okN} of ${RECON.length} checks pass · recomputed from the line data when this page loaded` : 'Workbook not found - nothing to reconcile', span: 's7', build: b => {
    if (!D.xl) { b.append(h('div', { class: 'note' }, 'Run scripts/sales_variance_analysis.py so outputs/Sales_Variance_Analysis.xlsx exists, then rebuild the dashboard.')); return; }
    b.append(h('div', { class: 'tbl-wrap' }, h('table', { class: 't' }, h('thead', null, h('tr', null, ['', 'Check', 'Items', 'Largest difference'].map(x => h('th', null, x)))), h('tbody', null, RECON.map(r => h('tr', null, h('td', { class: r.pass ? 'pass' : 'fail' }, r.pass ? '✔' : '✖'), h('td', { style: { whiteSpace: 'normal' } }, r.name), h('td', { class: 'r' }, numF(r.items)), h('td', { class: 'r' }, r.detail || (r.maxDiff === 0 ? 'identical' : r.maxDiff.toFixed(4)))))))));
  } }));
  grid.append(card({ id: 'dt-prof', title: 'Dataset profile', span: 's5', build: b => {
    const rep = Array.from(A.RP).reduce((a, c) => a + c, 0), notes = Array.from(A.NT).filter(v => v > 0).length, big = Array.from(A.DQ).reduce((a, c) => a + c, 0), cl = D.cleaning;
    const kv = [['Source', D.source], ['Data', D.note], ['Clean sales lines', numF(N)], ['Period', MLONG[0] + ' – ' + MLONG[NM - 1] + ` (${NM} months)`], ['Regions / categories / products', `${NR} / ${NC} / ${NP}`]];
    if (cl) kv.push(['Raw rows → clean lines', `${numF(cl.raw_rows)} → ${numF(cl.clean)} (${numF(cl.exact)} exact duplicates, ${numF(cl.merged)} merged copies, ${numF(cl.blank)} blank rows, ${numF(cl.review)} held for review)`], ['Values repaired by the cleaning step (all exact)', numF(cl.repaired)]);
    kv.push(['Lines with at least one repaired value', `${numF(rep)} (${pctF(rep / N * 100, 1)})`], ['Large orders flagged for checking (kept)', numF(big)], ['Lines with an operational note', numF(notes)],
      ['Status thresholds', `Favorable > +${CFG.favPct}% · Unfavorable < ${MINUS}${CFG.unfPct}%`], ['Material rule (separate from status)', `|var %| ≥ ${CFG.matPctV}% or |var $| ≥ ${usd(CFG.matUsd)}`], ['Arithmetic', 'Integer cents; percentages rounded half away from zero'], ['Dashboard built', D.built]);
    b.append(h('table', { class: 't' }, h('tbody', null, kv.map(([k, v]) => h('tr', null, h('td', { style: { color: 'var(--ink2)', whiteSpace: 'normal' } }, k), h('td', { style: { whiteSpace: 'normal', fontWeight: 600 } }, v))))));
  } }));
  const logRows = [].concat((D.cleaning && D.cleaning.checks || []).map(r => ({ src: 'Cleaning', s: r[0], c: r[1], d: r[2] })), (D.xl && D.xl.validation || []).map(r => ({ src: 'Analysis', s: r[0], c: r[1], d: r[2] })));
  if (logRows.length) grid.append(card({ id: 'dt-val', title: 'Validation log from the Python pipeline', sub: 'Every check the cleaning and analysis scripts ran', span: 's7', build: b => b.append(tableView({ name: 'validation_log', rows: logRows, search: true, pageSize: 20,
    cols: [{ k: 'src', label: 'Stage' }, { k: 's', label: 'Result', fmt: v => h('span', { class: v === 'PASS' ? 'pass' : v === 'FAIL' ? 'fail' : 'warnc' }, v) }, { k: 'c', label: 'Check', fmt: v => h('span', { style: { whiteSpace: 'normal' } }, v) }, { k: 'd', label: 'Detail', fmt: v => h('span', { style: { whiteSpace: 'normal' } }, v) }] })) }));
  if (D.xl) grid.append(card({ id: 'dt-dic', title: 'Metric dictionary', sub: 'How every measure is defined', span: 's5', build: b => b.append(tableView({ name: 'metric_dictionary', rows: D.xl.dictionary.map(r => ({ m: r[0], d: r[1], n: r[2] })), search: true, pageSize: 12,
    cols: [{ k: 'm', label: 'Metric', fmt: v => h('strong', null, v) }, { k: 'd', label: 'Definition', fmt: v => h('span', { style: { whiteSpace: 'normal' } }, v) }, { k: 'n', label: 'Notes', fmt: v => h('span', { style: { whiteSpace: 'normal' } }, v) }] })) }));
}

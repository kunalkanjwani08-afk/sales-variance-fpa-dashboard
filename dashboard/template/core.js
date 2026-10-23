'use strict';
/* ===========================================================================
   CORE: data decode, exact integer maths, filters, aggregation
   Money is held in integer CENTS. Every rule matches fpa_lib/finance.py.
   =========================================================================== */
const D = JSON.parse(document.getElementById('payload').textContent);
const DIM = D.dims, CFG = D.cfg, N = D.rows.R.length;
const REG = DIM.regions, CAT = DIM.cats, PRD = DIM.prods, MON = DIM.months, NOTES = DIM.notes;
const NR = REG.length, NC = CAT.length, NP = PRD.length, NM = MON.length;
const A = {}; for (const k in D.rows) A[k] = Float64Array.from(D.rows[k]);
const IDS = D.ids;                                   // Line_ID text (kept as text: leading zeros preserved)
const MNAME = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MYEAR = MON.map(m => +m.slice(0, 4)), MNUM = MON.map(m => +m.slice(5, 7));
const MQ = MNUM.map(n => Math.floor((n - 1) / 3) + 1);
const MORD = MON.map((m, i) => MYEAR[i] * 12 + MNUM[i] - 1);
const MLAB = MON.map((m, i) => MNAME[MNUM[i] - 1] + ' ' + String(MYEAR[i]).slice(2));
const MLONG = MON.map((m, i) => MNAME[MNUM[i] - 1] + ' ' + MYEAR[i]);
const YEARS = [...new Set(MYEAR)];
const $ = (s, el = document) => el.querySelector(s);
const SVGNS = 'http://www.w3.org/2000/svg';

/* ---------- line-level amounts (cents), computed once ----------
   Sales = Qty x Unit price ; Gross profit = Sales - Qty x Unit cost
   Volume = (AQ - TQ) x TP ; Price = AQ x (AP - TP) ; Cost = -AQ x (ACU - TCU)
   Profit volume = (AQ - TQ) x (TP - TCU)                                       */
(function derive() {
  const f = () => new Float64Array(N);
  Object.assign(A, { TS: f(), AS: f(), TG: f(), AG: f(), VAR: f(), GVAR: f(), VOL: f(), PRC: f(), COST: f(), PVOL: f(), AQTP: f(), AQTM: f() });
  for (let i = 0; i < N; i++) {
    const tq = A.TQ[i], aq = A.AQ[i], tp = A.TP[i], ap = A.AP[i], tcu = A.TCU[i], acu = A.ACU[i];
    A.TS[i] = tq * tp; A.AS[i] = aq * ap; A.TG[i] = A.TS[i] - tq * tcu; A.AG[i] = A.AS[i] - aq * acu;
    A.VAR[i] = A.AS[i] - A.TS[i]; A.GVAR[i] = A.AG[i] - A.TG[i];
    A.VOL[i] = (aq - tq) * tp; A.PRC[i] = aq * (ap - tp); A.COST[i] = -aq * (acu - tcu); A.PVOL[i] = (aq - tq) * (tp - tcu);
    A.AQTP[i] = aq * tp; A.AQTM[i] = aq * (tp - tcu);
  }
})();

/* ---------- DOM helpers (text always via textContent: data never becomes markup) ---------- */
function h(tag, props, ...kids) {
  const el = document.createElement(tag);
  if (props) for (const k in props) {
    const v = props[k];
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else if (k === 'text') el.textContent = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat()) if (c != null && c !== false) el.append(c.nodeType ? c : document.createTextNode(String(c)));
  return el;
}
function sv(tag, attrs, ...kids) {
  const el = document.createElementNS(SVGNS, tag);
  if (attrs) for (const k in attrs) {
    const v = attrs[k]; if (v == null) continue;
    if (typeof v === 'string' && v.startsWith('var(') && (k === 'fill' || k === 'stroke')) el.style.setProperty(k, v);
    else if (k === 'text') el.textContent = v;
    else el.setAttribute(k, v);
  }
  for (const c of kids.flat()) if (c != null && c !== false) el.append(c.nodeType ? c : document.createTextNode(String(c)));
  return el;
}

/* ---------- exact integer helpers (same rules as the Python analysis) ---------- */
function pct2(num, den) {            // num/den*100, 2 dp, half away from zero, exact on integers
  if (!(den > 0)) return NaN;
  const q = Math.floor((2 * Math.abs(num) * 10000 + den) / (2 * den));
  return (num < 0 ? -q : q) / 100;
}
function divRound(num, den) {        // BigInt division rounded half away from zero (exact for any size)
  const n = BigInt(num), d = BigInt(den), neg = (n < 0n) !== (d < 0n);
  const a = n < 0n ? -n : n, b = d < 0n ? -d : d;
  const q = (2n * a + b) / (2n * b);
  return Number(neg ? -q : q);
}
function purePart(tq, aq, val) {     // pure volume = (sum AQ - sum TQ) x group average target value
  return tq > 0 ? divRound(BigInt(Math.round(aq - tq)) * BigInt(Math.round(val)), Math.round(tq)) : 0;
}
const ST_FAV = 0, ST_WITHIN = 1, ST_UNF = 2, ST_NOTGT = 3;
function statusOf(vc, tc) {          // exact: compare variance x 10,000 with threshold (basis points) x target
  if (!(tc > 0)) return ST_NOTGT;
  if (vc * 10000 > CFG.favBp * tc) return ST_FAV;
  if (vc * 10000 < -CFG.unfBp * tc) return ST_UNF;
  return ST_WITHIN;
}
function matOf(vc, tc) {             // -> 0 none, 1 percent, 2 amount, 3 both
  const byPct = tc > 0 && Math.abs(vc) * 10000 >= CFG.matBp * tc, byUsd = Math.abs(vc) >= CFG.matUsdC;
  return (byPct ? 1 : 0) + (byUsd ? 2 : 0);
}
function isMat(vc, tc) { return matOf(vc, tc) > 0; }
const TRIG_NAME = ['', 'Percent', 'Amount', 'Percent and amount'];

/* ---------- formatting ---------- */
const MINUS = '–';
function sgnNum(v, body) { return (v < 0 ? MINUS : '') + body; }
function usd(v, dp = 0) { if (!isFinite(v)) return '—'; return sgnNum(v, '$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp })); }
function usdC(v, dp = 2) {
  if (!isFinite(v)) return '—'; const a = Math.abs(v);
  if (a >= 1e9) return sgnNum(v, '$' + (a / 1e9).toFixed(dp) + 'B');
  if (a >= 1e6) return sgnNum(v, '$' + (a / 1e6).toFixed(dp) + 'M');
  if (a >= 1e3) return sgnNum(v, '$' + (a / 1e3).toFixed(a >= 1e5 ? 0 : 1) + 'K');
  return sgnNum(v, '$' + a.toFixed(0));
}
function usdS(v, dp = 2) { return (v > 0 ? '+' : '') + usdC(v, dp); }          // signed compact
function pctF(v, dp = 2, plus = false) { if (!isFinite(v)) return '—'; return (v < 0 ? MINUS : (plus && v > 0 ? '+' : '')) + Math.abs(v).toFixed(dp) + '%'; }
function numF(v, dp = 0) { if (!isFinite(v)) return '—'; return sgnNum(v, Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp })); }
function ppF(v, dp = 2) { if (!isFinite(v)) return '—'; return (v > 0 ? '+' : v < 0 ? MINUS : '') + Math.abs(v).toFixed(dp) + ' pp'; }
const WITHIN_LABEL = CFG.favPct === CFG.unfPct ? `Within ±${CFG.favPct}%` : 'Within tolerance';
const ST_NAME = ['Favorable', WITHIN_LABEL, 'Unfavorable', 'No target'];
const ST_GLYPH = ['▲', '●', '▼', '○'];
const ST_CLS = ['fav', 'on', 'unf', 'nt'];
function stBadge(s) { return h('span', { class: 'st ' + ST_CLS[s] }, ST_GLYPH[s] + ' ' + ST_NAME[s]); }
function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
const regColor = i => `var(--s${(i % 8) + 1})`, catColor = i => `var(--s${(i % 8) + 1})`;

/* ---------- state & filters ---------- */
const S = { years: new Set(), qtrs: new Set(), regs: new Set(), cats: new Set(), prods: new Set(), stat: new Set(), mat: 'all', dq: 'all', m0: 0, m1: NM - 1 };
const maskCache = new Map();
function bump() { maskCache.clear(); }
function monthOk() {
  const ok = new Uint8Array(NM);
  for (let m = 0; m < NM; m++) ok[m] = (m >= S.m0 && m <= S.m1 && (!S.years.size || S.years.has(MYEAR[m])) && (!S.qtrs.size || S.qtrs.has(MQ[m]))) ? 1 : 0;
  return ok;
}
function getMask(excl = '') {          // excl: comma list of filters to ignore ('time','reg','cat','prod','stat','mat','dq')
  const key = excl; if (maskCache.has(key)) return maskCache.get(key);
  const ex = new Set(excl.split(',').filter(Boolean));
  const mOk = ex.has('time') ? null : monthOk();
  const useR = !ex.has('reg') && S.regs.size, useC = !ex.has('cat') && S.cats.size, useP = !ex.has('prod') && S.prods.size;
  const useS = !ex.has('stat') && S.stat.size, useM = !ex.has('mat') && S.mat !== 'all', useQ = !ex.has('dq') && S.dq !== 'all';
  const m = new Uint8Array(N), R = A.R, C = A.C, P = A.P, M = A.M, ST = A.ST, MT = A.MT, RP = A.RP;
  for (let i = 0; i < N; i++) {
    if (mOk && !mOk[M[i]]) continue;
    if (useR && !S.regs.has(R[i])) continue;
    if (useC && !S.cats.has(C[i])) continue;
    if (useP && !S.prods.has(P[i])) continue;
    if (useS && !S.stat.has(ST[i])) continue;
    if (useM && (S.mat === 'yes') !== (MT[i] === 1)) continue;
    if (useQ && (S.dq === 'repaired') !== (RP[i] === 1)) continue;
    m[i] = 1;
  }
  maskCache.set(key, m); return m;
}
function maskCount(m) { let c = 0; for (let i = 0; i < N; i++) c += m[i]; return c; }

/* ---------- aggregation ---------- */
const SUMS = ['tq', 'aq', 'ts', 'as', 'tg', 'ag', 'vol', 'prc', 'cost', 'pvol', 'aqtp', 'aqtm'];
const SRC = { tq: 'TQ', aq: 'AQ', ts: 'TS', as: 'AS', tg: 'TG', ag: 'AG', vol: 'VOL', prc: 'PRC', cost: 'COST', pvol: 'PVOL', aqtp: 'AQTP', aqtm: 'AQTM' };
function agg(key, K, mask) {
  const g = { K, n: new Float64Array(K), nf: new Float64Array(K), nw: new Float64Array(K), nu: new Float64Array(K), nn: new Float64Array(K), nm: new Float64Array(K), nb: new Float64Array(K) };
  SUMS.forEach(s => { g[s] = new Float64Array(K); });
  const arr = SUMS.map(s => [g[s], A[SRC[s]]]), ST = A.ST, MT = A.MT, VAR = A.VAR;
  for (let i = 0; i < N; i++) {
    if (!mask[i]) continue;
    const k = key[i];
    g.n[k]++;
    for (let j = 0; j < arr.length; j++) arr[j][0][k] += arr[j][1][i];
    const s = ST[i]; if (s === ST_FAV) g.nf[k]++; else if (s === ST_UNF) g.nu[k]++; else if (s === ST_NOTGT) g.nn[k]++; else g.nw[k]++;
    if (MT[i] === 1) g.nm[k]++;
    if (VAR[i] < 0) g.nb[k]++;
  }
  return g;
}
function dv(g, k) {              // derived metrics of group k (dollar fields in USD, *c fields in cents)
  const ts = g.ts[k], as = g.as[k], vc = as - ts, tq = g.tq[k], aq = g.aq[k], tg = g.tg[k], ag = g.ag[k], gvc = ag - tg, n = g.n[k];
  const pure = purePart(tq, aq, ts), ppure = purePart(tq, aq, tg), tgm = pct2(tg, ts), agm = pct2(ag, as);
  return {
    n, tq, aq, tc: ts, ac: as, vc, target: ts / 100, actual: as / 100, v: vc / 100,
    pct: pct2(vc, ts), att: pct2(aq, tq), qv: aq - tq, st: statusOf(vc, ts), mat: isMat(vc, ts), trig: TRIG_NAME[matOf(vc, ts)],
    vol: g.vol[k] / 100, pure: pure / 100, mix: (g.vol[k] - pure) / 100, prc: g.prc[k] / 100, cost: g.cost[k] / 100,
    tgc: tg, agc: ag, gvc, tgp: tg / 100, agp: ag / 100, gv: gvc / 100, gpct: pct2(gvc, tg),
    pvol: g.pvol[k] / 100, ppure: ppure / 100, pmix: (g.pvol[k] - ppure) / 100,
    tgm, agm, gmc: isFinite(tgm) && isFinite(agm) ? Math.round((agm - tgm) * 100) / 100 : NaN,
    real: pct2(g.prc[k], g.aqtp[k]),
    nf: g.nf[k], nw: g.nw[k], nu: g.nu[k], nn: g.nn[k], nm: g.nm[k], nb: g.nb[k],
    pb: n ? g.nb[k] / n * 100 : NaN, pu: n ? g.nu[k] / n * 100 : NaN,
  };
}
const ONE = new Float64Array(N);                 // all zeros: key for grand totals
function totals(mask) { return dv(agg(ONE, 1, mask), 0); }
const KEY_RC = Float64Array.from({ length: N }, (_, i) => A.R[i] * NC + A.C[i]);
const KEY_RM = Float64Array.from({ length: N }, (_, i) => A.R[i] * NM + A.M[i]);
const KEY_RCM = Float64Array.from({ length: N }, (_, i) => (A.R[i] * NC + A.C[i]) * NM + A.M[i]);

/* ---------- monthly series with growth metrics (exact, matched on calendar month) ---------- */
function monthlyFrom(g, off) {
  const out = [];
  for (let m = 0; m < NM; m++) out.push(dv(g, off + m));
  const lagIdx = (m, lag) => MORD.indexOf(MORD[m] - lag);
  let cum = 0, absSum = 0, cnt = 0;
  const ytd = {};
  for (let m = 0; m < NM; m++) {
    const d = out[m];
    const p1 = lagIdx(m, 1), p12 = lagIdx(m, 12), p2 = lagIdx(m, 2);
    d.mom = p1 >= 0 ? pct2(d.ac - out[p1].ac, out[p1].ac) : NaN;
    d.yoy = p12 >= 0 ? pct2(d.ac - out[p12].ac, out[p12].ac) : NaN;
    const y = MYEAR[m]; ytd[y] = ytd[y] || { t: 0, a: 0 }; ytd[y].t += d.tc; ytd[y].a += d.ac;
    d.ytdV = (ytd[y].a - ytd[y].t) / 100; d.ytdP = pct2(ytd[y].a - ytd[y].t, ytd[y].t);
    if (p1 >= 0 && p2 >= 0) { const t3 = d.tc + out[p1].tc + out[p2].tc, a3 = d.ac + out[p1].ac + out[p2].ac; d.r3 = (a3 - t3) / 100; d.r3p = pct2(a3 - t3, t3); } else { d.r3 = NaN; d.r3p = NaN; }
    cum += d.vc; absSum += Math.abs(d.vc); cnt++;
    d.cum = cum / 100; d.ts = absSum > 0 ? cum / (absSum / cnt) : NaN;
  }
  return out;
}
function monthlySeries(mask) { return monthlyFrom(agg(A.M, NM, mask), 0); }

/* ---------- tooltip (mouse and keyboard) ---------- */
const tipEl = document.getElementById('tip');
function showTip(ev, title, rows, foot) {
  tipEl.replaceChildren();
  if (title) tipEl.append(h('div', { class: 't' }, title));
  for (const r of rows || []) {
    const k = h('span', { class: 'k' }, r.color ? h('span', { class: 'sw', style: { borderTopColor: r.color } }) : null, r.label);
    tipEl.append(h('div', { class: 'r' }, k, h('span', { class: 'v' }, r.value)));
  }
  if (foot) tipEl.append(h('div', { class: 'f' }, foot));
  tipEl.style.opacity = 1;
  const pad = 14, w = tipEl.offsetWidth, hh = tipEl.offsetHeight;
  let x = ev.clientX + pad, y = ev.clientY + pad;
  if (x + w > innerWidth - 8) x = ev.clientX - w - pad;
  if (y + hh > innerHeight - 8) y = innerHeight - hh - 8;
  tipEl.style.left = Math.max(8, Math.min(x, innerWidth - w - 8)) + 'px'; tipEl.style.top = Math.max(8, y) + 'px';
}
function hideTip() { tipEl.style.opacity = 0; }
/* make an SVG mark reachable by keyboard: Tab to focus (shows the tooltip), Enter/Space to activate */
function focusable(el, label, activate, tip) {
  el.setAttribute('tabindex', '0'); el.setAttribute('role', activate ? 'button' : 'img'); el.setAttribute('aria-label', label);
  if (tip) {
    el.addEventListener('focus', () => { const r = el.getBoundingClientRect(); const t = tip(); showTip({ clientX: r.right, clientY: r.top }, t.title, t.rows, t.foot); });
    el.addEventListener('blur', hideTip);
  }
  if (activate) el.addEventListener('keydown', ev => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); activate(ev); } });
}

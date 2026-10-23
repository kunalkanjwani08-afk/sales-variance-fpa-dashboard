/* ===========================================================================
   BOOT
   =========================================================================== */
const TABS = [['overview', 'Overview', tabOverview], ['drivers', 'Drivers', tabDrivers], ['trends', 'Trends', tabTrends], ['regcat', 'Region × Category', tabRegCat], ['products', 'Products', tabProducts], ['exceptions', 'Exceptions & Lines', tabExceptions], ['insights', 'Insights & Recommendations', tabInsights], ['data', 'Data & Validation', tabData]];
let ACTIVE = 'overview';
function renderTabsNav() {
  const nav = $('#tabs'); nav.replaceChildren(...TABS.map(([id, label]) => h('button', { class: 'tab', role: 'tab', 'aria-selected': id === ACTIVE, onclick: () => { ACTIVE = id; renderTabsNav(); renderView(); window.scrollTo(0, Math.min(scrollY, $('#tabs').offsetTop - 60)); } }, label)));
}
function renderView() {
  const y = scrollY; pending = []; hideTip();
  const root = h('div'); const t = TABS.find(x => x[0] === ACTIVE);
  try { t[2](root); } catch (e) { console.error(e); root.append(h('div', { class: 'note' }, 'Something went wrong drawing this tab: ' + e.message)); }
  $('#view').replaceChildren(root); flushPending(); window.scrollTo(0, y);
  const c = $('#cnt'); if (c) c.textContent = `${numF(maskCount(getMask()))} of ${numF(N)} sales lines in view` + (isFiltered() ? '' : ' (no filters)');
}
function renderAll() { buildFilters(); renderView(); }
function exportLines() {
  const m = getMask(), head = ['Line_ID', 'Month', 'Region', 'Category', 'Product', 'Target_Qty', 'Actual_Qty', 'Target_Unit_Price_USD', 'Actual_Unit_Price_USD', 'Target_Unit_Cost_USD', 'Actual_Unit_Cost_USD',
    'Target_Sales_USD', 'Actual_Sales_USD', 'Sales_Variance_USD', 'Sales_Variance_Pct', 'Sales_Variance_Status', 'Material_Variance_Flag', 'Material_Trigger', 'Volume_Effect_USD', 'Price_Effect_USD',
    'Target_Gross_Profit_USD', 'Actual_Gross_Profit_USD', 'Gross_Profit_Variance_USD', 'Profit_Cost_Effect_USD', 'Notes'], rows = [];
  const c = v => (v / 100).toFixed(2);
  for (let i = 0; i < N; i++) { if (!m[i]) continue; rows.push([IDS[i], MON[A.M[i]], REG[A.R[i]], CAT[A.C[i]], PRD[A.P[i]], A.TQ[i], A.AQ[i], c(A.TP[i]), c(A.AP[i]), c(A.TCU[i]), c(A.ACU[i]), c(A.TS[i]), c(A.AS[i]), c(A.VAR[i]), pct2(A.VAR[i], A.TS[i]), ST_NAME[A.ST[i]], A.MT[i] ? 'Yes' : 'No', TRIG_NAME[A.TRG[i]], c(A.VOL[i]), c(A.PRC[i]), c(A.TG[i]), c(A.AG[i]), c(A.GVAR[i]), c(A.COST[i]), NOTES[A.NT[i]]]); }
  downloadCSV('sales_variance_filtered_lines.csv', head, rows);
}
function init() {
  let th = null; try { th = localStorage.getItem('svd-theme'); } catch (e) { } if (th) document.documentElement.setAttribute('data-theme', th);
  document.title = D.title; $('#ttl').textContent = D.title;
  $('#subttl').textContent = `${numF(N)} sales lines · ${NR} regions · ${NC} categories · ${NP} products · ${MLONG[0]} – ${MLONG[NM - 1]} · ${D.note}`;
  $('#foot').textContent = `Built ${D.built} from ${D.source}. All figures are computed in your browser in exact cents; nothing leaves your computer.`;
  $('#btnExport').addEventListener('click', exportLines); $('#btnPrint').addEventListener('click', () => window.print());
  $('#btnTheme').addEventListener('click', () => { setTheme(curTheme() === 'dark' ? 'light' : 'dark'); renderAll(); });
  RECON = reconcile(); const rb = $('#recon');
  if (!D.xl) { rb.textContent = 'Workbook not found: no reconciliation'; rb.className = 'badge note'; }
  else { const ok = RECON.filter(r => r.pass).length; rb.textContent = (ok === RECON.length ? '✔ ' : '✖ ') + `Reconciled to workbook: ${ok}/${RECON.length} checks`; rb.className = 'badge ' + (ok === RECON.length ? 'ok' : 'bad'); rb.style.cursor = 'pointer'; rb.title = 'Open the Data & Validation tab'; rb.addEventListener('click', () => { ACTIVE = 'data'; renderTabsNav(); renderView(); }); }
  renderTabsNav(); renderAll();
  let rt = null; addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(renderView, 160); });
  addEventListener('keydown', e => { if (e.key === 'Escape') hideTip(); });
  window.__dash = { S, getMask, totals, agg, dv, A, N, IDS, REG, CAT, PRD, MON, change, resetFilters, selectMonth, pickRegion, pickCat, pickProd, RECON, headlineClaims, liveInsights, setTab: t => { ACTIVE = t; renderTabsNav(); renderView(); }, renderAll, statusOf, matOf, pct2, persistenceRows, monthSeries, maskAll, discountSplit };
}
try { init(); } catch (e) { console.error(e); document.body.prepend(h('pre', { style: { color: 'red', padding: '20px', whiteSpace: 'pre-wrap' } }, 'Dashboard failed to start: ' + e.message + '\n' + e.stack)); }

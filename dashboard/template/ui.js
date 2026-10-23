/* ===========================================================================
   UI: filter bar, cards, tables, theme, export
   =========================================================================== */
const tableMode = new Map();          // card id -> show table instead of chart
let pending = [];                     // draw callbacks to run once the view is attached
function card(o) {                    // {id,title,sub,span,build(body),table(),tools[]}
  const id = o.id || o.title, body = h('div', { class: 'body' });
  const toolbar = h('div', { class: 'tools' });
  const el = h('section', { class: 'card ' + (o.span || 's6'), id: 'card-' + id.replace(/\W+/g, '-') },
    h('div', { class: 'card-h' }, h('div', null, h('h3', null, o.title), o.sub ? h('div', { class: 'sub' }, o.sub) : null), toolbar), body);
  const draw = () => { body.replaceChildren(); if (o.table && tableMode.get(id)) body.append(tableView(o.table())); else o.build(body); };
  if (o.tools) o.tools.forEach(t => toolbar.append(t));
  if (o.table) {
    const b = h('button', { class: 'tool', 'aria-pressed': !!tableMode.get(id), title: 'Show the numbers behind this chart', onclick: () => { tableMode.set(id, !tableMode.get(id)); b.setAttribute('aria-pressed', !!tableMode.get(id)); draw(); } }, '⊞ Table');
    toolbar.append(b);
  }
  pending.push(draw); return el;
}
function flushPending() { const p = pending; pending = []; p.forEach(f => { try { f(); } catch (e) { console.error(e); } }); }
function seg(options, current, onPick, aria) {
  return h('div', { class: 'seg', role: 'group', 'aria-label': aria || '' }, options.map(([v, l]) => h('button', { 'aria-pressed': v === current, onclick: () => onPick(v) }, l)));
}
function emptyMsg(msg) { return h('div', { class: 'empty' }, msg || 'No sales lines match the current filters. Use Reset to clear them.'); }

/* ---------- generic sortable / searchable / paged table ---------- */
function tableView(spec) {
  // spec: {cols:[{k,label,fmt,align,csv}], rows:[], pageSize, search:true, sortKey, sortDir, onRow(row), name}
  const st = { q: '', k: spec.sortKey || null, d: spec.sortDir || -1, page: 0 }, ps = spec.pageSize || 0;
  const root = h('div'), wrap = h('div', { class: 'tbl-wrap' }), pager = h('div', { class: 'pager' });
  const q = h('input', { type: 'search', placeholder: 'Search…', 'aria-label': 'Search table', oninput: e => { st.q = e.target.value.toLowerCase(); st.page = 0; draw(); } });
  const info = h('span', { class: 'sub' });
  const exp = h('button', { class: 'tool', onclick: () => downloadCSV((spec.name || 'table') + '.csv', spec.cols.map(c => c.label), view().map(r => spec.cols.map(c => c.csv ? c.csv(r) : r[c.k]))) }, '⬇ CSV');
  const bar = h('div', { class: 'tbar' }, spec.search === false ? null : q, info, exp);
  function view() {
    let rows = spec.rows;
    if (st.q) rows = rows.filter(r => spec.cols.some(c => String(c.csv ? c.csv(r) : r[c.k]).toLowerCase().includes(st.q)));
    if (st.k) { const c = spec.cols.find(c => c.k === st.k), g = c.sortVal || (r => r[c.k]); rows = rows.slice().sort((a, b) => { const x = g(a), y = g(b); const nx = x == null || (typeof x === 'number' && !isFinite(x)), ny = y == null || (typeof y === 'number' && !isFinite(y)); if (nx || ny) return nx === ny ? 0 : nx ? 1 : -1; return (x < y ? -1 : x > y ? 1 : 0) * st.d; }); }
    return rows;
  }
  function draw() {
    const rows = view(), tot = rows.length, pages = ps ? Math.max(1, Math.ceil(tot / ps)) : 1; st.page = Math.min(st.page, pages - 1);
    const slice = ps ? rows.slice(st.page * ps, st.page * ps + ps) : rows;
    const thead = h('thead', null, h('tr', null, spec.cols.map(c => h('th', { class: c.align === 'r' ? 'r' : '', onclick: () => { if (st.k === c.k) st.d = -st.d; else { st.k = c.k; st.d = c.align === 'r' ? -1 : 1; } draw(); }, title: 'Sort by ' + c.label }, c.label + (st.k === c.k ? (st.d > 0 ? ' ▲' : ' ▼') : '')))));
    const tb = h('tbody', null, slice.map(r => { const tr = h('tr', { class: spec.onRow ? 'clk' : '', onclick: spec.onRow ? () => spec.onRow(r) : null }, spec.cols.map(c => { const v = r[c.k], cell = c.fmt ? c.fmt(v, r) : v; return h('td', { class: c.align === 'r' ? 'r' : '' }, cell == null ? '' : cell); })); return tr; }));
    wrap.replaceChildren(h('table', { class: 't' }, thead, tb));
    info.textContent = `${numF(tot)} row${tot === 1 ? '' : 's'}` + (st.q ? ' (filtered)' : '');
    pager.replaceChildren();
    if (ps && pages > 1) pager.append(h('button', { class: 'tool', disabled: st.page === 0, onclick: () => { st.page--; draw(); } }, '‹ Prev'), `Page ${st.page + 1} of ${pages}`, h('button', { class: 'tool', disabled: st.page >= pages - 1, onclick: () => { st.page++; draw(); } }, 'Next ›'));
  }
  root.append(bar, wrap, pager); draw(); return root;
}
function downloadCSV(name, head, rows) {
  // Text that starts with = + - @ (and is not a plain number) is prefixed with ' so Excel never runs it as a formula.
  const q = v => {
    v = v == null ? '' : String(v);
    if (/^[=+\-@\t\r]/.test(v) && !/^[-+]?\d+(\.\d+)?$/.test(v)) v = "'" + v;
    return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  };
  const text = [head.map(q).join(','), ...rows.map(r => r.map(q).join(','))].join('\r\n');
  const blob = new Blob(['\uFEFF' + text], { type: 'text/csv;charset=utf-8' }), a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}

/* ---------- filter helpers ---------- */
function toggleSet(set, v) { set.has(v) ? set.delete(v) : set.add(v); }
function change(fn) { fn(); if (S.m0 > S.m1) [S.m0, S.m1] = [S.m1, S.m0]; bump(); renderAll(); }
function resetFilters() { change(() => { S.years.clear(); S.qtrs.clear(); S.regs.clear(); S.cats.clear(); S.prods.clear(); S.stat.clear(); S.mat = 'all'; S.dq = 'all'; S.m0 = 0; S.m1 = NM - 1; }); }
function selectMonth(m) { change(() => { S.years.clear(); S.qtrs.clear(); if (S.m0 === m && S.m1 === m) { S.m0 = 0; S.m1 = NM - 1; } else { S.m0 = S.m1 = m; } }); }
function pickRegion(r, ev) { change(() => { if (ev && (ev.shiftKey || ev.ctrlKey || ev.metaKey)) toggleSet(S.regs, r); else if (S.regs.size === 1 && S.regs.has(r)) S.regs.clear(); else { S.regs.clear(); S.regs.add(r); } }); }
function pickCat(c, ev) { change(() => { if (ev && (ev.shiftKey || ev.ctrlKey || ev.metaKey)) toggleSet(S.cats, c); else if (S.cats.size === 1 && S.cats.has(c)) S.cats.clear(); else { S.cats.clear(); S.cats.add(c); } }); }
function pickProd(p, ev) { change(() => { if (ev && (ev.shiftKey || ev.ctrlKey || ev.metaKey)) toggleSet(S.prods, p); else if (S.prods.size === 1 && S.prods.has(p)) S.prods.clear(); else { S.prods.clear(); S.prods.add(p); } }); }
function isFiltered() { return S.years.size || S.qtrs.size || S.regs.size || S.cats.size || S.prods.size || S.stat.size || S.mat !== 'all' || S.dq !== 'all' || S.m0 !== 0 || S.m1 !== NM - 1; }

let FILT_OPEN = innerWidth >= 760;           // on phones the filters start collapsed so the numbers are visible first
function buildFilters() {
  const root = $('#filters'); root.replaceChildren();
  const chip = (label, on, fn, dot) => h('button', { class: 'chip', 'aria-pressed': on, onclick: fn }, dot ? h('span', { class: 'dot', style: { background: dot } }) : null, label);
  const fg = (lab, ...kids) => h('div', { class: 'fg' }, h('div', { class: 'lbl' }, lab), ...kids);
  const monthSel = (which) => { const s = h('select', { class: 'sel', 'aria-label': which === 0 ? 'From month' : 'To month', onchange: e => change(() => { which === 0 ? S.m0 = +e.target.value : S.m1 = +e.target.value; S.years.clear(); S.qtrs.clear(); }) }, MON.map((m, i) => h('option', { value: i }, MLONG[i]))); s.value = which === 0 ? S.m0 : S.m1; return s; };
  const prodSel = h('select', { class: 'sel', 'aria-label': 'Product', onchange: e => change(() => { S.prods.clear(); if (e.target.value !== '') S.prods.add(+e.target.value); }) }, h('option', { value: '' }, 'All products (' + NP + ')'));
  CAT.forEach((c, ci) => { const og = h('optgroup', { label: c }); PRD.forEach((p, pi) => { if (DIM.prodCat[pi] === ci) og.append(h('option', { value: pi }, p)); }); prodSel.append(og); });
  prodSel.value = S.prods.size === 1 ? [...S.prods][0] : '';
  if (S.prods.size > 1) prodSel.firstChild.textContent = S.prods.size + ' products selected';
  const frow = h('div', { class: 'frow' },
    fg('Year', h('div', { class: 'chips' }, YEARS.map(y => chip(String(y), S.years.has(y), () => change(() => { toggleSet(S.years, y); S.m0 = 0; S.m1 = NM - 1; }))))),
    fg('Quarter', h('div', { class: 'chips' }, [1, 2, 3, 4].map(q => chip('Q' + q, S.qtrs.has(q), () => change(() => { toggleSet(S.qtrs, q); S.m0 = 0; S.m1 = NM - 1; }))))),
    fg('From', monthSel(0)), fg('To', monthSel(1)),
    fg('Region', h('div', { class: 'chips' }, REG.map((r, i) => chip(r, S.regs.has(i), () => change(() => toggleSet(S.regs, i)), `var(--s${i + 1})`)))),
    fg('Category', h('div', { class: 'chips' }, CAT.map((c, i) => chip(c, S.cats.has(i), () => change(() => toggleSet(S.cats, i)), `var(--s${i + 1})`)))),
    fg('Product', prodSel),
    fg('Line status', h('div', { class: 'chips' }, [0, 1, 2].map(s => chip(ST_GLYPH[s] + ' ' + ST_NAME[s], S.stat.has(s), () => change(() => toggleSet(S.stat, s)))))),
    fg('Material', seg([['all', 'All'], ['yes', 'Material'], ['no', 'Not material']], S.mat, v => change(() => { S.mat = v; }), 'Material filter')),
    fg('Data quality', seg([['all', 'All lines'], ['repaired', 'Repaired values'], ['clean', 'Untouched']], S.dq, v => change(() => { S.dq = v; }), 'Data quality filter')),
    h('div', { class: 'fg' }, h('div', { class: 'lbl' }, ' '), h('button', { class: 'btn', onclick: resetFilters, disabled: !isFiltered(), title: 'Clear every filter' }, '↺ Reset')));
  root.classList.toggle('open', FILT_OPEN); if (FILT_OPEN) root.append(frow);
  // active filter tags + live count
  const tags = h('div', { class: 'active' });
  const tag = (label, fn) => tags.append(h('span', { class: 'tag' }, label, h('button', { 'aria-label': 'Remove ' + label, onclick: fn }, '×')));
  S.years.forEach(y => tag('Year ' + y, () => change(() => S.years.delete(y)))); S.qtrs.forEach(q => tag('Q' + q, () => change(() => S.qtrs.delete(q))));
  if (S.m0 !== 0 || S.m1 !== NM - 1) tag(S.m0 === S.m1 ? MLONG[S.m0] : MLONG[S.m0] + ' → ' + MLONG[S.m1], () => change(() => { S.m0 = 0; S.m1 = NM - 1; }));
  S.regs.forEach(r => tag(REG[r], () => change(() => S.regs.delete(r)))); S.cats.forEach(c => tag(CAT[c], () => change(() => S.cats.delete(c)))); S.prods.forEach(p => tag(PRD[p], () => change(() => S.prods.delete(p))));
  S.stat.forEach(s => tag('Lines: ' + ST_NAME[s], () => change(() => S.stat.delete(s)))); if (S.mat !== 'all') tag(S.mat === 'yes' ? 'Material lines only' : 'Non-material lines only', () => change(() => { S.mat = 'all'; }));
  if (S.dq !== 'all') tag(S.dq === 'repaired' ? 'Lines with a repaired value' : 'Untouched lines', () => change(() => { S.dq = 'all'; }));
  tags.prepend(h('button', { class: 'ftoggle', 'aria-expanded': FILT_OPEN, title: FILT_OPEN ? 'Collapse the filters into a compact bar that stays on screen' : 'Show all filters', onclick: () => { FILT_OPEN = !FILT_OPEN; buildFilters(); } }, FILT_OPEN ? '▴ Collapse filters' : '▾ Show filters'));
  if (!FILT_OPEN && isFiltered()) tags.append(h('button', { class: 'ftoggle', onclick: resetFilters }, '↺ Clear all'));
  const cnt = maskCount(getMask());
  tags.append(h('span', { style: { marginLeft: 'auto' }, id: 'cnt' }, `${numF(cnt)} of ${numF(N)} sales lines in view` + (isFiltered() ? '' : ' (no filters)')));
  root.append(tags);
}

/* ---------- theme ---------- */
function setTheme(t) { if (t) document.documentElement.setAttribute('data-theme', t); else document.documentElement.removeAttribute('data-theme'); try { localStorage.setItem('svd-theme', t || ''); } catch (e) { } }
function curTheme() { return document.documentElement.getAttribute('data-theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'); }

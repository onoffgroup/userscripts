// ==UserScript==
// @name         Zoho Leads: Lead Type Counter (Fax) with Created Time filter (Lyte)
// @description  Counts Leads per Fax; filters by Created Time (All/This week/Last week); Lyte-aware; robust pagination and render waits.
// @namespace    https://on-off.group/
// @version      1.7.0
// @match        https://crm.zoho.com/*
// @match        https://crm.zoho.eu/*
// @match        https://crm.zoho.in/*
// @match        https://crm.zoho.com.au/*
// @match        https://*.zoho.com/crm/*
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // === CONFIG ===
  const LEAD_TYPE_HEADER = 'Fax';
  const CREATED_HEADER   = 'Created Time';
  const NEXT_BUTTON_SELECTOR = 'div.lyteSingleFront.lyteIconSingleFront[role="button"][aria-label="next"]';

  // Week starts on Monday (0=Sunday,1=Monday...). Change to 0 for Sunday-start weeks if you prefer.
  const WEEK_START_DAY = 1;

  // Render-wait tuning
  const RENDER_STABILITY_SAMPLES = 2;      // require this many consecutive equal signatures
  const RENDER_MAX_WAIT_MS = 20000;        // max wait for page to settle
  const POST_CHANGE_DELAY_MS = 500;        // small delay after change detected
  const MIN_NONBLANK_RATIO = 0.05;         // 5% non-blank or at least 1 non-blank

  const log = (...a)=>console.log('[OOG-FaxCounter]', ...a);

  // === UI ===
  const ui = document.createElement('div');
  Object.assign(ui.style, {
    position:'fixed', right:'16px', bottom:'16px', zIndex:999999,
    background:'#fff', border:'1px solid #ddd', borderRadius:'12px',
    padding:'12px', boxShadow:'0 6px 24px rgba(0,0,0,.12)', font:'14px system-ui',
    minWidth:'380px'
  });
  ui.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap">
      <strong>Lead Type Summary (Fax)</strong>
      <label style="font-size:13px;display:flex;align-items:center;gap:6px;">
        Created:
        <select id="oog_range" style="padding:4px 6px;border:1px solid #bbb;border-radius:6px;">
          <option value="this">This week</option>
          <option value="last">Last week</option>
          <option value="all">All</option>
        </select>
      </label>
    </div>
    <div style="display:flex;gap:8px;margin:8px 0;flex-wrap:wrap;">
      <button id="oog_page">Scan This Page</button>
      <button id="oog_all">Scan All Pages</button>
      <button id="oog_reset">Reset</button>
      <button id="oog_csv" disabled>Download CSV</button>
    </div>
    <div id="oog_status">Idle</div>
    <div id="oog_results" style="max-height:260px;overflow:auto;border-top:1px solid #eee;padding-top:8px;margin-top:6px;"></div>
  `;
  document.body.appendChild(ui);
  const statusEl = ui.querySelector('#oog_status');
  const resultsEl= ui.querySelector('#oog_results');
  const btnPage  = ui.querySelector('#oog_page');
  const btnAll   = ui.querySelector('#oog_all');
  const btnReset = ui.querySelector('#oog_reset');
  const btnCSV   = ui.querySelector('#oog_csv');
  const rangeSel = ui.querySelector('#oog_range');

  let counts = {}, total = 0;

  function render(){
    const rows = Object.entries(counts).sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>`${k||'(blank)'}: <b>${v}</b>`).join('<br>');
    resultsEl.innerHTML = `Total (filtered): <b>${total}</b><br>${rows || '<span style="color:#777;">No data yet.</span>'}`;
    btnCSV.disabled = total===0;
  }
  function reset(){ counts={}; total=0; render(); statusEl.textContent='Idle'; }

  function toCSV(){
    const esc=s=>`"${String(s).replace(/"/g,'""')}"`;
    const body = Object.entries(counts).map(([k,v])=>[k||'',v]);
    // Include selected range in filename for clarity
    const file = `lead_types_by_fax_${rangeSel.value||'all'}.csv`;
    return { name:file, data: [['LeadType(Fax)','Count'], ...body].map(r=>r.map(esc).join(',')).join('\n') };
  }
  btnCSV.addEventListener('click', ()=>{
    const {name, data} = toCSV();
    const blob = new Blob([data], {type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a=document.createElement('a'); a.href=url; a.download=name;
    document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(url); a.remove();},100);
  });

  // === Date helpers ===
  function startOfWeek(d){
    const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0,0,0,0);
    const day = dt.getDay(); // 0=Sun..6=Sat
    const diff = (day - WEEK_START_DAY + 7) % 7;
    dt.setDate(dt.getDate() - diff);
    return dt;
  }
  function weekRange(which){
    const now = new Date();
    const sow = startOfWeek(now);
    if (which === 'last') {
      const lastStart = new Date(sow); lastStart.setDate(sow.getDate() - 7);
      const lastEnd = new Date(lastStart); lastEnd.setDate(lastStart.getDate() + 7); lastEnd.setMilliseconds(-1);
      return {start:lastStart, end:lastEnd};
    }
    if (which === 'this') {
      const thisEnd = new Date(sow); thisEnd.setDate(sow.getDate() + 7); thisEnd.setMilliseconds(-1);
      return {start:sow, end:thisEnd};
    }
    return null; // 'all'
  }
  function parseDateLoose(str){
    if (!str) return null;
    // Try native parse first
    const d1 = new Date(str);
    if (!isNaN(d1)) return d1;
    // Common alternate formats cleanup
    const s = str.replace(/(\d{1,2})(st|nd|rd|th)/gi, '$1'); // 1st -> 1
    const d2 = new Date(s);
    if (!isNaN(d2)) return d2;
    return null;
  }

  function isInSelectedWeek(createdStr){
    const mode = rangeSel.value || 'this';
    if (mode === 'all') return true;
    const r = weekRange(mode);
    const d = parseDateLoose(createdStr);
    if (!d) return false;
    return d >= r.start && d <= r.end;
  }

  // === Lyte-aware DOM helpers ===
  function findLyteHeaderByLabel(label){
    // Try data-zcqa exact match first
    let hdr = document.querySelector(`lyte-exptable-th[data-zcqa="${CSS.escape(label)}"]`);
    if (hdr) return hdr;
    // Fallback to text match
    const headers = [...document.querySelectorAll('lyte-exptable-th')];
    hdr = headers.find(h => (h.innerText||'').trim().toLowerCase() === label.toLowerCase());
    if (!hdr) throw new Error(`Lyte header "${label}" not found. Ensure it is visible in the list view.`);
    return hdr;
  }

  function colZeroIdxFromHeader(hdr){
    const raw = hdr.getAttribute('cxcellcol');
    if (raw != null) {
      const n = parseInt(raw, 10);
      if (!Number.isNaN(n)) return n; // Lyte uses 0-based cxcellcol
    }
    const all = [...hdr.parentElement.querySelectorAll('lyte-exptable-th')];
    const i = all.indexOf(hdr);
    if (i < 0) throw new Error('Could not compute column index.');
    return i;
  }

  function getLyteRows(){
    // Prefer Lyte row tags
    let rows = [...document.querySelectorAll('lyte-exptable-tr')];
    if (rows.length) return rows;
    // Fallback: ARIA rows with cells
    rows = [...document.querySelectorAll('[role="row"]')].filter(r=>r.querySelector('[role="cell"]'));
    if (rows.length) return rows;
    // Fallback: table rows
    const tb = document.querySelector('tbody');
    return tb ? [...tb.querySelectorAll('tr')].filter(r=>r.querySelector('td')) : [];
  }

  function getCellFromRow(row, zeroIdx){
    // Lyte primary
    let cell = row.querySelector?.(`lyte-exptable-td[cxcellcol="${zeroIdx}"]`);
    if (cell) return cell;
    // Fallback nth-child
    cell = row.querySelector?.(`lyte-exptable-td:nth-child(${zeroIdx+1})`);
    if (cell) return cell;
    // ARIA fallback
    const aria = row.querySelector?.(`[role="cell"][aria-colindex="${zeroIdx+1}"]`);
    if (aria) return aria;
    // Table fallback
    const tds = row.querySelectorAll?.('td');
    return tds && tds[zeroIdx] ? tds[zeroIdx] : null;
  }

  function readLyteText(cell){
    if (!cell) return '';
    const lt = cell.querySelector?.('lyte-text[lt-prop-value]');
    if (lt) {
      const v = lt.getAttribute('lt-prop-value') || '';
      if (v.trim()) return v.trim();
      const t = (lt.textContent||'').trim();
      if (t) return t;
    }
    const anyLt = cell.querySelector?.('[lt-prop-value]');
    if (anyLt) {
      const v2 = anyLt.getAttribute('lt-prop-value');
      if (v2 && v2.trim()) return v2.trim();
    }
    const txt=(cell.innerText||cell.textContent||'').trim();
    if (txt) return txt;
    const tit=cell.getAttribute?.('title') || cell.querySelector?.('[title]')?.getAttribute('title') || '';
    return (tit||'').trim();
  }

  function pageSignature(colIdx){
    // Use Fax column cells for signature
    const rows = getLyteRows();
    if (!rows.length) return '';
    const first = readLyteText(getCellFromRow(rows[0], colIdx)).slice(0,120);
    const last  = readLyteText(getCellFromRow(rows[rows.length-1], colIdx)).slice(0,120);
    return `${first}|${last}|${rows.length}`;
  }

  function countThisPageFiltered(){
    const hdrFax = findLyteHeaderByLabel(LEAD_TYPE_HEADER);
    const hdrCrt = findLyteHeaderByLabel(CREATED_HEADER);
    const faxIdx = colZeroIdxFromHeader(hdrFax);
    const crtIdx = colZeroIdxFromHeader(hdrCrt);

    const rows = getLyteRows();
    log('Rows on page:', rows.length, 'faxIdx=', faxIdx, 'createdIdx=', crtIdx);

    let added = 0;
    rows.forEach(r=>{
      const faxCell = getCellFromRow(r, faxIdx);
      const crtCell = getCellFromRow(r, crtIdx);
      const faxVal  = readLyteText(faxCell);
      const crtVal  = readLyteText(crtCell);
      // Filter by Created Time
      if (isInSelectedWeek(crtVal)) {
        const key = faxVal || '(blank)';
        counts[key] = (counts[key]||0) + 1;
        total++; added++;
      }
    });
    return { faxIdx, added };
  }

  // === Robust change + render-stability wait ===
  async function waitForChange(prevURL, prevSig, colIdx, max=20000){
    const t0=performance.now();
    while (performance.now()-t0<max){
      await new Promise(r=>setTimeout(r,250));
      const urlChanged = location.href !== prevURL;
      const sigChanged = (() => {
        try { const s = pageSignature(colIdx); return s && s !== prevSig; }
        catch { return false; }
      })();
      if (urlChanged || sigChanged) {
        log('Change detected',{urlChanged,sigChanged});
        return true;
      }
    }
    log('Timeout waiting for change.');
    return false;
  }

  async function waitForRenderStable(colIdx, max=RENDER_MAX_WAIT_MS){
    const start = performance.now();
    let prev = '';
    let stableCount = 0;
    while (performance.now()-start < max){
      const rows = getLyteRows();
      if (rows.length) {
        const sig = pageSignature(colIdx);
        const firstRowFax = readLyteText(getCellFromRow(rows[0], colIdx));
        const nonBlank = rows.reduce((a,r)=> a + (readLyteText(getCellFromRow(r,colIdx)) ? 1 : 0), 0);
        const ratio = rows.length ? nonBlank/rows.length : 0;

        if (sig === prev) stableCount++; else stableCount = 0;
        prev = sig;

        if (stableCount >= (RENDER_STABILITY_SAMPLES-1) &&
            (nonBlank >= 1 || ratio >= MIN_NONBLANK_RATIO)) {
          log('Render stable:', {rows: rows.length, nonBlank, ratio: ratio.toFixed(2), firstRowFax});
          return true;
        }
      }
      await new Promise(r=>setTimeout(r,250));
    }
    log('Render stability timeout.');
    return false;
  }

  function findNextBtn(){
    const btn = document.querySelector(NEXT_BUTTON_SELECTOR);
    if (!btn) return { btn:null, disabled:true, reason:'not found' };
    const cs = getComputedStyle(btn);
    const disabled = btn.hasAttribute('disabled') || btn.getAttribute('aria-disabled')==='true' ||
                     /disabled/i.test(btn.className) || cs.pointerEvents==='none' || btn.offsetParent===null;
    return { btn, disabled, reason: disabled ? 'disabled/hidden' : 'ok' };
  }
async function goToFirstPage(faxIdx) {
  const btn = document.querySelector('div[role="button"][aria-label="first"]');
  if (!btn) { log('No First-page button found'); return; }
  const cs = getComputedStyle(btn);
  if (btn.hasAttribute('disabled') || cs.pointerEvents==='none' || btn.offsetParent===null) {
    log('Already at first page'); return;
  }
  log('Clicking First page…');
  btn.dispatchEvent(new MouseEvent('click', { bubbles:true, cancelable:true, view:window }));
  await new Promise(r=>setTimeout(r, POST_CHANGE_DELAY_MS));
  await waitForRenderStable(faxIdx, RENDER_MAX_WAIT_MS);
  log('Now at first page');
}

  // === Actions ===
  function renderAndStatus(added, prefix=''){
    render();
    const label = rangeSel.value === 'this' ? 'this week' :
                  rangeSel.value === 'last' ? 'last week' : 'all time';
    statusEl.textContent = added != null
      ? `${prefix}Counted ${added} leads on this page (${label}). Total: ${total}`
      : `${prefix}Total (${label}): ${total}`;
  }

  btnReset.addEventListener('click', reset);

  btnPage.addEventListener('click', ()=>{
    try{
      const r = countThisPageFiltered(); renderAndStatus(r.added);
    }catch(e){ statusEl.textContent='Error: '+e.message; log(e); }
  });

  btnAll.addEventListener('click', async ()=>{
    try{
      reset();
    statusEl.textContent='Jumping to page 1…';
    let faxIdx;
    try { faxIdx = colZeroIdxFromHeader(findLyteHeaderByLabel(LEAD_TYPE_HEADER)); } catch {}
    await goToFirstPage(faxIdx);

    statusEl.textContent='Scanning page 1…';
    let first = countThisPageFiltered(); renderAndStatus(first.added);
    faxIdx = first.faxIdx;

      let guard=500;
      while(guard-- > 0){
        const prevURL = location.href;
        const prevSig = pageSignature(faxIdx);

        const { btn, disabled, reason } = findNextBtn();
        log('Next state:', {disabled, reason});
        if (!btn || disabled){ statusEl.textContent='End reached (Next unavailable). Done.'; break; }

        log('Clicking Next…');
        btn.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,view:window}));

        // 1) Wait for URL/DOM change
        const changed = await waitForChange(prevURL, prevSig, faxIdx, 20000);
        if (!changed){ statusEl.textContent='No change after Next; assuming end. Done.'; break; }

        // 2) Small post-change delay
        await new Promise(r=>setTimeout(r, POST_CHANGE_DELAY_MS));

        // 3) Re-read Fax header (user might reorder columns) and wait for render to settle
        try { faxIdx = colZeroIdxFromHeader(findLyteHeaderByLabel(LEAD_TYPE_HEADER)); } catch {}
        await waitForRenderStable(faxIdx, RENDER_MAX_WAIT_MS);

        const before = total;
        const r = countThisPageFiltered(); renderAndStatus(null, '');
        if (total===before || r.added===0){ statusEl.textContent='No new rows (filtered); end reached. Done.'; break; }

        statusEl.textContent=`Count so far: ${total}`;
      }
      statusEl.textContent=`Done. Total: ${total}`;
    }catch(e){ statusEl.textContent='Error: '+e.message; log(e); }
  });

  if (!location.href.includes('/tab/Leads')) {
    statusEl.textContent = 'Navigate to Leads → List view to use.';
  }
})();

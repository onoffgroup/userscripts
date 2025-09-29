// ==UserScript==
// @name         Zoho Leads: Lead Type Counter (Fax) with Created Time filter (Lyte)
// @description  Counts Leads per Fax; filters by Created Time (All/This week/Last week); Lyte-aware; robust pagination and render waits.
// @namespace    https://on-off.group/
// @version      1.7.1
// @match        https://crm.zoho.com/*
// @match        https://crm.zoho.eu/*
// @match        https://crm.zoho.in/*
// @match        https://crm.zoho.com.au/*
// @match        https://*.zoho.com/crm/*
// @grant        none
// @downloadURL  https://raw.githubusercontent.com/onoffgroup/userscripts/main/zohoreport.user.js
// @updateURL    https://raw.githubusercontent.com/onoffgroup/userscripts/main/zohoreport.user.js
// @homepageURL  https://github.com/onoffgroup/userscripts
// @supportURL   https://github.com/onoffgroup/userscripts/issues
// ==/UserScript==

(function () {
  'use strict';

  // === CONFIG ===
  const LEAD_TYPE_HEADER = 'Fax';
  const CREATED_HEADER   = 'Created Time';
  const NEXT_BUTTON_SELECTOR = 'div.lyteSingleFront.lyteIconSingleFront[role="button"][aria-label="next" i]';
  const FIRST_PAGE_QUERY_VALUE = '1';
  const PER_PAGE_QUERY_VALUE = '10';

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
      <strong>Lead Type Summary</strong>
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
      <button id="oog_all">Count Leads</button>
      <button id="oog_reset">Reset</button>
      <button id="oog_csv" disabled>Download CSV</button>
    </div>
    <div id="oog_status">Idle</div>
    <div id="oog_results" style="max-height:260px;overflow:auto;border-top:1px solid #eee;padding-top:8px;margin-top:6px;"></div>
  `;
  document.body.appendChild(ui);
  const statusEl = ui.querySelector('#oog_status');
  const resultsEl= ui.querySelector('#oog_results');
  const btnAll   = ui.querySelector('#oog_all');
  const btnReset = ui.querySelector('#oog_reset');
  const btnCSV   = ui.querySelector('#oog_csv');
  const rangeSel = ui.querySelector('#oog_range');

  let counts = {}, total = 0;
  const seenRowSignatures = new Set();
  const seenPageSignatures = new Set();
  const seenPageNumbers = new Set();

  function render(){
    const rows = Object.entries(counts).sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>`${k||'(blank)'}: <b>${v}</b>`).join('<br>');
    resultsEl.innerHTML = `Total (filtered): <b>${total}</b><br>${rows || '<span style="color:#777;">No data yet.</span>'}`;
    btnCSV.disabled = total===0;
  }
  function reset(){ counts={}; total=0; seenRowSignatures.clear(); seenPageSignatures.clear(); seenPageNumbers.clear(); render(); statusEl.textContent='Idle'; }

  function getCurrentPageNumber(){
    try {
      const url = new URL(location.href);
      const raw = url.searchParams.get('page');
      if (raw != null && raw !== '') {
        const num = parseInt(raw, 10);
        if (!Number.isNaN(num)) return num;
      }
    } catch {}
    const pagerActive = document.querySelector('[data-zcqa="pager"] .active, lyte-pagination .active, .lytePagination .active');
    if (pagerActive){
      const text = (pagerActive.innerText||pagerActive.textContent||'').trim();
      const num = parseInt(text,10);
      if (!Number.isNaN(num)) return num;
    }
    return null;
  }

  function ensureFirstPageURL(){
    let url;
    try { url = new URL(location.href); }
    catch { return false; }
    if (!/\/list\b/.test(url.pathname)) return false;
    let changed = false;
    if (url.searchParams.get('page') !== FIRST_PAGE_QUERY_VALUE){
      url.searchParams.set('page', FIRST_PAGE_QUERY_VALUE);
      changed = true;
    }
    if (url.searchParams.get('per_page') !== PER_PAGE_QUERY_VALUE){
      url.searchParams.set('per_page', PER_PAGE_QUERY_VALUE);
      changed = true;
    }
    if (!changed) return false;
    const newUrl = url.toString();
    log('Navigating to enforced first page URL', newUrl);
    location.assign(newUrl);
    return true;
  }

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
      const lastStart = new Date(sow);
      lastStart.setDate(lastStart.getDate() - 7);
      lastStart.setHours(0,0,0,0);
      const lastEnd = new Date(lastStart.getTime() + 7*24*60*60*1000 - 1);
      return {start:lastStart, end:lastEnd};
    }
    if (which === 'this') {
      const start = new Date(sow);
      const end = new Date(now);
      return {start, end};
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

  function isRowCountable(row){
    if (!row) return false;
    if (row.getAttribute?.('aria-hidden') === 'true') return false;
    const hasCells = row.querySelector?.('lyte-exptable-td, [role="cell"], td');
    if (!hasCells) return false;
    if (row.offsetParent !== null) return true;
    const cs = getComputedStyle(row);
    if (cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity||'1') === 0) return false;
    return true;
  }

  function getCountableRows(){
    return getLyteRows().filter(isRowCountable);
  }

  function rowSignature(row){
    if (!row) return '';
    const attrCandidates = ['data-id','data-value','data-rowid','rowid','data-rid','data-lyte-id','cxeid','id'];
    for (const attr of attrCandidates){
      const v = row.getAttribute?.(attr);
      if (v) return `${attr}:${v}`;
    }
    const anchor = row.querySelector?.('a[href*="/crm/"]');
    if (anchor){
      const href = anchor.getAttribute('href') || '';
      if (href) return `href:${href}`;
    }
    const cells = row.querySelectorAll?.('lyte-exptable-td, [role="cell"], td') || [];
    const textSig = [...cells].map(c=>readLyteText(c)).join('|');
    return `text:${textSig}`;
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
    const rows = getCountableRows();
    if (!rows.length) return '';
    const firstRowSig = rowSignature(rows[0]);
    const lastRowSig  = rowSignature(rows[rows.length-1]);
    const midRowSig   = rows.length > 2 ? rowSignature(rows[Math.floor(rows.length/2)]) : '';
    const firstColVal = colIdx != null ? readLyteText(getCellFromRow(rows[0], colIdx)).slice(0,120) : '';
    return `${rows.length}|${firstRowSig}|${midRowSig}|${lastRowSig}|${firstColVal}`;
  }

  function countThisPageFiltered(){
    const hdrFax = findLyteHeaderByLabel(LEAD_TYPE_HEADER);
    const hdrCrt = findLyteHeaderByLabel(CREATED_HEADER);
    const faxIdx = colZeroIdxFromHeader(hdrFax);
    const crtIdx = colZeroIdxFromHeader(hdrCrt);

    const rows = getCountableRows();
    let duplicates = 0;
    let skipped = 0;
    log('Rows on page:', rows.length, 'faxIdx=', faxIdx, 'createdIdx=', crtIdx);

    let added = 0;
    rows.forEach(r=>{
      const sig = rowSignature(r);
      if (!sig) { skipped++; return; }
      if (seenRowSignatures.has(sig)) { duplicates++; return; }
      const faxCell = getCellFromRow(r, faxIdx);
      const crtCell = getCellFromRow(r, crtIdx);
      const faxVal  = readLyteText(faxCell);
      const crtVal  = readLyteText(crtCell);
      // Filter by Created Time
      if (isInSelectedWeek(crtVal)) {
        seenRowSignatures.add(sig);
        const key = faxVal || '(blank)';
        counts[key] = (counts[key]||0) + 1;
        total++; added++;
      } else {
        seenRowSignatures.add(sig); // Mark as seen even if filtered out to avoid recount on later pages
      }
    });
    if (duplicates || skipped) {
      log('Row filter stats:', {duplicates, skipped, seen: seenRowSignatures.size});
    }
    return { faxIdx, added, duplicates, skipped };
  }

  // === Robust change + render-stability wait ===
  async function waitForChange(prevURL, prevSig, colIdx, prevPageNum=null, max=20000){
    const t0=performance.now();
    log('waitForChange: watching for transition', { prevURL, prevSig, prevPageNum });
    while (performance.now()-t0<max){
      await new Promise(r=>setTimeout(r,250));
      const urlChanged = location.href !== prevURL;
      const sigChanged = (() => {
        try { const s = pageSignature(colIdx); return s && s !== prevSig; }
        catch { return false; }
      })();
      const pageNumChanged = (() => {
        if (prevPageNum == null) return false;
        try {
          const current = getCurrentPageNumber();
          return current != null && current !== prevPageNum;
        } catch {
          return false;
        }
      })();
      if (urlChanged || sigChanged || pageNumChanged) {
        log('waitForChange: detected',{urlChanged,sigChanged,pageNumChanged, elapsed: Math.round(performance.now()-t0)});
        return true;
      }
    }
    log('waitForChange: timeout', { waited: max });
    return false;
  }

  async function waitForRenderStable(colIdx, max=RENDER_MAX_WAIT_MS){
    const start = performance.now();
    let prev = '';
    let stableCount = 0;
    log('waitForRenderStable: watching', { colIdx, max });
    while (performance.now()-start < max){
      const rows = getCountableRows();
      if (rows.length) {
        const sig = pageSignature(colIdx);
        const firstRowFax = readLyteText(getCellFromRow(rows[0], colIdx));
        const nonBlank = rows.reduce((a,r)=> a + (readLyteText(getCellFromRow(r,colIdx)) ? 1 : 0), 0);
        const ratio = rows.length ? nonBlank/rows.length : 0;

        if (sig === prev) stableCount++; else stableCount = 0;
        prev = sig;

        if (stableCount >= (RENDER_STABILITY_SAMPLES-1) &&
            (nonBlank >= 1 || ratio >= MIN_NONBLANK_RATIO)) {
          log('waitForRenderStable: stable', {rows: rows.length, nonBlank, ratio: ratio.toFixed(2), firstRowFax, elapsed: Math.round(performance.now()-start)});
          return true;
        }
      }
      await new Promise(r=>setTimeout(r,250));
    }
    log('waitForRenderStable: timeout', { waited: max });
    return false;
  }

  function analyzePagerButton(btn){
    if (!btn) return { btn:null, disabled:true, reason:'not found' };
    const cs = getComputedStyle(btn);
    const hidden = btn.offsetParent===null || cs.display==='none' || cs.visibility==='hidden' || parseFloat(cs.opacity||'1')===0;
    const attrDisabled = btn.hasAttribute('disabled') || (btn.getAttribute('aria-disabled')||'').toLowerCase()==='true' || /disabled/i.test(btn.className);
    const pointerBlocked = cs.pointerEvents==='none';
    const tabindexBlocked = btn.getAttribute('tabindex') === '-1';
    const disabled = hidden || attrDisabled || pointerBlocked || tabindexBlocked;
    const reason = hidden ? 'hidden' : attrDisabled ? 'aria/attr disabled' : pointerBlocked ? 'pointer-events none' : tabindexBlocked ? 'tabindex -1' : 'ok';
    return { btn, disabled, reason };
  }

  function findNextBtn(){
    const candidates = [...document.querySelectorAll(NEXT_BUTTON_SELECTOR)];
    if (!candidates.length) return { btn:null, disabled:true, reason:'not found' };
    for (const cand of candidates){
      const info = analyzePagerButton(cand);
      if (!info.disabled) return info;
    }
    const fallback = analyzePagerButton(candidates[0]);
    return { ...fallback, reason: `all candidates disabled (${fallback.reason})` };
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

  btnAll.addEventListener('click', async ()=>{
    console.log('=== Zoho Leads Fax Counter started ===');
    try{
      reset();
    if (ensureFirstPageURL()){
      statusEl.textContent='Loading page 1 (10 per page)…';
      return;
    }
    statusEl.textContent='Jumping to page 1…';
    let faxIdx;
    try { faxIdx = colZeroIdxFromHeader(findLyteHeaderByLabel(LEAD_TYPE_HEADER)); } catch {}
    await goToFirstPage(faxIdx);

    statusEl.textContent='Scanning page 1…';
    let first = countThisPageFiltered(); renderAndStatus(first.added);
    faxIdx = first.faxIdx;
    try {
      const sig = pageSignature(faxIdx);
      if (sig) seenPageSignatures.add(sig);
    } catch {}
    const firstPageNum = getCurrentPageNumber();
    if (firstPageNum != null) seenPageNumbers.add(firstPageNum);

      let guard=500;
      while(guard-- > 0){
        log('--- Page loop iteration ---', {
          remainingGuard: guard,
          currentPage: getCurrentPageNumber(),
          totalSoFar: total
        });
        const prevURL = location.href;
        const prevSig = pageSignature(faxIdx);

        const { btn, disabled, reason } = findNextBtn();
        log('Next state:', {disabled, reason});
        if (!btn || disabled){ statusEl.textContent='End reached (Next unavailable). Done.'; break; }

        log('Clicking Next…');
        btn.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,view:window}));

        // 1) Wait for URL/DOM change
  const prevPageNum = getCurrentPageNumber();
  const changed = await waitForChange(prevURL, prevSig, faxIdx, prevPageNum, 20000);
        if (!changed){
          statusEl.textContent='No change after Next; assuming end. Done.';
          log('Pagination aborted: change not detected', { prevURL, prevSig, currentURL: location.href });
          break;
        }

        // 2) Small post-change delay
        await new Promise(r=>setTimeout(r, POST_CHANGE_DELAY_MS));

        // 3) Re-read Fax header (user might reorder columns) and wait for render to settle
        try { faxIdx = colZeroIdxFromHeader(findLyteHeaderByLabel(LEAD_TYPE_HEADER)); } catch {}
        await waitForRenderStable(faxIdx, RENDER_MAX_WAIT_MS);

        let currentSig = '';
        try {
          currentSig = pageSignature(faxIdx);
        } catch (err) {
          log('Failed to compute page signature', err);
        }
        const currentPageNum = getCurrentPageNumber();
        const pageAlreadySeen = currentPageNum != null && seenPageNumbers.has(currentPageNum);
        if (currentPageNum != null) seenPageNumbers.add(currentPageNum);
        if (currentSig) seenPageSignatures.add(currentSig);
        if (pageAlreadySeen) {
          statusEl.textContent = 'Reached a previously visited page; stopping to avoid loop.';
          log('Stopping because page number already processed', { currentPageNum, seenPageNumbers: [...seenPageNumbers] });
          break;
        }

        const before = total;
        const r = countThisPageFiltered();
        renderAndStatus(r.added);
        log('Page processed', {
          currentPageNum,
          added: r.added,
          duplicates: r.duplicates,
          skipped: r.skipped,
          total
        });

        if (total===before && r.added===0) {
          statusEl.textContent = `Count so far: ${total} (no matching leads on this page)`;
          log('No filtered leads on this page; continuing.');
        } else {
          statusEl.textContent=`Count so far: ${total}`;
        }
      }
      statusEl.textContent=`Done. Total: ${total}`;
    }catch(e){ statusEl.textContent='Error: '+e.message; log(e); }
  });

  if (!location.href.includes('/tab/Leads')) {
    statusEl.textContent = 'Navigate to Leads → List view to use.';
  }
})();

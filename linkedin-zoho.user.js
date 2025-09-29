// ==UserScript==
// @name         LinkedIn → Zoho Lead (On-Off Group) — Stable Modal
// @namespace    onoffgroup.linkedin.zoho
// @version      1.2
// @description  Add a LinkedIn profile to Zoho CRM Leads with a quick category picker.
// @match        https://www.linkedin.com/in/*
// @match        https://crm.zoho.com/crm/org900416758/tab/Leads/create*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_openInTab
// @run-at       document-idle
// @downloadURL  https://raw.githubusercontent.com/onoffgroup/userscripts/main/linkedin-zoho.user.js
// @updateURL    https://raw.githubusercontent.com/onoffgroup/userscripts/main/linkedin-zoho.user.js
// @homepageURL  https://github.com/onoffgroup/userscripts
// @supportURL   https://github.com/onoffgroup/userscripts/issues
// ==/UserScript==

(function () {
  const ZOHO_CREATE_URL = 'https://crm.zoho.com/crm/org900416758/tab/Leads/create';
  const Z_MAX = 2147483647; // keep me on top of everything

  const CATEGORIES = [
    { label: 'B2B', value: 'B2B' },
    { label: 'B2C', value: 'B2C' },
    { label: 'Partner', value: 'Partner' },
    { label: 'Educational', value: 'Educational' },
    { label: 'Association', value: 'Association' },
  ];

  const STORAGE_KEY = 'on_off_lead_payload';

  const onLinkedIn = location.hostname.includes('linkedin.com');
  const onZoho = location.hostname.includes('crm.zoho.com');

  // ---------- Utilities ----------
  const $ = (sel, root = document) => root.querySelector(sel);
  const waitFor = (ms) => new Promise(r => setTimeout(r, ms));
function injectCSS(css){ const s=document.createElement('style'); s.textContent=css; document.head.appendChild(s); }






  function createEl(tag, attrs = {}, children = []) {
    const el = document.createElement(tag);
    Object.entries(attrs).forEach(([k, v]) => {
      if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v);
    });
    (Array.isArray(children) ? children : [children]).forEach(c => {
      if (typeof c === 'string') el.appendChild(document.createTextNode(c));
      else if (c) el.appendChild(c);
    });
    return el;
  }

  function extractLinkedInData() {
    // Name (try main H1, then sticky header)
    let fullName = (document.querySelector('main h1')?.textContent || '').trim();
    if (!fullName) {
      fullName = (document.querySelector('.artdeco-entity-lockup__title')?.textContent || '').trim();
    }
    let firstName = '', lastName = '';
    if (fullName) {
      const parts = fullName.split(/\s+/);
      firstName = parts.shift() || '';
      lastName = parts.join(' ');
    }

    // Company: aria-label or sticky subtitle “... | Company”
    let company = '';
    const currentCompanyBtn = document.querySelector('button[aria-label^="Current company"]');
    if (currentCompanyBtn?.getAttribute('aria-label')) {
      const label = currentCompanyBtn.getAttribute('aria-label');
      const m = label.match(/Current company:\s*([^.]*)/i);
      if (m) company = m[1].trim();
    }
    if (!company) {
      const subtitle = document.querySelector('.artdeco-entity-lockup__subtitle')?.textContent || '';
      if (subtitle.includes('|')) company = subtitle.split('|').pop().trim();
    }
    if (!company) {
      const experienceSection = (() => {
        const headingSpan = Array.from(document.querySelectorAll('h2.pvs-header__title span[aria-hidden="true"]'))
          .find(span => (span.textContent || '').trim().toLowerCase() === 'experience');
        if (!headingSpan) return null;
        return headingSpan.closest('section') || headingSpan.closest('.artdeco-card') || null;
      })();
      if (experienceSection) {
        const norm = (el) => {
          if (!el) return '';
          const visible = el.querySelector('[aria-hidden="true"]');
          const text = (visible?.textContent ?? el.textContent ?? '');
          return text.replace(/\s+/g, ' ').trim();
        };
        const boldNodes = Array.from(experienceSection.querySelectorAll('.t-bold'));
        const companyBold = boldNodes.find(node => node.closest('a[data-field="experience_company_logo"]'));
        const candidateFromLogo = norm(companyBold);
        if (candidateFromLogo) {
          company = candidateFromLogo;
        } else if (boldNodes.length) {
          const candidateFallback = norm(boldNodes[0]);
          if (candidateFallback) company = candidateFallback;
        }
      }
    }

    const profileUrl = location.href.split('?')[0];
    return { firstName, lastName, company, profileUrl };
  }

  // ---------- LinkedIn UI ----------
  function btnStyle(bg, color) {
    return {
      background: bg, color,
      border: 'none', borderRadius: '8px', padding: '8px 12px',
      fontWeight: '600', cursor: 'pointer'
    };
  }

function openCategoryModal() {
  // host + shadow
  const host = document.createElement('div');
  host.id = 'onoff-host';
  host.style.position = 'fixed';
  host.style.inset = '0';
  host.style.zIndex = '2147483647';
  document.body.appendChild(host);

  const shadow = host.attachShadow({ mode: 'open' });

  // <style>
  const style = document.createElement('style');
  style.textContent = `
    :host { all: initial; }
    .overlay {
      position: fixed; inset: 0;
      background: rgba(0,0,0,0.25);
    }
    .card {
      position: fixed;
      right: 20px;
      bottom: 96px;
      transform: none;
      width: min(170px, calc(100vw - 40px));
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 16px 48px rgba(0,0,0,0.25);
      padding: 18px;
      font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, "Apple Color Emoji","Segoe UI Emoji";
      color: #111;
    }
    .title { font-size: 16px; font-weight: 700; margin: 0 0 10px 0; }
    .row { display: flex; align-items: center; gap: 8px; margin: 6px 0; line-height: 1.2; cursor: pointer; }
    input[type="radio"] {
      appearance: auto; -webkit-appearance: auto;
      width: 18px; height: 18px; margin: 0;
      accent-color: #0A66C2;
      cursor: pointer;
    }
    .actions { display:flex; gap:8px; justify-content:flex-end; margin-top:14px; }
    button { border: none; border-radius: 8px; padding: 8px 12px; font-weight: 600; cursor: pointer; }
    .cancel { background:#E0E0E0; color:#111; }
    .add { background:#0A66C2; color:#fff; }
  `;
  shadow.appendChild(style);

  // overlay
  const overlay = document.createElement('div');
  overlay.className = 'overlay';
  shadow.appendChild(overlay);

  // card
  const card = document.createElement('div');
  card.className = 'card';
  shadow.appendChild(card);

  // title
  const title = document.createElement('div');
  title.className = 'title';
  title.textContent = 'Select a lead type';
  card.appendChild(title);

  // options
  const options = document.createElement('div');
  options.className = 'options';
  card.appendChild(options);

  // build radios
  CATEGORIES.forEach(cat => {
    const row = document.createElement('label');
    row.className = 'row';

    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'onoff-category';
    radio.value = cat.value;

    const span = document.createElement('span');
    span.textContent = cat.label;

    row.appendChild(radio);
    row.appendChild(span);
    options.appendChild(row);
  });

  // actions
  const actions = document.createElement('div');
  actions.className = 'actions';
  card.appendChild(actions);

  const cancel = document.createElement('button');
  cancel.type = 'button';
  cancel.className = 'cancel';
  cancel.textContent = 'Cancel';
  actions.appendChild(cancel);

  const add = document.createElement('button');
  add.type = 'button';
  add.className = 'add';
  add.textContent = 'Add Lead';
  actions.appendChild(add);

  // handlers
  const close = () => host.remove();
  overlay.addEventListener('click', close);
  cancel.addEventListener('click', close);

  add.addEventListener('click', () => {
    const sel = shadow.querySelector('input[name="onoff-category"]:checked');
    if (!sel) { alert('Please pick a lead type.'); return; }

    const li = extractLinkedInData();
    GM_setValue(STORAGE_KEY, {
      createdAt: Date.now(),
      firstName: li.firstName || '',
      lastName: li.lastName || '',
      company: li.company || '',
      website: li.profileUrl || '',
      fax: sel.value
    });

    if (typeof GM_openInTab === 'function') {
      GM_openInTab('https://crm.zoho.com/crm/org900416758/tab/Leads/create', { active: true, insert: true });
    } else {
      window.open('https://crm.zoho.com/crm/org900416758/tab/Leads/create', '_blank');
    }
    close();
  });
}
async function setLeadSourceTo(valueText = 'X (Twitter)') {
  const wait = (ms) => new Promise(r => setTimeout(r, ms));
  const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();

  const dd = document.querySelector('#Crm_Leads_LEADSOURCE');
  if (!dd) { console.warn('LeadSource dropdown not found'); return false; }

  const trigger = dd.querySelector('.lyteDummyEventContainer') || dd.querySelector('lyte-drop-button');
  const labelSpan = dd.querySelector('.lyteDropdownLabel');
  if (!trigger) { console.warn('LeadSource trigger not found'); return false; }

  // 1) Open
  trigger.focus();
  trigger.click();
  for (let i = 0; i < 30; i++) {
    const bodyId = trigger.getAttribute('aria-controls');
    const bodyEl = bodyId && document.getElementById(bodyId);
    if (bodyEl && bodyEl.offsetParent !== null) {
      // 2) Find the desired option inside THIS body
      let opt = bodyEl.querySelector(`lyte-drop-item[data-value="${CSS.escape(valueText)}"]`);
      if (!opt) {
        const all = Array.from(bodyEl.querySelectorAll('lyte-drop-item[role="option"]'));
        opt = all.find(el => norm(el.textContent) === valueText) ||
              all.find(el => norm(el.textContent).includes('Twitter'));
      }
      if (!opt) { console.warn('LeadSource option not found'); return false; }

      // 3) Highlight it by setting aria-activedescendant to its id
      const targetId = opt.id;
      if (targetId) {
        trigger.setAttribute('aria-activedescendant', targetId);
        // Some builds require moving the visual highlight:
        opt.classList.add('lyteDropdownSelection');

        // 4) Commit with Enter (preferred path)
        ['keydown','keyup'].forEach(type => {
          trigger.dispatchEvent(new KeyboardEvent(type, {
            key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
          }));
        });

        await wait(200);
        const ok = !!labelSpan && norm(labelSpan.textContent) === valueText;
        if (ok) {
          // Optional: close with Esc
          ['keydown','keyup'].forEach(type => {
            trigger.dispatchEvent(new KeyboardEvent(type, {
              key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true
            }));
          });
          return true;
        }
      }

      // 5) Fallback: set value via click on the lyte-drop-item’s internal click handler
      // (some tenants only accept events on the actual OPTION node)
      opt.scrollIntoView({ block: 'nearest' });
      opt.click(); // try native click
      await wait(200);
      if (labelSpan && norm(labelSpan.textContent) === valueText) return true;

      // 6) Last fallback: type-to-filter then Enter (supported path)
      const box = bodyEl.closest('lyte-drop-box');
      const searchInput = box?.querySelector('lyte-input input[type="text"]#LEADSOURCE');
      if (searchInput) {
        searchInput.focus();
        searchInput.value = 'Twitter';
        searchInput.dispatchEvent(new Event('input', { bubbles: true }));
        await wait(250);
        ['keydown','keyup'].forEach(type => {
          searchInput.dispatchEvent(new KeyboardEvent(type, {
            key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
          }));
        });
        await wait(200);
        if (labelSpan && /twitter/i.test(labelSpan.textContent)) return true;
      }

      console.warn('LeadSource not committed by UI fallbacks.');
      return false;
    }
    await wait(100);
  }
  console.warn('LeadSource dropdown body did not appear.');
  return false;
}


// --- helpers ---
function setNativeValue(el, value) {
  const proto = Object.getPrototypeOf(el);
  const desc = Object.getOwnPropertyDescriptor(proto, 'value');
  const setter = desc && desc.set;
  if (setter) setter.call(el, value);
  else el.value = value;
}
async function commitInput(selector, value) {
  const el = document.querySelector(selector);
  if (!el) return false;

  // focus → set via native setter → input/change → blur
  el.focus();
  setNativeValue(el, value);
  el.dispatchEvent(new Event('input',  { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
  el.blur();

  // tiny nudge for stubborn Lyte fields: type+backspace
  el.focus();
  el.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, inputType: 'insertText', data: ' ' }));
  setNativeValue(el, value + ' ');
  el.dispatchEvent(new Event('input', { bubbles: true }));
  setNativeValue(el, value);
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.blur();

  return true;
}



  function initLinkedIn() {
    const btn = createEl('button', {
      id: 'onoff-add-to-zoho',
      style: {
        position: 'fixed',
        bottom: '20px',
        right: '20px',
        zIndex: Z_MAX,
        padding: '10px 14px',
        borderRadius: '999px',
        border: 'none',
        cursor: 'pointer',
        boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
        background: '#0A66C2',
        color: '#fff',
        fontWeight: '600',
        pointerEvents: 'auto'
      }
    }, 'Add to Zoho');

    document.body.appendChild(btn);
    btn.addEventListener('click', openCategoryModal);
  }

  // ---------- Zoho side ----------
  async function initZoho() {
    const payload = GM_getValue(STORAGE_KEY);
    if (!payload) return;

    GM_deleteValue(STORAGE_KEY);

    // Wait for Zoho inputs to render
    for (let i = 0; i < 40; i++) {
      if (document.querySelector('#Crm_Leads_LASTNAME_LInput')) break;
      await waitFor(300);
    }

   await commitInput('#Crm_Leads_FIRSTNAME_LInput', payload.firstName || '');
await commitInput('#Crm_Leads_LASTNAME_LInput',  payload.lastName  || '');   // <- required field
// Company (Zoho often uses an inner text input under the label)
{
  const companyInput =
    document.querySelector('input[aria-labelledby="Crm_Leads_COMPANY_label"]') ||
    document.querySelector('#Crm_Leads_COMPANY input[type="text"]');
  if (companyInput) {
    // commit by element instead of selector
    companyInput.focus();
    setNativeValue(companyInput, payload.company || '');
    companyInput.dispatchEvent(new Event('input', { bubbles: true }));
    companyInput.dispatchEvent(new Event('change', { bubbles: true }));
    companyInput.blur();
  }
}
await commitInput('input[aria-labelledby="Crm_Leads_WEBSITE_label"]', payload.website || '');
await commitInput('#Crm_Leads_FAX_LInput', payload.fax || '');

try {
  const ok = await setLeadSourceTo('X (Twitter)');
  if (!ok) console.warn('[OnOff] Lead Source still not set (X/Twitter not found/selected).');
} catch (e) {
  console.warn('[OnOff] Lead Source set error', e);
}

  }

  // ---------- Boot ----------
  if (onLinkedIn) {
    initLinkedIn();
  } else if (onZoho) {
    initZoho();
  }
})();

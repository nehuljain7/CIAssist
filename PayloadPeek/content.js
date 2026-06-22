// content.js — runs in the CPI tab.
// * Fetches same-origin so the session cookie (JSESSIONID) is sent automatically.
// * Renders a DOCKABLE panel injected into the page (left / right / bottom) that
//   persists across tab switches and new-tab opens, keeping its data.

if (!window.__cpiAttachSearch) {
  window.__cpiAttachSearch = true;

  const ATOM = 'http://www.w3.org/2005/Atom';
  const D = 'http://schemas.microsoft.com/ado/2007/08/dataservices';

  let running = false;
  let results = [];
  let host = null;     // shadow host element
  let root = null;     // shadow root
  let dock = 'right';

  // chrome.storage.session is trusted-only by default; background.js opens it to us,
  // but wrap anyway so a storage hiccup can never be reported as a search error.
  const ss = {
    async get(k) { try { return await chrome.storage.session.get(k); } catch (_) { return {}; } },
    async set(o) { try { await chrome.storage.session.set(o); } catch (_) {} }
  };

  /* =========================================================
   *  ENGINE  (emits events to a callback)
   * ========================================================= */
  function dText(entry, tag) { const e = entry.getElementsByTagNameNS(D, tag); return e.length ? (e[0].textContent || '').trim() : ''; }
  function atomText(entry, tag) { const e = entry.getElementsByTagNameNS(ATOM, tag); return e.length ? (e[0].textContent || '').trim() : ''; }
  function absolute(u, origin) { return !u ? u : (/^https?:\/\//i.test(u) ? u : origin + (u.startsWith('/') ? '' : '/') + u); }

  async function fetchText(url) {
    const res = await fetch(url, { credentials: 'include', headers: { 'Accept': 'application/atom+xml,application/xml,*/*' } });
    if (res.status === 401 || res.status === 403) throw new Error(`Session not authorised (${res.status}). Make sure you are logged in on this tab.`);
    if (!res.ok) throw new Error(`HTTP ${res.status} on ${url}`);
    return res.text();
  }
  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Could not parse XML response.');
    return doc;
  }
  function toUtcOData(local) { return new Date(local).toISOString().slice(0, 19); }
  function buildFilter(p) {
    const c = [];
    if (p.from) c.push(`LogStart ge datetime'${toUtcOData(p.from)}'`);
    if (p.to) c.push(`LogStart le datetime'${toUtcOData(p.to)}'`);
    if (p.status && p.status !== 'ANY') c.push(`Status eq '${p.status}'`);
    if (p.iflow) c.push(`IntegrationFlowName eq '${p.iflow.replace(/'/g, "''")}'`);
    return c.join(' and ');
  }

  async function listMpls(origin, params, emit) {
    const filter = buildFilter(params);
    let url = `${origin}/odata/api/v1/MessageProcessingLogs?$orderby=LogStart desc&$top=500`
            + (filter ? `&$filter=${encodeURIComponent(filter)}` : '');
    const out = []; const seen = new Set(); let page = 0;
    while (url) {
      const doc = parseXml(await fetchText(url));
      const entries = Array.from(doc.getElementsByTagNameNS(ATOM, 'entry'));
      page++;
      for (const e of entries) {
        const guid = dText(e, 'MessageGuid') || (atomText(e, 'id').match(/MessageProcessingLogs\('([^']+)'\)/) || [])[1];
        if (!guid || seen.has(guid)) continue;   // skip duplicate GUIDs (paging ties)
        seen.add(guid);
        out.push({ guid, logStart: dText(e, 'LogStart'), logEnd: dText(e, 'LogEnd'), status: dText(e, 'Status'), iflow: dText(e, 'IntegrationFlowName'), appMsgId: dText(e, 'ApplicationMessageId') });
      }
      emit({ type: 'progress', message: `Listed page ${page} — ${out.length} messages so far` });
      let next = null;
      for (const l of doc.getElementsByTagNameNS(ATOM, 'link')) { if (l.getAttribute('rel') === 'next') { next = l.getAttribute('href'); break; } }
      url = next ? absolute(next, origin) : null;
    }
    return out;
  }

  async function attachmentsFor(origin, guid) {
    const url = `${origin}/odata/api/v1/MessageProcessingLogs('${guid}')/Attachments`;
    let doc; try { doc = parseXml(await fetchText(url)); } catch (_) { return []; }
    const seen = new Set();
    return Array.from(doc.getElementsByTagNameNS(ATOM, 'entry'))
      .map(e => ({ name: dText(e, 'Name'), idUrl: absolute(atomText(e, 'id'), origin) }))
      .filter(a => a.idUrl && !seen.has(a.idUrl) && seen.add(a.idUrl)); // drop repeated attachment IDs
  }
  async function attachmentValue(idUrl) { return fetchText(idUrl.replace(/\/+$/, '') + '/$value'); }

  function makeMatcher(kw, useRegex, cs) {
    if (useRegex) { const re = new RegExp(kw, cs ? '' : 'i'); return (t) => { const m = re.exec(t); return m ? m.index : -1; }; }
    const n = cs ? kw : kw.toLowerCase(); return (t) => (cs ? t : t.toLowerCase()).indexOf(n);
  }
  function snippet(t, at, len) { if (at < 0) return ''; const s = Math.max(0, at - 60), e = Math.min(t.length, at + len + 60); return (s > 0 ? '…' : '') + t.slice(s, e).replace(/\s+/g, ' ') + (e < t.length ? '…' : ''); }

  async function mapLimit(items, limit, worker, onTick) {
    let idx = 0, active = 0, done = 0;
    return new Promise((resolve) => {
      const pump = () => {
        if (idx >= items.length && active === 0) return resolve();
        while (active < limit && idx < items.length) {
          const it = items[idx++]; active++;
          Promise.resolve(worker(it)).catch(() => {}).finally(() => { active--; done++; onTick && onTick(done, items.length); pump(); });
        }
      };
      pump();
    });
  }

  async function run(params, emit) {
    if (running) { emit({ type: 'error', message: 'A search is already running.' }); return; }
    running = true;
    const origin = location.origin;
    results = [];
    const matcher = makeMatcher(params.keyword, params.useRegex, params.caseSensitive);
    const wantName = (params.payloadName || '').toLowerCase().trim();
    try {
      const mpls = await listMpls(origin, params, emit);
      emit({ type: 'progress', message: `Scanning attachments in ${mpls.length} messages…` });
      const emitted = new Set();   // message IDs already claimed
      await mapLimit(mpls, 6, async (m) => {
        if (emitted.has(m.guid)) return;      // never the same message twice
        emitted.add(m.guid);                  // claim now, before any await (closes the parallel-worker race)
        const atts = await attachmentsFor(origin, m.guid);
        const cand = wantName ? atts.filter(a => (a.name || '').toLowerCase().includes(wantName)) : atts;
        for (const a of cand) {
          let content; try { content = await attachmentValue(a.idUrl); } catch (_) { continue; }
          const at = matcher(content);
          if (at >= 0) {
            const hit = { messageGuid: m.guid, applicationMessageId: m.appMsgId, logStart: m.logStart, logEnd: m.logEnd, status: m.status, iflow: m.iflow, attachmentName: a.name, valueUrl: a.idUrl.replace(/\/+$/, '') + '/$value', snippet: snippet(content, at, params.keyword.length) };
            emit({ type: 'hit', hit }); // the emit handler is the ONLY place that appends to results
            break; // one row per message, even if several attachments match
          }
        }
      }, (done, total) => { if (done % 5 === 0 || done === total) emit({ type: 'progress', message: `Scanned ${done}/${total} — ${results.length} hit(s)` }); });
      await ss.set({ lastResults: results });
      emit({ type: 'done' });
    } catch (e) { emit({ type: 'error', message: e.message || String(e) }); }
    finally { running = false; }
  }

  /* =========================================================
   *  UI  (dockable shadow-DOM panel)
   * ========================================================= */
  const esc = (s) => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function localNow() { const d = new Date(); d.setSeconds(0, 0); return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 16); }
  // CPI returns LogStart/LogEnd in UTC (Atom DateTime has no zone suffix). Mark it UTC, then show in the user's local time.
  function toLocalDisplay(v) {
    if (!v) return '';
    let s = String(v).trim();
    if (!/[zZ]$|[+-]\d\d:?\d\d$/.test(s)) s += 'Z';
    const d = new Date(s);
    if (isNaN(d.getTime())) return String(v);
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }
  const $ = (id) => root.getElementById(id);

  const CSS = `
    :host { all: initial; }
    *, *::before, *::after { box-sizing: border-box; }
    .panel { position: fixed; z-index: 2147483646; background:#fff; color:#1d2733;
      font:13px/1.45 -apple-system,Segoe UI,Roboto,sans-serif; box-shadow:0 0 16px rgba(0,0,0,.28);
      display:flex; flex-direction:column; border:1px solid #d0d7de; }
    .panel.dock-right { top:0; right:0; height:100vh; width:440px; border-right:none; }
    .panel.dock-left  { top:0; left:0;  height:100vh; width:440px; border-left:none; }
    .panel.dock-bottom{ left:0; bottom:0; width:100vw; height:44vh; border-bottom:none; }
    .hidden { display:none !important; }
    .hdr { display:flex; align-items:center; gap:6px; padding:8px 10px; background:#0a6ed1; color:#fff; }
    .hdr .title { font-weight:700; font-size:13px; flex:1; }
    .hdr button { background:rgba(255,255,255,.18); color:#fff; border:none; border-radius:5px; padding:3px 7px; cursor:pointer; font-size:13px; }
    .hdr button:hover { background:rgba(255,255,255,.32); }
    .hdr button.active { background:#fff; color:#0a6ed1; }
    .body { padding:10px 12px; overflow:auto; flex:1; }
    label { display:block; margin:7px 0; font-weight:600; }
    label.inline { display:inline-flex; align-items:center; gap:6px; font-weight:400; margin-right:16px; }
    input[type=text], input[type=datetime-local], select { width:100%; padding:6px 8px; margin-top:3px; border:1px solid #c4cdd6; border-radius:6px; font:inherit; font-weight:400; }
    .row { display:flex; gap:10px; } .row > label { flex:1; }
    .row.checks { margin:6px 0; }
    .actions { display:flex; gap:8px; margin:10px 0; flex-wrap:wrap; }
    button.act { padding:7px 12px; border:1px solid #c4cdd6; border-radius:6px; background:#f4f6f8; cursor:pointer; font-weight:600; font:inherit; }
    button.act.primary { background:#0a6ed1; color:#fff; border-color:#0a6ed1; }
    button.act:disabled { opacity:.5; cursor:default; }
    .status-line { margin:6px 0; min-height:18px; color:#4a5560; }
    table { width:100%; border-collapse:collapse; margin-top:6px; }
    th,td { text-align:left; padding:4px 6px; border-bottom:1px solid #eceff1; vertical-align:top; }
    th { font-size:11px; color:#6b7682; position:sticky; top:0; background:#fff; }
    .muted { color:#6b7682; } .small { font-size:11px; }
    .mono { font-family:ui-monospace,SFMono-Regular,Menlo,monospace; }
    a.open { color:#0a6ed1; cursor:pointer; text-decoration:none; } a.open:hover { text-decoration:underline; }
    .resizer { position:absolute; background:transparent; z-index:2; }
    .dock-right .resizer { left:-3px; top:0; width:6px; height:100%; cursor:ew-resize; }
    .dock-left  .resizer { right:-3px; top:0; width:6px; height:100%; cursor:ew-resize; }
    .dock-bottom .resizer{ left:0; top:-3px; height:6px; width:100%; cursor:ns-resize; }
  `;

  const HTML = `
    <div class="panel dock-right" id="panel">
      <div class="resizer" id="resizer"></div>
      <div class="hdr">
        <span class="title">PayloadPeek <span style="opacity:.65;font-weight:400;font-size:11px;">v3.1.4</span></span>
        <button id="dock-left"   title="Dock left">&#9706;</button>
        <button id="dock-bottom" title="Dock bottom">&#9645;</button>
        <button id="dock-right"  title="Dock right">&#9707;</button>
        <button id="close" title="Hide (use the toolbar icon to reopen)">&#10005;</button>
      </div>
      <div class="body">
        <label>Keyword <input id="keyword" type="text" placeholder="text to find inside the attachment payload"></label>
        <label>Payload (attachment) name <input id="payloadName" type="text" placeholder="matches d:Name — blank = all attachments"></label>
        <div class="row">
          <label>From <input id="from" type="datetime-local"></label>
          <label>To <input id="to" type="datetime-local"></label>
        </div>
        <div class="row">
          <label>Status
            <select id="status">
              <option value="ANY">Any</option><option value="COMPLETED">Completed</option>
              <option value="FAILED">Failed</option><option value="RETRY">Retry</option>
              <option value="PROCESSING">Processing</option><option value="ESCALATED">Escalated</option>
            </select>
          </label>
          <label>Integration flow (exact) <input id="iflow" type="text" placeholder="optional"></label>
        </div>
        <div class="row checks">
          <label class="inline"><input id="caseSensitive" type="checkbox"> Case sensitive</label>
          <label class="inline"><input id="useRegex" type="checkbox"> Regex</label>
        </div>
        <div class="actions">
          <button id="run" class="act primary">Search</button>
          <button id="exportCsv" class="act" disabled>Export CSV</button>
          <button id="copyIds" class="act" disabled>Copy message IDs</button>
        </div>
        <div class="status-line" id="status-line"></div>
        <div id="results"></div>
      </div>
    </div>`;

  function setStatus(t) { $('status-line').textContent = t; }

  function renderResults() {
    const el = $('results');
    const seenG = new Set();
    const view = results.filter(r => r && !seenG.has(r.messageGuid) && seenG.add(r.messageGuid)); // unique by message ID
    if (!view.length) { el.innerHTML = '<p class="muted">No matches yet.</p>'; $('exportCsv').disabled = $('copyIds').disabled = true; return; }
    el.innerHTML = `<table>
      <thead><tr><th>LogStart (local)</th><th>Status</th><th>IFlow</th><th>Attachment</th><th>Message GUID</th><th>Match</th></tr></thead>
      <tbody>${view.map((r, i) => `
        <tr>
          <td class="mono small">${esc(toLocalDisplay(r.logStart))}</td>
          <td>${esc(r.status || '')}</td>
          <td>${esc(r.iflow || '')}</td>
          <td><a class="open" data-i="${i}" title="Open payload in new tab">${esc(r.attachmentName || '(open)')}</a></td>
          <td class="mono small">${esc(r.messageGuid)}</td>
          <td class="small">${esc(r.snippet || '')}</td>
        </tr>`).join('')}</tbody></table>`;
    el.querySelectorAll('a.open').forEach(a => a.addEventListener('click', (e) => {
      e.preventDefault();
      window.open(view[+a.dataset.i].valueUrl, '_blank'); // new tab; panel stays put
    }));
    $('exportCsv').disabled = $('copyIds').disabled = false;
  }

  function toCsv() {
    const head = ['LogStart (local)', 'LogEnd (local)', 'Status', 'IntegrationFlow', 'AttachmentName', 'MessageGuid', 'ApplicationMessageId', 'ValueUrl', 'Snippet'];
    const q = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const seenG = new Set();
    const view = results.filter(r => r && !seenG.has(r.messageGuid) && seenG.add(r.messageGuid));
    return [head.join(',')].concat(view.map(r => [toLocalDisplay(r.logStart), toLocalDisplay(r.logEnd), r.status, r.iflow, r.attachmentName, r.messageGuid, r.applicationMessageId, r.valueUrl, r.snippet].map(q).join(','))).join('\n');
  }

  function setDock(pos) {
    dock = pos;
    const panel = $('panel');
    panel.classList.remove('dock-left', 'dock-right', 'dock-bottom');
    panel.classList.add('dock-' + pos);
    panel.style.width = ''; panel.style.height = '';
    ['left', 'right', 'bottom'].forEach(p => $('dock-' + p).classList.toggle('active', p === pos));
    chrome.storage.local.set({ dock: pos });
  }

  function wireResizer() {
    const panel = $('panel'), handle = $('resizer');
    let active = false;
    handle.addEventListener('mousedown', (e) => { active = true; e.preventDefault(); document.body.style.userSelect = 'none'; });
    window.addEventListener('mousemove', (e) => {
      if (!active) return;
      if (dock === 'right') panel.style.width = Math.min(900, Math.max(300, window.innerWidth - e.clientX)) + 'px';
      else if (dock === 'left') panel.style.width = Math.min(900, Math.max(300, e.clientX)) + 'px';
      else panel.style.height = Math.min(window.innerHeight - 40, Math.max(160, window.innerHeight - e.clientY)) + 'px';
    });
    window.addEventListener('mouseup', () => { active = false; document.body.style.userSelect = ''; });
  }

  function emit(msg) {
    if (msg.type === 'progress') setStatus(msg.message);
    else if (msg.type === 'hit') { results.push(msg.hit); renderResults(); }
    else if (msg.type === 'done') { renderResults(); setStatus(`Done — ${new Set(results.map(r => r.messageGuid)).size} match(es).`); $('run').disabled = false; }
    else if (msg.type === 'error') { setStatus('Error: ' + msg.message); $('run').disabled = false; }
  }

  function wire() {
    ['left', 'right', 'bottom'].forEach(p => $('dock-' + p).addEventListener('click', () => setDock(p)));
    $('close').addEventListener('click', () => hidePanel());
    $('run').addEventListener('click', () => {
      const keyword = $('keyword').value.trim();
      if (!keyword) { setStatus('Enter a keyword.'); return; }
      results = []; renderResults(); setStatus('Starting…'); $('run').disabled = true;
      run({
        keyword,
        payloadName: $('payloadName').value,
        from: $('from').value || null,
        to: $('to').value || null,
        status: $('status').value,
        iflow: $('iflow').value.trim(),
        caseSensitive: $('caseSensitive').checked,
        useRegex: $('useRegex').checked
      }, emit);
    });
    $('exportCsv').addEventListener('click', () => {
      const url = URL.createObjectURL(new Blob([toCsv()], { type: 'text/csv' }));
      const a = document.createElement('a'); a.href = url; a.download = `cpi-attachment-search-${Date.now()}.csv`; a.click();
      URL.revokeObjectURL(url);
    });
    $('copyIds').addEventListener('click', () => { navigator.clipboard.writeText([...new Set(results.map(r => r.messageGuid))].join('\n')); setStatus('Message GUIDs copied.'); });
    wireResizer();
  }

  async function buildPanel() {
    document.getElementById('__cpi_attach_search_host')?.remove(); // clear any stale panel from an old/duplicate instance
    host = document.createElement('div');
    host.id = '__cpi_attach_search_host';
    root = host.attachShadow({ mode: 'open' });
    root.innerHTML = `<style>${CSS}</style>${HTML}`;
    document.documentElement.appendChild(host);
    wire();

    const { dock: savedDock } = await chrome.storage.local.get('dock');
    setDock(savedDock || 'right');

    $('to').value = localNow();              // default "To" = current date/time
    const { lastResults } = await ss.get('lastResults');
    if (lastResults && lastResults.length) { results = lastResults; renderResults(); setStatus(`Showing ${results.length} result(s) from last run.`); }
    else renderResults();
  }

  async function showPanel() { if (!host) await buildPanel(); $('panel').classList.remove('hidden'); chrome.storage.local.set({ panelOpen: true }); }
  function hidePanel() { if (host) $('panel').classList.add('hidden'); chrome.storage.local.set({ panelOpen: false }); }
  async function togglePanel() {
    if (!host) { await showPanel(); return; }
    const hidden = $('panel').classList.toggle('hidden');
    chrome.storage.local.set({ panelOpen: !hidden });
  }

  chrome.runtime.onMessage.addListener((msg, _s, reply) => {
    if (msg?.type === 'ping') { reply({ ok: true }); return false; }
    if (msg?.type === 'toggle') { togglePanel(); reply({ ok: true }); return false; }
  });

  // Auto-restore the panel after a page reload if it was open before.
  chrome.storage.local.get('panelOpen').then(({ panelOpen }) => { if (panelOpen) showPanel(); });
}

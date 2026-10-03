const $ = (id) => document.getElementById(id);
const state = {
  ticker: 'MSFT',
  catalog: null,
  token: '',
  host: new URLSearchParams(location.search).get('office') === '1' ? 'office' : 'browser',
  plan: null,
  research: null,
  researchAbort: null,
  generation: 0,
  socket: null,
  reconnectTimer: null,
  stopped: false,
};
let officeBridge;
const commands = new Map();
let sequence = 0;

function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = String(text);
  return node;
}
function notice(message, error = false) {
  $('notice').textContent = message;
  $('notice').classList.toggle('error', error);
  $('notice').hidden = false;
}
function clearNotice() {
  $('notice').hidden = true;
}
function money(value, eps = false) {
  return Number(value).toLocaleString('en-US', {
    minimumFractionDigits: eps ? 2 : 0,
    maximumFractionDigits: eps ? 2 : 0,
  });
}
function billion(value) {
  return '$' + (Number(value) / 1000).toFixed(1) + 'B';
}
function fact(metric, period = 'FY2025') {
  return state.catalog.facts.find(
    (f) => f.ticker === state.ticker && f.metric === metric && f.period === period,
  );
}

async function api(path, { method = 'GET', body, signal } = {}) {
  const timeout = new AbortController();
  const timer = setTimeout(() => timeout.abort(), path === '/research' ? 85000 : 20000);
  const combined = signal ? AbortSignal.any([signal, timeout.signal]) : timeout.signal;
  try {
    const response = await fetch('/api' + path, {
      method,
      headers: {
        Authorization: 'Bearer ' + state.token,
        ...(body ? { 'Content-Type': 'application/json' } : {}),
      },
      body: body ? JSON.stringify(body) : undefined,
      signal: combined,
    });
    const content = await response.json();
    if (!response.ok) throw new Error(content.message ?? `Request failed (${response.status}).`);
    return content;
  } finally {
    clearTimeout(timer);
  }
}

function native(command, payload = {}) {
  if (state.host === 'office')
    return officeBridge
      ? officeBridge.execute(command, payload)
      : Promise.reject(new Error('The Office connection is not ready.'));
  if (!window.chrome?.webview)
    return Promise.reject(new Error('Open the LedgerLens pane in Excel to use workbook actions.'));
  const id = String(++sequence);
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      commands.delete(id);
      reject(
        new Error('Excel did not finish the operation. Close any open Excel dialog and try again.'),
      );
    }, 30000);
    commands.set(id, { resolve, reject, timer });
    window.chrome.webview.postMessage({ id, command, payload });
  });
}
if (state.host !== 'office')
  window.chrome?.webview?.addEventListener('message', (event) => {
    const data = event.data;
    if (data.type === 'bootstrap') {
      state.host = 'excel';
      state.token = data.token;
      $('host-state').textContent = 'Connected to Excel';
      initialize().catch((e) => notice(e.message, true));
    } else if (data.type === 'navigate') showTab(data.tab);
    else if (data.type === 'source') showSource(data.fact);
    else if (data.id && commands.has(data.id)) {
      const pending = commands.get(data.id);
      commands.delete(data.id);
      clearTimeout(pending.timer);
      if (data.error) pending.reject(new Error(data.error));
      else pending.resolve(data.result);
    }
  });

function showTab(name) {
  if (!['overview', 'research', 'refresh', 'health'].includes(name)) return;
  document.querySelectorAll('.tab').forEach((button) => {
    const active = button.dataset.tab === name;
    button.classList.toggle('active', active);
    if (active) button.setAttribute('aria-current', 'page');
    else button.removeAttribute('aria-current');
  });
  document.querySelectorAll('.view').forEach((view) => {
    const active = view.id === 'view-' + name;
    view.hidden = !active;
    view.classList.toggle('active', active);
  });
  clearNotice();
  if (name === 'health' && state.token) refreshHealth().catch((e) => notice(e.message, true));
}

function renderCompany() {
  const company = state.catalog.companies.find((c) => c.ticker === state.ticker);
  $('company-name').textContent = company.name;
  $('company-avatar').textContent = company.name[0];
  $('company-detail').textContent = `${company.ticker} · Fiscal year ends ${company.fiscalYearEnd}`;
  $('research-company').textContent = company.ticker;
  document
    .querySelectorAll('[data-company]')
    .forEach((b) => b.classList.toggle('selected', b.dataset.company === state.ticker));
  const revenue = fact('Revenue');
  const prior = fact('Revenue', 'FY2024');
  $('revenue-value').textContent = billion(revenue.value);
  const change = prior.value ? (revenue.value / prior.value - 1) * 100 : null;
  $('revenue-change').textContent =
    change === null ? 'No prior base' : `${change >= 0 ? '+' : ''}${change.toFixed(1)}% YoY`;
  const values = ['FY2023', 'FY2024', 'FY2025'].map((p) => fact('Revenue', p).value);
  const min = Math.min(...values) * 0.85;
  const max = Math.max(...values) * 1.05;
  const spread = max - min || 1;
  const points = values.map((v, i) => [i * 150 + 3, 59 - ((v - min) / spread) * 54]);
  const path = points.map((p, i) => `${i ? 'L' : 'M'} ${p[0]} ${p[1]}`).join(' ');
  // All SVG coordinates are computed from validated numeric source facts.
  $('revenue-chart').innerHTML =
    `<svg viewBox="0 0 306 64" preserveAspectRatio="none" role="img" aria-label="Annual revenue ${values.map((value) => money(value)).join(', ')} USD millions"><defs><linearGradient id="revenue-gradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#82cbb3" stop-opacity=".24"/><stop offset="100%" stop-color="#82cbb3" stop-opacity="0"/></linearGradient></defs><path d="${path} L303 64 L3 64 Z" fill="url(#revenue-gradient)"/><path d="${path}" stroke="#8cdac0" stroke-width="2" fill="none"/>${points.map((p) => `<circle cx="${p[0]}" cy="${p[1]}" r="3" fill="#baf5dc" stroke="#254a4e" stroke-width="2"/>`).join('')}</svg>`;
  $('metric-list').replaceChildren();
  for (const metric of [
    'OperatingIncome',
    'NetIncome',
    'OperatingCashFlow',
    'CapitalExpenditure',
  ]) {
    const f = fact(metric);
    const row = element('div', 'metric-row');
    row.append(element('span', '', f.label));
    const value = element('button', '', money(f.value));
    value.setAttribute('aria-label', `${f.label}: ${money(f.value)} ${f.unit}. Inspect source`);
    value.append(element('small', '', '↗'));
    value.onclick = () => showSource(f);
    row.append(value);
    $('metric-list').append(row);
  }
  updateFormula();
}
function updateFormula() {
  $('formula-preview').textContent =
    `=LL.METRIC("${state.ticker}","${$('formula-metric').value}","${$('formula-period').value}")`;
}

function showSource(source) {
  if (!source) return;
  const details = $('source-details');
  details.replaceChildren();
  details.append(
    element('h2', '', `${source.ticker} · ${source.label}`),
    element('div', 'source-value', money(source.value, source.metric === 'DilutedEPS')),
    element('p', 'fine-print', `${source.unit} · ${source.period}`),
  );
  const fields = element('dl', 'source-fields');
  for (const [label, value] of [
    ['Reporting period', source.start ? `${source.start} → ${source.end}` : source.end],
    ['Filed', source.filed],
    ['XBRL concept', source.concept],
    ['SEC accession', source.accession],
    ['Evidence acquired', new Date(source.acquiredAt).toLocaleDateString('en-CA')],
  ]) {
    const row = element('div');
    row.append(element('dt', '', label), element('dd', '', value));
    fields.append(row);
  }
  details.append(fields);
  const link = element('a', 'source-link', 'Open the original SEC filing ↗');
  try {
    const url = new URL(source.sourceUrl);
    if (url.protocol === 'https:' && url.hostname === 'www.sec.gov') link.href = url.href;
  } catch {}
  link.target = '_blank';
  link.rel = 'noopener noreferrer';
  details.append(link);
  details.append(
    element(
      'p',
      'fine-print',
      'The filing index provides the original report and its XBRL documents. Dates describe the fiscal reporting period, not a market quote.',
    ),
  );
  if (!$('source-dialog').open) $('source-dialog').showModal();
}

function renderAnswer(answer) {
  const card = element('article', 'answer-card');
  card.append(
    element(
      'div',
      'answer-meta',
      `${answer.isAiGenerated ? '✦' : '∑'} ${answer.provider} · ${answer.cached ? 'saved answer' : 'citation IDs validated'}`,
    ),
    element('h2', '', answer.headline),
    element('p', 'answer-summary', answer.summary),
  );
  for (const claim of answer.claims) {
    const row = element('div', 'claim');
    row.append(element('p', '', claim.text));
    for (const id of [...new Set(claim.sourceIds)]) {
      const source = answer.sources.find((f) => f.sourceId === id);
      if (!source) continue;
      const button = element('button', 'citation', `${source.period} · ${source.label}`);
      button.onclick = () => showSource(source);
      row.append(button);
    }
    card.append(row);
  }
  const caveats = element('div', 'caveats');
  answer.caveats.forEach((text) => caveats.append(element('p', '', text)));
  card.append(caveats);
  card.append(
    element(
      'p',
      'fine-print',
      `${answer.model} · ${new Date(answer.generatedAt).toLocaleTimeString()} · Verify interpretations before using them.`,
    ),
  );
  if (state.host === 'excel' || state.host === 'office') {
    const insert = element('button', 'button secondary full', 'Save research to workbook');
    insert.onclick = async () => {
      try {
        await native('saveResearch', { answer });
        notice('Research and its citations were saved to a new worksheet.');
      } catch (e) {
        notice(e.message, true);
      }
    };
    card.append(insert);
  }
  $('research-answer').replaceChildren(card);
  $('research-empty').hidden = true;
}

async function research() {
  state.researchAbort?.abort();
  const controller = new AbortController();
  state.researchAbort = controller;
  const generation = ++state.generation;
  $('run-research').disabled = true;
  $('research-progress').hidden = false;
  clearNotice();
  try {
    const answer = await api('/research', {
      method: 'POST',
      body: {
        ticker: state.ticker,
        question: $('research-question').value,
        useAi: $('use-ai').checked,
      },
      signal: controller.signal,
    });
    if (generation !== state.generation) return;
    state.research = answer;
    renderAnswer(answer);
  } catch (error) {
    if (generation !== state.generation) return;
    notice(
      error.name === 'AbortError'
        ? 'Research canceled. Your workbook was not changed.'
        : error.message,
      error.name !== 'AbortError',
    );
  } finally {
    if (generation === state.generation) {
      $('run-research').disabled = false;
      $('research-progress').hidden = true;
    }
  }
}

function renderPlan(plan) {
  const card = element('article', 'card');
  card.append(
    element('h3', '', `${plan.changes.length} reported values ready`),
    element(
      'p',
      'fine-print',
      `${plan.preservedAddresses.length} formula or assumption cells preserved. Values in USD millions, except EPS.`,
    ),
  );
  for (const change of plan.changes) {
    const row = element('div', 'change-row');
    const top = element('div');
    top.append(element('strong', '', change.source.label), element('span', '', change.address));
    const values = element('div', 'change-values');
    const before = change.expectedContent.startsWith('n:')
      ? money(Number(change.expectedContent.slice(2)))
      : 'Empty';
    values.append(
      element('del', '', before),
      element('strong', '', '→ ' + money(change.newValue, change.source.metric === 'DilutedEPS')),
    );
    row.append(top, values);
    card.append(row);
  }
  const apply = element(
    'button',
    'button primary full',
    plan.changes.length ? 'Apply reviewed values' : 'Model is already current',
  );
  apply.disabled = !plan.changes.length;
  apply.style.marginTop = '16px';
  apply.id = 'apply-refresh';
  apply.onclick = async () => {
    apply.disabled = true;
    try {
      const result = await native('applyRefresh', { planId: plan.id });
      notice(
        `${result.changed} reported values updated. Forecasts and assumptions were preserved.`,
      );
      $('refresh-plan').replaceChildren();
      state.plan = null;
    } catch (e) {
      notice(e.message, true);
      apply.disabled = false;
    }
  };
  card.append(apply);
  $('refresh-plan').replaceChildren(card);
}

async function refreshHealth() {
  const health = await api('/diagnostics');
  const online = health.mode === 'online';
  $('connection-state').classList.toggle('offline', !online);
  $('connection-state').replaceChildren(
    element('i'),
    document.createTextNode(online ? 'Connected' : 'Offline cache'),
  );
  $('health-headline').textContent = `Service ready · SEC circuit ${health.secCircuit}`;
  $('health-description').textContent =
    `${health.sourceFacts} sourced facts · ${health.aiConfigured ? 'OpenAI configured' : 'Calculated analysis available'}`;
  $('health-grid').replaceChildren();
  for (const [value, label] of [
    [health.snapshotReads, 'Snapshot reads'],
    [health.secRequests, 'SEC HTTP requests'],
    [health.secRetries, 'SEC retries'],
    [health.secActiveRequests, 'Active SEC requests'],
    [health.secPeakConcurrency + ' / 4', 'Peak SEC concurrency'],
    [health.aiCalls, 'AI requests'],
  ]) {
    const card = element('div', 'health-stat');
    card.append(element('strong', '', value), element('span', '', label));
    $('health-grid').append(card);
  }
  document
    .querySelectorAll('[data-mode]')
    .forEach((button) => button.classList.toggle('selected', button.dataset.mode === health.mode));
  $('event-list').replaceChildren();
  [...health.recentEvents]
    .reverse()
    .slice(0, 12)
    .forEach((event) => {
      const row = element('div', 'event');
      row.append(
        element('strong', '', event.message),
        element('small', '', new Date(event.at).toLocaleTimeString()),
      );
      $('event-list').append(row);
    });
}

function connectStream(attempt = 0) {
  if (state.stopped || !state.token) return;
  if (state.socket) {
    state.socket.onclose = null;
    state.socket.close();
  }
  const url = new URL('/api/events', location.href);
  url.protocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const socket = new WebSocket(url, ['ledgerlens.v1', 'll-auth.' + state.token]);
  state.socket = socket;
  socket.onopen = () => {
    attempt = 0;
    $('stream-status').textContent = '● Live';
  };
  socket.onmessage = () => {
    if (!$('view-health').hidden) refreshHealth().catch(() => {});
  };
  socket.onclose = () => {
    $('stream-status').textContent = 'Reconnecting…';
    state.reconnectTimer = setTimeout(
      () => connectStream(attempt + 1),
      Math.min(1000 * 2 ** attempt, 15000),
    );
  };
  socket.onerror = () => socket.close();
}

let initialized = false;
async function initialize() {
  if (initialized) return;
  state.catalog = await api('/catalog');
  initialized = true;
  $('company-switcher').replaceChildren();
  for (const company of state.catalog.companies) {
    const button = element('button', '', company.ticker);
    button.dataset.company = company.ticker;
    button.onclick = () => {
      if (state.ticker === company.ticker) return;
      state.ticker = company.ticker;
      state.researchAbort?.abort();
      ++state.generation;
      $('research-progress').hidden = true;
      $('run-research').disabled = false;
      $('research-answer').replaceChildren();
      $('research-empty').hidden = false;
      renderCompany();
    };
    $('company-switcher').append(button);
  }
  $('formula-metric').replaceChildren();
  for (const metric of state.catalog.metrics) {
    const option = element('option', '', metric.label);
    option.value = metric.id;
    $('formula-metric').append(option);
  }
  $('snapshot-note').textContent = 'Evidence snapshot: ' + state.catalog.snapshotDate + '.';
  if (state.host === 'browser') {
    $('host-state').textContent = 'Browser workspace';
    $('insert-formula').disabled = true;
    $('preview-refresh').disabled = true;
    $('rollback-refresh').disabled = true;
    $('insert-hint').textContent = 'Copy a formula here, or open the Excel pane to insert it.';
    $('refresh-hint').textContent =
      'Open the LedgerLens pane in Excel to preview and apply workbook changes.';
  } else if (state.host === 'office') {
    $('host-state').textContent = 'Connected via Office.js';
    $('insert-formula').textContent = 'Insert sourced value';
    $('insert-hint').textContent =
      'This Office.js preview inserts a snapshot value. LL formulas run in the Windows native add-in.';
    $('preview-refresh').disabled = true;
    $('rollback-refresh').disabled = true;
    $('refresh-hint').textContent =
      'Guarded model refresh is available in the Windows native add-in. This Office.js preview supports sourced values and research export.';
  }
  renderCompany();
  await refreshHealth();
  connectStream();
}

document
  .querySelectorAll('[data-tab]')
  .forEach((button) => (button.onclick = () => showTab(button.dataset.tab)));
document.querySelectorAll('[data-question]').forEach(
  (button) =>
    (button.onclick = () => {
      $('research-question').value = button.dataset.question;
      $('research-question').focus();
    }),
);
$('formula-metric').onchange = updateFormula;
$('formula-period').onchange = updateFormula;
$('copy-formula').onclick = async () => {
  try {
    await navigator.clipboard.writeText($('formula-preview').textContent);
    notice('Formula copied.');
  } catch {
    notice('Select the formula text and copy it with Ctrl+C.');
  }
};
$('insert-formula').onclick = async () => {
  try {
    await native('insertFormula', {
      ticker: state.ticker,
      metric: $('formula-metric').value,
      period: $('formula-period').value,
    });
    notice(
      state.host === 'office'
        ? 'Sourced snapshot value inserted into the selected empty cell.'
        : 'Formula inserted into the selected empty cell.',
    );
  } catch (e) {
    notice(e.message, true);
  }
};
$('sync-sec').onclick = async () => {
  $('sync-sec').disabled = true;
  try {
    const result = await api('/sync/' + state.ticker, { method: 'POST' });
    state.catalog = await api('/catalog');
    renderCompany();
    notice(
      `${result.count} ${result.ticker} facts checked against SEC. Review model changes before applying them.`,
    );
  } catch (e) {
    notice(e.message, true);
  } finally {
    $('sync-sec').disabled = false;
  }
};
$('run-research').onclick = research;
$('cancel-research').onclick = () => state.researchAbort?.abort();
$('close-source').onclick = () => $('source-dialog').close();
$('source-dialog').addEventListener('click', (e) => {
  if (e.target === $('source-dialog')) {
    const r = e.target.getBoundingClientRect();
    if (e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom)
      e.target.close();
  }
});
$('preview-refresh').onclick = async () => {
  $('preview-refresh').disabled = true;
  try {
    state.plan = await native('previewRefresh');
    renderPlan(state.plan);
  } catch (e) {
    notice(e.message, true);
  } finally {
    $('preview-refresh').disabled = false;
  }
};
$('rollback-refresh').onclick = async () => {
  try {
    const result = await native('rollbackRefresh');
    notice(`${result.changed} values restored. Later edits were preserved.`);
    $('refresh-plan').replaceChildren();
  } catch (e) {
    notice(e.message, true);
  }
};
document.querySelectorAll('[data-mode]').forEach(
  (button) =>
    (button.onclick = async () => {
      try {
        await api('/connection', { method: 'POST', body: { mode: button.dataset.mode } });
        if (state.host === 'excel') await native('recalculate');
        await refreshHealth();
        notice('Connection scenario: ' + button.textContent + '.');
      } catch (e) {
        notice(e.message, true);
      }
    }),
);
$('test-connection').onclick = async () => {
  try {
    await api('/events/test', { method: 'POST' });
    await refreshHealth();
  } catch (e) {
    notice(e.message, true);
  }
};
$('clear-cache').onclick = async () => {
  try {
    await api('/cache/clear', { method: 'POST' });
    await refreshHealth();
    notice('Request caches cleared; saved evidence is retained.');
  } catch (e) {
    notice(e.message, true);
  }
};
const healthTimer = setInterval(() => {
  if (initialized && !document.hidden && !$('view-health').hidden) refreshHealth().catch(() => {});
}, 5000);
window.addEventListener('beforeunload', () => {
  state.stopped = true;
  clearInterval(healthTimer);
  clearTimeout(state.reconnectTimer);
  state.researchAbort?.abort();
  if (state.socket) {
    state.socket.onclose = null;
    state.socket.close();
  }
  for (const p of commands.values()) clearTimeout(p.timer);
});

if (state.host !== 'office' && window.chrome?.webview)
  window.chrome.webview.postMessage({ command: 'ready' });
else {
  const parameters = new URLSearchParams(location.hash.slice(1));
  const incoming = parameters.get('session');
  if (incoming) {
    sessionStorage.setItem('ledgerlens-session', incoming);
    history.replaceState(null, '', location.pathname + location.search);
  }
  state.token = sessionStorage.getItem('ledgerlens-session') ?? '';
  if (state.token && state.host === 'office')
    import('./office-bridge.js')
      .then((module) => module.initializeOffice(api))
      .then((bridge) => {
        officeBridge = bridge;
        return initialize();
      })
      .catch((e) => notice(e.message, true));
  else if (state.token) initialize().catch((e) => notice(e.message, true));
  else {
    $('host-state').textContent = 'Workspace not connected';
    notice(
      'Launch LedgerLens from the included Start LedgerLens script to connect this workspace.',
    );
  }
}

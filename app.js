'use strict';

// ════════════════════════════════════════════════════════════
//  A. CONFIG
// ════════════════════════════════════════════════════════════
const CONFIG = {
  sheetUrl:      'https://script.google.com/macros/s/AKfycbw8cXJF40QNeW-0-IfgTpzZBMl8ydSL9GcyG-8Y0EWo7Aa5wnNzJMwImWIFbGGF9Cu8uw/exec',
  fetchInterval: 60 * 60 * 1000, // 1 hour in ms
  maxRows:       200,
  rowsPerPage:   10,
  chartMaxPoints:60,
};

// ════════════════════════════════════════════════════════════
//  B. STATE — single source of truth
// ════════════════════════════════════════════════════════════
const state = {
  // Current sensor reading
  current: { pm25:0, pm10:0, temp:0, hum:0, co2:0, volt:0, curr:0, pwr:0, energy:0 },
  // Previous sensor reading (for delta calculation)
  previous: { pm25:0, pm10:0, temp:0, hum:0, co2:0, volt:0, curr:0, pwr:0, energy:0 },
  // All records loaded from sheet (newest first)
  rows: [],
  // Filtered + sorted view for table rendering
  filteredRows: [],
  // Active alert objects
  alerts: [],
  // Alert timeline entries
  alertHistory: [],
  // Table UI state
  table: { sortKey: 'time', sortDir: 1, filterStatus: 'all', currentPage: 1 },
  // Export and date filter state
  exportRange: { preset: 'all', from: '', to: '' },
  // Active chart range
  chartRange: '1h',
  // Fetch metadata
  fetch: { isFetching: false, lastFetchAt: null, countdown: null },
  // Alert thresholds
  thresholds: { pm25: 35, pm10: 50, co2: 600, temp: 35, hum: 85, pwr: 3000 },
};

// ════════════════════════════════════════════════════════════
//  C. DOM CACHE
// ════════════════════════════════════════════════════════════
const DOM = {
  // Clock
  heroDate:      () => document.getElementById('hero-date'),
  heroClock:     () => document.getElementById('hero-clock'),
  // AQI Banner
  aqiScore:      () => document.getElementById('aqi-score'),
  aqiStatusText: () => document.getElementById('aqi-status-text'),
  aqiDot:        () => document.getElementById('aqi-dot'),
  bannerPm25:    () => document.getElementById('banner-pm25'),
  bannerPm10:    () => document.getElementById('banner-pm10'),
  bannerCo2:     () => document.getElementById('banner-co2'),
  gaugePm25:     () => document.getElementById('gauge-pm25'),
  gaugePm10:     () => document.getElementById('gauge-pm10'),
  gaugeCo2:      () => document.getElementById('gauge-co2'),
  miniTemp:      () => document.getElementById('mini-temp'),
  miniHum:       () => document.getElementById('mini-hum'),
  miniPwr:       () => document.getElementById('mini-pwr'),
  // Status strip
  statusLast:    () => document.getElementById('status-last'),
  statusFeed:    () => document.getElementById('status-feed'),
  statusFeedHealth: () => document.getElementById('status-feed-health'),
  statusFeedNote:   () => document.getElementById('status-feed-note'),
  statusSensorHealth: () => document.getElementById('status-sensor-health'),
  statusSensorNote:   () => document.getElementById('status-sensor-note'),
  statusTrend:   () => document.getElementById('status-trend'),
  statusAlerts:  () => document.getElementById('status-alerts'),
  // Metric cards
  valPm25:       () => document.getElementById('val-pm25'),
  valPm10:       () => document.getElementById('val-pm10'),
  valTemp:       () => document.getElementById('val-temp'),
  valHum:        () => document.getElementById('val-hum'),
  valCo2:        () => document.getElementById('val-co2'),
  valVolt:       () => document.getElementById('val-volt'),
  valCurr:       () => document.getElementById('val-curr'),
  deltaPm25:     () => document.getElementById('delta-pm25'),
  deltaPm10:     () => document.getElementById('delta-pm10'),
  deltaTemp:     () => document.getElementById('delta-temp'),
  deltaHum:      () => document.getElementById('delta-hum'),
  deltaCo2:      () => document.getElementById('delta-co2'),
  deltaVolt:     () => document.getElementById('delta-volt'),
  deltaCurr:     () => document.getElementById('delta-curr'),
  // Insight panel
  insightPwr:    () => document.getElementById('insight-pwr'),
  insightEnergy: () => document.getElementById('insight-energy'),
  insightPf:     () => document.getElementById('insight-pf'),
  insightEff:    () => document.getElementById('insight-eff'),
  barPwr:        () => document.getElementById('bar-pwr'),
  barEnergy:     () => document.getElementById('bar-energy'),
  barPf:         () => document.getElementById('bar-pf'),
  // Table
  tableSearch:   () => document.getElementById('table-search'),
  rangeSummary:  () => document.getElementById('range-summary'),
  rangeFrom:     () => document.getElementById('range-from'),
  rangeTo:       () => document.getElementById('range-to'),
  tableBody:     () => document.getElementById('table-body'),
  pageInfo:      () => document.getElementById('page-info'),
  pageNumbers:   () => document.getElementById('page-numbers'),
  btnPrev:       () => document.getElementById('btn-prev'),
  btnNext:       () => document.getElementById('btn-next'),
  // Toast
  toast:         () => document.getElementById('toast'),
  toastMessage:  () => document.getElementById('toast-message'),
  // Refresh button
  btnRefresh:    () => document.getElementById('btn-refresh'),
  // Skeleton overlays
  skeletons:     () => ['sk-pm25','sk-pm10','sk-temp','sk-hum','sk-volt','sk-curr','sk-insight'].map(id => document.getElementById(id)).filter(Boolean),
};

// ════════════════════════════════════════════════════════════
//  D. UTILITY FUNCTIONS
// ════════════════════════════════════════════════════════════
const fmt = (v, d = 1) => (isNaN(v) || v === null) ? '—' : (+v).toFixed(d);

function setHTML(el, html) { if (el) el.innerHTML = html; }
function setText(el, text) { if (el) el.textContent = text; }
function setWidth(el, pct)  { if (el) el.style.width = Math.min(100, Math.max(0, pct)) + '%'; }

function syncPowerUnitUI() {
  const miniUnit = DOM.miniPwr()?.querySelector('.aqi-mini-unit');
  if (miniUnit) miniUnit.textContent = 'W';

  const insightUnit = DOM.insightPwr()?.querySelector('span');
  if (insightUnit) insightUnit.textContent = 'W';

  const energyLegend = [...document.querySelectorAll('.chart-legend .legend-item')]
    .find(el => el.textContent.includes('Power'));
  if (energyLegend) energyLegend.innerHTML = '<div class="legend-dot" style="background:#006977"></div>Power (W)';

  const powerLabel = document.querySelector('label[for="thr-pwr"]') ||
    [...document.querySelectorAll('.modal-label')].find(el => el.textContent.includes('Power Max'));
  if (powerLabel) powerLabel.textContent = 'Power Max (W)';

  const powerInput = document.getElementById('thr-pwr');
  if (powerInput) {
    powerInput.value = '3000';
    powerInput.step = '10';
  }
}

function parseTimestamp(value) {
  if (!value) return new Date();
  const normalized = String(value).trim().replace(' ', 'T');
  const date = new Date(normalized);
  return Number.isNaN(date.getTime()) ? new Date() : date;
}

function normalizePower(value) {
  return +(value ?? 0);
}

function deltaClass(diff) {
  if (Math.abs(diff) < 0.01) return 'delta-flat';
  return diff > 0 ? 'delta-up' : 'delta-down';
}
function deltaLabel(diff, unit) {
  const arrow = diff > 0 ? '↑' : diff < 0 ? '↓' : '—';
  return `${arrow} ${Math.abs(diff).toFixed(1)}${unit}`;
}
function applyDelta(el, diff, unit) {
  if (!el) return;
  el.className = `delta ${deltaClass(diff)}`;
  el.textContent = deltaLabel(diff, unit);
}

function formatTime(date) {
  if (!(date instanceof Date) || isNaN(date)) return '—';
  return date.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}
function formatTimeShort(date) {
  if (!(date instanceof Date) || isNaN(date)) return '—';
  return date.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
}

// ════════════════════════════════════════════════════════════
//  E. AQI CALCULATION
// ════════════════════════════════════════════════════════════
function calcAqiFromPm25(pm) {
  const bp = [[0,12,0,50],[12.1,35.4,51,100],[35.5,55.4,101,150],[55.5,150.4,151,200],[150.5,250.4,201,300]];
  for (const [lo, hi, al, ah] of bp) {
    if (pm >= lo && pm <= hi) return Math.round(((ah - al) / (hi - lo)) * (pm - lo) + al);
  }
  return Math.round(pm * 1.5);
}
function calcAqiFromPm10(pm) {
  const bp = [[0,54,0,50],[55,154,51,100],[155,254,101,150],[255,354,151,200]];
  for (const [lo, hi, al, ah] of bp) {
    if (pm >= lo && pm <= hi) return Math.round(((ah - al) / (hi - lo)) * (pm - lo) + al);
  }
  return Math.round(pm * 0.7);
}
function calcAQI() {
  return Math.max(calcAqiFromPm25(state.current.pm25), calcAqiFromPm10(state.current.pm10));
}
function getAqiMeta(aqi) {
  if (aqi <= 50)  return { label: 'Good — Acceptable',              dot: '#a8e6c8' };
  if (aqi <= 100) return { label: 'Moderate',                       dot: '#fde68a' };
  if (aqi <= 150) return { label: 'Unhealthy for Sensitive Groups', dot: '#fca5a5' };
  return             { label: 'Unhealthy',                          dot: '#f87171' };
}

// ════════════════════════════════════════════════════════════
//  F. DATA FETCHING & NORMALIZATION
// ════════════════════════════════════════════════════════════
// Map status string from Sheet → dashboard display label
function mapSheetStatus(raw) {
  const v = String(raw || '').toLowerCase().trim();
  if (v === 'critical') return 'Attention';
  if (v === 'warning')  return 'Moderate';
  return 'Good'; // normal or empty
}

function normalizeRow(raw) {
  const time = parseTimestamp(raw.timestamp);
  return {
    time,
    pm25:   +(raw.pm25        ?? 0),
    pm10:   +(raw.pm10        ?? 0),
    temp:   +(raw.temperature ?? 0),
    hum:    Math.round(+(raw.humidity ?? 0)),
    co2:    Math.round(+(raw.co2      ?? 0)),
    volt:   +(raw.voltage     ?? 0),
    curr:   +(raw.current     ?? 0),
    pwr:    normalizePower(raw.power),
    energy: +(raw.energy      ?? 0),
    // Use status directly from Google Sheet (normal → Good, warning → Moderate, critical → Attention)
  };
}

function normalizeRows(rawArray) {
  return rawArray
    .map(normalizeRow)
    .reverse()                               // newest first
    .slice(0, CONFIG.maxRows);
}

function extractLatestRow(rawArray) {
  return rawArray[rawArray.length - 1];      // Apps Script sends oldest → newest
}

function applyLatestToState(latest) {
  Object.assign(state.previous, state.current);
  state.current.pm25   = +(latest.pm25        ?? 0);
  state.current.pm10   = +(latest.pm10        ?? 0);
  state.current.temp   = +(latest.temperature ?? 0);
  state.current.hum    = +(latest.humidity    ?? 0);
  state.current.co2    = +(latest.co2         ?? 0);
  state.current.volt   = +(latest.voltage     ?? 0);
  state.current.curr   = +(latest.current     ?? 0);
  state.current.pwr    = normalizePower(latest.power);
  state.current.energy = +(latest.energy      ?? 0);
}

function syncRowsState(rawArray) {
  state.rows = normalizeRows(rawArray);
}

function handleFetchError(err) {
  setText(DOM.statusFeed(), 'Fetch Error');
  setText(DOM.statusLast(), 'Failed');
  showToast('⚠ ' + err.message);
  console.error('[Atmosfera] fetch error:', err);
  // Retry after interval even on error
  startCountdown();
}

async function fetchSheetData() {
  const res = await fetch(CONFIG.sheetUrl);
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const json = await res.json();
  if (json.status !== 'ok') throw new Error(json.message || 'API returned error');
  if (!json.data || json.data.length === 0) throw new Error('No records in response');
  return json.data;
}

async function fetchSheet() {
  if (state.fetch.isFetching) return;
  state.fetch.isFetching = true;
  setText(DOM.statusFeed(), 'Fetching…');

  try {
    const rawData = await fetchSheetData();
    const latest  = extractLatestRow(rawData);
    applyLatestToState(latest);
    syncRowsState(rawData);

    renderDashboard();
    recomputeTableView();
    updateCharts();

    state.fetch.lastFetchAt = new Date();
    setText(DOM.statusFeed(), 'Online');
    setText(DOM.statusLast(), 'Just now');
    showToast(`Loaded ${rawData.length} records`);
    startCountdown();

  } catch (err) {
    handleFetchError(err);
  } finally {
    state.fetch.isFetching = false;
    hideSkeleton();
  }
}

// ════════════════════════════════════════════════════════════
//  G. STATUS / ALERT COMPUTATION — single source of truth
// ════════════════════════════════════════════════════════════
function computeRowStatus(row) {
  const t = state.thresholds;
  if (row.pm25 > t.pm25 || row.pm10 > t.pm10 || row.co2 > t.co2 ||
      row.temp > t.temp || row.hum  > t.hum  || row.pwr > t.pwr) {
    return 'Attention';
  }
  if (row.pm25 < 12 && row.pm10 < 20 && row.co2 < 450) return 'Good';
  return 'Moderate';
}

// ════════════════════════════════════════════════════════════
//  H. ALERT ENGINE
// ════════════════════════════════════════════════════════════
const ALERT_CHECKS = [
  { key: 'pm25', label: 'PM 2.5',      unit: 'µg/m³', getValue: () => state.current.pm25 },
  { key: 'pm10', label: 'PM 10',       unit: 'µg/m³', getValue: () => state.current.pm10 },
  { key: 'co2',  label: 'CO₂',         unit: 'ppm',   getValue: () => state.current.co2  },
  { key: 'temp', label: 'Temperature', unit: '°C',    getValue: () => state.current.temp },
  { key: 'hum',  label: 'Humidity',    unit: '%',     getValue: () => state.current.hum  },
  { key: 'pwr',  label: 'Power',       unit: 'W',     getValue: () => state.current.pwr  },
];

function evaluateAlerts() {
  ALERT_CHECKS.forEach(check => {
    const value     = check.getValue();
    const threshold = state.thresholds[check.key];
    const exceeded  = value > threshold;
    const idx       = state.alerts.findIndex(a => a.key === check.key);

    if (exceeded && idx === -1) {
      const now = new Date();
      const alert = {
        id: `${check.key}-${now.getTime()}`,
        key: check.key,
        label: check.label,
        value,
        peakValue: value,
        unit: check.unit,
        threshold,
        triggeredAt: now,
        acknowledgedAt: null,
      };
      state.alerts.push(alert);
      state.alertHistory.unshift({
        id: alert.id,
        key: alert.key,
        label: alert.label,
        unit: alert.unit,
        threshold: alert.threshold,
        startedAt: now,
        latestValue: value,
        peakValue: value,
        acknowledgedAt: null,
        resolvedAt: null,
      });
    } else if (!exceeded && idx !== -1) {
      const alert = state.alerts[idx];
      const entry = getAlertHistoryEntry(alert.id);
      if (entry) {
        entry.latestValue = value;
        entry.resolvedAt = new Date();
      }
      state.alerts.splice(idx, 1);
    } else if (exceeded && idx !== -1) {
      const alert = state.alerts[idx];
      alert.value = value;
      alert.threshold = threshold;
      alert.peakValue = Math.max(alert.peakValue, value);
      const entry = getAlertHistoryEntry(alert.id);
      if (entry) {
        entry.latestValue = value;
        entry.threshold = threshold;
        entry.peakValue = Math.max(entry.peakValue, value);
      }
    }
  });
  state.alertHistory = state.alertHistory.slice(0, 30);
}

// ════════════════════════════════════════════════════════════
//  I. RENDERING FUNCTIONS
// ════════════════════════════════════════════════════════════
function renderMetricCards() {
  const c = state.current;
  setHTML(DOM.valPm25(), `${fmt(c.pm25)}<span class="metric-unit">µg/m³</span>`);
  setHTML(DOM.valPm10(), `${fmt(c.pm10)}<span class="metric-unit">µg/m³</span>`);
  setHTML(DOM.valTemp(), `${fmt(c.temp)}<span class="metric-unit">°C</span>`);
  setHTML(DOM.valHum(),  `${fmt(c.hum, 0)}<span class="metric-unit">%</span>`);
  setHTML(DOM.valCo2(),  `${fmt(c.co2, 0)}<span class="metric-unit">ppm</span>`);
  setHTML(DOM.valVolt(), `${fmt(c.volt)}<span class="metric-unit">V</span>`);
  setHTML(DOM.valCurr(), `${fmt(c.curr)}<span class="metric-unit">A</span>`);
}

function renderDeltaBadges() {
  const c = state.current;
  const p = state.previous;
  applyDelta(DOM.deltaPm25(), c.pm25 - p.pm25, '');
  applyDelta(DOM.deltaPm10(), c.pm10 - p.pm10, '');
  applyDelta(DOM.deltaTemp(), c.temp - p.temp, '°');
  applyDelta(DOM.deltaHum(),  c.hum  - p.hum,  '%');
  applyDelta(DOM.deltaVolt(), c.volt - p.volt, 'V');
  applyDelta(DOM.deltaCurr(), c.curr - p.curr, 'A');
  // CO2 delta — styled inline for the featured card
  const co2El = DOM.deltaCo2();
  if (co2El) {
    const diff = c.co2 - p.co2;
    co2El.className = 'delta';
    co2El.style.background = 'rgba(255,255,255,.12)';
    co2El.style.color = 'rgba(255,255,255,.7)';
    co2El.textContent = deltaLabel(diff, '');
  }
}

function renderAqiBanner() {
  const c   = state.current;
  const aqi = calcAQI();
  const meta = getAqiMeta(aqi);

  setText(DOM.aqiScore(),      aqi);
  setText(DOM.aqiStatusText(), meta.label);
  if (DOM.aqiDot()) DOM.aqiDot().style.background = meta.dot;

  setText(DOM.bannerPm25(), `${fmt(c.pm25)} µg/m³`);
  setText(DOM.bannerPm10(), `${fmt(c.pm10)} µg/m³`);
  setText(DOM.bannerCo2(),  `${fmt(c.co2, 0)} ppm`);
  setHTML(DOM.miniTemp(),   `${fmt(c.temp)}<span class="aqi-mini-unit">°C</span>`);
  setHTML(DOM.miniHum(),    `${fmt(c.hum, 0)}<span class="aqi-mini-unit">%</span>`);
  setHTML(DOM.miniPwr(),    `${fmt(c.pwr, 1)}<span class="aqi-mini-unit">W</span>`);

  setWidth(DOM.gaugePm25(), (c.pm25 / 200) * 100);
  setWidth(DOM.gaugePm10(), (c.pm10 / 400) * 100);
  setWidth(DOM.gaugeCo2(),  ((c.co2 - 350) / 4650) * 100);
}

function renderInsightPanel() {
  const c  = state.current;
  const t  = state.thresholds;
  const pf = (c.volt > 0 && c.curr > 0) ? Math.min(1, c.pwr / (c.volt * c.curr)) : 0;
  const eff = pf >= .95 ? 'A+' : pf >= .9 ? 'A' : pf >= .85 ? 'B+' : 'B';
  const effColor = pf >= .95 ? 'var(--tertiary)' : pf >= .9 ? 'var(--secondary)' : 'var(--primary)';

  setHTML(DOM.insightPwr(),    `${fmt(c.pwr, 1)} <span style="font-size:.85rem;font-weight:400;opacity:.5">W</span>`);
  setHTML(DOM.insightEnergy(), `${fmt(c.energy, 2)} <span style="font-size:.85rem;font-weight:400;opacity:.5">kWh</span>`);
  setHTML(DOM.insightPf(),     `${pf.toFixed(2)} <span style="font-size:.85rem;font-weight:400;opacity:.5">PF</span>`);
  setText(DOM.insightEff(),    eff);
  if (DOM.insightEff()) DOM.insightEff().style.color = effColor;

  setWidth(DOM.barPwr(),    (c.pwr    / t.pwr) * 100);
  setWidth(DOM.barEnergy(), (c.energy / 30)     * 100);
  setWidth(DOM.barPf(),     pf                  * 100);

  // Update CO2 donut
  if (charts.donut) {
    const pct = Math.min(100, ((c.co2 - 350) / 4650) * 100);
    charts.donut.data.datasets[0].data = [pct, 100 - pct];
    charts.donut.update('none');
  }
}

function renderStatusStrip() {
  const c = state.current;
  const p = state.previous;
  const trendDiff = (c.pm25 - p.pm25) + (c.pm10 - p.pm10);
  const trend = trendDiff < -0.5 ? 'Improving' : trendDiff > 0.5 ? 'Worsening' : 'Stable';
  setText(DOM.statusTrend(), trend);
}

function renderTable() {
  const { currentPage } = state.table;
  const perPage = CONFIG.rowsPerPage;
  const start   = (currentPage - 1) * perPage;
  const end     = start + perPage;
  const page    = state.filteredRows.slice(start, end);
  const total   = state.filteredRows.length;
  const totalPages = Math.max(1, Math.ceil(total / perPage));

  const tbody = DOM.tableBody();
  if (!tbody) return;

  if (page.length === 0) {
    tbody.innerHTML = `<tr><td colspan="11" class="no-results">No matching records found</td></tr>`;
  } else {
    tbody.innerHTML = page.map(r => {
      const badgeClass = r.status === 'Good' ? 'badge-good' : r.status === 'Attention' ? 'badge-alert' : 'badge-moderate';
      const badgeIcon  = r.status === 'Good' ? '✓' : r.status === 'Attention' ? '⚠' : '~';
      return `<tr>
        <td class="td-time">${formatTime(r.time)}</td>
        <td class="td-value">${fmt(r.pm25)}</td>
        <td class="td-value">${fmt(r.pm10)}</td>
        <td class="td-value">${fmt(r.temp)}°</td>
        <td class="td-value">${r.hum}%</td>
        <td class="td-value">${r.co2}</td>
        <td class="td-value">${fmt(r.volt)} V</td>
        <td class="td-value">${fmt(r.curr, 2)} A</td>
        <td class="td-value">${fmt(r.pwr, 1)} W</td>
        <td class="td-value">${fmt(r.energy, 2)} kWh</td>
      </tr>`;
    }).join('');
  }

  const s = total === 0 ? 0 : start + 1;
  const e = Math.min(end, total);
  setText(DOM.pageInfo(), `Showing ${s}–${e} of ${total} records`);
  if (DOM.btnPrev()) DOM.btnPrev().disabled = currentPage <= 1;
  if (DOM.btnNext()) DOM.btnNext().disabled = currentPage >= totalPages;

  const pnEl = DOM.pageNumbers();
  if (pnEl) {
    pnEl.innerHTML = '';
    for (let i = 1; i <= totalPages; i++) {
      if (totalPages <= 7 || Math.abs(i - currentPage) < 3 || i === 1 || i === totalPages) {
        const btn = document.createElement('button');
        btn.className = 'page-btn' + (i === currentPage ? ' active' : '');
        btn.textContent = i;
        btn.addEventListener('click', () => { state.table.currentPage = i; renderTable(); });
        pnEl.appendChild(btn);
      }
    }
  }
}

// Master render entry point
function renderDashboard() {
  renderMetricCards();
  renderDeltaBadges();
  renderAqiBanner();
  renderInsightPanel();
  renderStatusStrip();
}

// ════════════════════════════════════════════════════════════
//  J. TABLE VIEW — search / filter / sort / paginate
// ════════════════════════════════════════════════════════════
function filterRows(rows) {
  const { filterStatus } = state.table;
  const query = DOM.tableSearch() ? DOM.tableSearch().value.toLowerCase().trim() : '';
  return rows.filter(r => {
    if (filterStatus !== 'all' && r.status !== filterStatus) return false;
    if (query) {
      const haystack = `${formatTime(r.time)} ${r.pm25} ${r.pm10} ${r.temp} ${r.hum} ${r.co2} ${r.volt} ${r.curr} ${r.pwr} ${r.energy} ${r.status}`.toLowerCase();
      if (!haystack.includes(query)) return false;
    }
    return true;
  });
}

function sortRows(rows) {
  const { sortKey, sortDir } = state.table;
  return [...rows].sort((a, b) => {
    const av = sortKey === 'time' ? a.time.getTime() : a[sortKey];
    const bv = sortKey === 'time' ? b.time.getTime() : b[sortKey];
    if (av === bv) return 0;
    return sortDir === 1 ? (av < bv ? 1 : -1) : (av < bv ? -1 : 1);
  });
}

function recomputeTableView() {
  state.table.currentPage = 1;
  state.filteredRows = sortRows(filterRows(state.rows));
  renderTable();
}

function applySortLegacy(key) {
  if (state.table.sortKey === key) {
    state.table.sortDir *= -1;
  } else {
    state.table.sortKey = key;
    state.table.sortDir = 1;
  }
  document.querySelectorAll('thead th').forEach(th => {
    th.classList.remove('sorted');
    const icon = th.querySelector('.sort-icon');
    if (icon) icon.textContent = '↕';
  });
  const th = document.getElementById('th-' + key);
  if (th) {
    th.classList.add('sorted');
    const icon = th.querySelector('.sort-icon');
    if (icon) icon.textContent = state.table.sortDir === 1 ? '↓' : '↑';
  }
  recomputeTableView();
  return;
  if (th) {
    th.classList.add('sorted');
    const icon = th.querySelector('.sort-icon');
    if (icon) icon.textContent = state.table.sortDir === -1 ? '↓' : '↑';
  }
  recomputeTableView();
}

function applyFilter(filterValue) {
  state.table.filterStatus = filterValue;
  document.querySelectorAll('.table-filter').forEach(b => b.classList.remove('active'));
  const btn = document.querySelector(`.table-filter[data-filter="${filterValue}"]`);
  if (btn) btn.classList.add('active');
  recomputeTableView();
}

function applySort(key) {
  if (state.table.sortKey === key) {
    state.table.sortDir *= -1;
  } else {
    state.table.sortKey = key;
    state.table.sortDir = 1;
  }

  document.querySelectorAll('thead th').forEach(th => {
    th.classList.remove('sorted');
    const icon = th.querySelector('.sort-icon');
    if (icon) icon.textContent = '↕';
  });

  const activeTh = document.getElementById('th-' + key);
  if (activeTh) {
    activeTh.classList.add('sorted');
    const icon = activeTh.querySelector('.sort-icon');
    if (icon) icon.textContent = state.table.sortDir === 1 ? '↓' : '↑';
  }

  recomputeTableView();
}

// ════════════════════════════════════════════════════════════
//  K. CHARTS
// ════════════════════════════════════════════════════════════
const charts = { aq: null, co2: null, energy: null, donut: null };

const chartData = {
  aq:     { labels: [], pm25: [], pm10: [] },
  co2:    { labels: [], co2:  [] },
  energy: { labels: [], volt: [], curr: [], pwr:  [] },
};

function makeGradient(ctx, colorTop, colorBottom, height = 220) {
  const g = ctx.createLinearGradient(0, 0, 0, height);
  g.addColorStop(0, colorTop);
  g.addColorStop(1, colorBottom);
  return g;
}

const CHART_TOOLTIP = {
  backgroundColor: '#fff', titleColor: '#2a3439', bodyColor: '#5a6a72',
  borderColor: 'rgba(90,106,114,.12)', borderWidth: 1, padding: 12, cornerRadius: 12,
};

function initCharts() {
  Chart.defaults.font.family = "'Inter', sans-serif";
  Chart.defaults.color = '#5a6a72';

  const aqCtx = document.getElementById('chart-aq').getContext('2d');
  charts.aq = new Chart(aqCtx, {
    type: 'line',
    data: { labels: chartData.aq.labels, datasets: [
      { label: 'PM 2.5 (µg/m³)', data: chartData.aq.pm25, borderColor: '#466370', borderWidth: 2, pointRadius: 0, fill: true, backgroundColor: makeGradient(aqCtx, 'rgba(70,99,112,.15)', 'rgba(70,99,112,0)'), tension: .45 },
      { label: 'PM 10 (µg/m³)',  data: chartData.aq.pm10, borderColor: '#4a9da8', borderWidth: 2, pointRadius: 0, fill: true, backgroundColor: makeGradient(aqCtx, 'rgba(74,157,168,.12)', 'rgba(74,157,168,0)'), tension: .45 },
      { label: 'WHO Limit',      data: [],                borderColor: 'rgba(93,94,97,.25)', borderWidth: 1.5, borderDash: [6,4], pointRadius: 0, fill: false, tension: 0 },
    ]},
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: { legend: { display: false }, tooltip: { ...CHART_TOOLTIP, callbacks: { label: ctx => `${ctx.dataset.label}: ${(+ctx.parsed.y).toFixed(1)}` } } },
      scales: {
        x: { grid: { color: 'rgba(90,106,114,.07)', drawBorder: false }, ticks: { maxTicksLimit: 6 } },
        y: { grid: { color: 'rgba(90,106,114,.07)', drawBorder: false }, ticks: { maxTicksLimit: 5 }, min: 0, title: { display: true, text: 'µg/m³', color: '#5a6a72', font: { size: 10 } } },
      }, animation: { duration: 200 },
    },
  });

  const co2Ctx = document.getElementById('chart-co2').getContext('2d');
  charts.co2 = new Chart(co2Ctx, {
    type: 'line',
    data: { labels: chartData.co2.labels, datasets: [
      { label: 'CO₂ (ppm)', data: chartData.co2.co2, borderColor: '#006977', borderWidth: 2, pointRadius: 0, fill: true, backgroundColor: makeGradient(co2Ctx, 'rgba(0,105,119,.15)', 'rgba(0,105,119,0)'), tension: .45 },
    ]},
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: { legend: { display: false }, tooltip: { ...CHART_TOOLTIP, callbacks: { label: ctx => `CO₂: ${Math.round(ctx.parsed.y)} ppm` } } },
      scales: {
        x: { grid: { color: 'rgba(0,0,0,.04)', drawBorder: false }, ticks: { maxTicksLimit: 5 } },
        y: { grid: { color: 'rgba(0,0,0,.04)', drawBorder: false }, ticks: { maxTicksLimit: 5 }, title: { display: true, text: 'ppm', color: '#5a6a72', font: { size: 10 } } },
      }, animation: { duration: 200 },
    },
  });

  const enCtx = document.getElementById('chart-energy').getContext('2d');
  charts.energy = new Chart(enCtx, {
    type: 'line',
    data: { labels: chartData.energy.labels, datasets: [
      { label: 'Voltage (V)', data: chartData.energy.volt, borderColor: '#5d5e61', borderWidth: 2, pointRadius: 0, fill: false, tension: .4 },
      { label: 'Current (A)', data: chartData.energy.curr, borderColor: '#466370', borderWidth: 2, pointRadius: 0, fill: false, tension: .4 },
      { label: 'Power (W)',   data: chartData.energy.pwr,  borderColor: '#006977', borderWidth: 2, pointRadius: 0, fill: true, backgroundColor: makeGradient(enCtx, 'rgba(0,105,119,.08)', 'rgba(0,105,119,0)', 200), tension: .4 },
    ]},
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: { legend: { display: false }, tooltip: { ...CHART_TOOLTIP, callbacks: { label: ctx => {
        const v = ctx.parsed.y;
        if (ctx.datasetIndex === 0) return `Voltage: ${(+v).toFixed(1)} V`;
        if (ctx.datasetIndex === 1) return `Current: ${(+v).toFixed(2)} A`;
        return `Power: ${(+v).toFixed(1)} W`;
      }}}},
      scales: {
        x: { grid: { color: 'rgba(0,0,0,.04)', drawBorder: false }, ticks: { maxTicksLimit: 6 } },
        y: { grid: { color: 'rgba(0,0,0,.04)', drawBorder: false }, ticks: { maxTicksLimit: 5 } },
      }, animation: { duration: 200 },
    },
  });

  // CO2 Donut
  const donutCtx = document.getElementById('chart-donut').getContext('2d');
  charts.donut = new Chart(donutCtx, {
    type: 'doughnut',
    data: { datasets: [{ data: [0, 100], backgroundColor: ['rgba(255,255,255,.7)', 'rgba(255,255,255,.12)'], borderWidth: 0, hoverOffset: 4 }] },
    options: { responsive: false, cutout: '74%', plugins: { legend: { display: false }, tooltip: { enabled: false } }, animation: { duration: 600 } },
  });

  Chart.register({
    id: 'donutLabel',
    afterDraw(chart) {
      if (chart.canvas.id !== 'chart-donut') return;
      const { ctx, width, height } = chart;
      ctx.save();
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillStyle = 'rgba(255,255,255,.9)';
      ctx.font = '700 18px Manrope, sans-serif';
      ctx.fillText(Math.round(state.current.co2), width / 2, height / 2 - 8);
      ctx.fillStyle = 'rgba(255,255,255,.45)';
      ctx.font = '400 10px Inter, sans-serif';
      ctx.fillText('ppm', width / 2, height / 2 + 10);
      ctx.restore();
    },
  });
}

function updateChartsLegacy() {
  const MAX = CONFIG.chartMaxPoints;
  const src = [...state.rows].reverse().slice(-MAX); // oldest → newest

  const labels = src.map(r => formatTimeShort(r.time));
  const pts    = src.length;

  // Sync data arrays
  const sync = (arr, values) => { arr.length = 0; values.forEach(v => arr.push(v)); };

  sync(chartData.aq.labels, labels);
  sync(chartData.aq.pm25,   src.map(r => r.pm25));
  sync(chartData.aq.pm10,   src.map(r => r.pm10));
  charts.aq.data.datasets[2].data = Array(pts).fill(15);
  charts.aq.update();

  sync(chartData.co2.labels, labels);
  sync(chartData.co2.co2,    src.map(r => r.co2));
  charts.co2.update();

  sync(chartData.energy.labels, labels);
  sync(chartData.energy.volt,   src.map(r => r.volt));
  sync(chartData.energy.curr,   src.map(r => r.curr));
  sync(chartData.energy.pwr,    src.map(r => r.pwr));
  charts.energy.update();
}

function applyChartRangeLegacy(rangeKey) {
  const pts = rangeKey === '1h' ? 10 : rangeKey === '6h' ? 20 : CONFIG.chartMaxPoints;
  const slicedLabels = chartData.aq.labels.slice(-pts);

  charts.aq.data.labels            = slicedLabels;
  charts.aq.data.datasets[0].data  = chartData.aq.pm25.slice(-pts);
  charts.aq.data.datasets[1].data  = chartData.aq.pm10.slice(-pts);
  charts.aq.data.datasets[2].data  = Array(Math.min(pts, chartData.aq.pm25.length)).fill(15);
  charts.aq.update();

  charts.co2.data.labels           = chartData.co2.labels.slice(-pts);
  charts.co2.data.datasets[0].data = chartData.co2.co2.slice(-pts);
  charts.co2.update();

  charts.energy.data.labels            = chartData.energy.labels.slice(-pts);
  charts.energy.data.datasets[0].data  = chartData.energy.volt.slice(-pts);
  charts.energy.data.datasets[1].data  = chartData.energy.curr.slice(-pts);
  charts.energy.data.datasets[2].data  = chartData.energy.pwr.slice(-pts);
  charts.energy.update();
}

function getRowsForChartRange(rangeKey) {
  const src = [...state.rows].reverse().slice(-CONFIG.chartMaxPoints);
  if (rangeKey === '24h' || src.length === 0) return src;

  const latest = src[src.length - 1].time;
  if (!(latest instanceof Date) || Number.isNaN(latest.getTime())) return src;

  const rangeMs = rangeKey === '1h' ? 60 * 60 * 1000 : 6 * 60 * 60 * 1000;
  return src.filter(row => (latest - row.time) <= rangeMs);
}

function updateCharts() {
  applyChartRange(state.chartRange);
}

function applyChartRange(rangeKey) {
  state.chartRange = rangeKey;

  const src = getRowsForChartRange(rangeKey);
  const labels = src.map(r => formatTimeShort(r.time));

  charts.aq.data.labels            = labels;
  charts.aq.data.datasets[0].data  = src.map(r => r.pm25);
  charts.aq.data.datasets[1].data  = src.map(r => r.pm10);
  charts.aq.data.datasets[2].data  = Array(src.length).fill(15);
  charts.aq.update();

  charts.co2.data.labels           = labels;
  charts.co2.data.datasets[0].data = src.map(r => r.co2);
  charts.co2.update();

  charts.energy.data.labels            = labels;
  charts.energy.data.datasets[0].data  = src.map(r => r.volt);
  charts.energy.data.datasets[1].data  = src.map(r => r.curr);
  charts.energy.data.datasets[2].data  = src.map(r => r.pwr);
  charts.energy.update();
}

// ════════════════════════════════════════════════════════════
//  L. COUNTDOWN TIMER — no overlap guard
// ════════════════════════════════════════════════════════════
function startCountdown() {
  clearInterval(state.fetch.countdown);
  let remaining = CONFIG.fetchInterval / 1000;

  state.fetch.countdown = setInterval(() => {
    remaining--;
    const m = Math.floor(remaining / 60);
    const s = remaining % 60;
    setText(DOM.statusLast(), `Next in ${m}:${String(s).padStart(2, '0')}`);
    if (remaining <= 0) {
      clearInterval(state.fetch.countdown);
      fetchSheet();
    }
  }, 1000);
}

// ════════════════════════════════════════════════════════════
//  M. SKELETON OVERLAY
// ════════════════════════════════════════════════════════════
function showSkeleton() {
  DOM.skeletons().forEach(el => {
    el.classList.add('visible');
    if (el.id === 'sk-insight') el.classList.add('flex');
  });
}

function hideSkeleton() {
  DOM.skeletons().forEach(el => {
    el.style.transition = 'opacity .3s ease';
    el.style.opacity    = '0';
    setTimeout(() => {
      el.classList.remove('visible', 'flex');
      el.style.opacity    = '';
      el.style.transition = '';
    }, 300);
  });
}

// ════════════════════════════════════════════════════════════
//  N. TOAST
// ════════════════════════════════════════════════════════════
let toastTimer = null;
function showToast(message) {
  const t = DOM.toast();
  setText(DOM.toastMessage(), message);
  if (t) t.classList.add('visible');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { if (t) t.classList.remove('visible'); }, 3000);
}

// ════════════════════════════════════════════════════════════
//  O. EXPORT CSV
// ════════════════════════════════════════════════════════════
function exportCSV() {
  if (state.rows.length === 0) { showToast('No data to export yet'); return; }
  const header = 'Time,PM2.5 (µg/m³),PM10 (µg/m³),Temperature (°C),Humidity (%),CO2 (ppm),Voltage (V),Current (A),Power (W),Energy (kWh),Status\n';
  const body   = state.rows.map(r => {
    const t = r.time instanceof Date && !isNaN(r.time) ? r.time.toLocaleString('en-GB') : '—';
    return `${t},${fmt(r.pm25)},${fmt(r.pm10)},${fmt(r.temp)},${r.hum},${r.co2},${fmt(r.volt)},${fmt(r.curr, 2)},${fmt(r.pwr, 1)},${fmt(r.energy, 2)},${r.status}`;
  }).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([header + body], { type: 'text/csv' }));
  a.download = `atmosfera_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  showToast(`Exported ${state.rows.length} records`);
}

function exportPDF() {
  if (state.rows.length === 0) { showToast('No data to export yet'); return; }

  const jsPDFCtor = window.jspdf?.jsPDF;
  if (!jsPDFCtor) {
    showToast('PDF export is unavailable');
    return;
  }

  const doc = new jsPDFCtor({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 12;
  const rowHeight = 6;
  const columns = [
    { label: 'Time',    width: 28, align: 'left'  },
    { label: 'PM2.5',   width: 18, align: 'right' },
    { label: 'PM10',    width: 18, align: 'right' },
    { label: 'Temp C',  width: 18, align: 'right' },
    { label: 'Hum %',   width: 16, align: 'right' },
    { label: 'CO2',     width: 18, align: 'right' },
    { label: 'Volt V',  width: 18, align: 'right' },
    { label: 'Curr A',  width: 18, align: 'right' },
    { label: 'Power W', width: 20, align: 'right' },
    { label: 'Energy',  width: 20, align: 'right' },
    { label: 'Status',  width: 24, align: 'left'  },
  ];

  const formatPdfDateTime = date => {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return 'N/A';
    return date.toLocaleString('en-GB', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  };

  const drawTableHeader = y => {
    let x = margin;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setFillColor(240, 244, 247);
    doc.rect(margin, y - 4.5, pageWidth - (margin * 2), rowHeight, 'F');

    columns.forEach(col => {
      const textX = col.align === 'right' ? x + col.width - 1 : x + 1;
      doc.text(col.label, textX, y, { align: col.align });
      x += col.width;
    });

    doc.setDrawColor(210, 218, 225);
    doc.line(margin, y + 1.5, pageWidth - margin, y + 1.5);
  };

  const drawPageHeader = pageNo => {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text('Atmosfera Monitoring Report', margin, margin);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.text(`Generated: ${formatPdfDateTime(new Date())}`, margin, margin + 6);
    doc.text(`Records: ${state.rows.length}`, margin, margin + 11);
    doc.text(`Page: ${pageNo}`, pageWidth - margin, margin + 6, { align: 'right' });

    if (pageNo === 1) {
      const c = state.current;
      const snapshot = [
        `PM2.5 ${fmt(c.pm25)}`,
        `PM10 ${fmt(c.pm10)}`,
        `Temp ${fmt(c.temp)} C`,
        `Hum ${fmt(c.hum, 0)} %`,
        `CO2 ${fmt(c.co2, 0)}`,
        `Power ${fmt(c.pwr, 1)} W`,
        `Alerts ${state.alerts.length}`,
      ].join('  |  ');
      doc.text(snapshot, margin, margin + 17);
    }
  };

  const drawRow = (row, y) => {
    const values = [
      formatPdfDateTime(row.time),
      fmt(row.pm25),
      fmt(row.pm10),
      fmt(row.temp),
      `${row.hum}`,
      `${row.co2}`,
      fmt(row.volt),
      fmt(row.curr, 2),
      fmt(row.pwr, 1),
      fmt(row.energy, 2),
      row.status,
    ];

    let x = margin;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);

    values.forEach((value, idx) => {
      const col = columns[idx];
      const textX = col.align === 'right' ? x + col.width - 1 : x + 1;
      doc.text(String(value), textX, y, { align: col.align });
      x += col.width;
    });

    doc.setDrawColor(232, 237, 241);
    doc.line(margin, y + 1.5, pageWidth - margin, y + 1.5);
  };

  let pageNo = 1;
  let y = margin;
  drawPageHeader(pageNo);
  y = 36;
  drawTableHeader(y);
  y += 8;

  state.rows.forEach(row => {
    if (y > pageHeight - margin) {
      doc.addPage();
      pageNo += 1;
      drawPageHeader(pageNo);
      y = 24;
      drawTableHeader(y);
      y += 8;
    }

    drawRow(row, y);
    y += rowHeight;
  });

  doc.save(`atmosfera_${new Date().toISOString().slice(0, 10)}.pdf`);
  showToast(`Exported ${state.rows.length} records to PDF`);
}

// ════════════════════════════════════════════════════════════
//  P. MODAL HELPERS
// ════════════════════════════════════════════════════════════
function openModal(id)  { const m = document.getElementById(id); if (m) m.classList.add('open'); }
function closeModal(id) { const m = document.getElementById(id); if (m) m.classList.remove('open'); }

function saveThresholds() {
  const read = (id, fallback) => {
    const value = parseFloat(document.getElementById(id)?.value);
    return Number.isFinite(value) ? value : fallback;
  };
  state.thresholds.pm25 = read('thr-pm25', 35);
  state.thresholds.pm10 = read('thr-pm10', 50);
  state.thresholds.co2  = read('thr-co2',  600);
  state.thresholds.temp = read('thr-temp', 35);
  state.thresholds.hum  = read('thr-hum',  85);
  state.thresholds.pwr  = read('thr-pwr',  3000);
  closeModal('modal-thresholds');
  renderDashboard();
  showToast('Thresholds updated');
}

// ════════════════════════════════════════════════════════════
//  Q. CLOCK
// ════════════════════════════════════════════════════════════
const DAYS   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

function tickClock() {
  const now = new Date();
  setText(DOM.heroClock(), now.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
  setText(DOM.heroDate(),  `${DAYS[now.getDay()]}, ${MONTHS[now.getMonth()]} ${now.getDate()}`);
}

// ════════════════════════════════════════════════════════════
//  R. MANUAL REFRESH
// ════════════════════════════════════════════════════════════
function manualRefresh() {
  if (state.fetch.isFetching) return;
  const btn  = DOM.btnRefresh();
  const icon = btn?.querySelector('svg');
  if (btn)  btn.disabled = true;
  if (icon) icon.style.animation = 'spin 0.8s linear infinite';
  clearInterval(state.fetch.countdown);
  showSkeleton();
  fetchSheet().finally(() => {
    if (icon) icon.style.animation = '';
    if (btn)  btn.disabled = false;
  });
}

// ════════════════════════════════════════════════════════════
//  S. EVENT BINDING
// ════════════════════════════════════════════════════════════
function formatDateTime(date) {
  if (!(date instanceof Date) || isNaN(date)) return 'â€”';
  return date.toLocaleString('en-GB', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

function formatDateInputValue(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function parseDateInput(value, endOfDay = false) {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) return null;
  if (endOfDay) date.setHours(23, 59, 59, 999);
  return date;
}

function humanizeDuration(ms) {
  if (!Number.isFinite(ms) || ms < 0) return '0m';
  const minutes = Math.round(ms / 60000);
  if (minutes < 60) return `${minutes}m`;
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  if (hours < 24) return mins === 0 ? `${hours}h` : `${hours}h ${mins}m`;
  const days = Math.floor(hours / 24);
  const remHours = hours % 24;
  return remHours === 0 ? `${days}d` : `${days}d ${remHours}h`;
}

function setStatusTone(valueEl, noteEl, tone) {
  const color =
    tone === 'good' ? 'var(--tertiary)' :
    tone === 'warn' ? '#b45309' :
    tone === 'bad'  ? 'var(--error)' :
    'var(--on-surface)';
  if (valueEl) valueEl.style.color = color;
  if (noteEl) noteEl.style.color = tone === 'neutral' ? 'var(--primary-fixed-dim)' : color;
}

function median(values) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function getLatestRowTime(rows = state.rows) {
  return rows[0]?.time instanceof Date ? rows[0].time : null;
}

function getDataCadenceMs(rows = state.rows) {
  if (!rows || rows.length < 2) return 15 * 60 * 1000;
  const diffs = [];
  for (let i = 0; i < Math.min(rows.length - 1, 10); i++) {
    const newer = rows[i]?.time?.getTime?.();
    const older = rows[i + 1]?.time?.getTime?.();
    if (Number.isFinite(newer) && Number.isFinite(older) && newer > older) diffs.push(newer - older);
  }
  return median(diffs) || 15 * 60 * 1000;
}

function ensureCustomRangeDefaults() {
  if (state.rows.length === 0) return;
  if (!state.exportRange.from) state.exportRange.from = formatDateInputValue(state.rows[state.rows.length - 1].time);
  if (!state.exportRange.to) state.exportRange.to = formatDateInputValue(state.rows[0].time);
}

function getSelectedRangeBounds() {
  if (state.rows.length === 0) return null;
  const latest = getLatestRowTime();
  if (!latest || state.exportRange.preset === 'all') return null;

  if (state.exportRange.preset === 'today') {
    const from = new Date(latest);
    from.setHours(0, 0, 0, 0);
    const to = new Date(latest);
    to.setHours(23, 59, 59, 999);
    return { from, to };
  }

  if (state.exportRange.preset === '7d' || state.exportRange.preset === '30d') {
    const days = state.exportRange.preset === '7d' ? 7 : 30;
    const from = new Date(latest.getTime() - ((days * 24 * 60 * 60 * 1000) - 1));
    return { from, to: latest };
  }

  ensureCustomRangeDefaults();
  const from = parseDateInput(state.exportRange.from);
  const to = parseDateInput(state.exportRange.to, true);
  if (from && to && from <= to) return { from, to };
  return null;
}

function getSelectedRangeLabel() {
  if (state.exportRange.preset === 'all') return 'All records';
  if (state.exportRange.preset === 'today') return 'Today';
  if (state.exportRange.preset === '7d') return 'Last 7 days';
  if (state.exportRange.preset === '30d') return 'Last 30 days';

  const bounds = getSelectedRangeBounds();
  if (!bounds) return 'Custom range';
  return `${formatDateInputValue(bounds.from)} to ${formatDateInputValue(bounds.to)}`;
}

function matchesSelectedRange(row) {
  const bounds = getSelectedRangeBounds();
  if (!bounds) return true;
  return row.time >= bounds.from && row.time <= bounds.to;
}

function getFilteredAndSortedRows(rows = state.rows) {
  return sortRows(filterRows(rows));
}

function getExportRows() {
  return getFilteredAndSortedRows(state.rows);
}

function getRangeSummaryText(count) {
  return `Date range: ${getSelectedRangeLabel()} (${count} records)`;
}

function syncRangeControls() {
  document.querySelectorAll('.range-chip').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.rangePreset === state.exportRange.preset);
  });

  const custom = state.exportRange.preset === 'custom';
  const fromEl = DOM.rangeFrom();
  const toEl = DOM.rangeTo();
  if (fromEl) {
    fromEl.disabled = !custom;
    fromEl.value = state.exportRange.from;
  }
  if (toEl) {
    toEl.disabled = !custom;
    toEl.value = state.exportRange.to;
  }
}

function applyRangePreset(preset) {
  state.exportRange.preset = preset;
  if (preset === 'custom') {
    ensureCustomRangeDefaults();
  }
  syncRangeControls();
  recomputeTableView();
}

function updateCustomRange() {
  state.exportRange.preset = 'custom';
  state.exportRange.from = DOM.rangeFrom()?.value || '';
  state.exportRange.to = DOM.rangeTo()?.value || '';
  syncRangeControls();
  recomputeTableView();
}

function getAlertHistoryEntry(id) {
  return state.alertHistory.find(entry => entry.id === id);
}

function acknowledgeAlert(id) {
  const alert = state.alerts.find(item => item.id === id);
  if (!alert || alert.acknowledgedAt) return;
  const now = new Date();
  alert.acknowledgedAt = now;
  const entry = getAlertHistoryEntry(id);
  if (entry && !entry.acknowledgedAt) entry.acknowledgedAt = now;
  renderAlerts();
}

function acknowledgeAllAlerts() {
  if (state.alerts.length === 0) {
    showToast('No active alerts to acknowledge');
    return;
  }
  const now = new Date();
  state.alerts.forEach(alert => {
    if (!alert.acknowledgedAt) {
      alert.acknowledgedAt = now;
      const entry = getAlertHistoryEntry(alert.id);
      if (entry && !entry.acknowledgedAt) entry.acknowledgedAt = now;
    }
  });
  renderAlerts();
  showToast('Active alerts acknowledged');
}

function computeFeedHealth() {
  if (state.rows.length === 0) {
    return { label: 'No data', note: 'Waiting for samples', tone: 'neutral' };
  }

  const latest = getLatestRowTime();
  if (!latest) return { label: 'Unknown', note: 'Latest timestamp unavailable', tone: 'warn' };

  const ageMs = Date.now() - latest.getTime();
  const cadenceMs = getDataCadenceMs();
  const freshLimit = Math.max(cadenceMs * 1.5, 20 * 60 * 1000);
  const delayedLimit = Math.max(cadenceMs * 4, 2 * 60 * 60 * 1000);

  if (ageMs <= freshLimit) return { label: 'Healthy', note: `Fresh sample ${humanizeDuration(ageMs)} ago`, tone: 'good' };
  if (ageMs <= delayedLimit) return { label: 'Delayed', note: `Latest sample ${humanizeDuration(ageMs)} ago`, tone: 'warn' };
  if (ageMs > 24 * 60 * 60 * 1000) return { label: 'Historical', note: `Newest sample is ${humanizeDuration(ageMs)} old`, tone: 'warn' };
  return { label: 'Stale', note: `No new sample for ${humanizeDuration(ageMs)}`, tone: 'bad' };
}

function computeSensorHealth() {
  if (state.rows.length === 0) {
    return { label: 'No data', note: 'Need samples before checking sensors', tone: 'neutral' };
  }

  const latest = state.rows[0];
  const outOfRange = [];
  const ranges = [
    ['Temperature', latest.temp, -10, 80],
    ['Humidity', latest.hum, 0, 100],
    ['CO2', latest.co2, 250, 5000],
    ['Voltage', latest.volt, 0, 260],
    ['Current', latest.curr, 0, 100],
    ['Power', latest.pwr, 0, 15000],
  ];

  ranges.forEach(([label, value, min, max]) => {
    if (value < min || value > max) outOfRange.push(label);
  });

  if (outOfRange.length) {
    return { label: 'Critical', note: `Out-of-range reading: ${outOfRange.join(', ')}`, tone: 'bad' };
  }

  const recent = state.rows.slice(0, 6);
  const stableMetrics = [
    ['PM2.5', recent.map(row => row.pm25.toFixed(1))],
    ['PM10', recent.map(row => row.pm10.toFixed(1))],
    ['Temp', recent.map(row => row.temp.toFixed(1))],
    ['Hum', recent.map(row => String(row.hum))],
    ['CO2', recent.map(row => String(row.co2))],
  ].filter(([, values]) => values.length >= 4 && new Set(values).size === 1)
    .map(([label]) => label);

  if (stableMetrics.length >= 3) {
    return { label: 'Watch', note: `Repeated values on ${stableMetrics.join(', ')}`, tone: 'warn' };
  }

  return { label: 'Healthy', note: 'Recent sensor values look consistent', tone: 'good' };
}

function renderHealthStatus() {
  const feed = computeFeedHealth();
  const sensor = computeSensorHealth();

  setText(DOM.statusFeedHealth(), feed.label);
  setText(DOM.statusFeedNote(), feed.note);
  setStatusTone(DOM.statusFeedHealth(), DOM.statusFeedNote(), feed.tone);

  setText(DOM.statusSensorHealth(), sensor.label);
  setText(DOM.statusSensorNote(), sensor.note);
  setStatusTone(DOM.statusSensorHealth(), DOM.statusSensorNote(), sensor.tone);
}

function bindEvents() {
  // Nav buttons
  document.getElementById('btn-refresh')?.addEventListener('click', manualRefresh);
  document.getElementById('btn-export-pdf')?.addEventListener('click', exportPDF);
  document.getElementById('btn-export')?.addEventListener('click', exportCSV);

  // Section actions
  document.getElementById('btn-scroll-table')?.addEventListener('click', () => {
    document.getElementById('section-table')?.scrollIntoView({ behavior: 'smooth' });
  });
  document.getElementById('btn-clear-alerts')?.addEventListener('click', () => {
    state.alerts = [];
    renderAlerts();
    showToast('Alerts cleared');
  });
  document.getElementById('btn-download-csv')?.addEventListener('click', exportCSV);

  // Table search
  DOM.tableSearch()?.addEventListener('input', recomputeTableView);

  // Table filters
  document.querySelectorAll('.table-filter').forEach(btn => {
    btn.addEventListener('click', () => applyFilter(btn.dataset.filter));
  });

  // Table sort headers
  document.querySelectorAll('thead th[data-sort]').forEach(th => {
    th.addEventListener('click', () => applySort(th.dataset.sort));
  });

  // Table pagination
  document.getElementById('btn-prev')?.addEventListener('click', () => {
    if (state.table.currentPage > 1) { state.table.currentPage--; renderTable(); }
  });
  document.getElementById('btn-next')?.addEventListener('click', () => {
    const total = Math.ceil(state.filteredRows.length / CONFIG.rowsPerPage);
    if (state.table.currentPage < total) { state.table.currentPage++; renderTable(); }
  });

  // Chart tabs
  document.querySelectorAll('.chart-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.chart-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      applyChartRange(tab.dataset.range);
    });
  });

  // Modal: close on backdrop click or [data-close] button
  document.querySelectorAll('.modal-backdrop').forEach(backdrop => {
    backdrop.addEventListener('click', e => { if (e.target === backdrop) backdrop.classList.remove('open'); });
  });
  document.querySelectorAll('[data-close]').forEach(btn => {
    btn.addEventListener('click', () => closeModal(btn.dataset.close));
  });

  // Logo scroll-to-top
  document.querySelector('.nav-logo')?.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
}

// ════════════════════════════════════════════════════════════
//  T. APP INITIALIZATION
// ════════════════════════════════════════════════════════════
function renderStatusStrip() {
  const c = state.current;
  const p = state.previous;
  const trendDiff = (c.pm25 - p.pm25) + (c.pm10 - p.pm10);
  const trend = trendDiff < -0.5 ? 'Improving' : trendDiff > 0.5 ? 'Worsening' : 'Stable';
  setText(DOM.statusTrend(), trend);
  renderHealthStatus();
}

function filterRows(rows) {
  const { filterStatus } = state.table;
  const query = DOM.tableSearch() ? DOM.tableSearch().value.toLowerCase().trim() : '';
  return rows.filter(r => {
    if (!matchesSelectedRange(r)) return false;
    if (filterStatus !== 'all' && r.status !== filterStatus) return false;
    if (query) {
      const haystack = `${formatTime(r.time)} ${r.pm25} ${r.pm10} ${r.temp} ${r.hum} ${r.co2} ${r.volt} ${r.curr} ${r.pwr} ${r.energy} ${r.status}`.toLowerCase();
      if (!haystack.includes(query)) return false;
    }
    return true;
  });
}

function recomputeTableView() {
  state.table.currentPage = 1;
  state.filteredRows = getFilteredAndSortedRows(state.rows);
  syncRangeControls();
  setText(DOM.rangeSummary(), getRangeSummaryText(state.filteredRows.length));
  renderTable();
}

function renderDashboard() {
  renderMetricCards();
  renderDeltaBadges();
  renderAqiBanner();
  renderInsightPanel();
  renderStatusStrip();
}

function exportCSV() {
  const rows = getExportRows();
  if (rows.length === 0) { showToast('No data to export for the selected range'); return; }
  const header = 'Time,PM2.5 (ug/m3),PM10 (ug/m3),Temperature (C),Humidity (%),CO2 (ppm),Voltage (V),Current (A),Power (W),Energy (kWh),Status\n';
  const body = rows.map(r => {
    const t = r.time instanceof Date && !isNaN(r.time) ? r.time.toLocaleString('en-GB') : '-';
    return `${t},${fmt(r.pm25)},${fmt(r.pm10)},${fmt(r.temp)},${r.hum},${r.co2},${fmt(r.volt)},${fmt(r.curr, 2)},${fmt(r.pwr, 1)},${fmt(r.energy, 2)},${r.status}`;
  }).join('\n');
  const blobUrl = URL.createObjectURL(new Blob([header + body], { type: 'text/csv' }));
  const a = document.createElement('a');
  a.href = blobUrl;
  a.download = `atmosfera_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(blobUrl);
  showToast(`Exported ${rows.length} records`);
}

function exportPDF() {
  const rows = getExportRows();
  if (rows.length === 0) { showToast('No data to export for the selected range'); return; }

  const jsPDFCtor = window.jspdf?.jsPDF;
  if (!jsPDFCtor) {
    showToast('PDF export is unavailable');
    return;
  }

  const doc = new jsPDFCtor({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 12;
  const rowHeight = 6;
  const rangeLabel = getSelectedRangeLabel();
  const columns = [
    { label: 'Time',    width: 28, align: 'left'  },
    { label: 'PM2.5',   width: 18, align: 'right' },
    { label: 'PM10',    width: 18, align: 'right' },
    { label: 'Temp C',  width: 18, align: 'right' },
    { label: 'Hum %',   width: 16, align: 'right' },
    { label: 'CO2',     width: 18, align: 'right' },
    { label: 'Volt V',  width: 18, align: 'right' },
    { label: 'Curr A',  width: 18, align: 'right' },
    { label: 'Power W', width: 20, align: 'right' },
    { label: 'Energy',  width: 20, align: 'right' },
    { label: 'Status',  width: 24, align: 'left'  },
  ];

  const drawTableHeader = y => {
    let x = margin;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setFillColor(240, 244, 247);
    doc.rect(margin, y - 4.5, pageWidth - (margin * 2), rowHeight, 'F');
    columns.forEach(col => {
      const textX = col.align === 'right' ? x + col.width - 1 : x + 1;
      doc.text(col.label, textX, y, { align: col.align });
      x += col.width;
    });
    doc.setDrawColor(210, 218, 225);
    doc.line(margin, y + 1.5, pageWidth - margin, y + 1.5);
  };

  const drawPageHeader = pageNo => {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text('Atmosfera Monitoring Report', margin, margin);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.text(`Generated: ${formatDateTime(new Date())}`, margin, margin + 6);
    doc.text(`Records: ${rows.length}`, margin, margin + 11);
    doc.text(`Range: ${rangeLabel}`, margin, margin + 16);
    doc.text(`Page: ${pageNo}`, pageWidth - margin, margin + 6, { align: 'right' });

    if (pageNo === 1) {
      const c = state.current;
      const snapshot = [
        `PM2.5 ${fmt(c.pm25)}`,
        `PM10 ${fmt(c.pm10)}`,
        `Temp ${fmt(c.temp)} C`,
        `Hum ${fmt(c.hum, 0)} %`,
        `CO2 ${fmt(c.co2, 0)}`,
        `Power ${fmt(c.pwr, 1)} W`,
        `Alerts ${state.alerts.length}`,
      ].join('  |  ');
      doc.text(snapshot, margin, margin + 22);
    }
  };

  const drawRow = (row, y) => {
    const values = [
      formatDateTime(row.time),
      fmt(row.pm25),
      fmt(row.pm10),
      fmt(row.temp),
      `${row.hum}`,
      `${row.co2}`,
      fmt(row.volt),
      fmt(row.curr, 2),
      fmt(row.pwr, 1),
      fmt(row.energy, 2),
      row.status,
    ];

    let x = margin;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    values.forEach((value, idx) => {
      const col = columns[idx];
      const textX = col.align === 'right' ? x + col.width - 1 : x + 1;
      doc.text(String(value), textX, y, { align: col.align });
      x += col.width;
    });
    doc.setDrawColor(232, 237, 241);
    doc.line(margin, y + 1.5, pageWidth - margin, y + 1.5);
  };

  let pageNo = 1;
  let y = margin;
  drawPageHeader(pageNo);
  y = 41;
  drawTableHeader(y);
  y += 8;

  rows.forEach(row => {
    if (y > pageHeight - margin) {
      doc.addPage();
      pageNo += 1;
      drawPageHeader(pageNo);
      y = 24;
      drawTableHeader(y);
      y += 8;
    }
    drawRow(row, y);
    y += rowHeight;
  });

  doc.save(`atmosfera_${new Date().toISOString().slice(0, 10)}.pdf`);
  showToast(`Exported ${rows.length} records to PDF`);
}

function saveThresholds() {
  const read = (id, fallback) => {
    const value = parseFloat(document.getElementById(id)?.value);
    return Number.isFinite(value) ? value : fallback;
  };
  state.thresholds.pm25 = read('thr-pm25', 35);
  state.thresholds.pm10 = read('thr-pm10', 50);
  state.thresholds.co2  = read('thr-co2',  600);
  state.thresholds.temp = read('thr-temp', 35);
  state.thresholds.hum  = read('thr-hum',  85);
  state.thresholds.pwr  = read('thr-pwr',  3000);
  closeModal('modal-thresholds');
  renderDashboard();
  recomputeTableView();
  updateCharts();
  showToast('Thresholds updated');
}

function bindEvents() {
  document.getElementById('btn-refresh')?.addEventListener('click', manualRefresh);
  document.getElementById('btn-export-pdf')?.addEventListener('click', exportPDF);
  document.getElementById('btn-export')?.addEventListener('click', exportCSV);

  document.getElementById('btn-scroll-table')?.addEventListener('click', () => {
    document.getElementById('section-table')?.scrollIntoView({ behavior: 'smooth' });
  });
  document.getElementById('btn-thresholds')?.addEventListener('click', () => openModal('modal-thresholds'));
  document.getElementById('btn-download-csv')?.addEventListener('click', exportCSV);

  DOM.tableSearch()?.addEventListener('input', recomputeTableView);
  DOM.rangeFrom()?.addEventListener('change', updateCustomRange);
  DOM.rangeTo()?.addEventListener('change', updateCustomRange);

  document.querySelectorAll('.range-chip').forEach(btn => {
    btn.addEventListener('click', () => applyRangePreset(btn.dataset.rangePreset));
  });

  document.querySelectorAll('.table-filter').forEach(btn => {
    btn.addEventListener('click', () => applyFilter(btn.dataset.filter));
  });

  document.querySelectorAll('thead th[data-sort]').forEach(th => {
    th.addEventListener('click', () => applySort(th.dataset.sort));
  });

  document.getElementById('btn-prev')?.addEventListener('click', () => {
    if (state.table.currentPage > 1) { state.table.currentPage--; renderTable(); }
  });
  document.getElementById('btn-next')?.addEventListener('click', () => {
    const total = Math.ceil(state.filteredRows.length / CONFIG.rowsPerPage);
    if (state.table.currentPage < total) { state.table.currentPage++; renderTable(); }
  });

  document.querySelectorAll('.chart-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.chart-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      applyChartRange(tab.dataset.range);
    });
  });

  document.querySelectorAll('.modal-backdrop').forEach(backdrop => {
    backdrop.addEventListener('click', e => { if (e.target === backdrop) backdrop.classList.remove('open'); });
  });
  document.querySelectorAll('[data-close]').forEach(btn => {
    btn.addEventListener('click', () => closeModal(btn.dataset.close));
  });
  document.getElementById('btn-save-thresholds')?.addEventListener('click', saveThresholds);
  document.querySelector('.nav-logo')?.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
}

function initApp() {
  // 1. Start clock
  tickClock();
  setInterval(tickClock, 1000);

  // 1.1 Align static UI labels with the power unit used by the data feed
  syncPowerUnitUI();
  syncRangeControls();
  setText(DOM.rangeSummary(), getRangeSummaryText(0));

  // 2. Bind all event listeners
  bindEvents();

  // 3. Initialize charts
  initCharts();

  // 4. Show skeleton loading state
  showSkeleton();
  setText(DOM.statusFeed(), 'Connecting…');
  setText(DOM.statusLast(), 'Loading…');

  // 5. First data fetch
  fetchSheet();
}

// Boot
document.addEventListener('DOMContentLoaded', initApp);

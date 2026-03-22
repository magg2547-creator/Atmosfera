'use strict';

const CONFIG = {
  sheetUrl: 'https://script.google.com/macros/s/AKfycbw8cXJF40QNeW-0-IfgTpzZBMl8ydSL9GcyG-8Y0EWo7Aa5wnNzJMwImWIFbGGF9Cu8uw/exec?token=zd;lkfgmhopgmhpzoifgjhz',
  fetchInterval: 60 * 60 * 1000,
  maxRows: 200,
  rowsPerPage: 10,
  chartMaxPoints: 60,
};

const state = {
  current: { pm25: 0, pm10: 0, temp: 0, hum: 0, co2: 0, volt: 0, curr: 0, pwr: 0, energy: 0 },
  previous: { pm25: 0, pm10: 0, temp: 0, hum: 0, co2: 0, volt: 0, curr: 0, pwr: 0, energy: 0 },
  rows: [],
  filteredRows: [],
  table: { sortKey: 'time', sortDir: 1, currentPage: 1 },
  exportRange: { preset: 'all' },
  chartRange: '1h',
  fetch: { isFetching: false, countdown: null },
};

const domIdCache = new Map();
const domQueryCache = new Map();

function byId(id) {
  if (!domIdCache.has(id)) {
    domIdCache.set(id, document.getElementById(id));
  }
  return domIdCache.get(id);
}

function queryAllCached(selector) {
  if (!domQueryCache.has(selector)) {
    domQueryCache.set(selector, [...document.querySelectorAll(selector)]);
  }
  return domQueryCache.get(selector);
}

const DOM = {
  heroDate: () => byId('hero-date'),
  heroClock: () => byId('hero-clock'),
  lastUpdate: () => byId('status-last'),
  aqiScore: () => byId('aqi-score'),
  aqiStatusText: () => byId('aqi-status-text'),
  aqiDot: () => byId('aqi-dot'),
  bannerPm25: () => byId('banner-pm25'),
  bannerPm10: () => byId('banner-pm10'),
  bannerCo2: () => byId('banner-co2'),
  gaugePm25: () => byId('gauge-pm25'),
  gaugePm10: () => byId('gauge-pm10'),
  gaugeCo2: () => byId('gauge-co2'),
  miniTemp: () => byId('mini-temp'),
  miniHum: () => byId('mini-hum'),
  miniPwr: () => byId('mini-pwr'),
  valPm25: () => byId('val-pm25'),
  valPm10: () => byId('val-pm10'),
  valTemp: () => byId('val-temp'),
  valHum: () => byId('val-hum'),
  valCo2: () => byId('val-co2'),
  valVolt: () => byId('val-volt'),
  valCurr: () => byId('val-curr'),
  deltaPm25: () => byId('delta-pm25'),
  deltaPm10: () => byId('delta-pm10'),
  deltaTemp: () => byId('delta-temp'),
  deltaHum: () => byId('delta-hum'),
  deltaCo2: () => byId('delta-co2'),
  deltaVolt: () => byId('delta-volt'),
  deltaCurr: () => byId('delta-curr'),
  insightPwr: () => byId('insight-pwr'),
  insightEnergy: () => byId('insight-energy'),
  insightPf: () => byId('insight-pf'),
  insightEff: () => byId('insight-eff'),
  barPwr: () => byId('bar-pwr'),
  barEnergy: () => byId('bar-energy'),
  barPf: () => byId('bar-pf'),
  tableSearch: () => byId('table-search'),
  tableSortSelect: () => byId('table-sort-select'),
  tableSortDirection: () => byId('table-sort-direction'),
  rangeSummary: () => byId('range-summary'),
  tableBody: () => byId('table-body'),
  pageInfo: () => byId('page-info'),
  pageNumbers: () => byId('page-numbers'),
  btnPrev: () => byId('btn-prev'),
  btnNext: () => byId('btn-next'),
  sortHeaders: () => queryAllCached('thead th[data-sort]'),
  rangeChips: () => queryAllCached('.range-chip'),
  chartTabs: () => queryAllCached('.chart-tab'),
  toast: () => byId('toast'),
  toastMessage: () => byId('toast-message'),
  btnRefresh: () => byId('btn-refresh'),
  skeletons: () => ['sk-pm25', 'sk-pm10', 'sk-temp', 'sk-hum', 'sk-volt', 'sk-curr', 'sk-insight']
    .map(id => byId(id))
    .filter(Boolean),
  errorBanner:     () => byId('error-banner'),
  errorBannerText: () => byId('error-banner-text'),
  btnErrorRetry:   () => byId('btn-error-retry'),
};

const TABLE_SORT_LABELS = Object.freeze({
  time: 'Recorded',
  pm25: 'PM2.5',
  pm10: 'PM10',
  temp: 'Temperature',
  hum: 'Humidity',
  co2: 'CO2',
  volt: 'Voltage',
  curr: 'Current',
  pwr: 'Power',
  energy: 'Energy',
});

const DATE_FORMATTERS = Object.freeze({
  heroClock: new Intl.DateTimeFormat('en-GB', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  }),
  heroDate: new Intl.DateTimeFormat('en-GB', {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
  }),
  timeShort: new Intl.DateTimeFormat('en-GB', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }),
  tableDate: new Intl.DateTimeFormat('en-GB', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }),
});

const charts = { aq: null, co2: null, energy: null, donut: null };

const DONUT_LABEL_PLUGIN = {
  id: 'donutLabel',
  afterDraw(chart) {
    if (chart.canvas.id !== 'chart-donut') return;

    const { ctx, width, height } = chart;
    ctx.save();
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillStyle = 'rgba(255,255,255,.9)';
    ctx.font = '700 22px Sora, sans-serif';       // ขึ้นจาก 18px
    ctx.fillText(Math.round(state.current.co2), width / 2, height / 2 - 10);
    ctx.fillStyle = 'rgba(255,255,255,.45)';
    ctx.font = '400 11px Instrument Sans, sans-serif'; // ขึ้นจาก 10px
    ctx.fillText('ppm', width / 2, height / 2 + 12);
    ctx.restore();
  },
};

const fmt = (value, digits = 1) => {
  const number = Number(value);
  return Number.isFinite(number) ? number.toFixed(digits) : '-';
};

function setHTML(element, html) {
  if (element) element.innerHTML = html;
}

function setText(element, text) {
  if (element) element.textContent = text;
}

function setWidth(element, percent) {
  if (!element) return;
  const clamped = Math.min(100, Math.max(0, percent));
  element.style.width = `${clamped}%`;
}

function parseTimestamp(value) {
  if (!value) return new Date();
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === 'number' && Number.isFinite(value)) {
    const numericDate = new Date(value);
    if (!Number.isNaN(numericDate.getTime())) return numericDate;
  }

  const raw = String(value).trim();
  const normalized = raw.replace(' ', 'T');
  const directDate = new Date(normalized);
  if (!Number.isNaN(directDate.getTime())) return directDate;

  const dateTimeMatch = raw.match(
    /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:[ T](\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?)?$/
  );

  if (dateTimeMatch) {
    let [, first, second, year, hour = '0', minute = '0', secondValue = '0'] = dateTimeMatch;
    let day = Number(first);
    let month = Number(second);
    const fullYear = Number(year.length === 2 ? `20${year}` : year);

    if (day <= 12 && month > 12) {
      day = Number(second);
      month = Number(first);
    }

    const parsedDate = new Date(
      fullYear,
      month - 1,
      day,
      Number(hour),
      Number(minute),
      Number(secondValue)
    );

    if (!Number.isNaN(parsedDate.getTime())) return parsedDate;
  }

  const dateOnlyMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dateOnlyMatch) {
    const [, year, month, day] = dateOnlyMatch;
    const parsedDate = new Date(Number(year), Number(month) - 1, Number(day));
    if (!Number.isNaN(parsedDate.getTime())) return parsedDate;
  }

  return new Date();
}

function normalizePower(value) {
  return Number(value ?? 0);
}

function deltaClass(diff) {
  if (Math.abs(diff) < 0.01) return 'delta-flat';
  return diff > 0 ? 'delta-up' : 'delta-down';
}

function deltaLabel(diff, unit) {
  const prefix = diff > 0 ? '^' : diff < 0 ? 'v' : '=';
  return `${prefix} ${Math.abs(diff).toFixed(1)}${unit}`;
}

function applyDelta(element, diff, unit) {
  if (!element) return;
  element.className = `delta ${deltaClass(diff)}`;
  element.textContent = deltaLabel(diff, unit);
}

function formatTimeShort(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '-';
  return DATE_FORMATTERS.timeShort.format(date);
}

function formatTableDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '-';
  return DATE_FORMATTERS.tableDate.format(date);
}

function formatTableClock(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '-';
  return DATE_FORMATTERS.timeShort.format(date);
}

function formatDateTime(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '-';
  return `${formatTableDate(date)} ${formatTableClock(date)}`;
}

function toCsvCell(value) {
  return `"${String(value ?? '').replace(/"/g, '""')}"`;
}

function calcAqiFromPm25(pm25) {
  const breakpoints = [
    [0,    12,    0,   50],
    [12.1, 35.4,  51,  100],
    [35.5, 55.4,  101, 150],
    [55.5, 150.4, 151, 200],
    [150.5, 250.4, 201, 300],
    [250.5, 350.4, 301, 400],  // เพิ่ม
    [350.5, 500.4, 401, 500],  // เพิ่ม
  ];

  for (const [low, high, aqiLow, aqiHigh] of breakpoints) {
    if (pm25 >= low && pm25 <= high) {
      return Math.round(((aqiHigh - aqiLow) / (high - low)) * (pm25 - low) + aqiLow);
    }
  }

  return 500; // PM2.5 > 500.4 → Hazardous ceiling
}

function calcAqiFromPm10(pm10) {
  const breakpoints = [
    [0, 54, 0, 50],
    [55, 154, 51, 100],
    [155, 254, 101, 150],
    [255, 354, 151, 200],
  ];

  for (const [low, high, aqiLow, aqiHigh] of breakpoints) {
    if (pm10 >= low && pm10 <= high) {
      return Math.round(((aqiHigh - aqiLow) / (high - low)) * (pm10 - low) + aqiLow);
    }
  }

  return Math.round(pm10 * 0.7);
}

function calcAQI() {
  return Math.max(
    calcAqiFromPm25(state.current.pm25),
    calcAqiFromPm10(state.current.pm10)
  );
}

function getAqiMeta(aqi) {
  if (aqi <= 50)  return { label: 'Good',                          dot: '#a8e6c8' };
  if (aqi <= 100) return { label: 'Moderate',                      dot: '#fde68a' };
  if (aqi <= 150) return { label: 'Unhealthy for Sensitive Groups', dot: '#fca5a5' };
  if (aqi <= 200) return { label: 'Unhealthy',                     dot: '#f87171' };
  if (aqi <= 300) return { label: 'Very Unhealthy',                dot: '#c084fc' }; // เพิ่ม
  return          { label: 'Hazardous',                            dot: '#9f1239' }; // เพิ่ม
}

function normalizeRow(raw) {
  return {
    time: parseTimestamp(raw.timestamp),
    pm25: Number(raw.pm25 ?? 0),
    pm10: Number(raw.pm10 ?? 0),
    temp: Number(raw.temperature ?? 0),
    hum: Math.round(Number(raw.humidity ?? 0)),
    co2: Math.round(Number(raw.co2 ?? 0)),
    volt: Number(raw.voltage ?? 0),
    curr: Number(raw.current ?? 0),
    pwr: normalizePower(raw.power),
    energy: Number(raw.energy ?? 0),
  };
}

function normalizeRows(rawRows) {
  return rawRows
    .map(normalizeRow)
    .reverse()
    .slice(0, CONFIG.maxRows);
}

function copyCurrentMetricsFromRow(row) {
  Object.assign(state.previous, state.current);
  state.current.pm25 = row.pm25;
  state.current.pm10 = row.pm10;
  state.current.temp = row.temp;
  state.current.hum = row.hum;
  state.current.co2 = row.co2;
  state.current.volt = row.volt;
  state.current.curr = row.curr;
  state.current.pwr = row.pwr;
  state.current.energy = row.energy;
}

async function fetchSheetData() {
  const response = await fetch(CONFIG.sheetUrl);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);

  const payload = await response.json();
  if (payload.status !== 'ok') throw new Error(payload.message || 'API returned an error');
  if (!Array.isArray(payload.data) || payload.data.length === 0) {
    throw new Error('No records in response');
  }

  return payload.data;
}

function handleFetchError(error) {
  const isFirstLoad = state.rows.length === 0;
  const message = isFirstLoad
    ? `Could not load data — ${error.message}`
    : `Refresh failed — showing last known data (${error.message})`;

  showErrorBanner(message);
  setText(DOM.lastUpdate(), 'Failed');
  showToast(`Warning: ${error.message}`);
  console.error('[Atmosfera] fetch error:', error);
  startCountdown();
}

function showErrorBanner(message) {
  const banner = DOM.errorBanner();
  const text   = DOM.errorBannerText();
  if (text)   text.textContent = message;
  if (banner) banner.hidden = false;
}

function hideErrorBanner() {
  const banner = DOM.errorBanner();
  if (banner) banner.hidden = true;
}

async function fetchSheet() {
  if (state.fetch.isFetching) return;
  state.fetch.isFetching = true;
  setText(DOM.lastUpdate(), 'Loading...');

  try {
    const rawRows        = await fetchSheetData();
    const normalizedRows = normalizeRows(rawRows);

    state.rows = normalizedRows;
    copyCurrentMetricsFromRow(normalizedRows[0]);

    renderDashboard();
    recomputeTableView();
    updateCharts();
    hideErrorBanner(); // ← เพิ่ม

    setText(DOM.lastUpdate(), 'Just now');
    showToast(`Loaded ${rawRows.length} records`);
    startCountdown();
    requestAnimationFrame(() => resizeAllCharts());
  } catch (error) {
    handleFetchError(error);
  } finally {
    state.fetch.isFetching = false;
    hideSkeleton();
  }
}

function renderMetricCards() {
  const current = state.current;
  setHTML(DOM.valPm25(), `${fmt(current.pm25)}<span class="metric-unit">&micro;g/m&sup3;</span>`);
  setHTML(DOM.valPm10(), `${fmt(current.pm10)}<span class="metric-unit">&micro;g/m&sup3;</span>`);
  setHTML(DOM.valTemp(), `${fmt(current.temp)}<span class="metric-unit">C</span>`);
  setHTML(DOM.valHum(), `${fmt(current.hum, 0)}<span class="metric-unit">%</span>`);
  setHTML(DOM.valCo2(), `${fmt(current.co2, 0)}<span class="metric-unit">ppm</span>`);
  setHTML(DOM.valVolt(), `${fmt(current.volt)}<span class="metric-unit">V</span>`);
  setHTML(DOM.valCurr(), `${fmt(current.curr)}<span class="metric-unit">A</span>`);
}

function renderDeltaBadges() {
  const current = state.current;
  const previous = state.previous;

  applyDelta(DOM.deltaPm25(), current.pm25 - previous.pm25, '');
  applyDelta(DOM.deltaPm10(), current.pm10 - previous.pm10, '');
  applyDelta(DOM.deltaTemp(), current.temp - previous.temp, 'C');
  applyDelta(DOM.deltaHum(), current.hum - previous.hum, '%');
  applyDelta(DOM.deltaVolt(), current.volt - previous.volt, 'V');
  applyDelta(DOM.deltaCurr(), current.curr - previous.curr, 'A');

  const co2Delta = DOM.deltaCo2();
  if (co2Delta) {
    co2Delta.className = 'delta';
    co2Delta.style.background = 'rgba(255,255,255,.12)';
    co2Delta.style.color = 'rgba(255,255,255,.7)';
    co2Delta.textContent = deltaLabel(current.co2 - previous.co2, '');
  }
}

function renderAqiBanner() {
  const current = state.current;
  const aqi = calcAQI();
  const meta = getAqiMeta(aqi);

  setText(DOM.aqiScore(), aqi);
  setText(DOM.aqiStatusText(), meta.label);

  const dot = DOM.aqiDot();
  if (dot) dot.style.background = meta.dot;

  setText(DOM.bannerPm25(), `${fmt(current.pm25)} \u00b5g/m\u00b3`);
  setText(DOM.bannerPm10(), `${fmt(current.pm10)} \u00b5g/m\u00b3`);
  setText(DOM.bannerCo2(), `${fmt(current.co2, 0)} ppm`);
  setHTML(DOM.miniTemp(), `${fmt(current.temp)}<span class="aqi-mini-unit">C</span>`);
  setHTML(DOM.miniHum(), `${fmt(current.hum, 0)}<span class="aqi-mini-unit">%</span>`);
  setHTML(DOM.miniPwr(), `${fmt(current.pwr, 1)}<span class="aqi-mini-unit">W</span>`);

  setWidth(DOM.gaugePm25(), (current.pm25 / 200) * 100);
  setWidth(DOM.gaugePm10(), (current.pm10 / 400) * 100);
  setWidth(DOM.gaugeCo2(), ((current.co2 - 350) / 4650) * 100);
}

function renderInsightPanel() {
  const current = state.current;
  const powerReference = 3000;
  const energyReference = 30;
  const powerFactor = current.volt > 0 && current.curr > 0
    ? Math.min(1, current.pwr / (current.volt * current.curr))
    : 0;
  const efficiency = powerFactor >= 0.95 ? 'A+' : powerFactor >= 0.9 ? 'A' : powerFactor >= 0.85 ? 'B+' : 'B';
  const efficiencyColor = powerFactor >= 0.95
    ? 'var(--tertiary)'
    : powerFactor >= 0.9
      ? 'var(--secondary)'
      : 'var(--primary)';

  setHTML(DOM.insightPwr(), `${fmt(current.pwr, 1)} <span style="font-size:.85rem;font-weight:400;opacity:.5">W</span>`);
  setHTML(DOM.insightEnergy(), `${fmt(current.energy, 2)} <span style="font-size:.85rem;font-weight:400;opacity:.5">kWh</span>`);
  setHTML(DOM.insightPf(), `${powerFactor.toFixed(2)} <span style="font-size:.85rem;font-weight:400;opacity:.5">PF</span>`);
  setText(DOM.insightEff(), efficiency);

  const insightEff = DOM.insightEff();
  if (insightEff) insightEff.style.color = efficiencyColor;

  setWidth(DOM.barPwr(), (current.pwr / powerReference) * 100);
  setWidth(DOM.barEnergy(), (current.energy / energyReference) * 100);
  setWidth(DOM.barPf(), powerFactor * 100);

  if (charts.donut) {
    const percent = Math.min(100, ((current.co2 - 350) / 4650) * 100);
    charts.donut.data.datasets[0].data = [percent, 100 - percent];
    charts.donut.update('none');
  }
}

function renderDashboard() {
  renderMetricCards();
  renderDeltaBadges();
  renderAqiBanner();
  renderInsightPanel();
}

function sortRows(rows) {
  const { sortKey, sortDir } = state.table;

  return [...rows].sort((left, right) => {
    const leftValue = sortKey === 'time' ? left.time.getTime() : left[sortKey];
    const rightValue = sortKey === 'time' ? right.time.getTime() : right[sortKey];

    if (leftValue === rightValue) return 0;
    return sortDir === 1
      ? (leftValue < rightValue ? 1 : -1)
      : (leftValue < rightValue ? -1 : 1);
  });
}

function getSelectedRangeBounds() {
  if (state.exportRange.preset === 'all') return null;

  const now = new Date();
  const from = new Date(now);
  from.setHours(0, 0, 0, 0);

  const to = new Date(now);
  to.setHours(23, 59, 59, 999);

  if (state.exportRange.preset === 'today') {
    return { from, to };
  }

  if (state.exportRange.preset === '7d') {
    from.setDate(from.getDate() - 6);
    return { from, to };
  }

  if (state.exportRange.preset === '30d') {
    from.setDate(from.getDate() - 29);
    return { from, to };
  }

  return null;
}

function getSelectedRangeLabel() {
  if (state.exportRange.preset === 'all') return 'All records';
  if (state.exportRange.preset === 'today') return 'Today';
  if (state.exportRange.preset === '7d') return 'Last 7 days';
  if (state.exportRange.preset === '30d') return 'Last 30 days';
  return 'All records';
}

function matchesSelectedRange(row) {
  const bounds = getSelectedRangeBounds();
  if (!bounds) return true;
  return row.time >= bounds.from && row.time <= bounds.to;
}

function filterRows(rows) {
  const query = DOM.tableSearch()?.value.toLowerCase().trim() || '';

  return rows.filter(row => {
    if (!matchesSelectedRange(row)) return false;

    if (!query) return true;

    const haystack = [
      formatDateTime(row.time),
      row.pm25,
      row.pm10,
      row.temp,
      row.hum,
      row.co2,
      row.volt,
      row.curr,
      row.pwr,
      row.energy,
    ].join(' ').toLowerCase();

    return haystack.includes(query);
  });
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

function getSortDirectionLabel() {
  if (state.table.sortKey === 'time') {
    return state.table.sortDir === 1 ? 'Newest first' : 'Oldest first';
  }
  return state.table.sortDir === 1 ? 'High to low' : 'Low to high';
}

function syncMobileSortControls() {
  const sortSelect = DOM.tableSortSelect();
  if (sortSelect) sortSelect.value = state.table.sortKey;

  const sortDirection = DOM.tableSortDirection();
  if (!sortDirection) return;

  const label = getSortDirectionLabel();
  const fieldName = TABLE_SORT_LABELS[state.table.sortKey] || 'records';
  sortDirection.textContent = label;
  sortDirection.dataset.sortDir = state.table.sortDir === 1 ? 'desc' : 'asc';
  sortDirection.setAttribute('aria-label', `Sort ${fieldName} ${label.toLowerCase()}`);
}

function syncRangeControls() {
  DOM.rangeChips().forEach(chip => {
    chip.classList.toggle('active', chip.dataset.rangePreset === state.exportRange.preset);
  });
}

function syncSortIndicators() {
  DOM.sortHeaders().forEach(header => {
    header.classList.remove('sorted');
    delete header.dataset.sortDir;
  });

  syncMobileSortControls();

  const activeHeader = byId(`th-${state.table.sortKey}`);
  if (!activeHeader) return;

  activeHeader.classList.add('sorted');
  activeHeader.dataset.sortDir = state.table.sortDir === 1 ? 'desc' : 'asc';
}

function recomputeTableView() {
  state.table.currentPage = 1;
  state.filteredRows = getFilteredAndSortedRows(state.rows);
  syncRangeControls();
  syncSortIndicators();
  setText(DOM.rangeSummary(), getRangeSummaryText(state.filteredRows.length));
  renderTable();
}

function applySort(key) {
  if (!TABLE_SORT_LABELS[key]) return;

  if (state.table.sortKey === key) {
    state.table.sortDir *= -1;
  } else {
    state.table.sortKey = key;
    state.table.sortDir = 1;
  }

  recomputeTableView();
}

function applySortKey(key) {
  if (!TABLE_SORT_LABELS[key]) return;

  if (state.table.sortKey === key) {
    syncMobileSortControls();
    return;
  }

  state.table.sortKey = key;
  state.table.sortDir = 1;
  recomputeTableView();
}

function toggleTableSortDirection() {
  state.table.sortDir *= -1;
  recomputeTableView();
}

function applyRangePreset(preset) {
  if (!['all', 'today', '7d', '30d'].includes(preset)) return;
  state.exportRange.preset = preset;
  recomputeTableView();
}

function renderTable() {
  const { currentPage } = state.table;
  const perPage = CONFIG.rowsPerPage;
  const start = (currentPage - 1) * perPage;
  const end = start + perPage;
  const pageRows = state.filteredRows.slice(start, end);
  const total = state.filteredRows.length;
  const totalPages = Math.max(1, Math.ceil(total / perPage));
  const tbody = DOM.tableBody();

  if (!tbody) return;

  if (pageRows.length === 0) {
    tbody.innerHTML = '<tr class="table-empty-state"><td colspan="10" class="no-results">No matching records found</td></tr>';
  } else {
    tbody.innerHTML = pageRows.map(row => `
      <tr>
        <td class="td-time" data-label="Recorded">
          <span class="td-date">${formatTableDate(row.time)}</span>
          <span class="td-clock">${formatTableClock(row.time)}</span>
        </td>
        <td class="td-value" data-label="PM2.5">${fmt(row.pm25)}</td>
        <td class="td-value" data-label="PM10">${fmt(row.pm10)}</td>
        <td class="td-value" data-label="Temperature">${fmt(row.temp)}</td>
        <td class="td-value" data-label="Humidity">${row.hum}</td>
        <td class="td-value" data-label="CO2">${row.co2}</td>
        <td class="td-value" data-label="Voltage">${fmt(row.volt)}</td>
        <td class="td-value" data-label="Current">${fmt(row.curr, 2)}</td>
        <td class="td-value" data-label="Power">${fmt(row.pwr, 1)}</td>
        <td class="td-value" data-label="Energy">${fmt(row.energy, 2)}</td>
      </tr>
    `).join('');
  }

  const rangeStart = total === 0 ? 0 : start + 1;
  const rangeEnd = Math.min(end, total);
  setText(DOM.pageInfo(), `Showing ${rangeStart}-${rangeEnd} of ${total} records`);

  const prevButton = DOM.btnPrev();
  const nextButton = DOM.btnNext();
  if (prevButton) prevButton.disabled = currentPage <= 1;
  if (nextButton) nextButton.disabled = currentPage >= totalPages;

  const pageNumbers = DOM.pageNumbers();
  if (!pageNumbers) return;

  const fragment = document.createDocumentFragment();

  for (let pageNumber = 1; pageNumber <= totalPages; pageNumber += 1) {
    if (totalPages <= 7 || Math.abs(pageNumber - currentPage) < 3 || pageNumber === 1 || pageNumber === totalPages) {
      const button = document.createElement('button');
      button.className = `page-btn${pageNumber === currentPage ? ' active' : ''}`;
      button.textContent = pageNumber;
      button.dataset.page = pageNumber;
      fragment.appendChild(button);
    }
  }

  pageNumbers.replaceChildren(fragment);
}

function makeGradient(context, topColor, bottomColor, height = 220) {
  const gradient = context.createLinearGradient(0, 0, 0, height);
  gradient.addColorStop(0, topColor);
  gradient.addColorStop(1, bottomColor);
  return gradient;
}

function initCharts() {
  Chart.register(DONUT_LABEL_PLUGIN);
  Chart.defaults.font.family = "'Instrument Sans', sans-serif";
  Chart.defaults.color = '#5a6a72';

  const aqContext = byId('chart-aq').getContext('2d');
  charts.aq = new Chart(aqContext, {
    type: 'line',
    data: {
      labels: [],
      datasets: [
        {
          label: 'PM 2.5 (\u00b5g/m\u00b3)',
          data: [],
          borderColor: '#466370',
          borderWidth: 2,
          pointRadius: 0,
          fill: true,
          backgroundColor: makeGradient(aqContext, 'rgba(70,99,112,.15)', 'rgba(70,99,112,0)'),
          tension: 0.45,
        },
        {
          label: 'PM 10 (\u00b5g/m\u00b3)',
          data: [],
          borderColor: '#4a9da8',
          borderWidth: 2,
          pointRadius: 0,
          fill: true,
          backgroundColor: makeGradient(aqContext, 'rgba(74,157,168,.12)', 'rgba(74,157,168,0)'),
          tension: 0.45,
        },
        {
          label: 'WHO Limit',
          data: [],
          borderColor: 'rgba(93,94,97,.25)',
          borderWidth: 1.5,
          borderDash: [6, 4],
          pointRadius: 0,
          fill: false,
          tension: 0,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#fff',
          titleColor: '#2a3439',
          bodyColor: '#5a6a72',
          borderColor: 'rgba(90,106,114,.12)',
          borderWidth: 1,
          padding: 12,
          cornerRadius: 12,
          callbacks: {
            label: context => `${context.dataset.label}: ${Number(context.parsed.y).toFixed(1)}`,
          },
        },
      },
      scales: {
        x: {
          grid: { color: 'rgba(90,106,114,.07)', drawBorder: false },
          ticks: { maxTicksLimit: 6 },
        },
        y: {
          grid: { color: 'rgba(90,106,114,.07)', drawBorder: false },
          ticks: { maxTicksLimit: 5 },
          min: 0,
          title: {
            display: true,
            text: '\u00b5g/m\u00b3',
            color: '#5a6a72',
            font: { size: 10 },
          },
        },
      },
      animation: { duration: 200 },
    },
  });

  const co2Context = byId('chart-co2').getContext('2d');
  charts.co2 = new Chart(co2Context, {
    type: 'line',
    data: {
      labels: [],
      datasets: [
        {
          label: 'CO2 (ppm)',
          data: [],
          borderColor: '#006977',
          borderWidth: 2,
          pointRadius: 0,
          fill: true,
          backgroundColor: makeGradient(co2Context, 'rgba(0,105,119,.15)', 'rgba(0,105,119,0)'),
          tension: 0.45,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#fff',
          titleColor: '#2a3439',
          bodyColor: '#5a6a72',
          borderColor: 'rgba(90,106,114,.12)',
          borderWidth: 1,
          padding: 12,
          cornerRadius: 12,
          callbacks: {
            label: context => `CO2: ${Math.round(context.parsed.y)} ppm`,
          },
        },
      },
      scales: {
        x: {
          grid: { color: 'rgba(0,0,0,.04)', drawBorder: false },
          ticks: { maxTicksLimit: 5 },
        },
        y: {
          grid: { color: 'rgba(0,0,0,.04)', drawBorder: false },
          ticks: { maxTicksLimit: 5 },
          title: {
            display: true,
            text: 'ppm',
            color: '#5a6a72',
            font: { size: 10 },
          },
        },
      },
      animation: { duration: 200 },
    },
  });

  const energyContext = byId('chart-energy').getContext('2d');
  charts.energy = new Chart(energyContext, {
    type: 'line',
    data: {
      labels: [],
      datasets: [
        {
          label: 'Voltage (V)',
          data: [],
          borderColor: '#5d5e61',
          borderWidth: 2,
          pointRadius: 0,
          fill: false,
          tension: 0.4,
        },
        {
          label: 'Current (A)',
          data: [],
          borderColor: '#466370',
          borderWidth: 2,
          pointRadius: 0,
          fill: false,
          tension: 0.4,
        },
        {
          label: 'Power (W)',
          data: [],
          borderColor: '#006977',
          borderWidth: 2,
          pointRadius: 0,
          fill: true,
          backgroundColor: makeGradient(energyContext, 'rgba(0,105,119,.08)', 'rgba(0,105,119,0)', 200),
          tension: 0.4,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#fff',
          titleColor: '#2a3439',
          bodyColor: '#5a6a72',
          borderColor: 'rgba(90,106,114,.12)',
          borderWidth: 1,
          padding: 12,
          cornerRadius: 12,
          callbacks: {
            label: context => {
              const value = context.parsed.y;
              if (context.datasetIndex === 0) return `Voltage: ${Number(value).toFixed(1)} V`;
              if (context.datasetIndex === 1) return `Current: ${Number(value).toFixed(2)} A`;
              return `Power: ${Number(value).toFixed(1)} W`;
            },
          },
        },
      },
      scales: {
        x: {
          grid: { color: 'rgba(0,0,0,.04)', drawBorder: false },
          ticks: { maxTicksLimit: 6 },
        },
        y: {
          grid: { color: 'rgba(0,0,0,.04)', drawBorder: false },
          ticks: { maxTicksLimit: 5 },
        },
      },
      animation: { duration: 200 },
    },
  });

  const donutContext = byId('chart-donut').getContext('2d');
  charts.donut = new Chart(donutContext, {
    type: 'doughnut',
    data: {
      datasets: [
        {
          data: [0, 100],
          backgroundColor: ['rgba(255,255,255,.7)', 'rgba(255,255,255,.12)'],
          borderWidth: 0,
          hoverOffset: 4,
        },
      ],
    },
    options: {
      responsive: false,
      cutout: '68',
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false },
      },
      animation: { duration: 600 },
    },
  });
}

function getRowsForChartRange(rangeKey) {
  const rows = [...state.rows].reverse().slice(-CONFIG.chartMaxPoints);
  if (rows.length === 0 || rangeKey === '24h') return rows;

  const latest = rows[rows.length - 1].time;
  if (!(latest instanceof Date) || Number.isNaN(latest.getTime())) return rows;

  const rangeMs = rangeKey === '1h' ? 60 * 60 * 1000 : 6 * 60 * 60 * 1000;
  return rows.filter(row => latest.getTime() - row.time.getTime() <= rangeMs);
}

function applyChartRange(rangeKey) {
  state.chartRange = rangeKey;

  DOM.chartTabs().forEach(tab => {
    tab.classList.toggle('active', tab.dataset.range === rangeKey);
  });

  const rows = getRowsForChartRange(rangeKey);
  const labels = rows.map(row => formatTimeShort(row.time));

  if (charts.aq) {
    charts.aq.data.labels = labels;
    charts.aq.data.datasets[0].data = rows.map(row => row.pm25);
    charts.aq.data.datasets[1].data = rows.map(row => row.pm10);
    charts.aq.data.datasets[2].data = Array(rows.length).fill(15);
    charts.aq.update();
  }

  if (charts.co2) {
    charts.co2.data.labels = labels;
    charts.co2.data.datasets[0].data = rows.map(row => row.co2);
    charts.co2.update();
  }

  if (charts.energy) {
    charts.energy.data.labels = labels;
    charts.energy.data.datasets[0].data = rows.map(row => row.volt);
    charts.energy.data.datasets[1].data = rows.map(row => row.curr);
    charts.energy.data.datasets[2].data = rows.map(row => row.pwr);
    charts.energy.update();
  }
}

function updateCharts() {
  applyChartRange(state.chartRange);
}

function startCountdown() {
  clearInterval(state.fetch.countdown);
  let remainingSeconds = CONFIG.fetchInterval / 1000;

  state.fetch.countdown = setInterval(() => {
    remainingSeconds -= 1;

    const minutes = Math.floor(remainingSeconds / 60);
    const seconds = remainingSeconds % 60;
    setText(DOM.lastUpdate(), `Next in ${minutes}:${String(seconds).padStart(2, '0')}`);

    if (remainingSeconds <= 0) {
      clearInterval(state.fetch.countdown);
      fetchSheet();
    }
  }, 1000);
}

function showSkeleton() {
  DOM.skeletons().forEach(element => {
    const pendingTimer = skeletonHideTimers.get(element);
    if (pendingTimer) {
      clearTimeout(pendingTimer);
      skeletonHideTimers.delete(element);
    }

    element.style.transition = '';
    element.style.opacity = '1';
    element.parentElement?.classList.add('is-loading');
    element.classList.add('visible');
    if (element.id === 'sk-insight') element.classList.add('flex');
  });
}

function hideSkeleton() {
  DOM.skeletons().forEach(element => {
    const pendingTimer = skeletonHideTimers.get(element);
    if (pendingTimer) clearTimeout(pendingTimer);

    element.style.transition = 'opacity .3s ease';
    element.style.opacity = '0';

    const timer = setTimeout(() => {
      element.parentElement?.classList.remove('is-loading');
      element.classList.remove('visible', 'flex');
      element.style.opacity = '';
      element.style.transition = '';
      skeletonHideTimers.delete(element);
    }, 300);

    skeletonHideTimers.set(element, timer);
  });
}

let toastTimer = null;
const skeletonHideTimers = new Map();

function showToast(message) {
  setText(DOM.toastMessage(), message);
  const toast = DOM.toast();
  if (toast) toast.classList.add('visible');

  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    if (toast) toast.classList.remove('visible');
  }, 3000);
}

function getExportRowValues(row) {
  return [
    formatDateTime(row.time),
    fmt(row.pm25),
    fmt(row.pm10),
    fmt(row.temp),
    String(row.hum),
    String(row.co2),
    fmt(row.volt),
    fmt(row.curr, 2),
    fmt(row.pwr, 1),
    fmt(row.energy, 2),
  ];
}

function exportAllCSV() {
  const rows = [...state.rows].sort((a, b) => b.time - a.time);
  if (rows.length === 0) {
    showToast('No data to export');
    return;
  }

  const headers = [
    'Time', 'PM2.5 (\u00b5g/m\u00b3)', 'PM10 (\u00b5g/m\u00b3)',
    'Temperature (C)', 'Humidity (%)', 'CO2 (ppm)',
    'Voltage (V)', 'Current (A)', 'Power (W)', 'Energy (kWh)',
  ];

  const csvRows = [
    headers.map(toCsvCell).join(','),
    ...rows.map(row => getExportRowValues(row).map(toCsvCell).join(',')),
  ];

  const blobUrl = URL.createObjectURL(
    new Blob([csvRows.join('\n')], { type: 'text/csv' })
  );
  const link = document.createElement('a');
  link.href = blobUrl;
  link.download = `atmosfera_all_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(blobUrl);

  showToast(`Exported ${rows.length} records`);
}

function exportPDF(rows) {
  if (!Array.isArray(rows)) rows = getExportRows();
  if (rows.length === 0) {
    showToast('No data to export for the selected range');
    return;
  }

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
    { label: 'Time', width: 28, align: 'left' },
    { label: 'PM2.5', width: 18, align: 'right' },
    { label: 'PM10', width: 18, align: 'right' },
    { label: 'Temp C', width: 18, align: 'right' },
    { label: 'Hum %', width: 16, align: 'right' },
    { label: 'CO2', width: 18, align: 'right' },
    { label: 'Volt V', width: 18, align: 'right' },
    { label: 'Curr A', width: 18, align: 'right' },
    { label: 'Power W', width: 20, align: 'right' },
    { label: 'Energy', width: 20, align: 'right' },
  ];

  const drawTableHeader = y => {
    let x = margin;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setFillColor(240, 244, 247);
    doc.rect(margin, y - 4.5, pageWidth - (margin * 2), rowHeight, 'F');

    columns.forEach(column => {
      const textX = column.align === 'right' ? x + column.width - 1 : x + 1;
      doc.text(column.label, textX, y, { align: column.align });
      x += column.width;
    });

    doc.setDrawColor(210, 218, 225);
    doc.line(margin, y + 1.5, pageWidth - margin, y + 1.5);
  };

  const drawPageHeader = pageNumber => {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text('Atmosfera Monitoring Report', margin, margin);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.text(`Generated: ${formatDateTime(new Date())}`, margin, margin + 6);
    doc.text(`Records: ${rows.length}`, margin, margin + 11);
    doc.text(`Range: ${rangeLabel}`, margin, margin + 16);
    doc.text(`Page: ${pageNumber}`, pageWidth - margin, margin + 6, { align: 'right' });

    if (pageNumber === 1) {
      const current = state.current;
      const snapshot = [
        `PM2.5 ${fmt(current.pm25)}`,
        `PM10 ${fmt(current.pm10)}`,
        `Temp ${fmt(current.temp)} C`,
        `Hum ${fmt(current.hum, 0)} %`,
        `CO2 ${fmt(current.co2, 0)}`,
        `Power ${fmt(current.pwr, 1)} W`,
      ].join('  |  ');

      doc.text(snapshot, margin, margin + 22);
    }
  };

  const drawRow = (row, y) => {
    const values = getExportRowValues(row);
    let x = margin;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);

    values.forEach((value, index) => {
      const column = columns[index];
      const textX = column.align === 'right' ? x + column.width - 1 : x + 1;
      doc.text(String(value), textX, y, { align: column.align });
      x += column.width;
    });

    doc.setDrawColor(232, 237, 241);
    doc.line(margin, y + 1.5, pageWidth - margin, y + 1.5);
  };

  let pageNumber = 1;
  let y = margin;
  drawPageHeader(pageNumber);
  y = 41;
  drawTableHeader(y);
  y += 8;

  rows.forEach(row => {
    if (y > pageHeight - margin) {
      doc.addPage();
      pageNumber += 1;
      drawPageHeader(pageNumber);
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

function tickClock() {
  const now = new Date();
  setText(DOM.heroClock(), DATE_FORMATTERS.heroClock.format(now));
  setText(DOM.heroDate(), DATE_FORMATTERS.heroDate.format(now));
}

function manualRefresh() {
  if (state.fetch.isFetching) return;

  const button = DOM.btnRefresh();
  const icon = button?.querySelector('svg');

  if (button) button.disabled = true;
  if (icon) icon.style.animation = 'spin 0.8s linear infinite';

  clearInterval(state.fetch.countdown);
  setText(DOM.lastUpdate(), 'Refreshing...');
  showSkeleton();

  fetchSheet().finally(() => {
    if (icon) icon.style.animation = '';
    if (button) button.disabled = false;
  });
}

function bindEvents() {
  DOM.btnRefresh()?.addEventListener('click', manualRefresh);
  byId('btn-export-pdf')?.addEventListener('click', openPdfModal);
  byId('btn-export')?.addEventListener('click', exportAllCSV);
  
  // modal
  byId('btn-modal-close')?.addEventListener('click', closePdfModal);
  byId('btn-modal-cancel')?.addEventListener('click', closePdfModal);
  byId('modal-pdf')?.addEventListener('click', e => {
    if (e.target === e.currentTarget) closePdfModal(); // คลิก overlay ปิด
  });

  byId('btn-modal-confirm')?.addEventListener('click', () => {
    const rows = getPdfModalRows();
    closePdfModal();
    exportPDF(rows);
  });

  byId('pdf-date-from')?.addEventListener('change', () => {
  syncPdfPresetButtons(null);
  updatePdfModalSummary();
});
  byId('pdf-date-to')?.addEventListener('change', () => {
  syncPdfPresetButtons(null);
  updatePdfModalSummary();
});

  document.querySelectorAll('[data-pdf-preset]').forEach(btn => {
    btn.addEventListener('click', () => applyPdfPreset(btn.dataset.pdfPreset));
  });
  
  byId('btn-scroll-table')?.addEventListener('click', () => {
    byId('section-table')?.scrollIntoView({ behavior: 'smooth' });
  });

  DOM.tableSearch()?.addEventListener('input', recomputeTableView);

  DOM.tableSortSelect()?.addEventListener('change', event => {
    applySortKey(event.target.value);
  });

  DOM.tableSortDirection()?.addEventListener('click', toggleTableSortDirection);

  DOM.rangeChips().forEach(chip => {
    chip.addEventListener('click', () => applyRangePreset(chip.dataset.rangePreset));
  });

  DOM.sortHeaders().forEach(header => {
    header.addEventListener('click', () => applySort(header.dataset.sort));
  });

  DOM.btnPrev()?.addEventListener('click', () => {
    if (state.table.currentPage > 1) {
      state.table.currentPage -= 1;
      renderTable();
    }
  });

  DOM.btnNext()?.addEventListener('click', () => {
    const totalPages = Math.ceil(state.filteredRows.length / CONFIG.rowsPerPage);
    if (state.table.currentPage < totalPages) {
      state.table.currentPage += 1;
      renderTable();
    }
  });

  DOM.pageNumbers()?.addEventListener('click', event => {
    const button = event.target.closest('.page-btn[data-page]');
    if (!button) return;
    state.table.currentPage = Number(button.dataset.page);
    renderTable();
  });

  DOM.chartTabs().forEach(tab => {
    tab.addEventListener('click', () => applyChartRange(tab.dataset.range));
  });

  DOM.btnErrorRetry()?.addEventListener('click', manualRefresh);
}

function resizeAllCharts() {
  Object.values(charts).forEach(chart => {
    if (!chart) return;
    chart.resize();
    chart.update('none');
  });
}

function initApp() {
  tickClock();
  setInterval(tickClock, 1000);

  initCharts();
  bindEvents();

  syncRangeControls();
  syncSortIndicators();
  setText(DOM.rangeSummary(), getRangeSummaryText(0));

  showSkeleton();
  setText(DOM.lastUpdate(), 'Loading...');

  fetchSheet();
  requestAnimationFrame(() => resizeAllCharts());
}

// ── PDF MODAL ──────────────────────────────────────────────

function openPdfModal() {
  const today = new Date().toISOString().slice(0, 10);
  const sevenDaysAgo = new Date(Date.now() - 6 * 864e5).toISOString().slice(0, 10);

  const fromInput = byId('pdf-date-from');
  const toInput   = byId('pdf-date-to');
  if (fromInput) fromInput.value = sevenDaysAgo;
  if (toInput)   toInput.value   = today;

  syncPdfPresetButtons('7d'); // highlight default preset
  updatePdfModalSummary();

  const modal = byId('modal-pdf');
  if (modal) modal.hidden = false;

  document.body.style.overflow = 'hidden'; // ล็อก scroll
}

function closePdfModal() {
  const modal = byId('modal-pdf');
  if (modal) modal.hidden = true;

  document.body.style.overflow = ''; // คืน scroll
}

function getPdfModalRows() {
  const fromVal = byId('pdf-date-from')?.value;
  const toVal   = byId('pdf-date-to')?.value;

  if (!fromVal || !toVal) return state.rows;

  const from = new Date(fromVal);
  const to   = new Date(toVal);
  to.setHours(23, 59, 59, 999);

  return state.rows.filter(row => row.time >= from && row.time <= to);
}

function updatePdfModalSummary() {
  const rows    = getPdfModalRows();
  const summary = byId('pdf-modal-summary');
  if (!summary) return;

  const fromVal = byId('pdf-date-from')?.value;
  const toVal   = byId('pdf-date-to')?.value;

  if (!fromVal || !toVal) {
    summary.textContent = `All records — ${rows.length} entries`;
    return;
  }

  if (rows.length === 0) {
    summary.textContent = 'No records found in this date range';
    return;
  }

  summary.textContent = `${fromVal} → ${toVal} — ${rows.length} entries`;
}

function syncPdfPresetButtons(activePreset) {
  document.querySelectorAll('[data-pdf-preset]').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.pdfPreset === activePreset);
  });
}

function applyPdfPreset(preset) {
  const today = new Date();
  const from  = new Date(today);
  const to    = today.toISOString().slice(0, 10);

  if (preset === 'today') {
    byId('pdf-date-from').value = to;
    byId('pdf-date-to').value   = to;
  } else if (preset === '7d') {
    from.setDate(from.getDate() - 6);
    byId('pdf-date-from').value = from.toISOString().slice(0, 10);
    byId('pdf-date-to').value   = to;
  } else if (preset === '30d') {
    from.setDate(from.getDate() - 29);
    byId('pdf-date-from').value = from.toISOString().slice(0, 10);
    byId('pdf-date-to').value   = to;
  } else {
    byId('pdf-date-from').value = '';
    byId('pdf-date-to').value   = '';
  }

  syncPdfPresetButtons(preset); // highlight ที่เลือก
  updatePdfModalSummary();
}

let resizeTimer = null;

window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    resizeAllCharts();
  }, 120);
});

document.addEventListener('DOMContentLoaded', initApp);

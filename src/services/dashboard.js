import { createDonutLabelPlugin } from '../charts/plugins.js';
import { createManualRefresh, fetchSheetData as requestSheetData } from './fetcher.js';
import { createPdfService } from './pdf.js';
import { createPdfDatePicker } from './pdfDatePicker.js';
import { escapeHtml } from '../utils/escapeHtml.js';
import { createReactiveState } from '../store/reactiveState.js';
import {
  DATE_FORMATTERS,
  addMonths,
  formatDateInputValue,
  formatDateTime,
  formatPdfDisplayDate,
  formatTableClock,
  formatTableDate,
  formatTimeShort,
  isSameDay,
  parseDateInputValue,
  parseTimestamp,
  startOfMonth,
} from '../utils/dateHelper.js';
import { calcAQI, getAqiMeta } from '../utils/aqiHelper.js';
import {
  applyChartRange,
  clearCharts,
  getCharts,
  initCharts,
  resizeAllCharts,
  updateCharts,
} from '../charts/dashboardCharts.js';
import { CONFIG } from '../config.js';

const METRICS_LIMITS = Object.freeze({
  pm25: { gaugeMax: 200 },
  pm10: { gaugeMax: 400 },
  co2: { baseline: 350, gaugeRange: 4650 },
  power: { reference: 3000 },
  energy: { reference: 30 },
});

/** Unicode text tokens (ASCII-safe source; avoids mojibake in editors) */
const TXT = Object.freeze({
  emDash: '\u2014',
  arrow: '\u2192',
  ellipsis: '\u2026',
  microgPerM3: '\u00b5g/m\u00b3',
});

const INITIAL_STATE = {
  current: { pm25: 0, pm10: 0, temp: 0, hum: 0, co2: 0, volt: 0, curr: 0, pwr: 0, energy: 0 },
  previous: { pm25: 0, pm10: 0, temp: 0, hum: 0, co2: 0, volt: 0, curr: 0, pwr: 0, energy: 0 },
  rows: [],
  filteredRows: [],
  table: { sortKey: 'time', sortDir: 1, currentPage: 1 },
  exportRange: { preset: 'all' },
  chartRange: '1h',
  pdfPicker: {
    activeField: 'from',
    pendingAnchor: '',
    suppressClick: false,
    viewDate: null,
    renderFrame: 0,
    drag: { active: false, anchor: '', current: '', moved: false },
  },
  fetch: { isFetching: false, countdown: null, uiState: 'idle', lastErrorMessage: null, initialStartedAt: 0 },
};

const store = createReactiveState(INITIAL_STATE);
const { state, setState, subscribe, batch } = store;

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
  shell: () => queryAllCached('.shell')[0] ?? null,
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
  pdfDateFrom: () => byId('pdf-date-from'),
  pdfDateTo: () => byId('pdf-date-to'),
  pdfDateFromDisplay: () => byId('pdf-date-from-display'),
  pdfDateToDisplay: () => byId('pdf-date-to-display'),
  pdfDateFieldButtons: () => queryAllCached('[data-pdf-field]'),
  pdfCalendarTitle: () => byId('pdf-calendar-title'),
  pdfCalendarSubtitle: () => byId('pdf-calendar-subtitle'),
  pdfCalendarMonths: () => byId('pdf-calendar-months'),
  btnPdfPrevMonth: () => byId('btn-pdf-prev-month'),
  btnPdfNextMonth: () => byId('btn-pdf-next-month'),
  pdfHealthPm25: () => byId('pdf-health-pm25'),
  pdfHealthCo2: () => byId('pdf-health-co2'),
  pdfHealthEnergy: () => byId('pdf-health-energy'),
  tableBody: () => byId('table-body'),
  pageInfo: () => byId('page-info'),
  pageNumbers: () => byId('page-numbers'),
  btnPrev: () => byId('btn-prev'),
  btnNext: () => byId('btn-next'),
  sortHeaders: () => queryAllCached('thead th[data-sort]'),
  rangeChips: () => queryAllCached('[data-range-preset]'),
  pdfPresetButtons: () => queryAllCached('[data-pdf-preset]'),
  chartTabs: () => queryAllCached('.chart-tab'),
  toast: () => byId('toast'),
  toastMessage: () => byId('toast-message'),
  btnRefresh: () => byId('btn-refresh'),
  skeletons: () => ['sk-pm25', 'sk-pm10', 'sk-temp', 'sk-hum', 'sk-volt', 'sk-curr', 'sk-insight']
    .map(id => byId(id))
    .filter(Boolean),
  errorBanner: () => byId('error-banner'),
  errorBannerTitle: () => byId('error-banner-title'),
  errorBannerText: () => byId('error-banner-text'),
  errorBannerMeta: () => byId('error-banner-meta'),
  emptyState: () => byId('empty-state'),
  btnErrorRetry: () => byId('btn-error-retry'),
  navStatus: () => byId('nav-status'),
  navStatusText: () => byId('nav-status-text'),
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

const METRIC_KEYS = Object.freeze([
  'pm25',
  'pm10',
  'temp',
  'hum',
  'co2',
  'volt',
  'curr',
  'pwr',
  'energy',
]);



let lastFocusedElementBeforeModal = null;
const MODAL_FOCUSABLE_SELECTOR = [
  'a[href]',
  'button:not([disabled])',
  'input:not([disabled])',
  'select:not([disabled])',
  'textarea:not([disabled])',
  '[tabindex]:not([tabindex="-1"])',
].join(',');



const fmt = (value, digits = 1) => {
  const number = Number(value);
  return Number.isFinite(number) ? number.toFixed(digits) : '-';
};

// â”€â”€ XSS Protection â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// à¹ƒà¸Šà¹‰à¸à¸±à¸šà¸‚à¹‰à¸­à¸¡à¸¹à¸¥à¸—à¸µà¹ˆà¸¡à¸²à¸ˆà¸²à¸ API à¸à¹ˆà¸­à¸™ render à¸¥à¸‡ innerHTML
// à¸›à¹‰à¸­à¸‡à¸à¸±à¸™ HTML injection à¸ˆà¸²à¸à¸„à¹ˆà¸²à¸—à¸µà¹ˆà¸œà¸´à¸”à¸›à¸à¸•à¸´à¹ƒà¸™ Google Sheet
function setHTML(element, html) {
  if (element) element.innerHTML = html;
}

function setText(element, text) {
  if (element) element.textContent = text;
}

// ── Animated Number Counter ─────────────────────────────────
const activeAnimations = new Map();

function animateCount(element, endValue, options = {}) {
  if (!element) return;
  const { digits = 1, duration = 600, suffix = '' } = options;
  const startText = element.textContent?.replace(/[^\d.-]/g, '') || '0';
  const startValue = parseFloat(startText) || 0;
  const end = Number(endValue);

  if (!Number.isFinite(end)) {
    element.textContent = '-';
    return;
  }

  // Skip animation if values are the same
  if (Math.abs(startValue - end) < Math.pow(10, -digits)) {
    element.textContent = end.toFixed(digits) + suffix;
    return;
  }

  // Cancel previous animation on this element
  const prev = activeAnimations.get(element);
  if (prev) cancelAnimationFrame(prev);

  const startTime = performance.now();
  const easeOutQuart = t => 1 - Math.pow(1 - t, 4);

  function tick(now) {
    const elapsed = now - startTime;
    const progress = Math.min(elapsed / duration, 1);
    const eased = easeOutQuart(progress);
    const current = startValue + (end - startValue) * eased;
    element.textContent = current.toFixed(digits) + suffix;

    if (progress < 1) {
      activeAnimations.set(element, requestAnimationFrame(tick));
    } else {
      activeAnimations.delete(element);
    }
  }

  activeAnimations.set(element, requestAnimationFrame(tick));
}

function setWidth(element, percent) {
  if (!element) return;
  const clamped = Math.min(100, Math.max(0, percent));
  element.style.width = `${clamped}%`;
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

function formatEntryCount(count) {
  return `${count} ${count === 1 ? 'entry' : 'entries'}`;
}

function getFriendlyFetchStatus(error, isFirstLoad) {
  const rawMessage = error?.message || 'Unknown connection issue';
  const isOffline = typeof navigator !== 'undefined' && navigator.onLine === false;
  const isTimeout = /timed out|timeout/i.test(rawMessage);
  const httpMatch = rawMessage.match(/HTTP\s+(\d+)/i);
  const statusCode = httpMatch ? Number(httpMatch[1]) : null;

  if (isOffline) {
    return {
      title: 'You appear to be offline',
      detail: 'Atmosfera cannot reach the network right now. Saved readings will stay visible while it retries.',
      meta: 'Network connection unavailable.',
      toast: 'Offline - retrying automatically',
    };
  }

  if (isTimeout) {
    return {
      title: 'Data source is slow to respond',
      detail: isFirstLoad
        ? 'The sensor feed did not respond in time. Atmosfera will try again shortly.'
        : 'Showing the latest saved readings while the sensor feed catches up.',
      meta: rawMessage,
      toast: 'Data source timed out',
    };
  }

  if (statusCode >= 500) {
    return {
      title: 'Sensor data is temporarily offline',
      detail: isFirstLoad
        ? 'Atmosfera could not reach the data source. The dashboard will retry automatically.'
        : 'Live sync paused. The dashboard is still showing the latest saved readings.',
      meta: `Data service returned ${rawMessage}.`,
      toast: 'Sensor data offline; retrying',
    };
  }

  if (statusCode === 401 || statusCode === 403) {
    return {
      title: 'Data access needs attention',
      detail: 'The data source rejected this request. Check the API token or Google Script permissions.',
      meta: `Authorization issue: ${rawMessage}.`,
      toast: 'Data access issue',
    };
  }

  return {
    title: isFirstLoad ? 'Unable to load sensor data' : 'Live sync paused',
    detail: isFirstLoad
      ? 'Atmosfera could not load readings yet. It will retry automatically.'
      : 'Showing the latest saved readings while Atmosfera reconnects.',
    meta: rawMessage,
    toast: 'Sync issue; retrying',
  };
}

function formatPdfMetric(value, digits = 1, unit = '') {
  if (!Number.isFinite(value)) return TXT.emDash;
  return `${fmt(value, digits)}${unit ? ` ${unit}` : ''}`;
}

function getFriendlyPdfRangeLabel(fromVal, toVal) {
  if (!fromVal || !toVal) return 'All matching records';

  const { preset } = state.exportRange;
  if (preset === 'today' && fromVal === toVal) return 'Today';

  const presetDates = getDateValuesForPreset(preset);
  if (preset === '7d' && fromVal === presetDates.fromValue && toVal === presetDates.toValue) {
    return 'Last 7 days';
  }
  if (preset === '30d' && fromVal === presetDates.fromValue && toVal === presetDates.toValue) {
    return 'Last 30 days';
  }

  if (fromVal === toVal) {
    return formatPdfDisplayDate(fromVal);
  }

  return `${formatPdfDisplayDate(fromVal)} ${TXT.arrow} ${formatPdfDisplayDate(toVal)}`;
}

function normalizeRow(raw) {
  const row = {
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
  row._search = [
    row.time?.toISOString()?.slice(0, 16) ?? '',
    row.pm25, row.pm10, row.temp, row.hum,
    row.co2, row.volt, row.curr, row.pwr, row.energy,
  ].join(' ').toLowerCase();
  return row;
}

function normalizeRows(rawRows) {
  const rows = rawRows.map(normalizeRow);
  const validRows = rows.filter(row => row.time !== null);

  if (validRows.length < rows.length) {
    console.warn(`[Atmosfera] Skipped ${rows.length - validRows.length} row(s) with invalid timestamps`);
  }

  return validRows.reverse();
}

function handleFetchError(error, options = {}) {
  const { trigger = 'auto' } = options;
  hideEmptyState();

  const isFirstLoad = state.rows.length === 0;
  const status = getFriendlyFetchStatus(error, isFirstLoad);

  showErrorBanner(status);
  if (trigger !== 'auto' || state.fetch.lastErrorMessage !== error.message) {
    showToast(status.toast);
  }
  setState({ fetch: { lastErrorMessage: error.message } });
  console.error('[Atmosfera] fetch error:', error);

  startCountdown();
}

function showErrorBanner(status) {
  const banner = DOM.errorBanner();
  const title = DOM.errorBannerTitle();
  const text = DOM.errorBannerText();
  const meta = DOM.errorBannerMeta();
  if (title) title.textContent = status.title;
  if (text) text.textContent = status.detail;
  if (meta) meta.textContent = status.meta;
  if (banner) banner.hidden = false;
}

function hideErrorBanner() {
  const banner = DOM.errorBanner();
  if (banner) banner.hidden = true;
}

function hideEmptyState() {
  const empty = DOM.emptyState();
  if (empty) empty.hidden = true;
}

function setSystemStatus(status) {
  const navStatus = DOM.navStatus();
  const navStatusText = DOM.navStatusText();
  if (navStatus) {
    navStatus.classList.toggle('is-offline', status === 'offline');
    navStatus.classList.toggle('is-waiting', status === 'waiting');
  }

  if (!navStatusText) return;
  if (status === 'offline') {
    navStatusText.textContent = 'System Status: Offline';
  } else if (status === 'waiting') {
    navStatusText.textContent = 'System Status: Waiting';
  } else {
    navStatusText.textContent = 'System Status: Live';
  }
}

function resetDeltaBadge(element) {
  if (!element) return;
  element.className = 'delta delta-flat delta--reset';
  element.textContent = TXT.emDash;
}

function resetMetricState() {
  const emptyMetrics = { pm25: 0, pm10: 0, temp: 0, hum: 0, co2: 0, volt: 0, curr: 0, pwr: 0, energy: 0 };
  setState({
    current: { ...emptyMetrics },
    previous: { ...emptyMetrics },
  });
}

function resetDashboardBadges() {
  [
    DOM.deltaPm25(),
    DOM.deltaPm10(),
    DOM.deltaTemp(),
    DOM.deltaHum(),
    DOM.deltaVolt(),
    DOM.deltaCurr(),
    DOM.deltaCo2(),
  ].forEach(resetDeltaBadge);
}

function resetInsightPanel() {
  setHTML(DOM.insightPwr(), `${TXT.emDash} <span style="font-size:.85rem;font-weight:400;opacity:.5">W</span>`);
  setHTML(DOM.insightEnergy(), `${TXT.emDash} <span style="font-size:.85rem;font-weight:400;opacity:.5">kWh</span>`);
  setHTML(DOM.insightPf(), `${TXT.emDash} <span style="font-size:.85rem;font-weight:400;opacity:.5">PF</span>`);
  setText(DOM.insightEff(), TXT.emDash);

  const insightEff = DOM.insightEff();
  if (insightEff) insightEff.style.color = '';

  setWidth(DOM.barPwr(), 0);
  setWidth(DOM.barEnergy(), 0);
  setWidth(DOM.barPf(), 0);
}

function showEmptyState() {
  const empty = DOM.emptyState();
  if (empty) empty.hidden = false;
}

function renderEmptyState() {
  batch(() => {
    setState({
      fetch: { uiState: 'empty' },
      rows: [],
      filteredRows: [],
      table: { currentPage: 1 },
    });
    resetMetricState();
  });

  setSystemStatus('waiting');
  hideErrorBanner();
  showEmptyState();

  setText(DOM.lastUpdate(), 'No records yet');

  setText(DOM.aqiScore(), TXT.emDash);
  setText(DOM.aqiStatusText(), 'No sensor data yet');

  const aqiDot = DOM.aqiDot();
  if (aqiDot) {
    aqiDot.style.background = 'rgba(90,106,114,.28)';
    aqiDot.style.boxShadow = 'none';
  }

  setText(DOM.bannerPm25(), TXT.emDash);
  setText(DOM.bannerPm10(), TXT.emDash);
  setText(DOM.bannerCo2(), TXT.emDash);

  setWidth(DOM.gaugePm25(), 0);
  setWidth(DOM.gaugePm10(), 0);
  setWidth(DOM.gaugeCo2(), 0);

  setText(DOM.miniTemp(), TXT.emDash);
  setText(DOM.miniHum(), TXT.emDash);
  setText(DOM.miniPwr(), TXT.emDash);

  setHTML(DOM.valPm25(), `${TXT.emDash}<span class="metric-unit">&micro;g/m&sup3;</span>`);
  setHTML(DOM.valPm10(), `${TXT.emDash}<span class="metric-unit">&micro;g/m&sup3;</span>`);
  setHTML(DOM.valTemp(), `${TXT.emDash}<span class="metric-unit">C</span>`);
  setHTML(DOM.valHum(), `${TXT.emDash}<span class="metric-unit">%</span>`);
  setHTML(DOM.valCo2(), `${TXT.emDash}<span class="metric-unit">ppm</span>`);
  setHTML(DOM.valVolt(), `${TXT.emDash}<span class="metric-unit">V</span>`);
  setHTML(DOM.valCurr(), `${TXT.emDash}<span class="metric-unit">A</span>`);
  resetDashboardBadges();
  resetInsightPanel();

  if (DOM.rangeSummary()) {
    setText(DOM.rangeSummary(), 'Date range: No records yet');
  }

  renderTable();
  clearCharts();
}

async function fetchSheet(options = {}) {
  const { showLoading = false, trigger = 'auto' } = options;

  if (state.fetch.isFetching) return;

  if (trigger === 'initial') {
    setState({ fetch: { initialStartedAt: performance.now() } });
  }

  if (showLoading) {
    setRefreshVisualState(true);
  }

  setState({ fetch: { isFetching: true } });
  setText(DOM.lastUpdate(), 'Loading...');

  try {
    const rawRows = await requestSheetData(CONFIG);

    if (rawRows.length === 0) {
      setState({
        fetch: { lastErrorMessage: null, uiState: 'empty' },
      });
      renderEmptyState();
      startCountdown();
      return;
    }

    const normalizedRows = normalizeRows(rawRows.slice(-CONFIG.maxRows));
    const [currentRow] = normalizedRows;
    const previousRow = normalizedRows[1] ?? currentRow;

    const nextCurrent = {};
    const nextPrevious = {};
    METRIC_KEYS.forEach(key => {
      nextCurrent[key] = currentRow?.[key] ?? 0;
      nextPrevious[key] = previousRow?.[key] ?? 0;
    });

    batch(() => {
      setState({
        fetch: { lastErrorMessage: null, uiState: 'ready' },
        rows: normalizedRows,
        current: nextCurrent,
        previous: nextPrevious,
      });
    });

    hideEmptyState();
    hideErrorBanner();

    setText(DOM.lastUpdate(), 'Just now');
    showToast(`Synced ${normalizedRows.length} records`);

    startCountdown();
    requestAnimationFrame(() => resizeAllCharts());
    requestAnimationFrame(() => syncActiveNavLink());
  } catch (error) {
    setState({ fetch: { uiState: 'error' } });
    hideEmptyState();
    handleFetchError(error, { trigger });
  } finally {
    setState({ fetch: { isFetching: false } });
    setRefreshVisualState(false);
    if (trigger === 'initial') {
      const elapsed = performance.now() - state.fetch.initialStartedAt;
      const remaining = Math.max(0, 650 - elapsed);
      window.setTimeout(() => {
        DOM.shell()?.classList.remove('is-starting');
        syncActiveNavLink();
      }, remaining);
    }
  }
}

function renderMetricCards() {
  const current = state.current;

  // Animate main number, then set unit via sibling
  const animateMetric = (el, value, digits, unitHtml) => {
    if (!el) return;
    let numSpan = el.querySelector('.metric-num');
    let unitSpan = el.querySelector('.metric-unit');
    if (!numSpan) {
      el.innerHTML = `<span class="metric-num"></span><span class="metric-unit">${unitHtml}</span>`;
      numSpan = el.querySelector('.metric-num');
    }
    animateCount(numSpan, value, { digits });
  };

  animateMetric(DOM.valPm25(), current.pm25, 1, '&micro;g/m&sup3;');
  animateMetric(DOM.valPm10(), current.pm10, 1, '&micro;g/m&sup3;');
  animateMetric(DOM.valTemp(), current.temp, 1, 'C');
  animateMetric(DOM.valHum(), current.hum, 0, '%');
  animateMetric(DOM.valCo2(), current.co2, 0, 'ppm');
  animateMetric(DOM.valVolt(), current.volt, 1, 'V');
  animateMetric(DOM.valCurr(), current.curr, 1, 'A');
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
    co2Delta.className = 'delta delta--reset';
    co2Delta.textContent = deltaLabel(current.co2 - previous.co2, '');
  }
}

function renderAqiBanner() {
  const current = state.current;
  const aqi = calcAQI(current.pm25, current.pm10);
  const meta = getAqiMeta(aqi);

  setText(DOM.aqiStatusText(), meta.label);

  animateCount(DOM.aqiScore(), aqi, { digits: 0, duration: 800 });

  const dot = DOM.aqiDot();
  if (dot) dot.style.background = meta.dot;

  setText(DOM.bannerPm25(), `${fmt(current.pm25)} \u00b5g/m\u00b3`);
  setText(DOM.bannerPm10(), `${fmt(current.pm10)} \u00b5g/m\u00b3`);
  setText(DOM.bannerCo2(), `${fmt(current.co2, 0)} ppm`);
  setHTML(DOM.miniTemp(), `${fmt(current.temp)}<span class="aqi-mini-unit">C</span>`);
  setHTML(DOM.miniHum(), `${fmt(current.hum, 0)}<span class="aqi-mini-unit">%</span>`);
  setHTML(DOM.miniPwr(), `${fmt(current.pwr, 1)}<span class="aqi-mini-unit">W</span>`);

  setWidth(DOM.gaugePm25(), (current.pm25 / METRICS_LIMITS.pm25.gaugeMax) * 100);
  setWidth(DOM.gaugePm10(), (current.pm10 / METRICS_LIMITS.pm10.gaugeMax) * 100);
  setWidth(DOM.gaugeCo2(), ((current.co2 - METRICS_LIMITS.co2.baseline) / METRICS_LIMITS.co2.gaugeRange) * 100);
}

function renderInsightPanel() {
  const current = state.current;
  const powerReference = METRICS_LIMITS.power.reference;
  const energyReference = METRICS_LIMITS.energy.reference;
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

  const chartsRef = getCharts();
  if (chartsRef.donut) {
    const percent = Math.max(0, Math.min(100, ((current.co2 - METRICS_LIMITS.co2.baseline) / METRICS_LIMITS.co2.gaugeRange) * 100));
    chartsRef.donut.data.datasets[0].data = [percent, 100 - percent];
    chartsRef.donut.update('none');
  }
}

function renderDashboard() {
  renderMetricCards();
  renderDeltaBadges();
  renderAqiBanner();
  renderInsightPanel();
}

let tableRenderFrame = 0;

function scheduleRenderTable() {
  if (tableRenderFrame) return;
  tableRenderFrame = requestAnimationFrame(() => {
    tableRenderFrame = 0;
    renderTable();
  });
}

// ── Reactive Subscriptions ──────────────────────────────────
subscribe('current', renderMetricCards);
subscribe('previous', renderDeltaBadges);
subscribe('current', renderAqiBanner);
subscribe('current', renderInsightPanel);
subscribe('rows', () => {
  recomputeTableView();
  updateCharts(state, DOM, formatTimeShort);
});
subscribe('filteredRows', scheduleRenderTable);
subscribe('table', scheduleRenderTable);
subscribe('fetch.uiState', (s) => {
  if (s.fetch.uiState === 'ready') setSystemStatus('live');
  else if (s.fetch.uiState === 'error') setSystemStatus('offline');
  else if (s.fetch.uiState === 'empty') setSystemStatus('waiting');
});

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

function getDateValuesForPreset(preset) {
  const today = new Date();
  const toValue = formatDateInputValue(today);

  if (preset === 'today') {
    return { fromValue: toValue, toValue };
  }

  if (preset === '7d') {
    const from = new Date(today);
    from.setDate(from.getDate() - 6);
    return { fromValue: formatDateInputValue(from), toValue };
  }

  if (preset === '30d') {
    const from = new Date(today);
    from.setDate(from.getDate() - 29);
    return { fromValue: formatDateInputValue(from), toValue };
  }

  return { fromValue: '', toValue: '' };
}

function matchesSelectedRange(row) {
  const bounds = getSelectedRangeBounds();
  if (!bounds) return true;
  return row.time >= bounds.from && row.time <= bounds.to;
}

function matchesSearchQuery(row, query = DOM.tableSearch()?.value.toLowerCase().trim() || '') {
  if (!query) return true;
  return row._search.includes(query);
}

function filterRows(rows) {
  const query = DOM.tableSearch()?.value.toLowerCase().trim() || '';

  return rows.filter(row => {
    if (!matchesSelectedRange(row)) return false;

    return matchesSearchQuery(row, query);
  });
}

function getFilteredAndSortedRows(rows = state.rows) {
  return sortRows(filterRows(rows));
}

function getExportRows() {
  return getFilteredAndSortedRows(state.rows);
}

function getRowsMatchingCurrentSearch(rows = state.rows) {
  const query = DOM.tableSearch()?.value.toLowerCase().trim() || '';
  return rows.filter(row => matchesSearchQuery(row, query));
}

function sortRowsByTime(rows, direction = 'asc') {
  const multiplier = direction === 'desc' ? -1 : 1;
  return [...rows].sort((left, right) => (left.time - right.time) * multiplier);
}

function getLatestRowByTime(rows) {
  if (!rows.length) return null;
  return rows.reduce((latestRow, row) => (row.time > latestRow.time ? row : latestRow), rows[0]);
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
  batch(() => {
    setState({ table: { currentPage: 1 } });

    if (state.rows.length === 0) {
      setState({ filteredRows: [] });
      syncRangeControls();
      syncSortIndicators();
      setText(DOM.rangeSummary(), 'Date range: No records yet');
      return;
    }

    setState({ filteredRows: getFilteredAndSortedRows(state.rows) });

    syncRangeControls();
    syncSortIndicators();

    setText(DOM.rangeSummary(), getRangeSummaryText(state.filteredRows.length));
  });
}

function applySort(key) {
  if (!TABLE_SORT_LABELS[key]) return;

  if (state.table.sortKey === key) {
    setState({ table: { sortDir: state.table.sortDir * -1 } });
  } else {
    setState({ table: { sortKey: key, sortDir: 1 } });
  }

  recomputeTableView();
}

function applySortKey(key) {
  if (!TABLE_SORT_LABELS[key]) return;

  if (state.table.sortKey === key) {
    syncMobileSortControls();
    return;
  }

  setState({ table: { sortKey: key, sortDir: 1 } });
  recomputeTableView();
}

function toggleTableSortDirection() {
  setState({ table: { sortDir: state.table.sortDir * -1 } });
  recomputeTableView();
}

function applyRangePreset(preset) {
  if (!['all', 'today', '7d', '30d'].includes(preset)) return;
  setState({ exportRange: { preset } });
  recomputeTableView();
}

function getPaginationItems(totalPages, currentPage) {
  if (totalPages <= 7) {
    return Array.from({ length: totalPages }, (_, index) => index + 1);
  }

  if (currentPage <= 4) {
    return [1, 2, 3, 4, 5, 'ellipsis', totalPages];
  }

  if (currentPage >= totalPages - 3) {
    return [1, 'ellipsis', totalPages - 4, totalPages - 3, totalPages - 2, totalPages - 1, totalPages];
  }

  return [1, 'ellipsis', currentPage - 1, currentPage, currentPage + 1, 'ellipsis', totalPages];
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

  // Fade out → update → fade in
  if (state.rows.length === 0) {
    tbody.innerHTML = `
      <tr class="table-empty-state">
        <td colspan="10" class="no-results">No sensor data available yet</td>
      </tr>
    `;
  } else if (pageRows.length === 0) {
    tbody.innerHTML = `
      <tr class="table-empty-state">
        <td colspan="10" class="no-results">No matching records found</td>
      </tr>
    `;
  } else {
    tbody.innerHTML = pageRows.map(row => `
      <tr>
        <td class="td-time" data-label="Recorded">
          <span class="td-date">${escapeHtml(formatTableDate(row.time))}</span>
          <span class="td-clock">${escapeHtml(formatTableClock(row.time))}</span>
        </td>
        <td class="td-value" data-label="PM2.5">${escapeHtml(fmt(row.pm25))}</td>
        <td class="td-value" data-label="PM10">${escapeHtml(fmt(row.pm10))}</td>
        <td class="td-value" data-label="Temperature">${escapeHtml(fmt(row.temp))}</td>
        <td class="td-value" data-label="Humidity">${escapeHtml(String(row.hum))}</td>
        <td class="td-value" data-label="CO2">${escapeHtml(String(row.co2))}</td>
        <td class="td-value" data-label="Voltage">${escapeHtml(fmt(row.volt))}</td>
        <td class="td-value" data-label="Current">${escapeHtml(fmt(row.curr, 2))}</td>
        <td class="td-value" data-label="Power">${escapeHtml(fmt(row.pwr, 1))}</td>
        <td class="td-value" data-label="Energy">${escapeHtml(fmt(row.energy, 2))}</td>
      </tr>
    `).join('');
  }

  const rangeStart = total === 0 ? 0 : start + 1;
  const rangeEnd = Math.min(end, total);

  setText(DOM.pageInfo(), `Showing ${rangeStart}-${rangeEnd} of ${total} records`);

  const prevButton = DOM.btnPrev();
  const nextButton = DOM.btnNext();

  if (prevButton) prevButton.disabled = currentPage <= 1 || total === 0;
  if (nextButton) nextButton.disabled = currentPage >= totalPages || total === 0;

  const pageNumbers = DOM.pageNumbers();
  if (!pageNumbers) return;

  if (total === 0) {
    pageNumbers.replaceChildren();
    return;
  }

  const fragment = document.createDocumentFragment();
  getPaginationItems(totalPages, currentPage).forEach(item => {
    if (item === 'ellipsis') {
      const ellipsis = document.createElement('span');
      ellipsis.className = 'page-ellipsis';
      ellipsis.textContent = TXT.ellipsis;
      ellipsis.setAttribute('aria-hidden', 'true');
      fragment.appendChild(ellipsis);
      return;
    }

    const button = document.createElement('button');
    button.className = `page-btn${item === currentPage ? ' active' : ''}`;
    button.textContent = item;
    button.dataset.page = item;
    fragment.appendChild(button);
  });

  pageNumbers.replaceChildren(fragment);

}


const FETCH_STATE_CONFIG = {
  ready: {
    interval: 3600,
    label: (s) => {
      const h = Math.floor(s / 3600);
      const m = Math.floor((s % 3600) / 60);
      const sec = s % 60;
      return h > 0
        ? `Refresh in ${h}h ${m}m`
        : `Refresh in ${m}:${String(sec).padStart(2, '0')}`;
    },
  },
  empty: { interval: 30, label: (s) => `Checking again in ${s}s` },
  error: { interval: 30, label: (s) => `Retrying in ${s}s` },
};

let hiddenAt = null;

function startCountdown(initialSeconds) {
  clearInterval(state.fetch.countdown);

  const cfg = FETCH_STATE_CONFIG[state.fetch.uiState] ?? FETCH_STATE_CONFIG.error;
  let remainingSeconds = initialSeconds ?? cfg.interval;

  setText(DOM.lastUpdate(), cfg.label(remainingSeconds));

  state.fetch.countdown = setInterval(() => {
    remainingSeconds -= 1;

    if (remainingSeconds <= 0) {
      clearInterval(state.fetch.countdown);
      fetchSheet({ showLoading: true, trigger: 'auto' });
      return;
    }

    setText(DOM.lastUpdate(), cfg.label(remainingSeconds));
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
    element.style.opacity = '';
    element.parentElement?.classList.add('is-loading');
  });
}

function hideSkeleton() {
  DOM.skeletons().forEach(element => {
    const pendingTimer = skeletonHideTimers.get(element);
    if (pendingTimer) clearTimeout(pendingTimer);

    element.style.transition = 'opacity .35s cubic-bezier(.22,.86,.32,1)';
    element.style.opacity = '0';

    const timer = setTimeout(() => {
      element.parentElement?.classList.remove('is-loading');
      element.style.opacity = '';
      element.style.transition = '';
      skeletonHideTimers.delete(element);
    }, 350);

    skeletonHideTimers.set(element, timer);
  });
}

function setRefreshButtonBusy(isBusy) {
  const button = DOM.btnRefresh();
  const icon = button?.querySelector('svg');

  if (button) button.disabled = isBusy;
  if (icon) icon.style.animation = isBusy ? 'spin 0.8s linear infinite' : '';
}

function setRefreshVisualState(isRefreshing) {
  DOM.shell()?.classList.toggle('is-refreshing', isRefreshing);
  setRefreshButtonBusy(isRefreshing);

  if (isRefreshing) {
    hideEmptyState();
    showSkeleton();
    return;
  }

  hideSkeleton();
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

const {
  exportAllCSV,
  exportPDF,
  setPdfExportBusy,
  waitForPaint,
} = createPdfService({
  byId,
  fmt,
  formatDateInputValue,
  formatDateTime,
  getExportRows,
  getLatestRowByTime,
  getSelectedRangeLabel,
  showToast,
  sortRowsByTime,
});

const pdfPicker = createPdfDatePicker({
  state,
  DOM,
  setText,
  byId,
  getRowsMatchingCurrentSearch,
  sortRowsByTime,
  getFriendlyPdfRangeLabel,
  formatEntryCount,
  formatPdfMetric,
  TXT,
});

function tickClock() {
  const now = new Date();
  setText(DOM.heroClock(), DATE_FORMATTERS.heroClock.format(now));
  setText(DOM.heroDate(), DATE_FORMATTERS.heroDate.format(now));
}

const manualRefresh = createManualRefresh({ state, fetchSheet });

let isScrollingFromNav = false;
let activeNavHash = '';

function setActiveNavLink(hash, options = {}) {
  const normalizedHash = hash || '#main-content';
  const force = options.force === true;
  if (!force && activeNavHash === normalizedHash) return;

  let activeLink = null;

  document.querySelectorAll('.nav-link').forEach(link => {
    const shouldBeActive = link.getAttribute('href') === normalizedHash;
    if (link.classList.contains('active') !== shouldBeActive) {
      link.classList.toggle('active', shouldBeActive);
    }
    if (shouldBeActive) activeLink = link;
  });

  if (activeLink && window.matchMedia('(max-width: 960px)').matches) {
    activeLink.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
  }

  // Slide the nav indicator to the active link
  moveNavIndicator(activeLink);
  activeNavHash = normalizedHash;
}

let navIndicator = null;

function moveNavIndicator(activeLink) {
  if (!activeLink) return;
  const menu = activeLink.closest('.nav-menu');
  if (!menu) return;

  if (!navIndicator) {
    navIndicator = document.createElement('div');
    navIndicator.className = 'nav-indicator';
    menu.style.position = 'relative';
    menu.appendChild(navIndicator);
  }

  // อัปเดตทันทีแบบ synchronous เพื่อความตอบสนองที่ไร้ดีเลย์
  const menuRect = menu.getBoundingClientRect();
  const linkRect = activeLink.getBoundingClientRect();

  navIndicator.style.transform = `translateY(${linkRect.top - menuRect.top}px)`;
  navIndicator.style.height = `${linkRect.height}px`;
}

let navSelectionLockHash = null;
let navSelectionUnlockTimer = null;

function unlockNavSelection() {
  navSelectionLockHash = null;
  navSelectionUnlockTimer = null;
  syncActiveNavLink();
}

function scheduleNavSelectionUnlock(delay = 180) {
  clearTimeout(navSelectionUnlockTimer);
  navSelectionUnlockTimer = setTimeout(unlockNavSelection, delay);
}

function getNavSections() {
  return [...document.querySelectorAll('.nav-link[href^="#"]')]
    .map(link => {
      const hash = link.getAttribute('href');
      const target = document.querySelector(hash);
      return target ? { hash, target } : null;
    })
    .filter(Boolean);
}

let navSectionMetrics = [];

function measureNavSectionMetrics() {
  const scrollY = window.scrollY;
  navSectionMetrics = getNavSections().map(section => ({
    hash: section.hash,
    top: section.target.getBoundingClientRect().top + scrollY,
  }));
}

function getScrollSpyProbe() {
  const isDesktopSidebar = window.matchMedia('(min-width: 961px)').matches;
  return window.scrollY + (isDesktopSidebar ? 48 : 128);
}

function pickActiveSectionFromScroll() {
  if (!navSectionMetrics.length) {
    measureNavSectionMetrics();
  }
  if (!navSectionMetrics.length) return;

  const scrollY = window.scrollY;
  const viewportBottom = scrollY + window.innerHeight;
  const documentHeight = document.documentElement.scrollHeight;

  if (viewportBottom >= documentHeight - 12) {
    setActiveNavLink(navSectionMetrics[navSectionMetrics.length - 1].hash);
    return;
  }

  const probe = getScrollSpyProbe();
  let activeHash = navSectionMetrics[0].hash;

  for (const section of navSectionMetrics) {
    if (probe >= section.top) {
      activeHash = section.hash;
    }
  }

  setActiveNavLink(activeHash);
}

function syncActiveNavLink() {
  if (navSelectionLockHash) {
    setActiveNavLink(navSelectionLockHash, { force: true });
    return;
  }

  measureNavSectionMetrics();
  pickActiveSectionFromScroll();
}

// ── Custom Smooth Scroll Engine (Lerp-based) ────────────────
const smoothScrollEngine = (() => {
  const SCROLL_PADDING = 80;

  let targetY = window.scrollY;
  let rafId = null;

  function getMaxScroll() {
    return Math.max(0, document.documentElement.scrollHeight - window.innerHeight);
  }

  function startAnimation() {
    targetY = Math.max(0, Math.min(targetY, getMaxScroll()));
    const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    window.scrollTo({ top: targetY, behavior: reduceMotion ? 'auto' : 'smooth' });
    rafId = requestAnimationFrame(() => { rafId = null; });
  }

  function init() {
    targetY = window.scrollY;
  }

  function scrollTo(element) {
    if (!element) return;
    const rect = element.getBoundingClientRect();
    targetY = window.scrollY + rect.top - SCROLL_PADDING;
    startAnimation();
  }

  function setActiveLink(link) {
    void link;
  }

  function getTargetY() {
    return targetY;
  }

  function isAnimating() {
    return rafId !== null;
  }

  function destroy() {
    if (rafId) cancelAnimationFrame(rafId);
    rafId = null;
  }

  return { init, scrollTo, setActiveLink, destroy, getTargetY, isAnimating };
})();

function navigateToSidebarSection(hash) {
  const target = document.querySelector(hash);
  if (!target) return;

  navSelectionLockHash = hash;
  setActiveNavLink(hash);
  smoothScrollEngine.scrollTo(target);
  history.replaceState(null, '', hash);
  scheduleNavSelectionUnlock(900);
}

function bindScrollPerf() {
  const root = document.documentElement;
  let scrollEndTimer = null;

  window.addEventListener('scroll', () => {
    if (!root.classList.contains('is-scrolling')) {
      root.classList.add('is-scrolling');
    }
    clearTimeout(scrollEndTimer);
    scrollEndTimer = setTimeout(() => {
      root.classList.remove('is-scrolling');
    }, 140);
  }, { passive: true });
}

function bindSidebarNavigation() {
  document.querySelectorAll('.nav-link[href^="#"]').forEach(link => {
    link.addEventListener('click', event => {
      const hash = link.getAttribute('href');
      const target = document.querySelector(hash);
      if (!target) return;

      event.preventDefault();

      isScrollingFromNav = true;
      setActiveNavLink(hash);
      smoothScrollEngine.setActiveLink(link);
      smoothScrollEngine.scrollTo(target);
      history.replaceState(null, '', hash);

      setTimeout(() => { isScrollingFromNav = false; }, 850);
    });
  });

  let scrollSpyFrame = null;
  const scheduleScrollSpyUpdate = () => {
    if (isScrollingFromNav || navSelectionLockHash) return;

    if (scrollSpyFrame) return;
    scrollSpyFrame = requestAnimationFrame(() => {
      pickActiveSectionFromScroll();
      scrollSpyFrame = null;
    });
  };

  window.addEventListener('scroll', scheduleScrollSpyUpdate, { passive: true });
  window.addEventListener('resize', () => {
    measureNavSectionMetrics();
    setActiveNavLink(activeNavHash, { force: true });
  });
  window.addEventListener('load', syncActiveNavLink);
  syncActiveNavLink();
}

function bindEvents() {
  byId('btn-scroll-top')?.addEventListener('click', () => {
    navigateToSidebarSection('#main-content');
  });
  bindScrollPerf();
  bindSidebarNavigation();
  DOM.btnRefresh()?.addEventListener('click', manualRefresh);
  byId('btn-export-pdf')?.addEventListener('click', openPdfModal);
  byId('btn-export')?.addEventListener('click', exportAllCSV);

  // â”€â”€ Mobile Export Dropdown â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  const exportToggle = byId('btn-export-toggle');
  const exportMenu   = byId('nav-export-menu');
  const closeExportMenu = () => {
    exportMenu?.setAttribute('hidden', '');
    exportToggle?.setAttribute('aria-expanded', 'false');
    if (document.activeElement === exportToggle) {
      exportToggle.blur();
    }
  };

  exportToggle?.addEventListener('click', e => {
    e.stopPropagation();
    const isHidden = exportMenu?.hasAttribute('hidden');
    if (isHidden) {
      exportMenu?.removeAttribute('hidden');
      exportToggle?.setAttribute('aria-expanded', 'true');
      return;
    }
    closeExportMenu();
  });
  byId('btn-export-pdf-mobile')?.addEventListener('click', () => {
    closeExportMenu();
    openPdfModal();
  });
  byId('btn-export-csv-mobile')?.addEventListener('click', () => {
    closeExportMenu();
    exportAllCSV();
  });
  document.addEventListener('click', e => {
    if (!byId('nav-export-dropdown')?.contains(e.target)) {
      closeExportMenu();
    }
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeExportMenu();
  });

  // modal
  byId('btn-modal-close')?.addEventListener('click', closePdfModal);
  byId('btn-modal-cancel')?.addEventListener('click', closePdfModal);
  byId('modal-pdf')?.addEventListener('click', e => {
    if (e.target === e.currentTarget) closePdfModal(); // à¸„à¸¥à¸´à¸ overlay à¸›à¸´à¸”
  });

  byId('btn-modal-confirm')?.addEventListener('click', async () => {
    if (byId('btn-modal-confirm')?.disabled) return;

    const selection = pdfPicker.getPdfModalSelection();
    setPdfExportBusy(true);

    try {
      await waitForPaint();
      const didExport = exportPDF(selection.rows, {
        rangeLabel: selection.rangeLabel,
        fromValue: selection.fromValue,
        toValue: selection.toValue,
      });

      if (didExport) closePdfModal();
    } finally {
      setPdfExportBusy(false);
    }
  });

  pdfPicker.bindCalendarEvents();

  let searchTimer = null;
  DOM.tableSearch()?.addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(recomputeTableView, 120);
  });

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
      setState({ table: { currentPage: state.table.currentPage - 1 } });
    }
  });

  DOM.btnNext()?.addEventListener('click', () => {
    const totalPages = Math.ceil(state.filteredRows.length / CONFIG.rowsPerPage);
    if (state.table.currentPage < totalPages) {
      setState({ table: { currentPage: state.table.currentPage + 1 } });
    }
  });

  DOM.pageNumbers()?.addEventListener('click', event => {
    const button = event.target.closest('.page-btn[data-page]');
    if (!button) return;
    setState({ table: { currentPage: Number(button.dataset.page) } });
  });

  DOM.chartTabs().forEach(tab => {
    tab.addEventListener('click', () => applyChartRange(tab.dataset.range, state, DOM, formatTimeShort, setState));
  });

  DOM.btnErrorRetry()?.addEventListener('click', manualRefresh);
}

export function initApp() {
  if ('scrollRestoration' in history) {
    history.scrollRestoration = 'manual';
  }

  if (window.location.hash) {
    history.replaceState(null, '', window.location.pathname + window.location.search);
  }

  window.scrollTo({ top: 0, left: 0, behavior: 'auto' });

  tickClock();
  setInterval(tickClock, 1000);

  initCharts(byId, state);
  bindEvents();
  smoothScrollEngine.init();

  syncRangeControls();
  syncSortIndicators();
  setText(DOM.rangeSummary(), getRangeSummaryText(0));

  setText(DOM.lastUpdate(), 'Loading...');

  fetchSheet({ showLoading: true, trigger: 'initial' });
  requestAnimationFrame(() => resizeAllCharts());
  requestAnimationFrame(() => {
    document.documentElement.classList.remove('is-booting');
    setActiveNavLink('#main-content', { force: true });
  });

  // ── Scroll Reveal Animations ──────────────────────────────
  initScrollReveal();

  // ── Hero Parallax ─────────────────────────────────────────
  initHeroParallax();

  // ── Page Visibility API ───────────────────────────────────
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
      hiddenAt = Date.now();
      clearInterval(state.fetch.countdown);
    } else if (hiddenAt !== null) {
      const elapsed = Math.floor((Date.now() - hiddenAt) / 1000);
      hiddenAt = null;
      const cfg = FETCH_STATE_CONFIG[state.fetch.uiState] ?? FETCH_STATE_CONFIG.error;
      if (elapsed >= cfg.interval) {
        fetchSheet({ showLoading: true, trigger: 'auto' });
      } else {
        startCountdown(cfg.interval - elapsed);
      }
    }
  });
}

function initScrollReveal() {
  const revealTargets = document.querySelectorAll(
    '.metric-card, .energy-card, .insight-panel, .chart-card, .table-card, .aqi-banner, .section-header'
  );

  const revealSet = new Set(revealTargets);

  // Set initial hidden state
  revealTargets.forEach(el => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(16px)';
    el.style.transition = 'opacity .45s cubic-bezier(.22,.86,.32,1), transform .45s cubic-bezier(.22,.86,.32,1)';
  });

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;

        const el = entry.target;
        const parent = el.parentElement;
        const siblings = parent ? [...parent.children].filter(c => revealSet.has(c)) : [];
        const idx = siblings.indexOf(el);
        const delay = idx >= 0 ? idx * 40 : 0;

        setTimeout(() => {
          el.style.opacity = '1';
          el.style.transform = 'translateY(0)';

          el.addEventListener('transitionend', () => {
            el.style.removeProperty('opacity');
            el.style.removeProperty('transform');
            el.style.removeProperty('transition');
          }, { once: true });
        }, delay);

        observer.unobserve(el);
      });
    },
    { threshold: 0.12, rootMargin: '0px 0px -20px 0px' }
  );

  revealTargets.forEach(el => observer.observe(el));
}

function initHeroParallax() {
  const hero = document.querySelector('.hero');
  if (!hero) return;

  // Promote to compositor layer
  hero.style.willChange = 'opacity';

  let ticking = false;
  window.addEventListener('scroll', () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => {
      // Only adjust opacity — transforms are expensive with backdrop-filter siblings
      const scrollY = window.scrollY;
      hero.style.opacity = String(Math.max(0, 1 - scrollY / 700));
      ticking = false;
    });
  }, { passive: true });
}

function openPdfModal() {
  const preset = ['all', 'today', '7d', '30d'].includes(state.exportRange.preset)
    ? state.exportRange.preset
    : 'all';
  const { fromValue, toValue } = getDateValuesForPreset(preset);

  pdfPicker.resetPdfDragState();
  setPdfExportBusy(false);
  pdfPicker.setPdfModalDates(
    fromValue,
    toValue,
    {
      anchorValue: fromValue || toValue || formatDateInputValue(new Date()),
      pendingAnchor: '',
    }
  );
  pdfPicker.syncPdfPresetButtons(preset);

  const modal = byId('modal-pdf');
  if (modal) {
    modal.hidden = false;
    lastFocusedElementBeforeModal = document.activeElement;
    modal.addEventListener('keydown', handlePdfModalKeydown);

    // Trigger reflow then animate in
    void modal.offsetHeight;
    modal.classList.add('is-active');

    const firstFocusable = modal.querySelector(MODAL_FOCUSABLE_SELECTOR);
    if (firstFocusable instanceof HTMLElement) firstFocusable.focus();
  }

  document.body.style.overflow = 'hidden'; // à¸¥à¹‡à¸­à¸ scroll
}

function closePdfModal() {
  const modal = byId('modal-pdf');
  if (modal) {
    modal.classList.remove('is-active');
    modal.removeEventListener('keydown', handlePdfModalKeydown);

    // Wait for animation to finish before hiding
    const onEnd = () => {
      modal.hidden = true;
      modal.removeEventListener('transitionend', onEnd);
    };
    modal.addEventListener('transitionend', onEnd, { once: true });
    // Fallback if transitionend doesn't fire
    setTimeout(() => { if (!modal.hidden) modal.hidden = true; }, 400);
  }

  pdfPicker.cancelScheduledPdfCalendarRender();
  pdfPicker.resetPdfDragState();
  pdfPicker.resetPdfPendingAnchor();
  document.body.style.overflow = ''; // à¸„à¸·à¸™ scroll

  if (lastFocusedElementBeforeModal instanceof HTMLElement) {
    lastFocusedElementBeforeModal.focus();
  }
  lastFocusedElementBeforeModal = null;
}

function handlePdfModalKeydown(event) {
  const modal = byId('modal-pdf');
  if (!modal || modal.hidden) return;

  if (event.key === 'Escape') {
    event.preventDefault();
    closePdfModal();
    return;
  }

  if (event.key !== 'Tab') return;

  const focusableElements = [...modal.querySelectorAll(MODAL_FOCUSABLE_SELECTOR)]
    .filter(element => !element.hasAttribute('hidden'));

  if (focusableElements.length === 0) return;

  const firstElement = focusableElements[0];
  const lastElement = focusableElements[focusableElements.length - 1];
  const activeElement = document.activeElement;

  if (event.shiftKey && activeElement === firstElement) {
    event.preventDefault();
    if (lastElement instanceof HTMLElement) lastElement.focus();
    return;
  }

  if (!event.shiftKey && activeElement === lastElement) {
    event.preventDefault();
    if (firstElement instanceof HTMLElement) firstElement.focus();
  }
}

let resizeTimer = null;

window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    resizeAllCharts();
    pdfPicker.handleResize();
  }, 120);
});

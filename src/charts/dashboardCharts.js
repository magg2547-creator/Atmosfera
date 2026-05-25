import { createDonutLabelPlugin } from './plugins.js';
import { CONFIG } from '../config.js';

const charts = { aq: null, co2: null, energy: null, donut: null };

let donutLabelPlugin = null;

function makeGradient(context, topColor, bottomColor, height = 220) {
  const gradient = context.createLinearGradient(0, 0, 0, height);
  gradient.addColorStop(0, topColor);
  gradient.addColorStop(1, bottomColor);
  return gradient;
}

export function getCharts() {
  return charts;
}

export function initCharts(byId, state) {
  donutLabelPlugin = createDonutLabelPlugin(() => state);
  Chart.register(donutLabelPlugin);
  Chart.defaults.font.family = "'Instrument Sans', sans-serif";
  Chart.defaults.color = '#5a6a72';
  Chart.defaults.devicePixelRatio = Math.min(window.devicePixelRatio || 1, 1.5);

  const lineChartAnimation = {
    duration: 500,
    easing: 'easeOutQuart',
  };
  const lineDecimation = {
    enabled: true,
    algorithm: 'lttb',
    samples: CONFIG.chartMaxPoints,
  };

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
          decimation: lineDecimation,
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
          decimation: lineDecimation,
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
          border: { display: false },
          grid: { color: 'rgba(90,106,114,.07)' },
          ticks: { maxTicksLimit: 6 },
        },
        y: {
          border: { display: false },
          grid: { color: 'rgba(90,106,114,.07)' },
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
      animation: lineChartAnimation,
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
          decimation: lineDecimation,
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
          border: { display: false },
          grid: { color: 'rgba(0,0,0,.04)' },
          ticks: { maxTicksLimit: 5 },
        },
        y: {
          border: { display: false },
          grid: { color: 'rgba(0,0,0,.04)' },
          ticks: { maxTicksLimit: 5 },
          title: {
            display: true,
            text: 'ppm',
            color: '#5a6a72',
            font: { size: 10 },
          },
        },
      },
      animation: lineChartAnimation,
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
          decimation: lineDecimation,
        },
        {
          label: 'Current (A)',
          data: [],
          borderColor: '#466370',
          borderWidth: 2,
          pointRadius: 0,
          fill: false,
          tension: 0.4,
          decimation: lineDecimation,
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
          decimation: lineDecimation,
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
          border: { display: false },
          grid: { color: 'rgba(0,0,0,.04)' },
          ticks: { maxTicksLimit: 6 },
        },
        y: {
          border: { display: false },
          grid: { color: 'rgba(0,0,0,.04)' },
          ticks: { maxTicksLimit: 5 },
        },
      },
      animation: lineChartAnimation,
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
      cutout: '68%',
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false },
      },
      animation: { duration: 600 },
    },
  });
}

export function getRowsForChartRange(rangeKey, state) {
  const allRows = [...state.rows].reverse();
  if (allRows.length === 0) return allRows;

  if (rangeKey === '24h') return allRows.slice(-CONFIG.chartMaxPoints);

  const latest = allRows[allRows.length - 1].time;
  if (!(latest instanceof Date) || Number.isNaN(latest.getTime())) {
    return allRows.slice(-CONFIG.chartMaxPoints);
  }

  const rangeMs = rangeKey === '1h' ? 60 * 60 * 1000 : 6 * 60 * 60 * 1000;
  const filtered = allRows.filter(row => latest.getTime() - row.time.getTime() <= rangeMs);
  return filtered.slice(-CONFIG.chartMaxPoints);
}

export function applyChartRange(rangeKey, state, DOM, formatTimeShort, setState = null) {
  if (setState) setState({ chartRange: rangeKey });

  DOM.chartTabs().forEach(tab => {
    tab.classList.toggle('active', tab.dataset.range === rangeKey);
  });

  const rows = getRowsForChartRange(rangeKey, state);
  const labels = rows.map(row => formatTimeShort(row.time));
  const hasSinglePoint = rows.length === 1;

  const applyPointVisibility = dataset => {
    if (!dataset) return;
    dataset.pointRadius = hasSinglePoint ? 3.5 : 0;
    dataset.pointHoverRadius = hasSinglePoint ? 5 : 3;
    dataset.pointHitRadius = hasSinglePoint ? 12 : 8;
  };

  if (charts.aq) {
    charts.aq.data.labels = labels;
    charts.aq.data.datasets[0].data = rows.map(row => row.pm25);
    charts.aq.data.datasets[1].data = rows.map(row => row.pm10);
    charts.aq.data.datasets[2].data = Array(rows.length).fill(15);
    charts.aq.data.datasets.forEach(applyPointVisibility);
    charts.aq.update();
  }

  if (charts.co2) {
    charts.co2.data.labels = labels;
    charts.co2.data.datasets[0].data = rows.map(row => row.co2);
    charts.co2.data.datasets.forEach(applyPointVisibility);
    charts.co2.update();
  }

  if (charts.energy) {
    charts.energy.data.labels = labels;
    charts.energy.data.datasets[0].data = rows.map(row => row.volt);
    charts.energy.data.datasets[1].data = rows.map(row => row.curr);
    charts.energy.data.datasets[2].data = rows.map(row => row.pwr);
    charts.energy.data.datasets.forEach(applyPointVisibility);
    charts.energy.update();
  }
}

export function updateCharts(state, DOM, formatTimeShort) {
  if (!state.rows.length) {
    Object.values(charts).forEach(chart => {
      if (!chart) return;
      chart.data.labels = [];
      chart.data.datasets.forEach(ds => ds.data = []);
      chart.update('none');
    });
    return;
  }
  applyChartRange(state.chartRange, state, DOM, formatTimeShort);
}

export function resizeAllCharts() {
  Object.values(charts).forEach(chart => {
    if (!chart) return;
    chart.resize();
    chart.update('none');
  });
}

export function clearCharts() {
  Object.values(charts).forEach(chart => {
    if (!chart) return;

    chart.data.labels = [];

    if (Array.isArray(chart.data.datasets)) {
      chart.data.datasets.forEach(dataset => {
        dataset.data = [];
      });
    }

    chart.update('none');
  });
}

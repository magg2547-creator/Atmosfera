export const DATE_FORMATTERS = Object.freeze({
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
  pdfDisplay: new Intl.DateTimeFormat('en-GB', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
  }),
  pdfMonth: new Intl.DateTimeFormat('en-GB', {
    month: 'long',
    year: 'numeric',
  }),
});

export function parseTimestamp(value) {
  if (!value) return new Date();
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === 'number' && Number.isFinite(value)) {
    const numericDate = new Date(value);
    if (!Number.isNaN(numericDate.getTime())) return numericDate;
  }

  const raw = String(value).trim();

  // Prefer the canonical Apps Script format: yyyy-MM-dd HH:mm:ss
  const isoLike = raw.match(
    /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (isoLike) {
    const [, year, month, day, hour = '0', minute = '0', second = '0'] = isoLike;
    const parsed = new Date(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second)
    );
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  const slashDate = raw.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?)?$/i
  );
  if (slashDate) {
    const [, first, second, yearRaw, hourRaw = '0', minute = '0', sec = '0', meridiem] = slashDate;
    const year = Number(yearRaw.length === 2 ? `20${yearRaw}` : yearRaw);
    const day = Number(first) > 12 ? Number(first) : Number(second);
    const month = Number(first) > 12 ? Number(second) : Number(first);
    let hour = Number(hourRaw);
    if (meridiem) {
      const upper = meridiem.toUpperCase();
      if (upper === 'PM' && hour < 12) hour += 12;
      if (upper === 'AM' && hour === 12) hour = 0;
    }

    const parsed = new Date(year, month - 1, day, hour, Number(minute), Number(sec));
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? new Date() : fallback;
}

export function formatTimeShort(date) {
  return DATE_FORMATTERS.timeShort.format(date);
}

export function formatTableDate(date) {
  return DATE_FORMATTERS.tableDate.format(date);
}

export function formatTableClock(date) {
  return DATE_FORMATTERS.timeShort.format(date);
}

export function formatDateTime(date) {
  return `${formatTableDate(date)} ${formatTableClock(date)}`;
}

export function formatDateInputValue(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

export function parseDateInputValue(value) {
  if (!value) return null;
  const [year, month, day] = String(value).split('-').map(Number);
  if (!year || !month || !day) return null;
  const date = new Date(year, month - 1, day);
  return Number.isNaN(date.getTime()) ? null : date;
}

export function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

export function addMonths(date, amount) {
  return new Date(date.getFullYear(), date.getMonth() + amount, 1);
}

export function isSameDay(left, right) {
  return left instanceof Date
    && right instanceof Date
    && left.getFullYear() === right.getFullYear()
    && left.getMonth() === right.getMonth()
    && left.getDate() === right.getDate();
}

export function formatPdfDisplayDate(value) {
  const date = typeof value === 'string' ? parseDateInputValue(value) : value;
  if (!date) return '-';
  return DATE_FORMATTERS.pdfDisplay.format(date);
}

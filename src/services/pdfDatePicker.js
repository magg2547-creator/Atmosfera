import {
  DATE_FORMATTERS,
  addMonths,
  formatDateInputValue,
  formatPdfDisplayDate,
  isSameDay,
  parseDateInputValue,
  startOfMonth,
} from '../utils/dateHelper.js';

const PDF_WEEKDAY_LABELS = Object.freeze(['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su']);

export function createPdfDatePicker(deps) {
  const {
    state,
    DOM,
    setText,
    getRowsMatchingCurrentSearch,
    sortRowsByTime,
    getFriendlyPdfRangeLabel,
    formatEntryCount,
    formatPdfMetric,
    TXT,
    byId,
  } = deps;

  function resetPdfDragState() {
    state.pdfPicker.drag = { active: false, anchor: '', current: '', moved: false };
  }

  function resetPdfPendingAnchor() {
    state.pdfPicker.pendingAnchor = '';
  }

  function clearPdfClickSuppression() {
    state.pdfPicker.suppressClick = false;
  }

  function armPdfClickSuppression() {
    state.pdfPicker.suppressClick = true;
    window.setTimeout(clearPdfClickSuppression, 0);
  }

  function getOrderedDateValues(left, right) {
    if (!left && !right) return ['', ''];
    if (!left) return [right, right];
    if (!right) return [left, left];
    return left <= right ? [left, right] : [right, left];
  }

  function getCalendarWeekdayIndex(date) {
    return (date.getDay() + 6) % 7;
  }

  function getPdfCalendarVisualRange() {
    if (state.pdfPicker.drag.active && state.pdfPicker.drag.anchor) {
      const [fromValue, toValue] = getOrderedDateValues(
        state.pdfPicker.drag.anchor,
        state.pdfPicker.drag.current || state.pdfPicker.drag.anchor
      );

      return {
        fromDate: parseDateInputValue(fromValue),
        toDate: parseDateInputValue(toValue),
      };
    }

    return {
      fromDate: parseDateInputValue(DOM.pdfDateFrom()?.value),
      toDate: parseDateInputValue(DOM.pdfDateTo()?.value),
    };
  }

  function getPdfPickerViewDate() {
    if (state.pdfPicker.viewDate instanceof Date && !Number.isNaN(state.pdfPicker.viewDate.getTime())) {
      return startOfMonth(state.pdfPicker.viewDate);
    }

    const anchorDate = parseDateInputValue(DOM.pdfDateFrom()?.value)
      ?? parseDateInputValue(DOM.pdfDateTo()?.value)
      ?? new Date();

    return startOfMonth(anchorDate);
  }

  function syncPdfDateCards() {
    setText(DOM.pdfDateFromDisplay(), formatPdfDisplayDate(DOM.pdfDateFrom()?.value));
    setText(DOM.pdfDateToDisplay(), formatPdfDisplayDate(DOM.pdfDateTo()?.value));

    const activeField = state.pdfPicker.pendingAnchor ? 'to' : 'from';
    DOM.pdfDateFieldButtons().forEach(button => {
      button.classList.toggle('active', button.dataset.pdfField === activeField);
    });
  }

  function renderPdfCalendarMonth(monthDate, fromDate, toDate) {
    const monthStart = startOfMonth(monthDate);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const calendarSlots = 42;

    const nextMonth = addMonths(monthStart, 1);
    const lastDayOfMonth = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), 0).getDate();
    const leadingDays = getCalendarWeekdayIndex(monthStart);

    const weekdayRow = PDF_WEEKDAY_LABELS.map(label => `<span>${label}</span>`).join('');
    const dayCells = [];

    for (let index = 0; index < leadingDays; index += 1) {
      dayCells.push('<span class="pdf-day-spacer" aria-hidden="true"></span>');
    }

    for (let day = 1; day <= lastDayOfMonth; day += 1) {
      const date = new Date(monthStart.getFullYear(), monthStart.getMonth(), day);
      const dateValue = formatDateInputValue(date);
      const isToday = isSameDay(date, today);
      const isStart = Boolean(fromDate) && isSameDay(date, fromDate);
      const isEnd = Boolean(toDate) && isSameDay(date, toDate);
      const isInRange = Boolean(fromDate && toDate && date > fromDate && date < toDate);
      const classes = ['pdf-day'];

      if (isToday) classes.push('is-today');
      if (isInRange) classes.push('is-in-range');
      if (isStart || isEnd) classes.push('is-edge');

      dayCells.push(`
        <button
          class="${classes.join(' ')}"
          type="button"
          data-date="${dateValue}"
          aria-label="${DATE_FORMATTERS.pdfDisplay.format(date)}"
        >${day}</button>
      `);
    }

    while (dayCells.length < calendarSlots) {
      dayCells.push('<span class="pdf-day-spacer" aria-hidden="true"></span>');
    }

    return `
      <section class="pdf-calendar-month">
        <div class="pdf-calendar-month-label">${DATE_FORMATTERS.pdfMonth.format(monthStart)}</div>
        <div class="pdf-calendar-weekdays">${weekdayRow}</div>
        <div class="pdf-calendar-days">${dayCells.join('')}</div>
      </section>
    `;
  }

  function renderPdfCalendar() {
    const viewDate = getPdfPickerViewDate();
    const isMobile = window.innerWidth < 600;
    const nextMonth = addMonths(viewDate, 1);
    const { fromDate, toDate } = getPdfCalendarVisualRange();

    state.pdfPicker.viewDate = viewDate;

    const titleText = isMobile
      ? DATE_FORMATTERS.pdfMonth.format(viewDate)
      : `${DATE_FORMATTERS.pdfMonth.format(viewDate)} \u2014 ${DATE_FORMATTERS.pdfMonth.format(nextMonth)}`;

    setText(DOM.pdfCalendarTitle(), titleText);
    setText(
      DOM.pdfCalendarSubtitle(),
      state.pdfPicker.drag.active
        ? 'Release to confirm the range.'
        : state.pdfPicker.pendingAnchor
          ? 'Tap another date to extend the range.'
          : 'Tap one date or drag across days.'
    );

    const months = DOM.pdfCalendarMonths();
    if (!months) return;

    if (isMobile) {
      months.innerHTML = renderPdfCalendarMonth(viewDate, fromDate, toDate);
    } else {
      months.innerHTML = [viewDate, nextMonth]
        .map(monthDate => renderPdfCalendarMonth(monthDate, fromDate, toDate))
        .join('');
    }
  }

  function cancelScheduledPdfCalendarRender() {
    if (!state.pdfPicker.renderFrame) return;
    cancelAnimationFrame(state.pdfPicker.renderFrame);
    state.pdfPicker.renderFrame = 0;
  }

  function schedulePdfCalendarRender() {
    if (state.pdfPicker.renderFrame) return;

    state.pdfPicker.renderFrame = requestAnimationFrame(() => {
      state.pdfPicker.renderFrame = 0;
      renderPdfCalendar();
    });
  }

  function refreshPdfPickerUi() {
    syncPdfDateCards();
    cancelScheduledPdfCalendarRender();
    renderPdfCalendar();
    updatePdfModalSummary();
  }

  function setPdfModalDates(fromValue, toValue, options = {}) {
    const {
      anchorValue = fromValue || toValue,
      pendingAnchor = '',
    } = options;

    const fromInput = DOM.pdfDateFrom();
    const toInput = DOM.pdfDateTo();
    if (fromInput) fromInput.value = fromValue || '';
    if (toInput) toInput.value = toValue || '';

    state.pdfPicker.activeField = pendingAnchor ? 'to' : 'from';
    state.pdfPicker.pendingAnchor = pendingAnchor || '';
    state.pdfPicker.viewDate = startOfMonth(
      parseDateInputValue(anchorValue)
      ?? parseDateInputValue(fromValue)
      ?? parseDateInputValue(toValue)
      ?? new Date()
    );
    resetPdfDragState();

    refreshPdfPickerUi();
  }

  function handlePdfCalendarDateSelection(value) {
    const pendingAnchor = state.pdfPicker.pendingAnchor;

    if (pendingAnchor) {
      const [nextFrom, nextTo] = getOrderedDateValues(pendingAnchor, value);
      syncPdfPresetButtons(null);
      setPdfModalDates(nextFrom, nextTo, {
        anchorValue: nextFrom,
        pendingAnchor: '',
      });
      return;
    }

    syncPdfPresetButtons(null);
    setPdfModalDates(value, value, {
      anchorValue: value,
      pendingAnchor: value,
    });
  }

  function beginPdfCalendarDrag(value) {
    if (!value) return;
    clearPdfClickSuppression();
    resetPdfPendingAnchor();
    state.pdfPicker.drag = {
      active: true,
      anchor: value,
      current: value,
      moved: false,
    };
    syncPdfPresetButtons(null);
    renderPdfCalendar();
  }

  function updatePdfCalendarDrag(value) {
    if (!state.pdfPicker.drag.active || !value) return;
    if (value === state.pdfPicker.drag.current) return;

    state.pdfPicker.drag.current = value;
    state.pdfPicker.drag.moved = true;
    schedulePdfCalendarRender();
  }

  function finishPdfCalendarDrag() {
    if (!state.pdfPicker.drag.active) return;

    const { anchor, current, moved } = state.pdfPicker.drag;

    if (!moved) {
      resetPdfDragState();
      return;
    }

    armPdfClickSuppression();
    cancelScheduledPdfCalendarRender();
    const [fromValue, toValue] = getOrderedDateValues(anchor, current || anchor);
    setPdfModalDates(fromValue, toValue, {
      anchorValue: fromValue,
      pendingAnchor: '',
    });
  }

  function cancelPdfCalendarDrag() {
    if (!state.pdfPicker.drag.active) return;
    cancelScheduledPdfCalendarRender();
    resetPdfDragState();
  }

  function syncPdfPresetButtons(activePreset) {
    DOM.pdfPresetButtons().forEach(btn => {
      btn.classList.toggle('active', btn.dataset.pdfPreset === activePreset);
    });
  }

  function applyPdfPreset(preset) {
    const today = new Date();
    const from = new Date(today);
    const to = formatDateInputValue(today);

    if (preset === 'today') {
      setPdfModalDates(to, to, { anchorValue: to, pendingAnchor: '' });
    } else if (preset === '7d') {
      from.setDate(from.getDate() - 6);
      setPdfModalDates(formatDateInputValue(from), to, {
        anchorValue: formatDateInputValue(from),
        pendingAnchor: '',
      });
    } else if (preset === '30d') {
      from.setDate(from.getDate() - 29);
      setPdfModalDates(formatDateInputValue(from), to, {
        anchorValue: formatDateInputValue(from),
        pendingAnchor: '',
      });
    } else {
      setPdfModalDates('', '', { anchorValue: formatDateInputValue(today), pendingAnchor: '' });
    }

    syncPdfPresetButtons(preset);
  }

  function getPdfModalSelection() {
    const fromVal = DOM.pdfDateFrom()?.value;
    const toVal = DOM.pdfDateTo()?.value;
    const baseRows = getRowsMatchingCurrentSearch(state.rows);

    if (!fromVal || !toVal) {
      return {
        rows: sortRowsByTime(baseRows, 'asc'),
        rangeLabel: getFriendlyPdfRangeLabel(fromVal, toVal),
        fromValue: fromVal || '',
        toValue: toVal || '',
      };
    }

    const from = parseDateInputValue(fromVal);
    const to = parseDateInputValue(toVal);
    if (!from || !to) {
      return {
        rows: sortRowsByTime(baseRows, 'asc'),
        rangeLabel: getFriendlyPdfRangeLabel(fromVal, toVal),
        fromValue: fromVal || '',
        toValue: toVal || '',
      };
    }

    to.setHours(23, 59, 59, 999);

    return {
      rows: sortRowsByTime(baseRows.filter(row => row.time >= from && row.time <= to), 'asc'),
      rangeLabel: getFriendlyPdfRangeLabel(fromVal, toVal),
      fromValue: fromVal,
      toValue: toVal,
    };
  }

  function getPdfModalRows() {
    return getPdfModalSelection().rows;
  }

  function updatePdfHealthSummary(rows) {
    if (!rows.length) {
      setText(DOM.pdfHealthPm25(), TXT.emDash);
      setText(DOM.pdfHealthCo2(), TXT.emDash);
      setText(DOM.pdfHealthEnergy(), TXT.emDash);
      return;
    }

    const avgPm25 = rows.reduce((sum, row) => sum + row.pm25, 0) / rows.length;
    const maxCo2 = rows.reduce((max, row) => Math.max(max, row.co2), 0);
    const energyValues = rows.map(row => row.energy).filter(Number.isFinite);
    const totalEnergy = energyValues.length > 0
      ? Math.max(...energyValues) - Math.min(...energyValues)
      : 0;

    setText(DOM.pdfHealthPm25(), formatPdfMetric(avgPm25, 1, TXT.microgPerM3));
    setText(DOM.pdfHealthCo2(), formatPdfMetric(maxCo2, 0, 'ppm'));
    setText(DOM.pdfHealthEnergy(), formatPdfMetric(totalEnergy, 2, 'kWh'));
  }

  function updatePdfModalSummary() {
    const { rows, rangeLabel } = getPdfModalSelection();
    const summary = byId('pdf-modal-summary');

    updatePdfHealthSummary(rows);
    if (!summary) return;

    const fromVal = DOM.pdfDateFrom()?.value;
    const toVal = DOM.pdfDateTo()?.value;

    if (!fromVal || !toVal) {
      summary.textContent = `All matching records ${TXT.emDash} ${formatEntryCount(rows.length)}`;
      return;
    }

    if (rows.length === 0) {
      summary.textContent = 'No records found in this date range';
      return;
    }

    summary.textContent = `${rangeLabel} ${TXT.emDash} ${formatEntryCount(rows.length)}`;
  }

  function bindCalendarEvents() {
    DOM.pdfDateFieldButtons().forEach(button => {
      button.addEventListener('click', () => {
        const value = button.dataset.pdfField === 'to'
          ? DOM.pdfDateTo()?.value
          : DOM.pdfDateFrom()?.value;

        if (!value) return;

        syncPdfPresetButtons(null);
        setPdfModalDates(value, value, {
          anchorValue: value,
          pendingAnchor: value,
        });
      });
    });

    DOM.btnPdfPrevMonth()?.addEventListener('click', () => {
      state.pdfPicker.viewDate = addMonths(getPdfPickerViewDate(), -1);
      renderPdfCalendar();
    });

    DOM.btnPdfNextMonth()?.addEventListener('click', () => {
      state.pdfPicker.viewDate = addMonths(getPdfPickerViewDate(), 1);
      renderPdfCalendar();
    });

    DOM.pdfCalendarMonths()?.addEventListener('pointerdown', event => {
      const button = event.target.closest('.pdf-day[data-date]');
      if (!button) return;
      beginPdfCalendarDrag(button.dataset.date);
    });

    DOM.pdfCalendarMonths()?.addEventListener('pointerover', event => {
      const button = event.target.closest('.pdf-day[data-date]');
      if (!button) return;
      updatePdfCalendarDrag(button.dataset.date);
    });

    DOM.pdfCalendarMonths()?.addEventListener('pointermove', event => {
      if (!state.pdfPicker.drag.active) return;
      const hovered = document.elementFromPoint(event.clientX, event.clientY)?.closest('.pdf-day[data-date]');
      if (!hovered) return;
      updatePdfCalendarDrag(hovered.dataset.date);
    });

    DOM.pdfCalendarMonths()?.addEventListener('click', event => {
      const button = event.target.closest('.pdf-day[data-date]');
      if (!button) return;
      if (state.pdfPicker.suppressClick) {
        clearPdfClickSuppression();
        return;
      }
      handlePdfCalendarDateSelection(button.dataset.date);
    });

    window.addEventListener('pointerup', finishPdfCalendarDrag);
    window.addEventListener('pointercancel', cancelPdfCalendarDrag);

    DOM.pdfPresetButtons().forEach(btn => {
      btn.addEventListener('click', () => applyPdfPreset(btn.dataset.pdfPreset));
    });
  }

  function unbindCalendarEvents() {
    window.removeEventListener('pointerup', finishPdfCalendarDrag);
    window.removeEventListener('pointercancel', cancelPdfCalendarDrag);
  }

  function handleResize() {
    const modal = DOM.pdfDateFrom()?.closest?.('#modal-pdf') ?? document.getElementById('modal-pdf');
    if (modal && !modal.hidden) renderPdfCalendar();
  }

  return {
    bindCalendarEvents,
    unbindCalendarEvents,
    handleResize,
    getPdfPickerViewDate,
    renderPdfCalendar,
    setPdfModalDates,
    syncPdfPresetButtons,
    applyPdfPreset,
    beginPdfCalendarDrag,
    updatePdfCalendarDrag,
    finishPdfCalendarDrag,
    cancelPdfCalendarDrag,
    clearPdfClickSuppression,
    resetPdfDragState,
    resetPdfPendingAnchor,
    cancelScheduledPdfCalendarRender,
    getPdfModalSelection,
    getPdfModalRows,
  };
}

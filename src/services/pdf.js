export function createPdfService(context) {
  const {
    byId,
    fmt,
    formatDateInputValue,
    formatDateTime,
    getExportRows,
    getLatestRowByTime,
    getSelectedRangeLabel,
    showToast,
    sortRowsByTime,
  } = context;

function toCsvCell(value) {
  return `"${String(value ?? '').replace(/"/g, '""')}"`;
}


function getPdfFilename(options = {}) {
  const today = formatDateInputValue(new Date());
  const fromValue = options.fromValue || '';
  const toValue = options.toValue || '';

  if (fromValue && toValue) {
    if (fromValue === toValue) return `atmosfera_report_${fromValue}.pdf`;
    return `atmosfera_report_${fromValue}_to_${toValue}.pdf`;
  }

  return `atmosfera_report_all_${today}.pdf`;
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

// â”€â”€ Download helper (All Platforms) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function fallbackDownload(blob, filename) {
  // IE11 / Legacy Edge fallback
  if (window.navigator && window.navigator.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(blob, filename);
    return;
  }
  
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.style.display = 'none';
  link.href = url;
  link.download = filename;
  link.rel = 'noopener'; // Security & browser context bypass
  
  document.body.appendChild(link);
  // Optional: Check if we need to force octet-stream for stubborn browsers
  link.click();
  
  // Cleanup â€” Give the browser enough time to process the download
  setTimeout(() => {
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }, 10000);
}

function setPdfExportBusy(isBusy) {
  const button = byId('btn-modal-confirm');
  if (!button) return;

  button.disabled = isBusy;
  button.classList.toggle('is-loading', isBusy);
  button.textContent = isBusy ? 'Preparing...' : 'Export PDF';
}

function waitForPaint() {
  return new Promise(resolve => {
    requestAnimationFrame(() => requestAnimationFrame(resolve));
  });
}

function exportAllCSV() {
  const rows = getExportRows();
  if (rows.length === 0) {
    showToast('Nothing to export');
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

  const filename = `atmosfera_export_${formatDateInputValue(new Date())}.csv`;
  const csvBlob  = new Blob([csvRows.join('\n')], { type: 'text/csv;charset=utf-8' });

  // iOS Safari: à¸•à¹‰à¸­à¸‡à¹ƒà¸Šà¹‰ Web Share API à¹€à¸žà¸·à¹ˆà¸­à¹ƒà¸«à¹‰à¸Šà¸·à¹ˆà¸­à¹„à¸Ÿà¸¥à¹Œà¹à¸¥à¸°à¸™à¸²à¸¡à¸ªà¸à¸¸à¸¥à¸–à¸¹à¸à¸•à¹‰à¸­à¸‡
  const isMobile = /Mobi|Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
  if (isMobile && navigator.share && navigator.canShare) {
    const csvFile = new File([csvBlob], filename, { type: 'text/csv' });
    if (navigator.canShare({ files: [csvFile] })) {
      navigator.share({ files: [csvFile], title: 'Atmosfera Export' })
        .catch(() => fallbackDownload(csvBlob, filename));
      showToast(`CSV ready (${rows.length} rows)`);
      return;
    }
  }

  // Desktop / Android à¹à¸šà¸šà¹€à¸à¹ˆà¸²: download à¸›à¸à¸•à¸´
  fallbackDownload(csvBlob, filename);
  showToast(`CSV ready (${rows.length} rows)`);
}

function exportPDF(rows, options = {}) {
  if (!Array.isArray(rows)) {
    options = rows ?? {};
    rows = getExportRows();
  }

  const exportRows = sortRowsByTime(rows, 'asc');
  if (exportRows.length === 0) {
    showToast('No records in selected range');
    return false;
  }

  const jsPDFCtor = window.jspdf?.jsPDF;
  if (!jsPDFCtor) {
    showToast('PDF export unavailable');
    return false;
  }

  const doc = new jsPDFCtor({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 12;
  const rowHeight = 6;
  const tableWidth = pageWidth - (margin * 2);
  const rangeLabel = options.rangeLabel || getSelectedRangeLabel();
  const snapshotRow = getLatestRowByTime(exportRows);
  const columns = [
    { label: 'Time', width: 36, align: 'left' },
    { label: 'PM2.5', width: 24, align: 'right' },
    { label: 'PM10', width: 24, align: 'right' },
    { label: 'Temp C', width: 24, align: 'right' },
    { label: 'Hum %', width: 22, align: 'right' },
    { label: 'CO2', width: 24, align: 'right' },
    { label: 'Volt V', width: 23, align: 'right' },
    { label: 'Curr A', width: 23, align: 'right' },
    { label: 'Power W', width: 25, align: 'right' },
    { label: 'Energy', width: 28, align: 'right' },
  ];

  const palette = {
    ink: [32, 50, 57],
    muted: [85, 112, 122],
    brand: [63, 111, 121],
    accent: [25, 121, 111],
    paper: [248, 251, 251],
    surface: [238, 245, 246],
    line: [211, 224, 227],
  };

  const fitPdfText = (value, maxLength) => {
    const text = String(value);
    return text.length > maxLength ? `${text.slice(0, maxLength - 1)}...` : text;
  };

  const drawPill = (label, value, x, y, width) => {
    doc.setFillColor(...palette.surface);
    doc.setDrawColor(...palette.line);
    doc.roundedRect(x, y, width, 12, 3, 3, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(6.8);
    doc.setTextColor(...palette.muted);
    doc.text(label.toUpperCase(), x + 4, y + 4.4);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.2);
    doc.setTextColor(...palette.ink);
    doc.text(fitPdfText(value, Math.max(8, Math.floor(width / 3))), x + 4, y + 9.2);
  };

  const drawMetricCard = (label, value, x, y, width) => {
    doc.setFillColor(248, 251, 251);
    doc.setDrawColor(...palette.line);
    doc.roundedRect(x, y, width, 18, 4, 4, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7);
    doc.setTextColor(...palette.muted);
    doc.text(label, x + 4, y + 5.3);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(...palette.ink);
    doc.text(String(value), x + 4, y + 13.2);
  };

  const drawTableHeader = y => {
    let x = margin;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setFillColor(...palette.ink);
    doc.roundedRect(margin, y - 5, tableWidth, rowHeight + 1, 2.5, 2.5, 'F');
    doc.setTextColor(248, 251, 251);

    columns.forEach(column => {
      const textX = column.align === 'right' ? x + column.width - 2 : x + 2;
      doc.text(column.label, textX, y, { align: column.align });
      x += column.width;
    });
  };

  const drawPageHeader = pageNumber => {
    doc.setFillColor(...palette.paper);
    doc.rect(0, 0, pageWidth, pageHeight, 'F');

    doc.setFillColor(...palette.ink);
    doc.rect(0, 0, pageWidth, 24, 'F');
    doc.setFillColor(...palette.accent);
    doc.rect(0, 0, 78, 24, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(248, 251, 251);
    doc.text('ATMOSFERA', margin, 10);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text('Air & Energy Monitor', margin, 16);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('Monitoring Report', pageWidth - margin, 10, { align: 'right' });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.text(`Generated ${formatDateTime(new Date())}`, pageWidth - margin, 16, { align: 'right' });

    doc.setTextColor(...palette.ink);
    if (pageNumber === 1) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(17);
      doc.text('Environmental snapshot', margin, 36);

      drawPill('Range', rangeLabel, margin, 42, 96);
      drawPill('Records', exportRows.length, margin + 100, 42, 44);
      drawPill('Page', pageNumber, margin + 148, 42, 32);

      if (snapshotRow) {
        const metrics = [
          ['PM2.5', fmt(snapshotRow.pm25)],
          ['PM10', fmt(snapshotRow.pm10)],
          ['Temp', `${fmt(snapshotRow.temp)} C`],
          ['Humidity', `${fmt(snapshotRow.hum, 0)} %`],
          ['CO2', fmt(snapshotRow.co2, 0)],
          ['Power', `${fmt(snapshotRow.pwr, 1)} W`],
        ];
        const cardWidth = (tableWidth - 25) / 6;
        metrics.forEach(([label, value], index) => {
          drawMetricCard(label, value, margin + (index * (cardWidth + 5)), 60, cardWidth);
        });
      }
    } else {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.text(`Monitoring Report / ${rangeLabel}`, margin, 34);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.setTextColor(...palette.muted);
      doc.text(`Page ${pageNumber}`, pageWidth - margin, 34, { align: 'right' });
    }
  };

  const drawPageFooter = pageNumber => {
    doc.setDrawColor(...palette.line);
    doc.line(margin, pageHeight - 10, pageWidth - margin, pageHeight - 10);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(...palette.muted);
    doc.text('Atmosfera Environmental Intelligence', margin, pageHeight - 5);
    doc.text(`Page ${pageNumber}`, pageWidth - margin, pageHeight - 5, { align: 'right' });
  };

  const drawRow = (row, y, index) => {
    const values = getExportRowValues(row);
    let x = margin;

    if (index % 2 === 0) {
      doc.setFillColor(244, 249, 249);
      doc.roundedRect(margin, y - 4.2, tableWidth, rowHeight, 1.5, 1.5, 'F');
    }

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.6);
    doc.setTextColor(...palette.ink);

    values.forEach((value, index) => {
      const column = columns[index];
      const textX = column.align === 'right' ? x + column.width - 2 : x + 2;
      doc.text(String(value), textX, y, { align: column.align });
      x += column.width;
    });

    doc.setDrawColor(224, 233, 235);
    doc.line(margin, y + 1.8, pageWidth - margin, y + 1.8);
  };

  let pageNumber = 1;
  let y = margin;
  drawPageHeader(pageNumber);
  y = 86;
  drawTableHeader(y);
  y += 8;

  exportRows.forEach((row, index) => {
    if (y > pageHeight - 16) {
      drawPageFooter(pageNumber);
      doc.addPage();
      pageNumber += 1;
      drawPageHeader(pageNumber);
      y = 44;
      drawTableHeader(y);
      y += 8;
    }

    drawRow(row, y, index);
    y += rowHeight;
  });
  drawPageFooter(pageNumber);

  const pdfFilename = getPdfFilename(options);
  const pdfBlob     = doc.output('blob');

  // iOS: à¹ƒà¸Šà¹‰ Web Share API à¹€à¸žà¸·à¹ˆà¸­à¸£à¸±à¸à¸©à¸²à¸Šà¸·à¹ˆà¸­à¹„à¸Ÿà¸¥à¹Œ .pdf à¹„à¸¡à¹ˆà¹ƒà¸«à¹‰à¸à¸¥à¸²à¸¢à¹€à¸›à¹‡à¸™à¹„à¸Ÿà¸¥à¹Œà¸›à¸£à¸°à¸«à¸¥à¸²à¸” (Limitation à¸‚à¸­à¸‡ iOS)
  const isMobile = /Mobi|Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
  if (isMobile && navigator.share && navigator.canShare) {
    const pdfFile = new File([pdfBlob], pdfFilename, { type: 'application/pdf' });
    if (navigator.canShare({ files: [pdfFile] })) {
      navigator.share({ files: [pdfFile], title: 'Atmosfera Report' })
        .catch(() => fallbackDownload(pdfBlob, pdfFilename));
      showToast(`PDF ready (${exportRows.length} rows)`);
      return true;
    }
  }

  // Desktop: download à¸›à¸à¸•à¸´ (jsPDF handle)
  doc.save(pdfFilename);
  showToast(`PDF ready (${exportRows.length} rows)`);
  return true;
}


  return {
    exportAllCSV,
    exportPDF,
    setPdfExportBusy,
    waitForPaint,
  };
}

/**
 * platform-comparison-transformer.js
 *
 * Transforms a Platform Analytics export (Android / iOS / Web) into a polished
 * workbook with:
 *   - The original source sheets preserved
 *   - A normalized "Daily Data" long-format pivot source
 *   - A side-by-side "Platform Comparison" summary
 *   - Visual comparison sheets (Volume, Conversion, Daily Trend, App Funnel, Web Funnel)
 *   - Live SUMIFS / AVERAGEIFS formulas throughout
 *   - Embedded bar and line charts (rendered by Excel/LibreOffice natively)
 *
 * Runs fully client-side. No server, no upload. The file never leaves the browser.
 *
 * Dependencies (both MIT, loaded from CDN on first use):
 *   - SheetJS (xlsx)       -> reads the source .xlsx
 *   - ExcelJS              -> writes the output .xlsx base (formulas, styling)
 *   - JSZip                -> post-processes the output to inject chart XML
 *                             (ExcelJS 4.x can't write charts natively, so we
 *                              splice them in after .xlsx.writeBuffer()).
 *
 * Usage:
 *   import { transformWorkbook, transformAndDownload } from './platform-comparison-transformer.js';
 *   const outBlob = await transformWorkbook(inputFile);  // File | Blob | ArrayBuffer -> Blob
 *   // or, one-call convenience:
 *   await transformAndDownload(inputFile);               // triggers browser download
 */

// =============================================================================
// Constants
// =============================================================================

const N_DAYS = 91;

const SOURCE_SHEETS = {
  androidFunnel: { name: 'Android purchase funnel',         startRow: 3 },
  iosFunnel:     { name: 'iOS purchase funnel',             startRow: 3 },
  webFlow:       { name: 'Web flow',                        startRow: 2 },
  iosEsim:       { name: 'iOS eSIM setup flow',             startRow: 2 },
  androidEsim:   { name: 'Android eSIM setup flow',         startRow: 2 },
  dailyCR:       { name: 'Daily CR (eSIM compatible devic', startRow: 3 },
  webVisitors:   { name: 'Web Daily Visitors',              startRow: 2 },
  dailySessions: { name: 'Daily Sessions',                  startRow: 3 },
};

const STYLE = {
  headerFont: { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
  headerFill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF305496' } },
  titleFont:  { name: 'Arial', bold: true, size: 16, color: { argb: 'FF1F4E78' } },
  subtitleFont: { name: 'Arial', italic: true, size: 10, color: { argb: 'FF595959' } },
  sectionFill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } },
  sectionFont: { name: 'Arial', bold: true, color: { argb: 'FF1F4E78' } },
  bodyFont: { name: 'Arial', size: 11 },
  bodyBoldFont: { name: 'Arial', bold: true, size: 11 },
  mutedFont: { name: 'Arial', size: 11, color: { argb: 'FF808080' } },
  thinBorder: {
    top:    { style: 'thin', color: { argb: 'FFBFBFBF' } },
    left:   { style: 'thin', color: { argb: 'FFBFBFBF' } },
    bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
    right:  { style: 'thin', color: { argb: 'FFBFBFBF' } },
  },
  numFmt: '#,##0;(#,##0);-',
  pctFmt: '0.00%;-0.00%;-',
  pctFmt1: '0.0%;-0.0%;-',
};

// =============================================================================
// CDN dependency loading (idempotent)
// =============================================================================

let _depsPromise = null;

function ensureDependencies() {
  if (_depsPromise) return _depsPromise;

  _depsPromise = (async () => {
    const loadScript = (src) => new Promise((resolve, reject) => {
      if (document.querySelector(`script[data-src="${src}"]`)) { resolve(); return; }
      const s = document.createElement('script');
      s.src = src;
      s.dataset.src = src;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error(`Failed to load ${src}`));
      document.head.appendChild(s);
    });

    if (typeof window.XLSX === 'undefined') {
      await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
    }
    if (typeof window.ExcelJS === 'undefined') {
      await loadScript('https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js');
    }
    // ExcelJS bundles JSZip internally, but we need the standalone one to
    // repack the final .xlsx after chart injection.
    if (typeof window.JSZip === 'undefined') {
      await loadScript('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js');
    }

    return { XLSX: window.XLSX, ExcelJS: window.ExcelJS, JSZip: window.JSZip };
  })();

  return _depsPromise;
}

// =============================================================================
// Source reading
// =============================================================================

function readSourceSheets(XLSX, arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });

  const readRows = (sheetName, startRow) => {
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`Missing sheet in source file: "${sheetName}"`);
    const all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
    return all.slice(startRow - 1).filter(r => r.some(v => v !== null && v !== ''));
  };

  const out = {};
  for (const [key, cfg] of Object.entries(SOURCE_SHEETS)) {
    out[key] = readRows(cfg.name, cfg.startRow);
  }
  return out;
}

// =============================================================================
// Build normalized records
// =============================================================================

function buildDailyDataRecords(src) {
  const get = (arr, i) => (i < arr.length ? arr[i] : []);
  const pick = (arr, idx) => (arr && arr[idx] != null ? arr[idx] : null);

  const records = [];
  for (let i = 0; i < N_DAYS; i++) {
    const day = i + 1;
    const a = get(src.androidFunnel, i);
    const o = get(src.iosFunnel, i);
    const w = get(src.webFlow, i);
    const ae = get(src.androidEsim, i);
    const oe = get(src.iosEsim, i);
    const cr = get(src.dailyCR, i);
    const wv = get(src.webVisitors, i);

    records.push({
      Day: day, Platform: 'Android',
      Sessions: pick(a, 0), ExplorePlans: pick(a, 1), LoginClick: pick(a, 2),
      LoginSuccess: pick(a, 3), ViewCountries: pick(a, 4), SelectedPlan: pick(a, 5),
      ContinueCheckout: pick(a, 6), WebLanding: null, WebPricing: null, WebCheckout: null,
      Transactions: pick(a, 7), NewUsers: null, ReturningUsers: null,
      EsimClick: pick(ae, 1), EsimSuccess: pick(ae, 2), EsimFailed: pick(ae, 3),
      CRSource: pick(cr, 0),
    });

    records.push({
      Day: day, Platform: 'iOS',
      Sessions: pick(o, 0), ExplorePlans: pick(o, 1), LoginClick: pick(o, 2),
      LoginSuccess: pick(o, 3), ViewCountries: pick(o, 4), SelectedPlan: pick(o, 5),
      ContinueCheckout: pick(o, 6), WebLanding: null, WebPricing: null, WebCheckout: null,
      Transactions: pick(o, 7), NewUsers: null, ReturningUsers: null,
      EsimClick: pick(oe, 1), EsimSuccess: pick(oe, 2), EsimFailed: pick(oe, 3),
      CRSource: pick(cr, 1),
    });

    records.push({
      Day: day, Platform: 'Web',
      Sessions: null, ExplorePlans: null, LoginClick: null, LoginSuccess: null,
      ViewCountries: null, SelectedPlan: null, ContinueCheckout: null,
      WebLanding: pick(w, 0), WebPricing: pick(w, 1), WebCheckout: pick(w, 2),
      Transactions: pick(w, 3), NewUsers: pick(wv, 0), ReturningUsers: pick(wv, 1),
      EsimClick: null, EsimSuccess: null, EsimFailed: null,
      CRSource: pick(cr, 2),
    });
  }
  return records;
}

// =============================================================================
// ExcelJS sheet builders
// =============================================================================

const DAILY_COLUMNS = [
  { key: 'Day', header: 'Day', width: 6 },
  { key: 'Platform', header: 'Platform', width: 10 },
  { key: 'Sessions', header: 'Sessions', width: 16 },
  { key: 'ExplorePlans', header: 'Explore Plans Click', width: 16 },
  { key: 'LoginClick', header: 'Login/Signup Click', width: 16 },
  { key: 'LoginSuccess', header: 'Login/Signup Success', width: 16 },
  { key: 'ViewCountries', header: 'View Countries', width: 16 },
  { key: 'SelectedPlan', header: 'Selected Plan', width: 16 },
  { key: 'ContinueCheckout', header: 'Continue to Checkout', width: 16 },
  { key: 'WebLanding', header: 'Web Landing Page', width: 16 },
  { key: 'WebPricing', header: 'Web Pricing', width: 16 },
  { key: 'WebCheckout', header: 'Web Checkout', width: 16 },
  { key: 'Transactions', header: 'Transactions', width: 16 },
  { key: 'NewUsers', header: 'New Users', width: 16 },
  { key: 'ReturningUsers', header: 'Returning Users', width: 16 },
  { key: 'EsimClick', header: 'eSIM Install Click', width: 16 },
  { key: 'EsimSuccess', header: 'eSIM Install Success', width: 16 },
  { key: 'EsimFailed', header: 'eSIM Install Failed', width: 16 },
  { key: 'CRSource', header: 'Conversion Rate (source)', width: 16 },
];

function styleHeaderRow(sheet, rowNum, colStart, colEnd) {
  for (let c = colStart; c <= colEnd; c++) {
    const cell = sheet.getCell(rowNum, c);
    cell.font = STYLE.headerFont;
    cell.fill = STYLE.headerFill;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = STYLE.thinBorder;
  }
}

function applyNum(cell, fmt = STYLE.numFmt, bold = false) {
  cell.font = bold ? STYLE.bodyBoldFont : STYLE.bodyFont;
  cell.numFmt = fmt;
  cell.alignment = { horizontal: 'right' };
  cell.border = STYLE.thinBorder;
}

function applyLabel(cell, bold = false) {
  cell.font = bold ? STYLE.bodyBoldFont : STYLE.bodyFont;
  cell.alignment = { horizontal: 'left' };
  cell.border = STYLE.thinBorder;
}

// ----- Daily Data -------------------------------------------------------------

function buildDailyDataSheet(workbook, records) {
  const sheet = workbook.addWorksheet('Daily Data', {
    views: [{ state: 'frozen', xSplit: 2, ySplit: 1 }],
  });

  sheet.getRow(1).values = DAILY_COLUMNS.map(c => c.header);
  styleHeaderRow(sheet, 1, 1, DAILY_COLUMNS.length);
  sheet.getRow(1).height = 32;
  DAILY_COLUMNS.forEach((c, i) => { sheet.getColumn(i + 1).width = c.width; });

  records.forEach((r, i) => {
    const rowNum = i + 2;
    DAILY_COLUMNS.forEach((col, ci) => {
      const cell = sheet.getCell(rowNum, ci + 1);
      cell.value = r[col.key];
      if (col.key === 'Day' || col.key === 'Platform') {
        cell.font = STYLE.bodyFont;
        cell.alignment = { horizontal: 'center' };
        cell.border = STYLE.thinBorder;
      } else {
        const fmt = (col.key === 'CRSource') ? STYLE.pctFmt : STYLE.numFmt;
        applyNum(cell, fmt);
      }
    });
  });

  sheet.addTable({
    name: 'DailyData',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: DAILY_COLUMNS.map(c => ({ name: c.header, filterButton: true })),
    rows: records.map(r => DAILY_COLUMNS.map(col => r[col.key])),
  });
}

// ----- Platform Comparison ----------------------------------------------------

function buildPlatformComparisonSheet(workbook) {
  const sheet = workbook.addWorksheet('Platform Comparison', { views: [{ showGridLines: false }] });

  sheet.getCell('B2').value = 'Platform Comparison Summary';
  sheet.getCell('B2').font = STYLE.titleFont;
  sheet.mergeCells('B2:E2');
  sheet.getCell('B3').value = `Data spans ${N_DAYS} days. All figures computed with formulas from the Daily Data sheet.`;
  sheet.getCell('B3').font = STYLE.subtitleFont;
  sheet.mergeCells('B3:E3');

  ['Metric', 'Android', 'iOS', 'Web'].forEach((h, i) => { sheet.getCell(5, 2 + i).value = h; });
  styleHeaderRow(sheet, 5, 2, 5);

  const sumifs = (col, plat) => ({ formula: `SUMIFS('Daily Data'!${col}:${col},'Daily Data'!B:B,"${plat}")` });
  const averageifs = (col, plat) => ({ formula: `IFERROR(AVERAGEIFS('Daily Data'!${col}:${col},'Daily Data'!B:B,"${plat}"),0)` });

  const sections = [
    { label: '— VOLUME (totals) —', header: true },
    { label: 'Total Sessions', col: 'C', kind: 'sum' },
    { label: 'Total Transactions', col: 'M', kind: 'sum' },
    { label: 'Total Web Landing Page', col: 'J', kind: 'sum' },
    { label: 'Total Web Pricing', col: 'K', kind: 'sum' },
    { label: 'Total Web Checkout', col: 'L', kind: 'sum' },
    { label: 'Total New Users (Web)', col: 'N', kind: 'sum' },
    { label: 'Total Returning Users (Web)', col: 'O', kind: 'sum' },
    { label: '— DAILY AVERAGES —', header: true },
    { label: 'Avg Daily Sessions', col: 'C', kind: 'avg' },
    { label: 'Avg Daily Transactions', col: 'M', kind: 'avg' },
    { label: '— APP PURCHASE FUNNEL TOTALS (Android/iOS) —', header: true },
    { label: 'Explore Plans Click', col: 'D', kind: 'sum' },
    { label: 'Login/Signup Click', col: 'E', kind: 'sum' },
    { label: 'Login/Signup Success', col: 'F', kind: 'sum' },
    { label: 'View Countries', col: 'G', kind: 'sum' },
    { label: 'Selected Plan', col: 'H', kind: 'sum' },
    { label: 'Continue to Checkout', col: 'I', kind: 'sum' },
    { label: '— eSIM SETUP TOTALS (Android/iOS) —', header: true },
    { label: 'eSIM Install Click', col: 'P', kind: 'sum' },
    { label: 'eSIM Install Success', col: 'Q', kind: 'sum' },
    { label: 'eSIM Install Failed', col: 'R', kind: 'sum' },
  ];

  const labelRow = {};
  let row = 6;
  for (const s of sections) {
    sheet.getCell(row, 2).value = s.label;
    if (s.header) {
      for (let c = 2; c <= 5; c++) {
        const cell = sheet.getCell(row, c);
        cell.fill = STYLE.sectionFill;
        cell.font = STYLE.sectionFont;
        cell.border = STYLE.thinBorder;
      }
    } else {
      applyLabel(sheet.getCell(row, 2));
      ['Android', 'iOS', 'Web'].forEach((plat, i) => {
        const cell = sheet.getCell(row, 3 + i);
        cell.value = s.kind === 'sum' ? sumifs(s.col, plat) : averageifs(s.col, plat);
        applyNum(cell);
      });
      labelRow[s.label] = row;
    }
    row++;
  }

  row++;
  sheet.getCell(row, 2).value = '— CONVERSION RATES (computed from totals) —';
  for (let c = 2; c <= 5; c++) {
    const cell = sheet.getCell(row, c);
    cell.fill = STYLE.sectionFill;
    cell.font = STYLE.sectionFont;
    cell.border = STYLE.thinBorder;
  }
  row++;

  const txnRow = labelRow['Total Transactions'];
  const sessRow = labelRow['Total Sessions'];
  const landRow = labelRow['Total Web Landing Page'];
  const loginClickRow = labelRow['Login/Signup Click'];
  const loginOkRow = labelRow['Login/Signup Success'];
  const esimClickRow = labelRow['eSIM Install Click'];
  const esimOkRow = labelRow['eSIM Install Success'];

  sheet.getCell(row, 2).value = 'Overall Conversion Rate';
  applyLabel(sheet.getCell(row, 2), true);
  ['C', 'D', 'E'].forEach((col, i) => {
    const plat = ['Android', 'iOS', 'Web'][i];
    const denomRow = plat === 'Web' ? landRow : sessRow;
    const cell = sheet.getCell(row, 3 + i);
    cell.value = { formula: `IFERROR(${col}${txnRow}/${col}${denomRow},0)` };
    applyNum(cell, STYLE.pctFmt, true);
  });
  row++;

  sheet.getCell(row, 2).value = 'Login Success Rate (app only)';
  applyLabel(sheet.getCell(row, 2));
  ['C', 'D'].forEach((col, i) => {
    const cell = sheet.getCell(row, 3 + i);
    cell.value = { formula: `IFERROR(${col}${loginOkRow}/${col}${loginClickRow},0)` };
    applyNum(cell, STYLE.pctFmt);
  });
  const naCell1 = sheet.getCell(row, 5);
  naCell1.value = 'n/a';
  naCell1.font = STYLE.mutedFont;
  naCell1.alignment = { horizontal: 'right' };
  naCell1.border = STYLE.thinBorder;
  row++;

  sheet.getCell(row, 2).value = 'eSIM Install Success Rate';
  applyLabel(sheet.getCell(row, 2));
  ['C', 'D'].forEach((col, i) => {
    const cell = sheet.getCell(row, 3 + i);
    cell.value = { formula: `IFERROR(${col}${esimOkRow}/${col}${esimClickRow},0)` };
    applyNum(cell, STYLE.pctFmt);
  });
  const naCell2 = sheet.getCell(row, 5);
  naCell2.value = 'n/a';
  naCell2.font = STYLE.mutedFont;
  naCell2.alignment = { horizontal: 'right' };
  naCell2.border = STYLE.thinBorder;

  sheet.getColumn(1).width = 2;
  sheet.getColumn(2).width = 44;
  [3, 4, 5].forEach(c => { sheet.getColumn(c).width = 16; });
}

// ----- Volume Comparison ------------------------------------------------------

function buildVolumeComparisonSheet(workbook) {
  const sheet = workbook.addWorksheet('Volume Comparison', { views: [{ showGridLines: false }] });

  sheet.getCell('B2').value = 'Volume Comparison: Android vs iOS vs Web';
  sheet.getCell('B2').font = STYLE.titleFont;
  sheet.mergeCells('B2:F2');
  sheet.getCell('B3').value = 'Totals and daily averages over the 91-day period. All values are live formulas.';
  sheet.getCell('B3').font = STYLE.subtitleFont;
  sheet.mergeCells('B3:F3');

  ['Metric', 'Android', 'iOS', 'Web'].forEach((h, i) => { sheet.getCell(5, 2 + i).value = h; });
  styleHeaderRow(sheet, 5, 2, 5);

  const rows = [
    ['Total Top-of-Funnel Visits',
      `SUMIFS('Daily Data'!C:C,'Daily Data'!B:B,"Android")`,
      `SUMIFS('Daily Data'!C:C,'Daily Data'!B:B,"iOS")`,
      `SUMIFS('Daily Data'!J:J,'Daily Data'!B:B,"Web")`],
    ['Total Transactions',
      `SUMIFS('Daily Data'!M:M,'Daily Data'!B:B,"Android")`,
      `SUMIFS('Daily Data'!M:M,'Daily Data'!B:B,"iOS")`,
      `SUMIFS('Daily Data'!M:M,'Daily Data'!B:B,"Web")`],
    ['Avg Daily Top-of-Funnel',
      `IFERROR(AVERAGEIFS('Daily Data'!C:C,'Daily Data'!B:B,"Android"),0)`,
      `IFERROR(AVERAGEIFS('Daily Data'!C:C,'Daily Data'!B:B,"iOS"),0)`,
      `IFERROR(AVERAGEIFS('Daily Data'!J:J,'Daily Data'!B:B,"Web"),0)`],
    ['Avg Daily Transactions',
      `IFERROR(AVERAGEIFS('Daily Data'!M:M,'Daily Data'!B:B,"Android"),0)`,
      `IFERROR(AVERAGEIFS('Daily Data'!M:M,'Daily Data'!B:B,"iOS"),0)`,
      `IFERROR(AVERAGEIFS('Daily Data'!M:M,'Daily Data'!B:B,"Web"),0)`],
  ];

  rows.forEach((r, i) => {
    const rowNum = 6 + i;
    sheet.getCell(rowNum, 2).value = r[0];
    applyLabel(sheet.getCell(rowNum, 2));
    for (let j = 0; j < 3; j++) {
      const cell = sheet.getCell(rowNum, 3 + j);
      cell.value = { formula: r[1 + j] };
      applyNum(cell);
    }
  });

  sheet.getColumn(1).width = 2;
  sheet.getColumn(2).width = 36;
  [3, 4, 5].forEach(c => { sheet.getColumn(c).width = 18; });
}

// ----- Conversion Comparison --------------------------------------------------

function buildConversionComparisonSheet(workbook) {
  const sheet = workbook.addWorksheet('Conversion Comparison', { views: [{ showGridLines: false }] });

  sheet.getCell('B2').value = 'Conversion Rate Comparison';
  sheet.getCell('B2').font = STYLE.titleFont;
  sheet.mergeCells('B2:F2');
  sheet.getCell('B3').value = 'Transactions ÷ top-of-funnel. Login & eSIM success rates are app-only.';
  sheet.getCell('B3').font = STYLE.subtitleFont;
  sheet.mergeCells('B3:F3');

  ['Metric', 'Android', 'iOS', 'Web'].forEach((h, i) => { sheet.getCell(5, 2 + i).value = h; });
  styleHeaderRow(sheet, 5, 2, 5);

  const rows = [
    ['Overall Conversion Rate',
      `IFERROR(SUMIFS('Daily Data'!M:M,'Daily Data'!B:B,"Android")/SUMIFS('Daily Data'!C:C,'Daily Data'!B:B,"Android"),0)`,
      `IFERROR(SUMIFS('Daily Data'!M:M,'Daily Data'!B:B,"iOS")/SUMIFS('Daily Data'!C:C,'Daily Data'!B:B,"iOS"),0)`,
      `IFERROR(SUMIFS('Daily Data'!M:M,'Daily Data'!B:B,"Web")/SUMIFS('Daily Data'!J:J,'Daily Data'!B:B,"Web"),0)`],
    ['Login Success Rate',
      `IFERROR(SUMIFS('Daily Data'!F:F,'Daily Data'!B:B,"Android")/SUMIFS('Daily Data'!E:E,'Daily Data'!B:B,"Android"),0)`,
      `IFERROR(SUMIFS('Daily Data'!F:F,'Daily Data'!B:B,"iOS")/SUMIFS('Daily Data'!E:E,'Daily Data'!B:B,"iOS"),0)`,
      0],
    ['eSIM Install Success Rate',
      `IFERROR(SUMIFS('Daily Data'!Q:Q,'Daily Data'!B:B,"Android")/SUMIFS('Daily Data'!P:P,'Daily Data'!B:B,"Android"),0)`,
      `IFERROR(SUMIFS('Daily Data'!Q:Q,'Daily Data'!B:B,"iOS")/SUMIFS('Daily Data'!P:P,'Daily Data'!B:B,"iOS"),0)`,
      0],
  ];

  rows.forEach((r, i) => {
    const rowNum = 6 + i;
    sheet.getCell(rowNum, 2).value = r[0];
    applyLabel(sheet.getCell(rowNum, 2));
    for (let j = 0; j < 3; j++) {
      const cell = sheet.getCell(rowNum, 3 + j);
      cell.value = (typeof r[1 + j] === 'string') ? { formula: r[1 + j] } : r[1 + j];
      applyNum(cell, STYLE.pctFmt);
    }
  });

  sheet.getColumn(1).width = 2;
  sheet.getColumn(2).width = 30;
  [3, 4, 5].forEach(c => { sheet.getColumn(c).width = 16; });
}

// ----- Daily Trend ------------------------------------------------------------

function buildDailyTrendSheet(workbook) {
  const sheet = workbook.addWorksheet('Daily Trend', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 5, showGridLines: false }],
  });

  sheet.getCell('B2').value = 'Daily Trend: Traffic & Transactions by Platform';
  sheet.getCell('B2').font = STYLE.titleFont;
  sheet.mergeCells('B2:F2');
  sheet.getCell('B3').value = 'Per-day series for each platform, built with SUMIFS from Daily Data.';
  sheet.getCell('B3').font = STYLE.subtitleFont;
  sheet.mergeCells('B3:F3');

  ['Day', 'Android Sessions', 'iOS Sessions', 'Web Landing',
   'Android Txn', 'iOS Txn', 'Web Txn'].forEach((h, i) => { sheet.getCell(5, 2 + i).value = h; });
  styleHeaderRow(sheet, 5, 2, 8);

  for (let i = 0; i < N_DAYS; i++) {
    const day = i + 1;
    const r = 6 + i;
    sheet.getCell(r, 2).value = day;
    sheet.getCell(r, 2).font = STYLE.bodyFont;
    sheet.getCell(r, 2).alignment = { horizontal: 'center' };
    sheet.getCell(r, 2).border = STYLE.thinBorder;

    const formulas = [
      `SUMIFS('Daily Data'!C:C,'Daily Data'!A:A,${day},'Daily Data'!B:B,"Android")`,
      `SUMIFS('Daily Data'!C:C,'Daily Data'!A:A,${day},'Daily Data'!B:B,"iOS")`,
      `SUMIFS('Daily Data'!J:J,'Daily Data'!A:A,${day},'Daily Data'!B:B,"Web")`,
      `SUMIFS('Daily Data'!M:M,'Daily Data'!A:A,${day},'Daily Data'!B:B,"Android")`,
      `SUMIFS('Daily Data'!M:M,'Daily Data'!A:A,${day},'Daily Data'!B:B,"iOS")`,
      `SUMIFS('Daily Data'!M:M,'Daily Data'!A:A,${day},'Daily Data'!B:B,"Web")`,
    ];
    formulas.forEach((f, j) => {
      const cell = sheet.getCell(r, 3 + j);
      cell.value = { formula: f };
      applyNum(cell);
    });
  }

  sheet.getColumn(1).width = 2;
  sheet.getColumn(2).width = 8;
  [3, 4, 5, 6, 7, 8].forEach(c => { sheet.getColumn(c).width = 16; });
}

// ----- App Funnel -------------------------------------------------------------

function buildAppFunnelSheet(workbook) {
  const sheet = workbook.addWorksheet('App Funnel', { views: [{ showGridLines: false }] });

  sheet.getCell('B2').value = 'In-App Purchase Funnel: Android vs iOS';
  sheet.getCell('B2').font = STYLE.titleFont;
  sheet.mergeCells('B2:F2');
  sheet.getCell('B3').value = 'Volume at each stage and each stage as % of Sessions.';
  sheet.getCell('B3').font = STYLE.subtitleFont;
  sheet.mergeCells('B3:F3');

  ['Stage', 'Android', 'iOS', 'Android % of Sessions', 'iOS % of Sessions']
    .forEach((h, i) => { sheet.getCell(5, 2 + i).value = h; });
  styleHeaderRow(sheet, 5, 2, 6);

  const stages = [
    ['Sessions', 'C'], ['Explore Plans Click', 'D'],
    ['Login/Signup Click', 'E'], ['Login/Signup Success', 'F'],
    ['View Countries', 'G'], ['Selected Plan', 'H'],
    ['Continue to Checkout', 'I'], ['Transactions', 'M'],
  ];

  stages.forEach(([label, col], i) => {
    const r = 6 + i;
    sheet.getCell(r, 2).value = label;
    applyLabel(sheet.getCell(r, 2));
    sheet.getCell(r, 3).value = { formula: `SUMIFS('Daily Data'!${col}:${col},'Daily Data'!B:B,"Android")` };
    sheet.getCell(r, 4).value = { formula: `SUMIFS('Daily Data'!${col}:${col},'Daily Data'!B:B,"iOS")` };
    applyNum(sheet.getCell(r, 3));
    applyNum(sheet.getCell(r, 4));
    sheet.getCell(r, 5).value = { formula: `IFERROR(C${r}/$C$6,0)` };
    sheet.getCell(r, 6).value = { formula: `IFERROR(D${r}/$D$6,0)` };
    applyNum(sheet.getCell(r, 5), STYLE.pctFmt1);
    applyNum(sheet.getCell(r, 6), STYLE.pctFmt1);
  });

  sheet.getColumn(1).width = 2;
  sheet.getColumn(2).width = 24;
  [3, 4, 5, 6].forEach(c => { sheet.getColumn(c).width = 18; });
}

// ----- Web Funnel -------------------------------------------------------------

function buildWebFunnelSheet(workbook) {
  const sheet = workbook.addWorksheet('Web Funnel', { views: [{ showGridLines: false }] });

  sheet.getCell('B2').value = 'Web Purchase Funnel';
  sheet.getCell('B2').font = STYLE.titleFont;
  sheet.mergeCells('B2:F2');
  sheet.getCell('B3').value = 'Web flow: Landing → Pricing → Checkout → Transactions.';
  sheet.getCell('B3').font = STYLE.subtitleFont;
  sheet.mergeCells('B3:F3');

  ['Stage', 'Volume', '% of Landing', 'Step Conversion']
    .forEach((h, i) => { sheet.getCell(5, 2 + i).value = h; });
  styleHeaderRow(sheet, 5, 2, 5);

  const webStages = [['Landing Page', 'J'], ['Pricing', 'K'], ['Checkout', 'L'], ['Transactions', 'M']];

  webStages.forEach(([label, col], i) => {
    const r = 6 + i;
    sheet.getCell(r, 2).value = label;
    applyLabel(sheet.getCell(r, 2));
    sheet.getCell(r, 3).value = { formula: `SUMIFS('Daily Data'!${col}:${col},'Daily Data'!B:B,"Web")` };
    applyNum(sheet.getCell(r, 3));
    sheet.getCell(r, 4).value = { formula: `IFERROR(C${r}/$C$6,0)` };
    applyNum(sheet.getCell(r, 4), STYLE.pctFmt1);

    if (i === 0) {
      const c = sheet.getCell(r, 5);
      c.value = '—';
      c.font = STYLE.mutedFont;
      c.alignment = { horizontal: 'right' };
      c.border = STYLE.thinBorder;
    } else {
      sheet.getCell(r, 5).value = { formula: `IFERROR(C${r}/C${r - 1},0)` };
      applyNum(sheet.getCell(r, 5), STYLE.pctFmt1);
    }
  });

  sheet.getColumn(1).width = 2;
  sheet.getColumn(2).width = 20;
  [3, 4, 5].forEach(c => { sheet.getColumn(c).width = 16; });
}

// ----- Notes ------------------------------------------------------------------

function buildNotesSheet(workbook) {
  const sheet = workbook.addWorksheet('Notes', { views: [{ showGridLines: false }] });

  const notes = [
    ['How to use this workbook', ''],
    ['', ''],
    ['Platform Comparison', 'Side-by-side summary with live formulas pointing at Daily Data.'],
    ['Visual sheets', 'Volume / Conversion / Daily Trend / App Funnel / Web Funnel each have tables AND embedded charts.'],
    ['Daily Data', `Long-format table — one row per Day × Platform. ${N_DAYS} days, wrapped in Excel Table "DailyData".`],
    ['', ''],
    ['Build your own pivot', ''],
    ['', '1. Open the Daily Data sheet.'],
    ['', '2. Insert → PivotTable on the DailyData table.'],
    ['', '3. Rows: Platform.  Values: any metric column.'],
    ['', ''],
    ['Data shape notes', ''],
    ['', 'Android/iOS have an 8-stage in-app purchase funnel.'],
    ['', 'Web has a 3-stage funnel: Landing → Pricing → Checkout → Transactions.'],
    ['', 'Android/iOS have eSIM setup data; Web does not.'],
    ['', 'Web has New vs Returning user counts; Android/iOS do not.'],
  ];

  notes.forEach((n, i) => {
    const row = i + 1;
    sheet.getCell(row, 1).value = n[0];
    sheet.getCell(row, 2).value = n[1];
    sheet.getCell(row, 1).font = row === 1 ? { name: 'Arial', bold: true, size: 14 } : { name: 'Arial', bold: true, size: 11 };
    sheet.getCell(row, 2).font = STYLE.bodyFont;
    sheet.getCell(row, 2).alignment = { wrapText: true, vertical: 'top' };
  });

  sheet.getColumn(1).width = 28;
  sheet.getColumn(2).width = 90;
}

// ----- Copy originals ---------------------------------------------------------

function copySourceSheets(XLSX, arrayBuffer, workbook) {
  const srcWb = XLSX.read(arrayBuffer, { type: 'array' });

  for (const cfg of Object.values(SOURCE_SHEETS)) {
    const ws = srcWb.Sheets[cfg.name];
    if (!ws) continue;

    const out = workbook.addWorksheet(cfg.name);
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[addr];
        if (!cell) continue;
        const outCell = out.getCell(R + 1, C + 1);
        outCell.value = cell.v;
        if (typeof cell.v === 'number') {
          outCell.numFmt = cell.z || '#,##0';
          outCell.alignment = { horizontal: 'right' };
        }
      }
    }

    if (ws['!merges']) {
      ws['!merges'].forEach(m => out.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1));
    }

    for (let C = range.s.c; C <= range.e.c; C++) {
      out.getColumn(C + 1).width = 14;
    }

    for (let R = 0; R < Math.min(2, range.e.r + 1); R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const cell = out.getCell(R + 1, C + 1);
        if (cell.value != null) {
          cell.font = { name: 'Arial', bold: true, size: 11 };
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        }
      }
    }
  }
}

// =============================================================================
// Chart injection (post-process OOXML)
//
// ExcelJS does not write charts. We build chart XML manually, then open the
// ExcelJS output as a zip, inject the chart/drawing parts, and repack.
// =============================================================================

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

/**
 * Build chart1.xml content.
 *
 * @param {object} spec
 * @param {string} spec.title
 * @param {'bar'|'col'|'line'} spec.type
 * @param {string} spec.catRef     full sheet-qualified reference, e.g. "'Volume Comparison'!$C$5:$E$5"
 * @param {Array<{name:string, valRef:string}>} spec.series  each series name is either a sheet ref or a literal string
 * @param {boolean} [spec.dataLabels]
 * @param {string}  [spec.dataLabelFmt]  e.g. '0%'
 */
function buildChartXml(spec) {
  const isBar = spec.type === 'bar' || spec.type === 'col';
  const chartTag = isBar ? 'barChart' : 'lineChart';
  const barDir = spec.type === 'bar' ? 'bar' : 'col';

  const seriesXml = spec.series.map((s, i) => {
    const txTag = s.nameRef
      ? `<tx><strRef><f>${xmlEscape(s.nameRef)}</f></strRef></tx>`
      : `<tx><v>${xmlEscape(s.name)}</v></tx>`;
    const catXml = spec.catRef
      ? `<cat><strRef><f>${xmlEscape(spec.catRef)}</f></strRef></cat>`
      : '';
    if (isBar) {
      return `<ser>
          <idx val="${i}"/><order val="${i}"/>
          ${txTag}
          <spPr><a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:prstDash val="solid"/></a:ln></spPr>
          ${catXml}
          <val><numRef><f>${xmlEscape(s.valRef)}</f></numRef></val>
        </ser>`;
    } else {
      return `<ser>
          <idx val="${i}"/><order val="${i}"/>
          ${txTag}
          <spPr><a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="19050"><a:prstDash val="solid"/></a:ln></spPr>
          ${catXml}
          <val><numRef><f>${xmlEscape(s.valRef)}</f></numRef></val>
          <smooth val="0"/>
        </ser>`;
    }
  }).join('');

  const dLbls = spec.dataLabels
    ? `<dLbls>${spec.dataLabelFmt ? `<numFmt formatCode="${xmlEscape(spec.dataLabelFmt)}" sourceLinked="0"/>` : ''}<showLegendKey val="0"/><showVal val="1"/><showCatName val="0"/><showSerName val="0"/><showPercent val="0"/><showBubbleSize val="0"/></dLbls>`
    : '';

  const plotInner = isBar
    ? `<${chartTag}>
         <barDir val="${barDir}"/>
         <grouping val="clustered"/>
         <varyColors val="0"/>
         ${seriesXml}
         ${dLbls}
         <gapWidth val="150"/>
         <axId val="10"/><axId val="100"/>
       </${chartTag}>`
    : `<${chartTag}>
         <grouping val="standard"/>
         <varyColors val="0"/>
         ${seriesXml}
         ${dLbls}
         <marker val="0"/>
         <axId val="10"/><axId val="100"/>
       </${chartTag}>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich>
        <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" wrap="square" anchor="ctr" anchorCtr="1"/>
        <a:lstStyle/>
        <a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>${xmlEscape(spec.title)}</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      ${plotInner.replace(/<ser>/g, '<c:ser>').replace(/<\/ser>/g, '</c:ser>')
                 .replace(/<idx /g, '<c:idx ').replace(/<\/idx>/g, '</c:idx>')
                 .replace(/<order /g, '<c:order ').replace(/<\/order>/g, '</c:order>')
                 .replace(/<tx>/g, '<c:tx>').replace(/<\/tx>/g, '</c:tx>')
                 .replace(/<strRef>/g, '<c:strRef>').replace(/<\/strRef>/g, '</c:strRef>')
                 .replace(/<numRef>/g, '<c:numRef>').replace(/<\/numRef>/g, '</c:numRef>')
                 .replace(/<f>/g, '<c:f>').replace(/<\/f>/g, '</c:f>')
                 .replace(/<v>/g, '<c:v>').replace(/<\/v>/g, '</c:v>')
                 .replace(/<cat>/g, '<c:cat>').replace(/<\/cat>/g, '</c:cat>')
                 .replace(/<val>/g, '<c:val>').replace(/<\/val>/g, '</c:val>')
                 .replace(/<spPr>/g, '<c:spPr>').replace(/<\/spPr>/g, '</c:spPr>')
                 .replace(/<dLbls>/g, '<c:dLbls>').replace(/<\/dLbls>/g, '</c:dLbls>')
                 .replace(/<numFmt /g, '<c:numFmt ').replace(/<showLegendKey /g, '<c:showLegendKey ')
                 .replace(/<showVal /g, '<c:showVal ').replace(/<showCatName /g, '<c:showCatName ')
                 .replace(/<showSerName /g, '<c:showSerName ').replace(/<showPercent /g, '<c:showPercent ')
                 .replace(/<showBubbleSize /g, '<c:showBubbleSize ')
                 .replace(/<gapWidth /g, '<c:gapWidth ').replace(/<axId /g, '<c:axId ')
                 .replace(/<barDir /g, '<c:barDir ').replace(/<grouping /g, '<c:grouping ')
                 .replace(/<varyColors /g, '<c:varyColors ').replace(/<smooth /g, '<c:smooth ')
                 .replace(/<marker /g, '<c:marker ')
                 .replace(/<(\/?)barChart>/g, '<$1c:barChart>')
                 .replace(/<(\/?)lineChart>/g, '<$1c:lineChart>')}
      <c:catAx>
        <c:axId val="10"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${spec.type === 'bar' ? 'l' : 'b'}"/>
        <c:crossAx val="100"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="100"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${spec.type === 'bar' ? 'b' : 'l'}"/>
        <c:majorGridlines/>
        <c:crossAx val="10"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`;
}

/**
 * Build drawing XML anchoring one or more charts into a sheet at given cells.
 *
 * @param {Array<{cell:string, chartRelId:string, width?:number, height?:number}>} anchors
 *    cell is a 1-indexed anchor like {col, row, colOffset?, rowOffset?}
 */
function buildDrawingXml(anchors) {
  const anchorXml = anchors.map((a, i) => {
    const cx = a.width  || 6000000;   // EMU, ~6.3 inches
    const cy = a.height || 3500000;   // EMU, ~3.6 inches
    return `
      <xdr:oneCellAnchor>
        <xdr:from>
          <xdr:col>${a.col}</xdr:col><xdr:colOff>0</xdr:colOff>
          <xdr:row>${a.row}</xdr:row><xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:ext cx="${cx}" cy="${cy}"/>
        <xdr:graphicFrame macro="">
          <xdr:nvGraphicFramePr>
            <xdr:cNvPr id="${i + 2}" name="Chart ${i + 1}"/>
            <xdr:cNvGraphicFramePr/>
          </xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       r:id="${a.chartRelId}"/>
            </a:graphicData>
          </a:graphic>
        </xdr:graphicFrame>
        <xdr:clientData/>
      </xdr:oneCellAnchor>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  ${anchorXml}
</xdr:wsDr>`;
}

/**
 * Inject charts into a .xlsx Blob produced by ExcelJS.
 *
 * @param {JSZip} zip           loaded zip of the ExcelJS output
 * @param {Array} chartSpecs    [{sheetName, anchor:{col,row}, chart:{...}}]
 *   where chart is a spec compatible with buildChartXml().
 */
async function injectCharts(JSZip, arrayBuffer, chartSpecs) {
  const zip = await JSZip.loadAsync(arrayBuffer);

  // Group specs by sheetName -> list of (anchor, chart)
  const bySheet = {};
  for (const spec of chartSpecs) {
    if (!bySheet[spec.sheetName]) bySheet[spec.sheetName] = [];
    bySheet[spec.sheetName].push(spec);
  }

  // Map sheet name -> xl/worksheets/sheetN.xml path via workbook.xml + rels
  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  const workbookRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');

  // Parse sheet names and rIds from workbook.xml
  // Attribute order varies between writers, so extract name and r:id independently from each <sheet ... /> tag
  const sheetNameToRid = {};
  const sheetTagRegex = /<sheet\s[^>]*\/>/g;
  let m;
  while ((m = sheetTagRegex.exec(workbookXml)) !== null) {
    const tag = m[0];
    const nameMatch = tag.match(/\sname="([^"]+)"/);
    const ridMatch = tag.match(/\sr:id="([^"]+)"/);
    if (nameMatch && ridMatch) {
      sheetNameToRid[nameMatch[1]] = ridMatch[1];
    }
  }

  // Parse rId -> target from workbook rels
  const relRegex = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>/g;
  const ridToTarget = {};
  while ((m = relRegex.exec(workbookRels)) !== null) {
    ridToTarget[m[1]] = m[2];
  }

  // Find highest existing chart number to avoid collisions
  let nextChartNum = 1;
  let nextDrawingNum = 1;
  zip.forEach((path) => {
    let match = path.match(/^xl\/charts\/chart(\d+)\.xml$/);
    if (match) nextChartNum = Math.max(nextChartNum, parseInt(match[1], 10) + 1);
    match = path.match(/^xl\/drawings\/drawing(\d+)\.xml$/);
    if (match) nextDrawingNum = Math.max(nextDrawingNum, parseInt(match[1], 10) + 1);
  });

  // Track new parts to register in [Content_Types].xml
  const newContentTypeEntries = [];

  for (const sheetName of Object.keys(bySheet)) {
    const specs = bySheet[sheetName];
    const rid = sheetNameToRid[sheetName];
    if (!rid) {
      console.warn(`Sheet "${sheetName}" not found, skipping chart injection.`);
      continue;
    }
    const target = ridToTarget[rid]; // e.g. "worksheets/sheet3.xml"
    const sheetXmlPath = 'xl/' + target;
    const sheetBase = target.split('/').pop(); // sheet3.xml
    const sheetRelsPath = `xl/worksheets/_rels/${sheetBase}.rels`;

    // Create one drawing that groups all charts for this sheet
    const drawingNum = nextDrawingNum++;
    const drawingPath = `xl/drawings/drawing${drawingNum}.xml`;
    const drawingRelsPath = `xl/drawings/_rels/drawing${drawingNum}.xml.rels`;

    // Create chart parts
    const anchors = [];
    const drawingRelEntries = [];
    specs.forEach((spec, idx) => {
      const chartNum = nextChartNum++;
      const chartPath = `xl/charts/chart${chartNum}.xml`;
      zip.file(chartPath, buildChartXml(spec.chart));
      newContentTypeEntries.push({
        partName: `/${chartPath}`,
        contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      });

      const relId = `rId${idx + 1}`;
      anchors.push({
        col: spec.anchor.col, row: spec.anchor.row,
        chartRelId: relId,
        width: spec.anchor.width, height: spec.anchor.height,
      });
      drawingRelEntries.push(
        `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNum}.xml"/>`
      );
    });

    // Write drawing + its rels
    zip.file(drawingPath, buildDrawingXml(anchors));
    zip.file(drawingRelsPath,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${drawingRelEntries.join('\n')}
</Relationships>`);
    newContentTypeEntries.push({
      partName: `/${drawingPath}`,
      contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
    });

    // Link drawing into the sheet's rels
    let sheetRelsXml;
    if (zip.file(sheetRelsPath)) {
      sheetRelsXml = await zip.file(sheetRelsPath).async('string');
    } else {
      sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    }
    // Find a free rId in this sheet's rels
    const existingIds = [...sheetRelsXml.matchAll(/Id="(rId\d+)"/g)].map(x => parseInt(x[1].slice(3), 10));
    const drawingRid = 'rId' + ((existingIds.length ? Math.max(...existingIds) : 0) + 1);
    const newRel = `<Relationship Id="${drawingRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingNum}.xml"/>`;
    sheetRelsXml = sheetRelsXml.replace('</Relationships>', newRel + '</Relationships>');
    zip.file(sheetRelsPath, sheetRelsXml);

    // Add <drawing r:id="..."/> to the sheet XML
    let sheetXml = await zip.file(sheetXmlPath).async('string');
    if (!sheetXml.includes('<drawing ')) {
      // Must be placed after specific elements per OOXML spec. Safe spot:
      // right before </worksheet>. If ExcelJS already has an element that
      // must come after drawing (legacyDrawing, oleObjects, picture), we
      // insert before it; otherwise before </worksheet>.
      const insertBefore = ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '</worksheet>'];
      let insertPos = -1;
      for (const tag of insertBefore) {
        insertPos = sheetXml.indexOf(tag);
        if (insertPos !== -1) break;
      }
      if (insertPos === -1) {
        console.warn(`Could not find insertion point in ${sheetXmlPath}`);
        continue;
      }
      sheetXml = sheetXml.slice(0, insertPos)
              + `<drawing r:id="${drawingRid}"/>`
              + sheetXml.slice(insertPos);
      zip.file(sheetXmlPath, sheetXml);
    }
  }

  // Update [Content_Types].xml with new parts
  let ctXml = await zip.file('[Content_Types].xml').async('string');
  for (const entry of newContentTypeEntries) {
    if (!ctXml.includes(`PartName="${entry.partName}"`)) {
      const override = `<Override PartName="${entry.partName}" ContentType="${entry.contentType}"/>`;
      ctXml = ctXml.replace('</Types>', override + '</Types>');
    }
  }
  zip.file('[Content_Types].xml', ctXml);

  return zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

// ----- Chart specs for each visual sheet --------------------------------------

function chartSpecs() {
  return [
    // Volume Comparison
    {
      sheetName: 'Volume Comparison',
      anchor: { col: 1, row: 11, width: 6500000, height: 3600000 },
      chart: {
        title: 'Top-of-Funnel Visits vs Transactions (Totals)',
        type: 'col',
        catRef: `'Volume Comparison'!$C$5:$E$5`,
        series: [
          { nameRef: `'Volume Comparison'!$B$6`, valRef: `'Volume Comparison'!$C$6:$E$6` },
          { nameRef: `'Volume Comparison'!$B$7`, valRef: `'Volume Comparison'!$C$7:$E$7` },
        ],
        dataLabels: true,
      },
    },
    {
      sheetName: 'Volume Comparison',
      anchor: { col: 1, row: 33, width: 6500000, height: 3600000 },
      chart: {
        title: 'Average Daily Volume',
        type: 'col',
        catRef: `'Volume Comparison'!$C$5:$E$5`,
        series: [
          { nameRef: `'Volume Comparison'!$B$8`, valRef: `'Volume Comparison'!$C$8:$E$8` },
          { nameRef: `'Volume Comparison'!$B$9`, valRef: `'Volume Comparison'!$C$9:$E$9` },
        ],
        dataLabels: true,
      },
    },
    // Conversion Comparison
    {
      sheetName: 'Conversion Comparison',
      anchor: { col: 1, row: 10, width: 6500000, height: 3600000 },
      chart: {
        title: 'Conversion Rate by Platform',
        type: 'col',
        catRef: `'Conversion Comparison'!$C$5:$E$5`,
        series: [
          { nameRef: `'Conversion Comparison'!$B$6`, valRef: `'Conversion Comparison'!$C$6:$E$6` },
          { nameRef: `'Conversion Comparison'!$B$7`, valRef: `'Conversion Comparison'!$C$7:$E$7` },
          { nameRef: `'Conversion Comparison'!$B$8`, valRef: `'Conversion Comparison'!$C$8:$E$8` },
        ],
        dataLabels: true,
        dataLabelFmt: '0.0%',
      },
    },
    // Daily Trend — traffic (col 9 = column J, row 4 = row 5)
    {
      sheetName: 'Daily Trend',
      anchor: { col: 9, row: 4, width: 8000000, height: 3600000 },
      chart: {
        title: 'Daily Top-of-Funnel Traffic',
        type: 'line',
        catRef: `'Daily Trend'!$B$6:$B$${5 + N_DAYS}`,
        series: [
          { nameRef: `'Daily Trend'!$C$5`, valRef: `'Daily Trend'!$C$6:$C$${5 + N_DAYS}` },
          { nameRef: `'Daily Trend'!$D$5`, valRef: `'Daily Trend'!$D$6:$D$${5 + N_DAYS}` },
          { nameRef: `'Daily Trend'!$E$5`, valRef: `'Daily Trend'!$E$6:$E$${5 + N_DAYS}` },
        ],
      },
    },
    {
      sheetName: 'Daily Trend',
      anchor: { col: 9, row: 26, width: 8000000, height: 3600000 },
      chart: {
        title: 'Daily Transactions',
        type: 'line',
        catRef: `'Daily Trend'!$B$6:$B$${5 + N_DAYS}`,
        series: [
          { nameRef: `'Daily Trend'!$F$5`, valRef: `'Daily Trend'!$F$6:$F$${5 + N_DAYS}` },
          { nameRef: `'Daily Trend'!$G$5`, valRef: `'Daily Trend'!$G$6:$G$${5 + N_DAYS}` },
          { nameRef: `'Daily Trend'!$H$5`, valRef: `'Daily Trend'!$H$6:$H$${5 + N_DAYS}` },
        ],
      },
    },
    // App Funnel
    {
      sheetName: 'App Funnel',
      anchor: { col: 1, row: 15, width: 6500000, height: 4500000 },
      chart: {
        title: 'App Funnel Volume by Stage',
        type: 'bar',
        catRef: `'App Funnel'!$B$6:$B$13`,
        series: [
          { nameRef: `'App Funnel'!$C$5`, valRef: `'App Funnel'!$C$6:$C$13` },
          { nameRef: `'App Funnel'!$D$5`, valRef: `'App Funnel'!$D$6:$D$13` },
        ],
      },
    },
    {
      sheetName: 'App Funnel',
      anchor: { col: 1, row: 44, width: 6500000, height: 4500000 },
      chart: {
        title: 'App Funnel: % of Sessions Reaching Each Stage',
        type: 'bar',
        catRef: `'App Funnel'!$B$6:$B$13`,
        series: [
          { nameRef: `'App Funnel'!$E$5`, valRef: `'App Funnel'!$E$6:$E$13` },
          { nameRef: `'App Funnel'!$F$5`, valRef: `'App Funnel'!$F$6:$F$13` },
        ],
        dataLabels: true,
        dataLabelFmt: '0%',
      },
    },
    // Web Funnel
    {
      sheetName: 'Web Funnel',
      anchor: { col: 1, row: 12, width: 6500000, height: 3600000 },
      chart: {
        title: 'Web Funnel: Volume by Stage',
        type: 'bar',
        catRef: `'Web Funnel'!$B$6:$B$9`,
        series: [
          { nameRef: `'Web Funnel'!$C$5`, valRef: `'Web Funnel'!$C$6:$C$9` },
        ],
        dataLabels: true,
      },
    },
    {
      sheetName: 'Web Funnel',
      anchor: { col: 1, row: 33, width: 6500000, height: 3600000 },
      chart: {
        title: 'Web Funnel: % of Landing Reaching Each Stage',
        type: 'bar',
        catRef: `'Web Funnel'!$B$6:$B$9`,
        series: [
          { nameRef: `'Web Funnel'!$D$5`, valRef: `'Web Funnel'!$D$6:$D$9` },
        ],
        dataLabels: true,
        dataLabelFmt: '0%',
      },
    },
  ];
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Transform a Platform export workbook into the polished comparison workbook.
 *
 * @param {File|Blob|ArrayBuffer} input
 * @returns {Promise<Blob>}  xlsx Blob
 */
async function transformWorkbook(input) {
  const { XLSX, ExcelJS, JSZip } = await ensureDependencies();

  let arrayBuffer;
  if (input instanceof ArrayBuffer) arrayBuffer = input;
  else if (input instanceof Blob)   arrayBuffer = await input.arrayBuffer();
  else throw new Error('transformWorkbook: input must be File, Blob, or ArrayBuffer');

  // 1. Read source
  const src = readSourceSheets(XLSX, arrayBuffer);

  // 2. Build records
  const records = buildDailyDataRecords(src);

  // 3. Assemble ExcelJS workbook
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Platform Comparison Transformer';
  workbook.created = new Date();

  copySourceSheets(XLSX, arrayBuffer, workbook);
  buildPlatformComparisonSheet(workbook);
  buildVolumeComparisonSheet(workbook);
  buildConversionComparisonSheet(workbook);
  buildDailyTrendSheet(workbook);
  buildAppFunnelSheet(workbook);
  buildWebFunnelSheet(workbook);
  buildDailyDataSheet(workbook, records);
  buildNotesSheet(workbook);

  // 4. Serialize
  const baseBuffer = await workbook.xlsx.writeBuffer();

  // 5. Inject charts
  const finalBuffer = await injectCharts(JSZip, baseBuffer, chartSpecs());

  return new Blob([finalBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Convenience: transform and trigger a browser download.
 *
 * @param {File|Blob|ArrayBuffer} input
 * @param {string} [filename='Platform_Comparison.xlsx']
 */
async function transformAndDownload(input, filename = 'Platform_Comparison.xlsx') {
  const blob = await transformWorkbook(input);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Expose to global scope so a plain inline <script> can use it.
// (We converted from an ES module to a classic script for easier integration.)
if (typeof window !== 'undefined') {
  window.PlatformComparisonTransformer = {
    transformWorkbook,
    transformAndDownload,
  };
}

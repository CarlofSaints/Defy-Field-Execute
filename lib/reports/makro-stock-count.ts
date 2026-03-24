import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import path from 'path';
import fs from 'fs';

// ─── Category mapping ───────────────────────────────────────────────────────
// Maps the category header in the Perigee raw export to subCat + cat values
// used in the output report.
const CATEGORY_MAP: Record<string, { subCat: string; cat: string }> = {
  'BUILT IN OVENS':              { subCat: 'OVEN',            cat: 'COOKING' },
  'ELECTRIC HOBS':               { subCat: 'HOB',             cat: 'COOKING' },
  'GAS HOBS':                    { subCat: 'HOB',             cat: 'COOKING' },
  'COOKER HOODS':                { subCat: 'COOKERHOOD',      cat: 'COOKING' },
  'ELECTRIC STOVES':             { subCat: 'STOVE',           cat: 'COOKING' },
  'GAS STOVES':                  { subCat: 'STOVE',           cat: 'COOKING' },
  'GAS - ELECTRIC STOVES':       { subCat: 'STOVE',           cat: 'COOKING' },
  'MICROWAVES':                  { subCat: 'MICROWAVE',       cat: 'MICROWAVE' },
  'MULTIFUNCTION MICROWAVES':    { subCat: 'MICROWAVE',       cat: 'MICROWAVE' },
  'FRONT LOADERS':               { subCat: 'WASHING MACHINE', cat: 'WASHING MACHINE' },
  'TOP LOADERS':                 { subCat: 'WASHING MACHINE', cat: 'WASHING MACHINE' },
  'TWIN TUBS':                   { subCat: 'WASHING MACHINE', cat: 'WASHING MACHINE' },
  'DRYERS':                      { subCat: 'DRYER',           cat: 'WASHING MACHINE' },
  'DISHWASHERS':                 { subCat: 'DISHWASHER',      cat: 'DISHWASHER' },
  'CHEST FREEZERS':              { subCat: 'FRIDGE',          cat: 'FRIDGE' },
  'SINGLE DOOR FRIDGE/FREEZER':  { subCat: 'FRIDGE',          cat: 'FRIDGE' },
  'DOUBLE DOORS / COMBI FRIDGES':{ subCat: 'FRIDGE',          cat: 'FRIDGE' },
  'UPRIGHT LARDER AND FREEZER':  { subCat: 'FRIDGE',          cat: 'FRIDGE' },
  'SIDE BY SIDES':               { subCat: 'FRIDGE',          cat: 'FRIDGE' },
  'BOX SETS':                    { subCat: 'BOX SET',         cat: 'BOX SET' },
};

// Codes that appear in the CRITICAL LINES section of the form
const CRITICAL_LINE_CODES = new Set([
  'DBO489E', 'DBO496E', 'DBO767', 'DBO768', 'DBO775',
  'DSS694',  'DGS906',  'DAC447', 'DAC627', 'DAC675',
  'DFF447',  'DFF436',  'DFF463', 'DAW389', 'DWD318', 'DTL160',
]);

// Brand colours
const DEFY_RED = 'E31837';
const WHITE    = 'FFFFFF';

// ─── Types ───────────────────────────────────────────────────────────────────
interface DatabaseRow {
  week:          string;
  storeName:     string;
  storeCode:     string;
  repName:       string;
  date:          string;
  subCategory:   string;
  productCode:   string;
  product:       string;
  soh:           number;
  onFloor:       string;
  reason:        string;
  province:      string;
  category:      string;
  keyCat:        string;
  criticalLines: string;
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
function getISOWeek(dateStr: string): { week: number; year: number } {
  const parts = dateStr.split('/');
  if (parts.length !== 3) return { week: 1, year: new Date().getFullYear() };
  const [d, m, y] = parts.map(Number);
  const date = new Date(y, m - 1, d);
  const thu = new Date(date);
  thu.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
  const yearStart = new Date(thu.getFullYear(), 0, 4);
  const week = 1 + Math.round(
    ((thu.getTime() - yearStart.getTime()) / 86400000 - 3 + (yearStart.getDay() + 6) % 7) / 7
  );
  return { week, year: thu.getFullYear() };
}

function extractProductCode(header: string): string {
  if (!header) return '';
  const words = header.trim().split(/\s+/);
  // Critical lines headers are "DEFY {CODE} SOH" — skip the brand prefix
  if (words[0].toUpperCase() === 'DEFY' && words.length > 1) return words[1].toUpperCase();
  return words[0].toUpperCase();
}

function applyHeaderStyle(cell: ExcelJS.Cell) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DEFY_RED } };
  cell.font = { bold: true, color: { argb: WHITE }, size: 10, name: 'Arial' };
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  cell.border = {
    top:    { style: 'thin', color: { argb: WHITE } },
    bottom: { style: 'thin', color: { argb: WHITE } },
    left:   { style: 'thin', color: { argb: WHITE } },
    right:  { style: 'thin', color: { argb: WHITE } },
  };
}

function applyDataStyle(cell: ExcelJS.Cell, isEven = false) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'F7F7F7' : WHITE } };
  cell.font = { size: 10, name: 'Arial' };
  cell.alignment = { vertical: 'middle', wrapText: false };
  cell.border = {
    top:    { style: 'thin', color: { argb: 'E0E0E0' } },
    bottom: { style: 'thin', color: { argb: 'E0E0E0' } },
    left:   { style: 'thin', color: { argb: 'E0E0E0' } },
    right:  { style: 'thin', color: { argb: 'E0E0E0' } },
  };
}

// ─── Main export ─────────────────────────────────────────────────────────────
export async function generateMakroStockCount(
  fileBuffer: Buffer,
  brand: string,
): Promise<{ buffer: Buffer; filename: string }> {

  // ── 1. Parse raw Perigee Excel ─────────────────────────────────────────────
  const inputWb = XLSX.read(fileBuffer, { type: 'buffer' });
  const inputWs = inputWb.Sheets[inputWb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(inputWs, { header: 1, defval: null }) as (string | number | null)[][];

  if (rawData.length < 2) throw new Error('No data rows found in uploaded file.');

  const headers = rawData[0] as (string | null)[];
  const dataRows = rawData.slice(1);

  // ── 2. Build category column map ──────────────────────────────────────────
  // Find where each category section starts; stop at CRITICAL LINES
  const categoryAtCol = new Map<number, { subCat: string; cat: string }>();
  let criticalLinesStartCol = Infinity;

  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (!h) continue;
    const hTrimmed = h.trim();
    if (hTrimmed === 'CRITICAL LINES') {
      criticalLinesStartCol = i;
      break;
    }
    if (CATEGORY_MAP[hTrimmed]) {
      categoryAtCol.set(i, CATEGORY_MAP[hTrimmed]);
    }
  }

  // ── 3. Identify all SOH columns (before CRITICAL LINES section) ───────────
  interface SohCol {
    colIdx:       number;
    header:       string;
    productCode:  string;
    subCat:       string;
    cat:          string;
    onFloorIdx:   number | null;
    reasonIdx:    number | null;
  }

  const sohCols: SohCol[] = [];

  for (let i = 0; i < headers.length && i < criticalLinesStartCol; i++) {
    const h = headers[i];
    if (!h || !h.toUpperCase().includes('SOH')) continue;

    const productCode = extractProductCode(h);

    // Determine category by finding the nearest category header to the left
    let subCat = 'UNKNOWN';
    let cat    = 'UNKNOWN';
    let bestCol = -1;
    for (const [catCol, catInfo] of categoryAtCol) {
      if (catCol < i && catCol > bestCol) {
        bestCol = catCol;
        subCat  = catInfo.subCat;
        cat     = catInfo.cat;
      }
    }

    // Find "IS THE ... ON THE FLOOR?" column for this product
    let onFloorIdx: number | null = null;
    let reasonIdx:  number | null = null;
    const pcUpper = productCode.toUpperCase();

    for (let j = Math.max(0, i - 5); j < i; j++) {
      const h2 = headers[j];
      if (h2 && h2.toUpperCase().startsWith('IS THE') && h2.toUpperCase().includes(pcUpper)) {
        onFloorIdx = j;
      }
    }
    for (let j = i + 1; j < Math.min(headers.length, i + 5); j++) {
      const h2 = headers[j];
      if (h2 && h2.toUpperCase().startsWith('WHAT IS THE REASON') && h2.toUpperCase().includes(pcUpper)) {
        reasonIdx = j;
      }
    }

    sohCols.push({ colIdx: i, header: h, productCode, subCat, cat, onFloorIdx, reasonIdx });
  }

  // ── 4. Build vertical database ────────────────────────────────────────────
  const db: DatabaseRow[] = [];
  const dates: string[] = [];

  for (const row of dataRows) {
    const repFirst = String(row[2] ?? '').trim();
    const repLast  = String(row[3] ?? '').trim();
    const repName  = [repFirst, repLast].filter(Boolean).join(' ') || 'UNKNOWN';
    const storeName = String(row[6] ?? '').trim() || 'UNKNOWN';
    const storeCode = String(row[7] ?? '').trim();
    const province  = String(row[8] ?? '').trim();
    const dateStr   = String(row[9] ?? '').trim();

    if (dateStr && dateStr.includes('/')) dates.push(dateStr);

    for (const col of sohCols) {
      const raw = row[col.colIdx];
      if (raw === null || raw === undefined || raw === '') continue;
      const sohNum = Number(raw);
      if (isNaN(sohNum)) continue;

      const onFloor = col.onFloorIdx !== null ? String(row[col.onFloorIdx] ?? '').trim() : '';
      const reason  = col.reasonIdx  !== null ? String(row[col.reasonIdx]  ?? '').trim() : '';

      let weekLabel = '';
      if (dateStr && dateStr.includes('/')) {
        const { week } = getISOWeek(dateStr);
        weekLabel = `WEEK${String(week).padStart(2, '0')}`;
      }

      db.push({
        week:          weekLabel,
        storeName,
        storeCode,
        repName,
        date:          dateStr,
        subCategory:   col.subCat,
        productCode:   col.productCode,
        product:       col.header,
        soh:           sohNum,
        onFloor,
        reason,
        province,
        category:      col.cat,
        keyCat:        col.cat,
        criticalLines: CRITICAL_LINE_CODES.has(col.productCode) ? 'YES' : 'NO',
      });
    }
  }

  // ── 5. Determine report week / year ───────────────────────────────────────
  let reportWeek = 1;
  let reportYear = new Date().getFullYear();
  if (dates.length > 0) {
    const latest = [...dates].sort().at(-1)!;
    const r = getISOWeek(latest);
    reportWeek = r.week;
    reportYear = r.year;
  }

  const weekLabel = `WK${String(reportWeek).padStart(2, '0')}`;
  const filename  = `${brand.toUpperCase()}_MAKRO_STOCK_COUNT_${weekLabel}_${reportYear}.xlsx`;

  // ── 6. Build output Excel ─────────────────────────────────────────────────
  const buf = await buildOutputExcel(db, brand, weekLabel, reportYear);
  return { buffer: buf, filename };
}

// ─── Excel builder ────────────────────────────────────────────────────────────
async function buildOutputExcel(
  db: DatabaseRow[],
  brand: string,
  weekLabel: string,
  year: number,
): Promise<Buffer> {

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Defy Field Execute';
  wb.created = new Date();

  // Load logos (graceful fallback if missing)
  const pub = path.join(process.cwd(), 'public');
  let defyLogoId:    number | null = null;
  let atomicLogoId:  number | null = null;
  let perigeeLogoId: number | null = null;

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    defyLogoId = wb.addImage({ buffer: fs.readFileSync(path.join(pub, 'defy-logo.png')) as any, extension: 'png' });
  } catch { /* logo unavailable */ }
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    atomicLogoId = wb.addImage({ buffer: fs.readFileSync(path.join(pub, 'atomic-logo.png')) as any, extension: 'png' });
  } catch { /* logo unavailable */ }
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    perigeeLogoId = wb.addImage({ buffer: fs.readFileSync(path.join(pub, 'perigee-logo.jpg')) as any, extension: 'jpeg' });
  } catch { /* logo unavailable */ }

  const reportTitle = `${brand.toUpperCase()}  |  MAKRO STOCK COUNT  |  ${weekLabel} ${year}`;

  // Adds a red logo/title banner as row 1, returns the column count used
  function addBannerRow(ws: ExcelJS.Worksheet, colCount: number) {
    ws.getRow(1).height = 48;

    // Fill entire banner row red
    for (let c = 1; c <= colCount; c++) {
      const cell = ws.getRow(1).getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DEFY_RED } };
      cell.border = {
        top:    { style: 'thin', color: { argb: DEFY_RED } },
        bottom: { style: 'thin', color: { argb: DEFY_RED } },
        left:   { style: 'thin', color: { argb: DEFY_RED } },
        right:  { style: 'thin', color: { argb: DEFY_RED } },
      };
    }

    // Title text
    const titleCell = ws.getRow(1).getCell(2);
    titleCell.value = reportTitle;
    titleCell.font  = { bold: true, color: { argb: WHITE }, size: 13, name: 'Arial' };
    titleCell.alignment = { vertical: 'middle', horizontal: 'left' };

    // Logos
    if (defyLogoId !== null) {
      ws.addImage(defyLogoId, { tl: { col: 0, row: 0 }, ext: { width: 80, height: 40 } });
    }
    if (atomicLogoId !== null) {
      ws.addImage(atomicLogoId, { tl: { col: colCount - 2, row: 0 }, ext: { width: 138, height: 39 } });
    }
    if (perigeeLogoId !== null) {
      ws.addImage(perigeeLogoId, { tl: { col: colCount - 1, row: 0 }, ext: { width: 80, height: 24 } });
    }
  }

  const zeroRows = db.filter(r => r.soh === 0);

  // ── Sheet 1: DATABASE ─────────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('DATABASE');
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3, topLeftCell: 'A4' }];

    const cols = [
      'WEEK','STORE NAME','STORE CODE','REP NAME','DATE',
      'SUB CATEGORY','PRODUCT CODE','PRODUCT','SOH',
      'ON FLOOR','REASON','PROVINCE','CATEGORY','KEY CAT','CRITICAL LINES',
    ];
    addBannerRow(ws, cols.length);
    ws.mergeCells(1, 1, 1, cols.length);
    ws.addRow([]); // blank spacer

    const hRow = ws.addRow(cols);
    hRow.height = 30;
    hRow.eachCell(cell => applyHeaderStyle(cell));

    db.forEach((row, i) => {
      const r = ws.addRow([
        row.week, row.storeName, row.storeCode, row.repName, row.date,
        row.subCategory, row.productCode, row.product, row.soh,
        row.onFloor, row.reason, row.province, row.category, row.keyCat, row.criticalLines,
      ]);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));
    });

    [10,28,12,22,12,18,14,52,8,10,32,15,15,15,14]
      .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  }

  // ── Sheet 2: DETAIL ───────────────────────────────────────────────────────
  const detailCols = [
    'STORE NAME','SUB CATEGORY','PRODUCT CODE','PRODUCT','SOH',
    'PROVINCE','CATEGORY','KEY CAT','CRITICAL LINES',
  ];
  const buildDetailSheet = (ws: ExcelJS.Worksheet, rows: DatabaseRow[]) => {
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3, topLeftCell: 'A4' }];
    addBannerRow(ws, detailCols.length);
    ws.mergeCells(1, 1, 1, detailCols.length);
    ws.addRow([]);

    const hRow = ws.addRow(detailCols);
    hRow.height = 30;
    hRow.eachCell(cell => applyHeaderStyle(cell));

    rows.forEach((row, i) => {
      const r = ws.addRow([
        row.storeName, row.subCategory, row.productCode, row.product,
        row.soh, row.province, row.category, row.keyCat, row.criticalLines,
      ]);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));
    });

    [28,18,14,52,8,15,15,15,14]
      .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  };

  buildDetailSheet(wb.addWorksheet('DETAIL'), db);

  // ── Sheet 3: ZERO SOH ─────────────────────────────────────────────────────
  buildDetailSheet(wb.addWorksheet('ZERO SOH'), zeroRows);

  // ── Sheet 4: SUMMARY (Product × Store pivot) ──────────────────────────────
  {
    const ws = wb.addWorksheet('SUMMARY');

    const stores   = [...new Set(db.map(r => r.storeName))].sort();
    const products = [...new Set(db.map(r => r.product))].sort();

    // Build pivot lookup: product → store → sum SOH
    const pivot = new Map<string, Map<string, number>>();
    for (const row of db) {
      if (!pivot.has(row.product)) pivot.set(row.product, new Map());
      const sm = pivot.get(row.product)!;
      sm.set(row.storeName, (sm.get(row.storeName) ?? 0) + row.soh);
    }

    const totalCols = stores.length + 2; // product col + store cols + total col
    addBannerRow(ws, totalCols);
    ws.mergeCells(1, 1, 1, totalCols);

    ws.addRow([]); // blank
    const weekRow = ws.addRow(['WEEK', weekLabel]);
    weekRow.getCell(1).font = { bold: true, size: 10, name: 'Arial' };
    weekRow.getCell(2).font = { bold: true, color: { argb: DEFY_RED }, size: 10, name: 'Arial' };
    ws.addRow([]); // blank

    // Header row with rotated store names
    const hRow = ws.addRow(['PRODUCT', ...stores, 'TOTAL']);
    hRow.height = 150;
    hRow.eachCell(cell => {
      applyHeaderStyle(cell);
      cell.alignment = { vertical: 'bottom', horizontal: 'center', textRotation: 90, wrapText: false };
    });
    hRow.getCell(1).alignment          = { vertical: 'middle', horizontal: 'left' };
    hRow.getCell(stores.length + 2).alignment = { vertical: 'middle', horizontal: 'center' };

    // Product rows
    products.forEach((product, i) => {
      const storeMap = pivot.get(product) ?? new Map();
      const vals: (string | number)[] = [product];
      let total = 0;
      for (const store of stores) {
        const soh = storeMap.get(store) ?? null;
        vals.push(soh !== null ? soh : '');
        if (soh !== null) total += soh;
      }
      vals.push(total);

      const r = ws.addRow(vals);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));

      // Highlight zero SOH red
      for (let c = 2; c <= stores.length + 1; c++) {
        const cell = r.getCell(c);
        if (cell.value === 0) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCC' } };
          cell.font = { color: { argb: DEFY_RED }, bold: true, size: 10, name: 'Arial' };
        }
      }
      r.getCell(stores.length + 2).font = { bold: true, size: 10, name: 'Arial' };
    });

    // Grand total row
    const gtVals: (string | number)[] = ['GRAND TOTAL'];
    let gtTotal = 0;
    for (const store of stores) {
      let storeTotal = 0;
      for (const sm of pivot.values()) storeTotal += sm.get(store) ?? 0;
      gtTotal += storeTotal;
      gtVals.push(storeTotal);
    }
    gtVals.push(gtTotal);
    const gtRow = ws.addRow(gtVals);
    gtRow.eachCell(cell => {
      applyHeaderStyle(cell);
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    gtRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };

    ws.getColumn(1).width = 52;
    for (let c = 2; c <= stores.length + 1; c++) ws.getColumn(c).width = 7.15;
    ws.getColumn(stores.length + 2).width = 10;
  }

  // ── Sheet 5: INSIGHTS ────────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('INSIGHTS');
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3, topLeftCell: 'A4' }];

    addBannerRow(ws, 9);
    ws.mergeCells(1, 1, 1, 9);
    ws.addRow([]);

    const hRow = ws.addRow([
      'CATEGORY', 'TOTAL SOH', 'SOH %', '',
      'STORE NAME', 'OOS CONT %', '',
      'PRODUCT CODE', 'OOS CONT %',
    ]);
    hRow.height = 30;
    hRow.eachCell((cell, c) => {
      if (c === 4 || c === 7) return; // spacer columns
      applyHeaderStyle(cell);
    });

    // Category totals (by subCategory)
    const catTotals = new Map<string, number>();
    for (const row of db) catTotals.set(row.subCategory, (catTotals.get(row.subCategory) ?? 0) + row.soh);
    const totalSoh = [...catTotals.values()].reduce((a, b) => a + b, 0);
    const sortedCats = [...catTotals.entries()].sort((a, b) => b[1] - a[1]);

    // Store OOS counts
    const storeOOS = new Map<string, number>();
    for (const row of zeroRows) storeOOS.set(row.storeName, (storeOOS.get(row.storeName) ?? 0) + 1);
    const totalOOS = zeroRows.length;
    const sortedStoreOOS = [...storeOOS.entries()].sort((a, b) => b[1] - a[1]);

    // Product OOS counts
    const prodOOS = new Map<string, number>();
    for (const row of zeroRows) prodOOS.set(row.productCode, (prodOOS.get(row.productCode) ?? 0) + 1);
    const sortedProdOOS = [...prodOOS.entries()].sort((a, b) => b[1] - a[1]);

    const maxR = Math.max(sortedCats.length + 1, sortedStoreOOS.length, sortedProdOOS.length);

    for (let i = 0; i < maxR; i++) {
      const vals: (string | number | null)[] = [null, null, null, null, null, null, null, null, null];

      if (i < sortedCats.length) {
        vals[0] = sortedCats[i][0];
        vals[1] = sortedCats[i][1];
        vals[2] = totalSoh > 0 ? sortedCats[i][1] / totalSoh : 0;
      } else if (i === sortedCats.length) {
        vals[0] = 'Grand Total';
        vals[1] = totalSoh;
        vals[2] = 1;
      }

      if (i < sortedStoreOOS.length) {
        vals[4] = sortedStoreOOS[i][0];
        vals[5] = totalOOS > 0 ? sortedStoreOOS[i][1] / totalOOS : 0;
      }

      if (i < sortedProdOOS.length) {
        vals[7] = sortedProdOOS[i][0];
        vals[8] = totalOOS > 0 ? sortedProdOOS[i][1] / totalOOS : 0;
      }

      const r = ws.addRow(vals);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));
      if (vals[2] !== null) r.getCell(3).numFmt = '0.00%';
      if (vals[5] !== null) r.getCell(6).numFmt = '0.00%';
      if (vals[8] !== null) r.getCell(9).numFmt = '0.00%';

      // Bold grand total row
      if (i === sortedCats.length) {
        r.getCell(1).font = { bold: true, size: 10, name: 'Arial' };
        r.getCell(2).font = { bold: true, size: 10, name: 'Arial' };
        r.getCell(3).font = { bold: true, size: 10, name: 'Arial' };
      }
    }

    [22, 12, 12, 3, 32, 14, 3, 20, 14]
      .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  }

  // ── Sheet 6: CRITICAL LINE OOS ────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('CRITICAL LINE OOS');
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 4, topLeftCell: 'A5' }];

    const critOOS = zeroRows.filter(r => r.criticalLines === 'YES');
    addBannerRow(ws, 8);
    ws.mergeCells(1, 1, 1, 8);
    ws.addRow([]);

    // Section headers
    const secRow = ws.addRow(['ARTICLE VIEW', null, null, null, 'ARTICLE BY SITE', null, null, null]);
    [1, 5].forEach(c => {
      const cell = secRow.getCell(c);
      applyHeaderStyle(cell);
      cell.font = { bold: true, color: { argb: WHITE }, size: 11, name: 'Arial' };
    });
    secRow.height = 28;

    const subRow = ws.addRow([
      'PRODUCT', 'SOH', null, null,
      'STORE NAME', 'PRODUCT CODE', 'SOH', 'OOS CONT %',
    ]);
    subRow.height = 28;
    subRow.eachCell((cell, c) => {
      if (c === 3 || c === 4) return;
      applyHeaderStyle(cell);
    });

    // Article view: product → OOS count
    const articleView = new Map<string, number>();
    for (const row of critOOS) articleView.set(row.product, (articleView.get(row.product) ?? 0) + 1);
    const articleViewArr = [...articleView.entries()].sort((a, b) => b[1] - a[1]);

    const totalCritOOS = critOOS.length;
    const maxR = Math.max(articleViewArr.length + 1, critOOS.length);

    for (let i = 0; i < maxR; i++) {
      const vals: (string | number | null)[] = [null, null, null, null, null, null, null, null];

      if (i < articleViewArr.length) {
        vals[0] = articleViewArr[i][0];
        vals[1] = articleViewArr[i][1];
      } else if (i === articleViewArr.length) {
        vals[0] = 'Grand Total';
        vals[1] = totalCritOOS;
      }

      if (i < critOOS.length) {
        const row = critOOS[i];
        vals[4] = row.storeName;
        vals[5] = row.productCode;
        vals[6] = row.soh;
        vals[7] = totalCritOOS > 0 ? 1 / totalCritOOS : 0;
      }

      const r = ws.addRow(vals);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));
      if (vals[7] !== null) r.getCell(8).numFmt = '0.00%';

      if (i === articleViewArr.length) {
        r.getCell(1).font = { bold: true, size: 10, name: 'Arial' };
        r.getCell(2).font = { bold: true, size: 10, name: 'Arial' };
      }
    }

    [52, 10, 3, 3, 32, 18, 10, 14]
      .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  }

  // Write and return
  const buf = await wb.xlsx.writeBuffer();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return Buffer.from(buf) as any;
}

import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import path from 'path';
import fs from 'fs';
import type { StoreMapEntry } from '@/lib/storeMapData';
import { buildMenuSheet, applyHeaderStyle, applyDataStyle, addNavRow } from './build-menu';

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

// Brand colour (used locally for pivot highlights etc.)
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

// ─── Local helpers ────────────────────────────────────────────────────────────
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


// ─── Province lookup helper ───────────────────────────────────────────────────
function lookupProvince(storeCode: string, storeName: string, storeMap: StoreMapEntry[]): string {
  if (storeCode) {
    const e = storeMap.find(x => x.storeCode.toUpperCase() === storeCode.toUpperCase());
    if (e) return e.province;
  }
  if (storeName) {
    const e = storeMap.find(x => x.storeName.toUpperCase() === storeName.toUpperCase());
    if (e) return e.province;
  }
  return '';
}

// ─── Perigee DD/MM/YYYY → Date ───────────────────────────────────────────────
function perigeeToDate(dateStr: string): Date {
  const parts = String(dateStr).split('/');
  if (parts.length !== 3) return new Date(0);
  const [d, m, y] = parts.map(Number);
  return new Date(y, m - 1, d);
}

// ─── Pre-generate analysis ───────────────────────────────────────────────────
// Parses only headers + row metadata — no Excel generation. Returns warnings
// the user should acknowledge, or a hardError that blocks generation entirely.
export function analyzeStockCount(fileBuffer: Buffer): { warnings: string[]; hardError: string | null } {
  const inputWb = XLSX.read(fileBuffer, { type: 'buffer' });
  const inputWs = inputWb.Sheets[inputWb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(inputWs, { header: 1, defval: null }) as (string | number | null)[][];

  if (rawData.length < 2) {
    return { warnings: [], hardError: 'No data rows found in the uploaded file.' };
  }

  const headers  = rawData[0] as (string | null)[];
  const rawRows  = rawData.slice(1);
  const warnings: string[] = [];

  // 1. Count rows where the rep did not complete store check-in
  const blankStoreCount = rawRows.filter(row =>
    !String(row[7] ?? '').trim() && !String(row[6] ?? '').trim()
  ).length;
  if (blankStoreCount > 0) {
    warnings.push(
      `${blankStoreCount} row${blankStoreCount === 1 ? '' : 's'} have no Store Name or Store Code — ` +
      `${blankStoreCount === 1 ? 'this rep' : 'these reps'} did not complete their store check-in ` +
      `and will appear as "UNKNOWN" in the report.`
    );
  }

  // 2. Find critical lines boundary and count SOH / on-floor / reason columns
  let criticalLinesStartCol = Infinity;
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (h && h.trim() === 'CRITICAL LINES') { criticalLinesStartCol = i; break; }
  }

  let sohCount     = 0;
  let onFloorCount = 0;
  let reasonCount  = 0;

  for (let i = 0; i < headers.length && i < criticalLinesStartCol; i++) {
    const h = headers[i];
    if (!h || !h.toUpperCase().includes('SOH')) continue;
    sohCount++;
    const pc = extractProductCode(h).toUpperCase();
    for (let j = Math.max(0, i - 5); j < i; j++) {
      const h2 = headers[j];
      if (h2 && h2.toUpperCase().startsWith('IS THE') && h2.toUpperCase().includes(pc)) { onFloorCount++; break; }
    }
    for (let j = i + 1; j < Math.min(headers.length, i + 5); j++) {
      const h2 = headers[j];
      if (h2 && h2.toUpperCase().startsWith('WHAT IS THE REASON') && h2.toUpperCase().includes(pc)) { reasonCount++; break; }
    }
  }

  if (sohCount === 0) {
    return {
      warnings: [],
      hardError:
        'No product SOH columns found in the uploaded file. ' +
        'Expected column headers containing "SOH" (e.g. "DBO489E 50cm Built-In Oven SOH"). ' +
        'Check that you have uploaded the correct Perigee raw export for this report.',
    };
  }

  if (onFloorCount === 0 && reasonCount === 0) {
    warnings.push(
      'This Perigee form does not include "Is the product on the floor?" or "What is the reason?" questions — ' +
      'the ON FLOOR and REASON columns will be blank in all sheets. ' +
      'This is normal for forms that only capture stock counts (e.g. Beko). ' +
      'All other sheets (SUMMARY, DETAIL, ZERO SOH, DATABASE) will still be fully populated.'
    );
  } else {
    if (onFloorCount === 0) warnings.push('No "Is the product on the floor?" questions found — the ON FLOOR column will be blank in all sheets.');
    if (reasonCount  === 0) warnings.push('No "What is the reason?" questions found — the REASON column will be blank in all sheets.');
  }

  return { warnings, hardError: null };
}

// ─── Main export ─────────────────────────────────────────────────────────────
export async function generateMakroStockCount(
  fileBuffer: Buffer,
  brand: string,
  storeMap: StoreMapEntry[] = [],
  retailer = 'MAKRO',
): Promise<{ buffer: Buffer; filename: string; rawDates: string[]; weekLabel: string }> {

  // ── 1. Parse raw Perigee Excel ─────────────────────────────────────────────
  const inputWb = XLSX.read(fileBuffer, { type: 'buffer' });
  const inputWs = inputWb.Sheets[inputWb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(inputWs, { header: 1, defval: null }) as (string | number | null)[][];

  if (rawData.length < 2) throw new Error('No data rows found in uploaded file.');

  const headers = rawData[0] as (string | null)[];
  const rawRows = rawData.slice(1);

  // ── Deduplicate: same store → keep most recent submission by DATE (col J = index 9) ──
  // Some reps fill in the form multiple times for the same store in the same period.
  const storeLatest = new Map<string, { row: typeof rawRows[0]; dateMs: number }>();
  for (let ri = 0; ri < rawRows.length; ri++) {
    const row = rawRows[ri];
    const storeCode = String(row[7] ?? '').trim();
    const storeName = String(row[6] ?? '').trim();
    // Rows with no store check-in get a unique key so they're included as UNKNOWN
    const key = storeCode || storeName || `__NOSTORE_${ri}`;
    const dateMs = perigeeToDate(String(row[9] ?? '')).getTime();
    if (!storeLatest.has(key) || dateMs > storeLatest.get(key)!.dateMs) {
      storeLatest.set(key, { row, dateMs });
    }
  }
  const dataRows = [...storeLatest.values()].map(v => v.row);

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

  if (sohCols.length === 0) {
    throw new Error(
      'No product SOH columns found in the uploaded file. ' +
      'Expected column headers containing "SOH" (e.g. "DBO489E 50cm Built-In Oven SOH"). ' +
      'Check that you have uploaded the correct Perigee raw export for this report.'
    );
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
    const dateStr   = String(row[9] ?? '').trim();
    // Province: look up from control file first; fall back to raw col 8 if present
    const rawProvince = String(row[8] ?? '').trim();
    const province = lookupProvince(storeCode, storeName, storeMap) || rawProvince || 'UNKNOWN';

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
        product:       col.header.replace(/\s+SOH\s*$/i, '').trim(),
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
  const retailerSlug = retailer.toUpperCase().replace(/\s+/g, '_');
  const filename  = `${brand.toUpperCase()}_${retailerSlug}_STOCK_COUNT_${weekLabel}_${reportYear}.xlsx`;

  // ── 6. Build output Excel ─────────────────────────────────────────────────
  const buf = await buildOutputExcel(db, brand, retailer, weekLabel, reportYear);
  return { buffer: buf, filename, rawDates: dates, weekLabel };
}

// ─── Internal hyperlink helper ───────────────────────────────────────────────
// Wraps sheet names with spaces in single quotes as Excel requires
function sheetHref(sheetName: string, cell = 'A1') {
  const safe = sheetName.includes(' ') ? `'${sheetName}'` : sheetName;
  return `#${safe}!${cell}`;
}


// ─── Excel builder ────────────────────────────────────────────────────────────
async function buildOutputExcel(
  db: DatabaseRow[],
  brand: string,
  retailer: string,
  weekLabel: string,
  year: number,
): Promise<Buffer> {

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Defy Field Execute';
  wb.created = new Date();

  // ── Load logos with correct aspect ratios ──────────────────────────────────
  // Actual dimensions verified from file headers:
  //   defy-logo.png    224×224  (square,   ratio 1.00)
  //   atomic-logo.png 1024×216  (wide,     ratio 4.74)
  //   perigee-logo.jpg 4151×4208 (square,  ratio 0.99)
  // Display sizes are 50% larger than previous values, maintaining true ratios.
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

  const zeroRows = db.filter(r => r.soh === 0);

  // ── Sheet definitions (order matches example report) ─────────────────────
  const SHEET_DEFS = [
    { name: 'CRITICAL LINE OOS', desc: 'Critical products with zero stock (OOS)' },
    { name: 'INSIGHTS',          desc: 'SOH by category · Store OOS % · Product OOS %' },
    { name: 'SUMMARY',           desc: 'Product × store SOH pivot · use ▼ CATEGORY / SUB CAT filters to slice by product type' },
    { name: 'DETAIL',            desc: 'All products across all visits · use ▼ filters to slice by Province, Category, etc.' },
    { name: 'ZERO SOH',          desc: 'Out-of-stock lines only (SOH = 0) · filterable' },
    { name: 'DATABASE',          desc: 'Full vertical database — all visits, all products · filterable' },
  ];

  // ── MENU sheet ─────────────────────────────────────────────────────────────
  {
    const genDate = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' });
    buildMenuSheet(wb, {
      title:        `${brand.toUpperCase()} — ${retailer.toUpperCase()} STOCK COUNT`,
      subtitle:     `${weekLabel} · ${year}  ·  Generated: ${genDate}`,
      sheetDefs:    SHEET_DEFS,
      note:         '💡  Tip: On SUMMARY use ▼ CATEGORY / SUB CAT filters to slice the pivot by product type. On DETAIL, ZERO SOH and DATABASE use ▼ filters to filter by Province, Category, Store and more.',
      defyLogoId,
      atomicLogoId,
      perigeeLogoId,
    });
  }

  // ── CRITICAL LINE OOS ─────────────────────────────────────────────────────
  // Flat filterable table: PROVINCE | STORE NAME | STORE CODE | REP NAME | DATE
  //   | PRODUCT CODE | PRODUCT | CATEGORY | SUB CAT | SOH
  {
    const ws = wb.addWorksheet('CRITICAL LINE OOS');
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2, topLeftCell: 'A3', showGridLines: false }];

    const critOOS = zeroRows
      .filter(r => r.criticalLines === 'YES')
      .sort((a, b) => {
        if (a.province  !== b.province)  return a.province.localeCompare(b.province);
        if (a.storeName !== b.storeName) return a.storeName.localeCompare(b.storeName);
        return a.product.localeCompare(b.product);
      });

    const cols = [
      'PROVINCE', 'STORE NAME', 'STORE CODE', 'REP NAME', 'DATE',
      'PRODUCT CODE', 'PRODUCT', 'CATEGORY', 'SUB CAT', 'SOH',
    ];

    addNavRow(ws, 'CRITICAL LINE OOS', cols.length);

    const hRow = ws.addRow(cols);
    hRow.height = 28;
    hRow.eachCell(cell => applyHeaderStyle(cell));

    ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: cols.length } };

    critOOS.forEach((row, i) => {
      const r = ws.addRow([
        row.province, row.storeName, row.storeCode, row.repName, row.date,
        row.productCode, row.product, row.category, row.subCategory, row.soh,
      ]);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));
    });

    [15, 30, 12, 22, 12, 14, 52, 15, 15, 8]
      .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  }

  // ── INSIGHTS ──────────────────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('INSIGHTS');
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2, topLeftCell: 'A3', showGridLines: false }];

    addNavRow(ws, 'INSIGHTS', 9);

    const hRow = ws.addRow([
      'CATEGORY', 'TOTAL SOH', 'SOH %', '',
      'STORE NAME', 'OOS CONT %', '',
      'PRODUCT CODE', 'OOS CONT %',
    ]);
    hRow.height = 28;
    hRow.eachCell((cell, c) => {
      if (c === 4 || c === 7) return;
      applyHeaderStyle(cell);
    });

    const catTotals = new Map<string, number>();
    for (const row of db) catTotals.set(row.subCategory, (catTotals.get(row.subCategory) ?? 0) + row.soh);
    const totalSoh = [...catTotals.values()].reduce((a, b) => a + b, 0);
    const sortedCats = [...catTotals.entries()].sort((a, b) => b[1] - a[1]);

    const storeOOS = new Map<string, number>();
    for (const row of zeroRows) storeOOS.set(row.storeName, (storeOOS.get(row.storeName) ?? 0) + 1);
    const totalOOS = zeroRows.length;
    const sortedStoreOOS = [...storeOOS.entries()].sort((a, b) => b[1] - a[1]);

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
        vals[0] = 'Grand Total'; vals[1] = totalSoh; vals[2] = 1;
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
      if (i === sortedCats.length) {
        [1, 2, 3].forEach(c => { r.getCell(c).font = { bold: true, size: 10, name: 'Arial' }; });
      }
    }
    [22, 12, 12, 3, 32, 14, 3, 20, 14]
      .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  }

  // ── SUMMARY (Product × Store pivot) ───────────────────────────────────────
  // Col layout: CATEGORY | SUB CAT | PRODUCT | [STORE1...N] | TOTAL
  // AutoFilter on row 5 lets users filter by CATEGORY or SUB CAT to slice the
  // pivot (e.g. show only FRIDGE rows). Province filtering belongs on DETAIL/DATABASE.
  {
    const ws = wb.addWorksheet('SUMMARY');

    const stores   = [...new Set(db.map(r => r.storeName))].sort();
    // Sort products by category then subCat then name for logical grouping
    const productMeta = new Map<string, { category: string; subCat: string }>();
    for (const row of db) {
      if (!productMeta.has(row.product)) {
        productMeta.set(row.product, { category: row.category, subCat: row.subCategory });
      }
    }
    const products = [...new Set(db.map(r => r.product))].sort((a, b) => {
      const ma = productMeta.get(a)!; const mb = productMeta.get(b)!;
      if (ma.category !== mb.category) return ma.category.localeCompare(mb.category);
      if (ma.subCat !== mb.subCat) return ma.subCat.localeCompare(mb.subCat);
      return a.localeCompare(b);
    });

    const pivot = new Map<string, Map<string, number>>();
    for (const row of db) {
      if (!pivot.has(row.product)) pivot.set(row.product, new Map());
      const sm = pivot.get(row.product)!;
      sm.set(row.storeName, (sm.get(row.storeName) ?? 0) + row.soh);
    }

    // 3 fixed cols (CATEGORY, SUB CAT, PRODUCT) + N store cols + TOTAL col
    const totalCols = stores.length + 4;
    addNavRow(ws, 'SUMMARY', totalCols);

    ws.addRow([]); // blank row 2
    const weekRow = ws.addRow(['WEEK', weekLabel]);
    weekRow.getCell(1).font = { bold: true, size: 10, name: 'Arial' };
    weekRow.getCell(2).font = { bold: true, color: { argb: DEFY_RED }, size: 10, name: 'Arial' };
    ws.addRow([]); // blank row 4

    // Header row with rotated store names (row 5)
    const hRow = ws.addRow(['CATEGORY', 'SUB CAT', 'PRODUCT', ...stores, 'TOTAL']);
    hRow.height = 150;
    hRow.eachCell(cell => {
      applyHeaderStyle(cell);
      cell.alignment = { vertical: 'bottom', horizontal: 'center', textRotation: 90, wrapText: false };
    });
    // CATEGORY, SUB CAT, PRODUCT headers — left-aligned, not rotated
    [1, 2, 3].forEach(c => {
      hRow.getCell(c).alignment = { vertical: 'middle', horizontal: 'left' };
    });
    hRow.getCell(stores.length + 4).alignment = { vertical: 'middle', horizontal: 'center' };

    // AutoFilter on row 5: users can filter CATEGORY or SUB CAT to slice the pivot
    ws.autoFilter = { from: { row: 5, column: 1 }, to: { row: 5, column: totalCols } };

    // Freeze first 3 cols + rows 1-5
    ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 5, topLeftCell: 'D6', showGridLines: false }];

    products.forEach((product, i) => {
      const meta = productMeta.get(product)!;
      const storeMap = pivot.get(product) ?? new Map();
      const vals: (string | number)[] = [meta.category, meta.subCat, product];
      let total = 0;
      for (const store of stores) {
        const soh = storeMap.get(store) ?? null;
        vals.push(soh !== null ? soh : '');
        if (soh !== null) total += soh;
      }
      vals.push(total);
      const r = ws.addRow(vals);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));
      // Highlight zero SOH cells (store columns only — cols 4 to stores.length+3)
      for (let c = 4; c <= stores.length + 3; c++) {
        const cell = r.getCell(c);
        if (cell.value === 0) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCC' } };
          cell.font = { color: { argb: DEFY_RED }, bold: true, size: 10, name: 'Arial' };
        }
      }
      r.getCell(stores.length + 4).font = { bold: true, size: 10, name: 'Arial' };
    });

    const gtVals: (string | number)[] = ['GRAND TOTAL', '', ''];
    let gtTotal = 0;
    for (const store of stores) {
      let st = 0;
      for (const sm of pivot.values()) st += sm.get(store) ?? 0;
      gtTotal += st;
      gtVals.push(st);
    }
    gtVals.push(gtTotal);
    const gtRow = ws.addRow(gtVals);
    gtRow.eachCell(cell => {
      applyHeaderStyle(cell);
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
    gtRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };

    ws.getColumn(1).width = 18;  // CATEGORY
    ws.getColumn(2).width = 18;  // SUB CAT
    ws.getColumn(3).width = 52;  // PRODUCT
    for (let c = 4; c <= stores.length + 3; c++) ws.getColumn(c).width = 7.15;
    ws.getColumn(stores.length + 4).width = 10;
  }

  // ── DETAIL + ZERO SOH shared builder ──────────────────────────────────────
  const detailCols = [
    'STORE NAME','SUB CATEGORY','PRODUCT CODE','PRODUCT','SOH',
    'PROVINCE','CATEGORY','KEY CAT','CRITICAL LINES',
  ];
  const buildDetailSheet = (ws: ExcelJS.Worksheet, title: string, rows: DatabaseRow[]) => {
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2, topLeftCell: 'A3', showGridLines: false }];
    addNavRow(ws, title, detailCols.length);

    const hRow = ws.addRow(detailCols);
    hRow.height = 28;
    hRow.eachCell(cell => applyHeaderStyle(cell));

    // AutoFilter on header row (row 2)
    ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: detailCols.length } };

    rows.forEach((row, i) => {
      const r = ws.addRow([
        row.storeName, row.subCategory, row.productCode, row.product,
        row.soh, row.province, row.category, row.keyCat, row.criticalLines,
      ]);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));
    });
    [28, 18, 14, 52, 8, 15, 15, 15, 14]
      .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  };

  buildDetailSheet(wb.addWorksheet('DETAIL'),   'DETAIL',   db);
  buildDetailSheet(wb.addWorksheet('ZERO SOH'), 'ZERO SOH', zeroRows);

  // ── DATABASE ──────────────────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('DATABASE');
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2, topLeftCell: 'A3', showGridLines: false }];

    const dbCols = [
      'WEEK','STORE NAME','STORE CODE','REP NAME','DATE',
      'SUB CATEGORY','PRODUCT CODE','PRODUCT','SOH',
      'ON FLOOR','REASON','PROVINCE','CATEGORY','KEY CAT','CRITICAL LINES',
    ];
    addNavRow(ws, 'DATABASE', dbCols.length);

    const hRow = ws.addRow(dbCols);
    hRow.height = 28;
    hRow.eachCell(cell => applyHeaderStyle(cell));

    ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: dbCols.length } };

    db.forEach((row, i) => {
      const r = ws.addRow([
        row.week, row.storeName, row.storeCode, row.repName, row.date,
        row.subCategory, row.productCode, row.product, row.soh,
        row.onFloor, row.reason, row.province, row.category, row.keyCat, row.criticalLines,
      ]);
      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));
    });
    [10, 28, 12, 22, 12, 18, 14, 52, 8, 10, 32, 15, 15, 15, 14]
      .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  }

  return Buffer.from(await wb.xlsx.writeBuffer());
}

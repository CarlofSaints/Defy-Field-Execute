/**
 * Service Call Report builder
 *
 * Input:  Perigee raw export ("Worksheet" sheet, 22 cols)
 * Output: Excel workbook — MENU + data sheet
 *
 * Column mapping (0-indexed, header at row 0, data from row 1):
 *   0   ID
 *   1   Email
 *   2   First Name
 *   3   Last Name
 *   4   Customer
 *   5   Channel
 *   6   Store
 *   7   Store Code
 *   8   Province
 *   9   Date (DD/MM/YYYY)
 *  10   Time
 *  11   Visit UUID         ← added by Perigee, shifts everything below by +1
 *  12   Tag
 *  13   Sync Date
 *  14   Sync Time
 *  15   IS THIS A NEW LOG OR RECURRING?
 *  16   PLEASE ENTER THE LOG NUMBER
 *  17   WHAT IS THE PRODUCT THAT NEEDS TO BE ATTENDED TO?
 *  18   THE DATE WHEN THE SERVICE CALL WAS LOGGED (DD/MM/YYYY)
 *  19   NAME OF THE CUSTOMER
 *  20   CUSTOMER PHONE NUMBER
 *  21   PLEASE GIVE A FULL DESCRIPTION OF THE MATTER
 *
 * No images in this form.
 */
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import fs   from 'fs';
import path from 'path';
import { buildMenuSheet, applyHeaderStyle, applyDataStyle, addNavRow } from './build-menu';

// ─── Constants ────────────────────────────────────────────────────────────────
const OUT_COLS = [
  'REPRESENTATIVE NAME',                       // 1
  'DATE',                                      // 2
  'PLACE',                                     // 3
  'NEW LOG OR RECURRING',                      // 4
  'LOG NUMBER',                                // 5
  'PRODUCT THAT NEEDS TO BE ATTENDED TO',      // 6
  'THE DATE WHEN THE SERVICE CALL WAS LOGGED', // 7
  'DAYS SINCE SERVICE CALL WAS LOGGED',        // 8
  'CUSTOMER NAME',                             // 9
  'PHONE NUMBER',                              // 10
  'FULL DESCRIPTION OF THE MATTER',            // 11
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
function s(row: (unknown)[], i: number): string {
  return String(row[i] ?? '').trim();
}

function parseDdMmYyyy(d: string): Date | null {
  if (!d || !d.includes('/')) return null;
  const [dd, mm, yyyy] = d.split('/');
  const dt = new Date(+yyyy, +mm - 1, +dd);
  return isNaN(dt.getTime()) ? null : dt;
}

function latestDateLabel(dates: string[]): string {
  const valid = dates
    .map(parseDdMmYyyy)
    .filter((d): d is Date => d !== null)
    .sort((a, b) => b.getTime() - a.getTime());
  if (!valid.length) {
    const t = new Date();
    return `${String(t.getDate()).padStart(2, '0')}-${String(t.getMonth() + 1).padStart(2, '0')}-${t.getFullYear()}`;
  }
  const d = valid[0];
  return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
}

/**
 * Calculate full days between two dates (ignoring time component).
 * Returns null if either date is invalid.
 */
function daysBetween(from: Date, to: Date): number | null {
  const msPerDay = 86_400_000;
  const fromDay = new Date(from.getFullYear(), from.getMonth(), from.getDate()).getTime();
  const toDay   = new Date(to.getFullYear(), to.getMonth(), to.getDate()).getTime();
  const diff = Math.floor((toDay - fromDay) / msPerDay);
  return diff >= 0 ? diff : null;
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generateServiceCallReport(
  fileBuffer: Buffer,
  brand: string,
): Promise<{ buffer: Buffer; filename: string; rawDates: string[] }> {

  // ── 1. Parse raw Perigee export ────────────────────────────────────────────
  const inputWb = XLSX.read(fileBuffer, { type: 'buffer' });
  const inputWs = inputWb.Sheets[inputWb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(inputWs, { header: 1, defval: null }) as unknown[][];

  if (rawData.length < 2) throw new Error('No data rows found in uploaded file.');

  // Row 0 = header labels; data starts at Row 1 (no numeric prefix row)
  const dataRows = rawData.slice(1).filter(r => s(r, 2) || s(r, 3)); // skip blank rows

  if (!dataRows.length) throw new Error('No service call records found in uploaded file.');

  // ── 2. Map rows to typed objects ───────────────────────────────────────────
  const now = new Date();

  interface ServiceCallRow {
    repName:     string;
    date:        string;
    place:       string;
    logType:     string;
    logNumber:   string;
    product:     string;
    loggedDate:  string;
    daysSince:   number | null;
    customer:    string;
    phone:       string;
    description: string;
  }

  const rows: ServiceCallRow[] = dataRows.map(r => {
    const loggedDateStr = s(r, 18);
    const loggedDate    = parseDdMmYyyy(loggedDateStr);
    return {
      repName:     [s(r, 2), s(r, 3)].filter(Boolean).join(' ') || 'UNKNOWN',
      date:        s(r, 9),
      place:       s(r, 6) || 'UNKNOWN',
      logType:     s(r, 15),
      logNumber:   s(r, 16),
      product:     s(r, 17),
      loggedDate:  loggedDateStr,
      daysSince:   loggedDate ? daysBetween(loggedDate, now) : null,
      customer:    s(r, 19),
      phone:       s(r, 20),
      description: s(r, 21),
    };
  });

  const rawDates  = rows.map(r => r.date).filter(d => d.includes('/'));
  const dateLabel = latestDateLabel(rawDates);
  const sheetName = `${brand.toUpperCase()} SERVICE CALLS`;

  // ── 3. Filename ────────────────────────────────────────────────────────────
  const filename = `${brand.toUpperCase()}_SERVICE_CALL_REPORT_${dateLabel}.xlsx`;

  // ── 4. Build output workbook ───────────────────────────────────────────────
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Defy Field Execute';
  wb.created = new Date();

  // Load logos
  const pub = path.join(process.cwd(), 'public');
  let defyLogoId:    number | null = null;
  let atomicLogoId:  number | null = null;
  let perigeeLogoId: number | null = null;

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    defyLogoId    = wb.addImage({ buffer: fs.readFileSync(path.join(pub, 'defy-logo.png'))    as any, extension: 'png'  });
  } catch { /* unavailable */ }
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    atomicLogoId  = wb.addImage({ buffer: fs.readFileSync(path.join(pub, 'atomic-logo.png'))  as any, extension: 'png'  });
  } catch { /* unavailable */ }
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    perigeeLogoId = wb.addImage({ buffer: fs.readFileSync(path.join(pub, 'perigee-logo.jpg')) as any, extension: 'jpeg' });
  } catch { /* unavailable */ }

  // ── MENU sheet ─────────────────────────────────────────────────────────────
  const genDate = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' });
  buildMenuSheet(wb, {
    title:    `${brand.toUpperCase()} SERVICE CALL REPORT`,
    subtitle: `Generated: ${genDate}`,
    sheetDefs: [{
      name: sheetName,
      desc: `${rows.length} service call${rows.length === 1 ? '' : 's'}`,
    }],
    defyLogoId,
    atomicLogoId,
    perigeeLogoId,
  });

  // ── DATA sheet ─────────────────────────────────────────────────────────────
  const ws = wb.addWorksheet(sheetName);
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2, topLeftCell: 'A3', showGridLines: false }];

  addNavRow(ws, `${brand.toUpperCase()} SERVICE CALL REPORT`, OUT_COLS.length);

  // Header row (ws row 2)
  const hRow = ws.addRow(OUT_COLS);
  hRow.height = 36;
  hRow.eachCell(cell => applyHeaderStyle(cell));
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: OUT_COLS.length } };

  // Data rows
  rows.forEach((row, i) => {
    const r = ws.addRow([
      row.repName,
      row.date,
      row.place,
      row.logType,
      row.logNumber,
      row.product,
      row.loggedDate,
      row.daysSince ?? '',
      row.customer,
      row.phone,
      row.description,
    ]);

    r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));

    // Style "DAYS SINCE" column as number
    const daysCell = r.getCell(8);
    if (typeof row.daysSince === 'number') {
      daysCell.numFmt = '0';
    }

    // Wrap the description column
    r.getCell(11).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

    r.height = 22;
  });

  // Column widths
  [22, 14, 26, 22, 16, 30, 14, 10, 20, 16, 40]
    .forEach((w, idx) => { ws.getColumn(idx + 1).width = w; });

  const buffer = Buffer.from(await wb.xlsx.writeBuffer());
  return { buffer, filename, rawDates };
}

/**
 * Defy Red Flag Report builder
 *
 * Input:  Perigee raw export (single "Worksheet" sheet, 19 columns A–S)
 * Output: Excel workbook — MENU + one sheet per escalation category
 *
 * Column mapping (0-indexed):
 *   0  A  ID
 *   1  B  Email
 *   2  C  First Name
 *   3  D  Last Name
 *   4  E  Customer
 *   5  F  Channel
 *   6  G  Store
 *   7  H  Store Code
 *   8  I  Province
 *   9  J  Date (DD/MM/YYYY)
 *  10  K  Time
 *  11  L  Tag
 *  12  M  Sync Date
 *  13  N  Sync Time
 *  14  O  DEFY RED FLAG REPORT  ← category (shown as output column)
 *  15  P  WHAT IS THE PROBLEM?  ← used for sheet grouping
 *  16  Q  WHAT IS THE MODEL NUMBER
 *  17  R  TAKE A PICTURE OF THE PROBLEM  ← Perigee image URL
 *  18  S  DO YOU WANT TO ESCALATE TO MANAGEMENT
 *
 * Image handling:
 *   Images are fetched from SharePoint before building the Excel file.
 *   The VBA saves each image as "{perigee-id}.jpg" where {perigee-id} is the
 *   last path segment of the Perigee image URL
 *   (e.g. "perigee-1307447LTnKUCFZBrdtVk").
 *
 *   SP folder path:  DFE_PICTURES_SP_PATH env var
 *   Default:         {DFE_SP_BASE_PATH}/PERIGEE IMAGE DOWNLOADS
 *
 *   If the file is not found on SP the cell keeps its clickable URL hyperlink.
 */
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import fs   from 'fs';
import path from 'path';
import { buildMenuSheet, applyHeaderStyle, applyDataStyle, addNavRow } from './build-menu';
import { downloadFileFromSP } from '@/lib/graph-oj';
import { loadAppSettings } from '@/lib/appSettings';

// ─── Constants ────────────────────────────────────────────────────────────────
const DEFY_RED = 'E31837';

// Output column headers
const OUT_COLS = [
  'REPRESENTATIVE NAME',   // 1
  'DATE',                  // 2
  'STORE',                 // 3
  'CATEGORY',              // 4
  'WHAT IS THE PROBLEM?',  // 5
  'MODEL NUMBER',          // 6
  'PICTURE',               // 7 ← image embedded here
  'ESCALATE?',             // 8
];
const PIC_COL     = 7;   // 1-indexed column for the picture
const PIC_COL_0   = 6;   // 0-indexed (for addImage)
const IMAGE_W     = 110; // px
const IMAGE_H     = 85;  // px
const ROW_H_IMAGE = 68;  // pts (accommodates IMAGE_H with small margin)
const ROW_H_PLAIN = 20;  // pts

// ─── Types ────────────────────────────────────────────────────────────────────
interface RedFlagRow {
  repName:  string;
  store:    string;
  date:     string;
  category: string;
  problem:  string;
  model:    string;
  imageUrl: string;
  imageId:  string;  // last URL segment = filename (without extension)
  escalate: string;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function safeSheet(name: string): string {
  return name.replace(/[/\\*?:[\]]/g, '').trim().slice(0, 31) || 'OTHER';
}

/**
 * Convert a Perigee image URL to the filename used by the VBA downloader.
 * Windows can't use "/" or ":" in filenames, so the VBA strips both characters.
 *
 * e.g. "https://live.perigeeportal.co.za/.../perigee-xxx/NONE/NONE"
 *   →  "httpslive.perigeeportal.co.za...perigee-xxxNONENONE"     (.jpg added by caller)
 */
function urlToFilename(url: string): string {
  if (!url) return '';
  return url.replace(/[/:]/g, '');
}

// ─── Quick problem extractor (used by generate route for validation) ──────────
export function extractRedFlagProblems(fileBuffer: Buffer): Set<string> {
  const inputWb = XLSX.read(fileBuffer, { type: 'buffer' });
  const inputWs = inputWb.Sheets[inputWb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(inputWs, { header: 1, defval: null }) as (string | null)[][];
  return new Set(
    rawData.slice(1).map(row => String(row[15] ?? '').trim()).filter(Boolean),
  );
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generateRedFlag(
  fileBuffer:    Buffer,
  brand:         string,
  label:         'SALES' | 'MARKETING',
  problemFilter: string[],   // only rows whose problem is in this list
): Promise<{ buffer: Buffer; filename: string; rawDates: string[] }> {

  // ── 1. Parse raw Perigee export ────────────────────────────────────────────
  const inputWb = XLSX.read(fileBuffer, { type: 'buffer' });
  const inputWs = inputWb.Sheets[inputWb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(inputWs, { header: 1, defval: null }) as (string | number | null)[][];

  if (rawData.length < 2) throw new Error('No data rows found in uploaded file.');

  const rawRows = rawData.slice(1);

  // ── 2. Map rows to typed objects ───────────────────────────────────────────
  const rows: RedFlagRow[] = rawRows.map(row => {
    const imageUrl = String(row[17] ?? '').trim();
    return {
      repName:  [String(row[2] ?? '').trim(), String(row[3] ?? '').trim()].filter(Boolean).join(' ') || 'UNKNOWN',
      store:    String(row[6] ?? '').trim() || String(row[7] ?? '').trim() || 'UNKNOWN',
      date:     String(row[9] ?? '').trim(),
      category: String(row[14] ?? '').trim(),
      problem:  String(row[15] ?? '').trim(),
      model:    String(row[16] ?? '').trim(),
      imageUrl,
      imageId:  urlToFilename(imageUrl),
      escalate: String(row[18] ?? '').trim(),
    };
  });

  if (rows.length === 0) throw new Error('No data rows found in uploaded file.');

  // Apply problem filter
  const filteredRows = problemFilter.length
    ? rows.filter(r => problemFilter.includes(r.problem))
    : rows;

  if (filteredRows.length === 0) {
    throw new Error(`No rows match the selected problems for the ${label} report.`);
  }

  const rawDates = filteredRows.map(r => r.date).filter(d => d.includes('/'));

  // ── 3. Pre-fetch images from SharePoint ────────────────────────────────────
  // Images saved by VBA as "{imageId}.jpg" in the PERIGEE IMAGE DOWNLOADS folder.
  const BASE_PATH      = (process.env.DFE_SP_BASE_PATH   || 'DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS').trim();
  const appSettings    = loadAppSettings();
  const PICTURES_FOLDER = (
    appSettings.picturesFolderPath ||
    process.env.DFE_PICTURES_SP_PATH ||
    `${BASE_PATH}/PERIGEE IMAGE DOWNLOADS`
  ).trim();

  // Collect unique image IDs across filtered rows
  const uniqueIds = [...new Set(filteredRows.map(r => r.imageId).filter(Boolean))];

  // Fetch all in parallel; missing/error → null
  const fetched = await Promise.all(
    uniqueIds.map(id => downloadFileFromSP(PICTURES_FOLDER, `${id}.jpg`))
  );
  const imageBuffers = new Map<string, Buffer>();
  uniqueIds.forEach((id, idx) => {
    const buf = fetched[idx];
    if (buf) imageBuffers.set(id, buf);
  });

  console.log(`[red-flag/${label}] SP images: ${imageBuffers.size}/${uniqueIds.length} fetched from "${PICTURES_FOLDER}"`);

  // ── 4. Group rows by "WHAT IS THE PROBLEM?" ───────────────────────────────
  const grouped = new Map<string, RedFlagRow[]>();
  for (const row of filteredRows) {
    const key = safeSheet(row.problem.toUpperCase()) || 'OTHER';
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key)!.push(row);
  }
  const problems = [...grouped.keys()].sort();

  // ── 5. Filename ────────────────────────────────────────────────────────────
  const today    = new Date();
  const dd       = String(today.getDate()).padStart(2, '0');
  const mm       = String(today.getMonth() + 1).padStart(2, '0');
  const yyyy     = today.getFullYear();
  const filename = `${brand.toUpperCase()}_${label}_RED_FLAG_${yyyy}${mm}${dd}.xlsx`;

  // ── 6. Build output workbook ───────────────────────────────────────────────
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
  const genDate   = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' });
  const sheetDefs = problems.map(prob => {
    const probRows  = grouped.get(prob)!;
    const escalated = probRows.filter(r => /^yes$/i.test(r.escalate)).length;
    const desc      = `${probRows.length} issue${probRows.length === 1 ? '' : 's'}${escalated > 0 ? ` · ${escalated} escalated to management` : ''}`;
    return { name: prob, desc };
  });

  buildMenuSheet(wb, {
    title:        `${brand.toUpperCase()} — ${label} RED FLAG REPORT`,
    subtitle:     `Generated: ${genDate}`,
    sheetDefs,
    defyLogoId,
    atomicLogoId,
    perigeeLogoId,
  });

  // ── Problem sheets ─────────────────────────────────────────────────────────
  for (const prob of problems) {
    const catRows = grouped.get(prob)!;
    const ws      = wb.addWorksheet(prob);
    ws.views      = [{ state: 'frozen', xSplit: 0, ySplit: 2, topLeftCell: 'A3', showGridLines: false }];

    addNavRow(ws, prob, OUT_COLS.length);

    // Header row (ws row 2)
    const hRow = ws.addRow(OUT_COLS);
    hRow.height = 28;
    hRow.eachCell(cell => applyHeaderStyle(cell));
    ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: OUT_COLS.length } };

    // Data rows
    // Image row index (0-indexed):  nav=0, header=1, first data=2, data_i = i+2
    catRows.forEach((row, i) => {
      const r = ws.addRow([
        row.repName,
        row.date,
        row.store,
        row.category,
        row.problem,
        row.model,
        '',           // PICTURE — set below
        row.escalate,
      ]);

      r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));

      // Always write a clickable hyperlink in the PICTURE cell
      if (row.imageUrl) {
        const picCell = r.getCell(PIC_COL);
        picCell.value = { text: '🔗 View', hyperlink: row.imageUrl, tooltip: 'View image on Perigee portal' };
        picCell.font  = { color: { argb: '0563C1' }, underline: true, size: 9, name: 'Arial' };
        picCell.alignment = { vertical: 'bottom', horizontal: 'left' };
      }

      // Embed image if we fetched it from SP
      const imgBuf = imageBuffers.get(row.imageId);
      if (imgBuf) {
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const imgId = wb.addImage({ buffer: imgBuf as any, extension: 'jpeg' });
          ws.addImage(imgId, {
            tl:  { col: PIC_COL_0, row: i + 2 + 0.05 },
            ext: { width: IMAGE_W, height: IMAGE_H },
          });
          r.height = ROW_H_IMAGE;
        } catch {
          r.height = ROW_H_PLAIN;
        }
      } else {
        r.height = ROW_H_PLAIN;
      }

      // Highlight escalated rows
      if (/^yes$/i.test(row.escalate)) {
        const escalCell = r.getCell(8);
        escalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5E5' } };
        escalCell.font = { bold: true, color: { argb: DEFY_RED }, size: 10, name: 'Arial' };
      }
    });

    // Column widths
    [22, 12, 28, 20, 35, 16, 18, 14]
      .forEach((w, idx) => { ws.getColumn(idx + 1).width = w; });
  }

  const buffer = Buffer.from(await wb.xlsx.writeBuffer());
  return { buffer, filename, rawDates };
}

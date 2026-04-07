/**
 * Activation Report builder
 *
 * Input:  Perigee raw export ("Worksheet" sheet, 30 cols)
 * Output: Excel workbook — MENU + ACTIVATIONS data sheet
 *
 * Column mapping (0-indexed, data rows start at raw index 2):
 *   Row 0  numeric prefix row (skip)
 *   Row 1  header row         (skip)
 *   Row 2+ data
 *
 *   2   First Name
 *   3   Last Name
 *   5   Channel
 *   6   Store (Place)
 *   9   Date (DD/MM/YYYY)
 *  11   Visit UUID  ← added by Perigee, shifts everything below by +1
 *  15   Activation Start date?
 *  16   Activation End date?
 *  17   Type of Activation?
 *  18   Image URL 1 (Activation Stand 1)
 *  19   Image URL 2 (Activation Stand 2)
 *  20   Image URL 3 (Activation Stand 3)
 *  21   Image URL 4 (Activation Stand 4)  — future Perigee expansion
 *  22   Image URL 5 (Activation Stand 5)  — future Perigee expansion
 *
 * Image handling:
 *   VBA saves each image as "{perigee-id}.jpg" where {perigee-id} is the
 *   last meaningful path segment of the Perigee image URL
 *   (e.g. "perigee-1307447LTnKUCFZBrdtVk").
 *
 *   SP folder path: DFE_ACTIVATION_PICTURES_SP_PATH env var
 *   Default:        {DFE_SP_BASE_PATH}/ACTIVATION IMAGE DOWNLOADS
 *
 *   If the file is not found on SP the cell keeps its clickable URL hyperlink.
 */
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import fs   from 'fs';
import path from 'path';
import { buildMenuSheet, applyHeaderStyle, applyDataStyle, addNavRow } from './build-menu';
import { listFilesInSPFolder, downloadSPFileById } from '@/lib/graph-oj';

// ─── Constants ────────────────────────────────────────────────────────────────
const SHEET_NAME = 'ACTIVATIONS';
const MAX_IMAGES = 5;

const OUT_COLS = [
  'REPRESENTATIVE NAME',   // 1
  'DATE',                  // 2
  'PLACE',                 // 3
  'ACTIVATION START DATE', // 4
  'ACTIVATION END DATE',   // 5
  'TYPE OF ACTIVATION',    // 6
  'ACTIVATION STAND 1',    // 7  ← image
  'ACTIVATION STAND 2',    // 8  ← image
  'ACTIVATION STAND 3',    // 9  ← image
  'ACTIVATION STAND 4',    // 10 ← image
  'ACTIVATION STAND 5',    // 11 ← image
  'CHANNEL',               // 12
];

// Image columns (1-indexed) and their 0-indexed equivalents for addImage
const IMG_COLS  = [7, 8, 9, 10, 11];          // 1-indexed
const IMG_COLS0 = [6, 7, 8,  9, 10];          // 0-indexed
const IMAGE_W     = 120; // px
const IMAGE_H     = 90;  // px
const ROW_H_IMAGE = 72;  // pts (accommodates IMAGE_H with small margin)
const ROW_H_PLAIN = 20;  // pts

// Raw data column indices for image URLs (0-indexed)
const RAW_IMAGE_COLS = [18, 19, 20, 21, 22];

// ─── Helpers ──────────────────────────────────────────────────────────────────
function s(row: (unknown)[], i: number): string {
  return String(row[i] ?? '').trim();
}

/**
 * Extract the Perigee image ID from a URL — the segment starting with "perigee-".
 * This matches the VBA ImageFilename() convention which extracts the token.
 *
 * e.g. "https://live.perigeeportal.co.za/.../perigee-1307447LTnKUCFZBrdtVk/NONE/NONE"
 *   →  "perigee-1307447LTnKUCFZBrdtVk"
 */
function urlToImageId(url: string): string {
  if (!url) return '';
  const seg = url.split('/').find(s => s.toLowerCase().startsWith('perigee-'));
  return seg ?? '';
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
    return `${String(t.getDate()).padStart(2,'0')}-${String(t.getMonth()+1).padStart(2,'0')}-${t.getFullYear()}`;
  }
  const d = valid[0];
  return `${String(d.getDate()).padStart(2,'0')}-${String(d.getMonth()+1).padStart(2,'0')}-${d.getFullYear()}`;
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generateActivationReport(
  fileBuffer: Buffer,
  brand: string,
): Promise<{ buffer: Buffer; filename: string; rawDates: string[] }> {

  // ── 1. Parse raw Perigee export ────────────────────────────────────────────
  const inputWb  = XLSX.read(fileBuffer, { type: 'buffer' });
  const inputWs  = inputWb.Sheets[inputWb.SheetNames[0]];
  const rawData  = XLSX.utils.sheet_to_json(inputWs, { header: 1, defval: null }) as unknown[][];

  if (rawData.length < 3) throw new Error('No data rows found in uploaded file.');

  // Row 0 = numeric prefix; Row 1 = header labels; data starts at Row 2
  const dataRows = rawData.slice(2).filter(r => s(r, 2) || s(r, 3)); // skip blank rows

  if (!dataRows.length) throw new Error('No activation records found in uploaded file.');

  // ── 2. Map rows to typed objects ───────────────────────────────────────────
  interface ActivationRow {
    repName:    string;
    date:       string;
    place:      string;
    startDate:  string;
    endDate:    string;
    type:       string;
    imageUrls:  string[];      // up to MAX_IMAGES
    imageIds:   string[];      // corresponding VBA filenames (no ext)
    channel:    string;
  }

  const rows: ActivationRow[] = dataRows.map(r => {
    const imageUrls = RAW_IMAGE_COLS
      .map(colIdx => s(r, colIdx))
      .filter(v => v.startsWith('http'));
    return {
      repName:   [s(r, 2), s(r, 3)].filter(Boolean).join(' ') || 'UNKNOWN',
      date:      s(r, 9),
      place:     s(r, 6) || s(r, 7) || 'UNKNOWN',
      startDate: s(r, 15),
      endDate:   s(r, 16),
      type:      s(r, 17),
      imageUrls,
      imageIds:  imageUrls.map(urlToImageId),
      channel:   s(r, 5),
    };
  });

  const rawDates  = rows.map(r => r.date).filter(d => d.includes('/'));
  const dateLabel = latestDateLabel(rawDates);

  // ── 3. Pre-fetch images from SharePoint ────────────────────────────────────
  const BASE_PATH       = (process.env.DFE_SP_BASE_PATH || 'DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS').trim();
  const PICTURES_FOLDER = (
    process.env.DFE_ACTIVATION_PICTURES_SP_PATH ||
    `${BASE_PATH}/ACTIVATION IMAGE DOWNLOADS`
  ).trim();

  // Collect unique non-empty image IDs
  const uniqueIds = [...new Set(rows.flatMap(r => r.imageIds).filter(Boolean))];

  // Recursively index all .jpg files in the pictures folder (including subfolders
  // created by older VBA versions) then fetch only the ones we need.
  const spFiles = await listFilesInSPFolder(PICTURES_FOLDER);
  console.log(
    `[activation-report] SP file index: ${spFiles.size} file(s) in "${PICTURES_FOLDER}" (incl. subfolders)`
  );

  // Diagnostic: log samples so we can verify naming matches
  const spSample = [...spFiles.keys()].slice(0, 5);
  if (spSample.length) console.log(`[activation-report] SP sample filenames:`, spSample);
  const idSample = uniqueIds.slice(0, 3).map(id => `${id}.jpg`.toLowerCase());
  if (idSample.length) console.log(`[activation-report] Expected image filenames:`, idSample);

  const fetched = await Promise.all(
    uniqueIds.map(async id => {
      const key    = `${id}.jpg`.toLowerCase();
      const itemId = spFiles.get(key);
      if (!itemId) return null;
      return downloadSPFileById(itemId);
    })
  );
  const imageBuffers = new Map<string, Buffer>();
  uniqueIds.forEach((id, idx) => {
    const buf = fetched[idx];
    if (buf) imageBuffers.set(id, buf);
  });

  console.log(
    `[activation-report] SP images: ${imageBuffers.size}/${uniqueIds.length} fetched`
  );

  // ── 4. Filename ────────────────────────────────────────────────────────────
  const filename = `${brand.toUpperCase()}_ACTIVATION_REPORT_${dateLabel}.xlsx`;

  // ── 5. Build output workbook ───────────────────────────────────────────────
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
    title:    `${brand.toUpperCase()} ACTIVATION REPORT`,
    subtitle: `Generated: ${genDate}`,
    sheetDefs: [{
      name: SHEET_NAME,
      desc: `${rows.length} activation${rows.length === 1 ? '' : 's'} — ${rawDates.length} with dates`,
    }],
    defyLogoId,
    atomicLogoId,
    perigeeLogoId,
  });

  // ── ACTIVATIONS sheet ──────────────────────────────────────────────────────
  const ws = wb.addWorksheet(SHEET_NAME);
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2, topLeftCell: 'A3', showGridLines: false }];

  addNavRow(ws, `${brand.toUpperCase()} ACTIVATION REPORT`, OUT_COLS.length);

  // Header row (ws row 2)
  const hRow = ws.addRow(OUT_COLS);
  hRow.height = 28;
  hRow.eachCell(cell => applyHeaderStyle(cell));
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: OUT_COLS.length } };

  // Data rows — ws row index = i + 2 (0-indexed), nav=0, header=1, first data=2
  rows.forEach((row, i) => {
    const r = ws.addRow([
      row.repName,
      row.date,
      row.place,
      row.startDate,
      row.endDate,
      row.type,
      '',   // STAND 1 — set below
      '',   // STAND 2 — set below
      '',   // STAND 3 — set below
      '',   // STAND 4 — set below
      '',   // STAND 5 — set below
      row.channel,
    ]);

    r.eachCell(cell => applyDataStyle(cell, i % 2 === 1));

    let anyImageEmbedded = false;

    // Handle each image slot (up to MAX_IMAGES)
    row.imageUrls.forEach((url, slot) => {
      if (!url || slot >= MAX_IMAGES) return;
      const colIdx1 = IMG_COLS[slot];   // 1-indexed
      const colIdx0 = IMG_COLS0[slot];  // 0-indexed for addImage

      const picCell = r.getCell(colIdx1);

      // Embed from SP if available — with clickable hyperlink on the image
      const imageId = row.imageIds[slot];
      const imgBuf  = imageBuffers.get(imageId);
      if (imgBuf) {
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const imgId = wb.addImage({ buffer: imgBuf as any, extension: 'jpeg' });
          ws.addImage(imgId, {
            tl:  { col: colIdx0, row: i + 2 + 0.05 },
            ext: { width: IMAGE_W, height: IMAGE_H },
            hyperlinks: { hyperlink: url, tooltip: 'Click to view full image' },
          });
          anyImageEmbedded = true;
        } catch { /* skip */ }
      }

      // Fallback: if image not embedded, show a clickable text link
      if (!imgBuf) {
        picCell.value = { text: '\u{1F517} View', hyperlink: url, tooltip: 'View image on Perigee portal' };
        picCell.font  = { color: { argb: '0563C1' }, underline: true, size: 9, name: 'Arial' };
        picCell.alignment = { vertical: 'bottom', horizontal: 'left' };
      }
    });

    r.height = anyImageEmbedded ? ROW_H_IMAGE : ROW_H_PLAIN;
  });

  // Column widths: [rep, date, place, start, end, type, img1-5, channel]
  [24, 14, 28, 20, 20, 28, 18, 18, 18, 18, 18, 18]
    .forEach((w, idx) => { ws.getColumn(idx + 1).width = w; });

  const buffer = Buffer.from(await wb.xlsx.writeBuffer());
  return { buffer, filename, rawDates };
}

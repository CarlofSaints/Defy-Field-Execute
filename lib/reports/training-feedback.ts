/**
 * Training Feedback Report builder
 *
 * Input:  Perigee raw export ("Worksheet" sheet, 55 cols)
 * Output: Multi-sheet Excel — MENU / OVERALL / CHANNEL VIEW / TRAINER SUMMARY
 *
 * Column mapping (0-indexed):
 *   2   First Name
 *   3   Last Name
 *   5   Channel
 *   6   Store (PLACE)
 *   8   Province (REGION)
 *   9   Date (DD/MM/YYYY)
 *  10   Time (HH:MM)
 *  11   Visit UUID
 *  16   HOW MANY FSPs DID YOU TRAIN? (1-30)
 *  18-52 Category questions + optional follow-ups (see CATEGORIES below)
 *  53   HOW LONG WAS THE TRAINING (IN MINUTES)
 *  54   Image URL (ignored)
 */
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import fs   from 'fs';
import path from 'path';
import { buildMenuSheet, addNavRow } from './build-menu';

// ─── Category definitions ─────────────────────────────────────────────────────
// yesIdx    = 0-based column index of the YES/NO question
// detailIdx = 0-based index of the "PLEASE SPECIFY" follow-up (null = not in form)
const CATEGORIES = [
  { name: 'AIR CONS',          yesIdx: 18, detailIdx: null },
  { name: 'CHEST FREEZERS',    yesIdx: 19, detailIdx: 20   },
  { name: 'COFFEE MACHINES',   yesIdx: 21, detailIdx: null },
  { name: 'COOKER HOODS',      yesIdx: 22, detailIdx: 23   },
  { name: 'DISHWASHERS',       yesIdx: 24, detailIdx: 25   },
  { name: 'DRYERS',            yesIdx: 26, detailIdx: 27   },
  { name: 'FANS',              yesIdx: 28, detailIdx: null },
  { name: 'FRIDGE',            yesIdx: 29, detailIdx: 30   },
  { name: 'FRONT LOADERS',     yesIdx: 31, detailIdx: 32   },
  { name: 'HAND BLENDERS',     yesIdx: 33, detailIdx: null },
  { name: 'HOBS',              yesIdx: 34, detailIdx: 35   },
  { name: 'IRONS',             yesIdx: 36, detailIdx: null },
  { name: 'JUICERS',           yesIdx: 37, detailIdx: null },
  { name: 'KETTLES',           yesIdx: 38, detailIdx: null },
  { name: 'MICROWAVES',        yesIdx: 39, detailIdx: 40   },
  { name: 'OVENS',             yesIdx: 41, detailIdx: 42   },
  { name: 'PRESSURE COOKERS',  yesIdx: 43, detailIdx: null },
  { name: 'SLOW COOKERS',      yesIdx: 44, detailIdx: null },
  { name: 'STOVES',            yesIdx: 45, detailIdx: 46   },
  { name: 'TABLE BLENDERS',    yesIdx: 47, detailIdx: null },
  { name: 'TOASTERS',          yesIdx: 48, detailIdx: null },
  { name: 'TOP LOADERS',       yesIdx: 49, detailIdx: 50   },
  { name: 'TWIN TUBS',         yesIdx: 51, detailIdx: 52   },
] as const;

// ─── Styling ──────────────────────────────────────────────────────────────────
const HEADER_FILL: ExcelJS.Fill = {
  type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF808080' },
};
const HEADER_FONT: Partial<ExcelJS.Font> = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
const TOTAL_FONT:  Partial<ExcelJS.Font> = { bold: true, name: 'Arial', size: 10 };
const DATA_FONT:   Partial<ExcelJS.Font> = { name: 'Arial', size: 10 };

// ─── Helpers ─────────────────────────────────────────────────────────────────
function s(row: (string | number | null)[], i: number): string {
  return String(row[i] ?? '').trim();
}

function parseDdMmYyyy(d: string): Date | null {
  if (!d || !d.includes('/')) return null;
  const [dd, mm, yyyy] = d.split('/');
  const dt = new Date(+yyyy, +mm - 1, +dd);
  return isNaN(dt.getTime()) ? null : dt;
}

/**
 * Combine DD/MM/YYYY date + HH:MM time → epoch ms (or null if invalid).
 * Used to detect submission timing anomalies.
 */
function parseDateTimeMs(dateStr: string, timeStr: string): number | null {
  if (!dateStr || !dateStr.includes('/')) return null;
  if (!timeStr || !timeStr.includes(':')) return null;
  const [dd, mm, yyyy] = dateStr.split('/').map(Number);
  const [hh, mi]       = timeStr.split(':').map(Number);
  const dt = new Date(yyyy, mm - 1, dd, hh, mi);
  return isNaN(dt.getTime()) ? null : dt.getTime();
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

function applyHeader(row: ExcelJS.Row) {
  row.eachCell(cell => {
    cell.font      = HEADER_FONT;
    cell.fill      = HEADER_FILL;
    cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
  });
  row.height = 30;
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generateTrainingFeedback(
  fileBuffer: Buffer,
  brand: string,
): Promise<{ buffer: Buffer; filename: string; rawDates: string[] }> {

  // ── 1. Parse raw Perigee Excel ─────────────────────────────────────────────
  const wb  = XLSX.read(fileBuffer, { type: 'buffer' });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (unknown)[][];

  if (raw.length < 2) throw new Error('No data rows found in uploaded file.');

  // ── 2. Map rows ────────────────────────────────────────────────────────────
  const rows = raw.slice(1).map(r => {
    const a = r as (string | number | null)[];

    const trainedOn = CATEGORIES.map(c => s(a, c.yesIdx));
    const details   = CATEGORIES.map(c =>
      c.detailIdx != null ? s(a, c.detailIdx) : '',
    );
    const noOfTrainings = trainedOn.filter(v => v.toLowerCase() === 'yes').length;

    const date = s(a, 9);
    const time = s(a, 10);
    const fspsRaw = a[16];
    const durRaw  = a[53];

    return {
      trainerName:  [s(a, 2), s(a, 3)].filter(Boolean).join(' '),
      date,
      time,
      store:        s(a, 6),
      channel:      s(a, 5),
      province:     s(a, 8),
      visitUuid:    s(a, 11),
      fsps:         fspsRaw != null && fspsRaw !== '' ? Number(fspsRaw) : null as number | null,
      trainedOn,
      details,
      noOfTrainings,
      durationMins: durRaw != null && durRaw !== '' ? Number(durRaw) : null as number | null,
      dateTimeMs:   parseDateTimeMs(date, time),
    };
  }).filter(r => r.trainerName);

  if (!rows.length) throw new Error('No valid training records found in uploaded file.');

  const rawDates  = rows.map(r => r.date).filter(d => d.includes('/'));
  const dateLabel = latestDateLabel(rawDates);
  const genDate   = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' });

  // ── 3. Build workbook ──────────────────────────────────────────────────────
  const ewb = new ExcelJS.Workbook();
  ewb.creator = 'Defy Field Execute';
  ewb.created = new Date();

  // Load logos (non-fatal if unavailable)
  const pub = path.join(process.cwd(), 'public');
  let defyLogoId:    number | null = null;
  let atomicLogoId:  number | null = null;
  let perigeeLogoId: number | null = null;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    defyLogoId    = ewb.addImage({ buffer: fs.readFileSync(path.join(pub, 'defy-logo.png'))    as any, extension: 'png'  });
  } catch { /* unavailable */ }
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    atomicLogoId  = ewb.addImage({ buffer: fs.readFileSync(path.join(pub, 'atomic-logo.png'))  as any, extension: 'png'  });
  } catch { /* unavailable */ }
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    perigeeLogoId = ewb.addImage({ buffer: fs.readFileSync(path.join(pub, 'perigee-logo.jpg')) as any, extension: 'jpeg' });
  } catch { /* unavailable */ }

  // ── MENU sheet (added first so it's the opening tab) ──────────────────────
  const totalVisitsAll  = rows.length;
  const totalTrainingsAll = rows.reduce((sum, r) => sum + r.noOfTrainings, 0);
  const totalFspsAll    = rows.reduce((sum, r) => sum + (r.fsps ?? 0), 0);

  buildMenuSheet(ewb, {
    title:    `${brand.toUpperCase()} TRAINING FEEDBACK REPORT`,
    subtitle: `Generated ${genDate} · ${totalVisitsAll} visit${totalVisitsAll === 1 ? '' : 's'} · ${totalFspsAll} FSPs trained · ${totalTrainingsAll} category trainings`,
    sheetDefs: [
      { name: 'OVERALL',         desc: 'Full training data — one row per visit, all categories' },
      { name: 'CHANNEL VIEW',    desc: 'Training count summary by channel' },
      { name: 'TRAINER SUMMARY', desc: 'Per-trainer breakdown: visits, category trainings, FSPs' },
    ],
    defyLogoId,
    atomicLogoId,
    perigeeLogoId,
  });

  // ── Col count helpers ─────────────────────────────────────────────────────
  // OVERALL: TRAINER + DATE + PLACE (3) + 1 col per category + 1 col per detail + 4 trailing
  const detailCount = CATEGORIES.filter(c => c.detailIdx != null).length;
  const ovColCount = 3 + CATEGORIES.length + detailCount + 4;
  const cvColCount = 3;
  const tsColCount = 7;

  // ─────────────────────────────────────────────────────────────────────────
  // Sheet 2: OVERALL
  // ─────────────────────────────────────────────────────────────────────────
  const ows = ewb.addWorksheet('OVERALL');
  ows.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];

  // Row 1: nav bar (red strip with sheet title + ⬅ MENU link)
  addNavRow(ows, `${brand.toUpperCase()} TRAINING FEEDBACK — ALL VISITS`, ovColCount);

  // Row 2: column headers — only add detail column when the form actually has one
  const ovHeaders: string[] = ['TRAINER NAME', 'DATE', 'PLACE'];
  for (const cat of CATEGORIES) {
    ovHeaders.push(`TRAINING ON ${cat.name}?`);
    if (cat.detailIdx != null) {
      ovHeaders.push(`PRODUCTS IN THE ${cat.name} CATEGORY YOU PROVIDED TRAINING ON?`);
    }
  }
  ovHeaders.push('NO OF TRAININGS DONE', 'TRAINING (IN MINUTES)', 'CHANNEL', 'REGION');
  applyHeader(ows.addRow(ovHeaders));

  // AutoFilter on header row (row 2)
  ows.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: ovHeaders.length } };

  // Column widths — walk in lock-step with header build
  ows.getColumn(1).width = 25;  // TRAINER NAME
  ows.getColumn(2).width = 14;  // DATE
  ows.getColumn(3).width = 30;  // PLACE
  let colCursor = 4;
  for (const cat of CATEGORIES) {
    ows.getColumn(colCursor).width = 12;  // YES/NO
    colCursor++;
    if (cat.detailIdx != null) {
      ows.getColumn(colCursor).width = 32;  // DETAIL
      colCursor++;
    }
  }
  ows.getColumn(colCursor).width     = 18;  // NO OF TRAININGS DONE
  ows.getColumn(colCursor + 1).width = 16;  // TRAINING (IN MINUTES)
  ows.getColumn(colCursor + 2).width = 18;  // CHANNEL
  ows.getColumn(colCursor + 3).width = 18;  // REGION

  // Data rows
  for (const row of rows) {
    const vals: (string | number | null)[] = [row.trainerName, row.date, row.store];
    for (let i = 0; i < CATEGORIES.length; i++) {
      vals.push(row.trainedOn[i] || null);
      if (CATEGORIES[i].detailIdx != null) {
        vals.push(row.details[i] || null);
      }
    }
    vals.push(row.noOfTrainings, row.durationMins, row.channel || null, row.province || null);
    const dr = ows.addRow(vals);
    dr.eachCell(cell => { cell.font = DATA_FONT; cell.alignment = { vertical: 'middle', wrapText: false }; });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // Sheet 3: CHANNEL VIEW
  // ─────────────────────────────────────────────────────────────────────────
  const cvws = ewb.addWorksheet('CHANNEL VIEW');
  cvws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  addNavRow(cvws, 'CHANNEL VIEW — TRAINING SUMMARY BY CHANNEL', cvColCount);

  cvws.getColumn(1).width = 28;
  cvws.getColumn(2).width = 26;
  cvws.getColumn(3).width = 14;

  applyHeader(cvws.addRow(['CHANNEL', 'SUM NO OF TRAININGS DONE', '% OF VISITS']));

  // Aggregate by channel
  const channelMap = new Map<string, { trainings: number; visits: number }>();
  for (const row of rows) {
    const ch  = row.channel || '(blank)';
    const cur = channelMap.get(ch) ?? { trainings: 0, visits: 0 };
    cur.trainings += row.noOfTrainings;
    cur.visits++;
    channelMap.set(ch, cur);
  }

  let totalTrainings = 0;
  for (const [ch, data] of [...channelMap.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    totalTrainings += data.trainings;
    const pct = totalVisitsAll > 0
      ? `${((data.visits / totalVisitsAll) * 100).toFixed(1)}%`
      : '0.0%';
    const dr = cvws.addRow([ch, data.trainings, pct]);
    dr.eachCell(cell => { cell.font = DATA_FONT; cell.alignment = { vertical: 'middle' }; });
  }

  const cvTotal = cvws.addRow(['Grand Total', totalTrainings, '100%']);
  cvTotal.eachCell(cell => { cell.font = TOTAL_FONT; cell.alignment = { vertical: 'middle' }; });

  // ─────────────────────────────────────────────────────────────────────────
  // Sheet 4: TRAINER SUMMARY
  // ─────────────────────────────────────────────────────────────────────────
  const tsws = ewb.addWorksheet('TRAINER SUMMARY');
  tsws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  addNavRow(tsws, 'TRAINER SUMMARY — PER-REP BREAKDOWN', tsColCount);

  tsws.getColumn(1).width = 28;  // TRAINER NAME
  tsws.getColumn(2).width = 18;  // PROVINCE
  tsws.getColumn(3).width = 30;  // STORE
  tsws.getColumn(4).width = 10;  // VISITS
  tsws.getColumn(5).width = 22;  // NO OF TRAININGS DONE
  tsws.getColumn(6).width = 18;  // TOTAL FSPs TRAINED
  tsws.getColumn(7).width = 60;  // NOTES

  applyHeader(tsws.addRow([
    'TRAINER NAME', 'PROVINCE', 'STORE',
    'VISITS', 'NO OF TRAININGS DONE', 'TOTAL FSPs TRAINED', 'NOTES',
  ]));

  // Group by trainer + province + store
  const trainerMap = new Map<string, {
    name: string; province: string; store: string;
    visitUuids: Set<string>;
    trainings: number; fsps: number;
    submissions: { time: number; duration: number }[];
  }>();

  for (const row of rows) {
    const key = `${row.trainerName}||${row.province}||${row.store}`;
    let cur = trainerMap.get(key);
    if (!cur) {
      cur = {
        name: row.trainerName, province: row.province, store: row.store,
        visitUuids: new Set(),
        trainings: 0, fsps: 0,
        submissions: [],
      };
      trainerMap.set(key, cur);
    }
    if (row.visitUuid) cur.visitUuids.add(row.visitUuid);
    cur.trainings += row.noOfTrainings;
    cur.fsps      += row.fsps ?? 0;
    if (row.dateTimeMs != null && row.durationMins != null && row.durationMins > 0) {
      cur.submissions.push({ time: row.dateTimeMs, duration: row.durationMins });
    }
  }

  // Detect timing anomalies: any consecutive submission gap (in minutes) shorter
  // than the claimed training duration of the earlier form is impossible.
  function hasTimingAnomaly(submissions: { time: number; duration: number }[]): boolean {
    if (submissions.length < 2) return false;
    const sorted = [...submissions].sort((a, b) => a.time - b.time);
    for (let i = 0; i < sorted.length - 1; i++) {
      const gapMin = (sorted[i + 1].time - sorted[i].time) / 60000;
      if (gapMin < sorted[i].duration) return true;
    }
    return false;
  }

  const WARNING_TEXT =
    'WARNING: it is likely that this form was submitted multiple times. Raw data should be checked! Training duration exceeds form submission time!';

  let gtVisits = 0, gtTrainings = 0, gtFsps = 0;
  for (const data of [...trainerMap.values()].sort((a, b) => a.name.localeCompare(b.name))) {
    const visitCount = data.visitUuids.size;
    const note       = hasTimingAnomaly(data.submissions) ? WARNING_TEXT : '';
    const dr = tsws.addRow([
      data.name, data.province, data.store,
      visitCount, data.trainings, data.fsps, note,
    ]);
    dr.eachCell(cell => { cell.font = DATA_FONT; cell.alignment = { vertical: 'middle' }; });
    if (note) {
      const noteCell = dr.getCell(7);
      noteCell.font      = { ...DATA_FONT, color: { argb: 'FFC00000' }, bold: true };
      noteCell.alignment = { vertical: 'middle', wrapText: true };
    }
    gtVisits    += visitCount;
    gtTrainings += data.trainings;
    gtFsps      += data.fsps;
  }

  const tsTotal = tsws.addRow(['Grand Total', '', '', gtVisits, gtTrainings, gtFsps, '']);
  tsTotal.eachCell(cell => { cell.font = TOTAL_FONT; cell.alignment = { vertical: 'middle' }; });

  // ── 4. Filename & buffer ───────────────────────────────────────────────────
  const filename = `${brand.toUpperCase()}_TRAINING_FEEDBACK_REPORT_${dateLabel}.xlsx`;
  const buffer   = Buffer.from(await ewb.xlsx.writeBuffer());

  return { buffer, filename, rawDates };
}

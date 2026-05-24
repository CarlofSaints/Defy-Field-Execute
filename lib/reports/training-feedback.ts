/**
 * Training Feedback Report builder
 *
 * Input:  Perigee raw export ("Worksheet" sheet, 55 cols)
 * Output: Multi-sheet Excel — MENU / OVERALL / CHANNEL VIEW /
 *         TRAINER SUMMARY (MTD) / TRAINER SUMMARY (WTD)
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
 *  15   DID YOU COMPLETE TRAINING? (Yes/No)
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
 * Combine DD/MM/YYYY date + HH:MM time -> epoch ms (or null if invalid).
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

function pad2(n: number) { return String(n).padStart(2, '0'); }

/**
 * Interpolate an orange gradient from light (low ratio) to full orange (high ratio).
 * Returns a 6-char hex string (no # or FF prefix).
 */
function interpolateOrange(ratio: number): string {
  // Full orange: F4A460, light: FCE4C8
  const r = Math.round(252 - (252 - 244) * ratio);
  const g = Math.round(228 - (228 - 164) * ratio);
  const b = Math.round(200 - (200 - 96)  * ratio);
  return [r, g, b].map(c => c.toString(16).padStart(2, '0')).join('');
}

function applyHeader(row: ExcelJS.Row) {
  row.eachCell(cell => {
    cell.font      = HEADER_FONT;
    cell.fill      = HEADER_FILL;
    cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
  });
  row.height = 30;
}

// ─── Timing anomaly detection ─────────────────────────────────────────────────
const WARNING_TEXT =
  'WARNING: it is likely that this form was submitted multiple times. Raw data should be checked! Training duration exceeds form submission time!';

function hasTimingAnomaly(submissions: { time: number; duration: number }[]): boolean {
  if (submissions.length < 2) return false;
  const sorted = [...submissions].sort((a, b) => a.time - b.time);
  for (let i = 0; i < sorted.length - 1; i++) {
    const gapMin = (sorted[i + 1].time - sorted[i].time) / 60000;
    if (gapMin < sorted[i].duration) return true;
  }
  return false;
}

// ─── Lightweight summary row for archiving + MTD accumulation ─────────────────
export interface TrainingSummaryRow {
  trainerName:       string;
  date:              string;
  channel:           string;
  province:          string;
  store:             string;
  visitUuid:         string;
  fsps:              number | null;
  trainingCompleted: boolean;
  noOfTrainings:     number;
}

// ─── Row type ─────────────────────────────────────────────────────────────────
interface TrainingRow {
  trainerName:       string;
  date:              string;
  time:              string;
  store:             string;
  channel:           string;
  province:          string;
  visitUuid:         string;
  trainingCompleted: boolean;
  fsps:              number | null;
  trainedOn:         string[];
  details:           string[];
  noOfTrainings:     number;
  durationMins:      number | null;
  dateTimeMs:        number | null;
}

// ─── KPI computation ──────────────────────────────────────────────────────────
function computeKpis(subset: TrainingRow[]) {
  const channels  = new Set(subset.map(r => r.channel).filter(Boolean));
  const bas       = new Set(subset.map(r => r.trainerName).filter(Boolean));
  const trainings = subset.filter(r => r.trainingCompleted).length;
  const fsps      = subset.reduce((sum, r) => sum + (r.fsps ?? 0), 0);
  const minutes   = subset.reduce((sum, r) => sum + (r.durationMins ?? 0), 0);
  return { channels: channels.size, bas: bas.size, trainings, fsps, minutes };
}

// ─── Trainer Summary sheet builder (reused for MTD & WTD) ─────────────────────
function buildTrainerSummarySheet(
  ewb: ExcelJS.Workbook,
  sheetName: string,
  navTitle: string,
  subset: TrainingRow[],
  colCount: number,
) {
  const ws = ewb.addWorksheet(sheetName);
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  addNavRow(ws, navTitle, colCount);

  ws.getColumn(1).width = 28;  // TRAINER NAME
  ws.getColumn(2).width = 18;  // CHANNEL
  ws.getColumn(3).width = 18;  // PROVINCE
  ws.getColumn(4).width = 30;  // STORE
  ws.getColumn(5).width = 10;  // VISITS
  ws.getColumn(6).width = 22;  // NO OF TRAININGS DONE
  ws.getColumn(7).width = 18;  // TOTAL FSPs TRAINED
  ws.getColumn(8).width = 60;  // NOTES

  applyHeader(ws.addRow([
    'TRAINER NAME', 'CHANNEL', 'PROVINCE', 'STORE',
    'VISITS', 'NO OF TRAININGS DONE', 'TOTAL FSPs TRAINED', 'NOTES',
  ]));

  // AutoFilter on header row (row 2)
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: colCount } };

  // Group by trainer + channel + province + store
  const map = new Map<string, {
    name: string; channel: string; province: string; store: string;
    visitUuids: Set<string>;
    trainings: number; fsps: number;
    submissions: { time: number; duration: number }[];
  }>();

  for (const row of subset) {
    const key = `${row.trainerName}||${row.channel}||${row.province}||${row.store}`;
    let cur = map.get(key);
    if (!cur) {
      cur = {
        name: row.trainerName, channel: row.channel, province: row.province, store: row.store,
        visitUuids: new Set(),
        trainings: 0, fsps: 0,
        submissions: [],
      };
      map.set(key, cur);
    }
    if (row.visitUuid) cur.visitUuids.add(row.visitUuid);
    cur.trainings += row.noOfTrainings;
    cur.fsps      += row.fsps ?? 0;
    if (row.dateTimeMs != null && row.durationMins != null && row.durationMins > 0) {
      cur.submissions.push({ time: row.dateTimeMs, duration: row.durationMins });
    }
  }

  let gtVisits = 0, gtTrainings = 0, gtFsps = 0;
  for (const data of [...map.values()].sort((a, b) =>
    a.name.localeCompare(b.name) || a.channel.localeCompare(b.channel) || a.store.localeCompare(b.store),
  )) {
    const visitCount = data.visitUuids.size;
    const note       = hasTimingAnomaly(data.submissions) ? WARNING_TEXT : '';
    const dr = ws.addRow([
      data.name, data.channel, data.province, data.store,
      visitCount, data.trainings, data.fsps, note,
    ]);
    dr.eachCell(cell => { cell.font = DATA_FONT; cell.alignment = { vertical: 'middle' }; });
    if (note) {
      const noteCell = dr.getCell(8);
      noteCell.font      = { ...DATA_FONT, color: { argb: 'FFC00000' }, bold: true };
      noteCell.alignment = { vertical: 'middle', wrapText: true };
    }
    gtVisits    += visitCount;
    gtTrainings += data.trainings;
    gtFsps      += data.fsps;
  }

  const totalRow = ws.addRow(['Grand Total', '', '', '', gtVisits, gtTrainings, gtFsps, '']);
  totalRow.eachCell(cell => { cell.font = TOTAL_FONT; cell.alignment = { vertical: 'middle' }; });
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generateTrainingFeedback(
  fileBuffer: Buffer,
  brand: string,
  historicalRows?: TrainingSummaryRow[],
): Promise<{ buffer: Buffer; filename: string; rawDates: string[]; archiveRows: TrainingSummaryRow[] }> {

  // ── 1. Parse raw Perigee Excel ─────────────────────────────────────────────
  const wb  = XLSX.read(fileBuffer, { type: 'buffer' });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (unknown)[][];

  if (raw.length < 2) throw new Error('No data rows found in uploaded file.');

  // ── 2. Map rows ────────────────────────────────────────────────────────────
  const rows: TrainingRow[] = raw.slice(1).map(r => {
    const a = r as (string | number | null)[];

    const trainedOn = CATEGORIES.map(c => s(a, c.yesIdx));
    const details   = CATEGORIES.map(c =>
      c.detailIdx != null ? s(a, c.detailIdx) : '',
    );

    const trainingCompleted = s(a, 15).toLowerCase() === 'yes';
    const noOfTrainings     = trainingCompleted ? 1 : 0;

    const date    = s(a, 9);
    const time    = s(a, 10);
    const fspsRaw = a[16];
    const durRaw  = a[53];

    return {
      trainerName:       [s(a, 2), s(a, 3)].filter(Boolean).join(' '),
      date,
      time,
      store:             s(a, 6),
      channel:           s(a, 5),
      province:          s(a, 8),
      visitUuid:         s(a, 11),
      trainingCompleted,
      fsps:              fspsRaw != null && fspsRaw !== '' ? Number(fspsRaw) : null as number | null,
      trainedOn,
      details,
      noOfTrainings,
      durationMins:      durRaw != null && durRaw !== '' ? Number(durRaw) : null as number | null,
      dateTimeMs:        parseDateTimeMs(date, time),
    };
  }).filter(r => r.trainerName);

  if (!rows.length) throw new Error('No valid training records found in uploaded file.');

  const rawDates  = rows.map(r => r.date).filter(d => d.includes('/'));
  const dateLabel = latestDateLabel(rawDates);
  const genDate   = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' });

  // ── 3. Date period filtering (MTD / WTD) ────────────────────────────────────
  const allParsedDates = rows
    .map(r => parseDdMmYyyy(r.date))
    .filter((d): d is Date => d !== null)
    .sort((a, b) => b.getTime() - a.getTime());

  const latestDate = allParsedDates.length > 0 ? allParsedDates[0] : null;

  // MTD: same month + year as latest date
  function isMtd(dateStr: string): boolean {
    if (!latestDate) return true;
    const d = parseDdMmYyyy(dateStr);
    return d != null && d.getMonth() === latestDate.getMonth() && d.getFullYear() === latestDate.getFullYear();
  }

  // WTD: Mon–Sun of the week containing the latest date
  let weekStart: Date | null = null;
  let weekEnd:   Date | null = null;
  if (latestDate) {
    const day = latestDate.getDay(); // 0=Sun … 6=Sat
    const mondayOffset = day === 0 ? -6 : 1 - day;
    weekStart = new Date(latestDate.getFullYear(), latestDate.getMonth(), latestDate.getDate() + mondayOffset);
    weekEnd   = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + 6);
  }

  function isWtd(dateStr: string): boolean {
    if (!weekStart || !weekEnd) return true;
    const d = parseDdMmYyyy(dateStr);
    if (!d) return false;
    const dNorm = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    return dNorm >= weekStart! && dNorm <= weekEnd!;
  }

  // Current upload's MTD and WTD rows
  const currentMtdRows = rows.filter(r => isMtd(r.date));
  const wtdRows        = rows.filter(r => isWtd(r.date));

  // Merge historical archived rows (from previous uploads this month) into MTD
  // Deduplicate by visitUuid to avoid double-counting re-uploads
  const seenUuids = new Set(currentMtdRows.map(r => r.visitUuid).filter(Boolean));
  const historicalMtd: TrainingRow[] = [];
  if (historicalRows?.length) {
    for (const hr of historicalRows) {
      if (!isMtd(hr.date)) continue;
      if (hr.visitUuid && seenUuids.has(hr.visitUuid)) continue;
      if (hr.visitUuid) seenUuids.add(hr.visitUuid);
      // Convert summary row to full TrainingRow (empty trainedOn/details)
      historicalMtd.push({
        trainerName:       hr.trainerName,
        date:              hr.date,
        time:              '',
        store:             hr.store,
        channel:           hr.channel,
        province:          hr.province,
        visitUuid:         hr.visitUuid,
        trainingCompleted: hr.trainingCompleted,
        fsps:              hr.fsps,
        trainedOn:         [],
        details:           [],
        noOfTrainings:     hr.noOfTrainings,
        durationMins:      null,
        dateTimeMs:        null,
      });
    }
  }
  const mtdRows = [...currentMtdRows, ...historicalMtd];

  // ── 4. KPI values ──────────────────────────────────────────────────────────
  const mtdKpi = computeKpis(mtdRows);
  const wtdKpi = computeKpis(wtdRows);

  // ── 5. Build workbook ──────────────────────────────────────────────────────
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
  const totalVisitsAll     = rows.length;
  const totalTrainingsAll  = rows.filter(r => r.trainingCompleted).length;
  const totalFspsAll       = rows.reduce((sum, r) => sum + (r.fsps ?? 0), 0);

  function kpiValues(kpi: ReturnType<typeof computeKpis>) {
    return [
      { title: 'Total Channels',      value: kpi.channels  },
      { title: 'Total BAs',           value: kpi.bas       },
      { title: 'Trainings Completed', value: kpi.trainings },
      { title: 'FSPs Trained',        value: kpi.fsps      },
      { title: 'Minutes Trained',     value: kpi.minutes   },
    ];
  }

  buildMenuSheet(ewb, {
    title:    `${brand.toUpperCase()} TRAINING FEEDBACK REPORT`,
    subtitle: `Generated ${genDate} \u00b7 ${totalVisitsAll} visit${totalVisitsAll === 1 ? '' : 's'} \u00b7 ${totalTrainingsAll} trainings completed \u00b7 ${totalFspsAll} FSPs trained`,
    sheetDefs: [
      { name: 'OVERALL',                desc: 'Full training data — one row per visit, all categories' },
      { name: 'CHANNEL VIEW',           desc: 'Training count summary by channel' },
      { name: 'BA SUMMARY',             desc: 'FSPs trained per BA — WTD vs MTD side-by-side with data bars' },
      { name: 'TRAINER SUMMARY (MTD)',   desc: 'Per-trainer breakdown for the current month' },
      { name: 'TRAINER SUMMARY (WTD)',   desc: 'Per-trainer breakdown for the current week' },
    ],
    kpiRows: [
      { label: 'MONTH TO DATE', values: kpiValues(mtdKpi) },
      { label: 'WEEK TO DATE',  values: kpiValues(wtdKpi) },
    ],
    defyLogoId,
    atomicLogoId,
    perigeeLogoId,
  });

  // ── Col count helpers ─────────────────────────────────────────────────────
  // OVERALL: TRAINER + DATE + PLACE + 1 per category + 1 per detail + 5 trailing
  const detailCount = CATEGORIES.filter(c => c.detailIdx != null).length;
  const ovColCount  = 3 + CATEGORIES.length + detailCount + 5;
  const cvColCount  = 5;
  const tsColCount  = 8;

  // ─────────────────────────────────────────────────────────────────────────
  // Sheet 2: OVERALL
  // ─────────────────────────────────────────────────────────────────────────
  const ows = ewb.addWorksheet('OVERALL');
  ows.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];

  addNavRow(ows, `${brand.toUpperCase()} TRAINING FEEDBACK \u2014 ALL VISITS`, ovColCount);

  // Row 2: column headers
  const ovHeaders: string[] = ['TRAINER NAME', 'DATE', 'PLACE'];
  for (const cat of CATEGORIES) {
    ovHeaders.push(`TRAINING ON ${cat.name}?`);
    if (cat.detailIdx != null) {
      ovHeaders.push(`PRODUCTS IN THE ${cat.name} CATEGORY YOU PROVIDED TRAINING ON?`);
    }
  }
  ovHeaders.push('TRAINING COMPLETED', 'NO OF TRAININGS DONE', 'TRAINING (IN MINUTES)', 'CHANNEL', 'REGION');
  applyHeader(ows.addRow(ovHeaders));

  // AutoFilter on header row (row 2)
  ows.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: ovHeaders.length } };

  // Column widths
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
  ows.getColumn(colCursor).width     = 18;  // TRAINING COMPLETED
  ows.getColumn(colCursor + 1).width = 18;  // NO OF TRAININGS DONE
  ows.getColumn(colCursor + 2).width = 16;  // TRAINING (IN MINUTES)
  ows.getColumn(colCursor + 3).width = 18;  // CHANNEL
  ows.getColumn(colCursor + 4).width = 18;  // REGION

  // Data rows
  for (const row of rows) {
    const vals: (string | number | null)[] = [row.trainerName, row.date, row.store];
    for (let i = 0; i < CATEGORIES.length; i++) {
      vals.push(row.trainedOn[i] || null);
      if (CATEGORIES[i].detailIdx != null) {
        vals.push(row.details[i] || null);
      }
    }
    vals.push(
      row.trainingCompleted ? 'Yes' : 'No',
      row.noOfTrainings,
      row.durationMins,
      row.channel || null,
      row.province || null,
    );
    const dr = ows.addRow(vals);
    dr.eachCell(cell => { cell.font = DATA_FONT; cell.alignment = { vertical: 'middle', wrapText: false }; });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // Sheet 3: CHANNEL VIEW
  // ─────────────────────────────────────────────────────────────────────────
  const cvws = ewb.addWorksheet('CHANNEL VIEW');
  cvws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
  addNavRow(cvws, 'CHANNEL VIEW \u2014 TRAINING SUMMARY BY CHANNEL', cvColCount);

  cvws.getColumn(1).width = 28;  // CHANNEL
  cvws.getColumn(2).width = 16;  // NO. OF STORES
  cvws.getColumn(3).width = 18;  // TRAININGS DONE
  cvws.getColumn(4).width = 16;  // NO. OF FSPs
  cvws.getColumn(5).width = 20;  // % OF TRAINING DONE

  applyHeader(cvws.addRow([
    'CHANNEL', 'NO. OF STORES', 'TRAININGS DONE', 'NO. OF FSPs', '% OF TRAINING DONE',
  ]));

  // Aggregate by channel (only completed trainings count for stores/trainings/FSPs)
  const channelMap = new Map<string, { stores: Set<string>; trainings: number; fsps: number }>();
  for (const row of rows) {
    const ch  = row.channel || '(blank)';
    let cur = channelMap.get(ch);
    if (!cur) {
      cur = { stores: new Set(), trainings: 0, fsps: 0 };
      channelMap.set(ch, cur);
    }
    if (row.trainingCompleted) {
      cur.stores.add(row.store);
      cur.trainings++;
      cur.fsps += row.fsps ?? 0;
    }
  }

  const totalTrainingsDone = [...channelMap.values()].reduce((sum, d) => sum + d.trainings, 0);

  for (const [ch, data] of [...channelMap.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    const pct = totalTrainingsDone > 0
      ? `${((data.trainings / totalTrainingsDone) * 100).toFixed(1)}%`
      : '0.0%';
    const dr = cvws.addRow([ch, data.stores.size, data.trainings, data.fsps, pct]);
    dr.eachCell(cell => { cell.font = DATA_FONT; cell.alignment = { vertical: 'middle' }; });
  }

  // Grand Total row
  const totalStores = new Set(
    rows.filter(r => r.trainingCompleted).map(r => `${r.channel}||${r.store}`),
  ).size;
  const totalFsps = [...channelMap.values()].reduce((sum, d) => sum + d.fsps, 0);
  const cvTotal = cvws.addRow(['Grand Total', totalStores, totalTrainingsDone, totalFsps, '100%']);
  cvTotal.eachCell(cell => { cell.font = TOTAL_FONT; cell.alignment = { vertical: 'middle' }; });

  // ─────────────────────────────────────────────────────────────────────────
  // Sheet 4: BA SUMMARY — WTD (left) + MTD (right) side-by-side
  // ─────────────────────────────────────────────────────────────────────────
  {
    const baColCount = 10; // A-J
    const baws = ewb.addWorksheet('BA SUMMARY');
    baws.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];
    addNavRow(baws, 'BA SUMMARY \u2014 FSPs TRAINED PER BA', baColCount);

    const ORANGE_FILL: ExcelJS.Fill = {
      type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4A460' },
    };
    const ORANGE_HEADER: ExcelJS.Fill = {
      type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8912D' },
    };
    const REGION_COLORS = [
      'FF4472C4', 'FFED7D31', 'FFA5A5A5', 'FFFFC000',
      'FF5B9BD5', 'FF70AD47', 'FF264478', 'FF9B57A0',
    ];

    // ── Aggregate WTD by trainer ────────────────────────────────────
    const wtdByTrainer = new Map<string, number>();
    const wtdRegions   = new Set<string>();
    for (const r of wtdRows) {
      if (!r.trainingCompleted) continue;
      wtdByTrainer.set(r.trainerName, (wtdByTrainer.get(r.trainerName) ?? 0) + (r.fsps ?? 0));
      if (r.province) wtdRegions.add(r.province);
    }
    const wtdSorted = [...wtdByTrainer.entries()].sort((a, b) => b[1] - a[1]);
    const wtdMax    = wtdSorted.length > 0 ? wtdSorted[0][1] : 1;

    // ── Aggregate MTD by trainer ────────────────────────────────────
    const mtdByTrainer = new Map<string, number>();
    const mtdRegions   = new Set<string>();
    for (const r of mtdRows) {
      if (!r.trainingCompleted) continue;
      mtdByTrainer.set(r.trainerName, (mtdByTrainer.get(r.trainerName) ?? 0) + (r.fsps ?? 0));
      if (r.province) mtdRegions.add(r.province);
    }
    const mtdSorted = [...mtdByTrainer.entries()].sort((a, b) => b[1] - a[1]);
    const mtdMax    = mtdSorted.length > 0 ? mtdSorted[0][1] : 1;

    // ── Date labels ─────────────────────────────────────────────────
    const wtdDateRange = weekStart && weekEnd
      ? `${pad2(weekStart.getDate())}/${pad2(weekStart.getMonth()+1)} - ${pad2(weekEnd.getDate())}/${pad2(weekEnd.getMonth()+1)}/${weekEnd.getFullYear()}`
      : '';
    const mtdMonthName = latestDate
      ? latestDate.toLocaleString('en-US', { month: 'long' }).toUpperCase()
      : '';

    // ── Column widths ───────────────────────────────────────────────
    baws.getColumn(1).width  = 28; // WTD Trainer Name
    baws.getColumn(2).width  = 20; // WTD FSPs
    baws.getColumn(3).width  = 3;  // spacer
    baws.getColumn(4).width  = 16; // WTD Region legend
    baws.getColumn(5).width  = 16; // WTD Region legend color
    baws.getColumn(6).width  = 3;  // spacer
    baws.getColumn(7).width  = 28; // MTD Trainer Name
    baws.getColumn(8).width  = 20; // MTD FSPs
    baws.getColumn(9).width  = 16; // MTD Region legend
    baws.getColumn(10).width = 16; // MTD Region legend color

    // ── Row 2: Titles ───────────────────────────────────────────────
    const titleRow = 2;
    const wtdTitleCell = baws.getRow(titleRow).getCell(1);
    wtdTitleCell.value = `TOTAL FSPS TRAINED BY BAS ${wtdDateRange}`;
    wtdTitleCell.font  = { bold: true, size: 11, name: 'Arial', color: { argb: 'FF333333' } };
    wtdTitleCell.alignment = { vertical: 'middle' };
    baws.mergeCells(titleRow, 1, titleRow, 2);

    const mtdTitleCell = baws.getRow(titleRow).getCell(7);
    mtdTitleCell.value = `TRAINING CONDUCTED BY BAS ${mtdMonthName}`;
    mtdTitleCell.font  = { bold: true, size: 11, name: 'Arial', color: { argb: 'FF333333' } };
    mtdTitleCell.alignment = { vertical: 'middle' };
    baws.mergeCells(titleRow, 7, titleRow, 8);

    // ── Row 3: blank (filter/search area placeholder) ───────────────
    baws.getRow(3).height = 10;

    // ── Row 4: Headers ──────────────────────────────────────────────
    const hdrRow = 4;

    // WTD headers
    const wtdH1 = baws.getRow(hdrRow).getCell(1);
    wtdH1.value = 'TRAINER NAME';
    wtdH1.fill  = ORANGE_HEADER;
    wtdH1.font  = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
    wtdH1.alignment = { vertical: 'middle', horizontal: 'left' };

    const wtdH2 = baws.getRow(hdrRow).getCell(2);
    wtdH2.value = 'Total FSPs Trained';
    wtdH2.fill  = ORANGE_HEADER;
    wtdH2.font  = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
    wtdH2.alignment = { vertical: 'middle', horizontal: 'center' };

    // Region legend header (WTD)
    const wtdRH = baws.getRow(hdrRow).getCell(4);
    wtdRH.value = 'REGION';
    wtdRH.fill  = ORANGE_HEADER;
    wtdRH.font  = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
    wtdRH.alignment = { vertical: 'middle', horizontal: 'center' };
    baws.mergeCells(hdrRow, 4, hdrRow, 5);

    // MTD headers
    const mtdH1 = baws.getRow(hdrRow).getCell(7);
    mtdH1.value = 'TRAINER NAME';
    mtdH1.fill  = ORANGE_HEADER;
    mtdH1.font  = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
    mtdH1.alignment = { vertical: 'middle', horizontal: 'left' };

    const mtdH2 = baws.getRow(hdrRow).getCell(8);
    mtdH2.value = 'Total FSPs Trained';
    mtdH2.fill  = ORANGE_HEADER;
    mtdH2.font  = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
    mtdH2.alignment = { vertical: 'middle', horizontal: 'center' };

    // Region legend header (MTD)
    const mtdRH = baws.getRow(hdrRow).getCell(9);
    mtdRH.value = 'REGION';
    mtdRH.fill  = ORANGE_HEADER;
    mtdRH.font  = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Arial', size: 10 };
    mtdRH.alignment = { vertical: 'middle', horizontal: 'center' };
    baws.mergeCells(hdrRow, 9, hdrRow, 10);

    baws.getRow(hdrRow).height = 28;

    // AutoFilter on WTD columns
    baws.autoFilter = { from: { row: hdrRow, column: 1 }, to: { row: hdrRow, column: 2 } };

    // ── Data rows ───────────────────────────────────────────────────
    const dataStart = hdrRow + 1;
    const maxRows   = Math.max(wtdSorted.length, mtdSorted.length, 1);

    for (let i = 0; i < maxRows; i++) {
      const rowNum = dataStart + i;

      // WTD data
      if (i < wtdSorted.length) {
        const [name, fsps] = wtdSorted[i];
        const nameCell = baws.getRow(rowNum).getCell(1);
        nameCell.value     = name;
        nameCell.font      = DATA_FONT;
        nameCell.alignment = { vertical: 'middle' };

        const fspsCell = baws.getRow(rowNum).getCell(2);
        fspsCell.value     = fsps;
        fspsCell.font      = DATA_FONT;
        fspsCell.alignment = { vertical: 'middle', horizontal: 'center' };

        // Orange data bar proportional fill
        const barWidth = wtdMax > 0 ? fsps / wtdMax : 0;
        if (barWidth > 0) {
          fspsCell.fill = {
            type: 'pattern', pattern: 'solid',
            fgColor: { argb: `FF${interpolateOrange(barWidth)}` },
          };
        }
      }

      // MTD data
      if (i < mtdSorted.length) {
        const [name, fsps] = mtdSorted[i];
        const nameCell = baws.getRow(rowNum).getCell(7);
        nameCell.value     = name;
        nameCell.font      = DATA_FONT;
        nameCell.alignment = { vertical: 'middle' };

        const fspsCell = baws.getRow(rowNum).getCell(8);
        fspsCell.value     = fsps;
        fspsCell.font      = DATA_FONT;
        fspsCell.alignment = { vertical: 'middle', horizontal: 'center' };

        const barWidth = mtdMax > 0 ? fsps / mtdMax : 0;
        if (barWidth > 0) {
          fspsCell.fill = {
            type: 'pattern', pattern: 'solid',
            fgColor: { argb: `FF${interpolateOrange(barWidth)}` },
          };
        }
      }
    }

    // ── Region legends ──────────────────────────────────────────────
    const wtdRegionArr = [...wtdRegions].sort();
    for (let i = 0; i < wtdRegionArr.length; i++) {
      const rowNum = dataStart + i;
      const colorIdx = i % REGION_COLORS.length;

      const labelCell = baws.getRow(rowNum).getCell(4);
      labelCell.value     = wtdRegionArr[i];
      labelCell.font      = { ...DATA_FONT, color: { argb: 'FFFFFFFF' }, bold: true };
      labelCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: REGION_COLORS[colorIdx] } };
      labelCell.alignment = { vertical: 'middle', horizontal: 'center' };
      baws.mergeCells(rowNum, 4, rowNum, 5);
    }

    const mtdRegionArr = [...mtdRegions].sort();
    for (let i = 0; i < mtdRegionArr.length; i++) {
      const rowNum = dataStart + i;
      const colorIdx = i % REGION_COLORS.length;

      const labelCell = baws.getRow(rowNum).getCell(9);
      labelCell.value     = mtdRegionArr[i];
      labelCell.font      = { ...DATA_FONT, color: { argb: 'FFFFFFFF' }, bold: true };
      labelCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: REGION_COLORS[colorIdx] } };
      labelCell.alignment = { vertical: 'middle', horizontal: 'center' };
      baws.mergeCells(rowNum, 9, rowNum, 10);
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // Sheet 5: TRAINER SUMMARY (MTD)
  // ─────────────────────────────────────────────────────────────────────────
  buildTrainerSummarySheet(
    ewb,
    'TRAINER SUMMARY (MTD)',
    'TRAINER SUMMARY (MTD) \u2014 PER-REP BREAKDOWN',
    mtdRows,
    tsColCount,
  );

  // ─────────────────────────────────────────────────────────────────────────
  // Sheet 5: TRAINER SUMMARY (WTD)
  // ─────────────────────────────────────────────────────────────────────────
  buildTrainerSummarySheet(
    ewb,
    'TRAINER SUMMARY (WTD)',
    'TRAINER SUMMARY (WTD) \u2014 PER-REP BREAKDOWN',
    wtdRows,
    tsColCount,
  );

  // ── 7. Build archive rows (lightweight, for blob persistence) ─────────────
  const archiveRows: TrainingSummaryRow[] = rows.map(r => ({
    trainerName:       r.trainerName,
    date:              r.date,
    channel:           r.channel,
    province:          r.province,
    store:             r.store,
    visitUuid:         r.visitUuid,
    fsps:              r.fsps,
    trainingCompleted: r.trainingCompleted,
    noOfTrainings:     r.noOfTrainings,
  }));

  // ── 8. Filename & buffer ───────────────────────────────────────────────────
  const filename = `${brand.toUpperCase()}_TRAINING_FEEDBACK_REPORT_${dateLabel}.xlsx`;
  const buffer   = Buffer.from(await ewb.xlsx.writeBuffer());

  return { buffer, filename, rawDates, archiveRows };
}

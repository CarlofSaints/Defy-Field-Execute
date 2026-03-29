/**
 * Shared MENU sheet builder for Defy Field Execute reports.
 *
 * Layout (header design — no standalone red banner):
 *  Row 1  │ thin top margin
 *  Row 2  │ report title (large red bold) + Defy/Atomic logos floating at right
 *  Row 3  │ subtitle (grey)
 *  Row 4  │ spacer
 *  Row 5  │ REPORT INDEX (red bar)
 *  Row 6  │ column headers (#, SHEET, DESCRIPTION)
 *  Row 7+ │ one row per sheet (alternating stripes, hyperlinked sheet names)
 *  …      │ optional note row
 *  last   │ spacer + "Powered by Perigee" row with Perigee logo bottom-right
 */
import ExcelJS from 'exceljs';

const DEFY_RED = 'E31837';
const WHITE    = 'FFFFFF';

export interface MenuSheetDef {
  name: string;
  desc: string;
}

export interface BuildMenuOpts {
  title:          string;
  subtitle:       string;
  sheetDefs:      MenuSheetDef[];
  note?:          string;
  defyLogoId?:    number | null;
  atomicLogoId?:  number | null;
  perigeeLogoId?: number | null;
}

// ─── Shared cell stylers (also used by data-sheet builders) ──────────────────
export function applyHeaderStyle(cell: ExcelJS.Cell) {
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

export function applyDataStyle(cell: ExcelJS.Cell, isEven = false) {
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

// ─── Nav row: slim red bar at top of every data sheet ────────────────────────
export function addNavRow(ws: ExcelJS.Worksheet, sheetTitle: string, colCount: number) {
  const row = ws.getRow(1);
  row.height = 22;
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DEFY_RED } };
  }
  const title = row.getCell(1);
  title.value = sheetTitle;
  title.font  = { bold: true, color: { argb: WHITE }, size: 10, name: 'Arial' };
  title.alignment = { vertical: 'middle', horizontal: 'left' };

  const menuCell = row.getCell(colCount);
  menuCell.value = { text: '⬅  MENU', hyperlink: '#MENU!A1', tooltip: 'Return to Menu' };
  menuCell.font  = { bold: true, color: { argb: WHITE }, size: 10, name: 'Arial', underline: true };
  menuCell.alignment = { vertical: 'middle', horizontal: 'right' };
}

// ─── MENU sheet builder ───────────────────────────────────────────────────────
export function buildMenuSheet(wb: ExcelJS.Workbook, opts: BuildMenuOpts): void {
  const { title, subtitle, sheetDefs, note, defyLogoId, atomicLogoId, perigeeLogoId } = opts;

  const ws = wb.addWorksheet('MENU');
  ws.views = [{ showGridLines: false }];

  // Row 1: thin top margin
  ws.getRow(1).height = 4;

  // Row 2: title (large red bold) + Defy/Atomic logos floating to the right
  // Logos are positioned in col C (0-indexed col 2); image row index = 1 (0-indexed)
  ws.getRow(2).height = 58;
  const titleCell = ws.getRow(2).getCell(1);
  titleCell.value = title;
  titleCell.font  = { bold: true, color: { argb: DEFY_RED }, size: 22, name: 'Arial' };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.mergeCells(2, 1, 2, 3); // title spans cols A–C so text has room at 22pt

  // Defy logo — top-right corner of row 2, clear of title text
  if (defyLogoId != null) {
    ws.addImage(defyLogoId, { tl: { col: 2.72, row: 1.04 }, ext: { width: 95, height: 52 } });
  }

  // Row 3: subtitle
  ws.getRow(3).height = 22;
  const subCell = ws.getRow(3).getCell(1);
  subCell.value   = subtitle;
  subCell.font    = { color: { argb: '666666' }, size: 11, name: 'Arial' };
  subCell.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.mergeCells(3, 1, 3, 3);

  // Row 4: spacer
  ws.getRow(4).height = 10;

  // Row 5: REPORT INDEX header
  ws.getRow(5).height = 26;
  const idxHeader = ws.getRow(5).getCell(1);
  idxHeader.value = 'REPORT INDEX';
  idxHeader.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: DEFY_RED } };
  idxHeader.font  = { bold: true, color: { argb: WHITE }, size: 11, name: 'Arial' };
  idxHeader.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.mergeCells(5, 1, 5, 3);

  // Row 6: column headers
  ws.getRow(6).height = 22;
  ['#', 'PROBLEM', 'DESCRIPTION'].forEach((h, i) => {
    const cell = ws.getRow(6).getCell(i + 1);
    cell.value = h;
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2D2D2D' } };
    cell.font  = { bold: true, color: { argb: WHITE }, size: 10, name: 'Arial' };
    cell.alignment = { vertical: 'middle', horizontal: i === 0 ? 'center' : 'left' };
  });

  // Rows 7+: one row per sheet
  sheetDefs.forEach(({ name, desc }, idx) => {
    const rowNum   = 7 + idx;
    const isEven   = idx % 2 === 0;
    const bg       = isEven ? 'F7F7F7' : WHITE;
    const safeName = name.includes(' ') ? `'${name}'` : name;

    ws.getRow(rowNum).height = 22;

    const numCell = ws.getRow(rowNum).getCell(1);
    numCell.value = idx + 1;
    numCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    numCell.font  = { bold: true, color: { argb: DEFY_RED }, size: 10, name: 'Arial' };
    numCell.alignment = { vertical: 'middle', horizontal: 'center' };
    numCell.border = {
      bottom: { style: 'thin', color: { argb: 'E0E0E0' } },
      right:  { style: 'thin', color: { argb: 'E0E0E0' } },
    };

    const nameCell = ws.getRow(rowNum).getCell(2);
    nameCell.value = { text: name, hyperlink: `#${safeName}!A1`, tooltip: `Go to ${name}` };
    nameCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    nameCell.font  = { bold: true, color: { argb: '0563C1' }, size: 10, name: 'Arial', underline: true };
    nameCell.alignment = { vertical: 'middle', horizontal: 'left' };
    nameCell.border = {
      bottom: { style: 'thin', color: { argb: 'E0E0E0' } },
      right:  { style: 'thin', color: { argb: 'E0E0E0' } },
    };

    const descCell = ws.getRow(rowNum).getCell(3);
    descCell.value = desc;
    descCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    descCell.font  = { color: { argb: '444444' }, size: 10, name: 'Arial' };
    descCell.alignment = { vertical: 'middle', horizontal: 'left' };
    descCell.border = { bottom: { style: 'thin', color: { argb: 'E0E0E0' } } };
  });

  // Optional note row
  let nextFreeRow = 7 + sheetDefs.length;
  if (note) {
    nextFreeRow++;
    ws.getRow(nextFreeRow).height = 32;
    const noteCell = ws.getRow(nextFreeRow).getCell(1);
    noteCell.value = note;
    noteCell.font  = { italic: true, color: { argb: '888888' }, size: 9, name: 'Arial' };
    noteCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    ws.mergeCells(nextFreeRow, 1, nextFreeRow, 3);
  }

  // Spacer
  nextFreeRow++;
  ws.getRow(nextFreeRow).height = 10;

  // "Powered by Perigee" row
  nextFreeRow++;
  ws.getRow(nextFreeRow).height = 65;
  const poweredByCell = ws.getRow(nextFreeRow).getCell(2);
  poweredByCell.value = 'Powered by:';
  poweredByCell.font  = { color: { argb: '888888' }, size: 9, name: 'Arial' };
  poweredByCell.alignment = { vertical: 'middle', horizontal: 'right' };

  if (perigeeLogoId != null) {
    // ws row nextFreeRow (1-indexed) → image row (nextFreeRow - 1) (0-indexed)
    ws.addImage(perigeeLogoId, {
      tl:  { col: 2.04, row: nextFreeRow - 1 + 0.1 },
      ext: { width: 55, height: 50 },
    });
  }
  // Atomic logo — bottom-right, alongside Perigee in the "Powered by" row
  if (atomicLogoId != null) {
    ws.addImage(atomicLogoId, {
      tl:  { col: 2.53, row: nextFreeRow - 1 + 0.3 },
      ext: { width: 130, height: 27 },
    });
  }

  // Column widths
  ws.getColumn(1).width = 5;
  ws.getColumn(2).width = 28;
  ws.getColumn(3).width = 68;
}

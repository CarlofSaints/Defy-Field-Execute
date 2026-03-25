/**
 * inject-slicers.ts
 *
 * Post-processes an xlsx buffer (produced by exceljs) to inject Excel Table
 * slicers via raw OOXML / JSZip manipulation.
 *
 * Excel slicer XML structure (x14 namespace, Excel 2010+):
 *
 *   xl/slicers/slicer{N}.xml            — visual properties for one slicer
 *   xl/slicerCaches/slicerCache{N}.xml  — connects slicer to a specific table column
 *   xl/drawings/drawing{N}.xml          — positions slicer shapes on the sheet
 *   xl/drawings/_rels/drawing{N}.xml.rels
 *   xl/worksheets/sheet{N}.xml          — gets <drawing> ref + slicer extLst
 *   xl/worksheets/_rels/sheet{N}.xml.rels
 *   xl/_rels/workbook.xml.rels          — workbook → slicer cache refs
 *   [Content_Types].xml                 — content types for all new files
 *
 * The "hidden column" trick:
 *   Dimension columns (PROVINCE, CATEGORY, SUB CAT) are included in the
 *   Excel Table but hidden via column.hidden = true in exceljs before
 *   generating the buffer.  Slicers reference these hidden columns by
 *   their 1-based table-column index.  The table looks like a clean summary
 *   (only STORE NAME / PRODUCT CODE / PRODUCT / SOH visible) while still
 *   being fully filterable by the hidden dimensions.
 */

import JSZip from 'jszip';

// ─── Public types ────────────────────────────────────────────────────────────

export interface SlicerSpec {
  /** Unique identifier used as the slicer name and in cache name, e.g. "Province" */
  name: string;
  /** Text shown in the slicer header, e.g. "PROVINCE" */
  caption: string;
  /** 1-based column index in the Excel Table (matches the column order in addTable) */
  tableColumn: number;
  /** Anchor position on the sheet (0-based col/row, like Excel drawing coords) */
  pos: { fromCol: number; fromRow: number; toCol: number; toRow: number };
}

// ─── Main export ─────────────────────────────────────────────────────────────

export async function injectTableSlicers(
  buffer: Buffer,
  sheetName: string,
  tableName: string,
  slicers: SlicerSpec[],
): Promise<Buffer> {

  const zip = await JSZip.loadAsync(buffer);

  // ── 1. Resolve sheet file path ────────────────────────────────────────────
  const wbXml     = await zip.file('xl/workbook.xml')!.async('text');
  const wbRelsXml = await zip.file('xl/_rels/workbook.xml.rels')!.async('text');

  const sheetEl = wbXml.match(
    new RegExp(`<sheet[^>]+name="${escAttr(sheetName)}"[^>]+r:id="(rId[^"]+)"`)
  );
  if (!sheetEl) throw new Error(`injectTableSlicers: sheet not found: "${sheetName}"`);
  const sheetRId = sheetEl[1];

  const relEl = wbRelsXml.match(new RegExp(`Id="${sheetRId}"[^>]+Target="([^"]+)"`));
  if (!relEl) throw new Error(`injectTableSlicers: sheet rel not found for ${sheetRId}`);

  // Target is relative to xl/, e.g. "worksheets/sheet2.xml"
  const sheetTarget   = relEl[1];                      // "worksheets/sheet2.xml"
  const sheetPath     = `xl/${sheetTarget}`;           // "xl/worksheets/sheet2.xml"
  const sheetFileName = sheetTarget.split('/').pop()!; // "sheet2.xml"
  const sheetRelsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;

  // ── 2. Find the Excel Table ID for tableName ──────────────────────────────
  let tableId = 1;
  const sheetRelsFile = zip.file(sheetRelsPath);

  if (sheetRelsFile) {
    const sheetRels = await sheetRelsFile.async('text');
    for (const m of sheetRels.matchAll(/Target="\.\.\/tables\/(table\d+\.xml)"/g)) {
      const tblFile = zip.file(`xl/tables/${m[1]}`);
      if (!tblFile) continue;
      const tblXml = await tblFile.async('text');
      if (tblXml.includes(`name="${tableName}"`)) {
        const idM = tblXml.match(/\bid="(\d+)"/);
        if (idM) tableId = parseInt(idM[1], 10);
        break;
      }
    }
  }

  // ── 3. Determine next free file indices ───────────────────────────────────
  const countFiles = (prefix: string, pattern: RegExp) =>
    Object.keys(zip.files).filter(k => pattern.test(k)).length;

  const firstSlicerIdx  = countFiles('slicer',      /^xl\/slicers\/slicer\d+\.xml$/)  + 1;
  const firstCacheIdx   = countFiles('slicerCache', /^xl\/slicerCaches\/slicerCache\d+\.xml$/) + 1;
  const drawingIdx      = countFiles('drawing',     /^xl\/drawings\/drawing\d+\.xml$/) + 1;

  // ── 4. Create slicer cache files (one per slicer) ─────────────────────────
  for (let i = 0; i < slicers.length; i++) {
    const s         = slicers[i];
    const cacheIdx  = firstCacheIdx + i;
    const cacheName = `SlicerCache_${s.name}`;
    zip.file(
      `xl/slicerCaches/slicerCache${cacheIdx}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<slicerCacheDefinition ` +
        `xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `name="${cacheName}" sourceName="${esc(s.caption)}">\r\n` +
      `  <tableSlicerCache tableId="${tableId}" column="${s.tableColumn}" ` +
        `sortOrder="ascending" customListSort="1"/>\r\n` +
      `  <slicerCacheData><items count="0"/></slicerCacheData>\r\n` +
      `</slicerCacheDefinition>`,
    );
  }

  // ── 5. Create slicer definition files (one per slicer) ────────────────────
  for (let i = 0; i < slicers.length; i++) {
    const s          = slicers[i];
    const slicerIdx  = firstSlicerIdx + i;
    const cacheName  = `SlicerCache_${s.name}`;
    zip.file(
      `xl/slicers/slicer${slicerIdx}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<slicers ` +
        `xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" ` +
        `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
        `mc:Ignorable="x14" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">\r\n` +
      `  <slicer name="${esc(s.name)}" cache="${cacheName}" caption="${esc(s.caption)}" ` +
        `rowHeight="21420" columnWidth="101220" style="SlicerStyleLight1" startItem="0">\r\n` +
      `    <slicerStyleInfo rowHeight="21420" columnWidth="101220" headerHeight="18570"/>\r\n` +
      `  </slicer>\r\n` +
      `</slicers>`,
    );
  }

  // ── 6. Create drawing file (positions all slicers as floating shapes) ──────
  const anchors = slicers.map((s, i) => {
    const p = s.pos;
    return [
      `  <xdr:twoCellAnchor editAs="oneCell">`,
      `    <xdr:from>`,
      `      <xdr:col>${p.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff>`,
      `      <xdr:row>${p.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff>`,
      `    </xdr:from>`,
      `    <xdr:to>`,
      `      <xdr:col>${p.toCol}</xdr:col><xdr:colOff>0</xdr:colOff>`,
      `      <xdr:row>${p.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff>`,
      `    </xdr:to>`,
      `    <mc:AlternateContent`,
      `      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">`,
      `      <mc:Choice Requires="x14">`,
      `        <x14:slicer`,
      `          xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"`,
      `          r:id="rId${i + 1}"/>`,
      `      </mc:Choice>`,
      `      <mc:Fallback>`,
      `        <xdr:sp macro="" textlink="">`,
      `          <xdr:nvSpPr>`,
      `            <xdr:cNvPr id="${i + 2}" name="${escAttr(s.name)} ${i + 1}"/>`,
      `            <xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr>`,
      `          </xdr:nvSpPr>`,
      `          <xdr:spPr>`,
      `            <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>`,
      `            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`,
      `          </xdr:spPr>`,
      `        </xdr:sp>`,
      `      </mc:Fallback>`,
      `    </mc:AlternateContent>`,
      `    <xdr:clientData/>`,
      `  </xdr:twoCellAnchor>`,
    ].join('\r\n');
  }).join('\r\n');

  zip.file(
    `xl/drawings/drawing${drawingIdx}.xml`,
    [
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
      `<xdr:wsDr`,
      `  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"`,
      `  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`,
      `  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`,
      `  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">`,
      anchors,
      `</xdr:wsDr>`,
    ].join('\r\n'),
  );

  // ── 7. Drawing rels (each anchor rId → its slicer file) ───────────────────
  const drawingRelEntries = slicers.map((_, i) =>
    `  <Relationship Id="rId${i + 1}" ` +
    `Type="http://schemas.microsoft.com/office/2007/relationships/slicer" ` +
    `Target="../slicers/slicer${firstSlicerIdx + i}.xml"/>`,
  ).join('\r\n');

  zip.file(
    `xl/drawings/_rels/drawing${drawingIdx}.xml.rels`,
    [
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`,
      drawingRelEntries,
      `</Relationships>`,
    ].join('\r\n'),
  );

  // ── 8. Patch the sheet XML ─────────────────────────────────────────────────
  let sheetXml = await zip.file(sheetPath)!.async('text');

  // Use high rId numbers (1000+) to avoid collision with exceljs-generated rIds
  const drawingRId   = `rId1000`;
  const slicerRIds   = slicers.map((_, i) => `rId${1001 + i}`);

  // Add <drawing> element before </worksheet> (only if not already present)
  if (!sheetXml.includes('<drawing ')) {
    sheetXml = sheetXml.replace(
      '</worksheet>',
      `<drawing r:id="${drawingRId}"/></worksheet>`,
    );
  }

  // Add slicer extLst entries
  const slicerListItems = slicerRIds.map(rid => `<x14:slicer r:id="${rid}"/>`).join('');
  const slicerExt =
    `<ext uri="{A8765BA9-456A-4daa-B4F3-9B584D80E87A}" ` +
    `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
    `<x14:slicerList>${slicerListItems}</x14:slicerList>` +
    `</ext>`;

  if (sheetXml.includes('<extLst>')) {
    sheetXml = sheetXml.replace('</extLst>', `${slicerExt}</extLst>`);
  } else {
    sheetXml = sheetXml.replace('</worksheet>', `<extLst>${slicerExt}</extLst></worksheet>`);
  }

  zip.file(sheetPath, sheetXml);

  // ── 9. Patch sheet rels ────────────────────────────────────────────────────
  const existingSheetRels = sheetRelsFile
    ? await sheetRelsFile.async('text')
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n` +
      `</Relationships>`;

  const newSheetRelEntries = [
    `  <Relationship Id="${drawingRId}" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ` +
    `Target="../drawings/drawing${drawingIdx}.xml"/>`,
    ...slicers.map((_, i) =>
      `  <Relationship Id="${slicerRIds[i]}" ` +
      `Type="http://schemas.microsoft.com/office/2007/relationships/slicer" ` +
      `Target="../slicers/slicer${firstSlicerIdx + i}.xml"/>`,
    ),
  ].join('\r\n');

  zip.file(
    sheetRelsPath,
    existingSheetRels.replace('</Relationships>', `${newSheetRelEntries}\r\n</Relationships>`),
  );

  // ── 10. Patch workbook rels (add slicer cache references) ─────────────────
  // Use high rId numbers (2000+) to avoid collision
  const cacheRelEntries = slicers.map((_, i) =>
    `  <Relationship Id="rId${2000 + i}" ` +
    `Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" ` +
    `Target="slicerCaches/slicerCache${firstCacheIdx + i}.xml"/>`,
  ).join('\r\n');

  zip.file(
    'xl/_rels/workbook.xml.rels',
    wbRelsXml.replace('</Relationships>', `${cacheRelEntries}\r\n</Relationships>`),
  );

  // ── 11. Patch [Content_Types].xml ─────────────────────────────────────────
  let ctXml = await zip.file('[Content_Types].xml')!.async('text');

  const newCTs = [
    ...slicers.map((_, i) =>
      `  <Override PartName="/xl/slicers/slicer${firstSlicerIdx + i}.xml" ` +
      `ContentType="application/vnd.ms-excel.slicer+xml"/>`,
    ),
    ...slicers.map((_, i) =>
      `  <Override PartName="/xl/slicerCaches/slicerCache${firstCacheIdx + i}.xml" ` +
      `ContentType="application/vnd.ms-excel.slicerCache+xml"/>`,
    ),
    `  <Override PartName="/xl/drawings/drawing${drawingIdx}.xml" ` +
    `ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
  ].join('\r\n');

  ctXml = ctXml.replace('</Types>', `${newCTs}\r\n</Types>`);
  zip.file('[Content_Types].xml', ctXml);

  // ── 12. Return modified buffer ─────────────────────────────────────────────
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ─── XML helpers ─────────────────────────────────────────────────────────────
function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function escAttr(s: string): string {
  return esc(s).replace(/"/g, '&quot;');
}

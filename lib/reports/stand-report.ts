/**
 * Defy Stand Report builder
 *
 * Input:  Perigee raw export (single sheet, columns include First Name, Last Name,
 *         Date, Store, and one or more image URL columns)
 * Output: PowerPoint .pptx — cover + one content slide per 2 images + closing slide
 *
 * Column mapping (0-indexed):
 *   2  C  First Name
 *   3  D  Last Name
 *   6  G  Store (may be empty → "UNKNOWN")
 *   9  J  Date (DD/MM/YYYY)
 *  14+ any column whose value starts with "http" → image URL
 *
 * Images fetched from SharePoint (same folder as Red Flag images).
 * Max 2 images per slide; additional images spill onto a new slide.
 * Rep name added small at bottom-right of each content slide.
 */
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import fs   from 'fs';
import path from 'path';
import { downloadFileFromSP } from '@/lib/graph-oj';
import { loadAppSettings }    from '@/lib/appSettings';

// ─── Slide dimensions (A4 landscape, EMU) ────────────────────────────────────
const SLIDE_W = 10693400;
const SLIDE_H = 7562850;

// ─── Layout constants (EMU) ───────────────────────────────────────────────────
// Title bar (rounded-rect, full-width header)
const TITLE_X  = 7384;   const TITLE_Y  = 148092;
const TITLE_CX = 10686016; const TITLE_CY = 509133;

// Logo 1 — wide (Defy/Atomic, top-right) — rId2 = image10.png from template
const LOGO1_X  = 8477138; const LOGO1_Y  = 307586;
const LOGO1_CX = 1774601; const LOGO1_CY = 207248;

// Logo 2 — small (top-right, overlapping logo area) — rId3 = image11.png from template
const LOGO2_X  = 8023757; const LOGO2_Y  = 200025;
const LOGO2_CX = 371422;  const LOGO2_CY = 385321;

// Data images
const IMG_Y   = 1651000;
const IMG_CX  = 5080000;
const IMG_CY  = 5080000;
const IMG1_2UP_X = 63500;   // left image (2-up layout)
const IMG2_2UP_X = 5207000; // right image (2-up layout)
const IMG1_1UP_X = Math.floor((SLIDE_W - IMG_CX) / 2); // centred (1-up)

// Label above each image
const LABEL_Y  = 1169889;
const LABEL_CX = 1300000;
const LABEL_CY = 246221;

// Date below each image
const DATE_Y   = 6858000;
const DATE_CX  = 1400000;
const DATE_CY  = 250000;

// Rep name (bottom-right, small)
const REP_X    = 8700000;
const REP_Y    = 7340000;
const REP_CX   = 1900000;
const REP_CY   = 210000;

// Cover title text box (overlaid on cover graphic)
const COVER_X  = 2848377; const COVER_Y  = 5013350;
const COVER_CX = 5061501; const COVER_CY = 457200;

// ─── XML namespace strings ────────────────────────────────────────────────────
const NS_A   = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const NS_R   = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const NS_P   = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const NS_REL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const T_LAY  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const T_IMG  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const T_SLD  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const SLD_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

// ─── Helpers ──────────────────────────────────────────────────────────────────
function xmlEsc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function urlToImageId(url: string): string {
  return url ? url.replace(/[/:]/g, '') : '';
}

function parseDdMmYyyy(date: string): Date | null {
  if (!date || !date.includes('/')) return null;
  const [dd, mm, yyyy] = date.split('/');
  const d = new Date(+yyyy, +mm - 1, +dd);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(d: Date): string {
  return `${d.getDate()}-${d.getMonth() + 1}-${d.getFullYear()}`;
}

function getDateRange(dates: string[]): string {
  const valid = dates
    .map(parseDdMmYyyy)
    .filter((d): d is Date => d !== null)
    .sort((a, b) => a.getTime() - b.getTime());
  if (!valid.length) return '';
  const s = fmtDate(valid[0]);
  const e = fmtDate(valid[valid.length - 1]);
  return s === e ? s : `${s} - ${e}`;
}

function formatDisplayDate(raw: string): string {
  const d = parseDdMmYyyy(raw);
  return d ? fmtDate(d) : raw;
}

// ─── XML element builders ─────────────────────────────────────────────────────
function spText(
  id: number,
  x: number, y: number, cx: number, cy: number,
  text: string,
  opts: { sz?: number; bold?: boolean; italic?: boolean; color?: string; align?: string } = {},
): string {
  const { sz = 1000, bold = false, italic = false, color = '414042', align = 'l' } = opts;
  return (
    `<p:sp>` +
      `<p:nvSpPr><p:cNvPr id="${id}" name="tb${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>` +
        `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
      `</p:spPr>` +
      `<p:txBody>` +
        `<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"><a:spAutoFit/></a:bodyPr>` +
        `<a:lstStyle/>` +
        `<a:p><a:pPr algn="${align}"/>` +
          `<a:r>` +
            `<a:rPr lang="en-ZA" sz="${sz}"${bold ? ' b="1"' : ''}${italic ? ' i="1"' : ''} dirty="0">` +
              `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
              `<a:latin typeface="Arial"/>` +
            `</a:rPr>` +
            `<a:t>${xmlEsc(text)}</a:t>` +
          `</a:r>` +
        `</a:p>` +
      `</p:txBody>` +
    `</p:sp>`
  );
}

function spTitleBar(id: number, store: string): string {
  return (
    `<p:sp>` +
      `<p:nvSpPr><p:cNvPr id="${id}" name="Title 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>` +
        `<a:xfrm><a:off x="${TITLE_X}" y="${TITLE_Y}"/><a:ext cx="${TITLE_CX}" cy="${TITLE_CY}"/></a:xfrm>` +
        `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10442"/></a:avLst></a:prstGeom>` +
        `<a:solidFill><a:schemeClr val="bg1"><a:lumMod val="95000"/></a:schemeClr></a:solidFill>` +
        `<a:ln><a:noFill/></a:ln>` +
      `</p:spPr>` +
      `<p:txBody>` +
        `<a:bodyPr anchor="ctr"/>` +
        `<a:lstStyle>` +
          `<a:lvl1pPr algn="l"><a:defRPr sz="4400"><a:solidFill><a:srgbClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr>` +
        `</a:lstStyle>` +
        `<a:p>` +
          `<a:r>` +
            `<a:rPr lang="en-ZA" sz="1600" b="1" dirty="0">` +
              `<a:solidFill><a:srgbClr val="414042"/></a:solidFill>` +
              `<a:latin typeface="Arial"/>` +
            `</a:rPr>` +
            `<a:t>${xmlEsc(store)}</a:t>` +
          `</a:r>` +
        `</a:p>` +
      `</p:txBody>` +
    `</p:sp>`
  );
}

function spPic(id: number, rId: string, x: number, y: number, cx: number, cy: number): string {
  return (
    `<p:pic>` +
      `<p:nvPicPr>` +
        `<p:cNvPr id="${id}" name="pic${id}"/>` +
        `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
        `<p:nvPr/>` +
      `</p:nvPicPr>` +
      `<p:blipFill>` +
        `<a:blip r:embed="${rId}" cstate="print"/>` +
        `<a:stretch><a:fillRect/></a:stretch>` +
      `</p:blipFill>` +
      `<p:spPr>` +
        `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `</p:spPr>` +
    `</p:pic>`
  );
}

// ─── Build cover slide ────────────────────────────────────────────────────────
function buildCoverSlide(tplXml: string, brand: string, dateRange: string): string {
  const titleXml = (
    `<p:sp>` +
      `<p:nvSpPr><p:cNvPr id="100" name="txtReportName"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr>` +
        `<a:xfrm><a:off x="${COVER_X}" y="${COVER_Y}"/><a:ext cx="${COVER_CX}" cy="${COVER_CY}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
      `</p:spPr>` +
      `<p:txBody>` +
        `<a:bodyPr wrap="square" rtlCol="0" anchor="b"><a:noAutofit/></a:bodyPr>` +
        `<a:lstStyle/>` +
        `<a:p><a:pPr algn="ctr"/>` +
          `<a:r><a:rPr lang="en-ZA" sz="2400" b="1" dirty="0">` +
            `<a:solidFill><a:srgbClr val="414042"/></a:solidFill>` +
            `<a:latin typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>` +
          `</a:rPr><a:t>${xmlEsc(brand.toUpperCase())} STAND REPORT</a:t></a:r>` +
        `</a:p>` +
        `<a:p><a:pPr algn="ctr"/>` +
          `<a:r><a:rPr lang="en-ZA" sz="2400" b="1" dirty="0">` +
            `<a:solidFill><a:srgbClr val="414042"/></a:solidFill>` +
            `<a:latin typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>` +
          `</a:rPr><a:t>(${xmlEsc(dateRange)})</a:t></a:r>` +
        `</a:p>` +
      `</p:txBody>` +
    `</p:sp>`
  );
  return tplXml.replace('</p:spTree>', `${titleXml}</p:spTree>`);
}

// ─── Build content slide ──────────────────────────────────────────────────────
interface ContentSlideParams {
  store:     string;
  repName:   string;
  dateStr:   string;   // formatted for display
  brand:     string;
  img1RId:   string | null;  // null = no image fetched
  img2RId:   string | null;
  hasImg2:   boolean;        // whether a second URL exists (even if not fetched)
}

function buildContentSlideXml(p: ContentSlideParams): string {
  const { store, repName, dateStr, brand, img1RId, img2RId, hasImg2 } = p;
  const isSingle = !hasImg2;
  const img1x = isSingle ? IMG1_1UP_X : IMG1_2UP_X;

  let id = 2;
  const parts: string[] = [];

  // 1. Title bar
  parts.push(spTitleBar(id++, store));

  // 2. Logo 1 (wide – rId2)
  parts.push(spPic(id++, 'rId2', LOGO1_X, LOGO1_Y, LOGO1_CX, LOGO1_CY));

  // 3. Logo 2 (small – rId3)
  parts.push(spPic(id++, 'rId3', LOGO2_X, LOGO2_Y, LOGO2_CX, LOGO2_CY));

  // 4. Image 1 label
  parts.push(spText(id++, img1x, LABEL_Y, LABEL_CX, LABEL_CY,
    `${brand.toUpperCase()} STAND - 1`, { sz: 1000, bold: true }));

  // 5. Image 1
  if (img1RId) {
    parts.push(spPic(id++, img1RId, img1x, IMG_Y, IMG_CX, IMG_CY));
  } else {
    parts.push(spText(id++, img1x, IMG_Y + Math.floor(IMG_CY / 2) - 200000, IMG_CX, 400000,
      'IMAGE UNAVAILABLE', { sz: 1400, color: '999999', align: 'ctr' }));
  }

  // 6. Date below image 1
  parts.push(spText(id++, img1x, DATE_Y, DATE_CX, DATE_CY, dateStr,
    { sz: 900, color: '808080' }));

  if (hasImg2) {
    // 7. Image 2 label
    parts.push(spText(id++, IMG2_2UP_X, LABEL_Y, LABEL_CX, LABEL_CY,
      `${brand.toUpperCase()} STAND - 2`, { sz: 1000, bold: true }));

    // 8. Image 2
    if (img2RId) {
      parts.push(spPic(id++, img2RId, IMG2_2UP_X, IMG_Y, IMG_CX, IMG_CY));
    } else {
      parts.push(spText(id++, IMG2_2UP_X, IMG_Y + Math.floor(IMG_CY / 2) - 200000, IMG_CX, 400000,
        'IMAGE UNAVAILABLE', { sz: 1400, color: '999999', align: 'ctr' }));
    }

    // 9. Date below image 2
    parts.push(spText(id++, IMG2_2UP_X, DATE_Y, DATE_CX, DATE_CY, dateStr,
      { sz: 900, color: '808080' }));
  }

  // 10. Rep name (bottom-right, small)
  if (repName) {
    parts.push(spText(id++, REP_X, REP_Y, REP_CX, REP_CY, repName,
      { sz: 700, color: '999999', align: 'r' }));
  }

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld ${NS_A} ${NS_R} ${NS_P}>` +
      `<p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm>` +
          `<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
          `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>` +
        `</a:xfrm></p:grpSpPr>` +
        parts.join('') +
      `</p:spTree></p:cSld>` +
      `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
    `</p:sld>`
  );
}

function buildContentSlideRels(
  img1Filename: string | null,
  img2Filename: string | null,
): string {
  let rel = (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships ${NS_REL}>` +
      `<Relationship Id="rId1" Type="${T_LAY}" Target="../slideLayouts/slideLayout4.xml"/>` +
      `<Relationship Id="rId2" Type="${T_IMG}" Target="../media/image10.png"/>` +
      `<Relationship Id="rId3" Type="${T_IMG}" Target="../media/image11.png"/>`
  );
  if (img1Filename) rel += `<Relationship Id="rId4" Type="${T_IMG}" Target="../media/${img1Filename}"/>`;
  if (img2Filename) rel += `<Relationship Id="rId5" Type="${T_IMG}" Target="../media/${img2Filename}"/>`;
  rel += `</Relationships>`;
  return rel;
}

// ─── Presentation XML helpers ─────────────────────────────────────────────────
function buildPresRels(tplRels: string, slideFiles: string[]): string {
  // Remove existing slide Relationship entries
  const stripped = tplRels.replace(/<Relationship[^>]+Type="[^"]*\/slide"[^>]*\/>/g, '');
  // Insert new slide entries before </Relationships>
  const slideRels = slideFiles
    .map((file, i) => {
      const rId = `rId_s${i + 1}`;
      return `<Relationship Id="${rId}" Type="${T_SLD}" Target="slides/${file}"/>`;
    })
    .join('');
  return stripped.replace('</Relationships>', `${slideRels}</Relationships>`);
}

function buildPresXml(tplXml: string, slideCount: number): string {
  // Build new sldIdLst (cover + content + closing = slideCount slides)
  let sldIdLst = '<p:sldIdLst>';
  for (let i = 0; i < slideCount; i++) {
    sldIdLst += `<p:sldId id="${256 + i}" r:id="rId_s${i + 1}"/>`;
  }
  sldIdLst += '</p:sldIdLst>';
  // Replace existing sldIdLst
  return tplXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, sldIdLst);
}

function buildContentTypes(tplCt: string, slideFiles: string[]): string {
  // Remove existing slide Override entries
  const stripped = tplCt.replace(
    /<Override PartName="\/ppt\/slides\/[^"]*" ContentType="[^"]*slide\+xml"[^/]*\/>/g, '',
  );
  // Insert new ones before </Types>
  const overrides = slideFiles
    .map(f => `<Override PartName="/ppt/slides/${f}" ContentType="${SLD_CT}"/>`)
    .join('');
  return stripped.replace('</Types>', `${overrides}</Types>`);
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generateStandReport(
  fileBuffer: Buffer,
  brand: string,
): Promise<{ buffer: Buffer; filename: string; rawDates: string[] }> {

  // ── 1. Parse Excel ─────────────────────────────────────────────────────────
  const wb = XLSX.read(fileBuffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (unknown)[][];
  if (rawData.length < 2) throw new Error('No data rows found in uploaded file.');

  // ── 2. Map rows ─────────────────────────────────────────────────────────────
  const rows = rawData.slice(1).map(raw => {
    const r = raw as (string | number | null)[];
    const imageUrls: string[] = [];
    r.forEach(cell => {
      const s = String(cell ?? '').trim();
      if (s.startsWith('http')) imageUrls.push(s);
    });
    return {
      repName:   [String(r[2] ?? '').trim(), String(r[3] ?? '').trim()].filter(Boolean).join(' ') || 'UNKNOWN',
      store:     String(r[6] ?? '').trim() || 'UNKNOWN',
      date:      String(r[9] ?? '').trim(),
      imageUrls,
      imageIds:  imageUrls.map(urlToImageId),
    };
  });

  if (rows.length === 0) throw new Error('No data rows found in uploaded file.');

  const rawDates = rows.map(r => r.date).filter(d => d.includes('/'));
  const dateRange = getDateRange(rawDates);

  // ── 3. Build slide groups (max 2 images per slide per row) ─────────────────
  interface SlideGroup { store: string; repName: string; date: string; imageIds: string[] }
  const slideGroups: SlideGroup[] = [];
  for (const row of rows) {
    if (row.imageIds.length === 0) continue;
    for (let i = 0; i < row.imageIds.length; i += 2) {
      slideGroups.push({
        store:    row.store,
        repName:  row.repName,
        date:     row.date,
        imageIds: row.imageIds.slice(i, i + 2),
      });
    }
  }

  if (slideGroups.length === 0) throw new Error('No image URLs found in uploaded file.');

  // ── 4. Fetch images from SharePoint ────────────────────────────────────────
  const BASE_PATH = (process.env.DFE_SP_BASE_PATH || 'DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS').trim();
  const appSettings = loadAppSettings();
  const PICTURES_FOLDER = (
    appSettings.picturesFolderPath ||
    process.env.DFE_PICTURES_SP_PATH ||
    `${BASE_PATH}/PERIGEE IMAGE DOWNLOADS`
  ).trim();

  const allIds = [...new Set(slideGroups.flatMap(g => g.imageIds).filter(Boolean))];
  const fetched = await Promise.all(allIds.map(id => downloadFileFromSP(PICTURES_FOLDER, `${id}.jpg`)));
  const imageBufs = new Map<string, Buffer>();
  allIds.forEach((id, i) => { if (fetched[i]) imageBufs.set(id, fetched[i]!); });

  console.log(`[stand-report] SP images: ${imageBufs.size}/${allIds.length} fetched from "${PICTURES_FOLDER}"`);

  // ── 5. Load template PPTX ──────────────────────────────────────────────────
  const tplPath = path.join(process.cwd(), 'public', 'PPT Template AM.pptx');
  const tplBuf  = fs.readFileSync(tplPath);
  const tplZip  = await JSZip.loadAsync(tplBuf);

  const tplSlide1Xml  = await tplZip.file('ppt/slides/slide1.xml')!.async('string');
  const tplSlide1Rels = await tplZip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
  const tplSlide3Xml  = await tplZip.file('ppt/slides/slide3.xml')!.async('string');
  const tplSlide3Rels = await tplZip.file('ppt/slides/_rels/slide3.xml.rels')!.async('string');
  const tplPresXml    = await tplZip.file('ppt/presentation.xml')!.async('string');
  const tplPresRels   = await tplZip.file('ppt/_rels/presentation.xml.rels')!.async('string');
  const tplCtXml      = await tplZip.file('[Content_Types].xml')!.async('string');

  // ── 6. Build output zip ────────────────────────────────────────────────────
  const outZip = new JSZip();

  // Copy all template files except slides & manifest files we'll rebuild
  const skipSet = new Set([
    'ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml',
    'ppt/slides/_rels/slide1.xml.rels', 'ppt/slides/_rels/slide2.xml.rels',
    'ppt/slides/_rels/slide3.xml.rels',
    'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels',
    '[Content_Types].xml',
  ]);
  for (const [name, file] of Object.entries(tplZip.files)) {
    if (file.dir || skipSet.has(name)) continue;
    outZip.file(name, await file.async('nodebuffer'));
  }

  // ── 7. Slide order: cover + content slides + closing ───────────────────────
  // slide1.xml = cover
  // slide2.xml…slide(N+1).xml = content
  // slide(N+2).xml = closing
  const N = slideGroups.length;
  const allSlideFiles: string[] = [];

  // Cover
  allSlideFiles.push('slide1.xml');
  outZip.file('ppt/slides/slide1.xml', buildCoverSlide(tplSlide1Xml, brand, dateRange));
  outZip.file('ppt/slides/_rels/slide1.xml.rels', tplSlide1Rels);

  // Content slides
  let globalImgCounter = 1;
  for (let i = 0; i < N; i++) {
    const g        = slideGroups[i];
    const slideNum = i + 2;
    const id1      = g.imageIds[0] ?? '';
    const id2      = g.imageIds[1] ?? '';
    const buf1     = id1 ? imageBufs.get(id1) : undefined;
    const buf2     = id2 ? imageBufs.get(id2) : undefined;
    const hasImg2  = g.imageIds.length > 1;

    let mediaName1: string | null = null;
    let mediaName2: string | null = null;

    if (buf1) {
      mediaName1 = `stand_img_${globalImgCounter++}.jpg`;
      outZip.file(`ppt/media/${mediaName1}`, buf1);
    }
    if (buf2) {
      mediaName2 = `stand_img_${globalImgCounter++}.jpg`;
      outZip.file(`ppt/media/${mediaName2}`, buf2);
    }

    const img1RId = buf1 ? 'rId4' : null;
    const img2RId = buf2 ? 'rId5' : null;

    const slideFile     = `slide${slideNum}.xml`;
    const slideRelsFile = `_rels/slide${slideNum}.xml.rels`;

    outZip.file(`ppt/slides/${slideFile}`, buildContentSlideXml({
      store:    g.store,
      repName:  g.repName,
      dateStr:  formatDisplayDate(g.date),
      brand,
      img1RId,
      img2RId,
      hasImg2,
    }));
    outZip.file(`ppt/slides/${slideRelsFile}`, buildContentSlideRels(mediaName1, mediaName2));
    allSlideFiles.push(slideFile);
  }

  // Closing slide
  const closingNum = N + 2;
  const closingFile = `slide${closingNum}.xml`;
  outZip.file(`ppt/slides/${closingFile}`, tplSlide3Xml);
  // Rewrite closing rels to update slide layout target (keep relative paths the same)
  outZip.file(`ppt/slides/_rels/${closingFile}.rels`, tplSlide3Rels);
  allSlideFiles.push(closingFile);

  // ── 8. Update presentation.xml, .rels, [Content_Types].xml ─────────────────
  outZip.file('ppt/presentation.xml',              buildPresXml(tplPresXml, allSlideFiles.length));
  outZip.file('ppt/_rels/presentation.xml.rels',   buildPresRels(tplPresRels, allSlideFiles));
  outZip.file('[Content_Types].xml',               buildContentTypes(tplCtXml, allSlideFiles));

  // ── 9. Filename & output ───────────────────────────────────────────────────
  const today = new Date();
  const dd    = String(today.getDate()).padStart(2, '0');
  const mm    = String(today.getMonth() + 1).padStart(2, '0');
  const yyyy  = today.getFullYear();
  const filename = `${brand.toUpperCase()}_STAND_REPORT_${yyyy}${mm}${dd}.pptx`;

  const buffer = Buffer.from(await outZip.generateAsync({
    type:       'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  }));

  return { buffer, filename, rawDates };
}

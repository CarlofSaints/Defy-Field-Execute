/**
 * Defy Field Execute — SharePoint folder path helpers
 *
 * Required Vercel env vars:
 *   OJ_TENANT_ID, OJ_CLIENT_ID, OJ_CLIENT_SECRET  (already used by graph-oj.ts)
 *   OJ_SP_HOST     = exceler8xl.sharepoint.com
 *   OJ_SP_LIBRARY  = Clients
 *   DFE_SP_BASE_PATH = DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS
 *
 * Folder convention:
 *   {BASE_PATH}/{Brand}/{MM}- {MON}{YY}/Week Ending - {YYYYMMDD}/
 *
 * Examples:
 *   DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS/DEFY/01- JAN26/Week Ending - 20260120
 *   DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS/GRUNDIG/03- MAR26/Week Ending - 20260313
 */

export type DfeBrand = 'DEFY' | 'GRUNDIG' | 'BEKO';

const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

const BASE_PATH = (
  process.env.DFE_SP_BASE_PATH || 'DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS'
).trim();

/**
 * Parse a Perigee date value — handles:
 *   - DD/MM/YYYY string (Perigee export format)
 *   - Excel serial number
 */
function parsePerigeeDate(val: unknown): Date | null {
  if (typeof val === 'string') {
    const m = val.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }
  if (typeof val === 'number' && val > 0) {
    // Excel serial: days since 1900-01-01 (with the Lotus 1-2-3 leap year bug offset)
    return new Date((val - 25569) * 86400 * 1000);
  }
  return null;
}

/**
 * Format month folder name: "01- JAN26"
 */
export function getMonthFolder(date: Date): string {
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const mon = MONTHS[date.getMonth()];
  const yy = String(date.getFullYear()).slice(-2);
  return `${mm}- ${mon}${yy}`;
}

/**
 * Format week-ending folder name: "Week Ending - 20260120"
 */
export function getWeekEndingFolder(date: Date): string {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `Week Ending - ${yyyy}${mm}${dd}`;
}

/**
 * Build the full SharePoint folder path for a report.
 *
 * @param brand     - 'DEFY' | 'GRUNDIG' | 'BEKO'
 * @param rawDates  - array of raw date values from Perigee data (DD/MM/YYYY strings or Excel serials)
 * @returns         - folder path string relative to the Clients document library root
 *
 * @throws if no valid dates found in rawDates
 */
export function buildDfeFolderPath(brand: DfeBrand, rawDates: unknown[]): string {
  const dates = rawDates
    .map(parsePerigeeDate)
    .filter((d): d is Date => d !== null);

  if (dates.length === 0) {
    throw new Error('buildDfeFolderPath: no valid dates found in rawDates array');
  }

  const latest = new Date(Math.max(...dates.map(d => d.getTime())));

  return `${BASE_PATH}/${brand}/${getMonthFolder(latest)}/${getWeekEndingFolder(latest)}`;
}

// Re-export the upload function so report routes only need to import from here
export { uploadToSharePoint } from './graph-oj';

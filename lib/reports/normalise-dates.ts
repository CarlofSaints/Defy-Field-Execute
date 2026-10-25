import * as XLSX from 'xlsx';

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;

/**
 * Field Goose exports dates as "2026-10-18"; Perigee exported "18/10/2026".
 * Every report, the week number and the SharePoint folder parse DD/MM/YYYY,
 * so rewrite ISO date cells once on upload. A file with none is returned untouched.
 */
export function normaliseUploadDates(fileBuffer: Buffer): Buffer {
  const wb = XLSX.read(fileBuffer, { type: 'buffer' });
  let changed = false;
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    for (const addr of Object.keys(ws)) {
      if (addr.startsWith('!')) continue;
      const cell = ws[addr] as XLSX.CellObject;
      const m = cell.t === 's' && typeof cell.v === 'string' ? cell.v.trim().match(ISO_DATE) : null;
      if (!m) continue;
      ws[addr] = { t: 's', v: `${m[3]}/${m[2]}/${m[1]}` };
      changed = true;
    }
  }
  return changed ? Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })) : fileBuffer;
}

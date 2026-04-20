import * as XLSX from 'xlsx';

export interface ChannelMismatchResult {
  expected:      string;
  mismatches:    Record<string, number>;
  totalMismatch: number;
}

/**
 * Scan an Excel buffer for rows whose "Channel" column doesn't match the
 * expected channel. Returns null when there is no mismatch (or no Channel
 * column found).
 */
export function analyzeChannelMismatch(
  fileBuffer: Buffer,
  expectedChannel: string,
): ChannelMismatchResult | null {
  const wb = XLSX.read(fileBuffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (string | number | null)[][];

  if (rows.length < 2) return null;

  const headers = rows[0] as (string | null)[];
  const channelIdx = headers.findIndex(
    h => typeof h === 'string' && h.trim().toLowerCase() === 'channel',
  );
  if (channelIdx === -1) return null;

  const expected = expectedChannel.trim().toLowerCase();
  const mismatches: Record<string, number> = {};
  let totalMismatch = 0;

  for (let i = 1; i < rows.length; i++) {
    const raw = String(rows[i][channelIdx] ?? '').trim();
    if (!raw) continue; // blank channel — keep
    if (raw.toLowerCase() === expected) continue;
    mismatches[raw.toUpperCase()] = (mismatches[raw.toUpperCase()] || 0) + 1;
    totalMismatch++;
  }

  return totalMismatch > 0 ? { expected: expectedChannel, mismatches, totalMismatch } : null;
}

/**
 * Return a new Excel buffer with mismatched-channel rows removed.
 * Blank-channel rows are kept.
 */
export function filterByChannel(
  fileBuffer: Buffer,
  expectedChannel: string,
): Buffer {
  const wb = XLSX.read(fileBuffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (string | number | null)[][];

  if (rows.length < 2) return fileBuffer;

  const headers = rows[0] as (string | null)[];
  const channelIdx = headers.findIndex(
    h => typeof h === 'string' && h.trim().toLowerCase() === 'channel',
  );
  if (channelIdx === -1) return fileBuffer; // no column to filter on

  const expected = expectedChannel.trim().toLowerCase();
  const filtered = [
    rows[0], // header
    ...rows.slice(1).filter(row => {
      const raw = String(row[channelIdx] ?? '').trim();
      return !raw || raw.toLowerCase() === expected;
    }),
  ];

  const newWs = XLSX.utils.aoa_to_sheet(filtered);
  const newWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWb, newWs, wb.SheetNames[0]);
  return Buffer.from(XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' }));
}

import fs from 'fs';
import path from 'path';

export interface ReportDef {
  id: string;
  name: string;
  dataFormat: string;    // which builder to use: stock-count, red-flag, stand-report, etc.
  channel?: string;      // optional retailer/channel: MAKRO, GAME, etc.
  outputTypes: string[]; // 'Excel' | 'PPT'
  brands: string[];      // 'Defy' | 'Beko'
}

/** All available data formats — maps to report builders in generate route */
export const DATA_FORMATS = [
  { value: 'stock-count',        label: 'Stock Count' },
  { value: 'red-flag',           label: 'Red Flag' },
  { value: 'stand-report',       label: 'Stand Report' },
  { value: 'service-call',       label: 'Service Call' },
  { value: 'training-feedback',  label: 'Training Feedback' },
  { value: 'activation-report',  label: 'Activation Report' },
] as const;

/** Human-readable label for a data format value */
export const DATA_FORMAT_LABELS: Record<string, string> = Object.fromEntries(
  DATA_FORMATS.map(f => [f.value, f.label]),
);

// Persisted in Vercel Blob on the server (the same store the upload archive and
// control files use), bundled JSON file in dev / as the seed. Existing report
// definitions in the bundled file or legacy env var are read until the first
// Blob save, so nothing is lost by the switch.
const FILE       = path.join(process.cwd(), 'data', 'reports.json');
const BLOB_KEY   = 'config/reports.json';
const VERCEL_KEY = 'DFE_REPORTS_JSON';
const useBlob    = !!process.env.BLOB_READ_WRITE_TOKEN;

export async function loadReports(): Promise<ReportDef[]> {
  if (useBlob) {
    const { list } = await import('@vercel/blob');
    const listing = await list({ prefix: BLOB_KEY });
    const match = listing.blobs.find(b => b.pathname === BLOB_KEY) ?? listing.blobs[0];
    if (match) {
      const res = await fetch(match.url, { cache: 'no-store' });
      if (res.ok) return await res.json() as ReportDef[];
    }
    // No Blob yet — fall through to the bundled file / legacy env var as seed.
  }

  // Bundled file is the canonical seed (committed to the repo).
  if (fs.existsSync(FILE)) return JSON.parse(fs.readFileSync(FILE, 'utf-8'));
  const env = process.env[VERCEL_KEY];
  if (env)                 return JSON.parse(env);
  return [];
}

export async function saveReports(reports: ReportDef[]): Promise<void> {
  if (useBlob) {
    const { put } = await import('@vercel/blob');
    await put(BLOB_KEY, JSON.stringify(reports), {
      access:          'public',
      contentType:     'application/json',
      addRandomSuffix: false,
    });
    return;
  }

  const dir = path.dirname(FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(FILE, JSON.stringify(reports, null, 2));
}

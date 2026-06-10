import fs from 'fs';
import path from 'path';

export interface RunLogEntry {
  id:               string;
  timestamp:        string;   // ISO 8601
  userName:         string;
  userEmail:        string;
  reportId:         string;
  reportName:       string;
  brand:            string;
  retailer:         string;
  filename:         string;
  status:           'success' | 'error';
  errorMessage?:    string;
  spPath?:          string;   // SharePoint file URL (set after upload)
  emailSent?:       boolean;
  emailRecipients?: string[];
}

// Persisted in Vercel Blob on the server (the same store the upload archive and
// control files use), local JSON file in dev. Existing history in the bundled
// file or legacy env var is read until the first Blob save.
const MAX_ENTRIES = 500;
const FILE        = path.join(process.cwd(), 'data', 'runLog.json');
const BLOB_KEY    = 'config/run-log.json';
const VERCEL_KEY  = 'DFE_RUN_LOG_JSON';
const useBlob     = !!process.env.BLOB_READ_WRITE_TOKEN;

export async function loadRunLog(): Promise<RunLogEntry[]> {
  if (useBlob) {
    const { list } = await import('@vercel/blob');
    const listing = await list({ prefix: BLOB_KEY });
    const match = listing.blobs.find(b => b.pathname === BLOB_KEY) ?? listing.blobs[0];
    if (match) {
      const res = await fetch(match.url, { cache: 'no-store' });
      if (res.ok) return await res.json() as RunLogEntry[];
    }
    // No Blob yet — fall through to legacy sources.
  }

  const env = process.env[VERCEL_KEY];
  if (process.env.VERCEL && env) return JSON.parse(env);
  if (fs.existsSync(FILE))       return JSON.parse(fs.readFileSync(FILE, 'utf-8'));
  if (env)                       return JSON.parse(env);
  return [];
}

export async function addRunEntry(entry: RunLogEntry): Promise<void> {
  const log = await loadRunLog();
  // Newest first; trim to max
  log.unshift(entry);
  if (log.length > MAX_ENTRIES) log.splice(MAX_ENTRIES);
  await saveRunLog(log);
}

async function saveRunLog(log: RunLogEntry[]): Promise<void> {
  if (useBlob) {
    const { put } = await import('@vercel/blob');
    await put(BLOB_KEY, JSON.stringify(log), {
      access:          'public',
      contentType:     'application/json',
      addRandomSuffix: false,
    });
    return;
  }

  const dir = path.dirname(FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(FILE, JSON.stringify(log, null, 2));
}

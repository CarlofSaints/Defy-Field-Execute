import fs from 'fs';
import path from 'path';

// Critical-line products per brand + channel. Persisted in Vercel Blob on the
// server (the same store the upload archive and product catalog use), local
// JSON file in dev. The earlier env-var store needed VERCEL_TOKEN and only
// reflected changes after a redeploy (env vars are baked at deploy time), so
// uploads didn't stick.

export interface CriticalLineEntry {
  modelNumber:        string;
  productDescription: string;
}

export interface CriticalLinesConfig {
  brand:   string;   // e.g. "DEFY" or "BEKO"
  channel: string;   // e.g. "MAKRO", "GAME", "INDEPENDENTS", "TAFELBERG", "HIRSCHS"
  lines:   CriticalLineEntry[];
}

const FILE     = path.join(process.cwd(), 'data', 'criticalLines.json');
const BLOB_KEY = 'config/critical-lines.json';
const useBlob  = !!process.env.BLOB_READ_WRITE_TOKEN;

export async function loadCriticalLines(): Promise<CriticalLinesConfig[]> {
  if (useBlob) {
    const { list } = await import('@vercel/blob');
    const listing = await list({ prefix: BLOB_KEY });
    const match = listing.blobs.find(b => b.pathname === BLOB_KEY) ?? listing.blobs[0];
    if (!match) return [];
    const res = await fetch(match.url, { cache: 'no-store' });
    if (!res.ok) return [];
    return await res.json() as CriticalLinesConfig[];
  }

  if (fs.existsSync(FILE)) {
    return JSON.parse(fs.readFileSync(FILE, 'utf-8'));
  }

  return [];
}

/**
 * Returns a Set of uppercase product codes that are critical lines for the
 * given brand + channel combination. Spaces are stripped from model numbers
 * to match Perigee product codes.
 */
export async function getCriticalLineSet(brand: string, channel: string): Promise<Set<string>> {
  const all = await loadCriticalLines();
  const b = brand.toUpperCase();
  const c = channel.toUpperCase();
  const config = all.find(
    cfg => cfg.brand.toUpperCase() === b && cfg.channel.toUpperCase() === c,
  );
  if (!config) return new Set();
  return new Set(config.lines.map(l => l.modelNumber.toUpperCase().replace(/\s+/g, '')));
}

export async function saveCriticalLines(configs: CriticalLinesConfig[]): Promise<void> {
  if (useBlob) {
    const { put } = await import('@vercel/blob');
    await put(BLOB_KEY, JSON.stringify(configs), {
      access:          'public',
      contentType:     'application/json',
      addRandomSuffix: false,
    });
    return;
  }

  const dir = path.dirname(FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(FILE, JSON.stringify(configs, null, 2));
}

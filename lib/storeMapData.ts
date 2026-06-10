import fs from 'fs';
import path from 'path';
import { blobEnabled } from './blobStore';

// Store → province mapping. Persisted in Vercel Blob on the server (the same
// store the upload archive and Product Management File use), local JSON file in dev.
// The earlier env-var store needed VERCEL_TOKEN and only reflected changes
// after a redeploy (env vars are baked at deploy time), so uploads didn't stick.

export interface StoreMapEntry {
  storeName: string;
  storeCode: string;
  province:  string;
}

const FILE       = path.join(process.cwd(), 'data', 'storeMap.json');
const BLOB_KEY   = 'config/store-map.json';
const VERCEL_KEY = 'DFE_STORE_MAP_JSON';
const useBlob    = blobEnabled;

export async function loadStoreMap(): Promise<StoreMapEntry[]> {
  if (useBlob) {
    const { list } = await import('@vercel/blob');
    const listing = await list({ prefix: BLOB_KEY });
    const match = listing.blobs.find(b => b.pathname === BLOB_KEY) ?? listing.blobs[0];
    if (match) {
      const res = await fetch(match.url, { cache: 'no-store' });
      if (res.ok) return await res.json() as StoreMapEntry[];
    }
    // No Blob yet — fall through to legacy sources so existing data survives
    // until the first upload writes to Blob.
  }

  const env = process.env[VERCEL_KEY];
  if (process.env.VERCEL && env) return JSON.parse(env);
  if (fs.existsSync(FILE))       return JSON.parse(fs.readFileSync(FILE, 'utf-8'));
  if (env)                       return JSON.parse(env);
  return [];
}

export async function saveStoreMap(entries: StoreMapEntry[]): Promise<void> {
  if (useBlob) {
    const { put } = await import('@vercel/blob');
    await put(BLOB_KEY, JSON.stringify(entries), {
      access:          'public',
      contentType:     'application/json',
      addRandomSuffix: false,
    });
    return;
  }

  const dir = path.dirname(FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(FILE, JSON.stringify(entries, null, 2));
}

import fs from 'fs';
import path from 'path';

export interface AppSettings {
  /**
   * SP path (or full URL) to the folder containing Perigee visit images
   * shared by the Red Flag and Stand Report builders. Empty = use default.
   */
  picturesFolderPath: string;
  /**
   * SP path (or full URL) to the folder containing Activation campaign
   * images used by the Activation Report builder. Empty = use default.
   */
  activationPicturesFolderPath: string;
}

const DEFAULT_SETTINGS: AppSettings = {
  picturesFolderPath: '',
  activationPicturesFolderPath: '',
};

// Persisted in Vercel Blob on the server (the same store the upload archive and
// control files use), local JSON file in dev. Existing settings in the legacy
// env var are still read until the first Blob save, so saved paths survive the
// switch.
const FILE       = path.join(process.cwd(), 'data', 'appSettings.json');
const BLOB_KEY   = 'config/app-settings.json';
const VERCEL_KEY = 'DFE_APP_SETTINGS_JSON';
const useBlob    = !!process.env.BLOB_READ_WRITE_TOKEN;

export async function loadAppSettings(): Promise<AppSettings> {
  if (useBlob) {
    const { list } = await import('@vercel/blob');
    const listing = await list({ prefix: BLOB_KEY });
    const match = listing.blobs.find(b => b.pathname === BLOB_KEY) ?? listing.blobs[0];
    if (match) {
      const res = await fetch(match.url, { cache: 'no-store' });
      if (res.ok) return { ...DEFAULT_SETTINGS, ...(await res.json()) };
    }
    // No Blob yet — fall through to legacy sources.
  }

  const env = process.env[VERCEL_KEY];
  if (process.env.VERCEL && env) return { ...DEFAULT_SETTINGS, ...JSON.parse(env) };
  if (fs.existsSync(FILE))       return { ...DEFAULT_SETTINGS, ...JSON.parse(fs.readFileSync(FILE, 'utf-8')) };
  if (env)                       return { ...DEFAULT_SETTINGS, ...JSON.parse(env) };
  return { ...DEFAULT_SETTINGS };
}

export async function saveAppSettings(settings: AppSettings): Promise<void> {
  if (useBlob) {
    const { put } = await import('@vercel/blob');
    await put(BLOB_KEY, JSON.stringify(settings), {
      access:          'public',
      contentType:     'application/json',
      addRandomSuffix: false,
    });
    return;
  }

  const dir = path.dirname(FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(FILE, JSON.stringify(settings, null, 2));
}

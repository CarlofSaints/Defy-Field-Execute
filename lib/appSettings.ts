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

const FILE      = path.join(process.cwd(), 'data', 'appSettings.json');
const VERCEL_KEY = 'DFE_APP_SETTINGS_JSON';
const PROJECT_ID = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';

let _cache: AppSettings | null = null;

export function loadAppSettings(): AppSettings {
  if (_cache !== null) return _cache;

  const env = process.env[VERCEL_KEY];
  if (process.env.VERCEL && env) {
    _cache = { ...DEFAULT_SETTINGS, ...JSON.parse(env) };
    return _cache!;
  }

  if (fs.existsSync(FILE)) {
    _cache = { ...DEFAULT_SETTINGS, ...JSON.parse(fs.readFileSync(FILE, 'utf-8')) };
    return _cache!;
  }

  if (env) {
    _cache = { ...DEFAULT_SETTINGS, ...JSON.parse(env) };
    return _cache!;
  }

  _cache = { ...DEFAULT_SETTINGS };
  return _cache!;
}

export async function saveAppSettings(settings: AppSettings) {
  _cache = settings;

  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(settings, null, 2));
    return;
  } catch {
    // Vercel read-only FS — fall through to Vercel API
  }

  const token = process.env.VERCEL_TOKEN;
  if (!token) return;

  const listRes = await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!listRes.ok) return;

  const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
  const existing = envs.find(e => e.key === VERCEL_KEY);
  const value    = JSON.stringify(settings);

  if (!existing) {
    await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        key: VERCEL_KEY,
        value,
        type: 'plain',
        target: ['production', 'preview', 'development'],
      }),
    });
  } else {
    await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env/${existing.id}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ value }),
    });
  }
}

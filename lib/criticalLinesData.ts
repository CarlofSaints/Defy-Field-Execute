import fs from 'fs';
import path from 'path';

export interface CriticalLineEntry {
  modelNumber:        string;
  productDescription: string;
}

export interface CriticalLinesConfig {
  brand:   string;   // e.g. "DEFY" or "BEKO"
  channel: string;   // e.g. "MAKRO", "GAME", "INDEPENDENTS", "TAFELBERG", "HIRSCHS"
  lines:   CriticalLineEntry[];
}

const FILE       = path.join(process.cwd(), 'data', 'criticalLines.json');
const VERCEL_KEY = 'DFE_CRITICAL_LINES_JSON';
const PROJECT_ID = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';

let _cache: CriticalLinesConfig[] | null = null;

export function loadCriticalLines(): CriticalLinesConfig[] {
  if (_cache !== null) return _cache;

  const env = process.env[VERCEL_KEY];
  if (process.env.VERCEL && env) {
    _cache = JSON.parse(env);
    return _cache!;
  }

  if (fs.existsSync(FILE)) {
    _cache = JSON.parse(fs.readFileSync(FILE, 'utf-8'));
    return _cache!;
  }

  if (env) {
    _cache = JSON.parse(env);
    return _cache!;
  }

  return [];
}

/**
 * Returns a Set of uppercase product codes that are critical lines
 * for the given brand + channel combination.
 * Spaces are stripped from model numbers to match Perigee product codes.
 */
export function getCriticalLineSet(brand: string, channel: string): Set<string> {
  const all = loadCriticalLines();
  const b = brand.toUpperCase();
  const c = channel.toUpperCase();
  const config = all.find(
    cfg => cfg.brand.toUpperCase() === b && cfg.channel.toUpperCase() === c,
  );
  if (!config) return new Set();
  return new Set(config.lines.map(l => l.modelNumber.toUpperCase().replace(/\s+/g, '')));
}

export async function saveCriticalLines(configs: CriticalLinesConfig[]) {
  _cache = configs;

  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(configs, null, 2));
    return;
  } catch {
    // Vercel read-only FS — fall through
  }

  const token = process.env.VERCEL_TOKEN;
  if (!token) return;

  const listRes = await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!listRes.ok) return;

  const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
  const existing = envs.find(e => e.key === VERCEL_KEY);
  const value    = JSON.stringify(configs);

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

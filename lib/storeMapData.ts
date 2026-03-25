import fs from 'fs';
import path from 'path';

export interface StoreMapEntry {
  storeName: string;
  storeCode: string;
  province:  string;
}

const FILE = path.join(process.cwd(), 'data', 'storeMap.json');
let _cache: StoreMapEntry[] | null = null;

export function loadStoreMap(): StoreMapEntry[] {
  if (_cache !== null) return _cache;

  const env = process.env.DFE_STORE_MAP_JSON;
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

export async function saveStoreMap(entries: StoreMapEntry[]) {
  _cache = entries;

  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(entries, null, 2));
    return;
  } catch {
    // Vercel read-only FS — fall through
  }

  const token = process.env.VERCEL_TOKEN;
  const projectId = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';
  if (token) {
    await upsertVercelEnvVar(token, projectId, entries);
  }
}

async function upsertVercelEnvVar(token: string, projectId: string, entries: StoreMapEntry[]) {
  const listRes = await fetch(`https://api.vercel.com/v9/projects/${projectId}/env`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!listRes.ok) return;

  const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
  const envRecord = envs.find(e => e.key === 'DFE_STORE_MAP_JSON');
  const value = JSON.stringify(entries);

  if (!envRecord) {
    // First upload — create the env var
    await fetch(`https://api.vercel.com/v9/projects/${projectId}/env`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        key: 'DFE_STORE_MAP_JSON',
        value,
        type: 'plain',
        target: ['production', 'preview', 'development'],
      }),
    });
    return;
  }

  await fetch(`https://api.vercel.com/v9/projects/${projectId}/env/${envRecord.id}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ value }),
  });
}

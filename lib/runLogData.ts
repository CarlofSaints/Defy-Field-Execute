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

const MAX_ENTRIES = 500;
const FILE        = path.join(process.cwd(), 'data', 'runLog.json');
let _cache: RunLogEntry[] | null = null;

export function loadRunLog(): RunLogEntry[] {
  if (_cache !== null) return _cache;

  const env = process.env.DFE_RUN_LOG_JSON;
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

export async function addRunEntry(entry: RunLogEntry): Promise<void> {
  const log = loadRunLog();
  // Newest first; trim to max
  log.unshift(entry);
  if (log.length > MAX_ENTRIES) log.splice(MAX_ENTRIES);
  await saveRunLog(log);
}

async function saveRunLog(log: RunLogEntry[]) {
  _cache = log;

  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(log, null, 2));
    return;
  } catch {
    // Vercel read-only FS — fall through
  }

  const token     = process.env.VERCEL_TOKEN;
  const projectId = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';
  if (token) await upsertVercelEnvVar(token, projectId, log);
}

async function upsertVercelEnvVar(token: string, projectId: string, log: RunLogEntry[]) {
  const listRes = await fetch(`https://api.vercel.com/v9/projects/${projectId}/env`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!listRes.ok) return;

  const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
  const envRecord = envs.find(e => e.key === 'DFE_RUN_LOG_JSON');
  const value     = JSON.stringify(log);

  if (!envRecord) {
    await fetch(`https://api.vercel.com/v9/projects/${projectId}/env`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        key: 'DFE_RUN_LOG_JSON',
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

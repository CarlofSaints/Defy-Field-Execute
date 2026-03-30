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

const FILE = path.join(process.cwd(), 'data', 'reports.json');
let _cache: ReportDef[] | null = null;

export function loadReports(): ReportDef[] {
  if (_cache !== null) return _cache;

  // Always prefer the bundled file — it is readable on Vercel serverless functions
  // and is the canonical source of truth (env var is stale after Admin Centre changes).
  if (fs.existsSync(FILE)) {
    _cache = JSON.parse(fs.readFileSync(FILE, 'utf-8'));
    return _cache!;
  }

  // Fallback for edge cases where the file isn't bundled
  const env = process.env.DFE_REPORTS_JSON;
  if (env) {
    _cache = JSON.parse(env);
    return _cache!;
  }

  return [];
}

export async function saveReports(reports: ReportDef[]) {
  _cache = reports;

  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(reports, null, 2));
    return;
  } catch {
    // Vercel read-only FS — fall through
  }

  const token = process.env.VERCEL_TOKEN;
  const projectId = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';
  if (token) {
    await updateVercelEnvVar(token, projectId, reports);
  }
}

async function updateVercelEnvVar(token: string, projectId: string, reports: ReportDef[]) {
  const listRes = await fetch(`https://api.vercel.com/v9/projects/${projectId}/env`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!listRes.ok) return;
  const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
  const envRecord = envs.find(e => e.key === 'DFE_REPORTS_JSON');
  if (!envRecord) return;

  await fetch(`https://api.vercel.com/v9/projects/${projectId}/env/${envRecord.id}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ value: JSON.stringify(reports) }),
  });
}

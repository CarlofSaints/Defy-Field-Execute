import fs from 'fs';
import path from 'path';

export interface User {
  id: string;
  name: string;
  email: string;
  password: string;
  isAdmin: boolean;
  forcePasswordChange: boolean;
  firstLoginAt: string | null;
  createdAt: string;
}

const FILE = path.join(process.cwd(), 'data', 'users.json');

// In-memory cache so all serverless function calls within the same warm instance
// see the latest user list without re-reading from disk/env.
let _cache: User[] | null = null;

export function loadUsers(): User[] {
  if (_cache !== null) return _cache;

  // On Vercel, the file in data/users.json is the build-time git version (stale).
  // The env var DFE_USERS_JSON is updated at runtime and is the source of truth.
  const env = process.env.DFE_USERS_JSON;
  if (process.env.VERCEL && env) {
    _cache = JSON.parse(env);
    return _cache!;
  }

  // Local dev, or first Vercel deployment before DFE_USERS_JSON is populated
  if (fs.existsSync(FILE)) {
    _cache = JSON.parse(fs.readFileSync(FILE, 'utf-8'));
    return _cache!;
  }

  // Fallback: env var on non-Vercel hosts
  if (env) {
    _cache = JSON.parse(env);
    return _cache!;
  }

  return [];
}

export async function saveUsers(users: User[]) {
  // Always update in-memory cache so the current warm instance sees changes immediately
  _cache = users;

  // Try filesystem (local dev)
  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(users, null, 2));
    return; // success — local dev, we're done
  } catch {
    // Vercel: read-only filesystem, fall through to API update
  }

  // Persist to Vercel env var so cold starts read the latest user list
  // Must be awaited — if fire-and-forget, Vercel kills the process before it completes
  const token = process.env.VERCEL_TOKEN;
  const projectId = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';
  if (token) {
    try {
      await updateVercelEnvVar(token, projectId, users);
    } catch (err) {
      // Log but don't throw — in-memory cache is already updated so the
      // current warm instance is consistent.  Cold starts may be stale until
      // the next successful save, but a save failure must NOT crash the caller.
      console.error('[userData] Vercel env var update failed:', err);
    }
  }
}

async function updateVercelEnvVar(token: string, projectId: string, users: User[]) {
  // Find the env var record ID for DFE_USERS_JSON
  const listRes = await fetch(`https://api.vercel.com/v9/projects/${projectId}/env`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!listRes.ok) return;
  const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
  const envRecord = envs.find(e => e.key === 'DFE_USERS_JSON');
  if (!envRecord) return;

  await fetch(`https://api.vercel.com/v9/projects/${projectId}/env/${envRecord.id}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ value: JSON.stringify(users) }),
  });
}

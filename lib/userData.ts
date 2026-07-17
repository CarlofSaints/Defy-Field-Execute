import fs from 'fs';
import path from 'path';

export interface User {
  id: string;
  name: string;
  email: string;
  password: string;   // bcrypt hash
  isAdmin: boolean;
  forcePasswordChange: boolean;
  firstLoginAt: string | null;
  createdAt: string;
  // Per-user notification preferences. Replaces the old hardcoded
  // "email every admin on every run" behaviour. Both default to off.
  notifyRunSuccess?: boolean;  // email this user when a report generates OK
  notifyRunError?: boolean;    // email this user when a report run fails
}

// Users are read from the DFE_USERS_JSON env var on Vercel (the value is the
// source of truth), or the local file in dev. This is the original, known-good
// behaviour — deliberately NOT Blob/encryption, so the login path is as simple
// and robust as possible. JSON parsing is guarded so a malformed value can
// never throw and 500 the auth route.

const FILE       = path.join(process.cwd(), 'data', 'users.json');
const VERCEL_KEY = 'DFE_USERS_JSON';
const PROJECT_ID = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';

function safeParse(raw: string | undefined): User[] | null {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as User[] : null;
  } catch (err) {
    console.error('[userData] failed to parse user list:', err);
    return null;
  }
}

export async function loadUsers(): Promise<User[]> {
  const env = process.env[VERCEL_KEY];

  if (process.env.VERCEL) {
    const fromEnv = safeParse(env);
    if (fromEnv) return fromEnv;
  }

  if (fs.existsSync(FILE)) {
    const fromFile = safeParse(fs.readFileSync(FILE, 'utf-8'));
    if (fromFile) return fromFile;
  }

  const fromEnv = safeParse(env);
  if (fromEnv) return fromEnv;

  return [];
}

export async function saveUsers(users: User[]): Promise<void> {
  // Never overwrite the stored list with an empty one — that would lock everyone
  // out. Treat an empty save as a no-op. Also never throw: a save failure must
  // not break login (the auth route awaits this).
  if (users.length === 0) {
    console.error('[userData] refusing to save an empty user list');
    return;
  }

  // On Vercel the DFE_USERS_JSON env var is the source of truth that loadUsers
  // reads. The deployment filesystem is per-container and ephemeral: a
  // writeFileSync here can *succeed* against a throwaway disk that loadUsers
  // never reads, in which case the old code returned early and the env var was
  // never updated — so newly created users and password changes silently
  // vanished. On Vercel we must go straight to the env-var API.
  if (process.env.VERCEL) {
    await saveUsersToEnv(users);
    return;
  }

  // Local dev: write the file.
  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(users, null, 2));
  } catch (err) {
    console.error('[userData] local file save failed:', err);
  }
}

// Persist the user list to the DFE_USERS_JSON project env var via the Vercel
// API. NOTE: env-var changes only take effect on the *next* deployment — the
// running functions keep the value baked in at deploy time. A create/reset done
// through the app therefore needs a redeploy before it goes live.
async function saveUsersToEnv(users: User[]): Promise<void> {
  const token = process.env.VERCEL_TOKEN;
  if (!token) {
    console.error('[userData] VERCEL_TOKEN missing — cannot persist users');
    return;
  }
  try {
    const listRes = await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!listRes.ok) {
      console.error('[userData] env list failed:', listRes.status);
      return;
    }
    const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
    const envRecord = envs.find(e => e.key === VERCEL_KEY);
    if (!envRecord) {
      console.error('[userData] DFE_USERS_JSON env record not found');
      return;
    }
    const patchRes = await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env/${envRecord.id}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ value: JSON.stringify(users) }),
    });
    if (!patchRes.ok) {
      console.error('[userData] env patch failed:', patchRes.status);
    }
  } catch (err) {
    console.error('[userData] env-var save failed:', err);
  }
}

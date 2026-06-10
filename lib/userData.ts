import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { blobEnabled } from './blobStore';

export interface User {
  id: string;
  name: string;
  email: string;
  password: string;   // bcrypt hash
  isAdmin: boolean;
  forcePasswordChange: boolean;
  firstLoginAt: string | null;
  createdAt: string;
}

// Users are credential data, so they are NOT stored in the public Blob as
// plaintext. When DFE_USERS_ENC_KEY (a 32-byte / 64-hex-char secret) is set,
// the user list is AES-256-GCM encrypted before being written to the Blob and
// decrypted on read — unreadable even though the Blob is public.
//
// Until that key is set, behaviour is unchanged from before: the legacy env var
// (DFE_USERS_JSON) / local file is used. Existing users are read from that
// legacy source as the seed, so enabling the key never locks anyone out and no
// passwords change.

const FILE       = path.join(process.cwd(), 'data', 'users.json');
const BLOB_KEY   = 'config/users.json.enc';
const VERCEL_KEY = 'DFE_USERS_JSON';
const PROJECT_ID = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';
const useBlob    = blobEnabled;

function getEncKey(): Buffer | null {
  const hex = process.env.DFE_USERS_ENC_KEY;
  if (!hex) return null;
  try {
    const buf = Buffer.from(hex.trim(), 'hex');
    return buf.length === 32 ? buf : null;
  } catch {
    return null;
  }
}

// AES-256-GCM. Output = base64( iv[12] | authTag[16] | ciphertext ).
function encrypt(plaintext: string, key: Buffer): string {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
  const enc = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, enc]).toString('base64');
}

function decrypt(b64: string, key: Buffer): string {
  const raw = Buffer.from(b64, 'base64');
  const iv  = raw.subarray(0, 12);
  const tag = raw.subarray(12, 28);
  const enc = raw.subarray(28);
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(enc), decipher.final()]).toString('utf8');
}

function loadLegacy(): User[] {
  const env = process.env[VERCEL_KEY];
  if (process.env.VERCEL && env) return JSON.parse(env);
  if (fs.existsSync(FILE))       return JSON.parse(fs.readFileSync(FILE, 'utf-8'));
  if (env)                       return JSON.parse(env);
  return [];
}

export async function loadUsers(): Promise<User[]> {
  const key = getEncKey();
  if (useBlob && key) {
    const { list } = await import('@vercel/blob');
    const listing = await list({ prefix: BLOB_KEY });
    const match = listing.blobs.find(b => b.pathname === BLOB_KEY) ?? listing.blobs[0];
    if (match) {
      const res = await fetch(match.url, { cache: 'no-store' });
      if (res.ok) {
        try {
          return JSON.parse(decrypt(await res.text(), key)) as User[];
        } catch (err) {
          console.error('[userData] failed to decrypt users blob:', err);
          // Fall through to legacy seed rather than locking everyone out
        }
      }
    }
    // No encrypted blob yet — seed from the legacy source.
  }
  return loadLegacy();
}

export async function saveUsers(users: User[]): Promise<void> {
  const key = getEncKey();

  // Preferred path: encrypted Blob (only when both Blob and a key are available)
  if (useBlob && key) {
    const { put } = await import('@vercel/blob');
    await put(BLOB_KEY, encrypt(JSON.stringify(users), key), {
      access:          'public',
      contentType:     'text/plain',
      addRandomSuffix: false,
    });
    return;
  }

  // Fallback (local dev, or before the key is set): unchanged from before —
  // local file, then the Vercel env var via API.
  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(users, null, 2));
    return;
  } catch {
    // Vercel read-only FS — fall through to env-var API
  }

  const token = process.env.VERCEL_TOKEN;
  if (!token) return;
  try {
    const listRes = await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!listRes.ok) return;
    const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
    const envRecord = envs.find(e => e.key === VERCEL_KEY);
    if (!envRecord) return;
    await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env/${envRecord.id}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ value: JSON.stringify(users) }),
    });
  } catch (err) {
    console.error('[userData] env-var fallback save failed:', err);
  }
}

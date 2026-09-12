import { cookies } from 'next/headers';
import crypto from 'crypto';
import { loadUsers, User } from '@/lib/userData';

// Server-side session.
//
// Before this existed the app kept its session in localStorage only, and the
// one requireAdmin() helper (in app-settings) read a `dfe_session` COOKIE that
// nothing ever set — so it always returned false, and every other API route
// simply trusted the client. /api/users was reachable unauthenticated on a
// public domain: GET listed every user, POST created an admin.
//
// The cookie is httpOnly and HMAC-signed, so a caller cannot forge one by
// knowing somebody's email address. isAdmin is deliberately NOT trusted from
// the cookie — it is looked up from the stored user list on every request, so
// demoting someone takes effect without waiting for their cookie to expire.

export const SESSION_COOKIE = 'dfe_session';
const MAX_AGE_SECONDS = 60 * 60 * 24 * 30; // 30 days

interface SessionPayload {
  id: string;
  email: string;
  iat: number;
}

function secret(): string | null {
  return process.env.DFE_SESSION_SECRET || null;
}

function sign(data: string, key: string): string {
  return crypto.createHmac('sha256', key).update(data).digest('base64url');
}

export function createSessionCookie(user: Pick<User, 'id' | 'email'>): string | null {
  const key = secret();
  if (!key) {
    console.error('[session] DFE_SESSION_SECRET missing — cannot issue a session');
    return null;
  }
  const payload: SessionPayload = { id: user.id, email: user.email, iat: Date.now() };
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  return `${body}.${sign(body, key)}`;
}

export const sessionCookieOptions = {
  httpOnly: true,
  secure: process.env.NODE_ENV === 'production',
  sameSite: 'lax' as const,
  path: '/',
  maxAge: MAX_AGE_SECONDS,
};

function verify(raw: string | undefined): SessionPayload | null {
  const key = secret();
  if (!key || !raw) return null;
  const dot = raw.lastIndexOf('.');
  if (dot < 1) return null;
  const body = raw.slice(0, dot);
  const mac = raw.slice(dot + 1);
  const expected = sign(body, key);
  // Constant-time compare; timingSafeEqual throws on a length mismatch.
  if (mac.length !== expected.length) return null;
  if (!crypto.timingSafeEqual(Buffer.from(mac), Buffer.from(expected))) return null;
  try {
    const payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf-8')) as SessionPayload;
    if (!payload?.email) return null;
    if (Date.now() - payload.iat > MAX_AGE_SECONDS * 1000) return null;
    return payload;
  } catch {
    return null;
  }
}

// The signed-in user, re-read from the stored list (never trusted from the
// cookie). Returns null for a missing, forged or expired cookie.
export async function getSessionUser(): Promise<User | null> {
  const store = await cookies();
  const payload = verify(store.get(SESSION_COOKIE)?.value);
  if (!payload) return null;
  const users = await loadUsers();
  return users.find(u => u.id === payload.id || u.email.toLowerCase() === payload.email.toLowerCase()) || null;
}

export async function isAdmin(): Promise<boolean> {
  const user = await getSessionUser();
  return user?.isAdmin === true;
}

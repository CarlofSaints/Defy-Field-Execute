import { NextRequest } from 'next/server';
import { loadUsers } from '@/lib/userData';
import { createSessionCookie, sessionCookieOptions, SESSION_COOKIE } from '@/lib/session';
import { hubMessage, verifyHubToken } from '@/lib/hubSso';

export const dynamic = 'force-dynamic';

const MODULE = 'defy-field-execute';

/**
 * Sign-in from the Atomic Marketing Hub (lib/hubSso.ts). Existing users only, matched by email.
 * Sets the same signed cookie as /api/auth, plus the localStorage copy the pages render from
 * (written by a tiny page, because a redirect alone can't reach localStorage).
 */
export async function GET(req: NextRequest) {
  const secret = process.env.AM_SSO_SECRET;
  if (!secret) return hubMessage('Hub sign-in is not set up', 'AM_SSO_SECRET is missing on this deployment.', 503);

  const hub = verifyHubToken(req.nextUrl.searchParams.get('token'), secret);
  if (!hub) return hubMessage('This sign-in link has expired', 'Go back to the Hub and open Defy Field Execute again.', 401);
  if (!hub.modules.includes(MODULE)) return hubMessage('No access', 'Your Hub account does not include Defy Field Execute. Ask a Hub admin to switch it on.', 403);

  const email = hub.email.toLowerCase().trim();
  const user = (await loadUsers()).find(u => u.email.toLowerCase().trim() === email);
  if (!user) {
    return hubMessage('No Defy Field Execute account', `There is no Defy Field Execute user for ${hub.email}. Ask your admin to add you.`, 403);
  }

  const token = createSessionCookie(user);
  if (!token) return hubMessage('Sign-in is not configured', 'DFE_SESSION_SECRET is missing on this deployment.', 503);

  // Same shape app/login/page.tsx stores. JSON in a <script>: escape '<' so no value can close the tag.
  const local = JSON.stringify({ id: user.id, name: user.name, email: user.email, isAdmin: user.isAdmin }).replace(/</g, '\\u003c');
  const next = user.forcePasswordChange ? '/change-password' : '/';
  const res = new Response(
    `<!doctype html><meta charset="utf-8"><title>Signing in…</title><body style="font-family:Arial,sans-serif">Signing in…` +
    `<script>localStorage.setItem('dfe_session', JSON.stringify(${local}));location.replace(${JSON.stringify(next)});</script></body>`,
    { headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' } },
  );
  const o = sessionCookieOptions;
  res.headers.append('Set-Cookie',
    `${SESSION_COOKIE}=${token}; Path=${o.path ?? '/'}; Max-Age=${o.maxAge}; HttpOnly; SameSite=Lax${o.secure ? '; Secure' : ''}`);
  return res;
}

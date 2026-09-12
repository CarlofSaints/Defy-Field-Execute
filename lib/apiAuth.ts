import { NextResponse } from 'next/server';
import { getSessionUser } from '@/lib/session';
import { User } from '@/lib/userData';

// Route guards. Each returns either a response to send straight back, or the
// signed-in user. Every route that reads or writes the user list must use one —
// the client-side useAuth() check is cosmetic and does not protect the API.

type Guard = { deny: NextResponse; user?: undefined } | { deny?: undefined; user: User };

const UNAUTHENTICATED = () =>
  NextResponse.json({ error: 'Not signed in' }, { status: 401 });
const FORBIDDEN = () =>
  NextResponse.json({ error: 'Forbidden' }, { status: 403 });

export async function requireAdmin(): Promise<Guard> {
  const user = await getSessionUser();
  if (!user) return { deny: UNAUTHENTICATED() };
  if (!user.isAdmin) return { deny: FORBIDDEN() };
  return { user };
}

// Some actions are legitimately done by a non-admin on their own record — the
// forced password change at /change-password is the one that matters here.
// Gated ANY-of: an admin, OR the owner of that exact record.
export async function requireSelfOrAdmin(targetId: string): Promise<Guard> {
  const user = await getSessionUser();
  if (!user) return { deny: UNAUTHENTICATED() };
  if (user.isAdmin || user.id === targetId) return { user };
  return { deny: FORBIDDEN() };
}

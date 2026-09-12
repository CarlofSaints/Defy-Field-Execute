import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import { loadUsers } from '@/lib/userData';
import { createSessionCookie, sessionCookieOptions, SESSION_COOKIE } from '@/lib/session';

export async function POST(req: NextRequest) {
  const { email, password } = await req.json();
  if (!email || !password) {
    return NextResponse.json({ error: 'Missing credentials' }, { status: 400 });
  }

  const users = await loadUsers();
  const user = users.find(u => u.email.toLowerCase() === email.toLowerCase());
  if (!user) {
    return NextResponse.json({ error: 'Invalid credentials' }, { status: 401 });
  }

  const valid = await bcrypt.compare(password, user.password);
  if (!valid) {
    return NextResponse.json({ error: 'Invalid credentials' }, { status: 401 });
  }

  // DO NOT write the user list back from this route.
  // On Vercel loadUsers() reads DFE_USERS_JSON out of process.env, which is
  // BAKED at deploy time. Saving the whole list here writes back that frozen
  // copy, which silently deletes any user created since the last deploy - and
  // because the baked copy never changes, firstLoginAt always reads null, so
  // this fired on EVERY login, not just the first. Verified 11 Sep 2026.
  // firstLoginAt cannot be tracked correctly until the user store moves to a
  // runtime store (Blob) that the running deployment reads live.

  const res = NextResponse.json({
    id: user.id,
    name: user.name,
    email: user.email,
    isAdmin: user.isAdmin,
    forcePasswordChange: user.forcePasswordChange,
  });

  // The API routes need a session they can actually trust. The client keeps its
  // own copy in localStorage for rendering; this httpOnly signed cookie is what
  // the server checks. See lib/session.ts.
  const token = createSessionCookie(user);
  if (token) res.cookies.set(SESSION_COOKIE, token, sessionCookieOptions);

  return res;
}

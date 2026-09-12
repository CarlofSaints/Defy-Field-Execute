import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import { randomUUID } from 'crypto';
import { loadUsers, saveUsers, User } from '@/lib/userData';
import { requireAdmin } from '@/lib/apiAuth';
import { getDeployState, triggerRedeploy } from '@/lib/deployState';

// GET — list all users. Admin only, enforced HERE: this route sits on a public
// domain, so a client-side check protects nothing.
export async function GET() {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  const users = (await loadUsers()).map(({ password: _p, ...u }) => u);
  return NextResponse.json(users);
}

// POST — create user. Admin only, and refused while an earlier user write is
// still waiting to go live (see lib/deployState.ts) because this write would be
// built on a stale read and would delete that earlier user.
export async function POST(req: NextRequest) {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  const state = await getDeployState();
  if (state.pending || state.unknown) {
    return NextResponse.json({ error: state.message, deployState: state }, { status: 409 });
  }

  const { name, email, password, isAdmin, forcePasswordChange, notifyRunSuccess, notifyRunError } = await req.json();
  if (!name || !email || !password) {
    return NextResponse.json({ error: 'Missing fields' }, { status: 400 });
  }

  const users = await loadUsers();
  if (users.find(u => u.email.toLowerCase() === email.toLowerCase())) {
    return NextResponse.json({ error: 'Email already exists' }, { status: 409 });
  }

  const hashed = await bcrypt.hash(password, 10);
  const user: User = {
    id: randomUUID(),
    name,
    email,
    password: hashed,
    isAdmin: !!isAdmin,
    forcePasswordChange: forcePasswordChange !== false,
    firstLoginAt: null,
    createdAt: new Date().toISOString(),
    notifyRunSuccess: !!notifyRunSuccess,
    notifyRunError: !!notifyRunError,
  };
  users.push(user);

  // Only claim the user was created if the store actually took the write. This
  // used to return 201 unconditionally, so the UI said "created" and the
  // welcome email went out even when nothing had been saved.
  const saved = await saveUsers(users);
  if (!saved) {
    return NextResponse.json(
      { error: 'The user could not be saved. Nothing was created and no email was sent.' },
      { status: 500 },
    );
  }

  // Saved, but not live until a new build bakes it in — so start that build now
  // rather than leaving the admin waiting on someone else to deploy.
  const deploy = await triggerRedeploy();

  const { password: _p, ...safe } = user;
  return NextResponse.json(
    { ...safe, deploy: { triggered: deploy.ok, error: deploy.error ?? null } },
    { status: 201 },
  );
}

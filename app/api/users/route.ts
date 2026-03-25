import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import { randomUUID } from 'crypto';
import { loadUsers, saveUsers, User } from '@/lib/userData';

// GET — list all users (admin only — caller must verify on client)
export async function GET() {
  const users = loadUsers().map(({ password: _p, ...u }) => u);
  return NextResponse.json(users);
}

// POST — create user
export async function POST(req: NextRequest) {
  const { name, email, password, isAdmin, forcePasswordChange } = await req.json();
  if (!name || !email || !password) {
    return NextResponse.json({ error: 'Missing fields' }, { status: 400 });
  }

  const users = loadUsers();
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
  };
  users.push(user);
  await saveUsers(users);

  const { password: _p, ...safe } = user;
  return NextResponse.json(safe, { status: 201 });
}

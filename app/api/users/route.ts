import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import fs from 'fs';
import path from 'path';
import { randomUUID } from 'crypto';

interface User {
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
const load = (): User[] => JSON.parse(fs.readFileSync(FILE, 'utf-8'));
const save = (u: User[]) => fs.writeFileSync(FILE, JSON.stringify(u, null, 2));

// GET — list all users (admin only — caller must verify on client)
export async function GET() {
  const users = load().map(({ password: _p, ...u }) => u);
  return NextResponse.json(users);
}

// POST — create user
export async function POST(req: NextRequest) {
  const { name, email, password, isAdmin } = await req.json();
  if (!name || !email || !password) {
    return NextResponse.json({ error: 'Missing fields' }, { status: 400 });
  }

  const users = load();
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
    forcePasswordChange: true,
    firstLoginAt: null,
    createdAt: new Date().toISOString(),
  };
  users.push(user);
  save(users);

  const { password: _p, ...safe } = user;
  return NextResponse.json(safe, { status: 201 });
}

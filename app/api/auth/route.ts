import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import fs from 'fs';
import path from 'path';

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

function loadUsers(): User[] {
  const file = path.join(process.cwd(), 'data', 'users.json');
  return JSON.parse(fs.readFileSync(file, 'utf-8'));
}

function saveUsers(users: User[]) {
  const file = path.join(process.cwd(), 'data', 'users.json');
  fs.writeFileSync(file, JSON.stringify(users, null, 2));
}

export async function POST(req: NextRequest) {
  const { email, password } = await req.json();
  if (!email || !password) {
    return NextResponse.json({ error: 'Missing credentials' }, { status: 400 });
  }

  const users = loadUsers();
  const user = users.find(u => u.email.toLowerCase() === email.toLowerCase());
  if (!user) {
    return NextResponse.json({ error: 'Invalid credentials' }, { status: 401 });
  }

  const valid = await bcrypt.compare(password, user.password);
  if (!valid) {
    return NextResponse.json({ error: 'Invalid credentials' }, { status: 401 });
  }

  if (!user.firstLoginAt) {
    user.firstLoginAt = new Date().toISOString();
    saveUsers(users);
  }

  return NextResponse.json({
    id: user.id,
    name: user.name,
    email: user.email,
    isAdmin: user.isAdmin,
    forcePasswordChange: user.forcePasswordChange,
  });
}

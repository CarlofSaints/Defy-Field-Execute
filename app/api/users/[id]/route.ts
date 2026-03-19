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

const FILE = path.join(process.cwd(), 'data', 'users.json');
const load = (): User[] => JSON.parse(fs.readFileSync(FILE, 'utf-8'));
const save = (u: User[]) => fs.writeFileSync(FILE, JSON.stringify(u, null, 2));

export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const body = await req.json();
  const users = load();
  const idx = users.findIndex(u => u.id === id);
  if (idx === -1) return NextResponse.json({ error: 'Not found' }, { status: 404 });

  if (body.name !== undefined) users[idx].name = body.name;
  if (body.email !== undefined) users[idx].email = body.email;
  if (body.isAdmin !== undefined) users[idx].isAdmin = body.isAdmin;
  if (body.password) {
    users[idx].password = await bcrypt.hash(body.password, 10);
    users[idx].forcePasswordChange = true;
  }
  save(users);

  const { password: _p, ...safe } = users[idx];
  return NextResponse.json(safe);
}

export async function DELETE(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const users = load();
  const filtered = users.filter(u => u.id !== id);
  if (filtered.length === users.length) {
    return NextResponse.json({ error: 'Not found' }, { status: 404 });
  }
  save(filtered);
  return NextResponse.json({ ok: true });
}

import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import { loadUsers, saveUsers } from '@/lib/userData';

export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const body = await req.json();
  const users = loadUsers();
  const idx = users.findIndex(u => u.id === id);
  if (idx === -1) return NextResponse.json({ error: 'Not found' }, { status: 404 });

  if (body.name !== undefined) users[idx].name = body.name;
  if (body.email !== undefined) users[idx].email = body.email;
  if (body.isAdmin !== undefined) users[idx].isAdmin = body.isAdmin;
  if (body.password) {
    users[idx].password = await bcrypt.hash(body.password, 10);
    users[idx].forcePasswordChange = true;
  }
  saveUsers(users);

  const { password: _p, ...safe } = users[idx];
  return NextResponse.json(safe);
}

export async function DELETE(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const users = loadUsers();
  const filtered = users.filter(u => u.id !== id);
  if (filtered.length === users.length) {
    return NextResponse.json({ error: 'Not found' }, { status: 404 });
  }
  saveUsers(filtered);
  return NextResponse.json({ ok: true });
}

import { NextRequest, NextResponse } from 'next/server';
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

export async function POST(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const { plainPassword, type } = await req.json();

  const users = load();
  const user = users.find(u => u.id === id);
  if (!user) return NextResponse.json({ error: 'Not found' }, { status: 404 });

  const { sendWelcomeEmail, sendPasswordResetEmail } = await import('@/lib/email');

  if (type === 'reset' && plainPassword) {
    await sendPasswordResetEmail(user.email, user.name, plainPassword);
  } else if (plainPassword) {
    await sendWelcomeEmail(user.email, user.name, plainPassword);
  }

  return NextResponse.json({ ok: true });
}

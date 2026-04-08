import { NextRequest, NextResponse } from 'next/server';
import { loadAppSettings, saveAppSettings, AppSettings } from '@/lib/appSettings';
import { cookies } from 'next/headers';
import { loadUsers } from '@/lib/userData';

async function requireAdmin(): Promise<boolean> {
  const cookieStore = await cookies();
  const raw = cookieStore.get('dfe_session')?.value;
  if (!raw) return false;
  try {
    const session = JSON.parse(raw);
    const users = loadUsers();
    const user = users.find(u => u.email === session.email);
    return user?.isAdmin === true;
  } catch {
    return false;
  }
}

export async function GET() {
  const settings = loadAppSettings();
  return NextResponse.json(settings);
}

export async function PATCH(req: NextRequest) {
  if (!(await requireAdmin())) {
    return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
  }

  let body: Partial<AppSettings>;
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: 'Invalid JSON' }, { status: 400 });
  }

  const current  = loadAppSettings();
  const updated: AppSettings = {
    ...current,
    ...(typeof body.picturesFolderPath === 'string'
      ? { picturesFolderPath: body.picturesFolderPath.trim() }
      : {}),
    ...(typeof body.activationPicturesFolderPath === 'string'
      ? { activationPicturesFolderPath: body.activationPicturesFolderPath.trim() }
      : {}),
  };

  await saveAppSettings(updated);
  return NextResponse.json(updated);
}

import { NextRequest, NextResponse } from 'next/server';
import { loadAppSettings, saveAppSettings, AppSettings } from '@/lib/appSettings';
import { requireAdmin } from '@/lib/apiAuth';

// The helper that used to live here read a `dfe_session` cookie that nothing
// ever set — the session was in localStorage — so it always returned false and
// this PATCH was a permanent 403. It now uses the shared guard in lib/apiAuth.

export async function GET() {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  const settings = await loadAppSettings();
  return NextResponse.json(settings);
}

export async function PATCH(req: NextRequest) {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  let body: Partial<AppSettings>;
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: 'Invalid JSON' }, { status: 400 });
  }

  const current  = await loadAppSettings();
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

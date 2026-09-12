import { NextResponse } from 'next/server';
import { loadRunLog } from '@/lib/runLogData';
import { requireAdmin } from '@/lib/apiAuth';

export async function GET() {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  return NextResponse.json(await loadRunLog());
}

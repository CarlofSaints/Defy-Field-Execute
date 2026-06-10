import { NextRequest, NextResponse } from 'next/server';
import { loadUsers } from '@/lib/userData';

// TEMPORARY diagnostic — no secrets returned (no password hashes, no full list).
// Delete this route once login is sorted.
export const dynamic = 'force-dynamic';

export async function GET(req: NextRequest) {
  const email = req.nextUrl.searchParams.get('email')?.toLowerCase().trim() || '';

  const raw = process.env.DFE_USERS_JSON;
  let envParses = false;
  let envCount = 0;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) { envParses = true; envCount = parsed.length; }
    } catch { /* envParses stays false */ }
  }

  let count = 0;
  let loadError: string | null = null;
  let found = false;
  let isAdmin: boolean | null = null;
  let hasHash = false;
  try {
    const users = await loadUsers();
    count = users.length;
    if (email) {
      const u = users.find(x => (x.email || '').toLowerCase() === email);
      found = !!u;
      isAdmin = u ? u.isAdmin : null;
      hasHash = !!(u && u.password);
    }
  } catch (e) {
    loadError = e instanceof Error ? e.message : String(e);
  }

  return NextResponse.json({
    commit: process.env.VERCEL_GIT_COMMIT_SHA ?? 'unknown',
    vercel: !!process.env.VERCEL,
    envVar: { present: !!raw, length: raw?.length ?? 0, parses: envParses, count: envCount },
    loadUsers: { count, error: loadError },
    lookup: email ? { email, found, isAdmin, hasHash } : 'no email param',
  });
}

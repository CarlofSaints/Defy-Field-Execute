import { NextResponse } from 'next/server';
import { requireAdmin } from '@/lib/apiAuth';
import { getDeployState, triggerRedeploy } from '@/lib/deployState';

// Tells the admin UI whether a saved user change is live yet, so the page can
// lock the forms rather than let a second write delete the first.
export async function GET() {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  return NextResponse.json(await getDeployState(), {
    headers: { 'Cache-Control': 'no-store' },
  });
}

// Manual fallback for the "Deploy now" button, used when the automatic trigger
// after a save did not fire.
export async function POST() {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  const result = await triggerRedeploy();
  if (!result.ok) {
    return NextResponse.json({ error: result.error || 'Could not start the deploy' }, { status: 502 });
  }
  return NextResponse.json({ ok: true });
}

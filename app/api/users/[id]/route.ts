import { NextRequest, NextResponse } from 'next/server';
import bcrypt from 'bcryptjs';
import { loadUsers, saveUsers } from '@/lib/userData';
import { requireAdmin, requireSelfOrAdmin } from '@/lib/apiAuth';
import { getDeployState, triggerRedeploy } from '@/lib/deployState';

export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;

    // ANY-of: an admin editing anyone, or a user editing their own record —
    // /change-password is a non-admin self-edit and must keep working.
    const guard = await requireSelfOrAdmin(id);
    if (guard.deny) return guard.deny;
    const actingAsAdmin = guard.user.isAdmin;

    // Any write rebuilds the whole list from a read, so it carries the same
    // lost-update hazard as a create and is refused for the same reason.
    const state = await getDeployState();
    if (state.pending || state.unknown) {
      return NextResponse.json({ error: state.message, deployState: state }, { status: 409 });
    }

    const body = await req.json();
    const users = await loadUsers();
    const idx = users.findIndex(u => u.id === id);
    if (idx === -1) return NextResponse.json({ error: 'Not found' }, { status: 404 });

    if (actingAsAdmin) {
      if (body.name !== undefined) users[idx].name = body.name;
      if (body.email !== undefined) users[idx].email = body.email;
      if (body.isAdmin !== undefined) users[idx].isAdmin = body.isAdmin;
      if (body.notifyRunSuccess !== undefined) users[idx].notifyRunSuccess = body.notifyRunSuccess;
      if (body.notifyRunError !== undefined) users[idx].notifyRunError = body.notifyRunError;
    } else if (body.isAdmin !== undefined && body.isAdmin !== users[idx].isAdmin) {
      // A non-admin editing their own record may change their password and
      // nothing else. Without this they could PATCH themselves to isAdmin.
      return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
    }

    if (body.password) {
      users[idx].password = await bcrypt.hash(body.password, 10);
      // If the caller explicitly passes forcePasswordChange: false (e.g. the
      // user completing their own forced reset), honour it.  Otherwise default
      // to true so admin-set passwords always prompt a change on next login.
      users[idx].forcePasswordChange = body.forcePasswordChange !== false;
    }

    const saved = await saveUsers(users);
    if (!saved) {
      return NextResponse.json(
        { error: 'The change could not be saved. Nothing was updated and no email was sent.' },
        { status: 500 },
      );
    }

    const deploy = await triggerRedeploy();

    const { password: _p, ...safe } = users[idx];
    return NextResponse.json({ ...safe, deploy: { triggered: deploy.ok, error: deploy.error ?? null } });
  } catch (err) {
    console.error('[PATCH /api/users/[id]]', err);
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 });
  }
}

export async function DELETE(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  const state = await getDeployState();
  if (state.pending || state.unknown) {
    return NextResponse.json({ error: state.message, deployState: state }, { status: 409 });
  }

  const { id } = await params;
  const users = await loadUsers();
  const filtered = users.filter(u => u.id !== id);
  if (filtered.length === users.length) {
    return NextResponse.json({ error: 'Not found' }, { status: 404 });
  }

  const saved = await saveUsers(filtered);
  if (!saved) {
    return NextResponse.json({ error: 'The user could not be deleted.' }, { status: 500 });
  }

  const deploy = await triggerRedeploy();
  return NextResponse.json({ ok: true, deploy: { triggered: deploy.ok, error: deploy.error ?? null } });
}

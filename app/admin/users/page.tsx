'use client';

import { useAuth } from '@/lib/useAuth';
import Header from '@/components/Header';
import { useEffect, useState } from 'react';

interface User {
  id: string;
  name: string;
  email: string;
  isAdmin: boolean;
  forcePasswordChange: boolean;
  firstLoginAt: string | null;
  createdAt: string;
  notifyRunSuccess?: boolean;
  notifyRunError?: boolean;
}

type Toast = { message: string; type: 'success' | 'error' };

// Mirrors lib/deployState.ts
interface DeployState {
  pending: boolean;
  building: boolean;
  unknown: boolean;
  savedAt: string | null;
  builtAt: string | null;
  message: string;
}

// Shown whenever a saved user change has not reached the running app yet.
// While it is up every user form on this page is locked, because a second write
// would be built on the stale list and would wipe the change already saved.
function PendingBanner({ state, onDeploy, deploying }: {
  state: DeployState; onDeploy: () => void; deploying: boolean;
}) {
  const amber = state.building || state.pending;
  return (
    <div className={`rounded-xl px-5 py-4 border flex flex-col sm:flex-row sm:items-center gap-3
      ${amber ? 'bg-amber-50 border-amber-200' : 'bg-red-50 border-red-200'}`}>
      <div className="flex items-start gap-3 flex-1">
        {state.building && (
          <span className="mt-0.5 h-4 w-4 rounded-full border-2 border-amber-500 border-t-transparent animate-spin shrink-0" />
        )}
        <div>
          <p className="text-sm font-semibold text-gray-900">
            {state.building ? 'Going live now' : state.unknown ? 'Cannot confirm the user list is live' : 'Waiting to go live'}
          </p>
          <p className="text-xs text-gray-700 mt-0.5">{state.message}</p>
          <p className="text-xs text-gray-500 mt-1">
            Adding, editing and deleting users is paused until this finishes. This page unlocks on its own.
          </p>
        </div>
      </div>
      {!state.building && (
        <button onClick={onDeploy} disabled={deploying}
          className="shrink-0 bg-gray-900 hover:bg-black disabled:opacity-50 text-white text-xs font-bold px-4 py-2 rounded-lg transition-colors">
          {deploying ? 'Starting…' : 'Deploy now'}
        </button>
      )}
    </div>
  );
}

function Toast({ toast, onClose }: { toast: Toast; onClose: () => void }) {
  useEffect(() => {
    const t = setTimeout(onClose, 4000);
    return () => clearTimeout(t);
  }, [onClose]);
  return (
    <div className={`fixed top-20 right-4 z-50 px-4 py-3 rounded-lg shadow-lg text-sm font-medium text-white
      ${toast.type === 'success' ? 'bg-green-600' : 'bg-red-600'}`}>
      {toast.message}
    </div>
  );
}

export default function AdminUsersPage() {
  const { session, loading, logout } = useAuth(true);
  const [users, setUsers] = useState<User[]>([]);
  const [toast, setToast] = useState<Toast | null>(null);
  const [deploy, setDeploy] = useState<DeployState | null>(null);
  const [deploying, setDeploying] = useState(false);
  const [sessionExpired, setSessionExpired] = useState(false);

  // Add user form
  const [addName, setAddName] = useState('');
  const [addEmail, setAddEmail] = useState('');
  const [addPw, setAddPw] = useState('');
  const [addAdmin, setAddAdmin] = useState(false);
  const [addForcePwChange, setAddForcePwChange] = useState(true);
  const [showAddPw, setShowAddPw] = useState(false);
  const [sendWelcome, setSendWelcome] = useState(true);
  const [addNotifySuccess, setAddNotifySuccess] = useState(false);
  const [addNotifyError, setAddNotifyError] = useState(false);
  const [addLoading, setAddLoading] = useState(false);

  // Edit modal
  const [editUser, setEditUser] = useState<User | null>(null);
  const [editName, setEditName] = useState('');
  const [editEmail, setEditEmail] = useState('');
  const [editAdmin, setEditAdmin] = useState(false);
  const [editNotifySuccess, setEditNotifySuccess] = useState(false);
  const [editNotifyError, setEditNotifyError] = useState(false);
  const [editPw, setEditPw] = useState('');
  const [showEditPw, setShowEditPw] = useState(false);
  const [sendReset, setSendReset] = useState(false);
  const [editLoading, setEditLoading] = useState(false);

  const notify = (message: string, type: 'success' | 'error' = 'success') =>
    setToast({ message, type });

  async function loadUsers() {
    const res = await fetch('/api/users');
    // The API is now gated server-side. Without this an expired session would
    // come back 401 and render as "no users", which reads as data loss.
    if (res.status === 401 || res.status === 403) { setSessionExpired(true); return; }
    if (res.ok) setUsers(await res.json());
  }

  async function refreshDeploy(): Promise<DeployState | null> {
    const res = await fetch('/api/deploy-state');
    if (res.status === 401 || res.status === 403) { setSessionExpired(true); return null; }
    if (!res.ok) return null;
    const state: DeployState = await res.json();
    setDeploy(state);
    return state;
  }

  useEffect(() => { if (session) { loadUsers(); refreshDeploy(); } }, [session]);

  // While a change is waiting to go live, poll until the new build is serving,
  // then unlock and pull the now-live list.
  const waiting = !!deploy && (deploy.pending || deploy.unknown);
  useEffect(() => {
    if (!session || !waiting) return;
    const t = setInterval(async () => {
      const state = await refreshDeploy();
      if (state && !state.pending && !state.unknown) {
        loadUsers();
        notify('Changes are live. You can add or edit users again.');
      }
    }, 10000);
    return () => clearInterval(t);
  }, [session, waiting]);

  async function handleDeployNow() {
    setDeploying(true);
    try {
      const res = await fetch('/api/deploy-state', { method: 'POST' });
      if (res.ok) { notify('Deploy started. This page unlocks when it finishes.'); refreshDeploy(); }
      else {
        const data = await res.json().catch(() => ({}));
        notify(data.error || 'Could not start the deploy', 'error');
      }
    } finally {
      setDeploying(false);
    }
  }

  async function handleAdd(e: React.FormEvent) {
    e.preventDefault();
    setAddLoading(true);
    try {
      const res = await fetch('/api/users', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: addName, email: addEmail, password: addPw, isAdmin: addAdmin, forcePasswordChange: addForcePwChange, notifyRunSuccess: addNotifySuccess, notifyRunError: addNotifyError }),
      });
      const data = await res.json();
      if (!res.ok) {
        // 409 with a deployState means an earlier change is still going live.
        if (data.deployState) setDeploy(data.deployState);
        notify(data.error || 'Failed to create user', 'error');
        return;
      }

      // Saved for real (the API only returns 201 once the store took the
      // write), so the welcome email is now safe to send.
      if (sendWelcome) {
        await fetch(`/api/users/${data.id}/notify`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ plainPassword: addPw, type: 'welcome', name: addName, email: addEmail }),
        });
      }
      notify(`User ${addName} created${sendWelcome ? ' — welcome email sent' : ''}. Going live now.`);
      setAddName(''); setAddEmail(''); setAddPw(''); setAddAdmin(false); setAddForcePwChange(true); setSendWelcome(true); setAddNotifySuccess(false); setAddNotifyError(false);
      // Lock the forms straight away rather than waiting for the next poll.
      setDeploy({
        pending: true,
        building: data.deploy?.triggered === true,
        unknown: false,
        savedAt: new Date().toISOString(),
        builtAt: null,
        message: data.deploy?.triggered
          ? `${addName} is saved and the deploy that makes them live has started.`
          : `${addName} is saved but the deploy did not start automatically${data.deploy?.error ? ` (${data.deploy.error})` : ''}. Use Deploy now.`,
      });
      refreshDeploy();
      loadUsers();
    } finally {
      setAddLoading(false);
    }
  }

  function openEdit(user: User) {
    setEditUser(user);
    setEditName(user.name);
    setEditEmail(user.email);
    setEditAdmin(user.isAdmin);
    setEditNotifySuccess(!!user.notifyRunSuccess);
    setEditNotifyError(!!user.notifyRunError);
    setEditPw('');
    setShowEditPw(false);
    setSendReset(false);
  }

  async function handleEdit(e: React.FormEvent) {
    e.preventDefault();
    if (!editUser) return;
    setEditLoading(true);
    try {
      const body: Record<string, unknown> = { name: editName, email: editEmail, isAdmin: editAdmin, notifyRunSuccess: editNotifySuccess, notifyRunError: editNotifyError };
      if (editPw) body.password = editPw;

      const res = await fetch(`/api/users/${editUser.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        if (err.deployState) setDeploy(err.deployState);
        notify(err.error || 'Failed to update user', 'error');
        return;
      }
      const data = await res.json().catch(() => ({}));

      if (editPw && sendReset) {
        await fetch(`/api/users/${editUser.id}/notify`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ plainPassword: editPw, type: 'reset', name: editName, email: editEmail }),
        });
      }
      notify(`User updated${editPw && sendReset ? ' — reset email sent' : ''}. Going live now.`);
      setEditUser(null);
      setDeploy({
        pending: true,
        building: data.deploy?.triggered === true,
        unknown: false,
        savedAt: new Date().toISOString(),
        builtAt: null,
        message: data.deploy?.triggered
          ? 'The change is saved and the deploy that makes it live has started.'
          : `The change is saved but the deploy did not start automatically${data.deploy?.error ? ` (${data.deploy.error})` : ''}. Use Deploy now.`,
      });
      refreshDeploy();
      loadUsers();
    } finally {
      setEditLoading(false);
    }
  }

  async function handleDelete(user: User) {
    if (!confirm(`Delete user ${user.name}? This cannot be undone.`)) return;
    const res = await fetch(`/api/users/${user.id}`, { method: 'DELETE' });
    const data = await res.json().catch(() => ({}));
    if (res.ok) {
      notify('User deleted. Going live now.');
      setDeploy({
        pending: true,
        building: data.deploy?.triggered === true,
        unknown: false,
        savedAt: new Date().toISOString(),
        builtAt: null,
        message: data.deploy?.triggered
          ? `${user.name} is deleted and the deploy that makes it live has started.`
          : 'The delete is saved but the deploy did not start automatically. Use Deploy now.',
      });
      refreshDeploy();
      loadUsers();
    } else {
      if (data.deployState) setDeploy(data.deployState);
      notify(data.error || 'Failed to delete user', 'error');
    }
  }

  // Every user write rebuilds the whole list from a read, so while a previous
  // write is still going live all of them must stay shut.
  const locked = !!deploy && (deploy.pending || deploy.unknown);

  if (loading || !session) return null;

  if (sessionExpired) {
    return (
      <div className="min-h-screen flex items-center justify-center px-4 bg-gray-50">
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-8 max-w-sm text-center">
          <h1 className="text-base font-bold text-gray-900">Your session has expired</h1>
          <p className="text-sm text-gray-600 mt-2">
            Sign in again to manage users. Nothing has been lost.
          </p>
          <button onClick={logout}
            className="mt-5 bg-[#E31837] hover:bg-[#c01430] text-white text-sm font-bold px-5 py-2 rounded-lg transition-colors">
            Sign in again
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen" style={{ backgroundImage: "url('/defy logo grey.png')", backgroundSize: '160px', backgroundRepeat: 'repeat', backgroundColor: 'rgb(252,252,252)', backgroundBlendMode: 'luminosity' }}>
      <Header session={session} onLogout={logout} />
      {toast && <Toast toast={toast} onClose={() => setToast(null)} />}

      <main className="max-w-screen-lg mx-auto px-4 py-8 flex flex-col gap-8">
        <div className="bg-white rounded-xl shadow-sm border-l-4 border-[#E31837] px-6 py-4 flex items-center gap-3">
          <h1 className="text-xl font-bold text-gray-900">User Management</h1>
        </div>

        {locked && deploy && (
          <PendingBanner state={deploy} onDeploy={handleDeployNow} deploying={deploying} />
        )}

        {/* Add User */}
        <section className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
          <h2 className="text-sm font-bold text-gray-700 uppercase tracking-wide mb-4">Add New User</h2>
          <form onSubmit={handleAdd} className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="flex flex-col gap-1">
              <label className="text-xs text-gray-500 font-medium">Full Name</label>
              <input value={addName} onChange={e => setAddName(e.target.value)} required
                className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]" />
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-xs text-gray-500 font-medium">Email</label>
              <input type="email" value={addEmail} onChange={e => setAddEmail(e.target.value)} required
                className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]" />
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-xs text-gray-500 font-medium">Password</label>
              <div className="relative">
                <input type={showAddPw ? 'text' : 'password'} value={addPw} onChange={e => setAddPw(e.target.value)} required
                  className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm pr-10 focus:outline-none focus:ring-2 focus:ring-[#E31837]" />
                <button type="button" onClick={() => setShowAddPw(v => !v)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 text-xs" tabIndex={-1}>
                  {showAddPw ? 'Hide' : 'Show'}
                </button>
              </div>
            </div>
            <div className="flex flex-col gap-3 justify-end">
              <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                <input type="checkbox" checked={addAdmin} onChange={e => setAddAdmin(e.target.checked)}
                  className="accent-[#E31837]" />
                Admin user
              </label>
              <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                <input type="checkbox" checked={addForcePwChange} onChange={e => setAddForcePwChange(e.target.checked)}
                  className="accent-[#E31837]" />
                Force password change on first login
              </label>
              <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                <input type="checkbox" checked={sendWelcome} onChange={e => setSendWelcome(e.target.checked)}
                  className="accent-[#E31837]" />
                Send welcome email
              </label>
              <div className="pt-1 border-t border-gray-100 mt-1">
                <p className="text-[11px] uppercase tracking-wide text-gray-400 font-semibold mb-1.5 mt-2">Report-run notifications</p>
                <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                  <input type="checkbox" checked={addNotifySuccess} onChange={e => setAddNotifySuccess(e.target.checked)}
                    className="accent-[#E31837]" />
                  Email on successful runs
                </label>
                <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer mt-2">
                  <input type="checkbox" checked={addNotifyError} onChange={e => setAddNotifyError(e.target.checked)}
                    className="accent-[#E31837]" />
                  Email on failed runs
                </label>
              </div>
            </div>
            <div className="sm:col-span-2">
              <button type="submit" disabled={addLoading || locked}
                className="bg-[#E31837] hover:bg-[#c01430] disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-bold px-6 py-2 rounded-lg transition-colors">
                {addLoading ? 'Creating…' : locked ? 'Waiting for the last change to go live…' : 'Create User'}
              </button>
            </div>
          </form>
        </section>

        {/* Users Table */}
        <section className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <h2 className="text-sm font-bold text-gray-700 uppercase tracking-wide p-6 pb-0">All Users</h2>
          <div className="overflow-x-auto mt-4">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-gray-100 bg-gray-50">
                  <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Name</th>
                  <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Email</th>
                  <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Role</th>
                  <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Notifications</th>
                  <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">First Login</th>
                  <th className="px-6 py-3" />
                </tr>
              </thead>
              <tbody>
                {users.map(u => (
                  <tr key={u.id} className="border-b border-gray-50 hover:bg-gray-50 transition-colors">
                    <td className="px-6 py-3 font-medium text-gray-900">{u.name}</td>
                    <td className="px-6 py-3 text-gray-600">{u.email}</td>
                    <td className="px-6 py-3">
                      <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${u.isAdmin ? 'bg-red-100 text-[#E31837]' : 'bg-gray-100 text-gray-600'}`}>
                        {u.isAdmin ? 'Admin' : 'User'}
                      </span>
                    </td>
                    <td className="px-6 py-3">
                      <div className="flex flex-wrap gap-1">
                        {u.notifyRunSuccess && (
                          <span className="text-[11px] font-medium px-2 py-0.5 rounded-full bg-green-100 text-green-700">Success</span>
                        )}
                        {u.notifyRunError && (
                          <span className="text-[11px] font-medium px-2 py-0.5 rounded-full bg-red-100 text-red-700">Errors</span>
                        )}
                        {!u.notifyRunSuccess && !u.notifyRunError && (
                          <span className="text-[11px] text-gray-400">—</span>
                        )}
                      </div>
                    </td>
                    <td className="px-6 py-3 text-gray-500 text-xs">
                      {u.firstLoginAt ? new Date(u.firstLoginAt).toLocaleDateString() : 'Never'}
                    </td>
                    <td className="px-6 py-3">
                      <div className="flex gap-2 justify-end">
                        <button onClick={() => openEdit(u)} disabled={locked}
                          className="text-xs text-blue-600 hover:text-blue-800 font-medium disabled:text-gray-300 disabled:cursor-not-allowed">Edit</button>
                        <button onClick={() => handleDelete(u)} disabled={locked}
                          className="text-xs text-red-500 hover:text-red-700 font-medium disabled:text-gray-300 disabled:cursor-not-allowed">Delete</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </main>

      {/* Edit Modal */}
      {editUser && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center px-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-md p-6">
            <h2 className="text-base font-bold text-gray-900 mb-5">Edit User</h2>
            <form onSubmit={handleEdit} className="flex flex-col gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-xs text-gray-500 font-medium">Full Name</label>
                <input value={editName} onChange={e => setEditName(e.target.value)} required
                  className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-xs text-gray-500 font-medium">Email</label>
                <input type="email" value={editEmail} onChange={e => setEditEmail(e.target.value)} required
                  className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]" />
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-xs text-gray-500 font-medium">New Password <span className="text-gray-400 font-normal">(leave blank to keep current)</span></label>
                <div className="relative">
                  <input type={showEditPw ? 'text' : 'password'} value={editPw} onChange={e => setEditPw(e.target.value)}
                    className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm pr-10 focus:outline-none focus:ring-2 focus:ring-[#E31837]"
                    placeholder="New password…" />
                  <button type="button" onClick={() => setShowEditPw(v => !v)}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 text-xs" tabIndex={-1}>
                    {showEditPw ? 'Hide' : 'Show'}
                  </button>
                </div>
              </div>
              <div className="flex flex-col gap-2">
                <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                  <input type="checkbox" checked={editAdmin} onChange={e => setEditAdmin(e.target.checked)} className="accent-[#E31837]" />
                  Admin user
                </label>
                <div className="pt-2 mt-1 border-t border-gray-100">
                  <p className="text-[11px] uppercase tracking-wide text-gray-400 font-semibold mb-1.5 mt-1">Report-run notifications</p>
                  <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                    <input type="checkbox" checked={editNotifySuccess} onChange={e => setEditNotifySuccess(e.target.checked)} className="accent-[#E31837]" />
                    Email on successful runs
                  </label>
                  <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer mt-2">
                    <input type="checkbox" checked={editNotifyError} onChange={e => setEditNotifyError(e.target.checked)} className="accent-[#E31837]" />
                    Email on failed runs
                  </label>
                </div>
                {editPw && (
                  <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                    <input type="checkbox" checked={sendReset} onChange={e => setSendReset(e.target.checked)} className="accent-[#E31837]" />
                    Send password reset email
                  </label>
                )}
              </div>
              <div className="flex gap-3 pt-2">
                <button type="submit" disabled={editLoading || locked}
                  className="bg-[#E31837] hover:bg-[#c01430] disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-bold px-5 py-2 rounded-lg transition-colors">
                  {editLoading ? 'Saving…' : locked ? 'Waiting to go live…' : 'Save Changes'}
                </button>
                <button type="button" onClick={() => setEditUser(null)}
                  className="text-sm text-gray-600 hover:text-gray-900 px-4 py-2 rounded-lg border border-gray-200 hover:bg-gray-50 transition-colors">
                  Cancel
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

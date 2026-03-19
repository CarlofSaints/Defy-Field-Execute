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
}

type Toast = { message: string; type: 'success' | 'error' };

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

  // Add user form
  const [addName, setAddName] = useState('');
  const [addEmail, setAddEmail] = useState('');
  const [addPw, setAddPw] = useState('');
  const [addAdmin, setAddAdmin] = useState(false);
  const [showAddPw, setShowAddPw] = useState(false);
  const [sendWelcome, setSendWelcome] = useState(true);
  const [addLoading, setAddLoading] = useState(false);

  // Edit modal
  const [editUser, setEditUser] = useState<User | null>(null);
  const [editName, setEditName] = useState('');
  const [editEmail, setEditEmail] = useState('');
  const [editAdmin, setEditAdmin] = useState(false);
  const [editPw, setEditPw] = useState('');
  const [showEditPw, setShowEditPw] = useState(false);
  const [sendReset, setSendReset] = useState(false);
  const [editLoading, setEditLoading] = useState(false);

  const notify = (message: string, type: 'success' | 'error' = 'success') =>
    setToast({ message, type });

  async function loadUsers() {
    const res = await fetch('/api/users');
    if (res.ok) setUsers(await res.json());
  }

  useEffect(() => { if (session) loadUsers(); }, [session]);

  async function handleAdd(e: React.FormEvent) {
    e.preventDefault();
    setAddLoading(true);
    try {
      const res = await fetch('/api/users', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: addName, email: addEmail, password: addPw, isAdmin: addAdmin }),
      });
      const data = await res.json();
      if (!res.ok) { notify(data.error || 'Failed to create user', 'error'); return; }

      if (sendWelcome) {
        await fetch(`/api/users/${data.id}/notify`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ plainPassword: addPw, type: 'welcome' }),
        });
      }
      notify(`User ${addName} created${sendWelcome ? ' — welcome email sent' : ''}`);
      setAddName(''); setAddEmail(''); setAddPw(''); setAddAdmin(false); setSendWelcome(true);
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
    setEditPw('');
    setShowEditPw(false);
    setSendReset(false);
  }

  async function handleEdit(e: React.FormEvent) {
    e.preventDefault();
    if (!editUser) return;
    setEditLoading(true);
    try {
      const body: Record<string, unknown> = { name: editName, email: editEmail, isAdmin: editAdmin };
      if (editPw) body.password = editPw;

      const res = await fetch(`/api/users/${editUser.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      if (!res.ok) { notify('Failed to update user', 'error'); return; }

      if (editPw && sendReset) {
        await fetch(`/api/users/${editUser.id}/notify`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ plainPassword: editPw, type: 'reset' }),
        });
      }
      notify(`User updated${editPw && sendReset ? ' — reset email sent' : ''}`);
      setEditUser(null);
      loadUsers();
    } finally {
      setEditLoading(false);
    }
  }

  async function handleDelete(user: User) {
    if (!confirm(`Delete user ${user.name}? This cannot be undone.`)) return;
    const res = await fetch(`/api/users/${user.id}`, { method: 'DELETE' });
    if (res.ok) { notify('User deleted'); loadUsers(); }
    else notify('Failed to delete user', 'error');
  }

  if (loading || !session) return null;

  return (
    <div className="min-h-screen bg-gray-50">
      <Header session={session} onLogout={logout} />
      {toast && <Toast toast={toast} onClose={() => setToast(null)} />}

      <main className="max-w-screen-lg mx-auto px-4 py-8 flex flex-col gap-8">
        <h1 className="text-xl font-bold text-gray-900">User Management</h1>

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
                <input type="checkbox" checked={sendWelcome} onChange={e => setSendWelcome(e.target.checked)}
                  className="accent-[#E31837]" />
                Send welcome email
              </label>
            </div>
            <div className="sm:col-span-2">
              <button type="submit" disabled={addLoading}
                className="bg-[#E31837] hover:bg-[#c01430] disabled:opacity-50 text-white text-sm font-bold px-6 py-2 rounded-lg transition-colors">
                {addLoading ? 'Creating…' : 'Create User'}
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
                    <td className="px-6 py-3 text-gray-500 text-xs">
                      {u.firstLoginAt ? new Date(u.firstLoginAt).toLocaleDateString() : 'Never'}
                    </td>
                    <td className="px-6 py-3">
                      <div className="flex gap-2 justify-end">
                        <button onClick={() => openEdit(u)}
                          className="text-xs text-blue-600 hover:text-blue-800 font-medium">Edit</button>
                        {u.id !== session.id && (
                          <button onClick={() => handleDelete(u)}
                            className="text-xs text-red-500 hover:text-red-700 font-medium">Delete</button>
                        )}
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
                {editPw && (
                  <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                    <input type="checkbox" checked={sendReset} onChange={e => setSendReset(e.target.checked)} className="accent-[#E31837]" />
                    Send password reset email
                  </label>
                )}
              </div>
              <div className="flex gap-3 pt-2">
                <button type="submit" disabled={editLoading}
                  className="bg-[#E31837] hover:bg-[#c01430] disabled:opacity-50 text-white text-sm font-bold px-5 py-2 rounded-lg transition-colors">
                  {editLoading ? 'Saving…' : 'Save Changes'}
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

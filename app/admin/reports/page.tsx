'use client';

import { useAuth } from '@/lib/useAuth';
import Header from '@/components/Header';
import { useEffect, useState, useCallback } from 'react';

interface ReportDef {
  id: string;
  name: string;
  outputTypes: string[];
  brands: string[];
}

const ALL_BRANDS      = ['Defy', 'Beko', 'Grundig'];
const ALL_OUTPUT_TYPES = ['Excel', 'PPT'];

type Toast = { message: string; type: 'success' | 'error' };

function ToastBanner({ toast, onClose }: { toast: Toast; onClose: () => void }) {
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

// Multi-select pill checkbox group
function PillGroup({
  options, selected, onChange,
}: { options: string[]; selected: string[]; onChange: (v: string[]) => void }) {
  const toggle = (v: string) =>
    onChange(selected.includes(v) ? selected.filter(x => x !== v) : [...selected, v]);
  return (
    <div className="flex flex-wrap gap-2">
      {options.map(opt => (
        <button
          key={opt}
          type="button"
          onClick={() => toggle(opt)}
          className={`px-3 py-1 rounded-full text-sm font-medium border transition-colors ${
            selected.includes(opt)
              ? 'bg-[#E31837] text-white border-[#E31837]'
              : 'bg-white text-gray-700 border-gray-300 hover:border-[#E31837] hover:text-[#E31837]'
          }`}
        >
          {opt}
        </button>
      ))}
    </div>
  );
}

export default function AdminReportsPage() {
  const { session, loading, logout } = useAuth(true);
  const [reports, setReports]   = useState<ReportDef[]>([]);
  const [toast, setToast]       = useState<Toast | null>(null);

  // Add form
  const [addName,        setAddName]        = useState('');
  const [addOutputTypes, setAddOutputTypes] = useState<string[]>(['Excel']);
  const [addBrands,      setAddBrands]      = useState<string[]>(['Defy', 'Beko', 'Grundig']);
  const [addLoading,     setAddLoading]     = useState(false);

  // Edit modal
  const [editReport,      setEditReport]      = useState<ReportDef | null>(null);
  const [editName,        setEditName]        = useState('');
  const [editOutputTypes, setEditOutputTypes] = useState<string[]>([]);
  const [editBrands,      setEditBrands]      = useState<string[]>([]);
  const [editLoading,     setEditLoading]     = useState(false);

  const notify = (message: string, type: 'success' | 'error' = 'success') =>
    setToast({ message, type });

  const loadReports = useCallback(async () => {
    const res = await fetch('/api/reports');
    if (res.ok) setReports(await res.json());
  }, []);

  useEffect(() => { loadReports(); }, [loadReports]);

  async function handleAdd(e: React.FormEvent) {
    e.preventDefault();
    if (!addName.trim() || !addOutputTypes.length || !addBrands.length) {
      notify('Fill in all fields', 'error');
      return;
    }
    setAddLoading(true);
    const res = await fetch('/api/reports', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: addName.trim(), outputTypes: addOutputTypes, brands: addBrands }),
    });
    setAddLoading(false);
    if (res.ok) {
      notify('Report added');
      setAddName('');
      setAddOutputTypes(['Excel']);
      setAddBrands(['Defy', 'Beko', 'Grundig']);
      loadReports();
    } else {
      const { error } = await res.json();
      notify(error || 'Failed to add', 'error');
    }
  }

  function openEdit(r: ReportDef) {
    setEditReport(r);
    setEditName(r.name);
    setEditOutputTypes([...r.outputTypes]);
    setEditBrands([...r.brands]);
  }

  async function handleEdit(e: React.FormEvent) {
    e.preventDefault();
    if (!editReport || !editName.trim() || !editOutputTypes.length || !editBrands.length) {
      notify('Fill in all fields', 'error');
      return;
    }
    setEditLoading(true);
    const res = await fetch(`/api/reports/${editReport.id}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: editName.trim(), outputTypes: editOutputTypes, brands: editBrands }),
    });
    setEditLoading(false);
    if (res.ok) {
      notify('Report updated');
      setEditReport(null);
      loadReports();
    } else {
      const { error } = await res.json();
      notify(error || 'Failed to update', 'error');
    }
  }

  async function handleDelete(id: string, name: string) {
    if (!confirm(`Delete "${name}"?`)) return;
    const res = await fetch(`/api/reports/${id}`, { method: 'DELETE' });
    if (res.ok) {
      notify('Report deleted');
      loadReports();
    } else {
      notify('Failed to delete', 'error');
    }
  }

  if (loading || !session) return null;

  return (
    <div
      className="min-h-screen"
      style={{
        backgroundImage: 'url(/defy logo grey.png)',
        backgroundSize: '160px',
        backgroundBlendMode: 'luminosity',
        backgroundColor: 'rgb(252,252,252)',
      }}
    >
      <Header session={session} onLogout={logout} />
      {toast && <ToastBanner toast={toast} onClose={() => setToast(null)} />}

      <main className="max-w-screen-lg mx-auto px-4 py-8 space-y-6">

        {/* Page header */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="border-l-4 border-[#E31837] px-6 py-4">
            <h1 className="text-xl font-bold text-gray-900">Reports Control Centre</h1>
            <p className="text-sm text-gray-500 mt-0.5">Add, edit or remove report types available to users.</p>
          </div>
        </div>

        {/* Add report form */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="border-l-4 border-[#E31837] px-6 py-4">
            <h2 className="font-semibold text-gray-800">Add New Report</h2>
          </div>
          <form onSubmit={handleAdd} className="px-6 pb-6 space-y-4">
            <div>
              <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-1">
                Report Name
              </label>
              <input
                value={addName}
                onChange={e => setAddName(e.target.value.toUpperCase())}
                placeholder="e.g. MAKRO STOCK COUNT"
                className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]/30 focus:border-[#E31837]"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-2">
                Output Types
              </label>
              <PillGroup options={ALL_OUTPUT_TYPES} selected={addOutputTypes} onChange={setAddOutputTypes} />
            </div>

            <div>
              <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-2">
                Available For Brands
              </label>
              <PillGroup options={ALL_BRANDS} selected={addBrands} onChange={setAddBrands} />
            </div>

            <button
              type="submit"
              disabled={addLoading}
              className="bg-[#E31837] hover:bg-[#c01430] text-white px-5 py-2 rounded-lg text-sm font-semibold disabled:opacity-50"
            >
              {addLoading ? 'Adding…' : 'Add Report'}
            </button>
          </form>
        </div>

        {/* Reports list */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="border-l-4 border-[#E31837] px-6 py-4">
            <h2 className="font-semibold text-gray-800">
              Report List <span className="text-gray-400 font-normal ml-2">{reports.length} reports</span>
            </h2>
          </div>

          {reports.length === 0 ? (
            <p className="px-6 pb-6 text-sm text-gray-400">No reports configured yet.</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-gray-50 border-b border-gray-100">
                  <tr>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Name</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Brands</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Output</th>
                    <th className="px-4 py-3" />
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {reports.map(r => (
                    <tr key={r.id} className="hover:bg-gray-50 transition-colors">
                      <td className="px-6 py-3 font-medium text-gray-900">{r.name}</td>
                      <td className="px-4 py-3">
                        <div className="flex flex-wrap gap-1">
                          {r.brands.map(b => (
                            <span key={b} className="px-2 py-0.5 bg-blue-50 text-blue-700 rounded text-xs font-medium">{b}</span>
                          ))}
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <div className="flex flex-wrap gap-1">
                          {r.outputTypes.map(t => (
                            <span key={t} className={`px-2 py-0.5 rounded text-xs font-medium ${
                              t === 'Excel' ? 'bg-green-50 text-green-700' : 'bg-orange-50 text-orange-700'
                            }`}>{t}</span>
                          ))}
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <div className="flex gap-2 justify-end">
                          <button
                            onClick={() => openEdit(r)}
                            className="text-xs text-blue-600 hover:text-blue-800 font-medium"
                          >
                            Edit
                          </button>
                          <button
                            onClick={() => handleDelete(r.id, r.name)}
                            className="text-xs text-red-600 hover:text-red-800 font-medium"
                          >
                            Delete
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </main>

      {/* Edit modal */}
      {editReport && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md">
            <div className="border-l-4 border-[#E31837] px-6 py-4">
              <h2 className="font-bold text-gray-900">Edit Report</h2>
            </div>
            <form onSubmit={handleEdit} className="px-6 pb-6 space-y-4">
              <div>
                <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-1">
                  Report Name
                </label>
                <input
                  value={editName}
                  onChange={e => setEditName(e.target.value.toUpperCase())}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]/30 focus:border-[#E31837]"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-2">
                  Output Types
                </label>
                <PillGroup options={ALL_OUTPUT_TYPES} selected={editOutputTypes} onChange={setEditOutputTypes} />
              </div>

              <div>
                <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-2">
                  Available For Brands
                </label>
                <PillGroup options={ALL_BRANDS} selected={editBrands} onChange={setEditBrands} />
              </div>

              <div className="flex gap-3 pt-2">
                <button
                  type="button"
                  onClick={() => setEditReport(null)}
                  className="flex-1 border border-gray-200 text-gray-700 hover:bg-gray-50 px-4 py-2 rounded-lg text-sm font-medium"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={editLoading}
                  className="flex-1 bg-[#E31837] hover:bg-[#c01430] text-white px-4 py-2 rounded-lg text-sm font-semibold disabled:opacity-50"
                >
                  {editLoading ? 'Saving…' : 'Save Changes'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

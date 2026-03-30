'use client';

import { useAuth } from '@/lib/useAuth';
import Header from '@/components/Header';
import { useEffect, useState, useCallback, useRef } from 'react';

interface ReportDef {
  id: string;
  name: string;
  dataFormat: string;
  channel?: string;
  outputTypes: string[];
  brands: string[];
}

const DATA_FORMATS = [
  { value: 'stock-count',        label: 'Stock Count' },
  { value: 'red-flag',           label: 'Red Flag' },
  { value: 'stand-report',       label: 'Stand Report' },
  { value: 'service-call',       label: 'Service Call' },
  { value: 'training-feedback',  label: 'Training Feedback' },
  { value: 'activation-report',  label: 'Activation Report' },
];

interface RunLogEntry {
  id:               string;
  timestamp:        string;
  userName:         string;
  userEmail:        string;
  reportId:         string;
  reportName:       string;
  brand:            string;
  retailer:         string;
  filename:         string;
  status:           'success' | 'error';
  errorMessage?:    string;
  spPath?:          string;
  emailSent?:       boolean;
  emailRecipients?: string[];
}

const ALL_BRANDS      = ['Defy', 'Beko'];
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

  // Store Province Mapping
  const [storeMapCount,      setStoreMapCount]      = useState<number | null>(null);
  const [storeMapUploading,  setStoreMapUploading]  = useState(false);

  // App Settings — SP image path
  const [picturesPath,        setPicturesPath]        = useState('');
  const [picturesPathSaving,  setPicturesPathSaving]  = useState(false);
  const [picturesPathLoaded,  setPicturesPathLoaded]  = useState(false);

  // Run Log
  const [runLog,        setRunLog]        = useState<RunLogEntry[]>([]);
  const [runLogLoading, setRunLogLoading] = useState(false);
  const runLogRef = useRef<HTMLDivElement>(null);

  // Add form
  const [addName,        setAddName]        = useState('');
  const [addDataFormat,  setAddDataFormat]  = useState('');
  const [addChannel,     setAddChannel]     = useState('');
  const [addOutputTypes, setAddOutputTypes] = useState<string[]>(['Excel']);
  const [addBrands,      setAddBrands]      = useState<string[]>(['Defy', 'Beko']);
  const [addLoading,     setAddLoading]     = useState(false);

  // Edit modal
  const [editReport,      setEditReport]      = useState<ReportDef | null>(null);
  const [editName,        setEditName]        = useState('');
  const [editDataFormat,  setEditDataFormat]  = useState('');
  const [editChannel,     setEditChannel]     = useState('');
  const [editOutputTypes, setEditOutputTypes] = useState<string[]>([]);
  const [editBrands,      setEditBrands]      = useState<string[]>([]);
  const [editLoading,     setEditLoading]     = useState(false);

  const notify = (message: string, type: 'success' | 'error' = 'success') =>
    setToast({ message, type });

  const loadReports = useCallback(async () => {
    const res = await fetch('/api/reports');
    if (res.ok) setReports(await res.json());
  }, []);

  const loadStoreMapCount = useCallback(async () => {
    const res = await fetch('/api/store-map');
    if (res.ok) {
      const { count } = await res.json();
      setStoreMapCount(count);
    }
  }, []);

  const loadAppSettings = useCallback(async () => {
    const res = await fetch('/api/app-settings');
    if (res.ok) {
      const data = await res.json();
      setPicturesPath(data.picturesFolderPath ?? '');
      setPicturesPathLoaded(true);
    }
  }, []);

  const loadRunLogData = useCallback(async () => {
    setRunLogLoading(true);
    const res = await fetch('/api/run-log');
    if (res.ok) setRunLog(await res.json());
    setRunLogLoading(false);
  }, []);

  useEffect(() => {
    loadReports();
    loadStoreMapCount();
    loadRunLogData();
    loadAppSettings();
  }, [loadReports, loadStoreMapCount, loadRunLogData, loadAppSettings]);

  async function handleSavePicturesPath() {
    setPicturesPathSaving(true);
    const res = await fetch('/api/app-settings', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ picturesFolderPath: picturesPath }),
    });
    setPicturesPathSaving(false);
    if (res.ok) {
      notify('Image folder path saved');
    } else {
      const { error } = await res.json().catch(() => ({ error: 'Save failed' }));
      notify(error || 'Save failed', 'error');
    }
  }

  async function handleStoreMapUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setStoreMapUploading(true);
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetch('/api/store-map', { method: 'POST', body: fd });
    setStoreMapUploading(false);
    // Reset input so same file can be re-uploaded
    e.target.value = '';
    if (res.ok) {
      const { count } = await res.json();
      setStoreMapCount(count);
      notify(`Store map updated — ${count} stores loaded`);
    } else {
      const { error } = await res.json().catch(() => ({ error: 'Upload failed' }));
      notify(error || 'Upload failed', 'error');
    }
  }

  async function handleAdd(e: React.FormEvent) {
    e.preventDefault();
    if (!addName.trim() || !addDataFormat || !addOutputTypes.length || !addBrands.length) {
      notify('Fill in all fields (name, data format, output type, brands)', 'error');
      return;
    }
    setAddLoading(true);
    const res = await fetch('/api/reports', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: addName.trim(),
        dataFormat: addDataFormat,
        channel: addChannel.trim() || undefined,
        outputTypes: addOutputTypes,
        brands: addBrands,
      }),
    });
    setAddLoading(false);
    if (res.ok) {
      notify('Report added');
      setAddName('');
      setAddDataFormat('');
      setAddChannel('');
      setAddOutputTypes(['Excel']);
      setAddBrands(['Defy', 'Beko']);
      loadReports();
    } else {
      const { error } = await res.json();
      notify(error || 'Failed to add', 'error');
    }
  }

  function openEdit(r: ReportDef) {
    setEditReport(r);
    setEditName(r.name);
    setEditDataFormat(r.dataFormat || '');
    setEditChannel(r.channel || '');
    setEditOutputTypes([...r.outputTypes]);
    setEditBrands([...r.brands]);
  }

  async function handleEdit(e: React.FormEvent) {
    e.preventDefault();
    if (!editReport || !editName.trim() || !editDataFormat || !editOutputTypes.length || !editBrands.length) {
      notify('Fill in all fields (name, data format, output type, brands)', 'error');
      return;
    }
    setEditLoading(true);
    const res = await fetch(`/api/reports/${editReport.id}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: editName.trim(),
        dataFormat: editDataFormat,
        channel: editChannel.trim() || '',
        outputTypes: editOutputTypes,
        brands: editBrands,
      }),
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
            <h1 className="text-xl font-bold text-gray-900">Control Centre</h1>
            <p className="text-sm text-gray-500 mt-0.5">Manage report types, brands, and reference data for the platform.</p>
          </div>
        </div>

        {/* Reports section heading */}
        <div className="px-1">
          <h2 className="text-base font-bold text-gray-800 uppercase tracking-wide">Report Management</h2>
          <p className="text-xs text-gray-400 mt-0.5">Add, edit or remove report types available to users.</p>
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

            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-1">
                  Data Format <span className="text-red-500">*</span>
                </label>
                <select
                  value={addDataFormat}
                  onChange={e => setAddDataFormat(e.target.value)}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]/30 focus:border-[#E31837] bg-white"
                >
                  <option value="">Select format…</option>
                  {DATA_FORMATS.map(f => (
                    <option key={f.value} value={f.value}>{f.label}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-1">
                  Channel <span className="text-gray-400 font-normal normal-case">(optional)</span>
                </label>
                <input
                  value={addChannel}
                  onChange={e => setAddChannel(e.target.value.toUpperCase())}
                  placeholder="e.g. MAKRO, GAME"
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]/30 focus:border-[#E31837]"
                />
              </div>
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
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Data Format</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Channel</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Brands</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Output</th>
                    <th className="px-4 py-3" />
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {reports.map(r => {
                    const fmt = DATA_FORMATS.find(f => f.value === r.dataFormat);
                    return (
                      <tr key={r.id} className="hover:bg-gray-50 transition-colors">
                        <td className="px-6 py-3 font-medium text-gray-900">{r.name}</td>
                        <td className="px-4 py-3">
                          {fmt ? (
                            <span className="px-2 py-0.5 bg-purple-50 text-purple-700 rounded text-xs font-medium">{fmt.label}</span>
                          ) : (
                            <span className="text-gray-400 text-xs">—</span>
                          )}
                        </td>
                        <td className="px-4 py-3">
                          {r.channel ? (
                            <span className="px-2 py-0.5 bg-amber-50 text-amber-700 rounded text-xs font-medium">{r.channel}</span>
                          ) : (
                            <span className="text-gray-300 text-xs">—</span>
                          )}
                        </td>
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
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
        {/* Run Log section heading */}
        <div className="px-1 pt-2">
          <h2 className="text-base font-bold text-gray-800 uppercase tracking-wide">Run Log</h2>
          <p className="text-xs text-gray-400 mt-0.5">Every report generation — who ran it, when, and what was produced.</p>
        </div>

        {/* Run Log table */}
        <div ref={runLogRef} className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="border-l-4 border-[#E31837] px-6 py-4 flex items-center justify-between">
            <div>
              <h2 className="font-semibold text-gray-800">
                Report Runs
                <span className="text-gray-400 font-normal ml-2">{runLog.length} entries</span>
              </h2>
            </div>
            <button
              onClick={loadRunLogData}
              disabled={runLogLoading}
              className="text-xs text-[#E31837] hover:text-[#c01430] font-medium disabled:opacity-40"
            >
              {runLogLoading ? 'Refreshing…' : '↻ Refresh'}
            </button>
          </div>

          {runLog.length === 0 ? (
            <p className="px-6 pb-6 text-sm text-gray-400">
              {runLogLoading ? 'Loading…' : 'No reports have been run yet.'}
            </p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-gray-50 border-b border-gray-100">
                  <tr>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Date / Time</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">User</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Report</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Brand</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">File</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">SP</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Email</th>
                    <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {runLog.map(entry => {
                    const ts = new Date(entry.timestamp);
                    const dateStr = ts.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
                    const timeStr = ts.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', hour12: false });
                    return (
                      <tr key={entry.id} className="hover:bg-gray-50 transition-colors">
                        <td className="px-4 py-3 text-gray-700 whitespace-nowrap">
                          <span className="font-medium">{dateStr}</span>
                          <span className="text-gray-400 ml-1.5">{timeStr}</span>
                        </td>
                        <td className="px-4 py-3">
                          <div className="font-medium text-gray-900 leading-tight">{entry.userName}</div>
                          <div className="text-xs text-gray-400 leading-tight">{entry.userEmail}</div>
                        </td>
                        <td className="px-4 py-3">
                          <div className="font-medium text-gray-900 leading-tight">{entry.reportName}</div>
                          {entry.retailer && <div className="text-xs text-gray-400 leading-tight">{entry.retailer}</div>}
                        </td>
                        <td className="px-4 py-3">
                          <span className="px-2 py-0.5 bg-blue-50 text-blue-700 rounded text-xs font-medium">{entry.brand}</span>
                        </td>
                        <td className="px-4 py-3 text-xs text-gray-500 font-mono max-w-[180px] truncate" title={entry.filename}>
                          {entry.filename || '—'}
                        </td>
                        <td className="px-4 py-3 text-center">
                          {entry.spPath ? (
                            <a
                              href={entry.spPath}
                              target="_blank"
                              rel="noopener noreferrer"
                              title={entry.spPath}
                              className="inline-flex items-center justify-center w-6 h-6 rounded bg-blue-50 hover:bg-blue-100 text-blue-600 transition-colors"
                            >
                              <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                                <path strokeLinecap="round" strokeLinejoin="round" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1" />
                              </svg>
                            </a>
                          ) : (
                            <span className="text-gray-300 text-xs">—</span>
                          )}
                        </td>
                        <td className="px-4 py-3">
                          {entry.emailSent ? (
                            <span
                              className="px-2 py-0.5 bg-green-50 text-green-700 rounded text-xs font-semibold"
                              title={entry.emailRecipients?.join(', ')}
                            >
                              Sent
                            </span>
                          ) : (
                            <span className="text-gray-300 text-xs">—</span>
                          )}
                        </td>
                        <td className="px-4 py-3">
                          {entry.status === 'success' ? (
                            <span className="px-2 py-0.5 bg-green-50 text-green-700 rounded text-xs font-semibold">Success</span>
                          ) : (
                            <span className="px-2 py-0.5 bg-red-50 text-red-700 rounded text-xs font-semibold" title={entry.errorMessage}>
                              Error
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Store Maintenance section heading */}
        <div className="px-1">
          <h2 className="text-base font-bold text-gray-800 uppercase tracking-wide">Store Maintenance</h2>
          <p className="text-xs text-gray-400 mt-0.5">Manage reference data used to enrich reports.</p>
        </div>

        {/* Store Province Mapping */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="border-l-4 border-[#E31837] px-6 py-4">
            <h2 className="font-semibold text-gray-800">Store Province Mapping</h2>
            <p className="text-sm text-gray-500 mt-0.5">
              Upload a control file to map stores to provinces. Reports use this to populate the PROVINCE column.
            </p>
          </div>
          <div className="px-6 pb-6 space-y-3">
            <p className="text-xs text-gray-500">
              Expected columns: <span className="font-mono bg-gray-100 px-1 rounded">STORE NAME</span>
              {' · '}<span className="font-mono bg-gray-100 px-1 rounded">STORE CODE</span>
              {' · '}<span className="font-mono bg-gray-100 px-1 rounded">PROVINCE</span>
            </p>

            {storeMapCount !== null && (
              <p className={`text-sm font-medium ${storeMapCount > 0 ? 'text-green-700' : 'text-gray-400'}`}>
                {storeMapCount > 0
                  ? `✓ ${storeMapCount} store${storeMapCount === 1 ? '' : 's'} currently mapped`
                  : 'No store map loaded yet'}
              </p>
            )}

            <div>
              <label className={`inline-block cursor-pointer px-5 py-2 rounded-lg text-sm font-semibold text-white transition-colors ${
                storeMapUploading
                  ? 'bg-gray-400 cursor-not-allowed'
                  : 'bg-[#E31837] hover:bg-[#c01430]'
              }`}>
                {storeMapUploading ? 'Uploading…' : storeMapCount && storeMapCount > 0 ? 'Replace Control File' : 'Upload Control File'}
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={handleStoreMapUpload}
                  disabled={storeMapUploading}
                />
              </label>
            </div>
          </div>
        </div>

        {/* Red Flag Image Folder */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="border-l-4 border-[#E31837] px-6 py-4">
            <h2 className="font-semibold text-gray-800">Red Flag Image Folder</h2>
            <p className="text-sm text-gray-500 mt-0.5">
              SharePoint path to the folder where red-flag images are stored. Leave blank to use the default.
            </p>
          </div>
          <div className="px-6 pb-6 space-y-3">
            <p className="text-xs text-gray-500">
              Default: <span className="font-mono bg-gray-100 px-1 rounded">
                {process.env.NEXT_PUBLIC_DFE_SP_BASE_PATH
                  ? `${process.env.NEXT_PUBLIC_DFE_SP_BASE_PATH}/PERIGEE IMAGE DOWNLOADS`
                  : 'DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS/PERIGEE IMAGE DOWNLOADS'}
              </span>
            </p>

            <input
              type="text"
              value={picturesPathLoaded ? picturesPath : ''}
              onChange={e => setPicturesPath(e.target.value)}
              placeholder="e.g. DEFY/PERIGEE - FG/2. EXTERNAL SYNC/REPORTS/PERIGEE IMAGE DOWNLOADS"
              className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-[#E31837]/30 focus:border-[#E31837]"
              disabled={!picturesPathLoaded}
            />

            <button
              type="button"
              onClick={handleSavePicturesPath}
              disabled={picturesPathSaving || !picturesPathLoaded}
              className="bg-[#E31837] hover:bg-[#c01430] text-white px-5 py-2 rounded-lg text-sm font-semibold disabled:opacity-50"
            >
              {picturesPathSaving ? 'Saving…' : 'Save Path'}
            </button>
          </div>
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

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-1">
                    Data Format <span className="text-red-500">*</span>
                  </label>
                  <select
                    value={editDataFormat}
                    onChange={e => setEditDataFormat(e.target.value)}
                    className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]/30 focus:border-[#E31837] bg-white"
                  >
                    <option value="">Select format…</option>
                    {DATA_FORMATS.map(f => (
                      <option key={f.value} value={f.value}>{f.label}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-gray-600 uppercase tracking-wide mb-1">
                    Channel <span className="text-gray-400 font-normal normal-case">(optional)</span>
                  </label>
                  <input
                    value={editChannel}
                    onChange={e => setEditChannel(e.target.value.toUpperCase())}
                    placeholder="e.g. MAKRO, GAME"
                    className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#E31837]/30 focus:border-[#E31837]"
                  />
                </div>
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

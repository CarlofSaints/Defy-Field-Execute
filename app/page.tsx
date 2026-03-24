'use client';

import { useAuth } from '@/lib/useAuth';
import Header from '@/components/Header';
import { useEffect, useRef, useState, useCallback } from 'react';

interface ReportDef {
  id: string;
  name: string;
  outputTypes: string[];
  brands: string[];
}

const ALL_BRANDS = ['Defy', 'Beko', 'Grundig'];

// ─── Brand pill (multi-select) ───────────────────────────────────────────────
function BrandPill({ label, selected, onClick }: { label: string; selected: boolean; onClick: () => void }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`px-4 py-2 rounded-full text-sm font-semibold border-2 transition-all ${
        selected
          ? 'bg-[#E31837] text-white border-[#E31837] shadow-sm'
          : 'bg-white text-gray-600 border-gray-200 hover:border-[#E31837] hover:text-[#E31837]'
      }`}
    >
      {label}
    </button>
  );
}

// ─── Searchable single-select report dropdown ────────────────────────────────
function ReportDropdown({
  reports,
  value,
  onChange,
}: {
  reports: ReportDef[];
  value: ReportDef | null;
  onChange: (r: ReportDef | null) => void;
}) {
  const [open, setOpen]     = useState(false);
  const [query, setQuery]   = useState('');
  const inputRef            = useRef<HTMLInputElement>(null);
  const containerRef        = useRef<HTMLDivElement>(null);

  const filtered = reports.filter(r =>
    r.name.toLowerCase().includes(query.toLowerCase())
  );

  // Close on outside click
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  function select(r: ReportDef) {
    onChange(r);
    setQuery('');
    setOpen(false);
  }

  return (
    <div ref={containerRef} className="relative">
      <button
        type="button"
        onClick={() => { setOpen(o => !o); setTimeout(() => inputRef.current?.focus(), 50); }}
        className={`w-full flex items-center justify-between px-4 py-2.5 rounded-xl border-2 text-sm transition-all ${
          open
            ? 'border-[#E31837] ring-2 ring-[#E31837]/20'
            : 'border-gray-200 hover:border-gray-300'
        } bg-white`}
      >
        <span className={value ? 'text-gray-900 font-medium' : 'text-gray-400'}>
          {value ? value.name : 'Search for a report…'}
        </span>
        <svg
          className={`w-4 h-4 text-gray-400 transition-transform ${open ? 'rotate-180' : ''}`}
          fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}
        >
          <path strokeLinecap="round" strokeLinejoin="round" d="M19 9l-7 7-7-7" />
        </svg>
      </button>

      {open && (
        <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-gray-200 rounded-xl shadow-xl z-20 overflow-hidden">
          {/* Search input */}
          <div className="p-2 border-b border-gray-100">
            <input
              ref={inputRef}
              value={query}
              onChange={e => setQuery(e.target.value)}
              placeholder="Type to filter…"
              className="w-full px-3 py-1.5 text-sm border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#E31837]/30 focus:border-[#E31837]"
            />
          </div>
          {/* Options */}
          <div className="max-h-56 overflow-y-auto">
            {filtered.length === 0 ? (
              <p className="px-4 py-3 text-sm text-gray-400">No reports match.</p>
            ) : (
              filtered.map(r => (
                <button
                  key={r.id}
                  type="button"
                  onClick={() => select(r)}
                  className={`w-full text-left px-4 py-2.5 text-sm transition-colors hover:bg-red-50 hover:text-[#E31837] ${
                    value?.id === r.id ? 'bg-red-50 text-[#E31837] font-semibold' : 'text-gray-700'
                  }`}
                >
                  {r.name}
                  <span className="ml-2 text-xs text-gray-400">{r.outputTypes.join(' / ')}</span>
                </button>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Output type pill (single-select) ────────────────────────────────────────
function TypePill({ label, selected, onClick }: { label: string; selected: boolean; onClick: () => void }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`px-4 py-2 rounded-full text-sm font-semibold border-2 transition-all ${
        selected
          ? 'bg-[#E31837] text-white border-[#E31837] shadow-sm'
          : 'bg-white text-gray-600 border-gray-200 hover:border-[#E31837] hover:text-[#E31837]'
      }`}
    >
      {label}
    </button>
  );
}

// ─── File drop zone ──────────────────────────────────────────────────────────
function DropZone({
  file,
  onChange,
}: {
  file: File | null;
  onChange: (f: File | null) => void;
}) {
  const inputRef  = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);

  function handleFiles(files: FileList | null) {
    if (files && files[0]) onChange(files[0]);
  }

  return (
    <div
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); handleFiles(e.dataTransfer.files); }}
      onClick={() => inputRef.current?.click()}
      className={`cursor-pointer rounded-xl border-2 border-dashed p-8 text-center transition-all ${
        drag
          ? 'border-[#E31837] bg-red-50'
          : file
          ? 'border-green-400 bg-green-50'
          : 'border-gray-200 hover:border-gray-300 bg-white hover:bg-gray-50'
      }`}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        className="hidden"
        onChange={e => handleFiles(e.target.files)}
      />
      {file ? (
        <div className="space-y-2">
          <div className="flex items-center justify-center gap-2 text-green-700">
            <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            <span className="font-semibold text-sm">{file.name}</span>
          </div>
          <p className="text-xs text-green-600">{(file.size / 1024).toFixed(0)} KB — click to change</p>
        </div>
      ) : (
        <div className="space-y-2">
          <svg className="w-10 h-10 mx-auto text-gray-300" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
          </svg>
          <p className="text-sm font-medium text-gray-500">Drop your Perigee Excel export here</p>
          <p className="text-xs text-gray-400">or click to browse (.xlsx)</p>
        </div>
      )}
    </div>
  );
}

// ─── Step number badge ────────────────────────────────────────────────────────
function StepBadge({ n, done }: { n: number; done: boolean }) {
  return (
    <div className={`w-7 h-7 rounded-full flex items-center justify-center text-sm font-bold shrink-0 ${
      done ? 'bg-green-500 text-white' : 'bg-[#E31837] text-white'
    }`}>
      {done ? (
        <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={3}>
          <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
        </svg>
      ) : n}
    </div>
  );
}

// ─── Page ─────────────────────────────────────────────────────────────────────
export default function DashboardPage() {
  const { session, loading, logout } = useAuth();

  const [reports,     setReports]     = useState<ReportDef[]>([]);
  const [brands,      setBrands]      = useState<string[]>(['Defy']);
  const [report,      setReport]      = useState<ReportDef | null>(null);
  const [outputType,  setOutputType]  = useState<string>('Excel');
  const [file,        setFile]        = useState<File | null>(null);
  const [generating,  setGenerating]  = useState(false);
  const [error,       setError]       = useState<string | null>(null);
  const [success,     setSuccess]     = useState<string | null>(null);

  const loadReports = useCallback(async () => {
    const res = await fetch('/api/reports');
    if (res.ok) setReports(await res.json());
  }, []);

  useEffect(() => { loadReports(); }, [loadReports]);

  // When a report is selected, default outputType to its first option
  useEffect(() => {
    if (report?.outputTypes?.length) setOutputType(report.outputTypes[0]);
  }, [report]);

  function toggleBrand(b: string) {
    setBrands(prev => prev.includes(b) ? prev.filter(x => x !== b) : [...prev, b]);
  }

  // Filter available output types based on selected report
  const availableOutputTypes = report?.outputTypes ?? ALL_BRANDS.map(() => 'Excel');
  const filteredReports = reports.filter(r =>
    r.brands.some(b => brands.includes(b))
  );

  async function handleGenerate() {
    setError(null);
    setSuccess(null);

    if (!brands.length) { setError('Select at least one brand.'); return; }
    if (!report)        { setError('Select a report type.'); return; }
    if (!file)          { setError('Upload a raw data file.'); return; }

    setGenerating(true);

    const fd = new FormData();
    fd.append('file', file);
    fd.append('brand', brands[0]); // primary brand
    fd.append('reportId', report.id);
    fd.append('outputType', outputType);

    try {
      const res = await fetch('/api/generate', { method: 'POST', body: fd });

      if (!res.ok) {
        const body = await res.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(body.error || `Server error ${res.status}`);
      }

      const disposition = res.headers.get('Content-Disposition') ?? '';
      const match = disposition.match(/filename="([^"]+)"/);
      const filename = match?.[1] ?? 'report.xlsx';

      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setSuccess(`${filename} downloaded successfully.`);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to generate report.');
    } finally {
      setGenerating(false);
    }
  }

  if (loading || !session) return null;

  const step1Done = brands.length > 0;
  const step2Done = !!report;
  const step3Done = true; // output type always has a default
  const step4Done = !!file;

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

      <main className="max-w-2xl mx-auto px-4 py-10 space-y-4">

        {/* Page title */}
        <div className="text-center mb-2">
          <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Generate Report</h1>
          <p className="text-gray-500 text-sm mt-1">Select your filters, upload the Perigee export, and download the formatted report.</p>
        </div>

        {/* ── Step 1: Brands ──────────────────────────────────────────────── */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <div className="flex items-center gap-3 mb-4">
            <StepBadge n={1} done={step1Done} />
            <div>
              <p className="font-semibold text-gray-900 text-sm">Select Brand(s)</p>
              <p className="text-xs text-gray-400">Multiple brands can be selected</p>
            </div>
          </div>
          <div className="flex flex-wrap gap-3">
            {ALL_BRANDS.map(b => (
              <BrandPill key={b} label={b} selected={brands.includes(b)} onClick={() => toggleBrand(b)} />
            ))}
          </div>
          {brands.length > 0 && (
            <p className="text-xs text-gray-400 mt-3">Selected: {brands.join(', ')}</p>
          )}
        </div>

        {/* ── Step 2: Report type ──────────────────────────────────────────── */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <div className="flex items-center gap-3 mb-4">
            <StepBadge n={2} done={step2Done} />
            <div>
              <p className="font-semibold text-gray-900 text-sm">Select Report</p>
              <p className="text-xs text-gray-400">One report at a time — searchable</p>
            </div>
          </div>
          {reports.length === 0 ? (
            <p className="text-sm text-gray-400">Loading reports…</p>
          ) : (
            <ReportDropdown
              reports={filteredReports.length > 0 ? filteredReports : reports}
              value={report}
              onChange={r => setReport(r)}
            />
          )}
        </div>

        {/* ── Step 3: Output type ──────────────────────────────────────────── */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <div className="flex items-center gap-3 mb-4">
            <StepBadge n={3} done={step3Done} />
            <div>
              <p className="font-semibold text-gray-900 text-sm">Output Format</p>
              <p className="text-xs text-gray-400">Select one output type</p>
            </div>
          </div>
          <div className="flex gap-3">
            {(report?.outputTypes ?? ['Excel', 'PPT']).map(t => (
              <TypePill key={t} label={t} selected={outputType === t} onClick={() => setOutputType(t)} />
            ))}
          </div>
        </div>

        {/* ── Step 4: Upload ───────────────────────────────────────────────── */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <div className="flex items-center gap-3 mb-4">
            <StepBadge n={4} done={step4Done} />
            <div>
              <p className="font-semibold text-gray-900 text-sm">Upload Raw Data</p>
              <p className="text-xs text-gray-400">Perigee Excel export (.xlsx)</p>
            </div>
          </div>
          <DropZone file={file} onChange={setFile} />
        </div>

        {/* ── Errors / Success ─────────────────────────────────────────────── */}
        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 rounded-xl px-4 py-3 text-sm flex items-start gap-2">
            <svg className="w-4 h-4 mt-0.5 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            {error}
          </div>
        )}
        {success && (
          <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl px-4 py-3 text-sm flex items-start gap-2">
            <svg className="w-4 h-4 mt-0.5 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            {success}
          </div>
        )}

        {/* ── Generate button ───────────────────────────────────────────────── */}
        <button
          type="button"
          onClick={handleGenerate}
          disabled={generating || !brands.length || !report || !file}
          className="w-full bg-[#E31837] hover:bg-[#c01430] disabled:bg-gray-200 disabled:text-gray-400 text-white font-bold py-3.5 rounded-2xl text-base transition-all shadow-sm disabled:shadow-none"
        >
          {generating ? (
            <span className="flex items-center justify-center gap-2">
              <svg className="animate-spin w-5 h-5" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
              </svg>
              Generating…
            </span>
          ) : (
            `Generate ${report?.name ?? 'Report'}`
          )}
        </button>

      </main>
    </div>
  );
}

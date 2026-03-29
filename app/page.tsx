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

const ALL_BRANDS = ['Defy', 'Beko'];

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

// ─── Multi-select problem dropdown ───────────────────────────────────────────
function ProblemMultiSelect({
  problems,
  selected,
  onChange,
  placeholder,
}: {
  problems:    string[];
  selected:    string[];
  onChange:    (v: string[]) => void;
  placeholder: string;
}) {
  const [open, setOpen] = useState(false);
  const ref             = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const toggle = (p: string) =>
    onChange(selected.includes(p) ? selected.filter(x => x !== p) : [...selected, p]);

  const selectAll   = () => onChange([...problems]);
  const deselectAll = () => onChange([]);

  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        className={`w-full flex items-center justify-between px-4 py-2.5 rounded-xl border-2 text-sm transition-all bg-white ${
          open ? 'border-[#E31837] ring-2 ring-[#E31837]/20' : 'border-gray-200 hover:border-gray-300'
        }`}
      >
        <span className={selected.length > 0 ? 'text-gray-900 font-medium' : 'text-gray-400'}>
          {selected.length === 0
            ? placeholder
            : `${selected.length} of ${problems.length} problem${problems.length === 1 ? '' : 's'} selected`}
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
          {/* Quick actions */}
          <div className="flex gap-3 px-4 py-2 border-b border-gray-100 bg-gray-50">
            <button type="button" onClick={selectAll}   className="text-xs text-[#E31837] font-medium hover:underline">Select all</button>
            <button type="button" onClick={deselectAll} className="text-xs text-gray-500 font-medium hover:underline">Clear</button>
          </div>
          {/* Options */}
          <div className="max-h-56 overflow-y-auto">
            {problems.map(p => (
              <label key={p} className="flex items-center gap-3 px-4 py-2.5 cursor-pointer hover:bg-red-50 transition-colors">
                <input
                  type="checkbox"
                  checked={selected.includes(p)}
                  onChange={() => toggle(p)}
                  className="accent-[#E31837] w-4 h-4 shrink-0"
                />
                <span className="text-sm text-gray-700">{p}</span>
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Page ─────────────────────────────────────────────────────────────────────
export default function DashboardPage() {
  const { session, loading, logout } = useAuth();

  const [reports,           setReports]           = useState<ReportDef[]>([]);
  const [brands,            setBrands]            = useState<string[]>(['Defy']);
  const [report,            setReport]            = useState<ReportDef | null>(null);
  const [outputType,        setOutputType]        = useState<string>('Excel');
  const [file,              setFile]              = useState<File | null>(null);
  const [sendEmail,         setSendEmail]         = useState(false);
  const [additionalEmail,   setAdditionalEmail]   = useState('');
  const [generating,        setGenerating]        = useState(false);
  const [error,             setError]             = useState<string | null>(null);
  const [success,           setSuccess]           = useState<string | null>(null);
  const [warnings,          setWarnings]          = useState<string[] | null>(null);
  const [largeFileWarning,  setLargeFileWarning]  = useState(false);

  // Red flag problem distribution
  const [problems,          setProblems]          = useState<string[]>([]);
  const [salesProblems,     setSalesProblems]     = useState<string[]>([]);
  const [marketingProblems, setMarketingProblems] = useState<string[]>([]);
  const [problemsLoading,   setProblemsLoading]   = useState(false);

  const loadReports = useCallback(async () => {
    const res = await fetch('/api/reports');
    if (res.ok) setReports(await res.json());
  }, []);

  useEffect(() => { loadReports(); }, [loadReports]);

  // Derived: is the selected report a red flag report?
  const isRedFlag = !!report && (report.id.endsWith('-red-flag') || report.id === 'red-flag');

  // When red flag + file changes, parse problem list from the file
  useEffect(() => {
    if (!isRedFlag || !file) {
      setProblems([]);
      setSalesProblems([]);
      setMarketingProblems([]);
      return;
    }
    let cancelled = false;
    async function parse() {
      setProblemsLoading(true);
      setSalesProblems([]);
      setMarketingProblems([]);
      try {
        const fd = new FormData();
        fd.append('file', file!);
        const res = await fetch('/api/parse-problems', { method: 'POST', body: fd });
        if (!cancelled && res.ok) {
          const { problems: parsed } = await res.json() as { problems: string[] };
          setProblems(parsed);
        }
      } finally {
        if (!cancelled) setProblemsLoading(false);
      }
    }
    parse();
    return () => { cancelled = true; };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isRedFlag, file]);

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

  async function handleGenerate(confirmed = false, ignoreLargeFile = false) {
    setError(null);
    setSuccess(null);
    if (!confirmed) setWarnings(null);

    if (!brands.length) { setError('Select at least one brand.'); return; }
    if (!report)        { setError('Select a report type.'); return; }
    if (!file)          { setError('Upload a raw data file.'); return; }

    // Warn before emailing if the uploaded file exceeds 4 MB
    if (!ignoreLargeFile && sendEmail && file && file.size > 4 * 1024 * 1024) {
      setLargeFileWarning(true);
      return;
    }

    if (isRedFlag) {
      if (problemsLoading) { setError('Still parsing problems from the file — please wait.'); return; }
      if (!salesProblems.length && !marketingProblems.length) {
        setError('Select at least one problem for the Sales or Marketing report.');
        return;
      }
    }

    setGenerating(true);

    const fd = new FormData();
    fd.append('file', file);
    fd.append('brand', brands[0]);
    fd.append('reportId', report.id);
    fd.append('outputType', outputType);
    fd.append('userName',  session?.name  ?? '');
    fd.append('userEmail', session?.email ?? '');
    fd.append('sendEmail', sendEmail ? 'true' : 'false');
    if (sendEmail && additionalEmail) fd.append('additionalEmail', additionalEmail);
    if (confirmed) fd.append('confirmed', 'true');
    if (isRedFlag) {
      fd.append('salesProblems',     JSON.stringify(salesProblems));
      fd.append('marketingProblems', JSON.stringify(marketingProblems));
    }

    try {
      const res  = await fetch('/api/generate', { method: 'POST', body: fd });
      const body = await res.json().catch(() => ({ error: 'Unknown error' }));

      if (!res.ok) throw new Error(body.error || `Server error ${res.status}`);

      // 200 with warnings — show confirmation card
      if (body.warnings?.length) {
        setWarnings(body.warnings as string[]);
        return;
      }

      const names: string[] = body.filenames ?? (body.filename ? [body.filename as string] : []);
      setSuccess(`${names.join(' + ')} saved to SharePoint${sendEmail ? ' and sent by email' : ''}.`);
      setWarnings(null);
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
          <p className="text-gray-500 text-sm mt-1">Select your filters, upload the Perigee export, and save the formatted report to SharePoint.</p>
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

        {/* ── Red Flag: Problem distribution ───────────────────────────────── */}
        {isRedFlag && (
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 bg-[#E31837]">
                <svg className="w-3.5 h-3.5 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M8 7h12m0 0l-4-4m4 4l-4 4M4 17h12M4 17l4 4M4 17l4-4" />
                </svg>
              </div>
              <div>
                <p className="font-semibold text-gray-900 text-sm">Split Problems by Team</p>
                <p className="text-xs text-gray-400">Assign each problem type to the Sales and/or Marketing report</p>
              </div>
            </div>

            {!file ? (
              <p className="text-sm text-gray-400">Upload a file above to see the available problems.</p>
            ) : problemsLoading ? (
              <p className="text-sm text-gray-400 flex items-center gap-2">
                <svg className="animate-spin w-4 h-4 text-[#E31837]" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                </svg>
                Parsing problems from file…
              </p>
            ) : problems.length === 0 ? (
              <p className="text-sm text-amber-600">No problem values found in the uploaded file. Check that column P is populated.</p>
            ) : (
              <div className="space-y-5">
                {/* Sales */}
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="text-xs font-semibold text-blue-700 uppercase tracking-wide">Sales Red Flag</label>
                    {salesProblems.length > 0 && (
                      <span className="text-xs text-gray-400">{salesProblems.length} selected</span>
                    )}
                  </div>
                  <ProblemMultiSelect
                    problems={problems}
                    selected={salesProblems}
                    onChange={setSalesProblems}
                    placeholder="Select problems for the sales report…"
                  />
                  {salesProblems.length > 0 && (
                    <div className="flex flex-wrap gap-1.5 mt-2">
                      {salesProblems.map(p => (
                        <span key={p} className="px-2 py-0.5 bg-blue-50 text-blue-700 rounded text-xs font-medium">{p}</span>
                      ))}
                    </div>
                  )}
                </div>

                {/* Marketing */}
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="text-xs font-semibold text-orange-700 uppercase tracking-wide">Marketing Red Flag</label>
                    {marketingProblems.length > 0 && (
                      <span className="text-xs text-gray-400">{marketingProblems.length} selected</span>
                    )}
                  </div>
                  <ProblemMultiSelect
                    problems={problems}
                    selected={marketingProblems}
                    onChange={setMarketingProblems}
                    placeholder="Select problems for the marketing report…"
                  />
                  {marketingProblems.length > 0 && (
                    <div className="flex flex-wrap gap-1.5 mt-2">
                      {marketingProblems.map(p => (
                        <span key={p} className="px-2 py-0.5 bg-orange-50 text-orange-700 rounded text-xs font-medium">{p}</span>
                      ))}
                    </div>
                  )}
                </div>

                {/* Hint: problems not assigned to either */}
                {(() => {
                  const unassigned = problems.filter(
                    p => !salesProblems.includes(p) && !marketingProblems.includes(p),
                  );
                  if (!unassigned.length) return null;
                  return (
                    <p className="text-xs text-amber-600">
                      {unassigned.length} problem{unassigned.length === 1 ? '' : 's'} not yet assigned to either report:&nbsp;
                      <span className="font-medium">{unassigned.join(', ')}</span>
                    </p>
                  );
                })()}
              </div>
            )}
          </div>
        )}

        {/* ── Step 5: Delivery ─────────────────────────────────────────────── */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <div className="flex items-center gap-3 mb-4">
            <StepBadge n={5} done={false} />
            <div>
              <p className="font-semibold text-gray-900 text-sm">Email Delivery</p>
              <p className="text-xs text-gray-400">Report always saves to SharePoint automatically</p>
            </div>
          </div>

          {/* Toggle */}
          <label className="flex items-center gap-3 cursor-pointer select-none">
            <button
              type="button"
              role="switch"
              aria-checked={sendEmail}
              onClick={() => setSendEmail(v => !v)}
              className={`relative inline-flex h-6 w-11 shrink-0 items-center rounded-full transition-colors focus:outline-none ${
                sendEmail ? 'bg-[#E31837]' : 'bg-gray-200'
              }`}
            >
              <span
                className={`inline-block h-4 w-4 transform rounded-full bg-white shadow transition-transform ${
                  sendEmail ? 'translate-x-6' : 'translate-x-1'
                }`}
              />
            </button>
            <span className="text-sm font-medium text-gray-700">
              Email this report to me
              {session?.email && (
                <span className="ml-1.5 text-xs text-gray-400 font-normal">({session.email})</span>
              )}
            </span>
          </label>

          {/* Additional email — only shown when toggle is on */}
          {sendEmail && (
            <div className="mt-4 space-y-2">
              <label className="block text-xs font-medium text-gray-600 uppercase tracking-wide">
                Additional recipient <span className="font-normal text-gray-400 normal-case">(optional)</span>
              </label>
              <input
                type="email"
                value={additionalEmail}
                onChange={e => setAdditionalEmail(e.target.value)}
                placeholder="colleague@example.com"
                className="w-full px-4 py-2.5 rounded-xl border-2 border-gray-200 text-sm focus:outline-none focus:border-[#E31837] focus:ring-2 focus:ring-[#E31837]/20 transition-all"
              />
            </div>
          )}
        </div>

        {/* ── Large file email warning ─────────────────────────────────────── */}
        {largeFileWarning && (
          <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-4 space-y-3">
            <div className="flex items-start gap-2">
              <svg className="w-5 h-5 mt-0.5 shrink-0 text-amber-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" />
              </svg>
              <div>
                <p className="font-semibold text-amber-800 text-sm">This file may be too large to email</p>
                <p className="text-sm text-amber-700 mt-1">
                  PPT reports with images can be very large and will likely fail as an email attachment.
                  Are you sure you want to email it? Remember — the report is always saved to SharePoint first, where you can WeTransfer it to your client.
                </p>
              </div>
            </div>
            <div className="flex gap-3 pt-1">
              <button
                type="button"
                onClick={() => { setLargeFileWarning(false); handleGenerate(false, true); }}
                disabled={generating}
                className="px-5 py-2 rounded-xl bg-[#E31837] hover:bg-[#c01430] disabled:bg-gray-200 disabled:text-gray-400 text-white text-sm font-semibold transition-all"
              >
                {generating ? 'Generating…' : 'Email it anyway'}
              </button>
              <button
                type="button"
                onClick={() => setLargeFileWarning(false)}
                disabled={generating}
                className="px-5 py-2 rounded-xl border border-gray-300 bg-white hover:bg-gray-50 text-gray-700 text-sm font-semibold transition-all"
              >
                Cancel
              </button>
            </div>
          </div>
        )}

        {/* ── Warnings confirmation card ───────────────────────────────────── */}
        {warnings && (
          <div className="bg-amber-50 border border-amber-300 rounded-xl px-4 py-4 space-y-3">
            <div className="flex items-start gap-2">
              <svg className="w-5 h-5 mt-0.5 shrink-0 text-amber-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" />
              </svg>
              <div>
                <p className="font-semibold text-amber-800 text-sm">Data warnings detected</p>
                <p className="text-xs text-amber-700 mt-0.5">The report can still be generated, but some data may be incomplete. Review the issues below before continuing.</p>
              </div>
            </div>
            <ul className="space-y-1.5 pl-2">
              {warnings.map((w, i) => (
                <li key={i} className="flex items-start gap-2 text-sm text-amber-900">
                  <span className="mt-1.5 w-1.5 h-1.5 rounded-full bg-amber-500 shrink-0" />
                  {w}
                </li>
              ))}
            </ul>
            <div className="flex gap-3 pt-1">
              <button
                type="button"
                onClick={() => handleGenerate(true)}
                disabled={generating}
                className="px-5 py-2 rounded-xl bg-[#E31837] hover:bg-[#c01430] disabled:bg-gray-200 disabled:text-gray-400 text-white text-sm font-semibold transition-all"
              >
                {generating ? 'Generating…' : 'Continue anyway'}
              </button>
              <button
                type="button"
                onClick={() => setWarnings(null)}
                disabled={generating}
                className="px-5 py-2 rounded-xl border border-gray-300 bg-white hover:bg-gray-50 text-gray-700 text-sm font-semibold transition-all"
              >
                Cancel
              </button>
            </div>
          </div>
        )}

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
          onClick={() => handleGenerate()}
          disabled={
            generating || !brands.length || !report || !file ||
            (isRedFlag && (problemsLoading || (!salesProblems.length && !marketingProblems.length)))
          }
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

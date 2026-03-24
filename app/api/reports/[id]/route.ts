import { NextRequest, NextResponse } from 'next/server';
import { loadReports, saveReports } from '@/lib/reportData';

export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const updates = await req.json();
  const reports = loadReports();
  const idx = reports.findIndex(r => r.id === id);
  if (idx === -1) return NextResponse.json({ error: 'Not found' }, { status: 404 });

  reports[idx] = { ...reports[idx], ...updates };
  if (updates.name) reports[idx].name = String(updates.name).toUpperCase().trim();
  await saveReports(reports);
  return NextResponse.json(reports[idx]);
}

export async function DELETE(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const reports = loadReports();
  const idx = reports.findIndex(r => r.id === id);
  if (idx === -1) return NextResponse.json({ error: 'Not found' }, { status: 404 });

  reports.splice(idx, 1);
  await saveReports(reports);
  return NextResponse.json({ ok: true });
}

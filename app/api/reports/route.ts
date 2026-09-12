import { NextRequest, NextResponse } from 'next/server';
import { randomUUID } from 'crypto';
import { loadReports, saveReports, ReportDef } from '@/lib/reportData';
import { requireAdmin, requireUser } from '@/lib/apiAuth';

// The main page lists reports for every signed-in user, so this is not
// admin-only - but it is no longer open to the public internet either.
export async function GET() {
  const guard = await requireUser();
  if (guard.deny) return guard.deny;

  return NextResponse.json(await loadReports());
}

// Defining a report is an admin action (only the admin page posts here).
export async function POST(req: NextRequest) {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  const { name, dataFormat, channel, outputTypes, brands } = await req.json();
  if (!name || !dataFormat || !outputTypes?.length || !brands?.length) {
    return NextResponse.json({ error: 'Missing required fields (name, dataFormat, outputTypes, brands)' }, { status: 400 });
  }

  const reports = await loadReports();
  const report: ReportDef = {
    id: randomUUID(),
    name: String(name).toUpperCase().trim(),
    dataFormat: String(dataFormat).trim(),
    ...(channel ? { channel: String(channel).toUpperCase().trim() } : {}),
    outputTypes,
    brands,
  };
  reports.push(report);
  await saveReports(reports);
  return NextResponse.json(report, { status: 201 });
}

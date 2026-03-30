import { NextRequest, NextResponse } from 'next/server';
import { randomUUID } from 'crypto';
import { loadReports, saveReports, ReportDef } from '@/lib/reportData';

export async function GET() {
  return NextResponse.json(loadReports());
}

export async function POST(req: NextRequest) {
  const { name, dataFormat, channel, outputTypes, brands } = await req.json();
  if (!name || !dataFormat || !outputTypes?.length || !brands?.length) {
    return NextResponse.json({ error: 'Missing required fields (name, dataFormat, outputTypes, brands)' }, { status: 400 });
  }

  const reports = loadReports();
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

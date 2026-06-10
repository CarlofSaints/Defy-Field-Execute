import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import {
  loadCriticalLines,
  saveCriticalLines,
  type CriticalLinesConfig,
  type CriticalLineEntry,
} from '@/lib/criticalLinesData';

export async function GET() {
  const configs = await loadCriticalLines();
  return NextResponse.json(configs);
}

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: 'Invalid form data' }, { status: 400 });
  }

  const file    = formData.get('file')    as File   | null;
  const brand   = formData.get('brand')   as string | null;
  const channel = formData.get('channel') as string | null;

  if (!file)    return NextResponse.json({ error: 'No file provided' }, { status: 400 });
  if (!brand)   return NextResponse.json({ error: 'Brand is required' }, { status: 400 });
  if (!channel) return NextResponse.json({ error: 'Channel is required' }, { status: 400 });

  const buffer = Buffer.from(await file.arrayBuffer());
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (string | null)[][];

  if (rows.length < 2) {
    return NextResponse.json({ error: 'File has no data rows' }, { status: 400 });
  }

  const headers = (rows[0] as (string | null)[]).map(h => (h ?? '').trim().toUpperCase());

  const modelIdx = headers.findIndex(h =>
    ['MODEL NUMBER', 'MODEL NO', 'MODEL', 'MODEL_NUMBER', 'MODELNUMBER', 'PRODUCT CODE', 'CODE'].includes(h),
  );
  const descIdx = headers.findIndex(h =>
    ['PRODUCT DESCRIPTION', 'DESCRIPTION', 'PRODUCT', 'PRODUCT_DESCRIPTION', 'DESC', 'PRODUCT NAME'].includes(h),
  );

  if (modelIdx === -1) {
    return NextResponse.json(
      { error: 'Could not find a MODEL NUMBER column. Expected header: MODEL NUMBER, MODEL NO, MODEL, or PRODUCT CODE' },
      { status: 400 },
    );
  }

  const entries: CriticalLineEntry[] = [];
  for (const row of rows.slice(1)) {
    const modelNumber = String(row[modelIdx] ?? '').trim().toUpperCase().replace(/\s+/g, '');
    if (!modelNumber) continue;
    const productDescription = descIdx >= 0 ? String(row[descIdx] ?? '').trim() : '';
    entries.push({ modelNumber, productDescription });
  }

  if (entries.length === 0) {
    return NextResponse.json({ error: 'No valid rows found in file' }, { status: 400 });
  }

  // Upsert: replace existing config for same brand+channel, or add new
  const configs = await loadCriticalLines();
  const b = brand.toUpperCase();
  const c = channel.toUpperCase();
  const idx = configs.findIndex(
    cfg => cfg.brand.toUpperCase() === b && cfg.channel.toUpperCase() === c,
  );

  const newConfig: CriticalLinesConfig = { brand: b, channel: c, lines: entries };
  if (idx >= 0) {
    configs[idx] = newConfig;
  } else {
    configs.push(newConfig);
  }

  await saveCriticalLines(configs);
  return NextResponse.json({ count: entries.length, brand: b, channel: c });
}

export async function DELETE(req: NextRequest) {
  const { brand, channel } = await req.json() as { brand?: string; channel?: string };
  if (!brand || !channel) {
    return NextResponse.json({ error: 'brand and channel are required' }, { status: 400 });
  }

  const configs = await loadCriticalLines();
  const b = brand.toUpperCase();
  const c = channel.toUpperCase();
  const filtered = configs.filter(
    cfg => !(cfg.brand.toUpperCase() === b && cfg.channel.toUpperCase() === c),
  );

  if (filtered.length === configs.length) {
    return NextResponse.json({ error: 'Config not found' }, { status: 404 });
  }

  await saveCriticalLines(filtered);
  return NextResponse.json({ ok: true });
}

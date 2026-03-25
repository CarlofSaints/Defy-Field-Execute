import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { loadStoreMap, saveStoreMap, StoreMapEntry } from '@/lib/storeMapData';

export async function GET() {
  const map = loadStoreMap();
  return NextResponse.json({ count: map.length, entries: map });
}

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: 'Invalid form data' }, { status: 400 });
  }

  const file = formData.get('file') as File | null;
  if (!file) return NextResponse.json({ error: 'No file provided' }, { status: 400 });

  const buffer = Buffer.from(await file.arrayBuffer());
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (string | null)[][];

  if (rows.length < 2) {
    return NextResponse.json({ error: 'File has no data rows' }, { status: 400 });
  }

  // Find column indices by header name (case-insensitive)
  const headers = (rows[0] as (string | null)[]).map(h => (h ?? '').trim().toUpperCase());
  const nameIdx     = headers.findIndex(h => ['STORE NAME', 'STORENAME', 'NAME'].includes(h));
  const codeIdx     = headers.findIndex(h => ['STORE CODE', 'STORECODE', 'CODE', 'SITE CODE', 'SITECODE'].includes(h));
  const provinceIdx = headers.findIndex(h => ['PROVINCE', 'REGION', 'PROV'].includes(h));

  if (provinceIdx === -1) {
    return NextResponse.json({ error: 'Could not find a PROVINCE column. Expected header: PROVINCE or REGION' }, { status: 400 });
  }
  if (nameIdx === -1 && codeIdx === -1) {
    return NextResponse.json({ error: 'Could not find STORE NAME or STORE CODE column' }, { status: 400 });
  }

  const entries: StoreMapEntry[] = [];
  for (const row of rows.slice(1)) {
    const storeName = nameIdx >= 0 ? String(row[nameIdx] ?? '').trim() : '';
    const storeCode = codeIdx >= 0 ? String(row[codeIdx] ?? '').trim() : '';
    const province  = String(row[provinceIdx] ?? '').trim().toUpperCase();
    if (!province || (!storeName && !storeCode)) continue;
    entries.push({ storeName, storeCode, province });
  }

  if (entries.length === 0) {
    return NextResponse.json({ error: 'No valid rows found in file' }, { status: 400 });
  }

  await saveStoreMap(entries);
  return NextResponse.json({ count: entries.length });
}

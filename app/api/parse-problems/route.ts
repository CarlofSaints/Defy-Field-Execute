import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

// POST — accepts a Perigee red-flag Excel file, returns unique problem values
// from col 15 (0-indexed) = "WHAT IS THE PROBLEM?"
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
  const wb     = XLSX.read(buffer, { type: 'buffer' });
  const ws     = wb.Sheets[wb.SheetNames[0]];
  const rows   = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (string | null)[][];

  if (rows.length < 2) return NextResponse.json({ problems: [] });

  const problems = [
    ...new Set(
      rows.slice(1)
        .map(row => String(row[15] ?? '').trim())
        .filter(Boolean),
    ),
  ].sort();

  return NextResponse.json({ problems });
}

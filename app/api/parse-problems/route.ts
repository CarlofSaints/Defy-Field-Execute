import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

// POST — accepts a Perigee red-flag Excel file, returns unique problem values
// looked up by header name "WHAT IS THE PROBLEM?" (order-independent).
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
  const rows   = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (string | number | null)[][];

  if (rows.length < 2) return NextResponse.json({ problems: [] });

  // Header-based column lookup (case/whitespace insensitive)
  const headerRow = rows[0] || [];
  const norm      = (v: unknown) => String(v ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
  const problemColIdx: number[] = [];
  headerRow.forEach((v, i) => { if (norm(v) === 'what is the problem?') problemColIdx.push(i); });

  if (problemColIdx.length === 0) {
    return NextResponse.json(
      { error: 'Could not find "WHAT IS THE PROBLEM?" column in uploaded file.' },
      { status: 400 },
    );
  }

  const problems = [
    ...new Set(
      rows.slice(1).map(row => {
        for (const c of problemColIdx) {
          const v = row[c];
          if (v != null) {
            const s = String(v).trim();
            if (s !== '') return s;
          }
        }
        return '';
      }).filter(Boolean),
    ),
  ].sort();

  return NextResponse.json({ problems });
}

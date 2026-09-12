import { NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { requireAdmin } from '@/lib/apiAuth';

// Returns a blank Store Mapping control-file template with the expected headers
// and a couple of example rows, so users know exactly what to fill in.
// Linked with a plain <a href> from the admin page. A cookie is sent on a
// same-origin navigation, so gating this does not break the download.
export async function GET() {
  const guard = await requireAdmin();
  if (guard.deny) return guard.deny;

  const rows: (string)[][] = [
    ['STORE NAME', 'STORE CODE', 'PROVINCE'],
    ['MAKRO WOODMEAD', 'M123', 'GAUTENG'],
    ['GAME GATEWAY MALL', 'G075', 'KWAZULU-NATAL'],
  ];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 34 }, { wch: 14 }, { wch: 22 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Store Map');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;

  return new NextResponse(new Uint8Array(buf), {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="Store Mapping Template.xlsx"',
    },
  });
}

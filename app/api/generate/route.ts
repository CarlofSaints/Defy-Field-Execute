import { NextRequest, NextResponse } from 'next/server';
import { generateMakroStockCount } from '@/lib/reports/makro-stock-count';

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: 'Invalid form data' }, { status: 400 });
  }

  const file      = formData.get('file') as File | null;
  const brand     = (formData.get('brand') as string | null)?.trim() || '';
  const reportId  = (formData.get('reportId') as string | null)?.trim() || '';
  const outputType = (formData.get('outputType') as string | null)?.trim() || 'Excel';

  if (!file || !brand || !reportId) {
    return NextResponse.json({ error: 'file, brand and reportId are required' }, { status: 400 });
  }

  const fileBuffer = Buffer.from(await file.arrayBuffer());

  try {
    let excelBuffer: Buffer;
    let filename: string;

    if (reportId === 'makro-stock-count') {
      ({ buffer: excelBuffer, filename } = await generateMakroStockCount(fileBuffer, brand));
    } else {
      return NextResponse.json(
        { error: `Report "${reportId}" is not yet implemented.` },
        { status: 422 },
      );
    }

    const mimeType = outputType === 'PPT'
      ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    return new Response(new Uint8Array(excelBuffer), {
      headers: {
        'Content-Type': mimeType,
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Unknown error';
    console.error('[generate]', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

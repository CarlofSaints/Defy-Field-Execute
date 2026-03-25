import { NextRequest, NextResponse } from 'next/server';
import { after } from 'next/server';
import { randomUUID } from 'crypto';
import { generateMakroStockCount } from '@/lib/reports/makro-stock-count';
import { loadStoreMap } from '@/lib/storeMapData';
import { loadReports } from '@/lib/reportData';
import { loadUsers } from '@/lib/userData';
import { addRunEntry } from '@/lib/runLogData';
import { sendRunNotification, sendReportEmail } from '@/lib/email';
import { buildDfeFolderPath, uploadToSharePoint } from '@/lib/sharepoint-dfe';
import type { DfeBrand } from '@/lib/sharepoint-dfe';
import type { RunLogEntry } from '@/lib/runLogData';

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: 'Invalid form data' }, { status: 400 });
  }

  const file            = formData.get('file') as File | null;
  const brand           = (formData.get('brand') as string | null)?.trim() || '';
  const reportId        = (formData.get('reportId') as string | null)?.trim() || '';
  const outputType      = (formData.get('outputType') as string | null)?.trim() || 'Excel';
  const userName        = (formData.get('userName') as string | null)?.trim() || 'Unknown';
  const userEmail       = (formData.get('userEmail') as string | null)?.trim() || '';
  const sendEmail       = formData.get('sendEmail') === 'true';
  const additionalEmail = (formData.get('additionalEmail') as string | null)?.trim() || '';

  if (!file || !brand || !reportId) {
    return NextResponse.json({ error: 'file, brand and reportId are required' }, { status: 400 });
  }

  const fileBuffer    = Buffer.from(await file.arrayBuffer());
  const rawFilename   = file.name;
  const storeMap      = loadStoreMap();

  const reports    = loadReports();
  const reportDef  = reports.find(r => r.id === reportId);
  const reportName = reportDef?.name ?? reportId.replace(/-/g, ' ').toUpperCase();

  const timestamp = new Date().toISOString();

  try {
    let excelBuffer: Buffer;
    let filename: string;
    let rawDates: string[] = [];
    let weekLabel = '';
    let retailer  = '';

    if (reportId.endsWith('-stock-count')) {
      retailer = reportId
        .replace(/-stock-count$/, '')
        .replace(/-/g, ' ')
        .toUpperCase();
      ({ buffer: excelBuffer, filename, rawDates, weekLabel } = await generateMakroStockCount(
        fileBuffer, brand, storeMap, retailer,
      ));
    } else {
      return NextResponse.json(
        { error: `Report "${reportId}" is not yet implemented.` },
        { status: 422 },
      );
    }

    // Capture for after() closure
    const _excelBuffer  = excelBuffer;
    const _rawBuffer    = fileBuffer;
    const _rawFilename  = rawFilename;
    const _filename     = filename;
    const _rawDates     = rawDates;
    const _weekLabel    = weekLabel;
    const _retailer     = retailer;

    after(async () => {
      let spPath          = '';
      let spFolderUrl     = '';
      let emailSent       = false;
      let emailRecipients: string[] = [];

      // ── Always upload to SharePoint ──────────────────────────────────────
      try {
        const dfeBrand  = brand.toUpperCase() as DfeBrand;
        const folderPath = buildDfeFolderPath(dfeBrand, _rawDates);
        const fileWebUrl = await uploadToSharePoint(folderPath, _filename, _excelBuffer);
        spPath      = fileWebUrl;
        spFolderUrl = fileWebUrl.substring(0, fileWebUrl.lastIndexOf('/'));
        // Also upload the original raw file alongside the generated report
        await uploadToSharePoint(folderPath, _rawFilename, _rawBuffer).catch((e: unknown) => {
          console.error('[generate] Raw file SP upload failed:', e instanceof Error ? e.message : e);
        });
      } catch (spErr) {
        console.error('[generate] SP upload failed:', spErr instanceof Error ? spErr.message : spErr);
      }

      // ── Optionally email the report ───────────────────────────────────────
      if (sendEmail && userEmail) {
        try {
          const recipients = [userEmail];
          if (additionalEmail) recipients.push(additionalEmail);
          const firstName = userName.split(' ')[0] || userName;
          await sendReportEmail({
            to:          recipients,
            firstName,
            reportName,
            brand,
            weekLabel:   _weekLabel,
            filename:    _filename,
            fileBuffer:  _excelBuffer,
            spFolderUrl: spFolderUrl || spPath,
          });
          emailSent       = true;
          emailRecipients = recipients;
        } catch (mailErr) {
          console.error('[generate] email send failed:', mailErr instanceof Error ? mailErr.message : mailErr);
        }
      }

      // ── Log the run ───────────────────────────────────────────────────────
      const entry: RunLogEntry = {
        id: randomUUID(),
        timestamp,
        userName,
        userEmail,
        reportId,
        reportName,
        brand,
        retailer: _retailer,
        filename: _filename,
        status:   'success',
        spPath:   spPath || undefined,
        emailSent,
        emailRecipients: emailSent ? emailRecipients : undefined,
      };
      await addRunEntry(entry);

      const adminEmails = loadUsers()
        .filter(u => u.isAdmin)
        .map(u => u.email);
      if (adminEmails.length) {
        await sendRunNotification(adminEmails, entry).catch(() => { /* non-fatal */ });
      }
    });

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

    after(async () => {
      const retailer = reportId.endsWith('-stock-count')
        ? reportId.replace(/-stock-count$/, '').replace(/-/g, ' ').toUpperCase()
        : '';
      const entry: RunLogEntry = {
        id: randomUUID(),
        timestamp,
        userName,
        userEmail,
        reportId,
        reportName,
        brand,
        retailer,
        filename:     '',
        status:       'error',
        errorMessage: message,
        emailSent:    false,
      };
      await addRunEntry(entry);

      const adminEmails = loadUsers()
        .filter(u => u.isAdmin)
        .map(u => u.email);
      if (adminEmails.length) {
        await sendRunNotification(adminEmails, entry).catch(() => { /* non-fatal */ });
      }
    });

    return NextResponse.json({ error: message }, { status: 500 });
  }
}

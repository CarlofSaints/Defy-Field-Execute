import { NextRequest, NextResponse } from 'next/server';
import { after } from 'next/server';
import { randomUUID } from 'crypto';
import { generateMakroStockCount, analyzeStockCount } from '@/lib/reports/makro-stock-count';
import { generateRedFlag, extractRedFlagProblems } from '@/lib/reports/red-flag';
import { generateStandReport } from '@/lib/reports/stand-report';
import { generateTrainingFeedback } from '@/lib/reports/training-feedback';
import { generateActivationReport } from '@/lib/reports/activation-report';
import { loadStoreMap } from '@/lib/storeMapData';
import { loadReports } from '@/lib/reportData';
import { loadUsers } from '@/lib/userData';
import { addRunEntry } from '@/lib/runLogData';
import { sendRunNotification, sendReportEmail } from '@/lib/email';
import { buildDfeFolderPath, uploadToSharePoint } from '@/lib/sharepoint-dfe';
import type { DfeBrand } from '@/lib/sharepoint-dfe';
import type { RunLogEntry } from '@/lib/runLogData';

interface GenerateResult {
  excelBuffer: Buffer;
  filename:    string;
  rawDates:    string[];
  weekLabel:   string;
  retailer:    string;
  label:       string;  // 'SALES' | 'MARKETING' | ''
  contentType?: string; // defaults to Excel if omitted
}

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
  const confirmed       = formData.get('confirmed') === 'true';

  if (!file || !brand || !reportId) {
    return NextResponse.json({ error: 'file, brand and reportId are required' }, { status: 400 });
  }

  const fileBuffer  = Buffer.from(await file.arrayBuffer());
  const rawFilename = file.name;
  const storeMap    = loadStoreMap();

  const reports    = loadReports();
  const reportDef  = reports.find(r => r.id === reportId);
  const reportName = reportDef?.name ?? reportId.replace(/-/g, ' ').toUpperCase();

  const timestamp = new Date().toISOString();

  // ── Pre-generate analysis (skipped if user already confirmed) ───────────────
  if (!confirmed && reportId.endsWith('-stock-count')) {
    const { warnings, hardError } = analyzeStockCount(fileBuffer);
    if (hardError) {
      return NextResponse.json({ error: hardError }, { status: 422 });
    }
    if (warnings.length > 0) {
      return NextResponse.json({ warnings }, { status: 200 });
    }
  }

  try {
    const results: GenerateResult[] = [];
    void outputType; // reserved for future output-type switching

    if (reportId.endsWith('-stock-count')) {
      const retailer = reportId
        .replace(/-stock-count$/, '')
        .replace(/-/g, ' ')
        .toUpperCase();
      const { buffer, filename, rawDates, weekLabel } = await generateMakroStockCount(
        fileBuffer, brand, storeMap, retailer,
      );
      results.push({ excelBuffer: buffer, filename, rawDates, weekLabel, retailer, label: '' });

    } else if (reportId.endsWith('-red-flag') || reportId === 'red-flag') {
      // Parse problem selections from the form
      let salesProblems:     string[] = [];
      let marketingProblems: string[] = [];
      try {
        salesProblems     = JSON.parse((formData.get('salesProblems')     as string | null) || '[]');
        marketingProblems = JSON.parse((formData.get('marketingProblems') as string | null) || '[]');
      } catch {
        return NextResponse.json({ error: 'Invalid problem selection data.' }, { status: 400 });
      }

      if (!salesProblems.length && !marketingProblems.length) {
        return NextResponse.json(
          { error: 'Select at least one problem for the Sales or Marketing report.' },
          { status: 400 },
        );
      }

      // Validate: check selected problems exist in the actual data
      const dataProblems = extractRedFlagProblems(fileBuffer);
      const validationErrors: string[] = [];
      for (const p of salesProblems) {
        if (!dataProblems.has(p)) {
          validationErrors.push(`You have selected "${p}" for Sales but no lines exist in your data matching that problem.`);
        }
      }
      for (const p of marketingProblems) {
        if (!dataProblems.has(p)) {
          validationErrors.push(`You have selected "${p}" for Marketing but no lines exist in your data matching that problem.`);
        }
      }
      if (validationErrors.length) {
        return NextResponse.json({ error: validationErrors.join('\n') }, { status: 422 });
      }

      // Generate Sales report
      if (salesProblems.length) {
        const { buffer, filename, rawDates } = await generateRedFlag(fileBuffer, brand, 'SALES', salesProblems);
        results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: 'SALES' });
      }
      // Generate Marketing report
      if (marketingProblems.length) {
        const { buffer, filename, rawDates } = await generateRedFlag(fileBuffer, brand, 'MARKETING', marketingProblems);
        results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: 'MARKETING' });
      }

    } else if (reportId === 'stand-report') {
      const { buffer, filename, rawDates } = await generateStandReport(fileBuffer, brand);
      results.push({
        excelBuffer: buffer,
        filename,
        rawDates,
        weekLabel: '',
        retailer:  '',
        label:     '',
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      });

    } else if (reportId === 'training-feedback-report') {
      const { buffer, filename, rawDates } = await generateTrainingFeedback(fileBuffer, brand);
      results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: '' });

    } else if (reportId === 'activation-report') {
      const { buffer, filename, rawDates } = await generateActivationReport(fileBuffer, brand);
      results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: '' });

    } else {
      return NextResponse.json(
        { error: `Report "${reportId}" is not yet implemented.` },
        { status: 422 },
      );
    }

    // ── Email size warning ────────────────────────────────────────────────────
    // Warn the user before emailing a large file. SP upload always happens.
    if (sendEmail && !confirmed) {
      const sizeWarnings: string[] = [];
      for (const r of results) {
        const mb = r.excelBuffer.length / (1024 * 1024);
        if (mb > 4) {
          sizeWarnings.push(
            `"${r.filename}" is ${mb.toFixed(1)} MB. ` +
            `Large files may fail to deliver by email. ` +
            `The file will still be saved to SharePoint — you can download it from there and use WeTransfer to send to your client instead.`
          );
        }
      }
      if (sizeWarnings.length > 0) {
        return NextResponse.json({ warnings: sizeWarnings }, { status: 200 });
      }
    }

    // Capture for after() closure
    const _results      = results;
    const _rawBuffer    = fileBuffer;
    const _rawFilename  = rawFilename;

    after(async () => {
      for (const result of _results) {
        let spPath          = '';
        let spFolderUrl     = '';
        let emailSent       = false;
        let emailRecipients: string[] = [];

        // Combine rawDates across all results for folder path
        const allRawDates = _results.flatMap(r => r.rawDates);

        // ── Always upload to SharePoint ────────────────────────────────────
        try {
          const dfeBrand   = brand.toUpperCase() as DfeBrand;
          const folderPath = buildDfeFolderPath(dfeBrand, allRawDates);
          const fileWebUrl = await uploadToSharePoint(folderPath, result.filename, result.excelBuffer, result.contentType);
          spPath      = fileWebUrl;
          spFolderUrl = fileWebUrl.substring(0, fileWebUrl.lastIndexOf('/'));

          // Upload raw file only alongside the first result (avoid duplicates)
          if (_results.indexOf(result) === 0) {
            await uploadToSharePoint(folderPath, _rawFilename, _rawBuffer).catch((e: unknown) => {
              console.error('[generate] Raw file SP upload failed:', e instanceof Error ? e.message : e);
            });
          }
        } catch (spErr) {
          console.error('[generate] SP upload failed:', spErr instanceof Error ? spErr.message : spErr);
        }

        // ── Optionally email the report ────────────────────────────────────
        if (sendEmail && userEmail) {
          try {
            const recipients = [userEmail];
            if (additionalEmail) recipients.push(additionalEmail);
            const firstName  = userName.split(' ')[0] || userName;
            const label      = result.label ? `${result.label} ` : '';
            await sendReportEmail({
              to:          recipients,
              firstName,
              reportName:  `${label}${reportName}`,
              brand,
              weekLabel:   result.weekLabel,
              filename:    result.filename,
              fileBuffer:  result.excelBuffer,
              spFolderUrl: spFolderUrl || spPath,
            });
            emailSent       = true;
            emailRecipients = recipients;
          } catch (mailErr) {
            console.error('[generate] email send failed:', mailErr instanceof Error ? mailErr.message : mailErr);
          }
        }

        // ── Log the run ────────────────────────────────────────────────────
        const entry: RunLogEntry = {
          id: randomUUID(),
          timestamp,
          userName,
          userEmail,
          reportId,
          reportName: result.label ? `${result.label} ${reportName}` : reportName,
          brand,
          retailer:   result.retailer,
          filename:   result.filename,
          status:     'success',
          spPath:     spPath || undefined,
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
      }
    });

    return NextResponse.json({ success: true, filenames: results.map(r => r.filename) });

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

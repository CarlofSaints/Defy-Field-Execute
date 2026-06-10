import { NextRequest, NextResponse } from 'next/server';
import { after } from 'next/server';
import { randomUUID } from 'crypto';
import { generateMakroStockCount, analyzeStockCount, type DatabaseRow } from '@/lib/reports/makro-stock-count';
import { generateRedFlag, extractRedFlagProblems, type RedFlagArchiveRow } from '@/lib/reports/red-flag';
import { generateStandReport, type StandReportArchiveRow } from '@/lib/reports/stand-report';
import { generateTrainingFeedback, type TrainingSummaryRow } from '@/lib/reports/training-feedback';
import { loadMonthUploads, archiveUpload } from '@/lib/uploadArchive';
import { generateActivationReport, type ActivationArchiveRow } from '@/lib/reports/activation-report';
import { generateServiceCallReport, type ServiceCallArchiveRow } from '@/lib/reports/service-call-report';
import { analyzeChannelMismatch, filterByChannel } from '@/lib/reports/channel-check';
import { loadStoreMap } from '@/lib/storeMapData';
import { loadReports, DATA_FORMAT_LABELS } from '@/lib/reportData';
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

// ── Date helpers ──────────────────────────────────────────────────────────────

function parseDdMmYyyy(raw: string): Date | null {
  const m = raw.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);
  if (!m) return null;
  const d = new Date(+m[3], +m[2] - 1, +m[1]);
  return isNaN(d.getTime()) ? null : d;
}

function pad2(n: number) { return String(n).padStart(2, '0'); }

/**
 * Build a date-range string from raw DD/MM/YYYY date strings.
 * Single date  → "23-03-2026"
 * Same year    → "23-03 - 30-03-2026"
 * Cross year   → "28-12-2025 - 04-01-2026"
 */
function buildDateRange(rawDates: string[]): string {
  const valid = rawDates
    .map(parseDdMmYyyy)
    .filter((d): d is Date => d !== null)
    .sort((a, b) => a.getTime() - b.getTime());
  if (!valid.length) {
    const t = new Date();
    return `${pad2(t.getDate())}-${pad2(t.getMonth() + 1)}-${t.getFullYear()}`;
  }
  const first = valid[0];
  const last  = valid[valid.length - 1];
  const fStr  = `${pad2(first.getDate())}-${pad2(first.getMonth() + 1)}-${first.getFullYear()}`;
  const lStr  = `${pad2(last.getDate())}-${pad2(last.getMonth() + 1)}-${last.getFullYear()}`;
  if (fStr === lStr) return fStr;
  // Same year → abbreviate start (omit year)
  if (first.getFullYear() === last.getFullYear()) {
    return `${pad2(first.getDate())}-${pad2(first.getMonth() + 1)} - ${lStr}`;
  }
  return `${fStr} - ${lStr}`;
}

/**
 * Central filename builder.
 * Pattern: BRAND - CHANNEL - REPORT TYPE LABEL - DATE RANGE.ext
 * e.g.  "DEFY - MAKRO - STOCK COUNT - 23-03 - 30-03-2026.xlsx"
 *       "DEFY - RED FLAG SALES - 30-03-2026.xlsx"
 *       "BEKO - STAND REPORT - 23-03 - 30-03-2026.pptx"
 */
function buildFilename(
  brand:      string,
  dataFormat: string,
  channel:    string | undefined,
  rawDates:   string[],
  label?:     string,  // e.g. 'SALES' | 'MARKETING' for red flag
): string {
  const ext = dataFormat === 'stand-report' ? 'pptx' : 'xlsx';
  const formatLabel = DATA_FORMAT_LABELS[dataFormat] ?? dataFormat.replace(/-/g, ' ');
  const dateRange   = buildDateRange(rawDates);

  const parts = [brand.toUpperCase()];
  if (channel) parts.push(channel.toUpperCase());
  // For red flag, label (SALES/MARKETING) goes after the report type
  if (label) {
    parts.push(`${formatLabel.toUpperCase()} ${label}`);
  } else {
    parts.push(formatLabel.toUpperCase());
  }
  parts.push(dateRange);

  return `${parts.join(' - ')}.${ext}`;
}

// ── Main route ────────────────────────────────────────────────────────────────

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
  const channelAction   = (formData.get('channelAction') as string | null)?.trim() || '';

  if (!file || !brand || !reportId) {
    return NextResponse.json({ error: 'file, brand and reportId are required' }, { status: 400 });
  }

  let fileBuffer  = Buffer.from(await file.arrayBuffer());
  const rawFilename = file.name;
  const storeMap    = await loadStoreMap();

  const reports    = await loadReports();
  const reportDef  = reports.find(r => r.id === reportId);
  const reportName = reportDef?.name ?? reportId.replace(/-/g, ' ').toUpperCase();
  const dataFormat = reportDef?.dataFormat ?? '';
  const channel    = reportDef?.channel;

  const timestamp = new Date().toISOString();

  // ── Channel mismatch check (before stock-count analysis) ────────────────────
  if (channel && !channelAction) {
    const mismatch = analyzeChannelMismatch(fileBuffer, channel);
    if (mismatch) {
      return NextResponse.json({ channelMismatch: mismatch }, { status: 200 });
    }
  }

  // If user chose to exclude mismatched rows, filter before any further processing
  if (channel && channelAction === 'exclude') {
    fileBuffer = filterByChannel(fileBuffer, channel) as Buffer<ArrayBuffer>;
  }

  // ── Pre-generate analysis (skipped if user already confirmed) ───────────────
  if (!confirmed && dataFormat === 'stock-count') {
    const { warnings, hardError, unmatchedProducts } = await analyzeStockCount(fileBuffer);
    if (hardError) {
      return NextResponse.json({ error: hardError }, { status: 422 });
    }
    if (warnings.length > 0) {
      return NextResponse.json({ warnings, unmatchedProducts }, { status: 200 });
    }
  }

  try {
    const results: GenerateResult[] = [];
    void outputType; // reserved for future output-type switching

    switch (dataFormat) {

      case 'stock-count': {
        const retailer = channel || 'UNKNOWN';
        const currentMonth = new Date().toISOString().slice(0, 7);
        const historicalStockRows = await loadMonthUploads<DatabaseRow>('stock-count', brand, currentMonth);
        const { buffer, rawDates, weekLabel, archiveRows: stockArchiveRows } = await generateMakroStockCount(
          fileBuffer, brand, storeMap, retailer, historicalStockRows,
        );
        const filename = buildFilename(brand, dataFormat, channel, rawDates);
        results.push({ excelBuffer: buffer, filename, rawDates, weekLabel, retailer, label: '' });
        await archiveUpload('stock-count', brand, currentMonth, stockArchiveRows).catch(e => {
          console.error('[generate] Stock count archive failed:', e instanceof Error ? e.message : e);
        });
        break;
      }

      case 'red-flag': {
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

        const currentMonthRf = new Date().toISOString().slice(0, 7);
        const historicalRfRows = await loadMonthUploads<RedFlagArchiveRow>('red-flag', brand, currentMonthRf);

        // Generate Sales report
        if (salesProblems.length) {
          const { buffer, rawDates, archiveRows: salesArchiveRows } = await generateRedFlag(fileBuffer, brand, 'SALES', salesProblems, historicalRfRows);
          const filename = buildFilename(brand, dataFormat, channel, rawDates, 'SALES');
          results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: 'SALES' });
          await archiveUpload('red-flag', brand, currentMonthRf, salesArchiveRows).catch(e => {
            console.error('[generate] Red flag SALES archive failed:', e instanceof Error ? e.message : e);
          });
        }
        // Generate Marketing report
        if (marketingProblems.length) {
          const { buffer, rawDates, archiveRows: mktgArchiveRows } = await generateRedFlag(fileBuffer, brand, 'MARKETING', marketingProblems, historicalRfRows);
          const filename = buildFilename(brand, dataFormat, channel, rawDates, 'MARKETING');
          results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: 'MARKETING' });
          await archiveUpload('red-flag', brand, currentMonthRf, mktgArchiveRows).catch(e => {
            console.error('[generate] Red flag MARKETING archive failed:', e instanceof Error ? e.message : e);
          });
        }
        break;
      }

      case 'stand-report': {
        const currentMonthSr = new Date().toISOString().slice(0, 7);
        const historicalSrRows = await loadMonthUploads<StandReportArchiveRow>('stand-report', brand, currentMonthSr);
        const { buffer, rawDates, archiveRows: srArchiveRows } = await generateStandReport(fileBuffer, brand, historicalSrRows);
        const filename = buildFilename(brand, dataFormat, channel, rawDates);
        results.push({
          excelBuffer: buffer,
          filename,
          rawDates,
          weekLabel: '',
          retailer:  '',
          label:     '',
          contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });
        await archiveUpload('stand-report', brand, currentMonthSr, srArchiveRows).catch(e => {
          console.error('[generate] Stand report archive failed:', e instanceof Error ? e.message : e);
        });
        break;
      }

      case 'training-feedback': {
        // Determine current month for MTD history
        const currentMonth = new Date().toISOString().slice(0, 7); // YYYY-MM
        const historicalRows = await loadMonthUploads<TrainingSummaryRow>(
          'training-feedback', brand, currentMonth,
        );

        const { buffer, rawDates, archiveRows } = await generateTrainingFeedback(
          fileBuffer, brand, historicalRows,
        );
        const filename = buildFilename(brand, dataFormat, channel, rawDates);
        results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: '' });

        // Archive this upload's parsed rows for future MTD accumulation
        await archiveUpload('training-feedback', brand, currentMonth, archiveRows).catch(e => {
          console.error('[generate] Training archive failed:', e instanceof Error ? e.message : e);
        });
        break;
      }

      case 'activation-report': {
        const currentMonthAr = new Date().toISOString().slice(0, 7);
        const historicalArRows = await loadMonthUploads<ActivationArchiveRow>('activation-report', brand, currentMonthAr);
        const { buffer, rawDates, archiveRows: arArchiveRows } = await generateActivationReport(fileBuffer, brand, historicalArRows);
        const filename = buildFilename(brand, dataFormat, channel, rawDates);
        results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: '' });
        await archiveUpload('activation-report', brand, currentMonthAr, arArchiveRows).catch(e => {
          console.error('[generate] Activation report archive failed:', e instanceof Error ? e.message : e);
        });
        break;
      }

      case 'service-call': {
        const currentMonthSc = new Date().toISOString().slice(0, 7);
        const historicalScRows = await loadMonthUploads<ServiceCallArchiveRow>('service-call', brand, currentMonthSc);
        const { buffer, rawDates, archiveRows: scArchiveRows } = await generateServiceCallReport(fileBuffer, brand, historicalScRows);
        const filename = buildFilename(brand, dataFormat, channel, rawDates);
        results.push({ excelBuffer: buffer, filename, rawDates, weekLabel: '', retailer: '', label: '' });
        await archiveUpload('service-call', brand, currentMonthSc, scArchiveRows).catch(e => {
          console.error('[generate] Service call archive failed:', e instanceof Error ? e.message : e);
        });
        break;
      }

      default:
        return NextResponse.json(
          { error: `Report "${reportName}" has no data format configured. Please set a data format in the Admin Centre.` },
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

        const adminEmails = (await loadUsers())
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
      const entry: RunLogEntry = {
        id: randomUUID(),
        timestamp,
        userName,
        userEmail,
        reportId,
        reportName,
        brand,
        retailer:     channel || '',
        filename:     '',
        status:       'error',
        errorMessage: message,
        emailSent:    false,
      };
      await addRunEntry(entry);

      const adminEmails = (await loadUsers())
        .filter(u => u.isAdmin)
        .map(u => u.email);
      if (adminEmails.length) {
        await sendRunNotification(adminEmails, entry).catch(() => { /* non-fatal */ });
      }
    });

    return NextResponse.json({ error: message }, { status: 500 });
  }
}

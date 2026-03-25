import { Resend } from 'resend';

let _resend: Resend | null = null;
function getResend() {
  if (!_resend) _resend = new Resend(process.env.RESEND_API_KEY!);
  return _resend;
}

const FROM = 'Defy Field Execute <report_sender@outerjoin.co.za>';
const APP_URL = process.env.NEXT_PUBLIC_SITE_URL || 'https://defy-field-execute.vercel.app';

function emailShell(bodyContent: string) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #e5e5e5;">
      <!-- Header -->
      <table width="100%" cellpadding="0" cellspacing="0" style="background:#E31837;">
        <tr>
          <td style="padding:20px 28px;">
            <div style="color:#fff;font-size:20px;font-weight:bold;letter-spacing:1px;margin:0;">DEFY FIELD EXECUTE</div>
            <div style="color:#fff;margin:3px 0 0;opacity:0.85;font-size:12px;">Powered by Atomic Marketing &amp; Perigee</div>
          </td>
          <td style="padding:12px 28px 12px 0;text-align:right;vertical-align:middle;width:100px;">
            <img src="${APP_URL}/defy-logo.png" width="84" alt="Defy" style="background:#fff;padding:6px 8px;border-radius:4px;display:inline-block;" />
          </td>
        </tr>
      </table>

      <!-- Body -->
      <div style="padding:32px 28px;background:#fff;">
        ${bodyContent}

        <!-- Partner logos -->
        <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:24px;padding-top:16px;border-top:1px solid #eee;">
          <tr>
            <td style="vertical-align:middle;">
              <img src="${APP_URL}/atomic-logo.png" height="32" alt="Atomic Marketing" style="display:block;" />
            </td>
            <td style="text-align:right;vertical-align:middle;">
              <img src="${APP_URL}/perigee-logo.jpg" height="32" alt="Perigee" style="display:block;margin-left:auto;" />
            </td>
          </tr>
        </table>
      </div>

      <!-- Footer -->
      <div style="padding:14px 28px;text-align:center;font-size:11px;color:#999;background:#f9f9f9;border-top:1px solid #eee;">
        Defy Field Execute &bull; Powered by Atomic Marketing &amp; Perigee &bull; OuterJoin
      </div>
    </div>
  `;
}

export async function sendWelcomeEmail(to: string, name: string, password: string) {
  const body = `
    <p style="margin:0 0 14px;">Hi <strong>${name}</strong>,</p>
    <p style="margin:0 0 8px;">Your account has been created on <strong>Defy Field Execute</strong>.</p>
    <p style="margin:0 0 20px;color:#555;font-size:14px;">This is the portal used to turn raw Perigee exports into beautiful reports, save them to SharePoint and email them to you.</p>
    <table style="background:#f9f9f9;border:1px solid #eee;border-radius:6px;padding:14px 16px;width:100%;margin-bottom:20px;">
      <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">Login URL</td><td style="font-size:13px;"><a href="${APP_URL}/login" style="color:#E31837;">${APP_URL}/login</a></td></tr>
      <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">Email</td><td style="font-size:13px;">${to}</td></tr>
      <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">Password</td><td style="font-size:13px;font-family:monospace;">${password}</td></tr>
    </table>
    <p style="margin:0 0 20px;color:#666;font-size:13px;">Please change your password after your first login.</p>
    <a href="${APP_URL}/login" style="background:#E31837;color:#fff;text-decoration:none;padding:12px 24px;border-radius:4px;font-weight:bold;font-size:14px;display:inline-block;">Login Now</a>
  `;

  return getResend().emails.send({
    from: FROM,
    to,
    subject: 'Welcome to Defy Field Execute',
    html: emailShell(body),
  });
}

export async function sendRunNotification(
  adminEmails: string[],
  entry: {
    userName:   string;
    userEmail:  string;
    reportName: string;
    brand:      string;
    retailer:   string;
    filename:   string;
    timestamp:  string;
    status:     'success' | 'error';
    errorMessage?: string;
  },
) {
  if (!adminEmails.length) return;

  const ts  = new Date(entry.timestamp);
  const dateStr = ts.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  const timeStr = ts.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', hour12: false });

  const statusBadge = entry.status === 'success'
    ? `<span style="background:#d1fae5;color:#065f46;padding:2px 8px;border-radius:4px;font-size:12px;font-weight:bold;">SUCCESS</span>`
    : `<span style="background:#fee2e2;color:#991b1b;padding:2px 8px;border-radius:4px;font-size:12px;font-weight:bold;">ERROR</span>`;

  const body = `
    <p style="margin:0 0 16px;color:#333;">A report was generated on <strong>Defy Field Execute</strong>.</p>
    <table style="background:#f9f9f9;border:1px solid #eee;border-radius:6px;padding:14px 16px;width:100%;margin-bottom:20px;border-collapse:collapse;">
      <tr><td style="padding:5px 12px 5px 0;color:#666;font-size:13px;white-space:nowrap;">Status</td><td style="font-size:13px;">${statusBadge}</td></tr>
      <tr><td style="padding:5px 12px 5px 0;color:#666;font-size:13px;white-space:nowrap;">User</td><td style="font-size:13px;">${entry.userName} &lt;${entry.userEmail}&gt;</td></tr>
      <tr><td style="padding:5px 12px 5px 0;color:#666;font-size:13px;white-space:nowrap;">Date</td><td style="font-size:13px;">${dateStr}</td></tr>
      <tr><td style="padding:5px 12px 5px 0;color:#666;font-size:13px;white-space:nowrap;">Time</td><td style="font-size:13px;">${timeStr}</td></tr>
      <tr><td style="padding:5px 12px 5px 0;color:#666;font-size:13px;white-space:nowrap;">Report</td><td style="font-size:13px;">${entry.reportName}</td></tr>
      <tr><td style="padding:5px 12px 5px 0;color:#666;font-size:13px;white-space:nowrap;">Brand</td><td style="font-size:13px;">${entry.brand}</td></tr>
      <tr><td style="padding:5px 12px 5px 0;color:#666;font-size:13px;white-space:nowrap;">Retailer</td><td style="font-size:13px;">${entry.retailer}</td></tr>
      <tr><td style="padding:5px 12px 5px 0;color:#666;font-size:13px;white-space:nowrap;">File</td><td style="font-size:13px;font-family:monospace;word-break:break-all;">${entry.filename}</td></tr>
      ${entry.errorMessage ? `<tr><td style="padding:5px 12px 5px 0;color:#991b1b;font-size:13px;white-space:nowrap;">Error</td><td style="font-size:13px;color:#991b1b;">${entry.errorMessage}</td></tr>` : ''}
    </table>
    <p style="margin:0;color:#999;font-size:12px;">This is an automated notification. View the full run log in the Control Centre.</p>
  `;

  return getResend().emails.send({
    from: FROM,
    to: adminEmails,
    subject: `DFE: ${entry.status === 'success' ? '✓' : '✗'} ${entry.reportName} — ${entry.brand} · ${entry.retailer}`,
    html: emailShell(body),
  });
}

export async function sendReportEmail(params: {
  to:          string[];
  firstName:   string;
  reportName:  string;
  brand:       string;
  weekLabel:   string;
  filename:    string;
  fileBuffer:  Buffer;
  spFolderUrl: string;
}) {
  const body = `
    <p style="margin:0 0 20px;">Hi <strong>${params.firstName}</strong>,</p>
    <p style="margin:0 0 20px;">Please find <strong>${params.reportName}</strong> attached for <strong>${params.brand}</strong> for <strong>${params.weekLabel}</strong>.</p>
    <p style="margin:0 0 8px;color:#333;">Remember that the report automatically saves in SharePoint, here:</p>
    <p style="margin:0 0 24px;">
      <a href="${params.spFolderUrl}" style="color:#E31837;word-break:break-all;">${params.spFolderUrl}</a>
    </p>
    <p style="margin:0 0 24px;color:#666;font-size:13px;">If you don't have access to the folder, reach out to your CAM at OuterJoin.</p>
    <p style="margin:0;color:#333;">Thank you<br>Team OJ</p>
  `;

  return getResend().emails.send({
    from: FROM,
    to: params.to,
    subject: `Defy Field Execute — ${params.reportName} · ${params.brand} · ${params.weekLabel}`,
    html: emailShell(body),
    attachments: [{ filename: params.filename, content: params.fileBuffer.toString('base64') }],
  });
}

export async function sendPasswordResetEmail(to: string, name: string, password: string) {
  const body = `
    <p style="margin:0 0 14px;">Hi <strong>${name}</strong>,</p>
    <p style="margin:0 0 20px;">Your password has been reset. Use the credentials below to log in.</p>
    <table style="background:#f9f9f9;border:1px solid #eee;border-radius:6px;padding:14px 16px;width:100%;margin-bottom:20px;">
      <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">Email</td><td style="font-size:13px;">${to}</td></tr>
      <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">New Password</td><td style="font-size:13px;font-family:monospace;">${password}</td></tr>
    </table>
    <a href="${APP_URL}/login" style="background:#E31837;color:#fff;text-decoration:none;padding:12px 24px;border-radius:4px;font-weight:bold;font-size:14px;display:inline-block;">Login Now</a>
  `;

  return getResend().emails.send({
    from: FROM,
    to,
    subject: 'Defy Field Execute — Password Reset',
    html: emailShell(body),
  });
}

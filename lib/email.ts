import { Resend } from 'resend';

let _resend: Resend | null = null;
function getResend() {
  if (!_resend) _resend = new Resend(process.env.RESEND_API_KEY!);
  return _resend;
}

const FROM = 'Defy Field Execute <noreply@outerjoin.co.za>';
const APP_URL = process.env.NEXT_PUBLIC_SITE_URL || 'https://defy-field-execute.vercel.app';

export async function sendWelcomeEmail(to: string, name: string, password: string) {
  return getResend().emails.send({
    from: FROM,
    to,
    subject: 'Welcome to Defy Field Execute',
    html: `
      <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
        <div style="background:#E31837;padding:24px 32px;">
          <h1 style="color:#fff;margin:0;font-size:22px;letter-spacing:1px;">DEFY FIELD EXECUTE</h1>
          <p style="color:#fff;margin:4px 0 0;opacity:0.85;font-size:13px;">Powered by Atomic Marketing &amp; Perigee</p>
        </div>
        <div style="padding:32px;background:#fff;border:1px solid #eee;">
          <p style="margin:0 0 16px;">Hi <strong>${name}</strong>,</p>
          <p style="margin:0 0 16px;">Your account has been created on <strong>Defy Field Execute</strong>.</p>
          <table style="background:#f9f9f9;border:1px solid #eee;border-radius:6px;padding:16px;width:100%;margin-bottom:24px;">
            <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">Login URL</td><td style="font-size:13px;"><a href="${APP_URL}/login" style="color:#E31837;">${APP_URL}/login</a></td></tr>
            <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">Email</td><td style="font-size:13px;">${to}</td></tr>
            <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">Password</td><td style="font-size:13px;font-family:monospace;">${password}</td></tr>
          </table>
          <p style="margin:0 0 16px;color:#666;font-size:13px;">Please change your password after your first login.</p>
          <a href="${APP_URL}/login" style="background:#E31837;color:#fff;text-decoration:none;padding:12px 24px;border-radius:4px;font-weight:bold;font-size:14px;">Login Now</a>
        </div>
        <div style="padding:16px 32px;text-align:center;font-size:11px;color:#999;">
          Defy Field Execute &bull; Powered by Atomic Marketing &amp; Perigee &bull; OuterJoin
        </div>
      </div>
    `,
  });
}

export async function sendPasswordResetEmail(to: string, name: string, password: string) {
  return getResend().emails.send({
    from: FROM,
    to,
    subject: 'Defy Field Execute — Password Reset',
    html: `
      <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
        <div style="background:#E31837;padding:24px 32px;">
          <h1 style="color:#fff;margin:0;font-size:22px;letter-spacing:1px;">DEFY FIELD EXECUTE</h1>
        </div>
        <div style="padding:32px;background:#fff;border:1px solid #eee;">
          <p style="margin:0 0 16px;">Hi <strong>${name}</strong>,</p>
          <p style="margin:0 0 16px;">Your password has been reset. Use the credentials below to log in.</p>
          <table style="background:#f9f9f9;border:1px solid #eee;border-radius:6px;padding:16px;width:100%;margin-bottom:24px;">
            <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">Email</td><td style="font-size:13px;">${to}</td></tr>
            <tr><td style="padding:4px 12px 4px 0;color:#666;font-size:13px;">New Password</td><td style="font-size:13px;font-family:monospace;">${password}</td></tr>
          </table>
          <a href="${APP_URL}/login" style="background:#E31837;color:#fff;text-decoration:none;padding:12px 24px;border-radius:4px;font-weight:bold;font-size:14px;">Login Now</a>
        </div>
      </div>
    `,
  });
}

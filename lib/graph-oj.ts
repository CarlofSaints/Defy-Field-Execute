// Microsoft Graph API — OJ tenant (exceler8xl.sharepoint.com)

const TENANT_ID = process.env.OJ_TENANT_ID!;
const CLIENT_ID = process.env.OJ_CLIENT_ID!;
const CLIENT_SECRET = process.env.OJ_CLIENT_SECRET!;
const SP_HOST = (process.env.OJ_SP_HOST || 'exceler8xl.sharepoint.com').trim();
const SP_LIBRARY = (process.env.OJ_SP_LIBRARY || 'Shared Documents').trim();

let _token: string | null = null;
let _tokenExpiry = 0;

async function getToken(): Promise<string> {
  if (_token && Date.now() < _tokenExpiry - 60_000) return _token;

  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
      }),
    }
  );
  if (!res.ok) throw new Error(`Token error: ${await res.text()}`);
  const data = await res.json();
  _token = data.access_token;
  _tokenExpiry = Date.now() + data.expires_in * 1000;
  return _token!;
}

async function getSiteId(): Promise<string> {
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SP_HOST}:/sites/root`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`getSiteId error: ${await res.text()}`);
  const data = await res.json();
  return data.id;
}

async function getDriveId(siteId: string): Promise<string> {
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`getDrives error: ${await res.text()}`);
  const data = await res.json();
  const drive = (data.value as { name: string; id: string }[]).find(
    d => d.name === SP_LIBRARY
  );
  if (!drive) throw new Error(`Drive "${SP_LIBRARY}" not found`);
  return drive.id;
}

/**
 * Upload a file buffer to SharePoint.
 * folderPath example: "Clients/DEFY/REPORTS/2026-03"
 */
export async function uploadToSharePoint(
  folderPath: string,
  filename: string,
  buffer: Buffer,
  contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
): Promise<string> {
  const token = await getToken();
  const siteId = await getSiteId();
  const driveId = await getDriveId(siteId);

  const encodedPath = folderPath
    .split('/')
    .map(p => encodeURIComponent(p))
    .join('/');
  const encodedFile = encodeURIComponent(filename);

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedPath}/${encodedFile}:/content`;

  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': contentType,
    },
    body: buffer as unknown as BodyInit,
  });
  if (!res.ok) throw new Error(`SP upload error: ${await res.text()}`);
  const data = await res.json();
  return data.webUrl as string;
}

/**
 * Send an email via Microsoft Graph (OJ mailbox).
 */
export async function sendGraphEmail(params: {
  from: string;
  to: string[];
  subject: string;
  html: string;
}) {
  const token = await getToken();
  const fromClean = params.from.trim();

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromClean)}/sendMail`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        message: {
          subject: params.subject,
          body: { contentType: 'HTML', content: params.html },
          toRecipients: params.to.map(addr => ({
            emailAddress: { address: addr.trim() },
          })),
        },
        saveToSentItems: true,
      }),
    }
  );
  if (!res.ok) throw new Error(`sendMail error: ${await res.text()}`);
}

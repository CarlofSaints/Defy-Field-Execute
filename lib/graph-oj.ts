// Microsoft Graph API — OJ tenant (exceler8xl.sharepoint.com)

const TENANT_ID = process.env.OJ_TENANT_ID!;
const CLIENT_ID = process.env.OJ_CLIENT_ID!;
const CLIENT_SECRET = process.env.OJ_CLIENT_SECRET!;
const SP_HOST = (process.env.OJ_SP_HOST || 'exceler8xl.sharepoint.com').trim();
const SP_LIBRARY = (process.env.OJ_SP_LIBRARY || 'Shared Documents').trim();

let _token: string | null = null;
let _tokenExpiry = 0;
let _siteId: string | null = null;
let _driveId: string | null = null;

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
  if (_siteId) return _siteId;
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SP_HOST}:/`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`getSiteId error: ${await res.text()}`);
  const data = await res.json();
  _siteId = data.id as string;
  return _siteId;
}

async function getDriveId(siteId: string): Promise<string> {
  if (_driveId) return _driveId;
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
  _driveId = drive.id;
  return _driveId;
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
 * Download a file from SharePoint. Returns null if the file is not found or
 * on any error (so callers can gracefully fall back).
 */
export async function downloadFileFromSP(
  folderPath: string,
  filename:   string,
): Promise<Buffer | null> {
  try {
    const token   = await getToken();
    const siteId  = await getSiteId();
    const driveId = await getDriveId(siteId);

    const encodedPath = folderPath.split('/').map(p => encodeURIComponent(p)).join('/');
    const encodedFile = encodeURIComponent(filename);
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedPath}/${encodedFile}:/content`;

    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) return null;
    return Buffer.from(await res.arrayBuffer());
  } catch {
    return null;
  }
}

/**
 * Recursively list all files inside a SharePoint folder (including subfolders).
 * Uses the /children endpoint (not search) so results are always current — no
 * indexing delay for recently added or moved files.
 * Returns a Map of lowercase filename → driveItem ID.
 * Falls back to an empty Map on any error.
 */
export async function listFilesInSPFolder(
  folderPath: string,
): Promise<Map<string, string>> {
  try {
    const token   = await getToken();
    const siteId  = await getSiteId();
    const driveId = await getDriveId(siteId);
    const fileMap = new Map<string, string>();

    async function listChildren(path: string): Promise<void> {
      const encoded = path.split('/').map(p => encodeURIComponent(p)).join('/');
      let nextUrl: string | null =
        `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encoded}:/children?$select=id,name,file,folder&$top=1000`;

      while (nextUrl) {
        const res = await fetch(nextUrl, { headers: { Authorization: `Bearer ${token}` } });
        if (!res.ok) {
          console.error(`[listFilesInSPFolder] children failed (${path}): ${res.status} ${await res.text()}`);
          break;
        }
        const data = await res.json() as {
          value?: { id: string; name: string; file?: object; folder?: object }[];
          '@odata.nextLink'?: string;
        };
        const subfolders: string[] = [];
        for (const item of data.value ?? []) {
          if (item.file && item.id) {
            fileMap.set(item.name.toLowerCase(), item.id);
          } else if (item.folder) {
            subfolders.push(`${path}/${item.name}`);
          }
        }
        // Recurse into subfolders in parallel
        await Promise.all(subfolders.map(listChildren));
        nextUrl = data['@odata.nextLink'] ?? null;
      }
    }

    await listChildren(folderPath);
    return fileMap;
  } catch (err) {
    console.error(`[listFilesInSPFolder] error:`, err instanceof Error ? err.message : err);
    return new Map();
  }
}

/**
 * Download a SharePoint file by its driveItem ID using authenticated Graph API.
 * More reliable than pre-auth URLs which aren't always returned for SP files.
 */
export async function downloadSPFileById(itemId: string): Promise<Buffer | null> {
  try {
    const token   = await getToken();
    const siteId  = await getSiteId();
    const driveId = await getDriveId(siteId);
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      console.error(`[downloadSPFileById] ${itemId}: ${res.status}`);
      return null;
    }
    return Buffer.from(await res.arrayBuffer());
  } catch {
    return null;
  }
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

/**
 * Generic blob-based archive for persisting parsed upload data across sessions.
 *
 * Blob key pattern: uploads/{reportType}/{brand}/{YYYY-MM}/{uploadId}.json
 *
 * On Vercel (BLOB_READ_WRITE_TOKEN set): uses @vercel/blob for persistence.
 * Locally (no token): archiving is a no-op and loadMonthUploads returns [].
 */
import { randomUUID } from 'crypto';
import { blobEnabled } from './blobStore';

const useBlob = blobEnabled;

function blobKey(reportType: string, brand: string, month: string, uploadId: string): string {
  return `uploads/${reportType}/${brand.toLowerCase()}/${month}/${uploadId}.json`;
}

function prefixKey(reportType: string, brand: string, month: string): string {
  return `uploads/${reportType}/${brand.toLowerCase()}/${month}/`;
}

/** Archive parsed upload data. Returns the uploadId. */
export async function archiveUpload<T>(
  reportType: string,
  brand: string,
  month: string, // YYYY-MM
  data: T[],
): Promise<string | null> {
  if (!useBlob) return null;

  const { put } = await import('@vercel/blob');
  const uploadId = randomUUID();
  const key = blobKey(reportType, brand, month, uploadId);
  const payload = JSON.stringify({
    id: uploadId,
    timestamp: new Date().toISOString(),
    rowCount: data.length,
    rows: data,
  });

  await put(key, payload, {
    access: 'public',
    contentType: 'application/json',
    addRandomSuffix: false,
  });

  return uploadId;
}

/** Load all archived uploads for a given month, returning merged rows. */
export async function loadMonthUploads<T>(
  reportType: string,
  brand: string,
  month: string, // YYYY-MM
): Promise<T[]> {
  if (!useBlob) return [];

  const { list } = await import('@vercel/blob');
  const prefix = prefixKey(reportType, brand, month);
  const listing = await list({ prefix });

  if (!listing.blobs.length) return [];

  const allRows: T[] = [];
  for (const blob of listing.blobs) {
    try {
      const res = await fetch(blob.url, { cache: 'no-store' });
      if (!res.ok) continue;
      const json = await res.json() as { rows: T[] };
      if (Array.isArray(json.rows)) {
        allRows.push(...json.rows);
      }
    } catch {
      // Skip corrupt/unreadable blobs
    }
  }

  return allRows;
}

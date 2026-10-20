import { downloadSPFileById } from '@/lib/graph-oj';

/**
 * Photo key for an export link.
 * Perigee: the "perigee-XXXX" segment, which is the file name the VBA saved to SharePoint.
 * Field Goose (since Oct 2026): the link itself. Its links are public, so we download direct.
 */
export function urlToImageId(url: string): string {
  if (!url) return '';
  if (url.startsWith('https://www.fieldgoose.com/')) return url;
  const seg = url.split('/').find(s => s.toLowerCase().startsWith('perigee-'));
  return seg ?? '';
}

/** Field Goose links straight from the web, Perigee ids from the VBA-saved SharePoint file. */
export async function fetchPhoto(id: string, spFiles: Map<string, string>): Promise<Buffer | null> {
  if (id.startsWith('https://')) {
    try {
      const res = await fetch(id);
      return res.ok ? Buffer.from(await res.arrayBuffer()) : null;
    } catch {
      return null;
    }
  }
  const itemId = spFiles.get(`${id}.jpg`.toLowerCase());
  return itemId ? downloadSPFileById(itemId) : null;
}

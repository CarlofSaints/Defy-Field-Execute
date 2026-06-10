import fs from 'fs';
import path from 'path';

// ─── Product catalog (control file) ─────────────────────────────────────────
// Maps a product CODE (e.g. "DDW242") to its CATEGORY + SUB CATEGORY.
// The code is the CLIENT PRODUCT ID column of the Product Management file; it
// also appears as the leading token of each product description in the Perigee
// stock-count export. Used to enrich CATEGORY / SUB CAT on reports.
//
// One combined catalog covers every brand — the control file carries Defy,
// Beko, Grundig, etc. in a single sheet and product codes are globally unique,
// so there is no per-brand split.
//
// Persistence: Vercel Blob on the server (the same store the upload archive
// uses), local JSON file in dev. The earlier env-var approach was dropped — it
// silently failed without VERCEL_TOKEN and suffered the baked-at-deploy
// stale-read bug, which is why the catalog "didn't persist".

export interface ProductCatalog {
  count:    number;                  // number of products in the map
  products: Record<string, string>;  // normalisedId -> "CATEGORY|SUBCATEGORY"
}

export interface ResolvedProduct {
  productCode: string;
  category:    string;
  subCategory: string;
  matched:     boolean;
}

export interface ProductLookup {
  size: number;
  resolve(description: string): ResolvedProduct;
}

const FILE     = path.join(process.cwd(), 'data', 'productCatalog.json');
const BLOB_KEY = 'config/product-catalog.json';
const useBlob  = !!process.env.BLOB_READ_WRITE_TOKEN;

/** Uppercase + strip everything that isn't a letter or digit. */
export function normalizeCode(s: string): string {
  return String(s ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

export async function loadProductCatalog(): Promise<ProductCatalog | null> {
  if (useBlob) {
    const { list } = await import('@vercel/blob');
    const listing = await list({ prefix: BLOB_KEY });
    const match = listing.blobs.find(b => b.pathname === BLOB_KEY) ?? listing.blobs[0];
    if (!match) return null;
    const res = await fetch(match.url, { cache: 'no-store' });
    if (!res.ok) return null;
    return await res.json() as ProductCatalog;
  }

  if (fs.existsSync(FILE)) {
    const parsed = JSON.parse(fs.readFileSync(FILE, 'utf-8'));
    // Tolerate the old array-of-per-brand-catalogs shape from earlier seeds.
    if (Array.isArray(parsed)) {
      const products: Record<string, string> = {};
      for (const c of parsed) Object.assign(products, c.products ?? {});
      return { count: Object.keys(products).length, products };
    }
    return parsed as ProductCatalog;
  }

  return null;
}

export async function saveProductCatalog(catalog: ProductCatalog): Promise<void> {
  if (useBlob) {
    const { put } = await import('@vercel/blob');
    await put(BLOB_KEY, JSON.stringify(catalog), {
      access:          'public',
      contentType:     'application/json',
      addRandomSuffix: false,
    });
    return;
  }

  const dir = path.dirname(FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(FILE, JSON.stringify(catalog, null, 2));
}

export async function deleteProductCatalog(): Promise<void> {
  if (useBlob) {
    const { list, del } = await import('@vercel/blob');
    const listing = await list({ prefix: BLOB_KEY });
    const match = listing.blobs.find(b => b.pathname === BLOB_KEY) ?? listing.blobs[0];
    if (match) await del(match.url);
    return;
  }

  if (fs.existsSync(FILE)) fs.rmSync(FILE);
}

/**
 * Build a category lookup. The returned `resolve()` matches a Perigee product
 * description to its catalog entry by finding the longest product code that
 * prefixes the (normalised) description — this tolerates codes that contain
 * spaces in the form data (e.g. "DDW 242" → "DDW242"). Falls back to a
 * substring match (code at the end of the description), then to the first token
 * with an UNKNOWN category if nothing matches.
 */
export async function buildProductLookup(): Promise<ProductLookup> {
  const catalog  = await loadProductCatalog();
  const products = catalog?.products ?? {};
  // Longest code first so the most specific match wins (e.g. DAC4470 before DAC447).
  const keys  = Object.keys(products).sort((a, b) => b.length - a.length);
  const cache = new Map<string, ResolvedProduct>();

  const make = (key: string): ResolvedProduct => {
    const [category, subCategory] = (products[key] ?? '|').split('|');
    return {
      productCode: key,
      category:    category || 'UNKNOWN',
      subCategory: subCategory || 'UNKNOWN',
      matched:     true,
    };
  };

  const resolve = (description: string): ResolvedProduct => {
    const cached = cache.get(description);
    if (cached) return cached;

    const nd = normalizeCode(description);
    let result: ResolvedProduct | null = null;

    // 1. Longest code that prefixes the normalised description.
    for (const k of keys) {
      if (k.length >= 3 && nd.startsWith(k)) { result = make(k); break; }
    }
    // 2. Fallback: longest code that appears anywhere (catches codes at the end).
    if (!result) {
      for (const k of keys) {
        if (k.length >= 5 && nd.includes(k)) { result = make(k); break; }
      }
    }
    // 3. No match: keep the leading token as the code, mark UNKNOWN category.
    if (!result) {
      result = {
        productCode: normalizeCode(description.split(/\s+/)[0] || ''),
        category:    'UNKNOWN',
        subCategory: 'UNKNOWN',
        matched:     false,
      };
    }

    cache.set(description, result);
    return result;
  };

  return { size: keys.length, resolve };
}

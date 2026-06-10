import fs from 'fs';
import path from 'path';

// ─── Product catalog (control file) ─────────────────────────────────────────
// Maps a product CODE (e.g. "DDW242") to its CATEGORY + SUB CATEGORY.
// The code is the CLIENT PRODUCT ID column of the DEFY Product Management file;
// it also appears as the leading token of each product description in the
// Perigee stock-count export. Used to enrich CATEGORY / SUB CAT on reports.
//
// Stored compactly as a map  normalisedId -> "CATEGORY|SUBCATEGORY"  to keep the
// Vercel env-var payload small (~33 KB for ~1100 products vs ~72 KB as an array).

export interface ProductCatalog {
  brand:    string;                  // e.g. "DEFY" or "BEKO"
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

const FILE       = path.join(process.cwd(), 'data', 'productCatalog.json');
const VERCEL_KEY = 'DFE_PRODUCT_CATALOG_JSON';
const PROJECT_ID = 'prj_FaBoeZxXminOA9W8gSwsrwuLTz2i';

let _cache: ProductCatalog[] | null = null;

/** Uppercase + strip everything that isn't a letter or digit. */
export function normalizeCode(s: string): string {
  return String(s ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

export function loadProductCatalogs(): ProductCatalog[] {
  if (_cache !== null) return _cache;

  const env = process.env[VERCEL_KEY];
  if (process.env.VERCEL && env) {
    _cache = JSON.parse(env);
    return _cache!;
  }

  if (fs.existsSync(FILE)) {
    _cache = JSON.parse(fs.readFileSync(FILE, 'utf-8'));
    return _cache!;
  }

  if (env) {
    _cache = JSON.parse(env);
    return _cache!;
  }

  return [];
}

/**
 * Build a category lookup for a brand. The returned `resolve()` matches a
 * Perigee product description to its catalog entry by finding the longest
 * product code that prefixes the (normalised) description — this tolerates
 * codes that contain spaces in the form data (e.g. "DDW 242" → "DDW242").
 * Falls back to a substring match (code at the end of the description), then
 * to the first token with an UNKNOWN category if nothing matches.
 */
export function buildProductLookup(brand: string): ProductLookup {
  const all = loadProductCatalogs();
  const cfg = all.find(c => c.brand.toUpperCase() === brand.toUpperCase());
  const products = cfg?.products ?? {};
  // Longest code first so the most specific match wins (e.g. DAC4470 before DAC447).
  const keys = Object.keys(products).sort((a, b) => b.length - a.length);
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

export async function saveProductCatalogs(catalogs: ProductCatalog[]) {
  _cache = catalogs;

  try {
    const dir = path.dirname(FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(FILE, JSON.stringify(catalogs, null, 2));
    return;
  } catch {
    // Vercel read-only FS — fall through to env-var persistence
  }

  const token = process.env.VERCEL_TOKEN;
  if (!token) return;

  const listRes = await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!listRes.ok) return;

  const { envs } = await listRes.json() as { envs: { id: string; key: string }[] };
  const existing = envs.find(e => e.key === VERCEL_KEY);
  const value    = JSON.stringify(catalogs);

  if (!existing) {
    await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        key: VERCEL_KEY,
        value,
        type: 'plain',
        target: ['production', 'preview', 'development'],
      }),
    });
  } else {
    await fetch(`https://api.vercel.com/v9/projects/${PROJECT_ID}/env/${existing.id}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ value }),
    });
  }
}

import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import {
  loadProductCatalogs,
  saveProductCatalogs,
  normalizeCode,
  type ProductCatalog,
} from '@/lib/productCatalogData';

export async function GET() {
  // Return summary only (brand + count) — the full map can be large.
  const catalogs = loadProductCatalogs();
  return NextResponse.json(catalogs.map(c => ({ brand: c.brand, count: c.count })));
}

export async function POST(req: NextRequest) {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json({ error: 'Invalid form data' }, { status: 400 });
  }

  const file  = formData.get('file')  as File   | null;
  const brand = formData.get('brand') as string | null;

  if (!file)  return NextResponse.json({ error: 'No file provided' }, { status: 400 });
  if (!brand) return NextResponse.json({ error: 'Brand is required' }, { status: 400 });

  const buffer = Buffer.from(await file.arrayBuffer());
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as (string | null)[][];

  if (rows.length < 2) {
    return NextResponse.json({ error: 'File has no data rows' }, { status: 400 });
  }

  const headers = (rows[0] as (string | null)[]).map(h => (h ?? '').trim().toUpperCase());

  const idIdx = headers.findIndex(h =>
    ['CLIENT PRODUCT ID', 'PRODUCT ID', 'PRODUCT CODE', 'MODEL NUMBER', 'MODEL', 'CODE', 'ID'].includes(h),
  );
  const catIdx = headers.findIndex(h =>
    ['PRODUCT CATEGORY', 'CATEGORY', 'CAT'].includes(h),
  );
  const subIdx = headers.findIndex(h =>
    ['PRODUCT SUB CATEGORY', 'SUB CATEGORY', 'SUBCATEGORY', 'SUB CAT', 'SUB-CATEGORY'].includes(h),
  );

  if (idIdx === -1) {
    return NextResponse.json(
      { error: 'Could not find a product code column. Expected header: CLIENT PRODUCT ID, PRODUCT CODE or MODEL NUMBER.' },
      { status: 400 },
    );
  }
  if (catIdx === -1 && subIdx === -1) {
    return NextResponse.json(
      { error: 'Could not find a PRODUCT CATEGORY or PRODUCT SUB CATEGORY column.' },
      { status: 400 },
    );
  }

  const products: Record<string, string> = {};
  for (const row of rows.slice(1)) {
    const normId = normalizeCode(String(row[idIdx] ?? ''));
    if (!normId) continue;
    const category    = catIdx >= 0 ? String(row[catIdx] ?? '').trim().toUpperCase() : '';
    const subCategory = subIdx >= 0 ? String(row[subIdx] ?? '').trim().toUpperCase() : '';
    if (!category && !subCategory) continue;
    // First occurrence wins (control files are ordered with the canonical row first)
    if (!(normId in products)) products[normId] = `${category}|${subCategory}`;
  }

  const count = Object.keys(products).length;
  if (count === 0) {
    return NextResponse.json({ error: 'No valid products found in file' }, { status: 400 });
  }

  // Upsert by brand
  const catalogs = loadProductCatalogs();
  const b = brand.toUpperCase();
  const idx = catalogs.findIndex(c => c.brand.toUpperCase() === b);
  const newCatalog: ProductCatalog = { brand: b, count, products };
  if (idx >= 0) catalogs[idx] = newCatalog;
  else catalogs.push(newCatalog);

  await saveProductCatalogs(catalogs);
  return NextResponse.json({ count, brand: b });
}

export async function DELETE(req: NextRequest) {
  const { brand } = await req.json() as { brand?: string };
  if (!brand) {
    return NextResponse.json({ error: 'brand is required' }, { status: 400 });
  }

  const catalogs = loadProductCatalogs();
  const b = brand.toUpperCase();
  const filtered = catalogs.filter(c => c.brand.toUpperCase() !== b);

  if (filtered.length === catalogs.length) {
    return NextResponse.json({ error: 'Catalog not found' }, { status: 404 });
  }

  await saveProductCatalogs(filtered);
  return NextResponse.json({ ok: true });
}

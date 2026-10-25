// Self-check: a Field Goose export (ISO dates) gets the right week and SharePoint folder.
// Run: npx tsx scripts/check-fg-dates.ts <fg-export.xlsx>
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { normaliseUploadDates } from '../lib/reports/normalise-dates';
import { buildDfeFolderPath } from '../lib/sharepoint-dfe';

const buf = normaliseUploadDates(readFileSync(process.argv[2]));
const rows = XLSX.utils.sheet_to_json(XLSX.read(buf).Sheets[XLSX.read(buf).SheetNames[0]], { header: 1 }) as unknown[][];
const dates = rows.slice(1).map(r => String(r[9]));
console.log('dates:', [...new Set(dates)]);
assert(dates.every(d => /^\d{2}\/\d{2}\/\d{4}$/.test(d)), 'dates not DD/MM/YYYY');
console.log('SP folder:', buildDfeFolderPath('DEFY', dates));

// Perigee files pass through byte-for-byte
const perigee = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(perigee, XLSX.utils.aoa_to_sheet([['Date'], ['18/10/2026']]), 'S');
const pBuf = Buffer.from(XLSX.write(perigee, { type: 'buffer', bookType: 'xlsx' }));
assert(normaliseUploadDates(pBuf) === pBuf, 'Perigee file was rewritten');
console.log('OK');

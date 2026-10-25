// Self-check: base columns are found by name, also when an export inserts a column.
// Run: npx tsx scripts/check-base-cols.ts <export.xlsx>
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { baseCols } from '../lib/reports/base-cols';

const wb = XLSX.read(readFileSync(process.argv[2]));
const headers = (XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }) as unknown[][])[0];
assert.deepEqual(baseCols(headers), { first: 2, last: 3, store: 6, code: 7, province: 8, date: 9 });

const shifted = [...headers.slice(0, 5), 'Region', ...headers.slice(5)];
assert.deepEqual(baseCols(shifted), { first: 2, last: 3, store: 7, code: 8, province: 9, date: 10 });
console.log('OK');

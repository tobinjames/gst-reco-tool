// xlsxbook: cells stay within Excel's limits (a longer cell makes Excel "repair" the file)
// Run: npx tsx test/xlsxbook.test.ts (or npm test for all)
import assert from 'assert/strict';
import JSZip from 'jszip';
import { simpleXlsx, writeBook } from '../src/xlsxbook';
import { readTable } from '../src/register';

const long = Array.from({ length: 6000 }, (_, i) => `BL${String(i).padStart(5, '0')}`).join(', '); // ~54,000 characters
const bytes = await writeBook([{ name: 'S', rows: [[long, { v: 'bad ￾ char \uD800 lone' }, { f: 'A1', v: long }]] }]);
const xml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string');
for (const m of xml.matchAll(/<(t|v)[^>]*>([^<]*)<\/\1>/g))
  assert.ok(m[2].length <= 32767, `cell text of ${m[2].length} characters`);
assert.ok(!/￾|\uD800/.test(xml), 'characters XML does not allow are dropped');

// one-sheet download used by the Review step: title lines, blank line, header, data; numbers stay numbers
const dl = await simpleXlsx(
  'Invoices',
  [
    ['Supplier', 'Taxable'],
    ['A & B <Traders>', 1000.5],
    ['Total', 1000.5],
  ],
  ['Purchases vs GSTR-2B', 'FY 2026-27'],
);
const back = (await readTable('dl.xlsx', dl))[0];
assert.equal(back.name, 'Invoices');
assert.deepEqual(
  back.rows.filter(r => r.some(v => v !== null && v !== '')).map(r => r.filter(v => v !== null && v !== '')),
  [['Purchases vs GSTR-2B'], ['FY 2026-27'], ['Supplier', 'Taxable'], ['A & B <Traders>', 1000.5], ['Total', 1000.5]],
);
const wbXml = await (await JSZip.loadAsync(dl)).file('xl/workbook.xml')!.async('string');
assert.ok(wbXml.includes("'Invoices'!$A$4:$B$6"), 'autofilter on the header row and data');
console.log('xlsxbook.ts: all checks passed');

// Tally columnar register: multi-column mapping (taxable = sum of ledger columns, each tax head = sum of tax-ledger columns)
// Run: npx tsx test/tally.test.ts (or npm test for all)   (fixture: test/fixtures/tally_columnar_purchase.xlsx, synthetic; regenerate with make_tally_columnar.py)
import { readFileSync } from 'fs';
import assert from 'assert/strict';
import {
  PURCHASE_FIELDS,
  colsOf,
  detectHeader,
  guessMappingFull,
  looksColumnar,
  normalise,
  readTable,
  unusedAmountCols,
} from '../src/register';

const sheets = await readTable(
  'tally_columnar_purchase.xlsx',
  new Uint8Array(readFileSync(new URL('./fixtures/tally_columnar_purchase.xlsx', import.meta.url))),
);
const rows = sheets[0].rows;
const h = detectHeader(rows, PURCHASE_FIELDS);
const hdr = rows[h];
const name = (cs: number[]) => cs.map(j => hdr[j]);
assert.equal(h, 6, 'header row');
assert.ok(looksColumnar(hdr), 'recognised as columnar');
const map = guessMappingFull(rows, h, PURCHASE_FIELDS);
console.log(Object.fromEntries(Object.entries(map).map(([k, v]) => [k, name(colsOf(v))])));
assert.deepEqual(name(colsOf(map.no)), ['Supplier Invoice No.']);
assert.deepEqual(name(colsOf(map.date)), ['Supplier Invoice Date']);
assert.deepEqual(name(colsOf(map.gstin)), ['GSTIN/UIN']);
assert.deepEqual(name(colsOf(map.party)), ['Particulars']);
assert.deepEqual(name(colsOf(map.gross)), ['Gross Total']);
assert.deepEqual(name(colsOf(map.taxable)), [
  'Purchase @ 18% (Local)',
  'Purchase @ 12% (Local)',
  'Purchase @ 18% (Interstate)',
  'Freight Inward',
  'Packing Charges',
  'Professional Charges',
]);
assert.deepEqual(name(colsOf(map.cgst)), ['Input CGST @ 9%', 'Input CGST @ 6%']);
assert.deepEqual(name(colsOf(map.sgst)), ['Input SGST @ 9%', 'Input SGST @ 6%']);
assert.deepEqual(name(colsOf(map.igst)), ['Input IGST @ 18%']);
assert.deepEqual(name(colsOf(map.cess)), ['Input Cess']);
assert.deepEqual(name(unusedAmountCols(rows, h, map)), [], 'round off / TDS are not flagged as missed ledgers');
assert.deepEqual(
  name(unusedAmountCols(rows, h, { ...map, taxable: colsOf(map.taxable).filter(j => hdr[j] !== 'Packing Charges') })),
  ['Packing Charges'],
);

const d = { noGstinZeroTax: 'EXEMPT', itcType: 'Inputs', status: 'Availed', claimFy: '2026-27', purNoGstin: 'RCM-UNREG' };
const n = normalise('purchases', rows, h, map, d, '2026-27');
console.log(n.stats);
n.rows.forEach(r => console.log(r.slice(0, 5).join(' | '), '|', r.slice(11, 17).join(' ')));
assert.equal(n.stats.lines, 6);
assert.equal(n.stats.docs, 6);
assert.equal(n.stats.badDates, 0);
assert.equal(n.stats.outsideFy, 0);
assert.equal(n.stats.grossOff, 0, 'gross total agrees (round off and TDS allowed for)');
const [r1, r2] = n.rows;
assert.deepEqual(r1.slice(1, 5), ['SST/26-27/1042', 'Sharma Steel Traders', '27AABCS1234F1Z5', 'B2B']);
assert.deepEqual(r1.slice(12, 17), [102500, 9225, 9225, 0, 0]);
assert.equal(r1[11], 18);
assert.deepEqual(r2.slice(12, 17), [51000, 3390, 3390, 0, 0]); // two rates in one voucher
const tot = (j: number) => Math.round(n.rows.reduce((a, r) => a + r[j], 0) * 100) / 100;
assert.equal(tot(12), 100000 + 2500 + 40000 + 10000 + 1000 + 75550.4 + 50000 + 3200 + 20000);
assert.equal(tot(15), 13599.07 + 3600);
assert.equal(tot(16), 1000);
assert.equal(n.rows[4][4], 'RCM-UNREG', 'no GSTIN → default');
// a missed ledger shows up in the gross-total check
const drop = { ...map, taxable: colsOf(map.taxable).filter(j => hdr[j] !== 'Freight Inward') };
const n2 = normalise('purchases', rows, h, drop, d, '2026-27');
assert.equal(n2.stats.grossOff, 1);
assert.deepEqual(n2.stats.grossOffEg, ['SST/26-27/1042']);

// ordinary single-column register still maps one column per field
const plain = [
  ['Bill Date', 'Bill No', 'Vendor Name', 'GSTIN', 'Taxable Value', 'CGST', 'SGST', 'IGST', 'CGST %'],
  ['01-10-2026', 'B1', 'X', '27AABCS1234F1Z5', 1000, 90, 90, 0, 9],
];
const pm = guessMappingFull(plain, 0, PURCHASE_FIELDS);
assert.ok(!looksColumnar(plain[0]));
assert.equal(pm.taxable, 4);
assert.equal(pm.cgst, 5);
assert.equal(pm.sgst, 6);
const pn = normalise('purchases', plain, 0, pm, d, '2026-27');
assert.deepEqual(pn.rows[0].slice(12, 17), [1000, 90, 90, 0, 0]);
console.log('tally.ts: all checks passed');

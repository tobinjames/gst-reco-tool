// Revenue reconciliation on the synthetic demo company (test/fixtures/revenue, made by make_revenue_fixtures.py):
// sales register route and GL dump route, both statements, user reasons, B2CS months, workbook formulas and read-back.
// Run: npx tsx test/revenue.test.ts (or npm test for all)
import { readFileSync, readdirSync } from 'fs';
import assert from 'assert/strict';
import { SALES_FIELDS, detectHeader, guessMappingFull, normalise, readTable } from '../src/register';
import { processGstr1 } from '../src/gstr1';
import { parseBooks } from '../src/flow/books';
import type { ThreeB } from '../src/gstr3b';
import { RevResult, emptyUser, runRevenue } from '../src/revenue/engine';
import { glDocs, guessRevenue } from '../src/revenue/ledgers';
import {
  applyCustomerAliases,
  customerSuggestions,
  customersFromG1,
  parseCustomerList,
  resolveCustomers,
} from '../src/revenue/customers';
import { nameKey } from '../src/suppliers';
import { classify } from '../src/build';
import { ITEMS_SHEET, readItemsSheet, revenueWorkbook } from '../src/revenue/workbook';
import { Reader } from '../src/xlsx';
import { readRows } from '../src/integrity';

const dir = new URL('./fixtures/revenue/', import.meta.url);
const fx = (f: string) => new Uint8Array(readFileSync(new URL(f, dir)));
const FY = '2025-26';

// ---- inputs
const g1Files = readdirSync(dir)
  .filter(f => f.startsWith('demo_rev_gstr1_'))
  .map(f => ({ name: f, json: JSON.parse(readFileSync(new URL(f, dir), 'utf8')) }));
const g1 = processGstr1(g1Files, FY, '2026-27', []).rows;
const tbJson = JSON.parse(readFileSync(new URL('demo_rev_3b.json', dir), 'utf8'));
const threeB: ThreeB[] = Object.entries(tbJson).map(([period, rows]: [string, any]) => ({
  gstin: '32AAACJ1234K1Z5',
  legalName: 'DEMO',
  year: FY,
  month: '',
  period,
  arn: '',
  arnDate: '',
  rows: Object.entries(rows).map(([key, v]: [string, any]) => ({
    key,
    amt: [v[0], v[2], v[3], v[1], v[4]],
    info: [FY, period, '', key, 'demo'],
  })),
  pay: [],
  problems: [],
}));
const sh = (await readTable('demo_rev_sales_register.xlsx', fx('demo_rev_sales_register.xlsx')))[0];
const hdr = detectHeader(sh.rows, SALES_FIELDS);
const map = guessMappingFull(sh.rows, hdr, SALES_FIELDS);
const reg = normalise('sales', sh.rows, hdr, map, { noGstinZeroTax: 'EXEMPT' } as any, FY).rows;
assert.ok(reg.length > 40, 'register read');

const line = (r: RevResult, code: string) => r.st1.find(l => l.code === code)?.amt ?? [0, 0];
const near = (a: number[], b: number[], msg: string) =>
  assert.ok(Math.abs(a[0] - b[0]) < 0.02 && Math.abs(a[1] - b[1]) < 0.02, `${msg}: got ${a} expected ${b}`);

// ---- 1. sales register route
const r = runRevenue({ fy: FY, books: { rows: reg, source: 'register' }, g1, threeB, user: emptyUser() });
near(r.books.turnover, [2118000, 2118000 * 0.18 - 0.18 * 8000], 'books turnover (exempt has no tax)');
near(line(r, 'NOT_REPORTED'), [-30000, -5400], 'P1 not reported + P6 credit note');
near(line(r, 'NEXT_FY'), [-30000, -5400], 'P2 next year');
near(line(r, 'PREV_FY'), [25000, 4500], 'P3 previous year');
near(line(r, 'VALUE'), [-1000, -180], 'P5 value');
near(line(r, 'B2C'), [-2000, -360], 'P7 B2C');
near(line(r, 'NOT_IN_BOOKS'), [15000, 2700], 'P8 not in books');
near(line(r, 'NIL'), [-3000, 0], 'P9 exempt');
near(line(r, 'unx'), [0, 0], 'nothing unexplained');
near(r.g1Turnover, [2092000, 2092000 * 0.18 - 0.18 * 5000], 'GSTR-1 turnover');
assert.equal(r.timing.count, 1, 'P4 Jul → Aug is timing, not in the statement');
assert.ok(!r.st1.some(l => l.code === 'ROUND'), 'no rounding lines');
// statement 2: Aug short / Sep excess cancel; Mar short 20,000
const t31a = r.st2.find(l => l.code === 'T3.1a')!;
near(t31a.amt, [-20000, -3600], '3B 3.1(a) short in Mar only');
near(r.st2.find(l => l.code === 'unx')!.amt, [0, 0], '3B: nothing unexplained');
assert.ok(!r.st2.some(l => l.code === 'T3.1c'), 'Table 8 agrees with 3.1(c)');
near(r.tbTotal, [2072000, 2072000 * 0.18 - 0.18 * 5000], '3B total');
const aug = r.months.find(m => m.ym === '2025-08')!;
near([aug.tb![0] - aug.g1[0], 0], [-10000, 0], 'Aug month view shows the 3B shortfall');

// ---- 2. user reasons and B2CS months
const p1 = r.items.find(i => i.bNo === 'SAL/1103')!;
const dec = r.items.find(i => i.bNo === 'CS/1225')!;
assert.equal(dec.b2cMonth, '2025-12');
const u = emptyUser();
u.reasons[p1.id] = { r: 'CANCELLED', note: 'cancelled, not reported' };
// the Dec cash sale: 2,000 of it was really reported nowhere -> moved out of B2CS
u.b2c[dec.bIds.join(',')] = '';
const r2 = runRevenue({ fy: FY, books: { rows: reg, source: 'register' }, g1, threeB, user: u });
near(line(r2, 'CANCELLED'), [-40000, -7200], 'user reason moves P1');
near(line(r2, 'NOT_REPORTED'), [10000 - 20000, 1800 - 3600], 'credit note + Dec cash sale now not reported');
near(line(r2, 'B2C'), [18000, 3240], 'B2C now: Dec B2CS 18,000 has no books entry');
near(line(r2, 'unx'), [0, 0], 'still ties');
assert.equal(r2.items.find(i => i.bNo === 'SAL/1103')!.remark, 'cancelled, not reported');

// ---- 3. GL dump route
const dump = parseBooks('demo_rev_gl_dump.xlsx', (await readTable('demo_rev_gl_dump.xlsx', fx('demo_rev_gl_dump.xlsx')))[0].rows);
assert.equal(guessRevenue('Interest Received'), 'nosupply');
assert.equal(guessRevenue('Local Sales 18%'), 'taxable');
assert.equal(guessRevenue('Exempt Sales'), 'exempt');
assert.equal(guessRevenue('Output CGST 9%'), null);
assert.equal(guessRevenue('HDFC Bank'), null);
const gl = glDocs([dump], {}, FY);
assert.deepEqual(gl.ledgers.map(l => [l.name, l.guess]).sort(), [
  ['Exempt Sales', 'exempt'],
  ['Interest Received', 'nosupply'],
  ['Interstate Sales 18%', 'taxable'],
  ['Local Sales 18%', 'taxable'],
]);
const rg = runRevenue({ fy: FY, books: { rows: gl.docs.map(d => d.row), source: 'gl' }, g1, threeB, user: emptyUser() });
near(rg.books.revenue, [2120000, r.books.turnover[1]], 'GL revenue includes interest');
near(rg.books.excluded, [2000, 0], 'interest excluded as not a supply');
near(rg.books.turnover, r.books.turnover, 'same turnover as the register');
near(line(rg, 'B2C'), [-2000 - 40000 + 10000, -360 - 7200 + 1800], 'GL: entries without GSTIN not in GSTR-1 are taken as B2C');
near(line(rg, 'NEXT_FY'), [-30000, -5400], 'GL: next year by invoice no.');
near(line(rg, 'unx'), [0, 0], 'GL: ties');
// customer names: GSTR-1 has no names, the GL dump has no GSTINs; a customer is mapped when its invoice no. AND amount agree
const glRows = gl.docs.map(d => d.row);
const dirC = customersFromG1(g1, glRows);
const sug = customerSuggestions(glRows, dirC);
const ker = sug.find(x => x.name === 'Kerala Customer Pvt Ltd')!;
assert.equal(ker.exact, undefined, 'no books entry carries the GSTIN: not a "same name" mapping');
assert.equal(ker.inv?.gstin, '32AAACK1111A1Z1', 'Kerala: invoice no. + amount agree');
assert.ok(ker.inv!.hits >= 1 && ker.inv!.sample.length >= 1, 'the agreeing invoices are kept as the basis');
assert.equal(sug.find(x => x.name === 'Chennai Customer Pvt Ltd')!.inv?.gstin, '33AAACC2222B1Z2');
assert.equal(sug.find(x => x.name === 'Cash Sales')!.cands.length, 0, 'cash sales: no GSTR-1 customer');
// a spelling variant with no invoice in GSTR-1 is suggested by name, for the user to confirm
const variant = [[45900, 'SAL/9999', 'Kerala Customers P. Ltd', '', '', 'B2B', 'INV', '', '', null, 18, 1000, 90, 90, 0, 0]];
const vs = customerSuggestions(variant, dirC);
assert.equal(vs[0].cands[0]?.gstin, '32AAACK1111A1Z1', 'name variant suggested');
assert.ok(!vs[0].exact && !vs[0].strong && !vs[0].inv, 'variant needs the user');
// same invoice no. but another amount: suggested (number only), never mapped
const kerNo = [...dirC.get('32AAACK1111A1Z1')!.inv.values()][0];
const wrongAmt = [[45900, kerNo.no, 'Some Other Buyer', '', '', 'B2B', 'INV', '', '', null, 18, kerNo.tx + 500, 90, 90, 0, 0]];
const wa = customerSuggestions(wrongAmt, dirC)[0];
assert.ok(!wa.inv && wa.cands[0].noHits === 1 && wa.cands[0].invHits === 0, 'number alone does not map');
// rules and their order: user > list > same name > invoice no. + amount; '?' = ask me
const RULES = { list: true, exact: true, invAmt: true, invAuto: true, invMin: 1 };
let mp = resolveCustomers(sug, RULES, new Map(), {});
assert.equal(mp.get(ker.key)?.how, 'invamt');
assert.match(mp.get(ker.key)!.detail, /agree with GSTR-1 under this GSTIN by number and amount/);
assert.equal(resolveCustomers(sug, { ...RULES, invAuto: false }, new Map(), {}).has(ker.key), false, 'suggest only');
assert.equal(resolveCustomers(sug, { ...RULES, invMin: 99 }, new Map(), {}).has(ker.key), false, 'too few invoices');
assert.equal(resolveCustomers(sug, RULES, new Map(), { [ker.key]: '?' }).has(ker.key), false, 'user asked to choose');
mp = resolveCustomers(sug, RULES, new Map([[ker.key, '32AAACK1111A1Z1']]), {});
assert.equal(mp.get(ker.key)?.how, 'list');
assert.equal(resolveCustomers(sug, RULES, new Map(), { [ker.key]: '' }).get(ker.key)?.how, 'user');
const aliases = new Map<string, string>(
  [...resolveCustomers(sug, RULES, new Map(), {})].map(([k, m]) => [k, m.gstin] as [string, string]),
);
const rgc = runRevenue({
  fy: FY,
  books: { rows: applyCustomerAliases(glRows, aliases), source: 'gl' },
  g1,
  threeB,
  user: emptyUser(),
});
near(line(rgc, 'NOT_REPORTED'), line(r, 'NOT_REPORTED'), 'GL + customer names: same as the register');
near(line(rgc, 'B2C'), line(r, 'B2C'), 'GL + customer names: B2C as the register');
near(line(rgc, 'unx'), [0, 0], 'GL + customer names: ties');
// a customer list file (name + GSTIN, Zoho export with extra columns): recognised on drop, names give the GSTINs
const cf = await classify('demo_rev_customers.xlsx', fx('demo_rev_customers.xlsx'), async () => {
  throw new Error('no pdf');
});
assert.equal(cf[0].kind, 'Customer list');
assert.ok(cf[0].customers!.some(x => x.name === 'Kerala Customer Pvt Ltd' && x.gstin === '32AAACK1111A1Z1'));
assert.ok(!cf[0].customers!.some(x => x.name === 'Cash Sales'), 'no GSTIN, not listed');
assert.equal(parseCustomerList(sh.rows), null, 'a sales register is not a customer list');
const rgl = runRevenue({
  fy: FY,
  books: {
    rows: applyCustomerAliases(glRows, new Map(cf[0].customers!.map(x => [nameKey(x.name), x.gstin] as [string, string]))),
    source: 'gl',
  },
  g1,
  threeB,
  user: emptyUser(),
});
near(line(rgl, 'NOT_REPORTED'), line(r, 'NOT_REPORTED'), 'GL + customer list: same as the register');
near(line(rgl, 'B2C'), line(r, 'B2C'), 'GL + customer list: B2C as the register');
near(line(rgl, 'unx'), [0, 0], 'GL + customer list: ties');
// ignore the interest ledger entirely
const rg2 = runRevenue({
  fy: FY,
  books: { rows: glDocs([dump], { 'Interest Received': 'none' }, FY).docs.map(d => d.row), source: 'gl' },
  g1,
  threeB,
  user: emptyUser(),
});
near(rg2.books.revenue, [2118000, r.books.turnover[1]], 'ignored ledger is left out');

// ---- 4. workbook: formulas, lists, read back
const bytes = await revenueWorkbook(r2, {
  gstin: '32AAACJ1234K1Z5',
  legalName: 'Demo',
  fy: FY,
  source: 'Sales register',
  files: [],
  audit: [],
  builtAt: 'test',
  g1,
  threeB,
  bookRows: reg,
});
const rd = await Reader.open(bytes);
for (const s of [
  'Read Me',
  'Dashboard',
  'Books_to_GSTR1',
  'GSTR1_to_3B',
  'Month_wise',
  'B2C_Month',
  'Month_3B',
  ITEMS_SHEET,
  'G1_Rows',
  '3B_Rows',
  'Lists',
])
  assert.ok(rd.has(s), s);
const back = readItemsSheet(await readRows(rd, ITEMS_SHEET, 1));
assert.equal(back.length, r2.items.length);
assert.equal(back.find(x => x.id === p1.id)!.reason, 'Cancelled invoice');
assert.equal(back.find(x => x.id === dec.id)!.b2c, '');
assert.equal(back.find(x => x.id === r2.items.find(i => i.bNo === 'CS/0425')!.id)!.b2c, 'Apr-25');
// write the workbook for the LibreOffice recalculation check (test/revenue_recalc.py)
if (process.env.REV_OUT) {
  const { writeFileSync } = await import('fs');
  writeFileSync(process.env.REV_OUT, bytes);
}
console.log('revenue: register, GL dump, reasons, B2CS months and workbook OK');

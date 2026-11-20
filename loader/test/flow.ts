// GST ledger flow: readers + engine on the synthetic demo company (test/fixtures/ledgers, made by make_ledger_fixtures.py)
// Run: npx tsx test/flow.ts
import { readFileSync } from 'fs';
import assert from 'assert/strict';
import { readTable } from '../src/register';
import { parsePortal } from '../src/flow/portal';
import { accountsOf, guessAccount, isGstName, parseBooks } from '../src/flow/books';
import { HEAD_KEYS, Inputs, bucketLabel, run, ymLabel } from '../src/flow/engine';

const fx = (f: string) => new Uint8Array(readFileSync(new URL('./fixtures/ledgers/' + f, import.meta.url)));
const portal = [];
for (const f of ['demo_credit_ledger.csv', 'demo_cash_ledger.csv', 'demo_liability_register.csv']) {
  const p = parsePortal((await readTable(f, fx(f)))[0].rows);
  assert.deepEqual(p.problems, [], f + ' parses without balance breaks');
  portal.push({ file: f, ledger: p });
}
const gl = parseBooks('demo_tally_gl_dump.xlsx', (await readTable('demo_tally_gl_dump.xlsx', fx('demo_tally_gl_dump.xlsx')))[0].rows);
assert.equal(gl.format, 'tally-dump');
assert.deepEqual(gl.skipped, [{ reason: '"Purchase Order" vouchers do not post to the books', count: 4 }]);
const accMap: Inputs['accMap'] = {};
for (const a of accountsOf([gl])) if (isGstName(a.name)) accMap[a.name] = guessAccount(a.name) as any;
console.log(accMap);
const inp: Inputs = { fy: '2025-26', from: '2025-04', to: '2026-03', books: [gl], portal, accMap, overrides: {},
  bal: { books: { 'Input CGST': { opening: 10000 }, 'Input SGST': { opening: 9500 } }, portal: {} } };
const r = run(inp);
// every difference row is fully accounted for by its reconciling items (opening rows by the opening items)
const keys = new Set([...Object.keys(r.flowBooks.rows), ...Object.keys(r.flowPortal.rows)]);
for (const k of keys) for (const h of HEAD_KEYS) {
  const d = (r.flowBooks.rows[k]?.[h] || 0) - (r.flowPortal.rows[k]?.[h] || 0);
  const it = r.items.filter(i => i.bucket === k).reduce((s, i) => s + i.amt[h], 0);
  assert.ok(Math.abs(d - it) < 0.01, `${k} ${h}: difference ${d} vs items ${it}`);
}
for (const blk of ['itc', 'liab', 'cash'] as const) for (const h of HEAD_KEYS) {
  const d = r.flowBooks.opening[blk][h] - r.flowPortal.opening[blk][h];
  const it = r.items.filter(i => i.bucket === 'opening' && i.block === blk).reduce((s, i) => s + i.amt[h], 0);
  assert.ok(Math.abs(d - it) < 0.01, `opening ${blk} ${h}`);
}
const find = (cat: string, bucket: string) => r.items.find(i => i.category === cat && i.bucket === bucket);
const amt = (i: any) => Object.fromEntries(HEAD_KEYS.filter(h => Math.abs(i.amt[h]) > 0.004).map(h => [h, i.amt[h]]));
assert.deepEqual(amt(find('Opening balance', 'opening')), { CGST: 500 }, '1. opening');
assert.deepEqual(amt(find('Netted on portal', 'itc_reversed')), { CGST: 1000, SGST: 1000 }, '2. reversal netted');
assert.deepEqual(find('Timing', 'itc_availed')!.months, ['2025-07', '2025-08'], '3. timing Jul → Aug');
assert.deepEqual(find('Timing', 'itc_availed')!.shift, { IGST: 0, CGST: 1000, SGST: 1000, CESS: 0, ALL: 0 });
assert.deepEqual(amt(find('Tax head', 'out_tax')), { IGST: 1800, CGST: -900, SGST: -900 }, '4. wrong head');
assert.deepEqual(amt(find('Fee booked to tax', 'paid_cash')), { CGST: 50, SGST: 50 }, '5. late fee');
const un = r.items.filter(i => !i.explained);
assert.equal(un.length, 1, 'only the planted unreported sale is unexplained');
assert.deepEqual(amt(un[0]), { CGST: 300, SGST: 300 }, '6. Nov sale');
assert.deepEqual(un[0].months, ['2025-11']);
assert.equal(r.books.filter(e => un[0].books.includes(e.id) && e.line.party === 'Walk-in Customer').length, 2, 'the Nov output-tax lines are listed behind the item');
assert.deepEqual(amt(find('Tax head', 'itc_setoff')), { IGST: -3000, CGST: 3000 }, '7. cross-utilisation');
assert.deepEqual(amt(find('After the period', 'itc_availed')), { IGST: 5000 }, '8. Mar ITC claimed in Apr-26');
assert.deepEqual(amt(find('After the period', 'cash_deposit')), { ALL: -9000 }, '10. Mar challan paid in Apr-26');
assert.equal(r.checks.filter(c => !c.ok).length, 0, 'portal ledgers agree with each other');
assert.equal(r.flowPortal.closing.liab.CGST, 0);
assert.equal(r.flowBooks.closing.itc.IGST, 8000, 'Mar ITC 5,000 + IGST credit 3,000 credited to CGST in books');

// user changes: moving the Nov sale's lines to another flow row is allowed only within the block, and is logged by the caller
const nov = r.books.filter(e => un[0].books.includes(e.id) && e.head !== 'ALL');
const r2 = run({ ...inp, overrides: Object.fromEntries(nov.map(e => [e.id, { period: '2025-12' }])) });
assert.ok(r2.items.some(i => !i.explained && i.months.includes('2025-12')), 'a changed period is still checked against the portal');

// reconciliation statement: books closing + items = portal closing, one line per real event
import { reconcile } from '../src/flow/recon';
const rc = reconcile(r);
for (const s of rc.statements) for (const h of HEAD_KEYS) assert.ok(Math.abs(s.gap[h]) < 0.01, `${s.block} ${h} reconciles to the portal`);
const st = (b: string) => rc.statements.find(s => s.block === b)!;
assert.deepEqual(st('itc').lines.map(e => e.category), ['Opening balance', 'After the period', 'Tax head'], 'ITC: opening, Mar ITC claimed in Apr, cross-utilisation');
assert.deepEqual(st('liab').lines.map(e => [e.category, e.explained]), [['Tax head', true], ['Fee booked to tax', true], ['Amount differs', false]], 'liability: wrong head, late fee, unreported Nov sale');
assert.deepEqual(st('cash').lines.map(e => e.category), ['After the period'], 'cash: challan paid after the period');
// opening differences in every ledger stay in their own statement (they once merged into the ITC one)
const rcO = reconcile(run({ ...inp, bal: { ...inp.bal, portal: { liability: { CGST: { opening: 700 } }, cash: { IGST: { opening: 450 } } } } }));
for (const s of rcO.statements) for (const h of HEAD_KEYS) assert.ok(Math.abs(s.gap[h]) < 0.01, `with openings everywhere: ${s.block} ${h} reconciles (gap ${s.gap[h]})`);
assert.deepEqual(rcO.statements.map(s => s.lines.filter(e => e.category === 'Opening balance').length), [1, 1, 1], 'one opening line per ledger');
const netted = rc.noEffect.filter(e => e.category === 'Netted on portal');
assert.equal(netted.length, 1, 'the June reversal is one event, not two rows');
assert.ok(rc.noEffect.every(e => HEAD_KEYS.every(h => Math.abs(e.effect[h]) < 0.005)), 'no-effect events do not move the closing');

const fmt = (x: number) => x.toFixed(2).padStart(12);
if (process.env.SHOW) {
for (const [name, st] of [['BOOKS', r.flowBooks], ['PORTAL', r.flowPortal]] as const) {
  console.log('\n' + name);
  for (const blk of ['itc', 'liab', 'cash', 'fee'] as const) {
    console.log(' ', blk, 'opening', HEAD_KEYS.map(h => fmt(st.opening[blk][h])).join(''));
    for (const [k, v] of Object.entries(st.rows)) if ((k.startsWith(blk + ':') || (!k.includes(':') && k.startsWith({ itc: 'itc', liab: '', cash: 'cash', fee: 'fee' }[blk]) && (blk !== 'liab' || /^(out|paid)/.test(k))))) console.log('   ', bucketLabel(k).padEnd(40), HEAD_KEYS.map(h => fmt(v[h])).join(''));
    console.log(' ', blk, 'closing', HEAD_KEYS.map(h => fmt(st.closing[blk][h])).join(''));
  }
}
console.log('\nITEMS');
for (const it of r.items) console.log(it.id, it.explained ? 'E' : 'U', it.category.padEnd(24), it.bucket.padEnd(14), HEAD_KEYS.map(h => it.amt[h] ? `${h} ${it.amt[h]}` : '').filter(Boolean).join(', '), '|', it.text);
console.log('\nCHECKS'); for (const c of r.checks.filter(c => !c.ok)) console.log(c.area, ymLabel(c.period), c.head, c.a, c.b, c.diff);
console.log('\nWARNS'); r.warns.forEach(w => console.log(w.level, w.msg));
if (process.env.DBG) for (const x of r.monthly.filter(x => x.bucket === process.env.DBG).sort((a, b) => a.month.localeCompare(b.month))) console.log(x.bucket, x.head, x.month, x.b, x.p);
}
console.log('flow: all checks passed', r.items.length, 'items');

// workbook: written with cached values; test/flow_recalc.py recalculates it in LibreOffice and compares
import { flowWorkbook } from '../src/flow/workbook';
import { writeFileSync } from 'fs';
const wbBytes = await flowWorkbook(r, { gstin: '32AAACJ1234K1Z5', legalName: 'DEMO TRADERS PRIVATE LIMITED', fy: '2025-26', from: '2025-04', to: '2026-03', builtAt: '20-11-2026', audit: [],
  files: [{ name: 'demo_credit_ledger.csv', kind: 'Electronic Credit Ledger', note: '' }] });
if (process.env.OUT) { writeFileSync(process.env.OUT, wbBytes); console.log('workbook written to', process.env.OUT); }

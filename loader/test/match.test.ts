// Matching options, manual matches and "keep apart" (code X). Run: npx tsx test/match.test.ts (or npm test for all)
import assert from 'assert/strict';
import { DEFAULT_MATCH, MATCH_PRESETS, MatchOptions, reconPurchases } from '../src/recon';

const G = '27AABCS1234F1Z5', G2 = '27AAFCK5678M1Z2';
// Books_Purchase: [date, no, name, gstin, type, itc, status, fy, hsn, uqc, qty, rate, taxable, C, S, I, cess, remarks]
const bk = (no: string, date: string, t: number, tax: number, g = G) => [date, no, 'Sup', g, 'B2B', 'Inputs', 'Availed', '2026-27', '', '', null, 18, t, tax / 2, tax / 2, 0, 0, ''];
// RAW_2B: [key, tx, c, s, i, cess, src, bfy, period, section, gstin, name, dtype, no, date, dfy, avl, rcm, ims]
const tb = (no: string, date: string, t: number, tax: number, g = G) => ['8A', t, tax / 2, tax / 2, 0, 0, '2B', '2026-27', '102026', 'B2B', g, 'SUP', 'INV', no, date, '2026-27', 'Y', 'N', ''];
const books = [
  bk('INV-001', '01-10-2026', 1000, 180),          // exact
  bk('A/26-27/0451', '02-10-2026', 2000, 360),     // last digits
  bk('SI 7781', '03-10-2026', 3000, 540),          // digits: 2B "SI-07781"
  bk('XYZ', '10-10-2026', 4000, 720),              // date+amount, 2B dated 12-10 (±2 days)
  bk('Q-9', '15-10-2026', 5000, 900, G2),          // other GSTIN
  bk('P1', '20-10-2026', 600, 108), bk('P2', '21-10-2026', 400, 72), // supplier total vs one 2B doc
];
const twob = [
  tb('INV-001', '01-10-2026', 1000, 180),
  tb('INV/26-27/00451', '09-10-2026', 2000, 360),
  tb('SI-07781', '03-10-2026', 3000.5, 540),
  tb('ABC-1', '12-10-2026', 4000, 720),
  tb('Q-9', '15-10-2026', 5000, 900, '27AAKPR4321J1Z3'),
  tb('TOT-1', '22-10-2026', 1000, 180),
];
const run = (o: MatchOptions = DEFAULT_MATCH, manual = new Map(), keep = new Set<string>()) => {
  const r = reconPurchases(books, twob, '2026-27', '2027-28', new Map(), manual, [], o, keep);
  return { r, basis: Object.fromEntries(r.rows.filter(x => x[4] && x[5]).map(x => [x[4], x[1]])), stats: r.stats };
};
const std = run();
console.log(std.basis, std.stats);
assert.equal(std.basis['INV-001'], 'Invoice no.');
assert.equal(std.basis['SI 7781'], 'Digits of invoice no.');
assert.equal(std.basis['A/26-27/0451'], 'Last digits of invoice no.');
assert.equal(std.basis['XYZ'], 'GSTIN + date (±2d) + amount');
assert.equal(std.basis['Q-9'], 'Invoice no. + amount (GSTIN differs)');
assert.equal(std.basis['P1 + P2'], 'Supplier total (2 books / 1 2B docs)');
assert.equal(std.stats.matched, 6); assert.equal(std.stats.booksOnly, 0);

// strict: invoice no. only
const st = run(MATCH_PRESETS.Strict);
assert.deepEqual(Object.keys(st.basis), ['INV-001']);
// date window 0 days: XYZ no longer matches
const d0 = run({ ...DEFAULT_MATCH, days: 0 });
assert.equal(d0.basis['XYZ'], undefined);
// rule switched off
const noTail = run({ ...DEFAULT_MATCH, rules: { ...DEFAULT_MATCH.rules, tail: false } });
assert.equal(noTail.basis['A/26-27/0451'], undefined);
// tolerance: taxable differs by 0.5 → fine at Rs 1, not at Rs 0.1 (with 0%)
const tight = run({ ...DEFAULT_MATCH, tol: 0.1, pct: 0 });
assert.equal(tight.basis['SI 7781'], undefined);
// keep apart: separate INV-001 → both sides unmatched with code X
const kept = run(DEFAULT_MATCH, new Map(), new Set([`B|${G}|INV001`, `R|${G}|INV001`]));
assert.equal(kept.basis['INV-001'], undefined);
const xr = kept.r.rows.filter(x => x[26] === 'X');
assert.equal(xr.length, 2);
// manual: pair XYZ with TOT-1 explicitly, even though kept apart
const man = new Map([['M001', { b: new Set([`B|${G}|XYZ`]), r: new Set([`R|${G}|TOT1`]) }]]);
const mm = run(DEFAULT_MATCH, man, new Set([`B|${G}|XYZ`]));
assert.equal(mm.basis['XYZ'], 'Manual (M001)');
console.log('match.ts: all checks passed');

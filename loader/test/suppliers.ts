// Supplier names: approximate matching of books names (no GSTIN) to GSTR-2B suppliers, applied only when confirmed.
// Run: npx tsx test/suppliers.ts
import { readFileSync } from 'fs';
import assert from 'assert/strict';
import { applyAliases, nameKey, nameScore, supplierSuggestions, suppliersFrom2B } from '../src/suppliers';
import { classify, buildItc } from '../src/build';

const j = JSON.parse(readFileSync(new URL('./fixtures/gstr2b_102026_synthetic.json', import.meta.url), 'utf8'));
const sup = suppliersFrom2B([j]);
assert.equal(sup.size, 6);

assert.equal(nameKey('M/s. Kaveri Packaging Pvt. Ltd.'), nameKey('KAVERI PACKAGING PRIVATE LIMITED'));
assert.equal(nameKey('Mehta and Associates'), nameKey('MEHTA & ASSOCIATES'));
assert.ok(nameScore(nameKey('Gujrat Polymer'), nameKey('GUJARAT POLYMERS LLP')) >= 0.9, 'typos');
assert.ok(nameScore(nameKey('SST'), nameKey('Sharma Steel Traders')) >= 0.8, 'initials');
assert.ok(nameScore(nameKey('Ramesh Hardware'), nameKey('RAMESH TRADING COMPANY')) < 0.6, 'shared first word only');

// Books_Purchase rows: [date, no, name, gstin, type, itc, status, fy, hsn, uqc, qty, rate, taxable, C, S, I, cess, remarks]
const row = (no: string, name: string, t: number, c: number, s: number, i: number) => [46296, no, name, '', 'RCM-UNREG', 'Inputs', 'Availed', '2026-27', '', '', null, 18, t, c, s, i, 0, ''];
const books = [
  row('SST/26-27/1042', 'SST', 102500, 9225, 9225, 0),
  row('KP-889', 'M/s Kaveri Packaging Pvt Ltd', 51000, 3390, 3390, 0),
  row('GP-2026-567', 'Gujrat Polymer', 75550.40, 0, 0, 13599.07),
  row('MA/112', 'Mehta and Associates', 50000, 4500, 4500, 0),
  row('455', 'Ramesh Hardware Stores', 3200, 288, 288, 0),
  row('X-1', 'Totally Different Vendor', 500, 45, 45, 0),
];
const sug = supplierSuggestions(books, sup);
const by = Object.fromEntries(sug.map(x => [x.name, x]));
console.log(sug.map(x => `${x.name} → ${x.exact ?? x.cands.map(c => `${c.name} ${c.score}${c.invHits ? ' inv' : ''}`).join(' | ')}`).join('\n'));
assert.equal(by['M/s Kaveri Packaging Pvt Ltd'].exact, '27AAFCK5678M1Z2');
assert.equal(by['Mehta and Associates'].exact, '27AAEFM9753K1Z8');
assert.equal(by['SST'].exact, undefined); assert.equal(by['SST'].cands[0].gstin, '27AABCS1234F1Z5'); assert.equal(by['SST'].cands[0].invHits, 1);
assert.equal(by['Gujrat Polymer'].cands[0].gstin, '24AAKFG2468H1ZQ');
assert.equal(by['Ramesh Hardware Stores'].cands[0].gstin, '27AAKPR4321J1Z3', 'invoice-no. hit ranks first');
assert.equal(by['Totally Different Vendor'].cands.length, 0);

// confirmed names fill the GSTIN; unconfirmed stay blank
const rows = books.map(r => [...r]);
const n = applyAliases(rows, new Map([[nameKey('SST'), '27AABCS1234F1Z5'], [nameKey('Gujrat Polymer'), '24AAKFG2468H1ZQ']]));
assert.equal(n, 2); assert.equal(rows[0][3], '27AABCS1234F1Z5'); assert.equal(rows[0][4], 'B2B'); assert.equal(rows[4][3], '');

// end to end: confirmed SST matches its 2B invoice; unconfirmed Ramesh still matches by invoice no. + amount; Gujrat (different number format) matches
const files = await classify('2b.json', new Uint8Array(readFileSync(new URL('./fixtures/gstr2b_102026_synthetic.json', import.meta.url))), async () => { throw new Error('no pdf'); });
const r = await buildItc(new Uint8Array(readFileSync(new URL('../template/ITC_Recon_Template.xlsx', import.meta.url))), files, '2026-27', { purchases: rows });
console.log(r.recon);
console.log('suppliers.ts: all checks passed');

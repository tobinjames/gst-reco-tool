// The Loader reads every section of GSTN's own sample files (test/fixtures/gstn, from the GST developer portal).
// Run: npx tsx test/gstn.test.ts (or npm test for all)
import { readFileSync } from 'fs';
import assert from 'assert/strict';
import { process2b } from '../src/gstr2b';
import { processGstr1 } from '../src/gstr1';
import { normalise, PURCHASE_FIELDS } from '../src/register';
import { fyOfPeriod, fyShift } from '../src/util';

const load = (f: string) => JSON.parse(readFileSync(new URL('./fixtures/gstn/' + f, import.meta.url), 'utf8'));
const tax = (a: number[]) => Math.round((a[1] + a[2] + a[3] + a[4]) * 100) / 100;

// ---- GSTR-2B (format v4.5)
{
  const j = load('gstr2b_api_v4.5_sample.json');
  j.data.rtnprd = '052025'; // sample dates are old; put the 2B in FY 2025-26
  // an ISD credit note is "ISDC" in the format
  j.data.docdata.isd[0].doclist.push({ ...j.data.docdata.isd[0].doclist[0], doctyp: 'ISDC', docnum: 'C-1' });
  const warns: any[] = [];
  const { rows } = process2b([{ name: '2b', json: j }], '2025-26', '2024-25', '2026-27', warns);
  const sec = (s: string) => rows.filter(r => r.info[3] === s);
  const dd = j.data.docdata;
  // every document in every section the Loader uses becomes a row
  assert.equal(sec('B2B').length, dd.b2b.flatMap((s: any) => s.inv).length);
  assert.equal(sec('CDNR').length, dd.cdnr.flatMap((s: any) => s.nt).length);
  assert.equal(sec('ECOM').length, dd.ecom.flatMap((s: any) => s.inv).length, 'e-commerce operator documents are read');
  assert.equal(sec('ECOMA').length, dd.ecoma.flatMap((s: any) => s.inv).length);
  assert.equal(sec('IMPG').length, dd.impg.length);
  // SEZ imports are grouped by supplier: { ctin, boe: [...] }
  const sez = sec('IMPGSEZ');
  assert.equal(sez.length, dd.impgsez.flatMap((s: any) => s.boe).length);
  assert.ok(
    sez.every(r => r.info[7] && tax(r.amt) > 0),
    'SEZ bills of entry keep their number and amounts',
  );
  assert.equal(sez[0].info[4], dd.impgsez[0].ctin, 'SEZ supplier GSTIN kept');
  assert.equal(sez[0].key, '8G', 'SEZ imports count in Table 8G');
  // credit notes are negative
  assert.ok(sec('CDNR').every(r => tax(r.amt) < 0));
  const isdc = sec('ISD').find(r => r.info[6] === 'ISDC')!;
  assert.ok(tax(isdc.amt) < 0, 'ISD credit note (ISDC) is negative');
  // the ITC reversal section is flagged, not silently dropped
  assert.ok(
    warns.some(w => w.level === 'warn' && /ITC reversal/.test(w.msg)),
    'ITC reversal flagged',
  );
  console.log(`2B: ${rows.length} rows from ${Object.keys(dd).length} sections`);
}

// ---- GSTR-1: the 2026 format (v5.0) splits advances into B2B / B2C; the 2025 format (v4.1) has plain lists
for (const f of ['gstr1_api_v5.0_sample.json', 'gstr1_api_v4.1_sample.json']) {
  const j = load(f);
  j.fp = '052025';
  const warns: any[] = [];
  const r: any = processGstr1([{ name: f, json: j } as any], fyOfPeriod(j.fp), fyShift(fyOfPeriod(j.fp), 1), warns);
  const docs = r.docs as any[];
  const count = (s: string) => docs.filter(d => d.section === s).length;
  const nested = (v: any): any[] => (Array.isArray(v) ? v : v && typeof v === 'object' ? Object.values(v).flatMap(nested) : []);
  for (const s of ['at', 'ata', 'txpd', 'txpda'])
    assert.equal(count(s.toUpperCase()), nested(j[s]).flatMap((p: any) => p.itms).length, `${f}: ${s} rows`);
  assert.equal(count('B2B'), j.b2b.flatMap((s: any) => s.inv).length, `${f}: b2b`);
  assert.ok(r.hsn.length >= 2, `${f}: HSN B2B and B2C both read`);
  console.log(`${f}: ${docs.length} documents, ${r.hsn.length} HSN lines`);
}

// ---- GSTIN pattern: GSTN allows 1-9 and A-J as the 14th character as well as Z
{
  const row = (g: string) => [['01-05-2025', 'X-1', 'Sup', g, 100, 9, 9]];
  const map = { date: 0, no: 1, party: 2, gstin: 3, taxable: 4, cgst: 5, sgst: 6 };
  const d = { noGstinZeroTax: 'EXEMPT', itcType: 'Inputs', status: 'Availed', claimFy: '2025-26', purNoGstin: 'RCM-UNREG' };
  for (const g of ['27AABCS1234F1Z5', '27AABCS1234F1A5', '27AABCS1234F15X'])
    assert.equal(normalise('purchases', [['h'], ...row(g)], 0, map, d, '2025-26').rows[0][4], 'B2B', g);
  assert.notEqual(normalise('purchases', [['h'], ...row('27AABCS1234F1K5')], 0, map, d, '2025-26').rows[0][4], 'B2B');
  void PURCHASE_FIELDS;
}
console.log('gstn.test.ts: all checks passed');

// Next-FY returns count for this FY only for April to October (filed up to 30 November):
// ICAI Technical Guide on GSTR-9 (Dec 2025), FAQ 3 and Tables 8A / 8C / 10 / 11 / 13.
// Run: npx tsx test/window.test.ts (or npm test for all)
import assert from 'assert/strict';
import { process2b } from '../src/gstr2b';
import { processGstr1 } from '../src/gstr1';

const cy = '2025-26',
  py = '2024-25',
  ny = '2026-27';
// a FY 2025-26 invoice reported late, in a next-FY 2B
const twoB = (rtnprd: string) => ({
  data: {
    rtnprd,
    gstin: '27AAACD0000A1Z9',
    docdata: {
      b2b: [
        {
          ctin: '27AABCS1234F1Z5',
          trdnm: 'S',
          inv: [{ inum: 'L-1', dt: '15-03-2026', txval: 100, cgst: 9, sgst: 9, igst: 0, cess: 0, itcavl: 'Y', rev: 'N' }],
        },
      ],
    },
  },
});
const key2b = (p: string) => process2b([{ name: p, json: twoB(p) }], cy, py, ny, []).rows[0].key;
assert.equal(key2b('102026'), '8A+8C', 'October 2B of next FY counts');
assert.equal(key2b('112026'), '', 'November 2B of next FY is filed after 30 November: not in 8A');
assert.equal(key2b('042026'), '8A+8C');

// a debit note for a FY 2025-26 invoice in a next-FY GSTR-1 (Table 10)
const g1 = (fp: string) => ({
  gstin: '27AAACD0000A1Z9',
  fp,
  cdnr: [
    {
      ctin: '27AABCS1234F1Z5',
      nt: [
        {
          ntty: 'D',
          nt_num: 'D-' + fp,
          nt_dt: '20-03-2026',
          val: 118,
          pos: '27',
          rchrg: 'N',
          inv_typ: 'R',
          itms: [{ num: 1, itm_det: { rt: 18, txval: 100, camt: 9, samt: 9, iamt: 0, csamt: 0 } }],
        },
      ],
    },
  ],
});
const key1 = (fp: string) =>
  (processGstr1([{ name: fp, json: g1(fp) } as any], cy, ny, []) as any).rows.find(
    (r: any) => r.info?.includes?.('D-' + fp) || JSON.stringify(r).includes('D-' + fp),
  )?.key;
assert.equal(key1('102026'), '10', 'October GSTR-1 of next FY: Table 10');
assert.notEqual(key1('112026'), '10', 'November GSTR-1 of next FY: outside the window');
console.log('window.test.ts: all checks passed');

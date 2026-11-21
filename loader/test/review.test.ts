// Fixes from the independent review of v0.12 (Nov 2026): 2B amendments, purchase debit notes, returns loaded twice,
// CSV invoice numbers with leading zeros, month/day dates, and matching speed. Run: npx tsx test/review.test.ts
import assert from 'assert/strict';
import { isSuperseded, process2b } from '../src/gstr2b';
import { processGstr1 } from '../src/gstr1';
import { dateOrder, normalise, readTable } from '../src/register';
import { reconPurchases } from '../src/recon';
import type { Warn } from '../src/util';

const cy = '2025-26',
  py = '2024-25',
  ny = '2026-27';
const S = '27AABCS1234F1Z5';
const inv = (inum: string, txval: number, extra: any = {}) => ({
  inum,
  dt: '10-04-2025',
  txval,
  cgst: txval * 0.09,
  sgst: txval * 0.09,
  igst: 0,
  cess: 0,
  itcavl: 'Y',
  rev: 'N',
  ...extra,
});
const twoB = (rtnprd: string, docdata: any) => ({ data: { rtnprd, gstin: '27AAACD0000A1Z9', docdata } });
const sum8A = (rows: { key: string; amt: number[] }[]) =>
  rows.filter(r => r.key.startsWith('8A')).reduce((a, r) => a + r.amt[0], 0);

// ---- 2B amendments: GSTN GSTR-2B FAQ 8 - the amendment carries the full revised value, so only the latest version counts
{
  const w: Warn[] = [];
  const { rows } = process2b(
    [
      // given out of order on purpose: the June amendment replaces the April original
      {
        name: 'jun',
        json: twoB('062025', { b2ba: [{ ctin: S, inv: [inv('A-1', 1200, { oinum: 'A-1', oidt: '10-04-2025' })] }] }),
      },
      { name: 'apr', json: twoB('042025', { b2b: [{ ctin: S, inv: [inv('A-1', 1000), inv('A-2', 500)] }] }) },
    ],
    cy,
    py,
    ny,
    w,
  );
  assert.equal(sum8A(rows), 1700, '8A = revised A-1 (1,200) + A-2 (500), not 2,700');
  const orig = rows.find(r => r.info[3] === 'B2B' && r.info[7] === 'A-1')!;
  assert.equal(orig.key, '', 'original is set aside');
  assert.match(String(orig.info[12]), /^Superseded by B2BA A-1 in 06\/2025/);
}
// amendment that changes the invoice number: linked by oinum
{
  const { rows } = process2b(
    [
      { name: 'apr', json: twoB('042025', { b2b: [{ ctin: S, inv: [inv('A-9', 1000)] }] }) },
      { name: 'may', json: twoB('052025', { b2ba: [{ ctin: S, inv: [inv('A-9R', 900, { oinum: 'A-9' })] }] }) },
    ],
    cy,
    py,
    ny,
    [],
  );
  assert.equal(sum8A(rows), 900);
  // Input Reco: the superseded original is not a document "in 2B, not in books"
  const raw = rows.map(x => [x.key || null, ...x.amt, ...x.info]);
  assert.equal(raw.filter(isSuperseded).length, 1);
  const bk = ['10-04-2025', 'A-9R', 'S', S, 'B2B', 'Inputs', 'Availed', cy, '', '', null, 18, 900, 81, 81, 0, 0, ''];
  const { rows: out } = reconPurchases([bk], raw, cy, ny, new Map());
  assert.deepEqual(
    out.map((r: any[]) => r[0]),
    ['Matched'],
  );
}
// amendment without its original loaded: not counted (the original was counted in its own month), with a warning
{
  const w: Warn[] = [];
  const { rows } = process2b(
    [{ name: 'jun', json: twoB('062025', { b2ba: [{ ctin: S, inv: [inv('A-5', 1200, { oinum: 'A-5' })] }] }) }],
    cy,
    py,
    ny,
    w,
  );
  assert.equal(sum8A(rows), 0);
  assert.ok(w.some(x => /original document is not in the 2B files loaded/.test(x.msg)));
}
// credit note amended twice: the last one counts
{
  const nt = (ntnum: string, txval: number, extra: any = {}) => ({ ...inv('', txval, extra), ntnum, typ: 'C', dt: '20-04-2025' });
  const { rows } = process2b(
    [
      { name: 'apr', json: twoB('042025', { cdnr: [{ ctin: S, nt: [nt('C-1', 100)] }] }) },
      { name: 'may', json: twoB('052025', { cdnra: [{ ctin: S, nt: [nt('C-1', 150, { ontnum: 'C-1' })] }] }) },
      { name: 'jun', json: twoB('062025', { cdnra: [{ ctin: S, nt: [nt('C-1', 120, { ontnum: 'C-1' })] }] }) },
    ],
    cy,
    py,
    ny,
    [],
  );
  assert.equal(sum8A(rows), -120);
}

// ---- the same 2B / GSTR-1 period loaded twice: the second copy is left out; parts of a large return are kept
{
  const w: Warn[] = [];
  const a = twoB('042025', { b2b: [{ ctin: S, inv: [inv('A-1', 1000), inv('A-2', 500)] }] });
  const part2 = twoB('042025', { b2b: [{ ctin: S, inv: [inv('A-3', 300)] }] });
  const { rows } = process2b(
    [
      { name: 'a', json: a },
      { name: 'a copy', json: structuredClone(a) },
      { name: 'part 2', json: part2 },
    ],
    cy,
    py,
    ny,
    w,
  );
  assert.equal(sum8A(rows), 1800);
  assert.equal(w.filter(x => x.level === 'error' && /already loaded/.test(x.msg)).length, 1);
}
{
  const g1 = {
    gstin: '27AAACD0000A1Z9',
    fp: '042025',
    b2b: [
      {
        ctin: S,
        inv: [{ inum: 'S-1', idt: '05-04-2025', pos: '27', itms: [{ itm_det: { txval: 1000, rt: 18, camt: 90, samt: 90 } }] }],
      },
    ],
  };
  const w: Warn[] = [];
  const once = processGstr1([{ name: 'g1', json: g1 }], cy, ny, []).rows;
  const twice = processGstr1(
    [
      { name: 'g1', json: g1 },
      { name: 'g1 (1)', json: structuredClone(g1) },
    ],
    cy,
    ny,
    w,
  ).rows;
  const t = (rs: any[]) => rs.filter(r => r.key === '4A').reduce((a, r) => a + r.amt[0], 0);
  assert.equal(t(twice), t(once));
  assert.ok(w.some(x => x.level === 'error' && /already loaded/.test(x.msg)));
}

// ---- purchase debit note / return reduces ITC, and is not paired with the invoice of the same number
{
  const rows = [
    ['Date', 'Supplier Invoice No', 'Supplier', 'GSTIN', 'Voucher Type', 'Taxable Value', 'CGST', 'SGST'],
    ['01-05-2025', 'A1', 'Sup', S, 'Purchase', 10000, 900, 900],
    ['10-05-2025', 'DN1', 'Sup', S, 'Debit Note', -2000, -180, -180],
    ['12-05-2025', 'DN2', 'Sup', S, 'Debit Note', 1000, 90, 90],
  ];
  const map = { date: 0, no: 1, party: 2, gstin: 3, dtype: 4, taxable: 5, cgst: 6, sgst: 7 };
  const d = { noGstinZeroTax: 'NIL', itcType: 'Inputs', status: 'Availed', claimFy: cy, purNoGstin: 'RCM-UNREG' };
  const n = normalise('purchases', rows, 0, map, d, cy);
  assert.deepEqual(
    n.rows.map(r => [r[1], r[12], r[13]]),
    [
      ['A1', 10000, 900],
      ['DN1', -2000, -180],
      ['DN2', -1000, -90],
    ],
  );
  assert.equal(n.stats.credit, 2);
  // books invoice 7 and debit note 7 vs 2B invoice 7 and credit note 7
  const b = (no: string, t: number) => [
    '01-05-2025',
    no,
    'Sup',
    S,
    'B2B',
    'Inputs',
    'Availed',
    cy,
    '',
    '',
    null,
    18,
    t,
    t * 0.09,
    t * 0.09,
    0,
    0,
    '',
  ];
  const r = (no: string, t: number, dt: string, sec: string) => [
    '8A',
    t,
    t * 0.09,
    t * 0.09,
    0,
    0,
    '2B',
    cy,
    '052025',
    sec,
    S,
    'Sup',
    dt,
    no,
    '01-05-2025',
    cy,
    'Y',
    'N',
    '',
  ];
  const { rows: out } = reconPurchases(
    [b('7', 1000), b('7', -200)],
    [r('7', 1000, 'INV', 'B2B'), r('7', -200, 'CN', 'CDNR')],
    cy,
    ny,
    new Map(),
  );
  assert.deepEqual(out.map((x: any[]) => [x[0], x[8]]).sort(), [
    ['Matched', -200],
    ['Matched', 1000],
  ]);
}

// ---- CSV: invoice numbers keep their leading zeros; dates written month/day are recognised
(async () => {
  const t = await readTable('a.csv', new TextEncoder().encode('Inv No,Amt,Date\n000123,10,04/13/2025\n0,5,04/01/2025\n'));
  assert.equal(t[0].rows[1][0], '000123');
  assert.equal(t[0].rows[1][1], 10);
  assert.equal(t[0].rows[2][0], 0);
  assert.equal(dateOrder(['04/13/2025', '04/01/2025']), 'mdy');
  assert.equal(dateOrder(['13/04/2025', '04/01/2025']), 'dmy');
  assert.equal(dateOrder(['04/01/2025']), 'dmy', 'ambiguous: day first (India)');
  const n = normalise(
    'sales',
    [
      ['Date', 'No', 'Taxable', 'IGST'],
      ['04/13/2025', '1', 100, 18],
      ['05/02/2025', '2', 100, 18],
    ],
    0,
    { date: 0, no: 1, taxable: 2, igst: 3 },
    { noGstinZeroTax: 'NIL', itcType: 'Inputs', status: 'Availed', claimFy: cy, purNoGstin: 'B2B' },
    cy,
  );
  assert.equal(n.stats.dateFmt, 'mdy');
  assert.equal(n.stats.outsideFy, 0);

  // ---- matching speed: 30,000 documents a side in well under the old ~6 s
  const N = 30000,
    books: any[][] = [],
    ret: any[][] = [];
  for (let i = 0; i < N; i++) {
    const g = `29ABCDE${String(i % 800).padStart(4, '0')}F1Z5`,
      tx = 1000 + (i % 97) * 10;
    books.push([
      '01-05-2025',
      `INV/${i}`,
      'S',
      g,
      'B2B',
      'Inputs',
      'Availed',
      cy,
      '',
      '',
      null,
      18,
      tx,
      tx * 0.09,
      tx * 0.09,
      0,
      0,
      '',
    ]);
    ret.push([
      '8A',
      tx,
      tx * 0.09,
      tx * 0.09,
      0,
      0,
      '2B',
      cy,
      '052025',
      'B2B',
      g,
      'S',
      'INV',
      i % 11 ? `INV/${i}` : `X${i}`,
      '01-05-2025',
      cy,
      'Y',
      'N',
      '',
    ]);
  }
  const t0 = Date.now();
  reconPurchases(books, ret, cy, ny, new Map());
  assert.ok(Date.now() - t0 < 4000, `30k documents took ${Date.now() - t0} ms`);
  console.log('review fixes: ok');
})().catch(e => {
  console.error(e);
  process.exit(1);
});

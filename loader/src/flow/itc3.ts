// ITC three ways: booked in the books, available in GSTR-2B, claimed in GSTR-3B (credit ledger).
// Books → 2B is split invoice by invoice with the Input Reco matching (recon.ts); 2B → 3B month by month, since 3B is a total.
// Invoices come from the purchase vouchers in the GL dump; a purchase register, when loaded, is a cross-check of their total.
import { HEADS, Head, Ym } from './types';
import { BEntry, Inputs, Result, ymLabel } from './engine';
import { contraOf } from './books';
import { DEFAULT_MATCH, MatchOptions, reconPurchases, toDate } from '../recon';
import { fyShift } from '../util';

export type H4 = Record<Head, number>;
export interface BridgeLine {
  label: string;
  amt: H4;
  kind: 'start' | 'less' | 'add' | 'total' | 'check';
  action?: string;
  status?: string;
}
export interface Itc3Invoice {
  status: string;
  supplier: string;
  gstin: string;
  bNo: string;
  rNo: string;
  bDate: string;
  rDate: string;
  bTax: number;
  rTax: number;
  period: string;
  note: string;
}
export interface Itc3 {
  books: BridgeLine[];
  claim: BridgeLine[];
  months: { m: Ym; booked: number; b2b: number; claimed: number; diff: number; run: number }[];
  invoices: Itc3Invoice[];
  nonInvoice: { date: string; voucher: string; account: string; party: string; tax: number; id: string }[];
  register?: { register: number; books: number };
  notes: string[];
}
const z4 = (): H4 => ({ IGST: 0, CGST: 0, SGST: 0, CESS: 0 });
const r2 = (x: number) => Math.round(x * 100) / 100;
const t4 = (a: H4) => r2(HEADS.reduce((s, h) => s + a[h], 0));
const add = (a: H4, b: H4, k = 1) => {
  for (const h of HEADS) a[h] = r2(a[h] + b[h] * k);
};
const ymOfPeriod = (p: string) => (/^\d{6}$/.test(p) ? `${p.slice(2)}-${p.slice(0, 2)}` : '');
const ymOfDd = (d: string) => (/^\d{2}-\d{2}-\d{4}$/.test(d) ? `${d.slice(6)}-${d.slice(3, 5)}` : '');

/** raw2b: RAW_2B rows (as build.ts writes them); register: normalised Books_Purchase rows */
export function itcThreeWay(
  r: Result,
  inp: Inputs,
  raw2b: any[][],
  register?: any[][],
  match: MatchOptions = DEFAULT_MATCH,
): Itc3 {
  const inP = (m: Ym) => m >= inp.from && m <= inp.to;
  const notes: string[] = [];
  // ---- books: the ITC-availed lines, grouped into vouchers
  const avail = r.books.filter(e => e.where === 'in' && e.bucket === 'itc_availed' && HEADS.includes(e.head as Head));
  const byV = new Map<string, BEntry[]>();
  for (const e of avail) {
    const k = `${e.line.file}|${e.line.vkey}`;
    (byV.get(k) || byV.set(k, []).get(k)!).push(e);
  }
  const allLines = new Map<string, Inputs['books'][number]['lines']>();
  for (const f of inp.books)
    for (const l of f.lines) {
      const k = `${f.name}|${l.vkey}`;
      (allLines.get(k) || allLines.set(k, []).get(k)!).push(l);
    }
  const gstAcc = new Set(Object.keys(inp.accMap));
  const bookRows: any[][] = [],
    nonInvoice: Itc3['nonInvoice'] = [];
  const booked = z4(),
    invTax = z4();
  for (const [k, es] of byV) {
    const tax = z4();
    for (const e of es) tax[e.head as Head] = r2(tax[e.head as Head] + e.effect);
    add(booked, tax);
    const lines = allLines.get(k) || [];
    // the supplier: a party account on the other side of the voucher (credited for a purchase, debited for a debit note)
    const sign = t4(tax) >= 0 ? 1 : -1;
    const party = lines
      .filter(l => !gstAcc.has(l.account) && contraOf(l.account, l.accGroup) === 'party' && (sign > 0 ? l.cr > 0 : l.dr > 0))
      .sort((a, b) => (sign > 0 ? b.cr - a.cr : b.dr - a.dr))[0];
    const l0 = es[0].line;
    if (!party && !l0.party) {
      nonInvoice.push({
        date: l0.date,
        voucher: `${l0.vchType} ${l0.vchNo}`.trim(),
        account: [...new Set(es.map(e => e.line.account))].join(', '),
        party: '',
        tax: t4(tax),
        id: es[0].id,
      });
      continue;
    }
    add(invTax, tax);
    const gross = party ? (sign > 0 ? party.cr : party.dr) : 0;
    const taxable = r2(Math.max(0, gross - Math.abs(t4(tax))) * sign);
    const ref = lines.find(l => l.ref)?.ref || l0.vchNo;
    bookRows.push([
      l0.date,
      ref,
      party?.account || l0.party,
      '',
      'B2B',
      'Inputs',
      'Availed',
      inp.fy,
      null,
      null,
      null,
      null,
      taxable,
      tax.CGST,
      tax.SGST,
      tax.IGST,
      tax.CESS,
    ]);
  }
  // ---- 2B of the period
  const ny = fyShift(inp.fy, 1);
  const pool = raw2b.filter(x => x && inP(ymOfPeriod(String(x[8]))));
  const b2bTot = z4();
  const b2bByM = new Map<Ym, number>();
  for (const x of pool) {
    if (
      !['B2B', 'B2BA', 'CDNR', 'CDNRA', 'IMPG', 'IMPGSEZ', 'ISD', 'ISDA'].includes(String(x[9])) ||
      String(x[16]) === 'N' ||
      String(x[17]) === 'Y'
    )
      continue;
    const a: H4 = { CGST: +x[2] || 0, SGST: +x[3] || 0, IGST: +x[4] || 0, CESS: +x[5] || 0 };
    add(b2bTot, a);
    const m = ymOfPeriod(String(x[8]));
    b2bByM.set(m, r2((b2bByM.get(m) || 0) + t4(a)));
  }
  // ---- match invoices (books period vs 2B period, both sides restricted to the period)
  const rec = reconPurchases(bookRows, pool, inp.fy, ny, new Map(), new Map(), notes, match);
  const matched = z4(),
    matched2b = z4(),
    bOnly = z4(),
    rOnly = z4();
  const invoices: Itc3Invoice[] = [];
  for (const row of rec.rows) {
    const st = String(row[0]);
    const B: H4 = { CGST: +row[9] || 0, SGST: +row[10] || 0, IGST: +row[11] || 0, CESS: +row[12] || 0 };
    const R: H4 = { CGST: +row[14] || 0, SGST: +row[15] || 0, IGST: +row[16] || 0, CESS: +row[17] || 0 };
    if (st.startsWith('Matched')) {
      add(matched, B);
      add(matched2b, R);
    } else if (st === 'In books, not in 2B') add(bOnly, B);
    else if (st === 'In 2B, not in books') add(rOnly, R);
    else continue;
    // a matched invoice in a different month is not a difference: list only what needs looking at
    const diff = Math.abs(t4(B) - t4(R)) > match.tol;
    if (st.startsWith('Matched') && !diff) continue;
    invoices.push({
      status: st.startsWith('Matched') ? (diff ? 'Matched, tax differs' : st) : st,
      supplier: String(row[3] || ''),
      gstin: String(row[2] || ''),
      bNo: String(row[4] || ''),
      rNo: String(row[5] || ''),
      bDate: String(row[6] || ''),
      rDate: String(row[7] || ''),
      bTax: t4(B),
      rTax: t4(R),
      period: String(row[20] || ''),
      note: String(row[25] || ''),
    });
  }
  const ord = ['In books, not in 2B', 'In 2B, not in books', 'Matched, tax differs'];
  invoices.sort(
    (a, b) =>
      (ord.indexOf(a.status) + 1 || 9) - (ord.indexOf(b.status) + 1 || 9) ||
      Math.abs(b.bTax || b.rTax) - Math.abs(a.bTax || a.rTax),
  );
  const nonInv = z4();
  for (const h of HEADS) nonInv[h] = r2(booked[h] - invTax[h]);
  const vdiff = z4();
  for (const h of HEADS) vdiff[h] = r2(matched2b[h] - matched[h]);
  const computed = z4();
  add(computed, booked);
  add(computed, nonInv, -1);
  add(computed, bOnly, -1);
  add(computed, rOnly);
  add(computed, vdiff);
  const other = z4();
  for (const h of HEADS) other[h] = r2(b2bTot[h] - computed[h]);
  const books: BridgeLine[] = [
    { label: 'ITC availed in books (purchase invoices, debit notes and other entries)', amt: booked, kind: 'start' },
    {
      label: 'Less: entries without a purchase invoice (journals, adjustments)',
      amt: neg(nonInv),
      kind: 'less',
      action: 'Check each entry: ITC needs a tax invoice (they are listed below).',
    },
    {
      label: 'Less: invoices in books, not in 2B',
      amt: neg(bOnly),
      kind: 'less',
      status: 'In books, not in 2B',
      action: 'Not claimable until the supplier files: follow up, or reverse the ITC.',
    },
    {
      label: 'Add: invoices in 2B, not in books',
      amt: rOnly,
      kind: 'add',
      status: 'In 2B, not in books',
      action: 'Book the invoice and claim it (by 30 November of the next year), or confirm it is not yours.',
    },
    {
      label: 'Add / less: tax differs on matched invoices (2B − books)',
      amt: vdiff,
      kind: 'add',
      status: 'Matched, tax differs',
      action: 'Correct the books, or ask the supplier to amend.',
    },
    ...(HEADS.some(h => Math.abs(other[h]) > 0.5)
      ? [
          {
            label: 'Add: other 2B entries (ISD, imports, not matched by type)',
            amt: other,
            kind: 'add' as const,
            action: 'Check them in the Input Reco tab.',
          },
        ]
      : []),
    { label: 'ITC available as per GSTR-2B for the period', amt: b2bTot, kind: 'total' },
  ];
  // ---- 2B → 3B, month by month (3B is a total, so the comparison is by month)
  const claimedByM = new Map<Ym, number>(),
    bookedByM = new Map<Ym, number>();
  const claimed = z4();
  for (const p of r.portal)
    if (p.where === 'in' && p.block === 'itc' && ['itc_availed', 'itc_rcm', 'itc_reclaimed', 'itc_reversed'].includes(p.bucket)) {
      const v = p.effect; // the change in the ITC balance (reversals are negative)
      claimed[p.head] = r2(claimed[p.head] + v);
      claimedByM.set(p.month, r2((claimedByM.get(p.month) || 0) + v));
    }
  for (const e of avail) bookedByM.set(e.month, r2((bookedByM.get(e.month) || 0) + e.effect));
  let run = 0;
  const months = r.months.filter(inP).map(m => {
    const b2b = b2bByM.get(m) || 0,
      claimed = claimedByM.get(m) || 0,
      diff = r2(b2b - claimed);
    run = r2(run + diff);
    return { m, booked: r2(bookedByM.get(m) || 0), b2b, claimed, diff, run };
  });
  const gap = z4();
  for (const h of HEADS) gap[h] = r2(b2bTot[h] - claimed[h]);
  const rcm = r.books.filter(e => e.where === 'in' && e.bucket === 'itc_rcm').reduce((s, e) => s + e.effect, 0);
  const claim: BridgeLine[] = [
    { label: 'ITC available as per GSTR-2B for the period', amt: b2bTot, kind: 'start' },
    { label: 'ITC claimed in GSTR-3B (credit ledger: ITC accrued, net of reversals)', amt: claimed, kind: 'total' },
    {
      label: 'Difference: available in 2B less claimed in 3B',
      amt: gap,
      kind: 'check',
      action:
        'The months below show where it arises. A month above zero has 2B credit not claimed yet (claim it by 30 November of the next year, unless it is ineligible or reversed); below zero, ITC was claimed without a 2B invoice of that month, or 2B of another month was claimed.',
    },
  ];
  if (Math.abs(rcm) > 0.5)
    notes.push(`The 3B claim includes ITC on reverse charge (${r2(rcm).toLocaleString('en-IN')} in books), which is not in 2B.`);
  const reversed = r.books.filter(e => e.where === 'in' && e.bucket === 'itc_reversed').reduce((s, e) => s - e.effect, 0);
  if (Math.abs(reversed) > 0.5)
    notes.push(
      `The credit ledger shows ITC net of reversals; the books reversed ${r2(reversed).toLocaleString('en-IN')} in the period.`,
    );
  // ---- purchase register: total tax of the period against the invoices taken from the GL
  let regCheck: Itc3['register'];
  if (register?.length) {
    const reg = register
      .filter(x => x && inP(ymOfDd(toDate(x[0]))))
      .reduce((s, x) => s + (+x[13] || 0) + (+x[14] || 0) + (+x[15] || 0) + (+x[16] || 0), 0);
    regCheck = { register: r2(reg), books: t4(invTax) };
  }
  // the books carry no supplier GSTIN: say how it was found, without the Input Reco sheet names
  const out = notes.map(n =>
    /had no supplier GSTIN in books/.test(n)
      ? n
          .replace(/ See "Match basis".*$/, '')
          .replace(
            /^(\d+) purchase invoice\(s\) had no supplier GSTIN in books; /,
            'The books do not carry supplier GSTINs: of $1 purchase invoices, ',
          )
      : n,
  );
  return { books, claim, months, invoices, nonInvoice, register: regCheck, notes: out };
}
const neg = (a: H4): H4 => ({ IGST: -a.IGST, CGST: -a.CGST, SGST: -a.SGST, CESS: -a.CESS });
export { ymLabel };

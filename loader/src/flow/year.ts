// Year-level reconciliation: set-up checks, one comparison line per component (books vs portal for the period, with the
// month where a difference arises), suggested journals for set-off / payment differences, and the closing statement
// books + journals vs portal. Built on the engine's month-wise cells (engine.ts) so every figure ties to the lines.
import { AccType, Block, HEADS, INTERNAL, UNMAPPED, Ym, bucketOf } from './types';
import { HEAD_KEYS, HeadK, Inputs, Result, blockOfType, ymLabel } from './engine';
import { inr } from '../util';

export type Kind = 'invoice' | 'itc' | 'journal' | 'setup';
export interface Comp {
  id: string;
  block: Block;
  label: string;
  buckets: string[];
  gsign: 1 | -1;
  kind: Kind;
  books: Record<HeadK, number>;
  portal: Record<HeadK, number>;
  diff: Record<HeadK, number>;
  /** part of the difference that is a year-end cut-off (in books in the period, on the portal after it, or the other way) */
  cutoff: Record<HeadK, number>;
  rest: Record<HeadK, number>;
  months: { m: Ym; b: number; p: number; d: number; run: number; any: boolean }[];
  action: string;
}
export interface Leg {
  account: string;
  head: HeadK;
  dr: number;
  cr: number;
  block: Block | 'other';
  existing: boolean;
}
export interface Journal {
  no: string;
  date: string;
  title: string;
  narration: string;
  legs: Leg[];
  comps: string[];
}
export interface SetupIssue {
  level: 'fix' | 'check' | 'note';
  what: string;
  action: string;
}
export interface ClosingLine {
  label: string;
  amt: Record<HeadK, number>;
  action: string;
  kind: 'opening' | 'cutoff' | 'comp' | 'split' | 'unexplained';
}
export interface Closing {
  block: 'itc' | 'liab' | 'cash';
  books: Record<HeadK, number>;
  journals: Record<HeadK, number>;
  adjusted: Record<HeadK, number>;
  portal: Record<HeadK, number>;
  diff: Record<HeadK, number>;
  lines: ClosingLine[];
}
export interface Year {
  setup: SetupIssue[];
  comps: Comp[];
  journals: Journal[];
  closing: Closing[];
}

const zero = (): Record<HeadK, number> => ({ IGST: 0, CGST: 0, SGST: 0, CESS: 0, ALL: 0 });
const r2 = (x: number) => Math.round(x * 100) / 100;
const nz = (x: number) => Math.abs(x) > 0.5;
const tot = (a: Record<HeadK, number>) => r2(HEAD_KEYS.reduce((s, h) => s + a[h], 0));
const fmt = inr;
export const BLOCK_NAME: Record<Block, string> = {
  itc: 'Input tax credit',
  liab: 'Tax payable',
  cash: 'Electronic cash ledger',
  fee: 'Interest, late fee and penalty',
};

/** comparison lines: the portal shows ITC net of reversals and output tax net of credit notes, so those rows are compared together */
const DEFS: { id: string; block: Block; label: string; buckets: string[]; gsign: 1 | -1; kind: Kind; action: string }[] = [
  {
    id: 'itc_net',
    block: 'itc',
    label: 'ITC availed, net of reversals (credit ledger: ITC accrued)',
    buckets: ['itc_availed', 'itc_rcm', 'itc_reclaimed', 'itc_reversed'],
    gsign: 1,
    kind: 'itc',
    action: 'Split invoice by invoice against GSTR-2B (below).',
  },
  {
    id: 'itc_setoff',
    block: 'itc',
    label: 'ITC used to pay tax (set-off)',
    buckets: ['itc_setoff'],
    gsign: -1,
    kind: 'journal',
    action: 'Post the suggested set-off journal.',
  },
  {
    id: 'itc_refund',
    block: 'itc',
    label: 'ITC refunded',
    buckets: ['itc_refund'],
    gsign: -1,
    kind: 'journal',
    action: 'Post the suggested refund journal.',
  },
  {
    id: 'out_net',
    block: 'liab',
    label: 'Output tax, net of credit notes',
    buckets: ['out_tax', 'out_cn'],
    gsign: 1,
    kind: 'invoice',
    action: 'Match sales with GSTR-1 invoice by invoice (Annual return tab) for the months shown.',
  },
  {
    id: 'out_rcm',
    block: 'liab',
    label: 'Tax on reverse charge',
    buckets: ['out_rcm'],
    gsign: 1,
    kind: 'invoice',
    action: 'Check the reverse-charge purchases of the months shown against GSTR-3B 3.1(d).',
  },
  {
    id: 'paid_itc',
    block: 'liab',
    label: 'Tax paid by ITC (set-off)',
    buckets: ['paid_itc'],
    gsign: -1,
    kind: 'journal',
    action: 'Post the suggested set-off journal.',
  },
  {
    id: 'paid_cash',
    block: 'liab',
    label: 'Tax paid in cash',
    buckets: ['paid_cash'],
    gsign: -1,
    kind: 'journal',
    action: 'Post the suggested cash payment journal.',
  },
  {
    id: 'cash_deposit',
    block: 'cash',
    label: 'Deposits into the cash ledger (challans)',
    buckets: ['cash_deposit'],
    gsign: 1,
    kind: 'journal',
    action: 'Post the suggested deposit journal (confirm the bank).',
  },
  {
    id: 'cash_used',
    block: 'cash',
    label: 'Cash ledger used to pay tax, interest and fees',
    buckets: ['cash_used'],
    gsign: -1,
    kind: 'journal',
    action: 'Post the suggested cash payment journal.',
  },
  {
    id: 'cash_refund',
    block: 'cash',
    label: 'Cash ledger refunded',
    buckets: ['cash_refund'],
    gsign: -1,
    kind: 'journal',
    action: 'Post the suggested refund journal (confirm the bank).',
  },
  {
    id: 'fee_interest',
    block: 'fee',
    label: 'Interest',
    buckets: ['fee_interest'],
    gsign: 1,
    kind: 'journal',
    action: 'Post the suggested cash payment journal.',
  },
  {
    id: 'fee_latefee',
    block: 'fee',
    label: 'Late fee',
    buckets: ['fee_latefee'],
    gsign: 1,
    kind: 'journal',
    action: 'Post the suggested cash payment journal.',
  },
  {
    id: 'fee_penalty',
    block: 'fee',
    label: 'Penalty and others',
    buckets: ['fee_penalty'],
    gsign: 1,
    kind: 'journal',
    action: 'Post the suggested cash payment journal.',
  },
];
const JGROUPS: { title: string; comps: string[]; balance: string; why: string }[] = [
  {
    title: 'ITC set-off as per the GST portal',
    comps: ['itc_setoff', 'paid_itc'],
    balance: 'Suspense (check the set-off)',
    why: 'set-off of ITC against tax as per the credit ledger and liability register',
  },
  {
    title: 'Tax, interest and fees paid in cash as per the GST portal',
    comps: ['paid_cash', 'cash_used', 'fee_interest', 'fee_latefee', 'fee_penalty'],
    balance: 'Suspense (check the payment)',
    why: 'tax, interest and fees paid from the electronic cash ledger as per the portal',
  },
  {
    title: 'Cash ledger deposits and refunds as per the GST portal',
    comps: ['cash_deposit', 'cash_refund'],
    balance: 'Bank (confirm the account)',
    why: 'challans deposited into / refunds from the electronic cash ledger as per the portal',
  },
  {
    title: 'ITC refund as per the GST portal',
    comps: ['itc_refund'],
    balance: 'GST refund receivable',
    why: 'ITC refunded as per the credit ledger',
  },
];
const signOf = (k: string) => bucketOf(k)?.sign ?? 1;

export function yearRecon(r: Result, inp: Inputs, opts: { has2b: boolean }): Year {
  const months = r.months.filter(m => m >= inp.from && m <= inp.to);
  // ---- components from the month-wise cells (same heads as the engine compared: accounts not split by head are in ALL)
  const cell = new Map<string, { b: number; p: number }>();
  for (const x of r.monthly) {
    if (!months.includes(x.month)) continue;
    const k = `${x.bucket}|${x.head}|${x.month}`;
    const c = cell.get(k) || { b: 0, p: 0 };
    c.b += x.b;
    c.p += x.p;
    cell.set(k, c);
  }
  const comps: Comp[] = [];
  const defs = [...DEFS];
  for (const blk of ['itc', 'liab', 'cash'] as Block[])
    for (const extra of [UNMAPPED, INTERNAL])
      defs.push({
        id: `${blk}:${extra}`,
        block: blk,
        label: extra === UNMAPPED ? 'Books lines not placed in a flow row' : 'Transfers between GST accounts in books',
        buckets: [`${blk}:${extra}`],
        gsign: 1,
        kind: 'setup',
        action:
          extra === UNMAPPED
            ? 'Pick a flow row for each line (Show lines).'
            : 'Usually nil over the year; check the lines if not.',
      });
  for (const d of defs) {
    const c: Comp = { ...d, books: zero(), portal: zero(), diff: zero(), cutoff: zero(), rest: zero(), months: [] };
    for (const m of months) {
      let b = 0,
        p = 0;
      const dh = zero();
      for (const k of d.buckets)
        for (const h of HEAD_KEYS) {
          const x = cell.get(`${k}|${h}|${m}`);
          if (!x) continue;
          const s = d.gsign * signOf(k);
          c.books[h] += x.b * s;
          c.portal[h] += x.p * s;
          b += x.b * s;
          p += x.p * s;
          dh[h] += (x.b - x.p) * s;
        }
      c.months.push({ m, b: r2(b), p: r2(p), d: r2(b - p), run: 0, any: HEAD_KEYS.some(h => nz(dh[h])) });
    }
    let run = 0;
    for (const x of c.months) {
      run = r2(run + x.d);
      x.run = run;
    }
    for (const it of r.items)
      if (it.category === 'After the period' && d.buckets.includes(it.bucket))
        for (const h of HEAD_KEYS) c.cutoff[h] += it.amt[h] * d.gsign * signOf(it.bucket);
    for (const h of HEAD_KEYS) {
      c.books[h] = r2(c.books[h]);
      c.portal[h] = r2(c.portal[h]);
      c.diff[h] = r2(c.books[h] - c.portal[h]);
      c.cutoff[h] = r2(c.cutoff[h]);
      c.rest[h] = r2(c.diff[h] - c.cutoff[h]);
    }
    if (d.kind === 'setup' && !HEAD_KEYS.some(h => nz(c.books[h]) || nz(c.portal[h]))) continue;
    if (d.block === 'fee' && !HEAD_KEYS.some(h => nz(c.books[h]) || nz(c.portal[h]))) continue;
    comps.push(c);
  }

  // ---- suggested journals: what the books need so the set-off / payment components agree with the portal
  const accOf = (t: AccType[], h: HeadK): [string, boolean] => {
    const hit = Object.entries(inp.accMap)
      .filter(([, m]) => t.includes(m.type) && (m.head === h || m.head === 'ALL' || !m.head))
      .sort((a, b) => +(a[1].head !== h) - +(b[1].head !== h) || a[0].localeCompare(b[0]));
    if (hit.length) return [hit[0][0], true];
    const g =
      t[0] === 'input'
        ? `Input ${h}`
        : t[0] === 'output'
          ? `Output ${h}`
          : t[0] === 'cash_ledger'
            ? 'GST electronic cash ledger'
            : t[0] === 'interest'
              ? 'Interest on GST'
              : t[0] === 'late_fee'
                ? 'GST late fee'
                : 'GST penalty';
    return [h === 'ALL' || t[0] === 'cash_ledger' || /fee|interest|penalty/.test(t[0]) ? g : g, false];
  };
  const typesOf = (c: Comp): AccType[] =>
    c.block === 'itc'
      ? ['input']
      : c.block === 'liab'
        ? ['output', 'payable', 'rcm_payable']
        : c.block === 'cash'
          ? ['cash_ledger']
          : c.id === 'fee_interest'
            ? ['interest']
            : c.id === 'fee_latefee'
              ? ['late_fee']
              : ['penalty'];
  const journals: Journal[] = [];
  const endDate = (() => {
    const [y, m] = inp.to.split('-').map(Number);
    const d = new Date(Date.UTC(y, m, 0));
    return `${String(d.getUTCDate()).padStart(2, '0')}-${String(m).padStart(2, '0')}-${y}`;
  })();
  for (const g of JGROUPS) {
    const legs = new Map<string, Leg>();
    const ms = new Set<string>();
    for (const id of g.comps) {
      const c = comps.find(x => x.id === id);
      if (!c) continue;
      for (const h of HEAD_KEYS) {
        const adj = -c.rest[h]; // portal − books, in the component's own direction
        if (!nz(adj)) continue;
        for (const x of c.months) if (x.any) ms.add(x.m);
        // effect on the account: assets (ITC, cash ledger) and expenses are debited when they go up; liabilities are credited
        const up = c.block === 'fee' ? adj : adj * c.gsign; // change in the block balance
        const [acc, existing] = accOf(typesOf(c), c.block === 'fee' ? 'ALL' : h);
        const debit = c.block === 'liab' ? up < 0 : up > 0;
        const lh: HeadK = c.block === 'fee' ? 'ALL' : h;
        const k = `${acc}|${lh}`,
          l = legs.get(k) || { account: acc, head: lh, dr: 0, cr: 0, block: c.block, existing };
        if (debit) l.dr += Math.abs(up);
        else l.cr += Math.abs(up);
        legs.set(k, l);
      }
    }
    const L = [...legs.values()]
      .map(l => {
        const n = r2(l.dr - l.cr);
        return { ...l, dr: n > 0 ? n : 0, cr: n < 0 ? -n : 0 };
      })
      .filter(l => nz(l.dr) || nz(l.cr));
    if (!L.length) continue;
    const off = r2(L.reduce((s, l) => s + l.dr - l.cr, 0));
    if (nz(off))
      L.push({ account: g.balance, head: 'ALL', dr: off < 0 ? -off : 0, cr: off > 0 ? off : 0, block: 'other', existing: false });
    L.sort((a, b) => b.dr - a.dr || a.cr - b.cr);
    const mlist = [...ms].sort();
    journals.push({
      no: `J${journals.length + 1}`,
      date: endDate,
      title: g.title,
      comps: g.comps.filter(id => comps.some(c => c.id === id)),
      narration: `Books brought in line with the GST portal: ${g.why}, ${ymLabel(inp.from)} to ${ymLabel(inp.to)}${mlist.length ? ` (differences in ${mlist.length > 4 ? `${mlist.length} months, ${ymLabel(mlist[0])} to ${ymLabel(mlist[mlist.length - 1])}` : mlist.map(ymLabel).join(', ')})` : ''}.`,
      legs: L,
    });
  }

  // ---- closing statement: books closing + journals = adjusted; what still differs from the portal, line by line
  const closing: Closing[] = (['itc', 'liab', 'cash'] as const).map(b => {
    const books = r.flowBooks.closing[b],
      portal = r.flowPortal.closing[b];
    const journalsE = zero();
    for (const j of journals)
      for (const l of j.legs) if (l.block === b) journalsE[l.head] += b === 'liab' ? l.cr - l.dr : l.dr - l.cr;
    const adjusted = zero(),
      diff = zero();
    for (const h of HEAD_KEYS) {
      journalsE[h] = r2(journalsE[h]);
      adjusted[h] = r2(books[h] + journalsE[h]);
      diff[h] = r2(adjusted[h] - portal[h]);
    }
    const lines: ClosingLine[] = [];
    const op = zero();
    for (const h of HEAD_KEYS) op[h] = r2(r.flowBooks.opening[b][h] - r.flowPortal.opening[b][h]);
    if (HEAD_KEYS.some(h => nz(op[h])))
      lines.push({
        label: 'Opening balance: books and portal differ',
        amt: op,
        kind: 'opening',
        action: 'Check the openings under Balances; a difference carried from last year needs its own reconciliation.',
      });
    const cut = zero();
    for (const c of comps.filter(c => c.block === b)) for (const h of HEAD_KEYS) cut[h] += c.cutoff[h] * c.gsign;
    if (HEAD_KEYS.some(h => nz(cut[h])))
      lines.push({
        label: 'Year-end cut-off: in books in the period, on the portal in a later return (or the other way)',
        amt: Object.fromEntries(HEAD_KEYS.map(h => [h, r2(cut[h])])) as any,
        kind: 'cutoff',
        action: 'None: clears in the next period.',
      });
    for (const c of comps.filter(c => c.block === b && c.kind !== 'journal')) {
      const a = Object.fromEntries(HEAD_KEYS.map(h => [h, r2(c.rest[h] * c.gsign)])) as Record<HeadK, number>;
      if (HEAD_KEYS.some(h => nz(a[h])))
        lines.push({
          label: c.label,
          amt: a,
          kind: 'comp',
          action: c.kind === 'itc' && !opts.has2b ? 'Load the GSTR-2B files to split this invoice by invoice.' : c.action,
        });
    }
    const left = zero();
    for (const h of HEAD_KEYS) left[h] = r2(diff[h] - lines.reduce((s, l) => s + l.amt[h], 0));
    if (HEAD_KEYS.some(h => nz(left[h]))) {
      const split = !nz(tot(left));
      lines.push({
        label: split ? 'Books account not split by tax head (the totals agree)' : 'Not explained',
        amt: left,
        kind: split ? 'split' : 'unexplained',
        action: split ? 'None: compare the Total column.' : 'Check the opening balances and the lines not placed in a flow row.',
      });
    }
    return { block: b, books, journals: journalsE, adjusted, portal, diff, lines };
  });

  // ---- set-up checks
  const setup: SetupIssue[] = [];
  const kinds = new Set(inp.portal.map(p => p.ledger.kind));
  const missing = (['credit', 'cash', 'liability'] as const).filter(k => !kinds.has(k));
  if (missing.length)
    setup.push({
      level: 'fix',
      what: `Portal ledger(s) not loaded: ${missing.map(k => (k === 'credit' ? 'Electronic Credit Ledger' : k === 'cash' ? 'Electronic Cash Ledger' : 'Electronic Liability Register')).join(', ')}`,
      action: 'Download them for the whole year and add them under Add files.',
    });
  const fileOpen = new Set(
    inp.books.filter(f => f.format === 'ledger' && f.opening !== undefined && f.ledger).map(f => f.ledger!),
  );
  // an account with no opening, where the portal has an opening balance for its ledger and head
  const noOpen = Object.entries(inp.accMap)
    .filter(([n, m]) => {
      const blk = blockOfType(m.type);
      if (!blk || blk === 'fee' || inp.bal.books[n]?.opening !== undefined || fileOpen.has(n)) return false;
      const po = r.flowPortal.opening[blk];
      return !m.head || m.head === 'ALL' ? nz(tot(po)) : nz(po[m.head as HeadK]);
    })
    .map(([n]) => n);
  if (noOpen.length)
    setup.push({
      level: 'fix',
      what: `No opening balance for ${noOpen.length} books account(s): ${noOpen.slice(0, 6).join(', ')}${noOpen.length > 6 ? ', …' : ''}. They start at nil while the portal has an opening balance.`,
      action: 'Enter the openings under Balances, from the trial balance of the day before the period.',
    });
  for (const b of ['itc', 'liab', 'cash'] as const) {
    const u = r.books.filter(e => e.where === 'in' && e.block === b && e.bucket === UNMAPPED);
    if (u.length)
      setup.push({
        level: 'fix',
        what: `${u.length} books line(s) of ${BLOCK_NAME[b].toLowerCase()} could not be placed in a flow row (total ${fmt(r2(u.reduce((s, e) => s + e.effect, 0)))})`,
        action: 'Pick a flow row for each: GST flow → Books lines not placed (Show lines), or Change on the line.',
      });
  }
  for (const w of r.warns)
    if (w.level !== 'info')
      setup.push({ level: 'check', what: w.msg, action: 'Check the files: a missing month shows up as a difference.' });
  if (!opts.has2b)
    setup.push({
      level: 'note',
      what: 'No GSTR-2B files loaded',
      action: 'Add the 2B JSONs of the period to split the ITC difference invoice by invoice.',
    });
  void HEADS;
  return { setup, comps, journals, closing };
}

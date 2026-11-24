// GST ledger flow workbook: detail sheets carry every input row with its classification; the flow statements,
// differences and checks are SUMIFS formulas over them (cached values are written too, Excel recalculates on open).
import { C, SheetSpec, col, serial, writeBook } from '../xlsxbook';
import { SheetInfo, Section, heading, how, indexSheet, instructionsSheet } from '../wpfront';
import { LOADER_VERSION } from '../version';
import { BUCKETS, BLOCK_LABEL, Block, INTERNAL, PORTAL_LABEL, UNMAPPED, ACC_TYPE_LABEL, bucketOf } from './types';
import { HEAD_KEYS, HeadK, Result, bucketLabel, ymLabel } from './engine';
import { FLOW_GUIDE } from './guide';
import { BLOCK_NAME, Year } from './year';
import { Itc3 } from './itc3';

export interface AuditRow {
  when: string;
  change: string;
  record: string;
  from: string;
  to: string;
  remark: string;
}
export interface FlowMeta {
  gstin: string;
  legalName: string;
  fy: string;
  from: string;
  to: string;
  files: { name: string; kind: string; note: string }[];
  audit: AuditRow[];
  builtAt: string;
  year: Year;
  itc: Itc3 | null;
  /** the Matching settings in use, in words */
  settings?: string[];
  /** notes the preparer wrote on the steps */
  notes?: [string, string][];
}

const HEADS_TXT: Record<HeadK, string> = { IGST: 'IGST', CGST: 'CGST', SGST: 'SGST', CESS: 'Cess', ALL: 'Not split by head' };
const BLOCKS: Block[] = ['itc', 'liab', 'cash', 'fee'];
const N = (v: number) => Math.round(v * 100) / 100;

export async function flowWorkbook(r: Result, meta: FlowMeta): Promise<Uint8Array> {
  // short references for every record
  const bref = new Map<string, string>(),
    pref = new Map<string, string>();
  r.books.forEach((e, i) => bref.set(e.id, `BL${String(i + 1).padStart(5, '0')}`));
  // one portal reference per file line (a line has several heads / components)
  const pLine = new Map<string, string>();
  r.portal.forEach(e => {
    const k = `${e.file}#${e.row}`;
    if (!pLine.has(k)) pLine.set(k, `PL${String(pLine.size + 1).padStart(5, '0')}`);
    pref.set(e.id, pLine.get(k)!);
  });
  const period = `${ymLabel(meta.from)} to ${ymLabel(meta.to)}`;
  const title = (t: string): C[][] => heading(t, { ...meta, period });
  // each books / portal line's movement in Year_Compare (a column on the line sheets, so Year_Compare is one SUMIFS a cell)
  const y0 = meta.year;
  const code0 = (k: string) => (k.includes(':') ? k.split(':')[1] : k);
  const dupLabel = new Set(y0.comps.map(c => c.label).filter((l, i, a) => a.indexOf(l) !== i));
  const moveName = (c: (typeof y0.comps)[number]) => (dupLabel.has(c.label) ? `${BLOCK_NAME[c.block]}: ${c.label}` : c.label);
  const movement = new Map<string, string>();
  for (const c of y0.comps) for (const k of c.buckets) movement.set(`${c.block}|${code0(k)}`, moveName(c));
  const moveOf = (block: string, bucket: string) => movement.get(`${block}|${code0(bucket)}`) ?? '';
  const hdr = (cols: string[]): C[] => cols.map(c => ({ v: c, s: 'hdr' }));

  // ---- Books_Lines
  const BL_H = 4; // header row
  const books: C[][] = [
    ...title('Books ledger lines (GST accounts)'),
    hdr([
      'Ref',
      'Source file',
      'Row',
      'Date',
      'Month',
      'Compared with return period',
      'In period',
      'Voucher type',
      'Voucher no.',
      'Party / particulars',
      'Account',
      'Account type',
      'Head',
      'Debit',
      'Credit',
      'Effect on balance',
      'Block',
      'Flow row code',
      'Flow row',
      'Row amount',
      'Classified by',
      'Narration',
      'Year_Compare movement',
    ]),
  ];
  r.books.forEach((e, i) => {
    const row = BL_H + 1 + i,
      l = e.line,
      sign = bucketOf(e.bucket)?.sign ?? 1;
    books.push([
      bref.get(e.id),
      l.file,
      l.row,
      { v: serial(l.date), s: 'date' },
      ymLabel(l.date.slice(6) + '-' + l.date.slice(3, 5)),
      ymLabel(e.month),
      e.where,
      l.vchType,
      l.vchNo,
      l.party,
      l.account,
      ACC_TYPE_LABEL[e.accType],
      e.head,
      l.dr || null,
      l.cr || null,
      { f: e.block === 'liab' ? `O${row}-N${row}` : `N${row}-O${row}`, v: e.effect },
      e.block,
      e.bucket,
      bucketLabel(e.bucket),
      { f: sign === 1 ? `P${row}` : `-P${row}`, v: N(e.effect * sign) },
      e.rule,
      l.narration,
      e.where === 'in' && e.block ? moveOf(e.block, e.bucket) : '',
    ]);
  });
  // ---- Portal_Lines
  const portal: C[][] = [
    ...title('Portal ledger lines'),
    hdr([
      'Ref',
      'Ledger',
      'Source file',
      'Row',
      'Date',
      'Reference',
      'Return period',
      'Compared in month',
      'In period',
      'Description',
      'Dr / Cr',
      'Ledger used',
      'Head',
      'Component',
      'Amount',
      'Effect on balance',
      'Block',
      'Flow row code',
      'Flow row',
      'Row amount',
      'Year_Compare movement',
    ]),
  ];
  r.portal.forEach((e, i) => {
    const row = BL_H + 1 + i,
      sign = bucketOf(e.bucket)?.sign ?? 1,
      es = e.amount ? Math.sign(e.effect / e.amount) : 1;
    portal.push([
      pref.get(e.id),
      PORTAL_LABEL[e.kind],
      e.file,
      e.row,
      { v: serial(e.date), s: 'date' },
      e.ref,
      ymLabel(e.month),
      ymLabel(e.month),
      e.block ? e.where : 'not in the flow',
      e.desc,
      e.dir,
      e.used,
      e.head,
      e.comp,
      e.amount,
      { f: es === 1 ? `O${row}` : `-O${row}`, v: e.effect },
      e.block || '',
      e.block ? e.bucket : '',
      e.block ? bucketLabel(e.bucket) : 'Discharge of interest / fee (shown in the cash ledger)',
      e.block ? { f: sign === 1 ? `P${row}` : `-P${row}`, v: N(e.effect * sign) } : null,
      e.where === 'in' && e.block ? moveOf(e.block, e.bucket) : '',
    ]);
  });
  // ---- Balances
  const bal: C[][] = [
    ...title('Opening balances used in the flow'),
    hdr(['Side', 'Block', 'Head', 'Ledger / account', 'Source', "Amount (in the block's direction)", 'As printed or entered']),
  ];
  for (const b of r.balances) bal.push([b.side, b.block, b.head, b.what, b.source, b.amount, b.raw ?? null]);
  bal.push(
    [],
    [
      {
        v: 'Liability balances are shown as payable (+). Ledger exports show debit balances as + and credit balances as −, so a credit balance of an output tax account appears here with the sign changed.',
        s: 'note',
      },
    ],
  );

  // ---- flow statements: the same rows on all three sheets
  const rowKeys = (b: Block) => {
    const present = new Set([...Object.keys(r.flowBooks.rows), ...Object.keys(r.flowPortal.rows)]);
    const ks = BUCKETS.filter(x => x.block === b)
      .map(x => x.id)
      .filter(k => present.has(k) || b !== 'fee');
    for (const extra of [INTERNAL, UNMAPPED]) if (present.has(`${b}:${extra}`)) ks.push(`${b}:${extra}`);
    return ks;
  };
  const HC = HEAD_KEYS.map((_, i) => col(i + 1)); // B..F
  type Layout = { block: Block; open?: number; rows: { key: string; at: number }[]; close?: number };
  const layout: Layout[] = [];
  const flowSheet = (side: 'Books' | 'Portal' | 'Difference'): C[][] => {
    const st = side === 'Books' ? r.flowBooks : side === 'Portal' ? r.flowPortal : null;
    const out: C[][] = [
      ...title(
        side === 'Difference'
          ? 'GST flow: difference (books − portal)'
          : `GST flow as per ${side === 'Books' ? 'books' : 'the GST portal'}`,
      ),
    ];
    out.push([
      {
        v:
          side === 'Books'
            ? 'Built only from the books ledgers (Books_Lines) and the opening balances (Balances). Books by voucher date.'
            : side === 'Portal'
              ? 'Built only from the credit ledger, cash ledger and liability register (Portal_Lines) and their opening balances. Portal by return period.'
              : 'Each cell is Flow_Books minus Flow_Portal. The closing difference is the sum of the row differences. Year_Compare and Closing explain them.',
        s: 'note',
      },
    ]);
    out.push(hdr(['Flow row', ...HEAD_KEYS.map(h => HEADS_TXT[h]), 'Total', 'Row code']));
    const lay = side === 'Books' || !layout.length;
    for (const b of BLOCKS) {
      const L: Layout = lay ? { block: b, rows: [] } : layout.find(x => x.block === b)!;
      out.push([
        { v: BLOCK_LABEL[b], s: 'sec' },
        ...HEAD_KEYS.map(() => ({ v: '', s: 'sec' }) as C),
        { v: '', s: 'sec' },
        { v: '', s: 'sec' },
      ]);
      const cellF = (key: string, h: HeadK, r1: number): C => {
        if (side === 'Difference')
          return {
            f: `Flow_Books!${col(HEAD_KEYS.indexOf(h) + 1)}${r1}-Flow_Portal!${col(HEAD_KEYS.indexOf(h) + 1)}${r1}`,
            v: N((r.flowBooks.rows[key]?.[h] || 0) - (r.flowPortal.rows[key]?.[h] || 0)),
          };
        const sh = side === 'Books' ? 'Books_Lines' : 'Portal_Lines';
        const hc = side === 'Books' ? 'M' : 'M',
          kc = 'R',
          vc = 'T',
          inc = side === 'Books' ? 'G' : 'I';
        return {
          f: `SUMIFS(${sh}!$${vc}:$${vc},${sh}!$${kc}:$${kc},$H${r1},${sh}!$${hc}:$${hc},"${h}",${sh}!$${inc}:$${inc},"in")`,
          v: N(st!.rows[key]?.[h] || 0),
        };
      };
      if (b !== 'fee') {
        const r1 = out.length + 1;
        if (lay) L.open = r1;
        out.push([
          'Opening balance',
          ...HEAD_KEYS.map(h =>
            side === 'Difference'
              ? ({
                  f: `Flow_Books!${col(HEAD_KEYS.indexOf(h) + 1)}${r1}-Flow_Portal!${col(HEAD_KEYS.indexOf(h) + 1)}${r1}`,
                  v: N(r.flowBooks.opening[b][h] - r.flowPortal.opening[b][h]),
                  s: 'numB',
                } as C)
              : ({
                  f: `SUMIFS(Balances!$F:$F,Balances!$A:$A,"${side}",Balances!$B:$B,"${b}",Balances!$C:$C,"${h}")`,
                  v: N(st!.opening[b][h]),
                  s: 'numB',
                } as C),
          ),
          { f: `SUM(B${r1}:F${r1})`, s: 'numB', v: 0 },
          'opening',
        ]);
      }
      for (const key of rowKeys(b)) {
        const r1 = out.length + 1;
        if (lay) L.rows.push({ key, at: r1 });
        const sign = bucketOf(key)?.sign ?? 1;
        out.push([
          `${b === 'fee' ? '' : sign === 1 ? 'Add: ' : 'Less: '}${bucketLabel(key)}`,
          ...HEAD_KEYS.map(h => cellF(key, h, r1)),
          { f: `SUM(B${r1}:F${r1})`, v: 0 },
          key,
        ]);
      }
      const r1 = out.length + 1;
      if (lay) L.close = r1;
      const terms = (c: string) =>
        [...(L.open ? [`${c}${L.open}`] : []), ...L.rows.map(x => `${(bucketOf(x.key)?.sign ?? 1) === 1 ? '+' : '-'}${c}${x.at}`)]
          .join('')
          .replace(/^\+/, '');
      const src = side === 'Books' ? r.flowBooks : side === 'Portal' ? r.flowPortal : null;
      out.push([
        b === 'fee' ? 'Total' : 'Closing balance',
        ...HEAD_KEYS.map(
          (h, i) =>
            ({
              f: terms(HC[i]) || '0',
              s: 'numTot',
              v: src ? N(src.closing[b][h]) : N(r.flowBooks.closing[b][h] - r.flowPortal.closing[b][h]),
            }) as C,
        ),
        { f: `SUM(B${r1}:F${r1})`, s: 'numTot', v: 0 },
        '',
      ]);
      out.push([]);
      if (lay) layout.push(L);
    }
    return out;
  };
  // fill the Total column cached values
  const fixTotals = (rows: C[][]) =>
    rows.forEach(rw => {
      const t = rw[6] as any;
      if (t && typeof t === 'object' && t.f?.startsWith('SUM(B'))
        t.v = N(rw.slice(1, 6).reduce((s: number, c: any) => s + (typeof c === 'object' && c ? +c.v || 0 : 0), 0));
    });
  const fBooks = flowSheet('Books'),
    fPortal = flowSheet('Portal'),
    fDiff = flowSheet('Difference');
  [fBooks, fPortal, fDiff].forEach(fixTotals);
  fDiff.push([
    {
      v: 'Year_Compare compares each movement for the year; Closing takes the closing as per books, with the suggested journals, to the closing as per portal.',
      s: 'note',
    },
  ]);

  const y = meta.year,
    it = meta.itc;
  const blank = (n: number): C[] =>
    Array(n)
      .fill(0)
      .map(() => ({ v: '', s: 'sec' }) as C);
  const sum = (a: Record<string, number>, ks: readonly string[]) => N(ks.reduce((t, k) => t + (a[k] || 0), 0));
  // ---- Setup_Checks
  const su: C[][] = [...title('Set-up checks: fix these first'), hdr(['Level', 'Issue', 'What to do'])];
  for (const x of y.setup)
    su.push([
      { v: x.level === 'fix' ? 'Fix first' : x.level === 'check' ? 'Check' : 'Note', s: x.level === 'fix' ? 'bad' : 'bold' },
      { v: x.what, s: 'wrap' },
      { v: x.action, s: 'wrap' },
    ]);
  if (!y.setup.length)
    su.push(['OK', 'All three portal ledgers are loaded, the openings are in, and every books line is in a flow row.']);

  // ---- Year_Compare: formulas over the line sheets (column P = effect on the balance; Q block; R flow row code; books F / portal H = month compared)
  const ms = r.months.filter(m => m >= meta.from && m <= meta.to);
  // one SUMIFS a cell: the "Year_Compare movement" column of the line sheets (W on Books_Lines, U on Portal_Lines)
  const sumF = (sh: 'Books_Lines' | 'Portal_Lines', c: { gsign: number }, row: number, monthCell?: string) => {
    const mc = sh === 'Books_Lines' ? 'F' : 'H',
      vc = sh === 'Books_Lines' ? 'W' : 'U';
    const f = `SUMIFS(${sh}!$P:$P,${sh}!$${vc}:$${vc},$A${row}${monthCell ? `,${sh}!$${mc}:$${mc},${monthCell}` : ''})`;
    return c.gsign === 1 ? f : `-${f}`;
  };
  const yc: C[][] = [
    ...title('Year comparison: each movement as per books and as per portal'),
    hdr([
      'Movement',
      'Books',
      'Portal',
      'Difference',
      'Of which year-end cut-off',
      'To look at',
      'Where it is worked',
      'What to do',
      ...ms.map(m => ymLabel(m)),
    ]),
  ];
  for (const b of BLOCKS) {
    const cs = y.comps.filter(c => c.block === b);
    if (!cs.length) continue;
    yc.push([{ v: BLOCK_NAME[b], s: 'sec' }, ...blank(7 + ms.length)]);
    for (const c of cs) {
      const row = yc.length + 1;
      yc.push([
        moveName(c),
        { f: sumF('Books_Lines', c, row), v: sum(c.books, HEAD_KEYS) },
        { f: sumF('Portal_Lines', c, row), v: sum(c.portal, HEAD_KEYS) },
        { f: `B${row}-C${row}`, v: sum(c.diff, HEAD_KEYS), s: 'numB' },
        sum(c.cutoff, HEAD_KEYS) || null,
        { f: `D${row}-E${row}`, v: sum(c.rest, HEAD_KEYS), s: 'numB' },
        Math.abs(sum(c.diff, HEAD_KEYS)) > 0.5
          ? c.kind === 'journal'
            ? 'Journals'
            : c.kind === 'itc'
              ? 'ITC_3way'
              : c.kind === 'setup'
                ? 'Setup_Checks'
                : 'Invoice level'
          : 'Agrees',
        { v: Math.abs(sum(c.diff, HEAD_KEYS)) > 0.5 ? c.action : '', s: 'wrap' },
        ...c.months.map((x, i) => {
          const mcell = `${col(8 + i)}$${BL_H}`,
            pf = sumF('Portal_Lines', c, row, mcell);
          return { f: `${sumF('Books_Lines', c, row, mcell)}${pf.startsWith('-') ? '+' + pf.slice(1) : '-' + pf}`, v: x.d } as C;
        }),
      ]);
    }
    yc.push([]);
  }
  yc.push([
    {
      v: 'How it is worked out: Books = SUMIFS of "Effect on balance" (column P) of Books_Lines whose "Year_Compare movement" (column W) is the movement in column A; Portal = the same over Portal_Lines (column U); the month columns add the month in row 4 as a second condition. A movement that reduces the balance is shown with its sign turned. Totals across tax heads. Flow_Difference and Month_wise show each head. The year-end cut-off is what was in the books in the period and on the portal in a later return (or the other way): it clears next period.',
      s: 'note',
    },
  ]);

  // ---- ITC_3way and ITC_Invoices
  const HD = ['IGST', 'CGST', 'SGST', 'CESS'] as const;
  const i3: C[][] = [[{ v: 'ITC: books, GSTR-2B and GSTR-3B', s: 'title' }], title('')[1], []];
  const inv: C[][] = [
    ...title('ITC: invoices that differ between books and GSTR-2B'),
    hdr([
      'Status',
      'Supplier',
      'GSTIN',
      'Books no.',
      '2B no.',
      'Books date',
      '2B date',
      '2B period',
      'Tax in books',
      'Tax in 2B',
      'Difference',
      'What to do',
    ]),
  ];
  if (!it)
    i3.push([
      {
        v: 'No GSTR-2B files were loaded. Add the 2B JSONs of the period in the Loader to split the ITC difference invoice by invoice.',
        s: 'note',
      },
    ]);
  else {
    const bridge = (t: string, ls: typeof it.books) => {
      i3.push([{ v: t, s: 'sec' }, ...blank(6)], hdr(['', 'IGST', 'CGST', 'SGST', 'Cess', 'Total', 'What to do']));
      for (const l of ls) {
        if (!(l.kind === 'start' || l.kind === 'total') && !HD.some(h => Math.abs(l.amt[h]) > 0.5)) continue;
        const row = i3.length + 1,
          st = l.kind === 'start' || l.kind === 'total' ? 'numTot' : l.kind === 'check' ? 'numB' : undefined;
        i3.push([
          { v: l.label, s: st ? 'bold' : 'wrap' },
          ...HD.map(h => ({ v: N(l.amt[h]), s: st }) as C),
          { f: `SUM(B${row}:E${row})`, v: sum(l.amt, HD), s: st },
          { v: l.action || '', s: 'wrap' },
        ]);
      }
      i3.push([]);
    };
    bridge('Books to GSTR-2B (invoice by invoice: ITC_Invoices)', it.books);
    bridge('GSTR-2B to GSTR-3B (claimed in the credit ledger)', it.claim);
    i3.push(
      [{ v: 'By month', s: 'sec' }, ...blank(6)],
      hdr(['Month', 'ITC booked', 'Available in 2B', 'Claimed in 3B', '2B − 3B', 'Running', '']),
    );
    for (const m of it.months) {
      const row = i3.length + 1;
      i3.push([
        ymLabel(m.m),
        m.booked,
        m.b2b,
        m.claimed,
        { f: `C${row}-D${row}`, v: m.diff, s: 'numB' },
        { f: `F${row - 1}+E${row}`, v: m.run },
      ]);
    }
    for (const n of it.notes) i3.push([{ v: n, s: 'note' }]);
    if (it.register)
      i3.push(
        [],
        [
          {
            v: `Purchase register cross-check: tax in the register ${it.register.register}, on the purchase invoices in the GL ${it.register.books}.`,
            s: 'note',
          },
        ],
      );
    for (const v of it.invoices) {
      const row = inv.length + 1;
      inv.push([
        v.status,
        v.supplier,
        v.gstin,
        v.bNo,
        v.rNo,
        v.bDate,
        v.rDate,
        v.period,
        v.bTax || null,
        v.rTax || null,
        { f: `I${row}-J${row}`, v: N(v.bTax - v.rTax) },
        { v: v.note, s: 'wrap' },
      ]);
    }
    if (!it.invoices.length) inv.push(['Every purchase invoice in the books matches GSTR-2B.']);
    if (it.nonInvoice.length) {
      inv.push(
        [],
        [{ v: 'ITC entries without a purchase invoice (journals, adjustments)', s: 'sec' }, ...blank(11)],
        hdr(['Date', 'Voucher', 'Account', '', '', '', '', '', 'ITC', '', '', 'Books line']),
      );
      for (const n of it.nonInvoice) inv.push([n.date, n.voucher, n.account, '', '', '', '', '', n.tax, '', '', bref.get(n.id)]);
    }
  }
  // fix the first running cell of the month table (no row above to add to)
  for (let i = 0; i < i3.length; i++)
    if (i3[i]?.[0] && (i3[i][0] as any).v === 'Month') {
      const f = i3[i + 1]?.[5] as any;
      if (f) f.f = `E${i + 2}`;
    }

  // ---- Journals
  const jr: C[][] = [
    ...title('Suggested journals: books brought in line with the GST portal'),
    hdr(['No.', 'Date', 'Account', 'Head', 'Debit', 'Credit', 'Block', 'Effect on the block', 'Narration']),
  ];
  for (const j of y.journals) {
    for (const [i, l] of j.legs.entries()) {
      const row = jr.length + 1;
      jr.push([
        j.no,
        { v: serial(j.date), s: 'date' },
        l.existing ? l.account : `${l.account} (pick the account)`,
        l.head,
        l.dr || null,
        l.cr || null,
        l.block,
        l.block === 'other'
          ? null
          : {
              f: l.block === 'liab' ? `F${row}-E${row}` : `E${row}-F${row}`,
              v: N(l.block === 'liab' ? l.cr - l.dr : l.dr - l.cr),
            },
        i === 0 ? { v: `${j.title}. ${j.narration}`, s: 'wrap' } : '',
      ]);
    }
    jr.push([]);
  }
  if (!y.journals.length) jr.push(['None needed: set-off, payments and deposits agree with the portal.']);
  jr.push([
    {
      v: 'Nothing here is posted. Accounts marked "pick the account" are not in the books loaded: use your own ledger. Suspense lines show a set-off or payment the books do not have on both sides.',
      s: 'note',
    },
  ]);

  // ---- Closing: books + journals, the remaining differences, portal (formulas to the flow sheets and Journals)
  const cl: C[][] = [
    ...title('Closing statement: books to the GST portal'),
    hdr(['', ...HEAD_KEYS.map(h => HEADS_TXT[h]), 'Total', 'What to do']),
  ];
  const closeAt: { block: Block; books: number; journals: number; adjusted: number; portal: number; unexplained: number[] }[] =
    [];
  for (const c of y.closing) {
    const L = layout.find(x => x.block === c.block)!;
    cl.push([{ v: BLOCK_NAME[c.block], s: 'sec' }, ...blank(7)]);
    const rb = cl.length + 1;
    cl.push([
      { v: 'Closing as per books', s: 'bold' },
      ...HEAD_KEYS.map((h, i) => ({ f: `Flow_Books!${HC[i]}${L.close}`, v: N(c.books[h]), s: 'numB' }) as C),
      { f: `SUM(B${rb}:F${rb})`, v: sum(c.books, HEAD_KEYS), s: 'numB' },
    ]);
    const rj = cl.length + 1;
    cl.push([
      'Add: suggested journals (Journals)',
      ...HEAD_KEYS.map(
        h => ({ f: `SUMIFS(Journals!$H:$H,Journals!$G:$G,"${c.block}",Journals!$D:$D,"${h}")`, v: N(c.journals[h]) }) as C,
      ),
      { f: `SUM(B${rj}:F${rj})`, v: sum(c.journals, HEAD_KEYS) },
    ]);
    const ra = cl.length + 1;
    cl.push([
      { v: 'Adjusted closing as per books', s: 'bold' },
      ...HEAD_KEYS.map((h, i) => ({ f: `${HC[i]}${rb}+${HC[i]}${rj}`, v: N(c.adjusted[h]), s: 'numTot' }) as C),
      { f: `SUM(B${ra}:F${ra})`, v: sum(c.adjusted, HEAD_KEYS), s: 'numTot' },
    ]);
    const lr: number[] = [];
    for (const l of c.lines) {
      const row = cl.length + 1;
      lr.push(row);
      cl.push([
        { v: `${l.kind === 'unexplained' ? 'NOT EXPLAINED: ' : 'Less: '}${l.label}`, s: 'wrap' },
        ...HEAD_KEYS.map(h => N(l.amt[h]) || null),
        { f: `SUM(B${row}:F${row})`, v: sum(l.amt, HEAD_KEYS) },
        { v: l.action, s: 'wrap' },
      ]);
    }
    const rp = cl.length + 1;
    cl.push([
      { v: 'Closing as per portal', s: 'bold' },
      ...HEAD_KEYS.map((h, i) => ({ f: `Flow_Portal!${HC[i]}${L.close}`, v: N(c.portal[h]), s: 'numTot' }) as C),
      { f: `SUM(B${rp}:F${rp})`, v: sum(c.portal, HEAD_KEYS), s: 'numTot' },
    ]);
    closeAt.push({
      block: c.block,
      books: rb,
      journals: rj,
      adjusted: ra,
      portal: rp,
      unexplained: lr.filter((_, i) => c.lines[i].kind === 'unexplained'),
    });
    const rc = cl.length + 1;
    cl.push([
      { v: 'Check (should be nil)', s: 'note' },
      ...HEAD_KEYS.map((h, i) => ({ f: `${HC[i]}${ra}${lr.map(x => `-${HC[i]}${x}`).join('')}-${HC[i]}${rp}`, v: 0 }) as C),
      { f: `SUM(B${rc}:F${rc})`, v: 0 },
    ]);
    cl.push([]);
  }
  cl.push([
    {
      v: 'Lines are differences between the adjusted books and the portal: positive means the books are higher. A line in the Total column only, with nil total, is a books account not split by tax head.',
      s: 'note',
    },
  ]);
  // ---- Cross_Checks
  const chk: C[][] = [
    ...title('Cross-checks between the portal ledgers and with GSTR-3B'),
    hdr(['Check', 'Return period', 'Head', 'A', 'Amount A', 'B', 'Amount B', 'Difference', 'Status', 'Detail']),
  ];
  r.checks.forEach((c, i) => {
    const row = BL_H + 1 + i;
    chk.push([
      c.area,
      ymLabel(c.period),
      c.head,
      c.aLabel,
      c.a,
      c.bLabel,
      c.b,
      { f: `E${row}-G${row}`, v: c.diff },
      { f: `IF(ABS(H${row})<=1,"OK","Check")`, v: c.ok ? 'OK' : 'Check', s: c.ok ? 'ok' : 'bad' },
      c.note || '',
    ]);
  });
  if (!r.checks.length) chk.push(['No checks: load at least two portal ledgers (and GSTR-3B PDFs for the 3B checks).']);
  if (r.warns.length) {
    chk.push([], [{ v: 'Messages', s: 'bold' }]);
    r.warns.forEach(w => chk.push([w.level, { v: w.msg, s: 'wrap' }]));
  }
  // ---- Month_wise
  const mw: C[][] = [
    ...title('Month-wise: books, portal and difference'),
    hdr([
      'Block',
      'Flow row',
      'Head',
      ...r.months.flatMap(m => [`${ymLabel(m)} books`, `${ymLabel(m)} portal`, `${ymLabel(m)} diff`]),
      'Row code',
    ]),
  ];
  const keyHeads = new Map<string, Set<HeadK>>();
  for (const e of r.books)
    if (e.block && e.where === 'in') {
      const k = e.bucket === INTERNAL || e.bucket === UNMAPPED ? `${e.block}:${e.bucket}` : e.bucket;
      (keyHeads.get(k) || keyHeads.set(k, new Set()).get(k)!).add(e.head as HeadK);
    }
  for (const e of r.portal)
    if (e.block && e.where === 'in') (keyHeads.get(e.bucket) || keyHeads.set(e.bucket, new Set()).get(e.bucket)!).add(e.head);
  const monthsCol = (m: string) => ymLabel(m);
  for (const b of BLOCKS)
    for (const key of rowKeys(b))
      for (const h of HEAD_KEYS) {
        if (!keyHeads.get(key)?.has(h)) continue;
        const row = mw.length + 1,
          cells: C[] = [BLOCK_LABEL[b], bucketLabel(key), h];
        r.months.forEach((m, i) => {
          const bc = col(3 + i * 3),
            pc = col(4 + i * 3);
          const bv = N(
            r.books
              .filter(
                e => e.where === 'in' && e.head === h && e.month === m && (e.bucket === key || `${e.block}:${e.bucket}` === key),
              )
              .reduce((s, e) => s + e.effect * (bucketOf(e.bucket)?.sign ?? 1), 0),
          );
          const pv = N(
            r.portal
              .filter(e => e.where === 'in' && e.head === h && e.month === m && e.bucket === key && e.block)
              .reduce((s, e) => s + e.effect * (bucketOf(e.bucket)?.sign ?? 1), 0),
          );
          cells.push(
            {
              f: `SUMIFS(Books_Lines!$T:$T,Books_Lines!$R:$R,$${col(3 + r.months.length * 3)}${row},Books_Lines!$M:$M,$C${row},Books_Lines!$F:$F,"${monthsCol(m)}",Books_Lines!$G:$G,"in")`,
              v: bv,
            },
            {
              f: `SUMIFS(Portal_Lines!$T:$T,Portal_Lines!$R:$R,$${col(3 + r.months.length * 3)}${row},Portal_Lines!$M:$M,$C${row},Portal_Lines!$H:$H,"${monthsCol(m)}",Portal_Lines!$I:$I,"in")`,
              v: pv,
            },
            { f: `${bc}${row}-${pc}${row}`, v: N(bv - pv), s: 'numB' },
          );
        });
        cells.push(key);
        mw.push(cells);
      }
  // ---- Audit
  const au: C[][] = [
    ...title('Your changes to the records (reclassifications, periods, balances, account types)'),
    hdr(['When', 'Change', 'Record', 'From', 'To', 'Remark']),
  ];
  meta.audit.forEach(a => au.push([a.when, a.change, a.record, a.from, a.to, { v: a.remark, s: 'wrap' }]));
  if (!meta.audit.length)
    au.push(["No changes: every classification is the tool's own, from the account types confirmed in the Loader."]);
  // ---- Summary: the closing position of each ledger and what is left, all links to Closing / Flow sheets
  const sm: C[][] = [
    ...title('Summary: GST ledgers as per books and as per the portal'),
    hdr([
      'Closing balance (all heads)',
      'As per books',
      'Suggested journals',
      'Books after journals',
      'As per portal',
      'Not explained',
      'How it is worked out',
    ]),
  ];
  for (const c of closeAt) {
    const row = sm.length + 1,
      cc = y.closing.find(x => x.block === c.block)!;
    const unx = cc.lines.filter(l => l.kind === 'unexplained').reduce((t, l) => t + sum(l.amt, HEAD_KEYS), 0);
    sm.push([
      { v: BLOCK_NAME[c.block], s: 'bold' },
      { f: `Closing!G${c.books}`, v: sum(cc.books, HEAD_KEYS) },
      { f: `Closing!G${c.journals}`, v: sum(cc.journals, HEAD_KEYS) },
      { f: `B${row}+C${row}`, v: sum(cc.adjusted, HEAD_KEYS), s: 'numB' },
      { f: `Closing!G${c.portal}`, v: sum(cc.portal, HEAD_KEYS) },
      c.unexplained.length
        ? { f: c.unexplained.map(x => `Closing!G${x}`).join('+'), v: N(unx), s: Math.abs(unx) > 0.5 ? 'bad' : 'ok' }
        : { v: 0, s: 'ok' },
      how(
        `Closing sheet, Total column: books (row ${c.books}), journals (row ${c.journals}), portal (row ${c.portal})${c.unexplained.length ? `, not explained (row ${c.unexplained.join(', ')})` : ''}.`,
      ),
    ]);
  }
  const nFix = y.setup.filter(x => x.level === 'fix').length,
    nChk = r.checks.filter(c => !c.ok).length;
  sm.push(
    [],
    [{ v: 'Status', s: 'sec' }, ...blank(6)],
    hdr(['Item', 'Count', '', '', '', '', 'How it is worked out']),
    [
      'Set-up issues to fix first',
      { f: 'COUNTIF(Setup_Checks!$A:$A,"Fix first")', v: nFix, s: nFix ? 'bad' : 'ok' },
      '',
      '',
      '',
      '',
      how('Count of "Fix first" in column A of Setup_Checks.'),
    ],
    [
      'Cross-checks that do not agree',
      { f: 'COUNTIF(Cross_Checks!$I:$I,"Check")', v: nChk, s: nChk ? 'bad' : 'ok' },
      '',
      '',
      '',
      '',
      how('Count of "Check" in the Status column of Cross_Checks (difference above ₹1).'),
    ],
    [
      'Suggested journals',
      y.journals.length,
      '',
      '',
      '',
      '',
      how('Journals sheet: review them and post them in the books; nothing is posted by the tool.'),
    ],
    [
      'Your changes to the records',
      meta.audit.length,
      '',
      '',
      '',
      '',
      how('Audit sheet: every change made in the Loader, with your remark.'),
    ],
    [],
    [{ v: 'Movements for the year (Year_Compare)', s: 'sec' }, ...blank(6)],
    hdr(['Movement', 'Books', 'Portal', 'Difference', 'Year-end cut-off', 'To look at', 'How it is worked out']),
  );
  for (let i = 0; i < yc.length; i++) {
    const rw = yc[i];
    const a0 = rw?.[0];
    const isLine = rw && rw.length > 6 && typeof rw[1] === 'object' && rw[1] && (rw[1] as any).f;
    if (!isLine) continue;
    const yr = i + 1,
      row = sm.length + 1;
    sm.push([
      typeof a0 === 'string' ? a0 : ((a0 as any)?.v ?? ''),
      { f: `Year_Compare!B${yr}`, v: (rw[1] as any).v },
      { f: `Year_Compare!C${yr}`, v: (rw[2] as any).v },
      { f: `B${row}-C${row}`, v: (rw[3] as any).v, s: 'numB' },
      { f: `Year_Compare!E${yr}`, v: typeof rw[4] === 'number' ? rw[4] : 0 },
      { f: `D${row}-E${row}`, v: (rw[5] as any).v, s: Math.abs((rw[5] as any).v || 0) > 0.5 ? 'bad' : 'numB' },
      how(`Year_Compare row ${yr}`),
    ]);
  }

  // ---- Instructions (one sheet) and Index
  const front = {
    title: 'GST ledger flow: books vs GST portal',
    legalName: meta.legalName,
    gstin: meta.gstin,
    fy: meta.fy,
    period,
    builtAt: meta.builtAt,
    version: LOADER_VERSION,
  };
  const guideLines: [string, string][] = FLOW_GUIDE.flatMap(g => [
    [
      `${g.source === 'portal' ? 'GST portal' : g.source === 'tally' ? 'Tally' : g.source === 'zoho' ? 'Zoho Books' : 'Other'}: ${g.title}`,
      g.steps.map((x, i) => `${i + 1}. ${x}`).join('  ') + (g.note ? `  (${g.note})` : ''),
    ] as [string, string],
  ]);
  const sections: Section[] = [
    {
      heading: 'What this workpaper does',
      lines: [
        [
          'Purpose',
          'Takes the GST ledgers in the books (ITC, output tax, cash ledger, interest / fees) and the portal Electronic Credit Ledger, Cash Ledger and Liability Register to the same flow: opening, movements, closing. Every difference is explained from the entries, journals are suggested, and the Closing sheet takes the books to the portal.',
        ],
        ['Start here', 'Summary, then Setup_Checks (fix these first), Year_Compare, ITC_3way, Journals and Closing.'],
      ],
    },
    {
      heading: 'How the figures are worked out',
      lines: [
        [
          'Books_Lines / Portal_Lines',
          'Every line read from the files, with its flow row (column R), its effect on the balance (column P) and its Year_Compare movement. Nothing else feeds the statements.',
        ],
        [
          'Flow_Books / Flow_Portal',
          'Each cell = SUMIFS of "Row amount" (column T) of the line sheet for that flow row code (column H of the flow sheet), tax head and "in period". Opening = SUMIFS over Balances.',
        ],
        ['Flow_Difference', 'Flow_Books minus Flow_Portal, cell by cell.'],
        [
          'Year_Compare',
          'One SUMIFS a cell over the "Year_Compare movement" column of the line sheets; month columns add the month as a condition.',
        ],
        [
          'Closing',
          'Closing as per books (link to Flow_Books) + suggested journals (SUMIFS over Journals) − each remaining difference = closing as per portal (link to Flow_Portal). The Check row must be nil.',
        ],
        [
          'Cross_Checks',
          'A − B for each pair (cash ledger vs liability paid in cash, credit ledger vs liability paid by ITC, ledgers vs GSTR-3B); "Check" when the difference is above ₹1.',
        ],
        [
          'Basis',
          'Books by voucher date; portal entries by return period (the April filing of the March return counts in March). A books line moved to another return period in the Loader is compared in that period, and the change is in Audit.',
        ],
      ],
    },
    {
      heading: 'Working with the tool',
      lines: [
        [
          'Change a classification',
          'In the Loader (GST ledger flow tab), open a month in the year comparison, Show lines, Change, with a remark. Changes go to Audit.',
        ],
        [
          'Next time',
          'Drop this saved workpaper into the Loader: it brings back its source files, the account types, balances and your changes. Drop a newer GL dump with it to replace the saved copy; portal ledgers stay as filed.',
        ],
        [
          'Annual return',
          'This workpaper, with the Input Reco and Revenue reco workpapers, is what the Annual return tab builds GSTR-9 / 9C from.',
        ],
      ],
    },
    { heading: 'Settings used', lines: (meta.settings ?? []).map(t => ['Matching settings', t] as [string, string]) },
    { heading: 'Notes by the preparer', lines: meta.notes ?? [] },
    {
      heading: 'Files used',
      lines: meta.files.map(f => [f.name, `${f.kind}${f.note ? ' · ' + f.note : ''}`] as [string, string]),
    },
    { heading: 'How to export each report', lines: guideLines },
  ];
  const info: SheetInfo[] = [
    {
      name: 'Summary',
      role: 'Summary',
      about:
        'Closing balances as per books and portal, what is not explained, the status of the checks, and each movement for the year.',
    },
    {
      name: 'Setup_Checks',
      role: 'Working',
      about: 'Fix these first: a missing ledger, opening or flow row shows up as a difference everywhere else.',
    },
    {
      name: 'Year_Compare',
      role: 'Working',
      about: 'Each movement for the year as per books and as per portal, the year-end cut-off and the month-wise difference.',
    },
    {
      name: 'ITC_3way',
      role: 'Working',
      about: 'ITC as per books, GSTR-2B and GSTR-3B: books to 2B invoice by invoice, 2B to 3B month by month.',
    },
    { name: 'ITC_Invoices', role: 'Working', about: 'The purchase invoices that differ between the books and GSTR-2B.' },
    { name: 'Journals', role: 'Working', about: 'Suggested entries that bring the books in line with the portal (not posted).' },
    {
      name: 'Closing',
      role: 'Working',
      about: 'Closing as per books + journals − each remaining difference = closing as per portal. Check must be nil.',
    },
    {
      name: 'Flow_Books',
      role: 'Working',
      about: 'Opening, movements and closing from the books only (SUMIFS over Books_Lines).',
    },
    { name: 'Flow_Portal', role: 'Working', about: 'The same rows from the portal ledgers only (SUMIFS over Portal_Lines).' },
    { name: 'Flow_Difference', role: 'Working', about: 'Flow_Books minus Flow_Portal, row by row.' },
    { name: 'Cross_Checks', role: 'Working', about: 'The portal ledgers against each other and against GSTR-3B.' },
    { name: 'Month_wise', role: 'Working', about: 'Books, portal and difference for each flow row, head and month.' },
    { name: 'Books_Lines', role: 'Source data', about: 'Every books GST ledger line with its classification.' },
    { name: 'Portal_Lines', role: 'Source data', about: 'Every portal ledger line with its classification.' },
    {
      name: 'Balances',
      role: 'Source data',
      about: 'Opening balances: entered, printed in the files, or carried from entries before the period.',
    },
    { name: 'Audit', role: 'Your input', about: 'Every change made in the Loader, with its remark.' },
  ];
  const sheets: SheetSpec[] = [
    indexSheet(front, info),
    instructionsSheet(front, sections),
    { name: 'Summary', rows: sm, widths: [46, 16, 16, 18, 16, 16, 70], freeze: { row: 4 }, tab: '2C388E' },
    { name: 'Setup_Checks', rows: su, widths: [12, 90, 80], freeze: { row: BL_H } },
    {
      name: 'Year_Compare',
      rows: yc,
      widths: [52, 15, 15, 15, 15, 15, 16, 60, ...ms.map(() => 12)],
      freeze: { row: BL_H, col: 1 },
      tab: '2C388E',
    },
    { name: 'ITC_3way', rows: i3, widths: [70, 14, 14, 14, 12, 14, 70], freeze: { row: 3 }, tab: '2C388E' },
    {
      name: 'ITC_Invoices',
      rows: inv,
      widths: [22, 32, 18, 18, 18, 11, 11, 10, 14, 14, 14, 80],
      freeze: { row: BL_H },
      filter: { row: BL_H, cols: 12 },
    },
    { name: 'Journals', rows: jr, widths: [6, 11, 40, 8, 14, 14, 8, 14, 90], freeze: { row: BL_H } },
    { name: 'Closing', rows: cl, widths: [70, 14, 14, 14, 12, 14, 14, 70], freeze: { row: BL_H, col: 1 }, tab: '2C388E' },
    { name: 'Flow_Books', rows: fBooks, widths: [52, 16, 16, 16, 16, 18, 18, 14], freeze: { row: 5, col: 1 }, tab: '2C388E' },
    { name: 'Flow_Portal', rows: fPortal, widths: [52, 16, 16, 16, 16, 18, 18, 14], freeze: { row: 5, col: 1 }, tab: '2C388E' },
    { name: 'Flow_Difference', rows: fDiff, widths: [70, 16, 16, 16, 16, 18, 18, 16], freeze: { row: 5, col: 1 }, tab: '2C388E' },
    {
      name: 'Cross_Checks',
      rows: chk,
      widths: [44, 12, 8, 30, 15, 32, 15, 14, 9, 80],
      freeze: { row: BL_H },
      filter: { row: BL_H, cols: 10 },
    },
    {
      name: 'Month_wise',
      rows: mw,
      widths: [24, 36, 8, ...r.months.flatMap(() => [13, 13, 13]), 14],
      freeze: { row: BL_H, col: 3 },
    },
    {
      name: 'Books_Lines',
      rows: books,
      widths: [10, 26, 6, 11, 8, 12, 8, 14, 14, 30, 26, 22, 7, 13, 13, 13, 6, 14, 30, 13, 40, 40, 34],
      freeze: { row: BL_H, col: 1 },
      filter: { row: BL_H, cols: 23 },
    },
    {
      name: 'Portal_Lines',
      rows: portal,
      widths: [10, 26, 26, 6, 11, 20, 10, 10, 8, 32, 6, 8, 7, 9, 13, 13, 6, 14, 30, 13, 34],
      freeze: { row: BL_H, col: 1 },
      filter: { row: BL_H, cols: 21 },
    },
    { name: 'Balances', rows: bal, widths: [8, 6, 8, 34, 50, 18, 16], freeze: { row: BL_H } },
    { name: 'Audit', rows: au, widths: [18, 26, 40, 24, 24, 60], freeze: { row: BL_H } },
  ];
  return writeBook(sheets);
}

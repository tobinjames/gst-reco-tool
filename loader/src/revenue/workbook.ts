// Revenue reconciliation workbook. Rev_Items carries every document / B2CS / Table 8 row with its books and GSTR-1 side;
// "Your reason", "B2CS month" and "Remark" are yours to fill (drop-down lists) and are read back when the saved workbook is
// dropped into the Loader again. The statements, month-wise and B2C sheets are SUMIFS formulas over Rev_Items,
// G1_Rows and 3B_Rows (cached values written too; Excel recalculates on open).
import { SheetInfo, Section, heading, how, indexSheet, instructionsSheet } from '../wpfront';
import { LOADER_VERSION } from '../version';
import { C, SheetSpec, col, serial, writeBook } from '../xlsxbook';
import {
  Item,
  REASON,
  REASONS,
  RevResult,
  TABLES,
  TABLE_LABEL,
  Table,
  monthsOf,
  tableOfKey,
  ymLabel,
  ymOfPeriod,
} from './engine';
import { REV_CLASS_LABEL, RevClass } from './ledgers';
import { REV_GUIDE } from './guide';
import type { RawRow } from '../util';
import type { ThreeB } from '../gstr3b';

export interface RevAudit {
  when: string;
  change: string;
  record: string;
  from: string;
  to: string;
  remark: string;
}
export interface RevMeta {
  gstin: string;
  legalName: string;
  fy: string;
  source: string;
  files: { name: string; kind: string; note: string }[];
  audit: RevAudit[];
  builtAt: string;
  g1: RawRow[];
  threeB: ThreeB[];
  /** books input rows (Books_Sales layout) */
  bookRows: any[][];
  ledgers?: { name: string; cls: RevClass | null; dr: number; cr: number; lines: number }[];
  /** the Matching settings in use, in words */
  settings?: string[];
  /** books customers without a GSTIN and the GSTIN given to them, with the basis */
  customers?: { name: string; gstin: string; basis: string; detail: string }[];
}
export const ITEMS_SHEET = 'Rev_Items';
export const ITEMS_HDR = 5; // header row of Rev_Items
export const ITEM_COLS = [
  'Id',
  'Status',
  'Customer',
  'GSTIN',
  'Type',
  'Doc type',
  'Books no.',
  'GSTR-1 no.',
  'Books date',
  'Books month',
  'GSTR-1 section',
  'GSTR-1 month',
  'Books taxable',
  'Books tax',
  'GSTR-1 taxable',
  'GSTR-1 tax',
  'Diff taxable (GSTR-1 - books)',
  'Diff tax (GSTR-1 - books)',
  "Tool's reason",
  'Your reason',
  'B2CS month',
  'Reason used',
  'Remark',
  'Note',
  'Match basis',
  'Source',
] as const;
const IX = (name: (typeof ITEM_COLS)[number]) => ITEM_COLS.indexOf(name);

const N = (v: number) => Math.round(v * 100) / 100;
/** cached values for a total row of SUM formulas: the sum of the cells above (rows are 1-based Excel rows) */
function cacheSums(rows: C[][], total: number, first: number, last: number) {
  rows[total - 1].forEach((c, j) => {
    if (!c || typeof c !== 'object' || !c.f || c.v !== undefined) return;
    let t = 0;
    for (let i = first; i <= last; i++) {
      const x = rows[i - 1]?.[j];
      const v = x && typeof x === 'object' ? x.v : x;
      if (typeof v === 'number') t += v;
    }
    c.v = N(t);
  });
}
const q = (s: string) => `"${s.replace(/"/g, '""')}"`;

export async function revenueWorkbook(r: RevResult, meta: RevMeta): Promise<Uint8Array> {
  const title = (t: string): C[][] => heading(t, meta);
  const hdr = (cols: readonly string[]): C[] => cols.map(c => ({ v: c, s: 'hdr' }));
  const months = monthsOf(meta.fy);
  const mLabels = months.map(ymLabel);
  const B2C = REASON.B2C.label,
    NR = REASON.NOT_REPORTED.label,
    EXCL = REASON.EXCLUDED.label;

  // ---- Lists (drop-downs)
  const pick = REASONS.filter(x => x.code !== 'MATCHED');
  const lists: C[][] = [
    [
      { v: 'Reasons', s: 'hdr' },
      { v: 'Months', s: 'hdr' },
    ],
  ];
  for (let i = 0; i < Math.max(pick.length, mLabels.length); i++) lists.push([pick[i]?.label ?? null, mLabels[i] ?? null]);
  const reasonList = `Lists!$A$2:$A$${pick.length + 1}`,
    monthList = `Lists!$B$2:$B$${mLabels.length + 1}`;

  // ---- Rev_Items
  const items = [...r.items].sort(
    (a, b) =>
      REASONS.findIndex(x => x.code === a.reason) - REASONS.findIndex(x => x.code === b.reason) ||
      (a.bYm || a.gYm).localeCompare(b.bYm || b.gYm) ||
      (a.bNo || a.gNo).localeCompare(b.bNo || b.gNo),
  );
  const first = ITEMS_HDR + 1,
    last = Math.max(first, ITEMS_HDR + items.length);
  const rng = (name: (typeof ITEM_COLS)[number]) => `${ITEMS_SHEET}!$${col(IX(name))}$${first}:$${col(IX(name))}$${last}`;
  const L = (name: (typeof ITEM_COLS)[number]) => col(IX(name));
  const itemRows: C[][] = [
    ...title('Revenue reconciliation: documents, B2CS and Table 8 rows'),
    [
      {
        v: 'Fill the yellow columns: Your reason (drop-down), B2CS month (the month whose GSTR-1 B2CS includes the entry; blank = not in B2CS) and Remark. Save and drop the workbook into the Loader with the same files to carry them over.',
        s: 'note',
      },
    ],
    hdr(ITEM_COLS),
  ];
  items.forEach((it, k) => {
    const rr = ITEMS_HDR + 1 + k;
    const c = (n: (typeof ITEM_COLS)[number]) => `${L(n)}${rr}`;
    const used = `IF(${c('Your reason')}<>"",${c('Your reason')},IF(AND(${c('B2CS month')}<>"",OR(${c('Books taxable')}<>0,${c('Books tax')}<>0)),${q(B2C)},IF(AND(${c("Tool's reason")}=${q(B2C)},OR(${c('Books taxable')}<>0,${c('Books tax')}<>0)),${q(NR)},${c("Tool's reason")})))`;
    itemRows.push([
      it.id,
      it.status,
      it.party,
      it.gstin,
      it.type,
      it.dt,
      it.bNo,
      it.gNo,
      it.bDate ? { v: serial(it.bDate), s: 'date' } : '',
      ymLabel(it.bYm),
      it.gSection,
      ymLabel(it.gYm),
      it.hasB ? N(it.b[0]) : 0,
      it.hasB ? N(it.b[1]) : 0,
      N(it.g[0]),
      N(it.g[1]),
      { f: `${c('GSTR-1 taxable')}-${c('Books taxable')}`, v: N(it.g[0] - it.b[0]) },
      { f: `${c('GSTR-1 tax')}-${c('Books tax')}`, v: N(it.g[1] - it.b[1]) },
      REASON[it.derived].label,
      { v: it.userReason ? REASON[it.userReason].label : '', s: 'input' },
      { v: ymLabel(it.b2cMonth), s: 'input' },
      { f: used, v: REASON[it.reason].label },
      { v: it.remark, s: 'input' },
      it.note || (it.hasG && !it.g[0] && it.gAny[0] ? `Reported in another year's return: GSTR-1 taxable ${N(it.gAny[0])}` : ''),
      it.basis,
      it.src,
    ]);
  });
  const itemsSheet: SheetSpec = {
    name: ITEMS_SHEET,
    rows: itemRows,
    widths: [22, 22, 28, 17, 10, 8, 16, 16, 11, 9, 10, 9, 14, 13, 14, 13, 14, 13, 34, 34, 11, 34, 30, 40, 22, 26],
    freeze: { row: ITEMS_HDR, col: 3 },
    filter: { row: ITEMS_HDR, cols: ITEM_COLS.length },
    lists: [
      { sqref: `${L('Your reason')}${first}:${L('Your reason')}${last + 200}`, list: reasonList },
      { sqref: `${L('B2CS month')}${first}:${L('B2CS month')}${last + 200}`, list: monthList },
    ],
    tab: '2C388E',
  };

  // ---- Statement 1: books → GSTR-1
  const s1: C[][] = [
    ...title('Revenue reconciliation: books to GSTR-1'),
    hdr(['', 'Taxable value', 'Tax', 'Entries', 'What to do']),
  ];
  const S1_FIRST = s1.length + 1;
  const sumR = (c: (typeof ITEM_COLS)[number]) => `SUM(${rng(c)})`;
  const sifs = (c: (typeof ITEM_COLS)[number], label: string) => `SUMIFS(${rng(c)},${rng('Reason used')},${q(label)})`;
  const val = (code: string) => r.st1.find(l => l.code === code)?.amt ?? [0, 0];
  s1.push([
    { v: 'Revenue as per books', s: 'bold' },
    { f: sumR('Books taxable'), v: N(r.books.revenue[0]), s: 'numB' },
    { f: sumR('Books tax'), v: N(r.books.revenue[1]), s: 'numB' },
  ]);
  const rowRev = s1.length;
  s1.push([
    'Less: entries you marked as not a supply',
    { f: `-${sifs('Books taxable', EXCL)}`, v: N(-r.books.excluded[0]) },
    { f: `-${sifs('Books tax', EXCL)}`, v: N(-r.books.excluded[1]) },
    { f: `COUNTIFS(${rng('Reason used')},${q(EXCL)})`, v: r.items.filter(i => i.reason === 'EXCLUDED').length },
    REASON.EXCLUDED.action,
  ]);
  s1.push([
    { v: 'Turnover as per books', s: 'bold' },
    { f: `B${rowRev}+B${rowRev + 1}`, v: N(r.books.turnover[0]), s: 'numTot' },
    { f: `C${rowRev}+C${rowRev + 1}`, v: N(r.books.turnover[1]), s: 'numTot' },
  ]);
  const rowTurn = s1.length;
  s1.push([
    { v: 'Add / (less): reasons for the difference, from Rev_Items "Reason used"', s: 'sec' },
    { v: '', s: 'sec' },
    { v: '', s: 'sec' },
    { v: '', s: 'sec' },
    { v: '', s: 'sec' },
  ]);
  const lineFirst = s1.length + 1;
  for (const R of REASONS) {
    if (R.code === 'EXCLUDED' || R.code === 'MATCHED') continue;
    const v = val(R.code);
    s1.push([
      R.label,
      { f: `${sifs('GSTR-1 taxable', R.label)}-${sifs('Books taxable', R.label)}`, v: N(v[0]) },
      { f: `${sifs('GSTR-1 tax', R.label)}-${sifs('Books tax', R.label)}`, v: N(v[1]) },
      { f: `COUNTIFS(${rng('Reason used')},${q(R.label)})`, v: r.items.filter(i => i.reason === R.code).length },
      R.action,
    ]);
  }
  const lineLast = s1.length;
  s1.push([
    { v: 'Difference not explained', s: 'bold' },
    { f: `B${lineLast + 2}-B${rowTurn}-SUM(B${lineFirst}:B${lineLast})`, v: N(val('unx')[0]), s: 'numB' },
    { f: `C${lineLast + 2}-C${rowTurn}-SUM(C${lineFirst}:C${lineLast})`, v: N(val('unx')[1]), s: 'numB' },
    '',
    'Should be nil',
  ]);
  s1.push([
    { v: 'Turnover as per GSTR-1', s: 'bold' },
    { f: sumR('GSTR-1 taxable'), v: N(r.g1Turnover[0]), s: 'numTot' },
    { f: sumR('GSTR-1 tax'), v: N(r.g1Turnover[1]), s: 'numTot' },
  ]);
  const rowG1 = s1.length;
  s1.push([]);
  s1.push([
    {
      v: `Not in this statement: ${r.timing.count} matched invoice(s) booked in one month and reported in another month of the same year (taxable ${N(r.timing.amt[0])}). They cancel out over the year; see Month_wise.`,
      s: 'note',
    },
  ]);
  s1.push([{ v: 'Books months by invoice date; GSTR-1 months by return period. Credit notes are negative.', s: 'note' }]);
  void S1_FIRST;

  // ---- G1_Rows and 3B_Rows
  const g1Rows: C[][] = [
    ...title('GSTR-1 rows of the year (from the JSON files)'),
    hdr([
      'Key',
      '3B table',
      'Month',
      'Return FY',
      'Section',
      'Doc type',
      'Number',
      'Date',
      'Party',
      'POS',
      'Rate',
      'Taxable',
      'Tax',
      'File',
    ]),
  ];
  const NEG = new Set(['4I', '4L', '5H', '5K']);
  for (const x of meta.g1) {
    if (!x.key || String(x.info[0]) !== meta.fy) continue;
    const sg = NEG.has(x.key) ? -1 : 1;
    const t = tableOfKey(x.key);
    g1Rows.push([
      x.key,
      t ?? '',
      ymLabel(ymOfPeriod(String(x.info[1]))),
      String(x.info[0]),
      String(x.info[2]),
      String(x.info[3]),
      String(x.info[4]),
      String(x.info[5]),
      String(x.info[7]),
      String(x.info[8]),
      x.info[9] as any,
      N(sg * x.amt[0]),
      N(sg * (x.amt[1] + x.amt[2] + x.amt[3] + x.amt[4])),
      String(x.info[11]),
    ]);
  }
  const G1F = 5,
    G1L = Math.max(G1F, g1Rows.length);
  const g1r = (c: string) => `G1_Rows!$${c}$${G1F}:$${c}$${G1L}`;
  const tbRows: C[][] = [
    ...title('GSTR-3B Table 3.1 rows (from the PDFs)'),
    hdr(['Month', '3B table', 'Description', 'Taxable', 'Tax', 'File']),
  ];
  for (const t of meta.threeB) {
    const ym = ymOfPeriod(t.period);
    if (!months.includes(ym)) continue;
    for (const row of t.rows)
      if (TABLES.includes(row.key as Table))
        tbRows.push([
          ymLabel(ym),
          row.key,
          TABLE_LABEL[row.key as Table],
          N(row.amt[0]),
          N(row.amt[1] + row.amt[2] + row.amt[3] + row.amt[4]),
          String(row.info[4] ?? ''),
        ]);
  }
  const TBF = 5,
    TBL = Math.max(TBF, tbRows.length);
  const tbr = (c: string) => `'3B_Rows'!$${c}$${TBF}:$${c}$${TBL}`;

  // ---- Month_3B: per month, GSTR-1 per 3B table vs 3B
  const m3hdr: string[] = ['Month', '3B loaded'];
  for (const side of ['taxable', 'tax']) {
    m3hdr.push(`Advances (GSTR-1) ${side}`, `Reverse charge supplies (GSTR-1) ${side}`);
    for (const t of TABLES) m3hdr.push(`${t} GSTR-1 ${side}`, `${t} 3B ${side}`, `${t} diff (3B - GSTR-1) ${side}`);
  }
  m3hdr.push('GSTR-1 in tables + advances, taxable', 'GSTR-1 in tables + advances, tax');
  const m3: C[][] = [
    ...title('GSTR-1 vs GSTR-3B by table and month (diff = 3B - GSTR-1; 3.1(a) includes advances)'),
    [],
    hdr(m3hdr),
  ];
  const M3F = m3.length + 1;
  const W = 2 + TABLES.length * 3; // columns per side
  r.months.forEach((m, k) => {
    const rr = M3F + k;
    const row: C[] = [ymLabel(m.ym), { f: `IF(COUNTIFS(${tbr('A')},A${rr})>0,"Yes","No")`, v: m.tbt ? 'Yes' : 'No' }];
    (['taxable', 'tax'] as const).forEach((_, si) => {
      const g1c = si ? 'M' : 'L',
        tbc = si ? 'E' : 'D';
      const base = 2 + si * W;
      row.push(
        { f: `SUMIFS(${g1r(g1c)},${g1r('B')},"ADV",${g1r('C')},$A${rr})`, v: N(m.adv[si]) },
        { f: `SUMIFS(${g1r(g1c)},${g1r('B')},"RCM",${g1r('C')},$A${rr})`, v: N(m.rcm[si]) },
      );
      TABLES.forEach((t, ti) => {
        const cg = col(base + 2 + ti * 3),
          ct = col(base + 3 + ti * 3),
          adv = col(base);
        const dv = m.tbt ? m.tbt[t][si] - m.g1t[t][si] - (t === '3.1a' ? m.adv[si] : 0) : 0;
        row.push(
          { f: `SUMIFS(${g1r(g1c)},${g1r('B')},${q(t)},${g1r('C')},$A${rr})`, v: N(m.g1t[t][si]) },
          { f: `SUMIFS(${tbr(tbc)},${tbr('B')},${q(t)},${tbr('A')},$A${rr})`, v: m.tbt ? N(m.tbt[t][si]) : 0 },
          { f: `IF($B${rr}="Yes",${ct}${rr}-${cg}${rr}${t === '3.1a' ? `-${adv}${rr}` : ''},0)`, v: N(dv) },
        );
      });
    });
    const tot = (si: number) => {
      const base = 2 + si * W;
      return [col(base), ...TABLES.map((_, ti) => col(base + 2 + ti * 3))].map(c => `${c}${rr}`).join('+');
    };
    row.push(
      { f: tot(0), v: N(TABLES.reduce((a, t) => a + m.g1t[t][0], m.adv[0])) },
      { f: tot(1), v: N(TABLES.reduce((a, t) => a + m.g1t[t][1], m.adv[1])) },
    );
    m3.push(row);
  });
  const M3L = M3F + r.months.length - 1;
  const m3tot: C[] = [{ v: 'Total', s: 'bold' }, ''];
  for (let j = 2; j < m3hdr.length; j++) m3tot.push({ f: `SUM(${col(j)}${M3F}:${col(j)}${M3L})`, s: 'numTot' });
  m3.push(m3tot);
  const M3T = m3.length;
  cacheSums(m3, M3T, M3F, M3L);
  const m3c = (side: number, what: 'adv' | 'rcm' | 'g1' | 'tb' | 'diff', t?: Table) => {
    const base = 2 + side * W;
    if (what === 'adv') return col(base);
    if (what === 'rcm') return col(base + 1);
    const ti = TABLES.indexOf(t!);
    return col(base + 2 + ti * 3 + (what === 'g1' ? 0 : what === 'tb' ? 1 : 2));
  };
  const totCol = (si: number) => col(2 + 2 * W + si);

  // ---- Statement 2: GSTR-1 → 3B
  const s2: C[][] = [
    ...title('Revenue reconciliation: GSTR-1 to GSTR-3B'),
    hdr(['', 'Taxable value', 'Tax', 'Months', 'What to do']),
  ];
  const v2 = (code: string) => r.st2.find(l => l.code === code)?.amt ?? [0, 0];
  s2.push([
    { v: 'Turnover as per GSTR-1', s: 'bold' },
    { f: `Books_to_GSTR1!B${rowG1}`, v: N(r.g1Turnover[0]), s: 'numB' },
    { f: `Books_to_GSTR1!C${rowG1}`, v: N(r.g1Turnover[1]), s: 'numB' },
  ]);
  const s2Start = s2.length;
  const allT = (si: number) => TABLES.map(t => `Month_3B!${m3c(si, 'g1', t)}${M3T}`).join('+');
  s2.push([
    'Less: supplies on reverse charge (tax paid by the recipient, not in 3B 3.1)',
    { f: `-Month_3B!${m3c(0, 'rcm')}${M3T}`, v: N(v2('rcm')[0]) },
    { f: `-Month_3B!${m3c(1, 'rcm')}${M3T}`, v: N(v2('rcm')[1]) },
    '',
    'Not reported in 3B by the supplier',
  ]);
  s2.push([
    'Add: advances received, net of advances adjusted',
    { f: `Month_3B!${m3c(0, 'adv')}${M3T}`, v: N(v2('adv')[0]) },
    { f: `Month_3B!${m3c(1, 'adv')}${M3T}`, v: N(v2('adv')[1]) },
    '',
    'In 3B 3.1(a), not turnover (GSTR-1 Tables 11A / 11B)',
  ]);
  s2.push([
    'Less: GSTR-1 documents not counted in a 3B table',
    { f: `-(B${s2Start}+B${s2Start + 1}-(${allT(0)}))`, v: N(v2('nontable')[0]) },
    { f: `-(C${s2Start}+C${s2Start + 1}-(${allT(1)}))`, v: N(v2('nontable')[1]) },
    '',
    'Should be nil; otherwise amendments whose original is not loaded',
  ]);
  s2.push([
    'Less: GSTR-1 of months with no 3B loaded',
    { f: `-SUMIFS(Month_3B!${totCol(0)}${M3F}:${totCol(0)}${M3L},Month_3B!B${M3F}:B${M3L},"No")`, v: N(v2('notb')[0]) },
    { f: `-SUMIFS(Month_3B!${totCol(1)}${M3F}:${totCol(1)}${M3L},Month_3B!B${M3F}:B${M3L},"No")`, v: N(v2('notb')[1]) },
    { f: `COUNTIFS(Month_3B!B${M3F}:B${M3L},"No")`, v: r.months.filter(m => !m.tbt).length },
    'Add those GSTR-3B PDFs to compare them',
  ]);
  for (const t of TABLES) {
    const l = r.st2.find(x => x.code === 'T' + t);
    s2.push([
      `${TABLE_LABEL[t]}: 3B more / (less) than GSTR-1`,
      { f: `Month_3B!${m3c(0, 'diff', t)}${M3T}`, v: N(l?.amt[0] ?? 0) },
      { f: `Month_3B!${m3c(1, 'diff', t)}${M3T}`, v: N(l?.amt[1] ?? 0) },
      { f: `COUNTIFS(Month_3B!${m3c(0, 'diff', t)}${M3F}:${m3c(0, 'diff', t)}${M3L},"<>0")`, v: l?.count ?? 0 },
      l?.action ?? '',
    ]);
  }
  const s2Last = s2.length;
  const tbSum = (si: number) => TABLES.map(t => `Month_3B!${m3c(si, 'tb', t)}${M3T}`).join('+');
  s2.push([
    { v: 'Difference not explained', s: 'bold' },
    { f: `B${s2Last + 2}-SUM(B${s2Start}:B${s2Last})`, v: N(v2('unx')[0]), s: 'numB' },
    { f: `C${s2Last + 2}-SUM(C${s2Start}:C${s2Last})`, v: N(v2('unx')[1]), s: 'numB' },
    '',
    'Should be nil',
  ]);
  s2.push([
    { v: 'Turnover as per GSTR-3B (Table 3.1 and 3.1.1(ii))', s: 'bold' },
    { f: tbSum(0), v: N(r.tbTotal[0]), s: 'numTot' },
    { f: tbSum(1), v: N(r.tbTotal[1]), s: 'numTot' },
  ]);
  const rowTb = s2.length;
  s2.push([]);
  s2.push([{ v: 'Month-to-month differences within the year cancel out here; see Month_3B for each month.', s: 'note' }]);

  // ---- Month_wise
  const mw: C[][] = [
    ...title('Month-wise: books vs GSTR-1 vs GSTR-3B'),
    hdr([
      'Month',
      'Books taxable',
      'GSTR-1 taxable',
      '3B taxable',
      'GSTR-1 - books',
      '3B - GSTR-1',
      'Running: GSTR-1 - books',
      'Running: 3B - GSTR-1',
      'Books tax',
      'GSTR-1 tax',
      '3B tax',
      'GSTR-1 - books (tax)',
      '3B - GSTR-1 (tax)',
      'Running (tax): GSTR-1 - books',
      'Running (tax): 3B - GSTR-1',
    ]),
  ];
  const MWF = mw.length + 1;
  const runs = [0, 0, 0, 0];
  r.months.forEach((m, k) => {
    const rr = MWF + k,
      m3r = M3F + k;
    const blk = (si: number, c0: number): C[] => {
      const B = col(c0),
        G = col(c0 + 1),
        T = col(c0 + 2),
        D1 = col(c0 + 3),
        D2 = col(c0 + 4);
      const bc = si ? 'Books tax' : 'Books taxable',
        gc = si ? 'GSTR-1 tax' : 'GSTR-1 taxable';
      const tb = TABLES.map(t => `Month_3B!${m3c(si, 'tb', t)}${m3r}`).join('+');
      runs[si * 2] += m.g1[si] - m.books[si];
      runs[si * 2 + 1] += m.tb ? m.tb[si] - m.g1[si] : 0;
      return [
        { f: `SUMIFS(${rng(bc)},${rng('Books month')},$A${rr},${rng('Reason used')},"<>"&${q(EXCL)})`, v: N(m.books[si]) },
        { f: `SUMIFS(${rng(gc)},${rng('GSTR-1 month')},$A${rr})`, v: N(m.g1[si]) },
        { f: `IF(Month_3B!$B${m3r}="Yes",${tb},0)`, v: N(m.tb ? m.tb[si] : 0) },
        { f: `${G}${rr}-${B}${rr}`, v: N(m.g1[si] - m.books[si]) },
        { f: `IF(Month_3B!$B${m3r}="Yes",${T}${rr}-${G}${rr},0)`, v: N(m.tb ? m.tb[si] - m.g1[si] : 0) },
        { f: k ? `${col(c0 + 5)}${rr - 1}+${D1}${rr}` : `${D1}${rr}`, v: N(runs[si * 2]) },
        { f: k ? `${col(c0 + 6)}${rr - 1}+${D2}${rr}` : `${D2}${rr}`, v: N(runs[si * 2 + 1]) },
      ];
    };
    mw.push([ymLabel(m.ym), ...blk(0, 1), ...blk(1, 8)]);
  });
  const MWL = MWF + r.months.length - 1;
  mw.push([
    { v: 'Total', s: 'bold' },
    ...[1, 2, 3, 4, 5, '', '', 8, 9, 10, 11, 12].map(j =>
      typeof j === 'number' ? { f: `SUM(${col(j)}${MWF}:${col(j)}${MWL})`, s: 'numTot' as const } : '',
    ),
  ]);
  cacheSums(mw, mw.length, MWF, MWL);
  mw.push([]);
  mw.push([{ v: 'Books: by invoice date, without entries marked as not a supply. GSTR-1 and 3B: by return period.', s: 'note' }]);

  // ---- B2C_Month
  const b2: C[][] = [
    ...title('B2C (small) sales: books entries ticked for each B2CS month vs GSTR-1 B2CS'),
    hdr(['B2CS month', 'Books taxable', 'GSTR-1 B2CS taxable', 'Difference', 'Books tax', 'GSTR-1 B2CS tax', 'Difference (tax)']),
  ];
  const B2F = b2.length + 1;
  r.months.forEach((m, k) => {
    const rr = B2F + k;
    const s = (c: (typeof ITEM_COLS)[number]) => `SUMIFS(${rng(c)},${rng('Reason used')},${q(B2C)},${rng('B2CS month')},$A${rr})`;
    b2.push([
      ymLabel(m.ym),
      { f: s('Books taxable'), v: N(m.b2cBooks[0]) },
      { f: s('GSTR-1 taxable'), v: N(m.b2cG1[0]) },
      { f: `C${rr}-B${rr}`, v: N(m.b2cG1[0] - m.b2cBooks[0]) },
      { f: s('Books tax'), v: N(m.b2cBooks[1]) },
      { f: s('GSTR-1 tax'), v: N(m.b2cG1[1]) },
      { f: `F${rr}-E${rr}`, v: N(m.b2cG1[1] - m.b2cBooks[1]) },
    ]);
  });
  const B2L = B2F + r.months.length - 1;
  b2.push([
    { v: 'Total', s: 'bold' },
    ...[1, 2, 3, 4, 5, 6].map(j => ({ f: `SUM(${col(j)}${B2F}:${col(j)}${B2L})`, s: 'numTot' as const })),
  ]);
  cacheSums(b2, b2.length, B2F, B2L);
  b2.push([]);
  b2.push([
    {
      v: 'GSTR-1 B2CS has no invoice numbers. In Rev_Items, set "B2CS month" on the books entries that were reported in a month\'s B2CS (B2C entries start with their invoice month; clear it if an entry was not reported). An entry moved to another month only changes the month view; one not reported at all goes to "In books, not reported in GSTR-1".',
      s: 'note',
    },
  ]);

  // ---- Dashboard
  const dash: C[][] = [
    ...title('Summary: turnover as per books, GSTR-1 and GSTR-3B'),
    hdr(['', 'Taxable value', 'Tax', 'How it is worked out']),
    [
      'Turnover as per books',
      { f: `Books_to_GSTR1!B${rowTurn}`, v: N(r.books.turnover[0]) },
      { f: `Books_to_GSTR1!C${rowTurn}`, v: N(r.books.turnover[1]) },
      how(`Books_to_GSTR1 row ${rowTurn}: books taxable value and tax of every Rev_Items entry, less entries marked "not a supply".`),
    ],
    [
      'Turnover as per GSTR-1',
      { f: `Books_to_GSTR1!B${rowG1}`, v: N(r.g1Turnover[0]) },
      { f: `Books_to_GSTR1!C${rowG1}`, v: N(r.g1Turnover[1]) },
      how(`Books_to_GSTR1 row ${rowG1}: sum of the GSTR-1 columns of Rev_Items (invoices, notes, B2CS, Table 8).`),
    ],
    [
      'Turnover as per GSTR-3B',
      { f: `GSTR1_to_3B!B${rowTb}`, v: N(r.tbTotal[0]) },
      { f: `GSTR1_to_3B!C${rowTb}`, v: N(r.tbTotal[1]) },
      how(`GSTR1_to_3B row ${rowTb}: Table 3.1 (a), (b), (c), (e) and 3.1.1(ii) of the 3B PDFs, months in Month_3B.`),
    ],
    [
      { v: 'GSTR-1 - books', s: 'bold' },
      { f: 'B6-B5', s: 'numB', v: N(r.g1Turnover[0] - r.books.turnover[0]) },
      { f: 'C6-C5', s: 'numB', v: N(r.g1Turnover[1] - r.books.turnover[1]) },
      how('Explained line by line in Books_to_GSTR1.'),
    ],
    [
      { v: '3B - GSTR-1', s: 'bold' },
      { f: 'B7-B6', s: 'numB', v: N(r.tbTotal[0] - r.g1Turnover[0]) },
      { f: 'C7-C6', s: 'numB', v: N(r.tbTotal[1] - r.g1Turnover[1]) },
      how('Explained table by table in GSTR1_to_3B.'),
    ],
    [],
    [
      { v: 'Main reasons (books to GSTR-1)', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
    ],
    ...r.st1
      .filter(l => l.kind === 'line')
      .sort((a, b) => Math.abs(b.amt[0]) - Math.abs(a.amt[0]))
      .map(l => {
        const i = s1.findIndex(x => x[0] === REASON[l.code as keyof typeof REASON]?.label);
        return [
          l.label.replace(/^(Add|Less): /, ''),
          i >= 0 ? { f: `Books_to_GSTR1!B${i + 1}`, v: N(l.amt[0]) } : N(l.amt[0]),
          i >= 0 ? { f: `Books_to_GSTR1!C${i + 1}`, v: N(l.amt[1]) } : N(l.amt[1]),
          how(i >= 0 ? `Books_to_GSTR1 row ${i + 1}` : ''),
        ] as C[];
      }),
    [],
    [
      { v: 'Unexplained (books to GSTR-1)', s: 'bold' },
      { f: `Books_to_GSTR1!B${lineLast + 1}`, v: N(val('unx')[0]), s: 'numB' },
      { f: `Books_to_GSTR1!C${lineLast + 1}`, v: N(val('unx')[1]), s: 'numB' },
      how('Should be nil'),
    ],
    [
      { v: 'Unexplained (GSTR-1 to 3B)', s: 'bold' },
      { f: `GSTR1_to_3B!B${rowTb - 1}`, v: N(v2('unx')[0]), s: 'numB' },
      { f: `GSTR1_to_3B!C${rowTb - 1}`, v: N(v2('unx')[1]), s: 'numB' },
      how('Should be nil'),
    ],
  ];

  // ---- Books_Lines
  const bl: C[][] = [
    ...title(`Books lines used (${meta.source})`),
    hdr([
      'Date',
      'Doc no.',
      'Customer',
      'GSTIN',
      'POS',
      'Type',
      'Doc type',
      'HSN',
      'Rate',
      'Taxable',
      'CGST',
      'SGST',
      'IGST',
      'Cess',
      'Ledgers',
      'Ledger class',
      'Voucher type',
      'Source rows',
    ]),
  ];
  for (const x of meta.bookRows)
    bl.push([
      typeof x[0] === 'number' ? { v: x[0], s: 'date' } : x[0],
      x[1],
      x[2],
      x[3],
      x[4],
      x[5],
      x[6],
      x[7],
      x[10],
      x[11],
      x[12],
      x[13],
      x[14],
      x[15],
      x[16] ?? '',
      x[17] ? (REV_CLASS_LABEL[x[17] as RevClass] ?? x[17]) : '',
      x[18] ?? '',
      x[19] ?? '',
    ]);

  // ---- Instructions (one sheet) and Index
  const front = {
    title: 'Revenue reconciliation: books vs GSTR-1 vs GSTR-3B',
    legalName: meta.legalName,
    gstin: meta.gstin,
    fy: meta.fy,
    builtAt: meta.builtAt,
    version: LOADER_VERSION,
    extra: [['Books source', meta.source]] as [string, string][],
  };
  const sections: Section[] = [
    {
      heading: 'What this workpaper does',
      lines: [
        ['Purpose', 'Reconciles turnover as per books with GSTR-1, and GSTR-1 with GSTR-3B, for the year, giving the reason for every difference.'],
        ['Start here', 'Summary, then Books_to_GSTR1 and GSTR1_to_3B (bank-reco style statements: start figure, reasons, end figure).'],
      ],
    },
    {
      heading: 'How the figures are worked out',
      lines: [
        ['Rev_Items', 'One row per invoice, credit note, B2CS month / rate / place and Table 8 row, with the books side and the GSTR-1 side. "Reason used" = your reason if you gave one, else the tool\'s reason.'],
        ['Books_to_GSTR1', 'Each reason line = SUMIFS of the GSTR-1 columns minus SUMIFS of the books columns of Rev_Items with that "Reason used". So books + reasons = GSTR-1; "Difference not explained" is what is left and should be nil.'],
        ['GSTR1_to_3B', 'Links to the totals row of Month_3B, which compares GSTR-1 and 3B month by month and table by table (SUMIFS over G1_Rows and 3B_Rows).'],
        ['Month_wise / Month_3B', 'The same by month. Books are by invoice date, GSTR-1 / 3B by return period: shifts between months of the same year cancel out over the year and show only here.'],
        ['B2C_Month', 'Books B2C entries ticked for each month against GSTR-1 B2CS.'],
        ['G1_Rows / 3B_Rows / Books_Lines', 'The data read from the files, untouched.'],
      ],
    },
    {
      heading: 'Reasons (what to do about each)',
      lines: REASONS.filter(x => x.code !== 'MATCHED').map(x => [x.label, x.action] as [string, string]),
    },
    {
      heading: 'Working with the tool',
      lines: [
        ['Your reasons in Excel', 'Fill the yellow columns of Rev_Items (Your reason, B2CS month, Remark) and save. The statements move with them.'],
        ['Next time', 'Drop this saved workpaper into the Loader (Revenue reco tab): it brings back its files, settings and reasons; your yellow columns are read back by Id. A newer sales register or GL dump dropped with it replaces the saved copy; GSTR-1 and 3B stay as filed.'],
        ['Annual return', 'This workpaper, with the Input Reco and GST ledger flow workpapers, is what the Annual return tab builds GSTR-9 / 9C from.'],
      ],
    },
    { heading: 'Settings used', lines: (meta.settings ?? []).map(t => ['Matching settings', t] as [string, string]) },
    { heading: 'Files used', lines: meta.files.map(f => [f.name, `${f.kind}${f.note ? ' · ' + f.note : ''}`] as [string, string]) },
    {
      heading: 'Where to get the files',
      lines: REV_GUIDE.map(
        g =>
          [
            `${g.title} (${g.source === 'portal' ? 'GST portal' : g.source === 'tally' ? 'Tally' : g.source === 'zoho' ? 'Zoho Books' : 'Other'})`,
            g.steps.join(' ') + (g.note ? ' ' + g.note : ''),
          ] as [string, string],
      ),
    },
  ];
  const info: SheetInfo[] = [
    { name: 'Summary', role: 'Summary', about: 'Turnover as per books, GSTR-1 and GSTR-3B, the differences and the main reasons (links to the statements).' },
    { name: 'Books_to_GSTR1', role: 'Working', about: 'Revenue as per books, each reason for the difference, turnover as per GSTR-1 (SUMIFS over Rev_Items).' },
    { name: 'GSTR1_to_3B', role: 'Working', about: 'Turnover as per GSTR-1 to GSTR-3B Table 3.1, table by table for the year.' },
    { name: 'Month_wise', role: 'Working', about: 'Books, GSTR-1 and GSTR-3B by month.' },
    { name: 'B2C_Month', role: 'Working', about: 'Books B2C entries for each month against GSTR-1 B2CS.' },
    { name: 'Month_3B', role: 'Working', about: 'GSTR-1 against GSTR-3B by table and month.' },
    { name: ITEMS_SHEET, role: 'Your input', about: 'Every document, B2CS and Table 8 row with both sides; yellow columns are yours (reason, B2CS month, remark).' },
    ...(meta.customers?.length
      ? [{ name: 'Customer_Map', role: 'Working', about: 'Books customers without a GSTIN, the GSTIN given and the basis.' } as SheetInfo]
      : []),
    { name: 'G1_Rows', role: 'Source data', about: 'GSTR-1 rows of the year, from the JSON files.' },
    { name: '3B_Rows', role: 'Source data', about: 'GSTR-3B Table 3.1 rows, from the PDFs.' },
    { name: 'Books_Lines', role: 'Source data', about: `Books lines used (${meta.source}).` },
    ...(meta.ledgers?.length
      ? [{ name: 'Ledgers', role: 'Working', about: 'Revenue ledgers in the GL dump and the class you gave each.' } as SheetInfo]
      : []),
    { name: 'Audit', role: 'Your input', about: 'Your decisions: reasons, B2CS months, ledger classes.' },
    { name: 'Lists', role: 'Reference', about: 'The drop-down lists used in Rev_Items.' },
  ];
  const sheets: SheetSpec[] = [
    indexSheet(front, info),
    instructionsSheet(front, sections),
    { name: 'Summary', rows: dash, widths: [52, 18, 16, 90], freeze: { row: 4 }, tab: '2C388E' },
    { name: 'Books_to_GSTR1', rows: s1, widths: [60, 18, 16, 9, 80], freeze: { row: 4 } },
    { name: 'GSTR1_to_3B', rows: s2, widths: [70, 18, 16, 9, 70], freeze: { row: 4 } },
    { name: 'Month_wise', rows: mw, widths: [10, ...Array(14).fill(15)], freeze: { row: 4, col: 1 } },
    { name: 'B2C_Month', rows: b2, widths: [12, 16, 18, 15, 14, 16, 15], freeze: { row: 4, col: 1 } },
    { name: 'Month_3B', rows: m3, widths: [10, 9, ...Array(m3hdr.length - 2).fill(14)], freeze: { row: 5, col: 2 } },
    itemsSheet,
    ...(meta.customers?.length
      ? [
          {
            name: 'Customer_Map',
            rows: [
              ...title('Customer GSTIN for books entries without one'),
              [],
              ['Customer in books', 'GSTIN given', 'Basis', 'Evidence'].map(v => ({ v, s: 'hdr' as const })),
              ...meta.customers.map(
                x => [x.name, x.gstin || 'Unregistered (B2C)', x.basis, { v: x.detail, s: 'wrap' as const }] as C[],
              ),
            ],
            widths: [40, 22, 26, 100],
            freeze: { row: 4 },
          } as SheetSpec,
        ]
      : []),
    {
      name: 'G1_Rows',
      rows: g1Rows,
      widths: [7, 8, 8, 9, 10, 8, 18, 11, 18, 6, 6, 14, 13, 40],
      freeze: { row: 4 },
      filter: { row: 4, cols: 14 },
    },
    { name: '3B_Rows', rows: tbRows, widths: [8, 8, 50, 14, 13, 40], freeze: { row: 4 }, filter: { row: 4, cols: 6 } },
    {
      name: 'Books_Lines',
      rows: bl,
      widths: [11, 16, 28, 17, 5, 10, 8, 9, 6, 14, 12, 12, 12, 10, 30, 22, 12, 30],
      freeze: { row: 4 },
      filter: { row: 4, cols: 18 },
    },
  ];
  if (meta.ledgers?.length)
    sheets.push({
      name: 'Ledgers',
      rows: [
        ...title('Revenue ledgers in the GL dump and what you said each one is'),
        hdr(['Ledger', 'Class', 'Lines', 'Debit', 'Credit', 'Net credit']),
        ...meta.ledgers.map(
          l => [l.name, l.cls ? REV_CLASS_LABEL[l.cls] : 'Not reviewed: left out', l.lines, l.dr, l.cr, N(l.cr - l.dr)] as C[],
        ),
      ],
      widths: [40, 34, 8, 16, 16, 16],
      freeze: { row: 4 },
    });
  sheets.push({
    name: 'Audit',
    rows: [
      ...title('Your decisions (reasons, B2CS months, ledger classes)'),
      hdr(['When', 'Change', 'Record', 'From', 'To', 'Remark']),
      ...meta.audit.map(a => [a.when, a.change, a.record, a.from, a.to, a.remark] as C[]),
    ],
    widths: [17, 22, 40, 30, 30, 40],
    freeze: { row: 4 },
  });
  sheets.push({ name: 'Lists', rows: lists, widths: [50, 10] });
  return writeBook(sheets);
}

/** a saved revenue workbook: the user's columns of Rev_Items, by Id */
export function readItemsSheet(rows: any[][]): { id: string; reason: string; b2c: string; remark: string }[] {
  const h = rows.findIndex(r => r && String(r[0]) === 'Id' && r.includes('Your reason'));
  if (h < 0) return [];
  const H = rows[h].map((x: any) => String(x ?? ''));
  const ci = (n: string) => H.indexOf(n);
  const out: { id: string; reason: string; b2c: string; remark: string }[] = [];
  for (const r of rows.slice(h + 1)) {
    if (!r || !r[0]) continue;
    out.push({
      id: String(r[0]),
      reason: String(r[ci('Your reason')] ?? '').trim(),
      b2c: String(r[ci('B2CS month')] ?? '').trim(),
      remark: String(r[ci('Remark')] ?? '').trim(),
    });
  }
  return out;
}
export type { Item };

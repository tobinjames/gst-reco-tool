// Revenue reconciliation: books (sales register, or revenue ledgers from a GL dump) vs GSTR-1 vs GSTR-3B for one FY.
// Invoice matching is recon.ts reconSales(); this module turns its pairs (plus B2CS, Table 8 and e-commerce rows, which
// have no invoices) into one list of items, each with a books side and a GSTR-1 side counted for the year, and a reason.
// The statements are sums of those items by reason, so they always tie: books + reasons = GSTR-1.
// Months: books by invoice date, GSTR-1 / 3B by return period. Shifts between months of the same year cancel out in the
// year statement and show only in the month-wise view.
import { DEFAULT_MATCH, KeepApart, Manual, MatchOptions, normNo, reconSales, toDate } from '../recon';
import type { ThreeB } from '../gstr3b';
import { RawRow, fyOfDate, fyOfPeriod, fyShift, r2 } from '../util';

export type Ym = string; // YYYY-MM
export const MON3 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export const ymLabel = (ym: Ym) => (ym ? `${MON3[+ym.slice(5) - 1]}-${ym.slice(2, 4)}` : '');
export const ymOfPeriod = (p: string): Ym => (p && /^\d{6}$/.test(p) ? `${p.slice(2)}-${p.slice(0, 2)}` : '');
export const ymOfDate = (d: string): Ym => {
  const m = String(d).match(/^(\d{2})-(\d{2})-(\d{4})$/);
  return m ? `${m[3]}-${m[2]}` : '';
};
export const monthsOf = (fy: string): Ym[] =>
  [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].map(
    m => `${m >= 4 ? fy.slice(0, 4) : +fy.slice(0, 4) + 1}-${String(m).padStart(2, '0')}`,
  );

// ---------------------------------------------------------------- reasons
export type ReasonCode =
  | 'NOT_REPORTED'
  | 'NEXT_FY'
  | 'NOT_IN_BOOKS'
  | 'PREV_FY'
  | 'VALUE'
  | 'B2C'
  | 'NIL'
  | 'ECO'
  | 'CANCELLED'
  | 'DUPLICATE'
  | 'OTHER_NO'
  | 'AMEND'
  | 'BOOK_ERR'
  | 'OTHER'
  | 'ROUND'
  | 'EXCLUDED'
  | 'MATCHED';
export interface Reason {
  code: ReasonCode;
  label: string;
  /** what to do about it (shown under the statement line and in Excel) */
  action: string;
  /** the tool sets it from the documents (the user can still pick it) */
  derived?: boolean;
}
/** statement order; the user can pick any of these except Matched */
export const REASONS: Reason[] = [
  {
    code: 'NOT_REPORTED',
    label: 'In books, not reported in GSTR-1',
    action: 'Report in the next GSTR-1 (Table 10 of GSTR-9 if after the year) or pay through DRC-03',
    derived: true,
  },
  {
    code: 'NEXT_FY',
    label: "Reported in next year's GSTR-1",
    action: "Included in the next year's returns: shows in GSTR-9 Table 10, tax in Table 14",
    derived: true,
  },
  {
    code: 'PREV_FY',
    label: "Previous year's invoice reported in this year's GSTR-1",
    action: "Belongs to last year's books: part of last year's GSTR-9 Table 10",
    derived: true,
  },
  {
    code: 'NOT_IN_BOOKS',
    label: 'In GSTR-1, not in books',
    action: 'Check for an invoice missing in books, or one wrongly reported',
    derived: true,
  },
  {
    code: 'VALUE',
    label: 'Value differs between books and GSTR-1',
    action: 'Amend in GSTR-1 or correct the books',
    derived: true,
  },
  {
    code: 'B2C',
    label: 'B2C (small) sales: books vs GSTR-1 B2CS',
    action: 'GSTR-1 B2CS has no invoices: tick books entries against the B2CS month they were reported in',
    derived: true,
  },
  {
    code: 'NIL',
    label: 'Nil-rated / exempt / non-GST: books vs GSTR-1 Table 8',
    action: 'Check the Table 8 figures in GSTR-1',
    derived: true,
  },
  {
    code: 'ECO',
    label: 'Supplies through e-commerce operators u/s 9(5)',
    action: 'Reported in GSTR-1 Table 14: tax paid by the operator',
    derived: true,
  },
  { code: 'CANCELLED', label: 'Cancelled invoice', action: 'Cancelled in books or on the portal: no action if both agree' },
  { code: 'DUPLICATE', label: 'Duplicate entry', action: 'Remove the duplicate' },
  { code: 'OTHER_NO', label: 'Reported under another invoice no.', action: 'Match it by hand or amend the number' },
  { code: 'AMEND', label: 'Amendment to be made in GSTR-1', action: 'Amend in the next GSTR-1' },
  { code: 'BOOK_ERR', label: 'Booking error in books', action: 'Correct the books' },
  { code: 'OTHER', label: 'Other reason (see remark)', action: '' },
  { code: 'ROUND', label: 'Rounding (within tolerance)', action: '', derived: true },
  {
    code: 'EXCLUDED',
    label: 'Not a supply (excluded by you)',
    action: 'Left out of turnover: interest, dividends, forex gains and the like',
  },
  { code: 'MATCHED', label: 'Matched', action: '', derived: true },
];
export const REASON = Object.fromEntries(REASONS.map(r => [r.code, r])) as Record<ReasonCode, Reason>;
export const reasonOfLabel = (s: string): ReasonCode | '' =>
  REASONS.find(
    r =>
      r.label.toLowerCase() ===
      String(s ?? '')
        .trim()
        .toLowerCase(),
  )?.code ?? '';

// ---------------------------------------------------------------- user decisions
export interface UserState {
  /** item id (or one of its document ids) → reason picked by the user */
  reasons: Record<string, { r?: ReasonCode; note?: string }>;
  /** books document id → B2CS month it was reported in ('' = not in B2CS) */
  b2c: Record<string, Ym | ''>;
}
export const emptyUser = (): UserState => ({ reasons: {}, b2c: {} });

// ---------------------------------------------------------------- items
export type ItemKind = 'doc' | 'b2c' | 'nil' | 'eco' | 'excl';
export interface Item {
  id: string;
  kind: ItemKind;
  status: string;
  basis: string;
  party: string;
  gstin: string;
  type: string;
  dt: string;
  bNo: string;
  gNo: string;
  bDate: string;
  gDate: string;
  bYm: Ym;
  gYm: Ym;
  gSection: string;
  /** books side [taxable, tax], signed (credit notes negative); 0 when not in this year's books */
  b: [number, number];
  /** GSTR-1 side [taxable, tax], signed, counted only from this year's returns */
  g: [number, number];
  /** GSTR-1 side as reported, whatever the year (for next-year returns) */
  gAny: [number, number];
  hasB: boolean;
  hasG: boolean;
  derived: ReasonCode;
  /** default B2CS month (books B2C documents: their invoice month) */
  b2cDefault: Ym | '';
  /** B2CS month in use (user's, else default) */
  b2cMonth: Ym | '';
  reason: ReasonCode;
  userReason: ReasonCode | '';
  remark: string;
  note: string;
  /** books ids / return ids behind the item */
  bIds: string[];
  rIds: string[];
  src: string;
}

export interface BooksIn {
  /** Books_Sales layout rows: A date, B no, C customer, D GSTIN, E POS, F type, G doc type, H HSN, I UQC, J qty, K rate,
   * L taxable, M CGST, N SGST, O IGST, P cess; GL rows add Q ledgers, R class, S voucher type, T source rows */
  rows: any[][];
  source: 'register' | 'gl' | 'none';
}
export interface RevInputs {
  fy: string;
  books: BooksIn;
  /** processGstr1() rows of every loaded GSTR-1 (any year) */
  g1: RawRow[];
  threeB: ThreeB[];
  user: UserState;
  match?: MatchOptions;
  manual?: Manual;
  keepApart?: KeepApart;
}

// ---------------------------------------------------------------- 3B tables
export type Table = '3.1a' | '3.1b' | '3.1c' | '3.1e' | '3.1.1ii';
export const TABLES: Table[] = ['3.1a', '3.1b', '3.1c', '3.1e', '3.1.1ii'];
export const TABLE_LABEL: Record<Table, string> = {
  '3.1a': '3.1(a) Taxable supplies (other than zero rated, nil, exempt)',
  '3.1b': '3.1(b) Zero rated supplies (exports, SEZ)',
  '3.1c': '3.1(c) Nil-rated and exempt supplies',
  '3.1e': '3.1(e) Non-GST supplies',
  '3.1.1ii': '3.1.1(ii) Supplies through e-commerce operators u/s 9(5)',
};
/** GSTR-9 key of a GSTR-1 row (gstr1.ts) → 3B table; null = not in 3B table 3.1 */
export function tableOfKey(key: string): Table | 'ADV' | 'RCM' | null {
  if (['4A', '4B', '4E', '4I', '4J', '4K', '4L'].includes(key)) return '3.1a';
  if (key === '4F') return 'ADV';
  if (['4C', '4D', '5A', '5B', '5H', '5I', '5J', '5K'].includes(key)) return '3.1b';
  if (key === '5D' || key === '5E') return '3.1c';
  if (key === '5F') return '3.1e';
  if (key === '5C1') return '3.1.1ii';
  if (key === '5C') return 'RCM';
  return null;
}
const NEG = new Set(['4I', '4L', '5H', '5K']);
const taxOf = (a: number[]) => r2(a[1] + a[2] + a[3] + a[4]);
const nz = (v: any) => (typeof v === 'number' ? v : parseFloat(String(v ?? '').replace(/,/g, '')) || 0);

export interface MonthRow {
  ym: Ym;
  books: [number, number];
  g1: [number, number];
  /** 3B table 3.1 total (null = 3B not loaded for the month) */
  tb: [number, number] | null;
  /** GSTR-1 per 3B table, and 3B per table */
  g1t: Record<Table, [number, number]>;
  tbt: Record<Table, [number, number]> | null;
  adv: [number, number];
  rcm: [number, number];
  /** B2C: books entries ticked for the month's B2CS vs GSTR-1 B2CS */
  b2cBooks: [number, number];
  b2cG1: [number, number];
}
export interface StLine {
  code: string;
  label: string;
  action: string;
  /** adjustment from the start figure towards the end figure: [taxable, tax] */
  amt: [number, number];
  count: number;
  /** reason codes behind the line (to open the items) */
  reasons?: ReasonCode[];
  kind?: 'start' | 'sub' | 'end' | 'line' | 'unexplained';
  hint?: string;
}
export interface RevResult {
  fy: string;
  items: Item[];
  months: MonthRow[];
  books: { revenue: [number, number]; excluded: [number, number]; turnover: [number, number] };
  g1Turnover: [number, number];
  tbTotal: [number, number];
  /** Books → GSTR-1 */
  st1: StLine[];
  /** GSTR-1 → GSTR-3B */
  st2: StLine[];
  /** matched documents booked in one month and reported in another month of the year (not in the statement) */
  timing: { count: number; amt: [number, number] };
  checks: { level: 'info' | 'warn' | 'error'; msg: string }[];
  notes: string[];
  tol: number;
}

const add2 = (a: [number, number], b: [number, number], s = 1): [number, number] => [r2(a[0] + s * b[0]), r2(a[1] + s * b[1])];
const z2 = (): [number, number] => [0, 0];
const tblZero = () => Object.fromEntries(TABLES.map(t => [t, z2()])) as Record<Table, [number, number]>;

/** B2CS / Table 8 have no invoices: count the months that differ */
function monthsWithDiff(l: Item[], code: ReasonCode) {
  const m = new Map<string, number>();
  for (const i of l) {
    const k = code === 'B2C' ? i.b2cMonth : i.bYm || i.gYm;
    m.set(k, (m.get(k) ?? 0) + i.g[0] - i.b[0] + i.g[1] - i.b[1]);
  }
  return [...m.values()].filter(v => Math.abs(v) > 0.004).length;
}
/** a user decision for an item: by the item id, else by any of its document ids */
function userOf<T>(m: Record<string, T>, id: string): T | undefined {
  if (m[id] !== undefined) return m[id];
  for (const p of id.split(/[~,]/)) if (p && m[p] !== undefined) return m[p];
  return undefined;
}
const bookId = (r: any[]) => {
  const dt = String(r[6] ?? '');
  return `B|${dt === 'CN' || dt === 'DN' ? 'N' : 'I'}|${normNo(r[1])}`;
};

export function runRevenue(inp: RevInputs): RevResult {
  const { fy } = inp;
  const ny = fyShift(fy, 1);
  const mo = inp.match ?? DEFAULT_MATCH;
  const tol = mo.tol;
  const checks: RevResult['checks'] = [];
  const notes: string[] = [];
  const months = monthsOf(fy);

  // ---- books: this year's rows only
  const allRows = inp.books.rows.filter(r => r && [11, 12, 13, 14, 15].some(i => nz(r[i]) !== 0));
  const rows = allRows.filter(r => fyOfDate(toDate(r[0])) === fy);
  if (allRows.length > rows.length)
    checks.push({ level: 'info', msg: `${allRows.length - rows.length} books line(s) dated outside ${fy} are left out.` });
  const adv = rows.filter(r => String(r[5]) === 'ADVANCE');
  if (adv.length)
    checks.push({
      level: 'info',
      msg: `${adv.length} advance receipt line(s) in the register are not turnover and are left out (advances in GSTR-1 are compared with 3B only).`,
    });
  const excl = rows.filter(r => String(r[5]) === 'NOSUPPLY');
  const bRows = rows.filter(r => !['ADVANCE', 'NOSUPPLY'].includes(String(r[5])));
  // documents (lines of one invoice added up), same grouping as reconSales
  interface BDoc {
    id: string;
    rows: any[][];
    no: string;
    date: string;
    party: string;
    gstin: string;
    type: string;
    dt: string;
    tx: number;
    tax: number;
    src: string;
  }
  const bdocs = new Map<string, BDoc>();
  for (const r of [...bRows, ...excl]) {
    const id = bookId(r) + (String(r[5]) === 'NOSUPPLY' ? '|X' : '');
    let d = bdocs.get(id);
    if (!d)
      bdocs.set(
        id,
        (d = {
          id,
          rows: [],
          no: String(r[1] ?? ''),
          date: toDate(r[0]),
          party: String(r[2] ?? ''),
          gstin: String(r[3] ?? '')
            .trim()
            .toUpperCase(),
          type: String(r[5] ?? ''),
          dt: String(r[6] ?? 'INV') || 'INV',
          tx: 0,
          tax: 0,
          src: String(r[16] ?? ''),
        }),
      );
    d.rows.push(r);
    const sg = d.dt === 'CN' ? -1 : 1;
    d.tx = r2(d.tx + sg * Math.abs(nz(r[11])));
    d.tax = r2(d.tax + sg * (Math.abs(nz(r[12])) + Math.abs(nz(r[13])) + Math.abs(nz(r[14])) + Math.abs(nz(r[15]))));
  }

  // ---- GSTR-1 rows as RAW_GSTR1: A key, B-F amounts, G return FY, H period, I section, J doc type, K no, L date, M doc FY, N party ...
  const g1 = inp.g1.filter(r => r.key).map(r => [r.key, ...r.amt, ...r.info]);
  const DOCSEC = ['B2B', 'B2BA', 'B2CL', 'B2CLA', 'EXP', 'EXPA', 'CDNR', 'CDNRA', 'CDNUR', 'CDNURA'];

  const sales = reconSales(bRows, g1, fy, ny, new Map(), inp.manual ?? new Map(), notes, mo, inp.keepApart ?? new Set());
  const items: Item[] = [];
  const used = new Set<string>();
  const mk = (p: Partial<Item> & Pick<Item, 'id' | 'kind' | 'derived'>): Item => {
    const it: Item = {
      status: '',
      basis: '',
      party: '',
      gstin: '',
      type: '',
      dt: '',
      bNo: '',
      gNo: '',
      bDate: '',
      gDate: '',
      bYm: '',
      gYm: '',
      gSection: '',
      b: z2(),
      g: z2(),
      gAny: z2(),
      hasB: false,
      hasG: false,
      b2cDefault: '',
      b2cMonth: '',
      reason: p.derived,
      userReason: '',
      remark: '',
      note: '',
      bIds: [],
      rIds: [],
      src: '',
      ...p,
    };
    items.push(it);
    return it;
  };
  // reconSales: [0 status, 1 basis, 2 b gstin, 3 r gstin, 4 name, 5 doc type, 6 b no, 7 r no, 8 b date, 9 r date, 10 section,
  //   11 period MM/YYYY, 12 b taxable, 13 b tax, 14 r taxable, 15 r tax, 16 d tx, 17 d tax, 18 b type, 19 effect, 20 code, 21 remark, 22 id]
  for (const s of sales.rows) {
    const [bPart, rPart] = String(s[22]).split('~');
    const bIds = bPart ? bPart.split(',') : [],
      rIds = rPart ? rPart.split(',') : [];
    bIds.forEach(x => used.add(x));
    const sg = String(s[5]) === 'CN' ? -1 : 1;
    const hasB = s[12] !== null && s[12] !== undefined,
      hasG = s[14] !== null && s[14] !== undefined;
    const per = String(s[11] ?? ''); // MM/YYYY
    const gYm = per ? `${per.slice(3)}-${per.slice(0, 2)}` : '';
    const gFy = gYm ? fyOfPeriod(per.replace('/', '')) : '';
    const b: [number, number] = hasB ? [r2(sg * nz(s[12])), r2(sg * nz(s[13]))] : z2();
    const gAny: [number, number] = hasG ? [r2(sg * nz(s[14])), r2(sg * nz(s[15]))] : z2();
    const g: [number, number] = gFy === fy ? gAny : z2();
    const st = String(s[0]);
    let derived: ReasonCode;
    let note = '';
    const bd = bIds.map(x => bdocs.get(x)).filter(Boolean) as BDoc[];
    if (hasB && hasG) {
      if (gFy === ny) derived = 'NEXT_FY';
      else {
        const dT = Math.abs(b[0] - g[0]),
          dX = Math.abs(b[1] - g[1]);
        derived = dT > tol || dX > tol ? 'VALUE' : dT > 0.004 || dX > 0.004 ? 'ROUND' : 'MATCHED';
        if (st === 'Matched - GSTIN differs') note = 'Customer GSTIN in GSTR-1 differs from books';
      }
    } else if (hasB) derived = 'NOT_REPORTED';
    else {
      const dFy = fyOfDate(String(s[9] ?? ''));
      derived = gFy === fy && dFy && dFy < fy ? 'PREV_FY' : 'NOT_IN_BOOKS';
      if (gFy !== fy) note = `Reported in a ${gFy} return: not counted in this year`;
    }
    const bDate = String(s[8] ?? '');
    const it = mk({
      id: String(s[22]),
      kind: 'doc',
      derived,
      status: st,
      basis: String(s[1] ?? ''),
      party: String(s[4] ?? '') || bd[0]?.party || '',
      gstin: String(s[2] || s[3] || ''),
      type: String(s[18] || '') || String(s[10] ?? ''),
      dt: String(s[5] ?? ''),
      bNo: String(s[6] ?? ''),
      gNo: String(s[7] ?? ''),
      bDate,
      gDate: String(s[9] ?? ''),
      bYm: ymOfDate(bDate),
      gYm,
      gSection: String(s[10] ?? ''),
      b,
      g,
      gAny,
      hasB,
      hasG,
      note,
      bIds,
      rIds,
      src: bd.map(d => d.src).filter(Boolean)[0] || '',
    });
    // books invoices with no GSTIN from a GL dump that are not in GSTR-1 B2B: most likely in B2CS of their month
    if (derived === 'NOT_REPORTED' && inp.books.source === 'gl' && !it.gstin && it.type === 'B2B') {
      it.b2cDefault = it.bYm;
      it.note = 'No GSTIN in the books and not found in GSTR-1 B2B / B2CL: taken as a B2C sale of its month';
    }
  }
  // books documents not in reconSales' output: B2C small, nil / exempt / non-GST, entries excluded as not a supply
  for (const d of bdocs.values()) {
    if (used.has(d.id)) continue;
    const isExcl = d.type === 'NOSUPPLY';
    const kind: ItemKind = isExcl
      ? 'excl'
      : d.type === 'B2C'
        ? 'b2c'
        : ['EXEMPT', 'NIL', 'NON-GST'].includes(d.type)
          ? 'nil'
          : 'doc';
    const derived: ReasonCode = isExcl ? 'EXCLUDED' : kind === 'b2c' ? 'B2C' : kind === 'nil' ? 'NIL' : 'NOT_REPORTED';
    mk({
      id: `${d.id}~`,
      kind,
      derived,
      status: isExcl
        ? 'Not a supply'
        : kind === 'b2c'
          ? 'B2C (small)'
          : kind === 'nil'
            ? 'Nil / exempt / non-GST'
            : 'In books only',
      party: d.party,
      gstin: d.gstin,
      type: isExcl ? 'Not a supply' : d.type,
      dt: d.dt,
      bNo: d.no,
      bDate: d.date,
      bYm: ymOfDate(d.date),
      b: [d.tx, d.tax],
      hasB: true,
      b2cDefault: kind === 'b2c' ? ymOfDate(d.date) : '',
      bIds: [d.id],
      src: d.src,
    });
  }
  // GSTR-1 rows without invoices: B2CS (and amendments), Table 8, e-commerce u/s 9(5)
  for (const r of g1) {
    const key = String(r[0]),
      rfy = String(r[6]),
      sec = String(r[8]);
    if (rfy !== fy || DOCSEC.includes(sec)) continue;
    const t = tableOfKey(key);
    const per = String(r[7]);
    const ym = ymOfPeriod(per);
    const amt: [number, number] = [r2(nz(r[1])), taxOf([0, nz(r[2]), nz(r[3]), nz(r[4]), nz(r[5])])];
    let kind: ItemKind, derived: ReasonCode;
    if (sec === 'B2CS' || sec === 'B2CSA') {
      kind = 'b2c';
      derived = 'B2C';
    } else if (sec === 'NIL') {
      kind = 'nil';
      derived = 'NIL';
    } else if (key === '5C1') {
      kind = 'eco';
      derived = 'ECO';
    } else continue; // advances (4F) and others: compared with 3B only
    void t;
    mk({
      id: `G|${sec}|${per}|${r[10]}|${r[9]}|${r[14]}`,
      kind,
      derived,
      status:
        sec === 'NIL' ? `Table 8 ${r[9]}` : sec === 'B2CSA' ? 'B2CS amendment' : sec === 'B2CS' ? 'B2CS' : 'E-commerce 9(5)',
      party: String(r[13] ?? ''),
      type: sec,
      gNo: sec.startsWith('B2CS') ? `POS ${String(r[10]).split('/')[0]} · ${r[15]}%` : String(r[10] ?? ''),
      gYm: ym,
      gSection: sec,
      g: amt,
      gAny: amt,
      hasG: true,
      b2cDefault: kind === 'b2c' ? ym : '',
      rIds: [`G|${sec}|${per}|${r[10]}`],
    });
  }
  // merge identical ids (B2CS rows of one month / POS / rate in two files)
  {
    const seen = new Map<string, Item>();
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const o = seen.get(it.id);
      if (!o) {
        seen.set(it.id, it);
        continue;
      }
      o.g = add2(o.g, it.g);
      o.gAny = add2(o.gAny, it.gAny);
      o.b = add2(o.b, it.b);
      items.splice(i--, 1);
    }
  }

  // ---- user decisions
  for (const it of items) {
    const u = userOf(inp.user.reasons, it.id);
    const ub = it.hasB ? userOf(inp.user.b2c, it.bIds.join(',')) : undefined;
    it.b2cMonth = it.kind === 'b2c' && it.hasG ? it.b2cDefault : ub !== undefined ? ub : it.b2cDefault;
    it.userReason = u?.r && u.r !== 'MATCHED' ? u.r : '';
    it.remark = u?.note ?? '';
    it.reason = it.userReason || (it.hasB && it.b2cMonth ? 'B2C' : it.derived === 'B2C' && it.hasB ? 'NOT_REPORTED' : it.derived);
  }

  // ---- totals
  const sumBy = (f: (i: Item) => boolean, side: 'b' | 'g') => items.filter(f).reduce((a, i) => add2(a, i[side]), z2());
  const revenue = sumBy(() => true, 'b');
  const excluded = sumBy(i => i.reason === 'EXCLUDED', 'b');
  const turnover = add2(revenue, excluded, -1);
  const g1Turnover = sumBy(() => true, 'g');

  // ---- statement 1: books → GSTR-1 (adjustment = GSTR-1 side − books side, per reason)
  const st1: StLine[] = [{ code: 'rev', label: 'Revenue as per books', action: '', amt: revenue, count: 0, kind: 'start' }];
  if (Math.abs(excluded[0]) + Math.abs(excluded[1]) > 0.004)
    st1.push({
      code: 'EXCLUDED',
      label: 'Less: entries you marked as not a supply',
      action: REASON.EXCLUDED.action,
      amt: [r2(-excluded[0]), r2(-excluded[1])],
      count: items.filter(i => i.reason === 'EXCLUDED').length,
      reasons: ['EXCLUDED'],
      kind: 'line',
    });
  st1.push({ code: 'turn', label: 'Turnover as per books', action: '', amt: turnover, count: 0, kind: 'sub' });
  let sumLines = z2();
  for (const R of REASONS) {
    if (R.code === 'EXCLUDED' || R.code === 'MATCHED') continue;
    const l = items.filter(i => i.reason === R.code);
    const amt: [number, number] = l.reduce((a, i) => add2(a, add2(i.g, i.b, -1)), z2());
    if (Math.abs(amt[0]) < 0.005 && Math.abs(amt[1]) < 0.005) continue;
    sumLines = add2(sumLines, amt);
    const more = amt[0] < 0 || (amt[0] === 0 && amt[1] < 0);
    st1.push({
      code: R.code,
      label: `${more ? 'Less' : 'Add'}: ${R.label}`,
      action: R.action,
      amt,
      count:
        R.code === 'B2C' || R.code === 'NIL'
          ? monthsWithDiff(l, R.code)
          : l.filter(i => Math.abs(i.g[0] - i.b[0]) > 0.004 || Math.abs(i.g[1] - i.b[1]) > 0.004).length,
      reasons: [R.code],
      kind: 'line',
    });
  }
  // what the reasons don't cover: GSTR-1 side of excluded entries, rounding of r2
  const unx = add2(add2(g1Turnover, turnover, -1), sumLines, -1);
  st1.push({
    code: 'unx',
    label: 'Difference not explained',
    action: 'Should be nil. If not, look at entries marked as not a supply that are also in GSTR-1.',
    amt: unx,
    count: 0,
    kind: 'unexplained',
  });
  st1.push({ code: 'g1', label: 'Turnover as per GSTR-1', action: '', amt: g1Turnover, count: 0, kind: 'end' });

  // ---- timing within the year (matched, different months): not in the statement
  const timing = { count: 0, amt: z2() };
  for (const it of items)
    if (it.kind === 'doc' && it.hasB && it.hasG && it.bYm && it.gYm && it.bYm !== it.gYm && months.includes(it.gYm)) {
      timing.count++;
      timing.amt = add2(timing.amt, it.b);
    }

  // ---- months
  const tbOf = new Map<Ym, Record<Table, [number, number]>>();
  for (const t of inp.threeB) {
    const ym = ymOfPeriod(t.period);
    if (!months.includes(ym)) continue;
    const rec = tbOf.get(ym) ?? tblZero();
    for (const row of t.rows) {
      const k = row.key as Table;
      if (!TABLES.includes(k)) continue;
      rec[k] = add2(rec[k], [r2(row.amt[0]), taxOf(row.amt)]);
    }
    if (tbOf.has(ym)) checks.push({ level: 'warn', msg: `Two GSTR-3B files for ${ymLabel(ym)}: both are added.` });
    tbOf.set(ym, rec);
  }
  const M = new Map<Ym, MonthRow>(
    months.map(ym => [
      ym,
      {
        ym,
        books: z2(),
        g1: z2(),
        tb: null,
        g1t: tblZero(),
        tbt: tbOf.get(ym) ?? null,
        adv: z2(),
        rcm: z2(),
        b2cBooks: z2(),
        b2cG1: z2(),
      },
    ]),
  );
  for (const it of items) {
    if (it.reason !== 'EXCLUDED' && M.has(it.bYm)) M.get(it.bYm)!.books = add2(M.get(it.bYm)!.books, it.b);
    if (M.has(it.gYm)) M.get(it.gYm)!.g1 = add2(M.get(it.gYm)!.g1, it.g);
    if (it.reason === 'B2C') {
      if (it.hasB && M.has(it.b2cMonth)) M.get(it.b2cMonth)!.b2cBooks = add2(M.get(it.b2cMonth)!.b2cBooks, it.b);
      if (it.hasG && M.has(it.b2cMonth)) M.get(it.b2cMonth)!.b2cG1 = add2(M.get(it.b2cMonth)!.b2cG1, it.g);
    }
  }
  for (const r of g1) {
    if (String(r[6]) !== fy) continue;
    const key = String(r[0]);
    const t = tableOfKey(key);
    const m = M.get(ymOfPeriod(String(r[7])));
    if (!t || !m) continue;
    const sg = NEG.has(key) ? -1 : 1;
    const a: [number, number] = [r2(sg * nz(r[1])), r2(sg * taxOf([0, nz(r[2]), nz(r[3]), nz(r[4]), nz(r[5])]))];
    if (t === 'ADV') m.adv = add2(m.adv, a);
    else if (t === 'RCM') m.rcm = add2(m.rcm, a);
    else m.g1t[t] = add2(m.g1t[t], a);
  }
  for (const m of M.values()) if (m.tbt) m.tb = TABLES.reduce((a, t) => add2(a, m.tbt![t]), z2());
  const mrows = months.map(ym => M.get(ym)!);
  const tbTotal = mrows.reduce((a, m) => (m.tb ? add2(a, m.tb) : a), z2());

  // ---- statement 2: GSTR-1 → 3B, year level (months with 3B loaded); month-to-month shifts cancel out
  const withTb = mrows.filter(m => m.tbt);
  const noTb = mrows.filter(m => !m.tbt && (Math.abs(m.g1[0]) > 0.004 || TABLES.some(t => Math.abs(m.g1t[t][0]) > 0.004)));
  const st2: StLine[] = [{ code: 'g1', label: 'Turnover as per GSTR-1', action: '', amt: g1Turnover, count: 0, kind: 'start' }];
  const yAdv = mrows.reduce((a, m) => add2(a, m.adv), z2());
  const yRcm = mrows.reduce((a, m) => add2(a, m.rcm), z2());
  // GSTR-1 per table, for the year (all months) and for the months with 3B
  const g1Tables = mrows.reduce(
    (a, m) =>
      add2(
        a,
        TABLES.reduce((b, t) => add2(b, m.g1t[t]), z2()),
      ),
    z2(),
  );
  // the GSTR-1 turnover above (items) and the table view of the same returns can differ only by rows the tables leave out
  const nonTable = add2(add2(g1Turnover, yRcm, -1), g1Tables, -1);
  if (Math.abs(yRcm[0]) + Math.abs(yRcm[1]) > 0.004)
    st2.push({
      code: 'rcm',
      label: 'Less: supplies on reverse charge (tax paid by the recipient, not in 3B 3.1)',
      action: 'Not reported in 3B by the supplier',
      amt: [r2(-yRcm[0]), r2(-yRcm[1])],
      count: 0,
      kind: 'line',
    });
  if (Math.abs(yAdv[0]) + Math.abs(yAdv[1]) > 0.004)
    st2.push({
      code: 'adv',
      label: `${yAdv[0] >= 0 ? 'Add' : 'Less'}: advances received, net of advances adjusted (in 3B 3.1(a), not turnover)`,
      action: 'GSTR-1 Tables 11A / 11B',
      amt: yAdv,
      count: 0,
      kind: 'line',
    });
  if (Math.abs(nonTable[0]) + Math.abs(nonTable[1]) > 0.004)
    st2.push({
      code: 'nontable',
      label: `${nonTable[0] > 0 ? 'Less' : 'Add'}: GSTR-1 documents not counted in a 3B table`,
      action: 'Amendments whose original is not loaded, or documents of other sections',
      amt: [r2(-nonTable[0]), r2(-nonTable[1])],
      count: 0,
      kind: 'line',
    });
  if (noTb.length) {
    const a = noTb.reduce(
      (s, m) =>
        add2(
          s,
          add2(
            TABLES.reduce((b, t) => add2(b, m.g1t[t]), z2()),
            m.adv,
          ),
        ),
      z2(),
    );
    st2.push({
      code: 'notb',
      label: `Less: GSTR-1 of months with no 3B loaded (${noTb.map(m => ymLabel(m.ym)).join(', ')})`,
      action: 'Add those GSTR-3B PDFs to compare them',
      amt: [r2(-a[0]), r2(-a[1])],
      count: noTb.length,
      kind: 'line',
    });
  }
  // per 3B table: 3B − GSTR-1 for the months compared (3.1(a) includes advances)
  const diffT = Object.fromEntries(
    TABLES.map(t => [
      t,
      withTb.reduce((a, m) => add2(a, add2(m.tbt![t], t === '3.1a' ? add2(m.g1t[t], m.adv) : m.g1t[t], -1)), z2()),
    ]),
  ) as Record<Table, [number, number]>;
  for (const t of TABLES) {
    const d = diffT[t];
    if (Math.abs(d[0]) < 0.005 && Math.abs(d[1]) < 0.005) continue;
    const other = TABLES.filter(x => x !== t).find(
      x => Math.abs(diffT[x][0] + d[0]) <= Math.max(tol, Math.abs(d[0]) * 0.01) && Math.abs(d[0]) > tol,
    );
    const short = d[0] < 0 || (d[0] === 0 && d[1] < 0);
    const ms = withTb
      .map(m => [m.ym, add2(m.tbt![t], t === '3.1a' ? add2(m.g1t[t], m.adv) : m.g1t[t], -1)] as const)
      .filter(([, x]) => Math.abs(x[0]) > 0.004 || Math.abs(x[1]) > 0.004);
    st2.push({
      code: 'T' + t,
      label: `${short ? 'Less' : 'Add'}: ${TABLE_LABEL[t]}: ${short ? 'less in 3B than in GSTR-1' : 'more in 3B than in GSTR-1'}`,
      action: other
        ? `Offset by ${TABLE_LABEL[other as Table].split(' ')[0]}: likely reported under the other 3B table`
        : short
          ? 'Short reported in 3B: pay through a later 3B or DRC-03 (with interest)'
          : 'Excess in 3B: check for GSTR-1 short reported, or adjust in a later 3B',
      amt: d,
      count: ms.length,
      kind: 'line',
      hint: ms.map(([ym, x]) => `${ymLabel(ym)} ${x[0] >= 0 ? '+' : ''}${x[0].toFixed(2)}`).join(' · '),
    });
  }
  const st2Sum = st2.slice(1).reduce((a, l) => add2(a, l.amt), z2());
  const unx2 = add2(add2(tbTotal, g1Turnover, -1), st2Sum, -1);
  st2.push({
    code: 'unx',
    label: 'Difference not explained',
    action: 'Should be nil',
    amt: unx2,
    count: 0,
    kind: 'unexplained',
  });
  st2.push({
    code: 'tb',
    label: `Turnover as per GSTR-3B (Table 3.1 and 3.1.1(ii), ${withTb.length} month(s))`,
    action: '',
    amt: tbTotal,
    count: 0,
    kind: 'end',
  });

  // ---- set-up checks
  if (inp.books.source === 'none')
    checks.push({ level: 'error', msg: 'No books: add the sales register, or a day book / GL dump.' });
  if (!g1.some(r => String(r[6]) === fy)) checks.push({ level: 'error', msg: `No GSTR-1 of ${fy} loaded.` });
  const g1M = new Set(g1.filter(r => String(r[6]) === fy).map(r => ymOfPeriod(String(r[7]))));
  const gaps = months.filter(m => !g1M.has(m));
  if (g1M.size && gaps.length)
    checks.push({
      level: 'warn',
      msg: `No GSTR-1 data for ${gaps.map(ymLabel).join(', ')} (quarterly filers report in the quarter's last month; a nil return is fine).`,
    });
  if (!inp.threeB.length) checks.push({ level: 'warn', msg: 'No GSTR-3B PDFs loaded: the GSTR-1 → 3B statement is empty.' });
  if (!g1.some(r => String(r[6]) === ny))
    checks.push({
      level: 'info',
      msg: `No GSTR-1 of ${ny} loaded: invoices of this year reported in April–October of the next year show as not reported.`,
    });

  return {
    fy,
    items,
    months: mrows,
    books: { revenue, excluded, turnover },
    g1Turnover,
    tbTotal,
    st1,
    st2,
    timing,
    checks,
    notes,
    tol,
  };
}

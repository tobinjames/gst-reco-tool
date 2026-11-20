// GST ledger flow engine: classify books and portal entries into flow rows, build the books / portal / difference statements,
// and explain the differences from the entries themselves (no free-text reasons).
import { contraOf } from './books';
import { lineSign, ymOfDate } from './portal';
import { AccHead, AccType, BUCKETS, Block, BookFile, BookLine, COMPONENTS, Component, HEADS, Head, INTERNAL, PORTAL_LABEL, PortalKind, PortalLedger, UNMAPPED, Ym, bucketOf, r2 } from './types';
import type { ThreeB } from '../gstr3b';

export type HeadK = Head | 'ALL';
export const HEAD_KEYS: HeadK[] = [...HEADS, 'ALL'];
const TOL = 1;

export interface AccMapEntry { type: AccType; head: AccHead | '' }
export interface LineOverride { bucket?: string; head?: AccHead; period?: Ym }
export interface Inputs {
  fy: string; from: Ym; to: Ym;
  books: BookFile[]; portal: { file: string; ledger: PortalLedger }[];
  accMap: Record<string, AccMapEntry>;
  overrides: Record<string, LineOverride>;
  /** entered balances: books per account; portal per ledger kind and head */
  bal: { books: Record<string, { opening?: number; closing?: number }>; portal: Partial<Record<PortalKind, Partial<Record<Head, { opening?: number }>>>> };
  threeB?: ThreeB[];
}

/** a classified books line */
export interface BEntry {
  id: string; line: BookLine; accType: AccType; head: HeadK | ''; block: Block | null;
  effect: number;           // effect on the block's balance
  bucket: string; rule: string; auto: string; // bucket in use, why, and what the rules said (before any user change)
  month: Ym;                // comparison period (user may change it)
  where: 'before' | 'in' | 'after';
}
export interface PEntry {
  id: string; kind: PortalKind; file: string; row: number; date: string; ref: string; desc: string; dir: 'Dr' | 'Cr'; used: string;
  head: Head; comp: Component; amount: number; effect: number; block: Block | null; bucket: string; month: Ym; where: 'before' | 'in' | 'after';
}
export type Cells = Record<string, Record<HeadK, number>>; // bucket → head → amount (row value, i.e. effect × bucket sign)

export interface Statement {
  opening: Record<Block, Record<HeadK, number>>;
  rows: Cells;                                    // bucket → head → value
  closing: Record<Block, Record<HeadK, number>>;
}
export interface Item {
  id: string; block: Block; bucket: string; // bucket 'opening' for opening differences
  category: 'Opening balance' | 'Netted on portal' | 'Tax head' | 'Timing' | 'After the period' | 'Fee booked to tax' | 'In books, not on portal' | 'On portal, not in books' | 'Amount differs' | 'Not classified' | 'Rounding';
  explained: boolean; text: string; amt: Record<HeadK, number>;
  /** timing items: amount that moved between months (nets to nil for the period) */
  shift: Record<HeadK, number>; months: Ym[];
  books: string[]; portal: string[];
}
export const CARRIED = 'Carried from entries before the period';
export interface BalLine { side: 'Books' | 'Portal'; block: Block; head: HeadK; what: string; source: string; amount: number; raw?: number }
export interface Check { area: string; period: Ym | ''; head: string; a: number; b: number; aLabel: string; bLabel: string; diff: number; ok: boolean; note?: string }
export interface Result {
  books: BEntry[]; portal: PEntry[]; flowBooks: Statement; flowPortal: Statement; balances: BalLine[];
  items: Item[]; checks: Check[]; warns: { level: 'info' | 'warn' | 'error'; msg: string }[];
  monthly: { bucket: string; head: HeadK; month: Ym; b: number; p: number }[];
  months: Ym[];
}

const zeroK = (): Record<HeadK, number> => ({ IGST: 0, CGST: 0, SGST: 0, CESS: 0, ALL: 0 });
const blocks: Block[] = ['itc', 'liab', 'cash', 'fee'];
export const blockOfType = (t: AccType): Block | null => t === 'input' ? 'itc' : (t === 'output' || t === 'rcm_payable' || t === 'payable') ? 'liab' : t === 'cash_ledger' ? 'cash' : (t === 'interest' || t === 'late_fee' || t === 'penalty') ? 'fee' : null;
export const monthsOf = (fy: string): Ym[] => [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].map(m => `${m >= 4 ? fy.slice(0, 4) : +fy.slice(0, 4) + 1}-${String(m).padStart(2, '0')}`);
export const ymLabel = (ym: Ym) => ym ? `${['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][+ym.slice(5) - 1]}-${ym.slice(2, 4)}` : '';
export const lineId = (l: BookLine) => `B:${l.file}#${l.row}`;
export const bucketLabel = (id: string) => id === INTERNAL ? 'Transfers between GST accounts' : id === UNMAPPED ? 'Not classified' : id === 'opening' ? 'Opening balance' : bucketOf(id)?.label || id;
export const bucketsOfBlock = (b: Block) => BUCKETS.filter(x => x.block === b).map(x => x.id);
/** buckets the user may pick for a line of this account type */
export const allowedBuckets = (t: AccType) => { const b = blockOfType(t); return b ? [...bucketsOfBlock(b), INTERNAL] : []; };

// ---------- books: rule-based bucket for each GST line
const isDN = (s: string) => /debit note|purchase return|vendor credit|\bdn\b/i.test(s);
const isCN = (s: string) => /credit note|sales return|\bcn\b/i.test(s);
const salesLike = (s: string) => /sales|invoice|\bsale\b/i.test(s);
function ruleBucket(type: AccType, effect: number, contra: Set<string>, vt: string, narr: string): [string, string] {
  const has = (...k: string[]) => k.some(x => contra.has(x));
  const only = (...k: string[]) => contra.size > 0 && [...contra].every(x => k.includes(x));
  if (type === 'interest') return ['fee_interest', 'interest account'];
  if (type === 'late_fee') return ['fee_latefee', 'late fee account'];
  if (type === 'penalty') return ['fee_penalty', 'penalty account'];
  if (type === 'input') {
    if (only('input')) return [INTERNAL, 'transfer between input tax accounts'];
    if (has('rcm_payable')) return ['itc_rcm', 'voucher also credits RCM payable'];
    if (has('output', 'payable')) return ['itc_setoff', 'set off against output tax / GST payable'];
    if (has('refund_recv')) return ['itc_refund', 'transferred to GST refund receivable'];
    if (effect >= 0) {
      if (/reclaim|re-?avail|re-?credit|reversal (written )?back/i.test(narr)) return ['itc_reclaimed', 'narration says reclaimed'];
      if (has('party', 'expense', 'other', 'income', 'bank')) return ['itc_availed', 'input tax debited with a purchase / expense'];
      return [UNMAPPED, 'no other account in the voucher'];
    }
    if (isDN(vt) || (has('party') && !has('expense'))) return ['itc_availed', 'debit note / purchase return (reduces ITC availed)'];
    if (has('bank')) return ['itc_refund', 'input tax credited against bank (refund received)'];
    if (has('expense', 'other', 'income')) return ['itc_reversed', 'input tax credited to an expense / other account'];
    return [UNMAPPED, 'credit to input tax with no recognised other side'];
  }
  if (type === 'output' || type === 'payable' || type === 'rcm_payable') {
    if (only('output', 'payable', 'rcm_payable')) return [INTERNAL, 'transfer between output / payable accounts'];
    if (type === 'rcm_payable' && effect >= 0) return ['out_rcm', 'RCM payable credited'];
    if (effect < 0) {
      if (has('input')) return ['paid_itc', 'set off against input tax'];
      if (has('bank', 'cash_ledger')) return ['paid_cash', 'paid from bank / electronic cash ledger'];
      if (isCN(vt) || (has('party', 'income') && salesLike(vt))) return ['out_cn', 'credit note / sales return'];
      if (has('party', 'income')) return ['out_cn', 'output tax debited against a customer'];
      return [UNMAPPED, 'debit to output tax with no recognised other side'];
    }
    if (has('input') && has('party', 'expense')) return ['out_rcm', 'output tax credited together with input tax and a supplier / expense (reverse charge)'];
    if (has('party', 'income') || salesLike(vt) || isDN(vt)) return ['out_tax', 'output tax credited with a sale'];
    if (has('expense')) return ['out_rcm', 'tax credited against an expense (reverse charge without ITC)'];
    return [UNMAPPED, 'credit to output tax with no recognised other side'];
  }
  if (type === 'cash_ledger') {
    if (effect >= 0) return has('bank') ? ['cash_deposit', 'deposited from bank'] : [UNMAPPED, 'debit to cash ledger not from bank'];
    if (has('output', 'payable', 'rcm_payable', 'interest', 'late_fee', 'penalty')) return ['cash_used', 'used to pay tax / interest / fee'];
    if (has('bank', 'refund_recv')) return ['cash_refund', 'refunded to bank'];
    return [UNMAPPED, 'credit to cash ledger with no recognised other side'];
  }
  return [INTERNAL, ''];
}

export function run(inp: Inputs): Result {
  const warns: Result['warns'] = [];
  const months = monthsOf(inp.fy).filter(m => m >= inp.from && m <= inp.to);
  const where = (m: Ym): 'before' | 'in' | 'after' => (!m ? 'in' : m < inp.from ? 'before' : m > inp.to ? 'after' : 'in');
  const accOf = (name: string): AccMapEntry => inp.accMap[name] || { type: 'ignore', head: '' };

  // ---- books
  const B: BEntry[] = [];
  for (const f of inp.books) {
    const byV = new Map<string, BookLine[]>();
    for (const l of f.lines) { const k = l.vkey; (byV.get(k) || byV.set(k, []).get(k)!).push(l); }
    for (const l of f.lines) {
      const acc = accOf(l.account);
      if (acc.type === 'ignore' || acc.type === 'refund_recv') continue;
      const block = blockOfType(acc.type)!;
      const effect = r2(block === 'liab' ? l.cr - l.dr : l.dr - l.cr);
      if (!effect) continue;
      const contra = new Set<string>();
      if (f.format === 'ledger') {
        const m = inp.accMap[l.party];
        contra.add(m && m.type !== 'ignore' ? m.type : contraOf(l.party, ''));
      } else for (const o of byV.get(l.vkey)!) {
        if (o === l || o.account === l.account) continue;
        const m = accOf(o.account);
        contra.add(m.type !== 'ignore' ? m.type : contraOf(o.account, o.accGroup));
      }
      let [bucket, rule] = ruleBucket(acc.type, effect, contra, l.vchType, l.narration);
      const auto = bucket;
      const id = lineId(l), ov = inp.overrides[id] || {};
      let head: HeadK | '' = (ov.head || acc.head) as any;
      if (ov.bucket) { bucket = ov.bucket; rule = 'changed by user'; }
      if (!head && (block === 'itc' || block === 'liab')) { bucket = UNMAPPED; rule = 'tax head of the account not set'; }
      if (!head) head = 'ALL';
      const month = ov.period || ymOfDate(l.date);
      B.push({ id, line: l, accType: acc.type, head, block, effect, bucket, rule, auto, month, where: where(ymOfDate(l.date)) });
    }
  }

  // ---- portal: one entry per head and component with an amount; duplicate lines across overlapping downloads are dropped
  const P: PEntry[] = [];
  const seen = new Set<string>();
  for (const { file, ledger } of inp.portal) for (const l of ledger.lines) {
    const sig = [ledger.kind, l.date, l.ref, l.desc, l.dir, l.used, ...HEADS.flatMap(h => COMPONENTS.map(c => l.amt[h][c]))].join('|');
    if (seen.has(sig)) continue; seen.add(sig);
    const month = l.taxPeriod || ymOfDate(l.date);
    for (const h of HEADS) for (const c of COMPONENTS) {
      const a = l.amt[h][c]; if (!a) continue;
      const s = lineSign(ledger.kind, l.dir);
      let block: Block | null = null, bucket = INTERNAL, effect = r2(s * a);
      const d = l.desc.toLowerCase();
      if (ledger.kind === 'credit') {
        block = 'itc';
        bucket = l.dir === 'Cr' ? (/reclaim|re-?credit(?!.*refund)/.test(d) ? 'itc_reclaimed' : /refund/.test(d) ? 'itc_refund' : 'itc_availed')
          : (/\breversal\b|\breversed\b|ineligible|drc-?0?3|17\s*\(5\)|\blapse/.test(d) ? 'itc_reversed' : /refund/.test(d) ? 'itc_refund' : 'itc_setoff');
      } else if (ledger.kind === 'cash') {
        block = 'cash';
        bucket = l.dir === 'Cr' ? (/refund/.test(d) ? 'cash_refund' : 'cash_deposit') : (/refund/.test(d) ? 'cash_refund' : 'cash_used');
      } else {
        if (c === 'tax') {
          block = 'liab';
          bucket = l.dir === 'Dr' ? (/reverse charge|\brcm\b/.test(d) && !/other than/.test(d) ? 'out_rcm' : 'out_tax') : (/credit/i.test(l.used) ? 'paid_itc' : 'paid_cash');
        } else if (l.dir === 'Dr') {
          block = 'fee'; effect = r2(a);
          bucket = c === 'interest' ? 'fee_interest' : c === 'fee' ? 'fee_latefee' : 'fee_penalty';
        }
      }
      P.push({ id: `P:${file}#${l.row}:${h}:${c}`, kind: ledger.kind, file, row: l.row, date: l.date, ref: l.ref, desc: l.desc, dir: l.dir, used: l.used, head: h, comp: c, amount: a, effect, block, bucket, month, where: where(month) });
    }
  }

  // ---- statements
  const rowVal = (bucket: string, effect: number) => { const b = bucketOf(bucket); return b ? r2(effect * b.sign) : effect; };
  const emptyStmt = (): Statement => ({ opening: Object.fromEntries(blocks.map(b => [b, zeroK()])) as any, rows: {}, closing: Object.fromEntries(blocks.map(b => [b, zeroK()])) as any });
  const rowKey = (block: Block, bucket: string) => (bucket === INTERNAL || bucket === UNMAPPED ? `${block}:${bucket}` : bucket);
  const addRow = (st: Statement, key: string, head: HeadK, v: number) => { (st.rows[key] ||= zeroK())[head] = r2(st.rows[key][head] + v); };

  const fb = emptyStmt();
  const balances: BalLine[] = [];
  const donePortal = new Set<string>();
  // books opening: entered (or printed in a single-ledger file) per account, plus lines before the period
  const accFile = new Map<string, number>();
  for (const f of inp.books) if (f.format === 'ledger' && f.ledger && f.opening !== undefined) accFile.set(f.ledger, f.opening);
  for (const [name, m] of Object.entries(inp.accMap)) {
    const block = blockOfType(m.type); if (!block || block === 'fee') continue;
    const raw = inp.bal.books[name]?.opening ?? accFile.get(name);
    if (raw === undefined) continue;
    const head = (m.head || 'ALL') as HeadK;
    // ledgers show Dr balances as +; liability accounts carry Cr balances
    const v = block === 'liab' ? -raw : raw;
    fb.opening[block][head] = r2(fb.opening[block][head] + v);
    balances.push({ side: 'Books', block, head, what: name, source: inp.bal.books[name]?.opening !== undefined ? 'Entered (balance on the first day)' : 'Opening Balance row of the ledger file', amount: r2(v), raw });
  }
  const carried = (side: 'Books' | 'Portal', block: Block, head: HeadK, v: number) => {
    let b = balances.find(x => x.side === side && x.block === block && x.head === head && x.source === CARRIED);
    if (!b) balances.push(b = { side, block, head, what: side === 'Books' ? 'Entries dated before the period' : 'Entries for return periods before the period', source: CARRIED, amount: 0 });
    b.amount = r2(b.amount + v);
  };
  for (const e of B) {
    if (!e.block) continue;
    const head = e.head as HeadK;
    if (e.where === 'before') { if (e.block !== 'fee') { fb.opening[e.block][head] = r2(fb.opening[e.block][head] + e.effect); carried('Books', e.block, head, e.effect); } continue; }
    if (e.where === 'after') continue;
    addRow(fb, rowKey(e.block, e.bucket), head, rowVal(e.bucket, e.effect));
  }
  const fp = emptyStmt();
  for (const { ledger } of [...inp.portal].sort((a, b) => ymOfDate(a.ledger.from).localeCompare(ymOfDate(b.ledger.from)) || a.ledger.from.localeCompare(b.ledger.from))) {
    const block: Block = ledger.kind === 'credit' ? 'itc' : ledger.kind === 'cash' ? 'cash' : 'liab';
    for (const h of HEADS) {
      if (donePortal.has(`${ledger.kind}|${h}`)) continue;
      const ent = inp.bal.portal[ledger.kind]?.[h]?.opening;
      const v = ent ?? (ledger.from && ymOfDate(ledger.from) <= inp.from ? ledger.opening?.[h] : undefined);
      if (v === undefined) continue;
      donePortal.add(`${ledger.kind}|${h}`);
      fp.opening[block][h] = r2(fp.opening[block][h] + v);
      if (v) balances.push({ side: 'Portal', block, head: h, what: PORTAL_LABEL[ledger.kind], source: ent !== undefined ? 'Entered (balance on the first day)' : `Opening Balance row of the downloaded ledger (from ${ledger.from})`, amount: v });
    }
  }
  for (const e of P) {
    if (!e.block) continue;
    if (e.where === 'before') { if (e.block !== 'fee') { fp.opening[e.block][e.head] = r2(fp.opening[e.block][e.head] + e.effect); carried('Portal', e.block, e.head, e.effect); } continue; }
    if (e.where === 'after') continue;
    addRow(fp, rowKey(e.block, e.bucket), e.head, rowVal(e.bucket, e.effect));
  }
  for (const st of [fb, fp]) for (const block of blocks) for (const h of HEAD_KEYS) {
    let c = block === 'fee' ? 0 : st.opening[block][h];
    for (const [k, v] of Object.entries(st.rows)) {
      const b = bucketOf(k), blk = b ? b.block : k.split(':')[0];
      if (blk !== block) continue;
      c += b ? v[h] * b.sign : v[h];
    }
    st.closing[block][h] = r2(c);
  }

  // ---- month-wise cells for the reasons
  // rows the books keep in an account not split by tax head (head ALL only): compare the portal in total for those rows
  const rowHeads = new Map<string, Set<HeadK>>();
  for (const e of B) if (e.block && e.where === 'in') { const k = rowKey(e.block, e.bucket); (rowHeads.get(k) || rowHeads.set(k, new Set()).get(k)!).add(e.head as HeadK); }
  const totalOnly = new Set([...rowHeads].filter(([, hs]) => hs.size === 1 && hs.has('ALL')).map(([k]) => k));
  type Cell = { b: number; p: number; bl: string[]; pl: string[] };
  const cells = new Map<string, Cell>(); // `${rowKey}|${head}|${month}`
  const cell = (k: string, h: HeadK, m: Ym) => { const kk = `${k}|${h}|${m}`; let c = cells.get(kk); if (!c) cells.set(kk, c = { b: 0, p: 0, bl: [], pl: [] }); return c; };
  for (const e of B) if (e.block && e.where === 'in') { const c = cell(rowKey(e.block, e.bucket), e.head as HeadK, e.month); c.b = r2(c.b + rowVal(e.bucket, e.effect)); c.bl.push(e.id); }
  const items: Item[] = [];
  const blockOfKey = (k: string): Block => (bucketOf(k)?.block || k.split(':')[0]) as Block;
  const newItem = (block: Block, bucket: string, category: Item['category'], explained: boolean, text: string, months: Ym[], b: string[], p: string[]): Item => {
    const it: Item = { id: '', block, bucket, category, explained, text, amt: zeroK(), shift: zeroK(), months, books: [...new Set(b)], portal: [...new Set(p)] };
    items.push(it); return it;
  };
  const collapsed = new Map<string, Item>();
  for (const e of P) if (e.block && e.where === 'in') {
    const k = rowKey(e.block, e.bucket), v = rowVal(e.bucket, e.effect);
    let h: HeadK = e.head;
    if (totalOnly.has(k)) {
      // one item per row moves the portal's head-wise amounts into the Total-only column the books use
      let it = collapsed.get(k);
      if (!it) collapsed.set(k, it = newItem(blockOfKey(k), k, 'Tax head', true, `${bucketLabel(k)}: the books account is not split by tax head, so this row is compared in total`, [], [], []));
      it.amt[e.head] = r2(it.amt[e.head] - v); it.amt.ALL = r2(it.amt.ALL + v); it.portal.push(e.id);
      h = 'ALL';
    }
    const c = cell(k, h, e.month); c.p = r2(c.p + v); c.pl.push(e.id);
  }
  const monthly: Result['monthly'] = [];
  for (const [k, c] of cells) { const [bucket, head, month] = k.split('|'); monthly.push({ bucket, head: head as HeadK, month, b: c.b, p: c.p }); }
  // residual differences, worked down step by step
  const R = new Map<string, number>(); for (const [k, c] of cells) R.set(k, r2(c.b - c.p));
  const res = (k: string, h: HeadK, m: Ym) => R.get(`${k}|${h}|${m}`) || 0;
  const take = (k: string, h: HeadK, m: Ym, x: number) => R.set(`${k}|${h}|${m}`, r2(res(k, h, m) - x));
  const refs = (k: string, h: HeadK, m: Ym) => { const c = cells.get(`${k}|${h}|${m}`); return c ? { b: c.bl, p: c.pl } : { b: [] as string[], p: [] as string[] }; };
  const sgn = (x: number) => (x > 0 ? 1 : -1);
  const big = (x: number) => Math.abs(x) > TOL;
  const keysOf = () => [...new Set([...cells.keys()].map(k => k.split('|')[0]))];

  // 1. opening balances
  for (const block of ['itc', 'liab', 'cash'] as Block[]) {
    const d = zeroK(); let any = false;
    for (const h of HEAD_KEYS) { d[h] = r2(fb.opening[block][h] - fp.opening[block][h]); if (Math.abs(d[h]) > 0.004) any = true; }
    if (any) { const it = newItem(block, 'opening', 'Opening balance', true, 'Opening balance as per books differs from the portal', [inp.from], [], []); it.amt = d; }
  }
  // 2. the portal nets some rows into another (GSTR-3B 4C is ITC net of reversals; output tax is net of credit notes)
  const NET: [string, string, string][] = [
    ['itc_reversed', 'itc_availed', 'ITC reversed in books; the credit ledger shows ITC net of reversals (GSTR-3B 4C)'],
    ['itc_rcm', 'itc_availed', 'Reverse charge ITC is part of "ITC accrued" on the credit ledger'],
    ['itc_reclaimed', 'itc_availed', 'Reclaimed ITC is part of "ITC accrued" on the credit ledger'],
    ['out_cn', 'out_tax', 'Credit notes booked separately in books; on the portal output tax is net of credit notes'],
  ];
  for (const [a, bk, text] of NET) for (const h of HEAD_KEYS) for (const m of months) {
    const ra = res(a, h, m), rb = res(bk, h, m), pa = cells.get(`${a}|${h}|${m}`)?.p || 0;
    const sa = bucketOf(a)!.sign, sb = bucketOf(bk)!.sign;
    if (!big(ra) || !big(rb) || big(pa) || sgn(ra * sa) === sgn(rb * sb)) continue;
    const x = Math.min(Math.abs(ra), Math.abs(rb));
    const ref1 = refs(a, h, m), ref2 = refs(bk, h, m);
    const it1 = newItem(bucketOf(a)!.block, a, 'Netted on portal', true, text, [m], ref1.b, ref2.p); it1.amt[h] = r2(x * sgn(ra));
    const it2 = newItem(bucketOf(bk)!.block, bk, 'Netted on portal', true, text, [m], ref1.b, ref2.p); it2.amt[h] = r2(x * sgn(rb));
    take(a, h, m, it1.amt[h]); take(bk, h, m, it2.amt[h]);
  }
  // 3. tax head differs (IGST vs CGST + SGST; CGST vs SGST): same row and month, amounts that offset
  for (const k of keysOf()) for (const m of months) {
    const pairs: [HeadK[], HeadK[]][] = [[['IGST'], ['CGST', 'SGST']], [['CGST'], ['SGST']], [['IGST', 'CGST', 'SGST', 'CESS'], ['ALL']]];
    for (const [A, Bh] of pairs) {
      const ra = A.reduce((s, h) => s + res(k, h, m), 0), rb = Bh.reduce((s, h) => s + res(k, h, m), 0);
      if (!big(ra) || !big(rb) || sgn(ra) === sgn(rb)) continue;
      const x = Math.min(Math.abs(ra), Math.abs(rb));
      const used = /setoff|paid_itc/.test(k);
      const text = used ? 'ITC used against a different tax head on the portal than in books (cross-utilisation)'
        : Bh[0] === 'ALL' ? 'Booked in an account not split by tax head in books' : 'Booked under a different tax head in books than on the portal';
      const r = [...A, ...Bh].map(h => refs(k, h, m));
      const it = newItem(blockOfKey(k), k, 'Tax head', true, text, [m], r.flatMap(x => x.b), r.flatMap(x => x.p));
      // spread x over the heads in proportion to their residuals
      for (const [grp, tot] of [[A, ra], [Bh, rb]] as [HeadK[], number][]) {
        const vals = grp.map(h => res(k, h, m));
        grp.forEach((h, i) => { const share = r2(x * sgn(tot) * (vals[i] / tot)); if (share) { it.amt[h] = r2(it.amt[h] + share); take(k, h, m, share); } });
      }
    }
  }
  // 4. timing: same row and head, opposite differences in nearby months. Nets to nil for the period; the amount moved is kept as `shift`
  // 4a. entry by entry: a books entry and a portal entry of the same amount in different months (challans, single invoices)
  {
    const groups = new Map<string, { b: { id: string; m: Ym; v: number }[]; p: { ids: string[]; m: Ym; v: number }[] }>();
    const g = (k: string, h: HeadK) => { const kk = `${k}|${h}`; let x = groups.get(kk); if (!x) groups.set(kk, x = { b: [], p: [] }); return x; };
    for (const e of B) if (e.block && e.where === 'in') g(rowKey(e.block, e.bucket), e.head as HeadK).b.push({ id: e.id, m: e.month, v: rowVal(e.bucket, e.effect) });
    const pl = new Map<string, { k: string; h: HeadK; ids: string[]; m: Ym; v: number }>();
    for (const e of P) if (e.block && e.where === 'in') {
      const k = rowKey(e.block, e.bucket), h: HeadK = totalOnly.has(k) ? 'ALL' : e.head, lk = `${e.file}#${e.row}|${k}|${h}`;
      const x = pl.get(lk) || pl.set(lk, { k, h, ids: [], m: e.month, v: 0 }).get(lk)!; x.ids.push(e.id); x.v = r2(x.v + rowVal(e.bucket, e.effect));
    }
    for (const x of pl.values()) g(x.k, x.h).p.push({ ids: x.ids, m: x.m, v: x.v });
    // only rows where both sides are single payments / journals; invoices in books against a return total on the portal are compared month-wise
    const ENTRY_ROWS = /^(cash_|paid_|itc_setoff|itc_refund)/;
    for (const [kk, { b, p }] of groups) {
      const [k, h] = kk.split('|') as [string, HeadK];
      if (!ENTRY_ROWS.test(k)) continue;
      const usedB = new Set<number>(), usedP = new Set<number>();
      // same month first (needs no explanation), then the nearest other month
      for (const sameMonth of [true, false]) b.forEach((be, bi) => {
        if (usedB.has(bi) || !big(be.v)) return;
        let best = -1, bestGap = 99;
        p.forEach((pe, pi) => {
          if (usedP.has(pi) || Math.abs(pe.v - be.v) > TOL) return;
          const gap = Math.abs(months.indexOf(be.m) - months.indexOf(pe.m));
          if (sameMonth ? gap !== 0 : gap === 0 || gap > 3) return;
          if (gap < bestGap) { best = pi; bestGap = gap; }
        });
        if (best < 0) return;
        const pe = p[best];
        if (!sameMonth) {
          // the same entry in different months: move it (same-month pairs were taken first, so equal monthly amounts pair within their month)
          const x = be.v;
          const lag = months.indexOf(be.m) - months.indexOf(pe.m);
          const it = newItem(blockOfKey(k), k, 'Timing', true, `${bucketLabel(k)}: booked ${Math.abs(lag)} month${Math.abs(lag) > 1 ? 's' : ''} ${lag < 0 ? 'earlier' : 'later'} in books than the return period on the portal`, [be.m, pe.m], [be.id], pe.ids);
          it.shift[h] = r2(Math.abs(x)); take(k, h, be.m, x); take(k, h, pe.m, -x);
        }
        usedB.add(bi); usedP.add(best);
      });
    }
  }
  // 4b. month by month: exact amounts first (next month, then two months on), then what is left, partly
  for (const [gap, exact] of [[1, true], [2, true], [1, false], [2, false]] as [number, boolean][])
  for (const k of keysOf()) for (const h of HEAD_KEYS) for (let i = 0; i + gap < months.length; i++) {
    const j = i + gap;
    const ri = res(k, h, months[i]), rj = res(k, h, months[j]);
    if (!big(ri) || !big(rj) || sgn(ri) === sgn(rj) || (exact && Math.abs(ri + rj) > TOL)) continue;
    const x = Math.min(Math.abs(ri), Math.abs(rj));
    const lag = ri > 0 ? i - j : j - i; // books month − portal month
    const r1 = refs(k, h, months[i]), r2_ = refs(k, h, months[j]);
    const it = newItem(blockOfKey(k), k, 'Timing', true, `${bucketLabel(k)}: booked ${Math.abs(lag)} month${Math.abs(lag) > 1 ? 's' : ''} ${lag < 0 ? 'earlier' : 'later'} in books than the return period on the portal`,
      [months[i], months[j]], [...r1.b, ...r2_.b], [...r1.p, ...r2_.p]);
    it.shift[h] = r2(x);
    take(k, h, months[i], r2(x * sgn(ri))); take(k, h, months[j], r2(x * sgn(rj)));
  }
  // 5. entries that fall outside the period on the other side (books in March, portal return in April of the next year, etc.)
  const outside = (side: 'b' | 'p', k: string, h: HeadK) => {
    const list: { id: string; v: number; month: Ym }[] = side === 'p'
      ? P.filter(e => e.block && e.where === 'after' && rowKey(e.block, e.bucket) === k && (totalOnly.has(k) ? h === 'ALL' : e.head === h)).map(e => ({ id: e.id, v: rowVal(e.bucket, e.effect), month: e.month }))
      : B.filter(e => e.block && e.where === 'after' && rowKey(e.block, e.bucket) === k && e.head === h).map(e => ({ id: e.id, v: rowVal(e.bucket, e.effect), month: ymOfDate(e.line.date) }));
    return { sum: r2(list.reduce((s, e) => s + e.v, 0)), ids: list.map(e => e.id), months: [...new Set(list.map(e => e.month))].sort() };
  };
  for (const k of keysOf()) for (const h of HEAD_KEYS) for (const side of ['p', 'b'] as const) {
    const tot = r2(months.reduce((s, m) => s + res(k, h, m), 0));
    if (!big(tot)) continue;
    const o = outside(side, k, h);
    // books higher in the period (tot > 0) is explained by portal entries after it; portal higher by books entries after it
    if (!big(o.sum) || (side === 'p' ? sgn(tot) !== sgn(o.sum) : sgn(tot) === sgn(o.sum))) continue;
    let x = r2(Math.min(Math.abs(tot), Math.abs(o.sum)) * sgn(tot));
    const it = newItem(blockOfKey(k), k, 'After the period', true, side === 'p'
      ? `${bucketLabel(k)}: in books in the period, on the portal in a later return period`
      : `${bucketLabel(k)}: on the portal for a return period in the period, in books after it`, [], side === 'b' ? o.ids : [], side === 'p' ? o.ids : []);
    for (let i = months.length - 1; i >= 0 && Math.abs(x) > 0.004; i--) {
      const r = res(k, h, months[i]); if (Math.abs(r) <= 0.004 || sgn(r) !== sgn(x)) continue;
      const t = Math.abs(r) < Math.abs(x) ? r : x; take(k, h, months[i], t); it.amt[h] = r2(it.amt[h] + t); x = r2(x - t);
      const rf = refs(k, h, months[i]); it.books.push(...rf.b); it.portal.push(...rf.p); it.months.push(months[i]);
    }
    it.months.push(...o.months);
  }
  // 6. interest / late fee paid in cash but debited to the tax ledger in books
  for (const h of HEAD_KEYS) for (const m of months) {
    const r = res('paid_cash', h, m); if (!(r > TOL)) continue;
    const feePaid = P.filter(e => e.kind === 'liability' && e.dir === 'Cr' && e.comp !== 'tax' && /cash/i.test(e.used) && e.head === h && e.month === m);
    const fpaid = r2(feePaid.reduce((s, e) => s + e.amount, 0));
    if (!big(fpaid) || B.some(e => e.block === 'fee' && e.month === m)) continue;
    const x = Math.min(r, fpaid);
    const it = newItem('liab', 'paid_cash', 'Fee booked to tax', true, 'Interest / late fee paid in cash on the portal, debited to the GST tax ledger in books', [m], refs('paid_cash', h, m).b, feePaid.map(e => e.id));
    it.amt[h] = r2(x); take('paid_cash', h, m, x);
    // the same fee then shows on the portal's fee row with no books expense: that is the other side of this item
    const fk = feePaid.map(e => rowKey('fee', e.comp === 'interest' ? 'fee_interest' : e.comp === 'fee' ? 'fee_latefee' : 'fee_penalty'));
    for (const k of new Set(fk)) {
      const rr = res(k, h, m); if (!(rr < -TOL)) continue;
      const y = Math.max(rr, -x);
      const it2 = newItem('fee', k, 'Fee booked to tax', true, 'Interest / late fee paid in cash on the portal, debited to the GST tax ledger in books', [m], [], refs(k, h, m).p);
      it2.amt[h] = r2(y); take(k, h, m, y);
    }
  }
  // 7. what is left is unexplained, listed by row and month with the entries behind it
  for (const k of keysOf()) for (const m of months) for (const h of HEAD_KEYS) {
    const r = res(k, h, m); if (Math.abs(r) <= 0.004) continue;
    const c = cells.get(`${k}|${h}|${m}`)!;
    const cat: Item['category'] = k.endsWith(UNMAPPED) ? 'Not classified' : !big(r) ? 'Rounding' : !big(c.p) ? 'In books, not on portal' : !big(c.b) ? 'On portal, not in books' : 'Amount differs';
    const text = cat === 'Not classified' ? 'Books entries not classified into a flow row: pick a row for each'
      : cat === 'Rounding' ? `Rounding (under ₹${TOL})` : `${bucketLabel(k)}: ${cat.toLowerCase()}`;
    const it = newItem(blockOfKey(k), k, cat, cat === 'Rounding', text, [m], c.bl, c.pl); it.amt[h] = r; take(k, h, m, r);
  }
  // one item per row, reason and month (timing items: per row and lag, across months)
  const merged = new Map<string, Item>();
  for (const it of items) {
    const mk = [it.category, it.bucket, it.text, it.category === 'Timing' || it.category === 'Tax head' && it.months.length === 0 ? '' : it.months.join(',')].join('|');
    const t = merged.get(mk);
    if (!t) { merged.set(mk, it); continue; }
    for (const h of HEAD_KEYS) { t.amt[h] = r2(t.amt[h] + it.amt[h]); t.shift[h] = r2(t.shift[h] + it.shift[h]); }
    t.months.push(...it.months); t.books.push(...it.books); t.portal.push(...it.portal);
  }
  const kept = [...merged.values()].filter(it => HEAD_KEYS.some(h => Math.abs(it.amt[h]) > 0.004 || Math.abs(it.shift[h]) > 0.004));
  const blkOrder: Block[] = ['itc', 'liab', 'cash', 'fee'];
  const bOrder = (k: string) => (k === 'opening' ? -1 : BUCKETS.findIndex(x => x.id === k) >= 0 ? BUCKETS.findIndex(x => x.id === k) : 99);
  kept.sort((a, b) => blkOrder.indexOf(a.block) - blkOrder.indexOf(b.block) || bOrder(a.bucket) - bOrder(b.bucket) || +b.explained - +a.explained || (a.months[0] || '').localeCompare(b.months[0] || ''));
  kept.forEach((it, i) => { it.id = `R${String(i + 1).padStart(3, '0')}`; it.months = [...new Set(it.months)].sort(); it.books = [...new Set(it.books)]; it.portal = [...new Set(it.portal)]; });

  return { books: B, portal: P, flowBooks: fb, flowPortal: fp, balances, items: kept, checks: checks(inp, P, months, warns), warns, monthly, months };
}
const fmt = (x: number) => x.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

// ---------- cross-checks between the portal ledgers, and with GSTR-3B
function checks(inp: Inputs, P: PEntry[], months: Ym[], warns: Result['warns']): Check[] {
  const out: Check[] = [];
  const have = (k: PortalKind) => inp.portal.some(p => p.ledger.kind === k);
  const sum = (f: (e: PEntry) => boolean) => r2(P.filter(f).reduce((s, e) => s + e.amount, 0));
  const push = (c: Omit<Check, 'diff' | 'ok'>) => { const diff = r2(c.a - c.b); out.push({ ...c, diff, ok: Math.abs(diff) <= TOL }); };
  for (const m of months) {
    if (have('cash') && have('liability')) for (const h of HEADS) {
      const a = sum(e => e.kind === 'cash' && e.bucket === 'cash_used' && e.head === h && e.month === m);
      const b = sum(e => e.kind === 'liability' && e.dir === 'Cr' && /cash/i.test(e.used) && e.head === h && e.month === m);
      if (a || b) push({ area: 'Cash ledger used vs liability paid in cash', period: m, head: h, a, b, aLabel: 'Cash ledger debits', bLabel: 'Liability register: paid in cash' });
    }
    if (have('credit') && have('liability')) {
      const a = sum(e => e.kind === 'credit' && e.bucket === 'itc_setoff' && e.month === m);
      const b = sum(e => e.kind === 'liability' && e.dir === 'Cr' && /credit/i.test(e.used) && e.month === m);
      if (a || b) {
        const per = HEADS.map(h => `${h} ${fmt(sum(e => e.kind === 'credit' && e.bucket === 'itc_setoff' && e.month === m && e.head === h))} used / ${fmt(sum(e => e.kind === 'liability' && e.dir === 'Cr' && /credit/i.test(e.used) && e.month === m && e.head === h))} paid`).join(' · ');
        push({ area: 'Credit ledger used vs liability paid by ITC', period: m, head: 'Total', a, b, aLabel: 'Credit ledger debits', bLabel: 'Liability register: paid by ITC', note: per });
      }
    }
    const tb = inp.threeB?.find(t => t.period === `${m.slice(5)}${m.slice(0, 4)}`);
    if (tb) {
      const val = (key: string, h: Head) => { const r = tb.rows.find(x => x.key === key); if (!r) return 0; return h === 'IGST' ? r.amt[3] : h === 'CGST' ? r.amt[1] : h === 'SGST' ? r.amt[2] : r.amt[4]; };
      const pay = (key: string) => tb.pay.find(p => p.key === key)?.amount || 0;
      const L: Record<Head, string> = { IGST: 'A', CGST: 'B', SGST: 'C', CESS: 'D' };
      for (const h of HEADS) {
        if (have('credit')) {
          const net = r2(['4A1', '4A2', '4A3', '4A4', '4A5', '4D1'].reduce((s, k) => s + val(k, h), 0) - ['4B1', '4B2'].reduce((s, k) => s + val(k, h), 0));
          const a = sum(e => e.kind === 'credit' && /itc_(availed|rcm|reclaimed)/.test(e.bucket) && e.dir === 'Cr' && e.head === h && e.month === m);
          if (a || net) push({ area: 'Credit ledger ITC vs GSTR-3B net ITC (4C)', period: m, head: h, a, b: net, aLabel: 'Credit ledger: ITC accrued', bLabel: 'GSTR-3B 4A − 4B (+4D1)' });
        }
        if (have('liability')) {
          const a = sum(e => e.kind === 'liability' && e.dir === 'Dr' && e.comp === 'tax' && e.head === h && e.month === m);
          const b = pay(`9${L[h]}-GROSS`);
          if (a || b) push({ area: 'Liability register vs GSTR-3B tax payable (6.1)', period: m, head: h, a, b, aLabel: 'Liability register: tax', bLabel: 'GSTR-3B 6.1 tax payable' });
          const ac = sum(e => e.kind === 'liability' && e.dir === 'Cr' && e.comp === 'tax' && /cash/i.test(e.used) && e.head === h && e.month === m), bc = pay(`9${L[h]}-CASH`);
          if (ac || bc) push({ area: 'Paid in cash vs GSTR-3B 6.1', period: m, head: h, a: ac, b: bc, aLabel: 'Liability register: paid in cash', bLabel: 'GSTR-3B 6.1 paid in cash' });
        }
      }
    }
  }
  // returns covered by the portal ledgers
  for (const k of ['credit', 'liability'] as PortalKind[]) {
    if (!have(k)) continue;
    const got = new Set(P.filter(e => e.kind === k && (k === 'credit' ? e.bucket === 'itc_availed' : e.dir === 'Dr')).map(e => e.month));
    const ledgerMonths = new Set(inp.portal.filter(p => p.ledger.kind === k).flatMap(p => p.ledger.lines.map(l => l.taxPeriod || ymOfDate(l.date))));
    const miss = months.filter(m => !got.has(m) && !ledgerMonths.has(m));
    if (miss.length) warns.push({ level: 'warn', msg: `${k === 'credit' ? 'Electronic Credit Ledger' : 'Electronic Liability Register'}: no entries for ${miss.map(ymLabel).join(', ')}. Download it for the whole year (up to the filing of the last return) or check that those returns were filed.` });
  }
  for (const { file, ledger } of inp.portal) {
    if (ledger.kind !== 'liability' && ledger.from && ymOfDate(ledger.from) > inp.from && inp.bal.portal[ledger.kind] === undefined)
      warns.push({ level: 'warn', msg: `${file}: starts on ${ledger.from}, after the start of the period. Enter the balance at the start of the period under Balances, or download the ledger from ${`01-${inp.from.slice(5)}-${inp.from.slice(0, 4)}`}.` });
    ledger.problems.forEach(p => warns.push({ level: 'warn', msg: `${file}: ${p}` }));
    if (ledger.kind === 'liability' && ledger.closing && HEADS.some(h => Math.abs(ledger.closing![h]) > TOL))
      warns.push({ level: 'warn', msg: `${file}: closing balance shows unpaid liability (${HEADS.filter(h => ledger.closing![h]).map(h => `${h} ${fmt(ledger.closing![h])}`).join(', ')})` });
  }
  return out;
}

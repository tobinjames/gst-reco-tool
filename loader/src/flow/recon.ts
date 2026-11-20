// Reconciliation statement: closing as per books → the events that explain the difference → closing as per portal.
// Reconciling items (engine.ts) are per flow row; here they are grouped into real events (one reversal netted on the portal is
// one event, not two rows), each with its effect on the closing balance, a plain description, the vouchers and an action.
import { Block, BLOCK_LABEL, bucketOf } from './types';
import { HEAD_KEYS, HeadK, Item, Result, bucketLabel, ymLabel } from './engine';

export type BalBlock = 'itc' | 'liab' | 'cash';
export const BAL_BLOCKS: BalBlock[] = ['itc', 'liab', 'cash'];
export const LEDGER_LABEL: Record<Block, string> = { itc: 'Input tax credit', liab: 'Tax payable', cash: 'Electronic cash ledger', fee: 'Interest / late fee' };

export interface ReconEvent {
  id: string; key: string; block: Block; category: Item['category']; explained: boolean;
  what: string; action: string;
  /** effect on the block's closing balance, books − portal (nil for events that reverse within the period) */
  effect: Record<HeadK, number>;
  /** amount involved when there is no effect on the closing (moved between months, netted, fee rows) */
  involved: number;
  affects: boolean; months: string[]; books: string[]; portal: string[]; vouchers: string; rows: string[]; items: Item[];
}
export interface ReconStatement { block: BalBlock; books: Record<HeadK, number>; portal: Record<HeadK, number>; lines: ReconEvent[]; unexplained: Record<HeadK, number>; heads: HeadK[]; gap: Record<HeadK, number> }
export interface Recon { statements: ReconStatement[]; noEffect: ReconEvent[]; events: ReconEvent[] }

const zero = (): Record<HeadK, number> => ({ IGST: 0, CGST: 0, SGST: 0, CESS: 0, ALL: 0 });
const r2 = (x: number) => Math.round(x * 100) / 100;
const nz = (x: number) => Math.abs(x) > 0.004;
const fmt = (x: number) => x.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const COLLAPSE = /not split by tax head, so this row is compared in total/;
/** strip the "<flow row>: " prefix the engine puts on some reasons */
const bare = (it: Item) => { const l = bucketLabel(it.bucket) + ': '; return it.text.startsWith(l) ? it.text.slice(l.length) : it.text; };
const monthsTxt = (ms: string[]) => { const s = [...new Set(ms)].sort().map(ymLabel); return s.length > 2 ? `${s[0]} to ${s[s.length - 1]}` : s.join(' and '); };

function describe(e: ReconEvent, r: Result): [string, string] {
  const it = e.items[0], lab = bucketLabel(it.bucket), row = /^ITC/.test(lab) ? lab : lab.charAt(0).toLowerCase() + lab.slice(1), Row = lab, ms = monthsTxt(e.months);
  const later = [...new Set(e.portal.map(id => r.portal.find(p => p.id === id)).filter(p => p && p.where === 'after').map(p => ymLabel(p!.month)))].join(', ');
  switch (e.category) {
    case 'Opening balance': return ['Opening balance as per books is not the same as the portal\'s opening balance', 'Check the opening under Balances against the trial balance and the portal ledger. A difference carried from last year needs its own reconciliation.'];
    case 'Netted on portal': return [`${bare(it)}${ms ? ` (${ms})` : ''}`, 'None: same amount, shown net on the portal.'];
    case 'Tax head':
      if (COLLAPSE.test(it.text)) return [`The books keep the ${LEDGER_LABEL[e.block].toLowerCase()} in total, the portal by tax head`, 'None: compare the totals.'];
      if (/cross-utilisation/.test(it.text)) return [`${ms}: ITC was set off against a different tax head on the portal than in books (cross-utilisation)`, 'Pass a journal between the input tax accounts so the books follow the portal\'s set-off.'];
      return [`${ms}: ${row} booked under a different tax head in books than in the return`, 'Check the place of supply: correct the books, or amend the invoice in GSTR-1.'];
    case 'Timing': return [`${ms}: ${row} ${bare(it).replace(/^booked/, 'booked')}`, 'None: timing within the period.'];
    case 'After the period':
      return /in books in the period, on the portal/.test(it.text)
        ? [`${Row} in books in the period, on the portal in ${later ? `the ${later} return period` : 'a later return period'}`, `None: timing, clears${later ? ` in ${later}` : ' after the period'}.`]
        : [`${Row} on the portal for a return period in the period, booked in books only after it`, 'None: timing. Check that the books entry date is right.'];
    case 'Fee booked to tax': return [`${ms}: interest / late fee paid on the portal was debited to the GST tax ledger in books`, 'Pass a journal from the tax ledger to the interest / late fee expense account.'];
    case 'In books, not on portal': return [`${ms}: ${row} in books, not on the portal`, e.block === 'itc' ? 'Check: claim it in a later return if eligible, or reverse it in books.' : e.block === 'liab' ? 'Check whether it was reported in GSTR-1 / 3B: report it, or correct the books.' : 'Check the entry against the portal.'];
    case 'On portal, not in books': return [`${ms}: ${row} on the portal, not in books`, 'Check: book the entry, or find out why the portal has it.'];
    case 'Amount differs': {
      const t = HEAD_KEYS.reduce((s, h) => s + e.items.reduce((u, i) => u + i.amt[h], 0), 0);
      return [`${ms}: ${row} in books is ${fmt(Math.abs(t))} ${t > 0 ? 'more' : 'less'} than on the portal`, e.block === 'liab' ? 'Check whether every sale of the month was reported in GSTR-1 / 3B (Show entries).' : e.block === 'itc' ? 'Check the purchases of the month against GSTR-2B and the 3B claim (Show entries).' : 'Check the entries of the month (Show entries).'];
    }
    case 'Not classified': return ['Books lines that could not be placed in a flow row', 'Pick a flow row for each line (Show entries).'];
    case 'Rounding': return ['Rounding', 'None.'];
  }
  return [it.text, ''];
}

export function reconcile(r: Result): Recon {
  const groups = new Map<string, ReconEvent>();
  for (const it of r.items) {
    const pair = it.category === 'Netted on portal' || it.category === 'Fee booked to tax';
    // events never span ledgers, except a fee booked to tax (the tax payable side and the fee row are one event)
    const ledger = it.category === 'Fee booked to tax' ? 'liab' : it.block;
    const key = [ledger, it.category, pair ? it.text : COLLAPSE.test(it.text) ? `collapse|${it.block}` : `${it.bucket}|${it.text}`, it.category === 'Timing' || COLLAPSE.test(it.text) ? '' : it.months.join(',')].join('|');
    let e = groups.get(key);
    if (!e) groups.set(key, e = { id: '', key, block: it.block, category: it.category, explained: it.explained, what: '', action: '', effect: zero(), involved: 0, affects: false, months: [], books: [], portal: [], vouchers: '', rows: [], items: [] });
    e.items.push(it); e.explained &&= it.explained;
    if (e.block === 'fee' && it.block !== 'fee') e.block = it.block;
    if (it.block !== 'fee') { const s = it.bucket === 'opening' ? 1 : bucketOf(it.bucket)?.sign ?? 1; for (const h of HEAD_KEYS) e.effect[h] = r2(e.effect[h] + it.amt[h] * s); }
    e.months.push(...it.months); e.books.push(...it.books); e.portal.push(...it.portal);
    if (!e.rows.includes(bucketLabel(it.bucket))) e.rows.push(bucketLabel(it.bucket));
  }
  const bById = new Map(r.books.map(b => [b.id, b])), pById = new Map(r.portal.map(p => [p.id, p]));
  const events = [...groups.values()];
  for (const e of events) {
    e.months = [...new Set(e.months)].sort(); e.books = [...new Set(e.books)]; e.portal = [...new Set(e.portal)];
    e.affects = HEAD_KEYS.some(h => nz(e.effect[h]));
    e.involved = COLLAPSE.test(e.items[0].text) ? 0 : r2(e.items.reduce((s, it) => s + HEAD_KEYS.reduce((t, h) => t + Math.abs(it.shift[h]) + (e.affects ? 0 : Math.abs(it.amt[h])), 0), 0) / (e.category === 'Netted on portal' || e.category === 'Fee booked to tax' ? 2 : 1));
    [e.what, e.action] = describe(e, r);
    const vs = [...new Set(e.books.map(id => bById.get(id)).filter(Boolean).map(b => `${b!.line.vchType} ${b!.line.vchNo}`.trim() + ` (${b!.line.date})`))];
    const ps = [...new Set(e.portal.map(id => pById.get(id)).filter(Boolean).map(p => `${p!.ref || p!.desc} (${p!.date})`))];
    const list = (a: string[], n: number) => a.length > n ? `${a.slice(0, n).join(', ')} and ${a.length - n} more` : a.join(', ');
    e.vouchers = [vs.length ? `Books: ${list(vs, 3)}` : '', ps.length ? `Portal: ${list(ps, 2)}` : ''].filter(Boolean).join(' · ');
  }
  const order: Item['category'][] = ['Opening balance', 'After the period', 'Tax head', 'Fee booked to tax', 'Netted on portal', 'Timing', 'In books, not on portal', 'On portal, not in books', 'Amount differs', 'Not classified', 'Rounding'];
  const blk: Block[] = ['itc', 'liab', 'cash', 'fee'];
  events.sort((a, b) => +b.affects - +a.affects || blk.indexOf(a.block) - blk.indexOf(b.block) || +b.explained - +a.explained || order.indexOf(a.category) - order.indexOf(b.category) || (a.months[0] || '').localeCompare(b.months[0] || ''));
  events.forEach((e, i) => (e.id = `E${String(i + 1).padStart(2, '0')}`));
  const statements = BAL_BLOCKS.map(b => {
    const books = r.flowBooks.closing[b], portal = r.flowPortal.closing[b];
    const lines = events.filter(e => e.block === b && e.affects);
    const unexplained = zero(), gap = zero();
    for (const h of HEAD_KEYS) {
      unexplained[h] = r2(lines.filter(e => !e.explained).reduce((s, e) => s + e.effect[h], 0));
      gap[h] = r2(books[h] - lines.reduce((s, e) => s + e.effect[h], 0) - portal[h]);
    }
    const heads = HEAD_KEYS.filter(h => nz(books[h]) || nz(portal[h]) || lines.some(e => nz(e.effect[h])));
    return { block: b, books, portal, lines, unexplained, heads: heads.length ? heads : (['IGST', 'CGST', 'SGST'] as HeadK[]), gap };
  });
  return { statements, noEffect: events.filter(e => !e.affects), events };
}
export const ledgerLabel = (b: Block) => LEDGER_LABEL[b] || BLOCK_LABEL[b];
export { fmt as fmtAmt };

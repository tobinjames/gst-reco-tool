// Annual return built from the other tabs' workpapers (Tobin, Nov 2026: users finish Input Reco, Revenue reco and GST
// ledger flow first, then drop those three saved workpapers into the Annual return tab).
// Their source files, settings and choices come back through state.ts; this file reads their headline figures for the
// Workpapers sheet of the GSTR-9 workbook (column B, "as per the workpaper", and F, "where in the workpaper"); the
// template (workbook/build_workbook.py) holds the GSTR-9 side, the difference and what to do.
import type { InFile } from './build';
import { readRows } from './integrity';
import { Reader, V } from './xlsx';
import { Warn, r2 } from './util';
import WP from './wprows.json';

const ROW = WP as Record<string, number>;
export const WORKPAPERS_SHEET = 'Workpapers';
export type WpTab = 'itc' | 'revenue' | 'flow';
export const wpTabOf = (f: InFile): WpTab | null =>
  f.saved && (f.saved.tab === 'itc' || f.saved.tab === 'revenue' || f.saved.tab === 'flow')
    ? f.saved.tab
    : f.kind === 'Revenue workbook'
      ? 'revenue'
      : f.kind === 'Flow workbook'
        ? 'flow'
        : f.kind === 'Workbook' && f.note === 'ITC reconciliation workbook'
          ? 'itc'
          : null;

const num = (v: any) => (typeof v === 'number' && isFinite(v) ? v : v === null || v === undefined || v === '' ? 0 : +v || 0);
const findRow = (rows: any[][], label: string) => rows.find(r => r && String(r[0] ?? '').trim() === label);

/** ITC position from the workpaper's Recon_Purchases rows (values written by the tool; status and amounts as saved) */
function itcFigures(rows: any[][]) {
  const tax = (r: any[], o: number) => [0, 1, 2, 3].reduce((a, i) => a + num(r[o + i]), 0);
  let avail = 0,
    books = 0,
    claim = 0,
    booksOnly = 0,
    twoBOnly = 0;
  for (const r of rows) {
    const st = String(r[0] ?? '');
    if (!st) continue;
    if (r[5] && r[22] !== 'N') avail += tax(r, 14);
    if (r[4] && !st.startsWith('Not compared')) books += tax(r, 9);
    if (st.startsWith('Matched') && r[22] !== 'N')
      claim += [0, 1, 2, 3].reduce((a, i) => a + Math.min(num(r[9 + i]), num(r[14 + i])), 0);
    if (st === 'In books, not in 2B') booksOnly += tax(r, 9);
    if (st === 'In 2B, not in books') twoBOnly += tax(r, 14);
  }
  return { avail: r2(avail), books: r2(books), claim: r2(claim), booksOnly: r2(booksOnly), twoBOnly: r2(twoBOnly) };
}

export async function workpaperCells(files: InFile[], warns: Warn[]): Promise<{ cells: Record<string, V>; loaded: WpTab[] }> {
  const cells: Record<string, V> = {};
  const loaded: WpTab[] = [];
  const put = (key: string, v: number | null, where: string) => {
    if (!ROW[key]) return;
    cells[`B${ROW[key]}`] = v === null ? null : r2(v);
    cells[`F${ROW[key]}`] = v === null ? 'not in this workpaper' : where;
  };
  const head = (t: WpTab, f: InFile) => {
    const key = t === 'revenue' ? 'rev' : t;
    cells[`F${ROW[key]}`] = `${f.name}${f.saved?.savedAt ? ` · saved ${f.saved.savedAt}` : ''}`;
  };
  // the latest workpaper of each tab
  const byTab = new Map<WpTab, InFile>();
  for (const f of files) {
    const t = wpTabOf(f);
    if (!t || !f.bytes) continue;
    const cur = byTab.get(t);
    if (!cur || String(f.saved?.savedAt ?? '') >= String(cur.saved?.savedAt ?? '')) byTab.set(t, f);
  }
  for (const [t, f] of byTab) {
    try {
      const rd = await Reader.open(f.bytes!);
      head(t, f);
      loaded.push(t);
      if (t === 'itc') {
        const x = itcFigures(await readRows(rd, 'Recon_Purchases', 16));
        put('itc_avail', x.avail, 'Recon_Purchases: 2B tax, ITC available');
        put('itc_books', x.books, 'Recon_Purchases: books tax (Summary, ITC as per books)');
        put('itc_claim', x.claim, 'Recon_Purchases: lower of books and 2B on matched invoices (Summary)');
        put('itc_booksonly', x.booksOnly, 'Recon_Purchases: In books, not in 2B');
        put('itc_2bonly', x.twoBOnly, 'Recon_Purchases: In 2B, not in books');
      } else if (t === 'revenue') {
        const sh = rd.has('Summary') ? 'Summary' : 'Dashboard';
        const rows = await readRows(rd, sh, 1);
        const v = (label: string, col: number) => {
          const r = findRow(rows, label);
          return r ? num(r[col]) : null;
        };
        put('rev_books', v('Turnover as per books', 1), `${sh}: Turnover as per books`);
        put('rev_g1', v('Turnover as per GSTR-1', 1), `${sh}: Turnover as per GSTR-1`);
        put('rev_g1tax', v('Turnover as per GSTR-1', 2), `${sh}: Turnover as per GSTR-1, tax`);
        put('rev_3b', v('Turnover as per GSTR-3B', 1), `${sh}: Turnover as per GSTR-3B`);
        put('rev_unx1', v('Unexplained (books to GSTR-1)', 1), `${sh}: Unexplained (books to GSTR-1)`);
        put('rev_unx2', v('Unexplained (GSTR-1 to 3B)', 1), `${sh}: Unexplained (GSTR-1 to 3B)`);
      } else {
        if (!rd.has('Summary')) {
          warns.push({
            level: 'warn',
            file: f.name,
            msg: 'GST ledger flow workpaper from before v0.15 (no Summary sheet): its figures are not on the Workpapers sheet. Drop it into the GST ledger flow tab and build it again.',
          });
          continue;
        }
        const rows = await readRows(rd, 'Summary', 1);
        const blk = (label: string) => findRow(rows, label);
        const itc = blk('Input tax credit'),
          liab = blk('Tax payable'),
          cash = blk('Electronic cash ledger');
        put('flow_itc_books', itc ? num(itc[1]) : null, 'Summary: Input tax credit, as per books');
        put('flow_itc_portal', itc ? num(itc[4]) : null, 'Summary: Input tax credit, as per portal');
        put('flow_itc_unx', itc ? num(itc[5]) : null, 'Summary: Input tax credit, not explained');
        put('flow_liab_unx', liab ? num(liab[5]) : null, 'Summary: Tax payable, not explained');
        put('flow_cash_unx', cash ? num(cash[5]) : null, 'Summary: Electronic cash ledger, not explained');
        const cnt = (label: string) => {
          const r = blk(label);
          return r ? num(r[1]) : null;
        };
        put('flow_fix', cnt('Set-up issues to fix first'), 'Summary: Set-up issues to fix first');
        put('flow_checks', cnt('Cross-checks that do not agree'), 'Summary: Cross-checks that do not agree');
        put('flow_journals', cnt('Suggested journals'), 'Summary: Suggested journals');
      }
    } catch (e: any) {
      warns.push({ level: 'warn', file: f.name, msg: `Could not read the workpaper's figures: ${e.message}` });
    }
  }
  return { cells, loaded };
}

// Revenue reconciliation, books side without a sales register: revenue ledgers picked from a day book / GL dump / detailed
// ledgers (read by flow/books.ts), classified by the user, and turned into invoice-level rows in the Books_Sales layout
// (the layout register.ts normalise() writes and recon.ts reconSales() reads).
import type { BookFile, BookLine } from '../flow/types';
import { guessAccount, headOfName, isGstName } from '../flow/books';
import { serial } from '../xlsxbook';
import { fyOfDate, r2 } from '../util';

/** what a revenue ledger holds */
export type RevClass = 'taxable' | 'export' | 'exempt' | 'nongst' | 'nosupply' | 'none';
export const REV_CLASS_LABEL: Record<RevClass, string> = {
  taxable: 'Taxable sales / income',
  export: 'Exports / SEZ (zero rated)',
  exempt: 'Nil-rated / exempt',
  nongst: 'Non-GST supply',
  nosupply: 'Not a supply: show as excluded',
  none: 'Not revenue: ignore the ledger',
};
export const REV_CLASSES = Object.keys(REV_CLASS_LABEL) as RevClass[];

const SALES =
  /\b(sales?|turnover|revenue|income|receipts?|commission|brokerage|service charges?|fees? received|job ?work|rent(al)?|scrap|export)\b/i;
const NOT_REVENUE =
  /\b(purchases?|expenses?|cost|payable|receivable|debtors?|creditors?|bank|cash|capital|loan|advance|deposit|provision|reserve|suspense|depreciation|salary|wages|tds|tcs)\b/i;
const NO_SUPPLY =
  /\b(interest|dividend|forex|foreign exchange|exchange (gain|difference|fluctuation)|discount (received|earned)|rebate|profit on sale of (investment|shares|mutual)|gain on|insurance claim|bad debts? recovered|round(ing)? off|subsidy|grant)\b/i;

/** a suggested class for an account that looks like revenue, else null (not shown by default) */
export function guessRevenue(name: string, group = ''): RevClass | null {
  if (isGstName(name)) return null;
  const g = group.toLowerCase();
  const income = /income/.test(g) && !/expense/.test(g);
  if (!income && (!SALES.test(name) || NOT_REVENUE.test(name)) && !NO_SUPPLY.test(name)) return null;
  if (NO_SUPPLY.test(name)) return income || /income|gain|received|earned|interest|dividend/i.test(name) ? 'nosupply' : null;
  if (/\b(export|sez|zero[- ]?rated|overseas)\b/i.test(name)) return 'export';
  if (/\b(exempt|nil[- ]?rated|nil)\b/i.test(name)) return 'exempt';
  if (/\bnon[- ]?gst\b/i.test(name)) return 'nongst';
  return 'taxable';
}

const isOutputTax = (n: string) => {
  if (!isGstName(n)) return false;
  const g = guessAccount(n);
  return g.type === 'output' && !!headOfName(n);
};

export interface GlDoc {
  /** Books_Sales layout row + [16] ledgers, [17] class, [18] voucher type, [19] file/rows */
  row: any[];
  cls: RevClass;
}
export interface GlResult {
  docs: GlDoc[];
  /** revenue lines used, for the Books_Lines sheet */
  lines: { l: BookLine; cls: RevClass; doc: string }[];
  /** totals per ledger (all revenue-looking or classified ledgers), for the ledger card and the Ledgers sheet */
  ledgers: { name: string; cls: RevClass | null; guess: RevClass | null; dr: number; cr: number; lines: number }[];
  outsideFy: number;
  /** vouchers with output tax but no revenue line (e.g. tax-only journals) */
  taxOnly: number;
}

/** turn GL lines into invoice-level rows: one row per voucher and class; tax from the output GST lines of the voucher */
export function glDocs(files: BookFile[], classes: Record<string, RevClass>, fy: string): GlResult {
  const lines = files.flatMap(f => f.lines);
  const ledgers = new Map<string, GlResult['ledgers'][number]>();
  for (const l of lines) {
    const guess = guessRevenue(l.account, l.accGroup);
    const cls = classes[l.account] ?? guess;
    if (!cls && !guess) continue;
    let a = ledgers.get(l.account);
    if (!a) ledgers.set(l.account, (a = { name: l.account, cls: cls ?? null, guess, dr: 0, cr: 0, lines: 0 }));
    a.dr = r2(a.dr + l.dr);
    a.cr = r2(a.cr + l.cr);
    a.lines++;
  }
  const clsOf = (acc: string): RevClass | null => {
    const c = classes[acc] ?? ledgers.get(acc)?.guess ?? null;
    return c && c !== 'none' ? c : null;
  };
  const byV = new Map<string, BookLine[]>();
  for (const l of lines) {
    const k = `${l.file}|${l.vkey}`;
    if (!byV.has(k)) byV.set(k, []);
    byV.get(k)!.push(l);
  }
  // single-ledger exports: each file is one ledger, so join the lines of one voucher across files by date + voucher no.
  if (files.every(f => f.format === 'ledger') && files.length > 1) {
    byV.clear();
    for (const l of lines) {
      const k = `${l.date}|${l.vchNo || l.file + '#' + l.row}`;
      if (!byV.has(k)) byV.set(k, []);
      byV.get(k)!.push(l);
    }
  }
  const out: GlResult = { docs: [], lines: [], ledgers: [...ledgers.values()], outsideFy: 0, taxOnly: 0 };
  for (const [, vl] of byV) {
    const rev = vl.filter(l => clsOf(l.account));
    const tax = vl.filter(l => isOutputTax(l.account));
    if (!rev.length) {
      if (tax.length) out.taxOnly++;
      continue;
    }
    const f = vl[0];
    if (fyOfDate(f.date) !== fy) {
      out.outsideFy++;
      continue;
    }
    const party =
      vl.find(l => l.party)?.party ||
      vl
        .filter(l => !clsOf(l.account) && !isGstName(l.account))
        .sort((a, b) => b.dr - a.dr)
        .map(l => l.account)[0] ||
      '';
    const vt = f.vchType.toLowerCase();
    const classes2 = [...new Set(rev.map(l => clsOf(l.account)!))];
    classes2.forEach((c, ci) => {
      const ls = rev.filter(l => clsOf(l.account) === c);
      const tx = r2(ls.reduce((s, l) => s + l.cr - l.dr, 0));
      // tax goes with the taxable (or export) part of the voucher
      const taxHere =
        ci === classes2.findIndex(x => x === 'taxable' || x === 'export') ||
        (ci === 0 && !classes2.some(x => x === 'taxable' || x === 'export'));
      const h = { CGST: 0, SGST: 0, IGST: 0, CESS: 0 } as Record<string, number>;
      if (taxHere) for (const t of tax) h[headOfName(t.account)!] = r2(h[headOfName(t.account)!] + t.cr - t.dr);
      const tot = tx + h.CGST + h.SGST + h.IGST + h.CESS;
      const dt = /credit note|sales return|rejection/.test(vt) || tot < 0 ? 'CN' : /debit note/.test(vt) ? 'DN' : 'INV';
      const taxSum = h.CGST + h.SGST + h.IGST + h.CESS;
      const type =
        c === 'export'
          ? taxSum
            ? 'EXP-WP'
            : 'EXP-WOP'
          : c === 'exempt'
            ? 'EXEMPT'
            : c === 'nongst'
              ? 'NON-GST'
              : c === 'nosupply'
                ? 'NOSUPPLY'
                : 'B2B';
      const no = (f.vchNo || f.ref || `${f.date}/${f.row}`) + (classes2.length > 1 && ci ? `-${c}` : '');
      const ab = (x: number) => r2(Math.abs(x));
      const row = [
        serial(f.date) ?? f.date,
        no,
        party,
        '',
        '',
        type,
        dt,
        '',
        '',
        null,
        tx ? Math.round((Math.abs(taxSum) / Math.abs(tx)) * 1000) / 10 : null,
        ab(tx),
        ab(h.CGST),
        ab(h.SGST),
        ab(h.IGST),
        ab(h.CESS),
        [...new Set(ls.map(l => l.account))].join(', '),
        c,
        f.vchType,
        `${f.file} rows ${vl.map(l => l.row).join(', ')}`,
      ];
      out.docs.push({ row, cls: c });
      for (const l of ls) out.lines.push({ l, cls: c, doc: no });
    });
  }
  out.ledgers.sort(
    (a, b) => +(b.cls !== null && b.cls !== 'none') - +(a.cls !== null && a.cls !== 'none') || b.cr - b.dr - (a.cr - a.dr),
  );
  return out;
}

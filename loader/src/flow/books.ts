// Books side of the GST ledger flow: read a GL dump (Tally voucher dump, Zoho Detailed General Ledger) or a single-ledger export
// (Tally Ledger Vouchers, Zoho Account Transactions, any "Date · Particulars · Vch type · Debit · Credit" layout), pick the GST accounts by name.
import { AccHead, AccType, BookFile, BookLine, Head } from './types';
import { ddmmyyyy } from './portal';

const txt = (v: any) =>
  String(v ?? '')
    .replace(/\s+/g, ' ')
    .trim();
const low = (v: any) => txt(v).toLowerCase();
const key = (v: any) =>
  low(v)
    .replace(/[\s_.]+/g, ' ')
    .replace(/[()₹#]/g, '')
    .trim();
const num = (v: any) => {
  if (typeof v === 'number') return v;
  let s = txt(v);
  if (!s || s === '-') return 0;
  const neg = /^\(.*\)$/.test(s);
  s = s.replace(/[^\d.-]/g, '');
  const x = parseFloat(s);
  return isFinite(x) ? (neg ? -x : x) : 0;
};
const r2 = (x: number) => Math.round(x * 100) / 100;

/** vouchers that do not post to the books (Tally keeps them in the same dump) */
export const NON_ACCOUNTING =
  /\b(purchase order|sales order|delivery note|receipt note|rejections? (in|out)|memorandum|memo\b|physical stock|stock journal|material (in|out)|job work|quotation|reversing journal|optional)/i;

type Fmt = BookFile['format'];
interface Layout {
  fmt: Fmt;
  header: number;
  col: Record<string, number>;
}

function layoutOf(rows: any[][]): Layout | null {
  for (let i = 0; i < Math.min(40, rows.length); i++) {
    const H = (rows[i] || []).map(key);
    const at = (...names: string[]) => H.findIndex(h => names.includes(h));
    if (
      at('gl account', 'gl_account', 'ledger name', 'ledgername') >= 0 &&
      at('amount') >= 0 &&
      at('vch type', 'vch_type', 'voucher type', 'vouchertypename') >= 0
    )
      return {
        fmt: 'tally-dump',
        header: i,
        col: {
          account: at('gl account', 'ledger name', 'ledgername'),
          amount: at('amount'),
          vchType: at('vch type', 'voucher type', 'vouchertypename'),
          vkey: at('voucher id', 'vch key', 'voucher key', 'masterid', 'guid'),
          date: at('date'),
          vchNo: at('vouchernumber', 'voucher number', 'vch no', 'voucher no'),
          party: at('partyname', 'party name', 'party ledger name'),
          narration: at('narration'),
          ref: at('reference', 'supplier invoice no', 'bill no'),
        },
      };
    if (at('account name', 'account_name') >= 0 && at('debit') >= 0 && at('credit') >= 0)
      return {
        fmt: 'zoho-gl',
        header: i,
        col: {
          account: at('account name'),
          debit: at('debit'),
          credit: at('credit'),
          vchType: at('transaction type'),
          vkey: at('transaction id'),
          date: at('date'),
          vchNo: at('entity number', 'transaction number', 'transaction#'),
          party: at('transaction details', 'contact name'),
          narration: at('description', 'notes'),
          group: at('account type'),
          ref: at('reference number'),
        },
      };
    const d = at('date', 'voucher date', 'txn date', 'transaction date'),
      dr = H.findIndex(h => /^(debit|dr)( amount)?$/.test(h)),
      cr = H.findIndex(h => /^(credit|cr)( amount)?$/.test(h));
    const part = at('particulars', 'account', 'transaction details', 'details', 'description', 'ledger');
    if (d >= 0 && dr >= 0 && cr >= 0 && part >= 0 && !H.some(h => /taxable|invoice value|hsn/.test(h)))
      return {
        fmt: 'ledger',
        header: i,
        col: {
          date: d,
          debit: dr,
          credit: cr,
          party: part,
          vchType: at('vch type', 'voucher type', 'transaction type', 'type'),
          vchNo: at('vch no', 'voucher no', 'vch no.', 'transaction#', 'transaction number', 'reference', 'ref no'),
          narration: at('narration', 'notes'),
        },
      };
  }
  return null;
}
export const looksLikeBooks = (rows: any[][]) => !!layoutOf(rows);

export function parseBooks(name: string, rows: any[][], ledgerName?: string): BookFile {
  const lay = layoutOf(rows);
  if (!lay)
    throw new Error(
      'This file does not look like a ledger or GL export. It needs a Date column, an Account or Particulars column and Debit and Credit (or Amount) columns. Open it in Excel and check the header row',
    );
  const { fmt, header, col } = lay;
  const out: BookFile = { name, format: fmt, lines: [], skipped: [], problems: [] };
  const skip = (reason: string) => {
    const s = out.skipped.find(x => x.reason === reason);
    if (s) s.count++;
    else out.skipped.push({ reason, count: 1 });
  };
  const g = (r: any[], k: string) => (col[k] !== undefined && col[k] >= 0 ? r[col[k]] : undefined);
  let lastDate = '',
    lastKey = '',
    seq = 0;
  if (fmt === 'ledger') {
    out.ledger = ledgerName || titleLedger(rows.slice(0, header)) || name.replace(/\.(xlsx|csv)$/i, '');
  }
  for (let i = header + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    if (!r.some(v => txt(v))) continue;
    if (fmt === 'tally-dump') {
      const vt = txt(g(r, 'vchType')),
        acc = txt(g(r, 'account'));
      if (!acc) continue;
      if (NON_ACCOUNTING.test(vt)) {
        skip(`"${vt}" vouchers do not post to the books`);
        continue;
      }
      const a = num(g(r, 'amount'));
      const date = ddmmyyyy(g(r, 'date'));
      if (!date) {
        skip('no date');
        continue;
      }
      out.lines.push({
        file: name,
        row: i + 1,
        vkey: txt(g(r, 'vkey')) || `${date}|${txt(g(r, 'vchNo'))}|${vt}`,
        date,
        vchType: vt,
        vchNo: txt(g(r, 'vchNo')),
        party: txt(g(r, 'party')),
        account: acc,
        narration: txt(g(r, 'narration')),
        dr: a < 0 ? r2(-a) : 0,
        cr: a > 0 ? r2(a) : 0,
        accGroup: '',
        ref: txt(g(r, 'ref')) || undefined,
      });
    } else if (fmt === 'zoho-gl') {
      const acc = txt(g(r, 'account'));
      if (!acc) continue;
      const date = ddmmyyyy(g(r, 'date'));
      if (!date) {
        skip('no date');
        continue;
      }
      const dr = num(g(r, 'debit')),
        cr = num(g(r, 'credit'));
      out.lines.push({
        file: name,
        row: i + 1,
        vkey: txt(g(r, 'vkey')) || `${date}|${txt(g(r, 'vchNo'))}`,
        date,
        vchType: txt(g(r, 'vchType')),
        vchNo: txt(g(r, 'vchNo')),
        party: txt(g(r, 'party')),
        account: acc,
        narration: txt(g(r, 'narration')),
        dr: r2(dr),
        cr: r2(cr),
        accGroup: low(g(r, 'group')),
        ref: txt(g(r, 'ref')) || undefined,
      });
    } else {
      const part = txt(g(r, 'party')),
        dr = num(g(r, 'debit')),
        cr = num(g(r, 'credit'));
      if (/^opening balance/i.test(part) || /^opening balance/i.test(txt(r.join(' ')).slice(0, 40))) {
        out.opening = r2(dr - cr);
        continue;
      }
      if (/^closing balance/i.test(part)) {
        out.closing = r2(dr - cr);
        continue;
      }
      if (/^(total|grand total|current total)/i.test(part)) continue;
      let date = ddmmyyyy(g(r, 'date'));
      // Tally prints the date once for several lines of the same day, and wraps long narrations into extra rows
      if (!date && (dr || cr)) date = lastDate;
      if (!date) continue;
      if (!dr && !cr) continue;
      lastDate = date;
      const vt = txt(g(r, 'vchType'));
      if (NON_ACCOUNTING.test(vt)) {
        skip(`"${vt}" vouchers do not post to the books`);
        continue;
      }
      const vno = txt(g(r, 'vchNo'));
      const vk = `${date}|${vno || ++seq}|${vt}`;
      lastKey = vk;
      out.lines.push({
        file: name,
        row: i + 1,
        vkey: vk,
        date,
        vchType: vt,
        vchNo: vno,
        party: part,
        account: out.ledger!,
        narration: txt(g(r, 'narration')),
        dr: r2(dr),
        cr: r2(cr),
        accGroup: '',
      });
    }
  }
  void lastKey;
  if (!out.lines.length) out.problems.push('No ledger lines found');
  return out;
}

/** single-ledger exports name the ledger in the title rows ("Ledger: Input CGST", or the line under the company name) */
function titleLedger(top: any[][]): string {
  const cells = top.flat().map(txt).filter(Boolean);
  for (const c of cells) {
    const m = c.match(/^(ledger|account)\s*[:\-]\s*(.+)$/i);
    if (m) return m[2].trim();
  }
  const g = cells.find(c => isGstName(c));
  return g || '';
}

// ---------- GST accounts by name
const HEAD_RE: [RegExp, Head][] = [
  [/\b(igst|integrated)\b/i, 'IGST'],
  [/\b(cgst|central)\b/i, 'CGST'],
  [/\b(sgst|utgst|state|ut)\b/i, 'SGST'],
  [/\bcess\b/i, 'CESS'],
];
export const headOfName = (n: string): Head | null => {
  for (const [re, h] of HEAD_RE) if (re.test(n)) return h;
  return null;
};
const VALUE_LEDGER = /\b(sales?|purchases?|services?|income|turnover|revenue|stock|inventory|freight|discount)\b/i;
const TAX_WORD = /\b(input|output|itc|payable|receivable|credit|liability|rcm|reverse charge|tax (paid|payable))\b/i;
const GST_WORD = /\b(gst|igst|cgst|sgst|utgst|cess|itc|input tax|output tax)\b/i;

export function isGstName(n: string) {
  if (/\btds\b|\btcs\b/i.test(n) && !/gst/i.test(n)) return false;
  if (/electronic cash|cash ledger|gst (cash|deposit)|pmt-?06/i.test(n)) return true;
  if (/late fee/i.test(n)) return true;
  if (!GST_WORD.test(n)) return false;
  if (VALUE_LEDGER.test(n) && !TAX_WORD.test(n)) return false; // "IGST Sales 12%", "Local Purchases 12%"
  return true;
}
export interface AccGuess {
  type: AccType;
  head: AccHead | '';
}
export function guessAccount(n: string): AccGuess {
  const head = headOfName(n);
  if (/electronic cash|cash ledger|gst (cash|deposit)|pmt-?06/i.test(n)) return { type: 'cash_ledger', head: head || 'ALL' };
  if (/late fee/i.test(n)) return { type: 'late_fee', head: head || 'ALL' };
  if (/\binterest\b/i.test(n)) return { type: 'interest', head: head || 'ALL' };
  if (/\bpenalty\b/i.test(n)) return { type: 'penalty', head: head || 'ALL' };
  if (/\brefund\b/i.test(n)) return { type: 'refund_recv', head: head || 'ALL' };
  if (/\b(rcm|reverse charge)\b/i.test(n) && /\b(payable|output|liability)\b/i.test(n))
    return { type: 'rcm_payable', head: head || '' };
  if (/\b(input|itc|receivable|credit)\b/i.test(n)) return { type: 'input', head: head || '' };
  if (/\boutput\b/i.test(n)) return { type: 'output', head: head || '' };
  if (/\b(payable|liability)\b/i.test(n)) return { type: head ? 'output' : 'payable', head: head || 'ALL' };
  return { type: head ? 'output' : 'ignore', head: head || '' };
}

/** what the other side of a GST entry is, from the account name (and Zoho's account type) */
export type Contra = 'bank' | 'party' | 'expense' | 'income' | 'other';
export function contraOf(name: string, group: string): Contra {
  const g = group.toLowerCase();
  if (/^(bank|cash)$/.test(g)) return 'bank';
  if (/accounts_(payable|receivable)/.test(g)) return 'party';
  if (/expense|cost_of_goods/.test(g)) return 'expense';
  if (/income/.test(g)) return 'income';
  if (/electronic cash/i.test(name)) return 'other';
  if (
    /\b(bank|hdfc|icici|sbi|axis|kotak|idfc|indusind|federal|canara|yes bank|cash in hand|petty cash|cash)\b/i.test(name) ||
    /\bcash\b/i.test(name)
  )
    return 'bank';
  if (/\b(sales?|income|revenue|turnover)\b/i.test(name)) return 'income';
  if (
    /\b(purchases?|expenses?|charges|fees?|rent|salary|wages|repairs|freight|travel|professional|consultancy|depreciation|write.?off|loss)\b/i.test(
      name,
    )
  )
    return 'expense';
  return 'party';
}

/** accounts in the loaded books with their totals, GST-looking ones first */
export function accountsOf(files: BookFile[]) {
  const m = new Map<string, { name: string; lines: number; dr: number; cr: number; files: Set<string>; gst: boolean }>();
  for (const f of files)
    for (const l of f.lines) {
      const k = l.account;
      let a = m.get(k);
      if (!a) m.set(k, (a = { name: k, lines: 0, dr: 0, cr: 0, files: new Set(), gst: isGstName(k) || f.format === 'ledger' }));
      a.lines++;
      a.dr = r2(a.dr + l.dr);
      a.cr = r2(a.cr + l.cr);
      a.files.add(f.name);
    }
  return [...m.values()].sort((a, b) => +b.gst - +a.gst || a.name.localeCompare(b.name));
}
export type { BookLine };

// GST ledger flow: shared types
export const HEADS = ['IGST', 'CGST', 'SGST', 'CESS'] as const;
export type Head = typeof HEADS[number];
export const HEAD_LABEL: Record<Head, string> = { IGST: 'Integrated', CGST: 'Central', SGST: 'State/UT', CESS: 'Cess' };
export const COMPONENTS = ['tax', 'interest', 'penalty', 'fee', 'others'] as const;
export type Component = typeof COMPONENTS[number];

/** a period as YYYY-MM (sorts as text) */
export type Ym = string;

export type PortalKind = 'cash' | 'credit' | 'liability';
export const PORTAL_LABEL: Record<PortalKind, string> = { cash: 'Electronic Cash Ledger', credit: 'Electronic Credit Ledger', liability: 'Electronic Liability Register' };

export interface PortalLine {
  row: number;               // 1-based row in the file
  sr: string; date: string;  // dd-mm-yyyy
  ref: string; taxPeriod: Ym | ''; desc: string;
  dir: 'Dr' | 'Cr';
  used: string;              // liability register: ledger used for discharging (Cash / Credit / -)
  /** amount per head and component (credit ledger: tax only) */
  amt: Record<Head, Record<Component, number>>;
  bal: Record<Head, number>;  // running balance after the line (total of components)
}
export interface PortalLedger {
  kind: PortalKind; gstin: string; legalName: string;
  from: string; to: string;    // dd-mm-yyyy (liability register: first and last day of its tax period range)
  opening?: Record<Head, number>; closing?: Record<Head, number>;
  lines: PortalLine[]; problems: string[];
}

/** what a books account is, for the flow */
export type AccType = 'input' | 'output' | 'rcm_payable' | 'payable' | 'cash_ledger' | 'refund_recv' | 'interest' | 'late_fee' | 'penalty' | 'ignore';
export const ACC_TYPE_LABEL: Record<AccType, string> = {
  input: 'Input tax (ITC)', output: 'Output tax', rcm_payable: 'RCM payable', payable: 'GST payable (control)', cash_ledger: 'Electronic cash ledger / GST deposit',
  refund_recv: 'GST refund receivable', interest: 'GST interest (expense)', late_fee: 'GST late fee (expense)', penalty: 'GST penalty (expense)', ignore: 'Not a GST account',
};
/** head 'ALL' = the account is not split by head (e.g. one "GST cash ledger" account): it counts in the Total column only */
export type AccHead = Head | 'ALL';

/** one line of a books ledger / GL dump */
export interface BookLine {
  file: string; row: number; vkey: string;   // vkey groups the lines of one voucher
  date: string; vchType: string; vchNo: string; party: string; account: string; narration: string;
  dr: number; cr: number;
  /** Zoho gives the account type (bank, cash, expense…), Tally does not */
  accGroup: string;
}
export interface BookFile {
  name: string; format: 'tally-dump' | 'zoho-gl' | 'ledger';
  lines: BookLine[];
  /** single-ledger exports: the ledger name and the opening / closing printed in the file */
  ledger?: string; opening?: number; closing?: number;
  skipped: { reason: string; count: number }[];
  problems: string[];
}

export type Block = 'itc' | 'liab' | 'cash' | 'fee';
export interface BucketDef { id: string; block: Block; label: string; sign: 1 | -1 }
/** flow rows: sign = effect on the block's closing balance */
export const BUCKETS: BucketDef[] = [
  { id: 'itc_availed', block: 'itc', label: 'ITC availed (net of debit notes)', sign: 1 },
  { id: 'itc_rcm', block: 'itc', label: 'ITC on reverse charge', sign: 1 },
  { id: 'itc_reclaimed', block: 'itc', label: 'ITC reclaimed', sign: 1 },
  { id: 'itc_reversed', block: 'itc', label: 'ITC reversed', sign: -1 },
  { id: 'itc_setoff', block: 'itc', label: 'ITC used to pay tax (set-off)', sign: -1 },
  { id: 'itc_refund', block: 'itc', label: 'ITC refunded', sign: -1 },
  { id: 'out_tax', block: 'liab', label: 'Output tax (net of credit notes)', sign: 1 },
  { id: 'out_rcm', block: 'liab', label: 'Tax on reverse charge', sign: 1 },
  { id: 'out_cn', block: 'liab', label: 'Credit notes / sales returns', sign: -1 },
  { id: 'paid_itc', block: 'liab', label: 'Paid by ITC', sign: -1 },
  { id: 'paid_cash', block: 'liab', label: 'Paid in cash', sign: -1 },
  { id: 'cash_deposit', block: 'cash', label: 'Deposited (challans)', sign: 1 },
  { id: 'cash_used', block: 'cash', label: 'Used to pay tax, interest and fees', sign: -1 },
  { id: 'cash_refund', block: 'cash', label: 'Refunded', sign: -1 },
  { id: 'fee_interest', block: 'fee', label: 'Interest', sign: 1 },
  { id: 'fee_latefee', block: 'fee', label: 'Late fee', sign: 1 },
  { id: 'fee_penalty', block: 'fee', label: 'Penalty and others', sign: 1 },
];
export const BLOCK_LABEL: Record<Block, string> = { itc: 'Input tax credit (asset)', liab: 'Tax liability', cash: 'Electronic cash ledger', fee: 'Interest, late fee and penalty' };
export const bucketOf = (id: string) => BUCKETS.find(b => b.id === id);
/** books lines that only move between GST accounts of the same block (e.g. output → GST payable) */
export const INTERNAL = 'internal';
export const UNMAPPED = 'unmapped';

export const zeroHeads = (): Record<Head, number> => ({ IGST: 0, CGST: 0, SGST: 0, CESS: 0 });
export const r2 = (x: number) => Math.round(x * 100) / 100;

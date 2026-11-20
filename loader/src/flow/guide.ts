// How to export each report for the GST ledger flow (shown in the tab and on the workbook's Read Me)
export interface GuideItem {
  title: string;
  source: 'portal' | 'tally' | 'zoho' | 'other';
  steps: string[];
  note?: string;
}
export const FLOW_GUIDE: GuideItem[] = [
  {
    source: 'portal',
    title: 'Electronic Credit Ledger',
    steps: [
      'Log in at gst.gov.in → Services → Ledgers → Electronic Credit Ledger.',
      'From: 1 April of the year (so the Opening Balance row is the opening of the year). To: a date after the March return was filed, e.g. 31 May of the next year, so March ITC and its set-off are included.',
      'Click GO, then download the CSV / Excel file. Drop it here as it is.',
    ],
  },
  {
    source: 'portal',
    title: 'Electronic Cash Ledger',
    steps: [
      'Services → Ledgers → Electronic Cash Ledger.',
      'Same dates as the credit ledger (1 April to after the March return was filed).',
      'GO → download the CSV / Excel file.',
    ],
  },
  {
    source: 'portal',
    title: 'Electronic Liability Register',
    steps: [
      'Services → Ledgers → Electronic Liability Register → Part I: Return related liabilities.',
      'Tax period from April to March of the year → GO → download the CSV / Excel file.',
    ],
    note: 'The register has no opening balance. If a liability of the previous year was unpaid on 1 April, enter it under Balances.',
  },
  {
    source: 'tally',
    title: 'GST ledgers, one by one (standard Tally)',
    steps: [
      'Gateway of Tally → Display More Reports → Account Books → Ledger → pick the ledger (Input CGST, Input SGST, Input IGST, Output CGST …, RCM payable, GST electronic cash ledger, GST interest / late fee).',
      'Alt+F2: set the period to the financial year.',
      'Ctrl+E (Export) → format Excel (Spreadsheet) → Send. The export carries the opening and closing balance.',
      'Repeat for every GST ledger, including rate-wise ledgers (Output CGST 6%, 9% …): they are added up by tax head.',
    ],
    note: 'Each line shows the other ledger of the voucher under Particulars: that is what decides the flow row (sale, purchase, set-off, payment).',
  },
  {
    source: 'tally',
    title: 'Full voucher / GL dump',
    steps: [
      'If your firm has a voucher-dump export (one row per ledger line with voucher type, date, ledger and amount, Dr shown as minus), export the whole year and drop it here.',
      'The tool lists the GST-looking ledgers by name; confirm the type and tax head of each.',
      'Non-accounting vouchers in the dump (purchase / sales orders, delivery and receipt notes, memorandum) are left out.',
    ],
    note: 'A dump has no opening balances: enter them under Balances from the trial balance of 31 March of the previous year.',
  },
  {
    source: 'zoho',
    title: 'Detailed General Ledger',
    steps: [
      'Reports → Accountant → Detailed General Ledger.',
      'Date range: the financial year. Export → XLSX.',
      'The tool picks the GST accounts (Input / Output CGST, SGST, IGST …) by name for you to confirm.',
    ],
    note: 'The export has no opening balances: take them from Reports → Trial Balance as on 31 March of the previous year and enter them under Balances.',
  },
  {
    source: 'zoho',
    title: 'Account Transactions (one account)',
    steps: [
      'Reports → Account Transactions → filter the account (e.g. Input CGST), date range the financial year.',
      'Export → XLSX. Repeat for each GST account.',
    ],
  },
  {
    source: 'other',
    title: 'Any other software',
    steps: [
      'Export each GST ledger (or the whole general ledger) to .xlsx or .csv with: Date, Particulars / Account, Voucher type, Voucher no., Debit, Credit.',
      'Title rows above the header are fine; "Opening Balance" and "Closing Balance" rows are read when present.',
    ],
  },
];

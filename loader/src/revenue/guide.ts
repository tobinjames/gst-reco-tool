// How to export each report for the revenue reconciliation (shown in the tab and on the workbook's Read Me).
// Menu paths as of November 2026: recheck them when the portal or the software changes.
import type { GuideItem } from '../flow/guide';

export const REV_GUIDE: GuideItem[] = [
  {
    source: 'portal',
    title: 'GSTR-1 (every month of the year)',
    steps: [
      'gst.gov.in → Services → Returns → Returns Dashboard → pick the year and the month → Search.',
      'On the GSTR-1 / IFF tile click Download → Generate JSON file to download. The link appears after a few minutes (refresh the page).',
      'Download the ZIP and drop it here as it is. Do this for April to March, and for April to October of the next year to catch invoices reported late.',
    ],
    note: 'Quarterly (QRMP) filers: one GSTR-1 per quarter (plus IFF months if used).',
  },
  {
    source: 'portal',
    title: 'GSTR-3B (every month of the year)',
    steps: [
      'Returns Dashboard → pick the year and the month → Search.',
      'On the GSTR-3B tile of a filed period click Download to save the PDF.',
      'One PDF per month (per quarter for QRMP filers). Drop them all here.',
    ],
  },
  {
    source: 'tally',
    title: 'Sales register (preferred)',
    steps: [
      'Gateway of Tally → Display More Reports → Account Books → Sales Register.',
      'Alt+F2: set the period to 1 April to 31 March.',
      'Switch to Columnar (Alt+F5). Turn on Voucher No., GSTIN/UIN, Gross Total and keep all ledger columns.',
      'Export (Ctrl+E) → Excel (Spreadsheet) → Send. Do the same for the Credit Note Register (and Debit Note Register, if used).',
      'Drop the files here, then check the column mapping under "Map register columns".',
    ],
  },
  {
    source: 'tally',
    title: 'Day book / GL dump (only if there is no proper sales register)',
    steps: [
      'Use the same voucher dump as the GST ledger flow tab: one row per ledger line with voucher id, voucher type, date, voucher number, party, ledger (GL account) and amount.',
      'Cover the whole year (1 April to 31 March).',
      'Drop it here: the tab lists the ledgers that look like revenue (sales, other income). Tick or add the ledgers that are revenue, and say what each one holds.',
    ],
    note: 'Other income such as interest or dividends is offered as "not a supply"; you can ignore a ledger altogether, or single entries.',
  },
  {
    source: 'tally',
    title: 'Detailed ledgers (alternative to the dump)',
    steps: [
      'Display More Reports → Account Books → Ledger → pick each sales / income ledger and each Output CGST / SGST / IGST / Cess ledger.',
      'Alt+F2: 1 April to 31 March. Alt+F1 (or F12 Configure) for detailed view with voucher numbers.',
      'Export each one to Excel and drop them all here. Lines of one voucher are joined by date and voucher number.',
    ],
  },
  {
    source: 'zoho',
    title: 'Invoices and credit notes (preferred)',
    steps: [
      'Sales → Invoices → menu (☰) → Export Invoices → date range of the year → XLSX → Export.',
      'Sales → Credit Notes → menu (☰) → Export Credit Notes, same dates.',
      'Drop both here and check the column mapping.',
    ],
  },
  {
    source: 'zoho',
    title: 'Detailed general ledger (only if invoices cannot be exported)',
    steps: [
      'Reports → Accountant → General Ledger → Detailed view (or Account Transactions for each income account).',
      'Date range: the year. Export as XLSX.',
      'Drop it here and confirm which income accounts are revenue.',
    ],
  },
  {
    source: 'tally',
    title: 'Customer list: name + GSTIN (optional, when the books have no GSTIN)',
    steps: [
      'Gateway of Tally → Display More Reports → List of Accounts → Alt+F5 detailed / F12 Configure: show GSTIN/UIN. Or Multi Masters → Multi Alter for Sundry Debtors.',
      'Export (Ctrl+E) → Excel. Any sheet with a customer name column and a GSTIN column works.',
      'Drop it here with the other files: names in the books are matched to the GSTIN on the list.',
    ],
  },
  {
    source: 'zoho',
    title: 'Customer list: name + GSTIN (optional)',
    steps: [
      'Sales → Customers → menu (☰) → Export Customers → XLSX or CSV.',
      'Drop it here: Display Name / Company Name and the GSTIN column are read.',
    ],
  },
  {
    source: 'other',
    title: 'Any other software or Excel',
    steps: [
      'Export the sales register (for Busy, Marg and others, the GSTR-1 register) to .xlsx or .csv: one row per invoice or per item line.',
      'Needed: invoice date, invoice no., taxable value. Strongly recommended: customer name, GSTIN, place of supply, CGST, SGST, IGST, cess and the document type (invoice / credit note).',
      'If only a ledger dump is available, it needs date, voucher no., account and debit / credit (or amount) on every line.',
      'If customer GSTINs are missing in the books, add a two-column sheet: Customer name, GSTIN.',
    ],
  },
];

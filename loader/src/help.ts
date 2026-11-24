// Help & guide: the in-app manual (how to use the tool, how each tab matches and explains, how to export each report,
// how to save and reuse a workpaper). The page is plain HTML built here; the export steps come from the same lists the
// tabs show (flow/guide.ts, revenue/guide.ts), so they cannot drift apart. Menu paths are as of November 2026.
import { FLOW_GUIDE, GuideItem } from './flow/guide';
import { REV_GUIDE } from './revenue/guide';
import { REASONS } from './revenue/engine';
import { MATCH_PRESETS } from './recon';
import { esc } from './util';

const SOURCE_NAME: Record<GuideItem['source'], string> = {
  portal: 'GST portal',
  tally: 'Tally Prime',
  zoho: 'Zoho Books',
  other: 'Other software / Excel',
};

function guideList(items: GuideItem[]): string {
  return (['portal', 'tally', 'zoho', 'other'] as const)
    .map(src => {
      const rows = items.filter(i => i.source === src);
      if (!rows.length) return '';
      return `<h4>${SOURCE_NAME[src]}</h4>${rows
        .map(
          g =>
            `<div class="hp-item"><b>${esc(g.title)}</b><ol>${g.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ol>${
              g.note ? `<p class="hp-note">${esc(g.note)}</p>` : ''
            }</div>`,
        )
        .join('')}`;
    })
    .join('');
}

const presetRow = (name: string) => {
  const p = MATCH_PRESETS[name];
  const on = Object.values(p.rules).filter(Boolean).length;
  return `<tr><td><b>${name}</b></td><td>${on === 1 ? 'Rule 1 only' : 'Rules 1 to 6'}</td><td>₹${p.tol}${p.pct ? ` and ${p.pct}%` : ''}</td><td>${p.days ? `±${p.days} days` : 'none'}</td><td>${p.compare === 'tax' ? 'Tax only' : 'Taxable value and tax'}</td></tr>`;
};

/** the sections of the manual: id, title (shown in the side menu) */
export const HELP_SECTIONS: [string, string][] = [
  ['h-start', 'Start here'],
  ['h-steps', 'The steps every tab shares'],
  ['h-itc', 'Input Reco'],
  ['h-flow', 'GST ledger flow'],
  ['h-rev', 'Revenue reco'],
  ['h-annual', 'Annual return (GSTR-9 / 9C)'],
  ['h-settings', 'Matching settings explained'],
  ['h-wp', 'Saving and reusing a workpaper'],
  ['h-files', 'Getting the files'],
  ['h-faq', 'Questions and problems'],
];

export function helpHtml(): string {
  const reasons = REASONS.filter(r => r.code !== 'ROUND')
    .map(
      r =>
        `<tr><td>${esc(r.label)}</td><td>${r.derived ? 'Worked out by the tool' : 'You pick it'}</td><td>${esc(r.action)}</td></tr>`,
    )
    .join('');
  return `
<div class="hp-wrap">
  <nav class="hp-nav" aria-label="Help sections">
    <b>Help &amp; guide</b>
    ${HELP_SECTIONS.map(([id, t]) => `<a href="#help/${id}" data-help="${id}">${esc(t)}</a>`).join('')}
    <button class="hp-close" id="helpClose" type="button">← Back to the tool</button>
  </nav>
  <article class="hp-body">

  <section id="h-start">
    <h2>Start here</h2>
    <p>This tool reconciles a client's GST returns with their books and builds an Excel <b>workpaper</b> you can review, sign off and keep. Everything runs in your browser: <b>the files never leave your computer</b>. File names do not matter, because every file is recognised by what is inside it.</p>
    <p><b>Which tab do I use?</b></p>
    <table>
      <tr><th>Tab</th><th>Question it answers</th><th>You need</th></tr>
      <tr><td><a href="#help/h-itc" data-help="h-itc">Input Reco</a></td><td>Is the ITC in my purchase books the same as the ITC the suppliers filed (GSTR-2B)?</td><td>GSTR-2B JSON, purchase register</td></tr>
      <tr><td><a href="#help/h-flow" data-help="h-flow">GST ledger flow</a></td><td>Do the credit, cash and liability ledgers on the portal agree with the GST ledgers in my books, and why not?</td><td>Three portal ledgers, GST ledgers or a GL dump from the books</td></tr>
      <tr><td><a href="#help/h-rev" data-help="h-rev">Revenue reco</a></td><td>Is the turnover in my books the same as in GSTR-1 and GSTR-3B?</td><td>GSTR-1, GSTR-3B, sales register (or day book / GL dump)</td></tr>
      <tr><td><a href="#help/h-annual" data-help="h-annual">Annual return</a></td><td>Prepare GSTR-9 and 9C, and the upload JSON</td><td>Best: the three workpapers above, plus the portal GSTR-9 JSON</td></tr>
    </table>
    <h3>The usual order for a year-end</h3>
    <ol>
      <li><b>Input Reco</b> for the year (or month by month during the year). Save the workpaper.</li>
      <li><b>GST ledger flow</b> for the year. Save the workpaper.</li>
      <li><b>Revenue reco</b> for the year. Save the workpaper.</li>
      <li><b>Annual return</b>: drop the three saved workpapers. The tool brings back their files, settings and choices, builds the GSTR-9 / 9C workpaper, and compares each workpaper's figures with the return.</li>
      <li>Open the GSTR-9 workpaper in Excel, review it, save it, drop it back into the last step of the Annual return tab, and download the GSTR-9 upload JSON.</li>
    </ol>
    <p class="hp-note">You can use any tab on its own. The tabs share one file pool: files you add in one tab are available in the others, and files of a GSTIN other than the client's are warned about and ignored.</p>
  </section>

  <section id="h-steps">
    <h2>The steps every tab shares</h2>
    <ol class="hp-steps">
      <li><b>Add files.</b> Drop everything for the client at once, or a few at a time. The coverage grid shows which months of each return are loaded; a cross (✕) is a month missing inside the range you loaded. If the same return is dropped twice, the duplicate is left out with a message.</li>
      <li><b>Map register columns</b> (when a register is loaded). The tool guesses which column is the invoice number, date, GSTIN, taxable value and each tax, and shows a preview. Correct any wrong guess. Tally columnar registers (one column per ledger) are understood. Credit notes, debit notes and returns are told apart through the <i>Document type</i> column, so that they reduce rather than add.</li>
      <li><b>Confirm client &amp; year.</b> The GSTIN and legal name are read from the files; change them if the tool picked the wrong one. Choose the financial year.</li>
      <li><b>Matching settings.</b> How strictly invoices are matched, and which rules may explain a difference. The standard settings suit most clients. <a href="#help/h-settings" data-help="h-settings">What each setting does</a>.</li>
      <li><b>Review, confirm and correct</b> (differs by tab, see below). Anything you change is logged and carried into the workbook.</li>
      <li><b>Build workbook.</b> Downloads the Excel workpaper. Open it, review it, sign it off on the Index sheet, and <b>keep it</b>: it is also how you resume next time.</li>
    </ol>
    <p>Step numbers on screen differ by tab, because each tab shows only the steps it needs.</p>
  </section>

  <section id="h-itc">
    <h2>Input Reco: GSTR-2B against the purchase register</h2>
    <p><b>Purpose:</b> find out which purchase bills can be claimed as ITC, which are in the books but not in the supplier's GSTR-2B (do not claim yet), and which are in 2B but missing from the books.</p>
    <p><b>Files:</b> GSTR-2B JSON for each month you are reconciling, and the purchase register (or several: Tally register, journals with GST, payment vouchers, Zoho bills and expenses). Include credit notes and debit notes.</p>
    <h3>How the matching works</h3>
    <p>Each purchase document is paired with a GSTR-2B document. The tool tries the rules in this order, and a document takes the first rule that pairs it:</p>
    <ol>
      <li><b>Manual matches</b> you made come first. Documents you "keep apart" are never paired.</li>
      <li><b>Rule 1, same invoice no.</b> Same supplier GSTIN and same invoice number, ignoring punctuation, spaces and case.</li>
      <li><b>Rule 2, same digits.</b> The digits of the number agree (INV-0045 = 45) and the amounts are within tolerance.</li>
      <li><b>Rule 3, last digits.</b> One number ends with the other, or the last three or more digits agree, and the amounts are within tolerance.</li>
      <li><b>Rule 4, GSTIN + date + amount.</b> Same GSTIN, dates within the window, same taxable value and tax. The number is ignored.</li>
      <li><b>Rule 5, number + amount, GSTIN differs.</b> Same invoice number and amount, but the supplier GSTIN differs (usually a keying error in the books).</li>
      <li><b>Rule 6, supplier totals.</b> What is left of one supplier agrees in total on both sides (for example one 2B document booked as three bills). Purchases only.</li>
    </ol>
    <p>Books lines without a supplier GSTIN take it from GSTR-2B by supplier name: exact names are applied automatically, similar names are suggested under <b>Confirm supplier names</b> and applied only when you confirm them.</p>
    <p>Amendments in GSTR-2B (B2BA, CDNRA) carry the full revised values. The tool sets the original aside and counts only the latest version, so an amended invoice is never counted twice. Credit notes, debit notes and returns reduce ITC and are never paired with an invoice of the same number.</p>
    <h3>What you see in Review matches</h3>
    <table>
      <tr><th>Status</th><th>Meaning</th><th>What to do</th></tr>
      <tr><td>Matched</td><td>Same invoice, amounts agree</td><td>Nothing</td></tr>
      <tr><td>Matched - value difference</td><td>Paired, but taxable value or tax differs by more than the tolerance</td><td>Check the bill; claim only what 2B shows</td></tr>
      <tr><td>Matched - invoice no. differs / supplier GSTIN differs / supplier total</td><td>Paired by a looser rule</td><td>Confirm it is the same bill, or unmatch it</td></tr>
      <tr><td>Matched - in next FY 2B</td><td>The supplier reported it in the next financial year</td><td>Look at it before claiming</td></tr>
      <tr><td>Matched - manual</td><td>You paired it (code M001, M002…)</td><td>Nothing</td></tr>
      <tr><td>In books, not in 2B</td><td>The supplier has not reported it</td><td>Do not claim until it appears; follow up with the supplier</td></tr>
      <tr><td>In 2B, not in books</td><td>Reported by the supplier, not booked</td><td>Book it, or ask the supplier if it is not yours</td></tr>
    </table>
    <p>You can filter by status, party or period with totals, select documents and <b>match</b> them by hand (one to many is fine) or <b>keep apart</b> (code X). Every change is listed under "Your changes" with an undo, and becomes a Match code in the workbook.</p>
    <h3>The workbook</h3>
    <p>Summary (the ITC position and <b>claimable ITC</b> per tax head: for matched documents, the lower of books and 2B), Recon_Purchases (one row per document, with your match codes and remarks), Recon_Suppliers, Month_wise, Books_Purchase and RAW_2B. Every summary figure is a visible SUMIFS over the detail sheet, so it can be followed without the tool.</p>
  </section>

  <section id="h-flow">
    <h2>GST ledger flow: portal ledgers against books</h2>
    <p><b>Purpose:</b> show the whole movement of GST for the year (opening, output tax, RCM, credit notes, ITC availed, reversed and reclaimed, set-off, cash deposits and use, interest, late fee, penalty, refunds, closing) once as per books and once as per the portal, and explain every difference from the entries themselves.</p>
    <p><b>Files:</b> the portal's Electronic Credit Ledger, Electronic Cash Ledger and Electronic Liability Register (CSV or Excel), and the GST ledgers from the books (Tally voucher dump, Zoho detailed GL, or one export per ledger). GSTR-3B PDFs and GSTR-2B JSON are optional and add cross-checks and the ITC comparison.</p>
    <h3>Steps</h3>
    <ol>
      <li><b>Portal ledgers.</b> The tool reads them and shows the period each covers. Portal entries are counted by <b>return period</b> (the period in the ARN, or the return the entry belongs to).</li>
      <li><b>Books: GST accounts.</b> GST-looking ledgers are picked by name (rate-wise ledgers such as Output CGST 9% are added up under their head). Confirm the type and tax head of each. Value ledgers such as "IGST Sales 12%" are not GST accounts. Orders, delivery and receipt notes and memorandum vouchers are left out. Books entries count by voucher date unless you move a line to another return period.</li>
      <li><b>Opening and closing balances.</b> Single-ledger exports and the portal's Opening Balance rows carry balances. A GL dump does not, so enter the openings (and closings, to check against the trial balance). They are recorded on the Audit sheet.</li>
      <li><b>GST flow.</b> Three statements: <i>Flow as per books</i>, <i>Flow as per portal</i> and the <i>Difference</i>, with the unexplained part of each row.</li>
      <li><b>Reconciliation: books to portal.</b> Five parts, the same on screen and in Excel:
        <ol type="a">
          <li><b>Set-up checks:</b> a missing portal ledger, a books account with no opening balance where the portal has one, lines not placed in any flow row, coverage gaps, no GSTR-2B loaded. Fix these first.</li>
          <li><b>Year comparison:</b> one line per movement, books against portal for the year. The year-end cut-off is separated, and the rest is "to look at". Open a line to see the month-by-month difference and the entries behind it.</li>
          <li><b>ITC three ways:</b> books invoices (built from the purchase vouchers) are matched to GSTR-2B with the Input Reco rules, which bridges books to 2B. 2B to 3B is compared month by month against the credit ledger.</li>
          <li><b>Suggested journals:</b> for set-off, cash payments and fees, deposits and refunds, dated at the period end. They are suggestions only; nothing is posted anywhere.</li>
          <li><b>Closing statement:</b> closing per books + journals = adjusted, less each remaining difference, = closing per portal, with a nil check.</li>
        </ol>
      </li>
      <li><b>Cross-checks:</b> cash ledger against liability paid in cash, credit ledger against liability paid by ITC, credit ledger against GSTR-3B 4C, and liability register against 3B 6.1 (when 3B PDFs are loaded).</li>
    </ol>
    <h3>How a difference gets its reason</h3>
    <p>Reasons are <b>worked out from the entries, never typed</b>. Each rule takes the part of a difference it can explain, in this order; what no rule explains stays <b>Unexplained</b>, with its entries listed.</p>
    <ol>
      <li><b>Opening balance:</b> always shown first.</li>
      <li><b>Netted on portal:</b> ITC reversals, RCM, reclaims and credit notes that the portal nets into another row.</li>
      <li><b>Tax head:</b> IGST against CGST + SGST, cross-utilisation of ITC, accounts not split by head.</li>
      <li><b>Timing:</b> the same entry (or offsetting month totals) falls in different months. How many months apart is a setting.</li>
      <li><b>After the period:</b> an entry on the other side falls after the period (booked in March, return in April).</li>
      <li><b>Fee booked to tax:</b> interest or late fee paid in cash but debited to the tax ledger in the books.</li>
    </ol>
    <p>You can change only the <i>records</i> behind a reason: the flow row of a line (within the rows allowed for that account type), its tax head or its return period, always with a remark. Changes go to the Audit sheet and are kept per GSTIN and year. Lines that no rule can place go to "Not mapped" and you pick a row.</p>
    <h3>The workbook</h3>
    <p>Index, Instructions, Summary, Setup_Checks, Year_Compare, ITC_3way, ITC_Invoices, Journals, Closing, Flow_Books, Flow_Portal, Flow_Difference, Cross_Checks, Month_wise, Books_Lines, Portal_Lines, Balances and Audit. The comparison sheets are SUMIFS over Books_Lines, Portal_Lines and Balances.</p>
  </section>

  <section id="h-rev">
    <h2>Revenue reco: books, GSTR-1 and GSTR-3B</h2>
    <p><b>Purpose:</b> show that turnover and output tax in the books agree with what was reported in GSTR-1 and paid through GSTR-3B for the year, and give the reason for every difference.</p>
    <p><b>Files:</b> GSTR-1 JSON or ZIP for every month of the year (and April to October of the next year, to catch invoices reported late), GSTR-3B PDFs, and the books: a <b>sales register</b> (preferred, with credit notes) or, if there is none, a day book / GL dump / detailed ledgers. Optionally a customer list (name and GSTIN).</p>
    <h3>Books side</h3>
    <ul>
      <li>If a sales register is loaded it is used, and a GL dump is then only a control total.</li>
      <li>Without a register, invoices are built from the GL dump. The tool lists ledgers that look like revenue; tick them and say what each holds: taxable sales, exports / SEZ, exempt or nil-rated, non-GST, or <i>not a supply</i> (interest, dividend, forex gains). Tax comes from the Output GST lines of the same voucher. You can ignore a whole ledger, or single entries.</li>
    </ul>
    <h3>How books are matched to GSTR-1</h3>
    <ul>
      <li>Invoices are matched with the same rules as Input Reco (rules 1 to 5; the supplier-totals rule does not apply). Books use the <b>invoice date</b>, GSTR-1 uses the <b>return period</b>.</li>
      <li><b>Timing inside the same financial year is not a reconciling item</b>: an invoice dated in March and reported in April of the same FY simply matches. It still shows in the month views.</li>
      <li>B2CS (small) sales, Table 8 (nil / exempt) and e-commerce 9(5) supplies have no invoices, so they are compared by <b>month</b> (B2CS by month, rate and place of supply). Each books entry has a <b>B2CS month</b>: it starts as the invoice month. Change or clear it on screen, or in the yellow column of the workbook, to say which GSTR-1 month the entry was reported in.</li>
    </ul>
    <h3>Customer GSTIN</h3>
    <p>GSTR-1 has the customer GSTIN but no name, so a books customer without a GSTIN gets one by these methods, in order (each can be switched off in Matching settings). Your own choice under <b>Confirm customer names</b> always wins.</p>
    <ol>
      <li><b>Customer list</b> you loaded (a name with one GSTIN).</li>
      <li><b>Same name in the books</b> on other entries that carry one GSTIN (spacing, case and "Pvt Ltd" ignored).</li>
      <li><b>Invoice no. and amount:</b> mapped only when the customer's invoices agree with GSTR-1 documents of exactly one GSTIN by number <i>and</i> amount. A number match alone never maps. The settings choose whether this is applied automatically or only suggested, and how many invoices must agree.</li>
    </ol>
    <p>Every mapping shows its basis on screen and on the Customer_Map sheet. Similar names are only suggested.</p>
    <h3>The statements</h3>
    <p><b>Dashboard:</b> tiles, a month chart (books, GSTR-1, GSTR-3B) and a month table with running differences, for taxable value or tax.</p>
    <p><b>Books → GSTR-1:</b> the books total, then one line per reason, ending at the GSTR-1 total, like a bank reconciliation. "Not explained" is always the last line. <b>GSTR-1 → GSTR-3B:</b> RCM supplies, advances and months without a 3B, then each 3B table.</p>
    <table>
      <tr><th>Reason</th><th>Set by</th><th>What to do</th></tr>
      ${reasons}
    </table>
    <p>In <b>Entries and reasons</b> you can filter (customer, type, month, invoice date, reason, search, differences only) and pick a reason and a remark for one entry or in bulk. In the workbook, the yellow columns <i>Your reason</i>, <i>B2CS month</i> and <i>Remark</i> have drop-downs and are read back when the workbook is dropped again.</p>
  </section>

  <section id="h-annual">
    <h2>Annual return: GSTR-9 and 9C</h2>
    <p><b>Purpose:</b> fill the GSTR-9 / 9C workpaper from the returns and the books, check it against the portal's auto-drafted GSTR-9, and produce the GSTR-9 upload JSON.</p>
    <h3>Recommended route: build from your three workpapers</h3>
    <ol>
      <li>Finish and save the <b>Input Reco</b>, <b>Revenue reco</b> and <b>GST ledger flow</b> workpapers first.</li>
      <li>On the Annual return tab, drop the three workpapers, the portal GSTR-9 JSON, and next year's GSTR-1 / 2B for April to October if they are not already in the workpapers. The tab shows which of the three are loaded.</li>
      <li>Each workpaper brings back its source files and settings. The Input Reco match codes and remarks carry over into the annual purchase reconciliation, so you do not redo them.</li>
      <li>Build the workbook. It has a <b>Workpapers</b> sheet comparing each workpaper's headline figures with the matching GSTR-9 cells, with the difference and OK or Look at.</li>
    </ol>
    <p>Building direct from GSTR-1, 2B, 3B and the registers also works, but is <b>not recommended</b>: it skips the cross-checks the other tabs give.</p>
    <h3>Where the figures come from</h3>
    <ul>
      <li>Outward supplies (Tables 4 and 5) come from GSTR-1, with amendments. Tables 10 and 11 (amendments for the year) come from the next year's GSTR-1, April to October.</li>
      <li>ITC (Tables 6, 7 and 8) comes from GSTR-3B and from GSTR-2B against the books: the current year's 2B, and next year's 2B for April to October only, as ICAI's Technical Guide on GSTR-9 says. Next-year credit notes are tagged for review.</li>
      <li>Tax paid (Table 9) comes from GSTR-3B 6.1, and the 9C reconciliation amounts from the books.</li>
      <li>The portal's auto-drafted GSTR-9 is loaded for comparison.</li>
      <li>A <b>Summary</b> sheet lists the figures to file (4N, 5N, 6O, 8A, Table 9, the 9C un-reconciled amounts) and the open checks.</li>
    </ul>
    <h3>Last step: generate the upload JSON</h3>
    <ol>
      <li>Open the workbook in Excel and review it. Enter your changes in the "To be filed" columns.</li>
      <li>Save it, then drop the saved file into the "Generate upload JSON" step.</li>
      <li>The tool runs an <b>edit check</b>: formulas against the template, a fingerprint and seal of the imported data, and a list of the adjustments you made. Then it produces the JSON.</li>
      <li>On the GST portal: GSTR-9 tile → Prepare Offline → Upload tab → choose the JSON → Upload. Wait for <b>Processed</b>. If it shows "Processed with Error", download the error report.</li>
    </ol>
  </section>

  <section id="h-settings">
    <h2>Matching settings explained</h2>
    <p>This step comes before any matching, so a result is never produced under settings you have not seen. Settings are saved <b>per tab</b> in this browser and written to the workbook's Read Me, so a workpaper says what rules it was made with.</p>
    <h3>Invoice rules (Input Reco, Annual return, Revenue reco, and ITC three ways in the flow tab)</h3>
    <table>
      <tr><th>Preset</th><th>Rules</th><th>Tolerance</th><th>Date window</th><th>Amounts compared</th></tr>
      ${presetRow('Strict')}${presetRow('Standard')}${presetRow('Lenient')}
    </table>
    <ul>
      <li><b>Tolerance ₹:</b> amounts within this many rupees are equal.</li>
      <li><b>Tolerance %</b> (rules 2 and 3): extra tolerance, as a percentage of the amount, when the number is only partly the same.</li>
      <li><b>Date window</b> (rule 4): dates may differ by this many days.</li>
      <li><b>Amounts compared:</b> taxable value and tax, or tax only.</li>
      <li><b>Supplier / customer name similarity:</b> how alike two names must be before they are suggested. Lower finds more suggestions and more wrong ones. Suggestions are never applied without you.</li>
    </ul>
    <p>Start with Standard. Use Strict when only exact numbers should pair, and Lenient for messy books. Either way, look at the rows matched by a looser rule.</p>
    <h3>GST ledger flow</h3>
    <p>The equal-within ₹ tolerance, which explaining rules run (netted on portal, tax head, timing, after the period, fee booked to tax) and the two timing windows. Switching a rule off makes those differences show as unexplained instead. Opening-balance differences are always shown first.</p>
    <h3>Revenue reco</h3>
    <p>The invoice rules, and the customer GSTIN methods described above.</p>
    <h3>Fixed, not a setting</h3>
    <p>Next-year returns count for April to October only; books by invoice date against GSTR-1 by return period; portal entries by return period and books by voucher date; and the order in which rules run.</p>
  </section>

  <section id="h-wp">
    <h2>Saving and reusing a workpaper</h2>
    <p>The Excel file the tool writes is both your <b>workpaper</b> and your <b>save file</b>. It carries two hidden sheets, <code>_JJState</code> and <code>_JJFiles</code>, with the settings and choices and a copy of the source files. So the next time you can drop <b>only the workpaper</b>, with no need to find the original downloads.</p>
    <h3>Save</h3>
    <ol>
      <li>Build the workbook and download it. The file name does not matter.</li>
      <li>Review it in Excel and save it. Do <b>not</b> delete, rename or edit the hidden <code>_JJState</code> and <code>_JJFiles</code> sheets.</li>
      <li>Keep it with the client's file. It holds the client's data, so treat it like the returns themselves.</li>
    </ol>
    <h3>Come back to it</h3>
    <ol>
      <li>Open the tool and drop the workpaper. The page switches to the workpaper's own tab.</li>
      <li>The Matching settings step says "Restored from …" and lists what came back:
        <ul>
          <li><b>Settings</b> of the tab, as they were when it was built.</li>
          <li><b>Choices</b>: register column mappings and defaults, supplier names, revenue ledger types, customer matches, reasons, B2CS months, flow account types, moved lines, balances, the month range, and Input Reco match codes and remarks.</li>
          <li><b>Source files</b>: every return, portal ledger, register and GL dump used. They show in the file list marked "from the workpaper".</li>
        </ul>
      </li>
      <li>Change anything as usual and build again.</li>
      <li>To go back to the settings you had before the restore, use <b>Keep my earlier settings instead</b>.</li>
    </ol>
    <h3>What can be replaced, and what cannot</h3>
    <table>
      <tr><th>Never changes (filed documents)</th><th>You may replace by dropping a newer file</th></tr>
      <tr><td>GSTR-1, GSTR-2B, GSTR-3B, the portal GSTR-9, and the portal credit, cash and liability ledgers</td><td>Sales register, purchase register, GL dump and customer list: the new file replaces the workpaper's copy of the same thing</td></tr>
    </table>
    <p>If a return is missing from the workpaper (a late month, say), drop it as well: it is added.</p>
    <h3>Things to know</h3>
    <ul>
      <li><b>Hand corrections in Excel:</b> on Input Reco and Annual return, the restored register rebuilds the Books sheets, so lines you corrected by hand there are replaced. To keep them, remove the register from the file list before you build.</li>
      <li><b>Revenue reco:</b> the yellow columns (<i>Your reason</i>, <i>B2CS month</i>, <i>Remark</i>) are read back after the saved state, so your Excel edits win.</li>
      <li><b>Annual return:</b> dropping the three workpapers there builds from them, and the page stays on the Annual return tab.</li>
      <li><b>Other client:</b> files of a different GSTIN are warned about and ignored. Use a fresh page for a different client.</li>
    </ul>
    <h3>Reading the workpaper without the tool</h3>
    <p>Every workbook opens on an <b>Index</b> (cover, contents with links, how to read it, colour key, Prepared / Reviewed sign-off). There is one <b>Instructions</b> sheet and a <b>Summary</b>. Each summary figure is a plain SUMIFS or link over a visible detail sheet, so a reviewer who never opens the tool can follow the working. The font is Aptos throughout.</p>
  </section>

  <section id="h-files">
    <h2>Getting the files</h2>
    <p>These are the same instructions as under <b>Add files → Where do I get these files?</b> and in each tab. Menu paths are as of November 2026; if a portal or software menu has moved, the report names still tell you what to look for.</p>
    <h3>GST portal (all tabs)</h3>
    <p>Log in at <b>gst.gov.in</b> → <b>Services → Returns → Returns Dashboard</b> → pick the financial year and period → <b>Search</b>.</p>
    <table>
      <tr><th>File</th><th>How</th><th>Used in</th></tr>
      <tr><td>GSTR-2B JSON</td><td>GSTR-2B tile → Download → Generate JSON file to download. One per month; for the annual return also April to October of the next year.</td><td>Input Reco, Flow, Annual</td></tr>
      <tr><td>GSTR-1 JSON / ZIP</td><td>GSTR-1 / IFF tile → Download → Generate JSON file to download (the link appears after a few minutes; refresh). Drop the ZIP as it is. Previous, current and next FY for the annual return.</td><td>Revenue, Annual</td></tr>
      <tr><td>GSTR-3B PDF</td><td>GSTR-3B tile of a filed period → Download. One per month (per quarter for QRMP filers).</td><td>Revenue, Flow (optional), Annual</td></tr>
      <tr><td>GSTR-9 auto-drafted JSON</td><td>Returns Dashboard → Annual Return → year → GSTR-9 tile → Prepare Offline → Download tab → Generate JSON file to download.</td><td>Annual</td></tr>
    </table>
    <h3>For Input Reco and the Annual return: purchases</h3>
    <p>Purchase register, journals and payment vouchers with GST, debit notes (Tally), and bills, expenses and manual journals (Zoho): the click-by-click steps are in <b>Add files → Where do I get these files?</b>, under the Tally Prime and Zoho Books tabs. Any other software or Excel works if the file has date, invoice number, supplier GSTIN, taxable value and tax columns.</p>
    <h3>For the GST ledger flow</h3>
    ${guideList(FLOW_GUIDE)}
    <h3>For the Revenue reco</h3>
    ${guideList(REV_GUIDE.filter(g => g.source !== 'portal'))}
    <h3>Good to know about registers</h3>
    <ul>
      <li>Any .xlsx or .csv works. Title rows above the header are fine; the tool finds the header row.</li>
      <li>Zoho exports one row per item line; lines are added up to invoice level.</li>
      <li>Include credit notes, debit notes and returns, and make sure the mapping step knows which rows they are.</li>
      <li>Load several registers together when ITC is booked outside the purchase register (journals, payment vouchers, expenses).</li>
    </ul>
  </section>

  <section id="h-faq">
    <h2>Questions and problems</h2>
    <dl>
      <dt>Is my data uploaded anywhere?</dt>
      <dd>No. The page runs in your browser and nothing is sent to a server.</dd>
      <dt>Which browser?</dt>
      <dd>Chrome or Edge. No install, and no Excel 365 needed.</dd>
      <dt>The tool says files of another GSTIN are ignored.</dt>
      <dd>It works with one client GSTIN at a time (the most frequent one in the files, unless you type one). Check the GSTIN under Confirm client &amp; year.</dd>
      <dt>A month shows ✕ in the coverage grid.</dt>
      <dd>That month's return is missing inside the range you loaded. Download it from the portal and drop it.</dd>
      <dt>Everything shows as "In books, not in 2B".</dt>
      <dd>Usually the register's GSTIN or invoice-number column is mapped wrongly, or the GSTR-2B belongs to a different month or GSTIN. Check the preview under Map register columns.</dd>
      <dt>The numbers look doubled or too high.</dt>
      <dd>Check the file list for a register loaded twice, and the Document type mapping for credit notes that are not marked as such.</dd>
      <dt>It is slow with large files.</dt>
      <dd>A progress bar shows while files are read and workbooks built. The tabs recalculate once after the last file, so add all the files before judging.</dd>
      <dt>Excel says the workbook is damaged.</dt>
      <dd>Workbooks saved by version 0.14 had this fault. Rebuild them with the current version.</dd>
      <dt>Can I change a formula in the workpaper?</dt>
      <dd>On the GSTR-9 workbook, the edit check reports every changed formula and adjustment when you generate the JSON. On the other workbooks, change the yellow input cells rather than formulas, so the working stays readable.</dd>
    </dl>
  </section>

  </article>
</div>`;
}

// ---------- the page: a view that replaces the tool's cards while it is open

let built = false;
const view = () => document.getElementById('helpview')!;

/** open the manual, at a section when given (an id from HELP_SECTIONS) */
export function openHelp(id?: string) {
  const v = view();
  if (!built) {
    v.innerHTML = helpHtml();
    built = true;
    v.querySelectorAll<HTMLElement>('[data-help]').forEach(a =>
      a.addEventListener('click', e => {
        e.preventDefault();
        openHelp(a.dataset.help);
      }),
    );
    document.getElementById('helpClose')!.addEventListener('click', () => closeHelp());
  }
  document.body.classList.add('helping');
  const sec = id && HELP_SECTIONS.some(s => s[0] === id) ? id : HELP_SECTIONS[0][0];
  v.querySelectorAll('.hp-nav a').forEach(a => a.classList.toggle('on', (a as HTMLElement).dataset.help === sec));
  history.replaceState(null, '', `#help/${sec}`);
  document.querySelectorAll('button.tab').forEach(b => b.classList.remove('on'));
  const t = document.getElementById(sec);
  window.scrollTo({ top: t ? Math.max(0, t.getBoundingClientRect().top + window.scrollY - 150) : 0 });
}

/** leave the manual (the caller shows the tab again) */
export function closeHelp() {
  if (!document.body.classList.contains('helping')) return;
  document.body.classList.remove('helping');
  history.replaceState(null, '', location.pathname + location.search);
  window.dispatchEvent(new Event('jj-help-closed'));
  window.scrollTo({ top: 0 });
}

/** wire the Help button and #help/<section> links */
export function initHelp() {
  document
    .getElementById('helpBtn')
    ?.addEventListener('click', () => (document.body.classList.contains('helping') ? closeHelp() : openHelp()));
  const fromHash = () => {
    const m = /^#help(?:\/([\w-]+))?$/.exec(location.hash);
    if (m) openHelp(m[1]);
  };
  window.addEventListener('hashchange', fromHash);
  fromHash();
}

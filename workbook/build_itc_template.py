"""Generates the GST Annual Return Workbook (GSTR-9 / GSTR-9C) - macro-free, Excel 2010+ compatible.

Layout conventions
- Every GSTR-9 data row carries a stable KEY (column AK) used by the loader (import) and JSON generator (export).
- Raw / books sheets are summed with SUMIFS on that key, so the workbook is fully traceable.
- Blocks per row: Portal | As per Returns | As per Books | Adjustment | To be Filed | Variance (Filed - Portal)
"""
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

import os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
LOADER = os.path.join(HERE, '..', 'loader')
sys.path.insert(0, HERE)
import jjfront
M = json.load(open(os.path.join(HERE, 'masters.json')))
# writes straight into the loader, so a rebuild needs no copying
OUT = os.path.join(LOADER, 'template', 'ITC_Recon_Template.xlsx')

# ---------------------------------------------------------------- styles
F = 'Aptos'
f_norm = Font(name=F, size=9)
f_bold = Font(name=F, size=9, bold=True)
f_title = Font(name=F, size=13, bold=True, color='2C388E')
f_sub = Font(name=F, size=9, italic=True, color='595959')
f_hdr = Font(name=F, size=9, bold=True, color='FFFFFF')
f_input = Font(name=F, size=9, color='0000FF')
f_link = Font(name=F, size=9, color='008000')
fill_hdr = PatternFill('solid', fgColor='2C388E')
fill_blk = {
    'Portal': PatternFill('solid', fgColor='DDEBF7'),
    'Returns': PatternFill('solid', fgColor='E2EFDA'),
    'Books': PatternFill('solid', fgColor='FCE4D6'),
    'Adjustment': PatternFill('solid', fgColor='FFF2CC'),
    'To be Filed': PatternFill('solid', fgColor='D9D2E9'),
    'Variance': PatternFill('solid', fgColor='EDEDED'),
}
fill_input = PatternFill('solid', fgColor='FFFF99')
fill_na = PatternFill('solid', fgColor='BFBFBF')
fill_sub = PatternFill('solid', fgColor='F2F2F2')
fill_sec = PatternFill('solid', fgColor='D6DCE4')
thin = Side(style='thin', color='BFBFBF')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
NUM = '#,##0.00;(#,##0.00);"-"'
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
wrap = Alignment(wrap_text=True, vertical='top')

wb = Workbook()
wb.remove(wb.active)


def cell(ws, ref, v=None, font=f_norm, fill=None, fmt=None, align=None, border=True):
    c = ws[ref]
    if v is not None:
        c.value = v
    c.font = font
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if align: c.alignment = align
    if border: c.border = box
    return c


def title(ws, text, sub=None):
    ws['A1'] = "='Basic'!$C$4"
    ws['A1'].font = f_bold
    ws['A2'] = text
    ws['A2'].font = f_title
    if sub:
        ws['A3'] = sub
        ws['A3'].font = f_sub
    ws.sheet_view.showGridLines = False



# ================================================================= ITC RECONCILIATION workbook (monthly mode of the Loader)
# ================================================================= RECONCILIATION sheets (filled by the Loader)
NREC = 20000
def recon_sheet(name, sub, headers, widths, num_cols, summary=None, start=16):
    ws = wb.create_sheet(name)
    title(ws, name.replace('_', ' '), sub)
    hr = start - 1
    for j, (h, w) in enumerate(zip(headers, widths)):
        cell(ws, f'{CL(j + 1)}{hr}', h, f_hdr, fill_hdr, align=center)
        ws.column_dimensions[CL(j + 1)].width = w
    ws.row_dimensions[hr].height = 30
    # style row: the Loader copies these cell styles onto every row it writes
    for j in range(len(headers)):
        c = ws[f'{CL(j + 1)}{start}']
        c.font = f_input if headers[j].startswith(('Remarks', 'Match code')) else f_norm
        c.border = box
        if CL(j + 1) in num_cols:
            c.number_format = NUM
        if headers[j].startswith(('Remarks', 'Match code')):
            c.fill = fill_input
    ws.freeze_panes = f'C{start}'
    ws.auto_filter.ref = f'A{hr}:{CL(len(headers))}{hr}'
    if summary:
        summary(ws)
    return ws

def status_summary(statuses, status_col, sum_cols, labels):
    def f(ws):
        # the summary must end above the table header (row 15): with 10 statuses it starts on row 3 and the sheet note moves beside the title
        top = 4 if len(statuses) <= 9 else 3
        if top == 3:
            ws['C2'] = ws['A3'].value; ws['C2'].font = f_sub; ws['A3'] = None
        cell(ws, f'A{top}', 'Status', f_hdr, fill_hdr, align=center)
        cell(ws, f'B{top}', 'Count', f_hdr, fill_hdr, align=center)
        for j, lab in enumerate(labels):
            cell(ws, f'{CL(3 + j)}{top}', lab, f_hdr, fill_hdr, align=center)
        for i, st in enumerate(statuses):
            r = top + 1 + i
            cell(ws, f'A{r}', st, f_bold if i == 0 else f_norm)
            cell(ws, f'B{r}', f'=COUNTIF(${status_col}$16:${status_col}${NREC},A{r})', fmt='#,##0')
            for j, sc in enumerate(sum_cols):
                cell(ws, f'{CL(3 + j)}{r}', f'=SUMIF(${status_col}$16:${status_col}${NREC},$A{r},${sc}$16:${sc}${NREC})', fmt=NUM)
        r = top + 1 + len(statuses)
        cell(ws, f'A{r}', 'Total', f_bold, fill_sub)
        cell(ws, f'B{r}', f'=SUM(B{top + 1}:B{r - 1})', f_bold, fill_sub, '#,##0')
        for j in range(len(sum_cols)):
            col = CL(3 + j)
            cell(ws, f'{col}{r}', f'=SUM({col}{top + 1}:{col}{r - 1})', f_bold, fill_sub, NUM)
        ws.column_dimensions['A'].width = max(ws.column_dimensions['A'].width or 0, 34)
    return f

HELPER = ['Books tax', '2B tax', 'Claimable CGST', 'Claimable SGST', 'Claimable IGST', 'Claimable cess', 'Claimable tax', 'Books month']
HELPER_NUM = {'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ'}
P_STATUS = ['Matched', 'Matched - value difference', 'Matched - invoice no. differs', 'Matched - supplier GSTIN differs',
            'Matched - in next FY 2B', 'Matched - supplier total', 'Matched - manual', 'In books, not in 2B', 'In 2B, not in books', 'Not compared (RCM / ISD / imports)']
recon_sheet('Recon_Purchases', 'Purchase register vs GSTR-2B, invoice by invoice. Manual match: give the rows that belong together the same Match code (e.g. M1), save, and drop the workbook into the Loader. Codes and remarks are kept on refresh.',
            ['Status', 'Match basis', 'Supplier GSTIN', 'Supplier name', 'Invoice no. (books)', 'Invoice no. (2B)', 'Date (books)', 'Date (2B)',
             'Books taxable', 'Books CGST', 'Books SGST', 'Books IGST', 'Books cess', '2B taxable', '2B CGST', '2B SGST', '2B IGST', '2B cess',
             'Diff taxable (books - 2B)', 'Diff tax (books - 2B)', '2B period', '2B section', '2B ITC available', 'Books ITC status', 'Books ITC claimed FY',
             'GSTR-9 effect', 'Match code', 'Remarks', 'ID'] + HELPER,
            [26, 20, 17, 26, 16, 16, 11, 11] + [12] * 12 + [9, 8, 8, 16, 10, 30, 10, 30, 22] + [12] * 7 + [10],
            set('IJKLMNOPQRST') | HELPER_NUM,
            status_summary(P_STATUS, 'A', ['I', 'N', 'S', 'T'], ['Books taxable', '2B taxable', 'Diff taxable', 'Diff tax']))
wb['Recon_Purchases']['G3'] = 'Tax = CGST + SGST + IGST + cess. Tolerance Rs 1.'
wb['Recon_Purchases']['G3'].font = f_sub
recon_sheet('Recon_Suppliers', 'Supplier-wise summary of the purchase reconciliation - use it for vendor follow-ups.',
            ['Supplier GSTIN', 'Supplier name', 'Invoices in books', 'Invoices in 2B', 'Books tax', '2B tax', 'Diff tax (books - 2B)',
             'In books, not in 2B', 'In 2B, not in books', 'Value differences', 'Totals match (±Rs 1)', 'Remarks'],
            [17, 34, 10, 10, 14, 14, 14, 11, 11, 11, 12, 34], set('EFG'), start=6)
S_STATUS = ['Matched', 'Matched - value difference', 'Matched - GSTIN differs', 'Matched - reported in next FY', 'Matched - manual', 'In books, not in GSTR-1', 'In GSTR-1, not in books']
recon_sheet('Recon_Sales', 'Sales register vs GSTR-1, document by document (B2B, exports, SEZ, deemed exports, notes, B2C large). Manual match: same Match code on the rows that belong together, then drop the workbook into the Loader.',
            ['Status', 'Match basis', 'Customer GSTIN (books)', 'Customer GSTIN (GSTR-1)', 'Customer', 'Doc type', 'Doc no. (books)', 'Doc no. (GSTR-1)',
             'Date (books)', 'Date (GSTR-1)', 'GSTR-1 section', 'GSTR-1 period', 'Books taxable', 'Books tax', 'GSTR-1 taxable', 'GSTR-1 tax',
             'Diff taxable (books - GSTR-1)', 'Diff tax (books - GSTR-1)', 'Supply type (books)', 'GSTR-9 effect', 'Match code', 'Remarks', 'ID'],
            [26, 18, 17, 17, 26, 8, 16, 16, 11, 11, 10, 9, 13, 12, 13, 12, 13, 12, 12, 30, 10, 30, 20],
            set('MNOPQR'),
            status_summary(S_STATUS, 'A', ['M', 'O', 'Q', 'R'], ['Books taxable', 'GSTR-1 taxable', 'Diff taxable', 'Diff tax']))
recon_sheet('Recon_Sales_Monthly', 'Month-wise: books (by document date) vs GSTR-1 (by return period), net of credit notes. Advances excluded.',
            ['Month', 'Books taxable', 'GSTR-1 taxable', 'Diff taxable', 'Books tax', 'GSTR-1 tax', 'Diff tax',
             'Books B2C taxable', 'GSTR-1 B2C taxable', 'Diff B2C', 'Books exempt / nil / non-GST', 'GSTR-1 exempt / nil / non-GST', 'Diff', 'Remarks'],
            [11, 15, 15, 14, 14, 14, 13, 15, 15, 13, 15, 15, 13, 30], set('BCDEFGHIJKLM'), start=6)

# ================================================================= MAP sheet

for n in ('Recon_Sales', 'Recon_Sales_Monthly'):
    wb.remove(wb[n])

def plain_sheet(name, sub, headers, widths, num_cols, input_style=False):
    ws = wb.create_sheet(name)
    title(ws, name.replace('_', ' '), sub)
    for j, (h, w) in enumerate(zip(headers, widths)):
        cell(ws, f'{CL(j + 1)}5', h, f_hdr, fill_hdr, align=center)
        ws.column_dimensions[CL(j + 1)].width = w
        c = ws[f'{CL(j + 1)}6']
        c.font = f_input if input_style else f_norm
        c.border = box
        if CL(j + 1) in num_cols: c.number_format = NUM
    ws['A6'].number_format = 'DD-MM-YYYY'
    ws.row_dimensions[5].height = 30
    ws.freeze_panes = 'A6'
    ws.auto_filter.ref = f'A5:{CL(len(headers))}5'
    return ws
plain_sheet('Books_Purchase', 'Purchase register as mapped by the Loader (one row per register line).',
            ['Doc date', 'Doc number', 'Supplier', 'Supplier GSTIN', 'Supply Type', 'ITC Type', 'ITC Status', 'ITC claimed in FY', 'HSN', 'UQC',
             'Quantity', 'Rate', 'Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', 'Remarks'],
            [11, 14, 26, 17, 12, 13, 14, 10, 10, 7, 10, 6, 14, 12, 12, 12, 10, 22], set('MNOPQ'), True)
plain_sheet('RAW_2B', 'GSTR-2B documents read by the Loader.',
            ['Key', 'Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', 'Source', '2B FY', '2B period', 'Section', 'Supplier GSTIN', 'Supplier name',
             'Doc type', 'Doc number', 'Doc date', 'Doc FY', 'ITC available', 'RCM', 'IMS'],
            [8, 14, 13, 13, 13, 11, 8, 9, 10, 9, 17, 28, 8, 16, 11, 9, 9, 6, 10], set('BCDEF'))

plain_sheet('Month_wise', 'Month by month (tax = CGST + SGST + IGST + cess). 2B figures by 2B period; books figures by invoice date.',
            ['Month', '2B ITC available', 'Claimable (matched, by 2B period)', 'In 2B, not in books', 'Books ITC (by invoice date)', 'In books, not in 2B (by invoice date)', 'Remarks'],
            [12, 16, 18, 16, 18, 18, 34], set('BCDEF'))
wb['Month_wise']['G6'].font = f_input; wb['Month_wise']['G6'].fill = fill_input
# ---- Summary (first sheet)
ws = wb.create_sheet('Summary', 0)
ws.sheet_view.showGridLines = False
ws['B1'] = 'JAMES & JAMES LLP  ·  CHARTERED ACCOUNTANTS'; ws['B1'].font = Font(name=F, size=10, bold=True, color='2C388E')
ws['B2'] = 'ITC Reconciliation - GSTR-2B vs purchase register'; ws['B2'].font = f_title
ws.column_dimensions['A'].width = 3; ws.column_dimensions['B'].width = 52
for c in 'CDEFG': ws.column_dimensions[c].width = 15
for i, lab in enumerate(['Legal name', 'GSTIN', 'GSTR-2B periods', 'Financial year', 'Prepared by', 'Built on']):
    cell(ws, f'B{4 + i}', lab, f_bold, fill_sub); cell(ws, f'C{4 + i}', 'James & James LLP, Chartered Accountants' if lab == 'Prepared by' else None, f_norm)
    ws.merge_cells(f'C{4 + i}:G{4 + i}')
r0 = 11
cell(ws, f'B{r0}', 'ITC position (tax)', f_hdr, fill_hdr)
for j, h in enumerate(['Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', 'Total']):
    cell(ws, f'{CL(3 + j)}{r0}', h, f_hdr, fill_hdr, align=center)
lines = ['ITC available as per GSTR-2B (forward charge, ITC available = Y)',
         'ITC as per books (invoices compared with 2B)',
         'ITC claimable - matched invoices (lower of books and 2B)',
         'In books, not in 2B - do not claim; follow up with supplier',
         'In 2B, not in books - book the invoice or treat as not availed',
         'Value differences on matched invoices (books - 2B)']
HOW = ['SUMIFS of the 2B tax columns (O-R) of Recon_Purchases: rows with a 2B invoice whose "2B ITC available" is not N.',
       'SUMIFS of the books tax columns (J-M) of Recon_Purchases: rows with a books invoice, other than "Not compared".',
       'SUM of the "Claimable" columns (AF-AI) of Recon_Purchases: on each matched invoice, the lower of books and 2B (MIN).',
       'SUMIFS of the books tax columns where Status = "In books, not in 2B".',
       'SUMIFS of the 2B tax columns where Status = "In 2B, not in books".',
       'Books minus 2B tax where Status = "Matched - value difference".']
cell(ws, f'H{r0}', 'How it is worked out', f_hdr, fill_hdr)
ws.column_dimensions['H'].width = 80
for i, lab in enumerate(lines):
    r = r0 + 1 + i
    cell(ws, f'H{r}', HOW[i], Font(name=F, size=9, color='4A5068'), align=wrap)
    cell(ws, f'B{r}', lab, f_bold if i == 2 else f_norm, fill_blk['To be Filed'] if i == 2 else None)
    for j in range(4): cell(ws, f'{CL(3 + j)}{r}', 0, f_bold if i == 2 else f_norm, fill_blk['To be Filed'] if i == 2 else None, NUM)
    cell(ws, f'G{r}', f'=SUM(C{r}:F{r})', f_bold, fill_sub, NUM)
r1 = r0 + len(lines) + 3
cell(ws, f'B{r1}', 'Invoice status', f_hdr, fill_hdr)
cell(ws, f'C{r1}', 'Count', f_hdr, fill_hdr, align=center)
cell(ws, f'D{r1}', 'Diff tax (books - 2B)', f_hdr, fill_hdr, align=center)
for i in range(len(P_STATUS)):
    r = r1 + 1 + i
    cell(ws, f'B{r}', f'=Recon_Purchases!A{4 + i}', f_link)
    cell(ws, f'C{r}', f'=Recon_Purchases!B{4 + i}', f_link, fmt='#,##0')
    cell(ws, f'D{r}', f'=Recon_Purchases!F{4 + i}', f_link, fmt=NUM)
r2 = r1 + len(P_STATUS) + 2
notes = ['How to use: drop the 2B JSON(s) and the purchase register into the Loader (Monthly ITC reconciliation), map the columns, build.',
         'Manual match: give rows that belong together the same Match code in Recon_Purchases, save, and drop this workbook into the Loader with the same files.',
         'Recon_Suppliers lists suppliers with differences - use it for vendor follow-up. Month_wise splits the position by month when several 2B months are loaded.',
         'ITC claimable is the lower of books and 2B on each matched invoice (section 16(2)(aa)); check eligibility (17(5)) and payment (rule 37) separately.']
ws[f'B{r2}'] = 'How to use this workpaper, and how each figure is worked out: see the Instructions sheet.'; ws[f'B{r2}'].font = f_sub
for w in wb.worksheets:
    if w.title != 'Summary':
        w['A1'] = "=Summary!$C$4&\"  |  \"&Summary!$C$5&\"  |  2B \"&Summary!$C$6&\"  |  Prepared by James & James LLP\""
        w['A1'].font = f_bold
    w.page_setup.orientation = 'landscape'; w.page_setup.fitToWidth = 1; w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr.fitToPage = True
    w.oddFooter.left.text = 'Prepared by James & James LLP, Chartered Accountants'; w.oddFooter.left.size = 8
    w.oddFooter.center.text = 'Page &P of &N'; w.oddFooter.center.size = 8; w.oddFooter.right.text = '&A'; w.oddFooter.right.size = 8
wb['Summary'].sheet_properties.tabColor = '2C388E'
for n in ('Recon_Purchases', 'Recon_Suppliers'): wb[n].sheet_properties.tabColor = 'FFC000'
wb['Books_Purchase'].sheet_properties.tabColor = 'ED7D31'; wb['Month_wise'].sheet_properties.tabColor = 'FFC000'; wb['RAW_2B'].sheet_properties.tabColor = '70AD47'
# ---- Index, Instructions (one sheet), Aptos everywhere, a link back to the Index on every sheet
jjfront.add_instructions(wb, 'ITC reconciliation: GSTR-2B vs purchase register', [
    ('What this workpaper does', [
        ('Purpose', 'Compares the purchase register with GSTR-2B invoice by invoice and works out the ITC that can be claimed: the lower of books and 2B on each matched invoice whose 2B says ITC is available.'),
        ('Start here', 'Summary (ITC position), then Recon_Purchases (every invoice with its status) and Recon_Suppliers (for vendor follow-up).')]),
    ('How the figures are worked out', [
        ('Recon_Purchases', 'One row per invoice: books side (columns I-M), 2B side (N-R), differences (S-T), status (A) and how it matched (B). Columns AD-AK work out books tax, 2B tax, the claimable ITC per head (MIN of books and 2B on matched rows) and the books month, so every summary is one SUMIFS.'),
        ('Status summary', 'Rows 3-14 of Recon_Purchases: COUNTIF / SUMIF of the rows below by Status.'),
        ('Summary', 'Each ITC position line is a SUMIFS (or SUM) over Recon_Purchases; column H says which.'),
        ('Recon_Suppliers', 'COUNTIFS / SUMIFS of Recon_Purchases by supplier GSTIN.'),
        ('Month_wise', 'SUMIFS of Recon_Purchases by 2B period (column U) for the 2B side and by books month (column AK) for the books side.'),
        ('Books_Purchase / RAW_2B', 'The purchase register as mapped in the Loader, and the GSTR-2B documents read from the JSON files.')]),
    ('Working with the tool', [
        ('Match by hand', 'Give rows that belong together the same Match code (e.g. M1) in Recon_Purchases; X keeps two rows apart. Save and drop the workbook into the Loader with the same files: codes and remarks are kept.'),
        ('Next time', 'Drop this saved workpaper alone into the Loader (Input Reco tab): it brings back its 2B files, register, mapping, settings and supplier names. A newer purchase register dropped with it replaces the saved copy; 2B stays as filed.'),
        ('Annual return', 'This workpaper, with the Revenue reco and GST ledger flow workpapers, is what the Annual return tab builds GSTR-9 / 9C from: its match codes carry over to the annual purchase reconciliation.'),
        ('Eligibility', 'ITC claimable here is the 2B test (section 16(2)(aa)) only; check eligibility (17(5)) and payment to suppliers (rule 37) separately.')]),
], 0)
jjfront.add_index(wb, 'ITC reconciliation: GSTR-2B vs purchase register',
                  [('Client', '=Summary!C4'), ('GSTIN', '=Summary!C5'), ('GSTR-2B periods', '=Summary!C6'), ('Financial year', '=Summary!C7'), ('Prepared on', '=Summary!C9')],
                  [('Instructions', 'Reference', 'How to use and read this workpaper.'),
                   ('Summary', 'Summary', 'ITC position by head: available in 2B, as per books, claimable, books only, 2B only, value differences.'),
                   ('Recon_Purchases', 'Working', 'Every invoice: books vs 2B, status, match basis, GSTR-9 effect. Match code / Remarks are yours.'),
                   ('Recon_Suppliers', 'Working', 'Supplier-wise totals and differences, for vendor follow-up.'),
                   ('Month_wise', 'Working', 'ITC position by month (2B by 2B period, books by invoice date).'),
                   ('Books_Purchase', 'Source data', 'The purchase register as mapped in the Loader.'),
                   ('RAW_2B', 'Source data', 'GSTR-2B documents read from the JSON files.')],
                  'J&J GST Reco Loader (see the Loader version on Summary)')
jjfront.aptos(wb)
jjfront.add_back_links(wb)
m = wb.create_sheet('_ITC'); m['A1'] = 'itc_recon_template'; m['B1'] = '2026.11.21-2'; m.sheet_state = 'veryHidden'
wb.properties.creator = 'James & James LLP'; wb.properties.title = 'ITC Reconciliation - GSTR-2B vs books'
wb.active = 0
for w in wb.worksheets:
    w.sheet_view.tabSelected = w.title == 'Index'
wb.save(OUT)
print('saved', OUT, 'summary rows', r0, r1)

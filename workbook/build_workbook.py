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

import os
HERE = os.path.dirname(os.path.abspath(__file__))
LOADER = os.path.join(HERE, '..', 'loader')
M = json.load(open(os.path.join(HERE, 'masters.json')))
# writes straight into the loader, so a rebuild needs no copying
OUT = os.path.join(LOADER, 'template', 'GST_Annual_Return_Workbook.xlsx')

# ---------------------------------------------------------------- styles
F = 'Arial'
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


# ---------------------------------------------------------------- raw / books sheet specs
# sheet: (key column, [Taxable, CGST, SGST, IGST, Cess] columns)
SRC = {
    'G1': ('RAW_GSTR1', 'A', ['B', 'C', 'D', 'E', 'F']),
    '3B': ('RAW_3B', 'A', ['B', 'C', 'D', 'E', 'F']),
    '2B': ('RAW_2B', 'A', ['B', 'C', 'D', 'E', 'F']),
    'P9': ('RAW_Portal9', 'A', ['B', 'C', 'D', 'E', 'F']),
    'BS': ('Books_Sales', 'Q', ['S', 'T', 'U', 'V', 'W']),      # signed helper columns
    'BP': ('Books_Purchase', 'S', ['M', 'N', 'O', 'P', 'Q']),
    'BP7': ('Books_Purchase', 'T', ['M', 'N', 'O', 'P', 'Q']),  # table-7 reversal key
}


def sumifs(src, key, i):
    sh, kc, cols = SRC[src]
    return f"SUMIFS('{sh}'!${cols[i]}:${cols[i]},'{sh}'!${kc}:${kc},\"{key}\")"


def src_formula(specs, i):
    if not specs:
        return None
    parts = []
    for s in specs:
        src, key = s[0], s[1]
        sign = s[2] if len(s) > 2 else 1
        parts.append(('-' if sign < 0 else '+') + sumifs(src, key, i))
    f = ''.join(parts)
    return '=' + (f[1:] if f.startswith('+') else f)


# ---------------------------------------------------------------- GSTR-9 row model
# kind: data | sum | hdr ; cols: A all, T taxable only, X tax only
# returns/books: list of (src, key[, sign]) ; sum: list of (sign, key)
def D(code, desc, key, cols='A', ret=None, books=None, typ='', basis='Returns', locked=False, path='', note=''):
    return dict(kind='data', code=code, desc=desc, key=key, cols=cols, ret=ret, books=books, typ=typ,
                basis=basis, locked=locked, path=path, note=note)


def S(code, desc, key, terms, cols='A', typ='', path='', note=''):
    return dict(kind='sum', code=code, desc=desc, key=key, cols=cols, terms=terms, typ=typ, path=path, note=note)


def H(code, desc):
    return dict(kind='hdr', code=code, desc=desc)


T4 = [
    H('4', 'Details of advances, inward and outward supplies made during the financial year on which tax is payable'),
    D('A', 'Supplies made to un-registered persons (B2C)', '4A', ret=[('G1', '4A')], books=[('BS', '4A')], path='table4.b2c', note='GSTR-1 B2CS + B2CL (net of amendments, CN/DN)'),
    D('B', 'Supplies made to registered persons (B2B)', '4B', ret=[('G1', '4B')], books=[('BS', '4B')], path='table4.b2b', note='GSTR-1 B2B (excl. RCM)'),
    D('C', 'Zero rated supply (Export) on payment of tax (except supplies to SEZs)', '4C', ret=[('G1', '4C')], books=[('BS', '4C')], path='table4.exp', note='GSTR-1 EXPWP'),
    D('D', 'Supply to SEZs on payment of tax', '4D', ret=[('G1', '4D')], books=[('BS', '4D')], path='table4.sez', note='GSTR-1 SEZWP'),
    D('E', 'Deemed Exports', '4E', ret=[('G1', '4E')], books=[('BS', '4E')], path='table4.deemed', note='GSTR-1 DE'),
    D('F', 'Advances on which tax has been paid but invoice has not been issued', '4F', ret=[('G1', '4F')], books=[('BS', '4F')], path='table4.at', note='GSTR-1 AT less ATADJ'),
    D('G', 'Inward supplies on which tax is to be paid on reverse charge basis', '4G', ret=[('3B', '3.1d')], books=[('BP', '6C-ip'), ('BP', '6C-cg'), ('BP', '6C-is'), ('BP', '6D-ip'), ('BP', '6D-cg'), ('BP', '6D-is'), ('BP', '6F')], path='table4.rchrg', note='GSTR-3B 3.1(d)'),
    D('G1', 'Supplies on which e-commerce operator is required to pay tax as per section 9(5) [ECO to report]', '4G1', ret=[('G1', '4G1')], books=[('BS', '4G1')], path='table4.ecom', note='GSTR-1 Table 15 (ECO)'),
    S('H', 'Sub-total (A to G1 above)', '4H', [(1, '4A'), (1, '4B'), (1, '4C'), (1, '4D'), (1, '4E'), (1, '4F'), (1, '4G'), (1, '4G1')]),
    D('I', 'Credit Notes issued in respect of transactions specified in (B) to (E) above (-)', '4I', ret=[('G1', '4I')], books=[('BS', '4I')], path='table4.cr_nt', note='GSTR-1 CDNR/CDNUR credit notes'),
    D('J', 'Debit Notes issued in respect of transactions specified in (B) to (E) above (+)', '4J', ret=[('G1', '4J')], books=[('BS', '4J')], path='table4.dr_nt', note='GSTR-1 CDNR/CDNUR debit notes'),
    D('K', 'Supplies / tax declared through Amendments (+)', '4K', ret=[('G1', '4K')], books=[('BS', '4K')], path='table4.amd_pos', note='GSTR-1 9A/9C amendments (+)'),
    D('L', 'Supplies / tax reduced through Amendments (-)', '4L', ret=[('G1', '4L')], books=[('BS', '4L')], path='table4.amd_neg', note='GSTR-1 9A/9C amendments (-)'),
    S('M', 'Sub-total (I to L above)', '4M', [(-1, '4I'), (1, '4J'), (1, '4K'), (-1, '4L')]),
    S('N', 'Supplies and advances on which tax is to be paid (H + M) above', '4N', [(1, '4H'), (1, '4M')]),
]
T5 = [
    H('5', 'Details of Outward supplies made during the financial year on which tax is not payable'),
    D('A', 'Zero rated supply (Export) without payment of tax', '5A', 'T', ret=[('G1', '5A')], books=[('BS', '5A')], path='table5.zero_rtd', note='GSTR-1 EXPWOP'),
    D('B', 'Supply to SEZs without payment of tax', '5B', 'T', ret=[('G1', '5B')], books=[('BS', '5B')], path='table5.sez', note='GSTR-1 SEZWOP'),
    D('C', 'Supplies on which tax is to be paid by the recipient on reverse charge basis', '5C', 'T', ret=[('G1', '5C')], books=[('BS', '5C')], path='table5.rchrg', note='GSTR-1 B2B with RCM = Y'),
    D('C1', 'Supplies on which tax is to be paid by e-commerce operators as per section 9(5) [Supplier to report]', '5C1', 'T', ret=[('G1', '5C1')], books=[('BS', '5C1')], path='table5.ecom_14', note='GSTR-1 Table 14(b)'),
    D('D', 'Exempted', '5D', 'T', ret=[('G1', '5D')], books=[('BS', '5D')], path='table5.exmt', note='GSTR-1 Table 8 exempt'),
    D('E', 'Nil Rated', '5E', 'T', ret=[('G1', '5E')], books=[('BS', '5E')], path='table5.nil', note='GSTR-1 Table 8 nil'),
    D('F', "Non-GST supply (includes 'no supply')", '5F', 'T', ret=[('G1', '5F')], books=[('BS', '5F')], path='table5.non_gst', note='GSTR-1 Table 8 non-GST'),
    S('G', 'Sub-total (A to F above)', '5G', [(1, '5A'), (1, '5B'), (1, '5C'), (1, '5C1'), (1, '5D'), (1, '5E'), (1, '5F')], 'T'),
    D('H', 'Credit Notes issued in respect of transactions specified in A to F above (-)', '5H', 'T', ret=[('G1', '5H')], books=[('BS', '5H')], path='table5.cr_nt'),
    D('I', 'Debit Notes issued in respect of transactions specified in A to F above (+)', '5I', 'T', ret=[('G1', '5I')], books=[('BS', '5I')], path='table5.dr_nt'),
    D('J', 'Supplies declared through Amendments (+)', '5J', 'T', ret=[('G1', '5J')], books=[('BS', '5J')], path='table5.amd_pos'),
    D('K', 'Supplies reduced through Amendments (-)', '5K', 'T', ret=[('G1', '5K')], books=[('BS', '5K')], path='table5.amd_neg'),
    S('L', 'Sub-Total (H to K above)', '5L', [(-1, '5H'), (1, '5I'), (1, '5J'), (-1, '5K')], 'T'),
    S('M', 'Turnover on which tax is not to be paid (G + L) above', '5M', [(1, '5G'), (1, '5L')], 'T'),
    S('N', 'Total Turnover (including advances) (4N + 5M - 4G - 4G1) above', '5N', [(1, '4N'), (1, '5M'), (-1, '4G'), (-1, '4G1')], 'T'),
]

def itc3(code, desc, base, ret_fn, path, note=''):
    rows = []
    for i, (t, s) in enumerate([('Inputs', 'ip'), ('Capital Goods', 'cg'), ('Input Services', 'is')]):
        rows.append(D(code if i == 0 else '', desc if i == 0 else '', f'{base}-{s}', 'X', ret=ret_fn(s),
                      books=[('BP', f'{base}-{s}')], typ=t, path=f'{path}[{s}]', note=note if i == 0 else ''))
    return rows

T6 = [H('6', 'Details of ITC availed during the financial year'),
      D('A', 'Total amount of input tax credit availed through FORM GSTR-3B (sum total of Table 4A of FORM GSTR-3B)', '6A', 'X',
        ret=[('3B', '4A1'), ('3B', '4A2'), ('3B', '4A3'), ('3B', '4A4'), ('3B', '4A5')], books=None, locked=True,
        path='table6.itc_3b', note='Auto-populated & non-editable on portal (not uploaded)'),
      D('A1', 'ITC of preceding financial year availed in the financial year (included in 6A) other than ITC reclaimed under rule 37 and rule 37A', '6A1', 'X',
        ret=[('2B', '6A1')], books=[('BP', '6A1')], basis='Books', path='table6.itc_otr', note='2B of this FY carrying previous-FY documents (indicative) / books'),
      S('A2', 'Net ITC of the financial year (A - A1)', '6A2', [(1, '6A'), (-1, '6A1')], 'X'),
      ] + itc3('B', 'Inward supplies (other than imports and inward supplies liable to reverse charge but includes services received from SEZs)', '6B',
               lambda s: [('3B', '4A5')] if s == 'ip' else None, 'table6.supp_non_rchrg', 'GSTR-3B 4A(5) shown under Inputs - split per books') \
        + itc3('C', 'Inward supplies received from unregistered persons liable to reverse charge (other than B above) on which tax is paid & ITC availed', '6C',
               lambda s: [('3B', '4A3')] if s == 'ip' else None, 'table6.supp_rchrg_unreg', 'GSTR-3B 4A(3) (reg + unreg) shown here - split per books') \
        + itc3('D', 'Inward supplies received from registered persons liable to reverse charge (other than B above) on which tax is paid and ITC availed', '6D',
               lambda s: None, 'table6.supp_rchrg_reg', 'Per books') \
        + [D('E', 'Import of goods (including supplies from SEZ)', '6E-ip', 'X', ret=[('3B', '4A1')], books=[('BP', '6E-ip')], typ='Inputs', path='table6.iog[ip]', note='GSTR-3B 4A(1)'),
           D('', '', '6E-cg', 'X', ret=None, books=[('BP', '6E-cg')], typ='Capital Goods', path='table6.iog[cg]'),
           D('F', 'Import of services (excluding inward supplies from SEZs)', '6F', 'X', ret=[('3B', '4A2')], books=[('BP', '6F')], path='table6.ios', note='GSTR-3B 4A(2)'),
           D('G', 'Input Tax credit received from ISD', '6G', 'X', ret=[('3B', '4A4')], books=[('BP', '6G')], path='table6.isd', note='GSTR-3B 4A(4)'),
           D('H', 'Amount of ITC reclaimed under the provisions of the Act', '6H', 'X', ret=[('3B', '4D1')], books=[('BP', '6H')], path='table6.itc_clmd', note='GSTR-3B 4D(1)'),
           S('I', 'Sub-total (B to H above)', '6I', [(1, k) for k in ['6B-ip', '6B-cg', '6B-is', '6C-ip', '6C-cg', '6C-is', '6D-ip', '6D-cg', '6D-is', '6E-ip', '6E-cg', '6F', '6G', '6H']], 'X'),
           S('J', 'Difference (I - A2 above)', '6J', [(1, '6I'), (-1, '6A2')], 'X', note='Should be zero'),
           D('K', 'Transition Credit through TRAN-1 (including revisions if any)', '6K', 'X', ret=None, books=None, basis='Books', path='table6.tran1'),
           D('L', 'Transition Credit through TRAN-2', '6L', 'X', ret=None, books=None, basis='Books', path='table6.tran2'),
           D('M', 'ITC availed through ITC-01, ITC-02, and ITC-02A (other than B above)', '6M', 'X', ret=None, books=None, basis='Books', path='table6.other'),
           S('N', 'Sub-total (K to M above)', '6N', [(1, '6K'), (1, '6L'), (1, '6M')], 'X'),
           S('O', 'Total ITC availed (I + N) above', '6O', [(1, '6I'), (1, '6N')], 'X')]

T7 = [H('7', 'Details of ITC Reversed and Ineligible ITC for the financial year')] + [
    D(c, d, k, 'X', ret=None, books=[('BP7', k)], basis='Books', path=p) for c, d, k, p in [
        ('A', 'As per Rule 37', '7A', 'table7.rule37'),
        ('A1', 'As per Rule 37A', '7A1', 'table7.rule37A'),
        ('A2', 'As per Rule 38', '7A2', 'table7.rule38'),
        ('B', 'As per Rule 39', '7B', 'table7.rule39'),
        ('C', 'As per Rule 42', '7C', 'table7.rule42'),
        ('D', 'As per Rule 43', '7D', 'table7.rule43'),
        ('E', 'As per section 17(5)', '7E', 'table7.sec17'),
        ('F', 'Reversal of TRAN-I credit', '7F', 'table7.revsl_tran1'),
        ('G', 'Reversal of TRAN-II credit', '7G', 'table7.revsl_tran2'),
        ('H1', 'Other reversals (specify)', '7H', 'table7.other')]] + [
    S('I', 'Total ITC Reversed (Sum of A to H above)', '7I', [(1, k) for k in ['7A', '7A1', '7A2', '7B', '7C', '7D', '7E', '7F', '7G', '7H']], 'X',
      note='Offline tool v2.2 computes this excluding 7E; check against portal'),
    S('J', 'Net ITC Available for Utilization (6O - 7I)', '7J', [(1, '6O'), (-1, '7I')], 'X')]

T8 = [H('8', 'Other ITC related information'),
      D('A', 'ITC as per GSTR-2B [Table 3(I) thereof]', '8A', 'X', ret=[('2B', '8A'), ('2B', '8A+8C')], books=None, locked=True,
        path='table8.itc_2b', note='Auto-populated & non-editable on portal; returns = 2B rows tagged 8A (indicative)'),
      S('B', 'ITC as per sum total of 6(B) above', '8B', [(1, '6B-ip'), (1, '6B-cg'), (1, '6B-is')], 'X'),
      D('C', 'ITC on inward supplies (other than imports and RCM, but incl. services from SEZs) of this FY availed in the next FY', '8C', 'X',
        ret=[('2B', '8A+8C')], books=[('BP', '13')], basis='Books', path='table8.itc_inwd_supp', note='Next-FY 2B with this-FY documents (indicative) / books'),
      S('D', 'Difference [A - (B + C)]', '8D', [(1, '8A'), (-1, '8B'), (-1, '8C')], 'X'),
      D('E', 'ITC available but not availed', '8E', 'X', ret=None, books=[('BP', '8E')], basis='Books', path='table8.itc_nt_availd'),
      D('F', 'ITC available but ineligible', '8F', 'X', ret=None, books=[('BP', '8F')], basis='Books', path='table8.itc_nt_eleg'),
      D('G', 'IGST paid on import of goods (including supplies from SEZ)', '8G', 'X', ret=[('2B', '8G'), ('2B', '8G+8H1')], books=[('BP', '8G')], path='table8.iog_taxpaid', note='2B IMPG/IMPGSEZ of this FY'),
      S('H', 'IGST credit availed on import of goods (as per 6(E) above)', '8H', [(1, '6E-ip'), (1, '6E-cg')], 'X'),
      D('H1', 'IGST credit availed on import of goods in the next financial year', '8H1', 'X', ret=[('2B', '8G+8H1')], books=[('BP', '8H1')], basis='Books', path='table8.iog_nxt_yr'),
      S('I', 'Difference (G - H - H1)', '8I', [(1, '8G'), (-1, '8H'), (-1, '8H1')], 'X'),
      S('J', 'ITC available but not availed on import of goods (Equal to I)', '8J', [(1, '8I')], 'X'),
      S('K', 'Total ITC to be lapsed in current financial year (E + F + J)', '8K', [(1, '8E'), (1, '8F'), (1, '8J')], 'X')]

T10 = [H('10-13', 'Particulars of the transactions for the FY declared in returns of the next FY till the specified period'),
       D('10', 'Supplies / tax declared through Invoices / Debit Notes / Amendments (+) (net of credit notes)', '10', ret=[('G1', '10')], books=[('BS', '10')], path='table10.dbn_amd', note='Next-FY GSTR-1 documents dated this FY'),
       D('11', 'Supplies / tax reduced through Credit Notes / Amendments (-)', '11', ret=[('G1', '11')], books=[('BS', '11')], path='table10.cdn_amd', note='Next-FY GSTR-1 CN / amendments of this FY'),
       D('12', 'Reversal of ITC availed during previous financial year', '12', 'X', ret=None, books=[('BP7', '12')], basis='Books', path='table10.itc_rvsl'),
       D('13', 'ITC of the financial year availed in the next financial year', '13', 'X', ret=[('2B', '8A+8C'), ('2B', '8G+8H1')], books=[('BP', '13')], basis='Books', path='table10.itc_availd'),
       S('', 'Total turnover (5N + 10 - 11)', 'TT', [(1, '5N'), (1, '10'), (-1, '11')], 'T')]

# ---------------------------------------------------------------- GSTR-9 table sheet writer
BLOCKS = ['Portal', 'Returns', 'Books', 'Adjustment', 'To be Filed', 'Variance']
BLK_TITLE = {'Portal': 'Pre-filled (Portal)', 'Returns': 'As per Returns', 'Books': 'As per Books',
             'Adjustment': 'Adjustments (+/-)', 'To be Filed': 'To be Filed', 'Variance': 'Variance (To be Filed - Portal)'}
AMT = ['Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess']
FIRST_BLK_COL = 5  # E
COL_SRC = FIRST_BLK_COL + 30          # AI
COL_REM = COL_SRC + 1                 # AJ
COL_KEY = COL_SRC + 2                 # AK
COL_PATH = COL_SRC + 3                # AL
ROWMAP = {}   # key -> (sheet, row)


def bcol(block, i):
    return CL(FIRST_BLK_COL + BLOCKS.index(block) * 5 + i)


def write_table_sheet(name, sections, sub):
    ws = wb.create_sheet(name)
    title(ws, 'GSTR-9 Annual Return - ' + name, sub)
    r0 = 5
    for c, w in zip('ABCD', [6, 58, 13, 10]):
        ws.column_dimensions[c].width = w
    for b in BLOCKS:
        c1 = bcol(b, 0)
        ws.merge_cells(f'{c1}{r0}:{bcol(b, 4)}{r0}')
        cell(ws, f'{c1}{r0}', BLK_TITLE[b], f_hdr, fill_hdr, align=center)
        for i, a in enumerate(AMT):
            col = bcol(b, i)
            cell(ws, f'{col}{r0 + 1}', a, f_bold, fill_blk[b], align=center)
            ws.column_dimensions[col].width = 13.5
    for col, t, w in [('A', 'S.No', 6), ('B', 'Nature of Supplies / Details', 58), ('C', 'Type', 13), ('D', 'Basis', 10)]:
        ws.merge_cells(f'{col}{r0}:{col}{r0 + 1}')
        cell(ws, f'{col}{r0}', t, f_hdr, fill_hdr, align=center)
    for ci, t, w in [(COL_SRC, 'Source of "As per Returns"', 34), (COL_REM, 'Remarks / Working reference', 34),
                     (COL_KEY, 'Key', 8), (COL_PATH, 'JSON path', 26)]:
        col = CL(ci)
        ws.merge_cells(f'{col}{r0}:{col}{r0 + 1}')
        cell(ws, f'{col}{r0}', t, f_hdr, fill_hdr, align=center)
        ws.column_dimensions[col].width = w
    ws.row_dimensions[r0 + 1].height = 26
    ws.freeze_panes = f'E{r0 + 2}'
    dv = DataValidation(type='list', formula1='"Portal,Returns,Books"', allow_blank=False)
    ws.add_data_validation(dv)

    r = r0 + 2
    placed = []
    for sec in sections:
        for row in sec:
            if row['kind'] == 'hdr':
                cell(ws, f'A{r}', row['code'], f_bold, fill_sec, align=center)
                ws.merge_cells(f'B{r}:{CL(COL_PATH)}{r}')
                cell(ws, f'B{r}', row['desc'], f_bold, fill_sec)
                r += 1
                continue
            row['r'] = r
            ROWMAP[row['key']] = (name, r)
            placed.append(row)
            r += 1
    for row in placed:
        r = row['r']
        is_sum = row['kind'] == 'sum'
        base_font = f_bold if is_sum else f_norm
        cell(ws, f'A{r}', row['code'], base_font, fill_sub if is_sum else None, align=center)
        cell(ws, f'B{r}', row['desc'], base_font, fill_sub if is_sum else None, align=wrap)
        cell(ws, f'C{r}', row.get('typ', ''), f_norm, fill_sub if is_sum else None)
        cell(ws, f'{CL(COL_KEY)}{r}', row['key'], f_sub)
        cell(ws, f'{CL(COL_PATH)}{r}', row.get('path', ''), f_sub)
        cell(ws, f'{CL(COL_SRC)}{r}', row.get('note', ''), f_sub, align=wrap)
        cell(ws, f'{CL(COL_REM)}{r}', None, f_input, align=wrap)
        if is_sum:
            cell(ws, f'D{r}', None, f_norm, fill_sub)
        elif row['locked']:
            cell(ws, f'D{r}', 'Portal', f_bold, fill_na, align=center)
        else:
            cell(ws, f'D{r}', row['basis'], f_input, fill_input, align=center)
            dv.add(f'D{r}')
        for i in range(5):
            applicable = row['cols'] == 'A' or (row['cols'] == 'T' and i == 0) or (row['cols'] == 'X' and i > 0)
            for b in BLOCKS:
                col = bcol(b, i)
                ref = f'{col}{r}'
                if not applicable:
                    cell(ws, ref, None, fill=fill_na)
                    continue
                if is_sum:
                    f = ''
                    for sgn, k in row['terms']:
                        sh, rr = ROWMAP[k]
                        target = f'{col}{rr}' if sh == name else f"'{sh}'!{col}{rr}"
                        f += ('+' if sgn > 0 else '-') + target
                    cell(ws, ref, '=' + f.lstrip('+'), f_bold, fill_sub, NUM)
                    continue
                if b == 'Portal':
                    cell(ws, ref, '=' + sumifs('P9', row['key'], i), f_norm, None, NUM)
                elif b == 'Returns':
                    fm = src_formula(row['ret'], i)
                    cell(ws, ref, fm, f_norm, None if fm else fill_na, NUM)
                elif b == 'Books':
                    fm = src_formula(row['books'], i)
                    if row['locked'] and not fm:
                        cell(ws, ref, None, fill=fill_na)
                    else:
                        cell(ws, ref, fm, f_norm, None if fm else fill_input, NUM)
                        if not fm:
                            ws[ref].font = f_input
                elif b == 'Adjustment':
                    if row['locked']:
                        cell(ws, ref, None, fill=fill_na)
                    else:
                        cell(ws, ref, None, f_input, fill_input, NUM)
                elif b == 'To be Filed':
                    p, rt, bk, ad = (f'{bcol(x, i)}{r}' for x in ['Portal', 'Returns', 'Books', 'Adjustment'])
                    if row['locked']:
                        fm = f'=IF(COUNTIF(\'RAW_Portal9\'!$A$6:$A$500,"?*")>0,{p},{rt})'
                    else:
                        fm = f'=IF($D{r}="Portal",{p},IF($D{r}="Books",{bk},{rt}))+{ad}'
                    cell(ws, ref, fm, f_bold, fill_blk['To be Filed'], NUM)
                elif b == 'Variance':
                    cell(ws, ref, f'={bcol("To be Filed", i)}{r}-{bcol("Portal", i)}{r}', f_norm, None, NUM)
    ws.column_dimensions[CL(COL_KEY)].hidden = False
    # highlight non-zero variance
    rng = f'{bcol("Variance", 0)}{r0 + 2}:{bcol("Variance", 4)}{r}'
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ABS({bcol("Variance", 0)}{r0 + 2})>=1'],
                                                   font=Font(name=F, size=9, bold=True, color='C00000')))
    return ws


# ================================================================= BASIC
ws = wb.create_sheet('Basic')
ws.sheet_view.showGridLines = False
ws['B2'] = 'GST Annual Return Workbook - GSTR-9 / GSTR-9C'
ws['B2'].font = f_title
rows = [('Legal name', None, 'Client legal name as per GST registration'),
        ('Trade name', None, ''),
        ('GSTIN', None, '15-character GSTIN'),
        ('Financial year (YYYY-YY)', '2025-26', 'e.g. 2025-26'),
        ('Previous financial year', '=(VALUE(LEFT(C7,4))-1)&"-"&RIGHT(VALUE(LEFT(C7,4)),2)', 'Auto'),
        ('Next financial year', '=(VALUE(LEFT(C7,4))+1)&"-"&RIGHT(VALUE(LEFT(C7,4))+2,2)', 'Auto'),
        ('FY start date', '=DATE(VALUE(LEFT(C7,4)),4,1)', 'Auto'),
        ('FY end date', '=DATE(VALUE(LEFT(C7,4))+1,3,31)', 'Auto'),
        ('Prepared by', None, ''),
        ('Reviewed by', None, ''),
        ('Date of preparation', None, 'DD-MM-YYYY')]
for i, (lab, val, hint) in enumerate(rows):
    r = 4 + i
    cell(ws, f'B{r}', lab, f_bold, fill_sub)
    c = cell(ws, f'C{r}', val, f_input if not (isinstance(val, str) and val.startswith('=')) else f_norm,
             fill_input if not (isinstance(val, str) and val.startswith('=')) else None)
    cell(ws, f'D{r}', hint, f_sub, border=False)
ws['C10'].number_format = ws['C11'].number_format = 'DD-MMM-YYYY'
ws['C14'].number_format = 'DD-MMM-YYYY'
ws['C4'] = None
ws['C4'].value = None
ws.column_dimensions['B'].width = 28
ws.column_dimensions['C'].width = 40
ws.column_dimensions['D'].width = 40
ws['C6'].value = None
for nm, ref in [('FY_CY', '$C$7'), ('FY_PY', '$C$8'), ('FY_NY', '$C$9'), ('FY_START', '$C$10'), ('FY_END', '$C$11'), ('GSTIN', '$C$6')]:
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'Basic'!{ref}")
cell(ws, 'B17', 'GSTIN format check', f_bold, fill_sub)
cell(ws, 'C17', '=IF(C6="","Enter GSTIN",IF(LEN(C6)=15,"OK","GSTIN must be 15 characters"))', f_norm)

# ================================================================= GSTR-9 table sheets
write_table_sheet('4-5 Outward', [T4, T5], 'Tables 4 & 5 - outward supplies. Choose Basis per row; add adjustments in the yellow block.')
write_table_sheet('6-8 ITC', [T6, T7, T8], 'Tables 6, 7 & 8 - ITC availed, reversed and other ITC. 6A and 8A are portal-locked.')
write_table_sheet('10-13 Next FY', [T10], 'Tables 10 to 13 - this FY transactions declared in next FY returns (till 30 Nov).')

# ---------------------------------------------------------------- Table 9 & 14
ws = wb.create_sheet('9 & 14 Tax Paid')
title(ws, 'GSTR-9 Annual Return - Table 9 (tax paid) & Table 14 (differential tax)',
      'Table 9: only "Tax payable" is uploaded; paid columns are auto-populated by the portal from GSTR-3B.')
hdr = ['S.No', 'Description', 'Payable - Portal', 'Payable - As per 3B', 'Payable - Adjustment', 'Payable - To be Filed',
       'Paid in cash', 'Paid ITC - Central', 'Paid ITC - State/UT', 'Paid ITC - Integrated', 'Paid ITC - Cess',
       'Total paid', 'Difference (Payable - Paid)', 'Key', 'Remarks']
for j, h in enumerate(hdr):
    cell(ws, f'{CL(j + 1)}5', h, f_hdr, fill_hdr, align=center)
ws.row_dimensions[5].height = 36
for j, w in enumerate([6, 26, 14, 14, 14, 14, 14, 13, 13, 13, 13, 14, 14, 8, 30]):
    ws.column_dimensions[CL(j + 1)].width = w
heads = [('A', 'Integrated Tax', '9A'), ('B', 'Central Tax', '9B'), ('C', 'State/UT Tax', '9C'), ('D', 'Cess', '9D'),
         ('E', 'Interest', '9E'), ('F', 'Late fee', '9F'), ('G', 'Penalty', '9G'), ('H', 'Other', '9H')]
sh9, kc, cols = SRC['P9']
for i, (c, d, k) in enumerate(heads):
    r = 6 + i
    ROWMAP[k] = ('9 & 14 Tax Paid', r)
    cell(ws, f'A{r}', c, align=center)
    cell(ws, f'B{r}', d)
    cell(ws, f'C{r}', f"=SUMIFS('RAW_Portal9'!$B:$B,'RAW_Portal9'!$A:$A,\"{k}-PAY\")", fmt=NUM)
    cell(ws, f'D{r}', f"=SUMIFS('RAW_3B_PAY'!$B:$B,'RAW_3B_PAY'!$A:$A,\"{k}-PAY\")", fmt=NUM)
    cell(ws, f'E{r}', None, f_input, fill_input, NUM)
    cell(ws, f'F{r}', f'=D{r}+E{r}', f_bold, fill_blk['To be Filed'], NUM)
    cell(ws, f'G{r}', f"=SUMIFS('RAW_Portal9'!$B:$B,'RAW_Portal9'!$A:$A,\"{k}-CASH\")", fmt=NUM)
    for j, col in enumerate('HIJK'):
        src = 'CDEF'[j]
        cell(ws, f'{col}{r}', f"=SUMIFS('RAW_Portal9'!${src}:${src},'RAW_Portal9'!$A:$A,\"{k}-ITC\")", fmt=NUM)
    cell(ws, f'L{r}', f'=SUM(G{r}:K{r})', f_bold, fmt=NUM)
    cell(ws, f'M{r}', f'=F{r}-L{r}', fmt=NUM)
    cell(ws, f'N{r}', k, f_sub)
    cell(ws, f'O{r}', None, f_input)
ws['B15'] = 'Payable as per 3B is summed from GSTR-3B Table 6.1 (tax payable incl. RCM). Paid columns come from the portal pre-filled JSON.'
ws['B15'].font = f_sub
# table 14
cell(ws, 'A18', '14', f_bold, fill_sec, align=center)
ws.merge_cells('B18:O18')
cell(ws, 'B18', 'Differential tax paid on account of declaration in 10 & 11 above', f_bold, fill_sec)
for j, h in enumerate(['S.No', 'Description', 'Payable (10 - 11)', 'Payable - Adjustment', 'Payable - To be Filed', 'Paid', 'Difference', 'Key']):
    cell(ws, f'{CL(j + 1)}19', h, f_hdr, fill_hdr, align=center)
ws.row_dimensions[19].height = 30
t10 = ROWMAP['10'][1]; t11 = ROWMAP['11'][1]
for i, (c, d, k, ti) in enumerate([('A', 'Integrated Tax', '14A', 3), ('B', 'Central Tax', '14B', 1), ('C', 'State/UT Tax', '14C', 2),
                                   ('D', 'Cess', '14D', 4), ('E', 'Interest', '14E', None)]):
    r = 20 + i
    cell(ws, f'A{r}', c, align=center)
    cell(ws, f'B{r}', d)
    if ti is not None:
        col = bcol('To be Filed', ti)
        cell(ws, f'C{r}', f"='10-13 Next FY'!{col}{t10}-'10-13 Next FY'!{col}{t11}", f_link, fmt=NUM)
    else:
        cell(ws, f'C{r}', None, f_input, fill_input, NUM)
    cell(ws, f'D{r}', None, f_input, fill_input, NUM)
    cell(ws, f'E{r}', f'=C{r}+D{r}', f_bold, fill_blk['To be Filed'], NUM)
    cell(ws, f'F{r}', None, f_input, fill_input, NUM)
    cell(ws, f'G{r}', f'=E{r}-F{r}', fmt=NUM)
    cell(ws, f'H{r}', k, f_sub)
ws.freeze_panes = 'C6'

# ---------------------------------------------------------------- Table 15 & 16
ws = wb.create_sheet('15-16 Other')
title(ws, 'GSTR-9 Annual Return - Tables 15 & 16', 'Manual entry (yellow cells).')
cell(ws, 'A5', '15', f_bold, fill_sec, align=center)
ws.merge_cells('B5:J5')
cell(ws, 'B5', 'Particulars of Demands and Refunds', f_bold, fill_sec)
h15 = ['S.No', 'Details', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', 'Interest', 'Penalty', 'Late Fee / Others', 'Key']
for j, h in enumerate(h15):
    cell(ws, f'{CL(j + 1)}6', h, f_hdr, fill_hdr, align=center)
for i, (c, d, k) in enumerate([('A', 'Total Refund claimed', 'rfd_clmd'), ('B', 'Total Refund sanctioned', 'rfd_sanc'),
                               ('C', 'Total Refund Rejected', 'rfd_rejt'), ('D', 'Total Refund Pending', 'rfd_pend'),
                               ('E', 'Total demand of taxes', 'tax_dmnd'), ('F', 'Total taxes paid in respect of E above', 'tax_paid'),
                               ('G', 'Total demands pending out of E above', 'dmnd_pend')]):
    r = 7 + i
    cell(ws, f'A{r}', c, align=center); cell(ws, f'B{r}', d)
    for col in 'CDEFGHI':
        cell(ws, f'{col}{r}', None, f_input, fill_input, NUM)
    cell(ws, f'J{r}', 'table15.' + k, f_sub)
cell(ws, 'A16', '16', f_bold, fill_sec, align=center)
ws.merge_cells('B16:J16')
cell(ws, 'B16', 'Information on supplies received from composition taxpayers, deemed supply under section 143 and goods sent on approval basis', f_bold, fill_sec)
for j, h in enumerate(['S.No', 'Details', 'Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', '', '', 'Key']):
    cell(ws, f'{CL(j + 1)}17', h, f_hdr, fill_hdr, align=center)
for i, (c, d, k) in enumerate([('A', 'Supplies received from Composition taxpayers', 'comp_supp'),
                               ('B', 'Deemed supply under Section 143', 'deemed_supp'),
                               ('C', 'Goods sent on approval basis but not returned', 'not_returned')]):
    r = 18 + i
    cell(ws, f'A{r}', c, align=center); cell(ws, f'B{r}', d)
    for col in 'CDEFG':
        cell(ws, f'{col}{r}', None, f_input, fill_input, NUM)
    cell(ws, f'J{r}', 'table16.' + k, f_sub)
ws['B19'].value = 'Deemed supply under Section 143'
ws.column_dimensions['A'].width = 6; ws.column_dimensions['B'].width = 52
for col in 'CDEFGHI':
    ws.column_dimensions[col].width = 14
ws.column_dimensions['J'].width = 22

# ---------------------------------------------------------------- HSN 17 / 18
NROWS_HSN = 1500
for nm, tno, desc, srcnote in [('17 HSN Outward', '17', 'HSN wise summary of outward supplies', 'Loader fills from GSTR-1 HSN (Table 12); edit / add rows as needed'),
                               ('18 HSN Inward', '18', 'HSN wise summary of inward supplies', 'Fill from purchase register (Books_Purchase) - optional for turnover up to 5 crore')]:
    ws = wb.create_sheet(nm)
    title(ws, f'GSTR-9 Table {tno} - {desc}', srcnote)
    hdr = ['HSN Code', 'Description', 'UQC (code)', 'Total Quantity', 'Taxable Value', 'Concessional rate? (Yes/No)', 'Rate (%)',
           'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess', 'Action', 'Source', 'HSN valid?', 'UQC valid?', 'Tax check']
    for j, h in enumerate(hdr):
        cell(ws, f'{CL(j + 1)}6', h, f_hdr, fill_hdr, align=center)
    ws.row_dimensions[6].height = 36
    for j, w in enumerate([12, 30, 10, 12, 15, 12, 8, 14, 14, 14, 12, 8, 11, 10, 10, 12]):
        ws.column_dimensions[CL(j + 1)].width = w
    cell(ws, 'D5', 'Totals', f_bold, fill_sub)
    for col in 'EHIJK':
        cell(ws, f'{col}5', f'=SUM({col}7:{col}{6 + NROWS_HSN})', f_bold, fill_sub, NUM)
    cell(ws, 'M5', 'Invalid HSN:', f_bold, fill_sub)
    cell(ws, 'N5', f'=COUNTIF(N7:N{6 + NROWS_HSN},"No")', f_bold, fill_sub)
    cell(ws, 'O5', f'=COUNTIF(O7:O{6 + NROWS_HSN},"No")', f_bold, fill_sub)
    dvy = DataValidation(type='list', formula1='"Yes,No"'); ws.add_data_validation(dvy)
    dva = DataValidation(type='list', formula1='"Add,Delete"'); ws.add_data_validation(dva)
    dvr = DataValidation(type='list', formula1='"' + ','.join(str(x) for x in M['rates']) + '"'); ws.add_data_validation(dvr)
    dvu = DataValidation(type='list', formula1=f"='Masters'!$D$2:$D${1 + len(M['uqc'])}"); ws.add_data_validation(dvu)
    for r in range(7, 7 + NROWS_HSN):
        for col in 'ABCDEFGHIJKLM':
            c = ws[f'{col}{r}']
            c.font = f_input
            if col in 'DEHIJK':
                c.number_format = NUM
        ws[f'A{r}'].number_format = '@'
        ws[f'N{r}'] = f'=IF(A{r}="","",IF(COUNTIF(Masters!$A:$A,A{r})>0,"Yes","No"))'
        ws[f'O{r}'] = f'=IF(A{r}="","",IF(LEFT(A{r},2)="99","N/A",IF(COUNTIF(Masters!$D:$D,C{r})>0,"Yes","No")))'
        ws[f'P{r}'] = f'=IF(OR(A{r}="",G{r}=""),"",IF(ABS((H{r}+I{r}+J{r})-E{r}*G{r}/100)<=MAX(1,E{r}*0.001),"OK","Check"))'
        for col in 'NOP':
            ws[f'{col}{r}'].font = f_norm
    dvy.add(f'F7:F{6 + NROWS_HSN}'); dva.add(f'L7:L{6 + NROWS_HSN}')
    dvr.add(f'G7:G{6 + NROWS_HSN}'); dvu.add(f'C7:C{6 + NROWS_HSN}')
    ws.conditional_formatting.add(f'N7:P{6 + NROWS_HSN}', FormulaRule(formula=['OR(N7="No",N7="Check")'], fill=PatternFill('solid', fgColor='F8CBAD')))
    ws.freeze_panes = 'B7'

# ================================================================= GSTR-9C
ws = wb.create_sheet('GSTR-9C')
title(ws, 'GSTR-9C Reconciliation Statement (Part A)', 'Yellow = enter from audited financial statements / books. Green = linked to GSTR-9 "To be Filed".')
ws.column_dimensions['A'].width = 6; ws.column_dimensions['B'].width = 70
for col in 'CDEFGH':
    ws.column_dimensions[col].width = 15
ws.column_dimensions['I'].width = 14; ws.column_dimensions['J'].width = 34


def g9(key, i):
    sh, rr = ROWMAP[key]
    return f"'{sh}'!{bcol('To be Filed', i)}{rr}"


r = 5
R9C = {}


def sec(code, text, heads=None):
    global r
    cell(ws, f'A{r}', code, f_bold, fill_sec, align=center)
    ws.merge_cells(f'B{r}:J{r}')
    cell(ws, f'B{r}', text, f_bold, fill_sec)
    r += 1
    if heads:
        for j, h in enumerate(heads):
            cell(ws, f'{CL(j + 1)}{r}', h, f_hdr, fill_hdr, align=center)
        ws.row_dimensions[r].height = 30
        r += 1


def line(code, desc, vals, key=None, bold=False, path=None):
    """vals: list of (col, value, kind) kind in input/link/calc"""
    global r
    cell(ws, f'A{r}', code, f_bold if bold else f_norm, align=center)
    cell(ws, f'B{r}', desc, f_bold if bold else f_norm, align=wrap)
    for col, v, kind in vals:
        if kind == 'input':
            cell(ws, f'{col}{r}', v, f_input, fill_input, NUM)
        elif kind == 'link':
            cell(ws, f'{col}{r}', v, f_link, None, NUM)
        else:
            cell(ws, f'{col}{r}', v, f_bold if bold else f_norm, fill_sub if bold else None, NUM)
    if key:
        R9C[key] = r
    if path:
        cell(ws, f'I{r}', path, f_sub)
    cell(ws, f'J{r}', None, f_input, align=wrap)
    r += 1


# Part II table 5
sec('5', 'Reconciliation of Gross Turnover', ['S.No', 'Description', 'Amount', '', '', '', '', '', 'JSON', 'Remarks'])
t5 = [('A', 'Turnover (including exports) as per audited financial statements for the State / UT (for multi-GSTIN units)', 'A'),
      ('B', 'Unbilled revenue at the beginning of Financial Year', 'B'), ('C', 'Unadjusted advances at the end of the Financial Year', 'C'),
      ('D', 'Deemed Supply under Schedule I', 'D'), ('E', 'Credit Notes issued after the end of the financial year but reflected in the annual return', 'E'),
      ('F', 'Trade Discounts accounted for in the audited Annual Financial Statement but are not permissible under GST', 'F'),
      ('G', 'Turnover from April 2017 to June 2017', 'G'), ('H', 'Unbilled revenue at the end of Financial Year', 'H'),
      ('I', 'Unadjusted Advances at the beginning of the Financial Year', 'I'),
      ('J', 'Credit notes accounted for in the audited Annual Financial Statement but are not permissible under GST', 'J'),
      ('K', 'Adjustments on account of supply of goods by SEZ units to DTA Units', 'K'), ('L', 'Turnover for the period under composition scheme', 'L'),
      ('M', 'Adjustments in turnover under section 15 and rules thereunder (+/-)', 'M'), ('N', 'Adjustments in turnover due to foreign exchange fluctuations (+/-)', 'N'),
      ('O', 'Adjustments in turnover due to reasons not listed above (+/-)', 'O')]
for c, d, k in t5:
    line(c, d, [('C', 0 if c == 'G' else None, 'input')], '5' + k, path='table5')
f5 = '=C{A}+C{B}+C{C}+C{D}-C{E}+C{F}-C{G}-C{H}-C{I}+C{J}-C{K}-C{L}+C{M}+C{N}+C{O}'.format(**{k: R9C['5' + k] for k in 'ABCDEFGHIJKLMNO'})
line('P', 'Annual turnover after adjustments as above (A+B+C+D-E+F-G-H-I+J-K-L+/-M+/-N+/-O)', [('C', f5, 'calc')], '5P', True)
line('Q', 'Turnover as declared in Annual Return (GSTR-9) [5N + 10 - 11]', [('C', '=' + g9('TT', 0), 'link')], '5Q')
line('R', 'Un-Reconciled turnover (Q - P)', [('C', f'=C{R9C["5Q"]}-C{R9C["5P"]}', 'calc')], '5R', True)
sec('6', 'Reasons for Un-Reconciled difference in Annual Gross Turnover', ['S.No', 'Reason', 'Amount', '', '', '', '', '', '', 'Remarks'])
for i in range(1, 4):
    line(f'Reason {i}', None, [('C', None, 'input')], f'6R{i}')
    ws[f'B{r - 1}'].font = f_input; ws[f'B{r - 1}'].fill = fill_input
# table 7
sec('7', 'Reconciliation of Taxable Turnover', ['S.No', 'Description', 'Amount', '', '', '', '', '', '', 'Remarks'])
line('A', 'Annual turnover after adjustments (from 5P above)', [('C', f'=C{R9C["5P"]}', 'calc')], '7A')
line('B', 'Value of Exempted, Nil Rated, Non-GST supplies, No-Supply turnover', [('C', f'={g9("5D", 0)}+{g9("5E", 0)}+{g9("5F", 0)}', 'link')], '7B')
line('C', 'Zero rated supplies without payment of tax', [('C', f'={g9("5A", 0)}+{g9("5B", 0)}', 'link')], '7C')
line('D', 'Supplies on which tax is to be paid by the recipient on reverse charge basis', [('C', '=' + g9('5C', 0), 'link')], '7D')
line('D1', 'Supplies on which tax is to be paid by e-commerce operators as per section 9(5)', [('C', '=' + g9('5C1', 0), 'link')], '7D1')
line('E', 'Taxable turnover as per adjustments above (A-B-C-D-D1)', [('C', '=C{0}-C{1}-C{2}-C{3}-C{4}'.format(*[R9C[k] for k in ['7A', '7B', '7C', '7D', '7D1']]), 'calc')], '7E', True)
line('F', 'Taxable turnover as per liability declared in Annual Return (GSTR-9) [4N - 4G + 10 - 11, per ICAI TG]',
     [('C', f'={g9("4N", 0)}-{g9("4G", 0)}+{g9("10", 0)}-{g9("11", 0)}', 'link')], '7F')
line('G', 'Unreconciled taxable turnover (F - E)', [('C', f'=C{R9C["7F"]}-C{R9C["7E"]}', 'calc')], '7G', True)
sec('8', 'Reasons for Un-Reconciled difference in taxable turnover', ['S.No', 'Reason', 'Amount', '', '', '', '', '', '', 'Remarks'])
for i in range(1, 4):
    line(f'Reason {i}', None, [('C', None, 'input')], f'8R{i}')
    ws[f'B{r - 1}'].font = f_input; ws[f'B{r - 1}'].fill = fill_input
# table 9
sec('9', 'Reconciliation of rate wise liability and amount payable thereon (default = Books_Sales / Books_Purchase by rate; overwrite if needed)',
    ['S.No', 'Rate', 'Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', '', 'JSON', 'Remarks'])
BS_ = "'Books_Sales'!"
BP_ = "'Books_Purchase'!"
rate_rows = [('A', '5%', 5, False), ('B', '5% (RC)', 5, True), ('B1', '6%', 6, False), ('C', '12%', 12, False), ('D', '12% (RC)', 12, True),
             ('E', '18%', 18, False), ('F', '18% (RC)', 18, True), ('G', '28%', 28, False), ('H', '28% (RC)', 28, True),
             ('H1', '40%', 40, False), ('H2', '40% (RC)', 40, True), ('I', '3%', 3, False), ('J', '0.25%', 0.25, False), ('K', '0.10%', 0.1, False)]
first9 = r
for c, lab, rt, rc in rate_rows:
    vals = []
    for col, (sc, pc) in zip('CDEFG', [('S', 'M'), ('T', 'N'), ('U', 'O'), ('V', 'P'), ('W', 'Q')]):
        if rc:
            f = '+'.join(f"SUMIFS({BP_}${pc}:${pc},{BP_}$L:$L,{rt},{BP_}$S:$S,\"{k}\")" for k in ['6C-*', '6D-*', '6F'])
        else:
            f = f"SUMIFS({BS_}${sc}:${sc},{BS_}$K:$K,{rt},{BS_}$Q:$Q,\"4*\")-2*SUMIFS({BS_}${sc}:${sc},{BS_}$K:$K,{rt},{BS_}$Q:$Q,\"4I\")"
        vals.append((col, '=' + f, 'calc'))
    line(c, lab, vals, '9' + c, path='table9')
line('K1', 'Others %', [(col, None, 'input') for col in 'CDEFG'], '9K1')
line('K2', 'Supplies on which e-commerce operator is required to pay tax as per section 9(5)', [(col, None, 'input') for col in 'CDEFG'], '9K2')
for c, d in [('L', 'Interest'), ('M', 'Late Fee'), ('N', 'Penalty'), ('O', 'Others')]:
    line(c, d, [(col, None, 'input') for col in 'DEFG'], '9' + c)
    ws[f'C{r - 1}'].fill = fill_na
last9 = r - 1
line('P', 'Total amount to be paid as per tables above (A to O)', [(col, f'=SUM({col}{first9}:{col}{last9})', 'calc') for col in 'DEFG'], '9P', True)
ws[f'C{r - 1}'].fill = fill_na
q = [('D', '9B'), ('E', '9C'), ('F', '9A'), ('G', '9D')]
q14 = {'9A': 20, '9B': 21, '9C': 22, '9D': 23}
line('Q', 'Total amount payable as declared in Annual Return (GSTR-9) - Table 9 payable + Table 14 differential tax',
     [(col, f"='9 & 14 Tax Paid'!F{ROWMAP[k][1]}+'9 & 14 Tax Paid'!E{q14[k]}", 'link') for col, k in q], '9Q')
ws[f'C{r - 1}'].fill = fill_na
line('R', 'Un-reconciled payment of amount (Q - P)', [(col, f'={col}{R9C["9Q"]}-{col}{R9C["9P"]}', 'calc') for col in 'DEFG'], '9R', True)
ws[f'C{r - 1}'].fill = fill_na
sec('10', 'Reasons for un-reconciled payment of amount', ['S.No', 'Reason', '', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', '', '', 'Remarks'])
for i in range(1, 4):
    line(f'Reason {i}', None, [(col, None, 'input') for col in 'DEFG'], f'10R{i}')
    ws[f'B{r - 1}'].font = f_input; ws[f'B{r - 1}'].fill = fill_input
sec('11', 'Additional amount payable but not paid (due to reasons specified in Tables 6, 8 and 10 above)',
    ['S.No', 'Rate', 'Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', '', 'JSON', 'Remarks'])
for c, lab in [('A', '5%'), ('A1', '6%'), ('B', '12%'), ('C', '18%'), ('D', '28%'), ('D1', '40%'), ('E', '3%'), ('F', '0.25%'), ('G', '0.10%'),
               ('G1', 'Others %'), ('G2', 'Supplies on which ECO is required to pay tax u/s 9(5)'), ('H', 'Interest'), ('I', 'Late Fee'),
               ('J', 'Penalty'), ('K', 'Others (please specify)')]:
    line(c, lab, [(col, None, 'input') for col in 'CDEFG'], '11' + c, path='table11')
# Part IV
sec('12', 'Reconciliation of Net Input Tax Credit (ITC)', ['S.No', 'Description', 'Amount', '', '', '', '', '', 'JSON', 'Remarks'])
line('A', 'ITC availed as per audited Annual Financial Statement for the State / UT (for multi-GSTIN units)', [('C', None, 'input')], '12A', path='table12')
line('B', 'ITC booked in earlier Financial Years claimed in current Financial Year', [('C', None, 'input')], '12B')
line('C', 'ITC booked in current Financial Year to be claimed in subsequent Financial Years', [('C', None, 'input')], '12C')
line('D', 'ITC availed as per audited financial statements or books of account (A+B-C)', [('C', f'=C{R9C["12A"]}+C{R9C["12B"]}-C{R9C["12C"]}', 'calc')], '12D', True)
sh7, r7 = ROWMAP['7J']
line('E', 'ITC claimed in Annual Return (GSTR-9) [7J]', [('C', '=' + '+'.join(f"'{sh7}'!{bcol('To be Filed', i)}{r7}" for i in range(1, 5)), 'link')], '12E')
line('F', 'Un-reconciled ITC (E - D)', [('C', f'=C{R9C["12E"]}-C{R9C["12D"]}', 'calc')], '12F', True)
sec('13', 'Reasons for un-reconciled difference in ITC', ['S.No', 'Reason', 'Amount', '', '', '', '', '', '', 'Remarks'])
for i in range(1, 4):
    line(f'Reason {i}', None, [('C', None, 'input')], f'13R{i}')
    ws[f'B{r - 1}'].font = f_input; ws[f'B{r - 1}'].fill = fill_input
sec('14', 'Reconciliation of ITC declared in Annual Return (GSTR-9) with ITC availed on expenses as per audited Annual Financial Statement or books of account',
    ['S.No', 'Description', 'Value', 'Amount of Total ITC', 'Eligible ITC availed', '', '', '', 'JSON', 'Remarks'])
first14 = r
for c, d in [('A', 'Purchases'), ('B', 'Freight / Carriage'), ('C', 'Power and Fuel'), ('D', 'Imported goods (Including received from SEZs)'),
             ('E', 'Rent and Insurance'), ('F', 'Goods lost, stolen, destroyed, written off or disposed of by way of gift or free samples'),
             ('G', 'Royalties'), ('H', "Employees' Cost (Salaries, wages, Bonus etc.)"), ('I', 'Conveyance charges'), ('J', 'Bank Charges'),
             ('K', 'Entertainment charges'), ('L', 'Stationery Expenses (including postage etc.)'), ('M', 'Repair and Maintenance'),
             ('N', 'Other Miscellaneous expenses'), ('O', 'Capital goods'), ('P', 'Any other expense 1'), ('Q', 'Any other expense 2')]:
    line(c, d, [(col, None, 'input') for col in 'CDE'], '14' + c, path='table14')
last14 = r - 1
line('R', 'Total amount of eligible ITC availed (A to Q)', [('E', f'=SUM(E{first14}:E{last14})', 'calc')], '14R', True)
line('S', 'ITC claimed in Annual Return (GSTR-9)', [('E', f'=C{R9C["12E"]}', 'calc')], '14S')
line('T', 'Un-reconciled ITC (S - R)', [('E', f'=E{R9C["14S"]}-E{R9C["14R"]}', 'calc')], '14T', True)
sec('15', 'Reasons for un-reconciled difference in ITC', ['S.No', 'Reason', 'Amount', '', '', '', '', '', '', 'Remarks'])
for i in range(1, 4):
    line(f'Reason {i}', None, [('C', None, 'input')], f'15R{i}')
    ws[f'B{r - 1}'].font = f_input; ws[f'B{r - 1}'].fill = fill_input
sec('16', 'Tax payable on un-reconciled difference in ITC (due to reasons specified in 13 and 15 above)', ['S.No', 'Description', 'Amount payable', '', '', '', '', '', 'JSON', 'Remarks'])
for c, d in [('A', 'Central Tax'), ('B', 'State/UT Tax'), ('C', 'Integrated Tax'), ('D', 'Cess'), ('E', 'Interest'), ('F', 'Penalty')]:
    line(c, d, [('C', None, 'input')], '16' + c, path='table16')
sec('Pt V', 'Additional liability due to non-reconciliation (paid through cash / ITC)',
    ['S.No', 'Description', 'Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', '', 'JSON', 'Remarks'])
for c, d in [('A', '5%'), ('A1', '6%'), ('B', '12%'), ('C', '18%'), ('D', '28%'), ('D1', '40%'), ('E', '3%'), ('F', '0.25%'), ('G', '0.10%'),
             ('G1', 'Others %'), ('G2', 'Supplies on which ECO is required to pay tax u/s 9(5)'), ('H', 'Input Tax Credit'), ('I', 'Interest'),
             ('J', 'Late Fee'), ('K', 'Penalty'), ('L', 'Any other amount paid for supplies not included in Annual Return (GSTR-9)'),
             ('M', 'Erroneous refund to be paid back'), ('N', 'Outstanding demands to be settled'), ('O', 'Other (Pl. specify)')]:
    line(c, d, [(col, None, 'input') for col in 'CDEFG'], 'PV' + c, path='add_liab')
ws.freeze_panes = 'C5'
END9C = r

# ================================================================= CHECKS
ws = wb.create_sheet('Checks')
title(ws, 'Review checks', 'Status turns red when a check fails. Tolerance Rs 1 unless stated.')
for j, h in enumerate(['#', 'Check', 'Value', 'Status', 'What to do']):
    cell(ws, f'{CL(j + 1)}5', h, f_hdr, fill_hdr, align=center)
for col, w in zip('ABCDE', [5, 62, 16, 12, 60]):
    ws.column_dimensions[col].width = w


def tf(key, i):
    return g9(key, i)


checks = []
for i, head in [(1, 'Central'), (2, 'State/UT'), (3, 'Integrated'), (4, 'Cess')]:
    checks.append((f'6J difference (I - A2) - {head} tax', f'={tf("6J", i)}', 'zero', 'Explain / correct split of ITC in Table 6'))
for i, head in [(1, 'Central'), (2, 'State/UT'), (3, 'Integrated'), (4, 'Cess')]:
    checks.append((f'8D difference [A-(B+C)] - {head} tax', f'={tf("8D", i)}', 'info', 'Positive = ITC in 2B not availed; negative = availed beyond 2B. Explain in working.'))
for i, (head, k) in enumerate([('Integrated', '9A'), ('Central', '9B'), ('State/UT', '9C'), ('Cess', '9D')]):
    checks.append((f'Table 9 payable vs 4N tax - {head}', f"='9 & 14 Tax Paid'!F{ROWMAP[k][1]}-{tf('4N', [3, 1, 2, 4][i])}", 'zero',
                   'Table 9 payable should equal tax in 4N (incl. RCM 4G). Differences go to Table 14 / DRC-03.'))
for i, (head, k) in enumerate([('Integrated', '9A'), ('Central', '9B'), ('State/UT', '9C'), ('Cess', '9D')]):
    checks.append((f'Table 9 payable not fully paid - {head} (after portal JSON is loaded)', f"=IF(COUNTIF('RAW_Portal9'!$A$6:$A$500,\"?*\")>0,'9 & 14 Tax Paid'!M{ROWMAP[k][1]},0)", 'zero', 'Unpaid liability - pay through DRC-03 before filing'))
checks.append(('CGST <> SGST in To be Filed rows (count of rows)',
               "=SUMPRODUCT(--(ROUND('4-5 Outward'!{0}7:{0}60,2)<>ROUND('4-5 Outward'!{1}7:{1}60,2)))+SUMPRODUCT(--(ROUND('6-8 ITC'!{0}7:{0}80,2)<>ROUND('6-8 ITC'!{1}7:{1}80,2)))".format(bcol('To be Filed', 1), bcol('To be Filed', 2)),
               'zero', 'Portal warns when Central and State tax differ'))
checks.append(('Invalid HSN codes in Table 17', "='17 HSN Outward'!N5", 'zero', 'From FY 2025-26 the offline tool blocks JSON generation for invalid HSN'))
checks.append(('Invalid UQC in Table 17', "='17 HSN Outward'!O5", 'zero', 'Use UQC codes from Masters sheet'))
checks.append(('Table 17 taxable value vs (4N - 4G - 4G1 - 4F + 5M) [indicative]',
               f"='17 HSN Outward'!E5-({tf('4N', 0)}-{tf('4G', 0)}-{tf('4G1', 0)}-{tf('4F', 0)}+{tf('5M', 0)})", 'info', 'Large gaps usually mean missing HSN rows'))
checks.append(('Portal GSTR-9 JSON loaded (RAW_Portal9)', "=IF(COUNTIF('RAW_Portal9'!$A$6:$A$500,\"?*\")>0,0,1)", 'zero', 'Load the auto-drafted GSTR-9 JSON through the Loader before finalising'))
checks.append(('9C 5R - un-reconciled gross turnover', f"='GSTR-9C'!C{R9C['5R']}", 'info', 'Explain in 9C Table 6'))
checks.append(('9C 7G - un-reconciled taxable turnover', f"='GSTR-9C'!C{R9C['7G']}", 'info', 'Explain in 9C Table 8'))
checks.append(('9C 9R - un-reconciled tax (total of heads)', f"=SUM('GSTR-9C'!D{R9C['9R']}:G{R9C['9R']})", 'info', 'Explain in 9C Table 10'))
checks.append(('9C 12F - un-reconciled ITC', f"='GSTR-9C'!C{R9C['12F']}", 'info', 'Explain in 9C Table 13'))
checks.append(('9C 12C vs GSTR-9 Table 13 (total tax)', f"='GSTR-9C'!C{R9C['12C']}-(" + '+'.join(g9('13', i) for i in range(1, 5)) + ')', 'info', 'ICAI TG: 12C should equal GSTR-9 Table 13'))
checks.append(('Purchases in books, not in 2B - tax (Recon_Purchases)', "=Recon_Purchases!F11", 'info', 'Follow up with suppliers; ITC not in 2B cannot be availed (section 16(2)(aa))'))
checks.append(('Purchases in 2B, not in books - tax (Recon_Purchases)', "=-Recon_Purchases!F12", 'info', 'Book the invoice, or report in 8E / 8F if not availed'))
checks.append(('Sales in books, not in GSTR-1 - taxable (Recon_Sales)', "=Recon_Sales!C10", 'info', 'Report through Table 10 / DRC-03 if tax unpaid'))
checks.append(('Sales in GSTR-1, not in books - taxable (Recon_Sales)', "=Recon_Sales!D11", 'info', 'Check for missing entries in books'))
checks.append(('Books_Sales rows without a valid GSTR-9 key', "=COUNTIF('Books_Sales'!$Q$6:$Q$20000,\"?\")", 'zero', 'Fix Supply Type / Doc Type on those rows'))
checks.append(('Books_Purchase rows without a valid GSTR-9 key', "=COUNTIF('Books_Purchase'!$S$6:$S$20000,\"?\")", 'zero', 'Fix Supply Type / ITC Type / ITC Status'))
checks.append(('GSTIN entered and 15 characters', "=IF(LEN(GSTIN)=15,0,1)", 'zero', 'Basic sheet'))
for i, (t, f, kind, todo) in enumerate(checks):
    r = 6 + i
    cell(ws, f'A{r}', i + 1, align=center)
    cell(ws, f'B{r}', t, align=wrap)
    cell(ws, f'C{r}', f, fmt=NUM)
    if kind == 'zero':
        cell(ws, f'D{r}', f'=IF(ABS(C{r})<1,"OK","Check")', f_bold, align=center)
    else:
        cell(ws, f'D{r}', f'=IF(ABS(C{r})<1,"OK","Review")', f_bold, align=center)
    cell(ws, f'E{r}', todo, f_sub, align=wrap)
nchk = 6 + len(checks) - 1
ws.conditional_formatting.add(f'D6:D{nchk}', FormulaRule(formula=['D6="Check"'], fill=PatternFill('solid', fgColor='F8CBAD'), font=Font(name=F, size=9, bold=True, color='C00000')))
ws.conditional_formatting.add(f'D6:D{nchk}', FormulaRule(formula=['D6="Review"'], fill=PatternFill('solid', fgColor='FFE699')))
ws.conditional_formatting.add(f'D6:D{nchk}', FormulaRule(formula=['D6="OK"'], fill=PatternFill('solid', fgColor='C6EFCE')))

# ================================================================= RAW & BOOKS sheets
def data_sheet(name, sub, headers, widths, example=None, input_cols='', formulas=None, nform=0):
    ws = wb.create_sheet(name)
    title(ws, name, sub)
    for j, (h, w) in enumerate(zip(headers, widths)):
        cell(ws, f'{CL(j + 1)}5', h, f_hdr, fill_hdr, align=center)
        ws.column_dimensions[CL(j + 1)].width = w
    ws.row_dimensions[5].height = 30
    ws.freeze_panes = 'A6'
    if example:
        for j, v in enumerate(example):
            c = ws[f'{CL(j + 1)}6']
            c.value = v
            c.font = f_input if CL(j + 1) in input_cols else f_norm
    if formulas:
        for rr in range(6, 6 + nform):
            for col, tmpl in formulas.items():
                ws[f'{col}{rr}'] = tmpl.format(r=rr)
                ws[f'{col}{rr}'].font = f_norm
                ws[f'{col}{rr}'].number_format = NUM if (name == 'Books_Sales' and col in 'STUVW') else 'General'
    ws.auto_filter.ref = f'A5:{CL(len(headers))}5'
    return ws

AH = ['Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess']
data_sheet('RAW_Portal9', 'Filled by the loader from the portal auto-drafted GSTR-9 JSON. Table 9 keys: 9A-PAY / 9A-CASH (col B), 9A-ITC (cols C-F = ITC used by head).',
           ['Key'] + AH + ['JSON path', 'Description'], [10, 15, 15, 15, 15, 15, 32, 50])
data_sheet('RAW_GSTR1', 'Filled by the loader from GSTR-1 / IFF / GSTR-1A JSONs of previous, current and next FY. Key = GSTR-9 row the document belongs to for the FY in Basic.',
           ['Key'] + AH + ['Return FY', 'Return period', 'GSTR-1 section', 'Doc type', 'Doc number', 'Doc date', 'Doc FY', 'Counterparty GSTIN', 'POS', 'Rate', 'RCM', 'Source file'],
           [8, 14, 13, 13, 13, 11, 10, 10, 11, 8, 16, 11, 9, 17, 6, 6, 5, 24])
data_sheet('RAW_GSTR1_HSN', 'Filled by the loader from GSTR-1 Table 12 (HSN) - summed into 17 HSN Outward.',
           ['HSN', 'Description', 'UQC', 'Quantity', 'Taxable Value', 'Rate', 'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess', 'Return period', 'B2B/B2C'],
           [12, 30, 8, 12, 15, 6, 14, 14, 14, 12, 11, 8])
data_sheet('RAW_3B', 'Filled by the loader from GSTR-3B PDFs. Key = 3B table (3.1a, 3.1b, 3.1c, 3.1d, 3.1e, 3.1.1i, 3.1.1ii, 4A1-4A5, 4B1, 4B2, 4D1, 4D2, 5a, 5b).',
           ['Key (3B table)'] + AH + ['Return FY', 'Return period', 'ARN date', '3B label', 'Source file'], [12, 15, 14, 14, 14, 12, 10, 11, 11, 40, 24])
data_sheet('RAW_3B_PAY', 'Filled by the loader from GSTR-3B Table 6.1 (payment of tax). Keys 9A-PAY..9D-PAY (tax payable), 9E-PAY interest, 9F-PAY late fee.',
           ['Key', 'Amount', 'Return period', 'Source file'], [12, 15, 12, 24])
data_sheet('RAW_2B', 'Filled by the loader from GSTR-2B JSONs (this FY + next FY up to Oct/Nov) and optional 2A. Key: 8A = this-FY doc in this-FY 2B; 8A+8C = this-FY doc in next-FY 2B; 6A1 = previous-FY doc in this-FY 2B; 8G / 8G+8H1 = imports; blank = not counted.',
           ['Key'] + AH + ['Source (2B/2A)', '2B FY', '2B period', 'Section', 'Supplier GSTIN', 'Supplier name', 'Doc type', 'Doc number', 'Doc date', 'Doc FY', 'ITC available', 'RCM', 'Match with books'],
           [8, 14, 13, 13, 13, 11, 9, 9, 10, 9, 17, 28, 8, 16, 11, 9, 9, 6, 16])

# Books_Sales
SALES_TYPES = ['B2B', 'B2C', 'EXP-WP', 'EXP-WOP', 'SEZ-WP', 'SEZ-WOP', 'DEEMED', 'ADVANCE', 'RCM-SUPPLY', 'ECO-9(5)', 'EXEMPT', 'NIL', 'NON-GST']
DOC_TYPES = ['INV', 'CN', 'DN', 'ADV', 'ADJ']
NB = 3000
sales_f = {
    'Q': '=IF(L{r}&M{r}&N{r}&O{r}&P{r}="","",IF(AND(A{r}<>"",A{r}>FY_END),IF(G{r}="CN","11","10"),IFERROR(INDEX(Map!$C:$C,MATCH(F{r}&"|"&G{r},Map!$A:$A,0)),"?")))',
    'R': '=IF(Q{r}="","",IFERROR(INDEX(Map!$D:$D,MATCH(F{r}&"|"&G{r},Map!$A:$A,0)),1))',
    'S': '=IF(Q{r}="",0,N(L{r})*R{r})', 'T': '=IF(Q{r}="",0,N(M{r})*R{r})', 'U': '=IF(Q{r}="",0,N(N{r})*R{r})',
    'V': '=IF(Q{r}="",0,N(O{r})*R{r})', 'W': '=IF(Q{r}="",0,N(P{r})*R{r})',
}
ws = data_sheet('Books_Sales', 'Paste the sales register here (cols A-P). Supply Type / Doc Type drive the GSTR-9 key (col Q). Documents dated after FY end go to Table 10/11.',
                ['Doc date', 'Doc number', 'Customer', 'Customer GSTIN', 'POS', 'Supply Type', 'Doc Type', 'HSN', 'UQC', 'Quantity', 'Rate',
                 'Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', 'GSTR-9 Key', 'Sign', 'Taxable (signed)', 'CGST (signed)',
                 'SGST (signed)', 'IGST (signed)', 'Cess (signed)', 'Remarks'],
                [11, 14, 24, 17, 6, 11, 8, 10, 7, 10, 6, 14, 12, 12, 12, 10, 9, 5, 14, 12, 12, 12, 10, 24],
                example=[None] * 16, input_cols='ABCDEFGHIJKLMNOP', formulas=sales_f, nform=NB)
dv = DataValidation(type='list', formula1='"' + ','.join(SALES_TYPES) + '"'); ws.add_data_validation(dv); dv.add(f'F6:F{5 + NB}')
dv = DataValidation(type='list', formula1='"' + ','.join(DOC_TYPES) + '"'); ws.add_data_validation(dv); dv.add(f'G6:G{5 + NB}')
for rr in range(6, 6 + NB):
    ws[f'A{rr}'].number_format = 'DD-MM-YYYY'
    for col in 'LMNOP':
        ws[f'{col}{rr}'].number_format = NUM
        ws[f'{col}{rr}'].font = f_input
    ws[f'Q{rr}'].number_format = '@'

# Books_Purchase
PUR_TYPES = ['B2B', 'RCM-REG', 'RCM-UNREG', 'IMPG', 'IMPG-SEZ', 'IMPS', 'SEZ-SERVICES', 'ISD']
ITC_TYPES = ['Inputs', 'Capital Goods', 'Input Services']
ITC_STATUS = ['Availed', 'Reclaimed-Rule37/37A', 'Reclaimed-Other', 'Not availed', 'Ineligible', 'Reversed-Rule37', 'Reversed-Rule37A', 'Reversed-Rule38', 'Reversed-Rule39',
              'Reversed-Rule42', 'Reversed-Rule43', 'Reversed-17(5)', 'Reversed-Other', 'Reversed-PY ITC (Table 12)']
pur_f = {
    'S': ('=IF(M{r}&N{r}&O{r}&P{r}&Q{r}="","",IF(H{r}=FY_NY,"13",IF(G{r}="Not availed","8E",IF(G{r}="Ineligible","8F",IF(G{r}="Reclaimed","6H",'
          'IF(LEFT(G{r},8)="Reversed",IF(G{r}="Reversed-PY ITC (Table 12)","",'
          'IFERROR(INDEX(Map!$H:$H,MATCH(E{r}&"|"&F{r},Map!$G:$G,0)),"?")),'
          'IF(AND(A{r}<>"",A{r}<FY_START),"6A1",IFERROR(INDEX(Map!$H:$H,MATCH(E{r}&"|"&F{r},Map!$G:$G,0)),"?"))))))))'),
    'T': '=IF(LEFT(G{r},8)<>"Reversed","",IFERROR(INDEX(Map!$K:$K,MATCH(G{r},Map!$J:$J,0)),"?"))',
}
ws = data_sheet('Books_Purchase', 'Paste the purchase / ITC register (cols A-R). Reversed lines feed Table 6 (as availed) AND Table 7 (key col T). Claimed in next FY -> Table 13 / 8C.',
                ['Doc date', 'Doc number', 'Supplier', 'Supplier GSTIN', 'Supply Type', 'ITC Type', 'ITC Status', 'ITC claimed in FY', 'HSN', 'UQC',
                 'Quantity', 'Rate', 'Taxable Value', 'Central Tax', 'State/UT Tax', 'Integrated Tax', 'Cess', 'Remarks', 'GSTR-9 Key (Table 6/8/13)', 'Table 7/12 Key'],
                [11, 14, 24, 17, 12, 13, 22, 10, 10, 7, 10, 6, 14, 12, 12, 12, 10, 22, 12, 10],
                input_cols='ABCDEFGHIJKLMNOPQR', formulas=pur_f, nform=NB)
# fix: Reversed lines should feed table 6 too (as availed). Use availed mapping for them:
pur_f['S'] = ('=IF(M{r}&N{r}&O{r}&P{r}&Q{r}="","",IF(H{r}=FY_NY,"13",IF(G{r}="Not availed","8E",IF(G{r}="Ineligible","8F",IF(G{r}="Reclaimed-Rule37/37A","6H",IF(AND(G{r}="Reclaimed-Other",A{r}<>"",A{r}<FY_START),"6A1",IF(G{r}="Reclaimed-Other","6H",'
              'IF(G{r}="Reversed-PY ITC (Table 12)","",IF(AND(A{r}<>"",A{r}<FY_START),"6A1",IFERROR(INDEX(Map!$H:$H,MATCH(E{r}&"|"&F{r},Map!$G:$G,0)),"?"))))))))))')
pur_f['T'] = '=IF(LEFT(G{r},8)<>"Reversed","",IFERROR(INDEX(Map!$K:$K,MATCH(G{r},Map!$J:$J,0)),"?"))'
for rr in range(6, 6 + NB):
    ws[f'S{rr}'] = pur_f['S'].format(r=rr)
    ws[f'T{rr}'] = pur_f['T'].format(r=rr)
    ws[f'S{rr}'].font = ws[f'T{rr}'].font = f_norm
    ws[f'S{rr}'].number_format = ws[f'T{rr}'].number_format = 'General'
    ws[f'A{rr}'].number_format = 'DD-MM-YYYY'
    for col in 'MNOPQ':
        ws[f'{col}{rr}'].number_format = NUM
        ws[f'{col}{rr}'].font = f_input
for col, lst in [('E', PUR_TYPES), ('F', ITC_TYPES), ('G', ITC_STATUS)]:
    dv = DataValidation(type='list', formula1=f"=Map!${ {'E': 'M', 'F': 'N', 'G': 'J'}[col]}$2:${ {'E': 'M', 'F': 'N', 'G': 'J'}[col]}${1 + len(lst)}")
    ws.add_data_validation(dv); dv.add(f'{col}6:{col}{5 + NB}')

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

P_STATUS = ['Matched', 'Matched - value difference', 'Matched - invoice no. differs', 'Matched - supplier GSTIN differs',
            'Matched - in next FY 2B', 'Matched - supplier total', 'Matched - manual', 'In books, not in 2B', 'In 2B, not in books', 'Not compared (RCM / ISD / imports)']
recon_sheet('Recon_Purchases', 'Purchase register vs GSTR-2B, invoice by invoice. Manual match: give the rows that belong together the same Match code (e.g. M1), save, and drop the workbook into the Loader. Codes and remarks are kept on refresh.',
            ['Status', 'Match basis', 'Supplier GSTIN', 'Supplier name', 'Invoice no. (books)', 'Invoice no. (2B)', 'Date (books)', 'Date (2B)',
             'Books taxable', 'Books CGST', 'Books SGST', 'Books IGST', 'Books cess', '2B taxable', '2B CGST', '2B SGST', '2B IGST', '2B cess',
             'Diff taxable (books - 2B)', 'Diff tax (books - 2B)', '2B period', '2B section', '2B ITC available', 'Books ITC status', 'Books ITC claimed FY',
             'GSTR-9 effect', 'Match code', 'Remarks', 'ID'],
            [26, 20, 17, 26, 16, 16, 11, 11] + [12] * 12 + [9, 8, 8, 16, 10, 30, 10, 30, 22],
            set('IJKLMNOPQRST'),
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
ws = wb.create_sheet('Map')
for j, h in enumerate(['Sales combo', 'Description', 'GSTR-9 key', 'Sign', '', '', 'Purchase combo', 'Table 6 key', '', 'ITC Status', 'Table 7 key', '', 'Purchase types', 'ITC types']):
    if h:
        cell(ws, f'{CL(j + 1)}1', h, f_hdr, fill_hdr, align=center)
smap = []
for st, inv, cn, dn in [('B2B', '4B', '4I', '4J'), ('B2C', '4A', '4A', '4A'), ('EXP-WP', '4C', '4I', '4J'), ('SEZ-WP', '4D', '4I', '4J'),
                        ('DEEMED', '4E', '4I', '4J'), ('EXP-WOP', '5A', '5H', '5I'), ('SEZ-WOP', '5B', '5H', '5I'), ('RCM-SUPPLY', '5C', '5H', '5I'),
                        ('ECO-9(5)', '5C1', '5C1', '5C1'), ('EXEMPT', '5D', '5H', '5I'), ('NIL', '5E', '5H', '5I'), ('NON-GST', '5F', '5H', '5I')]:
    smap.append((f'{st}|INV', f'{st} invoice', inv, 1))
    smap.append((f'{st}|CN', f'{st} credit note', cn, -1 if cn == inv else 1))
    smap.append((f'{st}|DN', f'{st} debit note', dn, 1))
smap += [('ADVANCE|ADV', 'Advance received (tax paid)', '4F', 1), ('ADVANCE|ADJ', 'Advance adjusted against invoice', '4F', -1)]
for i, (a, b, c, d) in enumerate(smap):
    for col, v in zip('ABCD', (a, b, c, d)):
        cell(ws, f'{col}{2 + i}', v, f_norm)
pmap = []
for st, base in [('B2B', '6B'), ('SEZ-SERVICES', '6B'), ('RCM-UNREG', '6C'), ('RCM-REG', '6D'), ('IMPG', '6E'), ('IMPG-SEZ', '6E')]:
    for t, s in [('Inputs', 'ip'), ('Capital Goods', 'cg'), ('Input Services', 'is')]:
        k = f'{base}-{s}'
        if base == '6E' and s == 'is':
            k = '?'
        pmap.append((f'{st}|{t}', k))
for t in ITC_TYPES:
    pmap.append((f'IMPS|{t}', '6F'))
    pmap.append((f'ISD|{t}', '6G'))
for i, (a, b) in enumerate(pmap):
    cell(ws, f'G{2 + i}', a); cell(ws, f'H{2 + i}', b)
t7 = [('Availed', ''), ('Reclaimed-Rule37/37A', ''), ('Reclaimed-Other', ''), ('Not availed', ''), ('Ineligible', ''), ('Reversed-Rule37', '7A'), ('Reversed-Rule37A', '7A1'),
      ('Reversed-Rule38', '7A2'), ('Reversed-Rule39', '7B'), ('Reversed-Rule42', '7C'), ('Reversed-Rule43', '7D'), ('Reversed-17(5)', '7E'),
      ('Reversed-Other', '7H'), ('Reversed-PY ITC (Table 12)', '12')]
for i, (a, b) in enumerate(t7):
    cell(ws, f'J{2 + i}', a); cell(ws, f'K{2 + i}', b)
for i, a in enumerate(PUR_TYPES):
    cell(ws, f'M{2 + i}', a)
for i, a in enumerate(ITC_TYPES):
    cell(ws, f'N{2 + i}', a)
for col, w in zip('ABCDEFGHIJKLMN', [16, 28, 10, 6, 2, 2, 24, 11, 2, 26, 11, 2, 14, 14]):
    ws.column_dimensions[col].width = w

# ================================================================= MASTERS
ws = wb.create_sheet('Masters')
for j, h in enumerate(['HSN code (GSTN master)', '', 'UQC description', 'UQC code', '', 'Rates (%)']):
    if h:
        cell(ws, f'{CL(j + 1)}1', h, f_hdr, fill_hdr, align=center)
for i, h in enumerate(M['hsn']):
    c = ws.cell(row=2 + i, column=1, value=h)
    c.number_format = '@'
for i, (code, desc) in enumerate(M['uqc']):
    ws.cell(row=2 + i, column=3, value=desc)
    ws.cell(row=2 + i, column=4, value=code)
for i, rt in enumerate(M['rates']):
    ws.cell(row=2 + i, column=6, value=rt)
ws.column_dimensions['A'].width = 22; ws.column_dimensions['C'].width = 26; ws.column_dimensions['D'].width = 12
ws['H1'] = 'Source: GSTR-9 Offline Tool v2.2 (GSTN) Master sheet - HSN list used for FY 2025-26 HSN validation.'
ws['H1'].font = f_sub

# ================================================================= READ ME (first sheet)
ws = wb.create_sheet('Read Me', 0)
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 3; ws.column_dimensions['B'].width = 26; ws.column_dimensions['C'].width = 100
ws['B2'] = 'GST Annual Return Workbook - GSTR-9 & GSTR-9C'; ws['B2'].font = f_title
ws['B3'] = 'Macro-free. Works in Excel 2010+, Excel for Mac and LibreOffice. Built for FY 2025-26 formats (GSTR-9 offline tool v2.2, GSTR-9C v2.9).'
ws['B3'].font = f_sub
rows = [
    ('HOW IT FLOWS', ''),
    ('1. Basic', 'Enter client name, GSTIN and FY. Previous / next FY are derived.'),
    ('2. Load data', 'The GST Annual Return Loader fills the RAW_ sheets (portal GSTR-9 JSON, GSTR-1 JSONs, GSTR-3B PDFs, GSTR-2B JSONs). They can also be pasted manually.'),
    ('3. Books', 'Paste sales and purchase registers into Books_Sales / Books_Purchase and pick Supply Type, Doc Type, ITC Type, ITC Status. The key columns map each line to a GSTR-9 row.'),
    ('4. Review rows', 'In 4-5 Outward, 6-8 ITC, 10-13 Next FY: compare Portal vs Returns vs Books. Pick the Basis (yellow, col D) and put any correction in the Adjustments block.'),
    ('5. Table 9 / 14 / 15 / 16 / HSN', 'Check tax payable vs paid; fill manual tables; review HSN tables 17/18 (invalid HSN/UQC flagged).'),
    ('6. GSTR-9C', 'Enter audited financial statement figures (yellow). GSTR-9 figures link automatically (green).'),
    ('7. Checks', 'Clear every "Check" and explain every "Review" before generating the JSON.'),
    ('7a. Reconcile', 'After pasting the registers, save the workbook and drop it into the Loader with the returns: Recon_Purchases / Recon_Suppliers (books vs 2B) and Recon_Sales / Recon_Sales_Monthly (books vs GSTR-1) are filled. Your remarks there are kept on refresh. To match invoices manually, type the same code (e.g. M1) in the Match code column of the rows that belong together, save, and drop the workbook into the Loader again.'),
    ('8. Generate JSON', 'Drop the saved workbook into the Loader -> GSTR-9 and GSTR-9C upload JSONs. Or paste "To be Filed" values into the GSTN offline tools.'),
    ('', ''),
    ('COLUMN BLOCKS', ''),
    ('Pre-filled (Portal)', 'From the auto-drafted GSTR-9 JSON downloaded from the portal (Prepare Offline > Generate JSON file to download).'),
    ('As per Returns', 'Computed from GSTR-1 / GSTR-3B / GSTR-2B raw sheets. Column AI says which return table feeds it.'),
    ('As per Books', 'Summed from Books_Sales / Books_Purchase by GSTR-9 key. Yellow when it must be typed in.'),
    ('Adjustments', 'Manual +/- corrections on top of the chosen Basis. Explain in Remarks.'),
    ('To be Filed', 'Chosen Basis + Adjustments. This is what goes into the JSON. 6A and 8A equal the portal once the portal JSON is loaded (locked); until then they show the returns figure.'),
    ('Variance', 'To be Filed minus Portal. Red when 1 rupee or more.'),
    ('', ''),
    ('COLOURS', ''),
    ('Blue text / yellow fill', 'Inputs you type or paste'),
    ('Black text', 'Formulas - do not overwrite'),
    ('Green text', 'Links to another sheet (e.g. GSTR-9 figures in 9C)'),
    ('Grey fill', 'Not applicable / locked on the portal'),
    ('', ''),
    ('NOTES', ''),
    ('Key column (AK)', 'Each GSTR-9 row has a fixed key used by the Loader. Do not edit keys or insert rows inside the table sheets.'),
    ('Indicative rows', '6A1, 8A, 8C and 13 from returns are indicative (based on 2B document dates). Books should drive the final figure.'),
    ('Table 7I', 'Sum of A to H including 7E, as per the form and the ICAI Technical Guide. GSTN offline tool v2.2 computes it excluding 7E; 7I is not uploaded.'),
    ('ICAI guidance applied', '7F = 4N - 4G + 10 - 11; 9Q = Table 9 payable + Table 14; 12E and 14S = 7J; 8B = 6B only; 9C 12C should equal GSTR-9 Table 13; reclaims under rule 37/37A go to 6H, other reclaims of previous-FY ITC go to 6A1.'),
    ('Table 9', 'Only "Tax payable" is uploaded; paid columns are auto-populated by the portal.'),
    ('GSTR-9C FY 2025-26', 'Self-certified. Part B (auditor certificate) and DSC signing inside the offline tool do not apply; sign on the portal while filing.'),
]
for i, (a, b) in enumerate(rows):
    r = 5 + i
    ws[f'B{r}'] = a; ws[f'C{r}'] = b
    ws[f'B{r}'].font = f_bold if a and (b == '' or a.isupper()) else f_bold
    ws[f'C{r}'].font = f_norm
    ws[f'C{r}'].alignment = wrap
    if a and b == '':
        ws[f'B{r}'].fill = fill_sec; ws[f'C{r}'].fill = fill_sec
for b, f in [('Pre-filled (Portal)', 'Portal'), ('As per Returns', 'Returns'), ('As per Books', 'Books'), ('Adjustments', 'Adjustment'),
             ('To be Filed', 'To be Filed'), ('Variance', 'Variance')]:
    for rr in range(5, 5 + len(rows)):
        if ws[f'B{rr}'].value == b:
            ws[f'B{rr}'].fill = fill_blk[f]
for rr in range(5, 5 + len(rows)):
    if ws[f'B{rr}'].value == 'Blue text / yellow fill':
        ws[f'B{rr}'].fill = fill_input; ws[f'B{rr}'].font = Font(name=F, size=9, bold=True, color='0000FF')

# sheet order & tab colours
order = ['Read Me', 'Basic', '4-5 Outward', '6-8 ITC', '9 & 14 Tax Paid', '10-13 Next FY', '15-16 Other', '17 HSN Outward', '18 HSN Inward',
         'GSTR-9C', 'Checks', 'Books_Sales', 'Books_Purchase', 'Recon_Purchases', 'Recon_Suppliers', 'Recon_Sales', 'Recon_Sales_Monthly', 'RAW_Portal9', 'RAW_GSTR1', 'RAW_GSTR1_HSN', 'RAW_3B', 'RAW_3B_PAY', 'RAW_2B', 'Map', 'Masters']
wb._sheets = [wb[n] for n in order]
for n in order:
    wb[n].sheet_properties.tabColor = ('2C388E' if n in ('Read Me', 'Basic') else '7030A0' if n[0].isdigit() or n == 'GSTR-9C' else
                                       'C00000' if n == 'Checks' else 'ED7D31' if n.startswith('Books') else 'FFC000' if n.startswith('Recon') else '70AD47' if n.startswith('RAW') else '808080')
for ws in wb.worksheets:
    if ws.title not in ('Read Me', 'Basic', 'Map', 'Masters'):
        ws['A1'] = "='Basic'!$C$4&\"  |  \"&'Basic'!$C$6&\"  |  FY \"&'Basic'!$C$7"
        ws['A1'].font = f_bold
# ---------------- James & James branding
FIRM = 'James & James LLP'
FIRM_FULL = 'James & James LLP, Chartered Accountants'
FIRM_ADDR = 'HB 48, 2nd Cross Rd, Panampilly Nagar, Kochi, Kerala 682036  ·  consult@jjglobal.in  ·  +91 94006 26941  ·  jjglobal.in'
brand_font = Font(name=F, size=9, bold=True, color='2C388E')
for ws in wb.worksheets:
    if ws.title not in ('Read Me', 'Basic', 'Map', 'Masters'):
        ws['A1'] = "='Basic'!$C$4&\"  |  \"&'Basic'!$C$6&\"  |  FY \"&'Basic'!$C$7&\"  |  Prepared by James & James LLP\""
    # print setup: landscape, fit to width, branded footer
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.left.text = 'Prepared by ' + FIRM_FULL
    ws.oddFooter.left.size = 8
    ws.oddFooter.center.text = 'Page &P of &N'
    ws.oddFooter.center.size = 8
    ws.oddFooter.right.text = '&A'
    ws.oddFooter.right.size = 8
    ws.oddHeader.right.text = 'GSTR-9 / GSTR-9C working papers'
    ws.oddHeader.right.size = 8
# cover branding
rm = wb['Read Me']
rm['B1'] = 'JAMES & JAMES LLP  ·  CHARTERED ACCOUNTANTS'
rm['B1'].font = Font(name=F, size=10, bold=True, color='2C388E')
rm['B3'] = 'Prepared by James & James LLP. ' + rm['B3'].value
last = rm.max_row + 2
rm[f'B{last}'] = FIRM
rm[f'B{last}'].font = brand_font
rm[f'C{last}'] = FIRM_ADDR
rm[f'C{last}'].font = Font(name=F, size=9, color='595959')
bs = wb['Basic']
bs['B1'] = 'JAMES & JAMES LLP  ·  CHARTERED ACCOUNTANTS'
bs['B1'].font = Font(name=F, size=10, bold=True, color='2C388E')
cell(bs, 'B16', 'Prepared by (firm)', f_bold, fill_sub)
cell(bs, 'C16', FIRM_FULL, brand_font)
wb.properties.creator = FIRM
wb.properties.lastModifiedBy = FIRM
wb.properties.title = 'GST Annual Return Workbook - GSTR-9 / GSTR-9C'
wb.properties.subject = 'GSTR-9 and GSTR-9C working papers'
wb.properties.keywords = 'GSTR-9, GSTR-9C, James & James LLP'
wb.properties.company = FIRM if hasattr(wb.properties, 'company') else None
# ---------------- integrity sheet (written by the Loader; used to detect edits before JSON generation)
TEMPLATE_VERSION = '2026.11.20-7'
ig = wb.create_sheet('_Integrity')
ig['A1'] = 'Item'; ig['B1'] = 'Value'; ig['C1'] = 'Rows'
ig['A2'] = 'template_version'; ig['B2'] = TEMPLATE_VERSION
ig['A3'] = 'note'; ig['B3'] = 'Maintained by the GST Annual Return Loader. Do not edit.'
ig.sheet_state = 'veryHidden'
ig.protection.sheet = True
ig.protection.password = 'JJ-GST9-integrity'
ig.protection.selectLockedCells = True
ig.protection.selectUnlockedCells = True
wb.active = 0
wb.save(OUT)
json.dump({k: v for k, v in ROWMAP.items()}, open(os.path.join(LOADER, 'src', 'rowmap.json'), 'w'), indent=0)
PATHS = {}
for sec_ in [T4, T5, T6, T7, T8, T10]:
    for row in sec_:
        if row['kind'] != 'hdr' and row.get('path'):
            PATHS[row['key']] = row['path']
json.dump(PATHS, open(os.path.join(LOADER, 'src', 'paths.json'), 'w'), indent=1)
print('saved', OUT)

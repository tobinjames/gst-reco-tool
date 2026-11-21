"""The front of every J&J workpaper template: Aptos fonts, an Index sheet (cover, contents with links, colour key,
sign-off), one Instructions sheet, and a "Back to Index" link on every sheet.
Tobin (Nov 2026): Aptos throughout, a proper index, summary sheets, instructions on one sheet, formulas a reader
without the tool can follow. The loader's own workbooks (flow, revenue) do the same in loader/src/wpfront.ts."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.hyperlink import Hyperlink

FONT = 'Aptos'
BRAND = '2C388E'
FIRM_LINE = 'JAMES & JAMES LLP  ·  CHARTERED ACCOUNTANTS'
f_band = Font(name=FONT, size=12, bold=True, color='FFFFFF')
f_bandsub = Font(name=FONT, size=10, color='DDE1F2')
f_norm = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_hdr = Font(name=FONT, size=10, bold=True, color='FFFFFF')
f_sec = Font(name=FONT, size=10, bold=True, color=BRAND)
f_link = Font(name=FONT, size=10, underline='single', color=BRAND)
f_note = Font(name=FONT, size=9, italic=True, color='737B8E')
f_input = Font(name=FONT, size=10, color='0000FF')
fill_brand = PatternFill('solid', fgColor=BRAND)
fill_sec = PatternFill('solid', fgColor='E3E6F5')
fill_input = PatternFill('solid', fgColor='FFFF99')
wrap = Alignment(wrap_text=True, vertical='top')
line = Border(bottom=Side(style='thin', color='D3D7E6'))

ROLE_NOTE = [
    ('Summary', 'Key figures. Each one is a link or a SUMIFS to a working sheet, so you can trace it back.'),
    ('Working', 'The return tables and reconciliations. Totals are SUM / SUMIFS over the source-data sheets.'),
    ('Source data', 'The rows read from the files, as they are, with the GSTR-9 key the tool gave each one.'),
]


def link(ws, ref, sheet, text=None):
    c = ws[ref]
    c.value = text or sheet
    c.hyperlink = Hyperlink(ref=ref, location=f"'{sheet}'!A1", display=c.value)
    c.font = f_link
    return c


def _sec(ws, r, text, ncol=5):
    for j in range(1, ncol + 1):
        ws.cell(row=r, column=j).fill = fill_sec
    ws.cell(row=r, column=1, value=text).font = f_sec


def add_index(wb, title, cover, sheets, version_note):
    """cover: [(label, value or formula)]; sheets: [(name, role, about)]. Inserted as the first sheet."""
    ws = wb.create_sheet('Index', 0)
    ws.sheet_view.showGridLines = False
    for c, w in zip('ABCDE', [16, 30, 13, 100, 4]):
        ws.column_dimensions[c].width = w
    for j in range(1, 6):
        ws.cell(row=1, column=j).fill = fill_brand
        ws.cell(row=2, column=j).fill = fill_brand
    ws['A1'] = FIRM_LINE; ws['A1'].font = f_band
    ws['A2'] = title; ws['A2'].font = f_bandsub
    ws.row_dimensions[1].height = 26; ws.row_dimensions[2].height = 20
    r = 4
    for lab, val in cover + [('Prepared with', version_note)]:
        ws.cell(row=r, column=1, value=lab).font = f_bold
        ws.cell(row=r, column=2, value=val).font = f_norm
        r += 1
    r += 1
    _sec(ws, r, 'Contents'); r += 1
    for j, h in enumerate(['No.', 'Sheet (click to open)', 'Type', 'What it shows / where the figures come from', '']):
        c = ws.cell(row=r, column=j + 1, value=h); c.font = f_hdr; c.fill = fill_brand
    r += 1
    for i, (name, role, about) in enumerate(sheets):
        ws.cell(row=r, column=1, value=i + 1).font = f_norm
        link(ws, f'B{r}', name)
        ws.cell(row=r, column=3, value=role).font = f_norm
        c = ws.cell(row=r, column=4, value=about); c.font = f_norm; c.alignment = wrap
        for j in range(1, 5):
            ws.cell(row=r, column=j).border = line
        r += 1
    r += 1
    _sec(ws, r, 'How to read the sheets'); r += 1
    for lab, t in ROLE_NOTE + [('Yellow cells', 'Yours to fill in (blue text). Everything else is worked out.'),
                               ('Formulas', 'Plain Excel only (SUM, SUMIFS, COUNTIFS, IF, links): no macros. Click any figure and read its formula in the formula bar; Formulas › Trace Precedents shows where it comes from.')]:
        a = ws.cell(row=r, column=1, value=lab); a.font = f_bold
        if lab == 'Yellow cells':
            a.fill = fill_input; a.font = Font(name=FONT, size=10, bold=True, color='0000FF')
        c = ws.cell(row=r, column=2, value=t); c.font = f_norm
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c.alignment = wrap
        r += 1
    r += 1
    _sec(ws, r, 'Sign-off'); r += 1
    for j, h in enumerate(['', 'Name', 'Date', 'Remarks']):
        c = ws.cell(row=r, column=j + 1, value=h or None); c.font = f_hdr; c.fill = fill_brand
    r += 1
    for lab in ('Prepared by', 'Reviewed by'):
        ws.cell(row=r, column=1, value=lab).font = f_bold
        for j in (2, 3, 4):
            c = ws.cell(row=r, column=j); c.fill = fill_input; c.font = f_input
        r += 1
    ws.sheet_properties.tabColor = BRAND
    return ws


def add_instructions(wb, title, sections, index_after=1):
    """sections: [(heading, [(label, text)])]: the tool and workpaper instructions on one sheet."""
    ws = wb.create_sheet('Instructions', index_after)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 34; ws.column_dimensions['B'].width = 120
    ws['A2'] = 'Instructions: ' + title; ws['A2'].font = Font(name=FONT, size=15, bold=True, color=BRAND)
    r = 5
    for head, lines in sections:
        _sec(ws, r, head, 2); r += 1
        for lab, t in lines:
            ws.cell(row=r, column=1, value=lab or None).font = f_bold
            c = ws.cell(row=r, column=2, value=t); c.font = f_norm; c.alignment = wrap
            r += 1
        r += 1
    ws.sheet_properties.tabColor = BRAND
    return ws


def add_back_links(wb, skip=('Index',)):
    """'‹ Back to Index' in row 1 of every other sheet, in the first free cell to the right of the heading text."""
    for ws in wb.worksheets:
        if ws.title in skip or ws.sheet_state != 'visible':
            continue
        tot, col = 0.0, 2
        for j in range(1, 60):
            w = ws.column_dimensions[CL(j)].width or 9
            tot += w
            if tot >= 75 and ws.cell(row=1, column=j + 1).value is None:
                col = j + 1
                break
        link(ws, f'{CL(col)}1', 'Index', '‹ Back to Index')


def aptos(wb):
    """every font in the workbook to Aptos (sizes, colours and weights kept)"""
    from copy import copy
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.has_style and c.font and c.font.name != FONT:
                    f = copy(c.font); f.name = FONT; c.font = f
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                if rule.dxf and rule.dxf.font is not None:
                    rule.dxf.font.name = FONT

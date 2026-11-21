"""Demo purchase register for the DEMO TRADERS company (same purchases as demo_tally_gl_dump.xlsx), so the
Input Reco tab can be tried on the same data as the GST ledger flow tab. Run from anywhere:
    python3 make_purchase_register.py
Columns as in a Tally Purchase Register export with GSTIN and Supplier Invoice No. switched on."""
import glob, json, os
from collections import OrderedDict
import openpyxl

here = os.path.dirname(os.path.abspath(__file__))
gstin = {}
for f in glob.glob(os.path.join(here, 'demo_2b_*.json')):
    for s in json.load(open(f))['data']['docdata'].get('b2b', []):
        gstin[s['trdnm']] = s['ctin']
ws = openpyxl.load_workbook(os.path.join(here, 'demo_tally_gl_dump.xlsx')).active
v = OrderedDict()
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[1] != 'Purchase':
        continue
    d = v.setdefault(r[0], {'date': r[3], 'vno': r[4], 'ref': r[5], 'party': r[6], 'tx': 0, 'IGST': 0, 'CGST': 0, 'SGST': 0, 'gross': 0})
    acc, amt = r[7], r[14] or 0
    if acc == r[6]:
        d['gross'] += amt
    elif acc.startswith('Input '):
        d[acc.split()[1]] += -amt
    else:
        d['tx'] += -amt
wb = openpyxl.Workbook()
o = wb.active
o.title = 'Purchase Register'
o.append(['DEMO TRADERS'])
o.append(['Purchase Register'])
o.append(['1-Apr-2025 to 31-Mar-2026'])
o.append([])
o.append(['Date', 'Particulars', 'Voucher Type', 'Voucher No.', 'Supplier Invoice No.', 'GSTIN/UIN', 'Gross Total', 'Taxable Value', 'IGST', 'CGST', 'SGST'])
for d in v.values():
    o.append([d['date'], d['party'], 'Purchase', d['vno'], d['ref'] or d['vno'], gstin.get(d['party'], ''), d['gross'], d['tx'], d['IGST'], d['CGST'], d['SGST']])
for row in o.iter_rows(min_row=6):
    row[0].number_format = 'dd-mmm-yy'
out = os.path.join(here, 'demo_purchase_register.xlsx')
wb.save(out)
print(out, len(v), 'invoices')

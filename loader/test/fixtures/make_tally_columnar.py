"""Synthetic Tally Prime columnar Purchase Register export (Excel), for tests.

Made up to look like Display More Reports > Purchase Register > (month) > Columnar,
with Supplier Invoice No./Date and GSTIN/UIN switched on. Not a real export: replace
with a real one when available. Run: python3 make_tally_columnar.py <out.xlsx>
"""
import sys, datetime as dt
from openpyxl import Workbook

H = ['Date', 'Particulars', 'Supplier Invoice No.', 'Supplier Invoice Date', 'GSTIN/UIN', 'Voucher Type', 'Voucher No.', 'Gross Total',
     'Purchase @ 18% (Local)', 'Purchase @ 12% (Local)', 'Purchase @ 18% (Interstate)', 'Freight Inward', 'Packing Charges', 'Professional Charges',
     'Input CGST @ 9%', 'Input SGST @ 9%', 'Input CGST @ 6%', 'Input SGST @ 6%', 'Input IGST @ 18%', 'Input Cess', 'Round Off', 'TDS Payable']
C = {h: i for i, h in enumerate(H)}
# (date, party, inv no, inv date (text, as Tally exports it), gstin, vno, {ledger: amount})
V = [
    (dt.date(2026, 10, 3), 'Sharma Steel Traders', 'SST/26-27/1042', '1-Oct-26', '27AABCS1234F1Z5', '1', {'Purchase @ 18% (Local)': 100000, 'Freight Inward': 2500, 'Input CGST @ 9%': 9225, 'Input SGST @ 9%': 9225}),
    (dt.date(2026, 10, 5), 'Kaveri Packaging Pvt Ltd', 'KP-889', '4-Oct-26', '27AAFCK5678M1Z2', '2', {'Purchase @ 12% (Local)': 40000, 'Purchase @ 18% (Local)': 10000, 'Packing Charges': 1000,
        'Input CGST @ 6%': 2400, 'Input SGST @ 6%': 2400, 'Input CGST @ 9%': 990, 'Input SGST @ 9%': 990, 'Round Off': -0.0}),
    (dt.date(2026, 10, 9), 'Gujarat Polymers LLP', 'GP/2026/0567', '7-Oct-26', '24AAKFG2468H1ZQ', '3', {'Purchase @ 18% (Interstate)': 75550.40, 'Input IGST @ 18%': 13599.07, 'Round Off': 0.53}),
    (dt.date(2026, 10, 14), 'Mehta & Associates', 'MA/112', '12-Oct-26', '27AAEFM9753K1Z8', '4', {'Professional Charges': 50000, 'Input CGST @ 9%': 4500, 'Input SGST @ 9%': 4500, 'TDS Payable': -5000}),
    (dt.date(2026, 10, 20), 'Ramesh Hardware', '455', '20-Oct-26', '', '5', {'Purchase @ 18% (Local)': 3200}),
    (dt.date(2026, 10, 24), 'Deccan Tobacco Co', 'DTC-7781', '22-Oct-26', '29AACCD1357P1ZR', '6', {'Purchase @ 18% (Interstate)': 20000, 'Input IGST @ 18%': 3600, 'Input Cess': 1000}),
]
wb = Workbook(); ws = wb.active; ws.title = 'Purchase Register'
for r, t in enumerate(['Demo Manufacturing Pvt Ltd', 'Plot 12, MIDC Bhosari, Pune', 'GSTIN: 27AAACD0000A1Z9', 'Purchase Register', '1-Oct-2026 to 31-Oct-2026'], 1):
    ws.cell(r, 1, t)
HR = 7
for j, h in enumerate(H, 1): ws.cell(HR, j, h)
r = HR
tot = [0.0] * len(H)
for d, party, no, idate, gstin, vno, led in V:
    r += 1
    row = [None] * len(H)
    row[C['Date']] = d; row[C['Particulars']] = party; row[C['Supplier Invoice No.']] = no; row[C['Supplier Invoice Date']] = idate
    row[C['GSTIN/UIN']] = gstin or None; row[C['Voucher Type']] = 'Purchase'; row[C['Voucher No.']] = vno
    for k, v in led.items(): row[C[k]] = v
    row[C['Gross Total']] = round(sum(v for v in led.values()), 2)
    for j, v in enumerate(row, 1):
        ws.cell(r, j, v)
        if isinstance(v, dt.date): ws.cell(r, j).number_format = 'd-mmm-yy'
    for j in range(C['Gross Total'], len(H)): tot[j] += row[j] or 0
r += 1
ws.cell(r, C['Particulars'] + 1, 'Grand Total')
for j in range(C['Gross Total'], len(H)): ws.cell(r, j + 1, round(tot[j], 2))
wb.save(sys.argv[1])

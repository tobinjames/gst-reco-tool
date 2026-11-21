"""Synthetic revenue-reconciliation fixtures for one made-up company, FY 2025-26: a sales register (.xlsx), the same sales
as a Tally voucher dump (plus interest income), GSTR-1 JSON per month (Apr-25 .. Apr-26) and GSTR-3B Table 3.1 figures
(demo_rev_3b.json: what the 3B PDFs would give; the tests turn it into ThreeB objects).

Planted differences (GSTR-1 - books, taxable / tax):
  P1 Nov-25: SAL/1103 Kerala 40,000 in books, not in GSTR-1                         -40,000 / -7,200  not reported
  P2 Mar-26: SAL/0303 Chennai 30,000 reported in the Apr-26 GSTR-1 (next year)      -30,000 / -5,400  next FY
  P3 Apr-25 GSTR-1 has PY/0999 dated 28-03-2025 (last year's invoice)               +25,000 / +4,500  previous FY
  P4 Jul-25: SAL/0703 booked 28-07-2025, reported in the Aug-25 GSTR-1              timing: not in the year statement
  P5 Sep-25: SAL/0901 books 100,000, GSTR-1 99,000                                   -1,000 / -180     value
  P6 Oct-25: credit note CN/1001 Kerala 10,000 in books, not in GSTR-1              +10,000 / +1,800  not reported
  P7 B2C: Dec-25 B2CS 18,000 vs books 20,000; Jan-26 cash sales reported in Feb B2CS -2,000 / -360    B2C (Jan->Feb nets)
  P8 Feb-26: X/77 Kerala 15,000 in GSTR-1, not in books                             +15,000 / +2,700  not in books
  P9 Mar-26: exempt sale EX/0301 3,000 in books, not in Table 8 (Jun 5,000 agrees)   -3,000            nil / exempt
  GL dump only: Jun-25 interest received 2,000 (suggested "not a supply")
3B: Aug-25 3.1(a) short by 10,000 / 1,800, made good in Sep-25 (timing); Mar-26 3.1(a) short by 20,000 / 3,600.
Run: python3 make_revenue_fixtures.py  (writes the files next to this script)
"""
import datetime as dt, json, os
from openpyxl import Workbook

OUT = os.path.dirname(os.path.abspath(__file__))
GSTIN = '32AAACJ1234K1Z5'
K, KG = 'Kerala Customer Pvt Ltd', '32AAACK1111A1Z1'
C, CG = 'Chennai Customer Pvt Ltd', '33AAACC2222B1Z2'
months = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in range(1, 4)]

# books documents: (date, no, party, gstin, type, doctype, taxable, rate, inter)
books = []
for (y, m) in months:
    d15 = dt.date(y, m, 15)
    books.append((d15, f'SAL/{m:02d}01', K, KG, 'B2B', 'Invoice', 100000, 18, False))
    books.append((d15, f'SAL/{m:02d}02', C, CG, 'B2B', 'Invoice', 50000, 18, True))
    books.append((dt.date(y, m, 28), f'CS/{m:02d}{y % 100}', 'Cash Sales', '', 'B2C', 'Invoice', 20000, 18, False))
books.append((dt.date(2025, 11, 20), 'SAL/1103', K, KG, 'B2B', 'Invoice', 40000, 18, False))  # P1
books.append((dt.date(2026, 3, 30), 'SAL/0303', C, CG, 'B2B', 'Invoice', 30000, 18, True))  # P2
books.append((dt.date(2025, 7, 28), 'SAL/0703', K, KG, 'B2B', 'Invoice', 10000, 18, False))  # P4
books.append((dt.date(2025, 10, 25), 'CN/1001', K, KG, 'B2B', 'Credit Note', 10000, 18, False))  # P6
books.append((dt.date(2025, 6, 10), 'EX/0601', K, KG, 'EXEMPT', 'Invoice', 5000, 0, False))
books.append((dt.date(2026, 3, 12), 'EX/0301', K, KG, 'EXEMPT', 'Invoice', 3000, 0, False))  # P9

def tax(tx, rate, inter):
    t = round(tx * rate / 100, 2)
    return (t, 0, 0) if inter else (0, round(t / 2, 2), round(t / 2, 2))

# ---- sales register
wb = Workbook(); ws = wb.active; ws.title = 'Sales'
ws.append(['Demo Traders Private Limited'])
ws.append(['Sales Register 1-Apr-2025 to 31-Mar-2026'])
ws.append(['Invoice Date', 'Invoice Number', 'Customer Name', 'GSTIN', 'Place of Supply', 'Supply Type', 'Document Type', 'Taxable Value', 'CGST', 'SGST', 'IGST'])
for (d, no, p, g, ty, dtp, tx, rt, inter) in books:
    i, c, s = tax(tx, rt, inter)
    ws.append([dt.datetime(d.year, d.month, d.day), no, p, g, '33' if inter else '32', ty, dtp, tx, c, s, i])
ws.append(['Total', None, None, None, None, None, None, sum(b[6] for b in books)])
wb.save(os.path.join(OUT, 'demo_rev_sales_register.xlsx'))

# ---- Tally voucher dump (Dr negative): sales ledgers by rate/type, output GST, interest income
wb = Workbook(); ws = wb.active; ws.title = 'Sheet1'
ws.append(['voucher_id', 'vch_type', 'vch_key', 'date', 'vouchernumber', 'reference', 'partyname', 'gl_account', 'stockitemname', 'rate', 'actualqty', 'billedqty', 'cost_centre', 'cost_centre_global', 'amount', 'narration'])
n = [0]
def vch(vt, d, no, party, lines, narr=None):
    n[0] += 1; vid = f'rev-{n[0]:05d}'
    for led, a in lines:
        ws.append([vid, vt, vid + ':1', dt.datetime(d.year, d.month, d.day), no, None, party, led, None, None, None, None, None, None, a, narr])
for (d, no, p, g, ty, dtp, tx, rt, inter) in books:
    i, c, s = tax(tx, rt, inter)
    sg = -1 if dtp == 'Credit Note' else 1
    led = 'Exempt Sales' if ty == 'EXEMPT' else 'Interstate Sales 18%' if inter else 'Local Sales 18%'
    lines = [(p, -sg * round(tx + i + c + s, 2)), (led, sg * tx)]
    if i: lines.append(('Output IGST', sg * i))
    if c: lines += [('Output CGST 9%', sg * c), ('Output SGST 9%', sg * s)]
    vch('Credit Note' if dtp == 'Credit Note' else 'Sales', d, no, p, lines)
vch('Receipt', dt.date(2025, 6, 30), 'RCT/0001', 'HDFC Bank', [('HDFC Bank', -2000), ('Interest Received', 2000)], 'FD interest')
vch('Purchase Order', dt.date(2025, 6, 5), 'PO/1', 'Supplier', [('Supplier', 1180), ('Local Sales 18%', -1000), ('Output CGST 9%', -90), ('Output SGST 9%', -90)])
wb.save(os.path.join(OUT, 'demo_rev_gl_dump.xlsx'))

# ---- customer list (Zoho "Export Customers" style): just names and GSTINs, plus columns the reader must not mind
wb = Workbook(); ws = wb.active; ws.title = 'Contacts'
ws.append(['Display Name', 'Company Name', 'GST Treatment', 'GST Identification Number (GSTIN)', 'Credit Limit', 'Opening Balance'])
seen = {}
for b in books:
    if b[3] and b[2] not in seen: seen[b[2]] = b[3]
for p, g in seen.items(): ws.append([p, p, 'business_gst', g, 500000, 0])
ws.append(['Cash Sales', None, 'consumer', None, 0, 0])
wb.save(os.path.join(OUT, 'demo_rev_customers.xlsx'))

# ---- GSTR-1: what was reported, per return month
def fp(y, m): return f'{m:02d}{y}'
g1 = {}  # fp -> {'b2b': {ctin: [inv]}, 'b2cs': [...], 'cdnr': ..., 'nil': ...}
def per(y, m): return g1.setdefault(fp(y, m), {'b2b': {}, 'b2cs': [], 'nil': []})
def b2b(y, m, ctin, no, d, tx, inter):
    i, c, s = tax(tx, 18, inter)
    per(y, m)['b2b'].setdefault(ctin, []).append({'inum': no, 'idt': d.strftime('%d-%m-%Y'), 'val': round(tx + i + c + s, 2), 'pos': ctin[:2], 'rchrg': 'N', 'inv_typ': 'R',
        'itms': [{'num': 1801, 'itm_det': {'rt': 18, 'txval': tx, 'iamt': i, 'camt': c, 'samt': s, 'csamt': 0}}]})
for (y, m) in months:
    d15 = dt.date(y, m, 15)
    b2b(y, m, KG, f'SAL/{m:02d}01', d15, 99000 if (y, m) == (2025, 9) else 100000, False)  # P5
    b2b(y, m, CG, f'SAL/{m:02d}02', d15, 50000, True)
    b2c = {(2025, 12): 18000, (2026, 1): 0, (2026, 2): 40000}.get((y, m), 20000)  # P7
    if b2c: per(y, m)['b2cs'].append({'sply_ty': 'INTRA', 'pos': '32', 'typ': 'OE', 'rt': 18, 'txval': b2c, 'iamt': 0, 'camt': round(b2c * 0.09, 2), 'samt': round(b2c * 0.09, 2), 'csamt': 0})
b2b(2025, 4, KG, 'PY/0999', dt.date(2025, 3, 28), 25000, False)  # P3
b2b(2025, 8, KG, 'SAL/0703', dt.date(2025, 7, 28), 10000, False)  # P4
b2b(2026, 2, KG, 'X/77', dt.date(2026, 2, 10), 15000, False)  # P8
b2b(2026, 4, CG, 'SAL/0303', dt.date(2026, 3, 30), 30000, True)  # P2 (next FY return)
b2b(2026, 4, KG, 'SAL/0401', dt.date(2026, 4, 15), 100000, False)
per(2025, 6)['nil'].append({'sply_ty': 'INTRAB2B', 'nil_amt': 0, 'expt_amt': 5000, 'ngsup_amt': 0})
for k, v in g1.items():
    j = {'gstin': GSTIN, 'fp': k, 'b2b': [{'ctin': c, 'inv': l} for c, l in v['b2b'].items()]}
    if v['b2cs']: j['b2cs'] = v['b2cs']
    if v['nil']: j['nil'] = {'inv': v['nil']}
    with open(os.path.join(OUT, f'demo_rev_gstr1_{k}.json'), 'w') as f: json.dump(j, f, indent=1)

# ---- GSTR-3B Table 3.1 per month (FY): GSTR-1 of the month, then the planted 3B differences
tb = {}
for k, v in g1.items():
    y, m = int(k[2:]), int(k[:2])
    if (y, m) not in months: continue
    a = [0, 0, 0, 0, 0]  # taxable, igst, cgst, sgst, cess (3B PDF order)
    for c, l in v['b2b'].items():
        for inv in l:
            it = inv['itms'][0]['itm_det']; a = [a[0] + it['txval'], a[1] + it['iamt'], a[2] + it['camt'], a[3] + it['samt'], 0]
    for r in v['b2cs']: a = [a[0] + r['txval'], a[1], a[2] + r['camt'], a[3] + r['samt'], 0]
    adj = {(2025, 8): -10000, (2025, 9): 10000, (2026, 3): -20000}.get((y, m), 0)
    if adj: a = [a[0] + adj, a[1], a[2] + adj * 0.09, a[3] + adj * 0.09, 0]
    rows = {'3.1a': [round(x, 2) for x in a], '3.1b': [0, 0, 0, 0, 0], '3.1c': [sum(x['expt_amt'] + x['nil_amt'] for x in v['nil']), 0, 0, 0, 0], '3.1e': [0, 0, 0, 0, 0], '3.1.1ii': [0, 0, 0, 0, 0]}
    tb[k] = rows
with open(os.path.join(OUT, 'demo_rev_3b.json'), 'w') as f: json.dump(tb, f, indent=1)
print('written to', OUT)

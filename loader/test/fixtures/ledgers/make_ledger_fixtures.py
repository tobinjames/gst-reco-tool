"""Synthetic GST ledger-flow fixtures for one made-up company, FY 2025-26, in the layouts of the real downloads:
portal Electronic Cash / Credit Ledger and Liability Register (CSV, Nov 2026 layout) and a Tally voucher dump (xlsx).

Planted differences (what the reasons engine must find):
  1. Opening: books Input CGST 10,000 vs credit ledger 9,500.
  2. Jun-25: ITC of 1,000 CGST + 1,000 SGST reversed by journal in books; the credit ledger shows ITC net of it.
  3. Jul-25 / Aug-25: an invoice with 1,000 CGST + 1,000 SGST booked in Jul, claimed in the Aug return (timing).
  4. Sep-25: a local sale booked with IGST 1,800 instead of CGST 900 + SGST 900 (tax head).
  5. Oct-25: return filed late, late fee 50 + 50 paid in cash; books debited it to Output CGST / SGST.
  6. Nov-25: a sale with CGST 300 + SGST 300 in books, not reported (unexplained).
  7. Dec-25: IGST credit used against CGST on the portal; the books set-off journal credits Input CGST (cross-utilisation).
  8. Mar-26: an invoice with IGST 5,000 booked in Mar, claimed in the Apr-26 return (after the period).
  9. A Purchase Order with GST lines (non-accounting, must be ignored).
  10. Mar-26 cash deposit made on 20-Apr-26: in books after the period, on the portal for the Mar-26 return.
Run: python3 make_ledger_fixtures.py  (writes the files next to this script)
"""
import csv, datetime as dt, itertools, os
_seq = itertools.count(1234567)
def rid(k): return next(_seq) % 10**k
from openpyxl import Workbook

OUT = os.path.dirname(os.path.abspath(__file__))
GSTIN, NAME, ST = '32AAACJ1234K1Z5', 'DEMO TRADERS PRIVATE LIMITED', '32'
H = ['IGST', 'CGST', 'SGST', 'CESS']
months = [(2025, m) for m in range(4, 13)] + [(2026, m) for m in range(1, 5)]  # Apr-25 .. Apr-26 (Apr-26 only on the portal)

def nxt(y, m): return (y + 1, 1) if m == 12 else (y, m + 1)
def lbl(y, m): return dt.date(y, m, 1).strftime('%b-%y')
def ds(d): return d.strftime('%d/%m/%Y')

# ---- true (portal) figures per return period
liab, itc = {}, {}
for (y, m) in months:
    liab[(y, m)] = {'IGST': 9000, 'CGST': 9000, 'SGST': 9000, 'CESS': 0}
    itc[(y, m)] = {'IGST': 7200, 'CGST': 5400, 'SGST': 5400, 'CESS': 0}
liab[(2025, 9)].update(CGST=9900, SGST=9900)                       # 4. correct heads on the portal
itc[(2025, 6)].update(CGST=4400, SGST=4400)                        # 2. net of reversal
itc[(2025, 8)].update(CGST=6400, SGST=6400)                        # 3. claimed in Aug
itc[(2026, 4)].update(IGST=12200)                                  # 8. claimed in Apr-26
itc[(2025, 12)].update(IGST=12000)                                 # 7. enough IGST credit to pay part of CGST
fee = {(2025, 10): {'CGST': 50, 'SGST': 50}}                       # 5.

# ---- portal: set-off simulation (IGST credit first: IGST, CGST, SGST; CGST credit: CGST, IGST; SGST credit: SGST, IGST)
bal = {'IGST': 0, 'CGST': 9500, 'SGST': 9500, 'CESS': 0}
credit_rows, liab_rows, cash_rows, util = [], [], [], {}
cash_bal = {h: {'tax': 0, 'fee': 0} for h in H}
for (y, m) in months:
    fy_, fm = nxt(y, m)
    fdate = dt.date(fy_, fm, 25 if (y, m) == (2025, 10) else 20)
    arn = f'AA{ST}{m:02d}{y % 100:02d}{rid(7):07d}X'
    for h in H: bal[h] += itc[(y, m)][h]
    credit_rows.append([fdate, arn, lbl(y, m), 'ITC accrued through - Inputs', 'Credit', dict(itc[(y, m)]), dict(bal)])
    need = dict(liab[(y, m)]); used = []  # (credit head, liability head, amount)
    for ch, order in [('IGST', ['IGST', 'CGST', 'SGST']), ('CGST', ['CGST', 'IGST']), ('SGST', ['SGST', 'IGST'])]:
        for lh in order:
            x = min(bal[ch], need[lh])
            if x > 0: used.append((ch, lh, x)); bal[ch] -= x; need[lh] -= x
    util[(y, m)] = used
    dref = f'DI{ST}{fm:02d}{fy_ % 100:02d}{rid(7):07d}'
    debit = {h: sum(x for ch, lh, x in used if ch == h) for h in H}
    credit_rows.append([fdate, dref, lbl(y, m), 'Other than reverse charge', 'Debit', debit, dict(bal)])
    if y == 2026 and m == 4: continue  # Apr-26: credit ledger only
    f = fee.get((y, m), {})
    lb = {h: {'tax': liab[(y, m)][h], 'fee': f.get(h, 0)} for h in H}
    liab_rows.append([fdate, arn, '-', 'Other than reverse charge', 'Debit', {h: {'tax': liab[(y, m)][h], 'fee': f.get(h, 0)} for h in H}, {h: dict(v) for h, v in lb.items()}])
    liab_rows.append([fdate, arn, '-', 'Reverse charge and supplies made u/s 9(5)', 'Debit', {h: {'tax': 0, 'fee': 0} for h in H}, {h: dict(v) for h, v in lb.items()}])
    cash_need = {h: {'tax': need[h], 'fee': f.get(h, 0)} for h in H}
    cref = f'DC{ST}{fm:02d}{fy_ % 100:02d}{rid(7):07d}'
    if any(v['tax'] or v['fee'] for v in cash_need.values()):
        # deposit, then debit from the cash ledger, then discharge in the liability register
        for h in H:
            for c in ('tax', 'fee'): cash_bal[h][c] += cash_need[h][c]
        cash_rows.append([fdate, f'{fy_}{fm:02d}20{rid(8):08d}', lbl(y, m), 'Deposit', 'Credit', cash_need, {h: dict(v) for h, v in cash_bal.items()}])
        for h in H:
            for c in ('tax', 'fee'): cash_bal[h][c] -= cash_need[h][c]
        cash_rows.append([fdate, cref, lbl(y, m), 'Other than reverse charge', 'Debit', cash_need, {h: dict(v) for h, v in cash_bal.items()}])
        for h in H:
            for c in ('tax', 'fee'): lb[h][c] -= cash_need[h][c]
        liab_rows.append([fdate, cref, 'Cash', 'Other than reverse charge', 'Credit', cash_need, {h: dict(v) for h, v in lb.items()}])
    paid_itc = {h: sum(x for ch, lh, x in used if lh == h) for h in H}
    for h in H: lb[h]['tax'] -= paid_itc[h]
    liab_rows.append([fdate, dref, 'Credit', 'Other than reverse charge', 'Credit', {h: {'tax': paid_itc[h], 'fee': 0} for h in H}, {h: dict(v) for h, v in lb.items()}])

COMP = ['Tax', 'Interest', 'Penalty', 'Fee', 'Others', 'Total']
def comps(d): t, fe = d.get('tax', 0), d.get('fee', 0); return [t, 0, 0, fe, 0, t + fe]
def head_block(title):
    g = []
    for h in ['Integrated Tax', 'Central Tax', 'State Tax', 'CESS']: g += [f'{h} {title}'] + [''] * 5
    return g

def write_csv(name, rows):
    with open(os.path.join(OUT, name), 'w', newline='') as fh:
        w = csv.writer(fh, quoting=csv.QUOTE_MINIMAL)
        for r in rows: w.writerow(r)

# liability register
rows = [['Electronic Liability Register'], [], [''] * 6 + ['GSTIN', GSTIN, ''], [''] * 6 + ['Legal Name', NAME, ''], [''] * 6 + ['TAX PERIOD', 'Apr-25 to Mar-26', ''], []]
rows.append(['Sr.No', 'Date', 'Reference No.', 'Ledger Used For Discharging Liability ', 'Description', 'Transaction Type (Debit/Credit)'] + head_block('Amount Debited / Credited') + head_block('Balance'))
rows.append([''] * 6 + COMP * 8)
for i, (d, ref, used, desc, typ, amt, b) in enumerate(liab_rows, 1):
    rows.append([str(i), ds(d), ref, used, desc, typ] + sum((comps(amt[h]) for h in H), []) + sum((comps(b[h]) for h in H), []) + [''])
rows.append(['-'] * 5 + ['Closing Balance'] + ['-'] * 24 + [0] * 24)
rows[-1][4] = 'Closing Balance'; rows[-1][5] = '-'
write_csv('demo_liability_register.csv', rows)

# credit ledger (two-level header: Credit / Debit and Balance Available, per head)
rows = [['Electronic Credit Ledger'], [], ['', '', '', 'GSTIN', GSTIN, ''], ['', '', '', 'Legal Name', NAME, ''], ['', '', '', 'From', '01/04/2025', ''], ['', '', '', 'To', '31/05/2026']]
rows.append(['Sr.No', 'Date', 'Reference No.', 'Tax Period if any ', 'Description ', 'Transaction Type[Debit (DR) / Credit (CR)]', '', '', 'Credit / Debit', '', '', '', 'Balance Available'])
rows.append([''] * 6 + ['Integrated Tax', 'Central Tax', 'State Tax', 'CESS', 'Total'] * 2)
ob = {'IGST': 0, 'CGST': 9500, 'SGST': 9500, 'CESS': 0}
rows.append([' -', '-', '-', '-', 'Opening Balance', '-', '-', '-', '-', '-', '-'] + [ob[h] for h in H] + [sum(ob.values())] + [''])
for i, (d, ref, per, desc, typ, amt, b) in enumerate(credit_rows, 1):
    rows.append([f' {i}', ds(d), ref, per, desc, typ] + [amt[h] for h in H] + [sum(amt.values())] + [b[h] for h in H] + [sum(b.values())] + [''])
rows.append([' -', '-', '-', '-', 'Closing Balance', '-', '-', '-', '-', '-', '-'] + [bal[h] for h in H] + [sum(bal.values())] + [''])
write_csv('demo_credit_ledger.csv', rows)

# cash ledger
rows = [['Electronic Cash Ledger'], [], [''] * 6 + ['GSTIN', GSTIN, ''], [''] * 6 + ['Name(Legal)', NAME, ''], [''] * 6 + ['From', '01/04/2025', ''], [''] * 6 + ['To', '31/05/2026']]
rows.append(['Sr.No', 'Date  of deposit/Debit ', 'Time of deposit', 'Reporting date (by bank)', 'Reference No.', 'Tax Period if applicable', 'Description', 'Transaction Type (Debit/Credit)'] + head_block('Amount Debited / Credited') + head_block('Balance'))
rows.append([''] * 8 + COMP * 8)
z = {h: {} for h in H}
rows.append(['1', '-', '-', '-', '-', '-', 'Opening Balance', '-'] + ['-'] * 24 + sum((comps(z[h]) for h in H), []) + [''])
for i, (d, ref, per, desc, typ, amt, b) in enumerate(cash_rows, 2):
    rows.append([str(i), ds(d), '12:00:00' if typ == 'Credit' else '-', ds(d) if typ == 'Credit' else '-', ref, per, desc, typ] + sum((comps(amt[h]) for h in H), []) + sum((comps(b[h]) for h in H), []) + [''])
rows.append([str(len(cash_rows) + 2), '-', '-', '-', '-', '-', 'Closing Balance', '-'] + ['-'] * 24 + sum((comps(cash_bal[h]) for h in H), []) + [''])
write_csv('demo_cash_ledger.csv', rows)

# ---- books: Tally voucher dump
wb = Workbook(); ws = wb.active; ws.title = 'Sheet1'
ws.append(['voucher_id', 'vch_type', 'vch_key', 'date', 'vouchernumber', 'reference', 'partyname', 'gl_account', 'stockitemname', 'rate', 'actualqty', 'billedqty', 'cost_centre', 'cost_centre_global', 'amount', 'narration'])
n = [0]
def vch(vt, d, party, lines, narr=None):
    """lines: (ledger, amount) with Dr negative, as Tally exports"""
    n[0] += 1; vid = f'demo-{n[0]:05d}'
    for led, a in lines: ws.append([vid, vt, vid + ':1', dt.datetime(d.year, d.month, d.day), f'{vt[:3].upper()}/{n[0]:04d}', None, party, led, None, None, None, None, None, None, a, narr])
    ws.append([None] * 16)
IN = {'IGST': 'Input IGST', 'CGST': 'Input CGST', 'SGST': 'Input SGST'}
OUTL = {'IGST': 'Output IGST', 'CGST': 'Output CGST 9%', 'SGST': 'Output SGST 9%'}
for (y, m) in months[:-1]:  # Apr-25 .. Mar-26
    d15, eom = dt.date(y, m, 15), (dt.date(*nxt(y, m), 1) - dt.timedelta(days=1))
    # sales
    if (y, m) == (2025, 9):
        vch('Sales', d15, 'Kerala Customer Pvt Ltd', [('Kerala Customer Pvt Ltd', -(100000 + 18000)), ('Local Sales 18%', 100000 - 10000), ('Local Sales 18%', 10000), (OUTL['CGST'], 9000), (OUTL['SGST'], 9000)])
        vch('Sales', d15, 'Kochi Buyer LLP', [('Kochi Buyer LLP', -11800), ('Local Sales 18%', 10000), (OUTL['IGST'], 1800)])  # 4. wrong head
    else:
        vch('Sales', d15, 'Kerala Customer Pvt Ltd', [('Kerala Customer Pvt Ltd', -118000), ('Local Sales 18%', 100000), (OUTL['CGST'], 9000), (OUTL['SGST'], 9000)])
    vch('Sales', d15, 'Chennai Customer Pvt Ltd', [('Chennai Customer Pvt Ltd', -59000), ('Interstate Sales 18%', 50000), (OUTL['IGST'], 9000)])
    if (y, m) == (2025, 11):
        vch('Sales', d15 + dt.timedelta(days=5), 'Walk-in Customer', [('Walk-in Customer', -3933.33), ('Local Sales 18%', 3333.33), (OUTL['CGST'], 300), (OUTL['SGST'], 300)])  # 6.
    # purchases
    c = 5400 + (1000 if (y, m) == (2025, 7) else 0)
    vch('Purchase', d15, 'Kerala Supplier Pvt Ltd', [('Kerala Supplier Pvt Ltd', round(c / 0.09 + 2 * c, 2)), ('Local Purchases 18%', -round(c / 0.09, 2)), (IN['CGST'], -c), (IN['SGST'], -c)])
    if (y, m) == (2025, 8): pass  # the Aug return claimed the Jul invoice: books already have it in Jul
    ig = (12000 if (y, m) == (2025, 12) else 7200) + (5000 if (y, m) == (2026, 3) else 0)
    vch('Purchase', d15, 'Mumbai Supplier Pvt Ltd', [('Mumbai Supplier Pvt Ltd', round(ig / 0.18 + ig, 2)), ('Interstate Purchases 18%', -round(ig / 0.18, 2)), (IN['IGST'], -ig)])
    if (y, m) == (2025, 6):
        vch('Journal', eom, '', [('Staff Welfare Expenses', -2000), (IN['CGST'], 1000), (IN['SGST'], 1000)], 'ITC reversed: blocked credit u/s 17(5)')  # 2.
    if (y, m) == (2025, 5):
        vch('Purchase Order', d15, 'Kerala Supplier Pvt Ltd', [('Kerala Supplier Pvt Ltd', 11800), ('Local Purchases 18%', -10000), (IN['CGST'], -900), (IN['SGST'], -900)])  # 9.
    # set-off journal mirrors the portal utilisation
    lines = []
    for ch, lh, x in util[(y, m)]:
        inh = 'CGST' if (y, m) == (2025, 12) and ch == 'IGST' and lh == 'CGST' else ch  # 7.
        lines += [(OUTL[lh], -x), (IN[inh], x)]
    vch('Journal', eom, '', lines, f'GST set-off {lbl(y, m)}')
    # cash: deposit on the filing date (Dr cash ledger, Cr bank), used by journal at month end
    f = fee.get((y, m), {})
    paid = {h: liab[(y, m)][h] - sum(x for ch, lh, x in util[(y, m)] if lh == h) for h in ['IGST', 'CGST', 'SGST']}
    tot = sum(paid.values()) + sum(f.values())
    if tot:
        fd = dt.date(*nxt(y, m), 25 if (y, m) == (2025, 10) else 20)
        vch('Payment', fd, 'HDFC Bank', [('GST Electronic Cash Ledger', -tot), ('HDFC Bank', tot)], f'GST challan {lbl(y, m)}')
        vch('Journal', eom, '', [(OUTL[h], -(paid[h] + f.get(h, 0))) for h in paid if paid[h] + f.get(h, 0)] + [('GST Electronic Cash Ledger', tot)], f'GST paid in cash {lbl(y, m)}')  # 5. fee debited to output
wb.save(os.path.join(OUT, 'demo_tally_gl_dump.xlsx'))
print('written to', OUT)

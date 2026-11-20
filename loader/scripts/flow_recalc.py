"""Recalculate the ledger-flow workbook in LibreOffice and compare every formula with the value the Loader cached.
Usage: python3 scripts/flow_recalc.py <workbook.xlsx>   (needs soffice and openpyxl)
The cached values are removed first, so LibreOffice has to compute every formula itself."""
import os, re, shutil, subprocess, sys, tempfile, zipfile
from openpyxl import load_workbook

src = sys.argv[1]
tmp = tempfile.mkdtemp()
stripped = os.path.join(tmp, 'stripped.xlsx')
with zipfile.ZipFile(src) as zi, zipfile.ZipFile(stripped, 'w', zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        data = zi.read(it.filename)
        if it.filename.startswith('xl/worksheets/'):
            data = re.sub(rb'(</f>)<v>[^<]*</v>', rb'\1', data).replace(b' t="str"><f>', b'><f>')
        zo.writestr(it, data)
out = os.path.join(tmp, 'out')
subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', out, stripped], check=True, capture_output=True, timeout=180)
mine = load_workbook(src, data_only=True)
lo = load_workbook(os.path.join(out, 'stripped.xlsx'), data_only=True)
form = load_workbook(src, data_only=False)
bad = n = 0
for ws in form.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                n += 1
                a, b = mine[ws.title][c.coordinate].value, lo[ws.title][c.coordinate].value
                if isinstance(a, (int, float)) or isinstance(b, (int, float)):
                    if abs((a or 0) - (b or 0)) > 0.005:
                        bad += 1
                        if bad <= 15: print('MISMATCH', ws.title, c.coordinate, c.value[:90], 'loader', a, 'libreoffice', b)
                elif (a or '') != (b or ''):
                    bad += 1
                    if bad <= 15: print('MISMATCH', ws.title, c.coordinate, c.value[:90], 'loader', a, 'libreoffice', b)
shutil.rmtree(tmp)
print(f'{n} formulas, {bad} mismatches')
sys.exit(1 if bad else 0)

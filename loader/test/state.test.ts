// Workpaper memory (src/state.ts): the hidden _JJState / _JJFiles sheets carry the settings, choices and source files,
// survive a second build and an Excel-style re-save, and classify() reads them back from the dropped workbook.
// Run: npx tsx test/state.test.ts (or npm test for all)
import { readFileSync, readdirSync, writeFileSync, mkdtempSync, existsSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { spawnSync } from 'child_process';
import assert from 'assert/strict';
import { classify } from '../src/build';
import { Workbook, Reader } from '../src/xlsx';
import { writeBook } from '../src/xlsxbook';
import { store } from '../src/util';
import {
  FILES_SHEET,
  STATE_SHEET,
  addStateSheets,
  applyState,
  describeKeys,
  packState,
  readSavedFiles,
  setStateProvider,
  undoState,
} from '../src/state';

const noPdf = async () => {
  throw new Error('no pdf here');
};
const rev = new URL('./fixtures/revenue/', import.meta.url);
const fx = (f: string) => new Uint8Array(readFileSync(new URL(f, rev)));
const G = '32AAACJ1234K1Z5',
  FY = '2025-26';

// ---- source files: every GSTR-1 of the demo (several chunks of base64 once zipped) and the sales register
const srcNames = readdirSync(rev).filter(f => f.startsWith('demo_rev_gstr1_') || f === 'demo_rev_sales_register.xlsx');
const src = srcNames.map(n => ({ name: n, kind: n.endsWith('.json') ? 'GSTR-1' : 'Register', bytes: fx(n) }));
// a big file too, so the zip spans many cells
const big = new TextEncoder().encode(JSON.stringify({ pad: Array.from({ length: 30000 }, (_, i) => `${i}-${Math.random()}`) }));
src.push({ name: 'big.json', kind: 'Unknown', bytes: big });

// ---- the browser had these settings when the workbook was built
store.set('jjset:revenue', { match: { tol: 5 }, marker: 'workpaper' });
store.set(`jjrev:${G}:${FY}`, { user: { reasons: { a: { r: 'TIMING' } }, b2c: { x: '2025-05' } }, audit: [{}], imported: [] });
store.set(`jjrev:cls:${G}`, { 'Sales Local': 'taxable' });
store.set(`jjrev:cust:${G}`, { 'ACME TRADERS': '32ABCDE1234F1Z5' });
setStateProvider(() => ({
  keys: { 'jjmap:sales:invoice no|date': { invNo: 0, date: 1 }, jjalias: { 'ACME TRADERS': '32ABCDE1234F1Z5' } },
  files: src,
  legal: 'DEMO PVT LTD',
}));
const pack = packState('revenue', G, FY, '21-11-2026 14:00');
assert.equal(pack.state.files.length, src.length);
assert.ok(
  pack.state.keys['jjset:revenue'] && pack.state.keys[`jjrev:cls:${G}`] && pack.state.keys['jjmap:sales:invoice no|date'],
);

// ---- a revenue-like workbook (the sheets classify() looks for) gets the hidden sheets
const book = await writeBook([
  { name: 'Read Me', rows: [['x']] },
  { name: 'Books_to_GSTR1', rows: [['x']] },
  { name: 'Rev_Items', rows: [['Id']] },
]);
const once = await addStateSheets(book, pack);
const twice = await addStateSheets(once, pack); // building again from a dropped workpaper: replaced, not added twice
for (const bytes of [once, twice]) {
  const rd = await Reader.open(bytes);
  assert.ok(rd.has(STATE_SHEET) && rd.has(FILES_SHEET));
  assert.equal(rd.sheetNames().filter(n => n === STATE_SHEET).length, 1);
  assert.equal(rd.sheetNames().length, 5);
}
const wbXml = await (await import('jszip')).default.loadAsync(twice).then(zz => zz.file('xl/workbook.xml')!.async('string'));
assert.match(wbXml, /name="_JJState"[^>]*state="hidden"/);

async function check(bytes: Uint8Array, label: string) {
  const [f] = await classify('Revenue_Reco.xlsx', bytes, noPdf as any);
  assert.equal(f.kind, 'Revenue workbook', label);
  assert.ok(f.saved, `${label}: state read`);
  assert.equal(f.gstin, G);
  assert.equal(f.fy, FY);
  assert.equal(f.saved!.legal, 'DEMO PVT LTD');
  assert.equal(f.raw, undefined, 'workbooks are never saved inside another workpaper');
  const back = await readSavedFiles(await Reader.open(f.bytes!), f.saved!);
  assert.equal(back.length, src.length, `${label}: all source files back`);
  for (const [i, s] of src.entries()) {
    assert.equal(back[i].name, s.name);
    assert.deepEqual(back[i].bytes, s.bytes, `${label}: ${s.name} unchanged`);
  }
  // the files come back as what they were
  const again = await classify(back[0].name, back[0].bytes, noPdf as any);
  assert.equal(again[0].kind, src[0].kind);
  assert.ok(again[0].raw, 'source files keep their bytes, for the next workpaper');
  return f;
}
const f = await check(twice, 'as written');

// ---- restore into a browser that has other values, then undo
store.set('jjset:revenue', { match: { tol: 1 }, marker: 'browser' });
store.set(`jjrev:cls:${G}`, { 'Other Income': 'nonsupply' });
store.del(`jjrev:cust:${G}`);
const r = applyState(f.name, f.saved!);
assert.equal(store.get('jjset:revenue').marker, 'workpaper', 'settings: the workpaper wins');
assert.deepEqual(store.get(`jjrev:cls:${G}`), { 'Other Income': 'nonsupply', 'Sales Local': 'taxable' }, 'choices are merged');
assert.equal(store.get(`jjrev:${G}:${FY}`).user.reasons.a.r, 'TIMING');
const lines = describeKeys(f.saved!.keys).join(' | ');
for (const t of [
  'Matching settings',
  '1 supplier name',
  '1 revenue ledger class',
  '1 customer name match',
  '1 reason',
  'Column mapping of 1 register',
])
  assert.ok(lines.includes(t), `restored list says "${t}": ${lines}`);
assert.equal(r.lines.length, describeKeys(f.saved!.keys).length);
undoState(r);
assert.equal(store.get('jjset:revenue').marker, 'browser', 'undo puts the browser settings back');
assert.deepEqual(store.get(`jjrev:cls:${G}`), { 'Other Income': 'nonsupply' });
assert.equal(store.get(`jjrev:cust:${G}`), null, 'a key the browser did not have is removed again');

// ---- the template workbooks (Input Reco) keep their classification
const tpl = new Uint8Array(readFileSync(new URL('../template/ITC_Recon_Template.xlsx', import.meta.url)));
const itc = await addStateSheets(await (await Workbook.open(tpl)).save(), packState('itc', G, FY, ''));
const [fi] = await classify('ITC.xlsx', itc, noPdf as any);
assert.equal(fi.kind, 'Workbook');
assert.equal(fi.note, 'ITC reconciliation workbook');
assert.equal(fi.saved?.tab, 'itc');

// ---- saved again by a spreadsheet program (LibreOffice here, as Excel would: shared strings, new sheet files)
if (existsSync('/usr/bin/soffice')) {
  const d = mkdtempSync(join(tmpdir(), 'jjstate-'));
  writeFileSync(join(d, 'in.xlsx'), twice);
  const out = join(d, 'out');
  const res = spawnSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', out, join(d, 'in.xlsx')], {
    encoding: 'utf8',
    timeout: 120000,
  });
  if (res.status === 0 && existsSync(join(out, 'in.xlsx'))) {
    const resaved = new Uint8Array(readFileSync(join(out, 'in.xlsx')));
    await check(resaved, 're-saved');
    console.log('re-saved by LibreOffice: state and files read back');
  } else console.log('LibreOffice conversion not available, skipped the re-save check');
}
console.log('state: ok');

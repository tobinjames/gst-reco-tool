// Make a test case from real client files without keeping real GSTINs in the repo.
// Every GSTIN in the files is replaced by a made-up one, the same way in every file (so matching still works):
// the state code and the entity number are kept, and two GSTINs with the same PAN still share one made-up PAN.
// Amounts, dates, invoice numbers and names are not changed (names drive the supplier-name matching).
//
// Usage: npx tsx scripts/anonymise.ts <folder with the real files> <case name> [--keep <client GSTIN>] [--mode itc|annual] [--fy 2025-26]
//   e.g. npx tsx scripts/anonymise.ts ~/Desktop/abc-oct test-case-abc-oct --mode itc --fy 2025-26
// Writes test/cases/<case name>/ with the changed files and a case.json, then run `npm run test:update` and check expected.json.
// --keep leaves the client's own GSTIN as it is: use it when the case has GSTR-3B PDFs, which this script cannot change.
import { createHash } from 'crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'fs';
import { join } from 'path';
import JSZip from 'jszip';

const args = process.argv.slice(2);
const opt = (k: string) => { const i = args.indexOf(k); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const keep = (opt('--keep') || '').toUpperCase(), mode = opt('--mode') || 'itc', fy = opt('--fy') || '2025-26';
const [src, name] = args;
if (!src || !name) { console.log('Usage: npx tsx scripts/anonymise.ts <folder with the real files> <case name> [--keep <client GSTIN>] [--mode itc|annual] [--fy 2025-26]'); process.exit(1); }

const GSTIN = /\b(\d{2})([A-Z]{5}\d{4}[A-Z])([1-9A-Z])Z([0-9A-Z])\b/g;
const L = 'ABCDEFGHJKLMNPQRSTUVWXYZ', D = '0123456789';
/** made-up PAN from the real one (stable, so the same PAN always becomes the same fake one) */
const fakePan = (pan: string) => {
  const h = createHash('sha256').update('jj-anon:' + pan).digest();
  // 4th letter (holder type: C company, P person, F firm...) is kept so supplier types still read right
  return [0, 1, 2].map(i => L[h[i] % L.length]).join('') + pan[3] + L[h[3] % L.length] + [4, 5, 6, 7].map(i => D[h[i] % 10]).join('') + L[h[8] % L.length];
};
const seen = new Map<string, string>();
const swap = (s: string) => s.replace(GSTIN, (g, st, pan, ent) => {
  if (g === keep) return g;
  const f = st + fakePan(pan) + ent + 'Z' + L[createHash('sha256').update(g).digest()[0] % L.length];
  seen.set(g, f);
  return f;
});

const out = new URL(`../test/cases/${name}/`, import.meta.url).pathname;
mkdirSync(out, { recursive: true });
for (const f of readdirSync(src).filter(f => !f.startsWith('.'))) {
  const bytes = readFileSync(join(src, f));
  const ext = f.toLowerCase().split('.').pop();
  if (ext === 'json' || ext === 'csv' || ext === 'txt') writeFileSync(join(out, f), swap(bytes.toString('utf8')));
  else if (ext === 'xlsx' || ext === 'zip') {
    // .xlsx is a zip of XML files: change the text inside, leave everything else (formats, layout) as it is
    const z = await JSZip.loadAsync(bytes);
    for (const p of Object.keys(z.files)) if (/\.(xml|json|csv|txt)$/i.test(p) && !z.files[p].dir) z.file(p, swap(await z.file(p)!.async('string')));
    writeFileSync(join(out, f), await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  } else if (ext === 'pdf') {
    writeFileSync(join(out, f), bytes);
    console.log(`! ${f}: PDFs are copied unchanged (GSTINs inside are not replaced). Use --keep <client GSTIN> so the other files still agree with it.`);
  } else {
    console.log(`! ${f}: skipped (only .json, .csv, .txt, .xlsx, .zip and .pdf are handled; save .xls files as .xlsx first)`);
    continue;
  }
  console.log(`✓ ${f}`);
}
if (!existsSync(join(out, 'case.json'))) writeFileSync(join(out, 'case.json'), JSON.stringify({ mode, fy, about: 'Describe the client situation this case covers (no client name)' }, null, 1) + '\n');
console.log(`\n${seen.size} GSTINs replaced${keep ? `, ${keep} kept` : ''}. Case written to test/cases/${name}/.\nNext: edit case.json (mode, FY, about), run npm run test:update, and read expected.json before committing.`);

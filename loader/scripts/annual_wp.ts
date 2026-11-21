// Annual return from saved workpapers (the recommended route), without the page: classify the workpapers, bring back their
// source files and build the GSTR-9 workbook. Usage: npx tsx scripts/annual_wp.ts <folder with workpapers> <out.xlsx>
import { readFileSync, writeFileSync, readdirSync } from 'fs';
import { InFile, classify, buildWorkbook } from '../src/build';
import { readSavedFiles } from '../src/state';
import { Reader } from '../src/xlsx';
const dir = process.argv[2];
const files: InFile[] = [];
const nopdf = async () => {
  throw new Error('no pdf');
};
for (const n of readdirSync(dir)) {
  const got = await classify(n, new Uint8Array(readFileSync(dir + '/' + n)), nopdf as any);
  for (const g of got) {
    files.push(g);
    if (g.saved)
      for (const sf of await readSavedFiles(await Reader.open(g.bytes!), g.saved))
        files.push(...(await classify(sf.name, sf.bytes, nopdf as any)));
  }
}
console.log(files.map(f => f.kind + ' ' + f.name).join('\n'));
const res = await buildWorkbook(
  new Uint8Array(readFileSync(new URL('../template/GST_Annual_Return_Workbook.xlsx', import.meta.url))),
  files.filter(f => !['Register', 'Portal ledger', 'Books ledger'].includes(f.kind)),
  '2025-26',
  {},
);
writeFileSync(process.argv[3], res.bytes);
console.log(res.warns, res.recon);

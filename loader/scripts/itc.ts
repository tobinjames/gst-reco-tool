// Monthly ITC mode on real files: classify them and build the ITC reconciliation workbook (FY 2026-27; registers need the loader page).
// Usage: npx tsx scripts/itc.ts <out.xlsx> <files...>
import { readFileSync, writeFileSync } from 'fs';
import { InFile, classify, buildItc } from '../src/build';
const files: InFile[] = [];
for (const p of process.argv.slice(3))
  files.push(
    ...(await classify(p.split('/').pop()!, new Uint8Array(readFileSync(p)), async () => {
      throw new Error('no pdf');
    })),
  );
const r = await buildItc(
  new Uint8Array(readFileSync(new URL('../template/ITC_Recon_Template.xlsx', import.meta.url))),
  files,
  '2026-27',
  {},
);
writeFileSync(process.argv[2], r.bytes);
console.log(r.recon, r.position.claimable, r.warns);

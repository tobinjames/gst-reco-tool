import { readFileSync, writeFileSync } from 'fs';
import { classify, buildItc } from '../src/build';
const files = [];
for (const p of process.argv.slice(3)) files.push(...await classify(p.split('/').pop()!, new Uint8Array(readFileSync(p)), async () => { throw new Error('no pdf'); }));
const r = await buildItc(new Uint8Array(readFileSync('template/ITC_Recon_Template.xlsx')), files, '2026-27', {});
writeFileSync(process.argv[2], r.bytes); console.log(r.recon, r.position.claimable, r.warns);

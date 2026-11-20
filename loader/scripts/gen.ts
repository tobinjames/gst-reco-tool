// Print the GSTR-9 upload JSON generated from a finished workbook.
// Usage: npx tsx scripts/gen.ts <workbook.xlsx>
import { readFileSync, writeFileSync } from 'fs';
import { generate9 } from '../src/gen9';
const r = await generate9(new Uint8Array(readFileSync(process.argv[2])));
console.log(JSON.stringify(r.json, null, 1).slice(0, 3000)); console.log(r.msgs);

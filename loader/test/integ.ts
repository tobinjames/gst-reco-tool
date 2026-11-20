import { readFileSync } from 'fs';
import { generate9 } from '../src/gen9';
const tpl = new Uint8Array(readFileSync('template/GST_Annual_Return_Workbook.xlsx'));
const r = await generate9(new Uint8Array(readFileSync(process.argv[2])), tpl);
console.log(r.integrity?.items.map(i => `[${i.level}] ${i.area}: ${i.msg}`).join('\n'));
console.log('ADJ', JSON.stringify(r.integrity?.adjustments));
console.log('MSGS', r.msgs.filter(m => m.level === 'error').map(m => m.msg));

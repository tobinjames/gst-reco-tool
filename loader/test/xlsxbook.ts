// xlsxbook: cells stay within Excel's limits (a longer cell makes Excel "repair" the file)
// Run: npx tsx test/xlsxbook.ts
import assert from 'assert/strict';
import JSZip from 'jszip';
import { writeBook } from '../src/xlsxbook';

const long = Array.from({ length: 6000 }, (_, i) => `BL${String(i).padStart(5, '0')}`).join(', '); // ~54,000 characters
const bytes = await writeBook([{ name: 'S', rows: [[long, { v: 'bad ￾ char \uD800 lone' }, { f: 'A1', v: long }]] }]);
const xml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string');
for (const m of xml.matchAll(/<(t|v)[^>]*>([^<]*)<\/\1>/g)) assert.ok(m[2].length <= 32767, `cell text of ${m[2].length} characters`);
assert.ok(!/￾|\uD800/.test(xml), 'characters XML does not allow are dropped');
console.log('xlsxbook.ts: all checks passed');

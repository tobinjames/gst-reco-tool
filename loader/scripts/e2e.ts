// Annual mode end to end on real files: classify them and build the GSTR-9 workbook.
// Usage: npx tsx scripts/e2e.ts <FY e.g. 2025-26> <out.xlsx> <files...>
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { readFileSync, writeFileSync } from 'fs';
import { InFile, classify, buildWorkbook } from '../src/build';
import { linesFromItems, parse3b } from '../src/gstr3b';
const parsePdf = async (b: Uint8Array, name: string) => {
  const doc = await pdfjs.getDocument({ data: b }).promise;
  const pages: any[] = [];
  for (let i = 1; i <= doc.numPages; i++) pages.push(await (await doc.getPage(i)).getTextContent());
  return parse3b(linesFromItems(pages), name);
};
const [fy, out, ...paths] = process.argv.slice(2);
const files: InFile[] = [];
for (const p of paths) files.push(...(await classify(p.split('/').pop()!, new Uint8Array(readFileSync(p)), parsePdf)));
console.log(files.map(f => [f.name, f.kind, f.gstin, f.period, f.fy, f.note].join(' | ')).join('\n'));
const res = await buildWorkbook(
  new Uint8Array(readFileSync(new URL('../template/GST_Annual_Return_Workbook.xlsx', import.meta.url))),
  files,
  fy,
  {},
);
writeFileSync(out, res.bytes);
console.log(res.counts, res.summary);
console.log(res.warns);

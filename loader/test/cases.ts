// Real-file test cases: each folder in test/cases/ holds the files you would drop into the Loader, a case.json
// saying which mode and FY, and expected.json, the answers the tool gave when the case was last checked by hand.
// Running it repeats the whole Loader flow (classify, read registers, match, build the workbook) and fails if any
// answer changed. See test/cases/README.md.
//   npm test                    check every case (part of the full test run)
//   npm run test:update         rewrite expected.json after you have checked a change is right
//   npx tsx test/cases.ts <name>   one case only
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'fs';
import { join } from 'path';
import { isDeepStrictEqual } from 'util';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { InFile, buildItc, buildWorkbook, classify, reconPreview } from '../src/build';
import { linesFromItems, parse3b } from '../src/gstr3b';
import {
  PURCHASE_FIELDS,
  RegKind,
  detectHeader,
  fieldsOf,
  guessMappingFull,
  isMapped,
  normalise,
  pickRegister,
  readTable,
} from '../src/register';

interface CaseDef {
  mode: 'itc' | 'annual';
  fy: string;
  gstin?: string;
  legalName?: string;
  about?: string;
}
const ROOT = new URL('./cases/', import.meta.url).pathname;
const SKIP = new Set(['case.json', 'expected.json', 'README.md', 'notes.md']);
const update = process.argv.includes('--update') || process.env.UPDATE === '1';
const only = process.argv.slice(2).filter(a => !a.startsWith('--'));
const tpl = (f: string) => new Uint8Array(readFileSync(new URL('../template/' + f, import.meta.url)));

async function parsePdf(b: Uint8Array, name: string) {
  const doc = await pdfjs.getDocument({ data: b, useSystemFonts: true }).promise;
  const pages: any[] = [];
  for (let i = 1; i <= doc.numPages; i++) pages.push(await (await doc.getPage(i)).getTextContent());
  return parse3b(linesFromItems(pages), name);
}
const round = (v: any): any =>
  typeof v === 'number'
    ? Math.round(v * 100) / 100
    : Array.isArray(v)
      ? v.map(round)
      : v && typeof v === 'object'
        ? Object.fromEntries(Object.entries(v).map(([k, x]) => [k, round(x)]))
        : v;
const warnTxt = (w: { level: string; file: string; msg: string }[]) =>
  w.map(x => `${x.level}: ${x.file ? x.file + ': ' : ''}${x.msg}`);

/** what the tool produced for one case, in a form that is easy to read in a diff */
async function runCase(dir: string, def: CaseDef) {
  const files: InFile[] = [];
  for (const f of readdirSync(dir)
    .filter(f => !SKIP.has(f) && !f.startsWith('.'))
    .sort())
    files.push(...(await classify(f, new Uint8Array(readFileSync(join(dir, f))), parsePdf)));
  const out: any = {
    files: files.map(f => [f.name, f.kind, f.gstin || '', f.period || '', f.fy || '', f.note || ''].join(' | ')),
  };
  // registers: same automatic choices as the Loader page (sheet, header row, column mapping, defaults)
  const regs: { sales?: any[][]; purchases?: any[][] } = {};
  out.registers = [];
  for (const f of files.filter(f => f.kind === 'Register')) {
    const sheets = await readTable(f.name, f.bytes!);
    let { sheet, header, kind } = pickRegister(f.name, sheets);
    const rows = sheets[sheet].rows;
    if (def.mode === 'itc' && kind !== 'purchases') {
      kind = 'purchases' as RegKind;
      header = detectHeader(rows, PURCHASE_FIELDS);
    }
    const map = guessMappingFull(rows, header, fieldsOf(kind));
    const missing = fieldsOf(kind)
      .filter(x => x.required && !isMapped(map[x.id]))
      .map(x => x.label);
    if (missing.length) {
      out.registers.push({ file: f.name, kind, problem: 'columns not recognised: ' + missing.join(', ') });
      continue;
    }
    const d = {
      noGstinZeroTax: 'EXEMPT',
      itcType: 'Inputs',
      status: 'Availed',
      claimFy: def.fy,
      purNoGstin: isMapped(map.gstin) ? 'RCM-UNREG' : 'B2B',
    };
    const n = normalise(kind, rows, header, map, d, def.fy);
    (regs[kind] ||= []).push(...n.rows);
    const { grossOffEg, ...stats } = n.stats;
    out.registers.push({
      file: f.name,
      sheet: sheets[sheet].name,
      headerRow: header + 1,
      kind,
      columns: Object.fromEntries(
        Object.entries(map).map(([k, v]) => [
          k,
          (Array.isArray(v) ? v : [v]).map(j => String(rows[header][j] ?? '')).join(' + '),
        ]),
      ),
      stats,
    });
  }
  const pv = await reconPreview(files, def.fy, def.mode, regs);
  if (pv.purchases) {
    out.purchases = {
      stats: pv.purchases.stats,
      lines: pv.purchases.rows.map(r => [r[0], r[1], r[2], r[4], r[5]].map(x => String(x ?? '')).join(' | ')),
    };
  }
  if (pv.sales)
    out.sales = {
      stats: pv.sales.stats,
      lines: pv.sales.rows.map(r => [r[0], r[1], r[2] || r[3], r[6], r[7]].map(x => String(x ?? '')).join(' | ')),
    };
  if (def.mode === 'itc') {
    const r = await buildItc(
      tpl('ITC_Recon_Template.xlsx'),
      files.filter(f => f.kind !== 'Register'),
      def.fy,
      { gstin: def.gstin, legalName: def.legalName, purchases: regs.purchases },
    );
    out.itc = { periods: r.periods, position: round(r.position), warnings: warnTxt(r.warns) };
  } else {
    const r = await buildWorkbook(
      tpl('GST_Annual_Return_Workbook.xlsx'),
      files.filter(f => f.kind !== 'Register'),
      def.fy,
      { gstin: def.gstin, legalName: def.legalName, ...regs },
    );
    out.annual = { counts: r.counts, summary: round(r.summary), warnings: warnTxt(r.warns) };
  }
  return round(out);
}

/** JSON with one line per item, and lists of numbers kept on one line */
const pretty = (v: any) =>
  JSON.stringify(v, null, 1).replace(/\[\s+([-\d.,\s]+?)\s+\]/g, (_, inner) => '[' + inner.replace(/\s+/g, ' ').trim() + ']') +
  '\n';

/** first few differences between two JSON values, as paths */
function diffs(a: any, b: any, path = '', acc: string[] = []): string[] {
  if (acc.length >= 15 || isDeepStrictEqual(a, b)) return acc;
  if (Array.isArray(a) && Array.isArray(b) && a.every(x => typeof x === 'string') && b.every(x => typeof x === 'string')) {
    const sa = new Set(a),
      sb = new Set(b);
    for (const x of a) if (!sb.has(x) && acc.length < 15) acc.push(`${path}: no longer   ${x}`);
    for (const x of b) if (!sa.has(x) && acc.length < 15) acc.push(`${path}: now          ${x}`);
    if (!acc.length) acc.push(`${path}: same lines, different order`);
    return acc;
  }
  if (a && b && typeof a === 'object' && typeof b === 'object') {
    for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) diffs(a[k], b[k], path ? `${path}.${k}` : k, acc);
    return acc;
  }
  acc.push(`${path}: expected ${JSON.stringify(a)}, got ${JSON.stringify(b)}`);
  return acc;
}

let failed = 0,
  ran = 0;
const names = existsSync(ROOT)
  ? readdirSync(ROOT, { withFileTypes: true })
      .filter(d => d.isDirectory() && !d.name.startsWith('_'))
      .map(d => d.name)
      .sort()
  : [];
for (const name of names.filter(n => !only.length || only.includes(n))) {
  const dir = join(ROOT, name);
  if (!existsSync(join(dir, 'case.json'))) {
    console.log(`- ${name}: skipped (no case.json)`);
    continue;
  }
  const def: CaseDef = JSON.parse(readFileSync(join(dir, 'case.json'), 'utf8'));
  ran++;
  let got: any;
  try {
    got = await runCase(dir, def);
  } catch (e: any) {
    failed++;
    console.log(`✗ ${name}: the tool stopped with an error: ${e?.stack || e}`);
    continue;
  }
  const expPath = join(dir, 'expected.json');
  if (update || !existsSync(expPath)) {
    writeFileSync(expPath, pretty(got));
    console.log(`✎ ${name}: expected.json ${update ? 'updated' : 'created'}. Check it before committing.`);
    continue;
  }
  const d = diffs(JSON.parse(readFileSync(expPath, 'utf8')), got);
  if (d.length) {
    failed++;
    console.log(`✗ ${name}: results changed\n  ${d.join('\n  ')}\n  If the new results are right, run: npm run test:update`);
  } else console.log(`✓ ${name}`);
}
console.log(`cases: ${ran - failed} of ${ran} unchanged`);
if (failed) process.exit(1);

// Edit detection between the workbook the Loader produced and the workbook dropped for JSON generation
import ROWMAP from './rowmap.json';
import { Reader, rc } from './xlsx';
import { sha1 } from './build';

export const RAW_SHEETS = ['RAW_Portal9', 'RAW_GSTR1', 'RAW_GSTR1_HSN', 'RAW_3B', 'RAW_3B_PAY', 'RAW_2B'];
// Signature over the fingerprint sheet: a change to it (or to the GSTIN / FY afterwards) breaks the seal
const SEAL_KEY = 'jj-gstr9-seal-7f3c1a9e5b2d4806-2026';
export async function seal(lines: string[]): Promise<string> {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(SEAL_KEY), { name: 'HMAC', hash: 'SHA-256' }, false, [
    'sign',
  ]);
  const sig = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(lines.join('\n')));
  return Array.from(new Uint8Array(sig))
    .map(b => b.toString(16).padStart(2, '0'))
    .join('')
    .slice(0, 32);
}
export const sealLines = (rows: any[][], gstin: string, fy: string) => [
  `gstin=${gstin.trim().toUpperCase()}`,
  `fy=${fy.trim()}`,
  ...rows
    .filter(r => r[0] !== 'signature')
    .map(r =>
      r
        .slice(0, 3)
        .map(v => (v === null || v === undefined ? '' : String(v)))
        .join('|'),
    ),
];

const norm = (v: any) =>
  v === null || v === undefined || v === '' ? '' : typeof v === 'number' ? v.toFixed(2) : String(v).trim();

/** order-independent fingerprint of data rows (sorting / filtering a sheet does not change it) */
export async function fingerprint(rows: any[][]): Promise<{ hash: string; count: number }> {
  const lines: string[] = [];
  for (const r of rows) {
    const vals = r.map(norm);
    while (vals.length && vals[vals.length - 1] === '') vals.pop();
    if (vals.length) lines.push(vals.join('\t'));
  }
  lines.sort();
  return { hash: (await sha1(new TextEncoder().encode(lines.join('\n')))).slice(0, 16), count: lines.length };
}
export async function readRows(rd: Reader, sheet: string, startRow: number) {
  const map = await rd.sheet(sheet);
  const rows: any[][] = [];
  for (const [ref, c] of map) {
    const [r, col] = rc(ref);
    if (r < startRow) continue;
    (rows[r - startRow] ||= [])[col - 1] = c.v;
  }
  return rows.filter(Boolean).map(r => Array.from(r, v => v ?? null));
}

const normF = (f: string) =>
  f
    .replace(/^=/, '')
    .split(/("(?:[^"]|"")*")/)
    .map((p, i) =>
      i % 2
        ? p
        : p
            .replace(/'((?:[^']|'')+)'!/g, '$1!')
            .replace(/_xlfn\./gi, '')
            .replace(/\s+/g, '')
            .toUpperCase(),
    )
    .join('');

export interface IntegrityItem {
  level: 'error' | 'warn' | 'info';
  area: string;
  msg: string;
}
export interface Integrity {
  items: IntegrityItem[];
  adjustments: { where: string; what: string; values: string; remark: string }[];
  builtAt: string;
  templateVersion: string;
}

const STRICT = new Set([
  'Basic',
  '4-5 Outward',
  '6-8 ITC',
  '9 & 14 Tax Paid',
  '10-13 Next FY',
  '15-16 Other',
  '17 HSN Outward',
  '18 HSN Inward',
  'Checks',
  'Books_Sales',
  'Books_Purchase',
  'Map',
]);
const SOFT = new Set(['GSTR-9C']); // formulas here may be overwritten on purpose (e.g. 9C Table 9 rate rows)

export async function checkIntegrity(up: Reader, tpl: Reader): Promise<Integrity> {
  const items: IntegrityItem[] = [];
  const res: Integrity = { items, adjustments: [], builtAt: '', templateVersion: '' };
  // ---- 1. template version & loader fingerprints
  const tplVer = String((await tpl.get('_Integrity', 'B2')).v ?? '');
  if (!up.has('_Integrity')) {
    items.push({
      level: 'warn',
      area: 'Fingerprint',
      msg: 'This workbook was not produced by the Loader (no fingerprint sheet). Imported-data checks skipped.',
    });
  } else {
    const upVer = String((await up.get('_Integrity', 'B2')).v ?? '');
    res.templateVersion = upVer;
    if (upVer !== tplVer)
      items.push({
        level: 'warn',
        area: 'Template',
        msg: `Workbook template ${upVer || '(unknown)'} differs from this Loader's template ${tplVer}. Formula comparison may show differences that come from the template update - rebuild the workbook with this Loader to be sure.`,
      });
    const ig = await up.sheet('_Integrity');
    // verify the seal
    const igRows: any[][] = [];
    let sig = '';
    for (let r = 4; r <= 40; r++) {
      const a = (await up.get('_Integrity', 'A' + r)).v;
      if (a === null || a === undefined || a === '') continue;
      const b = (await up.get('_Integrity', 'B' + r)).v,
        c = (await up.get('_Integrity', 'C' + r)).v;
      if (a === 'signature') sig = String(b ?? '');
      else igRows.push([a, b, c === null || c === undefined ? null : Number(c)]);
    }
    const gst = String((await up.get('Basic', 'C6')).v ?? ''),
      fyv = String((await up.get('Basic', 'C7')).v ?? '');
    const expect = await seal(sealLines(igRows, gst, fyv));
    if (!sig) items.push({ level: 'error', area: 'Seal', msg: 'The workbook has no Loader seal - rebuild it with this Loader.' });
    else if (sig !== expect)
      items.push({
        level: 'error',
        area: 'Seal',
        msg: 'The Loader seal does not match: the hidden fingerprint sheet, or the GSTIN / financial year on the Basic sheet, was changed after the workbook was built. Rebuild the workbook with the Loader.',
      });
    for (const [ref, c] of ig) {
      const [r, col] = rc(ref);
      if (col !== 1 || r < 4) continue;
      const sheet = String(c.v);
      if (sheet === 'built_at') {
        res.builtAt = String((await up.get('_Integrity', 'B' + r)).v ?? '');
        continue;
      }
      if (!RAW_SHEETS.includes(sheet)) continue;
      const want = String((await up.get('_Integrity', 'B' + r)).v ?? ''),
        wantN = Number((await up.get('_Integrity', 'C' + r)).v ?? 0);
      const got = await fingerprint(await readRows(up, sheet, 6));
      if (got.hash !== want)
        items.push({
          level: 'error',
          area: 'Imported data',
          msg: `${sheet} has been edited since the Loader built it (${wantN} rows then, ${got.count} now). Imported return data should not be edited - reload the returns, or put corrections in the Adjustments columns.`,
        });
    }
  }
  // ---- 2. formulas and structure vs the original template
  let strictCount = 0;
  for (const sheet of tpl.sheetNames()) {
    if (!STRICT.has(sheet) && !SOFT.has(sheet)) continue;
    if (!up.has(sheet)) {
      items.push({ level: 'error', area: 'Structure', msg: `Sheet "${sheet}" is missing` });
      continue;
    }
    const tmap = await tpl.sheet(sheet),
      umap = await up.sheet(sheet);
    const changed: string[] = [],
      replaced: string[] = [];
    for (const [ref, tc] of tmap) {
      if (!tc.f || !tc.ftext) continue;
      const uc = umap.get(ref);
      if (!uc || !uc.f) {
        replaced.push(ref);
        continue;
      }
      if (uc.ftext !== undefined && normF(uc.ftext) !== normF(tc.ftext)) changed.push(ref);
    }
    const lvl = SOFT.has(sheet) ? 'warn' : 'error';
    const show = (l: string[]) => l.slice(0, 12).join(', ') + (l.length > 12 ? ` … (+${l.length - 12} more)` : '');
    if (replaced.length)
      items.push({
        level: lvl,
        area: 'Formulas',
        msg: `${sheet}: ${replaced.length} formula cell(s) overwritten with typed values or cleared: ${show(replaced)}`,
      });
    if (changed.length)
      items.push({ level: lvl, area: 'Formulas', msg: `${sheet}: ${changed.length} formula(s) changed: ${show(changed)}` });
    if (lvl === 'error') strictCount += replaced.length + changed.length;
  }
  // row keys (catch inserted / deleted / moved rows on the GSTR-9 sheets)
  for (const [key, [sheet, row]] of Object.entries(ROWMAP as unknown as Record<string, [string, number]>)) {
    if (!['4-5 Outward', '6-8 ITC', '10-13 Next FY'].includes(sheet)) continue;
    const v = (await up.get(sheet, 'AK' + row)).v;
    if (String(v ?? '') !== key) {
      items.push({
        level: 'error',
        area: 'Structure',
        msg: `${sheet}: row ${row} should be key ${key} but holds "${v ?? ''}" - rows were inserted, deleted or moved`,
      });
      break;
    }
  }
  if (!items.some(i => i.area === 'Formulas' || i.area === 'Structure'))
    items.push({ level: 'info', area: 'Formulas', msg: 'All formulas and row keys match the original template.' });
  // ---- 3. audit list of intended edits: Basis changes and adjustments
  const BLK = { adj: ['T', 'U', 'V', 'W', 'X'] };
  for (const [key, [sheet, row]] of Object.entries(ROWMAP as unknown as Record<string, [string, number]>)) {
    if (!['4-5 Outward', '6-8 ITC', '10-13 Next FY'].includes(sheet)) continue;
    const desc =
      String((await up.get(sheet, 'B' + row)).v ?? '').slice(0, 60) || String((await up.get(sheet, 'C' + row)).v ?? '');
    const remark = String((await up.get(sheet, 'AJ' + row)).v ?? '');
    const tb = String((await tpl.get(sheet, 'D' + row)).v ?? ''),
      ub = String((await up.get(sheet, 'D' + row)).v ?? '');
    if (tb && ub && tb !== ub) res.adjustments.push({ where: `${key}`, what: `Basis ${tb} → ${ub}`, values: desc, remark });
    const adj = await Promise.all(BLK.adj.map(c => up.get(sheet, c + row)));
    const nz = adj
      .map((c, i) => [['Taxable', 'CGST', 'SGST', 'IGST', 'Cess'][i], !c.f && typeof c.v === 'number' ? c.v : 0] as const)
      .filter(([, v]) => Math.abs(v) >= 0.01);
    if (nz.length) {
      res.adjustments.push({
        where: key,
        what: 'Adjustment',
        values: nz.map(([h, v]) => `${h} ${v.toLocaleString('en-IN')}`).join(', '),
        remark,
      });
      if (!remark) items.push({ level: 'warn', area: 'Adjustments', msg: `${key}: adjustment without a remark` });
    }
  }
  for (const [ref, label] of [
    ['E6', '9 IGST payable'],
    ['E7', '9 CGST payable'],
    ['E8', '9 SGST payable'],
    ['E9', '9 Cess payable'],
    ['E10', '9 Interest'],
    ['E11', '9 Late fee'],
    ['E12', '9 Penalty'],
    ['E13', '9 Other'],
  ]) {
    const v = (await up.get('9 & 14 Tax Paid', ref)).v;
    if (typeof v === 'number' && Math.abs(v) >= 0.01)
      res.adjustments.push({
        where: label,
        what: 'Adjustment',
        values: v.toLocaleString('en-IN'),
        remark: String((await up.get('9 & 14 Tax Paid', 'O' + ref.slice(1))).v ?? ''),
      });
  }
  return res;
}

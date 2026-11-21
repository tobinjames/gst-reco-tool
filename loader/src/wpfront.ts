// The front of every J&J workpaper written with xlsxbook.ts: an Index (cover, contents with links, colour key,
// sign-off) and one Instructions sheet, plus the shared sheet heading with a link back to the Index.
// Tobin (Nov 2026): Aptos throughout, a proper index, summary sheets, instructions on one sheet, and formulas that
// someone without the tool can follow.
import { C, CellObj, SheetSpec } from './xlsxbook';

export const INDEX = 'Index';
export const INSTRUCTIONS = 'Instructions';
export const FIRM = 'JAMES & JAMES LLP  ·  CHARTERED ACCOUNTANTS';

export type SheetRole = 'Summary' | 'Working' | 'Your input' | 'Source data' | 'Reference';
export interface SheetInfo {
  name: string;
  role: SheetRole;
  /** what the sheet shows, and where its figures come from */
  about: string;
}
export interface FrontMeta {
  /** e.g. "Revenue reconciliation: books vs GSTR-1 vs GSTR-3B" */
  title: string;
  legalName: string;
  gstin: string;
  fy: string;
  period?: string;
  builtAt: string;
  version: string;
  /** extra cover lines (label, value) */
  extra?: [string, string][];
}
export interface Section {
  heading: string;
  /** [label, text]; a label of '' continues the text of the line above */
  lines: [string, string][];
}

const q = (sheet: string) => `'${sheet.replace(/'/g, "''")}'`;
/** a clickable link to a sheet (HYPERLINK works in every Excel version and in LibreOffice) */
export const sheetLink = (sheet: string, text = sheet, cell = 'A1'): CellObj => ({
  f: `HYPERLINK("#${q(sheet)}!${cell}","${text.replace(/"/g, '""')}")`,
  v: text,
  s: 'hlink',
});
export const backLink = (): CellObj => sheetLink(INDEX, '‹ Back to Index');

/** the three heading rows of a sheet: title, client line, link back to the Index */
export function heading(t: string, meta: Pick<FrontMeta, 'legalName' | 'gstin' | 'fy' | 'period'>): C[][] {
  return [
    [{ v: t, s: 'title' }],
    [
      {
        v: `${meta.legalName || 'Client'} · GSTIN ${meta.gstin || '-'} · FY ${meta.fy}${meta.period ? ' · ' + meta.period : ''} · Prepared by James & James LLP`,
        s: 'sub',
      },
    ],
    [backLink()],
  ];
}

const band = (v: string, s: 'band' | 'bandSub', n = 5): C[] => [{ v, s }, ...Array(n - 1).fill({ v: '', s })];

export function indexSheet(meta: FrontMeta, sheets: SheetInfo[]): SheetSpec {
  const rows: C[][] = [
    band(FIRM, 'band'),
    band(meta.title, 'bandSub'),
    [],
    [{ v: 'Client', s: 'bold' }, meta.legalName || '-'],
    [{ v: 'GSTIN', s: 'bold' }, meta.gstin || '-'],
    [{ v: 'Financial year', s: 'bold' }, meta.fy],
    ...(meta.period ? [[{ v: 'Period', s: 'bold' }, meta.period] as C[]] : []),
    ...(meta.extra ?? []).map(([k, v]) => [{ v: k, s: 'bold' }, v] as C[]),
    [{ v: 'Prepared on', s: 'bold' }, meta.builtAt],
    [{ v: 'Prepared with', s: 'bold' }, `J&J GST Reco Loader v${meta.version}`],
    [],
    [
      { v: 'Contents', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
    ],
    [
      { v: 'No.', s: 'hdr' },
      { v: 'Sheet (click to open)', s: 'hdr' },
      { v: 'Type', s: 'hdr' },
      { v: 'What it shows / where the figures come from', s: 'hdr' },
      { v: '', s: 'hdr' },
    ],
  ];
  const all: SheetInfo[] = [
    { name: INSTRUCTIONS, role: 'Reference', about: 'How to use and read this workpaper, the settings used and the files it was built from.' },
    ...sheets,
  ];
  all.forEach((s, i) => rows.push([i + 1, sheetLink(s.name), s.role, { v: s.about, s: 'wrap' }]));
  rows.push(
    [],
    [
      { v: 'How to read the sheets', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
    ],
    ['Summary', { v: 'Key figures. Each one is a link or a SUMIFS to a working sheet, so you can trace it back.', s: 'wrap' }],
    ['Working', { v: 'The reconciliation itself. Totals are SUM / SUMIFS over the source-data sheets.', s: 'wrap' }],
    ['Source data', { v: 'The rows read from the files, as they are, with how the tool classified each one.', s: 'wrap' }],
    [{ v: 'Yellow cells', s: 'input' }, { v: 'Yours to fill in. Everything else is worked out.', s: 'wrap' }],
    [
      { v: 'Formulas', s: 'bold' },
      {
        v: 'Plain Excel only (SUM, SUMIFS, COUNTIFS, IF, links): no macros, no hidden sheets in the working. Click any figure and read its formula in the formula bar; Formulas › Trace Precedents shows where it comes from.',
        s: 'wrap',
      },
    ],
    [],
    [
      { v: 'Sign-off', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
      { v: '', s: 'sec' },
    ],
    [{ v: '', s: 'hdr' }, { v: 'Name', s: 'hdr' }, { v: 'Date', s: 'hdr' }, { v: 'Remarks', s: 'hdr' }],
    [{ v: 'Prepared by', s: 'bold' }, { v: '', s: 'input' }, { v: '', s: 'input' }, { v: '', s: 'input' }],
    [{ v: 'Reviewed by', s: 'bold' }, { v: '', s: 'input' }, { v: '', s: 'input' }, { v: '', s: 'input' }],
  );
  return { name: INDEX, rows, widths: [16, 30, 13, 100, 4], tab: '2C388E', heights: { 1: 26, 2: 20 } };
}

export function instructionsSheet(meta: FrontMeta, sections: Section[]): SheetSpec {
  const rows: C[][] = [...heading(`Instructions: ${meta.title}`, meta), []];
  for (const sec of sections) {
    if (!sec.lines.length) continue;
    rows.push([
      { v: sec.heading, s: 'sec' },
      { v: '', s: 'sec' },
    ]);
    for (const [k, t] of sec.lines) rows.push([k ? { v: k, s: 'bold' } : '', { v: t, s: 'wrap' }]);
    rows.push([]);
  }
  return { name: INSTRUCTIONS, rows, widths: [34, 120] };
}

/** "how this figure is worked out" text for a Summary / statement row, in words next to the formula */
export const how = (t: string): CellObj => ({ v: t, s: 'formula' });

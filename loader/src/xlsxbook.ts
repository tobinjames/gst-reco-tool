// Small multi-sheet .xlsx writer: values, formulas (with cached values), a fixed set of J&J styles, column widths, frozen panes, autofilter.
// Excel 2010+ compatible; formulas are recalculated on open.
import JSZip from 'jszip';
import { xmlEsc } from './util';

export type Style =
  | 'title'
  | 'sub'
  | 'hdr'
  | 'num'
  | 'numB'
  | 'numTot'
  | 'date'
  | 'note'
  | 'bold'
  | 'wrap'
  | 'sec'
  | 'ok'
  | 'bad'
  | 'link'
  | 'input'
  | 'hlink'
  | 'band'
  | 'bandSub'
  | 'kpi'
  | 'formula';
const STYLE_ID: Record<Style, number> = {
  title: 1,
  sub: 2,
  hdr: 3,
  num: 4,
  numB: 5,
  numTot: 6,
  date: 7,
  note: 8,
  bold: 9,
  wrap: 10,
  sec: 11,
  ok: 12,
  bad: 13,
  link: 14,
  input: 15,
  hlink: 16,
  band: 17,
  bandSub: 18,
  kpi: 19,
  formula: 20,
};
export interface CellObj {
  v?: string | number | null;
  f?: string;
  s?: Style;
}
export type C = string | number | null | undefined | CellObj;
export interface SheetSpec {
  name: string;
  rows: C[][];
  widths?: number[];
  freeze?: { row: number; col?: number };
  filter?: { row: number; cols: number };
  tab?: string;
  /** row heights (points) by row number (1-based) */
  heights?: Record<number, number>;
  /** drop-down lists: cells (e.g. "T6:T500") and the list (a range like "Lists!$A$2:$A$20") */
  lists?: { sqref: string; list: string }[];
}

export const col = (i: number) => {
  let s = '';
  i++;
  while (i) {
    const m = (i - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
};
/** dd-mm-yyyy → Excel date serial */
export const serial = (d: string) => {
  const m = d.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  return m ? Math.round((Date.UTC(+m[3], +m[2] - 1, +m[1]) - Date.UTC(1899, 11, 30)) / 86400000) : null;
};

/** Excel refuses (and "repairs") a cell with more than 32,767 characters */
const MAX_CELL = 32000;
const fit = (t: string) => (t.length > MAX_CELL ? t.slice(0, MAX_CELL - 40) + ' … (cut: too long for one Excel cell)' : t);
function cellXml(ref: string, c: C) {
  const o: CellObj = c !== null && typeof c === 'object' ? c : { v: c as any };
  const s = o.s ? ` s="${STYLE_ID[o.s]}"` : typeof o.v === 'number' && !o.f ? ` s="${STYLE_ID.num}"` : '';
  if (o.f) {
    const f = `<f>${xmlEsc(o.f.replace(/^=/, ''))}</f>`;
    if (typeof o.v === 'number' && isFinite(o.v)) return `<c r="${ref}"${s || ` s="${STYLE_ID.num}"`}>${f}<v>${o.v}</v></c>`;
    if (typeof o.v === 'string') return `<c r="${ref}"${s} t="str">${f}<v>${xmlEsc(fit(o.v))}</v></c>`;
    return `<c r="${ref}"${s}>${f}</c>`;
  }
  if (o.v === null || o.v === undefined || o.v === '') return o.s ? `<c r="${ref}"${s}/>` : '';
  if (typeof o.v === 'number' && isFinite(o.v)) return `<c r="${ref}"${s}><v>${o.v}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(fit(String(o.v)))}</t></is></c>`;
}

function sheetXml(sh: SheetSpec) {
  const widths = sh.widths || (sh.rows[0] || []).map(() => 14);
  const fr = sh.freeze;
  const pane = fr
    ? `<pane${fr.col ? ` xSplit="${fr.col}"` : ''}${fr.row ? ` ySplit="${fr.row}"` : ''} topLeftCell="${col(fr.col || 0)}${(fr.row || 0) + 1}" activePane="${fr.row && fr.col ? 'bottomRight' : fr.row ? 'bottomLeft' : 'topRight'}" state="frozen"/>`
    : '';
  const rows = sh.rows
    .map(
      (r, i) =>
        `<row r="${i + 1}"${sh.heights?.[i + 1] ? ` ht="${sh.heights[i + 1]}" customHeight="1"` : ''}>${(r || []).map((c, j) => cellXml(col(j) + (i + 1), c)).join('')}</row>`,
    )
    .join('');
  const af =
    sh.filter && sh.rows.length > sh.filter.row
      ? `<autoFilter ref="A${sh.filter.row}:${col(sh.filter.cols - 1)}${sh.rows.length}"/>`
      : '';
  const dv = sh.lists?.length
    ? `<dataValidations count="${sh.lists.length}">${sh.lists.map(l => `<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorStyle="warning" sqref="${l.sqref}"><formula1>${xmlEsc(l.list)}</formula1></dataValidation>`).join('')}</dataValidations>`
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${sh.tab ? `<sheetPr><tabColor rgb="FF${sh.tab}"/></sheetPr>` : ''}<sheetViews><sheetView workbookViewId="0" showGridLines="0">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="13.5"/>
<cols>${widths.map((w, j) => `<col min="${j + 1}" max="${j + 1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${rows}</sheetData>${af}${dv}<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToHeight="0" paperSize="9"/></worksheet>`;
}

/** J&J workpaper look: Aptos throughout (Tobin, Nov 2026), brand blue #2C388E for titles, headers and links */
export const FONT = 'Aptos';
const fnt = (o: { b?: boolean; i?: boolean; u?: boolean; sz?: number; color?: string }) =>
  `<font>${o.b ? '<b/>' : ''}${o.i ? '<i/>' : ''}${o.u ? '<u/>' : ''}<sz val="${o.sz ?? 10}"/>${o.color ? `<color rgb="FF${o.color}"/>` : ''}<name val="${FONT}"/><family val="2"/></font>`;
const FONTS = [
  fnt({}), // 0 body
  fnt({ b: true, sz: 15, color: '2C388E' }), // 1 title
  fnt({ color: '737B8E' }), // 2 sub
  fnt({ b: true, color: 'FFFFFF' }), // 3 header
  fnt({ b: true }), // 4 bold
  fnt({ b: true, color: '2C388E' }), // 5 brand bold
  fnt({ i: true, sz: 9, color: '737B8E' }), // 6 note
  fnt({ u: true, color: '2C388E' }), // 7 hyperlink
  fnt({ b: true, sz: 12, color: 'FFFFFF' }), // 8 band
  fnt({ sz: 10, color: 'DDE1F2' }), // 9 band sub
  fnt({ b: true, sz: 13, color: '2C388E' }), // 10 kpi figure
  fnt({ sz: 9, color: '4A5068' }), // 11 formula text (Consolas-like look not needed: plain Aptos, smaller)
];
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##,##0.00;[Red]-#,##,##0.00;&quot;-&quot;"/><numFmt numFmtId="165" formatCode="dd-mm-yyyy"/></numFmts>
<fonts count="${FONTS.length}">${FONTS.join('')}</fonts>
<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2C388E"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE3E6F5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8F5EE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFDEEEC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF8E1"/></patternFill></fill></fills>
<borders count="3"><border/><border><top style="thin"><color rgb="FF2C388E"/></top><bottom style="double"><color rgb="FF2C388E"/></bottom></border><border><bottom style="thin"><color rgb="FFD3D7E6"/></bottom></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="21">
<xf/>
<xf fontId="1" applyFont="1"/>
<xf fontId="2" applyFont="1"/>
<xf fontId="3" fillId="2" applyFont="1" applyFill="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="4" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="164" fontId="4" borderId="1" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
<xf numFmtId="165" applyNumberFormat="1"/>
<xf fontId="6" applyFont="1"><alignment wrapText="1" vertical="top"/></xf>
<xf fontId="4" applyFont="1"/>
<xf applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf fontId="5" fillId="3" applyFont="1" applyFill="1"/>
<xf fillId="4" applyFill="1"/>
<xf fillId="5" applyFill="1"/>
<xf fontId="5" applyFont="1"/>
<xf fillId="6" borderId="2" applyFill="1" applyBorder="1"/>
<xf fontId="7" applyFont="1"/>
<xf fontId="8" fillId="2" applyFont="1" applyFill="1"><alignment vertical="center"/></xf>
<xf fontId="9" fillId="2" applyFont="1" applyFill="1"><alignment vertical="center"/></xf>
<xf numFmtId="164" fontId="10" applyNumberFormat="1" applyFont="1"/>
<xf fontId="11" applyFont="1"><alignment wrapText="1" vertical="top"/></xf>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

export async function writeBook(sheets: SheetSpec[], defined: { name: string; ref: string }[] = []): Promise<Uint8Array> {
  const z = new JSZip();
  const n = sheets.length;
  z.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  );
  z.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  const filters = sheets
    .map((s, i) =>
      s.filter && s.rows.length > s.filter.row
        ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${xmlEsc(s.name)}'!$A$${s.filter.row}:$${col(s.filter.cols - 1)}$${s.rows.length}</definedName>`
        : '',
    )
    .join('');
  const names = defined.map(d => `<definedName name="${xmlEsc(d.name)}">${xmlEsc(d.ref)}</definedName>`).join('');
  z.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${sheets.map((s, i) => `<sheet name="${xmlEsc(s.name.slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${filters || names ? `<definedNames>${filters}${names}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
  );
  z.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  );
  z.file('xl/styles.xml', STYLES);
  sheets.forEach((s, i) => z.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s)));
  return await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** One-sheet download (Review step "Download this view"): optional title lines, then rows[0] as the header row, frozen and filtered */
export function simpleXlsx(sheetName: string, rows: any[][], title: string[] = []): Promise<Uint8Array> {
  const top: C[][] = title.length ? [...title.map((t, i) => [{ v: t, s: (i === 0 ? 'title' : 'sub') as Style }]), []] : [];
  const header: C[] = (rows[0] || []).map(v => ({ v, s: 'hdr' as Style }));
  const widths = (rows[0] || []).map((_, j) =>
    Math.min(60, Math.max(10, ...rows.slice(0, 500).map(r => String(r[j] ?? '').length + 2))),
  );
  const hdrRow = top.length + 1;
  return writeBook([
    {
      name: sheetName,
      rows: [...top, header, ...rows.slice(1)],
      widths,
      freeze: { row: hdrRow },
      filter: rows.length > 1 ? { row: hdrRow, cols: (rows[0] || []).length } : undefined,
    },
  ]);
}

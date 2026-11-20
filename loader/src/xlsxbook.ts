// Small multi-sheet .xlsx writer: values, formulas (with cached values), a fixed set of J&J styles, column widths, frozen panes, autofilter.
// Excel 2010+ compatible; formulas are recalculated on open.
import JSZip from 'jszip';

export type Style = 'title' | 'sub' | 'hdr' | 'num' | 'numB' | 'numTot' | 'date' | 'note' | 'bold' | 'wrap' | 'sec' | 'ok' | 'bad' | 'link';
const STYLE_ID: Record<Style, number> = { title: 1, sub: 2, hdr: 3, num: 4, numB: 5, numTot: 6, date: 7, note: 8, bold: 9, wrap: 10, sec: 11, ok: 12, bad: 13, link: 14 };
export interface CellObj { v?: string | number | null; f?: string; s?: Style }
export type C = string | number | null | undefined | CellObj;
export interface SheetSpec { name: string; rows: C[][]; widths?: number[]; freeze?: { row: number; col?: number }; filter?: { row: number; cols: number }; tab?: string }

const xmlEsc = (s: any) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' } as any)[c]).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
export const col = (i: number) => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
/** dd-mm-yyyy → Excel date serial */
export const serial = (d: string) => { const m = d.match(/^(\d{2})-(\d{2})-(\d{4})$/); return m ? Math.round((Date.UTC(+m[3], +m[2] - 1, +m[1]) - Date.UTC(1899, 11, 30)) / 86400000) : null; };

function cellXml(ref: string, c: C) {
  const o: CellObj = c !== null && typeof c === 'object' ? c : { v: c as any };
  const s = o.s ? ` s="${STYLE_ID[o.s]}"` : typeof o.v === 'number' && !o.f ? ` s="${STYLE_ID.num}"` : '';
  if (o.f) {
    const f = `<f>${xmlEsc(o.f.replace(/^=/, ''))}</f>`;
    if (typeof o.v === 'number' && isFinite(o.v)) return `<c r="${ref}"${s || ` s="${STYLE_ID.num}"`}>${f}<v>${o.v}</v></c>`;
    if (typeof o.v === 'string') return `<c r="${ref}"${s} t="str">${f}<v>${xmlEsc(o.v)}</v></c>`;
    return `<c r="${ref}"${s}>${f}</c>`;
  }
  if (o.v === null || o.v === undefined || o.v === '') return o.s ? `<c r="${ref}"${s}/>` : '';
  if (typeof o.v === 'number' && isFinite(o.v)) return `<c r="${ref}"${s}><v>${o.v}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(o.v)}</t></is></c>`;
}

function sheetXml(sh: SheetSpec) {
  const widths = sh.widths || (sh.rows[0] || []).map(() => 14);
  const fr = sh.freeze;
  const pane = fr ? `<pane${fr.col ? ` xSplit="${fr.col}"` : ''}${fr.row ? ` ySplit="${fr.row}"` : ''} topLeftCell="${col(fr.col || 0)}${(fr.row || 0) + 1}" activePane="${fr.row && fr.col ? 'bottomRight' : fr.row ? 'bottomLeft' : 'topRight'}" state="frozen"/>` : '';
  const rows = sh.rows.map((r, i) => `<row r="${i + 1}">${(r || []).map((c, j) => cellXml(col(j) + (i + 1), c)).join('')}</row>`).join('');
  const af = sh.filter && sh.rows.length > sh.filter.row ? `<autoFilter ref="A${sh.filter.row}:${col(sh.filter.cols - 1)}${sh.rows.length}"/>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${sh.tab ? `<sheetPr><tabColor rgb="FF${sh.tab}"/></sheetPr>` : ''}<sheetViews><sheetView workbookViewId="0" showGridLines="0">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="13.5"/>
<cols>${widths.map((w, j) => `<col min="${j + 1}" max="${j + 1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${rows}</sheetData>${af}<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToHeight="0" paperSize="9"/></worksheet>`;
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##,##0.00;[Red]-#,##,##0.00;&quot;-&quot;"/><numFmt numFmtId="165" formatCode="dd-mm-yyyy"/></numFmts>
<fonts count="7"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="14"/><color rgb="FF2C388E"/><name val="Arial"/></font><font><sz val="10"/><color rgb="FF737B8E"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FF2C388E"/><name val="Arial"/></font><font><i/><sz val="9"/><color rgb="FF737B8E"/><name val="Arial"/></font></fonts>
<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2C388E"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE3E6F5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8F5EE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFDEEEC"/></patternFill></fill></fills>
<borders count="3"><border/><border><top style="thin"><color rgb="FF2C388E"/></top><bottom style="double"><color rgb="FF2C388E"/></bottom></border><border><bottom style="thin"><color rgb="FFD3D7E6"/></bottom></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="15">
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
</cellXfs></styleSheet>`;

export async function writeBook(sheets: SheetSpec[], defined: { name: string; ref: string }[] = []): Promise<Uint8Array> {
  const z = new JSZip();
  const n = sheets.length;
  z.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  z.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  const filters = sheets.map((s, i) => s.filter && s.rows.length > s.filter.row ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${xmlEsc(s.name)}'!$A$${s.filter.row}:$${col(s.filter.cols - 1)}$${s.rows.length}</definedName>` : '').join('');
  const names = defined.map(d => `<definedName name="${xmlEsc(d.name)}">${xmlEsc(d.ref)}</definedName>`).join('');
  z.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>${sheets.map((s, i) => `<sheet name="${xmlEsc(s.name.slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${filters || names ? `<definedNames>${filters}${names}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
  z.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  z.file('xl/styles.xml', STYLES);
  sheets.forEach((s, i) => z.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s)));
  return await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// Minimal, style-preserving writer for the workbook template (edits sheet XML inside the .xlsx zip)
import JSZip from 'jszip';
import { xmlEsc } from './util';

/** a value, or { f } for a formula (written without a cached value; the workbook recalculates on open) */
export type V = string | number | null | undefined | { f: string };
const esc = xmlEsc;
const colNum = (c: string) => c.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
export const colLetter = (n: number) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

function cellXml(ref: string, v: V, style?: string) {
  const s = style ? ` s="${style}"` : '';
  if (v === null || v === undefined || v === '') return `<c r="${ref}"${s}/>`;
  if (typeof v === 'object') return `<c r="${ref}"${s}><f>${esc(v.f.replace(/^=/, ''))}</f></c>`;
  if (typeof v === 'number' && isFinite(v)) return `<c r="${ref}"${s}><v>${v}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(String(v))}</t></is></c>`;
}

export class Workbook {
  zip!: JSZip;
  private sheets = new Map<string, string>();   // name -> path
  private xml = new Map<string, string>();      // path -> xml (edited)

  static async open(data: ArrayBuffer | Uint8Array) {
    const wb = new Workbook();
    wb.zip = await JSZip.loadAsync(data);
    const wbXml = await wb.zip.file('xl/workbook.xml')!.async('string');
    const rels = await wb.zip.file('xl/_rels/workbook.xml.rels')!.async('string');
    const relMap = new Map<string, string>();
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = m[0].match(/Id="([^"]+)"/)?.[1], tgt = m[0].match(/Target="([^"]+)"/)?.[1];
      if (id && tgt) relMap.set(id, tgt.replace(/^\/?xl\//, '').replace(/^\//, ''));
    }
    for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
      const name = m[0].match(/name="([^"]+)"/)?.[1]?.replace(/&amp;/g, '&');
      const rid = m[0].match(/r:id="([^"]+)"/)?.[1];
      if (name && rid && relMap.get(rid)) wb.sheets.set(name, 'xl/' + relMap.get(rid));
    }
    return wb;
  }
  has(name: string) { return this.sheets.has(name); }
  private async get(name: string) {
    const p = this.sheets.get(name); if (!p) throw new Error(`Sheet ${name} not in workbook`);
    if (!this.xml.has(p)) this.xml.set(p, await this.zip.file(p)!.async('string'));
    return p;
  }
  private split(xml: string) {
    const a = xml.indexOf('<sheetData'), aEnd = xml.indexOf('>', a) + 1;
    if (xml[aEnd - 2] === '/') return { head: xml.slice(0, a) + '<sheetData>', rows: [] as string[], tail: '</sheetData>' + xml.slice(aEnd) };
    const b = xml.indexOf('</sheetData>');
    const body = xml.slice(aEnd, b);
    const rows = body.match(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g) || [];
    return { head: xml.slice(0, aEnd), rows, tail: xml.slice(b) };
  }
  /** replace all rows >= startRow with the given data (row arrays start at column A) */
  async writeRows(name: string, startRow: number, data: V[][], styled = false) {
    const p = await this.get(name);
    const { head, rows, tail } = this.split(this.xml.get(p)!);
    // optional: reuse the cell styles of the first data row (the template's "style row") on every written row
    const styles: string[] = [];
    if (styled) {
      const sr = rows.find(r => +(r.match(/<row\b[^>]*\br="(\d+)"/)?.[1] || 0) === startRow);
      for (const c of sr?.match(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || []) {
        const ref = c.match(/\br="([A-Z]+)\d+"/)?.[1], st = c.match(/\bs="(\d+)"/)?.[1];
        if (ref && st) styles[colNum(ref) - 1] = st;
      }
    }
    const keep = rows.filter(r => +(r.match(/<row\b[^>]*\br="(\d+)"/)?.[1] || 0) < startRow);
    const out = data.map((vals, i) => {
      const r = startRow + i;
      if (styled) {
        const n = Math.max(vals.length, styles.length);
        return `<row r="${r}">` + Array.from({ length: n }, (_, j) => cellXml(colLetter(j + 1) + r, vals[j], styles[j])).join('') + '</row>';
      }
      return `<row r="${r}">` + vals.map((v, j) => (v === null || v === undefined || v === '') ? '' : cellXml(colLetter(j + 1) + r, v)).join('') + '</row>';
    });
    let xml = head + keep.join('') + out.join('') + tail;
    const last = startRow + data.length - 1;
    xml = xml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${colLetter(Math.max(...data.map(d => d.length), 1))}${Math.max(last, 5)}"/>`);
    this.xml.set(p, xml);
  }
  /** fill the input columns (A..) of a table from startRow, keep formula columns, extend formulas to extra rows, clear old input values */
  async fillTable(name: string, startRow: number, data: V[][], nInput: number) {
    const p = await this.get(name);
    const { head, rows, tail } = this.split(this.xml.get(p)!);
    const byRow = new Map<number, string>();
    rows.forEach(r => byRow.set(+(r.match(/<row\b[^>]*\br="(\d+)"/)?.[1] || 0), r));
    const cellsOf = (row: string | undefined) => {
      const m = new Map<number, string>();
      for (const c of row?.match(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || []) m.set(colNum(c.match(/\br="([A-Z]+)\d+"/)![1]), c);
      return m;
    };
    const master = cellsOf(byRow.get(startRow));
    const formulas: { col: number; style?: string; text: string }[] = [];
    for (const [col, c] of master) {
      const f = c.match(/<f\b[^>]*>([\s\S]*?)<\/f>/)?.[1];
      if (f !== undefined && col > nInput) formulas.push({ col, style: c.match(/\bs="(\d+)"/)?.[1], text: unescXml(f) });
    }
    const lastExisting = Math.max(startRow, ...[...byRow.keys()]);
    const last = Math.max(lastExisting, startRow + data.length - 1);
    for (let r = startRow; r <= last; r++) {
      const i = r - startRow, vals = data[i];
      if (!vals && !byRow.has(r)) continue;
      const cur = cellsOf(byRow.get(r));
      for (let j = 0; j < nInput; j++) {
        const old = cur.get(j + 1) ?? master.get(j + 1);
        const style = old?.match(/\bs="(\d+)"/)?.[1];
        cur.set(j + 1, cellXml(colLetter(j + 1) + r, vals ? vals[j] : null, style));
      }
      if (vals) for (const f of formulas) if (!cur.has(f.col) || !/<f\b/.test(cur.get(f.col)!)) {
        const ref = colLetter(f.col) + r;
        cur.set(f.col, `<c r="${ref}"${f.style ? ` s="${f.style}"` : ''}><f>${esc(shiftFormula(f.text, i, 0))}</f></c>`);
      }
      const open = (byRow.get(r)?.match(/^<row\b[^>]*>/)?.[0] ?? `<row r="${r}">`).replace(/\sspans="[^"]*"/, '').replace(/\/>$/, '>');
      byRow.set(r, open + [...cur.entries()].sort((a, b) => a[0] - b[0]).map(e => e[1]).join('') + '</row>');
    }
    const all = [...byRow.entries()].sort((a, b) => a[0] - b[0]).map(e => e[1]);
    let xml = head + all.join('') + tail;
    xml = xml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${colLetter(Math.max(nInput, ...formulas.map(f => f.col)))}${last}"/>`);
    this.xml.set(p, xml);
  }
  /** set individual cells, keeping the existing cell style */
  async setCells(name: string, cells: Record<string, V>, styleOf: Record<string, string | undefined> = {}) {
    const p = await this.get(name);
    const { head, rows, tail } = this.split(this.xml.get(p)!);
    const byRow = new Map<number, string>();
    rows.forEach(r => byRow.set(+(r.match(/<row\b[^>]*\br="(\d+)"/)?.[1] || 0), r));
    const groups = new Map<number, [string, V][]>();
    for (const [ref, v] of Object.entries(cells)) {
      const r = +ref.match(/\d+/)![0];
      if (!groups.has(r)) groups.set(r, []);
      groups.get(r)!.push([ref, v]);
    }
    for (const [r, list] of groups) {
      let row = byRow.get(r) || `<row r="${r}"></row>`;
      if (row.endsWith('/>')) row = row.slice(0, -2) + '></row>';
      const cellsInRow = row.match(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || [];
      const map = new Map<string, string>();
      for (const c of cellsInRow) map.set(c.match(/\br="([A-Z]+\d+)"/)![1], c);
      for (const [ref, v] of list) {
        const old = map.get(ref);
        const style = ref in styleOf ? styleOf[ref] : old?.match(/\bs="(\d+)"/)?.[1];
        map.set(ref, cellXml(ref, v, style));
      }
      const sorted = [...map.entries()].sort((a, b) => colNum(a[0].replace(/\d+/, '')) - colNum(b[0].replace(/\d+/, ''))).map(e => e[1]);
      const open = row.match(/^<row\b[^>]*>/)![0].replace(/\sspans="[^"]*"/, '');
      byRow.set(r, open + sorted.join('') + '</row>');
    }
    const all = [...byRow.entries()].sort((a, b) => a[0] - b[0]).map(e => e[1]);
    this.xml.set(p, head + all.join('') + tail);
  }
  /** style index of each cell (as the sheet stands now) */
  async styles(name: string, refs: string[]): Promise<Record<string, string | undefined>> {
    const xml = this.xml.get(await this.get(name))!, out: Record<string, string | undefined> = {};
    for (const ref of refs) out[ref] = xml.match(new RegExp(`<c r="${ref}"[^>]*?\\bs="(\\d+)"`))?.[1];
    return out;
  }
  async save(): Promise<Uint8Array> {
    // force full recalculation on open, and drop stale cached formula results everywhere
    const wbXml = await this.zip.file('xl/workbook.xml')!.async('string');
    let w = wbXml.replace(/<calcPr\b[^>]*\/>/, '<calcPr calcId="191029" fullCalcOnLoad="1"/>');
    if (!/<calcPr/.test(w)) w = w.replace('</workbook>', '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    this.zip.file('xl/workbook.xml', w);
    for (const p of this.sheets.values()) {
      let xml = this.xml.get(p) ?? await this.zip.file(p)!.async('string');
      xml = xml.replace(/(<c\b[^>]*>)(<f\b[^>]*(?:\/>|>[\s\S]*?<\/f>))<v>[^<]*<\/v>/g, '$1$2');
      this.zip.file(p, xml);
    }
    if (this.zip.file('xl/calcChain.xml')) {
      this.zip.remove('xl/calcChain.xml');
      const ct = await this.zip.file('[Content_Types].xml')!.async('string');
      this.zip.file('[Content_Types].xml', ct.replace(/<Override[^>]*calcChain[^>]*\/>/g, ''));
      const rels = await this.zip.file('xl/_rels/workbook.xml.rels')!.async('string');
      this.zip.file('xl/_rels/workbook.xml.rels', rels.replace(/<Relationship[^>]*calcChain[^>]*\/>/g, ''));
    }
    return this.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
}

/** Reads cached cell values from a saved workbook (values exist once Excel / LibreOffice has saved it) */
export class Reader {
  private zip!: JSZip; private sheets = new Map<string, string>(); private ss: string[] = [];
  private cache = new Map<string, Map<string, Cell>>();
  static async open(data: Uint8Array) {
    const r = new Reader(); r.zip = await JSZip.loadAsync(data);
    const wbXml = await r.zip.file('xl/workbook.xml')!.async('string');
    const rels = await r.zip.file('xl/_rels/workbook.xml.rels')!.async('string');
    const relMap = new Map<string, string>();
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = m[0].match(/Id="([^"]+)"/)?.[1], tgt = m[0].match(/Target="([^"]+)"/)?.[1];
      if (id && tgt) relMap.set(id, tgt.replace(/^\/?xl\//, '').replace(/^\//, ''));
    }
    for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
      const name = m[0].match(/name="([^"]+)"/)?.[1]?.replace(/&amp;/g, '&');
      const rid = m[0].match(/r:id="([^"]+)"/)?.[1];
      if (name && rid && relMap.get(rid)) r.sheets.set(name, 'xl/' + relMap.get(rid));
    }
    const ssf = r.zip.file('xl/sharedStrings.xml');
    if (ssf) {
      const x = await ssf.async('string');
      r.ss = (x.match(/<si>[\s\S]*?<\/si>/g) || []).map(si => unesc((si.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join('')));
    }
    return r;
  }
  has(name: string) { return this.sheets.has(name); }
  async sheet(name: string) {
    if (this.cache.has(name)) return this.cache.get(name)!;
    const p = this.sheets.get(name); if (!p) throw new Error(`Sheet "${name}" not found`);
    const xml = await this.zip.file(p)!.async('string');
    const map = new Map<string, Cell>();
    const shared = new Map<string, { ref: string; text: string }>();
    for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = m[1], body = m[2] || '';
      const ref = attrs.match(/\br="([A-Z]+\d+)"/)?.[1]; if (!ref) continue;
      const t = attrs.match(/\bt="(\w+)"/)?.[1];
      const f = /<f\b/.test(body);
      let ftext: string | undefined;
      if (f) {
        const fm = body.match(/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/)!;
        const fa = fm[1], ft = fm[2] !== undefined ? unesc(fm[2]) : '';
        const si = fa.match(/\bsi="(\d+)"/)?.[1];
        if (/t="shared"/.test(fa) && si !== undefined) {
          if (ft) { shared.set(si, { ref, text: ft }); ftext = ft; }
          else { const mst = shared.get(si); if (mst) ftext = shiftFormula(mst.text, rc(ref)[0] - rc(mst.ref)[0], rc(ref)[1] - rc(mst.ref)[1]); }
        } else ftext = ft;
      }
      const vRaw0 = body.match(/<v>([\s\S]*?)<\/v>/)?.[1];
      const vRaw = vRaw0 === '' ? undefined : vRaw0;
      let v: any = null;
      if (t === 's' && vRaw !== undefined) v = this.ss[+vRaw] ?? '';
      else if (t === 'inlineStr') v = unesc((body.match(/<t\b[^>]*>([\s\S]*?)<\/t>/)?.[1]) ?? '');
      else if (t === 'str') v = vRaw !== undefined ? unesc(vRaw) : null;
      else if (t === 'e') v = { error: vRaw };
      else if (t === 'b') v = vRaw === '1';
      else if (vRaw !== undefined) v = +vRaw;
      map.set(ref, { v, f, ftext });
    }
    this.cache.set(name, map);
    return map;
  }
  /** value of a cell; throws if it is a formula without a saved result */
  async get(name: string, ref: string): Promise<Cell> { return (await this.sheet(name)).get(ref) ?? { v: null, f: false }; }
  sheetNames() { return [...this.sheets.keys()]; }
}
export interface Cell { v: any; f: boolean; ftext?: string }
/** [row, col] of an A1 reference */
export const rc = (ref: string): [number, number] => { const m = ref.match(/^([A-Z]+)(\d+)$/)!; return [+m[2], colNumber(m[1])]; };
const colNumber = (c: string) => c.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
/** move relative references of a formula by (dr, dc) - used to expand Excel shared formulas */
export function shiftFormula(f: string, dr: number, dc: number): string {
  return f.split(/("(?:[^"]|"")*")/).map((part, i) => i % 2 ? part : part.replace(
    /(^|[^A-Za-z0-9_.!$'])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g,
    (_m, pre, cAbs, col, rAbs, row) => {
      const c = cAbs ? col : colLetter(colNumber(col) + dc);
      const r = rAbs ? row : String(+row + dr);
      return `${pre}${cAbs}${c}${rAbs}${r}`;
    })).join('');
}
const unescXml = (s: string) => unesc(s);
const unesc = (s: string) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

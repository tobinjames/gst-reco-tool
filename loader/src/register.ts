// Sales / purchase register import: read any .xlsx / .csv export, map its columns, normalise to Books_Sales / Books_Purchase rows
import { Reader, rc } from './xlsx';
import { toDate } from './recon';
import { fyOfDate } from './util';

export type RegKind = 'sales' | 'purchases';
export interface RegSheet { name: string; rows: any[][] }
export interface FieldDef { id: string; label: string; required?: boolean; multi?: boolean; syn: string[] }
/** column mapping: field id → column index, or several column indexes that are summed (Tally columnar registers) */
export type Mapping = Record<string, number | number[]>;
/** the columns a mapping entry points to (always an array; empty when not mapped) */
export const colsOf = (v: number | number[] | undefined | null): number[] => (Array.isArray(v) ? v.filter(x => x >= 0) : typeof v === 'number' && v >= 0 ? [v] : []);
export const isMapped = (v: number | number[] | undefined | null) => colsOf(v).length > 0;

export const SALES_FIELDS: FieldDef[] = [
  { id: 'date', label: 'Invoice date', required: true, syn: ['invoice date', 'inv date', 'date', 'posting date', 'bill date', 'voucher date', 'document date'] },
  { id: 'no', label: 'Invoice / document no.', required: true, syn: ['invoice number', 'invoice no', 'inv no', 'voucher no', 'bill no', 'document no', 'invoice', 'voucher number', 'doc no', 'number'] },
  { id: 'party', label: 'Customer name', syn: ['customer name', 'customer', 'party name', 'party', 'particulars', 'buyer', 'name', 'billing name'] },
  { id: 'gstin', label: 'Customer GSTIN', syn: ['gstin/uin', 'customer gstin', 'gst identification number (gstin)', 'gstin', 'gst no', 'gst number', 'billing gstin', 'buyer gstin'] },
  { id: 'pos', label: 'Place of supply', syn: ['place of supply', 'pos', 'state'] },
  { id: 'stype', label: 'Supply type', syn: ['supply type', 'gst treatment', 'invoice type', 'gst category', 'type of supply', 'nature'] },
  { id: 'dtype', label: 'Document type (invoice / credit note / debit note)', syn: ['document type', 'voucher type', 'doc type', 'transaction type', 'type'] },
  { id: 'hsn', label: 'HSN / SAC', syn: ['hsn/sac', 'hsn', 'sac', 'hsn code', 'hsn/sac code'] },
  { id: 'uqc', label: 'UQC / unit', syn: ['uqc', 'unit', 'uom', 'usage unit'] },
  { id: 'qty', label: 'Quantity', syn: ['quantity', 'qty'] },
  { id: 'rate', label: 'GST rate %', syn: ['gst rate', 'tax rate', 'rate', 'item tax %', 'gst %', 'tax %', 'rate of tax'] },
  { id: 'taxable', multi: true, label: 'Taxable value', required: true, syn: ['taxable value', 'taxable amount', 'assessable value', 'item total', 'subtotal', 'sub total', 'net amount', 'amount before tax', 'value'] },
  { id: 'cgst', multi: true, label: 'CGST', syn: ['cgst', 'central tax', 'central tax amount', 'cgst amount'] },
  { id: 'sgst', multi: true, label: 'SGST / UTGST', syn: ['sgst', 'state tax', 'state tax amount', 'sgst amount', 'utgst', 'state/ut tax'] },
  { id: 'igst', multi: true, label: 'IGST', syn: ['igst', 'integrated tax', 'integrated tax amount', 'igst amount'] },
  { id: 'cess', multi: true, label: 'Cess', syn: ['cess', 'cess amount', 'compensation cess'] },
  { id: 'gross', label: 'Gross / invoice total (check only)', syn: ['gross total', 'invoice value', 'invoice amount', 'bill amount'] },
];
export const PURCHASE_FIELDS: FieldDef[] = [
  { id: 'date', label: 'Invoice date (supplier)', required: true, syn: ['supplier invoice date', 'bill date', 'invoice date', 'date', 'posting date', 'voucher date', 'document date'] },
  { id: 'no', label: 'Supplier invoice no.', required: true, syn: ['supplier invoice no', 'supplier invoice number', 'bill number', 'bill no', 'bill#', 'invoice number', 'invoice no', 'reference no', 'ref no', 'supplier ref', 'voucher no'] },
  { id: 'party', label: 'Supplier name', syn: ['supplier name', 'vendor name', 'supplier', 'vendor', 'party name', 'party', 'particulars', 'name'] },
  { id: 'gstin', label: 'Supplier GSTIN', syn: ['supplier gstin', 'vendor gstin', 'gstin/uin', 'gst identification number (gstin)', 'gstin', 'gst no', 'gst number'] },
  { id: 'stype', label: 'Supply type (B2B / RCM / import / ISD)', syn: ['supply type', 'gst treatment', 'gst category', 'purchase type', 'reverse charge', 'type of supply'] },
  { id: 'itctype', label: 'ITC type (inputs / capital goods / services)', syn: ['itc type', 'eligibility for itc', 'itc eligibility', 'input type', 'nature of expense'] },
  { id: 'status', label: 'ITC status (availed / not availed / ineligible)', syn: ['itc status', 'itc availed', 'itc claimed', 'eligible'] },
  { id: 'claimfy', label: 'ITC claimed in (FY / period / date)', syn: ['itc claimed in', 'claimed in', 'itc period', 'claim period', 'gstr-3b period', 'return period'] },
  { id: 'hsn', label: 'HSN / SAC', syn: ['hsn/sac', 'hsn', 'sac', 'hsn code'] },
  { id: 'uqc', label: 'UQC / unit', syn: ['uqc', 'unit', 'uom', 'usage unit'] },
  { id: 'qty', label: 'Quantity', syn: ['quantity', 'qty'] },
  { id: 'rate', label: 'GST rate %', syn: ['gst rate', 'tax rate', 'rate', 'item tax %', 'gst %', 'tax %'] },
  { id: 'taxable', multi: true, label: 'Taxable value', required: true, syn: ['taxable value', 'taxable amount', 'assessable value', 'item total', 'subtotal', 'sub total', 'net amount', 'value'] },
  { id: 'cgst', multi: true, label: 'CGST', syn: ['cgst', 'central tax', 'central tax amount', 'cgst amount'] },
  { id: 'sgst', multi: true, label: 'SGST / UTGST', syn: ['sgst', 'state tax', 'state tax amount', 'sgst amount', 'utgst', 'state/ut tax'] },
  { id: 'igst', multi: true, label: 'IGST', syn: ['igst', 'integrated tax', 'integrated tax amount', 'igst amount'] },
  { id: 'cess', multi: true, label: 'Cess', syn: ['cess', 'cess amount', 'compensation cess'] },
  { id: 'gross', label: 'Gross / invoice total (check only)', syn: ['gross total', 'invoice value', 'invoice amount', 'bill amount'] },
];

/** read every sheet of an .xlsx, or a .csv */
export async function readTable(name: string, bytes: Uint8Array): Promise<RegSheet[]> {
  if (name.toLowerCase().endsWith('.csv')) {
    const text = new TextDecoder().decode(bytes).replace(/^﻿/, '');
    const rows: any[][] = []; let row: any[] = [], f = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; }
      else if (ch === '"') q = true;
      else if (ch === ',') { row.push(f); f = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
      else f += ch;
    }
    if (f || row.length) { row.push(f); rows.push(row); }
    return [{ name: 'CSV', rows: rows.map(r => r.map(v => (v !== '' && /^-?[\d,]+(\.\d+)?$/.test(v.trim()) ? +v.replace(/,/g, '') : v))) }];
  }
  const rd = await Reader.open(bytes);
  const out: RegSheet[] = [];
  for (const sn of rd.sheetNames()) {
    const map = await rd.sheet(sn);
    const rows: any[][] = [];
    for (const [ref, c] of map) { const [r, col] = rc(ref); (rows[r - 1] ||= [])[col - 1] = c.v && typeof c.v === 'object' ? null : c.v; }
    const dense = Array.from(rows, r => (r ? Array.from(r, v => v ?? null) : []));
    if (dense.some(r => r.some(v => v !== null && v !== ''))) out.push({ name: sn, rows: dense });
  }
  return out;
}

const clean = (s: any) => String(s ?? '').toLowerCase().replace(/[\s_.]+/g, ' ').replace(/[()₹]/g, '').trim();
/** header row = the row (within the first 40) that matches the most field synonyms */
export function detectHeader(rows: any[][], fields: FieldDef[]) {
  let best = 0, bestScore = -1;
  for (let i = 0; i < Math.min(40, rows.length); i++) {
    const cells = (rows[i] || []).map(clean).filter(Boolean);
    const score = cells.filter(c => fields.some(f => f.syn.some(s => c === s || c.startsWith(s + ' ') || c.endsWith(' ' + s)))).length;
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best;
}
export function guessMapping(headers: any[], fields: FieldDef[]): Record<string, number> {
  const H = headers.map(clean);
  const map: Record<string, number> = {}, used = new Set<number>();
  for (const pass of [0, 1, 2]) for (const f of fields) {
    if (map[f.id] !== undefined) continue;
    for (const s of f.syn) {
      const i = H.findIndex((h, j) => !used.has(j) && h && (pass === 0 ? h === s : pass === 1 ? h.startsWith(s) || h.endsWith(s) : h.includes(s)));
      if (i >= 0) { map[f.id] = i; used.add(i); break; }
    }
  }
  return map;
}

// ---- Tally columnar registers: one column per ledger (Purchase @ 18%, Freight Inward, Input CGST @ 9%, Input SGST @ 9% ...)
const TAX_HEAD: [string, RegExp][] = [['cgst', /\b(cgst|central tax)\b/], ['sgst', /\b(sgst|utgst|state tax|state\/ut tax)\b/], ['igst', /\b(igst|integrated tax)\b/], ['cess', /\bcess\b/]];
/** a column that holds a rate, not an amount ("CGST %", "CGST Rate") */
const RATE_COL = /\brate\b|percent|^%|%$/;
/** ledger names that are value, not tax (purchase / sales / expense ledgers in a columnar export) */
const VALUE_LEDGER = /purchase|sales?\b|expense|freight|carriage|cartage|transport|packing|labour|job ?work|repair|maintenance|rent|professional|consult|service|import|material|stores|spares|consumable|fees|charges|commission|advertis|insurance|hire|works contract|interstate|inter-state|local|exempt|nil rated|zero rated/;
/** columns that must never be added to taxable or tax (totals, round-off, TDS, discounts, identifiers) */
const NOT_VALUE = /gross|total|round|tds|tcs|discount|balance|narration|gstin|invoice|voucher|date|\bno\b|number|particulars|party|supplier|customer|address|state|place|type|hsn|sac|qty|quantity|uqc|unit|rate|%/;
export function taxHeadOf(h: any): string | null {
  const c = clean(h); if (!c || RATE_COL.test(c) && !/@/.test(c)) return null;
  for (const [id, re] of TAX_HEAD) if (re.test(c)) return id;
  return null;
}
/** true when a column holds amounts in the rows below the header */
function numericCol(rows: any[][], header: number, j: number) {
  let n = 0, num = 0;
  for (const r of rows.slice(header + 1, header + 200)) { const v = r?.[j]; if (v === null || v === undefined || v === '') continue; n++; if (typeof v === 'number' || /^\(?-?[\d,]+(\.\d+)?\)?( ?(dr|cr))?$/i.test(String(v).trim())) num++; }
  return n > 0 && num / n >= 0.9;
}
/** Tally columnar register? (several ledger columns for one tax head, or tax-ledger columns but no "taxable value" column) */
export function looksColumnar(headers: any[]) {
  const heads = headers.map(taxHeadOf).filter(Boolean) as string[];
  const dup = new Set(heads).size < heads.length;
  const hasTaxable = headers.map(clean).some(h => /^(taxable|assessable)/.test(h));
  const ledgerish = headers.map(clean).some(h => /@ ?\d/.test(h) || /^input |^output /.test(h));
  return !hasTaxable && (dup || ledgerish);
}
/** single-column guess plus, for columnar registers, every ledger column summed into taxable and each tax head */
export function guessMappingFull(rows: any[][], header: number, fields: FieldDef[]): Mapping {
  const hdr = rows[header] || [];
  const map: Mapping = guessMapping(hdr, fields);
  if (!looksColumnar(hdr)) return map;
  const single = new Set(Object.entries(map).filter(([k]) => !fields.find(f => f.id === k)?.multi).map(([, v]) => v as number));
  const tax: Record<string, number[]> = { cgst: [], sgst: [], igst: [], cess: [] }, val: number[] = [];
  hdr.forEach((h, j) => {
    if (single.has(j) || !numericCol(rows, header, j)) return;
    const head = taxHeadOf(h);
    if (head) { tax[head].push(j); return; }
    const c = clean(h);
    if (c && VALUE_LEDGER.test(c) && !NOT_VALUE.test(c.replace(/@ ?[\d.]+ ?%?/g, '').replace(/\b(purchases?|sales?|local|interstate|inter-state)\b/g, ''))) val.push(j);
  });
  for (const id of ['cgst', 'sgst', 'igst', 'cess']) if (tax[id].length) map[id] = tax[id];
  if (val.length) map.taxable = val;
  return map;
}
/** amount columns (numeric, below the header) that no field uses: in a columnar register these are usually a missed ledger */
export function unusedAmountCols(rows: any[][], header: number, map: Mapping): number[] {
  const used = new Set(Object.values(map).flatMap(colsOf));
  const hdr = rows[header] || [];
  return hdr.map((_, j) => j).filter(j => !used.has(j) && numericCol(rows, header, j) && !/date|\bdt\b|\bno\b|number|voucher|gstin|hsn|sac|qty|quantity|rate|round|\btds\b|\btcs\b|gross|balance|^%|%$/.test(clean(hdr[j])) && !(rows.slice(header + 1, header + 200).every(r => r?.[j] && typeof r[j] === 'number' && r[j] > 20000 && r[j] < 80000)));
}

export interface RegDefaults { noGstinZeroTax: string; itcType: string; status: string; claimFy: string; purNoGstin: string }
const amt = (v: any) => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  let s = String(v).trim(); const neg = /^\(.*\)$/.test(s) || /\bcr\b/i.test(s);
  s = s.replace(/[^\d.-]/g, '');
  const x = parseFloat(s) || 0; return neg ? -Math.abs(x) : x;
};
const GSTIN_RE = /^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;
const MON3 = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
export function parseDate(v: any): string {
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{1,2})[-/ ]([A-Za-z]{3})[A-Za-z]*[-/ ,]*(\d{2,4})$/);
    if (m) { const mi = MON3.indexOf(m[2].toLowerCase()); const y = m[3].length === 2 ? '20' + m[3] : m[3]; if (mi >= 0) return `${m[1].padStart(2, '0')}-${String(mi + 1).padStart(2, '0')}-${y}`; }
    const m2 = v.trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})$/);
    if (m2) return `${m2[1].padStart(2, '0')}-${m2[2].padStart(2, '0')}-20${m2[3]}`;
  }
  return toDate(v);
}
/** dd-mm-yyyy → Excel serial (so Books sheets get real dates) */
const serial = (d: string) => { const m = d.match(/^(\d{2})-(\d{2})-(\d{4})$/); return m ? Math.round((Date.UTC(+m[3], +m[2] - 1, +m[1]) - Date.UTC(1899, 11, 30)) / 86400000) : null; };

function salesType(raw: string, gstin: string, pos: string, tax: number, dflt: string) {
  const s = raw.toLowerCase();
  const codes = ['B2B', 'B2C', 'EXP-WP', 'EXP-WOP', 'SEZ-WP', 'SEZ-WOP', 'DEEMED', 'ADVANCE', 'RCM-SUPPLY', 'ECO-9(5)', 'EXEMPT', 'NIL', 'NON-GST'];
  if (codes.includes(raw.toUpperCase())) return raw.toUpperCase();
  if (s) {
    if (s.includes('export') || s.includes('overseas')) return tax ? 'EXP-WP' : 'EXP-WOP';
    if (s.includes('sez')) return tax ? 'SEZ-WP' : 'SEZ-WOP';
    if (s.includes('deemed')) return 'DEEMED';
    if (s.includes('exempt')) return 'EXEMPT';
    if (s.includes('nil')) return 'NIL';
    if (s.includes('non') && s.includes('gst')) return 'NON-GST';
    if (s.includes('reverse')) return 'RCM-SUPPLY';
    if (s.includes('unreg') || s.includes('consumer') || s.includes('b2c')) return 'B2C';
    if (s.includes('regist') || s.includes('b2b') || s.includes('business')) return 'B2B';
  }
  if (GSTIN_RE.test(gstin)) return 'B2B';
  if (/^(96|97)\b|other country|foreign/i.test(pos)) return tax ? 'EXP-WP' : 'EXP-WOP';
  return tax ? 'B2C' : dflt;
}
function purchaseType(raw: string, gstin: string, tax: number, dflt: string) {
  const s = raw.toLowerCase();
  const codes = ['B2B', 'RCM-REG', 'RCM-UNREG', 'IMPG', 'IMPG-SEZ', 'IMPS', 'SEZ-SERVICES', 'ISD'];
  if (codes.includes(raw.toUpperCase())) return raw.toUpperCase();
  const reg = GSTIN_RE.test(gstin);
  if (s) {
    if (s.includes('import') && (s.includes('service') || s.includes('overseas'))) return 'IMPS';
    if (s.includes('import') || s.includes('boe')) return s.includes('sez') ? 'IMPG-SEZ' : 'IMPG';
    if (s.includes('isd')) return 'ISD';
    if (s.includes('sez')) return 'SEZ-SERVICES';
    if (s.includes('reverse') || s.includes('rcm') || s === 'yes' || s === 'y') return reg ? 'RCM-REG' : 'RCM-UNREG';
  }
  return reg ? 'B2B' : (tax ? dflt : dflt);
}

export interface Normalised { rows: any[][]; stats: { lines: number; docs: number; skipped: number; badDates: number; outsideFy: number; credit: number; grossOff: number; grossOffEg: string[] } }
export function normalise(kind: RegKind, rows: any[][], headerRow: number, map: Mapping, d: RegDefaults, fy: string): Normalised {
  const get = (r: any[], id: string) => { const c = colsOf(map[id]); return c.length ? r[c[0]] : null; };
  /** amount field: sum of every mapped column (one column for ordinary registers, several ledger columns for Tally columnar) */
  const sum = (r: any[], id: string) => colsOf(map[id]).reduce((a, j) => a + amt(r[j]), 0);
  // round-off (signed), TDS (deducted from the party) and TCS (added) columns are allowed for in the gross-total check
  const adjCol = (re: RegExp) => (rows[headerRow] || []).map((h, j) => (re.test(clean(h)) ? j : -1)).filter(j => j >= 0 && !colsOf(map.gross).includes(j));
  const roundCols = adjCol(/round/), tdsCols = adjCol(/\btds\b/), tcsCols = adjCol(/\btcs\b/);
  const out: any[][] = []; const docs = new Set<string>();
  const st = { lines: 0, docs: 0, skipped: 0, badDates: 0, outsideFy: 0, credit: 0, grossOff: 0, grossOffEg: [] as string[] };
  for (const r of rows.slice(headerRow + 1)) {
    if (!r) continue;
    const no = String(get(r, 'no') ?? '').trim();
    let tx = sum(r, 'taxable'), c = sum(r, 'cgst'), s = sum(r, 'sgst'), i = sum(r, 'igst'), cs = sum(r, 'cess');
    [tx, c, s, i, cs] = [tx, c, s, i, cs].map(x => Math.round(x * 100) / 100);
    if (!no || (!tx && !c && !s && !i && !cs)) { if (no || tx) st.skipped++; continue; }
    if (/^(total|grand total|sub ?total)/i.test(no)) { st.skipped++; continue; }
    if (isMapped(map.gross)) {
      const g = amt(get(r, 'gross')), ro = roundCols.reduce((a, j) => a + amt(r[j]), 0) - tdsCols.reduce((a, j) => a + Math.abs(amt(r[j])), 0) + tcsCols.reduce((a, j) => a + Math.abs(amt(r[j])), 0);
      if (g && Math.abs(Math.abs(g) - Math.abs(tx + c + s + i + cs + ro)) > 1) { st.grossOff++; if (st.grossOffEg.length < 3) st.grossOffEg.push(no); }
    }
    const date = parseDate(get(r, 'date'));
    if (!/^\d{2}-\d{2}-\d{4}$/.test(date)) st.badDates++;
    else if (fyOfDate(date) !== fy) st.outsideFy++;
    const rawDt = String(get(r, 'dtype') ?? '').toLowerCase();
    let dt = rawDt.includes('credit') || rawDt === 'cn' || rawDt.includes('return') ? 'CN' : rawDt.includes('debit') || rawDt === 'dn' ? 'DN' : 'INV';
    if (dt === 'INV' && tx < 0) dt = 'CN';
    if (dt === 'CN') st.credit++;
    [tx, c, s, i, cs] = [tx, c, s, i, cs].map(Math.abs);
    const gstin = String(get(r, 'gstin') ?? '').trim().toUpperCase();
    const tax = c + s + i + cs;
    const rate = get(r, 'rate'); const rateN = rate === null || rate === '' ? (tx ? Math.round(tax / tx * 1000) / 10 : null) : amt(rate) < 1 && amt(rate) > 0 ? amt(rate) * 100 : amt(rate);
    const ds = serial(date) ?? date;
    if (kind === 'sales') {
      const stype = salesType(String(get(r, 'stype') ?? ''), gstin, String(get(r, 'pos') ?? ''), tax, d.noGstinZeroTax);
      out.push([ds, no, get(r, 'party') ?? '', gstin, String(get(r, 'pos') ?? '').slice(0, 2), stype, dt, String(get(r, 'hsn') ?? ''), get(r, 'uqc') ?? '', amt(get(r, 'qty')) || null, rateN, tx, c, s, i, cs]);
    } else {
      const ptype = purchaseType(String(get(r, 'stype') ?? ''), gstin, tax, d.purNoGstin);
      const it = String(get(r, 'itctype') ?? '').toLowerCase();
      const itc = it.includes('capital') ? 'Capital Goods' : it.includes('service') ? 'Input Services' : it.includes('input') || it.includes('goods') ? 'Inputs' : d.itcType;
      const sraw = String(get(r, 'status') ?? '').toLowerCase();
      const status = /ineligible|blocked|17\(5\)|17 5/.test(sraw) ? 'Ineligible' : /not|^no$|^n$/.test(sraw) ? 'Not availed' : d.status;
      const craw = get(r, 'claimfy');
      let cfy = d.claimFy;
      if (craw !== null && craw !== '') { const s2 = String(craw).trim(); cfy = /^\d{4}-\d{2}$/.test(s2) ? s2 : /^\d{6}$/.test(s2) ? fyOfDate(`01-${s2.slice(0, 2)}-${s2.slice(2)}`) : fyOfDate(parseDate(craw)) || d.claimFy; }
      out.push([ds, no, get(r, 'party') ?? '', gstin, ptype, itc, status, cfy, String(get(r, 'hsn') ?? ''), get(r, 'uqc') ?? '', amt(get(r, 'qty')) || null, rateN, tx, c, s, i, cs, '']);
    }
    st.lines++; docs.add(`${gstin}|${no}|${dt}`);
  }
  st.docs = docs.size;
  return { rows: out, stats: st };
}

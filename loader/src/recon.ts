import { nameKey } from './suppliers';
// Invoice-level reconciliation: purchase register vs GSTR-2B, sales register vs GSTR-1
import { fyOfDate, r2 } from './util';

const normNo = (s: any) => String(s ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const digits = (s: any) => String(s ?? '').replace(/\D/g, '').replace(/^0+/, '');
const nz = (v: any) => (typeof v === 'number' ? v : parseFloat(String(v ?? '').replace(/,/g, '')) || 0);
const TOL = 1;
const tax = (a: number[]) => r2(a[1] + a[2] + a[3] + a[4]);

/** Excel serial / dd-mm-yyyy / yyyy-mm-dd → dd-mm-yyyy */
export function toDate(v: any): string {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return `${String(d.getUTCDate()).padStart(2, '0')}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${d.getUTCFullYear()}`;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) return `${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}-${m[3]}`;
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  return s;
}

interface Doc { id: string; gstin: string; name: string; no: string; date: string; amt: number[]; extra: Record<string, any>; used?: boolean }
export type Manual = Map<string, { b: Set<string>; r: Set<string> }>;
const merge = (l: Doc[]): Doc => ({ ...l[0], id: l.map(d => d.id).join(','), no: l.map(d => d.no).join(' + '), amt: l.reduce((a, d) => a.map((v, i) => r2(v + d.amt[i])), [0, 0, 0, 0, 0]) });
/** last run of digits (leading zeros dropped) - "INV/26-27/00451" → "451" */
const tailDigits = (s: any) => (String(s ?? '').match(/(\d+)\D*$/)?.[1] ?? '').replace(/^0+/, '');
/** the shorter number is the tail of the longer one, or both end in the same run of digits (3+ digits) */
const suffixMatch = (a: any, b: any) => {
  const x = normNo(a), y = normNo(b);
  if (!x || !y) return false;
  const [sh, lo] = x.length <= y.length ? [x, y] : [y, x];
  if (sh.length >= 4 && lo.endsWith(sh)) return true;
  const ta = tailDigits(a), tb = tailDigits(b);
  return ta.length >= 3 && ta === tb;
};
/** rows given the same Match code are paired first; several documents on a side are added together */
function applyManual(books: Doc[], ret: Doc[], manual: Manual, notes: string[]) {
  const pairs: { b?: Doc; r?: Doc; basis: string; code?: string }[] = [];
  for (const [code, g] of manual) {
    const bs = books.filter(d => !d.used && g.b.has(d.id)), rs = ret.filter(d => !d.used && g.r.has(d.id));
    if (!bs.length || !rs.length) { notes.push(`Match code ${code}: needs at least one row from each side (found ${bs.length} books, ${rs.length} return) - ignored`); continue; }
    [...bs, ...rs].forEach(d => (d.used = true));
    pairs.push({ b: merge(bs), r: merge(rs), basis: `Manual (${code})`, code });
  }
  return pairs;
}
const remarkOf = (remarks: Map<string, string>, id: string) => remarks.get(id) ?? id.split(/[~,]/).map(x => remarks.get(x)).find(Boolean) ?? '';
const group = <T>(list: T[], key: (x: T) => string) => { const m = new Map<string, T[]>(); for (const x of list) { const k = key(x); if (!m.has(k)) m.set(k, []); m.get(k)!.push(x); } return m; };
const sumAmt = (rows: number[][]) => rows.reduce((a, b) => a.map((v, i) => r2(v + (b[i] || 0))), [0, 0, 0, 0, 0]);

/** pair docs: exact key, then digits-only number, then GSTIN+date+amount, then number+amount across GSTINs */
function pair(books: Doc[], ret: Doc[], gstinMatters: boolean) {
  const pairs: { b?: Doc; r?: Doc; basis: string }[] = [];
  const idx = new Map<string, Doc[]>();
  const add = (k: string, d: Doc) => { if (!idx.has(k)) idx.set(k, []); idx.get(k)!.push(d); };
  ret.forEach(d => add(`${gstinMatters ? d.gstin : ''}|${normNo(d.no)}`, d));
  const take = (k: string, ok: (d: Doc) => boolean = () => true) => { const l = idx.get(k); const d = l?.find(x => !x.used && ok(x)); if (d) d.used = true; return d; };
  for (const b of books.filter(x => !x.used)) { const r = take(`${gstinMatters ? b.gstin : ''}|${normNo(b.no)}`); if (r) { b.used = true; pairs.push({ b, r, basis: 'Invoice no.' }); } }
  const free = () => ret.filter(d => !d.used);
  for (const b of books.filter(x => !x.used)) {
    const dg = digits(b.no); if (!dg) continue;
    const r = free().find(d => (!gstinMatters || d.gstin === b.gstin) && digits(d.no) === dg && Math.abs(tax(d.amt) - tax(b.amt)) <= Math.max(TOL, Math.abs(tax(b.amt)) * 0.01));
    if (r) { r.used = b.used = true; pairs.push({ b, r, basis: 'Digits of invoice no.' }); }
  }
  for (const b of books.filter(x => !x.used)) {
    const r = free().find(d => (!gstinMatters || d.gstin === b.gstin) && suffixMatch(d.no, b.no) && Math.abs(tax(d.amt) - tax(b.amt)) <= Math.max(TOL, Math.abs(tax(b.amt)) * 0.01) && Math.abs(d.amt[0] - b.amt[0]) <= Math.max(TOL, Math.abs(b.amt[0]) * 0.01));
    if (r) { r.used = b.used = true; pairs.push({ b, r, basis: 'Last digits of invoice no.' }); }
  }
  for (const b of books.filter(x => !x.used)) {
    const r = free().find(d => d.gstin === b.gstin && d.date === b.date && Math.abs(tax(d.amt) - tax(b.amt)) <= TOL && Math.abs(d.amt[0] - b.amt[0]) <= TOL);
    if (r) { r.used = b.used = true; pairs.push({ b, r, basis: 'GSTIN + date + amount' }); }
  }
  if (gstinMatters) for (const b of books.filter(x => !x.used)) {
    const r = free().find(d => normNo(d.no) === normNo(b.no) && Math.abs(tax(d.amt) - tax(b.amt)) <= TOL);
    if (r) { r.used = b.used = true; pairs.push({ b, r, basis: 'Invoice no. + amount (GSTIN differs)' }); }
  }
  for (const b of books.filter(x => !x.used)) pairs.push({ b, basis: '' });
  for (const r of ret.filter(x => !x.used)) pairs.push({ r, basis: '' });
  return pairs;
}

// ---------------------------------------------------------------- purchases vs 2B
export function reconPurchases(bookRows: any[][], b2bRows: any[][], cy: string, ny: string, remarks: Map<string, string>, manual: Manual = new Map(), notes: string[] = []) {
  // Books_Purchase columns: A date, B no, C supplier, D GSTIN, E type, F ITC type, G status, H claimed FY, ... M taxable, N C, O S, P I, Q cess
  const COMPARED = new Set(['B2B', 'SEZ-SERVICES', 'RCM-REG', 'IMPG', 'IMPG-SEZ']);
  const books: Doc[] = [], notCompared: Doc[] = [];
  const bl = bookRows.filter(r => r && [12, 13, 14, 15, 16].some(i => nz(r[i]) !== 0));
  for (const [, rows] of group(bl, r => `${String(r[3] ?? '').trim().toUpperCase()}|${normNo(r[1])}|${String(r[4] ?? '')}`)) {
    const f = rows[0];
    const d: Doc = { id: `B|${String(f[3] ?? '').trim().toUpperCase()}|${normNo(f[1])}`, gstin: String(f[3] ?? '').trim().toUpperCase(), name: String(f[2] ?? ''), no: String(f[1] ?? ''), date: toDate(f[0]),
      amt: sumAmt(rows.map(r => [nz(r[12]), nz(r[13]), nz(r[14]), nz(r[15]), nz(r[16])])),
      extra: { type: String(f[4] ?? ''), status: String(f[6] ?? ''), claimFy: String(f[7] ?? '') } };
    const zeroTax = Math.abs(tax(d.amt)) < 0.01;
    if (zeroTax) d.extra.why = 'No GST charged - nothing to match in 2B';
    (!zeroTax && (COMPARED.has(d.extra.type) || !d.extra.type || (!d.gstin && d.extra.type === 'RCM-UNREG')) ? books : notCompared).push(d);
  }
  // RAW_2B columns: A key, B-F amounts (T,C,S,I,Cess), G source, H 2B FY, I period, J section, K GSTIN, L name, M doc type, N no, O date, P doc FY, Q avail, R rcm, S ims
  const pool = b2bRows.filter(r => r && ['B2B', 'B2BA', 'CDNR', 'CDNRA', 'IMPG', 'IMPGSEZ'].includes(String(r[9])) &&
    (String(r[7]) === cy || (String(r[7]) === ny && String(r[15]) !== ny)));
  const ret: Doc[] = [];
  for (const [, rows] of group(pool, r => `${String(r[10] ?? '').toUpperCase()}|${normNo(r[13])}`)) {
    const amended = rows.filter(r => String(r[9]).endsWith('A'));
    const use = amended.length ? amended.filter(r => String(r[8]).slice(2) + String(r[8]).slice(0, 2) === amended.map(x => String(x[8]).slice(2) + String(x[8]).slice(0, 2)).sort().pop()) : rows;
    const f = use[0];
    ret.push({ id: `R|${String(f[10] ?? '').toUpperCase()}|${normNo(f[13])}`, gstin: String(f[10] ?? '').toUpperCase(), name: String(f[11] ?? ''), no: String(f[13] ?? ''), date: String(f[14] ?? ''),
      amt: sumAmt(use.map(r => [nz(r[1]), nz(r[2]), nz(r[3]), nz(r[4]), nz(r[5])])),
      extra: { bfy: String(f[7]), period: String(f[8]), section: String(f[9]), avl: String(f[16] ?? ''), rcm: String(f[17] ?? ''), dfy: String(f[15] ?? '') } });
  }
  // books without supplier GSTIN: take it from 2B - by exact (normalised) supplier name, else by invoice no. + amount
  const byName = new Map<string, Set<string>>();
  for (const r of ret) if (r.gstin && r.name) { const k = nameKey(r.name); if (!byName.has(k)) byName.set(k, new Set()); byName.get(k)!.add(r.gstin); }
  let inferred = 0, noGstin = 0;
  for (const b of books) {
    if (b.gstin || /^IMPG/.test(b.extra.type)) continue;
    noGstin++;
    const k = nameKey(b.name);
    let g = k && byName.get(k)?.size === 1 ? [...byName.get(k)!][0] : '';
    let how = 'supplier name';
    // approximate names are not guessed here: the Loader suggests them and the user confirms (suppliers.ts)
    if (!g) {
      const r = ret.find(r => normNo(r.no) === normNo(b.no) && Math.abs(tax(r.amt) - tax(b.amt)) <= TOL);
      if (r) { g = r.gstin; how = 'invoice no. + amount'; }
    }
    if (g) { b.gstin = g; b.id = `B|${g}|${normNo(b.no)}`; b.extra.gstinFrom = how; inferred++; }
  }
  if (noGstin) notes.push(`${noGstin} purchase invoice(s) had no supplier GSTIN in books; ${inferred} were given the GSTIN found in 2B (by supplier name or invoice no. + amount)${noGstin - inferred ? `, ${noGstin - inferred} could not be identified` : ''}. See "Match basis" in Recon_Purchases.`);
  const isImp = (d: Doc) => /^IMPG/.test(d.extra.section || d.extra.type || '');
  // imports have no supplier GSTIN in books: match on BOE number only
  const man = applyManual(books, ret, manual, notes);
  const pairs = [...man, ...pair(books.filter(b => !isImp(b)), ret.filter(r => !isImp(r)), true), ...pair(books.filter(isImp), ret.filter(isImp), false)];
  // supplier total: leftover unmatched documents of a supplier that add up to the same taxable value and tax on both sides
  const leftB = pairs.filter(p => p.b && !p.r), leftR = pairs.filter(p => p.r && !p.b);
  for (const [g, bl] of group(leftB, p => p.b!.gstin)) {
    const rl = leftR.filter(p => p.r!.gstin === g);
    if (!g || !rl.length) continue;
    const B = merge(bl.map(p => p.b!)), R = merge(rl.map(p => p.r!));
    if (Math.abs(tax(B.amt) - tax(R.amt)) <= TOL && Math.abs(B.amt[0] - R.amt[0]) <= TOL) {
      for (const p of [...bl, ...rl]) pairs.splice(pairs.indexOf(p), 1);
      pairs.push({ b: B, r: R, basis: `Supplier total (${bl.length} books / ${rl.length} 2B docs)`, total: true } as any);
    }
  }
  // hints for what is still unmatched: same supplier, invoice no. ends the same (amounts differ)
  const hint = new Map<string, string>();
  const stillR = pairs.filter(p => p.r && !p.b).map(p => p.r!);
  for (const p of pairs.filter(p => p.b && !p.r)) {
    const c = stillR.filter(r => r.gstin === p.b!.gstin && suffixMatch(r.no, p.b!.no));
    if (c.length) {
      const txt = c.map(r => `${r.no} (tax ${tax(r.amt).toLocaleString('en-IN')})`).join(', ');
      hint.set(p.b!.id, `Possible match in 2B: ${txt} - give both rows the same Match code to confirm`);
      c.forEach(r => hint.set(r.id, `Possible match in books: ${p.b!.no} (tax ${tax(p.b!.amt).toLocaleString('en-IN')}) - same Match code to confirm`));
    }
  }
  // otherwise: point out unmatched documents of the same supplier on the other side
  const stillB = pairs.filter(p => p.b && !p.r).map(p => p.b!);
  const sameSup = (list: Doc[], d: Doc, side: string) => {
    const o = list.filter(x => x.gstin && x.gstin === d.gstin);
    return o.length ? `Same supplier has ${o.length} unmatched ${side} doc(s): ${o.slice(0, 3).map(x => `${x.no} (tax ${tax(x.amt).toLocaleString('en-IN')})`).join(', ')}${o.length > 3 ? ', …' : ''} - match manually with a Match code if they are the same` : '';
  };
  for (const b of stillB) if (!hint.has(b.id)) { const t = sameSup(stillR, b, '2B'); if (t) hint.set(b.id, t); }
  for (const r of stillR) if (!hint.has(r.id)) { const t = sameSup(stillB, r, 'books'); if (t) hint.set(r.id, t); }
  const availed = (s: string) => /^(Availed|Reclaimed|Reversed)/.test(s);
  const out: any[][] = [];
  for (const { b, r, basis, code, total } of pairs as { b?: Doc; r?: Doc; basis: string; code?: string; total?: boolean }[]) {
    let status: string, effect = '';
    if (b && r) {
      const dT = Math.abs(b.amt[0] - r.amt[0]), dX = Math.abs(tax(b.amt) - tax(r.amt));
      if (code) status = 'Matched - manual';
      else if (total) status = 'Matched - supplier total';
      else if (basis.includes('GSTIN differs')) status = 'Matched - supplier GSTIN differs';
      else if (basis !== 'Invoice no.') status = 'Matched - invoice no. differs';
      else if (r.extra.bfy === ny) status = 'Matched - in next FY 2B';
      else if (dT > TOL || dX > TOL) status = 'Matched - value difference';
      else status = 'Matched';
      if (total) effect = "Invoice numbers differ but the supplier's unmatched totals agree - correct the numbers in books if needed";
      else if (code && (dT > TOL || dX > TOL)) effect = `Manual match with difference: taxable ${r2(b.amt[0] - r.amt[0])}, tax ${r2(tax(b.amt) - tax(r.amt))} - ITC limited to 2B`;
      else if (r.extra.bfy === ny) effect = 'In next-FY 2B: if availed by 30 Nov next FY → 8C and 13; included in 8A';
      else if (r.extra.dfy && r.extra.dfy !== cy) effect = 'Previous-FY invoice in this FY 2B: if availed this FY → 6A1';
      else if (dX > TOL) effect = 'Tax differs from 2B: ITC limited to 2B; check 8D';
      if (r.extra.avl === 'N') effect = (effect ? effect + '; ' : '') + '2B shows ITC not available';
    } else if (b) {
      status = 'In books, not in 2B';
      effect = hint.get(b.id) ?? (availed(b.extra.status) ? 'ITC availed but invoice not in 2B - not eligible u/s 16(2)(aa): follow up with supplier or reverse' : 'Not availed: no GSTR-9 effect');
    } else {
      status = 'In 2B, not in books';
      effect = hint.get(r!.id) ?? (r!.extra.avl === 'N' ? '2B: ITC not available' : r!.extra.bfy === ny ? 'Next-FY 2B: book and avail by 30 Nov next FY (8C / 13)' : 'In 8A but not availed → 8E (or 8F if ineligible); or book the invoice');
    }
    const B = b?.amt ?? [null, null, null, null, null], R = r?.amt ?? [null, null, null, null, null];
    const id = `${b?.id ?? ''}~${r?.id ?? ''}`;
    const basisTxt = b?.extra.gstinFrom ? `${basis}${basis ? '; ' : ''}GSTIN from 2B (${b.extra.gstinFrom})` : basis;
    out.push([status, basisTxt, (b ?? r)!.gstin, b?.name || r?.name || '', b?.no ?? '', r?.no ?? '', b?.date ?? '', r?.date ?? '',
      ...B, ...R, r2((b?.amt[0] ?? 0) - (r?.amt[0] ?? 0)), r2((b ? tax(b.amt) : 0) - (r ? tax(r.amt) : 0)),
      r?.extra.period ? r.extra.period.slice(0, 2) + '/' + r.extra.period.slice(2) : '', r?.extra.section ?? '', r?.extra.avl ?? '',
      b?.extra.status ?? '', b?.extra.claimFy ?? '', effect, code ?? '', remarkOf(remarks, id), id]);
  }
  for (const d of notCompared) {
    const id = `${d.id}~`;
    out.push(['Not compared (RCM / ISD / imports)', d.extra.why ? 'No GST' : d.extra.type, d.gstin, d.name, d.no, '', d.date, '', ...d.amt, null, null, null, null, null,
      null, null, '', '', '', d.extra.status, d.extra.claimFy, d.extra.why ?? 'Not in 2B scope (RCM from unregistered, import of services, ISD)', '', remarkOf(remarks, id), id]);
  }
  const order = ['In books, not in 2B', 'In 2B, not in books', 'Matched - value difference', 'Matched - invoice no. differs', 'Matched - supplier GSTIN differs', 'Matched - supplier total', 'Matched - in next FY 2B', 'Matched - manual', 'Matched', 'Not compared (RCM / ISD / imports)'];
  out.sort((a, b) => order.indexOf(a[0]) - order.indexOf(b[0]) || String(a[2]).localeCompare(String(b[2])) || String(a[4] || a[5]).localeCompare(String(b[4] || b[5])));
  // supplier summary
  const sup = new Map<string, any>();
  for (const row of out) {
    if (row[0].startsWith('Not compared')) continue;
    const g = row[2] || '(no GSTIN)';
    const s = sup.get(g) || { name: row[3], nb: 0, nr: 0, tb: 0, tr: 0, ob: 0, or: 0, vd: 0 };
    if (row[4]) { s.nb++; s.tb += (row[9] || 0) + (row[10] || 0) + (row[11] || 0) + (row[12] || 0); }
    if (row[5]) { s.nr++; s.tr += (row[14] || 0) + (row[15] || 0) + (row[16] || 0) + (row[17] || 0); }
    if (row[0] === 'In books, not in 2B') s.ob++;
    if (row[0] === 'In 2B, not in books') s.or++;
    if (row[0] === 'Matched - value difference') s.vd++;
    if (!s.name) s.name = row[3];
    sup.set(g, s);
  }
  const supRows = [...sup.entries()].map(([g, s]) => [g, s.name, s.nb, s.nr, r2(s.tb), r2(s.tr), r2(s.tb - s.tr), s.ob, s.or, s.vd, Math.abs(s.tb - s.tr) <= TOL ? 'Yes' : 'No', remarks.get('S|' + g) ?? ''])
    .sort((a: any, b: any) => Math.abs(b[6]) - Math.abs(a[6]));
  const count = (st: string) => out.filter(r => r[0] === st).length;
  return { rows: out, supRows, stats: { matched: out.filter(r => r[0].startsWith('Matched')).length, booksOnly: count('In books, not in 2B'), twoBOnly: count('In 2B, not in books'), valueDiff: count('Matched - value difference') } };
}

// ---------------------------------------------------------------- sales vs GSTR-1
export function reconSales(bookRows: any[][], g1Rows: any[][], cy: string, ny: string, remarks: Map<string, string>, manual: Manual = new Map(), notes: string[] = []) {
  // Books_Sales: A date, B no, C customer, D GSTIN, E POS, F type, G doc type, H HSN, I UQC, J qty, K rate, L taxable, M C, N S, O I, P cess
  const bl = bookRows.filter(r => r && [11, 12, 13, 14, 15].some(i => nz(r[i]) !== 0));
  const cls = (dt: string) => (dt === 'CN' || dt === 'DN') ? 'N' : 'I';
  const books: Doc[] = [];
  for (const [, rows] of group(bl, r => `${cls(String(r[6] ?? ''))}|${normNo(r[1])}`)) {
    const f = rows[0];
    books.push({ id: `B|${cls(String(f[6] ?? ''))}|${normNo(f[1])}`, gstin: String(f[3] ?? '').trim().toUpperCase(), name: String(f[2] ?? ''), no: String(f[1] ?? ''), date: toDate(f[0]),
      amt: sumAmt(rows.map(r => [nz(r[11]), nz(r[12]), nz(r[13]), nz(r[14]), nz(r[15])])).map(Math.abs),
      extra: { type: String(f[5] ?? ''), dt: String(f[6] ?? 'INV') } });
  }
  // RAW_GSTR1: A key, B-F amounts, G return FY, H period, I section, J doc type, K no, L date, M doc FY, N party, O pos, P rate, Q rcm, R file
  const DOCSEC = ['B2B', 'B2BA', 'B2CL', 'B2CLA', 'EXP', 'EXPA', 'CDNR', 'CDNRA', 'CDNUR', 'CDNURA'];
  const pool = g1Rows.filter(r => r && DOCSEC.includes(String(r[8])) && (String(r[6]) === cy || (String(r[6]) === ny && ['10', '11'].includes(String(r[0])))));
  const ret: Doc[] = [];
  for (const [, rows] of group(pool, r => `${cls(String(r[9]))}|${normNo(r[10])}`)) {
    const orig = rows.filter(r => !String(r[8]).endsWith('A'));
    const amd = rows.filter(r => String(r[8]).endsWith('A'));
    let a = sumAmt(orig.map(r => [1, 2, 3, 4, 5].map(i => Math.abs(nz(r[i])))));
    for (const r of amd) { const sg = ['4L', '5K', '11'].includes(String(r[0])) ? -1 : 1; a = a.map((v, i) => r2(v + sg * Math.abs(nz(r[i + 1])))); }
    const f = (orig[0] ?? rows[0]);
    const latest = rows.map(r => String(r[7])).sort((x, y) => (x.slice(2) + x.slice(0, 2)).localeCompare(y.slice(2) + y.slice(0, 2))).pop() || '';
    ret.push({ id: `R|${cls(String(f[9]))}|${normNo(f[10])}`, gstin: String(f[13] ?? '').toUpperCase(), name: '', no: String(f[10] ?? ''), date: String(f[11] ?? ''), amt: a,
      extra: { section: String(f[8]), period: latest, rfy: String(f[6]), dt: String(f[9]) } });
  }
  // documents compared: everything except B2C small / exempt / advances in books; B2C books docs join only if GSTR-1 has them as B2CL
  const SKIP = new Set(['B2C', 'EXEMPT', 'NIL', 'NON-GST', 'ADVANCE']);
  const bMain = books.filter(b => !SKIP.has(b.extra.type));
  const bB2c = books.filter(b => b.extra.type === 'B2C');
  const man = applyManual(books, ret, manual, notes);
  const pairs = pair(bMain, ret, false);
  // B2C books documents reported as B2C large / CDNUR
  const b2cMatches: typeof pairs = [];
  const retLeft = new Map(ret.filter(r => !r.used).map(r => [`${r.extra.dt === 'CN' || r.extra.dt === 'DN' ? 'N' : 'I'}|${normNo(r.no)}`, r]));
  for (const b of bB2c) {
    const r = retLeft.get(`${cls(b.extra.dt)}|${normNo(b.no)}`);
    if (r && !r.used) { r.used = true; b2cMatches.push({ b, r, basis: 'Invoice no. (B2C large)' }); }
  }
  const all: { b?: Doc; r?: Doc; basis: string; code?: string }[] = [...man, ...pairs.filter(p => !(p.r && !p.b && p.r.used)), ...b2cMatches];
  const out: any[][] = [];
  for (const { b, r, basis, code } of all) {
    let status: string, effect = '';
    if (b && r) {
      const dT = Math.abs(b.amt[0] - r.amt[0]), dX = Math.abs(tax(b.amt) - tax(r.amt));
      if (code) { status = 'Matched - manual'; if (dT > TOL || dX > TOL) effect = `Manual match with difference: taxable ${r2(b.amt[0] - r.amt[0])}, tax ${r2(tax(b.amt) - tax(r.amt))}`; }
      else if (r.extra.rfy === ny) { status = 'Matched - reported in next FY'; effect = 'Reported in next-FY GSTR-1 → Table 10 / 11 (and 14 for tax paid)'; }
      else if (b.gstin && r.gstin && b.gstin !== r.gstin) { status = 'Matched - GSTIN differs'; effect = 'Customer GSTIN in GSTR-1 differs from books - amend if wrong'; }
      else if (dT > TOL || dX > TOL) { status = 'Matched - value difference'; effect = 'Value differs: amend through GSTR-1 / Table 10-11, or explain'; }
      else status = 'Matched';
    } else if (b) { status = 'In books, not in GSTR-1'; effect = 'Not reported: report by 30 Nov next FY (Table 10) or pay through DRC-03'; }
    else { status = 'In GSTR-1, not in books'; effect = 'Check books for missing / cancelled document'; }
    const id = `${b?.id ?? ''}~${r?.id ?? ''}`;
    out.push([status, basis, b?.gstin ?? '', r?.gstin ?? '', b?.name ?? '', b?.extra.dt ?? r?.extra.dt ?? '', b?.no ?? '', r?.no ?? '', b?.date ?? '', r?.date ?? '',
      r?.extra.section ?? '', r?.extra.period ? r.extra.period.slice(0, 2) + '/' + r.extra.period.slice(2) : '',
      b ? b.amt[0] : null, b ? tax(b.amt) : null, r ? r.amt[0] : null, r ? tax(r.amt) : null,
      r2((b?.amt[0] ?? 0) - (r?.amt[0] ?? 0)), r2((b ? tax(b.amt) : 0) - (r ? tax(r.amt) : 0)), b?.extra.type ?? '', effect, code ?? '', remarkOf(remarks, id), id]);
  }
  const order = ['In books, not in GSTR-1', 'In GSTR-1, not in books', 'Matched - value difference', 'Matched - GSTIN differs', 'Matched - reported in next FY', 'Matched - manual', 'Matched'];
  out.sort((a, b) => order.indexOf(a[0]) - order.indexOf(b[0]) || String(a[6] || a[7]).localeCompare(String(b[6] || b[7])));

  // ---- month-wise
  const months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].map(m => String(m).padStart(2, '0') + (m >= 4 ? cy.slice(0, 4) : String(+cy.slice(0, 4) + 1)));
  const M = new Map(months.map(m => [m, { bt: 0, gt: 0, bx: 0, gx: 0, bb: 0, gb: 0, be: 0, ge: 0 }]));
  for (const r of bl) {
    const d = toDate(r[0]); if (fyOfDate(d) !== cy) continue;
    const type = String(r[5] ?? ''), dt = String(r[6] ?? 'INV');
    if (type === 'ADVANCE') continue;
    const m = M.get(d.slice(3, 5) + d.slice(6, 10)); if (!m) continue;
    const sg = dt === 'CN' ? -1 : 1;
    const t = Math.abs(nz(r[11])) * sg, x = (Math.abs(nz(r[12])) + Math.abs(nz(r[13])) + Math.abs(nz(r[14])) + Math.abs(nz(r[15]))) * sg;
    if (['EXEMPT', 'NIL', 'NON-GST'].includes(type)) { m.be += t; continue; }
    m.bt += t; m.bx += x; if (type === 'B2C') m.bb += t;
  }
  const NEG = new Set(['4I', '4L', '5H', '5K']);
  for (const r of g1Rows) {
    if (!r || String(r[6]) !== cy) continue;
    const key = String(r[0] ?? ''); if (!key || key === '4F' || key.startsWith('1')) continue;
    const m = M.get(String(r[7])); if (!m) continue;
    const sg = NEG.has(key) ? -1 : 1;
    const t = nz(r[1]) * sg, x = (nz(r[2]) + nz(r[3]) + nz(r[4]) + nz(r[5])) * sg;
    if (['5D', '5E', '5F'].includes(key)) { m.ge += t; continue; }
    m.gt += t; m.gx += x; if (key === '4A') m.gb += t;
  }
  const MN = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'];
  const monthly = months.map((p, i) => { const m = M.get(p)!; return [`${MN[i]} ${p.slice(2)}`, r2(m.bt), r2(m.gt), r2(m.bt - m.gt), r2(m.bx), r2(m.gx), r2(m.bx - m.gx), r2(m.bb), r2(m.gb), r2(m.bb - m.gb), r2(m.be), r2(m.ge), r2(m.be - m.ge), remarks.get('M|' + p) ?? '', ]; });
  const tot = ['Total', ...[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map(j => r2(monthly.reduce((a, r) => a + (r[j] as number), 0))), ''];
  const count = (st: string) => out.filter(r => r[0] === st).length;
  return { rows: out, monthly: [...monthly, tot], months, stats: { matched: out.filter(r => r[0].startsWith('Matched')).length, booksOnly: count('In books, not in GSTR-1'), g1Only: count('In GSTR-1, not in books'), valueDiff: count('Matched - value difference') } };
}

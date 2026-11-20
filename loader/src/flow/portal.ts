// GST portal ledgers (Services → Ledgers), as downloaded in CSV or saved as Excel:
//   Electronic Cash Ledger, Electronic Credit Ledger, Electronic Liability Register (Part I / return related).
// Layout (Nov 2026): a title row, GSTIN / legal name / period rows, then a two-row header.
//   Cash and liability: "<Head> Amount Debited / Credited" groups with Tax / Interest / Penalty / Fee / Others / Total under each, then "<Head> Balance" groups.
//   Credit ledger: "Credit / Debit" and "Balance Available" groups with Integrated / Central / State / CESS / Total under each.
import { COMPONENTS, Component, HEADS, Head, PortalKind, PortalLedger, PortalLine, Ym, r2, zeroHeads } from './types';

const txt = (v: any) => String(v ?? '').replace(/\s+/g, ' ').trim();
const low = (v: any) => txt(v).toLowerCase();
const num = (v: any) => { if (typeof v === 'number') return v; const s = txt(v).replace(/,/g, ''); if (!s || s === '-') return 0; const x = parseFloat(s); return isFinite(x) ? x : 0; };
const headOf = (s: string): Head | null => /integrated|igst/.test(s) ? 'IGST' : /central|cgst/.test(s) ? 'CGST' : /state|sgst|utgst/.test(s) ? 'SGST' : /cess/.test(s) ? 'CESS' : null;
const MON3 = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

/** dd/mm/yyyy, dd-mm-yyyy, Excel serial → dd-mm-yyyy */
export function ddmmyyyy(v: any): string {
  if (typeof v === 'number' && v > 20000 && v < 80000) { const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000); return `${String(d.getUTCDate()).padStart(2, '0')}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${d.getUTCFullYear()}`; }
  if (v instanceof Date) return `${String(v.getDate()).padStart(2, '0')}-${String(v.getMonth() + 1).padStart(2, '0')}-${v.getFullYear()}`;
  const s = txt(v);
  let m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) return `${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}-${m[3]}`;
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  m = s.match(/^(\d{1,2})[-/ ]([A-Za-z]{3})[A-Za-z]*[-/ ,]*(\d{2,4})$/);
  if (m) { const mi = MON3.indexOf(m[2].toLowerCase()); if (mi >= 0) return `${m[1].padStart(2, '0')}-${String(mi + 1).padStart(2, '0')}-${m[3].length === 2 ? '20' + m[3] : m[3]}`; }
  return '';
}
export const ymOfDate = (d: string): Ym => { const m = d.match(/^(\d{2})-(\d{2})-(\d{4})$/); return m ? `${m[3]}-${m[2]}` : ''; };
/** "Jan-26", "January 2026", "012026", "01/2026" → 2026-01 */
export function ymOfPeriod(v: any): Ym | '' {
  const s = low(v);
  let m = s.match(/^([a-z]{3})[a-z]*[-/ ']*(\d{2}|\d{4})$/);
  if (m) { const mi = MON3.indexOf(m[1]); if (mi >= 0) return `${m[2].length === 2 ? '20' + m[2] : m[2]}-${String(mi + 1).padStart(2, '0')}`; }
  m = s.match(/^(\d{2})[-/]?(\d{4})$/);
  if (m && +m[1] >= 1 && +m[1] <= 12) return `${m[2]}-${m[1]}`;
  return '';
}
/** return ARNs (AA + state code + MMYY + …) carry the tax period */
export const ymOfArn = (ref: string): Ym | '' => { const m = ref.match(/^AA\d{2}(\d{2})(\d{2})/); return m && +m[1] >= 1 && +m[1] <= 12 ? `20${m[2]}-${m[1]}` : ''; };

export function portalKindOf(rows: any[][]): PortalKind | null {
  const top = rows.slice(0, 4).map(r => (r || []).map(low).join(' ')).join(' ');
  if (/electronic cash ledger/.test(top)) return 'cash';
  if (/electronic credit ledger/.test(top)) return 'credit';
  if (/electronic liability register/.test(top)) return 'liability';
  return null;
}

const emptyAmt = () => Object.fromEntries(HEADS.map(h => [h, Object.fromEntries(COMPONENTS.map(c => [c, 0]))])) as Record<Head, Record<Component, number>>;

export function parsePortal(rows: any[][]): PortalLedger {
  const kind = portalKindOf(rows);
  if (!kind) throw new Error('Not a GST portal ledger');
  const res: PortalLedger = { kind, gstin: '', legalName: '', from: '', to: '', lines: [], problems: [] };
  // meta: label in one cell, value in the next
  let hi = -1;
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const r = (rows[i] || []).map(txt);
    if (r.some(c => /^sr\.?\s*no/i.test(c))) { hi = i; break; }
    r.forEach((c, j) => {
      const v = r[j + 1] || '', lc = c.toLowerCase();
      if (lc === 'gstin') res.gstin = v.toUpperCase();
      else if (/^(legal name|name\s*\(legal\))/.test(lc)) res.legalName = v;
      else if (lc === 'from') res.from = ddmmyyyy(v);
      else if (lc === 'to') res.to = ddmmyyyy(v);
      else if (lc === 'tax period') {
        const [a, b] = v.split(/\s+to\s+/i).map(ymOfPeriod);
        if (a) res.from = `01-${a.slice(5)}-${a.slice(0, 4)}`;
        if (b) { const last = new Date(Date.UTC(+b.slice(0, 4), +b.slice(5), 0)).getUTCDate(); res.to = `${last}-${b.slice(5)}-${b.slice(0, 4)}`; }
      }
    });
  }
  if (hi < 0) { res.problems.push('Header row (Sr.No, Date, …) not found'); return res; }
  const h1 = (rows[hi] || []).map(txt), h2 = (rows[hi + 1] || []).map(txt);
  const width = Math.max(h1.length, h2.length);
  const find = (re: RegExp) => h1.findIndex(c => re.test(c.toLowerCase()));
  const cDate = find(/date/), cRef = find(/reference/), cPer = find(/tax period/), cDesc = find(/description/), cType = find(/transaction type/), cUsed = find(/ledger used/);
  // amount and balance columns
  const amtCol: { head: Head; comp: Component; col: number }[] = [], balCol: { head: Head; col: number }[] = [];
  if (h2.some(c => /^tax$/i.test(c))) {
    let group = '';
    for (let j = 0; j < width; j++) {
      if (h1[j]) group = h1[j].toLowerCase();
      const sub = (h2[j] || '').toLowerCase(), head = headOf(group);
      if (!head || !sub) continue;
      if (/balance/.test(group)) { if (sub === 'total') balCol.push({ head, col: j }); }
      else if ((COMPONENTS as readonly string[]).includes(sub)) amtCol.push({ head, comp: sub as Component, col: j });
    }
  } else {
    // credit ledger: first run of head columns = credit / debit, second = balance
    const seen = new Set<Head>();
    for (let j = 0; j < width; j++) {
      const head = headOf((h2[j] || '').toLowerCase()); if (!head) continue;
      if (!seen.has(head)) { seen.add(head); amtCol.push({ head, comp: 'tax', col: j }); } else balCol.push({ head, col: j });
    }
  }
  if (!amtCol.length) res.problems.push('Amount columns not recognised');
  let lastArn: Ym | '' = '';
  for (let i = hi + 2; i < rows.length; i++) {
    const r = rows[i] || [];
    if (!r.some(v => txt(v))) continue;
    const desc = txt(r[cDesc]);
    const bal = zeroHeads(); balCol.forEach(b => (bal[b.head] = r2(num(r[b.col]))));
    if (/^opening balance/i.test(desc)) { res.opening = bal; continue; }
    if (/^closing balance/i.test(desc)) { res.closing = bal; continue; }
    const date = ddmmyyyy(r[cDate]);
    const t = low(r[cType]);
    if (!date || !/debit|credit|\bdr\b|\bcr\b/.test(t)) continue;
    const amt = emptyAmt();
    amtCol.forEach(a => (amt[a.head][a.comp] = r2(num(r[a.col]))));
    const ref = txt(r[cRef]);
    let taxPeriod: Ym | '' = cPer >= 0 ? ymOfPeriod(r[cPer]) : '';
    if (!taxPeriod && kind === 'liability') { const a = ymOfArn(ref); if (a) lastArn = a; taxPeriod = a || lastArn; }
    res.lines.push({ row: i + 1, sr: txt(r[0]), date, ref, taxPeriod, desc, dir: /debit|\bdr\b/.test(t) ? 'Dr' : 'Cr', used: cUsed >= 0 ? txt(r[cUsed]) : '', amt, bal });
  }
  // the file's own running balance must agree with its lines
  if (res.lines.length) checkRunning(res);
  return res;
}

/** effect of a line on the ledger's balance. Cash and credit ledgers: credit adds. Liability register: debit (liability created) adds */
export const lineSign = (kind: PortalKind, dir: 'Dr' | 'Cr') => (kind === 'liability' ? (dir === 'Dr' ? 1 : -1) : dir === 'Cr' ? 1 : -1);
export const lineTotal = (l: PortalLine, h: Head) => COMPONENTS.reduce((a, c) => a + l.amt[h][c], 0);

function checkRunning(p: PortalLedger) {
  const run = { ...(p.opening || zeroHeads()) };
  let bad = 0;
  for (const l of p.lines) {
    for (const h of HEADS) run[h] = r2(run[h] + lineSign(p.kind, l.dir) * lineTotal(l, h));
    if (HEADS.some(h => Math.abs(run[h] - l.bal[h]) > 1)) bad++;
    Object.assign(run, l.bal);
  }
  if (bad) p.problems.push(`${bad} line(s) where the running balance does not follow from the amounts`);
  if (p.closing && HEADS.some(h => Math.abs(run[h] - p.closing![h]) > 1)) p.problems.push('Closing balance does not follow from the lines');
}

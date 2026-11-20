// Supplier name matching: books rows without a GSTIN → GSTIN from GSTR-2B, by approximate supplier name.
// Exact (normalised) names are applied automatically; approximate ones are suggested and applied only when the user confirms.

/** normalised supplier name: upper case, "&" → AND, without M/s, Pvt, Ltd, LLP, Co. and punctuation */
export const nameKey = (n: any) =>
  String(n ?? '')
    .toUpperCase()
    .replace(/&/g, ' AND ')
    .replace(/\bM\/S\b/g, ' ')
    .replace(/[^A-Z0-9 ]/g, ' ')
    .split(/\s+/)
    .filter(w => w && !['PRIVATE', 'PVT', 'LIMITED', 'LTD', 'LLP', 'CO', 'COMPANY', 'THE', 'MS', 'P', 'AND'].includes(w))
    .join(' ');
const normNo = (s: any) =>
  String(s ?? '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');

export interface Supplier2B {
  gstin: string;
  name: string;
  nos: Set<string>;
}
/** suppliers (GSTIN, trade name, document numbers) in GSTR-2B JSON files */
export function suppliersFrom2B(jsons: any[]): Map<string, Supplier2B> {
  const out = new Map<string, Supplier2B>();
  const add = (s: any, no: any) => {
    const g = String(s?.ctin ?? '').toUpperCase();
    if (!g) return;
    if (!out.has(g)) out.set(g, { gstin: g, name: String(s?.trdnm ?? ''), nos: new Set() });
    const e = out.get(g)!;
    if (!e.name && s?.trdnm) e.name = String(s.trdnm);
    if (no) e.nos.add(normNo(no));
  };
  for (const j of jsons) {
    const dd = (j?.data ?? j)?.docdata || {};
    for (const sec of ['b2b', 'b2ba']) for (const s of dd[sec] || []) for (const inv of s.inv || []) add(s, inv.inum);
    for (const sec of ['cdnr', 'cdnra'])
      for (const s of dd[sec] || []) for (const nt of s.nt || []) add(s, nt.ntnum ?? nt.nt_num);
  }
  return out;
}

const bigrams = (s: string) => {
  const t = s.replace(/ /g, '');
  const m = new Map<string, number>();
  for (let i = 0; i < t.length - 1; i++) {
    const b = t.slice(i, i + 2);
    m.set(b, (m.get(b) || 0) + 1);
  }
  return m;
};
function dice(a: string, b: string) {
  const A = bigrams(a),
    B = bigrams(b);
  let inter = 0,
    na = 0,
    nb = 0;
  A.forEach(v => (na += v));
  B.forEach(v => (nb += v));
  A.forEach((v, k) => (inter += Math.min(v, B.get(k) || 0)));
  return na + nb ? (2 * inter) / (na + nb) : 0;
}
function lev(a: string, b: string) {
  const d = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let p = d[0];
    d[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const t = d[j];
      d[j] = Math.min(d[j] + 1, d[j - 1] + 1, p + (a[i - 1] === b[j - 1] ? 0 : 1));
      p = t;
    }
  }
  return d[b.length];
}
const tokEq = (x: string, y: string) =>
  x === y || (x.length >= 4 && y.length >= 4 && (lev(x, y) <= 1 || x.startsWith(y) || y.startsWith(x)));
/** 0..1 similarity of two normalised names: best of token overlap (typo-tolerant), character bigrams and initials ("SST" ~ "SHARMA STEEL TRADERS") */
export function nameScore(a: string, b: string): number {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const ta = a.split(' '),
    tb = b.split(' ');
  const hit = ta.filter(x => tb.some(y => tokEq(x, y))).length;
  const tok = hit / Math.max(ta.length, tb.length);
  const ini = (t: string[]) => t.map(w => w[0]).join('');
  const initials =
    (ta.length === 1 && ta[0].length >= 2 && ta[0] === ini(tb)) || (tb.length === 1 && tb[0].length >= 2 && tb[0] === ini(ta))
      ? 0.8
      : 0;
  const contains = a.length >= 5 && b.length >= 5 && (a.includes(b) || b.includes(a)) ? 0.85 : 0;
  // never 100% unless the names are the same: a near-identical name still needs the user's confirmation
  return Math.min(0.95, Math.round(Math.max(tok, dice(a, b), initials, contains) * 100) / 100);
}

export interface Candidate {
  gstin: string;
  name: string;
  score: number;
  invHits: number;
}
export interface Suggestion {
  key: string;
  name: string;
  docs: number;
  tax: number;
  exact?: string;
  cands: Candidate[];
}
/**
 * books suppliers without a GSTIN, with 2B candidates.
 * rows: Books_Purchase rows (B no, C name, D GSTIN, N..Q tax). Names that match exactly (normalised, one GSTIN) get `exact`.
 */
export function supplierSuggestions(rows: any[][], sup: Map<string, Supplier2B>, minScore = 0.55): Suggestion[] {
  const by = new Map<string, { name: string; nos: Set<string>; tax: number }>();
  for (const r of rows) {
    if (String(r[3] ?? '').trim()) continue;
    if (/^IMP/.test(String(r[4] ?? ''))) continue;
    const k = nameKey(r[2]);
    if (!k) continue;
    if (!by.has(k)) by.set(k, { name: String(r[2]), nos: new Set(), tax: 0 });
    const e = by.get(k)!;
    e.nos.add(normNo(r[1]));
    e.tax += [13, 14, 15, 16].reduce((a, i) => a + (+r[i] || 0), 0);
  }
  const keyed = [...sup.values()].map(s => ({ ...s, key: nameKey(s.name) }));
  const out: Suggestion[] = [];
  for (const [k, e] of by) {
    const same = keyed.filter(s => s.key === k);
    const cands = keyed
      .map(s => {
        const invHits = [...e.nos].filter(n => n && s.nos.has(n)).length;
        return { gstin: s.gstin, name: s.name, score: nameScore(k, s.key), invHits };
      })
      .filter(c => c.score >= minScore || c.invHits > 0)
      .sort((x, y) => (y.invHits > 0 ? 1 : 0) - (x.invHits > 0 ? 1 : 0) || y.score - x.score || y.invHits - x.invHits)
      .slice(0, 5);
    out.push({
      key: k,
      name: e.name,
      docs: e.nos.size,
      tax: Math.round(e.tax * 100) / 100,
      exact: same.length === 1 ? same[0].gstin : undefined,
      cands,
    });
  }
  return out.sort((a, b) => b.tax - a.tax);
}

/** fill the GSTIN (and B2B type) into books rows whose supplier name the user confirmed: aliases = nameKey → GSTIN */
export function applyAliases(rows: any[][], aliases: Map<string, string>): number {
  let n = 0;
  for (const r of rows) {
    if (String(r[3] ?? '').trim()) continue;
    const g = aliases.get(nameKey(r[2]));
    if (!g) continue;
    r[3] = g;
    if (['RCM-UNREG', 'B2B', ''].includes(String(r[4] ?? ''))) r[4] = 'B2B';
    r[17] = (r[17] ? r[17] + '; ' : '') + 'GSTIN from 2B (supplier name confirmed in Loader)';
    n++;
  }
  return n;
}

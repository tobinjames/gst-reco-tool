// Customer names: books entries without a customer GSTIN → GSTIN from GSTR-1, the way Input Reco does it for suppliers
// (src/suppliers.ts nameScore / nameKey). GSTR-1 JSON carries only the customer GSTIN, no name, so names are learnt from the
// books: entries of the same customer that do carry the GSTIN, and invoices whose number is in GSTR-1 under that GSTIN.
// Exact names (one GSTIN) are applied automatically; the rest are suggested and applied only when the user confirms,
// and the user can pick any GSTR-1 customer by hand.
import { nameKey, nameScore } from '../suppliers';
import { normNo } from '../recon';
import type { RawRow } from '../util';

export interface G1Customer {
  gstin: string;
  /** names this GSTIN carries in the books (entries with the GSTIN, or invoices found under it in GSTR-1) */
  names: Set<string>;
  nos: Set<string>;
  docs: number;
  tx: number;
}
const DOC = /^(B2B|B2BA|CDNR|CDNRA)$/;

/** customers (GSTINs) in GSTR-1 with their document numbers; names learnt from the books rows (Books_Sales layout) */
export function customersFromG1(
  g1: RawRow[],
  bookRows: any[][],
  list: { name: string; gstin: string }[] = [],
): Map<string, G1Customer> {
  const out = new Map<string, G1Customer>();
  for (const r of g1) {
    const sec = String(r.info[2]),
      g = String(r.info[7] ?? '').toUpperCase();
    if (!DOC.test(sec) || !/^\d{2}[A-Z0-9]{13}$/.test(g)) continue;
    let c = out.get(g);
    if (!c) out.set(g, (c = { gstin: g, names: new Set(), nos: new Set(), docs: 0, tx: 0 }));
    const no = normNo(r.info[4]);
    if (no && !c.nos.has(no)) c.docs++;
    if (no) c.nos.add(no);
    c.tx += r.amt[0];
  }
  for (const x of list) out.get(x.gstin)?.names.add(x.name);
  for (const r of bookRows) {
    const g = String(r[3] ?? '')
      .trim()
      .toUpperCase();
    const k = nameKey(r[2]);
    if (!k) continue;
    if (g && out.has(g)) out.get(g)!.names.add(String(r[2]));
    else if (!g) {
      const no = normNo(r[1]);
      const hits = [...out.values()].filter(c => no && c.nos.has(no));
      if (hits.length === 1) hits[0].names.add(String(r[2]));
    }
  }
  return out;
}

export interface CustCandidate {
  gstin: string;
  name: string;
  score: number;
  invHits: number;
}
export interface CustSuggestion {
  key: string;
  name: string;
  docs: number;
  tx: number;
  /** exact (normalised) name of one GSTIN that carries it in the books: applied without asking */
  exact?: string;
  cands: CustCandidate[];
  /** invoice numbers found under the top candidate (2+), or 1 with a similar name */
  strong: boolean;
}

/** books customers without a GSTIN (taxable sales), with GSTR-1 candidates by name and invoice numbers */
export function customerSuggestions(bookRows: any[][], dir: Map<string, G1Customer>, minScore = 0.55): CustSuggestion[] {
  const by = new Map<string, { name: string; nos: Set<string>; tx: number }>();
  for (const r of bookRows) {
    if (String(r[3] ?? '').trim()) continue;
    if (['EXEMPT', 'NIL', 'NON-GST', 'NOSUPPLY', 'ADVANCE', 'EXP-WP', 'EXP-WOP'].includes(String(r[5] ?? ''))) continue;
    const k = nameKey(r[2]);
    if (!k) continue;
    if (!by.has(k)) by.set(k, { name: String(r[2]), nos: new Set(), tx: 0 });
    const e = by.get(k)!;
    e.nos.add(normNo(r[1]));
    e.tx += Math.abs(+r[11] || 0);
  }
  const keyed = [...dir.values()].map(c => ({ ...c, keys: [...c.names].map(nameKey) }));
  const out: CustSuggestion[] = [];
  for (const [k, e] of by) {
    const same = keyed.filter(c => c.keys.includes(k));
    const cands = keyed
      .map(c => ({
        gstin: c.gstin,
        name: [...c.names][0] ?? '',
        score: Math.max(0, ...c.keys.map(x => nameScore(k, x))),
        invHits: [...e.nos].filter(n => n && c.nos.has(n)).length,
      }))
      .filter(c => c.score >= minScore || c.invHits > 0)
      .sort((x, y) => y.invHits - x.invHits || y.score - x.score)
      .slice(0, 5);
    const top = cands[0];
    out.push({
      key: k,
      name: e.name,
      docs: e.nos.size,
      tx: Math.round(e.tx * 100) / 100,
      exact: same.length === 1 ? same[0].gstin : undefined,
      cands,
      strong: !!top && (top.invHits >= 2 || (top.invHits >= 1 && top.score >= minScore)),
    });
  }
  return out.sort((a, b) => b.tx - a.tx);
}

/** fill the GSTIN into books rows of a customer name the user confirmed (aliases: nameKey → GSTIN; '' = unregistered, B2C) */
export function applyCustomerAliases(rows: any[][], aliases: Map<string, string>): any[][] {
  return rows.map(r => {
    if (String(r[3] ?? '').trim()) return r;
    const g = aliases.get(nameKey(r[2]));
    if (g === undefined) return r;
    const x = [...r];
    if (g) {
      x[3] = g;
      if (['B2C', 'B2B', ''].includes(String(x[5] ?? ''))) x[5] = 'B2B';
    } else if (String(x[5]) === 'B2B') x[5] = 'B2C';
    return x;
  });
}

// ---------------------------------------------------------------- customer list file (name + GSTIN)
const GSTIN_RE = /^\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]$/;
const hkey = (v: any) =>
  String(v ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9/ ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
/** a customer master / list: a GSTIN column and a name column, and no invoice / voucher / tax columns (balances, credit limit etc. are fine). null = not one */
export function parseCustomerList(rows: any[][]): { name: string; gstin: string }[] | null {
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const H = (rows[i] || []).map(hkey);
    const g = H.findIndex(h =>
      /^(gstin( ?\/ ?uin)?|gst ?(no|number|in)|gst identification number gstin|gstin of (customer|party|buyer)|customer gstin|party gstin|gstin number)$/.test(
        h,
      ),
    );
    if (g < 0) continue;
    if (H.some(h => /^(date|.* date)$|invoice|taxable|hsn|voucher|cgst|sgst|igst|bill no/.test(h))) return null;
    const pref = [
      /^(customer|party|ledger|account|buyer|client)( name)?$/,
      /^(display name|company name|name|legal name|trade name|particulars)$/,
    ];
    let nm = -1;
    for (const re of pref) {
      nm = H.findIndex((h, j) => j !== g && re.test(h));
      if (nm >= 0) break;
    }
    if (nm < 0) continue;
    const out: { name: string; gstin: string }[] = [];
    for (const r of rows.slice(i + 1)) {
      const gs = String(r?.[g] ?? '')
        .trim()
        .toUpperCase();
      const name = String(r?.[nm] ?? '').trim();
      if (name && GSTIN_RE.test(gs)) out.push({ name, gstin: gs });
    }
    return out.length ? out : null;
  }
  return null;
}

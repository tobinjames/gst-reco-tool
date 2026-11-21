// Customer names: books entries without a customer GSTIN → GSTIN from GSTR-1, the way Input Reco does it for suppliers
// (src/suppliers.ts nameScore / nameKey). GSTR-1 JSON carries only the customer GSTIN, no name, so names are learnt from the
// books (entries of the same customer that carry the GSTIN) and the customer list. A books customer without a GSTIN gets one
// by the methods chosen under Matching settings: customer list, same name in books, or invoice no. AND amount agreeing with
// GSTR-1 documents of one GSTIN. Each mapping keeps its basis (shown on screen and in Customer_Map); similar names are only
// suggested, and the user can pick any GSTR-1 customer by hand.
import { nameKey, nameScore } from '../suppliers';
import { normNo } from '../recon';
import type { RawRow } from '../util';

export interface G1Customer {
  gstin: string;
  /** names shown for this GSTIN: from the customer list and from books entries carrying it */
  names: Set<string>;
  /** names (nameKey) that carry this GSTIN on books entries: the basis of "same name in books" */
  bookKeys: Set<string>;
  nos: Set<string>;
  /** GSTR-1 documents by invoice no. (normalised): number as reported, taxable value and tax (absolute) */
  inv: Map<string, { no: string; tx: number; tax: number }>;
  docs: number;
  tx: number;
}
const DOC = /^(B2B|B2BA|CDNR|CDNRA)$/;
const r2 = (x: number) => Math.round(x * 100) / 100;

/** customers (GSTINs) in GSTR-1 with their documents; names from the customer list and the books rows that carry a GSTIN */
export function customersFromG1(
  g1: RawRow[],
  bookRows: any[][],
  list: { name: string; gstin: string }[] = [],
  tol = 1,
): Map<string, G1Customer> {
  const out = new Map<string, G1Customer>();
  for (const r of g1) {
    const sec = String(r.info[2]),
      g = String(r.info[7] ?? '').toUpperCase();
    if (!DOC.test(sec) || !/^\d{2}[A-Z0-9]{13}$/.test(g)) continue;
    let c = out.get(g);
    if (!c) out.set(g, (c = { gstin: g, names: new Set(), bookKeys: new Set(), nos: new Set(), inv: new Map(), docs: 0, tx: 0 }));
    const no = normNo(r.info[4]);
    if (no && !c.nos.has(no)) c.docs++;
    if (no) {
      c.nos.add(no);
      const d = c.inv.get(no) ?? { no: String(r.info[4] ?? ''), tx: 0, tax: 0 };
      // amendments replace the original in GSTR-1; for matching a number, the latest figures are what the customer was billed
      if (/A$/.test(sec)) {
        d.tx = 0;
        d.tax = 0;
      }
      d.tx = r2(d.tx + Math.abs(r.amt[0] || 0));
      d.tax = r2(d.tax + Math.abs(r.amt[1] || 0) + Math.abs(r.amt[2] || 0) + Math.abs(r.amt[3] || 0) + Math.abs(r.amt[4] || 0));
      c.inv.set(no, d);
    }
    c.tx += r.amt[0];
  }
  for (const x of list) out.get(x.gstin)?.names.add(x.name);
  const noG = new Map<string, { name: string; no: string; tx: number; tax: number }>();
  for (const r of bookRows) {
    const g = String(r[3] ?? '')
      .trim()
      .toUpperCase();
    const k = nameKey(r[2]);
    if (!k) continue;
    if (g && out.has(g)) {
      out.get(g)!.names.add(String(r[2]));
      out.get(g)!.bookKeys.add(k);
    } else if (!g && normNo(r[1])) {
      const kk = `${k}|${normNo(r[1])}`;
      const d = noG.get(kk) ?? { name: String(r[2]), no: normNo(r[1]), tx: 0, tax: 0 };
      d.tx += Math.abs(+r[11] || 0);
      d.tax += [12, 13, 14, 15].reduce((a, i) => a + Math.abs(+r[i] || 0), 0);
      noG.set(kk, d);
    }
  }
  // a name is also learnt (to suggest similar names) from invoices whose number and amount agree with one GSTIN's document
  for (const d of noG.values()) {
    const hits = [...out.values()].filter(c => {
      const x = c.inv.get(d.no);
      return x && Math.abs(x.tx - d.tx) <= tol && Math.abs(x.tax - d.tax) <= tol;
    });
    if (hits.length === 1) hits[0].names.add(d.name);
  }
  return out;
}

export interface CustCandidate {
  gstin: string;
  name: string;
  score: number;
  /** invoices whose number AND amount agree with a GSTR-1 document of this GSTIN */
  invHits: number;
  /** invoices whose number agrees but the amount does not */
  noHits: number;
  /** the agreeing invoices, for the basis shown to the user */
  hits: string[];
}
export interface CustSuggestion {
  key: string;
  name: string;
  docs: number;
  tx: number;
  /** the name carries exactly one GSTIN on other books entries: "same name in books" */
  exact?: string;
  /** invoice no. + amount agree with GSTR-1 documents of exactly one GSTIN (none of another) */
  inv?: { gstin: string; hits: number; of: number; sample: string[] };
  cands: CustCandidate[];
  /** a candidate with invoices agreeing by number and amount */
  strong: boolean;
}
export interface CustMatchOpts {
  /** suggest names at least this similar (0-1) */
  minScore: number;
  /** amounts within this many rupees agree */
  tol: number;
  /** 'tax': only tax must agree; 'both': taxable value and tax */
  compare: 'tax' | 'both';
}
const inr0 = (v: number) => v.toLocaleString('en-IN', { maximumFractionDigits: 2 });

/** books customers without a GSTIN (taxable sales), with GSTR-1 candidates by name and by invoice no. + amount */
export function customerSuggestions(
  bookRows: any[][],
  dir: Map<string, G1Customer>,
  o: Partial<CustMatchOpts> = {},
): CustSuggestion[] {
  const minScore = o.minScore ?? 0.55,
    tol = o.tol ?? 1,
    compare = o.compare ?? 'both';
  const by = new Map<string, { name: string; inv: Map<string, { no: string; tx: number; tax: number }>; tx: number }>();
  for (const r of bookRows) {
    if (String(r[3] ?? '').trim()) continue;
    if (['EXEMPT', 'NIL', 'NON-GST', 'NOSUPPLY', 'ADVANCE', 'EXP-WP', 'EXP-WOP'].includes(String(r[5] ?? ''))) continue;
    const k = nameKey(r[2]);
    if (!k) continue;
    if (!by.has(k)) by.set(k, { name: String(r[2]), inv: new Map(), tx: 0 });
    const e = by.get(k)!;
    const no = normNo(r[1]);
    const tx = Math.abs(+r[11] || 0),
      tax = [12, 13, 14, 15].reduce((a, i) => a + Math.abs(+r[i] || 0), 0);
    if (no) {
      const d = e.inv.get(no) ?? { no: String(r[1] ?? ''), tx: 0, tax: 0 };
      d.tx = r2(d.tx + tx);
      d.tax = r2(d.tax + tax);
      e.inv.set(no, d);
    }
    e.tx += tx;
  }
  const agree = (a: { tx: number; tax: number }, b: { tx: number; tax: number }) =>
    Math.abs(a.tax - b.tax) <= tol && (compare === 'tax' || Math.abs(a.tx - b.tx) <= tol);
  const keyed = [...dir.values()].map(c => ({ ...c, keys: [...c.names].map(nameKey) }));
  const out: CustSuggestion[] = [];
  for (const [k, e] of by) {
    const same = keyed.filter(c => c.bookKeys.has(k));
    const cands = keyed
      .map(c => {
        let invHits = 0,
          noHits = 0;
        const hits: string[] = [];
        for (const [no, b] of e.inv) {
          const g = c.inv.get(no);
          if (!g) continue;
          if (agree(b, g)) {
            invHits++;
            hits.push(`${b.no} (₹${inr0(b.tx)} + tax ₹${inr0(b.tax)})`);
          } else noHits++;
        }
        return {
          gstin: c.gstin,
          name: [...c.names][0] ?? '',
          score: Math.max(0, ...c.keys.map(x => nameScore(k, x))),
          invHits,
          noHits,
          hits,
        };
      })
      .filter(c => c.score >= minScore || c.invHits > 0 || c.noHits > 0)
      .sort((x, y) => y.invHits - x.invHits || y.noHits - x.noHits || y.score - x.score)
      .slice(0, 5);
    const withHits = cands.filter(c => c.invHits > 0);
    const top = cands[0];
    out.push({
      key: k,
      name: e.name,
      docs: e.inv.size,
      tx: r2(e.tx),
      exact: same.length === 1 ? same[0].gstin : undefined,
      inv:
        withHits.length === 1
          ? { gstin: withHits[0].gstin, hits: withHits[0].invHits, of: e.inv.size, sample: withHits[0].hits.slice(0, 5) }
          : undefined,
      cands,
      strong: !!top && top.invHits > 0 && withHits.length === 1,
    });
  }
  return out.sort((a, b) => b.tx - a.tx);
}

export type CustHow = 'user' | 'list' | 'exact' | 'invamt';
export const CUST_HOW: Record<CustHow, string> = {
  user: 'Your choice',
  list: 'Customer list',
  exact: 'Same name in books',
  invamt: 'Invoice no. + amount',
};
export interface CustMapping {
  /** '' = unregistered (B2C) */
  gstin: string;
  how: CustHow;
  detail: string;
}
export interface CustRules {
  list: boolean;
  exact: boolean;
  invAmt: boolean;
  invAuto: boolean;
  invMin: number;
}
/**
 * the GSTIN each books customer without one gets, and why: the user's choice first ('?' = "ask me", so no automatic
 * mapping), then the customer list, the same name elsewhere in the books, and invoice no. + amount, as the settings allow
 */
export function resolveCustomers(
  sug: CustSuggestion[],
  rules: CustRules,
  listAliases: Map<string, string>,
  user: Record<string, string>,
): Map<string, CustMapping> {
  const out = new Map<string, CustMapping>();
  const keys = new Set([...sug.map(x => x.key), ...(rules.list ? listAliases.keys() : []), ...Object.keys(user)]);
  const byKey = new Map(sug.map(x => [x.key, x]));
  for (const k of keys) {
    const x = byKey.get(k);
    const u = user[k];
    if (u === '?') continue;
    if (u !== undefined) {
      out.set(k, { gstin: u, how: 'user', detail: u ? 'Picked and confirmed by you' : 'Marked by you as unregistered (B2C)' });
      continue;
    }
    if (rules.list && listAliases.has(k)) {
      out.set(k, { gstin: listAliases.get(k)!, how: 'list', detail: 'The customer list gives this name with one GSTIN' });
      continue;
    }
    if (!x) continue;
    if (rules.exact && x.exact) {
      out.set(k, { gstin: x.exact, how: 'exact', detail: 'Other entries in the books carry this name with this GSTIN' });
      continue;
    }
    if (rules.invAmt && rules.invAuto && x.inv && x.inv.hits >= Math.max(1, rules.invMin))
      out.set(k, {
        gstin: x.inv.gstin,
        how: 'invamt',
        detail: `${x.inv.hits} of ${x.inv.of} invoice(s) agree with GSTR-1 under this GSTIN by number and amount, and none with another GSTIN: ${x.inv.sample.join(', ')}${x.inv.hits > x.inv.sample.length ? ', …' : ''}`,
      });
  }
  return out;
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

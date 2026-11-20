// Portal auto-drafted GSTR-9 JSON (Prepare Offline → download) → RAW_Portal9 rows
import PATHS from './paths.json';
import { Amt, RawRow, n, r2 } from './util';

export function isPortal9(j: any) {
  const d = j?.gstr9 ?? j?.data ?? j;
  return !!(d && (d.table4 || d.table6 || d.table8) && d.gstin);
}
const amtOf = (o: any): Amt => [r2(n(o?.txval)), r2(n(o?.camt)), r2(n(o?.samt)), r2(n(o?.iamt)), r2(n(o?.csamt))];

export function processPortal9(j: any): { rows: RawRow[]; gstin: string; fy: string } {
  const d = j?.gstr9 ?? j?.data ?? j;
  const rows: RawRow[] = [];
  for (const [key, path] of Object.entries(PATHS as Record<string, string>)) {
    const m = path.match(/^(table\d+)\.(\w+)(?:\[(\w+)\])?$/);
    if (!m) continue;
    let node = d?.[m[1]]?.[m[2]];
    if (m[3]) node = Array.isArray(node) ? node.find((x: any) => x.itc_typ === m[3]) : undefined;
    if (node && typeof node === 'object') rows.push({ key, amt: amtOf(node), info: [path, ''] });
  }
  // Table 9: payable / cash / ITC by head
  const t9 = d?.table9 || {};
  for (const [k, head] of [
    ['iamt', '9A'],
    ['camt', '9B'],
    ['samt', '9C'],
    ['csamt', '9D'],
    ['intr', '9E'],
    ['fee', '9F'],
    ['pen', '9G'],
    ['other', '9H'],
  ]) {
    const o = t9[k];
    if (!o) continue;
    rows.push({ key: `${head}-PAY`, amt: [r2(n(o.txpyble)), 0, 0, 0, 0], info: [`table9.${k}.txpyble`, ''] });
    rows.push({ key: `${head}-CASH`, amt: [r2(n(o.txpaid_cash)), 0, 0, 0, 0], info: [`table9.${k}.txpaid_cash`, ''] });
    rows.push({
      key: `${head}-ITC`,
      amt: [0, r2(n(o.tax_paid_itc_camt)), r2(n(o.tax_paid_itc_samt)), r2(n(o.tax_paid_itc_iamt)), r2(n(o.tax_paid_itc_csamt))],
      info: [`table9.${k}.tax_paid_itc_*`, ''],
    });
  }
  const fp = String(d?.fp || '');
  const fy = fp.length === 6 ? `${+fp.slice(2) - 1}-${fp.slice(4)}` : '';
  return { rows, gstin: d?.gstin || '', fy };
}

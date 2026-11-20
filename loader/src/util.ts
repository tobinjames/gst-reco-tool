// Shared helpers
export type Amt = [number, number, number, number, number]; // Taxable, CGST, SGST, IGST, Cess

export const n = (v: unknown): number => {
  if (v === null || v === undefined || v === '' || v === '-') return 0;
  const x = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
  return isFinite(x) ? x : 0;
};
export const r2 = (x: number) => Math.round(x * 100) / 100;
export const zero = (): Amt => [0, 0, 0, 0, 0];
export const add = (a: Amt, b: Amt, sign = 1): Amt => a.map((v, i) => r2(v + sign * b[i])) as Amt;
export const neg = (a: Amt): Amt => a.map(v => r2(-v)) as Amt;
export const isZero = (a: Amt) => a.every(v => Math.abs(v) < 0.005);

/** "2025-26" for a date in that FY. Accepts dd-mm-yyyy or dd/mm/yyyy */
export function fyOfDate(d: string | undefined): string {
  if (!d) return '';
  const m = d.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (!m) return '';
  const mon = +m[2], yr = +m[3];
  const start = mon >= 4 ? yr : yr - 1;
  return `${start}-${String(start + 1).slice(2)}`;
}
/** "2025-26" for a return period MMYYYY */
export function fyOfPeriod(p: string | undefined): string {
  if (!p || p.length !== 6) return '';
  return fyOfDate(`01-${p.slice(0, 2)}-${p.slice(2)}`);
}
export function fyShift(fy: string, by: number): string {
  const s = +fy.slice(0, 4) + by;
  return `${s}-${String(s + 1).slice(2)}`;
}
/** item detail → Amt (GSTR-1 style keys) */
export function amtOf(o: any, txKey = 'txval'): Amt {
  return [r2(n(o?.[txKey])), r2(n(o?.camt)), r2(n(o?.samt)), r2(n(o?.iamt)), r2(n(o?.csamt))];
}
/** GSTR-2B style keys */
export function amt2b(o: any): Amt {
  return [r2(n(o?.txval)), r2(n(o?.cgst)), r2(n(o?.sgst)), r2(n(o?.igst)), r2(n(o?.cess))];
}
export const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];

export interface RawRow { key: string; amt: Amt; info: (string | number)[] }
export interface Warn { level: 'info' | 'warn' | 'error'; file: string; msg: string }

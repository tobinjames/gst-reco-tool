// Formulas for the summary sheets of the reconciliation workbooks, over the invoice rows of Recon_Purchases (from row 16),
// so the summaries follow any change made to the rows in Excel. Columns of Recon_Purchases (recon.ts reconPurchases):
// A status, C supplier GSTIN, E books no., F 2B no., G books date (dd-mm-yyyy), J-M books CGST/SGST/IGST/cess,
// O-R 2B CGST/SGST/IGST/cess, U 2B period (MM/YYYY), W 2B ITC available.
// Helper columns written beside every row (AD-AK, withHelpers) keep the summaries to one SUMIFS a cell (Tobin, Nov 2026:
// formulas a reader without the tool can follow): AD books tax, AE 2B tax, AF-AI claimable CGST/SGST/IGST/cess (the lower
// of books and 2B on a matched invoice whose 2B says ITC available), AJ claimable tax, AK books month (MM/YYYY).
import type { Reader, V, Workbook } from './xlsx';

const rg = (c: string) => `Recon_Purchases!$${c}$16:$${c}$20000`;
const NC = '"<>Not compared (RCM / ISD / imports)"';
const BK = ['J', 'K', 'L', 'M'],
  TB = ['O', 'P', 'Q', 'R'];
const f = (s: string): V => ({ f: s });
const sum4 = (cols: string[]) => `(${cols.map(rg).join('+')})`;
const CL = ['AF', 'AG', 'AH', 'AI'];
export const HELPER_HEADERS = [
  'Books tax',
  '2B tax',
  'Claimable CGST',
  'Claimable SGST',
  'Claimable IGST',
  'Claimable cess',
  'Claimable tax',
  'Books month',
];
/** the helper formulas AD-AK for the Recon_Purchases rows written from row 16 */
export function withHelpers(rows: any[][]): V[][] {
  return rows.map((row, i) => {
    const r = 16 + i,
      ok = `AND(LEFT($A${r},7)="Matched",$W${r}<>"N")`;
    const out: V[] = [...row];
    while (out.length < 29) out.push(null);
    out.length = 29;
    out.push(
      f(`J${r}+K${r}+L${r}+M${r}`),
      f(`O${r}+P${r}+Q${r}+R${r}`),
      ...BK.map((b, j) => f(`IF(${ok},MIN(${b}${r},${TB[j]}${r}),0)`)),
      f(`SUM(AF${r}:AI${r})`),
      f(`IF(LEN($G${r})=10,MID($G${r},4,2)&"/"&RIGHT($G${r},4),"")`),
    );
    return out;
  });
}

/** ITC position on Summary (C12:F17, one column per head) */
export function summaryCells(): Record<string, V> {
  const out: Record<string, V> = {};
  'CDEF'.split('').forEach((c, i) => {
    const b = rg(BK[i]),
      t = rg(TB[i]);
    out[`${c}12`] = f(`SUMIFS(${t},${rg('F')},"<>",${rg('W')},"<>N")`);
    out[`${c}13`] = f(`SUMIFS(${b},${rg('E')},"<>",${rg('A')},${NC})`);
    out[`${c}14`] = f(`SUM(${rg(CL[i])})`);
    out[`${c}15`] = f(`SUMIFS(${b},${rg('A')},"In books, not in 2B")`);
    out[`${c}16`] = f(`SUMIFS(${t},${rg('A')},"In 2B, not in books")`);
    out[`${c}17`] = f(
      `SUMIFS(${b},${rg('A')},"Matched - value difference")-SUMIFS(${t},${rg('A')},"Matched - value difference")`,
    );
  });
  return out;
}

/** Month_wise rows from row 6: keys are MMYYYY; column H carries the period (MM/YYYY) the formulas look up */
export function monthRows(keys: string[], label: (k: string) => string): V[][] {
  const rows: V[][] = keys.map((k, i) => {
    const r = 6 + i,
      per = `$H${r}`;
    return [
      label(k),
      f(`SUMIFS(${rg('AE')},${rg('U')},${per},${rg('F')},"<>",${rg('W')},"<>N")`),
      f(`SUMIFS(${rg('AJ')},${rg('U')},${per})`),
      f(`SUMIFS(${rg('AE')},${rg('U')},${per},${rg('A')},"In 2B, not in books")`),
      f(`SUMIFS(${rg('AD')},${rg('AK')},${per},${rg('E')},"<>",${rg('A')},${NC})`),
      f(`SUMIFS(${rg('AD')},${rg('AK')},${per},${rg('A')},"In books, not in 2B")`),
      null,
      `${k.slice(0, 2)}/${k.slice(2)}`,
    ];
  });
  if (rows.length > 1) {
    const n = 5 + rows.length;
    rows.push(['Total', ...'BCDEF'.split('').map(c => f(`SUM(${c}6:${c}${n})`)), null, null]);
  }
  return rows;
}

/** Recon_Suppliers rows from row 6 (order and remarks as computed; amounts and counts as formulas) */
export function supplierRows(sup: any[][], tol: number): V[][] {
  return sup.map((s, i) => {
    const r = 6 + i,
      k = s[0] === '(no GSTIN)' ? '"="' : `$A${r}`,
      c = rg('C');
    const cnt = (st: string) => f(`COUNTIFS(${c},${k},${rg('A')},"${st}")`);
    if (s[0] === '(no GSTIN)') {
      // rows without a GSTIN: an empty-cell criterion behaves differently across Excel and LibreOffice, so test the cells directly
      const g = `(${c}="")*(${rg('A')}<>"Not compared (RCM / ISD / imports)")`,
        sp = (x: string) => f(`SUMPRODUCT(${x})`);
      const st = (x: string) => sp(`(${c}="")*(${rg('A')}="${x}")`);
      return [
        s[0],
        s[1],
        sp(`${g}*(${rg('E')}<>"")`),
        sp(`${g}*(${rg('F')}<>"")`),
        sp(`${g}*${sum4(BK)}`),
        sp(`${g}*${sum4(TB)}`),
        f(`E${r}-F${r}`),
        st('In books, not in 2B'),
        st('In 2B, not in books'),
        st('Matched - value difference'),
        f(`IF(ABS(G${r})<=${tol},"Yes","No")`),
        s[11],
      ];
    }
    return [
      s[0],
      s[1],
      f(`COUNTIFS(${c},${k},${rg('E')},"<>",${rg('A')},${NC})`),
      f(`COUNTIFS(${c},${k},${rg('F')},"<>",${rg('A')},${NC})`),
      f(`SUMIFS(${rg('AD')},${c},${k},${rg('A')},${NC})`),
      f(`SUMIFS(${rg('AE')},${c},${k},${rg('A')},${NC})`),
      f(`E${r}-F${r}`),
      cnt('In books, not in 2B'),
      cnt('In 2B, not in books'),
      cnt('Matched - value difference'),
      f(`IF(ABS(G${r})<=${tol},"Yes","No")`),
      s[11],
    ];
  });
}

/** Recon_Sales_Monthly: the difference columns and the total row as formulas (books and GSTR-1 columns are the month totals) */
export function salesMonthRows(monthly: any[][]): V[][] {
  const n = monthly.length - 1; // last row is the total
  return monthly.map((m, i) => {
    const r = 6 + i;
    if (i === n) return ['Total', ...'BCDEFGHIJKLM'.split('').map(c => f(`SUM(${c}6:${c}${5 + n})`)), ''];
    return [
      m[0],
      m[1],
      m[2],
      f(`B${r}-C${r}`),
      m[4],
      m[5],
      f(`E${r}-F${r}`),
      m[7],
      m[8],
      f(`H${r}-I${r}`),
      m[10],
      m[11],
      f(`K${r}-L${r}`),
      m[13],
    ];
  });
}

/** Workbooks built before Nov 2026 had the status summary of Recon_Purchases on rows 4-15, so its Total row sat on the
 *  table's header row and columns A-F had no headings. When such a workbook is refreshed, move the summary to rows 3-14,
 *  put the headings back, and point Summary / Checks at the moved cells. */
export async function fixReconLayout(wb: Workbook, rd: Reader | null) {
  if (!rd || !wb.has('Recon_Purchases') || (await rd.get('Recon_Purchases', 'A15')).v !== 'Total') return;
  const S = 'Recon_Purchases',
    cols = 'ABCDEF'.split('');
  const refs = ['A3', 'G15', ...[4, 5, 6, 14, 15].flatMap(r => cols.map(c => c + r))];
  const st = await wb.styles(S, refs);
  const cells: Record<string, V> = {},
    sty: Record<string, string | undefined> = {};
  const put = (ref: string, v: V, from: string) => {
    cells[ref] = v;
    sty[ref] = st[from];
  };
  put('C2', (await rd.get(S, 'A3')).v, 'A3');
  put('A3', null, 'A3');
  for (const c of cols) put(`${c}3`, (await rd.get(S, `${c}4`)).v, `${c}4`);
  for (let r = 4; r <= 13; r++) {
    const from = r === 4 ? 5 : 6; // first status bold, the rest plain
    put(`A${r}`, (await rd.get(S, `A${r + 1}`)).v, `A${from}`);
    put(`B${r}`, f(`COUNTIF($A$16:$A$20000,A${r})`), `B${from}`);
    ['I', 'N', 'S', 'T'].forEach((sc, j) =>
      put(`${cols[2 + j]}${r}`, f(`SUMIF($A$16:$A$20000,$A${r},$${sc}$16:$${sc}$20000)`), `${cols[2 + j]}${from}`),
    );
  }
  put('A14', 'Total', 'A15');
  for (const c of cols.slice(1)) put(`${c}14`, f(`SUM(${c}4:${c}13)`), `${c}15`);
  ['Status', 'Match basis', 'Supplier GSTIN', 'Supplier name', 'Invoice no. (books)', 'Invoice no. (2B)'].forEach((h, j) =>
    put(`${cols[j]}15`, h, 'G15'),
  );
  await wb.setCells(S, cells, sty);
  if (wb.has('Summary') && rd.has('Summary') && (await rd.get('Summary', 'B21')).ftext === 'Recon_Purchases!A5') {
    const sm: Record<string, V> = {};
    for (let i = 0; i < 10; i++) {
      sm[`B${21 + i}`] = f(`Recon_Purchases!A${4 + i}`);
      sm[`C${21 + i}`] = f(`Recon_Purchases!B${4 + i}`);
      sm[`D${21 + i}`] = f(`Recon_Purchases!F${4 + i}`);
    }
    await wb.setCells('Summary', sm);
  }
  if (wb.has('Checks') && rd.has('Checks')) {
    const ck: Record<string, V> = {};
    for (const ref of ['C32', 'C33']) {
      const t = (await rd.get('Checks', ref)).ftext;
      if (t === 'Recon_Purchases!F12') ck[ref] = f('Recon_Purchases!F11');
      if (t === '-Recon_Purchases!F13') ck[ref] = f('-Recon_Purchases!F12');
    }
    if (Object.keys(ck).length) await wb.setCells('Checks', ck);
  }
}

// Finished workbook → GSTR-9 upload JSON (same structure as the GSTN offline tool v2.2 "Generate JSON File to upload")
import ROWMAP from './rowmap.json';
import { Reader } from './xlsx';
import { Integrity, checkIntegrity } from './integrity';

type Msg = { level: 'error' | 'warn' | 'info'; msg: string };
const TBF = ['Y', 'Z', 'AA', 'AB', 'AC']; // To be Filed: Taxable, CGST, SGST, IGST, Cess
const PORTAL = ['E', 'F', 'G', 'H', 'I'];
const FIELD_COL: Record<string, number> = { txval: 0, camt: 1, samt: 2, iamt: 3, csamt: 4 };
const R = (x: number) => Math.round(x * 100) / 100;

// per JSON object: [workbook key, fields] - field lists copied from the offline tool's export code
const SIMPLE: Record<string, [string, string][]> = {
  table4: [
    ['b2c', '4A'],
    ['b2b', '4B'],
    ['exp', '4C'],
    ['sez', '4D'],
    ['deemed', '4E'],
    ['at', '4F'],
    ['rchrg', '4G'],
    ['ecom', '4G1'],
    ['cr_nt', '4I'],
    ['dr_nt', '4J'],
    ['amd_pos', '4K'],
    ['amd_neg', '4L'],
  ],
  table5: [
    ['zero_rtd', '5A'],
    ['sez', '5B'],
    ['rchrg', '5C'],
    ['ecom_14', '5C1'],
    ['exmt', '5D'],
    ['nil', '5E'],
    ['non_gst', '5F'],
    ['cr_nt', '5H'],
    ['dr_nt', '5I'],
    ['amd_pos', '5J'],
    ['amd_neg', '5K'],
  ],
  table6: [
    ['itc_otr', '6A1'],
    ['ios', '6F'],
    ['isd', '6G'],
    ['itc_clmd', '6H'],
    ['tran1', '6K'],
    ['tran2', '6L'],
    ['other', '6M'],
  ],
  table7: [
    ['rule37', '7A'],
    ['rule37A', '7A1'],
    ['rule38', '7A2'],
    ['rule39', '7B'],
    ['rule42', '7C'],
    ['rule43', '7D'],
    ['sec17', '7E'],
    ['revsl_tran1', '7F'],
    ['revsl_tran2', '7G'],
  ],
  table8: [
    ['itc_inwd_supp', '8C'],
    ['itc_nt_availd', '8E'],
    ['itc_nt_eleg', '8F'],
    ['iog_taxpaid', '8G'],
    ['iog_nxt_yr', '8H1'],
  ],
  table10: [
    ['dbn_amd', '10'],
    ['cdn_amd', '11'],
    ['itc_rvsl', '12'],
    ['itc_availd', '13'],
  ],
};
const FIELDS: Record<string, string[]> = {
  'table4.exp': ['txval', 'iamt', 'csamt'],
  'table4.sez': ['txval', 'iamt', 'csamt'],
  'table6.ios': ['iamt', 'csamt'],
  'table6.tran1': ['camt', 'samt'],
  'table6.tran2': ['camt', 'samt'],
  'table7.revsl_tran1': ['camt', 'samt'],
  'table7.revsl_tran2': ['camt', 'samt'],
};
const defaultFields = (table: string) =>
  table === 'table4' || table === 'table10a'
    ? ['txval', 'iamt', 'camt', 'samt', 'csamt']
    : table === 'table5'
      ? ['txval']
      : ['iamt', 'camt', 'samt', 'csamt'];
const ARRAYS: [string, string, [string, string][], string[]][] = [
  [
    'table6',
    'supp_non_rchrg',
    [
      ['ip', '6B-ip'],
      ['cg', '6B-cg'],
      ['is', '6B-is'],
    ],
    ['iamt', 'camt', 'samt', 'csamt'],
  ],
  [
    'table6',
    'supp_rchrg_unreg',
    [
      ['ip', '6C-ip'],
      ['cg', '6C-cg'],
      ['is', '6C-is'],
    ],
    ['iamt', 'camt', 'samt', 'csamt'],
  ],
  [
    'table6',
    'supp_rchrg_reg',
    [
      ['ip', '6D-ip'],
      ['cg', '6D-cg'],
      ['is', '6D-is'],
    ],
    ['iamt', 'camt', 'samt', 'csamt'],
  ],
  [
    'table6',
    'iog',
    [
      ['ip', '6E-ip'],
      ['cg', '6E-cg'],
    ],
    ['iamt', 'csamt'],
  ],
];

export interface Gen9Result {
  integrity?: Integrity;
  json: any;
  text: string;
  msgs: Msg[];
  gstin: string;
  fy: string;
  lines: { table: string; field: string; values: Record<string, number> }[];
}

export async function generate9(bytes: Uint8Array, template?: Uint8Array): Promise<Gen9Result> {
  const rd = await Reader.open(bytes);
  if (!rd.has('4-5 Outward') || !rd.has('Basic')) throw new Error('This is not a GST Annual Return Workbook');
  const msgs: Msg[] = [];
  let uncalculated = 0;
  const num = async (sheet: string, ref: string) => {
    const c = await rd.get(sheet, ref);
    if (c.f && (c.v === null || c.v === undefined)) uncalculated++;
    if (c.v && typeof c.v === 'object' && 'error' in c.v) {
      msgs.push({ level: 'error', msg: `${sheet}!${ref} shows an error (${c.v.error})` });
      return 0;
    }
    return typeof c.v === 'number' ? R(c.v) : c.v === null || c.v === '' ? 0 : R(+c.v || 0);
  };
  const str = async (sheet: string, ref: string) => {
    const c = await rd.get(sheet, ref);
    return c.v === null || c.v === undefined ? '' : String(c.v).trim();
  };

  const gstin = (await str('Basic', 'C6')).toUpperCase();
  const fy = await str('Basic', 'C7');
  if (!/^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/.test(gstin))
    msgs.push({ level: 'error', msg: `GSTIN "${gstin}" on the Basic sheet is not valid` });
  if (!/^\d{4}-\d{2}$/.test(fy))
    msgs.push({ level: 'error', msg: `Financial year "${fy}" on the Basic sheet is not in YYYY-YY format` });
  const fp = '03' + fy.slice(0, 2) + fy.slice(-2);

  const out: any = { gstin, fp };
  const lines: Gen9Result['lines'] = [];
  const rowVals = async (key: string) => {
    const [sheet, row] = (ROWMAP as any)[key] as [string, number];
    const final = await Promise.all(TBF.map(c => num(sheet, c + row)));
    const portal = await Promise.all(PORTAL.map(c => num(sheet, c + row)));
    return { final, portal, sheet, row };
  };
  const pick = (vals: number[], fields: string[]) => Object.fromEntries(fields.map(f => [f, vals[FIELD_COL[f]]]));
  const anyNz = (a: number[], fields: string[]) => fields.some(f => Math.abs(a[FIELD_COL[f]]) >= 0.01);

  for (const [table, list] of Object.entries(SIMPLE)) {
    for (const [field, key] of list) {
      const fields =
        FIELDS[`${table}.${field}`] ||
        (table === 'table10'
          ? key === '10' || key === '11'
            ? ['txval', 'iamt', 'camt', 'samt', 'csamt']
            : ['iamt', 'camt', 'samt', 'csamt']
          : defaultFields(table));
      const { final, portal } = await rowVals(key);
      // include a row when it has a value, or when it clears a value the portal pre-filled
      if (!anyNz(final, fields) && !anyNz(portal, fields)) continue;
      (out[table] ||= {})[field] = pick(final, fields);
      lines.push({ table, field, values: pick(final, fields) });
    }
  }
  for (const [table, field, items, fields] of ARRAYS) {
    const arr: any[] = [];
    for (const [typ, key] of items) {
      const { final, portal } = await rowVals(key);
      if (!anyNz(final, fields) && !anyNz(portal, fields)) continue;
      arr.push({ itc_typ: typ, ...pick(final, fields) });
      lines.push({ table, field: `${field} [${typ}]`, values: pick(final, fields) });
    }
    if (arr.length) (out[table] ||= {})[field] = arr;
  }
  // 7H other reversals: array with description (Remarks column)
  {
    const { final, portal, sheet, row } = await rowVals('7H');
    const f = ['iamt', 'camt', 'samt', 'csamt'];
    if (anyNz(final, f) || anyNz(portal, f)) {
      const desc = (await str(sheet, 'AJ' + row)) || 'Other reversals';
      (out.table7 ||= {}).other = [{ desc, ...pick(final, f) }];
      lines.push({ table: 'table7', field: `other (${desc})`, values: pick(final, f) });
    }
  }
  // Table 9: tax payable only (paid columns are auto-populated by the portal)
  const T9 = '9 & 14 Tax Paid';
  const t9: any = {};
  for (const [field, row] of [
    ['iamt', 6],
    ['camt', 7],
    ['samt', 8],
    ['csamt', 9],
    ['intr', 10],
    ['fee', 11],
    ['pen', 12],
    ['other', 13],
  ] as const) {
    const v = await num(T9, 'F' + row),
      p = await num(T9, 'C' + row);
    if (Math.abs(v) >= 0.01 || Math.abs(p) >= 0.01) {
      t9[field] = { txpyble: v };
      lines.push({ table: 'table9', field, values: { txpyble: v } });
    }
  }
  if (Object.keys(t9).length) out.table9 = t9;
  // Table 14: tax paid (payable is computed by the portal from 10 - 11); interest payable + paid
  const t14: any = {};
  for (const [field, row] of [
    ['iamt', 20],
    ['camt', 21],
    ['samt', 22],
    ['csamt', 23],
  ] as const) {
    const paid = await num(T9, 'F' + row);
    if (Math.abs(paid) >= 0.01) t14[field] = { txpaid: paid };
  }
  {
    const pay = await num(T9, 'E24'),
      paid = await num(T9, 'F24');
    if (Math.abs(pay) >= 0.01 || Math.abs(paid) >= 0.01) t14.intr = { txpyble: pay, txpaid: paid };
  }
  if (Object.keys(t14).length) {
    out.table14 = t14;
    lines.push({ table: 'table14', field: Object.keys(t14).join(', '), values: {} });
  }
  // Table 15 & 16 (manual sheet)
  const O = '15-16 Other';
  const t15: any = {};
  for (const [i, field] of ['rfd_clmd', 'rfd_sanc', 'rfd_rejt', 'rfd_pend', 'tax_dmnd', 'tax_paid', 'dmnd_pend'].entries()) {
    const r = 7 + i;
    const [c, s, ig, cs, intr, pen, fee] = await Promise.all(['C', 'D', 'E', 'F', 'G', 'H', 'I'].map(col => num(O, col + r)));
    const o: any = { iamt: ig, camt: c, samt: s, csamt: cs };
    if (i >= 4) Object.assign(o, { intr, fee, pen });
    if (Object.values(o).some((v: any) => Math.abs(v) >= 0.01)) {
      t15[field] = o;
      lines.push({ table: 'table15', field, values: o });
    }
  }
  if (Object.keys(t15).length) out.table15 = t15;
  const t16: any = {};
  for (const [i, field] of ['comp_supp', 'deemed_supp', 'not_returned'].entries()) {
    const r = 18 + i;
    const [tx, c, s, ig, cs] = await Promise.all(['C', 'D', 'E', 'F', 'G'].map(col => num(O, col + r)));
    const o: any = field === 'comp_supp' ? { txval: tx } : { txval: tx, iamt: ig, camt: c, samt: s, csamt: cs };
    if (Object.values(o).some((v: any) => Math.abs(v) >= 0.01)) {
      t16[field] = o;
      lines.push({ table: 'table16', field, values: o });
    }
  }
  if (Object.keys(t16).length) out.table16 = t16;
  // Tables 17 / 18: HSN rows with Action = Add
  for (const [sheet, table] of [
    ['17 HSN Outward', 'table17'],
    ['18 HSN Inward', 'table18'],
  ] as const) {
    const items: any[] = [];
    let bad = 0;
    for (let r = 7; r <= 1506; r++) {
      const hsn = await str(sheet, 'A' + r);
      if (!hsn) continue;
      const action = (await str(sheet, 'L' + r)) || 'Add';
      if (action !== 'Add') continue;
      if ((await str(sheet, 'N' + r)) === 'No') bad++;
      const it: any = { hsn_sc: hsn };
      if (!hsn.startsWith('99')) {
        it.uqc = (await str(sheet, 'C' + r)).toUpperCase() || 'NA';
        it.qty = await num(sheet, 'D' + r);
      }
      it.txval = await num(sheet, 'E' + r);
      const conc = await str(sheet, 'F' + r);
      it.isconcesstional = conc === 'Yes' ? 'Y' : 'N';
      it.rt = await num(sheet, 'G' + r);
      it.iamt = await num(sheet, 'H' + r);
      it.camt = await num(sheet, 'I' + r);
      it.samt = await num(sheet, 'J' + r);
      it.csamt = await num(sheet, 'K' + r);
      items.push(it);
    }
    if (bad && table === 'table17')
      msgs.push({
        level: 'error',
        msg: `Table 17 has ${bad} HSN code(s) not in the GSTN master - the portal will reject them. Fix them on the "17 HSN Outward" sheet.`,
      });
    if (items.length) {
      out[table] = { items };
      lines.push({ table, field: `${items.length} HSN line(s)`, values: {} });
    }
  }

  if (uncalculated)
    msgs.unshift({
      level: 'error',
      msg: `${uncalculated} formula cell(s) have no calculated value. Open the workbook in Excel (or LibreOffice), let it calculate, save, and drop the saved file here again.`,
    });
  // carry over review checks
  if (rd.has('Checks')) {
    for (let r = 6; r <= 60; r++) {
      const st = await str('Checks', 'D' + r);
      if (!st) continue;
      if (st === 'Check')
        msgs.push({
          level: 'warn',
          msg: `Checks sheet: ${await str('Checks', 'B' + r)} = ${(await num('Checks', 'C' + r)).toLocaleString('en-IN')}`,
        });
    }
  }
  for (const l of lines)
    for (const [k, v] of Object.entries(l.values))
      if (v < 0)
        msgs.push({
          level: 'warn',
          msg: `${l.table}.${l.field}.${k} is negative (${v}); the portal accepts negatives only in some fields`,
        });
  const text = JSON.stringify(out);
  const integrity = template ? await checkIntegrity(rd, await Reader.open(template)) : undefined;
  return { json: out, text, msgs, gstin, fy, lines, integrity };
}

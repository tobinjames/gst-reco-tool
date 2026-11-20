// Orchestrates: classify input files → parse → write workbook
import JSZip from 'jszip';
import { G1File, isGstr1, processGstr1 } from './gstr1';
import { is2b, process2b } from './gstr2b';
import { ThreeB } from './gstr3b';
import { readTable } from './register';
import { parsePortal, portalKindOf } from './flow/portal';
import { looksLikeBooks, parseBooks } from './flow/books';
import { BookFile, PORTAL_LABEL, PortalLedger } from './flow/types';
import { isPortal9, processPortal9 } from './portal9';
import { RawRow, Warn, fyOfDate, fyOfPeriod, fyShift, r2 } from './util';
import { Reader, Workbook, colLetter as colLetterOf } from './xlsx';
import { RAW_SHEETS, fingerprint, seal, sealLines, readRows } from './integrity';
import { fixReconLayout, monthRows, salesMonthRows, summaryCells, supplierRows } from './reconsheets';
import { DEFAULT_MATCH, KeepApart, Manual, MatchOptions, describeMatch, reconPurchases, reconSales } from './recon';

/** what the user decided in the Loader's Review step: manual matches (Match codes) and documents kept apart (code X) */
export interface Decisions { manual: Manual; keepApart: KeepApart }
export interface ReconExtra { match?: MatchOptions; decisions?: { purchases?: Decisions; sales?: Decisions } }
export const LOADER_VERSION = '0.4';

export type Kind = 'GSTR-1' | 'GSTR-2B' | 'GSTR-3B' | 'GSTR-9 (portal)' | 'Workbook' | 'Register' | 'Portal ledger' | 'Books ledger' | 'Unknown';
export interface InFile { name: string; hash?: string; kind: Kind; json?: any; threeB?: ThreeB; bytes?: Uint8Array; gstin?: string; period?: string; fy?: string; note?: string; portal?: PortalLedger; books?: BookFile }

/** expand zips, parse JSON, detect kind. pdfs are parsed by the caller (pdf.js differs in browser / node) */
export async function sha1(bytes: Uint8Array) {
  const h = await crypto.subtle.digest('SHA-1', bytes as any);
  return Array.from(new Uint8Array(h)).map(b => b.toString(16).padStart(2, '0')).join('');
}

export async function classify(name: string, bytes: Uint8Array, parsePdf: (b: Uint8Array, name: string) => Promise<ThreeB>): Promise<InFile[]> {
  const out = await classifyInner(name, bytes, parsePdf);
  if (!name.toLowerCase().endsWith('.zip')) { const h = await sha1(bytes); out.forEach(f => (f.hash = h)); }
  return out;
}

async function classifyInner(name: string, bytes: Uint8Array, parsePdf: (b: Uint8Array, name: string) => Promise<ThreeB>): Promise<InFile[]> {
  const lower = name.toLowerCase();
  if (lower.endsWith('.zip')) {
    const z = await JSZip.loadAsync(bytes); const out: InFile[] = [];
    for (const f of Object.values(z.files)) if (!f.dir) out.push(...await classify(f.name.split('/').pop()!, await f.async('uint8array'), parsePdf));
    return out;
  }
  if (lower.endsWith('.pdf')) {
    try {
      const t = await parsePdf(bytes, name);
      return [{ name, kind: 'GSTR-3B', threeB: t, gstin: t.gstin, period: t.period, fy: fyOfPeriod(t.period), note: t.problems.length ? `${t.problems.length} issue(s)` : '' }];
    } catch (e: any) { return [{ name, kind: 'Unknown', note: 'PDF could not be read: ' + e.message }]; }
  }
  if (lower.endsWith('.xlsx') || lower.endsWith('.csv')) {
    if (lower.endsWith('.xlsx')) { try { const rd = await Reader.open(bytes); if (rd.has('RAW_GSTR1') && rd.has('Basic')) return [{ name, kind: 'Workbook', bytes, note: 'Annual return workbook' }];
      if (rd.has('_ITC') && rd.has('Recon_Purchases')) return [{ name, kind: 'Workbook', bytes, note: 'ITC reconciliation workbook' }]; } catch { return [{ name, kind: 'Unknown', note: 'Could not open the Excel file' }]; } }
    // GST portal ledgers and books ledgers / GL dumps (for the GST ledger flow tab); anything else is a register to map
    try {
      const sheets = await readTable(name, bytes);
      for (const sh of sheets) {
        if (portalKindOf(sh.rows)) {
          const p = parsePortal(sh.rows);
          return [{ name, kind: 'Portal ledger', bytes, portal: p, gstin: p.gstin, fy: fyOfDate(p.from), note: `${PORTAL_LABEL[p.kind]} · ${p.from || '?'} to ${p.to || '?'}${p.problems.length ? ` · ${p.problems.length} issue(s)` : ''}` }];
        }
      }
      for (const sh of sheets) if (looksLikeBooks(sh.rows)) {
        const b = parseBooks(name, sh.rows);
        const ds = b.lines.map(l => l.date.slice(6) + l.date.slice(3, 5) + l.date.slice(0, 2)).sort();
        const span = ds.length ? `${ds[0].slice(6)}-${ds[0].slice(4, 6)}-${ds[0].slice(0, 4)} to ${ds[ds.length - 1].slice(6)}-${ds[ds.length - 1].slice(4, 6)}-${ds[ds.length - 1].slice(0, 4)}` : '';
        return [{ name, kind: 'Books ledger', bytes, books: b, fy: ds.length ? fyOfDate(`${ds[0].slice(6)}-${ds[0].slice(4, 6)}-${ds[0].slice(0, 4)}`) : undefined,
          note: `${b.format === 'tally-dump' ? 'Tally voucher dump' : b.format === 'zoho-gl' ? 'Zoho detailed GL' : `Ledger: ${b.ledger}`} · ${b.lines.length} lines${span ? ' · ' + span : ''}` }];
      }
    } catch { /* fall through: treat as a register */ }
    return [{ name, kind: 'Register', bytes, note: 'map columns below' }];
  }
  if (lower.endsWith('.xls')) return [{ name, kind: 'Unknown', note: 'Old .xls format: open in Excel and save as .xlsx' }];
  if (lower.endsWith('.json')) {
    let j: any;
    try { j = JSON.parse(new TextDecoder().decode(bytes).replace(/^﻿/, '')); } catch { return [{ name, kind: 'Unknown', note: 'Not valid JSON' }]; }
    if (is2b(j)) { const d = j.data ?? j; return [{ name, kind: 'GSTR-2B', json: j, gstin: d.gstin, period: d.rtnprd, fy: fyOfPeriod(d.rtnprd) }]; }
    if (isPortal9(j)) { const p = processPortal9(j); return [{ name, kind: 'GSTR-9 (portal)', json: j, gstin: p.gstin, fy: p.fy }]; }
    if (isGstr1(j)) return [{ name, kind: 'GSTR-1', json: j, gstin: j.gstin, period: String(j.fp), fy: fyOfPeriod(String(j.fp)) }];
    return [{ name, kind: 'Unknown', note: 'JSON not recognised as GSTR-1 / GSTR-2B / GSTR-9' }];
  }
  return [{ name, kind: 'Unknown', note: 'Unsupported file type' }];
}

const p9Present = (fs: InFile[]) => fs.some(f => f.kind === 'GSTR-9 (portal)');
const row = (r: RawRow) => [r.key || null, ...r.amt, ...r.info];

export interface BuildResult { bytes: Uint8Array; warns: Warn[]; summary: Record<string, number>; counts: Record<string, number>; recon?: any }

export async function buildWorkbook(template: Uint8Array, files: InFile[], fy: string, extra: { legalName?: string; gstin?: string; sales?: any[][]; purchases?: any[][] } & ReconExtra): Promise<BuildResult> {
  const warns: Warn[] = [];
  const py = fyShift(fy, -1), ny = fyShift(fy, 1);
  const base = files.find(f => f.kind === 'Workbook' && f.note !== 'ITC reconciliation workbook');
  const wb = await Workbook.open(base?.bytes ?? template);
  const written: Record<string, any[][]> = {};
  const origWrite = wb.writeRows.bind(wb);
  wb.writeRows = async (name: string, start: number, data: any[][]) => { if (RAW_SHEETS.includes(name)) written[name] = data; return origWrite(name, start, data); };
  // previous fingerprints (refreshing an existing workbook keeps those of sheets not reloaded)
  const prevFp: Record<string, [string, number]> = {};
  if (base) {
    const rd = await Reader.open(base.bytes!);
    if (rd.has('_Integrity')) for (const [ref, c] of await rd.sheet('_Integrity')) {
      if (!ref.startsWith('A') || +ref.slice(1) < 4) continue;
      const r = ref.slice(1);
      prevFp[String(c.v)] = [String((await rd.get('_Integrity', 'B' + r)).v ?? ''), Number((await rd.get('_Integrity', 'C' + r)).v ?? 0)];
    }
  }
  if (!wb.has('RAW_GSTR1')) throw new Error(`${base?.name} is not a GST Annual Return Workbook`);

  // GSTIN consistency
  const gstins = new Set(files.filter(f => f.gstin).map(f => f.gstin));
  if (gstins.size > 1) warns.push({ level: 'error', file: '', msg: `Files belong to different GSTINs: ${[...gstins].join(', ')}` });
  const gstin = extra.gstin || [...gstins][0] || '';

  // ---- GSTR-1
  const g1 = files.filter(f => f.kind === 'GSTR-1').map(f => ({ name: f.name, json: f.json }) as G1File);
  const { rows: g1rows, hsn } = processGstr1(g1, fy, ny, warns);
  const keep = (n: number) => !!base && n === 0;   // refreshing an existing workbook: leave sheets alone when no new files of that type
  if (!keep(g1.length)) await wb.writeRows('RAW_GSTR1', 6, g1rows.map(row));
  // HSN raw + aggregated Table 17
  if (!keep(g1.length)) await wb.writeRows('RAW_GSTR1_HSN', 6, hsn.map(h => [String(h.hsn_sc ?? ''), h.desc ?? '', h.uqc ?? '', r2(+h.qty || 0), r2(+h.txval || 0), +h.rt || 0,
    r2(+h.iamt || 0), r2(+h.camt || 0), r2(+h.samt || 0), r2(+h.csamt || 0), h.period, h.tag]));
  const agg = new Map<string, any>();
  for (const h of hsn) {
    const k = `${h.hsn_sc}|${(h.uqc || '').toUpperCase()}|${+h.rt || 0}`;
    const a = agg.get(k) || { hsn: String(h.hsn_sc), desc: h.desc || '', uqc: (h.uqc || '').toUpperCase(), qty: 0, tx: 0, rt: +h.rt || 0, i: 0, c: 0, s: 0, cs: 0 };
    a.qty += +h.qty || 0; a.tx += +h.txval || 0; a.i += +h.iamt || 0; a.c += +h.camt || 0; a.s += +h.samt || 0; a.cs += +h.csamt || 0;
    agg.set(k, a);
  }
  const hsnCells: Record<string, any> = {};
  [...agg.values()].sort((a, b) => a.hsn.localeCompare(b.hsn)).forEach((a, i) => {
    const r = 7 + i; if (r > 1506) return;
    Object.assign(hsnCells, { [`A${r}`]: a.hsn, [`B${r}`]: a.desc, [`C${r}`]: a.uqc, [`D${r}`]: r2(a.qty), [`E${r}`]: r2(a.tx), [`F${r}`]: 'No', [`G${r}`]: a.rt,
      [`H${r}`]: r2(a.i), [`I${r}`]: r2(a.c), [`J${r}`]: r2(a.s), [`K${r}`]: r2(a.cs), [`L${r}`]: 'Add', [`M${r}`]: 'GSTR-1' });
  });
  if (agg.size > 1500) warns.push({ level: 'warn', file: '', msg: `Table 17 has ${agg.size} HSN lines; only the first 1500 were written` });
  if (Object.keys(hsnCells).length) await wb.setCells('17 HSN Outward', hsnCells);

  // ---- GSTR-3B
  const tb = files.filter(f => f.kind === 'GSTR-3B' && f.threeB);
  const seen3b = new Map<string, string>();
  const rows3b: any[] = [], pay: any[] = [];
  for (const f of tb) {
    const t = f.threeB!;
    for (const p of t.problems) warns.push({ level: 'warn', file: f.name, msg: p });
    if (fyOfPeriod(t.period) !== fy) { warns.push({ level: 'info', file: f.name, msg: `3B ${t.period} is not in FY ${fy}; skipped` }); continue; }
    if (seen3b.has(t.period)) { warns.push({ level: 'warn', file: f.name, msg: `3B for ${t.period} already loaded from ${seen3b.get(t.period)}; skipped` }); continue; }
    seen3b.set(t.period, f.name);
    rows3b.push(...t.rows.map(row));
    pay.push(...t.pay.map(p => [p.key, p.amount, t.period, f.name]));
  }
  const months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].map(m => String(m).padStart(2, '0') + (m >= 4 ? fy.slice(0, 4) : +fy.slice(0, 4) + 1));
  const missing3b = months.filter(m => !seen3b.has(m));
  if (missing3b.length) warns.push({ level: 'error', file: '', msg: `GSTR-3B ${fy} missing for: ${missing3b.map(m => m.slice(0, 2) + '/' + m.slice(2)).join(', ')} (QRMP filers: only quarter-end months are expected)` });
  if (!keep(tb.length)) { await wb.writeRows('RAW_3B', 6, rows3b); await wb.writeRows('RAW_3B_PAY', 6, pay); }

  // ---- GSTR-2B
  const b2 = files.filter(f => f.kind === 'GSTR-2B').map(f => ({ name: f.name, json: f.json }));
  const { rows: r2b, periods } = process2b(b2, fy, py, ny, warns);
  if (!keep(b2.length)) await wb.writeRows('RAW_2B', 6, r2b.map(row));
  const fmtM = (l: string[]) => l.map(m => m.slice(0, 2) + '/' + m.slice(2)).join(', ');
  const nyMonths = [4, 5, 6, 7, 8, 9, 10, 11].map(m => String(m).padStart(2, '0') + ny.slice(0, 4));
  const missing2b = months.filter(m => !periods.includes(m));
  if (missing2b.length) warns.push({ level: 'error', file: '', msg: `GSTR-2B ${fy} missing for: ${fmtM(missing2b)}` });
  const missing2bNy = nyMonths.filter(m => !periods.includes(m));
  if (missing2bNy.length) warns.push({ level: 'warn', file: '', msg: `GSTR-2B ${ny} (Apr–Nov, for 8A/8C/13) missing for: ${fmtM(missing2bNy)}` });
  // GSTR-1: quarterly returns cover their whole quarter
  const g1cov = new Set<string>();
  for (const f of files.filter(x => x.kind === 'GSTR-1')) {
    const p = String(f.json.fp);
    if (f.json.filing_typ === 'Q') { const m = +p.slice(0, 2); [m - 2, m - 1, m].forEach(x => g1cov.add(String(x).padStart(2, '0') + p.slice(2))); }
    else g1cov.add(p);
  }
  const missing1 = months.filter(m => !g1cov.has(m));
  if (missing1.length) warns.push({ level: 'error', file: '', msg: `GSTR-1 ${fy} missing for: ${fmtM(missing1)}` });
  const missing1Ny = nyMonths.filter(m => !g1cov.has(m));
  if (missing1Ny.length) warns.push({ level: 'warn', file: '', msg: `GSTR-1 ${ny} (Apr–Nov, for Tables 10/11) missing for: ${fmtM(missing1Ny)}` });
  if (!p9Present(files)) warns.push({ level: 'warn', file: '', msg: 'Portal GSTR-9 JSON not loaded - Pre-filled (Portal) columns will be blank' });

  // ---- Portal GSTR-9
  const p9 = files.filter(f => f.kind === 'GSTR-9 (portal)');
  if (p9.length > 1) warns.push({ level: 'warn', file: '', msg: 'More than one portal GSTR-9 JSON; using the first' });
  if (p9.length) {
    const p = processPortal9(p9[0].json);
    if (p.fy && p.fy !== fy) warns.push({ level: 'error', file: p9[0].name, msg: `Portal GSTR-9 is for FY ${p.fy}, workbook FY is ${fy}` });
    await wb.writeRows('RAW_Portal9', 6, p.rows.map(row));
  }

  // ---- Basic
  const legal = extra.legalName || tb.find(f => f.threeB?.legalName)?.threeB?.legalName || '';
  const basic: Record<string, any> = { C6: gstin, C7: fy };
  if (legal) basic.C4 = legal;
  await wb.setCells('Basic', basic);

  // ---- registers (mapped in the Loader) → Books sheets
  if (extra.sales) await wb.fillTable('Books_Sales', 6, extra.sales, 16);
  if (extra.purchases) await wb.fillTable('Books_Purchase', 6, extra.purchases, 18);

  // ---- reconciliation (needs books + returns; books come from the mapped registers or from the dropped workbook)
  const baseRd = base ? await Reader.open(base.bytes!) : null;
  await fixReconLayout(wb, baseRd);
  const rawOf = async (sh: string) => written[sh] ?? (baseRd && baseRd.has(sh) ? await readRows(baseRd, sh, 6) : []);
  const booksOf = async (sh: string, given?: any[][]) => given ?? (baseRd && baseRd.has(sh) ? await readRows(baseRd, sh, 6) : []);
  const { remarks, manual, keepApart } = await readReconState(baseRd);
  const mo = extra.match ?? DEFAULT_MATCH;
  const recon: any = {};
  const bp = (await booksOf('Books_Purchase', extra.purchases)).filter(r => r && [12, 13, 14, 15, 16].some(i => typeof r[i] === 'number' && r[i] !== 0));
  const bs = (await booksOf('Books_Sales', extra.sales)).filter(r => r && [11, 12, 13, 14, 15].some(i => typeof r[i] === 'number' && r[i] !== 0));
  const raw2b = await rawOf('RAW_2B'), raw1 = await rawOf('RAW_GSTR1');
  if (wb.has('Recon_Purchases')) {
    if (bp.length && raw2b.length) {
      const notes: string[] = [];
      const dp = extra.decisions?.purchases;
      const pr = reconPurchases(bp, raw2b, fy, ny, remarks, dp?.manual ?? manual.Recon_Purchases, notes, mo, dp?.keepApart ?? keepApart.Recon_Purchases);
      notes.forEach(m => warns.push({ level: 'warn', file: 'Recon_Purchases', msg: m }));
      await wb.writeRows('Recon_Purchases', 16, pr.rows, true);
      await wb.writeRows('Recon_Suppliers', 6, supplierRows(pr.supRows, mo.tol), true);
      await wb.setCells('Recon_Purchases', { G3: `Tax = CGST + SGST + IGST + cess. ${describeMatch(mo)}` });
      recon.purchases = pr.stats;
    } else if (bp.length || raw2b.length) warns.push({ level: 'info', file: '', msg: `Purchase reconciliation skipped: ${bp.length ? 'no GSTR-2B data' : 'no purchase register'} yet` });
    if (bs.length && raw1.length) {
      const notes: string[] = [];
      const ds = extra.decisions?.sales;
      const sr = reconSales(bs, raw1, fy, ny, remarks, ds?.manual ?? manual.Recon_Sales, notes, mo, ds?.keepApart ?? keepApart.Recon_Sales);
      await wb.setCells('Recon_Sales', { G3: describeMatch(mo) });
      notes.forEach(m => warns.push({ level: 'warn', file: 'Recon_Sales', msg: m }));
      await wb.writeRows('Recon_Sales', 16, sr.rows, true);
      await wb.writeRows('Recon_Sales_Monthly', 6, salesMonthRows(sr.monthly), true);
      recon.sales = sr.stats;
    } else if (bs.length || raw1.length) warns.push({ level: 'info', file: '', msg: `Sales reconciliation skipped: ${bs.length ? 'no GSTR-1 data' : 'no sales register'} yet` });
  }

  // fingerprints of the imported data, checked again when the JSON is generated
  const fpRows: any[][] = [['built_at', new Date().toISOString().slice(0, 16).replace('T', ' '), null], ['loader_version', LOADER_VERSION, null]];
  for (const sh of RAW_SHEETS) {
    if (written[sh]) { const f = await fingerprint(written[sh]); fpRows.push([sh, f.hash, f.count]); }
    else if (prevFp[sh]) fpRows.push([sh, ...prevFp[sh]]);
    else { const f = await fingerprint([]); fpRows.push([sh, f.hash, 0]); }
  }
  fpRows.push(['signature', await seal(sealLines(fpRows, gstin, fy)), null]);
  if (wb.has('_Integrity')) await wb.writeRows('_Integrity', 4, fpRows);
  else warns.push({ level: 'warn', file: base?.name ?? '', msg: 'Workbook has no fingerprint sheet (older version) - edit checks will be limited. Rebuild from the Loader template to enable them.' });

  const summary: Record<string, number> = {};
  for (const r of g1rows) if (r.key) summary['GSTR-1 ' + r.key] = r2((summary['GSTR-1 ' + r.key] || 0) + r.amt[0]);
  for (const r of r2b) if (r.key) summary['2B ' + r.key + ' (tax)'] = r2((summary['2B ' + r.key + ' (tax)'] || 0) + r.amt[1] + r.amt[2] + r.amt[3] + r.amt[4]);
  const counts = { 'GSTR-1 documents': g1rows.length, 'HSN lines': agg.size, 'GSTR-3B months': seen3b.size, 'GSTR-2B documents': r2b.length, 'Portal GSTR-9': p9.length ? 1 : 0 };
  return { bytes: await wb.save(), warns, summary, counts, recon };
}

/** remarks and manual Match codes from an earlier workbook (annual or ITC) */
export async function readReconState(baseRd: Reader | null) {
  const remarks = new Map<string, string>();
  const manual: Record<string, Manual> = { Recon_Purchases: new Map(), Recon_Sales: new Map() };
  const keepApart: Record<string, KeepApart> = { Recon_Purchases: new Set(), Recon_Sales: new Set() };
  if (baseRd) for (const [sh, start, codeCol, remCol, idCol] of [['Recon_Purchases', 16, 27, 28, 29], ['Recon_Sales', 16, 21, 22, 23]] as const) {
    if (!baseRd.has(sh)) continue;
    const hdrCode = String((await baseRd.get(sh, colLetterOf(codeCol) + '15')).v ?? '');
    if (hdrCode !== 'Match code') continue;          // workbook from an older template: no manual matching columns
    for (const r of await readRows(baseRd, sh, start)) {
      const id = String(r[idCol - 1] ?? ''), code = String(r[codeCol - 1] ?? '').trim().toUpperCase(), rem = r[remCol - 1];
      if (!id) continue;
      if (rem) { remarks.set(id, String(rem)); for (const part of id.split(/[~,]/)) if (part) remarks.set(part, String(rem)); }
      if (code === 'X') { for (const part of id.split(/[~,]/)) if (part) keepApart[sh].add(part); }
      else if (code) {
        const g = manual[sh].get(code) ?? { b: new Set<string>(), r: new Set<string>() };
        for (const part of id.split(/[~,]/)) { if (part.startsWith('B|')) g.b.add(part); else if (part.startsWith('R|')) g.r.add(part); }
        manual[sh].set(code, g);
      }
    }
  }
  if (baseRd?.has('Recon_Suppliers')) { const rc = String((await baseRd.get('Recon_Suppliers', 'K5')).v ?? '') === 'Remarks' ? 10 : 11; for (const r of await readRows(baseRd, 'Recon_Suppliers', 6)) if (r[0] && r[rc]) remarks.set('S|' + r[0], String(r[rc])); }
  return { remarks, manual, keepApart };
}

// ================================================================ monthly mode: ITC reconciliation workbook
export interface ItcResult { bytes: Uint8Array; warns: Warn[]; recon?: any; periods: string[]; position: Record<string, number[]> }
export async function buildItc(template: Uint8Array, files: InFile[], fy: string, extra: { legalName?: string; gstin?: string; purchases?: any[][] } & ReconExtra): Promise<ItcResult> {
  const warns: Warn[] = [];
  const py = fyShift(fy, -1), ny = fyShift(fy, 1);
  const base = files.find(f => f.kind === 'Workbook' && f.note === 'ITC reconciliation workbook');
  const wb = await Workbook.open(base?.bytes ?? template);
  const baseRd = base ? await Reader.open(base.bytes!) : null;
  await fixReconLayout(wb, baseRd);
  const gstins = new Set(files.filter(f => f.gstin).map(f => f.gstin));
  if (gstins.size > 1) warns.push({ level: 'error', file: '', msg: `Files belong to different GSTINs: ${[...gstins].join(', ')}` });
  // 2B
  const b2 = files.filter(f => f.kind === 'GSTR-2B').map(f => ({ name: f.name, json: f.json }));
  let raw2b: any[][];
  let periods: string[];
  if (b2.length) {
    const r = process2b(b2, fy, py, ny, warns);
    raw2b = r.rows.map(row); periods = r.periods;
    await wb.writeRows('RAW_2B', 6, raw2b);
  } else {
    raw2b = baseRd ? await readRows(baseRd, 'RAW_2B', 6) : [];
    periods = [...new Set(raw2b.map(r => String(r[8] ?? '')).filter(Boolean))].sort((a, b) => (a.slice(2) + a.slice(0, 2)).localeCompare(b.slice(2) + b.slice(0, 2)));
  }
  // books
  if (extra.purchases) await wb.fillTable('Books_Purchase', 6, extra.purchases, 18);
  const bp = (extra.purchases ?? (baseRd ? await readRows(baseRd, 'Books_Purchase', 6) : []))
    .filter(r => r && [12, 13, 14, 15, 16].some(i => typeof r[i] === 'number' && r[i] !== 0));
  if (!raw2b.length) warns.push({ level: 'error', file: '', msg: 'No GSTR-2B loaded' });
  if (!bp.length) warns.push({ level: 'error', file: '', msg: 'No purchase register loaded' });
  const position: Record<string, number[]> = {};
  let recon: any;
  if (raw2b.length && bp.length) {
    const { remarks, manual, keepApart } = await readReconState(baseRd);
    const notes: string[] = [];
    const mo = extra.match ?? DEFAULT_MATCH, dp = extra.decisions?.purchases;
    const pr = reconPurchases(bp, raw2b, fy, ny, remarks, dp?.manual ?? manual.Recon_Purchases, notes, mo, dp?.keepApart ?? keepApart.Recon_Purchases);
    await wb.setCells('Recon_Purchases', { G3: `Tax = CGST + SGST + IGST + cess. ${describeMatch(mo)}` });
    notes.forEach(m => warns.push({ level: 'warn', file: 'Recon_Purchases', msg: m }));
    await wb.writeRows('Recon_Purchases', 16, pr.rows, true);
    await wb.writeRows('Recon_Suppliers', 6, supplierRows(pr.supRows, mo.tol), true);
    recon = pr.stats;
    // ITC position by head (CGST, SGST, IGST, Cess)
    const H = [0, 1, 2, 3], B = (r: any[], i: number) => +r[9 + i] || 0, R = (r: any[], i: number) => +r[14 + i] || 0;
    const sum = (f: (r: any[]) => boolean, v: (r: any[], i: number) => number) => H.map(i => r2(pr.rows.filter(f).reduce((a, r) => a + v(r, i), 0)));
    const has2b = (r: any[]) => r[5] !== '' && r[5] !== null && r[22] !== 'N';
    position.available = sum(has2b, R);
    position.books = sum(r => r[4] !== '' && r[4] !== null && !String(r[0]).startsWith('Not compared'), B);
    position.claimable = sum(r => String(r[0]).startsWith('Matched') && r[22] !== 'N', (r, i) => Math.min(B(r, i), R(r, i)));
    position.booksOnly = sum(r => r[0] === 'In books, not in 2B', B);
    position.twoBOnly = sum(r => r[0] === 'In 2B, not in books', R);
    position.valueDiff = sum(r => r[0] === 'Matched - value difference', (r, i) => B(r, i) - R(r, i));
    // the Summary amounts are formulas over Recon_Purchases (position[] is the same figure, for the Loader's own display)
    await wb.setCells('Summary', summaryCells());
    // month-wise: 2B side by 2B period, books side by invoice date
    const tx = (r: any[], o: number) => [0, 1, 2, 3].reduce((a, i) => a + (+r[o + i] || 0), 0);
    const per = (p: string) => p.replace('/', '');
    const mm = new Map<string, number[]>();
    const bump = (k: string, j: number, v: number) => { if (!k) return; const a = mm.get(k) ?? [0, 0, 0, 0, 0]; a[j] += v; mm.set(k, a); };
    for (const r of pr.rows) {
      const st = String(r[0]); if (st.startsWith('Not compared')) continue;
      const p2 = per(String(r[20] ?? ''));
      const dB = String(r[6] ?? ''), mB = /^\d{2}-\d{2}-\d{4}$/.test(dB) ? dB.slice(3, 5) + dB.slice(6) : '';
      if (r[5] && r[22] !== 'N') bump(p2, 0, tx(r, 14));
      if (st.startsWith('Matched') && r[22] !== 'N') bump(p2, 1, [0, 1, 2, 3].reduce((a, i) => a + Math.min(+r[9 + i] || 0, +r[14 + i] || 0), 0));
      if (st === 'In 2B, not in books') bump(p2, 2, tx(r, 14));
      if (r[4]) bump(mB, 3, tx(r, 9));
      if (st === 'In books, not in 2B') bump(mB, 4, tx(r, 9));
    }
    const MN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const keys = [...mm.keys()].sort((a, b) => (a.slice(2) + a.slice(0, 2)).localeCompare(b.slice(2) + b.slice(0, 2)));
    if (wb.has('Month_wise')) {
      await wb.writeRows('Month_wise', 6, monthRows(keys, k => `${MN[+k.slice(0, 2) - 1]} ${k.slice(2)}`), true);
      await wb.setCells('Month_wise', { H5: 'Period (used by the formulas)' });
    }
  }
  const fmtP = (p: string) => `${p.slice(0, 2)}/${p.slice(2)}`;
  const legal = extra.legalName || (baseRd ? String((await baseRd.get('Summary', 'C4')).v ?? '') : '');
  await wb.setCells('Summary', { C4: legal || null, C5: extra.gstin || [...gstins][0] || (baseRd ? String((await baseRd.get('Summary', 'C5')).v ?? '') : ''),
    C6: periods.length ? (periods.length > 3 ? `${fmtP(periods[0])} to ${fmtP(periods[periods.length - 1])} (${periods.length} months)` : periods.map(fmtP).join(', ')) : '',
    C7: fy, C9: new Date().toISOString().slice(0, 16).replace('T', ' ') });
  return { bytes: await wb.save(), warns, recon, periods, position };
}

/**
 * Review step: run the reconciliations in the browser without writing a workbook.
 * Uses the same raw rows as the build (new 2B / GSTR-1 files, else those in the dropped workbook) and that workbook's codes and remarks.
 */
export async function reconPreview(files: InFile[], fy: string, mode: 'itc' | 'annual', extra: { sales?: any[][]; purchases?: any[][] } & ReconExtra) {
  const warns: Warn[] = [];
  const py = fyShift(fy, -1), ny = fyShift(fy, 1);
  const base = files.find(f => f.kind === 'Workbook' && (mode === 'itc') === (f.note === 'ITC reconciliation workbook'));
  const baseRd = base ? await Reader.open(base.bytes!) : null;
  const fromBase = async (sh: string) => (baseRd && baseRd.has(sh) ? await readRows(baseRd, sh, 6) : []);
  const b2 = files.filter(f => f.kind === 'GSTR-2B').map(f => ({ name: f.name, json: f.json }));
  const raw2b = b2.length ? process2b(b2, fy, py, ny, warns).rows.map(row) : await fromBase('RAW_2B');
  const g1 = files.filter(f => f.kind === 'GSTR-1').map(f => ({ name: f.name, json: f.json }) as G1File);
  const raw1 = mode === 'annual' ? (g1.length ? processGstr1(g1, fy, ny, warns).rows.map(row) : await fromBase('RAW_GSTR1')) : [];
  const nonzero = (rows: any[][], cols: number[]) => rows.filter(r => r && cols.some(i => typeof r[i] === 'number' && r[i] !== 0));
  const bp = nonzero(extra.purchases ?? await fromBase('Books_Purchase'), [12, 13, 14, 15, 16]);
  const bs = mode === 'annual' ? nonzero(extra.sales ?? await fromBase('Books_Sales'), [11, 12, 13, 14, 15]) : [];
  const { remarks, manual, keepApart } = await readReconState(baseRd);
  const mo = extra.match ?? DEFAULT_MATCH;
  const notes: string[] = [];
  const out: { purchases?: ReturnType<typeof reconPurchases>; sales?: ReturnType<typeof reconSales>; notes: string[]; saved: { purchases: Decisions; sales: Decisions } } =
    { notes, saved: { purchases: { manual: manual.Recon_Purchases, keepApart: keepApart.Recon_Purchases }, sales: { manual: manual.Recon_Sales, keepApart: keepApart.Recon_Sales } } };
  if (bp.length && raw2b.length) { const d = extra.decisions?.purchases; out.purchases = reconPurchases(bp, raw2b, fy, ny, remarks, d?.manual ?? manual.Recon_Purchases, notes, mo, d?.keepApart ?? keepApart.Recon_Purchases); }
  if (bs.length && raw1.length) { const d = extra.decisions?.sales; out.sales = reconSales(bs, raw1, fy, ny, remarks, d?.manual ?? manual.Recon_Sales, notes, mo, d?.keepApart ?? keepApart.Recon_Sales); }
  return out;
}

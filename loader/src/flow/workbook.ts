// GST ledger flow workbook: detail sheets carry every input row with its classification; the flow statements,
// differences and checks are SUMIFS formulas over them (cached values are written too, Excel recalculates on open).
import { C, SheetSpec, col, serial, writeBook } from '../xlsxbook';
import { BUCKETS, BLOCK_LABEL, Block, INTERNAL, PORTAL_LABEL, UNMAPPED, ACC_TYPE_LABEL, bucketOf } from './types';
import { HEAD_KEYS, HeadK, Result, bucketLabel, ymLabel } from './engine';
import { FLOW_GUIDE } from './guide';
import { ledgerLabel, reconcile } from './recon';

export interface AuditRow { when: string; change: string; record: string; from: string; to: string; remark: string }
export interface FlowMeta { gstin: string; legalName: string; fy: string; from: string; to: string; files: { name: string; kind: string; note: string }[]; audit: AuditRow[]; builtAt: string }

const HEADS_TXT: Record<HeadK, string> = { IGST: 'IGST', CGST: 'CGST', SGST: 'SGST', CESS: 'Cess', ALL: 'Not split by head' };
const BLOCKS: Block[] = ['itc', 'liab', 'cash', 'fee'];
const N = (v: number) => Math.round(v * 100) / 100;

export async function flowWorkbook(r: Result, meta: FlowMeta): Promise<Uint8Array> {
  // short references for every record
  const bref = new Map<string, string>(), pref = new Map<string, string>();
  r.books.forEach((e, i) => bref.set(e.id, `BL${String(i + 1).padStart(5, '0')}`));
  // one portal reference per file line (a line has several heads / components)
  const pLine = new Map<string, string>();
  r.portal.forEach(e => { const k = `${e.file}#${e.row}`; if (!pLine.has(k)) pLine.set(k, `PL${String(pLine.size + 1).padStart(5, '0')}`); pref.set(e.id, pLine.get(k)!); });
  const rc = reconcile(r);
  const itemsOf = new Map<string, string[]>();
  for (const ev of rc.events) for (const id of [...ev.books, ...ev.portal]) { const a = itemsOf.get(id) || []; if (!a.includes(ev.id)) a.push(ev.id); itemsOf.set(id, a); }
  const period = `${ymLabel(meta.from)} to ${ymLabel(meta.to)}`;
  const title = (t: string): C[][] => [[{ v: t, s: 'title' }], [{ v: `${meta.legalName || 'Client'} · GSTIN ${meta.gstin || '-'} · FY ${meta.fy} · ${period} · Prepared by James & James LLP`, s: 'sub' }], []];
  const hdr = (cols: string[]): C[] => cols.map(c => ({ v: c, s: 'hdr' }));

  // ---- Books_Lines
  const BL_H = 4; // header row
  const books: C[][] = [...title('Books ledger lines (GST accounts)'), hdr(['Ref', 'Source file', 'Row', 'Date', 'Month', 'Compared with return period', 'In period', 'Voucher type', 'Voucher no.', 'Party / particulars', 'Account', 'Account type', 'Head', 'Debit', 'Credit', 'Effect on balance', 'Block', 'Flow row code', 'Flow row', 'Row amount', 'Classified by', 'Reconciling items', 'Narration'])];
  r.books.forEach((e, i) => {
    const row = BL_H + 1 + i, l = e.line, sign = bucketOf(e.bucket)?.sign ?? 1;
    books.push([bref.get(e.id), l.file, l.row, { v: serial(l.date), s: 'date' }, ymLabel(l.date.slice(6) + '-' + l.date.slice(3, 5)), ymLabel(e.month), e.where, l.vchType, l.vchNo, l.party, l.account, ACC_TYPE_LABEL[e.accType], e.head,
      l.dr || null, l.cr || null, { f: e.block === 'liab' ? `O${row}-N${row}` : `N${row}-O${row}`, v: e.effect }, e.block, e.bucket, bucketLabel(e.bucket),
      { f: sign === 1 ? `P${row}` : `-P${row}`, v: N(e.effect * sign) }, e.rule, (itemsOf.get(e.id) || []).join(', '), l.narration]);
  });
  // ---- Portal_Lines
  const portal: C[][] = [...title('Portal ledger lines'), hdr(['Ref', 'Ledger', 'Source file', 'Row', 'Date', 'Reference', 'Return period', 'Compared in month', 'In period', 'Description', 'Dr / Cr', 'Ledger used', 'Head', 'Component', 'Amount', 'Effect on balance', 'Block', 'Flow row code', 'Flow row', 'Row amount', 'Reconciling items'])];
  r.portal.forEach((e, i) => {
    const row = BL_H + 1 + i, sign = bucketOf(e.bucket)?.sign ?? 1, es = e.amount ? Math.sign(e.effect / e.amount) : 1;
    portal.push([pref.get(e.id), PORTAL_LABEL[e.kind], e.file, e.row, { v: serial(e.date), s: 'date' }, e.ref, ymLabel(e.month), ymLabel(e.month), e.block ? e.where : 'not in the flow', e.desc, e.dir, e.used, e.head, e.comp,
      e.amount, { f: es === 1 ? `O${row}` : `-O${row}`, v: e.effect }, e.block || '', e.block ? e.bucket : '', e.block ? bucketLabel(e.bucket) : 'Discharge of interest / fee (shown in the cash ledger)',
      e.block ? { f: sign === 1 ? `P${row}` : `-P${row}`, v: N(e.effect * sign) } : null, (itemsOf.get(e.id) || []).join(', ')]);
  });
  // ---- Balances
  const bal: C[][] = [...title('Opening balances used in the flow'), hdr(['Side', 'Block', 'Head', 'Ledger / account', 'Source', 'Amount (in the block\'s direction)', 'As printed or entered'])];
  for (const b of r.balances) bal.push([b.side, b.block, b.head, b.what, b.source, b.amount, b.raw ?? null]);
  bal.push([], [{ v: 'Liability balances are shown as payable (+). Ledger exports show debit balances as + and credit balances as −, so a credit balance of an output tax account appears here with the sign changed.', s: 'note' }]);

  // ---- flow statements: the same rows on all three sheets
  const rowKeys = (b: Block) => {
    const present = new Set([...Object.keys(r.flowBooks.rows), ...Object.keys(r.flowPortal.rows)]);
    const ks = BUCKETS.filter(x => x.block === b).map(x => x.id).filter(k => present.has(k) || b !== 'fee');
    for (const extra of [INTERNAL, UNMAPPED]) if (present.has(`${b}:${extra}`)) ks.push(`${b}:${extra}`);
    return ks;
  };
  const HC = HEAD_KEYS.map((_, i) => col(i + 1)); // B..F
  type Layout = { block: Block; open?: number; rows: { key: string; at: number }[]; close?: number };
  const layout: Layout[] = [];
  const flowSheet = (side: 'Books' | 'Portal' | 'Difference'): C[][] => {
    const st = side === 'Books' ? r.flowBooks : side === 'Portal' ? r.flowPortal : null;
    const out: C[][] = [...title(side === 'Difference' ? 'GST flow: difference (books − portal)' : `GST flow as per ${side === 'Books' ? 'books' : 'the GST portal'}`)];
    out.push([{ v: side === 'Books' ? 'Built only from the books ledgers (Books_Lines) and the opening balances (Balances). Books by voucher date.' : side === 'Portal' ? 'Built only from the credit ledger, cash ledger and liability register (Portal_Lines) and their opening balances. Portal by return period.' : 'Each cell is Flow_Books minus Flow_Portal. The closing difference is the sum of the row differences. Recon_Items explains every row.', s: 'note' }]);
    out.push(hdr(['Flow row', ...HEAD_KEYS.map(h => HEADS_TXT[h]), 'Total', 'Row code']));
    const lay = side === 'Books' || !layout.length;
    for (const b of BLOCKS) {
      const L: Layout = lay ? { block: b, rows: [] } : layout.find(x => x.block === b)!;
      out.push([{ v: BLOCK_LABEL[b], s: 'sec' }, ...HEAD_KEYS.map(() => ({ v: '', s: 'sec' } as C)), { v: '', s: 'sec' }, { v: '', s: 'sec' }]);
      const cellF = (key: string, h: HeadK, r1: number): C => {
        if (side === 'Difference') return { f: `Flow_Books!${col(HEAD_KEYS.indexOf(h) + 1)}${r1}-Flow_Portal!${col(HEAD_KEYS.indexOf(h) + 1)}${r1}`, v: N((r.flowBooks.rows[key]?.[h] || 0) - (r.flowPortal.rows[key]?.[h] || 0)) };
        const sh = side === 'Books' ? 'Books_Lines' : 'Portal_Lines';
        const hc = side === 'Books' ? 'M' : 'M', kc = 'R', vc = 'T', inc = side === 'Books' ? 'G' : 'I';
        return { f: `SUMIFS(${sh}!$${vc}:$${vc},${sh}!$${kc}:$${kc},$H${r1},${sh}!$${hc}:$${hc},"${h}",${sh}!$${inc}:$${inc},"in")`, v: N(st!.rows[key]?.[h] || 0) };
      };
      if (b !== 'fee') {
        const r1 = out.length + 1; if (lay) L.open = r1;
        out.push(['Opening balance', ...HEAD_KEYS.map(h => side === 'Difference'
          ? { f: `Flow_Books!${col(HEAD_KEYS.indexOf(h) + 1)}${r1}-Flow_Portal!${col(HEAD_KEYS.indexOf(h) + 1)}${r1}`, v: N(r.flowBooks.opening[b][h] - r.flowPortal.opening[b][h]), s: 'numB' } as C
          : { f: `SUMIFS(Balances!$F:$F,Balances!$A:$A,"${side}",Balances!$B:$B,"${b}",Balances!$C:$C,"${h}")`, v: N(st!.opening[b][h]), s: 'numB' } as C), { f: `SUM(B${r1}:F${r1})`, s: 'numB', v: 0 }, 'opening']);
      }
      for (const key of rowKeys(b)) {
        const r1 = out.length + 1; if (lay) L.rows.push({ key, at: r1 });
        const sign = bucketOf(key)?.sign ?? 1;
        out.push([`${b === 'fee' ? '' : sign === 1 ? 'Add: ' : 'Less: '}${bucketLabel(key)}`, ...HEAD_KEYS.map(h => cellF(key, h, r1)), { f: `SUM(B${r1}:F${r1})`, v: 0 }, key]);
      }
      const r1 = out.length + 1; if (lay) L.close = r1;
      const terms = (c: string) => [...(L.open ? [`${c}${L.open}`] : []), ...L.rows.map(x => `${(bucketOf(x.key)?.sign ?? 1) === 1 ? '+' : '-'}${c}${x.at}`)].join('').replace(/^\+/, '');
      const src = side === 'Books' ? r.flowBooks : side === 'Portal' ? r.flowPortal : null;
      out.push([b === 'fee' ? 'Total' : 'Closing balance', ...HEAD_KEYS.map((h, i) => ({ f: terms(HC[i]) || '0', s: 'numTot', v: src ? N(src.closing[b][h]) : N(r.flowBooks.closing[b][h] - r.flowPortal.closing[b][h]) } as C)), { f: `SUM(B${r1}:F${r1})`, s: 'numTot', v: 0 }, '']);
      out.push([]);
      if (lay) layout.push(L);
    }
    return out;
  };
  // fill the Total column cached values
  const fixTotals = (rows: C[][]) => rows.forEach(rw => { const t = rw[6] as any; if (t && typeof t === 'object' && t.f?.startsWith('SUM(B')) t.v = N(rw.slice(1, 6).reduce((s: number, c: any) => s + (typeof c === 'object' && c ? +c.v || 0 : 0), 0)); });
  const fBooks = flowSheet('Books'), fPortal = flowSheet('Portal'), fDiff = flowSheet('Difference');
  [fBooks, fPortal, fDiff].forEach(fixTotals);
  fDiff.push([{ v: 'The Reconciliation sheet takes the closing as per books to the closing as per portal, item by item (Recon_Items).', s: 'note' }]);

  // ---- Recon_Items: one row per event; amounts are the adjustment that takes the books closing to the portal closing
  const IH = (h: HeadK) => col(7 + HEAD_KEYS.indexOf(h)); // H..L
  const items: C[][] = [...title('Reconciling items: books to portal'), hdr(['No.', 'Ledger', 'Effect on closing', 'Status', 'What happened', 'Return period(s)', 'Vouchers', ...HEAD_KEYS.map(h => HEADS_TXT[h]), 'Total', 'Amount involved (no effect on closing)', 'Action', 'Flow row(s)', 'Books lines', 'Portal lines'])];
  const evRow = new Map<string, number>();
  const evLine = (ev: typeof rc.events[number]) => {
    const row = items.length + 1; evRow.set(ev.id, row);
    const adj = (h: HeadK) => (ev.affects ? N(-ev.effect[h]) : 0);
    items.push([ev.id, ledgerLabel(ev.block), ev.affects ? 'Yes' : 'No', ev.explained ? 'Explained' : 'Unexplained', { v: ev.what, s: 'wrap' }, ev.months.map(ymLabel).join(', '), { v: ev.vouchers, s: 'wrap' },
      ...HEAD_KEYS.map(h => (Math.abs(adj(h)) > 0.004 ? adj(h) : null)), { f: `SUM(H${row}:L${row})`, v: N(HEAD_KEYS.reduce((t, h) => t + adj(h), 0)) },
      ev.affects ? null : ev.involved || null, { v: ev.action, s: 'wrap' }, ev.rows.join(', '),
      { v: [...new Set(ev.books.map(id => bref.get(id)).filter(Boolean))].join(', '), s: 'wrap' }, { v: [...new Set(ev.portal.map(id => pref.get(id)).filter(Boolean))].join(', '), s: 'wrap' }]);
  };
  rc.events.filter(e => e.affects).forEach(evLine);
  const ne = rc.events.filter(e => !e.affects);
  if (ne.length) { items.push([], [{ v: 'Differences that reversed within the period: no effect on the closing', s: 'sec' }, ...Array(17).fill(0).map(() => ({ v: '', s: 'sec' } as C))]); ne.forEach(evLine); }
  items.push([], [{ v: 'Amounts are what to add to the closing as per books to reach the closing as per portal. Unexplained items are differences the records do not explain: correct the classification of the lines in the Loader, or the books / returns themselves.', s: 'note' }]);

  // ---- Reconciliation: closing as per books → items → closing as per portal (formulas over the flow sheets and Recon_Items)
  const rec: C[][] = [...title('Reconciliation: closing balance as per books to the GST portal'), hdr(['', ...HEAD_KEYS.map(h => HEADS_TXT[h]), 'Total', 'Action', 'Item'])];
  for (const st of rc.statements) {
    const L = layout.find(x => x.block === st.block)!;
    rec.push([{ v: ledgerLabel(st.block), s: 'sec' }, ...Array(8).fill(0).map(() => ({ v: '', s: 'sec' } as C))]);
    const rb = rec.length + 1;
    rec.push([{ v: 'Closing balance as per books', s: 'bold' }, ...HEAD_KEYS.map((h, i) => ({ f: `Flow_Books!${HC[i]}${L.close}`, v: N(st.books[h]), s: 'numB' } as C)), { f: `SUM(B${rb}:F${rb})`, v: N(HEAD_KEYS.reduce((t, h) => t + st.books[h], 0)), s: 'numB' }]);
    const lr: number[] = [];
    for (const ev of st.lines) {
      const row = rec.length + 1, er = evRow.get(ev.id)!; lr.push(row);
      rec.push([{ v: `${ev.explained ? '' : 'UNEXPLAINED: '}${ev.what}`, s: 'wrap' }, ...HEAD_KEYS.map(h => ({ f: `Recon_Items!${IH(h)}${er}`, v: N(-ev.effect[h]) } as C)), { f: `SUM(B${row}:F${row})`, v: N(HEAD_KEYS.reduce((t, h) => t - ev.effect[h], 0)) }, { v: ev.action, s: 'wrap' }, ev.id]);
    }
    if (!st.lines.length) rec.push([{ v: 'No difference: books and portal agree.', s: 'note' }]);
    const rp = rec.length + 1;
    rec.push([{ v: 'Closing balance as per portal', s: 'bold' }, ...HEAD_KEYS.map((h, i) => ({ f: `Flow_Portal!${HC[i]}${L.close}`, v: N(st.portal[h]), s: 'numTot' } as C)), { f: `SUM(B${rp}:F${rp})`, v: N(HEAD_KEYS.reduce((t, h) => t + st.portal[h], 0)), s: 'numTot' }]);
    const ru = rec.length + 1, led = ledgerLabel(st.block);
    rec.push([{ v: 'Of which unexplained', s: 'note' }, ...HEAD_KEYS.map(h => ({ f: `SUMIFS(Recon_Items!$${IH(h)}:$${IH(h)},Recon_Items!$B:$B,"${led}",Recon_Items!$C:$C,"Yes",Recon_Items!$D:$D,"Unexplained")`, v: N(-st.unexplained[h]) } as C)), { f: `SUM(B${ru}:F${ru})`, v: N(-HEAD_KEYS.reduce((t, h) => t + st.unexplained[h], 0)) }]);
    const rn = rec.length + 1;
    rec.push([{ v: 'Not reconciled (should be nil)', s: 'note' }, ...HEAD_KEYS.map((h, i) => ({ f: `${HC[i]}${rb}${lr.map(x => `+${HC[i]}${x}`).join('')}-${HC[i]}${rp}`, v: N(st.gap[h]) } as C)), { f: `SUM(B${rn}:F${rn})`, v: N(HEAD_KEYS.reduce((t, h) => t + st.gap[h], 0)) }]);
    rec.push([]);
  }
  rec.push([{ v: `Differences that reversed within the period (timing between months, amounts shown net on the portal) do not change the closing: they are listed at the end of Recon_Items (${ne.length}).`, s: 'note' }]);
  // ---- Cross_Checks
  const chk: C[][] = [...title('Cross-checks between the portal ledgers and with GSTR-3B'), hdr(['Check', 'Return period', 'Head', 'A', 'Amount A', 'B', 'Amount B', 'Difference', 'Status', 'Detail'])];
  r.checks.forEach((c, i) => { const row = BL_H + 1 + i; chk.push([c.area, ymLabel(c.period), c.head, c.aLabel, c.a, c.bLabel, c.b, { f: `E${row}-G${row}`, v: c.diff }, { f: `IF(ABS(H${row})<=1,"OK","Check")`, v: c.ok ? 'OK' : 'Check', s: c.ok ? 'ok' : 'bad' }, c.note || '']); });
  if (!r.checks.length) chk.push(['No checks: load at least two portal ledgers (and GSTR-3B PDFs for the 3B checks).']);
  if (r.warns.length) { chk.push([], [{ v: 'Messages', s: 'bold' }]); r.warns.forEach(w => chk.push([w.level, { v: w.msg, s: 'wrap' }])); }
  // ---- Month_wise
  const mw: C[][] = [...title('Month-wise: books, portal and difference'), hdr(['Block', 'Flow row', 'Head', ...r.months.flatMap(m => [`${ymLabel(m)} books`, `${ymLabel(m)} portal`, `${ymLabel(m)} diff`]), 'Row code'])];
  const keyHeads = new Map<string, Set<HeadK>>();
  for (const e of r.books) if (e.block && e.where === 'in') { const k = e.bucket === INTERNAL || e.bucket === UNMAPPED ? `${e.block}:${e.bucket}` : e.bucket; (keyHeads.get(k) || keyHeads.set(k, new Set()).get(k)!).add(e.head as HeadK); }
  for (const e of r.portal) if (e.block && e.where === 'in') (keyHeads.get(e.bucket) || keyHeads.set(e.bucket, new Set()).get(e.bucket)!).add(e.head);
  const monthsCol = (m: string) => ymLabel(m);
  for (const b of BLOCKS) for (const key of rowKeys(b)) for (const h of HEAD_KEYS) {
    if (!keyHeads.get(key)?.has(h)) continue;
    const row = mw.length + 1, cells: C[] = [BLOCK_LABEL[b], bucketLabel(key), h];
    r.months.forEach((m, i) => {
      const bc = col(3 + i * 3), pc = col(4 + i * 3);
      const bv = N(r.books.filter(e => e.where === 'in' && e.head === h && e.month === m && (e.bucket === key || `${e.block}:${e.bucket}` === key)).reduce((s, e) => s + e.effect * (bucketOf(e.bucket)?.sign ?? 1), 0));
      const pv = N(r.portal.filter(e => e.where === 'in' && e.head === h && e.month === m && e.bucket === key && e.block).reduce((s, e) => s + e.effect * (bucketOf(e.bucket)?.sign ?? 1), 0));
      cells.push({ f: `SUMIFS(Books_Lines!$T:$T,Books_Lines!$R:$R,$${col(3 + r.months.length * 3)}${row},Books_Lines!$M:$M,$C${row},Books_Lines!$F:$F,"${monthsCol(m)}",Books_Lines!$G:$G,"in")`, v: bv },
        { f: `SUMIFS(Portal_Lines!$T:$T,Portal_Lines!$R:$R,$${col(3 + r.months.length * 3)}${row},Portal_Lines!$M:$M,$C${row},Portal_Lines!$H:$H,"${monthsCol(m)}",Portal_Lines!$I:$I,"in")`, v: pv },
        { f: `${bc}${row}-${pc}${row}`, v: N(bv - pv), s: 'numB' });
    });
    cells.push(key); mw.push(cells);
  }
  // ---- Audit
  const au: C[][] = [...title('Your changes to the records (reclassifications, periods, balances, account types)'), hdr(['When', 'Change', 'Record', 'From', 'To', 'Remark'])];
  meta.audit.forEach(a => au.push([a.when, a.change, a.record, a.from, a.to, { v: a.remark, s: 'wrap' }]));
  if (!meta.audit.length) au.push(['No changes: every classification is the tool\'s own, from the account types confirmed in the Loader.']);
  // ---- Read Me
  const rm: C[][] = [[{ v: 'GST ledger flow: books vs GST portal', s: 'title' }], [{ v: `${meta.legalName || 'Client'} · GSTIN ${meta.gstin || '-'} · FY ${meta.fy} · ${period}`, s: 'bold' }], [{ v: `Built by the James & James GST Reco Loader on ${meta.builtAt}. Prepared by James & James LLP.`, s: 'sub' }], [],
    [{ v: 'How to read this workbook', s: 'sec' }, { v: '', s: 'sec' }],
    ['Flow_Books', 'Opening, movements and closing of input tax credit, tax liability, the electronic cash ledger and interest / fees, from the books only.'],
    ['Flow_Portal', 'The same rows from the portal ledgers only (credit ledger, cash ledger, liability register), by return period.'],
    ['Reconciliation', 'For each ledger: closing balance as per books, each reconciling item, closing balance as per portal. "Not reconciled" must be nil; "Of which unexplained" shows what the records do not explain.'],
    ['Flow_Difference', 'Books minus portal, row by row.'],
    ['Recon_Items', 'One row per reconciling item, worked out from the records: what happened, the vouchers, the amount by tax head and what to do. Items that reversed within the period (no effect on the closing) are at the end.'],
    ['Cross_Checks', 'The portal ledgers against each other (cash ledger vs liability paid in cash; credit ledger vs liability paid by ITC) and against GSTR-3B where its PDFs were loaded.'],
    ['Month_wise', 'Books, portal and difference for each row, head and month.'],
    ['Books_Lines / Portal_Lines', 'Every input line with its flow row, how it was classified and the reconciling items it belongs to. All the figures above are SUMIFS over these sheets.'],
    ['Balances', 'Opening balances: entered, printed in the files, or carried from entries before the period.'],
    ['Audit', 'Every change you made in the Loader, with your remark.'], [],
    [{ v: 'Basis', s: 'sec' }, { v: '', s: 'sec' }],
    [{ v: 'Books are taken by voucher date. Portal entries are taken by return period (an April filing of the March return counts in March). Where a books line was moved to another return period in the Loader, that period is used to compare it (column "Compared with return period"), and the change is in Audit.', s: 'wrap' }], [],
    [{ v: 'Files used', s: 'sec' }, { v: '', s: 'sec' }], ...meta.files.map(f => [f.name, `${f.kind}${f.note ? ' · ' + f.note : ''}`] as C[]), [],
    [{ v: 'How to export each report', s: 'sec' }, { v: '', s: 'sec' }],
    ...FLOW_GUIDE.flatMap(g => [[{ v: `${g.source === 'portal' ? 'GST portal' : g.source === 'tally' ? 'Tally' : g.source === 'zoho' ? 'Zoho Books' : 'Other'}: ${g.title}`, s: 'bold' }] as C[], ...g.steps.map((s, i) => [`  ${i + 1}.`, { v: s, s: 'wrap' }] as C[]), ...(g.note ? [['', { v: g.note, s: 'note' }] as C[]] : [])]),
  ];
  const sheets: SheetSpec[] = [
    { name: 'Read Me', rows: rm, widths: [28, 120] },
    { name: 'Reconciliation', rows: rec, widths: [70, 14, 14, 14, 12, 14, 14, 60, 6], freeze: { row: BL_H, col: 1 }, tab: '2C388E' },
    { name: 'Flow_Books', rows: fBooks, widths: [52, 16, 16, 16, 16, 18, 18, 14], freeze: { row: 5, col: 1 }, tab: '2C388E' },
    { name: 'Flow_Portal', rows: fPortal, widths: [52, 16, 16, 16, 16, 18, 18, 14], freeze: { row: 5, col: 1 }, tab: '2C388E' },
    { name: 'Flow_Difference', rows: fDiff, widths: [70, 16, 16, 16, 16, 18, 18, 16], freeze: { row: 5, col: 1 }, tab: '2C388E' },
    { name: 'Recon_Items', rows: items, widths: [6, 20, 9, 12, 60, 16, 44, 13, 13, 13, 11, 14, 14, 16, 50, 26, 30, 30], freeze: { row: BL_H, col: 1 }, filter: { row: BL_H, cols: 18 } },
    { name: 'Cross_Checks', rows: chk, widths: [44, 12, 8, 30, 15, 32, 15, 14, 9, 80], freeze: { row: BL_H }, filter: { row: BL_H, cols: 10 } },
    { name: 'Month_wise', rows: mw, widths: [24, 36, 8, ...r.months.flatMap(() => [13, 13, 13]), 14], freeze: { row: BL_H, col: 3 } },
    { name: 'Books_Lines', rows: books, widths: [10, 26, 6, 11, 8, 12, 8, 14, 14, 30, 26, 22, 7, 13, 13, 13, 6, 14, 30, 13, 40, 14, 40], freeze: { row: BL_H, col: 1 }, filter: { row: BL_H, cols: 23 } },
    { name: 'Portal_Lines', rows: portal, widths: [10, 26, 26, 6, 11, 20, 10, 10, 8, 32, 6, 8, 7, 9, 13, 13, 6, 14, 30, 13, 14], freeze: { row: BL_H, col: 1 }, filter: { row: BL_H, cols: 21 } },
    { name: 'Balances', rows: bal, widths: [8, 6, 8, 34, 50, 18, 16], freeze: { row: BL_H } },
    { name: 'Audit', rows: au, widths: [18, 26, 40, 24, 24, 60], freeze: { row: BL_H } },
  ];
  return writeBook(sheets);
}

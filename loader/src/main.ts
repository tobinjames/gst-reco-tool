import * as pdfjs from 'pdfjs-dist/build/pdf.mjs';
import * as pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs';
import templateUrl from '../template/GST_Annual_Return_Workbook.xlsx?inline';
import itcTemplateUrl from '../template/ITC_Recon_Template.xlsx?inline';
import { InFile, buildItc, buildWorkbook, classify } from './build';
import { linesFromItems, parse3b } from './gstr3b';
import { fyShift } from './util';
import { generate9 } from './gen9';
import { Review } from './review';
import { FlowTab } from './flow/ui';
import { applyAliases, supplierSuggestions, suppliersFrom2B } from './suppliers';
import { FieldDef, Mapping, PURCHASE_FIELDS, RegKind, RegSheet, colsOf, fieldsOf, pickRegister, detectHeader, guessMappingFull, isMapped, looksColumnar, normalise, readTable, unusedAmountCols } from './register';

// run pdf.js in the page itself (no separate worker file) so the tool works when opened straight from disk
(globalThis as any).pdfjsWorker = pdfWorker;

const $ = (id: string) => document.getElementById(id)! as any;
// "where do I get these files" guide: tabs
document.querySelectorAll('button.gtab').forEach((b: any) => (b.onclick = () => {
  document.querySelectorAll('.gtab, .gpane').forEach((x: any) => x.classList.toggle('on', x.dataset.g === b.dataset.g));
}));
let files: InFile[] = [];
type Tab = 'itc' | 'annual' | 'flow';
type Mode = 'itc' | 'annual';
let tab: Tab = (() => { try { const t = localStorage.getItem('jjmode') as Tab; return ['itc', 'annual', 'flow'].includes(t) ? t : 'annual'; } catch { return 'annual'; } })();
const HERO: Record<Tab, [string, string, string]> = {
  itc: ['Input <em>reconciliation</em>', 'GSTR-2B against the purchase register, month by month: matched, books only, 2B only and value differences, with your match codes and remarks kept.', 'GSTR-2B JSON(s) for the months you are reconciling · purchase register (.xlsx / .csv, any format) · optionally the previous ITC reconciliation workbook to keep match codes and remarks'],
  annual: ['Annual return <em>loader</em>', 'GSTR-1, 2B, 3B and the registers into the GSTR-9 / 9C workbook, and the finished workbook back into the GSTR-9 upload JSON.', 'GSTR-1 JSON / ZIP (previous, current and next FY) · GSTR-2B JSON (this FY + next FY Apr–Nov) · GSTR-3B PDFs · portal GSTR-9 JSON · sales / purchase registers (.xlsx / .csv, any format) · optionally your existing workbook (.xlsx) to refresh it'],
  flow: ['GST <em>ledger flow</em>', 'Opening, output tax, input tax, payments, refunds, reversals and closing, as per books and as per the portal, with the reasons for each difference worked out from the entries.', 'Portal Electronic Credit Ledger, Cash Ledger and Liability Register (CSV / Excel) · GST ledgers from Tally, Zoho or any software, one by one or as a full GL dump (.xlsx / .csv) · optionally GSTR-3B PDFs for the 3B checks'],
};
/** the mode the reconciliation code works in (the flow tab has no matching review) */
const mode = (): Mode => (tab === 'itc' ? 'itc' : 'annual');
function setTab(t: Tab) {
  tab = t; try { localStorage.setItem('jjmode', t); } catch { /* ignore */ }
  document.querySelectorAll('button.tab').forEach((b: any) => { b.classList.toggle('on', b.dataset.tab === t); b.setAttribute('aria-selected', String(b.dataset.tab === t)); });
  let n = 0;
  document.querySelectorAll('main > section.card').forEach((c: any) => {
    const on = (c.dataset.tabs || '').split(/\s+/).includes(t);
    c.classList.toggle('offtab', !on);
    if (!on) return;
    const num = c.querySelector('h2 .num');
    if (c.dataset.sub) { if (num) num.textContent = `${n}${c.dataset.sub}`; } else if (num) num.textContent = String(++n);
  });
  $('heroT').innerHTML = HERO[t][0]; $('heroP').textContent = HERO[t][1]; $('dropHint').textContent = HERO[t][2];
  $('build').textContent = t === 'itc' ? 'Build ITC reconciliation' : 'Build workbook';
  $('result').innerHTML = '';
  render();
}
document.querySelectorAll('button.tab').forEach((b: any) => (b.onclick = () => setTab(b.dataset.tab)));

/** the client's GSTIN: as typed, else the most frequent one in the files */
function clientGstin(): string {
  const typed = String($('gstin').value || '').trim().toUpperCase();
  if (typed && $('gstin').dataset.touched) return typed;
  const c: Record<string, number> = {};
  for (const f of files) if (f.gstin) c[f.gstin] = (c[f.gstin] || 0) + 1;
  return Object.entries(c).sort((a, b) => b[1] - a[1])[0]?.[0] || typed;
}
/** files of another GSTIN are left out of every tab (single GSTIN for now) */
const otherGstin = () => { const g = clientGstin(); return g ? files.filter(f => f.gstin && f.gstin.toUpperCase() !== g) : []; };
const activeFiles = () => { const o = new Set(otherGstin()); return files.filter(f => !o.has(f)); };

async function parsePdf(b: Uint8Array, name: string) {
  const doc = await pdfjs.getDocument({ data: b }).promise;
  const pages: any[] = [];
  for (let i = 1; i <= doc.numPages; i++) pages.push(await (await doc.getPage(i)).getTextContent());
  return parse3b(linesFromItems(pages), name);
}

const esc = (s: any) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' } as any)[c]);
const fmtP = (p?: string) => p && p.length === 6 ? `${p.slice(0, 2)}/${p.slice(2)}` : (p || '');

function render() {
  const order = ['Workbook', 'Register', 'Portal ledger', 'Books ledger', 'GSTR-9 (portal)', 'GSTR-3B', 'GSTR-1', 'GSTR-2B', 'Unknown'];
  const sorted = [...files].sort((a, b) => order.indexOf(a.kind) - order.indexOf(b.kind) || (a.period || '').slice(2).localeCompare((b.period || '').slice(2)) || (a.period || '').localeCompare(b.period || ''));
  const other = new Set(otherGstin());
  $('files').innerHTML = files.length ? `<div class="tablewrap"><table><tr><th>File</th><th>Type</th><th>GSTIN</th><th>Period</th><th>FY</th><th>Note</th><th></th></tr>` +
    sorted.map(f => `<tr><td>${esc(f.name)}</td><td><span class="tag">${f.kind}</span></td><td>${esc(f.gstin)}${other.has(f) ? ' <span class="warn-error">other GSTIN: ignored</span>' : ''}</td><td>${fmtP(f.period)}</td><td>${esc(f.fy)}</td><td class="muted">${esc(f.note)}</td><td><button class="link" data-i="${files.indexOf(f)}">remove</button></td></tr>`).join('') + '</table></div>' : '';
  $('files').querySelectorAll('button[data-i]').forEach((b: any) => b.onclick = () => { files.splice(+b.dataset.i, 1); render(); });
  // FY options: from 3B / 2B / portal; default = most frequent 3B FY
  const counts: Record<string, number> = {};
  for (const f of files) if (f.fy && (f.kind === 'GSTR-3B' || f.kind === 'GSTR-9 (portal)')) counts[f.fy] = (counts[f.fy] || 0) + (f.kind === 'GSTR-9 (portal)' ? 100 : 1);
  if (!Object.keys(counts).length) for (const f of files) if (f.fy) counts[f.fy] = (counts[f.fy] || 0) + 1;
  const best = Object.entries(counts).sort((a, b) => b[1] - a[1])[0]?.[0];
  const sel = $('fy') as HTMLSelectElement, cur = sel.value;
  const now = new Date(); const curFy = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  const opts = Array.from({ length: 8 }, (_, i) => `${curFy - i}-${String(curFy - i + 1).slice(2)}`);
  sel.innerHTML = opts.map(o => `<option>${o}</option>`).join('');
  sel.value = cur && sel.dataset.touched ? cur : (best || fyShift(opts[0], -1));
  const g = clientGstin(); if (g && !$('gstin').dataset.touched) $('gstin').value = g;
  const l = files.find(f => f.threeB?.legalName)?.threeB?.legalName; if (l && !$('legal').dataset.touched) $('legal').value = l;
  $('build').disabled = !activeFiles().some(f => f.kind !== 'Unknown' && f.kind !== 'Portal ledger' && f.kind !== 'Books ledger');
  renderCoverage();
  renderRegisters();
  flow.refresh();
}

// ---------- month-wise coverage grid
const MON = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'];
const periodsOf = (fy: string) => [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3].map(m => String(m).padStart(2, '0') + (m >= 4 ? fy.slice(0, 4) : String(+fy.slice(0, 4) + 1)));
/** periods a file covers (quarterly GSTR-1 covers the 3 months of its quarter) */
function covered(f: InFile): string[] {
  if (!f.period) return [];
  if (f.kind === 'GSTR-1' && f.json?.filing_typ === 'Q') {
    const m = +f.period.slice(0, 2), y = f.period.slice(2);
    return [m - 2, m - 1, m].map(x => String(x).padStart(2, '0') + y);
  }
  return [f.period];
}
function renderCoverage() {
  const el = $('coverage');
  const files = activeFiles(), oth = otherGstin();
  if (!files.length && !oth.length) { el.innerHTML = ''; return; }
  const othHtml = oth.length ? `<div class="gaps"><span class="warn-error">${oth.length} file(s) of another GSTIN are ignored in every tab:</span> ${oth.map(f => `${esc(f.name)} (${esc(f.gstin)})`).join(', ')}. The client GSTIN is ${esc(clientGstin())}: change it under Confirm client &amp; year if that is wrong.</div>` : '';
  if (tab === 'flow') { el.innerHTML = flow.coverageHtml() + othHtml; return; }
  if (tab === 'itc') {
    const fy = $('fy').value, ps = periodsOf(fy);
    const g2 = new Set(files.filter(f => f.kind === 'GSTR-2B').flatMap(covered));
    const loaded = ps.filter(p => g2.has(p));
    const first = ps.indexOf(loaded[0]), last = ps.indexOf(loaded[loaded.length - 1]);
    const gaps = loaded.length ? ps.slice(first, last + 1).filter(p => !g2.has(p)) : [];
    const regs = files.filter(f => f.kind === 'Register').length;
    el.innerHTML = `<table class="cov"><tr><th class="cvl">Return</th>${MON.map(m => `<th class="cvh">${m}</th>`).join('')}</tr>
      <tr><th class="cvl">GSTR-2B<span>${fy}</span></th>${ps.map((p, i) => g2.has(p) ? `<td class="cv ok">✓</td>` : (i > first && i < last ? `<td class="cv miss">✕</td>` : `<td class="cv na">·</td>`)).join('')}</tr></table>
      <div class="gaps">${!loaded.length ? '<span class="warn-error">Add the GSTR-2B JSON for the month(s) you are reconciling.</span>' : gaps.length ? `<span class="warn-error">Gap in 2B months:</span> ${gaps.map(p => MON[ps.indexOf(p)]).join(', ')}` : `<span class="ok-text">2B loaded for ${loaded.map(p => MON[ps.indexOf(p)]).join(', ')}.</span>`}
      ${regs ? '' : ' · <span class="warn-error">Add the purchase register.</span>'}</div>${othHtml}`;
    return;
  }
  const fy = $('fy').value, ny = fyShift(fy, 1), py = fyShift(fy, -1);
  const have = (kind: string) => { const set = new Set<string>(); files.filter(f => f.kind === kind).forEach(f => covered(f).forEach(p => set.add(p))); return set; };
  const g1 = have('GSTR-1'), g3 = have('GSTR-3B'), g2 = have('GSTR-2B');
  type Row = { label: string; sub: string; set: Set<string>; fy: string; window: number; optional?: boolean };
  const rows: Row[] = [
    { label: 'GSTR-3B', sub: fy, set: g3, fy, window: 12 },
    { label: 'GSTR-1', sub: fy, set: g1, fy, window: 12 },
    { label: 'GSTR-2B', sub: fy, set: g2, fy, window: 12 },
    { label: 'GSTR-1', sub: `${ny} · Apr–Nov`, set: g1, fy: ny, window: 8 },
    { label: 'GSTR-2B', sub: `${ny} · Apr–Nov`, set: g2, fy: ny, window: 8 },
    { label: 'GSTR-1', sub: `${py} · optional`, set: g1, fy: py, window: 12, optional: true },
  ];
  const gaps: string[] = [];
  const body = rows.map(r => {
    const ps = periodsOf(r.fy);
    const missing: string[] = [];
    const cells = ps.map((p, i) => {
      if (i >= r.window) return `<td class="cv na" title="Not needed">·</td>`;
      if (r.set.has(p)) return `<td class="cv ok" title="${MON[i]} ${p.slice(2)} loaded">✓</td>`;
      missing.push(MON[i]);
      return r.optional ? `<td class="cv opt" title="Optional - needed only to value amendments">–</td>` : `<td class="cv miss" title="${MON[i]} ${p.slice(2)} missing">✕</td>`;
    }).join('');
    const got = ps.slice(0, r.window).filter(p => r.set.has(p)).length;
    if (missing.length && !r.optional) gaps.push(`<b>${r.label} ${r.sub}</b>: ${missing.join(', ')}`);
    return `<tr><th class="cvl">${r.label}<span>${r.sub}</span></th>${cells}<td class="cvc ${got === r.window ? 'full' : r.optional ? '' : 'part'}">${got}/${r.window}</td></tr>`;
  }).join('');
  const p9 = files.some(f => f.kind === 'GSTR-9 (portal)');
  const wbk = files.some(f => f.kind === 'Workbook');
  const outside = files.filter(f => f.fy && ![fy, ny, py].includes(f.fy) && f.kind !== 'GSTR-9 (portal)');
  el.innerHTML = `<table class="cov"><tr><th class="cvl">Return</th>${MON.map(m => `<th class="cvh">${m}</th>`).join('')}<th class="cvh">Loaded</th></tr>${body}
    <tr><th class="cvl">GSTR-9<span>portal pre-filled JSON</span></th><td colspan="12" class="cv ${p9 ? 'ok' : 'miss'} wide">${p9 ? '✓ loaded' : '✕ not loaded — download from the GSTR-9 tile (Prepare Offline › Generate JSON file to download)'}</td><td class="cvc ${p9 ? 'full' : 'part'}">${p9 ? 1 : 0}/1</td></tr></table>
    <div class="gaps">${gaps.length ? `<span class="warn-error">Missing:</span> ${gaps.join(' · ')}` : '<span class="ok-text">All required months loaded.</span>'}
      ${files.some(f => f.kind === 'GSTR-1' && f.json?.filing_typ === 'Q') ? '<br/><span class="muted">Quarterly GSTR-1 counts for all three months of its quarter. For quarterly (QRMP) 3B filers, only quarter-end months are expected.</span>' : ''}
      ${wbk ? '<br/><span class="muted">Existing workbook loaded: sheets for return types you don\'t add here are left as they are.</span>' : ''}
      ${outside.length ? `<br/><span class="warn-warn">${outside.length} file(s) outside ${py} to ${ny} will be ignored: ${outside.map(f => esc(f.name)).join(', ')}</span>` : ''}</div>${othHtml}`;
}
['fy', 'gstin', 'legal'].forEach(id => $(id).addEventListener('change', () => { $(id).dataset.touched = '1'; if (id === 'gstin') { render(); return; } renderCoverage(); renderRegisters(); flow.refresh(); }));

async function addFiles(list: FileList | File[]) {
  $('status').textContent = 'Reading files…';
  let dup = 0;
  for (const f of Array.from(list)) {
    try {
      for (const x of await classify(f.name, new Uint8Array(await f.arrayBuffer()), parsePdf)) {
        // file names don't matter (the portal reuses names); identical content is skipped
        if (x.hash && files.some(y => y.hash === x.hash)) { dup++; continue; }
        if (files.some(y => y.name === x.name)) x.name = `${x.name} (${x.kind}${x.period ? ' ' + fmtP(x.period) : ''})`;
        files.push(x);
      }
    } catch (e: any) { files.push({ name: f.name, kind: 'Unknown', note: e.message }); }
    render();
  }
  if (dup) { $('status').textContent = `${dup} duplicate file(s) skipped`; return; }
  $('status').textContent = '';
}
const drop = $('drop');
drop.onclick = () => $('pick').click();
$('pick').onchange = (e: any) => { addFiles(e.target.files); e.target.value = ''; };
drop.ondragover = (e: DragEvent) => { e.preventDefault(); drop.classList.add('over'); };
drop.ondragleave = () => drop.classList.remove('over');
drop.ondrop = (e: DragEvent) => { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer?.files) addFiles(e.dataTransfer.files); };

$('build').onclick = async () => {
  $('build').disabled = true; $('status').textContent = 'Building…'; $('result').innerHTML = '';
  try {
    const fy = $('fy').value;
    if (tab === 'itc') { await buildItcNow(fy); $('build').disabled = false; return; }
    const tpl = new Uint8Array(await (await fetch(templateUrl)).arrayBuffer());
    const res = await buildWorkbook(tpl, activeFiles().filter(f => f.kind !== 'Register' && f.kind !== 'Portal ledger' && f.kind !== 'Books ledger'), fy, { gstin: $('gstin').value.trim(), legalName: $('legal').value.trim(), ...registerRows(fy), ...review.extra() });
    const name = `GSTR9_${($('gstin').value || 'client').trim()}_${fy}.xlsx`;
    const url = URL.createObjectURL(new Blob([res.bytes as any], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a'); a.href = url; a.download = name; a.click();
    const w = [...res.warns].sort((a, b) => ['error', 'warn', 'info'].indexOf(a.level) - ['error', 'warn', 'info'].indexOf(b.level));
    $('result').innerHTML = `<p><b>Downloaded ${esc(name)}</b> · <a href="${url}" download="${esc(name)}">download again</a>.
      Open it in Excel and review the Checks sheet.</p>
      ${res.recon?.purchases || res.recon?.sales ? `<p>Reconciliation: ${res.recon.purchases ? `<b>purchases vs 2B</b> ${res.recon.purchases.matched} matched, ${res.recon.purchases.booksOnly} in books only, ${res.recon.purchases.twoBOnly} in 2B only, ${res.recon.purchases.valueDiff} value differences (Recon_Purchases, Recon_Suppliers)` : ''}${res.recon.purchases && res.recon.sales ? ' · ' : ''}${res.recon.sales ? `<b>sales vs GSTR-1</b> ${res.recon.sales.matched} matched, ${res.recon.sales.booksOnly} in books only, ${res.recon.sales.g1Only} in GSTR-1 only, ${res.recon.sales.valueDiff} value differences (Recon_Sales, Recon_Sales_Monthly)` : ''}.</p>` : ''}
      <div class="grid2"><div><table><tr><th>Loaded</th><th class="r">Count</th></tr>${Object.entries(res.counts).map(([k, v]) => `<tr><td>${k}</td><td class="r">${v}</td></tr>`).join('')}</table></div>
      <div><table><tr><th>Key totals (taxable for GSTR-1, tax for 2B)</th><th class="r">Amount</th></tr>${Object.entries(res.summary).map(([k, v]) => `<tr><td>${k}</td><td class="r">${v.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td></tr>`).join('')}</table></div></div>
      ${w.length ? `<h3 style="font-size:14px;margin:16px 0 0">Messages (${w.length})</h3><ul class="warns">${w.map(x => `<li class="warn-${x.level}">${x.file ? `<b>${esc(x.file)}:</b> ` : ''}${esc(x.msg)}</li>`).join('')}</ul>` : '<p class="muted">No warnings.</p>'}`;
    $('status').textContent = '';
  } catch (e: any) {
    $('status').innerHTML = `<span class="warn-error">${esc(e.message)}</span>`;
  }
  $('build').disabled = false;
};


// ---------- step 4: workbook -> GSTR-9 upload JSON
const drop2 = $('drop2');
drop2.onclick = () => $('pick2').click();
drop2.ondragover = (e: DragEvent) => { e.preventDefault(); drop2.classList.add('over'); };
drop2.ondragleave = () => drop2.classList.remove('over');
drop2.ondrop = (e: DragEvent) => { e.preventDefault(); drop2.classList.remove('over'); const f = e.dataTransfer?.files?.[0]; if (f) gen(f); };
$('pick2').onchange = (e: any) => { const f = e.target.files?.[0]; if (f) gen(f); e.target.value = ''; };
const fmtN = (v: number) => v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

async function gen(file: File) {
  const el = $('gen'); el.innerHTML = '<span class="muted">Reading workbook…</span>';
  try {
    const tpl = new Uint8Array(await (await fetch(templateUrl)).arrayBuffer());
    const r = await generate9(new Uint8Array(await file.arrayBuffer()), tpl);
    const errors = r.msgs.filter(m => m.level === 'error'), others = r.msgs.filter(m => m.level !== 'error');
    const name = `${r.gstin || 'GSTIN'}_${r.fy}_GSTR9_upload.json`;
    const url = URL.createObjectURL(new Blob([r.text], { type: 'application/json' }));
    const heads = ['txval', 'camt', 'samt', 'iamt', 'csamt'];
    const rows = r.lines.map(l => `<tr><td>${esc(l.table.replace('table', 'Table '))}</td><td>${esc(l.field)}</td>${heads.map(h => `<td class="r">${h in l.values ? fmtN(l.values[h]) : ''}</td>`).join('')}<td class="r">${'txpyble' in l.values ? fmtN(l.values.txpyble) : ''}</td></tr>`).join('');
    const ig = r.integrity;
    const igErr = ig?.items.filter(i => i.level === 'error') ?? [];
    const igHtml = ig ? `<h3 class="eyebrow" style="margin:6px 0 0">Edit check${ig.builtAt ? ` · built by the Loader on ${esc(ig.builtAt)}` : ''}</h3>
      <ul class="warns">${ig.items.map(i => `<li class="warn-${i.level === 'info' ? 'ok' : i.level}"><b>${esc(i.area)}:</b> ${esc(i.msg)}</li>`).join('')}</ul>
      ${ig.adjustments.length ? `<details style="margin-top:8px" open><summary class="eyebrow" style="cursor:pointer">Your adjustments and Basis changes (${ig.adjustments.length})</summary>
        <table style="margin-top:6px"><tr><th>Row</th><th>Change</th><th>Detail</th><th>Remark</th></tr>${ig.adjustments.map(a => `<tr><td>${esc(a.where)}</td><td>${esc(a.what)}</td><td>${esc(a.values)}</td><td>${esc(a.remark)}</td></tr>`).join('')}</table></details>` : '<p class="muted">No adjustments or Basis changes.</p>'}` : '';
    const dlBtn = `<a id="dl9" href="${url}" download="${esc(name)}"><button>Download ${esc(name)}</button></a>`;
    el.innerHTML = `
      <p><b>${esc(r.gstin)} · FY ${esc(r.fy)}</b> · ${r.lines.length} entries ready.</p>
      ${igHtml}
      ${errors.length ? `<ul class="warns">${errors.map(m => `<li class="warn-error">${esc(m.msg)}</li>`).join('')}</ul><p class="warn-error">Fix the errors above and drop the saved workbook again. The JSON is not offered while errors remain.</p>`
        : igErr.length ? `<p class="warn-error" style="margin-top:12px">The workbook differs from what the Loader produced (see the edit check). Review those changes before uploading.</p>
            <label style="display:flex;gap:8px;align-items:center;text-transform:none;letter-spacing:0;font-weight:300;font-size:14px;color:var(--strong)"><input type="checkbox" id="ack" style="min-width:0"/> I have reviewed these changes and want to generate the JSON anyway</label>
            <p id="dlwrap" style="display:none">${dlBtn}</p>`
        : `<p>${dlBtn}</p>`}
      ${others.length ? `<h3 class="eyebrow" style="margin:16px 0 0">Review before uploading (${others.length})</h3><ul class="warns">${others.map(m => `<li class="warn-${m.level}">${esc(m.msg)}</li>`).join('')}</ul>` : ''}
      <details style="margin-top:14px"><summary class="eyebrow" style="cursor:pointer">What goes into the JSON</summary>
        <table style="margin-top:8px"><tr><th>Table</th><th>Field</th><th class="r">Taxable</th><th class="r">Central</th><th class="r">State/UT</th><th class="r">Integrated</th><th class="r">Cess</th><th class="r">Payable</th></tr>${rows}</table>
        <p class="muted">6A, 8A and the paid columns of Table 9 are not uploaded: the portal fills them. Rows that are zero both here and on the portal are left out, as the GSTN offline tool does.</p></details>
      <h3 class="eyebrow" style="margin:18px 0 0">Uploading on the portal</h3>
      <ol class="steps"><li>Returns Dashboard › select the financial year › Annual Return (GSTR-9) › <b>Prepare Offline</b>.</li>
        <li>Upload tab › choose this JSON › Upload. Wait for the status to show <b>Processed</b> (or download the error report if it shows Processed with Error).</li>
        <li>Back on the GSTR-9 tile, <b>Prepare Online</b> › Preview draft GSTR-9 (PDF) and compare it with the workbook.</li>
        <li>Uploading only saves a draft. Compute liabilities, pay any late fee, and file with DSC / EVC when the preview matches.</li></ol>`;
    const ack = document.getElementById('ack') as HTMLInputElement | null;
    if (ack) ack.onchange = () => { (document.getElementById('dlwrap') as HTMLElement).style.display = ack.checked ? 'block' : 'none'; };
  } catch (e: any) { el.innerHTML = `<span class="warn-error">${esc(e.message)}</span>`; }
}


// ---------- registers: column mapping
interface RegState { sheets: RegSheet[]; sheet: number; header: number; kind: RegKind; map: Mapping; d: Record<string, string> }
const regState = new Map<InFile, RegState>();
const store = { get(k: string) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch { return null; } }, set(k: string, v: any) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* storage unavailable */ } } };
const sigOf = (st: RegState) => 'jjmap:' + st.kind + ':' + (st.sheets[st.sheet].rows[st.header] || []).map(v => String(v ?? '').trim().toLowerCase()).join('|');

async function initReg(f: InFile) {
  const sheets = await readTable(f.name, f.bytes!);
  const { sheet, header, kind } = pickRegister(f.name, sheets);
  const st: RegState = { sheets, sheet, header, kind, map: {}, d: {} };
  st.map = store.get(sigOf(st)) || guessMappingFull(sheets[sheet].rows, header, fieldsOf(kind));
  regState.set(f, st);
}
function remap(st: RegState) { st.map = store.get(sigOf(st)) || guessMappingFull(st.sheets[st.sheet].rows, st.header, fieldsOf(st.kind)); }
const defaults = (st: RegState) => ({ noGstinZeroTax: st.d.noGstinZeroTax || 'EXEMPT', itcType: st.d.itcType || 'Inputs', status: st.d.status || 'Availed',
  claimFy: st.d.claimFy || $('fy').value, purNoGstin: st.d.purNoGstin || (isMapped(st.map.gstin) ? 'RCM-UNREG' : 'B2B') });

function renderRegisters() {
  const regs = files.filter(f => f.kind === 'Register');
  if (tab === 'itc') for (const f of regs) { const st = regState.get(f); if (st && st.kind !== 'purchases') { st.kind = 'purchases'; st.header = detectHeader(st.sheets[st.sheet].rows, PURCHASE_FIELDS); remap(st); } }
  $('regcard').style.display = regs.length ? '' : 'none';
  const box = $('regs');
  box.innerHTML = '';
  for (const f of regs) {
    const st = regState.get(f);
    const div = document.createElement('div'); div.className = 'reg'; box.appendChild(div);
    if (!st) { div.innerHTML = `<h3>${esc(f.name)}</h3><span class="muted">Reading…</span>`; initReg(f).then(renderRegisters).catch(e => { f.kind = 'Unknown'; f.note = e.message; render(); }); continue; }
    const sh = st.sheets[st.sheet], hdr = sh.rows[st.header] || [];
    const isDateCol = (i: number) => /date|dt\b/i.test(String(hdr[i] ?? ''));
    const sample = (i: number) => { for (const r of sh.rows.slice(st.header + 1, st.header + 30)) { const v = r?.[i]; if (v !== null && v !== undefined && v !== '') return (typeof v === 'number' && isDateCol(i) && v > 20000 && v < 80000 ? new Date(Date.UTC(1899, 11, 30) + v * 86400000).toISOString().slice(0, 10).split('-').reverse().join('-') : String(v)).slice(0, 18); } return ''; };
    const colName = (i: number) => esc(hdr[i] ?? `Column ${i + 1}`);
    const colOpts = (sel: number | undefined, first = '— not in register —', skip: number[] = []) => `<option value="-1">${first}</option>` + hdr.map((h, i) => skip.includes(i) ? '' : `<option value="${i}" ${sel === i ? 'selected' : ''}>${esc(h ?? `Column ${i + 1}`)}${sample(i) ? ' · e.g. ' + esc(sample(i)) : ''}</option>`).join('');
    /** amount fields can take several columns (Tally columnar: one column per ledger), summed per row */
    const fieldCtl = (x: FieldDef) => {
      const cols = colsOf(st.map[x.id]);
      if (!x.multi || cols.length < 2) return `<select data-f="${x.id}">${colOpts(cols[0])}</select>` + (x.multi && cols.length ? `<select data-add="${x.id}" class="addcol">${colOpts(undefined, '+ add another column (summed)', cols)}</select>` : '');
      return `<div class="chips">${cols.map(j => `<span class="chip">${colName(j)}<button class="link" data-rm="${x.id}:${j}" title="remove">×</button></span>`).join('')}</div>
        <select data-add="${x.id}" class="addcol">${colOpts(undefined, `+ add column (sum of ${cols.length})`, cols)}</select>`;
    };
    const fields = fieldsOf(st.kind);
    const dsel = (key: string, label: string, opts: string[], cur: string) => `<label>${label}<select data-d="${key}">${opts.map(o => `<option ${o === cur ? 'selected' : ''}>${o}</option>`).join('')}</select></label>`;
    const dd = defaults(st);
    const missing = fields.filter(x => x.required && !isMapped(st.map[x.id]));
    const columnar = looksColumnar(hdr) || fields.some(x => x.multi && colsOf(st.map[x.id]).length > 1);
    let prevHtml = '';
    if (!missing.length) {
      const n = normalise(st.kind, sh.rows, st.header, st.map, dd, $('fy').value);
      const unused = columnar ? unusedAmountCols(sh.rows, st.header, st.map) : [];
      const cols = st.kind === 'sales' ? ['Date', 'No.', 'Customer', 'GSTIN', 'POS', 'Supply type', 'Doc', 'HSN', 'UQC', 'Qty', 'Rate', 'Taxable', 'CGST', 'SGST', 'IGST', 'Cess']
        : ['Date', 'No.', 'Supplier', 'GSTIN', 'Supply type', 'ITC type', 'ITC status', 'Claimed FY', 'HSN', 'UQC', 'Qty', 'Rate', 'Taxable', 'CGST', 'SGST', 'IGST', 'Cess'];
      const fmt = (v: any, i: number) => i === 0 && typeof v === 'number' ? new Date(Date.UTC(1899, 11, 30) + v * 86400000).toISOString().slice(0, 10).split('-').reverse().join('-') : (v ?? '');
      prevHtml = `<p style="margin:6px 0">${n.stats.lines} lines → ${n.stats.docs} documents${n.stats.credit ? ` · ${n.stats.credit} credit-note lines` : ''}${n.stats.skipped ? ` · ${n.stats.skipped} rows skipped (no number / no amount / totals)` : ''}
        ${n.stats.grossOff ? ` · <span class="warn-error">${n.stats.grossOff} documents where taxable + tax (+ round off) ≠ gross total, e.g. ${n.stats.grossOffEg.map(esc).join(', ')}: a ledger column is probably missing or mapped twice</span>` : ''}
        ${n.stats.badDates ? ` · <span class="warn-error">${n.stats.badDates} unreadable dates</span>` : ''}${n.stats.outsideFy ? ` · <span class="warn-warn">${n.stats.outsideFy} lines dated outside FY ${esc($('fy').value)}</span>` : ''}</p>
        <p class="muted" style="margin:4px 0">Totals: ${(st.kind === 'sales' ? [['Taxable', 11], ['CGST', 12], ['SGST', 13], ['IGST', 14], ['Cess', 15]] : [['Taxable', 12], ['CGST', 13], ['SGST', 14], ['IGST', 15], ['Cess', 16]] as [string, number][]).map(([l, j]) => `${l} <b>${fmtN(n.rows.reduce((a, r) => a + (+r[j as number] || 0), 0))}</b>`).join(' · ')}</p>
        ${unused.length ? `<p class="warn-warn" style="margin:6px 0">Amount columns not used: ${unused.map(colName).join(', ')}. If any is a purchase / expense or tax ledger, add it to Taxable value or the tax head.</p>` : ''}
        <div class="prev"><table><tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr>${n.rows.slice(0, 5).map(r => `<tr>${cols.map((_, i) => `<td>${esc(fmt(r[i], i))}</td>`).join('')}</tr>`).join('')}</table></div>`;
    } else prevHtml = `<p class="warn-error">Map the required fields: ${missing.map(m => esc(m.label)).join(', ')}</p>`;
    div.innerHTML = `<h3>${esc(f.name)}</h3>
      <div class="row">
        <label ${tab === 'itc' ? 'style="display:none"' : ''}>Register type<select data-k="kind"><option value="sales" ${st.kind === 'sales' ? 'selected' : ''}>Sales register</option><option value="purchases" ${st.kind === 'purchases' ? 'selected' : ''}>Purchase register</option></select></label>
        ${st.sheets.length > 1 ? `<label>Sheet<select data-k="sheet">${st.sheets.map((x, i) => `<option value="${i}" ${i === st.sheet ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select></label>` : ''}
        <label>Header row<input data-k="header" type="number" min="1" value="${st.header + 1}" style="min-width:0;width:90px"/></label>
      </div>
      ${columnar ? `<p class="callout">This looks like a <b>Tally columnar register</b> (one column per ledger). Taxable value is the sum of the purchase / sales / expense ledger columns picked below, and each tax head is the sum of its tax-ledger columns. Check that every ledger is included and that Gross Total, Round Off and TDS are not.</p>` : ''}
      <div class="mapgrid">${fields.map(x => `<label><span>${esc(x.label)}${x.required ? ' <span class="req">*</span>' : ''}</span>${fieldCtl(x)}</label>`).join('')}</div>
      <div class="row">${st.kind === 'sales'
        ? dsel('noGstinZeroTax', 'No-GSTIN rows with no tax are', ['EXEMPT', 'NIL', 'NON-GST', 'EXP-WOP'], dd.noGstinZeroTax)
        : dsel('itcType', 'Default ITC type', ['Inputs', 'Input Services', 'Capital Goods'], dd.itcType) + dsel('status', 'Default ITC status', ['Availed', 'Not availed'], dd.status)
          + dsel('claimFy', 'ITC claimed in (default)', [$('fy').value, fyShift($('fy').value, 1)], dd.claimFy) + dsel('purNoGstin', 'Rows without supplier GSTIN are', ['RCM-UNREG', 'IMPS', 'IMPG', 'B2B'], dd.purNoGstin)}</div>
      ${st.kind === 'purchases' && !isMapped(st.map.gstin) ? '<p class="warn-warn" style="margin:8px 0 0">No supplier GSTIN column: the GSTIN will be taken from GSTR-2B by supplier name (or invoice no. + amount) when reconciling. Map the GSTIN column if the register has one.</p>' : ''}
      <p class="muted" style="margin:8px 0 0">Supply type is taken from the mapped column when present; otherwise rows with a valid GSTIN are B2B and the rest follow the defaults. You can still correct individual lines in the workbook's Books sheets.</p>
      ${prevHtml}`;
    const save = () => { store.set(sigOf(st), st.map); renderRegisters(); };
    div.querySelectorAll('select[data-f]').forEach((el: any) => el.onchange = () => { st.map[el.dataset.f] = +el.value; save(); });
    div.querySelectorAll('select[data-add]').forEach((el: any) => el.onchange = () => { const id = el.dataset.add, j = +el.value; if (j >= 0) { st.map[id] = [...new Set([...colsOf(st.map[id]), j])]; save(); } });
    div.querySelectorAll('button[data-rm]').forEach((el: any) => el.onclick = () => { const [id, j] = el.dataset.rm.split(':'); const left = colsOf(st.map[id]).filter(c => c !== +j); st.map[id] = left.length === 1 ? left[0] : left.length ? left : -1; save(); });
    div.querySelectorAll('select[data-d]').forEach((el: any) => el.onchange = () => { st.d[el.dataset.d] = el.value; renderRegisters(); });
    (div.querySelector('select[data-k="kind"]') as any).onchange = (e: any) => { st.kind = e.target.value; st.header = detectHeader(st.sheets[st.sheet].rows, fieldsOf(st.kind)); remap(st); renderRegisters(); };
    const shSel = div.querySelector('select[data-k="sheet"]') as any;
    if (shSel) shSel.onchange = (e: any) => { st.sheet = +e.target.value; st.header = detectHeader(st.sheets[st.sheet].rows, fieldsOf(st.kind)); remap(st); renderRegisters(); };
    (div.querySelector('input[data-k="header"]') as any).onchange = (e: any) => { st.header = Math.max(0, +e.target.value - 1); remap(st); renderRegisters(); };
  }
  renderSuppliers();
}
/** normalised register rows per kind, for the build */
function registerRows(fy: string, lenient = false) {
  const out: { sales?: any[][]; purchases?: any[][] } = {};
  for (const f of files.filter(x => x.kind === 'Register')) {
    const st = regState.get(f); if (!st) continue;
    if (fieldsOf(st.kind).some(x => x.required && !isMapped(st.map[x.id]))) { if (lenient) continue; throw new Error(`${f.name}: map the required columns first`); }
    const n = normalise(st.kind, st.sheets[st.sheet].rows, st.header, st.map, defaults(st), fy);
    (out[st.kind] ||= []).push(...n.rows);
  }
  // supplier names the user confirmed against 2B → their GSTIN
  if (out.purchases && !lenient) applyAliases(out.purchases, confirmedAliases());
  return out;
}

// ---------- supplier names: books rows without GSTIN, matched to 2B suppliers by approximate name, confirmed by the user
const ALIAS_KEY = 'jjalias';
const aliasStore = (): Record<string, string> => store.get(ALIAS_KEY) || {};
const confirmedAliases = () => new Map(Object.entries(aliasStore()).filter(([, g]) => g));
const supPick = new Map<string, string>(); // name key → GSTIN chosen in the dropdown (not yet confirmed)
function renderSuppliers() {
  const card = $('supcard'), box = $('sups');
  const b2 = activeFiles().filter(f => f.kind === 'GSTR-2B');
  const pur = b2.length ? registerRows($('fy').value, true).purchases : undefined;
  queueMicrotask(() => review.refresh()); // after module init (review is declared further down)
  const sup2b = b2.length ? suppliersFrom2B(b2.map(f => f.json)) : new Map();
  const sug = pur ? supplierSuggestions(pur, sup2b) : [];
  const allSup = [...sup2b.values()].sort((a, b) => (a.name || a.gstin).localeCompare(b.name || b.gstin));
  const exact = sug.filter(x => x.exact), approx = sug.filter(x => !x.exact);
  card.style.display = sug.length ? '' : 'none';
  if (!sug.length) { box.innerHTML = ''; return; }
  const al = aliasStore();
  const pct = (x: number) => `${Math.round(x * 100)}%`;
  const rowsHtml = approx.map(x => {
    const conf = al[x.key];
    const cur = conf || supPick.get(x.key) || (x.cands[0]?.gstin || '');
    const strong = x.cands[0] && (x.cands[0].invHits > 0 || x.cands[0].score >= 0.9);
    const candSet = new Set(x.cands.map(c => c.gstin));
    // suggestions first, then every other 2B supplier so the user can pair names the system did not suggest
    const opts = (x.cands.length ? '' : `<option value="" ${cur ? '' : 'selected'}>— pick the supplier in 2B —</option>`)
      + (x.cands.length ? `<optgroup label="Suggested">${x.cands.map(c => `<option value="${c.gstin}" ${c.gstin === cur ? 'selected' : ''}>${esc(c.name || c.gstin)} · ${c.gstin} · ${pct(c.score)}${c.invHits ? ` · ${c.invHits} invoice no. match` : ''}</option>`).join('')}</optgroup>` : '')
      + `<optgroup label="All suppliers in 2B">${allSup.filter(c => !candSet.has(c.gstin)).map(c => `<option value="${c.gstin}" ${c.gstin === cur ? 'selected' : ''}>${esc(c.name || c.gstin)} · ${c.gstin}</option>`).join('')}</optgroup>`;
    const state = conf ? `<span class="tag ok-tag">Confirmed</span> <button class="link" data-undo="${esc(x.key)}">undo</button>`
      : conf === '' ? `<span class="tag no-tag">No match</span> <button class="link" data-undo="${esc(x.key)}">undo</button>`
      : `<button class="small" data-ok="${esc(x.key)}">Confirm</button> <button class="link" data-no="${esc(x.key)}">${x.cands.length ? 'Not this supplier' : 'Not in 2B'}</button>`;
    return `<tr class="${conf ? 'confirmed' : ''}"><td><b>${esc(x.name)}</b><div class="muted">${x.docs} invoice(s) · ITC ${fmtN(x.tax)}</div></td>
      <td><select data-pick="${esc(x.key)}" ${conf !== undefined ? 'disabled' : ''}>${opts}</select>${strong && conf === undefined ? '<div class="ok-text" style="font-size:12px">Strong match</div>' : !x.cands.length && conf === undefined ? '<div class="muted" style="font-size:12px">No similar name found: pick the supplier yourself if it is in 2B</div>' : ''}</td><td class="act">${state}</td></tr>`;
  }).join('');
  const pending = approx.filter(x => al[x.key] === undefined && x.cands.length);
  const unsuggested = approx.filter(x => !x.cands.length).length;
  const strongPending = pending.filter(x => x.cands[0].invHits > 0 || x.cands[0].score >= 0.9);
  box.innerHTML = `${exact.length ? `<p class="ok-text" style="margin:0 0 10px">${exact.length} supplier name(s) match a 2B supplier exactly and get its GSTIN automatically.</p>` : ''}
    ${approx.length ? `<div class="row" style="justify-content:space-between;margin-bottom:10px"><span class="muted">${approx.length} supplier name(s) without a GSTIN don't exactly match a name in 2B${unsuggested ? ` (${unsuggested} with no similar name)` : ''}. Pick the right supplier (suggestions first, or any supplier in 2B) and confirm. Confirmed names are remembered for next time.</span>
      ${strongPending.length ? `<button class="small" id="supAllStrong">Confirm ${strongPending.length} strong match(es)</button>` : ''}</div>
      <div class="tablewrap"><table class="sup"><tr><th>Name in books</th><th>Supplier in GSTR-2B</th><th></th></tr>${rowsHtml}</table></div>` : ''}`;
  const setAlias = (k: string, g: string | null) => { const a = aliasStore(); if (g === null) delete a[k]; else a[k] = g; store.set(ALIAS_KEY, a); renderSuppliers(); };
  const pickOf = (k: string) => supPick.get(k) ?? approx.find(x => x.key === k)?.cands[0]?.gstin ?? '';
  box.querySelectorAll('select[data-pick]').forEach((el: any) => (el.onchange = () => supPick.set(el.dataset.pick, el.value)));
  box.querySelectorAll('button[data-ok]').forEach((el: any) => (el.onclick = () => { const g = pickOf(el.dataset.ok); if (g) setAlias(el.dataset.ok, g); }));
  box.querySelectorAll('button[data-no]').forEach((el: any) => (el.onclick = () => setAlias(el.dataset.no, '')));
  box.querySelectorAll('button[data-undo]').forEach((el: any) => (el.onclick = () => setAlias(el.dataset.undo, null)));
  const all = box.querySelector('#supAllStrong') as any;
  if (all) all.onclick = () => { const a = aliasStore(); for (const x of strongPending) a[x.key] = pickOf(x.key); store.set(ALIAS_KEY, a); renderSuppliers(); };
}

// ---------- review matches (step 4): runs the reconciliation in the page; settings and manual decisions go into the build
const review = new Review($('rvcard'), $('rv'), () => {
  if (!files.length || tab === 'flow') return null;
  let regs: any;
  try { regs = registerRows($('fy').value); } catch (e: any) { regs = e.message; }
  return { files: activeFiles().filter(f => f.kind !== 'Portal ledger' && f.kind !== 'Books ledger'), fy: $('fy').value, mode: mode(), regs };
});



// ---------- monthly mode build
async function buildItcNow(fy: string) {
  try {
    const tpl = new Uint8Array(await (await fetch(itcTemplateUrl)).arrayBuffer());
    const regs = registerRows(fy);
    const res = await buildItc(tpl, activeFiles().filter(f => f.kind !== 'Register' && f.kind !== 'Portal ledger' && f.kind !== 'Books ledger'), fy, { gstin: $('gstin').value.trim(), legalName: $('legal').value.trim(), purchases: regs.purchases, ...review.extra() });
    const p = res.periods;
    const label = !p.length ? fy : p.length === 1 ? p[0] : `${p[0]}-${p[p.length - 1]}`;
    const name = `ITC_Recon_${($('gstin').value || 'client').trim()}_${label}.xlsx`;
    const url = URL.createObjectURL(new Blob([res.bytes as any], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const errs = res.warns.filter(w => w.level === 'error');
    if (!errs.length) { const a = document.createElement('a'); a.href = url; a.download = name; a.click(); }
    const P = res.position, heads = ['Central', 'State/UT', 'Integrated', 'Cess', 'Total'];
    const line = (lab: string, v?: number[], strong = false) => v ? `<tr${strong ? ' style="font-weight:500"' : ''}><td>${lab}</td>${[...v, v.reduce((a, b) => a + b, 0)].map(x => `<td class="r">${fmtN(x)}</td>`).join('')}</tr>` : '';
    const w = [...res.warns].sort((a, b) => ['error', 'warn', 'info'].indexOf(a.level) - ['error', 'warn', 'info'].indexOf(b.level));
    $('result').innerHTML = (errs.length ? '' : `<p><b>Downloaded ${esc(name)}</b> · <a href="${url}" download="${esc(name)}">download again</a>.</p>`)
      + (res.recon ? `<p>${res.recon.matched} matched · ${res.recon.booksOnly} in books only · ${res.recon.twoBOnly} in 2B only · ${res.recon.valueDiff} value differences</p>
        <table><tr><th>ITC position</th>${heads.map(h => `<th class="r">${h}</th>`).join('')}</tr>
        ${line('Available as per GSTR-2B', P.available)}${line('As per books (compared invoices)', P.books)}${line('Claimable: matched invoices (lower of books and 2B)', P.claimable, true)}
        ${line('In books, not in 2B (do not claim)', P.booksOnly)}${line('In 2B, not in books', P.twoBOnly)}${line('Value differences (books - 2B)', P.valueDiff)}</table>` : '')
      + (w.length ? `<ul class="warns">${w.map(x => `<li class="warn-${x.level}">${x.file ? `<b>${esc(x.file)}:</b> ` : ''}${esc(x.msg)}</li>`).join('')}</ul>` : '')
      + `<p class="muted">To match invoices manually: type the same Match code on the rows in Recon_Purchases, save, and drop the workbook here again with the same files.</p>`;
    $('status').textContent = '';
  } catch (e: any) { $('status').innerHTML = `<span class="warn-error">${esc(e.message)}</span>`; }
}

// ---------- GST ledger flow tab
const flow = new FlowTab(() => ({ files: activeFiles(), fy: $('fy').value, gstin: clientGstin(), legal: String($('legal').value || '').trim(), active: tab === 'flow', register: () => registerRows($('fy').value, true).purchases }), () => render());

setTab(tab);

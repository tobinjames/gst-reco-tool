// GST ledger flow tab: portal ledgers, books GST accounts, balances, the three flow statements, reasons, cross-checks, workbook
import type { InFile } from '../build';
import { AccHead, AccType, ACC_TYPE_LABEL, BLOCK_LABEL, BUCKETS, Block, BookFile, HEADS, HEAD_LABEL, Head, INTERNAL, PORTAL_LABEL, PortalKind, UNMAPPED, Ym, bucketOf } from './types';
import { accountsOf, guessAccount, isGstName, parseBooks } from './books';
import { ymOfDate } from './portal';
import { AccMapEntry, BEntry, HEAD_KEYS, HeadK, Inputs, Item, LineOverride, Result, allowedBuckets, blockOfType, bucketLabel, monthsOf, run, ymLabel } from './engine';
import { AuditRow, flowWorkbook } from './workbook';
import { ReconEvent, ReconStatement, ledgerLabel, reconcile } from './recon';
import { FLOW_GUIDE } from './guide';
import { readTable } from '../register';

const esc = (s: any) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' } as any)[c]);
const fmt = (v: number) => (Math.abs(v) < 0.005 ? '–' : v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const store = { get(k: string) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch { return null; } }, set(k: string, v: any) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* storage unavailable */ } } };
const $ = (id: string) => document.getElementById(id)! as any;
const HK_TXT: Record<HeadK, string> = { IGST: 'IGST', CGST: 'CGST', SGST: 'SGST', CESS: 'Cess', ALL: 'Not split' };
const BLOCKS: Block[] = ['itc', 'liab', 'cash', 'fee'];
const ACC_TYPES = Object.keys(ACC_TYPE_LABEL) as AccType[];
const nowTxt = () => { const d = new Date(); return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`; };

interface AccState extends AccMapEntry { confirmed?: boolean }
interface State { ov: Record<string, LineOverride>; bal: Inputs['bal']; audit: AuditRow[]; from?: Ym; to?: Ym }
export interface Ctx { files: InFile[]; fy: string; gstin: string; legal: string; active: boolean }

export class FlowTab {
  private acc: Record<string, AccState> = {};
  private st: State = { ov: {}, bal: { books: {}, portal: {} }, audit: [] };
  private key = '';
  private view: 'books' | 'portal' | 'diff' = 'diff';
  private showAll = false;
  private editing = '';
  private open = new Set<string>();
  private res: Result | null = null;
  private inp: Inputs | null = null;
  private ledgerNames: Record<string, string> = {};
  constructor(private ctx: () => Ctx, private onFilesChanged: () => void) {
    $('flBuild').onclick = () => this.build();
  }

  private load(c: Ctx) {
    const k = `${c.gstin || '-'}:${c.fy}`;
    if (k === this.key) return;
    this.key = k;
    this.acc = store.get(`jjflow:acc:${c.gstin || '-'}`) || {};
    this.st = { ov: {}, bal: { books: {}, portal: {} }, audit: [], ...(store.get(`jjflow:${k}`) || {}) };
  }
  private save() { const c = this.ctx(); store.set(`jjflow:acc:${c.gstin || '-'}`, this.acc); store.set(`jjflow:${this.key}`, this.st); }
  private log(change: string, record: string, from: string, to: string, remark: string) { this.st.audit.push({ when: nowTxt(), change, record, from, to, remark }); }

  /** files of the tab */
  private parts(c: Ctx) {
    const portal = c.files.filter(f => f.kind === 'Portal ledger' && f.portal).map(f => ({ file: f.name, ledger: f.portal! }));
    const books = c.files.filter(f => f.kind === 'Books ledger' && f.books).map(f => f.books!);
    return { portal, books, threeB: c.files.filter(f => f.kind === 'GSTR-3B' && f.threeB).map(f => f.threeB!), regs: c.files.filter(f => f.kind === 'Register') };
  }

  /** summary for the shared Add files card */
  coverageHtml(): string {
    const c = this.ctx(); const { portal, books, threeB } = this.parts(c);
    const has = (k: PortalKind) => portal.filter(p => p.ledger.kind === k);
    const row = (label: string, list: { file: string; ledger: any }[], opt = false) => `<tr><th class="cvl">${label}</th><td class="cv ${list.length ? 'ok' : opt ? 'opt' : 'miss'} wide">${list.length ? '✓ ' + list.map(p => `${esc(p.ledger.from)} to ${esc(p.ledger.to)}`).join(' · ') : opt ? '– optional' : '✕ not loaded'}</td></tr>`;
    return `<table class="cov">${row('Electronic Credit Ledger', has('credit'))}${row('Electronic Cash Ledger', has('cash'))}${row('Electronic Liability Register', has('liability'))}
      <tr><th class="cvl">Books ledgers</th><td class="cv ${books.length ? 'ok' : 'miss'} wide">${books.length ? `✓ ${books.length} file(s), ${books.reduce((s, b) => s + b.lines.length, 0)} lines` : '✕ GST ledgers or a GL dump from your accounting software'}</td></tr>
      <tr><th class="cvl">GSTR-3B<span>optional, for 3B checks</span></th><td class="cv ${threeB.length ? 'ok' : 'opt'} wide">${threeB.length ? `✓ ${threeB.length} month(s)` : '– optional'}</td></tr></table>`;
  }

  refresh() {
    const c = this.ctx();
    if (!c.active) return;
    this.load(c);
    const { portal, books, threeB, regs } = this.parts(c);
    // accounts: GST-looking ones get a suggested type and head until the user confirms
    const accs = accountsOf(books);
    for (const a of accs) if (!this.acc[a.name] && a.gst) { const g = guessAccount(a.name); this.acc[a.name] = { type: isGstName(a.name) || books.some(b => b.format === 'ledger' && b.ledger === a.name) ? g.type : 'ignore', head: g.head }; }
    const months = monthsOf(c.fy);
    const from = this.st.from && months.includes(this.st.from) ? this.st.from : months[0], to = this.st.to && months.includes(this.st.to) ? this.st.to : months[11];
    const accMap: Inputs['accMap'] = {};
    for (const a of accs) { const s = this.acc[a.name]; if (s && s.type !== 'ignore') accMap[a.name] = { type: s.type, head: s.head }; }
    this.inp = { fy: c.fy, from, to: to < from ? from : to, books, portal, accMap, overrides: this.st.ov, bal: this.st.bal, threeB };
    try { this.res = portal.length || books.length ? run(this.inp) : null; } catch (e: any) { this.res = null; $('flFlow').innerHTML = `<p class="warn-error">${esc(e.message)}</p>`; }
    this.renderPortal(portal);
    this.renderBooks(books, accs, regs);
    this.renderBalances(books, portal);
    this.renderFlow();
    this.renderReasons();
    this.renderChecks();
    $('flBuild').disabled = !this.res;
    this.save();
  }

  private guide(src: ('portal' | 'tally' | 'zoho' | 'other')[], title: string) {
    const items = FLOW_GUIDE.filter(g => src.includes(g.source));
    return `<details class="guide"><summary><span class="gq">?</span> ${title}</summary><div class="gpane on"><div class="gitems">${items.map(g => `<div class="gitem"><h4>${esc(g.title)} <span class="gfor">${g.source === 'portal' ? 'GST portal' : g.source === 'tally' ? 'Tally' : g.source === 'zoho' ? 'Zoho Books' : 'Other'}</span></h4><ol>${g.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ol>${g.note ? `<p class="muted">${esc(g.note)}</p>` : ''}</div>`).join('')}</div></div></details>`;
  }

  // ---------- 1. portal ledgers
  private renderPortal(portal: Inputs['portal']) {
    const c = this.ctx();
    const rows = portal.map(({ file, ledger: p }) => {
      const startsLate = p.kind !== 'liability' && p.from && ymOfDate(p.from) > this.inp!.from;
      return `<tr><td><b>${PORTAL_LABEL[p.kind]}</b><div class="muted">${esc(file)}</div></td><td>${esc(p.gstin)}${c.gstin && p.gstin && p.gstin !== c.gstin ? ' <span class="warn-error">other GSTIN</span>' : ''}</td><td>${esc(p.from)} to ${esc(p.to)}${startsLate ? `<div class="warn-warn">starts after ${ymLabel(this.inp!.from)}: enter the opening under Balances</div>` : ''}</td>
        <td class="r">${p.lines.length}</td><td class="r">${p.opening ? HEADS.map(h => `${h} ${fmt(p.opening![h])}`).join('<br/>') : '<span class="muted">none in file</span>'}</td><td class="r">${p.closing ? HEADS.map(h => `${h} ${fmt(p.closing![h])}`).join('<br/>') : ''}</td>
        <td>${p.problems.length ? `<span class="warn-warn">${p.problems.map(esc).join('; ')}</span>` : '<span class="ok-text">balances tie</span>'}</td></tr>`;
    }).join('');
    const missing = (['credit', 'cash', 'liability'] as PortalKind[]).filter(k => !portal.some(p => p.ledger.kind === k));
    $('flPortal').innerHTML = (portal.length ? `<div class="tablewrap"><table class="fl"><tr><th>Ledger</th><th>GSTIN</th><th>Period in file</th><th class="r">Lines</th><th class="r">Opening</th><th class="r">Closing</th><th>Check</th></tr>${rows}</table></div>` : '')
      + (missing.length ? `<p class="${portal.length ? 'warn-warn' : 'muted'}" style="margin:10px 0 0">${portal.length ? 'Not loaded: ' : 'Add '}${missing.map(k => PORTAL_LABEL[k]).join(', ')}.</p>` : '')
      + `<p class="muted" style="margin:8px 0 0">Portal entries are taken by return period: the March return filed in April counts in March. Download the credit and cash ledgers up to a date after the last return of the period was filed.</p>`
      + this.guide(['portal'], 'How to download the portal ledgers');
  }

  // ---------- 2. books accounts
  private renderBooks(books: BookFile[], accs: ReturnType<typeof accountsOf>, regs: InFile[]) {
    const shown = accs.filter(a => this.showAll || a.gst || (this.acc[a.name] && this.acc[a.name].type !== 'ignore'));
    const hidden = accs.length - shown.length;
    const opts = (cur: string, list: [string, string][]) => list.map(([v, l]) => `<option value="${v}" ${v === cur ? 'selected' : ''}>${esc(l)}</option>`).join('');
    const unconf = shown.filter(a => !this.acc[a.name]?.confirmed && this.acc[a.name]?.type !== 'ignore');
    const rows = shown.map(a => {
      const s = this.acc[a.name] || { type: 'ignore', head: '' };
      const needHead = (s.type === 'input' || s.type === 'output' || s.type === 'rcm_payable') && !s.head;
      return `<tr class="${s.confirmed ? 'confirmed' : ''}"><td><b>${esc(a.name)}</b><div class="muted">${[...a.files].map(esc).join(', ')}</div></td><td class="r">${a.lines}</td><td class="r">${fmt(a.dr)}</td><td class="r">${fmt(a.cr)}</td>
        <td><select data-acct="${esc(a.name)}">${opts(s.type, ACC_TYPES.map(t => [t, ACC_TYPE_LABEL[t]]))}</select></td>
        <td><select data-acch="${esc(a.name)}" ${s.type === 'ignore' ? 'disabled' : ''}>${opts(s.head || '', [['', '— pick —'], ...HEADS.map(h => [h, `${h} (${HEAD_LABEL[h]})`] as [string, string]), ['ALL', 'Not split by head']])}</select>${needHead ? '<div class="warn-error" style="font-size:12px">pick the tax head</div>' : ''}</td>
        <td class="act">${s.type === 'ignore' ? '<span class="tag no-tag">Not GST</span>' : s.confirmed ? '<span class="tag ok-tag">Confirmed</span>' : `<button class="small" data-accok="${esc(a.name)}">Confirm</button>`}</td></tr>`;
    }).join('');
    const files = books.map(b => `<li><b>${esc(b.name)}</b>: ${b.format === 'tally-dump' ? 'Tally voucher dump' : b.format === 'zoho-gl' ? 'Zoho detailed general ledger' : `single ledger <input data-lname="${esc(b.name)}" value="${esc(b.ledger)}" style="min-width:0;width:220px;padding:4px 8px;font-size:13px"/>${b.opening !== undefined ? ` · opening ${fmt(b.opening)}` : ''}${b.closing !== undefined ? ` · closing ${fmt(b.closing)}` : ''}`}, ${b.lines.length} lines${b.skipped.map(x => ` · <span class="warn-info">${x.count} line(s) left out: ${esc(x.reason)}</span>`).join('')}${b.problems.map(p => ` · <span class="warn-warn">${esc(p)}</span>`).join('')}</li>`).join('');
    $('flBooks').innerHTML = (books.length ? `<ul class="warns" style="margin:0 0 12px">${files}</ul>` : '<p class="muted">No books ledgers yet. Add GST ledgers or a GL dump under Add files.</p>')
      + (regs.length ? `<p class="muted">Read as registers (column mapping in the other tabs): ${regs.map(f => `${esc(f.name)} <button class="link" data-asbooks="${esc(f.name)}">read as a books ledger</button>`).join(' · ')}</p>` : '')
      + (shown.length ? `<div class="row" style="justify-content:space-between;margin:8px 0"><span class="muted">${unconf.length ? `${unconf.length} account(s) to confirm. ` : 'All accounts confirmed. '}Rate-wise ledgers (e.g. Output CGST 6%, 9%) are added up under their tax head.</span>
        <span>${unconf.length ? `<button class="small" id="flAccAll">Confirm all ${unconf.length}</button>` : ''} <button class="link" id="flAccShow">${this.showAll ? 'show GST accounts only' : `show all accounts${hidden ? ` (${hidden} more)` : ''}`}</button></span></div>
        <div class="tablewrap"><table class="fl sup"><tr><th>Account</th><th class="r">Lines</th><th class="r">Debit</th><th class="r">Credit</th><th>What it is</th><th>Tax head</th><th></th></tr>${rows}</table></div>` : '')
      + this.guide(['tally', 'zoho', 'other'], 'How to export the GST ledgers from Tally, Zoho and others');
    const box = $('flBooks');
    const setAcc = (name: string, patch: Partial<AccState>) => {
      const old = this.acc[name] || { type: 'ignore', head: '' }, nw = { ...old, ...patch };
      if (patch.type !== undefined && patch.type !== old.type || patch.head !== undefined && patch.head !== old.head)
        this.log('Account type / head', name, `${ACC_TYPE_LABEL[old.type]} · ${old.head || '-'}`, `${ACC_TYPE_LABEL[nw.type]} · ${nw.head || '-'}`, 'Set in the Loader (Books: GST accounts)');
      this.acc[name] = nw; this.refresh();
    };
    box.querySelectorAll('select[data-acct]').forEach((el: any) => (el.onchange = () => { const t = el.value as AccType; const g = guessAccount(el.dataset.acct); setAcc(el.dataset.acct, { type: t, head: (this.acc[el.dataset.acct]?.head || g.head || (blockOfType(t) === 'cash' || blockOfType(t) === 'fee' ? 'ALL' : '')) as any, confirmed: true }); }));
    box.querySelectorAll('select[data-acch]').forEach((el: any) => (el.onchange = () => setAcc(el.dataset.acch, { head: el.value as AccHead, confirmed: true })));
    box.querySelectorAll('button[data-accok]').forEach((el: any) => (el.onclick = () => setAcc(el.dataset.accok, { confirmed: true })));
    const all = box.querySelector('#flAccAll') as any; if (all) all.onclick = () => { for (const a of unconf) this.acc[a.name] = { ...this.acc[a.name], confirmed: true }; this.log('Account types confirmed', `${unconf.length} account(s)`, 'suggested', 'confirmed', 'Confirm all'); this.refresh(); };
    const sh = box.querySelector('#flAccShow') as any; if (sh) sh.onclick = () => { this.showAll = !this.showAll; this.refresh(); };
    box.querySelectorAll('input[data-lname]').forEach((el: any) => (el.onchange = async () => {
      const f = this.ctx().files.find(x => x.name === el.dataset.lname); if (!f?.bytes) return;
      const sheets = await readTable(f.name, f.bytes); f.books = parseBooks(f.name, sheets[0].rows, el.value.trim()); this.refresh();
    }));
    box.querySelectorAll('button[data-asbooks]').forEach((el: any) => (el.onclick = async () => {
      const f = this.ctx().files.find(x => x.name === el.dataset.asbooks); if (!f?.bytes) return;
      try { const sheets = await readTable(f.name, f.bytes); f.books = parseBooks(f.name, sheets[0].rows); f.kind = 'Books ledger'; f.note = `Ledger: ${f.books.ledger || f.books.format} · ${f.books.lines.length} lines`; this.onFilesChanged(); }
      catch (e: any) { alertBox(box, e.message); }
    }));
  }

  // ---------- 3. balances
  private renderBalances(books: BookFile[], portal: Inputs['portal']) {
    const inp = this.inp!, accs = Object.entries(inp.accMap).filter(([, m]) => { const b = blockOfType(m.type); return b && b !== 'fee'; });
    const fileOpen = new Map<string, number>(), fileClose = new Map<string, number>();
    for (const b of books) if (b.format === 'ledger' && b.ledger) { if (b.opening !== undefined) fileOpen.set(b.ledger, b.opening); if (b.closing !== undefined) fileClose.set(b.ledger, b.closing); }
    const inP = (d: string) => { const m = ymOfDate(d); return m >= inp.from && m <= inp.to; };
    const before = (d: string) => ymOfDate(d) < inp.from;
    let noOpen = 0;
    const rows = accs.map(([name, m]) => {
      const lines = books.flatMap(b => b.lines.filter(l => l.account === name));
      const dr = lines.filter(l => inP(l.date)).reduce((s, l) => s + l.dr, 0), cr = lines.filter(l => inP(l.date)).reduce((s, l) => s + l.cr, 0);
      const pre = lines.filter(l => before(l.date)).reduce((s, l) => s + l.dr - l.cr, 0);
      const ent = this.st.bal.books[name] || {};
      const op = ent.opening ?? fileOpen.get(name);
      if (op === undefined) noOpen++;
      const close = (op ?? 0) + pre + dr - cr;
      const tb = ent.closing ?? (inp.to === monthsOf(inp.fy)[11] ? fileClose.get(name) : undefined);
      const diff = tb !== undefined ? Math.round((close - tb) * 100) / 100 : undefined;
      return `<tr><td><b>${esc(name)}</b><div class="muted">${ACC_TYPE_LABEL[m.type]} · ${m.head || '-'}</div></td>
        <td class="r"><input class="amt" data-bo="${esc(name)}" value="${ent.opening ?? ''}" placeholder="${op !== undefined && ent.opening === undefined ? fmt(op) + ' (file)' : '0.00'}"/></td>
        <td class="r">${pre ? fmt(pre) : '–'}</td><td class="r">${fmt(dr)}</td><td class="r">${fmt(cr)}</td><td class="r"><b>${fmt(close)}</b></td>
        <td class="r"><input class="amt" data-bc="${esc(name)}" value="${ent.closing ?? ''}" placeholder="${tb !== undefined && ent.closing === undefined ? fmt(tb) + ' (file)' : 'from trial balance'}"/></td>
        <td class="r ${diff ? 'neg' : ''}">${diff === undefined ? '' : diff ? fmt(diff) : '<span class="ok-text">✓</span>'}</td></tr>`;
    }).join('');
    const kinds = [...new Set(portal.map(p => p.ledger.kind))];
    const prow = kinds.map(k => {
      const L = portal.filter(p => p.ledger.kind === k).map(p => p.ledger).sort((a, b) => ymOfDate(a.from).localeCompare(ymOfDate(b.from)))[0];
      const fromFile = L.from && ymOfDate(L.from) <= inp.from && L.opening;
      return `<tr><td><b>${PORTAL_LABEL[k]}</b><div class="muted">${fromFile ? `Opening Balance row (${esc(L.from)})` : k === 'liability' ? 'no opening in the register: nil unless entered' : `file starts ${esc(L.from)}: enter the balance on the first day`}</div></td>
        ${HEADS.map(h => { const ent = this.st.bal.portal[k]?.[h]?.opening; return `<td class="r"><input class="amt" data-po="${k}:${h}" value="${ent ?? ''}" placeholder="${fromFile ? fmt(L.opening![h]) + ' (file)' : '0.00'}"/></td>`; }).join('')}</tr>`;
    }).join('');
    $('flBal').innerHTML = (rows ? `<p class="grp">Books (debit +, credit −, as in the ledger)</p><div class="tablewrap"><table class="fl"><tr><th>Account</th><th class="r">Opening on ${ymLabel(inp.from)} start</th><th class="r">Entries before the period</th><th class="r">Debits</th><th class="r">Credits</th><th class="r">Closing (worked out)</th><th class="r">Closing per trial balance</th><th class="r">Check</th></tr>${rows}</table></div>` : '<p class="muted">Confirm the books GST accounts first.</p>')
      + (prow ? `<p class="grp">Portal (balance on the first day of the period)</p><div class="tablewrap"><table class="fl"><tr><th>Ledger</th>${HEADS.map(h => `<th class="r">${h}</th>`).join('')}</tr>${prow}</table></div>` : '')
      + (noOpen ? `<p class="warn-warn" style="margin:8px 0 0">${noOpen} books account(s) have no opening balance in the file and none entered, so they start at nil. If the period does not start with a nil balance, enter the opening from the trial balance.</p>` : '')
      + '<p class="muted" style="margin:8px 0 0">Leave a field empty to use the figure from the file. Every balance you type goes to the Audit sheet.</p>';
    const box = $('flBal');
    const num = (v: string) => { const s = v.replace(/[,\s₹]/g, ''); if (s === '') return undefined; const x = /^\(.*\)$/.test(s) ? -parseFloat(s.slice(1, -1)) : /cr$/i.test(s) ? -parseFloat(s) : parseFloat(s); return isFinite(x) ? x : undefined; };
    box.querySelectorAll('input[data-bo],input[data-bc]').forEach((el: any) => (el.onchange = () => {
      const name = el.dataset.bo || el.dataset.bc, kind = el.dataset.bo ? 'opening' : 'closing';
      const cur = (this.st.bal.books[name] ||= {}), old = (cur as any)[kind], v = num(el.value);
      if (v === undefined) delete (cur as any)[kind]; else (cur as any)[kind] = v;
      this.log(`Books ${kind} balance`, name, old === undefined ? '(file / none)' : String(old), v === undefined ? '(file / none)' : String(v), 'Entered in Balances');
      this.refresh();
    }));
    box.querySelectorAll('input[data-po]').forEach((el: any) => (el.onchange = () => {
      const [k, h] = el.dataset.po.split(':') as [PortalKind, Head];
      const kk = (this.st.bal.portal[k] ||= {}), cur = (kk[h] ||= {}), old = cur.opening, v = num(el.value);
      if (v === undefined) delete cur.opening; else cur.opening = v;
      this.log('Portal opening balance', `${PORTAL_LABEL[k]} ${h}`, old === undefined ? '(file / none)' : String(old), v === undefined ? '(file / none)' : String(v), 'Entered in Balances');
      this.refresh();
    }));
  }

  // ---------- 4. flow statements
  private renderFlow() {
    const r = this.res, inp = this.inp!;
    if (!r) { $('flFlow').innerHTML = '<p class="muted">Add the portal ledgers and the books GST ledgers to see the flow.</p>'; return; }
    const months = monthsOf(inp.fy);
    const sel = (id: string, cur: Ym) => `<select id="${id}">${months.map(m => `<option value="${m}" ${m === cur ? 'selected' : ''}>${ymLabel(m)}</option>`).join('')}</select>`;
    const useAll = HEAD_KEYS.includes('ALL') && [r.flowBooks, r.flowPortal].some(st => Object.values(st.rows).some(v => Math.abs(v.ALL) > 0.004) || BLOCKS.some(b => Math.abs(st.opening[b].ALL) > 0.004));
    const heads: HeadK[] = useAll ? HEAD_KEYS : [...HEADS];
    const unByRow = (b: Block, k: string) => r.items.filter(i => !i.explained && i.block === b && i.bucket === k).reduce((s, i) => s + HEAD_KEYS.reduce((t, h) => t + i.amt[h], 0), 0);
    const val = (which: 'books' | 'portal' | 'diff', get: (st: Result['flowBooks']) => Record<HeadK, number> | undefined) => {
      const b = get(r.flowBooks), p = get(r.flowPortal);
      return (h: HeadK) => which === 'books' ? b?.[h] || 0 : which === 'portal' ? p?.[h] || 0 : (b?.[h] || 0) - (p?.[h] || 0);
    };
    const tbl = (b: Block) => {
      const keys = [...BUCKETS.filter(x => x.block === b).map(x => x.id), `${b}:${INTERNAL}`, `${b}:${UNMAPPED}`].filter(k => r.flowBooks.rows[k] || r.flowPortal.rows[k] || (bucketOf(k) && b !== 'fee'));
      const line = (label: string, f: (h: HeadK) => number, cls = '', un?: number) => { const tot = heads.reduce((s, h) => s + f(h), 0); return `<tr class="${cls}"><td>${label}</td>${heads.map(h => `<td class="r">${fmt(f(h))}</td>`).join('')}<td class="r"><b>${fmt(tot)}</b></td>${this.view === 'diff' ? `<td class="r ${un && Math.abs(un) > 0.004 ? 'neg' : ''}">${un === undefined ? '' : fmt(un)}</td>` : ''}</tr>`; };
      let h = `<tr class="sec"><td colspan="${heads.length + 2 + (this.view === 'diff' ? 1 : 0)}">${BLOCK_LABEL[b]}</td></tr>`;
      if (b !== 'fee') h += line('Opening balance', val(this.view, st => st.opening[b]), '', this.view === 'diff' ? unByRow(b, 'opening') : undefined);
      for (const k of keys) { const s = bucketOf(k)?.sign ?? 1; h += line(`${b === 'fee' ? '' : s === 1 ? 'Add: ' : 'Less: '}${bucketLabel(k)}`, val(this.view, st => st.rows[k]), k.endsWith(UNMAPPED) ? 'warn-error' : '', this.view === 'diff' ? unByRow(b, k) : undefined); }
      h += line(b === 'fee' ? 'Total' : 'Closing balance', val(this.view, st => st.closing[b]), 'tot', this.view === 'diff' ? r.items.filter(i => !i.explained && i.block === b).reduce((s, i) => s + HEAD_KEYS.reduce((t, hh) => t + i.amt[hh] * (i.bucket === 'opening' ? 1 : bucketOf(i.bucket)?.sign ?? 1), 0), 0) : undefined);
      return h;
    };
    const tot = (st: Result['flowBooks'], b: Block) => HEAD_KEYS.reduce((s, h) => s + st.closing[b][h], 0);
    const unex = r.items.filter(i => !i.explained);
    const unTot = unex.reduce((s, i) => s + HEAD_KEYS.reduce((t, h) => t + Math.abs(i.amt[h]), 0), 0);
    const kpi = (label: string, b: Block) => { const bb = tot(r.flowBooks, b), pp = tot(r.flowPortal, b); return `<div class="kpi ${Math.abs(bb - pp) < 1 ? 'ok' : 'warn'}"><span>${label}: books / portal</span><b>${fmt(bb)}</b><em>portal ${fmt(pp)} · difference ${fmt(bb - pp)}</em></div>`; };
    $('flFlow').innerHTML = `<div class="rv-bar"><div class="row" style="gap:10px"><label>From${sel('flFrom', inp.from)}</label><label>To${sel('flTo', inp.to)}</label></div>
        <div class="seg">${(['books', 'portal', 'diff'] as const).map(v => `<button data-view="${v}" class="${this.view === v ? 'on' : ''}">${v === 'books' ? 'As per books' : v === 'portal' ? 'As per portal' : 'Difference'}</button>`).join('')}</div></div>
      <div class="kpis">${kpi('ITC closing', 'itc')}${kpi('Liability closing', 'liab')}${kpi('Cash ledger closing', 'cash')}<div class="kpi ${unex.length ? 'bad' : 'ok'}"><span>Unexplained</span><b>${unex.length}</b><em>${unex.length ? `item(s), ${fmt(unTot)} in total` : 'every difference is explained'}</em></div></div>
      <div class="tablewrap"><table class="fl"><tr><th>${this.view === 'books' ? 'Flow as per books' : this.view === 'portal' ? 'Flow as per GST portal' : 'Difference: books − portal'}</th>${heads.map(h => `<th class="r">${HK_TXT[h]}</th>`).join('')}<th class="r">Total</th>${this.view === 'diff' ? '<th class="r">Of which unexplained</th>' : ''}</tr>${BLOCKS.map(tbl).join('')}</table></div>
      ${useAll ? '<p class="muted" style="margin:8px 0 0">Books accounts that are not split by tax head (for example a single GST cash ledger account) are in "Not split", while the portal shows each head: compare the Total column for those rows.</p>' : ''}
      <p class="muted" style="margin:8px 0 0">${this.view === 'books' ? 'From the books ledgers only, by voucher date.' : this.view === 'portal' ? 'From the credit ledger, cash ledger and liability register only, by return period.' : 'The closing difference is the sum of the row differences. The Reconciliation below explains it item by item.'}
        ${r.warns.length ? ` <span class="warn-warn">${r.warns.length} message(s) under Cross-checks.</span>` : ''}</p>`;
    const box = $('flFlow');
    box.querySelectorAll('button[data-view]').forEach((el: any) => (el.onclick = () => { this.view = el.dataset.view; this.renderFlow(); }));
    ($('flFrom') as any).onchange = (e: any) => { this.st.from = e.target.value; this.refresh(); };
    ($('flTo') as any).onchange = (e: any) => { this.st.to = e.target.value; this.refresh(); };
  }

  // ---------- 5. reasons
  private renderReasons() {
    const r = this.res;
    if (!r) { $('flReasons').innerHTML = ''; return; }
    const byId = new Map(r.books.map(e => [e.id, e])), pById = new Map(r.portal.map(e => [e.id, e]));
    const months = monthsOf(this.inp!.fy), nextMonths = monthsOf(`${+this.inp!.fy.slice(0, 4) + 1}-${String(+this.inp!.fy.slice(2, 4) + 2).padStart(2, '0')}`).slice(0, 3), prev = monthsOf(`${+this.inp!.fy.slice(0, 4) - 1}-${this.inp!.fy.slice(2, 4)}`).slice(-3);
    const periodOpts = [...prev, ...months, ...nextMonths];
    const bookRow = (e: BEntry) => {
      const ov = this.st.ov[e.id], editing = this.editing === e.id, l = e.line;
      const val = (bucketOf(e.bucket)?.sign ?? 1) * e.effect;
      let h = `<tr><td>${esc(l.date)}</td><td>${esc(l.vchType)} ${esc(l.vchNo)}</td><td>${esc(l.party)}</td><td>${esc(l.account)}</td><td>${e.head}</td><td class="r">${fmt(val)}</td><td>${esc(bucketLabel(e.bucket))}<div class="muted" style="font-size:11.5px">${ov ? 'changed by you' : esc(e.rule)}${e.month !== ymOfDate(l.date) ? ` · compared with ${ymLabel(e.month)}` : ''}</div></td>
        <td class="act"><button class="link" data-edit="${esc(e.id)}">${editing ? 'close' : 'change'}</button></td></tr>`;
      if (editing) {
        const allowed = allowedBuckets(e.accType);
        h += `<tr><td colspan="8"><div class="edit" data-for="${esc(e.id)}">
          <label>Flow row<select data-k="bucket">${[...allowed, UNMAPPED].map(b => `<option value="${b}" ${b === e.bucket ? 'selected' : ''}>${esc(bucketLabel(b))}</option>`).join('')}</select></label>
          <label>Tax head<select data-k="head">${[...HEADS, 'ALL'].map(x => `<option value="${x}" ${x === e.head ? 'selected' : ''}>${HK_TXT[x as HeadK]}</option>`).join('')}</select></label>
          <label>Compare with return period<select data-k="period">${periodOpts.map(m => `<option value="${m}" ${m === e.month ? 'selected' : ''}>${ymLabel(m)}</option>`).join('')}</select></label>
          <label class="remark" style="flex:1 1 260px">Remark (required)<input class="remark" data-k="remark" placeholder="why: e.g. invoice claimed in the April return (2B)"/></label>
          <button class="small" data-save="${esc(e.id)}">Save</button>${ov ? ` <button class="link" data-reset="${esc(e.id)}">undo my change</button>` : ''}</div></td></tr>`;
      }
      return h;
    };
    const rc = reconcile(r);
    const detail = (ev: ReconEvent) => {
      const bl = ev.books.map(id => byId.get(id)).filter(Boolean) as BEntry[];
      const pl = ev.portal.map(id => pById.get(id)).filter(Boolean) as Result['portal'];
      return `<div class="item-in">${bl.length ? `<p class="grp">Books lines</p><div class="tablewrap"><table class="fl"><tr><th>Date</th><th>Voucher</th><th>Party / particulars</th><th>Account</th><th>Head</th><th class="r">Amount</th><th>Flow row</th><th></th></tr>${bl.slice(0, 200).map(bookRow).join('')}</table></div>${bl.length > 200 ? `<p class="muted">First 200 of ${bl.length}; all of them are in the workbook.</p>` : ''}` : ''}
          ${pl.length ? `<p class="grp">Portal lines</p><div class="tablewrap"><table class="fl"><tr><th>Ledger</th><th>Date</th><th>Reference</th><th>Return period</th><th>Description</th><th>Head</th><th>Part</th><th class="r">Amount</th></tr>${pl.slice(0, 200).map(p => `<tr><td>${PORTAL_LABEL[p.kind]}</td><td>${esc(p.date)}</td><td>${esc(p.ref)}</td><td>${ymLabel(p.month)}</td><td>${esc(p.desc)}${p.used && p.used !== '-' ? ` (${esc(p.used)})` : ''}</td><td>${p.head}</td><td>${p.comp}</td><td class="r">${fmt(p.amount)}</td></tr>`).join('')}</table></div>` : ''}
          ${!bl.length && !pl.length ? '<p class="muted">From the balances: see Opening and closing balances above.</p>' : ''}</div>`;
    };
    const tagUn = '<span class="tag" style="color:var(--err);background:var(--err-bg)">Unexplained</span> ';
    const evText = (ev: ReconEvent) => `${ev.explained ? '' : tagUn}<b>${esc(ev.what)}</b><div class="muted" style="font-size:12.5px">${esc(ev.action)}${ev.vouchers ? ` · ${esc(ev.vouchers)}` : ''}</div>
      <button class="link" data-ev="${esc(ev.key)}">${this.open.has(ev.key) ? 'Hide entries' : 'Show entries'}</button>`;
    const stmt = (s: ReconStatement) => {
      const hs = s.heads, n = hs.length + 2;
      const amts = (f: (h: HeadK) => number, cls = '') => hs.map(h => `<td class="r ${cls}">${fmt(f(h))}</td>`).join('') + `<td class="r ${cls}"><b>${fmt(hs.reduce((t, h) => t + f(h), 0))}</b></td>`;
      const anyUn = hs.some(h => Math.abs(s.unexplained[h]) > 0.004);
      return `<p class="grp">${esc(ledgerLabel(s.block))}</p><div class="tablewrap"><table class="fl rec"><tr><th></th>${hs.map(h => `<th class="r">${HK_TXT[h]}</th>`).join('')}<th class="r">Total</th></tr>
        <tr class="tot"><td>Closing as per books</td>${amts(h => s.books[h])}</tr>
        ${s.lines.map(ev => `<tr class="${ev.explained ? '' : 'unx'}"><td>${evText(ev)}</td>${amts(h => -ev.effect[h])}</tr>${this.open.has(ev.key) ? `<tr><td colspan="${n}">${detail(ev)}</td></tr>` : ''}`).join('')
          || `<tr><td colspan="${n}" class="ok-text">No difference: books and portal agree.</td></tr>`}
        <tr class="tot"><td>Closing as per portal</td>${amts(h => s.portal[h])}</tr>
        ${anyUn ? `<tr><td class="neg">Of which unexplained</td>${amts(h => -s.unexplained[h], 'neg')}</tr>` : ''}
        ${hs.some(h => Math.abs(s.gap[h]) > 0.004) ? `<tr><td class="neg">Not reconciled (should be nil)</td>${amts(h => -s.gap[h], 'neg')}</tr>` : ''}</table></div>`;
    };
    const ne = rc.noEffect;
    const unN = rc.events.filter(e => !e.explained).length;
    $('flReasons').innerHTML = `<p class="muted" style="margin:0 0 6px">Each statement starts from the closing balance as per books, adds or takes away each item, and ends at the closing balance as per the portal. ${unN ? `<span class="warn-error">${unN} item(s) are unexplained: check those first.</span>` : '<span class="ok-text">Every difference is explained.</span>'}</p>
      ${rc.statements.map(stmt).join('')}
      ${ne.length ? `<details class="rv-hist" ${this.open.has('__ne') ? 'open' : ''} data-ne="1" style="margin-top:14px"><summary class="eyebrow" style="cursor:pointer">Differences that reversed within the period: no effect on the closing (${ne.length})</summary>
        <div class="tablewrap" style="margin-top:8px"><table class="fl rec"><tr><th></th><th>Ledger</th><th class="r">Amount involved</th></tr>${ne.map(ev => `<tr class="${ev.explained ? '' : 'unx'}"><td>${evText(ev)}</td><td>${esc(ledgerLabel(ev.block))}</td><td class="r">${ev.involved ? fmt(ev.involved) : '–'}</td></tr>${this.open.has(ev.key) ? `<tr><td colspan="3">${detail(ev)}</td></tr>` : ''}`).join('')}</table></div></details>` : ''}
      ${this.st.audit.length ? `<details class="rv-hist"><summary class="eyebrow" style="cursor:pointer">Your changes (${this.st.audit.length})</summary><ol>${this.st.audit.slice().reverse().map(a => `<li>${esc(a.when)} · <b>${esc(a.change)}</b> · ${esc(a.record)}: ${esc(a.from)} → ${esc(a.to)}${a.remark ? ` · <i>${esc(a.remark)}</i>` : ''}</li>`).join('')}</ol></details>` : ''}`;
    const box = $('flReasons');
    box.querySelectorAll('button[data-ev]').forEach((el: any) => (el.onclick = (ev: Event) => { ev.preventDefault(); const k = el.dataset.ev; if (this.open.has(k)) this.open.delete(k); else this.open.add(k); this.renderReasons(); }));
    const neEl = box.querySelector('details[data-ne]') as any; if (neEl) neEl.addEventListener('toggle', () => { if (neEl.open) this.open.add('__ne'); else this.open.delete('__ne'); });
    box.querySelectorAll('button[data-edit]').forEach((el: any) => (el.onclick = (ev: Event) => { ev.preventDefault(); this.editing = this.editing === el.dataset.edit ? '' : el.dataset.edit; this.renderReasons(); }));
    box.querySelectorAll('button[data-reset]').forEach((el: any) => (el.onclick = () => {
      const e = byId.get(el.dataset.reset)!; const old = this.st.ov[e.id];
      delete this.st.ov[e.id]; this.editing = '';
      this.log('Line change undone', recTxt(e), describeOv(old), 'tool classification', 'Undone in the Loader'); this.refresh();
    }));
    box.querySelectorAll('button[data-save]').forEach((el: any) => (el.onclick = () => {
      const id = el.dataset.save, e = byId.get(id)!, f = box.querySelector(`.edit[data-for="${CSS.escape(id)}"]`) as HTMLElement;
      const g = (k: string) => (f.querySelector(`[data-k="${k}"]`) as any).value as string;
      const remark = g('remark').trim();
      if (!remark) { (f.querySelector('[data-k="remark"]') as any).focus(); alertBox(f, 'A remark is needed: say why this line belongs elsewhere.'); return; }
      const ov: LineOverride = {};
      if (g('bucket') !== e.auto) ov.bucket = g('bucket');
      const accHead = this.inp!.accMap[e.line.account]?.head || 'ALL';
      if (g('head') !== accHead) ov.head = g('head') as AccHead;
      if (g('period') !== ymOfDate(e.line.date)) ov.period = g('period');
      const before = `${bucketLabel(e.bucket)} · ${e.head} · ${ymLabel(e.month)}`;
      if (Object.keys(ov).length) this.st.ov[id] = ov; else delete this.st.ov[id];
      this.log('Line changed', recTxt(e), before, `${bucketLabel(g('bucket'))} · ${g('head')} · ${ymLabel(g('period'))}`, remark);
      this.editing = ''; this.refresh();
    }));
  }

  // ---------- 6. checks
  private renderChecks() {
    const r = this.res;
    if (!r) { $('flChecks').innerHTML = '<p class="muted">Load at least two portal ledgers.</p>'; return; }
    const bad = r.checks.filter(c => !c.ok), ok = r.checks.length - bad.length;
    const rows = [...bad, ...r.checks.filter(c => c.ok)].slice(0, 400).map(c => `<tr><td>${esc(c.area)}</td><td>${ymLabel(c.period)}</td><td>${esc(c.head)}</td><td class="r">${fmt(c.a)}</td><td class="r">${fmt(c.b)}</td><td class="r ${c.ok ? '' : 'neg'}">${fmt(c.diff)}</td><td>${c.ok ? '<span class="ok-text">OK</span>' : '<span class="warn-error">Check</span>'}${c.note ? `<div class="muted" style="font-size:11.5px">${esc(c.note)}</div>` : ''}</td></tr>`).join('');
    $('flChecks').innerHTML = `${r.warns.length ? `<ul class="warns" style="margin:0 0 12px">${r.warns.map(w => `<li class="warn-${w.level}">${esc(w.msg)}</li>`).join('')}</ul>` : ''}
      ${r.checks.length ? `<p class="${bad.length ? 'warn-warn' : 'ok-text'}" style="margin:0 0 8px">${ok} of ${r.checks.length} checks agree${bad.length ? `; ${bad.length} to look at` : ''}.</p>
      <div class="tablewrap" style="max-height:480px;overflow:auto"><table class="fl"><tr><th>Check</th><th>Return period</th><th>Head</th><th class="r">A</th><th class="r">B</th><th class="r">A − B</th><th></th></tr>${rows}</table></div>
      <p class="muted" style="margin:8px 0 0">A and B: cash ledger debits vs liability paid in cash; credit ledger debits vs liability paid by ITC; credit ledger ITC vs GSTR-3B 4A − 4B; liability register vs GSTR-3B 6.1 (where the 3B PDFs are loaded).</p>`
      : '<p class="muted">No cross-checks yet: they need two of the portal ledgers, or a portal ledger and GSTR-3B PDFs.</p>'}`;
  }

  // ---------- 7. workbook
  private async build() {
    const r = this.res, c = this.ctx(); if (!r || !this.inp) return;
    $('flStatus').textContent = 'Building…';
    try {
      const bytes = await flowWorkbook(r, { gstin: c.gstin, legalName: c.legal, fy: c.fy, from: this.inp.from, to: this.inp.to, builtAt: nowTxt(), audit: this.st.audit,
        files: c.files.filter(f => ['Portal ledger', 'Books ledger', 'GSTR-3B'].includes(f.kind)).map(f => ({ name: f.name, kind: f.kind, note: f.note || '' })) });
      const full = this.inp.from === monthsOf(c.fy)[0] && this.inp.to === monthsOf(c.fy)[11];
      const name = `GST_Flow_${c.gstin || 'client'}_${c.fy}${full ? '' : `_${ymLabel(this.inp.from)}_to_${ymLabel(this.inp.to)}`}.xlsx`;
      const url = URL.createObjectURL(new Blob([bytes as any], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a'); a.href = url; a.download = name; a.click();
      const un = r.items.filter(i => !i.explained).length;
      $('flResult').innerHTML = `<p><b>Downloaded ${esc(name)}</b> · <a href="${url}" download="${esc(name)}">download again</a>.</p>
        <p class="muted">Flow_Books, Flow_Portal and Flow_Difference are formulas over Books_Lines, Portal_Lines and Balances. ${r.items.length} reconciling item(s), ${un} unexplained, in Recon_Items. ${this.st.audit.length} change(s) of yours in Audit.</p>`;
      $('flStatus').textContent = '';
    } catch (e: any) { $('flStatus').innerHTML = `<span class="warn-error">${esc(e.message)}</span>`; }
  }
}

const recTxt = (e: BEntry) => `${e.line.file} row ${e.line.row} · ${e.line.date} · ${e.line.vchType} ${e.line.vchNo} · ${e.line.account}`;
const describeOv = (o?: LineOverride) => (o ? [o.bucket && bucketLabel(o.bucket), o.head, o.period && ymLabel(o.period)].filter(Boolean).join(' · ') : '-');
function alertBox(el: HTMLElement, msg: string) {
  let p = el.querySelector('.flmsg') as HTMLElement | null;
  if (!p) { p = document.createElement('p'); p.className = 'flmsg warn-error'; p.style.margin = '6px 0 0'; el.appendChild(p); }
  p.textContent = msg;
}

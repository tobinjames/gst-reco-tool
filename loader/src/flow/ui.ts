// GST ledger flow tab: portal ledgers, books GST accounts, balances, the three flow statements, reasons, cross-checks, workbook
import type { InFile } from '../build';
import {
  AccHead,
  AccType,
  ACC_TYPE_LABEL,
  BLOCK_LABEL,
  BUCKETS,
  Block,
  BookFile,
  HEADS,
  HEAD_LABEL,
  Head,
  INTERNAL,
  PORTAL_LABEL,
  PortalKind,
  UNMAPPED,
  Ym,
  bucketOf,
} from './types';
import { accountsOf, guessAccount, isGstName, parseBooks } from './books';
import { ymOfDate } from './portal';
import {
  AccMapEntry,
  BEntry,
  HEAD_KEYS,
  HeadK,
  Inputs,
  Item,
  LineOverride,
  Result,
  allowedBuckets,
  blockOfType,
  bucketLabel,
  monthsOf,
  run,
  ymLabel,
} from './engine';
import { AuditRow, flowWorkbook } from './workbook';
import { BLOCK_NAME, Comp, Year, yearRecon } from './year';
import { BridgeLine, Itc3, itcThreeWay } from './itc3';
import { process2b } from '../gstr2b';
import { esc, fyShift, inr, store } from '../util';
import { FLOW_GUIDE } from './guide';
import { readTable } from '../register';
import { Settings, describe } from '../settings';

const fmt = (v: number) => (Math.abs(v) < 0.005 ? '–' : inr(v));
const $ = (id: string) => document.getElementById(id)! as any;
const HK_TXT: Record<HeadK, string> = { IGST: 'IGST', CGST: 'CGST', SGST: 'SGST', CESS: 'Cess', ALL: 'Not split' };
const BLOCKS: Block[] = ['itc', 'liab', 'cash', 'fee'];
const ACC_TYPES = Object.keys(ACC_TYPE_LABEL) as AccType[];
const nowTxt = () => {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
};

interface AccState extends AccMapEntry {
  confirmed?: boolean;
}
interface State {
  ov: Record<string, LineOverride>;
  bal: Inputs['bal'];
  audit: AuditRow[];
  from?: Ym;
  to?: Ym;
}
export interface Ctx {
  files: InFile[];
  fy: string;
  gstin: string;
  legal: string;
  active: boolean;
  /** purchase register rows (optional cross-check) */ register?: () => any[][] | undefined;
  /** the tab's matching settings */
  settings?: () => Settings;
}

export class FlowTab {
  private acc: Record<string, AccState> = {};
  private st: State = { ov: {}, bal: { books: {}, portal: {} }, audit: [] };
  private key = '';
  private view: 'books' | 'portal' | 'diff' = 'diff';
  private showAll = false;
  private editing = '';
  private open = new Set<string>();
  private res: Result | null = null;
  private year: Year | null = null;
  private itc: Itc3 | null = null;
  private inp: Inputs | null = null;
  private ledgerNames: Record<string, string> = {};
  constructor(
    private ctx: () => Ctx,
    private onFilesChanged: () => void,
  ) {
    $('flBuild').onclick = () => this.build();
  }

  private load(c: Ctx) {
    const k = `${c.gstin || '-'}:${c.fy}`;
    if (k === this.key) return;
    this.key = k;
    this.acc = store.get(`jjflow:acc:${c.gstin || '-'}`) || {};
    this.st = { ov: {}, bal: { books: {}, portal: {} }, audit: [], ...(store.get(`jjflow:${k}`) || {}) };
  }
  private save() {
    const c = this.ctx();
    store.set(`jjflow:acc:${c.gstin || '-'}`, this.acc);
    store.set(`jjflow:${this.key}`, this.st);
  }
  private log(change: string, record: string, from: string, to: string, remark: string) {
    this.st.audit.push({ when: nowTxt(), change, record, from, to, remark });
  }

  /** files of the tab */
  private parts(c: Ctx) {
    const portal = c.files.filter(f => f.kind === 'Portal ledger' && f.portal).map(f => ({ file: f.name, ledger: f.portal! }));
    const books = c.files.filter(f => f.kind === 'Books ledger' && f.books).map(f => f.books!);
    return {
      portal,
      books,
      threeB: c.files.filter(f => f.kind === 'GSTR-3B' && f.threeB).map(f => f.threeB!),
      regs: c.files.filter(f => f.kind === 'Register'),
    };
  }

  /** summary for the shared Add files card */
  coverageHtml(): string {
    const c = this.ctx();
    const { portal, books, threeB } = this.parts(c);
    const has = (k: PortalKind) => portal.filter(p => p.ledger.kind === k);
    const row = (label: string, list: { file: string; ledger: any }[], opt = false) =>
      `<tr><th class="cvl">${label}</th><td class="cv ${list.length ? 'ok' : opt ? 'opt' : 'miss'} wide">${list.length ? '✓ ' + list.map(p => `${esc(p.ledger.from)} to ${esc(p.ledger.to)}`).join(' · ') : opt ? '– optional' : '✕ not loaded'}</td></tr>`;
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
    for (const a of accs)
      if (!this.acc[a.name] && a.gst) {
        const g = guessAccount(a.name);
        this.acc[a.name] = {
          type: isGstName(a.name) || books.some(b => b.format === 'ledger' && b.ledger === a.name) ? g.type : 'ignore',
          head: g.head,
        };
      }
    const months = monthsOf(c.fy);
    const from = this.st.from && months.includes(this.st.from) ? this.st.from : months[0],
      to = this.st.to && months.includes(this.st.to) ? this.st.to : months[11];
    const accMap: Inputs['accMap'] = {};
    for (const a of accs) {
      const s = this.acc[a.name];
      if (s && s.type !== 'ignore') accMap[a.name] = { type: s.type, head: s.head };
    }
    this.inp = {
      fy: c.fy,
      from,
      to: to < from ? from : to,
      books,
      portal,
      accMap,
      overrides: this.st.ov,
      bal: this.st.bal,
      threeB,
      opts: c.settings?.().flow,
    };
    try {
      this.res = portal.length || books.length ? run(this.inp) : null;
    } catch (e: any) {
      this.res = null;
      $('flFlow').innerHTML = `<p class="warn-error">${esc(e.message)}</p>`;
    }
    const b2 = c.files.filter(f => f.kind === 'GSTR-2B' && f.json).map(f => ({ name: f.name, json: f.json }));
    let raw2b: any[][] = [];
    try {
      raw2b = b2.length
        ? process2b(b2, c.fy, fyShift(c.fy, -1), fyShift(c.fy, 1), []).rows.map(x => [x.key || null, ...x.amt, ...x.info])
        : [];
    } catch {
      raw2b = [];
    }
    let reg: any[][] | undefined;
    try {
      reg = c.register?.();
    } catch {
      reg = undefined;
    }
    this.year = this.res ? yearRecon(this.res, this.inp, { has2b: raw2b.length > 0 }) : null;
    this.itc = this.res && raw2b.length ? itcThreeWay(this.res, this.inp, raw2b, reg, c.settings?.().match) : null;
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
    const rows = portal
      .map(({ file, ledger: p }) => {
        const startsLate = p.kind !== 'liability' && p.from && ymOfDate(p.from) > this.inp!.from;
        return `<tr><td><b>${PORTAL_LABEL[p.kind]}</b><div class="muted">${esc(file)}</div></td><td>${esc(p.gstin)}${c.gstin && p.gstin && p.gstin !== c.gstin ? ' <span class="warn-error">other GSTIN</span>' : ''}</td><td>${esc(p.from)} to ${esc(p.to)}${startsLate ? `<div class="warn-warn">starts after ${ymLabel(this.inp!.from)}: enter the opening under Balances</div>` : ''}</td>
        <td class="r">${p.lines.length}</td><td class="r">${p.opening ? HEADS.map(h => `${h} ${fmt(p.opening![h])}`).join('<br/>') : '<span class="muted">none in file</span>'}</td><td class="r">${p.closing ? HEADS.map(h => `${h} ${fmt(p.closing![h])}`).join('<br/>') : ''}</td>
        <td>${p.problems.length ? `<span class="warn-warn">${p.problems.map(esc).join('; ')}</span>` : '<span class="ok-text">balances tie</span>'}</td></tr>`;
      })
      .join('');
    const missing = (['credit', 'cash', 'liability'] as PortalKind[]).filter(k => !portal.some(p => p.ledger.kind === k));
    $('flPortal').innerHTML =
      (portal.length
        ? `<div class="tablewrap"><table class="fl"><tr><th>Ledger</th><th>GSTIN</th><th>Period in file</th><th class="r">Lines</th><th class="r">Opening</th><th class="r">Closing</th><th>Check</th></tr>${rows}</table></div>`
        : '') +
      (missing.length
        ? `<p class="${portal.length ? 'warn-warn' : 'muted'}" style="margin:10px 0 0">${portal.length ? 'Not loaded: ' : 'Add '}${missing.map(k => PORTAL_LABEL[k]).join(', ')}.</p>`
        : '') +
      `<p class="muted" style="margin:8px 0 0">Portal entries are taken by return period: the March return filed in April counts in March. Download the credit and cash ledgers up to a date after the last return of the period was filed.</p>` +
      this.guide(['portal'], 'How to download the portal ledgers');
  }

  // ---------- 2. books accounts
  private renderBooks(books: BookFile[], accs: ReturnType<typeof accountsOf>, regs: InFile[]) {
    const shown = accs.filter(a => this.showAll || a.gst || (this.acc[a.name] && this.acc[a.name].type !== 'ignore'));
    const hidden = accs.length - shown.length;
    const opts = (cur: string, list: [string, string][]) =>
      list.map(([v, l]) => `<option value="${v}" ${v === cur ? 'selected' : ''}>${esc(l)}</option>`).join('');
    const unconf = shown.filter(a => !this.acc[a.name]?.confirmed && this.acc[a.name]?.type !== 'ignore');
    const rows = shown
      .map(a => {
        const s = this.acc[a.name] || { type: 'ignore', head: '' };
        const needHead = (s.type === 'input' || s.type === 'output' || s.type === 'rcm_payable') && !s.head;
        return `<tr class="${s.confirmed ? 'confirmed' : ''}"><td><b>${esc(a.name)}</b><div class="muted">${[...a.files].map(esc).join(', ')}</div></td><td class="r">${a.lines}</td><td class="r">${fmt(a.dr)}</td><td class="r">${fmt(a.cr)}</td>
        <td><select data-acct="${esc(a.name)}">${opts(
          s.type,
          ACC_TYPES.map(t => [t, ACC_TYPE_LABEL[t]]),
        )}</select></td>
        <td><select data-acch="${esc(a.name)}" ${s.type === 'ignore' ? 'disabled' : ''}>${opts(s.head || '', [['', '— pick —'], ...HEADS.map(h => [h, `${h} (${HEAD_LABEL[h]})`] as [string, string]), ['ALL', 'Not split by head']])}</select>${needHead ? '<div class="warn-error" style="font-size:12px">pick the tax head</div>' : ''}</td>
        <td class="act">${s.type === 'ignore' ? '<span class="tag no-tag">Not GST</span>' : s.confirmed ? '<span class="tag ok-tag">Confirmed</span>' : `<button class="small" data-accok="${esc(a.name)}">Confirm</button>`}</td></tr>`;
      })
      .join('');
    const files = books
      .map(
        b =>
          `<li><b>${esc(b.name)}</b>: ${b.format === 'tally-dump' ? 'Tally voucher dump' : b.format === 'zoho-gl' ? 'Zoho detailed general ledger' : `single ledger <input data-lname="${esc(b.name)}" value="${esc(b.ledger)}" style="min-width:0;width:220px;padding:4px 8px;font-size:13px"/>${b.opening !== undefined ? ` · opening ${fmt(b.opening)}` : ''}${b.closing !== undefined ? ` · closing ${fmt(b.closing)}` : ''}`}, ${b.lines.length} lines${b.skipped.map(x => ` · <span class="warn-info">${x.count} line(s) left out: ${esc(x.reason)}</span>`).join('')}${b.problems.map(p => ` · <span class="warn-warn">${esc(p)}</span>`).join('')}</li>`,
      )
      .join('');
    $('flBooks').innerHTML =
      (books.length
        ? `<ul class="warns" style="margin:0 0 12px">${files}</ul>`
        : '<p class="muted">No books ledgers yet. Add GST ledgers or a GL dump under Add files.</p>') +
      (regs.length
        ? `<p class="muted">Read as registers (column mapping in the other tabs): ${regs.map(f => `${esc(f.name)} <button class="link" data-asbooks="${esc(f.name)}">read as a books ledger</button>`).join(' · ')}</p>`
        : '') +
      (shown.length
        ? `<div class="row" style="justify-content:space-between;margin:8px 0"><span class="muted">${unconf.length ? `${unconf.length} account(s) to confirm. ` : 'All accounts confirmed. '}Rate-wise ledgers (e.g. Output CGST 6%, 9%) are added up under their tax head.</span>
        <span>${unconf.length ? `<button class="small" id="flAccAll">Confirm all ${unconf.length}</button>` : ''} <button class="link" id="flAccShow">${this.showAll ? 'show GST accounts only' : `show all accounts${hidden ? ` (${hidden} more)` : ''}`}</button></span></div>
        <div class="tablewrap"><table class="fl sup"><tr><th>Account</th><th class="r">Lines</th><th class="r">Debit</th><th class="r">Credit</th><th>What it is</th><th>Tax head</th><th></th></tr>${rows}</table></div>`
        : '') +
      this.guide(['tally', 'zoho', 'other'], 'How to export the GST ledgers from Tally, Zoho and others');
    const box = $('flBooks');
    const setAcc = (name: string, patch: Partial<AccState>) => {
      const old = this.acc[name] || { type: 'ignore', head: '' },
        nw = { ...old, ...patch };
      if ((patch.type !== undefined && patch.type !== old.type) || (patch.head !== undefined && patch.head !== old.head))
        this.log(
          'Account type / head',
          name,
          `${ACC_TYPE_LABEL[old.type]} · ${old.head || '-'}`,
          `${ACC_TYPE_LABEL[nw.type]} · ${nw.head || '-'}`,
          'Set in the Loader (Books: GST accounts)',
        );
      this.acc[name] = nw;
      this.refresh();
    };
    box.querySelectorAll('select[data-acct]').forEach(
      (el: any) =>
        (el.onchange = () => {
          const t = el.value as AccType;
          const g = guessAccount(el.dataset.acct);
          setAcc(el.dataset.acct, {
            type: t,
            head: (this.acc[el.dataset.acct]?.head ||
              g.head ||
              (blockOfType(t) === 'cash' || blockOfType(t) === 'fee' ? 'ALL' : '')) as any,
            confirmed: true,
          });
        }),
    );
    box
      .querySelectorAll('select[data-acch]')
      .forEach((el: any) => (el.onchange = () => setAcc(el.dataset.acch, { head: el.value as AccHead, confirmed: true })));
    box
      .querySelectorAll('button[data-accok]')
      .forEach((el: any) => (el.onclick = () => setAcc(el.dataset.accok, { confirmed: true })));
    const all = box.querySelector('#flAccAll') as any;
    if (all)
      all.onclick = () => {
        for (const a of unconf) this.acc[a.name] = { ...this.acc[a.name], confirmed: true };
        this.log('Account types confirmed', `${unconf.length} account(s)`, 'suggested', 'confirmed', 'Confirm all');
        this.refresh();
      };
    const sh = box.querySelector('#flAccShow') as any;
    if (sh)
      sh.onclick = () => {
        this.showAll = !this.showAll;
        this.refresh();
      };
    box.querySelectorAll('input[data-lname]').forEach(
      (el: any) =>
        (el.onchange = async () => {
          const f = this.ctx().files.find(x => x.name === el.dataset.lname);
          if (!f?.bytes) return;
          const sheets = await readTable(f.name, f.bytes);
          f.books = parseBooks(f.name, sheets[0].rows, el.value.trim());
          this.refresh();
        }),
    );
    box.querySelectorAll('button[data-asbooks]').forEach(
      (el: any) =>
        (el.onclick = async () => {
          const f = this.ctx().files.find(x => x.name === el.dataset.asbooks);
          if (!f?.bytes) return;
          try {
            const sheets = await readTable(f.name, f.bytes);
            f.books = parseBooks(f.name, sheets[0].rows);
            f.kind = 'Books ledger';
            f.note = `Ledger: ${f.books.ledger || f.books.format} · ${f.books.lines.length} lines`;
            this.onFilesChanged();
          } catch (e: any) {
            alertBox(box, e.message);
          }
        }),
    );
  }

  // ---------- 3. balances
  private renderBalances(books: BookFile[], portal: Inputs['portal']) {
    const inp = this.inp!,
      accs = Object.entries(inp.accMap).filter(([, m]) => {
        const b = blockOfType(m.type);
        return b && b !== 'fee';
      });
    const fileOpen = new Map<string, number>(),
      fileClose = new Map<string, number>();
    for (const b of books)
      if (b.format === 'ledger' && b.ledger) {
        if (b.opening !== undefined) fileOpen.set(b.ledger, b.opening);
        if (b.closing !== undefined) fileClose.set(b.ledger, b.closing);
      }
    const inP = (d: string) => {
      const m = ymOfDate(d);
      return m >= inp.from && m <= inp.to;
    };
    const before = (d: string) => ymOfDate(d) < inp.from;
    let noOpen = 0;
    const rows = accs
      .map(([name, m]) => {
        const lines = books.flatMap(b => b.lines.filter(l => l.account === name));
        const dr = lines.filter(l => inP(l.date)).reduce((s, l) => s + l.dr, 0),
          cr = lines.filter(l => inP(l.date)).reduce((s, l) => s + l.cr, 0);
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
      })
      .join('');
    const kinds = [...new Set(portal.map(p => p.ledger.kind))];
    const prow = kinds
      .map(k => {
        const L = portal
          .filter(p => p.ledger.kind === k)
          .map(p => p.ledger)
          .sort((a, b) => ymOfDate(a.from).localeCompare(ymOfDate(b.from)))[0];
        const fromFile = L.from && ymOfDate(L.from) <= inp.from && L.opening;
        return `<tr><td><b>${PORTAL_LABEL[k]}</b><div class="muted">${fromFile ? `Opening Balance row (${esc(L.from)})` : k === 'liability' ? 'no opening in the register: nil unless entered' : `file starts ${esc(L.from)}: enter the balance on the first day`}</div></td>
        ${HEADS.map(h => {
          const ent = this.st.bal.portal[k]?.[h]?.opening;
          return `<td class="r"><input class="amt" data-po="${k}:${h}" value="${ent ?? ''}" placeholder="${fromFile ? fmt(L.opening![h]) + ' (file)' : '0.00'}"/></td>`;
        }).join('')}</tr>`;
      })
      .join('');
    $('flBal').innerHTML =
      (rows
        ? `<p class="grp">Books (debit +, credit −, as in the ledger)</p><div class="tablewrap"><table class="fl"><tr><th>Account</th><th class="r">Opening on ${ymLabel(inp.from)} start</th><th class="r">Entries before the period</th><th class="r">Debits</th><th class="r">Credits</th><th class="r">Closing (worked out)</th><th class="r">Closing per trial balance</th><th class="r">Check</th></tr>${rows}</table></div>`
        : '<p class="muted">Confirm the books GST accounts first.</p>') +
      (prow
        ? `<p class="grp">Portal (balance on the first day of the period)</p><div class="tablewrap"><table class="fl"><tr><th>Ledger</th>${HEADS.map(h => `<th class="r">${h}</th>`).join('')}</tr>${prow}</table></div>`
        : '') +
      (noOpen
        ? `<p class="warn-warn" style="margin:8px 0 0">${noOpen} books account(s) have no opening balance in the file and none entered, so they start at nil. If the period does not start with a nil balance, enter the opening from the trial balance.</p>`
        : '') +
      '<p class="muted" style="margin:8px 0 0">Leave a field empty to use the figure from the file. Every balance you type goes to the Audit sheet.</p>';
    const box = $('flBal');
    const num = (v: string) => {
      const s = v.replace(/[,\s₹]/g, '');
      if (s === '') return undefined;
      const x = /^\(.*\)$/.test(s) ? -parseFloat(s.slice(1, -1)) : /cr$/i.test(s) ? -parseFloat(s) : parseFloat(s);
      return isFinite(x) ? x : undefined;
    };
    box.querySelectorAll('input[data-bo],input[data-bc]').forEach(
      (el: any) =>
        (el.onchange = () => {
          const name = el.dataset.bo || el.dataset.bc,
            kind = el.dataset.bo ? 'opening' : 'closing';
          const cur = (this.st.bal.books[name] ||= {}),
            old = (cur as any)[kind],
            v = num(el.value);
          if (v === undefined) delete (cur as any)[kind];
          else (cur as any)[kind] = v;
          this.log(
            `Books ${kind} balance`,
            name,
            old === undefined ? '(file / none)' : String(old),
            v === undefined ? '(file / none)' : String(v),
            'Entered in Balances',
          );
          this.refresh();
        }),
    );
    box.querySelectorAll('input[data-po]').forEach(
      (el: any) =>
        (el.onchange = () => {
          const [k, h] = el.dataset.po.split(':') as [PortalKind, Head];
          const kk = (this.st.bal.portal[k] ||= {}),
            cur = (kk[h] ||= {}),
            old = cur.opening,
            v = num(el.value);
          if (v === undefined) delete cur.opening;
          else cur.opening = v;
          this.log(
            'Portal opening balance',
            `${PORTAL_LABEL[k]} ${h}`,
            old === undefined ? '(file / none)' : String(old),
            v === undefined ? '(file / none)' : String(v),
            'Entered in Balances',
          );
          this.refresh();
        }),
    );
  }

  // ---------- 4. flow statements
  private renderFlow() {
    const r = this.res,
      inp = this.inp!;
    if (!r) {
      $('flFlow').innerHTML = '<p class="muted">Add the portal ledgers and the books GST ledgers to see the flow.</p>';
      return;
    }
    const months = monthsOf(inp.fy);
    const sel = (id: string, cur: Ym) =>
      `<select id="${id}">${months.map(m => `<option value="${m}" ${m === cur ? 'selected' : ''}>${ymLabel(m)}</option>`).join('')}</select>`;
    const useAll =
      HEAD_KEYS.includes('ALL') &&
      [r.flowBooks, r.flowPortal].some(
        st => Object.values(st.rows).some(v => Math.abs(v.ALL) > 0.004) || BLOCKS.some(b => Math.abs(st.opening[b].ALL) > 0.004),
      );
    const heads: HeadK[] = useAll ? HEAD_KEYS : [...HEADS];
    const unByRow = (b: Block, k: string) =>
      r.items
        .filter(i => !i.explained && i.block === b && i.bucket === k)
        .reduce((s, i) => s + HEAD_KEYS.reduce((t, h) => t + i.amt[h], 0), 0);
    const val = (which: 'books' | 'portal' | 'diff', get: (st: Result['flowBooks']) => Record<HeadK, number> | undefined) => {
      const b = get(r.flowBooks),
        p = get(r.flowPortal);
      return (h: HeadK) => (which === 'books' ? b?.[h] || 0 : which === 'portal' ? p?.[h] || 0 : (b?.[h] || 0) - (p?.[h] || 0));
    };
    const tbl = (b: Block) => {
      const keys = [...BUCKETS.filter(x => x.block === b).map(x => x.id), `${b}:${INTERNAL}`, `${b}:${UNMAPPED}`].filter(
        k => r.flowBooks.rows[k] || r.flowPortal.rows[k] || (bucketOf(k) && b !== 'fee'),
      );
      const line = (label: string, f: (h: HeadK) => number, cls = '', un?: number) => {
        const tot = heads.reduce((s, h) => s + f(h), 0);
        return `<tr class="${cls}"><td>${label}</td>${heads.map(h => `<td class="r">${fmt(f(h))}</td>`).join('')}<td class="r"><b>${fmt(tot)}</b></td>${this.view === 'diff' ? `<td class="r ${un && Math.abs(un) > 0.004 ? 'neg' : ''}">${un === undefined ? '' : fmt(un)}</td>` : ''}</tr>`;
      };
      let h = `<tr class="sec"><td colspan="${heads.length + 2 + (this.view === 'diff' ? 1 : 0)}">${BLOCK_LABEL[b]}</td></tr>`;
      if (b !== 'fee')
        h += line(
          'Opening balance',
          val(this.view, st => st.opening[b]),
          '',
          this.view === 'diff' ? unByRow(b, 'opening') : undefined,
        );
      for (const k of keys) {
        const s = bucketOf(k)?.sign ?? 1;
        h += line(
          `${b === 'fee' ? '' : s === 1 ? 'Add: ' : 'Less: '}${bucketLabel(k)}`,
          val(this.view, st => st.rows[k]),
          k.endsWith(UNMAPPED) ? 'warn-error' : '',
          this.view === 'diff' ? unByRow(b, k) : undefined,
        );
      }
      h += line(
        b === 'fee' ? 'Total' : 'Closing balance',
        val(this.view, st => st.closing[b]),
        'tot',
        this.view === 'diff'
          ? r.items
              .filter(i => !i.explained && i.block === b)
              .reduce(
                (s, i) =>
                  s +
                  HEAD_KEYS.reduce((t, hh) => t + i.amt[hh] * (i.bucket === 'opening' ? 1 : (bucketOf(i.bucket)?.sign ?? 1)), 0),
                0,
              )
          : undefined,
      );
      return h;
    };
    const tot = (st: Result['flowBooks'], b: Block) => HEAD_KEYS.reduce((s, h) => s + st.closing[b][h], 0);
    const unex = r.items.filter(i => !i.explained);
    const unTot = unex.reduce((s, i) => s + HEAD_KEYS.reduce((t, h) => t + Math.abs(i.amt[h]), 0), 0);
    const kpi = (label: string, b: Block) => {
      const bb = tot(r.flowBooks, b),
        pp = tot(r.flowPortal, b);
      return `<div class="kpi ${Math.abs(bb - pp) < 1 ? 'ok' : 'warn'}"><span>${label}: books / portal</span><b>${fmt(bb)}</b><em>portal ${fmt(pp)} · difference ${fmt(bb - pp)}</em></div>`;
    };
    $('flFlow').innerHTML =
      `<div class="rv-bar"><div class="row" style="gap:10px"><label>From${sel('flFrom', inp.from)}</label><label>To${sel('flTo', inp.to)}</label></div>
        <div class="seg">${(['books', 'portal', 'diff'] as const).map(v => `<button data-view="${v}" class="${this.view === v ? 'on' : ''}">${v === 'books' ? 'As per books' : v === 'portal' ? 'As per portal' : 'Difference'}</button>`).join('')}</div></div>
      <div class="kpis">${kpi('ITC closing', 'itc')}${kpi('Liability closing', 'liab')}${kpi('Cash ledger closing', 'cash')}<div class="kpi ${unex.length ? 'bad' : 'ok'}"><span>Unexplained</span><b>${unex.length}</b><em>${unex.length ? `item(s), ${fmt(unTot)} in total` : 'every difference is explained'}</em></div></div>
      <div class="tablewrap"><table class="fl"><tr><th>${this.view === 'books' ? 'Flow as per books' : this.view === 'portal' ? 'Flow as per GST portal' : 'Difference: books − portal'}</th>${heads.map(h => `<th class="r">${HK_TXT[h]}</th>`).join('')}<th class="r">Total</th>${this.view === 'diff' ? '<th class="r">Of which unexplained</th>' : ''}</tr>${BLOCKS.map(tbl).join('')}</table></div>
      ${useAll ? '<p class="muted" style="margin:8px 0 0">Books accounts that are not split by tax head (for example a single GST cash ledger account) are in "Not split", while the portal shows each head: compare the Total column for those rows.</p>' : ''}
      <p class="muted" style="margin:8px 0 0">${this.view === 'books' ? 'From the books ledgers only, by voucher date.' : this.view === 'portal' ? 'From the credit ledger, cash ledger and liability register only, by return period.' : 'The closing difference is the sum of the row differences. The Reconciliation below explains it item by item.'}
        ${r.warns.length ? ` <span class="warn-warn">${r.warns.length} message(s) under Cross-checks.</span>` : ''}</p>`;
    const box = $('flFlow');
    box.querySelectorAll('button[data-view]').forEach(
      (el: any) =>
        (el.onclick = () => {
          this.view = el.dataset.view;
          this.renderFlow();
        }),
    );
    ($('flFrom') as any).onchange = (e: any) => {
      this.st.from = e.target.value;
      this.refresh();
    };
    ($('flTo') as any).onchange = (e: any) => {
      this.st.to = e.target.value;
      this.refresh();
    };
  }

  // ---------- 7. reconciliation: set-up, year comparison, ITC three ways, journals, closing
  private renderReasons() {
    const r = this.res,
      y = this.year;
    if (!r || !y) {
      $('flReasons').innerHTML = '';
      return;
    }
    const byId = new Map(r.books.map(e => [e.id, e]));
    const months = monthsOf(this.inp!.fy),
      nextMonths = monthsOf(`${+this.inp!.fy.slice(0, 4) + 1}-${String(+this.inp!.fy.slice(2, 4) + 2).padStart(2, '0')}`).slice(
        0,
        3,
      ),
      prev = monthsOf(`${+this.inp!.fy.slice(0, 4) - 1}-${this.inp!.fy.slice(2, 4)}`).slice(-3);
    const periodOpts = [...prev, ...months, ...nextMonths];
    const bookRow = (e: BEntry) => {
      const ov = this.st.ov[e.id],
        editing = this.editing === e.id,
        l = e.line;
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
    const keyOf = (blk: Block, b: string) => (b === UNMAPPED || b === INTERNAL ? `${blk}:${b}` : b);
    const lines = (c: Comp, m: Ym) => {
      const bl = r.books.filter(
        e => e.where === 'in' && e.block && e.month === m && c.buckets.includes(keyOf(e.block, e.bucket)),
      );
      const pl = r.portal.filter(
        p => p.where === 'in' && p.block && p.month === m && c.buckets.includes(keyOf(p.block, p.bucket)),
      );
      return `<div class="item-in">${bl.length ? `<p class="grp">Books lines</p><div class="tablewrap"><table class="fl"><tr><th>Date</th><th>Voucher</th><th>Party / particulars</th><th>Account</th><th>Head</th><th class="r">Amount</th><th>Flow row</th><th></th></tr>${bl.slice(0, 200).map(bookRow).join('')}</table></div>${bl.length > 200 ? `<p class="muted">First 200 of ${bl.length}; all of them are in Books_Lines.</p>` : ''}` : '<p class="muted">No books lines in this month.</p>'}
        ${
          pl.length
            ? `<p class="grp">Portal lines</p><div class="tablewrap"><table class="fl"><tr><th>Ledger</th><th>Date</th><th>Reference</th><th>Description</th><th>Head</th><th>Part</th><th class="r">Amount</th></tr>${pl
                .slice(0, 200)
                .map(
                  p =>
                    `<tr><td>${PORTAL_LABEL[p.kind]}</td><td>${esc(p.date)}</td><td>${esc(p.ref)}</td><td>${esc(p.desc)}${p.used && p.used !== '-' ? ` (${esc(p.used)})` : ''}</td><td>${p.head}</td><td>${p.comp}</td><td class="r">${fmt(p.amount)}</td></tr>`,
                )
                .join('')}</table></div>`
            : '<p class="muted">No portal lines in this month.</p>'
        }</div>`;
    };
    const T = (a: Record<HeadK, number>) => HEAD_KEYS.reduce((s, h) => s + a[h], 0);
    const btn = (k: string, show: string, hide: string) =>
      `<button class="link" data-tg="${esc(k)}">${this.open.has(k) ? hide : show}</button>`;
    // 1. set-up
    const lvl = {
      fix: ['Fix first', 'var(--err)', 'var(--err-bg)'],
      check: ['Check', 'var(--warn)', 'var(--warn-bg)'],
      note: ['Note', 'var(--muted)', 'var(--bg)'],
    } as const;
    const setup = y.setup.length
      ? `<ul class="warns" style="margin:0">${y.setup.map(s => `<li><span class="tag" style="color:${lvl[s.level][1]};background:${lvl[s.level][2]}">${lvl[s.level][0]}</span> <b>${esc(s.what)}</b><div class="muted" style="font-size:12.5px">${esc(s.action)}</div></li>`).join('')}</ul>`
      : '<p class="ok-text" style="margin:0">All three portal ledgers are loaded, the openings are in, and every books line is in a flow row.</p>';
    // 2. year comparison
    const yearRows = (blk: Block) =>
      y.comps
        .filter(c => c.block === blk)
        .map(c => {
          const d = T(c.diff),
            cut = T(c.cutoff),
            rest = T(c.rest),
            k = 'c:' + c.id,
            has = Math.abs(d) > 0.5;
          const kindTxt =
            c.kind === 'journal'
              ? 'Journal below'
              : c.kind === 'itc'
                ? 'ITC split below'
                : c.kind === 'setup'
                  ? 'Set-up'
                  : 'Invoice level';
          let h = `<tr class="${Math.abs(rest) > 0.5 && c.kind !== 'journal' ? 'unx' : ''}"><td><b>${esc(c.label)}</b>${has ? `<div class="muted" style="font-size:12.5px">${esc(c.action)}</div>` : ''}${btn(k, 'By month', 'Hide months')}</td>
        <td class="r">${fmt(T(c.books))}</td><td class="r">${fmt(T(c.portal))}</td><td class="r ${has ? 'neg' : ''}">${fmt(d)}</td><td class="r">${fmt(cut)}</td><td class="r"><b>${fmt(rest)}</b></td><td>${has ? esc(kindTxt) : '<span class="ok-text">Agrees</span>'}</td></tr>`;
          if (this.open.has(k)) {
            const heads = HEAD_KEYS.filter(hh => [c.books, c.portal, c.diff].some(a => Math.abs(a[hh]) > 0.004));
            h += `<tr><td colspan="7"><div class="item-in">${
              heads.length > 1
                ? `<p class="grp">By tax head</p><div class="tablewrap"><table class="fl"><tr><th></th>${heads.map(hh => `<th class="r">${HK_TXT[hh]}</th>`).join('')}</tr>
          ${(
            [
              ['Books', c.books],
              ['Portal', c.portal],
              ['Difference', c.diff],
              ['Of which cut-off', c.cutoff],
              ['To look at', c.rest],
            ] as const
          )
            .map(([n, a]) => `<tr><td>${n}</td>${heads.map(hh => `<td class="r">${fmt(a[hh])}</td>`).join('')}</tr>`)
            .join('')}</table></div>`
                : ''
            }
          <p class="grp">By month</p><div class="tablewrap"><table class="fl"><tr><th>Month</th><th class="r">Books</th><th class="r">Portal</th><th class="r">Difference</th><th class="r">Running difference</th><th></th></tr>
          ${c.months
            .map(x => {
              const mk = `m:${c.id}|${x.m}`;
              return `<tr class="${x.any ? 'unx' : ''}"><td>${ymLabel(x.m)}</td><td class="r">${fmt(x.b)}</td><td class="r">${fmt(x.p)}</td><td class="r ${x.any ? 'neg' : ''}">${fmt(x.d)}</td><td class="r">${fmt(x.run)}</td><td>${btn(mk, 'Show lines', 'Hide lines')}</td></tr>${this.open.has(mk) ? `<tr><td colspan="6">${lines(c, x.m)}</td></tr>` : ''}`;
            })
            .join('')}</table></div>
          <p class="muted" style="font-size:12.5px">A month in red differs in at least one tax head. The running difference shows where it starts and whether it clears later (a timing difference).</p></div></td></tr>`;
          }
          return h;
        })
        .join('');
    const yearTbl = (['itc', 'liab', 'cash', 'fee'] as Block[])
      .filter(b => y.comps.some(c => c.block === b))
      .map(
        b =>
          `<p class="grp">${esc(BLOCK_NAME[b])}</p><div class="tablewrap"><table class="fl rec"><tr><th></th><th class="r">Books</th><th class="r">Portal</th><th class="r">Difference</th><th class="r">Of which year-end cut-off</th><th class="r">To look at</th><th>Where</th></tr>${yearRows(b)}</table></div>`,
      )
      .join('');
    // 3. ITC three ways
    const it = this.itc;
    const h4 = (a: Record<string, number>) =>
      HEADS.map(h => `<td class="r">${fmt(a[h] || 0)}</td>`).join('') +
      `<td class="r"><b>${fmt(HEADS.reduce((s, h) => s + (a[h] || 0), 0))}</b></td>`;
    const bridge = (
      ls: BridgeLine[],
    ) => `<div class="tablewrap"><table class="fl rec"><tr><th></th>${HEADS.map(h => `<th class="r">${HK_TXT[h]}</th>`).join('')}<th class="r">Total</th></tr>
      ${ls
        .filter(l => l.kind === 'start' || l.kind === 'total' || HEADS.some(h => Math.abs(l.amt[h]) > 0.5))
        .map(
          l =>
            `<tr class="${l.kind === 'start' || l.kind === 'total' ? 'tot' : l.kind === 'check' ? 'unx' : ''}"><td>${l.kind === 'start' || l.kind === 'total' ? `<b>${esc(l.label)}</b>` : esc(l.label)}${l.action ? `<div class="muted" style="font-size:12.5px">${esc(l.action)}</div>` : ''}</td>${h4(l.amt)}</tr>`,
        )
        .join('')}</table></div>`;
    const itcHtml = !it
      ? `<p class="muted">Add the GSTR-2B JSON files of the period (the ones the Input Reco tab uses) to split the ITC difference invoice by invoice: what is in books but not in 2B, in 2B but not in books, and what was claimed in 3B.</p>`
      : `<p class="grp">Books to GSTR-2B (invoice by invoice)</p>${bridge(it.books)}
        ${
          it.invoices.length
            ? `<p style="margin:8px 0 4px">${btn('i3:inv', `Show the ${it.invoices.length} invoice(s) that differ`, 'Hide the invoices')}</p>${
                this.open.has('i3:inv')
                  ? `<div class="tablewrap" style="max-height:480px;overflow:auto"><table class="fl"><tr><th>Status</th><th>Supplier</th><th>GSTIN</th><th>Books no.</th><th>2B no.</th><th>Books date</th><th>2B period</th><th class="r">Tax in books</th><th class="r">Tax in 2B</th><th>What to do</th></tr>
          ${it.invoices
            .slice(0, 300)
            .map(
              v =>
                `<tr><td>${esc(v.status)}</td><td>${esc(v.supplier)}</td><td>${esc(v.gstin)}</td><td>${esc(v.bNo)}</td><td>${esc(v.rNo)}</td><td>${esc(v.bDate)}</td><td>${esc(v.period)}</td><td class="r">${fmt(v.bTax)}</td><td class="r">${fmt(v.rTax)}</td><td style="font-size:12.5px">${esc(v.note)}</td></tr>`,
            )
            .join(
              '',
            )}</table></div>${it.invoices.length > 300 ? `<p class="muted">First 300; all of them are in ITC_Invoices.</p>` : ''}`
                  : ''
              }`
            : ''
        }
        ${it.nonInvoice.length ? `<p style="margin:8px 0 4px">${btn('i3:ni', `Show the ${it.nonInvoice.length} ITC entries without a purchase invoice`, 'Hide them')}</p>${this.open.has('i3:ni') ? `<div class="tablewrap"><table class="fl"><tr><th>Date</th><th>Voucher</th><th>Account</th><th class="r">ITC</th></tr>${it.nonInvoice.map(n => `<tr><td>${esc(n.date)}</td><td>${esc(n.voucher)}</td><td>${esc(n.account)}</td><td class="r">${fmt(n.tax)}</td></tr>`).join('')}</table></div>` : ''}` : ''}
        <p class="grp">GSTR-2B to GSTR-3B (month by month)</p>${bridge(it.claim)}
        <div class="tablewrap"><table class="fl"><tr><th>Month</th><th class="r">ITC booked</th><th class="r">Available in 2B</th><th class="r">Claimed in 3B</th><th class="r">2B − 3B</th><th class="r">Running</th></tr>
          ${it.months.map(m => `<tr class="${Math.abs(m.diff) > 0.5 ? 'unx' : ''}"><td>${ymLabel(m.m)}</td><td class="r">${fmt(m.booked)}</td><td class="r">${fmt(m.b2b)}</td><td class="r">${fmt(m.claimed)}</td><td class="r ${Math.abs(m.diff) > 0.5 ? 'neg' : ''}">${fmt(m.diff)}</td><td class="r">${fmt(m.run)}</td></tr>`).join('')}</table></div>
        ${it.notes.length ? `<ul class="warns" style="margin:6px 0 0">${it.notes.map(n => `<li class="warn-info">${esc(n)}</li>`).join('')}</ul>` : ''}
        ${it.register ? `<p class="${Math.abs(it.register.register - it.register.books) > 1 ? 'warn-warn' : 'ok-text'}" style="margin:8px 0 0">Purchase register cross-check: tax in the register ${fmt(it.register.register)}, on the purchase invoices in the GL ${fmt(it.register.books)}${Math.abs(it.register.register - it.register.books) > 1 ? ` (differs by ${fmt(it.register.register - it.register.books)}: check the register's period and filters)` : ' (they agree)'}.</p>` : ''}`;
    // 4. journals
    const jHtml = y.journals.length
      ? y.journals
          .map(
            j => `<p class="grp">${esc(j.no)} · ${esc(j.date)} · ${esc(j.title)}</p><div class="tablewrap"><table class="fl"><tr><th>Account</th><th>Head</th><th class="r">Debit</th><th class="r">Credit</th></tr>
        ${j.legs.map(l => `<tr><td>${esc(l.account)}${l.existing ? '' : ' <span class="muted" style="font-size:11.5px">(not in your books: pick the account)</span>'}</td><td>${HK_TXT[l.head]}</td><td class="r">${fmt(l.dr)}</td><td class="r">${fmt(l.cr)}</td></tr>`).join('')}
        <tr class="tot"><td>Total</td><td></td><td class="r">${fmt(j.legs.reduce((s, l) => s + l.dr, 0))}</td><td class="r">${fmt(j.legs.reduce((s, l) => s + l.cr, 0))}</td></tr></table></div><p class="muted" style="font-size:12.5px;margin:4px 0 0">${esc(j.narration)}</p>`,
          )
          .join('')
      : '<p class="ok-text" style="margin:0">None needed: set-off, payments and deposits agree with the portal.</p>';
    // 5. closing
    const clo = y.closing
      .map(c => {
        const hs = HEAD_KEYS.filter(h =>
          [c.books, c.portal, c.journals, ...c.lines.map(l => l.amt)].some(a => Math.abs(a[h]) > 0.004),
        );
        const H = hs.length ? hs : (['IGST', 'CGST', 'SGST'] as HeadK[]);
        const am = (a: Record<HeadK, number>, cls = '') =>
          H.map(h => `<td class="r ${cls}">${fmt(a[h])}</td>`).join('') +
          `<td class="r ${cls}"><b>${fmt(H.reduce((s, h) => s + a[h], 0))}</b></td>`;
        const neg = (a: Record<HeadK, number>) => Object.fromEntries(HEAD_KEYS.map(h => [h, -a[h]])) as Record<HeadK, number>;
        return `<p class="grp">${esc(BLOCK_NAME[c.block])}</p><div class="tablewrap"><table class="fl rec"><tr><th></th>${H.map(h => `<th class="r">${HK_TXT[h]}</th>`).join('')}<th class="r">Total</th></tr>
        <tr class="tot"><td>Closing as per books</td>${am(c.books)}</tr>
        ${HEAD_KEYS.some(h => Math.abs(c.journals[h]) > 0.004) ? `<tr><td>Add: suggested journals</td>${am(c.journals)}</tr><tr class="tot"><td>Adjusted closing as per books</td>${am(c.adjusted)}</tr>` : ''}
        ${c.lines.map(l => `<tr class="${l.kind === 'unexplained' ? 'unx' : ''}"><td>${l.kind === 'unexplained' ? '<span class="tag" style="color:var(--err);background:var(--err-bg)">Not explained</span> ' : ''}${esc(l.label)}<div class="muted" style="font-size:12.5px">${esc(l.action)}</div></td>${am(neg(l.amt))}</tr>`).join('')}
        <tr class="tot"><td>Closing as per portal</td>${am(c.portal)}</tr></table></div>`;
      })
      .join('');
    const sec = (n: string, t: string, sub: string, body: string) =>
      `<h3 style="margin:22px 0 2px;font-size:16px">${n}. ${t}</h3><p class="muted" style="margin:0 0 8px">${sub}</p>${body}`;
    $('flReasons').innerHTML =
      sec(
        '1',
        'Set-up checks',
        'Fix these first: a missing ledger, opening or flow row shows up as a difference everywhere below.',
        setup,
      ) +
      sec(
        '2',
        'Year comparison',
        'Each line compares the year as per books with the portal. The year-end cut-off clears in the next period; "To look at" is the rest. Open a line for the month where it arises.',
        yearTbl,
      ) +
      sec(
        '3',
        'ITC: books, GSTR-2B and GSTR-3B',
        'The ITC difference split three ways. Books to 2B is matched invoice by invoice from the purchase vouchers in the books; 2B to 3B is by month, because 3B is a total.',
        itcHtml,
      ) +
      sec(
        '4',
        'Suggested journals',
        'Entries that would bring set-off, payments and deposits in the books in line with the portal. Nothing is posted: review them and post them in your books.',
        jHtml,
      ) +
      sec(
        '5',
        'Closing statement',
        'Books closing plus the journals, against the portal closing. What is left is explained line by line.',
        clo,
      ) +
      (this.st.audit.length
        ? `<details class="rv-hist" style="margin-top:14px"><summary class="eyebrow" style="cursor:pointer">Your changes (${this.st.audit.length})</summary><ol>${this.st.audit
            .slice()
            .reverse()
            .map(
              a =>
                `<li>${esc(a.when)} · <b>${esc(a.change)}</b> · ${esc(a.record)}: ${esc(a.from)} → ${esc(a.to)}${a.remark ? ` · <i>${esc(a.remark)}</i>` : ''}</li>`,
            )
            .join('')}</ol></details>`
        : '');
    const box = $('flReasons');
    box.querySelectorAll('button[data-tg]').forEach(
      (el: any) =>
        (el.onclick = (ev: Event) => {
          ev.preventDefault();
          const k = el.dataset.tg;
          if (this.open.has(k)) this.open.delete(k);
          else this.open.add(k);
          this.renderReasons();
        }),
    );
    box.querySelectorAll('button[data-edit]').forEach(
      (el: any) =>
        (el.onclick = (ev: Event) => {
          ev.preventDefault();
          this.editing = this.editing === el.dataset.edit ? '' : el.dataset.edit;
          this.renderReasons();
        }),
    );
    box.querySelectorAll('button[data-reset]').forEach(
      (el: any) =>
        (el.onclick = () => {
          const e = byId.get(el.dataset.reset)!;
          const old = this.st.ov[e.id];
          delete this.st.ov[e.id];
          this.editing = '';
          this.log('Line change undone', recTxt(e), describeOv(old), 'tool classification', 'Undone in the Loader');
          this.refresh();
        }),
    );
    box.querySelectorAll('button[data-save]').forEach(
      (el: any) =>
        (el.onclick = () => {
          const id = el.dataset.save,
            e = byId.get(id)!,
            f = box.querySelector(`.edit[data-for="${CSS.escape(id)}"]`) as HTMLElement;
          const g = (k: string) => (f.querySelector(`[data-k="${k}"]`) as any).value as string;
          const remark = g('remark').trim();
          if (!remark) {
            (f.querySelector('[data-k="remark"]') as any).focus();
            alertBox(f, 'A remark is needed: say why this line belongs elsewhere.');
            return;
          }
          const ov: LineOverride = {};
          if (g('bucket') !== e.auto) ov.bucket = g('bucket');
          const accHead = this.inp!.accMap[e.line.account]?.head || 'ALL';
          if (g('head') !== accHead) ov.head = g('head') as AccHead;
          if (g('period') !== ymOfDate(e.line.date)) ov.period = g('period');
          const before = `${bucketLabel(e.bucket)} · ${e.head} · ${ymLabel(e.month)}`;
          if (Object.keys(ov).length) this.st.ov[id] = ov;
          else delete this.st.ov[id];
          this.log(
            'Line changed',
            recTxt(e),
            before,
            `${bucketLabel(g('bucket'))} · ${g('head')} · ${ymLabel(g('period'))}`,
            remark,
          );
          this.editing = '';
          this.refresh();
        }),
    );
  }

  // ---------- 6. checks
  private renderChecks() {
    const r = this.res;
    if (!r) {
      $('flChecks').innerHTML = '<p class="muted">Load at least two portal ledgers.</p>';
      return;
    }
    const bad = r.checks.filter(c => !c.ok),
      ok = r.checks.length - bad.length;
    const rows = [...bad, ...r.checks.filter(c => c.ok)]
      .slice(0, 400)
      .map(
        c =>
          `<tr><td>${esc(c.area)}</td><td>${ymLabel(c.period)}</td><td>${esc(c.head)}</td><td class="r">${fmt(c.a)}</td><td class="r">${fmt(c.b)}</td><td class="r ${c.ok ? '' : 'neg'}">${fmt(c.diff)}</td><td>${c.ok ? '<span class="ok-text">OK</span>' : '<span class="warn-error">Check</span>'}${c.note ? `<div class="muted" style="font-size:11.5px">${esc(c.note)}</div>` : ''}</td></tr>`,
      )
      .join('');
    $('flChecks').innerHTML =
      `${r.warns.length ? `<ul class="warns" style="margin:0 0 12px">${r.warns.map(w => `<li class="warn-${w.level}">${esc(w.msg)}</li>`).join('')}</ul>` : ''}
      ${
        r.checks.length
          ? `<p class="${bad.length ? 'warn-warn' : 'ok-text'}" style="margin:0 0 8px">${ok} of ${r.checks.length} checks agree${bad.length ? `; ${bad.length} to look at` : ''}.</p>
      <div class="tablewrap" style="max-height:480px;overflow:auto"><table class="fl"><tr><th>Check</th><th>Return period</th><th>Head</th><th class="r">A</th><th class="r">B</th><th class="r">A − B</th><th></th></tr>${rows}</table></div>
      <p class="muted" style="margin:8px 0 0">A and B: cash ledger debits vs liability paid in cash; credit ledger debits vs liability paid by ITC; credit ledger ITC vs GSTR-3B 4A − 4B; liability register vs GSTR-3B 6.1 (where the 3B PDFs are loaded).</p>`
          : '<p class="muted">No cross-checks yet: they need two of the portal ledgers, or a portal ledger and GSTR-3B PDFs.</p>'
      }`;
  }

  // ---------- 7. workbook
  private async build() {
    const r = this.res,
      c = this.ctx();
    if (!r || !this.inp || !this.year) return;
    $('flStatus').textContent = 'Building…';
    try {
      const bytes = await flowWorkbook(r, {
        gstin: c.gstin,
        legalName: c.legal,
        fy: c.fy,
        from: this.inp.from,
        to: this.inp.to,
        builtAt: nowTxt(),
        audit: this.st.audit,
        year: this.year,
        itc: this.itc,
        settings: c.settings ? describe('flow', c.settings()) : undefined,
        files: c.files
          .filter(f => ['Portal ledger', 'Books ledger', 'GSTR-3B'].includes(f.kind))
          .map(f => ({ name: f.name, kind: f.kind, note: f.note || '' })),
      });
      const full = this.inp.from === monthsOf(c.fy)[0] && this.inp.to === monthsOf(c.fy)[11];
      const name = `GST_Flow_${c.gstin || 'client'}_${c.fy}${full ? '' : `_${ymLabel(this.inp.from)}_to_${ymLabel(this.inp.to)}`}.xlsx`;
      const url = URL.createObjectURL(
        new Blob([bytes as any], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      );
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      a.click();
      const unx = this.year.closing.flatMap(x => x.lines).filter(l => l.kind === 'unexplained').length;
      $('flResult').innerHTML =
        `<p><b>Downloaded ${esc(name)}</b> · <a href="${url}" download="${esc(name)}">download again</a>.</p>
        <p class="muted">Year_Compare, the flow sheets and Closing are formulas over Books_Lines, Portal_Lines and Balances. ${this.year.journals.length} suggested journal(s)${unx ? `, ${unx} closing line(s) not explained` : ''}. ${this.st.audit.length} change(s) of yours in Audit.</p>`;
      $('flStatus').textContent = '';
    } catch (e: any) {
      $('flStatus').innerHTML = `<span class="warn-error">${esc(e.message)}</span>`;
    }
  }
}

const recTxt = (e: BEntry) =>
  `${e.line.file} row ${e.line.row} · ${e.line.date} · ${e.line.vchType} ${e.line.vchNo} · ${e.line.account}`;
const describeOv = (o?: LineOverride) =>
  o ? [o.bucket && bucketLabel(o.bucket), o.head, o.period && ymLabel(o.period)].filter(Boolean).join(' · ') : '-';
function alertBox(el: HTMLElement, msg: string) {
  let p = el.querySelector('.flmsg') as HTMLElement | null;
  if (!p) {
    p = document.createElement('p');
    p.className = 'flmsg warn-error';
    p.style.margin = '6px 0 0';
    el.appendChild(p);
  }
  p.textContent = msg;
}

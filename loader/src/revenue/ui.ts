// Revenue reco tab: books source (sales register, else revenue ledgers from a GL dump), dashboard, the two statements,
// every entry with filters and reasons (one by one or in bulk), and the workbook.
import type { InFile } from '../build';
import { addStateSheets, packState } from '../state';
import { withProgress } from '../progress';
import { processGstr1 } from '../gstr1';
import type { MatchOptions } from '../recon';
import { Reader } from '../xlsx';
import { readRows } from '../integrity';
import { RawRow, Warn, esc, fyShift, inr, r2, store } from '../util';
import {
  Item,
  REASON,
  REASONS,
  ReasonCode,
  RevResult,
  StLine,
  UserState,
  Ym,
  emptyUser,
  monthsOf,
  reasonOfLabel,
  runRevenue,
  ymLabel,
} from './engine';
import { GlResult, REV_CLASSES, REV_CLASS_LABEL, RevClass, glDocs } from './ledgers';
import { ITEMS_SHEET, RevAudit, readItemsSheet, revenueWorkbook } from './workbook';
import { REV_GUIDE } from './guide';
import { nameKey } from '../suppliers';
import {
  CUST_HOW,
  CustMapping,
  CustSuggestion,
  G1Customer,
  applyCustomerAliases,
  customerSuggestions,
  customersFromG1,
  resolveCustomers,
} from './customers';
import { CustSettings, DEFAULT_CUST, describe, defaults } from '../settings';
import { DEFAULT_MATCH } from '../recon';

const $ = (id: string) => document.getElementById(id)! as any;
const fmt = (v: number) => (Math.abs(v) < 0.005 ? '–' : inr(v));
const nowTxt = () => {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
};
const PICKABLE = REASONS.filter(r => r.code !== 'MATCHED');

export interface RevCtx {
  files: InFile[];
  fy: string;
  gstin: string;
  legal: string;
  active: boolean;
  /** sales register rows mapped in "Map register columns" (undefined = none, or not mapped yet) */
  salesRows: () => any[][] | undefined;
  match?: () => MatchOptions | undefined;
  /** how customers without a GSTIN get one (Matching settings) */
  cust?: () => CustSettings | undefined;
}
interface State {
  user: UserState;
  audit: RevAudit[];
  imported: string[];
}
interface Filters {
  q: string;
  party: string;
  type: string;
  reason: string;
  month: string;
  from: string;
  to: string;
  diffOnly: boolean;
}
const NOF = (): Filters => ({ q: '', party: '', type: '', reason: '', month: '', from: '', to: '', diffOnly: true });

export class RevenueTab {
  private key = '';
  private st: State = { user: emptyUser(), audit: [], imported: [] };
  private cls: Record<string, RevClass> = {};
  /** customer name (nameKey) → GSTIN confirmed by the user ('' = unregistered / B2C) */
  private cust: Record<string, string> = {};
  private custSug: CustSuggestion[] = [];
  private listCount = 0;
  /** customer name (nameKey) → GSTIN in use and the basis */
  private custMap = new Map<string, CustMapping>();
  private custDir = new Map<string, G1Customer>();
  private custPick = new Map<string, string>();
  private showAllLedgers = false;
  private res: RevResult | null = null;
  private gl: GlResult | null = null;
  private g1raw: RawRow[] = [];
  /** GSTR-1 files left out (the same return loaded twice) */
  private g1warns: Warn[] = [];
  private bookRows: any[][] = [];
  private source: 'register' | 'gl' | 'none' = 'none';
  private f: Filters = NOF();
  private sel = new Set<string>();
  private limit = 200;
  private side: 0 | 1 = 0;
  private g1Key = '';
  constructor(private ctx: () => RevCtx) {
    $('rvBuild').onclick = () => this.build();
  }

  /** read the browser settings again (a workpaper was restored, or the restore undone) */
  reload() {
    this.key = '';
  }
  private load(c: RevCtx) {
    const k = `${c.gstin || '-'}:${c.fy}`;
    if (k === this.key) return;
    this.key = k;
    this.st = { user: emptyUser(), audit: [], imported: [], ...(store.get(`jjrev:${k}`) || {}) };
    this.st.user = { ...emptyUser(), ...this.st.user };
    this.cls = store.get(`jjrev:cls:${c.gstin || '-'}`) || {};
    this.cust = store.get(`jjrev:cust:${c.gstin || '-'}`) || {};
    this.f = NOF();
    this.sel.clear();
  }
  private save() {
    const c = this.ctx();
    store.set(`jjrev:${this.key}`, this.st);
    store.set(`jjrev:cls:${c.gstin || '-'}`, this.cls);
    store.set(`jjrev:cust:${c.gstin || '-'}`, this.cust);
  }
  private log(change: string, record: string, from: string, to: string, remark = '') {
    this.st.audit.push({ when: nowTxt(), change, record, from, to, remark });
  }

  coverageHtml(): string {
    const c = this.ctx();
    const fy = c.fy,
      ny = fyShift(fy, 1);
    const ms = monthsOf(fy);
    const g1 = c.files.filter(f => f.kind === 'GSTR-1');
    const per = (f: InFile) => (f.period ? `${f.period.slice(2)}-${f.period.slice(0, 2)}` : '');
    const covered = (f: InFile): string[] => {
      const p = per(f);
      if (!p) return [];
      if (f.json?.filing_typ === 'Q') {
        const y = +p.slice(0, 4),
          m = +p.slice(5);
        return [m - 2, m - 1, m].map(x => `${y}-${String(x).padStart(2, '0')}`);
      }
      return [p];
    };
    const g1s = new Set(g1.flatMap(covered));
    const g3 = new Set(c.files.filter(f => f.kind === 'GSTR-3B' && f.threeB).map(f => per({ period: f.threeB!.period } as any)));
    const nyMs = monthsOf(ny).slice(0, 7);
    const row = (label: string, sub: string, list: string[], set: Set<string>, opt = false) =>
      `<tr><th class="cvl">${label}<span>${sub}</span></th>${list.map(m => (set.has(m) ? '<td class="cv ok">✓</td>' : opt ? '<td class="cv opt">–</td>' : '<td class="cv miss">✕</td>')).join('')}${list.length < 12 ? `<td class="cv na" colspan="${12 - list.length}"></td>` : ''}<td class="cvc ${list.every(m => set.has(m)) ? 'full' : opt ? '' : 'part'}">${list.filter(m => set.has(m)).length}/${list.length}</td></tr>`;
    const regs = c.files.filter(f => f.kind === 'Register').length,
      books = c.files.filter(f => f.kind === 'Books ledger').length,
      cl = c.files.filter(f => f.kind === 'Customer list').reduce((n, f) => n + (f.customers?.length || 0), 0);
    return `<table class="cov"><tr><th class="cvl">Return</th>${['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'].map(m => `<th class="cvh">${m}</th>`).join('')}<th class="cvh">Loaded</th></tr>
      ${row('GSTR-1', fy, ms, g1s)}${row('GSTR-3B', fy, ms, g3)}${row('GSTR-1', `${ny} · Apr–Oct, optional`, nyMs, g1s, true)}
      <tr><th class="cvl">Books</th><td colspan="13" class="cv ${regs || books ? 'ok' : 'miss'} wide">${regs ? `✓ sales register: ${regs} file(s)` : books ? `✓ day book / GL dump: ${books} file(s) (no sales register)` : '✕ sales register (preferred), or a day book / GL dump'}</td></tr>
      <tr><th class="cvl">Customer list<span>optional</span></th><td colspan="13" class="cv ${cl ? 'ok' : 'opt'} wide">${cl ? `✓ ${cl} customers with GSTIN` : '– name + GSTIN of customers: helps when the books have no GSTIN'}</td></tr></table>`;
  }

  refresh() {
    const c = this.ctx();
    if (!c.active) return;
    this.load(c);
    // GSTR-1 rows (any year); processGstr1 is heavy, so keep it until the files change
    const g1f = c.files.filter(f => f.kind === 'GSTR-1' && f.json);
    const gk = g1f.map(f => f.hash || f.name).join('|') + '|' + c.fy;
    if (gk !== this.g1Key) {
      this.g1Key = gk;
      try {
        this.g1raw = g1f.length
          ? processGstr1(
              g1f.map(f => ({ name: f.name, json: f.json })),
              c.fy,
              fyShift(c.fy, 1),
              (this.g1warns = []),
            ).rows
          : [];
      } catch (e: any) {
        this.g1raw = [];
        $('rvDash').innerHTML = `<p class="warn-error">GSTR-1: ${esc(e.message)}</p>`;
      }
    }
    // books: sales register first, else revenue ledgers from the GL dump / ledgers
    let reg: any[][] | undefined;
    try {
      reg = c.salesRows();
    } catch {
      reg = undefined;
    }
    const bookFiles = c.files.filter(f => f.kind === 'Books ledger' && f.books).map(f => f.books!);
    this.gl = bookFiles.length ? glDocs(bookFiles, this.cls, c.fy) : null;
    if (reg && reg.length) {
      this.source = 'register';
      this.bookRows = reg;
    } else if (this.gl) {
      this.source = 'gl';
      this.bookRows = this.gl.docs.map(d => d.row);
    } else {
      this.source = 'none';
      this.bookRows = [];
    }
    // customers without a GSTIN: by the methods chosen under Matching settings, each with its basis; the user's choice wins
    const cs = c.cust?.() ?? DEFAULT_CUST,
      mo = c.match?.() ?? DEFAULT_MATCH;
    const list = cs.list ? c.files.filter(f => f.kind === 'Customer list').flatMap(f => f.customers || []) : [];
    this.custDir = customersFromG1(this.g1raw, this.bookRows, list, mo.tol);
    this.custSug =
      this.bookRows.length && this.custDir.size
        ? customerSuggestions(this.bookRows, this.custDir, { minScore: cs.nameMin / 100, tol: mo.tol, compare: mo.compare })
        : [];
    // the customer list: a name with one GSTIN there is applied as is (also customers that are not in GSTR-1 at all)
    const byName = new Map<string, Set<string>>();
    for (const x of list) {
      const k = nameKey(x.name);
      if (k) (byName.get(k) ?? byName.set(k, new Set()).get(k)!).add(x.gstin);
    }
    const listAliases = new Map<string, string>();
    for (const [k, gs] of byName) if (gs.size === 1) listAliases.set(k, [...gs][0]);
    this.listCount = list.length;
    this.custMap = resolveCustomers(this.custSug, cs, listAliases, this.cust);
    const aliases = new Map([...this.custMap].map(([k, m]) => [k, m.gstin] as [string, string]));
    if (aliases.size) this.bookRows = applyCustomerAliases(this.bookRows, aliases);
    this.run();
    this.importWorkbooks().then(changed => {
      if (changed) {
        this.run();
        this.render();
      }
    });
    this.render();
  }

  private run() {
    const c = this.ctx();
    const threeB = c.files.filter(f => f.kind === 'GSTR-3B' && f.threeB).map(f => f.threeB!);
    try {
      this.res =
        this.bookRows.length || this.g1raw.length
          ? runRevenue({
              fy: c.fy,
              books: { rows: this.bookRows, source: this.source },
              g1: this.g1raw,
              threeB,
              user: this.st.user,
              match: c.match?.(),
            })
          : null;
    } catch (e: any) {
      this.res = null;
      $('rvDash').innerHTML = `<p class="warn-error">${esc(e.message)}</p>`;
    }
    this.save();
  }

  /** a saved revenue workbook dropped with the files: read the user's columns back (once per file) */
  private async importWorkbooks(): Promise<boolean> {
    const c = this.ctx();
    const wbs = c.files.filter(f => f.kind === 'Revenue workbook' && f.bytes && !this.st.imported.includes(f.hash || f.name));
    if (!wbs.length || !this.res) return false;
    let n = 0;
    for (const f of wbs) {
      try {
        const rd = await Reader.open(f.bytes!);
        const rows = await readRows(rd, ITEMS_SHEET, 1);
        const byId = new Map(this.res.items.map(i => [i.id, i]));
        const lab2ym = new Map(monthsOf(c.fy).map(m => [ymLabel(m).toLowerCase(), m]));
        for (const x of readItemsSheet(rows)) {
          const it = byId.get(x.id);
          const code = reasonOfLabel(x.reason);
          const cur = this.st.user.reasons[x.id];
          if (code && code !== 'MATCHED') {
            if (cur?.r !== code || (cur?.note ?? '') !== x.remark) {
              this.st.user.reasons[x.id] = { r: code, note: x.remark || undefined };
              n++;
            }
          } else if (x.remark && cur?.note !== x.remark && it) {
            this.st.user.reasons[x.id] = { r: cur?.r, note: x.remark };
            n++;
          }
          if (it && it.hasB) {
            const ym = x.b2c ? (lab2ym.get(x.b2c.toLowerCase()) ?? '') : '';
            if (ym !== it.b2cMonth) {
              this.st.user.b2c[it.bIds.join(',')] = ym;
              n++;
            }
          }
        }
        this.log('Read back from workbook', f.name, '', `${n} change(s)`);
      } catch (e: any) {
        this.log('Workbook not read', f.name, '', e.message);
      }
      this.st.imported.push(f.hash || f.name);
    }
    this.save();
    return true;
  }

  private render() {
    this.renderBooks();
    this.renderCustomers();
    this.renderDash();
    this.renderStatements();
    this.renderItems();
    $('rvBuild').disabled = !this.res;
  }

  private guide(src: ('portal' | 'tally' | 'zoho' | 'other')[], title: string) {
    const items = REV_GUIDE.filter(g => src.includes(g.source));
    return `<details class="guide"><summary><span class="gq">?</span> ${title}</summary><div class="gpane on"><div class="gitems">${items.map(g => `<div class="gitem"><h4>${esc(g.title)} <span class="gfor">${g.source === 'portal' ? 'GST portal' : g.source === 'tally' ? 'Tally' : g.source === 'zoho' ? 'Zoho Books' : 'Other'}</span></h4><ol>${g.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ol>${g.note ? `<p class="muted">${esc(g.note)}</p>` : ''}</div>`).join('')}</div></div></details>`;
  }

  // ---------- books source
  private renderBooks() {
    const c = this.ctx();
    const regs = c.files.filter(f => f.kind === 'Register');
    let html = '';
    if (this.source === 'register') {
      const t = this.bookRows.reduce((a, r) => a + (String(r[6]) === 'CN' ? -1 : 1) * Math.abs(+r[11] || 0), 0);
      html += `<p><span class="tag ok-tag">Sales register</span> ${regs.map(f => esc(f.name)).join(', ')} · ${this.bookRows.length} lines · taxable ${fmt(t)}. Column mapping is under "Map register columns".</p>`;
      if (this.gl) {
        const glT = this.gl.docs
          .filter(d => d.cls !== 'nosupply')
          .reduce((a, d) => a + (d.row[6] === 'CN' ? -1 : 1) * d.row[11], 0);
        html += `<p class="muted">The day book / GL dump is loaded too: it is used only as a check. Revenue ledgers ticked below total ${fmt(r2(glT))} against ${fmt(r2(t))} in the register (difference ${fmt(r2(glT - t))}).</p>`;
      }
    } else if (this.source === 'gl')
      html += `<p><span class="tag">Day book / GL dump</span> No sales register is loaded, so invoices are built from the vouchers that touch the revenue ledgers below, with tax from the Output GST ledgers of the same voucher. Customer GSTINs are not in a GL dump: invoices not found in GSTR-1 B2B are taken as B2C of their month.</p>`;
    else if (regs.length)
      html +=
        '<p class="warn-warn">Map the required register columns (date, invoice no., taxable value) under "Map register columns".</p>';
    else
      html +=
        '<p class="muted">Add the sales register (preferred), or a day book / GL dump if there is no proper sales register.</p>';
    if (this.gl) {
      const ls = this.gl.ledgers.filter(l => this.showAllLedgers || l.guess || this.cls[l.name]);
      const opts = (cur: string) =>
        `<option value="">— review —</option>` +
        REV_CLASSES.map(k => `<option value="${k}" ${k === cur ? 'selected' : ''}>${esc(REV_CLASS_LABEL[k])}</option>`).join('');
      html += `<p class="grp">Revenue ledgers</p><p class="muted" style="margin-top:0">Ledgers that look like revenue are suggested. Say what each one holds; "Not revenue" leaves the ledger out entirely. To leave out single entries, use the reason "Not a supply" in the list below.</p>
        <div class="tablewrap"><table class="fl sup"><tr><th>Ledger</th><th class="r">Lines</th><th class="r">Debit</th><th class="r">Credit</th><th class="r">Net credit</th><th>What it holds</th></tr>
        ${ls
          .map(l => {
            const cur = this.cls[l.name] ?? '';
            return `<tr class="${cur ? 'confirmed' : ''}"><td><b>${esc(l.name)}</b>${!cur && l.guess ? `<div class="muted small">suggested: ${esc(REV_CLASS_LABEL[l.guess])}</div>` : ''}</td><td class="r">${l.lines}</td><td class="r">${fmt(l.dr)}</td><td class="r">${fmt(l.cr)}</td><td class="r">${fmt(r2(l.cr - l.dr))}</td>
            <td><select data-rcls="${esc(l.name)}">${opts(cur || l.guess || '')}</select>${!cur ? ` <button class="small" data-rok="${esc(l.name)}">Confirm</button>` : ''}</td></tr>`;
          })
          .join('')}</table></div>
        <div class="row" style="justify-content:space-between;margin:8px 0"><span class="muted">${this.gl.outsideFy ? `${this.gl.outsideFy} voucher(s) outside ${esc(c.fy)} left out. ` : ''}${this.gl.taxOnly ? `${this.gl.taxOnly} voucher(s) post output GST without a revenue ledger (e.g. tax journals): not turnover. ` : ''}</span>
        <span>${ls.some(l => !this.cls[l.name]) ? '<button class="small" id="rvClsAll">Confirm all suggestions</button> ' : ''}<button class="link" id="rvClsShow">${this.showAllLedgers ? 'show suggested ledgers only' : 'show all ledgers, to add one'}</button></span></div>`;
      if (this.showAllLedgers) {
        // every account in the dump, so the user can add a revenue ledger the guess missed
        const all = new Map<string, { dr: number; cr: number; n: number }>();
        for (const f of c.files.filter(x => x.kind === 'Books ledger' && x.books))
          for (const l of f.books!.lines) {
            if (this.gl.ledgers.some(x => x.name === l.account)) continue;
            const a = all.get(l.account) ?? { dr: 0, cr: 0, n: 0 };
            a.dr += l.dr;
            a.cr += l.cr;
            a.n++;
            all.set(l.account, a);
          }
        html += `<details open><summary class="muted">Other accounts (${all.size})</summary><div class="tablewrap" style="max-height:320px;overflow:auto"><table class="fl"><tr><th>Account</th><th class="r">Lines</th><th class="r">Net credit</th><th></th></tr>${[
          ...all,
        ]
          .sort((a, b) => b[1].cr - b[1].dr - (a[1].cr - a[1].dr))
          .map(
            ([n, a]) =>
              `<tr><td>${esc(n)}</td><td class="r">${a.n}</td><td class="r">${fmt(r2(a.cr - a.dr))}</td><td><button class="link" data-radd="${esc(n)}">add as revenue</button></td></tr>`,
          )
          .join('')}</table></div></details>`;
      }
    }
    html += this.guide(['tally', 'zoho', 'other'], 'How to export the sales register, day book or ledgers');
    const box = $('rvBooks');
    box.innerHTML = html;
    const setCls = (name: string, v: RevClass) => {
      const old = this.cls[name];
      this.cls[name] = v;
      this.log('Ledger class', name, old ? REV_CLASS_LABEL[old] : 'suggested', REV_CLASS_LABEL[v]);
    };
    box.querySelectorAll('select[data-rcls]').forEach(
      (el: any) =>
        (el.onchange = () => {
          if (!el.value) return;
          setCls(el.dataset.rcls, el.value);
          this.refresh();
        }),
    );
    box.querySelectorAll('button[data-rok]').forEach(
      (el: any) =>
        (el.onclick = () => {
          const l = this.gl!.ledgers.find(x => x.name === el.dataset.rok);
          const sel = box.querySelector(`select[data-rcls="${CSS.escape(el.dataset.rok)}"]`) as HTMLSelectElement;
          const v = (sel?.value || l?.guess) as RevClass;
          if (v) setCls(el.dataset.rok, v);
          this.refresh();
        }),
    );
    box.querySelectorAll('button[data-radd]').forEach(
      (el: any) =>
        (el.onclick = () => {
          setCls(el.dataset.radd, 'taxable');
          this.refresh();
        }),
    );
    const all = box.querySelector('#rvClsAll');
    if (all)
      all.onclick = () => {
        for (const l of this.gl!.ledgers) if (!this.cls[l.name] && l.guess) setCls(l.name, l.guess);
        this.refresh();
      };
    const show = box.querySelector('#rvClsShow');
    if (show)
      show.onclick = () => {
        this.showAllLedgers = !this.showAllLedgers;
        this.renderBooks();
      };
  }

  // ---------- customer names (no GSTIN in the books)
  private renderCustomers() {
    const card = $('rv-cust'),
      box = $('rvCust');
    card.style.display = this.custSug.length ? '' : 'none';
    if (!this.custSug.length) return;
    const cs = this.ctx().cust?.() ?? DEFAULT_CUST;
    const pending = this.custSug.filter(x => !this.custMap.has(x.key));
    const done = this.custSug.filter(x => this.custMap.has(x.key));
    const all = [...this.custDir.values()].sort((a, b) => b.tx - a.tx);
    const label = (g: string) => {
      const c = this.custDir.get(g);
      const n = c ? [...c.names][0] : '';
      return `${g}${n ? ' · ' + n : ''}${c ? ` · ${c.docs} doc(s) in GSTR-1` : ''}`;
    };
    const pickOf = (x: CustSuggestion) => {
      const p = this.custPick.get(x.key) ?? x.inv?.gstin ?? x.cands[0]?.gstin ?? '';
      return p;
    };
    const strong = pending.filter(x => x.inv);
    const noCand = pending.filter(x => !x.cands.length);
    const why = (c: { invHits: number; noHits: number; score: number }) =>
      [
        c.invHits ? `${c.invHits} invoice(s): no. + amount agree` : '',
        c.noHits ? `${c.noHits} no. only (amount differs)` : '',
        `name ${Math.round(c.score * 100)}%`,
      ]
        .filter(Boolean)
        .join(' · ');
    const rows = [...pending]
      .sort((a, b) => +!!b.inv - +!!a.inv || +!!b.cands.length - +!!a.cands.length)
      .map(x => {
        const cur = pickOf(x);
        const candG = new Set(x.cands.map(c => c.gstin));
        const top = x.cands[0];
        const basis = x.inv
          ? `<span class="tag ok-tag">invoice no. + amount</span> <span class="muted small">${x.inv.hits} of ${x.inv.of} invoice(s) agree with ${esc(x.inv.gstin)}: ${esc(x.inv.sample.join(', '))}${!cs.invAmt ? ' · rule off in settings' : !cs.invAuto ? ' · set to suggest only' : x.inv.hits < cs.invMin ? ` · settings need ${cs.invMin}` : ''}</span>`
          : top
            ? `<span class="muted small">${esc(why(top))}${top.noHits && !top.invHits ? ' · amounts differ, so not mapped' : ''}</span>`
            : '<span class="muted small">No GSTR-1 customer by name, invoice no. or amount</span>';
        return `<tr><td><b>${esc(x.name)}</b><div class="muted small">${x.docs} document(s) · taxable ${fmt(x.tx)}</div></td>
        <td><select data-cpick="${esc(x.key)}" style="max-width:520px">
          ${x.cands.length ? `<optgroup label="Suggested">${x.cands.map(c => `<option value="${c.gstin}" ${c.gstin === cur ? 'selected' : ''}>${esc(label(c.gstin))} · ${esc(why(c))}</option>`).join('')}</optgroup>` : ''}
          <optgroup label="Any customer in GSTR-1">${all
            .filter(c => !candG.has(c.gstin))
            .map(c => `<option value="${c.gstin}" ${c.gstin === cur ? 'selected' : ''}>${esc(label(c.gstin))}</option>`)
            .join('')}</optgroup>
          <option value="-" ${cur === '-' ? 'selected' : ''}>Unregistered customer (B2C)</option>
          ${!cur ? '<option value="" selected>— pick —</option>' : ''}</select><div style="margin-top:4px">${basis}</div></td>
        <td class="act"><button class="small" data-cok="${esc(x.key)}">Confirm</button></td></tr>`;
      })
      .join('');
    const howTag = (m: CustMapping) => `<span class="tag ${m.how === 'user' ? '' : 'ok-tag'}">${esc(CUST_HOW[m.how])}</span>`;
    const doneRows = done
      .map(x => {
        const m = this.custMap.get(x.key)!;
        return `<tr><td><b>${esc(x.name)}</b><div class="muted small">${x.docs} document(s) · taxable ${fmt(x.tx)}</div></td>
        <td>${m.gstin ? esc(label(m.gstin)) : 'Unregistered (B2C)'}</td><td>${howTag(m)}<div class="muted small">${esc(m.detail)}</div></td>
        <td class="act">${m.how === 'user' ? `<button class="link" data-cundo="${esc(x.key)}">undo</button>` : `<button class="link" data-cask="${esc(x.key)}">not right? choose</button>`}</td></tr>`;
      })
      .join('');
    const methods = describe('revenue', { ...defaults('revenue'), cust: cs })[1];
    box.innerHTML =
      `<p class="muted" style="margin-top:0">GSTR-1 has the customer GSTIN but no name, so a books customer without a GSTIN gets one from the customer list, the same name elsewhere in the books, or invoices whose number <b>and</b> amount agree with GSTR-1 documents of one GSTIN. <b>In use:</b> ${esc(methods.replace(/^Customer GSTIN by: /, ''))} (change it under Matching settings). A mapped customer's entries are matched invoice by invoice instead of being taken as B2C. ${this.listCount ? `Customer list loaded: ${this.listCount} name(s) with GSTIN.` : 'Tip: drop a customer list (just customer name and GSTIN, from Tally or Zoho) and names are matched from it.'}</p>` +
      (pending.length
        ? `<div class="row" style="justify-content:space-between;margin:6px 0"><span class="muted">${pending.length} customer(s) without a GSTIN to confirm.</span><span>${strong.length ? `<button class="small" id="rvCustStrong">Confirm ${strong.length} where invoice no. + amount agree</button> ` : ''}${noCand.length ? `<button class="small ghost" id="rvCustB2c">Mark ${noCand.length} with no match as unregistered (B2C)</button>` : ''}</span></div>
          <div class="tablewrap"><table class="fl sup"><tr><th>Customer in books</th><th>GSTIN in GSTR-1 and why it is suggested</th><th></th></tr>${rows}</table></div>`
        : '<p class="ok-text">All customers without a GSTIN are mapped.</p>') +
      (done.length
        ? `<details style="margin-top:10px" ${pending.length ? '' : 'open'}><summary class="muted">Mapped customers and the basis of each (${done.length})</summary>
          <div class="tablewrap"><table class="fl sup"><tr><th>Customer in books</th><th>GSTIN given</th><th>Basis</th><th></th></tr>${doneRows}</table></div></details>`
        : '');
    box
      .querySelectorAll('select[data-cpick]')
      .forEach((el: any) => (el.onchange = () => this.custPick.set(el.dataset.cpick, el.value)));
    const confirm = (x: CustSuggestion, v: string) => {
      if (!v) return;
      const g = v === '-' ? '' : v;
      const from = this.custMap.get(x.key);
      this.cust[x.key] = g;
      this.log(
        'Customer GSTIN',
        x.name,
        from ? `${from.gstin || 'B2C'} (${CUST_HOW[from.how]})` : '(none)',
        g ? label(g) : 'Unregistered (B2C)',
      );
    };
    box.querySelectorAll('button[data-cok]').forEach(
      (el: any) =>
        (el.onclick = () => {
          const x = pending.find(p => p.key === el.dataset.cok);
          if (x) confirm(x, pickOf(x));
          this.refresh();
        }),
    );
    box.querySelectorAll('button[data-cundo]').forEach(
      (el: any) =>
        (el.onclick = () => {
          const x = done.find(p => p.key === el.dataset.cundo);
          delete this.cust[el.dataset.cundo];
          if (x) this.log('Customer GSTIN', x.name, 'confirmed', 'undone');
          this.refresh();
        }),
    );
    box.querySelectorAll('button[data-cask]').forEach(
      (el: any) =>
        (el.onclick = () => {
          const x = done.find(p => p.key === el.dataset.cask);
          const m = this.custMap.get(el.dataset.cask);
          this.cust[el.dataset.cask] = '?';
          if (x && m) this.log('Customer GSTIN', x.name, `${m.gstin || 'B2C'} (${CUST_HOW[m.how]})`, 'to be chosen by you');
          this.refresh();
        }),
    );
    const nb = box.querySelector('#rvCustB2c');
    if (nb)
      nb.onclick = () => {
        for (const x of noCand) confirm(x, '-');
        this.refresh();
      };
    const st = box.querySelector('#rvCustStrong');
    if (st)
      st.onclick = () => {
        for (const x of strong) confirm(x, x.inv!.gstin);
        this.refresh();
      };
  }

  // ---------- dashboard
  private renderDash() {
    const r = this.res;
    const box = $('rvDash');
    if (!r) {
      box.innerHTML = '<p class="muted">Add GSTR-1 JSONs and the books to see the reconciliation.</p>';
      return;
    }
    const s = this.side;
    const unx1 = r.st1.find(l => l.code === 'unx')!.amt[s],
      unx2 = r.st2.find(l => l.code === 'unx')!.amt[s];
    const lab = s ? 'tax' : 'taxable value';
    const kp = (label: string, v: number, sub: string, cls = '') =>
      `<div class="kpi ${cls}"><span>${label}</span><b>${fmt(v)}</b><em>${sub}</em></div>`;
    const d1 = r2(r.g1Turnover[s] - r.books.turnover[s]),
      d2 = r2(r.tbTotal[s] - r.g1Turnover[s]);
    const open = r.items.filter(
      i =>
        ['NOT_REPORTED', 'NOT_IN_BOOKS', 'VALUE'].includes(i.reason) &&
        !i.userReason &&
        Math.abs(i.g[0] - i.b[0]) + Math.abs(i.g[1] - i.b[1]) > r.tol,
    ).length;
    const mine = r.items.filter(i => i.userReason).length;
    const kpis = `<div class="kpis">${kp('Books', r.books.turnover[s], `turnover, ${lab}`)}${kp('GSTR-1', r.g1Turnover[s], `${lab}`)}${kp('GSTR-3B', r.tbTotal[s], `Table 3.1, ${r.months.filter(m => m.tbt).length} month(s)`)}
      ${kp('GSTR-1 − books', d1, `${r.st1.filter(l => l.kind === 'line').length} reason(s)`, Math.abs(d1) > r.tol ? 'warn' : 'ok')}${r.months.some(m => m.tbt) ? kp('3B − GSTR-1', d2, 'see GSTR-1 → 3B', Math.abs(d2) > r.tol ? 'warn' : 'ok') : kp('3B − GSTR-1', 0, 'no GSTR-3B loaded', 'tot')}
      ${kp('Not explained', r2(unx1 + unx2), `${open} entr${open === 1 ? 'y' : 'ies'} to look at · ${mine} with your reason`, Math.abs(unx1 + unx2) > r.tol ? 'bad' : 'ok')}</div>`;
    // grouped bars per month: books, GSTR-1, 3B
    const ser = [
      { name: 'Books', color: '#4552B0', v: (m: (typeof r.months)[number]) => m.books[s] },
      { name: 'GSTR-1', color: '#1F9E89', v: (m: (typeof r.months)[number]) => m.g1[s] },
      { name: 'GSTR-3B', color: '#E08A1E', v: (m: (typeof r.months)[number]) => (m.tb ? m.tb[s] : 0) },
    ];
    const W = 760,
      H = 220,
      padL = 64,
      padB = 26,
      padT = 10;
    const vals = r.months.flatMap(m => ser.map(x => x.v(m)));
    const max = Math.max(1, ...vals),
      min = Math.min(0, ...vals);
    const y = (v: number) => padT + ((max - v) / (max - min)) * (H - padT - padB);
    const gw = (W - padL - 8) / 12,
      bw = Math.min(16, (gw - 10) / 3);
    const ticks = [0, 0.25, 0.5, 0.75, 1].map(t => min + (max - min) * t);
    const short = (v: number) =>
      Math.abs(v) >= 1e7
        ? `${r2(v / 1e7)} Cr`
        : Math.abs(v) >= 1e5
          ? `${r2(v / 1e5)} L`
          : inr(Math.round(v)).replace(/\.00$/, '');
    const bars = r.months
      .map((m, i) => {
        const x0 = padL + i * gw + (gw - 3 * bw - 4) / 2;
        return (
          ser
            .map((sr, k) => {
              const v = sr.v(m),
                top = y(Math.max(v, 0)),
                h = Math.max(0, Math.abs(y(v) - y(0)));
              return `<rect x="${x0 + k * (bw + 2)}" y="${top}" width="${bw}" height="${h}" rx="2" fill="${sr.color}"><title>${ymLabel(m.ym)} · ${sr.name}: ${inr(v)}${sr.name === 'GSTR-3B' && !m.tb ? ' (3B not loaded)' : ''}</title></rect>`;
            })
            .join('') + `<text x="${padL + i * gw + gw / 2}" y="${H - 8}" text-anchor="middle" class="ax">${ymLabel(m.ym)}</text>`
        );
      })
      .join('');
    const chart = `<div class="rvchart"><div class="legend">${ser.map(x => `<span><i style="background:${x.color}"></i>${x.name}</span>`).join('')}<span class="muted">by month, ${lab} (books by invoice date, returns by return period)</span></div>
      <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Monthly ${lab}: books, GSTR-1 and GSTR-3B">${ticks.map(t => `<line x1="${padL}" x2="${W - 4}" y1="${y(t)}" y2="${y(t)}" class="gl"/><text x="${padL - 6}" y="${y(t) + 4}" text-anchor="end" class="ax">${short(t)}</text>`).join('')}<line x1="${padL}" x2="${W - 4}" y1="${y(0)}" y2="${y(0)}" class="zl"/>${bars}</svg></div>`;
    let run1 = 0,
      run2 = 0;
    const table = `<div class="tablewrap"><table class="fl"><tr><th>Month</th><th class="r">Books</th><th class="r">GSTR-1</th><th class="r">GSTR-3B</th><th class="r">GSTR-1 − books</th><th class="r">Running</th><th class="r">3B − GSTR-1</th><th class="r">Running</th><th class="r">B2CS vs books B2C</th></tr>
      ${r.months
        .map(m => {
          const a = r2(m.g1[s] - m.books[s]),
            b = m.tb ? r2(m.tb[s] - m.g1[s]) : 0;
          run1 = r2(run1 + a);
          run2 = r2(run2 + b);
          const b2c = r2(m.b2cG1[s] - m.b2cBooks[s]);
          return `<tr><td>${ymLabel(m.ym)}</td><td class="r">${fmt(m.books[s])}</td><td class="r">${fmt(m.g1[s])}</td><td class="r">${m.tb ? fmt(m.tb[s]) : '<span class="muted">not loaded</span>'}</td><td class="r ${Math.abs(a) > r.tol ? 'neg' : ''}">${fmt(a)}</td><td class="r muted">${fmt(run1)}</td><td class="r ${Math.abs(b) > r.tol ? 'neg' : ''}">${m.tb ? fmt(b) : ''}</td><td class="r muted">${fmt(run2)}</td><td class="r ${Math.abs(b2c) > r.tol ? 'neg' : ''}">${fmt(b2c)}</td></tr>`;
        })
        .join('')}
      <tr class="tot"><td>Year</td><td class="r">${fmt(r.books.turnover[s])}</td><td class="r">${fmt(r.g1Turnover[s])}</td><td class="r">${fmt(r.tbTotal[s])}</td><td class="r">${fmt(d1)}</td><td></td><td class="r">${fmt(r2(r.months.reduce((x, m) => x + (m.tb ? m.tb[s] - m.g1[s] : 0), 0)))}</td><td></td><td class="r">${fmt(r2(r.months.reduce((x, m) => x + m.b2cG1[s] - m.b2cBooks[s], 0)))}</td></tr></table></div>
      <p class="muted" style="margin:6px 0 0">A difference that reverses in a later month of the year is timing: it shows here but not in the statements below${r.timing.count ? ` (${r.timing.count} invoice(s) booked in one month and reported in another, taxable ${fmt(r.timing.amt[0])})` : ''}.</p>`;
    const checks = r.checks.length
      ? `<ul class="warns">${r.checks.map(x => `<li class="warn-${x.level}">${esc(x.msg)}</li>`).join('')}${r.notes.map(n => `<li class="warn-warn">${esc(n)}</li>`).join('')}</ul>`
      : '';
    box.innerHTML =
      this.g1warns
        .filter(w => w.level === 'error')
        .map(w => `<p class="warn-error"><b>${esc(w.file)}:</b> ${esc(w.msg)}</p>`)
        .join('') +
      checks +
      `<div class="rv-bar"><div class="seg"><button class="${s ? '' : 'on'}" data-side="0">Taxable value</button><button class="${s ? 'on' : ''}" data-side="1">Tax</button></div></div>` +
      kpis +
      chart +
      table;
    box.querySelectorAll('button[data-side]').forEach(
      (b: any) =>
        (b.onclick = () => {
          this.side = +b.dataset.side as 0 | 1;
          this.renderDash();
        }),
    );
  }

  // ---------- statements
  private stTable(lines: StLine[], clickable: boolean) {
    return `<div class="tablewrap"><table class="fl rec"><tr><th></th><th class="r">Taxable value</th><th class="r">Tax</th><th class="r">Entries</th></tr>${lines
      .map(l => {
        const strong = l.kind === 'start' || l.kind === 'sub' || l.kind === 'end';
        const unx = l.kind === 'unexplained';
        const bad = unx && Math.abs(l.amt[0]) + Math.abs(l.amt[1]) > 0.5;
        return `<tr class="${strong ? 'tot' : ''} ${bad ? 'unx' : ''}"><td>${strong ? `<b>${esc(l.label)}</b>` : esc(l.label)}${l.action && !strong ? `<div class="muted small">${esc(l.action)}</div>` : ''}${l.hint ? `<div class="muted small">${esc(l.hint)}</div>` : ''}${clickable && l.reasons ? `<button class="link" data-show="${l.reasons.join(',')}">show entries</button>` : ''}</td>
        <td class="r">${unx && !bad ? '<span class="ok-text">nil</span>' : fmt(l.amt[0])}</td><td class="r">${unx && !bad ? '' : fmt(l.amt[1])}</td><td class="r">${l.count || ''}</td></tr>`;
      })
      .join('')}</table></div>`;
  }
  private renderStatements() {
    const r = this.res;
    if (!r) {
      $('rvSt1').innerHTML = $('rvSt2').innerHTML = '';
      return;
    }
    $('rvSt1').innerHTML =
      this.stTable(r.st1, true) +
      '<p class="muted" style="margin:6px 0 0">Add / less is the change needed to go from the books figure to the GSTR-1 figure. Reasons come from the matching; where you picked a reason for an entry, yours is used.</p>';
    $('rvSt2').innerHTML =
      this.stTable(r.st2, false) +
      `<p class="muted" style="margin:6px 0 0">Year totals of the months with a 3B loaded. A month's shortfall made good in a later month of the year cancels out here (see the month table above).</p>` +
      this.guide(['portal'], 'How to download GSTR-1 and GSTR-3B');
    $('rvSt1')
      .querySelectorAll('button[data-show]')
      .forEach(
        (b: any) =>
          (b.onclick = () => {
            this.f = { ...NOF(), reason: b.dataset.show.split(',')[0], diffOnly: true };
            this.sel.clear();
            this.renderItems();
            $('rv-items').scrollIntoView({ behavior: 'smooth', block: 'start' });
          }),
      );
  }

  // ---------- entries with reasons
  private filtered(): Item[] {
    const r = this.res!,
      f = this.f;
    const q = f.q.trim().toLowerCase();
    const dn = (d: string) => (d ? d.slice(6) + d.slice(3, 5) + d.slice(0, 2) : '');
    return r.items.filter(i => {
      if (f.diffOnly && Math.abs(i.g[0] - i.b[0]) < 0.005 && Math.abs(i.g[1] - i.b[1]) < 0.005 && !i.userReason) return false;
      if (f.reason && i.reason !== f.reason) return false;
      if (f.party && (i.party || '(no name)') !== f.party) return false;
      if (f.type && i.type !== f.type) return false;
      if (f.month && i.bYm !== f.month && i.gYm !== f.month) return false;
      if (f.from && (!i.bDate || dn(i.bDate) < f.from.replace(/-/g, ''))) return false;
      if (f.to && (!i.bDate || dn(i.bDate) > f.to.replace(/-/g, ''))) return false;
      if (q && ![i.party, i.gstin, i.bNo, i.gNo, i.remark].some(x => String(x).toLowerCase().includes(q))) return false;
      return true;
    });
  }
  private setReason(list: Item[], code: ReasonCode | '', note?: string) {
    for (const it of list) {
      const old = it.userReason ? REASON[it.userReason].label : `(tool) ${REASON[it.derived].label}`;
      const keep = note ?? this.st.user.reasons[it.id]?.note;
      if (code) this.st.user.reasons[it.id] = { r: code, note: keep };
      else if (keep) this.st.user.reasons[it.id] = { note: keep };
      else delete this.st.user.reasons[it.id];
      // reasons saved under a document id of the item (from an earlier pairing) give way too
      for (const p of it.id.split(/[~,]/)) if (p && p !== it.id && this.st.user.reasons[p]) delete this.st.user.reasons[p];
      this.log('Reason', `${it.bNo || it.gNo} ${it.party}`.trim(), old, code ? REASON[code].label : '(tool)', note ?? '');
    }
  }
  private setB2c(list: Item[], ym: Ym | '') {
    for (const it of list) {
      if (!it.hasB) continue;
      this.st.user.b2c[it.bIds.join(',')] = ym;
      this.log('B2CS month', `${it.bNo} ${it.party}`.trim(), ymLabel(it.b2cMonth) || '-', ymLabel(ym) || 'not in B2CS');
    }
  }
  private rerun() {
    this.run();
    this.render();
  }

  private renderItems() {
    const box = $('rvItems');
    const r = this.res;
    if (!r) {
      box.innerHTML = '';
      return;
    }
    const f = this.f;
    const l = this.filtered();
    const ms = monthsOf(r.fy);
    const parties = [...new Set(r.items.map(i => i.party || '(no name)'))].sort();
    const types = [...new Set(r.items.map(i => i.type).filter(Boolean))].sort();
    const reasons = REASONS.filter(x => r.items.some(i => i.reason === x.code));
    const opt = (v: string, label: string, cur: string) =>
      `<option value="${esc(v)}" ${v === cur ? 'selected' : ''}>${esc(label)}</option>`;
    const reasonSel = (cur: string, attr: string, blank = '— tool —') =>
      `<select ${attr}><option value="">${blank}</option>${PICKABLE.map(x => opt(x.code, x.label, cur)).join('')}</select>`;
    const monthSel = (cur: string, attr: string) =>
      `<select ${attr}><option value="">not in B2CS</option>${ms.map(m => opt(m, ymLabel(m), cur)).join('')}</select>`;
    const tot = l.reduce((a, i) => [a[0] + i.b[0], a[1] + i.g[0], a[2] + i.b[1], a[3] + i.g[1]], [0, 0, 0, 0]).map(r2);
    const selItems = r.items.filter(i => this.sel.has(i.id));
    const filters = `<div class="rv-filters"><div class="fchips">${reasons
      .map(
        x =>
          `<button class="fchip ${f.reason === x.code ? 'on' : ''}" data-fr="${x.code}">${esc(x.label)} <span>${r.items.filter(i => i.reason === x.code && (!f.diffOnly || Math.abs(i.g[0] - i.b[0]) + Math.abs(i.g[1] - i.b[1]) > 0.004 || i.userReason)).length}</span></button>`,
      )
      .join('')}</div>
      <div class="row">
        <label class="grow">Search<input data-f="q" value="${esc(f.q)}" placeholder="customer, GSTIN, invoice no. or remark"/></label>
        <label>Customer<select data-f="party"><option value="">All</option>${parties.map(p => opt(p, p, f.party)).join('')}</select></label>
        <label>Type<select data-f="type"><option value="">All</option>${types.map(t => opt(t, t, f.type)).join('')}</select></label>
        <label>Month<select data-f="month"><option value="">All</option>${ms.map(m => opt(m, ymLabel(m), f.month)).join('')}</select></label>
        <label>Invoice date from<input type="date" data-f="from" value="${f.from}" style="min-width:0"/></label>
        <label>to<input type="date" data-f="to" value="${f.to}" style="min-width:0"/></label>
        <label class="tg"><input type="checkbox" data-f="diffOnly" ${f.diffOnly ? 'checked' : ''}/> Differences only</label>
        <button class="link" id="rvClear">Clear filters</button></div></div>`;
    const foot = `<div class="rv-foot top"><span><b>${l.length}</b> entries shown</span><span>Taxable: books ${fmt(tot[0])} · GSTR-1 ${fmt(tot[1])} · diff <b>${fmt(r2(tot[1] - tot[0]))}</b></span><span>Tax: books ${fmt(tot[2])} · GSTR-1 ${fmt(tot[3])} · diff <b>${fmt(r2(tot[3] - tot[2]))}</b></span>${l.length ? `<button class="link" id="rvSelAll">select all ${l.length} shown</button>` : ''}</div>`;
    const bulk = selItems.length
      ? `<div class="rv-sel"><span><b>${selItems.length}</b> selected · taxable diff ${fmt(r2(selItems.reduce((a, i) => a + i.g[0] - i.b[0], 0)))}</span>
        <span class="acts">${reasonSel('', 'id="rvBulkR"', 'Set reason…')}<input id="rvBulkN" placeholder="remark (optional)" style="min-width:0;width:180px;padding:6px 10px;font-size:13px"/><button class="small" id="rvBulkApply">Apply</button>
        ${monthSel('', 'id="rvBulkM"').replace('<option value="">not in B2CS</option>', '<option value="-">B2CS month…</option><option value="">not in B2CS</option>')}<button class="small ghost" id="rvBulkMApply">Set month</button>
        <button class="small ghost" id="rvBulkClear">Back to tool's reason</button><button class="link" id="rvSelClear">Clear</button></span></div>`
      : '';
    const rows = l
      .slice(0, this.limit)
      .map(i => {
        const dT = r2(i.g[0] - i.b[0]),
          dX = r2(i.g[1] - i.b[1]);
        return `<tr class="${this.sel.has(i.id) ? 'sel' : ''}"><td class="cb"><input type="checkbox" data-sel="${esc(i.id)}" ${this.sel.has(i.id) ? 'checked' : ''}/></td>
        <td>${reasonSel(i.userReason, `data-ir="${esc(i.id)}"`, `— ${esc(REASON[i.derived].label)} —`)}${i.userReason ? '' : i.reason !== i.derived ? `<div class="muted small">${esc(REASON[i.reason].label)}</div>` : ''}${i.note ? `<div class="muted small">${esc(i.note)}</div>` : ''}</td>
        <td>${esc(i.status)}<div class="muted small">${esc(i.type)}${i.dt && i.dt !== 'INV' ? ' · ' + esc(i.dt) : ''}${i.basis ? ' · ' + esc(i.basis) : ''}</div></td>
        <td>${esc(i.party || '—')}<div class="muted small">${esc(i.gstin)}${this.custMap.get(nameKey(i.party))?.gstin ? ` · GSTIN by ${esc(CUST_HOW[this.custMap.get(nameKey(i.party))!.how].toLowerCase())}` : ''}</div></td>
        <td>${esc(i.bNo)}<div class="muted small">${esc(i.bDate)}</div></td>
        <td>${esc(i.gNo)}<div class="muted small">${esc(i.gSection)} ${ymLabel(i.gYm)}${i.hasG && !i.g[0] && i.gAny[0] ? ' · other year' : ''}</div></td>
        <td class="r">${i.hasB ? fmt(i.b[0]) : ''}</td><td class="r">${i.hasG ? fmt(i.gAny[0]) : ''}</td><td class="r ${Math.abs(dT) > r.tol ? 'neg' : ''}">${fmt(dT)}</td><td class="r ${Math.abs(dX) > r.tol ? 'neg' : ''}">${fmt(dX)}</td>
        <td>${i.hasB && i.reason !== 'EXCLUDED' ? monthSel(i.b2cMonth, `data-im="${esc(i.bIds.join(','))}"`) : i.kind === 'b2c' ? ymLabel(i.b2cMonth) : ''}</td>
        <td><input class="remark" data-in="${esc(i.id)}" value="${esc(i.remark)}" placeholder="remark" style="min-width:0;width:150px;padding:5px 8px;font-size:13px"/></td></tr>`;
      })
      .join('');
    box.innerHTML =
      filters +
      foot +
      bulk +
      `<div class="tablewrap rv-tw"><table class="rv"><thead><tr><th class="cb"><input type="checkbox" id="rvAllShown" title="Select all shown"/></th><th>Reason</th><th>Status</th><th>Customer</th><th>Books no.</th><th>GSTR-1 no.</th><th class="r">Books taxable</th><th class="r">GSTR-1 taxable</th><th class="r">Diff taxable</th><th class="r">Diff tax</th><th>B2CS month</th><th>Remark</th></tr></thead><tbody>${rows}</tbody></table></div>
      ${l.length > this.limit ? `<button class="link" id="rvMore">Show more (${l.length - this.limit} left)</button>` : ''}${!l.length ? '<p class="muted" style="margin:10px 0">Nothing matches these filters.</p>' : ''}
      <p class="muted" style="margin:8px 0 0">Diff = GSTR-1 − books. Pick a reason only where you know better than the tool; it moves the entry to that line of the statement. "B2CS month" is the month whose GSTR-1 B2CS includes a books entry (B2C entries start with their invoice month). Everything you change goes to the Audit sheet.</p>`;

    const on = (sel: string, ev: string, fn: (el: any) => void) =>
      box.querySelectorAll(sel).forEach((el: any) => (el[ev] = () => fn(el)));
    on('[data-f]', 'onchange', el => {
      (this.f as any)[el.dataset.f] = el.type === 'checkbox' ? el.checked : el.value;
      this.limit = 200;
      this.renderItems();
    });
    const q = box.querySelector('input[data-f="q"]');
    if (q)
      q.oninput = () => {
        this.f.q = q.value;
        clearTimeout((this as any)._t);
        (this as any)._t = setTimeout(() => {
          this.renderItems();
          const n = $('rvItems').querySelector('input[data-f="q"]');
          n?.focus();
          n?.setSelectionRange(n.value.length, n.value.length);
        }, 250);
      };
    on('button[data-fr]', 'onclick', el => {
      this.f.reason = this.f.reason === el.dataset.fr ? '' : el.dataset.fr;
      this.renderItems();
    });
    on('#rvClear', 'onclick', () => {
      this.f = NOF();
      this.renderItems();
    });
    on('#rvMore', 'onclick', () => {
      this.limit += 300;
      this.renderItems();
    });
    on('input[data-sel]', 'onchange', el => {
      if (el.checked) this.sel.add(el.dataset.sel);
      else this.sel.delete(el.dataset.sel);
      this.renderItems();
    });
    on('#rvAllShown', 'onchange', el => {
      for (const i of l.slice(0, this.limit)) el.checked ? this.sel.add(i.id) : this.sel.delete(i.id);
      this.renderItems();
    });
    on('#rvSelAll', 'onclick', () => {
      l.forEach(i => this.sel.add(i.id));
      this.renderItems();
    });
    on('#rvSelClear', 'onclick', () => {
      this.sel.clear();
      this.renderItems();
    });
    on('select[data-ir]', 'onchange', el => {
      const it = r.items.find(i => i.id === el.dataset.ir);
      if (it) this.setReason([it], el.value as ReasonCode | '');
      this.rerun();
    });
    on('select[data-im]', 'onchange', el => {
      const it = r.items.find(i => i.bIds.join(',') === el.dataset.im);
      if (it) this.setB2c([it], el.value);
      this.rerun();
    });
    on('input[data-in]', 'onchange', el => {
      const it = r.items.find(i => i.id === el.dataset.in);
      if (!it) return;
      const cur = this.st.user.reasons[it.id];
      this.log('Remark', `${it.bNo || it.gNo} ${it.party}`.trim(), cur?.note ?? '', el.value);
      if (cur) cur.note = el.value || undefined;
      else if (el.value) this.st.user.reasons[it.id] = { note: el.value };
      this.rerun();
    });
    on('#rvBulkApply', 'onclick', () => {
      const code = $('rvBulkR').value as ReasonCode;
      if (!code) return;
      this.setReason(selItems, code, $('rvBulkN').value || undefined);
      this.sel.clear();
      this.rerun();
    });
    on('#rvBulkMApply', 'onclick', () => {
      const v = $('rvBulkM').value;
      if (v === '-') return;
      this.setB2c(selItems, v);
      this.sel.clear();
      this.rerun();
    });
    on('#rvBulkClear', 'onclick', () => {
      this.setReason(selItems, '');
      this.sel.clear();
      this.rerun();
    });
  }

  // ---------- workbook
  private build() {
    return withProgress('Building the workbook…', () => this.buildInner());
  }
  private async buildInner() {
    const c = this.ctx();
    const r = this.res;
    if (!r) return;
    $('rvStatus').textContent = 'Building…';
    try {
      this.save();
      const book = await revenueWorkbook(r, {
        gstin: c.gstin,
        legalName: c.legal,
        fy: c.fy,
        source:
          this.source === 'register' ? 'Sales register' : this.source === 'gl' ? 'Day book / GL dump (revenue ledgers)' : 'None',
        files: c.files.filter(f => f.kind !== 'Unknown').map(f => ({ name: f.name, kind: f.kind, note: f.note || '' })),
        audit: this.st.audit,
        builtAt: nowTxt(),
        g1: this.g1raw,
        threeB: c.files.filter(f => f.kind === 'GSTR-3B' && f.threeB).map(f => f.threeB!),
        bookRows: this.bookRows,
        ledgers: this.source === 'gl' ? this.gl?.ledgers.map(l => ({ ...l, cls: this.cls[l.name] ?? null })) : undefined,
        settings: describe('revenue', {
          ...defaults('revenue'),
          match: c.match?.() ?? DEFAULT_MATCH,
          cust: c.cust?.() ?? DEFAULT_CUST,
        }),
        customers: this.custSug
          .filter(x => this.custMap.has(x.key))
          .map(x => {
            const m = this.custMap.get(x.key)!;
            return { name: x.name, gstin: m.gstin, basis: CUST_HOW[m.how], detail: m.detail };
          }),
      });
      const bytes = await addStateSheets(book, packState('revenue', c.gstin, c.fy, nowTxt()));
      const name = `Revenue_Reco_${(c.gstin || 'client').trim()}_${c.fy}.xlsx`;
      const url = URL.createObjectURL(
        new Blob([bytes as any], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      );
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      a.click();
      $('rvStatus').textContent = '';
      $('rvResult').innerHTML =
        `<p><b>Downloaded ${esc(name)}</b> · <a href="${url}" download="${esc(name)}">download again</a>.</p><p class="muted">To add reasons in Excel: fill the yellow columns of Rev_Items (Your reason, B2CS month, Remark), save, and drop the workbook here with the same files.</p>`;
    } catch (e: any) {
      $('rvStatus').innerHTML = `<span class="warn-error">${esc(e.message)}</span>`;
    }
  }
}

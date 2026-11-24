// Review step: run the reconciliation in the browser before building, let the user change how documents are matched,
// filter suppliers / invoices with totals, match or unmatch by hand, and download the filtered view.
import { Decisions, InFile, ReconExtra, reconPreview } from './build';
import { DEFAULT_MATCH, MatchOptions } from './recon';
import { describeInvoices } from './settings';
import { simpleXlsx } from './xlsxbook';
import { esc, inr, store } from './util';
import { meter, ranked, stackedBars, tile, short } from './dash';

type Side = 'purchases' | 'sales';
type View = 'suppliers' | 'invoices';
export interface ReviewInputs {
  files: InFile[];
  fy: string;
  mode: 'itc' | 'annual';
  regs: { sales?: any[][]; purchases?: any[][] } | string;
}

interface Item {
  key: string;
  status: string;
  group: string;
  basis: string;
  gstin: string;
  name: string;
  bno: string;
  rno: string;
  bdate: string;
  rdate: string;
  period: string;
  bTx: number | null;
  bTax: number | null;
  rTx: number | null;
  rTax: number | null;
  dTx: number;
  dTax: number;
  heads?: { b: number[]; r: number[] };
  code: string;
  effect: string;
  bIds: string[];
  rIds: string[];
  remark: string;
}
const LABEL: Record<Side, { ret: string; party: string; parties: string }> = {
  purchases: { ret: '2B', party: 'Supplier', parties: 'Suppliers' },
  sales: { ret: 'GSTR-1', party: 'Customer', parties: 'Customers' },
};
const GROUPS = [
  {
    id: 'matched',
    label: 'Matched',
    test: (s: string) => s === 'Matched' || s === 'Matched - in next FY 2B' || s === 'Matched - reported in next FY',
  },
  { id: 'diff', label: 'Value difference', test: (s: string) => s === 'Matched - value difference' },
  {
    id: 'other',
    label: 'Matched, no. / GSTIN differs',
    test: (s: string) => /^Matched - (invoice no\. differs|supplier GSTIN differs|GSTIN differs|supplier total)/.test(s),
  },
  { id: 'manual', label: 'Manual', test: (s: string) => s === 'Matched - manual' },
  { id: 'books', label: 'In books only', test: (s: string) => s.startsWith('In books, not in') },
  { id: 'ret', label: 'In return only', test: (s: string) => /^In (2B|GSTR-1), not in books/.test(s) },
  { id: 'nc', label: 'Not compared', test: (s: string) => s.startsWith('Not compared') },
];
const groupOf = (s: string) => GROUPS.find(g => g.test(s))?.id ?? 'other';
const n2 = (v: any) => (typeof v === 'number' ? v : v === null || v === undefined || v === '' ? null : +v || 0);
const r2 = (x: number) => Math.round(x * 100) / 100;
const fmt = (v: number | null | undefined) => (v === null || v === undefined ? '' : inr(v));
const dayNo = (d: string) => {
  const m = String(d).match(/^(\d{2})-(\d{2})-(\d{4})$/);
  return m ? Date.UTC(+m[3], +m[2] - 1, +m[1]) / 86400000 : NaN;
};
const isoToDay = (s: string) => (s ? Date.parse(s) / 86400000 : NaN);
const cloneD = (d: Decisions): Decisions => ({
  manual: new Map([...d.manual].map(([k, v]) => [k, { b: new Set(v.b), r: new Set(v.r) }])),
  keepApart: new Set(d.keepApart),
});

function itemsOf(side: Side, rows: any[][]): Item[] {
  const ids = (id: string) => {
    const [b = '', r = ''] = String(id).split('~');
    return { bIds: b ? b.split(',') : [], rIds: r ? r.split(',') : [] };
  };
  if (side === 'purchases')
    return rows.map(x => {
      const bh = [9, 10, 11, 12].map(i => n2(x[i]) ?? 0),
        rh = [14, 15, 16, 17].map(i => n2(x[i]) ?? 0);
      const has = (i: number) => x[i] !== null && x[i] !== undefined && x[i] !== '';
      const bTax = has(8) ? r2(bh.reduce((a, b) => a + b, 0)) : null,
        rTax = has(13) ? r2(rh.reduce((a, b) => a + b, 0)) : null;
      return {
        key: String(x[28]),
        status: x[0],
        group: groupOf(x[0]),
        basis: x[1] || '',
        gstin: x[2] || '',
        name: x[3] || '',
        bno: x[4] || '',
        rno: x[5] || '',
        bdate: x[6] || '',
        rdate: x[7] || '',
        period: x[20] || '',
        bTx: n2(x[8]),
        bTax,
        rTx: n2(x[13]),
        rTax,
        dTx: r2((n2(x[8]) ?? 0) - (n2(x[13]) ?? 0)),
        dTax: r2((bTax ?? 0) - (rTax ?? 0)),
        heads: { b: bh, r: rh },
        code: x[26] || '',
        effect: x[25] || '',
        remark: x[27] || '',
        ...ids(x[28]),
      };
    });
  return rows.map(x => ({
    key: String(x[22]),
    status: x[0],
    group: groupOf(x[0]),
    basis: x[1] || '',
    gstin: x[2] || x[3] || '',
    name: x[4] || '',
    bno: x[6] || '',
    rno: x[7] || '',
    bdate: x[8] || '',
    rdate: x[9] || '',
    period: x[11] || '',
    bTx: n2(x[12]),
    bTax: n2(x[13]),
    rTx: n2(x[14]),
    rTax: n2(x[15]),
    dTx: n2(x[16]) ?? 0,
    dTax: n2(x[17]) ?? 0,
    code: x[20] || '',
    effect: x[19] || '',
    remark: x[21] || '',
    ...ids(x[22]),
  }));
}

interface Filters {
  groups: Set<string>;
  q: string;
  basis: string;
  period: string;
  from: string;
  to: string;
  diffOnly: boolean;
  minDiff: number;
  supStatus: string;
  supUnmatched: boolean;
}
const newFilters = (): Filters => ({
  groups: new Set(),
  q: '',
  basis: '',
  period: '',
  from: '',
  to: '',
  diffOnly: false,
  minDiff: 0,
  supStatus: '',
  supUnmatched: false,
});

export class Review {
  private decisions: Partial<Record<Side, Decisions>> = {};
  private sig = '';
  private history: { side: Side; label: string; prev: Decisions }[] = [];
  private items: Partial<Record<Side, Item[]>> = {};
  private notes: string[] = [];
  private side: Side = 'purchases';
  private view: View = 'invoices';
  private f: Record<Side, Filters> = { purchases: newFilters(), sales: newFilters() };
  private sel = new Set<string>();
  private sort: { k: string; dir: 1 | -1 } = { k: '', dir: 1 };
  private limit = 300;
  private timer: any = 0;
  private busy = false;
  private err = '';
  constructor(
    private card: HTMLElement,
    private box: HTMLElement,
    private inputs: () => ReviewInputs | null,
    /** the tab's matching options, set under "Matching settings" */
    private matchOf: () => MatchOptions = () => DEFAULT_MATCH,
  ) {}
  private get match(): MatchOptions {
    return this.matchOf();
  }

  /** what the build needs: matching options and the user's decisions (only once the review has run) */
  extra(): ReconExtra {
    return { match: this.match, decisions: this.decisions.purchases || this.decisions.sales ? { ...this.decisions } : undefined };
  }

  refresh(delay = 200) {
    clearTimeout(this.timer);
    this.timer = setTimeout(() => this.run(), delay);
  }

  private async run() {
    const inp = this.inputs();
    if (!inp) {
      this.card.style.display = 'none';
      return;
    }
    const has2b = inp.files.some(f => f.kind === 'GSTR-2B'),
      has1 = inp.files.some(f => f.kind === 'GSTR-1');
    const hasBase = inp.files.some(f => f.kind === 'Workbook');
    if (!has2b && !has1 && !hasBase) {
      this.card.style.display = 'none';
      return;
    }
    if (typeof inp.regs === 'string') {
      this.card.style.display = '';
      this.err = inp.regs;
      this.items = {};
      this.draw();
      return;
    }
    const sig = [inp.fy, inp.mode, ...inp.files.map(f => f.hash || f.name)].join('|');
    if (sig !== this.sig) {
      this.sig = sig;
      this.decisions = {};
      this.history = [];
      this.sel.clear();
    }
    this.busy = true;
    this.err = '';
    try {
      const res = await reconPreview(
        inp.files.filter(f => f.kind !== 'Register'),
        inp.fy,
        inp.mode,
        { ...inp.regs, match: this.match, decisions: this.decisionsOrUndefined() },
      );
      if (!this.decisions.purchases) this.decisions.purchases = cloneD(res.saved.purchases);
      if (!this.decisions.sales) this.decisions.sales = cloneD(res.saved.sales);
      this.items = {
        purchases: res.purchases ? itemsOf('purchases', res.purchases.rows) : undefined,
        sales: res.sales ? itemsOf('sales', res.sales.rows) : undefined,
      };
      this.notes = res.notes;
      if (!this.items[this.side]) this.side = this.items.purchases ? 'purchases' : 'sales';
    } catch (e: any) {
      this.err = e.message;
      this.items = {};
    }
    this.busy = false;
    this.card.style.display = this.items.purchases || this.items.sales || this.err ? '' : 'none';
    this.draw();
    this.onItems();
  }
  private decisionsOrUndefined() {
    return this.decisions.purchases || this.decisions.sales ? { ...this.decisions } : undefined;
  }

  // ---------- filtering
  private filtered(): Item[] {
    const f = this.f[this.side],
      q = f.q.trim().toLowerCase();
    const from = isoToDay(f.from),
      to = isoToDay(f.to);
    let l = (this.items[this.side] || []).filter(x => {
      if (f.groups.size && !f.groups.has(x.group)) return false;
      if (f.basis && (f.basis === '(none)' ? x.basis : x.basis !== f.basis)) return false;
      if (f.period && x.period !== f.period) return false;
      if (q && ![x.name, x.gstin, x.bno, x.rno].some(v => String(v).toLowerCase().includes(q))) return false;
      const d = dayNo(x.bdate || x.rdate);
      if (!isNaN(from) && !(d >= from)) return false;
      if (!isNaN(to) && !(d <= to)) return false;
      const diff = Math.max(Math.abs(x.dTax), Math.abs(x.dTx));
      if (f.diffOnly && diff <= this.match.tol) return false;
      if (f.minDiff && Math.abs(x.dTax) < f.minDiff) return false;
      return true;
    });
    if (this.sort.k) {
      const k = this.sort.k as keyof Item,
        d = this.sort.dir;
      l = [...l].sort((a, b) => {
        const x = a[k] as any,
          y = b[k] as any;
        return (
          (typeof x === 'number' || typeof y === 'number'
            ? (x ?? -Infinity) - (y ?? -Infinity)
            : String(x ?? '').localeCompare(String(y ?? ''))) * d
        );
      });
    }
    return l;
  }
  private suppliers(items: Item[]) {
    const m = new Map<string, any>();
    for (const x of items) {
      if (x.group === 'nc') continue;
      const k = x.gstin || x.name || '(none)';
      const s = m.get(k) || {
        gstin: x.gstin,
        name: x.name,
        nb: 0,
        nr: 0,
        bTx: 0,
        rTx: 0,
        bTax: 0,
        rTax: 0,
        ub: 0,
        ur: 0,
        diffDocs: 0,
      };
      if (x.bno) {
        s.nb += x.bIds.length || 1;
        s.bTx += x.bTx ?? 0;
        s.bTax += x.bTax ?? 0;
      }
      if (x.rno) {
        s.nr += x.rIds.length || 1;
        s.rTx += x.rTx ?? 0;
        s.rTax += x.rTax ?? 0;
      }
      if (x.group === 'books') s.ub++;
      if (x.group === 'ret') s.ur++;
      if (x.group === 'diff') s.diffDocs++;
      if (!s.name) s.name = x.name;
      m.set(k, s);
    }
    const tol = this.match.tol;
    return [...m.values()].map(s => {
      const dTax = r2(s.bTax - s.rTax),
        dTx = r2(s.bTx - s.rTx);
      const status = !s.nr
        ? 'Only in books'
        : !s.nb
          ? 'Only in return'
          : Math.abs(dTax) <= tol && Math.abs(dTx) <= tol
            ? 'Agrees'
            : 'Differs';
      return { ...s, bTx: r2(s.bTx), rTx: r2(s.rTx), bTax: r2(s.bTax), rTax: r2(s.rTax), dTax, dTx, status };
    });
  }
  private filteredSuppliers() {
    const f = this.f[this.side],
      q = f.q.trim().toLowerCase();
    let l = this.suppliers(this.items[this.side] || []).filter(
      s =>
        (!f.supStatus || s.status === f.supStatus) &&
        (!q || String(s.name).toLowerCase().includes(q) || String(s.gstin).toLowerCase().includes(q)) &&
        (!f.minDiff || Math.abs(s.dTax) >= f.minDiff) &&
        (!f.supUnmatched || s.ub + s.ur > 0) &&
        (!f.diffOnly || s.status !== 'Agrees'),
    );
    const k = this.sort.k || 'dTaxAbs',
      d = this.sort.k ? this.sort.dir : -1;
    l = l.sort(
      (a: any, b: any) =>
        (k === 'dTaxAbs'
          ? Math.abs(a.dTax) - Math.abs(b.dTax)
          : typeof a[k] === 'number'
            ? a[k] - b[k]
            : String(a[k]).localeCompare(String(b[k]))) * d,
    );
    return l;
  }
  private totals(l: Item[]) {
    const t = { n: l.length, bTx: 0, bTax: 0, rTx: 0, rTax: 0, bh: [0, 0, 0, 0], rh: [0, 0, 0, 0] };
    for (const x of l) {
      t.bTx += x.bTx ?? 0;
      t.bTax += x.bTax ?? 0;
      t.rTx += x.rTx ?? 0;
      t.rTax += x.rTax ?? 0;
      x.heads?.b.forEach((v, i) => (t.bh[i] += v));
      x.heads?.r.forEach((v, i) => (t.rh[i] += v));
    }
    return t;
  }

  // ---------- decisions
  private d() {
    return (this.decisions[this.side] ||= { manual: new Map(), keepApart: new Set() });
  }
  private push(label: string) {
    this.history.push({ side: this.side, label, prev: cloneD(this.d()) });
  }
  private nextCode() {
    const used = new Set(this.d().manual.keys());
    let i = 1;
    while (used.has('M' + String(i).padStart(3, '0'))) i++;
    return 'M' + String(i).padStart(3, '0');
  }
  private matchSelected(items: Item[]) {
    const b = items.filter(x => x.group === 'books'),
      r = items.filter(x => x.group === 'ret');
    if (!b.length || !r.length) return;
    const code = this.nextCode();
    this.push(`Matched ${b.map(x => x.bno).join(' + ')} with ${r.map(x => x.rno).join(' + ')} (${code})`);
    const g = { b: new Set(b.flatMap(x => x.bIds)), r: new Set(r.flatMap(x => x.rIds)) };
    this.d().manual.set(code, g);
    this.sel.clear();
    this.refresh(0);
  }
  private unmatchSelected(items: Item[]) {
    const m = items.filter(x => x.bIds.length && x.rIds.length);
    if (!m.length) return;
    this.push(
      `Unmatched ${m.length} pair(s): ${m
        .slice(0, 3)
        .map(x => x.bno || x.rno)
        .join(', ')}${m.length > 3 ? ', …' : ''}`,
    );
    const d = this.d();
    for (const x of m) {
      if (x.code && x.code !== 'X' && d.manual.has(x.code)) d.manual.delete(x.code);
      else [...x.bIds, ...x.rIds].forEach(id => d.keepApart.add(id));
    }
    this.sel.clear();
    this.refresh(0);
  }
  private releaseSelected(items: Item[]) {
    const k = items.filter(x => x.code === 'X');
    if (!k.length) return;
    this.push(`Allowed automatic matching again for ${k.length} document(s)`);
    for (const x of k) [...x.bIds, ...x.rIds].forEach(id => this.d().keepApart.delete(id));
    this.sel.clear();
    this.refresh(0);
  }
  private undo(i: number) {
    const h = this.history[i];
    if (!h) return;
    this.decisions[h.side] = h.prev;
    this.history.splice(i);
    this.sel.clear();
    this.refresh(0);
  }

  // ---------- download
  private async download() {
    const L = LABEL[this.side],
      inp = this.inputs();
    let rows: any[][];
    if (this.view === 'suppliers') {
      rows = [
        [
          `${L.party} GSTIN`,
          L.party,
          'Docs in books',
          `Docs in ${L.ret}`,
          'Books taxable',
          `${L.ret} taxable`,
          'Books tax',
          `${L.ret} tax`,
          'Diff taxable',
          'Diff tax',
          'Unmatched in books',
          `Unmatched in ${L.ret}`,
          'Status',
        ],
        ...this.filteredSuppliers().map(s => [
          s.gstin,
          s.name,
          s.nb,
          s.nr,
          s.bTx,
          s.rTx,
          s.bTax,
          s.rTax,
          s.dTx,
          s.dTax,
          s.ub,
          s.ur,
          s.status,
        ]),
      ];
    } else {
      const l = this.filtered(),
        t = this.totals(l);
      rows = [
        [
          'Status',
          'Match basis',
          `${L.party} GSTIN`,
          L.party,
          'Books no.',
          `${L.ret} no.`,
          'Books date',
          `${L.ret} date`,
          `${L.ret} period`,
          'Books taxable',
          'Books tax',
          `${L.ret} taxable`,
          `${L.ret} tax`,
          'Diff taxable',
          'Diff tax',
          'Match code',
          'Effect / action',
          'Remarks',
        ],
        ...l.map(x => [
          x.status,
          x.basis,
          x.gstin,
          x.name,
          x.bno,
          x.rno,
          x.bdate,
          x.rdate,
          x.period,
          x.bTx,
          x.bTax,
          x.rTx,
          x.rTax,
          x.dTx,
          x.dTax,
          x.code,
          x.effect,
          x.remark,
        ]),
        [
          'Total',
          `${t.n} rows`,
          '',
          '',
          '',
          '',
          '',
          '',
          '',
          r2(t.bTx),
          r2(t.bTax),
          r2(t.rTx),
          r2(t.rTax),
          r2(t.bTx - t.rTx),
          r2(t.bTax - t.rTax),
          '',
          '',
          '',
        ],
      ];
    }
    const title = [
      `${this.side === 'purchases' ? 'Purchases vs GSTR-2B' : 'Sales vs GSTR-1'} - ${this.view === 'suppliers' ? L.parties.toLowerCase() : 'invoices'} (filtered view)`,
      `FY ${inp?.fy ?? ''} · filters: ${this.filterText() || 'none'} · prepared by James & James LLP`,
    ];
    const bytes = await simpleXlsx(this.view === 'suppliers' ? L.parties : 'Invoices', rows, title);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(
      new Blob([bytes as any], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    );
    a.download = `Review_${this.side}_${this.view}_${inp?.fy ?? ''}.xlsx`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  private filterText() {
    const f = this.f[this.side],
      p: string[] = [];
    if (f.groups.size) p.push([...f.groups].map(g => GROUPS.find(x => x.id === g)?.label).join(' / '));
    if (f.q) p.push(`search "${f.q}"`);
    if (f.basis) p.push(`rule ${f.basis}`);
    if (f.period) p.push(`period ${f.period}`);
    if (f.from || f.to) p.push(`dates ${f.from || '…'} to ${f.to || '…'}`);
    if (f.diffOnly) p.push('differences only');
    if (f.minDiff) p.push(`tax diff ≥ ${f.minDiff}`);
    if (this.view === 'suppliers' && f.supStatus) p.push(f.supStatus);
    if (this.view === 'suppliers' && f.supUnmatched) p.push('with unmatched');
    return p.join(', ');
  }

  // ---------- dashboard ("At a glance" card): built from the same items as the review, whole side, not filtered
  /** called after every run so the tab's dashboard card can redraw */
  onItems: () => void = () => {};
  hasItems() {
    return !!(this.items.purchases || this.items.sales);
  }
  dashHtml(): string {
    return (['purchases', 'sales'] as Side[])
      .filter(s => this.items[s])
      .map(s => this.dashSide(s))
      .join('');
  }
  private dashSide(side: Side): string {
    const L = LABEL[side],
      all = (this.items[side] || []).filter(x => x.group !== 'nc');
    if (!all.length) return '';
    const isM = (g: string) => g === 'matched' || g === 'diff' || g === 'other' || g === 'manual';
    const cnt = (g: string) => all.filter(x => x.group === g).length;
    const matched = all.filter(x => isM(x.group));
    const sum = (l: Item[], f: (x: Item) => number) => r2(l.reduce((a, x) => a + f(x), 0));
    const lo = (x: Item) => Math.min(x.bTax ?? 0, x.rTax ?? 0);
    const covered = sum(matched, lo),
      booksTax = sum(all, x => x.bTax ?? 0);
    const books = all.filter(x => x.group === 'books'),
      ret = all.filter(x => x.group === 'ret');
    const over = sum(matched, x => Math.max(0, (x.bTax ?? 0) - (x.rTax ?? 0)));
    const atRisk = r2(sum(books, x => x.bTax ?? 0) + over);
    const pctN = (matched.length / all.length) * 100,
      pctV = booksTax > 0 ? (covered / booksTax) * 100 : 100;
    const diffs = cnt('diff');
    const tone = (p: number) => (p >= 95 ? 'ok' : p >= 80 ? 'warn' : 'bad');
    const meters = `<div class="dash-grid">${meter(`Documents matched (${all.length} in all)`, pctN, `${cnt('matched') + cnt('manual') + cnt('other')} agree · ${diffs} value difference · ${books.length} in books only · ${ret.length} in ${L.ret} only`, tone(pctN))}
      ${meter(side === 'purchases' ? `Books tax backed by ${L.ret}` : `Books tax reported in ${L.ret}`, pctV, `${fmt(covered)} of ${fmt(booksTax)}`, tone(pctV))}</div>`;
    const tiles =
      side === 'purchases'
        ? `<div class="kpis">${tile('ITC claimable', fmt(covered), 'lower of books and 2B on matched invoices', 'ok')}
          ${tile('ITC at risk', fmt(atRisk), `${fmt(sum(books, x => x.bTax ?? 0))} in books only + ${fmt(over)} above 2B`, atRisk > 1 ? 'bad' : 'ok')}
          ${tile('In 2B, not in books', fmt(sum(ret, x => x.rTax ?? 0)), `${ret.length} document(s): credit available but not booked`, ret.length ? 'warn' : 'ok')}
          ${tile(
            'Value differences',
            String(diffs),
            `net tax ${fmt(
              sum(
                all.filter(x => x.group === 'diff'),
                x => x.dTax,
              ),
            )} (books − 2B)`,
            diffs ? 'warn' : 'ok',
          )}</div>`
        : `<div class="kpis">${tile('Tax matched', fmt(covered), `${matched.length} document(s)`, 'ok')}
          ${tile('In books, not in GSTR-1', fmt(sum(books, x => x.bTax ?? 0)), `${books.length} document(s): possibly not reported`, books.length ? 'bad' : 'ok')}
          ${tile('In GSTR-1, not in books', fmt(sum(ret, x => x.rTax ?? 0)), `${ret.length} document(s)`, ret.length ? 'warn' : 'ok')}
          ${tile(
            'Value differences',
            String(diffs),
            `net tax ${fmt(
              sum(
                all.filter(x => x.group === 'diff'),
                x => x.dTax,
              ),
            )} (books − GSTR-1)`,
            diffs ? 'warn' : 'ok',
          )}</div>`;
    // by return period (documents in books only have none)
    const per = [...new Set(all.map(x => x.period || ''))].sort((a, b) =>
      a === '' ? 1 : b === '' ? -1 : (a.slice(3) + a.slice(0, 2)).localeCompare(b.slice(3) + b.slice(0, 2)),
    );
    const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const lab = (p: string) => (/^\d\d\/\d{4}$/.test(p) ? `${MON[+p.slice(0, 2) - 1]} ${p.slice(5)}` : 'No period');
    const by = (pred: (g: string) => boolean) => per.map(p => all.filter(x => (x.period || '') === p && pred(x.group)).length);
    const chart = stackedBars(
      per.map(lab),
      [
        { name: 'Matched', color: '#1F9E89', vals: by(g => g === 'matched' || g === 'other' || g === 'manual') },
        { name: 'Value difference', color: '#E08A1E', vals: by(g => g === 'diff') },
        { name: 'In books only', color: '#B42318', vals: by(g => g === 'books') },
        { name: `In ${L.ret} only`, color: '#4552B0', vals: by(g => g === 'ret') },
      ],
      {
        aria: `Documents by ${L.ret} period and status`,
        note: `documents by ${L.ret} period (books-only documents have no period)`,
      },
    );
    // parties with the largest net difference
    const g = new Map<string, { name: string; gstin: string; net: number; un: number }>();
    for (const x of all) {
      const k = x.gstin || x.name;
      const e = g.get(k) || { name: x.name, gstin: x.gstin, net: 0, un: 0 };
      e.net = r2(e.net + (x.bTax ?? 0) - (x.rTax ?? 0));
      if (x.group === 'books' || x.group === 'ret') e.un++;
      if (!e.name) e.name = x.name;
      g.set(k, e);
    }
    const top = [...g.values()]
      .filter(e => Math.abs(e.net) > 0.5)
      .sort((a, b) => Math.abs(b.net) - Math.abs(a.net))
      .slice(0, 10)
      .map(e => ({ label: e.name || e.gstin, sub: `${e.gstin || 'no GSTIN'}${e.un ? ` · ${e.un} unmatched` : ''}`, val: e.net }));
    const rank = ranked(
      `Largest tax differences by ${L.party.toLowerCase()} (books − ${L.ret})`,
      top,
      `Every ${L.party.toLowerCase()} agrees.`,
      fmt as any,
    );
    return `<div class="dash-side"><h3>${side === 'purchases' ? 'Purchases vs GSTR-2B' : 'Sales vs GSTR-1'}</h3>${meters}${tiles}<div class="dash-grid">${chart}${rank}</div></div>`;
  }

  // ---------- drawing
  private draw() {
    const b = this.box;
    if (this.err) {
      b.innerHTML = `<p class="warn-error">${esc(this.err)}</p>`;
      return;
    }
    const sides = (['purchases', 'sales'] as Side[]).filter(s => this.items[s]);
    if (!sides.length) {
      b.innerHTML = '<p class="muted">Loading…</p>';
      return;
    }
    const L = LABEL[this.side],
      f = this.f[this.side],
      m = this.match,
      all = this.items[this.side] || [];
    const settings = `<p class="muted small" style="margin:0 0 6px">Matched with: ${esc(describeInvoices(m, this.side === 'purchases'))}. Change it under Matching settings.</p>`;
    // KPI tiles (whole side, not filtered)
    const kp = (g: string, label: string, cls = '') => {
      const l = all.filter(x => x.group === g || (g === 'matched' && ['matched', 'other', 'manual'].includes(x.group)));
      const t = this.totals(l);
      const amt = g === 'books' ? t.bTax : g === 'ret' ? t.rTax : g === 'diff' ? t.bTax - t.rTax : t.bTax;
      return `<button class="kpi ${cls} ${f.groups.size === 1 && f.groups.has(g) ? 'on' : ''}" data-kpi="${g}"><span>${label}</span><b>${l.length}</b><em>${g === 'diff' ? 'diff ' : 'tax '}${fmt(r2(amt))}</em></button>`;
    };
    const allT = this.totals(all.filter(x => x.group !== 'nc'));
    const kpis = `<div class="kpis">${kp('matched', 'Matched', 'ok')}${kp('diff', 'Value difference', 'warn')}${kp('books', 'In books only', 'bad')}${kp('ret', `In ${L.ret} only`, 'bad')}
      <div class="kpi tot"><span>Tax: books vs ${L.ret}</span><b>${fmt(r2(allT.bTax))}</b><em>${L.ret} ${fmt(r2(allT.rTax))} · diff ${fmt(r2(allT.bTax - allT.rTax))}</em></div></div>`;
    const sideTabs =
      sides.length > 1
        ? `<div class="gtabs" style="padding:0;margin:14px 0 4px">${sides.map(s => `<button class="gtab ${s === this.side ? 'on' : ''}" data-side="${s}">${s === 'purchases' ? 'Purchases vs GSTR-2B' : 'Sales vs GSTR-1'}</button>`).join('')}</div>`
        : '';
    const viewTabs = `<div class="rv-bar"><div class="seg">${(['suppliers', 'invoices'] as View[]).map(v => `<button class="${v === this.view ? 'on' : ''}" data-view="${v}">${v === 'suppliers' ? L.parties : 'Invoices'}</button>`).join('')}</div>
      <button class="ghost" id="rvDl">⤓ Download this view (.xlsx)</button></div>`;
    const periods = [...new Set(all.map(x => x.period).filter(Boolean))].sort((a, b) =>
      (a.slice(3) + a.slice(0, 2)).localeCompare(b.slice(3) + b.slice(0, 2)),
    );
    const bases = [...new Set(all.map(x => x.basis).filter(Boolean))].sort();
    const chips = GROUPS.filter(g => all.some(x => x.group === g.id))
      .map(
        g =>
          `<button class="fchip ${f.groups.has(g.id) ? 'on' : ''}" data-grp="${g.id}">${g.label.replace('return', L.ret)} <span>${all.filter(x => x.group === g.id).length}</span></button>`,
      )
      .join('');
    const filters =
      this.view === 'invoices'
        ? `<div class="rv-filters"><div class="fchips">${chips}</div><div class="row">
          <label class="grow">Search<input data-f="q" value="${esc(f.q)}" placeholder="${L.party.toLowerCase()}, GSTIN or invoice no."/></label>
          <label>Match rule<select data-f="basis"><option value="">All</option>${bases.map(x => `<option ${x === f.basis ? 'selected' : ''}>${esc(x)}</option>`).join('')}<option value="(none)" ${f.basis === '(none)' ? 'selected' : ''}>(unmatched)</option></select></label>
          <label>${L.ret} period<select data-f="period"><option value="">All</option>${periods.map(p => `<option ${p === f.period ? 'selected' : ''}>${p}</option>`).join('')}</select></label>
          <label>Date from<input type="date" data-f="from" value="${f.from}" style="min-width:0"/></label>
          <label>to<input type="date" data-f="to" value="${f.to}" style="min-width:0"/></label>
          <label>Tax diff at least ₹<input type="number" min="0" data-f="minDiff" value="${f.minDiff || ''}" style="min-width:0;width:110px"/></label>
          <label class="tg"><input type="checkbox" data-f="diffOnly" ${f.diffOnly ? 'checked' : ''}/> Differences only</label>
          <button class="link" id="rvClear">Clear filters</button></div></div>`
        : `<div class="rv-filters"><div class="row">
          <label class="grow">Search<input data-f="q" value="${esc(f.q)}" placeholder="name or GSTIN"/></label>
          <label>Status<select data-f="supStatus"><option value="">All</option>${['Agrees', 'Differs', 'Only in books', 'Only in return'].map(s => `<option value="${s}" ${s === f.supStatus ? 'selected' : ''}>${s.replace('return', L.ret)}</option>`).join('')}</select></label>
          <label>Tax diff at least ₹<input type="number" min="0" data-f="minDiff" value="${f.minDiff || ''}" style="min-width:0;width:110px"/></label>
          <label class="tg"><input type="checkbox" data-f="supUnmatched" ${f.supUnmatched ? 'checked' : ''}/> Has unmatched documents</label>
          <label class="tg"><input type="checkbox" data-f="diffOnly" ${f.diffOnly ? 'checked' : ''}/> Not agreeing only</label>
          <button class="link" id="rvClear">Clear filters</button></div></div>`;
    const th = (k: string, label: string, r = false) =>
      `<th class="${r ? 'r' : ''} sortable" data-sort="${k}">${label}${this.sort.k === k ? (this.sort.dir === 1 ? ' ▲' : ' ▼') : ''}</th>`;
    let table = '',
      foot = '';
    if (this.view === 'suppliers') {
      const l = this.filteredSuppliers();
      const t = l.reduce(
        (a, s) => {
          a.bTx += s.bTx;
          a.rTx += s.rTx;
          a.bTax += s.bTax;
          a.rTax += s.rTax;
          return a;
        },
        { bTx: 0, rTx: 0, bTax: 0, rTax: 0 },
      );
      table = `<div class="tablewrap rv-tw"><table class="rv"><thead><tr>${th('name', L.party)}${th('nb', 'Books', true)}${th('nr', L.ret, true)}${th('bTax', 'Books tax', true)}${th('rTax', `${L.ret} tax`, true)}${th('dTax', 'Diff tax', true)}${th('ub', 'Unmatched books', true)}${th('ur', `Unmatched ${L.ret}`, true)}${th('status', 'Status')}</tr></thead><tbody>
        ${l
          .slice(0, this.limit)
          .map(
            s =>
              `<tr class="click" data-sup="${esc(s.gstin || s.name)}"><td><b>${esc(s.name || '—')}</b><div class="muted">${esc(s.gstin || 'no GSTIN')}</div></td><td class="r">${s.nb}</td><td class="r">${s.nr}</td><td class="r">${fmt(s.bTax)}</td><td class="r">${fmt(s.rTax)}</td><td class="r ${Math.abs(s.dTax) > m.tol ? 'neg' : ''}">${fmt(s.dTax)}</td><td class="r">${s.ub || ''}</td><td class="r">${s.ur || ''}</td><td><span class="st st-${s.status.split(' ')[0].toLowerCase()}">${esc(s.status.replace('return', L.ret))}</span></td></tr>`,
          )
          .join('')}</tbody></table></div>
        ${l.length > this.limit ? `<button class="link" id="rvMore">Show more (${l.length - this.limit} left)</button>` : ''}`;
      foot = `<div class="rv-foot"><span><b>${l.length}</b> ${L.parties.toLowerCase()}</span><span>Taxable: books ${fmt(r2(t.bTx))} · ${L.ret} ${fmt(r2(t.rTx))} · diff <b>${fmt(r2(t.bTx - t.rTx))}</b></span><span>Tax: books ${fmt(r2(t.bTax))} · ${L.ret} ${fmt(r2(t.rTax))} · diff <b>${fmt(r2(t.bTax - t.rTax))}</b></span></div>`;
    } else {
      const l = this.filtered(),
        t = this.totals(l);
      const selItems = all.filter(x => this.sel.has(x.key)),
        st = this.totals(selItems);
      const canMatch =
        selItems.some(x => x.group === 'books') &&
        selItems.some(x => x.group === 'ret') &&
        selItems.every(x => x.group === 'books' || x.group === 'ret');
      const canUnmatch = selItems.some(x => x.bIds.length && x.rIds.length);
      const canRelease = selItems.some(x => x.code === 'X');
      const gstins = new Set(selItems.map(x => x.gstin).filter(Boolean));
      table = `<div class="tablewrap rv-tw"><table class="rv"><thead><tr><th class="cb"><input type="checkbox" id="rvAll" title="Select all shown"/></th>${th('status', 'Status')}${th('name', L.party)}${th('bno', 'Books no.')}${th('rno', `${L.ret} no.`)}${th('bdate', 'Date')}${th('bTx', 'Books taxable', true)}${th('rTx', `${L.ret} taxable`, true)}${th('bTax', 'Books tax', true)}${th('rTax', `${L.ret} tax`, true)}${th('dTax', 'Diff tax', true)}</tr></thead><tbody>
        ${l
          .slice(0, this.limit)
          .map(
            x => `<tr class="${this.sel.has(x.key) ? 'sel' : ''}"><td class="cb"><input type="checkbox" data-sel="${esc(x.key)}" ${this.sel.has(x.key) ? 'checked' : ''} ${x.group === 'nc' ? 'disabled' : ''}/></td>
          <td><span class="st st-${x.group}">${esc((t => t.charAt(0).toUpperCase() + t.slice(1))(x.status.replace(/^Matched - /, '')))}</span>${x.code ? ` <span class="tag" title="${x.code === 'X' ? 'Kept apart by you' : 'Manual match'}">${esc(x.code)}</span>` : ''}<div class="muted small">${esc(x.basis)}</div></td>
          <td>${esc(x.name)}<div class="muted small">${esc(x.gstin)}</div></td><td>${esc(x.bno)}</td><td>${esc(x.rno)}<div class="muted small">${esc(x.period)}</div></td>
          <td class="nw">${esc(x.bdate || x.rdate)}${x.bdate && x.rdate && x.bdate !== x.rdate ? `<div class="muted small">${L.ret} ${esc(x.rdate)}</div>` : ''}</td>
          <td class="r">${fmt(x.bTx)}</td><td class="r">${fmt(x.rTx)}</td><td class="r">${fmt(x.bTax)}</td><td class="r">${fmt(x.rTax)}</td><td class="r ${Math.abs(x.dTax) > m.tol ? 'neg' : ''}">${x.bTax !== null && x.rTax !== null ? fmt(x.dTax) : ''}</td></tr>`,
          )
          .join('')}</tbody></table></div>
        ${l.length > this.limit ? `<button class="link" id="rvMore">Show more (${l.length - this.limit} left)</button>` : ''}
        ${!l.length ? '<p class="muted" style="margin:10px 0">Nothing matches these filters.</p>' : ''}`;
      const heads =
        this.side === 'purchases'
          ? `<span class="muted">CGST ${fmt(r2(t.bh[0] - t.rh[0]))} · SGST ${fmt(r2(t.bh[1] - t.rh[1]))} · IGST ${fmt(r2(t.bh[2] - t.rh[2]))} · Cess ${fmt(r2(t.bh[3] - t.rh[3]))} (diff by head)</span>`
          : '';
      foot = `<div class="rv-foot"><span><b>${t.n}</b> rows shown by filters</span><span>Taxable: books ${fmt(r2(t.bTx))} · ${L.ret} ${fmt(r2(t.rTx))} · diff <b>${fmt(r2(t.bTx - t.rTx))}</b></span><span>Tax: books ${fmt(r2(t.bTax))} · ${L.ret} ${fmt(r2(t.rTax))} · diff <b>${fmt(r2(t.bTax - t.rTax))}</b></span>${heads}</div>
        ${
          selItems.length
            ? `<div class="rv-sel"><span><b>${selItems.length}</b> selected · books tax ${fmt(r2(st.bTax))} · ${L.ret} tax ${fmt(r2(st.rTax))} · diff <b>${fmt(r2(st.bTax - st.rTax))}</b> (taxable diff ${fmt(r2(st.bTx - st.rTx))})${canMatch && gstins.size > 1 ? ' · <span class="warn-warn">different GSTINs</span>' : ''}</span>
          <span class="acts">${canMatch ? '<button class="small" id="rvMatch">Match selected</button>' : ''}${canUnmatch ? '<button class="small ghost" id="rvUnmatch">Unmatch</button>' : ''}${canRelease ? '<button class="small ghost" id="rvRelease">Allow auto-match again</button>' : ''}<button class="link" id="rvSelClear">Clear</button></span>
          ${!canMatch && !canUnmatch && !canRelease ? `<span class="muted">To match by hand, select at least one "In books only" and one "In ${L.ret} only" row.</span>` : ''}</div>`
            : ''
        }`;
    }
    const hist = this.history.length
      ? `<details class="rv-hist"><summary>Your changes (${this.history.length}) - carried into the workbook as Match codes</summary><ol>${this.history.map((h, i) => `<li>${esc(h.label)} <button class="link" data-undo="${i}">undo${i < this.history.length - 1 ? ' (and later)' : ''}</button></li>`).join('')}</ol></details>`
      : '';
    const notes = this.notes.length
      ? `<ul class="warns">${this.notes.map(n => `<li class="warn-warn">${esc(n)}</li>`).join('')}</ul>`
      : '';
    b.innerHTML =
      settings +
      sideTabs +
      kpis +
      viewTabs +
      filters +
      foot.replace('rv-foot', 'rv-foot top') +
      table +
      hist +
      notes +
      (this.busy ? '<p class="muted">Updating…</p>' : '');
    this.wire();
  }

  private wire() {
    const b = this.box,
      f = this.f[this.side];
    const q = (s: string) => b.querySelectorAll(s) as NodeListOf<any>;
    q('[data-side]').forEach(
      el =>
        (el.onclick = () => {
          this.side = el.dataset.side;
          this.sel.clear();
          this.sort = { k: '', dir: 1 };
          this.limit = 300;
          this.draw();
        }),
    );
    q('[data-view]').forEach(
      el =>
        (el.onclick = () => {
          this.view = el.dataset.view;
          this.sort = { k: '', dir: 1 };
          this.limit = 300;
          this.draw();
        }),
    );
    q('[data-kpi]').forEach(
      el =>
        (el.onclick = () => {
          const g = el.dataset.kpi;
          const only = f.groups.size === 1 && f.groups.has(g);
          f.groups = only ? new Set() : new Set(g === 'matched' ? ['matched', 'other', 'manual'] : [g]);
          this.view = 'invoices';
          this.draw();
        }),
    );
    q('[data-grp]').forEach(
      el =>
        (el.onclick = () => {
          const g = el.dataset.grp;
          f.groups.has(g) ? f.groups.delete(g) : f.groups.add(g);
          this.limit = 300;
          this.draw();
        }),
    );
    q('[data-f]').forEach(el => {
      const k = el.dataset.f;
      const apply = () => {
        (f as any)[k] = el.type === 'checkbox' ? el.checked : k === 'minDiff' ? Math.max(0, +el.value || 0) : el.value;
        this.limit = 300;
        this.draw();
        const n = b.querySelector(`[data-f="${k}"]`) as any;
        if (n && el.type !== 'checkbox' && el.tagName === 'INPUT' && el.type !== 'date') {
          n.focus();
          try {
            n.setSelectionRange(n.value.length, n.value.length);
          } catch {
            /* number */
          }
        }
      };
      if (el.tagName === 'INPUT' && (el.type === 'text' || !el.type || el.type === 'search'))
        el.oninput = () => {
          clearTimeout((el as any)._t);
          (el as any)._t = setTimeout(apply, 250);
        };
      else el.onchange = apply;
    });
    const clear = b.querySelector('#rvClear') as any;
    if (clear)
      clear.onclick = () => {
        this.f[this.side] = newFilters();
        this.draw();
      };
    q('th[data-sort]').forEach(
      el =>
        (el.onclick = () => {
          const k = el.dataset.sort;
          this.sort = this.sort.k === k ? { k, dir: (this.sort.dir * -1) as 1 | -1 } : { k, dir: 1 };
          this.draw();
        }),
    );
    q('tr[data-sup]').forEach(
      el =>
        (el.onclick = () => {
          this.f[this.side] = { ...newFilters(), q: el.dataset.sup };
          this.view = 'invoices';
          this.draw();
        }),
    );
    q('[data-sel]').forEach(
      el =>
        (el.onchange = () => {
          el.checked ? this.sel.add(el.dataset.sel) : this.sel.delete(el.dataset.sel);
          this.draw();
        }),
    );
    const allCb = b.querySelector('#rvAll') as any;
    if (allCb)
      allCb.onchange = () => {
        const l = this.filtered()
          .slice(0, this.limit)
          .filter(x => x.group !== 'nc');
        if (allCb.checked) l.forEach(x => this.sel.add(x.key));
        else l.forEach(x => this.sel.delete(x.key));
        this.draw();
      };
    const selItems = () => (this.items[this.side] || []).filter(x => this.sel.has(x.key));
    const on = (id: string, fn: () => void) => {
      const el = b.querySelector('#' + id) as any;
      if (el) el.onclick = fn;
    };
    on('rvMatch', () => this.matchSelected(selItems()));
    on('rvUnmatch', () => this.unmatchSelected(selItems()));
    on('rvRelease', () => this.releaseSelected(selItems()));
    on('rvSelClear', () => {
      this.sel.clear();
      this.draw();
    });
    on('rvMore', () => {
      this.limit += 300;
      this.draw();
    });
    on('rvDl', () => {
      this.download().catch(e => console.error('download failed', e));
    });
    q('[data-undo]').forEach(el => (el.onclick = () => this.undo(+el.dataset.undo)));
  }
}

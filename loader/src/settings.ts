// Matching settings: one step per tab, shown before any matching, where the user picks the rules, tolerances and windows.
// Every tab reads its options from here (Input Reco and Annual return: invoice matching and supplier names; GST ledger flow:
// how differences are explained and how ITC invoices are paired with 2B; Revenue reco: invoice matching and customer GSTINs).
// Kept in the browser per tab; the chosen options are also written into each workbook so the reader knows how it was matched.
import { DEFAULT_MATCH, MATCH_PRESETS, MatchOptions } from './recon';
import { esc, store } from './util';

export type TabId = 'itc' | 'annual' | 'flow' | 'revenue';

/** Revenue reco: how a books customer without a GSTIN gets one from GSTR-1 */
export interface CustSettings {
  /** apply the customer list (name + GSTIN) when one is loaded */
  list: boolean;
  /** apply a name that appears in the books with exactly one GSTIN */
  exact: boolean;
  /** map a customer when its invoice number AND amount agree with a GSTR-1 document of one GSTIN */
  invAmt: boolean;
  /** ... without asking (otherwise it is only suggested) */
  invAuto: boolean;
  /** invoices that must agree before the mapping is applied without asking */
  invMin: number;
  /** suggest GSTR-1 customers whose learnt name is at least this similar (%) */
  nameMin: number;
}
/** GST ledger flow: which rules may explain a books vs portal difference */
export interface FlowRules {
  netted: boolean;
  head: boolean;
  timing: boolean;
  after: boolean;
  fee: boolean;
}
export interface FlowSettings {
  /** amounts within this many rupees are equal */
  tol: number;
  rules: FlowRules;
  /** the same payment / journal in another month: up to this many months apart */
  entryMonths: number;
  /** month totals that offset: up to this many months apart */
  monthGap: number;
}
export interface Settings {
  /** invoice matching (Input Reco, Annual return, Revenue reco), or ITC invoices vs 2B (GST ledger flow) */
  match: MatchOptions;
  /** Input Reco / Annual return: suggest a 2B supplier for a books name at least this similar (%) */
  supNameMin: number;
  cust: CustSettings;
  flow: FlowSettings;
}

export const DEFAULT_CUST: CustSettings = { list: true, exact: true, invAmt: true, invAuto: true, invMin: 1, nameMin: 55 };
export const DEFAULT_FLOW: FlowSettings = {
  tol: 1,
  rules: { netted: true, head: true, timing: true, after: true, fee: true },
  entryMonths: 3,
  monthGap: 2,
};
const clone = <T>(x: T): T => JSON.parse(JSON.stringify(x));
export function defaults(_tab?: TabId): Settings {
  return { match: clone(DEFAULT_MATCH), supNameMin: 55, cust: clone(DEFAULT_CUST), flow: clone(DEFAULT_FLOW) };
}
/** stored settings over the defaults (missing keys from older versions take the default) */
export function merge(tab: TabId, saved: any): Settings {
  const d = defaults(tab);
  if (!saved || typeof saved !== 'object') return d;
  return {
    match: { ...d.match, ...(saved.match || {}), rules: { ...d.match.rules, ...(saved.match?.rules || {}) } },
    supNameMin: typeof saved.supNameMin === 'number' ? saved.supNameMin : d.supNameMin,
    cust: { ...d.cust, ...(saved.cust || {}) },
    flow: { ...d.flow, ...(saved.flow || {}), rules: { ...d.flow.rules, ...(saved.flow?.rules || {}) } },
  };
}

// ---------------------------------------------------------------- what each tab offers
interface Field {
  path: string;
  label: string;
  hint: string;
  kind: 'bool' | 'num' | 'select';
  min?: number;
  step?: number;
  unit?: string;
  options?: [string, string][];
  /** shown only when this returns true */
  when?: (s: Settings) => boolean;
}
interface Section {
  id: string;
  title: string;
  intro: string;
  /** the Strict / Standard / Lenient presets of the invoice rules */
  presets?: boolean;
  fields: Field[];
  /** fixed rules the user cannot change, said once so the basis is complete */
  fixed?: string[];
}
const rule = (k: keyof MatchOptions['rules'], label: string, hint: string): Field => ({
  path: `match.rules.${k}`,
  label,
  hint,
  kind: 'bool',
});
const matchFields = (party: string, totals: boolean): Field[] => [
  rule('exact', '1 · Same invoice no.', `Same ${party} GSTIN and invoice number, ignoring punctuation, spaces and case`),
  rule('digits', '2 · Same digits of no.', 'The digits of the invoice number agree (INV-0045 = 45), amounts within tolerance'),
  rule(
    'tail',
    '3 · Last digits of no.',
    'One number ends with the other (or the last 3+ digits agree), amounts within tolerance',
  ),
  rule('dateAmt', '4 · GSTIN + date + amount', 'Same GSTIN, dates within the window, same taxable value and tax; number ignored'),
  rule('otherGstin', '5 · No. + amount, GSTIN differs', `Same invoice number and amount, but the ${party} GSTIN differs`),
  ...(totals
    ? [
        rule(
          'supplierTotal',
          '6 · Supplier totals',
          "A supplier's leftover documents agree in total on both sides (purchases only)",
        ),
      ]
    : []),
  { path: 'match.tol', label: 'Tolerance ₹', hint: 'Amounts within this many rupees are equal', kind: 'num', min: 0, step: 0.5 },
  {
    path: 'match.pct',
    label: 'Tolerance % (rules 2–3)',
    hint: 'Extra tolerance, as a % of the amount, when the number is only partly the same',
    kind: 'num',
    min: 0,
    step: 0.5,
  },
  {
    path: 'match.days',
    label: 'Date window ± days (rule 4)',
    hint: 'Dates may differ by this many days',
    kind: 'num',
    min: 0,
    step: 1,
  },
  {
    path: 'match.compare',
    label: 'Amounts compared',
    hint: 'What must agree for a match',
    kind: 'select',
    options: [
      ['both', 'Taxable value and tax'],
      ['tax', 'Tax only'],
    ],
  },
];
const supNames: Section = {
  id: 'sup',
  title: 'Supplier names (purchase lines without a GSTIN)',
  intro:
    'Books lines without a supplier GSTIN take it from GSTR-2B. Exact names are applied automatically; similar names are suggested under "Confirm supplier names" and applied only when you confirm them.',
  fields: [
    {
      path: 'supNameMin',
      label: 'Suggest names at least % similar',
      hint: 'Lower finds more suggestions (and more wrong ones)',
      kind: 'num',
      min: 30,
      step: 5,
    },
  ],
  fixed: ['Lines whose invoice no. and amount agree with one 2B document take that supplier GSTIN.'],
};
const DEFS: Record<TabId, Section[]> = {
  itc: [
    {
      id: 'inv',
      title: 'How invoices are matched: purchase register vs GSTR-2B',
      intro: 'Manual matches always come first. Then the rules below, in this order. Untick a rule to leave it out.',
      presets: true,
      fields: matchFields('supplier', true),
    },
    supNames,
  ],
  annual: [
    {
      id: 'inv',
      title: 'How invoices are matched: purchases vs GSTR-2B and sales vs GSTR-1',
      intro: 'Manual matches always come first. Then the rules below, in this order. Rule 6 applies to purchases only.',
      presets: true,
      fields: matchFields('supplier / customer', true),
      fixed: [
        'Next-year returns count for April to October only (ICAI Technical Guide on GSTR-9); next-year credit notes are tagged for review.',
      ],
    },
    supNames,
  ],
  flow: [
    {
      id: 'why',
      title: 'How differences between books and the portal are explained',
      intro:
        'Each rule below takes the part of a difference it can explain, in this order. What no rule explains stays "Unexplained", listed with its entries. Untick a rule to see those differences as unexplained instead.',
      fields: [
        {
          path: 'flow.tol',
          label: 'Equal within ₹',
          hint: 'Amounts within this many rupees are equal',
          kind: 'num',
          min: 0,
          step: 0.5,
        },
        {
          path: 'flow.rules.netted',
          label: 'Netted on portal',
          hint: 'ITC reversals, RCM, reclaims and credit notes that the portal nets into another row',
          kind: 'bool',
        },
        {
          path: 'flow.rules.head',
          label: 'Tax head',
          hint: 'IGST vs CGST + SGST, cross-utilisation of ITC, accounts not split by head',
          kind: 'bool',
        },
        {
          path: 'flow.rules.timing',
          label: 'Timing (month to month)',
          hint: 'The same entry, or offsetting month totals, in different months',
          kind: 'bool',
        },
        {
          path: 'flow.entryMonths',
          label: 'Same entry: up to months apart',
          hint: 'A payment or journal of the same amount in books and on the portal',
          kind: 'num',
          min: 1,
          step: 1,
          when: s => s.flow.rules.timing,
        },
        {
          path: 'flow.monthGap',
          label: 'Month totals: up to months apart',
          hint: 'Month totals that offset (books higher one month, portal the next)',
          kind: 'num',
          min: 1,
          step: 1,
          when: s => s.flow.rules.timing,
        },
        {
          path: 'flow.rules.after',
          label: 'After the period',
          hint: 'Entries that fall after the period on the other side (books in March, return in April)',
          kind: 'bool',
        },
        {
          path: 'flow.rules.fee',
          label: 'Fee booked to tax',
          hint: 'Interest or late fee paid in cash but debited to the tax ledger in books',
          kind: 'bool',
        },
      ],
      fixed: [
        'Opening balance differences are always shown first.',
        'Portal entries count by return period, books by voucher date (unless you move a line).',
      ],
    },
    {
      id: 'inv',
      title: 'ITC invoices: purchase vouchers in the books vs GSTR-2B',
      intro: 'Used for "ITC three ways". Manual matches do not apply here; the rules below run in this order.',
      presets: true,
      fields: matchFields('supplier', true),
    },
  ],
  revenue: [
    {
      id: 'inv',
      title: 'How invoices are matched: books vs GSTR-1',
      intro: 'The rules below run in this order. B2CS, Table 8 and e-commerce rows have no invoices and are compared by month.',
      presets: true,
      fields: matchFields('customer', false),
      fixed: [
        'Books by invoice date, GSTR-1 by return period; shifts between months of the same year are not reconciling items.',
      ],
    },
    {
      id: 'cust',
      title: 'Customer GSTIN for books entries without one',
      intro:
        'GSTR-1 has the customer GSTIN but no name. A books customer without a GSTIN gets one by the methods ticked here, in this order; your own choice under "Confirm customer names" always wins. Every mapping shows its basis there.',
      fields: [
        {
          path: 'cust.list',
          label: '1 · Customer list',
          hint: 'A name with one GSTIN on the customer list you loaded',
          kind: 'bool',
        },
        {
          path: 'cust.exact',
          label: '2 · Same name in the books',
          hint: 'The books carry this exact name (spacing, case and "Pvt Ltd" ignored) with one GSTIN on other entries',
          kind: 'bool',
        },
        {
          path: 'cust.invAmt',
          label: '3 · Invoice no. + amount',
          hint: 'Invoices of this customer whose number AND amount agree with GSTR-1 documents of one GSTIN',
          kind: 'bool',
        },
        {
          path: 'cust.invAuto',
          label: 'Apply rule 3 without asking',
          hint: 'Off: the GSTIN is only suggested, and you confirm it',
          kind: 'bool',
          when: s => s.cust.invAmt,
        },
        {
          path: 'cust.invMin',
          label: 'Invoices that must agree',
          hint: 'At least this many invoices must agree, all under the same GSTIN',
          kind: 'num',
          min: 1,
          step: 1,
          when: s => s.cust.invAmt,
        },
        {
          path: 'cust.nameMin',
          label: 'Suggest names at least % similar',
          hint: 'Similar names are only suggested, never applied without you',
          kind: 'num',
          min: 30,
          step: 5,
        },
      ],
      fixed: ['Amounts are compared with the tolerance and "Amounts compared" set above.'],
    },
  ],
};

// ---------------------------------------------------------------- plain-words description (screen and workbooks)
const presetOf = (m: MatchOptions) =>
  Object.entries(MATCH_PRESETS).find(([, p]) => JSON.stringify(p) === JSON.stringify(m))?.[0] ?? 'Custom';
export function describeInvoices(m: MatchOptions, totals = true): string {
  const names: [keyof MatchOptions['rules'], string][] = [
    ['exact', 'same invoice no.'],
    ['digits', 'digits of no.'],
    ['tail', 'last digits'],
    ['dateAmt', `GSTIN + date ±${m.days}d + amount`],
    ['otherGstin', 'no. + amount with another GSTIN'],
    ...(totals ? ([['supplierTotal', 'supplier totals']] as [keyof MatchOptions['rules'], string][]) : []),
  ];
  const on = names.filter(([k]) => m.rules[k]).map(([, n]) => n);
  return `${presetOf(m)}: ${on.length ? on.join(', ') : 'manual matches only'}; tolerance ₹${m.tol}${m.pct ? ` / ${m.pct}%` : ''}; ${m.compare === 'tax' ? 'tax only' : 'taxable value and tax'} compared`;
}
/** one line per section, saying what was used: shown in the card's header and written to the workbooks */
export function describe(tab: TabId, s: Settings): string[] {
  const out: string[] = [];
  if (tab === 'flow') {
    const r = s.flow.rules;
    const on = [
      r.netted && 'netted on portal',
      r.head && 'tax head',
      r.timing && `timing (entries ${s.flow.entryMonths} / month totals ${s.flow.monthGap} months apart)`,
      r.after && 'after the period',
      r.fee && 'fee booked to tax',
    ].filter(Boolean);
    out.push(`Differences explained by: opening balance, ${on.join(', ') || 'nothing else'}; equal within ₹${s.flow.tol}`);
    out.push(`ITC invoices vs 2B: ${describeInvoices(s.match)}`);
  } else if (tab === 'revenue') {
    out.push(`Invoices: ${describeInvoices(s.match, false)}`);
    const c = s.cust;
    const how = [
      c.list && 'customer list',
      c.exact && 'same name in books',
      c.invAmt &&
        `invoice no. + amount (${c.invMin}+ invoice${c.invMin > 1 ? 's' : ''}${c.invAuto ? ', applied' : ', suggested only'})`,
    ].filter(Boolean);
    out.push(`Customer GSTIN by: ${how.join(', ') || 'your confirmation only'}; similar names suggested from ${c.nameMin}%`);
  } else {
    out.push(`Invoices: ${describeInvoices(s.match)}`);
    out.push(`Supplier names: exact applied, similar suggested from ${s.supNameMin}%`);
  }
  return out;
}

// ---------------------------------------------------------------- the card
const getP = (o: any, p: string) => p.split('.').reduce((a, k) => a?.[k], o);
const setP = (o: any, p: string, v: any) => {
  const ks = p.split('.');
  const last = ks.pop()!;
  ks.reduce((a, k) => a[k], o)[last] = v;
};
export class SettingsCard {
  private s: Partial<Record<TabId, Settings>> = {};
  constructor(
    private box: HTMLElement,
    private onChange: (tab: TabId) => void,
  ) {}

  get(tab: TabId): Settings {
    let s = this.s[tab];
    if (!s) {
      let saved = store.get(`jjset:${tab}`);
      // before v0.12 the invoice rules were set in "Review matches" and shared by Input Reco and Annual return
      if (!saved && (tab === 'itc' || tab === 'annual') && store.get('jjmatch')) saved = { match: store.get('jjmatch') };
      s = this.s[tab] = merge(tab, saved);
    }
    return s;
  }
  private rendering = false;
  private save(tab: TabId) {
    store.set(`jjset:${tab}`, this.s[tab]);
    // after the event: re-rendering removes the input that fired it (a blur then would fire change inside innerHTML)
    setTimeout(() => {
      this.render(tab);
      this.onChange(tab);
    }, 0);
  }

  render(tab: TabId) {
    const s = this.get(tab);
    const open = store.get(`jjsetOpen:${tab}`) ?? true;
    const isDefault = JSON.stringify(s) === JSON.stringify(defaults(tab));
    const field = (f: Field) => {
      if (f.when && !f.when(s)) return '';
      const v = getP(s, f.path);
      if (f.kind === 'bool')
        return `<label class="tg" title="${esc(f.hint)}"><input type="checkbox" data-set="${f.path}" ${v ? 'checked' : ''}/> ${esc(f.label)}</label>`;
      if (f.kind === 'select')
        return `<label title="${esc(f.hint)}">${esc(f.label)}<select data-set="${f.path}">${f.options!.map(([k, l]) => `<option value="${k}" ${k === v ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select></label>`;
      return `<label title="${esc(f.hint)}">${esc(f.label)}<input type="number" min="${f.min ?? 0}" step="${f.step ?? 1}" data-set="${f.path}" value="${v}" style="min-width:0;width:110px"/></label>`;
    };
    const lines = describe(tab, s);
    const section = (x: Section, i: number) => {
      const bools = x.fields.filter(f => f.kind === 'bool'),
        rest = x.fields.filter(f => f.kind !== 'bool');
      const preset = x.presets ? presetOf(s.match) : '';
      return `<div class="set-sec"><h3>${esc(x.title)}${preset ? ` <span class="tag">${preset}</span>` : ''}</h3>
        <p class="muted">${esc(x.intro)}</p>
        ${
          x.presets
            ? `<div class="row" style="gap:8px">${Object.keys(MATCH_PRESETS)
                .map(p => `<button class="gtab ${p === preset ? 'on' : ''}" data-preset="${p}">${p}</button>`)
                .join('')}</div>`
            : ''
        }
        ${bools.length ? `<div class="rv-rules">${bools.map(field).join('')}</div>` : ''}
        ${rest.length ? `<div class="row">${rest.map(field).join('')}</div>` : ''}
        ${x.fixed?.length ? `<ul class="set-fixed">${x.fixed.map(t => `<li>${esc(t)}</li>`).join('')}</ul>` : ''}
        <p class="set-now"><b>In use:</b> ${esc(lines[i] ?? '')}</p></div>`;
    };
    this.rendering = true;
    this.box.innerHTML = `<details class="rv-set" ${open ? 'open' : ''}><summary><b>${isDefault ? 'Standard settings' : 'Your settings'}</b> <span class="muted">${esc(lines.join(' · '))}</span></summary>
      <div class="rv-set-in">${DEFS[tab].map(section).join('')}
      <div class="row" style="justify-content:space-between"><span class="muted">Saved in this browser for the ${tab === 'itc' ? 'Input Reco' : tab === 'annual' ? 'Annual return' : tab === 'flow' ? 'GST ledger flow' : 'Revenue reco'} tab, and written to the workbook's Read Me.</span>
      ${isDefault ? '' : '<button class="small ghost" data-setreset>Back to standard settings</button>'}</div></div></details>`;
    this.rendering = false;
    const q = (sel: string) => this.box.querySelectorAll(sel) as NodeListOf<any>;
    (this.box.querySelector('details') as any).ontoggle = (e: any) => store.set(`jjsetOpen:${tab}`, e.target.open);
    q('[data-preset]').forEach(
      el =>
        (el.onclick = () => {
          s.match = clone(MATCH_PRESETS[el.dataset.preset]);
          this.save(tab);
        }),
    );
    q('[data-set]').forEach(
      el =>
        (el.onchange = () => {
          if (this.rendering) return;
          const p = el.dataset.set;
          const v =
            el.type === 'checkbox' ? el.checked : el.tagName === 'SELECT' ? el.value : Math.max(+el.min || 0, +el.value || 0);
          setP(s, p, v);
          this.save(tab);
        }),
    );
    const reset = this.box.querySelector('[data-setreset]') as any;
    if (reset)
      reset.onclick = () => {
        this.s[tab] = defaults(tab);
        this.save(tab);
      };
  }
}

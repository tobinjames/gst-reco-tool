// Step-by-step view (wizard) over the existing cards. Each tab is a list of steps; a step shows some of the cards and
// hides the rest with the `offstep` class, so the cards and their logic are untouched. The old stacked page is still
// available with ?ui=classic.
import { store } from './util';
export type TabId = 'itc' | 'flow' | 'revenue' | 'annual';
/** what the wizard needs to know about the page to decide if a step is ready */
export interface Ctx {
  kinds: Set<string>; // file kinds loaded for this tab
  hasWorkpaper: boolean; // a saved workpaper is loaded: it carries its own files
  gstin: string;
  nFiles: number;
  /** how many files of each kind are loaded */
  count?: Record<string, number>;
}
export interface Step {
  id: string;
  title: string;
  /** one plain sentence: what to do on this step */
  hint: string;
  /** ids of the cards shown on this step */
  cards: string[];
  /** more cards, shown one at a time as sub-tabs under the step (results with several views) */
  sub?: string[];
  /** the manual's section for this step */
  help?: string;
  /** null when the step is complete, else the reason in plain words */
  missing?: (c: Ctx) => string | null;
}

const need = (c: Ctx, ...groups: string[][]) => {
  if (c.hasWorkpaper) return null;
  const miss = groups.filter(g => !g.some(k => c.kinds.has(k))).map(g => g.join(' or '));
  return miss.length ? `Add ${miss.join(' and ')} to continue` : null;
};
const client = (c: Ctx) => (c.gstin ? null : 'Enter the client GSTIN (it is filled in when a GSTR file is added)');

const SETTINGS: Step = {
  id: 'settings',
  title: 'Settings',
  hint: 'The defaults suit most clients. Change them only if you know you need to, then press Next.',
  cards: ['setcard'],
};

export const STEPS: Record<TabId, Step[]> = {
  itc: [
    {
      id: 'files',
      title: 'Add files',
      hint: 'Drop the GSTR-2B JSON files and your purchase register. Or drop a saved workpaper: it brings back everything.',
      cards: ['c-files'],
      missing: c => need(c, ['GSTR-2B'], ['Register', 'Workbook']),
    },
    {
      id: 'client',
      title: 'Client & year',
      hint: 'Check the client and the financial year, and tell us which register column holds what.',
      cards: ['c-client', 'regcard'],
      missing: client,
    },
    SETTINGS,
    {
      id: 'review',
      title: 'Review matches',
      hint: 'See what matched and what did not. Fix a match by hand if needed.',
      cards: ['supcard', 'rvcard'],
    },
    { id: 'build', title: 'Download', hint: 'Build the workbook and save it.', cards: ['c-build'] },
  ],
  annual: [
    {
      id: 'files',
      title: 'Add files',
      hint: 'Best: drop the saved workpapers of the other tabs. Otherwise drop GSTR-1, 2B, 3B and the registers.',
      cards: ['c-files'],
      missing: c => (c.nFiles ? null : 'Add at least one file to continue'),
    },
    {
      id: 'client',
      title: 'Client & year',
      hint: 'Check the client and the financial year, and map the register columns.',
      cards: ['c-client', 'regcard'],
      missing: client,
    },
    SETTINGS,
    {
      id: 'review',
      title: 'Review matches',
      hint: 'See what matched and what did not. Fix a match by hand if needed.',
      cards: ['supcard', 'rvcard'],
    },
    { id: 'build', title: 'Download', hint: 'Build the workbook and save it.', cards: ['c-build'] },
    {
      id: 'json',
      title: 'Upload JSON',
      hint: 'When the workbook is finished in Excel, drop it back here to get the GSTR-9 upload JSON.',
      cards: ['step4'],
    },
  ],
  flow: [
    {
      id: 'files',
      title: 'Add files',
      hint: 'Drop the portal Credit, Cash and Liability ledgers and the books GST ledgers. Or drop a saved workpaper.',
      cards: ['c-files'],
      missing: c => need(c, ['Portal ledger'], ['Books ledger', 'Flow workbook']),
    },
    {
      id: 'client',
      title: 'Client & year',
      hint: 'Check the client and the financial year.',
      cards: ['c-client'],
      missing: client,
    },
    {
      id: 'books',
      title: 'Ledgers & balances',
      hint: 'Confirm what each books account is, and enter opening balances if they are not in the files.',
      cards: ['fl-portal', 'fl-books', 'fl-bal'],
    },
    SETTINGS,
    {
      id: 'review',
      title: 'Results',
      hint: 'The flow, the reason for each difference, and the cross-checks.',
      cards: [],
      sub: ['fl-flow', 'fl-reasons', 'fl-checks'],
    },
    { id: 'build', title: 'Download', hint: 'Build the workbook and save it.', cards: ['fl-build'] },
  ],
  revenue: [
    {
      id: 'files',
      title: 'Add files',
      hint: 'Drop the GSTR-1 files, GSTR-3B PDFs and the sales register or GL dump. Or drop a saved workpaper.',
      cards: ['c-files'],
      missing: c => need(c, ['GSTR-1'], ['Register', 'Books ledger', 'Revenue workbook']),
    },
    {
      id: 'client',
      title: 'Client & year',
      hint: 'Check the client and the financial year.',
      cards: ['c-client'],
      missing: client,
    },
    {
      id: 'books',
      title: 'Books',
      hint: 'Tell us which column holds what, and confirm customer names.',
      cards: ['rv-books', 'regcard', 'rv-cust'],
    },
    SETTINGS,
    {
      id: 'review',
      title: 'Results',
      hint: 'The dashboard, the two statements and the reason for each difference.',
      cards: ['rv-dash'],
      sub: ['rv-st1', 'rv-st2', 'rv-items'],
    },
    { id: 'build', title: 'Download', hint: 'Build the workbook and save it.', cards: ['rv-build'] },
  ],
};

/** the wizard is the normal page; ?ui=classic brings back the old stacked one */
export const wizardOn = () => {
  try {
    return !(typeof location !== 'undefined' && /[?&]ui=classic\b/.test(location.search));
  } catch {
    return true;
  }
};

/** steps whose number is `i` (0-based) can be left only when every step before it is complete */
export function firstBlocked(steps: Step[], c: Ctx): number {
  const i = steps.findIndex(s => s.missing?.(c));
  return i < 0 ? steps.length : i;
}
/** can the user stand on step i? Step 0 (add files) is always open */
export const canVisit = (steps: Step[], i: number, c: Ctx) => i <= firstBlocked(steps, c);

/** what each tab needs, shown as a checklist on the Add files step */
export interface Need {
  label: string;
  kinds: string[];
  /** where it comes from, in a few words */
  from: string;
  /** shown as "x of 12" when this many files are expected */
  of?: number;
  optional?: boolean;
}
export const NEEDS: Record<TabId, Need[]> = {
  itc: [
    { label: 'GSTR-2B (one JSON per month)', kinds: ['GSTR-2B'], from: 'GST portal → Returns → GSTR-2B → Download', of: 12 },
    { label: 'Purchase register', kinds: ['Register'], from: 'Tally, Zoho or Excel' },
  ],
  flow: [
    { label: 'Portal ledgers (credit, cash, liability)', kinds: ['Portal ledger'], from: 'GST portal → Ledgers' },
    { label: 'Books GST ledgers or GL dump', kinds: ['Books ledger'], from: 'Tally, Zoho or Excel' },
    { label: 'GSTR-3B PDFs', kinds: ['GSTR-3B'], from: 'GST portal → Returns', optional: true },
  ],
  revenue: [
    { label: 'GSTR-1 (JSON or ZIP, every month)', kinds: ['GSTR-1'], from: 'GST portal → Returns → GSTR-1', of: 12 },
    { label: 'Sales register or revenue ledgers', kinds: ['Register', 'Books ledger'], from: 'Tally, Zoho or Excel' },
    { label: 'GSTR-3B PDFs', kinds: ['GSTR-3B'], from: 'GST portal → Returns', optional: true },
  ],
  annual: [
    {
      label: 'Saved workpapers of the other tabs',
      kinds: ['Workbook', 'Revenue workbook', 'Flow workbook'],
      from: 'the workbooks you built earlier',
      optional: true,
    },
    { label: 'GSTR-1, GSTR-2B, GSTR-3B', kinds: ['GSTR-1', 'GSTR-2B', 'GSTR-3B'], from: 'GST portal → Returns' },
    { label: 'Registers', kinds: ['Register'], from: 'Tally, Zoho or Excel', optional: true },
  ],
};
const HOME: Record<TabId, [string, string]> = {
  itc: ['Check purchase ITC', 'Compare GSTR-2B with your purchase register, month by month.'],
  revenue: ['Check sales and revenue', 'Compare books, GSTR-1 and GSTR-3B for the year.'],
  flow: ['Check the GST ledgers', 'Compare the portal ledgers with the GST accounts in your books.'],
  annual: ['Prepare GSTR-9', 'Build the annual return workbook and the upload file.'],
};
const TAB_NAME: Record<TabId, string> = {
  itc: 'Input Reco',
  revenue: 'Revenue reco',
  flow: 'GST ledger flow',
  annual: 'Annual return',
};
const HELP: Record<TabId, string> = { itc: 'h-itc', flow: 'h-flow', revenue: 'h-rev', annual: 'h-annual' };
const escH = (x: string) => x.replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
const el = (id: string) => document.getElementById(id);
export const noteKey = (tab: string, step: string, gstin: string, fy: string) => `jjnote:${tab}:${step}:${gstin || '-'}:${fy}`;

export interface Hooks {
  ctx: () => Ctx;
  /** open a tab (the page's own tab switch) */
  openTab: (t: TabId) => void;
  /** load the demo files of a tab */
  demo: (t: TabId) => void;
}

export class Wizard {
  private tab: TabId = 'itc';
  private idx: Record<string, number> = {};
  private subIdx: Record<string, number> = {};
  home = false;
  /** the user has saved a file since the last change: no warning when closing the page */
  saved = true;
  constructor(private h: Hooks) {}

  get steps() {
    return STEPS[this.tab];
  }
  get current() {
    return this.idx[this.tab] || 0;
  }
  setTab(t: TabId) {
    this.tab = t;
    this.apply();
  }
  showHome() {
    this.home = true;
    this.apply();
    if (typeof window !== 'undefined') window.scrollTo({ top: 0 });
  }
  go(i: number) {
    this.idx[this.tab] = Math.max(0, Math.min(this.steps.length - 1, i));
    this.apply();
    if (typeof window !== 'undefined') window.scrollTo({ top: 0 });
  }
  goTo(id: string) {
    const i = this.steps.findIndex(s => s.id === id);
    if (i >= 0) this.go(i);
  }
  /** after a saved workpaper is added: files and settings are back, so go on to the review step */
  afterRestore() {
    this.home = false;
    this.goTo('review');
  }
  private homeHtml(): string {
    const last = store.get('jjrecent') as { tab: TabId; gstin: string; legal: string; fy: string; at: string } | null;
    const lastTab = (store.get('jjmode') as TabId) || null;
    const cards = (Object.keys(HOME) as TabId[])
      .map(
        t =>
          `<button type="button" class="home-card" data-t="${t}"><b>${HOME[t][0]}</b><span>${HOME[t][1]}</span><em>${TAB_NAME[t]} →</em></button>`,
      )
      .join('');
    const rec = last
      ? `<div class="home-last"><div><b>Last worked on</b><br>${escH(last.legal || last.gstin)} · ${escH(last.fy)} · ${TAB_NAME[last.tab] || ''} <span class="muted">(${escH(last.at)})</span><br><span class="muted">Your files are not kept on this page. Add them again, or drop the workpaper you saved to carry on.</span></div>` +
        `<button type="button" class="wiz-next" data-t="${last.tab}">Continue in ${TAB_NAME[last.tab] || 'the tab'} →</button></div>`
      : lastTab
        ? ''
        : '';
    return `<h2>What do you want to do?</h2><p class="muted">Pick one. The page then takes you through it one step at a time: add files, check settings, review, download.</p><div class="home-grid">${cards}</div>${rec}<p class="muted">New here? Open <b>Help &amp; guide</b> at the top, or practise with demo files on the first step of any tab.</p>`;
  }
  private checklistHtml(c: Ctx): string {
    const cnt = c.count || {};
    const rows = NEEDS[this.tab]
      .map(n => {
        const have = n.kinds.reduce((a, k) => a + (cnt[k] || 0), 0);
        const ok = have > 0;
        const of = n.of
          ? ` <span class="muted">${have} of ${n.of}</span>`
          : have > 1
            ? ` <span class="muted">${have} files</span>`
            : '';
        return `<li class="${ok ? 'ok' : n.optional ? 'opt' : 'miss'}"><span class="ck">${ok ? '✓' : n.optional ? '○' : '✕'}</span><span><b>${n.label}</b>${n.optional ? ' <span class="muted">(optional)</span>' : ''}${of}${ok ? '' : `<br><span class="muted">Get it from: ${n.from}</span>`}</span></li>`;
      })
      .join('');
    return `<div class="wiz-check"><b>What you need on this tab</b><ul>${rows}</ul><div class="wiz-check-act"><button type="button" class="link" id="wizGuide">How do I download these files?</button> · <button type="button" class="link" id="wizDemo">Try with demo files</button></div></div>`;
  }
  /** redraw the stepper and show only this step's cards */
  apply() {
    if (typeof document === 'undefined') return;
    const c = this.h.ctx(),
      steps = this.steps;
    if (c.nFiles > 0) this.home = false;
    document.body.classList.toggle('wizhome', this.home);
    const homeEl = el('wizhome');
    if (this.home && homeEl) {
      homeEl.innerHTML = this.homeHtml();
      homeEl.querySelectorAll('button[data-t]').forEach(
        (b: any) =>
          (b.onclick = () => {
            this.home = false;
            this.h.openTab(b.dataset.t);
          }),
      );
      return;
    }
    if (c.gstin) {
      const fy = (el('fy') as HTMLSelectElement | null)?.value || '';
      store.set('jjrecent', {
        tab: this.tab,
        gstin: c.gstin,
        legal: (el('legal') as HTMLInputElement | null)?.value || '',
        fy,
        at: new Date().toISOString().slice(0, 10),
      });
    }
    // never stay on a step whose earlier steps are no longer complete (e.g. the files were removed)
    if (!canVisit(steps, this.current, c)) this.idx[this.tab] = Math.max(0, firstBlocked(steps, c));
    const cur = this.current,
      st = steps[cur];
    const subKey = `${this.tab}:${st.id}`;
    const sub = st.sub || [];
    const subSel = Math.min(this.subIdx[subKey] || 0, Math.max(0, sub.length - 1));
    const show = new Set([...st.cards, ...(sub.length ? [sub[subSel]] : [])]);
    const all = new Set(steps.flatMap(s => [...s.cards, ...(s.sub || [])]));
    document.querySelectorAll('main > section.card').forEach((e: any) => {
      e.classList.toggle('offstep', all.has(e.id) && !show.has(e.id));
    });
    const bar = el('wizbar'),
      foot = el('wizfoot');
    if (!bar || !foot) return;
    const blocked = firstBlocked(steps, c);
    const title = (id: string) =>
      (el(id)?.querySelector('h2')?.textContent || id)
        .replace(/^\s*\d+[a-z]?\s*/, '')
        .replace(/shared by all tabs/i, '')
        .trim();
    const helpId = st.help || (st.id === 'files' ? 'h-files' : st.id === 'settings' ? 'h-settings' : HELP[this.tab]);
    const fyv = (el('fy') as HTMLSelectElement | null)?.value || '';
    const nk = noteKey(this.tab, st.id, c.gstin, fyv);
    const note = (store.get(nk) as string) || '';
    bar.innerHTML =
      `<ol class="wiz-steps">` +
      steps
        .map((s, k) => {
          const cls = k === cur ? 'cur' : k < cur && k < blocked ? 'done' : 'todo';
          const ok = canVisit(steps, k, c);
          return `<li class="${cls}"><button type="button" data-k="${k}" ${ok ? '' : 'disabled'}><span class="wn">${cls === 'done' ? '✓' : k + 1}</span><span class="wt">${s.title}</span></button></li>`;
        })
        .join('') +
      `</ol><div class="wiz-head"><div><b>Step ${cur + 1} of ${steps.length} · ${st.title}</b><p>${st.hint} <a href="#help/${helpId}">Need help with this step?</a></p></div>` +
      `<div class="wiz-acts"><button type="button" class="wiz-add" id="wizAdd">＋ Add files <span>${c.nFiles}</span></button><button type="button" class="link" id="wizHome">⌂ Start</button></div></div>` +
      (st.id === 'files' ? this.checklistHtml(c) : '') +
      (sub.length
        ? `<div class="wiz-sub">${[...st.cards, ...sub]
            .filter(id => sub.includes(id))
            .map((id, k) => `<button type="button" data-s="${k}" class="${k === subSel ? 'on' : ''}">${escH(title(id))}</button>`)
            .join('')}</div>`
        : '');
    bar.querySelectorAll('button[data-k]').forEach((b: any) => (b.onclick = () => this.go(+b.dataset.k)));
    bar.querySelectorAll('button[data-s]').forEach(
      (b: any) =>
        (b.onclick = () => {
          this.subIdx[subKey] = +b.dataset.s;
          this.apply();
        }),
    );
    const on = (id: string, f: () => void) => {
      const b = el(id);
      if (b) b.onclick = f;
    };
    on('wizAdd', () => el('pick')?.click());
    on('wizHome', () => this.showHome());
    on('wizDemo', () => this.h.demo(this.tab));
    on('wizGuide', () => {
      const g = el('guide') as HTMLDetailsElement | null;
      if (g) {
        g.open = true;
        el('c-files')?.classList.remove('offstep');
        g.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
    const why = st.missing?.(c) || null,
      last = cur === steps.length - 1;
    foot.innerHTML =
      `<button type="button" class="wiz-back" id="wizBack" ${cur === 0 ? 'disabled' : ''}>← Back</button>` +
      `<details class="wiz-note" ${note ? 'open' : ''}><summary>✎ Note</summary><textarea id="wizNote" rows="3" placeholder="A remark for this step, for yourself or the reviewer. Saved in this browser and in the workpaper.">${escH(note)}</textarea></details>` +
      `<span class="wiz-why">${why || ''}</span>` +
      (last
        ? ''
        : `<button type="button" class="wiz-next" id="wizNext" ${why ? 'disabled' : ''}>Next: ${steps[cur + 1].title} →</button>`);
    on('wizBack', () => this.go(cur - 1));
    on('wizNext', () => this.go(cur + 1));
    const ta = el('wizNote') as HTMLTextAreaElement | null;
    if (ta) ta.oninput = () => store.set(nk, ta.value);
  }
}

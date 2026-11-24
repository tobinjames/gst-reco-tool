// Step-by-step view (wizard) over the existing cards. Each tab is a list of steps; a step shows some of the cards and
// hides the rest with the `offstep` class, so the cards and their logic are untouched. Switched on with ?ui=wizard.
export type TabId = 'itc' | 'flow' | 'revenue' | 'annual';
/** what the wizard needs to know about the page to decide if a step is ready */
export interface Ctx {
  kinds: Set<string>; // file kinds loaded for this tab
  hasWorkpaper: boolean; // a saved workpaper is loaded: it carries its own files
  gstin: string;
  nFiles: number;
}
export interface Step {
  id: string;
  title: string;
  /** one plain sentence: what to do on this step */
  hint: string;
  /** ids of the cards shown on this step */
  cards: string[];
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
      cards: ['fl-flow', 'fl-reasons', 'fl-checks'],
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
      cards: ['rv-dash', 'rv-st1', 'rv-st2', 'rv-items'],
    },
    { id: 'build', title: 'Download', hint: 'Build the workbook and save it.', cards: ['rv-build'] },
  ],
};

export const wizardOn = () => {
  try {
    return typeof location !== 'undefined' && /[?&]ui=wizard\b/.test(location.search);
  } catch {
    return false;
  }
};

/** steps whose number is `i` (0-based) can be left only when every step before it is complete */
export function firstBlocked(steps: Step[], c: Ctx): number {
  const i = steps.findIndex(s => s.missing?.(c));
  return i < 0 ? steps.length : i;
}
/** can the user stand on step i? Step 0 (add files) is always open */
export const canVisit = (steps: Step[], i: number, c: Ctx) => i <= firstBlocked(steps, c);

export class Wizard {
  private tab: TabId = 'itc';
  private idx: Record<string, number> = {};
  constructor(private ctx: () => Ctx) {}

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
    this.goTo('review');
  }
  /** redraw the stepper and show only this step's cards */
  apply() {
    if (typeof document === 'undefined') return;
    const c = this.ctx(),
      steps = this.steps,
      i = this.current;
    // never stay on a step whose earlier steps are no longer complete (e.g. the files were removed)
    if (!canVisit(steps, i, c)) this.idx[this.tab] = Math.max(0, firstBlocked(steps, c));
    const cur = this.current,
      show = new Set(steps[cur].cards);
    const all = new Set(steps.flatMap(s => s.cards));
    document.querySelectorAll('main > section.card').forEach((el: any) => {
      el.classList.toggle('offstep', all.has(el.id) && !show.has(el.id));
    });
    const bar = document.getElementById('wizbar');
    const foot = document.getElementById('wizfoot');
    if (!bar || !foot) return;
    const blocked = firstBlocked(steps, c);
    bar.innerHTML =
      `<ol class="wiz-steps">` +
      steps
        .map((s, k) => {
          const st = k === cur ? 'cur' : k < cur && k < blocked ? 'done' : 'todo';
          const ok = canVisit(steps, k, c);
          return `<li class="${st}"><button type="button" data-k="${k}" ${ok ? '' : 'disabled'}><span class="wn">${
            st === 'done' ? '✓' : k + 1
          }</span><span class="wt">${s.title}</span></button></li>`;
        })
        .join('') +
      `</ol><div class="wiz-head"><div><b>Step ${cur + 1} of ${steps.length} · ${steps[cur].title}</b><p>${steps[cur].hint}</p></div>` +
      `<button type="button" class="wiz-add" id="wizAdd">＋ Add files <span>${c.nFiles}</span></button></div>`;
    bar.querySelectorAll('button[data-k]').forEach((b: any) => (b.onclick = () => this.go(+b.dataset.k)));
    const addBtn = document.getElementById('wizAdd');
    if (addBtn) addBtn.onclick = () => document.getElementById('pick')?.click();
    const why = steps[cur].missing?.(c) || null,
      last = cur === steps.length - 1;
    foot.innerHTML =
      `<button type="button" class="wiz-back" id="wizBack" ${cur === 0 ? 'disabled' : ''}>← Back</button>` +
      `<span class="wiz-why">${why || ''}</span>` +
      (last
        ? ''
        : `<button type="button" class="wiz-next" id="wizNext" ${why ? 'disabled' : ''}>Next: ${steps[cur + 1].title} →</button>`);
    const back = document.getElementById('wizBack'),
      next = document.getElementById('wizNext');
    if (back) back.onclick = () => this.go(cur - 1);
    if (next) next.onclick = () => this.go(cur + 1);
  }
}

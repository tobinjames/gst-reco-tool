// Progress bar for long steps (reading files, matching, building workbooks). The work runs on the page's own thread, so
// each step() waits for the browser to paint before the next piece of work starts. Outside a browser (tests) it does nothing.
const hasDom = typeof document !== 'undefined';
let box: HTMLElement | null = null,
  depth = 0;

function el() {
  if (!hasDom) return null;
  if (!box) {
    box = document.createElement('div');
    box.id = 'progress';
    box.setAttribute('role', 'progressbar');
    box.setAttribute('aria-live', 'polite');
    box.innerHTML = `<div class="pg-card"><div class="pg-label"></div><div class="pg-track"><div class="pg-bar"></div></div><div class="pg-sub"></div></div>`;
    document.body.appendChild(box);
  }
  return box;
}
/** let the browser paint the bar before the next (blocking) piece of work */
const paint = () => (hasDom ? new Promise<void>(r => requestAnimationFrame(() => setTimeout(r, 0))) : Promise.resolve());

/**
 * Show where a long step is. frac 0..1 draws the bar; leave it out for a moving bar when the size is unknown.
 * Nested calls (a build that also reads files) share one bar.
 */
export async function step(label: string, frac?: number, sub = '') {
  const b = el();
  if (!b) return;
  b.classList.add('on');
  (b.querySelector('.pg-label') as HTMLElement).textContent = label;
  (b.querySelector('.pg-sub') as HTMLElement).textContent = sub;
  const bar = b.querySelector('.pg-bar') as HTMLElement;
  const known = frac !== undefined && isFinite(frac);
  bar.classList.toggle('indet', !known);
  bar.style.width = known ? `${Math.max(2, Math.min(100, frac! * 100)).toFixed(1)}%` : '';
  b.setAttribute('aria-valuenow', known ? String(Math.round(frac! * 100)) : '');
  await paint();
}
/** run a long task with the bar shown; the bar goes away when the outermost task ends, also on an error */
export async function withProgress<T>(label: string, task: () => Promise<T> | T): Promise<T> {
  depth++;
  try {
    await step(label);
    return await task();
  } finally {
    if (--depth === 0) box?.classList.remove('on');
  }
}

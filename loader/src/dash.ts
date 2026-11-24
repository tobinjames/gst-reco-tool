// Shared pieces of the dashboards ("At a glance" cards): tiles, a progress meter, stacked monthly bars, a ranked list.
// Styled with the .kpi / .rvchart rules in index.html plus the .dash-* rules next to them. Plain strings, no state.
import { esc, inr } from './util';

export type Tone = 'ok' | 'warn' | 'bad' | 'tot' | '';
const r2 = (x: number) => Math.round(x * 100) / 100;

/** a whole-rupee figure shortened for axes and labels: 12.5 L, 3.2 Cr */
export const short = (v: number) =>
  Math.abs(v) >= 1e7 ? `${r2(v / 1e7)} Cr` : Math.abs(v) >= 1e5 ? `${r2(v / 1e5)} L` : inr(Math.round(v)).replace(/\.00$/, '');

export const tile = (label: string, value: string, sub: string, tone: Tone = '') =>
  `<div class="kpi ${tone}"><span>${esc(label)}</span><b>${value}</b><em>${sub}</em></div>`;

/** a horizontal progress meter: pct 0..100 */
export function meter(label: string, pct: number, sub: string, tone: Tone = 'ok') {
  const p = Math.max(0, Math.min(100, isFinite(pct) ? pct : 0));
  return `<div class="dash-meter ${tone}"><div class="dm-top"><span>${esc(label)}</span><b>${p.toFixed(p === 100 || p === 0 ? 0 : 1)}%</b></div>
    <div class="dm-bar" role="img" aria-label="${esc(label)} ${p.toFixed(1)} percent"><i style="width:${p}%"></i></div><div class="dm-sub">${sub}</div></div>`;
}

export interface Series {
  name: string;
  color: string;
  vals: number[];
}
/**
 * Stacked bars, one per category (usually a month). `fmtVal` formats a value in the tooltips and on the axis.
 * Empty when there is nothing to draw.
 */
export function stackedBars(
  cats: string[],
  series: Series[],
  opts: { aria: string; note?: string; fmtVal?: (v: number) => string },
) {
  const f = opts.fmtVal ?? ((v: number) => String(Math.round(v)));
  const tot = cats.map((_, i) => series.reduce((a, s) => a + (s.vals[i] || 0), 0));
  const max = Math.max(0, ...tot);
  if (!cats.length || max <= 0) return '<p class="muted">Nothing to chart yet.</p>';
  const W = 760,
    H = 220,
    padL = 56,
    padB = 26,
    padT = 10;
  const y = (v: number) => padT + (1 - v / max) * (H - padT - padB);
  const gw = (W - padL - 8) / cats.length,
    bw = Math.min(30, gw - 8);
  const ticks = [0, 0.25, 0.5, 0.75, 1].map(t => max * t);
  const bars = cats
    .map((c, i) => {
      let acc = 0;
      const x0 = padL + i * gw + (gw - bw) / 2;
      const parts = series
        .map(s => {
          const v = s.vals[i] || 0;
          if (v <= 0) return '';
          const top = y(acc + v),
            h = y(acc) - top;
          acc += v;
          return `<rect x="${x0}" y="${top}" width="${bw}" height="${Math.max(0, h)}" fill="${s.color}"><title>${esc(c)} · ${esc(s.name)}: ${f(v)}</title></rect>`;
        })
        .join('');
      return `${parts}<text x="${padL + i * gw + gw / 2}" y="${H - 8}" text-anchor="middle" class="ax">${esc(c)}</text>`;
    })
    .join('');
  return `<div class="rvchart"><div class="legend">${series.map(s => `<span><i style="background:${s.color}"></i>${esc(s.name)}</span>`).join('')}${opts.note ? `<span class="muted">${esc(opts.note)}</span>` : ''}</div>
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(opts.aria)}">${ticks
      .map(
        t =>
          `<line x1="${padL}" x2="${W - 4}" y1="${y(t)}" y2="${y(t)}" class="gl"/><text x="${padL - 6}" y="${y(t) + 4}" text-anchor="end" class="ax">${f(t)}</text>`,
      )
      .join('')}${bars}</svg></div>`;
}

export interface RankRow {
  label: string;
  sub?: string;
  /** signed amount: the bar length is its size, the colour its sign */
  val: number;
  note?: string;
}
/** a ranked list with horizontal bars (largest first, as given) */
export function ranked(title: string, rows: RankRow[], empty: string, fmtVal: (v: number) => string = inr) {
  if (!rows.length)
    return `<div class="dash-rank"><h4>${esc(title)}</h4><p class="muted" style="margin:0">${esc(empty)}</p></div>`;
  const max = Math.max(...rows.map(r => Math.abs(r.val)), 1);
  return `<div class="dash-rank"><h4>${esc(title)}</h4>${rows
    .map(
      r => `<div class="dr-row"><div class="dr-l"><b>${esc(r.label || '—')}</b>${r.sub ? `<span class="muted">${esc(r.sub)}</span>` : ''}</div>
      <div class="dr-b" title="${esc(r.note || '')}"><i class="${r.val < 0 ? 'neg' : ''}" style="width:${Math.max(2, (Math.abs(r.val) / max) * 100)}%"></i></div><div class="dr-v">${fmtVal(r.val)}</div></div>`,
    )
    .join('')}</div>`;
}

/** one pass/fail / to-do line */
export const statusLine = (tone: 'ok' | 'warn' | 'bad', text: string) =>
  `<li class="dash-st ${tone}"><i>${tone === 'ok' ? '✓' : tone === 'warn' ? '!' : '✕'}</i><span>${text}</span></li>`;

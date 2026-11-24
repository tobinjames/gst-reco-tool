// Demo files for practising: a made-up company (not a real client), bundled so the page still works offline.
import type { TabId } from './wizard';

const text = import.meta.glob('./demo/*/*.{json,csv}', { query: '?raw', import: 'default', eager: true }) as Record<
  string,
  string
>;
const bin = import.meta.glob('./demo/*/*.xlsx', { query: '?inline', import: 'default', eager: true }) as Record<string, string>;

const MIME: Record<string, string> = {
  json: 'application/json',
  csv: 'text/csv',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};
const fromDataUrl = (u: string): Uint8Array => Uint8Array.from(atob(u.slice(u.indexOf(',') + 1)), c => c.charCodeAt(0));

/** the demo files of a tab, as File objects ready for the file drop */
export function demoFiles(tab: TabId): File[] {
  const dir = tab === 'annual' ? 'itc' : tab;
  const out: File[] = [];
  const mk = (path: string, data: BlobPart) => {
    const name = path.split('/').pop()!;
    out.push(new File([data], name, { type: MIME[name.split('.').pop()!] }));
  };
  for (const [p, t] of Object.entries(text)) if (p.startsWith(`./demo/${dir}/`)) mk(p, t);
  for (const [p, u] of Object.entries(bin)) if (p.startsWith(`./demo/${dir}/`)) mk(p, fromDataUrl(u) as unknown as BlobPart);
  return out;
}

// Workpaper memory: every workbook the tool writes carries two hidden sheets, so that dropping the saved workpaper back in
// (alone) brings everything back:
//   _JJState  your settings and choices (matching settings, register column mappings, supplier / customer names, revenue
//             ledger classes and reasons, flow account types, moved lines and balances), as JSON split over cells A2, A3...
//   _JJFiles  the source files the workpaper was built from (returns, portal ledgers, registers, GL dumps), as a zip in base64
//             split over cells A2, A3...
// Returns and portal ledgers never change once filed, so they stay; registers, GL dumps and customer lists can be replaced
// by dropping a newer one with the workpaper (see main.ts).
import JSZip from 'jszip';
import { Reader, rc } from './xlsx';
import { store, xmlEsc } from './util';
import type { TabId } from './settings';

export const STATE_SHEET = '_JJState';
export const FILES_SHEET = '_JJFiles';
const CHUNK = 32000; // Excel keeps up to 32,767 characters in a cell

export interface SavedFile {
  /** path inside the _JJFiles zip */
  p: string;
  name: string;
  kind: string;
  size: number;
}
export interface SavedState {
  v: 1;
  tab: TabId;
  gstin: string;
  fy: string;
  legal?: string;
  savedAt: string;
  /** browser settings (store keys) at the time of the build */
  keys: Record<string, any>;
  files: SavedFile[];
}

// ------------------------------------------------------------------ what is saved
/** store keys of a tab for a client + year (besides the ones main.ts adds: register mappings, supplier names) */
export function tabKeys(tab: TabId, gstin: string, fy: string): string[] {
  const g = gstin || '-';
  const k = [`jjset:${tab}`];
  if (tab === 'flow') k.push(`jjflow:acc:${g}`, `jjflow:${g}:${fy}`);
  if (tab === 'revenue') k.push(`jjrev:${g}:${fy}`, `jjrev:cls:${g}`, `jjrev:cust:${g}`);
  return k;
}
/** keys whose value is a map of choices (name → choice): restoring adds to what the browser has, the workpaper winning */
const isMerged = (k: string) =>
  k === 'jjalias' || k.startsWith('jjflow:acc:') || k.startsWith('jjrev:cls:') || k.startsWith('jjrev:cust:');

export interface PackFile {
  name: string;
  kind: string;
  bytes: Uint8Array;
}
/** main.ts supplies the shared parts: register mappings, supplier names, the files in use */
let provider: (tab: TabId) => { keys: Record<string, any>; files: PackFile[]; legal?: string } = () => ({ keys: {}, files: [] });
export function setStateProvider(fn: typeof provider) {
  provider = fn;
}

export function packState(tab: TabId, gstin: string, fy: string, savedAt: string): { state: SavedState; files: PackFile[] } {
  const extra = provider(tab);
  const keys: Record<string, any> = {};
  for (const k of tabKeys(tab, gstin, fy)) {
    const v = store.get(k);
    if (v !== null && v !== undefined) keys[k] = v;
  }
  Object.assign(keys, extra.keys);
  const files = extra.files;
  return {
    state: {
      v: 1,
      tab,
      gstin,
      fy,
      legal: extra.legal,
      savedAt,
      keys,
      files: files.map((f, i) => ({
        p: `${String(i + 1).padStart(3, '0')}_${f.name.replace(/[\\/:*?"<>|]/g, '_')}`,
        name: f.name,
        kind: f.kind,
        size: f.bytes.length,
      })),
    },
    files,
  };
}

// ------------------------------------------------------------------ writing the hidden sheets
const b64 = (u: Uint8Array) => {
  let s = '';
  for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode(...u.subarray(i, i + 0x8000));
  return btoa(s);
};
const unb64 = (s: string) => {
  const bin = atob(s);
  const u = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return u;
};
const chunks = (s: string) => {
  const out: string[] = [];
  for (let i = 0; i < s.length; i += CHUNK) out.push(s.slice(i, i + CHUNK));
  return out.length ? out : [''];
};
function sheetXml(note: string, parts: string[]) {
  const cell = (r: number, t: string) =>
    `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(t)}</t></is></c></row>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${cell(1, note)}${parts.map((p, i) => cell(i + 2, p)).join('')}</sheetData></worksheet>`;
}

/** add (or replace) the two hidden sheets in a finished .xlsx */
export async function addStateSheets(bytes: Uint8Array, pack: { state: SavedState; files: PackFile[] }): Promise<Uint8Array> {
  const z = await JSZip.loadAsync(bytes);
  const inner = new JSZip();
  pack.state.files.forEach((f, i) => inner.file(f.p, pack.files[i].bytes));
  const filesB64 = pack.files.length ? b64(await inner.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })) : '';
  let wb = await z.file('xl/workbook.xml')!.async('string');
  let rels = await z.file('xl/_rels/workbook.xml.rels')!.async('string');
  let ct = await z.file('[Content_Types].xml')!.async('string');
  const relTarget = (rid: string) =>
    [...rels.matchAll(/<Relationship\b[^>]*>/g)]
      .map(m => m[0])
      .find(r => r.includes(`Id="${rid}"`))
      ?.match(/Target="([^"]+)"/)?.[1];
  const sheets: [string, string, string][] = [
    [
      STATE_SHEET,
      'Saved by the James & James GST Reco Loader: your settings and choices, read back when this workbook is dropped into the Loader. Do not edit.',
      JSON.stringify(pack.state),
    ],
    [
      FILES_SHEET,
      'Saved by the James & James GST Reco Loader: the source files of this workpaper (zip, base64), so it can be dropped back alone. Do not edit.',
      filesB64,
    ],
  ];
  for (const [name, note, data] of sheets) {
    const xml = sheetXml(note, chunks(data));
    const tag = [...wb.matchAll(/<sheet\b[^>]*\/>/g)].map(m => m[0]).find(t => t.includes(`name="${name}"`));
    if (tag) {
      const tgt = relTarget(tag.match(/r:id="([^"]+)"/)![1])!;
      z.file(tgt.startsWith('/') ? tgt.slice(1) : 'xl/' + tgt, xml);
      continue;
    }
    const ids = [...wb.matchAll(/sheetId="(\d+)"/g)].map(m => +m[1]);
    const sid = Math.max(0, ...ids) + 1;
    let n = 1;
    while (rels.includes(`Id="rIdJJ${n}"`)) n++;
    const rid = `rIdJJ${n}`;
    let fn = 1;
    while (z.file(`xl/worksheets/jj${fn}.xml`)) fn++;
    const path = `worksheets/jj${fn}.xml`;
    z.file('xl/' + path, xml);
    wb = wb.replace('</sheets>', `<sheet name="${name}" sheetId="${sid}" state="hidden" r:id="${rid}"/></sheets>`);
    rels = rels.replace(
      '</Relationships>',
      `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${path}"/></Relationships>`,
    );
    ct = ct.replace(
      '</Types>',
      `<Override PartName="/xl/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    );
  }
  z.file('xl/workbook.xml', wb);
  z.file('xl/_rels/workbook.xml.rels', rels);
  z.file('[Content_Types].xml', ct);
  return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ------------------------------------------------------------------ reading them back
async function joined(rd: Reader, sheet: string) {
  const parts: [number, string][] = [];
  for (const [ref, c] of await rd.sheet(sheet)) {
    const [r, col] = rc(ref);
    if (col === 1 && r >= 2 && c.v !== null && c.v !== undefined) parts.push([r, String(c.v)]);
  }
  return parts
    .sort((a, b) => a[0] - b[0])
    .map(p => p[1])
    .join('');
}
export async function readState(rd: Reader): Promise<SavedState | null> {
  if (!rd.has(STATE_SHEET)) return null;
  try {
    const s = JSON.parse(await joined(rd, STATE_SHEET));
    return s && s.v === 1 && s.keys ? s : null;
  } catch {
    return null;
  }
}
/** the source files saved in the workpaper */
export async function readSavedFiles(rd: Reader, st: SavedState): Promise<PackFile[]> {
  if (!rd.has(FILES_SHEET) || !st.files?.length) return [];
  const data = await joined(rd, FILES_SHEET);
  if (!data) return [];
  const z = await JSZip.loadAsync(unb64(data));
  const out: PackFile[] = [];
  for (const f of st.files) {
    const e = z.file(f.p);
    if (e) out.push({ name: f.name, kind: f.kind, bytes: await e.async('uint8array') });
  }
  return out;
}

// ------------------------------------------------------------------ applying (with undo)
export interface Restored {
  /** the workpaper's file name */
  from: string;
  state: SavedState;
  /** what the browser had before, to put back (null = nothing) */
  undo: Record<string, any>;
  lines: string[];
  /** source files added from the workpaper */
  nFiles?: number;
}
const n = (o: any) => (o && typeof o === 'object' ? Object.keys(o).length : 0);
/** what a set of saved keys holds, in words */
export function describeKeys(keys: Record<string, any>): string[] {
  const out: string[] = [];
  const maps = Object.keys(keys).filter(k => k.startsWith('jjmap:')).length;
  for (const [k, v] of Object.entries(keys)) {
    if (k.startsWith('jjset:')) out.push('Matching settings');
    else if (k === 'jjalias') {
      const ok = Object.values(v || {}).filter(Boolean).length;
      if (n(v)) out.push(`${ok} supplier name(s) matched to a GSTIN${n(v) > ok ? `, ${n(v) - ok} marked "none of these"` : ''}`);
    } else if (k.startsWith('jjflow:acc:')) {
      if (n(v)) out.push(`${n(v)} books account type(s) / head(s)`);
    } else if (k.startsWith('jjflow:')) {
      const bal = n(v?.bal?.books) + n(v?.bal?.portal);
      const bits = [
        n(v?.ov) ? `${n(v.ov)} line(s) moved or changed` : '',
        bal ? `${bal} balance(s) entered` : '',
        v?.from && v?.to ? `months ${v.from} to ${v.to}` : '',
        v?.audit?.length ? `${v.audit.length} Audit entr(ies)` : '',
      ].filter(Boolean);
      if (bits.length) out.push('Ledger flow: ' + bits.join(', '));
    } else if (k.startsWith('jjrev:cls:')) {
      if (n(v)) out.push(`${n(v)} revenue ledger class(es)`);
    } else if (k.startsWith('jjrev:cust:')) {
      if (n(v)) out.push(`${n(v)} customer name match(es)`);
    } else if (k.startsWith('jjrev:')) {
      const bits = [
        n(v?.user?.reasons) ? `${n(v.user.reasons)} reason(s) / remark(s)` : '',
        n(v?.user?.b2c) ? `${n(v.user.b2c)} B2CS month(s)` : '',
        v?.audit?.length ? `${v.audit.length} Audit entr(ies)` : '',
      ].filter(Boolean);
      if (bits.length) out.push('Revenue reco: ' + bits.join(', '));
    }
  }
  if (maps) out.push(`Column mapping of ${maps} register(s)`);
  return out;
}
export function applyState(from: string, st: SavedState): Restored {
  const undo: Record<string, any> = {};
  for (const [k, v] of Object.entries(st.keys)) {
    const cur = store.get(k);
    undo[k] = cur;
    if (isMerged(k) && cur && typeof cur === 'object' && v && typeof v === 'object') store.set(k, { ...cur, ...v });
    else if (k.startsWith('jjrev:') && !k.startsWith('jjrev:cls:') && !k.startsWith('jjrev:cust:') && cur?.imported)
      // keep the list of workbooks already read back in this browser
      store.set(k, { ...v, imported: [...new Set([...(v?.imported || []), ...cur.imported])] });
    else store.set(k, v);
  }
  return { from, state: st, undo, lines: describeKeys(st.keys) };
}
export function undoState(r: Restored) {
  for (const [k, v] of Object.entries(r.undo)) {
    if (v === null || v === undefined) store.del(k);
    else store.set(k, v);
  }
}

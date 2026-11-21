// GSTR-2B monthly JSON → RAW_2B rows
import { Amt, RawRow, Warn, amt2b, fyOfDate, fyOfPeriod, neg, n, r2 } from './util';

export function is2b(j: any) {
  const d = j?.data ?? j;
  return !!(d && d.rtnprd && (d.docdata || d.itcsumm));
}

/** column S note on a 2B row that a later amendment replaced (not counted anywhere) */
export const SUPERSEDED = 'Superseded by';
/** a RAW_2B row as an array (A key, B-F amounts, G..S info): true when a later amendment replaced it */
export const isSuperseded = (r: any[]) => String(r?.[18] ?? '').startsWith(SUPERSEDED);
const docNo = (s: any) =>
  String(s ?? '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
/** one key per document of a 2B: the amendment of a document carries the key of its original (oinum / ontnum) */
const docSig = (d: any) => {
  const out = new Set<string>();
  const dd = d.docdata || {};
  for (const s of [...(dd.b2b || []), ...(dd.ecom || [])]) for (const i of s.inv || []) out.add(`I|${s.ctin}|${docNo(i.inum)}`);
  for (const s of dd.cdnr || []) for (const nt of s.nt || []) out.add(`N|${s.ctin}|${docNo(nt.ntnum ?? nt.nt_num)}`);
  for (const b of dd.impg || []) out.add(`E|${docNo(b.boenum)}`);
  return out;
};

export function process2b(files: { name: string; json: any }[], cy: string, py: string, ny: string, warns: Warn[]) {
  const rows: RawRow[] = [];
  const periods = new Set<string>();
  /** documents already read per period: a second file of the same period with the same documents is a duplicate */
  const seenDocs = new Map<string, Set<string>>();
  // amendments are applied in period order (an earlier original first, then the amendment that replaces it)
  const per = (f: { json: any }) => {
    const p = String((f.json.data ?? f.json).rtnprd ?? '');
    return p.slice(2) + p.slice(0, 2);
  };
  files = [...files].sort((a, b) => per(a).localeCompare(per(b)));
  /** live document → its row; amended documents are looked up by their original number */
  const live = new Map<string, RawRow>();
  for (const f of files) {
    const d = f.json.data ?? f.json;
    const prd = String(d.rtnprd),
      bfy = fyOfPeriod(prd);
    const sig = docSig(d);
    if (periods.has(prd)) {
      const prev = seenDocs.get(prd)!;
      const same = [...sig].filter(k => prev.has(k)).length;
      // large 2Bs come in several files with different documents; the same documents again = the same 2B loaded twice
      if (sig.size && same / sig.size >= 0.5) {
        warns.push({
          level: 'error',
          file: f.name,
          msg: `2B for ${prd.slice(0, 2)}/${prd.slice(2)} is already loaded (${same} of ${sig.size} documents are the same): this file is left out`,
        });
        continue;
      }
      sig.forEach(k => prev.add(k));
    } else seenDocs.set(prd, sig);
    periods.add(prd);
    if (bfy !== cy && bfy !== ny) {
      warns.push({ level: 'info', file: f.name, msg: `2B period ${prd} is outside ${cy} / ${ny}; rows kept but not counted` });
    }
    // within next FY, only the April-October 2B counts for this FY's 8A (sec 16(4): returns filed up to 30 November;
    // ICAI TG on GSTR-9, Dec 2025, FAQ 3)
    const nyMonth = +prd.slice(0, 2);
    const nyInWindow = bfy === ny && nyMonth >= 4 && nyMonth <= 10;
    const dd = d.docdata || {};
    const emit = (
      section: string,
      sup: any,
      doc: any,
      docNo: string,
      docDt: string,
      dtype: string,
      amt: Amt,
      rcm: string,
      avl: string,
      isImport = false,
      orig?: string,
    ) => {
      const dfy = fyOfDate(docDt);
      let key = '';
      const counted = avl === 'Y' && rcm !== 'Y';
      if (counted && isImport) {
        if (dfy === cy && bfy === cy) key = '8G';
        else if (dfy === cy && nyInWindow) key = '8G+8H1';
      } else if (counted && ['B2B', 'B2BA', 'CDNR', 'CDNRA', 'ECOM', 'ECOMA'].includes(section)) {
        if (dfy === cy && bfy === cy) key = '8A';
        else if (dfy === cy && nyInWindow) key = '8A+8C';
        else if (dfy === py && bfy === cy) key = '6A1';
      }
      const row: RawRow = {
        key,
        amt,
        info: [
          '2B',
          bfy,
          prd,
          section,
          sup?.ctin ?? '',
          sup?.trdnm ?? '',
          dtype,
          docNo,
          docDt,
          dfy,
          avl,
          rcm,
          doc?.imsStatus ? `IMS:${doc.imsStatus}` : '',
        ],
      };
      rows.push(row);
      supersede(row, section, sup?.ctin ?? '', dtype, docNo, orig);
    };
    /** GSTN (GSTR-2B FAQ 8): amendment tables carry the full revised values of the document, the summary only the
     * difference. So the original (or an earlier amendment) of an amended document is set aside and only the latest
     * version counts. Without the original loaded the amendment's full value is not counted (the original is in a 2B
     * that is not loaded and was counted there). */
    const supersede = (row: RawRow, section: string, ctin: string, dtype: string, no: string, orig?: string) => {
      const fam = section.replace(/A$/, '').replace(/^IMPGSEZ$/, 'IMPG');
      const kind = fam === 'CDNR' ? 'N' : fam === 'IMPG' ? 'E' : 'I';
      const k = (n: any) => `${fam}|${kind === 'E' ? '' : ctin}|${kind}|${docNo(n)}`;
      const own = k(no);
      const amended = section.endsWith('A') || (fam === 'IMPG' && orig !== undefined);
      if (amended) {
        const ok = k(orig || no);
        const old = live.get(ok) ?? live.get(own);
        const p = `${prd.slice(0, 2)}/${prd.slice(2)}`;
        if (old && old !== row) {
          old.key = '';
          old.info[12] = `${SUPERSEDED} ${section} ${no} in ${p}${old.info[12] ? '; ' + old.info[12] : ''}`;
          live.delete(ok);
          live.delete(own);
        } else if (section.endsWith('A')) {
          if (row.key)
            warns.push({
              level: 'warn',
              file: f.name,
              msg: `${section} ${no}${orig && docNo(orig) !== docNo(no) ? ` (original ${orig})` : ''} of ${ctin || 'import'}: the original document is not in the 2B files loaded, so the amended value is not counted in ${row.key} (load the 2B of the original month)`,
            });
          row.key = '';
          row.info[12] = `Original not loaded: not counted${row.info[12] ? '; ' + row.info[12] : ''}`;
        }
      }
      live.set(own, row);
    };
    for (const [sec, arr] of [
      ['B2B', dd.b2b],
      ['B2BA', dd.b2ba],
      // documents from suppliers selling through an e-commerce operator who pays tax u/s 9(5): same shape as B2B (format v4.x)
      ['ECOM', dd.ecom],
      ['ECOMA', dd.ecoma],
    ] as const)
      for (const s of arr || [])
        for (const inv of s.inv || [])
          emit(sec, s, inv, inv.inum, inv.dt, 'INV', amt2b(inv), inv.rev ?? 'N', inv.itcavl ?? 'Y', false, inv.oinum);
    for (const [sec, arr] of [
      ['CDNR', dd.cdnr],
      ['CDNRA', dd.cdnra],
    ] as const)
      for (const s of arr || [])
        for (const nt of s.nt || []) {
          const cn = (nt.typ ?? nt.ntty) === 'C';
          const a = amt2b(nt);
          emit(
            sec,
            s,
            nt,
            nt.ntnum ?? nt.nt_num,
            nt.dt ?? nt.nt_dt,
            cn ? 'CN' : 'DN',
            cn ? neg(a) : a,
            nt.rev ?? 'N',
            nt.itcavl ?? 'Y',
            false,
            nt.ontnum ?? nt.ont_num,
          );
        }
    for (const [sec, arr] of [
      ['ISD', dd.isd],
      ['ISDA', dd.isda],
    ] as const)
      for (const s of arr || [])
        for (const doc of s.doclist ?? [s]) {
          const a: Amt = [0, r2(n(doc.cgst)), r2(n(doc.sgst)), r2(n(doc.igst)), r2(n(doc.cess))];
          rows.push({
            key: '',
            amt: doc.doctyp === 'C' || doc.doctyp === 'ISDC' ? neg(a) : a, // ISD credit notes: 'ISDC' in the GSTN format
            info: [
              '2B',
              bfy,
              prd,
              sec,
              s.ctin ?? '',
              s.trdnm ?? '',
              doc.doctyp ?? '',
              doc.docnum ?? '',
              doc.docdt ?? '',
              fyOfDate(doc.docdt),
              doc.itcelg ?? '',
              'N',
              'ISD - from 3B 4A(4)',
            ],
          });
        }
    // IMPG is a flat list of bills of entry; IMPGSEZ groups them by SEZ supplier: [{ ctin, trdnm, boe: [...] }]
    for (const [sec, arr] of [
      ['IMPG', dd.impg],
      ['IMPGSEZ', dd.impgsez],
    ] as const)
      for (const doc of (arr || []).flatMap((s: any) =>
        Array.isArray(s.boe) ? s.boe.map((b: any) => ({ ...b, ctin: s.ctin, trdnm: s.trdnm })) : [s],
      )) {
        const a: Amt = [r2(n(doc.txval)), 0, 0, r2(n(doc.igst)), r2(n(doc.cess))];
        emit(
          sec,
          { ctin: doc.ctin ?? doc.portcd ?? '', trdnm: doc.trdnm ?? '' },
          doc,
          doc.boenum,
          doc.boedt,
          'BOE',
          a,
          'N',
          'Y',
          true,
          // a bill of entry amended at ICEGATE comes again with isamd = Y and replaces the earlier one
          doc.isamd === 'Y' ? String(doc.boenum) : undefined,
        );
      }
    for (const sec of Object.keys(dd))
      if (!['b2b', 'b2ba', 'cdnr', 'cdnra', 'isd', 'isda', 'impg', 'impgsez', 'ecom', 'ecoma'].includes(sec))
        warns.push({
          level: sec === 'itcrev' && dd[sec]?.length ? 'warn' : 'info',
          file: f.name,
          msg:
            sec === 'itcrev'
              ? `2B "ITC reversal" section (IMS) is not read yet: check the reversal against GSTR-3B 4B by hand`
              : `2B section "${sec}" not used`,
        });
  }
  return { rows, periods: [...periods].sort() };
}

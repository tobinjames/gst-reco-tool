// GSTR-2B monthly JSON → RAW_2B rows
import { Amt, RawRow, Warn, amt2b, fyOfDate, fyOfPeriod, neg, n, r2 } from './util';

export function is2b(j: any) {
  const d = j?.data ?? j;
  return !!(d && d.rtnprd && (d.docdata || d.itcsumm));
}

export function process2b(files: { name: string; json: any }[], cy: string, py: string, ny: string, warns: Warn[]) {
  const rows: RawRow[] = [];
  const periods = new Set<string>();
  for (const f of files) {
    const d = f.json.data ?? f.json;
    const prd = String(d.rtnprd),
      bfy = fyOfPeriod(prd);
    if (periods.has(prd)) warns.push({ level: 'warn', file: f.name, msg: `2B for ${prd} loaded more than once` });
    periods.add(prd);
    if (bfy !== cy && bfy !== ny) {
      warns.push({ level: 'info', file: f.name, msg: `2B period ${prd} is outside ${cy} / ${ny}; rows kept but not counted` });
    }
    // within next FY, only Apr-Nov 2B matters for this FY's ITC (sec 16(4) - 30 Nov)
    const nyMonth = +prd.slice(0, 2);
    const nyInWindow = bfy === ny && nyMonth >= 4 && nyMonth <= 11;
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
    ) => {
      const dfy = fyOfDate(docDt);
      let key = '';
      const counted = avl === 'Y' && rcm !== 'Y';
      if (counted && isImport) {
        if (dfy === cy && bfy === cy) key = '8G';
        else if (dfy === cy && nyInWindow) key = '8G+8H1';
      } else if (counted && ['B2B', 'B2BA', 'CDNR', 'CDNRA'].includes(section)) {
        if (dfy === cy && bfy === cy) key = '8A';
        else if (dfy === cy && nyInWindow) key = '8A+8C';
        else if (dfy === py && bfy === cy) key = '6A1';
      }
      rows.push({
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
      });
    };
    for (const [sec, arr] of [
      ['B2B', dd.b2b],
      ['B2BA', dd.b2ba],
    ] as const)
      for (const s of arr || [])
        for (const inv of s.inv || []) emit(sec, s, inv, inv.inum, inv.dt, 'INV', amt2b(inv), inv.rev ?? 'N', inv.itcavl ?? 'Y');
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
            amt: doc.doctyp === 'C' ? neg(a) : a,
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
    for (const [sec, arr] of [
      ['IMPG', dd.impg],
      ['IMPGSEZ', dd.impgsez],
    ] as const)
      for (const doc of arr || []) {
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
          doc.isamd === 'Y' ? 'Y' : 'Y',
          true,
        );
      }
    for (const sec of Object.keys(dd))
      if (!['b2b', 'b2ba', 'cdnr', 'cdnra', 'isd', 'isda', 'impg', 'impgsez'].includes(sec))
        warns.push({ level: 'info', file: f.name, msg: `2B section "${sec}" not used` });
  }
  return { rows, periods: [...periods].sort() };
}

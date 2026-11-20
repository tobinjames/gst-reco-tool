// GSTR-1 / IFF / GSTR-1A JSON (portal "download JSON" of filed return) → RAW_GSTR1 rows
import { Amt, RawRow, Warn, add, amtOf, fyOfDate, fyOfPeriod, isZero, n, neg, r2, zero } from './util';

export interface G1File {
  name: string;
  json: any;
}

interface Doc {
  file: string;
  rfy: string;
  period: string;
  section: string;
  dtype: string;
  num: string;
  date: string;
  dfy: string;
  party: string;
  pos: string;
  rate: number | string;
  rcm: string;
  amt: Amt;
  family: string; // B2B / B2CL / EXP / CDN / CDNUR / B2CS / AT / TXP / NIL / ECO
  invTyp?: string; // R, SEWP, SEWOP, DE ; exp WPAY/WOPAY ; cdnur B2CL/EXPWP/EXPWOP
  ntty?: string; // C / D
  amended?: boolean;
  orig?: string;
  omon?: string;
}

const sumItms = (itms: any[] | undefined, nested = true): Amt =>
  (itms || []).reduce((a: Amt, it: any) => add(a, amtOf(nested ? (it.itm_det ?? it) : it)), zero());
const rateOf = (itms: any[] | undefined) => {
  const rs = Array.from(new Set((itms || []).map((it: any) => n((it.itm_det ?? it).rt))));
  return rs.length === 1 ? rs[0] : rs.join('/');
};

export function isGstr1(j: any) {
  return (
    j &&
    typeof j === 'object' &&
    j.fp &&
    j.gstin &&
    ['b2b', 'b2cl', 'b2cs', 'exp', 'cdnr', 'cdnur', 'hsn', 'nil', 'at', 'txpd', 'doc_issue', 'b2ba', 'cdnra', 'supeco'].some(
      k => k in j,
    )
  );
}

/** a section as a list: GSTN format v5.0 (2026) splits AT / ATA / TXPD / TXPDA by recipient type,
 * e.g. { at_b2b: [...], at_b2c: [...], at_exp: [...], at_ecom: { at_ecomurp2b: [...], ... } } */
const listOf = (v: any): any[] => (Array.isArray(v) ? v : v && typeof v === 'object' ? Object.values(v).flatMap(listOf) : []);

function extract(f: G1File): { docs: Doc[]; hsn: any[] } {
  const j = f.json,
    period = String(j.fp),
    rfy = fyOfPeriod(period),
    docs: Doc[] = [];
  const base = { file: f.name, rfy, period };
  const push = (d: Partial<Doc>) =>
    docs.push({
      ...base,
      section: '',
      dtype: 'INV',
      num: '',
      date: '',
      dfy: '',
      party: '',
      pos: '',
      rate: '',
      rcm: 'N',
      amt: zero(),
      family: '',
      ...d,
    } as Doc);

  for (const [sec, amended] of [
    ['b2b', false],
    ['b2ba', true],
  ] as const)
    for (const p of j[sec] || [])
      for (const inv of p.inv || [])
        push({
          section: sec.toUpperCase(),
          family: 'B2B',
          num: inv.inum,
          date: inv.idt,
          dfy: fyOfDate(inv.idt),
          party: p.ctin,
          pos: inv.pos,
          rcm: inv.rchrg || 'N',
          invTyp: inv.inv_typ || 'R',
          rate: rateOf(inv.itms),
          amt: sumItms(inv.itms),
          amended,
          orig: inv.oinum,
        });
  for (const [sec, amended] of [
    ['b2cl', false],
    ['b2cla', true],
  ] as const)
    for (const p of j[sec] || [])
      for (const inv of p.inv || [])
        push({
          section: sec.toUpperCase(),
          family: 'B2CL',
          num: inv.inum,
          date: inv.idt,
          dfy: fyOfDate(inv.idt),
          pos: p.pos,
          rate: rateOf(inv.itms),
          amt: sumItms(inv.itms),
          amended,
          orig: inv.oinum,
        });
  for (const r of j.b2cs || [])
    push({
      section: 'B2CS',
      family: 'B2CS',
      dtype: 'SUM',
      dfy: rfy,
      pos: r.pos,
      rate: n(r.rt),
      invTyp: r.typ,
      amt: amtOf(r),
      num: `${r.pos}/${r.rt}/${r.typ}`,
    });
  for (const r of j.b2csa || [])
    for (const it of r.itms || [])
      push({
        section: 'B2CSA',
        family: 'B2CS',
        dtype: 'SUM',
        dfy: fyOfPeriod(r.omon),
        pos: r.pos,
        rate: n(it.rt),
        invTyp: r.typ,
        amt: amtOf(it),
        num: `${r.pos}/${it.rt}/${r.typ}`,
        amended: true,
        omon: r.omon,
      });
  for (const [sec, amended] of [
    ['exp', false],
    ['expa', true],
  ] as const)
    for (const p of j[sec] || [])
      for (const inv of p.inv || [])
        push({
          section: sec.toUpperCase(),
          family: 'EXP',
          num: inv.inum,
          date: inv.idt,
          dfy: fyOfDate(inv.idt),
          invTyp: p.exp_typ,
          rate: rateOf(inv.itms),
          amt: sumItms(inv.itms, false),
          amended,
          orig: inv.oinum,
        });
  for (const [sec, amended] of [
    ['cdnr', false],
    ['cdnra', true],
  ] as const)
    for (const p of j[sec] || [])
      for (const nt of p.nt || [])
        push({
          section: sec.toUpperCase(),
          family: 'CDN',
          dtype: nt.ntty === 'D' ? 'DN' : 'CN',
          ntty: nt.ntty,
          num: nt.nt_num,
          date: nt.nt_dt,
          dfy: fyOfDate(nt.nt_dt),
          party: p.ctin,
          pos: nt.pos,
          rcm: nt.rchrg || 'N',
          invTyp: nt.inv_typ || 'R',
          rate: rateOf(nt.itms),
          amt: sumItms(nt.itms),
          amended,
          orig: nt.ont_num,
        });
  for (const [sec, amended] of [
    ['cdnur', false],
    ['cdnura', true],
  ] as const)
    for (const nt of j[sec] || [])
      push({
        section: sec.toUpperCase(),
        family: 'CDNUR',
        dtype: nt.ntty === 'D' ? 'DN' : 'CN',
        ntty: nt.ntty,
        num: nt.nt_num,
        date: nt.nt_dt,
        dfy: fyOfDate(nt.nt_dt),
        pos: nt.pos,
        invTyp: nt.typ,
        rate: rateOf(nt.itms),
        amt: sumItms(nt.itms),
        amended,
        orig: nt.ont_num,
      });
  for (const [sec, fam, amended] of [
    ['at', 'AT', false],
    ['ata', 'AT', true],
    ['txpd', 'TXP', false],
    ['txpda', 'TXP', true],
  ] as const)
    for (const p of listOf(j[sec]))
      for (const it of p.itms || [])
        push({
          section: sec.toUpperCase(),
          family: fam,
          dtype: fam === 'AT' ? 'ADV' : 'ADJ',
          dfy: amended ? fyOfPeriod(p.omon) : rfy,
          pos: p.pos,
          rate: n(it.rt),
          amt: amtOf(it, 'ad_amt'),
          num: `${p.pos}/${it.rt}`,
          amended,
          omon: p.omon,
        });
  for (const it of j.nil?.inv || []) {
    for (const [k, lab] of [
      ['nil_amt', 'NIL'],
      ['expt_amt', 'EXEMPT'],
      ['ngsup_amt', 'NONGST'],
    ] as const)
      if (n(it[k]))
        push({ section: 'NIL', family: 'NIL', dtype: lab, num: it.sply_ty, dfy: rfy, amt: [r2(n(it[k])), 0, 0, 0, 0] });
  }
  for (const r of j.supeco?.paytx || [])
    push({
      section: 'SUPECO-PAYTX',
      family: 'ECO',
      dtype: 'SUM',
      party: r.etin,
      dfy: rfy,
      amt: [r2(n(r.suppval)), r2(n(r.cgst)), r2(n(r.sgst)), r2(n(r.igst)), r2(n(r.cess))],
    });
  for (const r of j.supeco?.clttx || [])
    push({
      section: 'SUPECO-CLTTX',
      family: 'ECO-TCS',
      dtype: 'SUM',
      party: r.etin,
      dfy: rfy,
      amt: [r2(n(r.suppval)), 0, 0, 0, 0],
    });

  const hsn: any[] = [];
  const h = j.hsn || {};
  for (const [arr, tag] of [
    [h.data, ''],
    [h.hsn_b2b, 'B2B'],
    [h.hsn_b2c, 'B2C'],
  ] as const)
    for (const r of arr || []) hsn.push({ ...r, period, tag, file: f.name });
  return { docs, hsn };
}

/** GSTR-9 key for an (un-amended view of a) document, for the selected FY */
function keyFor(d: Doc, cy: string, ny: string): string {
  const credit = d.dtype === 'CN';
  if (d.rfy === cy) {
    switch (d.family) {
      case 'B2B':
        if (d.rcm === 'Y') return '5C';
        return ({ R: '4B', SEWP: '4D', SEWOP: '5B', DE: '4E', CBW: '4B' } as any)[d.invTyp || 'R'] || '4B';
      case 'B2CL':
      case 'B2CS':
        return '4A';
      case 'EXP':
        return d.invTyp === 'WOPAY' ? '5A' : '4C';
      case 'CDN':
        if (d.rcm === 'Y' || d.invTyp === 'SEWOP') return credit ? '5H' : '5I';
        return credit ? '4I' : '4J';
      case 'CDNUR':
        if (d.invTyp === 'EXPWOP') return credit ? '5H' : '5I';
        if (d.invTyp === 'B2CL') return '4A'; // netted into 4A (sign handled by caller)
        return credit ? '4I' : '4J';
      case 'AT':
      case 'TXP':
        return '4F';
      case 'NIL':
        return d.dtype === 'NIL' ? '5E' : d.dtype === 'EXEMPT' ? '5D' : '5F';
      case 'ECO':
        return '5C1';
      default:
        return '';
    }
  }
  // Credit notes in next-FY returns dated Apr–Nov of next FY may relate to this FY's supplies (Table 11, per ICAI TG) - cannot be
  // tied to an invoice automatically, so they are tagged for review and not counted
  if (d.rfy === ny && d.dfy === ny && credit && (d.family === 'CDN' || d.family === 'CDNUR') && !d.amended) {
    const m = +(d.date || '').split(/[-/]/)[1];
    if (m >= 4 && m <= 11) return '11-REVIEW';
  }
  if (d.rfy === ny && d.dfy === cy && d.dfy) {
    if (['B2CS', 'AT', 'TXP', 'NIL', 'ECO', 'ECO-TCS'].includes(d.family)) return '';
    return credit ? '11' : '10';
  }
  return '';
}

export function processGstr1(files: G1File[], cy: string, ny: string, warns: Warn[]) {
  const all: Doc[] = [];
  const hsnRows: any[] = [];
  const seen = new Set<string>();
  for (const f of files) {
    const fp = `${f.json.gstin}|${f.json.fp}|${f.name.replace(/_\d+\.json$/i, '')}`;
    const { docs, hsn } = extract(f);
    all.push(...docs);
    if (Object.values(f.json.supecoa ?? {}).some((v: any) => Array.isArray(v) && v.length))
      warns.push({
        level: 'warn',
        file: f.name,
        msg: 'GSTR-1 Table 14A (amendments of supplies through e-commerce operators) is not read yet; check it by hand',
      });
    for (const sec of ['ecom', 'ecoma'])
      if (f.json[sec] && Object.values(f.json[sec]).some((v: any) => Array.isArray(v) && v.length))
        warns.push({
          level: 'warn',
          file: f.name,
          msg: `GSTR-1 section "${sec}" (Table 15: supplies on which an e-commerce operator pays tax u/s 9(5)) is not read yet; check Tables 4 / 5 by hand`,
        });
    if (fyOfPeriod(String(f.json.fp)) === cy) hsnRows.push(...hsn);
    seen.add(fp);
  }
  // index originals for amendment differentials
  const orig = new Map<string, Doc>();
  for (const d of all)
    if (!d.amended) {
      if (d.family === 'B2CS') orig.set(`B2CS|${d.period}|${d.num}`, d);
      else if (d.family === 'AT' || d.family === 'TXP') orig.set(`${d.family}|${d.period}|${d.num}`, d);
      else if (d.num) orig.set(`${d.family}|${d.party}|${d.num}`, d);
    }
  // later amendments supersede earlier ones of the same doc: process in period order and keep the latest value as "original"
  const byPeriod = (a: Doc, b: Doc) =>
    (a.period.slice(2) + a.period.slice(0, 2)).localeCompare(b.period.slice(2) + b.period.slice(0, 2));
  all.sort(byPeriod);

  const rows: RawRow[] = [];
  for (const d of all) {
    let key = keyFor(d, cy, ny);
    let amt = d.amt;
    let note = '';
    if (d.amended) {
      const ok =
        d.family === 'B2CS'
          ? `B2CS|${d.omon}|${d.num}`
          : d.family === 'AT' || d.family === 'TXP'
            ? `${d.family}|${d.omon}|${d.num}`
            : `${d.family}|${d.party}|${d.orig}`;
      const o = orig.get(ok);
      if (o) {
        amt = add(d.amt, o.amt, -1);
        note = `diff vs original in ${o.period}`;
      } else {
        note = 'ORIGINAL NOT LOADED - full amended value shown, not counted';
        key = '';
        warns.push({
          level: 'warn',
          file: d.file,
          msg: `${d.section} ${d.orig || d.num}: original document not found in loaded GSTR-1 files; excluded from 4K/4L`,
        });
      }
      orig.set(ok, { ...d, amt: d.amt }); // chain further amendments
      if (key && isZero(amt)) key = '';
      if (key && d.rfy === cy) {
        // amendment differentials go to 4K / 4L (taxable) or 5J / 5K (non-taxable); B2CS/AT stay netted in their own rows
        if (!['B2CS', 'AT', 'TXP', 'B2CL'].includes(d.family) && !(d.family === 'CDNUR' && d.invTyp === 'B2CL')) {
          const nonTax = key.startsWith('5');
          const credit = d.dtype === 'CN';
          const effect = credit ? -1 : 1; // a CN increase reduces supplies
          const total = amt.reduce((a, b) => a + b, 0) * effect;
          if (credit) amt = neg(amt);
          if (total >= 0) key = nonTax ? '5J' : '4K';
          else {
            key = nonTax ? '5K' : '4L';
            amt = neg(amt);
          }
        } else if (d.family === 'TXP') amt = neg(amt);
      } else if (key && d.rfy === ny) {
        const credit = d.dtype === 'CN';
        const total = amt.reduce((a, b) => a + b, 0) * (credit ? -1 : 1);
        if (credit) amt = neg(amt);
        if (total >= 0) key = '10';
        else {
          key = '11';
          amt = neg(amt);
        }
      }
    } else {
      // sign conventions: 4A nets CDNUR(B2CL) notes; 4F nets advance adjustments
      if (key === '4A' && d.family === 'CDNUR') amt = d.dtype === 'CN' ? neg(amt) : amt;
      if (key === '4F' && d.family === 'TXP') amt = neg(amt);
    }
    if (key === '11-REVIEW') note = 'next-FY credit note - include in Table 11 only if it relates to this FY supplies';
    rows.push({
      key,
      amt,
      info: [
        d.rfy,
        d.period,
        d.section,
        d.dtype,
        d.num,
        d.date,
        d.dfy,
        d.party,
        d.pos,
        d.rate as any,
        d.rcm,
        d.file + (note ? ' | ' + note : ''),
      ],
    });
  }
  const rev = rows.filter(r => r.key === '11-REVIEW');
  if (rev.length)
    warns.push({
      level: 'warn',
      file: '',
      msg: `${rev.length} credit note(s) in next-FY GSTR-1 (Apr–Nov) totalling taxable ${r2(rev.reduce((a, r) => a + r.amt[0], 0)).toLocaleString('en-IN')} are tagged 11-REVIEW in RAW_GSTR1. If they relate to this FY's supplies, add them to Table 11 (adjustment column).`,
    });
  return { rows, hsn: hsnRows, docs: all };
}

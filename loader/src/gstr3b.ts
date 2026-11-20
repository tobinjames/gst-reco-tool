// GSTR-3B (portal-generated PDF) → RAW_3B and RAW_3B_PAY rows.
// Input: text lines (tokens in reading order) built from pdf.js text items by position.
import { Amt, MONTHS, RawRow, n, r2 } from './util';

export interface Line { y: number; page: number; tokens: string[] }

/** group pdf.js text items into lines (ignores rotated watermark text) */
export function linesFromItems(pages: { items: { str: string; transform: number[] }[] }[]): Line[] {
  const out: Line[] = [];
  pages.forEach((pg, pi) => {
    const items = pg.items.filter(it => it.str && it.str.trim() && Math.abs(it.transform[1]) < 0.01 && Math.abs(it.transform[2]) < 0.01);
    items.sort((a, b) => b.transform[5] - a.transform[5] || a.transform[4] - b.transform[4]);
    const lines: { y: number; parts: { x: number; s: string }[] }[] = [];
    for (const it of items) {
      const y = it.transform[5], x = it.transform[4];
      let l = lines.find(L => Math.abs(L.y - y) < 2.5);
      if (!l) { l = { y, parts: [] }; lines.push(l); }
      l.parts.push({ x, s: it.str });
    }
    for (const l of lines) {
      l.parts.sort((a, b) => a.x - b.x);
      const tokens = l.parts.map(p => p.s).join(' ').split(/\s+/).filter(Boolean);
      // drop stray single-letter watermark fragments ("FIL", "ED", "E", "D"…)
      if (tokens.length === 1 && /^(F|I|L|E|D|FIL|ED|FILED)$/.test(tokens[0])) continue;
      out.push({ y: l.y, page: pi + 1, tokens });
    }
  });
  return out;
}

const isNum = (t: string) => /^-?[\d,]+(\.\d+)?$/.test(t) || t === '-';

export interface ThreeB {
  gstin: string; legalName: string; year: string; month: string; period: string; arn: string; arnDate: string;
  rows: RawRow[]; pay: { key: string; amount: number }[]; problems: string[];
}

export function parse3b(lines: Line[], file: string): ThreeB {
  const res: ThreeB = { gstin: '', legalName: '', year: '', month: '', period: '', arn: '', arnDate: '', rows: [], pay: [], problems: [] };
  let sec = '', sub = '';
  const vals: Record<string, number[]> = {};
  const payAgg: Record<string, number> = {};
  for (const L of lines) {
    const t = L.tokens, text = t.join(' ');
    // trailing numeric tokens
    let k = t.length;
    while (k > 0 && isNum(t[k - 1])) k--;
    const label = t.slice(0, k).join(' ');
    const nums = t.slice(k).map(n);

    if (/^Year\b/.test(text) && !res.year) res.year = t[t.length - 1];
    else if (/^Period\b/.test(text) && !res.month) res.month = t[t.length - 1];
    else if (/^GSTIN of the supplier/.test(text)) res.gstin = t[t.length - 1];
    else if (/^2\(a\)\./.test(text)) res.legalName = text.replace(/^2\(a\)\.\s*Legal name of the registered person\s*/, '');
    else if (/^2\(c\)\./.test(text)) res.arn = t[t.length - 1];
    else if (/^2\(d\)\./.test(text)) res.arnDate = t[t.length - 1];

    if (/^3\.1\.1 /.test(text)) { sec = '3.1.1'; continue; }
    if (/^3\.1 /.test(text)) { sec = '3.1'; continue; }
    if (/^3\.2 /.test(text)) { sec = '3.2'; continue; }
    if (/^4\. ?Eligible ITC/.test(text)) { sec = '4'; sub = ''; continue; }
    if (/^5 Values of exempt/.test(text)) { sec = '5'; continue; }
    if (/^5\.1 /.test(text)) { sec = '5.1'; continue; }
    if (/^6\.1 /.test(text)) { sec = '6.1'; sub = ''; continue; }
    if (/^Breakup of tax liability/.test(text) || /^Verification/.test(text)) { sec = 'end'; continue; }
    if (sec === '4') {
      if (/^A\. ITC Available/.test(text)) { sub = 'A'; continue; }
      if (/^B\. ITC Reversed/.test(text)) { sub = 'B'; continue; }
      if (/^C\. Net ITC/.test(text)) { sub = 'C'; continue; }
      if (/^\(D\) Other Details/.test(text)) { sub = 'D'; continue; }
    }
    if (sec === '6.1') {
      if (/^\(A\) Other than reverse charge/.test(text)) { sub = 'A'; continue; }
      if (/^\(B\) Reverse charge/.test(text)) { sub = 'B'; continue; }
    }
    if (!nums.length) continue;
    const lab = label.replace(/\s+/g, '');
    const set = (key: string, need: number) => {
      if (nums.length !== need) { res.problems.push(`${key}: expected ${need} values, found ${nums.length} ("${text}")`); return; }
      vals[key] = nums;
    };
    if (sec === '3.1') { const m = lab.match(/^\(([a-e])\)/); if (m) set('3.1' + m[1], 5); }
    else if (sec === '3.1.1') { const m = lab.match(/^\((i{1,2})\)/); if (m) set('3.1.1' + m[1], 5); }
    else if (sec === '4' && (sub === 'A' || sub === 'B' || sub === 'D')) { const m = lab.match(/^\((\d)\)/); if (m) set('4' + sub + m[1], 4); }
    else if (sec === '5') {
      if (/^Fromasupplier/i.test(lab)) set('5a', 2); else if (/^NonGST/i.test(lab)) set('5b', 2);
    } else if (sec === '6.1' && sub) {
      const head = /^Integrated/.test(label) ? 'A' : /^Central/.test(label) ? 'B' : /^State/.test(label) ? 'C' : /^Cess/.test(label) ? 'D' : '';
      if (!head) continue;
      if (nums.length !== 10) { res.problems.push(`6.1 ${sub}/${head}: expected 10 values, found ${nums.length}`); continue; }
      const [payable, , net, itcI, itcC, itcS, itcCs, cash, intr, fee] = nums;
      const addp = (key: string, v: number) => { payAgg[key] = r2((payAgg[key] || 0) + v); };
      addp(`9${head}-PAY`, net); addp(`9${head}-GROSS`, payable); addp(`9${head}-CASH`, cash);
      addp(`9${head}-ITC-IGST`, itcI); addp(`9${head}-ITC-CGST`, itcC); addp(`9${head}-ITC-SGST`, itcS); addp(`9${head}-ITC-CESS`, itcCs);
      addp('9E-PAY', intr); addp('9F-PAY', fee);
    }
  }
  const mi = MONTHS.indexOf((res.month || '').toLowerCase());
  if (mi >= 0 && res.year) {
    const startYr = +res.year.slice(0, 4);
    const yr = mi >= 3 ? startYr : startYr + 1;
    res.period = String(mi + 1).padStart(2, '0') + yr;
  } else res.problems.push('Return period not found');

  const labels: Record<string, string> = {
    '3.1a': '3.1(a) Outward taxable supplies (other than zero rated, nil, exempted)', '3.1b': '3.1(b) Outward taxable supplies (zero rated)',
    '3.1c': '3.1(c) Other outward supplies (nil rated, exempted)', '3.1d': '3.1(d) Inward supplies (liable to reverse charge)', '3.1e': '3.1(e) Non-GST outward supplies',
    '3.1.1i': '3.1.1(i) Supplies where ECO pays tax u/s 9(5)', '3.1.1ii': '3.1.1(ii) Supplies made through ECO u/s 9(5)',
    '4A1': '4A(1) Import of goods', '4A2': '4A(2) Import of services', '4A3': '4A(3) Inward supplies liable to RCM', '4A4': '4A(4) Inward supplies from ISD',
    '4A5': '4A(5) All other ITC', '4B1': '4B(1) Reversed - rules 38, 42, 43 & 17(5)', '4B2': '4B(2) Reversed - others', '4D1': '4D(1) ITC reclaimed',
    '4D2': '4D(2) Ineligible ITC u/s 16(4) & PoS', '5a': '5 Composition / exempt / nil inward', '5b': '5 Non-GST inward',
  };
  for (const key of Object.keys(labels)) {
    const v = vals[key];
    if (!v) { res.problems.push(`${labels[key]} not found`); continue; }
    let amt: Amt;
    if (v.length === 5) amt = [v[0], v[2], v[3], v[1], v[4]];          // taxable, igst, cgst, sgst, cess → T,C,S,I,Cess
    else if (v.length === 4) amt = [0, v[1], v[2], v[0], v[3]];        // igst, cgst, sgst, cess
    else amt = [r2(v[0] + v[1]), 0, 0, 0, 0];                          // inter + intra
    res.rows.push({ key, amt, info: [res.year, res.period, res.arnDate, labels[key], file] });
  }
  res.pay = Object.entries(payAgg).map(([key, amount]) => ({ key, amount }));
  return res;
}

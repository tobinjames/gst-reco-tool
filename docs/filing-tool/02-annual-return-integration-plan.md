# Plan: the current tool takes the client's filing workbook as input for the annual return

James & James LLP · planning document, 25 Nov 2026 · nothing here is built.
Companion: [01-filing-tool-prompt-and-handover.md](01-filing-tool-prompt-and-handover.md) (the workbook contract is §5 there).
Based on `docs/HANDOVER.md` as of Loader v0.19.

## 1. The idea and what it means

Tobin's idea: the sales register carries a **section code** (b2b, b2cs, …) and a **return period** per row, so the current tool no longer needs the GSTR-1 JSON files to know where each sale goes. Purchase register and 2B rows point at each other through **extra columns**, so no new matching run is needed at year end.

Today the Annual return tab takes GSTR-1 JSON/ZIP (PY, CY, NY), 2B JSON (CY plus next-year April to October), 3B PDFs, a purchase and sales register, the portal's auto-drafted GSTR-9 JSON, and (recommended) the saved Input Reco, Revenue reco and Ledger flow workpapers. The plan adds one more source, **the Filing workbook**, which can stand in for the GSTR-1, 2B, 3B and register files of every period it covers.

The change is mostly an **adapter layer**: read the workbook and produce the same internal rows the current code already builds from JSON and PDFs. GSTR-9 logic, ICAI treatments, the Recon sheets and the output workbook stay as they are. That keeps the risk small and gives a strong test: *annual output from the filing workbook must equal annual output from the raw files it was built from.*

## 2. Two routes, one adapter

| Route | Who | Input |
|---|---|---|
| A. Filing workbook | clients prepared with the new tool | `Filing_<GSTIN>.xlsx` (contract v1) |
| B. Plain registers with codes | clients not on the new tool, or a one-off | the user's own sales register with **Section Code** and **Return Period** columns, purchase register and 2B with **Match Code** columns (a small template workbook with these headings is provided) |

Both feed the same reader (`src/filingwb.ts`): route B is simply route A with only some sheets present. Tobin's literal idea is route B; route A is what the new tool produces automatically.

## 3. Mapping: filing workbook → what the annual tab needs

### 3.1 Sales and GSTR-1 (Tables 4, 5, 10, 11, 17)
`gstr1.ts` turns GSTR-1 JSON into RAW_GSTR1 rows carrying GSTR-9 keys (4A–4N, 5A–5N, 10/11, 11-REVIEW) and amendment differences. The adapter does the same from `Sales` rows.

| Section code | GSTR-9 table |
|---|---|
| `b2cl`, `b2cs`, `ecom_9_5` (supplier-side per operator rules) | 4A supplies to unregistered persons |
| `b2b` | 4B |
| `exp_wp` | 4C exports with payment of tax |
| `sez_wp` | 4D SEZ with payment |
| `deemed` | 4E |
| `at` (advance, tax paid, invoice not yet issued) | 4F |
| RCM inward is from purchases, not sales | 4G (from 3B/purchases as today) |
| `cdnr`, `cdnur` credit notes / debit notes | 4I / 4J (by Doc Type) |
| amendment rows (Original Doc filled) | 4K / 4L, and 10 / 11 when they relate to the previous year and fall in the allowed window |
| `exp_wop`, `sez_wop` | 5A, 5B |
| `exempt`, `nil`, `nongst` | 5D, 5E, 5F |
| credit/debit notes on nil, exempt | 5H / 5I |

The existing `gstr1.ts` is the authority on the exact keys and the ICAI treatments (e.g. what goes where for advances adjusted via `txpd`, amendments of the next year in April to October, the 11-REVIEW tag). **Implementation rule: write the adapter by factoring out of `gstr1.ts` the part that turns "documents" into RAW rows**, so JSON and workbook share one code path, rather than writing a second mapping. The table above is the intended result, to be checked against that code and the ICAI technical guide in `icai-gstr9-guidance`, not a new rule set.

Things the adapter needs that the register must carry: `HSN/SAC`, `UQC`, `Qty`, `Rate` (Table 17 outward HSN summary, Table 12 of GSTR-1), `POS`, `Doc Type`, `Original Doc No/Date`, `Return Period`. If HSN is missing the adapter fills what it can and shows a warning (Table 17 then comes from the JSON route if the user also drops GSTR-1 files).

### 3.2 Return period and financial year
- GSTR-9 follows the **return period of the filed return**, not the invoice date. The workbook's `Return Period` column is that period. A March invoice reported in the April return is April, belonging to the next FY, and is treated as Table 10/11 style next-year reporting by the same rules as today.
- FY, PY, CY and NY are derived from `Return Period` with `util.ts` FY maths. The NY April to October window (ICAI) is applied as today.
- Only periods with `Periods.Status = Filed` count as "GSTR-1 equivalent". Prepared-but-unfiled periods are listed as **not usable** with the reason. A period that was filed but whose JSON hash no longer matches the `Sales` rows (someone edited after filing) is flagged "changed after filing" and the user must choose: use the rows, or drop the portal JSON for that period.

### 3.3 Purchases, 2B and ITC (Tables 6, 7, 8, 12, 13)
- `GSTR2B` → RAW_2B rows (same keys 8A, 8A+8C, 6A1, 8G, 8G+8H1 and superseded logic). This is a straight read because the sheet is built from `gstr2b.ts` output in the first place. Next-year April to October 2B rows are in the same sheet by return period.
- `Purchases` → Books_Purchase rows as `register.ts` produces them today (documents types, negative notes, RCM flag).
- **Matching through the extra columns**: `Match Code` on `Purchases` and `GSTR2B` is passed to `reconPurchases` as manual match codes (the current Recon sheet already reads M-codes and X keep-apart codes). Rows with a code are paired exactly as the preparer confirmed; rows without a code are matched by the normal rules using the matching settings in the workpaper or the tab. Result: the annual Recon_Purchases is identical in structure and the preparer's year of decisions is respected. No second matching vocabulary is invented.
- `ITC` sheet decisions feed the claimed / reversed columns the annual tab already derives for Tables 6, 7 and 8: `Claim Period` decides the year in which ITC is taken (this is exactly what Tables 6A1/6H/7 rely on), and `Decision` (ineligible 17(5), rule 37, 42/43) maps to the 7A–7H reversal rows. *Check the current key names in `rowmap.json` and the 7E/7I treatment in the handover before writing the map; do not invent keys.*
- Blocked-credit decisions are only as good as the user's confirmations; the annual workbook shows them as a reviewable list with who confirmed and when (from `ITC` and `Audit`).

### 3.4 GSTR-3B
- `GSTR3B` sheet holds the **filed** figures per period (3.1, 3.1.1, 4A–4D, 5, 6.1). The adapter produces the same Line/ThreeB structures that `gstr3b.ts` builds from PDF text, so the existing 3B-based tables and the Revenue reco "GSTR-1 → 3B" bridge work without PDFs.
- Prepared-but-not-filed values are ignored; a period with no filed 3B shows as missing exactly as a missing PDF does today.

### 3.5 What stays external
- Portal auto-drafted GSTR-9 JSON (RAW_Portal9) is still downloaded from the portal; the workbook cannot replace it.
- GST ledger flow (credit / cash / liability ledgers from the portal, GL dump) is separate and unchanged.
- Books for GSTR-9C (trial balance mapping) are unchanged.

## 4. Precedence when the same period arrives from two sources

The workbook is *our record*; portal files are *portal truth*. Rule:
1. If the user also drops a GSTR-1 JSON, 2B JSON or 3B PDF for a period, **the portal file wins** for that period and the tool reports the differences against the workbook (count and value by section). That is the "portal edited online / amended" safety net.
2. If the same period exists in the workbook and in a dropped workpaper, the existing v0.14 rules apply (register replacements only).
3. Every sheet of the annual workbook shows a **Source** note per period ("filing workbook", "portal JSON", "3B PDF").
4. A summary on the Workpapers sheet lists period coverage (12 months CY, 7 months NY, PY for amendments) with source and any "changed after filing" flags.

## 5. Changes in the current tool, by file

| File | Change |
|---|---|
| `src/build.ts` `classify()` | new file kind `'Filing workbook'`, detected by a hidden metadata sheet (`_JJFiling`: schema, GSTIN, tool version) and sheet names; refuses an unknown schema version with a clear message; checks GSTIN against the client (existing warning for another GSTIN) |
| `src/filingwb.ts` (new) | reads the workbook (reader in `xlsx.ts`), validates the contract (required columns, allowed codes, one row id per row, match code symmetry), verifies the audit chain and period hashes (warn, never block, but show "chain broken / changed after filing"), returns typed tables |
| `src/gstr1.ts` | factor the "documents → RAW rows" step into a function callable by both the JSON reader and the workbook adapter |
| `src/filingadapt.ts` (new) | sales → RAW_GSTR1 rows; 2B → RAW_2B; purchases → Books_Purchase; 3B → ThreeB; Match Codes → manual codes; ITC decisions → claimed/reversal columns |
| `src/wizard.ts` | NEEDS lists: for GSTR-1, 2B, 3B, registers the file checklist shows "satisfied by filing workbook for N of 12 months"; one file instead of many |
| `src/state.ts` | the filing workbook is a workbook and is not embedded in `_JJFiles` (as for every workbook); the annual workpaper records the workbook name, schema, last Seq and head hash it was built from, so a rebuild can tell the workbook changed |
| `src/revenue/*`, Input Reco | optional second step: the Revenue reco and Input Reco tabs accept the same workbook (sales + filed 1/3B; purchases + 2B + matches), so the monthly tabs also start from one file. Not needed for the first annual release |
| `workbook/` | `build_filing_sample.py` fixture generator (also used by the new tool's tests) and the route-B template with the extra columns |
| `src/help.ts`, `flow/guide.ts`-style guide | new "Using the filing workbook" section; update "getting the files" |
| Tests | see §7 |

Version: **v0.20** (minor, additive). No change to existing workbook templates, so no migration of saved workpapers.

## 6. Scale and archives

- A multi-year workbook may contain several FYs. The adapter filters to the client's FY plus PY and NY windows before any heavy work; large sheets are read once in a streaming pass (check the current `xlsx.ts` Reader memory use with a 100,000-row sheet and set a limit with a clear message).
- If the new tool's "Close the year" archive (handover §5.6) is used, the annual tool accepts the **archive workbook of that FY** plus the current workbook for the next-year April to October window. The adapter treats archives as the same contract with a `Kind = Archive` flag.

## 7. Tests and acceptance

1. **Equivalence test (the main one).** Build a filing workbook from the demo set (`samples/input-reco`, `samples/revenue`, the demo GSTR-1/2B/3B), run annual from the workbook and from the raw files, and require: identical GSTR-9 cell values, identical Recon_Purchases match results, identical Summary counts. Differences listed explicitly and each explained or fixed.
2. Contract tests: missing column, unknown code, duplicate Row Id, one-sided match code, two FYs in one workbook, locked and unfiled periods, a period "changed after filing".
3. Precedence test: workbook plus a portal JSON with a different invoice → portal wins and the difference is reported.
4. Route B: a plain sales register with Section Code and Return Period columns gives the same Table 4/5 as the JSON.
5. Real-file case: one anonymised client year once Tobin provides it (needed; none yet).
6. LibreOffice recalculation of the annual workbook: zero mismatches.
7. Performance: a workbook of 3 years and 60,000 sales rows reads in an acceptable time (target set when measured).

## 8. Risks and honest notes

- **The workbook is not the portal.** If staff amend online without recording it, the annual numbers drift. Mitigation: portal file wins, the optional difference report, a prompt to download the final GSTR-1 JSON of each period at year end when possible.
- **Mapping errors are silent errors.** Hence the equivalence test and the rule to reuse `gstr1.ts` code, not to copy it.
- **Hand-edited workbooks.** Staff will edit cells. The adapter validates the contract and reports, but cannot prevent bad data. The audit chain warns about edits only if the edit happened outside the new tool.
- **HSN / quantity data** may be missing in small clients' manual entries; Table 17/18 then need the JSON route or manual input (as today).
- **Immutability of the log is limited** (handover §3); the annual tool shows verification results but does not claim more.
- **Schema changes.** A new GSTN format or a changed GSTR-9 means the contract and adapter must move together; version both.
- **Unverified here:** exact key names in `rowmap.json` and the full set of 7A–7H keys, how Tables 17/18 are filled today, and whether GSTR-9 auto-population rules for FY 2025-26 changed. The builder must check these in the code and the GSTN offline tool, not rely on this plan.

## 9. Steps and effort (rough)

| Step | Output | Approx. effort |
|---|---|---|
| 1 | Contract frozen with the new tool; fixture generator; route-B template | 1–2 days (shared with the filing tool's phase 0) |
| 2 | `filingwb.ts` reader and validations | 2 days |
| 3 | Factor `gstr1.ts`; sales adapter; equivalence test for Tables 4/5/10/11 | 3 days |
| 4 | 2B / purchases / match codes / ITC adapter; equivalence for Tables 6–8 | 3 days |
| 5 | 3B adapter; Workpapers/coverage sheet; precedence rules; wizard checklist and help | 2 days |
| 6 | Real-file case; review; release v0.20 | 1–2 days after samples arrive |

Order of work: the adapter can only be finished after the filing tool's phase 1 and 2 have fixed the sheets; steps 1 to 3 can start as soon as the contract is frozen.

## 10. Questions for Tobin

1. Is it acceptable that **only filed periods** count and that the portal file wins when both exist?
2. Should Revenue reco and Input Reco also take the filing workbook in the same release (§5, last row), or annual first?
3. For clients not on the new tool, do you want route B (template with codes) built at all, or is it enough to keep the JSON/PDF route?
4. May the annual tool read an **archive** workbook per FY (so history can leave the live file)?
5. Sample data: a real client's sales register, filed GSTR-1 JSON(s) with amendments, 2B and 3B for a year, anonymised, to use as the equivalence case.

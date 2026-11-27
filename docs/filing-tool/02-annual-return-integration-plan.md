# Plan: the current tool takes the client's filing workbook as input for the annual return

James & James LLP · planning document, 25 Nov 2026 · nothing here is built.
Companion: [01-filing-tool-prompt-and-handover.md](01-filing-tool-prompt-and-handover.md) (the workbook contract is §5 there).
Based on `docs/HANDOVER.md` as of Loader v0.19.

## Decisions by Tobin, 25 Nov 2026 (this version follows them)

- **Portal data is always required.** The annual return still loads the portal files (GSTR-1 JSON, GSTR-2B JSON, GSTR-3B PDFs, the auto-drafted GSTR-9 JSON). They are the source of truth for what was reported. The filing workbook does **not** replace them.
- The workbook adds the **books detail**: invoice-level sales with section code and return period, purchases, the purchase-to-2B links (match codes), ITC / blocked-credit decisions and the filed 3B figures it recorded. Unfiled (draft) periods do not count.
- One workbook per GSTIN. The filing features are added **inside the current tool** (same repo, same HTML, `loader/src/filing/`), so the annual adapter is an internal module, not a second app. A filed period is locked for good (corrections are amendments in a later period). Closed-year **archive** workbooks are accepted. Of the other tabs, only **GST ledger flow** will later take the workbook (Revenue reco and Input Reco already work from reconciled data and do not need it).
- Test data: the existing demo samples in the project files (`samples/input-reco`, `samples/revenue`); real anonymised files later.

## 1. The idea and what it means

Tobin's idea: the sales register carries a **section code** (b2b, b2cs, …) and a **return period** per row, so the current tool no longer needs the GSTR-1 JSON files to know where each sale goes. Purchase register and 2B rows point at each other through **extra columns**, so no new matching run is needed at year end.

Today the Annual return tab takes GSTR-1 JSON/ZIP (PY, CY, NY), 2B JSON (CY plus next-year April to October), 3B PDFs, a purchase and sales register, the portal's auto-drafted GSTR-9 JSON, and (recommended) the saved Input Reco, Revenue reco and Ledger flow workpapers. The plan adds one more source, **the Filing workbook**, which supplies the books side (registers, match codes, ITC decisions, section codes and return periods per sales row) for every period it covers. Portal files are still loaded for the same periods and win wherever the two disagree (§4).

The change is mostly an **adapter layer**: read the workbook and produce the same books-side rows the current code already builds from a sales / purchase register, plus the match codes and decisions. GSTR-9 logic, ICAI treatments, the Recon sheets and the output workbook stay as they are. That keeps the risk small and gives a strong test: *the annual output must not change when the same books arrive through the filing workbook instead of the registers, and the section totals of the workbook must agree with the portal GSTR-1.*

## 2. Two routes, one adapter

| Route | Who | Input |
|---|---|---|
| A. Filing workbook | clients prepared with the new tool | `Filing_<GSTIN>.xlsx` (contract v1) **plus the portal files** |
| B. Plain registers with codes | clients not on the new tool, or a one-off | **plus the portal files**; the user's own sales register with **Section Code** and **Return Period** columns, purchase register and 2B with **Match Code** columns (a small template workbook with these headings is provided) |

Both feed the same reader (`src/filingwb.ts`): route B is simply route A with only some sheets present. Tobin's literal idea is route B; route A is what the new tool produces automatically.

## 3. Mapping: filing workbook → what the annual tab needs

### 3.1 Sales and GSTR-1 (Tables 4, 5, 10, 11, 17)
`gstr1.ts` turns the portal GSTR-1 JSON into RAW_GSTR1 rows carrying GSTR-9 keys (4A–4N, 5A–5N, 10/11, 11-REVIEW) and amendment differences. **These rows stay the figures in the annual return.** The adapter turns the workbook's `Sales` rows into the same GSTR-9 keys by section code, and uses them to:
1. **check** each period and table: workbook total vs portal JSON total (count, taxable, tax heads), listing differences;
2. **explain and reconcile**: invoice-level detail with customer, section code and period for the Revenue reco and the Sales side of Recon (so the books-side sales register no longer needs separate mapping);
3. **fill what the JSON lacks** where it helps (e.g. customer names, HSN/UQC/quantity for Table 17 when the portal JSON's HSN table is missing or summarised differently), shown as workbook-sourced.

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
- Only periods with `Periods.Status = Filed` count. Draft (prepared or reviewed but not filed) periods are **left out** and listed with the reason. A filed period whose recorded JSON hash no longer matches the `Sales` rows (someone edited after filing) is flagged "changed after filing"; the portal JSON is used for the return figures either way and the difference is reported.

### 3.3 Purchases, 2B and ITC (Tables 6, 7, 8, 12, 13)
- The **portal 2B JSON files are still loaded** (CY plus next-year April to October) and give RAW_2B as today. The workbook's `GSTR2B` sheet carries the same documents with the row ids the `Match Code` links point to; the tool checks that every workbook 2B document exists in the loaded 2B JSON (same supplier, number, values) and re-attaches the match codes to the JSON rows by that key. A workbook 2B row not found in the portal file is reported and not used.
- `Purchases` → Books_Purchase rows as `register.ts` produces them today (documents types, negative notes, RCM flag).
- **Matching through the extra columns**: `Match Code` on `Purchases` and `GSTR2B` is passed to `reconPurchases` as manual match codes (the current Recon sheet already reads M-codes and X keep-apart codes). Rows with a code are paired exactly as the preparer confirmed; rows without a code are matched by the normal rules using the matching settings in the workpaper or the tab. Result: the annual Recon_Purchases is identical in structure and the preparer's year of decisions is respected. No second matching vocabulary is invented.
- `ITC` sheet decisions feed the claimed / reversed columns the annual tab already derives for Tables 6, 7 and 8: `Claim Period` decides the year in which ITC is taken (this is exactly what Tables 6A1/6H/7 rely on), and `Decision` (ineligible 17(5), rule 37, 42/43) maps to the 7A–7H reversal rows. *Check the current key names in `rowmap.json` and the 7E/7I treatment in the handover before writing the map; do not invent keys.*
- Blocked-credit decisions are only as good as the user's confirmations; the annual workbook shows them as a reviewable list with who confirmed and when (from `ITC` and `Audit`).

### 3.4 GSTR-3B
- The **portal 3B PDFs are still loaded** and drive the 3B-based tables and the Revenue reco "GSTR-1 → 3B" bridge as today.
- The `GSTR3B` sheet's recorded **filed** figures (3.1, 4A–4D, 6.1) are compared with the PDF per period; differences are listed (typing errors in the record, or a return filed differently from what was prepared). Prepared-but-not-filed values are ignored.

### 3.5 What stays external
- Portal auto-drafted GSTR-9 JSON (RAW_Portal9) is still downloaded from the portal; the workbook cannot replace it.
- GST ledger flow (credit / cash / liability ledgers from the portal, GL dump) is separate and unchanged.
- Books for GSTR-9C (trial balance mapping) are unchanged.

## 4. Portal files and the workbook together

The workbook is *our record of the books and our decisions*; portal files are *what was reported*. Rule:
1. **Portal files are required** for every period in the annual return. Without the GSTR-1, 2B and 3B of a period the tool says what is missing, as today. It does not fall back to the workbook's sales totals.
2. Figures for the GSTR-9 tables come from the portal files. The workbook contributes the books side (registers), the match codes, ITC decisions and per-row section codes.
3. Differences are shown, never silently resolved: a *Workbook vs portal* sheet lists, per period and section, workbook total, portal total, difference, and the invoices that explain it (in workbook not on portal, on portal not in workbook, value differences). That is the safety net for online amendments and edits after filing.
4. If the same period exists in a dropped workpaper, the existing v0.14 rules apply (register replacements only).
5. The Workpapers sheet lists period coverage (12 months CY, 7 months NY, PY for amendments) with, per period, portal files present, workbook present, status (filed / draft / missing), and any "changed after filing" flag.

## 5. Changes in the current tool, by file

The annual integration is its own module, **`loader/src/filingbridge/`** (handover §10.2): it reads the contract tables through `ClientStore` (the workbook today, a database snapshot or export later) and returns the rows the annual tab already uses; it does not import the filing modules' code and writes nothing to the workbook. Files named `src/filingwb.ts` and `src/filingadapt.ts` below live in that folder, and its own tests and README sit with it. The only edits outside it are the small hooks listed (`build.ts`, `gstr1.ts`, `wizard.ts`, `state.ts`).

| File | Change |
|---|---|
| `src/build.ts` `classify()` | new file kind `'Filing workbook'`, detected by a hidden metadata sheet (`_JJFiling`: schema, GSTIN, tool version) and sheet names; refuses an unknown schema version with a clear message; checks GSTIN against the client (existing warning for another GSTIN) |
| `src/filingwb.ts` (new) | reads the workbook (reader in `xlsx.ts`), validates the contract (required columns, allowed codes, one row id per row, match code symmetry), compares the recorded fingerprints of filed periods with the rows (warn, never block, show "changed after filing"), returns typed tables |
| `src/gstr1.ts` | factor the "documents → RAW rows" step into a function callable by both the JSON reader and the workbook adapter |
| `src/filingadapt.ts` (new) | sales → RAW_GSTR1 rows; 2B → RAW_2B; purchases → Books_Purchase; 3B → ThreeB; Match Codes → manual codes; ITC decisions → claimed/reversal columns |
| `src/wizard.ts` | NEEDS lists unchanged for the portal files; the filing workbook is an optional books source that satisfies the registers and match codes ("workbook covers N of 12 months") |
| `src/state.ts` | the filing workbook is a workbook and is not embedded in `_JJFiles` (as for every workbook); the annual workpaper records the workbook name, schema and last sign-off Seq it was built from, so a rebuild can tell the workbook changed |
| `src/flow/*` | later, optional: GST ledger flow reads the workbook for books-side GST entries and ITC / output detail (scope to be agreed when it is taken up). Revenue reco and Input Reco are unchanged (Tobin: they are already reconciled by then) |
| `workbook/` | `build_filing_sample.py` fixture generator (also used by the new tool's tests) and the route-B template with the extra columns |
| `src/help.ts`, `flow/guide.ts`-style guide | new "Using the filing workbook" section; update "getting the files" |
| Tests | see §7 |

Version: **v0.20** (minor, additive). No change to existing workbook templates, so no migration of saved workpapers.

## 6. Scale and archives

Storage will later move to a database (handover §11). The bridge therefore never opens sheets itself: it asks `ClientStore` for tables, so the annual return keeps working when a client's history lives in a database, an archive workbook, or an exported snapshot.

- A multi-year workbook may contain several FYs. The adapter filters to the client's FY plus PY and NY windows before any heavy work; large sheets are read once in a streaming pass (check the current `xlsx.ts` Reader memory use with a 100,000-row sheet and set a limit with a clear message).
- If the new tool's "Close the year" archive (handover §5.6) is used, the annual tool accepts the **archive workbook of that FY** plus the current workbook for the next-year April to October window. The adapter treats archives as the same contract with a `Kind = Archive` flag.

## 7. Tests and acceptance

1. **Equivalence test (the main one).** Build a filing workbook from the demo set (`samples/input-reco`, `samples/revenue`: demo registers plus the demo GSTR-1/2B/3B). Run annual with the portal files plus the **registers** (today's route) and with the portal files plus the **workbook**, and require: identical GSTR-9 cell values, identical Recon_Purchases and sales match results, identical Summary counts; the *Workbook vs portal* sheet shows no differences on the clean demo and the planted ones on a doctored copy. Differences listed explicitly and each explained or fixed.
2. Contract tests: missing column, unknown code, duplicate Row Id, one-sided match code, two FYs in one workbook, locked and unfiled periods, a period "changed after filing".
3. Portal-wins test: workbook with an invoice that is not in the portal GSTR-1 (and the reverse) → GSTR-9 figures follow the portal and the difference is listed. Missing portal file for a period → clear message, no silent fallback. Draft period → excluded and listed.
4. Route B: a plain sales register with Section Code and Return Period columns, plus the portal JSON, gives the same results as route A and the same section totals as the JSON.
5. Real-file case: the demo samples first (agreed); one anonymised real client year later.
6. LibreOffice recalculation of the annual workbook: zero mismatches.
7. Performance: a workbook of 3 years and 60,000 sales rows reads in an acceptable time (target set when measured).

## 8. Risks and honest notes

- **The workbook is not the portal.** If staff amend online without recording it, the annual numbers drift. Mitigation: portal file wins, the optional difference report, a prompt to download the final GSTR-1 JSON of each period at year end when possible.
- **Mapping errors are silent errors.** Hence the equivalence test and the rule to reuse `gstr1.ts` code, not to copy it.
- **Hand-edited workbooks.** Staff will edit cells. The adapter validates the contract and reports, but cannot prevent bad data. The sign-off sheet password stops accidental edits only; the fingerprint comparison reports changes to filed periods.
- **HSN / quantity data** may be missing in small clients' manual entries; Table 17/18 then need the JSON route or manual input (as today).
- **The sign-off log is only password-protected** (handover §8, Tobin's decision); the annual tool lists who prepared and reviewed each period but does not claim more.
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

## 10. Open questions for Tobin

Answered on 25 Nov 2026: one workbook per GSTIN; same repo; portal files always required and authoritative, drafts excluded; demo samples for testing; only GST ledger flow takes the workbook later (Revenue reco and Input Reco do not); archive workbooks accepted; filed periods are locked for good, no unlock.

1. Route B (a plain sales register with Section Code / Return Period columns, plus portal files) for clients not on the new tool: build it or not? Recommended: later, it is cheap once the adapter exists.
2. GST ledger flow with the workbook: what exactly should it read? To be defined when that work is taken up.
3. Real anonymised client data for a final check, when available.

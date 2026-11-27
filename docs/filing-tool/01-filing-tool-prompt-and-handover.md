# GST filing features (GSTR-1 and GSTR-3B preparation): build prompt and handover

James & James LLP · planning document, 25 Nov 2026 · for the build planned the week of 30 Nov 2026
Status: **plan only, nothing is built.** Decided 25 Nov 2026: the features are added **inside the current tool** (new tab(s) in the loader, same repo, same offline HTML), not as a separate app. See §10. Companion: [02-annual-return-integration-plan.md](02-annual-return-integration-plan.md).

This file has two parts. **Part A** is a self-contained brief for the person (or Claude session) who builds the tool. **Part B** is the prompt to paste into a new Claude session. Part A can be read without Part B and vice versa.

---

# Part A. Handover

## 1. Goal in one paragraph

New features of the current tool for the firm's staff to **prepare GSTR-1 and GSTR-3B for clients every month (or quarter)**. Each client has **one Excel workbook** that holds all of that client's data across months and years and **grows every time a return is filed**. Staff load a sales register (or key in sales by hand for small clients), the tool generates the **GSTR-1 JSON** for upload to the GST portal, optionally checks the portal's GSTR-1 preview against the data, matches the purchase register to GSTR-2B (same logic as the current tool), prepares GSTR-3B figures, and records **who prepared and who reviewed** each return in a password-protected sign-off log inside the workbook. At year end the same workbook, together with the portal files, is the **input to the current annual return tool** (see companion plan).

## 2. What the person asked for (checklist the build must cover)

| # | Requirement (Tobin, 25 Nov 2026) | Where it is covered |
|---|---|---|
| 1 | One Excel file per client holding multiple months and years; every filing adds to it | §5 workbook contract, §6 lifecycle |
| 2 | Today's clients are small; later the history may have to move out of Excel | §5.6 scale limits, §11 moving history out |
| 3 | Upload a sales register, **or** enter records one by one (small clients without a proper register) | §7.1 |
| 4 | Generate the GSTR-1 JSON for filing | §7.2 |
| 5 | Optional: upload the GSTR-1 **preview** before filing and validate it against the data | §7.3 |
| 6 | Record of who prepared and who reviewed, inside the file, that nobody changes by accident | §8 (sheet password; honest limits) |
| 7 | Hosting on Cloudflare free tier, about 50 users | §9 |
| 8 | Upload purchase register and GSTR-2B and match them, using the current tool's logic | §7.4 |
| 9 | Purchase register missing: client confirms vendor-wise; summary PDF plus detailed Excel for the client | §7.5 |
| 10 | Detect non-eligible ITC from purchase register / 2B by keywords (insurance, vehicle, auto, hotel, restaurant, others); only a suggestion, user confirms | §7.6 |
| 11 | At year end the single file feeds the current annual return tool | §5, companion plan |
| 12 | Also GSTR-3B preparation (title of the project) | §7.7 |

## 3. Honest limits (read before promising anything to staff)

1. **The sign-off log is protected by an Excel sheet password, not by cryptography.** Tobin's decision (25 Nov 2026): a simple password is enough because nobody will change it knowingly. Be clear to staff what that is: sheet protection stops accidental edits and casual changes; anyone who wants to can remove it in a minute, and the names in the log are chosen or typed by the user (there is no login), so they prove nothing against a determined person. A hash chain and a server record were considered and dropped as not needed. If a stronger claim is ever needed, they can be added later without changing the workbook layout (the log already has one row per event).
2. **The tool cannot file anything.** It produces a JSON that a person uploads on the GST portal and files with DSC/EVC. "Filed" status and ARN are recorded by the user (or read from a downloaded acknowledgement), not observed.
3. **The workbook is a record of what we prepared**, not necessarily of what the portal holds. Portal edits, amendments made online, IMS actions and late 2B changes can differ. The optional preview check (§7.3) and the annual tool's comparison with downloaded returns exist for this reason.
4. **GSTN formats and rules change** (the GSTR-1 JSON moved to v5.0 in Feb 2026; B2CL limits, HSN digit rules and 3B auto-population have all changed in the last two years). Every threshold and field list must come from a **config file or the downloaded GSTN sample**, never hard-coded, and be dated. Anything in this document marked *(verify)* must be checked against the GST portal / developer portal when building. Fetch recipes are in project memory (`gstn-api-schemas`).
5. **Excel is not a database.** Limits in §5.6. Design so that Excel is one storage adapter among several.
6. **Cloudflare free-tier numbers change.** Quotas quoted here are from memory *(verify)*.

## 4. Principles carried over from the current tool

- Browser-only processing; client files stay on the user's PC. There is no server component. This keeps the "files never leave the PC" promise and limits confidentiality risk.
- No macros, no Excel 365-only functions. SUMIFS / INDEX / MATCH, **one readable formula per cell**; someone without the tool must be able to follow the working (Tobin, workpaper-format).
- **Aptos** font throughout; **Index** sheet with links; **one Instructions sheet**; Summary sheets; "Back to Index" link on every sheet; Prepared/Reviewed sign-off block. Colour #2C388E, James & James logo (loader/src/brand/).
- Saved workbook restores everything (hidden state sheets, as `_JJState` / `_JJFiles` in the current tool). Portal returns are permanent; registers and ledgers may be replaced.
- Reasons / decisions are derived or chosen from lists, with logged overrides; no free-text accounting reasons.
- Keep `npm run check` green, add real-file test cases (anonymised), and update the in-app Help whenever logic changes.
- Do not produce source zips; the repo is the source of truth. Deliver the built HTML and the PR link.

## 5. Workbook contract (version 1)

This is the interface between the new tool and the current annual tool. Treat it as an API: **change it only by bumping `schema` and writing a migration.** Sheet and column names below are the proposal; the build may refine them, but the annual tool reads whatever is finally agreed, so freeze it early (week 1) and put the final version in `docs/filing-tool/workbook-contract.md`.

### 5.1 Files and naming
`Filing_<GSTIN>.xlsx`, **one workbook per GSTIN** (decided by Tobin, 25 Nov 2026; the current tool is single-GSTIN per run; a client with several GSTINs gets several workbooks and a shared client name).

### 5.2 Sheets

| Sheet | Purpose |
|---|---|
| `Index` | cover, contents with links, colour key, sign-off |
| `Instructions` | the single instructions sheet |
| `Client` | GSTIN, legal / trade name, state code, filing frequency (monthly / QRMP), FY start, composition no, e-invoice applicable (Y/N), HSN digits required, tool version, schema version |
| `Periods` | one row per return period (MMYYYY): status of GSTR-1, 3B (Not started / Prepared / Reviewed / Filed), JSON file name, JSON SHA-256, ARN, filing date, prepared by / at, reviewed by / at, tax payable, tax paid (cash / credit). **Formulas** summarise Sales / Purchases by period |
| `Sales` | **all sales rows, all periods** (§5.3) |
| `Purchases` | purchase register rows, all periods (§5.4) |
| `GSTR2B` | flattened 2B documents, all periods (§5.4) |
| `Match` | purchase ↔ 2B links, match codes, who confirmed (§5.4) |
| `ITC` | eligibility decision per purchase / 2B document (§5.5) |
| `Vendor_Confirm` | vendor-wise confirmation rounds (§7.5) |
| `GSTR3B` | one block per period: prepared table values, and the values actually filed (§7.7) |
| `SignOff` | sign-off and event log, protected with a sheet password (§8) |
| `Lists` | section codes, rates, states, units, blocked-credit keywords (user-editable), reasons |
| `_JJState`, `_JJFiles` | hidden: settings and optional source files, as in the current tool |

### 5.3 `Sales` columns (one row per invoice **line group**; per invoice and rate)
`Row Id` (stable, never reused, e.g. S000123) · **`Return Period`** (MMYYYY of the return in which it is/was reported; the tool proposes it from the invoice date and the user can move a row) · **`Section Code`** (see below) · `Doc Type` (Invoice / Credit note / Debit note / Advance / Advance adjustment) · `Doc No` · `Doc Date` · `Party Name` · `Party GSTIN` · `POS` (state code) · `Supply Type` (Intra / Inter, derived) · `Reverse Charge` (Y/N) · `HSN/SAC` · `UQC` · `Qty` · `Rate %` · `Taxable Value` · `IGST` · `CGST` · `SGST/UTGST` · `Cess` · `Invoice Value` · `Export type / Port / Shipping bill no and date` (exports only) · `Original Doc No / Date` (amendments and CN/DN references) · `E-invoice IRN` (optional) · `Status` (Draft / In JSON / Filed / Amended) · `Locked` (Y after the period is filed) · `Source` (register file name, "manual", batch id) · `Entered by` · `Entered at` · `Remarks`.

**Section codes** (Tobin's idea; the tool derives a proposal, the user confirms; override is logged): `b2b`, `b2cl`, `b2cs`, `exp_wp` (export with payment), `exp_wop` (export without payment), `sez_wp`, `sez_wop`, `deemed` (deemed export), `cdnr` (credit/debit note to registered), `cdnur` (to unregistered), `at` (advance received), `txpd` (advance adjusted), `nil`, `exempt`, `nongst`, `ecom_9_5` (supplies through e-commerce operators taxed under 9(5)), `b2ba`/`cdnra` style amendments are not separate codes: an amendment is a row with `Original Doc No/Date` filled and `Doc Type` unchanged, the tool places it in the amendment tables when generating JSON. Document-series totals (Table 13) and HSN summary (Table 12) are **derived**, not entered. *(verify the code list against GSTR-1 v5.0 sections.)*

Derivation rules the tool should propose (all editable in config): GSTIN present → b2b; no GSTIN, inter-state and invoice value above the B2CL threshold *(verify the current limit)* → b2cl; other unregistered → b2cs; export/SEZ from the supply type; credit/debit note from Doc Type; nil/exempt/non-GST from rate and a user flag.

### 5.4 `Purchases`, `GSTR2B`, `Match`
- `Purchases`: `Row Id` (P000123), `Doc No`, `Doc Date`, `Supplier Name`, `Supplier GSTIN`, `Doc Type` (invoice / debit note / credit note / bill of entry), taxable, IGST, CGST, SGST, Cess, invoice value, `RCM` (Y/N), `Ledger / Expense head` (from register, used for keywords), `HSN/SAC` (if present), `Books Period` (month booked), `Source`, `Entered by/at`.
- `GSTR2B`: the same flattened row as the current tool's `RAW_2B` (document, supplier GSTIN and name, doc type, date, values, 2B return period, ITC-available flag and reason from 2B, supersession status for amendments). Reuse `gstr2b.ts` exactly; amendments follow the v0.13 rules (latest version counts).
- `Match` (and mirror columns): to let both tables point at each other without a lookup, **`Purchases` and `GSTR2B` each carry `Match Code` (M0001…) and `Matched Row Id`**, and `Match` holds one row per code (`Match Code`, `Purchase Row Ids`, `2B Row Ids`, `Basis` (auto rule name / manual), `Confirmed by`, `Confirmed at`, `Remark`). Many-to-one and one-to-many are therefore possible. The current tool already reads manual match codes (M-codes) and "keep apart" codes (X) from its Recon sheet; use the same vocabulary so output is compatible. Integrity rule: a code in one table must exist in the other (or be X); the tool refuses to save otherwise.

### 5.5 `ITC` decision columns
`Row Id` (P or 2B id) · `Suggested` (Eligible / Blocked 17(5) / Review) · `Suggested because` (keyword and field it hit) · `Decision` (Eligible / Ineligible 17(5) / Reverse rule 37 / Reverse rule 42-43 / Other, from `Lists`) · `Decided by` · `Decided at` · `Claim Period` (the 3B period in which ITC is claimed; this is what the annual return needs for Tables 6, 7, 8) · `Remark`.

### 5.6 Scale
- Excel allows 1,048,576 rows a sheet, but a workbook full of SUMIFS slows well before that. Rule of thumb for the design: **comfortable up to about 50,000 rows per table, painful past about 150,000**, and file size above about 30 MB is a nuisance on shared drives *(measure with the real generator and set the thresholds from tests)*.
- Mitigations in v1: `Periods` formulas reference bounded named ranges; the tool shows rows per table and file size on open and **warns at 70% of the agreed limits**; one workbook per GSTIN; offer **"Close the year"**: write a read-only archive workbook for a finished FY and carry forward only open items (unmatched 2B, unfiled drafts, opening balances) into the current workbook. The annual tool then needs the archive for that FY (companion plan §6).
- Design the code around a `ClientStore` interface (read/write rows by table and period, append audit entry) with the Excel workbook as adapter #1. See §11.

## 6. Monthly lifecycle (what staff do)

1. Open the client workbook (or create it: GSTIN, name, frequency). Opening checks the schema version and shows warnings (missing period, rows changed in a filed period, near row limits).
2. Pick the return period. Load/enter **sales** (§7.1); the tool proposes section codes; the user fixes exceptions; validation list (§7.2) is shown.
3. Generate **GSTR-1 JSON**; optional preview check (§7.3). Log: *Prepared*.
4. A second person opens the workbook, reviews the validation and summaries, and records *Reviewed* (cannot be the preparer; configurable for one-person clients, then logged as "self-reviewed").
5. Staff upload the JSON on the portal and file. They record *Filed*, ARN and date in `Periods` (or drop the acknowledgement). Rows of that period become **Locked**. Corrections later go in as amendment rows in a later period, never as edits to locked rows (the tool refuses edits on locked rows; the log shows any bypass attempt it can detect).
6. Load **purchase register** and **GSTR-2B** JSON for the period; auto-match; confirm manual matches; keyword suggestions for blocked credit; vendor confirmation if there is no register (§7.5).
7. Prepare **GSTR-3B** (§7.7); review; record filed values.
8. Save. The workbook has grown by one period.

## 7. Functional specification

### 7.1 Sales input
- **Register upload** `.xlsx/.csv`: reuse `register.ts` (header detection, synonyms, column mapping, Tally multi-column mapping, Zoho formats, saved mappings per client). New mapping fields: Section Code (optional), Return Period (optional), HSN, UQC, Qty, POS, Reverse charge, Export fields, Original Doc No/Date. If the register already has Section Code and Return Period columns (Tobin's idea, also the annual tool's standalone route) they are used and only validated.
- **Manual entry**: a form (one invoice at a time, multiple rate lines) and a quick grid; same validations as upload; GSTIN checksum, state from GSTIN, POS defaults to supplier state, IGST vs CGST+SGST derived from POS, rate from a list, totals computed, duplicate invoice number check against all periods. Manual rows have `Source = manual`.
- Re-uploading a register for a period that already has rows: show a diff (new / changed / missing) and ask; never silently duplicate. Locked periods refuse.
- A row has a stable `Row Id`; edits keep it, deletions mark `Status = Void` (never delete rows from the workbook), all logged.

### 7.2 GSTR-1 generation
- Build the JSON from `Sales` rows of one return period (including amendments, CN/DN, advances, nil/exempt, exports, HSN summary, document series). Output file name `GSTR1_<GSTIN>_<MMYYYY>.json`. Format: **GSTR-1 v5.0** or whatever the then-current offline/API format is; take field names from the GSTN developer-portal sample (project memory `gstn-api-schemas`) and keep the schema as a versioned config. The existing `gstr1.ts` *reads* v5.0; the generator must round-trip: **generate → read with gstr1.ts → compare to the Sales rows = identical** (acceptance test).
- Pre-generation validation (blocking vs warning), each with a plain-language fix: GSTIN format and checksum; invoice number length (16) and allowed characters, uniqueness per FY; date within the period and not in the future; POS valid; tax heads consistent with intra/inter; rate in the allowed list; value consistency (taxable × rate vs tax, tolerance ₹1); B2CL threshold; HSN length rule for the client's turnover *(verify)*; negative values; CN/DN reference to an existing invoice; nil-rated section totals; document-series continuity gaps; QRMP quarter rules (IFF vs quarterly return) *(verify)*.
- JSON size: the portal limits an upload (about 5 MB / 19,000 invoices a file, older rule) *(verify)*; split automatically.
- Summary shown before download: section-wise count and values; this is the thing the user compares with the portal.

### 7.3 Optional: GSTR-1 preview validation
- The user, after uploading the JSON to the portal, downloads the portal's **GSTR-1 summary / preview** (PDF and/or JSON, whichever the portal offers *(verify; needs a real sample from Tobin before designing the parser)*) and drops it into the tool.
- The tool compares **section-wise totals** (count, taxable, each tax head) and, where the preview has invoice detail, **invoice-wise**: missing on portal, extra on portal, value differences, "processed with errors". Parsing PDFs uses the current tool's pdf.js approach (`gstr3b.ts`).
- Result is informational and logged (`Preview checked: N differences`). It never blocks filing. Skipping it is allowed.

### 7.4 Purchase register vs GSTR-2B
- Reuse `recon.ts` (reconPurchases, rules, presets, tolerances, supplier-name matching `suppliers.ts`, manual M / X codes, review UI `review.ts`, matching settings `settings.ts`). **Do not rewrite matching**: move or import it as a shared module (see §10.3) and add only the persistence into `Purchases` / `GSTR2B` / `Match`.
- Monthly flow: 2B JSON for the period is loaded into `GSTR2B` (dedupe rule of v0.13: a period loaded twice is refused); register rows loaded into `Purchases`; auto-match runs across **all unmatched rows of all open periods** (late-reported invoices match in later months); confirmed matches are stored with codes.
- Outputs per period: matched / in 2B only / in books only / mismatched, supplier-wise totals, claimable ITC by head (the current tool's Summary logic: MIN of books and 2B on matched rows).

### 7.5 Client vendor confirmation (when no purchase register)
- Trigger: a period (or range) with 2B loaded and no register. The tool builds a **vendor-wise statement from `GSTR2B`**: supplier, GSTIN, number of invoices, taxable, tax by head, with a blocked-credit hint column.
- **PDF summary** for the client (Aptos is not embeddable in PDF without the font file; use an embedded open font or the firm's licensed font file *(decision needed)*): client and period header, firm letterhead/logo, plain explanation ("These are purchases the GST portal shows against you; please confirm which are genuinely yours"), vendor-wise table (top N and "others"), totals, confirm/dispute tick boxes, signature and date, return-by date. Generate in the browser with a small PDF library (e.g. pdf-lib or jsPDF, pinned version).
- **Excel detail** for the client: one row per 2B invoice with columns *Client response* (Confirmed / Not ours / Different amount / Not received yet), *Amount per client*, *Remarks*, with drop-down lists, locked columns for the data, a Version/Round id and a cell protecting the round id so a wrong file is not re-imported. Readable by a non-accountant: supplier name first, plain headings, no codes.
- **Re-import**: the returned Excel is read; responses populate `Vendor_Confirm` and drive the match (Confirmed → treated as matched to "client confirmed", eligible unless keyword suggestion says blocked; Not ours → ineligible, flagged for supplier follow-up and IMS reject advice; Different amount → hold). Everything logged with the import file hash. Rounds can be repeated.
- The firm never decides alone: the confirmation is evidence kept in the workbook (`Vendor_Confirm` row ids, who imported).

### 7.6 Blocked-credit suggestions (Sec 17(5), keyword based)
- Suggest, never decide. Each suggestion carries the **keyword and the field it hit** so the reviewer can judge. States: Suggested → **Confirmed ineligible** or **Confirmed eligible** (reason mandatory) → logged.
- Fields searched, strongest first: HSN/SAC if present (e.g. motor cars 8703, accommodation and restaurant services under SAC 9963, insurance 9971 *(verify codes)*), ledger / expense head, supplier name, narration. **GSTR-2B has no HSN and no ledger**; for 2B-only items only supplier name is searched, so suggestions there will be weak and the tool must say so.
- Starting keyword groups (a `Lists` table, editable per firm; per client overrides): motor vehicle, car, auto, automobile, two wheeler, car hire / rent-a-cab; hotel, lodging, resort, restaurant, food, catering, canteen, beverages; insurance (life, health, mediclaim); club, gym, fitness, spa, beauty, salon, cosmetic, health services; works contract, construction, civil, cement, immovable property (needs "plant and machinery" judgement); gift, free sample, personal, lost / stolen / destroyed / written off (credit reversal items), travel benefit, vacation.
- **Every group must show its exceptions** so the reviewer thinks, because keyword hits are often *eligible*: motor vehicles for dealers, transporters and driving schools; insurance where obligatory by law or used to insure outward-supply goods; food and rent-a-cab where the client supplies the same service or the law obliges the employer; hotels where the client is in the hotel business; construction capitalised into plant and machinery.
- Auto-hint sources already in 2B (ITC not available flags, reason codes) are shown as a separate "2B says ineligible" column.
- Output: a **Blocked credit schedule** (sheet + summary in the 3B working) listing document, amount, tax, category (17(5)(a) … (i)), decision. This feeds 3B 4(D)(1) *(verify table)* and the annual return's Table 6/7 reversal rows.

### 7.7 GSTR-3B preparation
- Compute from workbook data: 3.1(a) outward taxable, (b) zero-rated, (c) nil/exempt, (d) inward RCM (purchases flagged RCM), (e) non-GST; 3.2 inter-state to unregistered / composition / UIN by POS; 4(A) ITC available (from 2B matched, claim period = this period; import of goods and services; RCM; ISD), 4(B) reversals (rules 38, 42, 43, 17(5) blocked, others), 4(D) ineligible; 5 inward nil/exempt/non-GST; 6.1 tax payable vs paid by ITC / cash with set-off order. Compare with the **auto-drafted 3B** values if the user supplies the portal's 3B JSON or PDF *(the portal's auto-population and editability rules changed in 2025; verify what is still editable)*.
- Values **actually filed** are entered/imported afterwards and stored beside the prepared ones with the difference. The annual tool reads the filed values, not the prepared ones.
- 3B JSON generation for the portal's offline upload: phase 3 (§12), after GSTR-1 and matching are solid.
- Interest and late fee are out of scope in v1 (show a reminder only).

## 8. Sign-off log: who prepared, who reviewed

Decision (Tobin, 25 Nov 2026): **a simple Excel sheet password is enough.** No hash chain, no server record, no roles table.

- `SignOff` is an append-only list by design: one row per event with `Seq`, `Date-time`, `Person` (picked from a staff list kept on `Lists`, or typed), `Role` (Preparer / Reviewer), `Action` (Created workbook, Imported register, Generated GSTR-1 JSON, Preview checked, Matched, Confirmed ITC decision, Prepared 3B, **Marked prepared**, **Marked reviewed**, Marked filed, Imported vendor confirmation, Exported archive), `Period`, `Object` (file name or row ids), `Content fingerprint` (SHA-256 of the generated JSON, shown for reference so a reviewer can see which file was reviewed), `Remark`. The tool only adds rows; it never edits or deletes old ones.
- The sheet is **protected with a password** (set in code, as the current tool does for its sheets; see the integrity note in `docs/HANDOVER.md`, the password ships in the built file, so this is a deterrent, not security).
- Preparer and reviewer must differ (a flag allows a self-review for one-person clients; it is written in the log as "self-reviewed").
- "Reviewed" is recorded against the fingerprint of the JSON and the period's sales rows; if those rows change afterwards, the period status goes back to Prepared and the log says why.
- **A filed period is locked for good; there is no unlock** (Tobin: once filed, nobody should reopen it). Corrections go in as amendment rows in a later period, as on the portal. The tool refuses edits, voids and re-uploads on rows of a filed period; a locked row's sheet cells are also protected.
- The Instructions sheet says plainly: *"The sign-off log is protected by a sheet password. It prevents accidental changes. It is not tamper-proof."*

## 9. Hosting on Cloudflare (free tier, about 50 users)

Because the filing features live inside the same single-file HTML as the rest of the tool and need no server (§8), hosting is just **static hosting of that one file**:
- **Cloudflare Pages** (free) can serve it, as the Pages preview already does for pull requests; the repo also deploys to GitHub Pages today (PR #10, no login). Both are fine for 50 users; the choice is about login, not capacity.
- **Login (optional):** Cloudflare Access (Zero Trust free plan, advertised as free for up to 50 users *(verify)*) in front of the page gives staff a firm-account login, so the HTML is not public. Not needed by the sign-off log. A later option is to prefill "Person" from the Access identity; this needs a tiny call and is not part of v1.
- Everyone gets the latest release automatically; client workbooks stay on the firm's shared drive or PCs. Nothing about clients is stored on Cloudflare.
- The same HTML still works opened from disk, offline.
- No Worker, no D1, no R2 in v1.

## 10. Architecture: inside the current tool (decided 25 Nov 2026)

### 10.1 Is it viable without making the tool hard to maintain? Yes, with rules.
Assessment against the current code (v0.19: about 17,000 lines of TypeScript in `loader/src`; each tab already owns a folder: `flow/` has engine, workbook, ui and guide; `revenue/` likewise; the wizard (`wizard.ts`) lists tabs, steps and file needs in tables, `STEPS` and `NEEDS`; tests per module):
- A new tab is the pattern the code already supports. The filing features go under **`loader/src/filing/`**, split into separate modules (§10.2), each with the same shape as `flow/` and `revenue/` (engine, workbook writer, ui, guide, tests).
- Matching purchases against 2B is **not a new feature**: Input Reco already does it. The filing tab reuses it and only adds saving the result into the client workbook (match codes) and the blocked-credit and vendor-confirmation steps. This keeps the most complex logic in one place.
- The workbook reader for the annual return sits in the same code base, so the annual adapter is an internal function call, with no second copy of the code to keep in step.
- Size: the filing work is roughly 6,000 to 9,000 new lines (rough estimate; mostly the GSTR-1 generator and validations, the 3B builder, the manual-entry grid and the workbook), i.e. about +40% on today's code. The built HTML is 3.7 MB; the main addition is a small PDF library (the vendor-confirmation PDF), which should be small *(measure before choosing)*.

**Main trade-offs.**
1. *Release coupling.* Every filing fix ships the whole tool, including the annual and reconciliation tabs, and the annual season and monthly filing run at the same time. Mitigation: module boundary rules below, one test suite, and a release checklist that runs the existing real-file cases before every release.
2. *The filing tab is used every month by 50 people; the other tabs a few times a year.* Bugs in the filing tab have a more frequent and more visible cost. Mitigation: more tests there than elsewhere (round-trip GSTR-1, golden matching) and a "known limits" panel.
3. *Complexity of one big wizard.* The tab list is already four tabs plus annual. Mitigation: keep the tab self-contained and put infrequently-used steps (vendor confirmation, preview check, close the year) behind collapsible sections rather than in the main path.

**Rules to keep it maintainable (include in CONTRIBUTING / the handover):**
- Only `wizard.ts` and `main.ts` may know about the filing tab; nothing in `recon.ts`, `register.ts`, `flow/`, `revenue/` imports from `filing/`. Shared helpers go to `util.ts` or a small `common/` folder.
- No behaviour change to existing tabs inside the filing PRs (separate PRs for any refactor, proved by the existing tests and real-file cases).
- Every GSTN-dependent value (field names, thresholds, HSN digit rules, keyword groups) lives in a dated config file (`filing/config/`), not in code.
- Each phase ends with `npm run check` green, tests added, in-app Help updated (`help.ts`), and a size report of the built HTML.
- A size budget: stop and review if the built HTML grows past an agreed limit (suggest 5 MB) or any single file passes about 1,200 lines.
- If it later gets too big or needs a server (more than one person editing the same workbook, or history moved to a database), the module can be split out: it is already self-contained.

**What would have to change first:** nothing blocking. Do before building: (a) freeze the workbook contract; (b) add the new tab id and an empty steps table to `wizard.ts` behind a flag; (c) agree the size budget.

### 10.2 Module layout (Tobin, 27 Nov 2026: each module separate for easy maintenance)

Each module is a folder with its own engine, screens (steps), workbook part, guide text, config and tests, and a **small public interface** (`index.ts`). **Modules do not call each other's code. They share data only through the client workbook tables defined in the contract (§5)**, so any module can be changed, tested or even switched off without touching another.

```
loader/src/filing/
  common/        client workbook open/create/save, contract (sheet and column names, schema version),
                 periods and lock rules, sign-off log writer, Lists/config loader, shared validators (GSTIN, state, rate)
  sales/         sales register mapping, manual entry grid, section-code and return-period derivation, Sales sheet writer
  gstr1/         GSTR-1 JSON generator, pre-generation validation, optional preview comparison, GSTR-1 summary
  purchases/     saving Input Reco matches (match codes) into Purchases / GSTR2B / Match sheets; no matching logic of its own
  itc/           blocked-credit (Sec 17(5)) keyword engine, suggestions, confirmations, schedule
  vendorconfirm/ vendor-wise statement, client PDF summary, client Excel detail, re-import of responses
  gstr3b/        3B working from workbook data, filed-value capture, later the 3B JSON
  config/        dated GSTN config: JSON schema versions, thresholds, HSN rules, keyword groups, section codes
loader/src/filingbridge/   annual-return reader and adapter (reads the workbook, returns the rows the annual tab already uses)
```

| Module | Owns (reads / writes workbook sheets) | May import | Must not import | Has its own |
|---|---|---|---|---|
| `common` | Client, Periods, SignOff, Lists | `util.ts`, `xlsx.ts`, `xlsxbook.ts`, `wpfront.ts`, `state.ts` | any other filing module | tests for contract, locking, sign-off |
| `sales` | writes Sales | `common`, `register.ts` | `gstr1`, `gstr3b`, `itc` | tests for derivation and mapping; Help section "Sales" |
| `gstr1` | reads Sales, Periods; writes Periods (JSON name, fingerprint) | `common`, `gstr1.ts` (the existing reader, for round-trip tests and preview) | `sales` (reads the Sales sheet, not its code) | tests: round trip, validations, preview; Help section |
| `purchases` | writes Purchases, GSTR2B, Match | `common`, `recon.ts`, `gstr2b.ts`, `review.ts`, `settings.ts` | `itc`, `vendorconfirm` | golden test equal to Input Reco |
| `itc` | reads Purchases, GSTR2B; writes ITC | `common`, `config` | `purchases` code | keyword tests with false-positive cases |
| `vendorconfirm` | reads GSTR2B; writes Vendor_Confirm | `common`, PDF library | `itc`, `purchases` code | PDF and Excel round trip |
| `gstr3b` | reads Sales, Purchases, ITC, Match; writes GSTR3B | `common`, `gstr3b.ts` reader | `gstr1` code | tests on demo samples |
| `filingbridge` | reads the whole workbook; writes nothing | `common` contract only, `gstr1.ts`, `gstr2b.ts`, `register.ts` shapes | UI code | equivalence tests (companion plan §7) |

Rules that make this work:
1. **One module per pull request** (plus its Help section and tests); the order is in §12. A module that is not finished can stay behind a flag, and the tab shows only finished modules.
2. A module's `index.ts` exports only: its steps (for `wizard.ts`), its workbook sheet writers and readers, and its config. Anything else is private. A lint / test check (an import-boundary test, run in `npm test`) fails the build if a module imports a forbidden one.
3. The contract (`common/contract.ts`, mirrored in `workbook-contract.md`) is the only coupling. Changing a column means bumping the schema version and updating every module that touches that sheet, so the module table above also says who is affected.
4. Each module has its own section in `help.ts` and its own Instructions text in the workbook, so a change to one module changes one place.
5. Each module carries a short `README.md` in its folder: purpose, sheets read and written, config files, tests, known limits, who to ask. New contributors start there.
6. Config files (`config/`) are dated and versioned per module; a GSTN format change touches `config/` and one module only.
7. Shared logic that two modules need goes to `common` (if it is about the workbook) or `util.ts` (if it is general), never copied.

Failure containment: if `gstr3b` has a bug, `gstr1` and the existing tabs still work. The annual return (`filingbridge`) reads only sheets, so it is unaffected by changes inside a module as long as the contract holds.

### 10.3 Stack and repository
TypeScript + Vite as today, no new framework unless the manual-entry grid needs one (decide in week 1). XLSX writing with `xlsx.ts` / `xlsxbook.ts` / `wpfront.ts`. PDF via a small pinned library, loaded as part of the single file. **Same repository, same loader**: `loader/src/filing/`. This reverses the earlier "new repo" and "new `filing/` folder" options. The workbook contract lives in `docs/filing-tool/workbook-contract.md`, with one fixture generator used by the tests.

### 10.4 Reuse map (what the filing modules take from existing code)
| Need | Reuse |
|---|---|
| register reading and mapping | `register.ts` (+ new sales fields) |
| purchase vs 2B matching, settings, review UI | `recon.ts`, `review.ts`, `settings.ts`, `suppliers.ts` |
| 2B parsing incl. amendments | `gstr2b.ts` |
| GSTR-1 parsing (for preview check and round-trip test) | `gstr1.ts` |
| 3B parsing (compare with filed) | `gstr3b.ts` |
| workbook writing, state sheets, Index/Instructions | `xlsx.ts`, `xlsxbook.ts`, `state.ts`, `wpfront.ts`, `workbook/jjfront.py` |
| progress bar, search, wizard shell | `progress.ts`, `tablesearch.ts`, `wizard.ts` |
| GSTIN checks, FY maths, INR formatting | `util.ts` |
New: sales normaliser with section-code derivation, GSTR-1 generator, validators, preview comparer, 3B builder, blocked-credit engine, vendor confirmation exports, sign-off log writer.

### 10.5 Testing (acceptance)
1. Generate → parse round trip for GSTR-1 on synthetic data covering every section.
2. The GSTN sample payloads in `test/fixtures/gstn/` are reproduced by the generator within tolerance.
3. Workbook grows: three periods loaded, then a fourth; earlier rows byte-identical; edits to a filed period's rows are refused by the tool; the sign-off sheet is protected.
4. Matching results equal the current tool's on the same inputs (golden case from `samples/input-reco`).
5. Real-file cases: anonymised client registers and portal files from Tobin (needed; none yet for sales registers other than the demo).
6. LibreOffice recalculation of the workbook: zero formula mismatches (as `flow_recalc.py`).
7. Performance: a 50,000-row Sales sheet loads, validates and generates JSON in an agreed time on an ordinary office PC.

## 11. Moving history out of Excel later

Do not wait for pain. Decide now:
- `ClientStore` interface: `listPeriods`, `readRows(table, filter)`, `appendRows`, `voidRow`, `appendAudit`, `lockPeriod`. Excel adapter first.
- Every row already has a stable `Row Id`, a period and a lock status, so exporting to a database later is a mechanical copy: tables `sales`, `purchases`, `gstr2b`, `match`, `itc`, `audit` with the same columns.
- Likely targets: Cloudflare **D1** (SQLite; fine for a few million rows across clients, check size limits) or Postgres elsewhere. Then the Excel file becomes a **view/export** (current FY plus open items) while history sits in the database, and the annual tool reads either the workbook or an export.
- The `SignOff` rows carry over unchanged as an event table (if stronger protection is wanted then, a server-side append-only log can be added at that point).
- Triggers to move: any client above the §5.6 thresholds, concurrent editing by several staff, or a need for cross-client queries.
- Costs of moving: client data on a server (confidentiality, backup, DPDP Act obligations *(take advice)*), access control per client, migration tests.

## 12. Suggested phases

| Phase | Content | Done when |
|---|---|---|
| 0 (2 days) | Freeze workbook contract; set up `filing/` skeleton, `common`, import-boundary test, tab behind a flag; decide open questions | contract doc merged, empty tab registered |
| 1 | `sales` (upload, manual entry, section codes) and `gstr1` (JSON, validation), sign-off log in `common` | JSON accepted by the portal on a test client |
| 2 | `purchases` (save Input Reco matches into the workbook), `itc` (blocked credit), `vendorconfirm` | matching equals Input Reco on the golden case |
| 3 | `gstr3b` (working, filed-value capture), optional preview check in `gstr1`, Close-the-year in `common`, optional Cloudflare Access login for the hosted copy | reviewer flow works end to end with 2 people |
| 4 | `filingbridge` (annual-return adapter, companion plan), later 3B JSON | annual numbers unchanged vs the register route; workbook vs portal differences listed |

Each row is one or more pull requests, one module each, with its own tests, Help section and module README.

## 13. Open questions for Tobin

Answered 25 Nov 2026: one workbook per GSTIN; the features go inside the current tool (same repo, `loader/src/filing/`); a filed period is locked for good with no unlock role (corrections are amendments in a later period); portal files stay required for the annual return and win over the workbook; the demo samples are the test data for now.

Still open (3 is answered, kept for numbering):
3. ~~Immutability~~ Answered 25 Nov 2026: a simple Excel sheet password on the sign-off log is enough (§8).
4. Real samples needed: a client sales register (anonymised), a GSTR-1 preview / summary the portal gives, a portal 3B JSON, a GSTR-1 filed JSON with CDNR and amendments.
5. Can a small one-person client be self-reviewed (logged as such)?
6. Which of QRMP quarterly filers are in scope for v1 (IFF)? 
7. PDF font: Aptos is not freely embeddable; which font for the client PDF?
8. Do clients ever have e-invoice IRN data to import (e-invoice is a separate JSON/portal flow)? Out of scope in v1 unless asked.

---

# Part B. Prompt to give the build session

> Copy everything between the lines into a new Claude session with the repo `tobinjames/gst-reco-tool` attached. It is meant to work on its own.

---

You are building a new tool inside the GitHub repo `tobinjames/gst-reco-tool` (private; owner is James & James LLP, Tobin James). Read first, in this order: `docs/HANDOVER.md`, `docs/filing-tool/01-filing-tool-prompt-and-handover.md` (Part A is your specification and says what is open), `docs/filing-tool/02-annual-return-integration-plan.md`. Then read `loader/src/register.ts`, `recon.ts`, `gstr2b.ts`, `gstr1.ts`, `xlsxbook.ts`, `state.ts`, `wpfront.ts`, `util.ts` and `loader/test/`.

**What to build.** New features inside the existing loader (`loader/`), as a new tab under `loader/src/filing/`, split into separate modules (`common`, `sales`, `gstr1`, `purchases`, `itc`, `vendorconfirm`, `gstr3b`, `config`, plus `filingbridge` for the annual return) as laid out in §10.2 of the handover: each module has its own folder, engine, steps, tests, Help section and README, and modules talk to each other only through the workbook sheets of the contract, never through each other's code. Follow the pattern of `flow/` and `revenue/` and register the tab through `wizard.ts`. Do not create a separate app. Reuse `register.ts`, `recon.ts`, `gstr2b.ts`, `gstr1.ts`, `xlsxbook.ts`, `state.ts`, `wpfront.ts`, `util.ts` and the Input Reco matching rather than rewriting them. One Excel workbook per client GSTIN holds all periods and grows with every filing. It must: take a sales register or manual entry; derive a section code (b2b, b2cl, b2cs, …) and return period per row, editable by the user; validate and generate the GSTR-1 JSON in the current GSTN format (use official sample payloads, never guess field names); optionally compare a portal GSTR-1 preview with the data; keep purchase-register vs 2B match codes in the workbook using the Input Reco matching; produce client vendor confirmation (PDF summary and Excel detail, re-importable) when there is no purchase register; suggest blocked credit (Sec 17(5)) by keywords showing the hit and the usual exceptions, which the user confirms or overrides; prepare GSTR-3B figures and record the values actually filed; keep a sign-off log (who prepared, who reviewed) in a sheet protected with a password. A filed period is locked for good (corrections are amendments in a later period). The workbook layout in §5 of the handover is a contract with the annual return: do not change it without a schema version bump and a migration.

**Rules.**
- Plan before coding. Post the plan (phases, risks, questions) and wait for Tobin's go. Build in the phases of §12; each phase is a pull request that passes `npm run check`.
- Be honest about limits in the UI and docs: the sign-off log is only sheet-password protected; the tool does not file; the workbook is not the portal.
- Never hard-code thresholds, HSN digit rules, field lists or keyword groups; keep them in dated config, and mark anything uncertain *(verify)* and check it against the GST portal or developer portal.
- Workbooks: Aptos font, Index sheet with links, one Instructions sheet, Summary sheet, "Back to Index" links, J&J colour #2C388E and the logo from `loader/src/brand/`, one plain readable formula per cell (SUMIFS / INDEX / MATCH; no macros, no Excel 365 functions). Prepared by / Reviewed by block on Index.
- Client files never leave the user's PC; there is no server component.
- Keep every module self-contained (§10.1 and §10.2 rules): nothing outside `wizard.ts` / `main.ts` imports from `filing/`; modules never import each other (an import-boundary test enforces it); one module per pull request; no behaviour change to existing tabs in filing PRs; report the built HTML size in each PR.
- Tests: round-trip GSTR-1 generation through `gstr1.ts`; locked-period tests (edits to a filed period are refused); matching golden case equal to the current tool; LibreOffice recalculation with zero mismatches; real-file cases only after anonymising GSTINs and names (scripts/anonymise.ts). Do not commit real client data.
- Do not produce source zips. Deliver the built HTML (release workflow) and the PR link. Update `loader/src/help.ts`-style in-app help for every user-visible change.
- Ask Tobin about the open questions in §13 before building the parts that depend on them; start with the parts that do not.

First deliverable: the frozen workbook contract (`docs/filing-tool/workbook-contract.md`), a fixture workbook generated by a script, the empty tab registered behind a flag, and the list of samples you still need from Tobin.

---

# GST Filing Tool (GSTR-1 and GSTR-3B preparation): build prompt and handover

James & James LLP · planning document, 25 Nov 2026 · for the build planned the week of 30 Nov 2026
Status: **plan only, nothing is built.** Companion: [02-annual-return-integration-plan.md](02-annual-return-integration-plan.md).

This file has two parts. **Part A** is a self-contained brief for the person (or Claude session) who builds the tool. **Part B** is the prompt to paste into a new Claude session. Part A can be read without Part B and vice versa.

---

# Part A. Handover

## 1. Goal in one paragraph

A new tool for the firm's staff to **prepare GSTR-1 and GSTR-3B for clients every month (or quarter)**. Each client has **one Excel workbook** that holds all of that client's data across months and years and **grows every time a return is filed**. Staff load a sales register (or key in sales by hand for small clients), the tool generates the **GSTR-1 JSON** for upload to the GST portal, optionally checks the portal's GSTR-1 preview against the data, matches the purchase register to GSTR-2B (same logic as the current tool), prepares GSTR-3B figures, and records **who prepared and who reviewed** each return in a tamper-evident log. At year end the same workbook is the **input to the current annual return tool** (see companion plan).

## 2. What the person asked for (checklist the build must cover)

| # | Requirement (Tobin, 25 Nov 2026) | Where it is covered |
|---|---|---|
| 1 | One Excel file per client holding multiple months and years; every filing adds to it | §5 workbook contract, §6 lifecycle |
| 2 | Today's clients are small; later the history may have to move out of Excel | §5.6 scale limits, §11 moving history out |
| 3 | Upload a sales register, **or** enter records one by one (small clients without a proper register) | §7.1 |
| 4 | Generate the GSTR-1 JSON for filing | §7.2 |
| 5 | Optional: upload the GSTR-1 **preview** before filing and validate it against the data | §7.3 |
| 6 | Immutable record of who prepared and who reviewed, inside the file | §8 (with honest limits) |
| 7 | Hosting on Cloudflare free tier, about 50 users | §9 |
| 8 | Upload purchase register and GSTR-2B and match them, using the current tool's logic | §7.4 |
| 9 | Purchase register missing: client confirms vendor-wise; summary PDF plus detailed Excel for the client | §7.5 |
| 10 | Detect non-eligible ITC from purchase register / 2B by keywords (insurance, vehicle, auto, hotel, restaurant, others); only a suggestion, user confirms | §7.6 |
| 11 | At year end the single file feeds the current annual return tool | §5, companion plan |
| 12 | Also GSTR-3B preparation (title of the project) | §7.7 |

## 3. Honest limits (read before promising anything to staff)

1. **A plain Excel file cannot be truly immutable.** Anyone with Excel can edit any cell, including a "log" sheet, and recompute any checksum. What we can offer:
   - *Tamper-evident*: a hash chain in the workbook, so any change to past entries is detectable by the tool when the file is opened.
   - *Tamper-resistant*: each log entry's hash is also sent to a small server log (Cloudflare, append-only, identity taken from Cloudflare Access, not typed by the user) which returns a signed receipt stored in the workbook. Editing the workbook then also needs the server record to be edited, which users cannot do. This is the only route to a meaningful "immutable" claim. Even then, the people who run the Cloudflare account can in principle alter the server log, so it is audit-grade for the firm, not court-grade.
   - Without server contact (offline use) entries are marked **"not anchored"** and anchored at the next online save.
2. **The tool cannot file anything.** It produces a JSON that a person uploads on the GST portal and files with DSC/EVC. "Filed" status and ARN are recorded by the user (or read from a downloaded acknowledgement), not observed.
3. **The workbook is a record of what we prepared**, not necessarily of what the portal holds. Portal edits, amendments made online, IMS actions and late 2B changes can differ. The optional preview check (§7.3) and the annual tool's comparison with downloaded returns exist for this reason.
4. **GSTN formats and rules change** (the GSTR-1 JSON moved to v5.0 in Feb 2026; B2CL limits, HSN digit rules and 3B auto-population have all changed in the last two years). Every threshold and field list must come from a **config file or the downloaded GSTN sample**, never hard-coded, and be dated. Anything in this document marked *(verify)* must be checked against the GST portal / developer portal when building. Fetch recipes are in project memory (`gstn-api-schemas`).
5. **Excel is not a database.** Limits in §5.6. Design so that Excel is one storage adapter among several.
6. **Cloudflare free-tier numbers change.** Quotas quoted here are from memory *(verify)*.

## 4. Principles carried over from the current tool

- Browser-only processing; client files stay on the user's PC. Only hashes (no client data) go to the server (§8, §9). This keeps the "files never leave the PC" promise and limits confidentiality risk.
- No macros, no Excel 365-only functions. SUMIFS / INDEX / MATCH, **one readable formula per cell**; someone without the tool must be able to follow the working (Tobin, workpaper-format).
- **Aptos** font throughout; **Index** sheet with links; **one Instructions sheet**; Summary sheets; "Back to Index" link on every sheet; Prepared/Reviewed sign-off block. Colour #2C388E, James & James logo (loader/src/brand/).
- Saved workbook restores everything (hidden state sheets, as `_JJState` / `_JJFiles` in the current tool). Portal returns are permanent; registers and ledgers may be replaced.
- Reasons / decisions are derived or chosen from lists, with logged overrides; no free-text accounting reasons.
- Keep `npm run check` green, add real-file test cases (anonymised), and update the in-app Help whenever logic changes.
- Do not produce source zips; the repo is the source of truth. Deliver the built HTML and the PR link.

## 5. Workbook contract (version 1)

This is the interface between the new tool and the current annual tool. Treat it as an API: **change it only by bumping `schema` and writing a migration.** Sheet and column names below are the proposal; the build may refine them, but the annual tool reads whatever is finally agreed, so freeze it early (week 1) and put the final version in `docs/filing-tool/workbook-contract.md`.

### 5.1 Files and naming
`Filing_<GSTIN>.xlsx`, **one workbook per GSTIN** (the current tool is single-GSTIN per run; a client with several GSTINs gets several workbooks and a shared client name). *Open question 1.*

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
| `Audit` | the hash-chained log (§8) |
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

1. Open the client workbook (or create it: GSTIN, name, frequency). Opening checks schema version, verifies the audit chain and server anchors, and shows warnings (broken chain, missing period, near row limits).
2. Pick the return period. Load/enter **sales** (§7.1); the tool proposes section codes; the user fixes exceptions; validation list (§7.2) is shown.
3. Generate **GSTR-1 JSON**; optional preview check (§7.3). Log: *Prepared*.
4. A second person opens the workbook, reviews the validation and summaries, and records *Reviewed* (cannot be the preparer; configurable for one-person clients, then logged as "self-reviewed").
5. Staff upload the JSON on the portal and file. They record *Filed*, ARN and date in `Periods` (or drop the acknowledgement). Rows of that period become **Locked**. Corrections later go in as amendment rows in a later period, never as edits to locked rows (the tool refuses edits on locked rows; the log shows any bypass attempt it can detect).
6. Load **purchase register** and **GSTR-2B** JSON for the period; auto-match; confirm manual matches; keyword suggestions for blocked credit; vendor confirmation if there is no register (§7.5).
7. Prepare **GSTR-3B** (§7.7); review; record filed values.
8. Save. The workbook has grown by one period. Optional: send a **hash receipt** to the server.

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

## 8. Audit log: who prepared, who reviewed

### 8.1 Event model
Append-only `Audit` sheet. One entry per event: `Seq`, `Timestamp (UTC)`, `Actor` (email), `Role` (Preparer / Reviewer / Admin), `Action` (Created workbook, Imported register, Changed row, Voided row, Generated GSTR-1 JSON, Preview checked, Matched, Confirmed ITC decision, Prepared 3B, **Marked prepared**, **Marked reviewed**, Marked filed, Unlocked, Imported vendor confirmation, Exported archive), `Period`, `Object` (row ids or file name), `Content hash` (SHA-256 of the file generated or of a canonical dump of the rows affected), `Prev hash`, `Entry hash`, `Server receipt`.

`Entry hash = SHA-256(Seq | Timestamp | Actor | Role | Action | Period | Object | Content hash | Prev hash)`. Seq 1's `Prev hash` is the hash of the Client sheet at creation.

### 8.2 Three layers (build in this order)
1. **Chain in the file.** On every save the tool appends entries and recomputes nothing old. On open, it re-verifies the chain and the content hashes it can recompute (e.g. the `Sales` rows of a filed period against the hash recorded at filing). A break shows a red banner and is logged on the next save. *Tamper-evident only.*
2. **Server anchor.** The browser posts `(workbook id, Seq, Entry hash)` to a Cloudflare Worker. The Worker takes the user's identity from the **Cloudflare Access** JWT (not from the request body), stores the record in an append-only table (D1) and returns a signed receipt (HMAC or Ed25519; key in a Worker secret). The receipt goes into `Audit`. On open, the tool asks the Worker to confirm the latest N receipts. Rewriting history now needs the server's cooperation. No client data crosses; only hashes, actor, time, GSTIN, period.
3. **Optional later**: periodic anchoring of the log head to something outside the firm's control (e.g. an e-mail to the client or a public timestamping service) if a stronger claim is ever needed.

### 8.3 Rules
- Preparer and reviewer must differ (config can allow a self-review flag).
- "Reviewed" is recorded against a **content hash of the JSON and of the period's Sales rows**; any later change to those rows invalidates the review and the Periods status returns to Prepared (visible in the log).
- Unlocking a filed period needs Admin role, a reason from a list, and creates an entry.
- Say plainly in the Instructions sheet: *"This log is tamper-evident. It is not tamper-proof."*
- Role assignment lives in a Worker-side table (KV/D1) keyed by Access email; the workbook only displays the roles it was told by the server (offline: last known roles, flagged).

## 9. Hosting on Cloudflare (free tier, about 50 users)

Recommended shape (verify current plans and quotas):
- **Cloudflare Pages** serves the static tool (single HTML as today, or small bundle). Free.
- **Cloudflare Access** (Zero Trust free plan, advertised as free for up to 50 users) in front of Pages and the Worker, with login by firm Google/Microsoft account or one-time PIN. Gives the verified user email used in the audit log and for roles. The handover already names `gst.jjglobal.in` behind Access as the intended host; the current tool is on GitHub Pages with no login (PR #10): moving to Access is a prerequisite for the audit log, not a nicety.
- **Pages Functions / Worker** with **D1** (SQLite) for the anchor log, role table and receipts. Free quotas are small but the audit traffic is tiny (a few entries per return) *(verify)*.
- **R2** (optional) for storing client workbooks on the server: not needed in v1 because staff keep the workbooks on the firm's shared drive as today. If chosen later, it must be private, access-controlled, backed up, and a fresh decision on confidentiality (client financial data in a third-party cloud).
- Do not put client data (names, invoices) in D1 or logs in v1.
- Offline use must still work (single HTML opened from disk): everything except anchoring and roles; entries are marked "not anchored".
- Cost warning: if the audit server becomes business-critical, the free tier has no SLA; plan for a paid plan or an alternative.

## 10. Architecture and build notes

### 10.1 Stack
TypeScript + Vite single-file build like `loader/`, no framework (or a small one if the manual-entry grid needs it; decide in week 1). XLSX writing with the existing `xlsx.ts` / `xlsxbook.ts` (formulas with cached values, Aptos, Index, Instructions via `wpfront.ts`). PDF via a small pinned library.

### 10.2 Repository
Recommendation: **same repo**, new top-level folder `filing/` next to `loader/`, importing the shared engines from `loader/src` by path alias at first (no refactor in the first PR), then extract a `core/` package when the second consumer is stable. Alternative: separate repo with a copied core (drift risk, rejected). *Open question 2.*

### 10.3 Reuse map (current tool → new tool)
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
New: sales normaliser with section-code derivation, GSTR-1 generator, validators, preview comparer, 3B builder, blocked-credit engine, vendor confirmation exports, audit chain + anchor client, Worker.

### 10.4 Testing (acceptance)
1. Generate → parse round trip for GSTR-1 on synthetic data covering every section.
2. The GSTN sample payloads in `test/fixtures/gstn/` are reproduced by the generator within tolerance.
3. Workbook grows: three periods loaded, then a fourth; earlier rows byte-identical; chain verifies; tamper (edit a value, delete a row, reorder) is detected.
4. Matching results equal the current tool's on the same inputs (golden case from `samples/input-reco`).
5. Real-file cases: anonymised client registers and portal files from Tobin (needed; none yet for sales registers other than the demo).
6. LibreOffice recalculation of the workbook: zero formula mismatches (as `flow_recalc.py`).
7. Performance: a 50,000-row Sales sheet loads, validates and generates JSON in an agreed time on an ordinary office PC.

## 11. Moving history out of Excel later

Do not wait for pain. Decide now:
- `ClientStore` interface: `listPeriods`, `readRows(table, filter)`, `appendRows`, `voidRow`, `appendAudit`, `lockPeriod`. Excel adapter first.
- Every row already has a stable `Row Id`, a period and a lock status, so exporting to a database later is a mechanical copy: tables `sales`, `purchases`, `gstr2b`, `match`, `itc`, `audit` with the same columns.
- Likely targets: Cloudflare **D1** (SQLite; fine for a few million rows across clients, check size limits) or Postgres elsewhere. Then the Excel file becomes a **view/export** (current FY plus open items) while history sits in the database, and the annual tool reads either the workbook or an export.
- The server hash-chain from §8 carries over unchanged and becomes the primary log.
- Triggers to move: any client above the §5.6 thresholds, concurrent editing by several staff, or a need for cross-client queries.
- Costs of moving: client data on a server (confidentiality, backup, DPDP Act obligations *(take advice)*), access control per client, migration tests.

## 12. Suggested phases

| Phase | Content | Done when |
|---|---|---|
| 0 (2 days) | Freeze workbook contract; decide open questions; get real samples | contract doc merged |
| 1 | Workbook create/open/append, Sales upload + manual entry, section codes, validation, **GSTR-1 JSON**, audit chain (layer 1), Periods | JSON accepted by the portal on a test client |
| 2 | Purchases + 2B + matching + `Match`; blocked-credit suggestions; vendor confirmation PDF+Excel | matching equals current tool on golden case |
| 3 | 3B working, filed-value capture, optional preview check, server anchor + roles (Cloudflare Access), Close-the-year | reviewer flow works end to end with 2 users |
| 4 | Annual-return adapter in the current tool (companion plan), 3B JSON | annual numbers equal the raw-file route |

## 13. Open questions for Tobin

1. One workbook per **GSTIN** or per **client** (several GSTINs)? Recommended: per GSTIN.
2. Same repo (`filing/` folder) or a new repo? Recommended: same repo.
3. Can staff accept that "immutable" means tamper-evident plus server anchor, with the server layer in phase 3 and Cloudflare Access login mandatory?
4. Real samples needed: a client sales register (anonymised), a GSTR-1 preview / summary the portal gives, a portal 3B JSON, a GSTR-1 filed JSON with CDNR and amendments.
5. Who may unlock a filed period (Admin only)? Can a small client be self-reviewed?
6. Which of QRMP quarterly filers are in scope for v1 (IFF)? 
7. PDF font: Aptos is not freely embeddable; which font for the client PDF?
8. Do clients ever have e-invoice IRN data to import (e-invoice is a separate JSON/portal flow)? Out of scope in v1 unless asked.

---

# Part B. Prompt to give the build session

> Copy everything between the lines into a new Claude session with the repo `tobinjames/gst-reco-tool` attached. It is meant to work on its own.

---

You are building a new tool inside the GitHub repo `tobinjames/gst-reco-tool` (private; owner is James & James LLP, Tobin James). Read first, in this order: `docs/HANDOVER.md`, `docs/filing-tool/01-filing-tool-prompt-and-handover.md` (Part A is your specification and says what is open), `docs/filing-tool/02-annual-return-integration-plan.md`. Then read `loader/src/register.ts`, `recon.ts`, `gstr2b.ts`, `gstr1.ts`, `xlsxbook.ts`, `state.ts`, `wpfront.ts`, `util.ts` and `loader/test/`.

**What to build.** A browser-based GST filing preparation tool (GSTR-1 first, then purchase/2B matching, then GSTR-3B), in a new folder `filing/`, importing the matching, register, 2B and workbook-writing code from `loader/src` rather than rewriting it. One Excel workbook per client GSTIN holds all periods and grows with every filing. It must: take a sales register or manual entry; derive a section code (b2b, b2cl, b2cs, …) and return period per row, editable by the user; validate and generate the GSTR-1 JSON in the current GSTN format (use official sample payloads, never guess field names); optionally compare a portal GSTR-1 preview with the data; match purchase register to GSTR-2B with the current tool's logic and store links as match codes in both tables; produce client vendor confirmation (PDF summary and Excel detail, re-importable) when there is no purchase register; suggest blocked credit (Sec 17(5)) by keywords showing the hit and the usual exceptions, which the user confirms or overrides; prepare GSTR-3B figures and record the values actually filed; keep a hash-chained audit log of who prepared and who reviewed in the workbook, with server anchoring through Cloudflare Access plus a Worker and D1 (hashes only, no client data). The workbook layout in §5 of the handover is a contract with the annual return tool: do not change it without a schema version bump and a migration.

**Rules.**
- Plan before coding. Post the plan (phases, risks, questions) and wait for Tobin's go. Build in the phases of §12; each phase is a pull request that passes `npm run check`.
- Be honest about limits in the UI and docs: the log is tamper-evident, not tamper-proof; the tool does not file; the workbook is not the portal.
- Never hard-code thresholds, HSN digit rules, field lists or keyword groups; keep them in dated config, and mark anything uncertain *(verify)* and check it against the GST portal or developer portal.
- Workbooks: Aptos font, Index sheet with links, one Instructions sheet, Summary sheet, "Back to Index" links, J&J colour #2C388E and the logo from `loader/src/brand/`, one plain readable formula per cell (SUMIFS / INDEX / MATCH; no macros, no Excel 365 functions). Prepared by / Reviewed by block on Index.
- Client files never leave the user's PC; only hashes go to the server. No client data in logs.
- Tests: round-trip GSTR-1 generation through `gstr1.ts`; tamper tests for the chain; matching golden case equal to the current tool; LibreOffice recalculation with zero mismatches; real-file cases only after anonymising GSTINs and names (scripts/anonymise.ts). Do not commit real client data.
- Do not produce source zips. Deliver the built HTML (release workflow) and the PR link. Update `loader/src/help.ts`-style in-app help for every user-visible change.
- Ask Tobin about the open questions in §13 before building the parts that depend on them; start with the parts that do not.

First deliverable: the frozen workbook contract (`docs/filing-tool/workbook-contract.md`) and a fixture workbook generated by a script, plus the list of samples you still need from Tobin.

---

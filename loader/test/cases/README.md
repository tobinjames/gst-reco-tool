# Real-file test cases

Each folder here is one client situation, kept as a permanent check. It holds:

- the files you would drop into the Loader (GSTR-2B / GSTR-1 JSON, purchase or sales register, 3B PDFs…)
- `case.json`: which mode and FY, e.g. `{ "mode": "itc", "fy": "2025-26", "about": "Tally columnar register, credit notes in 2B" }`
- `expected.json`: what the tool produced when the case was last checked by hand: how each file was recognised,
  which register columns were picked, every invoice's match status, the ITC position or GSTR-9 summary, and the warnings.

`npm test` re-runs every case the same way the Loader page does (automatic column mapping, Standard matching) and
fails if anything in `expected.json` changed, listing the lines that differ.

## Adding a case from real files

1. Put the client's files in a folder on your computer (only the ones the case needs).
2. From `loader/`, run
   `npx tsx scripts/anonymise.ts <that folder> <case-name> --mode itc --fy 2025-26`
   This copies them into `test/cases/<case-name>/` with every GSTIN replaced by a made-up one (consistently, so
   matching is unaffected). Amounts, dates, invoice numbers and names stay. If the case has 3B PDFs, add
   `--keep <client GSTIN>`, because PDFs are copied unchanged.
3. Edit `case.json` (`about`: what is special about this case, no client name).
4. Run `npm run test:update`. It writes `expected.json`. **Read it** and compare with what you know is right
   (matched / not matched invoices, ITC totals). Correct the tool first if something is wrong.
5. Commit the folder. From then on, any change that alters these results shows up as a failed test.

## When a test fails after a change

The output lists what changed (e.g. `purchases.lines: now Matched | ... `). If the change is an improvement you
expected, run `npm run test:update` and commit the new `expected.json` with the change. If not, it's a bug.

Folders starting with `_` are ignored (useful for work in progress). Register mappings the user would correct by
hand on the page are not supported yet: a case uses the automatic mapping.

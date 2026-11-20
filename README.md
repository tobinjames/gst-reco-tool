# GST Annual Return & ITC Reconciliation Tool

James & James LLP. Offline, single-page HTML tool for monthly ITC reconciliation (GSTR-2B vs purchase register) and GSTR-9 annual return preparation, with a macro-free Excel workbook template.

- `loader/` Vite + TypeScript app, builds to one HTML file
- `workbook/` Python (openpyxl) scripts that generate the Excel templates
- `release/GST_Annual_Return_Loader.html` the current built tool (Loader v0.7, template 2026.11.20-7)
- `docs/HANDOVER.md` full handover: architecture, key decisions, open items

## Build

```
cd loader
npm install
npx vite build   # outputs dist/index.html
```

The Python scripts in `workbook/` still use absolute paths (`/home/claude/...`); adjust them before running.

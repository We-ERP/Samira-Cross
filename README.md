# Samira Cross-Match v2.0

Advanced Data Cross-Referencing System for FTTH, DSL & IR Excel sheets.

## Features
- **Drag & Drop** upload for all 3 sheets
- **Progress bar** with real % during file load and cross-match
- **Chunked async processing** — handles large files (70MB+ FTTH) without freezing
- **Refresh** button — re-run match without re-uploading files
- **Reset** button — clear everything and start fresh
- **Stats dashboard** — Total / Same / Changed / Not Avail with percentages
- **Filter chips** — quickly filter by match result
- **Search** — search across all columns instantly
- **Sortable columns** — click any header to sort
- **Color-coded badges** in the table
- **Export to Excel** with full colors and formatting
- Preview capped at 500 rows in-browser (export for full data)

## File Structure
```
samira-crossmatch/
├── index.html   ← main page
├── style.css    ← all styles
├── script.js    ← all logic
└── README.md
```

## Usage
1. Open `index.html` in any modern browser (Chrome recommended for large files)
2. Upload FTTH, DSL, and IR sheets (`.xlsx`, `.xls`, or `.csv`)
3. Click **Run Cross-Match**
4. Use filter chips / search to explore results
5. Click **Export to Excel** for the full colored report
6. Use **Refresh** to re-run the match (e.g., after a config change)
7. Use **Reset** to clear all data and start over

## Required Columns
| Sheet | Required Columns |
|-------|-----------------|
| FTTH  | `Service Number`, `Service Name` |
| DSL   | `Service Number`, `Service Name` |
| IR    | `dsl_number`, `case_sub_type` |

## Matching Logic
- Service numbers are normalized (strips `FBB`/`FV` prefixes, trims whitespace, uppercased)
- `case_sub_type` from IR is compared (case-insensitive) to `Service Name` from DSL/FTTH
- Results: **The same** / **Changed** / **Not avail**

## Hosting on GitHub Pages
1. Push all files to a GitHub repo
2. Go to **Settings → Pages → Source: main branch**
3. Done — your tool is live at `https://username.github.io/repo-name`

## Tech Stack
| Library | Purpose |
|---------|---------|
| [SheetJS (xlsx)](https://sheetjs.com/) | Reading Excel/CSV files |
| [ExcelJS](https://github.com/exceljs/exceljs) | Writing colored Excel output |
| [FileSaver.js](https://github.com/eligrey/FileSaver.js/) | Browser file download |
| Plain HTML + CSS + JS | No framework, no build step needed |

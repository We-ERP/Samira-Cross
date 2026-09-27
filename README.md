# Samira Cross-Match v2.0

Advanced Data Cross-Referencing System for FTTH, DSL, and IR Excel sheets.

## Features
- Drag & Drop file upload for all 3 sheets
- Fast O(1) lookup using Maps
- Stats dashboard (Total / Same / Changed / Not Avail)
- Filter chips + search across all columns
- Sortable columns (click any header)
- Color-coded result badges
- Export to Excel with full formatting & colors
- Previews up to 500 rows in-browser (export for full data)

## Usage
1. Open `index.html` in any modern browser
2. Upload FTTH, DSL, and IR sheets (`.xlsx`, `.xls`, or `.csv`)
3. Click **Run Cross-Match**
4. Filter / search results as needed
5. Click **Export to Excel** for the full colored report

## Column Requirements
| Sheet | Key Column |
|-------|-----------|
| FTTH  | `Service Number`, `Service Name` |
| DSL   | `Service Number`, `Service Name` |
| IR    | `dsl_number`, `case_sub_type` |

## How Matching Works
- Service numbers are cleaned (removes `FBB`, `FV` prefixes, trims whitespace)
- `case_sub_type` from IR is compared to `Service Name` from DSL/FTTH
- Result: **The same** / **Changed** / **Not avail**

## Tech Stack
- [SheetJS (xlsx)](https://sheetjs.com/) — reading Excel files
- [ExcelJS](https://github.com/exceljs/exceljs) — writing colored Excel output
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) — file download
- Plain HTML + CSS + JS (no framework, no build step)

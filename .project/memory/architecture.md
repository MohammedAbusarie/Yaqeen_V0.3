# System Architecture

## Module Map
├── Static Web App (Vanilla JS / ESM)
│ ├── Purpose: Client-side attendance report generator (no backend)
│ ├── Entry point: `index.html` → `<script type="module" src="./app.js">`
│ ├── Responsibilities:
│ │  - Collect inputs (XLSX + student IDs + week + type)
│ │  - Process workbook in-browser (SheetJS)
│ │  - Generate report JSON/TXT and render a report view
│ │  - Export PDF (jsPDF + AutoTable)
│ └── External Dependencies (CDN, globals):
│    - SheetJS (XLSX): `window.XLSX`
│    - jsPDF: `window.jspdf`
│    - AutoTable plugin: `doc.autoTable(...)`
│
├── UI Wiring Layer
│ ├── File: `app.js`
│ ├── Responsibilities:
│ │  - Cache DOM element references
│ │  - Initialize state + viewer
│ │  - Wire event handlers
│ └── Depends on: `src/*`, `viewer.js`
│
├── Event Handlers / Orchestration
│ ├── File: `src/handlers.js`
│ ├── Responsibilities:
│ │  - Validate inputs
│ │  - Load XLSX (upload or URL download)
│ │  - Load/parse student IDs file
│ │  - Run processing, generate report, update viewer
│ │  - Export report files (JSON/TXT/PDF)
│ │  - **Editor workflow**: load workbook to memory, list selectable headers, build preview, allow fixups, confirm + download modified XLSX
│ └── Depends on: `attendance.js`, `report.js`, `src/*`
│
├── Core Processing (Pure-ish Logic + Explicit Errors)
│ ├── File: `attendance.js`
│ ├── Responsibilities:
│ │  - Normalize + parse student IDs
│ │  - Parse grades input (`id,grade`)
│ │  - Convert Google Sheet URL → XLSX export URL
│ │  - Download XLSX (best-effort; CORS may block)
│ │  - Read workbook via SheetJS and scan sheets for target week/type
│ │  - **Editor core**: list column headers (rows 2–5) across single/multi scope, detect lecture/section bounds, compute preview rows, apply edits
│ └── Key types:
│    - `FoundLogEntry`, `ParsedStudentIds`
│    - Errors: `ValidationError`, `DownloadError`, `FileError`, `ProcessingError`
│
├── Report Generation
│ ├── File: `report.js`
│ └── Responsibilities:
│    - Build `ReportData` JSON contract
│    - Generate human-readable TXT report
│
├── Report Rendering
│ ├── File: `viewer.js`
│ └── Responsibilities:
│    - Render report table + summary
│    - Optional “Ordered Input” mode if `ordered_rows` exists
│
├── Editor Preview Rendering (inline in Inputs view)
│ ├── Files: `index.html`, `src/handlers.js`
│ └── Responsibilities:
│    - Render preview table (grouped-by-sheet or input-order)
│    - Fix dialogs (search by ID/name) and manual grade edits (preview-only)
│    - **Search & pick** input method: search students by ID/name from workbook, build chosen list; supports attendance and grade (grade entered at add time); duplicates allowed with warning
│
├── Sheet Merger (New Feature)
│ ├── File: `src/sheetMerger.js`
│ ├── Responsibilities:
│ │  - Scan headers from rows 1-5 across all sheets
│ │  - Extract all columns with sample data
│ │  - Build drag-and-drop mapping matrix
│ │  - Merge columns sequentially (Sheet1 rows, then Sheet2 rows, etc.)
│ │  - Generate merged XLSX workbook
│ │  - Optional: eliminate duplicate header rows
│ └── Key functions:
│    - `scanHeadersFromRows()`: detect headers in first 5 rows
│    - `extractAllColumns()`: get all columns with metadata
│    - `mergeColumnsSequentially()`: concatenate mapped columns
│    - `generateMergedWorkbook()`: export merged data as XLSX
│
└── Internal Utilities
   ├── `src/state.js`: in-memory state container + reset
   ├── `src/navigation.js`: view switching (Inputs/Report/OCR/SheetMerger/About)
   ├── `src/metadata.js`: report metadata + safe filenames
   ├── `src/fileRead.js`: File → ArrayBuffer/Text helpers
   ├── `src/dom.js`: DOM id lookup + blob download helper
   ├── `src/uiStatus.js`: status + loading UI
   ├── `src/ocr.js`: OCR processing with Tesseract.js
   └── **Documentation**: `docs/` — project docs; legacy theme colors in `docs/theme-legacy-colors.md` (see manifest for current identity)

## Data Models
- **Student IDs input (`.txt`) → `ParsedStudentIds`** (`attendance.js`)
  - `orderedEntries`: array of `{type:'title'}` and `{type:'id'}` entries
  - `targetIdsSet`: unique IDs used for matching
  - `idCounts`: global counts for duplicate IDs
  - `sectionIdCounts`: per-section counts (if titles/sections used)
- **Search & pick input** (`state.editor.chosenStudents`): array of `{ id, name?, sheet?, row1?, grade? }`; converted to same shapes as file/textarea for `computeEditorPreview`
- **Workbook scan result** (`attendance.js`)
  - `found_log`: `FoundLogEntry[]` (sheet, id, name, cell, header/data row hints)
  - `not_found_ids`: `string[]`
  - `available_weeks`: `string[]` (e.g., `["W1","W2",...]`)
- **Report JSON contract → `ReportData`** (`report.js`)
  - `metadata`: timestamp/week/type/sheet_url + totals
  - `sheets`: per sheet `{name, column, header_row, students[]}`
  - `not_found`: `string[]`
  - optional: `input_order`, `ordered_rows`, `id_counts`
- **Sheet Merger Data Models** (`src/sheetMerger.js`)
  - `ColumnInfo`: `{sheet, columnIndex, headerText, headerRow, columnLetter, sampleValues, key}`
  - `MergedData`: `{rows, headers, totalRows, totalColumns, sourceSheets}`
  - Mapping structure: `{[sheetName]: {[position]: columnKey}}`

## API Contract
- **No backend API**. Everything runs in-browser.
- **External network request** (optional): Download XLSX export from a Google Sheet URL
  - Method: `GET`
  - URL: derived from `/edit` → `/export?format=xlsx` (`attendance.js`)
  - Auth: none (relies on sheet sharing settings)
  - Known limitation: can fail due to browser CORS restrictions

## Deployment & Environment
- **Static hosting**: Netlify (recommended)
  - Publish directory: `web` (per `README.md`)
- **Environment variables**: none used

## Scalability & Limits
- **User capacity**: client/browser-limited (single-user tool)
- **Request rate**: N/A (no backend)
- **Data limits**:
  - XLSX size limited by browser memory/CPU
  - Large sheets may take noticeable time to scan (loops across sheets/rows)

## Last Updated
2026-03-11 | Docs folder added; theme identity documented (see `.project/memory/manifest.md` and `docs/theme-legacy-colors.md` for legacy palette)



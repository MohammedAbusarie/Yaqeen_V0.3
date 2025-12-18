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
│
└── Internal Utilities
   ├── `src/state.js`: in-memory state container + reset
   ├── `src/navigation.js`: view switching (Inputs/Report/About)
   ├── `src/metadata.js`: report metadata + safe filenames
   ├── `src/fileRead.js`: File → ArrayBuffer/Text helpers
   ├── `src/dom.js`: DOM id lookup + blob download helper
   └── `src/uiStatus.js`: status + loading UI

## Data Models
- **Student IDs input (`.txt`) → `ParsedStudentIds`** (`attendance.js`)
  - `orderedEntries`: array of `{type:'title'}` and `{type:'id'}` entries
  - `targetIdsSet`: unique IDs used for matching
  - `idCounts`: global counts for duplicate IDs
  - `sectionIdCounts`: per-section counts (if titles/sections used)
- **Workbook scan result** (`attendance.js`)
  - `found_log`: `FoundLogEntry[]` (sheet, id, name, cell, header/data row hints)
  - `not_found_ids`: `string[]`
  - `available_weeks`: `string[]` (e.g., `["W1","W2",...]`)
- **Report JSON contract → `ReportData`** (`report.js`)
  - `metadata`: timestamp/week/type/sheet_url + totals
  - `sheets`: per sheet `{name, column, header_row, students[]}`
  - `not_found`: `string[]`
  - optional: `input_order`, `ordered_rows`, `id_counts`

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
2025-12-17 | Approved by [Human]



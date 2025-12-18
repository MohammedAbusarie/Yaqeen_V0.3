# Design Patterns & Coding Standards

## Layer 1: Architecture Pattern
- **Pattern**: Small-module, vanilla JS ES Modules (ESM) with a thin UI wiring layer and pure-ish core logic.
- **Core principle**: Keep processing logic deterministic and testable (pure functions + explicit error types), keep DOM side-effects in handler/view layers.
- **File organization**:
  - Entry point + wiring: `app.js`
  - Core domain logic: `attendance.js`, `report.js`
  - UI rendering: `viewer.js`
  - Internal helpers: `src/*.js` (DOM/state/navigation/metadata/file IO/status)
- **Import structure**:
  - Use ESM `import { ... } from "./relative.js"` (browser-native modules)
  - `index.html` loads CDN libs that expose globals (`window.XLSX`, `window.jspdf`)

## Layer 2: Code Style
- **Naming**:
  - Functions/vars: `camelCase` (`parseStudentIdsText`, `readWorkbookFromArrayBuffer`)
  - Classes/Errors: `PascalCase` (`ReportViewer`, `ValidationError`)
  - Constants: prefer `const`
- **Variable declarations**:
  - Prefer `const`; use `let` only when reassigned
  - Avoid implicit globals
- **Function structure**:
  - Pure helpers for parsing/scanning/report building (`attendance.js`, `report.js`)
  - Async only for I/O (`fetch`, file reads)
- **Error handling**:
  - Throw explicit error subclasses for expected failures:
    - `ValidationError`, `FileError`, `DownloadError`, `ProcessingError` (`attendance.js`)
  - UI handlers catch and convert errors into user-facing status text (`src/handlers.js`)

## Layer 3: Component/Module Patterns
- **Handler factory pattern**:
  - Pattern: `createHandlers({ els, state, viewer, setStatus, disableRun, switchView })` returns an object of event callbacks.
  - When to use: all DOM event wiring in `app.js` should call into handlers returned by `createHandlers`.
  - Example: `src/handlers.js`
- **State container pattern**:
  - Pattern: `createInitialState()` returns a plain object; reset via `resetAll(...)`.
  - When to use: keep in-memory state minimal (workbook buffer/name + current report).
  - Example: `src/state.js`
- **Viewer class pattern**:
  - Pattern: `ReportViewer` owns DOM refs for report view and re-renders on state change.
  - When to use: report rendering concerns (tables, filters, summary).
  - Example: `viewer.js`

## Layer 4: Data Flow
- **State management**: In-memory JS object (`src/state.js`); no persistence.
- **API calls**:
  - Direct `fetch()` in `attendance.js` for Google Sheets XLSX export (best-effort; can fail due to CORS).
  - No internal service layer (project is static-only).
- **Data validation**:
  - Validate inputs in `src/handlers.js` (week range, type, required files/URL)
  - Validate parsed content in `attendance.js` (`parseStudentIdsText` enforces numeric IDs, length ≥ 6)

## FORBIDDEN PATTERNS (ZERO TOLERANCE)
- **Adding a backend requirement**: This project is explicitly static-only (deployable to Netlify).
- **Storing uploaded data remotely**: Privacy-first; processing must stay in the browser.
- **Silent error swallowing**: Expected failures must surface as status messages; unexpected failures must include context.

## Last Approved By
[Human name] | 2025-12-17



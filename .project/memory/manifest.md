# Project Manifest

## Active Tech Stack
- Framework: Static site (no framework) — HTML/CSS + Vanilla JavaScript (ES Modules)
- Database: None (in-memory only)
- Language: JavaScript (browser ESM)
- Auth: None

## Critical Constraints
- Static-only deployment (Netlify-compatible), no backend required
- Privacy-first: files processed locally in the browser; no server upload
- External dependencies are loaded via CDN and used as globals:
  - SheetJS (`window.XLSX`)
  - jsPDF (`window.jspdf`) + AutoTable (`doc.autoTable`)

## Last Updated
2025-12-17

## Memory Freshness Check
- Design Patterns: ✓ Current
- Known Bugs: ✓ Current
- Architecture: ✓ Current

## Decision Log
| Date | Decision | Approved/Rejected | Why | Notes |
|---|---|---|---|---|
| 2025-12-17 | Create missing required memory files in `.project/memory/` | Approved | Required by workspace rules (pre-action protocol) | Added `manifest.md`, `design-patterns.md`, `bugs-known.md`, `architecture.md`, `features-roadmap.md` |
| 2025-12-17 | Populate memory files from current repo state | Approved | Required by workspace rules (keep memory current) | Filled tech stack, architecture map, design patterns, and known limitations (CORS) |
| 2025-12-17 | Add in-browser workbook editor (any column) with preview/fix/confirm download | Approved | Required feature: not limited to column W; supports grades and attendance | Implemented via new editor workflow in `index.html`, `src/handlers.js`, and `attendance.js` |



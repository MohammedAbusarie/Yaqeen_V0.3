# Features & Roadmap

## Current Development
- Feature: Sheet Merger - Drag-and-drop column mapping to merge multiple sheets
- Status: Completed (2025-12-27)
- Owner: Mohammed Abusarie
- Description: Allows users to upload XLSX/ODS files or Google Sheet links, map columns from multiple sheets using drag-and-drop interface, and download a single merged output. Supports header detection (rows 1-5), sequential row concatenation, and optional duplicate header elimination.

## Completed Features
- ✅ Baseline static attendance report tool (XLSX + Student IDs → Report JSON/TXT/PDF)
- ✅ In-browser workbook editor with preview/fix/confirm workflow
- ✅ OCR experimental feature for attendance sheet image processing
- ✅ Sheet Merger with drag-and-drop column mapping (2025-12-27)
- ✅ Sheet Merger UX improvements: accordion grouping, search filtering, column counts (2025-12-27)

## Approved Next
- Feature: Improve Google Sheet URL reliability (reduce CORS failures via clearer UX + guidance; no backend)
- Priority: P2
- Estimated scope: 0.5–1 day
- Blocked by: Browser/Google Sheets CORS limitations

## Rejected / On Hold
- Feature: Add server-side processing / upload files to server
- Reason: Conflicts with static-only + privacy-first constraints
- Revisit date: Only with explicit human approval (architecture change)

## Last Reviewed
2025-12-27 by Mohammed Abusarie



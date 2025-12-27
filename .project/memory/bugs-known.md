# Known Bugs & Solutions

## Active Bugs
| Bug ID | Description | Root Cause | Solution Applied | Status |
|-----|----|-----|---|-----|
| BUG-001 | Google Sheet URL download can fail (CORS/network), even when URL is valid | Browser CORS restrictions or sheet sharing settings prevent fetching `/export?format=xlsx` | User workaround: export from Google Sheets manually and upload `.xlsx` | Open |

## Fixed Bugs (Do Not Repeat)
| Bug ID | What Failed | Why It Failed | How We Fixed It | Prevention |
|-----|---|---|-----|-----|
| BUG-FIXED-001 | Header detection could find w1, w2, w3, w4 under "attendance section" but not w1, w2, w3 under "attendance lecture", and vice versa (only one type shown) | When same headers (W1, W2, etc.) appear in both section and lecture areas, the code was creating a single entry and losing one of the types. The upgrade logic was converting section entries to lecture, losing section data. | Fixed by splitting locations by column ranges in final conversion step. Now creates separate entries: `section::W1`, `lecture::W1`, etc., based on which column range each location falls into. This preserves both section and lecture week columns even when they have the same header text. | When headers appear in multiple column ranges (section vs lecture), split locations by range and create separate entries for each. Don't rely solely on upgrade logic - verify actual column positions. |
| BUG-FIXED-002 | ODS files with formulas or special cell formats fail to load with "Unsupported value type" error | SheetJS was trying to parse formulas and cell styles from ODS files, which can cause parsing errors when the file format doesn't match expected structure. ODS files may not store cached calculated values for formulas. | Added `cellFormula: false` and `cellStyles: false` options to both read attempts in `readWorkbookFromArrayBuffer()`. Enhanced raw mode handling to extract calculated values from formula cells using `cell.w` (formatted text) as fallback when `cell.v` (calculated value) is missing. | Always use `cellFormula: false` and `cellStyles: false` when reading workbooks. Handle formula cells by checking for `cell.w` as fallback when `cell.v` is undefined. |
| BUG-FIXED-003 | When expanding a sheet in column pool, only first 2 columns are visible and scrolling doesn't work | `.columnGroup` had `overflow: hidden` which was clipping content and preventing column items from being visible beyond the first few items. Scrolling should occur at `.columnPool__list` level, not at the group level. | Removed `overflow: hidden` from `.columnGroup`. Applied border-radius directly to `.columnGroup__header` (top corners) and `.columnGroup__body` (bottom corners) to maintain rounded appearance. Added rule for collapsed state to apply full border-radius to header. | When using accordion/expandable groups, don't apply `overflow: hidden` to the group container if content needs to scroll. Apply overflow and scrolling at the parent container level (`.columnPool__list` in this case). Use border-radius on child elements to maintain visual design. |
|| BUG-FIXED-004 | "Eliminate duplicate header rows" checkbox in Sheet Merger had no effect - headers always included regardless of state | `startDataRow` always used `minHeaderRow` which skipped headers. Checkbox logic tried text-matching (unreliable), then re-added headers anyway. Both paths produced same output. | Made `startDataRow` conditional: `eliminateHeaders ? minHeaderRow : (minHeaderRow - 1)`. When checked: skip header, add single header at top. When unchecked: include each sheet's header row. | When implementing toggles, ensure conditional logic affects data flow directly, not just post-processing. |

## Performance Issues Tracked
- Issue: [Description]
- Metric: [Baseline vs target]
- Fix: [Applied solution]

## Last Updated
2025-12-27 (BUG-FIXED-004 added)



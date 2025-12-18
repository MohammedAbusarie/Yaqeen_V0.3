# Known Bugs & Solutions

## Active Bugs
| Bug ID | Description | Root Cause | Solution Applied | Status |
|-----|----|-----|---|-----|
| BUG-001 | Google Sheet URL download can fail (CORS/network), even when URL is valid | Browser CORS restrictions or sheet sharing settings prevent fetching `/export?format=xlsx` | User workaround: export from Google Sheets manually and upload `.xlsx` | Open |

## Fixed Bugs (Do Not Repeat)
| Bug ID | What Failed | Why It Failed | How We Fixed It | Prevention |
|-----|---|---|-----|-----|
| BUG-FIXED-001 | Header detection could find w1, w2, w3, w4 under "attendance section" but not w1, w2, w3 under "attendance lecture", and vice versa (only one type shown) | When same headers (W1, W2, etc.) appear in both section and lecture areas, the code was creating a single entry and losing one of the types. The upgrade logic was converting section entries to lecture, losing section data. | Fixed by splitting locations by column ranges in final conversion step. Now creates separate entries: `section::W1`, `lecture::W1`, etc., based on which column range each location falls into. This preserves both section and lecture week columns even when they have the same header text. | When headers appear in multiple column ranges (section vs lecture), split locations by range and create separate entries for each. Don't rely solely on upgrade logic - verify actual column positions. |

## Performance Issues Tracked
- Issue: [Description]
- Metric: [Baseline vs target]
- Fix: [Applied solution]

## Last Updated
2025-01-27


